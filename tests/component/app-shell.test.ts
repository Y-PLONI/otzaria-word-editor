/**
 * המעטפת עצמה — `App.vue` — מורכבת ונלחצת.
 *
 * ## למה הקובץ הזה קיים
 *
 * 915 שורות של חיווט מעטפת (שמירה, פתיחה, autosave, שם המסמך, דיאלוגים,
 * מטריקות) היו מאומתות **רק** ב-`readFileSync` + regex. סריקת מקור אינה יכולה
 * להבחין בין „הפונקציה קיימת” ל„הפונקציה מחוברת לפקד ומגיעה למי שאמור לענות
 * לה”, וזה נמדד: מוטציה שהסירה את `save?.setAutosaveEnabled(...)` מ-
 * `toggleAutosave` — כלומר החזירה את מתג השמירה האוטומטית להיות דקורטיבי,
 * הבאג המקורי בדיוק — עברה 203 בדיקות בירוק.
 *
 * ## הכפילים, ומה שנשאר אמיתי
 *
 * `onMounted` של המעטפת מקים מנוע SuperDoc אמיתי, ולכן מוחלפים בכפיל בדיוק
 * הדברים שאין להם קיום ב-jsdom או שהתשובה שלהם היא מה שנבדק:
 *
 *   * `engine/create-editor` — מייבא `superdoc` ו-workers. לא מגיעים אליו כאן
 *     בכלל (ה-swap מוחלף), אבל הייבוא הסטטי לבדו מפיל את ההרכבה.
 *   * `sessions/editor-swap` — מחזיר session מזויף שמצליח מיד.
 *   * `sessions/save-coordinator` — **זה מה שנמדד**: כל קריאה אליו מוקלטת, וגם
 *     ה-deps שלו נשמרים כדי שהבדיקה תוכל לדחוף snapshot ולראות מה הפס מציג.
 *   * `engine/command-adapter`, `engine/search`, `engine/doc-metrics`,
 *     `engine/document-defaults` — נשענים על handle של מנוע חי.
 *   * `host/settings` — כדי שהמתג יימדד גם על השאלה אם הבחירה נשמרה.
 *
 * מה שנשאר אמיתי: `TitleBar`, `Ribbon`, `StatusBar` והדיאלוגים, כל הזרימה של
 * `openDocument`, מטפל המקלדת, ו-`host/files`/`otzaria-client` (הם ניגשים
 * ל-`window.Otzaria` שאינו קיים ומחזירים כשל בשקט — בדיוק כמו בדפדפן).
 */
import { describe, expect, it, vi, beforeEach } from 'vitest';
import { mount } from '@vue/test-utils';
import {
  autoUnmount,
  createCommandDouble,
  createSuperdocDouble,
  settle,
  type SuperdocDouble,
} from './harness';
import type { SaveCoordinatorDeps, SaveSnapshot } from '../../src/sessions/save-coordinator';

/**
 * המצב המשותף לכפילים. `vi.hoisted` נדרש: מפעלי ה-`vi.mock` מורמים אל מעל
 * הייבואים ורצים לפני גוף הקובץ, ולכן משתנה רגיל בהיקף המודול היה TDZ.
 */
const stub = vi.hoisted(() => ({
  /** מה שהקואורדינטור קיבל. `[]` אחרי מוטציה שמנתקת את המתג. */
  autosaveCalls: [] as boolean[],
  /** מה ש-`saveNow` קיבל, לפי הסדר. */
  saveNowCalls: [] as Array<{ forceSaveAs?: boolean; suggestedName?: string } | undefined>,
  markDirtyCalls: 0,
  resetCalls: 0,
  /** מה שנשמר להפעלה הבאה. */
  persistedAutosave: [] as boolean[],
  /** מה שההעדפה השמורה מחזירה בעלייה. */
  storedAutosave: true,
  searchOpens: 0,
  /** ה-deps שהמעטפת נתנה לקואורדינטור — דרך לדחוף snapshot כמו המנוע. */
  saveDeps: null as SaveCoordinatorDeps | null,
  /** ה-session שה-swap „פתח”. מוגדר בכל בדיקה מחדש. */
  session: null as unknown,
  /** כפיל המופע שבתוך ה-session, כדי לראות מה המעטפת ביקשה מהמנוע. */
  superdoc: null as SuperdocDouble | null,
  /** האדפטר שהמעטפת תזריק לרצועה. */
  adapter: null as unknown,
}));

vi.mock('../../src/engine/create-editor', () => ({
  createEditor: vi.fn(),
  OPEN_TIMEOUT_MS: 1_000,
}));

vi.mock('../../src/sessions/editor-swap', () => ({
  createEditorSwap: () => ({
    get current() {
      return stub.session;
    },
    get isOpening() {
      return false;
    },
    open: async () => ({ status: 'opened', session: stub.session }),
    destroy: () => {},
  }),
}));

vi.mock('../../src/sessions/save-coordinator', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../src/sessions/save-coordinator')>()),
  createSaveCoordinator: (deps: SaveCoordinatorDeps) => {
    stub.saveDeps = deps;
    return {
      snapshot: {
        state: 'idle',
        isDirty: false,
        isSaving: false,
        targetToken: null,
        name: null,
        lastError: null,
      },
      markDirty: () => {
        stub.markDirtyCalls += 1;
      },
      setAutosaveEnabled: (enabled: boolean) => {
        stub.autosaveCalls.push(enabled);
      },
      adoptTarget: () => {},
      reset: () => {
        stub.resetCalls += 1;
      },
      saveNow: async (options?: { forceSaveAs?: boolean; suggestedName?: string }) => {
        stub.saveNowCalls.push(options);
        return { status: 'saved', token: 'token-1', name: 'מסמך.docx' };
      },
      dispose: () => {},
    };
  },
}));

vi.mock('../../src/engine/command-adapter', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../src/engine/command-adapter')>()),
  createCommandAdapter: () => stub.adapter,
}));

vi.mock('../../src/engine/search', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../src/engine/search')>();
  return {
    ...actual,
    createSearchAdapter: () => ({
      getState: () => actual.idleSearchState(),
      subscribe: () => () => {},
      open: () => {
        stub.searchOpens += 1;
        return { ok: true, snapshot: actual.idleSearchState() };
      },
      close: () => {},
      clear: () => {},
      find: () => ({ ok: true, snapshot: actual.idleSearchState() }),
      findDebounced: () => {},
      replace: async () => ({ ok: true, snapshot: actual.idleSearchState() }),
      replaceAll: async () => ({ ok: true, snapshot: actual.idleSearchState() }),
      dispose: () => {},
    }),
  };
});

vi.mock('../../src/engine/doc-metrics', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../src/engine/doc-metrics')>();
  return {
    ...actual,
    createDocMetrics: () => ({
      getState: () => actual.emptyDocMetrics(),
      noteDocumentChanged: () => {},
      noteSelectionChanged: () => {},
      notePaginationUpdate: () => {},
      measureNow: () => {},
      dispose: () => {},
    }),
  };
});

vi.mock('../../src/engine/document-defaults', () => ({
  applyHebrewDocumentDefaults: async () => ({ failures: [] }),
  applyHebrewPaperSize: async () => ({ applied: true }),
}));

vi.mock('../../src/host/settings', () => ({
  loadLastDocument: async () => null,
  saveLastDocument: async () => {},
  forgetLastDocument: async () => {},
  loadAutosaveEnabled: async () => stub.storedAutosave,
  saveAutosaveEnabled: async (enabled: boolean) => {
    stub.persistedAutosave.push(enabled);
  },
  loadRulerVisible: async () => false,
  saveRulerVisible: async () => {},
}));

// הייבוא **אחרי** ה-mocks במכוון (הם מורמים בכל מקרה, וזה הסדר שקורא נכון).
const { default: App } = await import('../../src/App.vue');

autoUnmount();

/** מרכיבה את המעטפת ומחזירה בקרה רק אחרי שכל זרימת ה-`onMounted` נרגעה. */
async function mountShell(): Promise<ReturnType<typeof mount>> {
  const wrapper = mount(App, { attachTo: document.body });
  // זרימת העלייה היא שרשרת של await-ים (העדפה, swap, openDocument, ברירות
  // מחדל של מסמך חדש), ולכן nextTick אחד אינו מספיק.
  await settle(12);
  return wrapper;
}

beforeEach(() => {
  stub.autosaveCalls.length = 0;
  stub.saveNowCalls.length = 0;
  stub.persistedAutosave.length = 0;
  stub.markDirtyCalls = 0;
  stub.resetCalls = 0;
  stub.searchOpens = 0;
  stub.storedAutosave = true;
  stub.saveDeps = null;
  stub.adapter = createCommandDouble();
  stub.superdoc = createSuperdocDouble();
  stub.session = {
    superdoc: stub.superdoc.host,
    // ה-controller המזויף: רק מה שהמעטפת נוגעת בו ישירות. שאר הקוראים
    // (`observeZoom`, `observeFontOptions`, `observeStyleGallery`) מתוכננים
    // ליפול לברירת מחדל כשה-handle חסר, וזה מה שנמדד בבדיקות שלהם.
    ui: { selection: { observe: () => () => {} } },
    onDispose: () => {},
    destroy: () => {},
  };
});

describe('הרכבת המעטפת', () => {
  it('העלייה פותחת מסמך ומחווטת את הפס, הרצועה ושורת המצב', async () => {
    // בלי זה כל הבדיקות למטה יכולות לעבור בירוק על מעטפת שלא סיימה לעלות.
    const wrapper = await mountShell();

    expect(wrapper.find('.word-titlebar').exists()).toBe(true);
    expect(wrapper.find('.word-statusbar').exists()).toBe(true);
    expect(wrapper.find('.editor-stack').exists()).toBe(true);
    expect(stub.saveDeps, 'הקואורדינטור הוקם').not.toBeNull();
    expect(stub.resetCalls, 'הפתיחה איפסה את מצב השמירה').toBe(1);
  });

  it('המסמך שנפתח מקבל את הסמן — אפשר להקליד בלי קליק מקדים', async () => {
    // הבאג שהתיקון בא לו: העורך נפתח, ההקלדה לא הגיעה לשום מקום, והמשתמש היה
    // צריך ללחוץ עם העכבר בגוף הטקסט לפני שיכול היה לכתוב מילה.
    await mountShell();

    expect(stub.superdoc?.ops(), 'הפתיחה ביקשה מהמנוע להחזיר את הסמן').toContain('focus');
  });

  it('פתיחה אינה חוטפת את הפוקוס משדה שמקלידים בו', async () => {
    // הפתיחה אסינכרונית ויכולה להימשך שניות. אם בינתיים המשתמש הקליד בשורת
    // החיפוש (שאינה מודאלית ונשארת פתוחה מעל המסמך), קפיצה לגוף המסמך הייתה
    // מוחקת לו את ההקלדה באמצע.
    const wrapper = await mountShell();
    await wrapper.find('.search-box').trigger('click');
    await settle();
    document.querySelector<HTMLInputElement>('#fr-search-input')?.focus();
    stub.superdoc?.reset();

    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'n', code: 'KeyN', ctrlKey: true }));
    await settle(12);

    expect(stub.resetCalls, 'המסמך החדש אכן נפתח').toBe(2);
    expect(stub.superdoc?.ops(), 'הפוקוס נשאר בשדה החיפוש').not.toContain('focus');
  });
});

describe('מתג השמירה האוטומטית', () => {
  it('לחיצה מגיעה לקואורדינטור — לא רק לצבע של הפיל', async () => {
    // זו המוטציה שחמקה: הסרת `save?.setAutosaveEnabled(...)` השאירה מתג שמזיז
    // את הכפתור ואינו מכבה שום דבר, ו-203 בדיקות עברו.
    const wrapper = await mountShell();

    // העלייה טוענת את ההעדפה השמורה ומעבירה אותה לקואורדינטור.
    expect(stub.autosaveCalls).toEqual([true]);

    const toggle = wrapper.find('.autosave-toggle');
    expect(toggle.attributes('aria-checked')).toBe('true');

    await toggle.trigger('click');
    await settle();

    expect(stub.autosaveCalls, 'הכיבוי הגיע לקואורדינטור').toEqual([true, false]);
    expect(wrapper.find('.autosave-toggle').attributes('aria-checked')).toBe('false');
  });

  it('הבחירה נשמרת להפעלה הבאה', async () => {
    const wrapper = await mountShell();

    await wrapper.find('.autosave-toggle').trigger('click');
    await settle();

    expect(stub.persistedAutosave).toEqual([false]);
  });

  it('ההעדפה השמורה היא זו שנטענת, ולא ברירת המחדל', async () => {
    // כיבוי בהפעלה קודמת חייב להגיע לקואורדינטור **לפני** העריכה הראשונה,
    // אחרת סבב ה-autosave הראשון רץ לפי ברירת המחדל.
    stub.storedAutosave = false;

    const wrapper = await mountShell();

    expect(stub.autosaveCalls).toEqual([false]);
    expect(wrapper.find('.autosave-toggle').attributes('aria-checked')).toBe('false');
  });
});

describe('שמירה', () => {
  it('כפתור השמירה בסרגל המהיר מריץ שמירה על המסמך הפתוח', async () => {
    const wrapper = await mountShell();

    await wrapper.findAll('.qa-btn')[0]!.trigger('click');
    await settle();

    expect(stub.saveNowCalls).toHaveLength(1);
    expect(stub.saveNowCalls[0]).toMatchObject({ forceSaveAs: false });
  });

  it('Ctrl+S שומר, ו-Ctrl+Shift+S הוא „שמור בשם”', async () => {
    // המטפל יושב על `window`, ולכן זה מה שמעיד שהוא נרשם בפועל.
    await mountShell();

    window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', code: 'KeyS', ctrlKey: true }));
    await settle();
    window.dispatchEvent(
      new KeyboardEvent('keydown', { key: 's', code: 'KeyS', ctrlKey: true, shiftKey: true }),
    );
    await settle();

    expect(stub.saveNowCalls.map((call) => call?.forceSaveAs)).toEqual([false, true]);
  });

  it('Ctrl+S שומר גם בפריסת מקלדת עברית', async () => {
    // הרגרסיה שהתיקון בא לה: בפריסה עברית הדפדפן מדווח `key: 'ד'`, וההשוואה
    // הישנה (`event.key === 's'`) פשוט לא תפסה. בעורך לכתיבת חידושי תורה זה
    // אומר שהשמירה מתה בדיוק כשהמשתמש עשה את מה שהתוסף נועד לו.
    await mountShell();

    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'ד', code: 'KeyS', ctrlKey: true }));
    await settle();

    expect(stub.saveNowCalls).toHaveLength(1);
  });

  it('Ctrl+P מדפיס גם בפריסה עברית, ו-Ctrl+G אינו נבלע', async () => {
    await mountShell();

    // `cancelable` נדרש כדי ש-`defaultPrevented` יהיה מדיד. keydown אמיתי
    // בדפדפן הוא cancelable; אירוע מלאכותי בלי הדגל אינו, ו-preventDefault בו
    // הוא no-op שקט.
    const options = { ctrlKey: true, cancelable: true };
    const print = new KeyboardEvent('keydown', { key: 'פ', code: 'KeyP', ...options });
    const unknown = new KeyboardEvent('keydown', { key: 'ג', code: 'KeyG', ...options });
    window.dispatchEvent(print);
    window.dispatchEvent(unknown);
    await settle();

    expect(print.defaultPrevented).toBe(true);
    // צירוף שאינו שלנו נשאר של הדפדפן.
    expect(unknown.defaultPrevented).toBe(false);
  });

  it('שינוי שם המסמך מסמן אותו כלא-שמור', async () => {
    const wrapper = await mountShell();

    const input = wrapper.find('.doc-title-input');
    (input.element as HTMLInputElement).value = 'חידושי בבא קמא';
    await input.trigger('change');
    await settle();

    expect(stub.markDirtyCalls).toBe(1);
    expect((wrapper.find('.doc-title-input').element as HTMLInputElement).value).toBe(
      'חידושי בבא קמא',
    );
  });

  it('מצב השמירה שהקואורדינטור מדווח מגיע לפס הכותרת', async () => {
    // החיווט הזה (`:is-dirty`, `:save-state-text`) היה מאומת רק ב-regex, וכפיל
    // שמדווח „מלוכלך” ופס שאינו משתנה נראים בסריקת מקור זהים.
    const wrapper = await mountShell();
    expect(wrapper.find('.dirty-indicator').exists()).toBe(false);

    const dirty: SaveSnapshot = {
      state: 'idle',
      isDirty: true,
      isSaving: false,
      targetToken: null,
      name: null,
      lastError: null,
    };
    stub.saveDeps!.onStateChange!(dirty);
    await settle();

    expect(wrapper.find('.dirty-indicator').exists()).toBe(true);
    expect(wrapper.find('.save-state-pill').text()).toBe('שינויים לא שמורים');
  });
});

describe('חיפוש', () => {
  it('כפתור החיפוש בפס פותח session במנוע ולא רק דיאלוג', async () => {
    // פתיחת הדיאלוג בלי `searchAdapter.open()` היא דיאלוג שכל חיפוש בו נכשל
    // סגור — ואת זה רואים רק ממעטפת מורכבת.
    const wrapper = await mountShell();

    await wrapper.find('.search-box').trigger('click');
    await settle();

    expect(stub.searchOpens).toBe(1);
  });
});
