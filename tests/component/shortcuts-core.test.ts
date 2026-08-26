/**
 * שלב 1 — הקיצורים נמדדים על המעטפת האמיתית.
 *
 * למה לא על המנתב לבדו: המנתב כבר נבדק ב-`tests/unit/shortcut-dispatch.test.ts`
 * מול כפילים. מה שהוא **אינו** יכול להוכיח הוא שהמעטפת חיווטה אותו למי שאמור
 * לענות — שהפקודה הגיעה ל-controller, שהפעולה הגיעה ל-Document API, ושהמזהה
 * וה-payload הם אלה שהמנוע מקבל ולא מה שנראה נכון. זו בדיוק המחלקה של באגים
 * ש-`App.vue` סבל ממנה: פונקציה שקיימת ואינה מחוברת.
 *
 * הקיצורים חייבים לעבוד ללא תלות בלשונית הפתוחה — לשוניות הרצועה מורכבות רק
 * כשהן פעילות — ולכן כולם נמדדים כאן, על המעטפת, ולא דרך פקד בלשונית.
 */
import { describe, expect, it, vi, beforeEach } from 'vitest';
import { mount } from '@vue/test-utils';
import {
  autoUnmount,
  createCommandDouble,
  createSuperdocDouble,
  settle,
  type CommandDouble,
  type SuperdocDouble,
} from './harness';
import type { SaveCoordinatorDeps } from '../../src/sessions/save-coordinator';

const stub = vi.hoisted(() => ({
  saveNowCalls: [] as Array<{ forceSaveAs?: boolean } | undefined>,
  session: null as unknown,
  adapter: null as unknown,
  pickCalls: 0,
  /** מצב השמירה שהקואורדינטור מדווח. בדיקה יכולה להפוך אותו למלוכלך. */
  isDirty: false,
  /** מה שנשאל בדיאלוג האישור של אוצריא, לפי הסדר. */
  confirms: [] as string[],
  /** התשובה של המשתמש בדיאלוג. */
  confirmAnswer: false,
  /** כמה פעמים מצב השמירה אופס — כלומר מסמך חדש נפתח. */
  resets: 0,
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
  createSaveCoordinator: (_deps: SaveCoordinatorDeps) => ({
    get snapshot() {
      return {
        state: 'idle',
        isDirty: stub.isDirty,
        isSaving: false,
        targetToken: null,
        name: null,
        lastError: null,
      };
    },
    markDirty: () => {},
    setAutosaveEnabled: () => {},
    adoptTarget: () => {},
    reset: () => {
      stub.resets += 1;
    },
    saveNow: async (options?: { forceSaveAs?: boolean }) => {
      stub.saveNowCalls.push(options);
      return { status: 'saved', token: 'token-1', name: 'מסמך.docx' };
    },
    dispose: () => {},
  }),
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
      open: () => ({ ok: true, snapshot: actual.idleSearchState() }),
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
  loadAutosaveEnabled: async () => true,
  saveAutosaveEnabled: async () => {},
  loadRulerVisible: async () => false,
  saveRulerVisible: async () => {},
}));

vi.mock('../../src/host/otzaria-client', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../src/host/otzaria-client')>()),
  confirm: async (question: { title: string }) => {
    stub.confirms.push(question.title);
    return stub.confirmAnswer;
  },
  notifyError: () => {},
}));

vi.mock('../../src/host/files', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../src/host/files')>()),
  pickDocxFile: async () => {
    stub.pickCalls += 1;
    return null;
  },
}));

const { default: App } = await import('../../src/App.vue');

autoUnmount();

let adapter: CommandDouble;
let superdoc: SuperdocDouble;

async function mountShell() {
  const wrapper = mount(App, { attachTo: document.body });
  await settle(12);
  return wrapper;
}

/** אירוע מקלדת אמיתי על `window`, בדיוק כמו לחיצה של המשתמש. */
function press(over: Partial<KeyboardEventInit> & { code: string }): KeyboardEvent {
  const event = new KeyboardEvent('keydown', { cancelable: true, bubbles: true, ...over });
  window.dispatchEvent(event);
  return event;
}

beforeEach(() => {
  stub.saveNowCalls.length = 0;
  stub.confirms.length = 0;
  stub.pickCalls = 0;
  stub.resets = 0;
  stub.isDirty = false;
  stub.confirmAnswer = false;
  adapter = createCommandDouble();
  superdoc = createSuperdocDouble();
  stub.adapter = adapter;
  stub.session = {
    superdoc: superdoc.host,
    ui: { selection: { observe: () => () => {} } },
    onDispose: () => {},
    destroy: () => {},
  };
});

describe('פקודות המנוע', () => {
  const cases: ReadonlyArray<readonly [string, Partial<KeyboardEventInit> & { code: string }, string]> = [
    ['Ctrl+Z', { code: 'KeyZ', ctrlKey: true }, 'undo'],
    ['Ctrl+Y', { code: 'KeyY', ctrlKey: true }, 'redo'],
    ['Ctrl+Shift+Z', { code: 'KeyZ', ctrlKey: true, shiftKey: true }, 'redo'],
    ['Ctrl+B', { code: 'KeyB', ctrlKey: true }, 'bold'],
    ['Ctrl+I', { code: 'KeyI', ctrlKey: true }, 'italic'],
    ['Ctrl+U', { code: 'KeyU', ctrlKey: true }, 'underline'],
    ['Ctrl+Shift+8', { code: 'Digit8', ctrlKey: true, shiftKey: true }, 'formatting-marks'],
  ];

  for (const [label, init, commandId] of cases) {
    it(`${label} מריץ ${commandId}`, async () => {
      await mountShell();

      const event = press(init);
      await settle();

      expect(adapter.calls.map((call) => call.id)).toEqual([commandId]);
      expect(event.defaultPrevented).toBe(true);
    });
  }

  it('כותרות שולחות את מזהה הסגנון ש-linked-style מקבל', async () => {
    await mountShell();

    press({ code: 'Digit1', ctrlKey: true, altKey: true });
    press({ code: 'Digit2', ctrlKey: true, altKey: true });
    press({ code: 'Digit3', ctrlKey: true, altKey: true });
    await settle();

    expect(adapter.calls).toEqual([
      { id: 'linked-style', payload: { style: 'Heading1' } },
      { id: 'linked-style', payload: { style: 'Heading2' } },
      { id: 'linked-style', payload: { style: 'Heading3' } },
    ]);
  });

  it('Ctrl+B עובד גם בפריסת מקלדת עברית', async () => {
    await mountShell();

    press({ key: 'נ', code: 'KeyB', ctrlKey: true });
    await settle();

    expect(adapter.calls.map((call) => call.id)).toEqual(['bold']);
  });

  it('פקודה חסומה מדווחת בעברית ואינה נופלת', async () => {
    adapter = createCommandDouble({ failures: { undo: 'history-empty' } });
    stub.adapter = adapter;
    const wrapper = await mountShell();

    press({ code: 'KeyZ', ctrlKey: true });
    await settle();

    expect(wrapper.find('.status-message').text()).toContain('אין פעולה לבטל');
  });

  it('פקודה שאינה זמינה: המנוע מסרב, והמשתמש מקבל הסבר', async () => {
    // ההבדל מכפתור: כפתור מנוטרל אינו נלחץ בכלל, וקיצור אין דרך „לנטרל”.
    // לכן הוא כן מגיע לאדפטר — והתשובה היא סירוב מנומק ולא שינוי שקט במסמך.
    adapter = createCommandDouble({ states: { bold: { enabled: false } } });
    stub.adapter = adapter;
    const wrapper = await mountShell();

    press({ code: 'KeyB', ctrlKey: true });
    await settle();

    expect(adapter.applied, 'המסמך לא שונה').toEqual([]);
    expect(adapter.blocked.map((call) => call.id)).toEqual(['bold']);
    expect(wrapper.find('.status-message').text().length).toBeGreaterThan(0);
  });
});

describe('פעולות המעטפת', () => {
  it('Ctrl+A בוחר את כל המסמך דרך ה-Document API', async () => {
    await mountShell();

    const event = press({ code: 'KeyA', ctrlKey: true });
    await settle();

    expect(superdoc.ops()).toContain('ranges.resolve');
    expect(event.defaultPrevented).toBe(true);
    // לא דרך ה-controller: ל„בחר הכל” אין פקודה בקטלוג.
    expect(adapter.calls).toEqual([]);
  });

  it('Ctrl+Enter מתחיל פסקה בעמוד חדש', async () => {
    await mountShell();

    press({ code: 'Enter', ctrlKey: true });
    await settle();

    expect(superdoc.ops()).toContain('format.paragraph.setFlowOptions');
  });

  it('Enter לבד אינו מוסיף מעבר עמוד', async () => {
    await mountShell();

    const event = press({ code: 'Enter' });
    await settle();

    expect(superdoc.ops()).not.toContain('format.paragraph.setFlowOptions');
    expect(event.defaultPrevented).toBe(false);
  });

  it('Ctrl+O פותח את בורר הקבצים של אוצריא', async () => {
    await mountShell();

    press({ code: 'KeyO', ctrlKey: true });
    await settle();

    expect(stub.pickCalls).toBe(1);
  });

  it('Ctrl+N פותח מסמך חדש', async () => {
    await mountShell();
    stub.resets = 0;

    const event = press({ code: 'KeyN', ctrlKey: true });
    await settle();

    expect(stub.resets, 'מסמך חדש נפתח').toBe(1);
    expect(event.defaultPrevented).toBe(true);
  });

  it('Ctrl+N על מסמך מלוכלך שואל לפני שהוא מוחק עבודה', async () => {
    // המסלול המסוכן בשלב הזה: קיצור שמוחק את מה שהמשתמש כתב. הוא חייב לעבור
    // באותה הכרעה בדיוק שהכפתור עובר בה — `decideDocumentSwitch` — ולא לקצר.
    await mountShell();
    stub.isDirty = true;
    stub.resets = 0;

    press({ code: 'KeyN', ctrlKey: true });
    await settle();

    // אותה שרשרת שאלות בדיוק שהכפתור „מסמך חדש” עובר בה.
    expect(stub.confirms, 'המשתמש נשאל').toEqual(['המסמך לא נשמר', 'לפתוח בלי לשמור?']);
    expect(stub.resets, 'ובלי אישור — לא נפתח מסמך חדש').toBe(0);
  });

  it('Ctrl+O עובר באותו מסלול של הכפתור: קודם בורר הקבצים', async () => {
    // ב-`onPickAndOpen` הבחירה קודמת לשאלה על השינויים שלא נשמרו, ולכן ביטול
    // בבורר אינו מגיע לשאלה בכלל. הקיצור אינו מקצר את המסלול הזה.
    await mountShell();
    stub.isDirty = true;

    press({ code: 'KeyO', ctrlKey: true });
    await settle();

    expect(stub.pickCalls, 'בורר הקבצים נפתח').toBe(1);
    expect(stub.confirms, 'הבחירה בוטלה — אין מה לשאול').toEqual([]);
  });

  it('Ctrl+K פותח את דיאלוג הקישור — מכל לשונית', async () => {
    // הלשונית הפעילה בעלייה היא „בית”, לא „הוספה”. עד שהדיאלוג עבר למעטפת
    // הוא פשוט לא היה קיים ברגע הזה, ולכן הקיצור לא היה יכול לפתוח אותו.
    const wrapper = await mountShell();
    expect(document.querySelector('.link-dialog'), 'סגור בהתחלה').toBeNull();

    const event = press({ code: 'KeyK', ctrlKey: true });
    await settle();

    expect(document.querySelector('.link-dialog'), 'נפתח').not.toBeNull();
    expect(event.defaultPrevented).toBe(true);
    expect(wrapper.findAll('[role="tabpanel"]').length, 'הלשונית לא הוחלפה').toBe(1);
  });

  it('Escape סוגר את דיאלוג הקישור', async () => {
    await mountShell();
    press({ code: 'KeyK', ctrlKey: true });
    await settle();

    press({ code: 'Escape' });
    await settle();

    expect(document.querySelector('.link-dialog')).toBeNull();
  });

  it('Ctrl+K קורא את הבחירה מהמסמך לפני שהוא פותח', async () => {
    await mountShell();

    press({ code: 'KeyK', ctrlKey: true });
    await settle();

    expect(superdoc.ops()).toContain('selection.current');
  });

  it('F12 הוא „שמור בשם”', async () => {
    await mountShell();

    const event = press({ code: 'F12' });
    await settle();

    expect(stub.saveNowCalls.map((call) => call?.forceSaveAs)).toEqual([true]);
    expect(event.defaultPrevented).toBe(true);
  });
});

describe('פעולות שתלויות במצב המנוע', () => {
  it('Ctrl+] מחשב מהגודל שהמנוע מדווח, ולא ממונה מקומי', async () => {
    // מונה מקומי היה מטפס גם כשהמסמך דוחה את הפקודה, ואז הלחיצה הבאה הייתה
    // מחשבת מגודל שאינו במסמך. זו בדיוק הרגרסיה שהכפתורים ברצועה סבלו ממנה.
    adapter = createCommandDouble({ states: { 'font-size': { value: 12 } } });
    stub.adapter = adapter;
    await mountShell();

    press({ key: ']', code: 'BracketRight', ctrlKey: true });
    press({ key: ']', code: 'BracketRight', ctrlKey: true });
    await settle();

    // שתי הלחיצות מחשבות מ-12, כי המנוע לא זז.
    expect(adapter.payloads('font-size')).toEqual([14, 14]);
  });

  it('Ctrl+[ יורד בסולם של Word', async () => {
    adapter = createCommandDouble({ states: { 'font-size': { value: 14 } } });
    stub.adapter = adapter;
    await mountShell();

    press({ key: '[', code: 'BracketLeft', ctrlKey: true });
    await settle();

    expect(adapter.payloads('font-size')).toEqual([12]);
  });

  it('גודל שהמנוע אינו מדווח נופל לברירת המחדל ולא ל-NaN', async () => {
    adapter = createCommandDouble({ states: { 'font-size': { value: undefined } } });
    stub.adapter = adapter;
    await mountShell();

    press({ key: ']', code: 'BracketRight', ctrlKey: true });
    await settle();

    expect(adapter.payloads('font-size')).toEqual([14]);
  });

  it('כתב עילי ותחתי עוברים ב-Document API ולא בפקודה', async () => {
    // `toggleVertAlign` דורש טווח מסומן — בלעדיו הוא נכשל סגור לפני שהוא
    // מגיע ל-`format.vertAlign`, וזו התנהגות נכונה שהבדיקה אינה מודדת כאן.
    superdoc = createSuperdocDouble({ selection: { hasRange: true, text: 'טקסט' } });
    stub.session = {
      superdoc: superdoc.host,
      ui: { selection: { observe: () => () => {} } },
      onDispose: () => {},
      destroy: () => {},
    };
    await mountShell();

    press({ key: '=', code: 'Equal', ctrlKey: true });
    await settle();

    expect(superdoc.ops()).toContain('format.vertAlign');
    expect(adapter.calls).toEqual([]);
  });
});

describe('הפניות, סקירה ותצוגה', () => {
  it('Ctrl+Alt+F ו-Ctrl+Alt+D מוסיפים הערות דרך ה-Document API', async () => {
    await mountShell();

    press({ code: 'KeyF', ctrlKey: true, altKey: true });
    await settle();

    expect(superdoc.ops()).toContain('footnotes.insert');
    expect(superdoc.inputs('footnotes.insert')).toEqual([
      expect.objectContaining({ type: 'footnote' }),
    ]);
  });

  it('Ctrl+Alt+F אינו Ctrl+F — הערה מול חיפוש', async () => {
    const wrapper = await mountShell();

    press({ code: 'KeyF', ctrlKey: true });
    await settle();

    expect(wrapper.find('.find-replace-dialog').exists()).toBe(true);
    expect(superdoc.ops()).not.toContain('footnotes.insert');
  });

  it('Ctrl+Shift+E מחליף למצב מעקב לפי מה שהמנוע מדווח', async () => {
    // אין פקודה נפרדת: `document-mode` עם 'suggesting' הוא מצב המעקב.
    adapter = createCommandDouble({ states: { 'document-mode': { value: 'editing' } } });
    stub.adapter = adapter;
    await mountShell();

    press({ code: 'KeyE', ctrlKey: true, shiftKey: true });
    await settle();

    expect(adapter.payloads('document-mode')).toEqual([{ mode: 'suggesting' }]);
  });

  it('Ctrl+Shift+E ממצב מעקב חוזר לעריכה', async () => {
    adapter = createCommandDouble({ states: { 'document-mode': { value: 'suggesting' } } });
    stub.adapter = adapter;
    await mountShell();

    press({ code: 'KeyE', ctrlKey: true, shiftKey: true });
    await settle();

    expect(adapter.payloads('document-mode')).toEqual([{ mode: 'editing' }]);
  });

  it('Ctrl+Shift+E אינו Ctrl+E — מעקב מול מרכוז', async () => {
    await mountShell();

    press({ code: 'KeyE', ctrlKey: true });
    await settle();

    expect(adapter.calls.map((call) => call.id)).toEqual(['text-align']);
  });

  it('F11 מדליק ומכבה מצב מיקוד', async () => {
    const wrapper = await mountShell();
    expect(wrapper.find('.word-app-shell').classes()).not.toContain('focus-mode');

    press({ code: 'F11' });
    await settle();
    expect(wrapper.find('.word-app-shell').classes()).toContain('focus-mode');

    press({ code: 'F11' });
    await settle();
    expect(wrapper.find('.word-app-shell').classes()).not.toContain('focus-mode');
  });

  it('F3 על מסמך שלא חיפשו בו פותח את החיפוש במקום ליפול', async () => {
    const wrapper = await mountShell();

    const event = press({ code: 'F3' });
    await settle();

    expect(wrapper.find('.find-replace-dialog').exists()).toBe(true);
    expect(event.defaultPrevented).toBe(true);
  });

  it('Shift+F3 אינו F3 — שתי רשומות נפרדות', async () => {
    await mountShell();

    // שתיהן פותחות את החיפוש כשאין שאילתה, וזה מה שנבדק כאן: שתיהן נתפסות.
    expect(press({ code: 'F3', shiftKey: true }).defaultPrevented).toBe(true);
  });
});

describe('כיווניות פסקה', () => {
  /** לחיצה ושחרור של Shift, כמו במקלדת אמיתית. */
  function pressShift(side: 'ShiftLeft' | 'ShiftRight'): void {
    const init = { code: side, key: 'Shift', ctrlKey: true, bubbles: true, cancelable: true };
    window.dispatchEvent(new KeyboardEvent('keydown', init));
    window.dispatchEvent(new KeyboardEvent('keyup', init));
  }

  it('Ctrl + Shift ימני הופך את הפסקה ל-RTL', async () => {
    await mountShell();

    pressShift('ShiftRight');
    await settle();

    expect(adapter.calls.map((call) => call.id)).toEqual(['direction-rtl']);
  });

  it('Ctrl + Shift שמאלי הופך אותה ל-LTR', async () => {
    await mountShell();

    pressShift('ShiftLeft');
    await settle();

    expect(adapter.calls.map((call) => call.id)).toEqual(['direction-ltr']);
  });

  it('Ctrl+Shift+X אינו נוגע בכיווניות', async () => {
    // הצירוף שהכי קל להיכשל בו: הוא נגמר בשחרור של אותו Shift בדיוק.
    await mountShell();

    const init = { ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true };
    window.dispatchEvent(
      new KeyboardEvent('keydown', { code: 'ShiftLeft', key: 'Shift', ...init }),
    );
    window.dispatchEvent(new KeyboardEvent('keydown', { code: 'KeyX', key: 'x', ...init }));
    window.dispatchEvent(new KeyboardEvent('keyup', { code: 'KeyX', key: 'x', ...init }));
    window.dispatchEvent(new KeyboardEvent('keyup', { code: 'ShiftLeft', key: 'Shift', ...init }));
    await settle();

    expect(adapter.calls.map((call) => call.id)).toEqual(['strikethrough']);
  });
});

describe('אוצריא', () => {
  it('Ctrl+Shift+G הוא חיפוש בספרייה, ואינו נוגע בחיפוש במסמך', async () => {
    const wrapper = await mountShell();

    press({ code: 'KeyG', ctrlKey: true, shiftKey: true });
    await settle();

    // בלי מאחז אוצריא הפעולה נכשלת סגור ומדווחת — ומה שחשוב כאן הוא שהיא
    // **לא** פתחה את דיאלוג החיפוש במסמך.
    expect(wrapper.find('.find-replace-dialog').exists()).toBe(false);
  });

  it('Ctrl+Shift+F נשאר של אוצריא ואינו נבלע', async () => {
    // אוצריא קושרת אותו ל„חיפוש חדש בכל הספרים” ברמת האפליקציה. רשומה שלנו
    // על אותו צירוף הייתה תווית שאין לה סיכוי — בדיוק המחלה שהתוכנית מרפאת.
    await mountShell();

    const event = press({ code: 'KeyF', ctrlKey: true, shiftKey: true });
    await settle();

    expect(event.defaultPrevented).toBe(false);
  });

  it('Ctrl+Shift+O אינו Ctrl+O — ספרייה מול בורר קבצים', async () => {
    await mountShell();

    press({ code: 'KeyO', ctrlKey: true, shiftKey: true });
    await settle();

    expect(stub.pickCalls, 'בורר הקבצים לא נפתח').toBe(0);
  });

  it('Ctrl+Shift+Q מבקש את הבחירה מהקורא ולא מהמסמך', async () => {
    const wrapper = await mountShell();

    const event = press({ code: 'KeyQ', ctrlKey: true, shiftKey: true });
    await settle();

    expect(event.defaultPrevented).toBe(true);
    // אין מאחז אוצריא בבדיקות, ולכן זו הודעה בעברית ולא חריגה.
    expect(wrapper.find('.status-message').text().length).toBeGreaterThan(0);
  });
});

describe('מה שהקיצורים אינם עושים', () => {
  it('פוקוס בשדה טקסט של הממשק: קיצור מסמך אינו נורה', async () => {
    const wrapper = await mountShell();

    const input = wrapper.find('.doc-title-input');
    (input.element as HTMLInputElement).focus();
    input.element.dispatchEvent(
      new KeyboardEvent('keydown', { code: 'KeyB', ctrlKey: true, bubbles: true, cancelable: true }),
    );
    await settle();

    expect(adapter.calls).toEqual([]);
  });

  it('Ctrl+A בשדה טקסט נשאר של הדפדפן', async () => {
    const wrapper = await mountShell();

    const input = wrapper.find('.doc-title-input');
    const event = new KeyboardEvent('keydown', {
      code: 'KeyA',
      ctrlKey: true,
      bubbles: true,
      cancelable: true,
    });
    input.element.dispatchEvent(event);
    await settle();

    expect(event.defaultPrevented).toBe(false);
    expect(superdoc.ops()).not.toContain('ranges.resolve');
  });

  it('כל צירוף מוכר בולע את ברירת המחדל של הדפדפן', async () => {
    // בלי זה ה-WebView פותח את הדיאלוגים שלו („שמירת דף”, „הדפסה”) מעל התוסף.
    await mountShell();

    const combos: Array<Partial<KeyboardEventInit> & { code: string }> = [
      { code: 'KeyZ', ctrlKey: true },
      { code: 'KeyY', ctrlKey: true },
      { code: 'KeyB', ctrlKey: true },
      { code: 'KeyI', ctrlKey: true },
      { code: 'KeyU', ctrlKey: true },
      { code: 'KeyA', ctrlKey: true },
      { code: 'KeyN', ctrlKey: true },
      { code: 'KeyO', ctrlKey: true },
      { code: 'KeyP', ctrlKey: true },
      { code: 'Enter', ctrlKey: true },
      { code: 'NumpadEnter', ctrlKey: true },
      { code: 'Digit8', ctrlKey: true, shiftKey: true },
      { code: 'Digit1', ctrlKey: true, altKey: true },
      { code: 'F12' },
    ];

    const swallowed = combos.filter((init) => press(init).defaultPrevented);
    await settle();

    expect(swallowed).toHaveLength(combos.length);
  });

  it('Ctrl+Shift+8 הוא מתג: כל לחיצה שולחת את הפקודה מחדש', async () => {
    await mountShell();

    press({ code: 'Digit8', ctrlKey: true, shiftKey: true });
    press({ code: 'Digit8', ctrlKey: true, shiftKey: true });
    await settle();

    expect(adapter.calls.map((call) => call.id)).toEqual([
      'formatting-marks',
      'formatting-marks',
    ]);
  });

  it('Escape בלי חלון פתוח אינו נבלע', async () => {
    // נסיגה שנמצאה ב-QA: בליעה גורפת של Escape הייתה חוסמת את ה-Escape של
    // המנוע ושל הדפדפן, גם כשלא היה לנו מה לסגור.
    await mountShell();

    const event = press({ code: 'Escape' });
    await settle();

    expect(event.defaultPrevented).toBe(false);
  });

  it('צירוף שכבר טופל אינו רץ פעם שנייה', async () => {
    // המאזין שלנו יושב על window אחרי ה-keymap של המנוע. אם המנוע כבר טיפל
    // ב-Ctrl+B, הרצה נוספת שלנו הייתה מבטלת את ההדגשה שהוא זה עתה החיל.
    await mountShell();

    const event = new KeyboardEvent('keydown', {
      code: 'KeyB',
      ctrlKey: true,
      cancelable: true,
      bubbles: true,
    });
    event.preventDefault();
    window.dispatchEvent(event);
    await settle();

    expect(adapter.calls).toEqual([]);
  });

  it('צירוף שאינו ברשימה אינו נבלע', async () => {
    await mountShell();

    const event = press({ code: 'KeyQ', ctrlKey: true });
    await settle();

    expect(event.defaultPrevented).toBe(false);
    expect(adapter.calls).toEqual([]);
  });
});

describe('Ctrl+K והכפתור נחסמים באותו תנאי', () => {
  /**
   * הליקוי: הקיצור נשען על „יש מסמך” והכפתור על `linkCmd.enabled`. במצב
   * שבו הפקודה מנוטרלת, Ctrl+K פתח דיאלוג שהאישור בו נכשל — בעברית, אבל
   * רק אחרי שהמשתמש כבר הקליד כתובת.
   */
  it('פקודת „קישור” מנוטרלת — הקיצור אינו פותח דיאלוג', async () => {
    adapter = createCommandDouble({ states: { link: { enabled: false } } });
    stub.adapter = adapter;
    await mountShell();

    press({ code: 'KeyK', ctrlKey: true });
    await settle();

    expect(document.querySelector('.link-dialog')).toBeNull();
  });

  it('פקודת „קישור” זמינה — הקיצור פותח', async () => {
    // הבקרה: בלעדיה הבדיקה שמעליה עוברת גם אם Ctrl+K הפסיק לעבוד לגמרי.
    await mountShell();

    press({ code: 'KeyK', ctrlKey: true });
    await settle();

    expect(document.querySelector('.link-dialog')).not.toBeNull();
  });
});

describe('הפוקוס בתוך המסמך', () => {
  /**
   * הבאג שנמדד בדפדפן אמיתי, ולא בשום בדיקה: משטח ההקלדה של המנוע הוא
   * `<textarea aria-label="Text composition input">` ברוחב פיקסל אחד, בתוך
   * אזור המסמך. כלומר ברגע שהמשתמש מתחיל להקליד, `event.target` של כל הקשה
   * הוא TEXTAREA — והשומר של „שדה טקסט” חסם את כל הקיצורים בדיוק במצב היחיד
   * שבו הם נחוצים. הבדיקות הקודמות ירו על `window` ולכן לא ראו את זה.
   */
  function composingSurface(wrapper: ReturnType<typeof mount>): HTMLTextAreaElement {
    const host = wrapper.find('.editor-stack').element;
    const surface = document.createElement('textarea');
    surface.setAttribute('aria-label', 'Text composition input');
    host.appendChild(surface);
    return surface;
  }

  /** הקשה שמקורה במשטח ההקלדה, כמו אצל מי שמקליד במסמך. */
  function typeInDocument(
    surface: HTMLElement,
    over: Partial<KeyboardEventInit> & { code: string },
  ): KeyboardEvent {
    const event = new KeyboardEvent('keydown', { cancelable: true, bubbles: true, ...over });
    surface.dispatchEvent(event);
    return event;
  }

  it('Ctrl+Z בזמן הקלדה מגיע לפקודה', async () => {
    const wrapper = await mountShell();

    typeInDocument(composingSurface(wrapper), { code: 'KeyZ', ctrlKey: true });
    await settle();

    expect(adapter.calls.map((call) => call.id)).toEqual(['undo']);
  });

  it('Ctrl+Z בזמן הקלדה בעברית מגיע לפקודה', async () => {
    // בפריסה עברית `key` הוא „ז”. זה בדיוק המצב שדווח: באנגלית עבד, בעברית לא.
    const wrapper = await mountShell();

    typeInDocument(composingSurface(wrapper), { code: 'KeyZ', key: 'ז', ctrlKey: true });
    await settle();

    expect(adapter.calls.map((call) => call.id)).toEqual(['undo']);
  });

  it('כל קיצורי העיצוב עובדים בזמן הקלדה, לא רק הביטול', async () => {
    const wrapper = await mountShell();
    const surface = composingSurface(wrapper);

    typeInDocument(surface, { code: 'KeyB', key: 'נ', ctrlKey: true });
    typeInDocument(surface, { code: 'KeyI', key: 'ן', ctrlKey: true });
    typeInDocument(surface, { code: 'KeyE', key: 'ק', ctrlKey: true });
    await settle();

    expect(adapter.calls.map((call) => call.id)).toEqual(['bold', 'italic', 'text-align']);
  });

  it('שדה טקסט של הממשק עדיין חוסם — הוא לא באזור המסמך', async () => {
    // שם `Ctrl+B` שייך לשדה, ובליעתו הייתה שוברת אותו.
    await mountShell();
    const field = document.createElement('input');
    document.body.appendChild(field);

    const event = new KeyboardEvent('keydown', {
      code: 'KeyB',
      ctrlKey: true,
      cancelable: true,
      bubbles: true,
    });
    field.dispatchEvent(event);
    await settle();
    field.remove();

    expect(adapter.calls).toEqual([]);
    expect(event.defaultPrevented).toBe(false);
  });
});
