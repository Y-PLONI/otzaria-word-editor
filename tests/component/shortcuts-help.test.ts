/**
 * שלב 6 — גילוי ונגישות: דיאלוג „קיצורי מקלדת” ומעגל המיקוד של `F6`.
 *
 * הדיאלוג נבנה מהרג'יסטרי, ולכן מה שנמדד כאן אינו „יש רשימה” אלא **שהרשימה
 * היא הרג'יסטרי**: כל מזהה מיוצג, והספירה זהה. רשימה כתובה ביד הייתה עוברת
 * את הבדיקה הראשונה ונופלת בשתי האחרות — וזה בדיוק ההבדל שהתוכנית הזאת באה
 * לשמור עליו.
 *
 * שסריקת המקור אינה מוצאת תווית קיצור כתובה ביד באף קומפוננטה נבדק
 * ב-`tests/contract/shortcut-registry.test.ts`, ולא כאן: זו בדיקה על המקור
 * ולא על ההרכבה.
 */
import { describe, expect, it, vi, beforeEach } from 'vitest';
import { mount } from '@vue/test-utils';
import {
  autoUnmount,
  createCommandDouble,
  createSuperdocDouble,
  settle,
  tipMessage,
  tipOf,
  type CommandDouble,
  type SuperdocDouble,
} from './harness';
import type { CreateEditorOptions, EditorSession } from '../../src/engine/create-editor';
import type { SaveCoordinatorDeps } from '../../src/sessions/save-coordinator';

const stub = vi.hoisted(() => ({
  saveNowCalls: [] as Array<{ forceSaveAs?: boolean } | undefined>,
  openEditor: null as ((container: HTMLElement) => EditorSession) | null,
  adapters: new WeakMap<object, CommandDouble>(),
  pickCalls: 0,
  /** מצב השמירה שהקואורדינטור מדווח. בדיקה יכולה להפוך אותו למלוכלך. */
  isDirty: false,
  /** מה שנשאל בדיאלוג האישור של אוצריא, לפי הסדר. */
  confirms: [] as string[],
  /** התשובה של המשתמש בדיאלוג. */
  confirmAnswer: false,
  /** כמה פעמים מצב השמירה אופס — כלומר מסמך חדש נפתח. */
  resets: 0,
  /** רשומת ההפעלה שהעלייה תמצא. `null` = הפעלה ראשונה. */
  sessionRecord: null as unknown,
}));

/**
 * Keep the real swap lifecycle: current is null until the first open completes,
 * and replacing/closing an editor disposes that editor's interaction lease.
 * A pre-populated swap double prevents App from registering its first adapter.
 */
vi.mock('../../src/engine/create-editor', () => ({
  createEditor: vi.fn(({ container }: CreateEditorOptions) => stub.openEditor!(container)),
  OPEN_TIMEOUT_MS: 1_000,
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
    markRestored: () => {},
    hold: () => () => {},
    isHeld: false,
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
  createCommandAdapter: (ui: object) => stub.adapters.get(ui)!,
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
  forgetLastDocument: async () => {},
  loadAutosaveEnabled: async () => true,
  loadSessionRecord: async () => stub.sessionRecord,
  saveSessionRecord: async () => {},
  saveAutosaveEnabled: async () => {},
  loadRulerVisible: async () => false,
  saveRulerVisible: async () => {},
  loadSpellcheckEnabled: async () => false,
  loadListAutoformatEnabled: async () => true,
  saveListAutoformatEnabled: async () => {},
  loadNavigationPaneVisible: async () => false,
  saveNavigationPaneVisible: async () => {},
  loadAutoCorrect: async () => null,
  loadNewDocumentDefaults: async () => null,
  saveNewDocumentDefaults: async () => {},
  saveAutoCorrect: async () => {},
  saveSpellcheckEnabled: async () => {},
  loadCanvasColor: async () => null,
  loadPasteDefault: async () => null,
  savePasteDefault: async () => {},
  saveCanvasColor: async () => {},
  loadSpellcheckWords: async () => [],
  saveSpellcheckWords: async () => {},
  loadRecentDocuments: async () => null,
  saveRecentDocuments: async () => {},
  loadDiscardBackups: async () => null,
  saveDiscardBackups: async () => {},
  loadCustomShortcuts: async () => null,
  saveCustomShortcuts: async () => {},
  // העץ של „פתח מסמך” — `useOpenSources().init()` רץ בעלייה, ומוק בלי
  // הפונקציות האלה זורק מתוכו דחייה לא-מטופלת (כך נפל ה-CI של #78).
  loadDocFolders: async () => null,
  saveDocFolders: async () => {},
  loadLibraryCache: async () => null,
  saveLibraryCache: async () => {},
  loadOpenDialogPlace: async () => null,
  saveOpenDialogPlace: async () => {},
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

import {
  SHORTCUTS,
  shortcutsByGroup,
  type Shortcut,
} from '../../src/ui/shortcuts/registry';
import { emptySession } from '../../src/sessions/session-state';

const { default: App } = await import('../../src/App.vue');

autoUnmount();

let adapter: CommandDouble;
let superdoc: SuperdocDouble;

async function mountShell() {
  const wrapper = mount(App, { attachTo: document.body });
  await settle(12);
  expect(stub.resets, 'the initial document actually opened before testing focus').toBe(1);
  expect(wrapper.find('.status-load__skip').exists(), 'opening has finished').toBe(false);
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
  stub.sessionRecord = null;
  adapter = createCommandDouble({ states: { 'document-mode': { value: 'editing' } } });
  superdoc = createSuperdocDouble();
  stub.adapters = new WeakMap();
  let first = true;
  stub.openEditor = (container) => {
    // Each candidate owns its mode and subscriptions; it must not inherit the
    // previous editor's temporary viewing mode during a repeated open.
    const doc = first ? superdoc : createSuperdocDouble();
    const commands = first ? adapter : createCommandDouble();
    first = false;
    commands.setState('document-mode', { value: commands.getState('document-mode').value ?? 'editing' });
    doc.host.setDocumentMode = vi.fn(mode => commands.setState('document-mode', { value: mode }));
    const ui = { selection: { observe: () => () => {} } } as unknown as EditorSession['ui'];
    stub.adapters.set(ui, commands);
    const disposers: Array<() => void> = [];
    let destroyed = false;
    return {
      superdoc: doc.host, ui, container,
      onDispose(dispose) {
        if (destroyed) dispose();
        else disposers.push(dispose);
      },
      destroy() {
        if (destroyed) return;
        destroyed = true;
        for (const dispose of disposers.splice(0)) dispose();
      },
    };
  };
});


const HELP = { code: 'Slash', ctrlKey: true };

/** פותח את הדיאלוג ומחזיר את המעטפת. */
async function openHelp() {
  const wrapper = await mountShell();
  press(HELP);
  await settle();
  return wrapper;
}

describe('דיאלוג הקיצורים', () => {
  it('Ctrl+/ פותח, ו-Escape סוגר', async () => {
    const wrapper = await openHelp();
    expect(wrapper.find('.shortcuts-dialog').exists()).toBe(true);

    press({ code: 'Escape' });
    await settle();

    expect(wrapper.find('.shortcuts-dialog').exists()).toBe(false);
  });

  it('Ctrl+/ הוא מתג — אותו צירוף גם סוגר', async () => {
    // בלי `inModal` הרשומה נחסמה ברגע שהדיאלוג נפתח, כלומר הצירוף פתח
    // בלבד — בניגוד למה שמשתמש מצפה ממקש שהוא זה עתה לחץ.
    const wrapper = await openHelp();
    expect(wrapper.find('.shortcuts-dialog').exists()).toBe(true);

    press(HELP);
    await settle();

    expect(wrapper.find('.shortcuts-dialog').exists()).toBe(false);
  });

  it('Ctrl+/ מעל דיאלוג אחר אינו פותח חלון שני, ואינו נבלע', async () => {
    const wrapper = await mountShell();
    press({ code: 'Escape' });
    await settle();
    // „אודות” נפתח מהרצועה; הוא `aria-modal`, כמו דיאלוג הקישור.
    const tab = wrapper.findAll('[role="tab"]').find((item) => item.text() === 'קובץ');
    await tab!.trigger('click');
    await settle();
    const about = wrapper
      .findAll('button')
      .find((node) => tipOf(node).title.startsWith('אודות'));
    await about!.trigger('click');
    await settle();
    expect(wrapper.find('.about-dialog').exists(), '„אודות” פתוח').toBe(true);

    const event = press(HELP);
    await settle();

    expect(wrapper.find('.shortcuts-dialog').exists(), 'לא נפתח חלון שני').toBe(false);
    expect(wrapper.find('.about-dialog').exists(), '„אודות” נשאר').toBe(true);
    expect(event.defaultPrevented, 'הצירוף לא נבלע').toBe(false);
  });

  it('הלוכסן של הספרון פותח גם הוא', async () => {
    // למשתמש זה „אותו מקש”; רק הדפדפן יודע שאלה שני code שונים.
    const wrapper = await mountShell();

    press({ code: 'NumpadDivide', ctrlKey: true });
    await settle();

    expect(wrapper.find('.shortcuts-dialog').exists()).toBe(true);
  });

  /** הרצועה מרכיבה לשונית רק כשהיא פעילה, ולכן צריך לפתוח את „קובץ” קודם. */
  async function shortcutsButton(wrapper: ReturnType<typeof mount>) {
    const tab = wrapper.findAll('[role="tab"]').find((item) => item.text() === 'קובץ');
    await tab!.trigger('click');
    await settle();
    return wrapper
      .findAll('button')
      .find((node) => tipMessage(node).startsWith('רשימת קיצורי המקלדת'));
  }

  it('יש כפתור ברצועה שפותח את הרשימה', async () => {
    // דיאלוג שמגיעים אליו רק בקיצור הוא דיאלוג שאיש לא ימצא — והוא **כל
    // הרשימה** של הקיצורים, כלומר בדיוק מה שמי שאינו יודע אותם מחפש.
    const wrapper = await mountShell();

    const button = await shortcutsButton(wrapper);

    expect(button, 'הכפתור קיים בלשונית „קובץ”').toBeDefined();
    await button!.trigger('click');
    await settle();

    expect(wrapper.find('.shortcuts-dialog').exists()).toBe(true);
  });

  it('ה-tooltip של הכפתור מלמד את הצירוף, מהרג׳יסטרי', async () => {
    const wrapper = await mountShell();

    const button = await shortcutsButton(wrapper);

    expect(tipOf(button!).shortcut).toBe('Ctrl+/');
  });

  it('Ctrl+/ בפריסה עברית פותח גם הוא', async () => {
    // אותו מקש פיזי מפיק „.” בפריסה העברית.
    const wrapper = await mountShell();

    press({ code: 'Slash', key: '.', ctrlKey: true });
    await settle();

    expect(wrapper.find('.shortcuts-dialog').exists()).toBe(true);
  });

  it('מציג את כל הרשומות שברג׳יסטרי — הספירה זהה', async () => {
    const wrapper = await openHelp();

    expect(wrapper.findAll('.shortcut-combo')).toHaveLength(SHORTCUTS.length);
  });

  it('לכל רשומה יש שורה משלה — לא רק המספר נכון', async () => {
    // כלומר: רשומה חדשה ברג׳יסטרי מופיעה כאן בלי לגעת בקומפוננטה.
    const wrapper = await openHelp();
    const shown = wrapper.findAll('.shortcut-combo').map((node) => node.text());

    const missing = (SHORTCUTS as readonly Shortcut[])
      .filter((shortcut) => !shown.includes(shortcut.label))
      .map((shortcut) => shortcut.id);

    expect(missing).toEqual([]);
  });

  it('התיאורים הם אלה שברג׳יסטרי', async () => {
    const wrapper = await openHelp();
    const text = wrapper.find('.shortcuts-body').text();

    for (const shortcut of SHORTCUTS as readonly Shortcut[]) {
      expect(text, shortcut.id).toContain(shortcut.description);
    }
  });

  it('מקובץ לפי תחום, בסדר של הרג׳יסטרי', async () => {
    const wrapper = await openHelp();
    const titles = wrapper.findAll('.shortcuts-group-title').map((node) => node.text());

    expect(titles).toEqual(shortcutsByGroup().map((group) => group.title));
  });

  it('הצירוף מוצג משמאל לימין', async () => {
    // בהקשר RTL תווית מעורבת כמו „Ctrl + Shift ימני” נשברת בלי dir מפורש.
    const wrapper = await openHelp();
    const combos = wrapper.findAll('.shortcut-combo kbd');

    expect(combos.length).toBeGreaterThan(0);
    expect(combos.every((node) => node.attributes('dir') === 'ltr')).toBe(true);
  });

  it('נגיש: role, aria-modal ושם', async () => {
    const wrapper = await openHelp();
    const dialog = wrapper.find('.shortcuts-dialog');

    expect(dialog.attributes('role')).toBe('dialog');
    expect(dialog.attributes('aria-modal')).toBe('true');
    const labelledBy = dialog.attributes('aria-labelledby');
    expect(labelledBy).toBeTruthy();
    expect(wrapper.find(`#${labelledBy}`).text()).toBe('קיצורי מקלדת');
  });

  it('המיקוד נכנס לדיאלוג, וחוזר למי ששלח אותו', async () => {
    const wrapper = await mountShell();
    const trigger = wrapper.findAll('button')[0]!.element as HTMLElement;
    trigger.focus();

    press(HELP);
    await settle();
    expect(wrapper.find('.shortcuts-dialog').element.contains(document.activeElement)).toBe(true);

    press({ code: 'Escape' });
    await settle();

    expect(document.activeElement).toBe(trigger);
  });

  it('כשהוא פתוח, קיצורי המסמך אינם רצים', async () => {
    // `aria-modal` הוא הצהרה שמה שמאחור אינו קיים. קיצור שממשיך לעבוד מכחיש אותה.
    await openHelp();
    adapter.calls.length = 0;

    press({ code: 'KeyB', ctrlKey: true });
    await settle();

    expect(adapter.calls).toEqual([]);
  });
});

describe('מעגל המיקוד', () => {
  /** האזור שבו הפוקוס כרגע, לפי המחלקות של המעטפת. */
  function focusedRegion(wrapper: ReturnType<typeof mount>): string | null {
    const active = document.activeElement;
    if (!active) return null;
    for (const selector of [
      '.word-titlebar',
      '.word-ribbon-container',
      '.editor-stack',
      '.word-statusbar',
    ]) {
      const region = wrapper.find(selector);
      if (region.exists() && region.element.contains(active)) return selector;
    }
    return null;
  }

  /** משטח ההקלדה של המנוע, כדי שלאזור המסמך יהיה למה למקד. */
  function composingSurface(wrapper: ReturnType<typeof mount>): HTMLTextAreaElement {
    const surface = document.createElement('textarea');
    wrapper.find('.editor-stack').element.appendChild(surface);
    return surface;
  }

  /**
   * הכפיל מקליט את `focus` אבל אינו מזיז מיקוד. כאן הוא נקשר למשטח ההקלדה,
   * כלומר להתנהגות האמיתית: `superdoc.focus()` מחזיר את הסמן לטקסט. בלי זה
   * הבדיקה הייתה מודדת את הכפיל ולא את המעגל.
   */
  function withEngineFocus(wrapper: ReturnType<typeof mount>): HTMLTextAreaElement {
    const surface = composingSurface(wrapper);
    const host = superdoc.host as unknown as { focus?: (options?: unknown) => void };
    const recorded = host.focus?.bind(host);
    host.focus = (options?: unknown) => {
      recorded?.(options);
      surface.focus();
    };
    return surface;
  }

  it('F6 עובר בין ארבעת האזורים ומקיף חזרה', async () => {
    const wrapper = await mountShell();
    withEngineFocus(wrapper);
    (wrapper.find('.word-ribbon-container button').element as HTMLElement).focus();

    const visited: Array<string | null> = [];
    for (let step = 0; step < 4; step += 1) {
      press({ code: 'F6' });
      await settle();
      visited.push(focusedRegion(wrapper));
    }

    expect(visited).toEqual([
      '.editor-stack',
      '.word-statusbar',
      '.word-titlebar',
      '.word-ribbon-container',
    ]);
  });

  it('Shift+F6 עובר בכיוון ההפוך', async () => {
    const wrapper = await mountShell();
    withEngineFocus(wrapper);
    (wrapper.find('.word-ribbon-container button').element as HTMLElement).focus();

    press({ code: 'F6', shiftKey: true });
    await settle();

    expect(focusedRegion(wrapper)).toBe('.word-titlebar');
  });

  it('המסמך ממוקד דרך המנוע, ולא דרך המארח', async () => {
    // מיקוד ה-`<main>` מזיז פוקוס אבל אינו מחזיר את הסמן לטקסט.
    const wrapper = await mountShell();
    (wrapper.find('.word-ribbon-container button').element as HTMLElement).focus();
    superdoc.reset();

    press({ code: 'F6' });
    await settle();

    expect(superdoc.ops()).toContain('focus');
  });

  it('F6 בתוך שדה טקסט של הרצועה אינו נבלע', async () => {
    // שם המקש שייך לשדה, ולא לנו.
    await mountShell();
    const field = document.createElement('input');
    document.body.appendChild(field);

    const event = new KeyboardEvent('keydown', { code: 'F6', cancelable: true, bubbles: true });
    field.dispatchEvent(event);
    await settle();
    field.remove();

    expect(event.defaultPrevented).toBe(false);
  });

  it('Escape מהרצועה מחזיר את הפוקוס למסמך', async () => {
    const wrapper = await mountShell();
    (wrapper.find('.word-ribbon-container button').element as HTMLElement).focus();
    superdoc.reset();

    const event = press({ code: 'Escape' });
    await settle();

    expect(superdoc.ops()).toContain('focus');
    expect(event.defaultPrevented).toBe(true);
  });

  it('Escape כשהפוקוס כבר במסמך אינו נבלע', async () => {
    // אחרת היינו בולעים את ה-Escape של המנוע.
    const wrapper = await mountShell();
    composingSurface(wrapper).focus();

    const event = press({ code: 'Escape' });
    await settle();

    expect(event.defaultPrevented).toBe(false);
  });
});

describe('מצב מיקוד', () => {
  /** משטח ההקלדה של המנוע, כדי שלאזור המסמך יהיה למה למקד. */
  function surface(wrapper: ReturnType<typeof mount>): HTMLTextAreaElement {
    const element = document.createElement('textarea');
    wrapper.find('.editor-stack').element.appendChild(element);
    return element;
  }

  it('הכניסה מתחילה עם הלוח העליון פתוח', async () => {
    // מצב מיקוד שמעלים את כל הפקדים ברגע ההפעלה נקרא כתקלה ולא כבחירה. הלוח
    // נשאר עד תנועת העכבר הראשונה אל גוף המסמך.
    const wrapper = await mountShell();
    press({ code: 'F11' });
    await settle();

    expect(wrapper.find('.word-app-shell').classes()).toContain('reveal-top');
  });

  it('כפתור היציאה מופיע רק במצב מיקוד — ומכבה אותו', async () => {
    // `Esc` ו-`F11` עובדים, אבל שניהם דורשים לדעת אותם. זה הפקד היחיד שרואים.
    const wrapper = await mountShell();
    expect(wrapper.find('.focus-exit').exists(), 'מחוץ למצב מיקוד').toBe(false);

    press({ code: 'F11' });
    await settle();
    const exit = wrapper.find('.focus-exit');
    expect(exit.exists(), 'במצב מיקוד').toBe(true);
    expect(exit.attributes('aria-label')).toBe('יציאה ממצב מיקוד');
    // „יציאה ממצב מיקוד, לחוץ” — זה מה שקורא מסך היה מכריז. `aria-pressed`
    // מתאר מתג, והכפתור הזה אינו מתג אלא פעולה חד-כיוונית.
    expect(exit.attributes('aria-pressed'), 'אינו מוכרז כלחוץ').toBeUndefined();

    await exit.trigger('click');
    await settle();

    expect(wrapper.find('.word-app-shell').classes()).not.toContain('focus-mode');
    expect(wrapper.find('.focus-exit').exists(), 'אחרי היציאה').toBe(false);
  });

  it('מעבר טאב במצב מיקוד מחזיר את ההקלדה למסמך', async () => {
    // נמדד: הפוקוס נשאר על כפתור הטאב, והלוח נסגר מיד אחרי הלחיצה —
    // `visibility: hidden` מבריח אותו ל-`<body>` (`active: BODY`) וההקלדה
    // נעלמת בלי סימן.
    const wrapper = await mountShell();
    await wrapper.find('.word-doctabs-new').trigger('click');
    await settle(12);
    expect(wrapper.findAll('.word-doctab').length, 'שני טאבים').toBe(2);

    press({ code: 'F11' });
    await settle();
    superdoc.reset();

    await wrapper.findAll('.word-doctab')[0]!.trigger('click');
    await settle();

    expect(superdoc.ops(), 'המסמך קיבל את הפוקוס').toContain('focus');
  });

  it('מחוץ למצב מיקוד מעבר טאב אינו חוטף את הפוקוס', async () => {
    // הבקרה, ולא סתם: שם רצועת הטאבים נשארת על המסך, ופוקוס שממשיך לשבת על
    // הטאב שנלחץ הוא ההתנהגות הנכונה — וגם מה שמאפשר לעבור טאבים בחצים.
    const wrapper = await mountShell();
    await wrapper.find('.word-doctabs-new').trigger('click');
    await settle(12);
    superdoc.reset();

    await wrapper.findAll('.word-doctab')[0]!.trigger('click');
    await settle();

    expect(superdoc.ops()).not.toContain('focus');
  });

  it('מצב מיקוד ששוחזר מהפעלה קודמת מתחיל עם הלוח פתוח', async () => {
    // `toggleFocusMode` מציב `revealed = 'top'` בכניסה, ו-
    // `applyShellPreferences` לא — כלומר עלייה עם ההעדפה דלוקה נתנה בדיוק את
    // מה שהכניסה הידנית הוגדרה כתקלה: כל הפקדים נעלמים בבת אחת.
    const record = emptySession();
    record.view.focusMode = true;
    stub.sessionRecord = record;

    const wrapper = await mountShell();

    const classes = wrapper.find('.word-app-shell').classes();
    expect(classes, 'מצב מיקוד שוחזר').toContain('focus-mode');
    expect(classes, 'והלוח פתוח, כמו בכניסה הידנית').toContain('reveal-top');
  });

  it('Escape יוצא ממצב מיקוד', async () => {
    // המקש הראשון שכל משתמש מנסה. בלעדיו היציאה היחידה היא למצוא שוב את F11
    // או לרחף מעל קצה המסך.
    const wrapper = await mountShell();
    press({ code: 'F11' });
    await settle();
    expect(wrapper.find('.word-app-shell').classes()).toContain('focus-mode');

    const event = press({ code: 'Escape' });
    await settle();

    expect(wrapper.find('.word-app-shell').classes()).not.toContain('focus-mode');
    expect(event.defaultPrevented).toBe(true);
  });

  it('Escape מתוך דיאלוג החיפוש סוגר אותו בלבד', async () => {
    // הדיאלוג מטפל ב-Escape בעצמו ומאפס את המצב סינכרונית. בלי `.stop`
    // האירוע ממשיך ל-window, ושם `closeTopmost` כבר אינו רואה חיפוש פתוח —
    // ונופל לענף הבא. עד שנוסף מצב המיקוד הנפילה הזאת הייתה בלתי מזיקה.
    const wrapper = await mountShell();
    press({ code: 'F11' });
    press({ code: 'KeyF', ctrlKey: true });
    await settle();
    const dialog = wrapper.find('.find-replace-dialog');
    expect(dialog.exists(), 'החיפוש נפתח').toBe(true);
    expect(wrapper.find('.word-app-shell').classes()).toContain('focus-mode');

    dialog.element.dispatchEvent(
      new KeyboardEvent('keydown', { code: 'Escape', key: 'Escape', cancelable: true, bubbles: true }),
    );
    await settle();

    expect(wrapper.find('.find-replace-dialog').exists(), 'החיפוש נסגר').toBe(false);
    expect(
      wrapper.find('.word-app-shell').classes(),
      'מצב המיקוד נשאר — Escape אחד, פעולה אחת',
    ).toContain('focus-mode');
  });

  it('דיאלוג פתוח נסגר לפני מצב המיקוד, ולא יחד איתו', async () => {
    const wrapper = await mountShell();
    press({ code: 'F11' });
    press({ code: 'Slash', ctrlKey: true });
    await settle();

    press({ code: 'Escape' });
    await settle();

    expect(wrapper.find('.shortcuts-dialog').exists()).toBe(false);
    expect(wrapper.find('.word-app-shell').classes(), 'מצב המיקוד נשאר').toContain('focus-mode');
  });

  /**
   * ואותה טענה בדיוק, במסלול שבאמת קורה בדפדפן.
   *
   * הבדיקה שמעל עוברת ב-jsdom **ושקרית בכרום**: ל-jsdom אין מסך מלא, ולכן
   * ה-`keydown` שלה מגיע לדף. בכרום אמיתי `Escape` שמשמש ליציאה ממסך מלא
   * **נבלע** — נמדד `keys: []` ורק `fullscreenchange: [false]` — וכל מה
   * שהמעטפת מקבלת הוא האירוע הזה. לכן הדימוי כאן הוא `fullscreenchange`
   * ולא מקש: אחרת הטענה נבדקת במסלול שאינו קיים.
   */
  /**
   * ל-jsdom אין מסך מלא כלל — `document.fullscreenElement` אינו קיים אפילו
   * כתכונה — ולכן „יציאה” חייבת לבוא אחרי „כניסה”: `watchFullscreen` מדווח על
   * **מעבר** ולא על אירוע, מפני ששני שמות האירוע יכולים לירות שניהם על אותה
   * יציאה אחת והקורא סוגר שכבה בכל קריאה (ראו composables/window-fullscreen.ts).
   * דימוי שיורה „יצאנו” בלי שנכנסנו מודד מסלול שאין לו מקבילה בדפדפן.
   */
  function setFullscreenElement(value: Element | null): void {
    Object.defineProperty(document, 'fullscreenElement', { configurable: true, value });
  }

  function enterFullscreenFromOutside(): void {
    setFullscreenElement(document.documentElement);
    document.dispatchEvent(new Event('fullscreenchange'));
  }

  function exitFullscreenFromOutside(): void {
    setFullscreenElement(null);
    document.dispatchEvent(new Event('fullscreenchange'));
  }

  it('יציאה ממסך מלא עם דיאלוג פתוח סוגרת את הדיאלוג ומשאירה את מצב המיקוד', async () => {
    // נמדד לפני התיקון: `Ctrl+F` ואחריו `Escape` אחד נתנו
    // `{focus:false, fs:false, find:true}` — הדיאלוג נשאר פתוח ומצב המיקוד
    // כבה, בדיוק ההיפך משתי הבדיקות שמעל.
    const wrapper = await mountShell();
    press({ code: 'F11' });
    press({ code: 'KeyF', ctrlKey: true });
    await settle();
    expect(wrapper.find('.find-replace-dialog').exists(), 'החיפוש נפתח').toBe(true);
    // אחרי הכניסה למצב מיקוד, שגם היא ממקדת את המסמך: בלי האיפוס הטענה על
    // הפוקוס שלמטה הייתה מתקיימת מאליה.
    superdoc.reset();

    enterFullscreenFromOutside();
    await settle();
    exitFullscreenFromOutside();
    await settle();

    expect(wrapper.find('.find-replace-dialog').exists(), 'החיפוש נסגר').toBe(false);
    expect(
      wrapper.find('.word-app-shell').classes(),
      'ומצב המיקוד נשאר — מיקוד בלי מסך מלא הוא מצב נתמך',
    ).toContain('focus-mode');
    // ושדה החיפוש שהוסר לא לקח איתו את ההקלדה: בלי זה הפוקוס נשאר על
    // ה-`<body>` (נמדד `active: BODY`), ומצב מיקוד שאי אפשר להקליד בו.
    expect(superdoc.ops(), 'הפוקוס חזר למסמך').toContain('focus');
  });

  it('יציאה ממסך מלא בלי שכבה פתוחה כן מכבה את מצב המיקוד', async () => {
    // הבקרה: בלעדיה התיקון שמעל היה יכול לנטרל את המאזין לגמרי, ולהשאיר
    // מעטפת בלי פסים בתוך חלון רגיל אחרי `F11` של הדפדפן.
    const wrapper = await mountShell();
    press({ code: 'F11' });
    await settle();
    expect(wrapper.find('.word-app-shell').classes()).toContain('focus-mode');

    enterFullscreenFromOutside();
    await settle();
    exitFullscreenFromOutside();
    await settle();

    expect(wrapper.find('.word-app-shell').classes()).not.toContain('focus-mode');
  });

  it('שני שמות האירוע על אותה יציאה סוגרים שכבה אחת, לא שתיים', async () => {
    // מאחז יכול לחשוף גם `fullscreenchange` וגם `webkitfullscreenchange`
    // ולירות את שניהם. הקורא הוא `closeTopmostLayer`, כלומר קריאה שנייה
    // הייתה סוגרת גם את מצב המיקוד — היפוך סדר השכבות שהמאזין נכתב לתקן.
    const wrapper = await mountShell();
    press({ code: 'F11' });
    press({ code: 'KeyF', ctrlKey: true });
    await settle();
    enterFullscreenFromOutside();
    await settle();

    setFullscreenElement(null);
    document.dispatchEvent(new Event('fullscreenchange'));
    document.dispatchEvent(new Event('webkitfullscreenchange'));
    await settle();

    expect(wrapper.find('.find-replace-dialog').exists(), 'החיפוש נסגר').toBe(false);
    expect(
      wrapper.find('.word-app-shell').classes(),
      'ומצב המיקוד שרד — האירוע השני אינו מעבר',
    ).toContain('focus-mode');
  });

  /**
   * סוגר את הלוח שהכניסה למצב מיקוד פותחת.
   *
   * `pointerleave` על השורש — בדיוק מה שקורה כשהמצביע יוצא מהחלון — ולא השמה
   * ל-state פנימי: מה שנבדק הוא ההתנהגות של המעטפת, וזו הדרך שיש למשתמש.
   */
  async function hideRevealed(wrapper: ReturnType<typeof mount>): Promise<void> {
    await wrapper.find('.word-app-shell').trigger('pointerleave');
    await settle();
  }

  it('F6 במצב מיקוד אינו ממקד פקד בלתי נראה', async () => {
    // הרצועה ושורת המצב יוצאות מהזרימה ומוזזות אל מחוץ למסך — הן עדיין בעץ.
    // בלי סימון הזמינות המשתמש היה מקבל טבעת מיקוד על פס שאינו על המסך,
    // הקלדה שאינה מגיעה למסמך, ו-Enter שמפעיל כפתור שאינו נראה.
    const wrapper = await mountShell();
    const typing = surface(wrapper);
    typing.focus();
    press({ code: 'F11' });
    await settle();
    // הכניסה מתחילה עם הלוח פתוח, ופס פתוח **הוא** זמין (ראו הבדיקה למטה).
    // מה שנבדק כאן הוא המצב שאחרי שהוא נסגר.
    await hideRevealed(wrapper);

    press({ code: 'F6' });
    await settle();

    const ribbon = wrapper.find('.word-ribbon-container').element;
    const statusbar = wrapper.find('.word-statusbar').element;
    expect(ribbon.contains(document.activeElement), 'הרצועה').toBe(false);
    expect(statusbar.contains(document.activeElement), 'שורת המצב').toBe(false);
  });

  it('F6 במצב מיקוד מגיע לכפתור היציאה — הפקד היחיד שעל המסך', async () => {
    // המנוע בולע `Tab`, ולכן בלי הכפתור כאזור במעגל אין אליו שום דרך מהמקלדת
    // (נמדד בדפדפן), וכלל ה-`:focus-visible` שלו היה קוד מת.
    const wrapper = await mountShell();
    surface(wrapper).focus();
    press({ code: 'F11' });
    await settle();
    await hideRevealed(wrapper);

    const event = press({ code: 'F6' });
    await settle();

    expect(document.activeElement, 'הפוקוס על הכפתור').toBe(wrapper.find('.focus-exit').element);
    expect(event.defaultPrevented, 'ולכן המקש כן נבלע — יש לאן לעבור').toBe(true);
  });

  it('פס שנחשף חוזר להיות נגיש ל-F6', async () => {
    // מה ש-`FocusRegion.isAvailable` מבטיח בתיעוד שלו. נמדד לפני התיקון:
    // רצועה פרושה על כל המסך (הכניסה מתחילה ב-`reveal-top`) לא הייתה נגישה
    // כלל — הפוקוס נשאר במשטח ההקלדה של המנוע.
    const wrapper = await mountShell();
    surface(wrapper).focus();
    press({ code: 'F11' });
    await settle();
    expect(wrapper.find('.word-app-shell').classes(), 'הלוח פתוח').toContain('reveal-top');

    // מסמך → כפתור היציאה → (שורת המצב מדולגת, היא אינה חשופה) → סרגל הכותרת.
    press({ code: 'F6' });
    await settle();
    press({ code: 'F6' });
    await settle();

    expect(
      wrapper.find('.word-titlebar').element.contains(document.activeElement),
      'סרגל הכותרת שנחשף',
    ).toBe(true);
  });

  it('מחוץ למצב מיקוד F6 כן מגיע לרצועה', async () => {
    // הבקרה: בלעדיה הבדיקה שמעליה הייתה עוברת גם אם F6 הפסיק לעבוד לגמרי.
    const wrapper = await mountShell();
    surface(wrapper).focus();

    press({ code: 'F6' });
    await settle();

    expect(wrapper.find('.word-statusbar').element.contains(document.activeElement)).toBe(true);
  });
});

describe('סרגל הכותרת במעגל', () => {
  it('Escape משדה שם המסמך מחזיר את הפוקוס למסמך', async () => {
    // בלי סרגל הכותרת כאזור, `current()` החזיר null ומי שהגיע לשדה השם
    // ב-Tab נשאר תקוע: F6 חסום שם בכוונה, ו-Escape לא עשה דבר.
    const wrapper = await mountShell();
    const field = wrapper.find('.word-titlebar input');
    expect(field.exists(), 'יש שדה טקסט בסרגל הכותרת').toBe(true);
    (field.element as HTMLElement).focus();
    superdoc.reset();

    press({ code: 'Escape' });
    await settle();

    expect(superdoc.ops()).toContain('focus');
  });
});
