/**
 * המנתב — שלוש ההכרעות שקודם היו פזורות ב-App.vue ולכן לא נבדקו: פוקוס בשדה
 * טקסט, דיאלוג מודאלי פתוח, ומתי מותר לבלוע את ההתנהגות של הדפדפן.
 */
import { describe, it, expect, vi } from 'vitest';
import {
  createShortcutDispatcher,
  isTextEntryTarget,
  type ShortcutDispatcherDeps,
} from '../../src/ui/shortcuts/dispatch';
import { SHORTCUTS, type Shortcut } from '../../src/ui/shortcuts/registry';

const BOLD: Shortcut = {
  id: 'bold',
  label: 'Ctrl+B',
  description: 'מודגש',
  group: 'font',
  code: 'KeyB',
  ctrl: true,
  command: 'bold',
};

const SAVE: Shortcut = {
  id: 'save',
  label: 'Ctrl+S',
  description: 'שמירה',
  group: 'file',
  code: 'KeyS',
  ctrl: true,
  action: 'save',
  inTextEntry: true,
};

const ESCAPE: Shortcut = {
  id: 'escape',
  label: 'Esc',
  description: 'סגירה',
  group: 'app',
  code: 'Escape',
  action: 'escape',
  inTextEntry: true,
  inModal: true,
};

const PASTE: Shortcut = {
  id: 'paste',
  label: 'Ctrl+V',
  description: 'הדבקה',
  group: 'clipboard',
  code: 'KeyV',
  ctrl: true,
  native: true,
};

/** פעולת מעטפת עם payload — רשומת `tab-goto` היא המקרה האמיתי. */
const TAB_GOTO: Shortcut = {
  id: 'tab-goto-3',
  label: 'Alt+3',
  description: 'מעבר לטאב השלישי',
  group: 'tabs',
  code: 'Digit3',
  alt: true,
  action: 'tab-goto',
  payload: 3,
};

const ALL = [BOLD, SAVE, ESCAPE, PASTE, TAB_GOTO];

/** יעד מזויף למאזין, כדי שהבדיקה תוכל לוודא שהוא נרשם ושהוא נותק. */
function fakeTarget() {
  const listeners: EventListener[] = [];
  return {
    count: () => listeners.length,
    fire: (event: unknown) => listeners.forEach((listener) => listener(event as Event)),
    addEventListener: (_type: string, listener: unknown) => {
      listeners.push(listener as EventListener);
    },
    removeEventListener: (_type: string, listener: unknown) => {
      const index = listeners.indexOf(listener as EventListener);
      if (index >= 0) listeners.splice(index, 1);
    },
  };
}

function setup(over: Partial<ShortcutDispatcherDeps> = {}) {
  const runCommand = vi.fn();
  const runAction = vi.fn(() => true);
  const target = fakeTarget();
  const dispatcher = createShortcutDispatcher({
    runCommand,
    runAction,
    shortcuts: ALL,
    target,
    ...over,
  });
  return { runCommand, runAction, target, dispatcher };
}

function event(over: Record<string, unknown> = {}): KeyboardEvent {
  return {
    key: '',
    code: '',
    ctrlKey: false,
    metaKey: false,
    shiftKey: false,
    altKey: false,
    target: null,
    preventDefault: vi.fn(),
    stopPropagation: vi.fn(),
    ...over,
  } as unknown as KeyboardEvent;
}

function element(tag: string, attributes: Record<string, string> = {}): HTMLElement {
  const node = document.createElement(tag);
  for (const [name, value] of Object.entries(attributes)) node.setAttribute(name, value);
  return node;
}

describe('isTextEntryTarget', () => {
  it('שדות הטקסט של הממשק שלנו', () => {
    expect(isTextEntryTarget(element('input'))).toBe(true);
    expect(isTextEntryTarget(element('textarea'))).toBe(true);
    expect(isTextEntryTarget(element('select'))).toBe(true);
    expect(isTextEntryTarget(element('div', { role: 'textbox' }))).toBe(true);
  });

  it('אזור המסמך של המנוע אינו שדה טקסט', () => {
    // דווקא שם קיצורי המסמך חייבים לעבוד — זה כל העניין.
    expect(isTextEntryTarget(element('div'))).toBe(false);
    expect(isTextEntryTarget(element('button'))).toBe(false);
    expect(isTextEntryTarget(null)).toBe(false);
  });
});

describe('משטח ההקלדה של המנוע', () => {
  // הבאג שנמדד בדפדפן אמיתי: המנוע מקבל הקשות דרך `<textarea>` נסתר ברוחב
  // פיקסל אחד בתוך אזור המסמך, ולכן `event.target` של כל הקשה בזמן הקלדה הוא
  // TEXTAREA. בדיקת ה-tag לבדה חסמה את **כל** הקיצורים בדיוק כשהפוקוס במסמך.
  const composing = element('textarea');

  it('שדה טקסט בתוך אזור המסמך אינו חוסם', () => {
    const { dispatcher, runCommand } = setup({ isDocumentSurface: () => true });

    dispatcher.handle(event({ code: 'KeyB', ctrlKey: true, target: composing }));

    expect(runCommand).toHaveBeenCalledExactlyOnceWith('bold', undefined);
  });

  it('אותו שדה מחוץ לאזור המסמך כן חוסם', () => {
    const { dispatcher, runCommand } = setup({ isDocumentSurface: () => false });

    dispatcher.handle(event({ code: 'KeyB', ctrlKey: true, target: composing }));

    expect(runCommand).not.toHaveBeenCalled();
  });

  it('בלי ההכרעה הזאת ברירת המחדל היא לחסום — כדי ששדה החיפוש יישאר שלו', () => {
    const { dispatcher, runCommand } = setup();

    dispatcher.handle(event({ code: 'KeyB', ctrlKey: true, target: element('input') }));

    expect(runCommand).not.toHaveBeenCalled();
  });
});

describe('עורך פעיל שננעל לפתיחה', () => {
  it('חוסם פקודות ופעולות משני שלבי הניתוב, ומאפשר אותן שוב בשחרור', () => {
    let blocked = true;
    const { dispatcher, runCommand, runAction } = setup({
      shortcuts: SHORTCUTS, isDocumentBlocked: () => blocked,
    });
    const keys = [
      { code: 'KeyB', ctrlKey: true },
      { code: 'Enter', ctrlKey: true },
      { code: 'KeyE', ctrlKey: true, shiftKey: true },
    ];
    for (const init of keys) {
      const key = event(init);
      expect(dispatcher.handleCapture(key) || dispatcher.handle(key)).toBe(true);
      expect(key.preventDefault).toHaveBeenCalled();
      expect(key.stopPropagation).toHaveBeenCalled();
    }
    expect(runCommand).not.toHaveBeenCalled();
    expect(runAction).not.toHaveBeenCalled();
    blocked = false;
    dispatcher.handle(event(keys[0]));
    dispatcher.handle(event(keys[1]));
    dispatcher.handleCapture(event(keys[2]));
    expect(runCommand).toHaveBeenCalledWith('bold', undefined);
    expect(runAction).toHaveBeenCalledWith('page-break', undefined);
    expect(runAction).toHaveBeenCalledWith('track-changes', undefined);
  });

  it('משאיר Escape/F6 ו־Tab, הדבקה ושדות UI לפי החוזה המקורי', () => {
    const { dispatcher, runAction, runCommand } = setup({
      shortcuts: SHORTCUTS, isDocumentBlocked: () => true,
    });
    for (const init of [{ code: 'Escape' }, { code: 'F6' }, { code: 'F6', shiftKey: true }]) {
      expect(dispatcher.handle(event(init))).toBe(true);
    }
    expect(runAction.mock.calls).toEqual([
      ['escape', undefined], ['focus-next-region', undefined], ['focus-prev-region', undefined],
    ]);
    for (const init of [{ code: 'KeyQ', altKey: true }, { code: 'F8', altKey: true }]) {
      expect(dispatcher.handleCapture(event(init))).toBe(true);
    }
    expect(runAction).toHaveBeenCalledTimes(3);
    for (const init of [
      { code: 'Tab' },
      { code: 'KeyV', ctrlKey: true },
      { code: 'KeyB', ctrlKey: true, target: element('input') },
      { code: 'Enter', ctrlKey: true, target: element('textarea') },
    ]) {
      const key = event(init);
      expect(dispatcher.handleCapture(key) || dispatcher.handle(key)).toBe(false);
      expect(key.preventDefault).not.toHaveBeenCalled();
    }
    expect(runCommand).not.toHaveBeenCalled();
  });
});

describe('המנתב', () => {
  it('צירוף מוכר: הפקודה רצה דרך האדפטר, וברירת המחדל נבלעת', () => {
    const { dispatcher, runCommand } = setup();
    const keyEvent = event({ code: 'KeyB', ctrlKey: true });

    expect(dispatcher.handle(keyEvent)).toBe(true);
    expect(runCommand).toHaveBeenCalledWith('bold', undefined);
    expect(keyEvent.preventDefault).toHaveBeenCalled();
  });

  it('פעולת מעטפת רצה בשמה', () => {
    const { dispatcher, runAction } = setup();

    expect(dispatcher.handle(event({ code: 'KeyS', ctrlKey: true }))).toBe(true);
    expect(runAction).toHaveBeenCalledWith('save', undefined);
  });

  it('פעולת מעטפת מקבלת את ה-payload של הרשומה', () => {
    // בלי זה שמונה רשומות `Alt+1`…`Alt+8` היו מריצות את אותה פעולה בלי שום
    // דרך לדעת לאיזה טאב לעבור — כלומר כולן היו עוברות לאותו אחד.
    //
    // `handleCapture` ולא `handle`: `Alt+3` הוא `Alt` בלי `Ctrl`, כלומר צירוף
    // שהמנוע בולע — ראו `runsBeforeEngine`.
    const { dispatcher, runAction } = setup();
    const keyEvent = event({ code: 'Digit3', altKey: true });

    expect(dispatcher.handleCapture(keyEvent)).toBe(true);
    expect(runAction).toHaveBeenCalledWith('tab-goto', 3);
    // ההקשה אינה ממשיכה למנוע, אחרת הספרה נכתבת במסמך במקום מעבר לשונית.
    expect(keyEvent.stopPropagation).toHaveBeenCalled();
  });

  it('רשומת `Alt` אינה רצה פעמיים — ב-bubble היא כבר לא שלה', () => {
    const { dispatcher, runAction } = setup();

    expect(dispatcher.handle(event({ code: 'Digit3', altKey: true }))).toBe(false);
    expect(runAction).not.toHaveBeenCalled();
  });

  it('`Ctrl+Alt` נשאר ב-bubble — המנוע אינו נוגע בו', () => {
    // הכלל אינו „יש Alt” אלא „Alt בלי Ctrl”, וזה מה שנמדד: `Ctrl+Alt+M` עבר
    // דרך המנוע בלי שנגע בו.
    const CTRL_ALT: Shortcut = {
      id: 'footnote',
      label: 'Ctrl+Alt+F',
      description: 'הערת שוליים',
      group: 'insert',
      code: 'KeyF',
      ctrl: true,
      alt: true,
      action: 'save',
    };
    const { dispatcher, runAction } = setup({ shortcuts: [CTRL_ALT] });

    expect(dispatcher.handleCapture(event({ code: 'KeyF', ctrlKey: true, altKey: true }))).toBe(false);
    expect(runAction).not.toHaveBeenCalled();
    expect(dispatcher.handle(event({ code: 'KeyF', ctrlKey: true, altKey: true }))).toBe(true);
    expect(runAction).toHaveBeenCalledWith('save', undefined);
  });

  it('צירוף לא מוכר אינו נבלע', () => {
    const { dispatcher, runCommand, runAction } = setup();
    const keyEvent = event({ code: 'KeyG', ctrlKey: true });

    expect(dispatcher.handle(keyEvent)).toBe(false);
    expect(keyEvent.preventDefault).not.toHaveBeenCalled();
    expect(runCommand).not.toHaveBeenCalled();
    expect(runAction).not.toHaveBeenCalled();
  });

  it('צירוף של הדפדפן מתועד אך אינו נבלע', () => {
    // preventDefault על Ctrl+V היה מבטל את ההדבקה עצמה.
    const { dispatcher, runCommand } = setup();
    const keyEvent = event({ code: 'KeyV', ctrlKey: true });

    expect(dispatcher.handle(keyEvent)).toBe(false);
    expect(keyEvent.preventDefault).not.toHaveBeenCalled();
    expect(runCommand).not.toHaveBeenCalled();
  });

  it('פוקוס בשדה טקסט חוסם קיצור מסמך', () => {
    const { dispatcher, runCommand } = setup();
    const keyEvent = event({ code: 'KeyB', ctrlKey: true, target: element('input') });

    expect(dispatcher.handle(keyEvent)).toBe(false);
    expect(keyEvent.preventDefault).not.toHaveBeenCalled();
    expect(runCommand).not.toHaveBeenCalled();
  });

  it('פוקוס בשדה טקסט אינו חוסם רשומה שהוכרזה כמותרת שם', () => {
    const { dispatcher, runAction } = setup();

    expect(dispatcher.handle(event({ code: 'KeyS', ctrlKey: true, target: element('input') }))).toBe(
      true,
    );
    expect(dispatcher.handle(event({ code: 'Escape', target: element('input') }))).toBe(true);
    expect(runAction).toHaveBeenNthCalledWith(1, 'save', undefined);
    expect(runAction).toHaveBeenNthCalledWith(2, 'escape', undefined);
  });

  it('אזור המסמך אינו חוסם — הקיצורים עובדים בתוכו', () => {
    const { dispatcher, runCommand } = setup();

    expect(dispatcher.handle(event({ code: 'KeyB', ctrlKey: true, target: element('div') }))).toBe(
      true,
    );
    expect(runCommand).toHaveBeenCalledWith('bold', undefined);
  });

  it('role=textbox נחשב שדה טקסט', () => {
    const { dispatcher } = setup();
    const target = element('div', { role: 'textbox' });

    expect(dispatcher.handle(event({ code: 'KeyB', ctrlKey: true, target }))).toBe(false);
  });

  it('דיאלוג מודאלי פתוח חוסם הכול חוץ מ-Escape', () => {
    const { dispatcher, runAction, runCommand } = setup({ isModalOpen: () => true });

    expect(dispatcher.handle(event({ code: 'KeyB', ctrlKey: true }))).toBe(false);
    expect(dispatcher.handle(event({ code: 'KeyS', ctrlKey: true }))).toBe(false);
    expect(dispatcher.handle(event({ code: 'Escape' }))).toBe(true);
    expect(runCommand).not.toHaveBeenCalled();
    expect(runAction).toHaveBeenCalledExactlyOnceWith('escape', undefined);
  });

  it('נרשם ליעד ומנותק ב-dispose', () => {
    // שניים: הרשומות המובנות ב-bubble, והרשימה האישית ב-capture.
    const { dispatcher, target, runCommand } = setup();
    expect(target.count()).toBe(2);

    target.fire(event({ code: 'KeyB', ctrlKey: true }));
    expect(runCommand).toHaveBeenCalledTimes(1);

    dispatcher.dispose();
    expect(target.count()).toBe(0);

    target.fire(event({ code: 'KeyB', ctrlKey: true }));
    expect(runCommand).toHaveBeenCalledTimes(1);
  });

  it('אירוע שכבר טופל אינו רץ שוב', () => {
    // המאזין יושב על window בשלב ה-bubble, כלומר אחרי ה-keymap של מנוע
    // העריכה. בלי הבדיקה הזאת צירוף שהמנוע קושר בעצמו היה מופעל פעמיים —
    // והמשתמש היה רואה הדגשה שמתבטלת מיד.
    const { dispatcher, runCommand } = setup();
    const keyEvent = event({ code: 'KeyB', ctrlKey: true, defaultPrevented: true });

    expect(dispatcher.handle(keyEvent)).toBe(false);
    expect(runCommand).not.toHaveBeenCalled();
  });

  it('פעולה שלא טיפלה אינה בולעת את ברירת המחדל', () => {
    // `Escape` בלי חלון פתוח: אם נבלע אותו, שברנו את ה-Escape של המנוע.
    const { dispatcher } = setup({ runAction: vi.fn(() => false) });
    const keyEvent = event({ code: 'Escape' });

    expect(dispatcher.handle(keyEvent)).toBe(false);
    expect(keyEvent.preventDefault).not.toHaveBeenCalled();
  });

  it('פעולה שטיפלה בולעת', () => {
    const { dispatcher } = setup({ runAction: vi.fn(() => true) });
    const keyEvent = event({ code: 'Escape' });

    expect(dispatcher.handle(keyEvent)).toBe(true);
    expect(keyEvent.preventDefault).toHaveBeenCalled();
  });

  it('dispose חוזר אינו נופל', () => {
    const { dispatcher } = setup();
    dispatcher.dispose();
    expect(() => dispatcher.dispose()).not.toThrow();
  });
});

describe('held capture chords', () => {
  for (const code of ['KeyE', 'Digit8']) {
    for (const modifier of ['ctrlKey', 'metaKey']) {
      it(`${modifier}+Shift+${code}: toggle once and consume repeated keydowns`, () => {
        const { dispatcher, runAction, runCommand } = setup({ shortcuts: SHORTCUTS });
        const chord = { code, [modifier]: true, shiftKey: true };
        expect(dispatcher.handleCapture(event(chord))).toBe(true);
        for (let i = 0; i < 3; i++) {
          const repeated = event({ ...chord, repeat: true });
          expect(dispatcher.handleCapture(repeated)).toBe(true);
          expect(repeated.preventDefault).toHaveBeenCalledOnce();
          expect(repeated.stopPropagation).toHaveBeenCalledOnce();
        }
        expect(runAction.mock.calls.length + runCommand.mock.calls.length).toBe(1);
      });
    }
  }

  for (const guard of [
    { isModalOpen: () => true },
    { target: element('input') },
    { isComposing: true },
    { keyCode: 229 },
    { defaultPrevented: true },
  ]) {
    it(`preserves repeat guards: ${Object.keys(guard)[0]}`, () => {
      const { isModalOpen, ...init } = guard;
      const { dispatcher, runAction } = setup({ shortcuts: SHORTCUTS, isModalOpen });
      const repeated = event({ code: 'KeyE', ctrlKey: true, shiftKey: true, repeat: true, ...init });
      expect(dispatcher.handleCapture(repeated)).toBe(false);
      expect(repeated.preventDefault).not.toHaveBeenCalled();
      expect(runAction).not.toHaveBeenCalled();
    });
  }

  it('still consumes repeats from the engine composition textarea', () => {
    const { dispatcher } = setup({ shortcuts: SHORTCUTS, isDocumentSurface: () => true });
    expect(dispatcher.handleCapture(event({ code: 'KeyE', ctrlKey: true, shiftKey: true, repeat: true, target: element('textarea') }))).toBe(true);
  });

  it('keeps built-in priority over repeatable custom shortcuts', () => {
    const runCustom = vi.fn(() => true);
    const custom: Shortcut = { ...BOLD, id: 'custom', code: 'KeyE', shift: true, repeatable: true };
    const { dispatcher } = setup({ shortcuts: SHORTCUTS, customShortcuts: () => [custom], runCustom });
    expect(dispatcher.handleCapture(event({ code: 'KeyE', ctrlKey: true, shiftKey: true, repeat: true }))).toBe(true);
    expect(runCustom).not.toHaveBeenCalled();
  });

  it('leaves equivalent engine bindings and native repeat handling alone', () => {
    const { dispatcher, runCommand } = setup({ shortcuts: SHORTCUTS });
    for (const code of ['KeyX', 'KeyZ', 'KeyV']) {
      const repeated = event({ code, ctrlKey: true, shiftKey: true, repeat: true });
      expect(dispatcher.handleCapture(repeated)).toBe(false);
      expect(repeated.preventDefault).not.toHaveBeenCalled();
    }
    expect(runCommand).not.toHaveBeenCalled();
  });

  it('continues running repeatable capture commands', () => {
    const { dispatcher, runAction } = setup({ shortcuts: [{ ...TAB_GOTO, repeatable: true }] });
    expect(dispatcher.handleCapture(event({ code: 'Digit3', altKey: true, repeat: true }))).toBe(true);
    expect(runAction).toHaveBeenCalledExactlyOnceWith('tab-goto', 3);
  });
});

describe('קיצורים אישיים', () => {
  /**
   * הרשומה האישית. `Ctrl+Alt+K` פנוי, ו-`Ctrl+B` **תפוס** — הרשומה השנייה
   * קיימת בדיוק כדי לוודא שגם כשהיא נכנסת לרשימה האישית היא אינה יורה.
   */
  const CUSTOM: Shortcut = {
    id: 'cs-1',
    label: 'Ctrl+Alt+K',
    description: 'כותרת קטע',
    group: 'custom',
    code: 'KeyK',
    ctrl: true,
    alt: true,
  };

  const SHADOWING_BOLD: Shortcut = {
    id: 'cs-2',
    label: 'Ctrl+B',
    description: 'ניסיון לדרוס מודגש',
    group: 'custom',
    code: 'KeyB',
    ctrl: true,
  };

  /**
   * `handled` הוא מה ש-`runCustom` יחזיר. פרמטר ולא override על ה-deps:
   * מסירת מרגל אחר דרך `over` הייתה משאירה את המרגל שמוחזר כאן בלי
   * חיווט — כלומר בדיקה שמאשרת „לא נקרא” על מרגל שממילא לא היה מחובר.
   */
  function withCustom(
    over: Partial<ShortcutDispatcherDeps> = {},
    handled = true,
    shortcuts: readonly Shortcut[] = [CUSTOM, SHADOWING_BOLD],
  ) {
    const runCustom = vi.fn(() => handled);
    const harness = setup({ customShortcuts: () => shortcuts, runCustom, ...over });
    return { ...harness, runCustom };
  }

  it('קיצור אישי אינו משנה עורך נעול, בלי להשתלט על שדה UI', () => {
    let blocked = true;
    const { dispatcher, runCustom } = withCustom({ isDocumentBlocked: () => blocked });
    const held = event({ code: 'KeyK', ctrlKey: true, altKey: true });
    expect(dispatcher.handleCapture(held)).toBe(true);
    expect(held.preventDefault).toHaveBeenCalled();
    expect(runCustom).not.toHaveBeenCalled();
    const field = event({ code: 'KeyK', ctrlKey: true, altKey: true, target: element('input') });
    expect(dispatcher.handleCapture(field)).toBe(false);
    expect(field.preventDefault).not.toHaveBeenCalled();
    blocked = false;
    expect(dispatcher.handleCapture(event({ code: 'KeyK', ctrlKey: true, altKey: true }))).toBe(true);
    expect(runCustom).toHaveBeenCalledExactlyOnceWith('cs-1');
  });

  it('צירוף פנוי מגיע לקיצור האישי ונבלע', () => {
    const { dispatcher, runCustom } = withCustom();
    const keydown = event({ code: 'KeyK', ctrlKey: true, altKey: true });

    expect(dispatcher.handleCapture(keydown)).toBe(true);
    expect(runCustom).toHaveBeenCalledWith('cs-1');
    expect(keydown.preventDefault).toHaveBeenCalled();
  });

  it('**ההקשה אינה מגיעה למנוע** — לא רק ברירת המחדל נבלעת', () => {
    /*
     * המנוע מחזיק מאזין capture על `.v2-super-editor__stage` שמטפל
     * ב-`Alt+<מקש>` כהקלדת תו, והוא **אינו** בודק `defaultPrevented` לפני
     * שהוא מקליד. `preventDefault` לבדו השאיר את התו במסמך — נמדד
     * ב-scripts/qa/custom-shortcut-owner-probe.mjs. `stopPropagation` הוא מה
     * שמונע מההקשה להגיע אליו בכלל.
     */
    const { dispatcher } = withCustom({}, true, [
      { ...CUSTOM, id: 'cs-alt', label: 'Alt+X', code: 'KeyX', ctrl: false, alt: true },
    ]);
    const keydown = event({ code: 'KeyX', altKey: true });

    expect(dispatcher.handleCapture(keydown)).toBe(true);
    expect(keydown.stopPropagation).toHaveBeenCalled();
  });

  it('הרשימה האישית אינה נבחנת ב-bubble', () => {
    // המסלול היחיד שלה הוא capture. בדיקה שקוראת ל-`handle` ומצפה להרצה
    // הייתה מאשרת מסלול שאינו קיים יותר.
    const { dispatcher, runCustom } = withCustom();

    expect(dispatcher.handle(event({ code: 'KeyK', ctrlKey: true, altKey: true }))).toBe(false);
    expect(runCustom).not.toHaveBeenCalled();
  });

  it('**קיצור מובנה זוכה** — רשומה אישית על אותו צירוף אינה נבחנת', () => {
    // זו ההבטחה של כל המערכת: היא נוספת ואינה דורסת. `Ctrl+B` מריץ את
    // פקודת המנוע, ולא את מה שהמשתמש הצמיד לו בטעות — וגם כאן, שהמסלול
    // האישי הוא זה שרץ ראשון.
    const { dispatcher, runCommand, runCustom } = withCustom();
    const keydown = event({ code: 'KeyB', ctrlKey: true });

    expect(dispatcher.handleCapture(keydown)).toBe(false);
    expect(runCustom).not.toHaveBeenCalled();
    expect(keydown.stopPropagation).not.toHaveBeenCalled();

    expect(dispatcher.handle(keydown)).toBe(true);
    expect(runCommand).toHaveBeenCalledWith('bold', undefined);
  });

  it('צירוף שהדפדפן מטפל בו אינו נגזל בידי קיצור אישי', () => {
    // `Ctrl+V` מותאם לרשומה מובנית שמסומנת `native` ולכן היא **אינה רצה**.
    // די בכך שהיא הותאמה: אילו המסלול האישי היה נבחן אחרי סירוב, קיצור אישי
    // על Ctrl+V היה בולע את ההדבקה עצמה — הפעולה היחידה שאסור לגעת בה.
    const { dispatcher, runCustom } = withCustom({}, true, [
      { ...CUSTOM, id: 'cs-3', code: 'KeyV', ctrl: true, alt: false },
    ]);
    const keydown = event({ code: 'KeyV', ctrlKey: true });

    expect(dispatcher.handleCapture(keydown)).toBe(false);
    expect(runCustom).not.toHaveBeenCalled();
    // ולא נגזלת גם מהדפדפן: `stopPropagation` היה מונע את ההדבקה עצמה.
    expect(keydown.stopPropagation).not.toHaveBeenCalled();
  });

  it('דיאלוג מודאלי פתוח חוסם — לקיצור אישי אין „מותר במודאל”', () => {
    const { dispatcher, runCustom } = withCustom({ isModalOpen: () => true });

    expect(dispatcher.handleCapture(event({ code: 'KeyK', ctrlKey: true, altKey: true }))).toBe(false);
    expect(runCustom).not.toHaveBeenCalled();
  });

  it('שדה טקסט של הממשק חוסם, ואזור המסמך אינו חוסם', () => {
    const field = element('input');
    const surface = element('textarea');

    const blocked = withCustom();
    expect(
      blocked.dispatcher.handleCapture(
        event({ code: 'KeyK', ctrlKey: true, altKey: true, target: field }),
      ),
    ).toBe(false);
    expect(blocked.runCustom).not.toHaveBeenCalled();

    // משטח ההקלדה של המנוע הוא `<textarea>` בתוך אזור המסמך — ושם הקיצור
    // חייב לעבוד, אחרת הוא מת בדיוק כשמקלידים.
    const allowed = withCustom({ isDocumentSurface: (node) => node === surface });
    expect(
      allowed.dispatcher.handleCapture(
        event({ code: 'KeyK', ctrlKey: true, altKey: true, target: surface }),
      ),
    ).toBe(true);
  });

  it('„לא טופל” אינו נבלע — בלי מסמך הצירוף ממשיך הלאה', () => {
    const { dispatcher, runCustom } = withCustom({}, false);
    const keydown = event({ code: 'KeyK', ctrlKey: true, altKey: true });

    expect(dispatcher.handleCapture(keydown)).toBe(false);
    expect(runCustom).toHaveBeenCalledWith('cs-1');
    expect(keydown.preventDefault).not.toHaveBeenCalled();
    expect(keydown.stopPropagation).not.toHaveBeenCalled();
  });

  it('בלי רשימה אישית המנתב מתנהג כמו קודם', () => {
    const { dispatcher } = setup();
    expect(dispatcher.handleCapture(event({ code: 'KeyK', ctrlKey: true, altKey: true }))).toBe(false);
    expect(dispatcher.handle(event({ code: 'KeyK', ctrlKey: true, altKey: true }))).toBe(false);
  });

  it('אירוע שכבר טופל אינו מגיע לרשימה האישית', () => {
    const { dispatcher, runCustom } = withCustom();
    expect(
      dispatcher.handleCapture(
        event({ code: 'KeyK', ctrlKey: true, altKey: true, defaultPrevented: true }),
      ),
    ).toBe(false);
    expect(runCustom).not.toHaveBeenCalled();
  });
});
