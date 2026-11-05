/**
 * המנתב של קיצורי המקלדת: אירוע אחד נכנס, רשומה אחת מהרג'יסטרי רצה.
 *
 * הוא מחזיק את שלוש ההכרעות שקודם היו פזורות ב-`App.vue`, ולכן לא היו נבדקות:
 * מה קורה כשהפוקוס בשדה טקסט, מה קורה כשדיאלוג פתוח, ומתי מותר לבלוע את
 * ההתנהגות של הדפדפן.
 *
 * ## שני שלבים, ולא אחד
 *
 * ברירת המחדל היא **bubble**, אחרי המנוע, וזה מכוון: צירוף שהמנוע קושר בעצמו
 * (Ctrl+B) חייב להיות שלו, ו-`event.defaultPrevented` הוא מה שמונע הרצה
 * כפולה.
 *
 * היוצא מן הכלל הוא `Alt` בלי `Ctrl`, והוא רץ ב-**capture**, לפני המנוע.
 * המדידה: המנוע מחזיק מאזין `keydown` בשלב ה-capture על
 * `.v2-super-editor__stage`, והוא מטפל בצירוף כזה כ**הקלדת תו** — מכניס את
 * התו לבחירה וקורא `preventDefault`. כלומר המנתב שלנו, שיושב ב-bubble, נטש
 * על `defaultPrevented` בלי שהמשתמש ידע למה, ובמקום הפעולה נכתב תו במסמך.
 * מדידה: `scripts/qa/custom-shortcut-owner-probe.mjs`.
 *
 * זה פגע בשני מקומות: בקיצור אישי על `Alt+X` (שהקליד „ס” בפריסה עברית),
 * ו**גם** ב-10 רשומות מובנות — `Alt+1`…`Alt+9` למעבר בין המסמכים הפתוחים
 * ו-`Alt+Q` ל„ספר לי”. כולן לא עבדו כשהסמן במסמך.
 * ההכרעה עצמה יושבת ב-`runsBeforeEngine`, ולא בדגל על כל רשומה: דגל שצריך
 * לזכור להצמיד לרשומה חדשה הוא דגל ששוכחים.
 *
 * ולכן גם `stopPropagation` ולא רק `preventDefault`: המאזין ההוא אינו בודק
 * `defaultPrevented` לפני שהוא מקליד, ובליעה בלעדיו הייתה משאירה את התו
 * במסמך. אירוע שקיצור שלנו טיפל בו אינו אמור להגיע למנוע בכלל.
 *
 * מה **לא** משתנה: „מובנה זוכה”. גם במסלול ה-capture די בכך שהצירוף הותאם
 * לרשומה ברג'יסטרי כדי שהרשימה האישית לא תיבחן — אחרת קיצור אישי על `Ctrl+V`
 * היה בולע את ההדבקה עצמה, וזה בדיוק המסלול שהבדיקה שומרת עליו.
 */
import type { CommandId } from '../../engine/capabilities';
import { matchAny } from './match';
import { SHORTCUTS, type ShellAction, type Shortcut } from './registry';

/** פעולות מעטפת שזמינות גם כאשר גוף המסמך נעול להחלפה. */
const ACTIONS_WITHOUT_DOCUMENT = new Set<ShellAction>([
  'escape', 'focus-next-region', 'focus-prev-region',
  'shortcuts-help',
]);

/**
 * האם המנוע יבלע את הצירוף לפני שהוא יגיע ל-bubble.
 *
 * הכלל נמדד ולא נוחש: המאזין של המנוע מטפל ב-`Alt` **בלי** `Ctrl` כהקלדת
 * תו — נמדד על `Alt+X`, `Alt+Q` ו-`Alt+7`, ובכל השלושה התו נכנס למסמך
 * וה-`preventDefault` שלו הרג את המסלול שלנו. עם `Ctrl` הוא אינו נוגע
 * (נמדד על `Ctrl+Alt+M`), ולכן הכלל הוא בדיוק „Alt בלי Ctrl”.
 *
 * הנפגעות ברג'יסטרי הן 10 רשומות: `Alt+1`…`Alt+9` (מעבר בין המסמכים
 * הפתוחים) ו-`Alt+Q` („ספר לי”). כולן לא עבדו כשהסמן במסמך — כלומר במצב
 * הרגיל — ובמקומן נכתבה הספרה או האות במסמך.
 *
 * `Alt+F8` (ניהול מאקרו) שייך לשלב הזה אבל **לא** היה שבור, וזו מדידה ולא
 * הנחה: בעטיפת כל מאזיני ה-`keydown` בדף ונטרול המאזין שלנו, `Alt+7`
 * ו-`Alt+Q` נבלעו בידי המאזין של המנוע על `.v2-super-editor__stage` והתו
 * נכתב במסמך, ואילו `Alt+F8` עבר את כל המאזינים בלי שאיש קרא `preventDefault`
 * ובלי שנכתב דבר — מקש פונקציה אינו מפיק תו, ואין מה להקליד. הכלל כאן הוא על
 * צורת הצירוף ולא על רשומה-רשומה, ולכן הוא לוקח גם אותו; הוא נמדד בשער
 * (`custom-shortcut-focus-qa.mjs`) כדי שהמעבר שלו ל-capture לא יישאר בלי עדות.
 *
 * מיוצאת לבדיקות: זו ההכרעה שקובעת באיזה שלב רשומה רצה, והיא צריכה להיבדק
 * כקלט-פלט ולא דרך מאזין.
 */
export function runsBeforeEngine(shortcut: Shortcut): boolean {
  if (shortcut.alt === true && shortcut.ctrl !== true) return true;
  return conflictsWithEngineCtrlShift(shortcut);
}

/**
 * צירופי `Ctrl+Shift` שהמנוע קושר בעצמו, ומה הם עושים אצלו.
 *
 * נמדד בדפדפן עם זהות Windows (scripts/qa/ctrl-shift-survey-probe.mjs, על כל האותיות, הספרות
 * ו-`> < + _`): אלה בדיוק הצירופים שהמנוע קורא להם `preventDefault`. רובם
 * מוסכמות TipTap ולא Word — `Ctrl+Shift+E` מרכז, `L`/`R` מיישרים, `7`/`8`
 * פותחים רשימה. כש-`defaultPrevented` כבר דלוק ב-bubble המנתב נוטש, ולכן
 * רשומה שלנו על צירוף כזה לא רצה מעולם: `Ctrl+Shift+E` מירכז את הפסקה במקום
 * להפעיל מעקב שינויים, ו-`Ctrl+Shift+8` פתח רשימת תבליטים במקום סימני עיצוב.
 */
const ENGINE_CTRL_SHIFT: Readonly<Record<string, string>> = {
  KeyE: 'text-align',
  KeyL: 'text-align',
  KeyR: 'text-align',
  KeyH: 'highlight',
  KeyX: 'strikethrough',
  KeyZ: 'redo',
  Digit7: 'ordered-list',
  Digit8: 'bullet-list',
};

/**
 * רשומה שלנו על צירוף שהמנוע קושר **לפעולה אחרת** — ולכן חייבת לרוץ לפניו.
 *
 * כשהפעולה זהה (`Ctrl+Shift+X` קו חוצה, `Ctrl+Shift+Z` ביצוע חוזר) הצירוף
 * נשאר של המנוע, בדיוק כמו `Ctrl+B`: אין מה להרוויח מלעקוף אותו.
 */
function conflictsWithEngineCtrlShift(shortcut: Shortcut): boolean {
  if (shortcut.ctrl !== true || shortcut.shift !== true || shortcut.alt === true) return false;
  const codes = typeof shortcut.code === 'string' ? [shortcut.code] : (shortcut.code ?? []);
  return codes.some((code) => {
    const engine = ENGINE_CTRL_SHIFT[code];
    return engine !== undefined && engine !== shortcut.command;
  });
}

/**
 * האם היעד הוא שדה טקסט — לפי `tagName` ולפי `role`, ולא לפי מאפיין העריכה
 * (שאילתה כזאת על ה-DOM הפנימי של SuperDoc אסורה, ראו
 * tests/unit/engine-boundaries.test.ts).
 *
 * **זה אינו מספיק לבדו כדי לחסום קיצור**, וזה מה שנמדד בדפדפן אמיתי: משטח
 * ההקלדה של המנוע הוא `<textarea aria-label="Text composition input">` ברוחב
 * פיקסל אחד, בתוך אזור המסמך. כלומר ברגע שהמשתמש מתחיל להקליד, `event.target`
 * של **כל** הקשה הוא TEXTAREA — ובדיקה על ה-tag לבדה חסמה את כל הקיצורים בדיוק
 * במצב היחיד שבו הם נחוצים. לכן החסימה מצליבה עם `isDocumentSurface`.
 */
export function isTextEntryTarget(target: EventTarget | null): boolean {
  const element = target as HTMLElement | null;
  const tag = element?.tagName;
  if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return true;
  return element?.getAttribute?.('role') === 'textbox';
}

export interface ShortcutDispatcherDeps {
  /** מריצה פקודת מנוע. אותו מסלול בדיוק של לחיצת כפתור ברצועה. */
  runCommand: (id: CommandId, payload?: unknown) => void;
  /**
   * מריצה פעולת מעטפת ומחזירה האם טופלה. ה-payload הוא זה של הרשומה, והוא
   * `undefined` לרוב הפעולות — רק `tab-goto` נבדלת בו בין רשומה לרשומה.
   */
  runAction: (action: ShellAction, payload?: unknown) => boolean;
  /** האם דיאלוג מודאלי פתוח כרגע. */
  isModalOpen?: () => boolean;
  /**
   * העורך הפעיל עומד להתחלף ואינו מקבל פעולות. חוסם גם קיצורים מה-body;
   * שדות UI, ניווט מיקוד ופעולות שכבה כגון Escape נשארים זמינים.
   */
  isDocumentBlocked?: () => boolean;
  /**
   * האם היעד נמצא **בתוך אזור המסמך** — כלומר שייך למנוע ולא לממשק שלנו.
   *
   * ברירת המחדל היא „לא”, וכל מי שמרכיב את המנתב על מסמך אמיתי חייב למסור
   * אותה: בלעדיה הקיצורים מתים ברגע שמקלידים (ראו `isTextEntryTarget`).
   */
  isDocumentSurface?: (target: EventTarget | null) => boolean;
  /** הרשומות. ברירת המחדל היא הרג'יסטרי; הבדיקות מזריקות רשימה משלהן. */
  shortcuts?: readonly Shortcut[];
  /**
   * הקיצורים שהמשתמש הגדיר בעצמו, בצורת `Shortcut` (`customMatchers`).
   *
   * פונקציה ולא רשימה: הרשימה משתנה בזמן ריצה — הדיאלוג מוסיף ומוחק — ומאזין
   * שנרשם על העתק היה קופא על מה שהיה בפתיחת המסמך.
   */
  customShortcuts?: () => readonly Shortcut[];
  /**
   * מריצה קיצור אישי לפי מזהה. מחזירה האם טופל, כמו פעולת מעטפת: בלי מסמך
   * פתוח אין מה לעצב, ובליעת הצירוף הייתה לוקחת אותו מהדפדפן בלי לתת דבר.
   */
  runCustom?: (id: string) => boolean;
  /** היעד שאליו נרשם המאזין. ברירת מחדל `window`. */
  target?: Pick<EventTarget, 'addEventListener' | 'removeEventListener'>;
}

export interface ShortcutDispatcher {
  /**
   * הרשומות המובנות, בשלב ה-bubble. `true` פירושו „הקיצור רץ, וההתנהגות של
   * הדפדפן נבלעה”.
   */
  handle: (event: KeyboardEvent) => boolean;
  /**
   * שלב ה-capture: הרשומות שהמנוע היה בולע (`runsBeforeEngine`) והקיצורים
   * האישיים. ראו „שני שלבים” בראש הקובץ — זה המסלול היחיד שלהם, והם אינם
   * נבחנים ב-`handle`.
   */
  handleCapture: (event: KeyboardEvent) => boolean;
  /** מנתק את המאזינים. אידמפוטנטי. */
  dispose: () => void;
}

export function createShortcutDispatcher(deps: ShortcutDispatcherDeps): ShortcutDispatcher {
  const shortcuts = deps.shortcuts ?? SHORTCUTS;
  const isModalOpen = deps.isModalOpen ?? (() => false);
  const isDocumentBlocked = deps.isDocumentBlocked ?? (() => false);
  const isDocumentSurface = deps.isDocumentSurface ?? (() => false);

  /**
   * שדה טקסט **של הממשק שלנו** — שם המסמך, שדות החיפוש, בוררים. שם `Ctrl+F`
   * שייך לשדה ואסור לנו לדרוס אותו. משטח ההקלדה של המנוע הוא גם הוא
   * `<textarea>`, אבל הוא יושב בתוך אזור המסמך — ובו הקיצורים חייבים לעבוד.
   */
  const inUiTextEntry = (target: EventTarget | null): boolean =>
    isTextEntryTarget(target) && !isDocumentSurface(target);

  /**
   * הקיצורים האישיים — **רק** אחרי שלא נמצאה רשומה מובנית.
   *
   * הסדר הזה הוא כל ההבטחה של המערכת האישית: היא נוספת ואינה דורסת. די בכך
   * שצירוף **הותאם** לרשומה מובנית כדי שהמסלול האישי לא ייבחן — גם אם הרשומה
   * ההיא סירבה לרוץ (`native`, `onKeyUp`, דיאלוג פתוח). אחרת קיצור אישי על
   * `Ctrl+V` היה נורה בדיוק במצב שבו ההדבקה של הדפדפן חייבת לעבור.
   *
   * שאר ההכרעות זהות לאלה של הרשומות המובנות ואינן משוכפלות: דיאלוג פתוח חוסם
   * (לקיצור אישי אין „מותר במודאל”), ושדה טקסט של הממשק חוסם — הצמדת עיצוב
   * אינה שייכת לשדה שם המסמך.
   *
   * רץ ב-capture, לפני המנוע. ההנמקה והמדידה בראש הקובץ.
   */
  function handleCustom(event: KeyboardEvent): boolean {
    const runCustom = deps.runCustom;
    if (!runCustom || !deps.customShortcuts) return false;

    const match = matchAny(event, deps.customShortcuts());
    if (!match) return false;

    if (isModalOpen()) return false;
    if (inUiTextEntry(event.target)) return false;
    if (isDocumentBlocked()) {
      swallow(event);
      return true;
    }

    const handled = runCustom(match.id);
    if (handled) swallow(event);
    return handled;
  }

  /**
   * בליעה מלאה: גם ברירת המחדל של הדפדפן וגם המשך המסע אל המנוע.
   *
   * `stopPropagation` ולא `preventDefault` לבדו, מפני שהמאזין של המנוע אינו
   * בודק `defaultPrevented` לפני שהוא מקליד את התו — ראו את ראש הקובץ. הוא
   * נקרא גם במסלול ה-bubble, שם אין למי להפסיק את המסע, וזה בסדר: אירוע
   * שהקיצור שלנו טיפל בו אינו אמור להמשיך לאיש.
   */
  function swallow(event: KeyboardEvent): void {
    event.preventDefault();
    event.stopPropagation();
  }

  /**
   * ההרצה של רשומה שכבר הותאמה — משותפת לשני השלבים.
   *
   * ההכרעות כאן אינן תלויות בשלב, ושכפולן היה מאפשר לשני המסלולים להיפרד
   * בשקט: רשומה שנחסמה במודאל בשלב אחד ורצה בשני היא בדיוק סוג הפער שאיש
   * אינו מגלה עד שמשתמש מדווח עליו.
   */
  function runMatched(shortcut: Shortcut, event: KeyboardEvent): boolean {
    // כיווניות פסקה מזוהה בשחרור ה-Shift, ב-`direction.ts`. כאן היא הייתה
    // נורית ברגע שהמשתמש לוחץ Shift — כלומר גם באמצע `Ctrl+Shift+X`.
    if (shortcut.onKeyUp) return false;

    // צירוף שהדפדפן מטפל בו מתועד אצלנו כדי שהתווית תהיה אמיתית, אבל אסור
    // לגעת בו: `preventDefault` על Ctrl+V היה מבטל את ההדבקה עצמה.
    if (shortcut.native) return false;

    if (isModalOpen() && !shortcut.inModal) return false;
    if (!shortcut.inTextEntry && inUiTextEntry(event.target)) return false;
    if (isDocumentBlocked() &&
        (!shortcut.action || !ACTIONS_WITHOUT_DOCUMENT.has(shortcut.action))) {
      swallow(event);
      return true;
    }

    // Capture owns the entire held chord. Letting repeats reach the engine
    // would center the paragraph (E) or create bullets (8) after our toggle.
    // Keep all guards above, but never repeat a non-repeatable action.
    if (event.repeat && shortcut.repeatable !== true) {
      swallow(event);
      return true;
    }

    // פקודת מנוע נחשבת מטופלת תמיד: גם סירוב של ה-controller הוא תשובה,
    // והיא מגיעה למשתמש כהודעה בעברית. פעולת מעטפת מדווחת בעצמה — `Escape`
    // שלא היה לו מה לסגור אינו „מטופל”, ואסור לבלוע אותו.
    let handled = false;
    if (shortcut.command) {
      deps.runCommand(shortcut.command, shortcut.payload);
      handled = true;
    } else if (shortcut.action) {
      handled = deps.runAction(shortcut.action, shortcut.payload);
    }

    // הבליעה אחרי ההרצה, ובכוונה: `Ctrl+S` שאינו מריץ שמירה (כי שמירה כבר
    // רצה) עדיין נחשב מטופל, וחייב למנוע מה-WebView לפתוח את דיאלוג „שמירת
    // דף” שלו.
    if (handled) swallow(event);
    return handled;
  }

  /**
   * שלב ה-capture: הרשומות שהמנוע היה בולע, ואחריהן הרשימה האישית.
   *
   * הסדר הוא „מובנה זוכה”, בדיוק כמו ב-bubble: די בכך שהצירוף **הותאם**
   * לרשומה ברג'יסטרי כדי שהרשימה האישית לא תיבחן. רשומה מובנית שאינה שייכת
   * לשלב הזה (`runsBeforeEngine` שלילי) אינה רצה כאן ואינה מפנה את המקום
   * לרשימה האישית — היא תרוץ ב-bubble, וה-`return false` הוא מה ששומר על
   * שתי ההבטחות יחד.
   */
  function handleCapture(event: KeyboardEvent): boolean {
    // מאזין capture מוקדם משלנו שכבר טיפל. אותה הכרעה של `handle`.
    if (event.defaultPrevented) return false;

    const shortcut = matchAny(event, shortcuts, { ignoreRepeat: true });
    if (shortcut) return runsBeforeEngine(shortcut) ? runMatched(shortcut, event) : false;

    return handleCustom(event);
  }

  function handle(event: KeyboardEvent): boolean {
    // מישהו כבר טיפל. המאזין שלנו יושב על `window` בשלב ה-bubble, כלומר
    // **אחרי** ה-keymap של מנוע העריכה שיושב על אזור המסמך; בלי הבדיקה הזאת
    // צירוף שהמנוע קושר בעצמו (Ctrl+B, למשל) היה מופעל פעמיים — הדגשה
    // וביטולה — והמשתמש היה רואה „הקיצור לא עובד” בלי שום שגיאה.
    if (event.defaultPrevented) return false;

    const shortcut = matchAny(event, shortcuts);
    // הרשימה האישית אינה נבחנת כאן: היא רצה ב-capture, ובשלב הזה הצירוף שלה
    // כבר נבלע. ראו „שני שלבים” בראש הקובץ.
    if (!shortcut) return false;

    // רשומה של שלב ה-capture כבר קיבלה את ההזדמנות שלה. הרצה נוספת כאן
    // הייתה מריצה אותה פעמיים במסלול שבו `stopPropagation` אינו עוצר
    // (בדיקה שקוראת לשתי הפונקציות), ובמסלול האמיתי היא ממילא לא מגיעה.
    if (runsBeforeEngine(shortcut)) return false;

    return runMatched(shortcut, event);
  }

  const target = deps.target ?? (typeof window === 'undefined' ? undefined : window);
  const listener = (event: Event) => {
    handle(event as KeyboardEvent);
  };
  const captureListener = (event: Event) => {
    handleCapture(event as KeyboardEvent);
  };
  target?.addEventListener('keydown', captureListener, true);
  target?.addEventListener('keydown', listener);

  let disposed = false;
  return {
    handle,
    handleCapture,
    dispose() {
      if (disposed) return;
      disposed = true;
      target?.removeEventListener('keydown', captureListener, true);
      target?.removeEventListener('keydown', listener);
    },
  };
}
