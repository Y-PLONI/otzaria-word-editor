/**
 * המצב שהסרגל מצייר, והמקום היחיד שמודד את העמוד המצויר.
 *
 * ## שלושה מקורות, וכל אחד ולמה דווקא הוא
 *
 * 1. **גיאומטריית הדף** — `sections.list()` דרך `readPageMargins`
 *    (engine/page-setup.ts): רוחב העמוד, שולי ההתחלה והסוף וכיוון המקטע,
 *    באינצ'ים שהומרו ל-twips. זה ה-API הציבורי והמוקלד, וזה גם אותו מקום
 *    שהגלריה „שוליים” כותבת אליו — כלומר הסרגל והרצועה קוראים וכותבים לאותו
 *    מספר, ולא לשני מקורות שיכולים להיפרד.
 *
 * 2. **הכניסות של הפסקה** — `readParagraphIndents` (engine/paragraph-format.ts).
 *
 * 3. **המלבן של העמוד על המסך** — נמדד מה-DOM. זו החריגה, והיא מוסברת למטה.
 *
 * ## למה המלבן נמדד ולא מחושב
 *
 * הניסיון הראשון היה לחשב אותו: רוחב העמוד מוכפל בזום, ממורכז ברוחב הפנימי
 * של מיכל הגלילה. זה נמדד מול המנוע האמיתי (Chrome headless, ה-dist הארוז)
 * ונמצא **שגוי בכל זום שאינו 100%**:
 *
 *     זום  | מלבן העמוד בפועל | „ממורכז במיכל”
 *     50%  | ‎-625.0 … -228.2  | 176.1 … 572.9
 *     75%  | ‎-195.1 … 400.2   | 76.9 … 672.2
 *     140% | ‎-120.4 … 990.8   | ‎-44.7 … 749.0
 *
 * הסיבה היא איך שהמנוע מיישם זום: הוא נותן ל-wrapper שלו
 * `width: 100/zoom%` ואז `transform: scale(zoom)` עם `transform-origin: top
 * left` — כלומר תיבת הפריסה של ה-wrapper רחבה מהתוכן הנראה, והעמוד ממורכז
 * בתוך **תיבת הפריסה** ולא בתוך המיכל. כל נוסחה שלנו הייתה משכפלת פנימיות
 * של המנוע, ושדרוג שישנה אותן היה מזיז את הסרגל בשקט ביחס לטקסט. סרגל
 * שמוזז בחצי סנטימטר גרוע מסרגל שאינו קיים.
 *
 * לכן: `getBoundingClientRect()` על העמוד המצויר. **קריאה בלבד** — לא כתיבה,
 * לא בנייה ולא הזזה של DOM. העיגון הוא `data-page-index`, התכונה שהמנוע מסמן
 * בה כל עמוד (הוא עצמו מסנן לפיה: `Number.isInteger(Number(el.dataset.pageIndex))`),
 * וה-host מגיע מ-`ui.viewport.getHost()` — API ציבורי ומוקלד, שנמדד כמחזיר
 * את ה-`div` שלנו עצמו. tests/contract/engine-page-hooks.test.ts מאמת
 * שהעיגון עדיין קיים באריזת המנוע, ו-tests/unit/engine-boundaries.test.ts
 * מאמת שאיש מלבד הקובץ הזה אינו נוגע בו.
 *
 * ## מה שהסרגל **אינו** מציג, ולמה
 *
 * שני חלקים מהסרגל של Word אינם כאן, ושניהם מטעמי מדידה ולא מטעמי זמן:
 *
 *   - **כניסת שורה ראשונה וכניסה תלויה.** בפסקה עברית המנוע מצייר אותן
 *     הפוך: `firstLine: 1440` הזיז את השורה הראשונה **החוצה** אל תוך השוליים
 *     (הקצה הימני שלה עבר מ-96px מקצה הדף ל-0), ו-`hanging` הכניס אותה
 *     פנימה — כלומר בדיוק ההפך מהסמנטיקה של Word. `left`/`right` דווקא כן
 *     ממופים נכון לצד ההתחלה והסוף (נמדד). סמן שגורר ערך שמצויר הפוך הוא
 *     סרגל שמשקר, ולכן הוא אינו מוצג עד שהמנוע יתקן.
 *   - **עצירות טאב.** `setTabStop` מחזיר `success: true` וכותב ל-DOCX, אבל
 *     `doc.get()` **אינו מחזיר** את `tabs` בתכונות הפסקה — כלומר אין דרך
 *     לקרוא את העצירות הקיימות, ולכן אין דרך לצייר אותן. סרגל שמראה רק את
 *     העצירות שנוספו בו עצמו, ומעלים את אלה שהגיעו מקובץ Word, מטעה.
 *
 * שניהם מדווחים ב-docs/engine-gaps.md.
 *
 * ## מה שנמדד על הסרגל עצמו, אחרי שנכתב
 *
 * ה-`dist` הארוז ב-Chrome headless, עם לחיצות וגרירות אמיתיות
 * (`Input.dispatchMouseEvent`) ולא סימולציה של אירועים ב-DOM:
 *
 *   - **יישור.** אזור הטקסט בסרגל מול הטקסט שהמנוע צייר: הפרש של 0.0px
 *     ב-100% ו-0.1px ב-70%. הגרסה הראשונה הראתה 1.0px — `border` על מלבן
 *     העמוד הזיז את כל הילדים, מפני שמיקום מוחלט נמדד מתיבת הריפוד. מאז זה
 *     `box-shadow: inset`.
 *   - **זום.** `viewport.observe` **אינו** מדווח על שינוי זום: ב-70% העמוד
 *     הצטמצם ל-555px והסרגל נשאר על 794. מכאן ה-prop `zoom` בקומפוננטות,
 *     ומכאן גם `SETTLE_DELAYS_MS` — המנוע מצייר מחדש אחרי שהאירוע כבר הגיע.
 *   - **גרירה.** גרירת ידית שוליים 60px פנימה הזיזה את קצה הטקסט מ-697.7px
 *     ל-642.5px, והסרגל חזר ליישור מלא. גרירת סמן כניסה הכניסה את הטקסט
 *     פנימה ב-66px, וגרירת השוליים העליונים הורידה את הטקסט ב-64.6px.
 *   - **מה שהמדידה תפסה ובדיקות היחידה לא:** `same()` כאן השווה בגרסה
 *     ראשונה את רוחב הדף ואת שני השוליים האופקיים בלבד, ולכן שינוי בשוליים
 *     העליונים נחשב „ללא שינוי” — הטקסט זז והסרגל האנכי לא.
 */
import type { CommandOutcome } from './command-adapter';
import {
  applyParagraphIndentation,
  type ParagraphIndentReading,
  type ParagraphIndents,
  type ParagraphFormatTarget,
  type ParagraphTarget,
} from './paragraph-format';
import type { PageMarginsState } from './page-setup';
import type { RulerUnit } from './ruler-geometry';

/* ------------------------------------------------------------------ */
/* המלבן של העמוד                                                     */
/* ------------------------------------------------------------------ */

/** התכונה שהמנוע מסמן בה עמוד מצויר. ראו הערת הפתיחה. */
export const PAGE_INDEX_ATTRIBUTE = 'data-page-index';

/** המלבן של העמוד, ביחס לאלמנט הייחוס של הסרגל. */
export interface PageRect {
  leftPx: number;
  widthPx: number;
  topPx: number;
  heightPx: number;
}

/**
 * איזה עמוד נמדד, וגם על איזה ציר משווים שינוי.
 *
 * `'x'` — הסרגל האופקי: העמוד הראשון מספיק (כל העמודים ממורכזים באותו מקום
 * אופקית), ושינוי אנכי אינו מעניין אותו. `'y'` — הסרגל האנכי: העמוד **הנראה**
 * הוא הרלוונטי, בדיוק כמו ב-Word, שם הסרגל האנכי מתאר את העמוד שעל המסך.
 *
 * למה שני הצירים אינם מדווחים יחד: בגלילה אנכית `topPx` משתנה בכל פריים,
 * ודיווח עליו לסרגל האופקי היה מחשב מחדש מאה שנתות בכל פריים — בלי ששנתה
 * אחת זזה.
 */
export type RulerAxis = 'x' | 'y';

/** מה שנצרך מ-`superdoc.ui`: ה-host המצויר וההודעה על שינוי גיאומטריה. */
export interface ViewportSource {
  viewport?: {
    getHost?: () => HTMLElement | null;
    observe?: (listener: () => void) => () => void;
  };
}

/** ה-host שהמנוע מצייר לתוכו, או `null` לפני שהמסמך נטען. */
export function paintedHost(ui: ViewportSource | null | undefined): HTMLElement | null {
  const getHost = ui?.viewport?.getHost;
  if (typeof getHost !== 'function') return null;
  try {
    return getHost.call(ui?.viewport) ?? null;
  } catch {
    return null;
  }
}

/**
 * המלבן של העמוד המצויר, ביחס ל-`reference`.
 *
 * לסרגל האופקי נמדד העמוד **הראשון**, ובכוונה: כל העמודים ממורכזים באותו מקום
 * אופקית — הם ילדים של אותו wrapper — ולכן ה-x שלהם זהה, וזה כל מה שהוא צריך.
 * חיפוש „העמוד הפעיל” היה מוסיף תלות בבחירה בלי להזיז פיקסל.
 *
 * לסרגל האנכי נמדד העמוד **הנראה ביותר**, מפני ששם ההפך נכון: שוליים עליונים
 * של עמוד שגללנו ממנו הם מספר שאינו מתאר דבר על המסך.
 */
export function measurePageRect(
  host: HTMLElement | null,
  reference: HTMLElement | null,
  axis: RulerAxis = 'x',
): PageRect | null {
  if (!host || !reference) return null;
  const pages = host.querySelectorAll(`[${PAGE_INDEX_ATTRIBUTE}]`);
  if (pages.length === 0) return null;
  const referenceBox = reference.getBoundingClientRect();

  const page = axis === 'y' ? mostVisiblePage(pages, referenceBox) : pages[0];
  if (!(page instanceof HTMLElement)) return null;

  const pageBox = page.getBoundingClientRect();
  if (!(pageBox.width > 0)) return null;

  return {
    leftPx: pageBox.left - referenceBox.left,
    widthPx: pageBox.width,
    topPx: pageBox.top - referenceBox.top,
    heightPx: pageBox.height,
  };
}

/**
 * העמוד שהכי הרבה ממנו נראה בתוך אזור הייחוס.
 *
 * זה מה שהסרגל האנכי חייב: המשתמש גולל לעמוד השני, והשוליים שהסרגל מראה הם
 * של העמוד שמולו — לא של העמוד הראשון שנשאר מעליו. בתיקו נבחר הראשון, וכך
 * הסרגל אינו מהבהב בין שני עמודים בגבול המדויק.
 */
function mostVisiblePage(pages: NodeListOf<Element>, reference: DOMRect): Element | null {
  let best: Element | null = null;
  let bestVisible = -1;

  for (const page of Array.from(pages)) {
    const box = page.getBoundingClientRect();
    const visible = Math.min(box.bottom, reference.bottom) - Math.max(box.top, reference.top);
    if (visible > bestVisible) {
      bestVisible = visible;
      best = page;
    }
  }

  return best;
}

/**
 * שני מלבנים שנראים זהים **על הציר שנמדד**. חצי פיקסל אינו שינוי שכדאי לרנדר
 * עליו, ותזוזה בציר השני אינה עניינו של הסרגל הזה בכלל.
 */
function sameRect(a: PageRect | null, b: PageRect | null, axis: RulerAxis): boolean {
  if (a === null || b === null) return a === b;
  if (axis === 'y') {
    return Math.abs(a.topPx - b.topPx) < 0.5 && Math.abs(a.heightPx - b.heightPx) < 0.5;
  }
  return Math.abs(a.leftPx - b.leftPx) < 0.5 && Math.abs(a.widthPx - b.widthPx) < 0.5;
}

export interface PageRectWatchOptions {
  /** מיכל הגלילה שהמנוע מצייר בתוכו. מגיע מ-`paintedHost`. */
  host: HTMLElement | null;
  /** האלמנט שביחס אליו נמדד המלבן — מיכל הסרגל עצמו. */
  reference: HTMLElement | null;
  /** ה-controller, בשביל `viewport.observe`. */
  ui?: ViewportSource | null;
  /** הציר שהסרגל הקורא מצייר עליו. ברירת המחדל: אופקי. */
  axis?: RulerAxis;
  onChange: (rect: PageRect | null) => void;
}

export interface PageRectWatch {
  /** מדידה מיידית ועוד כמה אחריה. ראו SETTLE_DELAYS_MS. */
  measure(): void;
  dispose(): void;
}

/**
 * המדידות שאחרי מדידה שהתבקשה מבחוץ.
 *
 * למה זה נדרש, ונמדד: שינוי זום מגיע אלינו כשהמנוע רק **מתחיל** לצייר מחדש.
 * מדידה יחידה באותו רגע תופסת את הגיאומטריה הישנה, והסרגל נשאר ברוחב של
 * הזום הקודם עד הגלילה הבאה — כך זה נראה בבדיקה החיה על ה-dist הארוז: זום
 * 70% הקטין את העמוד ל-555px, והסרגל נשאר על 794.
 *
 * הסולם מכסה גם ציור מהיר וגם עימוד מחדש של מסמך ארוך, וכל מדידה שאינה
 * משנה דבר נזרקת ב-`sameRect` — כלומר המחיר של המדידות המיותרות הוא
 * `getBoundingClientRect` אחד, בלי רינדור.
 */
export const SETTLE_DELAYS_MS = [80, 250, 600] as const;

/**
 * עוקבת אחרי המלבן של העמוד ומדווחת על כל שינוי.
 *
 * שלושה מקורות עדכון, וכולם נדרשים: גלילה (העמוד זז מתחת לסרגל), שינוי גודל
 * של המיכל (`ResizeObserver`), ו-`viewport.observe` של ה-controller — שזו
 * ההודעה הציבורית על „הגיאומטריה כבר אינה תקפה”, כלומר זום, עימוד מחדש
 * ושינוי שוליים. בלי השלישי הסרגל היה נשאר במקומו הישן אחרי כל שינוי זום עד
 * הגלילה הבאה.
 *
 * המדידה עצמה מושהית ל-`requestAnimationFrame`: אירועי גלילה מגיעים בקצב
 * הפריים ממילא, ומדידה בכל אחד מהם היא layout thrash.
 */
export function watchPageRect(options: PageRectWatchOptions): PageRectWatch {
  const { host, reference, ui, onChange } = options;
  const axis: RulerAxis = options.axis ?? 'x';
  let last: PageRect | null = null;
  let frame: number | null = null;
  /**
   * דגל נפרד מה-handle, ולא `frame !== null`.
   *
   * זה נמדד בבדיקה: כש-`requestAnimationFrame` מריץ את ה-callback **מיד**
   * (כפיל בבדיקות, או polyfill), ההשמה `frame = requestAnimationFrame(...)`
   * מתרחשת *אחרי* שה-callback כבר איפס אותו — ולכן `frame` נשאר על ה-handle
   * לנצח וכל מדידה נוספת נחסמת. סרגל שמפסיק לעקוב אחרי הגלילה השנייה.
   */
  let pending = false;
  let disposed = false;

  function measureNow(): void {
    if (disposed) return;
    const next = measurePageRect(host, reference, axis);
    if (sameRect(next, last, axis)) return;
    last = next;
    onChange(next);
  }

  function schedule(): void {
    if (disposed || pending) return;
    if (typeof requestAnimationFrame !== 'function') {
      measureNow();
      return;
    }
    pending = true;
    frame = requestAnimationFrame(() => {
      pending = false;
      frame = null;
      measureNow();
    });
  }

  host?.addEventListener('scroll', schedule, { passive: true });

  let resize: ResizeObserver | null = null;
  if (typeof ResizeObserver === 'function' && host) {
    resize = new ResizeObserver(schedule);
    resize.observe(host);
    if (reference) resize.observe(reference);
  }

  let unobserve: (() => void) | null = null;
  const observe = ui?.viewport?.observe;
  if (typeof observe === 'function') {
    try {
      unobserve = observe.call(ui?.viewport, schedule) ?? null;
    } catch {
      unobserve = null;
    }
  }

  const timers = new Set<ReturnType<typeof setTimeout>>();

  /** מדידה עכשיו, ועוד כמה אחריה — ראו SETTLE_DELAYS_MS. */
  function measure(): void {
    schedule();
    for (const delay of SETTLE_DELAYS_MS) {
      const timer = setTimeout(() => {
        timers.delete(timer);
        schedule();
      }, delay);
      timers.add(timer);
    }
  }

  measureNow();

  return {
    measure,
    dispose() {
      disposed = true;
      host?.removeEventListener('scroll', schedule);
      resize?.disconnect();
      for (const timer of timers) clearTimeout(timer);
      timers.clear();
      if (frame !== null && typeof cancelAnimationFrame === 'function') cancelAnimationFrame(frame);
      try {
        unobserve?.();
      } catch {
        /* ביטול מנוי שנכשל אינו סיבה להפיל פירוק */
      }
    },
  };
}

/* ------------------------------------------------------------------ */
/* יחידת המידה                                                         */
/* ------------------------------------------------------------------ */

/** מה שנצרך מהמופע: `getMeasurementUnit` בלבד. */
export interface MeasurementUnitSource {
  getMeasurementUnit?: () => unknown;
}

/**
 * יחידת המידה של הסרגל, מהמנוע.
 *
 * `create-editor.ts` פותח ב-`'cm'`, אבל הפקודה `measurement-unit` קיימת
 * ב-registry — כלומר היחידה יכולה להתחלף בזמן ריצה, וסרגל שמקודד „ס\"מ” היה
 * מציג שנתות באינץ' עם תווית של סנטימטר. אותה הכרעה בדיוק כמו ביחידת המידה
 * בפאנל הכותרות (engine/hf-chrome.ts).
 */
export function readRulerUnit(host: MeasurementUnitSource | null | undefined): RulerUnit {
  const read = host?.getMeasurementUnit;
  if (typeof read !== 'function') return 'cm';
  try {
    return read.call(host) === 'in' ? 'in' : 'cm';
  } catch {
    return 'cm';
  }
}

/* ------------------------------------------------------------------ */
/* מצב הסרגל מצד המסמך                                                */
/* ------------------------------------------------------------------ */

/** מה שנקרא מהמסמך. `null` בכל שדה פירושו „אין מה לצייר”, לא כשל. */
export interface RulerReading {
  page: PageMarginsState;
  /** הכניסות של הפסקה שהסמן בה, או `null` כשאין סמן במסמך. */
  indents: ParagraphIndents | null;
  /** היעד לכתיבה חזרה. עולה ויורד יחד עם `indents`. */
  target: ParagraphTarget | null;
}

export interface RulerModelSource {
  readPage: () => Promise<PageMarginsState | null>;
  readIndents: () => Promise<ParagraphIndentReading | null>;
  onChange: (reading: RulerReading | null) => void;
}

export interface RulerModel {
  getState(): RulerReading | null;
  /**
   * הסרגל מוצג או מוסתר. סרגל מוסתר אינו קורא כלום — `doc.get()` סורק את
   * המסמך כולו, וקריאה שלו על כל תזוזת סמן כשאיש אינו רואה את התוצאה היא
   * עבודה מיותרת על מסמך של שמונים עמודים.
   */
  setEnabled(enabled: boolean): void;
  noteSelectionChanged(): void;
  noteDocumentChanged(): void;
  /** קריאה מיידית, בלי השהיה — אחרי פתיחת מסמך או הדלקת הסרגל. */
  refreshNow(): void;
  dispose(): void;
}

/**
 * השהיית הקריאה אחרי תזוזת סמן. זהה ל-CURRENT_PAGE_DEBOUNCE_MS שבשורת המצב,
 * ומאותו טעם: סמני כניסה שמתעדכנים חצי שנייה אחרי הסמן נראים תקועים.
 */
export const RULER_SELECTION_DEBOUNCE_MS = 150;

/**
 * השהיית הקריאה אחרי שינוי במסמך. ארוכה יותר — הקלדה אינה משנה כניסות,
 * והקריאה סורקת את המסמך כולו.
 */
export const RULER_DOCUMENT_DEBOUNCE_MS = 500;

/**
 * מרכיבה את מצב הסרגל וקוראת אותו בהשקטה.
 *
 * אותה תבנית כמו createDocMetrics (engine/doc-metrics.ts), ומאותן סיבות:
 * מונה דורות שזורק תשובה של מסמך שכבר נסגר, דיווח רק על שינוי אמיתי, ו-
 * `dispose` שמבטל את מה שבאוויר.
 */
export function createRulerModel(source: RulerModelSource): RulerModel {
  let reading: RulerReading | null = null;
  let enabled = false;
  let disposed = false;
  let generation = 0;
  let selectionTimer: ReturnType<typeof setTimeout> | undefined;
  let documentTimer: ReturnType<typeof setTimeout> | undefined;

  /**
   * כל שדה של הדף נבדק, ולא רק אלה שהסרגל האופקי צריך.
   *
   * זה נמדד: גרסה קודמת השוותה `pageWidthTwips`/`left`/`right`/`direction`
   * בלבד, וגרירה של השוליים העליונים נחשבה „ללא שינוי” — הטקסט זז במסמך
   * והסרגל האנכי נשאר במקומו. השוואה חלקית היא באג שקט מסוג שקשה לאתר,
   * מפני שהיא נראית כמו אופטימיזציה.
   */
  function same(a: RulerReading | null, b: RulerReading | null): boolean {
    if (a === null || b === null) return a === b;
    return (
      a.page.pageWidthTwips === b.page.pageWidthTwips &&
      a.page.pageHeightTwips === b.page.pageHeightTwips &&
      a.page.leftTwips === b.page.leftTwips &&
      a.page.rightTwips === b.page.rightTwips &&
      a.page.topTwips === b.page.topTwips &&
      a.page.bottomTwips === b.page.bottomTwips &&
      // גם הערכים האפקטיביים: הוספת כותרת עליונה מרימה את שולי הטקסט בלי
      // שאיש נגע ב-`w:top`, והשוואה חלקית הייתה משאירה את הסרגל על הישן.
      a.page.effectiveTopTwips === b.page.effectiveTopTwips &&
      a.page.effectiveBottomTwips === b.page.effectiveBottomTwips &&
      a.page.direction === b.page.direction &&
      a.indents?.leftTwips === b.indents?.leftTwips &&
      a.indents?.rightTwips === b.indents?.rightTwips &&
      a.indents?.firstLineTwips === b.indents?.firstLineTwips &&
      a.indents?.hangingTwips === b.indents?.hangingTwips &&
      a.indents?.bidi === b.indents?.bidi &&
      a.target?.nodeId === b.target?.nodeId
    );
  }

  function publish(next: RulerReading | null): void {
    if (disposed || same(next, reading)) return;
    reading = next;
    source.onChange(next);
  }

  async function read(): Promise<void> {
    const mine = ++generation;
    if (!enabled) {
      publish(null);
      return;
    }

    let page: PageMarginsState | null = null;
    try {
      page = await source.readPage();
    } catch {
      page = null;
    }
    if (disposed || mine !== generation) return;
    if (!page) {
      publish(null);
      return;
    }

    let paragraph: ParagraphIndentReading | null = null;
    try {
      paragraph = await source.readIndents();
    } catch {
      paragraph = null;
    }
    if (disposed || mine !== generation) return;

    publish({
      page,
      indents: paragraph?.indents ?? null,
      target: paragraph?.target ?? null,
    });
  }

  function schedule(which: 'selection' | 'document'): void {
    if (disposed || !enabled) return;
    if (which === 'selection') {
      clearTimeout(selectionTimer);
      selectionTimer = setTimeout(() => void read(), RULER_SELECTION_DEBOUNCE_MS);
      return;
    }
    clearTimeout(documentTimer);
    documentTimer = setTimeout(() => void read(), RULER_DOCUMENT_DEBOUNCE_MS);
  }

  return {
    getState: () => reading,

    setEnabled(next) {
      if (enabled === next) return;
      enabled = next;
      if (!enabled) {
        // הדור עולה כדי שקריאה שבאוויר לא תדווח אחרי הכיבוי.
        generation += 1;
        clearTimeout(selectionTimer);
        clearTimeout(documentTimer);
        publish(null);
        return;
      }
      void read();
    },

    noteSelectionChanged: () => schedule('selection'),
    noteDocumentChanged: () => schedule('document'),
    refreshNow: () => void read(),

    dispose() {
      disposed = true;
      generation += 1;
      clearTimeout(selectionTimer);
      clearTimeout(documentTimer);
    },
  };
}

/* ------------------------------------------------------------------ */
/* כתיבה חזרה                                                          */
/* ------------------------------------------------------------------ */

/**
 * כניסות מגרירה בסרגל.
 *
 * `startTwips`/`endTwips` הם הצד הלוגי, והם נכתבים ל-`left`/`right` **בלי
 * היפוך**: נמדד שבפסקה עברית `left: 1440` הזיז את הקצה הימני של הטקסט פנימה
 * ב-96px, ו-`right: 1440` הזיז את הקצה השמאלי. כלומר `left` הוא תמיד צד
 * ההתחלה ו-`right` תמיד צד הסוף — הסמנטיקה של `w:start`/`w:end` ב-OOXML.
 * ההיפוך היחיד הוא בציור, ב-`pixelOffset` (engine/ruler-geometry.ts).
 *
 * „מיוחד” נשמר כפי שהוא: `setIndentation` מחליף את `<w:ind>` כולו (נמדד),
 * ולכן גרירה של כניסת פסקה הייתה **מוחקת** כניסת שורה ראשונה שהגיעה מקובץ
 * Word. הערכים הקיימים נקראים מהמסמך ונשלחים בחזרה יחד עם החדשים.
 */
export function applyRulerIndents(
  host: ParagraphFormatTarget,
  target: ParagraphTarget,
  current: ParagraphIndents,
  next: { startTwips: number; endTwips: number },
): Promise<CommandOutcome> {
  const special: 'none' | 'firstLine' | 'hanging' =
    current.hangingTwips > 0 ? 'hanging' : current.firstLineTwips > 0 ? 'firstLine' : 'none';
  const amountTwips = special === 'hanging' ? current.hangingTwips : current.firstLineTwips;

  return applyParagraphIndentation(host, target, {
    leftTwips: next.startTwips,
    rightTwips: next.endTwips,
    special,
    amountTwips: special === 'none' ? 0 : amountTwips,
  });
}
