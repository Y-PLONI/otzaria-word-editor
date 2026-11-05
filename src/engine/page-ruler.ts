/**
 * המצב שהסרגל מצייר, והמקום היחיד שמודד את העמוד המצויר.
 *
 * **קורא שני, מאז „גבולות עמוד”:** `measureAllPageRects`/`watchAllPageRects`
 * בתחתית הקובץ נועדו לשכבת engine/page-border-layer.ts — היא צריכה את המלבן
 * של **כל** עמוד, לא רק הראשון או הנראה ביותר. העיגון זהה (`data-page-index`)
 * ולכן הוא יושב כאן ולא שם, לפי אותו כלל שהערת הפתיחה של הסרגל האנכי מנסחת:
 * מקום אחד בלבד נוגע בעיגון. tests/unit/engine-boundaries.test.ts אוכף זאת.
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
  sameParagraphTarget,
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
/* המלבנים של **כל** העמודים                                          */
/* ------------------------------------------------------------------ */

/**
 * מלבן עמוד אחד מתוך רשימת כל העמודים, עם האינדקס שהמנוע עצמו סימן אותו בו
 * (`data-page-index`) — לא סתם סדר ה-DOM, כדי ש-`v-for` ב-Vue יקבל מפתח יציב
 * גם אם עמודים מתווספים/יורדים מהאמצע בעימוד מחדש.
 */
export interface IndexedPageRect extends PageRect {
  pageIndex: number;
}

/**
 * המלבנים של כל העמודים המצוירים, ביחס ל-`reference` — לא רק העמוד הראשון
 * (כמו `measurePageRect`) אלא כולם. זה מה ששכבת „גבולות עמוד” צריכה: מסגרת
 * סביב **כל** עמוד במסמך רב-עמודי, לא רק סביב זה שהסרגל עוקב אחריו.
 *
 * אותו עיגון בדיוק כמו `measurePageRect` (`data-page-index`), ואותה קריאה
 * טהורה — `getBoundingClientRect` בלבד, בלי לגעת ב-DOM. ראו
 * tests/unit/engine-boundaries.test.ts: רק הקובץ הזה רשאי לגעת בעיגון.
 */
export function measureAllPageRects(
  host: HTMLElement | null,
  reference: HTMLElement | null,
): readonly IndexedPageRect[] {
  if (!host || !reference) return [];
  const pages = host.querySelectorAll(`[${PAGE_INDEX_ATTRIBUTE}]`);
  if (pages.length === 0) return [];
  const referenceBox = reference.getBoundingClientRect();

  const out: IndexedPageRect[] = [];
  pages.forEach((node) => {
    if (!(node instanceof HTMLElement)) return;
    const pageIndex = Number(node.getAttribute(PAGE_INDEX_ATTRIBUTE));
    if (!Number.isInteger(pageIndex)) return;
    const box = node.getBoundingClientRect();
    if (!(box.width > 0) || !(box.height > 0)) return;
    out.push({
      pageIndex,
      leftPx: box.left - referenceBox.left,
      widthPx: box.width,
      topPx: box.top - referenceBox.top,
      heightPx: box.height,
    });
  });
  return out;
}

/** כמו `sameRect`, על מערך שלם: אורך שונה או עמוד אחד ששינה דיו הם כבר שינוי. */
function sameRects(a: readonly IndexedPageRect[], b: readonly IndexedPageRect[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const x = a[i];
    const y = b[i];
    if (
      x.pageIndex !== y.pageIndex ||
      Math.abs(x.leftPx - y.leftPx) >= 0.5 ||
      Math.abs(x.topPx - y.topPx) >= 0.5 ||
      Math.abs(x.widthPx - y.widthPx) >= 0.5 ||
      Math.abs(x.heightPx - y.heightPx) >= 0.5
    ) {
      return false;
    }
  }
  return true;
}

export interface AllPageRectsWatchOptions {
  /** מיכל הגלילה שהמנוע מצייר בתוכו. מגיע מ-`paintedHost`. */
  host: HTMLElement | null;
  /** האלמנט שביחס אליו נמדדים המלבנים — שכבת הציור שלנו עצמה. */
  reference: HTMLElement | null;
  /** ה-controller, בשביל `viewport.observe`. */
  ui?: ViewportSource | null;
  onChange: (rects: readonly IndexedPageRect[]) => void;
}

/**
 * כמו `watchPageRect`, על כל העמודים יחד — אותה תשתית מעקב בדיוק (גלילה,
 * שינוי גודל, `viewport.observe`, מדידות התיישבות ב-`SETTLE_DELAYS_MS`), רק
 * שהמדידה וההשוואה הן על מערך ולא על עמוד יחיד. ראו `watchPageRect` להסבר
 * המלא על כל אחד מהמקורות.
 */
export function watchAllPageRects(options: AllPageRectsWatchOptions): PageRectWatch {
  const { host, reference, ui, onChange } = options;
  let last: readonly IndexedPageRect[] = [];
  let frame: number | null = null;
  let pending = false;
  let disposed = false;

  function measureNow(): void {
    if (disposed) return;
    const next = measureAllPageRects(host, reference);
    if (sameRects(next, last)) return;
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
/* מלבני טקסט גולמיים בכל עמוד — למספרי שורות                          */
/* ------------------------------------------------------------------ */

/**
 * מלבן טקסט גולמי אחד, כפי ש-`Range.getClientRects()` מחזיר אותו — לפני כל
 * קיבוץ לשורות או סינון קונטיינרים. engine/line-number-layer.ts הוא זה
 * שהופך רשימה כזאת לשורות אמיתיות; כאן רק המדידה.
 *
 * ## קורא שלישי, מאז „מספרי שורות” — ולמה **לא** דרך selector אל המנוע
 *
 * מספרי שורות דורשים לדעת את המיקום האנכי של כל שורת טקסט מוצגת בעמוד, לא
 * רק את מלבן העמוד השלם (`measureAllPageRects` למעלה). למנוע אין API
 * ציבורי שמחזיר את זה, ואין ב-DOM המצויר class name או תכונה שהחוזה שלנו
 * (tests/unit/engine-boundaries.test.ts) מרשה לנו לחפש: הבדיקה שם אוסרת
 * `querySelector`/`querySelectorAll`/`closest` עם מחרוזת שמכילה `superdoc`
 * או `.sd-` בכל קובץ מלבד שני העיגונים המתועדים (כותרות, ומלבן העמוד) —
 * ומבנה ה-DOM הפנימי של המנוע (class names כמו „קו”/„קטע” לכל שורה) הוא
 * בדיוק סוג הדבר שהחוזה הזה נועד להגן מפניו: המנוע עצמו מתעד את אותם class
 * names כ„חוזה פנימי בין הצייר לקורא… שינוי כאן הוא breaking change לשניהם”
 * — כלומר לא הבטחה ציבורית, ולא משהו שמותר לתוסף חיצוני להישען עליו.
 *
 * **הטכניקה שכן ציבורית:** `Range.selectNodeContents(pageEl)` ואז
 * `range.getClientRects()` — API תקני של ה-DOM, לא selector אל מבנה פנימי.
 * הוא מחזיר מלבן לכל תיבת טקסט שהטווח חוצה, כולל:
 *   - כמה מלבנים על אותה שורה חזותית ממש (ריצות טקסט נפרדות — כיווניות,
 *     או פשוט כמה `span` צמודים).
 *   - מלבן-קונטיינר גדול לכל אלמנט בלוקי שכולו בתוך הטווח (פסקה שלמה,
 *     כותרת) — זה בנוסף למלבנים של השורות שבתוכו, לא במקומם.
 * שני אלה נמדדו (לא הונחו): על עמוד עם 51 שורות חזותיות אמיתיות, הטווח החזיר
 * 107 מלבנים; רוב הפער (56) הם ריצות-טקסט נוספות על שורה קיימת וכמה מלבני
 * קונטיינר. הסינון בין שורה לקונטיינר (`groupLinesFromRects`,
 * engine/line-number-layer.ts) משתמש בגובה חריג ביחס לחציון — קונטיינר של
 * פסקה שלמה גבוה משמעותית משורה בודדת — וקיבוץ ריצות על אותה שורה לפי `top`
 * זהה. אחרי שני אלה: 51 קבוצות בדיוק, תואם אחד-לאחד (מיקום כולל) לשורות
 * האמיתיות. איפה עובר הגבול בין כותרת/שוליים לגוף הטקסט — לא כאן: זו שאלה
 * גיאומטרית (איזה פס גובה בעמוד), לא שאלת DOM, ונענית ב-line-number-layer.ts
 * מתוך `readPageMargins` (engine/page-setup.ts) — לא מ-selector נוסף.
 */
export interface RawTextRect {
  leftPx: number;
  topPx: number;
  widthPx: number;
  heightPx: number;
}

/** מלבני הטקסט הגולמיים של עמוד אחד, ביחס ל-`reference`. */
export function measurePageContentRects(
  pageEl: HTMLElement,
  reference: HTMLElement,
): readonly RawTextRect[] {
  if (typeof document === 'undefined' || typeof document.createRange !== 'function') return [];

  let range: Range;
  try {
    range = document.createRange();
    range.selectNodeContents(pageEl);
  } catch {
    return [];
  }

  const referenceBox = reference.getBoundingClientRect();
  const out: RawTextRect[] = [];
  const list = range.getClientRects();
  for (let i = 0; i < list.length; i++) {
    const r = list[i]!;
    if (!(r.width > 0) || !(r.height > 0)) continue;
    out.push({
      leftPx: r.left - referenceBox.left,
      topPx: r.top - referenceBox.top,
      widthPx: r.width,
      heightPx: r.height,
    });
  }
  return out;
}

/** מלבני הטקסט הגולמיים של עמוד אחד, עם ה-`data-page-index` שלו. */
export interface PageContentRects {
  pageIndex: number;
  rects: readonly RawTextRect[];
}

/** אותו עיגון בדיוק כמו `measureAllPageRects` — `data-page-index` תחת ה-host. */
export function measureAllPageContentRects(
  host: HTMLElement | null,
  reference: HTMLElement | null,
): readonly PageContentRects[] {
  if (!host || !reference) return [];
  const pages = host.querySelectorAll(`[${PAGE_INDEX_ATTRIBUTE}]`);
  if (pages.length === 0) return [];

  const out: PageContentRects[] = [];
  pages.forEach((node) => {
    if (!(node instanceof HTMLElement)) return;
    const pageIndex = Number(node.getAttribute(PAGE_INDEX_ATTRIBUTE));
    if (!Number.isInteger(pageIndex)) return;
    out.push({ pageIndex, rects: measurePageContentRects(node, reference) });
  });
  return out;
}

/** כמו `sameRects`, אבל על מבנה מקונן (עמוד → רשימת מלבנים). */
function sameContentRects(a: readonly PageContentRects[], b: readonly PageContentRects[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const pa = a[i]!;
    const pb = b[i]!;
    if (pa.pageIndex !== pb.pageIndex || pa.rects.length !== pb.rects.length) return false;
    for (let j = 0; j < pa.rects.length; j++) {
      const x = pa.rects[j]!;
      const y = pb.rects[j]!;
      if (
        Math.abs(x.leftPx - y.leftPx) >= 0.5 ||
        Math.abs(x.topPx - y.topPx) >= 0.5 ||
        Math.abs(x.widthPx - y.widthPx) >= 0.5 ||
        Math.abs(x.heightPx - y.heightPx) >= 0.5
      ) {
        return false;
      }
    }
  }
  return true;
}

export interface AllPageContentRectsWatchOptions {
  /** מיכל הגלילה שהמנוע מצייר בתוכו. מגיע מ-`paintedHost`. */
  host: HTMLElement | null;
  /** האלמנט שביחס אליו נמדדים המלבנים — שכבת הציור שלנו עצמה. */
  reference: HTMLElement | null;
  /** ה-controller, בשביל `viewport.observe`. */
  ui?: ViewportSource | null;
  onChange: (rects: readonly PageContentRects[]) => void;
}

/**
 * כמו `watchAllPageRects`, על מלבני הטקסט הגולמיים של כל עמוד — אותה תשתית
 * מעקב בדיוק (גלילה, שינוי גודל, `viewport.observe`, מדידות התיישבות
 * ב-`SETTLE_DELAYS_MS`). כפילות מכוונת ולא הפשטה משותפת: `watchAllPageRects`
 * כבר כפל את `watchPageRect` באותה צורה בדיוק, ושתי הפונקציות הקיימות
 * ממשיכות לעבוד — הפשטה משותפת כאן הייתה נוגעת בקוד נבדק כדי לשרת תכונה
 * שלישית, לא מתקנת דבר בשתיים הראשונות.
 */
export function watchAllPageContentRects(options: AllPageContentRectsWatchOptions): PageRectWatch {
  const { host, reference, ui, onChange } = options;
  let last: readonly PageContentRects[] = [];
  let frame: number | null = null;
  let pending = false;
  let disposed = false;

  function measureNow(): void {
    if (disposed) return;
    const next = measureAllPageContentRects(host, reference);
    if (sameContentRects(next, last)) return;
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
/* ריצות טקסט גולמיות בכל עמוד — לסימני עיצוב (¶)                     */
/* ------------------------------------------------------------------ */

/**
 * ריצת טקסט אחת — כל Text node גלוי בעמוד, בסדר מסמך, כפי שהוא. קורא יחיד:
 * engine/formatting-marks-layer.ts, שמיישר את הרצף הזה מול `doc.blocks.list()`
 * כדי למקם ¶ בסוף כל פסקה — ראו ההנמקה המלאה שם, ואת המדידה שעומדת מאחוריה
 * ב-docs/superdoc-2.10-review.md ("סימני עיצוב... נחקר לעומק"). זו הרחבה של
 * אותה טכניקה בדיוק כמו `measurePageContentRects` למעלה — `TreeWalker`/`Range`
 * תקניים על תוכן העמוד, לא selector אל מבנה פנימי של המנוע — רק ברזולוציה של
 * צומת בודד במקום טווח-עמוד שלם, כי כאן צריך גם את הטקסט של כל צומת ולא רק
 * את מלבנו.
 */
export interface PageTextRun {
  /** `nodeValue` הגולמי של הצומת. */
  text: string;
  /** מלבן אחד לכל שורה חזותית שהצומת פרוס עליה (עטיפה), בסדר מסמך. */
  rects: readonly RawTextRect[];
  /**
   * מאיזה צד "אחרי התו האחרון" של **הריצה הזאת עצמה** — ראו `directionFromText`
   * למטה. לא כיוון הפסקה/המסמך: פסקה RTL עם ריצה שכולה אנגלית (למשל שם-פרטי
   * או מספר) מצוירת מבפנים LTR למרות שהפסקה שלה RTL — נמדד (docs/superdoc-2.10-review.md):
   * הסתמכות על `getComputedStyle(...).direction` בלבד מיקמה ¶ בקצה הלא-נכון
   * בדיוק במקרה הזה (טקסט אנגלי בתוך פסקה RTL כברירת מחדל).
   */
  direction: 'ltr' | 'rtl';
}

/**
 * תווים "חזקים" (bidi) בעברית/ערבית — קובעים RTL מעצם הסקריפט, בלי קשר
 * ל-CSS. הטווחים (כתובים כ-\u כדי שיהיו ודאיים ולא תלויי-encoding של הקובץ
 * עצמו): עברית (U+0590-U+05FF), ערבית ונגזרותיה (U+0600-U+06FF, U+0750-U+077F,
 * U+08A0-U+08FF), וצורות-הצגה של שתיהן (U+FB1D-U+FDFF, U+FE70-U+FEFF).
 */
const RTL_STRONG_CHAR =
  /[\u0590-\u05FF\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/;

/**
 * תווים "חזקים" בלטינית/יווני/קיריליות/CJK/הנגול — קובעים LTR מעצם הסקריפט.
 * לא צריך להיות ממצה: כש-`directionFromText` לא מוצאת תו חזק כלל, הקורא
 * נופל ל-CSS (`getComputedStyle`), לא ל-`ltr` קשיח.
 */
const LTR_STRONG_CHAR =
  /[A-Za-z\u00C0-\u024F\u0370-\u03FF\u0400-\u04FF\u4E00-\u9FFF\uAC00-\uD7A3]/;

/**
 * כיוון הקריאה החזותי בפועל של **התוכן עצמו** — לא של הפסקה שהוא יושב בה.
 * סורקת מהסוף להתחלה ומחזירה את הכיוון של התו-החזק האחרון (כך שסוגריים/
 * פיסוק/רווח בסוף המחרוזת לא קובעים כיוון). `null` כשאין אף תו חזק (למשל
 * רק ספרות/פיסוק, או ה-placeholder של פסקה ריקה) — הקורא נופל אז ל-CSS.
 */
export function directionFromText(text: string): 'ltr' | 'rtl' | null {
  for (let i = text.length - 1; i >= 0; i--) {
    const ch = text[i]!;
    if (RTL_STRONG_CHAR.test(ch)) return 'rtl';
    if (LTR_STRONG_CHAR.test(ch)) return 'ltr';
  }
  return null;
}

/** ריצות הטקסט הגולמיות של עמוד אחד, ביחס ל-`reference`. */
export function measurePageTextRuns(pageEl: HTMLElement, reference: HTMLElement): readonly PageTextRun[] {
  if (typeof document === 'undefined' || typeof document.createTreeWalker !== 'function') return [];

  let walker: TreeWalker;
  try {
    walker = document.createTreeWalker(pageEl, NodeFilter.SHOW_TEXT);
  } catch {
    return [];
  }

  const referenceBox = reference.getBoundingClientRect();
  const out: PageTextRun[] = [];

  let node = walker.nextNode();
  while (node) {
    const value = node.nodeValue ?? '';
    if (value.length === 0) {
      node = walker.nextNode();
      continue;
    }

    let range: Range;
    try {
      range = document.createRange();
      range.selectNodeContents(node);
    } catch {
      node = walker.nextNode();
      continue;
    }

    const rects: RawTextRect[] = [];
    const list = range.getClientRects();
    for (let i = 0; i < list.length; i++) {
      const r = list[i]!;
      if (!(r.width > 0) || !(r.height > 0)) continue;
      rects.push({
        leftPx: r.left - referenceBox.left,
        topPx: r.top - referenceBox.top,
        widthPx: r.width,
        heightPx: r.height,
      });
    }

    if (rects.length > 0) {
      // קודם התוכן עצמו (ראו directionFromText) — ורק כשאין בו תו חזק כלל
      // (ספרות/פיסוק בלבד, או ה-placeholder של פסקה ריקה) נופלים ל-CSS
      // המחושב של האלמנט המכיל.
      let direction = directionFromText(value);
      if (direction === null) {
        direction = 'ltr';
        const container = node.parentElement;
        if (container && typeof getComputedStyle === 'function') {
          try {
            direction = getComputedStyle(container).direction === 'rtl' ? 'rtl' : 'ltr';
          } catch {
            direction = 'ltr';
          }
        }
      }
      out.push({ text: value, rects, direction });
    }

    node = walker.nextNode();
  }

  return out;
}

/**
 * ריצות הטקסט הגולמיות של **כל** העמודים, משורשרות לרשימה אחת בסדר מסמך —
 * ולכן לפי `data-page-index` עולה, לא לפי סדר ה-DOM (וירטואליזציה יכולה
 * לצייר עמודים שלא לפי סדר ההוספה). בלי המיון הזה יישור הרצף מול
 * `blocks.list()` (שתמיד בסדר מסמך) היה נשבר במסמך רב-עמודי.
 */
export function measureAllPageTextRuns(
  host: HTMLElement | null,
  reference: HTMLElement | null,
): readonly PageTextRun[] {
  if (!host || !reference) return [];
  const pages = host.querySelectorAll(`[${PAGE_INDEX_ATTRIBUTE}]`);
  if (pages.length === 0) return [];

  const ordered = Array.from(pages)
    .filter((node): node is HTMLElement => node instanceof HTMLElement)
    .map((el) => ({ el, pageIndex: Number(el.getAttribute(PAGE_INDEX_ATTRIBUTE)) }))
    .filter((p) => Number.isInteger(p.pageIndex))
    .sort((a, b) => a.pageIndex - b.pageIndex);

  const out: PageTextRun[] = [];
  for (const { el } of ordered) {
    out.push(...measurePageTextRuns(el, reference));
  }
  return out;
}

/** Pages intersecting the viewport plus a short scroll-ahead margin. */
export function visiblePageElements(host: HTMLElement, marginPx = 400): readonly HTMLElement[] {
  const viewport = host.getBoundingClientRect();
  return Array.from(host.querySelectorAll(`[${PAGE_INDEX_ATTRIBUTE}]`))
    .filter((node): node is HTMLElement => node instanceof HTMLElement)
    .map((el) => ({ el, index: Number(el.getAttribute(PAGE_INDEX_ATTRIBUTE)) }))
    .filter(({ el, index }) => {
      if (!Number.isInteger(index)) return false;
      const box = el.getBoundingClientRect();
      return box.bottom >= viewport.top - marginPx && box.top <= viewport.bottom + marginPx;
    })
    .sort((a, b) => a.index - b.index)
    .map(({ el }) => el);
}

/**
 * שתי רשימות מלבנים גולמיים שקולות. חצי פיקסל הוא הרזולוציה שכל המדידות כאן
 * עובדות בה — מתחתיה זו רעידה של `getClientRects` ולא תזוזה.
 */
function sameRawRects(a: readonly RawTextRect[], b: readonly RawTextRect[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const x = a[i]!;
    const y = b[i]!;
    if (
      Math.abs(x.leftPx - y.leftPx) >= 0.5 ||
      Math.abs(x.topPx - y.topPx) >= 0.5 ||
      Math.abs(x.widthPx - y.widthPx) >= 0.5 ||
      Math.abs(x.heightPx - y.heightPx) >= 0.5
    ) {
      return false;
    }
  }
  return true;
}

/** כמו `sameContentRects`, על רשימת ריצות-טקסט שטוחה. */
function sameTextRuns(a: readonly PageTextRun[], b: readonly PageTextRun[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const x = a[i]!;
    const y = b[i]!;
    if (x.text !== y.text || x.direction !== y.direction) return false;
    if (!sameRawRects(x.rects, y.rects)) return false;
  }
  return true;
}

export interface AllPageTextRunsWatchOptions {
  /** מיכל הגלילה שהמנוע מצייר בתוכו. מגיע מ-`paintedHost`. */
  host: HTMLElement | null;
  /** האלמנט שביחס אליו נמדדות הריצות — שכבת הציור שלנו עצמה. */
  reference: HTMLElement | null;
  /** ה-controller, בשביל `viewport.observe`. */
  ui?: ViewportSource | null;
  onChange: (runs: readonly PageTextRun[]) => void;
}

/**
 * כמו `watchAllPageContentRects`, על ריצות הטקסט הגולמיות של כל עמוד — אותה
 * תשתית מעקב בדיוק (גלילה, שינוי גודל, `viewport.observe`, מדידות התיישבות
 * ב-`SETTLE_DELAYS_MS`). כפילות מכוונת, מאותה סיבה שכבר הוסברה שם.
 */
export function watchAllPageTextRuns(options: AllPageTextRunsWatchOptions): PageRectWatch {
  const { host, reference, ui, onChange } = options;
  let last: readonly PageTextRun[] = [];
  let frame: number | null = null;
  let pending = false;
  let disposed = false;

  function measureNow(): void {
    if (disposed) return;
    const next = measureAllPageTextRuns(host, reference);
    if (sameTextRuns(next, last)) return;
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
/* טווחים בתוך שורת טקסט — לבדיקת האיות                                */
/* ------------------------------------------------------------------ */

/**
 * הקורא הרביעי, וזה שהוא צריך שונה מכל השלושה שלפניו: לא מלבן של עמוד, לא
 * מלבן של שורה ולא מלבן של צומת שלם — אלא מלבן של **קטע בתוך צומת**, כדי
 * למתוח קו גלי מתחת למילה אחת (ui/shell/SpellingOverlay.vue).
 *
 * ## למה מכאן, ולא דרך `ui.viewport.getRect`
 *
 * למנוע **יש** API ציבורי שפותר טווח טקסט לגיאומטריה: `ui.viewport.getRect
 * ({ target, relativeTo })` עם `SelectionTarget` של בלוק+היסטים, בדיוק אותו
 * יעד ש-engine/search.ts כבר בונה. הוא נמדד ועובד (Chrome headless על ה-dist
 * הארוז): 288 טווחים, כולם נפתרו, מלבן לכל אחד.
 *
 * ומה שנמדד יחד איתו הוא הסיבה שהוא לא נבחר: **0.15ms לקריאה**. עמוד עברי
 * ממוצע הוא ~400 מילים, מהן ~5% אינן במילון — 20 קריאות לעמוד זה עוד סביר,
 * אבל הסימון נמדד מחדש בכל גלילה, וב„הצג הכול” על מסמך בן עשרות עמודים זה
 * מגיע למאות מילישניות לפריים. המסלול כאן — TreeWalker ‏+ `Range` על אותו
 * תוכן — נמדד **1.2ms לעמוד שלם** (31 צמתים, 480 מילים נסרקו, 24 טווחים
 * נמדדו): פי 60 פחות, ובלי להישען על שום מבנה פנימי של המנוע.
 *
 * זו אותה טכניקה תקנית בדיוק שכבר עומדת מאחורי `measurePageContentRects`
 * ו-`measurePageTextRuns` למעלה, רק ברזולוציה של קטע בתוך צומת. `getRect`
 * נשאר המסלול הנכון לקריאה **בודדת** (למשל „גלול אל השגיאה הבאה”), ולא
 * למאות בפריים.
 *
 * ## קיבוץ: „מה נחשב שורת טקסט אחת”
 *
 * מילה יכולה להתפצל לשני צמתים כשהעיצוב משתנה באמצעה (חצי מודגש), וסריקה
 * צומת-צומת הייתה מסמנת שתי שגיאות במקום מילה אחת מוכרת. לכן צמתים עוקבים
 * מצורפים לפי **האב הראשון שאינו אלמנט inline** — לפי שם התגית, ולא לפי שם
 * מחלקה של המנוע (‏tests/unit/engine-boundaries.test.ts אוסר, ובצדק: המנוע
 * מתעד את שמות המחלקות שלו כחוזה פנימי). בפועל הטיפוס הוא צעד אחד: הצומת
 * יושב ב-`<span>`, וההורה שלו הוא כבר `<div>`.
 *
 * **תגית ולא `getComputedStyle`**, ושלוש סיבות: היא אינה עולה כלום בלולאה
 * שרצה על כל צומת בעמוד; היא נותנת אותה תשובה בכל סביבה (ב-jsdom
 * `getComputedStyle(span).display` הוא מחרוזת ריקה, כלומר בדיקה שרצה שם
 * הייתה מודדת התנהגות שאינה קיימת בדפדפן); ותגית שאינה מוכרת מטופלת כבלוק —
 * הכיוון הבטוח, שחותך קבוצה במקום להדביק שתי שורות למילה שאינה קיימת.
 *
 * הגבול הזה גם מספיק: המנוע מצייר **בלוק לכל שורה חזותית** (נמדד), ומילה
 * אינה נשברת בין שורות — כלומר קבוצה אחת מכילה מילים שלמות בלבד.
 *
 * ## היקף: רק העמודים שרואים
 *
 * `measureAllPageTextRuns` סורק את כל העמודים, כי הוא מיישר רצף מול
 * `blocks.list()` וחייב להיות שלם. כאן אין רצף לשמור — כל קבוצה עומדת בפני
 * עצמה — ולכן נסרקים רק עמודים שנחתכים עם החלון (בתוספת `marginPx`). זה מה
 * שהופך מסמך בן שמונים עמודים לאותה עלות כמו מסמך בן שניים.
 */

/** טווח בקואורדינטות-הטקסט של קבוצה אחת, כפי ש-`select` מחזירה. */
export interface TextSegment {
  readonly start: number;
  readonly end: number;
}

/** טווח שנמדד: הטקסט שלו, ומלבן לכל תיבה שהוא נפרס עליה. */
export interface MeasuredSegment {
  readonly text: string;
  readonly rects: readonly RawTextRect[];
}

export interface TextSegmentOptions {
  /** כמה פיקסלים מעבר לחלון עדיין נסרקים, כדי שגלילה קצרה לא תגלה שטח ריק. */
  readonly marginPx?: number;
  /** תקרת טווחים. מסמך פתולוגי לא יהפוך פריים אחד למאות מילישניות. */
  readonly limit?: number;
}

const SEGMENT_MARGIN_PX = 400;
const SEGMENT_LIMIT = 600;

/**
 * תגיות שאינן שוברות שורת טקסט — כלומר צמתים משני צידיהן מצטרפים. אלה
 * העוטפים שסימוני העיצוב של ProseMirror מייצרים (מודגש, נטוי, קישור, כתב
 * עילי/תחתי), בתוספת שאר ה-inline של HTML כדי שמסמך שהגיע מ-Word עם עוטף
 * פחות שכיח לא ייחתך לשווא.
 */
const INLINE_TAGS: ReadonlySet<string> = new Set([
  'SPAN', 'A', 'B', 'STRONG', 'I', 'EM', 'U', 'S', 'STRIKE', 'DEL', 'INS',
  'SUB', 'SUP', 'MARK', 'CODE', 'SMALL', 'BIG', 'TT', 'ABBR', 'CITE', 'Q',
  'BDI', 'BDO', 'FONT', 'RUBY', 'RT', 'RP', 'VAR', 'SAMP', 'KBD', 'TIME',
]);

/** האב הראשון שאינו אלמנט inline. ראו הערת הראש — תגית, ולא סגנון מחושב. */
function blockAncestor(element: Element, root: Element): Element {
  let current: Element | null = element;
  while (current && current !== root && INLINE_TAGS.has(current.tagName)) {
    current = current.parentElement;
  }
  return current ?? root;
}

/** חלק אחד בקבוצה: הצומת, וההיסט שבו הוא מתחיל בטקסט המחובר. */
interface SegmentPart {
  readonly node: Text;
  readonly offset: number;
}

/**
 * מקום בתוך קבוצה ⟵ (צומת, היסט בתוכו). `null` = ההיסט מחוץ לקבוצה.
 *
 * `select` הוא קלט חיצוני, ולא כל מי שיכתוב אחד יחזיר טווחים בתוך הטקסט
 * שנמסר לו. טווח שחורג פשוט מדולג — עדיף מ-`setEnd` שזורק ומפיל את מדידת
 * כל שאר העמוד.
 */
function locate(parts: readonly SegmentPart[], offset: number): { node: Text; at: number } | null {
  for (let i = parts.length - 1; i >= 0; i--) {
    const part = parts[i]!;
    if (offset >= part.offset) {
      const at = offset - part.offset;
      return at <= (part.node.nodeValue?.length ?? 0) ? { node: part.node, at } : null;
    }
  }
  return null;
}

/** הטווחים שנבחרו בעמוד אחד, ביחס ל-`reference`. */
function measurePageTextSegments(
  pageEl: HTMLElement,
  referenceBox: DOMRect,
  select: (text: string) => readonly TextSegment[],
  out: MeasuredSegment[],
  limit: number,
): void {
  let walker: TreeWalker;
  try {
    walker = document.createTreeWalker(pageEl, NodeFilter.SHOW_TEXT);
  } catch {
    return;
  }

  let parts: SegmentPart[] = [];
  let text = '';
  let group: Element | null = null;

  function flush(): void {
    if (text.length > 0 && out.length < limit) {
      for (const segment of select(text)) {
        if (out.length >= limit) break;
        const from = locate(parts, segment.start);
        const to = locate(parts, segment.end);
        if (!from || !to) continue;

        let range: Range;
        try {
          range = document.createRange();
          range.setStart(from.node, from.at);
          range.setEnd(to.node, to.at);
        } catch {
          continue;
        }

        const rects: RawTextRect[] = [];
        const list = range.getClientRects();
        for (let i = 0; i < list.length; i++) {
          const rect = list[i]!;
          if (!(rect.width > 0) || !(rect.height > 0)) continue;
          rects.push({
            leftPx: rect.left - referenceBox.left,
            topPx: rect.top - referenceBox.top,
            widthPx: rect.width,
            heightPx: rect.height,
          });
        }
        if (rects.length > 0) out.push({ text: text.slice(segment.start, segment.end), rects });
      }
    }
    parts = [];
    text = '';
  }

  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const value = node.nodeValue ?? '';
    if (value.length === 0) continue;

    const parent = (node as Text).parentElement;
    const owner = parent ? blockAncestor(parent, pageEl) : pageEl;
    if (owner !== group) {
      flush();
      group = owner;
    }
    parts.push({ node: node as Text, offset: text.length });
    text += value;
  }
  flush();
}

/**
 * הטווחים שנבחרו בעמודים הגלויים, ביחס ל-`reference`.
 *
 * `select` מקבלת את הטקסט המחובר של קבוצה אחת ומחזירה טווחים בתוכו — כך
 * המודול הזה נשאר גיאומטריה בלבד, ומי שיודע *מה* לסמן (engine/spellcheck.ts)
 * אינו נוגע ב-DOM.
 */
export function measureAllPageTextSegments(
  host: HTMLElement | null,
  reference: HTMLElement | null,
  select: (text: string) => readonly TextSegment[],
  options: TextSegmentOptions = {},
): readonly MeasuredSegment[] {
  if (!host || !reference) return [];
  if (typeof document === 'undefined' || typeof document.createTreeWalker !== 'function') return [];

  const pages = host.querySelectorAll(`[${PAGE_INDEX_ATTRIBUTE}]`);
  if (pages.length === 0) return [];

  const margin = options.marginPx ?? SEGMENT_MARGIN_PX;
  const limit = options.limit ?? SEGMENT_LIMIT;
  const viewport = host.getBoundingClientRect();
  const top = viewport.top - margin;
  const bottom = viewport.bottom + margin;

  // פעם אחת למעבר ולא פעם לכל עמוד: `getBoundingClientRect` על אלמנט שלא
  // נגע בו כלום הוא עדיין קריאה שמכריחה פריסה, והמעבר הזה רץ בכל גלילה.
  const referenceBox = reference.getBoundingClientRect();
  const out: MeasuredSegment[] = [];

  const ordered = Array.from(pages)
    .filter((node): node is HTMLElement => node instanceof HTMLElement)
    .map((el) => ({ el, pageIndex: Number(el.getAttribute(PAGE_INDEX_ATTRIBUTE)) }))
    .filter((page) => Number.isInteger(page.pageIndex))
    .sort((a, b) => a.pageIndex - b.pageIndex);

  for (const { el } of ordered) {
    if (out.length >= limit) break;
    const box = el.getBoundingClientRect();
    if (box.bottom < top || box.top > bottom) continue;
    measurePageTextSegments(el, referenceBox, select, out, limit);
  }

  return out;
}

/* ------------------------------------------------------------------ */
/* גליפים של עמוד — להצמדת לחיצה (engine/pointer-snap.ts)             */
/* ------------------------------------------------------------------ */

/** מלבני הגליפים ביחס לפינת העמוד ברגע המדידה, והעמוד עצמו — עכשיו. */
export interface PageGlyphMeasure {
  readonly rects: readonly RawTextRect[];
  /** מלבן העמוד ברגע הקריאה, בקואורדינטות חלון — גלילה מזיזה אותו. */
  pageBox(): RawTextRect;
}

/** אלמנטים מוחלפים: לחיצה עליהם היא בחירת אובייקט, לא של טקסט. */
const REPLACED_TAGS = new Set(['IMG', 'SVG', 'CANVAS', 'VIDEO', 'PICTURE', 'OBJECT', 'IFRAME', 'EMBED']);

/**
 * האם נקודה בקואורדינטות החלון נופלת על עמוד מצויר.
 *
 * `null` פירושו „אין תשובה” — אין עמודים בעץ (לפני שהמסמך נצבע, או אחרי
 * פירוק), ואז מי שקורא לא יסיק „מחוץ לעמוד” ממה שהוא בעצם „עוד לא יודע”.
 *
 * אותו עיגון ואותה קריאה טהורה כמו `measureAllPageRects`; זו הגרסה שמשרתת
 * מאזין לחיצה, ולכן היא יוצאת בעמוד הראשון שמכיל את הנקודה ואינה אוספת
 * מלבנים. הצורך: engine/click-focus.ts.
 */
export function pointOverPage(host: HTMLElement | null, xPx: number, yPx: number): boolean | null {
  if (!host) return null;
  const pages = host.querySelectorAll(`[${PAGE_INDEX_ATTRIBUTE}]`);
  if (pages.length === 0) return null;
  for (const node of pages) {
    if (!(node instanceof HTMLElement)) continue;
    const box = node.getBoundingClientRect();
    if (!(box.width > 0) || !(box.height > 0)) continue;
    if (xPx >= box.left && xPx <= box.right && yPx >= box.top && yPx <= box.bottom) return true;
  }
  return false;
}

/** התכונה שהמנוע מסמן בה לאיזה „סיפור” שייך פרגמנט: גוף, כותרת, הערה. */
const STORY_ATTRIBUTE = 'data-layout-story';

function rawBox(rect: DOMRect): RawTextRect {
  return { leftPx: rect.left, topPx: rect.top, widthPx: rect.width, heightPx: rect.height };
}

/**
 * מלבני הגליפים של צומת טקסט. ה-`Range` מגיע מבחוץ ומשמש מחדש: מדידת עמוד
 * שלם עוברת על מאות צומתי טקסט, ו-`createRange` לכל אחד מהם היה מקצה מאות
 * טווחים חיים בתוך מאזין של `mousedown`. `null` כשה-Range אינו יכול לכסות
 * את הצומת.
 */
function glyphRects(node: Text, range: Range): DOMRectList | null {
  try {
    range.selectNodeContents(node);
  } catch {
    return null;
  }
  return typeof range.getClientRects === 'function' ? range.getClientRects() : null;
}

function pushGlyphRects(node: Text, out: RawTextRect[], origin: DOMRect, range: Range): boolean {
  const list = glyphRects(node, range);
  if (!list) return false;
  let found = false;
  for (let i = 0; i < list.length; i++) {
    const rect = list[i]!;
    if (!(rect.width > 0) || !(rect.height > 0)) continue;
    found = true;
    out.push({ leftPx: rect.left - origin.left, topPx: rect.top - origin.top, widthPx: rect.width, heightPx: rect.height });
  }
  return found;
}

/**
 * האם יש תחת האלמנט טקסט מצויר — כלומר זה היקף שאפשר להצמיד אליו. יוצאת
 * במלבן הראשון ואינה אוספת כלום: התשובה היא כן/לא.
 */
function hasGlyphs(element: Element, range: Range): boolean {
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if ((node.nodeValue ?? '').length === 0) continue;
    const list = glyphRects(node as Text, range);
    if (!list) continue;
    for (let i = 0; i < list.length; i++) {
      const rect = list[i]!;
      if (rect.width > 0 && rect.height > 0) return true;
    }
  }
  return false;
}

/** צומת טקסט שאינו בגוף המסמך — כותרת עליונה/תחתונה, הערה — כשהתכונה קיימת. */
function outsideBody(node: Text): boolean {
  const story = node.parentElement?.closest(`[${STORY_ATTRIBUTE}]`)?.getAttribute(STORY_ATTRIBUTE);
  return typeof story === 'string' && story !== '' && story !== 'body';
}

/**
 * מלבני הגליפים שלחיצה בנקודה מוצמדת אליהם — ראו engine/pointer-snap.ts.
 *
 * ההיקף נגזר מ-`target`, מטרת האירוע:
 *
 *   - בתוך עמוד: האלמנט הראשון במעלה העץ שיש תחתיו טקסט מצויר — השורה
 *     שנלחצה, תא הטבלה, הפסקה. אלמנט מוחלף בדרך (תמונה) — `null`: זו
 *     בחירת אובייקט, ואין להזיז אותה. מטרה בלי טקסט תחתיה ושאינה העמוד
 *     עצמו (ידית, מסגרת) — גם כן `null`.
 *   - העמוד עצמו, או אלמנט שמכסה את העמוד כולו (שכבה של המנוע, ה-host):
 *     כל גליפי **הגוף** של העמוד שמתחת לנקודה — לא כותרות ולא הערות,
 *     כשהמנוע מסמן אותן.
 *   - `null` כמטרה — גרירה: העמוד שמתחת לנקודה, כל גליפי הגוף.
 *
 * קריאה בלבד, אותו עיגון כמו `measureAllPageRects`.
 */
export function measurePageGlyphs(
  host: HTMLElement | null,
  target: EventTarget | null,
  xPx: number,
  yPx: number,
): PageGlyphMeasure | null {
  if (!host) return null;
  if (typeof document === 'undefined' || typeof document.createTreeWalker !== 'function') return null;
  if (typeof document.createRange !== 'function') return null;
  const pages = Array.from(host.querySelectorAll(`[${PAGE_INDEX_ATTRIBUTE}]`)).filter(
    (node): node is HTMLElement => node instanceof HTMLElement,
  );
  if (pages.length === 0) return null;

  // Range אחד לכל המדידה — ראו `glyphRects`.
  let range: Range;
  try {
    range = document.createRange();
  } catch {
    return null;
  }

  const targetEl = target instanceof Element ? target : target instanceof Node ? target.parentElement : null;
  let page = targetEl ? (pages.find((candidate) => candidate.contains(targetEl)) ?? null) : null;
  let scope: Element | null = null;

  if (page && targetEl) {
    if (targetEl === page) {
      scope = page;
    } else {
      for (let el: Element | null = targetEl; el && el !== page; el = el.parentElement) {
        if (REPLACED_TAGS.has(el.tagName.toUpperCase())) return null;
        if (hasGlyphs(el, range)) {
          scope = el;
          break;
        }
      }
      if (!scope) return null;
    }
  } else {
    page = pages.find((candidate) => {
      const box = candidate.getBoundingClientRect();
      return xPx >= box.left && xPx <= box.right && yPx >= box.top && yPx <= box.bottom;
    }) ?? null;
    if (!page) return null;
    if (targetEl && targetEl !== host) {
      // מטרה מחוץ לעמוד שאינה מכסה אותו — פקד של המנוע מעל הדף. לא שלנו.
      const targetBox = targetEl.getBoundingClientRect();
      const pageBox = page.getBoundingClientRect();
      if (
        targetBox.left > pageBox.left + 1 ||
        targetBox.top > pageBox.top + 1 ||
        targetBox.right < pageBox.right - 1 ||
        targetBox.bottom < pageBox.bottom - 1
      ) {
        return null;
      }
    }
    scope = page;
  }

  const origin = page.getBoundingClientRect();
  const rects: RawTextRect[] = [];
  const bodyOnly = scope === page;
  let walker: TreeWalker;
  try {
    walker = document.createTreeWalker(scope, NodeFilter.SHOW_TEXT);
  } catch {
    return null;
  }
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if ((node.nodeValue ?? '').length === 0) continue;
    if (bodyOnly && outsideBody(node as Text)) continue;
    pushGlyphRects(node as Text, rects, origin, range);
  }

  const pageEl = page;
  return { rects, pageBox: () => rawBox(pageEl.getBoundingClientRect()) };
}

/** שני מערכי טווחים שקולים — כדי לא לצייר מחדש מדידה שלא זזה. */
export function sameTextSegments(a: readonly MeasuredSegment[], b: readonly MeasuredSegment[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const x = a[i]!;
    const y = b[i]!;
    if (x.text !== y.text || !sameRawRects(x.rects, y.rects)) return false;
  }
  return true;
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
  /** סרגל מוסתר אינו קורא את המסמך ואינו מחזיק תזמון רענון. */
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

/** שינוי תוכן בלבד יכול להמתין; אירוע בחירה באותו רצף מקצר את ההשהיה. */
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
  let timer: ReturnType<typeof setTimeout> | undefined;
  let scheduledDelay = Infinity;
  let running = false;
  let rerun = false;

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
      sameParagraphTarget(a.target, b.target)
    );
  }

  function publish(next: RulerReading | null): void {
    if (disposed || same(next, reading)) return;
    reading = next;
    source.onChange(next);
  }

  function cancelTimer(): void {
    clearTimeout(timer);
    timer = undefined;
    scheduledDelay = Infinity;
  }

  async function read(): Promise<void> {
    if (disposed || !enabled) return;
    if (running) {
      rerun = true;
      return;
    }
    running = true;
    const mine = generation;
    try {
      let page: PageMarginsState | null = null;
      try {
        page = await source.readPage();
      } catch {
        // A failed background read clears the readout without reporting a command error.
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
        // The page is still useful when there is no readable paragraph selection.
      }
      if (disposed || mine !== generation) return;
      publish({ page, indents: paragraph?.indents ?? null, target: paragraph?.target ?? null });
    } finally {
      running = false;
      // Keep at most one read in flight and one request to read the latest state.
      if (rerun && timer === undefined && !disposed && enabled) {
        rerun = false;
        void read();
      }
    }
  }

  function schedule(delay: number): void {
    if (disposed || !enabled) return;
    // Invalidate immediately, not only when the debounce expires.
    generation += 1;
    rerun = false;
    scheduledDelay = Math.min(scheduledDelay, delay);
    clearTimeout(timer);
    timer = setTimeout(() => {
      timer = undefined;
      scheduledDelay = Infinity;
      void read();
    }, scheduledDelay);
  }

  function refreshNow(): void {
    generation += 1;
    cancelTimer();
    void read();
  }

  return {
    getState: () => reading,

    setEnabled(next) {
      if (disposed || enabled === next) return;
      enabled = next;
      generation += 1;
      cancelTimer();
      rerun = false;
      if (!enabled) {
        publish(null);
        return;
      }
      void read();
    },

    noteSelectionChanged: () => schedule(RULER_SELECTION_DEBOUNCE_MS),
    noteDocumentChanged: () => schedule(RULER_DOCUMENT_DEBOUNCE_MS),
    refreshNow,

    dispose() {
      disposed = true;
      generation += 1;
      rerun = false;
      cancelTimer();
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

/* ------------------------------------------------------------------ */
/* ספירת עמודים ומילות הקצה — לכלי „שולחן העורך”                      */
/* ------------------------------------------------------------------ */

/**
 * מספר העמודים **המצוירים**: `max(data-page-index) + 1`. `null` כשאין
 * עמוד מצויר. זה הקירוב היחיד שיש לספירת העמודים של המסמך — למנוע אין
 * API ציבורי לזה (docs/shulchan-source/engine-issues/3970-layout-read-api.md),
 * ו-`totalPages` של `onPaginationUpdate` מגיע בלי התחייבות לעיתוי.
 */
export function countPaintedPages(host: HTMLElement | null): number | null {
  if (!host) return null;
  const pages = host.querySelectorAll(`[${PAGE_INDEX_ATTRIBUTE}]`);
  let max = -1;
  pages.forEach((node) => {
    const index = Number(node.getAttribute(PAGE_INDEX_ATTRIBUTE));
    if (Number.isInteger(index) && index > max) max = index;
  });
  return max >= 0 ? max + 1 : null;
}

/**
 * ספירת עמודים אחרי התיישבות: המנוע מעמד מחדש א-סינכרונית אחרי מוטציה,
 * וספירה מיידית רואה את הפריסה הקודמת. הסולם הוא `SETTLE_DELAYS_MS` של
 * הסרגל — ונעצר מוקדם כששתי קריאות עוקבות מסכימות. הקריאה האחרונה היא
 * התשובה; `null` כשאף פעם לא צויר עמוד.
 */
export async function settledPageCount(
  host: HTMLElement | null,
  delays: readonly number[] = SETTLE_DELAYS_MS,
  wait: (ms: number) => Promise<void> = (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
): Promise<number | null> {
  let last = countPaintedPages(host);
  for (const delay of delays) {
    await wait(delay);
    const next = countPaintedPages(host);
    if (next === last && next !== null) return next;
    last = next;
  }
  return last;
}

/** מילה בקצה עמוד: הטקסט שלה והמלבנים שהיא נפרסת עליהם, ביחס ל-`reference`. */
export interface EdgeWord {
  text: string;
  rects: readonly RawTextRect[];
}

/** המילה הראשונה והאחרונה של עמוד מצויר, והטקסט הפותח שלו. */
export interface PageEdgeWords {
  pageIndex: number;
  /** תחילת הטקסט של העמוד (עד 200 תווים גולמיים) — לזיהוי העמוד בתצלום. */
  head: string;
  first: EdgeWord | null;
  last: EdgeWord | null;
}

const HEAD_RAW_LENGTH = 200;

interface TextPart {
  node: Text;
  offset: number;
}

function rectsOf(from: TextPart | null, to: TextPart | null, start: number, end: number, referenceBox: DOMRect): RawTextRect[] {
  if (!from || !to) return [];
  try {
    const range = document.createRange();
    range.setStart(from.node, start - from.offset);
    range.setEnd(to.node, end - to.offset);
    const out: RawTextRect[] = [];
    const list = range.getClientRects();
    for (let i = 0; i < list.length; i++) {
      const r = list[i]!;
      if (!(r.width > 0) || !(r.height > 0)) continue;
      out.push({ leftPx: r.left - referenceBox.left, topPx: r.top - referenceBox.top, widthPx: r.width, heightPx: r.height });
    }
    return out;
  } catch {
    return [];
  }
}

function partAt(parts: readonly TextPart[], offset: number): TextPart | null {
  for (let i = parts.length - 1; i >= 0; i--) {
    const part = parts[i]!;
    if (offset >= part.offset) return part;
  }
  return null;
}

/**
 * המילה הראשונה והאחרונה בכל עמוד מצויר — ל„סימון עמודים” של שולחן העורך.
 *
 * הטקסט של העמוד נקרא ב-`TreeWalker` על צמתי הטקסט שלו, בסדר המסמך (זה
 * כולל כותרת/כותרת תחתונה ומספרי עמוד אם הם צמתי טקסט — הקירוב מתועד
 * ב-engine/shulchan/page-marking.ts). „מילה” = רצף תווים שאינם רווח.
 * קריאה בלבד, אותו עיגון כמו `measureAllPageRects`.
 */
export function measurePageEdgeWords(
  host: HTMLElement | null,
  reference: HTMLElement | null,
): readonly PageEdgeWords[] {
  if (!host || !reference) return [];
  if (typeof document === 'undefined' || typeof document.createTreeWalker !== 'function') return [];
  const pages = host.querySelectorAll(`[${PAGE_INDEX_ATTRIBUTE}]`);
  if (pages.length === 0) return [];
  const referenceBox = reference.getBoundingClientRect();

  const ordered = Array.from(pages)
    .filter((node): node is HTMLElement => node instanceof HTMLElement)
    .map((el) => ({ el, pageIndex: Number(el.getAttribute(PAGE_INDEX_ATTRIBUTE)) }))
    .filter((page) => Number.isInteger(page.pageIndex))
    .sort((a, b) => a.pageIndex - b.pageIndex);

  const out: PageEdgeWords[] = [];
  for (const { el, pageIndex } of ordered) {
    let walker: TreeWalker;
    try {
      walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    } catch {
      continue;
    }
    const parts: TextPart[] = [];
    let text = '';
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const value = node.nodeValue ?? '';
      if (value.length === 0) continue;
      parts.push({ node: node as Text, offset: text.length });
      text += value;
    }

    const firstMatch = /\S+/.exec(text);
    const lastMatch = /\S+\s*$/.exec(text);
    const edge = (match: RegExpExecArray | null): EdgeWord | null => {
      if (!match) return null;
      const word = match[0].trim();
      const start = match.index;
      const end = start + word.length;
      return { text: word, rects: rectsOf(partAt(parts, start), partAt(parts, Math.max(start, end - 1)), start, end, referenceBox) };
    };
    out.push({
      pageIndex,
      head: text.slice(0, HEAD_RAW_LENGTH),
      first: edge(firstMatch),
      last: edge(lastMatch),
    });
  }
  return out;
}
