/**
 * מה שהמנוע צייר, כמקור מידע: עמודים, פסקאות, שורות, והתיבה של כל תו.
 *
 * המנוע מצייר את המסמך ל-DOM אמיתי ותולה על כל אלמנט את טווח ה-pm שלו, ולכן
 * אפשר לקרוא מכאן **איפה כל דבר נמצא** בלי לדעת דבר על המודל הפנימי שלו. זו
 * התשתית של כל עקיפת סמן בריפו: היא חייבת להחליט **בתוך** ה-keydown לאן הסמן
 * צריך לזוז, וכל מקור מידע א-סינכרוני של המנוע מגיע מאוחר מדי.
 *
 * ## מה כאן ומה לא
 *
 * כאן — **עובדות** על מה שצויר: איזה fragment, איזו שורה, מי השכן, מה טווח
 * ה-pm של כל אחד, איפה מצוירת כל גרפמה, ובאיזה כיוון. שם — ב-`rtl-caret.ts`
 * וב-`rtl-vertical-caret.ts` — **מדיניות**: מה מקש מסוים צריך לעשות עם
 * העובדות האלה, ומתי לוותר ולמסור למנוע.
 *
 * הגבול הזה אינו קוסמטי. הכללים כאן נמדדו אחד-אחד מול המנוע — פסקה שנחצית בין
 * עמודים, פסקה ריקה שתופסת pm נוסף, טבלה בלי מזהה מקור, עמוד שלא צויר, תפר
 * דו-כיווני שיש לו שני בתים, אשכול גרפמה שאסור להיכנס לתוכו — וכל צרכן שמנחש
 * אותם מחדש מגלה אותם שוב דרך באג. ההסברים והמדידות נשארו צמודים לכל פונקציה,
 * כפי שנכתבו כשנמדדו.
 *
 * ## למה לא הכול
 *
 * ‏`rtl-line-end.ts` (מקש `End`) קדם לשני האחרים ומחזיק קורא משלו, מצומצם
 * יותר. הוא **לא** הוסב לכאן בכוונה: הוא יציב, מגודר בשער משלו, ואין בו באג —
 * והסבה היא סיכון בלי תמורה. אם יתווסף לו פער, זה המקום להסב אליו.
 */

/** תו מצויר אחד: ההיסט שלו (ב-pm של המנוע) והתיבה שלו על המסך. */
export interface PaintedChar {
  pm: number;
  left: number;
  right: number;
  /** אשכול הגרפמה עצמו — אות ואיתה הניקוד והטעמים שלה. חסר בטאב, שאין לו טקסט. */
  ch?: string;
  /**
   * כמה יחידות UTF-16 האשכול תופס. חסר = 1; ‏3 ב-„שָׁ”, 2 באימוג'י —
   * ההסבר ב-`readLineChars`.
   */
  units?: number;
}

/** יחידות ה-UTF-16 של התו. ראו `PaintedChar.units`. */
const unitsOf = (c: PaintedChar): number => c.units ?? 1;

/** מקום חוקי לסמן: ההיסט וה-x שבו הוא מצויר. */
export interface CaretSlot {
  pm: number;
  x: number;
}

/**
 * שתי תיבות נחשבות נוגעות עד לסף הזה. הן נמדדו נוגעות **בדיוק** (הפרש 0),
 * והסף קיים רק בשביל עיגול של תת-פיקסל.
 */
const TOUCH = 0.6;

/**
 * מערכות כתב שנקראות מימין לשמאל. אות שלהן היא R או AL ב-UBA.
 */
const RTL_SCRIPT =
  /[\p{Script=Hebrew}\p{Script=Arabic}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Adlam}]/u;

/** ספרה — EN או AN ב-UBA, ובשני המקרים נקראת משמאל לימין גם בשורה עברית. */
const DIGIT = /\p{Nd}/u;

/** אות — כולל סימני ניקוד שמצטרפים לאות (Other_Alphabetic). */
const LETTER = /\p{Alphabetic}/u;

/**
 * האם הדפדפן מסדר את התו הזה משמאל לימין גם בתוך שורה עברית.
 *
 * ספרה, או אות שאינה ממערכת כתב ימנית — כן. סימן — לא: ב-UBA הוא ניטרלי (ON),
 * הדפדפן מצייר אותו בכיוון הפסקה, ושני התפרים שלו נמצאים ב-x שונה — זו
 * אינה הנקודה העיוורת שלמטה, וקיפול שלהם מוחק מקום נגיש. נמדד: ב-„שלוש × ארבע” ההיסטים
 * 5 ו-6 מצוירים ב-1008.5 ו-999.4, וסיווג של × כלועזי הקפיץ את החץ מ-4 ישר ל-7.
 */
function isLtrIntrinsic(ch: string): boolean {
  if (ch === '') return false;
  return DIGIT.test(ch) || (LETTER.test(ch) && !RTL_SCRIPT.test(ch));
}

/**
 * הכיוון של כל תו.
 *
 * `true` = התו שייך לקטע ימין-לשמאל. הכיוון נגזר מהגאומטריה: שני תווים עוקבים
 * שהתיבות שלהם נוגעות שייכים לאותו קטע, והצד שבו הן נוגעות הוא הכיוון. תו
 * שאין לו שכן צמוד משני צדדיו — למשל רווח שמפריד בין עברית לאי לטיני, או
 * טאב — נופל לכיוון השורה.
 *
 * ## אי של תו אחד
 *
 * לגאומטריה יש נקודה עיוורת אחת: ספרה או אות לטינית **בודדת** בין שני תווים
 * עבריים („סעיף 3 בחוק”, „אות a אחת”) נוגעת בשכניה בדיוק כמו תו עברי, כי אין
 * לה שכן לטיני שיגלה את הכיוון שלה. היא סווגה כעברית, שני התפרים שלה נשארו
 * יעדים, והמנוע מצייר את שניהם באותו x — ולכן כל הקשה החזירה את הסמן לתפר
 * השני ולא עברה את התו (נמדד: „3” ב-[1005, 1013], והיסטים 5 ו-6 מצוירים שניהם
 * ב-1005; „a” ב-[1012.1, 1019.2], ו-4 ו-5 שניהם ב-1019.2). כשהגאומטריה אינה
 * מראה צמידות לטינית, התו עצמו מכריע.
 */
export function charDirections(chars: readonly PaintedChar[], lineRtl: boolean): boolean[] {
  const dirs: (boolean | null)[] = new Array(chars.length).fill(null);

  for (let i = 0; i + 1 < chars.length; i += 1) {
    const a = chars[i]!;
    const b = chars[i + 1]!;
    let rtl: boolean | null = null;
    if (Math.abs(b.left - a.right) < TOUCH) rtl = false;
    else if (Math.abs(a.left - b.right) < TOUCH) rtl = true;
    if (rtl === null) continue;
    if (dirs[i] === null) dirs[i] = rtl;
    dirs[i + 1] = rtl;
  }

  return dirs.map((value, i) => {
    if (value !== false && isLtrIntrinsic(chars[i]!.ch ?? '')) return false;
    return value ?? lineRtl;
  });
}

/**
 * כל המקומות שהסמן רשאי לעצור בהם בשורה, בלי אלה שיושבים על תפר.
 *
 * הסמן בהיסט o יושב על ה**קצה המוביל** של התו שבהיסט o — שמאל לתו לטיני,
 * ימין לתו עברי — ובסוף השורה על הקצה הנגרר של האחרון. היסט שהכיוון משתנה
 * בו מושמט; ההסבר בהערת הפתיחה.
 *
 * ## ושני קצות השורה **אינם** תפר — נמדד
 *
 * הדילוג חל על היסט שיש לו שני בתים, ולכאורה גם קצה של שורה שמתחילה או
 * נגמרת באי לועזי הוא כזה: מעבר לתו הראשון ולאחרון עומדת הפסקה, שכיוונה
 * הפוך. נמדד בשער (`check:arrows`, 17.9.2026) שזה **לא** המצב — המנוע מצייר
 * שם את הבית של **התו**, בדיוק מה שמחושב כאן:
 *
 *   • „ABC מילה כאן” (פסקה `w:bidi`, האי בקצה הימני ‏[1010.5 … 1043.3]):
 *     היסט 0 מצויר ב-**1010.5**, כלומר בקצה השמאלי של `A` — ולא בתחילת
 *     הפסקה שמימין. ההליכה שלמה: ‏…4@1006.4 → 0@1010.5 → 1@1022, וההפוכה
 *     ‏3@1043.3 → 2@1032.7 → 1@1022 → 0@1010.5 → 4@1006.4.
 *   • „לפני הטבלה 4” (הספרה בקצה השמאלי, ‏[967 … 975.4]): היסט הסיום 12
 *     מצויר ב-**975.4**, הקצה הימני של הספרה — ולא בקצה השמאלי של השורה.
 *     ‏12@975.4 → 10@979.4 → 9@986.9, בלי צעד לכיוון ההפוך.
 *
 * דילוג על החריצים האלה היה **מוחק** מקום נגיש שהמנוע מצייר נכון: לפי אותן
 * תיבות שנמדדו, בלי החריץ 0@1010.5 ההקשה מ-4@1006.4 הייתה נוחתת על 1@1022 —
 * כלומר מעל „A” כולו (גזירה מהמדידה, לא ריצה נוספת). לכן התו הראשון אינו
 * נבדק מול קודמו, וחריץ הסיום נפלט תמיד.
 */
export function caretSlots(chars: readonly PaintedChar[], lineRtl: boolean): CaretSlot[] {
  if (!chars.length) return [];
  const dirs = charDirections(chars, lineRtl);
  const slots: CaretSlot[] = [];

  for (let i = 0; i < chars.length; i += 1) {
    if (i > 0 && dirs[i] !== dirs[i - 1]) continue;
    const c = chars[i]!;
    slots.push({ pm: c.pm, x: dirs[i] ? c.right : c.left });
  }

  const last = chars[chars.length - 1]!;
  slots.push({
    pm: last.pm + unitsOf(last),
    x: dirs[chars.length - 1] ? last.left : last.right,
  });

  return slots;
}

/** הסמן המצויר של המנוע. */
export const CARET_SELECTOR = '.sd-v2-local-selection-caret';

/** ה-fragment של פסקה. */
const FRAGMENT_SELECTOR = '[data-source-node-id][data-pm-start]';

/** עמוד בגוף המסמך. ה-fragments של הגוף הם ילדיו הישירים (נמדד). */
const PAGE_CLASS = 'superdoc-page';

/** ה-fragment הוא המשך של פסקה שהתחילה בעמוד קודם. */
const CONTINUES_FROM_PREV = 'data-continues-from-prev';
const CONTINUES_ON_NEXT = 'data-continues-on-next';

/**
 * מה שנושא היסטים בתוך שורה. סמן המספור של פריט רשימה אינו אחד מאלה — הוא
 * מצויר בשורה ואינו נושא טווח pm, ולכן נופל מעצמו (נמדד).
 */
export const PM_CARRIER_SELECTOR = '.superdoc-text-run, .superdoc-tab';

/**
 * טווח pm מתכונה, או `NaN` כשאין תכונה.
 *
 * ‏`Number(null)` הוא **0**, ולא NaN — ולכן `Number(el.getAttribute(…))` לבדו
 * הופך אלמנט בלי טווח לאלמנט שטווחו 0..0. נמדד: טאב-הסיומת של סמן רשימה
 * (`SPAN.superdoc-tab.superdoc-marker-suffix-tab`) נושא רווח אחד ואין לו
 * טווח, וכך הוא פסל את כל השורה — כלומר החזיר את הבאג לכל פריטי הרשימה.
 */
export const attr = (el: Element, name: string): number => {
  const raw = el.getAttribute(name);
  return raw === null ? Number.NaN : Number(raw);
};

/**
 * מה שנצרך מ-`Intl.Segmenter`. הוא אינו בטיפוסי `ES2020` שב-`tsconfig`
 * (נוסף ב-`es2022.intl`), והצורה מוצהרת כאן ולא כהרחבה גלובלית — כדי שלא
 * תתנגש בהצהרה של מודול אחר.
 */
interface GraphemeSegmenter {
  segment(input: string): Iterable<{ index: number; segment: string }>;
}

/**
 * החלוקה לאשכולות גרפמה, פעם אחת למודול — ההסבר ב-`readLineChars`.
 *
 * בלי locale: כללי האשכול ב-UAX-29 אינם תלויים בשפה, והברירה נמדדה נותנת
 * את אותם גבולות בדיוק ל-`he` ולברירת המחדל על שש השורות בגשש.
 */
const GRAPHEMES: GraphemeSegmenter = new (
  Intl as unknown as {
    Segmenter: new (
      locales: undefined,
      options: { granularity: 'grapheme' },
    ) => GraphemeSegmenter;
  }
).Segmenter(undefined, { granularity: 'grapheme' });

/**
 * התיבה של כל אשכול גרפמה בשורה, עם ההיסט שלו.
 *
 * מחזירה `null` כשהטווחים המצוירים אינם מכסים את השורה ברצף — אז אין מיפוי
 * מהימן, וההקשה נמסרת למנוע.
 *
 * ## אשכול גרפמה, לא נקודת קוד ולא יחידת UTF-16
 *
 * ‏`pm` נספר ביחידות UTF-16, אבל היחידה שיש לה מקום על המסך היא **אשכול
 * הגרפמה**: „שָׁ” הוא שלוש נקודות קוד (ש, קמץ, שין-ימנית) ותיבה אחת, ואין
 * מקום סמן בין אות לניקוד שלה. הכלל של הדפדפן נמדד ב-Chrome אמיתי
 * (`node scripts/qa/rtl-caret-cluster-probe.mjs`, 17.9.2026) — הוא מחזיר
 * ל-`Range` שחותך באמצע אשכול את **תיבת האשכול המלאה**:
 *
 *   | האשכול                       | יחידות | רוחב התיבה | כל יחידה בנפרד |
 *   |------------------------------|--------|------------|----------------|
 *   | `😀` — זוג תחליף              |   2    |   22.0     | אותה תיבה      |
 *   | `שָׁ` ב-„מילה abcשָׁלום”        |   3    |    9.8     | אותה תיבה      |
 *   | `e`+U+0301 ב-„מילה éa מילה”  |   2    |    7.1     | אותה תיבה      |
 *
 * (ה-x המוחלט תלוי ברוחב החלון ולכן אינו רשום כאן; מה שקבוע הוא השוויון,
 * וזה מה שהגשש מוודא.)
 *
 * כלומר איטרציה לפי נקודת קוד אינה מספיקה: היא הייתה יוצרת ל-„שָׁ” שלושה
 * חריצים באותו x, ו**שניים מהם הם היסט בין האות לניקוד שלה**. נמדד מקצה
 * לקצה על המודול הזה עם התיבות שהגשש החזיר: ב-„מילה abcשָׁלום” שרדו החריצים
 * ‏`9` ו-`10` באותו x (הבסיס `8` נפל כתפר), וההליכה — לשני הכיוונים — נחתה
 * על `9`, כלומר בין ש לקמץ שלה. הקשה שם הייתה מכניסה תו בין האות לניקודה.
 * (‏„12שָׁלום” ו-„בְּ1948 שְנָה” יצאו שלמים במקרה בלבד: חריץ עם pm נמוך יותר
 * יושב באותו x ומנצח בשוויון של `<`/`>` ב-`visualTarget`.)
 *
 * החלוקה לאשכולות גם **מוזילה** את המסלול על הטקסט שהעורך הזה נועד לו: אותה
 * מדידה על שורה מנוקדת (89 יחידות UTF-16) — 54 קריאות `getBoundingClientRect`
 * במקום 89, חציון 0.3ms במקום 0.7ms; ועל שורה בלי ניקוד (269 יחידות) 269
 * קריאות בשני המסלולים, 1.8ms מול 1.9ms.
 */
export function readLineChars(line: Element): PaintedChar[] | null {
  const chars: PaintedChar[] = [];
  const range = document.createRange();

  for (const el of Array.from(line.querySelectorAll(PM_CARRIER_SELECTOR))) {
    const pmStart = attr(el, 'data-pm-start');
    const pmEnd = attr(el, 'data-pm-end');
    if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) continue;

    const text = el.firstChild;
    if (!text || text.nodeType !== Node.TEXT_NODE) {
      // טאב: יחידת pm אחת בלי טקסט, והתיבה היא של האלמנט עצמו.
      if (pmEnd - pmStart !== 1) return null;
      const rect = el.getBoundingClientRect();
      chars.push({ pm: pmStart, left: rect.left, right: rect.right });
      continue;
    }

    const data = (text as Text).data;
    if (data.length !== pmEnd - pmStart) return null;
    for (const { index, segment } of GRAPHEMES.segment(data)) {
      range.setStart(text, index);
      range.setEnd(text, index + segment.length);
      const rect = range.getBoundingClientRect();
      chars.push({
        pm: pmStart + index,
        left: rect.left,
        right: rect.right,
        ch: segment,
        ...(segment.length > 1 ? { units: segment.length } : {}),
      });
    }
  }

  if (!chars.length) return null;
  chars.sort((a, b) => a.pm - b.pm);
  if (chars[0]!.pm !== attr(line, 'data-pm-start')) return null;
  const last = chars[chars.length - 1]!;
  if (last.pm + unitsOf(last) !== attr(line, 'data-pm-end')) return null;
  for (let i = 1; i < chars.length; i += 1) {
    if (chars[i]!.pm !== chars[i - 1]!.pm + unitsOf(chars[i - 1]!)) return null;
  }

  return chars;
}

const flag = (el: Element, name: string): boolean => el.getAttribute(name) === 'true';

/**
 * פסקה או שורה בלי תוכן: טווח באורך אפס.
 *
 * נמדד (17.9.2026, superdoc 2.15.0): פסקה ריקה מצוירת כ-fragment עם
 * `data-pm-start` **שווה** ל-`data-pm-end` (14..14), ובתוכו
 * `DIV.superdoc-line` באותו טווח, עם `dir="rtl"` ובלי אף נושא היסט. יש לה
 * בדיוק מקום סמן אחד, וזה ההיסט הזה עצמו.
 */
export function isEmptyRange(el: Element): boolean {
  const from = attr(el, 'data-pm-start');
  return Number.isFinite(from) && from === attr(el, 'data-pm-end');
}

/**
 * החריץ היחיד של שורה ריקה, או `null` כשאינה ריקה.
 *
 * ה-x נלקח ממלבן השורה, והוא אינו נקרא במסלול שמשתמש בחריץ הזה (בקצה נקרא
 * ה-pm בלבד) — אבל חריץ בלי מיקום היה שקר בטיפוס.
 */
export function emptyLineSlot(line: Element, rtl: boolean): CaretSlot | null {
  if (!isEmptyRange(line)) return null;
  const rect = line.getBoundingClientRect();
  return { pm: attr(line, 'data-pm-start'), x: rtl ? rect.right : rect.left };
}

/**
 * ה-pm של תחילת הבלוק: ה-fragment הראשון שלו, זה שאינו המשך — או, לפסקה
 * בתוך תא שאין לה טווח משלה, השורה הראשונה שלה. `NaN` כשהבלוק אינו מצויר —
 * ואז אין מיפוי מהיסט ל-pm.
 */
export function blockStart(host: HTMLElement, blockId: string): number {
  let bare: Element | null = null;
  for (const el of Array.from(host.querySelectorAll('[data-source-node-id]'))) {
    if (el.getAttribute('data-source-node-id') !== blockId) continue;
    if (!Number.isFinite(attr(el, 'data-pm-start'))) {
      bare ??= el;
      continue;
    }
    if (flag(el, CONTINUES_FROM_PREV)) continue;
    return attr(el, 'data-pm-start');
  }
  const first = bare ? linesOf(bare)[0] : undefined;
  return first ? attr(first, 'data-pm-start') : Number.NaN;
}

/** המרחק האנכי מ-y לשורה הקרובה ביותר ב-fragment שמכילה את ה-pm. */
function lineDistance(fragment: Element, pm: number, y: number): number {
  let best = Infinity;
  for (const line of linesOf(fragment)) {
    if (pm < attr(line, 'data-pm-start') || pm > attr(line, 'data-pm-end')) continue;
    const rect = line.getBoundingClientRect();
    best = Math.min(best, Math.abs((rect.top + rect.bottom) / 2 - y));
  }
  return best;
}

export interface CaretHome {
  fragment: HTMLElement;
  caretPm: number;
}

/**
 * ה-fragment שהסמן בו, וה-pm שלו. פסקה שנחצית בין עמודים היא כמה fragments
 * עם אותו מזהה, ובגבול ביניהם ה-pm שייך לשניים — ה-y של הסמן מכריע.
 */
export function findFragment(host: HTMLElement, blockId: string, caretOffset: number, caretY: number): CaretHome | null {
  const start = blockStart(host, blockId);
  if (!Number.isFinite(start)) return null;
  const caretPm = start + caretOffset;

  let best: HTMLElement | null = null;
  let bestDistance = Infinity;
  for (const fragment of Array.from(host.querySelectorAll<HTMLElement>(FRAGMENT_SELECTOR))) {
    if (fragment.getAttribute('data-source-node-id') !== blockId) continue;
    const from = attr(fragment, 'data-pm-start');
    const to = attr(fragment, 'data-pm-end');
    if (!Number.isFinite(from) || caretPm < from || (Number.isFinite(to) && caretPm > to)) continue;
    const distance = lineDistance(fragment, caretPm, caretY);
    if (!best || distance < bestDistance) {
      best = fragment;
      bestDistance = distance;
    }
  }
  return best ? { fragment: best, caretPm } : null;
}

const isPage = (el: Element | null): boolean => !!el && el.classList.contains(PAGE_CLASS);

/**
 * שני fragments באותו מכל ציור: אותו הורה, או שניהם בגוף — עמוד אחרי עמוד.
 * כותרת עליונה ותא טבלה אינם כאלה.
 */
export function sameFlow(a: Element, b: Element): boolean {
  if (a.parentElement === b.parentElement) return true;
  return isPage(a.parentElement) && isPage(b.parentElement);
}

const hasRange = (el: Element): boolean => Number.isFinite(attr(el, 'data-pm-start'));

/** הבלוקים של הגוף שבעמוד — פסקאות וטבלאות — לפי הסדר. */
function pageFragments(page: Element): Element[] {
  return Array.from(page.children).filter(hasRange);
}

/**
 * הבלוק הבא (או הקודם) בזרימה: האח הבא שיש לו טווח, ואם אין — הראשון בעמוד
 * הסמוך. עמוד סמוך שאין בו אף בלוק מצויר אינו „אין שכן” אלא „לא ידוע”, ולכן
 * `null` — והקצה של המסמך נבדק בנפרד.
 */
export function neighbourOf(host: HTMLElement, fragment: Element, forward: boolean): Element | null {
  let el = forward ? fragment.nextElementSibling : fragment.previousElementSibling;
  while (el && !hasRange(el)) el = forward ? el.nextElementSibling : el.previousElementSibling;
  if (el) return el;

  const page = fragment.parentElement;
  if (!page || !isPage(page)) return null;
  const pages = Array.from(host.getElementsByClassName(PAGE_CLASS));
  const next = pages[pages.indexOf(page) + (forward ? 1 : -1)];
  if (!next) return null;
  const items = pageFragments(next);
  return items[forward ? 0 : items.length - 1] ?? null;
}

/**
 * האם ה-fragment הוא הקצה של המסמך בכיוון הזה: הראשון בעמוד הראשון, או
 * האחרון בעמוד האחרון — ולא רק האחרון שמצויר.
 */
export function atDocumentEdge(host: HTMLElement, fragment: Element, forward: boolean): boolean {
  const page = fragment.parentElement;
  if (!page || !isPage(page)) return false;
  const pages = Array.from(host.getElementsByClassName(PAGE_CLASS));
  if (pages[forward ? pages.length - 1 : 0] !== page) return false;
  if (flag(fragment, forward ? CONTINUES_ON_NEXT : CONTINUES_FROM_PREV)) return false;
  const siblings = pageFragments(page);
  return siblings[forward ? siblings.length - 1 : 0] === fragment;
}

/** השורות של ה-fragment — ילדיו הישירים שנושאים טווח pm. */
export function linesOf(fragment: Element): Element[] {
  return Array.from(fragment.children).filter(
    (child) =>
      Number.isFinite(attr(child, 'data-pm-start')) && Number.isFinite(attr(child, 'data-pm-end')),
  );
}

/**
 * השורה שהסמן בה.
 *
 * ההיסט לבדו אינו מספיק: גבול בין שתי שורות גולשות הוא **היסט אחד** ששייך
 * לשתיהן, ושתי השורות נגמרות ומתחילות באותו x. מה שמבדיל ביניהן הוא ה-y
 * שהסמן מצויר בו, ולכן הוא זה שמכריע כששתי מועמדות.
 */
export function lineAt(lines: readonly Element[], caretPm: number, caretY: number): number {
  const candidates: number[] = [];
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i]!;
    if (caretPm >= attr(line, 'data-pm-start') && caretPm <= attr(line, 'data-pm-end')) {
      candidates.push(i);
    }
  }
  if (candidates.length <= 1) return candidates[0] ?? -1;

  let best = candidates[0]!;
  let bestGap = Infinity;
  for (const i of candidates) {
    const rect = lines[i]!.getBoundingClientRect();
    const gap = Math.abs((rect.top + rect.bottom) / 2 - caretY);
    if (gap < bestGap) {
      bestGap = gap;
      best = i;
    }
  }
  return best;
}

/** מלבן הבחירה ש-`getSelectionAnchorRect()` מחזיר, בקואורדינטות המסך. */
export interface SelectionRect {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

/**
 * כיוון השורה שמתחת למלבן: `true` לעברית, `false` ללטינית, `null` כשאין שם
 * שורה של פסקה בגוף הזרימה — תא טבלה, למשל, שפסקאותיו בלי טווח משלהן.
 */
export function lineDirectionUnder(host: HTMLElement, rect: SelectionRect): boolean | null {
  for (const fragment of Array.from(host.querySelectorAll(FRAGMENT_SELECTOR))) {
    for (const line of linesOf(fragment)) {
      const box = line.getBoundingClientRect();
      if (box.bottom < rect.top || box.top > rect.bottom) continue;
      if (box.right < rect.left || box.left > rect.right) continue;
      return line.getAttribute('dir') === 'rtl';
    }
  }
  return null;
}

/**
 * האם שני בלוקים סמוכים הם באמת שכנים **במסמך**, ולא רק ברשימת הבורר.
 *
 * טבלה מצוירת כ-fragment בלי `data-source-node-id` (נמדד: „לפני” pm 104..114,
 * הטבלה 115..122, „אחרי” 123..133), ולכן בורר ה-fragments מדלג עליה והפסקה
 * שמעבר לה נראית שכנה — והחץ קפץ מעל הטבלה כולה. פער בטווחים פירושו שיש
 * ביניהן משהו, והמנוע — שנכנס לתא כראוי — הוא שמטפל בהקשה. המשך של אותה פסקה
 * בעמוד הבא מתחיל בדיוק היכן שהקודם נגמר, בלי הפרש.
 *
 * ופסקה ריקה תופסת מקום pm נוסף **אחריה**. נמדד: „לפני הריקה 4” 1..13, הריקה
 * 14..14, „אחרי הריקה 5” 16..28 — כלומר הפרש 1 אחרי פסקה עם תוכן, והפרש 2
 * אחרי ריקה. בלי זה הריקה נראית כמו „יש משהו ביניהן”.
 *
 * ‏`sameFlow` נבדק כאן ולא אצל הקורא: כותרת עליונה ותא טבלה אינם שכנים של
 * הסמן בשום מדיניות.
 */
export function contiguousWith(fragment: Element, sibling: Element, forward: boolean): boolean {
  if (!sameFlow(fragment, sibling)) return false;
  const siblingId = sibling.getAttribute('data-source-node-id');
  const sameBlock = siblingId !== null && siblingId === fragment.getAttribute('data-source-node-id');
  const earlier = forward ? fragment : sibling;
  const gap = sameBlock ? 0 : isEmptyRange(earlier) ? 2 : 1;
  return forward
    ? attr(sibling, 'data-pm-start') === attr(fragment, 'data-pm-end') + gap
    : attr(fragment, 'data-pm-start') === attr(sibling, 'data-pm-end') + gap;
}

/**
 * הקצוות שהטקסט של השורה **מצויר** בהם, או `null` כשאין בה טקסט.
 *
 * מלבן אלמנט השורה אינו זה: הוא נמתח על רוחב הפסקה גם כששלוש מילים בלבד
 * צוירו בה, וזו בדיוק השאלה שנשאלת כאן — האם עמודת המטרה נופלת על טקסט או
 * מעבר לו. לכן הקצוות נלקחים מנושאי ה-pm, ולא מהשורה.
 *
 * זול בכוונה: מספר קריאות `getBoundingClientRect` כמספר הריצות בשורה, ולא
 * כמספר הגרפמות שבה. מי שצריך את התיבה של כל תו קורא ל-`readLineChars`, וזה
 * מסלול יקר בהרבה — הבדיקה „האם צריך בכלל להתערב” נעשית קודם, וכאן.
 *
 * ‏**רק מה שנושא טווח pm**, בדיוק כמו ב-`readLineChars`. סמן המספור של פריט
 * רשימה מצויר בשורה ומותאם לבורר (טאב-הסיומת שלו הוא
 * `SPAN.superdoc-tab.superdoc-marker-suffix-tab`) אבל **אין לו טווח**, מפני
 * שאין בו מקום סמן. ספירתו כאן הייתה מותחת את הקצה אל אזור המספור, ואז עמודה
 * שנופלת שם הייתה נספרת כ„בתוך השורה” — כלומר מקום שאי אפשר להעמיד בו סמן
 * היה נחשב יעד תקין.
 */
export function lineExtent(line: Element): { left: number; right: number } | null {
  let left = Infinity;
  let right = -Infinity;
  for (const el of Array.from(line.querySelectorAll(PM_CARRIER_SELECTOR))) {
    if (!Number.isFinite(attr(el, 'data-pm-start'))) continue;
    const rect = el.getBoundingClientRect();
    if (!rect.width) continue;
    left = Math.min(left, rect.left);
    right = Math.max(right, rect.right);
  }
  return right < left ? null : { left, right };
}

/**
 * האם ההיסט הראשון של השורה הוא **תפר גלישה** — כלומר גם ההיסט האחרון של
 * השורה שמעליה.
 *
 * זה חשוב מפני שלתפר כזה יש בעלים אחד בלבד. נמדד (22.9.2026, superdoc 2.15.0,
 * ‏`scripts/qa/rtl-vertical-seam-probe.mjs`): פסקה שגלשה לשתי שורות, pm 1..94
 * ו-94..153; כתיבת ההיסט שבתפר דרך `authoring.setSelectionTarget` ציירה את
 * הסמן ב-y של השורה ה**ראשונה** ובקצה השמאלי שלה — כלומר בסופה, ולא בתחילת
 * השנייה. אותו כלל שכבר נמדד למקש `End` ב-`rtl-line-end.ts`.
 *
 * המסקנה למי שכותב בחירה: „לך לתחילת שורת המשך” **אינו ניתן לביטוי**, ומי
 * שמכוון לשם מקבל את סוף השורה שמעליה — כלומר הקשה שנראית כאילו לא עשתה דבר.
 * המקום היחיד שבו זה עולה הוא שורה מוזחת או ממורכזת שקצה ההתחלה שלה נסוג
 * פנימה; בפסקה רגילה עמודת המטרה אינה יכולה ליפול מעבר לקצה ההתחלה, שהוא אותו
 * x בכל השורות.
 */
export function startIsSeam(fragment: Element, lines: readonly Element[], index: number): boolean {
  if (index > 0) {
    const previous = lines[index - 1];
    return !!previous && attr(previous, 'data-pm-end') === attr(lines[index]!, 'data-pm-start');
  }
  return flag(fragment, CONTINUES_FROM_PREV);
}

/**
 * מזהה הבלוק שה-fragment מצייר, או `null` כשאין לו אחד.
 *
 * ‏`null` אינו „חסר” אלא **טבלה**: היא מצוירת כ-fragment בלי מזהה מקור, והבלוק
 * האמיתי הוא הפסקה שבתוך התא. מי שמקבל `null` צריך להחליט מה לעשות עם זה —
 * `rtl-caret.ts` יורד לתא, `rtl-vertical-caret.ts` מוסר למנוע — וזה הבדל
 * במדיניות, לא בקריאה.
 */
export function blockIdOf(fragment: Element): string | null {
  return fragment.getAttribute('data-source-node-id');
}

/** הבלוקים המצוירים שבתוך fragment — הפסקאות שבתאי טבלה, לפי הסדר. */
export function blocksInside(fragment: Element): Element[] {
  return Array.from(fragment.querySelectorAll('[data-source-node-id]')).filter(
    (el) => linesOf(el).length > 0,
  );
}

/**
 * מזהה הבלוק הראשון של גוף המסמך, כפי שהוא מצויר — או `null`.
 *
 * הצרכן: סמן הפתיחה (`applyDocumentStartCaret` ב-caret-anchor.ts). קריאות
 * המסמך של המנוע ממתינות לשקט, ובמסמך ארוך המנוע ממשיך לעמד ברקע עוד זמן רב
 * אחרי `onReady`, ואילו העמוד הראשון כבר מצויר בו.
 *
 * רק גוף המסמך (`data-layout-story="body"`) — כותרת עליונה ותחתונה מצוירות
 * באותו עמוד. והמינימום לפי `data-pm-start` ולא הראשון ב-DOM: המקטעים ממוקמים
 * באופן מוחלט, וסדר הצאצאים אינו מובטח להיות סדר המסמך. גם טבלה היא fragment,
 * אך בלי מזהה מקור: אם היא הבלוק הראשון, הקורא נופל למסלול של המנוע ולא
 * מדלג לפסקה שאחריה. פסקאות שבתאים אינן fragments של הגוף. כשחסר טווח
 * ל-fragment אי אפשר להוכיח מי הראשון, ולכן חוזרים למנוע.
 */
export function paintedFirstBlockId(host: ParentNode | null | undefined): string | null {
  if (typeof host?.querySelectorAll !== 'function') return null;
  let firstId: string | null = null;
  let firstStart = Infinity;
  for (const fragment of Array.from(host.querySelectorAll('.superdoc-fragment[data-layout-story="body"]'))) {
    const start = attr(fragment, 'data-pm-start');
    if (!Number.isFinite(start)) return null;
    const id = fragment.getAttribute('data-source-node-id');
    if (start > firstStart) continue;
    // טווח ראשון משותף בלי מזהה חד-משמעי אינו סיבה לבחור לפי סדר ה-DOM.
    if (start === firstStart) {
      if (id !== firstId) firstId = null;
      continue;
    }
    firstStart = start;
    firstId = id || null;
  }
  return firstId;
}

/* ------------------------------------------------------------------ */
/* מיקום טקסט שהוחלף — לסימון התיקון האוטומטי                           */
/* ------------------------------------------------------------------ */

/**
 * הבלוק המצויר שהצומת נמצא בו (הפסקה הקרובה ביותר, גם בתוך תא טבלה), או `null`.
 * מי שרוצה לבדוק רק מה שנוגע לבלוק שמתחת לעכבר, בלי לסרוק את המסמך כולו.
 */
export function blockIdAt(node: Node | null): string | null {
  const el = node instanceof Element ? node : (node?.parentElement ?? null);
  const fragment = el?.closest('[data-source-node-id]') ?? null;
  return fragment ? fragment.getAttribute('data-source-node-id') : null;
}

export interface Box {
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
}

export interface LocatedWord {
  /** האות הראשונה של הטקסט — מתחתיה יושב הפס. */
  readonly first: Box;
  /** כל מלבני המילה (יותר מאחד כשהיא נשברת בין שורות). */
  readonly rects: readonly Box[];
  /** ההיסט בבלוק שבו הטקסט נמצא עכשיו. שונה מהמבוקש כשעריכה לפניו הזיזה אותו. */
  readonly start: number;
}

/** מה שהיה סביב הטקסט ברגע ההחלפה, כדי למצוא אותו שוב אחרי עריכה לפניו. */
export interface TextContext {
  readonly before: string;
  readonly after: string;
}

function rectsOf(node: Node, offset: number, length: number): { first: Box; rects: Box[] } | null {
  const range = document.createRange();
  range.setStart(node, offset);
  range.setEnd(node, offset + length);
  const rects = Array.from(range.getClientRects()).filter((r) => r.width > 0 && r.height > 0);
  range.setEnd(node, offset + 1);
  const first = Array.from(range.getClientRects()).find((r) => r.width > 0 && r.height > 0);
  return first && rects.length > 0 ? { first, rects } : null;
}

/**
 * היכן טקסט מסוים מצויר עכשיו, או `null`: הבלוק, ההיסט בו, והטקסט שאמור להיות שם.
 *
 * המנוע מצייר כל פסקה כ-fragment עם `data-source-node-id` ו-`data-pm-start`, וכל
 * רצף טקסט בה כאלמנט עם `data-pm-start`/`data-pm-end` (נמדד; אותו מבנה ש-
 * `rtl-line-end.ts` קורא). ההמרה להיסט בבלוק היא חיסור אחד: `offset = pm −
 * fragmentPmStart`.
 *
 * בלי הקשר נבדק הטקסט בהיסט הרשום. עם הקשר מחפשים בכל הפסקה, ומקבלים רק
 * מופע יחיד שמה שלפניו ומה שאחריו תואם ל-`context`. יותר ממופע מתאים אחד
 * אינו מאפשר לזהות את המילה, גם בהיסט הרשום, ולכן מוחזר `null`.
 * התאמה זו מתארת את הטקסט הנוכחי; היא אינה עוגן זהות בין עריכות במסמך.
 */
export function locateText(
  blockId: string,
  start: number,
  expected: string,
  root: ParentNode = document,
  context?: TextContext,
): LocatedWord | null {
  return createTextLocator(blockId, root)(start, expected, context);
}

/** נתוני בלוק נאספים פעם אחת לסריקת ריחוף; אין לשמור את הקורא בין פריימים. */
export function createTextLocator(
  blockId: string,
  root: ParentNode = document,
): (start: number, expected: string, context?: TextContext) => LocatedWord | null {
  const readers: Array<(start: number, expected: string, context?: TextContext) => LocatedWord | null | 'ambiguous'> = [];
  const fragments = root.querySelectorAll<HTMLElement>('[data-source-node-id][data-pm-start]');
  for (const fragment of Array.from(fragments)) {
    if (fragment.getAttribute('data-source-node-id') !== blockId) continue;
    const base = attr(fragment, 'data-pm-start');
    const end = attr(fragment, 'data-pm-end');
    if (!Number.isFinite(base)) continue;

    // הרצפים של הפסקה לפי הסדר, ומהם הטקסט שלה. רצף שאורכו אינו טווח ה-pm שלו
    // (תו שהמנוע מייצג אחרת) מפסיק את הניחוש: אין מיפוי מהימן.
    const runs: { node: Node; pmStart: number; text: string }[] = [];
    let reliable = true;
    for (const run of Array.from(fragment.querySelectorAll<HTMLElement>('[data-pm-start]'))) {
      if (run.children.length > 0) continue;
      const runStart = attr(run, 'data-pm-start');
      const runEnd = attr(run, 'data-pm-end');
      const node = run.firstChild;
      if (!node || node.nodeType !== Node.TEXT_NODE || !Number.isFinite(runStart)) continue;
      const text = node.nodeValue ?? '';
      if (runEnd - runStart !== text.length) reliable = false;
      runs.push({ node, pmStart: runStart, text });
    }
    if (runs.length === 0) continue;

    // הטקסט של הפסקה כולה, מה-pm הראשון שלה — לבדיקת ההקשר ולחיפוש.
    const first = runs[0]!.pmStart;
    let paragraph = '';
    if (reliable) {
      for (const run of runs) {
        const gap = run.pmStart - first - paragraph.length;
        if (gap < 0) {
          reliable = false;
          break;
        }
        if (gap > 0) paragraph += '\u0000'.repeat(gap);
        paragraph += run.text;
      }
    }
    // מחרוזת והקשר זהים נבדקים פעם אחת בפריים, גם כשמאות תיקונים זהים נשמרו.
    // -1 = אין מופע; null = יותר ממופע אחד, ולכן אין זיהוי בטוח.
    const contextMatches = new Map<string, number | null>();

    readers.push((start, expected, context) => {
      if (expected.length === 0) return null;
      const at = (blockOffset: number): LocatedWord | null => {
        let pm = base + blockOffset;
        if (Number.isFinite(end) && (pm < base || pm >= end)) return null;
        const segments: { node: Node; offset: number; length: number }[] = [];
        let matched = 0;
        for (const run of runs) {
          if (matched === expected.length) break;
          const offset = pm - run.pmStart;
          if (offset < 0) return null; // רווח PM או סוף fragment אינם המשך טקסט.
          if (offset >= run.text.length) continue;
          const length = Math.min(expected.length - matched, run.text.length - offset);
          if (run.text.slice(offset, offset + length) !== expected.slice(matched, matched + length)) return null;
          segments.push({ node: run.node, offset, length });
          matched += length;
          pm += length;
        }
        if (matched !== expected.length) return null;
        let firstBox: Box | null = null;
        const rects: Box[] = [];
        for (const segment of segments) {
          const found = rectsOf(segment.node, segment.offset, segment.length);
          if (!found) return null;
          firstBox ??= found.first;
          rects.push(...found.rects);
        }
        return firstBox ? { first: firstBox, rects, start: blockOffset } : null;
      };
      if (!context) return at(start);
      if (!reliable) return null;
      const key = JSON.stringify([expected, context.before, context.after]);
      let match = contextMatches.get(key);
      if (match === undefined) {
        match = -1;
        for (let from = paragraph.indexOf(expected); from >= 0; from = paragraph.indexOf(expected, from + 1)) {
          if (!paragraph.slice(0, from).endsWith(context.before) ||
            !paragraph.slice(from + expected.length).startsWith(context.after)) continue;
          if (match >= 0) { match = null; break; }
          match = from;
        }
        contextMatches.set(key, match);
      }
      if (match === null) return 'ambiguous';
      return match < 0 ? null : at(first - base + match);
    });
  }
  return (start, expected, context) => {
    let located: LocatedWord | null = null;
    for (const read of readers) {
      const found = read(start, expected, context);
      if (found === 'ambiguous' || (found && located)) return null;
      if (found) located = found;
    }
    return located;
  };
}
