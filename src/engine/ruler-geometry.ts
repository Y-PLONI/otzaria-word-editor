/**
 * החשבון של הסרגל: שנתות, יחידות, הצמדה וחסמים. בלי DOM, בלי מנוע, בלי Vue.
 *
 * ## למה מודול נפרד ולא חשבון בתוך הקומפוננטה
 *
 * הסרגל הוא פקד שמתרגם פיקסלים למידות של מסמך ובחזרה, וכל טעות בתרגום הזה
 * נראית *נכון* על המסך: שנתה שזזה בחצי מילימטר, מספר שמתחיל בצד הלא נכון, או
 * גרירה שמייצרת שוליים שליליים. שום בדיקת קומפוננטה לא תופסת את זה — צריך
 * לחשב את המספרים ולהשוות אותם למה שהמסמך אמור לקבל. לכן החשבון כאן, נבדק
 * ב-tests/unit/ruler-geometry.test.ts, והקומפוננטה רק מציירת את מה שיוצא.
 *
 * ## היחידה הפנימית היא twips
 *
 * זו היחידה של OOXML (1440 לאינץ'), וזו גם היחידה של `format.paragraph.*`
 * (נמדד ב-engine/paragraph-format.ts: הערך שנשלח הוא מה שנכתב). `sections.*`
 * מקבל אינצ'ים ולכן ההמרה נעשית ברגע הקריאה למנוע, בדיוק כמו ב-page-setup.ts.
 * הצגה לעין המשתמש היא בסנטימטרים (או באינצ'ים, לפי `measurementUnit` של
 * המנוע) — ההמרה היחידה שמותרת בשכבת התצוגה.
 *
 * ## מערכת הצירים: „מתחילת העמוד”, לוגית
 *
 * כל מיקום כאן הוא מרחק **מקצה ההתחלה של העמוד** — הקצה הימני במסמך עברי
 * והשמאלי במסמך לועזי. כך אותו חשבון משרת את שני הכיוונים, וההיפוך היחיד הוא
 * במיפוי לפיקסלים (`pixelOffset`). הניסיון ההפוך — לחשב בפיקסלים פיזיים
 * ולהפוך את הסימנים בכל מקום — הוא בדיוק איך שנולדים סרגלים שבהם המספרים
 * גדלים לכיוון אחד והידיות זזות לכיוון השני.
 *
 * ## הנקודה שממנה סופרים היא תחילת הטקסט, לא קצה הדף
 *
 * ב-Word המספור מתחיל ב-0 בתחילת אזור הטקסט, גדל לכיוון הסוף, וגם **בתוך
 * השוליים** הוא גדל — כלומר 1, 2, 3 גם אחורה. אין מספרים שליליים בסרגל.
 * `rulerTicks` מייצר בדיוק את זה: הצעדים נמדדים מתחילת הטקסט לשני הכיוונים,
 * והתווית היא הערך המוחלט.
 */

/** 1440 twips לאינץ'. אותו קבוע שמופיע ב-page-setup.ts, ומאותה מדידה. */
export const TWIPS_PER_INCH = 1440;

/** אינץ' הוא 2.54 ס"מ בדיוק. */
export const CM_PER_INCH = 2.54;

/** נגזר, ולא מספר כתוב: 566.9291… ועיגול שלו היה מזיז שנתות לאורך הדף. */
export const TWIPS_PER_CM = TWIPS_PER_INCH / CM_PER_INCH;

/**
 * יחידת המידה של הסרגל. זהה ל-`SuperDocMeasurementUnit` של המנוע — הסרגל
 * הולך אחרי `measurementUnit` ולא אחרי הנחה שלנו, וזו אותה הכרעה כמו
 * ב-engine/hf-chrome.ts על יחידת המידה בפאנל הכותרות.
 */
export type RulerUnit = 'cm' | 'in';

/** שם היחידה כפי שהוא מוצג. הנוסח זהה ל-HF_UNIT_TEXT — קול אחד לכל הממשק. */
export const UNIT_TEXT: Record<RulerUnit, string> = {
  cm: 'ס"מ',
  in: "אינץ'",
};

export function twipsPerUnit(unit: RulerUnit): number {
  return unit === 'in' ? TWIPS_PER_INCH : TWIPS_PER_CM;
}

/**
 * הצעדים של השנתות, ביחידה.
 *
 * הסולם העשרוני לסנטימטר והחצאי־שמיניות לאינץ' הם מה ש-Word מצייר, וזו גם
 * הסיבה שאין כאן צעד אחד „כללי”: רבע סנטימטר הוא מרחק סביר לעין, ורבע אינץ'
 * הוא כבר גס מדי.
 */
export const TICK_STEPS: Record<RulerUnit, { major: number; mid: number; minor: number }> = {
  cm: { major: 1, mid: 0.5, minor: 0.25 },
  in: { major: 1, mid: 0.5, minor: 0.125 },
};

/** צעד ההצמדה בגרירה. עם Alt הגרירה חופשית — כמו ב-Word. */
export function snapStepUnits(unit: RulerUnit): number {
  return TICK_STEPS[unit].minor;
}

export type TickKind = 'major' | 'mid' | 'minor';

export interface RulerTick {
  /** מרחק מקצה ההתחלה של העמוד, ב-twips. */
  twips: number;
  kind: TickKind;
  /** מספר שמוצג ליד השנתה. קיים רק בשנתות ראשיות שאינן 0. */
  label?: string;
}

export interface TickInput {
  /** רוחב העמוד ב-twips. */
  pageWidthTwips: number;
  /** מרחק תחילת הטקסט מקצה ההתחלה — כלומר שולי ההתחלה. */
  textStartTwips: number;
  unit: RulerUnit;
  /** פיקסלים ליחידת תצוגה על המסך. קובע אילו שנתות בכלל נראות. */
  pxPerUnit: number;
}

/**
 * המרווח המינימלי בין שנתות שעוד אפשר לראות. מתחת לזה השנתות מתמזגות לקו
 * אפור אחיד, וזה גרוע משנתות שאינן שם: הסרגל נראה מלוכלך ולא מדויק.
 */
const MIN_TICK_GAP_PX = 4;

/** המרווח המינימלי בין מספרים. מתחת לזה הם נדבקים זה לזה. */
const MIN_LABEL_GAP_PX = 26;

/** כל כמה שנתות ראשיות מוצג מספר, כשאין מקום לכולם. */
const LABEL_INTERVALS = [1, 2, 5, 10] as const;

/** בוחרת כל כמה יחידות יוצג מספר, לפי המקום שיש בפועל. */
export function labelInterval(pxPerUnit: number): number {
  for (const interval of LABEL_INTERVALS) {
    if (pxPerUnit * interval >= MIN_LABEL_GAP_PX) return interval;
  }
  return LABEL_INTERVALS[LABEL_INTERVALS.length - 1];
}

/**
 * השנתות של הסרגל, מקצה העמוד ועד קצהו.
 *
 * הצעדים נמדדים **מתחילת הטקסט** לשני הכיוונים (ראו הערת הפתיחה), ולכן
 * ההתחלה של הלולאה היא האינדקס השלם הראשון שעדיין בתוך הדף ולא 0.
 *
 * שנתה שנופלת בדיוק על גבול העמוד נכללת: היא הקצה, ובלעדיה הסרגל נראה
 * קטוע חצי שנתה לפני הסוף.
 */
export function rulerTicks(input: TickInput): RulerTick[] {
  const { pageWidthTwips, textStartTwips, unit, pxPerUnit } = input;
  if (!(pageWidthTwips > 0) || !Number.isFinite(textStartTwips) || !(pxPerUnit > 0)) return [];

  const steps = TICK_STEPS[unit];
  const perUnit = twipsPerUnit(unit);
  const step = steps.minor;
  const stepTwips = step * perUnit;

  // סיבולת של רבע twip: `textStart` מגיע מהמרה של אינצ'ים ואינו שלם, ובלעדיה
  // שנתה שאמורה ליפול בדיוק על 1.0 נופלת על 0.9999 ומאבדת את התווית שלה.
  const epsilon = 0.25;
  const showMinor = pxPerUnit * steps.minor >= MIN_TICK_GAP_PX;
  const showMid = pxPerUnit * steps.mid >= MIN_TICK_GAP_PX;
  const every = labelInterval(pxPerUnit);

  const first = Math.ceil((-textStartTwips - epsilon) / stepTwips);
  const last = Math.floor((pageWidthTwips - textStartTwips + epsilon) / stepTwips);

  const ticks: RulerTick[] = [];
  for (let index = first; index <= last; index += 1) {
    const units = index * step;
    const twips = textStartTwips + units * perUnit;
    if (twips < -epsilon || twips > pageWidthTwips + epsilon) continue;

    const kind = tickKind(units, steps);
    if (kind === 'minor' && !showMinor) continue;
    if (kind === 'mid' && !showMid) continue;

    const labelled = kind === 'major' && units !== 0 && Math.abs(units) % every === 0;
    ticks.push({
      twips,
      kind,
      ...(labelled ? { label: String(Math.abs(Math.round(units))) } : {}),
    });
  }

  return ticks;
}

/** דירוג השנתה. ההשוואה היא על מספר הצעדים ולא על השבר, כדי לא לצבור שגיאה. */
function tickKind(units: number, steps: { major: number; mid: number; minor: number }): TickKind {
  const stepsFromZero = Math.round(units / steps.minor);
  const perMajor = Math.round(steps.major / steps.minor);
  const perMid = Math.round(steps.mid / steps.minor);
  if (stepsFromZero % perMajor === 0) return 'major';
  if (stepsFromZero % perMid === 0) return 'mid';
  return 'minor';
}

/**
 * מיקום על המסך של ערך „מתחילת העמוד”.
 *
 * זה המקום **היחיד** שבו הכיוון נכנס לחשבון. `pageLeftPx` ו-`pageWidthPx` הם
 * המלבן המצויר של העמוד, ולכן במסמך עברי ההתחלה היא הקצה הימני שלו.
 */
export function pixelOffset(
  twipsFromStart: number,
  page: { leftPx: number; widthPx: number; widthTwips: number },
  direction: 'rtl' | 'ltr',
): number {
  if (!(page.widthTwips > 0)) return page.leftPx;
  const ratio = twipsFromStart / page.widthTwips;
  const distance = ratio * page.widthPx;
  return direction === 'rtl' ? page.leftPx + page.widthPx - distance : page.leftPx + distance;
}

/** ההמרה ההפוכה: פיקסל על המסך → מרחק מתחילת העמוד ב-twips. */
export function twipsFromPixel(
  pixel: number,
  page: { leftPx: number; widthPx: number; widthTwips: number },
  direction: 'rtl' | 'ltr',
): number {
  if (!(page.widthPx > 0)) return 0;
  const distance = direction === 'rtl' ? page.leftPx + page.widthPx - pixel : pixel - page.leftPx;
  return (distance / page.widthPx) * page.widthTwips;
}

/**
 * הצמדה לשנתה. `free` (המקש Alt מוחזק) מדלג עליה ומחזיר עיגול ל-twip שלם —
 * המנוע דוחה כניסה שאינה מספר שלם, וזה נמדד: „must be a non-negative integer”.
 */
export function snapTwips(twips: number, unit: RulerUnit, free = false): number {
  if (!Number.isFinite(twips)) return 0;
  if (free) return Math.round(twips);
  const stepTwips = snapStepUnits(unit) * twipsPerUnit(unit);
  return Math.round(Math.round(twips / stepTwips) * stepTwips);
}

/**
 * הרוחב המינימלי שנשאר לטקסט. שוליים או כניסות שמצטמצמים אל מתחת לזה יוצרים
 * עמוד בלי עמודת טקסט — Word אינו מרשה את זה, וגם המנוע היה מקבל את הערך
 * ומצייר מסמך שאי אפשר לכתוב בו.
 */
export const MIN_TEXT_WIDTH_TWIPS = Math.round(TWIPS_PER_CM); // ס"מ אחד

export interface MarginBounds {
  pageWidthTwips: number;
  /** השוליים שבצד השני, שאינם זזים בגרירה הזאת. */
  otherMarginTwips: number;
}

/** גבולות הגרירה של ידית שוליים: לא שלילי, ולא על חשבון עמודת הטקסט. */
export function clampMargin(twips: number, bounds: MarginBounds): number {
  const max = bounds.pageWidthTwips - bounds.otherMarginTwips - MIN_TEXT_WIDTH_TWIPS;
  return Math.round(Math.min(Math.max(0, twips), Math.max(0, max)));
}

export interface IndentBounds {
  /** רוחב עמודת הטקסט — כלומר העמוד פחות שני השוליים. */
  columnWidthTwips: number;
  /** הכניסה שבצד השני. */
  otherIndentTwips: number;
}

/**
 * גבולות הגרירה של סמן כניסה.
 *
 * כניסה שלילית **אינה** מוצעת, ובכוונה: המנוע מקבל `left: -500` בשקט (נמדד,
 * ראו paragraph-format.ts), ודיאלוג הפסקה של Word אינו מציע אותה. סרגל שמאפשר
 * לגרור לשלילי היה יוצר מסמך שנראה תקין אצלנו ושונה ב-Word.
 */
export function clampIndent(twips: number, bounds: IndentBounds): number {
  const max = bounds.columnWidthTwips - bounds.otherIndentTwips - MIN_TEXT_WIDTH_TWIPS;
  return Math.round(Math.min(Math.max(0, twips), Math.max(0, max)));
}

/**
 * מספר לתצוגה: „2.5” בסנטימטרים, „0.98” באינצ'ים.
 *
 * שתי ספרות אחרי הנקודה ולא יותר — זו הדיוק ש-Word מציג, ומעבר לזה המספר
 * מתחיל לספר על שגיאות עיגול של ההמרה מ-twips ולא על המסמך.
 */
export function formatUnits(twips: number, unit: RulerUnit): string {
  const value = twips / twipsPerUnit(unit);
  const rounded = Math.round(value * 100) / 100;
  // בלי אפסים מיותרים: „2” ולא „2.00”.
  return String(rounded);
}

/** תווית מלאה עם יחידה, לקוראי מסך ולתווית הגרירה. */
export function measureLabel(twips: number, unit: RulerUnit): string {
  return `${formatUnits(twips, unit)} ${UNIT_TEXT[unit]}`;
}
