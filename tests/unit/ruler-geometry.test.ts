/**
 * החשבון של הסרגל.
 *
 * זו הבדיקה שמחליפה „הסתכלתי והוא נראה בסדר”. סרגל שגוי אינו נראה שגוי: שנתה
 * שזזה בחצי מילימטר, מספר שמתחיל בצד הלא נכון או גרירה שמייצרת שוליים
 * שליליים — כולם נראים כמו סרגל. לכן המספרים נבדקים כאן מול מידות שאפשר
 * לחשב ביד, ובראשן A4 עם שוליים של 2.54 ס"מ, שהוא בדיוק המסמך שהתוסף פותח.
 */
import { describe, it, expect } from 'vitest';
import {
  CM_PER_INCH,
  MIN_TEXT_WIDTH_TWIPS,
  TWIPS_PER_CM,
  TWIPS_PER_INCH,
  UNIT_TEXT,
  clampIndent,
  clampMargin,
  formatUnits,
  labelInterval,
  measureLabel,
  pixelOffset,
  rulerTicks,
  snapStepUnits,
  snapTwips,
  twipsFromPixel,
  twipsPerUnit,
} from '../../src/engine/ruler-geometry';

/** A4: 21 ס"מ. אלה המספרים שהמנוע החזיר בפועל על מסמך חדש. */
const A4_WIDTH_TWIPS = Math.round(21 * TWIPS_PER_CM); // 11 906
const INCH_MARGIN = TWIPS_PER_INCH; // 2.54 ס"מ — ברירת המחדל של המסמך הריק

describe('יחידות', () => {
  it('twips ליחידה', () => {
    expect(twipsPerUnit('in')).toBe(1440);
    expect(twipsPerUnit('cm')).toBeCloseTo(1440 / 2.54, 6);
    expect(TWIPS_PER_CM * CM_PER_INCH).toBeCloseTo(TWIPS_PER_INCH, 6);
  });

  it('הנוסח זהה לזה של פאנל הכותרות', () => {
    expect(UNIT_TEXT.cm).toBe('ס"מ');
    expect(UNIT_TEXT.in).toBe("אינץ'");
  });

  it('מספר לתצוגה בלי אפסים מיותרים', () => {
    expect(formatUnits(TWIPS_PER_CM * 2, 'cm')).toBe('2');
    expect(formatUnits(TWIPS_PER_CM * 2.5, 'cm')).toBe('2.5');
    expect(formatUnits(TWIPS_PER_INCH, 'cm')).toBe('2.54');
    expect(measureLabel(TWIPS_PER_INCH, 'in')).toBe("1 אינץ'");
  });
});

describe('שנתות', () => {
  const ticks = rulerTicks({
    pageWidthTwips: A4_WIDTH_TWIPS,
    textStartTwips: INCH_MARGIN,
    unit: 'cm',
    pxPerUnit: 37.8, // ס"מ ב-100%: 96/2.54
  });

  it('האפס יושב בתחילת הטקסט, ואין לו מספר', () => {
    const zero = ticks.find((tick) => Math.abs(tick.twips - INCH_MARGIN) < 1);
    expect(zero).toBeDefined();
    expect(zero?.kind).toBe('major');
    expect(zero?.label).toBeUndefined();
  });

  it('המספרים גדלים לשני הכיוונים, ואין ביניהם שלילי', () => {
    const labels = ticks.filter((tick) => tick.label).map((tick) => tick.label);
    expect(labels.some((label) => label?.startsWith('-'))).toBe(false);
    // 2.54 ס"מ שוליים → מספר „1” ו„2” גם בתוך השוליים, ו-„1”…„15” באזור הטקסט.
    expect(labels.filter((label) => label === '1')).toHaveLength(2);
    expect(labels.filter((label) => label === '2')).toHaveLength(2);
  });

  it('אף שנתה אינה חורגת מהדף', () => {
    for (const tick of ticks) {
      expect(tick.twips).toBeGreaterThanOrEqual(-0.5);
      expect(tick.twips).toBeLessThanOrEqual(A4_WIDTH_TWIPS + 0.5);
    }
  });

  it('שנתה ראשית בכל סנטימטר שלם', () => {
    const majors = ticks.filter((tick) => tick.kind === 'major');
    // 21 ס"מ → 22 גבולות שלמים, פחות אלה שנופלים מחוץ לדף בגלל היסט השוליים.
    expect(majors.length).toBe(21);
    const gaps = majors
      .map((tick) => tick.twips)
      .sort((a, b) => a - b)
      .slice(1)
      .map((twips, index) => twips - majors.map((t) => t.twips).sort((a, b) => a - b)[index]);
    for (const gap of gaps) expect(gap).toBeCloseTo(TWIPS_PER_CM, 3);
  });

  it('בזום נמוך השנתות הצפופות נעלמות ולא מתמזגות לקו אפור', () => {
    const dense = rulerTicks({
      pageWidthTwips: A4_WIDTH_TWIPS,
      textStartTwips: INCH_MARGIN,
      unit: 'cm',
      pxPerUnit: 9, // ‎25% זום בערך
    });
    expect(dense.some((tick) => tick.kind === 'minor')).toBe(false);
    expect(dense.some((tick) => tick.kind === 'major')).toBe(true);
  });

  it('מספר בכל שנתה כשיש מקום, ואחת לכמה כשאין', () => {
    expect(labelInterval(37.8)).toBe(1);
    expect(labelInterval(9)).toBe(5);
    expect(labelInterval(0.5)).toBe(10);
  });

  it('באינץ׳ הסולם הוא שמיניות', () => {
    const inches = rulerTicks({
      pageWidthTwips: TWIPS_PER_INCH * 4,
      textStartTwips: 0,
      unit: 'in',
      pxPerUnit: 96,
    });
    const minors = inches.filter((tick) => tick.kind === 'minor');
    expect(minors.length).toBeGreaterThan(0);
    // שמינית אינץ' = 180 twips.
    expect(minors[0].twips % 180).toBeCloseTo(0, 6);
  });

  it('קלט חסר מחזיר רשימה ריקה ולא זורק', () => {
    expect(rulerTicks({ pageWidthTwips: 0, textStartTwips: 0, unit: 'cm', pxPerUnit: 10 })).toEqual([]);
    expect(
      rulerTicks({ pageWidthTwips: 1000, textStartTwips: 0, unit: 'cm', pxPerUnit: 0 }),
    ).toEqual([]);
  });
});

describe('מיפוי לפיקסלים', () => {
  const page = { leftPx: 100, widthPx: 800, widthTwips: A4_WIDTH_TWIPS };

  it('במסמך לועזי ההתחלה היא הקצה השמאלי', () => {
    expect(pixelOffset(0, page, 'ltr')).toBe(100);
    expect(pixelOffset(A4_WIDTH_TWIPS, page, 'ltr')).toBe(900);
  });

  it('במסמך עברי ההתחלה היא הקצה הימני', () => {
    expect(pixelOffset(0, page, 'rtl')).toBe(900);
    expect(pixelOffset(A4_WIDTH_TWIPS, page, 'rtl')).toBe(100);
  });

  it('ההמרה ההפוכה מחזירה את אותו ערך, בשני הכיוונים', () => {
    for (const direction of ['rtl', 'ltr'] as const) {
      const px = pixelOffset(3000, page, direction);
      expect(twipsFromPixel(px, page, direction)).toBeCloseTo(3000, 6);
    }
  });
});

describe('הצמדה', () => {
  it('מצמידה לרבע סנטימטר', () => {
    const quarter = TWIPS_PER_CM / 4;
    expect(snapTwips(quarter * 3 + 20, 'cm')).toBe(Math.round(quarter * 3));
    expect(snapStepUnits('cm')).toBe(0.25);
  });

  it('Alt מבטל את ההצמדה אבל לא את השלמות', () => {
    // המנוע דוחה כניסה שאינה מספר שלם — „must be a non-negative integer”.
    expect(snapTwips(517.4, 'cm', true)).toBe(517);
    expect(Number.isInteger(snapTwips(517.4, 'cm', false))).toBe(true);
  });
});

describe('חסמים', () => {
  it('שוליים אינם שליליים', () => {
    expect(clampMargin(-500, { pageWidthTwips: A4_WIDTH_TWIPS, otherMarginTwips: 1440 })).toBe(0);
  });

  it('שוליים אינם בולעים את עמודת הטקסט', () => {
    const max = clampMargin(A4_WIDTH_TWIPS, {
      pageWidthTwips: A4_WIDTH_TWIPS,
      otherMarginTwips: 1440,
    });
    expect(max).toBe(A4_WIDTH_TWIPS - 1440 - MIN_TEXT_WIDTH_TWIPS);
    expect(A4_WIDTH_TWIPS - 1440 - max).toBeGreaterThanOrEqual(MIN_TEXT_WIDTH_TWIPS);
  });

  it('כניסה אינה שלילית ואינה בולעת את העמודה', () => {
    const column = 6000;
    expect(clampIndent(-10, { columnWidthTwips: column, otherIndentTwips: 0 })).toBe(0);
    expect(clampIndent(99999, { columnWidthTwips: column, otherIndentTwips: 1000 })).toBe(
      column - 1000 - MIN_TEXT_WIDTH_TWIPS,
    );
  });

  it('עמודה צרה מהמינימום אינה מייצרת חסם שלילי', () => {
    expect(clampIndent(500, { columnWidthTwips: 200, otherIndentTwips: 0 })).toBe(0);
  });
});
