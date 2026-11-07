/**
 * ההחלטה „הטקסט שהוקלד הוא סמן רשימה”.
 *
 * כל מקרה כאן הוא **שורה שנמדדה ב-Word חי** — Office 16, ממשק עברי, הקלדה
 * דרך `PostMessage(WM_CHAR)` לחלון העריכה שלו וקריאה בחזרה דרך COM. הטבלאות
 * המלאות ב-docs/list-autoformat-research.md, והמקרים כאן מכסים אותן במלואן:
 * גם מה שהומר, גם מה שנדחה, וגם מה ש-Word כתב בפועל ל-ListTemplate.
 *
 * השלילות אינן פחות חשובות מהחיוביות: `ב.` ו-`5.` ו-`א'` הם בדיוק ההבדל בין
 * „זיהוי רשימות” לבין „כל פסקה שנייה במסמך תורני הופכת לרשימה".
 */
import { describe, expect, it } from 'vitest';
import { planListAutoformat } from '../../src/engine/list-autoformat';

const ALEF = 'א';
const BET = 'ב';
const BULLET = '•';

describe('planListAutoformat — ממוספרות', () => {
  it.each([
    ['1. ', 'decimal', '%1.', 1],
    ['1) ', 'decimal', '%1)', 1],
    ['1- ', 'decimal', '%1-', 1],
    ['1> ', 'decimal', '%1>', 1],
    ['a. ', 'lowerLetter', '%1.', 1],
    ['a) ', 'lowerLetter', '%1)', 1],
    ['A. ', 'upperLetter', '%1.', 1],
  ])('%s → %s עם תבנית %s', (typed, numFmt, lvlText, startAt) => {
    expect(planListAutoformat(typed)).toEqual({
      kind: 'numbered',
      markerLength: typed.length,
      numFmt,
      lvlText,
      startAt,
    });
  });

  /*
   * הליבה של הבקשה. `hebrew1` הוא גימטריה (א ב ג … י יא יב) — נמדד
   * `NumberStyle=45` ב-Word, והוא גם ה-numFmt ש-lists.ts מודד כעובד במנוע.
   */
  it.each([
    [`${ALEF}. `, '%1.'],
    [`${ALEF}) `, '%1)'],
    [`${ALEF}- `, '%1-'],
    [`${ALEF}> `, '%1>'],
  ])('אות עברית %s → hebrew1 עם תבנית %s', (typed, lvlText) => {
    expect(planListAutoformat(typed)).toEqual({
      kind: 'numbered',
      markerLength: typed.length,
      numFmt: 'hebrew1',
      lvlText,
      startAt: 1,
    });
  });

  it.each([
    ['(1) ', 'decimal'],
    [`(${ALEF}) `, 'hebrew1'],
    ['(a) ', 'lowerLetter'],
  ])('הצורה העוטפת %s שומרת את שני הסוגריים', (typed, numFmt) => {
    expect(planListAutoformat(typed)).toEqual({
      kind: 'numbered',
      markerLength: typed.length,
      numFmt,
      lvlText: '(%1)',
      startAt: 1,
    });
  });

  it('אפס מתחיל מאפס — נמדד listValue=0 ב-Word', () => {
    expect(planListAutoformat('0. ')).toEqual({
      kind: 'numbered',
      markerLength: 3,
      numFmt: 'decimal',
      lvlText: '%1.',
      startAt: 0,
    });
  });

  it('Tab מפעיל כמו רווח', () => {
    expect(planListAutoformat('1.\t')).toMatchObject({ kind: 'numbered', numFmt: 'decimal' });
  });
});

describe('planListAutoformat — תבליטים', () => {
  it.each([
    ['* ', BULLET, 'Symbol'],
    ['> ', BULLET, 'Symbol'],
    ['-> ', BULLET, 'Symbol'],
    ['=> ', BULLET, 'Symbol'],
    // `+` נדחה ב-Word ומתקבל כאן — LibreOffice מקבל, ואין סיבה לשלילה.
    ['+ ', BULLET, 'Symbol'],
  ])('%s → תבליט', (typed, lvlText, markerFont) => {
    expect(planListAutoformat(typed)).toEqual({
      kind: 'bullet',
      markerLength: typed.length,
      lvlText,
      markerFont,
    });
  });

  it('מקף נשאר מקף בגופן Arial — כך Word צייר אותו', () => {
    expect(planListAutoformat('- ')).toEqual({
      kind: 'bullet',
      markerLength: 2,
      lvlText: '-',
      markerFont: 'Arial',
    });
  });

  it('מקף לבדו הוא תבליט, ואחרי ספרה הוא מפריד של מספור', () => {
    expect(planListAutoformat('- ')).toMatchObject({ kind: 'bullet' });
    expect(planListAutoformat('1- ')).toMatchObject({ kind: 'numbered', lvlText: '%1-' });
  });
});

describe('planListAutoformat — מה שנדחה', () => {
  it('אות עברית שאינה א — Word דוחה, וזו ההגנה על „ב׳ בניסן”', () => {
    expect(planListAutoformat(`${BET}. `)).toBeNull();
    expect(planListAutoformat(`${BET}) `)).toBeNull();
  });

  it('ספרה שאינה פותחת רצף', () => {
    expect(planListAutoformat('5. ')).toBeNull();
    expect(planListAutoformat('12. ')).toBeNull();
  });

  /*
   * הגרש הוא הצורה הנפוצה ביותר בכתיבה תורנית, ובדיוק לכן אינו מתקבל: „א׳
   * בניסן”, „ר' יוסי”. Word דוחה אותה גם הוא.
   */
  it('גרש אינו מפריד', () => {
    expect(planListAutoformat(`${ALEF}' `)).toBeNull();
    expect(planListAutoformat(`${ALEF}׳ `)).toBeNull();
  });

  /*
   * נמדד בשני כיווני קריאה של הפסקה: Word מקבל רק את התו הלוגי `)`. במסמך RTL
   * הוא מצטייר בצורת `(`, וזה אינו מה שמקלידים.
   */
  it('הסוגר הפותח אינו מפריד, גם אחרי אות עברית', () => {
    expect(planListAutoformat(`${ALEF}( `)).toBeNull();
    expect(planListAutoformat('1( ')).toBeNull();
    expect(planListAutoformat(`)${ALEF}( `)).toBeNull();
  });

  it('מילה אינה סמן', () => {
    expect(planListAutoformat('שלום. ')).toBeNull();
    expect(planListAutoformat('ראה. ')).toBeNull();
  });

  /*
   * תו ההפעלה הוא נקודת ההחלטה היחידה, וללא בדיקה מפורשת שלו כל תו שבא אחרי
   * סמן תקין היה מפעיל: „‏`1.5`” — ספרה שהוקלדה אחרי „‏`1.`” — הופך ל„סמן
   * `1.` ואחריו `5`”. זה נמצא בסבב מוטציה, אחרי ששלושת המקרים הראשונים כאן
   * נשארו ירוקים על מודול ששער ההפעלה הוסר ממנו.
   */
  it('בלי תו הפעלה אין המרה — זו נקודת ההחלטה היחידה', () => {
    expect(planListAutoformat('1.')).toBeNull();
    expect(planListAutoformat(`${ALEF})`)).toBeNull();
    expect(planListAutoformat('* ')).not.toBeNull();
    expect(planListAutoformat('*')).toBeNull();
  });

  it('תו רגיל אחרי סמן תקין אינו מפעיל', () => {
    expect(planListAutoformat('1.5')).toBeNull();
    expect(planListAutoformat(`${ALEF})${BET}`)).toBeNull();
    expect(planListAutoformat('-x')).toBeNull();
    expect(planListAutoformat('(1)x')).toBeNull();
  });

  it('הסמן חייב להיות כל מה שלפני הסמן בבלוק', () => {
    expect(planListAutoformat('שלום 1. ')).toBeNull();
    expect(planListAutoformat(' 1. ')).toBeNull();
    expect(planListAutoformat('1. שלום ')).toBeNull();
  });

  it('קינון אינו נתמך — Word עצמו אינו מזהה אותו בהקלדה', () => {
    expect(planListAutoformat('1.1. ')).toBeNull();
  });

  it('קלט ריק או פגום', () => {
    expect(planListAutoformat('')).toBeNull();
    expect(planListAutoformat(' ')).toBeNull();
    expect(planListAutoformat(undefined as unknown as string)).toBeNull();
    expect(planListAutoformat(null as unknown as string)).toBeNull();
  });

  /*
   * שער נגד ירושה מ-Object.prototype: `constructor` ו-`toString` הם מפתחות
   * שקיימים על כל אובייקט, והתאמה עליהם הייתה מחזירה פונקציה במקום `null`.
   */
  it('מפתח שירש מ-Object.prototype אינו סמן', () => {
    expect(planListAutoformat('constructor ')).toBeNull();
    expect(planListAutoformat('toString ')).toBeNull();
  });
});
