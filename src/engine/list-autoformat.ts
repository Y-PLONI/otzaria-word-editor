/**
 * זיהוי רשימות בהקלדה — „‏`א)` ורווח” הופך את הפסקה לרשימה ממוספרת בעברית.
 *
 * המודול **טהור**: מקבל את הטקסט שמתחילת הבלוק ועד הסמן, ומחזיר תוכנית החלה
 * או `null`. בלי DOM, בלי מנוע, בלי אסינכרוניות — ההחלה עצמה יושבת
 * ב-`list-autoformat-install.ts`.
 *
 * ## הכול כאן נמדד ב-Word, ולא הוסק ממנו
 *
 * הטבלאות המלאות ב-`docs/list-autoformat-research.md`. שלושת הדברים שקובעים
 * את הצורה של הקובץ הזה:
 *
 * 1. **נקודת ההחלטה היא הרווח (או ה-Tab), והפסקה עוד ריקה.** התיעוד הרווח
 *    אומר „הקלד סימן, רווח, טקסט, ואז Enter”; נמדד שזה לא נכון. אחרי `1` `.`
 *    `רווח` בלבד הפסקה כבר רשימה, והטקסט `1. ` **נבלע**. ולכן הקלט כאן הוא
 *    מחרוזת קצרה שעוגנת `^…$`, ולא חיפוש בתוך פסקה.
 *
 * 2. **המפריד שהוקלד נשמר כתבנית הסמן.** `א)` נותן `lvlText:'%1)'` ולא
 *    `'%1.'` — כל הרשימה תמשיך בסוגר. זה העיקר, לא קישוט: נמדד
 *    `NumberFormat='%1)'` על ה-ListTemplate שWord יצר.
 *
 * 3. **אות עברית ממופה ל-`hebrew1` — גימטריה** (א ב ג … י יא יב), לא לסדר
 *    האלף-בית. נמדד `NumberStyle=45` (`wdListNumberStyleHebrew1`). זה בדיוק
 *    ה-`numFmt` ש-`lists.ts` כבר מודד כעובד אצלנו.
 *
 * ## למה `ב.` ו-`5.` נדחים
 *
 * ב-Word הכלל אינו „רק 1” אלא „ערך שפותח רשימה, או ערך שממשיך סביר את הרשימה
 * הקודמת במסמך”: `5. ` במסמך ריק נדחה, אבל אחרי רשימה שהגיעה ל-2 ויצאו ממנה
 * `3. ` התקבל והתחיל מ-3 (נמדד). ההמשך-לפי-הקודמת אינו נבנה כאן, ולכן מתקבלים
 * **פותחי רצף בלבד**. זו גם ההגנה האמיתית מפני „ב׳ בניסן” ו„ג׳ ימים”.
 *
 * וזו התנהגות ה**מוצר** ולא רק של המודול הטהור, אבל לא בזכות הקובץ הזה: המנוע
 * ממיר „N. ” לכל מספר בעצמו, ולכן דחייה כאן לבדה הייתה משאירה את ההמרה בידיו —
 * „1948. ” בתחילת פסקה היה הופך לרשימה עשרונית. מה שסוגר את זה הוא
 * `interceptsSpace` ב-`list-autoformat-install.ts`, שעוצר כל צורה שהמנוע ממיר
 * בעצמו. שני הקבצים יחד הם הכלל; אף אחד מהם לבדו אינו.
 *
 * ## ‏`i.` ו-`a)` — מתקבלים, אף שהעורך עברי
 *
 * נשקל לדחות אותם מאותו טעם שדחה את `א'`: בעורך שכל תוכנו עברי, אות לטינית
 * בתחילת שורה היא סביר יותר ציטוט מכוונת מספור. הם נשארו, משתי סיבות. ראשית,
 * Word מקבל אותם, והחריגה ממנו כאן הייתה **בלי מדידה** — בעוד הדחייה של `א'`
 * נשענת על מדידה ש-Word עצמו דוחה אותה. ושנית, הסיכון אינו סימטרי: `א'` מופיע
 * בכל „ר' יוסי” ובכל תאריך עברי, ואילו פסקה שנפתחת ב-„`i. `” או ב-„`a) `”
 * ואינה מכוונת לרשימה היא נדירה. אם תימדד — הן יורדות.
 *
 * ## ו-`א'` בגרש — הצורה הנפוצה ביותר, ובכוונה אינה כאן
 *
 * Word דוחה אותה, ומאותה סיבה שאנחנו דוחים: „א׳ בניסן”, „ר' יוסי”, „ב' ימים”.
 * קבלה שלה הייתה הופכת פסקאות תמימות לרשימות בכל מסמך תורני. אם תידרש — מתג
 * נפרד, לא ברירת מחדל.
 */

/** רשימה ממוספרת: `numFmt` + תבנית הסמן + ערך ההתחלה. */
export interface NumberedAutoformat {
  kind: 'numbered';
  /** כמה תווים נבלעים, כולל תו ההפעלה. */
  markerLength: number;
  /** ערך `w:numFmt` — מהרשימה ש-`lists.ts` מאשר. */
  numFmt: string;
  /** תבנית הסמן, עם `%1` כמונה הרמה. */
  lvlText: string;
  /** `restartAt` נקרא רק כשזה אינו 1. */
  startAt: number;
}

/** רשימת תבליטים: הסמן עצמו והגופן שמצייר אותו. */
export interface BulletAutoformat {
  kind: 'bullet';
  markerLength: number;
  lvlText: string;
  markerFont: string;
}

export type ListAutoformat = NumberedAutoformat | BulletAutoformat;

/**
 * פותחי רצף בלבד — ראו ההערה על `ב.` למעלה. `0` הוא היחיד שמתחיל מאפס, וזה
 * נמדד: `0. ` הומר עם `listValue=0`.
 */
const SEEDS: Readonly<Record<string, { numFmt: string; startAt: number }>> = {
  '0': { numFmt: 'decimal', startAt: 0 },
  '1': { numFmt: 'decimal', startAt: 1 },
  a: { numFmt: 'lowerLetter', startAt: 1 },
  A: { numFmt: 'upperLetter', startAt: 1 },
  // א — `hebrew1`, גימטריה. נמדד NumberStyle=45 ב-Word.
  'א': { numFmt: 'hebrew1', startAt: 1 },
};

/**
 * ארבעת המפרידים שנמדדו מתקבלים ב-Word: `1.` `1)` `1-` `1>`.
 *
 * `(` **אינו** ביניהם, וזה נמדד ולא הונח: במסמך RTL הסוגר `)` שאחרי אות עברית
 * מצטייר בצורת `(`, ולכן נבדק אם Word מקבל גם את התו הלוגי השני. הוא אינו
 * מקבל — בשני כיווני הקריאה של הפסקה — והמקלדת העברית של Windows מייצרת
 * `)` ב-Shift+0 ממילא. אין פער, ואין מה להוסיף.
 */
const SEPARATORS = ['.', ')', '-', '>'] as const;

/**
 * תבליטים. `*` `-` `>` `->` `=>` נמדדו מתקבלים ב-Word; `+` נדחה שם, ומתקבל
 * כאן — LibreOffice מקבל אותו (`pBulletChar` ב-`autofmt.cxx`), ואין סיבה
 * לשחזר דווקא את השלילה הזאת.
 *
 * הסמן: `-` נשאר `-` בגופן Arial (נמדד), וכל השאר `•` בגופן Symbol — כמו
 * `BULLET_STYLE` ב-`lists.ts`. Word מצייר ל-`>`, `->` ו-`=>` חיצים בגופן
 * Wingdings; כאן לא, **בכוונה**: גופן סמלים שאינו מותקן מצייר סמן ריק, וזה
 * בדיוק הכשל שתועד ב-`lists.ts` על `hebrew1` בגרסה 2.8.0. לא נלקח סיכון כזה
 * בלי מדידה שהמנוע מצייר אותו.
 */
const BULLETS: Readonly<Record<string, { lvlText: string; markerFont: string }>> = {
  '*': { lvlText: '•', markerFont: 'Symbol' },
  '-': { lvlText: '-', markerFont: 'Arial' },
  '>': { lvlText: '•', markerFont: 'Symbol' },
  '->': { lvlText: '•', markerFont: 'Symbol' },
  '=>': { lvlText: '•', markerFont: 'Symbol' },
  '+': { lvlText: '•', markerFont: 'Symbol' },
};

/** רווח או Tab. שניהם נמדדו מפעילים ב-Word. */
function isTrigger(ch: string | undefined): boolean {
  return ch === ' ' || ch === '\t';
}

/**
 * התוכנית לטקסט שמתחילת הבלוק ועד הסמן, או `null` כשאין מה לעשות — וזה הרוב
 * המכריע של ההקשות.
 *
 * `text` חייב להיות **בדיוק** הסמן ותו ההפעלה. רווח מוביל („‏` 1. `”) נדחה:
 * הוא אינו מה שמקלידים, ו-Word עצמו מנקה אותו במסלול אחר (‏`DeleteLeadingBlanks`
 * ב-LibreOffice) שאינו חלק מהזיהוי. וגם המנוע אינו ממיר אותו במקומנו:
 * `interceptsSpace` עוצר את הרווח, ורווח שנעצר ואין לו תוכנית כאן נכתב כטקסט
 * רגיל — כלומר שתי השכבות דוחות אותו, ולא רק זו.
 *
 * הסדר — תבליטים לפני מספור — אינו שרירותי: `-` לבדו הוא תבליט, בעוד `1-` הוא
 * מפריד של רשימה ממוספרת. ההתאמה על הסמן **כולו** היא שמפרידה ביניהם.
 */
export function planListAutoformat(text: string): ListAutoformat | null {
  if (typeof text !== 'string' || text.length < 2) return null;
  if (!isTrigger(text[text.length - 1])) return null;

  const marker = text.slice(0, -1);
  const markerLength = text.length;

  const bullet = Object.prototype.hasOwnProperty.call(BULLETS, marker) ? BULLETS[marker] : undefined;
  if (bullet) {
    return { kind: 'bullet', markerLength, lvlText: bullet.lvlText, markerFont: bullet.markerFont };
  }

  // `(א)` — הצורה העוטפת. נמדדה מתקבלת, עם `lvlText:'(%1)'`.
  if (marker.length === 3 && marker[0] === '(' && marker[2] === ')') {
    const seed = SEEDS[marker[1]];
    return seed
      ? { kind: 'numbered', markerLength, numFmt: seed.numFmt, lvlText: '(%1)', startAt: seed.startAt }
      : null;
  }

  if (marker.length === 2) {
    const seed = SEEDS[marker[0]];
    const separator = marker[1];
    if (seed && (SEPARATORS as readonly string[]).includes(separator)) {
      return {
        kind: 'numbered',
        markerLength,
        numFmt: seed.numFmt,
        lvlText: `%1${separator}`,
        startAt: seed.startAt,
      };
    }
  }

  return null;
}
