/**
 * שכבת החבילה של DOCX: קורא וכותב ה-ZIP, הפורס והדוחס, וסורק ה-XML שבעזרתו
 * מתקנים חלקים — **בלי שום תלות אחרת**.
 *
 * ## למה מודול משלו
 *
 * שני כיוונים משתמשים באותה שכבה: `docx-preflight.ts` מתקן את מה שנכנס למנוע,
 * ו-`docx-postflight.ts` מתקן את מה שיוצא ממנו. כשהשכבה ישבה בתוך
 * ה-preflight, חיבור התיקון היוצא ל-`export.ts` יצר את שרשרת הייבוא
 * `vite.config.ts` → `blank-document.ts` → `export.ts` → preflight →
 * `vba-import.ts` → חבילת ESM-בלבד, וטעינת קובץ התצורה נשברה עליה. מודול בלי
 * תלויות אינו גורר איש, ולכן אף כיוון אינו גורר את השני.
 *
 * ## הכלל שחל על כל מה שכאן
 *
 * **לתקן, ולא לחסום.** כל כשל — ארכיון שאינו נקרא, חלק שאינו נפרס, דוחס שאינו
 * קיים — מחזיר `null` ולא זורק. מי שקורא מכאן ממשיך עם המקור כמות שהוא.
 */

/**
 * בייטים שגובים מ-`ArrayBuffer` רגיל.
 *
 * הכינוי מפורש מפני ש-TypeScript מבדיל מאז 5.7 בין `ArrayBuffer` ל-
 * `SharedArrayBuffer`, ו-`Uint8Array` סתם כולל את שניהם — צורה ש-`Blob`
 * ו-`DecompressionStream` אינם מקבלים.
 */
export type Bytes = Uint8Array<ArrayBuffer>;

const ZIP_LOCAL_SIGNATURE = 0x04034b50;
const ZIP_CENTRAL_SIGNATURE = 0x02014b50;
const ZIP_EOCD_SIGNATURE = 0x06054b50;
const ZIP_LOCAL_HEADER_SIZE = 30;
const ZIP_CENTRAL_HEADER_SIZE = 46;
const ZIP_EOCD_SIZE = 22;

/** דגל „הגדלים מגיעים אחרי הנתונים”. הכתיבה כאן תמיד יודעת אותם מראש. */
const ZIP_FLAG_DATA_DESCRIPTOR = 0x0008;
/** דגל הצפנה. ארכיון מוצפן אינו משהו שיש כאן מה לעשות איתו. */
const ZIP_FLAG_ENCRYPTED = 0x0001;

const METHOD_STORED = 0;
const METHOD_DEFLATE = 8;

/** „גרסה נדרשת” 2.0 — המינימום שרשומת deflate מצהירה עליו (APPNOTE 4.4.3.2). */
const ZIP_VERSION_DEFLATE = 20;

/** הסימן שהארכיון הוא ZIP64. אין כאלה ב-DOCX ריאלי, ולכן פשוט לא נוגעים בהם. */
const ZIP64_MARKER_32 = 0xffffffff;
const ZIP64_MARKER_16 = 0xffff;

export interface ZipEntry {
  /** שם החלק, לזיהוי. */
  name: string;
  /** בייטי השם כפי שהיו — כדי לא לקודד מחדש שם שאינו UTF-8. */
  nameBytes: Bytes;
  versionMadeBy: number;
  versionNeeded: number;
  flags: number;
  method: number;
  modTime: number;
  modDate: number;
  crc: number;
  internalAttrs: number;
  externalAttrs: number;
  /** התוכן כפי שהוא מאוחסן — דחוס או לא, לפי `method`. */
  data: Bytes;
  uncompressedSize: number;
  /**
   * שדה ה-extra של הכותרת המקומית ושל ספריית האינדקס, והערת הרשומה.
   *
   * שני שדות ה-extra נשמרים בנפרד מפני ש-APPNOTE מתיר להם להיות שונים,
   * ובפועל הם שונים (Info-ZIP כותב UT מלא מקומית ומקוצר במרכז). הם נכתבים
   * בחזרה כמות שהם, ורשומה שנבנתה כאן מאפס משאירה אותם ריקים.
   *
   * **מה שמותר להבטיח על כך, ומה שלא.** `readZip` דוחה את הסימן של ZIP64
   * בחמשת השדות שהוא **קורא**: `compressedSize`, `uncompressedSize`
   * ו-`localOffset` שברשומה המרכזית, ו-`entriesTotal` ו-`cdOffset` שב-EOCD.
   * אלה גם השדות היחידים שהכתיבה כאן מחשבת מחדש, ולכן שדה extra שתלוי
   * בהם אינו יכול לעבור. מה שאינו נבדק, ונמדד: `diskNumberStart` (היסט 34
   * ברשומה המרכזית) הוא מקום סימן חוקי לפי APPNOTE 4.5.3 ומתקבל כאן, וכך
   * גם ארבעת שדות ה-EOCD שאיש אינו קורא (`thisDisk`, `cdStartDisk`,
   * `entriesThisDisk`, `cdSize`). הארבעה אינם מזיקים; `diskNumberStart` כן
   * נכתב בחזרה כאפס קבוע בעוד ה-extra שנושא את הערך האמיתי נשמר — כלומר
   * חבילה מרובת-דיסקים תצא מכאן חסרת עקביות. לא נוסף לו שומר: חבילת OOXML
   * מרובת-דיסקים אינה מקרה שנמדד, וההבטחה מנוסחת קטנה במקומו.
   *
   * ומה שנשאר פתוח בלי מדידה: רשומת extra של ZIP64 (מזהה 0x0001) רשאית
   * להופיע **בלי שום סימן** — יש כותבים שפולטים אותה תמיד — ואיש כאן אינו
   * מפרש מזהי רשומות בתוך שדה ה-extra. דחיית הסימן אינה מוכיחה, אם כן,
   * שאין כאן רשומה כזאת; היא רק מוכיחה שהערכים שהכתיבה משנה אינם יושבים
   * בה.
   *
   * **ולפי שעה זו שמירה רדומה.** נמדד על שלושה-עשר קבצי ה-docx שבעץ
   * (`tmp/sp3946`): בכל רשומה בכולם אורך שני שדות ה-extra הוא 0, וכך גם
   * אורך הערת הרשומה והערת הארכיון. **והראיה חלשה מכפי שהמספר נשמע:**
   * שניים-עשר מהם נוצרו ב-`tmp/sp3946/make-docx.mjs`, שכותב zip מינימלי
   * משלו ואינו מסוגל לפלוט שדה extra כלל; רק `case.docx` הוא חבילה שיצאה
   * מ-Word. מה שהשמירה שומרת עליו הוא חבילה שנארזה בכלי אחר — Info-ZIP
   * כותב UT בכל רשומה — ולא חבילה של Word.
   */
  localExtra?: Bytes;
  centralExtra?: Bytes;
  comment?: Bytes;
}

/**
 * החלקים שבהם יושבות תכונות ריצה, וכולם באותה סכימה של WordprocessingML.
 *
 * `styles.xml` הוא הרוב המעשי — כותרת מודגשת היא כמעט תמיד סגנון ולא עיצוב
 * ישיר — אבל עיצוב ישיר קיים, וכך גם כותרת עליונה, הערת שוליים ומספור רשימה.
 * כולם עוברים באותו כלל אחד, מפני שזה **אותו** כלל: `rPr` היא `rPr`.
 *
 * מה שאינו כאן ובכוונה: `word/glossary/*` (בלוקים לשימוש חוזר; אינם מרונדרים)
 * ו-`word/settings.xml`, שיש לו תיקון משלו.
 *
 * `stylesWithEffects.xml` **כן** כאן: זה חלק אמיתי של Word 2010 ובו גיליון
 * סגנונות שלם. Word 2010 קורא מ-`bCs` בעצמו ולכן אינו נפגע, אבל להשאיר גיליון
 * סגנונות שלם מחוץ לכלל „`rPr` היא `rPr`” הוא חוסר עקביות, לא החלטה.
 *
 * הספרה אופציונלית בכל השמות שיכולים לשאת אותה, ולא רק בארבעה מהם — הצורה
 * הקודמת התירה `document2` ולא `footnotes2`, וזו הייתה השמטה ולא כלל.
 */
export const CONTENT_PARTS =
  /^word\/(?:document|styles|stylesWithEffects|numbering|footnotes|endnotes|comments|header|footer)\d*\.xml$/i;

/**
 * הכנסה אחת לתוך XML: ההיסט, הטקסט, והטווח שההכנסה חייבת ליפול בתוכו.
 *
 * הטווח נישא על ההכנסה ואינו מחושב מחדש — חיפוש שלו בזמן ההחלה היה סריקה על
 * כל הטווחים לכל הכנסה, כלומר ריבועי במספר הפסקאות (נמדד: 2,229ms על 40,000
 * פסקאות מול 171ms בלעדיו).
 */
export interface XmlInsert {
  at: number;
  text: string;
  span: { from: number; to: number };
}

/**
 * מחילה הכנסות על XML, או `null` כשהן מפרות את האינווריאנטה.
 *
 * **הפלט הוא הקלט ועוד הטקסטים שהוכנסו, או שאין פלט.** שום בייט של הקלט אינו
 * נמחק ואינו מוזז, וזה נאכף ולא מונח: כל היסט חייב להיות שלם, ממוין, בתוך
 * גבולות המחרוזת, ובתוך הטווח שההכנסה הצהירה עליו — והפלט נבנה מחיתוכים
 * רציפים שאורכם הכולל חייב להיות אורך הקלט.
 *
 * זו אינווריאנטה ולא נפילה-לאחור, ואין קלט שמגיע אליה דרך הקוראים של היום.
 * היא כאן מפני שמצב הכשל של הגישה שנפסלה (סימון `w:rtl`) לא היה זריקה אלא
 * פלט שגוי בשקט — `slice` על היסט שלילי שהחזיר מסמך משוסע בלי שאיש ישים לב.
 * נמדד שהטריגר הוא עריכה אחת מכאן: מסמך עם טקסט ואחריו תיבת טקסט מעוגנת סוגר
 * את הפסקה הפנימית ראשונה ודוחף את ההיסט הגדול לפני הקטן, ובלי המיון ובלי
 * הבדיקות אותו קלט מייצר פסקה משוכפלת וטקסט שנעלם.
 */
export function applyXmlInserts(xml: string, inserts: readonly XmlInsert[]): string | null {
  const sorted = [...inserts].sort((first, second) => first.at - second.at);
  const parts: string[] = [];
  let at = 0;
  let kept = 0;
  for (const { at: offset, text, span } of sorted) {
    if (!Number.isInteger(offset) || offset < at || offset > xml.length) return null;
    if (offset < span.from || offset > span.to) return null;
    const slice = xml.slice(at, offset);
    kept += slice.length;
    parts.push(slice, text);
    at = offset;
  }
  const rest = xml.slice(at);
  kept += rest.length;
  if (kept !== xml.length) return null;
  parts.push(rest);
  return parts.join('');
}

/**
 * עריכה אחת: החלפת `[at, end)` ב-`text` — הכנסה כש-`end === at`. כמו
 * `XmlInsert`, היא נושאת את הטווח שהיא חייבת ליפול בתוכו.
 */
export interface XmlEdit {
  at: number;
  end: number;
  text: string;
  span: { from: number; to: number };
}

/**
 * מחילה עריכות על XML, או `null` כשהן מפרות את האינווריאנטה.
 *
 * אותה אכיפה כמו ב-`applyXmlInserts`, מורחבת להחלפה: כל עריכה חייבת להיות
 * שלמה, בתוך גבולות המחרוזת ובתוך הטווח שהצהירה עליו, והעריכות אינן חופפות.
 * מה שאינו מוחלף מועתק ברצף, ואורכו ועוד אורך מה שהוחלף חייב להיות אורך הקלט.
 *
 * כמה עריכות באותו היסט נכתבות בסדר שבו נמסרו (מיון יציב). עריכה שמתחילה
 * לפני סוף ההחלפה שלפניה — חפיפה — מחזירה `null`.
 */
export function applyXmlEdits(xml: string, edits: readonly XmlEdit[]): string | null {
  const sorted = [...edits].sort((first, second) => first.at - second.at);
  const parts: string[] = [];
  let at = 0;
  let kept = 0;
  let replaced = 0;
  for (const { at: from, end, text, span } of sorted) {
    if (!Number.isInteger(from) || !Number.isInteger(end)) return null;
    if (from < at || end < from || end > xml.length) return null;
    if (from < span.from || end > span.to) return null;
    const slice = xml.slice(at, from);
    kept += slice.length;
    replaced += end - from;
    parts.push(slice, text);
    at = end;
  }
  const rest = xml.slice(at);
  kept += rest.length;
  if (kept + replaced !== xml.length) return null;
  parts.push(rest);
  return parts.join('');
}

/**
 * ערכי `ST_OnOff` שמשמעותם „כבוי”. כל ערך אחר — ובכלל זה היעדר `w:val` — דולק,
 * וזה מה שהתקן אומר: `<w:bCs/>` בלי מאפיין היא הדגשה פעילה.
 */
const OFF_VALUES = new Set(['0', 'false', 'off']);

/**
 * מאפיין `w:val`, לשני התיקונים כאחד.
 *
 * שני סוגי המרכאות, ולא רק כפולות: XML מתיר את שניהם, ו-`w:val='0'` שנקרא
 * כדולק היה הופך „לא מודגש” שנכתב במפורש למודגש — כלומר שינוי במסמך, בדיוק מה
 * שהמודול הזה מבטיח לא לעשות. הקידומת אופציונלית מאותו טעם שהסורק אינו נעול
 * על `w:` (ראו TOKEN_SOURCE).
 *
 * התחילית היא `\s` ולא `\b`, כדי שהמאפיין יתחיל במקום שבו מאפיין באמת מתחיל:
 * `\b` היה מתאים גם ל-`val='0'` **בתוך ערך** של מאפיין אחר. (`\s` אינו סוגר
 * את המקרה של רווח בתוך ערך כזה, למשל `w:foo="a val='0'"`, ולשם צריך פרסר
 * ולא רגקס. לאף אחד משני האלמנטים שכאן אין מאפיין מלבד `val`.)
 *
 * מה שנשאר לא-מטופל: ישות מספרית (`w:val="&#48;"`). היא חוקית, ואף כלי
 * מציאותי אינו כותב אותה.
 *
 * קבוצה 1 — ערך במרכאות כפולות; קבוצה 2 — בבודדות.
 */
export const VAL_ATTRIBUTE = /\s(?:[\w.-]+:)?val\s*=\s*(?:"([^"]*)"|'([^']*)')/;

/** הערך שב-`w:val`, או `null` כשאינו שם. */
export function valueOf(attributes: string): string | null {
  const match = VAL_ATTRIBUTE.exec(attributes);
  return match ? (match[1] ?? match[2]) : null;
}

/** האם דגל `ST_OnOff` דולק, לפי מאפייני התג. */
export function isOn(attributes: string): boolean {
  const value = valueOf(attributes);
  return value === null || !OFF_VALUES.has(value.trim().toLowerCase());
}

/**
 * הסורק: תג, פתיחת הערה, או פתיחת CDATA — לפי סדר הופעתם.
 *
 * **המאפיינים מודעים למרכאות** (`[^>"']` או מחרוזת מצוטטת), ולא `[^>]*`. זה
 * אינו הידור: `<w:rPrChange w:author="a>b">` הוא XML חוקי לגמרי, ורגקס שנעצר
 * על ה-`>` הראשון היה קורא אותו כתג אחר לגמרי.
 *
 * הקידומת נלכדת ואינה נעולה על `w`: החבילה רשאית לקשור את מרחב השמות של
 * WordprocessingML לכל קידומת. מי שנעול על `w:` גם מפספס מסמך כזה לגמרי, וגם —
 * גרוע יותר — אינו רואה `ns0:b` קיימת ומוסיף `w:b` שנייה לצדה.
 *
 * **אבל קידומת היא חובה כאן, וזו מגבלה:** חבילה שקושרת את מרחב השמות
 * כברירת מחדל (`<document xmlns="…/wordprocessingml/2006/main">` ואז `<bCs/>`
 * בלי קידומת) אינה מותאמת כלל, ושני התיקונים פשוט אינם קורים — בשקט. Word
 * כותב קידומת תמיד; מחולל צד-שלישי אינו חייב. הכיוון שמרני (לא לתקן, ולא
 * לתקן לא נכון), ולכן זה מתועד ולא נסגר: לזהות „האם התג הזה בכלל
 * WordprocessingML” בלי קידומת דורש מעקב אחר הכרזות מרחב שמות, כלומר פרסר.
 *
 * הערות, CDATA והוראות עיבוד נבלעות שלמות. שלושתן נראות לרגקס בדיוק כמו תגים
 * (`<!-- <w:rPr><w:bCs/></w:rPr> -->`), והתיקון בתוכן היה עריכה של טקסט
 * המשתמש — או של הצהרה — ולא של העיצוב. הצהרת ה-XML עצמה (`<?xml … ?>`) לא
 * הותאמה גם קודם, מפני ש-`?` אינו ב-`[\w.-]`; מה שנסגר כאן הוא הוראת עיבוד
 * שיש **בתוכה** משהו שנראה כמו תג.
 */
export const TOKEN_SOURCE =
  /<!--|<!\[CDATA\[|<\?|<(\/?)([\w.-]+):([\w.-]+)((?:[^>"']|"[^"]*"|'[^']*')*)>/;

/**
 * לכל פותח כזה — הסוגר שלו. מה שביניהם אינו XML שיש בו מה לתקן, וכל מה
 * שהסורק צריך לעשות איתו הוא לדלג עליו שלם.
 */
export const SKIPPED_SPANS = new Map([
  ['<!--', '-->'],
  ['<![CDATA[', ']]>'],
  ['<?', '?>'],
]);

/**
 * הסוגר של מה שהסורק חייב לבלוע שלם — או `undefined` לתג רגיל.
 *
 * זהה ל-`SKIPPED_SPANS.get(token)`, בלי לחשב hash של מחרוזת התג המלאה על כל
 * תג במסמך: הערה, CDATA והוראת עיבוד הן היחידות שהתו השני שלהן `!` או `?`.
 * נמדד כחלק ניכר מזמן הסריקה של מסמך גדול.
 */
export function skippedCloser(token: string): string | undefined {
  const second = token.charCodeAt(1);
  return second === 0x21 || second === 0x3f ? SKIPPED_SPANS.get(token) : undefined;
}

/** מרחב השמות של WordprocessingML. הקידומת נגזרת ממנו ואינה מונחת. */
const WORDPROCESSING_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * הקידומת שקשורה למרחב השמות של WordprocessingML, או `null` כשאין הצהרה.
 *
 * הקידומת ולא השם המקומי לבדו: ‏`<m:r>` של נוסחה ו-`<a:r>` של DrawingML נושאים
 * בדיוק את אותם שמות מקומיים ואינם ריצות Word. סורק שמתאים לפי השם בלבד היה
 * נוגע בהם; ול-`m:rPr` אין `m:bCs` בסכמה בכלל.
 *
 * ההצהרה נקראת מ**תג השורש בלבד**, ולא מהמחרוזת כולה. נמדד שהחיפוש החופשי
 * נשבר: הערה שיושבת לפני השורש וכתוב בה `xmlns:m="…/main"` גרמה לסורק לקרוא
 * את כל הנוסחאות שבמסמך כפסקאות Word ולהכניס RLM לתוכן, בעוד הפסקאות
 * האמיתיות לא נגעו. זה גם המקום הנכון: בחבילת OOXML ההצהרה יושבת על השורש.
 * הצהרה שיושבת עמוק יותר מחזירה `null`, כלומר אי-פעולה.
 */
export function wordPrefix(xml: string): string | null {
  const declaration = new RegExp(
    `\\sxmlns:([\\w.-]+)\\s*=\\s*"${WORDPROCESSING_NS}"|\\sxmlns:([\\w.-]+)\\s*=\\s*'${WORDPROCESSING_NS}'`,
  );
  const token = new RegExp(TOKEN_SOURCE.source, 'g');
  for (let match = token.exec(xml); match; match = token.exec(xml)) {
    const closer = SKIPPED_SPANS.get(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      if (end < 0) return null;
      token.lastIndex = end + closer.length;
      continue;
    }
    // התג האמיתי הראשון הוא השורש. תג סוגר במקום הזה הוא XML פגום.
    if (match[1]) return null;
    const found = declaration.exec(match[4] ?? '');
    return found ? (found[1] ?? found[2] ?? null) : null;
  }
  return null;
}

/** הביטוי של כל שם מאפיין, נבנה פעם אחת: `rFonts` נקרא כמעט בכל ריצה במסמך. */
const ATTRIBUTE_PATTERNS = new Map<string, RegExp>();

/** מאפיין לפי שם, בלי להיות נעול על קידומת — כמו `valueOf`, ומאותו טעם. */
export function attribute(attributes: string, name: string): string | null {
  let pattern = ATTRIBUTE_PATTERNS.get(name);
  if (!pattern) {
    pattern = new RegExp(`\\s(?:[\\w.-]+:)?${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`);
    ATTRIBUTE_PATTERNS.set(name, pattern);
  }
  const match = pattern.exec(attributes);
  return match ? (match[1] ?? match[2]) : null;
}

/**
 * ערך שנקרא ממאפיין וייכתב לתוך מאפיין חדש במרכאות כפולות.
 *
 * רק `"` מוחלף: מה שנקרא הוא הטקסט הגולמי שבין המרכאות, כלומר כבר מוגן —
 * `&amp;` שנקרא כך וייכתב כך הוא אותו ערך בדיוק. גרש כפול יכול להופיע בו רק
 * כשהמקור היה במרכאות בודדות, וזה המקרה היחיד שצריך הגנה.
 */
export function quoteAttribute(raw: string): string {
  return raw.replace(/"/g, '&quot;');
}

/**
 * החלת שינוי על חלקי XML של DOCX, והחזרת הבייטים החדשים — או `null` כששום חלק
 * לא השתנה.
 *
 * זהו אותו שלד ש-`repairEntries` רץ עליו, מיוצא כדי שגם תיקון **בכיוון היוצא**
 * (engine/docx-postflight.ts) יקבל את קורא ה-zip, את הפורס ואת הכותב בלי
 * להעתיק אף אחד מהם. מה ש-`repairEntries` מוסיף מעליו — רשימת ההערות וההודעה
 * למשתמש — שייך לפתיחה בלבד, ולכן אינו כאן.
 *
 * אותו כלל בדיוק: חלק שאינו נקרא מדולג ואינו מפיל את השאר, וכשלון מוחלט מחזיר
 * `null` — כלומר „אין לי מה לומר על המסמך הזה”, ולא מסמך פגום.
 */
export async function rewriteDocxXmlParts<Context = undefined>(
  bytes: Bytes,
  matches: (name: string) => boolean,
  transform: (xml: string, name: string, context: Context) => string | null,
  prepare?: (read: (name: string) => Promise<string | null>, names: readonly string[]) => Promise<Context>,
): Promise<Bytes | null> {
  const entries = readZip(bytes);
  if (!entries) return null;

  // חלק שנקרא ב-`prepare` אינו נפרס פעמיים: הטקסט נשמר עד שהלולאה מגיעה אליו.
  const texts = new Map<ZipEntry, Promise<string | null>>();
  const textOf = (entry: ZipEntry): Promise<string | null> => {
    let text = texts.get(entry);
    if (!text) {
      text = readEntryText(entry);
      texts.set(entry, text);
    }
    return text;
  };
  const byName = new Map(entries.map((entry) => [entry.name, entry]));
  // ‏`prepare` שנכשל — כמו `transform` שנכשל: הוא אינו מפיל את השמירה. מה שהוא
  // היה מכין חסר, ולכן אין גם במה לתקן; הבייטים יוצאים כמות שהם.
  let context: Context;
  try {
    context = prepare
      ? await prepare(
          (name) => {
            const entry = byName.get(name);
            return entry ? textOf(entry) : Promise.resolve(null);
          },
          entries.map((entry) => entry.name),
        )
      : (undefined as Context);
  } catch (error) {
    console.warn('[otzaria-word] ההכנה לתיקון החלקים זרקה, והמסמך נשאר כמות שהוא', error);
    return null;
  }

  const patched = new Map<ZipEntry, ZipEntry>();
  for (const entry of entries) {
    if (!matches(entry.name)) continue;
    const original = await textOf(entry);
    texts.delete(entry);
    if (original === null) continue;

    let next: string | null;
    try {
      next = transform(original, entry.name, context);
    } catch (error) {
      // ההבטחה בראש הקובץ היא `null` ולא זריקה, והיא לא הייתה נשמרת:
      // קריאה חוזרת שזורקת על חלק אחד הייתה דוחה את ההבטחה כולה, כלומר
      // מפילה את **כל** השמירה בגלל חלק אחד. כאן היא מפילה חלק אחד בלבד:
      // הוא נכתב כמות שהוא, והשאר ממשיך — בדיוק כמו `original === null`
      // שורה אחת למעלה. (זה **אינו** מה ש-`preflightSource` עושה סביב
      // `repairEntries`: שם ה-catch מחזיר את המסמך **כולו** בלי תיקון.)
      //
      // ואף `transform` בעץ אינו מגיע לכאן: `markNeutralParagraphEnds`
      // מחזיר `null` על כל הפרת חזקה, ו-`uniqueNumberingIds` הוא רגקסים
      // ומחרוזות. אין בשניהם `throw`. השורה נשארת מפני שהיא חוזה של
      // הפונקציה המיוצאת הזאת כלפי מי שיעביר לה תיקון שלישי.
      console.warn(`[otzaria-word] תיקון החלק ${entry.name} זרק, והחלק נשאר כמות שהוא`, error);
      continue;
    }
    // וזה הכשל שכן נמדד. `TextEncoder.encode` אינו זורק על מה שאינו
    // מחרוזת אלא **מקודד את הצורה שלו**: callback שהפך בטעות ל-`async`
    // מחזיר Promise והחלק יוצא `"[object Promise]"`, ו-`return` שנשמט
    // מחזיר `undefined` והחלק יוצא **ריק**. בשני המקרים הארכיון תקין
    // לגמרי — חתימות, CRC וגדלים נכונים — והמסמך של המשתמש הרוס בשקט.
    // „לתקן, ולא לחסום” פירושו כאן: לא לגעת בחלק שאיננו מבינים את הפלט
    // עליו. נכונות **התוכן** נשארת של מי שכותב את התיקון; מה שנבדק כאן
    // הוא רק שקיבלנו טקסט.
    if (typeof next !== 'string' && next !== null) {
      console.warn(`[otzaria-word] תיקון החלק ${entry.name} לא החזיר טקסט, והחלק נשאר כמות שהוא`);
      continue;
    }
    if (next === null || next === original) continue;
    patched.set(entry, await rewriteEntry(entry, new TextEncoder().encode(next)));
  }

  if (patched.size === 0) return null;
  return writeZip(entries.map((entry) => patched.get(entry) ?? entry), readZipComment(bytes));
}

/**
 * הרשומה של חלק שתוקן: דחוסה כשאפשר, ו-`STORED` כשלא.
 *
 * ה-CRC והגודל הלא-דחוס נמדדים על התוכן, ולא על מה שנכתב — כך ZIP מגדיר אותם,
 * וכך גם קורא שפורס את הרשומה יודע לאמת אותה. הדחיסה היא **בונוס**, לא תנאי:
 * אם `CompressionStream` חסר או נכשל, הרשומה נכתבת גלויה, ובשני המקרים הקורא
 * מקבל בדיוק את אותם בייטים אחרי הפריסה.
 *
 * `versionNeeded` מורם ל-2.0 כשנכתב deflate: רשומה שהמקור שלה היה `STORED`
 * (1.0) ועכשיו דחוסה מצהירה על מה שהיא. מקור שכבר היה 2.0 ומעלה נשאר כפי שהוא.
 */
export async function rewriteEntry(entry: ZipEntry, content: Bytes): Promise<ZipEntry> {
  const deflated = await deflateVerified(content);
  return {
    ...entry,
    flags: entry.flags & ~ZIP_FLAG_DATA_DESCRIPTOR,
    method: deflated ? METHOD_DEFLATE : METHOD_STORED,
    versionNeeded: deflated ? Math.max(entry.versionNeeded, ZIP_VERSION_DEFLATE) : entry.versionNeeded,
    crc: crc32(content),
    data: deflated ?? content,
    uncompressedSize: content.byteLength,
  };
}

/**
 * פענוח XML של חלק, **בלי לפשוט את ה-BOM**.
 *
 * ברירת המחדל של `TextDecoder` מוחקת U+FEFF, ו-`TextEncoder` אינו מחזיר אותו,
 * ולכן חלק שהתחיל ב-BOM היה נכתב מחדש בלעדיו. אין לזה נזק תפקודי — ההצהרה
 * אומרת UTF-8 — אבל ההבטחה בכותרת הקובץ היא „זהה למקור בכל מה שאינו התיקון
 * עצמו”, וזה חלק ממנה.
 */
export function decodeXml(bytes: Bytes): string {
  return new TextDecoder('utf-8', { ignoreBOM: true }).decode(bytes);
}

/** תוכן החלק כטקסט, או `null` כשאי אפשר לפרוס אותו. */
export async function readEntryText(entry: ZipEntry): Promise<string | null> {
  if (entry.method === METHOD_STORED) return decodeXml(entry.data);
  if (entry.method !== METHOD_DEFLATE) return null;

  const inflated = await inflateRaw(entry.data);
  return inflated && decodeXml(inflated);
}

/**
 * פריסת deflate גולמי דרך `DecompressionStream`.
 *
 * `null` כשהסביבה אינה מכירה אותו או כשהנתונים אינם נפרסים — שני מקרים שבהם
 * אין לנו מה לומר על המסמך, ולכן הוא נמסר למנוע כמות שהוא.
 */
function inflateRaw(data: Bytes): Promise<Bytes | null> {
  const Decompression = (globalThis as { DecompressionStream?: typeof DecompressionStream })
    .DecompressionStream;
  if (!Decompression) return Promise.resolve(null);
  return pipeThrough(data, () => new Decompression('deflate-raw'), 'פריסת חלק מהמסמך נכשלה');
}

/**
 * דחיסת deflate גולמי דרך `CompressionStream` — ההופכי של `inflateRaw`, באותו
 * API ובאותה נפילה-חזרה.
 *
 * `null` כשהסביבה אינה מכירה אותו או כשהדחיסה נכשלה. בשני המקרים החלק נכתב
 * `STORED`, ולכן זה אינו כשל של התיקון אלא רק ויתור על החיסכון בזיכרון.
 */
function deflateRaw(data: Bytes): Promise<Bytes | null> {
  const Compression = (globalThis as { CompressionStream?: typeof CompressionStream })
    .CompressionStream;
  if (!Compression) return Promise.resolve(null);
  return pipeThrough(data, () => new Compression('deflate-raw'), 'דחיסת חלק מהמסמך נכשלה');
}

/**
 * דחיסה שאומתה: הפלט נפרס בחזרה ומושווה למקור לפני שהוא נכתב.
 *
 * זה מה שעונה על החשש שבגללו לא הייתה כאן דחיסה עד עכשיו — „דוחס שמתנהג אחרת
 * מהצפוי הוא באג שקט במסמך של המשתמש”. אחרי סבב מלא של דחיסה-ופריסה עם
 * השוואת CRC, דוחס שגוי אינו יכול להיות שקט: הוא נופל כאן ל-`STORED`. המחיר
 * הוא פריסה נוספת של החלק, ובשביל `document.xml` של ספר זה עשרות מילישניות
 * על פתיחה שממילא נמדדת במאות.
 *
 * `null` גם כשהדחיסה אינה קטנה מהמקור: רשומה גלויה קצרה יותר היא פשוט
 * הרשומה הנכונה.
 */
async function deflateVerified(content: Bytes): Promise<Bytes | null> {
  const deflated = await deflateRaw(content);
  if (!deflated || deflated.byteLength >= content.byteLength) return null;

  const restored = await inflateRaw(deflated);
  if (!restored || restored.byteLength !== content.byteLength || crc32(restored) !== crc32(content)) {
    console.warn('[otzaria-word] הדחיסה לא שחזרה את החלק בדיוק — נכתב לא-דחוס');
    return null;
  }
  return deflated;
}

/** מעבירה בייטים דרך זרם-טרנספורמציה ואוספת את הפלט. `null` ומיומן על כשל. */
async function pipeThrough(
  data: Bytes,
  transform: () => ReadableWritablePair<Uint8Array, BufferSource>,
  failure: string,
): Promise<Bytes | null> {
  try {
    const source = new ReadableStream<BufferSource>({
      start(controller) {
        controller.enqueue(data);
        controller.close();
      },
    });
    const reader = source.pipeThrough(transform()).getReader();

    const chunks: Uint8Array[] = [];
    let total = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      total += value.byteLength;
    }

    const out = new Uint8Array(total);
    let at = 0;
    for (const chunk of chunks) {
      out.set(chunk, at);
      at += chunk.byteLength;
    }
    return out;
  } catch (error) {
    console.warn(`[otzaria-word] ${failure}`, error);
    return null;
  }
}

/**
 * קריאת הארכיון מהספרייה המרכזית שלו — ולא מסריקת כותרות מקומיות, שהיא ניחוש
 * כשיש בהן data descriptor. `null` פירושו „לא ארכיון שאני מבין”, וזו תשובה
 * חוקית לגמרי: המנוע יקבל את המקור.
 *
 * **הרשומות חוזרות בסדר הפיזי שלהן, ולא בסדר הספרייה המרכזית.** APPNOTE אינו
 * מחייב שהשניים יהיו זהים, ו-OPC כן מחייב ש-`[Content_Types].xml` יהיה הרשומה
 * הראשונה בארכיון. הכותב כאן מסדר את הכותרות המקומיות לפי הסדר שהוא מקבל,
 * ולכן קריאה בסדר הספרייה הייתה יכולה להזיז אותו ממקומו — נמדד: ארכיון
 * שהספרייה שלו מונה את `word/document.xml` ראשון, בעוד `[Content_Types].xml`
 * הוא הראשון פיזית, יצא מכאן עם `word/document.xml` ככותרת המקומית הראשונה.
 * בכל שלושה-עשר קבצי ה-docx שבעץ שני הסדרים כבר זהים, ולכן זו שמירה על מה
 * שעוד לא נשבר — אבל הכותב רץ עכשיו על כל שמירה, ולא רק כשתוקן משהו.
 * (שניים-עשר מהשלושה-עשר נכתבו באותו כלי מינימלי, ולכן אינם יכולים להראות
 * סדר חורג ממילא; הראיה היא `case.docx` בלבד. ראו הערת `localExtra`.)
 *
 * ולסדר הזה יש קורא שלישי, בלי קשר ל-OPC: `.find()` לפי שם רשומה —
 * `readDocxPart`/`partText` ב-`docx-preflight.ts` — מכריע מעכשיו כפילות שם
 * לפי הרשומה הפיזית הראשונה ולא לפי הראשונה שבספרייה. אין מקרה נמדד של שם
 * כפול בחבילת OOXML, והסדר הפיזי הוא ממילא זה שקוראי ZIP אחרים רואים.
 */
export function readZip(bytes: Bytes): ZipEntry[] | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = findEocd(view);
  if (eocd < 0) return null;

  const count = view.getUint16(eocd + 10, true);
  const centralOffset = view.getUint32(eocd + 16, true);
  if (count === ZIP64_MARKER_16 || centralOffset === ZIP64_MARKER_32) return null;

  const entries: { entry: ZipEntry; localOffset: number }[] = [];
  let at = centralOffset;
  for (let i = 0; i < count; i++) {
    // מה שהשורה הזאת שומרת עליו הוא ההבטחה „`null` ולא זריקה”, ולא נכונות:
    // `centralOffset` מגיע מה-EOCD ויכול להצביע לכל מקום, ו-`getUint32`
    // שחורג **זורק** RangeError. נמדד: היסט ספרייה בשלושת הבייטים האחרונים
    // של הקובץ, או 0xFFFFFFF0, מפיל את הקריאה בלי השורה.
    //
    // והגבול הוא אורך הקובץ ולא ה-EOCD **בכוונה**: הצורה ההדוקה נכתבה כאן
    // ונמדדה בלתי-ניתנת-להבחנה — כל רשומה שמתחילה אחרי ה-EOCD נדחית ממילא
    // בבדיקת אורך הרשומה שלמטה. 200,000 ארכיונים פגומים (6,212 מהם נקראו
    // בהצלחה), אפס הפרשים בין שתי הצורות.
    if (at + ZIP_CENTRAL_HEADER_SIZE > bytes.byteLength) return null;
    if (view.getUint32(at, true) !== ZIP_CENTRAL_SIGNATURE) return null;

    const flags = view.getUint16(at + 8, true);
    if (flags & ZIP_FLAG_ENCRYPTED) return null;

    const compressedSize = view.getUint32(at + 20, true);
    const uncompressedSize = view.getUint32(at + 24, true);
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    const localOffset = view.getUint32(at + 42, true);
    // `uncompressedSize` נמצא כאן אחרי מדידה: בלעדיו ארכיון שהסימן יושב בו
    // דווקא **שם** התקבל, והוא אחד מהגדלים שהכתיבה כאן מחשבת מחדש. שלושת
    // השדות האלה ושניים ב-EOCD הם כל מה שנבדק — לא „כל מקום שהסימן יכול
    // לשבת בו”. ראו את הערת `localExtra`.
    if (
      compressedSize === ZIP64_MARKER_32 ||
      uncompressedSize === ZIP64_MARKER_32 ||
      localOffset === ZIP64_MARKER_32
    ) {
      return null;
    }

    const nameAt = at + ZIP_CENTRAL_HEADER_SIZE;
    // הרשומה חייבת להיכנס שלמה, ולא רק הכותרת שלה. `subarray` שחורג אינו
    // זורק אלא **מקצר בשקט**, והבדיקה שבראש הלולאה תופסת את החריגה רק
    // בסיבוב הבא — כלומר הרשומה האחרונה בארכיון קטוע נקראת ויוצאת שגויה.
    // נמדד על `case.docx` שנחתך: שם הרשומה האחרונה בלע את ה-EOCD ויצא
    // `"customXml/_rels/itemPK\x05\x06…"`. ארכיון קטוע אינו ארכיון שאנחנו
    // מבינים, והתשובה עליו היא `null` — כמו כל כשל אחר כאן.
    if (nameAt + nameLength + extraLength + commentLength > eocd) return null;
    const nameBytes = bytes.subarray(nameAt, nameAt + nameLength);
    const centralExtra = bytes.subarray(nameAt + nameLength, nameAt + nameLength + extraLength);
    const comment = bytes.subarray(
      nameAt + nameLength + extraLength,
      nameAt + nameLength + extraLength + commentLength,
    );
    const data = entryData(bytes, view, localOffset, compressedSize);
    if (!data) return null;
    const localExtra = entryLocalExtra(bytes, view, localOffset);
    if (!localExtra) return null;

    entries.push({
      localOffset,
      entry: {
        name: new TextDecoder().decode(nameBytes),
        nameBytes: nameBytes.slice(),
        versionMadeBy: view.getUint16(at + 4, true),
        versionNeeded: view.getUint16(at + 6, true),
        flags,
        method: view.getUint16(at + 10, true),
        modTime: view.getUint16(at + 12, true),
        modDate: view.getUint16(at + 14, true),
        crc: view.getUint32(at + 16, true),
        internalAttrs: view.getUint16(at + 36, true),
        externalAttrs: view.getUint32(at + 38, true),
        data,
        uncompressedSize,
        localExtra: localExtra.slice(),
        centralExtra: centralExtra.slice(),
        comment: comment.slice(),
      },
    });

    at += ZIP_CENTRAL_HEADER_SIZE + nameLength + extraLength + commentLength;
  }
  return entries.sort((a, b) => a.localOffset - b.localOffset).map((read) => read.entry);
}

/** הבייטים המאוחסנים של רשומה, לפי הכותרת המקומית שלה. */
function entryData(
  bytes: Bytes,
  view: DataView,
  localOffset: number,
  compressedSize: number,
): Bytes | null {
  if (localOffset + ZIP_LOCAL_HEADER_SIZE > bytes.byteLength) return null;
  if (view.getUint32(localOffset, true) !== ZIP_LOCAL_SIGNATURE) return null;

  const nameLength = view.getUint16(localOffset + 26, true);
  const extraLength = view.getUint16(localOffset + 28, true);
  const start = localOffset + ZIP_LOCAL_HEADER_SIZE + nameLength + extraLength;
  if (start + compressedSize > bytes.byteLength) return null;

  return bytes.subarray(start, start + compressedSize);
}

/** שדה ה-extra שבכותרת המקומית של רשומה. */
function entryLocalExtra(bytes: Bytes, view: DataView, localOffset: number): Bytes | null {
  if (localOffset + ZIP_LOCAL_HEADER_SIZE > bytes.byteLength) return null;
  const nameLength = view.getUint16(localOffset + 26, true);
  const extraLength = view.getUint16(localOffset + 28, true);
  const start = localOffset + ZIP_LOCAL_HEADER_SIZE + nameLength;
  if (start + extraLength > bytes.byteLength) return null;
  return bytes.subarray(start, start + extraLength);
}

/**
 * הערת הארכיון שאחרי ה-EOCD, כדי שכתיבה מחדש תחזיר אותה במקומה.
 * ריק — אין הערה, או שהארכיון אינו נקרא.
 */
export function readZipComment(bytes: Bytes): Bytes {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = findEocd(view);
  if (eocd < 0) return new Uint8Array(0) as Bytes;
  const length = view.getUint16(eocd + 20, true);
  const start = eocd + ZIP_EOCD_SIZE;
  if (start + length > bytes.byteLength) return new Uint8Array(0) as Bytes;
  return bytes.subarray(start, start + length).slice();
}

/**
 * מיקום ה-EOCD. נסרק מהסוף, כי לארכיון מותרת הערה בת עד 64KB אחריו.
 */
function findEocd(view: DataView): number {
  const last = view.byteLength - ZIP_EOCD_SIZE;
  const first = Math.max(0, view.byteLength - ZIP_EOCD_SIZE - 0xffff);
  for (let at = last; at >= first; at--) {
    if (view.getUint32(at, true) === ZIP_EOCD_SIGNATURE) return at;
  }
  return -1;
}

/**
 * כתיבת הארכיון מחדש.
 *
 * כל מה שאינו התוכן שתוקן נכתב בחזרה כמות שהוא: הסדר שהתקבל (`readZip` מוסר
 * אותו פיזי), השמות, שיטת הדחיסה, הבייטים, שדות ה-extra של שתי הכותרות, הערת
 * כל רשומה והערת הארכיון. זו הדרישה מפני שהכותב רץ על **כל** מסמך שמיוצא, גם
 * כשהתיקון נוגע בחלק אחד בלבד: מה שאינו נשמר כאן נעלם מכל שמירה.
 *
 * ומה שאין: מסלול „עבר כאן ולא תוקן בו דבר”. גם `rewriteDocxXmlParts` וגם
 * `repairEntries` מחזירים `null` לפני שהם מגיעים לכאן כששום חלק לא השתנה,
 * ולכן הזהות נמדדת על חבילה שתוקן בה חלק. נמדד על `tmp/sp3946/case.docx`:
 * שני חלקים תוקנו, וארבע-עשרה הרשומות האחרות יצאו בייט-בבייט כמו שנכנסו —
 * אותה שיטת דחיסה, אותו CRC, אותם בייטים מאוחסנים.
 */
export function writeZip(entries: ZipEntry[], archiveComment?: Bytes): Bytes {
  const comment = archiveComment ?? (new Uint8Array(0) as Bytes);
  let size = ZIP_EOCD_SIZE + comment.byteLength;
  for (const entry of entries) {
    size += ZIP_LOCAL_HEADER_SIZE + entry.nameBytes.byteLength + entry.data.byteLength;
    size += entry.localExtra?.byteLength ?? 0;
    size += ZIP_CENTRAL_HEADER_SIZE + entry.nameBytes.byteLength;
    size += (entry.centralExtra?.byteLength ?? 0) + (entry.comment?.byteLength ?? 0);
  }

  const out = new Uint8Array(size);
  const view = new DataView(out.buffer);
  const offsets: number[] = [];
  let at = 0;

  for (const entry of entries) {
    offsets.push(at);
    view.setUint32(at, ZIP_LOCAL_SIGNATURE, true);
    view.setUint16(at + 4, entry.versionNeeded, true);
    view.setUint16(at + 6, entry.flags & ~ZIP_FLAG_DATA_DESCRIPTOR, true);
    view.setUint16(at + 8, entry.method, true);
    view.setUint16(at + 10, entry.modTime, true);
    view.setUint16(at + 12, entry.modDate, true);
    view.setUint32(at + 14, entry.crc, true);
    view.setUint32(at + 18, entry.data.byteLength, true);
    view.setUint32(at + 22, entry.uncompressedSize, true);
    view.setUint16(at + 26, entry.nameBytes.byteLength, true);
    view.setUint16(at + 28, entry.localExtra?.byteLength ?? 0, true);
    at += ZIP_LOCAL_HEADER_SIZE;
    out.set(entry.nameBytes, at);
    at += entry.nameBytes.byteLength;
    if (entry.localExtra) {
      out.set(entry.localExtra, at);
      at += entry.localExtra.byteLength;
    }
    out.set(entry.data, at);
    at += entry.data.byteLength;
  }

  const centralOffset = at;
  entries.forEach((entry, index) => {
    view.setUint32(at, ZIP_CENTRAL_SIGNATURE, true);
    view.setUint16(at + 4, entry.versionMadeBy, true);
    view.setUint16(at + 6, entry.versionNeeded, true);
    view.setUint16(at + 8, entry.flags & ~ZIP_FLAG_DATA_DESCRIPTOR, true);
    view.setUint16(at + 10, entry.method, true);
    view.setUint16(at + 12, entry.modTime, true);
    view.setUint16(at + 14, entry.modDate, true);
    view.setUint32(at + 16, entry.crc, true);
    view.setUint32(at + 20, entry.data.byteLength, true);
    view.setUint32(at + 24, entry.uncompressedSize, true);
    view.setUint16(at + 28, entry.nameBytes.byteLength, true);
    view.setUint16(at + 30, entry.centralExtra?.byteLength ?? 0, true);
    view.setUint16(at + 32, entry.comment?.byteLength ?? 0, true);
    view.setUint16(at + 34, 0, true);
    view.setUint16(at + 36, entry.internalAttrs, true);
    view.setUint32(at + 38, entry.externalAttrs, true);
    view.setUint32(at + 42, offsets[index], true);
    at += ZIP_CENTRAL_HEADER_SIZE;
    out.set(entry.nameBytes, at);
    at += entry.nameBytes.byteLength;
    if (entry.centralExtra) {
      out.set(entry.centralExtra, at);
      at += entry.centralExtra.byteLength;
    }
    if (entry.comment) {
      out.set(entry.comment, at);
      at += entry.comment.byteLength;
    }
  });

  view.setUint32(at, ZIP_EOCD_SIGNATURE, true);
  view.setUint16(at + 4, 0, true);
  view.setUint16(at + 6, 0, true);
  view.setUint16(at + 8, entries.length, true);
  view.setUint16(at + 10, entries.length, true);
  view.setUint32(at + 12, at - centralOffset, true);
  view.setUint32(at + 16, centralOffset, true);
  view.setUint16(at + 20, comment.byteLength, true);
  out.set(comment, at + ZIP_EOCD_SIZE);

  return out;
}

let crcTable: Uint32Array | null = null;

/**
 * CRC32 כפי ש-ZIP מגדיר אותו. טבלה אחת, בייט אחר בייט.
 *
 * מיוצאת בשביל הבדיקה בלבד — אין לה קורא אחר מחוץ למודול. מה שהבדיקה שומרת
 * עליו הוא שוויון עם מימוש ייחוס: CRC שגוי הוא ארכיון שבור, וזה כשל שקט.
 *
 * **„slice-by-4” נכתב כאן ונמדד איטי יותר, ולכן הוסר.** מאז שהתיקון השני
 * נכנס החלק שנכתב מחדש עשוי להיות `document.xml` של ספר שלם, ולכן נראה
 * שכדאי. נמדד ב-Node 24, שלוש הרצות, מינימום מתוך חמש חזרות בכל אחת:
 *
 *     גודל     בייט-בבייט     slice-by-4
 *     5.6MB    33–50ms        77–127ms
 *     64KB     0.31–0.63ms    0.81–1.05ms
 *     512B     0.002ms        0.018–0.024ms
 *
 * גם הווריאנט בלי `>>> 0` בתוך הלולאה — כלומר בלי לייצר uint32 שיוצא מטווח
 * ה-Smi בכל איטרציה — נשאר איטי מהפשוט בכל הגדלים. V8 מהדר את הלולאה הצרה
 * הזאת טוב יותר ממה שארבע טבלאות (4KB במקום 1KB) מרוויחות. **לא לכתוב את
 * זה שוב בלי למדוד.**
 */
export function crc32(bytes: Bytes): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let value = i;
      for (let bit = 0; bit < 8; bit++) {
        value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
      }
      crcTable[i] = value >>> 0;
    }
  }

  let crc = 0xffffffff;
  for (let i = 0; i < bytes.byteLength; i++) {
    crc = crcTable[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
