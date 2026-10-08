/**
 * ייצוא המסמך הפעיל, ובחירת השם והסיומת שהוא נשמר תחתם.
 *
 * `triggerDownload: false` — התוסף מקבל את ה-Blob ומחליט מה לעשות בו.
 * ההורדה האוטומטית של SuperDoc אינה מסלול שמירה אמין, ובעיקר אינה בסיס
 * ל-autosave.
 *
 * ## למה הסיומת יושבת כאן, ולמה היא לא תמיד `docx`
 *
 * מסמך עם מאקרו (`.docm`) הוא אותה חבילת OOXML בדיוק, ובתוכה חלק נוסף —
 * `word/vbaProject.bin`. המנוע **שומר** אותו בייט-בבייט בייצוא, כלומר הבייטים
 * שיוצאים מכאן עשויים להיות מסמך עם מאקרו גם כשלא נגענו במאקרו.
 *
 * וזאת בדיוק המלכודת: Word מסרב לפתוח חבילה שיש בה `vbaProject` אם שמה מסתיים
 * ב-`.docx`. כתיבה כ-`docx` הייתה הופכת מסמך תקין של המשתמש לקובץ שהוא מקבל
 * עליו אזהרה — או שהמאקרו שלו פשוט מפסיקים לעבוד, בלי ששום דבר אמר לו למה.
 * לכן הסיומת אינה קבועה אלא נגזרת משני דברים: מה הייתה סיומת המקור, והאם
 * החבילה נושאת חלק מאקרו בפועל.
 */
import type { SuperDoc } from 'superdoc';
import type { Bytes } from './docx-parts';
import { postflightDocx } from './docx-postflight';

/** הסיומות של חבילות OOXML לעיבוד תמלילים שהתוסף מכיר. */
export const WORD_EXTENSIONS = ['docx', 'docm', 'dotx', 'dotm'] as const;

export type WordExtension = (typeof WORD_EXTENSIONS)[number];

export const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/** ה-MIME של כל סיומת. `docm`/`dotm` הם טיפוסים נפרדים, לא וריאנט של `docx`. */
export const MIME_FOR_EXTENSION: Readonly<Record<WordExtension, string>> = {
  docx: DOCX_MIME,
  docm: 'application/vnd.ms-word.document.macroEnabled.12',
  dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
  dotm: 'application/vnd.ms-word.template.macroEnabled.12',
};

/** הסיומת המקבילה עם מאקרו. `docm`/`dotm` הן כבר כאלה. */
const MACRO_ENABLED: Readonly<Record<WordExtension, WordExtension>> = {
  docx: 'docm',
  dotx: 'dotm',
  docm: 'docm',
  dotm: 'dotm',
};

const EXTENSION_PATTERN = new RegExp(`\\.(${WORD_EXTENSIONS.join('|')})$`, 'i');

/**
 * ייצוא המסמך. `exportType: ['docx']` הוא משפחת הפורמט של המנוע ואינו נוגע
 * לסיומת: חבילה עם מאקרו יוצאת מכאן עם המאקרו שלה בפנים.
 *
 * זו נקודת החנק היחידה של הכתיבה — שמירה, „שמור בשם” והחלפת הקובץ כולן עוברות
 * כאן — ולכן זה גם המקום שבו הבייטים עוברים את התיקונים של הדרך החוצה (סימון
 * התו הניטרלי הסוגר, מזהי מספור) לפני שהם נכתבים. ראו engine/docx-postflight.ts.
 */
export async function exportDocx(superdoc: SuperDoc, options: ExportDocxOptions = {}): Promise<Blob> {
  const blob = await superdoc.export({ exportType: ['docx'], triggerDownload: false });
  if (!(blob instanceof Blob)) throw new Error('הייצוא לא החזיר קובץ');
  return options.postflight === false ? blob : applyPostflight(blob);
}

export interface ExportDocxOptions {
  /**
   * `false` — הבייטים כפי שהמנוע כתב אותם. לטיוטת השחזור בלבד: היא נטענת
   * חזרה **לעורך**, שאינו צריך את התיקונים של Word, והיא נכתבת עשר שניות
   * אחרי כל שינוי.
   *
   * נמדד על `postflightDocx` **מקצה לקצה** — כלומר כולל את שכבת החבילה:
   * פריסת ה-zip, סריקת החלקים, הדחיסה מחדש והפריסה שמאמתת אותה.
   *
   * **והקלט הוא ארכיון דחוס**, כמו שהמנוע כותב אותו. זה אינו פרט טכני:
   * הטבלה הקודמת נמדדה על ארכיון `STORED`, שאין בו פריסה בכניסה, בעוד
   * המשפט שמעליה הבטיח שהמספר כולל אותה. הפריסה הנכנסת נמדדה בנפרד, והיא
   * 42ms ל-8.4MB מול 11ms לאותו חלק לא-דחוס.
   *
   * מסמך עברי סינתטי שנכתב לתוך חבילת Word אמיתית, Node 24, מכונה של ארבע
   * ליבות ו-16GB **שרצים עליה דברים אחרים**; מינימום מתוך תשע הרצות:
   *
   *     פסקאות   document.xml   מקצה לקצה
   *      1,000      0.09MB          17ms
   *     10,000      0.87MB          89ms
   *     40,000      3.50MB         346ms
   *
   * כתיבת הארכיון עצמה זניחה, והשאר מתחלק בין סריקת
   * `docx-neutral-mark.ts` (‏6 / 40 / 171ms — ליניארית במסמך), ה-deflate,
   * הפריסה שמאמתת אותו ו**שלוש** מעבירות crc32: `deflateVerified` עושה
   * שתיים (על המשוחזר ועל המקור) ו-`rewriteEntry` שלישית, ביחד 3.04× גודל
   * החלק לפי ספירה, בקצב של ‎27ms למעבירה.
   *
   * (הטבלה נמדדה מחדש אחרי שבדיקת הטווח ב-`docx-neutral-mark.ts` תוקנה
   * מריבועית לליניארית — לפניה אותה סריקה לקחה 2,229ms ב-40,000.)
   *
   * לצורך ההחלטה שהיא מתעדת אין בכך הבדל: גם 89ms בטיוטה שנכתבת כל עשר
   * שניות הם עצירה של ההקלדה, ולכן `false`.
   *
   * **ונמדד מחדש אחרי `docx-run-direction.ts` (8.10.2026)**, על המקרה הגרוע
   * שלו: 20,000 פסקאות שבכל אחת ריצה מעורבת מודגשת (3.2MB) — כלומר כל פסקה
   * מפוצלת, מסומנת וממורה. מקצה לקצה, שלוש הרצות, ובמכונה **ב-100% עומס**
   * של סשן אחר: לפני 0.38–0.51 שניות, אחרי 1.8–2.9. לפי שלב (מינימום מחמש,
   * בלי שכבת ה-zip): RLM ‏150ms, הסימון 530–660ms, המראה 340ms — שהייתה
   * מדלגת על מסמך בלי `w:rtl` ועכשיו רצה עליו. זה מחזק את ההחלטה כאן.
   */
  postflight?: boolean;
}

/**
 * הבייטים של ה-Blob, או `null` כשאי אפשר לקרוא אותם.
 *
 * `arrayBuffer` נבדק ולא מונח: ב-jsdom הוא אינו קיים, ובלי הבדיקה כל בדיקת
 * יחידה שנוגעת בייצוא הייתה זורקת `TypeError` במקום לרוץ.
 */
async function blobBytes(blob: Blob): Promise<Bytes | null> {
  if (typeof blob.arrayBuffer !== 'function') return null;
  return new Uint8Array(await blob.arrayBuffer()) as Bytes;
}

/**
 * התיקונים של הדרך החוצה על הבייטים שיוצאים.
 *
 * אותו כלל כמו בכיוון הנכנס (`docx-preflight.ts`): **לתקן, ולא לחסום.** כל כשל
 * — Blob שאינו נקרא, zip שאינו נפרס, דוחס שאינו קיים — מחזיר את המקור כמות
 * שהוא. נקודה בצד הלא נכון היא באג; שמירה שנכשלת היא אובדן עבודה.
 */
async function applyPostflight(blob: Blob): Promise<Blob> {
  try {
    const bytes = await blobBytes(blob);
    const marked = bytes && (await postflightDocx(bytes));
    return marked ? new Blob([marked], { type: blob.type || DOCX_MIME }) : blob;
  } catch (error) {
    // **ומיומן.** בלי השורה הזאת, מקרה קצה שמפיל את התיקון על מסמך אחד
    // מוריד אותו מ**כל** שמירה של המשתמש ההוא בשקט: הסימפטום שדווח חוזר,
    // ואין שום שורה ביומן להתחיל ממנה — בעוד שער ה-QA, שרץ על קובץ משלו,
    // נשאר ירוק. כל כשל אחר בשכבה הזאת מיומן (docx-parts.ts,
    // docx-preflight.ts), וזה היה היחיד שלא.
    console.warn('[otzaria-word] התיקון של הדרך החוצה נכשל, והמסמך נשמר כמות שהוא', error);
    return blob;
  }
}

/** הסיומת שבשם הקובץ, או `null` כשאינה אחת מהמוכרות. */
export function extensionFromFileName(name: string | undefined): WordExtension | null {
  const match = name ? EXTENSION_PATTERN.exec(name.trim()) : null;
  return match ? (match[1]!.toLowerCase() as WordExtension) : null;
}

/** מסיר סיומת מוכרת משם קובץ. שם בלי סיומת מוכרת חוזר כמות שהוא. */
export function stripWordExtension(name: string): string {
  return name.replace(EXTENSION_PATTERN, '');
}

/**
 * הסיומת שתחתה יש לשמור.
 *
 * שומרת על סיומת המקור — משתמש שפתח תבנית מצפה לשמור תבנית — ומשדרגת לגרסת
 * המאקרו כשהחבילה נושאת חלק מאקרו. השדרוג הוא המסלול הבטוח: קובץ עם
 * `vbaProject` שנשמר כ-`docx` הוא קובץ ש-Word מתלונן עליו.
 *
 * הכיוון ההפוך לא נעשה בכוונה: מסמך `.docm` שאין בו (עוד) מאקרו נשמר כ-`.docm`,
 * מפני שזה מה שהמשתמש בחר — Word עצמו מתנהג כך.
 *
 * ## איפה הסיומת הזאת קובעת, ואיפה לא
 *
 * ב„שמור בשם”, בשם שמוצע ובסינון של הדיאלוג (`extension` ב-
 * `commitUserFileWrite`). **לא** בכתיבה במקום: שם
 * המאחז כותב ל-`targetToken` — הנתיב שממנו הקובץ נפתח — ומתעלם מהשם
 * המוצע, וזו ההתנהגות הנכונה: הקובץ נשאר היכן שהמשתמש שם אותו. מסמך
 * `.docx` שכבר נושא `vbaProject` נשמר לכן חזרה לאותו `.docx`; זה מצבו
 * מלפני שנפתח כאן, והעורך אינו „מתקן” אותו בלי שנתבקש.
 */
export function resolveSaveExtension(
  sourceName: string | undefined,
  hasMacros: boolean,
): WordExtension {
  const base = extensionFromFileName(sourceName) ?? 'docx';
  return hasMacros ? MACRO_ENABLED[base] : base;
}

/**
 * שם הקובץ לשמירה: מנקה תווים שאינם חוקיים בשם קובץ ב-Windows, מסיר סיומת
 * מוכרת אם יש, ומצמיד את הסיומת המבוקשת.
 */
export function documentFileName(title: string, extension: WordExtension = 'docx'): string {
  const clean = stripWordExtension(title.replace(/[\\/:*?"<>|]/g, '').trim()) || 'מסמך';
  return `${clean}.${extension}`;
}

/**
 * אותם בייטים עם ה-MIME הנכון לסיומת.
 *
 * `superdoc.export` מסמן את ה-Blob כ-`docx` תמיד, וה-MIME הזה הוא מה שנשלח
 * כ-`Content-Type` בכתיבה. מסמך עם מאקרו שנכתב כ-`docx` היה מוצהר לא נכון
 * למי שקורא את ההצהרה.
 */
export function retypeBlob(blob: Blob, extension: WordExtension): Blob {
  const type = MIME_FOR_EXTENSION[extension];
  // ההשוואה חסרת-רישיות מפני ש-`Blob` מנרמל את הטיפוס לאותיות קטנות לפי
  // התקן: `blob.type === type` על טיפוס עם אות גדולה (`macroEnabled`) לעולם
  // אינו אמת, וכל קריאה הייתה מייצרת Blob חדש לחינם.
  return blob.type.toLowerCase() === type.toLowerCase() ? blob : new Blob([blob], { type });
}
