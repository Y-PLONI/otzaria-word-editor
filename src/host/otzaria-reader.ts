/**
 * שילוב עם הקורא של אוצריא: ניווט לספרייה, פתיחת חיפוש, והבחירה בטאב הטקסט.
 *
 * למה מודול נפרד ולא קריאות מתוך `App.vue`: שלושת ה-RPC כאן חולקים שני דברים
 * שאין להם מקום בקומפוננטה — מיפוי המתודה להרשאה שהיא דורשת (`reader.open`,
 * `navigation.write`), והתשובה לשאלה „מה אומרים למשתמש כשההרשאה חסרה”. בלי
 * המיפוי הזה כשל הרשאה מגיע כהודעה של אוצריא באנגלית, או נבלע — וזה בדיוק
 * הכשל שהמודול הזה נכתב כדי למנוע: שלושת הכפתורים בלשונית „אוצריא” הציגו
 * הודעת סטטוס שמתארת פעולה שלא קרתה.
 *
 * ההרשאות מוצהרות ב-`public/manifest.json`, ו-tests/unit/manifest.test.ts
 * מקבע שמה שנצרך כאן אכן מוצהר שם.
 */
import type { SuperDoc } from 'superdoc';
import { call, isPermissionDenied, hostErrorCode } from './otzaria-client';
import { readDocSelection, type SelectionDocumentApi } from '../engine/doc-selection';
import {
  readDocCapabilities,
  type CapabilitiesDocumentApi,
} from '../engine/doc-capabilities';
import {
  receiptFailureText,
  thrownText,
  type DocReceipt,
  type MaybePromise,
} from '../engine/document-api';
import type {
  GetSectionTextMapArgs,
  NavigationTarget,
  OpenSearchTabArgs,
  ReaderRefState,
  ReaderSelection,
  SectionTextMapResult,
} from '../types/otzaria_plugin';

/**
 * תוצאת פעולה מול הקורא. מטופסת ולא זריקה, מאותו טעם כמו `ImageDataUrlResult`
 * ב-host/files.ts: הקורא בממשק צריך הודעה אחת בעברית ולא שלושה מסלולי טיפול.
 */
export type ReaderResult<T = void> =
  | { ok: true; value: T }
  | { ok: false; message: string; reason: string };

/**
 * ההרשאה שכל מתודה דורשת, לפי docs/plugin-sdk/API_REFERENCE.md. המפה היא גם
 * המקור לבדיקת המניפסט — כך „הכפתור נכשל בהרשאה” אינו יכול להגיע למשתמש
 * בגלל הצהרה שנשכחה.
 */
export const READER_PERMISSIONS: Record<string, string> = {
  'reader.getSelection': 'reader.open',
  'reader.openSearchTab': 'reader.open',
  'reader.getCurrentState': 'reader.open',
  'reader.getSectionTextMap': 'reader.open',
  'navigation.goTo': 'navigation.write',
};

/**
 * מנרמלת טקסט מסומן לשורה אחת.
 *
 * גם הבחירה במסמך וגם הבחירה בקורא מגיעות עם שברי שורות ורווחים כפולים —
 * פסקה בקורא היא שורה בקובץ המקור, ובחירה בשתי פסקאות מביאה את שתיהן. שאילתת
 * חיפוש עם שבר שורה בתוכה אינה מה שהמשתמש סימן, וציטוט שנכנס למסמך צריך
 * להיות פסקה אחת. הנרמול הוא של רווחים בלבד: ניקוד, טעמים וסימני פיסוק הם
 * חלק מהטקסט המצוטט ואין לגעת בהם.
 */
export function normalizeSelectedText(text: string): string {
  return typeof text === 'string' ? text.replace(/\s+/g, ' ').trim() : '';
}

/**
 * הודעה בעברית לכשל שחזר מאוצריא.
 *
 * הרשאה חסרה היא המקרה שחוזר בשטח, והיא צריכה לומר **מה** חסר: המשתמש יכול
 * לאשר הרשאה בהגדרות התוסף, אבל לא אם ההודעה אומרת „הפעולה נכשלה”. שאר
 * הקודים נמסרים כפי שהם לצד ההודעה של אוצריא — הם מיועדים לדיווח באג, לא
 * להוראה למשתמש.
 */
function hostFailure(
  method: string,
  failedAction: string,
  error: unknown,
): { ok: false; message: string; reason: string } {
  if (isPermissionDenied(error)) {
    const permission = READER_PERMISSIONS[method] ?? method;
    return {
      ok: false,
      reason: 'permission-denied',
      message: `${failedAction}: לתוסף חסרה ההרשאה „${permission}”. יש לאשר אותה לתוסף בהגדרות אוצריא ולטעון את הלשונית מחדש`,
    };
  }
  const code = hostErrorCode(error);
  const detail = error instanceof Error ? error.message : String(error);
  return {
    ok: false,
    reason: code ?? 'threw',
    message: code ? `${failedAction}: ${detail} (${code})` : `${failedAction}: ${detail}`,
  };
}

/**
 * קוראת ל-RPC שהתשובה שלו היא „בוצע” בוליאני (`navigation.goTo`,
 * `reader.openSearchTab` — שניהם מתועדים כמחזירים `true`).
 *
 * `false` הוא סירוב מפורש של אוצריא ולכן הוא כשל. כל צורה אחרת — `null`,
 * אובייקט, `undefined` — נחשבת הצלחה ונרשמת ללוג בלבד: התוצאה הנראית של שתי
 * הפעולות היא מסך שמתחלף, המשתמש רואה בעצמו אם זה קרה, והודעת שגיאה על גרסה
 * שהחזירה צורה אחרת היא אזעקת שקר. זה גם ההסבר ל-stub של הפיתוח, שמחזיר
 * `null` לכל מתודה שאינה ממומשת בו.
 */
async function callAck(
  method: string,
  failedAction: string,
  payload: Record<string, unknown>,
): Promise<ReaderResult> {
  let data: unknown;
  try {
    data = await call<unknown>(method, payload);
  } catch (error) {
    return hostFailure(method, failedAction, error);
  }

  if (data === false) {
    return { ok: false, reason: 'refused', message: `${failedAction}: אוצריא לא ביצעה את הפעולה` };
  }
  if (data !== true) {
    console.warn(`[otzaria-word] ${method} החזירה תשובה בצורה לא צפויה`, data);
  }
  return { ok: true, value: undefined };
}

/** מעבר למסך ראשי באוצריא. */
export function goTo(target: NavigationTarget): Promise<ReaderResult> {
  return callAck('navigation.goTo', 'המעבר באוצריא נכשל', { target });
}

/**
 * פותחת את מסך הספרייה של אוצריא.
 *
 * הפעולה מוציאה את המשתמש מלשונית התוסף — זה מה ש„פתח ספרייה” אומר, והמסמך
 * נשאר פתוח בלשונית שלו. השמירה אינה מופעלת כאן: היא כבר אוטומטית, ולכפות
 * שמירה על ניווט היה הופך כפתור ניווט לכפתור שכותב לדיסק.
 */
export function openLibrary(): Promise<ReaderResult> {
  return goTo('library');
}

/**
 * פותחת את מסך החיפוש הרגיל של אוצריא עם השאילתה.
 *
 * `autoSearch` נשאר בברירת המחדל (`true`): המשתמש סימן טקסט וביקש לחפש אותו,
 * ולפתוח לו את המסך בלי להריץ היה מבקש ממנו ללחוץ Enter על מה שהוא כבר בחר.
 */
export function openSearchTab(args: OpenSearchTabArgs): Promise<ReaderResult> {
  const { query, ...rest } = args;
  return callAck('reader.openSearchTab', 'פתיחת החיפוש באוצריא נכשלה', { query, ...rest });
}

/**
 * הבחירה הנוכחית בטאב הטקסט של הקורא.
 *
 * `null` הוא תשובה תקינה ולא כשל: אין בחירה, או שהטאב הפעיל אינו טאב טקסט
 * (PDF, למשל). הקורא בממשק מבקש מהמשתמש לסמן, ואינו מציג שגיאה.
 *
 * תשובה שאינה אובייקט מתורגמת גם היא ל-`null` — גרסת מארח שתחזיר מחרוזת או
 * מספר אינה סיבה לזרוק מתוך לחיצה על כפתור.
 */
export async function getReaderSelection(): Promise<ReaderResult<ReaderSelection | null>> {
  let data: unknown;
  try {
    data = await call<unknown>('reader.getSelection');
  } catch (error) {
    return hostFailure('reader.getSelection', 'קריאת הבחירה מאוצריא נכשלה', error);
  }

  if (!data || typeof data !== 'object') return { ok: true, value: null };
  return { ok: true, value: data as ReaderSelection };
}

/**
 * מצב הקריאה הנוכחי: ספר, מיקום (`currentRef`) וטאב פעיל. `null` תקין —
 * אין ספר פתוח (מסך ספרייה, למשל), וזה לא כשל.
 */
export async function getCurrentReaderState(): Promise<ReaderResult<ReaderRefState | null>> {
  let data: unknown;
  try {
    data = await call<unknown>('reader.getCurrentState');
  } catch (error) {
    return hostFailure('reader.getCurrentState', 'קריאת מצב הקורא נכשלה', error);
  }

  if (!data || typeof data !== 'object') return { ok: true, value: null };
  return { ok: true, value: data as ReaderRefState };
}

/**
 * טקסט ה-section הנוכחי בספר, לצורך "השלמה מהספר" — ראו
 * engine/book-completion.ts. `layer: 'source'` בלבד ובלי `includeWords`:
 * הנרמול וחלוקת המילים נעשים מקומית (book-completion.ts), וזה גם מה שחוסך
 * את משיכת שני שכבות הטקסט ואת מערך ה-tokens לכל section.
 */
export async function getSectionTextMap(
  args: GetSectionTextMapArgs,
): Promise<ReaderResult<SectionTextMapResult | null>> {
  let data: unknown;
  try {
    data = await call<unknown>('reader.getSectionTextMap', { ...args });
  } catch (error) {
    return hostFailure('reader.getSectionTextMap', 'קריאת טקסט הספר נכשלה', error);
  }

  if (!data || typeof data !== 'object') return { ok: true, value: null };
  return { ok: true, value: data as SectionTextMapResult };
}

/** הראשון מבין המועמדים שיש בו טקסט. `''` אינו „קיים” — הוא בחירה ריקה. */
function firstText(...candidates: Array<string | null | undefined>): string {
  for (const candidate of candidates) {
    if (typeof candidate === 'string' && candidate.trim() !== '') return candidate;
  }
  return '';
}

/**
 * הטקסט שמצטטים.
 *
 * `sourceSelectedText` לפני `renderedSelectedText`, כי לציטוט תורני נדרש
 * **טקסט המקור**: מה שהוצג בקורא תלוי בהגדרות התצוגה של מי שסימן — ניקוד
 * וטעמים מוסרים שם לפי העדפה — ולכן אותה בחירה בדיוק הייתה מייצרת ציטוט אחר
 * אצל שני משתמשים. הציטוט צריך לשקף את הספר, לא את המסך.
 *
 * `text` (השדה הוותיק) הוא הגיבוי היחיד: העוגן המורחב קיים מ-0.9.95 בלבד,
 * ובגרסה ישנה יותר הוא כל מה שיש.
 *
 * בחירה שפרושה על כמה פסקאות מגיעה ב-`sections` (מ-0.9.97), והשדות ברמה
 * העליונה אינם נושאים אותה במלואה. הפסקאות מחוברות ברווח ולא בשבר שורה:
 * `insert` עם `type: 'text'` מכניס טקסט לפסקה, ולא נכון להעמיס עליו פירוק
 * לפסקאות שהוא לא הובטח לעשות.
 */
function citationSourceText(selection: ReaderSelection | null | undefined): string {
  if (!selection || typeof selection !== 'object') return '';

  const sections = Array.isArray(selection.sections) ? selection.sections : [];
  if (sections.length > 0) {
    return sections
      .map((section) => firstText(section?.sourceSelectedText, section?.renderedSelectedText))
      .filter((text) => text !== '')
      .join(' ');
  }

  return firstText(
    selection.sourceSelectedText,
    selection.renderedSelectedText,
    selection.text,
  );
}

/**
 * המלל שנכנס למסמך: הטקסט המצוטט, ואחריו המקור בסוגריים —
 * `וַיֹּאמֶר אֱלֹהִים (בראשית פרק א)`.
 *
 * מצב אחד ולא שלושה: §14.1 בתכנית מבקשת „טקסט בלבד, טקסט + מקור, וציטוט
 * מעוצב”, ושלושת המצבים דורשים מקום בממשק שיבחר ביניהם (תפריט על הכפתור).
 * המצב שנבחר הוא זה שנכון כברירת מחדל לכתיבת חידושים — ציטוט בלי מקור אינו
 * ציטוט — והשניים האחרים הם המשך, לא חוב.
 *
 * בלי `currentRef` נכנס הטקסט לבדו: סוגריים ריקים גרועים מציטוט בלי מקור.
 * מחרוזת ריקה פירושה „אין מה להכניס”, והקורא בממשק הוא שמחליט מה לומר.
 */
export function buildCitationText(selection: ReaderSelection | null | undefined): string {
  const text = normalizeSelectedText(citationSourceText(selection));
  if (!text) return '';

  const sections = Array.isArray(selection?.sections) ? selection.sections : [];
  const ref = normalizeSelectedText(
    firstText(selection?.currentRef, sections[0]?.currentRef),
  );
  return ref ? `${text} (${ref})` : text;
}

/**
 * הצורה שנצרכת מ-`doc`: ההכנסה עצמה, הבחירה שקובעת לאן, והיכולות שנשאלות
 * לפניה. מוגדרת כאן ולא מיובאת מהמנוע — ההסבר המלא ב-engine/document-api.ts.
 * `CapabilitiesDocumentApi` מורש ולא משוכפל: בלעדיו TypeScript דוחה את המסירה
 * ל-`readDocCapabilities` כטיפוס חלש שאין לו שדה משותף.
 */
interface CitationDocumentApi extends SelectionDocumentApi, CapabilitiesDocumentApi {
  insert?: (input: {
    value: string;
    type: 'text';
    target?: unknown;
  }) => MaybePromise<DocReceipt>;
}

export interface CitationHost {
  activeEditor?: { doc?: CitationDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-engine/page-setup.ts. */
export type CitationTarget = SuperDoc | CitationHost | null | undefined;

/** לאן הציטוט נכנס בפועל. הקורא בממשק אומר את זה למשתמש. */
export type CitationPlacement = 'at-cursor' | 'document-end';

/**
 * האם יש למי להכניס ציטוט. נבדק בפקד, כדי שלא ייראה פעיל בלי מסמך.
 *
 * שתי בדיקות ולא אחת, ולא בכפל: זמינות הפעולה נשאלת דרך מרחב השאלות המשותף
 * (`canInsertText` ב-engine/doc-capabilities.ts), כדי שלא תהיה כאן תשובה שנייה
 * לאותה שאלה — אבל **נוכחות** `doc.insert` נבדקת לפניה. מפת ה-`operations` של
 * המנוע נבנית מקטלוג הפעולות, ולכן גרסה שהסירה את המימוש ועודה מכריזה על
 * הפעולה בקטלוג הייתה מחזירה „זמין” לפקד שאין לו למה לקרוא. אותו לקח בדיוק
 * כמו ב-engine/page-break.ts, ושם בדיקה תפסה את הסרתה.
 *
 * א-סינכרונית מפני ש-`capabilities.get()` א-סינכרוני, ולעולם אינה זורקת —
 * נכשלת סגור, כי „אולי כן” הוא בדיוק הכפתור המת.
 */
export async function canInsertText(host: CitationTarget): Promise<boolean> {
  const insert = (host as CitationHost | null | undefined)?.activeEditor?.doc?.insert;
  if (typeof insert !== 'function') return false;

  const report = await readDocCapabilities(host);
  return report.can('canInsertText');
}

/**
 * מכניסה את הציטוט למסמך דרך ה-Document API הציבורי.
 *
 * **המיקום**: החוזה של `insert` קובע ש„בלי `target` ההכנסה נעשית בסוף
 * המסמך”, וזה כמעט תמיד לא מה שהמשתמש התכוון אליו. לכן היעד נלקח מתצלום
 * הבחירה במסמך (engine/doc-selection.ts) — אותו מסלול שדיאלוג הקישור משתמש
 * בו, ומאותו טעם: לחיצה על פקד ברצועה גוזלת את המיקוד מהעורך, והיעד צריך
 * להיתפס ולהימסר במפורש. כשאין תצלום — למשל מסמך שנפתח ואיש לא הקליק בו —
 * ההכנסה נעשית בסוף המסמך, וההודעה למשתמש אומרת את זה במקום להשתיק.
 *
 * **RTL**: אין קביעת כיווניות כאן, ובכוונה. מסמך חדש נפתח עם `w:bidi`
 * ב-`docDefaults/w:pPrDefault` (engine/document-defaults.ts), וכל פסקה שנוצרת
 * בו יורשת אותו; הכנסה לתוך פסקה קיימת יורשת את הכיווניות שלה. מסמך שנפתח
 * מקובץ נושא את הכיווניות של מי שכתב אותו — לכפות עליה RTL בגלל ציטוט היה
 * שינוי בעיצוב המסמך של מישהו אחר, ולא הכנסת טקסט.
 *
 * לעולם אינה זורקת: פעולות ה-Document API זורקות `INVALID_INPUT` על קלט פסול
 * במקום להחזיר קבלה, וחריגה מתוך פקד ברצועה מפילה את רינדור הרצועה כולה.
 */
export async function insertCitation(
  host: CitationTarget,
  text: string,
): Promise<ReaderResult<CitationPlacement>> {
  const failedAction = 'הכנסת הציטוט נכשלה';
  const insert = (host as CitationHost | null | undefined)?.activeEditor?.doc?.insert;

  if (typeof insert !== 'function') {
    // אותו נוסח שהיכולות מחזירות (§12), כדי שהמשתמש יראה את אותו הסבר בין אם
    // הפקד מנוטרל ובין אם הוא נלחץ לפני שהמסמך סיים להיטען.
    return { ok: false, reason: 'command-unsupported', message: `${failedAction}: אינו זמין בגרסה זו` };
  }
  if (!text) {
    return { ok: false, reason: 'empty-text', message: `${failedAction}: אין טקסט להכנסה` };
  }

  const snapshot = await readDocSelection(host);
  // `selectionTarget` ולא `target`: `insert` מקבל **רק** את הצורה הזו, ומסירת
  // רשימת הקטעים נכשלה סגור עם `target must be a SelectionTarget object.` —
  // המשתמש ראה הודעת שגיאה ושום דבר לא נכתב. ההבחנה מתועדת ב-doc-selection.ts.
  const placement: CitationPlacement = snapshot.selectionTarget ? 'at-cursor' : 'document-end';

  let receipt: DocReceipt;
  try {
    receipt = await insert({
      value: text,
      type: 'text',
      ...(snapshot.selectionTarget ? { target: snapshot.selectionTarget } : {}),
    });
  } catch (error) {
    return { ok: false, reason: 'threw', message: thrownText(failedAction, error) };
  }

  if (receipt?.success === false) {
    return {
      ok: false,
      reason: receipt.failure?.code ?? 'receipt-failed',
      message: receiptFailureText(failedAction, receipt),
    };
  }

  return { ok: true, value: placement };
}
