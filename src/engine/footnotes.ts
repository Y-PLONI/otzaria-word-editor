/**
 * הערת שוליים והערת סיום, דרך `doc.footnotes.insert`.
 *
 * **`footnotes` הוא adapter אופציונלי בחוזה** (`footnotes?:` ב-
 * `DocumentApiAdapters`). כשהוא חסר, `doc.capabilities.get()` מסמן את כל
 * הפעולות של ה-namespace כ-`available: false` עם `NAMESPACE_UNAVAILABLE` —
 * ולכן בדיקת היכולת לבדה מספיקה כדי להחליט אם הפקד פעיל, ואין צורך בניחוש.
 * זה גם מה ש-§12 דורש: „פקד שאין לו API ציבורי אמין מסומן „לא זמין בגרסה
 * זו”; לא מממשים אותו דרך XML ידני או DOM פנימי.”
 *
 * `at` אינו נשלח: החוזה קובע שבהיעדרו ההוספה נעשית במקום הסמן — „the
 * toolbar/default editor path” — וזו בדיוק ההתנהגות שפקד בסרגל צריך. שליחת
 * `at` הייתה מחייבת אותנו לחשב כתובת טקסט מהבחירה, כלומר לשחזר בקוד שלנו את
 * מה שהמנוע כבר עושה.
 *
 * `content: ''` ולא טקסט מקום: Word מוסיף הערה ריקה ומעביר אליה את הסמן, ואין
 * שום טקסט שנכון לשתול במסמך של מישהו אחר. הוולידציה במנוע דורשת מחרוזת, לא
 * מחרוזת לא ריקה.
 *
 * ## גל 9 — מה נמדד לפני שנכתבה כאן עוד שורה
 *
 * Chrome headless, `file://`, ה-dist הארוז, חמישה סבבים, וכל אחד מהם פורק גם
 * ברמת ה-zip של `export.toDocx`. שלוש תשובות, וכל אחת שינתה את מה שנשלח:
 *
 * ### מה שנכתב נכון, ולכן נשלח
 *
 * ה-docx שיוצא הוא Word קנוני, אחד לאחד:
 *
 *     document.xml:  <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
 *                      <w:footnoteReference w:id="1"/></w:r>
 *     footnotes.xml: <w:footnote w:type="separator" w:id="-1">…
 *                    <w:footnote w:id="1"><w:p><w:pPr>
 *                      <w:pStyle w:val="FootnoteText"/></w:pPr>
 *                      <w:r><w:t>הערה</w:t></w:r></w:p></w:footnote>
 *
 * עם `separator` ו-`continuationSeparator` במקומם, וסגנונות `FootnoteText`
 * ו-`FootnoteReference` האמיתיים של Word. הערות סיום מקבלות את אותו טיפול
 * ב-`endnotes.xml`. **העברית עוברת שלמה**, כולל ניקוד וגרשיים:
 * „רַשִׁ״י בְּרֵאשִׁית א׳ א׳, ועיין ב"שו״ת הרמב״ם" סי׳ ק״י” חזר תו-בתו
 * מ-`get` ונכתב תו-בתו ל-`<w:t>`.
 *
 * ### `footnotes.update` — **מחליף**, ולא מוסיף
 *
 * זו הייתה השאלה הראשונה, מפני שב-`captions.update` של הגל הקודם התשובה
 * הייתה הפוכה והפילה את הפעולה כולה. כאן שלושה צעדים רצופים על אותה הערה:
 *
 *     insert 'הערה ראשונה'  → get 'הערה ראשונה'
 *     update 'הערה שנייה'   → get 'הערה שנייה'
 *     update 'הערה שלישית'  → get 'הערה שלישית'
 *
 * כלומר עריכה כאן היא קריאה אחת, בלי `remove`+`insert`, בלי בחירת עוגן ובלי
 * רשת שחזור. `patch: { content: '' }` **מוחק** את התוכן (ולא מוסיף מפריד
 * ריק), ו-`patch: {}` או שדה שאינו בחוזה מוחזרים `success: true` בלי לגעת
 * בכלום ובלי `NO_OP` — ולכן הוולידציה על התוכן יושבת כאן, לפני הקריאה.
 *
 * ### הכתובת אינה יודעת להבדיל בין הערת שוליים להערת סיום
 *
 * זה הממצא שקובע את צורת המודול. `FootnoteAddress` הוא
 * `{ kind:'entity', entityType:'footnote', noteId }` — **`entityType` הוא
 * תמיד `'footnote'`, גם עבור הערת סיום**, ושני הרצפים מתחילים מ-1 בנפרד.
 * במסמך שיש בו הערת שוליים 1 והערת סיום 1 שתיהן נושאות את אותה כתובת
 * בדיוק, ואת אותו `handle.ref` (`footnote:1`). מה שנמדד:
 *
 * - `get`/`update`/`remove` על הכתובת הזאת פוגעים ב**הערת השוליים**.
 * - אחרי שהערת השוליים הוסרה, אותה כתובת פוגעת בהערת הסיום.
 * - `entityType: 'endnote'` נזרק („target must be a FootnoteAddress …
 *   entityType 'footnote'”).
 *
 * כלומר לחיצה על „הסר” בשורה של הערת סיום הייתה מוחקת הערת שוליים אחרת,
 * עם `success: true`. **נעקף:** לפני כל עריכה והסרה נקרא `get` על הכתובת,
 * ואם הסוג שחזר אינו הסוג שהמשתמש בחר — הפעולה מסרבת **לפני** שנגעה
 * במסמך, ומסבירה שהערת הסיום אינה ניתנת לזיהוי כל עוד קיימת הערת שוליים
 * באותו מספר. זה גם המבחן היחיד שאפשר לסמוך עליו: הוא שואל את המנוע עצמו
 * למה הכתובת נפתרת, ולא מנחש מרשימה.
 *
 * ### `footnotes.configure` — לא נשלח, ולא מפני שהוא שבור
 *
 * ההפך: זה ה-`configure` הראשון מאז גל 3 שבאמת כותב. הוא כותב ל-settings.xml
 * `<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="4"/>
 * <w:numRestart w:val="eachPage"/><w:pos w:val="beneathText"/></w:footnotePr>`
 * — קנוני. ושלוש סיבות בכל זאת עוצרות אותו, וההנמקה המלאה
 * ב-docs/engine-gaps.md:
 *
 * 1. **אין קריאה.** אין בכל ה-API דרך לקרוא את ההגדרות שבמסמך, ולכן דיאלוג
 *    היה מציג ערכים שאינם של המסמך שעל המסך — בדיוק מה שאסרנו על עצמנו
 *    בדיאלוג של תוכן העניינים ובזה של המפתח.
 * 2. **כל קריאה מחליפה את האלמנט כולו.** `configure({ numbering: { start: 9 } })`
 *    אחרי `configure` מלא משאיר `<w:footnotePr><w:numStart w:val="9"/></w:footnotePr>`
 *    בלבד, ו-`numbering: {}` משאיר `<w:footnotePr></w:footnotePr>` ריק. כלומר
 *    אישור אחד על טופס שאינו יודע מה היה במסמך מוחק את מה שהוגדר ב-Word.
 * 3. **הערכים נכתבים גולמית.** `format: 'zigzag'` → `<w:numFmt w:val="zigzag"/>`,
 *    `start: 'א'` → `<w:numStart w:val="א"/>`. ושלושה ערכים **שכן בחוזה**
 *    נכתבים כאסימונים שאינם של Word: `restartPolicy: 'eachSection'` →
 *    `eachSection` (התקן: `eachSect`), `format: 'symbol'` → `symbol`
 *    (התקן: `chicago`), ומיקום הערת סיום → `sectionEnd`/`documentEnd`
 *    (התקן: `sectEnd`/`docEnd`).
 *
 * **מספור עברי:** אפשרי טכנית ואינו נשלח. `numFmt` נכתב גולמית, ולכן
 * `format: 'hebrew1'` מייצר `<w:numFmt w:val="hebrew1"/>` — אסימון תקני של
 * Word ובדיוק המספור שספר תורני רוצה. אבל `'hebrew1'` אינו ב-union של
 * `FootnoteNumberingConfig`, כלומר זו הישענות על ערך שאינו בטיפוסים
 * הציבוריים. הבריף אוסר את זה במפורש, ולכן הממצא מדווח למפקח ואינו ממומש.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';
import { captureNoteAnchor, carryFormatPastNote, type NoteRunFormatTarget } from './footnote-run-format';

export type NoteType = 'footnote' | 'endnote';

/**
 * `FootnoteAddress` — מה שכל פעולת הערה מקבלת כ-`target`.
 *
 * `entityType` הוא `'footnote'` גם עבור הערת סיום; זה אינו העתק-הדבק אלא
 * החוזה עצמו, וההסבר מה זה עושה בהערת הפתיחה.
 */
interface NoteAddress {
  kind: 'entity';
  entityType: 'footnote';
  noteId: string;
}

/** `FootnoteDomain` בחלק שנצרך כאן. הוא כבר עטוף ב-`DiscoveryItem`. */
interface RawNote {
  address?: { noteId?: string };
  type?: string;
  noteId?: string;
  displayNumber?: string;
  content?: string;
}

interface DiscoveryPage<T> {
  items?: readonly T[];
  total?: number;
}

export interface FootnotesDocumentApi {
  footnotes?: {
    insert?: (input: { type: NoteType; content: string }) => MaybePromise<DocReceipt & { footnote?: NoteAddress }>;
    list?: (query?: {
      type?: NoteType;
      limit?: number;
      offset?: number;
    }) => MaybePromise<DiscoveryPage<RawNote> | undefined>;
    get?: (input: { target: NoteAddress }) => MaybePromise<RawNote | undefined>;
    update?: (input: {
      target: NoteAddress;
      patch: { content: string };
    }) => MaybePromise<DocReceipt>;
    remove?: (input: { target: NoteAddress }) => MaybePromise<DocReceipt>;
  };
}

export interface FootnotesHost {
  activeEditor?: { doc?: FootnotesDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type FootnotesTarget = SuperDoc | FootnotesHost | null | undefined;

const NOTE_LABEL: Record<NoteType, string> = {
  footnote: 'הוספת הערת שוליים נכשלה',
  endnote: 'הוספת הערת סיום נכשלה',
};

/**
 * מוסיפה הערה במקום הסמן.
 *
 * לעולם אינה זורקת: `footnotes.insert` זורק `INVALID_INPUT` על קלט פסול במקום
 * להחזיר קבלה, וחריגה מפקד ב-Ribbon מפילה את רינדור הרצועה כולה.
 */
export async function insertNote(
  host: FootnotesTarget,
  type: NoteType,
): Promise<CommandOutcome> {
  const failedAction = NOTE_LABEL[type];
  const insert = (host as FootnotesHost | null | undefined)?.activeEditor?.doc?.footnotes?.insert;

  if (typeof insert !== 'function') {
    // אותו נוסח שהתכנית קובעת ב-§12, ואותו נוסח שהיכולת מחזירה — כדי שהמשתמש
    // יראה את אותו הסבר בין אם הפקד מנוטרל ובין אם הוא נלחץ לפני שהיכולות נקראו.
    return { ok: false, message: `${failedAction}: אינו זמין בגרסה זו`, reason: 'command-unsupported' };
  }

  // לפני ההכנסה: אחריה הסמן כבר בתוך ההערה. ראו footnote-run-format.ts.
  const anchor = await captureNoteAnchor(host as NoteRunFormatTarget);

  let receipt: DocReceipt & { footnote?: NoteAddress };
  try {
    receipt = await insert({ type, content: '' });
  } catch (error) {
    return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
  }

  if (receipt?.success === false) {
    return { ok: false, message: receiptFailureText(failedAction, receipt), reason: receipt.failure?.code };
  }

  if (anchor && typeof receipt?.footnote?.noteId === 'string') {
    carryFormatPastNote(host as NoteRunFormatTarget, anchor, { noteId: receipt.footnote.noteId, type });
  }
  return { ok: true };
}

/* ------------------------------------------------------------------ */
/* ניהול ההערות שכבר במסמך                                             */
/* ------------------------------------------------------------------ */

/**
 * נוסחי הכשל של העריכה וההסרה, לפי סוג ההערה.
 *
 * שתי טבלאות ולא ניסוח אחד גנרי: „הערת שוליים” ו„הערת סיום” הן שתי נקבות
 * שונות בשם, והמשתמש שלחץ על שורה של הערת סיום צריך לקרוא עליה — במיוחד
 * כשההודעה מסבירה למה דווקא היא אינה ניתנת לזיהוי.
 */
const UPDATE_FAILED: Record<NoteType, string> = {
  footnote: 'עריכת הערת השוליים נכשלה',
  endnote: 'עריכת הערת הסיום נכשלה',
};

const REMOVE_FAILED: Record<NoteType, string> = {
  footnote: 'הסרת הערת השוליים נכשלה',
  endnote: 'הסרת הערת הסיום נכשלה',
};

const READ_FAILED = 'קריאת ההערות נכשלה';

const LOADING_DETAIL = 'המסמך עדיין נטען';
const NOT_FOUND_DETAIL = 'ההערה אינה נמצאת במסמך';

/** מה שמוצג כשהתוכן נדחה. ריק אינו נשלח — ראו `normalizeNoteContent`. */
export const NOTE_CONTENT_HINT = 'יש להקליד את תוכן ההערה';

/**
 * הסירוב היחיד שנובע מהמנוע ולא מהמשתמש, והוא הממצא המרכזי של הגל: כתובת
 * ההערה אינה נושאת את הסוג, ולכן הערת סיום שמספרה זהה למספר של הערת שוליים
 * קיימת אינה ניתנת לפנייה כלל — כל פעולה עליה תיפול על הערת השוליים.
 * ההנמקה המלאה, כולל המדידה, בהערת הפתיחה.
 */
const AMBIGUOUS_DETAIL =
  'הערת הסיום נושאת את אותו מספר כמו הערת שוליים שבמסמך, והמנוע אינו יודע להבדיל ביניהן';

function unavailable(failedAction: string, detail: string, reason: string): CommandOutcome {
  return { ok: false, message: `${failedAction}: ${detail}`, reason };
}

/** אותו נוסח שהתכנית קובעת ב-§12, וזהה לזה שהיכולת מחזירה. ראו למעלה. */
function unsupported(failedAction: string): CommandOutcome {
  return {
    ok: false,
    message: `${failedAction}: אינו זמין בגרסה זו`,
    reason: 'command-unsupported',
  };
}

/** קריאה למנוע שאינה זורקת החוצה. ראו הערת ה„לעולם לא זורקת” שלמעלה. */
async function attempt<T>(
  failedAction: string,
  call: () => MaybePromise<T>,
): Promise<{ ok: true; value: T } | { ok: false; outcome: CommandOutcome }> {
  try {
    return { ok: true, value: await call() };
  } catch (error) {
    return {
      ok: false,
      outcome: { ok: false, message: thrownText(failedAction, error), reason: 'threw' },
    };
  }
}

/** כשל הקבלה, או `null` כשהיא הצליחה. `NO_OP` נחשב הצלחה — ראו header-footer.ts. */
function failureOf(failedAction: string, receipt: DocReceipt | undefined): CommandOutcome | null {
  const code = receipt?.failure?.code;
  if (receipt?.success !== false || code === 'NO_OP') return null;
  return { ok: false, message: receiptFailureText(failedAction, receipt), reason: code };
}

function docOf(host: FootnotesTarget): FootnotesDocumentApi | null {
  return (host as FootnotesHost | null | undefined)?.activeEditor?.doc ?? null;
}

function addressOf(noteId: string): NoteAddress {
  return { kind: 'entity', entityType: 'footnote', noteId };
}

/** גודל העמוד בכל שאיבה, ובלם מפני מנוע שיחזיר `total` שאינו יורד. ראו fields.ts. */
const PAGE_SIZE = 200;
const PAGE_GUARD = 1000;

/**
 * כל ההערות, בשאיבת עמודים עד `total`.
 *
 * `total` ולא `items.length`: `footnotes.list` הוא `DiscoveryOutput`, כלומר
 * `items` הוא עמוד תחת `limit`/`offset` (נמדד — `limit: 1, offset: 1` החזיר
 * את השנייה בלבד, עם `total` מלא). ספר תורני עם מאות הערות שוליים הוא בדיוק
 * המסמך שבו העמוד הראשון אינו הכול, וזה התרחיש של אוצריא ולא מקרה קצה.
 */
async function collectPages<T>(
  failedAction: string,
  list: (query: { limit: number; offset: number }) => MaybePromise<DiscoveryPage<T> | undefined>,
): Promise<{ ok: boolean; items: T[]; outcome?: CommandOutcome }> {
  const items: T[] = [];
  let offset = 0;
  let guard = 0;

  for (;;) {
    const listed = await attempt(failedAction, () => list({ limit: PAGE_SIZE, offset }));
    if (!listed.ok) return { ok: false, items, outcome: listed.outcome };

    const page = listed.value?.items ?? [];
    items.push(...page);
    if (page.length === 0) return { ok: true, items };

    offset += page.length;

    const total = listed.value?.total;
    if (!Number.isFinite(total) || offset >= (total as number)) return { ok: true, items };
    if (++guard > PAGE_GUARD) return { ok: true, items };
  }
}

/**
 * התוכן כפי שיישלח, או `null` כשהוא פסול.
 *
 * ריק **אינו** ערך חוקי כאן, ולא מפני שהמנוע דוחה אותו — הוא דווקא מקבל
 * אותו ומרוקן את ההערה. הערה ריקה במסמך היא סימן הפניה שאין מאחוריו דבר,
 * והדרך להיפטר ממנה היא „הסר הערה”; שמירה שמרוקנת בשקט הייתה מוחקת טקסט
 * במסלול שהמשתמש קורא לו „שמירה”.
 *
 * ירידת שורה מכווצת לרווח: נמדד שהיא נכתבת **גולמית לתוך `<w:t>`** (כמו
 * בכיתובים), כלומר היא אינה הופכת לפסקה שנייה בהערה אלא לרווח שאיש לא ביקש.
 */
export function normalizeNoteContent(raw: string): string | null {
  if (typeof raw !== 'string') return null;
  const content = raw.replace(/[\n\r\t]+/g, ' ').trim();
  return content === '' ? null : content;
}

/** הערה אחת כפי שהדיאלוג מציג אותה. */
export interface NoteSummary {
  /** `noteId` — ייחודי **בתוך** הסוג בלבד. ראו `NoteRef`. */
  id: string;
  type: NoteType;
  /** המספר שהמנוע מדווח. אינו בהכרח המספר ש-Word יציג — ראו `noteDisplay`. */
  number: string;
  content: string;
  /** מה שמוצג ברשימה: „הערת שוליים 1: תוכן ההערה”. */
  display: string;
}

/** מה שהממשק צריך לדעת. תצלום ולא מנוי, כמו header-footer.ts. */
export interface NotesState {
  notes: readonly NoteSummary[];
}

export function emptyNotesState(): NotesState {
  return { notes: [] };
}

const NOTE_NAME: Record<NoteType, string> = {
  footnote: 'הערת שוליים',
  endnote: 'הערת סיום',
};

/** „הערת שוליים 1: תוכן ההערה”. */
export function noteDisplay(type: NoteType, number: string, content: string): string {
  const head = number === '' ? NOTE_NAME[type] : `${NOTE_NAME[type]} ${number}`;
  return content === '' ? head : `${head}: ${content}`;
}

/**
 * ההפניה שהממשק מחזיק להערה: מספר **וסוג**.
 *
 * הסוג אינו קישוט ואינו חלק מהכתובת שנשלחת למנוע — הוא מה שמאפשר לוודא
 * שהכתובת נפתרה להערה שהמשתמש בחר ולא לשכנתה. ראו `resolveNote`.
 */
export interface NoteRef {
  id: string;
  type: NoteType;
}

/**
 * ההערות שבמסמך, בסדר שהמנוע מחזיר. לעולם אינה זורקת: כשל של קריאה מחזיר
 * רשימה ריקה, כלומר הדיאלוג יאמר „אין הערות” — ולא ימציא רשומות.
 *
 * הערה בלי `noteId` מדולגת: הצגתה הייתה מייצרת שורה שלחיצה עליה שולחת
 * `undefined` ל-`footnotes.remove`.
 */
async function collectNotes(
  host: FootnotesTarget,
  failedAction: string,
): Promise<{ ok: boolean; notes: NoteSummary[]; outcome?: CommandOutcome }> {
  const list = docOf(host)?.footnotes?.list;
  if (typeof list !== 'function') {
    return { ok: false, notes: [], outcome: unsupported(failedAction) };
  }

  const listed = await collectPages<RawNote>(failedAction, (query) => list(query));

  const notes: NoteSummary[] = [];
  for (const raw of listed.items) {
    const id = typeof raw.noteId === 'string' ? raw.noteId : raw.address?.noteId;
    if (typeof id !== 'string' || id === '') continue;
    // סוג שאינו אחד משניים מדולג ואינו „נחשב הערת שוליים”: הסוג הוא מה
    // שהעריכה וההסרה מאמתות מולו, וניחוש כאן הוא מחיקה של ההערה השכנה.
    const type = raw.type === 'endnote' ? 'endnote' : raw.type === 'footnote' ? 'footnote' : null;
    if (type === null) continue;
    const number = typeof raw.displayNumber === 'string' ? raw.displayNumber : '';
    const content = typeof raw.content === 'string' ? raw.content : '';
    notes.push({ id, type, number, content, display: noteDisplay(type, number, content) });
  }
  return { ok: listed.ok, notes, outcome: listed.outcome };
}

/**
 * „רשימת ההערות” לדיאלוג.
 *
 * כשל אמצע-שאיבה מחזיר **ריק** ולא את העמודים שהספיקו להיקרא, כמו
 * `readIndexState` בגל 5 — וכאן זה חמור יותר משם. רשימה חלקית אינה רק
 * ספירה שגויה: הדו-משמעיות של הערת סיום נקבעת מכך שיש ברשימה הערת שוליים
 * באותו מספר, ואם התאומה יושבת בעמוד שלא נשאב הדיאלוג יציע לערוך הערה
 * שהמודול יסרב לגעת בה — סירוב שהמשתמש רואה בלי שום סימן מקדים. „אין
 * הערות” הוא תשובה כנה; „שתיים מתוך מאתיים” אינה.
 */
export async function listNotes(host: FootnotesTarget): Promise<NoteSummary[]> {
  const listed = await collectNotes(host, READ_FAILED);
  return listed.ok ? listed.notes : [];
}

/**
 * קוראת את מצב ההערות במסמך. לעולם אינה זורקת: כשל של קריאה מחזיר „אין”,
 * כלומר ה-tooltip יאמר שאין מה לערוך — ולא ימציא מספר. כשל אמצע-שאיבה נכלל
 * בזה; ההנמקה ב-`listNotes`.
 */
export async function readNotesState(host: FootnotesTarget): Promise<NotesState> {
  if (!docOf(host)) return emptyNotesState();
  return { notes: await listNotes(host) };
}

/**
 * מאמתת שהכתובת נפתרת אל ההערה שהמשתמש בחר, **לפני** שנוגעים במסמך.
 *
 * זו רשת הביטחון של הגל, והיא נדרשת מפני שהכתובת אינה נושאת את הסוג: במסמך
 * שיש בו הערת שוליים 1 והערת סיום 1, שתיהן `{ noteId: '1' }`, והמנוע פותר
 * תמיד לטובת הערת השוליים (נמדד). בלי הבדיקה הזאת „הסר” על שורה של הערת
 * סיום היה מוחק הערת שוליים אחרת לגמרי ומדווח „בוצע”.
 *
 * `get` ולא השוואה מול `list`: הוא שואל את המנוע עצמו למה הכתובת נפתרת,
 * כלומר הוא מודד את אותו מסלול שהעריכה וההסרה ילכו בו — ולא מסלול מקביל
 * שעשוי לחלוק עליו.
 */
async function resolveNote(
  host: FootnotesTarget,
  failedAction: string,
  ref: NoteRef,
): Promise<{ ok: true; target: NoteAddress } | { ok: false; outcome: CommandOutcome }> {
  const get = docOf(host)?.footnotes?.get;
  if (typeof get !== 'function') return { ok: false, outcome: unsupported(failedAction) };

  const target = addressOf(ref.id);
  let info: RawNote | undefined;
  try {
    info = await get({ target });
  } catch {
    // `get` **זורק** על כתובת שאינה קיימת ואינו מחזיר קבלה (נמדד:
    // „footnote/endnote was not found.”), ורשימה שהתיישנה היא בדיוק המצב
    // שבו זה קורה. ההודעה שלנו עדיפה על ההודעה האנגלית של המנוע.
    return { ok: false, outcome: unavailable(failedAction, NOT_FOUND_DETAIL, 'note-not-found') };
  }

  if (info?.type !== ref.type) {
    if (ref.type === 'endnote') {
      return { ok: false, outcome: unavailable(failedAction, AMBIGUOUS_DETAIL, 'note-ambiguous') };
    }
    return { ok: false, outcome: unavailable(failedAction, NOT_FOUND_DETAIL, 'note-not-found') };
  }

  return { ok: true, target };
}

/**
 * „ערוך הערה” — מחליפה את תוכן ההערה.
 *
 * קריאה אחת ל-`footnotes.update`, ולא `remove`+`insert` כמו בכיתובים: נמדד
 * שהפעולה כאן **מחליפה** את התוכן ואינה מוסיפה עליו. ההנמקה, כולל שלושת
 * הצעדים שנמדדו, בהערת הפתיחה.
 */
export async function updateNote(
  host: FootnotesTarget,
  ref: NoteRef,
  content: string,
): Promise<CommandOutcome> {
  const failedAction = UPDATE_FAILED[ref.type] ?? UPDATE_FAILED.footnote;

  const text = normalizeNoteContent(content);
  if (text === null) {
    return { ok: false, message: `${failedAction}: ${NOTE_CONTENT_HINT}`, reason: 'invalid-content' };
  }

  const doc = docOf(host);
  if (!doc) return unavailable(failedAction, LOADING_DETAIL, 'document-api-unavailable');

  const update = doc.footnotes?.update;
  if (typeof update !== 'function') return unsupported(failedAction);

  const resolved = await resolveNote(host, failedAction, ref);
  if (!resolved.ok) return resolved.outcome;

  const updated = await attempt(failedAction, () =>
    update({ target: resolved.target, patch: { content: text } }),
  );
  if (!updated.ok) return updated.outcome;

  return failureOf(failedAction, updated.value) ?? { ok: true };
}

/**
 * „הסר הערה” — מוחקת את ההערה ואת סימן ההפניה שלה.
 *
 * צעד אחד: נמדד ש-`footnotes.remove` מוריד את ההערה מ-`footnotes.xml`
 * ומ-`list`, והסרה חוזרת על אותה כתובת מוחזרת `TARGET_NOT_FOUND`. מה שהוא
 * **כן** משאיר בגוף המסמך הוא ריצה ריקה בסגנון `FootnoteReference` בלי
 * `<w:footnoteReference>` בתוכה — שארית בלתי נראית שאין לתוסף דרך לנקות
 * (ראו docs/engine-gaps.md), ואינה פוגמת בתקינות הקובץ.
 */
export async function removeNote(host: FootnotesTarget, ref: NoteRef): Promise<CommandOutcome> {
  const failedAction = REMOVE_FAILED[ref.type] ?? REMOVE_FAILED.footnote;

  if (typeof ref?.id !== 'string' || ref.id === '') {
    return unavailable(failedAction, 'יש לבחור הערה', 'no-note');
  }

  const doc = docOf(host);
  if (!doc) return unavailable(failedAction, LOADING_DETAIL, 'document-api-unavailable');

  const remove = doc.footnotes?.remove;
  if (typeof remove !== 'function') return unsupported(failedAction);

  const resolved = await resolveNote(host, failedAction, ref);
  if (!resolved.ok) return resolved.outcome;

  const removed = await attempt(failedAction, () => remove({ target: resolved.target }));
  if (!removed.ok) return removed.outcome;

  return failureOf(failedAction, removed.value) ?? { ok: true };
}
