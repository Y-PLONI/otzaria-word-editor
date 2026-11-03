/**
 * שדות Word — מספר עמוד, מספר העמודים ותאריך — דרך `doc.fields`.
 *
 * ## למה `instruction` גולמי ולא „סוג שדה”
 *
 * `FieldInsertInput` בחוזה נושא `at`, `instruction`, `mode: 'raw'`, ועוד שני
 * שדות אופציונליים שאינם נצרכים כאן (`cachedResultText` ו-`updatePolicy`, ראו
 * „למה אין תאריך קבוע” למטה). אין בו `fieldType` ואין קטלוג של שדות מוכרים —
 * מה שנשלח הוא **קוד השדה של Word** כמחרוזת, בדיוק כפי שהוא נכתב בין הסוגריים
 * המסולסלים במסמך. לכן כל מחרוזת כאן היא קוד שדה תקני של Word ותו לא:
 * `PAGE`, `NUMPAGES`, `DATE \\@ "dd/MM/yyyy HH:mm"`.
 *
 * ## למה `DATE` נושא מתג `\@`, ומה בדיוק נמדד
 *
 * מי שמחשב את תוצאת השדה כאן אינו Word: `fields.insert` מכניס שדה עם תוצאה
 * ריקה, ו-`fields.rebuild` הוא שמחשב אותה בתוך המנוע (כך כתוב בחוזה:
 * „inserted with an empty cached result and is expected to be recomputed via
 * `fields.rebuild`”). החוזה אינו מצהיר על אילו מתגים ה-rebuild יודע לפרש,
 * ומתג שאינו מוכר אינו מייצר שגיאה אלא תוצאה שגויה בשקט — ולכן המתגים נמדדו
 * במנוע האמיתי (Chrome, `file://`, מסמך זרוע, `fields.insert` + `rebuild`
 * ואז `fields.list().items[].resolvedText`) ולא נוחשו. שלוש המדידות, ביום
 * 25.8.2026:
 *
 * - `DATE` ללא הלבשה            → `2026-08-24`      (ISO, ובאזור זמן UTC — יום אחורה)
 * - `DATE \@ "dd/MM/yyyy HH:mm"`  → `08/09/2026 12:54`      המתג מפורש, והתאריך מקומי
 * - `DATE \@ "d בMMMM yyyy"`→ `25 בAugust 2026` המתג מפורש, שמות החודשים לועזיים
 * - `DATE \* HEBREW`        → `2026-08-24`      המתג נבלע בשקט, בלי שגיאה
 *
 * כלומר: מתג תמונת-הפורמט `\@` מפורש כהלכה, ומתג לוח השנה אינו קיים. לכן
 * הפורמט המספרי הישראלי `dd/MM/yyyy HH:mm` נשלח — הוא נכון בעברית, אינו נשען על
 * שמות חודשים שהמנוע אינו מתרגם, והוא גם מה שמתקן את ה-ISO וההיסט של יום
 * שהתוצאה העירומה נותנת. לוח שנה עברי אינו אפשרי כאן, ואין להוסיף `\*` —
 * הוא נמדד כמתג שנבלע.
 *
 * הפורמט הזה הוא מה שנראה בתוסף. בפתיחת הקובץ ב-Word, Word מחשב את השדה
 * מחדש בעצמו לפי אותו מתג `\@` — זו תמונת פורמט תקנית שלו.
 *
 * ## למה אין „עמוד X מתוך Y”
 *
 * זה אינו שדה אחד אלא רצף של חמישה חלקים — טקסט „עמוד ”, שדה `PAGE`, טקסט
 * „ מתוך ”, שדה `NUMPAGES` — ולשם כך צריך להכניס טקסט **בין** שני שדות.
 * `fields.insert` מקבל `at: TextTarget`, כלומר כתובת טקסט מוחלטת, והקבלה
 * שהוא מחזיר נושאת `FieldAddress` (`blockId`/`occurrenceIndex`) ולא היסט
 * טקסט. כלומר אחרי ההכנסה הראשונה אין דרך ציבורית לדעת באיזה היסט להכניס את
 * החלק הבא, וכל ניחוש היה מייצר משפט מעורבב. שני הפקדים הנפרדים —
 * „מספר עמוד” ו„מספר העמודים במסמך” — הם מה שהחוזה מאפשר בוודאות.
 *
 * ## למה אין „תאריך קבוע” ואין „הסרת שדה”
 *
 * תאריך קבוע ב-Word הוא **טקסט**, לא שדה (זה מה ש„עדכן אוטומטית” הכבוי
 * בדיאלוג עושה), והחלופה בחוזה — `updatePolicy: 'preserveCached'` —
 * מוצהרת בקטלוג הפעולות ככזאת שמחזירה כרגע `CAPABILITY_UNAVAILABLE`. אין
 * כאן מה לממש.
 *
 * `fields.remove` אין לו פקד ברצועה — ב-Word מוחקים שדה כמו שמוחקים טקסט —
 * אבל הוא **כן** נצרך כאן, פנימית בלבד: ראו „מניעת קינון” למטה.
 *
 * ## מניעת קינון של שדה בתוך שדה, ולמה הבדיקה היא „הכנס ובדוק” ולא „בדוק מראש”
 *
 * `fields.insert` מקבל `at: TextTarget` בלי לבדוק אם הוא נופל **בתוך** שדה
 * קיים, ומכניס בשקט — שלא כמו `citations.insert`, שמסרב במצב הזה עם
 * `CAPABILITY_UNAVAILABLE` („text-range-in-field”, מתועד ב-`docs/engine-gaps.md`).
 * שלוש הכנסות רצופות באותה נקודה — בדיוק המסלול ש-hint התפריט „מספר עמוד”
 * מנחה עליו לצירוף „עמוד X מתוך Y” — קיננו שדה בתוך **תוצאת** שדה אחר: נמדד
 * ב-Chrome אמיתי, `PAGE` שהיה אמור לפתור ל-`"1"` פתר ל-`"28/08/202611"`.
 *
 * בדיקה **מראש** ("הסמן יושב על שדה?") הייתה הפתרון הנקי, אבל אין לה על מה
 * להישען: `fields.list` אינו מקבל סינון לפי בלוק או טווח (`FieldListInput`
 * הוא `type`/`limit`/`offset` בלבד — נבדק מול `fields.types.d.ts`), וכתובת
 * שדה (`FieldAddress`) היא `blockId`/`occurrenceIndex`, לא היסט טקסט — אין
 * דרך לדעת אילו תווים בבלוק שייכים לאיזה שדה. הדגל `nested` שה-`FieldDomain`
 * חושף **גם הוא לא עוזר**: נמדד ישירות (Chrome, אותו תרחיש קינון) ששני
 * השדות המקוננים חוזרים עם `nested: false` — הדגל תופס קינון של קוד שדה
 * בתוך קוד שדה (למשל `{ IF { PAGE } = 1 ... }`), לא את הסוג הזה של קינון,
 * שבו הכתובת החדשה נשתלת בתוך **הריצה של התוצאה המחושבת**.
 *
 * לכן הבדיקה כאן היא **אחרי** ההכנסה: תצלום של השדות הקיימים בבלוק
 * (`instruction`+`resolvedText`, לפי `fieldId` יציב) לפני הקריאה ל-`insert`,
 * ואותו תצלום שוב אחריה. שדה קיים שה-`resolvedText` או ה-`instruction` שלו
 * השתנו כתוצאה מההכנסה — הוכחה שהיא נשתלה בתוכו, ולא לידו. במצב הזה השדה
 * החדש **מוסר** (`fields.remove`, ולכן הוא כן נצרך) והפעולה מדווחת כסירוב
 * מנומק בעברית, לא כהצלחה שקטה. משתמש שהסמן שלו כבר בתוך שדה מקבל הודעה
 * שאומרת להזיז אותו — בדיוק מה ש-`citations.insert` נותן חינם מהמנוע.
 *
 * הבדיקה מדלגת (ולא בולמת) כשאין `fieldId` יציב על אף שדה בבלוק — גרסת מנוע
 * ישנה בלי handle יציב אין דרך אמינה להשוות בינה לבין עצמה אחרי מוטציה, ועדיף
 * לוותר על ההגנה מאשר לחסום הכנסה תקינה על בסיס ניחוש.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';
import { readDocSelection, type SelectionDocumentApi, type SelectionTarget } from './doc-selection';

/**
 * קודי השדות שהמודול מכיר, לפי שמם ב-Word. קבועים ולא נבנים: מחרוזת שנבנית
 * מקלט של הממשק היא בדיוק הדרך שבה נשתל קוד שדה שגוי במסמך.
 */
export const FIELD_INSTRUCTIONS = {
  pageNumber: 'PAGE',
  pageCount: 'NUMPAGES',
  date: 'DATE \\@ "dd/MM/yyyy HH:mm"',
  dateOnly: 'DATE \\@ "dd/MM/yyyy"',
  timeOnly: 'DATE \\@ "HH:mm"',
} as const;

export type FieldKind = keyof typeof FIELD_INSTRUCTIONS;

/** `FieldMutationResult` — הצלחה נושאת `field`, כשל נושא `failure` כמו כל קבלה. */
interface FieldReceipt extends DocReceipt {
  field?: unknown;
}

/**
 * `FieldDomain` בחלק שנצרך כאן. `address` הוא מה ש-`rebuild`/`remove` מקבלים
 * כ-`target`. `resolvedText` ו-`nested` נוספו לצורך גילוי הקינון — ראו הערת
 * הפתיחה.
 */
interface FieldEntry {
  address?: unknown;
  instruction?: string;
  fieldType?: string;
  resolvedText?: string;
  nested?: boolean;
}

export interface FieldsDocumentApi extends SelectionDocumentApi {
  fields?: {
    list?: (query?: {
      type?: string;
      limit?: number;
      offset?: number;
    }) => MaybePromise<{ items?: readonly FieldEntry[]; total?: number } | undefined>;
    insert?: (input: {
      at: unknown;
      instruction: string;
      mode: 'raw';
    }) => MaybePromise<FieldReceipt>;
    rebuild?: (input: { target: unknown }) => MaybePromise<DocReceipt>;
    /** נצרך רק פנימית, לביטול שדה שהתגלה כמקונן. ראו הערת הפתיחה. */
    remove?: (input: { target: unknown; mode: 'raw' }) => MaybePromise<DocReceipt>;
  };
}

type FieldsList = NonNullable<NonNullable<FieldsDocumentApi['fields']>['list']>;

/** מה שנדרש מ-SuperDoc: רק הפאסדה של המסמך. ראו page-setup.ts. */
export interface FieldsHost {
  activeEditor?: { doc?: FieldsDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type FieldsTarget = SuperDoc | FieldsHost | null | undefined;

/**
 * הטיית הכשל בעברית תקנית לכל שדה. „מספר העמוד” זכר, „התאריך” זכר — הביטוי
 * נשמר שלם ואינו נגזר מהמזהה. ראו document-api.ts.
 */
const INSERT_FAILED: Record<FieldKind, string> = {
  pageNumber: 'הוספת מספר העמוד נכשלה',
  pageCount: 'הוספת מספר העמודים נכשלה',
  date: 'הוספת התאריך נכשלה',
  dateOnly: 'הוספת התאריך נכשלה',
  timeOnly: 'הוספת השעה נכשלה',
};

const REBUILD_FAILED = 'עדכון השדות נכשל';
/**
 * החוזה מונה שלוש סיבות ל-`target: null`: מסמך ריק, בחירת node (תמונה, טבלה)
 * וחוסר מיקוד בעורך. „יש למקם את הסמן” לא הבחין ביניהן ונקרא כשקר למי שהסמן
 * שלו כבר במקום. הנוסח כאן מכוון לפעולה שמכסה את שלושתן: לחיצה בגוף המסמך
 * מחזירה מיקוד ומחליפה בחירת node בסמן טקסט, ותו אחד לפחות פותר מסמך ריק.
 */
const NO_TARGET_DETAIL = 'יש ללחוץ בגוף המסמך, על שורת טקסט שיש בה תו אחד לפחות, ואז להוסיף את השדה';
const READ_FAILED = 'קריאת שדות המסמך נכשלה';
/** ראו „מניעת קינון” בהערת הפתיחה: הבדיקה היא „הכנס ובדוק”, לא „בדוק מראש”. */
const NESTED_FIELD_DETAIL =
  'הסמן היה בתוך שדה קיים, וההכנסה הייתה נבלעת בתוכו. ההכנסה בוטלה — יש להזיז את הסמן אל מחוץ לשדה הקיים ולנסות שוב';

/** גודל העמוד בכל שאיבה, ובלם מפני מנוע שיחזיר `total` שאינו יורד. */
const PAGE_SIZE = 200;
const PAGE_GUARD = 1000;

function unavailable(failedAction: string, detail: string, reason: string): CommandOutcome {
  return { ok: false, message: `${failedAction}: ${detail}`, reason };
}

/** הנוסח שהתכנית קובעת ב-§12, וזהה לזה שהיכולת מחזירה. ראו footnotes.ts. */
function unsupported(failedAction: string): CommandOutcome {
  return {
    ok: false,
    message: `${failedAction}: אינו זמין בגרסה זו`,
    reason: 'command-unsupported',
  };
}

/** קריאה למנוע שאינה זורקת החוצה. ראו הערת ה„לעולם לא זורקת” ב-footnotes.ts. */
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

/**
 * כשל הקבלה, או `null` כשהיא הצליחה. `NO_OP` נחשב הצלחה — שדה שהמנוע כבר
 * חישב מחדש אינו טעון עדכון, וזה בדיוק מה שהמשתמש ביקש. ראו header-footer.ts.
 */
function failureOf(failedAction: string, receipt: DocReceipt | undefined): CommandOutcome | null {
  const code = receipt?.failure?.code;
  if (receipt?.success !== false || code === 'NO_OP') return null;
  return { ok: false, message: receiptFailureText(failedAction, receipt), reason: code };
}

function docOf(host: FieldsTarget): FieldsDocumentApi | null {
  return (host as FieldsHost | null | undefined)?.activeEditor?.doc ?? null;
}

/**
 * כל השדות שיש להם כתובת, בשאיבת עמודים עד `total`. `total` ולא
 * `items.length`: `fields.list` הוא `DiscoveryOutput`, ו„עדכן שדות” שרץ על
 * העמוד הראשון בלבד היה משאיר שדות לא מעודכנים במסמך גדול — ובלי שום סימן,
 * כי הפעולה מדווחת הצלחה.
 */
async function collectAllFields(
  failedAction: string,
  list: FieldsList,
): Promise<{ ok: true; items: FieldEntry[] } | { ok: false; outcome: CommandOutcome }> {
  const items: FieldEntry[] = [];
  let offset = 0;
  let guard = 0;

  for (;;) {
    const listed = await attempt(failedAction, () => list({ limit: PAGE_SIZE, offset }));
    if (!listed.ok) return { ok: false, outcome: listed.outcome };

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
 * מפתח יציב לשדה, מ-`fieldId` (עם `storyId` להבחין בין stories). `null` כשאין
 * `fieldId` — גרסת מנוע בלי handle יציב, ואין דרך אמינה להשוות כתובת כזאת בין
 * שתי קריאות `list` נפרדות (`occurrenceIndex` זז עם כל מוטציה בבלוק).
 */
function fieldIdentity(address: unknown): string | null {
  const addr = address as { fieldId?: unknown; storyId?: unknown } | null | undefined;
  if (!addr || typeof addr !== 'object' || typeof addr.fieldId !== 'string' || addr.fieldId === '') {
    return null;
  }
  return typeof addr.storyId === 'string' && addr.storyId !== ''
    ? `${addr.storyId} ${addr.fieldId}`
    : addr.fieldId;
}

function blockIdOf(address: unknown): string | null {
  const blockId = (address as { blockId?: unknown } | null | undefined)?.blockId;
  return typeof blockId === 'string' ? blockId : null;
}

/** תצלום השדות **הקיימים** בבלוק נתון, לפי `fieldId`. `null` = כשל קריאה, לא נבלם ההכנסה בגללו. */
async function snapshotFieldsIn(
  list: FieldsList,
  blockId: string,
): Promise<Map<string, { instruction?: string; resolvedText?: string }> | null> {
  const collected = await collectAllFields(READ_FAILED, list);
  if (!collected.ok) return null;

  const snapshot = new Map<string, { instruction?: string; resolvedText?: string }>();
  for (const entry of collected.items) {
    if (blockIdOf(entry.address) !== blockId) continue;
    const key = fieldIdentity(entry.address);
    if (key === null) continue;
    snapshot.set(key, { instruction: entry.instruction, resolvedText: entry.resolvedText });
  }
  return snapshot;
}

/**
 * בודקת אם ההכנסה שזה עתה בוצעה קיננה בתוך שדה קיים, ומבטלת אותה אם כן.
 *
 * ההשוואה היא מול `before`, תצלום שנלקח **לפני** הקריאה ל-`insert`. שדה
 * קיים ש-`instruction` או `resolvedText` שלו השתנו — מי שהוא לא, בהגדרה, לא
 * אמור היה להשתנות מהכנסת שדה **אחר** — הוא ההוכחה לקינון. ראו הערת הפתיחה.
 */
async function undoIfNested(
  doc: FieldsDocumentApi,
  list: FieldsList,
  blockId: string,
  before: Map<string, { instruction?: string; resolvedText?: string }>,
  newField: unknown,
  failedAction: string,
): Promise<CommandOutcome | null> {
  const after = await snapshotFieldsIn(list, blockId);
  if (after === null) return null;

  let corrupted = false;
  for (const [key, was] of before) {
    const now = after.get(key);
    if (!now || now.instruction !== was.instruction || now.resolvedText !== was.resolvedText) {
      corrupted = true;
      break;
    }
  }
  if (!corrupted) return null;

  const remove = doc.fields?.remove;
  if (typeof remove === 'function' && newField !== undefined && newField !== null) {
    // כשל הביטול אינו הופך את הסירוב לשקט: השדה שקינן נכתב בכל מקרה, וההודעה
    // חייבת לומר זאת — גם אם הניקוי עצמו לא הצליח.
    await attempt(failedAction, () => remove({ target: newField, mode: 'raw' }));
  }

  return unavailable(failedAction, NESTED_FIELD_DETAIL, 'field-in-field');
}

/* ------------------------------------------------------------------ */
/* הכנסת שדה                                                           */
/* ------------------------------------------------------------------ */

/**
 * מכניסה שדה במקום הסמן.
 *
 * `at` הוא פרמטר **חובה** בחוזה, ואין לו ברירת מחדל „במקום הסמן” כמו שיש
 * ל-`footnotes.insert`. לכן הבחירה נקראת כאן ונמסרת כמו שהיא: היא ה-
 * `TextTarget` שהמנוע עצמו הקרין, ובנייה מחדש שלה הייתה מקבעת אצלנו אחת
 * מכמה צורות אפשריות. ראו doc-selection.ts.
 *
 * סמן מכווץ הוא יעד חוקי לחלוטין — `TextTarget` דורש קטע אחד לפחות, לא טווח
 * לא-ריק — וזה המצב הרגיל של הפקד: לוחצים בלי לסמן כלום, והשדה נכנס בסמן.
 */
async function insertField(host: FieldsTarget, kind: FieldKind): Promise<CommandOutcome> {
  const failedAction = INSERT_FAILED[kind];
  const doc = docOf(host);
  if (!doc) return unavailable(failedAction, 'המסמך עדיין נטען', 'document-api-unavailable');

  const insert = doc.fields?.insert;
  if (typeof insert !== 'function') return unsupported(failedAction);

  const selection = await readDocSelection(host as SelectionTarget);
  if (!selection.target) {
    return unavailable(failedAction, NO_TARGET_DETAIL, 'no-selection');
  }

  // תצלום השדות שכבר בבלוק, **לפני** ההכנסה — הבסיס להשוואה שמגלה קינון.
  // ראו „מניעת קינון” בהערת הפתיחה. `blockId === null` לא אמור לקרות כאן
  // (`selection.target` דורש קטע עם `blockId`), אבל הבדיקה מדלגת ולא זורקת.
  const list = doc.fields?.list;
  const before =
    typeof list === 'function' && selection.blockId !== null
      ? await snapshotFieldsIn(list, selection.blockId)
      : null;

  const inserted = await attempt(failedAction, () =>
    insert({ at: selection.target, instruction: FIELD_INSTRUCTIONS[kind], mode: 'raw' }),
  );
  if (!inserted.ok) return inserted.outcome;

  const failure = failureOf(failedAction, inserted.value);
  if (failure) return failure;

  const newField = inserted.value?.field;

  if (before && before.size > 0 && list && selection.blockId !== null) {
    const nested = await undoIfNested(doc, list, selection.blockId, before, newField, failedAction);
    if (nested) return nested;
  }

  // השדה נכנס עם תוצאה ריקה, וה-rebuild הוא שמחשב אותה. בלעדיו המשתמש רואה
  // מקום ריק ולא „3”, ואין לו שום רמז שצריך ללחוץ על „עדכן שדות”.
  return rebuildInsertedField(doc, newField, failedAction);
}

/**
 * מחשבת את התוצאה של השדה שזה עתה נכנס.
 *
 * כשל כאן אינו כשל של ההכנסה: השדה **במסמך**, והוא ייראה ברגע שיעודכן —
 * מ„עדכן שדות”, מפתיחה מחדש ב-Word, או מהדפסה. הודעת כשל על פעולה שהצליחה
 * היא בדיוק סוג הרעש שמלמד את המשתמש להתעלם מהודעות, ולכן החזרה כאן היא
 * `{ ok: true }` בכל מקרה.
 */
async function rebuildInsertedField(
  doc: FieldsDocumentApi,
  field: unknown,
  failedAction: string,
): Promise<CommandOutcome> {
  const rebuild = doc.fields?.rebuild;
  if (typeof rebuild !== 'function' || field === undefined || field === null) return { ok: true };

  await attempt(failedAction, () => rebuild({ target: field }));
  return { ok: true };
}

/** „מספר עמוד” — `{ PAGE }`. */
export function insertPageNumber(host: FieldsTarget): Promise<CommandOutcome> {
  return insertField(host, 'pageNumber');
}

/** „מספר העמודים במסמך” — `{ NUMPAGES }`. ראו הערת הפתיחה על „עמוד X מתוך Y”. */
export function insertPageCount(host: FieldsTarget): Promise<CommandOutcome> {
  return insertField(host, 'pageCount');
}

/** „תאריך ושעה” — `{ DATE \\@ "dd/MM/yyyy HH:mm" }`. המתג נמדד, ראו הערת הפתיחה. */
export function insertDate(host: FieldsTarget): Promise<CommandOutcome> {
  return insertField(host, 'date');
}

/** תאריך בלבד ללא שעה — `{ DATE \\@ "dd/MM/yyyy" }`. המתג נמדד, ראו הערת הפתיחה. */
export function insertDateOnly(host: FieldsTarget): Promise<CommandOutcome> {
  return insertField(host, 'dateOnly');
}

/** שעה בלבד ללא תאריך — `{ DATE \\@ "HH:mm" }`. המתג נמדד, ראו הערת הפתיחה. */
export function insertTimeOnly(host: FieldsTarget): Promise<CommandOutcome> {
  return insertField(host, 'timeOnly');
}

/* ------------------------------------------------------------------ */
/* קריאה ועדכון                                                        */
/* ------------------------------------------------------------------ */

/** מה שהממשק צריך לדעת: האם יש בכלל מה לעדכן. תצלום ולא מנוי, כמו header-footer.ts. */
export interface FieldsState {
  count: number;
}

export function emptyFieldsState(): FieldsState {
  return { count: 0 };
}

/**
 * מונה את שדות המסמך. לעולם אינה זורקת: כשל של קריאה מחזיר „אין שדות”, כלומר
 * ה-tooltip יאמר שאין מה לעדכן — ולא ימציא מספר.
 */
export async function readFieldsState(host: FieldsTarget): Promise<FieldsState> {
  const list = docOf(host)?.fields?.list;
  if (typeof list !== 'function') return emptyFieldsState();

  const listed = await attempt(READ_FAILED, () => list());
  if (!listed.ok) return emptyFieldsState();

  // `total` ולא `items.length`: `FieldsListResult` הוא `DiscoveryOutput`, כלומר
  // `items` הוא עמוד תחת `limit`/`offset` ו-`total` הוא המספר במסמך כולו. מונה
  // שסופר עמוד אחד היה אומר „20 שדות” במסמך של מאות. הנפילה ל-`items.length`
  // היא לגרסה שאינה חושפת `total` — עדיף מספר חלקי על אפס שקרי.
  const { items, total } = listed.value ?? {};
  if (typeof total === 'number' && Number.isFinite(total)) return { count: total };
  return { count: Array.isArray(items) ? items.length : 0 };
}

/** מה שמוחזר כשלא כל השדות התעדכנו. ראו `rebuildAllFields`. */
function partialRebuildOutcome(succeeded: number, failed: number): CommandOutcome {
  const succeededText =
    succeeded === 0 ? 'אף שדה לא עודכן' : succeeded === 1 ? 'שדה אחד עודכן' : `${succeeded} שדות עודכנו`;
  const failedText = failed === 1 ? 'שדה אחד נכשל ולא עודכן' : `${failed} שדות נכשלו ולא עודכנו`;
  return {
    ok: false,
    message: `עדכון השדות לא הושלם: ${succeededText}, ו${failedText}`,
    reason: 'partial-rebuild',
  };
}

/**
 * „עדכן שדות” — מחשבת מחדש כל שדה במסמך, כמו F9 על מסמך שכולו מסומן.
 *
 * **לא** עוצרת בכשל הראשון, ו**מרעננת** את הרשימה אחרי כל כשל — שני שינויים
 * ששרשרת אחת אליה. שדה מקונן (ראו „מניעת קינון” בהערת הפתיחה, ומסמכים שהגיעו
 * כבר מקוננים מ-Word או מגרסה קודמת) גורם ל-`rebuild` על השדה שקינן בתוכו
 * **לשחזר/להסיר** את המבנה השגוי — ומי שנמדד: אחרי `rebuild` כזה מספר השדות
 * במסמך יורד (5→3), כלומר כתובות ששאבנו **לפני** אותה מוטציה עלולות
 * להצביע על מה שכבר אינו קיים.
 *
 * הגרסה הקודמת שאבה את כל הכתובות **פעם אחת** ומיטטה אחת-אחת, ועצרה בכשל
 * הראשון (`TARGET_NOT_FOUND` על כתובת שהתיישנה) — כלומר כל שדה **אחרי** זה
 * במסמך נשאר לא-מעודכן, בלי שום סימן חוץ מהודעת הכשל היחידה. שדה בודד שנכשל
 * אינו סיבה לא לעדכן את שאר המסמך.
 *
 * לכן כאן: כל שדה מנוסה **פעם אחת** (`attempted`, לפי `fieldIdentity` —
 * וכשאין `fieldId` יציב, לפי הכתובת עצמה), הכשל שלו נספר ולא עוצר את הלולאה,
 * ואחרי **כל** כשל נשאבת רשימה טרייה לפני שממשיכים — כך ששאר השדות מטופלים
 * על סמך מבנה עדכני, לא על סמך תצלום שאולי כבר לא נכון. בסוף מדווח כמה
 * הצליחו וכמה נכשלו; `ok:true` רק כשכולם הצליחו.
 *
 * מסמך בלי שדות מחזיר הצלחה שקטה: המשתמש ביקש שכל השדות יהיו מעודכנים, וכולם
 * מעודכנים. אין כאן כשל שיש לדווח עליו.
 */
export async function rebuildAllFields(host: FieldsTarget): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(REBUILD_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const list = doc.fields?.list;
  const rebuild = doc.fields?.rebuild;
  if (typeof list !== 'function' || typeof rebuild !== 'function') return unsupported(REBUILD_FAILED);

  const attempted = new Set<string>();
  let succeeded = 0;
  let failed = 0;
  let pending: FieldEntry[] = [];
  let needsFreshList = true;
  let guard = 0;

  for (;;) {
    if (needsFreshList) {
      const collected = await collectAllFields(REBUILD_FAILED, list);
      if (!collected.ok) {
        // כשל הקריאה עצמה: אם עוד לא הצלחנו/נכשלנו באף שדה זו הודעת הכשל
        // היחידה שיש; אחרת יש כבר תוצאה חלקית לדווח עליה, ועדיף עליה.
        if (succeeded === 0 && failed === 0) return collected.outcome;
        break;
      }

      pending = collected.items.filter((entry) => {
        // שדה בלי כתובת אינו יעד חוקי ל-`rebuild`, ושליחתו הייתה חריגת
        // `INVALID_INPUT` על שדה שאיש לא ביקש במיוחד.
        if (entry.address === undefined || entry.address === null) return false;
        return !attempted.has(fieldIdentity(entry.address) ?? JSON.stringify(entry.address));
      });
      needsFreshList = false;
    }

    const entry = pending.shift();
    if (!entry) break; // רשימה טרייה בלי אף שדה חדש — הכול נוסה, סיימנו.

    attempted.add(fieldIdentity(entry.address) ?? JSON.stringify(entry.address));

    const rebuilt = await attempt(REBUILD_FAILED, () => rebuild({ target: entry.address }));
    const failure = rebuilt.ok ? failureOf(REBUILD_FAILED, rebuilt.value) : rebuilt.outcome;
    if (failure) {
      failed++;
      // כתובת שנכשלה עשויה לסמן שהמבנה השתנה (בדיוק המצב שתועד למעלה) —
      // שאר הכתובות שכבר נשאבו עלולות להיות מיושנות גם הן.
      needsFreshList = true;
    } else {
      succeeded++;
    }

    if (++guard > PAGE_GUARD) break;
  }

  return failed === 0 ? { ok: true } : partialRebuildOutcome(succeeded, failed);
}
