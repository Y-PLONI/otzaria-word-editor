/**
 * „תאריך ושעה” בלוח העברי — „כ״ג בחשוון תשפ״ז”.
 *
 * ## למה טקסט ולא שדה `DATE`
 *
 * ב-Word תאריך עברי הוא שדה `DATE` עם המתג `\h` ("Hebrew lunar calendar").
 * כאן הוא נכתב כ**טקסט** קבוע במקום הסמן, כמו „תאריך שאינו מתעדכן” ב-Word.
 * בלוח הלועזי הוא נשאר שדה, כמו קודם (engine/fields.ts).
 * „שעה בלבד” אינה תלויה בלוח ונשארת שדה בשני המצבים.
 *
 * ## מאיפה התאריך
 *
 * `Intl.DateTimeFormat` עם `ca-hebrew` — סינכרוני, בלי רשת ובלי מאחז, ונותן
 * את המרת הלוח ואת שם החודש בעברית („חשוון”, „אדר א׳”). מה שהוא **לא** נותן
 * הוא אותיות: היום והשנה חוזרים כספרות („23”, „5787”), ולכן הם מומרים כאן
 * לגימטריה. הלוח מתחלף בחצות האזרחית ולא בשקיעה — אין לנו מיקום, ובזה
 * נוהג גם Word.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';
import { readDocSelection, type SelectionDocumentApi, type SelectionTarget } from './doc-selection';

/** הלוח שבו „תאריך ושעה” כותב את התאריך. */
export type DateCalendar = 'hebrew' | 'gregorian';

/** ברירת המחדל: עברי. משתמשי אוצריא כותבים תאריך עברי. */
export const DEFAULT_DATE_CALENDAR: DateCalendar = 'hebrew';

/** מפתח ההגדרה שזוכר את הבחירה בין הפעלות (host/settings.ts). */
export const DATE_CALENDAR_SETTING = 'insert-date-calendar';

/** ערך שמור שאינו אחד משני הלוחות — `null`, והקורא נופל לברירת המחדל. */
export function parseDateCalendar(raw: unknown): DateCalendar | null {
  return raw === 'hebrew' || raw === 'gregorian' ? raw : null;
}

const ONES = ['', 'א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט'];
const TENS = ['', 'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ'];
const HUNDREDS = ['', 'ק', 'ר', 'ש', 'ת'];

const GERESH = '׳';
const GERSHAYIM = '״';

/**
 * מספר 1–999 באותיות, עם גרש לאות אחת וגרשיים לפני האחרונה. 15 ו-16 הם
 * ט״ו וט״ז ולא י״ה וי״ו, ומאות מעל 400 נכתבות בצירוף ת׳ (800 = ת״ת).
 */
export function hebrewNumeral(value: number): string {
  let rest = Math.trunc(value) % 1000;
  if (rest <= 0) return '';
  let letters = '';
  while (rest >= 400) {
    letters += 'ת';
    rest -= 400;
  }
  letters += HUNDREDS[Math.floor(rest / 100)];
  rest %= 100;
  if (rest === 15 || rest === 16) {
    letters += rest === 15 ? 'טו' : 'טז';
  } else {
    letters += TENS[Math.floor(rest / 10)] + ONES[rest % 10];
  }
  return letters.length === 1
    ? letters + GERESH
    : letters.slice(0, -1) + GERSHAYIM + letters.slice(-1);
}

const HEBREW_CALENDAR = new Intl.DateTimeFormat('he-u-ca-hebrew', {
  day: 'numeric',
  month: 'long',
  year: 'numeric',
});

/** „כ״ג בחשוון תשפ״ז”. השנה בלי האלפים, כמקובל. */
export function formatHebrewDate(date: Date): string {
  const parts = HEBREW_CALENDAR.formatToParts(date);
  const part = (type: Intl.DateTimeFormatPartTypes) => parts.find((p) => p.type === type)?.value ?? '';
  const day = hebrewNumeral(Number(part('day')));
  const year = hebrewNumeral(Number(part('year')));
  return `${day} ב${part('month')} ${year}`;
}

/** „14:05” — אותה תמונה כמו `HH:mm` של השדה הלועזי. */
export function formatTime(date: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

/** הטקסט שנכתב: תאריך, או תאריך ושעה מופרדים ברווח — כמו `dd/MM/yyyy HH:mm`. */
export function hebrewDateText(date: Date, withTime: boolean): string {
  const text = formatHebrewDate(date);
  return withTime ? `${text} ${formatTime(date)}` : text;
}

interface HebrewDateDocumentApi extends SelectionDocumentApi {
  insert?: (input: { value: string; type: 'text'; target: unknown }) => MaybePromise<DocReceipt>;
}

interface SelectionPointLike {
  kind?: string;
  blockId?: unknown;
  offset?: unknown;
  story?: unknown;
}

interface HebrewDateHost {
  activeEditor?: { doc?: HebrewDateDocumentApi | null } | null;
  ui?: { selection?: { apply?: (target: unknown) => { ok?: boolean } | void } | null } | null;
}

export type HebrewDateTarget = SuperDoc | HebrewDateHost | null | undefined;

const INSERT_FAILED = 'הוספת התאריך נכשלה';
/** אותו נוסח כמו בשדות — ראו `NO_TARGET_DETAIL` ב-engine/fields.ts. */
const NO_TARGET_DETAIL = 'יש ללחוץ בגוף המסמך, על שורת טקסט שיש בה תו אחד לפחות, ואז להוסיף את התאריך';

/**
 * מציבים את הסמן אחרי הטקסט שנכתב, כמו ב-at-mention-overlay.ts
 * וב-voice-dictation.ts, כולל הניסיונות החוזרים כש-`apply` מחזיר `ok: false`.
 *
 * רק כשהיעד כולו בפסקה אחת — בבחירה שחוצה פסקאות ההיסט שאחרי ההכנסה אינו
 * נגזר מהיעד, ועדיף להשאיר את הסמן במקום שהמנוע בחר מאשר לנחש.
 */
async function placeCaretAfter(host: HebrewDateHost, target: unknown, length: number): Promise<void> {
  const selection = host.ui?.selection;
  if (typeof selection?.apply !== 'function') return;
  const { start, end, story } = target as { start?: SelectionPointLike; end?: SelectionPointLike; story?: unknown };
  if (typeof start?.blockId !== 'string' || start.blockId !== end?.blockId) return;
  if (typeof start.offset !== 'number' || typeof end.offset !== 'number') return;

  // בחירה הפוכה (גרירה משמאל לימין) מגיעה עם `start` אחרי `end`.
  const point = { ...start, offset: Math.min(start.offset, end.offset) + length };
  const caret = { kind: 'selection', start: point, end: point, ...(story !== undefined ? { story } : {}) };
  for (let tries = 0; tries < 10; tries++) {
    try {
      if (selection.apply.call(selection, caret)?.ok !== false) return;
    } catch (error) {
      // התאריך כבר נכתב; מיקום הסמן אינו שווה הודעת שגיאה למשתמש.
      console.warn('[otzaria-word] תאריך עברי: הצבת הסמן אחרי התאריך נכשלה', error);
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
}

/**
 * כותבת את התאריך העברי במקום הסמן. בחירה של טווח מוחלפת בתאריך, כמו
 * הקלדה. היעד הוא `selectionTarget` — ראו doc-selection.ts.
 */
export async function insertHebrewDate(
  host: HebrewDateTarget,
  withTime: boolean,
  now: Date = new Date(),
): Promise<CommandOutcome> {
  const doc = (host as HebrewDateHost | null | undefined)?.activeEditor?.doc ?? null;
  if (!doc) {
    return { ok: false, message: `${INSERT_FAILED}: המסמך עדיין נטען`, reason: 'document-api-unavailable' };
  }
  if (typeof doc.insert !== 'function') {
    return { ok: false, message: `${INSERT_FAILED}: אינו זמין בגרסה זו`, reason: 'command-unsupported' };
  }

  const selection = await readDocSelection(host as SelectionTarget);
  if (!selection.selectionTarget) {
    return { ok: false, message: `${INSERT_FAILED}: ${NO_TARGET_DETAIL}`, reason: 'no-selection' };
  }

  const value = hebrewDateText(now, withTime);
  let receipt: DocReceipt | undefined;
  try {
    receipt = await doc.insert({
      value,
      type: 'text',
      target: selection.selectionTarget,
    });
  } catch (error) {
    return { ok: false, message: thrownText(INSERT_FAILED, error), reason: 'threw' };
  }
  if (receipt?.success === false) {
    return {
      ok: false,
      message: receiptFailureText(INSERT_FAILED, receipt),
      reason: receipt.failure?.code,
    };
  }
  await placeCaretAfter(host as HebrewDateHost, selection.selectionTarget, value.length);
  return { ok: true };
}
