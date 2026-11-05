/**
 * הגנת מסמך (גל 19): `protection.get/setEditingRestriction/
 * clearEditingRestriction`.
 *
 * ## הבאג שהמודול הזה תוקן בגללו: ה-XML נכתב, שום דבר לא נחסם
 *
 * `setEditingRestriction({mode:'readOnly'})` כותב `w:documentProtection`
 * נכון — Word יכבד אותו כשהקובץ ייפתח מחדש — אבל בתוך הסשן החי שום דבר לא
 * נשען עליו. נמדד בדפדפן: אחרי ההפעלה `protection.get()` מחזיר
 * `enforced:true`, אבל `ui.commands.get('bold').getState()` עדיין
 * `enabled:true`, והקלדה נכתבת ל-`word/document.xml`.
 *
 * הסיבה: שער החסימה של המנוע (`unsupportedRuntimeRejection` ב-superdoc,
 * ודומיו) שוער **אך ורק** לפי `documentMode === 'viewing'` —
 * `editingRestriction.enforced` אינו נבדק באף מסלול חסימה. `document-mode`
 * הוא גם הדגל היחיד שהאפליקציה עצמה קוראת בשביל `isDocumentEditable`
 * (App.vue). כלומר יש כבר מסלול חסימה עובד — הוא רק לא מחובר להגנה.
 *
 * לכן `enableReadOnlyProtection`/`disableProtection` מקבלות גם `CommandAdapter`
 * ומעבירות את `document-mode` יחד עם כתיבת ה-XML: הפעלה → `'viewing'`,
 * ביטול → המצב שהיה לפני ההפעלה (`previousMode` שההפעלה מחזירה), כדי שמסמך
 * שהיה במעקב שינויים (`'suggesting'`) יחזור למעקב ולא ל-`'editing'` גורף.
 * `commands: null` (למשל בבדיקה שאינה עוסקת בכך) משאיר רק את כתיבת ה-XML,
 * ומחזיר ל-`'editing'` בברירת מחדל בביטול.
 *
 * **הסדר בין שני הצעדים הפוך בין הפעלה לביטול, ולא סתם משום סימטריה:**
 * כש-`document-mode === 'viewing'` ה-Document API כולו חוסם מוטציות — כולל
 * את `clearEditingRestriction` עצמה (`PERMISSION_DENIED`, "read-only review
 * mode") — נמדד בדפדפן. לכן ההפעלה כותבת XML **ואז** עוברת ל-`'viewing'`,
 * והביטול עובר מ-`'viewing'` **ואז** כותב את ה-XML.
 *
 * ## סדר המדידה שקבע את הבדיקה המקורית
 *
 * המדריך דרש למדוד את **מסלול הביטול לפני מסלול ההפעלה** — וכך נעשה:
 * `setEditingRestriction({mode:'readOnly'})` → `enforced:true`;
 * `capabilities.get()` אחרי ההפעלה: **4 פעולות נפלו ל-false**
 * (התוסף עצמו מוגבל!); `clearEditingRestriction()` → `enforced:false`
 * והמסמך חוזר לגמרי. הביטול עובד בלי סיסמה (v1 אינה תומכת סיסמה).
 *
 * לכן הפקד: מתג עם אישור מפורש, וההסבר אומר בדיוק מה יקרה — המסמך
 * יינעל לקריאה בלבד, 4 פעולות ייפלו, וניתן לבטל מהמתג עצמו.
 */
import type { SuperDoc } from 'superdoc';
import { reasonText, type CommandAdapter, type CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt } from './document-api';

const UNAVAILABLE_TEXT = 'אינו זמין בגרסה זו';

/** הפקודה שהמנוע גוזר ממנה חסימה בפועל — ראו הערת הראש. */
const DOCUMENT_MODE_COMMAND = 'document-mode';

/** ברירת המחדל לביטול, כשאין `previousMode` (למשל מסמך שנטען מוגן כבר). */
const DEFAULT_MODE = 'editing';

const MODE_CHANGE_UNAVAILABLE: Extract<CommandOutcome, { ok: false }> = {
  ok: false, message: reasonText('not-ready'), reason: 'not-ready',
};

/** Reading/writing protection and applying its mode cannot straddle an open/Skip. */
async function withModeChange<T>(
  commands: CommandAdapter | null,
  change: () => Promise<T>,
  unavailable: T,
): Promise<T> {
  const release = commands?.beginModeChange?.();
  if (commands?.beginModeChange && !release) {
    return unavailable;
  }
  try {
    return await change();
  } finally {
    release?.();
  }
}

export interface ProtectionState {
  mode: string;
  enforced: boolean;
  /**
   * האם המנוע **בפועל** חוסם קלט לפי ההגבלה, לא רק כותב אותה ל-XML. ראו
   * הערת הראש: נמדד `true` גם כשדבר לא היה חסום בפועל, ולכן השדה הזה
   * מוצג לצרכנים אך אינו הבסיס לחסימה כאן — היא נשענת על `document-mode`.
   */
  runtimeEnforced: boolean;
}

function currentDocumentMode(commands: CommandAdapter | null): string | undefined {
  if (!commands || !commands.has(DOCUMENT_MODE_COMMAND)) return undefined;
  const value = commands.getState(DOCUMENT_MODE_COMMAND).value;
  return typeof value === 'string' ? value : undefined;
}

interface ProtectionApiShape {
  protection?: {
    get?: () => MaybePromise<unknown>;
    setEditingRestriction?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
    clearEditingRestriction?: () => MaybePromise<DocReceipt>;
  };
}

export interface ProtectionHost {
  activeEditor?: { doc?: ProtectionApiShape | null } | null;
}

export type ProtectionTarget = SuperDoc | ProtectionHost | null | undefined;

type MaybePromise<T> = T | Promise<T>;

function docOf(host: ProtectionTarget): ProtectionApiShape | null {
  return (host as ProtectionHost | null | undefined)?.activeEditor?.doc ?? null;
}

/** מצב ההגנה הנוכחי. `null` = אין מנוע. */
export async function readProtectionState(host: ProtectionTarget): Promise<ProtectionState | null> {
  const get = docOf(host)?.protection?.get;
  if (typeof get !== 'function') return null;
  try {
    const raw = (await get()) as
      | { editingRestriction?: { mode?: string; enforced?: boolean; runtimeEnforced?: boolean } }
      | undefined;
    return {
      mode: raw?.editingRestriction?.mode ?? 'none',
      enforced: raw?.editingRestriction?.enforced === true,
      runtimeEnforced: raw?.editingRestriction?.runtimeEnforced === true,
    };
  } catch {
    return null;
  }
}

/** תוצאת ההפעלה: `previousMode` הוא מה שהביטול צריך להעביר אליו בחזרה. */
export type EnableProtectionOutcome =
  | { ok: true; previousMode: string }
  | { ok: false; message: string; reason?: string };

/**
 * הפעלת הגנה „קריאה בלבד" על המסמך כולו.
 *
 * שני צעדים ולא אחד: כתיבת ה-XML (`setEditingRestriction`), **וגם** מעבר
 * `document-mode` ל-`'viewing'` — המסלול היחיד שהמנוע בפועל שוער לפיו קלט
 * ופקודות (ראו הערת הראש). בלי הצעד השני ה-XML נכתב נכון וה-UI מציג מנעול,
 * אבל אין שום חסימה בסשן החי.
 *
 * `commands: null` מדלג על הצעד השני (רק כתיבת XML) — משמש בדיקה שבודקת
 * את מסלול ה-Document API בבידוד.
 */
export function enableReadOnlyProtection(
  host: ProtectionTarget,
  commands: CommandAdapter | null,
): Promise<EnableProtectionOutcome> {
  return withModeChange<EnableProtectionOutcome>(commands, () => enableProtection(host, commands), MODE_CHANGE_UNAVAILABLE);
}

async function enableProtection(host: ProtectionTarget, commands: CommandAdapter | null): Promise<EnableProtectionOutcome> {
  const failedAction = 'הפעלת הגנת המסמך נכשלה';
  const set = docOf(host)?.protection?.setEditingRestriction;
  if (typeof set !== 'function') {
    return { ok: false, message: `${failedAction}: ${UNAVAILABLE_TEXT}`, reason: 'command-unsupported' };
  }

  const previousMode = currentDocumentMode(commands) ?? DEFAULT_MODE;
  let receipt: DocReceipt;
  try {
    receipt = await set({ mode: 'readOnly' });
  } catch (error) {
    return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
  }

  // NO_OP = כבר מוגן. הצלחה מבחינת המשתמש.
  if (receipt?.success === false && receipt.failure?.code !== 'NO_OP') {
    return { ok: false, message: receiptFailureText(failedAction, receipt), reason: receipt.failure?.code };
  }

  // נשמר **לפני** שמעבירים ל-'viewing', כדי שהביטול ישחזר בדיוק את מה שהיה —
  // מסמך שהיה במעקב שינויים ('suggesting') לא אמור לצאת ממנו רק בגלל שהוגן.
  if (commands && previousMode !== 'viewing') {
    const modeOutcome = await commands.run(DOCUMENT_MODE_COMMAND, { mode: 'viewing' });
    if (!modeOutcome.ok) {
      // ה-XML כבר נכתב; החסימה בפועל היא זו שנכשלה. מדווחים ככשל — הבטחת
      // הפקד היא "המסמך יינעל", לא "ה-XML יכתב".
      return { ok: false, message: modeOutcome.message, reason: modeOutcome.reason };
    }
  }

  return { ok: true, previousMode };
}

/**
 * ביטול ההגנה — המסלול שנמדד ראשון.
 *
 * `restoreMode` הוא מה ש-`enableReadOnlyProtection` החזירה כ-`previousMode`;
 * ברירת המחדל (`'editing'`) משמשת רק כשאין ערך שמור (למשל מסמך שנטען מוגן
 * מלכתחילה, ראו `syncProtectionRuntime`).
 *
 * **הסדר כאן הפוך מ-`enableReadOnlyProtection` בכוונה: קודם `document-mode`,
 * ורק אחר-כך `clearEditingRestriction`.** נמדד בדפדפן (QA gate a): כשהמסמך
 * ב-`'viewing'`, ה-Document API **כולו** חוסם מוטציות — כולל את הקריאה
 * שאמורה לבטל את הנעילה עצמה — ומחזיר `PERMISSION_DENIED` ("is unavailable
 * while the document is in read-only review mode"). קריאה ל-`clear()` לפני
 * שהמסמך יצא מ-`'viewing'` נכשלת תמיד, כלומר "בטל" לא היה מבטל כלום.
 * שחרור המצב קודם פותח את השער, ורק אז ה-XML מתעדכן.
 */
export function disableProtection(
  host: ProtectionTarget,
  commands: CommandAdapter | null,
  restoreMode: string = DEFAULT_MODE,
): Promise<CommandOutcome> {
  return withModeChange<CommandOutcome>(commands, () => clearProtection(host, commands, restoreMode), MODE_CHANGE_UNAVAILABLE);
}

async function clearProtection(
  host: ProtectionTarget,
  commands: CommandAdapter | null,
  restoreMode: string,
): Promise<CommandOutcome> {
  const failedAction = 'ביטול הגנת המסמך נכשל';
  const clear = docOf(host)?.protection?.clearEditingRestriction;
  if (typeof clear !== 'function') {
    return { ok: false, message: `${failedAction}: ${UNAVAILABLE_TEXT}`, reason: 'command-unsupported' };
  }

  // ראו הערת הראש: אם לא יוצאים מ-'viewing' לפני הקריאה ל-clear(), הקריאה
  // עצמה נחסמת על ידי אותו שער שההגנה הפעילה.
  if (commands && currentDocumentMode(commands) === 'viewing') {
    const modeOutcome = await commands.run(DOCUMENT_MODE_COMMAND, { mode: restoreMode });
    if (!modeOutcome.ok) {
      return { ok: false, message: modeOutcome.message, reason: modeOutcome.reason };
    }
  }

  let receipt: DocReceipt;
  try {
    receipt = await clear();
  } catch (error) {
    return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
  }

  if (receipt?.success === false && receipt.failure?.code !== 'NO_OP') {
    return { ok: false, message: receiptFailureText(failedAction, receipt), reason: receipt.failure?.code };
  }

  return { ok: true };
}

/**
 * מוודאת שמצב הריצה תואם את מה שכתוב במסמך — למקרה שהוא **נטען כשהוא כבר
 * מוגן** (ה-XML נושא `w:documentProtection` מבעוד מועד, לא מפעולה בסשן הזה).
 * בלי הסנכרון הזה מסמך כזה נפתח עם התג "נעול" דלוק ושום דבר לא חסום —
 * בדיוק הבאג הזה, רק בנקודת כניסה אחרת.
 */
export function syncProtectionRuntime(
  host: ProtectionTarget,
  commands: CommandAdapter | null,
): Promise<ProtectionState | null> {
  return withModeChange<ProtectionState | null>(commands, async () => {
    const state = await readProtectionState(host);
    if (state?.enforced && commands && currentDocumentMode(commands) !== 'viewing') {
      await commands.run(DOCUMENT_MODE_COMMAND, { mode: 'viewing' });
    }
    return state;
  }, null);
}
