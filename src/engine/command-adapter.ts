/**
 * שכבת הפקודות. כל פקד ב-Ribbon עובר דרך כאן ולא קורא ל-`ui.commands` ישירות,
 * כדי שכשל יגיע למשתמש בעברית ולא ייעלם בשקט.
 *
 * האדפטר מחזיר תוצאה, ואינו מציג הודעה בעצמו: מי שיודע איך להציג הודעה
 * למשתמש הוא שכבת ה-Host, וערבוב שלה לכאן היה מונע בדיקה של השכבה הזאת
 * ומכריח כל צרכן להודעה מסוג אחד.
 */
import type { BorrowedSuperDocUI } from 'superdoc';
import type { CommandExecutionResult, CommandState, SuperDocUIReason } from 'superdoc/ui';

/**
 * תוצאת פקודה. `note` הוא ערוץ ל**הצלחה שיש עליה מה לומר** — פעולה שבוצעה
 * במלואה, אבל המנוע יצייר את התוצאה אחרת ממה שהמשתמש מצפה. בלעדיו יש רק שתי
 * אפשרויות, „שקט” או „נכשלה”, ושתיהן שקר: הפעולה לא נכשלה, ושתיקה משאירה את
 * המשתמש מול תצוגה שנראית שבורה בלי הסבר.
 */
export type CommandOutcome =
  | { ok: true; note?: string }
  | { ok: false; message: string; reason?: string };

/**
 * ה-reason של ה-controller — למה הפקודה חסומה. `Record` ולא `Partial<Record>`
 * בכוונה: מזהה חדש בגרסת superdoc עתידית יפיל את ה-typecheck במקום להגיע
 * למשתמש כהודעה גנרית.
 *
 * מיוצאת מפני ששכבת החיפוש (engine/search.ts) מתרגמת את אותם reasons; שתי
 * טבלאות היו נותנות שני נוסחים עבריים לאותו כשל.
 */
export const REASON_TEXT: Record<SuperDocUIReason, string> = {
  'not-ready': 'המסמך עדיין נטען',
  'document-api-unavailable': 'המסמך עדיין נטען',
  'document-readonly': 'המסמך פתוח לקריאה בלבד',
  'selection-required': 'יש למקם את הסמן במסמך',
  'range-selection-required': 'יש לסמן טקסט תחילה',
  'context-unavailable': 'הפעולה אינה זמינה במקום הזה במסמך',
  'geometry-unavailable': 'לא ניתן לאתר את המקום במסמך',
  'target-unresolved': 'לא ניתן לזהות את היעד של הפעולה',
  'target-not-visible': 'היעד של הפעולה אינו מוצג במסך',
  'command-unsupported': 'הפעולה אינה נתמכת בגרסה הזאת של המנוע',
  'command-deferred': 'הפעולה אינה נתמכת בגרסה הזאת של המנוע',
  'table-context-unavailable': 'יש למקם את הסמן בתוך תא בטבלה',
  'operation-unavailable': 'הפעולה אינה זמינה כרגע',
  'bulk-decisions-disabled': 'קבלה או דחייה של כל השינויים אינה מאופשרת',
  'tracked-change-decisions-disabled': 'קבלה או דחייה של שינויים אינה מאופשרת',
  'host-capability-unavailable': 'היכולת הדרושה לפעולה אינה זמינה',
  'history-empty': 'אין פעולה לבטל',
  'search-unavailable': 'החיפוש אינו זמין במסמך הזה',
  'search-invalid-pattern': 'תבנית החיפוש אינה חוקית',
  'search-truncated': 'יש יותר מדי תוצאות לספור את כולן, ולכן לא הוחלף דבר',
  'replace-unsupported': 'החלפת טקסט אינה נתמכת בגרסה הזאת של המנוע',
  'content-control-locked': 'החלק הזה במסמך מוגן מפני שינוי עיצוב',
  'permission-denied': 'אין הרשאה לבצע את הפעולה',
};

/**
 * קודי כשל שמגיעים ב-receipt של ה-Document API. אוצר המילים המלא מונה
 * עשרות קודים פנימיים; מתורגמים רק אלה שיש למשתמש מה לעשות איתם. כל השאר
 * מוצגים עם הקוד עצמו, כדי שאפשר יהיה לדווח עליהם — ולא נעלמים.
 */
/**
 * מיוצאת כדי ש-engine/document-api.ts יציג את אותו נוסח: המשתמש פוגש את אותו
 * כשל בשני מסלולים — פקודה מנותבת ו-Document API ישיר — וקול שני לאותו כשל
 * הוא באג בממשק.
 */
export const FAILURE_TEXT: Record<string, string> = {
  DOCUMENT_READONLY: 'המסמך פתוח לקריאה בלבד',
  NO_SELECTION: 'יש למקם את הסמן במסמך',
  PARTIAL_LINK_EDIT: 'הקישור עודכן אך הטקסט המוצג לא הוחלף',
  PERMISSION_DENIED: 'אין הרשאה לבצע את הפעולה',
  CAPABILITY_UNAVAILABLE: 'הפעולה אינה זמינה במסמך הזה',
  CAPABILITY_UNSUPPORTED: 'הפעולה אינה נתמכת בגרסה הזאת של המנוע',
  INVALID_TARGET: 'לא ניתן לבצע את הפעולה במקום הזה במסמך',
  TARGET_NOT_FOUND: 'היעד של הפעולה לא נמצא במסמך',
  LOCK_VIOLATION: 'החלק הזה במסמך מוגן מפני שינוי',
  // הנוסח נשאר כתיעוד ולשימוש engine/document-api.ts, אבל `run()` למטה
  // מיירט NO_OP לפני שהוא מגיע לטבלה הזאת — הוא אינו כשל שהמשתמש רואה.
  NO_OP: 'לא היה מה לשנות',
};

function isReceipt(
  result: CommandExecutionResult,
): result is Exclude<CommandExecutionResult, boolean> {
  return typeof result === 'object' && result !== null;
}

export function reasonText(reason: string | undefined): string {
  if (!reason) return 'הפעולה נכשלה';
  return REASON_TEXT[reason as SuperDocUIReason] ?? `הפעולה נכשלה (${reason})`;
}

export interface CommandAdapter {
  /** האם המנוע מכיר את הפקודה. */
  has(id: string): boolean;
  /** מצב הפקד — active / enabled / value. תמיד מהמנוע, לעולם לא מה-DOM. */
  getState(id: string): CommandState;
  /** מאזינה לשינויי מצב של פקד. מחזירה פונקציית ביטול. */
  observe(id: string, listener: (state: CommandState) => void): () => void;
  /** מריצה פקודה ומחזירה תוצאה עם הודעה בעברית בכשל. */
  run(id: string, payload?: unknown): Promise<CommandOutcome>;
  /** Reserve a write + runtime-mode transaction against document replacement. */
  beginModeChange?(): (() => void) | null;
}

export function createCommandAdapter(ui: BorrowedSuperDocUI): CommandAdapter {
  const getState = (id: string): CommandState => ui.commands.get(id).getState();

  return {
    has: (id) => ui.commands.has(id),
    getState,

    observe(id, listener) {
      return ui.commands.get(id).observe(listener);
    },

    async run(id, payload) {
      if (!ui.commands.has(id)) {
        // באג בקוד שלנו, לא מצב של המסמך: הפקד לא היה צריך להיבנות בכלל.
        return { ok: false, message: `הפעולה ${id} אינה מוכרת למנוע`, reason: 'unknown-command' };
      }

      let result: CommandExecutionResult;
      try {
        result = await ui.commands.executeAsync(id, payload);
      } catch (error) {
        return {
          ok: false,
          message: error instanceof Error ? error.message : 'הפעולה נכשלה',
          reason: 'threw',
        };
      }

      // false = ה-controller לא ניתב את הפקודה בכלל. ה-reason שבמצב הפקד
      // מסביר למה, ורק הוא יודע להבחין בין "אין בחירה" ל"לא נתמך".
      if (result === false) {
        const reason = getState(id).reason;
        return { ok: false, message: reasonText(reason), reason };
      }

      if (isReceipt(result) && result.success === false) {
        const { code, message } = result.failure;
        // NO_OP = הערכים המבוקשים כבר מוגדרים. זו אינה כשל מבחינת המשתמש —
        // הוא ביקש "הזחה 0" ו"RTL" והם כבר כך. הצגת שגיאה כאן היא בדיוק מה
        // שגרם לדיווח "הפעולה הכי טבעית במסמך עברי מציגה שגיאה" (בדיקת 9 בסקר
        // הפקדים). כל שאר מודולי ה-Document API הישיר (page-setup.ts,
        // header-footer.ts, footnotes.ts ועוד) כבר מכריעים כך; מסלול הפקודות
        // המנותב חייב להסכים איתם, אחרת אותו כשל מתנהג אחרת בשני המסלולים.
        if (code === 'NO_OP') return { ok: true };
        // אין תרגום? מציגים את ההסבר של המנוע ואת הקוד. ההסבר באנגלית, אבל הוא
        // אומר משהו — ובלעדיו נשארת רק הודעה גנרית שאי אפשר לעשות איתה כלום.
        const fallback = message ? `הפעולה נכשלה: ${message} (${code})` : `הפעולה נכשלה (${code})`;
        return { ok: false, message: FAILURE_TEXT[code] ?? fallback, reason: code };
      }

      return { ok: true };
    },
  };
}
