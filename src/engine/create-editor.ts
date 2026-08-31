/**
 * הקמת מנוע SuperDoc v2 במצב "מנוע בלבד" — `ui: false`.
 * הממשק כולו שלנו; SuperDoc אחראי למודל המסמך, לעימוד, ל-DOCX ולייצוא.
 *
 * חוזה הבעלות על ה-controller (superdoc@2.8.0):
 * המופע כבר מחזיק controller ומחזיר אותו ב-`superdoc.ui`, בטיפוס
 * `BorrowedSuperDocUI` — כלומר `Omit<SuperDocUI, 'destroy'>`. אין לקרוא
 * ל-`createSuperDocUI` מקוד התוסף: זה יוצר controller שני, בבעלותנו, שהמנוע
 * לא יודע עליו. הפירוק נעשה ב-`superdoc.destroy()` בלבד — הוא זה שמפרק גם את
 * ה-controller.
 */
import { SuperDoc } from 'superdoc';
import type { BorrowedSuperDocUI, SuperDocExceptionPayload } from 'superdoc';
import 'superdoc/style.css';
// אחרי גיליון המנוע, ובכוונה: הכללים שם מעברתים את שכבת הכותרות שהוא מצייר.
import '../styles/engine-chrome.css';
import { splashStage, SPLASH_STAGES } from '../host/splash';
import { installFormatPainter } from './format-painter';
import { localizeEngineChrome } from './hf-chrome';
import { installWordSelection } from './word-selection';
import { engineWorkerUrls } from './workers';

export interface EditorSession {
  superdoc: SuperDoc;
  /** ה-controller של המופע. מושאל — לא לפרק אותו, ואין לו `destroy`. */
  ui: BorrowedSuperDocUI;
  /** ה-container שהמנוע רינדר לתוכו. לשכבות חיצוניות שמותקנות/מפורקות מחוץ
   * ל-create-editor (למשל book-completion-overlay.ts, תלוי-טוגל). */
  container: HTMLElement;
  /** רושם ביטול של subscription שלנו. ירוץ ב-`destroy`, לפני פירוק המנוע. */
  onDispose(dispose: () => void): void;
  /** מבטל את ה-subscriptions שלנו ואז מפרק את המנוע. אידמפוטנטי. */
  destroy(): void;
}

export interface CreateEditorOptions {
  /** האלמנט שבתוכו SuperDoc מרנדר את המסמך. */
  container: HTMLElement;
  /**
   * המסמך לפתיחה: URL (מה-loopback של אוצריא) או File/Blob.
   * בלי מסמך נפתח מסמך ריק.
   */
  source?: string | File | Blob;
  /** נקרא על כל exception של המנוע, גם אחרי שהמסמך נטען. */
  onError?: (error: Error, payload: SuperDocExceptionPayload) => void;
  /** נקרא על כל שינוי במסמך. זה מה שמסמן אותו כלא-שמור. */
  onUpdate?: () => void;
  /**
   * מספר עמודי הפריסה, אחרי כל מעבר עימוד.
   *
   * זה **המקור היחיד** למספר העמודים: אין getter ציבורי לשאול בו „כמה עמודים
   * יש”, ו-`doc.info().counts.pages` מתועד כמי שנעדר „when pagination is
   * inactive or layout hasn't completed”. עמודים אינם ידועים לפני שהפריסה
   * רצה, ולכן מי שאינו מאזין לזה יכול רק להמציא מספר — וזה מה שקרה: שורת
   * המצב הציגה „עמוד 1 מתוך 1” על כל מסמך.
   *
   * ה-payload של המנוע נושא גם את המופע עצמו; הוא אינו מועבר הלאה — מי
   * שרושם את ה-callback מחזיק את ה-session ואינו צריך לקבל אותו בחזרה.
   */
  onPaginationUpdate?: (totalPages: number) => void;
  /** מעל הזמן הזה הפתיחה נכשלת. ראו OPEN_TIMEOUT_MS. */
  timeoutMs?: number;
}

/**
 * גבול הזמן לפתיחת מסמך.
 *
 * לא הגנה מפני איטיות אלא מפני שקט: `onReady` ו-`onException` הם שני המסלולים
 * היחידים שמסיימים את ההבטחה, ומסלול במנוע שלא יורה אף אחד מהם מקפיא את
 * הממשק בלי שום סימן. קרה בפועל עם דיאלוג הסיסמה המובנה. הערך נדיב ביחס
 * ל-boot שנמדד (485ms ארוז, 4.3 שניות בפיתוח) ומול workerStartupTimeoutMs
 * של המנוע (30 שניות).
 *
 * **מה הוא אינו מכסה, ואסור להניח שכן:** מנוע שנתקע בלולאה על החוט הראשי.
 * `setTimeout` צריך בדיוק את החוט החסום, ולכן הוא לא יירה — הפתיחה תישאר
 * תלויה לנצח, בלי שגיאה ובלי דרך לצאת. זה אינו תיאורטי: `w:defaultTabStop`
 * של אפס עושה בדיוק את זה (ראו docs/engine-gaps.md). הגנה כזאת חייבת לקרות
 * לפני שהמנוע רואה את הבייטים — engine/docx-preflight.ts.
 */
export const OPEN_TIMEOUT_MS = 120_000;

/**
 * ההודעה שמגיעה מ-exception של SuperDoc. ה-union מגיע מארבעה מקומות שונים
 * במנוע ואינו מנורמל, ולכן קוראים אותו בהגנה במקום להניח שדה.
 */
export function exceptionToError(payload: SuperDocExceptionPayload): Error {
  const raw: unknown = payload?.error;
  if (raw instanceof Error) return raw;
  if (typeof raw === 'string' && raw.trim() !== '') return new Error(raw);

  const code = payload && 'code' in payload ? payload.code : undefined;
  if (typeof code === 'string' && code !== '') {
    return new Error(`המנוע דיווח על שגיאה (${code})`);
  }
  return new Error('טעינת המסמך נכשלה');
}

export function createEditor(options: CreateEditorOptions): Promise<EditorSession> {
  // התחנה מדווחת כאן ולא ממי שקרא, ובכוונה: כאן המנוע כבר בזיכרון, וזה הרגע
  // שבו „פותח את המסמך” נכון. אחרי הפתיחה הראשונה מסך הטעינה סגור וזה no-op.
  splashStage(SPLASH_STAGES.documentOpening, 'פותח את המסמך…');

  const {
    container,
    source,
    onError,
    onUpdate,
    onPaginationUpdate,
    timeoutMs = OPEN_TIMEOUT_MS,
  } = options;

  return new Promise((resolve, reject) => {
    let instance: SuperDoc | undefined;
    let settled = false;
    let timer: ReturnType<typeof setTimeout> | undefined;
    /** כשל שהגיע לפני שהבנאי חזר, ולכן לא היה מה לפרק באותו רגע. */
    let pendingTeardown = false;

    const disposers: Array<() => void> = [];
    let destroyed = false;

    function destroy(target: SuperDoc): void {
      if (destroyed) return;
      destroyed = true;
      // קודם ה-subscriptions שלנו: listener שירוץ אחרי הפירוק יקרא state
      // של controller מפורק.
      for (const dispose of disposers.splice(0)) {
        try {
          dispose();
        } catch (error) {
          console.error('[otzaria-word] כשל בביטול subscription', error);
        }
      }
      target.destroy();
    }

    const superdoc = new SuperDoc({
      selector: container,
      document: source,

      // הממשק כולו שלנו — SuperDoc לא מרנדר שום toolbar, dialog או popover.
      ui: false,

      // יחידת המידה של הסרגל ושל שדות המרחק. ברירת המחדל של המנוע היא
      // `'in'` — ברירת המחדל של Word en-US — וזו הסיבה שפאנל הכותרות שלו הציג
      // „0.49 in”: 1.25 ס"מ, שהם ברירת המחדל של Word העברי (ראו
      // HEADER_DISTANCE_DEFAULT_CM ב-page-setup.ts), באינצ'ים. כל הממשק שלנו
      // מקליד ומציג סנטימטרים, ושתי יחידות באותו מסמך הן מספר שהמשתמש קורא
      // לא נכון — לא קוסמטיקה.
      measurementUnit: 'cm',

      // התוסף עובד אופליין וללא הרשאת רשת; טלמטריה תיצור קריאות שייחסמו.
      telemetry: { enabled: false },

      // דיאלוג הסיסמה המובנה של המנוע פועל גם כש-ui: false — הוא surface של
      // modules ולא של ui — והוא "לוקח אחריות" על DOCX מוצפן: הוא מטפל
      // ב-DOCX_PASSWORD_REQUIRED בעצמו ואינו פולט exception. התוצאה הייתה
      // הבטחה שאינה מסתיימת: דיאלוג באנגלית מעל הממשק שלנו, וביטול שלו משאיר
      // את הפתיחה תלויה לנצח. מכובה — כך שהכשל מגיע כ-exception ומטופל.
      // דיאלוג סיסמה בעברית הוא פיצ'ר לשלב מאוחר, לא תופעת לוואי.
      modules: { surfaces: { passwordPrompt: false } },

      // ב-file:// חייבים workers מ-blob: . undefined משאיר את ברירת המחדל
      // של SuperDoc, שנכונה בפיתוח מ-localhost.
      workerUrls: engineWorkerUrls(),

      // כל שינוי במסמך. ה-session מסמן ממנו dirty; אין קריאה ל-DOM.
      onEditorUpdate: onUpdate ? () => onUpdate() : undefined,

      // מספר העמודים, אחרי כל מעבר פריסה. ראו onPaginationUpdate למעלה.
      onPaginationUpdate: onPaginationUpdate
        ? ({ totalPages }) => onPaginationUpdate(totalPages)
        : undefined,

      // ה-payload נושא את המופע המוכן. משתמשים בו, ולא ב-closure, כדי לא
      // להישען על סדר ההשמה של הבנאי.
      onReady: ({ superdoc: ready }) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        const session: EditorSession = {
          superdoc: ready,
          ui: ready.ui,
          container,
          onDispose(dispose) {
            if (destroyed) {
              dispose();
              return;
            }
            disposers.push(dispose);
          },
          destroy: () => destroy(ready),
        };
        // ידית QA: סקריפטי השער ב-scripts/*.mjs (zoom-qa וכדומה) מריצים את
        // ה-dist הארוז וצריכים להגיע למופע החי; Vue ב-production אינו חושף
        // את עץ הרכיבים, ולכן אין דרך אחרת. אותו אובייקט session בדיוק —
        // לא עותק, שלא ייווצרו שני חוזי פירוק. קוד האפליקציה אינו קורא כאן.
        (window as unknown as { __otzariaEditor?: EditorSession }).__otzariaEditor = session;
        // לחיצה כפולה שבוחרת מילה שלמה גם בטקסט מנוקד, ושלוש לחיצות שבוחרות קטע —
        // ראו word-selection.ts. כאן ולא ליד `localizeEngineChrome`, מפני שזה זקוק
        // למופע עצמו: המופע מוכר רק ב-onReady, ו-`session.onDispose` הוא גם מה
        // שמטפל במקרה שהפירוק כבר רץ.
        session.onDispose(installWordSelection(container, ready).dispose);
        // מברשת עיצוב: המנוע מחיל רק כשמישהו קורא ל-notifyPointerUp/notifyKeyUp
        // של ui.formatPainter, וזה בדיוק מה ש-SuperToolbar עושה — אבל הוא לא קם
        // כש-ui: false. ראו format-painter.ts.
        session.onDispose(installFormatPainter(container, ready).dispose);
        resolve(session);
      },

      onException: (payload) => {
        const error = exceptionToError(payload);
        onError?.(error, payload);
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        // כשל לפני onReady משאיר מופע חצי-בנוי עם workers פתוחים. אם ה-exception
        // נורה מתוך הבנאי עצמו — והטיפוסים מתעדים מסלול כזה, שבו הריצה "mounts
        // only enough state to report that error" — instance עדיין undefined,
        // ולכן הפירוק נדחה לרגע שאחרי הבנאי.
        if (instance) destroy(instance);
        else pendingTeardown = true;
        reject(error);
      },
    });

    instance = superdoc;
    if (pendingTeardown) destroy(superdoc);

    // עברות שכבת הכותרות שהמנוע מצייר — ראו hf-chrome.ts ו-engine-chrome.css.
    // מותקן על ה-container ולא על המסמך, ונרשם כ-disposer: destroy מפרק את
    // ה-observer *לפני* המנוע, כדי שלא ירוץ על עץ שכבר נפרק.
    disposers.push(localizeEngineChrome(container).dispose);

    if (settled) return;
    timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      destroy(superdoc);
      reject(new Error('פתיחת המסמך לא הסתיימה בזמן סביר'));
    }, timeoutMs);
  });
}
