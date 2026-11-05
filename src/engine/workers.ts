/**
 * Workers של מנוע ה-DOCX, כ-blob: URLs.
 *
 * ה-build מטמיע את קוד ה-worker כמחרוזת (inlineEngineWorkers ב-vite.config.ts)
 * וכאן הוא הופך ל-blob: שנמסר ל-`config.workerUrls`. שלוש עובדות נמדדו
 * ב-Chromium וקובעות את הצורה הזאת (docs/spike.md §שער A):
 *
 * 1. קוד ה-workers נארז כמקורות קלאסיים. workerUrls מוסר למנוע כתובות blob
 *    שלהם, גם כשהקוד הראשי מפוצל ל-AMD; כתובת של chunk אינה מתירה module worker.
 * 2. מ-file:// (origin opaque) module worker אינו נטען — לא מ-blob ולא, בגודל
 *    הזה, מ-data. worker **קלאסי** מ-blob דווקא נטען.
 * 3. הקוד ש-Vite מפיק ל-workers הוא IIFE (`worker.format`), כלומר תואם-קלאסי:
 *    אין בו import/export ואין import.meta. נבדק בבנייה, ולא רק בהנחה.
 *
 * מכאן `asClassicWorker`: המנוע מקודד `{ type: 'module' }` בכל קריאה ל-Worker,
 * ואנחנו עוטפים את הבנאי ומסירים את האופציה — אבל **רק** ל-blob URLs שאנחנו
 * עצמנו בנינו. זו עטיפה של API של הדפדפן, לא שינוי של המנוע ולא של ה-workers
 * שלו; הרישיון אוסר לשנות אותם, וכאן הם נטענים בייט-בבייט כמו שהם.
 */
declare global {
  interface Window {
    __SUPERDOC_WORKER_SOURCES__?: Record<string, string>;
    /**
     * ה-blob URL של worker המסמך, כשהטוען שב-index.html כבר בנה אותו לחימום
     * מוקדם (ראו deferredEntry ב-vite.config.ts — השם משותף איתו). אימוץ
     * ה-URL הזה, ולא בנייה של חדש מאותו תוכן, הוא מה שמפגיש את ה-Worker של
     * המנוע עם הקומפילציה החמה: ה-cache ממופתח לפי URL.
     */
    __otzariaDocWorkerUrl?: string;
  }
}

/**
 * מפתחות של `Config.workerUrls`. שני התפקידים שהתוסף צריך: המסמך עצמו,
 * וה-index של הערות ו-track changes. `collaboration` אינו נארז — התוסף
 * עובד אופליין וללא הרשאת רשת.
 */
export interface EngineWorkerUrls {
  document?: string;
  reviewIndex?: string;
}

const ROLES = ['document', 'reviewIndex'] as const;

let cached: EngineWorkerUrls | undefined;
let patched = false;

/** ה-blob URLs שאנחנו בנינו. רק הם מקבלים את הטיפול הקלאסי. */
const ownUrls = new Set<string>();

/**
 * מסיר `type: 'module'` מבנאי ה-Worker — רק ל-URL שאנחנו בנינו.
 *
 * מ-file:// אין דרך אחרת: module worker נחסם שם, וקוד ה-worker באריזה הוא
 * IIFE ולכן רץ כקלאסי. כל URL אחר עובר לבנאי המקורי בדיוק כפי שהתקבל, כדי
 * שקוד אחר בדף לא יושפע.
 */
function patchWorkerConstructor(): void {
  if (patched) return;

  const Native = window.Worker;
  // סביבה בלי Worker (למשל jsdom בבדיקות) — אין מה לעטוף, ואין לזרוק.
  if (typeof Native !== 'function') return;

  class OtzariaWorker extends Native {
    constructor(scriptUrl: string | URL, options?: WorkerOptions) {
      if (options?.type === 'module' && ownUrls.has(String(scriptUrl))) {
        const { type: _module, ...classic } = options;
        super(scriptUrl, classic);
        return;
      }
      super(scriptUrl, options);
    }
  }

  window.Worker = OtzariaWorker;
  patched = true;
}

export function engineWorkerUrls(): EngineWorkerUrls | undefined {
  if (cached) return cached;

  const sources = window.__SUPERDOC_WORKER_SOURCES__;
  if (!sources) return undefined;

  const urls: EngineWorkerUrls = {};
  const created: string[] = [];
  try {
    for (const role of ROLES) {
      const code = sources[role];
      // תפקיד חסר נשאר undefined בכוונה: SuperDoc ייפול חזרה ל-URL המובנה שלו
      // במקום לקבל blob: ריק שייכשל בטעינה.
      if (typeof code === 'string' && code !== '') {
        // ה-URL שהטוען כבר בנה לחימום קודם לבנייה חדשה — ראו את ההצהרה למעלה.
        const warmUrl = role === 'document' ? window.__otzariaDocWorkerUrl : undefined;
        let url: string;
        if (typeof warmUrl === 'string' && warmUrl.startsWith('blob:')) {
          url = warmUrl;
        } else {
          url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' }));
          created.push(url);
        }
        ownUrls.add(url);
        urls[role] = url;
      }
    }

    if (ownUrls.size > 0) patchWorkerConstructor();
  } catch (error) {
    // אף עורך לא קיבל את הכתובות החדשות. החימום, לעומתן, עשוי להיות פעיל;
    // משאירים אותו ואת המקורות לניסיון חוזר, בלי לצבור blobs יתומים.
    for (const url of created) {
      ownUrls.delete(url);
      URL.revokeObjectURL(url);
    }
    throw error;
  }

  cached = urls;
  // ה-Blob URLs נשארים חיים לכל העורכים הבאים. רק העותק המחרוזתי המיותר
  // משתחרר, ורק לאחר השלמת המטמון בהצלחה.
  delete window.__SUPERDOC_WORKER_SOURCES__;
  return urls;
}

/**
 * לבדיקות בלבד: משחרר את ה-blob URLs ומאפס את המצב, כולל דגל העטיפה — כך שכל
 * בדיקה עוטפת את ה-Worker שהיא עצמה התקינה.
 */
export function resetEngineWorkerUrlsCache(): void {
  // ה-URL המאומץ מהטוען עלול להיות בין ה-ownUrls שנשרפים כאן; מחיקת הגלובל
  // מונעת אימוץ חוזר של URL שבוטל.
  if (typeof window !== 'undefined') delete window.__otzariaDocWorkerUrl;
  for (const url of ownUrls) URL.revokeObjectURL(url);
  ownUrls.clear();
  cached = undefined;
  patched = false;
}
