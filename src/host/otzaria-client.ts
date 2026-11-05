/**
 * עטיפה טיפוסית סביב `window.Otzaria` — האובייקט שאוצריא מזריקה ל-WebView.
 * כל הקוד בתוסף עובר דרך כאן ולא נוגע ב-window ישירות, כדי שיהיה מקום אחד
 * לטיפול בשגיאות, ל-stub של פיתוח בדפדפן, ולהמתנה ל-plugin.boot.
 *
 * הטיפוסים מגיעים מ-src/types/otzaria_plugin.d.ts — העתק verbatim של ה-d.ts
 * הרשמי מ-docs/plugin-sdk של אוצריא. לעדכן משם, לא לערוך ידנית.
 */
import type {
  BootPayload,
  OtzariaEventMap,
  OtzariaGlobal,
  ThemePayload,
} from '../types/otzaria_plugin';
import { beginPerformanceApi } from '../diagnostics/performance';

/** ב-WebView של אוצריא ה-SDK תמיד קיים; בדפדפן רגיל הוא עשוי לא להיות. */
function bridge(): OtzariaGlobal {
  const sdk = (window as Partial<Window>).Otzaria;
  if (!sdk) throw new Error('ה-SDK של אוצריא אינו זמין — התוסף נטען מחוץ לאוצריא?');
  return sdk;
}

export function isAvailable(): boolean {
  return Boolean((window as Partial<Window>).Otzaria);
}

/**
 * שגיאה שחזרה מאוצריא, עם ה-`code` שלה לצד ההודעה.
 *
 * למה נדרש: ה-envelope של ה-SDK מפריד בין `code` (`error.permission_denied`)
 * ל-`message`, וההודעה היא טקסט חופשי שאינו מבטיח להזכיר את הקוד. בלי הקוד,
 * זיהוי „ההרשאה חסרה” נעשה בחיפוש מחרוזת בהודעה — כלומר תלוי בנוסח שאוצריא
 * בחרה, ומשתנה בלי התראה. ההודעה נשארת בדיוק כשהייתה, ולכן `isPermissionDenied`
 * כאן הוא המקום היחיד שמזהה את הכשל: גם `host/otzaria-reader.ts` וגם
 * `host/files.ts` נשענים עליו, ואין שתי תשובות לאותה שאלה.
 */
export class HostCallError extends Error {
  constructor(
    message: string,
    readonly code: string | null,
    readonly method: string,
  ) {
    super(message);
    this.name = 'HostCallError';
  }
}

/** הקוד שאוצריא נתנה, או `null` כשהשגיאה לא באה ממנה. */
export function hostErrorCode(error: unknown): string | null {
  return error instanceof HostCallError ? error.code : null;
}

/**
 * האם הכשל הוא הרשאה חסרה.
 *
 * `endsWith` ולא השוואה מדויקת: התיעוד כותב `error.permission_denied` בגוף
 * הטקסט ו-`permission_denied` בטבלת ה-RPC bridge, ושתי הצורות חוזרות בשטח.
 * ההודעה נבדקת כגיבוי, למקרה שהשגיאה נוצרה במסלול שאינו `call`.
 */
export function isPermissionDenied(error: unknown): boolean {
  const code = hostErrorCode(error);
  if (code) return code.endsWith('permission_denied');
  return error instanceof Error && error.message.includes('permission_denied');
}

/**
 * קריאה ל-Host API. זורקת שגיאה עם ההודעה שהגיעה מאוצריא.
 *
 * ממתינה לאתחול כשה-SDK עדיין אינו מוכן — ראו `untilHostReady`.
 */
export async function call<T>(method: string, payload?: Record<string, unknown>): Promise<T> {
  if (isUnbootedStub(bridge())) await untilHostReady();
  return callNow<T>(method, payload);
}

/** `call` בלי ההמתנה לאתחול. ל-`recoverBoot` בלבד: הוא זה שמגלה שהאתחול כבר קרה. */
async function callNow<T>(method: string, payload?: Record<string, unknown>): Promise<T> {
  const measured = beginPerformanceApi('host', method, payload);
  let res;
  try {
    res = await bridge().call<T>(method, payload);
  } catch (error) {
    measured?.(undefined, 'error');
    throw error;
  }
  measured?.(res, res.success ? 'ok' : 'error');
  if (!res.success) {
    throw new HostCallError(
      res.error?.message ?? `הקריאה ל-${method} נכשלה`,
      res.error?.code ?? null,
      method,
    );
  }
  return res.data as T;
}

/** כמו call, אבל מחזירה null במקום לזרוק — לשימושים לא-קריטיים. */
export async function tryCall<T>(
  method: string,
  payload?: Record<string, unknown>,
): Promise<T | null> {
  try {
    return await call<T>(method, payload);
  } catch {
    return null;
  }
}

/** נרשמת לאירוע ומחזירה פונקציית ביטול. `off` דורש בדיוק את אותה הפניה. */
export function on<K extends keyof OtzariaEventMap>(
  event: K,
  callback: (detail: OtzariaEventMap[K]) => void,
): () => void {
  const sdk = bridge();
  sdk.on(event, callback);
  return () => sdk.off(event, callback);
}

/**
 * השם שה-latch ב-index.html כותב אליו. מוגדר כאן וב-index.html, ו-
 * tests/unit/otzaria-client.test.ts קורא את ה-HTML ומקבע שהשניים זהים.
 */
export const BOOT_LATCH_KEY = '__otzariaBoot';

interface BootLatch {
  payload: BootPayload | null;
  /** `performance.now()` בזמן הירייה — כדי שאפשר יהיה למדוד את הפער. */
  at: number | null;
}

function readLatch(): BootLatch | undefined {
  return (window as unknown as Record<string, BootLatch | undefined>)[BOOT_LATCH_KEY];
}

/**
 * ה-latch של plugin.boot.
 *
 * האירוע נורה פעם אחת. אוצריא אינה שומרת את ה-payload ואין `getBootInfo`, ו-
 * `on` של ה-SDK האמיתי הוא `window.addEventListener` בלי replay — כלומר מי
 * שנרשם אחרי הירייה לא יקבל אותו לעולם.
 *
 * ההרשמה כאן, בזמן טעינת המודול, אינה מספיקה: הבאנדל הוא 10MB, ובפיתוח הוא
 * גרף מודולים שנטען ברשת, ולכן הוא עלול להיטען אחרי שהאירוע נורה. לכן ה-latch
 * האמיתי הוא סקריפט inline ב-`<head>` של index.html, שרץ בזמן פריסת ה-HTML.
 * מה שנקרא כאן הוא מה שהוא שמר. ההרשמה שלמטה נשארת לשני מצבים: ה-latch לא רץ
 * (בדיקות), או שהאירוע טרם נורה.
 */
const bootPayload = new Promise<BootPayload>((resolve) => {
  const latched = readLatch()?.payload;
  if (latched) {
    resolve(latched);
    return;
  }
  window.addEventListener(
    'plugin.boot',
    (event) => resolve((event as CustomEvent<BootPayload>).detail),
    { once: true },
  );
});

/**
 * האם אוצריא כבר אתחלה את ה-SDK בדף הזה.
 *
 * ## למה `call` צריך לדעת את זה
 *
 * אוצריא מזריקה ב-AT_DOCUMENT_START stub שדוחה **כל** קריאה ב-„Otzaria SDK
 * not ready yet”, עד ש-`_boot` מחליף אותו ב-SDK האמיתי (plugin_sdk_scripts.dart).
 * עד 6.9 העלייה הגיעה לקריאה הראשונה אחרי ה-boot במקרה — החבילה הייתה כבדה
 * מספיק. מאז שהעלייה הואצה ב-7.0, המעטפת קראה את ההעדפות ואת רשומת ההפעלה
 * כ-150ms **לפני** ה-boot (נמדד באוצריא: הקריאות ב-1042ms, plugin.boot
 * ב-1192ms). `tryCall` הפך את הדחייה ל„אין רשומה”, נפתח מסמך ריק, והרשומה
 * הריקה נכתבה מעל זו שהחזיקה את הטאבים והטיוטות.
 *
 * הגידור כאן ולא אצל כל קורא: התיעוד של `waitForBoot` כבר קבע ש„כל קריאה
 * ל-Host API חייבת לרוץ אחריה”, ושישים קוראים שכל אחד זוכר את זה בעצמו הם
 * בדיוק מה שנשבר.
 */
let hostReady = false;
let markHostReady: () => void = () => undefined;
const hostReadyPromise = new Promise<void>((resolve) => {
  markHostReady = () => {
    hostReady = true;
    resolve();
  };
});
void bootPayload.then(() => markHostReady());

/**
 * ה-stub של אוצריא שעוד לא אותחל. מזוהה לפי `_boot`, שקיים רק עליו — ה-SDK
 * בבדיקות, ב-dev-stub ובשערי ה-CDP אינו נושא אותו, ולכן אינו ממתין לכלום.
 * `_booted` הוא הדגל של אוצריא חדשה; בישנה ממתינים לאירוע עצמו.
 */
function isUnbootedStub(sdk: OtzariaGlobal): boolean {
  if (hostReady) return false;
  const stub = sdk as OtzariaGlobal & { _boot?: unknown; _booted?: unknown };
  return typeof stub._boot === 'function' && stub._booted !== true;
}

/**
 * עד האתחול, או עד `BOOT_TIMEOUT_MS`. אחרי התקרה הקריאה יוצאת בכל זאת ונכשלת
 * כמו קודם — המתנה אינסופית הייתה מחליפה כשל גלוי במסך תלוי.
 */
function untilHostReady(): Promise<void> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  return Promise.race([
    hostReadyPromise,
    new Promise<void>((resolve) => {
      timer = setTimeout(resolve, BOOT_TIMEOUT_MS);
    }),
  ]).finally(() => clearTimeout(timer));
}

/** ברירת מחדל לשעון-שומר של ה-boot. ה-Host מקומי; המתנה ארוכה היא כשל. */
export const BOOT_TIMEOUT_MS = 15_000;

/** כמה להמתין לאירוע לפני שמנסים לשחזר את ה-payload בקריאות RPC. */
export const BOOT_GRACE_MS = 2_500;

/** כל כמה זמן לחזור ולנסות את השחזור, כל עוד ה-SDK עדיין לא ענה. */
export const BOOT_POLL_MS = 1_000;

/**
 * ממתינה ל-plugin.boot. כל קריאה ל-Host API חייבת לרוץ אחריה — קריאה
 * לפני ה-boot אינה מובטחת. נכשלת בזמן קצוב במקום להשאיר מסך תלוי.
 */
export function waitForBoot(timeoutMs = BOOT_TIMEOUT_MS): Promise<BootPayload> {
  let timer: ReturnType<typeof setTimeout>;

  return Promise.race([
    bootPayload.then((payload) => {
      clearTimeout(timer);
      return payload;
    }),
    new Promise<never>((_, reject) => {
      timer = setTimeout(
        () => reject(new Error('אוצריא לא סיימה לאתחל את התוסף')),
        timeoutMs,
      );
    }),
  ]);
}

/** מה שהתוסף באמת צריך מה-boot, ומאיפה זה הגיע. */
export interface BootInfo {
  theme: ThemePayload;
  app: BootPayload['app'];
  /** `'recovered'` פירושו שהאירוע לא הגיע וה-payload נבנה מקריאות RPC. */
  source: 'event' | 'recovered';
}

/**
 * בונה את ה-boot מקריאות RPC, בשביל המצב שבו האירוע אבד.
 *
 * `app.getInfo` ו-`app.getTheme` דורשים את ההרשאה `app.info.read`, שהיא הרשאת
 * בסיס — מוענקת לכל תוסף אוטומטית, בלי הצהרה במניפסט ובלי לשאול את המשתמש.
 * כלומר המסלול הזה אינו עולה שום הרשאה חדשה.
 *
 * הקריאות נכשלות כל עוד ה-SDK האמיתי לא הוזרק (ה-stub של אוצריא דוחה הכול עם
 * „not ready yet”), וזה בדיוק הסימן שממתינים לו: הצלחה כאן פירושה ש-`_boot`
 * כבר רץ — ולכן שהאירוע נורה ואבד, ואין טעם להמתין לו עוד.
 */
export async function recoverBoot(): Promise<BootInfo> {
  const [app, theme] = await Promise.all([
    callNow<BootPayload['app']>('app.getInfo'),
    callNow<ThemePayload>('app.getTheme'),
  ]);
  if (!app || !theme) throw new Error('אוצריא החזירה מידע אתחול חסר');
  // הצלחה כאן היא ההוכחה שהאתחול קרה גם כשהאירוע אבד — ובלי השורה הזאת כל
  // קריאה אחרת הייתה ממתינה לאירוע שלא יגיע, עד התקרה.
  markHostReady();
  return { app, theme, source: 'recovered' };
}

/**
 * מה שהמעטפת קוראת לו: או האירוע, או שחזור.
 *
 * הסדר הוא העיקר. קודם ממתינים לאירוע — הוא המסלול התקין, והוא מגיע תוך
 * מילישניות. רק אחרי `graceMs` מתחיל השחזור, כדי לא להעמיס קריאות RPC על
 * מסלול שעובד. השחזור חוזר ומנסה כל `pollMs`, כי כשל שלו פירושו „ה-SDK עוד לא
 * כאן” ולא „אין טעם”. האירוע גובר בכל רגע, גם אם הגיע באיחור.
 */
export function resolveBoot(
  options: { graceMs?: number; pollMs?: number; timeoutMs?: number } = {},
): Promise<BootInfo> {
  const {
    graceMs = BOOT_GRACE_MS,
    pollMs = BOOT_POLL_MS,
    timeoutMs = BOOT_TIMEOUT_MS,
  } = options;
  const startedAt = Date.now();

  return new Promise<BootInfo>((resolve, reject) => {
    let settled = false;
    /** הודעת הכשל האחרונה מהשחזור — היא שמפרידה בין ההשערות באבחון. */
    let lastRpcError = 'טרם נוסה';
    const timers: Array<ReturnType<typeof setTimeout>> = [];

    /** תמיד גם מנקה את הטיימרים: סבב polling שנשאר חי דולף קריאות RPC. */
    function finish(action: () => void): void {
      if (settled) return;
      settled = true;
      for (const timer of timers) clearTimeout(timer);
      action();
    }

    void bootPayload.then((payload) => {
      finish(() => resolve({ theme: payload.theme, app: payload.app, source: 'event' }));
    });

    timers.push(
      setTimeout(() => {
        finish(() => {
          // האבחון נכתב ל-console מפני שאוצריא מעבירה console של תוסף ללוג
          // שלה. בלי זה, כשל אתחול אצל משתמש הוא הודעה בעברית בלי שום נתון,
          // ואי אפשר להפריד בין „האירוע אבד” ל„הגשר מת”.
          console.error('[otzaria-word] כשל אתחול', describeBootState(startedAt, lastRpcError));
          reject(
            new Error('אוצריא לא סיימה לאתחל את התוסף. נסו לטעון את הלשונית מחדש.'),
          );
        });
      }, timeoutMs),
    );

    function attempt(): void {
      if (settled) return;
      void recoverBoot().then(
        (info) => finish(() => resolve(info)),
        (error: unknown) => {
          lastRpcError = error instanceof Error ? error.message : String(error);
          if (!settled) timers.push(setTimeout(attempt, pollMs));
        },
      );
    }

    timers.push(setTimeout(attempt, graceMs));
  });
}

/**
 * תמונת מצב לאבחון כשל אתחול. כל שדה כאן מפריד בין הסברים:
 * `latchRan: false` — הסקריפט ב-`<head>` לא רץ (אריזה שבורה).
 * `latchCaught: true` והגענו לכאן — באג אצלנו, ה-payload היה ולא נקרא.
 * `sdkBooted: false` + `lastRpcError: '…not ready yet'` — `_boot` לא רץ ב-context
 *   הזה כלל. זה צד אוצריא: ה-JS context שהתוסף רץ בו אינו זה שקיבל את ה-boot
 *   (נצפה ב-macOS כשה-platform view נבנה מחדש), ו-reload הוא מה שמרפא.
 * `bridgeAlive: false` — ערוץ ה-JS↔Dart מת; אין למי לדבר.
 */
function describeBootState(startedAt: number, lastRpcError: string): Record<string, unknown> {
  const latch = readLatch();
  // `_booted` אינו בטיפוס הציבורי — אוצריא מציבה אותו ב-_boot בלבד, ולכן הוא
  // הסימן המדויק ביותר לשאלה „האם ה-SDK האמיתי הוזרק ל-context הזה”.
  const sdk = (window as Partial<Window>).Otzaria as (OtzariaGlobal & { _booted?: boolean }) | undefined;
  const bridge = (window as unknown as { flutter_inappwebview?: { callHandler?: unknown } })
    .flutter_inappwebview;

  return {
    waitedMs: Date.now() - startedAt,
    latchRan: latch !== undefined,
    latchCaught: latch?.payload != null,
    latchAt: latch?.at ?? null,
    sdkPresent: sdk !== undefined,
    sdkBooted: sdk?._booted === true,
    bridgeAlive: typeof bridge?.callHandler === 'function',
    lastRpcError,
  };
}

export function onThemeChanged(callback: (theme: ThemePayload) => void): () => void {
  return on('theme.changed', callback);
}

/** הודעות למשתמש. נכשלות בשקט — הודעה שלא נראתה אינה סיבה להפיל פעולה. */
export function notify(message: string): void {
  void tryCall('ui.showMessage', { message });
}

export function notifyError(message: string): void {
  void tryCall('ui.showError', { message });
}

/**
 * שאלת אישור של אוצריא. **דו-כפתורי**, וזו המגבלה שלו: בחירה משלושה מצבים
 * („שמור” / „לא לשמור” / „ביטול”) אינה יכולה לעבור כאן, ולכן היא נשאלת בדיאלוג
 * שלנו — `ui/panels/UnsavedChangesDialog.vue`. מה שנשאר כאן הוא שאלות כן/לא
 * אמיתיות, כמו „הקובץ השתנה מבחוץ — לפתוח את מה שלא נשמר?”.
 *
 * כשל בדיאלוג נחשב „לא”: עדיף לא לעשות פעולה הרסנית מאשר לעשות אותה בלי
 * שהמשתמש אישר.
 *
 * `subtitle` אינו כאן בכוונה — `showConfirm` בצד אוצריא מעביר `title`
 * ו-`content` בלבד, והוא קיים רק ב-`showWarning`.
 */
export async function confirm(options: {
  title: string;
  content: string;
}): Promise<boolean> {
  const res = await tryCall<{ confirmed?: boolean }>('ui.showConfirm', {
    title: options.title,
    content: options.content,
  });
  return res?.confirmed === true;
}
