/**
 * קבוצת „לוח” — העתק, גזור והדבק — ו„בחר הכל”.
 *
 * ## למה מודול, ולא שלוש שורות בקומפוננטה
 *
 * מה שהיה ב-HomeTab: `doCut` ו-`doCopy` בגוף **ריק** עם הערה שהפעולה „נתמכת
 * דרך קיצור מקלדת”, ו-`doPaste` שקורא `navigator.clipboard.readText()` ו-
 * **זורק את התוצאה** (וב-`file://` גם מייצר unhandled rejection, כי הוא
 * `void`-ed). לפני זה היו שם `document.execCommand('cut'/'copy'/'paste')`,
 * שאסורים (tests/unit/engine-boundaries: „אין עריכת מסמך דרך ה-DOM”), והתגובה
 * לאיסור הייתה למחוק את הגוף. שני הכיוונים נובעים מאותה טעות: הלוח נראה כמו
 * יכולת של הדפדפן.
 *
 * הוא אינו. תוכן מעוצב שעובר ל-Word הוא עניין של מודל המסמך, ולמנוע יש לזה
 * משטח ציבורי — `doc.clipboard` — שממיר בין `ClipboardPayload` (סדרת פריטים
 * לפי MIME) לבין fragment של המסמך. `navigator.clipboard` הוא רק הצינור אל
 * מערכת ההפעלה, ומה שנשלח בו הוא מה שהמנוע סידר.
 *
 * `doSelectAll` היה אותה טעות בצורתה החריפה: `selectAllChildren(document.body)`
 * סימן את **הרצועה, שורת המצב וכל הממשק** — ולא את המסמך. בחירה במסמך היא
 * `SelectionTarget` במודל של המנוע, ולא טווח DOM.
 *
 * ## מה אינו ודאי בזמן ריצה
 *
 * 1. **`clipboard` הוא adapter אופציונלי** (`clipboard?:` ב-
 *    `DocumentApiAdapters`). כשהוא חסר, המנוע מסמן את כל פעולות ה-namespace
 *    כ-`NAMESPACE_UNAVAILABLE`, ולכן בדיקת היכולת (engine/doc-capabilities)
 *    מספיקה כדי להחליט אם הפקד פעיל — בדיוק כמו בהערות שוליים.
 * 2. **`navigator.clipboard` עשוי להיחסם.** התוסף רץ מ-`file://` בתוך
 *    WebView2, ושם קריאת הלוח נדחית גם בתוך user-gesture. לכן יש כאן **לוח
 *    פנימי**: העתקה שהמערכת חסמה עדיין נשמרת, וההדבקה בתוך התוסף תמצא אותה.
 *    ההודעה למשתמש אומרת בדיוק מה קרה — „הצלחה” במצב הזה היא טקסט שנעלם.
 *
 *    מה שנמדד ב-Chromium על `file://`: `isSecureContext` הוא `true`, ומה שחוסם
 *    את `read()`/`readText()` הוא **ההרשאה בלבד** — `clipboard-read` במצב
 *    `prompt` נדחה מיד, וברגע שהיא מוענקת שתי הקריאות עוברות. ההענקה אינה
 *    בידינו: היא של המאכסן, ב-`onPermissionRequest` של ה-WebView. עד שהיא
 *    תגיע, כפתור „הדבק” רואה לוח חסום — ו-Ctrl+V, שהמנוע מטפל בו בתוך המסמך
 *    עם ה-`dataTransfer` של האירוע, אינו תלוי בהרשאה הזאת כלל.
 * 3. **הפאסדה בדפדפן א-סינכרונית**: `BrowserDocumentApi` עוטף כל מתודה
 *    ב-`MaybePromise`, ולכן כל קריאה כאן ב-`await` ובתוך `try`.
 *
 * ## למה הבחירה נקראת לפני הסדרוּר
 *
 * „גזור” = העתק + מחיקה, ולמחיקה (`doc.delete`) נדרש `target` מפורש — החוזה
 * מקבל `target` או `ref`, ולא כלום. לכן הבחירה נקראת **פעם אחת** מ-
 * `ui.selection`, ואותו target נשלח גם ל-`serializeSelection` וגם ל-`delete`:
 * שתי קריאות נפרדות ל„הבחירה הנוכחית” היו יכולות לתאר שני טווחים שונים,
 * כלומר להעתיק דבר אחד ולמחוק אחר.
 *
 * הטיפוסים של הלוח מוגדרים כאן ואינם מיובאים: הם יושבים תחת
 * `superdoc/dist/document-api/...`, ו-import מנתיב פנימי אסור. `SelectionTarget`
 * כן מיובא — מ-`superdoc/ui`, שהוא export ציבורי.
 */
import type { SuperDoc } from 'superdoc';
import type { SelectionTarget } from 'superdoc/ui';
import { reasonText, type CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';
import { destinationSample, marksFromDestinationCapture, marksFromEmptyDocument, type DestinationCapture } from './paste-destination';
import { PASTE_STARTED, PASTE_FINISHED } from './paste-events';
import { readFormat, type FormatReadingSource } from './format-reading';
import {
  destinationMarks,
  mergeFragment,
  textOnlyFragment,
  type PasteFragment,
  type PasteMode,
  type PasteRunMarks,
} from './paste-formatting';

/* ------------------------------------------------------------------ */
/* צורת המנוע כפי שהמודול צורך אותה                                    */
/* ------------------------------------------------------------------ */

/** פריט בודד בלוח, לפי MIME. `bytes` קיים בחוזה ואינו נוצר כאן — ראו `writeSystemClipboard`. */
export interface ClipboardPayloadItem {
  type: string;
  kind: 'string' | 'bytes';
  data: string | Uint8Array;
  name?: string;
}

/** מה שהמנוע מסדר, ומה שהוא מקבל בחזרה. `items` הוא החלק היחיד שיש לו משמעות כאן. */
export interface ClipboardPayload {
  source?: string;
  items: readonly ClipboardPayloadItem[];
}

export interface ClipboardSerializeResult {
  payload?: ClipboardPayload;
}

/**
 * התכנית ש-`parse` מחזיר ו-`insert` מקבל. שני השדות אינם עוברים כאן עיבוד —
 * הם מוצהרים מפני שהוולידטור של המנוע דורש `fragment.blocks` כמערך ו-
 * `diagnostics` כמערך, וכפיל שאינו מכריז עליהם היה מאשר תכנית פסולה.
 */
export interface ClipboardPlan {
  fragment: { blocks: readonly unknown[] };
  diagnostics: readonly unknown[];
  summary?: { sourceKind?: string; plainFallback?: boolean };
}

/** צורת הכשל של משפחת `clipboard.*`. `details.unsupportedReason` הוא ההסבר האמיתי. */
export interface ClipboardFailure {
  code?: string;
  message?: string;
  details?: { unsupportedReason?: string };
}

export type ClipboardParseResult =
  | { success: true; plan: ClipboardPlan }
  | { success: false; failure?: ClipboardFailure };

/**
 * טווח טקסט שהמוטציה **יצרה**, כפי שהוא בנתיב ה-`effects` של הקבלה.
 * `selectionTarget` הוא בדיוק הצורה שפעולות הבחירה מקבלות — החוזה מתאר אותו
 * כ„selection target covering the created span, for selection-consuming APIs”.
 */
export interface ClipboardInsertedText {
  selectionTarget?: SelectionTarget;
}

/** הקבלה של `insert`, בתוספת מה שהיא מספרת על מה שהודבק בפועל. */
export interface ClipboardInsertReceipt extends DocReceipt {
  failure?: ClipboardFailure;
  plan?: { sourceKind?: string; plainFallback?: boolean };
  diagnostics?: readonly unknown[];
  effects?: { insertedText?: readonly ClipboardInsertedText[] };
}

/** מה שנצרך מ-`activeEditor.doc`. כל שדה אופציונלי: גרסה אחרת עשויה לא לחשוף אותו. */
export interface ClipboardDocumentApi {
  info?: () => MaybePromise<{ revision?: string; counts?: { characters?: number; paragraphs?: number; tables?: number; images?: number; lists?: number; sdtFields?: number } }>;
  export?: { toDocx?: (input: Record<string, never>) => MaybePromise<{ contentBase64?: string }> };
  selection?: { extractOoxml?: (input: { at: SelectionTarget }) => MaybePromise<DestinationCapture> };
  clipboard?: {
    serializeSelection?: (input?: {
      target?: SelectionTarget;
      includeHtml?: boolean;
    }) => MaybePromise<ClipboardSerializeResult | undefined>;
    parse?: (payload: ClipboardPayload) => MaybePromise<ClipboardParseResult | undefined>;
    insert?: (input: {
      payload?: ClipboardPayload;
      plan?: ClipboardPlan;
      fragment?: ClipboardPlan['fragment'];
      target?: SelectionTarget;
    }) => MaybePromise<ClipboardInsertReceipt | undefined>;
  };
  delete?: (input: { target: SelectionTarget }) => MaybePromise<DocReceipt | undefined>;
  ranges?: {
    resolve?: (input: {
      start: DocumentEdgeAnchor;
      end: DocumentEdgeAnchor;
    }) => MaybePromise<{ target?: SelectionTarget | null } | undefined>;
  };
}

/** קצה המסמך, כעוגן ל-`ranges.resolve`. אחת משלוש צורות העוגן שהחוזה מגדיר. */
export interface DocumentEdgeAnchor {
  kind: 'document';
  edge: 'start' | 'end';
}

/**
 * מצב הבחירה כפי ש-`ui.selection.getSnapshot()` מדווח. `status` הוא מה שמפריד
 * בין „אין בחירה” לבין „הקריאה טרם הסתיימה”, ולכן הוא נקרא ולא רק ה-target.
 */
export interface SelectionSnapshot {
  status?: string;
  empty?: boolean;
  selectionTarget?: SelectionTarget | null;
}

export interface SelectionSurface {
  getSnapshot?: () => SelectionSnapshot | undefined;
  /** מחילה בחירה על העורך החי. נכשלת סגור עם `reason`, ואינה זורקת. */
  apply?: (target: SelectionTarget) => MaybePromise<{ ok?: boolean; reason?: string } | undefined>;
}

/**
 * מצבי הפקודות של המנוע — מה שהרצועה קוראת. „מזג עיצוב” קורא מהם את העיצוב
 * שבסמן (ההנמקה ב-engine/paste-formatting.ts).
 */
export interface CommandsSurface {
  has?: (id: string) => boolean;
  get?: (id: string) => { getState?: () => { value?: unknown; active?: boolean } | undefined } | undefined;
}

export interface ClipboardHost {
  element?: Element | null;
  activeEditor?: { doc?: ClipboardDocumentApi | null; container?: HTMLElement } | null;
  ui?: { selection?: SelectionSurface | null; commands?: CommandsSurface | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type ClipboardHostTarget = SuperDoc | ClipboardHost | null | undefined;

/* ------------------------------------------------------------------ */
/* הלוח הפנימי                                                        */
/* ------------------------------------------------------------------ */

let buffered: { payload: ClipboardPayload; mirrored: boolean } | null = null;

/**
 * רשת הביטחון כשלוח המערכת אינו זמין.
 *
 * לא מטמון ולא אופטימיזציה: ב-`file://` הרשאת הלוח נדחית, ובלי הלוח הזה
 * „גזור” היה מוחק טקסט שאי אפשר להחזיר בהדבקה. הוא חי כל עוד הדף חי, ואינו
 * מסונכרן עם לוח המערכת — בכוונה: לוח המערכת הוא המקור כשהוא זמין.
 *
 * ## `mirrored`: לא כל תוכן ששמור כאן הוא תוכן שמותר להדביק
 *
 * הלוח הפנימי הוא **תמונת ראי** של מה שנשלח ללוח המערכת, ולכן יש לו שני מצבים
 * שונים בתכלית:
 *
 * - `mirrored: false` — הכתיבה ללוח המערכת נחסמה, והעותק כאן הוא **העותק
 *   היחיד**. בהדבקה שאין בה הרשאת קריאה הוא הידע הטוב ביותר שיש, וזה מה שמציל
 *   „גזור” מלמחוק טקסט שאין לו החזרה.
 * - `mirrored: true` — אותו תוכן יושב **גם** בלוח המערכת, ומאותו רגע לוח
 *   המערכת הוא המקור. המשתמש עשוי להעתיק שם משהו חדש, או לבחור פריט אחר
 *   מהיסטוריית Win+V, ואנחנו לא נדע. הדבקת הראי במצב הזה מתעלמת ממה שהוא בחר
 *   ומחזירה תוכן ישן — וזו בדיוק התלונה שהאבחנה הזאת מסירה.
 *
 * מיוצא כאובייקט (ולא כמשתנה) כדי שבדיקה תוכל לנקות אותו בין מקרים בלי שם
 * שמכריז „for tests”.
 */
export const internalClipboard = {
  read(): ClipboardPayload | null {
    return buffered?.payload ?? null;
  },
  /**
   * התוכן, ורק כשהוא העותק היחיד. זה מה שההדבקה קוראת: ראי של לוח המערכת אינו
   * תשובה לשאלה „מה יש בלוח **עכשיו**”.
   */
  readUnmirrored(): ClipboardPayload | null {
    return buffered && !buffered.mirrored ? buffered.payload : null;
  },
  write(payload: ClipboardPayload, { mirrored = false }: { mirrored?: boolean } = {}): void {
    buffered = { payload, mirrored };
  },
  clear(): void {
    buffered = null;
  },
};

/* ------------------------------------------------------------------ */
/* עזרים                                                              */
/* ------------------------------------------------------------------ */

const COPY_FAILED = 'ההעתקה נכשלה';
const CUT_FAILED = 'הגזירה נכשלה';
const PASTE_FAILED = 'ההדבקה נכשלה';

/**
 * הנוסח כשהמנוע אינו חושף את הלוח. זהה למה שהיכולת מחזירה על
 * `NAMESPACE_UNAVAILABLE`, כדי שהמשתמש יראה את אותו הסבר בין אם הפקד מנוטרל
 * ובין אם הוא נלחץ לפני שהיכולות נקראו.
 */
const UNAVAILABLE = 'אינו זמין בגרסה זו';

/**
 * מה שנאמר למשתמש כשהתוכן נשמר בתוך התוסף אבל לא הגיע ללוח המערכת.
 *
 * מגיע אליו כ-`ok: false` ולא כהצלחה שקטה: הוא ביקש להעתיק, וההעתקה **לא**
 * הגיעה למקום שממנו מדביקים ל-Word. הקיצור שמוזכר כאן עובד — הוא מטופל בתוך
 * המסמך על ידי המנוע עצמו, לא על ידינו.
 */
const SYSTEM_BLOCKED_REASON = 'system-clipboard-blocked';

function docOf(host: ClipboardHostTarget): ClipboardDocumentApi | undefined {
  return (host as ClipboardHost | null | undefined)?.activeEditor?.doc ?? undefined;
}

/**
 * `superdoc.ui` הוא getter שיוצר את ה-controller בקריאה הראשונה. קריאתו
 * מתועדת כבטוחה גם לפני שהמסמך מוכן, אבל היא בכל זאת קריאה אל קוד זר —
 * וכשל שלה אינו סיבה להפיל לחיצה על כפתור.
 */
function selectionOf(host: ClipboardHostTarget): SelectionSurface | undefined {
  try {
    return (host as ClipboardHost | null | undefined)?.ui?.selection ?? undefined;
  } catch (error) {
    console.warn('[otzaria-word] קריאת משטח הבחירה מהמנוע נכשלה', error);
    return undefined;
  }
}

/** כשל עם הודעה בעברית מ-reason של ה-controller. */
function failed(failedAction: string, reason: string): CommandOutcome {
  return { ok: false, message: `${failedAction}: ${reasonText(reason)}`, reason };
}

/** `empty` = סמן בלי טווח, כפי שהתצלום מדווח. */
type SelectionRead = { ok: true; target: SelectionTarget; empty: boolean; ready: boolean } | { ok: false; reason: string };

/**
 * הטווח שהמשתמש סימן, כ-`SelectionTarget` — הצורה שפעולות הכתיבה של המנוע
 * מקבלות ישירות.
 *
 * `selectionTarget` ולא `target`: הראשון שומר את קצות הבחירה בדיוק, והשני הוא
 * מודל הכתובות של הערות ומיועד לצרכן אחר.
 */
function readSelection(
  surface: SelectionSurface | undefined,
  { allowCollapsed = false }: { allowCollapsed?: boolean } = {},
): SelectionRead {
  const getSnapshot = surface?.getSnapshot;
  if (typeof getSnapshot !== 'function') return { ok: false, reason: 'host-capability-unavailable' };

  let slice: SelectionSnapshot | undefined;
  try {
    slice = getSnapshot();
  } catch {
    return { ok: false, reason: 'not-ready' };
  }

  // סמן בלי טווח הוא בחירה **חוקית** להדבקה ופסולה להעתקה, ולכן גם נוסח הכשל
  // שונה: „יש למקם את הסמן במסמך” מול „יש לסמן טקסט תחילה”.
  const missing = allowCollapsed ? 'selection-required' : 'range-selection-required';

  const target = slice?.selectionTarget ?? null;
  if (!target || typeof target !== 'object') {
    // `status` שאינו `ready` אומר שקריאת הבחירה עוד לא הסתיימה — וזה מצב אחר
    // מ„המשתמש לא סימן כלום”, גם אם שניהם מגיעים בתור `selectionTarget: null`.
    return { ok: false, reason: slice?.status === 'ready' ? missing : 'not-ready' };
  }
  if (!allowCollapsed && slice?.empty === true) {
    // אותה הבחנה כמו בענף שלמעלה, ולא רק עליו: `empty` יכול לדווח `true`
    // מפני שהתצלום פיגר אחרי בחירה חיה שכבר קיימת במסמך, ולא רק מפני
    // שה-`selectionTarget` עצמו עוד לא הגיע. נמדד (ראו הערת הראש של הקובץ):
    // Shift+חץ מציג `empty:true` יחד עם `status` שאינו `ready` במשך עד
    // כ-820ms אחרי שיש כבר טווח אמיתי במסמך. בלי הבדיקה הזאת „העתק”/„גזור”
    // שנלחצים בחלון הזה מקבלים „יש לסמן טקסט תחילה” על בחירה שקיימת.
    return { ok: false, reason: slice?.status === 'ready' ? 'range-selection-required' : 'not-ready' };
  }

  return { ok: true, target, empty: slice?.empty === true, ready: slice?.status === 'ready' };
}

/** כמה פעמים, ובאיזה מרווח, מותר להמתין שהבחירה תתייצב. ר' `readSelectionSettled`. */
const SELECTION_SETTLE_RETRIES = 10;
const SELECTION_SETTLE_INTERVAL_MS = 100;

/**
 * כמו `readSelection`, אבל ל„העתק” ול„גזור”: `status` שאינו `ready` הוא „עוד
 * לא ידוע”, לא „אין בחירה”, וזה עניין של זמן בלבד — לכן ממתינים לו במקום
 * להיכשל עליו.
 *
 * ## למה לא המסלול בלי `target`
 *
 * `serializeSelection` מקבל `target` אופציונלי, וההנחה המקורית כאן הייתה
 * שבלעדיו המנוע מסדר את הבחירה החיה בעצמו — ולכן „העתק” ויתר על ההמתנה
 * וקרא בלי target כשלא היה טווח נקרא. נמדד מול המנוע האמיתי (ה-dist הארוז)
 * שזה **אינו** המצב: הקריאה נכשלת עם `"…requires an explicit target until
 * current-selection serialization is wired"`. כלומר גם „העתק” זקוק לטווח
 * שנפתר, בדיוק כמו „גזור” שצריך אותו למחיקה (`doc.delete` מקבל `target` או
 * `ref` בלבד) — ולכן שניהם ממתינים כאן באותה לולאה.
 *
 * זו לולאת בדיקה חוזרת ולא המתנה קבועה: היא בודקת מיד, ומחכה רק כשהתשובה
 * „עוד לא ידוע” — מכונה מהירה יוצאת מהלולאה בסיבוב הראשון, איטית ממשיכה, ואף
 * אחת לא משלמת מחיר קבוע שאינו שלה. הגבול העליון (כ-1000ms) הוא רשת ביטחון
 * ולא המנגנון עצמו: נמדד ש-820ms הוא הזמן הארוך ביותר שבו הבחירה עדיין לא
 * התיישבה.
 */
async function readSelectionSettled(
  surface: SelectionSurface | undefined,
  { allowCollapsed = false, requireReady = false }: { allowCollapsed?: boolean; requireReady?: boolean } = {},
): Promise<SelectionRead> {
  for (let attempt = 0; attempt <= SELECTION_SETTLE_RETRIES; attempt += 1) {
    if (attempt > 0) await wait(SELECTION_SETTLE_INTERVAL_MS);
    const result = readSelection(surface, { allowCollapsed });
    if (result.ok) {
      if (!requireReady || result.ready) return result;
    } else if (result.reason !== 'not-ready') return result;
    if (attempt === SELECTION_SETTLE_RETRIES) return { ok: false, reason: 'not-ready' };
  }
  return { ok: false, reason: 'not-ready' };
}

/**
 * אותו טווח, מכווץ לאחד מקצותיו — כלומר סמן בלי בחירה.
 *
 * `story` ו-`coordinateSpace` נשמרים: פעולה בכותרת עליונה או בהערת שוליים
 * אינה בגוף המסמך, ובחירה בלי הזהות הזאת הייתה מצביעה על בלוק אחר לגמרי.
 */
function collapsedTo(target: SelectionTarget, edge: 'start' | 'end'): SelectionTarget {
  const point = target[edge];

  return {
    kind: 'selection',
    start: point,
    end: point,
    ...(target.story ? { story: target.story } : {}),
    ...(target.coordinateSpace ? { coordinateSpace: target.coordinateSpace } : {}),
  };
}

/** הפריט מה-payload לפי MIME, כשהוא מחרוזת. `null` = אין כזה. */
function textItem(payload: ClipboardPayload | undefined, type: string): string | null {
  const items = payload?.items;
  if (!Array.isArray(items)) return null;
  for (const item of items) {
    if (item?.type === type && item.kind === 'string' && typeof item.data === 'string') return item.data;
  }
  return null;
}

/**
 * ה-API של הדפדפן, או `undefined`.
 *
 * הצהרת ה-DOM ב-TypeScript מבטיחה ש-`navigator.clipboard` קיים; ב-jsdom
 * ובהקשרים לא-מאובטחים הוא פשוט אינו שם. ה-cast הוא ההודאה בפער הזה.
 */
function systemClipboard(): Clipboard | undefined {
  if (typeof navigator === 'undefined') return undefined;
  return (navigator as Navigator & { clipboard?: Clipboard }).clipboard;
}

/**
 * כותבת ללוח המערכת. מחזירה האם זה הצליח — ולא זורקת.
 *
 * `write()` עם `text/html` **ו**-`text/plain` הוא מה שמאפשר להדביק ל-Word
 * אמיתי ולשמור עיצוב; `writeText` לבדו היה משטח את הכול לטקסט. `write`
 * ו-`ClipboardItem` אינם קיימים בכל סביבה, ולכן הם נבדקים ולא מונחים, ויש
 * נפילה מדורגת אל `writeText`.
 *
 * ה-MIME הפרטי של המנוע (`application/x-superdoc-v2-fragment`) אינו נשלח:
 * לוח המערכת מסנן טיפוסים שאינם ברשימה המותרת, וניסיון לכתוב אותו מפיל את
 * כל הקריאה. העתקה בתוך המסמך שומרת נאמנות דרך הלוח הפנימי.
 */
async function writeSystemClipboard(payload: ClipboardPayload): Promise<boolean> {
  const plain = textItem(payload, 'text/plain');
  const html = textItem(payload, 'text/html');
  if (plain === null && html === null) return false;

  const api = systemClipboard();
  if (!api) return false;

  if (html !== null && typeof api.write === 'function' && typeof ClipboardItem === 'function') {
    const parts: Record<string, Blob> = { 'text/html': new Blob([html], { type: 'text/html' }) };
    if (plain !== null) parts['text/plain'] = new Blob([plain], { type: 'text/plain' });
    try {
      await api.write([new ClipboardItem(parts)]);
      return true;
    } catch {
      // הרשאה שנדחתה, או טיפוס שהדפדפן סינן. `writeText` הוא המסלול המצומצם
      // שלעתים כן עובר, ולכן ממשיכים אליו במקום לדווח כשל מיד.
    }
  }

  if (plain === null || typeof api.writeText !== 'function') return false;
  try {
    await api.writeText(plain);
    return true;
  } catch {
    return false;
  }
}

/* ------------------------------------------------------------------ */
/* העתק וגזור                                                         */
/* ------------------------------------------------------------------ */

type Serialized = { ok: true; payload: ClipboardPayload } | { ok: false; outcome: CommandOutcome };

/**
 * מסדרת את הבחירה דרך המנוע ומחזירה את ה-payload.
 *
 * השמירה בלוח הפנימי אינה כאן אלא באתרי הקריאה, מפני שהיא צריכה לדעת דבר אחד
 * שטרם ידוע בשלב הזה: האם אותו תוכן הגיע גם ללוח המערכת. ההסבר על ההבדל —
 * `internalClipboard`.
 *
 * `target` אופציונלי: החוזה קובע ש-`serializeSelection` מסדר „the current or
 * supplied model selection”, אבל נמדד מול המנוע האמיתי (ה-dist הארוז) שבלי
 * target הקריאה **נכשלת** — „…requires an explicit target until
 * current-selection serialization is wired”. כלומר זה אינו מסלול חלופי תקין
 * להעתקה רגילה (קוראי `copySelection`/`cutSelection` ממתינים לטווח שנפתר,
 * ר' `readSelectionSettled`), אלא מה שנשאר לקרוא בו כשאין בכלל דרך לקרוא את
 * הבחירה (`ui.selection` עצמו חסר) — שם אין ברירה חוץ מלנסות בלי target ולתת
 * למנוע להחליט.
 */
async function serializeSelection(
  host: ClipboardHostTarget,
  failedAction: string,
  target: SelectionTarget | undefined,
): Promise<Serialized> {
  const serialize = docOf(host)?.clipboard?.serializeSelection;
  if (typeof serialize !== 'function') {
    return {
      ok: false,
      outcome: {
        ok: false,
        message: `${failedAction}: ${UNAVAILABLE}`,
        reason: 'command-unsupported',
      },
    };
  }

  let result: ClipboardSerializeResult | undefined;
  try {
    result = await serialize(target ? { target, includeHtml: true } : { includeHtml: true });
  } catch (error) {
    return { ok: false, outcome: { ok: false, message: thrownText(failedAction, error), reason: 'threw' } };
  }

  const payload = result?.payload;
  // אין items = אין מה להעתיק. זה המצב של סמן בלי טווח, והוא בקשה שאין לה
  // תוצאה — לא תקלה במנוע.
  if (!payload || !Array.isArray(payload.items) || payload.items.length === 0) {
    return { ok: false, outcome: failed(failedAction, 'range-selection-required') };
  }

  return { ok: true, payload };
}

/** ההודעה כשהתוכן בלוח הפנימי בלבד. `gesture` הוא הקיצור שכן מגיע לתוכנה אחרת. */
function systemBlocked(what: string, gesture: string): CommandOutcome {
  return {
    ok: false,
    message: `${what} בתוך התוסף בלבד — לוח המערכת חסם את הפעולה. הדבקה כאן תעבוד; להעברה ל-Word יש להשתמש ב-${gesture}.`,
    reason: SYSTEM_BLOCKED_REASON,
  };
}

/**
 * „העתק”. מסדרת את הבחירה, שומרת בלוח הפנימי, וכותבת ללוח המערכת.
 *
 * לעולם אינה זורקת: כל הקריאות למנוע ולדפדפן עטופות, וחריגה מפקד ב-Ribbon
 * מפילה את רינדור הרצועה כולה.
 */
export async function copySelection(host: ClipboardHostTarget): Promise<CommandOutcome> {
  // `readSelectionSettled` ולא `readSelection`: נמדד מול המנוע האמיתי (ה-dist
  // הארוז) ש-`serializeSelection` בלי `target` **נכשל** — "requires an
  // explicit target until current-selection serialization is wired" — ולא
  // מסדר את הבחירה החיה כפי שההערה למטה הניחה. לכן `status` שאינו `ready`
  // ממתין להתייצבות כאן, בדיוק כמו ב„גזור"; המסלול בלי target נשאר למטה רק
  // למקרה שאין בכלל משטח לקרוא ממנו (`host-capability-unavailable`).
  const selection = await readSelectionSettled(selectionOf(host));
  // בחירה ריקה נעצרת כאן ולא במנוע, כדי שההודעה תהיה „יש לסמן טקסט תחילה”
  // ולא כשל סדרוּר. חוסר **מסלול** לקרוא את הבחירה (למשל אין `ui.selection`
  // בכלל) אינו עוצר: הקריאה למטה עדיין יוצאת בלי target, וזה המסלול שנשאר
  // תלוי במנוע — לא הבחירה הרגילה.
  if (!selection.ok && selection.reason === 'range-selection-required') {
    return failed(COPY_FAILED, selection.reason);
  }

  const serialized = await serializeSelection(
    host,
    COPY_FAILED,
    selection.ok ? selection.target : undefined,
  );
  if (!serialized.ok) return serialized.outcome;

  const onSystemClipboard = await writeSystemClipboard(serialized.payload);
  internalClipboard.write(serialized.payload, { mirrored: onSystemClipboard });

  if (!onSystemClipboard) return systemBlocked('הטקסט הועתק', 'Ctrl+C');
  return { ok: true };
}

/**
 * „גזור” = העתק ואחר כך מחק.
 *
 * הסדר אינו שרירותי: המחיקה רצה **רק** אחרי שהתוכן נשמר — בלוח המערכת, ואם
 * הוא נחסם אז בלוח הפנימי — כדי שכשל בסדרוּר לא ימחוק טקסט שאין לו עותק. לוח
 * מערכת שנחסם כן מאפשר להמשיך למחיקה: התוכן קיים בתוך התוסף, מסומן שם כעותק
 * היחיד, וההודעה אומרת את זה במפורש.
 */
export async function cutSelection(host: ClipboardHostTarget): Promise<CommandOutcome> {
  const doc = docOf(host);
  const remove = doc?.delete;
  if (typeof remove !== 'function') {
    return { ok: false, message: `${CUT_FAILED}: ${UNAVAILABLE}`, reason: 'command-unsupported' };
  }

  // ל„גזור” הבחירה היא חובה, ולא נוחות: `doc.delete` מקבל `target` או `ref`,
  // ובלי אחד מהם אין מה למחוק. `readSelectionSettled` ולא `readSelection`:
  // בניגוד ל„העתק”, אין כאן מסלול חלופי בלי target, ולכן `status` שעוד לא
  // הסתיים ממתין להתייצבות במקום להיכשל על „not-ready”.
  const selection = await readSelectionSettled(selectionOf(host));
  if (!selection.ok) return failed(CUT_FAILED, selection.reason);

  const serialized = await serializeSelection(host, CUT_FAILED, selection.target);
  if (!serialized.ok) return serialized.outcome;

  const onSystemClipboard = await writeSystemClipboard(serialized.payload);
  internalClipboard.write(serialized.payload, { mirrored: onSystemClipboard });

  let receipt: DocReceipt | undefined;
  try {
    receipt = await remove({ target: selection.target });
  } catch (error) {
    return { ok: false, message: thrownText(CUT_FAILED, error), reason: 'threw' };
  }

  if (receipt?.success === false) {
    return {
      ok: false,
      message: receiptFailureText(CUT_FAILED, receipt),
      reason: receipt.failure?.code,
    };
  }

  // הסמן חוזר למקום שממנו נגזר. ההנמקה המלאה ב-`caretAfterCut`.
  await moveCaret(selectionOf(host), caretAfterCut(selection.target), {
    what: 'למקום הגזירה',
    retries: 0,
  });

  if (!onSystemClipboard) return systemBlocked('הטקסט נגזר', 'Ctrl+X');
  return { ok: true };
}

/**
 * הסמן שאחרי „גזור”: תחילת הטווח שנמחק.
 *
 * ## התלונה שזה נולד ממנה
 *
 * „גוזר, וזה גוזר — אבל הסמן יוצא מהטקסט.” אחרי `doc.delete` המנוע נשאר
 * **בלי בחירה בכלל**. נמדד מול המנוע האמיתי (Chrome על ה-dist הארוז): גזירת
 * מילה מסומנת מחזירה `getSnapshot()` = `{status:'ready', empty:true,
 * selectionTarget:null}`, אין סמן מצויר במסמך, וההקלדה הבאה **אינה נכנסת לשום
 * מקום** — הטקסט נשאר כפי שהיה אחרי המחיקה.
 *
 * המיקוד עצמו אינו הבעיה: `document.activeElement` נשאר ה-textarea המוסתר של
 * המנוע לאורך כל המסלול, מפני שהפקד מונע את גזילת המיקוד (`@pointerdown.prevent`
 * ב-RibbonButton, ו-tests/unit/clipboard: „הפקד עצמו מונע מהלחיצה לגזול את
 * המיקוד”). מה שחסר הוא בחירה להקליד לתוכה.
 *
 * ## למה **תחילת** הטווח, ולמה היא שורדת את המחיקה
 *
 * זה גם מה ש-Word עושה: הסמן נשאר במקום שממנו הטקסט נעלם. נקודת ההתחלה נשארת
 * ניתנת לאיתור גם כשהמחיקה חוצה פסקאות — המנוע מותיר את בלוק ההתחלה ומכריז על
 * בלוק **הסיום** כ-`invalidatedRefs`. שלוש הצורות נמדדו, וכולן החזירו
 * `apply → {ok:true}` עם סמן מצויר והקלדה שנכנסת במקום הנכון:
 *
 * - טווח בתוך פסקה אחת;
 * - טווח חוצה פסקאות בכיסוי חלקי (הפסקאות מתמזגות לבלוק ההתחלה);
 * - טווח חוצה פסקאות שמכסה את שתיהן במלואן (בלוק ההתחלה שורד, ריק).
 *
 * ## למה בלי ניסיון חוזר
 *
 * להבדיל מההדבקה (`moveCaretAfterPaste`), היעד כאן אינו בלוק שזה עתה נוצר אלא
 * בלוק שכבר היה במסמך, ולכן אין ממה להמתין: בשלוש הצורות שנמדדו ההחלה
 * **מיד** אחרי שהמחיקה חזרה — בלי השהיה — עברה בניסיון הראשון.
 */
function caretAfterCut(target: SelectionTarget): SelectionTarget {
  return collapsedTo(target, 'start');
}

/* ------------------------------------------------------------------ */
/* הדבק                                                               */
/* ------------------------------------------------------------------ */

/**
 * ה-MIME-ים שמסלול הכפתור מעביר למנוע.
 *
 * `text/html` לפני `text/plain` — הראשון הוא זה שנושא עיצוב, וה-`parse` של
 * המנוע בוחר את הייצוג הנאמן ביותר מבין הפריטים. התמונות נקראות כ-`bytes`,
 * כי זה מה שהוולידטור דורש (`data instanceof Uint8Array`), והמנוע יודע להפוך
 * אותן ל-assets בתוך ה-fragment.
 *
 * מה שמחוץ לרשימה — RTF, קבצים, `image/gif` — אינו מוברח בכל זאת: `parse`
 * היה דוחה אותו, וההדבקה המקורית של המנוע (Ctrl+V) כן מטפלת בו.
 */
const READ_AS_TEXT = ['text/html', 'text/plain'] as const;
const READ_AS_BYTES = ['image/png', 'image/jpeg'] as const;

/**
 * קודי הכשל של משפחת `clipboard.*` שאינם בטבלה המשותפת שב-document-api.ts.
 *
 * שכבה דקה ולא טבלה שנייה: כל השאר נופל ל-`receiptFailureText`, כדי שלא
 * ייווצרו שני נוסחים עבריים לאותו קוד. (הטבלה שם פרטית ואינה מיוצאת.)
 */
const CLIPBOARD_FAILURE_TEXT: Record<string, string> = {
  INVALID_PAYLOAD: 'תוכן הלוח אינו במבנה שהמנוע מקבל',
  EMPTY_FRAGMENT: 'אין תוכן להדבקה',
  INVALID_FRAGMENT: 'תוכן הלוח פגום ואינו ניתן להדבקה',
};

/**
 * למה המנוע לא ידע להדביק. אלה קודי `SDPasteUnsupportedReason`, והם ההסבר
 * **המדויק** — `failure.code` במקרים האלה הוא רק `CAPABILITY_UNSUPPORTED`.
 */
const PASTE_UNSUPPORTED_TEXT: Record<string, string> = {
  'paste-empty': 'אין תוכן להדבקה',
  'paste-payload-too-large': 'התוכן שבלוח גדול מדי להדבקה',
  'paste-source-unsupported': 'המקור שהתוכן הועתק ממנו אינו נתמך',
  'paste-no-faithful-representation': 'אין דרך להדביק את התוכן הזה בלי לשנות אותו',
  'paste-structure-unsupported': 'מבנה התוכן שבלוח אינו נתמך בהדבקה',
  'paste-media-unsupported-type': 'סוג התמונה שבלוח אינו נתמך',
  'paste-media-too-large': 'התמונה שבלוח גדולה מדי',
  'paste-media-bad-magic': 'התמונה שבלוח פגומה',
  'paste-security-rejected': 'התוכן שבלוח נדחה מטעמי אבטחה',
  'paste-target-unsupported': 'לא ניתן להדביק במקום הזה במסמך',
  'paste-tracked-structural-unsupported': 'הדבקה כזאת אינה נתמכת במצב מעקב אחר שינויים',
  'paste-cross-story-unsupported': 'לא ניתן להדביק בין חלקים שונים של המסמך',
  'paste-fragment-version-unsupported': 'התוכן הועתק מגרסה אחרת של המנוע',
  'paste-legacy-slice-unsupported': 'התוכן הועתק מגרסה ישנה של המנוע',
  'paste-depth-exceeded': 'התוכן שבלוח מקונן עמוק מדי',
};

type SystemRead =
  | { kind: 'payload'; payload: ClipboardPayload }
  /** הלוח נקרא, ואין בו דבר שאפשר להעביר למנוע. */
  | { kind: 'empty' }
  /** אין API, או שההרשאה נדחתה. זה המצב הנפוץ ב-`file://`. */
  | { kind: 'blocked' };

/**
 * קוראת את לוח המערכת ובונה ממנו `ClipboardPayload`.
 *
 * ההבחנה בין `empty` ל-`blocked` היא מה שמכריע אם ליפול ללוח הפנימי: לוח
 * מערכת שנקרא **והיה ריק** הוא תשובה, והדבקה של משהו שהועתק לפני חצי שעה
 * בתוך התוסף הייתה שגויה. הרשאה שנדחתה היא היעדר מידע, ושם הלוח הפנימי הוא
 * הידע הטוב ביותר שיש.
 */
async function readSystemClipboard(): Promise<SystemRead> {
  const api = systemClipboard();
  if (!api) return { kind: 'blocked' };

  let readable = false;

  if (typeof api.read === 'function') {
    try {
      const items = (await api.read()) ?? [];
      readable = true;
      const collected: ClipboardPayloadItem[] = [];
      for (const item of items) {
        const types: readonly string[] = item?.types ?? [];
        for (const type of READ_AS_TEXT) {
          if (!types.includes(type)) continue;
          const data = await (await item.getType(type)).text();
          if (data) collected.push({ type, kind: 'string', data });
        }
        for (const type of READ_AS_BYTES) {
          if (!types.includes(type)) continue;
          const data = new Uint8Array(await (await item.getType(type)).arrayBuffer());
          if (data.length > 0) collected.push({ type, kind: 'bytes', data });
        }
      }
      if (collected.length > 0) {
        return { kind: 'payload', payload: { source: 'browser', items: collected } };
      }
    } catch {
      // `NotAllowedError` הוא המסלול הנפוץ ב-`file://`. `readText` מבקש הרשאה
      // מצומצמת יותר ולעתים כן עובר, ולכן ממשיכים אליו.
    }
  }

  if (typeof api.readText === 'function') {
    try {
      const text = await api.readText();
      readable = true;
      if (text) {
        return {
          kind: 'payload',
          payload: { source: 'browser', items: [{ type: 'text/plain', kind: 'string', data: text }] },
        };
      }
    } catch {
      // אותו דבר; כאן נגמרו המסלולים.
    }
  }

  return readable ? { kind: 'empty' } : { kind: 'blocked' };
}

/** ההודעה כשאין הרשאה לקרוא את הלוח. Ctrl+V כן עובד — המנוע מטפל בו במסמך. */
function readBlocked(): CommandOutcome {
  return {
    ok: false,
    message: 'ההדבקה נכשלה: אין הרשאה לקרוא את לוח המערכת. יש להדביק עם Ctrl+V — המנוע מטפל בו בתוך המסמך.',
    reason: SYSTEM_BLOCKED_REASON,
  };
}

/**
 * הסמן שאחרי ההדבקה: קצה הטקסט שנוצר, מכווץ לנקודה.
 *
 * ## התלונה שזה נולד ממנה
 *
 * „מדביק משה, ואז כהן — ויוצא כהןמשה.” המנוע **אינו** מזיז את הבחירה החיה
 * אחרי `clipboard.insert`. נמדד מול המנוע האמיתי (Chrome על ה-dist הארוז):
 * לפני ההדבקה הסמן ב-`41964671:0`, אחריה — `41964671:0` שוב, בזמן שהקבלה כבר
 * מדווחת על טווח שנוצר בין 0 ל-3. ההדבקה הבאה קוראת את אותו סמן, ולכן נכנסת
 * **לפני** קודמתה.
 *
 * Ctrl+V אינו סובל מזה: שם המנוע מטפל באירוע `paste` בעצמו ומזיז את הסמן
 * כחלק מאותה עסקה. מסלול הכפתור עובר ב-Document API, ושם הזזת הבחירה היא
 * באחריות מי שקרא — בדיוק כמו שליחת ה-`target` להדבקה.
 *
 * ## למה מהקבלה, ולא בחישוב
 *
 * אין כאן ספירת תווים משלנו: הדבקה של טקסט רב-שורות או של HTML נפרסת לכמה
 * פסקאות, ואורך מה שהודבק אינו האורך שנכנס לפסקה שבה הסמן היה. הקבלה מדווחת
 * על מה שנוצר בפועל, ו-`selectionTarget` שבתוכה הוא כבר הצורה שפעולות הבחירה
 * מקבלות.
 *
 * נמדד: גם ב-„אחת\nשתיים\nשלוש” וגם ב-`<p>אלף</p><p>בית</p>` חוזר **פריט
 * אחד** ב-`insertedText`, והוא הטווח האחרון שנוצר („שלוש”, „בית”) — כלומר
 * בדיוק המקום שבו הסמן צריך לשבת. `insertedBlocks` היה ריק בכל המקרים, ולכן
 * הוא אינו נקרא כאן. הקצה נלקח מהאיבר האחרון ולא מהראשון: החוזה מבטיח סדר
 * מסמך, וגרסה שתדווח על כמה טווחים תיקרא נכון.
 *
 * `null` = הקבלה אינה מדווחת על טווח שנוצר. אז אין מה להזיז, וההדבקה עצמה
 * עדיין הצליחה.
 */
function caretAfterPaste(receipt: ClipboardInsertReceipt | undefined): SelectionTarget | null {
  const spans = receipt?.effects?.insertedText;
  if (!Array.isArray(spans) || spans.length === 0) return null;

  const span = spans[spans.length - 1]?.selectionTarget;
  if (!span?.end) return null;

  return collapsedTo(span, 'end');
}

/** כמה ניסיונות חוזרים, ובאיזה מרווח, כשהבלוק החדש טרם אותר. ההנמקה למטה. */
const CARET_RETRIES = 3;
const CARET_RETRY_MS = 80;

const wait = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * מזיזה את הסמן לסוף מה שהודבק.
 *
 * best-effort בכוונה: התוכן כבר במסמך, וכשל בהזזת הסמן אינו הופך הדבקה
 * שהצליחה לכישלון — הוא נרשם ללוג, כמו הדבקה שנפלה לטקסט בלבד. `apply` הוא
 * adapter אופציונלי (ההסבר ב-`selectWholeDocument`), ולכן היעדרו הוא מסלול
 * חוקי ולא שגיאה.
 *
 * ## למה ניסיון חוזר
 *
 * הדבקה שיוצרת פסקה חדשה מחזירה בקבלה בלוק שברגע החזרה של `insert` עדיין
 * אינו ניתן לאיתור: `apply` נכשל סגור עם `target-unresolved`. נמדד מול המנוע
 * האמיתי, אותה הדבקה רב-שורתית בארבע השהיות — 0ms נדחה (וניסיון חוזר
 * 250ms אחריו עבר), ו-50ms / 300ms / 1000ms עברו בראשון.
 *
 * לכן הניסיון הראשון מיידי: הדבקה שנשארת בתוך הפסקה — הרוב המוחלט — מסתדרת
 * בו, והשהיה קבועה לפני `apply` הייתה מחייבת **כל** הדבקה לשלם את מחיר המקרה
 * הנדיר. אירוע „הבלוק נקלט” אינו בחוזה, ולכן ההמתנה היא ניסיונות קצרים.
 *
 * `apply` שזורק אינו חוזר על עצמו: חריגה אינה „עדיין לא מוכן” אלא מסלול שבור.
 */
async function moveCaretAfterPaste(
  surface: SelectionSurface | undefined,
  receipt: ClipboardInsertReceipt | undefined,
): Promise<void> {
  const caret = caretAfterPaste(receipt);
  if (!caret) return;

  await moveCaret(surface, caret, { what: 'לסוף ההדבקה', retries: CARET_RETRIES });
}

/**
 * מחילה סמן על העורך החי, ומדווחת ללוג בלבד.
 *
 * `retries` הוא ההבדל בין שני הקוראים, ולא גחמה: אחרי הדבקה היעד עשוי להיות
 * בלוק שזה עתה נוצר וטרם ניתן לאיתור (ההנמקה למעלה), ואחרי גזירה היעד הוא
 * בלוק **שכבר היה שם** — ולכן ניסיון אחד (ההנמקה ב-`caretAfterCut`).
 */
async function moveCaret(
  surface: SelectionSurface | undefined,
  caret: SelectionTarget,
  { what, retries }: { what: string; retries: number },
): Promise<void> {
  const apply = surface?.apply;
  if (typeof apply !== 'function') return;

  for (let attempt = 0; attempt <= retries; attempt += 1) {
    if (attempt > 0) await wait(CARET_RETRY_MS);

    let result: { ok?: boolean; reason?: string } | undefined;
    try {
      result = await apply(caret);
    } catch (error) {
      console.warn(`[otzaria-word] הזזת הסמן ${what} נכשלה`, error);
      return;
    }

    if (result?.ok !== false) return;
    if (attempt === retries) {
      console.warn(`[otzaria-word] הזזת הסמן ${what} נדחתה`, result.reason);
    }
  }
}

/** ההסבר בעברית לכשל של `parse` או `insert`. */
function pasteFailureText(failure: ClipboardFailure | undefined): string {
  const unsupported = failure?.details?.unsupportedReason;
  const explained = unsupported ? PASTE_UNSUPPORTED_TEXT[unsupported] : undefined;
  if (explained) return `${PASTE_FAILED}: ${explained}`;

  const known = failure?.code ? CLIPBOARD_FAILURE_TEXT[failure.code] : undefined;
  if (known) return `${PASTE_FAILED}: ${known}`;

  return receiptFailureText(PASTE_FAILED, { success: false, failure });
}

/**
 * „הדבק”.
 *
 * המסלול הוא זה שהחוזה מתאר: `ClipboardPayload` → `parse` → `insert`. ה-`parse`
 * נפרד מה-`insert` בכוונה — הוא **אינו** משנה את המסמך, ולכן תוכן שאי אפשר
 * להדביק נעצר לפני שנגענו במסמך, והמשתמש מקבל את הסיבה המדויקת
 * (`unsupportedReason`) ולא רק „הפעולה נכשלה”.
 *
 * ## למה `target` **כן** נשלח
 *
 * כאן הייתה ההנחה שהפילה את הפעולה: „המנוע פותר את מקום ההדבקה מהבחירה החיה
 * בעצמו — בדיוק כמו ב-Ctrl+V”. הוא אינו. נמדד מול המנוע האמיתי, עם בחירה חיה
 * בתחילת המסמך:
 *
 * - `insert({ plan })` — הקבלה חזרה `success: true` עם `atChar` בסוף המסמך.
 *   התוכן נכנס אחרי הפסקה האחרונה, והבחירה לא נלקחה בחשבון בכלל.
 * - `insert({ plan, target })` — `atChar: 0`, כלומר בדיוק בסמן.
 *
 * זה מתיישב עם מה שהמנוע עושה בעצמו: `pasteClipboardPayload` שבפאסדה שלו קורא
 * את הבחירה ו**נכשל** כשאין ממה לגזור target, במקום לתת ל-`insert` להחליט.
 * `ClipboardTarget` בחוזה כולל `{ kind: 'documentEnd' }`, וזאת התנהגות ברירת
 * המחדל — לא נפילה.
 *
 * לכן מקום ההדבקה נקרא **ראשון**, לפני הלוח: הוא המצב ברגע הלחיצה, בעוד קריאת
 * הלוח היא `await` שבמהלכו הבחירה יכולה להשתנות. חסר סמן — הפעולה נכשלת עם
 * „יש למקם את הסמן במסמך”, ולא מוסיפה בשקט בסוף המסמך.
 *
 * ההדבקה המקורית של המנוע אינה נוגעת בזה: המודול הזה אינו רושם מאזינים
 * ואינו חוטף אירועי `paste`. מי שכן — כשברירת המחדל של Ctrl+V אינה „שמירה
 * על עיצוב המקור” — הוא engine/paste-intercept.ts, והוא מעביר לכאן את התוכן
 * של האירוע ב-`payload`, במקום קריאה מלוח המערכת.
 *
 * ## `mode`
 *
 * „שמירה על עיצוב המקור” הוא המסלול שתואר למעלה, כמות שהוא. „מזג עיצוב”
 * ו„שמירת טקסט בלבד” משכתבים את ה-fragment שה-`parse` בנה לפני ה-`insert` —
 * למה בעצמנו ולא ב-`formattingMode` של המנוע, ב-engine/paste-formatting.ts.
 */
export async function pasteFromClipboard(
  host: ClipboardHostTarget,
  options: { mode?: PasteMode; payload?: ClipboardPayload } = {},
): Promise<CommandOutcome> {
  const editor = (host as ClipboardHost | null | undefined)?.activeEditor;
  const operation = {};
  const container = (host as ClipboardHost | null | undefined)?.element ?? editor?.container;
  container?.dispatchEvent(new CustomEvent(PASTE_STARTED, { bubbles: true, detail: { operation } }));
  try {
    return await performPaste(host, options);
  } catch (error) {
    return { ok: false, message: thrownText(PASTE_FAILED, error), reason: 'threw' };
  } finally {
    container?.dispatchEvent(new CustomEvent(PASTE_FINISHED, { bubbles: true, detail: { operation, evaluate: (host as ClipboardHost | null | undefined)?.activeEditor === editor } }));
  }
}

async function performPaste(
  host: ClipboardHostTarget,
  { mode = 'keep-source', payload: given }: { mode?: PasteMode; payload?: ClipboardPayload } = {},
): Promise<CommandOutcome> {
  const destinationDoc = docOf(host);
  const clipboard = destinationDoc?.clipboard;
  const insert = clipboard?.insert;
  if (typeof insert !== 'function') {
    return { ok: false, message: `${PASTE_FAILED}: ${UNAVAILABLE}`, reason: 'command-unsupported' };
  }

  // `allowCollapsed`: סמן בלי טווח הוא המצב הרגיל של הדבקה. טווח מסומן גם הוא
  // חוקי — המנוע מחליף אותו בתוכן שהודבק, כמו ב-Word.
  let at = readSelection(selectionOf(host), { allowCollapsed: true });
  if (!at.ok) return failed(PASTE_FAILED, at.reason);
  if (mode !== 'keep-source' && !at.ready) {
    const originalTarget = at.target;
    const before = await destinationDoc?.info?.();
    const settled = await readSelectionSettled(selectionOf(host), { allowCollapsed: true, requireReady: true });
    const after = await destinationDoc?.info?.();
    const latest = readSelection(selectionOf(host), { allowCollapsed: true });
    if (!settled.ok || !latest.ok || !latest.ready || docOf(host) !== destinationDoc || before?.revision === undefined || before.revision !== after?.revision ||
        JSON.stringify(settled.target) !== JSON.stringify(originalTarget) ||
        JSON.stringify(latest.target) !== JSON.stringify(originalTarget)) return failed(PASTE_FAILED, 'not-ready');
    at = latest;
  }
  if (mode !== 'keep-source' && (!at.empty || at.target.start.kind !== 'text' || at.target.end.kind !== 'text' || at.target.start.blockId !== at.target.end.blockId || at.target.start.offset !== at.target.end.offset)) {
    return { ok: false, reason: 'paste-destination-unsupported', message: 'ההדבקה לא בוצעה: יש לבטל את הבחירה ולמקם סמן במסמך, או להשתמש בשמירה על עיצוב המקור.' };
  }

  // העיצוב שבסמן נקרא **עכשיו**, באותו רגע כמו המקום: שניהם המצב ברגע
  // הלחיצה, וקריאת הלוח שמתחת היא `await` שבמהלכו הסמן יכול לזוז.
  const caret = mode === 'keep-source' ? null : readCaretFormat(host, at.empty);
  // Invoke the capture now, before awaiting clipboard IO or parsing.
  const destinationRead = caret ? readDestination(host, at.target, caret) : null;

  const read: SystemRead = given ? { kind: 'payload', payload: given } : await readSystemClipboard();
  // `readUnmirrored` ולא `read`: כשלוח המערכת קריא-אך-חסום, תוכן שכבר נשלח אליו
  // אינו „מה שבלוח עכשיו”. ההסבר — `internalClipboard`.
  const raw =
    read.kind === 'payload' ? read.payload : read.kind === 'blocked' ? internalClipboard.readUnmirrored() : null;
  const payload = raw && mode === 'text-only' ? textOnlyPayload(raw) : raw;

  if (!payload) {
    if (read.kind === 'blocked') return readBlocked();
    if (raw) {
      return { ok: false, message: `${PASTE_FAILED}: אין בלוח טקסט להדבקה`, reason: 'paste-empty' };
    }
    return {
      ok: false,
      message: 'ההדבקה נכשלה: לוח המערכת ריק, או שהתוכן שבו אינו נתמך כאן. Ctrl+V מדביק אותו דרך המנוע.',
      reason: 'paste-source-unsupported',
    };
  }

  const destination = destinationRead && await destinationRead;
  if (destination && !destination.ok) return destination.outcome;
  let input: {
    payload?: ClipboardPayload;
    plan?: ClipboardPlan;
    fragment?: ClipboardPlan['fragment'];
    target: SelectionTarget;
  };
  const parse = clipboard?.parse;
  if (typeof parse === 'function') {
    let parsed: ClipboardParseResult | undefined;
    try {
      parsed = await parse(payload);
    } catch (error) {
      return { ok: false, message: thrownText(PASTE_FAILED, error), reason: 'threw' };
    }
    if (!parsed || parsed.success !== true) {
      const failure = parsed?.success === false ? parsed.failure : undefined;
      return { ok: false, message: pasteFailureText(failure), reason: failure?.code ?? 'parse-failed' };
    }
    const plan = destination?.ok ? reformatPlan(parsed.plan, mode, destination.marks) : parsed.plan;
    // מ-superdoc 2.23, `insert` של תכנית `plain-text` מתעלם מה-marks שבריצות
    // ומעתיק את עיצוב הריצה שבסמן. בסמן בלי ריצה (מסמך ריק, גופן שנבחר ועוד
    // לא הוקלד) הטקסט נוחת בלי עיצוב בכלל. נמדד: אותו fragment עם David 18
    // נכתב ב-2.20 ולא ב-2.23. כ-`fragment` ישיר ה-marks נכתבים בשתי הגרסאות.
    input =
      destination?.ok && parsed.plan.summary?.sourceKind === 'plain-text'
        ? { fragment: plan.fragment, target: at.target }
        : { plan, target: at.target };
  } else if (mode !== 'keep-source') {
    // בלי `parse` אין fragment לשכתב, ו-`insert` על payload גולמי היה מדביק
    // את עיצוב המקור — כלומר בדיוק את מה שהמשתמש ביקש שלא.
    return { ok: false, message: `${PASTE_FAILED}: ${UNAVAILABLE}`, reason: 'command-unsupported' };
  } else {
    // גרסה שאינה חושפת `parse`: `insert` מקבל payload גולמי ומפרק אותו בעצמו.
    input = { payload, target: at.target };
  }

  if (mode !== 'keep-source') {
    const current = await readSelectionSettled(selectionOf(host), { allowCollapsed: true, requireReady: true });
    const revision = await destinationDoc?.info?.();
    const latest = readSelection(selectionOf(host), { allowCollapsed: true });
    if (!latest.ok || !latest.ready || JSON.stringify(latest.target) !== JSON.stringify(at.target) || docOf(host) !== destinationDoc || !current.ok || !current.ready || JSON.stringify(current.target) !== JSON.stringify(at.target) || !destination?.ok || revision?.revision !== destination.revision) {
      return failed(PASTE_FAILED, 'not-ready');
    }
  }
  let receipt: ClipboardInsertReceipt | undefined;
  try {
    receipt = await insert(input);
  } catch (error) {
    return { ok: false, message: thrownText(PASTE_FAILED, error), reason: 'threw' };
  }

  if (receipt?.success === false) {
    return { ok: false, message: pasteFailureText(receipt.failure), reason: receipt.failure?.code };
  }

  // הסמן עובר לסוף מה שהודבק — אחרת ההדבקה הבאה נכנסת לפני זו. ההנמקה
  // המלאה ב-`caretAfterPaste`.
  if (docOf(host) === destinationDoc) await moveCaretAfterPaste(selectionOf(host), receipt);

  // הדבקה שנפלה לטקסט בלבד היא **הצלחה** עם אובדן עיצוב. אין לה ערוץ נפרד
  // אל המשתמש (המדווח מציג כשלים), ולכן היא נרשמת ללוג ולא נצבעת כשגיאה.
  if (receipt?.plan?.plainFallback === true) {
    console.info('[otzaria-word] ההדבקה נפלה לטקסט בלבד', receipt.diagnostics);
  }

  return { ok: true };
}

/* ------------------------------------------------------------------ */
/* מזג עיצוב ושמירת טקסט בלבד                                         */
/* ------------------------------------------------------------------ */

/**
 * „שמירת טקסט בלבד” מתחיל מהטקסט עצמו: `text/plain` בלבד. HTML, תמונות
 * וה-fragment הפרטי של המנוע נזרקים כבר כאן, כך שגם טבלה מגיעה כשורות טקסט
 * ולא כטבלה ריקה מעיצוב — כמו ב-Word. `null` = אין בלוח טקסט.
 */
function textOnlyPayload(payload: ClipboardPayload): ClipboardPayload | null {
  const text = textItem(payload, 'text/plain');
  if (!text) return null;
  return { source: payload.source, items: [{ type: 'text/plain', kind: 'string', data: text }] };
}

/**
 * העיצוב שבסמן, כפי שהרצועה מציגה אותו — דרך `readFormat`, אותה קריאה שמתג
 * הקיצורים האישיים משתמש בה. `empty` הוא צורת הבחירה: על טווח מעורב צבע שלא
 * דווח אינו „אין צבע” (ההנמקה שם).
 */
function readCaretFormat(host: ClipboardHostTarget, empty: boolean) {
  let commands: CommandsSurface | null | undefined;
  try {
    commands = (host as ClipboardHost | null | undefined)?.ui?.commands;
  } catch {
    commands = undefined;
  }

  const source: FormatReadingSource = {
    has: (id) => {
      try {
        return commands?.has?.(id) === true;
      } catch {
        return false;
      }
    },
    getState: (id) => {
      try {
        return (commands?.get?.(id)?.getState?.() ?? {}) as ReturnType<FormatReadingSource['getState']>;
      } catch {
        return {} as ReturnType<FormatReadingSource['getState']>;
      }
    },
  };

  return readFormat(source, { empty, settled: true });
}

type DestinationRead = { ok: true; marks: PasteRunMarks; revision: string } | { ok: false; outcome: CommandOutcome };

async function readDestination(
  host: ClipboardHostTarget,
  target: SelectionTarget,
  caret: ReturnType<typeof readCaretFormat>,
): Promise<DestinationRead> {
  try {
    const doc = docOf(host);
    const sample = destinationSample(target);
    if (!sample || !doc?.selection?.extractOoxml) throw new Error('Destination capture unavailable');
    const revisionRead = doc.info?.();
    const captureRead = Promise.resolve(doc.selection.extractOoxml({ at: sample })).then((capture) => ({ ok: true as const, capture }), (error: unknown) => ({ ok: false as const, error }));
    const info = await revisionRead;
    const revision = info?.revision;
    if (typeof revision !== 'string') throw new Error('Destination revision unavailable');
    let marks: PasteRunMarks;
    try {
      const result = await captureRead;
      if (!result.ok) throw result.error;
      if (result.capture.evaluatedRevision && result.capture.evaluatedRevision !== revision) throw new Error('Destination changed during capture');
      marks = marksFromDestinationCapture(result.capture);
    } catch (error) {
      // Export only a wholly empty, small standalone document. This reads
      // actual styles/default-character properties, never assumes missing CS
      // channels equal a Latin toolbar value. Populated documents are not exported.
      if (!info?.counts || info.counts.characters !== 0 || info.counts.paragraphs !== 1 || info.counts.tables !== 0 ||
          info.counts.images !== 0 || info.counts.lists !== 0 || info.counts.sdtFields !== 0 || !doc.export?.toDocx) throw error;
      if ((target.story && target.story.storyType !== 'body') || (target.start.kind === 'text' && target.start.story && target.start.story.storyType !== 'body')) throw error;
      const packageRead = await doc.export.toDocx({});
      if (!packageRead.contentBase64) throw error;
      marks = await marksFromEmptyDocument(packageRead.contentBase64, sample.start.kind === 'text' ? sample.start.blockId : '');
    }
    if (marks.fontSizePt === undefined || marks.fontSizeCsPt === undefined) throw new Error('Destination size unavailable');
    if (marks.bold !== undefined && marks.boldCs !== undefined && marks.bold !== marks.boldCs) throw new Error('Split destination bold state');
    if (marks.italic !== undefined && marks.italicCs !== undefined && marks.italic !== marks.italicCs) throw new Error('Split destination italic state');
    if (marks.highlight && caret.highlight == null) throw new Error('Pending highlight cancellation unsupported');
    if (marks.boldCs !== undefined) marks.boldCs = caret.bold;
    if (marks.italicCs !== undefined) marks.italicCs = caret.italic;
    // The toolbar shows one scalar. When Latin and Hebrew differ, a toolbar value
    // equal to either channel is that channel and not a pending choice; the
    // rewrite then gives each script its own (paste-formatting `scriptSegments`).
    // A value matching neither was chosen at the caret and wins for both.
    const families = [marks.fontFamily, marks.fontFamilyCs ?? marks.fontFamily].map((family) => family?.toLowerCase());
    if (caret.fontFamily && !families.includes(caret.fontFamily.toLowerCase())) {
      marks.fontFamily = caret.fontFamily;
      delete marks.fontFamilyCs;
    } else if (caret.fontFamily && marks.fontFamilyCs === undefined) marks.fontFamily = caret.fontFamily;
    if (caret.fontSizePt && (marks.fontSizePt === marks.fontSizeCsPt || ![marks.fontSizePt, marks.fontSizeCsPt].includes(caret.fontSizePt))) {
      marks.fontSizePt = marks.fontSizeCsPt = caret.fontSizePt;
    }
    return { ok: true, marks: { ...marks, ...destinationMarks(caret) }, revision };
  } catch (error) {
    // כל הסירובים מגיעים לאותה הודעה למשתמש; הסיבה עצמה נרשמת ל-console.error,
    // ומשם „דיווח על באג” מצרף אותה (host/recent-errors.ts).
    console.error(`[otzaria-word] מיזוג עיצוב נדחה: ${error instanceof Error ? error.message : String(error)}`);
    return { ok: false, outcome: { ok: false, reason: 'paste-destination-unsupported', message: 'ההדבקה לא בוצעה: אי אפשר לשמר בבטחה את עיצוב התווים בנקודה זו. אפשר להשתמש בשמירה על עיצוב המקור.' } };
  }
}

/** Rewriting is pure: no additional history/format mutations after insertion. */
function reformatPlan(plan: ClipboardPlan, mode: PasteMode, destination: PasteRunMarks): ClipboardPlan {
  const fragment = plan.fragment as unknown as PasteFragment;
  const rewritten = mode === 'merge' ? mergeFragment(fragment, destination, { mirrorHtmlEmphasis: plan.summary?.sourceKind === 'html' }) : textOnlyFragment(fragment, destination);
  return { ...plan, fragment: rewritten as unknown as ClipboardPlan['fragment'] };
}

/* ------------------------------------------------------------------ */
/* בחר הכל                                                            */
/* ------------------------------------------------------------------ */

const SELECT_ALL_FAILED = 'בחירת כל המסמך נכשלה';

/**
 * „בחר הכל”.
 *
 * שני מסלולים ציבוריים משורשרים, ואף אחד מהם אינו DOM:
 *
 * 1. `doc.ranges.resolve` עם עוגני `{kind:'document', edge:'start'|'end'}`.
 *    אלה בדיוק העוגנים שהחוזה מגדיר לגבולות גוף המסמך, והפלט הוא
 *    `SelectionTarget` — „transparent selection target”, כלומר הצורה שפעולות
 *    הבחירה והכתיבה מקבלות. אין צורך לרוץ על `blocks.list()` ולהרכיב קצוות
 *    בעצמנו; המנוע כבר יודע איפה המסמך מתחיל ונגמר.
 * 2. `ui.selection.apply(target)` — „apply a public selection target through
 *    the host-owned selection helper”. זה גם מה שהמנוע עצמו עושה כשהוא מחיל
 *    בחירה מבחוץ, והוא נכשל סגור עם `reason` מהטקסונומיה הציבורית במקום
 *    לזרוק.
 *
 * שני השלבים אינם מובטחים: `ranges` אינו חשוף בכל גרסה, ו-`apply` תלוי
 * ב-helper שה-host מספק. כשאחד מהם חסר הפקד מנוטרל עם הסבר — §12 — ולא
 * מנסה מסלול עקיף.
 */
export async function selectWholeDocument(host: ClipboardHostTarget): Promise<CommandOutcome> {
  const resolve = docOf(host)?.ranges?.resolve;
  if (typeof resolve !== 'function') {
    return { ok: false, message: `${SELECT_ALL_FAILED}: ${UNAVAILABLE}`, reason: 'command-unsupported' };
  }

  const apply = selectionOf(host)?.apply;
  if (typeof apply !== 'function') {
    return failed(SELECT_ALL_FAILED, 'host-capability-unavailable');
  }

  let target: SelectionTarget | null | undefined;
  try {
    const resolved = await resolve({
      start: { kind: 'document', edge: 'start' },
      end: { kind: 'document', edge: 'end' },
    });
    target = resolved?.target;
  } catch (error) {
    return { ok: false, message: thrownText(SELECT_ALL_FAILED, error), reason: 'threw' };
  }

  if (!target || typeof target !== 'object') {
    return failed(SELECT_ALL_FAILED, 'target-unresolved');
  }

  let result: { ok?: boolean; reason?: string } | undefined;
  try {
    result = await apply(target);
  } catch (error) {
    return { ok: false, message: thrownText(SELECT_ALL_FAILED, error), reason: 'threw' };
  }

  if (result?.ok !== true) return failed(SELECT_ALL_FAILED, result?.reason ?? 'target-unresolved');
  return { ok: true };
}
