/**
 * ההחלה של זיהוי הרשימות בהקלדה. ההחלטה עצמה — מי סמן ומי לא — יושבת
 * ב-`list-autoformat.ts` ונבדקת ביחידה; כאן רק החיווט למנוע.
 *
 * ## מה נבדק במנוע, ומה זה מכתיב (scripts/qa/list-autoformat-qa.mjs)
 *
 * 1. **קריאות מחכות לשקט.** `ranges.resolve`, `blocks.list`, `getNodeById`
 *    ו-`lists.getState` עונים רק כשהמנוע נרגע מההקשה האחרונה — ו-`getState`
 *    מאוחר מכולן. בהקלדה רציפה אף אחת מהן אינה חוזרת עד שההקלדה נעצרת. גרסה
 *    שקראה את הטקסט אחרי הרווח המירה רק כשהמשתמש הפסיק להקליד — בדיוק מה
 *    שדווח („זה עובר רק אחרי שאני גומר להקליד”).
 * 2. **מוטציות אינן מחכות.** `doc.insert` ו-`lists.create` באמצע הקלדה רציפה
 *    הוחלו מיד, הסמן החי עבר איתן, ואף תו לא אבד. `mutations.apply` — לא:
 *    אחריה הסמן נשאר מחוץ לטקסט והמנוע דחה כל הקשה שבאה אחריה.
 * 3. **`ui.selection.get()` סינכרוני.** ב-`keydown` (לפני שהמנוע רואה את
 *    המקש) הוא מחזיר את מצב הסמן אחרי ההקשות הקודמות — מדויק בקצב אנושי,
 *    אבל בהקלדה מהירה מאוד הוא מפגר **ומדווח `ready`**. אחרי Enter הוא
 *    `stale` עד שהמנוע קלט אותו.
 *
 * ## ולכן: ההקשות עצמן הן המקור, לא קריאת המסמך
 *
 * - **רצף** (`run`) — התווים שנלחצו מאז האיפוס האחרון (כל מקש אחר, לחיצת
 *   עכבר, הדבקה).
 * - **עוגן** — הוכחה שהרצף התחיל בהיסט 0 של בלוק ידוע. התו הראשון מראה
 *   היסט 0, **ו**הוכח שמקש האיפוס כבר נקלט: המצב שונה ממה שנראה כשנלחץ,
 *   והמנוע היה מדויק כשנלחץ. כל תו אחר כך חייב להראות בדיוק את מספר התווים
 *   שלפניו. כל אי-התאמה בכל הקשה — גם מחוץ לרצף — מכניסה ל„מצב זהיר”.
 * - **הנתיב המהיר**: עוגן, ו„פסקה רגילה” ידועה → ברגע שהמנוע קלט את כל
 *   ההקשות (היסט = אורך הסמן + מה שנלחץ אחריו) מוחקים ויוצרים. בלי קריאה.
 * - **„האם זו כבר רשימה”** נשמר לכל בלוק מקריאות רקע, שעונות ברגעי שקט.
 *   פסקה שנוצרה ב-Enter מפסקה רגילה יורשת „פסקה רגילה” בלי קריאה. כל מה
 *   שעשוי לשנות רשימות (עכבר, מחיקה, Tab, קיצורים, הדבקה) מוחק את הזיכרון.
 * - **הנתיב המאומת** — כשחסר עוגן או סוג: ממתין לשקט, קורא את סוג הבלוק ואת
 *   תחילתו, ומוודא שהסמן עומד בדיוק אחרי כל מה שהוקלד מאז האיפוס. כשהמנוע
 *   שקט, השוויון הזה הוא עצמו ההוכחה שהרצף התחיל בהיסט 0.
 *
 * ## למה `lists.create` ולא פקודת הרצועה
 *
 * `numbered-list` היא טוגל שיוצרת רשימה **עשרונית**, ותיקון הסגנון אחריה הוא
 * מוטציה שנייה עם ציור ביניהן („`א.` נהיה `1.` ורק אחרי זה מחליף”).
 * `lists.create` יוצר ישר בסגנון המבוקש, ואינו טוגל — ולכן גם אינו מסרב על
 * פריט רשימה קיים: נמדד שהוא מנתק אותו לרשימה חדשה. מכאן בדיקת סוג הבלוק.
 *
 * ## Backspace שייך למנוע
 *
 * Backspace בתחילת פריט מסיר את הרשימה ומשאיר את הטקסט
 * (`scripts/qa/list-backspace-owner-probe.mjs`), ולכן המודול אינו נוגע בו.
 *
 * ## Ctrl+Z מיד אחרי ההמרה — צעד אחד, כמו ב-Word
 *
 * ההמרה היא שלושה צעדי היסטוריה (נמדד, לכל סוג: המחיקה, יצירת הרשימה, והסגנון).
 * בלי קיבוץ, Ctrl+Z הראשון הציג רשימה **עשרונית** „1.”, השני פסקה ריקה, ורק
 * השלישי החזיר את „א) ”. ב-Word הקשה אחת מחזירה את הטקסט. לכן העומק נמדד
 * לפני ההמרה ואחריה (`history.get` עונה תוך ‎~1ms גם באמצע הקלדה — נמדד),
 * וכשהביטול מגיע לאותו עומק — מיד, או אחרי שבוטלה ההקלדה שבאה אחריה — שלושת
 * הצעדים מתבטלים יחד. שלוש קריאות `undo` שנשלחות יחד מסתיימות ב-‎~85ms בלי
 * שהסמן העשרוני יצויר בדרך (נמדד). ‏Ctrl+Y מחזיר את ההמרה באותה צורה.
 *
 * כל עוד יש המרה כזו בהיסטוריה, Ctrl+Z/Ctrl+Y עוברים כאן: ההחלטה אם לקבץ
 * תלויה בעומק, והעומק נקרא רק בקריאה א-סינכרונית — מאוחר מכדי להחליט אם לתת
 * למקש להמשיך. ביטול שאינו מגיע לעומק הזה הוא צעד אחד, דרך אותה היסטוריה.
 * הקבוצה נזרקת כשההיסטוריה ירדה מתחתיה (בוטלה בדרך אחרת) או התרחקה ממנה
 * יותר מ-`GROUP_HORIZON` צעדים.
 *
 * ‏Ctrl+Z אינו עריכה, ולכן אינו קוטם את צד ה„חזור” של ההיסטוריה — גם כשהקבוצה
 * כבר נצרכה והמקש חזר למנוע. בלי ההבחנה הזאת ביטול **שני** מחק את קבוצת
 * ה„חזור”, ו-Ctrl+Y שאחריו החזיר את מחיקת הסמן בלי הרשימה: פסקה ריקה שאינה
 * רשימה, שדרשה עוד שני Ctrl+Y כדי להתאושש (נמדד בבדיקת היחידה).
 *
 * ## הצורות שהמנוע ממיר בעצמו
 *
 * המנוע ממיר בעצמו „- ”, „* ”, „+ ” ו-„N. ” לכל מספר (נמדד, בלי המודול). שני
 * דברים יוצאים מזה: המתג „כבוי” לא כיבה אותן — מי שכותב „1. ” כטקסט לא קיבל את
 * מה שהמתג הבטיח — ובמצב „פעיל” „- ” יצא `•` של המנוע ולא ה-`-` המתועד כאן.
 *
 * אין הגדרה במנוע שמכבה את זה, והכלל שלו יושב במסלול הכנסת הטקסט: עצירת
 * `keydown` לבדה אינה עוזרת (נמדד), וגם הכנסת הרווח דרך ה-API אינה פתרון —
 * הסמן החי נשאר **לפני** הרווח, וההקלדה שאחריו נכנסת לפניו (נמדד, בשלושה
 * מסלולים). מה שכן עובד: עצירת `keydown` של הרווח, ובמקום ה-`beforeinput`
 * שנוצר ממנו — `beforeinput` מסוג `insertReplacementText` עם אותו רווח. המנוע
 * מכניס אותו במקום הסמן, לפי סדר ההקשות, **בלי** הכלל (נמדד גם בהקשות של
 * 15ms). אם המנוע אינו מטפל בו, נשלח `insertText` רגיל — הרווח לא הולך לאיבוד.
 *
 * ## הרווח נעצר לפי אותו כלל בשני מצבי המתג
 *
 * כל צורה שהמנוע ממיר בעצמה נעצרת — גם כשהמתג דלוק. קודם נעצרו במצב „דלוק”
 * רק הצורות שיש להן תוכנית כאן, והתוצאה הייתה שהמתג הדלוק נתן **יותר** רשימות
 * לא-רצויות מהכבוי: „5. ”, „12. ” ו„1948. ” נדחים ב-`list-autoformat.ts`
 * („פותחי רצף בלבד”) — והמנוע המיר אותם במקומנו. מי שפתח פסקה ב„1948. ” קיבל
 * רשימה עשרונית. עכשיו מה שאין לו תוכנית כאן נשאר טקסט, ו„רק פותח רצף הופך
 * לרשימה” הוא התנהגות המוצר ולא רק של המודול הטהור.
 *
 * ## ההחלטה לעצור היא סינכרונית
 *
 * „מה שלפני הסמן” אינו רק מה שהוקלד מאז האיפוס: תיקון באמצע הסימן — „1x”,
 * Backspace, „. ” — מאפס את הרצף, והמנוע המשיך להמיר גם כשהמתג כבוי (נמדד).
 * הקריאה שנוספה לכך (תחילת הפסקה מהמנוע, והזנב מהמקלדת מצטרף אליה) אינה
 * יכולה **לשלול** עצירה: היא אינה מגיעה בזמן כשהקריאות חסומות, ונמדד שהיא גם
 * חוזרת על מצב ישן מהמקש האחרון (Backspace איטי → הקריאה החזירה „1x”, הזנב
 * „1.” נדבק אחריו, ו„1x1.” אינו סימן — הרווח עבר, והמנוע המיר עם מתג כבוי).
 *
 * לכן העצירה נשענת על הזנב לבדו: `ENGINE_MARKER` עוגן ב-`$`, ולכן זנב שאינו
 * סיומת אפשרית שלו הוא **הוכחה** שאין סימן לפני הסמן, בלי שום קריאה. זנב
 * שהוא סיומת אפשרית — עוצרים. הקריאה נשארת המקור של ה**המרה** (מה הסימן
 * במלואו), ושם היא בטוחה: ההחלה עוברת בנתיב המאומת, שקורא שוב ומשווה.
 *
 * ## החלון שבין ההחלטה למחיקה
 *
 * ‏`apply` מחליטה מסמן שנקרא לפניה, וממתינה ל-`history.get` (‎~1ms) לפני
 * שהיא מוחקת. המחיקה חותכת **היסטים קבועים** — 0 עד אורך הסימן — בבלוק ידוע,
 * ולכן היא אינה מתגוננת בעצמה: לחיצה או הקשה בתוך החלון מאפסות את המצב אבל
 * אינן יכולות לבטל קריאה שכבר רצה, ומה שיימחק יהיה טקסט אחר. `markerHolds`
 * הוא ההוכחה שנלקחת שוב רגע לפני המחיקה, והיא **סינכרונית בהכרח**: כל קריאה
 * מהמסמך כאן מחזירה את „ממיר רק אחרי שמפסיקים להקליד” (סעיף 1 למעלה).
 *
 * כל `await` שיתווסף ל-`apply` לפני המחיקה מרחיב את החלון הזה.
 *
 * ## כשל
 *
 * המשתמש לא ביקש רשימה — הוא הקליד. כשל בין מחיקת הסמן ליצירת הרשימה מחזיר
 * את הטקסט שנמחק ושותק.
 */
import { PASTE_STARTED } from './paste-events';
import { invalidateDocumentReads } from './document-read-state';
import type { SuperDoc } from 'superdoc';
import type { DocReceipt, MaybePromise } from './document-api';
import { planListAutoformat, type ListAutoformat } from './list-autoformat';
import { resolveListItemAt, type ListsTarget } from './lists';
import type { ResolvedRangeLike, SelectionPointLike, SelectionTargetLike } from './word-selection';

/** הסמן הארוך ביותר בלי תו ההפעלה: `(א)`. */
const LONGEST_MARKER = 3;

/** כל כמה לבדוק אם המנוע כבר קלט את הרווח. הבדיקה סינכרונית וזולה. */
const WATCH_INTERVAL_MS = 15;

/** אחרי כמה זמן בלי התאמה מדויקת עוברים לנתיב המאומת. */
const WATCH_LIMIT_MS = 1200;

/** כמה זמן אחרי אי-התאמה הנתיב המהיר כבוי. */
const CAUTIOUS_MS = 2000;

/** אחרי שקט כזה המנוע כבר קלט כל מה שנלחץ. */
const IDLE_MS = 300;

/** אחרי לחיצת עכבר הבחירה `stale` כמה מאות מילישניות; זה המרווח שמעליהן. */
const POINTER_SETTLE_MS = 500;
const SETTLE_POLL_MS = 60;
const SETTLE_POLLS = 10;

/** נסיונות של הנתיב המאומת: המשתמש עשוי לחזור להקליד בין הקריאות. */
const VERIFY_ATTEMPTS = 3;

const MODIFIER_KEYS = new Set(['Shift', 'Control', 'Alt', 'AltGraph', 'Meta', 'CapsLock', 'NumLock', 'ScrollLock', 'Fn', 'OS']);

/** מזיזים סמן בלי לשנות אף בלוק — מאפסים את הרצף, לא את הזיכרון. */
const NAVIGATION_KEYS = new Set(['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End', 'PageUp', 'PageDown']);

const EDIT_EVENTS = [PASTE_STARTED, 'paste', 'cut', 'drop', 'compositionstart'];

/** כמה תווים נשמרים מאז האיפוס לבדיקת `ENGINE_MARKER`. */
const ENGINE_MARKER_MAX = 12;

/**
 * הרווחים שלפני הסימן — חסומים ל-`ENGINE_MARKER_MAX - 1`, ולא `\s*`.
 *
 * ‏`\s*` בלתי-חסום סתר את התקרה: `typedSince` הופך ל-`null` אחרי
 * ‏`ENGINE_MARKER_MAX` תווים, ואז `interceptsSpace` מחזיר `false` — כלומר
 * הקוד הכריז על כיסוי שהוא אינו יכול לספק. שתי דרכים לפתור, והמדידה
 * הכריעה: **המנוע עצמו אינו ממיר שם**. נמדד (17.9.2026, superdoc 2.15.0,
 * המתג כבוי, על ה-dist הארוז) בדיוק בטווח שבו המודול אינו מיירט —
 * ‏11, 12 ו-16 רווחים מובילים ואז „1. ” — ובכל השלושה הבלוק נשאר
 * ‏`paragraph`. (`docs/engine-gaps.md` מדד קודם שעם רווח מוביל **אחד** הוא
 * כן ממיר.) כלומר אין חור לכסות, והחסם כאן רק אומר את האמת על הכיסוי.
 *
 * במסלול הזנב החסם אינו משנה דבר — שם התקרה כבר חותכת — והמקום היחיד שבו
 * הוא נצפה הוא `ENGINE_MARKER.test(known.text)` כשהסמן אינו נקרא (‏`at`
 * הוא `null`, ואז שומר ההיסט מדולג). גם שם החדש הוא הנכון, מאותה מדידה.
 *
 * הוא נבנה מהתקרה ולא נכתב כמספר — כדי ששינוי של אחד מהם לא ישאיר את השני
 * מאחור.
 */
const MARKER_SPACES = `\\s{0,${ENGINE_MARKER_MAX - 1}}`;

/** מה שהמנוע ממיר בעצמו כשבא אחריו רווח (נמדד). */
const ENGINE_MARKER = new RegExp(`^${MARKER_SPACES}(?:[-+*]|\\d{1,9}\\.)$`);

/**
 * כל מה שעשוי **להשלים** את `ENGINE_MARKER` — כלומר כל סיומת שלו, כולל הריקה.
 *
 * זה השער הסינכרוני היחיד: `ENGINE_MARKER` עוגן ב-`$`, ולכן אם הזנב שהוקלד
 * מאז האיפוס אינו מתאים כאן, גם המחרוזת המלאה שלפני הסמן אינה סימן — ואין
 * צורך לדעת מה היה לפניה. הספרות כאן `{0,9}` ולא `{1,9}` בדיוק בשביל „.”
 * שהוקלד אחרי ש„1” כבר נכתב לפני האיפוס.
 *
 * ## הרווח הוא חלק מהסוגריים, ולא סיומת בפני עצמה
 *
 * הצורה הקודמת — `/^\s*(?:[-+*]|\d{0,9}\.)?$/` — קיבלה גם זנב שכולו רווחים,
 * ואין כזה: כל מחרוזת שתואמת את `ENGINE_MARKER` נגמרת ב-`-`, `+`, `*` או
 * „.”, ולכן אף סיומת לא-ריקה שלה אינה רווח בלבד. הסלחנות הזאת הייתה **נראית
 * למשתמש**: אחרי `Enter` הרווח הראשון אינו נעצר (זנב ריק, היסט 0), אבל הזנב
 * הופך ל-„ ” — ומכאן הרווח ה**שני** וכל רווח אחריו נעצרו. שומרי ההיסט וסוג
 * הבלוק יושבים רק בענף של הזנב הריק, ולכן זה חל בכל היסט ובכל בלוק: רווח
 * כפול בהיסט 400, בתוך תא טבלה, בכותרת או בתוך פריט רשימה קיים — כולם עברו
 * דרך `stopImmediatePropagation` ודרך ה-`beforeinput` הסינתטי, מסלול שנמדד
 * לסימן בתחילת פסקה רגילה בלבד.
 *
 * הצורה כאן נבדקה בכוח גס מול ההגדרה שהיא אמורה לכסות: 3991 סימנים (כל
 * קידומת רווחים באורך ≤3 מעל שישה תווי `\s` שונים, וגם קידומות של תו חוזר
 * עד קצה החסם, × 13 גופי סימן), 4011 סיומות נבדלות — **0 חורים**, ו-39
 * התאמות שקר פחות, כולן מהמחלקה
 * „רווחים בלבד”. מיקום ה-`?` הוא כל ההבדל: הוא עוטף עכשיו את הרווחים יחד
 * עם הגוף, ולא רק את הגוף. הבדיקה חוזרת ב-`tests/unit/`.
 */
const ENGINE_MARKER_TAIL = new RegExp(`^(?:${MARKER_SPACES}(?:[-+*]|\\d{0,9}\\.))?$`);

/** Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z — לפי `code`, כדי שיעבדו גם בפריסה עברית. */
function historyKey(event: KeyboardEvent): 'undo' | 'redo' | null {
  if (!(event.ctrlKey || event.metaKey) || event.altKey) return null;
  // `code` ולא `key`: בפריסה עברית `key` של Ctrl+Z הוא „ז” (ראו
  // tests/contract/shortcut-registry.test.ts).
  const letter = event.code === 'KeyZ' ? 'z' : event.code === 'KeyY' ? 'y' : null;
  if (letter === 'z') return event.shiftKey ? 'redo' : 'undo';
  if (letter === 'y' && !event.shiftKey) return 'redo';
  return null;
}

/** הרווח, כהכנסה שהמנוע מבצע בלי כללי ההקלדה שלו. `false` — לא טופל. */
function insertPlainSpace(target: EventTarget, ours: WeakSet<Event>): boolean {
  const send = (inputType: string, withData: boolean): boolean => {
    const init: InputEventInit = { bubbles: true, cancelable: true, composed: true, inputType, data: ' ' };
    if (withData && typeof DataTransfer === 'function') {
      try {
        const transfer = new DataTransfer();
        transfer.setData('text/plain', ' ');
        init.dataTransfer = transfer;
      } catch {
        /* בלי העברה — המנוע קורא גם את `data` */
      }
    }
    const event = new InputEvent('beforeinput', init);
    ours.add(event);
    // `dispatchEvent` מחזיר `false` כשמישהו ביטל את ברירת המחדל — כלומר המנוע
    // קלט את ההכנסה.
    return !target.dispatchEvent(event);
  };
  if (send('insertReplacementText', true)) return true;
  return send('insertText', false);
}

/** `paragraph` — פסקה רגילה, בלי מספור. כל השאר אינו מומר. */
type BlockKind = 'paragraph' | 'other';

interface BlockListing {
  nodeId?: string;
  nodeType?: string;
  paragraphNumbering?: unknown;
  numbering?: unknown;
}

export interface ListAutoformatDoc {
  ranges?: { resolve?: (input: unknown) => MaybePromise<ResolvedRangeLike | undefined> };
  blocks?: {
    list?: (input?: Record<string, unknown>) => MaybePromise<{ blocks?: BlockListing[] } | undefined>;
  } | null;
  insert?: (input: { value: string; type: 'text'; target?: unknown }) => MaybePromise<DocReceipt>;
  lists?: {
    create?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
  } | null;
  history?: {
    get?: () => MaybePromise<{ undoDepth?: unknown; redoDepth?: unknown } | null | undefined>;
    undo?: () => MaybePromise<unknown>;
    redo?: () => MaybePromise<unknown>;
  } | null;
}

export interface UiSelectionLike {
  status?: string;
  selectionTarget?: SelectionTargetLike | null;
}

export interface ListAutoformatHost {
  activeEditor?: { doc?: ListAutoformatDoc | null } | null;
  ui?: { selection?: { get?: () => UiSelectionLike | null | undefined } | null } | null;
}

export interface ListAutoformatOptions {
  /** ה-container של המסמך — ראו create-editor.ts:EditorSession.container. */
  container: HTMLElement;
  host: SuperDoc | ListAutoformatHost | null | undefined;
  /**
   * המתג. `false` — אין המרה, ורק ההמרות של המנוע עצמו נחסמות. ברירת המחדל
   * דלוקה, כמו ב-host/settings.ts.
   */
  enabled?: boolean;
}

export interface ListAutoformatHandle {
  /**
   * „בטל” שאינו עובר במקלדת (פס הכותרת). `true` — יש המרה בהיסטוריה, והביטול
   * יצא לדרך כאן (מקובץ או צעד אחד); אין להריץ עוד ביטול.
   */
  undo(): boolean;
  /** „חזור”, באותו תנאי. */
  redo(): boolean;
  /**
   * הזזת המתג בלי להתקין מחדש — ההתקנה מחדש היתה מאבדת את קבוצת
   * הביטול של ההמרה האחרונה, ו-Ctrl+Z אחרי כיבוי המתג השאיר רשימה על המסך (נמדד).
   */
  setEnabled(value: boolean): void;
  dispose(): void;
}

function docOf(host: ListAutoformatOptions['host']): ListAutoformatDoc | null {
  return (host as ListAutoformatHost | null | undefined)?.activeEditor?.doc ?? null;
}

function point(blockId: string, offset: number, story: unknown): SelectionPointLike {
  const at: SelectionPointLike = { kind: 'text', blockId, offset };
  if (story !== undefined && story !== null) at.story = story;
  return at;
}

/** הסמן המכווץ בלבד: בחירה של טווח אינה הקלדה. */
interface Caret {
  blockId: string;
  offset: number;
  story: unknown;
}

function readCaret(target: SelectionTargetLike | null | undefined): Caret | null {
  if (!target || target.kind !== 'selection') return null;
  if (target.coordinateSpace !== undefined && target.coordinateSpace !== 'visible') return null;

  const { start, end } = target;
  if (start?.kind !== 'text' || end?.kind !== 'text') return null;
  if (typeof start.blockId !== 'string' || start.blockId !== end.blockId) return null;
  if (typeof start.offset !== 'number' || start.offset !== end.offset) return null;

  return { blockId: start.blockId, offset: start.offset, story: target.story ?? start.story ?? null };
}

function sameCaret(a: Caret, b: Caret): boolean {
  return a.blockId === b.blockId && a.offset === b.offset;
}

function textTarget(at: Caret, from: number, to: number): SelectionTargetLike {
  return {
    kind: 'selection',
    start: point(at.blockId, from, at.story),
    end: point(at.blockId, to, at.story),
    ...(at.story ? { story: at.story } : {}),
  };
}

function isBodyStory(story: unknown): boolean {
  return !story || (story as { storyType?: unknown }).storyType === 'body';
}

/**
 * הרמות שנכתבות. **רמה 0 בלבד** — בשונה מ-`setListNumberStyle` ב-lists.ts,
 * שנקראת מבחירה מפורשת בגלריה ולכן דורסת את הקסקדה כולה. גם Word כותב את
 * התבנית שהוקלדה ל-`ListLevels(1)` בלבד (נמדד).
 */
function levelsFor(plan: ListAutoformat): Array<Record<string, unknown>> {
  if (plan.kind === 'bullet') {
    return [{ level: 0, numFmt: 'bullet', lvlText: plan.lvlText, markerFont: plan.markerFont }];
  }
  // `markerFont: ''` הוא הריקון היחיד שהחוזה מקבל (`null` נדחה) — בלעדיו סמן
  // ממוספר יורש את גופן הסמלים של תבליט שקדם לו.
  return [{ level: 0, numFmt: plan.numFmt, lvlText: plan.lvlText, markerFont: '' }];
}

/** קריאה אחת שלעולם אינה זורקת. `NO_OP` הוא הצלחה. */
async function call(run: () => MaybePromise<DocReceipt>): Promise<boolean> {
  try {
    const receipt = await run();
    return !(receipt?.success === false && receipt.failure?.code !== 'NO_OP');
  } catch {
    return false;
  }
}

/**
 * ידית QA: בלעדיה „לא קרה כלום” אינו ניתן להבחנה מ„המודול לא הותקן”. אין לה
 * קורא בקוד האפליקציה.
 */
interface AutoformatDebug {
  installs: number;
  evaluates: number;
  applies: number;
  last: string;
  /** שתים-עשרה הסיבות האחרונות. */
  trail: string[];
}

function debugHandle(): AutoformatDebug | null {
  if (typeof window === 'undefined') return null;
  const holder = window as unknown as { __otzariaListAutoformat?: AutoformatDebug };
  holder.__otzariaListAutoformat ??= { installs: 0, evaluates: 0, applies: 0, last: '', trail: [] };
  return holder.__otzariaListAutoformat;
}

/**
 * איפה הרצף מתחיל בהיסט 0: בלוק ידוע, או „הבלוק ש-Enter יצר” — שאינו ידוע
 * עד שהמנוע קלט את ה-Enter, אבל הסמן בו מתחיל תמיד בהיסט 0.
 */
type Anchor = Caret | { afterEnter: string };

function anchoredAt(anchor: Anchor | null, at: Caret | null): Caret | null {
  if (!anchor || !at) return null;
  if ('afterEnter' in anchor) return at.blockId !== anchor.afterEnter ? at : null;
  return at.blockId === anchor.blockId ? at : null;
}

/** סמן שהוקלד ומחכה להמרה. */
interface Pending {
  typed: string;
  plan: ListAutoformat;
  anchor: Anchor | null;
  /** תווים שנלחצו אחרי תו ההפעלה — חלק מההיסט הצפוי. ממשיך לספור עד ההמרה. */
  keysAfter: number;
  since: number;
  /** Tab: רק דרך הנתיב המאומת — השלמה מהספר עשויה לבלוע אותו ולהכניס טקסט. */
  verifyOnly: boolean;
}

/** האיפוס שלפני הרצף, והמצב שנראה כשקרה. */
interface ResetMark {
  kind: 'install' | 'key' | 'pointer';
  seen: Caret | null;
  /** המנוע היה מדויק כשהאיפוס נלחץ. */
  exact: boolean;
  /** כמה מקשי איפוס ברצף, בלי הקלדה או שקט ביניהם. */
  count: number;
  at: number;
  enter: boolean;
}

/** המרה שאפשר לבטל בהקשה אחת: כמה צעדים, ובאיזה עומק היסטוריה. */
interface HistoryGroup {
  steps: number;
  depth: number;
  at: Caret;
  /** מצב המסמך שבצד ההיסטוריה הנוכחי, והמצב שיחזור אחריו. */
  current: { kind: BlockKind; text: string };
  other: { kind: BlockKind; text: string };
}

interface AutoformatHistory {
  undoGroup: HistoryGroup | null;
  redoGroup: HistoryGroup | null;
  replaying: Promise<void>;
}

// The active-tab watcher reinstalls keyboard listeners, but the document's
// history survives. Keep grouping beside that document, with weak ownership so
// closing a document does not retain its editor or history.
const documentHistories = new WeakMap<ListAutoformatDoc, AutoformatHistory>();

function historyFor(doc: ListAutoformatDoc | null): AutoformatHistory {
  let state = doc ? documentHistories.get(doc) : undefined;
  if (!state) {
    state = { undoGroup: null, redoGroup: null, replaying: Promise.resolve() };
    if (doc) documentHistories.set(doc, state);
  }
  return state;
}

/** כמה צעדים מעל ההמרה עוד שווה לחכות לה. מעבר לזה — המקש שוב של המנוע. */
const GROUP_HORIZON = 200;

async function historyDepth(doc: ListAutoformatDoc | null): Promise<{ undo: number; redo: number } | null> {
  const get = doc?.history?.get;
  if (typeof get !== 'function') return null;
  try {
    const state = await get.call(doc!.history);
    const undo = state?.undoDepth;
    const redo = state?.redoDepth;
    return typeof undo === 'number' && typeof redo === 'number' ? { undo, redo } : null;
  } catch {
    return null;
  }
}

export function installListAutoformat(options: ListAutoformatOptions): ListAutoformatHandle {
  const { container, host } = options;
  let enabled = options.enabled !== false;
  const debug = debugHandle();
  if (debug) debug.installs += 1;

  let disposed = false;
  let applying = false;
  /** הוקלד משהו בזמן ההמרה — ההיסטוריה שלה אינה רצף נקי. */
  let typedWhileApplying = false;
  /**
   * התווים מאז האיפוס, ‏ש-Backspace מוחק מהם תו במקום לאפס אותם. תיקון טעות
   * באמצע סימן הוא המקרה השכיח של רצף שנקטע, והוא ידוע מהמקלדת בלבד —
   * בלי להמתין לקריאה מהמנוע (שנמדדה ב-242ms אחרי Backspace). `null` — ארוך
   * מכדי להיות סימן.
   */
  let typedSince: string | null = '';
  /** תחילת הפסקה כמו שנקראה מהמנוע אחרי האיפוס האחרון. `null` — לא נקראה. */
  let resetPrefix: { blockId: string; text: string } | null = null;
  /** כל איפוס מבטל קריאה שלא חזרה עדיין. בנפרד מ-`generation`, ששייך לסוגי הבלוקים. */
  let prefixToken = 0;
  /**
   * כל איפוס מבטל גם המרה שכבר החלה. `apply` לוקחת אותו בכניסה ובודקת אותו
   * שוב רגע לפני המחיקה — ראו `markerHolds`.
   */
  let resetToken = 0;
  /** ה-`keydown` של רווח נעצר, וה-`beforeinput` שלו יוחלף. */
  let spaceArmed = false;
  const ours = new WeakSet<Event>();
  const historyState = historyFor(docOf(host));
  /** התווים מאז האיפוס. `null` — הוקלד משהו שאינו יכול להיות סמן. */
  let run: string | null = '';
  let anchor: Anchor | null = null;
  let pending: Pending | null = null;
  let watchTimer: ReturnType<typeof setTimeout> | undefined;
  const timers = new Set<ReturnType<typeof setTimeout>>();

  let resetMark: ResetMark = { kind: 'install', seen: null, exact: true, count: 1, at: now(), enter: false };
  /** המצב שנראה בתו הקודם, כשלא היה איפוס מאז. */
  let lastTyped: Caret | null = null;
  /** ההקשה או הלחיצה האחרונה. */
  let lastInputAt = -Infinity;
  let cautiousUntil = 0;
  /** למה הרצף האחרון לא עוגן — לידית ה-QA בלבד. */
  let anchorMiss = '';

  const kinds = new Map<string, BlockKind>();
  const reading = new Set<string>();
  let generation = 0;
  /** הבלוק שבו נלחץ Enter, כשידוע שהוא פסקה רגילה. חי עד האיפוס הבא. */
  let inheritFrom: string | null = null;

  function now(): number {
    return typeof performance !== 'undefined' ? performance.now() : Date.now();
  }

  function note(reason: string): void {
    if (!debug) return;
    debug.last = reason;
    debug.trail.push(reason);
    if (debug.trail.length > 12) debug.trail.shift();
  }

  function later(task: () => void, ms: number): void {
    const timer = setTimeout(() => {
      timers.delete(timer);
      if (!disposed) task();
    }, ms);
    timers.add(timer);
  }

  function caretNow(): Caret | null {
    let selection: UiSelectionLike | null | undefined;
    try {
      selection = (host as ListAutoformatHost | null | undefined)?.ui?.selection?.get?.();
    } catch {
      return null;
    }
    if (!selection || selection.status !== 'ready') return null;
    return readCaret(selection.selectionTarget);
  }

  function caretSoon(polls = SETTLE_POLLS): Promise<Caret | null> {
    const at = caretNow();
    if (at || polls <= 0) return Promise.resolve(at);
    return new Promise((resolve) => later(() => void caretSoon(polls - 1).then(resolve), SETTLE_POLL_MS));
  }

  function forgetKinds(): void {
    generation += 1;
    kinds.clear();
    reading.clear();
  }

  function startRun(): void {
    run = '';
    anchor = null;
    lastTyped = null;
    inheritFrom = null;
    typedSince = '';
  }

  /**
   * הסימן המלא שלפני הסמן — מה שנקרא מהפסקה אחרי האיפוס, ואחריו הזנב
   * שהוקלד מאז. `null` — הזנב ארוך מכדי להיות חלק מסימן.
   *
   * זהו המקור של **ההמרה** כשרצף ההקלדה נקטע („1x”, Backspace, „. ” — מה
   * שלפני הסמן הוא „1.”, והמקלדת לבדה יודעת רק „.”). הוא אינו המקור של
   * העצירה: הקריאה עשויה לאחר או לחזור על מצב ישן, ולכן `interceptsSpace`
   * אינו נשען עליה. כל המרה שנשענת עליה עוברת דרך הנתיב המאומת, שקורא שוב.
   */
  function markerBeforeCaret(at: Caret | null): string | null {
    const typed = typedSince;
    const known = resetPrefix;
    if (!known || (at && at.blockId !== known.blockId)) return typed;
    const tail = typed ?? '';
    // הקריאה עשויה להכלול חלק מהזנב או את כולו: מחברים רק את מה שחסר.
    let overlap = Math.min(tail.length, known.text.length);
    while (overlap > 0 && !known.text.endsWith(tail.slice(0, overlap))) overlap -= 1;
    return known.text + tail.slice(overlap);
  }

  /**
   * האם לעצור את הרווח הזה — ההסבר בהערת הפתיחה. אותו כלל בשני מצבי המתג:
   * כל צורה שהמנוע ממיר בעצמו נעצרת, וההמרה (אם יש תוכנית) היא שלנו.
   *
   * הכול כאן סינכרוני, לפי סדר הוודאות:
   *
   * 1. זנב שאינו סיומת אפשרית של `ENGINE_MARKER` — הוכחה שאין סימן.
   * 2. זנב שהוא סיומת אפשרית ואינו ריק — עוצרים. זה כל מה שהמנוע יראה
   *    בקצה, ואין צורך לדעת מה קדם לו.
   * 3. זנב ריק — הסימן, אם ישנו, קדם כולו לאיפוס. סמן בהיסט 0 (אין מה
   *    שיקדם) או רחוק מכל סימן אפשרי — שולל; קריאה שכבר חזרה מכריעה;
   *    ובלעדיה עוצרים, כי „1.” שהוקלד לפני לחיצה הוא בדיוק המקרה שבו
   *    המנוע ממיר (נמדד).
   */
  function interceptsSpace(at: Caret | null): boolean {
    const tail = typedSince;
    if (tail === null || !ENGINE_MARKER_TAIL.test(tail)) return false;
    if (tail !== '') return true;
    if (at && (at.offset === 0 || at.offset > ENGINE_MARKER_MAX)) return false;
    const known = resetPrefix;
    if (!known || (at && at.blockId !== known.blockId)) return true;
    return ENGINE_MARKER.test(known.text);
  }

  /**
   * ביטול או חזרה כשיש המרה בהיסטוריה: כל הקבוצה, כשהעומק הוא שלה; אחרת צעד
   * אחד. `null` אחרי זה — הקבוצה אינה רלוונטית עוד.
   */
  async function replay(kind: 'undo' | 'redo'): Promise<void> {
    const doc = docOf(host);
    const history = doc?.history;
    const step = kind === 'undo' ? history?.undo : history?.redo;
    if (typeof step !== 'function') return;
    const group = kind === 'undo' ? historyState.undoGroup : historyState.redoGroup;
    const before = await historyDepth(doc);
    const depth = before && (kind === 'undo' ? before.undo : before.redo);
    let whole = group !== null && depth === group.depth;
    // עומק לבדו אינו זהות של צעדי היסטוריה: ביטול זר ואחריו הקלדה יכולים
    // להחזיר אותו לאותו מספר. לפני שמבטלים קבוצה שלמה דורשים את מצב המסמך
    // שהקבוצה עצמה יצרה. כשאין ודאות, צעד אחד הוא התנהגות המנוע הבטוחה.
    if (whole && group) {
      const actual = await readKind(doc!, group.at);
      const text = await readPrefix(doc!, group.at, group.current.text.length + 1);
      if (actual !== group.current.kind || text !== group.current.text) {
        whole = false;
        if (kind === 'undo') historyState.undoGroup = null;
        else historyState.redoGroup = null;
      }
    }

    const calls: Promise<unknown>[] = [];
    const steps = whole && group ? group.steps : 1;
    for (let i = 0; i < steps; i += 1) {
      try {
        calls.push(Promise.resolve(step.call(history)));
      } catch {
        break;
      }
    }
    // This app-owned history path bypasses the engine's first-party keyboard
    // route. Invalidate after settlement, even when no editor-update was emitted.
    await Promise.allSettled(calls).finally(() => invalidateDocumentReads(doc));
    forgetKinds();
    const after = await historyDepth(doc);
    if (whole && group && after) {
      // הקבוצה עברה לצד השני של ההיסטוריה.
      const moved: HistoryGroup = {
        steps: group.steps,
        depth: kind === 'undo' ? after.redo : after.undo,
        at: group.at,
        current: group.other,
        other: group.current,
      };
      if (kind === 'undo') {
        historyState.undoGroup = null;
        historyState.redoGroup = moved;
      } else {
        historyState.redoGroup = null;
        historyState.undoGroup = moved;
      }
      note(`${kind}:grouped`);
      return;
    }
    note(`${kind}:single`);
    if (after) pruneGroups(after);
  }

  /** קבוצה שההיסטוריה עברה אותה, או התרחקה ממנה, אינה רלוונטית עוד. */
  function pruneGroups(depth: { undo: number; redo: number }): void {
    const { undoGroup, redoGroup } = historyState;
    if (undoGroup && (depth.undo < undoGroup.depth || depth.undo - undoGroup.depth > GROUP_HORIZON)) {
      historyState.undoGroup = null;
    }
    if (redoGroup && (depth.redo < redoGroup.depth || depth.redo - redoGroup.depth > GROUP_HORIZON)) {
      historyState.redoGroup = null;
    }
  }

  /** `true` — המקש שלנו: יש המרה בצד הזה של ההיסטוריה. */
  function takeHistory(kind: 'undo' | 'redo'): boolean {
    if (!(kind === 'undo' ? historyState.undoGroup : historyState.redoGroup)) return false;
    historyState.replaying = historyState.replaying.then(() => replay(kind)).catch(() => {});
    return true;
  }

  /** עריכה חדשה מוחקת את צד ה„חזור” של ההיסטוריה, ואיתו את הקבוצה שבו. */
  function dropRedoGroup(): void {
    historyState.redoGroup = null;
  }

  async function readKind(doc: ListAutoformatDoc, at: Caret): Promise<BlockKind | null> {
    const list = doc.blocks?.list;
    if (typeof list === 'function') {
      try {
        const input: Record<string, unknown> = { nodeIds: [at.blockId] };
        if (!isBodyStory(at.story)) input.in = at.story;
        const listed = await list(input);
        const block = listed?.blocks?.find((b) => b.nodeId === at.blockId);
        if (block) {
          return block.nodeType === 'paragraph' && !block.paragraphNumbering && !block.numbering
            ? 'paragraph'
            : 'other';
        }
      } catch {
        // נופלים ל-getState: היא מכסה גם פסקאות בתוך טבלה.
      }
    }
    const state = await resolveListItemAt(host as ListsTarget, at.blockId);
    if (state.kind === 'not-list') return 'paragraph';
    return state.kind === 'item' ? 'other' : null;
  }

  function learnKind(at: Caret): void {
    // כשהמתג כבוי אין המרה, ולכן גם אין בשביל מה לקרוא את סוג הבלוק.
    if (!enabled || kinds.has(at.blockId) || reading.has(at.blockId)) return;
    const doc = docOf(host);
    if (!doc) return;
    const asked = generation;
    reading.add(at.blockId);
    void readKind(doc, at).then((kind) => {
      if (disposed || asked !== generation) return;
      reading.delete(at.blockId);
      if (kind) kinds.set(at.blockId, kind);
    });
  }

  function learnWhenSettled(): void {
    void caretSoon().then((at) => {
      if (at && !disposed) learnKind(at);
    });
  }

  /**
   * מה שכתוב בתחילת הפסקה, אחרי שרצף ההקלדה נקטע.
   *
   * Backspace, חץ או לחיצה מאפסים את הרצף שנשמר מהמקלדת, אבל המנוע קורא
   * את הפסקה וממשיך להמיר (נמדד: „1x”, Backspace, „. ” — פריט רשימה עם סמן „1.”,
   * גם כשהמתג כבוי). לכן קוראים את התחלה מהמנוע כשהוא נרגע, ועד שהיא חוזרת
   * ההתנהגות היא הקודמת. רק סמוך לתחילת הפסקה — רחוק משם אין סימן אפשרי.
   */
  function learnPrefix(): void {
    resetPrefix = null;
    const doc = docOf(host);
    if (!doc) return;
    const token = (prefixToken += 1);
    void caretSoon().then(async (first) => {
      // הסמן זז בהיותר תו אחד באיפוס עצמו, וסימן ארוך מכאן אינו קיים.
      if (disposed || token !== prefixToken || !first || first.offset > ENGINE_MARKER_MAX + 1) return;
      // כל קריאה מהמנוע חוזרת רק כשהוא נרגע (סעיף 4.4 בתוכנית), ורק
      // אחריה הסמן משקף את העריכה שהאיפוס היה חלק ממנה.
      if ((await readPrefix(doc, first, 0)) === null) return;
      const at = caretNow();
      if (disposed || token !== prefixToken || !at || at.offset > ENGINE_MARKER_MAX) return;
      if (at.offset === 0) {
        resetPrefix = { blockId: at.blockId, text: '' };
        return;
      }
      const text = await readPrefix(doc, at, at.offset);
      if (disposed || token !== prefixToken || text === null) return;
      resetPrefix = { blockId: at.blockId, text };
    });
  }

  async function readPrefix(doc: ListAutoformatDoc, at: Caret, length: number): Promise<string | null> {
    const resolve = doc.ranges?.resolve;
    if (typeof resolve !== 'function') return null;
    const request: Record<string, unknown> = {
      start: { kind: 'point', point: point(at.blockId, 0, at.story) },
      end: { kind: 'point', point: point(at.blockId, length, at.story) },
    };
    if (at.story) request.in = at.story;
    try {
      const resolved = await resolve(request);
      const text = resolved?.preview?.text;
      if (typeof text !== 'string' || resolved?.preview?.truncated === true) return null;
      if ((resolved?.target?.start?.offset ?? 0) !== 0) return null;
      return text;
    } catch {
      return null;
    }
  }

  /**
   * ההוכחה, רגע לפני המחיקה, שהיא עדיין חותכת את מה שהוקלד. `''` — מחזיקה.
   *
   * ‏`apply` ממתינה ל-`history.get` לפני שהיא מוחקת, והחלון הזה פתוח לקלט:
   * לחיצה, Backspace או Ctrl+Z מאפסים את המצב אבל אינם יכולים לבטל קריאה
   * שכבר רצה. המחיקה חותכת **היסטים קבועים** (0 עד אורך הסימן) בבלוק ידוע,
   * ולכן מה שהיה שם כשהוחלט אינו בהכרח מה שיש שם כשהיא מגיעה.
   *
   * - **האסימון** סופר כל איפוס, והוא זה שתופס „הסמן חזר לאותו מקום והתוכן
   *   השתנה”: Backspace ואחריו תו אחר משאירים את ההיסט כפי שהיה.
   * - **הסמן** תופס את מה שאינו עובר באיפוס. הקלדה רגילה אינה מאפסת — ובה
   *   ההיסט רק גדל ותחילת הפסקה אינה זזה, ולכן החסם הוא „לא אחורה” ולא
   *   שוויון: שוויון היה מבטל בדיוק את ההמרה שבאמצע הקלדה רציפה.
   * - **אחרי Enter** הסמן כבר בבלוק החדש ואינו מעיד על הישן; שם האסימון הוא
   *   כל ההוכחה, והוא נלקח אחרי האיפוס של אותו Enter.
   */
  function markerHolds(at: Caret, token: number, afterEnter: boolean): string {
    if (disposed) return 'disposed';
    if (token !== resetToken) return 'reset';
    if (afterEnter) return '';
    const still = caretNow();
    if (!still) return 'stale';
    if (still.blockId !== at.blockId) return 'block';
    if (still.offset < at.offset) return 'back';
    return '';
  }

  async function apply(p: Pending, at: Caret, how: string, afterEnter = false): Promise<void> {
    const doc = docOf(host);
    const insert = doc?.insert;
    const create = doc?.lists?.create;
    if (disposed || applying || typeof insert !== 'function' || typeof create !== 'function') return;

    const token = resetToken;
    applying = true;
    typedWhileApplying = false;
    dropRedoGroup();
    // לפני המוטציה: Enter שנלחץ עכשיו יוצר פריט רשימה, לא פסקה רגילה.
    kinds.set(at.blockId, 'other');
    // המחיקה מזיזה את ההיסט בין שתי הקשות — לא אי-התאמה של המנוע.
    lastTyped = null;
    try {
      const before = await historyDepth(doc!);
      const moved = markerHolds(at, token, afterEnter);
      if (moved) {
        kinds.delete(at.blockId);
        note(`cancelled:${moved}`);
        return;
      }
      const erased = await call(() =>
        insert({ value: '', type: 'text', target: textTarget(at, 0, p.plan.markerLength) }),
      );
      if (!erased) {
        kinds.delete(at.blockId);
        note('erase-failed');
        return;
      }

      const created = await call(() =>
        create({
          mode: 'fromParagraphs',
          target: { kind: 'block', nodeType: 'paragraph', nodeId: at.blockId },
          kind: p.plan.kind === 'bullet' ? 'bullet' : 'ordered',
          style: { version: 1, levels: levelsFor(p.plan) },
          sequence: { mode: 'new', startAt: p.plan.kind === 'numbered' ? p.plan.startAt : 1 },
        }),
      );

      if (!created) {
        await call(() => insert({ value: p.typed, type: 'text', target: textTarget(at, 0, 0) }));
        kinds.delete(at.blockId);
        note('create-failed');
        return;
      }

      if (debug) debug.applies += 1;
      note(`applied:${p.plan.kind}:${how}`);

      const after = await historyDepth(doc!);
      if (before && after && after.undo > before.undo && !typedWhileApplying) {
        historyState.undoGroup = {
          steps: after.undo - before.undo,
          depth: after.undo,
          at: { ...at },
          current: { kind: 'other', text: '' },
          other: { kind: 'paragraph', text: p.typed },
        };
      }
    } finally {
      applying = false;
      lastTyped = null;
      invalidateDocumentReads(doc);
    }
  }

  /**
   * הנתיב שאינו סומך על העוגן: סוג הבלוק והטקסט שבתחילתו, מהמסמך. הקריאות
   * עונות רק כשהמנוע שקט, ואז היסט הסמן שווה בדיוק לכל מה שהוקלד מאז האיפוס
   * רק אם הרצף התחיל בהיסט 0.
   */
  async function verify(p: Pending): Promise<void> {
    const doc = docOf(host);
    if (!doc) return;
    for (let attempt = 0; attempt < VERIFY_ATTEMPTS; attempt += 1) {
      const known = p.anchor && !('afterEnter' in p.anchor) ? p.anchor : null;
      const guess = known ?? (await caretSoon());
      if (disposed || pending !== p) return;
      if (!guess) break;
      // תמיד קריאה טרייה: המנוע שקט ממילא, וזיכרון ישן כאן היה מחמיץ המרה.
      const kind = await readKind(doc, guess);
      const text = await readPrefix(doc, guess, p.typed.length);
      if (disposed || pending !== p) return;

      const at = await caretSoon();
      if (disposed || pending !== p) return;
      if (!at || at.blockId !== guess.blockId || at.offset !== p.typed.length + p.keysAfter) continue;
      pending = null;
      if (kind !== 'paragraph') {
        note(`kind:${kind ?? 'unknown'}`);
        return;
      }
      if (text !== p.typed) {
        note(`changed:${JSON.stringify(text)}`);
        return;
      }
      await apply(p, at, 'verified');
      return;
    }
    if (pending === p) pending = null;
    note('verify:gave-up');
  }

  function watch(p: Pending): void {
    if (watchTimer !== undefined) clearTimeout(watchTimer);
    const tick = (): void => {
      watchTimer = undefined;
      if (disposed || pending !== p) return;
      const at = anchoredAt(p.anchor, caretNow());
      if (at && at.offset === p.typed.length + p.keysAfter) {
        inherit(at);
        if (kinds.get(at.blockId) === 'paragraph') {
          pending = null;
          void apply(p, at, 'typed');
        } else {
          note(`kind-unknown:${kinds.get(at.blockId) ?? (reading.has(at.blockId) ? 'reading' : 'none')}`);
          void verify(p);
        }
        return;
      }
      if (now() - p.since > WATCH_LIMIT_MS) {
        note('watch-timeout');
        void verify(p);
        return;
      }
      watchTimer = setTimeout(tick, WATCH_INTERVAL_MS);
    };
    watchTimer = setTimeout(tick, 0);
  }

  /**
   * העוגן של רצף חדש, מהמצב שנראה בתו הראשון שלו (לפני שהמנוע ראה אותו).
   *
   * המצב הזה מוכיח משהו רק אם מקש האיפוס כבר נקלט: המנוע היה מדויק כשהאיפוס
   * נלחץ (`exact`), היה איפוס אחד בלבד, והמצב השתנה מאז. Enter יוצר בלוק שבו
   * הסמן בהיסט 0 — ולכן עוגן אחריו אינו תלוי בשאלה אם כבר נקלט.
   */
  function anchorFor(at: Caret | null, time: number): Anchor | null {
    const mark = resetMark;
    if (mark.kind !== 'install') {
      if (mark.count !== 1 || !mark.exact) return null;
      if (mark.enter && mark.seen) {
        const moved = at !== null && at.blockId !== mark.seen.blockId;
        return moved && at.offset !== 0 ? null : { afterEnter: mark.seen.blockId };
      }
    }
    if (!at || at.offset !== 0 || time < cautiousUntil) return null;
    if (mark.kind === 'install') return at;
    const moved = mark.seen !== null && !sameCaret(at, mark.seen);
    return moved || (mark.kind === 'pointer' && time - mark.at >= POINTER_SETTLE_MS) ? at : null;
  }

  /** בלוק ש-Enter יצר מפסקה רגילה הוא פסקה רגילה. */
  function inherit(at: Caret | null): void {
    if (!at || inheritFrom === null || at.blockId === inheritFrom) return;
    if (!kinds.has(at.blockId)) kinds.set(at.blockId, 'paragraph');
    inheritFrom = null;
  }

  /**
   * שני מקורות לסימן, לפי הסדר:
   *
   * 1. **רצף ההקלדה.** רק הוא מאפשר את הנתיב המהיר — רק הוא מעיד שהכול הוקלד
   *    עכשיו, ורק לו יש עוגן.
   * 2. **מה שלפני הסמן** (`markerBeforeCaret`), כשלרצף אין תוכנית. הרצף נקטע
   *    (Backspace, חץ, לחיצה) אבל מה שנכתב הוא סימן: „1x”, Backspace, „. ”
   *    הוא „1. ” לכל דבר. ההחלה שלו עוברת **תמיד** בנתיב המאומת — הרצף אינו
   *    מכסה את הסימן כולו, ולכן העוגן שלו אינו מעיד עליו — והנתיב הזה קורא
   *    את הפסקה שוב ומשווה, כך שקריאה ישנה אינה מוחקת דבר.
   */
  function onMarker(typedRun: string | null, anchored: Anchor | null, trigger: ' ' | '\t', at: Caret | null): void {
    if (!enabled || applying || pending) return;
    if (debug) debug.evaluates += 1;

    const fromRun = typedRun ? `${typedRun}${trigger}` : '';
    let typed = fromRun;
    let anchor2 = anchored;
    let plan = fromRun ? planListAutoformat(fromRun) : null;
    if (!plan) {
      const before = markerBeforeCaret(at);
      const fromPrefix = before ? `${before}${trigger}` : '';
      if (fromPrefix && fromPrefix !== fromRun) {
        plan = planListAutoformat(fromPrefix);
        if (plan) {
          typed = fromPrefix;
          anchor2 = null;
        }
      }
    }
    if (!plan) {
      note(`no-plan:${JSON.stringify(typed)}`);
      return;
    }
    const p: Pending = {
      typed,
      plan,
      anchor: anchor2,
      keysAfter: 0,
      since: now(),
      verifyOnly: trigger === '\t',
    };
    pending = p;
    if (p.verifyOnly || !p.anchor) {
      note(p.verifyOnly ? 'tab' : `no-anchor:${anchorMiss}`);
      void verify(p);
    } else {
      watch(p);
    }
  }

  function onPrintable(key: string, at: Caret | null, time: number): void {
    if (applying) typedWhileApplying = true;
    dropRedoGroup();
    if (lastTyped && (!at || at.blockId !== lastTyped.blockId || at.offset !== lastTyped.offset + 1)) {
      cautiousUntil = time + CAUTIOUS_MS;
    }
    lastTyped = at;
    if (pending) pending.keysAfter += 1;

    if (run === '') {
      anchor = anchorFor(at, time);
      if (!anchor) {
        const m = resetMark;
        anchorMiss = `${m.kind}/${m.count}/${m.exact ? 'exact' : 'inexact'}/${at ? at.offset : 'stale'}`;
      }
    } else if (
      anchor && run !== null && !('afterEnter' in anchor) &&
      (!at || at.blockId !== anchor.blockId || at.offset !== run.length)
    ) {
      anchor = null;
      anchorMiss = `lag/${at ? at.offset : 'stale'}/${run.length}`;
    }
    inherit(at);

    if (key === ' ') {
      const typedRun = run;
      const anchored = anchor;
      run = null;
      anchor = null;
      onMarker(typedRun, anchored, ' ', at);
      return;
    }
    if (run !== null) run = run.length < LONGEST_MARKER ? run + key : null;
    if (at) learnKind(at);
  }

  function onReset(kind: ResetMark['kind'], at: Caret | null, time: number, isEnter: boolean): void {
    const exact = lastTyped
      ? !!at && at.blockId === lastTyped.blockId && at.offset === lastTyped.offset + 1
      : !!at && time - lastInputAt >= IDLE_MS;
    if (lastTyped && !exact) cautiousUntil = time + CAUTIOUS_MS;
    const chained = !lastTyped && resetMark.kind !== 'install' && time - lastInputAt < IDLE_MS;
    resetMark = chained
      ? { ...resetMark, kind, count: resetMark.count + 1, at: time, enter: isEnter }
      : { kind, seen: at, exact, count: 1, at: time, enter: isEnter };
    lastInputAt = time;

    const waiting = pending;
    pending = null;
    // המרה שכבר החלה אינה נעצרת מעצמה: `pending` כבר אופס כשהיא יצאה לדרך.
    // האסימון הוא מה שעוצר אותה, רגע לפני המחיקה (`markerHolds`).
    resetToken += 1;
    startRun();
    learnPrefix();

    if (!isEnter) return;
    // Enter בפריט רשימה ריק מוציא אותו מהרשימה. מפסקה רגילה — אינו משנה דבר.
    if (!at || kinds.get(at.blockId) !== 'paragraph') forgetKinds();
    // Enter אחרי הרווח ולפני ההמרה: אם הכול כבר נקלט, הבלוק ידוע.
    if (waiting) {
      const ready = anchoredAt(waiting.anchor, at);
      if (
        ready && exact && !waiting.verifyOnly &&
        ready.offset === waiting.typed.length + waiting.keysAfter &&
        kinds.get(ready.blockId) === 'paragraph'
      ) {
        // ‏Enter מעביר את הסמן לבלוק החדש; העדות על הישן היא האסימון בלבד.
        void apply(waiting, ready, 'typed', true);
      }
      return;
    }
    if (at && exact && !applying && kinds.get(at.blockId) === 'paragraph') inheritFrom = at.blockId;
  }

  const onKeyDown = (event: KeyboardEvent): void => {
    if (disposed) return;
    const { key } = event;
    if (MODIFIER_KEYS.has(key)) return;
    const inside = container.contains(event.target as Node | null);
    const history = historyKey(event);

    if (history && inside && takeHistory(history)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      onReset('key', caretNow(), now(), false);
      return;
    }

    // מקש ברצועה, בתפריט או בדיאלוג עשוי להפעיל פקודת רשימות.
    if (!inside) {
      onEdit();
      return;
    }
    const time = now();
    const at = caretNow();
    // ‏`keyCode === 229` הוא החצי השני של התקן בריפו (`src/ui/shortcuts/match.ts`):
    // דפדפן שאינו מציב `isComposing` מדווח אותו. בלעדיו רווח שנלחץ בתוך הרכבה
    // היה נעצר ב-`stopImmediatePropagation` ומוחלף בסינתטי — כלומר שבירת ההרכבה.
    const composing = event.isComposing || event.keyCode === 229;
    const plain = !event.ctrlKey && !event.metaKey && !event.altKey && !composing;

    if (plain && [...key].length === 1) {
      const intercept = key === ' ' && interceptsSpace(at);
      onPrintable(key, at, time);
      if (typedSince !== null) typedSince = typedSince.length < ENGINE_MARKER_MAX ? typedSince + key : null;
      lastInputAt = time;
      if (intercept) {
        // המנוע לא יראה את ה-keydown; את ה-beforeinput מחליף `onBeforeInput`.
        event.stopImmediatePropagation();
        spaceArmed = true;
      }
      return;
    }
    const tab = plain && key === 'Tab' && !event.shiftKey;
    const typedRun = run;
    const anchored = anchor;
    // מה שיישאר מהסימן אחרי מחיקה של תו — `null` כשהנמחק אינו מהרצף הזה.
    const shrunk = plain && key === 'Backspace' && typedSince ? typedSince.slice(0, -1) : null;
    onReset('key', at, time, plain && key === 'Enter' && !event.shiftKey);
    if (plain && key === 'Backspace') typedSince = shrunk;
    // ‏Ctrl+Z/Ctrl+Y שהגיעו לכאן הם אלה שהקבוצה שלהם כבר נצרכה, והמקש חזר
    // למנוע. הם אינם עריכה ואינם קוטמים את צד ה„חזור”, ולכן גם אינם זורקים
    // את הקבוצה שבו — ראו „‏Ctrl+Z אינו עריכה” בהערת הפתיחה.
    if (!NAVIGATION_KEYS.has(key) && !history) dropRedoGroup();
    if (!NAVIGATION_KEYS.has(key) && key !== 'Enter') {
      // פעולה שמשנה תוכן יכולה גם להחליף את סוג הבלוק בלי להחליף את מזההו.
      // רענון הבחירה נדחה: מיד אחרי Delete/Backspace המנוע עוד עשוי לדווח
      // את הסמן הישן, ולכן קריאה מיידית הייתה מתבטלת או נלמדת לבלוק שגוי.
      forgetKinds();
      learnWhenSettled();
    }
    if (tab) onMarker(typedRun, anchored, '\t', at);
  };

  const onBeforeInput = (event: Event): void => {
    if (disposed || !spaceArmed || ours.has(event)) return;
    spaceArmed = false;
    const input = event as InputEvent;
    if (input.inputType !== 'insertText' || input.data !== ' ' || !event.target) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    // שני האירועים הסינתטיים נמדדים נקלטים (השער: „הרווח נכנס במקומו”).
    // אם שדרוג מנוע יפסיק לקלוט אותם הרווח ילך לאיבוד בשקט — ואין כאן קוד
    // גיבוי שלא נמדד לו טריגר, אלא רישום בידית האבחון שהשער קורא.
    if (!insertPlainSpace(event.target, ours)) note('space-lost');
  };

  const onKeyUp = (event: KeyboardEvent): void => {
    if (event.key === ' ') spaceArmed = false;
    if (disposed || !container.contains(event.target as Node | null)) return;
    const at = caretNow();
    if (at) learnKind(at);
  };

  /** לחיצה בתוך המסמך אינה משנה רשימות; לחיצה ברצועה או בתפריט — אולי. */
  const onPointerDown = (event: PointerEvent): void => {
    onReset('pointer', caretNow(), now(), false);
    if (!container.contains(event.target as Node | null)) {
      forgetKinds();
      // פקודת סרגל יכולה לשנות רשימה בלי לשנות את מזהה הפסקה. מחכים לבחירה
      // היציבה כדי שההקלדה הבאה לא תיפול לנתיב המאומת רק בגלל מטמון ישן.
      learnWhenSettled();
    }
  };

  const onEdit = (): void => {
    onReset('key', null, now(), false);
    dropRedoGroup();
    forgetKinds();
  };

  document.addEventListener('keydown', onKeyDown, true);
  document.addEventListener('beforeinput', onBeforeInput, true);
  document.addEventListener('keyup', onKeyUp, true);
  document.addEventListener('pointerdown', onPointerDown, true);
  container.addEventListener('pointerup', learnWhenSettled);
  container.addEventListener('focusin', learnWhenSettled);
  for (const type of EDIT_EVENTS) container.addEventListener(type, onEdit, true);
  learnWhenSettled();

  return {
    undo: () => !disposed && takeHistory('undo'),
    redo: () => !disposed && takeHistory('redo'),
    setEnabled(value: boolean) {
      if (disposed || value === enabled) return;
      enabled = value;
      // הרצף שבאמצע נכתב תחת מצב אחר; קבוצת הביטול נשמרת.
      pending = null;
      resetToken += 1;
      if (watchTimer !== undefined) clearTimeout(watchTimer);
      watchTimer = undefined;
      startRun();
      forgetKinds();
      learnPrefix();
    },
    dispose() {
      disposed = true;
      if (watchTimer !== undefined) clearTimeout(watchTimer);
      for (const timer of timers) clearTimeout(timer);
      timers.clear();
      document.removeEventListener('keydown', onKeyDown, true);
      document.removeEventListener('beforeinput', onBeforeInput, true);
      document.removeEventListener('keyup', onKeyUp, true);
      document.removeEventListener('pointerdown', onPointerDown, true);
      container.removeEventListener('pointerup', learnWhenSettled);
      container.removeEventListener('focusin', learnWhenSettled);
      for (const type of EDIT_EVENTS) container.removeEventListener(type, onEdit, true);
    },
  };
}
