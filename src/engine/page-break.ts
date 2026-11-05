/**
 * „מעבר עמוד” בלשונית „הוספה”, דרך `doc.format.paragraph.setFlowOptions`.
 *
 * ## המצב הקודם
 *
 * `insertPageBreak()` הייתה **פונקציה ריקה**, והכפתור הציג `Ctrl+Enter` כקיצור
 * שלו. שני הדברים לא היו נכונים: הפונקציה לא עשתה כלום, והתוסף אינו רושם את
 * הקיצור בשום מקום.
 *
 * ## למה `pageBreakBefore` ולא מעבר עמוד אמיתי
 *
 * מה ש-Word קורא לו Insert ▸ Page Break הוא `<w:br w:type="page"/>` בתוך ריצה —
 * מעבר **בסמן**, שמפצל את הפסקה הנוכחית. אין לזה API ציבורי ב-2.8.0:
 *
 * - אין פקודה כזאת ב-`COMMAND_CATALOG` של ה-controller (נבדקו כל ה-descriptors;
 *   `create`, `link`, `inline`, `blockParagraph`, `list`, `table` — ואין break).
 * - אין פעולה כזאת ב-`OperationId` של ה-Document API.
 * - במנוע עצמו **כן** קיים `insertPageBreakAtSelection`, קשור ל-`Mod-Enter`
 *   בצינור הקלט הפנימי שלו. הוא אינו נחשף לא על `doc` ולא על `ui.commands`,
 *   ולכן אין ממנו מסלול לפקד ברצועה. (מכאן גם ההערה שהייתה בקוד על
 *   `Ctrl+Enter` — כנראה הקיצור אכן עובד בתוך העורך, אבל זה **לא אומת**,
 *   ולכן הוא אינו מוצג כקיצור של הכפתור. §12: לא מממשים דרך DOM פנימי.)
 *
 * שתי חלופות ציבוריות נשקלו:
 *
 * 1. **`create.sectionBreak({ breakType: 'nextPage' })`** — נדחתה. היא אינה
 *    מעבר עמוד אלא מעבר **מקטע**: החוזה מתעד שהיא פולטת פסקת-נשא
 *    `<w:p><w:pPr><w:sectPr/></w:pPr></w:p>` „just before the existing final
 *    body section”, כלומר בסוף הגוף ולא בסמן. מקטע חדש גם מנתק כותרות
 *    עליונות/תחתונות והגדרות עמוד, וזה שינוי מבני שהמשתמש לא ביקש.
 * 2. **`format.paragraph.setFlowOptions({ pageBreakBefore: true })`** —
 *    נבחרה. זה `w:pageBreakBefore`, בדיוק התכונה ש-Word מציג כ„מעבר עמוד
 *    לפני” בדיאלוג הפסקה. OOXML אמיתי שעושה round-trip, בלי שינוי מבנה.
 *
 * ההבדל הסמנטי אינו מוסתר: `pageBreakBefore` מזיז את **כל הפסקה** לעמוד הבא,
 * ולא מפצל אותה בסמן. לכן התווית היא „התחל בעמוד חדש” ולא „מעבר עמוד”, וה-
 * tooltip אומר בדיוק מה יקרה. תווית שמבטיחה את התנהגות Word ומבצעת אחרת גרועה
 * מתווית מדויקת.
 *
 * ## NO_OP אינה שגיאה
 *
 * `possibleFailureCodes` של הפעולה הוא `['NO_OP']` בלבד, והמנוע מחזיר אותו כש-
 * ה-XML לא השתנה — כלומר הפסקה כבר מתחילה בעמוד חדש. מבחינת המשתמש זו הצלחה,
 * ולחיצה שנייה אינה אמורה להראות שגיאה. אותה הכרעה כמו ב-page-setup.ts.
 *
 * ## „אינו מתג” (docs/button-audit.md, שורה ד') — ומה שהמדידה גילתה
 *
 * הביקורת תיעדה שהכפתור מחיל את `pageBreakBefore` אך אינו מראה שהפסקה כבר
 * מסומנת ואין דרך לבטל מהרצועה, עם ההמלצה „לקרוא את המצב ולשלוח `false`
 * בכיבוי”. כתיבת `false` היא שורה אחת (`setFlowOptions({target,
 * pageBreakBefore:false})`, נבחן ועובד). **הקריאה היא המכשול האמיתי**, ונמדד
 * ישירות (Chrome headless, ה-dist הארוז): אחרי `setFlowOptions({pageBreakBefore:
 * true})` מוצלח (`success:true`, וה-DOCX המיוצא מכיל בפועל `<w:pageBreakBefore/>`),
 * תכונות אותה פסקה בדיוק שחוזרות מ-`doc.get()` הן `{bidi:true}` בלבד — בלי
 * `pageBreakBefore` כלל. זה בדיוק אותו פער שכבר מתועד ל-`tabs`/
 * `keepWithNext`/`keepLines`/`widowControl` ב-docs/engine-gaps.md (ראו
 * paragraph-format.ts), והמדידה הזאת מוסיפה אליו את `pageBreakBefore`. אין
 * גם API ייעודי לקרוא אותו (בניגוד ל-`titlePg`/`evenAndOddHeaders` שנקראים
 * מ-`sections.list()`, header-footer.ts) — ואין לקרוא `w:pPr` גולמי דרך DOM
 * או מנוע פנימי (tests/unit/engine-boundaries.test.ts).
 *
 * **המשמעות:** אי אפשר לבנות `:active` שמשקף את **המסמך**, כמו „שונה בעמוד
 * ראשון”. מה שכן אפשר לבנות הוא מתג אמיתי מבחינת ה**רצועה**: `PageBreakTracker`
 * למטה זוכר, פר-פסקה (`nodeId`), מה נכתב עליה מהרצועה בסשן הזה, ומשתמש בזה
 * כדי להחליט אם הלחיצה הבאה שולחת `true` או `false`. מסמך שנפתח עם
 * `pageBreakBefore` שכבר קיים בו מ-Word יוצג כ„כבוי” עד שהמשתמש ילחץ (הלחיצה
 * הראשונה תהיה NO_OP מבחינת המנוע ותסמן את המצב כ„פעיל” מכאן ואילך) — פער
 * קטן מול „המצב האמיתי מהמסמך”, אבל עדיף בהרבה על מתג שאינו יודע לכבות בכלל.
 *
 * ## QA עצמאי על ה-`PageBreakTracker`: שני פערי תצוגה, ומה נעשה בהם
 *
 * QA שני תקף את ההכרעה שמעל ומצא שני פערים אמיתיים במעקב המקומי — שניהם
 * תצוגתיים-בלבד (הלחיצה הבאה תמיד מיישרת לאמת, ראו „NO_OP אינה שגיאה” למעלה
 * ול-`setParagraphPageBreak`), אבל שניהם יכולים להראות למשתמש „פעיל” על
 * פסקה שה-DOCX שלה כבר לא מסומנת.
 *
 * ### „מסמך אחר” — הטענה שנבדקה ולא שוחזרה, וההחלטה שבכל זאת התקבלה
 *
 * ה-QA טען ש-`activeSuperdoc` הוא **אותו** מופע `SuperDoc` גם אחרי „מסמך
 * חדש”/„פתח קובץ”, ולכן `syncDocument(host)` (השוואת זהות `host`) כמעט
 * אף פעם לא מזהה מסמך אחר. **נמדד מחדש ולא שוחזר**: תיוג
 * `window.__otzariaEditor.superdoc.__marker` לפני „מסמך חדש” ובדיקה אחריו —
 * הסימון לא שרד, כלומר `createEditor` (`new SuperDoc(...)`) **כן** בונה מופע
 * חדש בכל פתיחה מוצלחת, ו-„מסמך חדש” ו„פתח קובץ” עוברים באותו קוד בדיוק
 * (`openDocument` → `swap.open` → `createEditor`, App.vue) — אין מסלול
 * שמתנהג אחרת בין השניים.
 *
 * זהות `host` הישנה, אם כן, **הייתה** מזהה מסמך אחר נכון בפועל. ההחלפה
 * ל-`DOCUMENT_GENERATION` (composables/keys.ts, נגזר מ-`EditorSwap.
 * documentGeneration`, sessions/editor-swap.ts) בוצעה בכל זאת: היא מונה
 * מפורש שאינו נשען על התנהגות לא-מתועדת של מחלקת `SuperDoc` (זהות אובייקט
 * אחרי פתיחה נמדדה, אך אינה הבטחה בחוזה שלה), ומקורה במקום אחד ומוסמך
 * (`editor-swap.ts`) שכבר מנהל בדיוק את השאלה „האם המסמך הפעיל הוחלף”. ראו
 * `PageBreakTracker.syncDocument`.
 *
 * ### Undo/Redo — פער אמיתי, ושלושה סיבובי תיקון
 *
 * זה כן אושר: לחיצה מסמנת פסקה (נכתב `pageBreakBefore`, הכפתור „פעיל”),
 * Ctrl+Z מסיר אותו מה-XML בפועל — אבל `PageBreakTracker` לא שומע Undo/Redo
 * בכלל (הם לא עוברים דרך `setParagraphPageBreak`), ולכן הכפתור נשאר מציג
 * „פעיל” על פסקה שכבר אינה מסומנת, עד ללחיצה הבאה (שאז NO_OP מיישר).
 *
 * **אין איתות ציבורי** ל„הפעולה שרצה הייתה Undo/Redo, ואלה ה-nodeId
 * שהושפעו” — `doc.get()` וה-`onEditorUpdate` הכללי (App.vue) מדווחים רק
 * „המסמך השתנה”, בלי לפרש מה קרה או איפה. הפתרון: ניקוי **כל** `PageBreakTracker`
 * על Undo. תצוגת „לא ידוע” (נופלת ל„כבוי”) עדיפה על תצוגה כוזבת של „פעיל” —
 * בדיוק כמו שהמתג כולו כבר מתפשר על „כבוי” עד לחיצה ראשונה במסמך שנפתח עם
 * `pageBreakBefore` קיים. **המחיר:** Undo על פעולה שאינה נוגעת ב-`pageBreakBefore`
 * בכלל גם הוא מוחק את הידע על כל שאר הפסקאות במסמך — פשרה מכוונת, כי אין
 * דרך זולה יותר לדעת שהוא לא נגע.
 *
 * **סיבוב 1 — התיקון הראשון לא הספיק (נמדד ב-QA בדפדפן אמיתי):** ניקוי
 * מ-`runShortcutCommand` (App.vue) על `undo`/`redo` מוצלח לא רץ בכלל כש-
 * Ctrl+Z מגיע עם הפוקוס בתוך המסמך — `createShortcutDispatcher` מדלג
 * בכוונה על אירוע `defaultPrevented`, וה-`history` המובנה של ProseMirror
 * כבר טיפל בלחיצה וביטל אותה **לפני** שהיא מגיעה לשם. הפתרון:
 * `watchUndoRedoKeys` (ui/shortcuts/undo-redo-watch.ts) — מאזין נפרד
 * ב-**capture**, שרואה את הלחיצה לפני המנוע.
 *
 * **סיבוב 2 — תקלה נוספת שנחשפה תוך כדי אימות סיבוב 1, ותוקנה:** ניקוי
 * `known` בלבד לא הספיק. `Ctrl+Z` הסיר את `w:pageBreakBefore` וה-forget רץ
 * (נמדד), אבל הכפתור המשיך להציג „פעיל” — כי הניקוי קורה ב-App.vue, מחוץ
 * ל-`InsertTab.vue` שמחזיק את `pageBreakOn` המוצג, ו-Undo אינו מזיז את
 * הסמן (ולכן `ui.selection.observe` לא יורה ולא מרענן). הפתרון:
 * `PageBreakTracker.onChange` — מנוי ש-`InsertTab.vue` נרשם אליו, ורץ מחדש
 * `refreshPageBreakOn` בכל שינוי בידע, לא רק בתזוזת סמן.
 *
 * **סיבוב 3 — QA שלישי, שני פערים נוספים ותיקוניהם:**
 *
 * 1. **Redo לא החזיר את הסימון (א-סימטריה).** `watchUndoRedoKeys` קרא
 *    `forgetAll()` גם על Redo, אבל שום דבר לא קרא `remember(nodeId, true)`
 *    בחזרה — כלומר Redo (פעולה שהמשתמש **ביקש**) הציג כפתור „לא פעיל” על
 *    פסקה שה-DOCX שלה כן חזר להיות מסומן, מיד אחרי הפעולה. נמדד 3/3 ריצות
 *    עקביות. **אין דרך לדעת בוודאות** מה בדיוק Redo החזיר בלי לקרוא את
 *    המסמך (אותה מגבלת `doc.get()` המקורית) — אבל אפשר לפחות לא לנקות
 *    בעיוורון: `forgetAllKeepingSnapshot`/`restoreSnapshot` שומרים תצלום
 *    יחיד (לא מחסנית) של הידע *לפני* ה-Undo, ומחזירים אותו ב-Redo המיידי
 *    שאחריו. זה נכון **בדיוק** לתרחיש הנפוץ ביותר — Undo בטעות ו-Redo מיד
 *    בחזרה — ומתפרק בעדינות (`restoreSnapshot` מחזירה `false`, ואז
 *    `forgetAll` רגילה) על שרשראות ארוכות יותר: שני Undo רצופים דורסים את
 *    התצלום של הראשון בתצלום ריק (כי אחרי Undo ראשון `known` כבר ריק), ולכן
 *    Redo כפול לא ישחזר את מה שהיה **לפני** ה-Undo הראשון. פשרה מכוונת:
 *    תצוגה שגויה-כלפי-מטה (Redo שלא משחזר) עדיפה על שגויה-כלפי-מעלה
 *    (שחזור שגוי שמראה „פעיל” כשלא צריך) — ולכן `restoreSnapshot` לעולם לא
 *    מנחשת, רק צורכת תצלום ודאי או נכשלת בבטחה. **סיכון שיורי מתועד ולא
 *    נפתר:** אם המשתמש עורך משהו אחר בין ה-Undo לבין ה-Redo (מה שבעורך רגיל
 *    היה מבטל את אפשרות ה-Redo האמיתית), ולוחץ Ctrl+Y בכל זאת, אין לנו
 *    איתות "עוד יש Redo אמיתי לעשות" — `restoreSnapshot` עדיין תחזיר את
 *    התצלום הישן, אף שהמנוע עצמו לא באמת ביצע Redo. זה שחזור **שגוי**,
 *    אך תרחיש נדיר (המשתמש צריך ללחוץ Redo בלי סיבה, ממש אחרי עריכה חדשה),
 *    ולא נמצאה דרך זולה לחסום אותו בלי לעקוב אחרי כל מוטציה במסמך (`onUpdate`
 *    הכללי אינו מבחין בין Undo/Redo לעריכה רגילה, ומעקב מדויק יותר דורש
 *    מחסנית undo/redo משלנו שמסונכרנת עם זו של המנוע — לא קיימת גישה
 *    ציבורית לזה).
 * 2. **Ctrl+Z בתוך שדה טקסט לא-קשור ניקה בטעות.** `watchUndoRedoKeys`
 *    תפס כל `keydown` תואם ב-`window`, בלי לבדוק את `event.target` —
 *    כלומר Ctrl+Z בתוך `#fr-search-input` (חיפוש-והחלפה) או כל שדה טקסט
 *    אחר של הממשק ניקה את `PageBreakTracker` בלי שום קשר למסמך. נמדד
 *    פעמיים בעקביות. הפתרון: `watchUndoRedoKeys` מקבל `isBlocked` — אותה
 *    בדיקה בדיוק ש-`createDirectionShortcut` כבר משתמש בה ב-App.vue
 *    (`isModalDialogOpen() || (isTextEntryTarget(target) && !isDocumentSurface(target))`),
 *    ומדלג על ההודעה כשהיא `true`.
 *
 * **גם תוקן באותו סיבוב, לא ממצא QA אלא רגרסיה עצמית שנחשפה בזמן התיקון:**
 * כפתורי „בטל”/„חזור” בפס הכותרת (`onUndo`/`onRedo`, App.vue) עברו
 * מ-`runShortcutCommand` (שהחזיק את הניקוי בעבר) ל-`watchUndoRedoKeys`
 * (שרק מאזין למקלדת) — כלומר **לחיצה על הכפתורים בפס הכותרת הפסיקה לנקות
 * את המעקב בכלל**. תוקן: `onUndo`/`onRedo` קוראים ל-`forgetAllKeepingSnapshot`/
 * `restoreSnapshot` ישירות, כמו `watchUndoRedoKeys` — בלי בדיקת `isBlocked`
 * (לחיצה מפורשת על כפתור תמיד עוסקת במסמך, לא במקום שהפוקוס היה בו קודם).
 *
 * ### מה שלא נבדק (תרחיש 3): גזירה/הדבקה של פסקה מסומנת
 *
 * לא נמדד עד הסוף — QA בדפדפן headless נתקע על מסלול ה-clipboard (מגבלת
 * הסביבה, לא של המוצר). זו אותה קטגוריית פער בדיוק: פעולה שמשנה `pageBreakBefore`
 * (או מעתיקה פסקה עם `nodeId` חדש שנושא את התכונה) בלי לעבור דרך הכפתור.
 * ברירת המחדל הבטוחה (`isOn` מחזירה `false` ל-`nodeId` לא-מוכר) חלה גם כאן:
 * פסקה חדשה שנוצרה מהדבקה ומעולם לא נלחצה עליה תוצג כ„כבוי” — נכון אם
 * ההדבקה לא כללה `pageBreakBefore`, ושגוי (באותו אופן בדיוק כמו Undo) אם
 * כן. לא תוקן: אין לו מסלול תיקון זול יותר מזה של Undo, וההשפעה נשארת
 * תצוגתית-בלבד.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';
import { readDocSelection, type SelectionDocumentApi } from './doc-selection';
import { readDocCapabilities } from './doc-capabilities';

/** מזהה הפעולה בקטלוג של המנוע. מיוצא כדי שהבדיקה תקבע אותו מול החבילה. */
export const PAGE_BREAK_OPERATION = 'format.paragraph.setFlowOptions';

/** הנוסח שהתכנית קובעת ב-§12 לפקד שאין לו API זמין. */
const UNAVAILABLE_TEXT = 'אינו זמין בגרסה זו';

/** ההסבר כשאין בכלל מה לשאול. זהה לנוסח ב-doc-capabilities.ts. */
const LOADING_TEXT = 'המסמך עדיין נטען';

const FAILED_ACTION = 'הגדרת התחלה בעמוד חדש נכשלה';
const FAILED_ACTION_OFF = 'ביטול התחלת העמוד החדש נכשל';

/** `ParagraphTarget` — `nodeType: 'paragraph'` גם לכותרת ולפריט רשימה. ראו למטה. */
interface ParagraphTarget {
  kind: 'block';
  nodeType: 'paragraph';
  nodeId: string;
  story?: unknown;
}

interface FlowOptionsInput {
  target: ParagraphTarget;
  pageBreakBefore: boolean;
}

/** הצורה שנצרכת מ-`doc`. ראו ההסבר ב-document-defaults.ts למה מוגדרת ולא מיובאת. */
export interface PageBreakDocumentApi extends SelectionDocumentApi {
  capabilities?: {
    get?: () => MaybePromise<
      { operations?: Partial<Record<string, { available?: boolean } | undefined>> } | undefined
    >;
  };
  format?: {
    paragraph?: {
      setFlowOptions?: (input: FlowOptionsInput) => MaybePromise<DocReceipt>;
    };
  };
}

export interface PageBreakHost {
  activeEditor?: { doc?: PageBreakDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type PageBreakTarget = SuperDoc | PageBreakHost | null | undefined;

/** האם הפקד זמין, ומה להציג ב-tooltip כשלא. */
export interface PageBreakSupport {
  available: boolean;
  /** הסבר בעברית, מוכן ל-tooltip. מחרוזת ריקה כשזמין. */
  explanation: string;
}

/**
 * קוראת את זמינות הפעולה.
 *
 * **למה כאן ולא ב-doc-capabilities.ts**, שזה מקומה הנכון: המודול ההוא מחזיק
 * מרחב שאלות סגור (`CAPABILITY_SPECS`), ואין בו שאלה על
 * `format.paragraph.setFlowOptions`. הוא היה בבעלות גל אחר בזמן כתיבת הקומיט
 * הזה, ולכן לא נגעתי בו. התוספת שנדרשת שם היא שורה אחת —
 * השאלה `canSetPageBreakBefore` נוספה ל-doc-capabilities.ts, ולכן הפונקציה
 * הזאת רק מתרגמת את הדוח שלו לשני המצבים שה-tooltip צריך.
 *
 * לעולם אינה זורקת: כשל של קריאת יכולות אינו סיבה להפיל את רינדור הרצועה.
 * נכשלת **סגור** — „אולי כן” הוא בדיוק הכפתור המת.
 */
export async function readPageBreakSupport(host: PageBreakTarget): Promise<PageBreakSupport> {
  const doc = (host as PageBreakHost | null | undefined)?.activeEditor?.doc;
  if (!doc) return { available: false, explanation: LOADING_TEXT };

  // נוכחות הפונקציה נבדקת **לפני** היכולות, ולא רק בגללן: מפת ה-`operations`
  // של המנוע נבנית מקטלוג הפעולות, ולכן גרסה שהסירה את המימוש ועוד מכריזה על
  // הפעולה בקטלוג הייתה מחזירה „זמין” לפקד שאין לו למה לקרוא. בדיקה שכיסתה
  // בדיוק את המקרה הזה תפסה את ההסרה שלה.
  if (typeof doc.format?.paragraph?.setFlowOptions !== 'function') {
    return { available: false, explanation: UNAVAILABLE_TEXT };
  }

  const report = await readDocCapabilities(host);

  // אין Document API לשאול — המסמך עדיין נטען. שונה מ„הפעולה אינה נתמכת”,
  // ולכן ההסבר שונה: פקד שנעלם לרגע בזמן פתיחה אינו פקד חסר.
  if (!report.available) return { available: false, explanation: LOADING_TEXT };

  if (report.can('canSetPageBreakBefore')) return { available: true, explanation: '' };

  // ההסבר של הדוח נושא את קוד הסיבה שהמנוע נתן; UNAVAILABLE_TEXT הוא הנוסח
  // כשהוא לא נתן כלום.
  return { available: false, explanation: report.explain('canSetPageBreakBefore') || UNAVAILABLE_TEXT };
}

/**
 * כותבת `pageBreakBefore` (בכיוון כלשהו) על הפסקה שהסמן בה, ומחזירה גם את
 * ה-`nodeId` שנפתר — כדי ש-`PageBreakTracker` יוכל לזכור מה נכתב על מי.
 *
 * הפסקה היא זו שהבחירה **מתחילה** בה. בבחירה שפרושה על כמה פסקאות זו הפסקה
 * הראשונה, וזו גם ההתנהגות המתבקשת: „שהחלק הזה יתחיל בעמוד חדש”.
 *
 * `nodeType: 'paragraph'` נשלח גם כשהסמן בכותרת או בפריט רשימה. זה אינו קיצור
 * דרך: פתרון היעד במנוע נעשה לפי `nodeId` ו-`story` בלבד ואינו מסתכל על
 * `nodeType` (נמדד), וזה בדיוק מה ש-`paragraphTarget` של ה-controller עצמו
 * שולח לכל פקודות `format.paragraph.*`.
 *
 * לעולם אינה זורקת: הוולידטורים של ה-Document API **זורקים** `INVALID_INPUT`
 * על קלט פסול במקום להחזיר קבלה, וחריגה מפקד ב-Ribbon מפילה את רינדור הרצועה.
 */
async function applyPageBreak(
  host: PageBreakTarget,
  requested: boolean | ((nodeId: string) => boolean),
  isCurrent: () => boolean = () => true,
): Promise<{ outcome: CommandOutcome; nodeId: string | null; pageBreakBefore: boolean }> {
  let pageBreakBefore = typeof requested === 'boolean' ? requested : true;
  let failedAction = pageBreakBefore ? FAILED_ACTION : FAILED_ACTION_OFF;
  const setFlowOptions = (host as PageBreakHost | null | undefined)?.activeEditor?.doc?.format
    ?.paragraph?.setFlowOptions;

  if (typeof setFlowOptions !== 'function') {
    // אותו נוסח שהיכולת מחזירה, כדי שהמשתמש יראה את אותו הסבר בין אם הפקד
    // מנוטרל ובין אם הוא נלחץ לפני שהיכולות נקראו.
    return {
      outcome: {
        ok: false,
        message: `${failedAction}: ${UNAVAILABLE_TEXT}`,
        reason: 'command-unsupported',
      },
      nodeId: null,
      pageBreakBefore,
    };
  }

  const selection = await readDocSelection(host);
  if (!isCurrent()) {
    return {
      outcome: { ok: false, message: 'המסמך התחלף; יש להפעיל את הפעולה שוב', reason: 'not-ready' },
      nodeId: null,
      pageBreakBefore,
    };
  }
  if (!selection.blockId) {
    return {
      outcome: { ok: false, message: 'יש למקם את הסמן במסמך', reason: 'selection-required' },
      nodeId: null,
      pageBreakBefore,
    };
  }

  // A toggle derives its value from the same resolved paragraph that is written,
  // never from a paused visual readout or a second selection request.
  if (typeof requested === 'function') pageBreakBefore = requested(selection.blockId);
  failedAction = pageBreakBefore ? FAILED_ACTION : FAILED_ACTION_OFF;

  const target: ParagraphTarget = {
    kind: 'block',
    nodeType: 'paragraph',
    nodeId: selection.blockId,
    // נשלח רק כשיש: בהיעדרו היעד מתפרש כגוף המסמך, וזו ברירת המחדל הנכונה.
    // `story: null` מפורש היה נכשל בוולידציה.
    ...(selection.story ? { story: selection.story } : {}),
  };

  let receipt: DocReceipt;
  try {
    receipt = await setFlowOptions({ target, pageBreakBefore });
  } catch (error) {
    return {
      outcome: { ok: false, message: thrownText(failedAction, error), reason: 'threw' },
      nodeId: selection.blockId,
      pageBreakBefore,
    };
  }

  // הערך המבוקש כבר מוגדר. הצלחה מבחינת המשתמש.
  if (receipt?.success === false && receipt.failure?.code === 'NO_OP') {
    return { outcome: { ok: true }, nodeId: selection.blockId, pageBreakBefore };
  }

  if (receipt?.success === false) {
    return {
      outcome: {
        ok: false,
        message: receiptFailureText(failedAction, receipt),
        reason: receipt.failure?.code,
      },
      nodeId: selection.blockId,
      pageBreakBefore,
    };
  }

  return { outcome: { ok: true }, nodeId: selection.blockId, pageBreakBefore };
}

/**
 * מסמנת את הפסקה שבה הסמן כ„מתחילה בעמוד חדש”. שולחת `true` תמיד — זו
 * הפעולה שהקיצור `Ctrl+Enter` מריץ (App.vue, `runPageBreak`), ומשמעותה שם
 * זהה למשמעות של „מעבר עמוד” ב-Word: פעולה קדימה, לא מתג. הכפתור ברצועה
 * (InsertTab.vue) הוא זה שמתנהג כמתג, דרך `toggleParagraphPageBreak` למטה.
 */
export async function startParagraphOnNewPage(host: PageBreakTarget): Promise<CommandOutcome> {
  return (await applyPageBreak(host, true)).outcome;
}

/**
 * גרסת המתג: כותבת את הערך המבוקש (`true`/`false`) ומעדכנת את `tracker`
 * בהצלחה — כולל NO_OP, שהוא הצלחה. ראו „„אינו מתג”” בהערת הפתיחה של הקובץ
 * למה המעקב הזה מקומי-בלבד ואינו נקרא מהמסמך.
 */
export async function setParagraphPageBreak(
  host: PageBreakTarget,
  pageBreakBefore: boolean,
  tracker: PageBreakTracker,
): Promise<CommandOutcome> {
  const version = tracker.documentVersion;
  const isCurrent = () => tracker.documentVersion === version;
  const { outcome, nodeId } = await applyPageBreak(host, pageBreakBefore, isCurrent);
  if (outcome.ok && isCurrent()) tracker.remember(nodeId, pageBreakBefore);
  return outcome;
}

/** Resolve the current paragraph once, then invert its tracked value. */
export async function toggleParagraphPageBreak(
  host: PageBreakTarget,
  tracker: PageBreakTracker,
): Promise<CommandOutcome> {
  const version = tracker.documentVersion;
  const isCurrent = () => tracker.documentVersion === version;
  const { outcome, nodeId, pageBreakBefore } = await applyPageBreak(host, id => !tracker.isOn(id), isCurrent);
  // A write already sent to the old document may finish after a tab switch.
  // Its success must never enter the new document's tracked state.
  if (outcome.ok && isCurrent()) tracker.remember(nodeId, pageBreakBefore);
  return outcome;
}

/**
 * מזהה הפסקה שהסמן בה עכשיו — למעקב המתג (`PageBreakTracker.isOn`). `null`
 * כשאין בחירה. לעולם אינה זורקת: `readDocSelection` עצמה אינה זורקת.
 */
export async function readPageBreakNodeId(host: PageBreakTarget): Promise<string | null> {
  return (await readDocSelection(host)).blockId;
}

/* ------------------------------------------------------------------ */
/* המתג: מעקב מקומי, כי doc.get() אינו מחזיר pageBreakBefore            */
/* ------------------------------------------------------------------ */

/**
 * מה שהכפתור ב-Ribbon צריך כדי להתנהג כמתג — ראו ההסבר המלא ב„„אינו מתג””
 * בהערת הפתיחה של הקובץ.
 */
export interface PageBreakTracker {
  /** Changes only when syncDocument switches document, fencing pending commands. */
  readonly documentVersion: number;
  /** האם ידוע לנו שהפסקה הזאת מסומנת. `false` גם כש„לא ידוע” — זו ברירת המחדל הבטוחה. */
  isOn(nodeId: string | null | undefined): boolean;
  /** קוראת אחרי כתיבה מוצלחת (כולל NO_OP) — ראו `setParagraphPageBreak`. */
  remember(nodeId: string | null | undefined, on: boolean): void;
  /**
   * מאפסת את הידע **רק** כש-`key` שונה מזה שנמסר בקריאה הקודמת — כלומר
   * מסמך אחר נפתח. לא כשזהה: `InsertTab.vue` נהרס ונבנה מחדש בכל מעבר בין
   * לשוניות (`v-if` ב-Ribbon.vue), ו-`watch(superdoc, ..., {immediate:true})`
   * שלו רץ בכל הרכבה כזאת — קריאה שמאפסת תמיד הייתה שוכחת את המתג שסומן
   * ברגע שהמשתמש עוזב את הלשונית וחוזר אליה.
   *
   * **`key` הוא `DOCUMENT_GENERATION` (composables/keys.ts), לא `host`** —
   * ראו הערת ה-QA שהוסיפה את זה, ולמה זה שינוי אמיתי ולא קוסמטי:
   *
   * גרסה קודמת השוותה את `host` (מופע ה-`SuperDoc`) עצמו לפי זהות אובייקט.
   * נמדד ישירות בדפדפן (תיוג `window.__otzariaEditor.superdoc` לפני „מסמך
   * חדש”/„פתח קובץ” ובדיקה אחרי) ש**כן** נוצר מופע `SuperDoc` חדש בכל פתיחה
   * מוצלחת — כלומר ההשוואה לפי זהות `host` בפועל **כן** תפסה מסמך אחר, ולא
   * שוכפלה בשקט כפי שה-QA חשד. עם זאת ההחלפה ל-`DOCUMENT_GENERATION` בוצעה
   * בכל זאת, כי היא איתות מפורש ובלתי-תלוי בפרט מימוש שאינו מתועד בחוזה של
   * `SuperDoc` (זהות אובייקט אחרי פתיחה היא התנהגות שנמדדה, לא הבטחה כתובה) —
   * וכי `sessions/editor-swap.ts` הוא כתובת אחת ומוסמכת ל„האם זה מסמך אחר”,
   * שעדיף על כל צרכן שיישם את השאלה הזאת בעצמו לפי הנחות על גרסת מנוע.
   */
  syncDocument(key: unknown): void;
  /**
   * מוחקת את **כל** הידע, בלי קשר למסמך. לפעולות שמשנות `pageBreakBefore`
   * מחוץ לכפתור הזה ואין דרך לדעת אילו `nodeId` הן נגעו בהן. ראו הסבר
   * המכשול המלא בהערת הפתיחה של הקובץ, תחת „„QA עצמאי”” → „Undo/Redo”.
   *
   * גם מוחקת תצלום ממתין (`forgetAllKeepingSnapshot`) אם היה — קריאה סתמית
   * ל-`forgetAll` היא איפוס מלא ומכוון, לא צריך שהוא ישאיר משהו להחזרה.
   */
  forgetAll(): void;
  /**
   * כמו `forgetAll`, אבל שומרת תצלום של הידע **לפני** המחיקה — כדי ש-Redo
   * מיידי שאחריו (`restoreSnapshot`) יוכל להחזיר אותו. נקראת על Undo, לא על
   * Redo — ראו „א-סימטריית Undo/Redo” בהערת הפתיחה של הקובץ למה זה בכלל
   * נחוץ ומה הגבולות שלו (תצלום יחיד, לא מחסנית).
   */
  forgetAllKeepingSnapshot(): void;
  /**
   * "צורכת" (חד-פעמית) את התצלום האחרון שנשמר על ידי `forgetAllKeepingSnapshot`,
   * אם יש כזה, ומחזירה אותו לתוך הידע הנוכחי — כאילו ה-Undo שיצר אותו לא
   * קרה. מחזירה `true` אם היה תצלום ו-`false` אם לא (ואז לא עשתה כלום —
   * הקורא אמור ליפול חזרה ל-`forgetAll`, ראו App.vue). „חד-פעמית”: אחרי
   * קריאה — מוצלחת או לא — אין תצלום להחזיר יותר עד ה-Undo הבא.
   */
  restoreSnapshot(): boolean;
  /**
   * נרשמת לכל שינוי בידע (`remember`, `syncDocument`/`forgetAll` שבאמת
   * מחקו) ומחזירה פונקציית ביטול-מנוי.
   *
   * למה זה קיים: `forgetAll` נקראת מ-App.vue, מחוץ ל-`InsertTab.vue` שמחזיק
   * את `pageBreakOn` המוצג. בלי מנוי, ה-Undo *כן* מנקה את `known` בפועל —
   * אבל הכפתור עדיין מציג את הערך הישן עד לרענון הבא (תזוזת סמן/לחיצה),
   * שעלול לא לקרות כלל אם Undo אינו מזיז את הסמן. נמדד: ב-QA בדפדפן אמיתי
   * `Ctrl+Z` הסיר את `w:pageBreakBefore` ו-`forgetAll` רץ, אך הכפתור נשאר
   * „פעיל” — בדיוק התקלה הזאת, לפני שהמנוי הזה נוסף.
   */
  onChange(listener: () => void): () => void;
}

export function createPageBreakTracker(): PageBreakTracker {
  const known = new Map<string, boolean>();
  const UNSET = Symbol('page-break-tracker-unset');
  let lastKey: unknown = UNSET;
  let documentVersion = 0;
  const listeners = new Set<() => void>();
  /** ראו `forgetAllKeepingSnapshot`/`restoreSnapshot`. `null` = אין מה להחזיר. */
  let pendingSnapshot: Map<string, boolean> | null = null;

  function notify(): void {
    for (const listener of listeners) {
      try {
        listener();
      } catch (error) {
        console.error('[otzaria-word] מאזין למעקב page-break נכשל', error);
      }
    }
  }

  return {
    get documentVersion() { return documentVersion; },
    isOn: (nodeId) => typeof nodeId === 'string' && nodeId !== '' && known.get(nodeId) === true,
    remember(nodeId, on) {
      if (typeof nodeId !== 'string' || nodeId === '') return;
      known.set(nodeId, on);
      notify();
    },
    syncDocument(key) {
      if (key === lastKey) return;
      lastKey = key;
      documentVersion += 1;
      known.clear();
      // מסמך אחר: תצלום ממתין ממסמך שכבר אינו פתוח אינו רלוונטי להחזרה.
      pendingSnapshot = null;
      notify();
    },
    forgetAll() {
      known.clear();
      pendingSnapshot = null;
      notify();
    },
    forgetAllKeepingSnapshot() {
      pendingSnapshot = new Map(known);
      known.clear();
      notify();
    },
    restoreSnapshot() {
      if (!pendingSnapshot) return false;
      known.clear();
      for (const [nodeId, on] of pendingSnapshot) known.set(nodeId, on);
      pendingSnapshot = null;
      notify();
      return true;
    },
    onChange(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
}

/**
 * מופע יחיד, משותף לכל מי שקורא ל-Ribbon בסשן — ולא נוצר מחדש בכל הרכבה של
 * `InsertTab.vue`. ראו `syncDocument` למה השיתוף הזה בטוח: הוא זה שמבטיח
 * שמסמך אחר מאפס את הידע, ואותו מסמך (רק לשונית שהוחלפה) לא.
 */
export const pageBreakTracker: PageBreakTracker = createPageBreakTracker();
