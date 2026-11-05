/**
 * מי שזוכר. אוסף את מצב ההפעלה, מחליט מתי לכתוב אותו, ומחזיק את טיוטת
 * השחזור.
 *
 * ## שני קצבים, ולא אחד
 *
 * מה שנשמר כאן נחלק לשניים שעלותם שונה בסדרי גודל:
 *
 * - **הרשומה** — מסמך, סמן, גודל תצוגה, לשונית. כתיבה אחת ל-`storage`,
 *   מיקרו-שניות. השהיה קצרה (`PERSIST_DELAY_MS`).
 * - **הטיוטה** — ייצוא DOCX מלא, המרה ל-base64, ומעבר בגשר. עשרות עד מאות
 *   מילישניות על מסמך אמיתי. השהיה ארוכה (`DRAFT_DELAY_MS`).
 *
 * קצב אחד לשניהם היה מכריח לבחור בין „הסמן נשמר לעיתים רחוקות” לבין „ייצוא
 * מלא כל שתי שניות של הקלדה”. שני קצבים הם מה שמאפשר לשניהם להיות נכונים.
 *
 * ## למה הטיוטה קיימת בכלל, כשיש שמירה אוטומטית
 *
 * השמירה האוטומטית כותבת לקובץ של המשתמש, ולכן היא פועלת רק כשיש לאן: מסמך
 * חדש שטרם נשמר אין לו יעד, ומסמך שנפתח לקריאה בלבד — גם כן. גם המתג עצמו
 * עשוי להיות כבוי, וזו בחירה לגיטימית שאין לעקוף. בכל שלושת המצבים אין שום
 * דבר בדיסק, וסגירת אוצריא מחקה את העבודה.
 *
 * הטיוטה היא במקום אחר לגמרי — המרחב הפרטי של התוסף — ולכן היא **אינה**
 * עוקפת את המתג: המשתמש אמר „אל תכתוב לקובץ שלי”, לא „תשכח מה שכתבתי”. זו
 * בדיוק ההפרדה ש-Word עושה בין „שמירה אוטומטית” ל„שחזור אוטומטי”.
 *
 * ## מה מוחק את הטיוטה — ומה **אינו** מוחק
 *
 * מסלול אחד בלבד: **`discardDraft`**, כלומר המשתמש אמר במפורש „למחוק את
 * השינויים”. בקובץ הזה מוחקים רק כשנאמר.
 *
 * **`noteSaved` אינו מוחק — הוא מעביר.** טיוטה שנכללה ב-snapshot השמור
 * עוברת ל-`previousDraft`; טיוטה מאוחרת ממנו נשארת פעילה. נקודת השמירה
 * נתפסת לפני הייצוא. העותק בקודמת נשאר עד השמירה הבאה. בעבר הוא נמחק,
 * בהיגיון נכון כשלעצמו — העבודה
 * בדיסק, והשארתה הייתה גורמת לפתיחה הבאה להעדיף עותק ישן על הקובץ עצמו. מה
 * שההיגיון הזה פספס: כשמה שהגיע לדיסק הוא בדיוק מה שהמשתמש רצה לבטל, לא
 * נשאר בשום מקום עותק של „לפני”. ההפרדה לשני שדות היא מה שמאפשר את שניהם —
 * `draft` ממשיך לומר „יש עבודה לא-שמורה”, ו-`previousDraft` מחזיק את העותק
 * בלי להתחזות לעבודה.
 *
 * ההעברה חינם מפני שיש שתי משבצות: ב-SDK אין `move`, וב„נתיב אחד” היא הייתה
 * קריאה, כתיבה ומחיקה של כל המסמך בכל שמירה. ראו `nextDraftPath`.
 *
 * **כתיבה שנכשלה אינה מוחקת.** טיוטה שעל הדיסק מאוחרת תמיד לשמירה האחרונה,
 * ולכן היא מחזיקה עבודה שאין בקובץ. „ישנה” היא ביחס למה שעל המסך, לא ביחס
 * לדיסק — ומולו היא תמיד החדשה מבין השתיים.
 *
 * **מעבר מסמך אינו מוחק.** זה נראה סביר — ובפועל השמיד עבודה בשלושה מסלולים
 * שאיש לא התכוון אליהם: פתיחה שנכשלה ונפלה למסמך ריק, טיוטה של מסמך אחר
 * שהמשתמש עוד יחזור אליו, ותשובת „לא” לשאלה על קובץ שהשתנה מבחוץ. בשלושתם
 * לא נאמר „למחוק”. המצביע ברשומה נושא את ה-token של המסמך שהטיוטה שייכת לו,
 * ולכן הוא מזדהה בעצמו ואינו זקוק לניקוי מונע.
 *
 * ## מה שאינו מובטח, ונאמר במפורש
 *
 * `flush` הוא הזדמנות, לא ערובה: הרגע שבו אוצריא נסגרת בכוח או קורסת אינו
 * מריץ שום קוד. מה שמגן שם הם שני הקבועים — `DRAFT_DELAY_MS` בהפוגה הרגילה,
 * ו-`DRAFT_MAX_WAIT_MS` כשההקלדה אינה נעצרת כלל. יחד הם ההבטחה: **לכל היותר
 * דקה של עבודה נמצאת באוויר**, מול ברירת המחדל של Word שהיא עשר דקות.
 */
import type { CaretAnchor } from '../engine/caret-anchor';
import type { WorkspaceWrite } from '../host/workspace';
import {
  activeEntry,
  emptySessionWithId,
  nextDraftPath,
  withActiveEntry,
  type DocumentSessionId,
  type SessionDocument,
  type SessionState,
  type SessionView,
} from './session-state';

/** השהיית הכתיבה של הרשומה, אחרי השינוי האחרון. */
export const PERSIST_DELAY_MS = 2_000;

/** השהיית כתיבת הטיוטה, אחרי העריכה האחרונה. ראו „שני קצבים” בראש הקובץ. */
export const DRAFT_DELAY_MS = 10_000;

/**
 * התקרה על ההשהיה: כמה זמן מותר לעבור בין טיוטה לטיוטה, גם כשההקלדה אינה
 * נעצרת.
 *
 * בלעדיה ההשהיה הייתה נדחית שוב ושוב — זו בדיוק המשמעות של debounce — ומי
 * שכותב ברצף עשר דקות היה מגלה שלא נכתב דבר. עם התקרה, ההבטחה נאמרת במספר
 * אחד: **לכל היותר דקה של עבודה נמצאת באוויר**, ולא משנה כיצד הוא מקליד.
 */
export const DRAFT_MAX_WAIT_MS = 60_000;

export interface SessionKeeperDeps {
  /**
   * מזהה הטאב שהזוכר הזה שייך לו (`DocumentSession.id`, App.vue). זורע בו
   * את ה-state הפנימי (`emptySessionWithId`) — לא `emptySession()` הסתמי —
   * כדי ש-`activeId` שנכתב ל-storage תמיד יצביע על רשומה שבאמת קיימת
   * באוסף שלו. ראו את ההנמקה המלאה ליד `emptySessionWithId`.
   */
  id: DocumentSessionId;
  /** כותבת את הרשומה. כשל אינו מדווח לממשק — הוא כבר בלוג. */
  persist: (state: SessionState, requireSuccess?: boolean) => Promise<void>;
  /** מייצא את המסמך הפעיל, לטיוטה. */
  exportDocument: () => Promise<Blob>;
  /** כותבת את הטיוטה לנתיב שנבחר, ומדווחת אם לא נכתבה ומדוע. */
  writeDraft: (content: Blob, path: string) => Promise<WorkspaceWrite>;
  /** מוחקת קובץ טיוטה. */
  removeDraft: (path: string) => Promise<void>;
  /**
   * שתי המשבצות שהטיוטה נכתבת אליהן (`draftSlotPaths`). הבחירה ביניהן
   * היא של המודול הזה ולא של הקורא — ראו `nextDraftPath`.
   */
  draftPaths: readonly [string, string];
  /**
   * קוראת את מקום הסמן. מקבלת את העוגן הקודם כדי שהקריאה תוכל לדלג על פתירת
   * סדר הפסקאות כשהסמן לא עזב את הפסקה — ראו engine/caret-anchor.ts.
   */
  readCaret: (previous: CaretAnchor | null) => Promise<CaretAnchor | null>;
  /** האם יש עבודה שאינה בדיסק. */
  isDirty: () => boolean;
  /** האם שמירה רצה כרגע. ייצוא נוסף בזמן הזה נדחה ואינו מבוטל. */
  isSaving: () => boolean;
  /**
   * האם פעולה כבדה רצה כרגע (`SaveCoordinator.hold`).
   *
   * שער נפרד מ-`isSaving`, ונחוץ בדיוק כמוהו: תקרת ההשהיה
   * (`DRAFT_MAX_WAIT_MS`) אינה נדחית על ידי הקלדה, ולכן ייצוא מלא יורה גם
   * באמצע כלי שרץ דקה על כל המסמך — כלומר הקפאה שחלה על ה-autosave בלבד
   * הייתה מותירה את המסלול היקר פתוח. וביציאה (`flush`) הסכנה גדולה יותר:
   * ייצוא של מסמך חצי-מוחל היה נכתב לאותה משבצת שבה יושב הצ׳קפוינט שנכתב
   * לפני הכלי, ודורס את העותק שכל התכונה נבנתה כדי לשמור.
   */
  isHeld: () => boolean;
  /**
   * ממתין לסבב השמירה שרץ. בשימוש ב-`flush` בלבד — ראו את ההנמקה שם.
   */
  settleSave: () => Promise<void>;
  /**
   * נקרא כשמסמך גדול מהמכסה, ולכן אין לו טיוטה כלל.
   *
   * זו אינה תקלה חולפת אלא תכונה של המסמך, ולכן היא מדווחת: ההבטחה של
   * המודול הזה — „לכל היותר דקה של עבודה באוויר” — פשוט אינה חלה כאן, ומי
   * שאינו יודע זאת בוטח ברשת ביטחון שאינה פרושה. פעם אחת לכל מסמך, מפני
   * שהתשובה קבועה וחזרה עליה כל עשר שניות היא הטרדה.
   */
  onDraftTooLarge?: () => void;
}

/** ראו `writeDraftNow`. */
type DraftWriteMode = 'timer' | 'final' | 'checkpoint';

export interface SetDocumentOptions {
  /** גודל הקובץ בדיסק, כדי שנדע לזהות עריכה חיצונית בפתיחה הבאה. */
  sourceSize?: number | null;
  /**
   * לא למחוק את הטיוטה.
   *
   * למסלול אחד בלבד: המסמך **נפתח מהטיוטה**. שם הטיוטה אינה שריד של מסמך
   * קודם אלא בדיוק מה שעל המסך, והמחיקה הרגילה הייתה משמידה עבודה שנייה
   * לפני שהיא נכתבת מחדש — ובחלון הזה קריסה מוחקת הכול.
   */
  keepDraft?: boolean;
}

declare const savePointBrand: unique symbol;

/** נקודת הייצוא שייכת לזוכר ולמסמך שלו, ואינה מונה של SaveCoordinator. */
export interface KeeperSavePoint {
  readonly [savePointBrand]: true;
}

export interface NoteSavedOptions {
  sourceSize?: number | null;
  /** „שמור בשם” מחליף יעד של אותה עריכה חיה, בלי לאפס את מוני הטיוטה. */
  document?: SessionDocument;
}

export interface SessionKeeper {
  /** הרשומה כפי שהיא בזיכרון. */
  readonly state: SessionState;
  /**
   * האם יש עבודה שאינה בדיסק **וגם** אינה בטיוטה — כלומר עבודה שקיימת רק
   * בזיכרון של המנוע.
   *
   * זהו השער שלפני „טאב נרדם” (App.vue, `sleepTab`): שחרור מנוע של טאב ברקע
   * הוא מחיקה של כל מה שיש בו בזיכרון, ולכן מותר רק כשכל מה שבו כבר נמצא
   * במקום שאפשר לקרוא ממנו בחזרה. אחרי `flush()` התשובה היא `false` בכל
   * מקרה שבו הטיוטה נכתבה בהצלחה; היא נשארת `true` בדיוק במקרים שבהם
   * הכתיבה לא הצליחה — מסמך גדול מהמכסה, גשר שנפל — ואלה בדיוק המקרים שבהם
   * אסור לשחרר.
   *
   * שני התנאים ולא אחד: `isDirty` לבדו נשאר דלוק גם מיד אחרי כתיבת טיוטה
   * (הוא מדבר על הדיסק, לא על הטיוטה), ומונה השינויים לבדו זז גם מתזוזת סמן
   * במסמך שנשמר כבר.
   */
  readonly hasUnwrittenWork: boolean;
  /** מאמצת רשומה שנקראה בעלייה, לפני שמשהו נכתב עליה. */
  adopt(state: SessionState | null): void;
  /** מסמך נפתח או נסגר. מבטלת עבודה מתוזמנת של המסמך הקודם. */
  setDocument(document: SessionDocument | null, options?: SetDocumentOptions): void;
  /** שומרת בייטים שנפתחו בלי קובץ מקור. אינה מייצאת ואינה תלויה בהקפאת autosave. */
  storeOpenedDraft(content: Blob, name: string): Promise<WorkspaceWrite>;
  /** שינוי במצב התצוגה. */
  updateView(patch: Partial<SessionView>): void;
  /** **המסמך** השתנה. מתזמן גם את רשומת ההפעלה וגם את הטיוטה. */
  noteChange(): void;
  /**
   * **הסמן** זז, והמסמך לא. מתזמן את רשומת ההפעלה בלבד.
   *
   * ההפרדה תוקנה אחרי מדידה: הבחירה מדווחת בכל תו שנקלד **ובנוסף** לעדכון
   * המסמך, ו-`noteChange` לשניהם פירושו שני איפוסים של שני הטיימרים לכל תו.
   * גרוע מזה, הוא קידם את מונה השינויים גם על תזוזת סמן — כלומר
   * `hasUnwrittenWork` הפך אמת, והטיוטה נכתבה שוב על מסמך שלא השתנה מאז
   * הטיוטה הקודמת. הטיוטה היא בייטים של מסמך; הסמן אינו שם.
   *
   * מה שכן שם הוא ברשומה (`caret`, נקרא ב-`persistNow`), ולכן ההשהיה של
   * הרשומה נשארת.
   */
  noteCaretMoved(): void;
  /** נקראת מיד לפני ייצוא snapshot לשמירה, ולא אחרי שהכתיבה הסתיימה. */
  captureSave(): KeeperSavePoint;
  /** רק הטיוטה שנכללה ב-snapshot השמור עוברת לקודמת; עריכות מאוחרות נשארות. */
  noteSaved(point: KeeperSavePoint, options?: NoteSavedOptions): Promise<void>;
  /**
   * כותבת עכשיו עותק של המסמך כפי שהוא — לפני פעולה שאין ממנה דרך חזרה אם
   * התוסף קורס תוך כדי.
   *
   * ההבדל מ-`flush`: הוא כותב רק עבודה שאינה בדיסק, וכאן נדרש עותק גם
   * כשהמסמך **נקי**. בתרחיש שבשבילו זה נבנה — „שמור” ואז כלי כבד — המסמך
   * נקי בדיוק ברגע שהכלי מתחיל, והשמירה האוטומטית שאחריו תכתוב את התוצאה
   * לקובץ. בלי העותק הזה לא נשאר בשום מקום „לפני”.
   *
   * מסמך מלוכלך נכתב כ-`draft`, כי זו באמת עבודה לא-שמורה ושחזור
   * אוטומטי נכון לה. מסמך נקי נכתב ישירות כ-`previousDraft`: הוא זהה
   * לדיסק, ורישום שלו כ-`draft` היה נפתח בעלייה הבאה כ„שינויים שלא נשמרו” שאינם קיימים.
   *
   * מתעלמת מ-`isHeld` — הקורא מקפיא לפניה בכוונה, כדי שסבב autosave לא
   * יתחיל תוך כדי הייצוא. כשל ייצוא נרשם ללוג ואינו מפיל את הפעולה: זו רשת
   * ביטחון, ומי שביקש להריץ כלי יריץ אותו גם בלעדיה.
   */
  checkpoint(): Promise<void>;
  /** המשתמש ביקש במפורש למחוק שינויים שלא נשמרו. ראו את המימוש. */
  discardDraft(): Promise<void>;
  /** כותב עכשיו את כל מה שיש. אידמפוטנטי — ראו host/lifecycle.ts. */
  flush(): Promise<void>;
  dispose(): void;
}

export function createSessionKeeper(deps: SessionKeeperDeps): SessionKeeper {
  let state: SessionState = emptySessionWithId(deps.id);
  let disposed = false;

  let persistTimer: ReturnType<typeof setTimeout> | undefined;
  let draftTimer: ReturnType<typeof setTimeout> | undefined;

  /**
   * סבב אחד בכל רגע, ובסדר. שתי כתיבות במקביל אינן מסוכנות ל-`storage`, אבל
   * שני ייצואים במקביל הם עבודה כפולה על אותו מסמך — ובעיקר: הן יכולות
   * להסתיים בסדר הפוך, כלומר טיוטה ישנה שנכתבת אחרי חדשה.
   */
  let chain: Promise<void> = Promise.resolve();

  /** מזהה המסמך הפתוח. `setDocument` מעלה אותו, וסבב עם ערך ישן נזרק. */
  let epoch = 0;

  /** גודל הקובץ בדיסק, כפי שדווח בפתיחה או בשמירה האחרונה. */
  let sourceSize: number | null = null;

  /** האם כבר נאמר על המסמך הזה שהוא גדול מכדי שתיכתב לו טיוטה. */
  let reportedTooLarge = false;

  /**
   * המועד שאחריו אסור לדחות עוד את כתיבת הטיוטה. `null` = אין עבודה שממתינה
   * לכתיבה, והספירה תתחיל מהעריכה הבאה.
   */
  let draftDeadline: number | null = null;

  /**
   * מונה השינויים, ומה שממנו כבר נכתבה טיוטה.
   *
   * `isDirty` אינו מספיק לשאלה „יש מה לכתוב?”: הוא נשאר `true` כל עוד המסמך
   * לא נשמר לדיסק, כלומר גם מיד אחרי שהטיוטה נכתבה. בלי המונה, שלושת מקורות
   * היציאה (ראו host/lifecycle.ts) היו מייצרים שלושה ייצואים מלאים של אותו
   * מסמך בדיוק, ברגע שהמשתמש עוזב.
   */
  let revision = 0;
  let draftedRevision = 0;
  const savePoints = new WeakMap<KeeperSavePoint, { epoch: number; revision: number }>();

  function clearTimer(timer: ReturnType<typeof setTimeout> | undefined): undefined {
    if (timer !== undefined) clearTimeout(timer);
    return undefined;
  }

  function run(task: () => Promise<void>): Promise<void> {
    chain = chain.then(task, task);
    return chain;
  }

  /** מוסיפה את הסבב לתור ובולעת כשל: זו רשת ביטחון, לא פעולה של המשתמש. */
  function schedule(task: () => Promise<void>): void {
    void run(task).catch((error: unknown) => {
      console.warn('[otzaria-word] שמירת מצב ההפעלה נכשלה', error);
    });
  }

  async function persistNow(): Promise<void> {
    if (disposed) return;
    persistTimer = clearTimer(persistTimer);

    const mine = epoch;
    const caret = await deps.readCaret(activeEntry(state)?.caret ?? null);
    // הסמן נקרא אסינכרונית; אם בינתיים הוחלף המסמך, הוא שייך למסמך שכבר אינו
    // פתוח ואסור לו להיכנס לרשומה של החדש.
    if (disposed || mine !== epoch) return;

    // `caret` שהוא `null` אינו מוחק את מה שידענו: הבחירה עשויה להיות מחוץ
    // למסמך ברגע הזה (מיקוד בשדה בדיאלוג), וזו אינה סיבה לשכוח איפה היה.
    if (caret) state = withActiveEntry(state, { caret });
    await deps.persist(state);
  }

  /**
   * שלושה מצבים, ולכל אחד כללים משלו:
   *
   * - `timer` — הסבב הרגיל. שמירה שרצה או הקפאה דוחות אותו לסבב הבא.
   * - `final` — ההזדמנות האחרונה (יציאה), ואין סבב נוסף אחריה: לשמירה שרצה
   *   הוא **ממתין** במקום לדחות. הקפאה, לעומת זאת, עוצרת גם אותו — ראו
   *   `isHeld` ב-[SessionKeeperDeps].
   * - `checkpoint` — עותק מפורש לפני פעולה כבדה. כותב גם מסמך נקי, ומתעלם
   *   מההקפאה. ראו [SessionKeeper.checkpoint].
   */
  async function writeDraftNow(mode: DraftWriteMode = 'timer'): Promise<void> {
    if (disposed) return;
    draftTimer = clearTimer(draftTimer);

    // הצ׳קפוינט של מסמך נקי אינו „עבודה לא-שמורה”, ולכן הוא נרשם במצביע
    // הקודמת ולא בטיוטה: שלושה צרכנים קוראים „יש טיוטה” כ„יש מה לשמור”.
    const target: 'draft' | 'previousDraft' =
      mode === 'checkpoint' && !deps.isDirty() ? 'previousDraft' : 'draft';

    if (target === 'draft' && (!deps.isDirty() || revision === draftedRevision)) {
      // בצ׳קפוינט זה אינו קיצור סתמי אלא בדיוק הנכון: העותק כבר על הדיסק.
      draftDeadline = null;
      return;
    }

    // פעולה כבדה רצה. הייצוא נדחה — ובמסלול היציאה נעצר לגמרי, כי הוא היה
    // כותב מסמך חצי-מוחל לאותה משבצת שבה יושב הצ׳קפוינט שלפני הכלי.
    if (mode !== 'checkpoint' && deps.isHeld()) {
      if (mode === 'timer') {
        // דוחפים את התקרה, בדיוק כמו מול שמירה שרצה: בלי זה הטיימר שירה
        // מפני שהתקרה הגיעה היה מתזמן את עצמו מחדש בלולאה לכל אורך הכלי.
        draftDeadline = Date.now() + DRAFT_DELAY_MS;
        scheduleDraft();
      }
      return;
    }

    // שמירה שרצה כבר מייצאת את אותו מסמך. הטיוטה ממתינה לסבב הבא במקום
    // להריץ ייצוא שני במקביל לה.
    //
    // הדחייה **דוחפת את התקרה קדימה**, ולא מנסה שוב מיד: כשהטיימר ירה דווקא
    // מפני שהתקרה הגיעה, `scheduleDraft` היה מחשב המתנה של אפס ומתזמן את
    // עצמו מחדש בלולאה — עשרות בדיקות בשנייה לכל אורך הייצוא.
    if (deps.isSaving()) {
      /**
       * ביציאה אין „סבב הבא”. דחייה כאן פירושה שהעבודה שנעשתה מאז הטיוטה
       * האחרונה תלויה כולה בהצלחת השמירה שרצה — וכשל שלה (הקובץ נעול, הרשת
       * נפלה) היה משאיר אותה בלי שום עותק. לכן ממתינים לסבב ובודקים מחדש:
       * אם הוא הצליח, `noteSaved` כבר ניקה ואין מה לכתוב; אם לא, הטיוטה
       * נכתבת. ההמתנה בטוחה גם מפני שהיא מונעת שני ייצואים במקביל על אותו
       * מנוע — בדיוק מה שהדחייה נועדה למנוע.
       */
      if (mode !== 'timer') {
        const mineBeforeWait = epoch;
        await deps.settleSave();
        if (disposed || mineBeforeWait !== epoch) return;
        // הצ׳קפוינט ממשיך גם אם המסמך נקי עכשיו: השמירה שהסתיימה כתבה את
        // „לפני” לדיסק, אבל העותק עדיין נדרש — הכלי עוד לא רץ.
        if (target === 'draft' && (!deps.isDirty() || revision === draftedRevision)) {
          draftDeadline = null;
          return;
        }
      } else {
        draftDeadline = Date.now() + DRAFT_DELAY_MS;
        scheduleDraft();
        return;
      }
    }
    // מכאן והלאה נכתבת טיוטה, ולכן הספירה לתקרה מתחילה מחדש בעריכה הבאה.
    draftDeadline = null;
    const writing = revision;

    const mine = epoch;
    let exported: Blob;
    try {
      exported = await deps.exportDocument();
    } catch (error) {
      console.warn('[otzaria-word] ייצוא הטיוטה נכשל', error);
      return;
    }
    if (disposed || mine !== epoch) return;

    // הנתיב נבחר **אחרי** הייצוא ולא לפניו: `noteSaved` עשוי היה להזיז
    // מצביעים בזמן שהייצוא רץ, והבחירה חייבת לראות את המצב העדכני.
    const entry = activeEntry(state);
    const path = nextDraftPath(
      deps.draftPaths,
      entry?.[target] ?? null,
      entry?.[target === 'draft' ? 'previousDraft' : 'draft'] ?? null,
    );
    const written = await deps.writeDraft(exported, path);
    if (disposed || mine !== epoch) return;

    if (written !== 'written') {
      /**
       * הטיוטה הקודמת **נשארת**, ובכוונה.
       *
       * טיוטה שעל הדיסק מאוחרת תמיד לשמירה האחרונה — `noteSaved` מוחק אותה
       * בכל שמירה — ולכן היא מחזיקה עבודה שאין בקובץ. „ישנה” כאן היא ביחס
       * למה שעל המסך, לא ביחס לדיסק; מולו היא תמיד החדשה מבין השתיים.
       * מחיקתה בגלל כתיבה שנכשלה הייתה משמידה את העבודה היחידה ששרדה, וזה
       * ההפך הגמור ממה שהמודול הזה קיים בשבילו. גיל הטיוטה נאמר למשתמש
       * בשחזור (`savedAt`), וכך „חלקית” אינה מגיעה בהפתעה.
       *
       * `draftedRevision` אינו מתקדם, ולכן הסבב הבא ינסה שוב את אותה עבודה.
       */
      if (written === 'too-large' && !reportedTooLarge) {
        reportedTooLarge = true;
        deps.onDraftTooLarge?.();
      }
      return;
    }

    // מתקדם רק על הטיוטה: צ׳קפוינט של מסמך נקי אינו „העבודה עד כאן נכתבה”,
    // והזזתו הייתה גורמת לסבב הבא לדלג על עבודה אמיתית.
    if (target === 'draft') draftedRevision = writing;
    state = withActiveEntry(state, {
      [target]: {
        path,
        ...(entry?.draft?.name ? { name: entry.draft.name } : {}),
        savedAt: Date.now(),
        documentToken: activeEntry(state)?.document?.token ?? null,
        sourceSize,
      },
    });
    await persistNow();
  }

  function schedulePersist(): void {
    if (disposed) return;
    persistTimer = clearTimer(persistTimer);
    persistTimer = setTimeout(() => schedule(persistNow), PERSIST_DELAY_MS);
  }

  /**
   * מתזמנת כתיבת טיוטה: השהיה מהעריכה האחרונה, אבל לא מעבר ל-`DRAFT_MAX_WAIT_MS`
   * מהעריכה **הראשונה** שטרם נכתבה. ראו את ההנמקה ליד הקבועים.
   */
  function scheduleDraft(): void {
    if (disposed || !deps.isDirty()) return;

    const now = Date.now();
    if (draftDeadline === null) draftDeadline = now + DRAFT_MAX_WAIT_MS;

    draftTimer = clearTimer(draftTimer);
    const wait = Math.max(0, Math.min(DRAFT_DELAY_MS, draftDeadline - now));
    draftTimer = setTimeout(() => schedule(() => writeDraftNow('timer')), wait);
  }

  return {
    get state() {
      return state;
    },

    get hasUnwrittenWork() {
      return deps.isDirty() && revision !== draftedRevision;
    },

    adopt(loaded) {
      state = loaded ?? emptySessionWithId(deps.id);
    },

    setDocument(document, options = {}) {
      const { sourceSize: size = null, keepDraft = false } = options;
      // כל סבב שבאוויר שייך למסמך הקודם, ומכאן והלאה התוצאה שלו נזרקת.
      epoch += 1;
      persistTimer = clearTimer(persistTimer);
      draftTimer = clearTimer(draftTimer);
      draftDeadline = null;
      revision = 0;
      draftedRevision = 0;
      sourceSize = size;
      reportedTooLarge = false;

      /**
       * הסמן נמחק רק כשהמסמך באמת התחלף.
       *
       * שתי הטעויות שהתנאי הזה מונע, וכל אחת מהן בכיוון אחר: מחיקה תמידית
       * הייתה מאבדת את המקום בכל פתיחה חוזרת של **אותו** מסמך — כלומר בדיוק
       * המסלול שהתכונה נכתבה בשבילו. אי-מחיקה הייתה משאירה את הסמן של מסמך
       * א' ברשומה שכבר מזוהה כמסמך ב', ואז הוא היה מוחל עליו: קפיצה שרירותית
       * לאמצע מסמך אחר.
       */
      const previousEntry = activeEntry(state);
      const sameDocument = (previousEntry?.document?.token ?? null) === (document?.token ?? null);

      /**
       * טיוטה שנשמרת עוברת לבעלות המסמך שנפתח ממנה.
       *
       * בלי ההעברה היא הייתה נשארת רשומה על המסמך הקודם — וזה בדיוק המסלול
       * שבו קובץ שלא נמצא נפתח כמסמך חדש: התוכן על המסך, אבל הרשומה טוענת
       * שהוא שייך ל-token שאינו נפתר, ולכן הפתיחה **הבאה** הייתה פוסלת אותו.
       *
       * בלי `keepDraft` המצביע נשאר כפי שהוא — ולא מתאפס. הוא נושא את ה-token
       * של הבעלים שלו, ולכן טיוטה שאינה שייכת למסמך שנפתח נפסלת בקריאה
       * (`decideDraftRecovery`) ואינה זקוקה לניקוי מונע. איפוס כאן היה מנתק
       * את הרשומה מקובץ שעדיין מחזיק עבודה.
       */
      const previousDraft = previousEntry?.draft ?? null;
      const draft =
        keepDraft && previousDraft
          ? { ...previousDraft, documentToken: document?.token ?? null, sourceSize: size }
          : previousDraft;

      state = withActiveEntry(state, {
        document,
        caret: sameDocument ? (previousEntry?.caret ?? null) : null,
        draft,
      });
      schedule(async () => {
        await deps.persist(state);
      });
    },

    async storeOpenedDraft(content, name) {
      // גם בייטים שהגיעו מבחוץ הם עבודה בזיכרון בלבד עד שהכתיבה הצליחה.
      // בכשל אסור ל-hasUnwrittenWork להתיר להרדים ולפרק את המנוע.
      const mine = epoch;
      // המקור הוא גרסת הפתיחה, לא תוכן שכבר הוקלד בזמן סידור התצוגה.
      const writing = revision === 0 ? 1 : 0;
      if (revision === 0) revision = 1;
      let result: WorkspaceWrite = 'failed';
      await run(async () => {
        if (disposed || mine !== epoch) return;
        const entry = activeEntry(state);
        const path = nextDraftPath(deps.draftPaths, entry?.draft ?? null, entry?.previousDraft ?? null);
        result = await deps.writeDraft(content, path);
        if (disposed || mine !== epoch) {
          result = 'failed';
          return;
        }
        if (result !== 'written') return;
        state = withActiveEntry(state, {
          draft: {
            path,
            name,
            savedAt: Date.now(),
            documentToken: entry?.document?.token ?? null,
            sourceSize,
          },
        });
        await deps.persist(state, true);
        if (disposed || mine !== epoch) {
          result = 'failed';
          return;
        }
        // הבייטים והרשומה צריכים להיכתב לפני שמותר להרדים את המנוע.
        // עריכה שהתרחשה בזמן הכתיבה עדיין זקוקה לייצוא משלה.
        draftedRevision = writing;
      }).catch((error: unknown) => {
        console.warn('[otzaria-word] שמירת הטיוטה שנפתחה נכשלה', error);
        result = 'failed';
      });
      return result;
    },

    /**
     * המשתמש אמר במפורש „אל תשמור” — ולכן מה שהטיוטה מחזיקה אינו מבוקש יותר.
     *
     * זהו המסלול **היחיד** שמוחק טיוטה מלבד שמירה מוצלחת, וזו החלטה: מחיקה
     * על כל מעבר מסמך נראתה סבירה ובפועל השמידה עבודה בשלושה מסלולים שאיש
     * לא התכוון אליהם — פתיחה שנכשלה, טיוטה של מסמך אחר שהמשתמש עוד יחזור
     * אליו, ותשובת „לא” לשאלה על קובץ שהשתנה מבחוץ. בכל שלושתם לא נאמר
     * „למחוק”, ובקובץ הזה מוחקים רק כשנאמר.
     */
    async discardDraft() {
      const entry = activeEntry(state);
      const recorded = [entry?.draft?.path, entry?.previousDraft?.path].filter(
        (path): path is string => path !== undefined && path !== null,
      );
      // כשיש מה למחוק, מוחקים גם את שתי המשבצות ולא רק את מה שברשומה: קובץ
      // שנכתב ורגע לפני שנרשם נקטע נשאר אחרת יתום לנצח, ואין ניקוי יתומים
      // בעלייה. נתיב מגרסה קודמת נכלל דרך הרשומה.
      const paths = recorded.length > 0 ? new Set([...recorded, ...deps.draftPaths]) : new Set<string>();
      state = withActiveEntry(state, { draft: null, previousDraft: null });
      await run(async () => {
        for (const path of paths) await deps.removeDraft(path);
        await deps.persist(state);
      });
    },

    updateView(patch) {
      const next = { ...state.view, ...patch };
      const unchanged = (Object.keys(patch) as Array<keyof SessionView>).every(
        (key) => state.view[key] === next[key],
      );
      if (unchanged) return;

      state = { ...state, view: next };
      schedulePersist();
    },

    noteChange() {
      revision += 1;
      schedulePersist();
      scheduleDraft();
    },

    noteCaretMoved() {
      schedulePersist();
    },

    captureSave() {
      const point = {} as KeeperSavePoint;
      savePoints.set(point, { epoch, revision });
      return point;
    },

    async noteSaved(point, options = {}) {
      const saved = savePoints.get(point);
      savePoints.delete(point);
      if (!saved || disposed || saved.epoch !== epoch) return;

      // גם הזזת המצביעים וגם שינוי היעד רצים בתור: טיוטה מאוחרת עשויה
      // להמתין לסיום השמירה ולהיכתב לפני שהרישום הזה מגיע לתורו.
      await run(async () => {
        if (disposed || saved.epoch !== epoch) return;
        sourceSize = options.sourceSize ?? sourceSize;
        const entry = activeEntry(state);
        const document = options.document ?? entry?.document ?? null;
        const sameDocument = (entry?.document?.token ?? null) === (document?.token ?? null);
        state = withActiveEntry(state, {
          document,
          caret: sameDocument ? (entry?.caret ?? null) : null,
        });
        if (entry?.draft) {
          state = draftedRevision > saved.revision
            ? withActiveEntry(state, {
              draft: { ...entry.draft, documentToken: document?.token ?? null, sourceSize },
            })
            : withActiveEntry(state, { draft: null, previousDraft: entry.draft });
        }
        if (revision <= saved.revision) {
          draftTimer = clearTimer(draftTimer);
          draftDeadline = null;
        } else if (draftTimer === undefined && revision !== draftedRevision) {
          // אין איפוס של timer/deadline של הקלדה מאוחרת. אם סבב שכבר היה
          // בתור נכשל, נחדש ניסיון בלי להכריז שהעבודה החדשה נמצאת בטיוטה.
          scheduleDraft();
        }
        await persistNow();
      });
    },

    checkpoint() {
      if (disposed) return Promise.resolve();
      return run(() => writeDraftNow('checkpoint')).catch((error: unknown) => {
        // רשת ביטחון, לא תנאי: מי שביקש להריץ כלי יריץ אותו גם בלי העותק.
        console.warn('[otzaria-word] כתיבת עותק לפני הפעולה נכשלה', error);
      });
    },

    flush() {
      if (disposed) return Promise.resolve();
      persistTimer = clearTimer(persistTimer);
      draftTimer = clearTimer(draftTimer);
      return run(async () => {
        // הטיוטה קודם, והרשומה תמיד: `writeDraftNow` עשוי לצאת מוקדם — מסמך
        // נקי, או שכבר נכתבה טיוטה מאותו שינוי — ואז מצב תצוגה שהמתין
        // להשהיה שבוטלה כאן היה הולך לאיבוד.
        if (deps.isDirty()) await writeDraftNow('final');
        await persistNow();
      }).catch((error: unknown) => {
        console.warn('[otzaria-word] שמירת מצב ההפעלה נכשלה', error);
      });
    },

    dispose() {
      disposed = true;
      epoch += 1;
      persistTimer = clearTimer(persistTimer);
      draftTimer = clearTimer(draftTimer);
      draftDeadline = null;
    },
  };
}
