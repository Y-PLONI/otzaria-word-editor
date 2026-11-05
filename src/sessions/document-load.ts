/**
 * מחוון הטעינה של שורת המצב: פס שמתמלא בזמן שמסמך נפתח, ולצדו „דלג” שמפסיק
 * את ההמתנה.
 *
 * ## למה מודול, ולא מספר ב-App.vue
 *
 * שלוש ההתנהגויות כאן אינן קוסמטיות, ושתיים מהן כבר נשברו פעם אחת במסך
 * הטעינה שב-`index.html` — ולכן הן נכתבות כאן באותה צורה בדיוק:
 *
 * 1. **ההתקדמות אינה נסוגה.** יעד נמוך מהנוכחי נבלע, ואיתו הטקסט שלו. פתיחה
 *    אחת מדווחת משני מקורות שרצים במקביל (המעטפת, ו-`createEditor` מבפנים),
 *    ופס שחוזר אחורה נקרא כתקלה ולא כדיווח שהגיע לא בסדר.
 * 2. **בין תחנה לתחנה הפס זוחל לבדו,** בקצב דועך. השלב שבו המנוע בונה את
 *    המסמך הוא כמעט כל הזמן; פס שעומד בו שתי שניות נקרא כקריסה, לא כאיטיות.
 * 3. **אף תחנה אינה 100%.** רק `finish()` מגיע לשם. „99% שנתקע” הוא בדיוק מה
 *    שגורם למשתמש ללחוץ „דלג” על טעינה שכבר כמעט נגמרה.
 *
 * ## הביטול: מה הוא מבטל ומה לא
 *
 * ה-`attempt` הוא מה שמפריד „הפתיחה הזאת” מ„פתיחה בכלל”: פתיחה שבוטלה, או
 * שפתיחה חדשה יותר החליפה אותה, אינה יכולה יותר לקדם את הפס ואינה יכולה לסגור
 * אותו — אחרת ניסיון שהסתיים באיחור היה מנקה את המחוון של הניסיון שרץ עכשיו.
 * `cancelled` נבדק בין שלב לשלב במעטפת, כי חלק מהשלבים אינם ניתנים לקטיעה
 * באמצע (קריאת הבייטים, התקנת הגופנים) — ומה שאפשר להבטיח הוא שהתוצאה שלהם
 * לא תיושם.
 *
 * מה שהמודול הזה **אינו** יכול לעשות: לשחרר את המנוע. זה תלוי במי שמחזיק את
 * הפתיחה (`EditorSwap.cancel` → ה-`signal` של `createEditor`), ומי שמבטל
 * חייב לקרוא גם לשם. ומה שגם הוא ולא הם אינם יכולים: לשחרר חוט ראשי שנתקע
 * בלולאה של המנוע. שם אין כפתור שיֵירה בכלל — ההגנה היחידה היא לפני שהמנוע
 * רואה את הבייטים (engine/docx-preflight.ts).
 */

/**
 * תחנות ההתקדמות של פתיחת מסמך, באחוזים.
 *
 * החלוקה היא לפי הזמנים שנמדדו ולא לפי מספר השלבים — אותה הכרעה כמו
 * ב-`SPLASH_STAGES` (host/splash.ts), ומאותה מדידה: פרישת המנוע ובניית המסמך
 * הן כמעט כל הזמן. פס שמחלק את הטווח שווה בשווה בין ארבעה שלבים היה קופא על
 * אחד מהם ומדלג על השלושה האחרים בהבזק.
 */
export const LOAD_STAGES = {
  /** קריאת הבייטים והבדיקה המקדימה. */
  reading: 12,
  /** התקנת הגופנים שהמסמך מבקש. */
  fonts: 26,
  /** המנוע בונה את המסמך ומעמד אותו — השלב הארוך. */
  engine: 88,
  /** חיווט המסמך שנפתח: מודדים, סרגל, שחזור התצוגה. */
  arranging: 96,
} as const;

/** תקרת התחנות. `finish()` הוא הדרך היחידה ל-100. */
const CEILING = 99;

/** קצב הזחילה: כמה מהמרווח שנשאר נסגר בכל tick. */
const CREEP_RATE = 0.08;

/** מרווח ה-tick של הזחילה, במילישניות. */
const CREEP_INTERVAL_MS = 60;

/** מתחת למרווח הזה הפס נחשב שהגיע ליעד, והזחילה נעצרת. */
const CREEP_EPSILON = 0.05;

/** כמה זמן הפס נשאר על 100% לפני שהוא נעלם. */
const FINISH_HOLD_MS = 240;

/** מה שהממשק מצייר. `active: false` = אין מחוון על המסך. */
export interface LoadSnapshot {
  active: boolean;
  /** 0–100. משמעותי רק כש-`active`. */
  percent: number;
  /** שם המסמך הנטען, כפי שנמסר ל-`begin`. */
  name: string;
  /** מה קורה כרגע, למשל „בונה את המסמך…”. */
  stage: string;
  /**
   * האם „דלג” אמור להיות על המסך.
   *
   * זה אינו `active`, וההבדל נמדד ברגע אחד: אחרי `finish()` המחוון נשאר על
   * המסך עוד רגע על 100% (ראו `FINISH_HOLD_MS`), ולחיצה על „דלג” בדיוק שם
   * הייתה מבטלת פתיחה שכבר הצליחה — כלומר מוציאה הודעת „בוטלה” על מסמך
   * שנפתח, ואף פותחת מסמך ריק במקומו.
   */
  cancellable: boolean;
}

export function idleLoadSnapshot(): LoadSnapshot {
  return { active: false, percent: 0, name: '', stage: '', cancellable: false };
}

/**
 * ידית לפתיחה אחת. כל הקריאות שלה הן no-op ברגע שהיא אינה הפתיחה הנוכחית —
 * בוטלה, או הוחלפה על ידי פתיחה חדשה יותר.
 */
export interface LoadAttempt {
  /** האם הפתיחה הזאת כבר אינה רלוונטית. נבדק בין שלב לשלב. */
  readonly cancelled: boolean;
  /** מבטל קריאה ועיבוד חבילה שעדיין מתנהלים, גם כשהפתיחה הוחלפה. */
  readonly signal: AbortSignal;
  /** מקדם את הפס לתחנה, עם טקסט מצב. נסיגה נבלעת. */
  stage(target: number, text?: string): void;
  /** הפתיחה הצליחה: 100%, ואז המחוון נעלם. */
  finish(): void;
  /**
   * הפתיחה נכשלה: המחוון נעלם מיד. ההודעה היא של שורת המצב, לא שלו.
   *
   * בטוח לקריאה על פתיחה שכבר הוכרעה — הצליחה, נכשלה או בוטלה — וזה מה
   * שמאפשר לקרוא לו מ-`finally` שסוגר את הפתיחה. בלעדיו חריגה שנזרקה אחרי
   * שהמנוע כבר הצליח הייתה משאירה פס תקוע על 96%.
   */
  fail(): void;
}

export interface DocumentLoadSource {
  /** נקרא על כל שינוי במה שמוצג. */
  onChange(snapshot: LoadSnapshot): void;
  /** להזרקה בבדיקות. ברירת המחדל היא השעון של הדפדפן. */
  timers?: {
    setInterval: (fn: () => void, ms: number) => unknown;
    clearInterval: (handle: unknown) => void;
    setTimeout: (fn: () => void, ms: number) => unknown;
    clearTimeout: (handle: unknown) => void;
  };
}

export interface DocumentLoad {
  /** מה שמוצג כרגע. */
  readonly snapshot: LoadSnapshot;
  /**
   * פותח מחוון חדש ומחזיר את הידית שלו. פתיחה שהייתה בדרך מתבטלת — היא כבר
   * אינה מה שהמשתמש מחכה לו.
   */
  begin(name: string, stage?: string): LoadAttempt;
  /**
   * „דלג”. מבטל את הפתיחה שבדרך ומסלק את המחוון. מחזיר `false` כשלא הייתה
   * פתיחה לבטל — כלומר אין למי שקרא מה לדווח.
   */
  cancel(): boolean;
  dispose(): void;
}

export function createDocumentLoad(source: DocumentLoadSource): DocumentLoad {
  const timers = source.timers ?? {
    setInterval: (fn, ms) => setInterval(fn, ms),
    clearInterval: (handle) => clearInterval(handle as ReturnType<typeof setInterval>),
    setTimeout: (fn, ms) => setTimeout(fn, ms),
    clearTimeout: (handle) => clearTimeout(handle as ReturnType<typeof setTimeout>),
  };

  let published = idleLoadSnapshot();
  let value = 0;
  let target = 0;
  let name = '';
  let stageText = '';
  let active = false;
  /** ראו `LoadSnapshot.cancellable`. */
  let cancellable = false;
  let creep: unknown = null;
  let hold: unknown = null;
  let disposed = false;
  /**
   * דור הפתיחה. ידית שאינה מהדור הנוכחי היא no-op — זה כל מה שמפריד פתיחה
   * שהמשתמש מחכה לה מפתיחה שהוא כבר ויתר עליה.
   */
  let generation = 0;
  let pending: AbortController | null = null;

  function abortPending(): void {
    const controller = pending;
    pending = null;
    controller?.abort();
  }

  function publish(): void {
    if (disposed) return;
    const next: LoadSnapshot = {
      active,
      // הפס נצבע במספר שלם: `width` בשבר עשרוני אינו מזיז פיקסל, והוא כן
      // מייצר דיווח על כל tick.
      percent: Math.round(value),
      name,
      stage: stageText,
      cancellable,
    };
    if (
      next.active === published.active &&
      next.percent === published.percent &&
      next.name === published.name &&
      next.stage === published.stage &&
      next.cancellable === published.cancellable
    ) {
      return;
    }
    published = next;
    source.onChange(next);
  }

  function stopCreep(): void {
    if (creep !== null) {
      timers.clearInterval(creep);
      creep = null;
    }
  }

  function stopHold(): void {
    if (hold !== null) {
      timers.clearTimeout(hold);
      hold = null;
    }
  }

  function tick(): void {
    const gap = target - value;
    if (gap <= CREEP_EPSILON) {
      value = target;
      stopCreep();
      publish();
      return;
    }
    // דעיכה גיאומטרית: תמיד זזה, לעולם אינה עוברת את היעד.
    value += gap * CREEP_RATE;
    publish();
  }

  function startCreep(): void {
    if (creep === null && active && !disposed) creep = timers.setInterval(tick, CREEP_INTERVAL_MS);
  }

  /** מסלק את המחוון מהמסך. אינו נוגע בדור — הקורא הוא זה שמקדם אותו. */
  function clear(): void {
    stopCreep();
    stopHold();
    active = false;
    cancellable = false;
    value = 0;
    target = 0;
    name = '';
    stageText = '';
    publish();
  }

  return {
    get snapshot() {
      return published;
    },

    begin(startName, startStage = 'פותח את המסמך…') {
      // פתיחה חדשה מבטלת את הקודמת: המחוון מתאר מסמך אחד, וזה החדש.
      const mine = (generation += 1);
      abortPending();
      const controller = new AbortController();
      pending = controller;
      if (disposed) abortPending();
      stopHold();
      active = true;
      cancellable = true;
      name = startName;
      stageText = startStage;
      value = 0;
      target = LOAD_STAGES.reading;
      publish();
      startCreep();

      const mineIsCurrent = (): boolean => mine === generation && !disposed;
      /**
       * האם הפתיחה הזאת כבר הוכרעה על ידי עצמה. נפרד מ-`mineIsCurrent`, שהוא
       * „האם היא עוד הפתיחה הנוכחית”: הצלחה **אינה** מקדמת את הדור — היא
       * משאירה את הפס על 100% לרגע — ולכן בלי הדגל הזה `fail()` שאחרי
       * `finish()` היה קוטע בדיוק את הרגע הזה.
       */
      let settled = false;

      return {
        signal: controller.signal,
        get cancelled() {
          return !mineIsCurrent();
        },

        stage(next, text) {
          if (!mineIsCurrent() || settled) return;
          // נסיגה נבלעת — גם באחוזים וגם בטקסט. ראו (1) בראש הקובץ.
          if (!Number.isFinite(next) || next < target) return;
          target = Math.min(next, CEILING);
          if (text) stageText = text;
          publish();
          startCreep();
        },

        finish() {
          if (!mineIsCurrent() || settled) return;
          settled = true;
          pending = null;
          stopCreep();
          // לפני ה-publish: מרגע שהפתיחה הצליחה אין יותר מה לבטל, וההחזקה
          // הקצרה שאחריה אינה חלון שבו „דלג” עוד תופס.
          cancellable = false;
          value = 100;
          target = 100;
          stageText = 'מוכן';
          publish();
          // ההחזקה הקצרה היא מה שמפריד „נגמר” מ„נעלם”: פס שקופץ מ-88 לאין
          // אינו מראה שהוא הגיע לסוף.
          hold = timers.setTimeout(() => {
            hold = null;
            if (mine !== generation) return;
            clear();
          }, FINISH_HOLD_MS);
        },

        fail() {
          if (!mineIsCurrent() || settled) return;
          settled = true;
          // בלי החזקה, ובכוונה: הודעת הכשל יושבת בשורת המצב עצמה, ופס אדום
          // שנשאר לצדה רק גונב את המקום שצריך להיקרא.
          generation += 1;
          abortPending();
          clear();
        },
      };
    },

    cancel() {
      if (!active || !cancellable || disposed) return false;
      generation += 1;
      abortPending();
      clear();
      return true;
    },

    dispose() {
      if (disposed) return;
      stopCreep();
      stopHold();
      disposed = true;
      cancellable = false;
      generation += 1;
      abortPending();
    },
  };
}
