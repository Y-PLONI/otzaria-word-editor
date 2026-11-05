/**
 * מחוון הטעינה של שורת המצב.
 *
 * מה שנמדד כאן הוא לא „הפס זז” אלא ארבע ההכרעות שבלעדיהן מחוון טעינה מזיק
 * יותר מכלום:
 *
 * 1. **ההתקדמות אינה נסוגה.** פתיחה אחת מדווחת משני מקורות במקביל, ופס שחוזר
 *    אחורה נקרא כתקלה.
 * 2. **אף תחנה אינה 100%.** „99% שנתקע” הוא בדיוק מה שגורם ללחוץ „דלג” על
 *    פתיחה שכמעט נגמרה.
 * 3. **„דלג” נעלם ברגע שהפתיחה הצליחה.** ההחזקה הקצרה על 100% היא חלון שבו
 *    לחיצה הייתה מבטלת מסמך שכבר נפתח — ואף פותחת מסמך ריק במקומו.
 * 4. **פתיחה שאינה הנוכחית היא no-op.** ניסיון שהסתיים באיחור אינו מקדם ואינו
 *    סוגר את המחוון של הניסיון שרץ עכשיו.
 *
 * השעונים מוזרקים ולא מדומים גלובלית: הזחילה היא `setInterval` ואילו ההחזקה
 * שאחרי `finish` היא `setTimeout`, והבדיקות כאן צריכות להריץ כל אחד מהם לבד.
 */
import { describe, it, expect, vi } from 'vitest';
import {
  createDocumentLoad,
  idleLoadSnapshot,
  LOAD_STAGES,
  type LoadSnapshot,
} from '../../src/sessions/document-load';

function fakeTimers() {
  let nextId = 1;
  const intervals = new Map<number, () => void>();
  const timeouts = new Map<number, () => void>();

  return {
    api: {
      setInterval: (fn: () => void): unknown => {
        const id = nextId++;
        intervals.set(id, fn);
        return id;
      },
      clearInterval: (handle: unknown): void => {
        intervals.delete(handle as number);
      },
      setTimeout: (fn: () => void): unknown => {
        const id = nextId++;
        timeouts.set(id, fn);
        return id;
      },
      clearTimeout: (handle: unknown): void => {
        timeouts.delete(handle as number);
      },
    },
    /** מריצה ticks של הזחילה. */
    tick(times = 1): void {
      for (let i = 0; i < times; i += 1) {
        for (const fn of [...intervals.values()]) fn();
      }
    },
    /** מריצה את ההחזקה שאחרי `finish`. */
    fireTimeouts(): void {
      for (const [id, fn] of [...timeouts]) {
        timeouts.delete(id);
        fn();
      }
    },
    get creeping(): boolean {
      return intervals.size > 0;
    },
    get holding(): boolean {
      return timeouts.size > 0;
    },
  };
}

function setup() {
  const timers = fakeTimers();
  const changes: LoadSnapshot[] = [];
  const load = createDocumentLoad({
    onChange: (snapshot) => changes.push(snapshot),
    timers: timers.api,
  });
  return { load, timers, changes };
}

/** מרווח ה-ticks שמספיק לזחילה להגיע ליעד ולעצור. */
const SETTLED = 80;

describe('מצב ההתחלה', () => {
  it('בלי פתיחה אין מחוון', () => {
    const { load } = setup();

    expect(load.snapshot).toEqual(idleLoadSnapshot());
    expect(load.snapshot.active).toBe(false);
  });

  it('„דלג” בלי פתיחה אינו מדווח שביטל משהו', () => {
    const { load } = setup();

    // מי שקרא נשען על התשובה כדי להחליט אם להוציא הודעה ולפתוח מסמך ריק.
    expect(load.cancel()).toBe(false);
  });
});

describe('פתיחה', () => {
  it('מציגה את השם, את השלב ואת האפשרות לדלג', () => {
    const { load } = setup();

    load.begin('בראשית.docx', 'קורא את הקובץ…');

    expect(load.snapshot.active).toBe(true);
    expect(load.snapshot.name).toBe('בראשית.docx');
    expect(load.snapshot.stage).toBe('קורא את הקובץ…');
    expect(load.snapshot.cancellable).toBe(true);
    expect(load.snapshot.percent).toBe(0);
  });

  it('הפס זוחל לעבר התחנה גם בלי דיווח נוסף', () => {
    const { load, timers } = setup();
    load.begin('בראשית.docx');

    timers.tick(5);

    // התנועה העצמאית היא מה שמפריד „איטי” מ„תקוע”.
    expect(load.snapshot.percent).toBeGreaterThan(0);
    expect(load.snapshot.percent).toBeLessThanOrEqual(LOAD_STAGES.reading);
  });

  it('הזחילה אינה עוברת את התחנה, ונעצרת כשהגיעה', () => {
    const { load, timers } = setup();
    load.begin('בראשית.docx');

    timers.tick(SETTLED);

    expect(load.snapshot.percent).toBe(LOAD_STAGES.reading);
    // שעון שממשיך לרוץ על פס שהגיע ליעד הוא דיווח מיותר בכל 60 מילישניות.
    expect(timers.creeping).toBe(false);
  });

  it('תחנה חדשה מחדשת את הזחילה', () => {
    const { load, timers } = setup();
    const attempt = load.begin('בראשית.docx');
    timers.tick(SETTLED);
    expect(timers.creeping).toBe(false);

    attempt.stage(LOAD_STAGES.engine, 'בונה את המסמך…');
    expect(load.snapshot.stage).toBe('בונה את המסמך…');

    timers.tick(SETTLED);
    expect(load.snapshot.percent).toBe(LOAD_STAGES.engine);
  });

  it('אף תחנה אינה מגיעה ל-100', () => {
    const { load, timers } = setup();
    const attempt = load.begin('בראשית.docx');

    // גם דיווח מוגזם אינו יכול להגיע לסוף: רק `finish` מגיע לשם.
    attempt.stage(300, 'כמעט…');
    timers.tick(SETTLED * 2);

    expect(load.snapshot.percent).toBe(99);
    expect(load.snapshot.active).toBe(true);
  });
});

describe('ההתקדמות אינה נסוגה', () => {
  it('תחנה נמוכה מהנוכחית נבלעת — היא והטקסט שלה', () => {
    const { load, timers } = setup();
    const attempt = load.begin('בראשית.docx');
    attempt.stage(LOAD_STAGES.engine, 'בונה את המסמך…');
    timers.tick(SETTLED);

    attempt.stage(LOAD_STAGES.fonts, 'מכין את הגופנים…');

    // דיווח שהגיע לא בסדר אינו מזיז את הפס אחורה, וגם אינו מחליף את הטקסט
    // בטקסט של שלב שכבר עבר.
    expect(load.snapshot.percent).toBe(LOAD_STAGES.engine);
    expect(load.snapshot.stage).toBe('בונה את המסמך…');
  });

  it('תחנה שאינה מספר נבלעת', () => {
    const { load, timers } = setup();
    const attempt = load.begin('בראשית.docx');
    timers.tick(SETTLED);

    attempt.stage(Number.NaN, 'שטויות');

    expect(load.snapshot.percent).toBe(LOAD_STAGES.reading);
    expect(load.snapshot.stage).not.toBe('שטויות');
  });
});

describe('סיום', () => {
  it('מגיע ל-100, ואז המחוון נעלם', () => {
    const { load, timers } = setup();
    const attempt = load.begin('בראשית.docx');

    attempt.finish();
    expect(load.snapshot.percent).toBe(100);
    expect(load.snapshot.active).toBe(true);

    timers.fireTimeouts();
    expect(load.snapshot.active).toBe(false);
    expect(load.snapshot).toEqual(idleLoadSnapshot());
  });

  it('„דלג” נעלם ברגע הסיום, ולא רק כשהמחוון יורד', () => {
    const { load } = setup();
    const attempt = load.begin('בראשית.docx');

    attempt.finish();

    // הרגע הזה הוא כל הטעם ב-`cancellable`: הפס עוד על המסך, והפתיחה כבר
    // הצליחה.
    expect(load.snapshot.cancellable).toBe(false);
  });

  it('לחיצה על „דלג” בזמן ההחזקה על 100% אינה מבטלת מסמך שנפתח', () => {
    const { load } = setup();
    const attempt = load.begin('בראשית.docx');
    attempt.finish();

    // `false` הוא מה שמונע מהמעטפת להוציא „הפתיחה הופסקה” ולפתוח מסמך ריק
    // על גבי מסמך שכבר נפתח בהצלחה.
    expect(load.cancel()).toBe(false);
  });

  it('כשל אחרי הצלחה אינו קוטע את ההחזקה על 100%', () => {
    const { load, timers } = setup();
    const attempt = load.begin('בראשית.docx');
    attempt.finish();

    // זה מה שמאפשר ל-`finally` של הפתיחה לקרוא ל-`fail` בלי לדעת אם היא
    // הצליחה. בלי ההגנה הזאת כל פתיחה מוצלחת הייתה מאבדת את רגע הסיום שלה.
    attempt.fail();

    expect(load.snapshot.percent).toBe(100);
    expect(load.snapshot.active).toBe(true);

    timers.fireTimeouts();
    expect(load.snapshot.active).toBe(false);
  });

  it('תחנה אחרי הצלחה נבלעת', () => {
    const { load } = setup();
    const attempt = load.begin('בראשית.docx');
    attempt.finish();

    attempt.stage(LOAD_STAGES.engine, 'בונה את המסמך…');

    expect(load.snapshot.percent).toBe(100);
    expect(load.snapshot.stage).toBe('מוכן');
  });

  it('כשל כפול אינו מדווח פעמיים', () => {
    const { load, changes } = setup();
    const attempt = load.begin('בראשית.docx');
    attempt.fail();
    const after = changes.length;

    attempt.fail();

    expect(changes).toHaveLength(after);
  });

  it('כשל מסלק את המחוון מיד, בלי החזקה', () => {
    const { load, timers } = setup();
    const attempt = load.begin('בראשית.docx');
    attempt.stage(LOAD_STAGES.engine);

    attempt.fail();

    // הודעת הכשל יושבת בשורת המצב עצמה; פס שנשאר לצדה גונב את המקום שלה.
    expect(load.snapshot.active).toBe(false);
    expect(timers.creeping).toBe(false);
    expect(timers.holding).toBe(false);
  });
});

describe('„דלג”', () => {
  it('מסלק את המחוון ומדווח שהיה מה לבטל', () => {
    const { load, timers } = setup();
    load.begin('בראשית.docx');
    timers.tick(3);

    expect(load.cancel()).toBe(true);
    expect(load.snapshot.active).toBe(false);
    expect(timers.creeping).toBe(false);
  });

  it('הפתיחה שבוטלה יודעת שהיא בוטלה', () => {
    const { load } = setup();
    const attempt = load.begin('בראשית.docx');
    expect(attempt.cancelled).toBe(false);

    load.cancel();

    // זה מה שהמעטפת בודקת בין שלב לשלב, כדי לא לכתוב מצב של מסמך שנזנח.
    expect(attempt.cancelled).toBe(true);
    expect(attempt.signal.aborted).toBe(true);
  });

  it('פתיחה שבוטלה אינה יכולה לקדם או לסגור את המחוון', () => {
    const { load, changes } = setup();
    const attempt = load.begin('בראשית.docx');
    load.cancel();
    const after = changes.length;

    attempt.stage(LOAD_STAGES.engine, 'בונה את המסמך…');
    attempt.finish();

    expect(load.snapshot.active).toBe(false);
    expect(changes).toHaveLength(after);
  });

  it('לחיצה שנייה אינה מדווחת שביטלה עוד פתיחה', () => {
    const { load } = setup();
    load.begin('בראשית.docx');

    expect(load.cancel()).toBe(true);
    // אחרת „דלג” כפול היה פותח שני מסמכים ריקים.
    expect(load.cancel()).toBe(false);
  });
});

describe('שתי פתיחות', () => {
  it('פתיחה חדשה מבטלת את הקודמת ומציגה את המסמך החדש', () => {
    const { load } = setup();
    const first = load.begin('בראשית.docx');

    const second = load.begin('שמות.docx');

    expect(first.cancelled).toBe(true);
    expect(second.cancelled).toBe(false);
    expect(first.signal.aborted).toBe(true);
    expect(second.signal.aborted).toBe(false);
    expect(load.snapshot.name).toBe('שמות.docx');
    expect(load.snapshot.percent).toBe(0);
  });

  it('הפתיחה הישנה שמסתיימת באיחור אינה מסלקת את המחוון של החדשה', () => {
    const { load, timers } = setup();
    const first = load.begin('בראשית.docx');
    const second = load.begin('שמות.docx');
    second.stage(LOAD_STAGES.engine, 'בונה את המסמך…');
    timers.tick(SETTLED);

    first.finish();
    timers.fireTimeouts();

    // זה בדיוק המסלול שבו מחוון נעלם באמצע טעינה בלי סיבה נראית.
    expect(load.snapshot.active).toBe(true);
    expect(load.snapshot.name).toBe('שמות.docx');
    expect(load.snapshot.percent).toBe(LOAD_STAGES.engine);
  });

  it('ההחזקה של סיום קודם אינה מסלקת פתיחה שהתחילה מיד אחריו', () => {
    const { load, timers } = setup();
    const first = load.begin('בראשית.docx');
    first.finish();

    load.begin('שמות.docx');
    timers.fireTimeouts();

    expect(load.snapshot.active).toBe(true);
    expect(load.snapshot.name).toBe('שמות.docx');
  });
});

describe('דיווח', () => {
  it('מדווח רק על שינוי אמיתי', () => {
    const { load, timers, changes } = setup();
    const attempt = load.begin('בראשית.docx');
    timers.tick(SETTLED);
    const settled = changes.length;

    // אותה תחנה, אותו טקסט, ועוד ticks על פס שהגיע ליעד.
    attempt.stage(LOAD_STAGES.reading);
    timers.tick(10);

    expect(changes).toHaveLength(settled);
  });

  it('האחוז המדווח שלם — פס אינו זז בשבר עשרוני', () => {
    const { load, timers, changes } = setup();
    load.begin('בראשית.docx');
    timers.tick(20);

    for (const change of changes) {
      expect(Number.isInteger(change.percent)).toBe(true);
    }
  });
});

describe('פירוק', () => {
  it('מפסיק עיבוד של פתיחה קיימת וגם פתיחה שנקראה אחרי פירוק', () => {
    const { load } = setup();
    const attempt = load.begin('בראשית.docx');
    load.dispose();
    expect(attempt.signal.aborted).toBe(true);
    expect(load.begin('שמות.docx').signal.aborted).toBe(true);
  });

  it('עוצר את הזחילה ואינו מדווח יותר', () => {
    const { load, timers, changes } = setup();
    const attempt = load.begin('בראשית.docx');
    const before = changes.length;

    load.dispose();

    expect(timers.creeping).toBe(false);
    timers.tick(10);
    attempt.stage(LOAD_STAGES.engine);
    attempt.finish();
    expect(changes).toHaveLength(before);
  });

  it('פירוק כפול אינו זורק', () => {
    const { load } = setup();
    load.begin('בראשית.docx');

    expect(() => {
      load.dispose();
      load.dispose();
    }).not.toThrow();
  });

  it('„דלג” אחרי פירוק אינו מדווח שביטל', () => {
    const { load } = setup();
    load.begin('בראשית.docx');
    load.dispose();

    expect(load.cancel()).toBe(false);
  });
});

describe('התחנות', () => {
  it('עולות בסדר, ואף אחת אינה 100', () => {
    const stages = [
      LOAD_STAGES.reading,
      LOAD_STAGES.fonts,
      LOAD_STAGES.engine,
      LOAD_STAGES.arranging,
    ];

    for (let i = 1; i < stages.length; i += 1) {
      expect(stages[i]).toBeGreaterThan(stages[i - 1]);
    }
    expect(stages[stages.length - 1]).toBeLessThan(100);
  });

  it('השלב הארוך מקבל את רוב הטווח', () => {
    // המדידה שמאחורי החלוקה: פרישת המנוע ובניית המסמך הן כמעט כל הזמן. פס
    // שמחלק את הטווח שווה בשווה בין ארבעה שלבים קופא על אחד מהם.
    const engineShare = LOAD_STAGES.engine - LOAD_STAGES.fonts;
    const rest = 100 - engineShare;

    expect(engineShare).toBeGreaterThan(rest / 2);
  });
});

describe('אין תלות בשעון גלובלי', () => {
  it('בלי הזרקה נופל לשעון של הסביבה', () => {
    vi.useFakeTimers();
    try {
      const changes: LoadSnapshot[] = [];
      const load = createDocumentLoad({ onChange: (snapshot) => changes.push(snapshot) });

      load.begin('בראשית.docx');
      vi.advanceTimersByTime(600);

      expect(load.snapshot.percent).toBeGreaterThan(0);
      load.dispose();
    } finally {
      vi.useRealTimers();
    }
  });
});
