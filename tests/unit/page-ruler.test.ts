/**
 * מצב הסרגל: המדידה של העמוד המצויר, הקריאה מהמסמך, והכתיבה חזרה.
 *
 * שלוש המשפחות שנבדקות כאן הן בדיוק שלוש הדרכים שבהן סרגל נשבר בלי להיראות
 * שבור:
 *
 *   1. **מדידה** — מלבן שאינו מתעדכן בגלילה, או שמתעדכן על כל פיקסל.
 *   2. **קריאה** — תשובה של מסמך שכבר נסגר שנוחתת על המסמך שנפתח אחריו, וקריאה
 *      שרצה כשהסרגל בכלל מוסתר.
 *   3. **כתיבה** — גרירה של כניסת פסקה ש**מוחקת** כניסת שורה ראשונה, מפני
 *      ש-`setIndentation` מחליף את `<w:ind>` כולו. זה נמדד על המנוע, וזו
 *      הסיבה ש-`applyRulerIndents` קיים בכלל.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  applyRulerIndents,
  createRulerModel,
  measurePageRect,
  paintedHost,
  readRulerUnit,
  watchPageRect,
  RULER_SELECTION_DEBOUNCE_MS,
  SETTLE_DELAYS_MS,
  type RulerReading,
} from '../../src/engine/page-ruler';
import type { PageMarginsState } from '../../src/engine/page-setup';
import type { ParagraphIndentReading } from '../../src/engine/paragraph-format';

/* ------------------------------------------------------------------ */
/* עזרי DOM                                                            */
/* ------------------------------------------------------------------ */

/** jsdom אינו מפריס, ולכן כל מלבן כאן נקבע במפורש. */
function withRect<T extends HTMLElement>(
  element: T,
  left: number,
  width: number,
  top = 0,
  height = 22,
): T {
  element.getBoundingClientRect = () =>
    ({ left, right: left + width, width, top, bottom: top + height, height, x: left, y: top }) as DOMRect;
  return element;
}

function pageHost(pageLeft = 120, pageWidth = 794): { host: HTMLElement; page: HTMLElement } {
  const host = withRect(document.createElement('div'), 0, 900);
  const page = withRect(document.createElement('div'), pageLeft, pageWidth);
  page.setAttribute('data-page-index', '0');
  host.appendChild(page);
  document.body.appendChild(host);
  return { host, page };
}

afterEach(() => {
  document.body.innerHTML = '';
});

/* ------------------------------------------------------------------ */

describe('paintedHost', () => {
  it('מחזיר את ה-host של המנוע', () => {
    const element = document.createElement('div');
    expect(paintedHost({ viewport: { getHost: () => element } })).toBe(element);
  });

  it('גרסה בלי `viewport`, או קריאה שזורקת, אינן מפילות דבר', () => {
    expect(paintedHost(null)).toBeNull();
    expect(paintedHost({})).toBeNull();
    expect(
      paintedHost({
        viewport: {
          getHost: () => {
            throw new Error('לא היום');
          },
        },
      }),
    ).toBeNull();
  });
});

describe('measurePageRect', () => {
  it('מודד את העמוד ביחס לאלמנט הייחוס', () => {
    const { host } = pageHost(120, 794);
    const reference = withRect(document.createElement('div'), 20, 900);

    expect(measurePageRect(host, reference)).toEqual({
      leftPx: 100,
      widthPx: 794,
      topPx: 0,
      heightPx: 22,
    });
  });

  it('בלי עמוד מצויר אין מלבן — וזה מצב רגיל, לא כשל', () => {
    const host = withRect(document.createElement('div'), 0, 900);
    expect(measurePageRect(host, host)).toBeNull();
    expect(measurePageRect(null, host)).toBeNull();
  });

  it('עמוד ברוחב אפס אינו מלבן', () => {
    const { host, page } = pageHost();
    withRect(page, 0, 0);
    expect(measurePageRect(host, host)).toBeNull();
  });
});

describe('watchPageRect', () => {
  beforeEach(() => {
    // רוב הבדיקות כאן מודדות סינכרונית; ה-rAF של jsdom היה דוחה אותן לפריים.
    vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
      callback(0);
      return 1;
    });
    vi.stubGlobal('cancelAnimationFrame', () => {});
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('מודד מיד, ומדווח שוב רק על שינוי אמיתי', () => {
    const { host, page } = pageHost(120, 794);
    const seen: Array<{ leftPx: number; widthPx: number } | null> = [];
    const watch = watchPageRect({ host, reference: host, onChange: (rect) => seen.push(rect) });

    expect(seen[0]).toMatchObject({ leftPx: 120, widthPx: 794 });

    watch.measure();
    expect(seen).toHaveLength(1); // אותו מלבן — אין דיווח שני

    withRect(page, 60, 794);
    watch.measure();
    expect(seen).toHaveLength(2);
    expect(seen[1]).toMatchObject({ leftPx: 60, widthPx: 794 });

    watch.dispose();
  });

  it('גלילה של ה-host מזמינה מדידה', () => {
    const { host, page } = pageHost(120, 794);
    const seen: unknown[] = [];
    const watch = watchPageRect({ host, reference: host, onChange: (rect) => seen.push(rect) });

    withRect(page, 20, 794);
    host.dispatchEvent(new Event('scroll'));

    expect(seen).toHaveLength(2);
    watch.dispose();
  });

  it('`viewport.observe` של המנוע מחובר, ומנותק בפירוק', () => {
    const { host, page } = pageHost();
    let listener: null | (() => void) = null;
    const capture = (callback: () => void): void => {
      listener = callback;
    };
    let unsubscribed = false;
    const watch = watchPageRect({
      host,
      reference: host,
      ui: {
        viewport: {
          observe: (callback) => {
            capture(callback);
            return () => {
              unsubscribed = true;
            };
          },
        },
      },
      onChange: () => {},
    });

    expect(listener).toBeTypeOf('function');
    withRect(page, 5, 794);
    (listener as unknown as () => void)();

    watch.dispose();
    expect(unsubscribed).toBe(true);
  });

  it('`measure` מודדת שוב אחרי שהמנוע סיים לצייר', () => {
    // נמדד על ה-dist הארוז: שינוי זום מגיע אלינו כשהציור מחדש רק מתחיל,
    // ומדידה יחידה באותו רגע תופסת את הגיאומטריה הישנה.
    vi.useFakeTimers();
    try {
      const { host, page } = pageHost(120, 794);
      const seen: unknown[] = [];
      const watch = watchPageRect({ host, reference: host, onChange: (rect) => seen.push(rect) });

      watch.measure();
      expect(seen).toHaveLength(1); // עדיין אותו מלבן

      // המנוע מסיים לצייר רק עכשיו, אחרי שהמדידה המיידית כבר רצה.
      withRect(page, 250, 555);
      vi.advanceTimersByTime(SETTLE_DELAYS_MS[SETTLE_DELAYS_MS.length - 1] + 10);

      expect(seen).toHaveLength(2);
      expect(seen[1]).toMatchObject({ leftPx: 250, widthPx: 555 });
      watch.dispose();
    } finally {
      vi.useRealTimers();
    }
  });

  it('מדידות ההמתנה מבוטלות בפירוק', () => {
    vi.useFakeTimers();
    try {
      const { host, page } = pageHost();
      const seen: unknown[] = [];
      const watch = watchPageRect({ host, reference: host, onChange: (rect) => seen.push(rect) });
      watch.measure();
      watch.dispose();

      withRect(page, 400, 794);
      vi.advanceTimersByTime(2000);

      expect(seen).toHaveLength(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it('אחרי הפירוק אין דיווח — גם לא מגלילה שנקלטה באותו רגע', () => {
    const { host, page } = pageHost();
    const seen: unknown[] = [];
    const watch = watchPageRect({ host, reference: host, onChange: (rect) => seen.push(rect) });
    watch.dispose();

    withRect(page, 400, 794);
    host.dispatchEvent(new Event('scroll'));
    watch.measure();

    expect(seen).toHaveLength(1); // רק המדידה הראשונה, מלפני הפירוק
  });
});

/* ------------------------------------------------------------------ */
/* המודל                                                               */
/* ------------------------------------------------------------------ */

const PAGE: PageMarginsState = {
  pageWidthTwips: 11906,
  pageHeightTwips: 16838,
  leftTwips: 1440,
  rightTwips: 1440,
  topTwips: 1440,
  bottomTwips: 1440,
  effectiveTopTwips: 1440,
  effectiveBottomTwips: 1440,
  direction: 'rtl',
};

const PARAGRAPH: ParagraphIndentReading = {
  target: { kind: 'block', nodeType: 'paragraph', nodeId: 'p1' },
  indents: { leftTwips: 720, rightTwips: 0, firstLineTwips: 0, hangingTwips: 0, bidi: true },
};

describe('createRulerModel', () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  function model(overrides: Partial<Parameters<typeof createRulerModel>[0]> = {}) {
    const readings: Array<RulerReading | null> = [];
    const readPage = vi.fn(async (): Promise<PageMarginsState | null> => PAGE);
    const readIndents = vi.fn(async (): Promise<ParagraphIndentReading | null> => PARAGRAPH);
    const source = {
      readPage,
      readIndents,
      onChange: (next: RulerReading | null) => readings.push(next),
      ...overrides,
    };
    return { adapter: createRulerModel(source), readings, source, readPage };
  }

  it('סרגל מוסתר אינו קורא כלום — `doc.get` סורק את המסמך כולו', async () => {
    const { adapter, source } = model();
    adapter.noteSelectionChanged();
    await vi.advanceTimersByTimeAsync(1000);

    expect(source.readPage).not.toHaveBeenCalled();
    expect(source.readIndents).not.toHaveBeenCalled();
    adapter.dispose();
  });

  it('הדלקה קוראת מיד', async () => {
    const { adapter, readings } = model();
    adapter.setEnabled(true);
    await vi.advanceTimersByTimeAsync(0);

    expect(readings).toHaveLength(1);
    expect(readings[0]?.page).toEqual(PAGE);
    expect(readings[0]?.indents?.leftTwips).toBe(720);
    expect(adapter.getState()?.target).toEqual(PARAGRAPH.target);
    adapter.dispose();
  });

  it('כיבוי מנקה את המצב מיד', async () => {
    const { adapter, readings } = model();
    adapter.setEnabled(true);
    await vi.advanceTimersByTimeAsync(0);
    adapter.setEnabled(false);

    expect(readings[readings.length - 1]).toBeNull();
    expect(adapter.getState()).toBeNull();
    adapter.dispose();
  });

  it('תזוזת סמן מושהית, ושלוש תזוזות רצופות הן קריאה אחת', async () => {
    const { adapter, readPage } = model();
    adapter.setEnabled(true);
    await vi.advanceTimersByTimeAsync(0);
    readPage.mockClear();

    adapter.noteSelectionChanged();
    adapter.noteSelectionChanged();
    adapter.noteSelectionChanged();
    expect(readPage).not.toHaveBeenCalled();

    await vi.advanceTimersByTimeAsync(RULER_SELECTION_DEBOUNCE_MS + 5);
    expect(readPage).toHaveBeenCalledTimes(1);
    adapter.dispose();
  });

  it('מדווח רק על שינוי אמיתי — אחרת כל הקשה הייתה מרנדרת את הסרגל', async () => {
    const { adapter, readings } = model();
    adapter.setEnabled(true);
    await vi.advanceTimersByTimeAsync(0);
    adapter.refreshNow();
    await vi.advanceTimersByTimeAsync(0);

    expect(readings).toHaveLength(1);
    adapter.dispose();
  });

  it('שינוי בשוליים העליונים בלבד מדווח — הסרגל האנכי תלוי בו', async () => {
    // נמדד על ה-dist הארוז: השוואה חלקית שדילגה על `topTwips` השאירה את
    // הסרגל האנכי במקומו בזמן שהטקסט במסמך זז 64px למטה.
    let page: PageMarginsState = PAGE;
    const { adapter, readings } = model({ readPage: vi.fn(async () => page) });
    adapter.setEnabled(true);
    await vi.advanceTimersByTimeAsync(0);
    expect(readings).toHaveLength(1);

    page = { ...PAGE, topTwips: 2880 };
    adapter.refreshNow();
    await vi.advanceTimersByTimeAsync(0);

    expect(readings).toHaveLength(2);
    expect(readings[1]?.page.topTwips).toBe(2880);
    adapter.dispose();
  });

  it('כשל בקריאת המקטע מוחזר כ„אין מה לצייר”, ולא כחריגה', async () => {
    const { adapter, readings } = model({
      readPage: vi.fn(async () => {
        throw new Error('המסמך נסגר');
      }),
    });
    adapter.setEnabled(true);
    await vi.advanceTimersByTimeAsync(0);

    expect(readings[readings.length - 1] ?? null).toBeNull();
    adapter.dispose();
  });

  it('אין סמן במסמך — יש עמוד, אין סמני כניסה', async () => {
    const { adapter, readings } = model({ readIndents: vi.fn(async () => null) });
    adapter.setEnabled(true);
    await vi.advanceTimersByTimeAsync(0);

    expect(readings[0]?.page).toEqual(PAGE);
    expect(readings[0]?.indents).toBeNull();
    expect(readings[0]?.target).toBeNull();
    adapter.dispose();
  });

  it('אחרי הפירוק אין דיווח — גם מקריאה שכבר הייתה באוויר', async () => {
    const pending: Array<(value: PageMarginsState) => void> = [];
    const { adapter, readings } = model({
      readPage: vi.fn(
        () =>
          new Promise<PageMarginsState>((resolve) => {
            pending.push(resolve);
          }),
      ),
    });
    adapter.setEnabled(true);
    adapter.dispose();
    for (const resolve of pending) resolve(PAGE);
    await vi.advanceTimersByTimeAsync(10);

    expect(readings).toEqual([]);
  });
});

/* ------------------------------------------------------------------ */
/* כתיבה                                                               */
/* ------------------------------------------------------------------ */

describe('applyRulerIndents', () => {
  function host() {
    const calls: Array<Record<string, unknown>> = [];
    return {
      calls,
      superdoc: {
        activeEditor: {
          doc: {
            format: {
              paragraph: {
                setIndentation: (input: Record<string, unknown>) => {
                  calls.push(input);
                  return { success: true };
                },
              },
            },
          },
        },
      },
    };
  }

  const target = { kind: 'block', nodeType: 'paragraph', nodeId: 'p1' } as const;

  it('צד ההתחלה נכתב ל-`left` וצד הסוף ל-`right`, גם בפסקה עברית', async () => {
    const { calls, superdoc } = host();
    await applyRulerIndents(
      superdoc,
      target,
      { leftTwips: 0, rightTwips: 0, firstLineTwips: 0, hangingTwips: 0, bidi: true },
      { startTwips: 720, endTwips: 360 },
    );

    expect(calls[0]).toMatchObject({ left: 720, right: 360 });
  });

  it('כניסת שורה ראשונה קיימת נשמרת — `setIndentation` מחליף את האלמנט כולו', async () => {
    const { calls, superdoc } = host();
    await applyRulerIndents(
      superdoc,
      target,
      { leftTwips: 0, rightTwips: 0, firstLineTwips: 567, hangingTwips: 0, bidi: false },
      { startTwips: 720, endTwips: 0 },
    );

    expect(calls[0]).toMatchObject({ left: 720, right: 0, firstLine: 567 });
    expect(calls[0]).not.toHaveProperty('hanging');
  });

  it('כניסה תלויה קיימת נשמרת אף היא', async () => {
    const { calls, superdoc } = host();
    await applyRulerIndents(
      superdoc,
      target,
      { leftTwips: 100, rightTwips: 0, firstLineTwips: 0, hangingTwips: 283, bidi: false },
      { startTwips: 1440, endTwips: 0 },
    );

    expect(calls[0]).toMatchObject({ left: 1440, hanging: 283 });
    expect(calls[0]).not.toHaveProperty('firstLine');
  });

  it('בלי „מיוחד” לא נשלח לא זה ולא זה', async () => {
    const { calls, superdoc } = host();
    await applyRulerIndents(
      superdoc,
      target,
      { leftTwips: 0, rightTwips: 0, firstLineTwips: 0, hangingTwips: 0, bidi: false },
      { startTwips: 0, endTwips: 0 },
    );

    expect(calls[0]).not.toHaveProperty('firstLine');
    expect(calls[0]).not.toHaveProperty('hanging');
  });
});

describe('readRulerUnit', () => {
  it('הולך אחרי המנוע', () => {
    expect(readRulerUnit({ getMeasurementUnit: () => 'in' })).toBe('in');
    expect(readRulerUnit({ getMeasurementUnit: () => 'cm' })).toBe('cm');
  });

  it('גרסה בלי הפונקציה, ערך זר או קריאה שזורקת — סנטימטרים', () => {
    expect(readRulerUnit(null)).toBe('cm');
    expect(readRulerUnit({})).toBe('cm');
    expect(readRulerUnit({ getMeasurementUnit: () => 'parsec' })).toBe('cm');
    expect(
      readRulerUnit({
        getMeasurementUnit: () => {
          throw new Error('לא היום');
        },
      }),
    ).toBe('cm');
  });
});
