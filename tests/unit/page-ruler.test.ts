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
  directionFromText,
  measureAllPageTextSegments,
  measurePageGlyphs,
  measurePageRect,
  paintedHost,
  sameTextSegments,
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

/**
 * `directionFromText` — באג אמיתי שנתפס בצילום מסך אמיתי (לא בבדיקת יחידה):
 * גרסה ראשונה קבעה כיוון לפי `getComputedStyle(container).direction` בלבד,
 * וזה מיקם ¶ בקצה הלא-נכון על טקסט אנגלי בתוך פסקה RTL (ברירת המחדל של
 * מסמך חדש כאן) — ה-¶ צויר **לפני** המילה הראשונה במקום **אחרי** האחרונה.
 * התיקון: הכיוון נגזר מהתוכן עצמו (התו-החזק האחרון), ו-CSS הוא רק נפילה
 * לאחור כשאין בתוכן אף תו חזק. ראו engine/formatting-marks-layer.ts
 * ו-scripts/qa/pilcrow-overlay-qa.mjs.
 */
describe('directionFromText', () => {
  it('טקסט אנגלי בלבד — ltr, גם אם יש בסופו פיסוק', () => {
    expect(directionFromText('paragraph one is short')).toBe('ltr');
    expect(directionFromText('hello.')).toBe('ltr');
  });

  it('טקסט עברי בלבד — rtl', () => {
    expect(directionFromText('שלום עולם בעברית')).toBe('rtl');
  });

  it('התו-החזק **האחרון**, לא הראשון — קובע את הכיוון', () => {
    // המקרה שנתפס בצילום המסך: פסקה אנגלית שגולשת, השורה האחרונה מסתיימת
    // ב"page" — חייב rtl=false גם אם משהו אחר בתחילת הפסקה היה עברי.
    expect(directionFromText('שלום world')).toBe('ltr');
    expect(directionFromText('hello עולם')).toBe('rtl');
  });

  it('בלי אף תו חזק (ספרות/פיסוק/רווח בלבד, או placeholder של פסקה ריקה) — null, לא ניחוש', () => {
    expect(directionFromText('123')).toBeNull();
    expect(directionFromText('!? .,')).toBeNull();
    expect(directionFromText(' ')).toBeNull();
    expect(directionFromText('')).toBeNull();
  });
});

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
    return { adapter: createRulerModel(source), readings, source, readPage, readIndents };
  }

  it('סרגל מוסתר אינו קורא כלום', async () => {
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
  it('coalesces document and selection events into one trailing read', async () => {
    const { adapter, readPage, readIndents } = model();
    adapter.setEnabled(true);
    await vi.advanceTimersByTimeAsync(0);
    readPage.mockClear();
    readIndents.mockClear();
    adapter.noteDocumentChanged();
    adapter.noteSelectionChanged();
    adapter.noteDocumentChanged();
    await vi.advanceTimersByTimeAsync(1000);
    expect(readPage).toHaveBeenCalledTimes(1);
    expect(readIndents).toHaveBeenCalledTimes(1);
    adapter.dispose();
  });

  it('refreshNow cancels the pending debounced read', async () => {
    const { adapter, readPage } = model();
    adapter.setEnabled(true);
    await vi.advanceTimersByTimeAsync(0);
    readPage.mockClear();
    adapter.noteDocumentChanged();
    adapter.noteSelectionChanged();
    adapter.refreshNow();
    await vi.advanceTimersByTimeAsync(1000);
    expect(readPage).toHaveBeenCalledTimes(1);
    adapter.dispose();
  });

  it('rejects an old response immediately after a new event, before its debounce expires', async () => {
    let resolve!: (value: ParagraphIndentReading) => void;
    const { adapter, readings } = model({
      readIndents: vi.fn(() => new Promise<ParagraphIndentReading>((done) => { resolve = done; })),
    });
    adapter.setEnabled(true);
    await vi.advanceTimersByTimeAsync(0);
    adapter.noteSelectionChanged();
    resolve(PARAGRAPH);
    await vi.advanceTimersByTimeAsync(0);
    expect(readings).toEqual([]);
    adapter.dispose();
  });

  it('serializes refreshes while an old read is pending and reads the latest state once', async () => {
    let resolve!: (value: PageMarginsState) => void;
    const readPage = vi.fn(async () => PAGE).mockImplementationOnce(() => new Promise((done) => { resolve = done; }));
    const { adapter, readings } = model({ readPage });
    adapter.setEnabled(true);
    adapter.noteDocumentChanged();
    adapter.noteSelectionChanged();
    await vi.advanceTimersByTimeAsync(1000);
    adapter.refreshNow();
    adapter.refreshNow();
    expect(readPage).toHaveBeenCalledTimes(1);
    resolve(PAGE);
    await vi.advanceTimersByTimeAsync(0);
    expect(readPage).toHaveBeenCalledTimes(2);
    expect(readings).toHaveLength(1);
    adapter.dispose();
  });

  it('publishes a changed write target even when the story has an identical paragraph ID and indents', async () => {
    let paragraph = PARAGRAPH;
    const { adapter, readings } = model({ readIndents: vi.fn(async () => paragraph) });
    adapter.setEnabled(true);
    await vi.advanceTimersByTimeAsync(0);
    paragraph = { ...PARAGRAPH, target: { ...PARAGRAPH.target, story: { kind: 'story', storyType: 'headerFooterPart', refId: 'header1' } } };
    adapter.refreshNow();
    await vi.advanceTimersByTimeAsync(0);
    expect(readings).toHaveLength(2);
    expect(adapter.getState()?.target).toEqual(paragraph.target);
    paragraph = { ...paragraph, target: { ...paragraph.target, nodeType: 'heading' } };
    adapter.refreshNow();
    await vi.advanceTimersByTimeAsync(0);
    expect(readings).toHaveLength(3);
    adapter.dispose();
  });

  it('does not dispatch a queued refresh after disabling the ruler', async () => {
    let resolve!: (value: PageMarginsState) => void;
    const readPage = vi.fn(() => new Promise<PageMarginsState>((done) => { resolve = done; }));
    const { adapter, readings } = model({ readPage });
    adapter.setEnabled(true);
    adapter.refreshNow();
    adapter.setEnabled(false);
    resolve(PAGE);
    await vi.advanceTimersByTimeAsync(1000);
    expect(readPage).toHaveBeenCalledTimes(1);
    expect(readings).toEqual([]);
    adapter.dispose();
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

/* ------------------------------------------------------------------ */
/* טווחים בתוך שורת טקסט — לבדיקת האיות                                */
/* ------------------------------------------------------------------ */

/**
 * מה שנמדד כאן הוא בדיוק מה שהופך „מילה מסומנת” ל„חצי מילה מסומנת”: הקיבוץ
 * של צמתי הטקסט. מילה שהעיצוב משתנה באמצעה יושבת בשני צמתים, ומדידה
 * צומת-צומת הייתה מסמנת שתי שגיאות במקום ערך מוכר אחד — ולהפך, קיבוץ שחוצה
 * גבול בלוק היה מדביק סוף פסקה לתחילת הבאה ויוצר „מילה” שאינה קיימת.
 *
 * `getClientRects` מזויף: jsdom אינו מפריס, ולכן הרוחב נגזר מאורך הטווח.
 */
describe('measureAllPageTextSegments', () => {
  const CHAR_PX = 10;
  let restoreRects: (() => void) | null = null;

  function fakeRects(): void {
    const original = Range.prototype.getClientRects;
    Range.prototype.getClientRects = function fake(this: Range) {
      const length = this.toString().length;
      const rect = {
        left: this.startOffset * CHAR_PX,
        top: 0,
        width: length * CHAR_PX,
        height: 20,
        right: (this.startOffset + length) * CHAR_PX,
        bottom: 20,
      } as DOMRect;
      return Object.assign([rect], { item: () => rect }) as unknown as DOMRectList;
    };
    restoreRects = () => {
      Range.prototype.getClientRects = original;
    };
  }

  /** עמוד עם בלוק אחד שבתוכו הצמתים שנמסרו, כל אחד ב-`<span>` משלו. */
  function pageWithRuns(...runs: string[][]): { host: HTMLElement; root: HTMLElement } {
    const { host, page } = pageHost();
    for (const block of runs) {
      const div = document.createElement('div');
      for (const text of block) {
        const span = document.createElement('span');
        span.textContent = text;
        div.appendChild(span);
      }
      page.appendChild(div);
    }
    const root = withRect(document.createElement('div'), 0, 900);
    document.body.appendChild(root);
    return { host, root };
  }

  beforeEach(fakeRects);
  afterEach(() => {
    restoreRects?.();
    restoreRects = null;
  });

  it('מוצאת טווח בתוך צומת יחיד, עם המלבן שלו', () => {
    const { host, root } = pageWithRuns(['אבגד הוזח']);
    const found = measureAllPageTextSegments(host, root, (text) => [
      { start: text.indexOf('הוזח'), end: text.indexOf('הוזח') + 4 },
    ]);

    expect(found).toHaveLength(1);
    expect(found[0]!.text).toBe('הוזח');
    expect(found[0]!.rects[0]).toMatchObject({ widthPx: 40, heightPx: 20 });
  });

  it('שני צמתים באותו בלוק הם טקסט אחד — מילה שהעיצוב משתנה באמצעה', () => {
    const { host, root } = pageWithRuns(['תוס', 'פות']);
    const seen: string[] = [];
    measureAllPageTextSegments(host, root, (text) => {
      seen.push(text);
      return [];
    });
    expect(seen).toEqual(['תוספות']);
  });

  it('שני בלוקים אינם מתחברים — סוף פסקה אינו נדבק לתחילת הבאה', () => {
    const { host, root } = pageWithRuns(['ראשונה'], ['שנייה']);
    const seen: string[] = [];
    measureAllPageTextSegments(host, root, (text) => {
      seen.push(text);
      return [];
    });
    expect(seen).toEqual(['ראשונה', 'שנייה']);
  });

  it('טווח שחוצה שני צמתים נמדד כמלבן אחד', () => {
    const { host, root } = pageWithRuns(['תוס', 'פות']);
    const found = measureAllPageTextSegments(host, root, () => [{ start: 0, end: 6 }]);
    expect(found[0]!.text).toBe('תוספות');
    expect(found[0]!.rects).toHaveLength(1);
  });

  it('עמוד שמחוץ לחלון אינו נסרק כלל', () => {
    const { host, page } = pageHost();
    // הרבה מתחת ל-host (שגובהו 22px, ברירת המחדל של `withRect`) ומעבר לשוליים.
    withRect(page, 120, 794, 5_000, 1_000);
    const div = document.createElement('div');
    div.textContent = 'רחוק';
    page.appendChild(div);
    const root = withRect(document.createElement('div'), 0, 900);
    document.body.appendChild(root);

    const seen: string[] = [];
    measureAllPageTextSegments(host, root, (text) => {
      seen.push(text);
      return [];
    });
    expect(seen).toEqual([]);
  });

  it('גם עמוד שנגלל מעל החלון אינו נסרק', () => {
    // הענף השני של הסינון. בלעדיו מסמך שנגללו בו עשרה עמודים היה ממשיך
    // למדוד את כולם בכל פריים.
    const { host, page } = pageHost();
    withRect(page, 120, 794, -5_000, 1_000);
    const div = document.createElement('div');
    div.textContent = 'למעלה';
    page.appendChild(div);
    const root = withRect(document.createElement('div'), 0, 900);
    document.body.appendChild(root);

    const seen: string[] = [];
    measureAllPageTextSegments(host, root, (text) => {
      seen.push(text);
      return [];
    });
    expect(seen).toEqual([]);
  });

  it('יש תקרה גם בלי `limit` מפורש', () => {
    const { host, root } = pageWithRuns(['א'.repeat(1_000)]);
    const found = measureAllPageTextSegments(host, root, (text) =>
      [...text].map((_, index) => ({ start: index, end: index + 1 })),
    );
    expect(found.length).toBeLessThan(1_000);
    expect(found.length).toBeGreaterThan(0);
  });

  it('טווח שחורג מהטקסט מדולג, ואינו מפיל את שאר המדידה', () => {
    // `select` הוא קוד חיצוני. `Range.setEnd` על היסט מחוץ לצומת זורק, וזריקה
    // כאן הייתה מבטלת את המדידה של כל שאר העמוד — לא רק של הטווח הפגום.
    const { host, root } = pageWithRuns(['אבגד']);
    const found = measureAllPageTextSegments(host, root, () => [
      { start: 0, end: 99 },
      { start: 0, end: 2 },
    ]);
    expect(found.map((item) => item.text)).toEqual(['אב']);
  });

  it('`limit` חוסם מסמך פתולוגי', () => {
    const { host, root } = pageWithRuns(['אבגדהוזחט']);
    const found = measureAllPageTextSegments(
      host,
      root,
      (text) => [...text].map((_, index) => ({ start: index, end: index + 1 })),
      { limit: 3 },
    );
    expect(found).toHaveLength(3);
  });

  it('בלי host או בלי reference — אין מדידה', () => {
    expect(measureAllPageTextSegments(null, document.createElement('div'), () => [])).toEqual([]);
    expect(measureAllPageTextSegments(document.createElement('div'), null, () => [])).toEqual([]);
  });
});

describe('sameTextSegments', () => {
  const segment = (text: string, leftPx: number) => ({
    text,
    rects: [{ leftPx, topPx: 0, widthPx: 40, heightPx: 20 }],
  });

  it('אותה מדידה — שקולה', () => {
    expect(sameTextSegments([segment('אבג', 10)], [segment('אבג', 10)])).toBe(true);
  });

  it('הזזה של פחות מחצי פיקסל אינה שינוי', () => {
    expect(sameTextSegments([segment('אבג', 10)], [segment('אבג', 10.4)])).toBe(true);
  });

  it('טקסט אחר, מלבן שזז, או מספר אחר — שינוי', () => {
    expect(sameTextSegments([segment('אבג', 10)], [segment('דהו', 10)])).toBe(false);
    expect(sameTextSegments([segment('אבג', 10)], [segment('אבג', 30)])).toBe(false);
    expect(sameTextSegments([segment('אבג', 10)], [])).toBe(false);
  });
});

/**
 * `measurePageGlyphs` — הגליפים שלחיצה מוצמדת אליהם (engine/pointer-snap.ts).
 *
 * jsdom אינו מפריס, ולכן כל צומת טקסט „מצויר” במלבן של ה-`<span>` שמכיל אותו
 * (`withRect`), ו-`getClientRects` של טווח מחזיר אותו. מה שנמדד: ההיקף נגזר
 * מהמטרה (שורה, עמוד, גרירה), אובייקטים אינם טקסט, וכותרות אינן גוף.
 */
describe('measurePageGlyphs', () => {
  let restoreRects: (() => void) | null = null;

  beforeEach(() => {
    const original = Range.prototype.getClientRects;
    Range.prototype.getClientRects = function fake(this: Range) {
      const node = this.startContainer;
      const element = node.nodeType === Node.TEXT_NODE ? node.parentElement : (node as Element);
      const rect = element ? element.getBoundingClientRect() : null;
      const list = rect && rect.width > 0 && rect.height > 0 ? [rect] : [];
      return Object.assign(list, { item: (i: number) => list[i] ?? null }) as unknown as DOMRectList;
    };
    restoreRects = () => {
      Range.prototype.getClientRects = original;
    };
  });
  afterEach(() => {
    restoreRects?.();
    restoreRects = null;
  });

  /** עמוד בפינה (100, 200) עם שורה עברית ארוכה, שורה ריקה, ושורה קצרה — כפי שנמדד. */
  function page(): {
    host: HTMLElement;
    page: HTMLElement;
    lines: HTMLElement[];
    runs: HTMLElement[];
  } {
    const host = withRect(document.createElement('div'), 0, 1000, 0, 1500);
    const pageEl = withRect(document.createElement('div'), 100, 800, 200, 1100);
    pageEl.setAttribute('data-page-index', '0');
    host.appendChild(pageEl);
    document.body.appendChild(host);

    const lines: HTMLElement[] = [];
    const runs: HTMLElement[] = [];
    const add = (text: string, left: number, width: number, top: number): void => {
      const fragment = document.createElement('div');
      const line = withRect(document.createElement('div'), 196, 600, top, 18);
      const run = withRect(document.createElement('span'), left, width, top, 17);
      run.textContent = text;
      line.appendChild(run);
      fragment.appendChild(line);
      pageEl.appendChild(fragment);
      lines.push(line);
      runs.push(run);
    };
    add('שורה ראשונה ארוכה יותר', 540, 158, 314);
    add(' ', 693, 5, 332);
    add('קצר', 673, 25, 351);
    return { host, page: pageEl, lines, runs };
  }

  it('מטרה על שורה — הגליפים של אותה שורה בלבד, ביחס לפינת העמוד', () => {
    const { host, lines } = page();
    const measured = measurePageGlyphs(host, lines[0]!, 150, 322);
    expect(measured).not.toBeNull();
    expect(measured!.rects).toEqual([{ leftPx: 440, topPx: 114, widthPx: 158, heightPx: 17 }]);
    expect(measured!.pageBox()).toEqual({ leftPx: 100, topPx: 200, widthPx: 800, heightPx: 1100 });
  });

  it('מטרה על צומת הטקסט עצמו — ההורה הוא ההיקף', () => {
    const { host, runs } = page();
    const measured = measurePageGlyphs(host, runs[2]!.firstChild, 680, 360);
    expect(measured!.rects).toHaveLength(1);
    expect(measured!.rects[0]).toMatchObject({ leftPx: 573 });
  });

  it('מטרה על העמוד — כל השורות', () => {
    const { host, page: pageEl } = page();
    const measured = measurePageGlyphs(host, pageEl, 300, 900);
    expect(measured!.rects.map((rect) => rect.topPx)).toEqual([114, 132, 151]);
  });

  it('כותרת עליונה אינה גוף: מדולגת בהיקף העמוד, ונמדדת כשלוחצים עליה', () => {
    const { host, page: pageEl } = page();
    const header = document.createElement('div');
    header.setAttribute('data-layout-story', 'header');
    const run = withRect(document.createElement('span'), 600, 90, 230, 17);
    run.textContent = 'כותרת';
    header.appendChild(run);
    pageEl.insertBefore(header, pageEl.firstChild);

    expect(measurePageGlyphs(host, pageEl, 300, 900)!.rects.map((rect) => rect.topPx)).toEqual([114, 132, 151]);
    expect(measurePageGlyphs(host, header, 300, 238)!.rects).toEqual([{ leftPx: 500, topPx: 30, widthPx: 90, heightPx: 17 }]);
  });

  it('גרירה (בלי מטרה) — העמוד שמתחת לנקודה; מחוץ לכל עמוד — כלום', () => {
    const { host } = page();
    expect(measurePageGlyphs(host, null, 300, 900)!.rects).toHaveLength(3);
    expect(measurePageGlyphs(host, null, 50, 900)).toBeNull();
    expect(measurePageGlyphs(host, null, 300, 1400)).toBeNull();
  });

  it('תמונה — בחירת אובייקט, לא טקסט', () => {
    const { host, lines } = page();
    const image = withRect(document.createElement('img'), 560, 40, 314, 17);
    lines[0]!.appendChild(image);
    expect(measurePageGlyphs(host, image, 570, 320)).toBeNull();
  });

  it('אלמנט בתוך העמוד בלי טקסט תחתיו — ידית או מסגרת — לא מוצמד', () => {
    const { host, page: pageEl } = page();
    const handle = withRect(document.createElement('div'), 500, 8, 400, 8);
    pageEl.appendChild(handle);
    expect(measurePageGlyphs(host, handle, 504, 404)).toBeNull();
  });

  it('שכבה מעל העמוד: מכסה אותו — כמו לחיצה על העמוד; פקד קטן — לא', () => {
    const { host } = page();
    const layer = withRect(document.createElement('div'), 0, 1000, 0, 1500);
    host.appendChild(layer);
    expect(measurePageGlyphs(host, layer, 300, 900)!.rects).toHaveLength(3);

    const widget = withRect(document.createElement('div'), 300, 20, 900, 20);
    host.appendChild(widget);
    expect(measurePageGlyphs(host, widget, 310, 910)).toBeNull();

    // ה-host עצמו הוא שכבה שמכסה הכול.
    expect(measurePageGlyphs(host, host, 300, 900)!.rects).toHaveLength(3);
  });

  it('בלי host או בלי עמודים — אין מדידה', () => {
    expect(measurePageGlyphs(null, null, 0, 0)).toBeNull();
    expect(measurePageGlyphs(document.createElement('div'), null, 0, 0)).toBeNull();
  });
});
