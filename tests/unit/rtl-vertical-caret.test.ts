/**
 * חצים אנכיים בשורה עברית — התיקון.
 *
 * ## מאיפה המספרים
 *
 * הקצוות של כל שורה כאן הם מה ש-Chrome צייר בגששים
 * ‏`scripts/qa/rtl-vertical-caret-probe.mjs` ו-`rtl-vertical-seam-probe.mjs`
 * (superdoc 2.15.0, ה-dist הארוז, 22.9.2026):
 *
 *   | השורה                 | הקצוות שנמדדו   | היסט הסיום שנמדד |
 *   |-----------------------|-----------------|------------------|
 *   | עברית ארוכה           | 548.2 … 1043.4  | —                |
 *   | „שורה קצרה”           | 980.1 … 1043.4  | 9                |
 *   | „קצרה וממורכזת”       | 699.8 … 785.2   | 13               |
 *   | בקרה: לטינית קצרה     | 441.6 … 504.3   | 10               |
 *
 * החלוקה **בתוך** הקצוות אחידה, ולא נמדדה תו-תו: ההחלטה שנבדקת כאן תלויה
 * בקצוות ובחריצים שנגזרים מהם, ולא ברוחב של אות מסוימת. מה שכן מגיע ישירות
 * מהמדידה הוא היסט הסיום של כל שורה — הערך שהמנוע החזיר כשהתבקש לסמן את סופה
 * — וזה בדיוק מה שהבדיקות קובעות שנכתב.
 *
 * ובכל אחד מהמקרים העבריים המנוע נחת ב**קצה ההפוך**: ירידה שעמודת המטרה שלה
 * מעבר לסוף השורה נחתה על היסט 0, וירידה לשורה ממורכזת שהעמודה מימין לתחילתה
 * נחתה על היסט הסיום. חמש הבקרות הלועזיות באותו מסמך נחתו נכון.
 *
 * ## הדמה כאן **מציירת**, ולא רק מדווחת
 *
 * זה לא נוי. הגרסה הראשונה של הקובץ הזה החזיקה תצלום קבוע וסמן שאינו זז,
 * ו**שלוש מוטציות שרדו אותו**: ביטול האיפוס במקש שאינו חץ, ביטול `owns`,
 * וזריעה מחדש של העמודה בכל הקשה. כולן עברו ירוק, מפני שבדמה סטטית ההקשה
 * השנייה זהה לראשונה ואין לה ממה להיכשל.
 *
 * לכן `setSelectionTarget` כאן מזיז את הסמן: הוא מעדכן את התצלום **ואת ה-x
 * וה-y של הסמן המצויר**, לפי אותו כלל שנמדד — כולל שתפר גלישה נפתר לשורה
 * הקודמת. רק כך „ההקשה השנייה” היא באמת הקשה שנייה.
 */
import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  goalInside,
  installRtlVerticalArrows,
  isVerticalArrow,
  landingSlot,
  nearestSlot,
  type VerticalCaretHost,
} from '../../src/engine/rtl-vertical-caret';
import type { CaretSlot, PaintedChar } from '../../src/engine/painted-lines';

/* ------------------------------------------------------------------ */
/* הפונקציות הטהורות                                                    */
/* ------------------------------------------------------------------ */

const key = (over: Partial<KeyboardEvent> = {}) =>
  ({
    key: 'ArrowDown',
    ctrlKey: false,
    metaKey: false,
    altKey: false,
    shiftKey: false,
    ...over,
  }) as KeyboardEvent;

describe('isVerticalArrow', () => {
  it('חץ למטה וחץ למעלה נקיים — שלנו', () => {
    expect(isVerticalArrow(key({ key: 'ArrowDown' }))).toBe(true);
    expect(isVerticalArrow(key({ key: 'ArrowUp' }))).toBe(true);
  });

  it('חצים אופקיים אינם שלנו', () => {
    expect(isVerticalArrow(key({ key: 'ArrowRight' }))).toBe(false);
    expect(isVerticalArrow(key({ key: 'ArrowLeft' }))).toBe(false);
  });

  it('Shift, Ctrl, Alt ו-Meta נשארים למנוע', () => {
    expect(isVerticalArrow(key({ shiftKey: true }))).toBe(false);
    expect(isVerticalArrow(key({ ctrlKey: true }))).toBe(false);
    expect(isVerticalArrow(key({ altKey: true }))).toBe(false);
    expect(isVerticalArrow(key({ metaKey: true }))).toBe(false);
  });

  it('הרכבה במקלדת — החץ שייך לחלונית ההרכבה', () => {
    expect(isVerticalArrow(key({ isComposing: true } as Partial<KeyboardEvent>))).toBe(false);
    expect(isVerticalArrow(key({ keyCode: 229 } as Partial<KeyboardEvent>))).toBe(false);
  });
});

describe('goalInside', () => {
  const short = { left: 980.1, right: 1043.4 };

  it('עמודה בתוך השורה — המנוע צודק', () => {
    expect(goalInside(short, 1021.6)).toBe(true);
  });

  it('עמודה מעבר לסוף השורה (משמאל, בעברית) — המנוע הפוך', () => {
    expect(goalInside(short, 916.1)).toBe(false);
  });

  it('עמודה לפני תחילת השורה (מימין) — גם זה מחוץ לשורה', () => {
    expect(goalInside({ left: 699.8, right: 785.2 }, 1027.2)).toBe(false);
  });

  it('הסבילות מכווצת את „בפנים”, ואינה מרחיבה אותו', () => {
    // „בטעות בחוץ” עולה כלום — אנחנו מחזירים בדיוק את החריץ שהמנוע היה בוחר.
    // „בטעות בפנים” עולה רוחב שורה: מוסרים למנוע במקום שנמדד בו הפוך.
    expect(goalInside(short, 980.5)).toBe(false);
    expect(goalInside(short, 1042.9)).toBe(false);
    expect(goalInside(short, 1000)).toBe(true);
  });

  it('הפרש של עשירית פיקסל בין שתי שורות אינו נחשב „בפנים”', () => {
    // צירוף מקרים שגרתי ברוחבי טקסט: שורת המקור נגמרת ב-548.2 ושורת היעד
    // ב-548.0. הרחבת הגבול הייתה מוסרת כאן למנוע, והוא מקפיץ לקצה ההפוך.
    expect(goalInside({ left: 548.0, right: 1043.4 }, 548.2)).toBe(false);
  });
});

describe('nearestSlot', () => {
  const slots: CaretSlot[] = [
    { pm: 79, x: 1043.4 },
    { pm: 83, x: 1015.3 },
    { pm: 88, x: 980.1 },
  ];

  it('בוחר את החריץ הקרוב ביותר לעמודה', () => {
    expect(nearestSlot(slots, 1014)?.pm).toBe(83);
  });

  it('עמודה מחוץ לשורה נופלת על החריץ שבקצה הקרוב', () => {
    expect(nearestSlot(slots, 500)?.pm).toBe(88);
    expect(nearestSlot(slots, 2000)?.pm).toBe(79);
  });

  it('בשוויון נשמר החריץ הראשון, כדי שהקשה חוזרת לא תקפיץ', () => {
    const tie: CaretSlot[] = [
      { pm: 5, x: 100 },
      { pm: 9, x: 100 },
    ];
    expect(nearestSlot(tie, 100)?.pm).toBe(5);
  });

  it('שורה בלי חריצים', () => {
    expect(nearestSlot([], 100)).toBeNull();
  });
});

describe('landingSlot — תפר הגלישה', () => {
  const slots: CaretSlot[] = [
    { pm: 94, x: 1043.4 },
    { pm: 95, x: 1036.3 },
    { pm: 96, x: 1029.2 },
  ];

  it('בלי תפר — החריץ הקרוב ביותר, גם כשהוא הראשון', () => {
    expect(landingSlot(slots, 1043.4, false)?.pm).toBe(94);
  });

  it('עם תפר — החריץ הראשון מוחלף בזה שאחריו, כי אותו אי אפשר לבטא', () => {
    expect(landingSlot(slots, 1043.4, true)?.pm).toBe(95);
  });

  it('חריץ אחר אינו מושפע מהתפר', () => {
    expect(landingSlot(slots, 1029, true)?.pm).toBe(96);
  });

  it('שורה שכל מה שיש בה הוא התפר — אין מה לבטא, וההקשה נמסרת', () => {
    expect(landingSlot([{ pm: 94, x: 1043.4 }], 1043.4, true)).toBeNull();
  });
});

/* ------------------------------------------------------------------ */
/* מסמך מצויר, ודמה שמציירת                                            */
/* ------------------------------------------------------------------ */

/**
 * jsdom אינו פורס, ולכן `getBoundingClientRect` מחזיר אפסים — גם ל-`Range`.
 * הדריסה מזינה את התיבות שנבנו, בדיוק כמו ב-`rtl-caret.test.ts`.
 */
const originalRangeRect = Range.prototype.getBoundingClientRect;

function stubRects(): void {
  Range.prototype.getBoundingClientRect = function (this: Range) {
    const parent = (this.startContainer as Text).parentElement as
      | (HTMLElement & { __chars?: PaintedChar[]; __shift?: () => number })
      | null;
    const box = parent?.__chars?.[this.startOffset];
    // `__shift` הוא הגלילה האופקית של ה-host: מלבני החלון זזים איתה.
    const shift = parent?.__shift?.() ?? 0;
    return { left: (box?.left ?? 0) - shift, right: (box?.right ?? 0) - shift } as DOMRect;
  };
}

afterEach(() => {
  Range.prototype.getBoundingClientRect = originalRangeRect;
  document.body.innerHTML = '';
  vi.restoreAllMocks();
});

interface LineSpec {
  pmStart: number;
  /** כמה אותיות בשורה. כל אחת יחידת pm אחת. */
  count: number;
  left: number;
  right: number;
  rtl?: boolean;
  /** שורה בלי תוכן — טווח באורך אפס. */
  empty?: boolean;
}

interface ParaSpec {
  /** `null` = טבלה: fragment בלי `data-source-node-id`. */
  id: string | null;
  lines: LineSpec[];
}

/** שורה מצוירת, כפי שהדמה צריכה לדעת אותה כדי להזיז את הסמן. */
interface PaintedLine {
  blockId: string | null;
  blockStart: number;
  spec: LineSpec;
  top: number;
  element: HTMLElement;
}

/** התיבות של שורה, מחולקות אחיד בין הקצוות שנמדדו. */
function boxesOf(spec: LineSpec): PaintedChar[] {
  const width = (spec.right - spec.left) / spec.count;
  return Array.from({ length: spec.count }, (_, i) => ({
    pm: spec.pmStart + i,
    left: spec.rtl === false ? spec.left + i * width : spec.right - (i + 1) * width,
    right: spec.rtl === false ? spec.left + (i + 1) * width : spec.right - i * width,
  }));
}

/**
 * ה-x שהמנוע מצייר בו את ההיסט הזה בתוך השורה.
 *
 * הסמן יושב על הקצה **המוביל** של התו — ימין בשורה עברית — ובסוף השורה על
 * הקצה הנגרר של האחרון. זו אותה נוסחה שב-`caretSlots`, וכאן היא משמשת לכיוון
 * ההפוך: מהיסט אל פיקסל, כדי שהדמה תוכל לצייר.
 */
function slotX(spec: LineSpec, pm: number): number {
  const index = pm - spec.pmStart;
  const width = (spec.right - spec.left) / spec.count;
  if (spec.rtl === false) {
    return index >= spec.count ? spec.right : spec.left + index * width;
  }
  return index >= spec.count ? spec.left : spec.right - index * width;
}

function buildLine(spec: LineSpec, top: number, shiftOf: () => number): HTMLElement {
  const line = document.createElement('div');
  const pmEnd = spec.empty ? spec.pmStart : spec.pmStart + spec.count;
  line.setAttribute('data-pm-start', String(spec.pmStart));
  line.setAttribute('data-pm-end', String(pmEnd));
  if (spec.rtl !== false) line.setAttribute('dir', 'rtl');
  line.getBoundingClientRect = () =>
    ({
      top,
      bottom: top + 20,
      left: spec.left - shiftOf(),
      right: spec.right - shiftOf(),
    }) as DOMRect;
  if (spec.empty) return line;

  const chars = boxesOf(spec);
  const span = document.createElement('span') as HTMLElement & {
    __chars?: PaintedChar[];
    __shift?: () => number;
  };
  span.className = 'superdoc-text-run';
  span.setAttribute('data-pm-start', String(spec.pmStart));
  span.setAttribute('data-pm-end', String(pmEnd));
  span.textContent = 'א'.repeat(spec.count);
  span.__chars = chars;
  span.__shift = shiftOf;
  // הריצה המצוירת נמתחת על התווים שבה — זה מה ש-`lineExtent` קורא.
  span.getBoundingClientRect = () =>
    ({
      left: Math.min(...chars.map((c) => c.left)) - shiftOf(),
      right: Math.max(...chars.map((c) => c.right)) - shiftOf(),
      width: spec.right - spec.left,
      top,
      bottom: top + 20,
    }) as DOMRect;
  line.append(span);
  return line;
}

interface Setup {
  host: HTMLElement;
  /** השורה שהסמן עליה בתחילת הבדיקה. */
  source: HTMLElement;
  superdoc: VerticalCaretHost;
  setSelectionTarget: ReturnType<typeof vi.fn>;
  /** כמה פעמים היירוט קרא את התצלום — כלומר ראה את ההקשה והגיע להכרעה. */
  seen: () => number;
  /** איפה הסמן עכשיו, אחרי כל מה שנכתב. */
  caret: () => { blockId: string; offset: number; x: number };
  /** אלמנטי השורות שנבנו, לפי הסדר במסמך. */
  lineEls: HTMLElement[];
  /** הזזת הסמן בידי **מישהו אחר** — המנוע, המאחז, מודול שכן. */
  moveCaret: (blockId: string, offset: number) => void;
  /** גלילה אופקית של ה-host, כמו זו ש-`caret-visibility.ts` מבצע. */
  scrollHost: (dx: number) => void;
}

/**
 * בונה מסמך מצויר מרשימת פסקאות, וממקם את הסמן בשורה ובעמודה נתונות.
 *
 * הדמה **מציירת**: כתיבה דרך `setSelectionTarget` מזיזה את הסמן המצויר אל
 * ה-x וה-y של ההיסט שנכתב, לפי אותם כללים שנמדדו — כולל שתפר גלישה נפתר
 * לשורה הקודמת. בלי זה אין בדיקה אמיתית להקשה שנייה ברצף.
 */
function setup({
  paras,
  sourcePara = 0,
  sourceLine = 0,
  caretOffset,
  caretX,
  withCaret = true,
}: {
  paras: ParaSpec[];
  sourcePara?: number;
  sourceLine?: number;
  caretOffset: number;
  caretX: number;
  withCaret?: boolean;
}): Setup {
  stubRects();

  const host = document.createElement('div');
  const page = document.createElement('div');
  page.className = 'superdoc-page';

  let top = 0;
  let shift = 0;
  const shiftOf = () => shift;
  let source: HTMLElement | null = null;
  const painted: PaintedLine[] = [];

  paras.forEach((para, paraIndex) => {
    const fragment = document.createElement('div');
    if (para.id !== null) fragment.setAttribute('data-source-node-id', para.id);
    const first = para.lines[0]!;
    const last = para.lines[para.lines.length - 1]!;
    fragment.setAttribute('data-pm-start', String(first.pmStart));
    fragment.setAttribute(
      'data-pm-end',
      String(last.empty ? last.pmStart : last.pmStart + last.count),
    );

    para.lines.forEach((spec, lineIndex) => {
      const line = buildLine(spec, top, shiftOf);
      painted.push({ blockId: para.id, blockStart: first.pmStart, spec, top, element: line });
      top += 20;
      fragment.append(line);
      if (paraIndex === sourcePara && lineIndex === sourceLine) source = line;
    });
    page.append(fragment);
  });

  host.append(page);

  const sourceId = paras[sourcePara]!.id!;
  let current = { blockId: sourceId, offset: caretOffset };
  // ה-x נשמר בקואורדינטות תוכן, בדיוק כמו במנוע — הציור מחסיר את הגלילה.
  let caretPos = { x: caretX, top: (source as unknown as HTMLElement).getBoundingClientRect().top };

  const caret = document.createElement('div');
  caret.className = 'sd-v2-local-selection-caret';
  caret.getBoundingClientRect = () =>
    ({
      left: caretPos.x - shift,
      right: caretPos.x - shift,
      top: caretPos.top,
      bottom: caretPos.top + 20,
    }) as DOMRect;
  if (withCaret) host.append(caret);
  document.body.append(host);

  /** מזיזה את הסמן המצויר להיסט שנכתב — התפר שייך לשורה הקודמת. */
  const repaint = (blockId: string, offset: number): void => {
    const candidates = painted.filter((line) => line.blockId === blockId);
    const pm = (candidates[0]?.blockStart ?? 0) + offset;
    let chosen: PaintedLine | null = null;
    for (const line of candidates) {
      if (pm < line.spec.pmStart || pm > line.spec.pmStart + line.spec.count) continue;
      chosen = line;
      break; // הראשונה שמכילה אותו — כלומר התפר שייך לקודמת, כפי שנמדד
    }
    if (!chosen) return;
    /*
     * שורה ריקה: אין בה תו, ולכן אין חריץ — הסמן מצויר על קצה ההתחלה שלה.
     * בלי המקרה הזה הדמה **אינה מזיזה** את הסמן אל פסקה ריקה, וכל בדיקה שעוברת
     * דרך אחת כזאת בודקת פחות ממה שהיא מצהירה: מוטציה שהחזירה `forget()` למסלול
     * המסירה שרדה כאן בשקט, מפני שהסמן נשאר במקרה על העמודה המקורית.
     */
    caretPos = chosen.spec.empty
      ? { x: chosen.spec.rtl === false ? chosen.spec.left : chosen.spec.right, top: chosen.top }
      : { x: slotX(chosen.spec, pm), top: chosen.top };
  };

  const story = { kind: 'story', storyType: 'body' };
  let seen = 0;
  const setSelectionTarget = vi.fn(
    (input: { target: { end: { blockId?: string; offset?: number } } }) => {
      current = {
        blockId: input.target.end.blockId ?? current.blockId,
        offset: input.target.end.offset ?? current.offset,
      };
      repaint(current.blockId, current.offset);
    },
  );

  const superdoc: VerticalCaretHost = {
    activeEditor: {
      host: {
        readLiveSelectionSyncSnapshot: () => {
          seen += 1;
          const point = { kind: 'text', blockId: current.blockId, offset: current.offset, story };
          return { selectionTarget: { kind: 'selection', start: point, end: point, story } };
        },
      },
      authoring: { setSelectionTarget },
    },
  };

  return {
    host,
    source: source as unknown as HTMLElement,
    superdoc,
    setSelectionTarget,
    seen: () => seen,
    caret: () => ({ ...current, x: caretPos.x }),
    lineEls: painted.map((line) => line.element),
    moveCaret: (blockId, offset) => {
      current = { blockId, offset };
      repaint(blockId, offset);
    },
    scrollHost: (dx) => {
      shift += dx;
      host.scrollLeft = shift;
    },
  };
}

const press = (target: HTMLElement, key: string, init: KeyboardEventInit = {}): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...init });
  target.dispatchEvent(event);
  return event;
};

/** מה שנכתב למנוע בקריאה ה-n. */
const written = (spy: ReturnType<typeof vi.fn>, call = 0) =>
  spy.mock.calls[call]![0] as {
    target: { start: { offset: number }; end: { offset: number; blockId: string } };
  };

/* הגאומטריה שנמדדה. */
const LONG = { pmStart: 1, count: 70, left: 548.2, right: 1043.4 };
const SHORT = { pmStart: 72, count: 9, left: 980.1, right: 1043.4 };
const CENTERED = { pmStart: 72, count: 13, left: 699.8, right: 785.2 };
const LTR_SHORT = { pmStart: 72, count: 10, left: 441.6, right: 504.3, rtl: false as const };

/**
 * ארוכה → קצרה → ארוכה, והסמן בעמודה 700: מחוץ לשורה הקצרה, ובתוך שתי
 * הארוכות. זה הרצף שבו נמדדה שמירת עמודת המטרה, והוא חוזר בכל בדיקה שצריכה
 * הקשה שנייה.
 */
const threeParagraphs = () =>
  setup({
    paras: [
      { id: 'p1', lines: [LONG] },
      { id: 'p2', lines: [SHORT] },
      { id: 'p3', lines: [{ ...LONG, pmStart: 82 }] },
    ],
    caretOffset: 40,
    caretX: 700,
  });

describe('installRtlVerticalArrows — התקלה שדווחה', () => {
  it('חץ למטה לשורה קצרה: הסמן נקבע ל**סוף** השורה, ולא לתחילתה', () => {
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [SHORT] },
      ],
      caretOffset: 40,
      // עמודה עמוק בתוך השורה הארוכה, משמאל לסוף השורה הקצרה.
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    const event = press(state.source, 'ArrowDown');

    expect(event.defaultPrevented, 'המנוע לא יראה את ההקשה').toBe(true);
    expect(state.setSelectionTarget).toHaveBeenCalledTimes(1);
    expect(written(state.setSelectionTarget).target.end.blockId).toBe('p2');
    // 9 הוא היסט הסיום שנמדד לשורה הזאת. המנוע נתן 0.
    expect(written(state.setSelectionTarget).target.end.offset).toBe(9);
    expect(written(state.setSelectionTarget).target.start.offset, 'בחירה מכווצת').toBe(9);
    handle.dispose();
  });

  it('ה-story וה-coordinateSpace של התצלום נכתבים בחזרה', () => {
    // ברירת מחדל מומצאת ל-`story` הייתה נכתבת למנוע דווקא במצב שבו הוא עצמו
    // לא דיווח אחד (הערת שוליים? כותרת?) — אותה הכרעה כמו ב-`rtl-caret.ts`.
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [SHORT] },
      ],
      caretOffset: 40,
      caretX: 700,
    });
    const editor = state.superdoc.activeEditor!;
    const inner = editor.host!.readLiveSelectionSyncSnapshot!;
    const story = { kind: 'story', storyType: 'footnote' };
    editor.host!.readLiveSelectionSyncSnapshot = () => {
      const snapshot = inner()!;
      return {
        selectionTarget: { ...snapshot.selectionTarget!, story, coordinateSpace: 'page' },
      };
    };
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');

    const input = state.setSelectionTarget.mock.calls[0]![0] as {
      target: { story: unknown; coordinateSpace?: string; end: { story: unknown } };
      focus?: boolean;
    };
    expect(input.target.story).toBe(story);
    expect(input.target.end.story).toBe(story);
    expect(input.target.coordinateSpace).toBe('page');
    expect(input.focus).toBe(true);
    handle.dispose();
  });

  it('חץ למעלה לשורה קצרה — אותו תיקון', () => {
    const state = setup({
      paras: [
        { id: 'p0', lines: [{ ...SHORT, pmStart: 1 }] },
        { id: 'p1', lines: [{ ...LONG, pmStart: 11 }] },
      ],
      sourcePara: 1,
      caretOffset: 40,
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    const event = press(state.source, 'ArrowUp');

    expect(event.defaultPrevented).toBe(true);
    expect(written(state.setSelectionTarget).target.end.blockId).toBe('p0');
    expect(written(state.setSelectionTarget).target.end.offset).toBe(9);
    handle.dispose();
  });

  it('שורה ממורכזת, עמודה מימין לתחילתה: הסמן נקבע ל**תחילת** השורה', () => {
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [CENTERED] },
      ],
      caretOffset: 5,
      // עמודה בקצה הימני של השורה הארוכה — מימין לקצה הימני של הממורכזת.
      caretX: 1027.2,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');

    // תחילת השורה. המנוע נתן את היסט הסיום, 13.
    expect(written(state.setSelectionTarget).target.end.offset).toBe(0);
    handle.dispose();
  });

  it('שורה גולשת בתוך אותה פסקה — אותה תקלה, בלי פסקה שנייה', () => {
    const state = setup({
      paras: [
        {
          id: 'p1',
          lines: [
            { pmStart: 1, count: 70, left: 454.5, right: 1043.4 },
            { pmStart: 71, count: 40, left: 666.1, right: 1043.4 },
          ],
        },
      ],
      caretOffset: 40,
      caretX: 500,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');

    // סוף השורה השנייה: pm 111, פחות תחילת הבלוק 1.
    expect(written(state.setSelectionTarget).target.end.offset).toBe(110);
    handle.dispose();
  });

  it('תפר הגלישה: ירידה אל תחילת שורת המשך נוחתת חריץ אחד פנימה', () => {
    // שורת המשך שקצה ההתחלה שלה נסוג פנימה — רק אז עמודת המטרה יכולה ליפול
    // מימין לתחילתה. בפסקה רגילה זה בלתי אפשרי, ולכן גם התפר אינו עולה.
    const state = setup({
      paras: [
        {
          id: 'p1',
          lines: [
            { pmStart: 1, count: 70, left: 454.5, right: 1043.4 },
            { pmStart: 71, count: 20, left: 666.1, right: 800 },
          ],
        },
      ],
      caretOffset: 5,
      caretX: 1030,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');

    // תחילת שורת ההמשך היא pm 71 — אותו היסט כמו סוף השורה הראשונה, וכתיבתו
    // הייתה מציירת את הסמן שם. לכן היעד הוא pm 72, כלומר היסט 71.
    expect(written(state.setSelectionTarget).target.end.offset).toBe(71);
    handle.dispose();
  });
});

describe('installRtlVerticalArrows — מה שנמסר למנוע', () => {
  /**
   * „נמסר” ולא „המודול מת”.
   *
   * בדיקה שמסתפקת ב-`defaultPrevented === false` ובמרגל שלא נקרא עוברת ירוק
   * גם על התקנה שאינה עושה דבר. לכן כל אחת מאלה קובעת גם ש**היירוט רץ והגיע
   * להכרעה** — כלומר שהתצלום נקרא.
   */
  const expectHandedOver = (event: KeyboardEvent, state: Setup): void => {
    expect(event.defaultPrevented, 'ההקשה ממשיכה למנוע').toBe(false);
    expect(state.setSelectionTarget, 'שום בחירה לא נכתבה').not.toHaveBeenCalled();
    expect(state.seen(), 'היירוט רץ והגיע להכרעה').toBeGreaterThan(0);
  };

  it('עמודה שנופלת בתוך שורת היעד — המנוע נמדד תקין', () => {
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [SHORT] },
      ],
      caretOffset: 5,
      caretX: 1021.6,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    expectHandedOver(press(state.source, 'ArrowDown'), state);
    handle.dispose();
  });

  it('סמן המספור של פריט רשימה אינו מרחיב את השורה', () => {
    // הסמן מצויר בשורה אבל אין לו טווח pm, כלומר אין בו מקום סמן. אילו הוא
    // נספר בקצה, עמודה שנופלת עליו הייתה נחשבת „בתוך השורה” והייתה נמסרת
    // למנוע — כלומר בדיוק המקרה השבור, במסווה.
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [SHORT] },
      ],
      caretOffset: 40,
      caretX: 700,
    });
    // טאב-הסיומת של סמן מספור: `PM_CARRIER_SELECTOR` תופס אותו, ואין לו טווח.
    // הוא נמתח שמאלה מהשורה הקצרה, בדיוק אל העמודה שהסמן יוצא ממנה.
    const marker = document.createElement('span');
    marker.className = 'superdoc-tab superdoc-marker-suffix-tab';
    marker.getBoundingClientRect = () => ({ left: 600, right: 980.1, width: 380.1 }) as DOMRect;
    state.lineEls[1]!.append(marker);

    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });
    const event = press(state.source, 'ArrowDown');

    expect(event.defaultPrevented, 'הסמן עדיין מחוץ לשורה — ההקשה שלנו').toBe(true);
    expect(written(state.setSelectionTarget).target.end.offset).toBe(9);
    handle.dispose();
  });

  it('שורת יעד לטינית — המנוע נמדד תקין בכל חמש הבקרות', () => {
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [LTR_SHORT] },
      ],
      caretOffset: 40,
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    expectHandedOver(press(state.source, 'ArrowDown'), state);
    handle.dispose();
  });

  it('שורת יעד ריקה — מקום סמן אחד, ואין מה לבחור', () => {
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [{ pmStart: 72, count: 0, left: 1043.4, right: 1043.4, empty: true }] },
      ],
      caretOffset: 40,
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    expectHandedOver(press(state.source, 'ArrowDown'), state);
    handle.dispose();
  });

  it('טבלה שכנה — fragment בלי מזהה מקור, והמנוע נכנס לתא כראוי', () => {
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: null, lines: [SHORT] },
      ],
      caretOffset: 40,
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    expectHandedOver(press(state.source, 'ArrowDown'), state);
    handle.dispose();
  });

  it('בלוק שאינו צמוד במסמך — יש משהו ביניהם שלא צויר', () => {
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        // פער של 8 יחידות pm: לא 1 ולא 2.
        { id: 'p2', lines: [{ ...SHORT, pmStart: 80 }] },
      ],
      caretOffset: 40,
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    expectHandedOver(press(state.source, 'ArrowDown'), state);
    handle.dispose();
  });

  it('ArrowDown בשורה האחרונה במסמך עברי מעביר לסופה, והקשה נוספת אינה כותבת שוב', () => {
    const state = setup({
      paras: [{ id: 'p1', lines: [LONG] }],
      caretOffset: 40,
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    const first = press(state.source, 'ArrowDown');
    expect(first.defaultPrevented).toBe(true);
    expect(state.caret()).toMatchObject({ blockId: 'p1', offset: 70 });
    expect(state.setSelectionTarget).toHaveBeenCalledTimes(1);

    const second = press(state.source, 'ArrowDown');
    expect(second.defaultPrevented).toBe(true);
    expect(state.setSelectionTarget, 'כבר בסוף השורה — לא שולחים כתיבה מיותרת').toHaveBeenCalledTimes(1);
    handle.dispose();
  });

  it('ArrowDown בסוף המסמך עובד גם בשורה לטינית', () => {
    const state = setup({
      paras: [{ id: 'p1', lines: [{ ...LTR_SHORT, pmStart: 1 }] }],
      caretOffset: 3,
      caretX: 460,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    expect(press(state.source, 'ArrowDown').defaultPrevented).toBe(true);
    expect(written(state.setSelectionTarget).target.end.offset).toBe(10);
    handle.dispose();
  });

  it('הסוף הוא סוף השורה החזותית האחרונה, גם בפסקה שנשברת לשתי שורות', () => {
    const state = setup({
      paras: [{
        id: 'p1',
        lines: [LONG, { ...SHORT, pmStart: 71 }],
      }],
      sourceLine: 1,
      caretOffset: 74,
      caretX: 1000,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    expect(press(state.source, 'ArrowDown').defaultPrevented).toBe(true);
    // pm 80 (סוף השורה השנייה) פחות בסיס הבלוק 1.
    expect(state.caret().offset).toBe(79);
    handle.dispose();
  });

  it('חץ למעלה בקצה המסמך נשאר למנוע', () => {
    const state = setup({
      paras: [{ id: 'p1', lines: [LONG] }],
      caretOffset: 40,
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    expectHandedOver(press(state.source, 'ArrowUp'), state);
    handle.dispose();
  });

  it('סוף עמוד מצויר אינו סוף המסמך כשקיים עמוד נוסף', () => {
    const state = setup({
      paras: [{ id: 'p1', lines: [LONG] }],
      caretOffset: 40,
      caretX: 700,
    });
    const nextPage = document.createElement('div');
    nextPage.className = 'superdoc-page';
    state.host.append(nextPage);
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    expectHandedOver(press(state.source, 'ArrowDown'), state);
    handle.dispose();
  });

  it('fragment שממשיך בעמוד הבא אינו קצה מסמך גם אם העמוד הבא טרם צויר', () => {
    const state = setup({
      paras: [{ id: 'p1', lines: [LONG] }],
      caretOffset: 40,
      caretX: 700,
    });
    state.source.parentElement!.setAttribute('data-continues-on-next', 'true');
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    expectHandedOver(press(state.source, 'ArrowDown'), state);
    handle.dispose();
  });

  it('אין סמן מצויר — אין עמודה, ואין ממה לגזור', () => {
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [SHORT] },
      ],
      caretOffset: 40,
      caretX: 700,
      withCaret: false,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    expectHandedOver(press(state.source, 'ArrowDown'), state);
    handle.dispose();
  });

  it('Shift+חץ נשאר למנוע', () => {
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [SHORT] },
      ],
      caretOffset: 40,
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    const event = press(state.source, 'ArrowDown', { shiftKey: true });

    expect(event.defaultPrevented).toBe(false);
    expect(state.setSelectionTarget).not.toHaveBeenCalled();
    // כאן `seen()` הוא 0 בכוונה: המקש נדחה לפני קריאת התצלום. מה שמוכיח
    // שההתקנה חיה הוא שהקשה נקייה באותה התקנה כן נבלעת.
    const clean = press(state.source, 'ArrowDown');
    expect(clean.defaultPrevented, 'הקשה נקייה באותה התקנה נבלעת — המודול חי').toBe(true);
    handle.dispose();
  });

  it('בחירה שאינה מכווצת — התצלום מחזיר null, ואין קריסה', () => {
    // זה מה שהמנוע מחזיר לכל בחירה שאינה מכווצת, ולכן זה מה שקורה בהקשה
    // האנכית הראשונה שאחרי בחירה בגרירה.
    const state = threeParagraphs();
    const editor = state.superdoc.activeEditor!;
    editor.host!.readLiveSelectionSyncSnapshot = () => ({ selectionTarget: null });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    const event = press(state.source, 'ArrowDown');

    expect(event.defaultPrevented).toBe(false);
    expect(state.setSelectionTarget).not.toHaveBeenCalled();
    handle.dispose();
  });

  it('קצה בחירה שאינו טקסט, או בלי היסט — נמסר ולא קורס', () => {
    for (const head of [
      { kind: 'table', blockId: 'p1', offset: 40 },
      { kind: 'text', offset: 40 },
      { kind: 'text', blockId: 'p1' },
    ]) {
      const state = threeParagraphs();
      const editor = state.superdoc.activeEditor!;
      editor.host!.readLiveSelectionSyncSnapshot = () => ({
        selectionTarget: { kind: 'selection', start: head, end: head, story: {} },
      });
      const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

      expect(press(state.source, 'ArrowDown').defaultPrevented).toBe(false);
      expect(state.setSelectionTarget).not.toHaveBeenCalled();
      handle.dispose();
      document.body.innerHTML = '';
    }
  });

  it('אין מסמך — אין מה לקרוא, ואין קריסה', () => {
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [SHORT] },
      ],
      caretOffset: 40,
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: { activeEditor: null } });

    const event = press(state.source, 'ArrowDown');

    expect(event.defaultPrevented).toBe(false);
    handle.dispose();
  });
});

describe('installRtlVerticalArrows — עמודת המטרה', () => {
  /**
   * המדידה שמכתיבה את ההתנהגות הזאת: הכתיבה שלנו **דורסת** את עמודת המטרה של
   * המנוע. לכן מרגע שהתערבנו, כל הקשה אנכית היא שלנו — גם כזו שהמנוע היה
   * נוחת בה נכון — והיעד נגזר מהעמודה **המקורית**, ולא מהמקום שכתבנו.
   *
   * שלוש הפסקאות: ארוכה, קצרה, ארוכה. הסמן מתחיל בעמודה 700 — מחוץ לשורה
   * הקצרה, ובתוך שתי הארוכות.
   */

  it('אחרי התערבות, ההקשה השנייה שלנו — ונגזרת מהעמודה המקורית', () => {
    const state = threeParagraphs();
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');
    expect(state.caret(), 'הסמן ירד לסוף השורה הקצרה').toMatchObject({
      blockId: 'p2',
      offset: 9,
    });

    const second = press(state.source, 'ArrowDown');

    expect(second.defaultPrevented, 'גם ההקשה הזאת שלנו — העמודה של המנוע נדרסה').toBe(true);
    expect(state.setSelectionTarget).toHaveBeenCalledTimes(2);
    expect(written(state.setSelectionTarget, 1).target.end.blockId).toBe('p3');
    // העמודה המקורית, 700, נופלת על החריץ ה-49 של השורה הארוכה. אילו העמודה
    // הייתה נזרעת מחדש מהסמן (שיושב על 980.1) היה יוצא חריץ אחר לגמרי.
    expect(written(state.setSelectionTarget, 1).target.end.offset).toBe(49);
    handle.dispose();
  });

  it('מקש שאינו חץ אנכי מאפס את העמודה — וההקשה הבאה חוזרת למנוע', () => {
    const state = threeParagraphs();
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');
    expect(state.setSelectionTarget).toHaveBeenCalledTimes(1);

    press(state.source, 'a');
    const after = press(state.source, 'ArrowDown');

    // העמודה נזרעה מחדש מהסמן, שיושב על 980.1 — בתוך השורה הארוכה שמתחת,
    // ושם המנוע נמדד תקין.
    expect(after.defaultPrevented, 'נמסר למנוע').toBe(false);
    expect(state.setSelectionTarget, 'ולא נכתב דבר נוסף').toHaveBeenCalledTimes(1);
    handle.dispose();
  });

  it('לחיצת עכבר מאפסת את העמודה', () => {
    const state = threeParagraphs();
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');
    state.host.dispatchEvent(new Event('pointerdown', { bubbles: true }));

    const after = press(state.source, 'ArrowDown');

    expect(after.defaultPrevented).toBe(false);
    expect(state.setSelectionTarget).toHaveBeenCalledTimes(1);
    handle.dispose();
  });

  it('Shift+חץ מסיים את התנועה האנכית, כמו כל מקש אחר', () => {
    const state = threeParagraphs();
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');
    press(state.source, 'ArrowDown', { shiftKey: true });
    const after = press(state.source, 'ArrowDown');

    expect(after.defaultPrevented).toBe(false);
    expect(state.setSelectionTarget).toHaveBeenCalledTimes(1);
    handle.dispose();
  });

  it('מודול שבולע על ה-document אינו מסתיר מאיתנו מקש', () => {
    // זה בדיוק `list-autoformat-install.ts`: מאזין `keydown` בשלב capture על
    // ה-`document`, שבולע `Ctrl+Z`/`Ctrl+Y` ורווח ב-`stopImmediatePropagation`.
    // שניהם מזיזים את הסמן, ולכן חייבים לאפס את העמודה. איפוס שיושב על
    // ה-`document` היה תלוי בשאלה מי נרשם קודם; על ה-`window` אין שאלה כזאת.
    // בולע בדיוק את מה ש-list-autoformat בולע: את הצירוף עם Ctrl, ולא הכול.
    const swallow = (event: Event): void => {
      if ((event as KeyboardEvent).ctrlKey) event.stopImmediatePropagation();
    };
    document.addEventListener('keydown', swallow, true);
    const state = threeParagraphs();
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });
    try {
      press(state.source, 'ArrowDown');
      expect(state.setSelectionTarget, 'הירידה הראשונה שלנו').toHaveBeenCalledTimes(1);

      press(state.source, 'z', { ctrlKey: true });
      const after = press(state.source, 'ArrowDown');

      // העמודה נזרעה מחדש מהסמן שעל 980.1, שבתוך השורה הארוכה שמתחת.
      expect(after.defaultPrevented, 'נמסר למנוע אחרי האיפוס').toBe(false);
      expect(state.setSelectionTarget).toHaveBeenCalledTimes(1);
    } finally {
      document.removeEventListener('keydown', swallow, true);
      handle.dispose();
    }
  });

  it('מודול אחר בולע את החץ האופקי — והעמודה מתאפסת בכל זאת', () => {
    // זה בדיוק מה ש-`rtl-caret.ts` עושה לחץ אופקי: מאזין על אותו host שעוצר
    // את האירוע ב-`stopImmediatePropagation`. אילו האיפוס היה יושב על ה-host,
    // הוא היה תלוי בשאלה מי נרשם קודם.
    const state = threeParagraphs();
    // נרשם **לפני** המודול, כמו `rtl-caret.ts` ב-App.vue, ובולע רק אופקי.
    state.host.addEventListener(
      'keydown',
      (event) => {
        if (event.key === 'ArrowRight') event.stopImmediatePropagation();
      },
      true,
    );
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');
    expect(state.setSelectionTarget, 'הירידה הראשונה שלנו').toHaveBeenCalledTimes(1);

    // החץ האופקי נבלע — המאזין שלנו על ה-host אינו רואה אותו כלל.
    press(state.source, 'ArrowRight');

    // ואף על פי כן העמודה התאפסה: הסמן יושב עכשיו על 980.1, שנמצא בתוך השורה
    // הארוכה שמתחת, ולכן ההקשה נמסרת למנוע. אילו העמודה נשארה 700, המודול היה
    // כותב שוב.
    const after = press(state.source, 'ArrowDown');
    expect(after.defaultPrevented, 'נמסר למנוע אחרי האיפוס').toBe(false);
    expect(state.setSelectionTarget).toHaveBeenCalledTimes(1);
    handle.dispose();
  });

  it('מודול אחר הזיז את הסמן — העמודה נזרעת מחדש, בלי לראות את המקש', () => {
    // זה בדיוק מה שקורה בחץ אופקי או ב-`End`: `rtl-caret.ts` ו-
    // `rtl-line-end.ts` בולעים אותם ב-`stopImmediatePropagation`, ולכן המודול
    // הזה **אינו רואה** את ההקשה. מה שהוא כן רואה הוא שהסמן אינו היכן שהשאיר
    // אותו.
    const state = threeParagraphs();
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');
    expect(state.setSelectionTarget).toHaveBeenCalledTimes(1);

    // מודול אחר מזיז את הסמן — אותו מסלול בדיוק, בלי שאף מקש יגיע אלינו.
    state.setSelectionTarget({ target: { end: { blockId: 'p2', offset: 4 } } });
    state.setSelectionTarget.mockClear();

    const after = press(state.source, 'ArrowDown');

    // העמודה נזרעה מחדש מהסמן, שיושב עכשיו בתוך השורה הקצרה — ומשם השורה
    // הארוכה שמתחת מכילה את העמודה, כלומר המנוע צודק וההקשה נמסרת לו.
    expect(after.defaultPrevented, 'נמסר למנוע אחרי הזריעה מחדש').toBe(false);
    expect(state.setSelectionTarget).not.toHaveBeenCalled();
    handle.dispose();
  });

  it('מסירה למנוע אינה מאפסת — פסקה ריקה באמצע אינה שוברת את העמודה', () => {
    // הרצף „ארוכה → ריקה → קצרה” שגרתי במסמך עברי, והוא זה שחושף למה מסירה
    // אסור לה לאפס: המנוע **לא** נדרס, ולכן העמודה שלו שרדה. אילו שלנו הייתה
    // מתאפסת, ההקשה הבאה הייתה נזרעת מהסמן שעל הפסקה הריקה — קצה ההתחלה —
    // ומכריזה בטעות „בתוך השורה הקצרה”, מוסרת, והמנוע היה נוחת בקצה ההפוך.
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [{ pmStart: 72, count: 0, left: 1043.4, right: 1043.4, empty: true }] },
        { id: 'p3', lines: [{ ...SHORT, pmStart: 74 }] },
      ],
      caretOffset: 40,
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    const first = press(state.source, 'ArrowDown');
    expect(first.defaultPrevented, 'שורה ריקה נמסרת למנוע').toBe(false);

    // המנוע הזיז את הסמן לפסקה הריקה — קצה ההתחלה של שורה ריקה.
    state.moveCaret('p2', 0);

    const second = press(state.source, 'ArrowDown');

    expect(second.defaultPrevented, 'העמודה שרדה, ולכן ההקשה הזאת שלנו').toBe(true);
    expect(written(state.setSelectionTarget).target.end.blockId).toBe('p3');
    expect(written(state.setSelectionTarget).target.end.offset).toBe(9);
    handle.dispose();
  });

  it('מסירה אחרי התערבות שומרת את העמודה דרך פסקה ריקה', () => {
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [SHORT] },
        { id: 'p3', lines: [{ pmStart: 82, count: 0, left: 1043.4, right: 1043.4, empty: true }] },
        { id: 'p4', lines: [{ ...SHORT, pmStart: 84 }] },
      ],
      caretOffset: 40,
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');
    expect(state.caret()).toMatchObject({ blockId: 'p2', offset: 9 });

    const throughEmpty = press(state.source, 'ArrowDown');
    expect(throughEmpty.defaultPrevented, 'החץ אל הפסקה הריקה נמסר למנוע').toBe(false);
    state.moveCaret('p3', 0);

    const afterEmpty = press(state.source, 'ArrowDown');

    expect(afterEmpty.defaultPrevented, 'המודול שומר את עמודת המטרה המקורית').toBe(true);
    expect(written(state.setSelectionTarget, 1).target.end.blockId).toBe('p4');
    expect(written(state.setSelectionTarget, 1).target.end.offset).toBe(9);
    handle.dispose();
  });

  it('גלילה אופקית של ה-host אינה מזיזה את העמודה', () => {
    // `caret-visibility.ts` יושב על אותו host וגולל אופקית כשהסמן יוצא
    // מהתצוגה — כלומר הכתיבה שלנו עצמה מייצרת את הגלילה. עמודה שנשמרה
    // בקואורדינטות חלון הייתה מצביעה אחרי זה למקום אחר לגמרי.
    const state = threeParagraphs();
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');
    state.scrollHost(200);

    press(state.source, 'ArrowDown');

    expect(state.setSelectionTarget).toHaveBeenCalledTimes(2);
    // אותו חריץ בדיוק כמו בלי הגלילה. בקואורדינטות חלון היה יוצא 20.
    expect(written(state.setSelectionTarget, 1).target.end.offset).toBe(49);
    handle.dispose();
  });

  it('חץ למטה ואז חץ למעלה — חוזרים לעמודה, לא למקום', () => {
    const state = threeParagraphs();
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');
    const back = press(state.source, 'ArrowUp');

    expect(back.defaultPrevented).toBe(true);
    expect(written(state.setSelectionTarget, 1).target.end.blockId).toBe('p1');
    // העמודה המקורית, 700 — ולא ההיסט שהתחלנו ממנו (40).
    expect(written(state.setSelectionTarget, 1).target.end.offset).toBe(49);
    handle.dispose();
  });

  it('אחרי התערבות, גם שורת יעד לטינית היא שלנו', () => {
    // התנאי הזול מוסר למנוע שורה לטינית — אבל רק כל עוד לא כתבנו. מרגע
    // שכתבנו, העמודה שלו היא מה שכתבנו, ולטינית אינה יוצאת מן הכלל.
    const state = setup({
      paras: [
        { id: 'p1', lines: [LONG] },
        { id: 'p2', lines: [SHORT] },
        { id: 'p3', lines: [{ pmStart: 82, count: 70, left: 441.6, right: 936.8, rtl: false }] },
      ],
      caretOffset: 40,
      caretX: 700,
    });
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });

    press(state.source, 'ArrowDown');
    const second = press(state.source, 'ArrowDown');

    expect(second.defaultPrevented, 'לא נמסר למנוע, שעמודתו נדרסה').toBe(true);
    expect(state.setSelectionTarget).toHaveBeenCalledTimes(2);
    expect(written(state.setSelectionTarget, 1).target.end.blockId).toBe('p3');
    handle.dispose();
  });

  it('dispose מסיר את שלושת המאזינים — כולל השניים שעל החלון', () => {
    // שניים מהם אינם על ה-host, ולכן „ההקשה כבר לא נבלעת” אינו מוכיח עליהם
    // דבר: מסמך שנסגר היה מדליף שני מאזיני capture שמחזיקים סגור על host
    // ו-superdoc מתים, אחד לכל פתיחת מסמך.
    const removals: string[] = [];
    const original = window.removeEventListener.bind(window);
    vi.spyOn(window, 'removeEventListener').mockImplementation(((
      type: string,
      listener: EventListenerOrEventListenerObject,
      options?: boolean | EventListenerOptions,
    ) => {
      removals.push(type);
      return original(type, listener, options);
    }) as typeof window.removeEventListener);

    const state = threeParagraphs();
    const handle = installRtlVerticalArrows({ host: state.host, superdoc: state.superdoc });
    handle.dispose();

    expect(removals).toContain('keydown');
    expect(removals).toContain('pointerdown');

    const event = press(state.source, 'ArrowDown');
    expect(event.defaultPrevented).toBe(false);
    expect(state.seen()).toBe(0);
  });

});
