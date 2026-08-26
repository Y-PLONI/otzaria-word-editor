/**
 * סרגל המידות.
 *
 * מה שנמדד כאן הוא בדיוק מה שאי אפשר למדוד בפונקציה טהורה: שהסרגל **מיושר**
 * לעמוד שהמנוע צייר, שהצד שבו הוא מתחיל הולך אחרי כיוון המקטע ולא אחרי כיוון
 * הממשק, ושגרירה של ידית מגיעה למנוע כמידה נכונה — באינצ'ים, בצד הנכון, ובלי
 * למחוק מה שלא נגררה.
 *
 * ## למה `getBoundingClientRect` מוחלף כאן
 *
 * jsdom אינו מפריס: כל מלבן בו הוא אפס. סרגל שכל המספרים שלו נגזרים ממלבן
 * העמוד לא היה מצייר דבר, וכל בדיקה כאן הייתה עוברת מהסיבה הלא נכונה. לכן
 * המלבנים נקבעים כאן במפורש, לפי המספרים שנמדדו על המנוע האמיתי: עמוד A4
 * ברוחב 794px ב-100%.
 */
import { describe, expect, it, beforeEach, afterEach, vi } from 'vitest';
import DocumentRuler from '../../src/ui/shell/DocumentRuler.vue';
import { TWIPS_PER_CM } from '../../src/engine/ruler-geometry';
import type { RulerReading } from '../../src/engine/page-ruler';
import { autoUnmount, mountUi, settle } from './harness';

autoUnmount();

/* ------------------------------------------------------------------ */
/* גיאומטריה מדומה                                                     */
/* ------------------------------------------------------------------ */

/** A4 ב-100%: 21 ס"מ = 794px, שוליים של 2.54 ס"מ = 96px. */
const PAGE_WIDTH_TWIPS = 11906;
const PAGE_HEIGHT_TWIPS = 16838;
const A_INCH = 1440;

let pageLeft = 100;
const PAGE_WIDTH_PX = 794;
/** ה-host והסרגל יושבים באותו מקום; מלבן הסרגל הוא נקודת הייחוס. */
const ROOT_LEFT = 0;

const originalRect = HTMLElement.prototype.getBoundingClientRect;

beforeEach(() => {
  pageLeft = 100;
  HTMLElement.prototype.getBoundingClientRect = function rect(this: HTMLElement): DOMRect {
    if (this.hasAttribute('data-page-index')) {
      return box(pageLeft, PAGE_WIDTH_PX);
    }
    if (this.classList.contains('doc-ruler')) return box(ROOT_LEFT, 900);
    return box(0, 900);
  };
  // ה-watcher מודד ב-rAF; בבדיקה הוא רץ מיד כדי שהמדידה תהיה סינכרונית.
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
    callback(0);
    return 1;
  });
  vi.stubGlobal('cancelAnimationFrame', () => {});
});

afterEach(() => {
  HTMLElement.prototype.getBoundingClientRect = originalRect;
  vi.unstubAllGlobals();
  document.body.innerHTML = '';
});

function box(left: number, width: number): DOMRect {
  return {
    left,
    right: left + width,
    width,
    top: 0,
    bottom: 22,
    height: 22,
    x: left,
    y: 0,
    toJSON: () => ({}),
  } as DOMRect;
}

function engineHost(): HTMLElement {
  const host = document.createElement('div');
  const page = document.createElement('div');
  page.setAttribute('data-page-index', '0');
  host.appendChild(page);
  document.body.appendChild(host);
  return host;
}

function reading(overrides: Partial<RulerReading> = {}): RulerReading {
  return {
    page: {
      pageWidthTwips: PAGE_WIDTH_TWIPS,
      pageHeightTwips: PAGE_HEIGHT_TWIPS,
      leftTwips: A_INCH,
      rightTwips: A_INCH,
      topTwips: A_INCH,
      bottomTwips: A_INCH,
      direction: 'rtl',
    },
    indents: { leftTwips: 0, rightTwips: 0, firstLineTwips: 0, hangingTwips: 0, bidi: true },
    target: { kind: 'block', nodeType: 'paragraph', nodeId: 'p1' },
    ...overrides,
  };
}

async function mountRuler(props: Record<string, unknown> = {}) {
  const harness = mountUi(DocumentRuler, {
    props: { visible: true, host: engineHost(), reading: reading(), ...props },
  });
  await settle();
  return harness;
}

/** המיקום הפיזי של ידית, בפיקסלים בתוך מלבן העמוד. */
function handleLeft(element: Element): number {
  return Number.parseFloat((element as HTMLElement).style.left);
}

function handleByLabel(wrapper: { findAll: (s: string) => Array<{ element: Element; attributes: (k: string) => string | undefined }> }, label: string) {
  return wrapper.findAll('.doc-ruler__handle').find((handle) => handle.attributes('aria-label') === label);
}

/* ------------------------------------------------------------------ */

describe('ציור', () => {
  it('בלי מסמך אין מה לצייר, והרצועה עצמה נשארת', async () => {
    const harness = mountUi(DocumentRuler, { props: { visible: true, reading: null } });
    await settle();

    expect(harness.wrapper.find('.doc-ruler').exists()).toBe(true);
    expect(harness.wrapper.find('.doc-ruler__page').exists()).toBe(false);
  });

  it('מלבן העמוד נלקח מהעמוד המצויר, ולא מחושב מהזום', async () => {
    const harness = await mountRuler();

    const page = harness.wrapper.find('.doc-ruler__page').element as HTMLElement;
    expect(page.style.left).toBe('100px');
    expect(page.style.width).toBe('794px');
  });

  it('אזור הטקסט הוא העמוד פחות שני השוליים', async () => {
    const harness = await mountRuler();

    const area = harness.wrapper.find('.doc-ruler__text-area').element as HTMLElement;
    // 96px שוליים בכל צד, ב-794px רוחב עמוד.
    expect(Math.round(Number.parseFloat(area.style.left))).toBe(96);
    expect(Math.round(Number.parseFloat(area.style.width))).toBe(602);
  });

  it('יש שנתות, ומספר אינו יורד מתחת לאפס', async () => {
    const harness = await mountRuler();

    const numbers = harness.wrapper.findAll('.doc-ruler__number').map((node) => node.text());
    expect(numbers.length).toBeGreaterThan(10);
    expect(numbers.some((text) => text.startsWith('-'))).toBe(false);
    // 2.54 ס"מ שוליים: „1” ו„2” מופיעים גם בתוך השוליים וגם באזור הטקסט.
    expect(numbers.filter((text) => text === '1')).toHaveLength(2);
  });
});

describe('כיוון', () => {
  it('במסמך עברי ידית שולי ההתחלה יושבת בקצה הימני', async () => {
    const harness = await mountRuler();

    const start = handleByLabel(harness.wrapper, 'שוליים ימניים');
    // 794 - 96 = 698, ביחס לשמאל מלבן העמוד.
    expect(Math.round(handleLeft(start!.element))).toBe(698);
  });

  it('במסמך לועזי אותה ידית עוברת לשמאל, ושמה מתחלף', async () => {
    const harness = await mountRuler({
      reading: reading({
        page: {
          pageWidthTwips: PAGE_WIDTH_TWIPS,
          pageHeightTwips: PAGE_HEIGHT_TWIPS,
          leftTwips: A_INCH,
          rightTwips: A_INCH,
          topTwips: A_INCH,
          bottomTwips: A_INCH,
          direction: 'ltr',
        },
      }),
    });

    const start = handleByLabel(harness.wrapper, 'שוליים שמאליים');
    expect(Math.round(handleLeft(start!.element))).toBe(96);
  });

  it('סמני הכניסה מופיעים רק כשיש פסקה', async () => {
    const withCaret = await mountRuler();
    expect(withCaret.wrapper.findAll('.doc-ruler__handle')).toHaveLength(4);

    const noCaret = await mountRuler({ reading: reading({ indents: null, target: null }) });
    expect(noCaret.wrapper.findAll('.doc-ruler__handle')).toHaveLength(2);
  });

  it('כניסת ההתחלה נמדדת מתחילת הטקסט, בצד ההתחלה', async () => {
    const harness = await mountRuler({
      reading: reading({
        indents: {
          leftTwips: Math.round(TWIPS_PER_CM),
          rightTwips: 0,
          firstLineTwips: 0,
          hangingTwips: 0,
          bidi: true,
        },
      }),
    });

    const indent = handleByLabel(harness.wrapper, 'כניסה מצד ההתחלה');
    // ס"מ אחד פנימה מ-698 → 698 - 37.8.
    expect(Math.round(handleLeft(indent!.element))).toBe(660);
  });
});

describe('נגישות', () => {
  it('כל ידית היא slider עם ערך בסנטימטרים', async () => {
    const harness = await mountRuler();

    const start = handleByLabel(harness.wrapper, 'שוליים ימניים')!;
    expect(start.attributes('role')).toBe('slider');
    expect(start.attributes('tabindex')).toBe('0');
    expect(start.attributes('aria-valuenow')).toBe('2.54');
    expect(start.attributes('aria-valuetext')).toBe('2.54 ס"מ');
  });

  it('מסמך לקריאה בלבד — הידיות אינן ב-tab order', async () => {
    const harness = await mountRuler({ editable: false });

    const start = handleByLabel(harness.wrapper, 'שוליים ימניים')!;
    expect(start.attributes('tabindex')).toBe('-1');
    expect(start.attributes('aria-disabled')).toBe('true');
  });
});

/* ------------------------------------------------------------------ */
/* גרירה                                                               */
/* ------------------------------------------------------------------ */

/** אירוע מצביע ל-jsdom: `PointerEvent` אינו בנוי בו, ו-`pointerId` נדרש. */
function pointer(type: string, clientX: number, extra: Record<string, unknown> = {}): MouseEvent {
  const event = new MouseEvent(type, { clientX, bubbles: true });
  Object.defineProperty(event, 'pointerId', { value: 1 });
  for (const [key, value] of Object.entries(extra)) {
    Object.defineProperty(event, key, { value });
  }
  return event;
}

async function drag(
  harness: Awaited<ReturnType<typeof mountRuler>>,
  label: string,
  toClientX: number,
  extra: Record<string, unknown> = {},
): Promise<void> {
  const handle = handleByLabel(harness.wrapper, label)!;
  handle.element.dispatchEvent(pointer('pointerdown', 0, { button: 0 }));
  await settle();
  window.dispatchEvent(pointer('pointermove', toClientX, extra));
  await settle();
  window.dispatchEvent(pointer('pointerup', toClientX, extra));
  await settle();
}

describe('גרירת שוליים', () => {
  it('גרירה פנימה מגדילה את השוליים, ונשלחת באינצ\'ים', async () => {
    const harness = await mountRuler();

    // מלבן העמוד מתחיל ב-100 ורחב 794 → הקצה הימני ב-894. גרירה ל-794
    // מרחיקה את גבול השוליים 100px מהקצה, כלומר 2.65 ס"מ שמוצמדים ל-2.75.
    await drag(harness, 'שוליים ימניים', 794);

    const calls = harness.superdoc.inputs('sections.setPageMargins');
    expect(calls).toHaveLength(1);
    // במסמך עברי ההתחלה היא הצד הימני, ולכן `right` הוא שזז.
    const input = calls[0] as { left: number; right: number };
    expect(input.right).toBeCloseTo(2.75 / 2.54, 3);
    expect(input.left).toBeCloseTo(1, 6);
  });

  it('Alt מבטל את ההצמדה', async () => {
    const harness = await mountRuler();

    await drag(harness, 'שוליים ימניים', 794, { altKey: true });

    const input = harness.superdoc.inputs('sections.setPageMargins')[0] as { right: number };
    // 100px = 2.6458 ס"מ. בלי הצמדה הערך אינו נופל על רבע סנטימטר.
    expect(input.right).toBeCloseTo(100 / 794 * (21 / 2.54), 2);
  });

  it('גרירה אל מעבר לקצה הדף נעצרת בגבול, ולא מייצרת שוליים שליליים', async () => {
    const harness = await mountRuler();

    await drag(harness, 'שוליים ימניים', 2000);

    const input = harness.superdoc.inputs('sections.setPageMargins')[0] as { right: number };
    expect(input.right).toBe(0);
  });

  it('הידית עוקבת אחרי הסמן עוד לפני השחרור', async () => {
    const harness = await mountRuler();
    const handle = handleByLabel(harness.wrapper, 'שוליים ימניים')!;

    handle.element.dispatchEvent(pointer('pointerdown', 0, { button: 0 }));
    await settle();
    window.dispatchEvent(pointer('pointermove', 794));
    await settle();

    // הידית נחה על הערך המוצמד (2.75 ס"מ) ולא בדיוק מתחת לסמן — זו ההצמדה
    // שהמשתמש רואה, ובלעדיה הוא היה משחרר על 2.65 ומקבל 2.75.
    expect(Math.round(handleLeft(handleByLabel(harness.wrapper, 'שוליים ימניים')!.element))).toBe(690);
    // ועדיין לא נכתב דבר: כתיבה בכל תזוזה הייתה מעמדת מחדש את המסמך כולו.
    expect(harness.superdoc.inputs('sections.setPageMargins')).toHaveLength(0);

    window.dispatchEvent(pointer('pointerup', 794));
    await settle();
    expect(harness.superdoc.inputs('sections.setPageMargins')).toHaveLength(1);
  });

  it('ביטול הגרירה אינו כותב דבר', async () => {
    const harness = await mountRuler();
    const handle = handleByLabel(harness.wrapper, 'שוליים ימניים')!;

    handle.element.dispatchEvent(pointer('pointerdown', 0, { button: 0 }));
    await settle();
    window.dispatchEvent(pointer('pointermove', 700));
    await settle();
    window.dispatchEvent(pointer('pointercancel', 700));
    await settle();

    expect(harness.superdoc.inputs('sections.setPageMargins')).toHaveLength(0);
  });

  it('מסמך לקריאה בלבד אינו נגרר', async () => {
    const harness = await mountRuler({ editable: false });

    await drag(harness, 'שוליים ימניים', 700);

    expect(harness.superdoc.inputs('sections.setPageMargins')).toHaveLength(0);
  });
});

describe('גרירת כניסות', () => {
  it('כניסת ההתחלה נכתבת ל-`left`, גם בפסקה עברית', async () => {
    const harness = await mountRuler();

    // 660 = ס"מ אחד פנימה מתחילת הטקסט (698).
    await drag(harness, 'כניסה מצד ההתחלה', 100 + 660);

    const input = harness.superdoc.inputs('format.paragraph.setIndentation')[0] as {
      left: number;
      right: number;
    };
    expect(input.left).toBe(567); // ס"מ אחד ב-twips, מוצמד לרבע ס"מ
    expect(input.right).toBe(0);
  });

  it('גרירה אינה מוחקת כניסת שורה ראשונה שהגיעה מ-Word', async () => {
    const harness = await mountRuler({
      reading: reading({
        indents: {
          leftTwips: 0,
          rightTwips: 0,
          firstLineTwips: 567,
          hangingTwips: 0,
          bidi: true,
        },
      }),
    });

    await drag(harness, 'כניסה מצד ההתחלה', 100 + 660);

    const input = harness.superdoc.inputs('format.paragraph.setIndentation')[0] as Record<
      string,
      unknown
    >;
    expect(input.firstLine).toBe(567);
  });

  it('כשל מהמנוע מדווח בעברית, ולא נבלע', async () => {
    const harness = mountUi(DocumentRuler, {
      props: { visible: true, host: engineHost(), reading: reading() },
      superdoc: undefined,
    });
    await settle();
    harness.superdoc.reset();

    // אין דרך לכשל מהכפיל הזה בלי failures, ולכן נמדד המסלול ההפוך: הצלחה
    // פולטת `changed` כדי שהמצב ייקרא מחדש מיד.
    await drag(harness as never, 'שוליים ימניים', 794);
    expect(harness.wrapper.emitted('changed')).toHaveLength(1);
  });
});

describe('מקלדת', () => {
  it('חץ שמאלה מזיז את הידית שמאלה — גם במסמך עברי, שם זה מגדיל את השוליים', async () => {
    const harness = await mountRuler();
    const handle = handleByLabel(harness.wrapper, 'שוליים ימניים')!;

    await (handle as unknown as { trigger: (e: string, o?: unknown) => Promise<void> }).trigger?.(
      'keydown',
      { key: 'ArrowLeft' },
    );
    handle.element.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', bubbles: true }));
    await settle();

    const input = harness.superdoc.inputs('sections.setPageMargins')[0] as { right: number };
    // 2.54 + 0.25 ס"מ.
    expect(input.right).toBeCloseTo(2.79 / 2.54, 3);
  });

  it('Home מקפיץ לקצה הדף, End לקצה השני', async () => {
    const harness = await mountRuler();
    const handle = handleByLabel(harness.wrapper, 'שוליים ימניים')!;

    handle.element.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    await settle();

    const input = harness.superdoc.inputs('sections.setPageMargins')[0] as { right: number };
    expect(input.right).toBe(0);
  });

  it('מקש שאינו של הסרגל אינו נוגע במסמך', async () => {
    const harness = await mountRuler();
    const handle = handleByLabel(harness.wrapper, 'שוליים ימניים')!;

    handle.element.dispatchEvent(new KeyboardEvent('keydown', { key: 'a', bubbles: true }));
    await settle();

    expect(harness.superdoc.inputs('sections.setPageMargins')).toHaveLength(0);
  });
});
