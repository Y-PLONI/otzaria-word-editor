/**
 * כיווץ שורה מצוירת שחורגת מהשוליים — ראו src/engine/justify-last-line.ts.
 *
 * ב-jsdom אין פריסה, ולכן הקופסה של השורה, הקופסאות של הטקסט ו-`offsetWidth`
 * מוזרקים לכל שורה. ההתנהגות מול המנוע האמיתי נמדדת בשער
 * scripts/qa/justify-last-line-qa.mjs.
 */
import { afterEach, describe, expect, it } from 'vitest';
import {
  compressSpillingLine,
  installJustifyLastLine,
  lineSpill,
  type JustifyLastLineHandle,
} from '../../src/engine/justify-last-line';

interface Geometry {
  box: { left: number; right: number };
  ink: { left: number; right: number };
  /** רוחב השורה לפני ה-transform של הזום. ברירת מחדל: רוחב הקופסה. */
  cssWidth?: number;
}

const geometry = new WeakMap<Node, Geometry>();

const rect = (left: number, right: number): DOMRect =>
  ({ left, right, width: right - left, top: 0, bottom: 10, height: 10, x: left, y: 0 }) as DOMRect;

const realCreateRange = document.createRange.bind(document);
document.createRange = () => {
  let target: Node | null = null;
  const range = realCreateRange();
  range.selectNodeContents = (node: Node) => {
    target = node;
  };
  range.getClientRects = () => {
    const g = target ? geometry.get(target) : undefined;
    return (g ? [rect(g.ink.left, g.ink.right)] : []) as unknown as DOMRectList;
  };
  return range;
};

function paintLine(text: string, g: Geometry, wordSpacing = ''): HTMLElement {
  const line = document.createElement('div');
  line.className = 'superdoc-line';
  line.textContent = text;
  line.style.wordSpacing = wordSpacing;
  setGeometry(line, g);
  return line;
}

function setGeometry(line: HTMLElement, g: Geometry): void {
  geometry.set(line, g);
  line.getBoundingClientRect = () => rect(g.box.left, g.box.right);
  Object.defineProperty(line, 'offsetWidth', {
    configurable: true,
    get: () => g.cssWidth ?? g.box.right - g.box.left,
  });
}

/** שורה עברית ברוחב 600 שהטקסט שלה יוצא שמאלה ב-10 פיקסלים. */
const SPILLING: Geometry = { box: { left: 100, right: 700 }, ink: { left: 90, right: 700 } };
const FOUR_SPACES = 'אבג אבג אבג אבג להדיא';

describe('lineSpill', () => {
  it('measures the spill on the left, where a Hebrew line overflows', () => {
    expect(lineSpill(paintLine(FOUR_SPACES, SPILLING))).toBe(10);
  });

  it('measures the spill on the right, where a Latin line overflows', () => {
    expect(lineSpill(paintLine('a b', { box: { left: 100, right: 700 }, ink: { left: 100, right: 706 } }))).toBe(6);
  });

  it('converts screen pixels back to CSS pixels under zoom', () => {
    const zoomed: Geometry = { box: { left: 200, right: 1400 }, ink: { left: 180, right: 1400 }, cssWidth: 600 };
    expect(lineSpill(paintLine(FOUR_SPACES, zoomed))).toBe(10);
  });

  it('is 0 for a line inside its box', () => {
    expect(lineSpill(paintLine(FOUR_SPACES, { box: { left: 100, right: 700 }, ink: { left: 140, right: 700 } }))).toBe(
      0,
    );
  });
});

describe('compressSpillingLine', () => {
  it('spreads the spill over the spaces as negative word-spacing', () => {
    const line = paintLine(FOUR_SPACES, SPILLING);
    expect(compressSpillingLine(line)).toBe(true);
    expect(line.style.wordSpacing).toBe('-2.5px');
  });

  it('counts no-break spaces, which the engine also compresses', () => {
    const line = paintLine('אבג אבג אבג אבג להדיא', SPILLING);
    compressSpillingLine(line);
    expect(line.style.wordSpacing).toBe('-2.5px');
  });

  it('leaves a line the engine already spaces alone', () => {
    const line = paintLine(FOUR_SPACES, SPILLING, '-1px');
    expect(compressSpillingLine(line)).toBe(false);
    expect(line.style.wordSpacing).toBe('-1px');
  });

  it('leaves a single word wider than the column alone', () => {
    const line = paintLine('להדיאלהדיאלהדיא', SPILLING);
    expect(compressSpillingLine(line)).toBe(false);
    expect(line.style.wordSpacing).toBe('');
  });

  it('ignores sub-pixel rounding', () => {
    const line = paintLine(FOUR_SPACES, { box: { left: 100, right: 700 }, ink: { left: 99.6, right: 700 } });
    expect(compressSpillingLine(line)).toBe(false);
  });

  it('re-measures its own compression after the engine updates the line in place', () => {
    const line = paintLine(FOUR_SPACES, SPILLING);
    compressSpillingLine(line);
    // המנוע החליף את הריצה: השורה קצרה יותר, ובלי הכיווץ שלנו היא כבר בפנים.
    line.textContent = 'אבג אבג';
    setGeometry(line, { box: { left: 100, right: 700 }, ink: { left: 400, right: 700 } });
    expect(compressSpillingLine(line)).toBe(false);
    expect(line.style.wordSpacing).toBe('');
  });
});

describe('installJustifyLastLine', () => {
  let handle: JustifyLastLineHandle | null = null;
  afterEach(() => {
    handle?.dispose();
    handle = null;
    document.body.replaceChildren();
  });

  const nextFrame = () => new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));

  it('compresses a line the engine paints after install', async () => {
    const root = document.createElement('div');
    document.body.append(root);
    handle = installJustifyLastLine(root);
    const page = document.createElement('div');
    const line = paintLine(FOUR_SPACES, SPILLING);
    page.append(line);
    root.append(page);
    await nextFrame();
    expect(line.style.wordSpacing).toBe('-2.5px');
  });

  it('compresses lines that were already painted', async () => {
    const root = document.createElement('div');
    const line = paintLine(FOUR_SPACES, SPILLING);
    root.append(line);
    document.body.append(root);
    handle = installJustifyLastLine(root);
    await nextFrame();
    expect(line.style.wordSpacing).toBe('-2.5px');
  });

  it('stops after dispose', async () => {
    const root = document.createElement('div');
    document.body.append(root);
    handle = installJustifyLastLine(root);
    handle.dispose();
    const line = paintLine(FOUR_SPACES, SPILLING);
    root.append(line);
    await nextFrame();
    expect(line.style.wordSpacing).toBe('');
  });
});
