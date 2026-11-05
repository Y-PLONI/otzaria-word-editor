import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { measureAllPageTextRuns, measurePageTextRuns } from '../../src/engine/page-ruler';
import { buildPilcrowMarks, type FormattingMarksBlock } from '../../src/engine/formatting-marks-layer';

const block = (nodeId: string, text = ''): FormattingMarksBlock => ({
  nodeId, nodeType: 'paragraph', text, isEmpty: text === '',
});

// Range.getClientRects() recorded in Chromium with SuperDoc 2.23.0, before/after
// the empty-run CSS rule. The preceding mixed run ends LTR; the empty paragraph is RTL.
const captures = [
  { css: 'enabled', width: 0, previousTop: -218.203125, emptyTop: -199.8125 },
  { css: 'disabled', width: 5, previousTop: -213.203125, emptyTop: -194.8125 },
];

describe('empty paragraph measurement → paragraph marks', () => {
  let host: HTMLElement;
  let page: HTMLElement;
  let reference: HTMLElement;
  let rectangles: WeakMap<Node, DOMRect[]>;

  beforeEach(() => {
    host = document.createElement('div');
    page = document.createElement('div');
    page.dataset.pageIndex = '0';
    host.append(page);
    reference = document.createElement('div');
    reference.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1000);
    document.body.append(host, reference);
    rectangles = new WeakMap();
    vi.spyOn(document, 'createRange').mockImplementation(() => {
      const range = new Range();
      range.getClientRects = () => {
        const rects = rectangles.get(range.startContainer) ?? [];
        return Object.assign(rects, { item: (i: number) => rects[i] ?? null }) as unknown as DOMRectList;
      };
      return range;
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
    document.body.replaceChildren();
  });

  function addRun(text: string, rects: DOMRect[], emptyRun = false): void {
    const span = document.createElement('span');
    span.style.direction = 'rtl';
    if (emptyRun) span.className = 'superdoc-empty-run';
    const node = document.createTextNode(text);
    span.append(node);
    page.append(span);
    rectangles.set(node, rects);
  }

  it.each(captures)('keeps an exact mark for a leading empty paragraph (CSS $css)', ({ width, emptyTop }) => {
    addRun('\u00a0', [new DOMRect(87, emptyTop, width, 19)], true);

    const runs = measureAllPageTextRuns(host, reference);
    expect(runs).toEqual([{
      text: '\u00a0', direction: 'rtl',
      rects: [{ leftPx: 87, topPx: emptyTop, widthPx: width, heightPx: 19 }],
    }]);
    expect(buildPilcrowMarks([block('empty')], runs)).toEqual([{
      nodeId: 'empty', topPx: emptyTop, heightPx: 19, anchorXPx: 87, direction: 'rtl', approximate: false,
    }]);
  });

  it.each(captures)('does not inherit the preceding LTR anchor or direction (CSS $css)', ({ width, previousTop, emptyTop }) => {
    const text = 'rtl-left אבג abc';
    addRun(text, [new DOMRect(87, previousTop, 29, 19), new DOMRect(116, previousTop, 37, 19), new DOMRect(153, previousTop, 49, 19)]);
    addRun('\u00a0', [new DOMRect(87, emptyTop, width, 19)], true);

    const runs = measureAllPageTextRuns(host, reference);
    expect(runs.map((run) => run.direction)).toEqual(['ltr', 'rtl']);
    expect(buildPilcrowMarks([block('before', text), block('empty')], runs)).toEqual([
      { nodeId: 'before', topPx: previousTop, heightPx: 19, anchorXPx: 202, direction: 'ltr', approximate: false },
      { nodeId: 'empty', topPx: emptyTop, heightPx: 19, anchorXPx: 87, direction: 'rtl', approximate: false },
    ]);
  });

  it('keeps a larger empty paragraph’s measured height and reference-relative anchor', () => {
    // The 24pt empty-run capture has height 38, while the preceding text has height 19.
    reference.getBoundingClientRect = () => new DOMRect(10, 20, 800, 1000);
    const text = 'rtl-large אבג abc';
    addRun(text, [new DOMRect(558.71875, 76.1875, 29, 19), new DOMRect(587.71875, 76.1875, 37, 19), new DOMRect(624.71875, 76.1875, 64, 19)]);
    addRun('\u00a0', [new DOMRect(688.71875, 94.59375, 0, 38)], true);

    const runs = measurePageTextRuns(page, reference);
    expect(runs[1]?.rects).toEqual([{ leftPx: 678.71875, topPx: 74.59375, widthPx: 0, heightPx: 38 }]);
    const marks = buildPilcrowMarks([block('before', text), block('empty')], runs);
    expect(marks[marks.length - 1]).toEqual({
      nodeId: 'empty', topPx: 74.59375, heightPx: 38, anchorXPx: 678.71875, direction: 'rtl', approximate: false,
    });
  });

  it.each([
    { name: 'unmarked NBSP', text: '\u00a0', emptyRun: false, height: 19, hasRect: true },
    { name: 'marked ZWSP', text: '\u200b', emptyRun: true, height: 19, hasRect: true },
    { name: 'multiple NBSPs', text: '\u00a0\u00a0', emptyRun: true, height: 19, hasRect: true },
    { name: 'zero-height placeholder', text: '\u00a0', emptyRun: true, height: 0, hasRect: true },
    { name: 'absent range rectangles', text: '\u00a0', emptyRun: true, height: 19, hasRect: false },
  ])('still excludes $name', ({ text, emptyRun, height, hasRect }) => {
    addRun(text, hasRect ? [new DOMRect(87, 100, 0, height)] : [], emptyRun);
    const runs = measureAllPageTextRuns(host, reference);
    expect(runs).toEqual([]);
    expect(buildPilcrowMarks([block('empty')], runs)).toEqual([]);
  });
});
