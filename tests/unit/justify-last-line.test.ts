import { afterEach, describe, expect, it, vi } from 'vitest';
import { compressSpillingLine, installJustifyLastLine, isFinalTextLine, lineSpill, type JustifyLastLineHandle } from '../../src/engine/justify-last-line';
import { invalidateDocumentReads } from '../../src/engine/document-read-state';

interface Geometry { box: { left: number; right: number }; ink: { left: number; right: number }; cssWidth?: number; }
const geometries = new WeakMap<Node, Geometry>();
const ink = new WeakMap<Node, { left: number; right: number }>();
const rect = (left: number, right: number): DOMRect => ({ left, right, width: right - left, top: 0, bottom: 10, height: 10, x: left, y: 0 }) as DOMRect;
const realCreateRange = document.createRange.bind(document);
document.createRange = () => {
  let target: Node | null = null, from = 0, to = 0;
  const range = realCreateRange();
  range.setStart = (node, offset) => { target = node; from = offset; };
  range.setEnd = (_node, offset) => { to = offset; };
  range.selectNodeContents = (node) => { target = node; };
  const bounds = () => {
    if (!target) return [];
    const line = target instanceof Element ? target.closest('.superdoc-line') : target.parentElement?.closest('.superdoc-line');
    const g = geometries.get(target) ?? (line ? geometries.get(line) : undefined);
    const value = ink.get(target) ?? g?.ink;
    return value ? [rect(value.left, value.right)] : [];
  };
  range.getClientRects = () => bounds() as unknown as DOMRectList;
  range.getBoundingClientRect = () => to - from === 1 ? rect(0, 8) : (bounds()[0] ?? rect(0, 0));
  return range;
};
const SPILLING: Geometry = { box: { left: 100, right: 700 }, ink: { left: 90, right: 700 } };
const FOUR_SPACES = 'אבג אבג אבג אבג להדיא';
function setGeometry(line: HTMLElement, g: Geometry) {
  geometries.set(line, g);
  line.getBoundingClientRect = () => rect(g.box.left, g.box.right);
  Object.defineProperty(line, 'offsetWidth', { configurable: true, get: () => g.cssWidth ?? g.box.right - g.box.left });
}
function paintLine(text = FOUR_SPACES, g = SPILLING, spacing = '') {
  const fragment = document.createElement('div');
  fragment.className = 'superdoc-fragment'; fragment.dataset.layoutStory = 'body'; fragment.dataset.sourceNodeId = 'p1';
  const line = document.createElement('div'); line.className = 'superdoc-line'; line.dir = 'rtl';
  const run = document.createElement('span'); run.className = 'superdoc-text-run'; run.textContent = text;
  line.append(run); fragment.append(line); line.style.wordSpacing = spacing; setGeometry(line, g);
  return line;
}
function api(alignment = 'justify') {
  return { getNodeById: vi.fn(async ({ nodeId }: { nodeId: string }) => ({
    address: { kind: 'block', nodeId }, node: { kind: 'paragraph', paragraph: { props: { alignment } } },
  })) };
}
const nextFrame = () => new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));

describe('text-only overflow', () => {
  it('measures the ending edge of RTL text', () => expect(lineSpill(paintLine())).toBe(10));
  it('measures the ending edge of LTR text', () => {
    const line = paintLine('a b', { box: { left: 100, right: 700 }, ink: { left: 100, right: 706 } }); line.dir = 'ltr';
    expect(lineSpill(line)).toBe(6);
  });
  it('converts transformed pixels to CSS pixels', () => expect(lineSpill(paintLine(FOUR_SPACES,
    { box: { left: 200, right: 1400 }, ink: { left: 180, right: 1400 }, cssWidth: 600 }))).toBe(10));
  it('does not treat a hanging leading indent as ending overflow', () => {
    expect(lineSpill(paintLine(FOUR_SPACES, { box: { left: 100, right: 700 }, ink: { left: 100, right: 720 } }))).toBe(0);
  });
  it('excludes a hanging list marker even when the whole line Range would overflow', () => {
    const line = paintLine(); const text = line.querySelector('span')!.firstChild!;
    ink.set(text, { left: 140, right: 700 });
    const marker = document.createElement('span'); marker.className = 'superdoc-list-marker'; marker.textContent = '1.';
    line.prepend(marker);
    expect(lineSpill(line)).toBe(0); expect(compressSpillingLine(line, true)).toBe(false);
  });
});

describe('bounded, verified final-line compression', () => {
  it('requires explicit justification approval', () => expect(compressSpillingLine(paintLine())).toBe(false));
  it('distributes overflow over actual text spaces', () => {
    const line = paintLine(); expect(compressSpillingLine(line, true)).toBe(true); expect(line.style.wordSpacing).toBe('-2.5px');
  });
  it('counts no-break text spaces', () => {
    const line = paintLine('אבג\u00a0אבג אבג\u00a0אבג להדיא'); compressSpillingLine(line, true); expect(line.style.wordSpacing).toBe('-2.5px');
  });
  it('excludes marker suffix spaces from the divisor', () => {
    const line = paintLine(); const suffix = document.createElement('span');
    suffix.className = 'superdoc-tab superdoc-marker-suffix-tab'; suffix.textContent = '\u00a0'; suffix.style.wordSpacing = '0px'; line.prepend(suffix);
    expect(compressSpillingLine(line, true)).toBe(true); expect(line.style.wordSpacing).toBe('-2.5px');
  });
  it('does not compress a normal fixed-position tab line', () => {
    const line = paintLine(); const tab = document.createElement('span'); tab.className = 'superdoc-tab'; tab.textContent = '\u00a0'; line.prepend(tab);
    expect(compressSpillingLine(line, true)).toBe(false);
  });
  it('does not count outside whitespace as compressible content', () => {
    const line = paintLine(`  ${FOUR_SPACES}  `); compressSpillingLine(line, true); expect(parseFloat(line.style.wordSpacing)).toBeCloseTo(-10 / 6);
  });
  it('leaves a single word with trailing spaces alone', () => expect(compressSpillingLine(paintLine('להדיא  '), true)).toBe(false));
  it('leaves extreme overflow alone rather than overlapping words', () => expect(compressSpillingLine(paintLine('א ב'), true)).toBe(false));
  it('leaves engine-owned spacing alone', () => {
    const line = paintLine(FOUR_SPACES, SPILLING, '-1px'); expect(compressSpillingLine(line, true)).toBe(false); expect(line.style.wordSpacing).toBe('-1px');
  });
  it('leaves runs with their own spacing alone', () => {
    const line = paintLine(); (line.firstChild as HTMLElement).style.wordSpacing = '0px'; expect(compressSpillingLine(line, true)).toBe(false);
  });
  it('ignores subpixel rounding', () => expect(compressSpillingLine(paintLine(FOUR_SPACES,
    { box: { left: 100, right: 700 }, ink: { left: 99.6, right: 700 } }), true)).toBe(false));
  it('excludes a non-final line and a fragment continued on the next page', () => {
    const line = paintLine(); const second = paintLine(); line.parentElement!.append(second);
    expect(isFinalTextLine(line)).toBe(false); line.parentElement!.setAttribute('data-continues-on-next', 'true'); expect(isFinalTextLine(second)).toBe(false);
  });
  it('removes its previous spacing after justification is lost', () => {
    const line = paintLine(); compressSpillingLine(line, true); compressSpillingLine(line, false); expect(line.style.wordSpacing).toBe('');
  });
  it('removes its previous spacing after the text fits', () => {
    const line = paintLine(); compressSpillingLine(line, true); setGeometry(line, { box: SPILLING.box, ink: { left: 200, right: 700 } });
    expect(compressSpillingLine(line, true)).toBe(false); expect(line.style.wordSpacing).toBe('');
  });
  it('does not overwrite an engine replacement of its old spacing', () => {
    const line = paintLine(); compressSpillingLine(line, true); line.style.wordSpacing = '2px'; compressSpillingLine(line, false); expect(line.style.wordSpacing).toBe('2px');
  });
});

describe('layout lifecycle', () => {
  let handle: JustifyLastLineHandle | undefined;
  afterEach(() => { handle?.dispose(); handle = undefined; document.body.replaceChildren(); });
  function mount(line = paintLine(), doc = api()) {
    const root = document.createElement('div'); root.append(line.parentElement!); document.body.append(root);
    handle = installJustifyLastLine(root, doc); return { root, line, doc };
  }
  it('processes an already painted line using a targeted paragraph read', async () => {
    const { line, doc } = mount(); await nextFrame(); expect(line.style.wordSpacing).toBe('-2.5px'); expect(doc.getNodeById).toHaveBeenCalledWith({ nodeId: 'p1' });
  });
  it('processes newly added lines', async () => {
    const { root } = mount(); await nextFrame(); const line = paintLine(); root.append(line.parentElement!); await nextFrame(); expect(line.style.wordSpacing).toBe('-2.5px');
  });
  it('does not enter an observer loop on its own style writes', async () => {
    const { doc } = mount(); await nextFrame(); await nextFrame(); await nextFrame(); expect(doc.getNodeById).toHaveBeenCalledTimes(1);
  });
  it('recalculates on characterData without a replaced text node', async () => {
    const { line } = mount(); await nextFrame(); setGeometry(line, { box: SPILLING.box, ink: { left: 200, right: 700 } });
    line.querySelector('span')!.firstChild!.textContent = 'קצר'; await nextFrame(); expect(line.style.wordSpacing).toBe('');
  });
  it('recalculates after removal without added nodes', async () => {
    const { line } = mount(); await nextFrame(); setGeometry(line, { box: SPILLING.box, ink: { left: 200, right: 700 } });
    line.firstChild!.remove(); await nextFrame(); expect(line.style.wordSpacing).toBe('');
  });
  it('recalculates when the same line receives a width change', async () => {
    const { line } = mount(); await nextFrame(); setGeometry(line, { box: SPILLING.box, ink: { left: 200, right: 700 } });
    line.style.width = '800px'; await nextFrame(); expect(line.style.wordSpacing).toBe('');
  });
  it('invalidates a cached justification after a document edit', async () => {
    const { line, doc } = mount(); await nextFrame(); doc.getNodeById.mockResolvedValue({ address: { kind: 'block', nodeId: 'p1' }, node: { kind: 'paragraph', paragraph: { props: { alignment: 'right' } } } });
    invalidateDocumentReads(doc); await nextFrame(); expect(line.style.wordSpacing).toBe(''); expect(doc.getNodeById).toHaveBeenCalledTimes(2);
  });
  it('discards an approval returned after DOM changes', async () => {
    let resolve!: (value: unknown) => void;
    const doc = { getNodeById: () => new Promise((r) => { resolve = r; }) };
    const line = paintLine(), root = document.createElement('div'); root.append(line.parentElement!); document.body.append(root);
    handle = installJustifyLastLine(root, doc); await nextFrame();
    line.querySelector('span')!.firstChild!.textContent = 'קצר'; setGeometry(line, { box: SPILLING.box, ink: { left: 200, right: 700 } });
    await Promise.resolve(); resolve({ address: { kind: 'block', nodeId: 'p1' }, node: { kind: 'paragraph', paragraph: { props: { alignment: 'justify' } } } });
    await nextFrame(); expect(line.style.wordSpacing).toBe('');
  });
  it('refreshes a previously rejected candidate after a style-only edit', async () => {
    const { line, doc } = mount(paintLine(), api('right')); await nextFrame(); expect(line.style.wordSpacing).toBe('');
    doc.getNodeById.mockResolvedValue({ address: { kind: 'block', nodeId: 'p1' }, node: { kind: 'paragraph', paragraph: { props: { alignment: 'justify' } } } });
    invalidateDocumentReads(doc); await nextFrame(); expect(line.style.wordSpacing).toBe('-2.5px');
  });
  it('retries a candidate invalidated during its paragraph read without needing a DOM edit', async () => {
    let resolve!: (value: unknown) => void;
    const doc = { getNodeById: vi.fn().mockImplementationOnce(() => new Promise((r) => { resolve = r; }))
      .mockResolvedValue({ address: { kind: 'block', nodeId: 'p1' }, node: { kind: 'paragraph', paragraph: { props: { alignment: 'justify' } } } }) };
    const line = paintLine(), root = document.createElement('div'); root.append(line.parentElement!); document.body.append(root);
    handle = installJustifyLastLine(root, doc); await nextFrame(); invalidateDocumentReads(doc); await nextFrame();
    resolve({ address: { kind: 'block', nodeId: 'p1' }, node: { kind: 'paragraph', paragraph: { props: { alignment: 'right' } } } });
    await nextFrame(); expect(doc.getNodeById).toHaveBeenCalledTimes(2); expect(line.style.wordSpacing).toBe('-2.5px');
  });
  it('does not rescan lines for selection overlays or root focus classes', async () => {
    const { root, line, doc } = mount(); await nextFrame(); await nextFrame();
    const measure = vi.spyOn(line, 'getBoundingClientRect'); const overlay = document.createElement('div'); overlay.className = 'selection'; root.append(overlay); root.classList.add('focused');
    await nextFrame(); expect(measure).not.toHaveBeenCalled(); expect(doc.getNodeById).toHaveBeenCalledTimes(1);
  });
  it('restores spacing and rejects a pending read after disposal', async () => {
    const { line } = mount(); await nextFrame(); handle!.dispose(); expect(line.style.wordSpacing).toBe('');
    line.style.width = '800px'; await nextFrame(); expect(line.style.wordSpacing).toBe('');
  });
});
