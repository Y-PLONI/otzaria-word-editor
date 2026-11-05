import { afterEach, describe, expect, it, vi } from 'vitest';
import { createFormattingMarksModel, readFormattingMarksBlocks } from '../../src/engine/formatting-marks';
import { invalidateDocumentReads, observeDocumentReadInvalidation } from '../../src/engine/document-read-state';
import { buildMappedPilcrowMarks, type FormattingMarksBlock, type FormattingMarksEnd } from '../../src/engine/formatting-marks-layer';
import { measureFormattingMarksWindow, watchFormattingMarksWindow } from '../../src/engine/formatting-marks-viewport';

const block = (nodeId: string, text = 'same', nodeType = 'paragraph'): FormattingMarksBlock => ({ nodeId, text, nodeType, isEmpty: text === '' });
const end = (nodeId: string, text = 'same', continuesFromPrevious = false, topPx = 10): FormattingMarksEnd => ({
  nodeId, continuesFromPrevious, runs: [{ text, direction: 'rtl', rects: [{ topPx, heightPx: 20, leftPx: 70, widthPx: 30 }] }],
});
const flush = async () => { for (let i = 0; i < 8; i++) await Promise.resolve(); };
const deferred = <T>() => { let resolve!: (value: T) => void; const promise = new Promise<T>(r => { resolve = r; }); return { promise, resolve }; };
afterEach(() => { vi.restoreAllMocks(); vi.useRealTimers(); document.body.replaceChildren(); });

describe('identity-mapped paragraph marks', () => {
  it('does not align repeated or empty paragraphs to a different visible ID', () => {
    const marks = buildMappedPilcrowMarks([block('offscreen'), block('visible'), block('empty', '')], [end('visible'), end('empty', ' ', false, 50)]);
    expect(marks.map(m => [m.nodeId, m.topPx, m.approximate])).toEqual([['visible', 10, false], ['empty', 50, false]]);
  });
  it('accepts a final continued fragment only when it is the matching paragraph suffix', () => {
    expect(buildMappedPilcrowMarks([block('p', 'beginning ending')], [end('p', 'ending', true)])).toHaveLength(1);
    expect(buildMappedPilcrowMarks([block('p', 'beginning ending')], [end('p', 'beginning', true)])).toEqual([]);
    expect(buildMappedPilcrowMarks([block('p', 'beginning ending')], [end('p', 'ending', false)])).toEqual([]);
  });
  it('rejects stale text, duplicate endings, headings, and missing empty placeholders', () => {
    expect(buildMappedPilcrowMarks([block('p', 'edited')], [end('p', 'old')])).toEqual([]);
    expect(buildMappedPilcrowMarks([block('p')], [end('p'), end('p')])).toEqual([]);
    expect(buildMappedPilcrowMarks([block('p', 'same', 'heading')], [end('p')])).toEqual([]);
    expect(buildMappedPilcrowMarks([block('p', '')], [{ nodeId: 'p', continuesFromPrevious: false, runs: [] }])).toEqual([]);
  });
});

describe('filtered paragraph text reads', () => {
  it('requests only missing visible IDs and invalidates after a document change', async () => {
    const list = vi.fn(async ({ nodeIds }: { nodeIds?: string[] } = {}) => ({ blocks: (nodeIds ?? []).map(id => block(id)) }));
    const doc = { blocks: { list } }; const host = { activeEditor: { doc } };
    expect((await readFormattingMarksBlocks(host, ['a', 'b']))?.map(b => b.nodeId)).toEqual(['a', 'b']);
    await readFormattingMarksBlocks(host, ['b', 'c']);
    expect(list.mock.calls.map(([x]) => x?.nodeIds)).toEqual([['a', 'b'], ['c']]);
    expect(list.mock.calls.every(([x]) => (x?.nodeIds?.length ?? 0) > 0)).toBe(true);
    invalidateDocumentReads(doc);
    await readFormattingMarksBlocks(host, ['b']);
    expect(list.mock.calls[list.mock.calls.length - 1]?.[0]?.nodeIds).toEqual(['b']);
    await readFormattingMarksBlocks(host, []);
    expect(list).toHaveBeenCalledTimes(3);
  });
  it('does not cache a failed or invalidated read, and isolates identical IDs in different documents', async () => {
    const pending = deferred<{ blocks: FormattingMarksBlock[] }>();
    const list = vi.fn().mockReturnValueOnce(pending.promise).mockResolvedValue({ blocks: [block('p', 'new')] });
    const doc = { blocks: { list } }; const host = { activeEditor: { doc } };
    const old = readFormattingMarksBlocks(host, ['p']);
    invalidateDocumentReads(doc); pending.resolve({ blocks: [block('p', 'old')] });
    expect(await old).toBeNull();
    expect((await readFormattingMarksBlocks(host, ['p']))?.[0].text).toBe('new');
    const other = { activeEditor: { doc: { blocks: { list: vi.fn().mockResolvedValue({ blocks: [block('p', 'other')] }) } } } };
    expect((await readFormattingMarksBlocks(other, ['p']))?.[0].text).toBe('other');
  });
  it('cancels between batches when a hidden feature no longer needs the result', async () => {
    let current = true;
    const list = vi.fn(async () => { current = false; return { blocks: [block('p')] }; });
    expect(await readFormattingMarksBlocks({ activeEditor: { doc: { blocks: { list } } } }, Array.from({ length: 1001 }, (_, i) => String(i)), () => current)).toBeNull();
    expect(list).toHaveBeenCalledTimes(1);
  });
});

describe('paragraph-mark scheduling', () => {
  it('starts with no full-document read and ignores late results after disable', async () => {
    const pending = deferred<FormattingMarksBlock[]>(); const read = vi.fn(() => pending.promise); const change = vi.fn();
    const model = createFormattingMarksModel({ read, onChange: change });
    model.setEnabled(true); await flush(); expect(read).not.toHaveBeenCalled();
    model.setBlockIds(['a']); expect(read).toHaveBeenCalledTimes(1);
    model.setEnabled(false); pending.resolve([block('a')]); await flush();
    expect(model.getState()).toBeNull(); expect(change).toHaveBeenLastCalledWith(null);
    model.noteDocumentChanged(); model.refreshNow(); await flush(); expect(read).toHaveBeenCalledTimes(1);
    model.dispose();
  });
  it('keeps one running read and only the latest queued viewport, rejecting an old reply immediately', async () => {
    const pending = deferred<FormattingMarksBlock[]>();
    const read = vi.fn().mockReturnValueOnce(pending.promise).mockResolvedValue([block('c')]); const change = vi.fn();
    const model = createFormattingMarksModel({ read, onChange: change });
    model.setBlockIds(['a']); model.setEnabled(true);
    model.setBlockIds(['b']); model.setBlockIds(['c']);
    expect(read).toHaveBeenCalledTimes(1); pending.resolve([block('a')]); await flush();
    expect(read.mock.calls.map(([ids]) => ids)).toEqual([['a'], ['c']]);
    expect(model.getState()).toEqual([block('c')]);
    expect(change.mock.calls.some(([value]) => value?.[0]?.nodeId === 'a')).toBe(false);
    model.dispose();
  });
  it('refreshes unchanged visible IDs after a style/history invalidation without an onUpdate event', async () => {
    vi.useFakeTimers();
    const doc = {}; let type = 'paragraph';
    const read = vi.fn(async () => [block('p', 'same', type)]);
    const model = createFormattingMarksModel({ read, onChange: () => {} });
    const stop = observeDocumentReadInvalidation(doc, model.noteDocumentChanged);
    model.setBlockIds(['p']); model.setEnabled(true); await flush();
    type = 'heading'; invalidateDocumentReads(doc); model.setBlockIds(['p']);
    await vi.advanceTimersByTimeAsync(301);
    expect(read).toHaveBeenCalledTimes(2); expect(model.getState()?.[0].nodeType).toBe('heading');
    type = 'paragraph'; invalidateDocumentReads(doc); model.dispose(); stop();
    await vi.advanceTimersByTimeAsync(301); expect(read).toHaveBeenCalledTimes(2);
  });
});

describe('paragraph-mark viewport measurement', () => {
  function box(el: HTMLElement, top: number, height: number): void {
    vi.spyOn(el, 'getBoundingClientRect').mockReturnValue({ top, bottom: top + height, left: 0, right: 500, x: 0, y: top, width: 500, height, toJSON() {} });
  }
  function page(host: HTMLElement, index: number, top: number, text: string, extra = ''): HTMLElement {
    const el = document.createElement('div'); el.setAttribute('data-page-index', String(index));
    el.innerHTML = `<div class="superdoc-fragment" data-layout-story="body" data-source-node-id="p${index}" data-pm-start="10" data-pm-end="20" ${extra}><span>${text}</span></div>`;
    host.append(el); box(el, top, 1000); return el;
  }
  function rangeRects() {
    const seen: string[] = [];
    const original = Range.prototype.getClientRects;
    Range.prototype.getClientRects = function () {
      seen.push(this.startContainer.nodeValue ?? '');
      return [{ left: 10, top: 10, width: 30, height: 20 }] as unknown as DOMRectList;
    };
    return { seen, restore: () => { Range.prototype.getClientRects = original; } };
  }
  it('measures text only in visible pages plus margin and never measures a continuing fragment', () => {
    const host = document.createElement('div'); document.body.append(host); box(host, 0, 800);
    page(host, 0, -3000, 'above'); page(host, 1, 0, 'visible'); page(host, 2, 3000, 'below');
    page(host, 3, 1000, 'continues', 'data-continues-on-next="true"');
    const ranges = rangeRects();
    try {
      const result = measureFormattingMarksWindow(host, host);
      expect(result.mode).toBe('mapped');
      if (result.mode === 'mapped') expect(result.endings.map(e => e.nodeId)).toEqual(['p1']);
      expect(ranges.seen).toEqual(['visible']);
    } finally { ranges.restore(); }
  });
  it('keeps both complete sequences when identity mapping is unavailable', () => {
    const host = document.createElement('div'); document.body.append(host); box(host, 0, 800);
    const first = page(host, 0, 0, 'unknown'); first.firstElementChild!.removeAttribute('class');
    page(host, 1, 3000, 'offscreen'); const ranges = rangeRects();
    try {
      const result = measureFormattingMarksWindow(host, host);
      expect(result.mode).toBe('legacy'); expect(ranges.seen).toEqual(['unknown', 'offscreen']);
    } finally { ranges.restore(); }
  });
  it('falls back when one body paragraph loses its ID but still recognizes a table', () => {
    const host = document.createElement('div'); document.body.append(host); box(host, 0, 800);
    const first = page(host, 0, 0, 'known');
    const unknown = first.firstElementChild!.cloneNode(true) as HTMLElement;
    unknown.removeAttribute('data-source-node-id'); first.append(unknown);
    const ranges = rangeRects();
    try {
      expect(measureFormattingMarksWindow(host, host).mode).toBe('legacy');
      unknown.setAttribute('data-table-boundaries', '{}');
      const mapped = measureFormattingMarksWindow(host, host);
      expect(mapped.mode).toBe('mapped');
      if (mapped.mode === 'mapped') expect(mapped.endings.map(e => e.nodeId)).toEqual(['p0']);
    } finally { ranges.restore(); }
  });
  it('removes listeners and settling timers on dispose', async () => {
    vi.useFakeTimers(); const host = document.createElement('div'); box(host, 0, 800);
    const change = vi.fn(); const unobserve = vi.fn();
    const watch = watchFormattingMarksWindow({ host, reference: host, ui: { viewport: { observe: () => unobserve } }, onChange: change });
    expect(change).toHaveBeenCalledTimes(1);
    watch.measure(); watch.dispose(); host.dispatchEvent(new Event('scroll'));
    await vi.advanceTimersByTimeAsync(1000);
    expect(change).toHaveBeenCalledTimes(1); expect(unobserve).toHaveBeenCalledTimes(1);
  });
});
