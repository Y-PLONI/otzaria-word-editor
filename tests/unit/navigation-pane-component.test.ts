import { afterEach, describe, expect, it, vi } from 'vitest';
import { mount, flushPromises } from '@vue/test-utils';
import NavigationPane from '../../src/ui/shell/NavigationPane.vue';

const blocks = [
  { nodeId: 'a', text: 'Chapter', headingLevel: 1 },
  { nodeId: 'b', text: 'Section', headingLevel: 2 },
  { nodeId: 'c', text: 'Next chapter', headingLevel: 1 },
];
function host(entries = blocks) {
  const list = vi.fn(({ offset, limit, nodeIds, includeText }: { offset: number; limit: number; nodeIds?: string[]; includeText: boolean }) => ({
    blocks: (nodeIds ? entries.filter((entry) => nodeIds.includes(entry.nodeId)) : entries).slice(offset, offset + limit)
      .map((entry) => includeText ? entry : { ...entry, text: undefined }),
  }));
  const current = vi.fn(() => ({ target: { segments: [{ blockId: 'a' }] } }));
  const apply = vi.fn();
  return { activeEditor: { doc: { blocks: { list }, selection: { current } } }, ui: { selection: { apply } }, list, current, apply };
}
const mounted: ReturnType<typeof mount>[] = [];
function pane(h = host()) {
  const w = mount(NavigationPane, { attachTo: document.body, props: { host: h as never, generation: 1, docRevision: 0, caretRevision: 0 } });
  mounted.push(w);
  return w;
}
afterEach(() => { mounted.splice(0).forEach((w) => w.unmount()); vi.useRealTimers(); });

describe('navigation pane interaction and async safety', () => {
  it('roves through visible headings, expands in RTL, and activates with Enter/Space', async () => {
    const h = host(); const w = pane(h); await flushPromises();
    const items = () => w.findAll('[role=treeitem]');
    const focused = () => (document.activeElement as HTMLElement)?.textContent?.trim();
    expect(items().map((item) => item.attributes('tabindex'))).toEqual(['0', '-1', '-1']);
    (items()[0].element as HTMLElement).focus();
    await items()[0].trigger('keydown', { key: 'ArrowRight' });
    expect(items()).toHaveLength(2);
    await items()[0].trigger('keydown', { key: 'ArrowDown' });
    expect(focused()).toBe('Next chapter');
    await items()[1].trigger('keydown', { key: 'Home' });
    expect(focused()).toBe('Chapter');
    await items()[0].trigger('keydown', { key: 'ArrowLeft' });
    expect(items()).toHaveLength(3);
    await items()[0].trigger('keydown', { key: 'ArrowLeft' });
    expect(focused()).toBe('Section');
    await items()[1].trigger('keydown', { key: 'ArrowRight' });
    expect(focused()).toBe('Chapter');
    await items()[0].trigger('keydown', { key: 'End' });
    expect(focused()).toBe('Next chapter');
    await items()[2].trigger('keydown', { key: 'ArrowUp' });
    expect(focused()).toBe('Section');
    await items()[1].trigger('keydown', { key: 'Enter' }); await flushPromises();
    expect(h.apply).toHaveBeenLastCalledWith(expect.objectContaining({ start: { kind: 'text', blockId: 'b', offset: 0 } }));
    await items()[1].trigger('keydown', { key: ' ' }); await flushPromises();
    expect(w.emitted('navigated')).toHaveLength(2);
  });

  it('announces a partial empty outline and never claims there are no headings in the document', async () => {
    const h = host(Array.from({ length: 25001 }, (_, i) => ({ nodeId: `b${i}`, text: `text${i}`, headingLevel: i === 25000 ? 1 : 0 })));
    const w = pane(h); await flushPromises();
    expect(w.get('[role=status]').text()).toContain('ניווט חלקי');
    expect(w.get('.nav-empty').text()).toBe('לא נמצאו כותרות בחלק המסמך שנקרא.');
    expect(h.list).toHaveBeenCalledTimes(51);
  });

  it('clears a previous current section when caret moves beyond the known map without rescanning', async () => {
    vi.useFakeTimers(); const h = host(); const w = pane(h); await flushPromises();
    expect(w.get('[aria-current=location]').text()).toContain('Chapter');
    h.current.mockReturnValue({ target: { segments: [{ blockId: 'unscanned' }] } });
    await w.setProps({ caretRevision: 1 }); await vi.advanceTimersByTimeAsync(151);
    expect(w.find('[aria-current=location]').exists()).toBe(false);
    expect(h.list).toHaveBeenCalledTimes(2);
  });

  it('ignores an out-of-order selection read', async () => {
    vi.useFakeTimers(); const h = host(); const w = pane(h); await flushPromises();
    let resolve!: (value: ReturnType<typeof h.current>) => void;
    h.current.mockImplementationOnce(() => new Promise((r) => { resolve = r; }) as never);
    await w.setProps({ caretRevision: 1 }); await vi.advanceTimersByTimeAsync(151);
    h.current.mockReturnValue({ target: { segments: [{ blockId: 'c' }] } });
    await w.setProps({ caretRevision: 2 }); await vi.advanceTimersByTimeAsync(151);
    resolve({ target: { segments: [{ blockId: 'a' }] } }); await flushPromises();
    expect(w.get('[aria-current=location]').text()).toContain('Next chapter');
  });

  it('cancels old pages on generation change even when the host object is reused', async () => {
    const h = host(); let resolve!: (value: { blocks: typeof blocks }) => void;
    h.list.mockImplementationOnce(() => new Promise((r) => { resolve = r; }) as never);
    const w = pane(h); await w.setProps({ generation: 2 }); await flushPromises();
    resolve({ blocks: Array.from({ length: 500 }, () => blocks[0]) }); await flushPromises();
    expect(h.list).toHaveBeenCalledTimes(3);
    expect(w.findAll('[role=treeitem]')).toHaveLength(3);
  });

  it('unmount cancels both scheduled refresh and pending pagination', async () => {
    vi.useFakeTimers(); const h = host(); let resolve!: (value: { blocks: typeof blocks }) => void;
    h.list.mockImplementationOnce(() => new Promise((r) => { resolve = r; }) as never);
    const w = pane(h); await w.setProps({ docRevision: 1 }); w.unmount(); mounted.splice(mounted.indexOf(w), 1);
    resolve({ blocks: Array.from({ length: 500 }, () => blocks[0]) }); await flushPromises();
    await vi.advanceTimersByTimeAsync(401);
    expect(h.list).toHaveBeenCalledTimes(1); expect(h.current).not.toHaveBeenCalled();
  });
});
