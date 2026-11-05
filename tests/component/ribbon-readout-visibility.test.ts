import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { invalidateDocumentReads } from '../../src/engine/document-read-state';
import type { ParagraphFormatDocumentApi } from '../../src/engine/paragraph-format';
import { pageBreakTracker } from '../../src/engine/page-break';
import { createFontMemory } from '../../src/composables/use-font-controls';
import Ribbon from '../../src/ui/ribbon/Ribbon.vue';
import HomeTab from '../../src/ui/ribbon/tabs/HomeTab.vue';
import InsertTab from '../../src/ui/ribbon/tabs/InsertTab.vue';
import ContextMenu from '../../src/ui/menu/ContextMenu.vue';
import { contextMenuModel } from '../../src/ui/menu/context-menu-model';
import { autoUnmount, buttonByTip, createCommandDouble, createSuperdocDouble, mountUi, pickerValue, setPicker, settle } from './harness';

autoUnmount();
beforeEach(() => { vi.useFakeTimers(); });
afterEach(() => { vi.useRealTimers(); });

const target = (blockId = 'block-1') => ({
  empty: true,
  target: { kind: 'text', segments: [{ blockId, range: { start: 0, end: 0 } }] },
});

function documentWithSelection(bidi = false) {
  const paragraphProps = { bidi };
  const superdoc = createSuperdocDouble({ paragraphProps });
  const doc = superdoc.host.activeEditor!.doc as unknown as ParagraphFormatDocumentApi;
  const current = vi.fn(() => target());
  doc.selection = { current };
  const listeners = new Set<() => void>();
  const removed: Array<() => void> = [];
  const observe = vi.fn((callback: () => void) => {
    listeners.add(callback);
    callback(); // The real SnapshotSubscribable publishes an initial snapshot.
    return () => { listeners.delete(callback); removed.push(callback); };
  });
  const selection = superdoc.host.ui.selection as unknown as { observe: typeof observe };
  selection.observe = observe;
  return { superdoc, doc, current, paragraphProps, listeners, removed, observe,
    notify: () => { for (const listener of listeners) listener(); } };
}

const pageBreakButton = (harness: ReturnType<typeof mountUi>) =>
  harness.wrapper.find('button[data-tip-title="התחל בעמוד חדש"]');

describe('collapsed ribbon readouts', () => {
  for (const activeTab of ['home', 'insert'] as const) {
    it(`${activeTab}: stays mounted, releases readout subscriptions, and refreshes once per reveal`, async () => {
      const model = documentWithSelection();
      const harness = mountUi(Ribbon, { superdoc: model.superdoc, props: { activeTab, collapsed: true } });
      await settle();
      const tab = harness.wrapper.findComponent(activeTab === 'home' ? HomeTab : InsertTab).element;
      expect(model.observe).not.toHaveBeenCalled();
      expect(model.current).not.toHaveBeenCalled();

      for (let cycle = 0; cycle < 3; cycle += 1) {
        model.notify();
        invalidateDocumentReads(model.doc);
        pageBreakTracker.forgetAll();
        await vi.advanceTimersByTimeAsync(500);
        const before = model.current.mock.calls.length;
        await harness.wrapper.setProps({ collapsed: false });
        await settle();
        expect(model.current).toHaveBeenCalledTimes(before + 1);
        expect(model.listeners.size).toBe(1);
        await vi.advanceTimersByTimeAsync(500);
        expect(model.current).toHaveBeenCalledTimes(before + 1);

        model.notify(); // Queue a refresh, then fold before the debounce fires.
        await harness.wrapper.setProps({ collapsed: true });
        expect(model.listeners.size).toBe(0);
        for (const callback of model.removed) callback(); // A queued old notification is harmless.
        await vi.advanceTimersByTimeAsync(500);
        expect(model.current).toHaveBeenCalledTimes(before + 1);
        expect(harness.wrapper.findComponent(activeTab === 'home' ? HomeTab : InsertTab).element).toBe(tab);
      }
      harness.wrapper.unmount();
      expect(model.listeners.size).toBe(0);
      expect(model.observe).toHaveBeenCalledTimes(3);
    });
  }

  it('Home rejects a response arriving after collapse and reads current formatting when revealed', async () => {
    const model = documentWithSelection(true);
    let resolveOld!: (value: unknown) => void;
    const original = model.doc.getNodeById!;
    model.doc.getNodeById = vi.fn(() => new Promise(done => { resolveOld = done; }));
    const harness = mountUi(Ribbon, { superdoc: model.superdoc });
    await vi.advanceTimersByTimeAsync(150);
    await harness.wrapper.setProps({ collapsed: true });
    resolveOld({ node: { kind: 'paragraph', paragraph: { props: { bidi: true } } }, address: { kind: 'block', nodeType: 'paragraph', nodeId: 'block-1' } });
    await settle();
    expect(buttonByTip(harness.wrapper, 'כיוון פסקה מימין לשמאל').attributes('aria-pressed')).toBe('false');

    model.paragraphProps.bidi = false;
    model.doc.getNodeById = original;
    invalidateDocumentReads(model.doc);
    const before = model.current.mock.calls.length;
    await harness.wrapper.setProps({ collapsed: false });
    await settle();
    expect(model.current).toHaveBeenCalledTimes(before + 1);
    expect(buttonByTip(harness.wrapper, 'כיוון פסקה משמאל לימין').attributes('aria-pressed')).toBe('true');
  });

  it('Insert keeps Undo/Redo tracking live without reading the hidden selection', async () => {
    const model = documentWithSelection();
    const harness = mountUi(Ribbon, { superdoc: model.superdoc, props: { activeTab: 'insert' } });
    await vi.advanceTimersByTimeAsync(150);
    pageBreakTracker.remember('block-1', true);
    await settle();
    expect(pageBreakButton(harness).attributes('aria-pressed')).toBe('true');

    await harness.wrapper.setProps({ collapsed: true });
    model.current.mockClear();
    pageBreakTracker.forgetAllKeepingSnapshot();
    await vi.advanceTimersByTimeAsync(500);
    expect(model.current).not.toHaveBeenCalled();
    await harness.wrapper.setProps({ collapsed: false });
    await settle();
    expect(pageBreakButton(harness).attributes('aria-pressed')).toBe('false');
    expect(model.current).toHaveBeenCalledTimes(1);

    await harness.wrapper.setProps({ collapsed: true });
    model.current.mockClear();
    expect(pageBreakTracker.restoreSnapshot()).toBe(true);
    await vi.advanceTimersByTimeAsync(500);
    expect(model.current).not.toHaveBeenCalled();
    await harness.wrapper.setProps({ collapsed: false });
    await settle();
    expect(pageBreakButton(harness).attributes('aria-pressed')).toBe('true');
    expect(model.current).toHaveBeenCalledTimes(1);
  });

  it('Home waits until reveal to read a replacement document', async () => {
    const first = documentWithSelection(true);
    const harness = mountUi(Ribbon, { superdoc: first.superdoc });
    await vi.advanceTimersByTimeAsync(150);
    await harness.wrapper.setProps({ collapsed: true });
    const second = documentWithSelection(false);
    await harness.setSuperdoc(second.superdoc);
    expect(second.observe).not.toHaveBeenCalled();
    expect(second.current).not.toHaveBeenCalled();
    await harness.wrapper.setProps({ collapsed: false });
    await settle();
    expect(second.current).toHaveBeenCalledTimes(1);
    expect(buttonByTip(harness.wrapper, 'כיוון פסקה משמאל לימין').attributes('aria-pressed')).toBe('true');
    expect(first.listeners.size).toBe(0);
  });

  it('Insert synchronizes a document replacement while hidden, clearing old tracking and pending history', async () => {
    const first = documentWithSelection();
    const harness = mountUi(Ribbon, { superdoc: first.superdoc, props: { activeTab: 'insert' } });
    await vi.advanceTimersByTimeAsync(150);
    pageBreakTracker.remember('block-1', true);
    await harness.wrapper.setProps({ collapsed: true });
    pageBreakTracker.forgetAllKeepingSnapshot();
    const second = documentWithSelection();
    await harness.setSuperdoc(second.superdoc);
    expect(pageBreakTracker.restoreSnapshot()).toBe(false);
    expect(second.observe).not.toHaveBeenCalled();
    expect(second.current).not.toHaveBeenCalled();
    await harness.wrapper.setProps({ collapsed: false });
    await settle();
    expect(pageBreakButton(harness).attributes('aria-pressed')).toBe('false');
    expect(second.current).toHaveBeenCalledTimes(1);
    expect(first.listeners.size).toBe(0);
  });

  it('Insert rejects a response arriving after collapse', async () => {
    const model = documentWithSelection();
    let resolveOld!: (value: ReturnType<typeof target>) => void;
    model.doc.selection = { current: () => new Promise(done => { resolveOld = done; }) };
    const harness = mountUi(Ribbon, { superdoc: model.superdoc, props: { activeTab: 'insert' } });
    await vi.advanceTimersByTimeAsync(150);
    await harness.wrapper.setProps({ collapsed: true });
    pageBreakTracker.remember('block-1', true);
    resolveOld(target());
    await settle();
    expect(pageBreakButton(harness).attributes('aria-pressed')).toBe('false');
    model.doc.selection = { current: model.current };
    await harness.wrapper.setProps({ collapsed: false });
    await settle();
    expect(pageBreakButton(harness).attributes('aria-pressed')).toBe('true');
  });

  it('Insert toggles the current paragraph while the reveal read is pending', async () => {
    const model = documentWithSelection();
    const harness = mountUi(Ribbon, { superdoc: model.superdoc, props: { activeTab: 'insert' } });
    await vi.advanceTimersByTimeAsync(150);
    pageBreakTracker.remember('block-1', true);
    await settle();
    expect(pageBreakButton(harness).attributes('aria-pressed')).toBe('true');
    await harness.wrapper.setProps({ collapsed: true });

    let resolveOld!: (value: ReturnType<typeof target>) => void;
    const current = vi.fn<() => ReturnType<typeof target> | Promise<ReturnType<typeof target>>>()
      .mockImplementationOnce(() => new Promise(done => { resolveOld = done; }))
      .mockImplementation(() => target('block-2'));
    model.doc.selection = { current };
    await harness.wrapper.setProps({ collapsed: false });
    expect(current).toHaveBeenCalledTimes(1);
    // The old paragraph's visible "on" state must not determine this write.
    await pageBreakButton(harness).trigger('click');
    await settle();
    expect(model.superdoc.inputs('format.paragraph.setFlowOptions')).toEqual([
      expect.objectContaining({ target: expect.objectContaining({ nodeId: 'block-2' }), pageBreakBefore: true }),
    ]);
    await harness.wrapper.setProps({ collapsed: true });
    pageBreakTracker.forgetAll();
    resolveOld(target());
    await settle();
    expect(pageBreakButton(harness).attributes('aria-pressed')).toBe('true');
    await harness.wrapper.setProps({ collapsed: false });
    await settle();
    expect(pageBreakButton(harness).attributes('aria-pressed')).toBe('false');
    expect(harness.failures()).toEqual([]);
  });

  it('preserves an open paragraph dialog and its unsubmitted input across collapse', async () => {
    const harness = mountUi(Ribbon);
    await settle();
    await buttonByTip(harness.wrapper, 'תפריט פסקה').trigger('click');
    await settle();
    const input = document.querySelector<HTMLInputElement>('.para-dialog input[type="number"]')!;
    expect(input).not.toBeNull();
    input.value = '3.5';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    await settle();
    await harness.wrapper.setProps({ collapsed: true });
    expect(document.querySelector('.para-dialog input[type="number"]')).toBe(input);
    expect(input.value).toBe('3.5');
    await harness.wrapper.setProps({ collapsed: false });
    await settle();
    expect(document.querySelector('.para-dialog input[type="number"]')).toBe(input);
    expect(input.value).toBe('3.5');
  });

  it('keeps command state and the shared context-menu font state live while collapsed', async () => {
    const adapter = createCommandDouble();
    const fontMemory = createFontMemory();
    const ribbon = mountUi(Ribbon, { adapter, fontMemory, props: { collapsed: true } });
    const sections = contextMenuModel({ hasDocument: true, hasRange: true, storyType: 'body', misspelledWord: null, can: () => true });
    const menu = mountUi(ContextMenu, { adapter, fontMemory, props: { open: true, point: { x: 200, y: 200 }, sections } });
    adapter.setState('bold', { active: true });
    await settle();
    expect(buttonByTip(ribbon.wrapper, 'מודגש').attributes('aria-pressed')).toBe('true');
    await menu.wrapper.find('[data-entry-id="bold"]').trigger('click');
    expect(adapter.applied.map(call => call.id)).toContain('bold');
    await setPicker(menu.wrapper, 'גופן', 'TaameyDavidCLM');
    await settle();
    expect(pickerValue(ribbon.wrapper, 'גופן')).toBe('TaameyDavidCLM');
    await ribbon.wrapper.setProps({ collapsed: false });
    await settle();
    expect(pickerValue(ribbon.wrapper, 'גופן')).toBe('TaameyDavidCLM');
    expect(menu.failures()).toEqual([]);
  });
});
