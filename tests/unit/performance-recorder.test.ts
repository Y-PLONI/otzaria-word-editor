import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { installPerformanceRecorder, type Recorder } from '../../scripts/performance-recorder.mjs';
import { beginPerformanceApi, beginPerformanceSpan, markPerformance } from '../../src/diagnostics/performance';

let recorder: Recorder | undefined;
let frames: Map<number, FrameRequestCallback>;
let nextFrame: number;
function paintOpportunity() {
  for (let tick = 0; tick < 2; tick += 1) {
    const pending = [...frames.values()]; frames.clear();
    for (const callback of pending) callback(performance.now());
  }
}
function fixture(initial = '') {
  let model = initial, offset = initial.length;
  const disposers: Array<() => void> = [];
  const container = document.createElement('div');
  container.innerHTML = '<div class="superdoc-fragment" data-source-node-id="block"><span class="superdoc-text-run"></span></div><textarea></textarea>';
  document.body.append(container);
  const run = container.querySelector('span')!;
  run.textContent = initial;
  const doc = {
    getNodeById: vi.fn(async () => ({ address: { nodeId: 'block' }, node: { kind: 'paragraph', paragraph: { inlines: [{ kind: 'run', run: { text: model } }] } } })),
    selection: { current: vi.fn(async () => ({ empty: true, target: { segments: [{ blockId: 'block', range: { start: offset, end: offset } }] } })) },
  };
  const editor = { container, superdoc: { activeEditor: { doc } },
    ui: { selection: { getSnapshot: () => ({ selectionTarget: {
      story: { storyType: 'body' },
      start: { kind: 'text', blockId: 'block', offset }, end: { kind: 'text', blockId: 'block', offset },
    } }) } }, onDispose: (dispose: () => void) => disposers.push(dispose) };
  Reflect.set(window, '__otzariaEditor', editor);
  Reflect.set(window, '__otzariaEditors', new Map());
  return { editor, container, doc, run, disposers,
    down(key: string, target = container.querySelector('textarea')!) {
      target.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true }));
    },
    async render(text: string, canonical = text) {
      model = canonical; offset = canonical.length; run.textContent = text;
      await Promise.resolve(); paintOpportunity();
    },
  };
}
beforeEach(() => {
  frames = new Map(); nextFrame = 0;
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => { const id = ++nextFrame; frames.set(id, callback); return id; });
  vi.stubGlobal('cancelAnimationFrame', (id: number) => frames.delete(id));
  vi.stubGlobal('getComputedStyle', () => ({ display: 'block', visibility: 'visible', opacity: '1' }));
  vi.stubGlobal('CSS', { escape: (value: string) => value });
  vi.spyOn(document, 'visibilityState', 'get').mockReturnValue('visible');
  vi.spyOn(Element.prototype, 'getBoundingClientRect').mockReturnValue({ x: 1, y: 1, top: 1, left: 1, bottom: 21, right: 101, width: 100, height: 20, toJSON() {} });
});
afterEach(() => {
  recorder?.stop(); recorder = undefined;
  Reflect.deleteProperty(window, '__otzariaPerformance');
  Reflect.deleteProperty(window, '__otzariaEditor');
  Reflect.deleteProperty(window, '__otzariaEditors');
  document.body.innerHTML = '';
  vi.restoreAllMocks(); vi.unstubAllGlobals();
});

describe('optional production hooks', () => {
  it('does not serialize payloads or install listeners when disabled', () => {
    const serialize = vi.fn(); const listener = vi.spyOn(document, 'addEventListener');
    expect(beginPerformanceApi('host', 'storage.set', { toJSON: serialize })).toBeUndefined();
    expect(beginPerformanceSpan('document.open')).toBeUndefined();
    markPerformance('test');
    expect(serialize).not.toHaveBeenCalled(); expect(listener).not.toHaveBeenCalled();
  });
  it('contains recorder failures, including failures after a span has begun', () => {
    const fail = () => { throw new Error('diagnostics failed'); };
    window.__otzariaPerformance = { mark: fail, api: () => fail, begin: () => ({ mark: fail, end: fail, bind: fail }) };
    expect(() => { markPerformance('test'); const span = beginPerformanceSpan('open'); span?.mark('step'); span?.bind({}); span?.end('done'); beginPerformanceApi('host', 'x', {})?.({}, 'ok'); }).not.toThrow();
  });
});

describe('document-validated typing recorder', () => {
  it('keeps every coalesced key, validates the block/caret, and exports no document contents', async () => {
    const f = fixture('private document'); recorder = installPerformanceRecorder({ apiBytes: true });
    const span = recorder.begin('document.open'); span.bind(f.editor); span.end('opened');
    await recorder.prepare('continuous', 'אב'); f.down('א'); f.down('ב');
    await f.render('private documentאב');
    expect(await recorder.finishBurst()).toEqual({ valid: true, received: 2, expected: 2, displayed: 2 });
    const result = recorder.snapshot();
    expect(result.keys).toHaveLength(2);
    expect(result.keys.every(key => key.coalesced && key.validated && key.renderOpportunityAt !== null)).toBe(true);
    expect(result.firstEdit).toMatchObject({ openingId: 1, documentApiValidated: true });
    expect(result.api.filter(row => row.scope === 'probe').reduce((n, row) => n + row.calls, 0)).toBe(3);
    expect(JSON.stringify(result)).not.toContain('private document');
    expect(JSON.stringify(result)).not.toContain('אב');
  });
  it('rejects a DOM-only change that did not reach the canonical document', async () => {
    const f = fixture(); recorder = installPerformanceRecorder();
    await recorder.prepare('first-edit', 'א'); f.down('א'); await f.render('א', '');
    expect((await recorder.finishBurst()).valid).toBe(false);
    expect(recorder.snapshot().firstEdit).toBeNull();
  });
  it('rejects duplicate/split visible fragments instead of counting their text twice', async () => {
    const f = fixture(); recorder = installPerformanceRecorder();
    f.container.append(f.container.firstElementChild!.cloneNode(true));
    await expect(recorder.prepare('first-edit', 'א')).rejects.toThrow('full sampled paragraph');
    expect(recorder.snapshot().keys).toHaveLength(0);
  });
  it('does not accept a paragraph clipped by the viewport', async () => {
    const f = fixture(); recorder = installPerformanceRecorder();
    const fragment = f.container.firstElementChild!;
    const bounds = fragment.getBoundingClientRect();
    vi.spyOn(fragment, 'getBoundingClientRect').mockReturnValue({ ...bounds, top: -1 });
    await expect(recorder.prepare('first-edit', 'א')).rejects.toThrow('full sampled paragraph');
  });
  it('rejects a document switch during an awaited preparation read', async () => {
    fixture(); recorder = installPerformanceRecorder();
    const prepare = recorder.prepare('first-edit', 'א');
    Reflect.set(window, '__otzariaEditor', {});
    await expect(prepare).rejects.toThrow('active document or caret changed');
  });
  it('does not attribute input from another element and rejects a switch after rendering', async () => {
    const f = fixture(); recorder = installPerformanceRecorder();
    await recorder.prepare('continuous', 'א');
    document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'א', bubbles: true }));
    expect(recorder.pending()?.received).toBe(0);
    f.down('א'); await f.render('א'); Reflect.set(window, '__otzariaEditor', {});
    expect((await recorder.finishBurst()).valid).toBe(false);
  });
  it('does not assign an earlier frame opportunity to a key arriving between callbacks', async () => {
    const f = fixture(); recorder = installPerformanceRecorder();
    await recorder.prepare('continuous', 'אב'); f.down('א'); f.run.textContent = 'א'; await Promise.resolve();
    const firstFrame = [...frames.values()]; frames.clear();
    for (const callback of firstFrame) callback(performance.now());
    f.down('ב'); f.run.textContent = 'אב'; await Promise.resolve();
    const secondFrame = [...frames.values()]; frames.clear();
    for (const callback of secondFrame) callback(performance.now());
    expect(recorder.pending()).toEqual({ received: 2, displayed: 1 });
    await f.render('אב');
    expect((await recorder.finishBurst()).valid).toBe(true);
  });
  it('rejects a final header selection even if its block and offset match the body', async () => {
    const f = fixture(); const selectionCurrent = f.doc.selection.current;
    recorder = installPerformanceRecorder();
    await recorder.prepare('first-edit', 'א'); f.down('א'); await f.render('א');
    const selection = { empty: true, selectionTarget: null, target: {
      story: { kind: 'story', storyType: 'header' },
      segments: [{ blockId: 'block', range: { start: 1, end: 1 } }],
    } };
    selectionCurrent.mockResolvedValue(selection);
    expect((await recorder.finishBurst()).valid).toBe(false);
    expect(recorder.snapshot().firstEdit).toBeNull();
  });
  it('invalidates hidden recordings and cancels pending rendering work on stop', async () => {
    const f = fixture(); recorder = installPerformanceRecorder();
    await recorder.prepare('continuous', 'א'); f.down('א');
    document.dispatchEvent(new Event('visibilitychange'));
    await f.render('א'); expect((await recorder.finishBurst()).valid).toBe(false);
    await recorder.prepare('continuous', 'ב'); f.down('ב'); f.run.textContent = 'אב'; await Promise.resolve();
    expect(frames.size).toBeGreaterThan(0); recorder.stop(); expect(frames.size).toBe(0);
  });
  it('preserves the receiver and original promise; restores methods when disposed', async () => {
    const f = fixture(); let receiver: unknown;
    const pending = Promise.resolve({ secret: 'not exported' });
    const original = function (this: unknown) { receiver = this; return pending; };
    const doc = f.doc as typeof f.doc & { get: typeof original }; doc.get = original;
    recorder = installPerformanceRecorder(); recorder.bind(f.editor);
    expect(doc.get()).toBe(pending); expect(receiver).toBe(doc); await pending;
    expect(recorder.snapshot().api.find(row => row.method === 'get')).toMatchObject({ calls: 1, completed: 1, responseBytes: null });
    for (const dispose of f.disposers) dispose(); expect(doc.get).toBe(original);
  });
  it('bounds exported events and makes truncated measurements detectable', () => {
    recorder = installPerformanceRecorder({ limit: 2 });
    for (let i = 0; i < 10; i += 1) recorder.mark('phase');
    expect(recorder.snapshot().events).toHaveLength(2);
    expect(recorder.snapshot().dropped).toBe(9);
  });
});
