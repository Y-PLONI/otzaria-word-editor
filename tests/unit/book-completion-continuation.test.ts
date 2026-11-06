import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { installBookCompletion, type AnchorRectLike, type BookCompletionHandle } from '../../src/engine/book-completion-overlay';
import type { SelectionTargetLike } from '../../src/engine/word-selection';

const reader = vi.hoisted(() => ({
  sectionGate: null as null | ((bookId: string, sectionIndex: number) => Promise<void>),
  state: {} as Record<string, unknown>,
  text: {} as Record<string, string>,
  staticMatch: vi.fn(() => null as { text: string } | null),
  acronymLookup: vi.fn(() => null as string | null),
}));

vi.mock('../../src/host/otzaria-reader', async (original) => ({
  ...await original<typeof import('../../src/host/otzaria-reader')>(),
  getCurrentReaderState: async () => ({ ok: true, value: reader.state }),
  getSectionTextMap: async ({ bookId, sectionIndex }: { bookId: string; sectionIndex: number }) => {
    if (reader.sectionGate) await reader.sectionGate(bookId, sectionIndex);
    const text = sectionIndex === 0 ? reader.text[bookId] ?? '' : '';
    return { ok: true, value: {
      renderedText: text, sourceText: text, renderedTextHash: text,
      hasMore: false, nextCursor: null,
    } };
  },
}));
vi.mock('../../src/host/otzaria-client', () => ({ on: () => () => {} }));
vi.mock('../../src/engine/static-completion-dictionary', () => ({ loadStaticSources: async () => [] }));
vi.mock('../../src/engine/static-completion', () => ({ matchStaticCompletion: reader.staticMatch }));
vi.mock('../../src/engine/acronym-dictionary', () => ({
  loadAcronymDictionary: async () => ({ lookup: reader.acronymLookup }),
}));

const PRIMARY = 'common1 common2 common3 common4 Aunique Aend';
const BACKGROUND = 'unique seed common1 common2 common3 common4 Bunique Bend';
const RECT = { left: 100, top: 100, width: 1, height: 20 };
let container: HTMLElement;
let handle: BookCompletionHandle | undefined;

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function fakeDoc(initialText = 'unique s') {
  const state = {
    text: initialText,
    caret: initialText.length,
    blockId: 'p',
    story: null as unknown,
    rect: RECT as AnchorRectLike | null,
    insertReceipt: undefined as Promise<{ success: boolean }> | undefined,
    moveCaretOnInsert: 'replacement' as 'replacement' | 'always' | 'never',
  };
  const doc = {
    selection: { current: vi.fn(async () => ({ selectionTarget: {
      kind: 'selection',
      start: { kind: 'text', blockId: state.blockId, offset: state.caret },
      end: { kind: 'text', blockId: state.blockId, offset: state.caret },
      story: state.story,
    } })) },
    ranges: { resolve: vi.fn(async (request: unknown) => {
      const range = request as { start: { point: { offset: number } }; end: { point: { offset: number } } };
      const start = Math.min(state.text.length, range.start.point.offset);
      const end = Math.min(state.text.length, range.end.point.offset);
      return { preview: { text: state.text.slice(start, end) }, target: { start: { offset: start }, end: { offset: end } } };
    }) },
    insert: vi.fn(async ({ value, target }: { value: string; target?: unknown }) => {
      const range = target as SelectionTargetLike;
      const start = range.start!.offset!;
      const end = range.end!.offset!;
      state.text = state.text.slice(0, start) + value + state.text.slice(end);
      // The engine maps the end of a replacement, but a caret at a collapsed
      // insertion boundary stays before the inserted text until selection.apply.
      if (state.moveCaretOnInsert === 'always' || (state.moveCaretOnInsert === 'replacement' && start !== end)) {
        state.caret = start + value.length;
      }
      return await (state.insertReceipt ?? Promise.resolve({ success: true }));
    }),
  };
  const apply = vi.fn((target: unknown) => {
    const range = target as SelectionTargetLike;
    state.caret = range.end!.offset!;
    state.blockId = range.end!.blockId!;
    state.story = range.story ?? range.end!.story ?? null;
    return { ok: true };
  });
  const host = {
    activeEditor: { doc },
    ui: { selection: { getAnchorRect: () => state.rect, apply } },
  };
  return { state, doc, host, apply };
}

beforeEach(() => {
  vi.useFakeTimers();
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) =>
    window.setTimeout(() => callback(performance.now()), 16));
  vi.setSystemTime(new Date(10_000));
  reader.sectionGate = null;
  reader.text = { primary: PRIMARY, background: BACKGROUND };
  reader.staticMatch.mockReset().mockReturnValue(null);
  reader.acronymLookup.mockReset().mockReturnValue(null);
  reader.state = {
    currentBook: 'Primary Book', currentBookId: 'primary', bookUid: 'primary',
    currentType: 'text', currentIndex: 0, currentRef: null,
    openTabs: [{ toolId: null, isSelf: false, type: 'text', bookId: 'background', bookUid: 'background', index: 0 }],
  };
  container = document.createElement('div');
  document.body.appendChild(container);
  Object.defineProperty(document, 'elementFromPoint', { configurable: true, value: () => null });
});

afterEach(() => {
  handle?.dispose();
  handle = undefined;
  document.body.innerHTML = '';
  vi.restoreAllMocks();
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

const ghost = () => document.querySelector('.otzaria-book-completion-ghost')?.textContent ?? '';
const keydown = (key: string) => {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true });
  container.dispatchEvent(event);
  return event;
};
const keyup = (key: string) => container.dispatchEvent(new KeyboardEvent('keyup', { key, bubbles: true }));

async function start(initialText?: string) {
  const fixture = fakeDoc(initialText);
  handle = installBookCompletion(container, fixture.host);
  await vi.advanceTimersByTimeAsync(1);
  container.dispatchEvent(new Event('input'));
  await vi.advanceTimersByTimeAsync(200);
  return fixture;
}

async function tab() {
  expect(keydown('Tab').defaultPrevented).toBe(true);
  keyup('Tab');
  await vi.advanceTimersByTimeAsync(200);
}

describe('book completion continuation', () => {
  it('keeps the selected background source across Tab keyup and repeated Tab', async () => {
    const { state, apply } = await start();
    expect(ghost()).toBe('seed common1 common2 common3 common4');

    await tab();
    expect(state.text).toBe('unique seed common1 common2 common3 common4');
    expect(ghost()).toBe(' Bunique Bend');

    await tab();
    expect(state.text).toBe(BACKGROUND);
    expect(state.caret).toBe(BACKGROUND.length);
    expect(apply).toHaveBeenCalledTimes(1);
    expect(ghost()).toBe('');

    container.dispatchEvent(new Event('beforeinput'));
    state.text = state.text.slice(0, state.caret) + ' typed' + state.text.slice(state.caret);
    state.caret += ' typed'.length;
    container.dispatchEvent(new Event('input'));
    expect(state.text).toBe(`${BACKGROUND} typed`);
  });

  it('keeps the selected source after an unchanged reader window is refreshed', async () => {
    const { state } = await start();
    await tab();
    expect(ghost()).toBe(' Bunique Bend');

    await vi.advanceTimersByTimeAsync(31_000);
    keyup('Shift');
    await vi.advanceTimersByTimeAsync(1_000);
    expect(ghost()).toBe(' Bunique Bend');

    await tab();
    expect(state.text).toBe(BACKGROUND);
  });

  it('keeps the source when an unchanged refresh finishes during a pending validation', async () => {
    const { state } = await start();
    await tab();
    expect(ghost()).toBe(' Bunique Bend');
    await vi.advanceTimersByTimeAsync(31_000);

    const validation = deferred<void>();
    let backgroundReads = 0;
    reader.sectionGate = async (bookId, index) => {
      if (bookId === 'background' && index === 0) {
        backgroundReads += 1;
        // The continuation check waits while the independent TTL load publishes
        // the same text with a new window object. It still has the same source.
        if (backgroundReads === 1) await validation.promise;
      }
    };
    keyup('Shift');
    await vi.advanceTimersByTimeAsync(150);
    expect(backgroundReads).toBe(2);

    validation.resolve();
    await vi.advanceTimersByTimeAsync(200);
    expect(ghost()).toBe(' Bunique Bend');
    await tab();
    expect(state.text).toBe(BACKGROUND);
  });

  it('repaints the same continuation when the engine anchor is temporarily unavailable', async () => {
    const { state } = await start();
    state.rect = null;
    keydown('Tab');
    keyup('Tab');
    await vi.advanceTimersByTimeAsync(30);
    expect(ghost()).toBe('');

    state.rect = RECT;
    await vi.advanceTimersByTimeAsync(200);
    expect(ghost()).toBe(' Bunique Bend');
  });

  it('does not consume Tab for a continuation that has not been painted', async () => {
    const { state, doc } = await start();
    state.rect = null;
    keydown('Tab');
    keyup('Tab');
    await vi.advanceTimersByTimeAsync(30);
    expect(ghost()).toBe('');

    expect(keydown('Tab').defaultPrevented).toBe(false);
    expect(doc.insert).toHaveBeenCalledTimes(1);
  });

  it('restarts matching after typing instead of retaining the previous source', async () => {
    const { state } = await start();
    await tab();
    container.dispatchEvent(new Event('beforeinput'));
    state.text = 'common1 c';
    state.caret = state.text.length;
    container.dispatchEvent(new Event('input'));
    await vi.advanceTimersByTimeAsync(200);
    expect(ghost()).toBe('common2 common3 common4 Aunique Aend');

    await tab();
    expect(state.text).toBe(PRIMARY);
  });

  it('clears the continuation as soon as the user points elsewhere', async () => {
    await start();
    await tab();
    container.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    expect(ghost()).toBe('');
    expect(keydown('Tab').defaultPrevented).toBe(false);
  });

  it.each(['beforeinput', 'pointerdown'])('does not resurrect continuation after %s during an insertion', async (event) => {
    const { state } = await start();
    const receipt = deferred<{ success: boolean }>();
    state.insertReceipt = receipt.promise;
    expect(keydown('Tab').defaultPrevented).toBe(true);
    keyup('Tab');
    await vi.advanceTimersByTimeAsync(0);
    expect(state.text).toBe('unique seed common1 common2 common3 common4');
    container.dispatchEvent(new Event(event, { bubbles: true }));
    state.caret = 0;

    receipt.resolve({ success: true });
    await vi.advanceTimersByTimeAsync(30);
    expect(ghost()).toBe('');
    expect(keydown('Tab').defaultPrevented).toBe(false);
  });

  it('invalidates an in-flight evaluation even when no suggestion is visible yet', async () => {
    const { state, doc } = await start('');
    const pending = deferred<Awaited<ReturnType<typeof doc.ranges.resolve>>>();
    doc.ranges.resolve.mockImplementationOnce(() => pending.promise);
    state.text = 'unique s';
    state.caret = state.text.length;
    container.dispatchEvent(new Event('input'));
    await vi.advanceTimersByTimeAsync(150);

    container.dispatchEvent(new Event('beforeinput'));
    state.text = '';
    state.caret = 0;
    pending.resolve({ preview: { text: 'unique s' }, target: { start: { offset: 0 }, end: { offset: 8 } } });
    await vi.advanceTimersByTimeAsync(0);
    expect(ghost()).toBe('');
  });

  it('does not let an older evaluation overwrite a newly accepted continuation', async () => {
    const { doc } = await start();
    const pending = deferred<Awaited<ReturnType<typeof doc.ranges.resolve>>>();
    doc.ranges.resolve.mockImplementationOnce(() => pending.promise);
    keyup('Shift');
    await vi.advanceTimersByTimeAsync(150);
    expect(keydown('Tab').defaultPrevented).toBe(true);
    keyup('Tab');
    await vi.advanceTimersByTimeAsync(30);
    expect(doc.insert).toHaveBeenCalledTimes(1);
    expect(ghost()).toBe(' Bunique Bend');

    pending.resolve({ preview: { text: 'unique s' }, target: { start: { offset: 0 }, end: { offset: 8 } } });
    await vi.advanceTimersByTimeAsync(0);
    expect(ghost()).toBe(' Bunique Bend');
  });

  it('validates the caret before repainting a continuation after a programmatic move', async () => {
    const { state } = await start();
    await tab();
    state.caret = 0;
    keyup('Shift');
    await vi.advanceTimersByTimeAsync(200);
    expect(ghost()).toBe('');
    expect(keydown('Tab').defaultPrevented).toBe(false);
  });

  it('does not insert a suggestion whose source changed after it was displayed', async () => {
    const { state, doc } = await start();
    reader.text.background = 'unique seed replacement text';

    expect(keydown('Tab').defaultPrevented).toBe(true);
    keyup('Tab');
    await vi.advanceTimersByTimeAsync(200);
    expect(doc.insert).not.toHaveBeenCalled();
    expect(state.text).toBe('unique s');
    expect(ghost()).toBe('seed replacement text');
  });

  it('does not insert after the text changes at the same caret without an input event', async () => {
    const { state, doc } = await start();
    state.text = 'changed!';
    expect(state.text.length).toBe(state.caret);

    expect(keydown('Tab').defaultPrevented).toBe(true);
    await vi.advanceTimersByTimeAsync(200);
    expect(doc.insert).not.toHaveBeenCalled();
    expect(state.text).toBe('changed!');
    expect(ghost()).toBe('');
  });

  it.each(['beforeinput', 'pointerdown'])('cancels acceptance if %s happens during its document check', async (event) => {
    const { doc } = await start();
    const pending = deferred<Awaited<ReturnType<typeof doc.ranges.resolve>>>();
    doc.ranges.resolve.mockImplementationOnce(() => pending.promise);

    expect(keydown('Tab').defaultPrevented).toBe(true);
    await vi.advanceTimersByTimeAsync(0);
    container.dispatchEvent(new Event(event, { bubbles: true }));
    pending.resolve({ preview: { text: 'unique s' }, target: { start: { offset: 0 }, end: { offset: 8 } } });
    await vi.advanceTimersByTimeAsync(200);
    expect(doc.insert).not.toHaveBeenCalled();
    expect(ghost()).toBe('');
  });

  it('does not insert at an old caret if it moves during the document check', async () => {
    const { state, doc } = await start();
    const pending = deferred<Awaited<ReturnType<typeof doc.ranges.resolve>>>();
    doc.ranges.resolve.mockImplementationOnce(() => pending.promise);

    expect(keydown('Tab').defaultPrevented).toBe(true);
    await vi.advanceTimersByTimeAsync(0);
    state.caret = 0;
    pending.resolve({ preview: { text: 'unique s' }, target: { start: { offset: 0 }, end: { offset: 8 } } });
    await vi.advanceTimersByTimeAsync(200);
    expect(doc.insert).not.toHaveBeenCalled();
    expect(state.text).toBe('unique s');
    expect(ghost()).toBe('');
  });

  it('does not reoffer a non-book completion on Tab keyup', async () => {
    reader.staticMatch.mockReturnValue({ text: 'known phrase' });
    const { state } = await start('known p');
    expect(ghost()).toBe('known phrase');
    await tab();
    expect(state.text).toBe('known known phrase');
    expect(ghost()).toBe('');
  });
});

describe('completion caret placement', () => {
  it('places the caret after a static completion even when insertion preserves it', async () => {
    reader.staticMatch.mockReturnValue({ text: 'known phrase' });
    const { state, apply } = await start('known p');
    state.moveCaretOnInsert = 'never';

    await tab();
    expect(state.text).toBe('known known phrase');
    expect(state.caret).toBe(state.text.length);
    expect(apply).toHaveBeenCalledTimes(1);
    expect(ghost()).toBe('');
  });

  it('places the caret after an acronym expansion at a collapsed target', async () => {
    reader.acronymLookup.mockReturnValue('חכמינו זכרונם לברכה');
    const { state, apply } = await start('חז"ל ');
    expect(ghost()).toBe('חכמינו זכרונם לברכה');

    await tab();
    expect(state.text).toBe('חז"ל חכמינו זכרונם לברכה');
    expect(state.caret).toBe(state.text.length);
    expect(apply).toHaveBeenCalledTimes(1);
    expect(ghost()).toBe('');
  });

  it('retries temporary layout failures before showing the continuation', async () => {
    const { state, apply } = await start();
    state.moveCaretOnInsert = 'never';
    apply.mockReturnValueOnce({ ok: false }).mockReturnValueOnce({ ok: false });
    expect(keydown('Tab').defaultPrevented).toBe(true);
    keyup('Tab');
    await vi.advanceTimersByTimeAsync(0);
    expect(apply).toHaveBeenCalledTimes(1);
    expect(state.caret).toBe(8);
    expect(ghost()).toBe('');

    await vi.advanceTimersByTimeAsync(200);
    expect(apply).toHaveBeenCalledTimes(3);
    expect(state.caret).toBe(state.text.length);
    expect(ghost()).toBe(' Bunique Bend');
  });

  it('bounds retries and leaves no live continuation after permanent layout failure', async () => {
    const { state, apply } = await start();
    state.moveCaretOnInsert = 'never';
    apply.mockReturnValue({ ok: false });
    expect(keydown('Tab').defaultPrevented).toBe(true);
    keyup('Tab');
    await vi.advanceTimersByTimeAsync(1_000);
    expect(apply).toHaveBeenCalledTimes(10);
    expect(state.caret).toBe(8);
    expect(ghost()).toBe('');
    expect(keydown('Tab').defaultPrevented).toBe(false);
    await vi.advanceTimersByTimeAsync(1_000);
    expect(apply).toHaveBeenCalledTimes(10);
  });

  it('clears continuation if the public selection API throws', async () => {
    const { state, apply } = await start();
    state.moveCaretOnInsert = 'never';
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => {});
    apply.mockImplementation(() => { throw new Error('selection unavailable'); });

    await tab();
    expect(apply).toHaveBeenCalledTimes(1);
    expect(warning).toHaveBeenCalledTimes(1);
    expect(state.caret).toBe(8);
    expect(ghost()).toBe('');
    expect(keydown('Tab').defaultPrevented).toBe(false);
  });

  it('does not claim continuation without apply when insertion leaves the caret behind', async () => {
    const { state, host, apply } = await start();
    state.moveCaretOnInsert = 'never';
    Reflect.deleteProperty(host.ui.selection, 'apply');

    await tab();
    expect(state.text).toBe('unique seed common1 common2 common3 common4');
    expect(state.caret).toBe(8);
    expect(apply).not.toHaveBeenCalled();
    expect(ghost()).toBe('');
    expect(keydown('Tab').defaultPrevented).toBe(false);
  });

  it('supports hosts without apply when insertion already places the caret correctly', async () => {
    const { state, host, apply } = await start();
    state.moveCaretOnInsert = 'always';
    Reflect.deleteProperty(host.ui.selection, 'apply');

    await tab();
    expect(ghost()).toBe(' Bunique Bend');
    await tab();
    expect(state.text).toBe(BACKGROUND);
    expect(state.caret).toBe(BACKGROUND.length);
    expect(apply).not.toHaveBeenCalled();
    expect(ghost()).toBe('');
  });

  it.each(['failed receipt', 'exception'])('does not move the caret after an insertion %s', async (failure) => {
    const { state, doc, apply } = await start();
    vi.spyOn(console, 'warn').mockImplementation(() => {});
    if (failure === 'failed receipt') doc.insert.mockResolvedValueOnce({ success: false });
    else doc.insert.mockRejectedValueOnce(new Error('insertion failed'));

    await tab();
    expect(state.text).toBe('unique s');
    expect(state.caret).toBe(8);
    expect(apply).not.toHaveBeenCalled();
    expect(ghost()).toBe('');
  });

  it.each(['beforeinput', 'pointerdown', 'dispose'])('stops caret retries after %s', async (event) => {
    const { state, apply } = await start();
    state.moveCaretOnInsert = 'never';
    apply.mockReturnValueOnce({ ok: false });
    expect(keydown('Tab').defaultPrevented).toBe(true);
    keyup('Tab');
    await vi.advanceTimersByTimeAsync(0);
    expect(apply).toHaveBeenCalledTimes(1);

    if (event === 'dispose') handle?.dispose();
    else container.dispatchEvent(new Event(event, { bubbles: true }));
    state.caret = 0;
    await vi.advanceTimersByTimeAsync(1_000);
    expect(apply).toHaveBeenCalledTimes(1);
    expect(state.caret).toBe(0);
    expect(ghost()).toBe('');
  });

  it.each(['offset', 'block', 'story'])('does not steal a changed %s while insertion is pending', async (change) => {
    const { state, apply } = await start();
    state.moveCaretOnInsert = 'never';
    const receipt = deferred<{ success: boolean }>();
    state.insertReceipt = receipt.promise;
    expect(keydown('Tab').defaultPrevented).toBe(true);
    keyup('Tab');
    await vi.advanceTimersByTimeAsync(0);

    if (change === 'offset') state.caret = 0;
    else if (change === 'block') state.blockId = 'another paragraph';
    else state.story = { storyType: 'footnote', noteId: 'another story' };
    const selected = { caret: state.caret, blockId: state.blockId, story: state.story };
    receipt.resolve({ success: true });
    await vi.advanceTimersByTimeAsync(200);
    expect(apply).not.toHaveBeenCalled();
    expect({ caret: state.caret, blockId: state.blockId, story: state.story }).toEqual(selected);
    expect(ghost()).toBe('');
  });

  it.each(['offset', 'block', 'story'])('rechecks a changed %s before retrying caret placement', async (change) => {
    const { state, apply } = await start();
    state.moveCaretOnInsert = 'never';
    apply.mockReturnValueOnce({ ok: false });
    expect(keydown('Tab').defaultPrevented).toBe(true);
    keyup('Tab');
    await vi.advanceTimersByTimeAsync(0);
    expect(apply).toHaveBeenCalledTimes(1);

    if (change === 'offset') state.caret = 0;
    else if (change === 'block') state.blockId = 'another paragraph';
    else state.story = { storyType: 'footnote', noteId: 'another story' };
    const selected = { caret: state.caret, blockId: state.blockId, story: state.story };
    await vi.advanceTimersByTimeAsync(1_000);
    expect(apply).toHaveBeenCalledTimes(1);
    expect({ caret: state.caret, blockId: state.blockId, story: state.story }).toEqual(selected);
    expect(ghost()).toBe('');
  });
});
