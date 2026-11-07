/**
 * Integration of the completion overlay with the reader contract. The editor is
 * a public-API double; assertions inspect the visible ghost or the source RPCs.
 * Static dictionaries are disabled so they cannot mask a missing book context.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  installBookCompletion,
  type BookCompletionHandle,
  type CompletionHost,
} from '../../src/engine/book-completion-overlay';
import type { ReaderResult } from '../../src/host/otzaria-reader';
import type {
  BookType,
  GetSectionTextMapArgs,
  ReaderState,
  ReaderSectionContentChangedEvent,
  SectionTextMapResult,
} from '../../src/types/otzaria_plugin';

const reader = vi.hoisted(() => ({
  getState: vi.fn<() => Promise<ReaderResult<ReaderState | null>>>(),
  getSection: vi.fn<(args: GetSectionTextMapArgs) => Promise<ReaderResult<SectionTextMapResult | null>>>(),
  listeners: new Map<string, (detail: Record<string, unknown>) => void>(),
}));

vi.mock('../../src/host/otzaria-reader', async (original) => ({
  ...await original<typeof import('../../src/host/otzaria-reader')>(),
  getCurrentReaderState: reader.getState,
  getSectionTextMap: reader.getSection,
}));
vi.mock('../../src/host/otzaria-client', () => ({
  on: (topic: string, listener: (detail: Record<string, unknown>) => void) => {
    reader.listeners.set(topic, listener);
    return () => reader.listeners.delete(topic);
  },
}));
vi.mock('../../src/engine/static-completion-dictionary', () => ({ loadStaticSources: async () => [] }));
vi.mock('../../src/engine/static-completion', () => ({ matchStaticCompletion: () => null }));
vi.mock('../../src/engine/acronym-dictionary', () => ({ loadAcronymDictionary: async () => null }));

const POINTED = 'אָבִי אָמַר דְּבָרִים יָפִים מְאֹד';
const PLAIN = 'אבי אמר דברים יפים מאד';
const PLAIN_GHOST = 'אמר דברים יפים מאד';
const READABLE_TYPES: BookType[] = [
  'text', 'epub', 'docx', 'docm', 'dotx', 'dotm', 'doc', 'dot', 'wbk', 'xml', 'rtf', 'odt',
  'html', 'htm', 'xhtml',
];

function bookTab(bookId: string, overrides: Partial<ReaderState['openTabs'][number]> = {}): ReaderState['openTabs'][number] {
  return {
    tabId: null,
    correctionSessionId: null,
    id: 1,
    toolId: null,
    isSelf: false,
    type: 'text',
    source: 'library',
    bookId,
    bookUid: bookId,
    book: bookId,
    index: 0,
    currentRef: null,
    ...overrides,
  };
}

function currentState(overrides: Partial<ReaderState> = {}): ReaderState {
  return {
    currentTabId: null,
    currentCorrectionSessionId: null,
    currentBook: 'Source Book',
    currentBookId: 'source',
    bookUid: 'source',
    currentId: 1,
    currentType: 'text',
    currentSource: 'library',
    currentIndex: 0,
    currentRef: null,
    openTabs: [],
    ...overrides,
  };
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

let handle: BookCompletionHandle | undefined;
let container: HTMLElement;
let state: ReaderState;
let typedText: string;
let sections: Map<string, Partial<SectionTextMapResult>>;
let onStatus: ReturnType<typeof vi.fn>;

function sectionKey(book: string, index: number, cursor = ''): string {
  return `${book}:${index}:${cursor}`;
}

function setSection(book: string, index: number, text: string, overrides: Partial<SectionTextMapResult> = {}, cursor = ''): void {
  sections.set(sectionKey(book, index, cursor), {
    renderedText: text,
    sourceText: text,
    renderedTextHash: text,
    sourceTextHash: text,
    ...overrides,
  });
}

function sectionResult(args: GetSectionTextMapArgs): ReaderResult<SectionTextMapResult> {
  return {
    ok: true,
    value: {
      schemaVersion: 1,
      bookId: args.bookId,
      sectionIndex: args.sectionIndex,
      currentRef: null,
      renderedText: '',
      sourceText: '',
      renderedTextHash: '',
      hasMore: false,
      ...sections.get(sectionKey(args.bookUid ?? args.bookId, args.sectionIndex, args.cursor)),
    },
  };
}

const requests = () => reader.getSection.mock.calls.map(([request]) => request);
const loadedBooks = () => new Set(requests().map((request) => request.bookUid ?? request.bookId));
const ghost = () => document.querySelector('.otzaria-book-completion-ghost')?.textContent ?? '';

beforeEach(() => {
  vi.useFakeTimers();
  vi.setSystemTime(new Date(10_000));
  reader.getState.mockReset();
  reader.getSection.mockReset();
  reader.listeners.clear();
  state = currentState();
  sections = new Map();
  typedText = 'אבי א';
  onStatus = vi.fn();
  // Each RPC sees a snapshot, as it would across the reader bridge.
  reader.getState.mockImplementation(async () => ({
    ok: true,
    value: { ...state, openTabs: state.openTabs.map((tab) => ({ ...tab })) },
  }));
  reader.getSection.mockImplementation(async (args) => sectionResult(args));
  container = document.createElement('div');
  document.body.appendChild(container);
  Object.defineProperty(document, 'elementFromPoint', { configurable: true, value: () => null });
});

afterEach(() => {
  handle?.dispose();
  handle = undefined;
  document.body.innerHTML = '';
  vi.useRealTimers();
});

async function mount(): Promise<void> {
  const host: CompletionHost = {
    activeEditor: {
      doc: {
        selection: {
          current: async () => ({ selectionTarget: {
            kind: 'selection',
            start: { kind: 'text', blockId: 'paragraph', offset: typedText.length },
            end: { kind: 'text', blockId: 'paragraph', offset: typedText.length },
          } }),
        },
        ranges: {
          resolve: async () => ({
            preview: { text: typedText },
            target: { start: { offset: 0 }, end: { offset: typedText.length } },
          }),
        },
      },
    },
    ui: { selection: { getAnchorRect: () => ({ left: 100, top: 100, width: 1, height: 20 }) } },
  };
  handle = installBookCompletion(container, host, { onStatus });
  await vi.advanceTimersByTimeAsync(1);
}

async function evaluate(): Promise<void> {
  container.dispatchEvent(new Event('input'));
  await vi.advanceTimersByTimeAsync(200);
}

async function start(): Promise<void> {
  await mount();
  await evaluate();
}

/** Polling refreshes in the background; the next input can use the new cache. */
async function refresh(): Promise<void> {
  await vi.advanceTimersByTimeAsync(1_300);
  await evaluate();
  await evaluate();
}

function bookChanged(): void {
  reader.listeners.get('reader.current_book_changed')?.({});
}

function sectionChanged(bookId: string, sectionIndex: number, overrides: Partial<ReaderSectionContentChangedEvent> = {}): void {
  const event: ReaderSectionContentChangedEvent = {
    schemaVersion: 1, bookId, sectionIndex,
    changeType: 'source-content', newSourceTextHash: 'updated-source',
    ...overrides,
  };
  reader.listeners.get('reader.sectionContentChanged')?.({ ...event });
}

describe('book completion: reader formats and tab selection', () => {
  it.each(READABLE_TYPES)('completes text from the current %s document', async (type) => {
    state.currentType = type as BookType;
    setSection('source', 0, POINTED);
    await start();
    expect(ghost()).toContain('אָמַר');
  });

  it.each(READABLE_TYPES)('uses a contextual match in a background %s document', async (type) => {
    state.openTabs = [bookTab('background', { type: type as BookType })];
    setSection('source', 0, 'טקסט אחר שאין בו התאמה');
    setSection('background', 0, POINTED);
    await start();
    expect(ghost()).toContain('אָמַר');
    expect(loadedBooks()).toContain('background');
  });

  it('supports an older host without currentType or tab.type', async () => {
    delete (state as Partial<ReaderState>).currentType;
    const legacyTab = bookTab('legacy');
    delete (legacyTab as Partial<typeof legacyTab>).type;
    state.openTabs = [legacyTab];
    setSection('source', 0, 'ראשית חכמה');
    setSection('legacy', 0, POINTED);
    await start();
    expect(loadedBooks()).toContain('source');
    expect(loadedBooks()).toContain('legacy');
    expect(ghost()).toContain('אָמַר');
  });

  it.each([null, 'pdf', 'external'] as const)('does not use the current %s tab as a text book', async (type) => {
    state.currentType = type;
    state.openTabs = [bookTab('background')];
    setSection('source', 0, POINTED);
    setSection('background', 0, POINTED);
    await start();
    expect(ghost()).toBe('');
    expect(reader.getSection).not.toHaveBeenCalled();
  });

  it('does not let search, PDF, external or tool tabs consume background-book slots', async () => {
    state.openTabs = [
      ...Array.from({ length: 5 }, (_, i) => bookTab(`search-${i}`, { type: null, id: null, bookUid: null })),
      bookTab('pdf', { type: 'pdf' }),
      bookTab('external', { type: 'external' }),
      bookTab('tool', { toolId: 'some-plugin' }),
      bookTab('self', { isSelf: true }),
      bookTab('background'),
    ];
    setSection('background', 0, POINTED);
    await start();
    expect(loadedBooks()).toEqual(new Set(['source', 'background']));
    expect(ghost()).toContain('אָמַר');
  });

  it('limits the context to six distinct books including the current book', async () => {
    state.openTabs = [
      bookTab('source'),
      bookTab('background-1'),
      bookTab('background-1'),
      ...Array.from({ length: 6 }, (_, i) => bookTab(`background-${i + 2}`)),
    ];
    setSection('background-5', 0, POINTED);
    await start();
    expect(loadedBooks()).toEqual(new Set([
      'source', 'background-1', 'background-2', 'background-3', 'background-4', 'background-5',
    ]));
    expect(ghost()).toContain('אָמַר');
    // The seventh book must neither be read nor supply a completion.
    setSection('background-5', 0, '');
    setSection('background-6', 0, 'יוסף סיפר סיפור ארוך');
    typedText = 'יוסף ס';
    await evaluate();
    expect(ghost()).toBe('');
    expect(loadedBooks()).not.toContain('background-6');
  });

  it('passes the stable bookUid when a tab title names an alternate edition', async () => {
    state.currentBookId = 'Source Book — alternate edition';
    state.bookUid = 'id:42';
    setSection('id:42', 0, POINTED);
    await start();
    expect(ghost()).toContain('אָמַר');
    expect(requests()).toContainEqual(expect.objectContaining({
      bookId: 'Source Book — alternate edition', bookUid: 'id:42',
    }));
  });

  it('never offers a background word from a partial word without context', async () => {
    state.openTabs = [bookTab('background')];
    typedText = 'אמ';
    setSection('source', 0, 'ראשית חכמה');
    setSection('background', 0, PLAIN);
    await start();
    expect(loadedBooks()).toContain('background');
    expect(ghost()).toBe('');
  });
});

describe('book completion: bounded context and display refresh', () => {
  it('keeps the current line usable when a preceding line fills the character cap', async () => {
    state.currentIndex = 5;
    setSection('source', 0, 'word '.repeat(4_000));
    setSection('source', 5, POINTED);
    await start();
    expect(requests()).toContainEqual(expect.objectContaining({ sectionIndex: 5 }));
    expect(ghost()).toContain('אָמַר');
  });

  it('does not reload an unchanged long window on every short reader poll', async () => {
    state.currentIndex = 5;
    setSection('source', 0, 'word '.repeat(4_000));
    setSection('source', 5, POINTED);
    await start();
    expect(ghost()).toContain('אָמַר');
    reader.getSection.mockClear();
    await refresh();
    await refresh();
    await refresh();
    // The matching center may be validated again; the huge preceding paragraph
    // must not force another context load every 1.2-second state poll.
    expect(requests().filter((request) => request.sectionIndex === 0)).toEqual([]);
    expect(ghost()).toContain('אָמַר');
  });

  it('keeps an unchanged long window cached when it contains no matching candidate', async () => {
    state.currentIndex = 5;
    setSection('source', 0, 'word '.repeat(4_000));
    setSection('source', 5, 'center remains unchanged');
    await start();
    expect(ghost()).toBe('');
    reader.getSection.mockClear();
    await refresh();
    await refresh();
    await refresh();
    expect(requests().filter((request) => request.sectionIndex === 0)).toEqual([]);
    expect(ghost()).toBe('');
  });

  it('reads only the local line window, never negative indexes or the whole book', async () => {
    state.currentIndex = 2;
    setSection('source', 2, POINTED);
    await start();
    expect(ghost()).toContain('אָמַר');
    expect(requests().every((request) => request.sectionIndex >= 0 && request.sectionIndex <= 27)).toBe(true);
  });

  it('does not offer a word beyond the context character cap', async () => {
    typedText = 'יחזקאל י';
    setSection('source', 0, `${'word '.repeat(4_000)}יחזקאל יאמר דברים ברורים`);
    await start();
    expect(ghost()).toBe('');
  });

  it('uses rendered text even when source text retains pointing', async () => {
    setSection('source', 0, PLAIN, { sourceText: POINTED, sourceTextHash: POINTED });
    await start();
    expect(ghost()).toBe(PLAIN_GHOST);
  });

  it('falls back to source text when an older host omits rendered text', async () => {
    setSection('source', 0, POINTED, { renderedText: undefined, renderedTextHash: undefined });
    await start();
    expect(ghost()).toContain('אָמַר');
  });

  it('refreshes the displayed text when the center line changes', async () => {
    setSection('source', 0, POINTED);
    await start();
    expect(ghost()).toContain('אָמַר');
    setSection('source', 0, PLAIN, { sourceText: POINTED, sourceTextHash: POINTED });
    await refresh();
    expect(ghost()).toBe(PLAIN_GHOST);
  });

  it('refreshes neighboring text when the center is an unchanged heading', async () => {
    setSection('source', 0, 'פרק א');
    setSection('source', 1, POINTED);
    await start();
    expect(ghost()).toContain('אָמַר');
    setSection('source', 1, PLAIN, { sourceText: POINTED, sourceTextHash: POINTED });
    await refresh();
    expect(ghost()).toBe(PLAIN_GHOST);
  });

  it('refreshes background display settings beside an unchanged heading', async () => {
    state.openTabs = [bookTab('background')];
    setSection('source', 0, 'ראשית חכמה');
    setSection('background', 0, 'פרק א');
    setSection('background', 1, POINTED);
    await start();
    expect(ghost()).toContain('אָמַר');
    setSection('background', 1, PLAIN, { sourceText: POINTED, sourceTextHash: POINTED });
    await refresh();
    expect(ghost()).toBe(PLAIN_GHOST);
  });

  it('loads a newly opened background book on the next reader poll', async () => {
    setSection('source', 0, 'ראשית חכמה');
    await start();
    expect(ghost()).toBe('');
    state.openTabs = [bookTab('background')];
    setSection('background', 0, POINTED);
    await refresh();
    expect(ghost()).toContain('אָמַר');
  });

  it('removes a closed background book from subsequent matches', async () => {
    state.openTabs = [bookTab('background')];
    setSection('background', 0, POINTED);
    await start();
    expect(ghost()).toContain('אָמַר');
    state.openTabs = [];
    await refresh();
    expect(ghost()).toBe('');
  });

  it('keeps the current book available while a background RPC is pending', async () => {
    const pending = deferred<ReaderResult<SectionTextMapResult | null>>();
    state.openTabs = [bookTab('background')];
    setSection('source', 0, POINTED);
    reader.getSection.mockImplementation((args) => args.bookUid === 'background'
      ? pending.promise
      : Promise.resolve(sectionResult(args)));
    await start();
    expect(ghost()).toContain('אָמַר');
    handle?.dispose();
    pending.resolve({ ok: true, value: null });
    await vi.advanceTimersByTimeAsync(0);
  });
});

describe('book completion: content events and cache lifetime', () => {
  it('refreshes a displayed neighbor after a rendering event on an unchanged heading', async () => {
    setSection('source', 0, 'פרק א');
    setSection('source', 1, POINTED);
    await start();
    expect(ghost()).toContain('אָמַר');
    expect(reader.listeners.has('reader.sectionContentChanged')).toBe(true);
    setSection('source', 1, PLAIN, { sourceText: POINTED, sourceTextHash: POINTED });
    sectionChanged('source', 0, {
      changeType: 'rendering-only', reason: 'nikud-toggle',
      oldSourceTextHash: 'פרק א', newSourceTextHash: 'פרק א',
      oldRenderedTextHash: 'פרק א', newRenderedTextHash: 'פרק א',
    });
    expect(ghost()).toBe('');
    // The next input starts a lazy reload. Publishing the new window must
    // restore the ghost after its debounce, without another user input.
    await evaluate();
    await vi.advanceTimersByTimeAsync(200);
    expect(ghost()).toBe(PLAIN_GHOST);
  });

  it('discovers a changed background match after one input and the reload debounce', async () => {
    state.openTabs = [bookTab('background')];
    setSection('source', 0, 'ראשית חכמה');
    setSection('background', 0, 'פרק א');
    setSection('background', 1, 'טקסט אחר בלי התאמה');
    await start();
    expect(ghost()).toBe('');
    setSection('background', 1, POINTED);
    sectionChanged('background', 1);
    await vi.advanceTimersByTimeAsync(200);
    await evaluate();
    await vi.advanceTimersByTimeAsync(200);
    expect(ghost()).toContain('אָמַר');
  });

  it('discards a page invalidated by a content event while its window is loading', async () => {
    const pending = deferred<ReaderResult<SectionTextMapResult | null>>();
    setSection('source', 0, 'פרק א');
    setSection('source', 1, POINTED, { hasMore: true, nextCursor: 'obsolete-page' });
    const oldPage = sectionResult({ bookId: 'source', bookUid: 'source', sectionIndex: 1 });
    let delayed = false;
    reader.getSection.mockImplementation((args) => {
      if (args.bookUid === 'source' && args.sectionIndex === 1 && !delayed) {
        delayed = true;
        return pending.promise;
      }
      return Promise.resolve(sectionResult(args));
    });
    await mount();
    expect(delayed).toBe(true);
    setSection('source', 1, PLAIN);
    sectionChanged('source', 1);
    pending.resolve(oldPage);
    await vi.advanceTimersByTimeAsync(200);
    await evaluate();
    await vi.advanceTimersByTimeAsync(200);
    expect(ghost()).toBe(PLAIN_GHOST);
    expect(requests().some((request) => request.cursor === 'obsolete-page')).toBe(false);
  });

  it('coalesces a burst of content events without repeatedly scanning a warm window', async () => {
    state.openTabs = [bookTab('background')];
    setSection('source', 0, 'פרק א');
    setSection('source', 1, POINTED);
    setSection('background', 0, 'ראשית חכמה');
    await start();
    expect(ghost()).toContain('אָמַר');
    reader.getSection.mockClear();
    setSection('source', 1, PLAIN);
    for (let index = 0; index < 30; index += 1) sectionChanged('source', 1);
    await vi.advanceTimersByTimeAsync(200);
    await evaluate();
    await vi.advanceTimersByTimeAsync(200);
    expect(ghost()).toBe(PLAIN_GHOST);
    // A distant, nonmatching paragraph is read by context rebuilds, never by
    // candidate validation. One event burst must not scan it thirty times.
    expect(requests().filter((request) => request.bookUid === 'source' && request.sectionIndex === 20).length)
      .toBeLessThanOrEqual(1);
    expect(requests().filter((request) => request.bookUid === 'background' && request.sectionIndex === 20))
      .toEqual([]);
  });

  it('eventually discovers a new match on hosts that provide no content event', async () => {
    setSection('source', 0, 'פרק א');
    setSection('source', 1, 'טקסט אחר בלי התאמה');
    await start();
    expect(ghost()).toBe('');
    reader.getSection.mockClear();
    setSection('source', 1, POINTED);
    await refresh();
    expect(ghost()).toBe('');
    expect(requests().filter((request) => request.sectionIndex === 1)).toEqual([]);
    // Legacy hosts get a bounded discovery refresh, even with an unchanged
    // center and no previously matching candidate to validate.
    await vi.advanceTimersByTimeAsync(31_000);
    await evaluate();
    await evaluate();
    await vi.advanceTimersByTimeAsync(200);
    expect(ghost()).toContain('אָמַר');
  });

  it('bounds automatic reloads when a source changes on every validation read', async () => {
    let revision = 0;
    reader.getSection.mockImplementation(async (args) => {
      if (args.sectionIndex === 0) {
        revision += 1;
        setSection('source', 0, `${POINTED} גרסה ${revision}`);
      }
      return sectionResult(args);
    });
    await mount();
    // Exactly one user input; later timers must not invent more retry budgets.
    await evaluate();
    await vi.advanceTimersByTimeAsync(800);
    expect(ghost()).toBe('');
    const settledReads = requests().length;
    expect(requests().filter((request) => request.sectionIndex === 0).length).toBeLessThanOrEqual(6);
    await vi.advanceTimersByTimeAsync(10_000);
    expect(requests()).toHaveLength(settledReads);
  });

  it('unsubscribes from content events and does not restart work after disposal', async () => {
    setSection('source', 0, POINTED);
    await start();
    expect(reader.listeners.has('reader.sectionContentChanged')).toBe(true);
    handle?.dispose();
    expect(reader.listeners.size).toBe(0);
    reader.getSection.mockClear();
    sectionChanged('source', 0);
    bookChanged();
    await vi.advanceTimersByTimeAsync(2_000);
    await evaluate();
    expect(requests()).toEqual([]);
    expect(ghost()).toBe('');
    expect(onStatus).not.toHaveBeenCalled();
  });
});

describe('book completion: cancellation and recoverable reader errors', () => {
  it('stops an old window after navigation without exposing its delayed content', async () => {
    const pending = deferred<ReaderResult<SectionTextMapResult | null>>();
    setSection('source', 0, POINTED);
    reader.getSection.mockImplementation((args) => args.bookUid === 'source'
      ? pending.promise
      : Promise.resolve(sectionResult(args)));
    await mount();
    expect(loadedBooks()).toContain('source');
    state = currentState({ currentBook: 'New Book', currentBookId: 'new', bookUid: 'new' });
    setSection('new', 0, 'אבי אשמח לשמוע חדשות טובות');
    bookChanged();
    await evaluate();
    expect(ghost()).toBe('אשמח לשמוע חדשות טובות');
    const oldRequests = requests().filter((request) => request.bookUid === 'source').length;
    pending.resolve({
      ok: true,
      value: {
        schemaVersion: 1, bookId: 'source', sectionIndex: 0, currentRef: null,
        renderedText: POINTED, renderedTextHash: POINTED, hasMore: true, nextCursor: 'old-page-2',
      },
    });
    await vi.advanceTimersByTimeAsync(200);
    await evaluate();
    expect(ghost()).toBe('אשמח לשמוע חדשות טובות');
    expect(requests().filter((request) => request.bookUid === 'source')).toHaveLength(oldRequests);
  });

  it('does not continue a paginated RPC or recreate the overlay after disposal', async () => {
    const pending = deferred<ReaderResult<SectionTextMapResult | null>>();
    reader.getSection.mockReturnValue(pending.promise);
    await mount();
    expect(reader.getSection).toHaveBeenCalled();
    const requestsBeforeDispose = requests().length;
    handle?.dispose();
    pending.resolve({
      ok: true,
      value: {
        schemaVersion: 1, bookId: 'source', sectionIndex: 0, currentRef: null,
        renderedText: POINTED, renderedTextHash: POINTED, hasMore: true, nextCursor: 'next-page',
      },
    });
    await vi.advanceTimersByTimeAsync(200);
    await evaluate();
    expect(requests()).toHaveLength(requestsBeforeDispose);
    expect(ghost()).toBe('');
    expect(onStatus).not.toHaveBeenCalled();
  });

  it('discards a delayed background result after that tab closes', async () => {
    const pending = deferred<ReaderResult<SectionTextMapResult | null>>();
    state.openTabs = [bookTab('background')];
    setSection('background', 0, POINTED);
    reader.getSection.mockImplementation((args) => args.bookUid === 'background'
      ? pending.promise
      : Promise.resolve(sectionResult(args)));
    await start();
    expect(loadedBooks()).toContain('background');
    state.openTabs = [];
    bookChanged();
    await vi.advanceTimersByTimeAsync(0);
    const readsBeforeClose = requests().filter((request) => request.bookUid === 'background').length;
    pending.resolve(sectionResult({ bookId: 'background', bookUid: 'background', sectionIndex: 0 }));
    await evaluate();
    expect(ghost()).toBe('');
    expect(requests().filter((request) => request.bookUid === 'background')).toHaveLength(readsBeforeClose);
  });

  it('retries a failed initial load when the reader becomes available', async () => {
    reader.getSection.mockResolvedValue({ ok: false, reason: 'bridge', message: 'Reader temporarily unavailable' });
    await start();
    expect(ghost()).toBe('');
    expect(onStatus).toHaveBeenCalledWith('Reader temporarily unavailable', true);
    setSection('source', 0, POINTED);
    reader.getSection.mockImplementation(async (args) => sectionResult(args));
    await refresh();
    expect(ghost()).toContain('אָמַר');
  });

  it('keeps the loaded center usable if a surrounding line fails', async () => {
    state.currentIndex = 5;
    setSection('source', 5, POINTED);
    reader.getSection.mockImplementation(async (args) => args.sectionIndex === 0
      ? { ok: false, reason: 'unavailable', message: 'Neighbor unavailable' }
      : sectionResult(args));
    await start();
    expect(ghost()).toContain('אָמַר');
  });

  it('keeps a failed background book silent and preserves current-book completion', async () => {
    state.openTabs = [bookTab('background')];
    setSection('source', 0, POINTED);
    reader.getSection.mockImplementation(async (args) => args.bookUid === 'background'
      ? { ok: false, reason: 'permission', message: 'Background book is not readable' }
      : sectionResult(args));
    await start();
    expect(ghost()).toContain('אָמַר');
    expect(onStatus).not.toHaveBeenCalled();
    setSection('background', 0, 'יוסף סיפר סיפור ארוך');
    reader.getSection.mockImplementation(async (args) => sectionResult(args));
    typedText = 'יוסף ס';
    await refresh();
    expect(ghost()).toBe('סיפר סיפור ארוך');
  });
});
