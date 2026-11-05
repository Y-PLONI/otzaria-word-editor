import { describe, expect, it, vi } from 'vitest';
import { createRecentDocumentStore } from '../../src/sessions/recent-document-store';
import type { RecentDocument, RecentOpen } from '../../src/sessions/recent-documents';
import { loadRecentDocuments, saveRecentDocuments } from '../../src/host/settings';

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

const entry = (token: string, openedAt = 10): RecentOpen => ({
  token, name: `${token}.docx`, size: 1, openedAt, writable: true,
});

function setup(load: () => Promise<unknown>, save = vi.fn(async (_: readonly RecentDocument[]) => {})) {
  const state = { list: [] as RecentDocument[], loading: false };
  const onError = vi.fn();
  const store = createRecentDocumentStore({
    load, save, onError,
    onChange: (list) => { state.list = list; },
    onLoading: (loading) => { state.loading = loading; },
  });
  return { store, state, save, onError };
}

describe('recent history loaded outside the opening barrier', () => {
  it.each(['reply', 'throw'])('preserves unread history through a real host %s failure', async (failure) => {
    let failing = true;
    let current: RecentDocument[] = [];
    const errors = vi.fn();
    const saved: unknown[] = [];
    window.Otzaria = {
      call: async (method: string, payload: { value?: unknown }) => {
        if (method === 'storage.get') {
          if (failing) {
            if (failure === 'throw') throw new Error('offline');
            return { success: false, error: { message: 'offline' } };
          }
          return { success: true, data: [{ ...entry('old'), pinned: true }] };
        }
        saved.push(payload.value);
        return { success: true, data: null };
      },
    } as never;
    try {
      const store = createRecentDocumentStore({
        load: loadRecentDocuments, save: saveRecentDocuments,
        onChange: list => { current = list; }, onLoading: () => {}, onError: errors,
      });
      store.remember(entry('new'));
      await vi.waitFor(() => expect(errors).toHaveBeenCalledOnce());
      expect(saved).toEqual([]);
      failing = false;
      await store.load();
      await vi.waitFor(() => expect(saved).toHaveLength(1));
      expect(current.map(item => item.token)).toEqual(['new', 'old']);
      expect(current.find(item => item.token === 'old')?.pinned).toBe(true);
      expect(saved[0]).toEqual(current);
    } finally {
      delete (window as Partial<Window>).Otzaria;
    }
  });

  it('merges an early opening with stored history and preserves a stored pin', async () => {
    const gate = deferred<unknown>();
    const load = vi.fn(() => gate.promise);
    const { store, state, save } = setup(load);
    const ready = store.load();
    store.remember(entry('known', 100));
    expect(state.list.map((item) => item.token)).toEqual(['known']);
    expect(state.loading).toBe(true);
    expect(save).not.toHaveBeenCalled();
    gate.resolve([{ ...entry('known', 1), pinned: true }, { ...entry('older', 2), pinned: false }]);
    await ready;
    await vi.waitFor(() => expect(save).toHaveBeenCalledOnce());
    expect(state.list.map((item) => item.token)).toEqual(['known', 'older']);
    expect(state.list[0]).toMatchObject({ pinned: true, openedAt: 100, writable: true });
    expect(save).toHaveBeenLastCalledWith(state.list);
    await store.load();
    expect(load).toHaveBeenCalledOnce();
    expect(state.loading).toBe(false);
  });

  it('replays early removals, pin changes and openings in order', async () => {
    const gate = deferred<unknown>();
    const { store, state, save } = setup(() => gate.promise);
    const ready = store.load();
    store.remember(entry('new'));
    store.forget('removed');
    store.setPinned('stored', true);
    store.setPinned('new', true);
    store.forget('new');
    gate.resolve([
      { ...entry('removed'), pinned: true },
      { ...entry('stored'), pinned: false },
    ]);
    await ready;
    await vi.waitFor(() => expect(save).toHaveBeenCalledOnce());
    expect(state.list).toEqual([{ ...entry('stored'), pinned: true }]);
  });

  it('serializes writes so an older save cannot complete after a newer save', async () => {
    const gate = deferred<void>();
    const save = vi.fn(async (_: readonly RecentDocument[]) => {});
    save.mockImplementationOnce(() => gate.promise);
    const { store } = setup(async () => [], save);
    await store.load();
    store.remember(entry('first'));
    await vi.waitFor(() => expect(save).toHaveBeenCalledOnce());
    store.remember(entry('second', 20));
    await Promise.resolve();
    expect(save).toHaveBeenCalledOnce();
    gate.resolve();
    await vi.waitFor(() => expect(save).toHaveBeenCalledTimes(2));
    expect(save.mock.calls[0]?.[0].map((item) => item.token)).toEqual(['first']);
    expect(save.mock.calls[1]?.[0].map((item) => item.token)).toEqual(['second', 'first']);
  });

  it('does not write over unread history after a load failure and retries pending changes', async () => {
    const load = vi.fn<() => Promise<unknown>>()
      .mockRejectedValueOnce(new Error('load failed'))
      .mockResolvedValueOnce([{ ...entry('old'), pinned: true }]);
    const { store, state, save, onError } = setup(load);
    store.remember(entry('new'));
    await vi.waitFor(() => expect(onError).toHaveBeenCalledOnce());
    expect(save).not.toHaveBeenCalled();
    expect(state.list.map((item) => item.token)).toEqual(['new']);
    await store.load();
    await vi.waitFor(() => expect(save).toHaveBeenCalledOnce());
    expect(state.list.map((item) => item.token)).toEqual(['new', 'old']);
  });

  it('continues saving after a rejected write', async () => {
    const save = vi.fn(async (_: readonly RecentDocument[]) => {});
    save.mockRejectedValueOnce(new Error('save failed'));
    const { store, onError } = setup(async () => [], save);
    await store.load();
    store.remember(entry('first'));
    store.remember(entry('second', 20));
    await vi.waitFor(() => expect(save).toHaveBeenCalledTimes(2));
    expect(onError).toHaveBeenCalledOnce();
    expect(save.mock.calls[1]?.[0].map((item) => item.token)).toEqual(['second', 'first']);
  });

  it('does not persist an unmodified hydrated list', async () => {
    const { store, state, save } = setup(async () => [{ ...entry('old'), pinned: true }]);
    await store.load();
    expect(state.list).toHaveLength(1);
    expect(save).not.toHaveBeenCalled();
  });
});
