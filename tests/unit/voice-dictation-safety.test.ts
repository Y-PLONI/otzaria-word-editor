import { describe, expect, it, vi } from 'vitest';
import type { NetworkFetchStreamChunk } from '../../src/types/otzaria_plugin';
import { createVoiceDictation, superdocTarget, type DictationTarget } from '../../src/engine/voice-dictation';
import { explainError, voiceStatus, type FetchStream } from '../../src/host/voice-service';

function gate() {
  let release!: () => void;
  const promise = new Promise<void>((resolve) => { release = resolve; });
  return { promise, release };
}

/** A finite SDK stream, with an optional pause after the first text update. */
function service(events: object[], between?: () => Promise<void>): FetchStream {
  return (request) => (async function* (): AsyncGenerator<NetworkFetchStreamChunk> {
    yield { type: 'response', sequence: 0, status: 200, ok: true, headers: {} };
    if (request.url.endsWith('/status')) {
      yield { type: 'data', sequence: 1, body: JSON.stringify({ service: 'otzaria-voice', pack: true, runtime: true }) + '\n' };
    } else if (request.url.endsWith('/dictate')) {
      for (let i = 0; i < events.length; i++) {
        yield { type: 'data', sequence: i + 1, body: JSON.stringify(events[i]) + '\n' };
        if (i === 1) await between?.();
      }
    } else {
      yield { type: 'data', sequence: 1, body: '{}\n' };
    }
  })();
}

function documentTarget() {
  let text = '';
  let offset = 0;
  const writes: string[] = [];
  const target: DictationTarget = {
    caret: () => ({ blockId: 'b', offset, story: null }),
    textBefore: async () => text,
    async replace(_at, from, to, next, guard) {
      if (guard && (!guard.isActive() || text.slice(from, to) !== guard.expectedText)) return false;
      writes.push(next);
      text = text.slice(0, from) + next + text.slice(to);
      offset = from + next.length;
      return true;
    },
    breakLine: async () => null,
  };
  return {
    target, writes,
    text: () => text,
    userReplace(next: string) { text = next; offset = next.length; },
  };
}

function engine(target: DictationTarget, fetchStream: FetchStream) {
  const say = vi.fn();
  const dictation = createVoiceDictation({ target: () => target, fetchStream: () => fetchStream, say, onState: () => {} });
  return { dictation, say };
}

const READY = { type: 'ready' };
const END = { type: 'end', reason: 'stopped' };
const update = (segment: number, text: string) => ({ type: 'update', segment, text });
const at = { blockId: 'b', offset: 7, story: { kind: 'story', storyType: 'body' } };

function superdocMock() {
  const selection = { getSnapshot: vi.fn(() => null), apply: vi.fn(() => ({ ok: true })) };
  const doc = {
    insert: vi.fn(async () => ({ success: true })),
    replace: vi.fn(async () => ({ success: true })),
    blocks: { split: vi.fn(async () => ({ success: true, inserted: { nodeId: 'next' } })) },
    ranges: { resolve: vi.fn(async (_input: unknown) => ({ preview: { text: 'Heading', truncated: false } })) },
  };
  return { sd: { activeEditor: { doc }, ui: { selection } }, doc, selection };
}

describe('dictation preserves user edits and cancelled sessions', () => {
  it('does not replace a manual edit of the draft, even when its length is unchanged', async () => {
    const document = documentTarget();
    const written = gate();
    const target = { ...document.target, replace: async (...args: Parameters<DictationTarget['replace']>) => {
      const result = await document.target.replace(...args);
      written.release();
      return result;
    } };
    const fetch = service([READY, update(0, 'abc'), update(0, 'abcd'), END], async () => {
      await written.promise;
      document.userReplace('XYZ');
    });
    const { dictation } = engine(target, fetch);
    await dictation.toggle();
    expect(document.text()).toBe('XYZ');
    expect(document.writes).toEqual(['abc']);
  });

  it('discards queued updates when cancellation arrives during an existing write', async () => {
    const document = documentTarget();
    const blocked = gate();
    const emitted = gate();
    const target = { ...document.target, replace: async (...args: Parameters<DictationTarget['replace']>) => {
      // The first mutation has already been submitted; cancellation must discard
      // later mutations without trying to roll back this one.
      const result = await document.target.replace(...args);
      await blocked.promise;
      return result;
    } };
    const fetch = service([READY, update(0, 'one'), update(1, 'two'), END]);
    const wrapped: FetchStream = (request) => (async function* () {
      yield* fetch(request);
      if (request.url.endsWith('/dictate')) emitted.release();
    })();
    const { dictation, say } = engine(target, wrapped);
    const running = dictation.toggle();
    await emitted.promise;
    dictation.cancel();
    blocked.release();
    await running;
    expect(document.writes).toEqual(['one']);
    expect(say).not.toHaveBeenCalled();
  });

  it('does not write the next line after cancellation during a paragraph split', async () => {
    const document = documentTarget();
    const entered = gate();
    const blocked = gate();
    const target = { ...document.target, breakLine: async () => {
      entered.release();
      await blocked.promise;
      return { blockId: 'next', offset: 0, story: null };
    } };
    const fetch = service([READY, { ...update(0, 'one תרד שורה two'), final: true }, END]);
    const { dictation, say } = engine(target, fetch);
    const running = dictation.toggle();
    await entered.promise;
    dictation.cancel();
    blocked.release();
    await running;
    expect(document.writes).toEqual(['one']);
    expect(say).not.toHaveBeenCalled();
  });

  it('drains the old writer before returning to idle when the stream fails', async () => {
    const document = documentTarget();
    const entered = gate();
    const blocked = gate();
    const failed = gate();
    let guard: Parameters<DictationTarget['replace']>[4];
    const target = { ...document.target, replace: async (...args: Parameters<DictationTarget['replace']>) => {
      guard = args[4];
      entered.release();
      await blocked.promise;
      return document.target.replace(...args);
    } };
    const fetch = service([READY, update(0, 'one')]);
    const broken: FetchStream = (request) => (async function* () {
      yield* fetch(request);
      if (request.url.endsWith('/dictate')) {
        failed.release();
        throw new Error('connection reset');
      }
    })();
    const { dictation } = engine(target, broken);
    const running = dictation.toggle();
    await entered.promise;
    await failed.promise;
    await vi.waitFor(() => expect(guard?.isActive()).toBe(false));
    expect(dictation.state()).toBe('listening');
    blocked.release();
    await running;
    expect(dictation.state()).toBe('idle');
    expect(document.writes).toEqual([]);
  });
});

describe('SuperDoc dictation target', () => {
  it.each([
    { kind: 'story', storyType: 'footnote', noteId: '1' },
    { kind: 'story', storyType: 'endnote', noteId: '1' },
    { kind: 'story', storyType: 'headerFooterPart', refId: 'rId1' },
  ])('keeps paragraph splits and the new caret in $storyType', async (story) => {
    const { sd, doc, selection } = superdocMock();
    const point = { ...at, story };
    const next = await superdocTarget(sd as never)!.breakLine(point, 7);
    expect(doc.blocks.split).toHaveBeenCalledWith({
      target: { kind: 'block', nodeType: 'paragraph', nodeId: 'b', story },
      offset: 7,
    });
    expect(next).toEqual({ blockId: 'next', offset: 0, story });
    expect(selection.apply).toHaveBeenCalledWith(expect.objectContaining({
      story,
      start: expect.objectContaining({ blockId: 'next', offset: 0, story }),
      end: expect.objectContaining({ blockId: 'next', offset: 0, story }),
    }));
  });

  it('requires the previous draft text in the atomic replace request', async () => {
    const { sd, doc } = superdocMock();
    const target = superdocTarget(sd as never)!;
    expect(await target.replace(at, 7, 10, 'abcd', { expectedText: 'abc', isActive: () => true })).toBe(true);
    expect(doc.replace).toHaveBeenCalledWith(expect.objectContaining({ text: 'abcd', expectedText: 'abc' }));
  });

  it('does not submit a mutation after cancellation', async () => {
    const { sd, doc, selection } = superdocMock();
    const target = superdocTarget(sd as never)!;
    const guard = { isActive: () => false };
    expect(await target.replace(at, 7, 7, 'abc', guard)).toBe(false);
    expect(await target.breakLine(at, 7, guard)).toBeNull();
    expect(doc.insert).not.toHaveBeenCalled();
    expect(doc.blocks.split).not.toHaveBeenCalled();
    expect(selection.apply).not.toHaveBeenCalled();
  });

  it('does not move the caret when an existing mutation completes after cancellation', async () => {
    const { sd, doc, selection } = superdocMock();
    const blocked = gate();
    let active = true;
    doc.insert.mockImplementation(async () => { await blocked.promise; return { success: true }; });
    const writing = superdocTarget(sd as never)!.replace(at, 7, 7, 'abc', { isActive: () => active });
    active = false;
    blocked.release();
    expect(await writing).toBe(false);
    expect(selection.apply).not.toHaveBeenCalled();
  });

  it('retries a caret move rejected while the new text is being laid out', async () => {
    const { sd, selection } = superdocMock();
    selection.apply.mockReturnValueOnce({ ok: false });
    expect(await superdocTarget(sd as never)!.replace(at, 7, 7, 'abc')).toBe(true);
    expect(selection.apply).toHaveBeenCalledTimes(2);
    expect(selection.apply).toHaveBeenLastCalledWith(expect.objectContaining({
      start: expect.objectContaining({ offset: 10 }),
      end: expect.objectContaining({ offset: 10 }),
    }));
  });

  it('stops retrying caret placement as soon as the recording is cancelled', async () => {
    const { sd, selection } = superdocMock();
    let active = true;
    selection.apply.mockImplementation(() => { active = false; return { ok: false }; });
    const target = superdocTarget(sd as never)!;
    expect(await target.replace(at, 7, 7, 'abc', { isActive: () => active })).toBe(false);
    expect(selection.apply).toHaveBeenCalledTimes(1);
  });

  it('reads a bounded text range, including heading and hyperlink text', async () => {
    const { sd, doc } = superdocMock();
    const target = superdocTarget(sd as never)!;
    expect(await target.textBefore(at)).toBe('Heading');
    expect(doc.ranges.resolve).toHaveBeenCalledWith({
      start: { kind: 'point', point: { kind: 'text', blockId: 'b', offset: 0, story: at.story } },
      end: { kind: 'point', point: { kind: 'text', blockId: 'b', offset: 7, story: at.story } },
      in: at.story,
    });
    doc.ranges.resolve.mockResolvedValue({ preview: { text: 'x'.repeat(40), truncated: false } });
    expect(await target.textBefore({ ...at, offset: 10000 })).toHaveLength(40);
    expect(doc.ranges.resolve).toHaveBeenLastCalledWith(expect.objectContaining({
      start: { kind: 'point', point: expect.objectContaining({ offset: 9960 }) },
    }));
  });

  it('cancels on user input while starting and removes the listeners afterwards', async () => {
    const { sd } = superdocMock();
    const container = document.createElement('div');
    const base = documentTarget();
    const real = superdocTarget(sd as never, container)!;
    const blocked = gate();
    const entered = gate();
    const fetch = service([READY, END]);
    const slow: FetchStream = (request) => (async function* () {
      if (request.url.endsWith('/status')) { entered.release(); await blocked.promise; }
      yield* fetch(request);
    })();
    const spy = vi.fn(slow);
    const { dictation } = engine({ ...base.target, onUserInteraction: real.onUserInteraction }, spy);
    const running = dictation.toggle();
    await entered.promise;
    container.dispatchEvent(new Event('beforeinput'));
    blocked.release();
    await running;
    expect(spy).toHaveBeenCalledTimes(1);
    expect(dictation.state()).toBe('idle');
    // Starting again proves that the cancelled recording left no stale listener.
    await dictation.toggle();
    expect(dictation.state()).toBe('idle');
    expect(spy.mock.calls.filter(([request]) => request.url.endsWith('/dictate'))).toHaveLength(1);
  });
});

describe('SDK stream failures', () => {
  it.each([
    'לתוסף אין הרשאת גישה לשירותים מקומיים (network.localhost). ניתן להפעיל את ההרשאה בהגדרות.',
    'היעד המקומי אינו מוצהר ב-network.allowlist של התוסף',
    'Request timed out',
    'Stream consumer is too slow',
  ])('preserves a message-only SDK failure: %s', async (message) => {
    // createRpcStream rejects with new Error(response.error.message): code is
    // absent on this SDK path, even for permission_denied and forbidden.
    const fetch: FetchStream = () => (async function* () { throw new Error(message); })();
    const error = await voiceStatus(fetch).catch((failure) => failure);
    expect(explainError(error)).toBe(message);
  });

  it('recognizes an actual connection refusal', async () => {
    const fetch: FetchStream = () => (async function* () { throw new Error('SocketException: Connection refused'); })();
    const error = await voiceStatus(fetch).catch((failure) => failure);
    expect(explainError(error)).toContain('שירות הדיבור של אוצריא אינו פועל');
  });
});
