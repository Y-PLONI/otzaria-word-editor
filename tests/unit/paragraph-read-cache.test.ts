import { describe, expect, it, vi } from 'vitest';
import {
  applyParagraphIndentation,
  readParagraphIndents,
  type ParagraphFormatDocumentApi,
  type ParagraphTarget,
} from '../../src/engine/paragraph-format';
import { forgetDocumentReads, invalidateDocumentReads } from '../../src/engine/document-read-state';

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function node(nodeId = 'p1', bidi = true, kind = 'paragraph') {
  return {
    node: { kind, paragraphIds: { paraId: nodeId }, [kind]: { props: { bidi, indent: { start: 36, end: 18 } } } },
    address: { kind: 'block', nodeType: kind, nodeId },
  };
}

function fixture() {
  let blockId = 'p1';
  let offset = 0;
  let story: unknown;
  const get = vi.fn(() => { throw new Error('Full-document reads are forbidden on the readout path'); });
  const getNodeById = vi.fn(async ({ nodeId }: { nodeId: string }) => node(nodeId));
  const getNode = vi.fn(async (address: ParagraphTarget) => ({ ...node(address.nodeId), address }));
  const list = vi.fn(async (input: { nodeIds: readonly string[]; in?: unknown }) => {
    expect(input.nodeIds).toHaveLength(1);
    return { blocks: [{ nodeId: input.nodeIds[0], nodeType: 'paragraph' }] };
  });
  const doc = {
    get,
    getNodeById,
    getNode,
    blocks: { list },
    selection: {
      current: vi.fn(async () => ({
        empty: true,
        target: { kind: 'text', segments: [{ blockId, range: { start: offset, end: offset } }], ...(story ? { story } : {}) },
      })),
    },
  };
  const host: { activeEditor: { doc: ParagraphFormatDocumentApi } } = { activeEditor: { doc } };
  return {
    doc, host, get, getNodeById, getNode, list,
    select(id: string, at = 0, inStory?: unknown) { blockId = id; offset = at; story = inStory; },
  };
}

const HEADER = { kind: 'story', storyType: 'headerFooterPart', refId: 'rIdHeader' };
const FOOTER = { kind: 'story', storyType: 'headerFooterPart', refId: 'rIdFooter' };

describe('targeted paragraph read shared by Home and ruler', () => {
  it('shares the pending request and reuses it across offsets in the same paragraph', async () => {
    const f = fixture();
    const pending = deferred<ReturnType<typeof node>>();
    f.getNodeById.mockReturnValueOnce(pending.promise);
    const home = readParagraphIndents(f.host);
    const ruler = readParagraphIndents(f.host);
    await vi.waitFor(() => expect(f.getNodeById).toHaveBeenCalledTimes(1));
    pending.resolve(node());
    expect(await home).toEqual(await ruler);
    for (let offset = 1; offset <= 20; offset++) {
      f.select('p1', offset);
      expect((await readParagraphIndents(f.host))?.indents.leftTwips).toBe(720);
    }
    expect(f.getNodeById).toHaveBeenCalledExactlyOnceWith({ nodeId: 'p1' });
    expect(f.list).toHaveBeenCalledExactlyOnceWith({ nodeIds: ['p1'] });
    expect(f.get).not.toHaveBeenCalled();
  });

  it('filters by the requested ID even for a paragraph far beyond the first page of blocks', async () => {
    const f = fixture();
    f.select('p9999');
    expect((await readParagraphIndents(f.host))?.target.nodeId).toBe('p9999');
    expect(f.list).toHaveBeenCalledExactlyOnceWith({ nodeIds: ['p9999'] });
    f.select('p10000');
    expect((await readParagraphIndents(f.host))?.target.nodeId).toBe('p10000');
    expect(f.getNodeById).toHaveBeenCalledTimes(2);
  });

  it('isolates identical IDs in body/header/footer and passes the story to both APIs', async () => {
    const f = fixture();
    f.getNode.mockImplementation(async (address) => ({ ...node(address.nodeId, (address.story as typeof HEADER).refId !== 'rIdHeader'), address }));
    expect((await readParagraphIndents(f.host))?.indents.bidi).toBe(true);
    f.select('p1', 0, HEADER);
    const header = await readParagraphIndents(f.host);
    expect(header?.indents.bidi).toBe(false);
    expect(header?.target.story).toEqual(HEADER);
    expect(f.getNode).toHaveBeenLastCalledWith({ kind: 'block', nodeType: 'paragraph', nodeId: 'p1', story: HEADER });
    expect(f.list).toHaveBeenLastCalledWith({ nodeIds: ['p1'], in: HEADER });
    f.select('p1', 3, { refId: 'rIdHeader', storyType: 'headerFooterPart', kind: 'story' });
    await readParagraphIndents(f.host);
    expect(f.getNode).toHaveBeenCalledTimes(1);
    f.select('p1', 0, FOOTER);
    expect((await readParagraphIndents(f.host))?.indents.bidi).toBe(true);
    expect(f.getNodeById).toHaveBeenCalledTimes(1);
    expect(f.getNode).toHaveBeenCalledTimes(2);
    expect(f.get).not.toHaveBeenCalled();
  });

  it('never falls back to a body-only ID read when a story-aware API is unavailable', async () => {
    const f = fixture();
    delete f.host.activeEditor.doc.getNode;
    f.select('p1', 0, HEADER);
    expect(await readParagraphIndents(f.host)).toBeNull();
    expect(f.getNodeById).not.toHaveBeenCalled();
    expect(f.get).not.toHaveBeenCalled();
  });

  it('supports the public getNode address when getNodeById is absent', async () => {
    const f = fixture();
    delete f.host.activeEditor.doc.getNodeById;
    expect((await readParagraphIndents(f.host))?.indents.rightTwips).toBe(360);
    expect(f.getNode).toHaveBeenCalledExactlyOnceWith({ kind: 'block', nodeType: 'paragraph', nodeId: 'p1' });
  });

  it('preserves the listItem write type although 2.23 projects its node as a paragraph', async () => {
    const f = fixture();
    f.list.mockResolvedValue({ blocks: [{ nodeId: 'p1', nodeType: 'listItem' }] });
    expect((await readParagraphIndents(f.host))?.target.nodeType).toBe('listItem');
  });

  it('keeps a table heading resolved by getNode even when blocks.list omits nested paragraphs', async () => {
    const f = fixture();
    f.getNodeById.mockResolvedValue(node('p1', true, 'heading'));
    f.list.mockResolvedValue({ blocks: [] });
    const reading = await readParagraphIndents(f.host);
    expect(reading?.target.nodeType).toBe('heading');
    expect(reading?.indents.bidi).toBe(true);
  });

  it('resolves a list in a table with lists.getState in the selected story', async () => {
    const f = fixture();
    f.select('p1', 0, HEADER);
    f.list.mockResolvedValue({ blocks: [] });
    const getState = vi.fn(async () => ({ success: true, isListItem: true }));
    f.host.activeEditor.doc.lists = { getState };
    expect((await readParagraphIndents(f.host))?.target.nodeType).toBe('listItem');
    expect(getState).toHaveBeenCalledExactlyOnceWith({ target: { kind: 'block', nodeType: 'paragraph', nodeId: 'p1', story: HEADER } });
  });

  it('retains API receivers for getNode, blocks.list and lists.getState', async () => {
    const f = fixture();
    f.select('p1', 0, HEADER);
    const doc = f.host.activeEditor.doc;
    doc.getNode = async function(address) { expect(this).toBe(doc); return { ...node(), address }; };
    doc.blocks = { list: async function() { expect(this).toBe(doc.blocks); return { blocks: [] }; } };
    doc.lists = { getState: async function() { expect(this).toBe(doc.lists); return { success: true, isListItem: false }; } };
    expect(await readParagraphIndents(f.host)).not.toBeNull();
  });

  it('returns null for malformed or mismatched envelopes without a whole-document fallback', async () => {
    for (const value of [null, node().node, node('other'), { ...node(), node: { kind: 'table' } }]) {
      const f = fixture();
      f.host.activeEditor.doc.getNodeById = async () => value;
      expect(await readParagraphIndents(f.host)).toBeNull();
      expect(f.get).not.toHaveBeenCalled();
    }
  });

  it('retries a failed targeted read instead of caching its failure', async () => {
    const f = fixture();
    f.getNodeById.mockRejectedValueOnce(new Error('temporarily unavailable'));
    expect(await readParagraphIndents(f.host)).toBeNull();
    expect(await readParagraphIndents(f.host)).not.toBeNull();
    expect(f.getNodeById).toHaveBeenCalledTimes(2);
  });

  it('invalidates an old pending result and keeps the new revision even if the old call finishes last', async () => {
    const f = fixture();
    const pending = deferred<ReturnType<typeof node>>();
    f.getNodeById.mockReturnValueOnce(pending.promise);
    const old = readParagraphIndents(f.host);
    await vi.waitFor(() => expect(f.getNodeById).toHaveBeenCalledTimes(1));
    invalidateDocumentReads(f.doc);
    f.getNodeById.mockResolvedValue(node('p1', false));
    expect((await readParagraphIndents(f.host))?.indents.bidi).toBe(false);
    pending.resolve(node('p1', true));
    expect(await old).toBeNull();
    expect((await readParagraphIndents(f.host))?.indents.bidi).toBe(false);
    expect(f.getNodeById).toHaveBeenCalledTimes(2);
  });

  it('discards a pending result after replacement or disposal of its document', async () => {
    for (const close of [false, true]) {
      const f = fixture();
      const pending = deferred<ReturnType<typeof node>>();
      f.getNodeById.mockReturnValueOnce(pending.promise);
      const reading = readParagraphIndents(f.host);
      await vi.waitFor(() => expect(f.getNodeById).toHaveBeenCalledTimes(1));
      if (close) forgetDocumentReads(f.doc);
      else f.host.activeEditor.doc = fixture().doc;
      pending.resolve(node());
      expect(await reading).toBeNull();
    }
  });

  it.each([false, true])('invalidates paragraph formatting even without onUpdate (throws=%s)', async (throws) => {
    const f = fixture();
    await readParagraphIndents(f.host);
    f.host.activeEditor.doc.format = { paragraph: { setIndentation: () => {
      f.getNodeById.mockResolvedValue(node('p1', false));
      if (throws) throw new Error('failed after a partial change');
      return { success: true };
    } } };
    const outcome = await applyParagraphIndentation(f.host, { kind: 'block', nodeType: 'paragraph', nodeId: 'p1' }, {
      leftTwips: 720, rightTwips: 360, special: 'none', amountTwips: 0,
    });
    expect(outcome.ok).toBe(!throws);
    expect((await readParagraphIndents(f.host))?.indents.bidi).toBe(false);
    expect(f.getNodeById).toHaveBeenCalledTimes(2);
  });

  it('invalidates the document captured by an async mutation, not a tab activated meanwhile', async () => {
    const f = fixture();
    const next = fixture();
    await Promise.all([readParagraphIndents(f.host), readParagraphIndents(next.host)]);
    const pending = deferred<{ success: true }>();
    f.host.activeEditor.doc.format = { paragraph: { setIndentation: () => pending.promise } };
    const writing = applyParagraphIndentation(f.host, { nodeId: 'p1' }, {
      leftTwips: 0, rightTwips: 0, special: 'none', amountTwips: 0,
    });
    f.host.activeEditor.doc = next.doc;
    pending.resolve({ success: true });
    await writing;
    await readParagraphIndents(next.host);
    expect(next.getNodeById).toHaveBeenCalledTimes(1);
    await readParagraphIndents({ activeEditor: { doc: f.doc } });
    expect(f.getNodeById).toHaveBeenCalledTimes(2);
  });
});
