import { describe, expect, it, vi } from 'vitest';
import type { SuperDocExtensionContext, SuperDocMutationEvent, SuperDocMutationFilter } from 'superdoc';
import { MacroKit, createMemoryStorage, createSuperdocHost, emptyState, type SuperdocLike } from 'superdoc-macros';
import { getDocumentReadRevision } from '../../src/engine/document-read-state';
import { readParagraphIndents } from '../../src/engine/paragraph-format';
import { createDocumentUpdateBridge } from '../../src/engine/document-updates';

vi.mock('superdoc', () => ({ defineSuperDocExtension: (extension: unknown) => extension }));

/** The public extension event contract, including the origin filter. */
function runtime(bridge: ReturnType<typeof createDocumentUpdateBridge>) {
  const listeners = new Set<{ filter: SuperDocMutationFilter; handler: (event: SuperDocMutationEvent) => void }>();
  const unsubscribed = vi.fn();
  const context = {
    // This facade is intentionally not the cache's Document API identity.
    doc: {},
    onMutation: (filter: SuperDocMutationFilter, handler: (event: SuperDocMutationEvent) => void) => {
      const listener = { filter, handler };
      listeners.add(listener);
      return () => { listeners.delete(listener); unsubscribed(); };
    },
    onSelection: vi.fn(),
    onPaint: vi.fn(),
  };
  bridge.extension.activate(context as unknown as SuperDocExtensionContext);
  return {
    context, unsubscribed,
    emit(origin: SuperDocMutationEvent['origin'] = 'extension') {
      const event: SuperDocMutationEvent = { id: 'mutation', origin, affects: new Set(['block']), stories: [] };
      for (const { filter, handler } of listeners) {
        if (!filter.origin || filter.origin === origin) handler(event);
      }
    },
    staleHandler: [...listeners][0]!.handler,
  };
}

function paragraph() {
  let bidi = true;
  let offset = 0;
  const doc = {
    get: vi.fn(() => { throw new Error('A readout must not read the whole document'); }),
    getNodeById: vi.fn(async ({ nodeId }: { nodeId: string }) => ({
      node: { kind: 'paragraph', paragraphIds: { paraId: nodeId }, paragraph: { props: { bidi } } },
      address: { kind: 'block', nodeType: 'paragraph', nodeId },
    })),
    blocks: { list: vi.fn(async ({ nodeIds }: { nodeIds: readonly string[] }) => ({
      blocks: nodeIds.map((nodeId) => ({ nodeId, nodeType: 'paragraph' })),
    })) },
    selection: { current: async () => ({
      empty: true,
      target: { kind: 'text', segments: [{ blockId: 'p1', range: { start: offset, end: offset } }] },
    }) },
  };
  return {
    doc,
    host: { activeEditor: { doc } },
    changeDirection() { bidi = false; },
    moveCaret() { offset += 1; },
  };
}

describe('document mutation updates', () => {
  it.each([false, true])('refreshes a shared reading after real MacroKit replay (later step fails: %s)', async (failLater) => {
    const f = paragraph();
    const onUpdate = vi.fn();
    const bridge = createDocumentUpdateBridge(onUpdate);
    const events = runtime(bridge);
    bridge.bind(f.doc);
    const superdoc = {
      ...f.host,
      ui: { commands: {
        has: () => true,
        get: () => ({ getState: () => ({ enabled: true }) }),
        executeAsync: async (id: string) => {
          if (id === 'fail') throw new Error('The second step failed');
          f.changeDirection();
          // 2.23 emits this public hook but omits onEditorUpdate for this origin.
          events.emit();
          return { success: true };
        },
      } },
    };
    const host = createSuperdocHost({ superdoc: superdoc as unknown as SuperdocLike });
    const storage = createMemoryStorage();
    storage.save({ ...emptyState(), recordings: [{
      version: 1, id: 'ltr', name: 'LTR', steps: [
        { type: 'command', id: 'direction-ltr' },
        ...(failLater ? [{ type: 'command' as const, id: 'fail' }] : []),
      ],
    }] });
    const kit = new MacroKit({ host, storage, runner: 'eval', scriptsEnabled: false });
    try {
      const [home, ruler] = await Promise.all([readParagraphIndents(f.host), readParagraphIndents(f.host)]);
      expect(home?.indents.bidi).toBe(true);
      expect(ruler).toEqual(home);
      expect(f.doc.getNodeById).toHaveBeenCalledTimes(1);

      const result = await kit.replayRecording('ltr');
      expect(result.ok).toBe(!failLater);
      // This callback is also the App's navDocRevision/dirty notification.
      expect(onUpdate).toHaveBeenCalledTimes(1);
      f.moveCaret();
      expect((await readParagraphIndents(f.host))?.indents.bidi).toBe(false);
      expect(f.doc.getNodeById).toHaveBeenCalledTimes(2);
      for (let i = 0; i < 5; i++) {
        f.moveCaret();
        expect((await readParagraphIndents(f.host))?.indents.bidi).toBe(false);
      }
      expect(f.doc.getNodeById).toHaveBeenCalledTimes(2);
      expect(f.doc.blocks.list.mock.calls).toEqual([[{ nodeIds: ['p1'] }], [{ nodeIds: ['p1'] }]]);
      expect(f.doc.get).not.toHaveBeenCalled();
    } finally {
      host.dispose();
      bridge.dispose();
    }
  });

  it('invalidates the captured document when a pending write commits after a tab switch', async () => {
    const a = paragraph();
    const b = paragraph();
    let active = a;
    let navDocRevision = 0;
    const bridge = createDocumentUpdateBridge(() => { if (active === a) navDocRevision++; });
    const events = runtime(bridge);
    bridge.bind(a.doc);
    await Promise.all([readParagraphIndents(a.host), readParagraphIndents(b.host)]);
    const revisionA = getDocumentReadRevision(a.doc);
    const revisionB = getDocumentReadRevision(b.doc);
    let finish!: () => void;
    const pending = new Promise<void>((resolve) => { finish = resolve; }).then(() => {
      a.changeDirection();
      events.emit();
    });

    active = b;
    finish();
    await pending;
    expect(getDocumentReadRevision(a.doc)).not.toBe(revisionA);
    expect(getDocumentReadRevision(b.doc)).toBe(revisionB);
    expect(navDocRevision).toBe(0);
    expect((await readParagraphIndents(b.host))?.indents.bidi).toBe(true);
    expect(b.doc.getNodeById).toHaveBeenCalledTimes(1);
    active = a;
    expect((await readParagraphIndents(a.host))?.indents.bidi).toBe(false);
    expect(a.doc.getNodeById).toHaveBeenCalledTimes(2);
    bridge.dispose();
  });

  it('holds pre-ready extension commits until the actual document is bound', () => {
    const doc = {};
    const onUpdate = vi.fn(() => expect(getDocumentReadRevision(doc)).not.toBe(revision));
    const bridge = createDocumentUpdateBridge(onUpdate);
    const events = runtime(bridge);
    const revision = getDocumentReadRevision(doc);
    const facadeRevision = getDocumentReadRevision(events.context.doc);
    events.emit();
    events.emit();
    expect(onUpdate).not.toHaveBeenCalled();
    bridge.bind(doc);
    expect(onUpdate).toHaveBeenCalledTimes(1);
    expect(getDocumentReadRevision(events.context.doc)).toBe(facadeRevision);
    bridge.dispose();
  });

  it('filters out non-extension origins and does not subscribe to selection/paint', () => {
    const doc = {};
    const onUpdate = vi.fn();
    const bridge = createDocumentUpdateBridge(onUpdate);
    const events = runtime(bridge);
    bridge.bind(doc);
    const revision = getDocumentReadRevision(doc);
    for (const origin of ['local', 'history', 'remote'] as const) events.emit(origin);
    expect(getDocumentReadRevision(doc)).toBe(revision);
    expect(onUpdate).not.toHaveBeenCalled();
    bridge.onEditorUpdate(doc);
    expect(onUpdate).toHaveBeenCalledTimes(1);
    expect(events.context.onSelection).not.toHaveBeenCalled();
    expect(events.context.onPaint).not.toHaveBeenCalled();
    bridge.dispose();
  });

  it('ignores late native/extension events after disposal, including a pending early commit', () => {
    const doc = {};
    const onUpdate = vi.fn();
    const bridge = createDocumentUpdateBridge(onUpdate);
    const events = runtime(bridge);
    events.emit();
    bridge.dispose();
    bridge.dispose();
    const revision = getDocumentReadRevision(doc);
    bridge.bind(doc);
    bridge.onEditorUpdate(doc);
    events.staleHandler({ id: 'late', origin: 'extension', affects: new Set(['block']), stories: [] });
    expect(getDocumentReadRevision(doc)).toBe(revision);
    expect(onUpdate).not.toHaveBeenCalled();
    expect(events.unsubscribed).toHaveBeenCalledTimes(1);
  });
});
