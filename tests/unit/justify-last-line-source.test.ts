import { describe, expect, it, vi } from 'vitest';
import { createJustificationReader } from '../../src/engine/justify-last-line-source';
import { invalidateDocumentReads } from '../../src/engine/document-read-state';
const node = (paragraph: object, id = 'p1') => ({ address: { kind: 'block', nodeId: id }, node: { kind: 'paragraph', paragraph } });
const catalog = (textAlign = 'justify') => ({ sourceStatus: { styles: 'present', preview: 'available' }, defaults: { paragraphStyleId: 'Normal' },
  items: [{ id: 'Normal', type: 'paragraph', provenance: 'authored', preview: { available: true, css: { textAlign } } }] });

describe('verified paragraph justification', () => {
  it.each(['justify', 'both'])('accepts direct %s without a style read', async (alignment) => {
    const doc = { getNodeById: vi.fn(async () => node({ props: { alignment } })), styles: { getCatalog: vi.fn() } };
    expect((await createJustificationReader(doc).isJustified('p1')).justified).toBe(true); expect(doc.styles.getCatalog).not.toHaveBeenCalled();
  });
  it.each(['left', 'right', 'center', 'distributed', 'unknown'])('rejects %s even if the style is justified', async (alignment) => {
    const doc = { getNodeById: async () => node({ props: { alignment } }), styles: { getCatalog: vi.fn(async () => catalog()) } };
    expect((await createJustificationReader(doc).isJustified('p1')).justified).toBe(false); expect(doc.styles.getCatalog).not.toHaveBeenCalled();
  });
  it('resolves an inherited/default style through the catalog without scanning body usage', async () => {
    const getCatalog = vi.fn(async () => catalog());
    const doc = { getNodeById: async () => node({}), styles: { getCatalog } };
    expect((await createJustificationReader(doc).isJustified('p1')).justified).toBe(true);
    expect(getCatalog).toHaveBeenCalledWith({ includePreview: true, includeUsage: false, includeHidden: true, includeLatent: false });
  });
  it('uses the named style before the paragraph default', async () => {
    const doc = { getNodeById: async () => node({ styleRef: 'Missing' }), styles: { getCatalog: async () => catalog() } };
    expect((await createJustificationReader(doc).isJustified('p1')).justified).toBe(false);
  });
  it('does not infer conditional table formatting from a paragraph style preview', async () => {
    const doc = { getNodeById: async () => node({}), styles: { getCatalog: vi.fn(async () => catalog()) } };
    expect((await createJustificationReader(doc).isJustified('p1', true)).justified).toBe(false); expect(doc.styles.getCatalog).not.toHaveBeenCalled();
  });
  it.each(['missing', 'malformed'])('rejects a %s styles part', async (status) => {
    const styles = catalog(); styles.sourceStatus.styles = status;
    const doc = { getNodeById: async () => node({}), styles: { getCatalog: async () => styles } };
    expect((await createJustificationReader(doc).isJustified('p1')).justified).toBe(false);
  });
  it('rejects synthetic default-floor previews', async () => {
    const styles = catalog(); styles.items[0].provenance = 'default-floor';
    expect((await createJustificationReader({ getNodeById: async () => node({}), styles: { getCatalog: async () => styles } }).isJustified('p1')).justified).toBe(false);
  });
  it('rejects an incorrect address or unavailable API', async () => {
    expect((await createJustificationReader({ getNodeById: async () => node({ props: { alignment: 'justify' } }, 'other') }).isJustified('p1')).justified).toBe(false);
    expect((await createJustificationReader({}).isJustified('p1')).justified).toBe(false);
  });
  it('shares a paragraph read and style catalog within the revision', async () => {
    const doc = { getNodeById: vi.fn(async ({ nodeId }: { nodeId: string }) => node({}, nodeId)), styles: { getCatalog: vi.fn(async () => catalog()) } };
    const reader = createJustificationReader(doc); await Promise.all([reader.isJustified('p1'), reader.isJustified('p1'), reader.isJustified('p2')]);
    expect(doc.getNodeById).toHaveBeenCalledTimes(2); expect(doc.styles.getCatalog).toHaveBeenCalledTimes(1);
    invalidateDocumentReads(doc); await reader.isJustified('p1'); expect(doc.styles.getCatalog).toHaveBeenCalledTimes(2);
  });
  it('rejects a delayed read from the previous document revision', async () => {
    let resolve!: (value: unknown) => void;
    const doc = { getNodeById: () => new Promise((r) => { resolve = r; }) }; const reader = createJustificationReader(doc);
    const pending = reader.isJustified('p1'); invalidateDocumentReads(doc); resolve(node({ props: { alignment: 'justify' } }));
    const approval = await pending; expect(approval.justified).toBe(false); expect(reader.isCurrent(approval)).toBe(false);
  });
});
