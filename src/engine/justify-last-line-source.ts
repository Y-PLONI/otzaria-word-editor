/** Targeted paragraph/style reads for the last-line workaround. No full-body
 * export, unfiltered block list, selection change or document write is needed.
 */
import { getDocumentReadRevision, type DocumentReadRevision } from './document-read-state';
import type { MaybePromise } from './document-api';

export interface JustificationDocument {
  getNodeById?: (input: { nodeId: string }) => MaybePromise<unknown>;
  styles?: { getCatalog?: (input: {
    includePreview: true; includeUsage: false; includeHidden: true; includeLatent: false;
  }) => MaybePromise<unknown> };
}
interface Approval { revision: DocumentReadRevision; justified: boolean; }
interface Paragraph { props?: { alignment?: unknown }; resolved?: { alignment?: unknown }; styleRef?: string; }
interface Catalog {
  defaults?: { paragraphStyleId?: string | null };
  sourceStatus?: { styles?: string; preview?: string };
  items?: Array<{ id: string; type?: string; provenance?: string; preview?: { available?: boolean; css?: { textAlign?: unknown } } }>;
}

export function createJustificationReader(doc: JustificationDocument) {
  let revision = getDocumentReadRevision(doc);
  let reads = new Map<string, Promise<Approval>>();
  let catalog: Promise<Catalog | null> | undefined;
  function current(): DocumentReadRevision {
    const next = getDocumentReadRevision(doc);
    if (next !== revision) { revision = next; reads = new Map(); catalog = undefined; }
    return next;
  }
  async function read(nodeId: string, inTable: boolean, captured: DocumentReadRevision): Promise<Approval> {
    const no = { revision: captured, justified: false };
    try {
      const result = await doc.getNodeById?.({ nodeId }) as {
        address?: { nodeId?: string; kind?: string };
        node?: { kind?: string; paragraph?: Paragraph; heading?: Paragraph };
      } | null | undefined;
      if (current() !== captured || result?.address?.nodeId !== nodeId || result.address.kind !== 'block') return no;
      const paragraph = result.node?.paragraph ?? result.node?.heading;
      if (!paragraph || !['paragraph', 'heading'].includes(result.node?.kind ?? '')) return no;
      const alignment = paragraph.resolved?.alignment ?? paragraph.props?.alignment;
      if (alignment !== undefined) return { revision: captured, justified: alignment === 'justify' || alignment === 'both' };
      if (inTable || !doc.styles?.getCatalog) return no;
      catalog ??= Promise.resolve(doc.styles.getCatalog({ includePreview: true, includeUsage: false,
        includeHidden: true, includeLatent: false })).then((value) => value as Catalog, () => null);
      const styles = await catalog;
      if (current() !== captured || !styles || styles.sourceStatus?.styles !== 'present' || styles.sourceStatus.preview !== 'available') return no;
      const styleId = paragraph.styleRef ?? styles.defaults?.paragraphStyleId;
      const style = styles.items?.find((item) => item.id === styleId);
      if (!style?.preview?.available || style.provenance !== 'authored' || !['paragraph', 'linked'].includes(style.type ?? '')) return no;
      const inherited = style.preview.css?.textAlign;
      return { revision: captured, justified: inherited === 'justify' };
    } catch { return no; }
  }
  return {
    isCurrent(approval: Approval) { return current() === approval.revision; },
    isJustified(nodeId: string, inTable = false): Promise<Approval> {
      const captured = current();
      const key = `${inTable ? 'table:' : 'body:'}${nodeId}`;
      let pending = reads.get(key);
      if (!pending) {
        if (reads.size >= 128) reads.clear();
        pending = read(nodeId, inTable, captured);
        reads.set(key, pending);
      }
      return pending;
    },
  };
}
