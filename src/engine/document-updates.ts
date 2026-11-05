import { defineSuperDocExtension } from 'superdoc';
import { invalidateDocumentReads } from './document-read-state';

/**
 * SuperDoc 2.23 omits extension-origin commits from onEditorUpdate. MacroKit,
 * extension commands and direct Document API writes use that origin, so the
 * public mutation hook must also reach the session's dirty/readout callback.
 * Selection/paint events do not invalidate document reads.
 */
export function createDocumentUpdateBridge(onUpdate?: () => void) {
  let documentApi: object | null = null;
  let pending = false;
  let disposed = false;
  const subscriptions = new Set<() => void>();

  function notify(doc: object | null = documentApi): void {
    if (disposed) return;
    // Extensions activate before onReady. Keep an early commit until we can
    // bind the real Document API identity (context.doc is a guarded facade).
    if (!doc) {
      pending = true;
      return;
    }
    invalidateDocumentReads(doc);
    onUpdate?.();
  }

  return {
    extension: defineSuperDocExtension({
      id: 'otzaria.document-updates',
      activate(context) {
        if (disposed) return;
        let active = true;
        // Native local/history commits already reach onEditorUpdate.
        const subscription = context.onMutation({ origin: 'extension' }, () => {
          if (active) notify();
        });
        const stop = () => {
          if (!active) return;
          active = false;
          subscriptions.delete(stop);
          if (typeof subscription === 'function') subscription();
          else subscription.dispose();
        };
        subscriptions.add(stop);
        return stop;
      },
    }),
    /** Bind once to this opened document, never to the currently active tab. */
    bind(doc: object | null | undefined): void {
      if (disposed) return;
      documentApi = doc ?? null;
      if (pending && documentApi) {
        pending = false;
        notify();
      }
    },
    /** The native event payload is usable even before createEditor resolves. */
    onEditorUpdate(doc?: object | null): void {
      notify(doc ?? documentApi);
    },
    dispose(): void {
      if (disposed) return;
      disposed = true;
      pending = false;
      documentApi = null;
      for (const stop of [...subscriptions]) stop();
    },
  };
}
