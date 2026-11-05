import PackageWorker from './docx-package.worker?worker&inline';
import {
  PACKAGE_PROTOCOL_VERSION,
  isPackageResult,
  type PackageOperation,
  type PackageRequest,
  type PackageResponse,
  type PackageResults,
} from './docx-package-protocol';

let nextRequestId = 0;

/**
 * One worker per request: obsolete document work can terminate immediately,
 * without cancelling another document. Vite's IIFE inline worker is classic and
 * starts from a Blob URL, including in the opaque origin of the packaged file:// UI.
 */
export function runPackageTask<K extends PackageOperation>(
  operation: K,
  bytes: ArrayBuffer,
  signal?: AbortSignal,
): Promise<PackageResults[K]> {
  return new Promise((resolve, reject) => {
    const aborted = (): DOMException => new DOMException('Document processing cancelled', 'AbortError');
    if (signal?.aborted) { reject(aborted()); return; }
    const id = ++nextRequestId;
    const worker = new PackageWorker({ name: `otzaria-docx-${operation}` });
    let settled = false;

    function finish(error: Error | null, result?: PackageResults[K]): void {
      if (settled) return;
      settled = true;
      signal?.removeEventListener('abort', cancel);
      worker.onmessage = worker.onerror = worker.onmessageerror = null;
      worker.terminate();
      if (error) reject(error);
      else resolve(result!);
    }
    function cancel(): void { finish(aborted()); }

    worker.onmessage = ({ data }: MessageEvent<PackageResponse<K>>) => {
      // A stale or unrelated response can never settle another request.
      if (!data || data.id !== id || data.operation !== operation) return;
      if (data.version !== PACKAGE_PROTOCOL_VERSION) {
        finish(new Error('Unsupported package worker protocol')); return;
      }
      if (data.ok !== true) {
        finish(new Error(typeof data.error === 'string' ? data.error : 'Package worker failed')); return;
      }
      const result = data.result;
      if (!isPackageResult(operation, result)) {
        finish(new Error('Invalid package worker result')); return;
      }
      finish(null, result);
    };
    worker.onerror = (event) => {
      event.preventDefault();
      finish(new Error(event.message || 'Package worker could not run'));
    };
    worker.onmessageerror = () => finish(new Error('Package worker result could not be read'));
    signal?.addEventListener('abort', cancel, { once: true });
    if (signal?.aborted) { cancel(); return; }
    const request: PackageRequest<K> = { version: PACKAGE_PROTOCOL_VERSION, id, operation, bytes };
    try { worker.postMessage(request, [bytes]); }
    catch (error) { finish(error instanceof Error ? error : new Error('Package worker dispatch failed')); }
  });
}
