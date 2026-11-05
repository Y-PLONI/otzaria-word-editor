import { preflightPackage } from './docx-preflight-core';
import { isPackageRequest, type PackageResponse } from './docx-package-protocol';

// DOM and WebWorker lib declarations overlap. Describe only the dedicated port
// used here, so this entry has no dependency on window or document.
const port = self as unknown as {
  onmessage: (event: MessageEvent<unknown>) => void;
  postMessage(message: PackageResponse, transfer: Transferable[]): void;
};

port.onmessage = async ({ data }) => {
  if (!isPackageRequest(data)) return;
  const envelope = { version: data.version, id: data.id, operation: data.operation };
  try {
    const result = await preflightPackage(new Uint8Array(data.bytes));
    const bytes = result.bytes?.buffer ?? null;
    port.postMessage({ ...envelope, ok: true, result: { ...result, bytes } }, bytes ? [bytes] : []);
  } catch (error) {
    port.postMessage({
      ...envelope, ok: false, error: error instanceof Error ? error.message : 'Package processing failed',
    }, []);
  }
};
