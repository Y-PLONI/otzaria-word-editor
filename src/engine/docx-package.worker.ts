import { preflightPackage } from './docx-preflight-core';
import { postflightDocx } from './docx-postflight';
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
    const result = data.operation === 'preflight'
      ? await preflightPackage(new Uint8Array(data.bytes))
      : { bytes: await postflightDocx(new Uint8Array(data.bytes)) };
    const output = result.bytes;
    // Transfer exactly the returned view, without copying a whole-buffer result.
    const bytes = output === null ? null
      : output.byteOffset === 0 && output.byteLength === output.buffer.byteLength
        ? output.buffer : output.slice().buffer;
    port.postMessage({ ...envelope, ok: true, result: { ...result, bytes } }, bytes ? [bytes] : []);
  } catch (error) {
    port.postMessage({
      ...envelope, ok: false, error: error instanceof Error ? error.message : 'Package processing failed',
    }, []);
  }
};
