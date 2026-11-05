import { DOCX_MIME } from './docx-mime';
import { preflightPackage } from './docx-preflight-core';
import { runPackageTask } from './docx-package-client';
import type { PackageResults } from './docx-package-protocol';
import { NO_VBA, type DocumentVba } from './vba-import';

// Keep the existing pure helpers available to repair callers and regression tests.
export * from './docx-preflight-core';

export interface PreflightResult {
  source: string | File | Blob | undefined;
  fontTable: string | null;
  vba: DocumentVba;
  notice: string | null;
}

/**
 * Read once, retain the immutable source, and transfer only its ArrayBuffer to the
 * package worker. The main thread installs fonts and connects the returned source
 * to the engine; ZIP, XML, CRC and VBA processing run in the worker.
 *
 * If a worker cannot run, use the same core locally. Skipping the repair solely
 * because a worker failed would expose defaultTabStop=0 to the engine and freeze
 * the UI. This exceptional fallback can block, just like the previous preflight.
 */
export async function preflightSource(
  source: string | File | Blob | undefined,
  options: { signal?: AbortSignal } = {},
): Promise<PreflightResult> {
  const { signal } = options;
  const unchanged = (original = source): PreflightResult => ({
    source: original, fontTable: null, vba: NO_VBA, notice: null,
  });
  if (source === undefined || signal?.aborted) return unchanged();

  let original: Blob;
  let buffer: ArrayBuffer;
  try {
    if (typeof source === 'string') {
      const response = await fetch(source, { signal });
      const blob = await response.blob();
      original = blob.slice(0, blob.size, DOCX_MIME);
    } else original = source;
    if (signal?.aborted) return unchanged(original);
    buffer = await original.arrayBuffer();
  } catch (error) {
    if (!signal?.aborted) console.warn('[otzaria-word] הבדיקה המקדימה לא קראה את המסמך', error);
    return unchanged();
  }
  if (signal?.aborted) return unchanged(original);

  let result: PackageResults['preflight'];
  try {
    result = await runPackageTask('preflight', buffer, signal);
  } catch (error) {
    if (signal?.aborted) return unchanged(original);
    console.warn('[otzaria-word] worker הבדיקה המקדימה נכשל; מנסה עיבוד מקומי', error);
    try {
      // Constructor/dispatch failures may leave the buffer attached. Read the
      // retained Blob again only if ownership was already transferred; no fetch.
      const recovered = buffer.byteLength || original.size === 0 ? buffer : await original.arrayBuffer();
      if (signal?.aborted) return unchanged(original);
      const local = await preflightPackage(new Uint8Array(recovered));
      result = { ...local, bytes: local.bytes?.buffer ?? null };
    } catch (failure) {
      console.warn('[otzaria-word] הבדיקה המקדימה נכשלה, והמסמך נמסר כמות שהוא', failure);
      return unchanged(original);
    }
  }
  if (signal?.aborted) return unchanged(original);
  for (const note of result.notes) console.warn(`[otzaria-word] ${note}`);
  return {
    source: result.bytes === null ? original : new Blob([result.bytes], { type: DOCX_MIME }),
    fontTable: result.fontTable, vba: result.vba, notice: result.notice,
  };
}
