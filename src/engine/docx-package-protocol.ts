import type { PreflightPackage } from './docx-preflight-core';

/** Versioned package operations; postflight can extend this map and the dispatcher. */
export interface PackageResults {
  preflight: Omit<PreflightPackage, 'bytes'> & { bytes: ArrayBuffer | null };
}

export type PackageOperation = keyof PackageResults;
export const PACKAGE_PROTOCOL_VERSION = 1;

export interface PackageRequest<K extends PackageOperation = PackageOperation> {
  version: typeof PACKAGE_PROTOCOL_VERSION;
  id: number;
  operation: K;
  /** Ownership is transferred, never structured-cloned. The caller keeps its Blob. */
  bytes: ArrayBuffer;
}

export type PackageResponse<K extends PackageOperation = PackageOperation> = {
  version: typeof PACKAGE_PROTOCOL_VERSION;
  id: number;
  operation: K;
} & ({ ok: true; result: PackageResults[K] } | { ok: false; error: string });

const validators: { [K in PackageOperation]: (value: unknown) => boolean } = {
  preflight(value) {
    const result = value as PackageResults['preflight'] | null;
    return !!result && (result.bytes === null || result.bytes instanceof ArrayBuffer)
      && (result.fontTable === null || typeof result.fontTable === 'string')
      && (result.notice === null || typeof result.notice === 'string')
      && Array.isArray(result.notes) && result.notes.every(note => typeof note === 'string')
      && !!result.vba && typeof result.vba.hasMacroPart === 'boolean';
  },
};

export function isPackageResult<K extends PackageOperation>(operation: K, value: unknown): value is PackageResults[K] {
  return validators[operation](value);
}

export function isPackageRequest(value: unknown): value is PackageRequest {
  const request = value as Partial<PackageRequest> | null;
  return request !== null && typeof request === 'object'
    && request.version === PACKAGE_PROTOCOL_VERSION
    && Number.isSafeInteger(request.id) && (request.id ?? 0) > 0
    && typeof request.operation === 'string' && Object.prototype.hasOwnProperty.call(validators, request.operation)
    && request.bytes instanceof ArrayBuffer;
}
