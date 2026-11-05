import type { Buffer } from 'node:buffer';

export function zipStored(entries: Record<string, string | Uint8Array>): Buffer<ArrayBuffer>;
export function zipDeflated(entries: Record<string, string | Uint8Array>): Buffer<ArrayBuffer>;
