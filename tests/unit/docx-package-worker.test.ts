import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Blob as NativeBlob } from 'node:buffer';
import { readFileSync } from 'node:fs';
import { URL as NodeURL } from 'node:url';
import { zipStored } from '../../scripts/qa/docx-fixtures.mjs';
import { runPackageTask } from '../../src/engine/docx-package-client';
import { preflightSource, preflightPackage, readDocxPart, SETTINGS_PART, FONT_TABLE_PART } from '../../src/engine/docx-preflight';
import { readZip } from '../../src/engine/docx-parts';
import { NO_VBA, WARNING_TEXT } from '../../src/engine/vba-import';
import type { PackageRequest, PackageResponse, PackageResults } from '../../src/engine/docx-package-protocol';

const constructWorker = vi.hoisted(() => vi.fn());
vi.mock('../../src/engine/docx-package.worker?worker&inline', () => ({ default: constructWorker }));

class WorkerDouble {
  onmessage: ((event: MessageEvent<PackageResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror: (() => void) | null = null;
  request: PackageRequest | null = null;
  sent: PackageRequest | null = null;
  terminate = vi.fn();
  postMessage = vi.fn((request: PackageRequest, transfer: Transferable[]) => {
    this.sent = request;
    // Real ownership transfer, so recovery cannot accidentally reuse detached bytes.
    this.request = structuredClone(request, { transfer });
  });
  reply(result: PackageResults['preflight'], id = this.request!.id): void {
    this.onmessage?.(new MessageEvent<PackageResponse>('message', {
      data: { ...this.request!, id, ok: true, result },
    }));
  }
  fail(): void {
    this.onerror?.(new ErrorEvent('error', { message: 'worker stopped' }));
  }
}

const unchanged: PackageResults['preflight'] = {
  bytes: null, notes: [], notice: null, fontTable: null, vba: NO_VBA,
};
let workers: WorkerDouble[];

beforeEach(() => {
  workers = [];
  constructWorker.mockReset().mockImplementation(function () {
    const worker = new WorkerDouble(); workers.push(worker); return worker;
  });
  vi.stubGlobal('Blob', NativeBlob);
  vi.spyOn(console, 'warn').mockImplementation(() => {});
});
afterEach(() => { vi.unstubAllGlobals(); vi.restoreAllMocks(); });

describe('package worker ownership', () => {
  it('accepts the postflight result without preflight metadata and rejects a foreign operation', async () => {
    const pending = runPackageTask('postflight', new ArrayBuffer(1));
    const completed = vi.fn(); void pending.then(completed);
    const worker = workers[0];
    worker.onmessage!(new MessageEvent<PackageResponse>('message', { data: { ...worker.request!, operation: 'preflight', ok: true, result: unchanged } }));
    await Promise.resolve(); expect(completed).not.toHaveBeenCalled();
    worker.onmessage!(new MessageEvent<PackageResponse>('message', { data: { ...worker.request!, operation: 'postflight', ok: true, result: { bytes: null } } }));
    await expect(pending).resolves.toEqual({ bytes: null });
    expect(worker.terminate).toHaveBeenCalledOnce();
  });

  it('transfers input bytes and terminates after returning metadata', async () => {
    const bytes = new Uint8Array([1, 2, 3]).buffer;
    const pending = runPackageTask('preflight', bytes);
    expect(bytes.byteLength).toBe(0);
    expect(new Uint8Array(workers[0].request!.bytes)).toEqual(new Uint8Array([1, 2, 3]));
    workers[0].reply({ ...unchanged, fontTable: '<fonts/>' });
    await expect(pending).resolves.toEqual({ ...unchanged, fontTable: '<fonts/>' });
    expect(workers[0].terminate).toHaveBeenCalledOnce();
  });

  it('keeps concurrent requests separate and ignores a foreign request id', async () => {
    const a = runPackageTask('preflight', new ArrayBuffer(1));
    const b = runPackageTask('preflight', new ArrayBuffer(2));
    const completed = vi.fn(); void a.then(completed);
    workers[0].reply({ ...unchanged, fontTable: 'wrong' }, workers[1].request!.id);
    workers[1].reply({ ...unchanged, fontTable: 'B' });
    await expect(b).resolves.toMatchObject({ fontTable: 'B' });
    expect(completed).not.toHaveBeenCalled();
    workers[0].reply({ ...unchanged, fontTable: 'A' });
    await expect(a).resolves.toMatchObject({ fontTable: 'A' });
  });

  it('terminates a cancelled opening without accepting its queued result or cancelling B', async () => {
    const controller = new AbortController();
    const a = runPackageTask('preflight', new ArrayBuffer(1), controller.signal);
    const rejected = expect(a).rejects.toMatchObject({ name: 'AbortError' });
    const lateMessage = workers[0].onmessage!;
    const b = runPackageTask('preflight', new ArrayBuffer(2));
    controller.abort();
    lateMessage(new MessageEvent<PackageResponse>('message', {
      data: { ...workers[0].request!, ok: true, result: unchanged },
    }));
    await rejected;
    expect(workers[0].terminate).toHaveBeenCalledOnce();
    expect(workers[1].terminate).not.toHaveBeenCalled();
    workers[1].reply(unchanged);
    await expect(b).resolves.toEqual(unchanged);
  });

  it.each(['error', 'messageerror', 'malformed'] as const)('releases a failed worker (%s)', async (kind) => {
    const pending = runPackageTask('preflight', new ArrayBuffer(1));
    const rejected = expect(pending).rejects.toBeInstanceOf(Error);
    if (kind === 'error') workers[0].fail();
    if (kind === 'messageerror') workers[0].onmessageerror!();
    if (kind === 'malformed') workers[0].reply({ ...unchanged, bytes: 'invalid' } as unknown as PackageResults['preflight']);
    await rejected;
    expect(workers[0].terminate).toHaveBeenCalledOnce();
    expect(workers[0].onmessage).toBeNull();
  });
});

describe('preflight source recovery', () => {
  const fontTable = '<w:fonts xmlns:w="ns"><w:font w:name="David"/></w:fonts>';
  const originalBytes = () => zipStored({
    [SETTINGS_PART]: '<w:settings xmlns:w="ns"><w:defaultTabStop w:val="0"/></w:settings>',
    [FONT_TABLE_PART]: fontTable,
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
    'word/vbaProject.bin': new Uint8Array(600),
    'word/media/untouched.bin': new Uint8Array([11, 22, 33, 44]),
  });

  it('recovers transferred bytes locally without fetching again, retaining repairs, fonts and VBA', async () => {
    const bytes = new Uint8Array(originalBytes());
    const fetch = vi.fn(async () => new Response(bytes)); vi.stubGlobal('fetch', fetch);
    const pending = preflightSource('http://loopback/document.docm');
    await vi.waitFor(() => expect(workers).toHaveLength(1));
    expect(workers[0].sent!.bytes.byteLength).toBe(0);
    workers[0].fail();
    const result = await pending;
    expect(fetch).toHaveBeenCalledOnce();
    expect(result.fontTable).toBe(fontTable);
    expect(result.vba).toMatchObject({ hasMacroPart: true, unreadable: true });
    const repaired = new Uint8Array(await (result.source as Blob).arrayBuffer());
    await expect(readDocxPart(repaired, SETTINGS_PART)).resolves.toContain('w:val="720"');
    for (const name of ['word/vbaProject.bin', 'word/media/untouched.bin']) {
      expect(readZip(repaired)!.find(entry => entry.name === name))
        .toEqual(readZip(bytes)!.find(entry => entry.name === name));
    }
  });

  it.each(['office-macros.docm', 'office-macros-repair.docm', 'auto-open-macros.docm'])(
    'preserves complete VBA metadata in local recovery after transfer (%s)', async (fixture) => {
      const bytes = new Uint8Array(readFileSync(new NodeURL(`../fixtures/preflight/${fixture}`, import.meta.url)));
      const expected = await preflightPackage(bytes);
      expect(expected.vba.hasMacroPart).toBe(true);
      if (fixture === 'auto-open-macros.docm') {
        expect(expected.vba.autoRun).toEqual(['Module1.AutoOpen']);
        expect(expected.vba.warnings).toContain(WARNING_TEXT['auto-run-macros']);
        expect(expected.vba.modules.some(module => module.source.includes('בדיקת תעבורה בלבד'))).toBe(true);
      } else expect(expected.vba.modules.some(module => module.source.includes('say_hello'))).toBe(true);
      if (fixture !== 'office-macros.docm') {
        expect(expected.bytes).not.toBeNull();
        expect(expected.notice).not.toBeNull();
      }
      const pending = preflightSource(new Blob([bytes]));
      await vi.waitFor(() => expect(workers).toHaveLength(1));
      workers[0].fail();
      const result = await pending;
      expect(result.vba).toEqual(expected.vba);
      expect(result.notice).toEqual(expected.notice);
      const repaired = new Uint8Array(await (result.source as Blob).arrayBuffer());
      expect(repaired).toEqual(expected.bytes ?? bytes);
      expect(readZip(repaired)!.find(entry => entry.name === 'word/vbaProject.bin'))
        .toEqual(readZip(bytes)!.find(entry => entry.name === 'word/vbaProject.bin'));
    }
  );

  it('returns complete, independently owned repair buffers for transfer', async () => {
    const result = await preflightPackage(new Uint8Array(originalBytes()));
    expect(result.bytes).not.toBeNull();
    expect(result.bytes!.byteOffset).toBe(0);
    expect(result.bytes!.byteLength).toBe(result.bytes!.buffer.byteLength);
  });

  it('reuses the original Blob when the worker finds no changes', async () => {
    const original = new Blob([originalBytes()]);
    const pending = preflightSource(original);
    await vi.waitFor(() => expect(workers).toHaveLength(1));
    workers[0].reply(unchanged);
    await expect(pending).resolves.toMatchObject({ source: original });
    expect((await pending).source).toBe(original);
  });

  it('does not process locally or fetch again after cancellation', async () => {
    const controller = new AbortController();
    const original = new Blob([originalBytes()]);
    const read = vi.spyOn(original, 'arrayBuffer');
    const pending = preflightSource(original, { signal: controller.signal });
    await vi.waitFor(() => expect(workers).toHaveLength(1));
    controller.abort();
    const result = await pending;
    expect(result.source).toBe(original);
    expect(result.fontTable).toBeNull();
    expect(read).toHaveBeenCalledOnce();
    expect(workers[0].terminate).toHaveBeenCalledOnce();
  });

  it('falls back when construction fails without making a second full Blob read', async () => {
    constructWorker.mockImplementationOnce(() => { throw new Error('Worker unavailable'); });
    const original = new Blob([originalBytes()]);
    const read = vi.spyOn(original, 'arrayBuffer');
    const result = await preflightSource(original);
    expect(read).toHaveBeenCalledOnce();
    expect(result.vba.hasMacroPart).toBe(true);
    await expect(readDocxPart(new Uint8Array(await (result.source as Blob).arrayBuffer()), SETTINGS_PART))
      .resolves.toContain('w:val="720"');
  });

  it('cancels the source fetch before starting a worker', async () => {
    const controller = new AbortController();
    const fetch = vi.fn((_url: string, { signal }: RequestInit) => new Promise<Response>((_resolve, reject) => {
      signal!.addEventListener('abort', () => reject(new DOMException('cancelled', 'AbortError')));
    }));
    vi.stubGlobal('fetch', fetch);
    const pending = preflightSource('http://loopback/document.docx', { signal: controller.signal });
    controller.abort();
    await expect(pending).resolves.toMatchObject({ source: 'http://loopback/document.docx' });
    expect(workers).toHaveLength(0);
    expect(console.warn).not.toHaveBeenCalled();
  });
});
