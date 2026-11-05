import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Blob as NativeBlob } from 'node:buffer';
import type { SuperDoc } from 'superdoc';
import { zipStored } from '../../scripts/qa/docx-fixtures.mjs';
import { exportDocx, DOCX_MIME } from '../../src/engine/export';
import * as postflight from '../../src/engine/docx-postflight';
import { createSaveCoordinator, type SaveCommitInput, type SaveCoordinator, type SaveCoordinatorDeps } from '../../src/sessions/save-coordinator';
import { createSessionKeeper, DRAFT_DELAY_MS, type KeeperSavePoint, type SessionKeeper } from '../../src/sessions/session-keeper';
import { activeEntry } from '../../src/sessions/session-state';

const task = vi.hoisted(() => vi.fn());
vi.mock('../../src/engine/docx-package-client', () => ({ runPackageTask: task }));

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((accept, fail) => { resolve = accept; reject = fail; });
  return { promise, resolve, reject };
}

function engine(source: Blob | (() => Blob | Promise<Blob>)) {
  const exportDocument = vi.fn(async () => typeof source === 'function' ? source() : source);
  return { host: { export: exportDocument } as unknown as SuperDoc, exportDocument };
}

/** The transport is tested separately; detach real bytes here to exercise recovery ownership. */
function holdTask() {
  const pending = deferred<{ bytes: ArrayBuffer | null }>();
  let request: { bytes: ArrayBuffer; signal?: AbortSignal } | undefined;
  task.mockImplementationOnce((_operation: string, bytes: ArrayBuffer, signal?: AbortSignal) => {
    request = { bytes: structuredClone(bytes, { transfer: [bytes] }), signal };
    const cancel = () => pending.reject(new DOMException('cancelled', 'AbortError'));
    signal?.addEventListener('abort', cancel, { once: true });
    if (signal?.aborted) cancel();
    return pending.promise.finally(() => signal?.removeEventListener('abort', cancel));
  });
  return { ...pending, get request() { return request; } };
}

function repairable() {
  return new Uint8Array(zipStored({
    'word/document.xml': '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>שלום.</w:t></w:r></w:p></w:body></w:document>',
    'word/media/untouched.bin': new Uint8Array([1, 3, 5, 7]),
    'word/vbaProject.bin': new Uint8Array([12, 34, 56, 78]),
  }));
}

const coordinators: SaveCoordinator[] = [];
const keepers: SessionKeeper[] = [];
function saving(host: SuperDoc, onSaved?: SaveCoordinatorDeps['onSaved'], beforeExport?: () => void) {
  const blobs: Blob[] = [];
  let ticket = 0;
  const beginWrite = vi.fn(async (_size: number) => ({ writeToken: `w${++ticket}`, uploadUrl: 'upload' }));
  const upload = vi.fn(async (_url: string, blob: Blob) => { blobs.push(blob); });
  const commit = vi.fn(async (input: SaveCommitInput) => ({ cancelled: false, token: input.targetToken ?? 'new' }));
  const abort = vi.fn(async (_token: string) => {});
  const coordinator = createSaveCoordinator({
    exportDocument: signal => {
      beforeExport?.();
      return exportDocx(host, { signal });
    },
    beginWrite, upload, commit, abort, onSaved,
  });
  coordinators.push(coordinator);
  coordinator.adoptTarget({ token: 'A', name: 'A.docx' });
  return { coordinator, beginWrite, upload, commit, abort, blobs };
}

function savingWithDrafts(host: SuperDoc) {
  let keeper!: SessionKeeper;
  let savePoint!: KeeperSavePoint;
  let noted = Promise.resolve();
  const h = saving(host, info => {
    // Same export point and save binding as App.initSaveCoordinator.
    noted = keeper.noteSaved(savePoint, {
      sourceSize: info.size,
      document: { token: info.token, name: info.name, writable: true },
    });
  }, () => { savePoint = keeper.captureSave(); });
  const drafts: Blob[] = [];
  const writeDraft = vi.fn(async (blob: Blob, _path: string) => { drafts.push(blob); return 'written' as const; });
  const settleSave = vi.fn(() => h.coordinator.settled());
  keeper = createSessionKeeper({
    id: 'typing-during-postflight', persist: async () => {},
    exportDocument: () => exportDocx(host, { postflight: false }),
    writeDraft, removeDraft: async () => {}, draftPaths: ['draft-a.docx', 'draft-b.docx'],
    readCaret: async () => null,
    isDirty: () => h.coordinator.snapshot.isDirty,
    isSaving: () => h.coordinator.snapshot.isSaving,
    isHeld: () => h.coordinator.isHeld,
    settleSave,
  });
  keepers.push(keeper);
  keeper.setDocument({ token: 'A', name: 'A.docx', writable: true });
  h.coordinator.setAutosaveEnabled(false);
  return { ...h, keeper, drafts, writeDraft, settleSave, get noted() { return noted; } };
}

beforeEach(() => {
  task.mockReset().mockResolvedValue({ bytes: null });
  vi.stubGlobal('Blob', NativeBlob);
  vi.spyOn(console, 'warn').mockImplementation(() => {});
});
afterEach(() => {
  keepers.splice(0).forEach(keeper => keeper.dispose());
  coordinators.splice(0).forEach(coordinator => coordinator.dispose());
  vi.restoreAllMocks(); vi.unstubAllGlobals(); vi.useRealTimers();
});

describe('export postflight ownership and failure policy', () => {
  it('runs postflight on the exported snapshot, preserves MIME, and never re-exports', async () => {
    const original = new Blob(['snapshot'], { type: 'application/vnd.ms-word.document.macroEnabled.12' });
    const source = engine(original); const held = holdTask();
    const result = exportDocx(source.host);
    await vi.waitFor(() => expect(held.request).toBeDefined());
    expect(new TextDecoder().decode(held.request!.bytes)).toBe('snapshot');
    expect(task.mock.calls[0][0]).toBe('postflight');
    expect(source.exportDocument).toHaveBeenCalledWith({ exportType: ['docx'], triggerDownload: false });
    held.resolve({ bytes: new TextEncoder().encode('repaired snapshot').buffer });
    const output = await result;
    expect(await output.text()).toBe('repaired snapshot');
    expect(output.type).toBe(original.type);
    expect(source.exportDocument).toHaveBeenCalledOnce();
  });

  it('retains the original Blob when unchanged and bypasses postflight for recovery drafts', async () => {
    const original = new Blob(['snapshot']); const source = engine(original);
    expect(await exportDocx(source.host)).toBe(original);
    expect(task).toHaveBeenCalledOnce();
    task.mockClear();
    expect(await exportDocx(source.host, { postflight: false })).toBe(original);
    expect(task).not.toHaveBeenCalled();
  });

  it.each([false, true])('recovers from worker failure using the same export (transferred=%s)', async (transfer) => {
    const bytes = repairable();
    const expected = await postflight.postflightDocx(bytes);
    expect(expected).not.toBeNull();
    const original = new Blob([bytes], { type: DOCX_MIME });
    const read = vi.spyOn(original, 'arrayBuffer'); const source = engine(original);
    task.mockImplementationOnce((_operation: string, buffer: ArrayBuffer) => {
      if (transfer) structuredClone(buffer, { transfer: [buffer] });
      throw new Error('worker unavailable');
    });
    const output = await exportDocx(source.host);
    expect(new Uint8Array(await output.arrayBuffer())).toEqual(expected);
    expect(read).toHaveBeenCalledTimes(transfer ? 2 : 1);
    expect(source.exportDocument).toHaveBeenCalledOnce();
  });

  it('keeps the original Blob when local recovery fails', async () => {
    const original = new Blob([repairable()]);
    task.mockRejectedValueOnce(new Error('worker failed'));
    vi.spyOn(postflight, 'postflightDocx').mockRejectedValueOnce(new DOMException('codec stopped', 'AbortError'));
    expect(await exportDocx(engine(original).host)).toBe(original);
  });

  it('treats a Blob AbortError with a live signal as a read failure, not caller cancellation', async () => {
    const original = new Blob(['snapshot']);
    const signal = new AbortController().signal;
    vi.spyOn(original, 'arrayBuffer').mockRejectedValueOnce(new DOMException('read stopped', 'AbortError'));
    expect(await exportDocx(engine(original).host, { signal })).toBe(original);
    expect(task).not.toHaveBeenCalled();
  });

  it('still repairs after a worker AbortError when the owning signal is live', async () => {
    const bytes = repairable(); const expected = await postflight.postflightDocx(bytes);
    const original = new Blob([bytes]); const signal = new AbortController().signal;
    task.mockRejectedValueOnce(new DOMException('worker transport stopped', 'AbortError'));
    const result = await exportDocx(engine(original).host, { signal });
    expect(new Uint8Array(await result.arrayBuffer())).toEqual(expected);
  });

  it('does not start engine export or a worker after caller cancellation', async () => {
    const controller = new AbortController(); controller.abort();
    const source = engine(new Blob(['snapshot']));
    await expect(exportDocx(source.host, { signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    expect(source.exportDocument).not.toHaveBeenCalled(); expect(task).not.toHaveBeenCalled();
  });

  it('does not start package work when cancellation happens during engine export', async () => {
    const pending = deferred<Blob>(); const source = engine(() => pending.promise);
    const controller = new AbortController();
    const result = exportDocx(source.host, { signal: controller.signal });
    const rejected = expect(result).rejects.toMatchObject({ name: 'AbortError' });
    controller.abort(); pending.resolve(new Blob(['old snapshot']));
    await rejected; expect(task).not.toHaveBeenCalled();
  });

  it('does not start package work when cancellation happens while reading the export Blob', async () => {
    const pending = deferred<ArrayBuffer>(); const original = new Blob(['snapshot']);
    const read = vi.spyOn(original, 'arrayBuffer').mockReturnValueOnce(pending.promise);
    const controller = new AbortController();
    const result = exportDocx(engine(original).host, { signal: controller.signal });
    const rejected = expect(result).rejects.toMatchObject({ name: 'AbortError' });
    await vi.waitFor(() => expect(read).toHaveBeenCalledOnce());
    controller.abort(); pending.resolve(new ArrayBuffer(1));
    await rejected; expect(task).not.toHaveBeenCalled();
    expect(console.warn).not.toHaveBeenCalled();
  });

  it('does not read the Blob again or process locally after cancelling transferred work', async () => {
    const original = new Blob([repairable()]); const read = vi.spyOn(original, 'arrayBuffer');
    const local = vi.spyOn(postflight, 'postflightDocx');
    const controller = new AbortController(); const held = holdTask();
    const result = exportDocx(engine(original).host, { signal: controller.signal });
    const rejected = expect(result).rejects.toMatchObject({ name: 'AbortError' });
    await vi.waitFor(() => expect(held.request).toBeDefined());
    controller.abort(); await rejected;
    held.resolve({ bytes: new TextEncoder().encode('late output').buffer });
    expect(read).toHaveBeenCalledOnce(); expect(local).not.toHaveBeenCalled();
    expect(console.warn).not.toHaveBeenCalled();
  });
});

describe('save coordination while postflight is pending', () => {
  it.each([false, true])('keeps the recovery draft due after typing during an old snapshot save (saveAs=%s)', async (forceSaveAs) => {
    vi.useFakeTimers();
    let text = 'before typing';
    const source = engine(() => new Blob([text]));
    const h = savingWithDrafts(source.host);
    const { keeper, drafts, writeDraft } = h;
    h.coordinator.markDirty(); keeper.noteChange();
    const held = holdTask();
    const saved = h.coordinator.saveNow({ forceSaveAs });
    await vi.waitFor(() => expect(held.request).toBeDefined());
    text = 'new typing while the snapshot worker is pending';
    h.coordinator.markDirty(); keeper.noteChange();
    const changedAt = Date.now();
    await vi.advanceTimersByTimeAsync(1000);
    expect(writeDraft).not.toHaveBeenCalled();
    held.resolve({ bytes: null });
    await expect(saved).resolves.toMatchObject({ status: 'saved' });
    await h.noted;
    expect(await h.blobs[0].text()).toBe('before typing');
    expect(h.coordinator.snapshot.isDirty).toBe(true);
    expect(keeper.hasUnwrittenWork, 'newer typing must remain eligible for a recovery draft').toBe(true);
    expect(task).toHaveBeenCalledOnce(); task.mockClear();
    await vi.advanceTimersByTimeAsync(DRAFT_DELAY_MS - (Date.now() - changedAt) - 1);
    expect(writeDraft).not.toHaveBeenCalled();
    await vi.advanceTimersByTimeAsync(1);
    await vi.waitFor(() => expect(writeDraft).toHaveBeenCalledOnce());
    expect(await drafts[0].text()).toBe(text);
    expect(activeEntry(keeper.state)?.draft?.documentToken).toBe(forceSaveAs ? 'new' : 'A');
    expect(task, 'recovery drafts must continue bypassing postflight').not.toHaveBeenCalled();
    expect(h.commit).toHaveBeenCalledOnce();
  });

  it('retains a newer queued draft when the older snapshot commit is recorded later', async () => {
    let text = 'old saved snapshot';
    const h = savingWithDrafts(engine(() => new Blob([text])).host);
    h.coordinator.markDirty(); h.keeper.noteChange();
    const held = holdTask(); const saved = h.coordinator.saveNow();
    await vi.waitFor(() => expect(held.request).toBeDefined());
    text = 'newer recovery snapshot';
    h.coordinator.markDirty(); h.keeper.noteChange();
    const queuedDraft = h.keeper.flush();
    await vi.waitFor(() => expect(h.settleSave).toHaveBeenCalledOnce());
    expect(h.writeDraft).not.toHaveBeenCalled();
    held.resolve({ bytes: null });
    await expect(saved).resolves.toMatchObject({ status: 'saved' });
    await queuedDraft; await h.noted;
    expect(await h.blobs[0].text()).toBe('old saved snapshot');
    expect(h.writeDraft).toHaveBeenCalledOnce();
    expect(await h.drafts[0].text()).toBe(text);
    expect(h.coordinator.snapshot.isDirty).toBe(true);
    expect(activeEntry(h.keeper.state)?.draft, 'the newer draft must stay discoverable for recovery').not.toBeNull();
    expect(h.keeper.hasUnwrittenWork).toBe(false);
    expect(task).toHaveBeenCalledOnce();
  });

  it('keeps recovery ownership across consecutive Save As operations with newer typing', async () => {
    vi.useFakeTimers();
    let text = 'snapshot A';
    const h = savingWithDrafts(engine(() => new Blob([text])).host);
    // Restore/engine events need not advance the two independent revision counters equally.
    h.coordinator.markDirty();
    h.coordinator.markDirty(); h.keeper.noteChange();
    h.commit.mockResolvedValueOnce({ cancelled: false, token: 'B' });
    h.commit.mockResolvedValueOnce({ cancelled: false, token: 'C' });
    const firstWorker = holdTask();
    const first = h.coordinator.saveNow({ forceSaveAs: true });
    await vi.waitFor(() => expect(firstWorker.request).toBeDefined());
    text = 'snapshot B'; h.coordinator.markDirty(); h.keeper.noteChange();
    firstWorker.resolve({ bytes: null });
    await expect(first).resolves.toMatchObject({ status: 'saved', token: 'B' });
    await h.noted;
    expect(activeEntry(h.keeper.state)?.document?.token).toBe('B');

    const secondWorker = holdTask();
    const second = h.coordinator.saveNow({ forceSaveAs: true });
    await vi.waitFor(() => expect(secondWorker.request).toBeDefined());
    text = 'new typing C'; h.coordinator.markDirty(); h.keeper.noteChange();
    const changedAt = Date.now();
    await vi.advanceTimersByTimeAsync(1_000);
    secondWorker.resolve({ bytes: null });
    await expect(second).resolves.toMatchObject({ status: 'saved', token: 'C' });
    await h.noted;
    expect(await h.blobs[0].text()).toBe('snapshot A');
    expect(await h.blobs[1].text()).toBe('snapshot B');
    expect(h.coordinator.snapshot.isDirty).toBe(true);
    expect(h.keeper.hasUnwrittenWork).toBe(true);
    await vi.advanceTimersByTimeAsync(DRAFT_DELAY_MS - (Date.now() - changedAt));
    await vi.waitFor(() => expect(h.writeDraft).toHaveBeenCalledOnce());
    expect(await h.drafts[0].text()).toBe(text);
    const draft = activeEntry(h.keeper.state)?.draft;
    expect(draft).toMatchObject({ documentToken: 'C', sourceSize: h.blobs[1].size });
    expect(h.keeper.hasUnwrittenWork).toBe(false);
    expect(task).toHaveBeenCalledTimes(2);

    await expect(h.coordinator.saveNow()).resolves.toMatchObject({ status: 'saved', token: 'C' });
    await h.noted;
    expect(await h.blobs[2].text()).toBe(text);
    expect(h.coordinator.snapshot.isDirty).toBe(false);
    expect(activeEntry(h.keeper.state)?.draft).toBeNull();
    expect(activeEntry(h.keeper.state)?.previousDraft).toEqual(draft);
  });

  it('keeps typing dirty after an old snapshot commits, then saves the newer snapshot', async () => {
    let text = 'before typing'; const source = engine(() => new Blob([text]));
    const h = saving(source.host); const held = holdTask();
    h.coordinator.markDirty();
    const first = h.coordinator.saveNow();
    await vi.waitFor(() => expect(held.request).toBeDefined());
    text = 'after typing'; h.coordinator.markDirty();
    expect(h.beginWrite).not.toHaveBeenCalled();
    held.resolve({ bytes: null });
    await expect(first).resolves.toMatchObject({ status: 'saved', token: 'A' });
    expect(await h.blobs[0].text()).toBe('before typing');
    expect(h.coordinator.snapshot.isDirty).toBe(true);
    await expect(h.coordinator.saveNow()).resolves.toMatchObject({ status: 'saved', token: 'A' });
    expect(await h.blobs[1].text()).toBe('after typing');
    expect(h.coordinator.snapshot.isDirty).toBe(false);
    expect(source.exportDocument).toHaveBeenCalledTimes(2);
  });

  it.each(['reset', 'dispose'] as const)('cancels package work on %s before allocating an upload ticket', async (operation) => {
    const source = engine(new Blob(['old snapshot'])); const h = saving(source.host);
    const held = holdTask(); h.coordinator.markDirty();
    const result = h.coordinator.saveNow();
    await vi.waitFor(() => expect(held.request).toBeDefined());
    if (operation === 'reset') {
      h.coordinator.reset({ token: 'B', name: 'B.docx' }); h.coordinator.markDirty();
    } else h.coordinator.dispose();
    await expect(result).resolves.toEqual({ status: 'stale' });
    expect(held.request!.signal!.aborted).toBe(true);
    expect(h.beginWrite).not.toHaveBeenCalled(); expect(h.upload).not.toHaveBeenCalled();
    expect(h.commit).not.toHaveBeenCalled(); expect(h.abort).not.toHaveBeenCalled();
    expect(h.coordinator.snapshot.isDirty).toBe(true);
    expect(console.warn).not.toHaveBeenCalled();
  });

  it('skips an obsolete queued epoch before starting export or a worker', async () => {
    let text = 'A'; const source = engine(() => new Blob([text])); const h = saving(source.host);
    const held = holdTask(); h.coordinator.markDirty();
    const a = h.coordinator.saveNow(); await vi.waitFor(() => expect(held.request).toBeDefined());
    h.coordinator.reset({ token: 'B', name: 'B.docx' }); h.coordinator.markDirty();
    const b = h.coordinator.saveNow();
    text = 'C'; h.coordinator.reset({ token: 'C', name: 'C.docx' }); h.coordinator.markDirty();
    const c = h.coordinator.saveNow();
    await expect(a).resolves.toEqual({ status: 'stale' });
    await expect(b).resolves.toEqual({ status: 'stale' });
    await expect(c).resolves.toMatchObject({ status: 'saved', token: 'C' });
    expect(source.exportDocument).toHaveBeenCalledTimes(2); expect(task).toHaveBeenCalledTimes(2);
    expect(h.blobs).toHaveLength(1); expect(await h.blobs[0].text()).toBe('C');
    expect(h.commit).toHaveBeenCalledWith(expect.objectContaining({ targetToken: 'C' }));
  });

  it('does not cancel an in-flight save when autosave is disabled or temporarily held', async () => {
    const h = saving(engine(new Blob(['snapshot'])).host); const held = holdTask();
    h.coordinator.markDirty(); const result = h.coordinator.saveNow();
    await vi.waitFor(() => expect(held.request).toBeDefined());
    h.coordinator.setAutosaveEnabled(false); const release = h.coordinator.hold(); release();
    expect(held.request!.signal!.aborted).toBe(false);
    held.resolve({ bytes: null });
    await expect(result).resolves.toMatchObject({ status: 'saved' });
  });

  it('cancelling one document does not cancel another document’s package work', async () => {
    const a = saving(engine(new Blob(['A'])).host); const heldA = holdTask();
    a.coordinator.markDirty(); const first = a.coordinator.saveNow();
    await vi.waitFor(() => expect(heldA.request).toBeDefined());
    const b = saving(engine(new Blob(['B'])).host); const heldB = holdTask();
    b.coordinator.markDirty(); const second = b.coordinator.saveNow();
    await vi.waitFor(() => expect(heldB.request).toBeDefined());
    a.coordinator.dispose();
    await expect(first).resolves.toEqual({ status: 'stale' });
    expect(heldB.request!.signal!.aborted).toBe(false);
    heldB.resolve({ bytes: null });
    await expect(second).resolves.toMatchObject({ status: 'saved' });
    expect(await b.blobs[0].text()).toBe('B'); expect(a.commit).not.toHaveBeenCalled();
  });
});
