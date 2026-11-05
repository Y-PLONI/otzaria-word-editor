import { describe, expect, it, vi } from 'vitest';
import {
  forgetDocumentReads,
  getDocumentReadRevision,
  invalidateDocumentReads,
  observeDocumentReadInvalidation,
} from '../../src/engine/document-read-state';

describe('document read revisions', () => {
  it('keeps one revision until this document changes', () => {
    const first = {}, second = {};
    const firstRevision = getDocumentReadRevision(first);
    const secondRevision = getDocumentReadRevision(second);
    expect(getDocumentReadRevision(first)).toBe(firstRevision);
    invalidateDocumentReads(first);
    expect(getDocumentReadRevision(first)).not.toBe(firstRevision);
    expect(getDocumentReadRevision(second)).toBe(secondRevision);
  });

  it('notifies only the changed document, after replacing the revision', () => {
    const first = {}, second = {};
    const before = getDocumentReadRevision(first);
    const seen: object[] = [];
    const other = vi.fn();
    observeDocumentReadInvalidation(first, () => seen.push(getDocumentReadRevision(first)));
    observeDocumentReadInvalidation(second, other);
    invalidateDocumentReads(first);
    expect(seen).toEqual([getDocumentReadRevision(first)]);
    expect(seen[0]).not.toBe(before);
    expect(other).not.toHaveBeenCalled();
  });

  it('unsubscribes each registration independently and idempotently', () => {
    const doc = {}, listener = vi.fn();
    const stopFirst = observeDocumentReadInvalidation(doc, listener);
    const stopSecond = observeDocumentReadInvalidation(doc, listener);
    stopFirst();
    stopFirst();
    invalidateDocumentReads(doc);
    expect(listener).toHaveBeenCalledTimes(1);
    stopSecond();
    invalidateDocumentReads(doc);
    expect(listener).toHaveBeenCalledTimes(1);
  });

  it('disposal rejects the old revision and drops all listeners', () => {
    const doc = {}, listener = vi.fn();
    const pendingRevision = getDocumentReadRevision(doc);
    const stop = observeDocumentReadInvalidation(doc, listener);
    forgetDocumentReads(doc);
    expect(getDocumentReadRevision(doc)).not.toBe(pendingRevision);
    invalidateDocumentReads(doc);
    stop();
    expect(listener).not.toHaveBeenCalled();
  });

  it('isolates a broken readout so it cannot break the mutation or other readouts', () => {
    const doc = {}, next = vi.fn();
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      observeDocumentReadInvalidation(doc, () => { throw new Error('readout failed'); });
      observeDocumentReadInvalidation(doc, next);
      expect(() => invalidateDocumentReads(doc)).not.toThrow();
      expect(next).toHaveBeenCalledTimes(1);
      expect(warning).toHaveBeenCalledTimes(1);
    } finally { warning.mockRestore(); }
  });

  it('ignores absent documents during opening and teardown', () => {
    const listener = vi.fn();
    const stop = observeDocumentReadInvalidation(undefined, listener);
    invalidateDocumentReads(undefined);
    invalidateDocumentReads(null);
    forgetDocumentReads(undefined);
    stop();
    expect(listener).not.toHaveBeenCalled();
  });
});
