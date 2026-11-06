import { describe, expect, it, vi } from 'vitest';
import {
  loadCompletionWindow,
  validateCompletionRange,
  type CompletionSectionText,
  type CompletionWindow,
} from '../../src/engine/book-completion-window';

const current = () => true;
const section = (text: string, fingerprint = text): CompletionSectionText => ({ text, fingerprint });

async function windowFrom(
  content: Record<number, CompletionSectionText>,
  centerIndex: number,
  charCap?: number,
): Promise<CompletionWindow> {
  const window = await loadCompletionWindow(
    centerIndex,
    async (index) => content[index] ?? null,
    current,
    charCap,
  );
  if (window === null) throw new Error('Expected a loaded window');
  return window;
}

describe('loadCompletionWindow', () => {
  it('reads the center first and records source offsets in book order', async () => {
    const content: Record<number, CompletionSectionText> = {
      9: section('before'),
      10: section('center', 'center-hash'),
      11: section('after'),
    };
    const read = vi.fn(async (index: number) => content[index] ?? null);
    const window = await loadCompletionWindow(10, read, current);

    expect(read.mock.calls.map(([index]) => index)).toEqual([10, 11, 12, 9, 8]);
    expect(window?.cache.sourceText).toBe('before\ncenter\nafter');
    expect(window?.sections).toEqual([
      { index: 9, start: 0, end: 6, fingerprint: 'before' },
      { index: 10, start: 7, end: 13, fingerprint: 'center-hash' },
      { index: 11, start: 14, end: 19, fingerprint: 'after' },
    ]);
    expect(window?.centerIndex).toBe(10);
    expect(window?.centerFingerprint).toBe('center-hash');
  });

  it('keeps the center when a preceding section exceeds the entire character cap', async () => {
    const content: Record<number, CompletionSectionText> = {
      4: section(`${'x'.repeat(25_000)} previous ending`, 'previous-hash'),
      5: section('central phrase', 'center-hash'),
    };
    const read = vi.fn(async (index: number) => content[index] ?? null);
    const window = await loadCompletionWindow(5, read, current, 40);

    expect(read.mock.calls.map(([index]) => index)).toEqual([5, 6, 4]);
    expect(window?.cache.sourceText).toHaveLength(40);
    expect(window?.cache.sourceText).toMatch(/previous ending\ncentral phrase$/);
    expect(window?.centerFingerprint).toBe('center-hash');
    const center = window!.sections.find(({ index }) => index === 5)!;
    expect(window?.cache.sourceText.slice(center.start, center.end)).toBe('central phrase');
  });

  it('bounds the center itself and does not read neighbors when the cap is full', async () => {
    const read = vi.fn(async () => section('a'.repeat(20_001), 'full-section-hash'));
    const window = await loadCompletionWindow(7, read, current);

    expect(read).toHaveBeenCalledTimes(1);
    expect(window?.cache.sourceText).toHaveLength(20_000);
    expect(window?.sections).toEqual([
      { index: 7, start: 0, end: 20_000, fingerprint: 'full-section-hash' },
    ]);
    expect(window?.centerFingerprint).toBe('full-section-hash');
  });

  it('reads no more than five previous and twenty-five following sections', async () => {
    const read = vi.fn(async (index: number) => section(`word ${index}`));
    const window = await loadCompletionWindow(10, read, current);

    expect(read).toHaveBeenCalledTimes(31);
    expect(read.mock.calls[0]).toEqual([10]);
    expect(window?.sections.map(({ index }) => index)).toEqual(
      Array.from({ length: 31 }, (_, offset) => offset + 5),
    );
  });

  it('does not index a word cut by the character cap at either edge', async () => {
    const following = await windowFrom({ 0: section('אבי'), 1: section('אמר בראשית') }, 0, 11);
    const preceding = await windowFrom({ 0: section('בראשית אמר'), 1: section('אבי') }, 1, 11);
    const words = (window: CompletionWindow) => window.cache.words.map(({ start, end }) =>
      window.cache.normalizedText.slice(start, end));

    expect(following.cache.sourceText).toBe('אבי\nאמר ברא');
    expect(words(following)).toEqual(['אבי', 'אמר']);
    expect(preceding.cache.sourceText).toBe('שית אמר\nאבי');
    expect(words(preceding)).toEqual(['אמר', 'אבי']);
  });

  it('does not discard a complete word when only detached marks remain at the clipped edge', async () => {
    const window = await windowFrom({ 0: section('אְ ב'), 1: section('ג') }, 1, 5);
    expect(window.cache.sourceText).toBe('ְ ב\nג');
    expect(window.cache.words.map(({ start, end }) => window.cache.normalizedText.slice(start, end)))
      .toEqual(['ב', 'ג']);
  });

  it('includes joining newlines in the cap and keeps empty section provenance', async () => {
    const window = await windowFrom({
      0: section(''),
      1: section('א'),
      2: section(''),
      3: section('ב'),
    }, 0, 5);

    expect(window.cache.sourceText).toBe('\nא\n\nב');
    expect(window.cache.sourceText.length).toBeLessThanOrEqual(5);
    expect(window.centerFingerprint).toBe('');
    expect(window.sections).toEqual([
      { index: 0, start: 0, end: 0, fingerprint: '' },
      { index: 1, start: 1, end: 2, fingerprint: 'א' },
      { index: 2, start: 3, end: 3, fingerprint: '' },
      { index: 3, start: 4, end: 5, fingerprint: 'ב' },
    ]);
  });

  it('ends each direction at null without losing the other direction', async () => {
    const read = vi.fn(async (index: number) => index >= 3 && index <= 5 ? section(`${index}`) : null);
    const window = await loadCompletionWindow(5, read, current);

    expect(read.mock.calls.map(([index]) => index)).toEqual([5, 6, 4, 3, 2]);
    expect(window?.cache.sourceText).toBe('3\n4\n5');
  });

  it('returns no replacement when the center is unavailable or a read throws', async () => {
    const missing = vi.fn(async () => null);
    expect(await loadCompletionWindow(3, missing, current)).toBeNull();
    expect(missing).toHaveBeenCalledTimes(1);

    const failing = vi.fn(async (index: number) => {
      if (index === 4) throw new Error('Reader disconnected');
      return section('center');
    });
    expect(await loadCompletionWindow(3, failing, current)).toBeNull();
    expect(failing.mock.calls.map(([index]) => index)).toEqual([3, 4]);
  });

  it('stops immediately if the reader identity changes during an awaited read', async () => {
    let active = true;
    const read = vi.fn(async (index: number) => {
      if (index === 4) active = false;
      return section('text');
    });

    expect(await loadCompletionWindow(3, read, () => active)).toBeNull();
    expect(read.mock.calls.map(([index]) => index)).toEqual([3, 4]);
    read.mockClear();
    expect(await loadCompletionWindow(3, read, () => active)).toBeNull();
    expect(read).not.toHaveBeenCalled();
  });

  it.each([
    [-1, 20_000],
    [1.5, 20_000],
    [Number.NaN, 20_000],
    [0, 0],
    [0, Number.POSITIVE_INFINITY],
  ])('rejects invalid center/cap (%s, %s) before calling the reader', async (center, cap) => {
    const read = vi.fn(async () => section('text'));
    expect(await loadCompletionWindow(center, read, current, cap)).toBeNull();
    expect(read).not.toHaveBeenCalled();
  });
});

describe('validateCompletionRange', () => {
  it('validates the matched body without polling its unchanged heading or other sections', async () => {
    const content: Record<number, CompletionSectionText> = {
      10: section('פרק א', 'heading'),
      11: section('בְּרֵאשִׁית בָּרָא', 'body-with-nikud'),
      12: section('המשך אחר', 'unrelated'),
    };
    const window = await windowFrom(content, 10);
    const read = vi.fn(async (index: number) => content[index] ?? null);

    expect(await validateCompletionRange(window, 2, 4, read, current)).toBe('valid');
    expect(read.mock.calls.map(([index]) => index)).toEqual([11]);
    content[11] = section('בראשית ברא', 'body-without-nikud');
    read.mockClear();
    expect(await validateCompletionRange(window, 2, 4, read, current)).toBe('changed');
    expect(read.mock.calls.map(([index]) => index)).toEqual([11]);
  });

  it('also validates punctuation, whitespace and empty sections between the endpoint words', async () => {
    const content: Record<number, CompletionSectionText> = {
      0: section('א'),
      1: section(''),
      2: section('  —  '),
      3: section('ב'),
      4: section('מיותר'),
    };
    const window = await windowFrom(content, 0);
    const read = vi.fn(async (index: number) => content[index] ?? null);

    expect(await validateCompletionRange(window, 0, 2, read, current)).toBe('valid');
    expect(read.mock.calls.map(([index]) => index)).toEqual([0, 1, 2, 3]);
    content[1] = section('מילה חדשה');
    read.mockClear();
    expect(await validateCompletionRange(window, 0, 2, read, current)).toBe('changed');
    expect(read.mock.calls.map(([index]) => index)).toEqual([0, 1]);
  });

  it('treats an empty rendered section as a real change', async () => {
    const window = await windowFrom({ 0: section('מילה', 'old') }, 0);
    expect(await validateCompletionRange(window, 0, 1, async () => section('', ''), current)).toBe('changed');
  });

  it('never authorizes a match after an unreadable section or a bridge error', async () => {
    const window = await windowFrom({ 0: section('א'), 1: section('ב') }, 0);
    const missing = vi.fn(async (index: number) => index === 0 ? section('א') : null);
    expect(await validateCompletionRange(window, 0, 2, missing, current)).toBe('failed');
    expect(missing).toHaveBeenCalledTimes(2);
    expect(await validateCompletionRange(window, 0, 2, async () => {
      throw new Error('Reader disconnected');
    }, current)).toBe('failed');
  });

  it('cancels during an awaited validation and does not read subsequent sections', async () => {
    const window = await windowFrom({ 0: section('א'), 1: section('ב') }, 0);
    let active = true;
    const read = vi.fn(async () => {
      active = false;
      return section('א');
    });

    expect(await validateCompletionRange(window, 0, 2, read, () => active)).toBe('cancelled');
    expect(read).toHaveBeenCalledTimes(1);
    read.mockClear();
    expect(await validateCompletionRange(window, 0, 2, read, () => active)).toBe('cancelled');
    expect(read).not.toHaveBeenCalled();
  });

  it.each([
    [-1, 1],
    [0, 0],
    [1, 0],
    [0, 3],
    [0.5, 1],
    [0, Number.NaN],
  ])('rejects invalid word ranges [%s, %s) without reading', async (from, to) => {
    const window = await windowFrom({ 0: section('א ב') }, 0);
    const read = vi.fn(async () => section('א ב'));
    expect(await validateCompletionRange(window, from, to, read, current)).toBe('failed');
    expect(read).not.toHaveBeenCalled();
  });

  it('rejects a range in an empty cache without reading', async () => {
    const window = await windowFrom({ 0: section('') }, 0);
    const read = vi.fn(async () => section(''));
    expect(await validateCompletionRange(window, 0, 1, read, current)).toBe('failed');
    expect(read).not.toHaveBeenCalled();
  });
});
