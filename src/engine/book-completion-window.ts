import { buildSectionCache, type SectionWordCache } from './book-completion';
import { WORD_INNER } from './word-selection';

/** The reader's effective text and its hash (or the text itself on older hosts). */
export interface CompletionSectionText {
  text: string;
  fingerprint: string;
}

export type ReadCompletionSection = (index: number) => Promise<CompletionSectionText | null>;

export interface CompletionWindowSection {
  index: number;
  /** Half-open offsets in `cache.sourceText`, excluding the joining newline. */
  start: number;
  end: number;
  fingerprint: string;
}

export interface CompletionWindow {
  cache: SectionWordCache;
  sections: CompletionWindowSection[];
  centerIndex: number;
  centerFingerprint: string;
}

export type CompletionRangeValidation = 'valid' | 'changed' | 'failed' | 'cancelled';

function cutsWord(text: string, offset: number): boolean {
  return offset > 0 && offset < text.length &&
    WORD_INNER.test(text[offset - 1]) && WORD_INNER.test(text[offset]);
}

/**
 * Builds a replacement without modifying the published window. Read the center
 * first so a large preceding section cannot use up its text budget. Forward
 * context has priority; prepend up to five preceding sections with the remainder.
 * `null` at the center fails the load; at a neighbor it ends that direction.
 */
export async function loadCompletionWindow(
  centerIndex: number,
  readSection: ReadCompletionSection,
  isCurrent: () => boolean,
  charCap = 20_000,
): Promise<CompletionWindow | null> {
  if (!Number.isSafeInteger(centerIndex) || centerIndex < 0 ||
      !Number.isSafeInteger(charCap) || charCap <= 0 || !isCurrent()) return null;

  try {
    const center = await readSection(centerIndex);
    if (!isCurrent() || center === null) return null;

    const loaded = [{ index: centerIndex, ...center, text: center.text.slice(0, charCap) }];
    let remaining = charCap - loaded[0].text.length;
    let clippedFirstWord = false;
    let clippedLastWord = cutsWord(center.text, charCap);

    // Reserve one character for every newline, including around empty sections.
    for (let offset = 1; offset <= 25 && remaining > 1; offset += 1) {
      if (!isCurrent()) return null;
      const index = centerIndex + offset;
      const section = await readSection(index);
      if (!isCurrent()) return null;
      if (section === null) break;
      const text = section.text.slice(0, remaining - 1);
      loaded.push({ index, ...section, text });
      remaining -= text.length + 1;
      if (text.length < section.text.length) {
        clippedLastWord = cutsWord(section.text, text.length);
        break;
      }
    }

    for (let offset = 1; offset <= 5 && offset <= centerIndex && remaining > 1; offset += 1) {
      if (!isCurrent()) return null;
      const index = centerIndex - offset;
      const section = await readSection(index);
      if (!isCurrent()) return null;
      if (section === null) break;
      // A partial preceding section must keep the text adjacent to the center.
      const text = section.text.slice(Math.max(0, section.text.length - (remaining - 1)));
      loaded.unshift({ index, ...section, text });
      remaining -= text.length + 1;
      if (text.length < section.text.length) {
        clippedFirstWord = cutsWord(section.text, section.text.length - text.length);
        break;
      }
    }

    if (!isCurrent()) return null;
    let cursor = 0;
    const sections = loaded.map(({ index, text, fingerprint }) => {
      const start = cursor;
      cursor += text.length + 1;
      return { index, start, end: start + text.length, fingerprint };
    });
    const cache = buildSectionCache(loaded.map((section) => section.text).join('\n'));
    // A character cap may split a word or its marks. Keep the offsets, but never
    // offer that boundary fragment as though it were a complete source word.
    if (clippedLastWord && cache.words[cache.words.length - 1]?.end === cache.normalizedText.length) {
      cache.words.pop();
    }
    if (clippedFirstWord && cache.words[0]?.start === 0) cache.words.shift();
    return {
      cache,
      sections,
      centerIndex,
      centerFingerprint: center.fingerprint,
    };
  } catch {
    return null;
  }
}

/**
 * Validate the source of [fromWord, toWord), including sections containing only
 * punctuation or whitespace between those words. No unrelated section is read.
 * A failed/cancelled check must not authorize displaying or inserting a match.
 */
export async function validateCompletionRange(
  window: CompletionWindow,
  fromWord: number,
  toWord: number,
  readSection: ReadCompletionSection,
  isCurrent: () => boolean,
): Promise<CompletionRangeValidation> {
  if (!isCurrent()) return 'cancelled';
  const { cache } = window;
  if (!Number.isSafeInteger(fromWord) || !Number.isSafeInteger(toWord) ||
      fromWord < 0 || fromWord >= toWord || toWord > cache.words.length) return 'failed';

  const start = cache.toOriginal[cache.words[fromWord].start];
  const end = cache.toOriginal[cache.words[toWord - 1].end - 1] + 1;
  const sections = window.sections.filter((section) => section.end > start && section.start < end);
  if (sections.length === 0 || sections[0].start > start || sections[sections.length - 1].end < end) {
    return 'failed';
  }

  const checked = new Set<number>();
  try {
    for (const section of sections) {
      if (!isCurrent()) return 'cancelled';
      if (checked.has(section.index)) continue;
      checked.add(section.index);
      const current = await readSection(section.index);
      if (!isCurrent()) return 'cancelled';
      if (current === null) return 'failed';
      if (current.fingerprint !== section.fingerprint) return 'changed';
    }
    return 'valid';
  } catch {
    return isCurrent() ? 'failed' : 'cancelled';
  }
}
