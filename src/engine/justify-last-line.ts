/**
 * Work around SuperDoc #4136 without changing document text or formatting.
 * Only a final, verified justified line is eligible. The renderer labels that
 * line text-align:right/left, so CSS alone cannot distinguish it from an
 * intentionally ragged paragraph. See justify-last-line-source.ts.
 */
import { observeDocumentReadInvalidation } from './document-read-state';
import { createJustificationReader, type JustificationDocument } from './justify-last-line-source';

export const LINE_SELECTOR = '.superdoc-line';
const FRAGMENT_SELECTOR = '.superdoc-fragment';
const TEXT_RUN_SELECTOR = '.superdoc-text-run';
const MIN_SPILL_PX = 0.5;
const OURS = 'oweSpillSpacing';
const SPACES = /[ \u00a0]/g;
const ownedStyles = new WeakMap<HTMLElement, string | null>();

/** Real text only: list markers, marker suffixes, tab placeholders and images
 * must not contribute either ink bounds or compressible spaces. Trim the
 * paragraph's outside whitespace for bounds; a trailing space is not an
 * overflowing word. Leading spaces still count: shrinking them moves all
 * following ink, whereas trailing spaces are after the measured ending edge.
 */
function textRanges(line: HTMLElement): { ranges: Range[]; spaces: number; text: string } {
  const view = line.ownerDocument.defaultView!;
  const nodes: Text[] = [];
  for (const run of line.querySelectorAll(TEXT_RUN_SELECTOR)) {
    const walker = line.ownerDocument.createTreeWalker(run, view.NodeFilter.SHOW_TEXT);
    let node: Node | null;
    while ((node = walker.nextNode())) nodes.push(node as Text);
  }
  const text = nodes.map((node) => node.data).join('');
  const start = text.length - text.trimStart().length;
  const end = text.trimEnd().length;
  const ranges: Range[] = [];
  let offset = 0;
  for (const node of nodes) {
    const from = Math.max(0, start - offset);
    const to = Math.min(node.length, end - offset);
    if (to > from) {
      const range = line.ownerDocument.createRange();
      range.setStart(node, from);
      range.setEnd(node, to);
      ranges.push(range);
    }
    offset += node.length;
  }
  return { ranges, spaces: text.slice(0, end).match(SPACES)?.length ?? 0, text };
}

/** Text overflow in pre-transform CSS pixels; deliberately excludes markers. */
export function lineSpill(line: HTMLElement): number {
  const box = line.getBoundingClientRect();
  if (box.width === 0 || line.offsetWidth === 0) return 0;
  const rects = textRanges(line).ranges.flatMap((range) => Array.from(range.getClientRects())).filter((r) => r.width > 0);
  if (rects.length === 0) return 0;
  const inkLeft = Math.min(...rects.map((r) => r.left));
  const inkRight = Math.max(...rects.map((r) => r.right));
  const rtl = (line.dir || line.ownerDocument.defaultView!.getComputedStyle(line).direction) === 'rtl';
  // A hanging first-line indent can intentionally cross the leading edge.
  // This workaround repairs overflow at the paragraph's ending edge only.
  const leadingSpill = rtl ? inkRight - box.right : box.left - inkLeft;
  if (leadingSpill > MIN_SPILL_PX) return 0;
  return Math.max(rtl ? box.left - inkLeft : inkRight - box.right, 0) * line.offsetWidth / box.width;
}

function resetOwnedSpacing(line: HTMLElement): void {
  const owned = line.dataset[OURS];
  if (owned === undefined) return;
  if (owned === line.style.wordSpacing) line.style.wordSpacing = '';
  delete line.dataset[OURS];
  ownedStyles.set(line, line.getAttribute('style'));
}

/** DOM eligibility is deliberately conservative. A fragment continued on the
 * next page/column does not contain the paragraph's final line. Ordinary tabs
 * have fixed positions; changing preceding spaces would move text relative to
 * their stops. A list's synthetic suffix tab is safe to leave untouched.
 */
export function isFinalTextLine(line: HTMLElement): boolean {
  const fragment = line.closest<HTMLElement>(FRAGMENT_SELECTOR);
  if (!fragment || fragment.hasAttribute('data-continues-on-next')) return false;
  const lines = fragment.querySelectorAll(LINE_SELECTOR);
  if (lines.item(lines.length - 1) !== line) return false;
  if (line.querySelector('.superdoc-tab:not(.superdoc-marker-suffix-tab)')) return false;
  const allowed = `${TEXT_RUN_SELECTOR}, .superdoc-list-marker, .superdoc-marker-suffix-tab, .superdoc-marker-suffix-space`;
  if (Array.from(line.children).some((child) => !child.matches(allowed))) return false;
  const runs = Array.from(line.querySelectorAll<HTMLElement>(TEXT_RUN_SELECTOR));
  return runs.length > 0 && runs.every((run) => !run.style.wordSpacing);
}

/** `justified` must come from the document, never from painted text-align. */
export function compressSpillingLine(line: HTMLElement, justified = false): boolean {
  resetOwnedSpacing(line);
  if (!justified || !isFinalTextLine(line) || line.style.wordSpacing !== '') return false;
  const spill = lineSpill(line);
  const { spaces } = textRanges(line);
  if (spill <= MIN_SPILL_PX || spaces === 0) return false;
  // Leave unfamiliar, extreme overflows alone rather than overlapping words.
  // The measured cases need at most a quarter of the original space width.
  const scale = line.offsetWidth / line.getBoundingClientRect().width;
  let smallestSpace = Infinity;
  for (const run of line.querySelectorAll(TEXT_RUN_SELECTOR)) {
    const walker = line.ownerDocument.createTreeWalker(run, 4);
    let node: Node | null;
    while ((node = walker.nextNode())) {
      const text = node as Text;
      for (let i = 0; i < text.length; i++) if (text.data[i] === ' ' || text.data[i] === '\u00a0') {
        const range = line.ownerDocument.createRange();
        range.setStart(text, i); range.setEnd(text, i + 1);
        smallestSpace = Math.min(smallestSpace, range.getBoundingClientRect().width * scale);
      }
    }
  }
  if (!Number.isFinite(smallestSpace) || spill / spaces > smallestSpace / 2) return false;
  line.style.wordSpacing = `${-spill / spaces}px`;
  line.dataset[OURS] = line.style.wordSpacing;
  ownedStyles.set(line, line.getAttribute('style'));
  return true;
}

export interface JustifyLastLineHandle { dispose(): void; }

/** Read layout after paint; recheck identity/revision after an async API read.
 * Own style writes are recognized by their full serialized value, so a width
 * or alignment change on the same node is not mistaken for our mutation.
 */
export function installJustifyLastLine(root: HTMLElement, doc: JustificationDocument): JustifyLastLineHandle {
  const view = root.ownerDocument.defaultView ?? window;
  const pending = new Set<HTMLElement>();
  const compressed = new Set<HTMLElement>();
  const candidates = new Set<HTMLElement>();
  const generations = new WeakMap<HTMLElement, number>();
  const reader = createJustificationReader(doc);
  let frame = 0;
  let disposed = false;

  const enqueue = (line: HTMLElement) => {
    generations.set(line, (generations.get(line) ?? 0) + 1);
    pending.add(line);
  };
  const collect = (node: Node) => {
    const el = node instanceof view.HTMLElement ? node : node.parentElement;
    if (!el) return;
    const fragment = el.closest<HTMLElement>(FRAGMENT_SELECTOR);
    if (fragment) fragment.querySelectorAll<HTMLElement>(LINE_SELECTOR).forEach(enqueue);
    else el.querySelectorAll<HTMLElement>(LINE_SELECTOR).forEach(enqueue);
  };
  const schedule = () => {
    if (!disposed && pending.size > 0 && frame === 0) frame = view.requestAnimationFrame(flush);
  };
  const process = async (line: HTMLElement, generation: number) => {
    resetOwnedSpacing(line);
    compressed.delete(line);
    if (!line.isConnected || !isFinalTextLine(line) || line.style.wordSpacing || lineSpill(line) <= MIN_SPILL_PX) { candidates.delete(line); return; }
    candidates.add(line);
    const fragment = line.closest<HTMLElement>(FRAGMENT_SELECTOR)!;
    // Body only. Non-body and conditional table-style cascades need their own
    // story/context-aware proof before they can be enabled here.
    if (fragment.getAttribute('data-layout-story') !== 'body') return;
    const nodeId = fragment.getAttribute('data-source-node-id');
    if (!nodeId) return;
    const inTable = !!line.closest('.superdoc-table, .superdoc-table-cell, [data-table-id], table, td');
    const approved = await reader.isJustified(nodeId, inTable);
    if (disposed || generations.get(line) !== generation || !line.isConnected || !reader.isCurrent(approved)) return;
    if (compressSpillingLine(line, approved.justified)) compressed.add(line);
  };
  function flush() {
    frame = 0;
    const lines = [...pending];
    pending.clear();
    for (const line of lines) void process(line, generations.get(line)!);
  }

  const observer = new view.MutationObserver((records) => {
    for (const record of records) {
      const target = record.target;
      if (record.type === 'attributes' && record.attributeName === 'style' && target instanceof view.HTMLElement
        && ownedStyles.has(target) && ownedStyles.get(target) === target.getAttribute('style')) continue;
      // A selection overlay added next to pages is not a document change.
      // Only collect the target for edits inside a paragraph; page/root style
      // changes can affect descendant line widths and are handled separately.
      const element = target instanceof view.HTMLElement ? target : target.parentElement;
      const inFragment = !!element?.closest(FRAGMENT_SELECTOR);
      if (inFragment || (record.type === 'attributes' && record.attributeName !== 'class')) collect(target);
      record.addedNodes.forEach(collect);
    }
    for (const line of compressed) if (!line.isConnected) compressed.delete(line);
    for (const line of candidates) if (!line.isConnected) candidates.delete(line);
    schedule();
  });
  observer.observe(root, { childList: true, characterData: true, attributes: true,
    attributeFilter: ['style', 'class', 'dir', 'data-continues-on-next'], subtree: true });
  const stopReads = observeDocumentReadInvalidation(doc, () => {
    for (const line of candidates) enqueue(line);
    schedule();
  });
  const resize = typeof view.ResizeObserver === 'function' ? new view.ResizeObserver((entries) => {
    for (const entry of entries) collect(entry.target);
    schedule();
  }) : null;
  resize?.observe(root);
  const fontsChanged = () => { collect(root); schedule(); };
  root.ownerDocument.fonts?.addEventListener('loadingdone', fontsChanged);
  collect(root);
  schedule();

  return { dispose() {
    disposed = true;
    observer.disconnect();
    resize?.disconnect();
    stopReads();
    root.ownerDocument.fonts?.removeEventListener('loadingdone', fontsChanged);
    if (frame !== 0) view.cancelAnimationFrame(frame);
    pending.clear();
    for (const line of compressed) resetOwnedSpacing(line);
    compressed.clear();
    candidates.clear();
  } };
}
