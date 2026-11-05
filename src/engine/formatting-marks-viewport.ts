import { paintedParagraphEnds } from './painted-lines';
import {
  measureAllPageTextRuns, measurePageTextRuns, visiblePageElements,
  SETTLE_DELAYS_MS,
  type PageRectWatch, type PageTextRun, type ViewportSource,
} from './page-ruler';
import type { FormattingMarksEnd } from './formatting-marks-layer';

export type FormattingMarksWindow =
  | { mode: 'mapped'; endings: readonly FormattingMarksEnd[] }
  | { mode: 'legacy'; runs: readonly PageTextRun[] };

/** Do not measure text outside the window, or join unrelated paragraph text. */
export function measureFormattingMarksWindow(host: HTMLElement, reference: HTMLElement): FormattingMarksWindow {
  const endings: FormattingMarksEnd[] = [];
  for (const page of visiblePageElements(host)) {
    const mapped = paintedParagraphEnds(page);
    if (mapped === null) return { mode: 'legacy', runs: measureAllPageTextRuns(host, reference) };
    for (const end of mapped) {
      endings.push({
        nodeId: end.blockId, continuesFromPrevious: end.continuesFromPrevious,
        runs: measurePageTextRuns(end.element, reference),
      });
    }
  }
  return { mode: 'mapped', endings };
}

/** One scheduled frame and one set of settling timers, even during rapid edits. */
export function watchFormattingMarksWindow(options: {
  host: HTMLElement;
  reference: HTMLElement;
  ui?: ViewportSource | null;
  onChange: (window: FormattingMarksWindow) => void;
}): PageRectWatch {
  const { host, reference, ui, onChange } = options;
  let disposed = false;
  let frame: number | null = null;
  const timers = new Set<ReturnType<typeof setTimeout>>();
  function measureNow(): void {
    if (!disposed) onChange(measureFormattingMarksWindow(host, reference));
  }
  function schedule(): void {
    if (disposed || frame !== null) return;
    if (typeof requestAnimationFrame !== 'function') { measureNow(); return; }
    frame = requestAnimationFrame(() => { frame = null; measureNow(); });
  }
  host.addEventListener('scroll', schedule, { passive: true });
  const resize = typeof ResizeObserver === 'function' ? new ResizeObserver(schedule) : null;
  resize?.observe(host);
  resize?.observe(reference);
  let unobserve: (() => void) | null = null;
  try { unobserve = ui?.viewport?.observe?.(schedule) ?? null; } catch { /* no viewport subscription */ }
  measureNow();
  return {
    measure() {
      schedule();
      for (const timer of timers) clearTimeout(timer);
      timers.clear();
      for (const delay of SETTLE_DELAYS_MS) {
        const timer = setTimeout(() => { timers.delete(timer); schedule(); }, delay);
        timers.add(timer);
      }
    },
    dispose() {
      disposed = true;
      host.removeEventListener('scroll', schedule);
      resize?.disconnect();
      try { unobserve?.(); } catch { /* disposal must finish */ }
      for (const timer of timers) clearTimeout(timer);
      timers.clear();
      if (frame !== null && typeof cancelAnimationFrame === 'function') cancelAnimationFrame(frame);
    },
  };
}
