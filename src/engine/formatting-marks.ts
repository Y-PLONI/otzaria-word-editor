/**
 * קריאת בלוקי המסמך (`doc.blocks.list`) עבור שכבת „סימני עיצוב" (¶) —
 * ui/shell/PilcrowOverlay.vue. אותה תבנית בדיוק כמו `createLineNumberingModel`/
 * `createPageBorderModel` (engine/page-setup.ts): מונה דורות שזורק תשובה של
 * מסמך שכבר נסגר, דיווח רק על שינוי אמיתי, `setEnabled` שנמנע מקריאה כשאיש
 * אינו רואה את התוצאה. הקריאה מוגבלת לפסקאות שסופן נמצא בחלון הציור;
 * הכיבוי מנתק את הצורך בקריאה).
 *
 * הקובץ הזה נפרד מ-`engine/formatting-marks-layer.ts` באותה חלוקה בדיוק כמו
 * page-setup.ts מול line-number-layer.ts: כאן א-סינכרוני ותלוי-Document API,
 * שם טהור וסינכרוני.
 */
import type { MaybePromise } from './document-api';
import type { FormattingMarksBlock } from './formatting-marks-layer';
import { getDocumentReadRevision, type DocumentReadRevision } from './document-read-state';

/** צורת כניסה אחת מ-`doc.blocks.list({includeText:true})` — רק מה שנדרש כאן. */
interface FormattingMarksBlockEntry {
  nodeId: string;
  nodeType: string;
  text?: string | null;
  isEmpty: boolean;
}

/** מה שנדרש מ-SuperDoc: רק ה-Document API של הבלוקים. */
export interface FormattingMarksHost {
  activeEditor?: {
    doc?: {
      blocks?: {
        list?: (input?: {
          includeText?: boolean;
          nodeIds?: string[];
          nodeTypes?: Array<'paragraph'>;
          offset?: number;
          limit?: number;
        }) => MaybePromise<{ blocks?: readonly FormattingMarksBlockEntry[] } | undefined>;
      } | null;
    } | null;
  } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל בבדיקות — אותה תבנית כמו `PageSetupTarget`. */
export type FormattingMarksTarget = FormattingMarksHost | null | undefined;

type MarksDocument = NonNullable<NonNullable<FormattingMarksHost['activeEditor']>['doc']>;
const cachedBlocks = new WeakMap<MarksDocument, {
  revision: DocumentReadRevision;
  entries: Map<string, FormattingMarksBlock>;
}>();
const TEXT_BATCH_SIZE = 500;
const TEXT_CACHE_LIMIT = 512;

/** טקסט הפסקאות שבחלון, במטמון חסום לפי מסמך/גרסה; null מפעיל fallback מלא. */
export async function readFormattingMarksBlocks(
  host: FormattingMarksTarget,
  /** null keeps the legacy full-sequence fallback when no painted identity is available. */
  nodeIds: readonly string[] | null = null,
  isCurrent: () => boolean = () => true,
): Promise<readonly FormattingMarksBlock[] | null> {
  const doc = host?.activeEditor?.doc;
  const blocks = doc?.blocks;
  const list = blocks?.list;
  if (typeof list !== 'function' || !isCurrent()) return null;
  if (nodeIds?.length === 0) return [];
  const revision = getDocumentReadRevision(doc!);
  let cache = cachedBlocks.get(doc!);
  if (cache?.revision !== revision) {
    cache = { revision, entries: new Map() };
    cachedBlocks.set(doc!, cache);
  }
  const ids = nodeIds === null ? null : [...new Set(nodeIds)];
  const missing = ids?.filter((id) => !cache.entries.has(id));
  const fresh: FormattingMarksBlock[] = [];
  try {
    for (let offset = 0; offset < (missing?.length ?? 1); offset += TEXT_BATCH_SIZE) {
      if (!isCurrent()) return null;
      const batch = missing?.slice(offset, offset + TEXT_BATCH_SIZE);
      const result = await list.call(blocks, {
        includeText: true, nodeTypes: ['paragraph'],
        ...(batch ? { nodeIds: batch, offset: 0, limit: batch.length } : {}),
      });
      if (!isCurrent() || getDocumentReadRevision(doc!) !== revision || !Array.isArray(result?.blocks)) return null;
      for (const entry of result.blocks) {
        // Never let a response to a filtered request leak unrelated paragraphs.
        if (batch && !batch.includes(entry.nodeId)) continue;
        fresh.push({ nodeId: entry.nodeId, nodeType: entry.nodeType, text: entry.text ?? '', isEmpty: entry.isEmpty });
      }
    }
  } catch {
    return null;
  }
  if (ids === null) return fresh;
  for (const entry of fresh) cache.entries.set(entry.nodeId, entry);
  const selected = ids.flatMap((id) => {
    const entry = cache.entries.get(id);
    return entry ? [entry] : [];
  });
  // Retain at most a small reading window, even after scrolling a whole book.
  while (cache.entries.size > TEXT_CACHE_LIMIT) cache.entries.delete(cache.entries.keys().next().value!);
  return selected;
}

/** השקטה בין קריאה לקריאה — אותו ערך כמו „מספרי שורות", מאותה סיבה. */
export const FORMATTING_MARKS_DEBOUNCE_MS = 300;

export interface FormattingMarksModelSource {
  read: (nodeIds: readonly string[] | null, isCurrent: () => boolean) => Promise<readonly FormattingMarksBlock[] | null>;
  onChange: (blocks: readonly FormattingMarksBlock[] | null) => void;
}

export interface FormattingMarksModel {
  getState(): readonly FormattingMarksBlock[] | null;
  /** Visible, identity-mapped paragraph endings; null requests the legacy fallback. */
  setBlockIds(nodeIds: readonly string[] | null): void;
  /**
   * סימני עיצוב מוצגים או מוסתרים. כבוי אינו קורא כלום; הדלקה קוראת את
   * החלון הנוכחי מיד, בלי לחכות לעריכה הראשונה.
   */
  setEnabled(enabled: boolean): void;
  /** קריאה מיידית, בלי השהיה. */
  refreshNow(): void;
  /** קריאה מושהית — אחרי שינוי כלשהו במסמך (עריכה). */
  noteDocumentChanged(): void;
  dispose(): void;
}

export function createFormattingMarksModel(source: FormattingMarksModelSource): FormattingMarksModel {
  let state: readonly FormattingMarksBlock[] | null = null;
  let enabled = false;
  let disposed = false;
  let generation = 0;
  let nodeIds: readonly string[] | null = [];
  let reading = false;
  let queued = false;
  let timer: ReturnType<typeof setTimeout> | undefined;

  async function read(): Promise<void> {
    if (disposed || !enabled || !queued || reading) return;
    queued = false;
    reading = true;
    const mine = generation;
    try {
      const next = nodeIds?.length === 0 ? [] : await source.read(nodeIds, () => !disposed && enabled && mine === generation);
      if (disposed || !enabled || mine !== generation) return;
      state = next;
      source.onChange(next);
    } catch {
      if (!disposed && enabled && mine === generation) {
        state = null;
        source.onChange(null);
      }
    } finally {
      reading = false;
      if (queued) void read();
    }
  }

  function refreshNow(): void {
    if (disposed || !enabled) return;
    clearTimeout(timer);
    generation += 1;
    queued = true;
    void read();
  }

  return {
    getState: () => state,

    setBlockIds(next) {
      if (disposed || (nodeIds === null && next === null)) return;
      if (nodeIds && next && nodeIds.length === next.length && nodeIds.every((id, i) => id === next[i])) return;
      nodeIds = next === null ? null : [...next];
      generation += 1;
      state = null;
      source.onChange(null);
      refreshNow();
    },

    setEnabled(next) {
      if (disposed || enabled === next) return;
      enabled = next;
      if (!enabled) {
        // הדור עולה כדי שקריאה שבאוויר לא תדווח אחרי הכיבוי.
        generation += 1;
        queued = false;
        clearTimeout(timer);
        state = null;
        source.onChange(null);
        return;
      }
      refreshNow();
    },

    refreshNow,

    noteDocumentChanged() {
      if (disposed || !enabled) return;
      generation += 1;
      queued = false;
      clearTimeout(timer);
      timer = setTimeout(refreshNow, FORMATTING_MARKS_DEBOUNCE_MS);
    },

    dispose() {
      disposed = true;
      generation += 1;
      queued = false;
      clearTimeout(timer);
    },
  };
}
