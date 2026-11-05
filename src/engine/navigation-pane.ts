/**
 * „חלונית ניווט” — רשימת הכותרות שבצד המסמך, כמו לשונית „כותרות” בחלונית
 * הניווט של Word. ui/shell/NavigationPane.vue מצייר; כאן הקריאה, העץ
 * והקפיצה.
 *
 * ## מה נחשב כותרת
 *
 * אותה הכרעה כמו בייצוא לספר אוצריא (`otzariaHeadingLevel` ב-otzaria-book.ts):
 * `headingLevel` של המנוע קודם — הוא מגיע מ-outline level אמיתי — ו-`styleId`
 * (`Heading1`…`Heading9`, `Title`, ושמות עבריים) הוא הגיבוי. ההבדל היחיד הוא
 * התקרה: אוצריא מכירה `<h1>`–`<h6>` ולכן חותכת שם, ו-Word מציג בחלונית תשע
 * רמות. כותרת ריקה אינה נכנסת — שורה ריקה ברשימה אינה יעד שאפשר לזהות.
 *
 * ## הקפיצה
 *
 * אותם שני API-ים ציבוריים שהחיפוש משתמש בהם כדי להגיע למופע (search.ts):
 * `ui.selection.apply` מציב את הסמן בתחילת הכותרת, ו-`ui.viewport.scrollIntoView`
 * גולל אליה. סמן ולא טווח: טווח שנכתב מבחוץ אינו מצויר, והסמן הוא גם מה
 * ש-Word עושה — לחיצה בחלונית מעבירה את נקודת ההקלדה לכותרת.
 */
import type { SuperDoc } from 'superdoc';
import type { MaybePromise } from './document-api';
import { getDocumentReadRevision, type DocumentReadRevision } from './document-read-state';

/** כותרת אחת ברשימה. */
export interface NavHeading {
  blockId: string;
  text: string;
  /** 1–9. */
  level: number;
  /** המיקום של הבלוק במסמך, בין **כל** הבלוקים — לא בין הכותרות. */
  ordinal: number;
}

interface NavBlockEntry {
  nodeId?: unknown;
  text?: unknown;
  styleId?: unknown;
  headingLevel?: unknown;
  isEmpty?: unknown;
}

interface NavSelectionInfo {
  target?: { segments?: readonly { blockId?: string }[] } | null;
}

export interface NavigationDocumentApi {
  blocks?: {
    list?: (input: {
      includeText: boolean;
      offset: number;
      limit: number;
      nodeIds?: string[];
    }) => MaybePromise<{ blocks?: readonly NavBlockEntry[]; revision?: string } | undefined>;
  } | null;
  selection?: {
    current?: () => MaybePromise<NavSelectionInfo | undefined>;
  } | null;
}

export interface NavigationUi {
  selection?: {
    apply?: (target: unknown) => unknown;
  } | null;
  viewport?: {
    scrollIntoView?: (input: {
      target: { kind: 'text'; blockId: string; range: { start: number; end: number } };
      block?: 'start' | 'center' | 'end' | 'nearest';
      behavior?: 'auto' | 'instant' | 'smooth';
    }) => MaybePromise<unknown>;
  } | null;
}

export interface NavigationHost {
  activeEditor?: { doc?: NavigationDocumentApi | null } | null;
  ui?: NavigationUi | null;
}

/** מה שהחלונית מחזיקה: הכותרות, ומיקום כל בלוק — כדי לדעת באיזה פרק הסמן. */
export interface NavOutline {
  headings: NavHeading[];
  ordinals: ReadonlyMap<string, number>;
  /** More blocks exist beyond the bounded scan. */
  partial: boolean;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type NavigationTarget = SuperDoc | NavigationHost | null | undefined;

/** אותם גבולות דפדוף כמו ב-otzaria-book.ts וב-search.ts. */
const BLOCKS_PAGE_SIZE = 500;
const BLOCKS_MAX_PAGES = 50;

/** תקרת Word: „רמה 9”. */
const MAX_LEVEL = 9;

const outlines = new WeakMap<NavigationDocumentApi, { revision: DocumentReadRevision; outline: NavOutline }>();

/** Also used by the component's explicit revision prop, including hosts without update events. */
export function invalidateNavOutline(target: NavigationTarget): void {
  const doc = hostOf(target)?.activeEditor?.doc;
  if (doc) outlines.delete(doc);
}

function hostOf(target: NavigationTarget): NavigationHost | null {
  return (target as NavigationHost | null | undefined) ?? null;
}

/**
 * רמת הכותרת של בלוק, 1–9, או `null` לפסקת גוף. הסדר והנרמול זהים ל-
 * `otzariaHeadingLevel` — ראו שם — בלי תקרת ה-6 של אוצריא.
 */
export function navHeadingLevel(
  styleId: string | null | undefined,
  headingLevel: number | undefined,
): number | null {
  if (typeof headingLevel === 'number' && Number.isInteger(headingLevel) && headingLevel >= 1) {
    return Math.min(headingLevel, MAX_LEVEL);
  }

  const key =
    typeof styleId === 'string' ? styleId.trim().toLowerCase().replace(/[\s_-]+/g, '') : '';
  if (key === 'title') return 1;

  const match = /^(?:heading|כותרת)([1-9])$/.exec(key);
  return match ? Number(match[1]) : null;
}

/** שורה אחת לתצוגה: שבירות רכות ורווחים כפולים מתכווצים לרווח אחד. */
function headingText(raw: string): string {
  return raw.replace(/\s+/g, ' ').trim();
}

/**
 * כל הכותרות במסמך, לפי הסדר. `null` כשאין Document API או שהקריאה נכשלה —
 * החלונית מראה אז את מה שהיה בה, ולא רשימה ריקה שאומרת „אין כותרות”.
 *
 * The scan is bounded; a one-block lookahead marks incomplete outlines explicitly.
 * The caller can cancel between pages after an edit, document switch or unmount.
 */
export async function readNavOutline(
  target: NavigationTarget,
  isCurrent: () => boolean = () => true,
): Promise<NavOutline | null> {
  const doc = hostOf(target)?.activeEditor?.doc;
  const blocks = doc?.blocks;
  const list = blocks?.list;
  if (typeof list !== 'function') return null;

  const revision = getDocumentReadRevision(doc!);
  const current = () => isCurrent() && getDocumentReadRevision(doc!) === revision;
  if (!current()) return null;
  const cached = outlines.get(doc!);
  if (cached?.revision === revision) return cached.outline;

  const headings: NavHeading[] = [];
  const ordinals = new Map<string, number>();
  const candidates: Array<Omit<NavHeading, 'text'>> = [];
  let offset = 0;
  let complete = false;
  let engineRevision: string | undefined;
  const sameRevision = (value: string | undefined): boolean => {
    if (engineRevision === undefined) engineRevision = value;
    return value === undefined || value === engineRevision;
  };
  try {
    for (let page = 0; page < BLOCKS_MAX_PAGES; page += 1) {
      const result = await list.call(blocks, { includeText: false, offset, limit: BLOCKS_PAGE_SIZE });
      if (!current() || !Array.isArray(result?.blocks) || !sameRevision(result?.revision)) return null;
      const entries = result.blocks;
      entries.forEach((entry, index) => {
        if (typeof entry?.nodeId !== 'string') return;
        ordinals.set(entry.nodeId, offset + index);
        const level = navHeadingLevel(
          typeof entry.styleId === 'string' ? entry.styleId : null,
          typeof entry.headingLevel === 'number' ? entry.headingLevel : undefined,
        );
        if (level === null) return;
        candidates.push({ blockId: entry.nodeId, level, ordinal: offset + index });
      });
      if (entries.length < BLOCKS_PAGE_SIZE) {
        complete = true;
        break;
      }
      offset += entries.length;
    }
    if (!complete) {
      const extra = await list.call(blocks, { includeText: false, offset, limit: 1 });
      if (!current() || !Array.isArray(extra?.blocks) || !sameRevision(extra?.revision)) return null;
      complete = extra.blocks.length === 0;
    }
    // Identity filtering happens before pagination in the public API. Keep the
    // ordinals from the metadata pass: filtered results have a different order.
    for (let start = 0; start < candidates.length; start += BLOCKS_PAGE_SIZE) {
      if (!current()) return null;
      const batch = candidates.slice(start, start + BLOCKS_PAGE_SIZE);
      const result = await list.call(blocks, {
        includeText: true, nodeIds: batch.map((entry) => entry.blockId), offset: 0, limit: batch.length,
      });
      if (!current() || !Array.isArray(result?.blocks) || !sameRevision(result?.revision)) return null;
      const texts = new Map(result.blocks.map((entry) => [entry.nodeId, entry]));
      for (const entry of batch) {
        // Missing text is an incomplete read, not proof that a heading is empty.
        const full = texts.get(entry.blockId);
        if (!full) return null;
        const raw = full.text;
        if (raw == null && full.isEmpty === true) continue;
        if (typeof raw !== 'string') return null;
        const text = headingText(raw);
        if (text) headings.push({ ...entry, text });
      }
    }
  } catch {
    return null;
  }
  const outline = { headings, ordinals, partial: !complete };
  outlines.set(doc!, { revision, outline });
  return outline;
}

/**
 * ההורה של כל כותרת ברשימה: האינדקס של הכותרת הקרובה שלפניה שרמתה נמוכה
 * ממנה, או `-1`. „נמוכה” ולא „בדיוק אחת פחות” — מסמך שקופץ מ„כותרת 1”
 * ל„כותרת 3” עדיין מקנן את השלישית מתחת לראשונה, כמו ב-Word.
 */
export function navParents(headings: readonly Pick<NavHeading, 'level'>[]): number[] {
  const parents: number[] = [];
  const stack: number[] = [];
  headings.forEach((heading, index) => {
    while (stack.length > 0 && headings[stack[stack.length - 1]].level >= heading.level) stack.pop();
    parents.push(stack.length > 0 ? stack[stack.length - 1] : -1);
    stack.push(index);
  });
  return parents;
}

/** האם לכותרת יש ילדים — כלומר האם יש לה חץ קיפול. */
export function navHasChildren(parents: readonly number[]): boolean[] {
  const has = parents.map(() => false);
  for (const parent of parents) if (parent >= 0) has[parent] = true;
  return has;
}

/** האם הכותרת מוסתרת מפני שאחד מאבותיה מקופל. */
export function navHidden(parents: readonly number[], collapsed: ReadonlySet<string>, ids: readonly string[]): boolean[] {
  // ההורה תמיד לפני הילד, ולכן `hidden[parent]` כבר חושב.
  const hidden: boolean[] = [];
  for (const parent of parents) {
    hidden.push(parent >= 0 && (hidden[parent] || collapsed.has(ids[parent])));
  }
  return hidden;
}

/**
 * הכותרת שהבלוק נמצא תחתיה: האחרונה שה-`ordinal` שלה אינו אחרי הבלוק.
 * `-1` כשהבלוק לפני הכותרת הראשונה. הרשימה ממוינת לפי `ordinal`, ולכן חיפוש
 * בינארי.
 */
export function navSectionIndex(headings: readonly Pick<NavHeading, 'ordinal'>[], ordinal: number): number {
  let low = 0;
  let high = headings.length - 1;
  let found = -1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    if (headings[mid].ordinal <= ordinal) {
      found = mid;
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }
  return found;
}

/** הבלוק שהסמן נמצא בו, או `null`. */
export async function readCaretBlockId(target: NavigationTarget): Promise<string | null> {
  const current = hostOf(target)?.activeEditor?.doc?.selection?.current;
  if (typeof current !== 'function') return null;
  try {
    const info = await current();
    const blockId = info?.target?.segments?.[0]?.blockId;
    return typeof blockId === 'string' ? blockId : null;
  } catch {
    return null;
  }
}

/**
 * מעבירה את הסמן לתחילת הכותרת וגוללת אליה. `false` כשאין למנוע את ה-API.
 * כשל בגלילה אינו כשל: הסמן כבר זז, וזה החלק שהמשתמש ממשיך ממנו.
 */
export async function goToNavHeading(target: NavigationTarget, blockId: string): Promise<boolean> {
  const ui = hostOf(target)?.ui;
  const apply = ui?.selection?.apply;
  if (typeof apply !== 'function') return false;
  try {
    apply.call(ui?.selection, {
      kind: 'selection',
      start: { kind: 'text', blockId, offset: 0 },
      end: { kind: 'text', blockId, offset: 0 },
    });
  } catch {
    return false;
  }
  try {
    await ui?.viewport?.scrollIntoView?.({
      target: { kind: 'text', blockId, range: { start: 0, end: 0 } },
      block: 'start',
      behavior: 'auto',
    });
  } catch {
    /* ויזואלי בלבד */
  }
  return true;
}
