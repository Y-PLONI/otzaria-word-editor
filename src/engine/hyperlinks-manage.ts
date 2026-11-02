/**
 * ניהול היפר-קישורים (גל 22): רשימה, הסרה ועריכה — דרך
 * `hyperlinks.list/wrap/remove`. `hyperlinks.patch` אינה זמינה ב-2.8.0,
 * ולכן **עריכה = הסרה + עטיפה מחדש** — אותה תבנית של `captions.update`
 * בגל 8: שתי פעולות לא-אטומיות מעל תוכן המשתמש.
 *
 * ## מה שנמדד
 *
 * - **היעד ל-wrap הוא `TextAddress`** (`{kind:'text',blockId,range}`) —
 *   SelectionTarget נדחה ("requires a valid TextAddress target").
 * - **היעד ל-remove הוא `HyperlinkTarget`** — כתובת של **צומת הקישור עצמו**
 *   (`{kind:'inline', nodeType:'hyperlink', anchor:{start,end}}`), לא טווח
 *   טקסט. `remove({ within: TextAddress })` זורק `Cannot read properties of
 *   undefined (reading 'anchor')` — הכתובת הנכונה נלקחת מ-`list()`.
 * - **wrap דורש מפרט קישור:** `{ link: { destination: { href } } }` — href בלבד
 *   זרק "requires a link specification object".
 * - `list()` מחזיר `{ items: [{ address, properties, text }] }` (מעטפת
 *   discovery סטנדרטית) — לא `stories[].hyperlinks[]`. הסרה על כתובת שאינה
 *   קישור מחזירה `TARGET_NOT_FOUND`.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt } from './document-api';

const UNAVAILABLE_TEXT = 'אינו זמין בגרסה זו';

export interface HyperlinkInfo {
  id?: string;
  href?: string;
  anchor?: string;
}

/** מיקום בתוך המסמך — `blockId` + היסט טקסט שטוח (נמדד: `InlineAnchor.start/end`). */
interface HyperlinkPosition {
  blockId?: string;
  offset?: number;
}

/** כתובת צומת קישור, כפי שהיא חוזרת מ-`list()` ונשלחת חזרה ל-`remove`. */
export interface HyperlinkTargetAddress {
  kind: 'inline';
  nodeType: 'hyperlink';
  anchor: { start: HyperlinkPosition; end: HyperlinkPosition };
  story?: unknown;
}

/** פריט גולמי בודד מתוך `items[]` של `hyperlinks.list()`. */
interface RawHyperlinkItem {
  id?: string;
  address?: HyperlinkTargetAddress;
  properties?: { href?: string; anchor?: string };
  text?: string;
}

interface HyperlinksApiShape {
  selection?: {
    current?: () => MaybePromise<SelectionInfoLike | undefined>;
  };
  clipboard?: {
    insert?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
  };
  hyperlinks?: {
    list?: (input?: Record<string, unknown>) => MaybePromise<unknown>;
    wrap?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
    remove?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
  };
}

export interface HyperlinksHost {
  activeEditor?: { doc?: HyperlinksApiShape | null } | null;
  ui?: {
    selection?: {
      /** מציבה סמן. אותה ידית ש-at-mention-overlay.ts מציב בה את הסמן אחרי קישור. */
      apply?: (target: Record<string, unknown>) => unknown;
    } | null;
  } | null;
}

export type HyperlinksTarget = SuperDoc | HyperlinksHost | null | undefined;

interface SelectionInfoLike {
  empty?: boolean;
  target?: {
    segments?: ReadonlyArray<{ blockId?: string; range?: { start?: number; end?: number } }>;
  } | null;
}

type MaybePromise<T> = T | Promise<T>;

function docOf(host: HyperlinksTarget): HyperlinksApiShape | null {
  return (host as HyperlinksHost | null | undefined)?.activeEditor?.doc ?? null;
}

/** TextAddress מהבחירה — הקטע הראשון עם blockId וטווח. `null` אם אין. */
async function readTextAddress(
  doc: HyperlinksApiShape,
): Promise<{ kind: 'text'; blockId: string; range: { start: number; end: number } } | null> {
  try {
    const info = await doc.selection?.current?.();
    const segment = info?.target?.segments?.find(
      (s) =>
        typeof s?.blockId === 'string' &&
        typeof s.range?.start === 'number' &&
        typeof s.range.end === 'number',
    );
    if (!segment) return null;
    return {
      kind: 'text',
      blockId: segment.blockId as string,
      range: { start: segment.range!.start as number, end: segment.range!.end as number },
    };
  } catch {
    return null;
  }
}

/** פריטי `items[]` הגולמיים של `hyperlinks.list()` (מעטפת discovery סטנדרטית). */
function rawHyperlinkItems(raw: unknown): RawHyperlinkItem[] {
  if (!raw || typeof raw !== 'object') return [];
  const items = (raw as { items?: unknown }).items;
  return Array.isArray(items) ? (items as RawHyperlinkItem[]) : [];
}

/** דולה את הקישורים ל-`HyperlinkInfo[]` חיצוני שטוח. */
function flattenLinks(raw: unknown): HyperlinkInfo[] {
  return rawHyperlinkItems(raw).map((item) => ({
    id: item.id,
    href: item.properties?.href,
    anchor: item.properties?.anchor,
  }));
}

/**
 * מאתרת את כתובת צומת הקישור שחופפת לטווח המסומן (`TextAddress`), כדי
 * לשלוח אותה ל-`remove`. `null` = אין קישור בטווח.
 */
async function findHyperlinkTarget(
  doc: HyperlinksApiShape,
  address: { blockId: string; range: { start: number; end: number } },
): Promise<HyperlinkTargetAddress | null> {
  const list = doc.hyperlinks?.list;
  if (typeof list !== 'function') return null;

  let raw: unknown;
  try {
    raw = await list();
  } catch {
    return null;
  }

  for (const item of rawHyperlinkItems(raw)) {
    const anchor = item.address?.anchor;
    const start = anchor?.start?.offset;
    const end = anchor?.end?.offset;
    if (
      typeof start !== 'number' ||
      typeof end !== 'number' ||
      anchor?.start?.blockId !== address.blockId ||
      anchor?.end?.blockId !== address.blockId
    ) {
      continue;
    }
    // חפיפה (כולל מגע בקצה) בין טווח הקישור לטווח המסומן.
    if (address.range.start <= end && address.range.end >= start) {
      return item.address ?? null;
    }
  }
  return null;
}

/** רשימת כל הקישורים במסמך. מחזירה `null` כשאין מנוע. */
export async function listHyperlinks(host: HyperlinksTarget): Promise<HyperlinkInfo[] | null> {
  const list = docOf(host)?.hyperlinks?.list;
  if (typeof list !== 'function') return null;
  try {
    return flattenLinks(await list());
  } catch {
    return null;
  }
}

/**
 * „הטווח נוגע בקישור קיים”, בכל הצורות שהמנוע מדווח בהן.
 *
 * הבדיקה היא על ההודעה ולא על הקוד בלבד, וזה נמדד: הקוד שחזר הוא
 * `INVALID_CONTEXT` (ומ-`wrap` — `INVALID_TARGET`), בעוד
 * `hyperlink-nested-unsupported` הוא **ההודעה**. בדיקה על הקוד לבדו לא
 * התאימה מעולם, ולכן המשתמש קיבל בשורת המצב את המחרוזת האנגלית הגולמית
 * במקום את ההודעה בעברית. משותפת לאזכור „@” ולדיאלוג הקישור.
 */
export function isHyperlinkOverlapFailure(receipt: DocReceipt | null | undefined): boolean {
  const failure = receipt?.failure;
  const text = `${failure?.code ?? ''} ${failure?.message ?? ''}`;
  return /hyperlink-nested-unsupported/i.test(text) || /overlap.*existing hyperlink/i.test(text);
}

/** סמן יחיד בבלוק טקסט, מתוך `selectionTarget`. `null` לכל צורה אחרת. */
function readCaretPoint(
  target: unknown,
): { blockId: string; offset: number; story: unknown } | null {
  const t = target as {
    kind?: string;
    story?: unknown;
    start?: { kind?: string; blockId?: unknown; offset?: unknown; story?: unknown };
    end?: { kind?: string; blockId?: unknown; offset?: unknown };
  } | null;
  if (!t || t.kind !== 'selection') return null;
  const { start, end } = t;
  if (start?.kind !== 'text' || end?.kind !== 'text') return null;
  if (typeof start.blockId !== 'string' || start.blockId !== end.blockId) return null;
  if (typeof start.offset !== 'number' || start.offset !== end.offset) return null;
  return { blockId: start.blockId, offset: start.offset, story: t.story ?? start.story ?? null };
}

/**
 * Insert a complete inline hyperlink in one public clipboard transaction.
 * `hyperlinks.insert` rejects even a distant second link in the same paragraph
 * on 2.23. `doc.insert` + `hyperlinks.wrap` works, but creates two Undo entries
 * and needs a compensating delete on failure. A v2.2 clipboard fragment carries
 * the hyperlink itself; match-destination preserves authored run properties.
 * It also works beside empty hyperlink remnants without a document-wide scan.
 * Rich inbound content rejects otzaria: URLs; keep those on the native link
 * command, so the supported custom-link workflow keeps its original behavior.
 * `null` means this is not a single caret (or the API is unavailable).
 */
export async function insertHyperlinkAtCaret(
  host: HyperlinksTarget,
  input: { href: string; text?: string; at: unknown },
): Promise<CommandOutcome | null> {
  const failedAction = 'הוספת הקישור נכשלה';
  const doc = docOf(host);
  const clipboard = doc?.clipboard;
  const insert = clipboard?.insert;
  const caret = readCaretPoint(input.at);
  if (typeof insert !== 'function' || !caret) return null;
  if (!/^(?:https?:|mailto:)/i.test(input.href)) return null;

  const text = input.text?.trim() || input.href;
  const { blockId, offset, story } = caret;
  const point = (at: number) => ({ kind: 'text', blockId, offset: at, ...(story ? { story } : {}) });
  const over = (at: number) => ({
    kind: 'selection',
    start: point(at),
    end: point(at),
    ...(story ? { story } : {}),
  });
  // v2.2 is the public clipboard schema that supports inline hyperlinks.
  // runs must equal the text projection of inlines, including hyperlink text.
  let receipt: DocReceipt;
  try {
    receipt = await insert.call(clipboard, {
      fragment: {
        kind: 'superdoc.clipboard.fragment',
        version: 'v2.2',
        blocks: [{
          kind: 'paragraph',
          runs: [{ text }],
          inlines: [{
            kind: 'hyperlink',
            target: { kind: 'external', url: input.href },
            inlines: [{ kind: 'text', text }],
          }],
        }],
      },
      formattingMode: 'match-destination',
      target: over(offset),
    });
  } catch (error) {
    return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
  }
  if (receipt?.success === false) {
    return {
      ok: false,
      message: isHyperlinkOverlapFailure(receipt)
        ? `${failedAction}: אי אפשר להוסיף קישור בתוך קישור קיים`
        : receiptFailureText(failedAction, receipt),
      reason: receipt.failure?.code,
    };
  }

  const selection = (host as HyperlinksHost | null | undefined)?.ui?.selection;
  try {
    // Engine text offsets use UTF-16 code units, like String.length.
    selection?.apply?.call(selection, over(offset + text.length));
  } catch (error) {
    console.warn('[otzaria-word] קישור: הצבת הסמן נכשלה', error);
  }
  return { ok: true };
}

/** „הסר קישור" על הטווח המסומן. TARGET_NOT_FOUND = כבר אין קישור = הצלחה. */
export async function removeHyperlink(host: HyperlinksTarget): Promise<CommandOutcome> {
  const failedAction = 'הסרת הקישור נכשלה';

  const doc = docOf(host);
  const remove = doc?.hyperlinks?.remove;
  if (!doc || typeof remove !== 'function') {
    return { ok: false, message: `${failedAction}: ${UNAVAILABLE_TEXT}`, reason: 'command-unsupported' };
  }

  const address = await readTextAddress(doc);
  if (!address) {
    return { ok: false, message: `${failedAction}: יש לסמן טקסט תחילה`, reason: 'selection-required' };
  }

  const target = await findHyperlinkTarget(doc, address);
  if (!target) {
    // אין קישור בטווח = המצב המבוקש כבר מתקיים. לא כשל.
    return { ok: true };
  }

  let receipt: DocReceipt;
  try {
    receipt = await remove({ target });
  } catch (error) {
    return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
  }

  // אין קישור בטווח = המצב המבוקש כבר מתקיים. לא כשל.
  if (receipt?.success === false && receipt.failure?.code === 'TARGET_NOT_FOUND') {
    return { ok: true };
  }
  if (receipt?.success === false) {
    return { ok: false, message: receiptFailureText(failedAction, receipt), reason: receipt.failure?.code };
  }

  return { ok: true };
}

/**
 * „עריכת קישור" — remove + wrap. אימות לפני מגע (`list`), ואם העטיפה
 * נכשלת אחרי הסרה — הודעה שאומרת שהקישור אבד והטקסט נשמר.
 */
export async function editHyperlink(host: HyperlinksTarget, newHref: string): Promise<CommandOutcome> {
  const failedAction = 'עריכת הקישור נכשלה';

  const href = typeof newHref === 'string' ? newHref.trim() : '';
  if (href === '' || href.length > 2048) {
    return { ok: false, message: `${failedAction}: כתובת הקישור נדרשת`, reason: 'invalid-href' };
  }

  const doc = docOf(host);
  const wrap = doc?.hyperlinks?.wrap;
  const remove = doc?.hyperlinks?.remove;
  if (!doc || typeof wrap !== 'function' || typeof remove !== 'function') {
    return { ok: false, message: `${failedAction}: ${UNAVAILABLE_TEXT}`, reason: 'command-unsupported' };
  }

  const address = await readTextAddress(doc);
  if (!address) {
    return { ok: false, message: `${failedAction}: יש לסמן טקסט תחילה`, reason: 'selection-required' };
  }

  // אימות לפני מגע: אין קישור בטווח → עטיפה ישירה, בלי remove מיותר.
  const existingTarget = await findHyperlinkTarget(doc, address);
  const hadLink = existingTarget !== null;

  if (hadLink) {
    let removed: DocReceipt;
    try {
      removed = await remove({ target: existingTarget });
    } catch (error) {
      return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
    }
    if (removed?.success === false && removed.failure?.code !== 'TARGET_NOT_FOUND') {
      return { ok: false, message: receiptFailureText(failedAction, removed), reason: removed.failure?.code };
    }
  }

  let wrappedReceipt: DocReceipt;
  try {
    wrappedReceipt = await wrap({ target: address, link: { destination: { href } } });
  } catch (error) {
    if (hadLink) {
      return {
        ok: false,
        message: `${failedAction}: הקישור הוסר אך החדש לא נוצר — הטקסט נשמר ללא קישור`,
        reason: 'link-lost',
      };
    }
    return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
  }

  if (wrappedReceipt?.success === false) {
    if (hadLink) {
      return {
        ok: false,
        message: `${failedAction}: הקישור הוסר אך החדש לא נוצר — הטקסט נשמר ללא קישור`,
        reason: wrappedReceipt.failure?.code,
      };
    }
    return { ok: false, message: receiptFailureText(failedAction, wrappedReceipt), reason: wrappedReceipt.failure?.code };
  }

  return { ok: true };
}

