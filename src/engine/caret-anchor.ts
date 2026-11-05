/**
 * „לחזור לאן שהייתי” — הסמן, הבחירה, והעמוד שרואים.
 *
 * ## מה הבעיה שהמודול הזה פותר
 *
 * מסמך שנפתח מחדש נפתח בתחילתו. זה נכון גם כשמדובר באותו קובץ בדיוק שהיה
 * פתוח לפני רגע, ולמשתמש שכתב בעמוד 40 זה אומר לגלול לשם בכל פעם מחדש.
 *
 * ## למה לא לשמור „מיקום” מספרי
 *
 * המיקום של ProseMirror (`from`/`to`) הוא אינדקס בתוך עץ שנבנה מחדש בכל
 * פתיחה, והוא גם אינו חלק מהמשטח הציבורי של המנוע. השמירה כאן היא במודל
 * הכתובות של ה-Document API: **פסקה + היסט תו** (`{blockId, offset}`), שהוא
 * המודל שהמנוע עצמו מייצא ומקבל — אותו מודל שבו נשמרות תגובות והערות.
 *
 * ## שני עוגנים, ולא אחד
 *
 * `blockId` הוא מזהה שהמנוע נותן לפסקה. השאלה אם הוא שורד סבב של
 * ייצוא-ופתיחה-מחדש היא שאלה על המנוע, לא עלינו, ואין לה תשובה בתיעוד. לכן
 * נשמרים **שניים**:
 *
 * 1. `blockId` — מנוסה ראשון, ועולה אפס: אם הוא שרד, השחזור מדויק ומיידי.
 * 2. `ordinal` — מקומה של הפסקה בסדר המסמך. הוא אינו יכול „לא לשרוד”: פסקה
 *    שלישית נשארת שלישית. הוא נפתר דרך `doc.blocks.list({offset, limit: 1})`,
 *    ולכן הוא הגיבוי ולא הראשון — קריאה לכל קצה ייחודי של הבחירה בלבד.
 *
 * הצורה הזאת גם אינה תלויה בתשובה: אם המזהים שורדים, המסלול השני לא ירוץ אף
 * פעם; אם לא — הוא זה שיעבוד, ואין שינוי קוד בין המקרים.
 *
 * ## למה `offset` מנוסה פעמיים
 *
 * הקובץ בדיסק יכול היה להשתנות בין ההפעלות (Word, עורך אחר, גיבוי). פסקה
 * שהתקצרה מקבלת היסט שאינו קיים בה, והמנוע נכשל סגור. במקום לוותר, הניסיון
 * השני הוא תחילת אותה פסקה: „העמוד הנכון” הוא כמעט כל מה שהמשתמש רצה, ובוודאי
 * יותר מתחילת המסמך.
 */
import type { BorrowedSuperDocUI, SuperDoc } from 'superdoc';
import type { SelectionTarget, TextAddress } from 'superdoc/ui';
import type { MaybePromise } from './document-api';
import { paintedFirstBlockId } from './painted-lines';

/** קצה אחד של הבחירה: פסקה, מקומה בסדר, והיסט התו בתוכה. */
export interface CaretPoint {
  blockId: string;
  /** `null` כשלא הצלחנו לפתור את הסדר — אז נשאר רק ה-`blockId`. */
  ordinal: number | null;
  offset: number;
}

/**
 * המקום שהמשתמש היה בו. `end` קיים רק כשהייתה בחירת טווח — סמן מכווץ הוא
 * `end: null`, ולא עותק של `start`, כדי שהשאלה „היה טווח?” תיקרא מהצורה.
 */
export interface CaretAnchor {
  start: CaretPoint;
  end: CaretPoint | null;
}

/** ה-slice של הבחירה, בחלק שנצרך כאן. הכול אופציונלי — נקרא בהגנה. */
interface SelectionSliceLike {
  selectionTarget?: unknown;
  target?: unknown;
}

/** מה שנצרך מ-`superdoc.ui`. גרסת מנוע בלי ההידיות האלה נופלת בחן. */
export interface CaretUiSource {
  selection?: {
    getSnapshot?: () => SelectionSliceLike | null | undefined;
    apply?: (target: SelectionTarget) => { ok?: unknown } | null | undefined;
  };
  viewport?: {
    scrollIntoView?: (input: {
      target: TextAddress;
      block?: 'start' | 'center' | 'end' | 'nearest';
      behavior?: 'auto' | 'instant' | 'smooth';
    }) => MaybePromise<{ success?: unknown } | null | undefined>;
  };
}

/**
 * ה-union מאפשר גם את ה-controller האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts.
 */
export type CaretUi = BorrowedSuperDocUI | CaretUiSource | null | undefined;

/** רשומת פסקה כפי ש-`blocks.list` מחזיר אותה, בחלק שנצרך כאן. */
interface BlockEntryLike {
  ordinal?: unknown;
  nodeId?: unknown;
}

interface BlocksListLike {
  total?: unknown;
  blocks?: unknown;
}

/** מה שנצרך מ-`superdoc.activeEditor.doc`. */
export interface CaretDocSource {
  activeEditor?: {
    doc?: {
      blocks?: {
        list?: (input?: {
          offset?: number;
          limit?: number;
        }) => MaybePromise<BlocksListLike | null | undefined>;
      } | null;
    } | null;
  } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. */
export type CaretDoc = SuperDoc | CaretDocSource | null | undefined;

/**
 * כמה פסקאות לבקש בכל קריאה בעת שמירת העוגן, וכמה קריאות לכל היותר.
 *
 * המנוע רשאי להחזיר פחות מהמבוקש, ולכן ההתקדמות היא לפי מה שחזר בפועל ולא
 * לפי הבקשה — לולאה שמתקדמת ב-`limit` הייתה מדלגת על פסקאות בכל מנוע שמגביל
 * את הדף. התקרה מגבילה את עבודת שמירת המיקום: 50 קריאות של 500 הן 25,000
 * פסקאות, ומעליהן נשמר רק המזהה. שחזור ordinal שכבר ידוע אינו סורק את הרשימה.
 */
const PAGE_SIZE = 500;
const MAX_PAGES = 50;

function asPoint(value: unknown): { blockId: string; offset: number } | null {
  if (!value || typeof value !== 'object') return null;
  const point = value as { kind?: unknown; blockId?: unknown; offset?: unknown };
  if (point.kind !== 'text') return null;
  if (typeof point.blockId !== 'string' || point.blockId === '') return null;
  const offset =
    typeof point.offset === 'number' && Number.isFinite(point.offset) ? point.offset : 0;
  return { blockId: point.blockId, offset: Math.max(0, Math.trunc(offset)) };
}

/**
 * הקצוות מתוך `selectionTarget` של המנוע.
 *
 * `nodeEdge` (הגבול של טבלה או תמונה) אינו נתמך כאן בכוונה: הוא אינו „מקום
 * בטקסט”, ואין לו היסט שאפשר לשחזר. במצב כזה לא נשמר עוגן, והמסמך נפתח
 * בתחילתו — בדיוק כפי שהיה לפני המודול הזה.
 */
function pointsFromSelectionTarget(
  value: unknown,
): { start: { blockId: string; offset: number }; end: { blockId: string; offset: number } } | null {
  if (!value || typeof value !== 'object') return null;
  const target = value as { kind?: unknown; start?: unknown; end?: unknown };
  if (target.kind !== 'selection') return null;

  const start = asPoint(target.start);
  if (!start) return null;
  return { start, end: asPoint(target.end) ?? start };
}

/**
 * גיבוי ל-`selectionTarget`: הקטע הראשון של `target`.
 *
 * `TextTarget` הוא מה שהמנוע מחזיר כשהבחירה נפתרת לטקסט מעוגן, והוא קיים גם
 * במצבים שבהם ה-`selectionTarget` המפורש אינו נפתר (המנוע מתעד אותו כמי
 * ש„עשוי להיות null כשהריצה אינה יכולה להקרין אותו באמת”). הוא נושא טווחים
 * ולא נקודות, ולכן הקצוות נגזרים מ-`range`.
 */
function pointsFromTextTarget(
  value: unknown,
): { start: { blockId: string; offset: number }; end: { blockId: string; offset: number } } | null {
  if (!value || typeof value !== 'object') return null;
  const target = value as { kind?: unknown; segments?: unknown };
  if (target.kind !== 'text' || !Array.isArray(target.segments) || target.segments.length === 0) {
    return null;
  }

  const first = target.segments[0] as { blockId?: unknown; range?: { start?: unknown } };
  const last = target.segments[target.segments.length - 1] as {
    blockId?: unknown;
    range?: { end?: unknown };
  };
  if (typeof first?.blockId !== 'string' || typeof last?.blockId !== 'string') return null;

  const startOffset = typeof first.range?.start === 'number' ? first.range.start : 0;
  const endOffset = typeof last.range?.end === 'number' ? last.range.end : startOffset;
  return {
    start: { blockId: first.blockId, offset: Math.max(0, Math.trunc(startOffset)) },
    end: { blockId: last.blockId, offset: Math.max(0, Math.trunc(endOffset)) },
  };
}

/** כל הפסקאות בסדר המסמך, כ-`nodeId` לפי `ordinal`. ראו PAGE_SIZE. */
async function listBlockIds(host: CaretDoc): Promise<string[] | null> {
  const list = (host as CaretDocSource | null | undefined)?.activeEditor?.doc?.blocks?.list;
  if (typeof list !== 'function') return null;

  const ids: string[] = [];
  try {
    for (let page = 0; page < MAX_PAGES; page += 1) {
      const result = await list({ offset: ids.length, limit: PAGE_SIZE });
      const blocks = Array.isArray(result?.blocks) ? (result.blocks as BlockEntryLike[]) : [];
      if (blocks.length === 0) break;

      for (const block of blocks) {
        // פסקה בלי `nodeId` עדיין תופסת מקום בסדר, ולכן היא נכנסת כמחרוזת
        // ריקה: דילוג עליה היה מזיז את כל מי שאחריה בעוגן אחד.
        ids.push(typeof block?.nodeId === 'string' ? block.nodeId : '');
      }

      const total = typeof result?.total === 'number' ? result.total : ids.length;
      if (ids.length >= total) break;
    }
  } catch (error) {
    console.warn('[otzaria-word] קריאת סדר הפסקאות נכשלה', error);
    return ids.length > 0 ? ids : null;
  }

  return ids;
}

/**
 * אותו סדר גוף מסמך לא מסונן שבו `listBlockIds` שומר ordinal, ברשומה אחת.
 * אם מנוע החזיר רשומה ממיקום אחר (למשל אחרי הגבלת offset), אין לשחזר אליה.
 */
async function blockIdAt(host: CaretDoc, ordinal: number): Promise<string | null> {
  if (!Number.isSafeInteger(ordinal) || ordinal < 0) return null;
  const blocks = (host as CaretDocSource | null | undefined)?.activeEditor?.doc?.blocks;
  const list = blocks?.list;
  if (typeof list !== 'function') return null;

  try {
    const result = await list.call(blocks, { offset: ordinal, limit: 1 });
    if (typeof result?.total === 'number' && ordinal >= result.total) return null;
    const entry = Array.isArray(result?.blocks) ? (result.blocks[0] as BlockEntryLike) : undefined;
    if (entry?.ordinal !== undefined && entry.ordinal !== ordinal) return null;
    return typeof entry?.nodeId === 'string' && entry.nodeId !== '' ? entry.nodeId : null;
  } catch (error) {
    console.warn('[otzaria-word] קריאת הפסקה לשחזור המקום נכשלה', error);
    return null;
  }
}

/**
 * המקום שבו הסמן נמצא עכשיו, או `null` כשאין מקום שאפשר לתאר — אין בחירה,
 * הבחירה אינה בטקסט, או שהמנוע אינו חושף את ה-handle.
 *
 * נקראת ברגע השמירה בלבד (השהיה או מעבר לרקע) ולא על כל הקלדה: היא פותרת את
 * סדר הפסקאות, וזו עבודה שאין שום סיבה לעשות שישים פעם בדקה.
 */
export async function readCaretAnchor(
  host: CaretUi,
  doc: CaretDoc,
  previous?: CaretAnchor | null,
): Promise<CaretAnchor | null> {
  const ui = host as CaretUiSource | null | undefined;
  const getSnapshot = ui?.selection?.getSnapshot;
  if (typeof getSnapshot !== 'function') return null;

  let slice: SelectionSliceLike | null | undefined;
  try {
    slice = getSnapshot.call(ui?.selection);
  } catch (error) {
    console.warn('[otzaria-word] קריאת הבחירה לשמירת המקום נכשלה', error);
    return null;
  }

  const points =
    pointsFromSelectionTarget(slice?.selectionTarget) ?? pointsFromTextTarget(slice?.target);
  if (!points) return null;

  const collapsed =
    points.end.blockId === points.start.blockId && points.end.offset === points.start.offset;

  /**
   * הסדר של הפסקאות שהעוגן הקודם כבר פתר.
   *
   * למה זה כאן ולא אופטימיזציה של הקורא: הקריאה הזאת נעשית בכל הפוגה בהקלדה,
   * והמקרה הרווח ביותר הוא סמן שזז **בתוך אותה פסקה** — אותו `blockId`, היסט
   * אחר. סריקת כל הפסקאות מחדש בשביל מספר שכבר ידוע היא העבודה היחידה
   * שהמודול הזה עושה שאפשר פשוט לא לעשות.
   */
  const known = (blockId: string): number | null | undefined => {
    if (previous?.start.blockId === blockId) return previous.start.ordinal;
    if (previous?.end?.blockId === blockId) return previous.end.ordinal;
    return undefined;
  };

  const startKnown = known(points.start.blockId);
  const endKnown = collapsed ? startKnown : known(points.end.blockId);
  const resolved = startKnown !== undefined && endKnown !== undefined;

  const ids = resolved ? null : await listBlockIds(doc);
  const ordinalOf = (blockId: string, cached: number | null | undefined): number | null => {
    if (cached !== undefined) return cached;
    if (!ids) return null;
    const index = ids.indexOf(blockId);
    return index >= 0 ? index : null;
  };

  const start: CaretPoint = {
    ...points.start,
    ordinal: ordinalOf(points.start.blockId, startKnown),
  };

  return {
    start,
    end: collapsed ? null : { ...points.end, ordinal: ordinalOf(points.end.blockId, endKnown) },
  };
}

function selectionTargetOf(start: CaretPoint, end: CaretPoint): SelectionTarget {
  return {
    kind: 'selection',
    start: { kind: 'text', blockId: start.blockId, offset: start.offset },
    end: { kind: 'text', blockId: end.blockId, offset: end.offset },
  };
}

function applied(ui: CaretUiSource, target: SelectionTarget): boolean {
  const apply = ui.selection?.apply;
  if (typeof apply !== 'function') return false;
  try {
    return apply.call(ui.selection, target)?.ok === true;
  } catch (error) {
    console.warn('[otzaria-word] החזרת הסמן למקומו נכשלה', error);
    return false;
  }
}

/**
 * גוללת אל המקום שהסמן הוחזר אליו.
 *
 * `behavior: 'instant'` ולא ברירת המחדל `'smooth'`: זו אינה ניווט שהמשתמש
 * ביקש אלא המשך של מה שהיה על המסך, ואנימציה שמתחילה בתחילת המסמך ורצה
 * לעמוד 40 היא בדיוק התחושה של „הוא איבד את המקום ועכשיו מחפש אותו”.
 */
async function scrollTo(ui: CaretUiSource, point: CaretPoint): Promise<void> {
  const scroll = ui.viewport?.scrollIntoView;
  if (typeof scroll !== 'function') return;
  try {
    await scroll.call(ui.viewport, {
      target: {
        kind: 'text',
        blockId: point.blockId,
        range: { start: point.offset, end: point.offset },
      },
      block: 'center',
      behavior: 'instant',
    });
  } catch (error) {
    console.warn('[otzaria-word] הגלילה למקום השמור נכשלה', error);
  }
}

/**
 * מחזירה את הסמן — ואת התצוגה — למקום השמור. `false` פירושו שהמקום לא נמצא,
 * והמסמך נשאר פתוח בתחילתו.
 *
 * סדר הניסיונות הוא סדר העלות: המזהה השמור, ואחריו קריאה ממוקדת לפי הסדר.
 * ההנמקה המלאה בראש הקובץ.
 */
export async function applyCaretAnchor(
  host: CaretUi,
  doc: CaretDoc,
  anchor: CaretAnchor | null | undefined,
): Promise<boolean> {
  const ui = host as CaretUiSource | null | undefined;
  if (!ui || !anchor) return false;

  /** מנסה זוג קצוות, ואם ההיסט נדחה — את תחילת אותה פסקה. ראו ראש הקובץ. */
  const attempt = async (start: CaretPoint, end: CaretPoint): Promise<boolean> => {
    if (applied(ui, selectionTargetOf(start, end))) {
      await scrollTo(ui, start);
      return true;
    }
    const head: CaretPoint = { ...start, offset: 0 };
    if (start.offset !== 0 && applied(ui, selectionTargetOf(head, head))) {
      await scrollTo(ui, head);
      return true;
    }
    return false;
  };

  const end = anchor.end ?? anchor.start;
  if (await attempt(anchor.start, end)) return true;

  // המזהים לא שרדו את סבב הייצוא. הסדר תמיד שורד — ראו „שני עוגנים” בראש.
  if (anchor.start.ordinal === null) return false;
  const startId = await blockIdAt(doc, anchor.start.ordinal);
  if (!startId) return false;
  const endOrdinal = end.ordinal ?? anchor.start.ordinal;
  const endId =
    endOrdinal === anchor.start.ordinal ? startId : (await blockIdAt(doc, endOrdinal)) ?? startId;

  return attempt({ ...anchor.start, blockId: startId }, { ...end, blockId: endId });
}

/**
 * האם יש עכשיו סמן או בחירה בטקסט המסמך — קריאה סינכרונית של תצלום הבחירה
 * בלבד, בלי לפתור סדר פסקאות.
 *
 * למי זה משמש: סמן הפתיחה (`applyDocumentStartCaret`). אם המשתמש כבר הספיק
 * ללחוץ בגוף הטקסט בזמן שהפתיחה עוד רצה, הצבת „תחילת המסמך" הייתה דורסת
 * את הקליק שלו — ולכן בודקים לפני.
 */
export function hasTextCaret(host: CaretUi): boolean {
  const ui = host as CaretUiSource | null | undefined;
  const getSnapshot = ui?.selection?.getSnapshot;
  if (typeof getSnapshot !== 'function') return false;

  let slice: SelectionSliceLike | null | undefined;
  try {
    slice = getSnapshot.call(ui?.selection);
  } catch {
    return false;
  }

  return (
    (pointsFromSelectionTarget(slice?.selectionTarget) ?? pointsFromTextTarget(slice?.target)) !==
    null
  );
}

/**
 * מציבה סמן מכווץ בתחילת המסמך — הפסקה הראשונה, היסט 0.
 *
 * למה זה קיים: `superdoc.focus()` ממקד את קלט המקלדת אבל **אינו** מציב סמן
 * כשאין למסמך בחירה קודמת לשחזר — נמדד על ה-dist הארוז: שדה הקלט של המנוע
 * נשאר „חונה" (parked), וכל הקלדה נבלעת עד קליק עכבר. מסמך חדש ומסמך שנפתח
 * בלי מקום שמור מקבלים כאן את מה ש-Word נותן בחינם: סמן שאפשר להקליד אחריו
 * מיד.
 *
 * שני מסלולים. הראשון — מהעמוד המצויר (`paintedFirstBlockId`) — **סינכרוני**:
 * כשהוא מצליח הסמן כבר במקומו כשהפונקציה מחזירה את ההבטחה, ולכן מי שאינו
 * ממתין לה עדיין מקבל סמן לפני השורה הבאה שלו. השני — `doc.blocks.list` —
 * נשאר למקרה שהעמוד אינו מצויר או שהמנוע דחה את המזהה (פסקה ראשונה שהיא
 * טבלה, למשל), והוא זה שעלול להמתין שניות ארוכות. לכן הבחירה נבדקת שוב אחריו:
 * קליק שהמשתמש הציב בזמן ההמתנה אינו נדרס.
 *
 * בלי גלילה, בכוונה: המסמך ממילא נפתח בתחילתו. `false` פירושו שאין פסקה
 * עם מזהה או שהמנוע דחה — והמסמך נשאר כפי שהיה, בלי סמן, כמו קודם.
 */
export async function applyDocumentStartCaret(
  host: CaretUi,
  doc: CaretDoc,
  container?: ParentNode | null,
): Promise<boolean> {
  const ui = host as CaretUiSource | null | undefined;
  if (typeof ui?.selection?.apply !== 'function') return false;

  const painted = paintedFirstBlockId(container);
  if (painted) {
    const head: CaretPoint = { blockId: painted, ordinal: 0, offset: 0 };
    if (applied(ui, selectionTargetOf(head, head))) return true;
  }

  const list = (doc as CaretDocSource | null | undefined)?.activeEditor?.doc?.blocks?.list;
  if (typeof list !== 'function') return false;

  let result: BlocksListLike | null | undefined;
  try {
    result = await list({ offset: 0, limit: 1 });
  } catch {
    return false;
  }

  const first = Array.isArray(result?.blocks) ? (result.blocks[0] as BlockEntryLike) : undefined;
  if (typeof first?.nodeId !== 'string' || first.nodeId === '') return false;
  if (hasTextCaret(ui)) return false;

  const head: CaretPoint = { blockId: first.nodeId, ordinal: 0, offset: 0 };
  return applied(ui, selectionTargetOf(head, head));
}
