/**
 * „פסקה” — כניסות, ריווח, אפשרויות שמירה וטאבים, דרך `doc.format.paragraph.*`.
 *
 * ## מה שנמדד לפני שנכתבה כאן השורה הראשונה
 *
 * Chrome headless על ה-dist הארוז, כל סבב מלווה בפירוק ה-zip של `export.toDocx`
 * וקריאת `document.xml` עצמו:
 *
 * ### היחידות הן twips גולמיים, אחד לאחד
 *
 *     setIndentation({left:720, right:360, firstLine:250})
 *       → <w:ind w:firstLine="250" w:left="720" w:right="360"/>
 *     setSpacing({before:240, after:120, line:480, lineRule:'exact'})
 *       → <w:spacing w:after="120" w:before="240" w:line="480" w:lineRule="exact"/>
 *
 * שונה מ-`sections.*` שם ה-API מקבל אינצ'ים וממיר לבד (`Math.round(v*1440)`).
 * כאן הערך שנשלח הוא **מה שנכתב**. לכן ההמרות מסנטימטרים ונקודות יושבות כאן,
 * בקריאה למנוע, ולא בדיאלוג.
 *
 * ### כל קריאה מחליפה את האלמנט כולו
 *
 * `setIndentation({left:-500})` אחרי הקריאה הקודמת השאיר
 * `<w:ind w:left="-500"/>` **בלבד** — `firstLine` ו-`right` נמחקו. אותו דין
 * על `w:spacing`. כלומר הפעולות האלה אינן patch אלא replace, והדיאלוג חייב
 * לשלוח **מצב מלא** של האלמנט בקריאה אחת. מסיבה זו הדיאלוג נפתח על תצלום
 * המצב הקיים (`readParagraphFormat`), ולא על ערכים ריקים.
 *
 * ### מה המנוע מאמת, ומה נשאר אצלנו
 *
 * נזרק (`INVALID_INPUT`, ולכן כל קריאה עטופה ב-catch):
 *   - ערך שאינו מספר שלם: `hanging: 0.5` → „must be a non-negative integer”.
 *   - שלילי בריווח: `before: -240` → זריקה.
 *   - enum: `lineRule:'zigzag'`, alignment `'zigzag'` בטאב → זריקה.
 *
 * **עובר בשקט** ולכן נאסר כאן לפני הקריאה:
 *   - `left: -500` → `success:true` ו-`<w:ind w:left="-500"/>`. ב-OOXML
 *     `w:left` הוא `ST_SignedTwipsMeasure`, אבל דיאלוג הפסקה של Word אינו
 *     מציע כניסה שלילית, וגם אנחנו לא נציע.
 *   - `setTabStop({position:-100})` → `success:true`. `w:pos` שלילי אינו
 *     חוקי ב-ECMA-376 — הוולידציה על ערכי הטאב יושבת כאן.
 *
 * ### NO_OP אינה שגיאה
 *
 * קריאה חוזרת עם ערכים זהים מחזירה `success:false, code:'NO_OP'` — הערכים
 * כבר מוגדרים, וזו הצלחה מבחינת המשתמש. אותה הכרעה בכל הגלים.
 *
 * ### טאבים הם רשימה, ולא ערך
 *
 * `setTabStop` **מוסיף** עצירה ואינו נוגע באחרות (נמדד: שתי קריאות השאירו
 * `<w:tab w:val="center" w:pos="1440" w:leader="dot"/>` ו-
 * `<w:tab w:val="right" w:pos="2880"/>` יחד); `clearTabStop({position})`
 * מוריד יעד יחיד; `clearAllTabStops` מוריד את `<w:tabs>` כולו. לכן הטאבים
 * הם החלק היחיד שאפשר לערוך בתוספות בטוח, וגם הקריאה של הרשימה הקיימת
 * יושבת ב-`readParagraphFormat` — מהמסמך עצמו, ולא מהנחה.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';
import { readDocSelection, type SelectionDocumentApi } from './doc-selection';

/** הנוסח שהתכנית קובעת ב-§12 לפקד שאין לו API זמין. זהה ל-page-break.ts. */
const UNAVAILABLE_TEXT = 'אינו זמין בגרסה זו';

const LOADING_TEXT = 'המסמך עדיין נטען';

/** 1440 twips לאינץ', ואינץ' הוא 2.54 ס\"מ בדיוק. */
export const TWIPS_PER_CM = 1440 / 2.54;

/** 20 twips בנקודה אחת. */
export const TWIPS_PER_PT = 20;

export type TabAlignment = 'left' | 'center' | 'right' | 'decimal' | 'bar';
export type TabLeader = 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
export type LineSpacingRule = 'auto' | 'exact' | 'atLeast';

/** עצירת טאב כפי שהיא מגיעה מהמסמך וחוזרת אליו. `positionTwips` שלם > 0. */
export interface TabStop {
  positionTwips: number;
  alignment: TabAlignment;
  leader?: TabLeader;
}

/** מצב הפסקה כפי שהוא נקרא מהמסמך. כל שדה קיים — זו תשובה של המנוע, לא שלנו. */
export interface ParagraphFormatSnapshot {
  indentation: { leftTwips: number; rightTwips: number; firstLineTwips: number; hangingTwips: number };
  spacing: { beforeTwips: number; afterTwips: number; lineTwips: number; rule: LineSpacingRule };
  keepNext: boolean;
  keepLines: boolean;
  widowControl: boolean;
  tabs: readonly TabStop[];
}

function docOf(host: ParagraphFormatTarget): ParagraphFormatDocumentApi | null {
  return (host as ParagraphFormatHost | null | undefined)?.activeEditor?.doc ?? null;
}

function unavailable(failedAction: string, detail: string, reason: string): CommandOutcome {
  return { ok: false, message: `${failedAction}: ${detail}`, reason };
}

function unsupported(failedAction: string): CommandOutcome {
  return unavailable(failedAction, UNAVAILABLE_TEXT, 'command-unsupported');
}

/** קריאה בודדת לפעולת מנוע: לעולם לא זורקת, ו-NO_OP היא הצלחה. */
async function call(
  failedAction: string,
  run: () => MaybePromise<DocReceipt>,
): Promise<CommandOutcome> {
  let receipt: DocReceipt;
  try {
    receipt = await run();
  } catch (error) {
    return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
  }
  if (receipt?.success === false && receipt.failure?.code !== 'NO_OP') {
    return { ok: false, message: receiptFailureText(failedAction, receipt), reason: receipt.failure?.code };
  }
  return { ok: true };
}

/** הצורה שנצרכת מ-`format.paragraph.*` ומ-`get`. מוגדרת כאן ולא מיובאת — ראו document-api.ts. */
export interface ParagraphFormatDocumentApi extends SelectionDocumentApi {
  format?: {
    paragraph?: {
      setIndentation?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
      clearIndentation?: (input: { target: unknown }) => MaybePromise<DocReceipt>;
      setSpacing?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
      clearSpacing?: (input: { target: unknown }) => MaybePromise<DocReceipt>;
      setKeepOptions?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
      setTabStop?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
      clearTabStop?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
      clearAllTabStops?: (input: { target: unknown }) => MaybePromise<DocReceipt>;
    };
  };
  get?: () => MaybePromise<unknown>;
}

/** מה שנדרש מ-SuperDoc: רק הפאסדה של המסמך. */
export interface ParagraphFormatHost {
  activeEditor?: { doc?: ParagraphFormatDocumentApi | null } | null;
}

export type ParagraphFormatTarget = SuperDoc | ParagraphFormatHost | null | undefined;

export interface ParagraphTarget {
  kind: 'block';
  nodeType: 'paragraph';
  nodeId: string;
  story?: unknown;
}

/**
 * הפסקה שהבחירה **מתחילה** בה — אותו פתרון יעד שנמדד ב-page-break.ts:
 * `blockId` מהבחירה, `nodeType:'paragraph'` תמיד (פתרון היעד במנוע נעשה לפי
 * `nodeId` ו-`story` בלבד), ו-`story` נשלח רק כשיש.
 */
async function resolveTarget(
  host: ParagraphFormatTarget,
): Promise<{ target: ParagraphTarget } | { error: CommandOutcome }> {
  const selection = await readDocSelection(host);
  if (!selection.blockId) {
    return { error: { ok: false, message: 'יש למקם את הסמן במסמך', reason: 'selection-required' } };
  }
  return {
    target: {
      kind: 'block',
      nodeType: 'paragraph',
      nodeId: selection.blockId,
      ...(selection.story ? { story: selection.story } : {}),
    },
  };
}

/** מאמת מספר שלם לא-שלילי ב-twips. מחזיר `null` כשהקלט פסול. */
function nonNegativeInt(value: number): number | null {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0 ? value : null;
}

/** התצלום כשאין מה לקרוא. לא קבוע משותף — כדי שקורא לא ישנה אותו לכולם. */
export function emptyParagraphFormat(): ParagraphFormatSnapshot {
  return {
    indentation: { leftTwips: 0, rightTwips: 0, firstLineTwips: 0, hangingTwips: 0 },
    spacing: { beforeTwips: 0, afterTwips: 0, lineTwips: 240, rule: 'auto' },
    keepNext: false,
    keepLines: false,
    widowControl: true,
    tabs: [],
  };
}

/**
 * המודל שנקרא מהמסמך (`doc.get`): נקודות, לא twips — ההמרה כאן ולא בדיאלוג.
 *
 * ## מה שנמדד על המודל עצמו (Chrome headless, ה-dist הארוז, מסמך חדש)
 *
 * הצורה שחוזרת מ-`doc.get()` היא:
 *
 *     { kind: 'paragraph',
 *       paragraphIds: { paraId: '41964671' },
 *       paragraph: { inlines: [...], props: { indent: {...}, spacing: {...}, bidi: true } } }
 *
 * שלושה דברים שם שונים ממה שהקוד כאן הניח, וכל אחד מהם לבדו הספיק כדי
 * שהקריאה תחזור **ריקה**:
 *
 *   1. **אין `id` על הצומת.** המזהה יושב ב-`paragraphIds.paraId`, והוא
 *      גם בדיוק ה-`blockId` שהבחירה מחזירה (נמדד: שניהם `'41964671'`).
 *   2. **המפתח הוא `indent` ולא `indentation`** — `setIndentation({left:720})`
 *      החזיר במודל `indent: { left: 36 }` (נקודות).
 *   3. **`tabs` ו-`keepNext`/`keepLines`/`widowControl` אינם במודל כלל**, גם
 *      אחרי קריאה מוצלחת ל-`setTabStop`/`setKeepOptions` שהחזירה
 *      `success: true`. המנוע כותב אותם ל-DOCX ואינו מחזיר אותם בקריאה.
 *
 * המשמעות של (1)+(2) הייתה חמורה: הדיאלוג נפתח על אפסים, ואישור שלו כתב
 * `<w:ind>` ריק — כלומר **מחק** כניסות שהמשתמש הגדיר ב-Word. זה בדיוק הכשל
 * שהערת הפתיחה של הקובץ מזהירה ממנו („דיאלוג שנפתח ריק ואושר היה מוחק בשקט”).
 *
 * שני שמות המפתחות נקראים כאן, הישן והנמדד: אין תיעוד שקובע איזה מהם החוזה,
 * וגרסה שתחזיר את השני לא תשבור את הקריאה.
 */
interface RawParagraphProps {
  /** השם שנמדד במנוע 2.8.0. */
  indent?: { left?: number; right?: number; firstLine?: number; hanging?: number };
  /** השם שהונח קודם. נשמר כרשת ביטחון — ראו הערת הפתיחה של הטיפוס. */
  indentation?: { left?: number; right?: number; firstLine?: number; hanging?: number };
  spacing?: { before?: number; after?: number; line?: number; lineRule?: string };
  keepWithNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  /** אינו מוחזר במנוע 2.8.0. ראו docs/engine-gaps.md. */
  tabs?: readonly { kind?: string; position?: number; alignment?: string; leader?: string }[];
  /** `true` בפסקה שכיוונה מימין לשמאל. נכתב על ידי `<w:bidi>`. */
  bidi?: boolean;
}

/**
 * מזהה הפסקה של הצומת, לפי שני המקומות שהוא יכול לשבת בהם.
 * `paragraphIds.paraId` הוא מה שנמדד; `id` נשאר כרשת ביטחון.
 */
function nodeParagraphId(node: object): string | undefined {
  const direct = (node as { id?: unknown }).id;
  if (typeof direct === 'string' && direct !== '') return direct;
  const paraId = (node as { paragraphIds?: { paraId?: unknown } }).paragraphIds?.paraId;
  return typeof paraId === 'string' && paraId !== '' ? paraId : undefined;
}

/**
 * התכונות של הפסקה שמזהה שלה `blockId`, מתוך המסמך שהוחזר מ-`doc.get()`.
 *
 * מיוצאת מפני שהסרגל (engine/page-ruler.ts) קורא את אותה פסקה בדיוק, ושני
 * מאתרים לאותו צומת היו נפרדים ביום שבו המודל ישנה צורה.
 */
export function findParagraphProps(document: unknown, blockId: string): RawParagraphProps | undefined {
  if (!document || typeof document !== 'object') return undefined;
  const body = (document as { body?: unknown }).body;
  if (!Array.isArray(body)) return undefined;

  for (const node of body) {
    if (!node || typeof node !== 'object') continue;
    if (nodeParagraphId(node) !== blockId) continue;
    // פסקה/כותרת/פריט רשימה נושאות את התכונות תחת המפתח של סוגן.
    const inner =
      (node as { paragraph?: { props?: RawParagraphProps } }).paragraph ??
      (node as { heading?: { props?: RawParagraphProps } }).heading ??
      (node as { list?: { props?: RawParagraphProps } }).list;
    return inner?.props;
  }
  return undefined;
}

/** הכניסות של הפסקה, ב-twips. `left`/`right` הם צד ההתחלה והסוף — ראו page-ruler.ts. */
export interface ParagraphIndents {
  leftTwips: number;
  rightTwips: number;
  firstLineTwips: number;
  hangingTwips: number;
  /** האם הפסקה מימין לשמאל (`<w:bidi>`). */
  bidi: boolean;
}

/** נקודות מהמודל → twips שלמים. ערך פסול נקרא כאפס ולא מפיל את הקריאה. */
function pointsToTwips(value: unknown): number {
  const points = typeof value === 'number' && Number.isFinite(value) ? value : 0;
  return Math.round(points * TWIPS_PER_PT);
}

/** הכניסות בלבד, מתוך תכונות פסקה שכבר נקראו. */
export function indentsFromProps(props: RawParagraphProps | undefined): ParagraphIndents {
  const indent = props?.indent ?? props?.indentation ?? {};
  return {
    leftTwips: Math.max(0, pointsToTwips(indent.left)),
    rightTwips: Math.max(0, pointsToTwips(indent.right)),
    firstLineTwips: Math.max(0, pointsToTwips(indent.firstLine)),
    hangingTwips: Math.max(0, pointsToTwips(indent.hanging)),
    bidi: props?.bidi === true,
  };
}

const LINE_RULES: readonly LineSpacingRule[] = ['auto', 'exact', 'atLeast'];
const TAB_ALIGNMENTS: readonly TabAlignment[] = ['left', 'center', 'right', 'decimal', 'bar'];
const TAB_LEADERS: readonly TabLeader[] = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'];

/**
 * קוראת את מצב הפסקה שבה הסמן, למילוי מוקדם של הדיאלוג.
 *
 * למה בכלל לקרוא: `setIndentation`/`setSpacing` **מחליפים** את האלמנט כולו
 * (ראו הערת הפתיחה). דיאלוג שנפתח ריק ואושר היה מוחק בשקט כניסות וריווח
 * שהוגדרו קודם — הרסני-למראית-עין, בדיוק התבנית שהגלים הקודמים סגרו.
 *
 * `doc.get()` מחזיר את המסמך כולו במודל SDM/1 — הפסקה מזוהה לפי
 * `paragraphIds.paraId` (ראו `findParagraphProps` ואת המדידה שמעליה), ואותן
 * התכונות ב**נקודות** ולא ב-twips. ערך שאינו מובן מוחזר כברירת מחדל ולא כשגיאה:
 * הקריאה נכשלת רק כשאין בכלל מה לפעול עליו.
 */
export async function readParagraphFormat(
  host: ParagraphFormatTarget,
): Promise<{ ok: true; target: ParagraphTarget; snapshot: ParagraphFormatSnapshot } | { ok: false; outcome: CommandOutcome }> {
  const doc = docOf(host);
  if (!doc) return { ok: false, outcome: unavailable('פתיחת תפריט הפסקה נכשלה', LOADING_TEXT, 'document-api-unavailable') };

  const resolved = await resolveTarget(host);
  if ('error' in resolved) return { ok: false, outcome: resolved.error };

  const get = doc.get;
  if (typeof get !== 'function') {
    return { ok: false, outcome: unsupported('פתיחת תפריט הפסקה נכשלה') };
  }

  let document: unknown;
  try {
    document = await get();
  } catch (error) {
    return { ok: false, outcome: { ok: false, message: thrownText('פתיחת תפריט הפסקה נכשלה', error), reason: 'threw' } };
  }

  const raw = findParagraphProps(document, resolved.target.nodeId);

  const defaults = emptyParagraphFormat();
  const ind = raw?.indent ?? raw?.indentation ?? {};
  const sp = raw?.spacing ?? {};
  const rule = LINE_RULES.includes(sp.lineRule as LineSpacingRule) ? (sp.lineRule as LineSpacingRule) : 'auto';
  const ptToTwips = (value: number | undefined): number =>
    nonNegativeInt(Math.round((value ?? 0) * TWIPS_PER_PT)) ?? 0;

  const tabs: TabStop[] = [];
  for (const tab of Array.isArray(raw?.tabs) ? raw.tabs : []) {
    if (tab?.kind !== 'set') continue;
    const positionTwips = typeof tab.position === 'number' && tab.position > 0 ? Math.round(tab.position * TWIPS_PER_PT) : 0;
    if (positionTwips <= 0) continue;
    tabs.push({
      positionTwips,
      alignment: TAB_ALIGNMENTS.includes(tab.alignment as TabAlignment) ? (tab.alignment as TabAlignment) : 'left',
      ...(TAB_LEADERS.includes(tab.leader as TabLeader) && tab.leader !== 'none'
        ? { leader: tab.leader as TabLeader }
        : {}),
    });
  }

  return {
    ok: true,
    target: resolved.target,
    snapshot: {
      indentation: {
        leftTwips: ptToTwips(ind.left),
        rightTwips: ptToTwips(ind.right),
        firstLineTwips: ptToTwips(ind.firstLine),
        hangingTwips: ptToTwips(ind.hanging),
      },
      spacing: {
        beforeTwips: ptToTwips(sp.before),
        afterTwips: ptToTwips(sp.after),
        lineTwips: nonNegativeInt(sp.line != null ? Math.round(sp.line * TWIPS_PER_PT) : NaN) ?? defaults.spacing.lineTwips,
        rule,
      },
      keepNext: raw?.keepWithNext === true,
      keepLines: raw?.keepLines === true,
      widowControl: raw?.widowControl !== false,
      tabs,
    },
  };
}

/** מה שהסרגל צריך לדעת על הפסקה שהסמן בה. */
export interface ParagraphIndentReading {
  /** היעד לכתיבה חזרה, כפי ש-`applyParagraphIndentation` מצפה לו. */
  target: ParagraphTarget;
  indents: ParagraphIndents;
}

/**
 * הכניסות של הפסקה שהסמן בה — הקריאה שהסרגל עושה על כל תזוזת סמן.
 *
 * למה קריאה נפרדת ולא `readParagraphFormat`: זו מחזירה תצלום מלא ומנסחת
 * הודעות כשל בעברית („פתיחת תפריט הפסקה נכשלה”), וזה נכון לדיאלוג שנפתח
 * בלחיצה. הסרגל אינו פעולה של המשתמש אלא תצוגה שרצה ברקע: אין לו על מה
 * להתלונן, ו„אין סמן במסמך” הוא אצלו מצב רגיל — פשוט אין סמני כניסה לצייר.
 * לכן `null` בכל מסלול שאינו מצליח, ואף הודעה.
 *
 * המחיר של `doc.get()` מוכר — הוא סורק את המסמך כולו — ולכן הקריאה מושהית
 * ב-`createRulerModel` (engine/page-ruler.ts), בדיוק כמו ספירת המילים.
 */
export async function readParagraphIndents(
  host: ParagraphFormatTarget,
): Promise<ParagraphIndentReading | null> {
  const doc = docOf(host);
  const get = doc?.get;
  if (typeof get !== 'function') return null;

  const selection = await readDocSelection(host);
  if (!selection.blockId) return null;

  let document: unknown;
  try {
    document = await get();
  } catch {
    return null;
  }

  return {
    target: {
      kind: 'block',
      nodeType: 'paragraph',
      nodeId: selection.blockId,
      ...(selection.story ? { story: selection.story } : {}),
    },
    indents: indentsFromProps(findParagraphProps(document, selection.blockId)),
  };
}

/** כניסות הפסקה, ב-twips. `special` קובע אם `firstLine` או `hanging` נשלחים — בדיוק כמו „מיוחד” ב-Word. */
export interface IndentationSettings {
  leftTwips: number;
  rightTwips: number;
  special: 'none' | 'firstLine' | 'hanging';
  amountTwips: number;
}

/**
 * שינוי הכניסות. מצב מלא בקריאה אחת — ראו „כל קריאה מחליפה את האלמנט כולו”
 * בהערת הפתיחה.
 */
export async function applyParagraphIndentation(
  host: ParagraphFormatTarget,
  target: unknown,
  settings: IndentationSettings,
): Promise<CommandOutcome> {
  const failedAction = 'שינוי הכניסות נכשל';
  const left = nonNegativeInt(settings.leftTwips);
  const right = nonNegativeInt(settings.rightTwips);
  const amount = nonNegativeInt(settings.amountTwips);
  if (left === null || right === null || amount === null) {
    return { ok: false, message: `${failedAction}: הערכים חייבים להיות מספרים לא-שליליים`, reason: 'invalid-input' };
  }

  const paragraph = docOf(host)?.format?.paragraph;
  const setIndentation = paragraph?.setIndentation;
  if (typeof setIndentation !== 'function') return unsupported(failedAction);

  // „מיוחד”: או שורה ראשונה או תלויה, לעולם לא שניהם — זו גם הסמנטיקה של Word.
  const payload: Record<string, unknown> = { target, left, right };
  if (settings.special === 'firstLine') payload.firstLine = amount;
  if (settings.special === 'hanging') payload.hanging = amount;

  return call(failedAction, () => setIndentation(payload));
}

export function clearParagraphIndentation(host: ParagraphFormatTarget, target: unknown): Promise<CommandOutcome> {
  const clear = docOf(host)?.format?.paragraph?.clearIndentation;
  if (typeof clear !== 'function') return Promise.resolve(unsupported('ניקוי הכניסות נכשל'));
  return call('ניקוי הכניסות נכשל', () => clear({ target }));
}

/** ריווח הפסקה, ב-twips. `rule:'auto'` עם `lineTwips` הוא הכפל (240=שורה, 480=כפולה). */
export interface SpacingSettings {
  beforeTwips: number;
  afterTwips: number;
  lineTwips: number;
  rule: LineSpacingRule;
}

export async function applyParagraphSpacing(
  host: ParagraphFormatTarget,
  target: unknown,
  settings: SpacingSettings,
): Promise<CommandOutcome> {
  const failedAction = 'שינוי הריווח נכשל';
  const before = nonNegativeInt(settings.beforeTwips);
  const after = nonNegativeInt(settings.afterTwips);
  const line = nonNegativeInt(settings.lineTwips);
  if (before === null || after === null || line === null) {
    return { ok: false, message: `${failedAction}: הערכים חייבים להיות מספרים לא-שליליים`, reason: 'invalid-input' };
  }
  if (!LINE_RULES.includes(settings.rule)) {
    return { ok: false, message: `${failedAction}: סוג מרווח השורות אינו חוקי`, reason: 'invalid-input' };
  }

  const setSpacing = docOf(host)?.format?.paragraph?.setSpacing;
  if (typeof setSpacing !== 'function') return unsupported(failedAction);

  return call(failedAction, () => setSpacing({ target, before, after, line, lineRule: settings.rule }));
}

export function clearParagraphSpacing(host: ParagraphFormatTarget, target: unknown): Promise<CommandOutcome> {
  const clear = docOf(host)?.format?.paragraph?.clearSpacing;
  if (typeof clear !== 'function') return Promise.resolve(unsupported('ניקוי הריווח נכשל'));
  return call('ניקוי הריווח נכשל', () => clear({ target }));
}

export async function applyParagraphKeepOptions(
  host: ParagraphFormatTarget,
  target: unknown,
  options: { keepNext: boolean; keepLines: boolean; widowControl: boolean },
): Promise<CommandOutcome> {
  const failedAction = 'שינוי אפשרויות השמירה נכשל';
  const setKeepOptions = docOf(host)?.format?.paragraph?.setKeepOptions;
  if (typeof setKeepOptions !== 'function') return unsupported(failedAction);
  return call(failedAction, () =>
    setKeepOptions({
      target,
      keepNext: options.keepNext === true,
      keepLines: options.keepLines === true,
      widowControl: options.widowControl !== false,
    }),
  );
}

const ALIGNMENT_SET: readonly string[] = TAB_ALIGNMENTS;
const LEADER_SET: readonly string[] = TAB_LEADERS;

export async function addParagraphTabStop(
  host: ParagraphFormatTarget,
  target: unknown,
  tab: TabStop,
): Promise<CommandOutcome> {
  const failedAction = 'הוספת עצירת הטאב נכשלה';
  const position = nonNegativeInt(tab.positionTwips);
  if (position === null || position <= 0) {
    // המנוע עצמו קיבל `position:-100` בחיוב (נמדד) — `w:pos` שלילי אינו חוקי
    // ב-ECMA-376, ולכן השער כאן ולא במנוע.
    return { ok: false, message: `${failedAction}: מיקום העצירה חייב להיות מספר חיובי`, reason: 'invalid-input' };
  }
  if (!ALIGNMENT_SET.includes(tab.alignment)) {
    return { ok: false, message: `${failedAction}: סוג היישור אינו חוקי`, reason: 'invalid-input' };
  }
  if (tab.leader !== undefined && !LEADER_SET.includes(tab.leader)) {
    return { ok: false, message: `${failedAction}: סוג המוביל אינו חוקי`, reason: 'invalid-input' };
  }

  const setTabStop = docOf(host)?.format?.paragraph?.setTabStop;
  if (typeof setTabStop !== 'function') return unsupported(failedAction);

  return call(failedAction, () =>
    setTabStop({
      target,
      position,
      alignment: tab.alignment,
      ...(tab.leader ? { leader: tab.leader } : {}),
    }),
  );
}

export async function removeParagraphTabStop(
  host: ParagraphFormatTarget,
  target: unknown,
  positionTwips: number,
): Promise<CommandOutcome> {
  const failedAction = 'הסרת עצירת הטאב נכשלה';
  const position = nonNegativeInt(positionTwips);
  if (position === null || position <= 0) {
    return { ok: false, message: `${failedAction}: מיקום העצירה חייב להיות מספר חיובי`, reason: 'invalid-input' };
  }
  const clearTabStop = docOf(host)?.format?.paragraph?.clearTabStop;
  if (typeof clearTabStop !== 'function') return unsupported(failedAction);
  return call(failedAction, () => clearTabStop({ target, position }));
}

export function clearAllParagraphTabStops(host: ParagraphFormatTarget, target: unknown): Promise<CommandOutcome> {
  const clearAll = docOf(host)?.format?.paragraph?.clearAllTabStops;
  if (typeof clearAll !== 'function') return Promise.resolve(unsupported('ניקוי עצירות הטאב נכשל'));
  return call('ניקוי עצירות הטאב נכשל', () => clearAll({ target }));
}




