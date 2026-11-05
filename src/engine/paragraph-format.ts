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
import { getDocumentReadRevision, invalidateDocumentReads, type DocumentReadRevision } from './document-read-state';

/** הנוסח שהתכנית קובעת ב-§12 לפקד שאין לו API זמין. זהה ל-page-break.ts. */
const UNAVAILABLE_TEXT = 'אינו זמין בגרסה זו';

const LOADING_TEXT = 'המסמך עדיין נטען';

/** הקלט של `doc.get` לכל המסמך. ריק, אבל חובה — ההנמקה ב-`get` שבטיפוס. */
const GET_ALL = Object.freeze({});

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
  /**
   * `<w:bidi>` של הפסקה — כלומר איזה צד הוא „ההתחלה” שלה. **`null` = הפסקה
   * אינה מצהירה, ואז הכיוון יורש מהמקטע.**
   *
   * של הפסקה ולא של המקטע, מפני ש-`w:start`/`w:end` שהכניסות נכתבות בהם הם
   * לוגיים לפסקה: פסקה LTR בתוך מקטע RTL מקבלת „לפני טקסט” בצד השמאלי. פס
   * התצוגה המקדימה הוא הצרן היחיד, והוא מצייר לפיו את `direction` שלו.
   *
   * ## ולמה שלושה מצבים ולא שניים
   *
   * `false` שמשמעו „לא הצהירה” הוא בדיוק התקלה ש„הכרזת bidi אינה כיוון”
   * מתארת: מסמך עברי שנוצר ב-Word אינו מצהיר `<w:bidi>` על כל פסקה — הוא
   * יורש — ופס שהיה קורא היעדר כ-LTR היה מצייר את הכניסה בצד ההפוך על
   * מסמכים רגילים לגמרי.
   *
   * שלושת המצבים אפשריים מפני שהמודל **מבדיל** ביניהם, וזה נמדד בכרום על
   * המנוע האמיתי: פסקה שנקבעה LTR מחזירה `props: { bidi: false }` — הערך
   * `false` ולא היעדר — ולכן `'bidi' in props` הוא הבחנה אמיתית. זה גם
   * עולה בקנה אחד עם מה שהערת `RawParagraphProps` מתעדת: המודל מוציא
   * מפתחות שאין לו מה לומר עליהם (‏`tabs`, `keepNext` ואחרים אינם בו כלל).
   *
   * `null` גם כשאין תכונות פסקה בכלל — כולל הפער שהערת `findParagraphProps`
   * מתעדת, פסקה בתוך תא טבלה שאינה נפתרת שם.
   */
  bidi: boolean | null;
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
  doc: ParagraphFormatDocumentApi | null,
  failedAction: string,
  run: () => MaybePromise<DocReceipt>,
): Promise<CommandOutcome> {
  let receipt: DocReceipt;
  try {
    receipt = await run();
  } catch (error) {
    return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
  } finally {
    // גם כשל עשוי להגיע אחרי שינוי חלקי. שומרים את המסמך שעליו הפעולה התחילה.
    invalidateDocumentReads(doc);
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
      /**
       * אפשרויות הזרימה. נצרך כאן ל-`contextualSpacing` בלבד; `pageBreakBefore`
       * עובר דרך engine/page-break.ts, ושתי הקריאות אינן מתנגשות — הפעולה היא
       * patch (נמדד, ראו `applyParagraphContextualSpacing`).
       */
      setFlowOptions?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
      setTabStop?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
      clearTabStop?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
      clearAllTabStops?: (input: { target: unknown }) => MaybePromise<DocReceipt>;
    };
  };
  /**
   * הקלט חובה, גם כשהוא ריק: מ-superdoc 2.23 `doc.get()` בלי ארגומנט זורק
   * „Cannot read properties of undefined (reading 'options')”, ו-`doc.get({})`
   * מחזיר את המסמך (נמדד; ב-2.20 שתיהן עבדו). מכאן `GET_ALL`.
   */
  get?: (input: typeof GET_ALL) => MaybePromise<unknown>;
  /** Public 2.23 node reads return an SDNodeResult envelope, not the node itself. */
  getNodeById?: (input: { nodeId: string }) => MaybePromise<unknown>;
  getNode?: (address: ParagraphTarget) => MaybePromise<unknown>;
  blocks?: {
    list?: (input: { nodeIds: readonly string[]; in?: unknown }) => MaybePromise<
      { blocks?: readonly { nodeId?: string; nodeType?: string }[] } | undefined
    >;
  };
  lists?: {
    getState?: (input: {
      target: { kind: 'block'; nodeType: 'paragraph' | 'listItem'; nodeId: string; story?: unknown };
    }) => MaybePromise<{ success?: boolean; isListItem?: boolean } | undefined>;
  };
}

/** מה שנדרש מ-SuperDoc: רק הפאסדה של המסמך. */
export interface ParagraphFormatHost {
  activeEditor?: { doc?: ParagraphFormatDocumentApi | null } | null;
}

export type ParagraphFormatTarget = SuperDoc | ParagraphFormatHost | null | undefined;

/** שלושת סוגי הבלוק שדיאלוג ההפסקה יכול לחול עליהם — `ParagraphBlockType` בחוזה. */
export type ParagraphBlockType = 'paragraph' | 'heading' | 'listItem';

export interface ParagraphTarget {
  kind: 'block';
  nodeType: ParagraphBlockType;
  nodeId: string;
  story?: unknown;
}

/**
 * סוג הבלוק בפועל של המזהה — פסקה, כותרת או פריט רשימה — לפי `blocks.list`,
 * ולפריט שאינו מופיע שם — לפי `lists.getState`.
 *
 * החוזה (`ParagraphTarget` ב-paragraphs.types.d.ts) דורש את הסוג האמיתי,
 * ולא ליטרל מקובע: כתובת עם `nodeType:'paragraph'` על בלוק שהוא בפועל
 * `heading` או `listItem` היא כתובת פסולה, וכתיבה חזרה אליה נכשלת.
 *
 * `blocks.list({nodeIds})` מסנן לפני העימוד, אך מונה בלוקים עליונים בלבד,
 * ולכן פריט רשימה בתוך תא טבלה אינו שם; שם
 * מכריע `lists.getState`, כמו ב-`resolveListItem` ב-lists.ts. כותרת
 * ממוספרת נשארת `heading`: זה מה ש-`blocks.list` מדווח, וכתובת ההפסקה
 * שלה היא כתובת כותרת.
 *
 * ברירת המחדל `'paragraph'` — כשאין אף אחת מהפעולות, וכשהמזהה לא נמצא —
 * היא התאמה לאחור: פסקה היא הסוג הנפוץ, וגרסת מנוע ישנה שאינה חושפת את
 * הפעולות לא הייתה מפסיקה לעבוד על המקרה הרגיל.
 */
async function resolveBlockType(
  doc: ParagraphFormatDocumentApi,
  blockId: string,
  story?: unknown,
  fallback: ParagraphBlockType = 'paragraph',
): Promise<ParagraphBlockType> {
  const list = doc.blocks?.list;
  if (typeof list === 'function') {
    try {
      const listed = await list.call(doc.blocks, { nodeIds: [blockId], ...(story ? { in: story } : {}) });
      const found = (listed?.blocks ?? []).find((b) => b.nodeId === blockId);
      if (found) {
        return found.nodeType === 'heading' || found.nodeType === 'listItem'
          ? found.nodeType
          : 'paragraph';
      }
    } catch {
      // נופלים ל-lists.getState.
    }
  }

  // A targeted read resolves headings inside tables, which blocks.list omits.
  if (fallback !== 'paragraph') return fallback;
  const getState = doc.lists?.getState;
  if (typeof getState !== 'function') return 'paragraph';
  try {
    const state = await getState.call(doc.lists, {
      target: { kind: 'block', nodeType: 'paragraph', nodeId: blockId, ...(story ? { story } : {}) },
    });
    return state?.success === true && state.isListItem ? 'listItem' : 'paragraph';
  } catch {
    return 'paragraph';
  }
}

/**
 * הפסקה/כותרת/פריט הרשימה שהבחירה **מתחילה** בה — אותו פתרון יעד שנמדד
 * ב-page-break.ts, עם תיקון אחד: `nodeType` נגזר מ-`blocks.list` ולא מקובע
 * ל-`'paragraph'` — ראו `resolveBlockType`. `blockId` מהבחירה, ו-`story`
 * נשלח רק כשיש.
 */
async function resolveTarget(
  host: ParagraphFormatTarget,
): Promise<{ target: ParagraphTarget } | { error: CommandOutcome }> {
  const selection = await readDocSelection(host);
  if (!selection.blockId) {
    return { error: { ok: false, message: 'יש למקם את הסמן במסמך', reason: 'selection-required' } };
  }
  const doc = docOf(host);
  const nodeType = doc ? await resolveBlockType(doc, selection.blockId, selection.story) : 'paragraph';
  return {
    target: {
      kind: 'block',
      nodeType,
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
    bidi: null,
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
  /**
   * השם שנמדד במנוע 2.8.0. `start`/`end` הם הצד הלוגי (`SDParagraphProps`
   * ב-sd-props.d.ts) — ראו ההערה על `indentsFromProps` למטה למה שניהם נקראים.
   */
  indent?: { left?: number; right?: number; start?: number; end?: number; firstLine?: number; hanging?: number };
  /** השם שהונח קודם. נשמר כרשת ביטחון — ראו הערת הפתיחה של הטיפוס. */
  indentation?: { left?: number; right?: number; start?: number; end?: number; firstLine?: number; hanging?: number };
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
 * Explicit dialog/selection reads still consume the full document model here.
 * This legacy walker covers top-level body nodes; background direction/ruler
 * reads use readTargetedParagraph, which also resolves table and other stories.
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

/**
 * הכניסות בלבד, מתוך תכונות פסקה שכבר נקראו.
 *
 * `start`/`end` נקראים **לפני** `left`/`right`, ולא כתחליף גיבוי גרידא:
 * `docs/engine-gaps.md` מדד ש-`setIndentation({left,right})` שלנו נכתב
 * כ-`w:start`/`w:end` הלוגיים (הצד ההתחלה/סוף, לא פיזי ימין/שמאל), ו-
 * `SDParagraphProps` (sd-props.d.ts) חושף את `indent.start`/`indent.end` בדיוק
 * בשביל זה. בלי הקדימה הזאת — משתמש שקובע כניסה בדיאלוג, סוגר ופותח מחדש,
 * היה רואה אפסים: הערך יושב תחת `start`/`end` ולא תחת `left`/`right`.
 */
export function indentsFromProps(props: RawParagraphProps | undefined): ParagraphIndents {
  const indent = props?.indent ?? props?.indentation ?? {};
  return {
    leftTwips: Math.max(0, pointsToTwips(indent.start ?? indent.left)),
    rightTwips: Math.max(0, pointsToTwips(indent.end ?? indent.right)),
    firstLineTwips: Math.max(0, pointsToTwips(indent.firstLine)),
    hangingTwips: Math.max(0, pointsToTwips(indent.hanging)),
    bidi: props?.bidi === true,
  };
}

/**
 * שלושת המצבים של `<w:bidi>`: הצהירה RTL, הצהירה LTR, או שתקה (`null`).
 *
 * `'bidi' in props` ולא `props.bidi === true`, וזו ההבחנה כולה — ראו `bidi`
 * ב-`ParagraphFormatSnapshot`. `indentsFromProps` לידו **אינו** משתמש בזה
 * בכוונה: הצרן שלו הוא חיווי הכיוון ברצועה, ושם „לא הצהירה” אינו מדליק
 * כפתור, כלומר שני מצבים מספיקים.
 */
function declaredBidi(props: RawParagraphProps | undefined): boolean | null {
  if (!props || !Object.prototype.hasOwnProperty.call(props, 'bidi')) return null;
  return props.bidi === true;
}

/**
 * אילו מהערכים ב-`snapshot` הפסקה **מצהירה בעצמה** (עיצוב ישיר). ה-snapshot
 * מחזיר 0 לערך שאינו מוצהר — וזה נכון לדיאלוג „פסקה”, שכותב מצב מלא — אבל
 * „סגנון מהבחירה” חייב להבחין: 0 שאינו מוצהר הוא „מה שהסגנון נותן”, ולא
 * „אפס”. בלי ההבחנה „עדכן כך שיתאים” מחק את הריווח של כותרת 1 (נמדד ב-QA).
 */
export interface ParagraphDeclarations {
  spaceBefore: boolean;
  spaceAfter: boolean;
  line: boolean;
  indentStart: boolean;
  indentEnd: boolean;
  firstLine: boolean;
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
): Promise<
  | { ok: true; target: ParagraphTarget; snapshot: ParagraphFormatSnapshot; declared: ParagraphDeclarations }
  | { ok: false; outcome: CommandOutcome }
> {
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
    document = await get(GET_ALL);
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
    declared: {
      spaceBefore: typeof sp.before === 'number',
      spaceAfter: typeof sp.after === 'number',
      line: typeof sp.line === 'number',
      indentStart: typeof (ind.start ?? ind.left) === 'number',
      indentEnd: typeof (ind.end ?? ind.right) === 'number',
      firstLine: typeof ind.firstLine === 'number' || typeof ind.hanging === 'number',
    },
    snapshot: {
      indentation: {
        // `start`/`end` לפני `left`/`right` — ראו ההערה על `indentsFromProps`.
        leftTwips: ptToTwips(ind.start ?? ind.left),
        rightTwips: ptToTwips(ind.end ?? ind.right),
        firstLineTwips: ptToTwips(ind.firstLine),
        hangingTwips: ptToTwips(ind.hanging),
      },
      spacing: {
        beforeTwips: ptToTwips(sp.before),
        afterTwips: ptToTwips(sp.after),
        lineTwips: lineTwipsFromModel(sp.line, rule) ?? defaults.spacing.lineTwips,
        rule,
      },
      keepNext: raw?.keepWithNext === true,
      keepLines: raw?.keepLines === true,
      widowControl: raw?.widowControl !== false,
      tabs,
      bidi: declaredBidi(raw),
    },
  };
}

/** מה שהסרגל צריך לדעת על הפסקה שהסמן בה. */
export interface ParagraphIndentReading {
  /** היעד לכתיבה חזרה, כפי ש-`applyParagraphIndentation` מצפה לו. */
  target: ParagraphTarget;
  indents: ParagraphIndents;
}

/** A canonical story key keeps equal paragraph IDs in different stories separate. */
function storyKey(story: unknown): string | null {
  if (!story || (story as { storyType?: unknown }).storyType === 'body') return '';
  try {
    return JSON.stringify(story, (_key, value: unknown) => {
      if (!value || typeof value !== 'object' || Array.isArray(value)) return value;
      return Object.fromEntries(Object.entries(value).sort(([a], [b]) => a.localeCompare(b)));
    }) ?? null;
  } catch {
    return null;
  }
}

/** Ruler equality must include the write target's type and story, not just its ID. */
export function sameParagraphTarget(a: ParagraphTarget | null, b: ParagraphTarget | null): boolean {
  if (!a || !b) return a === b;
  const story = storyKey(a.story);
  return a.nodeId === b.nodeId && a.nodeType === b.nodeType && story !== null && story === storyKey(b.story);
}

interface ParagraphReadCache {
  revision: DocumentReadRevision;
  lastKey?: string;
  last?: ParagraphIndentReading;
  pending: Map<string, Promise<ParagraphIndentReading | null>>;
}

// One completed paragraph per document, plus only currently pending requests.
// A new revision never reuses a previous revision's cached reading.
const paragraphReads = new WeakMap<ParagraphFormatDocumentApi, ParagraphReadCache>();

async function readTargetedParagraph(
  doc: ParagraphFormatDocumentApi,
  nodeId: string,
  story: unknown,
  revision: DocumentReadRevision,
): Promise<ParagraphIndentReading | null> {
  const address: ParagraphTarget = {
    kind: 'block', nodeType: 'paragraph', nodeId, ...(story ? { story } : {}),
  };
  let envelope: unknown;
  // getNodeById has no story parameter in 2.23. A non-body read must use getNode.
  if (storyKey(story) === '' && typeof doc.getNodeById === 'function') {
    envelope = await doc.getNodeById({ nodeId });
  } else if (typeof doc.getNode === 'function') {
    envelope = await doc.getNode(address);
  } else {
    return null;
  }
  if (getDocumentReadRevision(doc) !== revision || !envelope || typeof envelope !== 'object') return null;
  const { node, address: resolved } = envelope as { node?: unknown; address?: ParagraphTarget };
  if (!node || typeof node !== 'object' || resolved?.kind !== 'block' || resolved.nodeId !== nodeId) return null;
  const id = nodeParagraphId(node);
  if (id && id !== nodeId) return null;
  const projected = node as { kind?: string; paragraph?: { props?: RawParagraphProps }; heading?: { props?: RawParagraphProps } };
  if (projected.kind !== 'paragraph' && projected.kind !== 'heading') return null;

  // A list item is projected as a paragraph by 2.23. Resolve its write type with
  // a filtered list, with lists.getState as the fallback for paragraphs in cells.
  const hint = resolved.nodeType === 'heading' || resolved.nodeType === 'listItem' ? resolved.nodeType : 'paragraph';
  const nodeType = await resolveBlockType(doc, nodeId, story, hint);
  return {
    target: { ...address, nodeType },
    indents: indentsFromProps(projected.paragraph?.props ?? projected.heading?.props),
  };
}

/**
 * Home's direction indicator and the ruler share one targeted read per paragraph
 * and document revision. Caret offsets do not affect the key. No full-document
 * read or unfiltered blocks.list is allowed on this background path.
 *
 * A dispatched engine call cannot be cancelled; its result is discarded when
 * its document/revision changes. UI consumers additionally reject old selections.
 */
export async function readParagraphIndents(
  host: ParagraphFormatTarget,
): Promise<ParagraphIndentReading | null> {
  const doc = docOf(host);
  if (!doc) return null;
  const revision = getDocumentReadRevision(doc);
  const current = () => docOf(host) === doc && getDocumentReadRevision(doc) === revision;
  const selection = await readDocSelection(host);
  if (!selection.blockId || !current()) return null;
  const story = storyKey(selection.story);
  if (story === null) return null;
  const key = JSON.stringify([selection.blockId, story]);

  let cache = paragraphReads.get(doc);
  if (!cache || cache.revision !== revision) {
    cache = { revision, pending: new Map() };
    paragraphReads.set(doc, cache);
  }
  if (cache.lastKey === key && cache.last) return cache.last;
  cache.lastKey = key;
  cache.last = undefined;
  let pending = cache.pending.get(key);
  if (!pending) {
    const owner = cache;
    pending = readTargetedParagraph(doc, selection.blockId, selection.story, revision)
      .catch(() => null)
      .then((reading) => {
        owner.pending.delete(key);
        if (reading && getDocumentReadRevision(doc) === revision && owner.lastKey === key) owner.last = reading;
        return reading;
      });
    cache.pending.set(key, pending);
  }
  const reading = await pending;
  return current() ? reading : null;
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

  const doc = docOf(host);
  const paragraph = doc?.format?.paragraph;
  const setIndentation = paragraph?.setIndentation;
  if (typeof setIndentation !== 'function') return unsupported(failedAction);

  // „מיוחד”: או שורה ראשונה או תלויה, לעולם לא שניהם — זו גם הסמנטיקה של Word.
  const payload: Record<string, unknown> = { target, left, right };
  if (settings.special === 'firstLine') payload.firstLine = amount;
  if (settings.special === 'hanging') payload.hanging = amount;

  return call(doc, failedAction, () => setIndentation(payload));
}

export function clearParagraphIndentation(host: ParagraphFormatTarget, target: unknown): Promise<CommandOutcome> {
  const doc = docOf(host);
  const clear = doc?.format?.paragraph?.clearIndentation;
  if (typeof clear !== 'function') return Promise.resolve(unsupported('ניקוי הכניסות נכשל'));
  return call(doc, 'ניקוי הכניסות נכשל', () => clear({ target }));
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

  const doc = docOf(host);
  const setSpacing = doc?.format?.paragraph?.setSpacing;
  if (typeof setSpacing !== 'function') return unsupported(failedAction);

  return call(doc, failedAction, () => setSpacing({ target, before, after, line, lineRule: settings.rule }));
}

export function clearParagraphSpacing(host: ParagraphFormatTarget, target: unknown): Promise<CommandOutcome> {
  const doc = docOf(host);
  const clear = doc?.format?.paragraph?.clearSpacing;
  if (typeof clear !== 'function') return Promise.resolve(unsupported('ניקוי הריווח נכשל'));
  return call(doc, 'ניקוי הריווח נכשל', () => clear({ target }));
}

/* ------------------------------------------------------------------ */
/* ריווח על **כל** הבחירה — מה שתפריט „מרווח שורות וריווח” ברצועה מפעיל */
/* ------------------------------------------------------------------ */

/**
 * סוג הבלוק של כמה מזהים, בקריאת `blocks.list` **אחת**.
 *
 * `resolveBlockType` שמעליו פותר מזהה יחיד וקורא ל-`blocks.list` בכל פעם.
 * לפסקה אחת זה נכון; לבחירה של ארבעים פסקאות זה ארבעים סריקות של אותו סיפור
 * בדיוק. כאן הרשימה נקראת פעם אחת ונשלפת ממנה מפה.
 *
 * הנפילה לאחור שונה מזו של היחיד בכוונה: שם, מזהה שאינו ב-`blocks.list`
 * נבדק ב-`lists.getState` (פריט רשימה בתוך תא טבלה). כאן לא — זו קריאה לכל
 * מזהה חסר, כלומר בדיוק העלות שהפונקציה הזאת קיימת כדי למנוע. מזהה שאינו
 * ברשימה מקבל `'paragraph'`, וכתיבה אליו שתיכשל תדווח ככשל ככל כתיבה אחרת.
 */
async function resolveBlockTypes(
  doc: ParagraphFormatDocumentApi,
  blockIds: readonly string[],
  story?: unknown,
): Promise<Map<string, ParagraphBlockType>> {
  const types = new Map<string, ParagraphBlockType>();
  const list = doc.blocks?.list;
  if (typeof list !== 'function') return types;
  try {
    const listed = await list.call(doc.blocks, { nodeIds: blockIds, ...(story ? { in: story } : {}) });
    for (const block of listed?.blocks ?? []) {
      if (typeof block.nodeId !== 'string' || !blockIds.includes(block.nodeId)) continue;
      types.set(
        block.nodeId,
        block.nodeType === 'heading' || block.nodeType === 'listItem' ? block.nodeType : 'paragraph',
      );
    }
  } catch {
    // רשימה שנכשלה אינה סיבה לוותר: הכל ייקרא כ„פסקה”, כמו ביחיד.
  }
  return types;
}

/**
 * הריווח כפי שהוא **מוצהר בפסקה עצמה**. `null` = לא הוצהר, כלומר יורש
 * מהסגנון — וזו הבחנה שאי אפשר לוותר עליה, ראו `applySelectionSpacing`.
 */
export interface DeclaredSpacing {
  beforeTwips: number | null;
  afterTwips: number | null;
  lineTwips: number | null;
  rule: LineSpacingRule | null;
}

/** פסקה אחת בבחירה: לאן לכתוב, ומה מוצהר בה עכשיו. */
export interface ParagraphSpacingEntry {
  target: ParagraphTarget;
  spacing: DeclaredSpacing;
}

/**
 * הריווח של **כל** הפסקאות שהבחירה נוגעת בהן.
 *
 * ## למה קריאה, ולא רק כתיבה
 *
 * `setSpacing` מחליף את `<w:spacing>` כולו (ראו הערת הפתיחה). „הוסף רווח לפני
 * הפסקה” שהיה שולח `{before}` בלבד היה מוחק בשקט את מרווח השורות ואת הריווח
 * שאחרי — כלומר פעולה שנראית תוספת ומתנהגת כמחיקה. לכן כל פריט ברשימה נושא
 * את **המצב המלא**, והכותב מחליף בו שדה אחד.
 *
 * ## ולמה לא `readParagraphFormat` בלולאה
 *
 * `doc.get()` סורק את המסמך כולו. הוא נקרא כאן **פעם אחת** לכל הבחירה, ולא
 * פעם לכל פסקה — שלושים פסקאות מסומנות היו שלושים סריקות מלאות.
 *
 * `null` בכל מסלול שאינו מצליח, ובלי הודעה: הצרכן הראשון הוא תפריט שנפתח,
 * ותפריט אינו המקום להתלונן בו על „אין סמן במסמך”. מי שכותב בפועל כן מדווח.
 */
export async function readSelectionSpacing(
  host: ParagraphFormatTarget,
): Promise<readonly ParagraphSpacingEntry[] | null> {
  const doc = docOf(host);
  if (!doc || typeof doc.get !== 'function') return null;

  const selection = await readDocSelection(host);
  if (!selection.blockIds.length) return null;

  let document: unknown;
  try {
    document = await doc.get(GET_ALL);
  } catch {
    return null;
  }

  const types = await resolveBlockTypes(doc, selection.blockIds, selection.story);
  return selection.blockIds.map((blockId) => ({
    target: {
      kind: 'block' as const,
      nodeType: types.get(blockId) ?? 'paragraph',
      nodeId: blockId,
      ...(selection.story ? { story: selection.story } : {}),
    },
    spacing: spacingFromProps(findParagraphProps(document, blockId)),
  }));
}

/**
 * מה **מוצהר** בפסקה, ולא מה שיוצא ממנה.
 *
 * `readParagraphFormat` שלידו ממלא כל שדה חסר בברירת מחדל, וזה נכון שם: הוא
 * ממלא דיאלוג, ודיאלוג חייב להציג מספר בכל תיבה. כאן ההפך — שדה שלא הוצהר
 * **חייב** להישאר `null`, אחרת הכתיבה חזרה תקבע אותו, והפסקה תפסיק לרשת
 * אותו מהסגנון. ההסבר המלא ב-`applySelectionSpacing`.
 *
 * המודל אינו מחזיר מפתחות שאין לו מה לומר עליהם — ראו הערת `RawParagraphProps`
 * — ולכן היעדר מפתח הוא תשובה, לא חוסר מידע.
 */
/**
 * `spacing.line` שהמודל מחזיר → twips, לפי הכלל שלצדו.
 *
 * ## למה שתי יחידות ולא אחת
 *
 * ב-OOXML ל-`w:line` יש שתי משמעויות לפי `w:lineRule`: ב-`exact`/`atLeast`
 * הוא מרחק, וב-`auto` הוא כפולה ב-240ths. המודל משקף בדיוק את ההבחנה הזאת,
 * וזה **נמדד** (superdoc 2.14.0-next.5, Chrome, ה-dist הארוז —
 * `scripts/line-unit-probe.mjs`):
 *
 * | נכתב | lineRule | המודל מחזיר |
 * |---|---|---|
 * | 480 | `auto` | **2** — כפולה |
 * | 720 | `auto` | **3** — כפולה |
 * | 360 | `exact` | **18** — נקודות |
 * | 360 | `atLeast` | **18** — נקודות |
 *
 * ## ומה זה תיקן
 *
 * המרה אחידה של `× 20` הייתה כאן קודם, והיא שגויה ב-`auto` פי 12. שני
 * כשלים אמיתיים נגזרו ממנה, ושניהם שקטים:
 *
 * 1. **הדיאלוג הציג „בודדת” לכל פסקה במרווח אוטומטי.** פסקה ב-1.5 שורות
 *    חוזרת מהמודל כ-`1.5`, יצאה מכאן כ-30 twips, לא התאימה לאף אחד מ-
 *    240/360/480, והבורר נפל לברירת המחדל. „אישור” על אותו דיאלוג **כתב**
 *    240 — כלומר שינה את מרווח השורות של המשתמש בלי שביקש.
 * 2. **„הוסף רווח לפני הפסקה” כיווץ את מרווח השורות בכל לחיצה.** הפעולה
 *    משמרת את מה שמוצהר, ולכן היא קראה 3, שלחה 60, קראה 0.25, ושלחה 5 —
 *    נתפס בשער `home-paragraph-qa` על `<w:spacing w:line="5"/>`.
 */
function lineTwipsFromModel(value: number | undefined, rule: LineSpacingRule): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value)) return null;
  return Math.max(0, Math.round(value * (rule === 'auto' ? 240 : TWIPS_PER_PT)));
}

function spacingFromProps(props: RawParagraphProps | undefined): DeclaredSpacing {
  const sp = props?.spacing;
  const declared = (value: number | undefined): number | null =>
    typeof value === 'number' && Number.isFinite(value) ? Math.max(0, pointsToTwips(value)) : null;
  const rule = LINE_RULES.includes(sp?.lineRule as LineSpacingRule) ? (sp?.lineRule as LineSpacingRule) : null;
  return {
    beforeTwips: declared(sp?.before),
    afterTwips: declared(sp?.after),
    // `rule ?? 'auto'` ולא `rule`: פסקה שמצהירה `line` בלי `lineRule` היא
    // `auto` ב-OOXML, וזו גם ברירת המחדל שהכתיבה חזרה תשלח.
    lineTwips: lineTwipsFromModel(sp?.line, rule ?? 'auto'),
    rule,
  };
}

/**
 * קובעת ריווח על רשימת פסקאות — צד אחד (`before`/`after`), והשאר נשמר.
 *
 * ## מה נשלח, ולמה דווקא זה
 *
 * `setSpacing` **מחליף** את `<w:spacing>` כולו (הערת הפתיחה), ולכן יש כאן שתי
 * דרכים ליפול, לא אחת:
 *
 * 1. לשלוח את הצד שהשתנה בלבד — ו**למחוק** מרווח שורות שהיה מוצהר בפסקה.
 * 2. לשלוח מצב מלא תמיד — ו**לקבע** בפסקה מרווח שורות שהיא ירשה מהסגנון.
 *    הפסקה תיראה זהה באותו רגע, ותפסיק להשתנות ביום שהסגנון ישתנה. זו בדיוק
 *    התקלה ש-`FONT_PREVIEW_ENABLED` מתעדת בצד הריצות (`rFonts` מפורש שנשאר
 *    בקובץ), והיא חמורה יותר בפסקה: בספר שכולו בנוי על סגנון אחד, „הוסף רווח”
 *    על פרק שלם היה מנתק את כל הפסקאות שלו מהסגנון.
 *
 * לכן נשלח **בדיוק מה שהיה מוצהר**, ועליו הצד שהשתנה: `null` אינו נשלח כלל.
 * זה מה ש-`DeclaredSpacing` קיים בשבילו.
 *
 * ## ושלוש הכרעות על הלולאה
 *
 * 1. **סדרתי ולא `Promise.all`.** כל כתיבה היא מוטציה על אותו מסמך; שליחת
 *    ארבעים במקביל אינה מהירה יותר אלא בלתי-צפויה, ודיווח הכשל שלה אינו ניתן
 *    לשיוך לפסקה.
 * 2. **כשל אחד עוצר.** המשך אחרי כשל היה משאיר בחירה שחציה קיבלה את הריווח
 *    וחציה לא — מצב שאין למשתמש דרך לראות ואין לו דרך לבטל בצעד אחד.
 * 3. **`NO_OP` אינה כשל** — `call` כבר בולעת אותה. פסקה שכבר בערך המבוקש
 *    (מצב רגיל בבחירה מרובה) אינה מפילה את השאר.
 */
export async function applySelectionSpacing(
  host: ParagraphFormatTarget,
  entries: readonly ParagraphSpacingEntry[],
  patch: { beforeTwips?: number; afterTwips?: number },
): Promise<CommandOutcome> {
  const failedAction = 'שינוי הריווח נכשל';
  if (!entries.length) {
    return { ok: false, message: 'יש למקם את הסמן במסמך', reason: 'selection-required' };
  }
  const doc = docOf(host);
  const setSpacing = doc?.format?.paragraph?.setSpacing;
  if (typeof setSpacing !== 'function') return unsupported(failedAction);

  for (const entry of entries) {
    const before = patch.beforeTwips ?? entry.spacing.beforeTwips;
    const after = patch.afterTwips ?? entry.spacing.afterTwips;
    if ((before !== null && nonNegativeInt(before) === null) || (after !== null && nonNegativeInt(after) === null)) {
      return { ok: false, message: `${failedAction}: הערכים חייבים להיות מספרים לא-שליליים`, reason: 'invalid-input' };
    }
    const outcome = await call(doc, failedAction, () =>
      setSpacing({
        target: entry.target,
        ...(before === null ? {} : { before }),
        ...(after === null ? {} : { after }),
        ...(entry.spacing.lineTwips === null ? {} : { line: entry.spacing.lineTwips }),
        ...(entry.spacing.rule === null ? {} : { lineRule: entry.spacing.rule }),
      }),
    );
    if (!outcome.ok) return outcome;
  }
  return { ok: true };
}

/**
 * מתג „הוסף/הסר רווח לפני/אחרי הפסקה”, על כל הבחירה.
 *
 * הכלל עצמו יושב כאן ולא בפקד, מפני שיש לו **שני** אתרי קריאה: פריט התפריט
 * ברצועה והקיצור Ctrl+0. שני מימושים של „מתי זה מוסיף ומתי מסיר” היו מתפצלים
 * ביום שאחד מהם יתוקן, והמשתמש היה מקבל תשובה אחרת מהמקלדת ומהעכבר.
 *
 * ההכרעה על הכיוון נופלת על **הפסקה הראשונה בבחירה**, וכל השאר מקבלות את
 * אותה תוצאה — כמו ב-Word. החלופה, החלטה לכל פסקה בנפרד, הייתה הופכת בחירה
 * מעורבת לפעולה שמשאירה אותה מעורבת בדיוק כפי שהייתה, רק הפוך.
 */
export async function toggleSelectionSpacing(
  host: ParagraphFormatTarget,
  side: 'before' | 'after',
  stepTwips: number,
): Promise<CommandOutcome> {
  const entries = await readSelectionSpacing(host);
  if (!entries?.length) {
    return { ok: false, message: 'יש למקם את הסמן במסמך', reason: 'selection-required' };
  }
  const current = side === 'before' ? entries[0].spacing.beforeTwips : entries[0].spacing.afterTwips;
  const next = (current ?? 0) > 0 ? 0 : stepTwips;
  return applySelectionSpacing(host, entries, {
    [side === 'before' ? 'beforeTwips' : 'afterTwips']: next,
  });
}

/**
 * „אל תוסיף רווח בין פסקאות מאותו סגנון” — `w:contextualSpacing`.
 *
 * ## שתי מדידות שקבעו את צורת הפקד, ולא את המימוש
 *
 * נמדד ב-superdoc 2.14.0-next.5 על ה-dist הארוז (`scripts/flow-options-probe.mjs`):
 *
 * 1. **`setFlowOptions` הוא patch, לא replace.** `pageBreakBefore: true`
 *    ואחריו `contextualSpacing: true` על אותה פסקה השאירו את שתיהן:
 *    `<w:pPr><w:pageBreakBefore/><w:bidi/><w:contextualSpacing/></w:pPr>`.
 *    זה שונה מ-`setSpacing` שלידו, ולכן הכתיבה כאן **אינה** חייבת לשלוח מצב
 *    מלא — ומפתח שאינו נשלח אינו נוגע במה שקיים. זו גם הסיבה שכתיבה מכאן
 *    אינה מתנגשת ב„מעבר עמוד לפני” שנכתב מ-page-break.ts.
 * 2. **`doc.get()` אינו מחזיר את הערך.** אחרי ששתי הכתיבות הצליחו ונחתו
 *    ב-XML, ה-props של אותה פסקה במודל היו `{ bidi: true }` בלבד — בדיוק
 *    כמו `pageBreakBefore` (ראו page-break.ts) ו-`keepNext` (ראו
 *    `RawParagraphProps`).
 *
 * המסקנה של (2) היא על **הפקד**: תיבת סימון דו-מצבית הייתה נפתחת תמיד ריקה,
 * ו„אישור” היה שולח `false` — כלומר מכבה הגדרה שהמשתמש קבע ב-Word בלי שביקש.
 * לכן בדיאלוג יושב `TriToggle` ולא תיבה, ומצב „ללא שינוי” אינו קורא לכאן
 * בכלל. ההנמקה המלאה של התבנית ב-`ui/panels/common/TriToggle.vue`.
 */
export async function applyParagraphContextualSpacing(
  host: ParagraphFormatTarget,
  target: unknown,
  contextualSpacing: boolean,
): Promise<CommandOutcome> {
  const failedAction = 'שינוי הריווח בין פסקאות מאותו סגנון נכשל';
  const doc = docOf(host);
  const setFlowOptions = doc?.format?.paragraph?.setFlowOptions;
  if (typeof setFlowOptions !== 'function') return unsupported(failedAction);
  return call(doc, failedAction, () => setFlowOptions({ target, contextualSpacing }));
}

export async function applyParagraphKeepOptions(
  host: ParagraphFormatTarget,
  target: unknown,
  options: { keepNext: boolean; keepLines: boolean; widowControl: boolean },
): Promise<CommandOutcome> {
  const failedAction = 'שינוי אפשרויות השמירה נכשל';
  const doc = docOf(host);
  const setKeepOptions = doc?.format?.paragraph?.setKeepOptions;
  if (typeof setKeepOptions !== 'function') return unsupported(failedAction);
  return call(doc, failedAction, () =>
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

  const doc = docOf(host);
  const setTabStop = doc?.format?.paragraph?.setTabStop;
  if (typeof setTabStop !== 'function') return unsupported(failedAction);

  return call(doc, failedAction, () =>
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
  const doc = docOf(host);
  const clearTabStop = doc?.format?.paragraph?.clearTabStop;
  if (typeof clearTabStop !== 'function') return unsupported(failedAction);
  return call(doc, failedAction, () => clearTabStop({ target, position }));
}

export function clearAllParagraphTabStops(host: ParagraphFormatTarget, target: unknown): Promise<CommandOutcome> {
  const doc = docOf(host);
  const clearAll = doc?.format?.paragraph?.clearAllTabStops;
  if (typeof clearAll !== 'function') return Promise.resolve(unsupported('ניקוי עצירות הטאב נכשל'));
  return call(doc, 'ניקוי עצירות הטאב נכשל', () => clearAll({ target }));
}




