/**
 * רשימות — המספור העברי, התחלה מחדש, המשך מספור קודם והמרה לטקסט
 * (גל 14א), דרך `doc.lists.*`.
 *
 * ## ממצא הדגל שנמדד: `hebrew1` עובד
 *
 * `numFmt` בחוזה הוא **string חופשי** ולא union (שונה מ-
 * `sections.setPageNumbering.format`, שם ה-union נאכף בזמן ריצה). נמדד:
 * `numFmt:'hebrew1'` על רשימה קיימת כתב `<w:numFmt w:val="hebrew1"/>`
 * ב-numbering.xml — מספור א׳ ב׳ ג׳ אמיתי.
 * ולכן המודול מציע אותו, וחוסם ערכים שאינם ברשימת `numFmt` של ECMA-376.
 *
 * ## שאר מה שנמדד
 *
 * - `restartAt({startAt})` עובד; `continuePrevious` מחזיר קבלת כשל
 *   (`INVALID_CONTEXT / NO_PREVIOUS_LIST`) ולא זורק; `canContinuePrevious`
 *   הוא בוליאן ולכן TOCTOU — קוראים לפעולה ומדווחים את הקבלה, בלי להסתמך
 *   על הבוליאן.
 * - `convertToText({includeMarker:true})` **מעתיק את הסמן לתוך הטקסט**
 *   ('a. ') והפריט הופך לפסקה — בלתי-הפיך למעשה, ולכן הפקד דורש אישור
 *   דו-לחיצה בממשק.
 * - כתובת פריט היא `{kind:'block', nodeType:'listItem', nodeId}`; פסקה
 *   שאינה פריט מחזירה `TARGET_NOT_FOUND`. היעד נפתר מהבחירה + `lists.getState`
 *   (ראו `resolveListItem`: `blocks.list` לבדו מפספס רשימות בטבלה וכותרות
 *   ממוספרות).
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt } from './document-api';

const UNAVAILABLE_TEXT = 'אינו זמין בגרסה זו';

/**
 * ערכי `numFmt` של ECMA-376 בהם נעשה שימוש הגיוני במסמך עברי/כללי.
 *
 * `hebrew1` היה כאן מלכתחילה — החוזה מקבל `numberStyle` כמחרוזת חופשית,
 * ונמדד ש-`w:numFmt="hebrew1"` אכן נכתב ל-numbering.xml. מה שלא נמדד אז הוא
 * שה**סמן צויר ריק**: על superdoc@2.8.0 המשתמש קיבל „. ” בלי אות, כלומר
 * מסמך נכון ומסך ריק. במעבר ל-2.10.0 הסמנים מצוירים, ולכן `hebrew2` מצטרף.
 *
 * `bullet` אינו סגנון מספור: כ-`numFmt` בלי `lvlText` הסמן שהצטייר היה „%1.”.
 */
export const NUMBER_STYLES: readonly string[] = [
  'decimal',
  'upperLetter',
  'lowerLetter',
  'hebrew1',
  'hebrew2',
];

/**
 * תוויות לתצוגה.
 *
 * התווית של `hebrew1` הייתה „א׳, ב׳, ג׳ — עברי” והיא תוקנה: המנוע אינו מצייר
 * גרש. עשרים פריטים ברצף נמדדו על ה-dist הבנוי, וזה מה שיצא:
 *
 *     hebrew1 → א ב ג … י יא יב יג יד טו טז יז יח יט כ   (גימטריה)
 *     hebrew2 → א ב ג … י כ  ל  מ  נ  ס  ע  פ  צ  ק  ר   (סדר האלף-בית)
 *
 * שניהם זהים בעשרת הראשונים, ולכן התוויות מראות היכן הם נפרדים.
 */
export const NUMBER_STYLE_LABELS: Record<string, string> = {
  decimal: '1, 2, 3',
  upperLetter: 'A, B, C',
  lowerLetter: 'a, b, c',
  hebrew1: 'א, ב, ג … יא, יב (גימטריה)',
  hebrew2: 'א, ב, ג … כ, ל (אלף־בית)',
};

interface ListsApiShape {
  selection?: {
    current?: () => MaybePromise<SelectionInfoLike | undefined>;
  };
  blocks?: {
    list?: () => MaybePromise<{
      blocks?: Array<{ nodeId?: string; nodeType?: string; paragraphNumbering?: unknown }>;
    }>;
  };
  lists?: {
    getState?: (input: {
      target: { kind: 'block'; nodeType: 'paragraph' | 'listItem'; nodeId: string };
    }) => MaybePromise<{ success?: boolean; isListItem?: boolean } | undefined>;
    getStyle?: (input: {
      target: { kind: 'block'; nodeType: 'listItem'; nodeId: string };
    }) => MaybePromise<{ success?: boolean; style?: { levels?: ListLevelStyleLike[] } } | undefined>;
    applyStyle?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
    restartAt?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
    continuePrevious?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
    convertToText?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
  };
}

export interface ListsHost {
  activeEditor?: { doc?: ListsApiShape | null } | null;
}

/** רמה אחת כפי ש-`lists.getStyle` מחזירה אותה — רק מה שנקרא כאן. */
interface ListLevelStyleLike {
  level?: number;
  numFmt?: string;
  lvlText?: string;
  markerFont?: string;
}

export type ListsTarget = SuperDoc | ListsHost | null | undefined;

interface SelectionInfoLike {
  empty?: boolean;
  target?: {
    segments?: ReadonlyArray<{ blockId?: string }>;
  } | null;
}

type MaybePromise<T> = T | Promise<T>;

function docOf(host: ListsTarget): ListsApiShape | null {
  return (host as ListsHost | null | undefined)?.activeEditor?.doc ?? null;
}

/**
 * פותרת את הבלוק שבו הסמן לאחד משלושה מצבים (ראו `ListItemResolution`).
 * `blockId` מהבחירה (הבחירה אינה מדווחת listItem), וההכרעה „פריט רשימה או
 * לא” היא של `lists.getState` — סמכות הרשימות של המנוע — ולא של `blocks.list`.
 *
 * ההפרדה בין „אינו רשימה” ל„לא ידוע” קיימת בשביל `createList` בלבד, והיא
 * הכרחית: הפקודה ההיא טוגל, ועל „לא ידוע” היא הייתה מסירה מספור מפריט קיים.
 *
 * למה לא `blocks.list` (issue #14 ג׳, נמדד על superdoc 2.11.0): הוא מונה
 * בלוקים **עליונים** בלבד, ולכן פריט רשימה בתוך תא טבלה אינו מופיע בו כלל —
 * וזה `unknown`, לא „אין רשימה”; וכותרת ממוספרת של Word (Heading1 + numPr)
 * מדווחת בו כ-`heading` ולא `listItem`. בשני המקרים `getState` אומר
 * `isListItem:true`, ו-`setLevelNumberStyle` עם כתובת `listItem` מצליח —
 * כלומר הבלוק כן פריט רשימה, ורק הזיהוי כשל ב„יש למקם את הסמן בתוך רשימה”.
 *
 * `blocks.list` נשאר כנפילה לגרסת מנוע בלי `getState`, בקריאה אחת: בלי
 * ארגומנטים הוא מחזיר את כל הסיפור (אין עמוד של 50, נמדד), ו-
 * `paragraphNumbering` על הבלוק מכסה שם את הכותרת הממוספרת.
 */
async function resolveListItem(host: ListsTarget): Promise<ListItemResolution> {
  const doc = docOf(host);
  if (!doc) return { kind: 'unknown' };

  let blockId: string | null = null;
  try {
    const info = await doc.selection?.current?.();
    blockId =
      info?.target?.segments?.find((s) => typeof s?.blockId === 'string')?.blockId ?? null;
  } catch {
    return { kind: 'unknown' };
  }
  if (!blockId) return { kind: 'unknown' };
  return resolveListItemAt(host, blockId);
}

/**
 * אותה הכרעה, על בלוק שהקורא כבר מחזיק את המזהה שלו.
 *
 * מיוצאת בשביל `list-autoformat-install.ts`: זיהוי הרשימות בהקלדה כבר קרא את
 * הסמן בעצמו (הוא צריך גם את ההיסט), וקריאה שנייה ל-`selection.current` שם
 * הייתה גם מיותרת וגם מרוץ — בין שתי הקריאות המשתמש ממשיך להקליד. חשוב מכך:
 * `numbered-list` הוא **טוגל**, ושם בדיוק נדרשת ההפרדה בין „אינו רשימה”
 * להכרעה שלא התקבלה.
 */
export async function resolveListItemAt(
  host: ListsTarget,
  blockId: string,
): Promise<ListItemResolution> {
  const doc = docOf(host);
  if (!doc) return { kind: 'unknown' };
  const address: ListItemAddress = { kind: 'block', nodeType: 'listItem', nodeId: blockId };

  const getState = doc.lists?.getState;
  if (typeof getState === 'function') {
    try {
      // `nodeType:'paragraph'` תקף לכל בלוק פסקתי, כולל כותרת; `'heading'` נזרק.
      const state = await getState({ target: { kind: 'block', nodeType: 'paragraph', nodeId: blockId } });
      // `=== false` ולא truthiness: `success:true` בלי השדה אינו הכרעה, והוא
      // אסור להפעלת טוגל.
      if (state?.success === true && state.isListItem === true) return { kind: 'item', address };
      if (state?.success === true && state.isListItem === false) return { kind: 'not-list' };
    } catch {
      // נופלים ל-blocks.list.
    }
  }

  const list = doc.blocks?.list;
  if (typeof list !== 'function') return { kind: 'unknown' };
  try {
    const listed = await list();
    const block = (listed?.blocks ?? []).find((b) => b.nodeId === blockId);
    if (!block) return { kind: 'unknown' };
    return block.nodeType === 'listItem' || block.paragraphNumbering
      ? { kind: 'item', address }
      : { kind: 'not-list' };
  } catch {
    return { kind: 'unknown' };
  }
}

/** הכתובת לבדה, לקוראים שמתייחסים ל„לא רשימה” ול„לא ידוע” אותו דבר. */
async function resolveAddress(host: ListsTarget): Promise<ListItemAddress | null> {
  const resolution = await resolveListItem(host);
  return resolution.kind === 'item' ? resolution.address : null;
}

export interface ListItemAddress {
  kind: 'block';
  nodeType: 'listItem';
  nodeId: string;
}

/**
 * `not-list` הוא הכרעה חיובית — `getState` ענה `isListItem:false`, או שהבלוק
 * נמנה ב-`blocks.list` ואינו פריט. כל השאר `unknown`: אין `doc`, הבחירה
 * זרקה, אין `blockId`, `getState` חסרה/זרקה/לא הכריעה, `blocks.list`
 * חסרה/זרקה, או שאינה מונה את הבלוק כלל (פריט בתא טבלה).
 */
export type ListItemResolution =
  | { kind: 'item'; address: ListItemAddress }
  | { kind: 'not-list' }
  | { kind: 'unknown' };

function unsupported(failedAction: string): CommandOutcome {
  return { ok: false, message: `${failedAction}: ${UNAVAILABLE_TEXT}`, reason: 'command-unsupported' };
}

/** עוטף קריאה אחת: לעולם לא זורק, NO_OP הצלחה. */
async function call(failedAction: string, run: () => MaybePromise<DocReceipt>): Promise<CommandOutcome> {
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

/**
 * מוסיפה את קידומת הכשל של המודול להודעה שכבר מנוסחת (כשל שחוזר משכבה
 * אחרת). ה-`startsWith` מונע קידומת כפולה אם יזרימו לכאן הודעה שכבר נושאת
 * אותה.
 */
function prefixed(failedAction: string, message: string): string {
  if (!message) return failedAction;
  return message.startsWith(failedAction) ? message : `${failedAction}: ${message}`;
}

/** „יש למקם את הסמן ברשימה" — תשובה משותפת לכל הפעולות. */
function notInList(failedAction: string): CommandOutcome {
  return { ok: false, message: `${failedAction}: יש למקם את הסמן בתוך רשימה`, reason: 'selection-required' };
}

export interface SetNumberStyleOptions {
  /**
   * איך להפוך את הפסקה שבסמן לרשימה כשהיא עדיין אינה כזאת — הפקודה של
   * הכפתור „מספור”, שיושבת ב-registry של המנוע ולא ב-Document API, ולכן
   * מוזרקת מהרצועה ולא נקראת מכאן.
   *
   * בלעדיה הבחירה בסגנון מסרבת על פסקה רגילה. זה מה שהמשתמש קיבל (issue
   * #14 ג׳, שוחזר ב-scripts/qa/list-caret-qa.mjs): פתח את תפריט המספור על
   * פסקה, בחר „א, ב, ג”, וקרא „יש למקם את הסמן בתוך רשימה” — כשכל מה שביקש
   * הוא למספר את הפסקה, כמו שגלריית המספור של Word עושה.
   *
   * הפקודה היא **טוגל**: על בלוק שכבר רשימה מאותו סוג היא מסירה את המספור
   * (`lists.remove`, ובכותרת/תחתית `lists.removeInStory`). לכן היא נקראת רק
   * על הכרעה חיובית „אינו רשימה” — ולא כשהזיהוי החזיר „לא ידוע” (פריט
   * בתא טבלה, מנוע בלי `getState`), שם טוגל היה הורס את המספור הקיים.
   *
   * `useCommand.run` מדווח את הכשל בעצמו לפני שהוא מחזיר אותו, והקורא מדווח
   * שוב את מה שחזר מכאן — ולכן הכתיבה השנייה לפס המצב, זו שנושאת את שם
   * הפעולה, היא זו שהמשתמש רואה.
   */
  createList?: () => Promise<CommandOutcome>;
}

/** ה-`levels` שנכתבים ל-`applyStyle`: רק שדות הסמן, ולא ההזחות. */
interface PlannedLevel {
  level: number;
  numFmt: string;
  lvlText: string;
  markerFont?: string;
}

/** תשע הרמות של הגדרת רשימה ב-ECMA-376. */
const LEVEL_COUNT = 9;

/**
 * הרמות שסגנון המספור נכתב אליהן: רמה 0 היא הבקשה, ורמה עמוקה רק כשהיא
 * תבליט — שם אין פקד אחר שיגיע אליה, וקסקדה ממוספרת שהמשתמש קבע אינה נדרסת.
 */
function plannedLevels(numberStyle: string, current: ListLevelStyleLike[] | null): PlannedLevel[] {
  const at = (level: number) => current?.find((l) => l?.level === level);
  const levels: PlannedLevel[] = [];
  for (let level = 0; level < LEVEL_COUNT; level += 1) {
    if (level > 0 && at(level)?.numFmt !== 'bullet') continue;
    levels.push({
      level,
      numFmt: numberStyle,
      // `%N.` ולא `%1.`: הסמן מפנה למונה של הרמה שלו.
      lvlText: `%${level + 1}.`,
      // `''` הוא הריקון היחיד שהחוזה מקבל — `null` נדחה („must be a string”).
      ...(current === null || at(level)?.markerFont ? { markerFont: '' } : {}),
    });
  }
  return levels;
}

/** `lists.getStyle` כמדידה ולא כתלות: גרסה בלעדיה מחזירה `null`. */
async function readLevels(
  lists: ListsApiShape['lists'],
  target: ListItemAddress,
): Promise<ListLevelStyleLike[] | null> {
  const getStyle = lists?.getStyle;
  if (typeof getStyle !== 'function') return null;
  try {
    const result = await getStyle({ target });
    const levels = result?.success === true ? result.style?.levels : undefined;
    return Array.isArray(levels) ? levels : null;
  } catch {
    return null;
  }
}

/**
 * הרמות שנכתבו וחזרו אחרות. `numFmt` בלי `lvlText` הוא הכשל השקט שהיה כאן:
 * המסמך אמר `decimal`, המסך המשיך לצייר „•”, ופס המצב שתק.
 */
function unwrittenLevels(wanted: PlannedLevel[], after: ListLevelStyleLike[]): number[] {
  return wanted
    .filter((want) => {
      const got = after.find((l) => l?.level === want.level);
      return got !== undefined && (got.numFmt !== want.numFmt || got.lvlText !== want.lvlText);
    })
    .map((want) => want.level);
}

/**
 * מגדירה את סגנון המספור ברשימה שבה הסמן. `hebrew1` הוא המספור העברי
 * (א׳ ב׳ ג׳) — נמדד שנכתב ל-numbering.xml.
 *
 * `applyStyle` ולא `setLevelNumberStyle`: זה כותב `numFmt` לבדו, ואחרי המרה
 * לתבליטים ה-`lvlText` נשאר „•” — סגנון שנבחר, מסמך שאומר `decimal`, ומסך
 * שממשיך לצייר תבליט. שניהם sequence-local בחוזה, ולכן הבידוד נשמר.
 *
 * על פסקה שאינה רשימה: אם נמסרה `createList`, קודם יוצרים את הרשימה ואז
 * מחילים את הסגנון — בחירת סגנון היא בקשה למספר, לא שאלה על רשימה קיימת.
 * בלי `createList` (או כשהיצירה נכשלה) — הסירוב המפורש, כמו קודם. כשהזיהוי
 * לא הכריע (`unknown`) גם הסירוב, בלי לגעת במסמך: `createList` היא טוגל.
 */
export async function setListNumberStyle(
  host: ListsTarget,
  numberStyle: string,
  options: SetNumberStyleOptions = {},
): Promise<CommandOutcome> {
  const failedAction = 'שינוי סגנון המספור נכשל';

  if (!NUMBER_STYLES.includes(numberStyle)) {
    // string חופשי בחוזה — המנוע כנראה בולע כל ערך; רק numFmt תקני יוצא.
    return { ok: false, message: `${failedAction}: סגנון המספור אינו חוקי`, reason: 'invalid-number-style' };
  }

  const resolution = await resolveListItem(host);
  let item = resolution.kind === 'item' ? resolution.address : null;
  if (resolution.kind === 'not-list' && options.createList) {
    const created = await options.createList();
    // הכשל חוזר עם הקידומת של המודול: ההודעה של `createList` היא של שכבת
    // הפקודות („המנוע אינו מוכן”) ובלי שם הפעולה המשתמש אינו יודע מה נכשל.
    // ה-`reason` נשמר — הוא מה שמבדיל בין הכשלים אצל הצרכן.
    if (!created.ok) return { ...created, message: prefixed(failedAction, created.message) };
    // הכתובת יציבה (נמדד: אותו blockId לפני היצירה ואחריה); הפתרון מחדש הוא
    // אימות שהיצירה אכן תפסה, ולא הסתמכות על יציבות שאינה בחוזה.
    item = await resolveAddress(host);
  }
  if (!item) return notInList(failedAction);

  const lists = docOf(host)?.lists;
  const applyStyle = lists?.applyStyle;
  if (typeof applyStyle !== 'function') return unsupported(failedAction);

  const levels = plannedLevels(numberStyle, await readLevels(lists, item));
  const applied = await call(failedAction, () => applyStyle({ target: item, style: { version: 1, levels } }));
  if (!applied.ok) return applied;

  // הכתיבה נבדקת אחריה: „הצליח” בקבלה אינו „הסמן התחלף”.
  const after = await readLevels(lists, item);
  const unwritten = after ? unwrittenLevels(levels, after) : [];
  if (unwritten.length > 0) {
    return {
      ok: false,
      message: `${failedAction}: הסגנון לא נכתב לרמות ${unwritten.join(', ')}`,
      reason: 'level-not-written',
    };
  }
  return { ok: true };
}

/** רק שדות הסמן, ובכל תשע הרמות — שדה שאינו נמסר נשאר כשהיה, כולל ההזחות. */
const BULLET_STYLE = {
  version: 1,
  levels: Array.from({ length: LEVEL_COUNT }, (_, level) => ({
    level,
    numFmt: 'bullet',
    lvlText: '•',
    markerFont: 'Symbol',
  })),
};

/**
 * `applyStyle` ולא `setType`: `setType` כותב ל-`abstractNum` המשותף ומהפך גם
 * רשימה שכנה; `applyStyle` הוא sequence-local — הגדרה משותפת משובטת.
 */
export async function setListToBullets(host: ListsTarget): Promise<CommandOutcome> {
  const failedAction = 'המרת הרשימה לתבליטים נכשלה';

  const item = await resolveAddress(host);
  if (!item) return notInList(failedAction);

  const applyStyle = docOf(host)?.lists?.applyStyle;
  if (typeof applyStyle !== 'function') return unsupported(failedAction);

  return call(failedAction, () => applyStyle({ target: item, style: BULLET_STYLE }));
}

/** „התחל מחדש": מגדיר את ערך ההתחלה של הרשימה שבה הסמן. */
export async function restartListAt(host: ListsTarget, startAt: number): Promise<CommandOutcome> {
  const failedAction = 'התחלה מחדש של הרשימה נכשלה';

  if (typeof startAt !== 'number' || !Number.isInteger(startAt) || startAt < 0) {
    return { ok: false, message: `${failedAction}: הערך חייב להיות מספר שלם לא-שלילי`, reason: 'invalid-start' };
  }

  const item = await resolveAddress(host);
  if (!item) return notInList(failedAction);

  const restartAt = docOf(host)?.lists?.restartAt;
  if (typeof restartAt !== 'function') return unsupported(failedAction);

  return call(failedAction, () => restartAt({ target: item, startAt }));
}

/**
 * „המשך מספור קודם". כשאין רשימה קודמת המנוע מחזיר קבלת כשל
 * (`NO_PREVIOUS_LIST`) וההודעה מתורגמת — הבוליאן canContinue אינו בשימוש
 * (TOCTOU).
 */
export async function continuePreviousList(host: ListsTarget): Promise<CommandOutcome> {
  const failedAction = 'המשך המספור מהרשימה הקודמת נכשל';

  const item = await resolveAddress(host);
  if (!item) return notInList(failedAction);

  const continuePrevious = docOf(host)?.lists?.continuePrevious;
  if (typeof continuePrevious !== 'function') return unsupported(failedAction);

  return call(failedAction, () => continuePrevious({ target: item }));
}

/**
 * „המר לטקסט" — בלתי-הפיך למעשה: סמן הרשימה מועתק לתוך הטקסט ('a. ')
 * והפריט הופך לפסקה (נמדד). הפקד חייב אישור דו-לחיצה לפני הקריאה.
 */
export async function convertListToText(host: ListsTarget): Promise<CommandOutcome> {
  const failedAction = 'המרת הרשימה לטקסט נכשלה';

  const item = await resolveAddress(host);
  if (!item) return notInList(failedAction);

  const convertToText = docOf(host)?.lists?.convertToText;
  if (typeof convertToText !== 'function') return unsupported(failedAction);

  return call(failedAction, () => convertToText({ target: item, includeMarker: true }));
}


