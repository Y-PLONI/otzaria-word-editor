/**
 * מיקום ¶ בסוף כל פסקה — לוגיקה טהורה, בלי DOM. אין כאן `Range`/`TreeWalker`
 * וגם לא `data-page-index`: כל מה שהקובץ הזה מקבל הוא כבר נתונים גולמיים —
 * `engine/page-ruler.ts` (`measureAllPageTextRuns`) מודד את ה-DOM, וכאן רק
 * החשבון. בדיוק אותה חלוקה כמו `engine/line-number-layer.ts` מול `page-ruler.ts`.
 *
 * ## הרקע: למה ¶ בלבד, ולא · (רווח) ו-→ (טאב)
 *
 * `docs/superdoc-2.10-review.md` ("סימני עיצוב (¶/·/→) — נחקר לעומק") מתעד
 * שני חסמים: (1) נתיב הציור המובנה של המנוע מת מבנית ב-2.10.0, ו-(2) טאבים
 * אינם קיימים כתו בשום Text node, וסמן רשימה מזריק טקסט שאינו קיים בטקסט
 * הקנוני של `blocks.list()` — שני חסמים שפוסלים מיקום מדויק של · ו-→. ¶
 * עוקף את שניהם כי הוא נוגע רק ב**סוף** הפסקה, ולעולם לא צריך לפתור מיקום
 * פנימי: הוא לא אכפת לו איפה בדיוק עובר הגבול בין שני טאבים רצופים, ואת
 * הרשימות הוא פשוט לא מנסה למקם (ראו "היקף" למטה) — בלי לגעת בבעיה שאין לה
 * פתרון.
 *
 * ## המסלול הממוקד
 *
 * formatting-marks-viewport מודד רק סופי פסקאות בעמודים הנראים וסביבתם.
 * painted-lines מאמת זהות של פסקת גוף וסיום fragment; הטקסט מגיע מבקשת
 * blocks.list מסוננת. buildMappedPilcrowMarks מתאים כל סוף למזהה שלו,
 * ובפסקה חוצה עמוד מאמת את סיומת הטקסט. אין תלות ברצף שמחוץ לחלון.
 * כשמיפוי הזהות חסר, שני הרצפים המלאים נשארים במסלול הישן להלן.
 *
 * ## ה-fallback: התאמת-רצף בין TreeWalker ל-blocks.list()
 *
 * שני רצפים, שניהם בסדר מסמך: רשימת ה-`FormattingMarksRun` (כל Text node
 * גלוי, `measureAllPageTextRuns`) ורשימת ה-`FormattingMarksBlock` (הטקסט
 * הקנוני של כל פסקה, `doc.blocks.list({includeText:true})`). לכל פסקה
 * מתאימים ("צורכים") כמה ריצות עוקבות עד שהטקסט המצורף שלהן שווה בדיוק
 * לטקסט הקנוני של הפסקה (אחרי הסרת טאבים — הם ממילא לא קיימים בצד ה-DOM).
 * ה-Text node **האחרון** שנצרך הוא העוגן ל-¶ שלה: `Text node` אינו חוצה גבול
 * פסקה בדפדפן (שני `<p>` נפרדים לא חולקים צומת טקסט), ולכן צריכת אורך מדויק
 * מגיעה תמיד **בדיוק** לסוף איזשהו node — לא באמצעו. זה מה שהופך את ¶
 * לפתיר בלי הבעיה שפוסלת · ו-→: אין כאן צורך לחתוך node באמצע.
 *
 * ## היקף: רק `nodeType === 'paragraph'`
 *
 * כותרות ופריטי-רשימה לא מטופלים בכוונה: פריט רשימה מזריק ל-DOM טקסט של
 * סמן+מפריד שאינו קיים בטקסט הקנוני שלו (נמדד — superdoc-2.10-review.md),
 * והתבנית שנמדדה (סמן ואז רווח בודד) אומתה על תצורה אחת בלבד. ניסיון למקם ¶
 * גם שם היה מחייב לנחש כמה תווים לדלג — בדיוק סוג הניחוש שהוחלט לא לעשות.
 * הבלוקים האלה פשוט מדולגים: `findMatch` לא מנסה ליישר אותם, אבל **גם לא
 * צורך את הריצות שלהם** — הן נשארות ברצף עבור `findMatch` הבא (ראו "סנכרון
 * מחדש"), כלומר בלוק שמדולג לא שובר את הבלוקים אחריו.
 *
 * ## פסקה ריקה
 *
 * `blocks.list()` מדווח `text:""` על פסקה ריקה (`isEmpty:true`), אבל נמדד
 * (superdoc-2.10-review.md) שהמנוע בפועל מצייר לה placeholder אמיתי — Text
 * node יחיד שמכיל **רווח בודד** (`" "`), עם מלבן ומיקום משלו. הטיפול: אם
 * הריצה הבאה שטרם נצרכה היא whitespace-בלבד, זו ה-placeholder — נצרכת
 * ונשמשת עוגן אמיתי (`approximate:false`). אם לא (המנוע השתנה, או שהתצורה
 * הזאת לא מצטיירת) — נפילה-לאחור גיאומטרית: שורה אחת מתחת לעוגן המוצלח
 * הקודם, מסומנת `approximate:true` כדי שהצרכן (הבדיקות, ה-QA) יידע שזו
 * הערכה ולא מדידה. בלי עוגן קודם בכלל (הבלוק הריק הוא הראשון במסמך) —
 * מדולג: פער קטן ומתועד, לא ניחוש.
 *
 * ## סנכרון מחדש
 *
 * `findMatch` לא רק בודק אם ההתאמה מתחילה בדיוק בסמן הנוכחי — הוא מנסה כל
 * היסט התחלה אפשרי בחלון חסום (`RESYNC_LOOKAHEAD_RUNS`). זה מה שמאפשר
 * לבלוק שדולג (כותרת, פריט רשימה) או לכל תוכן לא-צפוי אחר (כותרת עליונה של
 * עמוד הבא, למשל) להיבלע בלי לשבור את הפסקאות שאחריו: הפסקה הבאה פשוט
 * נמצאת קצת יותר רחוק ברצף. אם לא נמצאה התאמה בכלל בחלון — הבלוק מדולג
 * בלי סימון, לא מנוחש.
 */
import type { RawTextRect } from './page-ruler';

/** מה שנדרש מכל בלוק שנקרא מ-`doc.blocks.list({includeText:true})`. */
export interface FormattingMarksBlock {
  nodeId: string;
  nodeType: string;
  /** הטקסט הקנוני — כולל טאבים (`\t`), שאינם קיימים בצד ה-DOM. */
  text: string;
  isEmpty: boolean;
}

/** ריצת טקסט אחת, כפי ש-`measureAllPageTextRuns` (page-ruler.ts) מודדת. */
export interface FormattingMarksRun {
  text: string;
  rects: readonly RawTextRect[];
  direction: 'ltr' | 'rtl';
}

/** A final painted fragment, validated against the matching Document API ID. */
export interface FormattingMarksEnd {
  nodeId: string;
  continuesFromPrevious: boolean;
  runs: readonly FormattingMarksRun[];
}

/** מיקום ¶ אחד לציור. */
export interface PilcrowMark {
  nodeId: string;
  topPx: number;
  heightPx: number;
  /** הקצה שממנו "מתחיל" ה-¶ בכיוון הקריאה — ראו `direction`. */
  anchorXPx: number;
  direction: 'ltr' | 'rtl';
  /** `true` = נפילה-לאחור גיאומטרית (פסקה ריקה בלי placeholder מזוהה). */
  approximate: boolean;
}

/** רק בלוקים מהסוג הזה מסומנים — ראו הערת הפתיחה, "היקף". */
const ELIGIBLE_NODE_TYPE = 'paragraph';

/** כמה ריצות קדימה מותר לחפש נקודת-סנכרון. ראו הערת הפתיחה, "סנכרון מחדש". */
export const RESYNC_LOOKAHEAD_RUNS = 60;

/** טאבים אינם קיימים כתו בשום Text node גלוי (נמדד) — מוסרים לפני השוואה. */
function canonicalTarget(text: string): string {
  return text.replace(/\t/g, '');
}

function isWhitespaceOnly(text: string): boolean {
  return text.length > 0 && /^\s+$/.test(text);
}

interface ConsumeResult {
  matched: boolean;
  /** האינדקס שממנו ממשיכים אחרי ההתאמה (בלעדי). */
  endIndex: number;
  lastRun: FormattingMarksRun | null;
}

/** צורכת ריצות עוקבות מ-`fromIndex` עד שאורך הטקסט המצורף שווה ל-`target`. */
function tryConsume(
  runs: readonly FormattingMarksRun[],
  fromIndex: number,
  target: string,
): ConsumeResult {
  let consumed = '';
  let idx = fromIndex;
  while (consumed.length < target.length && idx < runs.length) {
    consumed += runs[idx]!.text;
    idx += 1;
  }
  if (idx > fromIndex && consumed === target) {
    return { matched: true, endIndex: idx, lastRun: runs[idx - 1]! };
  }
  return { matched: false, endIndex: fromIndex, lastRun: null };
}

/**
 * מוצאת מאיפה בדיוק להתחיל לצרוך כדי להגיע ל-`target` בדיוק — מנסה כל היסט
 * אפשרי בחלון חסום, לא רק את הסמן הנוכחי. ראו הערת הפתיחה, "סנכרון מחדש".
 */
function findMatch(
  runs: readonly FormattingMarksRun[],
  fromIndex: number,
  target: string,
): (ConsumeResult & { startIndex: number }) | null {
  const limit = Math.min(runs.length, fromIndex + RESYNC_LOOKAHEAD_RUNS);
  for (let start = fromIndex; start <= limit; start++) {
    const result = tryConsume(runs, start, target);
    if (result.matched) return { ...result, startIndex: start };
  }
  return null;
}

/** עוגן ה-¶ בפועל, מתוך הריצה האחרונה שנצרכה — השורה/מלבן האחרון שלה. */
function anchorFromRun(run: FormattingMarksRun, approximate: boolean): Omit<PilcrowMark, 'nodeId'> | null {
  const rect = run.rects[run.rects.length - 1];
  if (!rect) return null;
  return {
    topPx: rect.topPx,
    heightPx: rect.heightPx,
    anchorXPx: run.direction === 'rtl' ? rect.leftPx : rect.leftPx + rect.widthPx,
    direction: run.direction,
    approximate,
  };
}

/**
 * מרכיבה את כל מיקומי ה-¶ לציור, מהבלוקים (Document API) ומריצות הטקסט
 * (DOM, `measureAllPageTextRuns`). קלט ריק בכל צד → פלט ריק, לא שגיאה.
 */
export function buildPilcrowMarks(
  blocks: readonly FormattingMarksBlock[],
  runs: readonly FormattingMarksRun[],
): PilcrowMark[] {
  const marks: PilcrowMark[] = [];
  let cursor = 0;
  let lastAnchor: Omit<PilcrowMark, 'nodeId'> | null = null;

  for (const block of blocks) {
    if (block.nodeType !== ELIGIBLE_NODE_TYPE) continue;

    const target = canonicalTarget(block.text);

    if (target.length === 0) {
      // פסקה ריקה — ראו הערת הפתיחה, "פסקה ריקה".
      const next = runs[cursor];
      if (next && isWhitespaceOnly(next.text)) {
        const anchor = anchorFromRun(next, false);
        cursor += 1;
        if (anchor) {
          marks.push({ nodeId: block.nodeId, ...anchor });
          lastAnchor = anchor;
        }
        continue;
      }
      if (lastAnchor) {
        const anchor: Omit<PilcrowMark, 'nodeId'> = {
          ...lastAnchor,
          topPx: lastAnchor.topPx + lastAnchor.heightPx,
          approximate: true,
        };
        marks.push({ nodeId: block.nodeId, ...anchor });
        lastAnchor = anchor;
      }
      continue; // לא נצרכה אף ריצה — לא היה מה לצרוך.
    }

    const found = findMatch(runs, cursor, target);
    if (!found) continue; // לא נמצאה התאמה בחלון — מדולג, לא מנוחש.

    cursor = found.endIndex;
    const anchor = found.lastRun ? anchorFromRun(found.lastRun, false) : null;
    if (anchor) {
      marks.push({ nodeId: block.nodeId, ...anchor });
      lastAnchor = anchor;
    }
  }

  return marks;
}

/** Match each paragraph to its own ending; identical text cannot shift identity. */
export function buildMappedPilcrowMarks(
  blocks: readonly FormattingMarksBlock[],
  endings: readonly FormattingMarksEnd[],
): PilcrowMark[] {
  const byId = new Map(blocks.map((block) => [block.nodeId, block]));
  const counts = new Map<string, number>();
  for (const end of endings) counts.set(end.nodeId, (counts.get(end.nodeId) ?? 0) + 1);
  const marks: PilcrowMark[] = [];
  for (const end of endings) {
    const block = byId.get(end.nodeId);
    if (!block || block.nodeType !== ELIGIBLE_NODE_TYPE || counts.get(end.nodeId) !== 1) continue;
    const target = canonicalTarget(block.text);
    const painted = end.runs.map((run) => run.text).join('');
    if (target.length === 0) {
      if (!isWhitespaceOnly(painted)) continue;
    } else if (!painted || (end.continuesFromPrevious ? !target.endsWith(painted) : target !== painted)) {
      continue;
    }
    const last = end.runs[end.runs.length - 1];
    const anchor = last ? anchorFromRun(last, false) : null;
    if (anchor) marks.push({ nodeId: block.nodeId, ...anchor });
  }
  return marks;
}
