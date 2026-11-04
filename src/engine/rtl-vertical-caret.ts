/**
 * חצים אנכיים בשורה עברית — „יורד לשורה קצרה יותר והסמן קופץ לתחילתה”.
 *
 * ## מה שבור במנוע
 *
 * בפסקה עברית, כשעמודת המטרה נופלת **מחוץ** לשורה שאליה יורדים (או עולים),
 * הסמן נוחת ב**קצה ההפוך** שלה. שתי החציים הם תמונת ראי מדויקת:
 *
 *   | עמודת המטרה ביחס לשורת היעד | היכן הסמן נחת | היכן הוא צריך |
 *   |-----------------------------|---------------|---------------|
 *   | מעבר ל**סוף** השורה         | **תחילתה**    | סופה          |
 *   | לפני **תחילת** השורה        | **סופה**      | תחילתה        |
 *   | בתוך השורה                  | נכון          | נכון          |
 *
 * הסימטריה הזאת היא מה שהופך את זה לתקלה **אחת** ולא לשתיים: עמודת המטרה
 * מתורגמת להיסט בתוך שורת היעד כאילו השורה רצה משמאל לימין, ואז מקוצצת אל
 * השורה. בשורה לטינית שני קצוות הקיצוץ יוצאים נכון; בשורה עברית שניהם יוצאים
 * הפוכים.
 *
 * נמדד ב-`scripts/qa/rtl-vertical-caret-probe.mjs` (superdoc 2.15.0, Chrome,
 * על ה-dist הארוז): עשרה מקרים, ולכל מקרה עברי יש בקרה לועזית באותו מסמך —
 * פסקה גולשת, פסקה ארוכה מול קצרה, חץ למעלה, שורה מוזחת ושורה ממורכזת. **כל**
 * חמש הבקרות הלועזיות עברו, וכל חמשת המקרים העבריים נכשלו. הפסקה הממורכזת היא
 * ההוכחה שזו תמונת ראי: באותה שורה נמדדו שני הצדדים, ושניהם הפוכים.
 *
 * דווח למעלה כ-[#4018](https://github.com/superdoc/docx-editor/issues/4018).
 * הקוד שמזיז את הסמן אינו בחלקים הפתוחים של SuperDoc, ולכן אין PR — רק יירוט.
 * כשה-issue ייסגר, למחוק את המודול הזה ואת השער שלו.
 *
 * ## למה זה לא „רק הסמן במקום הלא נכון”
 *
 * מי שיורד בחץ לשורה שמתחת וממשיך להקליד, מקליד ב**ראש** הפסקה — ובפסקה
 * הראשונה זו תחילת המסמך. הנזק נוחת בקובץ שנשמר, בדיוק כמו במקש `End`
 * ([#3986](https://github.com/superdoc/docx-editor/issues/3986)), ולא רק על
 * המסך.
 *
 * ## שלוש מדידות שקבעו את המדיניות כאן
 *
 * העקיפה נראית במבט ראשון כמו „תקן את הנחיתה”. היא אינה, ושלוש מדידות הן
 * הסיבה:
 *
 * **1. למנוע יש עמודת מטרה, והיא שורדת שורה קצרה.** ארוכה → קצרה → ארוכה
 * החזיר את הסמן לעמודה המקורית בשורה השלישית, גם בעברית וגם בלועזית. כלומר
 * המנוע עושה כאן את הדבר הנכון, וזה מה שצריך לשמר.
 *
 * **2. כתיבה דרך `authoring.setSelectionTarget` דורסת את העמודה הזאת.** נמדד
 * ב-`scripts/qa/rtl-vertical-seam-probe.mjs`: אחרי כתיבת בחירה אל סוף השורה
 * הקצרה, ה-`ArrowDown` הבא המשיך מהמקום **שנכתב**, ולא מהעמודה המקורית.
 *
 * שתי אלה יחד קובעות את הצורה: עקיפה שרק „מתקנת את הנחיתה” הייתה מתקנת הקשה
 * אחת ושוברת את הבאה אחריה. לכן המודול הזה מחזיק **עמודת מטרה משלו**, ומרגע
 * שהתערב פעם אחת הוא ממשיך לטפל בהקשות האנכיות עד שהסמן זז ממקום אחר. כל עוד
 * לא התערב — העמודה שלו והעמודה של המנוע זהות, שתיהן נזרעו מאותו סמן מצויר,
 * והוא מוסר לו כל הקשה שהוא צודק בה.
 *
 * ‏„זז ממקום אחר” נבדק בהשוואה למה שכתבנו, ולא בהאזנה למקשים — ההסבר ב-
 * `lastWrite`. זה מה שהופך את המודול לבלתי תלוי ב**סדר** שבו שלושת מאזיני
 * הסמן מותקנים ב-App.vue.
 *
 * **3. תפר גלישה שייך לשורה הקודמת.** ההיסט שבין שתי שורות גולשות הוא
 * `data-pm-end` של הראשונה וגם `data-pm-start` של השנייה, וכתיבתו מציירת את
 * הסמן ב**סוף הראשונה**. כלומר „רד לתחילת שורת המשך” אינו ניתן לביטוי דרך
 * ה-API, ומי שמכוון לשם מקבל הקשה שנראית כאילו לא עשתה דבר. ההסבר המלא
 * ב-`startIsSeam`, והטיפול ב-`landingSlot`.
 *
 * ## מה נשאר למנוע, ובמודע
 *
 * * **`Shift+חץ`** — התצלום הסינכרוני מחזיר `selectionTarget: null` לכל בחירה
 *   שאינה מכווצת. בחצים האופקיים זה נפתר בזיכרון של ראש הבחירה („בחירה” ב-
 *   `rtl-caret.ts`); כאן זה טרם נעשה, וההקשה מכל מקום מסיימת את אותו זיכרון.
 * * **`Ctrl/Alt/Meta+חץ`** ו-`PageUp`/`PageDown` — תנועה אחרת, מדידה אחרת.
 * * **שורה לטינית**, כל עוד לא התערבנו — שם המנוע נמדד תקין בכל חמש הבקרות.
 * * **טבלה שכנה, כותרת עליונה, בלוק שאינו צמוד במסמך, עמוד שלא צויר** —
 *   `nextVisualLine` מחזיר `null`, וההקשה עוברת כפי שהיא.
 * * **שורת יעד ריקה** — יש בה מקום סמן אחד בלבד, והמנוע נוחת עליו בכל עמודה.
 *
 * ובכל אחד מאלה ההקשה נמסרת, וה**עמודה נשמרת**. זה נראה כמו פרט ואינו: אילו
 * מסירה הייתה מאפסת, הרצף „פסקה ארוכה → פסקה ריקה → פסקה קצרה” — צירוף שגרתי
 * במסמך עברי — היה נשבר. המסירה הראשונה (השורה הריקה) הייתה מאפסת, ההקשה
 * הבאה הייתה נזרעת מהסמן שעל הפסקה הריקה — קצה ההתחלה — ומכריזה בטעות
 * „העמודה בתוך השורה הקצרה” ומוסרת למנוע, שעמודתו **לא** נדרסה ולכן נחת
 * בקצה ההפוך. מסירה אינה משנה את עמודת המנוע, ולכן אסור לה לשנות גם את שלנו.
 */

import {
  CARET_SELECTOR,
  blockIdOf,
  blockStart,
  caretSlots,
  contiguousWith,
  findFragment,
  isEmptyRange,
  lineAt,
  lineExtent,
  linesOf,
  neighbourOf,
  readLineChars,
  startIsSeam,
  type CaretSlot,
} from './painted-lines';

/* ------------------------------------------------------------------ */
/* מה שנקרא מהמנוע                                                      */
/* ------------------------------------------------------------------ */

/** נקודת קצה של בחירה, כפי שהמנוע מחזיר אותה. */
interface TextPoint {
  kind?: string;
  blockId?: string;
  offset?: number;
  story?: unknown;
}

interface LiveSelectionSnapshot {
  selectionTarget?: {
    kind?: string;
    start?: TextPoint | null;
    end?: TextPoint | null;
    story?: unknown;
    coordinateSpace?: string;
  } | null;
}

/** מה שנצרך מ-`activeEditor`. מוגדר כאן ולא מיובא — ההסבר ב-document-api.ts. */
export interface VerticalCaretEditor {
  host?: {
    readLiveSelectionSyncSnapshot?: () => LiveSelectionSnapshot | null | undefined;
  } | null;
  authoring?: {
    setSelectionTarget?: (input: {
      target: {
        kind: 'selection';
        start: TextPoint;
        end: TextPoint;
        story: unknown;
        coordinateSpace?: string;
      };
      focus?: boolean;
    }) => unknown;
  } | null;
}

export interface VerticalCaretHost {
  activeEditor?: VerticalCaretEditor | null;
}

export interface VerticalCaretOptions {
  /** ה-div שהמנוע מצייר בתוכו — `paintedHost(editor.ui)`. */
  host: HTMLElement | null;
  /** ה-SuperDoc של אותו מסמך. */
  superdoc: VerticalCaretHost | null | undefined;
}

export interface VerticalCaretHandle {
  dispose(): void;
}

/* ------------------------------------------------------------------ */
/* ההחלטות — פונקציות טהורות                                            */
/* ------------------------------------------------------------------ */

/** מה שנקרא מהאירוע. `isComposing`/`keyCode` חסרים בדמויות של הבדיקות. */
type VerticalKeyEvent = Pick<
  KeyboardEvent,
  'key' | 'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey'
> &
  Partial<Pick<KeyboardEvent, 'isComposing' | 'keyCode'>>;

/**
 * האם ההקשה היא חץ אנכי „נקי” — בלי Ctrl/Meta/Alt ובלי Shift.
 *
 * הרכבה במקלדת (IME) יוצאת: החץ שייך לחלונית ההרכבה ולא לסמן, והמסלול שלמטה
 * בולע את האירוע ב-`stopImmediatePropagation` — כלומר החלונית לא הייתה רואה
 * אותו כלל. התקן של הריפו ב-`src/ui/shortcuts/match.ts`: דפדפן שאינו מציב
 * `isComposing` מדווח `keyCode === 229`.
 */
export function isVerticalArrow(event: VerticalKeyEvent): boolean {
  if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return false;
  if (event.isComposing === true || event.keyCode === 229) return false;
  return !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey;
}

/** שורת היעד, וה-fragment שהיא יושבת בו. */
export interface VerticalTarget {
  fragment: Element;
  lines: readonly Element[];
  index: number;
  line: Element;
}

/**
 * השורה החזותית הבאה בכיוון הזה, או `null` כשאין כזו שאפשר לאמת.
 *
 * בתוך הפסקה זו פשוט השורה הסמוכה. בקצה הפסקה זה הבלוק הבא **באותו מכל
 * ציור** — ורק אם הוא צמוד במסמך ולא רק ברשימת הבורר (`contiguousWith`).
 *
 * טבלה נשארת למנוע במפורש: היא מצוירת כ-fragment בלי `data-source-node-id`,
 * והמנוע נכנס לתא כראוי. ‏`rtl-caret.ts` כן יורד לתא בחצים האופקיים, מפני ששם
 * מסירה למנוע נמדדה **לא** ניטרלית; כאן היא כן, ולכן אין סיבה לקחת אחריות על
 * גאומטריה של טבלה.
 */
export function nextVisualLine(
  host: HTMLElement,
  fragment: Element,
  lines: readonly Element[],
  index: number,
  forward: boolean,
): VerticalTarget | null {
  const insideIndex = index + (forward ? 1 : -1);
  const inside = lines[insideIndex];
  if (inside) return { fragment, lines, index: insideIndex, line: inside };

  const sibling = neighbourOf(host, fragment, forward);
  if (!sibling) return null;
  if (!contiguousWith(fragment, sibling, forward)) return null;
  if (blockIdOf(sibling) === null) return null;

  const siblingLines = linesOf(sibling);
  const edgeIndex = forward ? 0 : siblingLines.length - 1;
  const line = siblingLines[edgeIndex];
  if (!line) return null;
  return { fragment: sibling, lines: siblingLines, index: edgeIndex, line };
}

/**
 * האם עמודת המטרה נופלת על הטקסט שצויר בשורה.
 *
 * זו בדיוק השאלה שמפרידה בין „המנוע צודק” ל„המנוע הפוך”: נמדד שכל עוד העמודה
 * בתוך השורה הוא נוחת נכון, ומחוצה לה — בקצה ההפוך.
 *
 * **הסבילות מכווצת את „בפנים”, ואינה מרחיבה אותו** — ושתי הטעויות אינן
 * סימטריות. „בטעות בחוץ” עולה **כלום**: אנחנו מתערבים, ו-`nearestSlot` מחזיר
 * בדיוק את החריץ שהמנוע היה בוחר בעצמו באזור הזה. „בטעות בפנים” עולה **רוחב
 * שורה שלם**: ההקשה נמסרת למנוע דווקא במקום שנמדד בו הפוך. לכן הגבול נסוג
 * פנימה: הפרש של עשירית פיקסל בין קצה השורה שמעליה לקצה השורה שמתחתיה —
 * צירוף מקרים שגרתי לחלוטין ברוחבי טקסט — אינו יכול לגרור מסירה שגויה.
 */
export function goalInside(
  extent: { left: number; right: number },
  goalX: number,
  tolerance = 1,
): boolean {
  return goalX >= extent.left + tolerance && goalX <= extent.right - tolerance;
}

/**
 * החריץ שהכי קרוב לעמודת המטרה.
 *
 * שוויון נשמר לראשון שנמצא, שהוא בעל ה-pm הנמוך יותר — אותה הכרעה כמו
 * ב-`visualTarget`, ומאותה סיבה: שני חריצים באותו x פירושם שהמנוע מצייר את
 * שניהם שם, וההעדפה חייבת להיות יציבה כדי שהקשה חוזרת לא תקפיץ ביניהם.
 */
export function nearestSlot(slots: readonly CaretSlot[], goalX: number): CaretSlot | null {
  let best: CaretSlot | null = null;
  let bestGap = Infinity;
  for (const slot of slots) {
    const gap = Math.abs(slot.x - goalX);
    if (gap < bestGap) {
      bestGap = gap;
      best = slot;
    }
  }
  return best;
}

/**
 * החריץ שהסמן צריך לנחות עליו בשורת היעד, אחרי שמירת תפר הגלישה.
 *
 * החריץ הראשון של שורת המשך הוא ההיסט שבתפר, והמנוע מצייר אותו ב**סוף השורה
 * שמעליה** (נמדד — ראו `startIsSeam`). נחיתה עליו הייתה נראית כמו הקשה שלא
 * עשתה דבר, ולכן היעד מוזז חריץ אחד פנימה: זה המקום הקרוב ביותר לעמודת המטרה
 * שאפשר בכלל **לבטא**. המחיר הוא תו אחד, והוא משולם רק בשורת המשך שקצה
 * ההתחלה שלה נסוג פנימה — כלומר פסקה מוזחת או ממורכזת שגלשה.
 *
 * ‏`null` פירושו „אין חריץ שאפשר לבטא”, וההקשה נמסרת למנוע.
 */
export function landingSlot(
  slots: readonly CaretSlot[],
  goalX: number,
  firstIsSeam: boolean,
): CaretSlot | null {
  const nearest = nearestSlot(slots, goalX);
  if (!nearest) return null;
  if (!firstIsSeam || nearest !== slots[0]) return nearest;
  return slots[1] ?? null;
}

/* ------------------------------------------------------------------ */
/* היירוט                                                               */
/* ------------------------------------------------------------------ */

/** מתקין את היירוט. המצב היחיד שלו הוא עמודת המטרה, והוא נפרק עם המסמך. */
export function installRtlVerticalArrows({
  host,
  superdoc,
}: VerticalCaretOptions): VerticalCaretHandle {
  if (!host) return { dispose() {} };

  /**
   * עמודת המטרה שלנו, ו„האם כבר כתבנו”.
   *
   * נזרעת מה-x שהסמן **מצויר** בו בהקשה האנכית הראשונה — אותו מקור שהמנוע
   * זורע ממנו את שלו, ולכן השתיים זהות עד שאנחנו כותבים. ‏`owns` הוא מה שקורה
   * אחרי: מהרגע שכתבנו, העמודה של המנוע היא מה שכתבנו ולא המקורית, ולכן אי
   * אפשר עוד למסור לו הקשה — גם כזו שהוא היה צודק בה.
   *
   * ## ולמה **לא** ב-x של החלון
   *
   * ‏`getBoundingClientRect` מחזיר קואורדינטות חלון, ומיכל הגלילה הוא ה-host
   * שלנו — שגולל אופקית **בזמן** התנועה האנכית: `caret-visibility.ts` יושב על
   * אותו אלמנט, מנוי על הבחירה, וכותב `scrollLeft` כשהסמן יוצא מהתצוגה. כלומר
   * הכתיבה שלנו עצמה מזיזה את כל הגאומטריה מתחת לעמודה, ובחלון צר — מצב
   * הבסיס בפאנל של אוצריא — ההקשה הבאה הייתה מחושבת מול מספר ממסגרת ישנה,
   * ונוחתת בקצה ההפוך. כלומר בדיוק התקלה שהמודול בא לתקן, מיוצרת על ידו.
   *
   * לכן העמודה נשמרת בקואורדינטות **תוכן**: `x + host.scrollLeft` בשמירה,
   * ופחות `scrollLeft` בקריאה. גלילה אופקית — של המודול, של המשתמש, או של
   * המאחז — אינה נוגעת בה. (גלילה אנכית אינה משנה x, ולכן אינה עניין כאן.)
   */
  let goalContentX: number | null = null;
  let owns = false;
  // A vertical key we pass through may move the engine selection. Remember
  // that one movement so it is not mistaken for an unrelated caret move on
  // the next key (which would discard the goal column we still own).
  let engineHandoffPending = false;

  /**
   * מה שכתבנו בפעם האחרונה — ואיתו התשובה לשאלה „האם מישהו אחר הזיז את הסמן
   * מאז”.
   *
   * זה לא ייתור של האיפוס שבמקש: **מקש שמודול אחר בולע איננו רואים.**
   * ‏`rtl-caret.ts` עוצר חץ אופקי ב-`stopImmediatePropagation`, ו-
   * `rtl-line-end.ts` עושה אותו דבר ל-`End` — שניהם על אותו host. אילו
   * הנכונות כאן הייתה נשענת על לראות את המקש, היא הייתה נשענת על **סדר
   * ההתקנה** ב-App.vue, ומשתנה בלי שאיש ישים לב ברגע שמישהו יסדר שם מחדש.
   *
   * ההשוואה מזהה תזוזות של מודולים אחרים, והמאזין הגלובלי מאפס על הקלדה,
   * לחיצה ומקשים שאינם אנכיים. חץ אנכי שמסרנו למנוע מסומן בנפרד, כדי שתנועה
   * צפויה שלו — למשל דרך פסקה ריקה — לא תיחשב כאיפוס של עמודת המטרה.
   */
  let lastWrite: { blockId: string; offset: number } | null = null;

  const forget = (): void => {
    goalContentX = null;
    owns = false;
    engineHandoffPending = false;
    lastWrite = null;
  };

  const passToEngine = (): void => {
    if (owns) engineHandoffPending = true;
  };

  const onKeyDown = (event: KeyboardEvent): void => {
    if (!isVerticalArrow(event)) return;

    const editor = superdoc?.activeEditor;
    const setSelectionTarget = editor?.authoring?.setSelectionTarget;
    const readSnapshot = editor?.host?.readLiveSelectionSyncSnapshot;
    if (typeof setSelectionTarget !== 'function' || typeof readSnapshot !== 'function') {
      passToEngine();
      return;
    }

    /* בלי try/catch, כמו ב-`rtl-caret.ts`: נמדד שהקריאה אינה זורקת. מה שהיא
       כן עושה הוא להחזיר `selectionTarget: null`, וזה מטופל כאן. */
    const snapshot: LiveSelectionSnapshot | null | undefined = readSnapshot.call(editor?.host);
    const selection = snapshot?.selectionTarget;
    const head = selection?.end;
    if (!head || head.kind !== 'text' || typeof head.blockId !== 'string') {
      passToEngine();
      return;
    }
    if (typeof head.offset !== 'number') {
      passToEngine();
      return;
    }

    const caretRect = host.querySelector(CARET_SELECTOR)?.getBoundingClientRect();
    if (!caretRect) {
      passToEngine();
      return;
    }
    const caretY = (caretRect.top + caretRect.bottom) / 2;

    // הסמן אינו היכן שהשארנו אותו — מישהו אחר הזיז אותו, והעמודה שלנו ישנה.
    if (owns && (lastWrite?.blockId !== head.blockId || lastWrite.offset !== head.offset)) {
      if (engineHandoffPending) {
        // A previous vertical key was deliberately given to the engine (for
        // example, to cross an empty paragraph). Track its new caret position
        // while retaining our original goal column.
        lastWrite = { blockId: head.blockId, offset: head.offset };
      } else {
        forget();
      }
    }
    engineHandoffPending = false;

    // מקומי, ולא קריאה חוזרת של השדה: `forget()` שבין לבין היה מאפס אותו.
    // ההמרה לקואורדינטות חלון ובחזרה — ההסבר ב-`goalContentX`.
    const scroll = host.scrollLeft;
    goalContentX ??= caretRect.left + scroll;
    const goal = goalContentX - scroll;

    const home = findFragment(host, head.blockId, head.offset, caretY);
    if (!home) {
      passToEngine();
      return;
    }

    const lines = linesOf(home.fragment);
    const index = lineAt(lines, home.caretPm, caretY);
    if (index < 0) {
      passToEngine();
      return;
    }

    const target = nextVisualLine(host, home.fragment, lines, index, event.key === 'ArrowDown');
    if (!target) {
      passToEngine();
      return;
    }

    /* שורת יעד ריקה: מקום סמן אחד, והמנוע נוחת עליו בכל עמודה. אין מה לתקן,
       וגם אין מה לבחור — מסירה כאן היא ניטרלית. */
    if (isEmptyRange(target.line)) {
      passToEngine();
      return;
    }

    const targetRtl = target.line.getAttribute('dir') === 'rtl';

    // LTR destinations were measured to behave correctly. When we have not
    // taken ownership, hand them back before measuring every run in the line.
    if (!owns && !targetRtl) {
      passToEngine();
      return;
    }

    const extent = lineExtent(target.line);
    if (!extent) {
      passToEngine();
      return;
    }

    /*
     * כיוון השורה נבדק לפני המדידה: שורה לטינית נמסרת למנוע בלי למדוד את
     * הריצות שבה. בשורה עברית, `lineExtent` בודק אם עמודת המטרה בפנים — ואז
     * גם היא נמסרת. זה זול מ-`readLineChars`: מלבן אחד לכל ריצה ולא לכל גרפמה.
     */
    if (!owns && goalInside(extent, goal)) {
      passToEngine();
      return;
    }

    /* `null` כשהטווחים המצוירים אינם מכסים את השורה ברצף — אין מיפוי, וההקשה
       נמסרת. שורה בלי תוכן כבר יצאה למעלה, ולכן אין כאן מסלול „ריקה”. */
    const chars = readLineChars(target.line);
    if (!chars) {
      passToEngine();
      return;
    }
    const slots = caretSlots(chars, targetRtl);

    const slot = landingSlot(slots, goal, startIsSeam(target.fragment, target.lines, target.index));
    if (!slot) {
      passToEngine();
      return;
    }

    const blockId = blockIdOf(target.fragment);
    if (!blockId) {
      passToEngine();
      return;
    }

    /* `NaN` כשהחלק הראשון של הבלוק אינו מצויר — פסקה שהתחילה בעמוד שהמנוע לא
       צייר. אין מיפוי מ-pm להיסט, וההקשה נמסרת. */
    const offset = slot.pm - blockStart(host, blockId);
    if (!Number.isFinite(offset)) {
      passToEngine();
      return;
    }

    /*
     * מכאן ואילך זה שלנו: המנוע לא יראה את ההקשה.
     *
     * ‏`stopPropagation` ולא `stopImmediatePropagation`, בשונה משני המודולים
     * האחים — וזו הכרעה ולא חוסר עקביות. מה שצריך להיחסם הוא **המנוע**, והוא
     * מאזין מתחתינו (`.v2-super-editor__stage`); עצירה בשלב ה-capture על
     * ה-host שלנו כבר מונעת ממנו לראות את האירוע. מה שגרסת ה-`Immediate`
     * הייתה מוסיפה הוא לחסום את המאזינים האחרים על **אותו** אלמנט.
     *
     * וזה כן משנה: חסימה כזאת הייתה מונעת מ-`rtl-line-end.ts` לאפס את הרמז
     * שלו („‏`Home` קדם ל-`End`”) כשהוא מותקן **אחרינו**, והרצף `Home` → חץ
     * למטה → `End` היה מגיע לסוף השורה **הבאה** במקום זו שהסמן בה. אין סיבה
     * להחזיק תלות בסדר ההתקנה כשהיא לא קונה כלום.
     */
    event.preventDefault();
    event.stopPropagation();
    engineHandoffPending = false;

    /*
     * היעד הוא המקום שהסמן כבר עליו. זה קורה בקצה המסמך, וגם כשהציור מפגר
     * אחרי התצלום. בולעים ולא כותבים: כתיבה חוזרת של אותה בחירה אינה מזיזה
     * דבר, ומסירה למנוע כאן הייתה מחזירה אותו לעמודה שדרסנו.
     */
    if (blockId === head.blockId && offset === head.offset) return;

    // ה-`story` מגיע מהתצלום — נמדד שהוא קיים גם על ה-target וגם על הקצה.
    const story = selection?.story ?? head.story;
    const coordinateSpace = selection?.coordinateSpace;
    const point: TextPoint = { kind: 'text', blockId, offset, story };

    owns = true;
    lastWrite = { blockId, offset };

    /* בלי try/catch ובלי `.catch`, כמו ב-`rtl-caret.ts`: כתובת פסולה אינה
       זורקת ואינה דוחה — היא מחזירה קבלה עם `ok:false`. ההקשה כבר נבלעה,
       והסמן פשוט נשאר. */
    void setSelectionTarget.call(editor?.authoring, {
      target: {
        kind: 'selection',
        start: point,
        end: point,
        story,
        ...(typeof coordinateSpace === 'string' ? { coordinateSpace } : {}),
      },
      focus: true,
    });

  };

  /**
   * האיפוס — ולמה דווקא על ה-`document`, ולא על ה-host כמו ההקשה עצמה.
   *
   * כל מקש שאינו חץ אנכי מסיים את התנועה האנכית, ואיתה את העמודה: הקלדה,
   * ‏`Shift+חץ`, `Ctrl+חץ`, וגם **חץ אופקי ו-`End`**. אבל שני האחרונים נבלעים
   * על ה-host ב-`stopImmediatePropagation` — `rtl-caret.ts` ו-
   * `rtl-line-end.ts` יושבים שם — ולכן מאזין על ה-host היה רואה אותם רק אם
   * הוא נרשם **לפניהם**. כלומר הנכונות הייתה נשענת על סדר שלוש השורות
   * ב-App.vue, ומשתנה בשקט ברגע שמישהו מסדר שם מחדש.
   *
   * שלב ה-capture על ה-`window` קודם לכל מאזין אחר בדף — הוא הצומת הראשון
   * במסלול — ולכן האיפוס אינו תלוי בסדר ההתקנה כלל.
   *
   * ‏**ולמה `window` ולא `document`.** גם `document` היה קודם לכל מאזין על
   * ה-host, אבל הוא **אינו** ריק: `list-autoformat-install.ts` רושם שם
   * `keydown` בשלב capture ובולע ב-`stopImmediatePropagation` שני מסלולים —
   * ‏`Ctrl+Z`/`Ctrl+Y` שנצרכו, ורווח שיורט. שניהם מזיזים את הסמן, ואילו הוא
   * היה נרשם לפנינו, `Ctrl+Z` אחרי חץ אנכי לא היה מאפס. כלומר על `document`
   * התלות בסדר לא הייתה נעלמת אלא רק עוברת מקום.
   *
   * ולחיצת עכבר קובעת סמן חדש, ואיתו עמודה חדשה — כמו ב-Word — וגם היא
   * נרשמת שם: לחיצה על הרצועה או על דיאלוג מסיימת את התנועה בדיוק כמו לחיצה
   * בתוך המסמך.
   */
  const onAnyKeyDown = (event: KeyboardEvent): void => {
    if (!isVerticalArrow(event)) forget();
  };
  const onPointerDown = (): void => forget();

  host.addEventListener('keydown', onKeyDown, true);
  window.addEventListener('keydown', onAnyKeyDown, true);
  window.addEventListener('pointerdown', onPointerDown, true);

  return {
    dispose() {
      host.removeEventListener('keydown', onKeyDown, true);
      window.removeEventListener('keydown', onAnyKeyDown, true);
      window.removeEventListener('pointerdown', onPointerDown, true);
    },
  };
}
