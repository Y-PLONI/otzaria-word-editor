/**
 * אזכור „@”: הקלדת „@פסחים לד” פותחת רשימת הצעות, ובחירה בה מחליפה את האזכור
 * בקישור עומק לאוצריא.
 *
 * המבנה זהה ל-book-completion-overlay.ts, ומאותם טעמים שנמדדו שם:
 *
 * - **הרשימה היא DOM חיצוני לעץ של SuperDoc.** אין API להוסיף decoration בתוך
 *   המסמך, ושאילתה על מחלקות פנימיות של המנוע אסורה (engine-boundaries).
 * - **המיקום מגיע מ-`ui.selection.getAnchorRect`.** משטח ההקלדה של המנוע הוא
 *   `<textarea>` ברוחב פיקסל, וה-`getBoundingClientRect` של ה-anchorNode
 *   הדפדפני הוא אפס תמיד — `window.getSelection()` אינו כלי אמין כאן.
 * - **הטריגר נקרא ב-`input`/`keyup` ולא ב-`keydown`.** ב-keydown הטקסט עדיין
 *   לא נכתב למודל, ו-`doc.selection` היה מחזיר את המצב הישן.
 * - **`keydown` נתפס ב-capture, ו-`preventDefault` רק כשהרשימה פתוחה** — אחרת
 *   חצים ו-Tab מפסיקים להתנהג רגיל בשאר המסמך.
 *
 * ## כתיבת הקישור: מחיקת האזכור, כתיבת הטקסט, ואז `hyperlinks.wrap`
 *
 * המסלול הזה החליף את `hyperlinks.insert`, ולא מטעמי סגנון — `insert` אינה
 * יכולה לכתוב את הקישור השני בפסקה. מה שנמדד על 2.15.0:
 *
 * 1. **`hyperlinks.insert` נדחית בכל פסקה שכבר יש בה קישור.** לא רק בקצה
 *    שלו: הכנסה בהיסט 24, אחרי קישור שנגמר ב-19 ועם „ וכן ” מפריד ביניהם,
 *    חזרה `INVALID_CONTEXT / hyperlink-nested-unsupported` בדיוק כמו הכנסה
 *    בהיסט 19 עצמו. זה מה שהפך „שני קישורים בשורה אחת” לבלתי-אפשרי.
 * 2. **`hyperlinks.insert` אינה מזיזה את הסמן.** אחרי הכנסה בהיסט 8 הסמן
 *    נשאר ב-8 — כלומר **לפני** הקישור שנכתב, וכל מה שהוקלד אחריו נדחף לפניו.
 *    ‏„ראה @פסחים לד” ואז „ וכן …” יצא „ראה  וכן …פסחים דף לד”.
 * 3. **`hyperlinks.wrap` כן מצליחה** על טווח שנכתב זה עתה ב-`doc.insert`,
 *    ובאותה פסקה שבה `insert` נדחתה. זה שינוי מול מה שנמדד ב-2.12.0, שם
 *    היא החזירה `INVALID_TARGET` על טווח כזה; ההערה הישנה ב-
 *    `docs/engine-gaps.md` תוארכה מחדש בהתאם.
 * 4. **`doc.insert` עם `value: ''` מוחקת טווח**, ואף היא אינה מזיזה את הסמן.
 *
 * לכן: מוחקים את האזכור, מנקים צומתי קישור ריקים שחוסמים את המקום
 * (hyperlink-blanks.ts), כותבים את הטקסט, עוטפים אותו, ומציבים את הסמן
 * אחריו — כדי שההקלדה הבאה תמשיך מהמקום הנכון.
 *
 * ## הקישור אינו לחיץ מהמנוע, ולכן יש גשר
 *
 * נמדד: ה-DomPainter של המנוע מסנן כל href מול רשימת סכימות קבועה
 * (`http`, `https`, `mailto`, `tel`, `sms`), חוסם את השאר, ואינו מקבל
 * קונפיגורציה בנקודת הקריאה. הקישור נכתב ל-DOCX תקין ועובד בכל תוכנה
 * שפותחת אותו — אבל בתוך העורך `onActivate` (otzaria-link-activation.ts)
 * אינו נקרא, כי המנוע אינו מצייר את הריצה כקישור כלל. הלחיצה מטופלת
 * ב-otzaria-link-click.ts, שם גם המדידה המלאה. ראו docs/engine-gaps.md.
 */
import { PASTE_STARTED, PASTE_FINISHED } from './paste-events';
import type { SuperDoc } from 'superdoc';
import {
  WORD_WINDOW_RADIUS,
  type ResolvedRangeLike,
  type SelectionPointLike,
  type SelectionTargetLike,
  type WordSelectionDoc,
} from './word-selection';
import type { DocReceipt, MaybePromise } from './document-api';
import {
  buildLinkText,
  buildRefHref,
  isQueryable,
  parseAtTrigger,
  suggestionSubtitle,
  type ResolvedRefHit,
} from './at-mention';
import { isCaretFollowScroll } from './caret-visibility';
import { removeBlankHyperlinks, type HyperlinkBlanksDoc } from './hyperlink-blanks';
import { isHyperlinkOverlapFailure } from './hyperlinks-manage';
import { resolveRef } from '../host/otzaria-reader';

/** כתובת טווח טקסט, כפי ש-`hyperlinks.wrap`/`insert` מקבלים אותה. */
interface TextAddressLike {
  kind: 'text';
  blockId: string;
  range: { start: number; end: number };
  story?: unknown;
}

interface HyperlinkSpecLike {
  destination: { href: string };
}

export interface AtMentionDoc extends WordSelectionDoc, HyperlinkBlanksDoc {
  insert?: (input: { value: string; type: 'text'; target?: unknown }) => MaybePromise<DocReceipt>;
  hyperlinks?: {
    wrap?: (input: {
      target: TextAddressLike;
      link: HyperlinkSpecLike;
    }) => MaybePromise<DocReceipt>;
    list?: (input?: Record<string, unknown>) => MaybePromise<unknown>;
    remove?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
  } | null;
}

/** מלבן צבוע, כפי ש-`ui.selection.getAnchorRect` מחזירה אותו. */
export interface AnchorRectLike {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface AtMentionHost {
  activeEditor?: { doc?: AtMentionDoc | null } | null;
  ui?: {
    selection?: {
      getAnchorRect?: (input?: { placement?: 'start' | 'end' | 'center' }) => AnchorRectLike | null;
      /** מציבה סמן. אותה ידית ש-caret-anchor.ts משתמשת בה להחזרת המקום. */
      apply?: (target: SelectionTargetLike) => unknown;
    } | null;
  } | null;
}

export type AtMentionTarget = SuperDoc | AtMentionHost | null | undefined;

export interface AtMentionHandle {
  dispose(): void;
}

export interface AtMentionOptions {
  /** הודעת סטטוס למשתמש — כשל הרשאה, כשל כתיבה. לא נקראת על „אין התאמות”. */
  onStatus?: (message: string, isError: boolean) => void;
}

/** אותו debounce כמו בהשלמה מהספר; כאן הוא גם חוסך קריאות RPC. */
const INPUT_DEBOUNCE_MS = 180;

/**
 * התו שפותח אזכור. קבוע מפני שהוא נקרא בשני מסלולים שונים — `event.key`
 * של המקש, ו-`InputEvent.data` של הטקסט — ומחרוזת אחת בשניהם.
 */
const AT_SIGN = '@';

/** כמה הצעות להציג. מעבר לזה הרשימה מכסה את הטקסט שמעליה. */
const MAX_SUGGESTIONS = 8;

/**
 * ניקוי שאריות במסמך כולו אינו חלק מהפעולה שהמשתמש מחכה לה. הוא נשמר,
 * אך נדחה לזמן שקט כדי שקישור ישן במסמך אחר לא ינעל את ההקלדה כאן.
 */
const BLANK_HYPERLINK_CLEANUP_DELAY_MS = 750;

const POPUP_CLASS = 'otzaria-at-mention';
const POPUP_MAX_HEIGHT_PX = 260;
const POPUP_WIDTH_PX = 320;
/** מרווח בין הסמן לרשימה, וגם השוליים מקצה החלון. */
const POPUP_GAP_PX = 4;

function pointAt(blockId: string, offset: number, story: unknown): SelectionPointLike {
  const point: SelectionPointLike = { kind: 'text', blockId, offset };
  if (story !== undefined && story !== null) point.story = story;
  return point;
}

interface Seed {
  blockId: string;
  offset: number;
  story: unknown;
}

/**
 * הזרע נקרא מ-`selectionTarget` ולא מ-`target`: רק לו יש נקודות קצה, והוא
 * הצורה ש-`doc.insert` מקבל (ראו doc-selection.ts).
 */
function readCaretSeed(target: SelectionTargetLike | null | undefined): Seed | null {
  if (!target || target.kind !== 'selection') return null;
  if (target.coordinateSpace !== undefined && target.coordinateSpace !== 'visible') return null;
  const { start, end } = target;
  if (start?.kind !== 'text' || end?.kind !== 'text') return null;
  if (typeof start.blockId !== 'string' || start.blockId !== end.blockId) return null;
  if (typeof start.offset !== 'number' || start.offset !== end.offset) return null;
  return { blockId: start.blockId, offset: start.offset, story: target.story ?? start.story ?? null };
}

interface CaretText {
  /** הטקסט מתחילת החלון ועד הסמן. */
  beforeCaret: string;
  /** ההיסט בבלוק שממנו מתחיל `beforeCaret`. */
  base: number;
  cursorOffset: number;
  blockId: string;
  story: unknown;
}

/** קוראת את הטקסט שלפני הסמן. `null` = אין סמן יחיד בבלוק טקסט. */
async function readCaretText(doc: AtMentionDoc): Promise<CaretText | null> {
  if (typeof doc.selection?.current !== 'function') return null;
  if (typeof doc.ranges?.resolve !== 'function') return null;

  const info = await doc.selection.current();
  const seed = readCaretSeed(info?.selectionTarget);
  if (!seed) return null;

  const from = Math.max(0, seed.offset - WORD_WINDOW_RADIUS);
  const request: Record<string, unknown> = {
    start: { kind: 'point', point: pointAt(seed.blockId, from, seed.story) },
    end: { kind: 'point', point: pointAt(seed.blockId, seed.offset, seed.story) },
  };
  if (seed.story) request.in = seed.story;

  const resolved: ResolvedRangeLike | undefined = await doc.ranges.resolve(request);
  const text = resolved?.preview?.text;
  // `preview.text` נחתך ב-200 תווים; חלון של 90 בטוח מתחת לגבול, וקטיעה כאן
  // הייתה מזיזה את היסט ה-@ ביחס לטקסט.
  if (typeof text !== 'string' || resolved?.preview?.truncated === true) return null;

  return {
    beforeCaret: text,
    base: resolved?.target?.start?.offset ?? from,
    cursorOffset: seed.offset,
    blockId: seed.blockId,
    story: seed.story,
  };
}

type Session =
  | { kind: 'idle' }
  | {
      kind: 'suggesting';
      hits: ResolvedRefHit[];
      activeIndex: number;
      /** ההפניה כפי שהוקלדה — נשלחת ל-detection כשאין id חד-משמעי. */
      query: string;
      /** היסט ה-„@” בבלוק: תחילת הטווח שיוחלף. */
      replaceStart: number;
      cursorOffset: number;
      blockId: string;
      story: unknown;
    };

/** מתקינה את הפיצ'ר על ה-container של מסמך יחיד. ראו create-editor.ts:EditorSession.container. */
export function installAtMention(
  container: HTMLElement,
  host: AtMentionTarget,
  options: AtMentionOptions = {},
): AtMentionHandle {
  const doc = (host as AtMentionHost | null | undefined)?.activeEditor?.doc;
  const selectionHandle = (host as AtMentionHost | null | undefined)?.ui?.selection;
  let disposed = false;
  /** כיבוי לכל אורך המסמך אחרי כשל שאינו חולף — ראו `evaluate`. */
  let stopped = false;
  /**
   * היסט ה-„@” שהמשתמש דחה ב-Escape. בלעדיו התו הבא שיוקלד היה פותח את אותה
   * רשימה מחדש, וה-Escape היה נראה כאילו לא עשה דבר.
   */
  let dismissedAt: number | null = null;
  /**
   * האם יש בכלל סיכוי שהסמן יושב בתוך אזכור — **וזה מה שמוציא את הפיצ'ר
   * ממסלול ההקלדה.**
   *
   * `evaluate` הוא שתי קריאות RPC למנוע (`selection.current` ואז
   * `ranges.resolve`), והוא רץ 180ms אחרי כל תו שנקלד — גם באמצע טקסט רגיל
   * שאין בו „@” בשום מקום, כלומר כמעט תמיד. הבדיקה עצמה זולה רק בדיעבד:
   * התשובה היא „אין טריגר”, אבל היא הגיעה אחרי שהמנוע פתר טווח.
   *
   * מה שמחמש: הקלדת „@” (נקראת מ-`InputEvent.data`), הדבקה או כל
   * `inputType` שאין בו מה לקרוא, ותנועת סמן/לחיצה — מסלולים נדירים ביחס
   * לתו. מה שמפרק: `evaluate` שמצא שאין טריגר. כלומר מרגע „@” והלאה
   * המסלול המלא רץ על כל תו, עד שהאזכור נסגר — בדיוק החלון שבו הוא נחוץ.
   */
  let armed = false;
  let session: Session = { kind: 'idle' };
  /**
   * כתיבת קישור היא שתי פעולות של המנוע (מחיקה ואז insert). בזמן הקצר הזה
   * אסור לתת להקשה נוספת להזיז את הסמן או לערוך את הטווח, אחרת אי אפשר לדעת
   * לאן לשחזר את האזכור אם פעולת הקישור נדחית.
   */
  let writing = false;
  let debounceTimer: ReturnType<typeof setTimeout> | undefined;
  let blankCleanupTimer: ReturnType<typeof setTimeout> | undefined;
  let popupEl: HTMLDivElement | null = null;
  /** אחרון שדווח, כדי לא להציף את שורת הסטטוס באותה הודעה. */
  let lastReported: string | null = null;
  /**
   * טוקן ריצה: הערכה אחת עושה קריאת מנוע וקריאת RPC, וקריאה איטית שמסתיימת
   * אחרי מאוחרת ממנה הייתה מחזירה הצעות של מחרוזת שכבר לא מוקלדת.
   */
  let evalToken = 0;
  const pendingPastes = new Set<object>();

  function report(message: string, isError: boolean): void {
    if (lastReported === message) return;
    lastReported = message;
    options.onStatus?.(message, isError);
  }

  function hidePopup(): void {
    popupEl?.remove();
    popupEl = null;
  }

  function closeSession(): void {
    session = { kind: 'idle' };
    hidePopup();
  }

  function ensurePopup(): HTMLDivElement {
    if (popupEl) return popupEl;
    const el = document.createElement('div');
    el.className = POPUP_CLASS;
    el.setAttribute('role', 'listbox');
    el.setAttribute('aria-label', 'הצעות מקורות');
    Object.assign(el.style, {
      position: 'fixed',
      zIndex: '2147483647',
      width: `${POPUP_WIDTH_PX}px`,
      maxHeight: `${POPUP_MAX_HEIGHT_PX}px`,
      overflowY: 'auto',
      direction: 'rtl',
      textAlign: 'right',
      background: 'var(--otzaria-surface, #ffffff)',
      color: 'var(--otzaria-text, #202124)',
      border: '1px solid rgba(0, 0, 0, 0.16)',
      borderRadius: '8px',
      boxShadow: '0 6px 20px rgba(0, 0, 0, 0.18)',
      padding: '4px 0',
      font: '13px system-ui, sans-serif',
    } satisfies Partial<CSSStyleDeclaration>);
    // בלי זה הלחיצה גוזלת את המיקוד מהעורך, והסמן — ואיתו טווח ההחלפה — אובד.
    el.addEventListener('mousedown', (event) => event.preventDefault());
    document.body.appendChild(el);
    popupEl = el;
    return el;
  }

  function renderPopup(): void {
    // reference מקומי: `session` הוא משתנה משתנה, ובתוך ה-callbacks TypeScript
    // מאבד את הצמצום. זה אותו אובייקט, ולכן עדכון `activeIndex` דרכו תקף.
    const open = session;
    if (open.kind !== 'suggesting') return hidePopup();
    const rect = selectionHandle?.getAnchorRect?.({ placement: 'end' }) ?? null;
    if (!rect || (rect.width === 0 && rect.height === 0)) return hidePopup();

    const el = ensurePopup();
    el.textContent = '';
    open.hits.forEach((hit, index) => {
      const row = document.createElement('div');
      row.id = `${POPUP_CLASS}-${index}`;
      row.setAttribute('role', 'option');
      row.setAttribute('aria-selected', String(index === open.activeIndex));
      Object.assign(row.style, {
        padding: '6px 12px',
        cursor: 'pointer',
        background: index === open.activeIndex ? 'rgba(26, 115, 232, 0.12)' : 'transparent',
      } satisfies Partial<CSSStyleDeclaration>);

      const title = document.createElement('div');
      title.textContent = buildLinkText(hit, open.query);
      title.style.whiteSpace = 'nowrap';
      title.style.overflow = 'hidden';
      title.style.textOverflow = 'ellipsis';
      row.appendChild(title);

      const subtitle = suggestionSubtitle(hit);
      if (subtitle) {
        const sub = document.createElement('div');
        sub.textContent = subtitle;
        Object.assign(sub.style, {
          fontSize: '11px',
          opacity: '0.65',
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
        } satisfies Partial<CSSStyleDeclaration>);
        row.appendChild(sub);
      }

      row.addEventListener('mouseenter', () => {
        if (session !== open || open.activeIndex === index) return;
        open.activeIndex = index;
        renderPopup();
      });
      row.addEventListener('click', () => void accept(index));
      el.appendChild(row);
    });

    el.setAttribute('aria-activedescendant', `${POPUP_CLASS}-${open.activeIndex}`);
    positionPopup(el, rect);
    // גלילה לפריט הפעיל היא נוחות בלבד, והיא אינה קיימת בכל סביבה — כשל שלה
    // לא אמור להפיל את הציור עצמו.
    const active = el.children[open.activeIndex];
    if (active && typeof active.scrollIntoView === 'function') {
      active.scrollIntoView({ block: 'nearest' });
    }
  }

  /**
   * עוגן `right` ולא `left`: בעברית הרשימה צריכה להיפתח מנקודת הסמן שמאלה,
   * ועוגן שמאלי היה דוחף אותה לכיוון ההפוך. אנכית — מתחת לסמן אם יש מקום.
   */
  function positionPopup(el: HTMLDivElement, rect: AnchorRectLike): void {
    const right = Math.min(
      Math.max(POPUP_GAP_PX, window.innerWidth - rect.left),
      window.innerWidth - POPUP_WIDTH_PX - POPUP_GAP_PX,
    );
    el.style.right = `${Math.max(POPUP_GAP_PX, right)}px`;
    el.style.left = 'auto';

    const below = rect.top + rect.height + POPUP_GAP_PX;
    const height = Math.min(el.scrollHeight, POPUP_MAX_HEIGHT_PX);
    if (below + height <= window.innerHeight - POPUP_GAP_PX) {
      el.style.top = `${below}px`;
      el.style.bottom = 'auto';
    } else {
      el.style.top = 'auto';
      el.style.bottom = `${Math.max(POPUP_GAP_PX, window.innerHeight - rect.top + POPUP_GAP_PX)}px`;
    }
  }

  async function evaluate(): Promise<void> {
    if (disposed || stopped || !doc) return;
    const token = ++evalToken;

    let caret: CaretText | null;
    try {
      caret = await readCaretText(doc);
    } catch {
      caret = null;
    }
    if (token !== evalToken || disposed) return;
    if (!caret) return closeSession();

    const trigger = parseAtTrigger(caret.beforeCaret);
    if (!trigger || !isQueryable(trigger)) {
      dismissedAt = null;
      // אין „@” פתוח לפני הסמן — ואין טעם להמשיך לשאול על כל תו. ראו `armed`.
      // `isQueryable` הוא היוצא מן הכלל: „@” שהוקלד הרגע ועוד אין אחריו
      // שאילתה נשאר מחומש בכוונה — התו הבא הוא בדיוק מה שממתינים לו.
      if (!trigger) armed = false;
      return closeSession();
    }

    const replaceStart = caret.base + trigger.atIndex;
    if (dismissedAt === replaceStart) return closeSession();
    dismissedAt = null;

    const result = await resolveRef(trigger.query.trim(), MAX_SUGGESTIONS);
    if (token !== evalToken || disposed) return;

    if (!result.ok) {
      report(result.message, true);
      // הרשאה חסרה אינה משתנה תוך כדי הפעלה, וכל הקלדת „@” נוספת הייתה
      // מייצרת עוד קריאת RPC שנדחית.
      if (result.reason === 'permission-denied') stopped = true;
      return closeSession();
    }
    lastReported = null;
    if (result.value.length === 0) return closeSession();

    session = {
      kind: 'suggesting',
      hits: result.value,
      activeIndex: 0,
      query: trigger.query,
      replaceStart,
      cursorOffset: caret.cursorOffset,
      blockId: caret.blockId,
      story: caret.story,
    };
    renderPopup();
  }

  function scheduleEvaluate(): void {
    if (pendingPastes.size > 0) return;
    if (debounceTimer !== undefined) clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => void evaluate(), INPUT_DEBOUNCE_MS);
  }

  /**
   * הניקוי הרחב הוא תחזוקה של המסמך, לא תנאי לכתיבת האזכור הנוכחי. הצומת
   * החוסם באותו בלוק כבר הוסר לפני המחיקה; כאן מנקים שאריות אחרות רק אחרי
   * שהמשתמש חדל להקליד לרגע. טיימר יחיד מאחד כמה אזכורים רצופים לסריקה אחת.
   */
  function scheduleBlankHyperlinkCleanup(): void {
    clearTimeout(blankCleanupTimer);
    blankCleanupTimer = setTimeout(() => {
      blankCleanupTimer = undefined;
      void removeBlankHyperlinks(doc).catch((error) => {
        console.warn('[otzaria-word] אזכור: ניקוי קישורים ריקים נכשל', error);
      });
    }, BLANK_HYPERLINK_CLEANUP_DELAY_MS);
  }

  function move(delta: number): void {
    if (session.kind !== 'suggesting') return;
    const count = session.hits.length;
    session.activeIndex = (session.activeIndex + delta + count) % count;
    renderPopup();
  }

  /**
   * מציבה את הסמן אחרי הקישור שנכתב.
   *
   * בלי זה הסמן נשאר במקום שבו האזכור התחיל — כלומר לפני הקישור — וההקלדה
   * הבאה נדחפת לפניו. נמדד ש-`apply` על ההיסט שאחרי הקישור אינו מצרף את
   * ההקלדה הבאה לקישור: הטקסט שנכתב אחריו נשאר מחוצה לו, והעוגן אינו זז.
   */
  function placeCaretAfter(blockId: string, offset: number, story: unknown): void {
    const apply = selectionHandle?.apply;
    if (typeof apply !== 'function') return;
    try {
      apply.call(selectionHandle, {
        kind: 'selection',
        start: pointAt(blockId, offset, story),
        end: pointAt(blockId, offset, story),
      } satisfies SelectionTargetLike);
    } catch (error) {
      // הקישור כבר נכתב; מיקום הסמן אינו שווה הודעת שגיאה למשתמש.
      console.warn('[otzaria-word] אזכור: הצבת הסמן אחרי הקישור נכשלה', error);
    }
  }

  /** כותבת את הקישור. ראו הערת המודול: מחיקה, כתיבת הטקסט, ואז עטיפה. */
  async function writeLink(
    address: TextAddressLike,
    text: string,
    href: string,
    originalText: string,
  ): Promise<DocReceipt | null> {
    const wrap = doc?.hyperlinks?.wrap;
    if (typeof doc?.insert !== 'function' || typeof wrap !== 'function') return null;

    const { story } = address;
    const over = (blockId: string, start: number, end: number): SelectionTargetLike => ({
      kind: 'selection',
      start: pointAt(blockId, start, story),
      end: pointAt(blockId, end, story),
      ...(story ? { story } : {}),
    });

    /**
     * צומת קישור ריק ששרד מחיקה קודמת יושב במקום שעומדים לכתוב אליו וחוסם
     * את העטיפה — זה „מחקתי קישור ואני מתייג מחדש באותה שורה”. כאן סורקים
     * רק את הבלוק הזה: זה כל מה שנדרש לכתיבה, ומסמך בעל שאריות רבות במקום
     * אחר אינו צריך לנעול את ההקלדה. ניקוי שאר המסמך נשמר, אבל נדחה לזמן
     * שקט אחרי הצלחת הכתיבה. ראו hyperlink-blanks.ts.
     *
     * **לפני המחיקה, ובמכוון.** מכאן והלאה האזכור של המשתמש כבר לא במסמך,
     * וכל קריאת מנוע בחלון הזה היא קריאה שאם לא תחזור — הטקסט שלו אבד בלי
     * שנכתב דבר במקומו. הסריקה אינה תלויה במחיקה (צומת ריק אינו נוצר ממנה,
     * והסרתו אינה מזיזה היסטים), ולכן אין סיבה להחזיק אותה שם.
     */
    await removeBlankHyperlinks(doc, { blockId: address.blockId });

    const deleted = await doc.insert({
      value: '',
      type: 'text',
      target: over(address.blockId, address.range.start, address.range.end),
    });
    if (deleted?.success === false) return deleted;

    // הנקודה שנפתחה נקראת מהבחירה ולא מחושבת: זו גם נקודת הכתיבה האמיתית
    // וגם הסנכרון מול המחיקה שקדמה לה.
    let at = { blockId: address.blockId, offset: address.range.start };
    try {
      const after = readCaretSeed((await doc.selection?.current?.())?.selectionTarget);
      if (after) at = { blockId: after.blockId, offset: after.offset };
    } catch {
      /* נשארים עם תחילת הטווח שנמחק */
    }

    /**
     * שתי הפעולות אינן טרנזקציוניות עם המחיקה שלפניהן. אם אחת מהן נכשלת,
     * האזכור שהמשתמש הקליד חייב לחזור למקומו; הודעת שגיאה בלי שחזור הייתה
     * מאבדת אותו, ובלי האזכור אין גם מה לנסות שוב.
     */
    const restore = async (start: number, end: number): Promise<void> => {
      try {
        await doc.insert?.({ value: originalText, type: 'text', target: over(at.blockId, start, end) });
      } catch (error) {
        // זו תקלה כפולה ונדירה; הכשל המקורי עדיין מדווח ל-caller.
        console.warn('[otzaria-word] אזכור: שחזור הטקסט נכשל', error);
      }
    };

    let written: DocReceipt | undefined;
    try {
      written = await doc.insert({
        value: text,
        type: 'text',
        target: over(at.blockId, at.offset, at.offset),
      });
    } catch (error) {
      // זריקה כאן היא בדיוק המקרה שבו האזכור כבר נמחק ודבר לא נכתב במקומו.
      await restore(at.offset, at.offset);
      throw error;
    }
    if (written?.success === false) {
      await restore(at.offset, at.offset);
      return written;
    }

    // היסטי המנוע נמדדו כיחידות UTF-16, בדיוק כמו `String.length`: כתיבת
    // „פסחים דף לד” (11) בהיסט 8 נתנה עוגן 8..19.
    const start = at.offset;
    const end = at.offset + text.length;

    try {
      const receipt = await wrap({
        target: { kind: 'text', blockId: at.blockId, range: { start, end }, ...(story ? { story } : {}) },
        link: { destination: { href } },
      });
      if (receipt?.success === false) {
        // הטקסט נכתב אך אינו קישור — מחליפים אותו בחזרה באזכור.
        await restore(start, end);
        return receipt;
      }
      placeCaretAfter(at.blockId, end, story);
      scheduleBlankHyperlinkCleanup();
      return receipt;
    } catch (error) {
      await restore(start, end);
      throw error;
    }
  }

  async function accept(index?: number): Promise<void> {
    if (session.kind !== 'suggesting' || !doc) return;
    const hit = session.hits[index ?? session.activeIndex];
    if (!hit) return;

    const { replaceStart, cursorOffset, blockId, story, query } = session;
    const text = buildLinkText(hit, query);
    const href = buildRefHref(hit, query);
    const originalText = `@${query}`;
    closeSession();

    const address: TextAddressLike = {
      kind: 'text',
      blockId,
      range: { start: replaceStart, end: cursorOffset },
      ...(story ? { story } : {}),
    };

    let receipt: DocReceipt | null;
    writing = true;
    try {
      receipt = await writeLink(address, text, href, originalText);
    } catch (error) {
      console.warn('[otzaria-word] אזכור: כתיבת הקישור נכשלה', error);
      report('הוספת הקישור נכשלה', true);
      return;
    } finally {
      writing = false;
    }
    if (receipt === null) {
      report('הוספת קישור אינה זמינה במסמך זה', true);
      return;
    }
    if (receipt?.success === false) {
      report(
        isHyperlinkOverlapFailure(receipt)
          ? 'אי אפשר להוסיף קישור בתוך קישור קיים'
          : (receipt.failure?.message ?? 'הוספת הקישור נכשלה'),
        true,
      );
    }
  }

  /**
   * מה שנקלד, כשה-DOM מספר. `null` לכל מה שאינו הכנסת טקסט מוכרת — הדבקה,
   * גרירה, IME — ואז מחמשים בלי לשאול שאלות: זה מסלול שאינו לכל תו.
   */
  function insertedText(event: Event): string | null {
    if (!(event instanceof InputEvent)) return null;
    if (typeof event.data === 'string') return event.data;
    // מחיקה היא הכנסה של כלום, ולא „לא ידוע”: היא אינה יכולה לייצר „@”
    // חדש, ומה שהיא כן עושה — מחיקת האזכור — מטופל כשהמסלול מחומש ממילא.
    return event.inputType.startsWith('delete') ? '' : null;
  }

  /**
   * `input` — **המסלול המשני, וזה נמדד.**
   *
   * המנוע מטפל בהקלדה ב-`keydown` ומכניס את הטקסט בעצמו, ולכן במנוע האמיתי
   * **אין `input` ואין `beforeinput` בכלל**: probe על ה-dist הארוז הקליט
   * שלוש הקשות והחזיר שלושה `keyup` ואפס `input`. המאזין נשאר בשביל סביבה
   * שכן מדווחת (jsdom בבדיקות, ומסלול IME אפשרי), ולא כמקור היחיד.
   */
  const onInput = (event: Event): void => {
    if (writing) return;
    const inserted = insertedText(event);
    if (inserted === null || inserted.includes(AT_SIGN)) armed = true;
    if (!armed) return;
    scheduleEvaluate();
  };

  /**
   * המקשים שהרשימה מטפלת בהם אינם משנים טקסט, ולכן `keyup` שלהם אינו אמור
   * להעריך מחדש. נמדד: בלי הסינון הזה חץ למטה בנה session חדש שהחזיר את
   * הבחירה לפריט הראשון, ו-Escape נסגר ונפתח מיד.
   */
  const HANDLED_KEYS = new Set(['ArrowDown', 'ArrowUp', 'Enter', 'Tab', 'Escape']);

  /**
   * המקשים שמזיזים סמן בלי לערוך טקסט. הם מחמשים בעצמם: תנועת סמן היא הדרך
   * להיכנס לאזכור שכבר בטקסט (ולצאת ממנו), והיא אינה מגיעה בשום מסלול אחר.
   */
  const CARET_KEYS = new Set([
    'ArrowLeft',
    'ArrowRight',
    'Home',
    'End',
    'PageUp',
    'PageDown',
    'Backspace',
    'Delete',
  ]);

  /**
   * `keyup` — **המסלול העיקרי.** ראו `onInput` למה הוא ולא `input`.
   *
   * שלושה מחמשים כאן: „@” עצמו (המקש, ולא הטקסט שנכנס — זה מה שקיים),
   * מקשי הסמן, ומצב מחומש קיים. תו רגיל שאינו אחד מאלה יוצא בשורה אחת בלי
   * לגעת במנוע — וזה כל הפיצ'ר במסלול ההקלדה.
   */
  const onKeyUp = (event: KeyboardEvent): void => {
    if (writing) return;
    if (HANDLED_KEYS.has(event.key)) return;
    if (event.key === AT_SIGN || CARET_KEYS.has(event.key)) armed = true;
    if (!armed) return;
    scheduleEvaluate();
  };

  /**
   * לחיצה מציבה סמן, ואפשר שבתוך „@…” שכבר בטקסט; הדבקה מכניסה טקסט שאין
   * לו מקש. פעם אחת לכל אחד מהם — לא מסלול חם — וזה מה שמשלים את הצמצום.
   */
  const onArm = (): void => {
    if (writing) return;
    armed = true;
    scheduleEvaluate();
  };

  const onPasteStarted = (event: Event): void => {
    if (!(event instanceof CustomEvent) || typeof event.detail?.operation !== 'object') return;
    pendingPastes.add(event.detail.operation);
    evalToken += 1;
    if (debounceTimer !== undefined) clearTimeout(debounceTimer);
    closeSession();
  };
  const onPasteFinished = (event: Event): void => {
    if (!(event instanceof CustomEvent) || !pendingPastes.delete(event.detail?.operation)) return;
    if (!disposed && pendingPastes.size === 0 && event.detail?.evaluate !== false) onArm();
  };

  const onKeyDown = (event: KeyboardEvent): void => {
    if (writing) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if (session.kind !== 'suggesting') return;
    switch (event.key) {
      case 'ArrowDown':
        move(1);
        break;
      case 'ArrowUp':
        move(-1);
        break;
      case 'Enter':
      case 'Tab':
        void accept();
        break;
      case 'Escape':
        dismissedAt = session.replaceStart;
        closeSession();
        break;
      default:
        return;
    }
    event.preventDefault();
    event.stopPropagation();
  };

  // „המשתמש גלל משם” — אבל גלילה שהסמן גרר אחריו היא המשך של ההקלדה, ולא
  // ניווט. בלי ההבחנה, ‏`@פסחים` בחלון צר סגר את עצמו באמצע השאילתה ברגע
  // שהסמן חצה את שולי התצוגה. ראו `isCaretFollowScroll`.
  //
  // האירוע עצמו עובר, ולא המיכל: ההבחנה נשענת על ה-`target` שלו. אותו מאזין
  // רשום גם ל-`resize` (‏`target` = ה-`window`) ושומע גם גלילה של צאצא
  // (‏`target` = הצאצא) — נמדד, ושתיהן חייבות להמשיך לסגור.
  const onScroll = (event: Event): void => {
    if (isCaretFollowScroll(event)) return;
    closeSession();
  };
  const onBlur = (): void => {
    // המיקוד עזב את המסמך — אין סמן לעקוב אחריו. חזרה אליו מגיעה דרך
    // `mouseup` או תו חדש, ושניהם מחמשים מחדש.
    armed = false;
    closeSession();
  };

  container.addEventListener(PASTE_STARTED, onPasteStarted);
  container.addEventListener(PASTE_FINISHED, onPasteFinished);
  container.addEventListener('input', onInput);
  container.addEventListener('keyup', onKeyUp);
  container.addEventListener('mouseup', onArm);
  container.addEventListener('paste', onArm);
  container.addEventListener('keydown', onKeyDown, true);
  container.addEventListener('scroll', onScroll, true);
  container.addEventListener('focusout', onBlur);
  window.addEventListener('resize', onScroll);

  return {
    dispose() {
      disposed = true;
      if (debounceTimer !== undefined) clearTimeout(debounceTimer);
      clearTimeout(blankCleanupTimer);
      closeSession();
      container.removeEventListener(PASTE_STARTED, onPasteStarted);
      container.removeEventListener(PASTE_FINISHED, onPasteFinished);
      container.removeEventListener('input', onInput);
      container.removeEventListener('keyup', onKeyUp);
      container.removeEventListener('mouseup', onArm);
      container.removeEventListener('paste', onArm);
      container.removeEventListener('keydown', onKeyDown, true);
      container.removeEventListener('scroll', onScroll, true);
      container.removeEventListener('focusout', onBlur);
      window.removeEventListener('resize', onScroll);
    },
  };
}
