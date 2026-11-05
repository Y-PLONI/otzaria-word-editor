/**
 * "השלמה מהספר": ghost text אפור שמציע להשלים טקסט מתוך הספר הפתוח בקורא.
 *
 * ## למה overlay חיצוני ולא decoration של המנוע
 *
 * מנוע ה-DOCX קנייני ואסור לשנותו (ראו word-selection.ts). אין API ציבורי
 * להוסיף decoration בתוך המסמך, ולכן ה-ghost הוא `<span>` חיצוני לגמרי לעץ
 * ה-DOM של SuperDoc.
 *
 * ## מיקום ה-ghost: `ui.selection.getAnchorRect`, לא `window.getSelection()`
 *
 * נמדד: המנוע אינו מרנדר את הטקסט כתוכן native רגיל — ההקלדה נתפסת ב-
 * `<textarea>` נסתר, וה-DOM שה-selection הדפדפני מצביע עליו (anchorNode) הוא
 * mirror שה-`getBoundingClientRect` שלו אפס תמיד. `window.getSelection()`
 * אינו כלי אמין כאן. `ui.selection.getAnchorRect({ placement: 'end' })` הוא
 * ה-API הציבורי המקביל — הוא שואל את שכבת הגיאומטריה **של המנוע עצמו**
 * (layout-bridge/cursor-renderer), ולכן מחזיר את מה שבאמת מצויר על המסך, גם
 * כשאין רלוונטיות לסלקציה של הדפדפן.
 *
 * ## זיהוי הקלדה: `input`/`keyup`, לא `keydown`
 *
 * ב-`keydown` הטקסט עדיין לא נכתב למודל של המנוע; קריאת `doc.selection` שם
 * הייתה מחזירה את המצב הישן.
 *
 * ## Tab נתפס רק כשיש הצעה פעילה
 *
 * `keydown` על שלב ה-capture (כמו format-painter.ts/word-selection.ts), אבל
 * `preventDefault` נקרא רק כש-session.kind === 'suggesting' — אחרת Tab ממשיך
 * להתנהג רגיל (מעבר בין תאים בטבלה, למשל).
 *
 * ## למה ההחלפה היא "מחליפה את המילה החלקית", לא "מוסיפה אחריה"
 *
 * טקסט הספר מנוקד; מה שהמשתמש הקליד לרוב לא. `accept` שולח `target` שמכסה
 * מתחילת המילה החלקית ועד הסמן, ו-`insert` מחליף אותו בטקסט המקור המלא —
 * ראו ההסבר המקביל ב-book-completion.ts.
 */
import { PASTE_STARTED, PASTE_FINISHED } from './paste-events';
import type { SuperDoc } from 'superdoc';
import {
  WORD_INNER,
  WORD_WINDOW_RADIUS,
  wordBoundsIn,
  type ResolvedRangeLike,
  type SelectionPointLike,
  type SelectionTargetLike,
  type TextWindow,
  type WordSelectionDoc,
} from './word-selection';
import type { DocReceipt, MaybePromise } from './document-api';
import {
  buildSectionCache,
  matchAcrossBooks,
  matchBookTitle,
  sliceWords,
  type SectionWordCache,
} from './book-completion';
import { isCaretFollowScroll } from './caret-visibility';
import { loadAcronymDictionary } from './acronym-dictionary';
import { looksLikeAcronym } from './acronyms';
import { loadStaticSources } from './static-completion-dictionary';
import { matchStaticCompletion } from './static-completion';
import { parseAtTrigger } from './at-mention';
import {
  getCurrentReaderState,
  getSectionTextMap,
  normalizeSelectedText,
} from '../host/otzaria-reader';
import { on } from '../host/otzaria-client';
import type { ReaderState } from '../types/otzaria_plugin';

export interface CompletionDoc extends WordSelectionDoc {
  insert?: (input: { value: string; type: 'text'; target?: unknown }) => MaybePromise<DocReceipt>;
}

/** מלבן צבוע, כפי ש-`ui.selection.getAnchorRect` מחזירה אותו. */
export interface AnchorRectLike {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface CompletionHost {
  activeEditor?: { doc?: CompletionDoc | null } | null;
  ui?: {
    selection?: {
      getAnchorRect?: (input?: { placement?: 'start' | 'end' | 'center' }) => AnchorRectLike | null;
    } | null;
  } | null;
}

export type CompletionTarget = SuperDoc | CompletionHost | null | undefined;

export interface BookCompletionHandle {
  dispose(): void;
}

export interface BookCompletionOptions {
  /** הודעת סטטוס למשתמש — כשל הרשאה, כשל טעינה. לא נקראת לניקוי הצעה רגיל. */
  onStatus?: (message: string, isError: boolean) => void;
}

/** כמה מילות הקשר וכמה מילים להציע בכל סבב — ראו book-completion.ts. */
const CONTEXT_WORDS = 3;
const WORDS_TO_SHOW = 5;
const INPUT_DEBOUNCE_MS = 150;

/**
 * חלון השורות שנטען סביב המיקום בקורא.
 *
 * `currentIndex` של אוצריא הוא **אינדקס שורה** (פסקה), לא עמוד — נמדד: טעינה
 * של שורה בודדת נתנה 24-28 מילים בלבד, בעוד המשתמש רואה על המסך פסקאות
 * שלמות, וההשלמות הגיעו מטקסט שאינו מול עיניו. החלון הוא הקירוב ל"טקסט
 * המוצג": יותר מאחת, והרבה פחות מהספר כולו — וזה גם החיסכון ב-RAM.
 */
const CONTEXT_LINES_BACK = 5;
const CONTEXT_LINES_FORWARD = 25;
const CONTEXT_CHAR_CAP = 20_000;

/**
 * כל כמה זמן לבדוק מחדש איזה ספר/מיקום פתוח בקורא.
 *
 * האירועים `reader.current_book_changed`/`current_ref_changed` דורשים הרשאות
 * `events.subscribe:*` שאינן ב-manifest — ולכן ההרשמה אליהם אינה מקבלת דבר,
 * וזה בדיוק הבאג שנמדד: מעבר ספר השאיר את המילים של הספר הקודם. `getCurrentState`
 * כבר תחת `reader.open` הקיים, ולכן בדיקה עצלה בזמן הקלדה פותרת בלי הרשאה חדשה.
 */
const STATE_TTL_MS = 1200;

/**
 * כמה ספרים פתוחים נטענים להשלמה, כולל הספר שנקרא אחרון. כל ספר הוא חלון של
 * עד `CONTEXT_CHAR_CAP` תווים ו-31 קריאות לטעינה, ולכן יש תקרה.
 */
const MAX_OPEN_BOOKS = 6;

/** ספר פתוח בקורא, כפי שהוא מזוהה מול `getSectionTextMap`. */
interface BookRef {
  bookId: string;
  /** מזהה יציב. `bookId` הוא כותרת הטאב, ובמהדורה חלופית היא אינה שם הספר. */
  bookUid: string | null;
  index: number;
}

/** החלון הטעון של ספר פתוח אחד. */
interface BookContext extends BookRef {
  key: string;
  cache: SectionWordCache | null;
  /** טביעת התצוגה של שורת המרכז — ראו `displayChanged`. */
  fingerprint: string | null;
  loading: boolean;
  /** הוצא מרשימת הספרים הפתוחים; טעינה שבאוויר אינה כותבת אליו. */
  dropped: boolean;
}

/** מקום שמור לספר שעוד לא נטען, כדי שהאינדקס ב-`matchAcrossBooks` יישאר הספר. */
const EMPTY_CACHE = buildSectionCache('');

/** רק ספר טקסט: ל-PDF אין `getSectionTextMap`. סוג חסר = גרסת אוצריא ישנה. */
function isTextBook(type: string | null | undefined): boolean {
  return type === undefined || type === null || type === 'text';
}

/**
 * הספר שנקרא אחרון ראשון, ואחריו שאר טאבי הספרים. ריק כשאין ספר טקסט שנקרא
 * אחרון: ספר רקע היה תופס את מקומו, כולל ההשלמה בלי הקשר ששמורה לראשון בלבד.
 */
function openBookRefs(state: ReaderState): BookRef[] {
  if (state.currentBookId == null || typeof state.currentIndex !== 'number' || !isTextBook(state.currentType)) {
    return [];
  }
  const refs: BookRef[] = [
    { bookId: state.currentBookId, bookUid: state.bookUid ?? null, index: state.currentIndex },
  ];
  for (const tab of state.openTabs ?? []) {
    if (tab.toolId != null || tab.isSelf === true) continue;
    if (typeof tab.bookId !== 'string' || tab.bookId === '' || typeof tab.index !== 'number') continue;
    if (!isTextBook(tab.type)) continue;
    refs.push({ bookId: tab.bookId, bookUid: tab.bookUid ?? null, index: tab.index });
  }
  return refs;
}

function bookKey(ref: BookRef): string {
  return `${ref.bookUid ?? ref.bookId}:${ref.index}`;
}

const GHOST_CLASS = 'otzaria-book-completion-ghost';

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

/** כמו readSeed ב-word-selection.ts, אבל לסמן מכונס (לא טווח בחירה). */
function readCaretSeed(target: SelectionTargetLike | null | undefined): Seed | null {
  if (!target || target.kind !== 'selection') return null;
  if (target.coordinateSpace !== undefined && target.coordinateSpace !== 'visible') return null;

  const { start, end } = target;
  if (start?.kind !== 'text' || end?.kind !== 'text') return null;
  if (typeof start.blockId !== 'string' || start.blockId !== end.blockId) return null;
  if (typeof start.offset !== 'number' || start.offset !== end.offset) return null;

  return { blockId: start.blockId, offset: start.offset, story: target.story ?? start.story ?? null };
}

/** פירוק לרצף מילים (WORD_INNER) — לא מנרמל, רק מפריד. */
const WORD_RUN = new RegExp(`(?:${WORD_INNER.source})+`, 'gu');

interface TypedSnapshot {
  precedingWords: string[];
  /** ההיסטים בבלוק של תחילת כל מילה ב-`precedingWords`, באותו סדר. */
  precedingWordStarts: number[];
  partialWord: string;
  /** ההיסט בבלוק שממנו מתחילה ההחלפה: תחילת המילה החלקית, או הסמן עצמו. */
  replaceStart: number;
  cursorOffset: number;
  blockId: string;
  story: unknown;
  /** הטקסט הגולמי מתחילת החלון עד הסמן, לשם בידוד אזכור `@` מהשלמות ghost. */
  rawBeforeText: string;
}

async function readTypedSnapshot(doc: CompletionDoc): Promise<TypedSnapshot | null> {
  if (typeof doc.selection?.current !== 'function') {
    return null;
  }
  if (typeof doc.ranges?.resolve !== 'function') {
    return null;
  }

  const info = await doc.selection.current();
  const seed = readCaretSeed(info?.selectionTarget);
  if (!seed) {
    return null;
  }

  const from = Math.max(0, seed.offset - WORD_WINDOW_RADIUS);
  const to = seed.offset + WORD_WINDOW_RADIUS;
  const request: Record<string, unknown> = {
    start: { kind: 'point', point: pointAt(seed.blockId, from, seed.story) },
    end: { kind: 'point', point: pointAt(seed.blockId, to, seed.story) },
  };
  if (seed.story) request.in = seed.story;

  const resolved: ResolvedRangeLike | undefined = await doc.ranges.resolve(request);
  const text = resolved?.preview?.text;
  if (typeof text !== 'string' || resolved?.preview?.truncated === true) {
    return null;
  }

  const base = resolved?.target?.start?.offset ?? from;
  const windowEnd = resolved?.target?.end?.offset ?? base + text.length;
  const window: TextWindow = { text, base, atBlockStart: base === 0, atBlockEnd: windowEnd < to };

  const bounds = wordBoundsIn(window, seed.offset);
  const beforeText = text.slice(0, seed.offset - base);
  const runs = [...beforeText.matchAll(WORD_RUN)];
  const allWords = runs.map((run) => run[0]);
  const allStarts = runs.map((run) => base + (run.index ?? 0));

  const partialWord = bounds ? (allWords[allWords.length - 1] ?? '') : '';
  const precedingWords = bounds ? allWords.slice(0, -1) : allWords;
  const precedingWordStarts = bounds ? allStarts.slice(0, -1) : allStarts;
  const replaceStart = bounds ? bounds.start : seed.offset;

  return {
    precedingWords,
    precedingWordStarts,
    partialWord,
    replaceStart,
    cursorOffset: seed.offset,
    blockId: seed.blockId,
    story: seed.story,
    rawBeforeText: beforeText,
  };
}

type Suggestion =
  | { kind: 'idle' }
  | {
      kind: 'suggesting';
      /** מה שמוצג באפור: רק ההמשך שטרם הוקלד. */
      ghostText: string;
      /**
       * מה שנכתב בפועל בקבלה. יכול להיות ארוך מ-`ghostText`: הוא כולל גם את
       * מילות ההקשר שהמשתמש כבר הקליד, כדי להחליף אותן בגרסה המדויקת מהספר.
       */
      insertText: string;
      /** תחילת הטווח שיוחלף ב-`insertText`. */
      replaceStart: number;
      cursorOffset: number;
      blockId: string;
      story: unknown;
      /** אינדקס המילה הבאה בספר, כדי להציע רצף בלי חיפוש נוסף אחרי קבלה. */
      continueFrom: number | null;
      /** החלון שממנו באה ההצעה — `continueFrom` הוא אינדקס בו, ולא בספר אחר. */
      source: SectionWordCache | null;
    };

/** מתקינה את הפיצ'ר על ה-container של מסמך יחיד. ראו create-editor.ts:EditorSession.container. */
export function installBookCompletion(
  container: HTMLElement,
  host: CompletionTarget,
  options: BookCompletionOptions = {},
): BookCompletionHandle {
  const doc = (host as CompletionHost | null | undefined)?.activeEditor?.doc;
  const selectionHandle = (host as CompletionHost | null | undefined)?.ui?.selection;
  let disposed = false;
  let suggestion: Suggestion = { kind: 'idle' };
  /** הספרים הפתוחים, הספר שנקרא אחרון ראשון. */
  let books: BookContext[] = [];
  let backgroundLoading = false;
  let currentBook: string | null = null;
  let currentRef: string | null = null;
  let debounceTimer: ReturnType<typeof setTimeout> | undefined;
  let ghostEl: HTMLSpanElement | null = null;
  let lastStateAt = 0;
  let stateRefreshing = false;
  let lastReported: string | null = null;
  /**
   * טוקן ריצה: `evaluate` היא א-סינכרונית (שתי קריאות מנוע), וקריאה איטית
   * שמסיימת אחרי קריאה מאוחרת ומהירה ממנה הייתה דורסת הצעה עדכנית (או ריקון)
   * בהצעה ישנה — בדיוק התסמין שנמדד: טקסט/מיקום מהקשה קודמת, וניקוי המסמך
   * שלא מנקה את ה-ghost.
   */
  let evalToken = 0;
  const pendingPastes = new Set<object>();

  function hideGhost(): void {
    ghostEl?.remove();
    ghostEl = null;
  }

  function clearSuggestion(): void {
    suggestion = { kind: 'idle' };
    hideGhost();
  }

  /**
   * הגופן של הטקסט שמתחת לסמן. `elementFromPoint` ולא `querySelector` על
   * מחלקות פנימיות של המנוע (אסור — ראו engine-boundaries.test.ts): שאילתה
   * לפי קואורדינטה אינה תלויה במבנה ה-DOM הפנימי. `null` = לא נמצא גופן סביר.
   */
  function fontAtCaret(rect: AnchorRectLike): Partial<CSSStyleDeclaration> | null {
    const element = document.elementFromPoint(rect.left, rect.top + rect.height / 2);
    if (!element) return null;
    const style = getComputedStyle(element);
    const size = Number.parseFloat(style.fontSize);
    // גופן גדול משורת הסמן פירושו שנתפס אב־קדמון (מעטפת העורך), לא הטקסט.
    if (!Number.isFinite(size) || size <= 0 || size > rect.height * 1.5) return null;
    return {
      fontFamily: style.fontFamily,
      fontSize: style.fontSize,
      fontWeight: style.fontWeight,
      fontStyle: style.fontStyle,
      letterSpacing: style.letterSpacing,
    };
  }

  function showGhost(text: string): void {
    const rect = selectionHandle?.getAnchorRect?.({ placement: 'end' }) ?? null;
    if (!rect || (rect.width === 0 && rect.height === 0)) {
      hideGhost();
      return;
    }

    if (!ghostEl) {
      ghostEl = document.createElement('span');
      ghostEl.className = GHOST_CLASS;
      Object.assign(ghostEl.style, {
        position: 'fixed',
        display: 'block',
        pointerEvents: 'none',
        color: '#9aa0a6',
        whiteSpace: 'pre',
        direction: 'rtl',
        textAlign: 'right',
        zIndex: '2147483647',
      } satisfies Partial<CSSStyleDeclaration>);
      document.body.appendChild(ghostEl);
    }

    ghostEl.textContent = text;
    // הגובה נלקח מהסמן עצמו ולא מ-`getComputedStyle` של ה-container: ה-container
    // הוא מעטפת העורך, וגובה השורה שלו גדול מזה של הפסקה — וזה בדיוק מה שדחף
    // את ה-ghost שורה אחת מתחת לטקסט. `rect.height` הוא גובה שורת הסמן בפועל.
    const font = fontAtCaret(rect);
    if (font) Object.assign(ghostEl.style, font);
    // גובה הסמן הוא גובה שורה, לא גודל גופן — הקירוב הזה משמש רק כשלא נמצא
    // הגופן האמיתי, ובלעדיו ה-ghost היה בגודל אחר מהטקסט שהמשתמש מקליד.
    else ghostEl.style.fontSize = `${Math.round(rect.height * 0.72)}px`;
    ghostEl.style.lineHeight = `${rect.height}px`;
    ghostEl.style.height = `${rect.height}px`;
    // `right` ולא `left`: בעברית התווים הבאים ממשיכים **משמאל** לסמן, ועוגן
    // מימין עם `direction:rtl` הוא מה שגורם לקופסה לגדול שמאלה מנקודת הסמן,
    // ולא לצוף ימינה מעליו.
    ghostEl.style.right = `${window.innerWidth - rect.left}px`;
    ghostEl.style.left = 'auto';
    ghostEl.style.top = `${rect.top}px`;
  }

  /**
   * טקסט של שורה אחת בספר, כולל ה-pagination הפנימית שלה, וטביעת התצוגה
   * שלה. `null` = כשל. `report` רק לספר שנקרא אחרון: כשל בטאב רקע (ספר מוגן,
   * למשל) אינו מה שהמשתמש עושה עכשיו, ודיווח עליו היה מציף את שורת המצב.
   */
  async function readSectionText(
    ref: BookRef,
    sectionIndex: number,
    report: boolean,
  ): Promise<{ text: string; fingerprint: string } | null> {
    let text = '';
    let fingerprint: string | null = null;
    let cursor: string | undefined;
    for (let page = 0; page < 8 && text.length < CONTEXT_CHAR_CAP; page += 1) {
      const result = await getSectionTextMap({
        bookId: ref.bookId,
        ...(ref.bookUid ? { bookUid: ref.bookUid } : {}),
        sectionIndex,
        layer: 'both',
        cursor,
      });
      if (!result.ok) {
        if (report) options.onStatus?.(result.message, true);
        return null;
      }
      if (!result.value) break;
      // `renderedText` לפני `sourceText`: זו שכבת התצוגה בפועל — אוצריא בונה
      // אותה מהגדרות החלונית החיה (הסתרת ניקוד/טעמים, החלפת שמות קדושים).
      // המשתמש ביקש "הטקסט המוצג", וכשהניקוד מוסתר אצלו `sourceText` היה
      // מנקד לו את מה שהקליד. `layer: 'both'` נשמר כדי שתמיד תהיה נפילה אחורה.
      const pageText = result.value.renderedText || result.value.sourceText || '';
      fingerprint ??= result.value.renderedTextHash ?? pageText;
      text += pageText;
      if (!result.value.hasMore || !result.value.nextCursor) break;
      cursor = result.value.nextCursor;
    }
    return { text, fingerprint: fingerprint ?? '' };
  }

  /** חלון השורות סביב המיקום בקורא — ראו CONTEXT_LINES_BACK/FORWARD. */
  async function loadWindow(book: BookContext, report: boolean): Promise<void> {
    book.loading = true;
    const from = Math.max(0, book.index - CONTEXT_LINES_BACK);
    const to = book.index + CONTEXT_LINES_FORWARD;
    const lines: string[] = [];
    let fingerprint: string | null = null;
    let total = 0;

    for (let index = from; index <= to && total < CONTEXT_CHAR_CAP; index += 1) {
      const line = await readSectionText(book, index, report);
      // ניווט חדש בזמן ההמתנה — מה שנטען שייך למיקום שכבר אינו הנוכחי.
      if (book.dropped || disposed) return;
      // כשל (או סוף הספר) עוצר את החלון, אבל מה שכבר נאסף נשאר שמיש.
      if (line === null) break;
      if (index === book.index) fingerprint = line.fingerprint;
      if (line.text === '') continue;
      lines.push(line.text);
      total += line.text.length;
    }

    // המטמון הקודם משרת עד כאן: טעינה מחדש אחרי שינוי תצוגה אינה משאירה חור.
    book.cache = buildSectionCache(lines.join('\n').slice(0, CONTEXT_CHAR_CAP));
    book.fingerprint = fingerprint;
    book.loading = false;
  }

  /**
   * האם התצוגה של הספר השתנתה מאז שנטען — הסתרת ניקוד או טעמים, החלפת שם
   * הוי"ה. המפתח (ספר + שורה) אינו משתנה כשמסתירים ניקוד, ובלי הבדיקה הזאת
   * ההשלמה המשיכה להציע את הטקסט המנוקד שנטען לפני כן, עד גלילה למקום אחר —
   * הדיווח „מעתיק עם ניקוד ולא לפי איך שזה מוצג”. שורת המרכז מייצגת את החלון:
   * ההגדרות חלות על החלונית כולה.
   */
  async function displayChanged(book: BookContext): Promise<boolean> {
    if (book.loading || book.cache === null) return false;
    const line = await readSectionText(book, book.index, false);
    if (book.dropped || disposed || line === null) return false;
    // שורה שאוצריא שחררה מטאב רקע חוזרת ריקה — זה לא שינוי תצוגה.
    return line.text !== '' && line.fingerprint !== book.fingerprint;
  }

  /** ממזגת את רשימת הספרים הפתוחים: חלון שכבר נטען נשמר, ספר שנסגר יוצא. */
  function syncBooks(refs: BookRef[]): void {
    const next: BookContext[] = [];
    for (const ref of refs) {
      if (next.length >= MAX_OPEN_BOOKS) break;
      const key = bookKey(ref);
      if (next.some((book) => book.key === key)) continue;
      const existing = books.find((book) => book.key === key);
      next.push(
        existing ?? { ...ref, key, cache: null, fingerprint: null, loading: false, dropped: false },
      );
    }
    for (const book of books) {
      if (!next.includes(book)) book.dropped = true;
    }
    books = next;
  }

  /** טעינה או רענון של ספר אחד: חלון חסר נטען, תצוגה שהשתנתה נטענת מחדש. */
  async function refreshBook(book: BookContext, report: boolean): Promise<void> {
    if (book.loading) return;
    if (book.cache === null || (await displayChanged(book))) await loadWindow(book, report);
  }

  /** שאר הספרים — ברקע ובזה אחר זה, כדי לא לעכב את הספר שנקרא אחרון. */
  async function refreshBackgroundBooks(): Promise<void> {
    if (backgroundLoading) return;
    backgroundLoading = true;
    try {
      for (const book of books.slice(1)) {
        if (disposed) return;
        if (!book.dropped) await refreshBook(book, false);
      }
    } finally {
      backgroundLoading = false;
    }
  }

  /**
   * דיווח פעם אחת למעבר מצב בלבד: `refreshReaderState` רצה שוב ושוב בזמן
   * הקלדה, ודיווח בכל פעם היה מציף את שורת המצב באותה הודעה.
   */
  function reportOnce(message: string | null): void {
    if (message === lastReported) return;
    lastReported = message;
    if (message !== null) options.onStatus?.(message, true);
  }

  async function refreshReaderState(): Promise<void> {
    const result = await getCurrentReaderState();
    if (!result.ok) {
      reportOnce(result.message);
      return;
    }
    const state = result.value;
    currentBook = state?.currentBook ?? null;
    currentRef = state?.currentRef ?? null;

    syncBooks(state ? openBookRefs(state) : []);
    if (books.length === 0) {
      reportOnce('השלמה מהספר: לא זוהה ספר פתוח בקורא של אוצריא');
      return;
    }
    reportOnce(null);
    await refreshBook(books[0], true);
    void refreshBackgroundBooks();
  }

  /** בדיקה עצלה וממותנת של מצב הקורא — ראו STATE_TTL_MS. */
  function refreshReaderStateIfStale(): void {
    if (stateRefreshing || Date.now() - lastStateAt < STATE_TTL_MS) return;
    stateRefreshing = true;
    void refreshReaderState().finally(() => {
      stateRefreshing = false;
      lastStateAt = Date.now();
    });
  }

  async function evaluate(): Promise<void> {
    if (disposed || !doc) return;
    // לא ממתינים: ההקשה הזאת תשתמש במטמון שכבר יש, וההקשה הבאה כבר תראה את
    // הספר החדש. המתנה כאן הייתה מוסיפה השהיית RPC לכל הקלדה.
    refreshReaderStateIfStale();

    const token = ++evalToken;
    const snapshot = await readTypedSnapshot(doc);
    // ריצה מאוחרת יותר כבר התחילה בזמן שחיכינו למנוע — התוצאה שלנו ישנה.
    if (token !== evalToken) return;
    if (!snapshot || disposed) {
      clearSuggestion();
      return;
    }

    // אזכור המקור מקבל את כל הקשות ה-@; אחרת שתי שכבות נפרדות היו יכולות
    // להציג הצעה ולתפוס Tab מול הרשימה של at-mention-overlay.
    if (parseAtTrigger(snapshot.rawBeforeText)) {
      clearSuggestion();
      return;
    }

    const titleMatch = currentBook
      ? matchBookTitle(currentBook, currentRef, snapshot)
      : null;
    if (titleMatch) {
      suggestion = {
        kind: 'suggesting',
        ghostText: titleMatch.completionText,
        insertText: titleMatch.completionText,
        replaceStart: snapshot.replaceStart,
        cursorOffset: snapshot.cursorOffset,
        blockId: snapshot.blockId,
        story: snapshot.story,
        continueFrom: null,
        source: null,
      };
      showGhost(titleMatch.completionText);
      return;
    }

    const found = matchAcrossBooks(
      books.map((book) => book.cache ?? EMPTY_CACHE),
      snapshot,
      { maxContextWords: CONTEXT_WORDS, wordsToShow: WORDS_TO_SHOW },
    );
    if (found) {
      const { match: bookMatch } = found;
      const cache = books[found.bookIndex].cache ?? EMPTY_CACHE;
      // לשורה אחת: חלון ההקשר מחבר שורות ב-`\n`, והצעה שחוצה שורה בספר
      // הייתה נכנסת למסמך עם שבר שורה בתוכה — `insert` עם `type: 'text'`
      // מכניס טקסט לפסקה. אותו נימוק בדיוק כמו בציטוט (otzaria-reader.ts).
      const ghostText = normalizeSelectedText(bookMatch.text);
      if (ghostText === '') {
        clearSuggestion();
        return;
      }
      // ההחלפה נמתחת אחורה גם על מילות ההקשר שתאמו: המשתמש הקליד אותן בלי
      // ניקוד, וללא זה יצא משפט שרק סופו מנוקד — התסמין שנמדד ("כתבתי שתי
      // מילים, ניקד רק את השנייה").
      const used = Math.min(bookMatch.contextWordsUsed, snapshot.precedingWordStarts.length);
      const withContext =
        used > 0 ? sliceWords(cache, bookMatch.matchedWordIndex - used, used + WORDS_TO_SHOW) : null;
      const insertText = withContext ? normalizeSelectedText(withContext.text) : ghostText;
      const replaceStart =
        withContext && insertText !== ''
          ? snapshot.precedingWordStarts[snapshot.precedingWordStarts.length - used]
          : snapshot.replaceStart;

      suggestion = {
        kind: 'suggesting',
        ghostText,
        insertText: insertText === '' ? ghostText : insertText,
        replaceStart,
        cursorOffset: snapshot.cursorOffset,
        blockId: snapshot.blockId,
        story: snapshot.story,
        continueFrom: bookMatch.nextWordIndex,
        source: cache,
      };
      showGhost(ghostText);
      return;
    }

    /*
     * הרשימות הסטטיות — ביטויים תלמודיים ושמות מחברים — הן ה-fallback של
     * הספר, ולפני ראשי-התיבות.
     *
     * אין כאן מתיחה אחורה על מילות ההקשר כמו בספר: הרשומות אינן מנוקדות,
     * המשתמש הקליד בדיוק את מה שכתוב בהן, ואין מה לתקן בדיעבד. לכן ההחלפה
     * היא על המילה החלקית בלבד, ואין `continueFrom` — ההשלמה כבר מגיעה עד סוף
     * הרשומה, ומעבר לה אין מה להציע.
     */
    const staticSources = await loadStaticSources();
    if (token !== evalToken || disposed) return;
    if (staticSources) {
      const staticMatch = matchStaticCompletion(staticSources, snapshot);
      const ghostText = staticMatch ? normalizeSelectedText(staticMatch.text) : '';
      if (ghostText !== '') {
        suggestion = {
          kind: 'suggesting',
          ghostText,
          insertText: ghostText,
          replaceStart: snapshot.replaceStart,
          cursorOffset: snapshot.cursorOffset,
          blockId: snapshot.blockId,
          story: snapshot.story,
          continueFrom: null,
          source: null,
        };
        showGhost(ghostText);
        return;
      }
    }

    // ראשי-תיבות הם מילה שלמה, ולכן נבדקים רק אחרי רווח ולא כהשלמת prefix.
    const lastWord =
      snapshot.partialWord === '' ? (snapshot.precedingWords[snapshot.precedingWords.length - 1] ?? null) : null;
    // הצורה נבדקת **לפני** הטעינה: אחרת כל מילה רגילה שההשלמה מהספר לא מצאה
    // לה המשך הייתה מזריקה נכס של כמעט מגה-בייט, אחרי כל השהיה.
    if (lastWord !== null && looksLikeAcronym(lastWord)) {
      const dictionary = await loadAcronymDictionary();
      if (token !== evalToken || disposed) return;
      const expansion = dictionary?.lookup(lastWord) ?? null;
      if (expansion) {
        // רווח מפריד רק כשאין אחד. הר"ת נבדק גם אחרי רווח שהוקלד וגם כשהסמן
        // צמוד לאות האחרונה, והפרדה עיוורת הכניסה רווח כפול (נמדד בשער).
        const ghostText = /\s$/.test(snapshot.rawBeforeText) ? expansion : ` ${expansion}`;
        suggestion = {
          kind: 'suggesting',
          ghostText,
          insertText: ghostText,
          replaceStart: snapshot.cursorOffset,
          cursorOffset: snapshot.cursorOffset,
          blockId: snapshot.blockId,
          story: snapshot.story,
          continueFrom: null,
          source: null,
        };
        showGhost(ghostText);
        return;
      }
    }

    clearSuggestion();
  }

  function scheduleEvaluate(): void {
    if (pendingPastes.size > 0) return;
    if (debounceTimer !== undefined) clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => void evaluate(), INPUT_DEBOUNCE_MS);
  }

  async function accept(): Promise<void> {
    if (suggestion.kind !== 'suggesting' || !doc || typeof doc.insert !== 'function') return;
    const { insertText, replaceStart, cursorOffset, blockId, story, continueFrom, source } = suggestion;
    clearSuggestion();

    const target: SelectionTargetLike = {
      kind: 'selection',
      start: pointAt(blockId, replaceStart, story),
      end: pointAt(blockId, cursorOffset, story),
      ...(story ? { story } : {}),
    };

    let receipt: DocReceipt;
    try {
      receipt = await doc.insert({ value: insertText, type: 'text', target });
    } catch (error) {
      console.warn('[otzaria-word] השלמה מהספר: הכנסת הטקסט נכשלה', error);
      return;
    }
    if (receipt?.success === false || disposed) return;

    /*
     * להשלמה בלי המשך — ר"ת, רשימה סטטית, שם ספר — ההכנסה היא סוף העניין,
     * וכל הערכה שנקבעה או שרצה לפניה מסתמכת על הטקסט שלפניה: היא הייתה
     * מציעה שוב את אותה השלמה בדיוק. נמדד: ה-ghost חזר מיד אחרי Tab ונשאר
     * על המסך עד ההקשה הבאה. הטוקן פוסל את מי שבאוויר, וה-timer את מי שעוד
     * לא יצא.
     *
     * **ורק שם.** להשלמה מהספר יש המשך, וההערכה המושהית היא מי שמצייר אותו
     * בפועל: ה-`requestAnimationFrame` שלמטה נופל כשהמנוע עדיין לא מחזיר
     * `getAnchorRect` מיד אחרי ההכנסה, ואין לו ניסיון שני. ביטול עיוור של
     * ההערכה הותיר הצעה חיה ובלתי נראית — Tab נתפס והכניס חמש מילים בלי
     * ששום דבר הוצג (נמדד ב-scripts/qa/book-completion-qa.mjs).
     */
    if (continueFrom === null || source === null) {
      evalToken += 1;
      if (debounceTimer !== undefined) clearTimeout(debounceTimer);
      clearSuggestion();
      return;
    }

    const next = sliceWords(source, continueFrom, WORDS_TO_SHOW);
    // רווח מפריד ולא הצמדה: ההשלמה הקודמת הסתיימה במילה, וה-5 הבאות הן
    // המשך המשפט. `normalizeSelectedText` מטפל גם כאן בשבר שורה מהספר.
    const nextText = next ? ` ${normalizeSelectedText(next.text)}` : '';
    if (next && nextText.trim() !== '') {
      const newCursor = replaceStart + insertText.length;
      suggestion = {
        kind: 'suggesting',
        ghostText: nextText,
        insertText: nextText,
        replaceStart: newCursor,
        cursorOffset: newCursor,
        blockId,
        story,
        continueFrom: next.nextWordIndex,
        source,
      };
      requestAnimationFrame(() => {
        if (!disposed && suggestion.kind === 'suggesting') showGhost(suggestion.ghostText);
      });
    }
  }

  const onPasteStarted = (event: Event): void => {
    if (!(event instanceof CustomEvent) || typeof event.detail?.operation !== 'object') return;
    pendingPastes.add(event.detail.operation);
    evalToken += 1;
    if (debounceTimer !== undefined) clearTimeout(debounceTimer);
    clearSuggestion();
  };
  const onPasteFinished = (event: Event): void => {
    if (!(event instanceof CustomEvent) || !pendingPastes.delete(event.detail?.operation)) return;
    if (!disposed && pendingPastes.size === 0 && event.detail?.evaluate !== false) scheduleEvaluate();
  };

  const onInput = (): void => scheduleEvaluate();
  /*
   * כל הקשה אחרת פוסלת את ההצעה **מיד**, ולא אחרי ההשהיה. ה-ghost מצויר
   * בעמדת הסמן של ההערכה האחרונה, והאותיות החדשות מצוירות בדיוק שם — מתחתיו.
   * ההערכה הבאה מחכה ל-debounce, וקריאות המנוע בתוכה מחכות לשקט בהקלדה, ולכן
   * בהקלדה רציפה ה-ghost נשאר קפוא מעל הטקסט לאורך כל הרצף. נמדד: 11 מתוך 18
   * דגימות באמצע הקלדה הראו ghost שאינו בעמדת הסמן. הטוקן פוסל גם הערכה
   * שכבר באוויר עם הקשר שקדם להקשה.
   */
  const invalidate = (): void => {
    if (suggestion.kind === 'idle' && !ghostEl) return;
    evalToken += 1;
    clearSuggestion();
  };
  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.key === 'Tab' && suggestion.kind === 'suggesting') {
      event.preventDefault();
      event.stopPropagation();
      void accept();
      return;
    }
    // מקש שינוי לבדו אינו משנה טקסט ואינו מזיז את הסמן.
    if (event.key === 'Shift' || event.key === 'Control' || event.key === 'Alt' || event.key === 'Meta') return;
    invalidate();
  };
  // גלילה שהסמן גרר אחריו אינה „המשתמש גלל משם” — ראו `isCaretFollowScroll`.
  // האירוע עצמו עובר, ולא המיכל: ההבחנה נשענת על ה-`target` שלו, ולכן גלילה
  // של צאצא (שנשמעת כאן על שלב ה-capture) ממשיכה להסתיר את ה-ghost.
  const onScroll = (event: Event): void => {
    if (isCaretFollowScroll(event)) return;
    hideGhost();
  };

  container.addEventListener(PASTE_STARTED, onPasteStarted);
  container.addEventListener(PASTE_FINISHED, onPasteFinished);
  container.addEventListener('input', onInput);
  container.addEventListener('keyup', onInput);
  container.addEventListener('keydown', onKeyDown, true);
  // הדבקה מהתפריט והכתבה אינן עוברות ב-keydown.
  container.addEventListener('beforeinput', invalidate, true);
  container.addEventListener('scroll', onScroll, true);

  const offRefChanged = on('reader.current_ref_changed', (detail) => {
    currentBook = detail.currentBook;
    currentRef = detail.currentRef;
    clearSuggestion();
    void refreshReaderState();
  });
  const offBookChanged = on('reader.current_book_changed', () => {
    clearSuggestion();
    void refreshReaderState();
  });

  void refreshReaderState();

  return {
    dispose() {
      disposed = true;
      if (debounceTimer !== undefined) clearTimeout(debounceTimer);
      clearSuggestion();
      container.removeEventListener(PASTE_STARTED, onPasteStarted);
      container.removeEventListener(PASTE_FINISHED, onPasteFinished);
      container.removeEventListener('input', onInput);
      container.removeEventListener('keyup', onInput);
      container.removeEventListener('keydown', onKeyDown, true);
      container.removeEventListener('beforeinput', invalidate, true);
      container.removeEventListener('scroll', onScroll, true);
      offRefChanged();
      offBookChanged();
    },
  };
}
