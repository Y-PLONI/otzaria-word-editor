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
  matchAtCursor,
  matchBookTitle,
  sliceWords,
  type SectionWordCache,
} from './book-completion';
import {
  getCurrentReaderState,
  getSectionTextMap,
  normalizeSelectedText,
} from '../host/otzaria-reader';
import { on } from '../host/otzaria-client';

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
  let cache: SectionWordCache | null = null;
  let contextKey: string | null = null;
  let currentBook: string | null = null;
  let currentBookId: string | null = null;
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

  /** טקסט של שורה אחת בספר, כולל ה-pagination הפנימית שלה. `null` = כשל. */
  async function readSectionText(bookId: string, sectionIndex: number): Promise<string | null> {
    let text = '';
    let cursor: string | undefined;
    for (let page = 0; page < 8 && text.length < CONTEXT_CHAR_CAP; page += 1) {
      const result = await getSectionTextMap({ bookId, sectionIndex, layer: 'both', cursor });
      if (!result.ok) {
        options.onStatus?.(result.message, true);
        return null;
      }
      if (!result.value) break;
      // `renderedText` לפני `sourceText`: זו שכבת התצוגה בפועל — אוצריא בונה
      // אותה מהגדרות החלונית החיה (הסתרת ניקוד/טעמים, החלפת שמות קדושים).
      // המשתמש ביקש "הטקסט המוצג", וכשהניקוד מוסתר אצלו `sourceText` היה
      // מנקד לו את מה שהקליד. `layer: 'both'` נשמר כדי שתמיד תהיה נפילה אחורה.
      text += result.value.renderedText || result.value.sourceText || '';
      if (!result.value.hasMore || !result.value.nextCursor) break;
      cursor = result.value.nextCursor;
    }
    return text;
  }

  /** חלון השורות סביב המיקום בקורא — ראו CONTEXT_LINES_BACK/FORWARD. */
  async function ensureContext(bookId: string, centerIndex: number): Promise<void> {
    const key = `${bookId}:${centerIndex}`;
    if (key === contextKey) return;
    contextKey = key;
    cache = null;

    const from = Math.max(0, centerIndex - CONTEXT_LINES_BACK);
    const to = centerIndex + CONTEXT_LINES_FORWARD;
    const lines: string[] = [];
    let total = 0;

    for (let index = from; index <= to && total < CONTEXT_CHAR_CAP; index += 1) {
      const line = await readSectionText(bookId, index);
      // ניווט חדש בזמן ההמתנה — מה שנטען שייך למיקום שכבר אינו הנוכחי.
      if (key !== contextKey || disposed) return;
      // כשל (או סוף הספר) עוצר את החלון, אבל מה שכבר נאסף נשאר שמיש.
      if (line === null) break;
      if (line === '') continue;
      lines.push(line);
      total += line.length;
    }

    cache = buildSectionCache(lines.join('\n').slice(0, CONTEXT_CHAR_CAP));
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
    currentBookId = state?.currentBookId ?? null;

    if (currentBookId !== null && typeof state?.currentIndex === 'number') {
      reportOnce(null);
      await ensureContext(currentBookId, state.currentIndex);
    } else {
      contextKey = null;
      cache = null;
      reportOnce('השלמה מהספר: לא זוהה ספר פתוח בקורא של אוצריא');
    }
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
      };
      showGhost(titleMatch.completionText);
      return;
    }

    if (cache) {
      const bookMatch = matchAtCursor(cache, snapshot, {
        maxContextWords: CONTEXT_WORDS,
        wordsToShow: WORDS_TO_SHOW,
      });
      if (bookMatch) {
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
        };
        showGhost(ghostText);
        return;
      }
    }

    clearSuggestion();
  }

  function scheduleEvaluate(): void {
    if (debounceTimer !== undefined) clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => void evaluate(), INPUT_DEBOUNCE_MS);
  }

  async function accept(): Promise<void> {
    if (suggestion.kind !== 'suggesting' || !doc || typeof doc.insert !== 'function') return;
    const { insertText, replaceStart, cursorOffset, blockId, story, continueFrom } = suggestion;
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

    if (continueFrom !== null && cache) {
      const next = sliceWords(cache, continueFrom, WORDS_TO_SHOW);
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
        };
        requestAnimationFrame(() => {
          if (!disposed && suggestion.kind === 'suggesting') showGhost(suggestion.ghostText);
        });
      }
    }
  }

  const onInput = (): void => scheduleEvaluate();
  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.key !== 'Tab' || suggestion.kind !== 'suggesting') return;
    event.preventDefault();
    event.stopPropagation();
    void accept();
  };
  const onScroll = (): void => hideGhost();

  container.addEventListener('input', onInput);
  container.addEventListener('keyup', onInput);
  container.addEventListener('keydown', onKeyDown, true);
  container.addEventListener('scroll', onScroll, true);

  const offRefChanged = on('reader.current_ref_changed', (detail) => {
    currentBook = detail.currentBook;
    currentBookId = detail.currentBookId;
    currentRef = detail.currentRef;
    clearSuggestion();
    if (currentBookId !== null) void ensureContext(currentBookId, detail.currentIndex);
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
      container.removeEventListener('input', onInput);
      container.removeEventListener('keyup', onInput);
      container.removeEventListener('keydown', onKeyDown, true);
      container.removeEventListener('scroll', onScroll, true);
      offRefChanged();
      offBookChanged();
    },
  };
}
