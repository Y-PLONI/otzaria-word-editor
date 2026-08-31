/**
 * התאמת טקסט מוקלד לטקסט הספר הפתוח בקורא, לצורך "השלמה מהספר".
 *
 * מודול טהור (בלי DOM, בלי RPC): מקבל את טקסט ה-section הנוכחי ואת ההקשר
 * המוקלד סביב הסמן, ומחזיר השלמה או `null`. כל הפונקציות כאן נבדקות ביחידה.
 *
 * ## למה השוואה על טקסט מנורמל, וגזירה על טקסט מקור
 *
 * טקסט הספר מנוקד ומוטעם; מה שהמשתמש מקליד — לרוב לא. השוואה ישירה הייתה
 * כמעט תמיד נכשלת. הפתרון: בונים אינדקס בין הטקסט המנורמל (בלי ניקוד/טעמים)
 * לבין ההיסטים בטקסט המקור (`toOriginal`), מחפשים התאמה במנורמל, וחוזרים
 * למקור כדי לגזור את ההשלמה. ה-`accept` מחליף את המילה החלקית שהמשתמש הקליד
 * בטקסט המקור המלא (עם הניקוד) ולא רק מוסיף אחריה — אחרת שרשור „בר” (בלי
 * ניקוד) + „ֵאשִׁית” (עם ניקוד, מהמקום שההתאמה נמצאה) היה יוצר תערובת שבורה.
 *
 * ## למה השלמה בגבול מילה בלבד ב-`precedingWords`, ו-loop יורד ב-context
 *
 * השוואה נגד מילה בודדת חלקית ("בר") הייתה תואמת עשרות מקומות בפרק. הקשר של
 * עד 3 מילים קודמות מכוון את ההתאמה למקום הנכון; ניסיון בהקשר ארוך ואז קצר
 * (ולא הפוך) מעדיף את ההתאמה המדויקת ביותר שיש.
 */
import { WORD_INNER, WORD_MARK } from './word-selection';

const STRIP_MARKS = new RegExp(WORD_MARK.source, 'g');

/** מסירה ניקוד/טעמים בלבד. אותיות, גרש וגרשיים נשארים — הם חלק מהמילה. */
export function stripDiacritics(text: string): string {
  return text.replace(STRIP_MARKS, '');
}

/** נרמול מילה בודדת מוקלדת: ניקוד (אם הוקלד בכל זאת) והיקפי רווח. */
export function normalizeTypedWord(text: string): string {
  return stripDiacritics(text).trim();
}

/** גבול מילה, כמו ב-word-selection.ts: כל מה ש-WORD_INNER תופס. */
function isWordChar(ch: string | undefined): boolean {
  return ch !== undefined && WORD_INNER.test(ch);
}

/** גבולות המילים בטקסט מנורמל, כהיסטים ]start, end[ בתוך אותו טקסט. */
export interface WordSpan {
  start: number;
  end: number;
}

export interface SectionWordCache {
  sourceText: string;
  normalizedText: string;
  /** `toOriginal[i]` הוא ההיסט במקור של `normalizedText[i]`. */
  toOriginal: number[];
  words: WordSpan[];
}

/**
 * בונה את המטמון: נרמול + מיפוי היסטים + פירוק למילים. פעם אחת ל-section,
 * לא בכל הקשה — ראו book-completion-overlay.ts.
 */
export function buildSectionCache(sourceText: string): SectionWordCache {
  let normalizedText = '';
  const toOriginal: number[] = [];

  for (let i = 0; i < sourceText.length; i += 1) {
    const ch = sourceText[i];
    if (WORD_MARK.test(ch)) continue;
    normalizedText += ch;
    toOriginal.push(i);
  }

  const words: WordSpan[] = [];
  let start = -1;
  for (let i = 0; i <= normalizedText.length; i += 1) {
    const atWordChar = isWordChar(normalizedText[i]);
    if (atWordChar && start === -1) start = i;
    else if (!atWordChar && start !== -1) {
      words.push({ start, end: i });
      start = -1;
    }
  }

  return { sourceText, normalizedText, toOriginal, words };
}

/** מילה מנורמלת בהיסט `wordIndex`. */
function wordAt(cache: SectionWordCache, wordIndex: number): string {
  const span = cache.words[wordIndex];
  return cache.normalizedText.slice(span.start, span.end);
}

export interface WordSlice {
  /** הטקסט המדויק מהמקור, כולל ניקוד ופיסוק ביניים. */
  text: string;
  /** אינדקס המילה הבאה בתור — ה"המשך" לסבב ההשלמה הבא. */
  nextWordIndex: number;
}

/** גוזרת `count` מילים מהמקור, החל מ-`fromWordIndex`. `null` כשאין עוד. */
export function sliceWords(
  cache: SectionWordCache,
  fromWordIndex: number,
  count: number,
): WordSlice | null {
  if (fromWordIndex < 0 || fromWordIndex >= cache.words.length) return null;

  const lastIndex = Math.min(fromWordIndex + count, cache.words.length) - 1;
  const startNorm = cache.words[fromWordIndex].start;
  const endNorm = cache.words[lastIndex].end;

  const originalStart = cache.toOriginal[startNorm];
  // הקצה הסוגר: ההיסט שאחרי התו המקורי האחרון של המילה. `toOriginal` אינו
  // מוגדר בדיוק ב-`endNorm` (הוא מצביע על תו שאחרי המילה, שעשוי היה תו ניקוד
  // שהוסר) — ולכן לוקחים את התו המקורי האחרון בפועל ומוסיפים 1.
  const originalEnd = cache.toOriginal[endNorm - 1] + 1;

  return {
    text: cache.sourceText.slice(originalStart, originalEnd),
    nextWordIndex: lastIndex + 1,
  };
}

export interface TypedContext {
  /** מילים שהושלמו לפני המילה הנוכחית, מהרחוקה לקרובה. */
  precedingWords: string[];
  /** מה שהוקלד מהמילה הנוכחית. ריק = הסמן בגבול מילה (אחרי רווח/תחילת פסקה). */
  partialWord: string;
}

export interface BookCompletionMatch extends WordSlice {
  /** אינדקס המילה שממנה מתחילה ההשלמה (כדי לדעת מה בדיוק הוחלף). */
  matchedWordIndex: number;
  /**
   * כמה מילות הקשר קודמות נמצאו תואמות בפועל. הקורא משתמש בזה כדי להרחיב את
   * ההחלפה גם עליהן: בספר מנוקד המשתמש הקליד אותן בלי ניקוד, והשארתן כמות
   * שהן הייתה משאירה משפט חצי־מנוקד.
   */
  contextWordsUsed: number;
}

export interface MatchOptions {
  /** כמה מילות הקשר קודמות לנסות, מהארוך לקצר. ברירת מחדל 3. */
  maxContextWords?: number;
  /** כמה מילים להציע בכל סבב. ברירת מחדל 5. */
  wordsToShow?: number;
  /** אורך מזערי למילה חלקית שאין לפניה שום הקשר תואם. ברירת מחדל 3. */
  minStandalonePartial?: number;
}

/**
 * האם יש כאן די אות כדי להציע.
 *
 * נמדד: בלי הסף הזה אות בודדת בלי הקשר („ח” אחרי „קנה די”, שלא נמצא בספר)
 * התאימה למילה הראשונה בספר שמתחילה באות הזאת, וההצעה שהוצגה הייתה טקסט
 * אקראי מהעמוד. הקשר תואם הוא האות החזקה, ובלעדיו נדרשות שלוש אותיות.
 */
function hasEnoughSignal(contextLen: number, partial: string, minStandalone: number): boolean {
  if (contextLen >= 2) return true;
  if (contextLen === 1) return partial.length >= 1;
  return partial.length >= minStandalone;
}

/**
 * מחפשת התאמה בטקסט הספר להקשר המוקלד. `null` פירושו שאין מה להציע —
 * לא מפני שההקלדה שגויה, אלא מפני שהיא לא נמצאת (עדיין) בטקסט הזה.
 */
export function matchAtCursor(
  cache: SectionWordCache,
  context: TypedContext,
  options: MatchOptions = {},
): BookCompletionMatch | null {
  const maxContextWords = options.maxContextWords ?? 3;
  const wordsToShow = options.wordsToShow ?? 5;
  const minStandalone = options.minStandalonePartial ?? 3;

  const queryWords = context.precedingWords
    .map(normalizeTypedWord)
    .filter((word) => word !== '')
    .slice(-maxContextWords);
  const partial = normalizeTypedWord(context.partialWord);

  for (let contextLen = queryWords.length; contextLen >= 0; contextLen -= 1) {
    if (!hasEnoughSignal(contextLen, partial, minStandalone)) continue;
    const slice = queryWords.slice(queryWords.length - contextLen);
    const found = matchWithContext(cache, slice, partial, wordsToShow);
    if (found) return found;
  }
  return null;
}

function matchWithContext(
  cache: SectionWordCache,
  contextWords: string[],
  partial: string,
  wordsToShow: number,
): BookCompletionMatch | null {
  const targetOffset = contextWords.length;
  for (let w = 0; w + targetOffset < cache.words.length; w += 1) {
    let matches = true;
    for (let k = 0; k < contextWords.length; k += 1) {
      if (wordAt(cache, w + k) !== contextWords[k]) {
        matches = false;
        break;
      }
    }
    if (!matches) continue;

    const targetIndex = w + targetOffset;
    if (partial !== '' && !wordAt(cache, targetIndex).startsWith(partial)) continue;

    const slice = sliceWords(cache, targetIndex, wordsToShow);
    if (!slice) continue;
    return { ...slice, matchedWordIndex: targetIndex, contextWordsUsed: contextWords.length };
  }
  return null;
}

/** התאמת שם הספר הנוכחי + מיקום, כשהמשתמש מתחיל להקליד את שם הספר. */
export interface BookTitleMatch {
  /** מה שיש להוסיף/להחליף: שארית השם, ואז המיקום. */
  completionText: string;
}

/**
 * `currentBook`/`currentRef` הם טקסט רגיל בלי ניקוד ברוב המקרים, ולכן כאן
 * ההשוואה על התווים המקוריים (בלי מיפוי `toOriginal`) — פשוט ומספיק.
 */
export function matchBookTitle(
  currentBook: string,
  currentRef: string | null,
  context: TypedContext,
): BookTitleMatch | null {
  if (!currentBook) return null;

  const query = [...context.precedingWords, context.partialWord]
    .map((word) => stripDiacritics(word).trim())
    .filter((word) => word !== '')
    .join(' ');
  if (query === '') return null;

  const title = stripDiacritics(currentBook).trim();
  if (!title.toLowerCase().startsWith(query.toLowerCase())) return null;

  const ref = currentRef?.trim();
  const rest = title.slice(query.length);
  // ה-conditional: כש-`rest` ריק (הכותרת כבר הוקלדה במלואה) הרווח שב-
  // `${rest} ${ref}` הוא בדיוק המפריד החסר בין הכותרת שכבר בדף לבין המיקום.
  const completionText = ref ? `${rest} ${ref}` : rest;
  if (completionText === '') return null;

  return { completionText };
}
