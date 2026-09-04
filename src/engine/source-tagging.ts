/**
 * תיוג מקורות עם `@`: זיהוי token פתוח, פענוח סימון עמוד, והתאמה לערך
 * תוכן-עניינים אמיתי — לשם בניית קישור `otzaria://open/book/<id>?index=<n>`.
 *
 * מודול טהור (בלי RPC, בלי DOM): מקבל טקסט/רשימות שכבר נשלפו על ידי
 * `host/otzaria-library.ts`, ומחזיר התאמה או `null`. כל הפונקציות כאן
 * נבדקות ביחידה — ראו tests/unit/source-tagging.test.ts.
 *
 * ## הכלל שנבדק ואמור להישמר: אין סימון עמוד → אין קישור
 *
 * "פסחים יג" בלי `.`/`:`, בלי `א'`/`ב'` ובלי `ע"א`/`ע"ב` הוא התייחסות
 * עמומה (יכול להיות דף, סימן, פרק) ולא הופך לקישור. רק שלוש הצורות האלה
 * מזוהות, ורק הן מנורמלות לטופס הקנוני `"דף <אותיות><.|:>"` — בדיוק הצורה
 * שבה `tocEntry` של אוצריא עצמה שומרת ערכי דף/עמוד (נמדד: `seforim.db`,
 * ספר יומא, `tocEntry.text = 'דף יג.' / 'דף יג:'`, `level = 1`).
 */
import { stripDiacritics } from './book-completion';

/**
 * שם ספר יכול להיות בן כמה מילים ("בבא קמא"), ואחריו רווח ואז סימון עמוד —
 * כלומר רווח בודד **בתוך** ה-query הוא חלק לגיטימי מ-token פתוח, ולא סוגר
 * אותו. מה שכן סוגר: שבר שורה, `@` נוסף (תיוג חדש התחיל), או רווח כפול
 * (המשתמש כבר המשיך הלאה בפרוזה רגילה). אורך מרבי מונע סריקה בלתי חסומה
 * על "@ שנשאר פתוח" לאורך פסקה שלמה שלא הייתה מיועדת כתיוג בכלל.
 */
const TOKEN_BREAK = /\n|\r|@|  /;
const MAX_QUERY_LENGTH = 60;

export interface OpenTagToken {
  /** ההיסט (בתוך הטקסט שנבדק) שבו נמצא ה-`@` הפותח את ה-token. */
  atOffset: number;
  /** מה שהוקלד אחרי ה-`@`, עד הסמן. עשוי לכלול רווחים בודדים בתוכו. */
  query: string;
}

/**
 * מאתרת token פתוח של `@` שמסתיים בדיוק בסמן (סוף `textBeforeCursor`).
 *
 * "סגירה" בפועל (המרה לקישור, או ויתור על ההצעה) אינה כאן — היא תלויה
 * בתוצאה של `resolveSourceTag` על אותו `query`. הפונקציה הזו רק מגבילה את
 * טווח החיפוש: שבר שורה, `@` נוסף, רווח כפול או אורך חריג — `null`.
 */
export function findOpenTagToken(textBeforeCursor: string): OpenTagToken | null {
  const atOffset = textBeforeCursor.lastIndexOf('@');
  if (atOffset === -1) return null;

  const query = textBeforeCursor.slice(atOffset + 1);
  if (query.length > MAX_QUERY_LENGTH) return null;
  if (TOKEN_BREAK.test(query)) return null;

  return { atOffset, query };
}

export interface BookNameCandidate {
  bookId: string;
  title: string;
}

export interface BookNameMatch extends BookNameCandidate {
  /** מה שכבר הוקלד משם הספר (התחילית התואמת). */
  matchedPrefix: string;
  /** מה שנשאר להשלים משם הספר, לפני שמגיעים לסימון העמוד. */
  restOfTitle: string;
}

/**
 * ספרים ששם שלהם מתחיל ב-`query` (בלי ניקוד, לא תלוי-רישיות). ממוינת לפי
 * אורך כותרת עולה — כותרת קצרה יותר שמתאימה היא ההתאמה החדה יותר, ולכן
 * "יומא" קודמת ל"יומא טוב" (אם יש כזה) כשה-query הוא "יומא".
 */
export function matchBookNames(
  candidates: readonly BookNameCandidate[],
  query: string,
): BookNameMatch[] {
  const needle = stripDiacritics(query).trim();
  if (needle === '') return [];

  const matches: BookNameMatch[] = [];
  for (const candidate of candidates) {
    const title = stripDiacritics(candidate.title).trim();
    if (!title.toLowerCase().startsWith(needle.toLowerCase())) continue;
    matches.push({
      ...candidate,
      matchedPrefix: title.slice(0, needle.length),
      restOfTitle: title.slice(needle.length),
    });
  }
  return matches.sort((a, b) => a.title.length - b.title.length);
}

/**
 * ההפך מ-`matchBookNames`: כאן `query` כבר ארוך משם הספר (יש אחריו סימון
 * עמוד), ומחפשים איזה שם ספר ידוע הוא **תחילית** שלו — `query` מתחיל בשם
 * הספר ואז רווח. "יומא יג:" תואם "יומא"; "יומא" לבדו (בלי מה שאחריו) לא —
 * אין עדיין מה להשלים. בין שני מועמדים תואמים נבחר בעל הכותרת הארוכה יותר,
 * כדי ש-"בבא קמא" לא ייבלע על ידי "בבא" אם שניהם קיימים.
 */
export function findBookNameInQuery(
  candidates: readonly BookNameCandidate[],
  query: string,
): BookNameMatch | null {
  const normalizedQuery = stripDiacritics(query).trim();
  if (normalizedQuery === '') return null;

  let best: BookNameMatch | null = null;
  for (const candidate of candidates) {
    const title = stripDiacritics(candidate.title).trim();
    if (title === '' || !normalizedQuery.startsWith(`${title} `)) continue;
    if (!best || title.length > best.matchedPrefix.length) {
      best = { ...candidate, matchedPrefix: title, restOfTitle: '' };
    }
  }
  return best;
}

export type Amud = 'a' | 'b';

export interface PageMarker {
  /** האותיות (גימטריה) של הדף, בדיוק כפי שהוקלדו — "יג". */
  letters: string;
  amud: Amud;
  /** אורך הטקסט הגולמי שזוהה כסימון עמוד, מ-`letters` ועד סוף הסימן. */
  consumedLength: number;
}

/**
 * מזהה סימון עמוד בתחילת `raw` (שרק הורחק ממנו הרווח המוביל): גימטריה,
 * ואחריה **מיד** אחד מ-`.`/`:`/`א'`/`ב'`/`ע"א`/`ע"ב` (עם/בלי רווח לפני
 * הסימן המילולי). `null` = אין סימון מוכר, ולכן אין קישור.
 *
 * הגימטריה עצמה אינה מאומתת כמספר תקין (לא בודקים "אין יותר מארבע יו"דים
 * ברצף" וכדומה) — התאמה לרשימת ה-toc בפועל היא האימות; אות שלא קיימת
 * כדף פשוט לא תימצא שם.
 */
export function parsePageMarker(raw: string): PageMarker | null {
  const match = /^([א-ת]+)(\s*)(\.|:|א'|ב'|ע"א|ע"ב)/.exec(raw);
  if (!match) return null;

  const [full, letters, , marker] = match;
  const amud: Amud = marker === '.' || marker === "א'" || marker === 'ע"א' ? 'a' : 'b';

  return { letters: letters!, amud, consumedLength: full!.length };
}

/** הטופס הקנוני, זהה לצורת ה-`tocEntry` של אוצריא: `"דף יג:"`. */
export function canonicalDafText(letters: string, amud: Amud): string {
  return `דף ${letters}${amud === 'a' ? '.' : ':'}`;
}

export interface TocEntryLike {
  text: string;
  index: number;
  level: number;
}

export interface TocIndex {
  /** טקסט קנוני (`"דף יג:"`) → `index` של הערך המתאים ב-toc. */
  byCanonicalText: ReadonlyMap<string, number>;
}

/** בונה אינדקס חיפוש מהיר מתוצאת `library.getBookToc`. */
export function buildTocIndex(entries: readonly TocEntryLike[]): TocIndex {
  const byCanonicalText = new Map<string, number>();
  for (const entry of entries) {
    if (!byCanonicalText.has(entry.text)) byCanonicalText.set(entry.text, entry.index);
  }
  return { byCanonicalText };
}

export interface ResolvedTag {
  /** מספר ה-index של ערך ה-toc שנמצא, לבניית `otzaria://open/book/<id>?index=`. */
  tocIndex: number;
  /** אורך הטקסט (מ-`@` ועד סוף הסימון) שיש לעטוף כקישור. */
  consumedLength: number;
}

/**
 * הפתרון המלא: `query` הוא מה שהוקלד אחרי `@` (כולל שם הספר), `match` הוא
 * שם הספר שכבר זוהה, ו-`toc` הוא האינדקס של אותו ספר. `null` = אין קישור —
 * או שאין סימון עמוד מוכר, או שהוא לא נמצא ב-toc בפועל (למשל ספר בלי דפים,
 * כמו שו"ע, שם רמת ה-toc היא "סימן"/"סעיף" ולא "דף").
 */
export function resolveSourceTag(
  query: string,
  match: BookNameMatch,
  toc: TocIndex,
): ResolvedTag | null {
  const afterTitle = query.slice(match.matchedPrefix.length);
  const trimmed = afterTitle.replace(/^\s+/, '');
  const leadingSpace = afterTitle.length - trimmed.length;
  if (leadingSpace === 0 && afterTitle !== '') return null; // "יומאיג:" — אין הפרדה בין הכותרת לעמוד

  const marker = parsePageMarker(trimmed);
  if (!marker) return null;

  const canonical = canonicalDafText(marker.letters, marker.amud);
  const tocIndex = toc.byCanonicalText.get(canonical);
  if (tocIndex === undefined) return null;

  return {
    tocIndex,
    consumedLength: match.matchedPrefix.length + leadingSpace + marker.consumedLength,
  };
}

/** בניית ה-URL עצמו, פורמט `docs/plugin-sdk/API_REFERENCE.md` §"קישורי otzaria://". */
export function buildSourceTagHref(numericBookId: number, tocIndex: number): string {
  return `otzaria://open/book/${numericBookId}?index=${tocIndex}`;
}
