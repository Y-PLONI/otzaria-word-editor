/**
 * שילוב עם ספריית אוצריא לצורך תיוג מקורות (`engine/source-tagging.ts`):
 * רשימת ספרים, מטא-דאטה ותוכן עניינים — בלי לגעת בתוכן עצמו.
 *
 * שלושה מטמונים עצלים, ולא שליפה בעליית התוסף: `getLibraryTree` נטענת רק
 * בפעם הראשונה שמשתמש מקליד `@` (עשרות אלפי ספרים, מיותר למי שלא משתמש
 * בתכונה), ו-`getBookToc` פר-ספר, גם היא רק כשספר זוהה בפועל. ההרשאות —
 * `library.books.read` ו-`library.content.read` — מוצהרות ב-manifest.json
 * ומקובעות מול המפה כאן ב-tests/unit/manifest.test.ts, באותה שיטה כמו
 * READER_PERMISSIONS ב-host/otzaria-reader.ts.
 */
import { call, isPermissionDenied, hostErrorCode } from './otzaria-client';
import type { BookMeta, TocEntry } from '../types/otzaria_plugin';
import type { BookNameCandidate } from '../engine/source-tagging';

export const LIBRARY_PERMISSIONS: Record<string, string> = {
  'library.getTree': 'library.books.read',
  'library.getBookMetadata': 'library.books.read',
  'library.getBookToc': 'library.content.read',
};

export type LibraryResult<T> =
  | { ok: true; value: T }
  | { ok: false; message: string; reason: string };

function hostFailure(method: string, failedAction: string, error: unknown): LibraryResult<never> {
  if (isPermissionDenied(error)) {
    const permission = LIBRARY_PERMISSIONS[method] ?? method;
    return {
      ok: false,
      reason: 'permission-denied',
      message: `${failedAction}: לתוסף חסרה ההרשאה „${permission}”. יש לאשר אותה לתוסף בהגדרות אוצריא ולטעון את הלשונית מחדש`,
    };
  }
  const code = hostErrorCode(error);
  const detail = error instanceof Error ? error.message : String(error);
  return {
    ok: false,
    reason: code ?? 'threw',
    message: code ? `${failedAction}: ${detail} (${code})` : `${failedAction}: ${detail}`,
  };
}

/** צומת בעץ הספרייה, כפי שמוחזר מ-`library.getTree` — ר' API_REFERENCE.md §library.getTree. */
export interface LibraryTreeNode {
  title: string;
  path: string;
  categories?: LibraryTreeNode[];
  books?: BookNameCandidate[];
}

let treeCache: Promise<LibraryResult<readonly BookNameCandidate[]>> | null = null;

/** אוספת רקורסיבית את כל הספרים (`bookId`+`title`) מעץ הספרייה. */
function flattenBooks(node: LibraryTreeNode, out: BookNameCandidate[]): void {
  for (const book of node.books ?? []) {
    if (typeof book?.bookId === 'string' && typeof book?.title === 'string') {
      out.push({ bookId: book.bookId, title: book.title });
    }
  }
  for (const category of node.categories ?? []) flattenBooks(category, out);
}

/**
 * כל שמות הספרים בספרייה, שטוחים. עצלה וחד-פעמית לכל הפעלת התוסף — קריאה
 * שנייה (מסמך שני, למשל) מקבלת את אותה הבטחה, לא שליפה חוזרת.
 */
export function getLibraryBookNames(): Promise<LibraryResult<readonly BookNameCandidate[]>> {
  if (treeCache) return treeCache;

  treeCache = (async (): Promise<LibraryResult<readonly BookNameCandidate[]>> => {
    let data: LibraryTreeNode | null | undefined;
    try {
      data = await call<LibraryTreeNode | null>('library.getTree', { includeBooks: true });
    } catch (error) {
      treeCache = null; // כשל אינו נזכר — ניסיון הבא יטען מחדש, לא יישאר תקוע
      return hostFailure('library.getTree', 'טעינת רשימת הספרים נכשלה', error);
    }
    if (!data) {
      treeCache = null;
      return { ok: false, reason: 'empty-tree', message: 'טעינת רשימת הספרים נכשלה: התקבל מבנה ריק' };
    }

    const books: BookNameCandidate[] = [];
    flattenBooks(data, books);
    return { ok: true, value: books };
  })();

  return treeCache;
}

/** מטא-דאטה של ספר (בעיקר בשביל ה-`id` המספרי, לבניית `otzaria://open/book/<id>`). */
export async function getBookMetadata(bookId: string): Promise<LibraryResult<BookMeta>> {
  let data: unknown;
  try {
    data = await call<BookMeta>('library.getBookMetadata', { bookId });
  } catch (error) {
    return hostFailure('library.getBookMetadata', 'שליפת מזהה הספר נכשלה', error);
  }
  if (!data || typeof data !== 'object' || typeof (data as BookMeta).id !== 'number') {
    return { ok: false, reason: 'no-numeric-id', message: 'שליפת מזהה הספר נכשלה: לא התקבל מזהה מספרי' };
  }
  return { ok: true, value: data as BookMeta };
}

const tocCache = new Map<string, Promise<LibraryResult<readonly TocEntry[]>>>();

/** תוכן עניינים של ספר, עצלה ו-cached פר-`bookId` (לא נטען עד שספר זוהה בפועל). */
export function getBookToc(bookId: string): Promise<LibraryResult<readonly TocEntry[]>> {
  const cached = tocCache.get(bookId);
  if (cached) return cached;

  const promise = (async (): Promise<LibraryResult<readonly TocEntry[]>> => {
    let data: unknown;
    try {
      data = await call<TocEntry[]>('library.getBookToc', { bookId });
    } catch (error) {
      tocCache.delete(bookId);
      return hostFailure('library.getBookToc', 'טעינת תוכן העניינים נכשלה', error);
    }
    if (!Array.isArray(data)) {
      tocCache.delete(bookId);
      return { ok: false, reason: 'not-array', message: 'טעינת תוכן העניינים נכשלה: התקבל מבנה לא צפוי' };
    }
    return { ok: true, value: data };
  })();

  tocCache.set(bookId, promise);
  return promise;
}

/** לבדיקות בלבד — בזמן ריצה אין מסלול שמפרק את המטמונים. */
export function resetLibraryCaches(): void {
  treeCache = null;
  tocCache.clear();
}
