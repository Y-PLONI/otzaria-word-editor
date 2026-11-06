/**
 * „ערוך בוורד לאוצריא” — כפתור בסרגל העליון של הקורא, שפותח את ספר ה-Word
 * המוצג בעורך.
 *
 * ## המסלול
 *
 * הפקד מוצהר ב-`contributes.startup.toolbarItems` עם `openPlugin: true`. לחיצה
 * מעבירה את המשתמש ללשונית התוסף, ואוצריא מוסרת לה `reader.toolbar_item_clicked`
 * עם זהות הספר (`bookUid`, `currentId`, `currentType`, `currentSource`). משם
 * הספר נפתח בדיוק כמו מענף „הספרייה” של „פתח מסמך”: `library.openBookFile`
 * (`openLibraryBook` ב-host/open-sources.ts). ספר אישי נפתח לכתיבה, ו„שמור”
 * כותב לקובץ עצמו; ספר של הספרייה נפתח לקריאה, והשמירה הראשונה היא „שמור בשם”.
 *
 * ## רק על ספר Word
 *
 * `contexts` מבחין רק בין טקסט ל-PDF, ו-`when` בודק רק הגדרות ואחסון. מה שכן
 * רואה את סוג הספר הוא תכנית Host (`startup.programs`): אוצריא מריצה אותה בכל
 * החלפת ספר, בלי להעיר את התוסף, עם `reader.book.type` — `docx`/`docm` לספר
 * Word. הפקד קשור אליה ב-`binding.visibleOutput`, ולכן מוצג רק כשהיא מחזירה
 * ערך. `openPlugin` בפקד קשור נתמך מ-0.9.98.1, ומכאן `minAppVersion`.
 *
 * הבדיקה אחרי הלחיצה (`readEditBookClick`) נשארת: היא מה שמגן אם התכנית לא
 * חושבה מחדש, ובכל מקרה היא שעושה את הפענוח.
 *
 * ## ה-latch
 *
 * `openPlugin: true` מוסר את האירוע מיד בסוף ה-boot, לפני שהבאנדל נטען — אותו
 * מרוץ כמו ב„שלח למסמך” וב-`plugin.page_opened`, ואותו פתרון (`index.html`).
 */
import { EDITABLE_BOOK_TYPES, type LibraryBook } from '../sessions/open-sources';
import type { ToolbarItemClickedEvent } from '../types/otzaria_plugin';

export const EDIT_BOOK_ITEM_ID = 'otzaria-word-edit-book';

/** מזהה התכנית שמכריעה אם הספר המוצג הוא Word. */
export const WORD_BOOK_PROGRAM_ID = 'word-book';

/**
 * התכנית כפי שהיא מוצהרת ב-`contributes.startup.programs`. נגזרת מ-
 * `EDITABLE_BOOK_TYPES`, כך שהסוגים שהכפתור מוצג עליהם הם בדיוק הסוגים
 * שהלחיצה מקבלת; tests/unit/edit-book.test.ts מקבע שהמניפסט זהה לה.
 *
 * `data.choose` מחזיר `null` על ספר שאינו Word, ו-`visibleOutput` מסתיר פקד
 * שהפלט שלו ריק. בלי ספר פתוח `reader.book.type` חסר, וההשוואה שקרית.
 */
export const WORD_BOOK_PROGRAM = {
  id: WORD_BOOK_PROGRAM_ID,
  version: 1,
  triggers: ['reader.activeBookChanged'],
  commands: [
    {
      id: 'isWord',
      type: 'data.choose',
      args: {
        condition: {
          op: 'or',
          conditions: EDITABLE_BOOK_TYPES.map((type) => ({
            op: 'equals',
            left: { $context: 'reader.book.type' },
            right: type,
          })),
        },
        whenTrue: true,
        whenFalse: null,
      },
    },
  ],
  outputs: { isWord: { $result: 'isWord' } },
};

/**
 * הפקד כפי שהוא מוצהר ב-`contributes.startup.toolbarItems`.
 *
 * אין לו רישום מ-JS כגיבוי, בניגוד ל„שלח למסמך”: `reader.addToolbarItem` אינו
 * מקבל `binding`, ופקד באותו מזהה היה חוזר להופיע על כל ספר. המחיר: עד
 * שהמשתמש מדליק את `app.startup_contributions` הכפתור אינו מופיע כלל.
 */
export const EDIT_BOOK_ITEM = {
  id: EDIT_BOOK_ITEM_ID,
  title: 'ערוך בוורד לאוצריא',
  icon: 'document_edit_24_regular',
  contexts: ['reader-text'],
  binding: { program: WORD_BOOK_PROGRAM_ID, visibleOutput: 'isWord' },
  openPlugin: true,
};

/**
 * הצורה שנצרכת מהאירוע. `bookUid` אינו ב-d.ts, אבל אוצריא שולחת אותו
 * (`ReaderLocationSnapshot.toJson`), והוא הזהות היחידה שאינה תלויה בכותרת.
 */
export type EditBookEvent = Partial<ToolbarItemClickedEvent> & { bookUid?: string | null };

export type EditBookRequest =
  | { kind: 'ignore' }
  | { kind: 'rejected'; message: string }
  | { kind: 'book'; book: LibraryBook };

function text(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}

/**
 * מפענחת את הלחיצה לספר שאפשר לפתוח, להודעה, או ל„לא שלנו”.
 *
 * `ignore` — פקד אחר של התוסף (אין כרגע, אבל האירוע משותף לכל הפקדים).
 * `rejected` — הלחיצה הייתה על הפקד שלנו, והספר אינו Word או אינו מזוהה.
 */
export function readEditBookClick(event: EditBookEvent | null | undefined): EditBookRequest {
  if (!event || event.itemId !== EDIT_BOOK_ITEM_ID) return { kind: 'ignore' };

  const title = text(event.currentBook) || text(event.currentBookId) || 'הספר';
  const type = text(event.currentType).toLowerCase();
  if (!(EDITABLE_BOOK_TYPES as readonly string[]).includes(type)) {
    return { kind: 'rejected', message: `„${title}” אינו קובץ Word — „ערוך בוורד לאוצריא” פותח ספרי Word בלבד` };
  }

  const bookUid = text(event.bookUid) || null;
  const id = typeof event.currentId === 'number' ? event.currentId : null;
  const bookId = text(event.currentBookId) || text(event.currentBook);
  if (!bookUid && id === null && !bookId) {
    return { kind: 'rejected', message: 'אוצריא לא מסרה איזה ספר מוצג, ולכן לא נפתח דבר' };
  }

  return {
    kind: 'book',
    book: {
      key: bookUid ?? `${bookId}|${type}`,
      title,
      type,
      bookId,
      bookUid,
      id,
      source: text(event.currentSource) || null,
      author: '',
    },
  };
}

/**
 * השם של ה-latch ב-`index.html`. מוגדר כאן ושם, ו-
 * tests/unit/edit-book.test.ts מקבע את הזהות.
 */
export const TOOLBAR_LATCH_KEY = '__otzariaToolbarClicks';

interface ToolbarLatch {
  queue: EditBookEvent[];
  live: boolean;
}

/** החוזה של `takePendingPageOpens`: המאזין החי קודם, הריקון אחריו. */
export function takePendingToolbarClicks(): EditBookEvent[] {
  const latch = (window as unknown as Record<string, ToolbarLatch | undefined>)[TOOLBAR_LATCH_KEY];
  if (!latch) return [];
  latch.live = true;
  return latch.queue.splice(0, latch.queue.length);
}
