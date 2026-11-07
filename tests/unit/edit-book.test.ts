/**
 * „ערוך בוורד לאוצריא” — הכפתור בסרגל הקורא שפותח את ספר ה-Word המוצג בעורך.
 *
 * שלושה חלקים: הפענוח של אירוע הלחיצה (ספר Word נפתח, ספר אחר מקבל הודעה,
 * פקד אחר נבלע), ה-latch שב-`index.html` — הלחיצה מגיעה בסוף ה-boot, לפני
 * הבאנדל — וההצהרה במניפסט, שהיא המסלול שבו הכפתור קיים לפני שהתוסף רץ.
 */
import { describe, it, expect, afterEach, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  EDIT_BOOK_ITEM,
  EDIT_BOOK_ITEM_ID,
  TOOLBAR_LATCH_KEY,
  readEditBookClick,
  takePendingToolbarClicks,
  type EditBookEvent,
} from '../../src/host/edit-book';
import { EDITABLE_BOOK_TYPES } from '../../src/sessions/open-sources';
import { registerEditBookItem, READER_PERMISSIONS } from '../../src/host/otzaria-reader';

/** בדיוק מה ש-`_dispatchItemClick` באוצריא בונה, לספר Word אישי. */
function click(overrides: Partial<EditBookEvent> = {}): EditBookEvent {
  return {
    itemId: EDIT_BOOK_ITEM_ID,
    context: 'reader-text',
    currentBook: 'חידושים',
    currentBookId: 'חידושים',
    bookUid: 'user:42',
    currentId: 42,
    currentType: 'docx',
    currentSource: 'user',
    currentIndex: 3,
    currentRef: 'חידושים, פרק א',
    param: null,
    ...overrides,
  };
}

describe('readEditBookClick', () => {
  it('ספר Word אישי — נפתח בזהות שאוצריא מסרה, עם המקור שקובע כתיבה', () => {
    expect(readEditBookClick(click())).toEqual({
      kind: 'book',
      book: {
        key: 'user:42',
        title: 'חידושים',
        type: 'docx',
        bookId: 'חידושים',
        bookUid: 'user:42',
        id: 42,
        source: 'user',
        author: '',
      },
    });
  });

  it('docm נפתח גם הוא, וסוג באותיות גדולות מנורמל', () => {
    const request = readEditBookClick(click({ currentType: 'DOCM' }));
    expect(request.kind).toBe('book');
    if (request.kind === 'book') expect(request.book.type).toBe('docm');
  });

  it('ספר של הספרייה — המקור נמסר כמו שהוא, וההחלטה על קריאה בלבד נשארת ל-openLibraryBook', () => {
    const request = readEditBookClick(click({ currentSource: 'library' }));
    expect(request.kind === 'book' && request.book.source).toBe('library');
  });

  it('בלי bookUid (גרסה ישנה) — המפתח נבנה מהכותרת והסוג', () => {
    const request = readEditBookClick(click({ bookUid: undefined }));
    expect(request.kind === 'book' && request.book.key).toBe('חידושים|docx');
    expect(request.kind === 'book' && request.book.bookUid).toBe(null);
  });

  it.each([
    ['ספר טקסט', 'text'],
    ['PDF', 'pdf'],
    ['EPUB', 'epub'],
    ['סוג חסר', null],
  ])('%s — הודעה שאומרת שהכפתור פותח ספרי Word בלבד', (_label, type) => {
    const request = readEditBookClick(click({ currentType: type }));
    expect(request.kind).toBe('rejected');
    if (request.kind === 'rejected') expect(request.message).toContain('„חידושים” אינו קובץ Word');
  });

  it('Word בלי שום זהות — הודעה, ולא פתיחה של ספר אקראי', () => {
    const request = readEditBookClick(
      click({ bookUid: null, currentId: null, currentBook: null, currentBookId: null }),
    );
    expect(request.kind).toBe('rejected');
  });

  it.each([
    ['פקד אחר', click({ itemId: 'something-else' })],
    ['אירוע ריק', {}],
    ['null', null],
  ])('%s — מתעלמים', (_label, event) => {
    expect(readEditBookClick(event as EditBookEvent | null)).toEqual({ kind: 'ignore' });
  });
});

describe('ה-latch של reader.toolbar_item_clicked', () => {
  const html = readFileSync(join(process.cwd(), 'index.html'), 'utf8').replace(/\r\n/g, '\n');

  afterEach(() => {
    delete (window as unknown as Record<string, unknown>)[TOOLBAR_LATCH_KEY];
  });

  it('ה-latch ב-index.html נושא את השם שהקוד קורא לו, ומאזין לאירוע של אוצריא', () => {
    expect(html).toContain(`window.${TOOLBAR_LATCH_KEY} =`);
    expect(html).toContain("window.addEventListener('reader.toolbar_item_clicked'");
  });

  it('הוא רץ ב-head, לפני הבאנדל', () => {
    expect(html.indexOf(TOOLBAR_LATCH_KEY)).toBeLessThan(html.indexOf('src/main.ts'));
  });

  function runLatchScript(): void {
    const at = html.indexOf(`window.${TOOLBAR_LATCH_KEY} =`);
    const start = html.lastIndexOf('(function () {', at);
    const end = html.indexOf('})();', at) + '})();'.length;
    new Function(html.slice(start, end))();
  }

  function toolbarClicked(detail: unknown): void {
    window.dispatchEvent(new CustomEvent('reader.toolbar_item_clicked', { detail }));
  }

  it('צובר את ה-detail ומוסר אותו בריקון — כך שהוא עובר את הפענוח', () => {
    runLatchScript();
    toolbarClicked(click());

    const pending = takePendingToolbarClicks();
    expect(pending).toEqual([click()]);
    expect(readEditBookClick(pending[0]).kind).toBe('book');
  });

  it('אחרי הריקון הוא מפסיק לצבור', () => {
    runLatchScript();
    takePendingToolbarClicks();
    toolbarClicked(click());

    expect(takePendingToolbarClicks()).toEqual([]);
  });

  it('בלי latch — אין ממתינים, ואין זריקה', () => {
    expect(takePendingToolbarClicks()).toEqual([]);
  });
});

describe('ההצהרה במניפסט', () => {
  const manifest = JSON.parse(readFileSync(join(process.cwd(), 'public/manifest.json'), 'utf8')) as {
    permissions?: string[];
    minAppVersion?: string;
    contributes?: { startup?: { toolbarItems?: unknown[] } };
  };

  it('הפקד מוצהר כתרומת עלייה, זהה לזה שנרשם מ-JS', () => {
    expect(manifest.contributes?.startup?.toolbarItems).toEqual([
      { ...EDIT_BOOK_ITEM, contexts: [...EDIT_BOOK_ITEM.contexts], bookTypes: [...EDIT_BOOK_ITEM.bookTypes] },
    ]);
  });

  it('מסנן בדיוק את הסוגים שהפתיחה מקבלת ודורש מארח שתומך בסינון', () => {
    expect(EDIT_BOOK_ITEM.bookTypes).toEqual([...EDITABLE_BOOK_TYPES]);
    expect(manifest.minAppVersion).toBe('0.9.99');
  });

  it('הגיבוי הדינמי שולח אותו סינון גם בלי הרשאת תרומות עלייה', async () => {
    const call = vi.fn(async (method: string) => {
      if (method !== 'reader.addToolbarItem') throw new Error('רק reader.toolbar מאושרת');
      return { success: true, data: true, error: null };
    });
    window.Otzaria = { call } as never;
    try {
      await expect(registerEditBookItem()).resolves.toEqual({ ok: true, value: undefined });
      expect(call).toHaveBeenCalledExactlyOnceWith('reader.addToolbarItem',
        manifest.contributes?.startup?.toolbarItems?.[0]);
    } finally {
      delete (window as Partial<Window>).Otzaria;
    }
  });

  it('ההרשאה שהפקד דורש מוצהרת', () => {
    expect(READER_PERMISSIONS['reader.addToolbarItem']).toBe('reader.toolbar');
    expect(manifest.permissions).toContain('reader.toolbar');
  });
});
