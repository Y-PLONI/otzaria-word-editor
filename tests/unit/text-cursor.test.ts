/**
 * סמן-הטקסט של העכבר — ה-I-beam של Word על כל עמודת הטקסט.
 *
 * שתי המשפחות שנבדקות: הגיאומטריה הטהורה (נקודה מול פס הטקסט של כל עמוד —
 * בין השוליים הצדדיים, על כל גובה העמוד כולל אזורי הכותרות), וההחלה על
 * ה-host — שהסמן נדלק רק בתוך הפס, כבה בשוליים הצדדיים וביציאה, ואינו
 * נשאר דלוק אחרי פירוק.
 */
import { describe, it, expect, vi } from 'vitest';
import {
  createTextCursorWatch,
  pointInTextArea,
  TEXT_CURSOR_DEBOUNCE_MS,
  type TextBandGeometry,
} from '../../src/engine/text-cursor';
import type { IndexedPageRect } from '../../src/engine/page-ruler';
import type { PageMarginsState } from '../../src/engine/page-setup';

/** עמוד 500×1000 פיקסלים, שוליים צדדיים של 10%. */
const geometry: TextBandGeometry = {
  pageWidthTwips: 1000,
  leftTwips: 100,
  rightTwips: 100,
};

const page: IndexedPageRect = { pageIndex: 0, leftPx: 0, topPx: 0, widthPx: 500, heightPx: 1000 };

describe('pointInTextArea', () => {
  it('בתוך פס הטקסט — גם בעמוד ריק לגמרי', () => {
    expect(pointInTextArea(250, 500, [page], geometry)).toBe(true);
  });

  it('השוליים הצדדיים אינם פס טקסט', () => {
    expect(pointInTextArea(25, 500, [page], geometry), 'שוליים שמאליים').toBe(false);
    expect(pointInTextArea(475, 500, [page], geometry), 'שוליים ימניים').toBe(false);
  });

  it('השוליים האנכיים כן — הם אזור הכותרת העליונה/תחתונה', () => {
    // לחיצה כפולה שם נכנסת לעריכת הכותרת, ובזמן שהיא פתוחה מקלידים שם ממש.
    // Word מציג I-beam על כל גובה העמוד — ראו הערת הפתיחה של המודול.
    expect(pointInTextArea(250, 50, [page], geometry), 'אזור כותרת עליונה').toBe(true);
    expect(pointInTextArea(250, 950, [page], geometry), 'אזור כותרת תחתונה').toBe(true);
  });

  it('הרווח שבין עמודים אינו פס טקסט, והעמוד השני כן', () => {
    const second: IndexedPageRect = { ...page, pageIndex: 1, topPx: 1040 };
    expect(pointInTextArea(250, 1020, [page, second], geometry), 'בין העמודים').toBe(false);
    expect(pointInTextArea(250, 1040 + 500, [page, second], geometry), 'בעמוד השני').toBe(true);
  });

  it('בלי עמודים או בלי מידות דף — אין פס', () => {
    expect(pointInTextArea(250, 500, [], geometry)).toBe(false);
    expect(pointInTextArea(250, 500, [page], { ...geometry, pageWidthTwips: 0 })).toBe(false);
  });
});

/** host עם עמוד מצויר אחד, במלבנים קבועים — ה-DOM המזערי שהמעקב צריך. */
function hostWithPage() {
  const host = document.createElement('div');
  const pageEl = document.createElement('div');
  pageEl.setAttribute('data-page-index', '0');
  host.appendChild(pageEl);
  document.body.appendChild(host);

  const box = (left: number, top: number, width: number, height: number) =>
    ({ left, top, right: left + width, bottom: top + height, width, height }) as DOMRect;
  vi.spyOn(host, 'getBoundingClientRect').mockReturnValue(box(0, 0, 600, 1200));
  vi.spyOn(pageEl, 'getBoundingClientRect').mockReturnValue(box(0, 0, 500, 1000));
  return host;
}

const margins: PageMarginsState = {
  pageWidthTwips: 1000,
  pageHeightTwips: 1000,
  leftTwips: 100,
  rightTwips: 100,
  topTwips: 100,
  bottomTwips: 100,
  effectiveTopTwips: 100,
  effectiveBottomTwips: 100,
  direction: 'rtl',
};

function moveTo(host: HTMLElement, clientX: number, clientY: number): void {
  host.dispatchEvent(new MouseEvent('mousemove', { clientX, clientY }));
}

describe('createTextCursorWatch', () => {
  it('הסמן נדלק בתוך פס הטקסט, כבה בשוליים וביציאה', async () => {
    const host = hostWithPage();
    const watch = createTextCursorWatch({ host, readMargins: () => Promise.resolve(margins) });
    watch.refreshNow();
    await new Promise((resolve) => setTimeout(resolve, 0));

    moveTo(host, 250, 500);
    expect(host.style.cursor, 'בתוך הפס').toBe('text');

    moveTo(host, 25, 500);
    expect(host.style.cursor, 'בשוליים').toBe('');

    moveTo(host, 250, 500);
    host.dispatchEvent(new MouseEvent('mouseleave'));
    expect(host.style.cursor, 'העכבר יצא').toBe('');

    watch.dispose();
    host.remove();
  });

  it('בלי שוליים מהמסמך (עוד נטען) — אין סמן, ואין זריקה', async () => {
    const host = hostWithPage();
    const watch = createTextCursorWatch({
      host,
      readMargins: () => Promise.reject(new Error('המנוע עוד לא מוכן')),
    });
    watch.refreshNow();
    await new Promise((resolve) => setTimeout(resolve, 0));

    moveTo(host, 250, 500);
    expect(host.style.cursor).toBe('');

    watch.dispose();
    host.remove();
  });

  it('פירוק בזמן שהסמן דלוק מחזיר את ברירת המחדל', async () => {
    const host = hostWithPage();
    const watch = createTextCursorWatch({ host, readMargins: () => Promise.resolve(margins) });
    watch.refreshNow();
    await new Promise((resolve) => setTimeout(resolve, 0));

    moveTo(host, 250, 500);
    expect(host.style.cursor).toBe('text');

    watch.dispose();
    expect(host.style.cursor, 'מסמך שנסגר אינו משאיר סמן-טקסט').toBe('');

    moveTo(host, 250, 500);
    expect(host.style.cursor, 'אחרי הפירוק המאזין מנותק').toBe('');
    host.remove();
  });

  it('ה-host שזז נמדד מחדש בפריים הבא — המטמון אינו מתיישן', async () => {
    // ה-origin נשמר לפריים אחד (`hostOrigin`), ולכן קיפול הרצועה — שמזיז את
    // ה-host מבלי לשנות את מלבני העמודים ביחס אליו — אינו יכול להשאיר את
    // הסמן על מיפוי ישן.
    const host = hostWithPage();
    const watch = createTextCursorWatch({ host, readMargins: () => Promise.resolve(margins) });
    watch.refreshNow();
    await new Promise((resolve) => setTimeout(resolve, 0));

    moveTo(host, 250, 500);
    expect(host.style.cursor, 'בתוך הפס').toBe('text');

    // ה-host **והעמוד שבתוכו** ירדו 300 פיקסלים — כלומר המלבנים ביחס
    // ל-host לא השתנו, ו-`onChange` של המעקב אינו נקרא בכלל. רק ה-origin
    // הוא שזז, וזה בדיוק המצב שבו מטמון נצחי היה משקר.
    const pageEl = host.querySelector('[data-page-index]') as HTMLElement;
    vi.spyOn(host, 'getBoundingClientRect').mockReturnValue({
      left: 0, top: 300, right: 600, bottom: 1500, width: 600, height: 1200,
    } as DOMRect);
    vi.spyOn(pageEl, 'getBoundingClientRect').mockReturnValue({
      left: 0, top: 300, right: 500, bottom: 1300, width: 500, height: 1000,
    } as DOMRect);
    await new Promise((resolve) => requestAnimationFrame(() => resolve(null)));

    // אותה נקודה בחלון היא עכשיו 50 פיקסלים **מעל** ראש העמוד.
    moveTo(host, 250, 250);
    expect(host.style.cursor, 'מעל העמוד אחרי שה-host זז').toBe('');

    watch.dispose();
    host.remove();
  });

  it('host שאינו קיים — עצם דומם, בלי זריקות', () => {
    const watch = createTextCursorWatch({ host: null, readMargins: () => Promise.resolve(null) });
    watch.refreshNow();
    watch.noteDocumentChanged();
    watch.dispose();
  });

  it('תשובת שוליים ישנה אינה משנה סמן בזמן ההמתנה לרענון', async () => {
    vi.useFakeTimers();
    const host = hostWithPage();
    let resolve!: (value: PageMarginsState) => void;
    const readMargins = vi.fn<() => Promise<PageMarginsState | null>>()
      .mockImplementationOnce(() => new Promise((done) => { resolve = done; }))
      .mockResolvedValue({ ...margins, leftTwips: 600 });
    const watch = createTextCursorWatch({ host, readMargins });
    try {
      watch.refreshNow();
      moveTo(host, 250, 500);
      watch.noteDocumentChanged();
      resolve(margins);
      await vi.advanceTimersByTimeAsync(0);
      expect(host.style.cursor).toBe('');
      await vi.advanceTimersByTimeAsync(TEXT_CURSOR_DEBOUNCE_MS);
      expect(readMargins).toHaveBeenCalledTimes(2);
      expect(host.style.cursor).toBe('');
    } finally {
      watch.dispose();
      host.remove();
      vi.useRealTimers();
    }
  });
});
