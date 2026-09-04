/**
 * ה-overlay עצמו, עם כפיל `doc` ו-container אמיתי ב-jsdom. מה שנבדק כאן הוא
 * מה שאפשר לבדוק בלי מנוע: הטווח שנשלח לכתיבה, שני מסלולי הכתיבה, וניווט
 * המקלדת. המראה והמיקום נבדקים בשער ה-QA מול Chrome אמיתי.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { installAtMention } from '../../src/engine/at-mention-overlay';
import type { ResolvedRefHit } from '../../src/types/otzaria_plugin';

const { resolveRefMock } = vi.hoisted(() => ({ resolveRefMock: vi.fn() }));

vi.mock('../../src/host/otzaria-reader', () => ({
  resolveRef: resolveRefMock,
}));

function hit(overrides: Partial<ResolvedRefHit> = {}): ResolvedRefHit {
  return {
    id: 42,
    bookId: 'פסחים',
    bookUid: 'id:42',
    title: 'פסחים',
    reference: 'פסחים דף לד',
    index: 1234,
    isPdf: false,
    isSourceLine: true,
    isUserBook: false,
    bookPath: 'ש"ס, בבלי',
    ...overrides,
  };
}

const BLOCK = 'b1';

/** כפיל מנוע: הסמן יושב בסוף `text`, וחלון הקריאה מחזיר את הטקסט כולו. */
function fakeDoc(text: string, options: { docInsert?: () => unknown } = {}) {
  const calls = new Map<string, unknown[]>();
  const record = (name: string, input: unknown) =>
    calls.set(name, [...(calls.get(name) ?? []), input]);

  // הסמן זז עם המסמך: `insert` מחליף טווח, ואחריו הוא יושב בסוף מה שנכתב —
  // בדיוק מה שקוד ההחלפה קורא כדי לבנות את טווח העטיפה.
  let cursor = text.length;
  const hyperlinks: Record<string, unknown> = {
    insert: (input: unknown) => {
      record('hyperlinks.insert', input);
      return { success: true };
    },
  };

  const doc = {
    selection: {
      current: async () => ({
        selectionTarget: {
          kind: 'selection',
          start: { kind: 'text', blockId: BLOCK, offset: cursor },
          end: { kind: 'text', blockId: BLOCK, offset: cursor },
        },
      }),
    },
    ranges: {
      resolve: async () => ({
        preview: { text, truncated: false },
        target: { start: { offset: 0 }, end: { offset: cursor } },
      }),
    },
    insert: (input: unknown) => {
      record('insert', input);
      const { value, target } = input as {
        value: string;
        target: { start: { offset: number } };
      };
      cursor = target.start.offset + value.length;
      return options.docInsert?.() ?? { success: true };
    },
    hyperlinks,
  };

  const host = {
    activeEditor: { doc },
    ui: {
      selection: {
        getAnchorRect: () => ({ left: 200, top: 100, width: 1, height: 18 }),
      },
    },
  };
  return { host, calls };
}

/** מריצה את מחזור ה-debounce וההערכה עד שהרשימה מצוירת. */
async function settle(): Promise<void> {
  await vi.advanceTimersByTimeAsync(200);
  await vi.advanceTimersByTimeAsync(0);
}

function popup(): HTMLElement | null {
  return document.querySelector('.otzaria-at-mention');
}

function options(): HTMLElement[] {
  return [...(popup()?.querySelectorAll('[role="option"]') ?? [])] as HTMLElement[];
}

let container: HTMLElement;

beforeEach(() => {
  vi.useFakeTimers();
  resolveRefMock.mockReset();
  resolveRefMock.mockResolvedValue({ ok: true, value: [hit()] });
  container = document.createElement('div');
  document.body.appendChild(container);
});

afterEach(() => {
  vi.useRealTimers();
  container.remove();
  popup()?.remove();
});

describe('installAtMention', () => {
  it('פותח רשימה על אזכור, ושולח לאוצריא את ההפניה בלבד', async () => {
    const { host } = fakeDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();

    expect(resolveRefMock).toHaveBeenCalledWith('פסחים לד', 8);
    expect(options()).toHaveLength(1);
    expect(options()[0]!.textContent).toContain('פסחים דף לד');
    handle.dispose();
  });

  it('אינו נפתח על טקסט שאינו אזכור', async () => {
    const { host } = fakeDoc('סתם טקסט בלי סימן');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();

    expect(resolveRefMock).not.toHaveBeenCalled();
    expect(popup()).toBeNull();
    handle.dispose();
  });

  it('אינו נפתח כשאין התאמות', async () => {
    resolveRefMock.mockResolvedValue({ ok: true, value: [] });
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();

    expect(popup()).toBeNull();
    handle.dispose();
  });

  it('מחליף את הטווח מה-@ ועד הסמן, וכותב קישור עומק', async () => {
    const { host, calls } = fakeDoc('ראה @פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    // קודם מחיקת האזכור — `hyperlinks.insert` מוסיפה ואינה מחליפה.
    const deleted = calls.get('insert')?.[0] as {
      value: string;
      target: { start: { offset: number }; end: { offset: number } };
    };
    // "ראה " הוא 4 תווים, ולכן ה-@ יושב ב-4 והסמן ב-13.
    expect(deleted.value).toBe('');
    expect(deleted.target.start.offset).toBe(4);
    expect(deleted.target.end.offset).toBe(13);

    // ואז הקישור נכנס בנקודה שנפתחה.
    const linked = calls.get('hyperlinks.insert')?.[0] as {
      target: { blockId: string; range: { start: number; end: number } };
      text: string;
      link: { destination: { href: string } };
    };
    expect(linked.target).toMatchObject({ blockId: BLOCK, range: { start: 4, end: 4 } });
    expect(linked.text).toBe('פסחים דף לד');
    expect(linked.link.destination.href).toBe('otzaria://open/book/42?index=1234');
    handle.dispose();
  });

  it('אות שימוש נשארת במסמך — מוחלף רק מה-@', async () => {
    const { host, calls } = fakeDoc('כמובא ב@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    const deleted = calls.get('insert')?.[0] as { target: { start: { offset: number } } };
    // "כמובא ב" הוא 7 תווים; המחיקה מתחילה ב-@ שאחריהם.
    expect(deleted.target.start.offset).toBe(7);
    handle.dispose();
  });

  it('בלי hyperlinks.insert אין כתיבה, ויש דיווח', async () => {
    const onStatus = vi.fn();
    const { host, calls } = fakeDoc('@פסחים לד');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    delete (host.activeEditor.doc as any).hyperlinks.insert;
    const handle = installAtMention(container, host as never, { onStatus });

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    expect(calls.has('insert')).toBe(false);
    expect(onStatus).toHaveBeenCalledWith('הוספת קישור אינה זמינה במסמך זה', true);
    handle.dispose();
  });

  it('החצים מדלגים מעגלית, ו-Tab מקבל את הפריט הפעיל', async () => {
    resolveRefMock.mockResolvedValue({
      ok: true,
      value: [hit(), hit({ id: 7, reference: 'פסחים דף לה', index: 99 })],
    });
    const { host, calls } = fakeDoc('@פסחים');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();

    expect(options()[0]!.getAttribute('aria-selected')).toBe('true');
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(options()[1]!.getAttribute('aria-selected')).toBe('true');
    // מעגלי: עוד צעד אחד חוזר לראשון.
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(options()[0]!.getAttribute('aria-selected')).toBe('true');
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
    expect(options()[1]!.getAttribute('aria-selected')).toBe('true');

    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
    await vi.advanceTimersByTimeAsync(0);

    const linked = calls.get('hyperlinks.insert')?.[0] as { link: { destination: { href: string } } };
    expect(linked.link.destination.href).toBe('otzaria://open/book/7?index=99');
    handle.dispose();
  });

  it('לחיצת עכבר בוחרת את השורה שנלחצה', async () => {
    resolveRefMock.mockResolvedValue({
      ok: true,
      value: [hit(), hit({ id: 7, reference: 'פסחים דף לה', index: 99 })],
    });
    const { host, calls } = fakeDoc('@פסחים');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    options()[1]!.click();
    await vi.advanceTimersByTimeAsync(0);

    const linked = calls.get('hyperlinks.insert')?.[0] as { link: { destination: { href: string } } };
    expect(linked.link.destination.href).toBe('otzaria://open/book/7?index=99');
    handle.dispose();
  });

  it('Escape סוגר בלי לכתוב', async () => {
    const { host, calls } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));

    expect(popup()).toBeNull();
    expect(calls.has('insert')).toBe(false);
    handle.dispose();
  });

  it('keyup של מקשי הרשימה אינו מעריך מחדש', async () => {
    // נמדד בשער: בלי הסינון, ה-keyup של חץ למטה בנה session חדש שהחזיר את
    // הבחירה לפריט הראשון.
    resolveRefMock.mockResolvedValue({
      ok: true,
      value: [hit(), hit({ id: 7, reference: 'פסחים דף לה', index: 99 })],
    });
    const { host } = fakeDoc('@פסחים');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    container.dispatchEvent(new KeyboardEvent('keyup', { key: 'ArrowDown', bubbles: true }));
    await settle();

    expect(options()[1]!.getAttribute('aria-selected')).toBe('true');
    handle.dispose();
  });

  it('אחרי Escape אותו אזכור אינו נפתח מחדש', async () => {
    // נמדד בשער: ה-keyup של Escape עצמו פתח את הרשימה מיד מחדש.
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    container.dispatchEvent(new KeyboardEvent('keyup', { key: 'Escape', bubbles: true }));
    await settle();
    expect(popup()).toBeNull();

    // גם הקלדה נוספת באותו אזכור אינה מחזירה אותו.
    container.dispatchEvent(new Event('input'));
    await settle();
    expect(popup()).toBeNull();
    handle.dispose();
  });

  it('אזכור חדש נפתח גם אחרי שקודמו נדחה', async () => {
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await settle();
    expect(popup()).toBeNull();
    handle.dispose();

    // אזכור במיקום אחר — הדחייה נקשרת להיסט ה-@, לא לפיצ'ר כולו.
    const second = fakeDoc('ראה @ברכות ב');
    const handle2 = installAtMention(container, second.host as never);
    container.dispatchEvent(new Event('input'));
    await settle();
    expect(popup()).not.toBeNull();
    handle2.dispose();
  });

  it('מקש שאינו של הרשימה אינו נבלע', async () => {
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();

    const event = new KeyboardEvent('keydown', { key: 'a', bubbles: true, cancelable: true });
    container.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(false);
    handle.dispose();
  });

  it('Tab אינו נבלע כשאין רשימה פתוחה', async () => {
    const { host } = fakeDoc('סתם טקסט');
    const handle = installAtMention(container, host as never);

    const event = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
    container.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(false);
    handle.dispose();
  });

  it('כשל הרשאה מדווח פעם אחת ולא פותח רשימה', async () => {
    resolveRefMock.mockResolvedValue({
      ok: false,
      reason: 'permission-denied',
      message: 'חסרה הרשאה',
    });
    const onStatus = vi.fn();
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never, { onStatus });

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new Event('input'));
    await settle();

    expect(onStatus).toHaveBeenCalledTimes(1);
    expect(onStatus).toHaveBeenCalledWith('חסרה הרשאה', true);
    expect(popup()).toBeNull();
    // הרשאה חסרה אינה חולפת: אין טעם לשלוח קריאה נוספת בכל הקלדה.
    expect(resolveRefMock).toHaveBeenCalledTimes(1);
    handle.dispose();
  });

  it('dispose מסיר את הרשימה ומפסיק להאזין', async () => {
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    expect(popup()).not.toBeNull();

    handle.dispose();
    expect(popup()).toBeNull();

    resolveRefMock.mockClear();
    container.dispatchEvent(new Event('input'));
    await settle();
    expect(resolveRefMock).not.toHaveBeenCalled();
  });

  it('גלילה סוגרת את הרשימה — היא אינה נגררת עם הטקסט', async () => {
    const { host } = fakeDoc('@פסחים לד');
    const handle = installAtMention(container, host as never);

    container.dispatchEvent(new Event('input'));
    await settle();
    container.dispatchEvent(new Event('scroll'));

    expect(popup()).toBeNull();
    handle.dispose();
  });
});
