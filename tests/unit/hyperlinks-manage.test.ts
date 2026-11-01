/**
 * ניהול היפר-קישורים (גל 22). הבדיקה: `wrap` מקבל `TextAddress` (לא
 * SelectionTarget — נמדד שנדחה) ועוטף מפרט `link.destination`; `remove`
 * מקבל `target: HyperlinkTarget` — כתובת צומת הקישור כפי שחוזרת מ-`list()`,
 * **לא** `within: TextAddress` (הבאג המקורי: `Cannot read properties of
 * undefined (reading 'anchor')`); TARGET_NOT_FOUND בהסרה = הצלחה („אין
 * קישור" הוא המצב המבוקש); ועריכה מדווחת אמת כשהעטיפה נכשלת אחרי הסרה.
 */
import { describe, expect, it, vi } from 'vitest';
import {
  editHyperlink,
  insertHyperlinkAtCaret,
  listHyperlinks,
  removeHyperlink,
} from '../../src/engine/hyperlinks-manage';

const RANGE_SELECTION = {
  target: {
    kind: 'text',
    segments: [{ blockId: 'b1', range: { start: 0, end: 5 } }],
  },
};

/** כתובת קישור כפי ש-`list()` מחזיר בפועל — חופפת ל-RANGE_SELECTION. */
const LINK_ADDRESS = {
  kind: 'inline',
  nodeType: 'hyperlink',
  anchor: {
    start: { blockId: 'b1', offset: 0 },
    end: { blockId: 'b1', offset: 5 },
  },
};

function fakeDoc(options: {
  links?: unknown[];
  wrap?: () => unknown;
  remove?: () => unknown;
  selection?: unknown;
} = {}) {
  const calls = new Map<string, unknown[]>();
  const make =
    (name: string, fallback: () => unknown, override?: () => unknown) =>
    (input: unknown) => {
      calls.set(name, [...(calls.get(name) ?? []), input]);
      return (override ?? fallback)() as never;
    };

  const doc = {
    selection: { current: vi.fn(async () => options.selection ?? RANGE_SELECTION) },
    hyperlinks: {
      list: make('list', () => ({
        items: options.links ?? [],
      })),
      wrap: make('wrap', () => ({ success: true }), options.wrap),
      remove: make('remove', () => ({ success: true }), options.remove),
    },
  } as never;

  return { doc, calls, host: { activeEditor: { doc } } };
}

describe('listHyperlinks', () => {
  it('משטח items לרשימה אחת', async () => {
    const { host } = fakeDoc({
      links: [
        { id: 'h1', address: LINK_ADDRESS, properties: { href: 'https://a' } },
        { id: 'h2', address: LINK_ADDRESS, properties: { anchor: 'mark' } },
      ],
    });

    const links = await listHyperlinks(host);

    expect(links).toHaveLength(2);
    expect(links?.[0]).toMatchObject({ id: 'h1', href: 'https://a' });
  });

  it('אין מנוע — null', async () => {
    expect(await listHyperlinks(null)).toBeNull();
  });
});

describe('removeHyperlink', () => {
  it('מאתר את כתובת הקישור החופפת ושולח target=HyperlinkTarget (לא within)', async () => {
    const { host, calls } = fakeDoc({
      links: [{ id: 'h1', address: LINK_ADDRESS, properties: { href: 'https://a' } }],
    });

    const outcome = await removeHyperlink(host);

    expect(outcome).toEqual({ ok: true });
    expect(calls.get('remove')?.[0]).toEqual({ target: LINK_ADDRESS });
    expect(calls.get('remove')?.[0]).not.toHaveProperty('within');
  });

  it('אין קישור בטווח — הצלחה בלי לגעת ב-remove', async () => {
    const { host, calls } = fakeDoc({ links: [] });

    const outcome = await removeHyperlink(host);

    expect(outcome).toEqual({ ok: true });
    expect(calls.get('remove')).toBeUndefined();
  });

  it('TARGET_NOT_FOUND = הצלחה — אין קישור הוא המצב המבוקש', async () => {
    const { host } = fakeDoc({
      links: [{ id: 'h1', address: LINK_ADDRESS, properties: { href: 'https://a' } }],
      remove: () => ({ success: false, failure: { code: 'TARGET_NOT_FOUND', message: 'not found' } }),
    });

    await expect(removeHyperlink(host)).resolves.toEqual({ ok: true });
  });

  it('בלי טווח — „יש לסמן טקסט תחילה"', async () => {
    const { host } = fakeDoc({ selection: { target: null } });

    const outcome = await removeHyperlink(host);

    expect(outcome).toMatchObject({ ok: false, reason: 'selection-required' });
  });
});

describe('editHyperlink', () => {
  it('כשיש קישור בטווח — remove עם target=HyperlinkTarget ואז wrap עם link.destination', async () => {
    const { host, calls } = fakeDoc({
      links: [{ id: 'h1', address: LINK_ADDRESS, properties: { href: 'https://old' } }],
    });

    const outcome = await editHyperlink(host, 'https://new');

    expect(outcome).toEqual({ ok: true });
    expect(calls.get('remove')).toEqual([{ target: LINK_ADDRESS }]);
    expect(calls.get('wrap')?.[0]).toMatchObject({
      link: { destination: { href: 'https://new' } },
    });
  });

  it('כשאין קישור בטווח — wrap ישיר בלי remove מיותר', async () => {
    const { host, calls } = fakeDoc({ links: [] });

    await editHyperlink(host, 'https://new');

    expect(calls.get('remove')).toBeUndefined();
    expect(calls.get('wrap')).toHaveLength(1);
  });

  it('עטיפה שנכשלת אחרי הסרה — הודעה שאומרת שהקישור אבד', async () => {
    const { host } = fakeDoc({
      links: [{ id: 'h1', address: LINK_ADDRESS, properties: {} }],
      wrap: () => ({ success: false, failure: { code: 'INVALID_INPUT', message: 'bad' } }),
    });

    const outcome = await editHyperlink(host, 'https://new');

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.message).toContain('הטקסט נשמר');
  });

  it('href ריק נעצר לפני המגע', async () => {
    const { host, calls } = fakeDoc();

    const outcome = await editHyperlink(host, '   ');

    expect(outcome).toMatchObject({ ok: false, reason: 'invalid-href' });
    expect(calls.size).toBe(0);
  });
});

describe('insertHyperlinkAtCaret', () => {
  const caret = (offset: number) => ({
    kind: 'selection',
    start: { kind: 'text', blockId: 'b1', offset },
    end: { kind: 'text', blockId: 'b1', offset },
  });

  function caretHost(options: { wrap?: () => unknown } = {}) {
    const { doc, calls } = fakeDoc(options);
    const inserts: unknown[] = [];
    (doc as { insert: unknown }).insert = vi.fn(async (input: unknown) => {
      inserts.push(input);
      return { success: true };
    });
    const apply = vi.fn();
    return { host: { activeEditor: { doc }, ui: { selection: { apply } } }, calls, inserts, apply };
  }

  it('כותב טקסט ועוטף אותו — לא hyperlinks.insert, שנדחה בפסקה שיש בה קישור', async () => {
    const { host, calls, inserts, apply } = caretHost();

    const outcome = await insertHyperlinkAtCaret(host, { href: 'https://b.org/', text: 'שני', at: caret(10) });

    expect(outcome).toEqual({ ok: true });
    expect(inserts).toEqual([{ value: 'שני', type: 'text', target: caret(10) }]);
    expect(calls.get('wrap')).toEqual([
      { target: { kind: 'text', blockId: 'b1', range: { start: 10, end: 13 } }, link: { destination: { href: 'https://b.org/' } } },
    ]);
    // הסמן אחרי הקישור, ולא לפניו.
    expect(apply).toHaveBeenLastCalledWith(caret(13));
  });

  it('בלי טקסט להצגה — הכתובת היא הטקסט', async () => {
    const { host, inserts } = caretHost();

    await insertHyperlinkAtCaret(host, { href: 'https://a.org/', at: caret(0) });

    expect(inserts[0]).toMatchObject({ value: 'https://a.org/' });
  });

  it('עטיפה שנדחתה בתוך קישור — הטקסט נמחק, הסמן חוזר, וההודעה בעברית', async () => {
    const { host, inserts, apply } = caretHost({
      wrap: () => ({ success: false, failure: { code: 'INVALID_TARGET', message: 'hyperlink-nested-unsupported' } }),
    });

    const outcome = await insertHyperlinkAtCaret(host, { href: 'https://c.org/', text: 'abc', at: caret(4) });

    expect(outcome).toMatchObject({ ok: false, message: expect.stringContaining('בתוך קישור קיים') });
    expect(inserts[1]).toEqual({ value: '', type: 'text', target: { ...caret(4), end: { kind: 'text', blockId: 'b1', offset: 7 } } });
    expect(apply).toHaveBeenLastCalledWith(caret(4));
  });

  it('בחירה שאינה סמן יחיד — null, והקורא נופל לפקודת המנוע', async () => {
    const { host, inserts } = caretHost();
    const range = { ...caret(0), end: { kind: 'text', blockId: 'b1', offset: 3 } };

    expect(await insertHyperlinkAtCaret(host, { href: 'https://a.org/', at: range })).toBeNull();
    expect(await insertHyperlinkAtCaret(host, { href: 'https://a.org/', at: null })).toBeNull();
    expect(inserts).toHaveLength(0);
  });
});
