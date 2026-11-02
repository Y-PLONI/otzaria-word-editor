/**
 * ניהול היפר-קישורים (גל 22). הבדיקה: `wrap` מקבל `TextAddress` (לא
 * SelectionTarget — נמדד שנדחה) ועוטף מפרט `link.destination`; `remove`
 * מקבל `target: HyperlinkTarget` — כתובת צומת הקישור כפי שחוזרת מ-`list()`,
 * **לא** `within: TextAddress` (הבאג המקורי: `Cannot read properties of
 * undefined (reading 'anchor')`); TARGET_NOT_FOUND בהסרה = הצלחה („אין
 * קישור" הוא המצב המבוקש); ועריכה מדווחת אמת כשהעטיפה נכשלת אחרי הסרה.
 */
import { describe, expect, it, vi } from 'vitest';
import type { DocReceipt } from '../../src/engine/document-api';
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
  const caret = (offset: number, story?: unknown) => ({
    kind: 'selection',
    start: { kind: 'text', blockId: 'b1', offset },
    end: { kind: 'text', blockId: 'b1', offset },
    ...(story ? { story } : {}),
  });

  function caretHost(result: DocReceipt = { success: true }) {
    const insert = vi.fn(async (_input: Record<string, unknown>) => result);
    const apply = vi.fn();
    const list = vi.fn();
    return { host: { activeEditor: { doc: { clipboard: { insert }, hyperlinks: { list } } }, ui: { selection: { apply } } }, insert, apply, list };
  }

  it('inserts the link and text in one transaction, without scanning all hyperlinks', async () => {
    const { host, insert, apply, list } = caretHost();
    expect(await insertHyperlinkAtCaret(host, { href: 'https://b.org/', text: 'שני', at: caret(10) })).toEqual({ ok: true });
    expect(insert).toHaveBeenCalledExactlyOnceWith({
      fragment: { kind: 'superdoc.clipboard.fragment', version: 'v2.2', blocks: [{
        kind: 'paragraph', runs: [{ text: 'שני' }],
        inlines: [{ kind: 'hyperlink', target: { kind: 'external', url: 'https://b.org/' }, inlines: [{ kind: 'text', text: 'שני' }] }],
      }] },
      formattingMode: 'match-destination', target: caret(10),
    });
    expect(list).not.toHaveBeenCalled();
    expect(apply).toHaveBeenLastCalledWith(caret(13));
  });

  it('uses the href when display text is empty and preserves story/UTF-16 offsets', async () => {
    const { host, insert, apply } = caretHost();
    const story = { kind: 'story', storyType: 'footnote', noteId: '1' };
    await insertHyperlinkAtCaret(host, { href: 'https://a.org/', text: '  ', at: caret(0, story) });
    expect(insert.mock.calls[0]).toEqual([expect.objectContaining({ fragment: expect.objectContaining({ blocks: [expect.objectContaining({ runs: [{ text: 'https://a.org/' }] })] }) })]);
    await insertHyperlinkAtCaret(host, { href: 'https://a.org/', text: 'א😀', at: caret(4, story) });
    expect(apply).toHaveBeenLastCalledWith({ ...caret(7, story), start: { kind: 'text', blockId: 'b1', offset: 7, story }, end: { kind: 'text', blockId: 'b1', offset: 7, story } });
  });

  it('reports rejection without a compensating document edit or moving the caret', async () => {
    const { host, insert, apply } = caretHost({ success: false, failure: { code: 'INVALID_TARGET', message: 'hyperlink-nested-unsupported' } });
    const outcome = await insertHyperlinkAtCaret(host, { href: 'https://c.org/', text: 'abc', at: caret(4) });
    expect(outcome).toMatchObject({ ok: false, message: expect.stringContaining('בתוך קישור קיים') });
    expect(insert).toHaveBeenCalledTimes(1);
    expect(apply).not.toHaveBeenCalled();
  });

  it('reports exceptions without another edit', async () => {
    const { host, insert, apply } = caretHost();
    insert.mockRejectedValueOnce(new Error('unavailable'));
    expect(await insertHyperlinkAtCaret(host, { href: 'https://a.org/', at: caret(0) })).toMatchObject({ ok: false, reason: 'threw' });
    expect(insert).toHaveBeenCalledTimes(1);
    expect(apply).not.toHaveBeenCalled();
  });

  it('awaits asynchronous caret placement before completing the insertion', async () => {
    const { host, insert, apply } = caretHost();
    let finish!: (result: { ok: boolean }) => void;
    apply.mockReturnValueOnce(new Promise((resolve) => { finish = resolve; }));
    let finished = false;
    const pending = insertHyperlinkAtCaret(host, { href: 'https://a.org/', text: 'abc', at: caret(4) }).then((outcome) => {
      finished = true;
      return outcome;
    });
    await vi.waitFor(() => expect(apply).toHaveBeenCalled());
    expect(finished).toBe(false);
    finish({ ok: true });
    await expect(pending).resolves.toEqual({ ok: true });
    expect(insert).toHaveBeenCalledTimes(1);
  });

  it('logs a rejected caret target while keeping the successful insertion outcome', async () => {
    const { host, insert, apply } = caretHost();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      apply.mockReturnValueOnce({ ok: false, reason: 'target-unresolved' });
      await expect(insertHyperlinkAtCaret(host, { href: 'https://a.org/', text: 'abc', at: caret(4) })).resolves.toEqual({ ok: true });
      expect(warn).toHaveBeenCalledWith(expect.stringContaining('הצבת הסמן נדחתה'), 'target-unresolved');
      expect(insert).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });

  it('handles asynchronous caret rejection without retrying the insertion', async () => {
    const { host, insert, apply } = caretHost();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const error = new Error('selection unavailable');
    try {
      apply.mockRejectedValueOnce(error);
      await expect(insertHyperlinkAtCaret(host, { href: 'https://a.org/', text: 'abc', at: caret(4) })).resolves.toEqual({ ok: true });
      expect(warn).toHaveBeenCalledWith(expect.stringContaining('הצבת הסמן נכשלה'), error);
      expect(insert).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });

  it('keeps custom otzaria links on the native command because rich inbound rejects them', async () => {
    const { host, insert, apply } = caretHost();
    expect(await insertHyperlinkAtCaret(host, { href: 'otzaria://open/book/42?index=7', at: caret(0) })).toBeNull();
    expect(insert).not.toHaveBeenCalled();
    expect(apply).not.toHaveBeenCalled();
  });

  it('leaves noncollapsed selections and unavailable APIs to the native command', async () => {
    const { host, insert } = caretHost();
    const range = { ...caret(0), end: { kind: 'text', blockId: 'b1', offset: 3 } };
    expect(await insertHyperlinkAtCaret(host, { href: 'https://a.org/', at: range })).toBeNull();
    expect(await insertHyperlinkAtCaret(host, { href: 'https://a.org/', at: null })).toBeNull();
    expect(await insertHyperlinkAtCaret({ activeEditor: { doc: {} } }, { href: 'https://a.org/', at: caret(0) })).toBeNull();
    expect(insert).not.toHaveBeenCalled();
  });
});
