/**
 * „לחזור לאן שהייתי”.
 *
 * המבחן המרכזי כאן הוא **שני העוגנים**: המזהה של הפסקה נוסה ראשון כי הוא זול,
 * והסדר שלה בסדר המסמך הוא הגיבוי — כי הוא זה ששורד סבב של ייצוא ופתיחה
 * מחדש, גם אם המזהים לא. הבדיקות מקבעות ששני המסלולים באמת עובדים, ושהיקר
 * שבהם אינו רץ כשאין בו צורך.
 */
import { describe, it, expect, vi } from 'vitest';
import type { SelectionTarget } from 'superdoc/ui';
import {
  applyCaretAnchor,
  applyDocumentStartCaret,
  hasTextCaret,
  readCaretAnchor,
  type CaretAnchor,
  type CaretDocSource,
  type CaretUiSource,
} from '../../src/engine/caret-anchor';
import { paintedFirstBlockId } from '../../src/engine/painted-lines';

/** כפיל של `doc.blocks.list`, כולל דפדוף כמו שהמנוע רשאי לעשות. */
function docWithBlocks(ids: string[], pageSize = ids.length || 1) {
  const calls: Array<{ offset?: number; limit?: number }> = [];
  const doc: CaretDocSource = {
    activeEditor: {
      doc: {
        blocks: {
          list: (input) => {
            calls.push(input ?? {});
            const offset = input?.offset ?? 0;
            const limit = Math.min(pageSize, input?.limit ?? pageSize);
            const slice = ids.slice(offset, offset + limit);
            return {
              total: ids.length,
              blocks: slice.map((nodeId, index) => ({ nodeId, ordinal: offset + index })),
            };
          },
        },
      },
    },
  };
  return { doc, calls };
}

function uiWithSelection(slice: unknown): CaretUiSource {
  return { selection: { getSnapshot: () => slice as never } };
}

const caretAt = (blockId: string, offset: number) => ({
  kind: 'selection',
  start: { kind: 'text', blockId, offset },
  end: { kind: 'text', blockId, offset },
});

describe('readCaretAnchor', () => {
  it('קורא סמן מכווץ, עם מקומו בסדר המסמך', async () => {
    const { doc } = docWithBlocks(['b1', 'b2', 'b3']);
    const ui = uiWithSelection({ selectionTarget: caretAt('b3', 7) });

    await expect(readCaretAnchor(ui, doc)).resolves.toEqual({
      start: { blockId: 'b3', ordinal: 2, offset: 7 },
      end: null,
    });
  });

  it('בחירת טווח נשמרת עם שני הקצוות', async () => {
    const { doc } = docWithBlocks(['b1', 'b2', 'b3']);
    const ui = uiWithSelection({
      selectionTarget: {
        kind: 'selection',
        start: { kind: 'text', blockId: 'b1', offset: 2 },
        end: { kind: 'text', blockId: 'b3', offset: 5 },
      },
    });

    await expect(readCaretAnchor(ui, doc)).resolves.toEqual({
      start: { blockId: 'b1', ordinal: 0, offset: 2 },
      end: { blockId: 'b3', ordinal: 2, offset: 5 },
    });
  });

  it('נופל ל-target כשהמנוע לא הקרין selectionTarget', async () => {
    // המנוע מתעד במפורש ש-`selectionTarget` עשוי להיות `null` „כשהריצה אינה
    // יכולה להקרין אותו באמת”, ואז `target` הוא מה שנשאר.
    const { doc } = docWithBlocks(['b1', 'b2']);
    const ui = uiWithSelection({
      selectionTarget: null,
      target: { kind: 'text', segments: [{ blockId: 'b2', range: { start: 4, end: 4 } }] },
    });

    await expect(readCaretAnchor(ui, doc)).resolves.toEqual({
      start: { blockId: 'b2', ordinal: 1, offset: 4 },
      end: null,
    });
  });

  it('בחירה שאינה בטקסט אינה נשמרת', async () => {
    // גבול של טבלה או תמונה אינו „מקום בטקסט” ואין לו היסט לשחזר. במצב כזה
    // אין עוגן, והמסמך ייפתח בתחילתו — כפי שהיה תמיד.
    const { doc } = docWithBlocks(['b1']);
    const ui = uiWithSelection({
      selectionTarget: {
        kind: 'selection',
        start: { kind: 'nodeEdge', node: { kind: 'block', nodeId: 't1' }, edge: 'before' },
        end: { kind: 'nodeEdge', node: { kind: 'block', nodeId: 't1' }, edge: 'after' },
      },
    });

    await expect(readCaretAnchor(ui, doc)).resolves.toBeNull();
  });

  it('גרסת מנוע בלי ה-handle אינה מפילה — היא פשוט אינה זוכרת', async () => {
    await expect(readCaretAnchor({}, {})).resolves.toBeNull();
    await expect(readCaretAnchor(null, null)).resolves.toBeNull();
  });

  it('פסקה בלי מזהה עדיין תופסת מקום בסדר', async () => {
    // דילוג עליה היה מזיז את כל מי שאחריה בעוגן אחד — כלומר שחזור לשורה
    // הלא נכונה, שקשה יותר להבחין בו מאשר אי-שחזור.
    const doc: CaretDocSource = {
      activeEditor: {
        doc: {
          blocks: {
            list: () => ({
              total: 3,
              blocks: [{ nodeId: 'b1' }, { ordinal: 1 }, { nodeId: 'b3' }],
            }),
          },
        },
      },
    };

    const anchor = await readCaretAnchor(uiWithSelection({ selectionTarget: caretAt('b3', 0) }), doc);
    expect(anchor?.start.ordinal).toBe(2);
  });

  it('דפדוף: פסקה בעמוד השני נמצאת', async () => {
    // המנוע רשאי להחזיר פחות מהמבוקש; לולאה שמתקדמת לפי הבקשה ולא לפי מה
    // שחזר הייתה מדלגת על כל מי שמעבר לעמוד הראשון.
    const { doc, calls } = docWithBlocks(['b1', 'b2', 'b3', 'b4'], 2);

    const anchor = await readCaretAnchor(uiWithSelection({ selectionTarget: caretAt('b4', 1) }), doc);

    expect(anchor?.start.ordinal).toBe(3);
    expect(calls.length).toBeGreaterThan(1);
  });

  it('סמן שלא עזב את הפסקה אינו סורק את המסמך מחדש', async () => {
    // המקרה הרווח ביותר: הקלדה. סריקה מלאה בכל הפוגה בשביל מספר שכבר ידוע
    // היא העבודה היחידה כאן שאפשר פשוט לא לעשות.
    const { doc, calls } = docWithBlocks(['b1', 'b2', 'b3']);
    const previous: CaretAnchor = { start: { blockId: 'b2', ordinal: 1, offset: 3 }, end: null };

    const anchor = await readCaretAnchor(
      uiWithSelection({ selectionTarget: caretAt('b2', 9) }),
      doc,
      previous,
    );

    expect(anchor).toEqual({ start: { blockId: 'b2', ordinal: 1, offset: 9 }, end: null });
    expect(calls, 'הסדר היה ידוע — אין מה לקרוא').toHaveLength(0);
  });

  it('סמן שעבר פסקה כן סורק', async () => {
    const { doc, calls } = docWithBlocks(['b1', 'b2', 'b3']);
    const previous: CaretAnchor = { start: { blockId: 'b2', ordinal: 1, offset: 3 }, end: null };

    const anchor = await readCaretAnchor(
      uiWithSelection({ selectionTarget: caretAt('b3', 0) }),
      doc,
      previous,
    );

    expect(anchor?.start.ordinal).toBe(2);
    expect(calls.length).toBeGreaterThan(0);
  });

  it('קריאה שזרקה אינה מפילה את מי שקרא', async () => {
    const ui: CaretUiSource = {
      selection: {
        getSnapshot: () => {
          throw new Error('המנוע התפרק');
        },
      },
    };
    await expect(readCaretAnchor(ui, {})).resolves.toBeNull();
  });
});

/** כפיל של `ui.selection.apply` שמקבל רק את המזהים שהוא מכיר. */
function uiThatAccepts(known: string[]) {
  const applied: Array<{ start: string; offset: number }> = [];
  const selections: SelectionTarget[] = [];
  const scrolled: string[] = [];
  const ui: CaretUiSource = {
    selection: {
      apply: (target) => {
        const start = target.start as { blockId?: string; offset?: number };
        const end = target.end as { blockId?: string };
        if (!start.blockId || !end.blockId || !known.includes(start.blockId) || !known.includes(end.blockId)) {
          return { ok: false };
        }
        applied.push({ start: start.blockId, offset: start.offset ?? 0 });
        selections.push(target);
        return { ok: true };
      },
    },
    viewport: {
      scrollIntoView: (input) => {
        scrolled.push(input.target.blockId);
        return { success: true };
      },
    },
  };
  return { ui, applied, selections, scrolled };
}

describe('applyCaretAnchor', () => {
  const anchor: CaretAnchor = { start: { blockId: 'old-3', ordinal: 2, offset: 6 }, end: null };

  it('המזהה השמור נוסה ראשון, ובהצלחה אין קריאה לסדר הפסקאות', async () => {
    const { ui, applied, scrolled } = uiThatAccepts(['old-3']);
    const { doc, calls } = docWithBlocks(['x', 'y', 'z']);

    await expect(applyCaretAnchor(ui, doc, anchor)).resolves.toBe(true);
    expect(applied).toEqual([{ start: 'old-3', offset: 6 }]);
    expect(scrolled, 'גם התצוגה חוזרת, לא רק הסמן').toEqual(['old-3']);
    expect(calls, 'המסלול היקר לא רץ בכלל').toHaveLength(0);
  });

  it('מזהים שלא שרדו — הסדר מציל', async () => {
    // בדיוק המצב שהעוגן השני נכתב בשבילו: המסמך יוצא ונפתח מחדש, והמנוע
    // חילק מזהים חדשים.
    const { ui, applied } = uiThatAccepts(['new-1', 'new-2', 'new-3']);
    const { doc, calls } = docWithBlocks(['new-1', 'new-2', 'new-3']);

    await expect(applyCaretAnchor(ui, doc, anchor)).resolves.toBe(true);
    expect(applied).toEqual([{ start: 'new-3', offset: 6 }]);
    expect(calls).toEqual([{ offset: 2, limit: 1 }]);
  });

  it('מיקום עמוק במסמך נפתר בקריאה אחת, גם מעבר לתקרת הסריקה של שמירת העוגן', async () => {
    const ordinal = 40_003;
    const { ui, applied, scrolled } = uiThatAccepts([`new-${ordinal}`]);
    const list = vi.fn((input?: { offset?: number; limit?: number }) => ({
      total: 50_000,
      blocks: [{ nodeId: `new-${input?.offset ?? 0}`, ordinal: input?.offset ?? 0 }],
    }));
    const deep: CaretAnchor = { start: { blockId: 'old-deep', ordinal, offset: 7 }, end: null };

    await expect(
      applyCaretAnchor(ui, { activeEditor: { doc: { blocks: { list } } } }, deep),
    ).resolves.toBe(true);

    expect(list.mock.calls).toEqual([[{ offset: ordinal, limit: 1 }]]);
    expect(applied).toEqual([{ start: `new-${ordinal}`, offset: 7 }]);
    expect(scrolled).toEqual([`new-${ordinal}`]);
  });

  it.each([[0, 3], [3, 0]])('טווח %i→%i נשמר עם שני ההיסטים בשתי קריאות ממוקדות', async (start, end) => {
    const ids = ['new-1', 'new-2', 'new-3', 'new-4'];
    const { doc, calls } = docWithBlocks(ids, 2);
    const { ui, selections, scrolled } = uiThatAccepts(ids);
    const range: CaretAnchor = {
      start: { blockId: 'old-start', ordinal: start, offset: 2 },
      end: { blockId: 'old-end', ordinal: end, offset: 5 },
    };

    await expect(applyCaretAnchor(ui, doc, range)).resolves.toBe(true);

    expect(calls).toEqual([{ offset: start, limit: 1 }, { offset: end, limit: 1 }]);
    expect(selections).toEqual([{
      kind: 'selection',
      start: { kind: 'text', blockId: ids[start], offset: 2 },
      end: { kind: 'text', blockId: ids[end], offset: 5 },
    }]);
    expect(scrolled).toEqual([ids[start]]);
  });

  it.each([1, null])('קצה סיום בסדר %s חולק את קריאת ההתחלה כשאין לו מיקום אחר', async (ordinal) => {
    const { doc, calls } = docWithBlocks(['new-1', 'new-2']);
    const { ui, selections } = uiThatAccepts(['new-2']);
    const range: CaretAnchor = {
      start: { blockId: 'old-start', ordinal: 1, offset: 2 },
      end: { blockId: 'old-end', ordinal, offset: 9 },
    };

    await expect(applyCaretAnchor(ui, doc, range)).resolves.toBe(true);

    expect(calls).toEqual([{ offset: 1, limit: 1 }]);
    expect(selections).toEqual([{
      kind: 'selection',
      start: { kind: 'text', blockId: 'new-2', offset: 2 },
      end: { kind: 'text', blockId: 'new-2', offset: 9 },
    }]);
  });

  it('קצה סיום שנעלם נופל לפסקת ההתחלה בלי סריקה נוספת', async () => {
    const { doc, calls } = docWithBlocks(['new-1', 'new-2']);
    const { ui, selections } = uiThatAccepts(['new-2']);
    const range: CaretAnchor = {
      start: { blockId: 'old-start', ordinal: 1, offset: 2 },
      end: { blockId: 'old-end', ordinal: 8, offset: 5 },
    };

    await expect(applyCaretAnchor(ui, doc, range)).resolves.toBe(true);

    expect(calls).toEqual([{ offset: 1, limit: 1 }, { offset: 8, limit: 1 }]);
    expect(selections).toEqual([{
      kind: 'selection',
      start: { kind: 'text', blockId: 'new-2', offset: 2 },
      end: { kind: 'text', blockId: 'new-2', offset: 5 },
    }]);
  });

  it('גם לאחר פתרון מזהה חדש היסט שנדחה נופל לתחילת אותה פסקה', async () => {
    const { doc, calls } = docWithBlocks(['new-1', 'new-2', 'new-3']);
    const { ui, selections, scrolled } = uiThatAccepts(['new-3']);
    const apply = ui.selection!.apply!;
    ui.selection!.apply = (target) => {
      if ((target.start as { offset: number }).offset !== 0) return { ok: false };
      return apply(target);
    };

    await expect(applyCaretAnchor(ui, doc, anchor)).resolves.toBe(true);

    expect(calls).toEqual([{ offset: 2, limit: 1 }]);
    expect(selections).toEqual([caretAt('new-3', 0)]);
    expect(scrolled).toEqual(['new-3']);
  });

  it.each([-1, 1.5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1])('סדר לא תקין (%s) אינו נשלח למנוע', async (ordinal) => {
    const { doc, calls } = docWithBlocks(['new-1']);
    const { ui } = uiThatAccepts(['new-1']);

    await expect(
      applyCaretAnchor(ui, doc, { start: { ...anchor.start, ordinal }, end: null }),
    ).resolves.toBe(false);
    expect(calls).toEqual([]);
  });

  it.each([
    { total: 3, blocks: [] },
    { total: 3, blocks: [null] },
    { total: 3, blocks: [{ ordinal: 2 }] },
    { total: 3, blocks: [{ ordinal: 2, nodeId: '' }] },
    { total: 3, blocks: [{ ordinal: 1, nodeId: 'wrong-position' }] },
    { total: 2, blocks: [{ nodeId: 'clamped-position' }] },
  ])('תוצאת התחלה חסרה או ממיקום אחר אינה משחזרת אל פסקה שגויה: %j', async (result) => {
    const { ui, selections, scrolled } = uiThatAccepts(['wrong-position', 'clamped-position']);
    const list = vi.fn(() => result);
    const range: CaretAnchor = { ...anchor, end: { blockId: 'old-end', ordinal: 4, offset: 9 } };

    await expect(
      applyCaretAnchor(ui, { activeEditor: { doc: { blocks: { list } } } }, range),
    ).resolves.toBe(false);
    expect(list.mock.calls).toEqual([[{ offset: 2, limit: 1 }]]);
    expect(selections).toEqual([]);
    expect(scrolled).toEqual([]);
  });

  it('מנוע שאינו מחזיר ordinal עדיין נשען על אותו offset לא מסונן', async () => {
    const { ui, applied } = uiThatAccepts(['new-3']);
    const list = vi.fn(() => ({ total: 3, blocks: [{ nodeId: 'new-3' }] }));

    await expect(
      applyCaretAnchor(ui, { activeEditor: { doc: { blocks: { list } } } }, anchor),
    ).resolves.toBe(true);
    expect(list.mock.calls).toEqual([[{ offset: 2, limit: 1 }]]);
    expect(applied).toEqual([{ start: 'new-3', offset: 6 }]);
  });

  it.each([2, 4])('כשל קריאה בסדר %i אינו מפיל פתיחה ושומר את גיבוי הקצה האחרון', async (failedOrdinal) => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const { ui, selections } = uiThatAccepts(['new-3']);
    const list = vi.fn(async (input?: { offset?: number; limit?: number }) => {
      if (input?.offset === failedOrdinal) throw new Error('קריאת פסקה נכשלה');
      return { total: 5, blocks: [{ nodeId: 'new-3', ordinal: 2 }] };
    });
    const range: CaretAnchor = { ...anchor, end: { blockId: 'old-end', ordinal: 4, offset: 9 } };

    try {
      await expect(
        applyCaretAnchor(ui, { activeEditor: { doc: { blocks: { list } } } }, range),
      ).resolves.toBe(failedOrdinal !== 2);
      expect(list.mock.calls).toEqual(failedOrdinal === 2
        ? [[{ offset: 2, limit: 1 }]]
        : [[{ offset: 2, limit: 1 }], [{ offset: 4, limit: 1 }]]);
      expect(selections).toEqual(failedOrdinal === 2 ? [] : [{
        kind: 'selection',
        start: { kind: 'text', blockId: 'new-3', offset: 6 },
        end: { kind: 'text', blockId: 'new-3', offset: 9 },
      }]);
    } finally {
      warn.mockRestore();
    }
  });

  it('היסט שאינו קיים יותר נופל לתחילת אותה פסקה', async () => {
    // הקובץ נערך מבחוץ והפסקה התקצרה. „העמוד הנכון” הוא כמעט כל מה שהמשתמש
    // רצה, ובוודאי יותר מתחילת המסמך.
    const applied: number[] = [];
    const ui: CaretUiSource = {
      selection: {
        apply: (target) => {
          const offset = (target.start as { offset?: number }).offset ?? 0;
          if (offset !== 0) return { ok: false };
          applied.push(offset);
          return { ok: true };
        },
      },
    };

    await expect(applyCaretAnchor(ui, docWithBlocks(['a']).doc, anchor)).resolves.toBe(true);
    expect(applied).toEqual([0]);
  });

  it('לא נמצא — מדווח ולא זורק', async () => {
    const { ui } = uiThatAccepts([]);
    await expect(applyCaretAnchor(ui, docWithBlocks(['a']).doc, anchor)).resolves.toBe(false);
  });

  it('עוגן בלי סדר ובלי מזהה שנמצא — אין לאן ליפול', async () => {
    const { ui } = uiThatAccepts([]);
    const orphan: CaretAnchor = { start: { blockId: 'gone', ordinal: null, offset: 0 }, end: null };
    await expect(applyCaretAnchor(ui, docWithBlocks(['a']).doc, orphan)).resolves.toBe(false);
  });

  it('בלי עוגן, ובלי handle — לא קורה כלום', async () => {
    await expect(applyCaretAnchor(uiThatAccepts([]).ui, {}, null)).resolves.toBe(false);
    await expect(applyCaretAnchor({}, {}, anchor)).resolves.toBe(false);
  });

  it('גלילה שזרקה אינה הופכת שחזור מוצלח לכשל', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const ui: CaretUiSource = {
      selection: { apply: () => ({ ok: true }) },
      viewport: {
        scrollIntoView: () => {
          throw new Error('אין גיאומטריה');
        },
      },
    };

    await expect(applyCaretAnchor(ui, {}, anchor)).resolves.toBe(true);
    warn.mockRestore();
  });
});

describe('hasTextCaret', () => {
  it('סמן בטקסט — יש', () => {
    expect(hasTextCaret(uiWithSelection({ selectionTarget: caretAt('b1', 3) }))).toBe(true);
  });

  it('נופל ל-target, כמו readCaretAnchor', () => {
    const ui = uiWithSelection({
      selectionTarget: null,
      target: { kind: 'text', segments: [{ blockId: 'b2', range: { start: 0, end: 0 } }] },
    });
    expect(hasTextCaret(ui)).toBe(true);
  });

  it('אין בחירה, בחירה שאינה בטקסט, או אין handle — אין', () => {
    expect(hasTextCaret(uiWithSelection(null))).toBe(false);
    expect(
      hasTextCaret(
        uiWithSelection({
          selectionTarget: {
            kind: 'selection',
            start: { kind: 'nodeEdge', node: { kind: 'block', nodeId: 't1' }, edge: 'before' },
            end: { kind: 'nodeEdge', node: { kind: 'block', nodeId: 't1' }, edge: 'after' },
          },
        }),
      ),
    ).toBe(false);
    expect(hasTextCaret({})).toBe(false);
    expect(hasTextCaret(null)).toBe(false);
  });

  it('קריאה שזרקה אינה מפילה', () => {
    const ui: CaretUiSource = {
      selection: {
        getSnapshot: () => {
          throw new Error('המנוע התפרק');
        },
      },
    };
    expect(hasTextCaret(ui)).toBe(false);
  });
});

describe('applyDocumentStartCaret', () => {
  it('סמן מכווץ בפסקה הראשונה, היסט 0', async () => {
    const { ui, applied } = uiThatAccepts(['b1']);
    const { doc } = docWithBlocks(['b1', 'b2']);

    await expect(applyDocumentStartCaret(ui, doc)).resolves.toBe(true);
    expect(applied).toEqual([{ start: 'b1', offset: 0 }]);
  });

  it('בלי גלילה — המסמך ממילא נפתח בתחילתו', async () => {
    const { ui, scrolled } = uiThatAccepts(['b1']);
    await applyDocumentStartCaret(ui, docWithBlocks(['b1']).doc);
    expect(scrolled).toEqual([]);
  });

  it('מסמך בלי פסקאות, או פסקה ראשונה בלי מזהה — אין סמן ואין זריקה', async () => {
    const { ui } = uiThatAccepts(['b1']);
    await expect(applyDocumentStartCaret(ui, docWithBlocks([]).doc)).resolves.toBe(false);

    const noId: CaretDocSource = {
      activeEditor: {
        doc: { blocks: { list: () => ({ total: 1, blocks: [{ ordinal: 0 }] }) } },
      },
    };
    await expect(applyDocumentStartCaret(ui, noId)).resolves.toBe(false);
  });

  it('מהעמוד המצויר, בלי לחכות ל-blocks.list — שבמסמך ארוך חוזר רק אחרי סיום העימוד', async () => {
    const { ui, applied } = uiThatAccepts(['p1']);
    const list = vi.fn(() => new Promise<never>(() => {}));
    const pending = applyDocumentStartCaret(
      ui,
      { activeEditor: { doc: { blocks: { list } } } },
      paintedPage([['body', 'p1', 1]]),
    );

    // סינכרוני: הסמן כבר במקומו לפני שמישהו ממתין להבטחה.
    expect(applied).toEqual([{ start: 'p1', offset: 0 }]);
    expect(list).not.toHaveBeenCalled();
    await expect(pending).resolves.toBe(true);
  });

  it('המנוע דחה את המזהה המצויר — חוזרים ל-blocks.list', async () => {
    const { ui, applied } = uiThatAccepts(['b1']);
    const { doc, calls } = docWithBlocks(['b1']);

    await expect(applyDocumentStartCaret(ui, doc, paintedPage([['body', 'tbl', 1]]))).resolves.toBe(true);
    expect(calls).toHaveLength(1);
    expect(applied).toEqual([{ start: 'b1', offset: 0 }]);
  });

  it('טבלה ראשונה אינה מציבה סמן בפסקה שאחריה, גם כשהמנוע אינו יכול להציב סמן בטבלה', async () => {
    const { ui, applied } = uiThatAccepts(['after-table', 'cell']);
    const { doc, calls } = docWithBlocks(['tbl:1', 'after-table']);
    const page = paintedPage([['body', 'after-table', 11], ['body', null, 1]]);
    const cell = document.createElement('div');
    cell.dataset.layoutStory = 'body';
    cell.dataset.sourceNodeId = 'cell';
    page.lastElementChild!.appendChild(cell);

    await expect(applyDocumentStartCaret(ui, doc, page)).resolves.toBe(false);
    expect(calls).toEqual([{ offset: 0, limit: 1 }]);
    expect(applied).toEqual([]);
  });

  it('קליק שהמשתמש הציב בזמן ההמתנה ל-blocks.list אינו נדרס', async () => {
    let clicked = false;
    const { ui, applied } = uiThatAccepts(['b1']);
    ui.selection!.getSnapshot = () => (clicked ? { selectionTarget: caretAt('b9', 4) } : null);
    let release: (value: unknown) => void = () => {};
    const doc: CaretDocSource = {
      activeEditor: {
        doc: {
          blocks: {
            list: () =>
              new Promise((resolve) => {
                release = resolve;
              }) as never,
          },
        },
      },
    };

    const pending = applyDocumentStartCaret(ui, doc, null);
    clicked = true;
    release({ total: 1, blocks: [{ nodeId: 'b1', ordinal: 0 }] });

    await expect(pending).resolves.toBe(false);
    expect(applied).toEqual([]);
  });

  it('המנוע דחה, או שאין handles — false, לא חריגה', async () => {
    await expect(
      applyDocumentStartCaret(uiThatAccepts([]).ui, docWithBlocks(['b1']).doc),
    ).resolves.toBe(false);
    await expect(applyDocumentStartCaret({}, docWithBlocks(['b1']).doc)).resolves.toBe(false);
    await expect(applyDocumentStartCaret(uiThatAccepts(['b1']).ui, {})).resolves.toBe(false);
  });
});

/** עמוד מצויר כמו של המנוע: `[story, sourceNodeId, pmStart]` לכל מקטע, בסדר ה-DOM. */
function paintedPage(fragments: Array<[string, string | null, number | null]>): HTMLElement {
  const page = document.createElement('div');
  for (const [story, id, start] of fragments) {
    const fragment = document.createElement('div');
    fragment.className = 'superdoc-fragment';
    fragment.dataset.layoutStory = story;
    if (id !== null) fragment.dataset.sourceNodeId = id;
    if (start !== null) fragment.dataset.pmStart = String(start);
    page.appendChild(fragment);
  }
  return page;
}

describe('paintedFirstBlockId', () => {
  it('המקטע של הגוף עם ה-pmStart הנמוך — לא הראשון ב-DOM, ולא הכותרת העליונה', () => {
    const page = paintedPage([
      ['header', 'h1', 0],
      ['body', 'p2', 40],
      ['body', 'p1', 1],
      ['footer', 'f1', 0],
    ]);
    expect(paintedFirstBlockId(page)).toBe('p1');
  });

  it('טבלה בלי מזהה שמתחילה לפני הפסקה חוסמת בחירה בפסקה, בכל סדר DOM', () => {
    expect(paintedFirstBlockId(paintedPage([['body', null, 1], ['body', 'p2', 20]]))).toBeNull();
    expect(paintedFirstBlockId(paintedPage([['body', 'p2', 20], ['body', null, 1]]))).toBeNull();
  });

  it('טבלה מאוחרת אינה מבטלת את המסלול המהיר של הפסקה הראשונה', () => {
    expect(paintedFirstBlockId(paintedPage([['body', null, 20], ['body', 'p1', 1]]))).toBe('p1');
  });

  it('בלי טווח לטבלה אין דרך להוכיח שהפסקה ראשונה', () => {
    expect(paintedFirstBlockId(paintedPage([['body', null, null], ['body', 'p2', 20]]))).toBeNull();
  });

  it('fragment בלי טווח או בלי מזהה אינו נעקף בדרך לפסקה מאוחרת', () => {
    expect(paintedFirstBlockId(paintedPage([['body', 'p1', null], ['body', 'p2', 20]]))).toBeNull();
    expect(paintedFirstBlockId(paintedPage([['body', '', 1], ['body', 'p2', 20]]))).toBeNull();
  });

  it('פסקה שנחצתה בין עמודים נבחרת לפי מקטע ההתחלה שלה', () => {
    expect(paintedFirstBlockId(paintedPage([['body', 'p1', 40], ['body', 'p2', 70], ['body', 'p1', 1]]))).toBe('p1');
  });

  it('טווח ראשון עם מזהים סותרים חוזר למנוע, בלי תלות בסדר DOM', () => {
    expect(paintedFirstBlockId(paintedPage([['body', 'p1', 1], ['body', null, 1]]))).toBeNull();
    expect(paintedFirstBlockId(paintedPage([['body', null, 1], ['body', 'p1', 1]]))).toBeNull();
  });

  it('אין עמוד מצויר — null', () => {
    expect(paintedFirstBlockId(paintedPage([]))).toBeNull();
    expect(paintedFirstBlockId(paintedPage([['header', 'h1', 0]]))).toBeNull();
    expect(paintedFirstBlockId(null)).toBeNull();
    expect(paintedFirstBlockId(undefined)).toBeNull();
  });
});
