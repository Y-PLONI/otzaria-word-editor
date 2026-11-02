/**
 * הקלדה אחרי סימן הערה שהוכנס זה עתה. הממצא המלא ב-footnote-run-format.ts:
 * ריצת ההפניה נכתבת בלי העיצוב הישיר של הגוף, וההקלדה שאחריה נופלת לגופן
 * ברירת המחדל. התיקון קובע את הסימונים של ההקלדה הבאה — רק כשהסמן חוזר
 * בדיוק למקום שאחרי הסימן.
 */
import { describe, expect, it, vi } from 'vitest';
import {
  captureNoteAnchor,
  carryFormatPastNote,
  createNoteRunFormatExtension,
  type NoteRunFormatHost,
} from '../../src/engine/footnote-run-format';
import { insertNote } from '../../src/engine/footnotes';

const BODY = { kind: 'story', storyType: 'body' };
const NOTE = { kind: 'story', storyType: 'footnote', noteId: '1' };

function caret(blockId: string, offset: number, story: unknown = BODY) {
  const point = { kind: 'text', blockId, offset, story };
  return { kind: 'selection', start: point, end: point, story };
}

function fakeHost(options: { family?: unknown; size?: unknown; at?: unknown; initial?: unknown } = {}) {
  const executed: Array<[string, unknown]> = [];
  let listener: ((slice: unknown) => void) | null = null;
  const unsubscribe = vi.fn(() => {
    listener = null;
  });
  let insideNote = false;
  let noteType = 'footnote';
  const values: Record<string, unknown> = { 'font-family': options.family ?? 'TaameyDavidCLM', 'font-size': options.size ?? '16' };
  const host = {
    activeEditor: {
      doc: {
        selection: { current: async () => ({ selectionTarget: options.at ?? caret('b1', 9) }) },
        footnotes: {
          insert: (input: { type: string }) => {
            insideNote = true;
            noteType = input.type;
            return { success: true, footnote: { kind: 'entity', entityType: 'footnote', noteId: '1' } };
          },
          get: vi.fn(async () => ({ noteId: '1', type: noteType })),
        },
      },
    },
    ui: {
      commands: {
        has: () => true,
        // אחרי ההכנסה הרצועה מדווחת את ההקשר של ההערה, לא של הגוף.
        get: (id: string) => ({ getState: () => ({ value: insideNote ? (id === 'font-size' ? '8' : 'Arial') : values[id] }) }),
        executeAsync: vi.fn(async (id: string, payload: unknown) => {
          executed.push([id, payload]);
          return true;
        }),
      },
      selection: {
        observe: (fn: (slice: unknown) => void) => {
          listener = fn;
          fn({ status: 'ready', selectionTarget: options.initial ?? caret('n1', 0, NOTE) });
          return unsubscribe;
        },
      },
    },
  };
  const move = (target: unknown, status = 'ready') => listener?.({ status, selectionTarget: target });
  return { host, executed, move, unsubscribe, isListening: () => listener !== null };
}

describe('footnote-run-format', () => {
  it('הסמן חוזר למקום שאחרי הסימן — גופן וגודל הגוף נקבעים, פעם אחת', async () => {
    const { host, executed, move, isListening } = fakeHost();

    expect(await insertNote(host as never, 'footnote')).toEqual({ ok: true });
    move(caret('b1', 10));
    await vi.waitFor(() => expect(executed).toHaveLength(2));

    // מה שנקרא לפני ההכנסה, ולא „Arial / 8” שהרצועה מדווחת מתוך ההערה.
    expect(executed).toEqual([
      ['font-family', 'TaameyDavidCLM'],
      ['font-size', 16],
    ]);
    expect(isListening()).toBe(false);
    move(caret('b1', 10));
    expect(executed).toHaveLength(2);
  });

  it('חזרה למקום אחר בגוף מבטלת לצמיתות את התיקון', async () => {
    const { host, executed, move, isListening } = fakeHost();

    await insertNote(host as never, 'footnote');
    move(caret('b1', 9)); // לפני הסימן
    move(caret('b2', 10)); // פסקה אחרת
    move(caret('b1', 10), 'stale'); // הקריאה עוד לא התיישבה

    expect(executed).toEqual([]);
    expect(isListening()).toBe(false);
    move(caret('b1', 10));
    await Promise.resolve();
    expect(executed).toEqual([]);
  });

  it('ignores a settled pre-insertion snapshot until the selection actually enters the note', async () => {
    const f = fakeHost({ initial: caret('b1', 9) });
    await insertNote(f.host as never, 'footnote');
    expect(f.isListening()).toBe(true);
    f.move(caret('n1', 0, NOTE));
    f.move(caret('b1', 10));
    await vi.waitFor(() => expect(f.executed).toHaveLength(2));
  });

  it('a later body selection cancels even when no note snapshot was observed', async () => {
    const f = fakeHost({ initial: caret('b1', 9) });
    await insertNote(f.host as never, 'footnote');
    f.move(caret('b2', 0));
    f.move(caret('n1', 0, NOTE));
    f.move(caret('b1', 10));
    await Promise.resolve();
    expect(f.executed).toEqual([]);
    expect(f.isListening()).toBe(false);
  });

  it('הערה שנייה מחליפה את המעקב של הראשונה', async () => {
    const first = fakeHost();
    const anchor = { blockId: 'b1', offset: 9, family: 'David', size: 12 };

    carryFormatPastNote(first.host as never, anchor, { noteId: '1', type: 'footnote' });
    carryFormatPastNote(first.host as never, { ...anchor, offset: 20 }, { noteId: '1', type: 'footnote' });

    expect(first.unsubscribe).toHaveBeenCalledTimes(1);
    first.move(caret('b1', 21));
    await vi.waitFor(() => expect(first.executed).toHaveLength(2));
  });

  it('סמן בתוך הערה או בחירת טווח — אין עוגן', async () => {
    const range = { ...caret('b1', 2), end: { kind: 'text', blockId: 'b1', offset: 5, story: BODY } };
    expect(await captureNoteAnchor(fakeHost({ at: caret('n1', 0, NOTE) }).host as never)).toBeNull();
    expect(await captureNoteAnchor(fakeHost({ at: range }).host as never)).toBeNull();
  });

  it('בלי ui — ההכנסה עצמה לא נפגעת', async () => {
    const insert = vi.fn(() => ({ success: true }));
    const host: NoteRunFormatHost & { activeEditor: unknown } = {
      activeEditor: { doc: { footnotes: { insert } } } as never,
    };

    expect(await insertNote(host as never, 'footnote')).toEqual({ ok: true });
    expect(insert).toHaveBeenCalledTimes(1);
  });
});

function installMutationGuard(host: ReturnType<typeof fakeHost>['host']) {
  const controller = createNoteRunFormatExtension();
  let mutate!: (event: { stories: Array<{ kind: 'story'; storyType: 'body' | 'footnote' | 'endnote' }> }) => void;
  const dispose = controller.extension.activate({
    onMutation: (_filter: unknown, listener: typeof mutate) => { mutate = listener; },
  } as never);
  controller.bind(host as never);
  return { mutate, controller, dispose };
}

describe('note anchor lifetime', () => {
  for (const edit of ['delete reference', 'move reference', 'replace reference', 'undo insertion', 'explicit body formatting']) {
    it(`${edit} cancels the old anchor before returning to its offset`, async () => {
      const f = fakeHost();
      const { mutate } = installMutationGuard(f.host);
      await insertNote(f.host as never, 'footnote');
      mutate({ stories: [BODY as { kind: 'story'; storyType: 'body' }] });
      f.move(caret('b1', 10));
      await Promise.resolve();
      expect(f.executed).toEqual([]);
      expect(f.isListening()).toBe(false);
    });
  }

  it('a removed note cancels even when its mutation omits the body story', async () => {
    const f = fakeHost();
    await insertNote(f.host as never, 'footnote');
    f.host.activeEditor.doc.footnotes.get.mockRejectedValueOnce(new Error('not found'));
    f.move(caret('b1', 10));
    await vi.waitFor(() => expect(f.isListening()).toBe(false));
    expect(f.executed).toEqual([]);
  });

  it('an ambiguous note identity does not apply formatting', async () => {
    const f = fakeHost();
    await insertNote(f.host as never, 'endnote');
    f.host.activeEditor.doc.footnotes.get.mockResolvedValueOnce({ noteId: '1', type: 'footnote' });
    f.move(caret('b1', 10));
    await vi.waitFor(() => expect(f.isListening()).toBe(false));
    expect(f.executed).toEqual([]);
  });

  it('editing the footnote content preserves the first return correction', async () => {
    const f = fakeHost();
    const { mutate } = installMutationGuard(f.host);
    await insertNote(f.host as never, 'footnote');
    mutate({ stories: [NOTE as { kind: 'story'; storyType: 'footnote' }] });
    f.move(caret('b1', 10));
    await vi.waitFor(() => expect(f.executed).toHaveLength(2));
  });

  it('a document replacement or disposal cancels the pending correction', async () => {
    const f = fakeHost();
    const { controller } = installMutationGuard(f.host);
    await insertNote(f.host as never, 'footnote');
    controller.dispose();
    f.move(caret('b1', 10));
    await Promise.resolve();
    expect(f.executed).toEqual([]);
    expect(f.isListening()).toBe(false);
  });

  it('a selection change during font-family prevents font-size affecting another caret', async () => {
    const f = fakeHost();
    let finish!: () => void;
    f.host.ui.commands.executeAsync.mockImplementationOnce(async (id, payload) => {
      f.executed.push([id, payload]);
      await new Promise<void>((resolve) => { finish = resolve; });
      return true;
    });
    await insertNote(f.host as never, 'footnote');
    f.move(caret('b1', 10));
    await vi.waitFor(() => expect(f.executed).toHaveLength(1));
    f.move(caret('b2', 0));
    finish();
    await Promise.resolve();
    await Promise.resolve();
    expect(f.executed).toEqual([['font-family', 'TaameyDavidCLM']]);
    expect(f.isListening()).toBe(false);
  });
});
