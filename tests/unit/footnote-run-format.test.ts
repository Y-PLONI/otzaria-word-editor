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
  type NoteRunFormatHost,
} from '../../src/engine/footnote-run-format';
import { insertNote } from '../../src/engine/footnotes';

const BODY = { kind: 'story', storyType: 'body' };
const NOTE = { kind: 'story', storyType: 'footnote', noteId: '1' };

function caret(blockId: string, offset: number, story: unknown = BODY) {
  const point = { kind: 'text', blockId, offset, story };
  return { kind: 'selection', start: point, end: point, story };
}

function fakeHost(options: { family?: unknown; size?: unknown; at?: unknown } = {}) {
  const executed: Array<[string, unknown]> = [];
  let listener: ((slice: unknown) => void) | null = null;
  const unsubscribe = vi.fn(() => {
    listener = null;
  });
  let insideNote = false;
  const values: Record<string, unknown> = { 'font-family': options.family ?? 'TaameyDavidCLM', 'font-size': options.size ?? '16' };
  const host = {
    activeEditor: {
      doc: {
        selection: { current: async () => ({ selectionTarget: options.at ?? caret('b1', 9) }) },
        footnotes: {
          insert: () => {
            insideNote = true;
            return { success: true };
          },
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
          fn({ status: 'ready', selectionTarget: caret('n1', 0, NOTE) });
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

  it('במקום אחר בגוף — לא נוגעים בעיצוב, וממשיכים לחכות', async () => {
    const { host, executed, move, isListening } = fakeHost();

    await insertNote(host as never, 'footnote');
    move(caret('b1', 9)); // לפני הסימן
    move(caret('b2', 10)); // פסקה אחרת
    move(caret('b1', 10), 'stale'); // הקריאה עוד לא התיישבה

    expect(executed).toEqual([]);
    expect(isListening()).toBe(true);
  });

  it('הערה שנייה מחליפה את המעקב של הראשונה', async () => {
    const first = fakeHost();
    const anchor = { blockId: 'b1', offset: 9, family: 'David', size: 12 };

    carryFormatPastNote(first.host as never, anchor);
    carryFormatPastNote(first.host as never, { ...anchor, offset: 20 });

    expect(first.unsubscribe).toHaveBeenCalledTimes(1);
    first.move(caret('b1', 10));
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
