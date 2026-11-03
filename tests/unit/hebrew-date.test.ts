/**
 * התאריך העברי של „תאריך ושעה”: המרת הלוח, הגימטריה, והכתיבה במקום הסמן.
 * ההסבר למה זה טקסט ולא שדה `DATE \h` — ב-engine/hebrew-date.ts.
 */
import { describe, expect, it } from 'vitest';
import {
  formatHebrewDate,
  hebrewDateText,
  hebrewNumeral,
  insertHebrewDate,
  parseDateCalendar,
} from '../../src/engine/hebrew-date';
import type { DocReceipt } from '../../src/engine/document-api';

/** צהריים, כדי שאזור הזמן של המכונה לא יזיז את היום. */
const at = (iso: string, time = '12:00') => new Date(`${iso}T${time}:00`);

describe('hebrewNumeral', () => {
  it.each([
    [1, 'א׳'],
    [5, 'ה׳'],
    [10, 'י׳'],
    [11, 'י״א'],
    [15, 'ט״ו'],
    [16, 'ט״ז'],
    [20, 'כ׳'],
    [23, 'כ״ג'],
    [30, 'ל׳'],
    [115, 'קט״ו'],
    [400, 'ת׳'],
    [787, 'תשפ״ז'],
    [800, 'ת״ת'],
    [5787, 'תשפ״ז'],
  ])('%i → %s', (value, expected) => {
    expect(hebrewNumeral(value)).toBe(expected);
  });

  it('גרש וגרשיים הם התווים העבריים ולא גרש/מירכאות של ASCII', () => {
    expect(hebrewNumeral(23)).toBe('כ״ג');
    expect(hebrewNumeral(1)).toBe('א׳');
  });
});

describe('formatHebrewDate', () => {
  it.each([
    ['2026-11-03', 'כ״ג בחשוון תשפ״ז'],
    ['2026-09-12', 'א׳ בתשרי תשפ״ז'],
    ['2026-09-26', 'ט״ו בתשרי תשפ״ז'],
    ['2027-02-15', 'ח׳ באדר א׳ תשפ״ז'],
    ['2027-03-23', 'י״ד באדר ב׳ תשפ״ז'],
  ])('%s → %s', (iso, expected) => {
    expect(formatHebrewDate(at(iso))).toBe(expected);
  });

  it('שעה מצורפת ב-HH:mm, עם אפס מוביל', () => {
    expect(hebrewDateText(at('2026-11-03', '09:05'), true)).toBe('כ״ג בחשוון תשפ״ז 09:05');
    expect(hebrewDateText(at('2026-11-03', '09:05'), false)).toBe('כ״ג בחשוון תשפ״ז');
  });
});

describe('parseDateCalendar', () => {
  it('מקבל רק את שני הלוחות', () => {
    expect(parseDateCalendar('hebrew')).toBe('hebrew');
    expect(parseDateCalendar('gregorian')).toBe('gregorian');
    expect(parseDateCalendar('julian')).toBeNull();
    expect(parseDateCalendar(true)).toBeNull();
  });
});

describe('insertHebrewDate', () => {
  const caret = {
    kind: 'selection',
    start: { kind: 'text', blockId: 'p1', offset: 3 },
    end: { kind: 'text', blockId: 'p1', offset: 3 },
  };

  function host(receipt: DocReceipt = { success: true }, selectionTarget: unknown = caret) {
    const inserted: unknown[] = [];
    const applied: unknown[] = [];
    return {
      inserted,
      applied,
      host: {
        ui: {
          selection: {
            apply: (target: unknown) => {
              applied.push(target);
              return { ok: true };
            },
          },
        },
        activeEditor: {
          doc: {
            selection: {
              current: () => ({
                empty: true,
                target: { kind: 'text', segments: [{ blockId: 'p1', range: { start: 3, end: 3 } }] },
                selectionTarget,
              }),
            },
            insert: (input: unknown) => {
              inserted.push(input);
              return receipt;
            },
          },
        },
      },
    };
  }

  it('כותב את התאריך כטקסט ב-selectionTarget', async () => {
    const double = host();
    const outcome = await insertHebrewDate(double.host, true, at('2026-11-03', '14:30'));
    expect(outcome).toEqual({ ok: true });
    expect(double.inserted).toEqual([
      { value: 'כ״ג בחשוון תשפ״ז 14:30', type: 'text', target: caret },
    ]);
  });

  it('הסמן עובר אל אחרי התאריך שנכתב', async () => {
    const double = host();
    await insertHebrewDate(double.host, true, at('2026-11-03', '14:30'));
    const after = { kind: 'text', blockId: 'p1', offset: 3 + 'כ״ג בחשוון תשפ״ז 14:30'.length };
    expect(double.applied).toEqual([{ kind: 'selection', start: after, end: after }]);
  });

  it('בחירה הפוכה מוחלפת, והסמן נוחת אחרי התאריך מתחילתה', async () => {
    const backward = {
      kind: 'selection',
      start: { kind: 'text', blockId: 'p1', offset: 9 },
      end: { kind: 'text', blockId: 'p1', offset: 2 },
    };
    const double = host({ success: true }, backward);
    await insertHebrewDate(double.host, false, at('2026-11-03'));
    const offset = 2 + 'כ״ג בחשוון תשפ״ז'.length;
    expect(double.applied).toEqual([
      { kind: 'selection', start: { kind: 'text', blockId: 'p1', offset }, end: { kind: 'text', blockId: 'p1', offset } },
    ]);
  });

  it('בחירה שחוצה פסקאות — הסמן אינו מוזז לפי ניחוש', async () => {
    const across = {
      kind: 'selection',
      start: { kind: 'text', blockId: 'p1', offset: 2 },
      end: { kind: 'text', blockId: 'p2', offset: 1 },
    };
    const double = host({ success: true }, across);
    await insertHebrewDate(double.host, false, at('2026-11-03'));
    expect(double.inserted).toHaveLength(1);
    expect(double.applied).toEqual([]);
  });

  it('בלי סמן — סירוב מנומק, בלי כתיבה', async () => {
    const double = host({ success: true }, null);
    const outcome = await insertHebrewDate(double.host, false, at('2026-11-03'));
    expect(outcome.ok).toBe(false);
    expect(double.inserted).toEqual([]);
  });

  it('קבלה שנכשלה מדווחת ככשל', async () => {
    const double = host({ success: false, failure: { code: 'INVALID_TARGET' } });
    const outcome = await insertHebrewDate(double.host, false, at('2026-11-03'));
    expect(outcome.ok).toBe(false);
  });

  it('בלי מסמך — סירוב', async () => {
    expect((await insertHebrewDate(null, true)).ok).toBe(false);
  });
});
