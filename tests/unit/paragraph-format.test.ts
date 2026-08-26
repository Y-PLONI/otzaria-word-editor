/**
 * „פסקה” — כניסות, ריווח, שמירה וטאבים. הבדיקה על **מה נשלח למנוע** ועל
 * מסלולי הכשל; שההחלה עצמה עובדת נבדק במדידת הדפדפן (ראו הערת הפתיחה במודול).
 *
 * מה שנבדק כאן במיוחד:
 * - היחידות: twips גולמיים, אחד לאחר — הערך שנשלח הוא מה שנכתב ל-docx.
 * - מצב מלא: `setIndentation`/`setSpacing` מחליפים אלמנט, ולכן „מיוחד” לא
 *   נשלח כשהוא 'none' — שליחת `firstLine: 0` הייתה חוקית אך מטעה.
 * - השערים שלנו: טאב שלילי/שברוני נעצר **לפני** הקריאה, כי המנוע מקבל אותו
 *   בשקט (נמדד) ו-`w:pos` שלילי אינו חוקי.
 */
import { describe, expect, it, vi } from 'vitest';
import {
  TWIPS_PER_CM,
  addParagraphTabStop,
  applyParagraphIndentation,
  applyParagraphKeepOptions,
  applyParagraphSpacing,
  clearAllParagraphTabStops,
  clearParagraphIndentation,
  clearParagraphSpacing,
  emptyParagraphFormat,
  readParagraphFormat,
  readParagraphIndents,
  removeParagraphTabStop,
} from '../../src/engine/paragraph-format';

const CARET = {
  target: { kind: 'text', segments: [{ blockId: 'p3', range: { start: 2, end: 2 } }] },
};

type OpName =
  | 'setIndentation'
  | 'clearIndentation'
  | 'setSpacing'
  | 'clearSpacing'
  | 'setKeepOptions'
  | 'setTabStop'
  | 'clearTabStop'
  | 'clearAllTabStops';

/**
 * כפיל של `format.paragraph.*`. `ops` הם המימושים (null להשמטה), ו-`calls`
 * אוסף את הקלט לכל אחת — ההקלטה בעטיפה בלבד.
 */
function fakeDoc(options: { ops?: Partial<Record<OpName, ((input: unknown) => unknown) | null>>; get?: unknown; selection?: unknown } = {}) {
  const calls = new Map<OpName, unknown[]>();
  const paragraph: Record<string, unknown> = {};
  for (const name of ['setIndentation', 'clearIndentation', 'setSpacing', 'clearSpacing', 'setKeepOptions', 'setTabStop', 'clearTabStop', 'clearAllTabStops'] as const) {
    const impl = options.ops?.[name];
    if (impl === undefined) continue;
    calls.set(name, []);
    if (impl === null) continue;
    paragraph[name] = (input: unknown) => {
      calls.get(name)?.push(input);
      return impl(input);
    };
  }

  const doc = {
    selection: { current: vi.fn(async () => options.selection ?? CARET) },
    format: { paragraph },
    ...(options.get === undefined ? {} : { get: async () => options.get }),
  } as never;

  return { doc, calls, host: { activeEditor: { doc } } };
}

const ok = () => ({ success: true });

describe('applyParagraphIndentation', () => {
  it('הערכים נשלחים גולמיים ב-twips — אחד לאחר, בלי המרה', async () => {
    const { host, calls } = fakeDoc({ ops: { setIndentation: ok } });

    await applyParagraphIndentation(host, { nodeId: 'p3' }, {
      leftTwips: 720,
      rightTwips: 360,
      special: 'firstLine',
      amountTwips: 250,
    });

    expect(calls.get('setIndentation')?.[0]).toEqual({
      target: { nodeId: 'p3' },
      left: 720,
      right: 360,
      firstLine: 250,
    });
  });

  it('„מיוחד: תלויה” שולח hanging ולא firstLine', async () => {
    const { host, calls } = fakeDoc({ ops: { setIndentation: ok } });

    await applyParagraphIndentation(host, {}, {
      leftTwips: 0,
      rightTwips: 0,
      special: 'hanging',
      amountTwips: 567,
    });

    const sent = calls.get('setIndentation')?.[0] as Record<string, unknown>;
    expect(sent.hanging).toBe(567);
    expect(sent).not.toHaveProperty('firstLine');
  });

  it('„מיוחד: ללא” אינו שולח אף אחת מהתכונות', async () => {
    // `firstLine: 0` היה חוקי במנוע אך מטעה בקריאה; ההיעדר הוא המסמך.
    const { host, calls } = fakeDoc({ ops: { setIndentation: ok } });

    await applyParagraphIndentation(host, {}, { leftTwips: 1, rightTwips: 2, special: 'none', amountTwips: 0 });

    const sent = calls.get('setIndentation')?.[0] as Record<string, unknown>;
    expect(sent).not.toHaveProperty('hanging');
    expect(sent).not.toHaveProperty('firstLine');
  });

  it('ערך שאינו שלם נעצר אצלנו ולא נשלח', async () => {
    const { host, calls } = fakeDoc({ ops: { setIndentation: ok } });

    const outcome = await applyParagraphIndentation(host, {}, {
      leftTwips: 720.5,
      rightTwips: 0,
      special: 'none',
      amountTwips: 0,
    });

    expect(outcome).toMatchObject({ ok: false, reason: 'invalid-input' });
    expect(calls.get('setIndentation')).toHaveLength(0);
  });

  it('NO_OP הוא הצלחה — הערכים כבר מוגדרים', async () => {
    const { host } = fakeDoc({
      ops: { setIndentation: () => ({ success: false, failure: { code: 'NO_OP' } }) },
    });

    await expect(
      applyParagraphIndentation(host, {}, { leftTwips: 0, rightTwips: 0, special: 'none', amountTwips: 0 }),
    ).resolves.toEqual({ ok: true });
  });
});

describe('applyParagraphSpacing', () => {
  it('lineRule נשלח לצד line', async () => {
    const { host, calls } = fakeDoc({ ops: { setSpacing: ok } });

    await applyParagraphSpacing(host, {}, { beforeTwips: 240, afterTwips: 120, lineTwips: 480, rule: 'exact' });

    expect(calls.get('setSpacing')?.[0]).toMatchObject({ before: 240, after: 120, line: 480, lineRule: 'exact' });
  });

  it('rule שאינו ב-union נעצר לפני הקריאה', async () => {
    const { host, calls } = fakeDoc({ ops: { setSpacing: ok } });

    const outcome = await applyParagraphSpacing(host, {}, {
      beforeTwips: 0,
      afterTwips: 0,
      lineTwips: 480,
      rule: 'zigzag' as never,
    });

    expect(outcome).toMatchObject({ ok: false, reason: 'invalid-input' });
    expect(calls.get('setSpacing')).toHaveLength(0);
  });
});

describe('applyParagraphKeepOptions', () => {
  it('שלושת הדגלים נשלחים תמיד — booleans, ולא היעדרות', async () => {
    const { host, calls } = fakeDoc({ ops: { setKeepOptions: ok } });

    await applyParagraphKeepOptions(host, {}, { keepNext: true, keepLines: false, widowControl: false });

    expect(calls.get('setKeepOptions')?.[0]).toEqual({
      target: {},
      keepNext: true,
      keepLines: false,
      widowControl: false,
    });
  });
});

describe('tab stops', () => {
  it('עצירה חוקית נשלחת עם leader רק כשיש', async () => {
    const { host, calls } = fakeDoc({ ops: { setTabStop: ok } });

    await addParagraphTabStop(host, {}, { positionTwips: 1440, alignment: 'center', leader: 'dot' });
    await addParagraphTabStop(host, {}, { positionTwips: 2880, alignment: 'right' });

    expect(calls.get('setTabStop')?.[0]).toMatchObject({ position: 1440, alignment: 'center', leader: 'dot' });
    const second = calls.get('setTabStop')?.[1] as Record<string, unknown>;
    expect(second).not.toHaveProperty('leader');
  });

  it('מיקום שלילי נעצר אצלנו — המנוע מקבל אותו בשקט והוא פסול ב-ECMA-376', async () => {
    const { host, calls } = fakeDoc({ ops: { setTabStop: ok } });

    const outcome = await addParagraphTabStop(host, {}, { positionTwips: -100, alignment: 'left' });

    expect(outcome).toMatchObject({ ok: false, reason: 'invalid-input' });
    expect(calls.get('setTabStop')).toHaveLength(0);
  });

  it('מיקום אפס נעצר אף הוא — `w:pos` חייב להיות חיובי', async () => {
    // תפסה מוטציה: `<= 0` שהפך `< 0` שרד בלי הבדיקה הזאת.
    const { host, calls } = fakeDoc({ ops: { setTabStop: ok } });

    const outcome = await addParagraphTabStop(host, {}, { positionTwips: 0, alignment: 'left' });

    expect(outcome).toMatchObject({ ok: false, reason: 'invalid-input' });
    expect(calls.get('setTabStop')).toHaveLength(0);
  });

  it('alignment מחוץ ל-union נעצר לפני הקריאה', async () => {
    const { host, calls } = fakeDoc({ ops: { setTabStop: ok } });

    const outcome = await addParagraphTabStop(host, {}, { positionTwips: 500, alignment: 'zigzag' as never });

    expect(outcome).toMatchObject({ ok: false, reason: 'invalid-input' });
    expect(calls.get('setTabStop')).toHaveLength(0);
  });

  it('clearTabStop שולח את המיקום בלבד', async () => {
    const { host, calls } = fakeDoc({ ops: { clearTabStop: ok } });

    await removeParagraphTabStop(host, {}, 1440);

    expect(calls.get('clearTabStop')?.[0]).toEqual({ target: {}, position: 1440 });
  });

  it('clearAllTabStops שולח יעד בלבד', async () => {
    const { host, calls } = fakeDoc({ ops: { clearAllTabStops: ok } });

    await clearAllParagraphTabStops(host, {});

    expect(calls.get('clearAllTabStops')?.[0]).toEqual({ target: {} });
  });
});

describe('readParagraphFormat', () => {
  /** מסמך SDM/1 מזערי: הפסקה שהבחירה מצביעה עליה, עם props בנקודות. */
  function documentWith(props: Record<string, unknown>) {
    return {
      body: [
        { id: 'other', kind: 'paragraph', paragraph: { inlines: [] } },
        { id: 'p3', kind: 'paragraph', paragraph: { inlines: [], props } },
      ],
    };
  }

  it('הנקודות מהמודל הופכות twips — ×20', async () => {
    const { host } = fakeDoc({
      get: documentWith({ indentation: { left: 36, right: 18, firstLine: 12.5 }, keepWithNext: true }),
    });

    const result = await readParagraphFormat(host);

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.snapshot.indentation).toEqual({
        leftTwips: 720,
        rightTwips: 360,
        firstLineTwips: 250,
        hangingTwips: 0,
      });
      expect(result.snapshot.keepNext).toBe(true);
    }
  });

  it('עצירות טאב: רק kind=set, עמדות חיוביות, leader none אינו מוחזר', async () => {
    const { host } = fakeDoc({
      get: documentWith({
        tabs: [
          { kind: 'set', position: 72, alignment: 'center', leader: 'dot' },
          { kind: 'clear', position: 100 },
          { kind: 'set', position: 144, alignment: 'right' },
          { kind: 'set', position: -5, alignment: 'left' },
        ],
      }),
    });

    const result = await readParagraphFormat(host);

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.snapshot.tabs).toEqual([
        { positionTwips: 1440, alignment: 'center', leader: 'dot' },
        { positionTwips: 2880, alignment: 'right' },
      ]);
    }
  });

  it('הצומת מזוהה גם לפי `paragraphIds.paraId` — זו הצורה שהמנוע מחזיר', async () => {
    // נמדד על המנוע: `doc.get()` מחזיר
    // `{ kind:'paragraph', paragraphIds:{ paraId:'p3' }, paragraph:{ props } }`
    // — בלי `id` בכלל, ועם `indent` ולא `indentation`. הקריאה שחיפשה `id`
    // ו-`indentation` החזירה אפסים על **כל** מסמך, והדיאלוג שאושר אחריה מחק
    // כניסות שהגיעו מ-Word.
    const { host } = fakeDoc({
      get: {
        body: [
          { kind: 'paragraph', paragraphIds: { paraId: 'other' }, paragraph: { inlines: [] } },
          {
            kind: 'paragraph',
            paragraphIds: { paraId: 'p3' },
            paragraph: { inlines: [], props: { indent: { left: 36, hanging: 18 }, bidi: true } },
          },
        ],
      },
    });

    const result = await readParagraphFormat(host);

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.snapshot.indentation).toEqual({
        leftTwips: 720,
        rightTwips: 0,
        firstLineTwips: 0,
        hangingTwips: 360,
      });
    }
  });

  it('מסמך בלי הפסקה מחזיר ברירות מחדל ולא זריקה', async () => {
    const { host } = fakeDoc({ get: { body: [] } });

    const result = await readParagraphFormat(host);

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.snapshot).toEqual(emptyParagraphFormat());
  });

  it('בלי get במנוע — כשל מטופל עם הנוסח של §12', async () => {
    const { host } = fakeDoc({});

    const result = await readParagraphFormat(host);
    if (result.ok || result.outcome.ok) throw new Error('הקריאה הייתה אמורה להיכשל');

    expect(result.outcome.message).toContain('אינו זמין בגרסה זו');
  });

  it('get שזורק הופך להודעה מטופסת, ולא מפיל את הקורא', async () => {
    const throwingHost = {
      activeEditor: {
        doc: {
          selection: { current: async () => CARET },
          format: { paragraph: {} },
          get: () => {
            throw new Error('boom');
          },
        },
      },
    };

    const result = await readParagraphFormat(throwingHost as never);
    if (result.ok || result.outcome.ok) throw new Error('הקריאה הייתה אמורה להיכשל');

    expect(result.outcome.message).toContain('boom');
  });
});

describe('clears and failure paths', () => {
  it('clearIndentation/clearSpacing שולחים יעד בלבד', async () => {
    const { host, calls } = fakeDoc({ ops: { clearIndentation: ok, clearSpacing: ok } });

    await clearParagraphIndentation(host, {});
    await clearParagraphSpacing(host, {});

    expect(calls.get('clearIndentation')?.[0]).toEqual({ target: {} });
    expect(calls.get('clearSpacing')?.[0]).toEqual({ target: {} });
  });

  it('פעולה שאינה חשופה במנוע מדווחת „אינו זמין בגרסה זו”', async () => {
    const { host } = fakeDoc({ ops: { setIndentation: null } });

    const outcome = await applyParagraphIndentation(host, {}, {
      leftTwips: 0,
      rightTwips: 0,
      special: 'none',
      amountTwips: 0,
    });

    expect(outcome).toMatchObject({ ok: false, reason: 'command-unsupported' });
  });

  it('אין Document API — תוצאה מטופסת, לא זריקה', async () => {
    for (const host of [null, undefined, { activeEditor: null }] as never[]) {
      const outcome = await applyParagraphIndentation(host, {}, {
        leftTwips: 0,
        rightTwips: 0,
        special: 'none',
        amountTwips: 0,
      });
      expect(outcome.ok).toBe(false);
    }
  });

  it('בלי סמן — קריאת המצב מסרבת בהודעה מדויקת, ולא קריאה למנוע', async () => {
    const { host } = fakeDoc({ ops: { setIndentation: ok }, selection: { target: null } });

    const result = await readParagraphFormat(host);

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.outcome).toEqual({
        ok: false,
        message: 'יש למקם את הסמן במסמך',
        reason: 'selection-required',
      });
    }
  });

  it('story של הבחירה נושא את היעד (כותרת עליונה/תחתונה)', async () => {
    const story = { kind: 'header', index: 1 };
    const { host } = fakeDoc({
      get: { body: [] },
      selection: { target: { kind: 'text', segments: [{ blockId: 'p9', range: { start: 0, end: 1 } }], story } },
    });

    const result = await readParagraphFormat(host);

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.target).toMatchObject({ nodeId: 'p9', story });
    }
  });
});

describe('unit constants', () => {
  it('TWIPS_PER_CM הוא 1440 / 2.54 — הקובע את ההמרה של הדיאלוג', () => {
    expect(TWIPS_PER_CM).toBeCloseTo(566.93, 2);
  });
});

describe('readParagraphIndents', () => {
  it('מחזירה את הכניסות ואת היעד לכתיבה', async () => {
    const { host } = fakeDoc({
      get: {
        body: [
          {
            kind: 'paragraph',
            paragraphIds: { paraId: 'p3' },
            paragraph: { props: { indent: { left: 36, right: 18 }, bidi: true } },
          },
        ],
      },
    });

    const reading = await readParagraphIndents(host);

    expect(reading?.indents).toEqual({
      leftTwips: 720,
      rightTwips: 360,
      firstLineTwips: 0,
      hangingTwips: 0,
      bidi: true,
    });
    expect(reading?.target).toMatchObject({ kind: 'block', nodeType: 'paragraph', nodeId: 'p3' });
  });

  it('פסקה בלי כניסות היא אפסים, לא היעדר', async () => {
    const { host } = fakeDoc({
      get: { body: [{ kind: 'paragraph', paragraphIds: { paraId: 'p3' }, paragraph: {} }] },
    });

    expect((await readParagraphIndents(host))?.indents.leftTwips).toBe(0);
  });

  it('אין סמן במסמך — `null`, ובלי הודעת כשל', async () => {
    // הסרגל רץ ברקע ואינו פעולה של המשתמש: „יש למקם את הסמן במסמך” בשורת
    // המצב על כל לחיצה מחוץ למסמך היה רעש, לא עזרה.
    const host = {
      activeEditor: {
        doc: {
          selection: { current: async () => ({ empty: true, target: null }) },
          get: async () => ({ body: [] }),
        },
      },
    };

    expect(await readParagraphIndents(host as never)).toBeNull();
  });

  it('`get` שזורק מוחזר כ-`null` ואינו מפיל את הסרגל', async () => {
    const host = {
      activeEditor: {
        doc: {
          selection: { current: async () => CARET },
          get: () => {
            throw new Error('boom');
          },
        },
      },
    };

    expect(await readParagraphIndents(host as never)).toBeNull();
  });

  it('גרסה בלי `get` אינה זורקת', async () => {
    expect(await readParagraphIndents({ activeEditor: { doc: {} } } as never)).toBeNull();
  });
});
