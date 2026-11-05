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
  applyParagraphContextualSpacing,
  applySelectionSpacing,
  readSelectionSpacing,
  removeParagraphTabStop,
  toggleSelectionSpacing,
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
  | 'setFlowOptions'
  | 'setTabStop'
  | 'clearTabStop'
  | 'clearAllTabStops';

/**
 * כפיל של `format.paragraph.*`. `ops` הם המימושים (null להשמטה), ו-`calls`
 * אוסף את הקלט לכל אחת — ההקלטה בעטיפה בלבד.
 */
function fakeDoc(options: {
  ops?: Partial<Record<OpName, ((input: unknown) => unknown) | null>>;
  get?: unknown;
  selection?: unknown;
  /** מה ש-`blocks.list` מדווח — כדי לגזור את `nodeType` האמיתי של הבלוק שהסמן בו. */
  blocks?: readonly { nodeId: string; nodeType: string }[];
  /** `lists.getState`: מזהה בלוק → isListItem, לבלוק שאינו ב-`blocks.list` (טבלה). */
  listState?: Record<string, boolean>;
} = {}) {
  const calls = new Map<OpName, unknown[]>();
  const paragraph: Record<string, unknown> = {};
  for (const name of ['setIndentation', 'clearIndentation', 'setSpacing', 'clearSpacing', 'setKeepOptions', 'setFlowOptions', 'setTabStop', 'clearTabStop', 'clearAllTabStops'] as const) {
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
    // כמו superdoc 2.23: `get()` בלי קלט זורק, ו-`get({})` מחזיר את המסמך.
    ...(options.get === undefined
      ? {}
      : {
          get: async (input?: unknown) => {
            if (input === undefined) throw new TypeError("Cannot read properties of undefined (reading 'options')");
            return options.get;
          },
          getNodeById: async ({ nodeId }: { nodeId: string }) => {
            const body = (options.get as { body?: Array<Record<string, unknown>> }).body ?? [];
            const node = body.find((entry) => entry.id === nodeId || (entry.paragraphIds as { paraId?: string })?.paraId === nodeId);
            return { node, address: { kind: 'block', nodeType: node?.kind === 'heading' ? 'heading' : 'paragraph', nodeId } };
          },
        }),
    ...(options.blocks === undefined ? {} : { blocks: { list: async () => ({ blocks: options.blocks }) } }),
    ...(options.listState === undefined
      ? {}
      : {
          lists: {
            getState: async (input: { target: { nodeId: string } }) => {
              const state = options.listState as Record<string, boolean>;
              const id = input.target.nodeId;
              return id in state ? { success: true, isListItem: state[id] } : { success: false };
            },
          },
        }),
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

  it('`bidi` של הפסקה נקרא לתצלום — הצד שבו „לפני טקסט” יושב', async () => {
    /*
     * זהו הצרן היחיד של השדה: פס התצוגה המקדימה בדיאלוג מצייר לפיו את
     * `direction` שלו. הכניסות נכתבות ב-`w:start`/`w:end`, שהם לוגיים
     * ל**פסקה** — פסקה LTR בתוך מקטע RTL מקבלת „לפני טקסט” בצד השמאלי —
     * ולכן כיוון המקטע אינו התשובה כאן.
     */
    const rtl = await readParagraphFormat(
      fakeDoc({
        get: {
          body: [
            {
              kind: 'paragraph',
              paragraphIds: { paraId: 'p3' },
              paragraph: { inlines: [], props: { bidi: true } },
            },
          ],
        },
      }).host,
    );
    expect(rtl.ok && rtl.snapshot.bidi).toBe(true);

    // פסקה שהצהירה LTR חוזרת `false` — הערך, ולא היעדר. זו ההבחנה שמתירה
    // ליפול לכיוון המקטע רק כשהפסקה שותקת, ונמדדה בכרום על המנוע האמיתי.
    const declaredLtr = await readParagraphFormat(
      fakeDoc({
        get: {
          body: [
            {
              kind: 'paragraph',
              paragraphIds: { paraId: 'p3' },
              paragraph: { inlines: [], props: { bidi: false } },
            },
          ],
        },
      }).host,
    );
    expect(declaredLtr.ok && declaredLtr.snapshot.bidi).toBe(false);

    // ופסקה שאינה מצהירה כלל היא `null` ולא `false`: מסמך עברי שנוצר ב-Word
    // יורש את הכיוון, והיעדר שנקרא כ-LTR היה מצייר את הפס בצד ההפוך.
    const ltr = await readParagraphFormat(
      fakeDoc({
        get: {
          body: [
            {
              kind: 'paragraph',
              paragraphIds: { paraId: 'p3' },
              paragraph: { inlines: [], props: { indent: { start: 36 } } },
            },
          ],
        },
      }).host,
    );
    expect(ltr.ok && ltr.snapshot.bidi).toBe(null);

    // וגם פסקה שאין לה תכונות בכלל — כולל פסקה בתוך תא טבלה, שאינה נפתרת
    // ב-`findParagraphProps` (ראו ההערה שם).
    const noProps = await readParagraphFormat(
      fakeDoc({
        get: { body: [{ kind: 'paragraph', paragraphIds: { paraId: 'p3' }, paragraph: { inlines: [] } }] },
      }).host,
    );
    expect(noProps.ok && noProps.snapshot.bidi).toBe(null);
  });

  it('הבחירה יושבת בכותרת — היעד לכתיבה חזרה נושא nodeType:heading ולא paragraph מקובע', async () => {
    // באג 1: כתיבה חזרה עם nodeType:'paragraph' מקובע על כותרת היא כתובת
    // פסולה ונכשלת. nodeType כאן נגזר מ-blocks.list, בדיוק כמו resolveListItem.
    const { host } = fakeDoc({
      get: documentWith({ indentation: { left: 36 } }),
      blocks: [{ nodeId: 'p3', nodeType: 'heading' }],
    });

    const result = await readParagraphFormat(host);

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.target).toMatchObject({ nodeType: 'heading', nodeId: 'p3' });
  });

  it('הבחירה יושבת בפריט רשימה — היעד נושא nodeType:listItem', async () => {
    const { host } = fakeDoc({
      get: documentWith({ indentation: { left: 36 } }),
      blocks: [{ nodeId: 'p3', nodeType: 'listItem' }],
    });

    const result = await readParagraphFormat(host);

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.target).toMatchObject({ nodeType: 'listItem', nodeId: 'p3' });
  });

  it('בלי blocks.list, או כשהבלוק לא מדווח שם — nodeType נשאר paragraph', async () => {
    const withoutBlocks = await readParagraphFormat(fakeDoc({ get: documentWith({}) }).host);
    expect(withoutBlocks.ok && withoutBlocks.ok && withoutBlocks.target.nodeType).toBe('paragraph');

    const notFound = await readParagraphFormat(
      fakeDoc({ get: documentWith({}), blocks: [{ nodeId: 'other', nodeType: 'heading' }] }).host,
    );
    expect(notFound.ok && notFound.target.nodeType).toBe('paragraph');
  });

  it('פריט רשימה בתוך תא טבלה: אינו ב-blocks.list, ו-lists.getState קובע listItem', async () => {
    // issue #14 ג׳: blocks.list מונה בלוקים עליונים בלבד. הסמן ב-p3 שבתוך טבלה.
    const inTable = await readParagraphFormat(
      fakeDoc({
        get: documentWith({}),
        blocks: [{ nodeId: 'other', nodeType: 'paragraph' }],
        listState: { p3: true },
      }).host,
    );
    expect(inTable.ok && inTable.target.nodeType).toBe('listItem');

    // פסקה רגילה בטבלה — getState אומר שאינה פריט רשימה.
    const plainInTable = await readParagraphFormat(
      fakeDoc({ get: documentWith({}), blocks: [], listState: { p3: false } }).host,
    );
    expect(plainInTable.ok && plainInTable.target.nodeType).toBe('paragraph');
  });

  it('הכניסות נקראות מ-indent.start/indent.end כשהן קיימות — לא רק left/right', async () => {
    // באג 2: setIndentation({left,right}) שלנו נכתב לוגית ל-w:start/w:end
    // (נמדד ב-docs/engine-gaps.md). קריאה שמתעלמת מהם חוזרת אפסים אחרי
    // שהמשתמש קבע כניסה, סגר את הדיאלוג ופתח אותו מחדש.
    const { host } = fakeDoc({
      get: documentWith({ indent: { start: 36, end: 18, left: 999, right: 999 } }),
    });

    const result = await readParagraphFormat(host);

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.snapshot.indentation.leftTwips).toBe(720);
      expect(result.snapshot.indentation.rightTwips).toBe(360);
    }
  });

  it('בלי start/end — עדיין נופל חזרה על left/right', async () => {
    const { host } = fakeDoc({ get: documentWith({ indent: { left: 36, right: 18 } }) });

    const result = await readParagraphFormat(host);

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.snapshot.indentation.leftTwips).toBe(720);
      expect(result.snapshot.indentation.rightTwips).toBe(360);
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

  it('היעד לכתיבה נושא את ה-nodeType האמיתי — כותרת ולא paragraph מקובע', async () => {
    // באג 1: הסרגל כותב כניסות גם על כותרת/פריט רשימה, וכתיבה חזרה עם
    // nodeType:'paragraph' מקובע על כתובת כזאת פסולה ונכשלת.
    const { host } = fakeDoc({
      get: {
        body: [{ kind: 'paragraph', paragraphIds: { paraId: 'p3' }, paragraph: { props: {} } }],
      },
      blocks: [{ nodeId: 'p3', nodeType: 'listItem' }],
    });

    const reading = await readParagraphIndents(host);

    expect(reading?.target).toMatchObject({ nodeType: 'listItem', nodeId: 'p3' });
  });

  it('הכניסות נקראות מ-indent.start/indent.end כשהן קיימות', async () => {
    const { host } = fakeDoc({
      get: {
        body: [
          {
            kind: 'paragraph',
            paragraphIds: { paraId: 'p3' },
            paragraph: { props: { indent: { start: 36, end: 18, left: 999, right: 999 } } },
          },
        ],
      },
    });

    const reading = await readParagraphIndents(host);

    expect(reading?.indents.leftTwips).toBe(720);
    expect(reading?.indents.rightTwips).toBe(360);
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

  it('`getNodeById` שזורק מוחזר כ-`null` ואינו מפיל את הסרגל', async () => {
    const host = {
      activeEditor: {
        doc: {
          selection: { current: async () => CARET },
          getNodeById: () => {
            throw new Error('boom');
          },
        },
      },
    };

    expect(await readParagraphIndents(host as never)).toBeNull();
  });

  it('גרסה בלי קריאת צומת ממוקדת אינה זורקת', async () => {
    expect(await readParagraphIndents({ activeEditor: { doc: {} } } as never)).toBeNull();
  });
});

/* ------------------------------------------------------------------ */
/* ריווח על כל הבחירה — מה שתפריט „מרווח שורות וריווח” מפעיל            */
/* ------------------------------------------------------------------ */

/** בחירה שנוגעת בשלוש פסקאות, כפי שהמנוע מחזיר אותה — קטע לכל ריצה. */
const RANGE_OVER_THREE = {
  target: {
    kind: 'text',
    segments: [
      { blockId: 'p1', range: { start: 4, end: 9 } },
      { blockId: 'p2', range: { start: 0, end: 11 } },
      // אותה פסקה פעמיים: פסקה עם שתי ריצות מסומנות מחזירה שני קטעים.
      { blockId: 'p2', range: { start: 11, end: 14 } },
      { blockId: 'p3', range: { start: 0, end: 3 } },
    ],
  },
};

function bodyOf(entries: readonly (readonly [string, Record<string, unknown> | undefined])[]) {
  return {
    body: entries.map(([id, props]) => ({
      id,
      kind: 'paragraph',
      paragraph: { inlines: [], ...(props === undefined ? {} : { props }) },
    })),
  };
}

describe('readSelectionSpacing', () => {
  it('מחזירה פסקה לכל בלוק בבחירה, בלי כפילויות, ובקריאת `get` אחת', async () => {
    const { host, doc } = fakeDoc({
      selection: RANGE_OVER_THREE,
      get: bodyOf([
        ['p1', { spacing: { before: 6, after: 0, line: 18, lineRule: 'exact' } }],
        ['p2', undefined],
        ['p3', { spacing: { after: 12 } }],
      ]),
    });
    const getSpy = vi.spyOn(doc as unknown as { get: () => unknown }, 'get');

    const entries = await readSelectionSpacing(host);

    expect(entries?.map((entry) => entry.target.nodeId)).toEqual(['p1', 'p2', 'p3']);
    expect(getSpy).toHaveBeenCalledTimes(1);
  });

  it('שדה שאינו מוצהר חוזר `null` — ולא אפס', async () => {
    const { host } = fakeDoc({
      selection: RANGE_OVER_THREE,
      get: bodyOf([
        ['p1', { spacing: { before: 6 } }],
        ['p2', undefined],
        ['p3', undefined],
      ]),
    });

    const entries = await readSelectionSpacing(host);

    // זו ההבחנה שכל המודול נשען עליה: `null` = יורש מהסגנון, `0` = הוצהר אפס.
    expect(entries?.[0].spacing).toEqual({
      beforeTwips: 120,
      afterTwips: null,
      lineTwips: null,
      rule: null,
    });
    expect(entries?.[1].spacing).toEqual({
      beforeTwips: null,
      afterTwips: null,
      lineTwips: null,
      rule: null,
    });
  });

  it('`nodeType` נקרא מ-`blocks.list` בקריאה אחת, ולא אחת לכל פסקה', async () => {
    const { host, doc } = fakeDoc({
      selection: RANGE_OVER_THREE,
      get: bodyOf([['p1', undefined], ['p2', undefined], ['p3', undefined]]),
      blocks: [
        { nodeId: 'p1', nodeType: 'heading' },
        { nodeId: 'p2', nodeType: 'listItem' },
      ],
    });
    const listSpy = vi.spyOn((doc as unknown as { blocks: { list: () => unknown } }).blocks, 'list');

    const entries = await readSelectionSpacing(host);

    expect(entries?.map((entry) => entry.target.nodeType)).toEqual([
      'heading',
      'listItem',
      // מזהה שאינו ברשימה נופל ל„פסקה” ואינו קורא שוב.
      'paragraph',
    ]);
    expect(listSpy).toHaveBeenCalledTimes(1);
  });

  it('בלי בחירה — `null`, ובלי הודעה', async () => {
    const { host } = fakeDoc({ selection: { target: { kind: 'text', segments: [] } }, get: bodyOf([]) });
    expect(await readSelectionSpacing(host)).toBeNull();
  });
});

describe('applySelectionSpacing', () => {
  it('נשלח רק מה שהיה מוצהר, ועליו הצד שהשתנה', async () => {
    const { host, calls } = fakeDoc({
      ops: { setSpacing: ok },
      selection: RANGE_OVER_THREE,
      get: bodyOf([
        ['p1', { spacing: { line: 18, lineRule: 'exact' } }],
        ['p2', undefined],
      ]),
    });
    const entries = await readSelectionSpacing(host);

    await applySelectionSpacing(host, entries ?? [], { beforeTwips: 240 });

    const sent = calls.get('setSpacing') as Record<string, unknown>[];
    // מרווח השורות של p1 נשלח בחזרה כפי שהיה — בלעדיו `setSpacing` היה מוחק אותו.
    expect(sent[0]).toEqual({
      target: { kind: 'block', nodeType: 'paragraph', nodeId: 'p1' },
      before: 240,
      line: 360,
      lineRule: 'exact',
    });
    // ו-p2 לא הצהיר דבר, ולכן נשלח בו רק מה שהמשתמש ביקש: קיבוע מרווח שורות
    // שיורש מהסגנון היה מנתק את הפסקה ממנו לתמיד.
    expect(sent[1]).toEqual({
      target: { kind: 'block', nodeType: 'paragraph', nodeId: 'p2' },
      before: 240,
    });
  });

  it('כשל בפסקה אחת עוצר, ואינו משאיר את השאר חצי-מוחל בשקט', async () => {
    let call = 0;
    const { host, calls } = fakeDoc({
      ops: {
        setSpacing: () => {
          call += 1;
          return call === 2 ? { success: false, failure: { code: 'LOCKED', message: 'נעול' } } : { success: true };
        },
      },
      selection: RANGE_OVER_THREE,
      get: bodyOf([['p1', undefined], ['p2', undefined], ['p3', undefined]]),
    });
    const entries = await readSelectionSpacing(host);

    const outcome = await applySelectionSpacing(host, entries ?? [], { afterTwips: 240 });

    expect(outcome.ok).toBe(false);
    expect(calls.get('setSpacing')).toHaveLength(2);
  });

  it('בלי פסקאות — כשל „יש למקם את הסמן”, ובלי קריאה למנוע', async () => {
    const { host, calls } = fakeDoc({ ops: { setSpacing: ok } });

    const outcome = await applySelectionSpacing(host, [], { beforeTwips: 240 });

    expect(outcome).toEqual({
      ok: false,
      message: 'יש למקם את הסמן במסמך',
      reason: 'selection-required',
    });
    expect(calls.get('setSpacing')).toHaveLength(0);
  });
});

describe('toggleSelectionSpacing', () => {
  function withHead(props: Record<string, unknown> | undefined) {
    return fakeDoc({
      ops: { setSpacing: ok },
      selection: RANGE_OVER_THREE,
      get: bodyOf([['p1', props], ['p2', undefined], ['p3', undefined]]),
    });
  }

  it('אין רווח → מוסיף את הצעד, בכל הפסקאות', async () => {
    const { host, calls } = withHead(undefined);

    const outcome = await toggleSelectionSpacing(host, 'before', 240);

    expect(outcome.ok).toBe(true);
    const sent = calls.get('setSpacing') as Record<string, unknown>[];
    expect(sent).toHaveLength(3);
    expect(sent.every((input) => input.before === 240)).toBe(true);
  });

  it('יש רווח → מסיר, וההכרעה נופלת על הפסקה הראשונה בלבד', async () => {
    const { host, calls } = withHead({ spacing: { before: 12 } });

    await toggleSelectionSpacing(host, 'before', 240);

    const sent = calls.get('setSpacing') as Record<string, unknown>[];
    // כל השלוש מקבלות אפס — גם אלה שלא היה בהן רווח מלכתחילה. בחירה מעורבת
    // שהייתה מוכרעת לכל פסקה בנפרד הייתה נשארת מעורבת, רק הפוך.
    expect(sent.map((input) => input.before)).toEqual([0, 0, 0]);
  });

  it('„אחרי” הוא צד נפרד, ואינו נוגע ב„לפני”', async () => {
    const { host, calls } = withHead({ spacing: { before: 12 } });

    await toggleSelectionSpacing(host, 'after', 240);

    const sent = calls.get('setSpacing') as Record<string, unknown>[];
    expect(sent[0]).toEqual({
      target: { kind: 'block', nodeType: 'paragraph', nodeId: 'p1' },
      before: 240,
      after: 240,
    });
  });

  it('בלי סמן — כשל, ובלי כתיבה', async () => {
    const { host, calls } = fakeDoc({
      ops: { setSpacing: ok },
      selection: { target: { kind: 'text', segments: [] } },
      get: bodyOf([]),
    });

    const outcome = await toggleSelectionSpacing(host, 'before', 240);

    expect(outcome).toMatchObject({ ok: false, reason: 'selection-required' });
    expect(calls.get('setSpacing')).toHaveLength(0);
  });
});

describe('applyParagraphContextualSpacing', () => {
  it('נשלח `contextualSpacing` בלבד — ולא מצב מלא של אפשרויות הזרימה', async () => {
    const { host, calls } = fakeDoc({ ops: { setFlowOptions: ok } });

    const outcome = await applyParagraphContextualSpacing(host, { nodeId: 'p3' }, true);

    expect(outcome.ok).toBe(true);
    // `setFlowOptions` הוא patch (נמדד), ולכן מפתח שאינו נשלח אינו נוגע במה
    // שקיים. שליחת `pageBreakBefore: false` „ליתר ביטחון” הייתה מכבה מעבר
    // עמוד שהמשתמש הגדיר.
    expect(calls.get('setFlowOptions')?.[0]).toEqual({
      target: { nodeId: 'p3' },
      contextualSpacing: true,
    });
  });

  it('כיבוי נשלח כ-`false` מפורש, ולא בהשמטה', async () => {
    const { host, calls } = fakeDoc({ ops: { setFlowOptions: ok } });

    await applyParagraphContextualSpacing(host, { nodeId: 'p3' }, false);

    expect(calls.get('setFlowOptions')?.[0]).toEqual({
      target: { nodeId: 'p3' },
      contextualSpacing: false,
    });
  });

  it('גרסה בלי הפעולה מדווחת „אינו זמין” ואינה זורקת', async () => {
    const { host } = fakeDoc({ ops: { setFlowOptions: null } });

    const outcome = await applyParagraphContextualSpacing(host, { nodeId: 'p3' }, true);

    expect(outcome).toMatchObject({ ok: false, reason: 'command-unsupported' });
  });
});

/**
 * יחידת `spacing.line` מהמודל — שתי יחידות לפי `lineRule`.
 *
 * הטבלה כאן היא בדיוק מה שנמדד ב-`scripts/line-unit-probe.mjs` על המנוע
 * האמיתי: ב-`auto` המודל מחזיר כפולה, ובשאר נקודות. ההמרה האחידה שהייתה
 * כאן קודם הקטינה כל מרווח אוטומטי פי 12, ושתי תקלות שקטות נגזרו ממנה —
 * ראו `lineTwipsFromModel`.
 */
describe('יחידת מרווח השורות שהמודל מחזיר', () => {
  const declaring = (spacing: Record<string, unknown>) =>
    fakeDoc({
      selection: RANGE_OVER_THREE,
      get: bodyOf([['p1', { spacing }], ['p2', undefined], ['p3', undefined]]),
    });

  it('`auto` — המודל מחזיר כפולה, ו-twips הם ×240', async () => {
    const { host } = declaring({ line: 3, lineRule: 'auto' });

    const entries = await readSelectionSpacing(host);

    expect(entries?.[0].spacing.lineTwips).toBe(720);
  });

  it('`exact` — המודל מחזיר נקודות, ו-twips הם ×20', async () => {
    const { host } = declaring({ line: 18, lineRule: 'exact' });

    const entries = await readSelectionSpacing(host);

    expect(entries?.[0].spacing.lineTwips).toBe(360);
  });

  it('`line` בלי `lineRule` נקרא כ-`auto` — כמו ב-OOXML', async () => {
    const { host } = declaring({ line: 1.5 });

    const entries = await readSelectionSpacing(host);

    expect(entries?.[0].spacing.lineTwips).toBe(360);
    expect(entries?.[0].spacing.rule).toBeNull();
  });

  it('סבב קריאה-כתיבה-קריאה אינו מכווץ את המרווח', async () => {
    // זו התקלה שהשער תפס: 3 → 60 → 0.25 → 5. הכתיבה חייבת להחזיר את אותם
    // twips שנקראו, אחרת כל „הוסף רווח” מקטין את מרווח השורות בשקט.
    const { host, calls } = fakeDoc({
      ops: { setSpacing: ok },
      selection: RANGE_OVER_THREE,
      get: bodyOf([['p1', { spacing: { line: 3, lineRule: 'auto' } }]]),
    });
    const entries = await readSelectionSpacing(host);

    await applySelectionSpacing(host, entries ?? [], { beforeTwips: 240 });

    expect((calls.get('setSpacing')?.[0] as Record<string, unknown>).line).toBe(720);
  });

  it('הדיאלוג נפתח על המרווח האמיתי, ולא על „בודדת”', async () => {
    // `readParagraphFormat` הוא מה שממלא את הבורר בדיאלוג. 1.5 שהוחזר כ-30
    // twips לא התאים לאף אפשרות, הבורר נפל ל-240, ו„אישור” כתב 240 — כלומר
    // שינה את המסמך בלי שאיש ביקש.
    const { host } = fakeDoc({
      get: {
        body: [
          {
            id: 'p3',
            kind: 'paragraph',
            paragraph: { inlines: [], props: { spacing: { line: 1.5, lineRule: 'auto' } } },
          },
        ],
      },
    });

    const result = await readParagraphFormat(host);

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.snapshot.spacing.lineTwips).toBe(360);
  });
});
