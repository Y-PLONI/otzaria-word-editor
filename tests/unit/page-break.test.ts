/**
 * „התחל בעמוד חדש”. הבדיקה היא על **מה נשלח למנוע** ועל הדיווח — שההחלה
 * עצמה עובדת נבדק באימות בדפדפן.
 *
 * מה שנבדק כאן במיוחד: `NO_OP` אינו שגיאה, וכל מסלול כשל מחזיר תוצאה מטופסת
 * ולא זריקה. הוולידטורים של ה-Document API זורקים `INVALID_INPUT` על קלט
 * פסול במקום להחזיר קבלה, וחריגה מפקד ב-Ribbon מפילה את רינדור הרצועה כולה.
 */
import { describe, expect, it, vi } from 'vitest';
import {
  PAGE_BREAK_OPERATION,
  readPageBreakSupport,
  startParagraphOnNewPage,
  setParagraphPageBreak,
  toggleParagraphPageBreak,
  readPageBreakNodeId,
  createPageBreakTracker,
  type PageBreakDocumentApi,
} from '../../src/engine/page-break';

const CARET = {
  target: { kind: 'text', segments: [{ blockId: 'p3', range: { start: 2, end: 2 } }] },
};

/**
 * מופע כפול. `flow` הוא מה ש-`setFlowOptions` יחזיר, ו-`calls` אוסף את הקלט
 * כדי שהבדיקה תראה מה בדיוק נשלח.
 */
function fakeDoc(
  options: {
    flow?: (input: unknown) => unknown;
    available?: boolean;
    capabilities?: PageBreakDocumentApi['capabilities'];
    selection?: unknown;
    omitFlow?: boolean;
  } = {},
) {
  // הדחיפה ל-`calls` נעשית **רק** בעטיפה, ולא גם ב-`flow`: כפילות שם הייתה
  // מייצרת שתי רשומות לקריאה אחת, וטענה על „נקרא פעם אחת” הייתה עוברת בטעות.
  const calls: unknown[] = [];
  const flow = options.flow ?? (() => ({ success: true }));

  const doc = {
    selection: { current: vi.fn(async () => options.selection ?? CARET) },
    capabilities:
      options.capabilities === undefined
        ? {
            get: async () => ({
              operations: { [PAGE_BREAK_OPERATION]: { available: options.available ?? true } },
            }),
          }
        : options.capabilities,
    ...(options.omitFlow
      ? {}
      : {
          format: {
            paragraph: {
              setFlowOptions: (input: unknown) => {
                calls.push(input);
                return flow(input) as never;
              },
            },
          },
        }),
  } as unknown as PageBreakDocumentApi;

  return { doc, calls, host: { activeEditor: { doc } } };
}

describe('readPageBreakSupport', () => {
  it('זמין כשהמנוע מדווח שהפעולה זמינה', async () => {
    const { host } = fakeDoc();

    await expect(readPageBreakSupport(host)).resolves.toEqual({
      available: true,
      explanation: '',
    });
  });

  it('פעולה שאינה זמינה מקבלת את הנוסח של §12', async () => {
    const { host } = fakeDoc({ available: false });

    await expect(readPageBreakSupport(host)).resolves.toEqual({
      available: false,
      explanation: 'הפעולה אינה זמינה בגרסה הזאת של המנוע',
    });
  });

  it('פעולה שאינה בטבלה בכלל = גרסה שאינה מכירה אותה', async () => {
    const { host } = fakeDoc({ capabilities: { get: async () => ({ operations: {} }) } });

    await expect(readPageBreakSupport(host)).resolves.toMatchObject({ available: false });
  });

  it('גרסה שאינה חושפת setFlowOptions מסומנת „אינו זמין בגרסה זו”', async () => {
    const { host } = fakeDoc({ omitFlow: true });

    await expect(readPageBreakSupport(host)).resolves.toEqual({
      available: false,
      explanation: 'אינו זמין בגרסה זו',
    });
  });

  it('אין Document API — „המסמך עדיין נטען”, ולא זריקה', async () => {
    for (const host of [null, undefined, { activeEditor: null }, { activeEditor: { doc: null } }]) {
      await expect(readPageBreakSupport(host)).resolves.toEqual({
        available: false,
        explanation: 'המסמך עדיין נטען',
      });
    }
  });

  it('אין capabilities לשאול — נכשל סגור', async () => {
    const { host } = fakeDoc({ capabilities: {} });

    await expect(readPageBreakSupport(host)).resolves.toMatchObject({ available: false });
  });

  it('קריאת יכולות שזורקת אינה מפילה את הרצועה', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const { host } = fakeDoc({
      capabilities: {
        get: () => {
          throw new Error('boom');
        },
      },
    });

    await expect(readPageBreakSupport(host)).resolves.toMatchObject({ available: false });
    expect(warn).toHaveBeenCalled();
    warn.mockRestore();
  });

  it('תשובה שאינה אובייקט אינה תשובה', async () => {
    const { host } = fakeDoc({ capabilities: { get: async () => undefined } });

    await expect(readPageBreakSupport(host)).resolves.toMatchObject({ available: false });
  });
});

describe('startParagraphOnNewPage', () => {
  it('שולחת pageBreakBefore על הפסקה שבה הסמן', async () => {
    const { host, calls } = fakeDoc();

    await expect(startParagraphOnNewPage(host)).resolves.toEqual({ ok: true });
    expect(calls).toEqual([
      {
        target: { kind: 'block', nodeType: 'paragraph', nodeId: 'p3' },
        pageBreakBefore: true,
      },
    ]);
  });

  it('`nodeType: paragraph` גם לכותרת — פתרון היעד לפי nodeId בלבד', async () => {
    // זה מה ש-`paragraphTarget` של ה-controller עצמו שולח לכל
    // `format.paragraph.*`, ולכן זו הצורה הנכונה ולא קיצור דרך.
    const { host, calls } = fakeDoc({
      selection: { target: { kind: 'text', segments: [{ blockId: 'h1', range: { start: 0, end: 0 } }] } },
    });

    await startParagraphOnNewPage(host);

    expect(calls[0]).toMatchObject({ target: { nodeType: 'paragraph', nodeId: 'h1' } });
  });

  it('בבחירה על כמה פסקאות מוחל על זו שהבחירה מתחילה בה', async () => {
    const { host, calls } = fakeDoc({
      selection: {
        target: {
          kind: 'text',
          segments: [
            { blockId: 'p1', range: { start: 4, end: 9 } },
            { blockId: 'p2', range: { start: 0, end: 3 } },
          ],
        },
      },
    });

    await startParagraphOnNewPage(host);

    expect(calls[0]).toMatchObject({ target: { nodeId: 'p1' } });
  });

  it('story נשלח כשהבחירה אינה בגוף המסמך', async () => {
    const story = { kind: 'story', storyType: 'headerFooterSlot' };
    const { host, calls } = fakeDoc({
      selection: {
        target: { kind: 'text', segments: [{ blockId: 'h1', range: { start: 0, end: 0 } }], story },
      },
    });

    await startParagraphOnNewPage(host);

    expect(calls[0]).toMatchObject({ target: { story } });
  });

  it('בגוף המסמך `story` אינו נשלח בכלל', async () => {
    // `story: null` מפורש היה נכשל בוולידציה; היעדרו פירושו גוף המסמך.
    const { host, calls } = fakeDoc();

    await startParagraphOnNewPage(host);

    const target = (calls[0] as { target: Record<string, unknown> }).target;
    expect(Object.prototype.hasOwnProperty.call(target, 'story')).toBe(false);
    expect(Object.keys(target).sort()).toEqual(['kind', 'nodeId', 'nodeType']);
  });

  it('NO_OP הוא הצלחה — הפסקה כבר מתחילה בעמוד חדש', async () => {
    // לחיצה שנייה אינה אמורה להראות שגיאה. `possibleFailureCodes` של הפעולה
    // הוא `['NO_OP']` בלבד.
    const { host } = fakeDoc({
      flow: () => ({ success: false, failure: { code: 'NO_OP', message: 'no changes' } }),
    });

    await expect(startParagraphOnNewPage(host)).resolves.toEqual({ ok: true });
  });

  it('קבלה שנכשלה מתורגמת לעברית עם הקוד', async () => {
    const { host } = fakeDoc({
      flow: () => ({ success: false, failure: { code: 'DOCUMENT_READONLY', message: 'readonly' } }),
    });

    const outcome = await startParagraphOnNewPage(host);

    expect(outcome).toMatchObject({ ok: false, reason: 'DOCUMENT_READONLY' });
    if (!outcome.ok) expect(outcome.message).toContain('לקריאה בלבד');
  });

  it('קוד שאין לו תרגום מוצג עם ההסבר של המנוע — ולא נעלם', async () => {
    const { host } = fakeDoc({
      flow: () => ({ success: false, failure: { code: 'WEIRD_CODE', message: 'something odd' } }),
    });

    const outcome = await startParagraphOnNewPage(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.message).toContain('something odd');
      expect(outcome.message).toContain('WEIRD_CODE');
    }
  });

  it('Promise נתמך בדיוק כמו קבלה סינכרונית', async () => {
    const { host } = fakeDoc({ flow: () => Promise.resolve({ success: true }) });

    await expect(startParagraphOnNewPage(host)).resolves.toEqual({ ok: true });
  });

  it('זריקה הופכת להודעה ולא מפילה את הרצועה', async () => {
    const { host } = fakeDoc({
      flow: () => {
        throw new Error('INVALID_INPUT: target is required');
      },
    });

    const outcome = await startParagraphOnNewPage(host);

    expect(outcome).toMatchObject({ ok: false, reason: 'threw' });
    if (!outcome.ok) expect(outcome.message).toContain('INVALID_INPUT');
  });

  it('בלי סמן במסמך — הודעה מדויקת ולא קריאה למנוע', async () => {
    const { host, calls } = fakeDoc({ selection: { target: null } });

    await expect(startParagraphOnNewPage(host)).resolves.toEqual({
      ok: false,
      message: 'יש למקם את הסמן במסמך',
      reason: 'selection-required',
    });
    expect(calls).toEqual([]);
  });

  it('גרסה שאינה חושפת את הפעולה מדווחת „אינו זמין בגרסה זו”', async () => {
    const { host } = fakeDoc({ omitFlow: true });

    const outcome = await startParagraphOnNewPage(host);

    expect(outcome).toMatchObject({ ok: false, reason: 'command-unsupported' });
    if (!outcome.ok) expect(outcome.message).toContain('אינו זמין בגרסה זו');
  });

  it('אין Document API — תוצאה מטופסת, לא זריקה', async () => {
    await expect(startParagraphOnNewPage(null)).resolves.toMatchObject({ ok: false });
    await expect(startParagraphOnNewPage({ activeEditor: { doc: null } })).resolves.toMatchObject({
      ok: false,
    });
  });
});

describe('readPageBreakNodeId', () => {
  it('מזהה הפסקה שהסמן בה', async () => {
    const { host } = fakeDoc();
    await expect(readPageBreakNodeId(host)).resolves.toBe('p3');
  });

  it('null כשאין בחירה — לא זורקת', async () => {
    const { host } = fakeDoc({ selection: { target: null } });
    await expect(readPageBreakNodeId(host)).resolves.toBeNull();
  });

  it('null כשאין Document API', async () => {
    await expect(readPageBreakNodeId(null)).resolves.toBeNull();
  });
});

describe('createPageBreakTracker', () => {
  it('פסקה שלא נזכרה מעולם — כבויה', () => {
    const tracker = createPageBreakTracker();
    expect(tracker.isOn('p3')).toBe(false);
    expect(tracker.isOn(null)).toBe(false);
    expect(tracker.isOn(undefined)).toBe(false);
  });

  it('remember קובע את המצב שנקרא ב-isOn, לכל nodeId בנפרד', () => {
    const tracker = createPageBreakTracker();
    tracker.remember('p3', true);
    expect(tracker.isOn('p3')).toBe(true);
    expect(tracker.isOn('p4')).toBe(false);

    tracker.remember('p3', false);
    expect(tracker.isOn('p3')).toBe(false);
  });

  it('remember עם nodeId ריק/חסר אינה עושה כלום', () => {
    const tracker = createPageBreakTracker();
    tracker.remember(null, true);
    tracker.remember(undefined, true);
    tracker.remember('', true);
    expect(tracker.isOn(null)).toBe(false);
    expect(tracker.isOn('')).toBe(false);
  });

  it('syncDocument אינה מאפסת כשה-host זהה (אותו מסמך, לשונית שהוחלפה וחזרה)', () => {
    const tracker = createPageBreakTracker();
    const host = {};
    tracker.syncDocument(host);
    tracker.remember('p3', true);
    tracker.syncDocument(host);
    expect(tracker.isOn('p3')).toBe(true);
  });

  it('syncDocument מאפסת כשה-host שונה (מסמך אחר נפתח, או נסגר ל-null)', () => {
    const tracker = createPageBreakTracker();
    tracker.syncDocument({});
    tracker.remember('p3', true);
    expect(tracker.isOn('p3')).toBe(true);

    tracker.syncDocument({});
    expect(tracker.isOn('p3')).toBe(false);

    tracker.remember('p3', true);
    tracker.syncDocument(null);
    expect(tracker.isOn('p3')).toBe(false);
  });

  /**
   * ממצא QA: `syncDocument` הוחלפה מהשוואת זהות `host` להשוואת
   * `DOCUMENT_GENERATION` (composables/keys.ts, נגזר מ-`EditorSwap.
   * documentGeneration`) — מונה עולה שמקורו אחד ומוסמך (sessions/editor-swap.ts)
   * ולא נשען על התנהגות לא-מתועדת של `SuperDoc`. הבדיקות כאן על התרחיש
   * שה-QA חשש ממנו: מפתח (host, בעבר) שחוזר על עצמו בין "מסמכים" — מונה
   * עולה, בניגוד לזהות אובייקט, לעולם לא חוזר על עצמו, ולכן תמיד מזהה נכון.
   */
  it('syncDocument עם generation עולה (מספרים) מאפסת גם אם "host" ישן היה חוזר על עצמו', () => {
    const tracker = createPageBreakTracker();

    tracker.syncDocument(1); // דור 1 — "מסמך" ראשון
    tracker.remember('p3', true);
    expect(tracker.isOn('p3')).toBe(true);

    // "מסמך אחר" נפתח: generation עולה תמיד, גם אם אובייקט ה-host התאורטי
    // מאחוריו היה יכול (בטעות, או במימוש מנוע עתידי) לחזור על עצמו — בדיוק
    // התרחיש שהשוואת זהות בלבד הייתה עלולה לפספס.
    tracker.syncDocument(2); // דור 2
    expect(tracker.isOn('p3')).toBe(false);

    tracker.remember('p3', true);
    tracker.syncDocument(3); // דור 3 — עדיין עולה, עדיין מאפס
    expect(tracker.isOn('p3')).toBe(false);
  });

  it('syncDocument עם אותו generation (אותו מסמך) אינה מאפסת, גם בקריאות חוזרות', () => {
    const tracker = createPageBreakTracker();

    tracker.syncDocument(5);
    tracker.remember('p3', true);
    tracker.syncDocument(5);
    tracker.syncDocument(5);

    expect(tracker.isOn('p3')).toBe(true);
  });

  it('forgetAll מוחקת הכול בלי קשר למסמך — למסלול Undo/Redo (App.vue)', () => {
    const tracker = createPageBreakTracker();
    tracker.syncDocument(1);
    tracker.remember('p3', true);
    tracker.remember('p4', true);

    tracker.forgetAll();

    expect(tracker.isOn('p3')).toBe(false);
    expect(tracker.isOn('p4')).toBe(false);

    // generation לא השתנתה — forgetAll אינה תלויה בה כלל, ואינה מעדכנת אותה.
    tracker.remember('p3', true);
    tracker.syncDocument(1);
    expect(tracker.isOn('p3')).toBe(true);
  });
});

describe('setParagraphPageBreak', () => {
  it('מתג: שולחת true, וה-tracker זוכר שהפסקה פעילה', async () => {
    const tracker = createPageBreakTracker();
    const { host, calls } = fakeDoc();

    const outcome = await setParagraphPageBreak(host, true, tracker);

    expect(outcome).toEqual({ ok: true });
    expect(calls).toEqual([
      { target: { kind: 'block', nodeType: 'paragraph', nodeId: 'p3' }, pageBreakBefore: true },
    ]);
    expect(tracker.isOn('p3')).toBe(true);
  });

  it('מתג: שולחת false אחרי true, וה-tracker זוכר שהיא כבויה — זה הביטול', async () => {
    const tracker = createPageBreakTracker();
    const { host, calls } = fakeDoc();

    await setParagraphPageBreak(host, true, tracker);
    const outcome = await setParagraphPageBreak(host, false, tracker);

    expect(outcome).toEqual({ ok: true });
    expect(calls[1]).toEqual({
      target: { kind: 'block', nodeType: 'paragraph', nodeId: 'p3' },
      pageBreakBefore: false,
    });
    expect(tracker.isOn('p3')).toBe(false);
  });

  it('NO_OP נחשבת הצלחה, וה-tracker עדיין מתעדכן', async () => {
    const tracker = createPageBreakTracker();
    const { host } = fakeDoc({
      flow: () => ({ success: false, failure: { code: 'NO_OP', message: 'no changes' } }),
    });

    const outcome = await setParagraphPageBreak(host, true, tracker);

    expect(outcome).toEqual({ ok: true });
    expect(tracker.isOn('p3')).toBe(true);
  });

  it('כשל אמיתי — ה-tracker אינו מתעדכן', async () => {
    const tracker = createPageBreakTracker();
    const { host } = fakeDoc({
      flow: () => ({ success: false, failure: { code: 'DOCUMENT_READONLY', message: 'readonly' } }),
    });

    const outcome = await setParagraphPageBreak(host, true, tracker);

    expect(outcome.ok).toBe(false);
    expect(tracker.isOn('p3')).toBe(false);
  });

  it('הודעת הכשל של הכיבוי שונה מזו של ההפעלה', async () => {
    const tracker = createPageBreakTracker();
    const { host } = fakeDoc({ omitFlow: true });

    const onOutcome = await setParagraphPageBreak(host, true, tracker);
    const offOutcome = await setParagraphPageBreak(host, false, tracker);

    expect(onOutcome.ok).toBe(false);
    expect(offOutcome.ok).toBe(false);
    if (!onOutcome.ok && !offOutcome.ok) {
      expect(onOutcome.message).not.toEqual(offOutcome.message);
      expect(onOutcome.message).toContain('אינו זמין בגרסה זו');
      expect(offOutcome.message).toContain('אינו זמין בגרסה זו');
    }
  });

  it('בלי סמן במסמך — הודעה מדויקת, ולא נכתב כלום', async () => {
    const tracker = createPageBreakTracker();
    const { host, calls } = fakeDoc({ selection: { target: null } });

    const outcome = await setParagraphPageBreak(host, true, tracker);

    expect(outcome).toEqual({
      ok: false,
      message: 'יש למקם את הסמן במסמך',
      reason: 'selection-required',
    });
    expect(calls).toEqual([]);
  });
});

describe('toggleParagraphPageBreak', () => {
  it('derives the toggle from the one selection it writes, independent of a visual readout', async () => {
    const tracker = createPageBreakTracker();
    tracker.remember('p3', true);
    const { host, doc, calls } = fakeDoc();
    const selection = vi.mocked(doc.selection!.current!);
    selection.mockResolvedValueOnce(CARET).mockResolvedValue({
      target: { segments: [{ blockId: 'p4', range: { start: 0, end: 0 } }] },
    });

    await expect(toggleParagraphPageBreak(host, tracker)).resolves.toEqual({ ok: true });
    expect(selection).toHaveBeenCalledTimes(1);
    expect(calls).toEqual([{ target: { kind: 'block', nodeType: 'paragraph', nodeId: 'p3' }, pageBreakBefore: false }]);
    expect(tracker.isOn('p3')).toBe(false);

    await expect(toggleParagraphPageBreak(host, tracker)).resolves.toEqual({ ok: true });
    expect(selection).toHaveBeenCalledTimes(2);
    expect(calls[1]).toMatchObject({ target: { nodeId: 'p4' }, pageBreakBefore: true });
    expect(tracker.isOn('p4')).toBe(true);
  });

  it('keeps tracked state when turning a known paragraph off fails', async () => {
    const tracker = createPageBreakTracker();
    tracker.remember('p3', true);
    const { host } = fakeDoc({ flow: () => ({ success: false, failure: { code: 'DOCUMENT_READONLY' } }) });
    const outcome = await toggleParagraphPageBreak(host, tracker);
    expect(outcome).toMatchObject({ ok: false, reason: 'DOCUMENT_READONLY' });
    if (!outcome.ok) expect(outcome.message).toContain('ביטול התחלת העמוד החדש');
    expect(tracker.isOn('p3')).toBe(true);
  });

  it('does not change tracked state or write when there is no current paragraph', async () => {
    const tracker = createPageBreakTracker();
    tracker.remember('p3', true);
    const { host, calls } = fakeDoc({ selection: { target: null } });
    await expect(toggleParagraphPageBreak(host, tracker)).resolves.toMatchObject({ ok: false, reason: 'selection-required' });
    expect(calls).toEqual([]);
    expect(tracker.isOn('p3')).toBe(true);
  });
});

describe('page-break commands across document replacement', () => {
  for (const mode of ['toggle', 'set'] as const) {
    const run = mode === 'toggle'
      ? toggleParagraphPageBreak
      : (host: Parameters<typeof setParagraphPageBreak>[0], tracker: Parameters<typeof setParagraphPageBreak>[2]) => setParagraphPageBreak(host, false, tracker);

    it(`${mode}: cancels before writing when the tracker switches during selection resolution`, async () => {
      const tracker = createPageBreakTracker();
      tracker.syncDocument('A');
      tracker.remember('p3', true);
      let resolveSelection!: (value: unknown) => void;
      const selection = new Promise(done => { resolveSelection = done; });
      const { host, calls } = fakeDoc({ selection });
      const pending = run(host, tracker);
      tracker.syncDocument('B');
      tracker.remember('p3', false);
      resolveSelection(CARET);

      await expect(pending).resolves.toMatchObject({ ok: false, reason: 'not-ready' });
      expect(calls).toEqual([]);
      expect(tracker.isOn('p3')).toBe(false);
    });

    it(`${mode}: does not remember an old document's completed write in the new document`, async () => {
      const tracker = createPageBreakTracker();
      tracker.syncDocument('A');
      tracker.remember('p3', true);
      let entered!: () => void;
      const started = new Promise<void>(done => { entered = done; });
      let resolveWrite!: (value: unknown) => void;
      const { host, calls } = fakeDoc({ flow: () => {
        entered();
        return new Promise(done => { resolveWrite = done; });
      } });
      const pending = run(host, tracker);
      await started;
      expect(calls).toEqual([{ target: { kind: 'block', nodeType: 'paragraph', nodeId: 'p3' }, pageBreakBefore: false }]);
      tracker.syncDocument('B');
      tracker.remember('p3', true);
      resolveWrite({ success: true });

      await expect(pending).resolves.toEqual({ ok: true });
      expect(tracker.isOn('p3')).toBe(true);
    });
  }
});

describe('PageBreakTracker.onChange', () => {
  /**
   * ממצא QA שני: `forgetAll` נקראת מ-App.vue (אחרי Undo/Redo), מחוץ ל-
   * `InsertTab.vue` שמחזיק את החיווי המוצג — ובלי מנוי, הכפתור נשאר מציג
   * „פעיל” עד לתזוזת סמן שאינה מובטחת (Undo אינו מזיז את הסמן, נמדד).
   * `onChange` הוא התיקון: הרכיב נרשם אליו ומרענן בכל שינוי בידע.
   */
  it('remember מיידעת מאזינים', () => {
    const tracker = createPageBreakTracker();
    const calls: void[] = [];
    tracker.onChange(() => calls.push(undefined));

    tracker.remember('p3', true);

    expect(calls).toHaveLength(1);
  });

  it('forgetAll מיידעת מאזינים', () => {
    const tracker = createPageBreakTracker();
    let notified = false;
    tracker.onChange(() => {
      notified = true;
    });

    tracker.forgetAll();

    expect(notified).toBe(true);
  });

  it('syncDocument מיידעת רק כשהיא באמת מאפסת (מסמך אחר)', () => {
    const tracker = createPageBreakTracker();
    let count = 0;
    tracker.onChange(() => {
      count += 1;
    });

    tracker.syncDocument(1); // ראשון — תמיד מאפס (UNSET → 1)
    expect(count).toBe(1);

    tracker.syncDocument(1); // אותו מסמך — לא מיידעת
    expect(count).toBe(1);

    tracker.syncDocument(2); // מסמך אחר — מיידעת
    expect(count).toBe(2);
  });

  it('פונקציית הביטול מפסיקה את ההודעות', () => {
    const tracker = createPageBreakTracker();
    let count = 0;
    const unsubscribe = tracker.onChange(() => {
      count += 1;
    });

    tracker.remember('p3', true);
    unsubscribe();
    tracker.remember('p3', false);
    tracker.forgetAll();

    expect(count).toBe(1);
  });

  it('מאזין שזורק אינו מפיל מאזינים אחרים ואינו מפיל את הקריאה', () => {
    const tracker = createPageBreakTracker();
    const warn = vi.spyOn(console, 'error').mockImplementation(() => {});
    let secondRan = false;
    tracker.onChange(() => {
      throw new Error('boom');
    });
    tracker.onChange(() => {
      secondRan = true;
    });

    expect(() => tracker.forgetAll()).not.toThrow();
    expect(secondRan).toBe(true);
    warn.mockRestore();
  });
});

describe('PageBreakTracker.forgetAllKeepingSnapshot / restoreSnapshot', () => {
  /**
   * ממצא QA שלישי: Redo לא החזיר את הסימון — `watchUndoRedoKeys` ניקה
   * (`forgetAll`) גם על Redo, ושום דבר לא זכר להחזיר. הבדיקות כאן על
   * המנגנון שמטפל בתרחיש הנפוץ (Undo, ואז Redo מיידי מחזיר בדיוק אותו
   * דבר), ועל מה שקורה כשהוא לא — ראו „א-סימטריית Undo/Redo” ב-
   * engine/page-break.ts למה זו הכרעה מכוונת ולא פספוס.
   */
  it('Undo ואז Redo מיידי משחזר את הידע המדויק', () => {
    const tracker = createPageBreakTracker();
    tracker.remember('p3', true);
    tracker.remember('p4', false);

    tracker.forgetAllKeepingSnapshot();
    expect(tracker.isOn('p3')).toBe(false); // כמו forgetAll רגילה — הידע נעלם עד ה-Redo

    const restored = tracker.restoreSnapshot();

    expect(restored).toBe(true);
    expect(tracker.isOn('p3')).toBe(true);
    expect(tracker.isOn('p4')).toBe(false);
  });

  it('restoreSnapshot היא חד-פעמית — Redo שני בלי Undo נוסף לא מוצא מה להחזיר', () => {
    const tracker = createPageBreakTracker();
    tracker.remember('p3', true);
    tracker.forgetAllKeepingSnapshot();

    expect(tracker.restoreSnapshot()).toBe(true);
    expect(tracker.restoreSnapshot()).toBe(false); // "נצרך" כבר
    expect(tracker.isOn('p3')).toBe(true); // לא נמחק — restoreSnapshot שנכשלה לא נוגעת בידע
  });

  it('restoreSnapshot בלי Undo קודם מחזירה false ואינה עושה כלום', () => {
    const tracker = createPageBreakTracker();
    tracker.remember('p3', true);

    expect(tracker.restoreSnapshot()).toBe(false);
    expect(tracker.isOn('p3')).toBe(true); // נשאר כמו שהיה
  });

  it('שני Undo רצופים: התצלום השני דורס את הראשון (ריק, כי הראשון כבר ניקה)', () => {
    // בדיוק הגבול המתועד: מחסנית אין כאן, רק תא יחיד. Undo שני "רואה" את
    // הידע כפי שהוא אחרי ה-Undo הראשון — כבר ריק — ולכן אין מה להחזיר על
    // Redo כפול שאמור לשחזר למצב שלפני ה-Undo הראשון. זה נשאר safe (isOn
    // מחזירה false), לא שגוי-כלפי-מעלה.
    const tracker = createPageBreakTracker();
    tracker.remember('p3', true);

    tracker.forgetAllKeepingSnapshot(); // תצלום: {p3:true}
    tracker.forgetAllKeepingSnapshot(); // תצלום: {} — דורס את הקודם

    expect(tracker.restoreSnapshot()).toBe(true); // יש תצלום (ריק), אז true
    expect(tracker.isOn('p3')).toBe(false); // אבל הוא לא מכיל את p3
  });

  it('forgetAll רגילה מוחקת גם תצלום ממתין', () => {
    const tracker = createPageBreakTracker();
    tracker.remember('p3', true);
    tracker.forgetAllKeepingSnapshot();

    tracker.forgetAll(); // איפוס מלא ומכוון — לא אמור להשאיר משהו להחזרה

    expect(tracker.restoreSnapshot()).toBe(false);
  });

  it('syncDocument (מסמך אחר) מוחקת תצלום ממתין', () => {
    const tracker = createPageBreakTracker();
    tracker.remember('p3', true);
    tracker.syncDocument(1);
    tracker.forgetAllKeepingSnapshot();

    tracker.syncDocument(2); // מסמך אחר — התצלום מהמסמך הקודם לא רלוונטי

    expect(tracker.restoreSnapshot()).toBe(false);
  });

  it('forgetAllKeepingSnapshot מיידעת מאזינים (onChange), כמו forgetAll', () => {
    const tracker = createPageBreakTracker();
    let count = 0;
    tracker.onChange(() => (count += 1));

    tracker.forgetAllKeepingSnapshot();

    expect(count).toBe(1);
  });

  it('restoreSnapshot מיידעת מאזינים רק כשהיא באמת משחזרת', () => {
    const tracker = createPageBreakTracker();
    let count = 0;
    tracker.forgetAllKeepingSnapshot();
    tracker.onChange(() => (count += 1));

    tracker.restoreSnapshot(); // יש תצלום — מיידעת
    expect(count).toBe(1);

    tracker.restoreSnapshot(); // אין (כבר נצרך) — לא מיידעת
    expect(count).toBe(1);
  });
});
