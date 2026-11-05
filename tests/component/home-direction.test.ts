/** Direction readouts must follow document changes even when the engine emits no selection/onUpdate event. */
import { describe, expect, it, vi } from 'vitest';
import { invalidateDocumentReads } from '../../src/engine/document-read-state';
import type { ParagraphFormatDocumentApi } from '../../src/engine/paragraph-format';
import HomeTab from '../../src/ui/ribbon/tabs/HomeTab.vue';
import { autoUnmount, buttonByTip, createSuperdocDouble, mountUi, settle } from './harness';

autoUnmount();

const RTL_TIP = 'כיוון פסקה מימין לשמאל';
const LTR_TIP = 'כיוון פסקה משמאל לימין';

describe('כיוון פסקה — חיווי „פעיל” נקרא מה-pPr, לא מהפקודה', () => {
  it('לחיצה על „מימין לשמאל” מדליקה את החיווי מיד, בלי לחכות לתזוזת סמן', async () => {
    // הכפיל מייצג את מה ש-doc.get() מחזיר **אחרי** כתיבה מוצלחת של <w:bidi/>
    // — בדיוק כמו שנמדד במנוע האמיתי (ראו ההערה בראש הקובץ, ובראש
    // paragraphBidi ב-HomeTab.vue).
    const superdoc = createSuperdocDouble({ paragraphProps: { bidi: true } });
    const harness = mountUi(HomeTab, { superdoc });
    await settle();

    // לפני הלחיצה: אין עדיין קריאה שהתיישבה (ה-debounce עוד לא רץ), ולכן
    // נופלים לברירת המחדל של כפיל הפקודה — כמו לפני התיקון.
    expect(buttonByTip(harness.wrapper, RTL_TIP).attributes('aria-pressed')).toBe('false');

    await buttonByTip(harness.wrapper, RTL_TIP).trigger('click');
    await settle();

    expect(buttonByTip(harness.wrapper, RTL_TIP).attributes('aria-pressed')).toBe('true');
    expect(buttonByTip(harness.wrapper, LTR_TIP).attributes('aria-pressed')).toBe('false');
    expect(harness.failures()).toEqual([]);
  });

  it('לחיצה על „משמאל לימין” מדליקה אותו ומכבה את „מימין לשמאל”', async () => {
    const superdoc = createSuperdocDouble({ paragraphProps: { bidi: false } });
    const harness = mountUi(HomeTab, { superdoc });
    await settle();

    await buttonByTip(harness.wrapper, LTR_TIP).trigger('click');
    await settle();

    expect(buttonByTip(harness.wrapper, LTR_TIP).attributes('aria-pressed')).toBe('true');
    expect(buttonByTip(harness.wrapper, RTL_TIP).attributes('aria-pressed')).toBe('false');
    expect(harness.failures()).toEqual([]);
  });

  it('מסמך שנפתח עם bidi כבר קיים מציג „פעיל” אחרי שהשהיית תזוזת הסמן חולפת — בלי לחיצה כלל', async () => {
    // זה בדיוק המקרה שהחיווי הישן (מ-`dirRtlCmd.active`) אף פעם לא היה
    // מראה: פסקה שהגיעה RTL מ-Word, בלי שהמשתמש לחץ על שום דבר.
    vi.useFakeTimers();
    try {
      const superdoc = createSuperdocDouble({ paragraphProps: { bidi: true } });
      const harness = mountUi(HomeTab, { superdoc });
      await settle();
      await vi.advanceTimersByTimeAsync(500);
      await settle();

      expect(buttonByTip(harness.wrapper, RTL_TIP).attributes('aria-pressed')).toBe('true');
      expect(buttonByTip(harness.wrapper, LTR_TIP).attributes('aria-pressed')).toBe('false');
    } finally {
      vi.useRealTimers();
    }
  });

  it('מעבר למסמך אחר עם bidi שונה מתעדכן אחרי אותה השהיה', async () => {
    vi.useFakeTimers();
    try {
      const rtlDoc = createSuperdocDouble({ paragraphProps: { bidi: true } });
      const harness = mountUi(HomeTab, { superdoc: rtlDoc });
      await settle();
      await vi.advanceTimersByTimeAsync(500);
      await settle();
      expect(buttonByTip(harness.wrapper, RTL_TIP).attributes('aria-pressed')).toBe('true');

      const ltrDoc = createSuperdocDouble({ paragraphProps: { bidi: false } });
      await harness.setSuperdoc(ltrDoc);
      await vi.advanceTimersByTimeAsync(500);
      await settle();

      expect(buttonByTip(harness.wrapper, RTL_TIP).attributes('aria-pressed')).toBe('false');
      expect(buttonByTip(harness.wrapper, LTR_TIP).attributes('aria-pressed')).toBe('true');
    } finally {
      vi.useRealTimers();
    }
  });

  it('אין סמן במסמך: readParagraphIndents מחזירה null, והחיווי נופל לברירת המחדל בלי לזרוק', async () => {
    vi.useFakeTimers();
    try {
      const superdoc = createSuperdocDouble({ selection: { blockId: null } });
      const harness = mountUi(HomeTab, { superdoc });
      await settle();
      await vi.advanceTimersByTimeAsync(500);
      await settle();

      expect(buttonByTip(harness.wrapper, RTL_TIP).attributes('aria-pressed')).toBe('false');
      expect(buttonByTip(harness.wrapper, LTR_TIP).attributes('aria-pressed')).toBe('false');
    } finally {
      vi.useRealTimers();
    }
  });
  it('refreshes after formatting/Undo/Redo invalidation without a selection event', async () => {
    vi.useFakeTimers();
    try {
      const props = { bidi: true };
      const superdoc = createSuperdocDouble({ paragraphProps: props });
      const doc = superdoc.host.activeEditor!.doc;
      const harness = mountUi(HomeTab, { superdoc });
      await vi.advanceTimersByTimeAsync(150);
      await settle();
      expect(buttonByTip(harness.wrapper, RTL_TIP).attributes('aria-pressed')).toBe('true');
      for (const bidi of [false, true, false]) {
        props.bidi = bidi;
        invalidateDocumentReads(doc);
        await vi.advanceTimersByTimeAsync(150);
        await settle();
        expect(buttonByTip(harness.wrapper, RTL_TIP).attributes('aria-pressed')).toBe(String(bidi));
        expect(buttonByTip(harness.wrapper, LTR_TIP).attributes('aria-pressed')).toBe(String(!bidi));
      }
      expect(superdoc.inputs('get')).toEqual([]);
    } finally {
      vi.useRealTimers();
    }
  });

  it('rejects a pending old paragraph as soon as selection moves, while the next read is debounced', async () => {
    vi.useFakeTimers();
    try {
      const superdoc = createSuperdocDouble();
      const doc = superdoc.host.activeEditor!.doc as unknown as ParagraphFormatDocumentApi;
      let notifySelection!: () => void;
      const selection = superdoc.host.ui.selection as unknown as { observe: (callback: () => void) => () => void };
      const unsubscribe = vi.fn();
      selection.observe = (callback) => { notifySelection = callback; return unsubscribe; };
      let blockId = 'p1';
      doc.selection = { current: () => ({ target: { segments: [{ blockId, range: { start: 0, end: 0 } }] } }) };
      let resolve!: (value: unknown) => void;
      doc.getNodeById = () => new Promise((done) => { resolve = done; });
      const harness = mountUi(HomeTab, { superdoc });
      await vi.advanceTimersByTimeAsync(150);
      blockId = 'p2';
      notifySelection();
      resolve({ node: { kind: 'paragraph', paragraph: { props: { bidi: true } } }, address: { kind: 'block', nodeType: 'paragraph', nodeId: 'p1' } });
      await settle();
      expect(buttonByTip(harness.wrapper, RTL_TIP).attributes('aria-pressed')).toBe('false');
      await vi.advanceTimersByTimeAsync(150);
      resolve({ node: { kind: 'paragraph', paragraph: { props: { bidi: false } } }, address: { kind: 'block', nodeType: 'paragraph', nodeId: 'p2' } });
      await settle();
      expect(buttonByTip(harness.wrapper, LTR_TIP).attributes('aria-pressed')).toBe('true');
      harness.wrapper.unmount();
      expect(unsubscribe).toHaveBeenCalledTimes(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it('does not publish a pending read from the previous document after switching tabs', async () => {
    vi.useFakeTimers();
    try {
      const old = createSuperdocDouble();
      const doc = old.host.activeEditor!.doc as unknown as ParagraphFormatDocumentApi;
      let resolve!: (value: unknown) => void;
      doc.getNode = (address) => new Promise((done) => { resolve = () => done({ node: { kind: 'paragraph', paragraph: { props: { bidi: true } } }, address }); });
      const harness = mountUi(HomeTab, { superdoc: old });
      await vi.advanceTimersByTimeAsync(150);
      await harness.setSuperdoc(createSuperdocDouble({ paragraphProps: { bidi: false } }));
      await vi.advanceTimersByTimeAsync(150);
      await settle();
      expect(buttonByTip(harness.wrapper, LTR_TIP).attributes('aria-pressed')).toBe('true');
      resolve(undefined);
      await settle();
      expect(buttonByTip(harness.wrapper, LTR_TIP).attributes('aria-pressed')).toBe('true');
    } finally {
      vi.useRealTimers();
    }
  });

});
