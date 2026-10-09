/**
 * לשונית „תצוגה” — כפתור כפתור.
 *
 * זהו השער שנדרש אחרי התלונה „רוחב עמוד” ו„100%” לא עובדות: כל פקד בלשונית
 * נלחץ ומה שיצא ממנו נמדד — האירוע, הפקודה, ה-payload, והנטרול בזמן שאין
 * מנוע. חוזה ה-payload עצמו (הוולידטורים של superdoc) נבדק ב-ribbon-payloads;
 * כאן השאלה היא מה **כל לחיצה** עושה.
 *
 * „רוחב עמוד” אינו שולח את `zoom-fit-width` של המנו אלא מחשב את האחוז בעצמו
 * (engine/fit-width.ts — לולאת המשוב שנמדדה במנוע), ולכן הבדיקה שלו היא
 * דו-שלבית: גיאומטריה ידועה בכפיל + מאגס ברוחב ידוע, ואז ה-payload שיצא.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import ViewTab from '../../src/ui/ribbon/tabs/ViewTab.vue';
import { ZOOM_PERCENT_MAX } from '../../src/engine/zoom';
import {
  CANVAS_COLOR_VAR,
  applyCanvasColor,
  setCanvasDefaultColor,
} from '../../src/composables/canvas-color';
import {
  autoUnmount,
  buttonByTip,
  createSuperdocDouble,
  mountUi,
  settle,
  type SuperdocDoubleOptions,
} from './harness';

/**
 * הכתיבה ל-`storage` של אוצריא היא הצד השני של „הצבע נזכר”, והיא אינה
 * נראית ב-DOM. `tryCall` בולעת כשל בשקט מחוץ לאוצריא (host/otzaria-client.ts),
 * ולכן בלי הכפיל הזה הבדיקה הייתה עוברת בירוק גם על פקד ששוכח כל בחירה.
 */
const saveCanvasColor = vi.hoisted(() => vi.fn(async () => {}));
vi.mock('../../src/host/settings', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../src/host/settings')>()),
  saveCanvasColor,
}));

autoUnmount();

const FIT_TITLE = 'התאם את תצוגת העמוד לרוחב החלון';
const HUNDRED_TITLE = 'הצג את המסמך בגודלו האמיתי (100%)';
const RULER_TITLE = 'הצג או הסתר את סרגל המידות';
const MARKS_TITLE = 'הצג סימני פסקאות ותווים נסתרים';

/** A4 באינצ'ים — הצורה ש-`sections.list` מפרויקט (ראו engine/print.ts). */
const A4_WIDTH_IN = 8.268;

/** מעמידה מאגס `.editor-stack` ברוחב ידוע — jsdom אינו ממשיח layout. */
function installEditorStack(widthPx: number): () => void {
  const stack = document.createElement('main');
  stack.className = 'editor-stack';
  Object.defineProperty(stack, 'clientWidth', { value: widthPx });
  document.body.appendChild(stack);
  return () => stack.remove();
}

function mountWithPageWidth(widthIn?: number) {
  const options: SuperdocDoubleOptions = {
    ...(widthIn === undefined
      ? {}
      : { sections: { pageSize: { width: widthIn, height: 11.694 } } }),
  };
  return mountUi(ViewTab, { superdoc: createSuperdocDouble(options) });
}

describe('כפתורי לשונית „תצוגה”', () => {
  it('„100%” שולח את הפקודה עם 100 בדיוק', async () => {
    const harness = mountUi(ViewTab);
    await settle();

    await buttonByTip(harness.wrapper, HUNDRED_TITLE).trigger('click');
    await settle();

    expect(harness.adapter.payloads('zoom')).toEqual([100]);
    expect(harness.adapter.rejected).toEqual([]);
    expect(harness.failures()).toEqual([]);
  });

  it('„רוחב עמוד” מחשב את האחוז מרוחב המאגס ומידות הדף — ושולח אותו כ-payload', async () => {
    // 740px מול A4 (793.73px) → 93%. חלון צר: זו התאמה להקטנה.
    const removeStack = installEditorStack(740);
    try {
      const harness = mountWithPageWidth(A4_WIDTH_IN);
      await settle();

      await buttonByTip(harness.wrapper, FIT_TITLE).trigger('click');
      await settle();

      expect(harness.adapter.payloads('zoom')).toEqual([93]);
      expect(harness.adapter.rejected).toEqual([]);
      expect(harness.failures()).toEqual([]);
      expect(harness.adapter.calls.some((call) => call.id === 'zoom-fit-width')).toBe(false);
    } finally {
      removeStack();
    }
  });

  it('„רוחב עמוד” בחלון רחב מגדיל מעבר ל-100%', async () => {
    // לכפיל אין `getZoomState`, ואז נופלים לגבולות ברירת המחדל (10–500):
    // 1480px מול A4 → 186%, בתוך הטווח.
    const removeStack = installEditorStack(1480);
    try {
      const harness = mountWithPageWidth(A4_WIDTH_IN);
      await settle();

      await buttonByTip(harness.wrapper, FIT_TITLE).trigger('click');
      await settle();

      expect(harness.adapter.payloads('zoom')).toEqual([186]);
    } finally {
      removeStack();
    }
  });

  it('„רוחב עמוד” נצמד לתקרת ההיקף של Word (500%)', async () => {
    // 4200px מול A4 → 529% → התקרה. ה-max שהמנוע מדווח הוא גבול ה-fit-width
    // שלו ולא מגבלת זום ידני (setZoom אינו מצמצם), ולכן התקרה שלנו היא 500.
    const removeStack = installEditorStack(4200);
    try {
      const harness = mountWithPageWidth(A4_WIDTH_IN);
      await settle();

      await buttonByTip(harness.wrapper, FIT_TITLE).trigger('click');
      await settle();

      expect(harness.adapter.payloads('zoom')).toEqual([ZOOM_PERCENT_MAX]);
    } finally {
      removeStack();
    }
  });

  it('„רוחב עמוד” בלי מסמך פתוח מדווח ולא שולח פקודה', async () => {
    const harness = mountUi(ViewTab, { superdoc: null });
    await settle();

    await buttonByTip(harness.wrapper, FIT_TITLE).trigger('click');
    await settle();

    expect(harness.adapter.payloads('zoom')).toEqual([]);
    const failures = harness.failures();
    expect(failures).toHaveLength(1);
    if (!failures[0].outcome.ok) expect(failures[0].outcome.message).toContain('אין מסמך פתוח');
  });

  it('„רוחב עמוד” שמידות הדף שלו אינן קריאות מדווח, ולא מנחש אחוז', async () => {
    // 11906 הוא רוחב A4 ב-twips: יחידות שנשכחו בדרך, והקורא מסנן אותן במקום
    // לחשב מהן אחוזי הזוי (`isSaneInches` ב-engine/print.ts). המידה נמסרת כאן
    // מפורשות ואינה ברירת המחדל של הכפיל — ברירת המחדל היא A4 תקין, מפני
    // ש-`layOutTocRows` גוזר ממנה את רוחב אזור הטקסט.
    const removeStack = installEditorStack(740);
    try {
      const harness = mountWithPageWidth(11906);
      await settle();

      await buttonByTip(harness.wrapper, FIT_TITLE).trigger('click');
      await settle();

      expect(harness.adapter.payloads('zoom')).toEqual([]);
      expect(harness.failures()).toHaveLength(1);
      const [failure] = harness.failures();
      if (!failure.outcome.ok) expect(failure.outcome.reason).toBe('geometry-unavailable');
    } finally {
      removeStack();
    }
  });

  it('„סרגל” מריץ את פקודת ה-ruler', async () => {
    const harness = mountUi(ViewTab);
    await settle();

    await buttonByTip(harness.wrapper, RULER_TITLE).trigger('click');
    await settle();

    expect(harness.adapter.applied.filter((c) => c.id === 'ruler')).toHaveLength(1);
  });

  it('„סימני עיצוב” מריץ את פקודת formatting-marks', async () => {
    const harness = mountUi(ViewTab);
    await settle();

    await buttonByTip(harness.wrapper, MARKS_TITLE).trigger('click');
    await settle();

    expect(harness.adapter.applied.filter((c) => c.id === 'formatting-marks')).toHaveLength(1);
  });

  it('„מצב מיקוד” פולט toggle-focus-mode ולא פונה למנוע', async () => {
    const harness = mountUi(ViewTab);
    await settle();

    await buttonByTip(harness.wrapper, 'מצב קריאה ומיקוד').trigger('click');
    await settle();

    expect(harness.wrapper.emitted('toggle-focus-mode')).toHaveLength(1);
    expect(harness.adapter.calls).toEqual([]);
  });

  it('כשהמנוע אינו זמין — כל כפתורי הפקודות מנוטרלים, ולחיצה לא מגיעה לאדפטר', async () => {
    const harness = mountUi(ViewTab);
    await settle();

    harness.adapter.setState('zoom', { enabled: false });
    harness.adapter.setState('ruler', { enabled: false });
    harness.adapter.setState('formatting-marks', { enabled: false });
    await settle();

    expect(buttonByTip(harness.wrapper, HUNDRED_TITLE).attributes('disabled')).toBeDefined();
    expect(buttonByTip(harness.wrapper, FIT_TITLE).attributes('disabled')).toBeDefined();
    expect(buttonByTip(harness.wrapper, RULER_TITLE).attributes('disabled')).toBeDefined();
    expect(buttonByTip(harness.wrapper, MARKS_TITLE).attributes('disabled')).toBeDefined();

    await buttonByTip(harness.wrapper, HUNDRED_TITLE).trigger('click');
    await settle();
    expect(harness.adapter.calls).toEqual([]);
  });
});
/**
 * צבע הבד — המשטח שסביב הדף.
 *
 * הפקד הזה אינו פקודת מנוע: הוא כותב טוקן CSS על שורש המסמך ושומר את הבחירה
 * (composables/canvas-color.ts). לכן מה שנמדד כאן הוא שלוש הטענות שהמשתמש
 * רואה — הבד נצבע, „ברירת מחדל” מחזירה אותו לחום הבהיר של tokens.css, והפס
 * מתחת לאייקון מבטיח את מה שהלחיצה תחיל.
 *
 * הקבוצה בלשונית „תצוגה” ולא ב„בית”: זו העדפת תצוגה, לא תכונה של המסמך.
 */
describe('צבע רקע העורך', () => {
  /**
   * מה שפותח את הפלטה. הפקד כאן הוא `variant="large"` — הדפוס המפוצל הגדול
   * של „תאריך ושעה” — ולכן רצועת החץ היא המחלקה הגלובלית של הרצועה ולא
   * `.color-arrow-btn` המקומית. הסלקטור בקבוע אחד כדי שהחלפת הדפוס תיפול
   * כאן פעם אחת, ולא בארבע בדיקות נפרדות.
   */
  const TRIGGER = '.word-split--large .word-split__arrow';

  /** הצהרת הטוקן על שורש המסמך, כפי שהדפדפן היה קורא אותה. */
  function declared(): string {
    return document.documentElement.style.getPropertyValue(CANVAS_COLOR_VAR);
  }

  beforeEach(() => {
    saveCanvasColor.mockClear();
  });

  afterEach(() => {
    applyCanvasColor(null);
    setCanvasDefaultColor(null);
  });

  it('בחירת צבע צובעת את הבד ונזכרת', async () => {
    const harness = mountUi(ViewTab);
    await settle();

    await harness.wrapper.find(TRIGGER).trigger('click');
    await harness.wrapper.find('.standard-colors-row .color-swatch').trigger('click');
    await settle();

    // הצבע הראשון ב„צבעים רגילים” (ColorPickerPopover.vue).
    expect(declared()).toBe('#c00000');
    expect(saveCanvasColor).toHaveBeenCalledWith('#c00000');
  });

  it('„ברירת מחדל” מסירה את ההצהרה — הבד חוזר לערך שב-tokens.css', async () => {
    // היעדר ההצהרה, ולא כתיבה של החום הבהיר: כך ברירת המחדל נשארת במקום
    // אחד, ושינוי שלה מגיע גם למי שבחר בה פעם.
    const harness = mountUi(ViewTab);
    await settle();

    await harness.wrapper.find(TRIGGER).trigger('click');
    await harness.wrapper.find('.standard-colors-row .color-swatch').trigger('click');
    await settle();
    expect(declared()).not.toBe('');

    await harness.wrapper.find(TRIGGER).trigger('click');
    await harness.wrapper.find('.palette-clear-btn').trigger('click');
    await settle();

    expect(declared()).toBe('');
    expect(saveCanvasColor).toHaveBeenLastCalledWith(null);
  });

  it('הפריט המנקה אומר „ברירת מחדל” ולא „ללא צבע”', async () => {
    // „ללא צבע” הוא תיאור שגוי של מה שהלחיצה עושה: היא מחזירה את צבע ברירת
    // המחדל, ומשטח חסר צבע אינו קיים כאן. ראו `clearLabel`.
    const harness = mountUi(ViewTab);
    await settle();

    await harness.wrapper.find(TRIGGER).trigger('click');
    await settle();

    expect(harness.wrapper.find('.palette-clear-btn').text()).toBe('ברירת מחדל');
  });

  it('הפקד הוא הדפוס המפוצל הגדול — אותן מחלקות כמו „תאריך ושעה”', async () => {
    // הוא לבדו בקבוצה בת 70px, ופקד בן 22px נראה שם אבוד. המחלקות הן
    // הגלובליות של הרצועה ולא עותק מקומי שלהן: כך הגאומטריה של שני הדפוסים
    // אינה יכולה להיפרד. ראו `variant` ב-ColorPickerPopover.
    const harness = mountUi(ViewTab);
    await settle();

    const split = harness.wrapper.find('.word-split.word-split--large');
    expect(split.exists()).toBe(true);
    expect(split.find('.word-btn.btn-large').exists()).toBe(true);
    expect(split.find('.word-split__arrow').exists()).toBe(true);
    expect(split.find('.btn-label').text()).toBe('צבע רקע');
  });

  it('בלי העדפה הפס מראה את הבד כפי שהוא נראה — משטח הנושא המוכהה — ולא שחור', async () => {
    // ברירת המחדל של הבורר היא `#000000`, והפס הוא ההבטחה של הכפתור הראשי —
    // כלומר בלי הענף הזה הפקד היה מבטיח „לחיצה תצבע את הבד בשחור”.
    setCanvasDefaultColor('#ffffff');
    const harness = mountUi(ViewTab);
    await settle();

    // 255 × 0.93 — הלבן של הנושא מתחת לשכבת `--color-shell-dim`.
    expect(harness.wrapper.find('.color-indicator-bar').attributes('style')).toContain(
      'rgb(237, 237, 237)',
    );
  });
});
