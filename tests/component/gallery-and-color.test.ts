/**
 * הפקדים המורכבים ברצועה: גלריית הסגנונות, בורר הצבעים, כפתור התפריט ובורר
 * הטבלה.
 *
 * שניהם מקבלים את מה שהם מציגים מהמנוע: הגלריה את הקטלוג של המסמך הפתוח
 * (`STYLE_GALLERY`), והבורר את הצבע שהמנוע מדווח על הבחירה. הרשימה הקשיחה
 * שהייתה בגלריה לא ידעה מה יש במסמך, ו„ללא צבע” שלח מחרוזת ריקה — שהמנוע דוחה
 * סגור, כלומר כפתור שנראה כאילו הוא מנקה ואינו מנקה.
 *
 * מה שנמדד כאן הוא מה שדורש רינדור: הכרטיסים שנבנו מהקטלוג, הכרטיס המסומן
 * (כולל בחירה מעורבת, שבה **אין** כרטיס מסומן), כפתורי הגלילה שמופיעים רק
 * כשיש לאן לגלול, וה-payload שיוצא מהשבב.
 *
 * שני הפופאוברים האחרונים חולקים מכניקה אחת, וזו הסיבה שהם באותו קובץ:
 * `@pointerdown.prevent` על הפקד (בלעדיו הלחיצה גוזלת את המיקוד מהעורך והבחירה
 * במסמך אובדת), `.stop` על התוכן כדי שהמאזין הגלובלי לא יסגור אותו ברגע
 * שנוגעים בו, ו-Escape שסוגר ומחזיר מיקוד. מאז שהמיקום עבר לקוד
 * (composables/popover-position.ts) הם חולקים גם אותו — וזה נבדק בסוף הקובץ.
 */
import { describe, expect, it, vi } from 'vitest';
import ColorPickerPopover from '../../src/ui/ribbon/common/ColorPickerPopover.vue';
import RibbonMenuButton from '../../src/ui/ribbon/common/RibbonMenuButton.vue';
import TablePicker from '../../src/ui/ribbon/common/TablePicker.vue';
import StyleGallery from '../../src/ui/ribbon/common/StyleGallery.vue';
import { GALLERY_SCROLL_STEP_PX, type StyleGalleryState } from '../../src/engine/style-gallery';
import { POPOVER_GAP_PX } from '../../src/composables/popover-position';
import { autoUnmount, clickOutside, mountUi, settle, tipOf } from './harness';

autoUnmount();

/** מצב גלריה כפי שהמנוע מדווח אותו למסמך פתוח. */
function documentGallery(
  activeId: string | null,
  ids: readonly string[] = ['Normal', 'Heading1', 'MyStyle'],
): StyleGalleryState {
  return {
    items: ids.map((id) => ({
      id,
      label: id === 'MyStyle' ? 'סגנון של המסמך' : id,
      previewText: 'AaBbCc',
      previewStyle: {},
    })),
    activeId,
    fromDocument: true,
  };
}

/**
 * jsdom אינו מודד פריסה: `scrollWidth` ו-`clientWidth` הם 0, ולכן גלריה נראית
 * שם תמיד כמי שנכנסת כולה. המידות נמסרות במפורש כדי למדוד את ההחלטה עצמה.
 */
function fakeMetrics(
  element: HTMLElement,
  metrics: { scrollLeft: number; scrollWidth: number; clientWidth: number },
): void {
  for (const [key, value] of Object.entries(metrics)) {
    Object.defineProperty(element, key, { value, configurable: true });
  }
}

describe('StyleGallery', () => {
  it('בלי גלריה מהמנוע — הרשימה הבנויה, בתוויות עבריות', async () => {
    // רשת הביטחון: רצועה שעולה לפני שנפתח מסמך צריכה גלריה עובדת.
    const harness = mountUi(StyleGallery);
    await settle();

    expect(harness.wrapper.findAll('.style-card').map((card) => tipOf(card).title)).toEqual([
      'רגיל',
      'ללא מרווח',
      'כותרת 1',
      'כותרת 2',
      // „כותרת 3” מסונתזת: ראו `withHeadingLadder` ב-engine/style-gallery.ts.
      'כותרת 3',
      'כותרת משנה',
      'ציטוט',
    ]);
  });

  it('הכרטיסים מגיעים מהקטלוג של המסמך, כולל סגנון שאינו מובנה', async () => {
    const harness = mountUi(StyleGallery, { styleGallery: documentGallery('Heading1') });
    await settle();

    const cards = harness.wrapper.findAll('.style-card');
    expect(cards.map((card) => tipOf(card).title)).toEqual([
      'Normal',
      'Heading1',
      'סגנון של המסמך',
    ]);
    expect(cards.map((card) => card.attributes('aria-pressed'))).toEqual([
      'false',
      'true',
      'false',
    ]);
  });

  it('בחירה מעורבת — אין כרטיס מסומן', async () => {
    // המנוע מחזיר `null`, ו-Word מציג גלריה בלי בחירה. הסגנון של הפסקה
    // הראשונה כאילו הוא של כולן היה שקר.
    const harness = mountUi(StyleGallery, { styleGallery: documentGallery(null) });
    await settle();

    expect(
      harness.wrapper.findAll('.style-card').map((card) => card.attributes('aria-pressed')),
    ).toEqual(['false', 'false', 'false']);
  });

  it('בלי גלריה מהמסמך, ה-prop של הפקודה קובע את הכרטיס הפעיל', async () => {
    const harness = mountUi(StyleGallery, { props: { currentStyle: 'Heading2' } });
    await settle();

    const pressed = harness.wrapper
      .findAll('.style-card')
      .filter((card) => card.attributes('aria-pressed') === 'true');
    expect(pressed).toHaveLength(1);
    expect(tipOf(pressed[0]).title).toBe('כותרת 2');
  });

  it('לחיצה פולטת את מזהה הסגנון — לא את התווית', async () => {
    const harness = mountUi(StyleGallery, { styleGallery: documentGallery('Normal') });
    await settle();

    await harness.wrapper.findAll('.style-card')[2].trigger('click');
    expect(harness.wrapper.emitted('select-style')).toEqual([['MyStyle']]);
  });

  it('בלי בחירה במסמך כל הכרטיסים מנוטרלים', async () => {
    // `linked-style` נכשל בלי בחירה, וגלריה שנראית פעילה מזמינה לחיצה סרק.
    const harness = mountUi(StyleGallery, { props: { disabled: true } });
    await settle();

    for (const card of harness.wrapper.findAll('.style-card')) {
      expect(card.attributes('disabled')).toBeDefined();
    }
  });

  it('כפתורי הגלילה מופיעים רק כשיש לאן לגלול', async () => {
    const harness = mountUi(StyleGallery);
    await settle();

    // גלריה שנכנסת כולה: אין כפתורים בכלל.
    expect(harness.wrapper.findAll('.nav-btn')).toHaveLength(0);

    const scroller = harness.wrapper.find('.style-cards-scroll');
    fakeMetrics(scroller.element as HTMLElement, {
      scrollLeft: 0,
      scrollWidth: 600,
      clientWidth: 300,
    });
    await scroller.trigger('scroll');
    await settle();

    // בקצה ההתחלה יש רק „הבאים”.
    expect(harness.wrapper.findAll('.nav-btn').map((button) => button.attributes('aria-label'))).toEqual([
      'הסגנונות הבאים',
    ]);

    fakeMetrics(scroller.element as HTMLElement, {
      scrollLeft: -150,
      scrollWidth: 600,
      clientWidth: 300,
    });
    await scroller.trigger('scroll');
    await settle();

    expect(harness.wrapper.findAll('.nav-btn').map((button) => button.attributes('aria-label'))).toEqual([
      'הסגנונות הקודמים',
      'הסגנונות הבאים',
    ]);
  });

  it('הגלילה ב-RTL זזה לכיוון החזותי הנכון', async () => {
    const harness = mountUi(StyleGallery);
    await settle();

    const scroller = harness.wrapper.find('.style-cards-scroll');
    const element = scroller.element as HTMLElement;
    // jsdom אינו מממש `scrollBy`, ולכן הוא מוזרק — מה שנמדד הוא הארגומנט.
    const scrollBy = vi.fn();
    Object.defineProperty(element, 'scrollBy', { value: scrollBy, configurable: true });
    fakeMetrics(element, { scrollLeft: -150, scrollWidth: 600, clientWidth: 300 });
    await scroller.trigger('scroll');
    await settle();

    await harness.wrapper.find('.nav-btn[aria-label="הסגנונות הבאים"]').trigger('click');
    await harness.wrapper.find('.nav-btn[aria-label="הסגנונות הקודמים"]').trigger('click');

    // ב-RTL הגלילה קדימה היא לכיוון השלילי — ההיפוך שנשבר בלי שרואים.
    expect(scrollBy.mock.calls.map((call) => call[0].left)).toEqual([
      -GALLERY_SCROLL_STEP_PX,
      GALLERY_SCROLL_STEP_PX,
    ]);
  });
});

describe('ColorPickerPopover', () => {
  const open = async (harness: ReturnType<typeof mountUi>) => {
    await harness.wrapper.find('.color-arrow-btn').trigger('click');
    await settle();
  };

  /** הפס שמתחת לאייקון — הצבע שהלחיצה על הכפתור הראשי תחיל. */
  const bar = (harness: ReturnType<typeof mountUi>) =>
    harness.wrapper.find('.color-indicator-bar');

  /** jsdom מנרמל צבע ל-`rgb(...)`, ולכן גם הערך המצופה עובר את אותו נרמול. */
  const asCss = (color: string): string => {
    const probe = document.createElement('div');
    probe.style.backgroundColor = color;
    return probe.style.backgroundColor;
  };

  it('שבב מהפלטה פולט את הצבע, ובעקבותיו הפופאובר נסגר', async () => {
    const harness = mountUi(ColorPickerPopover, {
      props: { icon: 'fontColor', title: 'צבע גופן' },
    });
    await open(harness);

    expect(harness.wrapper.find('.color-palette-popover').exists()).toBe(true);
    // הקוד ירד ל-`data-tip-desc`, והכותרת היא השם — זה מה שקורא מסך מכריז.
    const swatch = harness.wrapper.find('.color-swatch[data-tip-desc="#c00000"]');
    expect(swatch.attributes('data-tip-title')).toBe('אדום כהה');
    expect(swatch.attributes('aria-label')).toBe('אדום כהה');
    await swatch.trigger('click');
    await settle();

    expect(harness.wrapper.emitted('change')).toEqual([['#c00000']]);
    expect(harness.wrapper.emitted('update:modelValue')).toEqual([['#c00000']]);
    expect(harness.wrapper.find('.color-palette-popover').exists()).toBe(false);
    // issue #14 (ד): אחרי הבחירה הסמן חוזר לטקסט, ולא נשאר „מחוץ למסמך".
    expect(harness.superdoc.ops()).toContain('focus');
  });

  it('„ללא צבע” פולט `null` — ולא מחרוזת ריקה', async () => {
    // זה החוזה של המנוע: `value === null` הוא מסלול הניקוי, ומחרוזת ריקה
    // נדחית שם במפורש. ה-`modelValue` נשאר מחרוזת כי הוא מזין CSS.
    const harness = mountUi(ColorPickerPopover, {
      props: { icon: 'highlight', title: 'צבע סימון', modelValue: '#FFFF00' },
    });
    await open(harness);

    await harness.wrapper.find('.palette-clear-btn').trigger('click');
    await settle();

    expect(harness.wrapper.emitted('change')).toEqual([[null]]);
    expect(harness.wrapper.emitted('update:modelValue')).toEqual([['']]);
  });

  it('`allowClear: false` אינו מציע ניקוי', async () => {
    const harness = mountUi(ColorPickerPopover, {
      props: { icon: 'fontColor', title: 'צבע גופן', allowClear: false },
    });
    await open(harness);

    expect(harness.wrapper.find('.palette-clear-btn').exists()).toBe(false);
  });

  /*
   * הפס והלחיצה הם ערך אחד — וזה **לא** צבע המסמך.
   *
   * הפס היה קשור ל-`modelValue`, כלומר לצבע שהמנוע מדווח על הבחירה, ולכן
   * הכפתור הראשי החיל על הטקסט את הצבע שכבר יש לו: בסימון זה לא נראה בכלל
   * (ברירת המחדל צהובה גם כך), ובצבע הגופן הצבע שנבחר נעלם מהפס ברגע שהסמן
   * עבר לטקסט שחור. הפס הוא ההבטחה של הכפתור, ולכן הוא מראה את מה שהלחיצה
   * תעשה — הבחירה האחרונה בפקד, שנדבקת אליו כמו ב-Word.
   */
  it('הפס והלחיצה הם הצבע שנבחר, ולא צבע הטקסט שהסמן עומד עליו', async () => {
    const harness = mountUi(ColorPickerPopover, {
      props: {
        icon: 'highlight',
        title: 'צבע סימון',
        defaultColor: '#FFFF00',
        // מה שהמנוע מדווח על הבחירה — טקסט שכבר מסומן בכחול.
        modelValue: '#0055FF',
      },
    });
    await open(harness);
    await harness.wrapper.find('.color-swatch[data-tip-desc="#ffff00"]').trigger('click');
    await settle();

    expect((bar(harness).element as HTMLElement).style.backgroundColor).toBe(asCss('#ffff00'));

    await harness.wrapper.find('.color-main-btn').trigger('click');
    await settle();

    // הצבע נדבק: הלחיצה השנייה שולחת אותו שוב, ולא את הכחול של המסמך.
    expect(harness.wrapper.emitted('change')).toEqual([['#ffff00'], ['#ffff00']]);
    expect(harness.wrapper.find('.color-palette-popover').exists()).toBe(false);
  });

  it('טרם נבחר צבע — הפס והלחיצה הם ברירת המחדל', async () => {
    const harness = mountUi(ColorPickerPopover, {
      props: { icon: 'highlight', title: 'צבע סימון', defaultColor: '#FFFF00' },
    });
    await harness.wrapper.find('.color-main-btn').trigger('click');

    expect(harness.wrapper.emitted('change')).toEqual([['#FFFF00']]);
    expect((bar(harness).element as HTMLElement).style.backgroundColor).toBe(asCss('#FFFF00'));
  });

  it('`applyOnClick: false` — החצי הראשי פותח את הפלטה ואינו מחיל דבר', async () => {
    // הפיצול נכון לצבע של בחירה, ולא לצבע של משטח יחיד: בבד (רקע העורך)
    // הפס מראה את הצבע שהמשטח כבר צבוע בו, ו„החל את הצבע שמוצג” הוא לחיצה
    // שאינה עושה כלום. נמדד בשער „אין כפתור מת” ב-ribbon-tabs.
    const harness = mountUi(ColorPickerPopover, {
      props: {
        icon: 'shading',
        title: 'צבע רקע העורך',
        defaultColor: '#EDEBE9',
        applyOnClick: false,
      },
    });

    await harness.wrapper.find('.color-main-btn').trigger('click');

    expect(harness.wrapper.emitted('change')).toBeUndefined();
    expect(harness.wrapper.find('.color-palette-popover').exists()).toBe(true);

    // ולחיצה שנייה סוגרת — אותו מתג כמו החץ, ולא פתיחה בלבד.
    await harness.wrapper.find('.color-main-btn').trigger('click');
    expect(harness.wrapper.find('.color-palette-popover').exists()).toBe(false);
  });

  it('בורר משטח עוקב אחרי המודל וחוזר להציג את ברירת המחדל באיפוס', async () => {
    const harness = mountUi(ColorPickerPopover, {
      props: {
        icon: 'shading', title: 'צבע רקע העורך', defaultColor: '#eeece1',
        modelValue: '#ffffff', followModelValue: true, applyOnClick: false,
      },
    });
    expect((bar(harness).element as HTMLElement).style.backgroundColor).toBe(asCss('#ffffff'));

    await open(harness);
    await harness.wrapper.find('.palette-clear-btn').trigger('click');
    expect(harness.wrapper.emitted('change')).toEqual([[null]]);
    await harness.wrapper.setProps({ modelValue: '' });
    expect(bar(harness).classes()).not.toContain('is-none');
    expect((bar(harness).element as HTMLElement).style.backgroundColor).toBe(asCss('#eeece1'));

    // גם שינוי חיצוני של ההעדפה חייב לעדכן את הפס בלי להרכיב את הבורר מחדש.
    await harness.wrapper.setProps({ modelValue: '#123456' });
    expect((bar(harness).element as HTMLElement).style.backgroundColor).toBe(asCss('#123456'));
  });

  it('„ללא צבע” נדבק אף הוא: הלחיצה הבאה מנקה ואינה מחילה את ברירת המחדל', async () => {
    const harness = mountUi(ColorPickerPopover, {
      props: { icon: 'highlight', title: 'צבע סימון', defaultColor: '#FFFF00' },
    });
    await open(harness);
    await harness.wrapper.find('.palette-clear-btn').trigger('click');
    await settle();

    // פס ריק — קו היקפי בלבד; „ללא צבע” הוא היעדר ולא צהוב.
    expect(bar(harness).classes()).toContain('is-none');

    await harness.wrapper.find('.color-main-btn').trigger('click');
    await settle();

    expect(harness.wrapper.emitted('change')).toEqual([[null], [null]]);
  });

  it('הצבע שיוחל נאמר גם במילים — הפס אינו נראה לקורא מסך', async () => {
    const harness = mountUi(ColorPickerPopover, {
      props: { icon: 'fontColor', title: 'צבע גופן', defaultColor: '#c00000' },
    });
    const main = () => harness.wrapper.find('.color-main-btn');

    expect(tipOf(main()).title).toBe('צבע גופן');
    expect(tipOf(main()).description).toBe('אדום כהה');
    expect(main().attributes('aria-label')).toBe('צבע גופן, אדום כהה');

    await open(harness);
    await harness.wrapper.find('.color-swatch[data-tip-desc="#4f81bd"]').trigger('click');
    await settle();

    expect(tipOf(main()).description).toBe('כחול');
    expect(main().attributes('aria-label')).toBe('צבע גופן, כחול');
  });

  it('לחיצה מחוץ לפקד סוגרת את הפלטה', async () => {
    const harness = mountUi(ColorPickerPopover, {
      props: { icon: 'fontColor', title: 'צבע גופן' },
    });
    await open(harness);
    expect(harness.wrapper.find('.color-palette-popover').exists()).toBe(true);

    clickOutside();
    await settle();

    expect(harness.wrapper.find('.color-palette-popover').exists()).toBe(false);
  });

  it('פקד מנוטרל אינו נפתח ואינו מחיל דבר', async () => {
    const harness = mountUi(ColorPickerPopover, {
      props: { icon: 'fontColor', title: 'צבע גופן', disabled: true, modelValue: '#0055FF' },
    });

    for (const selector of ['.color-main-btn', '.color-arrow-btn']) {
      const button = harness.wrapper.find(selector);
      expect(button.attributes('disabled'), selector).toBeDefined();
      await button.trigger('click');
    }
    await settle();

    expect(harness.wrapper.emitted('change')).toBeUndefined();
    expect(harness.wrapper.find('.color-palette-popover').exists()).toBe(false);
  });
});

describe('RibbonMenuButton', () => {
  const MENU_PROPS = {
    icon: 'margins',
    label: 'שוליים',
    tooltip: 'הגדרת שולי הדף',
    items: [
      { id: 'normal', label: 'רגיל', hint: '2.54 ס"מ מכל צד' },
      { id: 'narrow', label: 'צר', hint: '1.27 ס"מ מכל צד' },
    ],
  };

  it('נפתח, מכריז על עצמו כתפריט, ומציג את הפריטים עם ההסבר', async () => {
    const harness = mountUi(RibbonMenuButton, { props: MENU_PROPS });
    const button = harness.wrapper.find('button');

    expect(button.attributes('aria-haspopup')).toBe('menu');
    expect(button.attributes('aria-expanded')).toBe('false');

    await button.trigger('click');
    await settle();

    expect(harness.wrapper.find('button').attributes('aria-expanded')).toBe('true');
    const items = harness.wrapper.findAll('[role="menuitem"]');
    expect(items).toHaveLength(2);
    expect(items[0].text()).toContain('רגיל');
    expect(items[0].text()).toContain('2.54');
    expect(harness.wrapper.find('[role="menu"]').attributes('aria-label')).toBe('שוליים');
  });

  it('בחירה פולטת את המזהה וסוגרת', async () => {
    const harness = mountUi(RibbonMenuButton, { props: MENU_PROPS });
    await harness.wrapper.find('button').trigger('click');
    await settle();

    await harness.wrapper.findAll('[role="menuitem"]')[1].trigger('click');
    await settle();

    expect(harness.wrapper.emitted('select')).toEqual([['narrow']]);
    expect(harness.wrapper.find('[role="menu"]').exists()).toBe(false);
  });

  it('Escape סוגר ומחזיר את המיקוד לכפתור', async () => {
    const harness = mountUi(RibbonMenuButton, { props: MENU_PROPS });
    await harness.wrapper.find('button').trigger('click');
    await settle();

    await harness.wrapper.find('.ribbon-menu').trigger('keydown.escape');
    await settle();

    expect(harness.wrapper.find('[role="menu"]').exists()).toBe(false);
    expect(document.activeElement).toBe(harness.wrapper.find('button').element);
  });

  it('לחיצה בחוץ סוגרת, ולחיצה בתוך התפריט אינה', async () => {
    const harness = mountUi(RibbonMenuButton, { props: MENU_PROPS });
    await harness.wrapper.find('button').trigger('click');
    await settle();

    // `.stop` על הפופאובר הוא מה שמונע מהמאזין הגלובלי לסגור אותו מיד.
    await harness.wrapper.find('[role="menu"]').trigger('pointerdown');
    await settle();
    expect(harness.wrapper.find('[role="menu"]').exists()).toBe(true);

    clickOutside();
    await settle();
    expect(harness.wrapper.find('[role="menu"]').exists()).toBe(false);
  });

  it('פקד מנוטרל אינו נפתח', async () => {
    const harness = mountUi(RibbonMenuButton, {
      props: { ...MENU_PROPS, disabled: true },
    });

    const button = harness.wrapper.find('button');
    expect(button.attributes('disabled')).toBeDefined();
    await button.trigger('click');
    await settle();

    expect(harness.wrapper.find('[role="menu"]').exists()).toBe(false);
  });
});

describe('TablePicker', () => {
  /** פותחת את הבורר ומחזירה את הגריד. */
  async function openGrid(harness: ReturnType<typeof mountUi>) {
    await harness.wrapper.find('button').trigger('click');
    await settle();
    return harness.wrapper.find('[role="grid"]');
  }

  it('הגריד הוא פקד עם סמנטיקה, ולא מאה div ריקים', async () => {
    // זה מה שהיה: `<div class="grid-cell" @click>` ×100, בלי role, בלי שם
    // ובלי דרך להגיע אליו במקלדת.
    const harness = mountUi(TablePicker);
    const grid = await openGrid(harness);

    expect(grid.attributes('aria-label')).toBe('בחירת מידות הטבלה');
    expect(harness.wrapper.findAll('[role="row"]')).toHaveLength(10);

    const cells = harness.wrapper.findAll('[role="gridcell"]');
    expect(cells).toHaveLength(100);
    expect(cells[0].attributes('aria-label')).toBe('עמודה אחת על שורה אחת');
    expect(cells[12].attributes('aria-label')).toBe('3 עמודות על 2 שורות');

    const header = harness.wrapper.find('.table-picker-header');
    expect(header.attributes('role')).toBe('status');
    expect(header.attributes('aria-live')).toBe('polite');
  });

  it('נקודת Tab אחת לכל הגריד, ולא אחת לכל תא', async () => {
    const harness = mountUi(TablePicker);
    const grid = await openGrid(harness);

    expect(grid.attributes('tabindex')).toBe('0');
    for (const cell of harness.wrapper.findAll('[role="gridcell"]')) {
      expect(cell.attributes('tabindex')).toBeUndefined();
    }
  });

  it('כניסה במקלדת מעמידה את הסמן על 1×1, כדי שיהיה מה לאשר', async () => {
    const harness = mountUi(TablePicker);
    const grid = await openGrid(harness);
    expect(harness.wrapper.find('.table-picker-header').text()).toBe('הוסף טבלה');

    (grid.element as HTMLElement).focus();
    await settle();

    expect(grid.attributes('aria-activedescendant')).toBe('table-picker-cell-1-1');
    expect(harness.wrapper.find('.table-picker-header').text()).toBe('טבלה 1 × 1');
    expect(
      harness.wrapper.findAll('[aria-selected="true"]'),
      'בדיוק תא אחת מסומן',
    ).toHaveLength(1);
  });

  it('החצים משנים מידות, ו-ArrowLeft מוסיף עמודה — הכיוון החזותי ב-RTL', async () => {
    const harness = mountUi(TablePicker);
    const grid = await openGrid(harness);
    (grid.element as HTMLElement).focus();
    await settle();

    await grid.trigger('keydown', { key: 'ArrowLeft' });
    await grid.trigger('keydown', { key: 'ArrowLeft' });
    await grid.trigger('keydown', { key: 'ArrowDown' });

    expect(harness.wrapper.find('.table-picker-header').text()).toBe('טבלה 3 × 2');
    expect(grid.attributes('aria-activedescendant')).toBe('table-picker-cell-2-3');
    expect(harness.wrapper.findAll('[aria-selected="true"]')).toHaveLength(6);

    // ArrowRight חוזר, כמו בסרגל הלשוניות.
    await grid.trigger('keydown', { key: 'ArrowRight' });
    expect(harness.wrapper.find('.table-picker-header').text()).toBe('טבלה 2 × 2');
  });

  it('הניווט אינו יוצא מהגריד לשום כיוון', async () => {
    const harness = mountUi(TablePicker);
    const grid = await openGrid(harness);
    (grid.element as HTMLElement).focus();
    await settle();

    for (let step = 0; step < 4; step += 1) {
      await grid.trigger('keydown', { key: 'ArrowUp' });
      await grid.trigger('keydown', { key: 'ArrowRight' });
    }
    expect(harness.wrapper.find('.table-picker-header').text()).toBe('טבלה 1 × 1');

    for (let step = 0; step < 14; step += 1) {
      await grid.trigger('keydown', { key: 'ArrowDown' });
      await grid.trigger('keydown', { key: 'ArrowLeft' });
    }
    expect(harness.wrapper.find('.table-picker-header').text()).toBe('טבלה 10 × 10');
  });

  it('Enter ורווח מאשרים את המידות שנבחרו במקלדת', async () => {
    for (const key of ['Enter', ' ']) {
      const harness = mountUi(TablePicker);
      const grid = await openGrid(harness);
      (grid.element as HTMLElement).focus();
      await settle();

      await grid.trigger('keydown', { key: 'ArrowLeft' });
      await grid.trigger('keydown', { key: 'ArrowDown' });
      await grid.trigger('keydown', { key });
      await settle();

      expect(harness.wrapper.emitted('select'), key).toEqual([[{ rows: 2, cols: 2 }]]);
      expect(harness.wrapper.find('[role="grid"]').exists(), key).toBe(false);
      harness.wrapper.unmount();
    }
  });

  it('Escape סוגר ומחזיר את המיקוד לכפתור', async () => {
    const harness = mountUi(TablePicker);
    const grid = await openGrid(harness);
    (grid.element as HTMLElement).focus();
    await settle();

    await grid.trigger('keydown', { key: 'Escape' });
    await settle();

    expect(harness.wrapper.find('[role="grid"]').exists()).toBe(false);
    expect(harness.wrapper.emitted('select')).toBeUndefined();
    expect(document.activeElement).toBe(harness.wrapper.find('button').element);
  });

  it('עכבר שחלף על הגריד אינו מוחק את בחירת המקלדת', async () => {
    const harness = mountUi(TablePicker);
    const grid = await openGrid(harness);
    (grid.element as HTMLElement).focus();
    await settle();

    await grid.trigger('keydown', { key: 'ArrowDown' });
    await grid.trigger('mouseleave');
    await settle();

    expect(harness.wrapper.find('.table-picker-header').text()).toBe('טבלה 1 × 2');
  });

  it('הכותרת מדווחת את המידות שמעל הסמן, ובחירה פולטת אותן', async () => {
    const harness = mountUi(TablePicker);
    await harness.wrapper.find('button').trigger('click');
    await settle();

    expect(harness.wrapper.find('.table-picker-header').text()).toBe('הוסף טבלה');

    // התא ה-13 בגריד 10×10 הוא שורה 2 עמודה 3.
    const cells = harness.wrapper.findAll('.grid-cell');
    await cells[12].trigger('mouseenter');
    expect(harness.wrapper.find('.table-picker-header').text()).toBe('טבלה 3 × 2');

    await cells[12].trigger('click');
    await settle();

    expect(harness.wrapper.emitted('select')).toEqual([[{ rows: 2, cols: 3 }]]);
    expect(harness.wrapper.find('.table-picker-popover').exists()).toBe(false);
  });

  it('לחיצה בחוץ סוגרת בלי לבחור', async () => {
    const harness = mountUi(TablePicker);
    await harness.wrapper.find('button').trigger('click');
    await settle();

    clickOutside();
    await settle();

    expect(harness.wrapper.find('.table-picker-popover').exists()).toBe(false);
    expect(harness.wrapper.emitted('select')).toBeUndefined();
  });
});

/* ------------------------------------------------------------------ */
/* מיקום הפופאוברים                                                   */
/* ------------------------------------------------------------------ */

/**
 * שלושת הפופאוברים היו `position: absolute; top: 100%` בתוך
 * `.word-ribbon-body`, שמוגדר `overflow-x: auto; overflow-y: hidden` — כלומר
 * ההורה חותך אנכית בגובה הרצועה, וכל מה שמתחתיו לא נראה. מהפלטה, שגובהה
 * ~150px, הוצגה בפועל רק שורת „ללא צבע”.
 *
 * החשבון עצמו נבדק ב-tests/unit/popover-position.test.ts. מה שנמדד כאן הוא
 * **החיווט**: שהמלבן של הכפתור באמת מגיע לפופאובר כ-`style`. חיווט שנשבר
 * (`ref` שלא נקשר, `:style` שנשמט) אינו נראה בבדיקה הטהורה בכלל.
 */
describe('מיקום הפופאוברים ברצועה', () => {
  const POPOVERS = [
    {
      name: 'פלטת הצבעים',
      component: ColorPickerPopover,
      props: { icon: 'fontColor', title: 'צבע גופן' },
      opener: '.color-arrow-btn',
      container: '.color-picker-container',
      popover: '.color-palette-popover',
    },
    {
      name: 'תפריט הכפתור',
      component: RibbonMenuButton,
      props: {
        icon: 'margins',
        label: 'שוליים',
        items: [{ id: 'normal', label: 'רגיל' }],
      },
      opener: 'button',
      container: '.ribbon-menu',
      popover: '.ribbon-menu__popover',
    },
    {
      name: 'בורר הטבלה',
      component: TablePicker,
      props: {},
      opener: 'button',
      container: '.table-picker-container',
      popover: '.table-picker-popover',
    },
  ] as const;

  /** כפתור בשליש התחתון של חלון נמוך — בדיוק המצב שבו הפופאובר נחתך. */
  const ANCHOR = { top: 400, bottom: 424, left: 900, right: 960 };
  const SIZE = { width: 200, height: 150 };
  const VIEWPORT = { width: 1000, height: 500 };

  /**
   * jsdom אינו מודד פריסה: כל `getBoundingClientRect` מחזיר אפסים, ולכן שם כל
   * פופאובר „נכנס למטה” ואף החלטת מיקום אינה נצפית. הזרקת מלבנים לפי סלקטור היא
   * מה שמעמיד את המצב האמיתי.
   */
  function fakeLayout(container: string, popover: string): () => void {
    const rects: Record<string, DOMRect> = {
      [container]: { ...ANCHOR, width: 60, height: 24, x: ANCHOR.left, y: ANCHOR.top } as DOMRect,
      [popover]: { top: 0, bottom: 0, left: 0, right: 0, x: 0, y: 0, ...SIZE } as DOMRect,
    };
    const originalRect = Element.prototype.getBoundingClientRect;
    Element.prototype.getBoundingClientRect = function (this: Element): DOMRect {
      for (const [selector, rect] of Object.entries(rects)) {
        if (this.matches(selector)) return rect;
      }
      return originalRect.call(this);
    };

    const window_ = { width: window.innerWidth, height: window.innerHeight };
    const resize = (width: number, height: number): void => {
      Object.defineProperty(window, 'innerWidth', { value: width, configurable: true });
      Object.defineProperty(window, 'innerHeight', { value: height, configurable: true });
    };
    resize(VIEWPORT.width, VIEWPORT.height);

    // האפליקציה רצה ב-`<html dir="rtl">` (index.html), ובלי זה ההצמדה כאן
    // הייתה נמדדת לקצה השמאלי — כלומר בדיקה שמאשרת את הכיוון ההפוך.
    const dir = document.documentElement.getAttribute('dir');
    document.documentElement.setAttribute('dir', 'rtl');

    return () => {
      Element.prototype.getBoundingClientRect = originalRect;
      resize(window_.width, window_.height);
      if (dir === null) document.documentElement.removeAttribute('dir');
      else document.documentElement.setAttribute('dir', dir);
    };
  }

  for (const popover of POPOVERS) {
    it(`${popover.name}: נפתח בקואורדינטות חלון, ולא בתוך המכל שחותך`, async () => {
      const restore = fakeLayout(popover.container, popover.popover);
      try {
        const harness = mountUi(popover.component, { props: popover.props });
        await harness.wrapper.find(popover.opener).trigger('click');
        await settle();

        const element = harness.wrapper.find(popover.popover).element as HTMLElement;
        expect(element.style.position).toBe('fixed');
        // אין מקום מתחת לכפתור בחלון בגובה 500, ולכן הפופאובר מתהפך למעלה
        // ונוגע בקצה העליון של הכפתור.
        expect(element.style.top).toBe(`${ANCHOR.top - POPOVER_GAP_PX - SIZE.height}px`);
        // RTL: הקצה הימני של הפופאובר מיושר לקצה הימני של הכפתור.
        expect(element.style.left).toBe(`${ANCHOR.right - SIZE.width}px`);
      } finally {
        restore();
      }
    });
  }
});
