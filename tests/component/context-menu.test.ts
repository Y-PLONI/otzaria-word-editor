/**
 * תפריט ההקשר, מורכב.
 *
 * מה שנמדד כאן הוא מה שסריקת מקור אינה יכולה לתפוס: שהלחיצה **מגיעה למנוע**
 * דרך אותו אדפטר של הרצועה, שפריט מנוטרל אינו רץ, ושהחצים מזיזים מיקוד. הכפיל
 * מריץ את ה-payload דרך הוולידטורים האמיתיים של superdoc, ולכן פקודה שהמנוע
 * היה דוחה נופלת כאן ולא אצל המשתמש.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { nextTick } from 'vue';
import ContextMenu from '../../src/ui/menu/ContextMenu.vue';
import {
  contextMenuModel,
  type ContextMenuSection,
  type ContextMenuSnapshot,
} from '../../src/ui/menu/context-menu-model';
import HomeTab from '../../src/ui/ribbon/tabs/HomeTab.vue';
import { createFontMemory } from '../../src/composables/use-font-controls';
import RibbonCombo from '../../src/ui/ribbon/common/RibbonCombo.vue';
import type { PickerOption } from '../../src/composables/picker-value';
import * as fontAvailability from '../../src/engine/docx-fonts';
import {
  autoUnmount,
  createCommandDouble,
  createSuperdocDouble,
  mountUi,
  pickerValue,
  setPicker,
  settle,
  type Harness,
} from './harness';

autoUnmount();

function sections(over: Partial<ContextMenuSnapshot> = {}): readonly ContextMenuSection[] {
  return contextMenuModel({
    hasDocument: true,
    hasRange: true,
    storyType: 'body',
    misspelledWord: null,
    can: () => true,
    ...over,
  });
}

function open(
  list: readonly ContextMenuSection[] = sections(),
  options: Parameters<typeof mountUi>[1] = {},
): Harness {
  return mountUi(ContextMenu, {
    ...options,
    props: { open: true, point: { x: 400, y: 300 }, sections: list, ...(options.props ?? {}) },
  });
}

/**
 * לפי `data-entry-id` ולא לפי אינדקס. הגרסה הראשונה כאן מיפתה מזהה למקום
 * ברשימה של הדגם **הראשוני** — כלומר הייתה שקטה ושגויה ברגע שהקשר מסתיר פריט
 * (בכותרת עליונה נופלים „הערת שוליים” ו„ציטוט”, וכל האינדקסים שאחריהם זזים).
 */
function buttonById(harness: Harness, id: string) {
  const button = harness.wrapper.find(`[data-entry-id="${id}"]`);
  if (!button.exists()) throw new Error(`אין פריט בתפריט עם המזהה "${id}"`);
  return button;
}

/** תיבת הבורר שבכרטיס, לפי הטולטיפ שלה. */
function boxInCard(harness: Harness, tip: string) {
  const box = harness.wrapper.find(`input[role="combobox"][data-tip-title="${tip}"]`);
  if (!box.exists()) throw new Error(`אין תיבה בכרטיס עם הטולטיפ „${tip}”`);
  return box;
}

/** האפשרויות שהתיבה בכרטיס קיבלה — הצורה שלהן, ולא רק מה שנצבע ממנה. */
function optionsOf(harness: Harness, tip: string): readonly PickerOption[] {
  const combo = harness.wrapper
    .findAllComponents(RibbonCombo)
    .find((instance) => instance.props('title') === tip);
  if (!combo) throw new Error(`אין בורר בכרטיס עם הכותרת „${tip}”`);
  return combo.props('options') as readonly PickerOption[];
}

/** הערכים שברשימה הנפתחת, בסדר שבו הם מצוירים. */
function rowValues(harness: Harness): (string | undefined)[] {
  return harness.wrapper.findAll('[role="option"]').map((row) => row.attributes('data-value'));
}

describe('ContextMenu', () => {
  let harness: Harness;

  beforeEach(() => {
    harness = open();
  });

  it('מצייר את המקטעים כתפריט נגיש', () => {
    expect(harness.wrapper.attributes('role')).toBe('menu');
    // שישה: לוח, גופן, עיצוב, הוספה, אוצריא, עריכה.
    expect(harness.wrapper.findAll('[role="group"]')).toHaveLength(6);
    expect(harness.wrapper.findAll('[role="separator"]')).toHaveLength(5);
  });

  it('מתג מדווח menuitemcheckbox, ופעולה מדווחת menuitem', async () => {
    const bold = buttonById(harness, 'bold');
    const link = buttonById(harness, 'link');

    expect(bold.attributes('role')).toBe('menuitemcheckbox');
    expect(bold.attributes('aria-checked')).toBeDefined();
    expect(link.attributes('role')).toBe('menuitem');
    expect(link.attributes('aria-checked')).toBeUndefined();
  });

  it('לחיצה על אייקון של פקודה מגיעה למנוע', async () => {
    await buttonById(harness, 'bold').trigger('click');

    expect(harness.adapter.applied.map((call) => call.id)).toContain('bold');
    expect(harness.failures()).toEqual([]);
  });

  it('לחיצה על פריט לוח נמסרת למעלה ואינה רצה כפקודה', async () => {
    await buttonById(harness, 'copy').trigger('click');

    expect(harness.adapter.calls).toHaveLength(0);
    expect(harness.wrapper.emitted('run')?.[0]?.[0]).toMatchObject({ id: 'copy' });
  });

  it('לחיצה על שורת כתיבה נמסרת למעלה עם הפעולה', async () => {
    await buttonById(harness, 'link').trigger('click');

    expect(harness.wrapper.emitted('run')?.[0]?.[0]).toMatchObject({
      id: 'link',
      run: { kind: 'action', action: 'link' },
    });
  });

  it('כל לחיצה סוגרת את התפריט', async () => {
    await buttonById(harness, 'link').trigger('click');

    expect(harness.wrapper.emitted('close')).toHaveLength(1);
  });

  it('פריט מנוטרל אינו רץ ואינו סוגר', async () => {
    const disabled = open(sections({ hasRange: false }));
    await buttonById(disabled, 'copy').trigger('click');

    expect(disabled.wrapper.emitted('run')).toBeUndefined();
    expect(disabled.wrapper.emitted('close')).toBeUndefined();
  });

  it('פריט מנוטרל נשאר בר-מיקוד — aria-disabled ולא disabled', () => {
    const disabled = open(sections({ hasRange: false }));
    const copy = buttonById(disabled, 'copy');

    expect(copy.attributes('aria-disabled')).toBe('true');
    expect(copy.attributes('disabled')).toBeUndefined();
  });

  it('בפתיחה שום פריט אינו מסומן — המיקוד על הכרטיס עצמו, כמו ב-Word', async () => {
    await nextTick();

    expect(document.activeElement).toBe(harness.wrapper.element);
    expect(harness.wrapper.findAll('[tabindex="0"]')).toHaveLength(0);
  });

  /**
   * המיקוד עצמו נבדק ולא רק ה-tabindex: `registerButton` יכול היה להימחק כולו
   * וכל הבדיקות היו נשארות ירוקות — ה-attribute הוא החיווי, `document.activeElement`
   * הוא מה שהמשתמש מקבל.
   */
  it('חץ למטה מהפתיחה בוחר את הפריט הראשון, ומזיז מיקוד אמיתי', async () => {
    await harness.wrapper.trigger('keydown', { key: 'ArrowDown' });
    await nextTick();

    expect(buttonById(harness, 'cut').attributes('tabindex')).toBe('0');
    expect(document.activeElement).toBe(buttonById(harness, 'cut').element);

    await harness.wrapper.trigger('keydown', { key: 'ArrowDown' });
    await nextTick();

    expect(document.activeElement).toBe(buttonById(harness, 'copy').element);
    expect(buttonById(harness, 'cut').attributes('tabindex')).toBe('-1');
  });

  it('End קופץ לפריט האחרון, ומשם החץ מתגלגל להתחלה', async () => {
    await harness.wrapper.trigger('keydown', { key: 'End' });
    await nextTick();
    expect(document.activeElement).toBe(buttonById(harness, 'select-all').element);

    await harness.wrapper.trigger('keydown', { key: 'ArrowDown' });
    await nextTick();
    expect(document.activeElement).toBe(buttonById(harness, 'cut').element);
  });

  it('בכותרת עליונה הפריטים שאינם שייכים נעלמים, והשאר עדיין נגישים', async () => {
    const header = open(sections({ storyType: 'header' }));

    expect(header.wrapper.find('[data-entry-id="footnote"]').exists()).toBe(false);
    expect(header.wrapper.find('[data-entry-id="insert-citation"]').exists()).toBe(false);
    expect(buttonById(header, 'link').exists()).toBe(true);
  });

  it('פתיחה מחדש בנקודה אחרת מאפסת את המיקוד', async () => {
    await harness.wrapper.trigger('keydown', { key: 'ArrowDown' });
    await nextTick();
    expect(harness.wrapper.findAll('[tabindex="0"]')).toHaveLength(1);

    await harness.wrapper.setProps({ point: { x: 700, y: 500 } });
    await nextTick();

    expect(harness.wrapper.findAll('[tabindex="0"]')).toHaveLength(0);
    expect(document.activeElement).toBe(harness.wrapper.element);
  });

  it('Tab סוגר — התפריט אינו אזור בממשק', async () => {
    await harness.wrapper.trigger('keydown', { key: 'Tab' });

    expect(harness.wrapper.emitted('close')).toHaveLength(1);
  });

  it('בלי מקטעים אין כרטיס בכלל', () => {
    const empty = open([]);

    expect(empty.wrapper.find('[role="menu"]').exists()).toBe(false);
  });

  /**
   * שורה שכתוב בה „קישור…” ולצדה „Ctrl+K” אינה צריכה כרטיס טולטיפ שאומר
   * „קישור…”, ולכן היא אינה מצהירה על `data-tip-*` כלל — מה שהופך פקד לעוגן.
   * `title` אינו נבדק כאן כי הוא אינו קיים באף אלמנט בתוכנה
   * (tests/unit/native-title.test.ts).
   */
  it('לשורת כתיבה אין טולטיפ — התווית כבר על המסך', () => {
    expect(buttonById(harness, 'link').attributes('data-tip-title')).toBeUndefined();
    expect(buttonById(harness, 'bold').attributes('data-tip-title')).toBe('מודגש');
    expect(buttonById(harness, 'bold').attributes('data-tip-shortcut')).toBe('Ctrl+B');
    expect(buttonById(harness, 'bold').attributes('aria-label')).toBe('מודגש (Ctrl+B)');
  });

  it('תווית הקיצור מוצגת ב-LTR', () => {
    const shortcut = buttonById(harness, 'link').find('.ctx-btn__shortcut');

    expect(shortcut.attributes('dir')).toBe('ltr');
    expect(shortcut.text()).toBe('Ctrl+K');
  });
});

/**
 * שורת הגופן.
 *
 * מה שנמדד כאן הוא הדבר שבגללו היא נבנתה כך ולא אחרת: הערך שהיא מציגה אינו
 * שלה. `FONT_MEMORY` מסופק מהמעטפת (App.vue) לרצועה ולתפריט גם יחד, ובדיקה
 * שמוסרת אותו זיכרון לשתי ההרכבות היא הבדיקה היחידה שיכולה לתפוס חזרה לעותק
 * פרטי — מצב שבו התפריט מציג „Assistant 12” בזמן שהרצועה מציגה את גופן המסמך.
 */
describe('שורת הגופן בתפריט ההקשר', () => {
  it('מציגה את מה שהמנוע מדווח על הבחירה', async () => {
    const adapter = createCommandDouble();
    adapter.setState('font-family', { value: 'TaameyDavidCLM' });
    adapter.setState('font-size', { value: 20 });

    const menu = open(sections(), { adapter });
    await settle();

    expect(pickerValue(menu.wrapper, 'גופן')).toBe('TaameyDavidCLM');
    expect(pickerValue(menu.wrapper, 'גודל גופן')).toBe('20');
  });

  it('בחירת גופן מגיעה למנוע עם payload שהוא מאשר, וסוגרת את הכרטיס', async () => {
    const menu = open();
    await settle();

    await setPicker(menu.wrapper, 'גופן', 'TaameyDavidCLM');
    await settle();

    expect(menu.adapter.payloads('font-family')).toEqual(['TaameyDavidCLM']);
    expect(menu.adapter.rejected).toEqual([]);
    expect(menu.wrapper.emitted('close')).toHaveLength(1);
  });

  /**
   * גודל שאינו בסולם של Word הוא הסיבה שהתיבה היא תיבת ערך ולא בורר סגור,
   * וההתנהגות הזאת חייבת להיות זהות לזו שברצועה — היא מגיעה מאותו `normalize`.
   */
  it('גודל שהוקלד ואינו ברשימה מוחל, ואינו נעלם לטובת ההתאמה הראשונה', async () => {
    const menu = open();
    await settle();

    await setPicker(menu.wrapper, 'גודל גופן', '13');
    await settle();

    expect(menu.adapter.payloads('font-size')).toEqual([13]);
  });

  it('מה שהרצועה מציגה הוא מה שהתפריט מציג — זיכרון אחד לשניהם', async () => {
    const fontMemory = createFontMemory();
    const adapter = createCommandDouble();

    const ribbon = mountUi(HomeTab, { adapter, fontMemory });
    await settle();
    await setPicker(ribbon.wrapper, 'גופן', 'TaameyDavidCLM');
    await settle();
    expect(pickerValue(ribbon.wrapper, 'גופן')).toBe('TaameyDavidCLM');

    // המנוע אינו מדווח ערך (כמו מיד אחרי שהתפריט הזיז את הסמן), ולכן זה בדיוק
    // המצב שבו עותק פרטי היה נופל לברירת המחדל.
    const menu = open(sections(), { adapter, fontMemory });
    await settle();

    expect(pickerValue(menu.wrapper, 'גופן')).toBe('TaameyDavidCLM');
  });

  /**
   * Y-PLONI#14 סעיף א, בתת-המקרה „בחרתי את מה שכבר היה”.
   *
   * `RibbonCombo.choose()` פולט `done` תמיד ו-`update:modelValue` רק כשהערך
   * **שונה** — ולכן כרטיס שהאזין להחלה בלבד נשאר פתוח בדיוק במקרה הזה: מילה
   * ב-Arial, פתיחת הרשימה, לחיצה על „Arial”. המיקוד נשאר ב-`input`, וההקלדה
   * הבאה נכנסה לתיבת הגופן ולא למסמך.
   *
   * הסגירה **היא** מה שמחזיר את המיקוד (`App.vue`, `closeContextMenu`), ולכן
   * זו הפליטה שנמדדת כאן.
   */
  it('בחירת הגופן שהתיבה כבר מציגה סוגרת את הכרטיס, אף שאין מה להחיל', async () => {
    const menu = open();
    await settle();
    expect(pickerValue(menu.wrapper, 'גופן')).toBe('Assistant');

    const box = boxInCard(menu, 'גופן');
    await box.trigger('focus');
    await settle();

    const current = menu.wrapper
      .findAll('[role="option"]')
      .find((row) => row.attributes('data-value') === 'Assistant');
    expect(current).toBeDefined();
    await current?.trigger('pointerdown');
    await settle();

    // שום דבר לא הוחל — הערך זהה — ובכל זאת הכרטיס נסגר פעם אחת בדיוק.
    expect(menu.adapter.payloads('font-family')).toEqual([]);
    expect(menu.wrapper.emitted('close')).toHaveLength(1);
  });

  /**
   * המקבילה של „וגם אחרי Escape בתיבה” שבבדיקת הרצועה
   * (tests/component/picker-state.test.ts): `RibbonCombo` עוצר את `Escape`
   * ופולט `done`, ובלי מאזין ה-`Escape` הראשון נבלע — כלומר נדרש שני.
   */
  it('Escape בתיבה שבכרטיס סוגר אותו, בלי להחיל דבר', async () => {
    const menu = open();
    await settle();

    const box = boxInCard(menu, 'גודל גופן');
    await box.trigger('focus');
    await box.trigger('keydown', { key: 'Escape' });
    await settle();

    expect(menu.adapter.payloads('font-size')).toEqual([]);
    expect(menu.wrapper.emitted('close')).toHaveLength(1);
  });

  /**
   * שני חצאים של אותו באג, ושניהם נראו רק מרגע שבורר הגודל הפך מ-`<select>`
   * לרשימה נפתחת: `preview` הוא `font-family` של CSS, ורשימה מספרית שהערך
   * הנוכחי בראשה אינה סולם.
   */
  it('גודל שאינו בסולם נכנס במקומו לפי הסדר, ובלי לטעון שהוא גופן', async () => {
    const adapter = createCommandDouble();
    adapter.setState('font-size', { value: 13 });

    const menu = open(sections(), { adapter });
    await settle();
    expect(pickerValue(menu.wrapper, 'גודל גופן')).toBe('13');

    const added = optionsOf(menu, 'גודל גופן').find((option) => option.value === '13');
    expect(added).toBeDefined();
    expect(added?.preview).toBeUndefined();

    await boxInCard(menu, 'גודל גופן').trigger('focus');
    await settle();

    const values = rowValues(menu);
    expect(values).toContain('13');
    expect(values).toEqual([...values].sort((a, b) => Number(a) - Number(b)));
    expect(values[values.indexOf('13') - 1]).toBe('12');
    expect(values[values.indexOf('13') + 1]).toBe('14');
  });

  it.each([
    { name: 'טרם נבדק', available: undefined, pending: true, unavailable: false },
    { name: 'זמין', available: true, pending: false, unavailable: false },
    { name: 'חסר', available: false, pending: false, unavailable: true },
  ])('גופן נוכחי שאינו ברשימה נוסף בראשה לפי מצב המטמון: $name', async ({ available, pending, unavailable }) => {
    const cached = vi.spyOn(fontAvailability, 'cachedFamilyAvailability').mockReturnValue(available);
    const measure = vi.spyOn(fontAvailability, 'isFamilyAvailable');
    try {
      const adapter = createCommandDouble();
      adapter.setState('font-family', { value: 'Guttman Yad' });
      const menu = open(sections(), { adapter });
      await settle();

      const current = optionsOf(menu, 'גופן')[0];
      expect(current?.value).toBe('Guttman Yad');
      expect(current?.preview).toBe(available === true ? 'Guttman Yad' : undefined);
      expect(current?.pending === true).toBe(pending);
      expect(current?.unavailable === true).toBe(unavailable);
      expect(current?.measured).toBe(available !== undefined);
      expect(measure).not.toHaveBeenCalled();
    } finally {
      measure.mockRestore();
      cached.mockRestore();
    }
  });

  it('גופן שהוחל מהתפריט מופיע ברצועה מיד', async () => {
    const fontMemory = createFontMemory();
    const adapter = createCommandDouble();

    const ribbon = mountUi(HomeTab, { adapter, fontMemory });
    const menu = open(sections(), { adapter, fontMemory });
    await settle();

    await setPicker(menu.wrapper, 'גופן', 'TaameyDavidCLM');
    await settle();

    expect(pickerValue(ribbon.wrapper, 'גופן')).toBe('TaameyDavidCLM');
  });
});

/**
 * זיכרון הבוררים והחלפת מסמך.
 *
 * שתי שכבות הזיכרון (`FONT_MEMORY`) קיימות בשביל רגע אחד: המנוע אינו מדווח
 * ערך — בחירה מעורבת, או בחירה שטרם נפתרה — ואז הבורר מציג את „האחרון שידענו”
 * במקום להתרוקן. ההצדקה שלהן היא שהערך נמדד **באותו מסמך**, ובהחלפה היא
 * נופלת: הזיכרון של הספר שנסגר אינו ידיעה על הטאב שנפתח, והרגע שבו הוא מוצג
 * הוא בדיוק הרגע הנפוץ — מסמך טרי, לפני שהבחירה התיישבה.
 *
 * ההחלפה נעשית כאן בשני צעדים, כמו במעטפת: `session` חדש הוא **אדפטר** חדש
 * (`useCommand` מנקה אז את ההחזקה שלו — ראו composables/useCommand.ts), ומונה
 * `DOCUMENT_GENERATION` עולה. בלי שני הצעדים אין מה למדוד: ההחזקה של הקריאה
 * הייתה מציגה את הערך הישן במקום הזיכרון, ושתי התקלות היו מתחפשות זו לזו.
 */
describe('החלפת מסמך', () => {
  it('מאפסת את זיכרון הבוררים — הכרטיס אינו מציג את הגופן של המסמך הקודם', async () => {
    const first = createCommandDouble();
    const menu = open(sections(), { adapter: first });
    await settle();

    // הדיווח **אחרי** ההרכבה, ולא כמצב פתיחה: זה מה שכותב לזיכרון. מצב שהיה
    // שם עוד לפני שהבורר עלה אינו „שינוי”, ואף watch אינו רואה אותו.
    first.setState('font-family', { value: 'TaameyDavidCLM' });
    first.setState('font-size', { value: 20 });
    await settle();
    expect(pickerValue(menu.wrapper, 'גופן')).toBe('TaameyDavidCLM');
    expect(pickerValue(menu.wrapper, 'גודל גופן')).toBe('20');

    // מסמך אחר, ועליו המנוע עדיין שותק — כלומר הזיכרון הוא כל מה שיש.
    await menu.setAdapter(createCommandDouble());
    await menu.setSuperdoc(createSuperdocDouble());
    await settle();

    expect(pickerValue(menu.wrapper, 'גופן')).toBe('Assistant');
    expect(pickerValue(menu.wrapper, 'גודל גופן')).toBe('12');
  });
});

/**
 * כרטיס שנפתח למעלה.
 *
 * לחיצה בתחתית החלון — סוף הטקסט, כלומר המקום הנפוץ ביותר — אינה מותירה מקום
 * מתחתיה, והכרטיס מתהפך: הקצה **התחתון** שלו נוגע בסמן. בסדר הרגיל זה מרחיק
 * מהסמן בדיוק את מה שצמוד אליו בפתיחה למטה — שורת הלוח, שורת הגופן ושורת
 * העיצוב — ומחייב לחזור עם העכבר לאורך כל הכרטיס.
 */
describe('כרטיס שנפתח למעלה', () => {
  /** רק שמות המקטעים, בסדר שבו הם מצוירים. */
  function groups(harness: Harness): (string | undefined)[] {
    return harness.wrapper.findAll('[role="group"]').map((group) => group.attributes('aria-label'));
  }

  it('בפתיחה למטה הסדר הוא זה של הדגם', async () => {
    const down = open(sections(), { props: { point: { x: 400, y: 200 } } });
    await nextTick();

    expect(groups(down)[0]).toBe('לוח');
    expect(groups(down)[1]).toBe('גופן');
  });

  it('בפתיחה למעלה סדר המקטעים מתהפך, והאייקונים יורדים ליד הסמן', async () => {
    const up = open(sections(), { props: { point: { x: 400, y: window.innerHeight - 3 } } });
    await nextTick();

    const order = groups(up);
    expect(order[order.length - 1]).toBe('לוח');
    expect(order[order.length - 2]).toBe('גופן');
    expect(order[0]).toBe('עריכה');
  });

  /**
   * ההיפוך לבדו דוחף אל **מתחת לקו** בדיוק את מה שהוא נועד לקרב.
   *
   * לכרטיס יש `max-height` שנגזר מהמקום שמעל הנקודה, ו-`overflow-y: auto`.
   * מה שצמוד לסמן יושב אחרי ההיפוך בסוף ה-DOM, ו-`scrollTop: 0` מראה את הקצה
   * הרחוק. נמדד ב-Chrome על ה-dist בחלון 756×413: `scrollHeight: 326` מול
   * `clientHeight: 268`, ו„הוסף למילון” — הפריט היחיד בכרטיס שנוגע במילה
   * שנלחצה — נחתך לגמרי; לחיצה לפי המלבן שלו נחתה על המסמך
   * (`scripts/qa/spellcheck-qa.mjs` ירד ל-11/12).
   *
   * ב-jsdom אין פריסה, ולכן `scrollHeight` מוזרק: מה שנמדד כאן הוא ההשמה
   * עצמה, כלומר שהכרטיס נגלל אל הקצה ולא נשאר על אפס.
   */
  it('כרטיס שנגלל נפתח על הקצה הצמוד לסמן, ולא על הרחוק', async () => {
    const up = open(sections(), { props: { point: { x: 400, y: window.innerHeight - 3 } } });
    const card = up.wrapper.find('[role="menu"]').element as HTMLElement;
    Object.defineProperty(card, 'scrollHeight', { configurable: true, value: 326 });
    Object.defineProperty(card, 'clientHeight', { configurable: true, value: 268 });

    // המדידה קובעת את הצד, וההיפוך נכנס ל-DOM אחריה; הגלילה באה אחרי שניהם.
    await nextTick();
    await nextTick();

    expect(card.scrollTop, 'הקצה שצמוד לסמן הוא מה שנראה').toBe(326);
  });

  it('בפתיחה למטה הכרטיס נשאר בראש — שם הקצה הצמוד לסמן הוא הראשון', async () => {
    const down = open(sections(), { props: { point: { x: 400, y: 200 } } });
    const card = down.wrapper.find('[role="menu"]').element as HTMLElement;
    Object.defineProperty(card, 'scrollHeight', { configurable: true, value: 326 });

    await nextTick();
    await nextTick();

    expect(card.scrollTop).toBe(0);
  });

  /**
   * החץ למטה חייב להזיז מיקוד למה שנמצא למטה **על המסך**, ולא למה שהמודל בנה
   * אחריו — אחרת בכרטיס שהתהפך הוא מזיז אותו למעלה.
   */
  it('החצים עוברים בסדר שעל המסך, ולא בסדר הדגם', async () => {
    const up = open(sections(), { props: { point: { x: 400, y: window.innerHeight - 3 } } });
    // המדידה קובעת את הצד, והחץ אחריה: לפניה הכרטיס עדיין מצויר בסדר הרגיל.
    await nextTick();
    await up.wrapper.trigger('keydown', { key: 'ArrowDown' });
    await nextTick();

    expect(document.activeElement).toBe(buttonById(up, 'find').element);
  });
});
