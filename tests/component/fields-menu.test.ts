/**
 * מה שהמנוע **קיבל בפועל** מפריטי התפריט של „מספר עמוד”.
 *
 * הסורק הגנרי ב-ribbon-tabs.test.ts לוחץ על הכפתור העליון של כל פקד ובודק
 * שמשהו קרה. בכפתור תפריט הלחיצה הזאת רק פותחת את הפופאובר — ה-handler של
 * `@select` אינו רץ בכלל, ולכן החלפה בין `'page'` ל-`'count'` ב-InsertTab
 * עברה את כל חבילת הבדיקות בירוק ושתלה `NUMPAGES` בכפתור „מספר עמוד”.
 *
 * כאן הפריט עצמו נלחץ, וה-`instruction` שהגיע ל-`fields.insert` הוא מה
 * שנמדד — כלומר בדיוק המחרוזת שתיכתב לקובץ של המשתמש. הבחירה היא לפי מיקום
 * בפופאובר, ולכן יש גם בדיקה שמקבעת איזו תווית יושבת בכל מיקום: בלעדיה
 * החלפת שתי התוויות ביניהן הייתה עוברת בירוק.
 */
import { describe, expect, it } from 'vitest';
import InsertTab from '../../src/ui/ribbon/tabs/InsertTab.vue';
import { autoUnmount, mountUi, settle, tipStartsSelector } from './harness';

autoUnmount();

/** הפריטים בסדר שבו הם מופיעים בפופאובר, ומה שכל אחד מהם מכניס למסמך. */
const MENU_ITEMS = [
  { label: 'מספר עמוד', instruction: 'PAGE' },
  { label: 'מספר העמודים במסמך', instruction: 'NUMPAGES' },
] as const;

const MENU_BUTTON = tipStartsSelector('הכנסת שדה מספר עמוד');

describe('תפריט „מספר עמוד”', () => {
  it.each(MENU_ITEMS.map((item, index) => ({ index, ...item })))(
    'הפריט „$label” שולח `$instruction`',
    async ({ index, instruction }) => {
      const harness = mountUi(InsertTab);
      await settle();

      await harness.wrapper.find(MENU_BUTTON).trigger('click');
      await settle();

      const items = harness.wrapper.findAll('.ribbon-menu__item');
      expect(items).toHaveLength(MENU_ITEMS.length);
      await items[index].trigger('click');
      await settle();

      const inputs = harness.superdoc.inputs('fields.insert') as { instruction: string }[];
      expect(inputs.map((input) => input.instruction)).toEqual([instruction]);
    },
  );

  it('לכל מיקום בפופאובר יש התווית שהבדיקה שלמעלה מניחה', async () => {
    const harness = mountUi(InsertTab);
    await settle();

    await harness.wrapper.find(MENU_BUTTON).trigger('click');
    await settle();

    expect(harness.wrapper.findAll('.ribbon-menu__item-label').map((node) => node.text())).toEqual(
      MENU_ITEMS.map((item) => item.label),
    );
  });
});

/**
 * „תאריך ושעה” הוא כפתור מפוצל (`RibbonMenuButton` עם `split`): הגוף מכניס
 * תאריך+שעה, והחץ פותח תפריט עם שתי ברירות נוספות — תאריך בלבד ושעה בלבד —
 * ומתחת לקו בחירת הלוח. אותה סכנה כמו ב„מספר עמוד” למעלה: לחיצה על הגוף רק
 * פותחת תפריט בכפתור לא-מפוצל, ולכן כאן נבדק גם שהגוף מפעיל בלי לפתוח פופאובר.
 *
 * ברירת המחדל היא הלוח העברי, ובו התאריך נכתב כטקסט (`insert`) ולא כשדה —
 * ראו engine/hebrew-date.ts.
 */
const DATE_MENU_LABELS = ['תאריך ושעה', 'תאריך בלבד', 'שעה בלבד', 'לוח עברי', 'לוח לועזי'];

const GREGORIAN_ITEMS = [
  { label: 'תאריך ושעה', instruction: 'DATE \\@ "dd/MM/yyyy HH:mm"' },
  { label: 'תאריך בלבד', instruction: 'DATE \\@ "dd/MM/yyyy"' },
  { label: 'שעה בלבד', instruction: 'DATE \\@ "HH:mm"' },
] as const;

const HEBREW_DATE = /^[א-ת]{1,2}[׳״][א-ת]? ב[א-ת]+( [אב]׳)? [א-ת]{2,3}״[א-ת]/;

async function openDateMenu(harness: ReturnType<typeof mountUi>) {
  await harness.wrapper.find('.word-split__arrow').trigger('click');
  await settle();
  return harness.wrapper.findAll('.ribbon-menu__item');
}

async function chooseGregorian(harness: ReturnType<typeof mountUi>) {
  const items = await openDateMenu(harness);
  await items[DATE_MENU_LABELS.indexOf('לוח לועזי')].trigger('click');
  await settle();
}

/**
 * כפיל הבחירה מחזיר `selectionTarget` רק כשיש טווח (ראו harness.ts), והכתיבה
 * של התאריך העברי צריכה אותו.
 */
function mountWithWritableSelection() {
  const harness = mountUi(InsertTab);
  harness.superdoc.setSelection({ hasRange: true, text: 'abcd' });
  return harness;
}

const insertedText = (harness: ReturnType<typeof mountUi>) =>
  (harness.superdoc.inputs('insert') as { value: string; type: string }[]).map((input) => input.value);

describe('כפתור „תאריך ושעה” המפוצל', () => {
  it('ברירת המחדל: הגוף כותב תאריך עברי ושעה כטקסט, בלי שדה ובלי לפתוח תפריט', async () => {
    const harness = mountWithWritableSelection();
    await settle();

    await harness.wrapper.find('.word-split .word-btn').trigger('click');
    await settle();

    const texts = insertedText(harness);
    expect(texts).toHaveLength(1);
    expect(texts[0]).toMatch(HEBREW_DATE);
    expect(texts[0]).toMatch(/ \d\d:\d\d$/);
    expect(harness.superdoc.inputs('fields.insert')).toEqual([]);
    expect(harness.wrapper.find('.ribbon-menu__popover').exists()).toBe(false);
  });

  it('החץ פותח את התפריט, ואינו מכניס דבר; „לוח עברי” מסומן', async () => {
    const harness = mountUi(InsertTab);
    await settle();

    const items = await openDateMenu(harness);

    expect(harness.wrapper.find('.ribbon-menu__popover').exists()).toBe(true);
    expect(harness.superdoc.inputs('fields.insert')).toEqual([]);
    expect(harness.superdoc.inputs('insert')).toEqual([]);
    expect(harness.wrapper.findAll('.ribbon-menu__item-label').map((node) => node.text())).toEqual(
      DATE_MENU_LABELS,
    );
    expect(items.map((item) => item.attributes('aria-checked'))).toEqual([
      undefined,
      undefined,
      undefined,
      'true',
      'false',
    ]);
  });

  it('„תאריך בלבד” בלוח העברי כותב תאריך בלי שעה', async () => {
    const harness = mountWithWritableSelection();
    await settle();

    const items = await openDateMenu(harness);
    await items[1].trigger('click');
    await settle();

    const texts = insertedText(harness);
    expect(texts).toHaveLength(1);
    expect(texts[0]).toMatch(HEBREW_DATE);
    expect(texts[0]).not.toMatch(/\d/);
  });

  it('„שעה בלבד” נשארת שדה גם בלוח העברי', async () => {
    const harness = mountUi(InsertTab);
    await settle();

    const items = await openDateMenu(harness);
    await items[2].trigger('click');
    await settle();

    const inputs = harness.superdoc.inputs('fields.insert') as { instruction: string }[];
    expect(inputs.map((input) => input.instruction)).toEqual(['DATE \\@ "HH:mm"']);
    expect(harness.superdoc.inputs('insert')).toEqual([]);
  });

  it('אחרי בחירת „לוח לועזי” הגוף מכניס את השדה הלועזי, והסימון עובר', async () => {
    const harness = mountUi(InsertTab);
    await settle();
    await chooseGregorian(harness);
    expect(harness.superdoc.inputs('fields.insert')).toEqual([]);

    await harness.wrapper.find('.word-split .word-btn').trigger('click');
    await settle();

    const inputs = harness.superdoc.inputs('fields.insert') as { instruction: string }[];
    expect(inputs.map((input) => input.instruction)).toEqual(['DATE \\@ "dd/MM/yyyy HH:mm"']);
    expect(harness.superdoc.inputs('insert')).toEqual([]);

    const items = await openDateMenu(harness);
    expect(items.slice(3).map((item) => item.attributes('aria-checked'))).toEqual(['false', 'true']);
  });

  it.each(GREGORIAN_ITEMS.map((item, index) => ({ index, ...item })))(
    'בלוח לועזי, הפריט „$label” שולח `$instruction`',
    async ({ index, instruction }) => {
      const harness = mountUi(InsertTab);
      await settle();
      await chooseGregorian(harness);

      const items = await openDateMenu(harness);
      await items[index].trigger('click');
      await settle();

      const inputs = harness.superdoc.inputs('fields.insert') as { instruction: string }[];
      expect(inputs.map((input) => input.instruction)).toEqual([instruction]);
    },
  );
});
