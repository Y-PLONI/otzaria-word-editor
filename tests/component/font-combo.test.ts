/**
 * בורר הגופן עם החיפוש (`RibbonCombo`).
 *
 * ההכרעות הטהורות נבדקות ב-`tests/unit/font-search.test.ts`. מה שנבדק **כאן**
 * הוא מה שרק הרכבה יכולה לתפוס: שהמקלדת באמת מחוברת, שהבחירה באמת נפלטת,
 * ושיציאה מהשדה **אינה** מחילה גופן — שלושתם „HTML תקין לחלוטין” בכל סריקת
 * מקור, ושלושתם היו שוברים את הפקד בשקט.
 *
 * ובנוסף רגרסיה אחת שנצפתה בפועל: הרשימה נחתכה בגובה הרצועה. `position` נבדק
 * ולא המראה — jsdom אינו מודד פריסה, אבל הוא כן אומר איזה `position` הוחל,
 * וזו ההבחנה בין „נמדד ב-fixed” לבין „absolute שנחתך”.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { mount, type VueWrapper } from '@vue/test-utils';
import { nextTick, shallowRef } from 'vue';
import type { SuperDoc } from 'superdoc';
import { ACTIVE_SUPERDOC } from '../../src/engine/document-api';
import { FONT_PROBE_PRIORITIES } from '../../src/composables/keys';
import { withCurrent } from '../../src/composables/picker-value';
import RibbonCombo from '../../src/ui/ribbon/common/RibbonCombo.vue';
import type { ComboOption } from '../../src/ui/ribbon/font-search';

const OPTIONS: readonly ComboOption[] = [
  { value: 'Assistant', label: 'Assistant', group: '', hebrew: true },
  { value: 'TaameyDavidCLM', label: 'David (TaameyDavidCLM)', group: '', hebrew: true },
  { value: 'Arial', label: 'Arial', group: '', hebrew: true },
  { value: 'David', label: 'David', group: 'עברית', hebrew: true },
  { value: 'Narkisim', label: 'Narkisim', group: 'עברית', hebrew: true },
  { value: 'Arial Black', label: 'Arial Black', group: 'כל הגופנים', hebrew: false },
];

let combo: VueWrapper;

function open(): Promise<void> {
  return combo.find('input').trigger('focus').then(() => nextTick());
}

async function type(text: string): Promise<void> {
  const input = combo.find('input');
  (input.element as HTMLInputElement).value = text;
  await input.trigger('input');
  await nextTick();
}

const key = (name: string) => combo.find('input').trigger('keydown', { key: name });

const optionValues = () =>
  combo.findAll('[role="option"]').map((el) => el.attributes('data-value'));

const emitted = () => combo.emitted('update:modelValue') as string[][] | undefined;

beforeEach(() => {
  combo = mount(RibbonCombo, {
    props: { modelValue: 'Arial', options: OPTIONS, title: 'גופן' },
  });
});

describe('פתיחה', () => {
  it('הרשימה אינה ב-DOM עד שפותחים — 294 שורות אינן נבנות בכל רינדור', async () => {
    expect(combo.find('[role="listbox"]').exists()).toBe(false);
    await open();
    expect(combo.find('[role="listbox"]').exists()).toBe(true);
  });

  it('התיבה מציגה את הגופן הנוכחי, לא ריקה', () => {
    expect((combo.find('input').element as HTMLInputElement).value).toBe('Arial');
  });

  it('נפתחת על הגופן הנוכחי ולא על ראש הרשימה', async () => {
    // כך „פתח, חץ אחד” מגיע לשכן — ולא קופץ ל-Assistant מכל מקום.
    await open();
    expect(combo.find('[role="option"].active').attributes('data-value')).toBe('Arial');
  });

  it('הקבוצות מוצגות ככותרות ולא כאפשרויות', async () => {
    await open();
    const headings = combo.findAll('.ribbon-combo-group').map((el) => el.text());
    expect(headings).toEqual(['עברית', 'כל הגופנים']);
    expect(optionValues()).toHaveLength(OPTIONS.length);
  });
});

/**
 * הדגימה העברית. הטקסט עצמו הוא `content` של `::before` — jsdom אינו מחשב
 * גיליונות, ולכן מה שנבדק הוא מה שהתבנית כן מכריעה: **אילו** שורות מקבלות
 * אותה, ושהשם הנגיש נשאר השם ולא „אבגד Narkisim”.
 */
describe('דגימת אותיות עבריות', () => {
  const classesOf = (value: string) =>
    combo.find(`[role="option"][data-value="${value}"]`).classes();

  it('רק שורה של גופן שמכסה עברית מקבלת דגימה', async () => {
    await open();
    expect(classesOf('Narkisim')).toContain('hebrew');
    expect(classesOf('Arial')).toContain('hebrew');
    expect(classesOf('Arial Black')).not.toContain('hebrew');
  });

  it('אפשרות בלי הדגל אינה מקבלת דגימה — בורר הגודל, וגופן שאינו מהמנייה', async () => {
    await combo.setProps({ options: [{ value: '12', label: '12' }], modelValue: '12' });
    await open();
    expect(classesOf('12')).not.toContain('hebrew');
  });

  /**
   * עמודת הדגימה. ה-CSS הוא שמצמיד את האותיות לקצה אחד ואת השם לשני, ו-jsdom
   * אינו מחשב אותו — אבל **התנאי** להחלתו הוא מחלקה על הרשימה, וזו הכרעה של
   * התבנית: בלעדיה השורות חוזרות לצמוד-לצמוד, ובורר הגודל היה מקבל עמודה
   * ריקה של 3em לפני „12”.
   */
  describe('עמודת הדגימה', () => {
    const listClasses = () => combo.find('[role="listbox"]').classes();

    it('רשימה שיש בה גופנים עבריים מקבלת את העמודה', async () => {
      await open();
      expect(listClasses()).toContain('has-glyphs');
    });

    it('בורר הגודל אינו מקבל אותה — אין לו מה להדגים', async () => {
      await combo.setProps({
        options: [
          { value: '12', label: '12' },
          { value: '14', label: '14' },
        ],
        modelValue: '12',
      });
      await open();
      expect(listClasses()).not.toContain('has-glyphs');
    });

    it('חיפוש שכל התאמותיו לטיניות אינו מבטל אותה', async () => {
      // אחרת השמות היו קופצים לצד השני באמצע ההקלדה — בדיוק בשורה שהמשתמש
      // מסתכל עליה. העמודה נגזרת מהמלאי, לא מהתוצאות.
      await open();
      await type('Arial Black');
      expect(optionValues()).toEqual(['Arial Black']);
      expect(listClasses()).toContain('has-glyphs');
    });
  });

  it('השם הנגיש הוא שם הגופן, בלי הדגימה שלפניו', async () => {
    // Chrome מכליל תוכן של `::before` בשם הנגיש, וקורא מסך היה מכריז את
    // אותיות הדגימה לפני כל שם עברי ברשימה.
    await open();
    const option = combo.find('[role="option"][data-value="Narkisim"]');
    expect(option.attributes('aria-label')).toBe('Narkisim');
    expect(combo.find('[role="option"][data-value="Arial Black"]').attributes('aria-label'))
      .toBeUndefined();
  });
});

/**
 * מה שהפקד מוסר לתצוגה החיה. ההחלה עצמה אינה כאן — היא ב-
 * tests/unit/font-preview.test.ts — אבל בלי הפליטות האלה אין לה קלט בכלל,
 * וזו בדיוק תקלה ש„HTML תקין לחלוטין” אינו מראה.
 */
/**
 * הרשימה קופאת ברגע הפתיחה. הרגרסיה שדווחה: התצוגה החיה משנה את הגופנים
 * שהמסמך משתמש בהם, המנוע מרכיב מהם את קבוצת „אחרונים”, וכל שורה שרוחפים
 * מעליה קפצה לראש הרשימה — כלומר השורה שמתחת לעכבר התחלפה באמצע התנועה.
 */
describe('הקפאת הרשימה בזמן שהיא פתוחה', () => {
  it('אפשרויות שהתחלפו תוך כדי אינן משנות את הסדר', async () => {
    await open();
    await combo.setProps({
      options: [{ value: 'Narkisim', label: 'Narkisim', group: '' }, ...OPTIONS],
    });
    expect(optionValues()).toEqual(OPTIONS.map((option) => option.value));
  });

  it('הפתיחה הבאה כבר מציגה את מה שהתחלף', async () => {
    await open();
    const next = [{ value: 'Gisha', label: 'Gisha', group: '' }, ...OPTIONS];
    await combo.setProps({ options: next });
    await combo.find('input').trigger('blur');
    await open();
    expect(optionValues()).toEqual(next.map((option) => option.value));
  });

  it('גם רשימה שנפתחה מהקלדה מוקפאת — לא רק מפתיחה', async () => {
    // המסלול שנמדד: מיקוד בתיבה (נפתח ומוקפא) → לחיצה על החץ (סוגר ומאפס
    // את ההקפאה, והמיקוד **נשאר** בתיבה) → תו אחד. `onInput` פתח את הרשימה
    // בלי להקפיא, ומשם השורות זזות תחת העכבר — בדיוק הרגרסיה שההקפאה נוספה
    // בשבילה.
    await open();
    await combo.find('.ribbon-combo-arrow').trigger('pointerdown');
    await nextTick();
    expect(combo.find('[role="listbox"]').exists()).toBe(false);

    await type('a');
    expect(combo.find('[role="listbox"]').exists()).toBe(true);

    await combo.setProps({
      options: [{ value: 'Aharoni', label: 'Aharoni', group: '' }, ...OPTIONS],
    });
    // „a” תואם ל-Aharoni; רשימה לא מוקפאת הייתה מציגה אותו.
    expect(optionValues()).not.toContain('Aharoni');
  });

  it('והקלדה נוספת אינה מאפסת את ההקפאה שכבר קיימת', async () => {
    await open();
    await type('a');
    await combo.setProps({
      options: [{ value: 'Aharoni', label: 'Aharoni', group: '' }, ...OPTIONS],
    });
    await type('ar');
    expect(optionValues()).not.toContain('Aharoni');
  });
});

describe('פליטות לתצוגה החיה', () => {
  const last = <T,>(rows: T[][]): T[] | undefined => rows[rows.length - 1];
  const previews = () => last((combo.emitted('preview') as (string | null)[][] | undefined) ?? []);
  const ends = () => last((combo.emitted('previewEnd') as boolean[][] | undefined) ?? []);

  it('פתיחה מסמנת את הגופן הנוכחי, וחץ מוסר את השכן', async () => {
    await open();
    expect(previews()).toEqual(['Arial']);
    await key('ArrowDown');
    await nextTick();
    expect(previews()).toEqual(['David']);
  });

  it('עכבר על שורה מוסר אותה', async () => {
    await open();
    await combo.find('[role="option"][data-value="Narkisim"]').trigger('mousemove');
    await nextTick();
    expect(previews()).toEqual(['Narkisim']);
  });

  it('הקלדה מוסרת את ההתאמה הראשונה', async () => {
    await open();
    await type('nar');
    expect(previews()).toEqual(['Narkisim']);
  });

  it('בחירה נסגרת כ„נבחר”, ו-Escape כ„יצא”', async () => {
    await open();
    await combo.find('[role="option"][data-value="Narkisim"]').trigger('pointerdown');
    expect(ends()).toEqual([true]);

    await open();
    await key('Escape');
    expect(ends()).toEqual([false]);
  });

  it('יציאה מהשדה נסגרת כ„יצא” — מה שמחזיר את הגופן שהיה', async () => {
    await open();
    await combo.find('input').trigger('blur');
    expect(ends()).toEqual([false]);
  });
});

describe('חיפוש', () => {
  it('הקלדה מסננת', async () => {
    await open();
    await type('ari');
    expect(optionValues()).toEqual(['Arial', 'Arial Black']);
  });

  it('שאילתה בלי התאמות מציגה הודעה ולא רשימה ריקה בשקט', async () => {
    await open();
    await type('zzzz');
    expect(optionValues()).toEqual([]);
    expect(combo.find('.ribbon-combo-empty').exists()).toBe(true);
  });

  it('אחרי הקלדה הסימון על ההתאמה הראשונה — Enter בלי חץ נוסף', async () => {
    await open();
    await type('nar');
    await key('Enter');
    expect(emitted()?.[0]).toEqual(['Narkisim']);
  });
});

describe('מקלדת', () => {
  it('חץ למטה מזיז את הסימון, ו-Enter מחיל', async () => {
    await open();
    await key('ArrowDown');
    await key('Enter');
    expect(emitted()?.[0]).toEqual(['David']);
  });

  it('Escape סוגר בלי להחיל, והתיבה חוזרת לגופן שהיה', async () => {
    await open();
    await type('nar');
    await key('Escape');
    await nextTick();
    expect(emitted()).toBeUndefined();
    expect(combo.find('[role="listbox"]').exists()).toBe(false);
    expect((combo.find('input').element as HTMLInputElement).value).toBe('Arial');
  });

  it('בחירה בגופן שכבר נבחר אינה פולטת אירוע', async () => {
    // אחרת כל פתיחה-וסגירה הייתה מסמנת את המסמך כמלוכלך.
    await open();
    await key('Enter');
    expect(emitted()).toBeUndefined();
  });
});

describe('עכבר ופוקוס', () => {
  it('לחיצה על שורה מחילה אותה', async () => {
    await open();
    await combo.findAll('[role="option"]')[4].trigger('pointerdown');
    expect(emitted()?.[0]).toEqual(['Narkisim']);
  });

  it('`mousedown` לבדו אינו בוחר — הרשימה מאזינה ל-`pointerdown` בלבד', async () => {
    // ביטול `pointerdown` מדכא בדפדפן את `mousedown` התואם, ולכן מאזין שני על
    // `mousedown` היה חי רק ב-jsdom — ושם היה בוחר פעמיים. הבדיקה מקבעת שיש
    // מאזין אחד, כדי שהכפילות לא תחזור בשם „גם וגם".
    await open();
    await combo.findAll('[role="option"]')[4].trigger('mousedown');
    expect(emitted()).toBeUndefined();
    expect(combo.find('[role="listbox"]').exists()).toBe(true);
  });

  it('בחירת גופן מחזירה פוקוס למסמך דרך ACTIVE_SUPERDOC', async () => {
    const focusSpy = vi.fn();
    const fakeSuperdoc = shallowRef({ focus: focusSpy } as unknown as SuperDoc);
    const wrapper = mount(RibbonCombo, {
      props: { modelValue: 'Arial', options: OPTIONS, title: 'גופן' },
      global: {
        provide: {
          [ACTIVE_SUPERDOC as symbol]: fakeSuperdoc,
        },
      },
    });

    await wrapper.find('input').trigger('focus');
    await nextTick();
    await wrapper.findAll('[role="option"]')[4].trigger('pointerdown');
    expect(focusSpy).toHaveBeenCalledWith({ restoreSelection: true });
  });

  it('יציאה מהשדה סוגרת ואינה מחילה', async () => {
    // הקלדה בלי אישור אינה בחירה — עדיף על להחיל גופן בטעות על טקסט מסומן.
    await open();
    await type('nar');
    await combo.find('input').trigger('blur');
    await nextTick();
    expect(emitted()).toBeUndefined();
    expect(combo.find('[role="listbox"]').exists()).toBe(false);
  });
});

/**
 * המצב השני של הפקד — תיבת ערך (בורר הגודל).
 *
 * מה שנבדק כאן ולא ב-`font-search`: שהסימון האוטומטי אחרי הקלדה **אינו**
 * קורה במצב הזה. זו נקודת החיבור היחידה, והיא גם הבאג היחיד שאפשר להחזיר
 * בטעות — בלעדיה „13” היה מוחל כ„10”, ו-`commitValue` היה מחזיר את הערך
 * הנכון בלי שאיש יקרא לו איתו.
 */
describe('תיבת ערך', () => {
  const SIZES: readonly ComboOption[] = [8, 10, 11, 12, 14, 18].map((size) => ({
    value: String(size),
    label: String(size),
  }));

  /** נרמול מקוצר: מספר בלבד, וכל השאר נדחה. */
  const asSize = (typed: string) =>
    typed.trim() !== '' && Number.isFinite(Number(typed)) ? typed : null;

  beforeEach(() => {
    combo = mount(RibbonCombo, {
      props: { modelValue: '12', options: SIZES, normalize: asSize, title: 'גודל גופן' },
    });
  });

  it('גודל שאינו ברשימה מוחל כפי שהוקלד', async () => {
    await open();
    await type('13');
    await key('Enter');
    expect(emitted()?.[0]).toEqual(['13']);
  });

  it('ואינו נגרר להתאמה הראשונה: „1” הוא 1, לא 10', async () => {
    await open();
    await type('1');
    expect(optionValues()).toEqual(['10', '11', '12', '14', '18']);
    expect(combo.find('[role="option"].active').exists()).toBe(false);
    await key('Enter');
    expect(emitted()?.[0]).toEqual(['1']);
  });

  it('הרשימה עדיין רשימה — חץ ואז Enter בוחרים ממנה', async () => {
    await open();
    await type('1');
    await key('ArrowDown');
    await key('Enter');
    expect(emitted()?.[0]).toEqual(['10']);
  });

  it('קלט שאינו גודל אינו מוחל דבר, והתיבה חוזרת לגודל שהיה', async () => {
    await open();
    await type('גדול');
    await key('Enter');
    await nextTick();
    expect(emitted()).toBeUndefined();
    expect((combo.find('input').element as HTMLInputElement).value).toBe('12');
  });

  it('הודעת „אין התאמה” היא של הפקד הזה, ולא „אין גופן בשם הזה”', async () => {
    combo = mount(RibbonCombo, {
      props: {
        modelValue: '12',
        options: SIZES,
        normalize: asSize,
        emptyText: 'Enter מחיל את הגודל שהוקלד',
      },
    });
    await open();
    await type('13');
    expect(combo.find('.ribbon-combo-empty').text()).toBe('Enter מחיל את הגודל שהוקלד');
  });
});

/**
 * „סיימתי” — האירוע שמי שמרכיב את הפקד מחזיר לפיו את המיקוד למסמך.
 *
 * הסדר מול `update:modelValue` הוא **ההכרעה** ולא פרט: ההחלה צריכה לרוץ
 * כשהמסמך כבר ממוקד והבחירה שוחזרה. נמדד כאן דרך מאזינים אמיתיים, כי
 * `emitted()` של Vue Test Utils מקבץ לפי שם אירוע ואינו שומר סדר ביניהם.
 */
describe('סיום העבודה עם הפקד', () => {
  function withLog(props: Record<string, unknown> = {}) {
    const order: string[] = [];
    combo = mount(RibbonCombo, {
      props: {
        modelValue: 'Arial',
        options: OPTIONS,
        onDone: () => order.push('done'),
        'onUpdate:modelValue': (value: string) => order.push(`value:${value}`),
        ...props,
      },
    });
    return order;
  }

  it('בחירה פולטת „סיימתי” — ולפני הערך', async () => {
    const order = withLog();
    await open();
    await type('nar');
    await key('Enter');
    expect(order).toEqual(['done', 'value:Narkisim']);
  });

  it('גם בחירה בעכבר', async () => {
    const order = withLog();
    await open();
    await combo.findAll('[role="option"]')[4].trigger('pointerdown');
    expect(order).toEqual(['done', 'value:Narkisim']);
  });

  it('Escape הוא סיום גם הוא — מי שוויתר רוצה לחזור לכתוב', async () => {
    const order = withLog();
    await open();
    await type('nar');
    await key('Escape');
    expect(order).toEqual(['done']);
  });

  it('יציאה מהשדה אינה סיום — המשתמש כבר העביר את המיקוד בעצמו', async () => {
    const order = withLog();
    await open();
    await type('nar');
    await combo.find('input').trigger('blur');
    await nextTick();
    expect(order).toEqual([]);
  });

  it('בחירה בערך שכבר נבחר פולטת „סיימתי” בלי ערך', async () => {
    // אחרת „פתחתי, לחצתי על מה שכבר מסומן” היה משאיר את המיקוד ברצועה.
    const order = withLog();
    await open();
    await key('Enter');
    expect(order).toEqual(['done']);
  });
});

describe('רגרסיה: הרשימה נחתכה בגובה הרצועה', () => {
  it('הרשימה ממוקמת ב-fixed ולא ב-absolute', async () => {
    // `.word-ribbon-body` הוא `overflow-y: hidden`, ולכן `absolute` נחתך —
    // מה שנצפה כרשימה של שלוש שורות עם פס גלילה.
    // ראו composables/popover-position.ts.
    await open();
    expect(combo.find('[role="listbox"]').attributes('style')).toContain('position: fixed');
  });
});

/**
 * פס הדגימה — הטקסט שהמשתמש סימן במסמך, בגופן שהסימון עומד עליו.
 *
 * זו התצוגה שנשלחת **במקום** צביעה של המסמך עצמו, אחרי שזו נמדדה כמטעה
 * (`docs/engine-gaps.md`, „תצוגה חיה של גופן”). מה שנבדק כאן הוא מה שאי אפשר
 * לראות בקריאת התבנית: שהפס עוקב אחרי הסימון גם במקלדת ולא רק בעכבר, שהוא
 * אינו הופך לאפשרות שאפשר לבחור, ושבורר הגודל אינו מקבל אותו בכלל.
 */
describe('פס הדגימה', () => {
  const SAMPLED: readonly ComboOption[] = [
    { value: 'David', label: 'David', preview: 'David', hebrew: true },
    { value: 'Narkisim', label: 'Narkisim', preview: 'Narkisim', hebrew: true },
    { value: 'Rubik', label: 'Rubik', preview: 'Rubik', hebrew: true },
  ];

  const bar = () => combo.find('.ribbon-combo-sample');
  const barFamily = () => (bar().element as HTMLElement).style.fontFamily;
  const barSize = () => (bar().element as HTMLElement).style.fontSize;

  it('בלי טקסט אין פס — וזה בורר הגודל, שאין לו מה להדגים', async () => {
    await open();
    expect(bar().exists()).toBe(false);
    expect(combo.find('[role="listbox"]').classes()).not.toContain('has-sample');
  });

  it('עם טקסט — הפס מציג אותו כשהרשימה פתוחה, ולא לפני כן', async () => {
    await combo.setProps({ sample: 'בראשית ברא' });
    expect(bar().exists()).toBe(false);

    await open();
    expect(bar().text()).toBe('בראשית ברא');
    expect(combo.find('[role="listbox"]').classes()).toContain('has-sample');
  });

  it('הפס מצייר בגופן שהסימון עומד עליו, ועוקב אחרי החצים', async () => {
    // חץ ולא עכבר: `mousemove` הוא המסלול המובן מאליו, וניווט מקלדת הוא זה
    // שנשבר בשקט אם הפס נשען על אירוע העכבר במקום על הסימון עצמו.
    await combo.setProps({ options: SAMPLED, modelValue: 'David', sample: 'שלום' });
    await open();
    expect(barFamily()).toBe('David');

    await key('ArrowDown');
    await nextTick();
    expect(barFamily()).toBe('Narkisim');

    await key('ArrowUp');
    await nextTick();
    expect(barFamily()).toBe('David');
  });

  it('הפס נפתח על הגופן הנוכחי — ולא ריק עד התזוזה הראשונה', async () => {
    await combo.setProps({ options: SAMPLED, modelValue: 'Rubik', sample: 'שלום' });
    await open();
    expect(barFamily()).toBe('Rubik');
  });

  it('גופן שאינו ברשימה — הפס מוצג בגופן הרגיל, ולא ב-`font-family` מומצא', async () => {
    await combo.setProps({ options: SAMPLED, modelValue: 'אין כזה', sample: 'שלום' });
    await open();
    expect(bar().exists()).toBe(true);
    expect(barFamily()).toBe('');
  });

  it('חיפוש בלי התאמות אינו מציג פס — אין אפשרות מסומנת שהוא מדגים', async () => {
    await combo.setProps({ sample: 'בראשית ברא' });
    await open();
    await type('אין גופן כזה בעולם');
    expect(combo.find('.ribbon-combo-empty').exists()).toBe(true);
    expect(bar().exists()).toBe(false);
  });

  it('הפס אינו אפשרות: לא נספר, לא נבחר, ולא מוכרז', async () => {
    // `li` בתוך `listbox` שאינו `presentation` הוא אפשרות לכל דבר — כלומר
    // שורה שהחצים עוצרים עליה ושקורא מסך מונה. והכרזה שלו הייתה קוראת מחדש
    // את הטקסט שהמשתמש עצמו סימן, בכל תזוזת חץ.
    await combo.setProps({ options: SAMPLED, modelValue: 'David', sample: 'שלום' });
    await open();

    expect(optionValues()).toEqual(['David', 'Narkisim', 'Rubik']);
    expect(bar().attributes('role')).toBe('presentation');
    expect(bar().attributes('aria-hidden')).toBe('true');
    expect(bar().attributes('aria-selected')).toBeUndefined();

    await key('ArrowDown');
    await key('ArrowDown');
    await key('ArrowDown');
    await nextTick();
    // גלישה חוזרת לראש הרשימה, ולא נעצרת על הפס.
    expect(combo.find('input').attributes('aria-activedescendant')).toBe(
      combo.find('[role="option"][data-value="David"]').attributes('id'),
    );
  });

  it('`dir="auto"` על הפס — ציטוט לטיני במסמך עברי אינו מקבל את הפיסוק בצד הלא נכון', async () => {
    // הדגימה היא הטקסט של המשתמש, והמעטפת RTL. בלי `dir="auto"` הנקודה של
    // „Hello world.” נוחתת בקצה השמאלי של המחרוזת.
    await combo.setProps({ sample: 'Hello world.' });
    await open();
    expect(bar().attributes('dir')).toBe('auto');
  });

  it('הפס מצייר בגודל שהמסמך מדווח — לא בגודל קבוע של הרשימה', async () => {
    // זה מה שהופך את הדגימה ל„איך הטקסט שלי ייראה”: הגופן שהסימון עומד עליו,
    // ובגודל שבמסמך. הערך מגיע מוכן מ-`use-font-controls` (`sampleSize`).
    await combo.setProps({ options: SAMPLED, modelValue: 'David', sample: 'שלום', sampleSize: '21px' });
    await open();
    expect(barSize()).toBe('21px');
    expect(barFamily()).toBe('David');
  });

  it('בלי `sampleSize` הפס נשאר בגודלו הקבוע — קורא ישן אינו נשבר', async () => {
    // התיבה של בורר הגודל, וכל קורא שלא העביר את הפרופ, אינם אמורים לקבל
    // `font-size` inline — הגודל הקבוע שב-CSS נשאר.
    await combo.setProps({ options: SAMPLED, modelValue: 'David', sample: 'שלום' });
    await open();
    expect(barSize()).toBe('');
  });

  /**
   * שני מצבי הפס. ה-CSS הוא שנותן לאחד שורה אחת ולשני שתיים, ו-jsdom אינו
   * מחשב אותו; מה שנבדק כאן הוא ההכרעה שמפעילה אותו — ושהיא נגזרת מ-
   * `sampleSize` ולא מדגל שני שיכול לסתור אותו.
   */
  describe('מצב הדגימה מול מצב הבחירה', () => {
    const listClasses = () => combo.find('[role="listbox"]').classes();

    it('בלי גודל — הפס במצב דגימה, והרשימה יודעת שהוא גבוה יותר', async () => {
      await combo.setProps({ options: SAMPLED, modelValue: 'David', sample: 'פסוק' });
      await open();
      expect(bar().classes()).toContain('specimen');
      expect(listClasses()).toContain('specimen');
    });

    it('עם גודל שבמסמך — שורה אחת, בלי מצב דגימה', async () => {
      await combo.setProps({
        options: SAMPLED,
        modelValue: 'David',
        sample: 'שלום',
        sampleSize: '21px',
      });
      await open();
      expect(bar().classes()).not.toContain('specimen');
      expect(listClasses()).not.toContain('specimen');
    });

    it('הטקסט יושב ב-`<span>` פנימי — בלעדיו אין קטיעה בשלוש שורות', async () => {
      // `-webkit-line-clamp` דורש `display: -webkit-box` על האלמנט שנקטע,
      // ו-`text-overflow` של מצב הבחירה חל על מכל בלוק בלבד. הצומת הפנימי
      // אינו קישוט: בלעדיו שורה רביעית הייתה גולשת מהפס בגופן רחב, בלי שום
      // סימן ב-jsdom.
      await combo.setProps({ options: SAMPLED, modelValue: 'David', sample: 'פסוק' });
      await open();
      expect(bar().find('span').text()).toBe('פסוק');
    });

    it('בלי פס אין מחלקה על הרשימה — בורר הגודל אינו מקבל גובה של דגימה', async () => {
      await combo.setProps({ options: SAMPLED, modelValue: 'David', sample: '' });
      await open();
      expect(bar().exists()).toBe(false);
      expect(listClasses()).not.toContain('specimen');
    });
  });
});

describe('נגישות', () => {
  it('התיבה מכריזה על עצמה כ-combobox ומצביעה על הרשימה', async () => {
    const input = combo.find('input');
    expect(input.attributes('role')).toBe('combobox');
    expect(input.attributes('aria-expanded')).toBe('false');

    await open();
    expect(combo.find('input').attributes('aria-expanded')).toBe('true');
    expect(combo.find('input').attributes('aria-controls')).toBe(
      combo.find('[role="listbox"]').attributes('id'),
    );
  });

  it('כפתור החץ נפתח גם בהפעלה שאינה מעכבר', async () => {
    // `pointerdown` לבדו נראה נכון עד שמפעילים אחרת: `click()` תכנותי, והפעלה
    // במקלדת, אינם מייצרים `pointerdown` כלל — כלומר כפתור מת. `detail === 0`
    // הוא מה שמפריד ביניהם ללחיצת עכבר, שכבר טופלה.
    // `element.click()` ולא `trigger`: זו בדיוק ההפעלה שאינה מעכבר — היא
    // מייצרת `detail === 0`, ו-`trigger` אינו יכול לקבוע את השדה הזה.
    (combo.find('.ribbon-combo-arrow').element as HTMLElement).click();
    await nextTick();
    expect(combo.find('[role="listbox"]').exists()).toBe(true);
  });

  it('לחיצת עכבר על החץ אינה נסגרת מיד אחרי שנפתחה', async () => {
    // הרצף האמיתי הוא `pointerdown` ואז `click`. בלי ההבחנה השני היה מבטל את
    // הראשון, והרשימה הייתה מהבהבת במקום להיפתח.
    const arrow = combo.find('.ribbon-combo-arrow').element;
    await combo.find('.ribbon-combo-arrow').trigger('pointerdown');
    arrow.dispatchEvent(new MouseEvent('click', { bubbles: true, detail: 1 }));
    await nextTick();
    expect(combo.find('[role="listbox"]').exists()).toBe(true);
  });

  it('התיבה נושאת `data-tip-title` — טולטיפ הרצועה, ולפיו גם שערי ה-QA מאתרים אותה', () => {
    // `data-tip-title` ולא `title` נייטיב: זו המערכת שכל שאר פקדי הרצועה
    // משתמשים בה (ui/tooltip), ו-`nameOf` ב-`scripts/qa/qa-api.js` קורא אותה
    // ראשונה. פקד בלי אף אחד מהם הוא פקד שאף שער לא יוכל ללחוץ עליו.
    expect(combo.find('input').attributes('data-tip-title')).toBe('גופן');
    expect(combo.find('input').attributes('aria-label')).toBe('גופן');
  });

  it('`aria-activedescendant` מצביע על השורה המסומנת', async () => {
    await open();
    await key('ArrowDown');
    await nextTick();
    expect(combo.find('input').attributes('aria-activedescendant')).toBe(
      combo.find('[role="option"].active').attributes('id'),
    );
  });

  it('הגופן הנבחר מסומן `aria-selected`, ורק הוא', async () => {
    await open();
    const selected = combo
      .findAll('[role="option"]')
      .filter((el) => el.attributes('aria-selected') === 'true')
      .map((el) => el.attributes('data-value'));
    expect(selected).toEqual(['Arial']);
  });
});

/**
 * גופן שהדפדפן אינו פותר — מה שמוצג עליו, ומה ש**אינו** מוצג.
 *
 * הבאג שנמדד (`scripts/qa/font-availability-qa.mjs`, 67 שורות בכרום
 * אמיתי): שורת `Aptos` הכריזה „Aptos” וצוירה ב-Arial, ופס הדגימה הראה את
 * הטקסט של המשתמש בגודל שבמסמך תחת שם של גופן אחר.
 *
 * שני חצאים, והשני הוא זה שנותן לתיקון ערך: הורדת הציור לבדה מותירה את הפס
 * מצייר בגופן הממשק ב-24px, וזה נראה בדיוק כמו דגימה אמיתית. לכן גם הסימון.
 */
describe('גופן שאינו מותקן', () => {
  /*
   * ל-Aptos יש כאן `preview` **וגם** `unavailable`, וזה לא מקרי: בלי ה-
   * `preview` הבדיקות „אינו מצויר” היו עוברות גם על פקד שמתעלם מהדגל לגמרי,
   * כי ממילא לא היה מה לצייר. הצירוף הזה הוא מה שמודד את `paintFamily`.
   *
   * והוא גם המצב האמיתי של המסלול השני: `withCurrent` מוסיף את גופן המסמך עם
   * `preview` מלא, ורק המדידה קובעת אם הוא זמין.
   */
  const WITH_MISSING: readonly ComboOption[] = [
    { value: 'Arial', label: 'Arial', group: '', preview: 'Arial, sans-serif' },
    { value: 'Aptos', label: 'Aptos', group: '', preview: "'Aptos', sans-serif", unavailable: true },
  ];

  const mountMissing = (sample = '') =>
    mount(RibbonCombo, {
      props: { modelValue: 'Arial', options: WITH_MISSING, title: 'גופן', sample },
    });

  const rowFor = (wrapper: VueWrapper, value: string) =>
    wrapper.findAll('[role="option"]').find((el) => el.attributes('data-value') === value);

  it('השורה מסומנת, ומסבירה למי שעוצר עליה', async () => {
    const wrapper = mountMissing();
    await wrapper.find('input').trigger('focus');
    await nextTick();

    const row = rowFor(wrapper, 'Aptos');
    expect(row?.classes()).toContain('unavailable');
    expect(row?.attributes('data-tip-title')).toBe('הגופן אינו מותקן במכונה — אין דגימה להציג');
  });

  /**
   * גופן ש**כן** מותקן ושהדפדפן אינו מצייר — הסבר אחר, ולא אותו אחד.
   *
   * זו התלונה שהגיעה מהשטח: המשתמש התקין את הגופנים שנעדרו מהבורר, ולכן
   * „הגופן אינו מותקן במכונה” היה הופך את השורה ממחוסרת-תועלת למבלבלת. נמדד
   * ש-43 מ-287 המשפחות שהמנייה מדווחת הן מהסוג הזה, חמש מהן עבריות.
   */
  describe('גופן מותקן שהדפדפן אינו מצייר', () => {
    const WITH_UNDRAWABLE: readonly ComboOption[] = [
      { value: 'Arial', label: 'Arial', group: '', preview: 'Arial, sans-serif' },
      {
        value: 'Guttman Kav-Light',
        label: 'Guttman Kav-Light',
        group: '',
        preview: "'Guttman Kav-Light', serif",
        unavailable: true,
        installedNotDrawable: true,
      },
    ];

    const openWith = async (options: readonly ComboOption[]) => {
      const wrapper = mount(RibbonCombo, {
        props: { modelValue: 'Arial', options, title: 'גופן', sample: '' },
      });
      await wrapper.find('input').trigger('focus');
      await nextTick();
      return wrapper;
    };

    it('ההסבר אומר „מותקן” ולא „אינו מותקן”', async () => {
      const wrapper = await openWith(WITH_UNDRAWABLE);
      const row = rowFor(wrapper, 'Guttman Kav-Light');

      expect(row?.attributes('data-tip-title')).toBe(
        'הגופן מותקן אך הדפדפן אינו מצייר אותו — אין דגימה להציג',
      );
      // ועדיין מסומנת: אין דגימה אמיתית להציג, וזה לא השתנה.
      expect(row?.classes()).toContain('unavailable');
    });

    it('גם ה-aria-label אומר את זה — קורא מסך אינו רואה את הטולטיפ', async () => {
      const wrapper = await openWith(WITH_UNDRAWABLE);
      expect(rowFor(wrapper, 'Guttman Kav-Light')?.attributes('aria-label')).toBe(
        'Guttman Kav-Light — הגופן מותקן אך הדפדפן אינו מצייר אותו — אין דגימה להציג',
      );
    });

    it('ובלי הדגל חוזר ההסבר של „אינו מותקן” — שני מצבים ולא אחד', async () => {
      const wrapper = await openWith(
        WITH_UNDRAWABLE.map((option) =>
          option.value === 'Guttman Kav-Light'
            ? { ...option, installedNotDrawable: undefined }
            : option,
        ),
      );
      expect(rowFor(wrapper, 'Guttman Kav-Light')?.attributes('data-tip-title')).toBe(
        'הגופן אינו מותקן במכונה — אין דגימה להציג',
      );
    });
  });

  it('והיא אינה מצוירת בגופן שאינו קיים — זה השקר עצמו', async () => {
    const wrapper = mountMissing();
    await wrapper.find('input').trigger('focus');
    await nextTick();

    // `style` חסר לגמרי, ולא `font-family` ריק: `previewFamily` נופל במיזוג
    // (engine/font-options.ts), ולכן הפקד אינו מקבל מה לצייר בו מלכתחילה.
    expect(rowFor(wrapper, 'Aptos')?.attributes('style')).toBeUndefined();
  });

  it('שורה שכן נפתרת אינה נוגעת — לא סימון ולא title', async () => {
    const wrapper = mountMissing();
    await wrapper.find('input').trigger('focus');
    await nextTick();

    const row = rowFor(wrapper, 'Arial');
    expect(row?.classes()).not.toContain('unavailable');
    expect(row?.attributes('data-tip-title')).toBeUndefined();
    expect(row?.attributes('style')).toContain('Arial');
  });

  it('פס הדגימה מעומעם כשהסימון עומד על גופן שאינו מותקן', async () => {
    const wrapper = mountMissing('שלום עולם');
    await wrapper.find('input').trigger('focus');
    await nextTick();

    // נפתח על Arial, ולכן הפס נקי; חץ אחד מעביר ל-Aptos.
    expect(wrapper.find('.ribbon-combo-sample').classes()).not.toContain('unavailable');

    await wrapper.find('input').trigger('keydown', { key: 'ArrowDown' });
    await nextTick();

    expect(wrapper.find('[role="option"].active').attributes('data-value')).toBe('Aptos');
    expect(wrapper.find('.ribbon-combo-sample').classes()).toContain('unavailable');
  });

  it('והפס מצייר בגודל שבמסמך אבל לא בגופן שאינו קיים', async () => {
    /*
     * הטענה חייבת להיות **חיובית** כדי שתוכל להיכשל. בלי `sampleSize` הסגנון
     * של הפס יוצא `undefined`, ואז „אינו מכיל Aptos” נכון גם על פקד ריק ועל
     * פקד שמתעלם מהדגל לגמרי. הגודל הוא מה שמוכיח שהסגנון אכן נבנה, ושדווקא
     * הגופן הוא שנשמט ממנו.
     */
    const wrapper = mount(RibbonCombo, {
      props: { modelValue: 'Arial', options: WITH_MISSING, title: 'גופן', sample: 'שלום עולם', sampleSize: '24px' },
    });
    await wrapper.find('input').trigger('focus');
    await nextTick();
    await wrapper.find('input').trigger('keydown', { key: 'ArrowDown' });
    await nextTick();

    const style = wrapper.find('.ribbon-combo-sample').attributes('style') ?? '';
    expect(style).toContain('24px');
    expect(style).not.toContain('Aptos');
  });

  it('התיבה הסגורה אינה מצוירת בגופן שאינו מותקן — מקום הציור השלישי', async () => {
    const wrapper = mount(RibbonCombo, {
      props: { modelValue: 'Aptos', options: WITH_MISSING, title: 'גופן' },
    });
    expect(wrapper.find('input').attributes('style')).toBeUndefined();

    await wrapper.setProps({ modelValue: 'Arial' });
    expect(wrapper.find('input').attributes('style')).toContain('Arial');
  });

  it('שורה שאינה מותקנת עדיין נבחרת — מסמך נודד, וזו בחירה חוקית', async () => {
    // ה-CSS מנמק באורך למה **לא** `disabled`. בלי בדיקה, `aria-disabled` או
    // שומר ב-`choose` היו עוברים בשקט והופכים את הנימוק לשקר.
    const wrapper = mountMissing();
    await wrapper.find('input').trigger('focus');
    await nextTick();

    await rowFor(wrapper, 'Aptos')?.trigger('pointerdown');
    await nextTick();

    expect(wrapper.emitted('update:modelValue')?.[0]).toEqual(['Aptos']);
  });

  it('השם הנגיש אומר שאין דגימה — העמעום והטולטיפ אינם מגיעים לקורא מסך', async () => {
    const options: readonly ComboOption[] = [
      ...WITH_MISSING,
      { value: 'Narkisim', label: 'Narkisim', group: 'עברית', preview: 'Narkisim', hebrew: true },
    ];
    const wrapper = mount(RibbonCombo, { props: { modelValue: 'Arial', options, title: 'גופן' } });
    await wrapper.find('input').trigger('focus');
    await nextTick();

    expect(rowFor(wrapper, 'Aptos')?.attributes('aria-label')).toBe(
      'Aptos — הגופן אינו מותקן במכונה — אין דגימה להציג',
    );
    // שורה עברית זמינה מקבלת את שמה בלבד — זה הכלל שהיה כאן קודם, והוא נשמר.
    expect(rowFor(wrapper, 'Narkisim')?.attributes('aria-label')).toBe('Narkisim');
    // ושורה רגילה אינה מקבלת תווית, כדי שהתוכן שלה יוכרז כמו שהוא.
    expect(rowFor(wrapper, 'Arial')?.attributes('aria-label')).toBeUndefined();
  });

  it('חיפוש אינו מאבד את הסימון', async () => {
    const wrapper = mountMissing();
    const input = wrapper.find('input');
    await input.trigger('focus');
    (input.element as HTMLInputElement).value = 'apt';
    await input.trigger('input');
    await nextTick();

    const rows = wrapper.findAll('[role="option"]');
    expect(rows).toHaveLength(1);
    expect(rows[0].classes()).toContain('unavailable');
    expect(rows[0].attributes('data-tip-title')).toBeTruthy();
  });
});

describe('background font verification', () => {
  it.each([true, false])('refreshes an open current-only family when availability becomes %s', async (available) => {
    const current = 'CurrentOnlyFont';
    const pending = withCurrent([], current, { preview: true, available: () => undefined });
    const verified = withCurrent([], current, { preview: true, available: () => available });
    // היצרן משמיט pending כשהתשובה ידועה; הוא אינו כותב pending:false.
    expect(verified[0]).not.toHaveProperty('pending');
    await combo.setProps({ modelValue: current, options: [...pending, ...OPTIONS], sample: 'טקסט לדוגמה' });
    await open();
    const before = optionValues();
    const row = () => combo.find(`[data-value="${current}"]`);
    expect(row().attributes('data-availability')).toBe('pending');
    await combo.setProps({ options: [...OPTIONS, ...verified] });
    expect(optionValues()).toEqual(before);
    expect(row().attributes('data-availability')).toBe(available ? 'measured' : 'missing');
    expect(row().classes()).not.toContain('pending');
    expect(row().classes().includes('unavailable')).toBe(!available);
    expect(combo.find('.ribbon-combo-sample').text()).toBe('טקסט לדוגמה');
    if (available) {
      expect(row().attributes('style')).toContain(`font-family: ${current}`);
    } else {
      expect(row().attributes('style')).toBeUndefined();
      expect(row().attributes('aria-label')).toContain('הגופן אינו מותקן');
    }
    combo.unmount();
  });

  it('updates pending rows and samples while preserving the open list order', async () => {
    const pending: ComboOption = { value: 'New font', label: 'New font', hebrew: true, pending: true, measured: false };
    await combo.setProps({ modelValue: 'New font', options: [pending, ...OPTIONS], sample: 'טקסט לדוגמה' });
    await open();
    const before = optionValues();
    const row = () => combo.find('[data-value="New font"]');
    expect(row().attributes('data-availability')).toBe('pending');
    expect(row().classes()).not.toContain('unavailable');
    expect(row().attributes('style')).toBeUndefined();
    expect(combo.find('.ribbon-combo-sample').text()).toBe('הדגימה טרם נבדקה');
    await combo.setProps({ options: [...OPTIONS, { ...pending, group: 'עברית', pending: false, measured: true, preview: '"New font"' }] });
    expect(optionValues()).toEqual(before);
    expect(row().attributes('data-availability')).toBe('measured');
    expect(row().attributes('style')).toContain('font-family: "New font"');
    expect(combo.find('.ribbon-combo-sample').text()).toBe('טקסט לדוגמה');
    expect(combo.find('.ribbon-combo-sample').attributes('style')).toContain('font-family: "New font"');
    combo.unmount();
  });

  it('prioritizes selected and visible families, and disconnects on close', async () => {
    const prioritize = vi.fn();
    const disconnect = vi.fn();
    const observe = vi.fn();
    let notify!: IntersectionObserverCallback;
    vi.stubGlobal('IntersectionObserver', class {
      constructor(callback: IntersectionObserverCallback) { notify = callback; }
      disconnect = disconnect;
      observe = observe;
    });
    const picker = mount(RibbonCombo, {
      props: { modelValue: 'B', options: ['A', 'B', 'C'].map((value) => ({ value, label: value, pending: true })) },
      global: { provide: { [FONT_PROBE_PRIORITIES as symbol]: prioritize } },
    });
    try {
      expect(prioritize).toHaveBeenCalledWith(['B']);
      await picker.find('input').trigger('focus');
      await nextTick();
      expect(observe).toHaveBeenCalledTimes(3);
      notify([
        { isIntersecting: true, target: picker.find('[data-value="C"]').element },
        { isIntersecting: false, target: picker.find('[data-value="A"]').element },
      ] as unknown as IntersectionObserverEntry[], {} as IntersectionObserver);
      expect(prioritize).toHaveBeenLastCalledWith(['C']);
      await picker.find('input').trigger('keydown', { key: 'Escape' });
      expect(disconnect).toHaveBeenCalledTimes(1);
    } finally {
      picker.unmount();
      vi.unstubAllGlobals();
    }
  });
});
