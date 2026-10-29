/**
 * השער שאם היה קיים, גל התיקונים הזה לא היה נדרש: **כל** כפתור בכל לשונית
 * מורכב, נלחץ, ונמדד — או שהוא מנוטרל, או שהלחיצה עליו עשתה משהו נצפה.
 * „משהו נצפה” הוא אחד מארבעה: פקודה לאדפטר, קריאה ל-Document API, event
 * שנפלט, או שינוי ב-DOM (פופאובר שנפתח). כפתור שנלחץ ואף אחד מהם לא קרה הוא
 * כפתור מת, וזה כשל.
 *
 * למה זה תופס מה שסריקת המקור אינה תופסת: `tests/unit/tab-controls.test.ts`
 * שואל „האם יש `@click`”, וזה כל מה שהוא יכול לשאול. `doCut(){}`,
 * `insertPageBreak(){}` ו-`doSelectAll` שסימן את ממשק האפליקציה במקום את
 * המסמך — כולם עברו את השאלה ההיא בהצלחה מלאה.
 *
 * כל כפתור נמדד בהרכבה **טרייה**: פופאובר שנפתח בלחיצה קודמת היה משנה את מה
 * שהלחיצה הבאה פוגשת, וכשל כזה היה תלוי בסדר.
 *
 * ה-events נספרים דרך `emittedCount` של ה-harness, שמדלג על `click`: VTU רושם
 * ב-`emitted()` גם אירועי DOM שעברו דרך השורש, ולכן `click` מופיע שם גם על
 * כפתור מת. „בדיקת הבקרה” שבסוף הקובץ היא מה שמקבע את ההבחנה הזאת.
 */
import { beforeEach, afterEach, describe, expect, it } from 'vitest';
import type { DOMWrapper } from '@vue/test-utils';
import { defineComponent, h, type Component } from 'vue';
import RibbonButton from '../../src/ui/ribbon/common/RibbonButton.vue';
import HomeTab from '../../src/ui/ribbon/tabs/HomeTab.vue';
import FileTab from '../../src/ui/ribbon/tabs/FileTab.vue';
import InsertTab from '../../src/ui/ribbon/tabs/InsertTab.vue';
import LayoutTab from '../../src/ui/ribbon/tabs/LayoutTab.vue';
import ReferencesTab from '../../src/ui/ribbon/tabs/ReferencesTab.vue';
import ReviewTab from '../../src/ui/ribbon/tabs/ReviewTab.vue';
import ViewTab from '../../src/ui/ribbon/tabs/ViewTab.vue';
import ShulchanTab from '../../src/ui/ribbon/tabs/ShulchanTab.vue';
import DeveloperTab from '../../src/ui/ribbon/tabs/DeveloperTab.vue';
import OtzariaTab from '../../src/ui/ribbon/tabs/OtzariaTab.vue';
import Ribbon from '../../src/ui/ribbon/Ribbon.vue';
import {
  autoUnmount,
  createSuperdocDouble,
  emittedCount,
  installSystemClipboard,
  mountUi,
  settle,
  tipMessage,
} from './harness';

autoUnmount();

/** מסמך עם בחירה חיה: בלעדיה פעולות הלוח נכשלות לפני שהן נוגעות במנוע. */
const withSelection = () =>
  createSuperdocDouble({ selection: { hasRange: true, text: 'טקסט נבחר' } });

/**
 * `props` נדרש ללשונית „קובץ” בלבד: פקדיה הם פעולות מעטפת (מסמך פתוח, שמירה
 * שרצה, פתיחה שרצה) ולא פקודות מנוע, ולכן המצב שלהם מגיע מ-App.vue ולא
 * מהכפילים. ברירת המחדל שלהם היא „אין מסמך”, כלומר הרכבה בלי props הייתה
 * מודדת חצי מהלשונית מנוטרלת — ומאשרת בירוק בדיוק את מה שהשער הזה נבנה לתפוס.
 */
const TABS: ReadonlyArray<{
  name: string;
  component: Component;
  props?: Record<string, unknown>;
}> = [
  // `hasPdfExport` הוא זמינות ה-Host ולא היעדר API: `ui.exportPdf` קיים,
  // והפקד מנוטרל רק כשאוצריא ישנה מ-0.9.97. הכפיל מדמה Host עדכני, אחרת
  // הבדיקה הזאת הייתה מדווחת „אין API” על משהו שיש לו.
  { name: 'קובץ', component: FileTab, props: { hasDocument: true, hasPdfExport: true } },
  { name: 'בית', component: HomeTab },
  { name: 'הוספה', component: InsertTab },
  { name: 'פריסה', component: LayoutTab },
  { name: 'הפניות', component: ReferencesTab },
  { name: 'סקירה', component: ReviewTab },
  { name: 'תצוגה', component: ViewTab },
  { name: 'שולחן העורך', component: ShulchanTab },
  // „מפתחים” נסרקת כאן מאז שפקדי המאקרו ירדו מ„אוצריא”: הם היו נמדדים דרך
  // הלשונית ההיא, וכפילות בין שתי לשוניות אינה כיסוי.
  { name: 'מפתחים', component: DeveloperTab },
];

/** מה שמזהה כפתור בהודעת כשל — כדי שאפשר יהיה למצוא אותו בקובץ. */
function nameOf(button: DOMWrapper<Element>): string {
  return (
    // ההסבר ולא הכותרת, כשיש: בפקד מנוטרל הוא נושא את *הסיבה*, וזה מה
    // שהבדיקות כאן מזהות לפיו („הוספת תגובה — תתווסף בשלב הבא…”).
    tipMessage(button) ||
    button.attributes('aria-label') ||
    button.text().trim() ||
    button.html().slice(0, 70)
  );
}

interface Probe {
  name: string;
  disabled: boolean;
  /** מה שקרה בלחיצה, ריק = כלום. */
  effects: string[];
}

/**
 * מרכיבה את הלשונית, לוחצת על הכפתור ה-index, ומחזירה מה קרה.
 *
 * `count` מוחזר בהרכבה הראשונה כדי שהסוקר ידע כמה כפתורים יש; אין דרך לדעת
 * את זה בלי להרכיב, וספירה קשיחה כאן הייתה מתיישנת בכל פקד שנוסף.
 */
async function probe(
  component: Component,
  index: number,
  props?: Record<string, unknown>,
): Promise<Probe & { count: number }> {
  const harness = mountUi(component, { superdoc: withSelection(), props });
  await settle();

  const buttons = harness.wrapper.findAll('button');
  const button = buttons[index];
  const name = nameOf(button);
  const disabled = button.attributes('disabled') !== undefined;

  const before = {
    commands: harness.adapter.calls.length,
    doc: harness.superdoc.calls.length,
    reports: harness.reports.length,
    emitted: emittedCount(harness.wrapper),
    html: harness.wrapper.html(),
    // דיאלוגים מרונדרים ב-Teleport לגוף הדף ואינם נראים ב-wrapper.html().
    // בלי הצצה ל-body, כפתור שכל מה שהוא עושה הוא פתיחת דיאלוג נמדד כמת.
    bodyHtml: document.body.innerHTML,
  };

  await button.trigger('click');
  await settle();

  const effects: string[] = [];
  if (harness.adapter.calls.length > before.commands) effects.push('פקודה');
  if (harness.superdoc.calls.length > before.doc) effects.push('Document API');
  if (emittedCount(harness.wrapper) > before.emitted) effects.push('event');
  if (harness.reports.length > before.reports) effects.push('דיווח');
  if (harness.wrapper.html() !== before.html) effects.push('DOM');
  if (document.body.innerHTML !== before.bodyHtml) effects.push('DOM-teleport');

  return { name, disabled, effects, count: buttons.length };
}

/**
 * תקרת זמן לבדיקות שרצות דרך `probeAll`. ברירת המחדל של vitest, 5 שניות,
 * מכוונת לבדיקה שעושה דבר אחד — וכאן בדיקה אחת מרכיבה את הלשונית מחדש לכל
 * כפתור, כלומר „בית” היא כשלושים הרכבות של קומפוננטה בת 24KB ב-jsdom.
 *
 * נמדד על המכונה הזאת: 2.2 שניות ל„בית”, כלומר מרווח של פי שניים בלבד מול
 * 5 שניות — ומכונה איטית פי שניים תפיל את הבדיקה על עומס ולא על הקוד. התקרה
 * מורמת כאן, על הבדיקות שסוקרות לשונית שלמה, ולא גלובלית: 53 קובצי הבדיקה
 * האחרים ממשיכים להיות מוגנים ב-5 שניות, וזה מה שתופס בדיקה שנתקעת באמת.
 *
 * ## למה 30 ולא 20
 *
 * „בית” קיבלה תפריט („מרווח שורות וריווח”), ותפריט פתוח הוא תשעה כפתורים
 * נוספים — כלומר תשע הרכבות נוספות של אותה קומפוננטה. נמדד ב-A/B על המכונה
 * הזאת בזמן שהיא עמוסה: **10.7 שניות בלי הפקד, 19.0 אתו**. זה עבר בתקרה של
 * 20 והפיל אותה בריצה המלאה, כלומר הפער נבלע בעומס ולא בקוד — בדיוק המצב
 * שהפסקה שמעל מתארת. היחס נשמר: התקרה גדלה באותו שיעור שהבדיקה גדלה בו.
 */
const PROBE_TIMEOUT = 30_000;

/** סוקרת את כל הכפתורים בלשונית, כל אחד בהרכבה נפרדת. */
async function probeAll(
  component: Component,
  props?: Record<string, unknown>,
): Promise<Probe[]> {
  const first = await probe(component, 0, props);
  const probes: Probe[] = [first];
  for (let index = 1; index < first.count; index += 1) {
    probes.push(await probe(component, index, props));
  }
  return probes;
}

let restoreClipboard: () => void;

beforeEach(() => {
  restoreClipboard = installSystemClipboard();
});

afterEach(() => {
  restoreClipboard();
});

describe('אין כפתור מת באף לשונית', () => {
  for (const tab of TABS) {
    it(`„${tab.name}”: כל כפתור מנוטרל, או שלחיצה עליו עושה משהו`, { timeout: PROBE_TIMEOUT }, async () => {
      const probes = await probeAll(tab.component, tab.props);

      expect(probes.length, 'נמצאו כפתורים לבדוק').toBeGreaterThan(0);

      const dead = probes
        .filter((item) => !item.disabled && item.effects.length === 0)
        .map((item) => item.name);
      expect(dead, `כפתורים שנלחצו ולא קרה כלום ב„${tab.name}”`).toEqual([]);
    });
  }
});

describe('הפקדים שמנוטרלים בכוונה', () => {
  /**
   * מנוע עם כל היכולות, בחירה חיה ולוח מערכת: כל מה שנשאר מנוטרל כאן מנוטרל
   * **בכוונה**, ולא מפני שהיכולת חסרה. הרשימה היא לכן חוזה: פקד שנעלם ממנה
   * הוא פקד שהופעל, ופקד שנוסף אליה הוא פקד שהושתק — ובשני המקרים זו החלטה
   * שצריכה להיראות בבדיקה.
   */
  const EXPECTED_DISABLED: Record<string, readonly string[]> = {
    // „כתב תחתי” ו„כתב עליון” היו כאן, מנוטרלים קשיח עם tooltip שהאשים את
    // המנוע. הם עברו ל-Document API (`format.vertAlign`) והם חיים מעכשיו —
    // ולכן הרשימה של „בית” ריקה.
    בית: [],
    // „בדיקת איות” הייתה כאן, מנוטרלת עם tooltip שהודה שאין מילון. המילון
    // נכנס (issue #25) והיא מתג חי מעכשיו — ולכן נשארה רשומה אחת.
    סקירה: ['הוספת תגובה — תתווסף בשלב הבא, יחד עם זהות המחבר ופאנל התגובות'],
    קובץ: [],
    // „קשר לקודם” אינו פקד מת: ה-API שלו קיים והוא נדלק ברגע שיש במסמך מקטע
    // שני. הוא מנוטרל כאן מפני שהמסמך שהכפיל מציג הוא בעל מקטע יחיד, ולמקטע
    // הראשון אין קודם לקשר אליו — בדיוק כמו ב-Word. זו הרשומה היחידה ברשימה
    // שסיבתה מצב המסמך ולא היעדר API, ולכן היא מנומקת כאן במפורש.
    הוספה: ['אין במסמך מקטע נוסף — הקישור נוגע רק במקטעים שאחרי הראשון'],
    פריסה: [],
    הפניות: [],
    תצוגה: [],
    'שולחן העורך': [],
    // המאקרו אינם תלויים ב-SDK של אוצריא אלא במסמך פתוח בלבד, וכאן יש מסמך.
    מפתחים: [],
  };

  for (const tab of TABS) {
    it(`„${tab.name}”: רק הפקדים שאין להם API נשארים מנוטרלים`, { timeout: PROBE_TIMEOUT }, async () => {
      const probes = await probeAll(tab.component, tab.props);
      const disabled = probes.filter((item) => item.disabled).map((item) => item.name);
      expect(disabled).toEqual(EXPECTED_DISABLED[tab.name]);
    });
  }
});

describe('לשונית „קובץ” נשענת על מצב המעטפת', () => {
  /**
   * זו הלשונית האחרונה שהיו בה שבעה פקדים ואפס `:disabled`. התנאי כאן שונה
   * מכל שאר הלשוניות ולכן הוא נמדד בנפרד: הפקדים אינם פקודות מנוע אלא פעולות
   * מעטפת, ומה שקובע הוא שלושה מצבים שרק App.vue מחזיק. לפני התיקון „שמור”
   * נראה זמין על מסמך שאינו קיים, ו„פתח קובץ” נראה זמין בזמן שהשמירה רצה —
   * ובשני המקרים הלחיצה נבלעה בלי הודעה.
   */
  const LABELS = [
    'מסמך חדש',
    'פתח קובץ',
    'שמור',
    'שמור בשם...',
    'ייצוא ל-PDF',
    'הדפסה',
    'יציאה',
    'אודות',
    'קיצורים',
    'ניהול קיצורים',
    'דיווח על באג',
  ] as const;

  /** תווית → האם הפקד מנוטרל. התווית היא הטקסט שהמשתמש רואה על הכפתור. */
  async function states(props: Record<string, unknown>): Promise<Record<string, boolean>> {
    const harness = mountUi(FileTab, { superdoc: withSelection(), props });
    await settle();

    const byLabel: Record<string, boolean> = {};
    for (const button of harness.wrapper.findAll('button')) {
      byLabel[button.text().trim()] = button.attributes('disabled') !== undefined;
    }
    return byLabel;
  }

  it('אחד-עשר פקדים, וכל התוויות נמצאו — אחרת הבדיקות למטה מודדות אוויר', async () => {
    const byLabel = await states({ hasDocument: true });

    expect(Object.keys(byLabel)).toHaveLength(LABELS.length);
    for (const label of LABELS) expect(byLabel[label], label).toBeDefined();
  });

  it('אין מסמך פתוח: שמירה, ייצוא והדפסה מנוטרלים; פתיחה ו„אודות” לא', async () => {
    const byLabel = await states({ hasDocument: false });

    expect(byLabel['שמור']).toBe(true);
    expect(byLabel['שמור בשם...']).toBe(true);
    expect(byLabel['הדפסה']).toBe(true);
    // „מסמך חדש” ו„פתח קובץ” הם בדיוק מה שעושים כשאין מסמך.
    expect(byLabel['מסמך חדש']).toBe(false);
    expect(byLabel['פתח קובץ']).toBe(false);
    expect(byLabel['אודות']).toBe(false);
    // „יציאה” אינו דורש מסמך: יציאה ממסך ריק היא בקשה תקפה.
    expect(byLabel['יציאה']).toBe(false);
  });

  it('שמירה שרצה: אין שמירה נוספת ואין מעבר מסמך', async () => {
    // `decideDocumentSwitch` מחזיר cancel עם reason 'saving', ומפעיל הפעולות
    // של הקיצורים חוסם את Ctrl+S. הפקד מראה את זה מראש במקום לבלוע לחיצה.
    // `hasPdfExport` דלוק כאן ולא בשאר הבדיקות: בלעדיו „ייצוא ל-PDF” מנוטרל
    // מטעם אחר לגמרי (אין `ui.exportPdf` ב-Host), והטענה שהשמירה אינה חוסמת
    // אותו לא הייתה נמדדת בכלל.
    const byLabel = await states({ hasDocument: true, isSaving: true, hasPdfExport: true });

    expect(byLabel['שמור']).toBe(true);
    expect(byLabel['שמור בשם...']).toBe(true);
    expect(byLabel['מסמך חדש']).toBe(true);
    expect(byLabel['פתח קובץ']).toBe(true);
    // ייצוא והדפסה קוראים את המסמך ואינם מתנגשים בשמירה.
    expect(byLabel['ייצוא ל-PDF']).toBe(false);
    expect(byLabel['הדפסה']).toBe(false);
    // „יציאה” כן: הוא שואל „לשמור לפני יציאה?”, ובזמן סבב שמירה השאלה הזאת
    // הייתה מציעה לשמור שוב את מה שנשמר כרגע.
    expect(byLabel['יציאה']).toBe(true);
  });

  it('פתיחה שרצה: אין מעבר מסמך נוסף', async () => {
    const byLabel = await states({ hasDocument: true, isOpening: true });

    expect(byLabel['מסמך חדש']).toBe(true);
    expect(byLabel['פתח קובץ']).toBe(true);
    expect(byLabel['שמור']).toBe(false);
  });

  it('ה-tooltip של פקד מנוטרל אומר למה, ולא חוזר על התווית', async () => {
    const harness = mountUi(FileTab, { superdoc: withSelection(), props: { hasDocument: false } });
    await settle();

    const save = harness.wrapper.findAll('button').find((b) => b.text().trim() === 'שמור');
    expect(tipMessage(save!)).toContain('אין מסמך פתוח');
  });

  it('„יציאה” פולט exit-app, וההודעה עוברת דרך הרצועה', async () => {
    // בלי המסלול הזה הכפתור היה כפתור מת: הוא נראה, הוא נלחץ, ואף אחד לא
    // מקשיב. שני השלבים נמדדים — הלשונית פולטת, והרצועה מעבירה הלאה.
    const tab = mountUi(FileTab, { superdoc: withSelection(), props: { hasDocument: true } });
    await settle();
    const exitButton = tab.wrapper
      .findAll('button')
      .find((button) => button.text().trim() === 'יציאה');
    await exitButton!.trigger('click');
    expect(tab.wrapper.emitted('exit-app')).toHaveLength(1);

    const ribbon = mountUi(Ribbon, { props: { hasDocument: true } });
    await settle();
    const fileTab = ribbon.wrapper.findAll('[role="tab"]').find((t) => t.text() === 'קובץ');
    await fileTab!.trigger('click');
    await settle();
    const inRibbon = ribbon.wrapper
      .findAll('button')
      .find((button) => button.text().trim() === 'יציאה');
    await inRibbon!.trigger('click');
    expect(ribbon.wrapper.emitted('exit-app'), 'הרצועה לא העבירה את exit-app').toHaveLength(1);
  });

  it('ה-tooltip של „יציאה” אומר שהמסמך נסגר', async () => {
    // „יציאה” מלשונית בתוך אוצריא אינו מובן מאליו: `navigation.goTo` משהה את
    // ה-WebView ואינו הורס אותו, ולכן היציאה סוגרת את המסמכים בעצמה (`onExit`
    // ב-App.vue). כל עוד ה-tooltip הבטיח „המסמך יישאר פתוח” הוא סתר את השאלה
    // שהלחיצה שואלת („לצאת בלי לשמור?”), וזו הטענה שנמדדת כאן.
    const harness = mountUi(FileTab, { superdoc: withSelection(), props: { hasDocument: true } });
    await settle();

    const exitButton = harness.wrapper
      .findAll('button')
      .find((button) => button.text().trim() === 'יציאה');
    expect(tipMessage(exitButton!)).toContain('סגירת המסמך');
  });

  it('„יציאה” מנוטרל בזמן פתיחה, ואומר למה', async () => {
    // סגירה באמצע פתיחה משאירה את `openDocumentInto` כותב לתוך טאב מפורק —
    // `onExit` חוסם אותה ב-`isOpenBusy`, וכפתור שנלחץ ולא קורה בו דבר הוא
    // בדיוק מה שהלשונית הזאת נמנעת ממנו בשאר הפקדים.
    const harness = mountUi(FileTab, {
      superdoc: withSelection(),
      props: { hasDocument: true, isOpening: true },
    });
    await settle();

    const exitButton = harness.wrapper
      .findAll('button')
      .find((button) => button.text().trim() === 'יציאה');
    expect(exitButton!.attributes('disabled')).not.toBeUndefined();
    expect(tipMessage(exitButton!)).toContain('פתיחת מסמך רצה כרגע');
  });

  it('„יציאה” מנוטרל בזמן שסגירה קודמת רצה', async () => {
    const harness = mountUi(FileTab, {
      superdoc: withSelection(),
      props: { hasDocument: true, isExiting: true },
    });
    await settle();

    const exitButton = harness.wrapper
      .findAll('button')
      .find((button) => button.text().trim() === 'יציאה');
    expect(exitButton!.attributes('disabled')).not.toBeUndefined();
    expect(tipMessage(exitButton!)).toContain('סגירת המסמכים רצה כרגע');
  });

  it('הרצועה מעבירה את שלושת המצבים — אחרת ה-props כאן הם קוד מת', async () => {
    const harness = mountUi(Ribbon, { props: { hasDocument: false } });
    await settle();
    const fileTab = harness.wrapper
      .findAll('[role="tab"]')
      .find((tab) => tab.text() === 'קובץ');
    await fileTab!.trigger('click');
    await settle();

    const save = harness.wrapper.findAll('button').find((b) => b.text().trim() === 'שמור');
    expect(save?.attributes('disabled'), 'המצב לא עבר דרך הרצועה').not.toBeUndefined();
  });
});

describe('„אוצריא”', () => {
  /**
   * שלושת הפקדים תלויים ב-SDK, וההרכבה שלה בלעדיו מודדת רק את הכיבוי. לכן
   * ה-SDK מותקן — כפיל שמחזיר כשל על כל קריאה, כי מה שנמדד כאן הוא שהפקד
   * **מגיע** אליו.
   */
  beforeEach(() => {
    Reflect.set(window, 'Otzaria', {
      call: async () => ({ success: false, error: { message: 'לא זמין בבדיקה' } }),
      on: () => {},
      off: () => {},
    });
  });

  afterEach(() => {
    Reflect.deleteProperty(window, 'Otzaria');
  });

  it('בתוך אוצריא: אין בלשונית פקד מת, ואין בה פקד מנוטרל לצמיתות', async () => {
    const probes = await probeAll(OtzariaTab);

    const dead = probes
      .filter((item) => !item.disabled && item.effects.length === 0)
      .map((item) => item.name);
    expect(dead).toEqual([]);

    // „סגנון תורני” — שלושה כפתורים מנוטרלים, שירדו מהלשונית כפקד שאינו
    // נצרך (הבקשה למנוע נשארה פתוחה: superdoc/docx-editor#3975). בלעדיהם
    // אין כאן פקד שהניטרול שלו קבוע: כל מה שמנוטרל, מנוטרל מפני שאין SDK
    // או שאין מסמך.
    const permanentlyDisabled = probes.filter((item) =>
      item.name.includes('סגנונות תורניים'),
    );
    expect(permanentlyDisabled).toEqual([]);
  });

  it.each([false, true])('אפשר לבטל אתחול הכתבה מהכפתור גם כשהמסמך ננעל: %s', async (dictationReadOnly) => {
    const { wrapper } = mountUi(OtzariaTab, {
      superdoc: withSelection(),
      props: { dictationState: 'starting', dictationReadOnly },
    });
    await settle();
    const button = wrapper.get('.ribbon-dictation');
    expect(button.attributes('disabled')).toBeUndefined();
    expect(button.attributes('aria-busy')).toBe('true');
    expect(button.text()).toBe('בטל הקלטה');
    expect(tipMessage(button)).toContain('ביטול התחלת ההקלטה');
    await button.trigger('click');
    expect(wrapper.emitted('toggle-dictation')).toHaveLength(1);
  });

  it('מסמך לקריאה בלבד עדיין אינו מאפשר להתחיל הכתבה', async () => {
    const { wrapper } = mountUi(OtzariaTab, {
      superdoc: withSelection(),
      props: { dictationState: 'idle', dictationReadOnly: true },
    });
    await settle();
    const button = wrapper.get('.ribbon-dictation');
    expect(button.attributes('disabled')).toBeDefined();
    await button.trigger('click');
    expect(wrapper.emitted('toggle-dictation')).toBeUndefined();
  });

  it('מחוץ לאוצריא: פקדי ה-SDK מנוטרלים, והייצוא — שאינו תלוי בו — נשאר חי', async () => {
    Reflect.deleteProperty(window, 'Otzaria');
    const harness = mountUi(OtzariaTab, { superdoc: withSelection() });
    await settle();

    const live = harness.wrapper
      .findAll('button')
      .filter((button) => button.attributes('disabled') === undefined)
      .map((button) => button.text().trim());
    // פקדי הציטוט/חיפוש/ספרייה מנוטרלים — הם בדיוק מה שאין בלי ה-SDK.
    // „ייצוא לאוצריא” נשאר חי מחוץ לאוצריא: מסלול השמירה ממומש ב-dev-stub,
    // וכך הוא נבדק בדפדפן — כמו „שמור בשם” בלשונית „קובץ”.
    // המאקרו — שאינו תלוי ב-SDK ולכן נשאר חי גם כאן — עברו ל„מפתחים”,
    // ונמדדים שם.
    expect(live).toEqual(['ייצוא לאוצריא']);
  });
});

describe('„מפתחים”', () => {
  /**
   * שלושת פקדי המאקרו ישבו ב„אוצריא”, וזה מה שנמדד כאן אחרי שירדו משם:
   * מערכת המאקרו רצה כולה בעורך ואינה נוגעת ב-SDK של אוצריא, ולכן הם חיים
   * גם מחוץ לאוצריא — ומנוטרלים בלי מסמך, כי אין מה להקליט.
   */
  it('בלי ה-SDK של אוצריא: שלושת פקדי המאקרו חיים', async () => {
    const harness = mountUi(DeveloperTab, { superdoc: withSelection() });
    await settle();

    const live = harness.wrapper
      .findAll('button')
      .filter((button) => button.attributes('disabled') === undefined)
      .map((button) => button.text().trim());
    expect(live).toEqual(['ניהול מאקרו', 'הקלט מאקרו', 'נגן אחרון']);
  });

  it('בלי מסמך: שלושתם מנוטרלים, ולא נכשלים בלחיצה', async () => {
    const harness = mountUi(DeveloperTab, { superdoc: null });
    await settle();

    const buttons = harness.wrapper.findAll('button');
    expect(buttons).toHaveLength(3);
    for (const button of buttons) {
      expect(button.attributes('disabled'), button.text().trim()).not.toBeUndefined();
      expect(tipMessage(button)).toBe('יש לפתוח מסמך תחילה');
    }
  });
});

describe('בדיקת הבקרה של השער', () => {
  /**
   * לשונית מלאכותית עם שני כפתורים: אחד שהלחיצה עליו נשמעת, ואחד בלי מטפל
   * בכלל — בדיוק שלוש-עשרה הכפתורים שהגל הזה תיקן. אם השער אינו מסמן את השני
   * כמת ואת הראשון כחי, הוא אינו מודד כלום, וכל הבדיקות שלמעלה עוברות מהסיבה
   * הלא נכונה.
   */
  const TabWithDeadButton = defineComponent({
    name: 'TabWithDeadButton',
    emits: ['acted'],
    setup(_props, { emit }) {
      return () =>
        h('div', [
          h(RibbonButton, { label: 'חי', icon: 'bold', onClick: () => emit('acted') }),
          h(RibbonButton, { label: 'מת', icon: 'italic' }),
        ]);
    },
  });

  it('כפתור בלי מטפל מסומן כמת, וכפתור שנשמע אינו', async () => {
    const probes = await probeAll(TabWithDeadButton);
    expect(probes).toHaveLength(2);

    expect(probes[0].effects, 'הכפתור שהלחיצה עליו נשמעת').toContain('event');
    expect(probes[1].effects, 'הכפתור שאין לו מטפל').toEqual([]);
    expect(probes[1].disabled, 'והוא אינו מנוטרל — כלומר נראה עובד').toBe(false);
  });
});
