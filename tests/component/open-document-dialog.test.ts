/**
 * דיאלוג „פתח מסמך”, כפי שהמשתמש פוגש אותו.
 *
 * מה שנמדד כאן ולא במקום אחר:
 *
 * 1. **הציור של התבניות.** חמש התצוגות המקדימות הן SVG שנבנה בקוד מדגלי
 *    `TemplatePreview`, ולכן דגל שיישבר — או קצב שורות שישתנה — לא ייתפס בשום
 *    בדיקת מקור. ספירת ה-`<rect>` וה-`<line>` לכל תבנית היא הטבלה ב-§5.5 של
 *    מפרט העיצוב, ואם הציור משתנה בכוונה יש לעדכן גם אותה.
 * 2. **שהאירועים נושאים את ה-token/id הנכון.** החלפה בין `open-recent`
 *    ל-`forget-recent`, או בין שני כרטיסים, היא בדיוק סוג הבאג שסריקת מקור
 *    מאשרת בירוק.
 * 3. **שני מצבי הריק אינם אותו מצב.** „אין אחרונים בכלל” ו„הסינון לא מצא”
 *    נראים דומה ומתנהגים אחרת — בראשון תיבת החיפוש אינה מרונדרת כלל.
 */
import { describe, expect, it } from 'vitest';
import OpenDocumentDialog from '../../src/ui/panels/OpenDocumentDialog.vue';
import { DOCUMENT_TEMPLATES } from '../../src/engine/templates';
import type { RecentDocument } from '../../src/sessions/recent-documents';
import { autoUnmount, mountUi, settle } from './harness';

autoUnmount();

/** שעה לפני עכשיו — `draftAgeLabel` מחזירה „לפני שעה”, ולא מספר שמשתנה בין ריצות. */
const HOUR = 3_600_000;

function recent(over: Partial<RecentDocument> & { token: string }): RecentDocument {
  return {
    name: `${over.token}.docx`,
    size: 2048,
    openedAt: Date.now() - HOUR,
    writable: true,
    pinned: false,
    ...over,
  };
}

const RECENTS: RecentDocument[] = [
  recent({ token: 'alef', name: 'בראשית.docx' }),
  recent({ token: 'bet', name: 'שמות.docx', pinned: true }),
  recent({ token: 'gimel', name: 'ויקרא.docx', openedAt: 0, size: 0 }),
];

function open(props: Record<string, unknown> = {}) {
  return mountUi(OpenDocumentDialog, {
    props: {
      isOpen: true,
      templates: DOCUMENT_TEMPLATES,
      recents: RECENTS,
      busy: false,
      searchQuery: '',
      ...props,
    },
  });
}

describe('OpenDocumentDialog — מעטפת', () => {
  it('היסטוריה שעדיין נטענת אינה מוצגת כריקה ואינה חוסמת עיון בקבצים', async () => {
    const { wrapper } = open({ recents: [], recentsLoading: true, discardedLoading: true });
    expect(wrapper.find('.rec-section').attributes('aria-busy')).toBe('true');
    expect(wrapper.text()).toContain('טוען מסמכים אחרונים');
    expect(wrapper.text()).not.toContain('עדיין אין מסמכים אחרונים');
    expect(wrapper.find('.open-discarded').attributes('disabled')).toBeDefined();
    expect(wrapper.find('.open-browse').attributes('disabled')).toBeUndefined();
    await wrapper.setProps({ recentsLoading: false, discardedLoading: false });
    expect(wrapper.text()).toContain('עדיין אין מסמכים אחרונים');
    expect(wrapper.find('.open-discarded').exists()).toBe(false);
  });

  it('סגור אינו מרונדר בכלל', () => {
    const { wrapper } = mountUi(OpenDocumentDialog, { props: { isOpen: false } });
    expect(wrapper.find('.open-dialog').exists()).toBe(false);
  });

  it('פתוח מכריז על עצמו כמודאל עם שם נגיש', () => {
    const { wrapper } = open();
    const dialog = wrapper.find('.open-dialog');
    expect(dialog.attributes('role')).toBe('dialog');
    expect(dialog.attributes('aria-modal')).toBe('true');
    expect(wrapper.find(`#${dialog.attributes('aria-labelledby')}`).text()).toBe('פתח מסמך');
  });

  it('Escape סוגר', async () => {
    const { wrapper } = open();
    await wrapper.find('.open-dialog').trigger('keydown.esc');
    expect(wrapper.emitted('close')).toHaveLength(1);
  });

  it('לחיצה על הרקע סוגרת, ולחיצה בתוך החלון אינה', async () => {
    const { wrapper } = open();
    await wrapper.find('.open-dialog').trigger('click');
    expect(wrapper.emitted('close')).toBeUndefined();
    await wrapper.find('.modal-backdrop').trigger('click');
    expect(wrapper.emitted('close')).toHaveLength(1);
  });

  it('„סגור” בתחתית ו-✕ בכותרת שניהם סוגרים', async () => {
    const { wrapper } = open();
    await wrapper.find('.open-close-btn').trigger('click');
    await wrapper.find('.open-btn').trigger('click');
    expect(wrapper.emitted('close')).toHaveLength(2);
  });

  it('מלכודת המיקוד: Tab מהאחרון חוזר לראשון', async () => {
    const { wrapper } = open();
    await settle();
    const focusable = wrapper.findAll(
      '.open-dialog button:not([disabled]), .open-dialog input:not([disabled])',
    );
    const last = focusable[focusable.length - 1]!;
    (last.element as HTMLElement).focus();
    await last.trigger('keydown.tab');
    expect(document.activeElement).toBe(focusable[0]!.element);
  });
});

describe('OpenDocumentDialog — כרטיסי התבניות', () => {
  it('מרנדר כרטיס לכל תבנית, עם התווית והרמז', () => {
    const { wrapper } = open();
    const cards = wrapper.findAll('.tpl-card');
    expect(cards).toHaveLength(DOCUMENT_TEMPLATES.length);
    expect(cards.map((card) => card.find('.tpl-label').text())).toEqual(
      DOCUMENT_TEMPLATES.map((template) => template.label),
    );
  });

  it('לחיצה פולטת את המזהה של אותה תבנית', async () => {
    const { wrapper } = open();
    const index = DOCUMENT_TEMPLATES.findIndex((template) => template.id === 'two-column');
    await wrapper.findAll('.tpl-card')[index]!.trigger('click');
    expect(wrapper.emitted('create-from-template')).toEqual([['two-column']]);
  });

  /** הכרטיס הוא אייקון וכיתוב בלבד — הרמז וההערה עברו לטולטיפ. */
  it('הכרטיס אינו מציג טקסט משני', () => {
    const { wrapper } = open();
    const card = wrapper.findAll('.tpl-card')[0]!;
    expect(card.text().trim()).toBe(DOCUMENT_TEMPLATES[0]!.label);
  });

  /**
   * עד superdoc 2.11.0 נשא התיאור גם את ההערה „הטורים מצוירים הפוך בעורך;
   * הקובץ נשמר נכון”. ‏2.12.0 מצייר את הטור הראשון בימין, ההערה ירדה
   * (engine/templates.ts), ואין כרגע תבנית עם `note` — ולכן התיאור הוא הרמז
   * לבדו. חיבור הרמז וההערה עצמו נשאר ב-`tipDescription`.
   */
  it('הרמז מגיע בטולטיפ', () => {
    const { wrapper } = open();
    const index = DOCUMENT_TEMPLATES.findIndex((template) => template.id === 'two-column');
    const card = wrapper.findAll('.tpl-card')[index]!;
    const template = DOCUMENT_TEMPLATES[index]!;

    expect(card.attributes('data-tip-title')).toBe(template.label);
    expect(card.attributes('data-tip-desc')).toBe('גוף בשני טורים שווים, עם כותרת רצה');
  });

  /**
   * השם הנגיש מגיע מהתווית הגלויה (הכלל של RibbonButton: `aria-label` על פקד
   * עם תווית דורס אותה ושובר שליטה קולית), והתיאור הוא בדיוק מה שהטולטיפ
   * מראה — מקור אחד לשניהם.
   */
  it('התיאור הנגיש זהה לטולטיפ, והשם נשאר התווית', () => {
    const { wrapper } = open();
    const index = DOCUMENT_TEMPLATES.findIndex((template) => template.id === 'two-column');
    const card = wrapper.findAll('.tpl-card')[index]!;

    expect(card.attributes('aria-label'), 'אין aria-label שידרוס את התווית').toBeUndefined();
    const described = wrapper.find(`#${card.attributes('aria-describedby')}`);
    expect(described.text()).toBe(card.attributes('data-tip-desc'));
    // והתיאור יושב **מחוץ** לכפתור, אחרת הוא נספר לתוך השם הנגיש שלו.
    expect(card.element.contains(described.element)).toBe(false);
  });

  it('רק כרטיס אחד בסדר ה-Tab (roving tabindex)', () => {
    const { wrapper } = open();
    const stops = wrapper.findAll('.tpl-card').filter((card) => card.attributes('tabindex') === '0');
    expect(stops).toHaveLength(1);
  });

  it('ArrowLeft מתקדם — הכיוון של ממשק עברי', async () => {
    const { wrapper } = open();
    await wrapper.find('.tpl-grid').trigger('keydown', { key: 'ArrowLeft' });
    expect(wrapper.findAll('.tpl-card')[1]!.attributes('tabindex')).toBe('0');
  });

  /**
   * הטבלה ב-§5.5 של מפרט העיצוב. שינוי מכוון בציור מעדכן גם אותה — ושינוי
   * לא מכוון נופל כאן.
   */
  it.each([
    ['blank', 14, 0, 0],
    ['two-column', 29, 0, 0],
    ['annotated', 16, 1, 0],
    // עמוד שער בלבד: שלוש פסקאות ממורכזות, בלי גוף ובלי קו. ראו buildSheet.
    ['title-page', 4, 0, 0],
    ['kuntres-a5', 16, 0, 1],
  ])('התצוגה המקדימה של %s: %i מלבנים, %i קווים, %i קנה מידה', (id, rects, lines, groups) => {
    const { wrapper } = open();
    const index = DOCUMENT_TEMPLATES.findIndex((template) => template.id === id);
    const svg = wrapper.findAll('.tpl-sheet')[index]!;
    expect(svg.findAll('rect')).toHaveLength(rects);
    expect(svg.findAll('line')).toHaveLength(lines);
    expect(svg.findAll('g[transform]')).toHaveLength(groups);
  });

  it('לכל גיליון יש מסגרת בדיוק אחת, והיא במידות A4', () => {
    const { wrapper } = open();
    for (const svg of wrapper.findAll('.tpl-sheet')) {
      const sheets = svg.findAll('.pv-sheet');
      expect(sheets).toHaveLength(1);
      expect(sheets[0]!.attributes('width')).toBe('210');
      expect(sheets[0]!.attributes('height')).toBe('297');
    }
  });

  it('ה-SVG מוסתר מקורא מסך ואינו נושא title', () => {
    const { wrapper } = open();
    for (const svg of wrapper.findAll('.tpl-sheet')) {
      expect(svg.attributes('aria-hidden')).toBe('true');
      expect(svg.attributes('focusable')).toBe('false');
      expect(svg.find('title').exists()).toBe(false);
    }
  });
});

describe('OpenDocumentDialog — עיון בקבצים', () => {
  it('פולט browse', async () => {
    const { wrapper } = open();
    await wrapper.find('.open-browse').trigger('click');
    expect(wrapper.emitted('browse')).toHaveLength(1);
  });
});

describe('OpenDocumentDialog — מסמכים אחרונים', () => {
  it('מוצמד ראשון, והשאר לפי זמן יורד', () => {
    const { wrapper } = open();
    expect(wrapper.findAll('.rec-name').map((name) => name.text())).toEqual([
      'שמות.docx',
      'בראשית.docx',
      'ויקרא.docx',
    ]);
  });

  it('הקו מתחת למוצמדת האחרונה, ורק שם', () => {
    const { wrapper } = open();
    const marked = wrapper.findAll('.rec-row--last-pinned');
    expect(marked).toHaveLength(1);
    expect(marked[0]!.find('.rec-name').text()).toBe('שמות.docx');
  });

  it('פתיחה, הצמדה והסרה נושאות את ה-token של השורה שנלחצה', async () => {
    const { wrapper } = open();
    const row = wrapper.findAll('.rec-row')[1]!; // „בראשית”, לא מוצמד
    await row.find('.rec-open').trigger('click');
    await row.find('.rec-pin').trigger('click');
    await row.find('.rec-forget').trigger('click');
    expect(wrapper.emitted('open-recent')).toEqual([['alef']]);
    // ה-`pinned` שנשלח הוא **החדש**, לא הנוכחי.
    expect(wrapper.emitted('toggle-pin')).toEqual([['alef', true]]);
    expect(wrapper.emitted('forget-recent')).toEqual([['alef']]);
  });

  it('ביטול הצמדה שולח false', async () => {
    const { wrapper } = open();
    await wrapper.findAll('.rec-row')[0]!.find('.rec-pin').trigger('click');
    expect(wrapper.emitted('toggle-pin')).toEqual([['bet', false]]);
  });

  it('כפתור ההצמדה הוא מתג שמצבו מוכרז', () => {
    const { wrapper } = open();
    const pins = wrapper.findAll('.rec-pin');
    expect(pins[0]!.attributes('aria-pressed')).toBe('true');
    expect(pins[1]!.attributes('aria-pressed')).toBe('false');
  });

  /** WCAG 2.5.3 — השם הנגיש מתחיל בשם הנראה, וזה מה שמאפשר פקודת קול. */
  it('השם הנגיש של הפתיחה מתחיל בשם הקובץ ומסתיים ב„מוצמד”', () => {
    const { wrapper } = open();
    expect(wrapper.findAll('.rec-open')[0]!.attributes('aria-label')).toBe(
      'שמות.docx, לפני שעה, 2 ק״ב, מוצמד',
    );
  });

  /** `openedAt: 0` ו-`size: 0` הם „לא ידוע” ולא „ריק”. */
  it('שורה בלי גיל ובלי גודל אינה מציגה מטא כלל', () => {
    const { wrapper } = open();
    const row = wrapper.findAll('.rec-row')[2]!; // „ויקרא”
    expect(row.find('.rec-meta').exists()).toBe(false);
    expect(row.find('.rec-open').attributes('aria-label')).toBe('ויקרא.docx');
  });

  it('השם המלא מגיע ב-data-tip-title ולא ב-title', () => {
    const { wrapper } = open();
    const open0 = wrapper.findAll('.rec-open')[0]!;
    expect(open0.attributes('data-tip-title')).toBe('שמות.docx');
    expect(open0.attributes('title')).toBeUndefined();
  });

  it('רק שורה אחת בסדר ה-Tab', () => {
    const { wrapper } = open();
    const stops = wrapper.findAll('.rec-open').filter((el) => el.attributes('tabindex') === '0');
    expect(stops).toHaveLength(1);
  });

  it('ArrowDown מזיז שורה', async () => {
    const { wrapper } = open();
    await wrapper.find('.rec-list').trigger('keydown', { key: 'ArrowDown' });
    expect(wrapper.findAll('.rec-open')[1]!.attributes('tabindex')).toBe('0');
  });
});

describe('OpenDocumentDialog — חיפוש ומצבי ריק', () => {
  it('הקלדה פולטת update:searchQuery', async () => {
    const { wrapper } = open();
    const input = wrapper.find('.rec-search__input');
    (input.element as HTMLInputElement).value = 'שמ';
    await input.trigger('input');
    expect(wrapper.emitted('update:searchQuery')).toEqual([['שמ']]);
  });

  it('סינון מצמצם את הרשימה ומעדכן את המונה', () => {
    const { wrapper } = open({ searchQuery: 'שמות' });
    expect(wrapper.findAll('.rec-row')).toHaveLength(1);
    expect(wrapper.find('.rec-count').text()).toBe('1 מתוך 3');
  });

  it('בלי סינון המונה סופר את הכול', () => {
    const { wrapper } = open();
    expect(wrapper.find('.rec-count').text()).toBe('3 מסמכים');
  });

  it('כפתור הניקוי מופיע רק כשיש שאילתה, ופולט מחרוזת ריקה', async () => {
    expect(open().wrapper.find('.rec-search__clear').exists()).toBe(false);
    const { wrapper } = open({ searchQuery: 'שמות' });
    await wrapper.find('.rec-search__clear').trigger('click');
    expect(wrapper.emitted('update:searchQuery')).toEqual([['']]);
  });

  it('סינון בלי תוצאות: הרשימה נעלמת, החיפוש נשאר, ויש „נקה סינון”', async () => {
    const { wrapper } = open({ searchQuery: 'איוב' });
    expect(wrapper.find('.rec-list').exists()).toBe(false);
    expect(wrapper.find('.rec-search').exists()).toBe(true);
    expect(wrapper.find('.rec-empty').text()).toContain('איוב');
    await wrapper.find('.rec-empty__clear').trigger('click');
    expect(wrapper.emitted('update:searchQuery')).toEqual([['']]);
  });

  /** ההבדל מהמצב שמעליו: פקד סינון על רשימה ריקה אינו יכול לעשות דבר. */
  it('אין אחרונים כלל: אין תיבת חיפוש ואין מונה', () => {
    const { wrapper } = open({ recents: [] });
    expect(wrapper.find('.rec-search').exists()).toBe(false);
    expect(wrapper.find('.rec-count').exists()).toBe(false);
    expect(wrapper.find('.rec-empty__title').text()).toBe('עדיין אין מסמכים אחרונים');
  });
});

describe('OpenDocumentDialog — נסגרו בלי לשמור', () => {
  it('אינו מרונדר כשאין מה לשחזר', () => {
    expect(open().wrapper.find('.open-discarded').exists()).toBe(false);
  });

  it('מופיע עם המונה, ופולט show-discarded', async () => {
    const { wrapper } = open({ discardedCount: 3 });
    const link = wrapper.find('.open-discarded');
    expect(link.text()).toBe('נסגרו בלי לשמור (3)');
    await link.trigger('click');
    expect(wrapper.emitted('show-discarded')).toHaveLength(1);
  });
});

describe('OpenDocumentDialog — busy', () => {
  it('מנטרל את כל פקדי הפעולה, ומשאיר את הסגירה', () => {
    const { wrapper } = open({ busy: true });
    for (const card of wrapper.findAll('.tpl-card')) {
      expect(card.attributes('disabled')).toBeDefined();
    }
    expect(wrapper.find('.open-browse').attributes('disabled')).toBeDefined();
    expect(wrapper.find('.rec-search__input').attributes('disabled')).toBeDefined();
    expect(wrapper.find('.rec-open').attributes('disabled')).toBeDefined();
    expect(wrapper.find('.rec-pin').attributes('disabled')).toBeDefined();

    // מודאל שאי-אפשר לצאת ממנו גרוע מפתיחה שמסתיימת ברקע.
    expect(wrapper.find('.open-close-btn').attributes('disabled')).toBeUndefined();
    expect(wrapper.find('.open-btn').attributes('disabled')).toBeUndefined();
  });

  it('מכריז על עצמו, ואומר מה קורה', () => {
    const { wrapper } = open({ busy: true });
    expect(wrapper.find('.open-dialog').attributes('aria-busy')).toBe('true');
    expect(wrapper.find('.open-status').text()).toBe('פותח מסמך…');
  });
});

/**
 * הכוונה — „מסמך חדש” מול „פתח קובץ” — אינה שני מסכים אלא שתי נחיתות באותו
 * מסך. מה שנמדד כאן הוא **המיקוד והמבנה**; הגלילה עצמה נמדדת בדפדפן אמיתי
 * (`scripts/open-dialog-probe.mjs`), כי ל-jsdom אין פריסה ולכן `scrollTop`
 * שלו הוא תמיד אפס — בדיקה שהייתה נכתבת כאן הייתה עוברת ירוקה על קוד שבור.
 */
describe('OpenDocumentDialog — הכוונה שממנה נפתח', () => {
  /** המיקוד נקבע במעבר סגור→פתוח, ולכן ההרכבה מתחילה סגורה. */
  async function reopen(props: Record<string, unknown> = {}) {
    const { wrapper } = mountUi(OpenDocumentDialog, {
      props: {
        isOpen: false,
        templates: DOCUMENT_TEMPLATES,
        recents: RECENTS,
        busy: false,
        searchQuery: '',
        ...props,
      },
    });
    await wrapper.setProps({ isOpen: true });
    await settle();
    return wrapper;
  }

  it('„עיון בקבצים…” והאחרונים הם פן אחד, בתוך אותה עטיפה', () => {
    const { wrapper } = open();
    const pane = wrapper.find('.open-pane');
    expect(pane.exists()).toBe(true);
    expect(pane.find('.open-browse').exists()).toBe(true);
    expect(pane.find('.rec-section').exists()).toBe(true);
    // הרצועה נשארת **מחוץ** לפן — היא מה שנשאר לגלול מעליו.
    expect(pane.find('.tpl-section').exists()).toBe(false);
  });

  it("'open' ממקד את „עיון בקבצים…”", async () => {
    const wrapper = await reopen({ intent: 'open' });
    expect(document.activeElement).toBe(wrapper.find('.open-browse').element);
  });

  it("'new' ממקד את הכרטיס הראשון", async () => {
    const wrapper = await reopen({ intent: 'new' });
    expect(document.activeElement).toBe(wrapper.find('.tpl-card').element);
  });

  it('ברירת המחדל היא „מסמך חדש”', async () => {
    const wrapper = await reopen();
    expect(document.activeElement).toBe(wrapper.find('.tpl-card').element);
  });

  /**
   * מיקוד על כפתור מנוטרל הוא no-op שקט, והמיקוד היה נשאר מאחורי המודאל —
   * בדיוק הכשל שרשת הביטחון בקומפוננטה נכתבה בשבילו. הוא נבדק בשני המסלולים
   * כי `intent` הוסיף יעד שני שאף בדיקה לא כיסתה.
   */
  it('busy: המיקוד נופל לחלון עצמו ולא לכפתור מנוטרל', async () => {
    const openWrapper = await reopen({ intent: 'open', busy: true });
    expect(document.activeElement).toBe(openWrapper.find('.open-dialog').element);

    const newWrapper = await reopen({ intent: 'new', busy: true });
    expect(document.activeElement).toBe(newWrapper.find('.open-dialog').element);
  });
});

describe('OpenDocumentDialog — גרסה קודמת', () => {
  it('אינו מרונדר כשאין גרסה קודמת', () => {
    expect(open().wrapper.find('.open-previous').exists()).toBe(false);
  });

  it('מופיע עם הגיל, ופולט open-previous', async () => {
    const { wrapper } = open({ previousVersion: { name: 'חידושים.docx', age: 'לפני 4 דקות' } });
    const link = wrapper.find('.open-previous');
    expect(link.text()).toBe('גרסה קודמת (לפני 4 דקות)');
    await link.trigger('click');
    expect(wrapper.emitted('open-previous')).toHaveLength(1);
  });

  it('בלי גיל — הכפתור עדיין מופיע, בלי סוגריים ריקים', () => {
    const { wrapper } = open({ previousVersion: { name: 'חידושים.docx', age: '' } });
    expect(wrapper.find('.open-previous').text()).toBe('גרסה קודמת');
  });
});
