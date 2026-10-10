/**
 * פס הכותרת.
 *
 * שלוש התקלות שהבדיקות כאן שומרות עליהן אינן מפילות כלום ואינן נראות ב-
 * typecheck — הן נראות רק בעין, ורק אם מסתכלים:
 *
 *   1. **פקד שאינו פקד.** תיבת החיפוש הייתה `input readonly` שנראה כמו מקום
 *      להקליד בו, וה-`@click` יושב על ה-`div` העוטף — כלומר המקלדת לא הגיעה
 *      אליו. מתג השמירה האוטומטית היה `div` בלי `role`, בלי פוקוס ובלי הכרזה
 *      אם הוא דלוק.
 *   2. **תנועה שאינה הגיונית.** `translateX(-12px)` על כפתור המתג הוא תנועה
 *      שמאלה בשתי הכיווניות, ולכן ב-LTR הוא יצא מהפיל. הכלל היה גם משוכפל
 *      ל-`[dir="rtl"]` באותם ערכים בדיוק.
 *   3. **שני מקורות אמת לאותו אלמנט.** `.topbar` (shell.css) ו-`.word-titlebar`
 *      (הבלוק ה-scoped) מעצבים את אותה כותרת. הם כבר נפרדו בשקט: 16px מול 12px
 *      ריפוד, ושני צבעי גבול שונים.
 *
 * שתי הבדיקות על הפקדים עצמם (החיפוש שהוא כפתור ולא `input readonly`, והמתג
 * שהוא `switch` נגיש) **אינן כאן יותר**: הן נמדדות בהרכבה
 * (tests/component/shell-bars.test.ts), שבודקת גם את מה שסריקה אינה יכולה —
 * שהמתג מקבל מיקוד ושהלחיצה נפלטת. מה שנשאר כאן הוא החלטות ה-CSS והפריסה,
 * ושתי טענות היעדר על המקור.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  DOC_TITLE_MAX_CH,
  DOC_TITLE_MIN_CH,
  docTitleWidthCh,
} from '../../src/composables/shell-format';

/** vitest רץ משורש המאגר, ולכן src/ נמצא ביחס ל-cwd. */
const SRC = join(process.cwd(), 'src');

/**
 * ההערות מוסרות לפני הסריקה: התיעוד בקומפוננטה מסביר במפורש מה **היה** שם
 * („`input readonly`”, „הכלל הכפול ל-[dir=\"rtl\"]”), וההסבר הזה אינו הפרה.
 * בלי ההסרה כל בדיקה שמחפשת היעדר הייתה מוצאת את התיעוד של עצמה.
 */
function stripComments(source: string): string {
  return source.replace(/<!--[\s\S]*?-->/g, '').replace(/\/\*[\s\S]*?\*\//g, '');
}

const TITLEBAR = stripComments(readFileSync(join(SRC, 'ui/shell/TitleBar.vue'), 'utf8'));
const SHELL_CSS = stripComments(readFileSync(join(SRC, 'styles/shell.css'), 'utf8'));

/** התבנית בלבד — עד `</template>`. הסגנונות אינם פקדים. */
const TEMPLATE = TITLEBAR.slice(0, TITLEBAR.indexOf('</template>'));

/**
 * שמות המאפיינים שבלוק ה-CSS של סלקטור מגדיר. הסלקטור נלקח בהתאמה מדויקת
 * (`^` על תחילת שורה) כדי שכלל צר יותר עם אותו שם לא ייספר.
 */
function declaredProperties(css: string, selector: string): string[] {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const block = css.match(new RegExp(`^${escaped}\\s*\\{([^}]*)\\}`, 'm'));
  if (!block) throw new Error(`לא נמצא בלוק לסלקטור ${selector}`);
  return [...block[1].matchAll(/^\s*([a-z-]+)\s*:/gm)].map((m) => m[1]);
}

describe('רוחב שדה שם המסמך', () => {
  it('גדל עם השם', () => {
    expect(docTitleWidthCh('חידושים')).toBeLessThan(docTitleWidthCh('חידושים על מסכת בבא מציעא'));
  });

  it('שם ארוך נעצר בתקרה, ושם קצר ברצפה', () => {
    expect(DOC_TITLE_MAX_CH).toBe(48);
    expect(docTitleWidthCh('א'.repeat(200))).toBe(DOC_TITLE_MAX_CH);
    expect(docTitleWidthCh('א')).toBe(DOC_TITLE_MIN_CH);
    expect(docTitleWidthCh('')).toBe(DOC_TITLE_MIN_CH);
  });

  it('רוחב קשיח בפיקסלים אינו חוזר לשדה', () => {
    // `width: 110px` חתך „חידושים על מסכת בבא מציעא” אחרי שש אותיות.
    const block = TITLEBAR.match(/\.doc-title-input\s*\{[^}]*\}/)?.[0] ?? '';
    // `min-width` אינו „רוחב קשיח” — הוא מה שמאפשר לשדה להיצמד לגריד.
    expect(block).not.toMatch(/(?<![-\w])width\s*:/);
    expect(TEMPLATE).toContain('docTitleWidthCh(title)');
  });

  it('שם המסמך מיושר לשמאל גם בממשק RTL', () => {
    const block = TITLEBAR.match(/\.doc-title-input\s*\{[^}]*\}/)?.[0] ?? '';
    expect(block).toMatch(/text-align:\s*left/);
  });
});

describe('פקדים אמיתיים בפס הכותרת', () => {
  it('אין יותר כפתור נסתר #open', () => {
    // שער ה-boot מודד `data-boot` (scripts/boot-check.mjs), ולא את קיום הכפתור.
    expect(TITLEBAR).not.toMatch(/id="open"/);
    // `TEMPLATE` ולא `TITLEBAR`: הכוונה המקורית הייתה פקד שמוסתר לצמיתות
    // ב-DOM (`display: none` על אלמנט בתבנית עצמה, כמו #open שהיה). זה שונה
    // מ-`display: none` בתוך `@media` בבלוק ה-`<style>` — שם הוא הכלי
    // הרגיל להסתרה **מותנית-רוחב** של תוכן משני (תווית „שמירה אוטומטית”,
    // סיומת „- Word”), ולא פקד נסתר-לצמיתות. בדיקה על הקובץ כולו הייתה
    // פוסלת כל שימוש לגיטימי ב-@media לרספונסיביות.
    expect(TEMPLATE).not.toMatch(/display:\s*none/);
  });

  it('אין חץ פתיחה מדומה ליד שם המסמך', () => {
    // הוא רמז על תפריט שלא היה קיים.
    expect(TITLEBAR).not.toMatch(/title-dropdown-icon/);
  });
});

describe('תנועת המתג', () => {
  it('התנועה לוגית, ולכן נכונה בשתי הכיווניות', () => {
    const active = TITLEBAR.match(/\.autosave-toggle\.active \.toggle-thumb\s*\{[^}]*\}/)?.[0] ?? '';
    expect(active).toMatch(/inset-inline-start/);
    expect(active).not.toMatch(/translateX/);
  });

  it('אין כלל [dir="rtl"] כפול למתג', () => {
    expect(TITLEBAR).not.toMatch(/\[dir="rtl"\][^{]*\.toggle-thumb/);
  });

  it('הכפתור נשאר בתוך הפיל', () => {
    const pill = TITLEBAR.match(/\.toggle-pill\s*\{[^}]*\}/)?.[0] ?? '';
    const thumb = TITLEBAR.match(/\.toggle-thumb\s*\{[^}]*\}/)?.[0] ?? '';
    const active = TITLEBAR.match(/\.autosave-toggle\.active \.toggle-thumb\s*\{[^}]*\}/)?.[0] ?? '';

    const pillWidth = Number(pill.match(/width:\s*(\d+)px/)?.[1]);
    const thumbWidth = Number(thumb.match(/width:\s*(\d+)px/)?.[1]);
    const offset = Number(thumb.match(/inset-inline-start:\s*(\d+)px/)?.[1]);
    const travel = Number(active.match(/inset-inline-start:\s*(\d+)px/)?.[1]);

    expect(travel + thumbWidth + offset).toBe(pillWidth);
  });
});

describe('פריסת הפס', () => {
  it('החיפוש ממורכז בין עמודות קצה שוות', () => {
    const bar = TITLEBAR.match(/\.word-titlebar\s*\{[^}]*\}/)?.[0] ?? '';
    expect(bar).toMatch(/display:\s*grid/);
    const columns = bar.match(/grid-template-columns:\s*([^;]+);/)?.[1] ?? '';
    const [start, center, end] = columns.trim().split(/\s+(?![^(]*\))/);
    expect(start).toBe('minmax(160px, 1fr)');
    expect(center).toBe('minmax(0, 320px)');
    expect(start).toBe(end);
    expect(bar).not.toMatch(/justify-content/);
  });

  it('תיבת שם המסמך מתרחבת למקום הפנוי בלי לדחוק את הפקדים', () => {
    const wrappers = TITLEBAR.match(/\.doc-title-wrapper\s*\{[^}]*\}/g) ?? [];
    const wrapper = wrappers[wrappers.length - 1] ?? '';
    expect(wrapper).toMatch(/flex:\s*1 1 0/);
    const input = TITLEBAR.match(/\.doc-title-input\s*\{[^}]*\}/)?.[0] ?? '';
    expect(input).toMatch(/flex:\s*1 1 auto/);
  });

  it('הבלוק ה-scoped אינו מגדיר מחדש מה ש-.topbar כבר מגדיר', () => {
    // display הוא החריג המכוון: grid דורס את ה-flex לטובת המרכוז.
    const shared = declaredProperties(TITLEBAR, '.word-titlebar').filter((property) =>
      declaredProperties(SHELL_CSS, '.topbar').includes(property),
    );
    expect(shared).toEqual(['display']);
  });

  it('.topbar נשאר ב-shell.css — ולידציית העיצוב מחריגה אותו בשם הזה', () => {
    expect(SHELL_CSS).toMatch(/^\.topbar\s*\{/m);
    expect(TEMPLATE).toMatch(/class="topbar word-titlebar"/);
  });
});
