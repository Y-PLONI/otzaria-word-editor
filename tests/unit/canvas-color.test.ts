/**
 * צבע הבד — המשטח שסביב הדף.
 *
 * שתי החלטות נמדדות כאן, ושתיהן היו יכולות להיכשל בשקט:
 *
 * 1. **מה נחשב צבע.** הערך מגיע מ-`storage` של אוצריא, כלומר JSON שנכתב
 *    בהפעלה קודמת. צבע פגום שמגיע עד `setProperty` אינו זורק ואינו מדווח —
 *    הדפדפן פשוט מתעלם מההצהרה, והמשתמש מקבל בד בלי רקע כלל.
 *
 * 2. **„אין העדפה” הוא היעדר ההצהרה, ולא צבע שני.** `removeProperty` מחזיר
 *    את הבד לברירת המחדל שב-tokens.css (חום בהיר). אילו „ברירת מחדל” הייתה
 *    נכתבת כצבע, היא הייתה נשמרת ב-`storage` כבחירה, ושינוי עתידי של ברירת
 *    המחדל לא היה מגיע למי שלחץ עליה.
 */
import { beforeEach, describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  CANVAS_COLOR_VAR,
  DEFAULT_CANVAS_COLOR,
  applyCanvasColor,
  canvasColor,
  canvasDefaultColor,
  normalizeCanvasColor,
  setCanvasDefaultColor,
} from '../../src/composables/canvas-color';

function source(...parts: string[]): string {
  return readFileSync(join(process.cwd(), 'src', ...parts), 'utf8');
}

/** ההצהרה על שורש המסמך, כפי שהדפדפן היה קורא אותה. */
function declared(): string {
  return document.documentElement.style.getPropertyValue(CANVAS_COLOR_VAR);
}

beforeEach(() => {
  applyCanvasColor(null);
  setCanvasDefaultColor(DEFAULT_CANVAS_COLOR);
});

describe('normalizeCanvasColor', () => {
  it('צבע מהפלטה ומהדו-שיח עובר כמו שהוא', () => {
    // שני המקורות היחידים: הפלטה של ColorPickerPopover כתובה באותיות קטנות,
    // ו-`input[type=color]` מחזיר „simple color” — שבעה תווים.
    expect(normalizeCanvasColor('#edebe9')).toBe('#edebe9');
    expect(normalizeCanvasColor('#000000')).toBe('#000000');
  });

  it('אותיות גדולות ורווחים מסביב מיושרים לצורה אחת', () => {
    // ההשוואה בפלטה („איזו משבצת מסומנת”) היא על מחרוזת, ולכן #FFFF00 ו-
    // #ffff00 שנשמרים כשניים היו מציגים משבצת לא מסומנת על צבע שנבחר.
    expect(normalizeCanvasColor('#EDEBE9')).toBe('#edebe9');
    expect(normalizeCanvasColor('  #EdEbE9\n')).toBe('#edebe9');
  });

  it('כל מה שאינו #rrggbb נדחה', () => {
    // המסלול האמיתי: ערך שנכתב בגרסה קודמת, ערך שנקטע, או קריאה שהחזירה
    // צורה אחרת. `null` ו-`undefined` הם גם „אין ערך במפתח”.
    for (const raw of [
      null,
      undefined,
      '',
      '   ',
      'red',
      '#fff',
      '#edebe',
      '#edebe99',
      'edebe9',
      '#gggggg',
      'rgb(237, 235, 233)',
      237,
      { hex: '#edebe9' },
      ['#edebe9'],
    ]) {
      expect(normalizeCanvasColor(raw)).toBeNull();
    }
  });
});

describe('applyCanvasColor', () => {
  it('צבע נכתב כהצהרה על שורש המסמך', () => {
    // inline על :root הוא מה שמנצח את הכלל שב-tokens.css — אותו מנגנון
    // שבו host/theme.ts כותב את צבעי אוצריא.
    applyCanvasColor('#123456');

    expect(declared()).toBe('#123456');
    expect(canvasColor.value).toBe('#123456');
  });

  describe('ברירת המחדל של הבד', () => {
    it('מתעדכנת מערך הנושא ומנרמלת את הצבע', () => {
      setCanvasDefaultColor('#AABBCC');

      expect(canvasDefaultColor.value).toBe('#aabbcc');
    });

    it('חוזרת לצבע ה-fallback כשערכת הנושא אינה מספקת hex תקין', () => {
      setCanvasDefaultColor('var(--color-surface)');

      expect(canvasDefaultColor.value).toBe(DEFAULT_CANVAS_COLOR);
    });

    it('מכהה את ברירת המחדל בלבד ומשאירה בחירה מותאמת ללא שכבה', () => {
      const app = source('App.vue');

      expect(app).toContain("'canvas-default': canvasColor === null");
      expect(app).toContain('.editor-stack.canvas-default');
      expect(app).toContain('background-image: linear-gradient(var(--color-shell-dim)');
    });
  });

  it('`null` מסיר את ההצהרה ואינו כותב צבע אחר', () => {
    // הלב של „ברירת מחדל”: הבד חוזר לערך שב-tokens.css, ולכן אסור שיישאר
    // כאן ערך כלשהו — גם לא הצבע שהיה שם רגע קודם.
    applyCanvasColor('#123456');
    applyCanvasColor(null);

    expect(declared()).toBe('');
    expect(canvasColor.value).toBeNull();
  });

  it('החלפת צבע דורסת ואינה מצטברת', () => {
    applyCanvasColor('#123456');
    applyCanvasColor('#abcdef');

    expect(declared()).toBe('#abcdef');
    expect(canvasColor.value).toBe('#abcdef');
  });
});

/**
 * שני קצוות שאין ביניהם קשר שהמהדר רואה: שם הטוקן כמחרוזת ב-TypeScript, ואותו
 * שם בגיליון הסגנון. שינוי צד אחד בלבד אינו נופל בשום שער — הכתיבה מצליחה,
 * הכלל ממשיך לקרוא את הטוקן הישן, והבד פשוט מפסיק להיצבע. נמדד: שינוי השם
 * ב-TS בלבד השאיר את כל שאר הבדיקות בקובץ הזה ירוקות.
 */
describe('הטוקן שב-TypeScript הוא הטוקן שב-CSS', () => {
  it('הכלל של הבד צורך בדיוק את `CANVAS_COLOR_VAR`', () => {
    // שני הכללים: ה-`scoped` ב-App.vue והגלובלי ב-shell.css.
    expect(source('App.vue')).toContain(`background: var(${CANVAS_COLOR_VAR});`);
    expect(source('styles', 'shell.css')).toContain(
      `background-color: var(${CANVAS_COLOR_VAR});`,
    );
  });

  it('ברירת המחדל של הטוקן עוקבת אחרי ערכת הנושא', () => {
    expect(source('styles', 'tokens.css')).toContain(
      `${CANVAS_COLOR_VAR}: var(--color-surface-container-lowest);`,
    );
  });

  it('המסילה משתמשת בצבע משטח נפרד מצבע הקנבס', () => {
    expect(source('styles', 'shell.css')).toContain(
      'scrollbar-color: var(--color-outline) var(--color-surface-container-high);',
    );
  });
});
