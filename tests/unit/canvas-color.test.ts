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
  BAND_MIX,
  CANVAS_COLOR_VAR,
  DEFAULT_CANVAS_COLOR,
  FALLBACK_SURFACE,
  FALLBACK_SURFACE_CONTAINER_HIGH,
  applyCanvasColor,
  bandColor,
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
  setCanvasDefaultColor({});
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

  it('הבד בלי העדפה הוא צבע הפס העליון, והפס אינו נגזר מהבד', () => {
    const tokens = source('styles', 'tokens.css');
    expect(tokens).toContain(`${CANVAS_COLOR_VAR}: var(--word-shell-band-bg);`);
    // הכיוון הפוך היה מעביר את „צבע רקע” של המשתמש גם לפקדי המעטפת.
    expect(tokens).not.toMatch(/--word-shell-band-bg:\s*var\(--word-canvas-bg\)/);
  });

  it('התערובת שב-CSS היא התערובת שב-TypeScript', () => {
    // הפס בבורר מחושב ב-TS; שני יחסים שונים פירושם פס שמבטיח צבע אחד ובד
    // שמצויר באחר.
    expect(source('styles', 'tokens.css')).toContain(
      `--word-shell-band-bg: color-mix(in srgb, var(--color-surface-container-high) ${BAND_MIX * 100}%, var(--color-surface));`,
    );
    expect(source('styles', 'tokens.css')).toContain(`--color-surface: ${FALLBACK_SURFACE};`);
    expect(source('styles', 'tokens.css')).toContain(
      `--color-surface-container-high: ${FALLBACK_SURFACE_CONTAINER_HIGH};`,
    );
  });
});

describe('ברירת המחדל שבבורר', () => {
  it('בהירה משורת הטאבים ונגזרת ממנה', () => {
    // 0x40×0.35 + 0x80×0.65 = 105.6 → 0x6a: בין שני הצבעים, קרוב למשטח.
    expect(bandColor('#404040', '#808080')).toBe('#6a6a6a');
    setCanvasDefaultColor({ surface: '#ffffff', surfaceContainerHigh: '#e0d0c0' });
    expect(canvasDefaultColor.value).toBe(bandColor('#e0d0c0', '#ffffff'));
    expect(canvasDefaultColor.value).toBe('#f4efe9');
  });

  it('בלי `surfaceContainerHigh` — `surfaceContainerHighest`, כמו ב-host/theme.ts', () => {
    setCanvasDefaultColor({ surface: '#ffffff', surfaceContainerHighest: '#000000' });
    expect(canvasDefaultColor.value).toBe(bandColor('#000000', '#ffffff'));
  });

  it('ערכת נושא בלי צבעים תקינים משאירה את ברירת המחדל של tokens.css', () => {
    setCanvasDefaultColor({ surface: 'var(--x)' });
    expect(canvasDefaultColor.value).toBe(DEFAULT_CANVAS_COLOR);
  });
});

describe('פס הגלילה של הבד', () => {
  it('המסילה משתמשת בצבע משטח נפרד מצבע הקנבס', () => {
    expect(source('styles', 'shell.css')).toContain(
      'scrollbar-color: var(--color-outline) var(--color-surface);',
    );
  });

  it('אין כללי `::-webkit-scrollbar` על מיכל הגלילה — `scrollbar-color` מבטל אותם', () => {
    // נמדד ב-Chrome 154: כלל webkit של 40px נתן 40px, ועם `scrollbar-color`
    // על אותו אלמנט — 15px, ברירת המחדל. כלל כזה כאן הוא קוד מת שנראה חי,
    // ו-getComputedStyle על הפסאודו מחזיר אותו גם כשאינו מצויר.
    expect(source('styles', 'shell.css')).not.toMatch(/\.editor-stack__host[^{]*::-webkit-scrollbar/);
  });
});
