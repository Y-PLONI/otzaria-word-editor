/**
 * שער הגאומטריה של הרצועה.
 *
 * ## מה נשבר, ולמה סריקה ולא עין
 *
 * גובה הרצועה לא היה קבוע. `.word-ribbon-body` נשען על `min-height` בלבד,
 * ולכן גובהו בפועל נקבע לפי הקבוצה הגבוהה **בלשונית הפעילה**: כפתור גדול היה
 * 68px, מחסנית של שלושה כפתורים קטנים 76px, ומחסנית של ארבעה — שהייתה בקבוצת
 * „שינויים” של „סקירה” — 102px. התוצאה: 96px ברוב הלשוניות, 100px ב„בית”
 * וב„אוצריא”, ו-126px ב„סקירה”. כל החלפת לשונית הזיזה את המסמך כולו.
 *
 * אף אחד מהמספרים האלה לא נכתב בשום מקום. הם נגזרו מארבעה כללי CSS שנקבעו
 * בנפרד, ולכן גם אי-אפשר היה לראות את הבעיה בקריאת הקוד — רק בהחלפת לשונית
 * מול מסמך פתוח. זה בדיוק סוג הדבר שבדיקה תופסת ועין לא.
 *
 * ## החוזה שנשמר כאן
 *
 * 1. **כפתור גדול = שלוש שורות קטנות.** זה הכלל של Word, וכאן הוא נגזר מטוקן
 *    אחד (`--ribbon-row-h`) במקום להיכתב פעמיים.
 * 2. **מחסנית מחזיקה עד שלושה.** זה מה שמחזיק את (1): רביעי גולש מהקבוצה.
 *    זו הבדיקה החשובה כאן — היא היחידה שמונעת את החזרה של „סקירה”.
 * 3. **גובה הגוף קבוע ונגזר**, ולא `min-height` שהתוכן קובע.
 * 4. **אין מספרי גובה קשיחים** בפקדים שיושבים באותה שורה: בורר הגופן ובורר
 *    הצבע נמצאים ב-`.word-group-row` אחת ב„גופן”, ו-24 מול 22 השאיר אותם
 *    לא מיושרים.
 * 5. **`.column-items` אינו קיים יותר** — הוא היה משוכפל מילה במילה בארבעה
 *    בלוקים `scoped`, ו-`RibbonStack.vue` הוא הבעלים היחיד שלו עכשיו.
 *
 * ## מה השער הזה **אינו** בודק
 *
 * הוא קורא את המקור, ולכן הוא שומר על הכתיבה — לא על התוצאה. jsdom אינו מחשב
 * `calc`, `flex` או גלישה, ולכן CSS תקין-למראה שהגובה שלו קופץ עובר כאן.
 * הגובה בפועל, בכל שמונה הלשוניות ואחרי החלפה אמיתית, נמדד בדפדפן:
 * `npm run check:ribbon` (scripts/ribbon-geometry-probe.mjs). נמדד שהוא באמת
 * נופל על הקוד שלפני התיקון — 96/118/126/100, והמסמך זז 30px.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

/** vitest רץ משורש המאגר, ולכן src/ נמצא ביחס ל-cwd. */
const SRC = join(process.cwd(), 'src');
const TABS_DIR = join(SRC, 'ui/ribbon/tabs');

const TAB_FILES = readdirSync(TABS_DIR).filter((file) => file.endsWith('.vue'));
const TAB_SOURCE = new Map(
  TAB_FILES.map((file) => [file, readFileSync(join(TABS_DIR, file), 'utf8')]),
);

function read(relative: string): string {
  return readFileSync(join(SRC, relative), 'utf8');
}

/** גוש הכללים של סלקטור אחד, בלי גוש אחר שמכיל אותו כתחילית. */
function block(css: string, selector: string): string {
  const pattern = new RegExp(`${selector}\\s*\\{([^}]*)\\}`);
  return pattern.exec(css)?.[1] ?? '';
}

const TOKENS = read('styles/tokens.css');
const RIBBON_CSS = read('styles/ribbon.css');

/* ------------------------------------------------------------------ */
/* תקרת השלושה                                                        */
/* ------------------------------------------------------------------ */

/**
 * מספר הפקדים בכל `<RibbonStack>` בקובץ.
 *
 * ספירת תגי פתיחה ולא ניתוח DOM: אין כאן DOM, והתבניות כתובות בסגנון אחיד —
 * כל פקד נפתח בשורה משלו. מחסנית מקוננת אינה קיימת ולא תהיה, וזה בדיוק מה
 * שהתקרה אומרת.
 */
function stackSizes(source: string): number[] {
  const sizes: number[] = [];
  for (const found of source.matchAll(/<RibbonStack>([\s\S]*?)<\/RibbonStack>/g)) {
    sizes.push([...found[1].matchAll(/<Ribbon(?:Button|MenuButton|Select)\b/g)].length);
  }
  return sizes;
}

describe('תקרת המחסנית', () => {
  it('נמצאו מחסניות לבדוק', () => {
    // בלי זה כל הבדיקות למטה עוברות בירוק על רשימה ריקה — למשל אם שם התג ישונה.
    const total = TAB_FILES.reduce(
      (count, file) => count + stackSizes(TAB_SOURCE.get(file) ?? '').length,
      0,
    );
    expect(total).toBeGreaterThan(0);
  });

  it('אף מחסנית אינה מחזיקה יותר משלושה פקדים', () => {
    // רביעי גולש מ---ribbon-content-h, וזה בדיוק מה שהחזיק את „סקירה” ב-126px.
    const over: string[] = [];
    for (const file of TAB_FILES) {
      stackSizes(TAB_SOURCE.get(file) ?? '').forEach((size, index) => {
        if (size > 3) over.push(`${file}: מחסנית #${index + 1} מחזיקה ${size}`);
      });
    }
    expect(over).toEqual([]);
  });

  it('אף מחסנית אינה ריקה', () => {
    const empty: string[] = [];
    for (const file of TAB_FILES) {
      stackSizes(TAB_SOURCE.get(file) ?? '').forEach((size, index) => {
        if (size === 0) empty.push(`${file}: מחסנית #${index + 1}`);
      });
    }
    expect(empty).toEqual([]);
  });
});

/* ------------------------------------------------------------------ */
/* הגובה נגזר, ואינו נכתב                                             */
/* ------------------------------------------------------------------ */

describe('גאומטריה נגזרת', () => {
  it('גובה הכפתור הגדול הוא בדיוק גובה שלוש שורות קטנות', () => {
    // הטענה נמדדת על ההגדרה ולא על מספר: `--ribbon-content-h` חייב להיגזר
    // מ---ribbon-row-h, ושני גדלי הכפתור חייבים לצרוך את שני הטוקנים האלה.
    expect(TOKENS).toMatch(
      /--ribbon-content-h:\s*calc\(\s*3\s*\*\s*var\(--ribbon-row-h\)\s*\+\s*2\s*\*\s*var\(--ribbon-row-gap\)\s*\)/,
    );
    expect(block(RIBBON_CSS, '\\.word-btn\\.btn-large')).toMatch(
      /height:\s*var\(--ribbon-content-h\)/,
    );
    expect(block(RIBBON_CSS, '\\.word-btn\\.btn-small')).toMatch(/height:\s*var\(--ribbon-row-h\)/);
  });

  it('הריפוד והגבול של הגוף נגזרים מאותם טוקנים שהגובה בולע', () => {
    // `+ 9px` מחובר ביד היה נכון רק כל עוד איש אינו נוגע בריפוד. שינוי שלו
    // אינו מרחיב את הגוף (הגובה קבוע) אלא מקטין את המקום הפנוי — כלומר גלישה
    // שקטה, ובדיוק בלי שום סימן.
    const height = /--ribbon-height:([^;]*);/.exec(TOKENS)?.[1] ?? '';
    expect(height).toMatch(/2\s*\*\s*var\(--ribbon-body-pad\)/);
    expect(height).toMatch(/var\(--ribbon-body-border\)/);
    expect(height, 'מספר קשיח בחשבון הגובה').not.toMatch(/\d+px/);
    const body = block(RIBBON_CSS, '\\.word-ribbon-body');
    expect(body).toMatch(/padding:\s*var\(--ribbon-body-pad\)/);
    expect(body).toMatch(/border-block-end:\s*var\(--ribbon-body-border\)/);
  });

  it('גוף הרצועה בגובה קבוע, ולא בגובה שהתוכן קובע', () => {
    // `min-height` הוא מה שנתן לכל לשונית את הגובה של הקבוצה הגבוהה שלה.
    const body = block(RIBBON_CSS, '\\.word-ribbon-body');
    expect(body).toMatch(/(?:^|[\s;])height:\s*var\(--ribbon-height\)/);
    expect(body).not.toMatch(/min-height/);
  });

  it('לגובה הרצועה הגדרה אחת — אין רוחב חלון שדורס אותה', () => {
    // `@media (max-width: 600px) { --ribbon-height: auto }` נשאר מהשלד הראשון,
    // ומתחת ל-600px כל לשונית קיבלה שוב את גובה התוכן שלה: 49px ללשונית שכל
    // קבוצותיה צ'יפים, 94px לשכנתה (scripts/qa/ribbon-collapse-qa.mjs).
    const code = TOKENS.replace(/\/\*[\s\S]*?\*\//g, '');
    expect([...code.matchAll(/--ribbon-height\s*:/g)]).toHaveLength(1);
  });

  it('תוכן הקבוצה בגובה קבוע — זה מה שמשווה בין הלשוניות', () => {
    const content = block(RIBBON_CSS, '\\.word-group-content');
    expect(content).toMatch(/height:\s*var\(--ribbon-content-h\)/);
    expect(content).not.toMatch(/flex:\s*1/);
  });

  it('רזרבת פס הגלילה זהה לגובה הפס שמצויר', () => {
    // אם השניים יתפצלו, הרזרבה שבטוקן היא שקר והגוף גדל כשהרצועה גולשת.
    expect(block(RIBBON_CSS, '\\.word-ribbon-body::-webkit-scrollbar')).toMatch(
      /height:\s*var\(--ribbon-scrollbar-h\)/,
    );
    // `scrollbar-width` על אותו אלמנט מבטל את ההתאמה ב-Chromium, ואז הפס
    // אינו בגובה שהרזרבה מניחה.
    expect(block(RIBBON_CSS, '\\.word-ribbon-body')).not.toMatch(/scrollbar-width/);
  });

  it('הלשונית הפעילה אינה מודגשת — ההדגשה הרחיבה אותה והזיזה את הסרגל', () => {
    expect(block(RIBBON_CSS, '\\.word-tab-btn\\.active')).not.toMatch(/font-weight/);
  });
});

/* ------------------------------------------------------------------ */
/* אין גבהים קשיחים בשורה אחת                                         */
/* ------------------------------------------------------------------ */

describe('פקדים שיושבים באותה שורה', () => {
  /** הפקדים שיושבים יחד ב-`.word-group-row` של „גופן”, וחייבים אותו גובה. */
  const ROW_MATES = [
    'ui/ribbon/common/RibbonSelect.vue',
    'ui/ribbon/common/ColorPickerPopover.vue',
  ];

  /** השורש של כל אחד מהם — האלמנט שגובהו *הוא* גובה השורה. */
  const ROW_ROOTS = [
    { file: 'ui/ribbon/common/RibbonSelect.vue', selector: '\\.ribbon-select' },
    { file: 'ui/ribbon/common/ColorPickerPopover.vue', selector: '\\.color-btn-wrapper' },
  ];

  it('כולם נגזרים מ---ribbon-row-h ואינם כותבים מספר', () => {
    const missing = ROW_MATES.filter((file) => !read(file).includes('height: var(--ribbon-row-h)'));
    expect(missing, 'הגובה חייב להיגזר מהטוקן').toEqual([]);

    // הנוכחות של הטוקן אינה מספיקה: 24px שנשאר על **שורש** הפקד הוא בדיוק
    // מה שהשער נבנה למנוע, והבדיקה למעלה עוברת גם אז. מה שבתוך הפקד — פס
    // החיווי, האייקון, המשבצת — הוא קישוט במידה קבועה, ואינו נוגע לשורה.
    const hardcoded = ROW_ROOTS.flatMap(({ file, selector }) =>
      [...block(read(file), selector).matchAll(/(?<![a-z-])height:\s*(\d+)px/g)].map(
        (found) => `${file} ${selector}: height: ${found[1]}px`,
      ),
    );
    expect(hardcoded, 'גובה בפיקסלים על שורש פקד שיושב בשורה משותפת').toEqual([]);
  });

  it('גלריית הסגנונות בגובה תוכן הקבוצה', () => {
    // היא יושבת לצד כפתורים גדולים, ו-68px קשיח היה משאיר אותה נמוכה מהם.
    expect(read('ui/ribbon/common/StyleGallery.vue')).toMatch(
      /height:\s*var\(--ribbon-content-h\)/,
    );
  });
});

/* ------------------------------------------------------------------ */
/* בעלים אחד למחסנית                                                  */
/* ------------------------------------------------------------------ */

describe('המחסנית היא קומפוננטה', () => {
  it('„column-items” אינו מוגדר עוד באף לשונית', () => {
    // הוא היה משוכפל מילה במילה בארבעה בלוקים scoped, והקיבוץ היה מכפיל
    // אותו לשמונה. מחלקה משוכפלת אין לה בעלים.
    const leftovers = TAB_FILES.filter((file) => (TAB_SOURCE.get(file) ?? '').includes('column-items'));
    expect(leftovers).toEqual([]);
  });

  it('כל לשונית שמשתמשת במחסנית גם מייבאת אותה', () => {
    // רכיב שאינו מיובא מרונדר כאלמנט לא-פתור, וזה כבר קרה כאן: תפריט
    // „רשימה” ב„בית” לא היה ב-DOM בכלל (docs/button-audit.md).
    const missing = TAB_FILES.filter((file) => {
      const source = TAB_SOURCE.get(file) ?? '';
      return (
        source.includes('<RibbonStack>') && !source.includes("from '../common/RibbonStack.vue'")
      );
    });
    expect(missing).toEqual([]);
  });
});
