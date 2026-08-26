/**
 * הגבולות שהתכנית קובעת מול SuperDoc (§2, §4) הם החלטות ארכיטקטורה, ולא
 * העדפת סגנון: חריגה מהם מחזירה את התוסף הישן (עריכת DOM ידנית) או מפרה את
 * רישיון המנוע (import ישיר אליו). לכן הם נבדקים על המקור עצמו, בקובץ אחד
 * שנכשל ברור, ולא נסמכים על זכירה בזמן code review.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { join, relative, sep } from 'node:path';

// vitest רץ משורש המאגר, ולכן cwd הוא השורש.
const SRC = join(process.cwd(), 'src');

function sourceFiles(dir = SRC): string[] {
  const files: string[] = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) files.push(...sourceFiles(full));
    else if (/\.(ts|vue)$/.test(entry.name) && !entry.name.endsWith('.d.ts')) files.push(full);
  }
  return files;
}

/**
 * הערות מוסרות לפני הבדיקה: התיעוד בקוד מסביר במפורש מה אסור (למשל
 * "אין לקרוא ל-createSuperDocUI"), וההסבר הזה אינו הפרה. השורות נשמרות
 * כדי שמספרי השורות בכשל יישארו נכונים.
 */
function stripComments(text: string): string {
  return text
    .replace(/\/\*[\s\S]*?\*\//g, (block) => block.replace(/[^\n]/g, ' '))
    .replace(/\/\/.*$/gm, '');
}

const sources = sourceFiles().map((path) => ({
  path: relative(SRC, path),
  text: stripComments(readFileSync(path, 'utf8')),
}));

/** התאמות בקובץ, בפורמט "נתיב:שורה" — כדי שכשל יצביע למקום ולא רק לכלל. */
function hits(pattern: RegExp): string[] {
  const found: string[] = [];
  for (const { path, text } of sources) {
    text.split('\n').forEach((line, index) => {
      if (pattern.test(line)) found.push(`${path}:${index + 1}`);
    });
  }
  return found;
}

describe('גבולות מול SuperDoc', () => {
  it('יש קבצי מקור לבדוק', () => {
    expect(sources.length).toBeGreaterThan(0);
  });

  it('אין import ישיר למנוע ה-DOCX', () => {
    // רישיון המנוע מתיר אותו "solely as a dependency of SuperDoc".
    expect(hits(/from\s+['"]@superdoc\/docx-engine/)).toEqual([]);
  });

  it('אין import מנתיב פנימי של החבילה', () => {
    // רק ה-exports הציבוריים: superdoc, superdoc/ui, superdoc/style.css.
    expect(hits(/from\s+['"]superdoc\/(?!ui['"]|style\.css['"])/)).toEqual([]);
  });

  it('אין יצירת controller שני ל-UI', () => {
    // ה-controller יושב ב-superdoc.ui ובבעלות המופע. ראו create-editor.ts.
    expect(hits(/createSuperDocUI/)).toEqual([]);
  });

  it('אין עריכת מסמך דרך ה-DOM', () => {
    expect(hits(/execCommand|contentEditable|contenteditable/)).toEqual([]);
  });

  it('אין selector אל ה-DOM הפנימי של SuperDoc', () => {
    expect(hits(/(querySelector|querySelectorAll|closest)\s*\(\s*['"][^'"]*(\.sd-|superdoc)/i)).toEqual(
      [],
    );
  });
});

/**
 * החריגה היחידה מהכלל שמעליו, ומנוסחת כאן במפורש ולא נשענת על כך שה-regex
 * שלמעלה מחפש `.sd-`/`superdoc` ואינו תופס תכונות `data-`.
 *
 * מה שמצדיק אותה: שכבת הכותרות שהמנוע מצייר („Different First Page”, „Header
 * from Top”) אינה חלק מ-`ui`, אין לה מתג ב-`modules.surfaces` ואין לה הגדרת
 * טקסטים — והרישיון אוסר לשנות את האריזה. עברות מבחוץ היא הדרך היחידה, וכל
 * מה שהיא עושה הוא להחליף תוויות תצוגה.
 *
 * מה שהחריגה **אינה** מתירה, וזה מה שנמדד: מקום שני שנוגע באותה שכבה, וכתיבה
 * שאינה תווית. עריכת תוכן דרך ה-DOM היא בדיוק התוסף הישן שהתכנית באה להחליף.
 */
describe('עברות שכבת הכותרות', () => {
  const LOCALIZER = 'engine/hf-chrome.ts';

  /** הנתיבים ב-hits הם של המערכת; ההשוואה חייבת להיות אחידה. */
  function normalize(path: string): string {
    return path.split(sep).join('/');
  }

  it('רק hf-chrome.ts נוגע בעיגונים של השכבה', () => {
    const offenders = hits(/data-sd-h(?:f|eader-footer)-/).filter(
      (hit) => !normalize(hit).startsWith(LOCALIZER),
    );

    expect(offenders).toEqual([]);
  });

  it('העברות מחליפה תוויות בלבד — אינה בונה, מוחקת או מזיזה DOM', () => {
    const localizer = sources.find(({ path }) => normalize(path) === LOCALIZER);
    expect(localizer, LOCALIZER).toBeDefined();

    const source = localizer?.text ?? '';
    // מה שמותר: textContent ו-setAttribute. כל השאר הוא כבר עריכת DOM.
    expect(source).toMatch(/textContent/);
    expect(/innerHTML|insertAdjacent|appendChild|removeChild|createElement|\.remove\(/.test(source)).toBe(
      false,
    );
  });
});

/**
 * החריגה השנייה: מדידת המלבן של העמוד המצויר.
 *
 * מה שמצדיק אותה: לסרגל אין ערך אם הוא אינו מיושר לטקסט, ואין API ציבורי
 * שמחזיר את המלבן של העמוד על המסך. החישוב החלופי — רוחב עמוד כפול זום,
 * ממורכז במיכל — נמדד מול המנוע ונמצא שגוי בכל זום שאינו 100%, מפני שהמנוע
 * מיישם זום ב-`transform: scale()` על wrapper רחב מהתוכן הנראה. הפירוט,
 * כולל המספרים שנמדדו, בהערת הפתיחה של engine/page-ruler.ts.
 *
 * מה שהחריגה **אינה** מתירה, וזה מה שנמדד כאן: מקום שני שנוגע באותו עיגון,
 * וכל דבר שאינו קריאת גיאומטריה. `getBoundingClientRect` הוא קריאה טהורה;
 * בנייה, מחיקה או כתיבה אל תוך ה-DOM של המנוע היא כבר התוסף הישן.
 */
describe('מדידת העמוד המצויר', () => {
  const MEASURER = 'engine/page-ruler.ts';

  function normalize(path: string): string {
    return path.split(sep).join('/');
  }

  it('רק page-ruler.ts נוגע בעיגון של העמוד', () => {
    const offenders = hits(/data-page-index/).filter(
      (hit) => !normalize(hit).startsWith(MEASURER),
    );

    expect(offenders).toEqual([]);
  });

  it('המדידה קוראת גיאומטריה בלבד — אינה בונה, מוחקת או כותבת', () => {
    const measurer = sources.find(({ path }) => normalize(path) === MEASURER);
    expect(measurer, MEASURER).toBeDefined();

    const source = measurer?.text ?? '';
    expect(source).toMatch(/getBoundingClientRect/);
    expect(
      /innerHTML|textContent|setAttribute|insertAdjacent|appendChild|removeChild|createElement/.test(
        source,
      ),
    ).toBe(false);
  });
});
