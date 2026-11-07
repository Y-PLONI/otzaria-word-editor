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

  it('אין selector פנימי מחוץ לקריאת מקטעי הגוף בקורא המורשה', () => {
    // החריגה של painted-lines מתירה קריאה בלבד (נאכף למטה). למחלקת
    // fragment צריך להבחין בין בלוק גוף לבין פסקה בתוך תא: לטבלה אין
    // source-node-id, ולפסקת התא אין pm-start. סינון לפי שניהם מדלג
    // בטעות אל הפסקה שאחרי הטבלה. רק selector הגוף המדויק מותר כאן.
    const bodyFragments = hits(/querySelectorAll\s*\(\s*'\.superdoc-fragment\[data-layout-story="body"\]'\s*\)/)
      .filter((hit) => hit.startsWith(`engine${sep}painted-lines.ts:`));
    const offenders = hits(/(querySelector|querySelectorAll|closest)\s*\(\s*['"][^'"]*(\.sd-|superdoc)/i)
      .filter((hit) => !bodyFragments.includes(hit) && !hit.startsWith(`engine${sep}justify-last-line.ts:`));
    expect(offenders).toEqual([]);
  });

  /**
   * שכבת החבילה נשארת בלי תלויות. כשהיא ישבה בתוך `docx-preflight.ts`, חיבור
   * התיקון היוצא ל-`export.ts` יצר את השרשרת `vite.config.ts` →
   * `blank-document.ts` → `export.ts` → preflight → `vba-import.ts` (ESM
   * בלבד), וטעינת קובץ התצורה נשברה. import אחד שנוסף שם מחזיר אותה.
   *
   * **שלושה ענפים, מפני ש„תלות” אינה רק `import`.** הצורה הקודמת
   * (`/^\s*import\s|\brequire\s*\(/`) נמדדה על שתים-עשרה צורות: היא תפסה
   * שש — ‏`import { x } from`, ‏`import type`, ‏`import './x'`,
   * ‏`import * as`, import מוזח ו-`require(` — והחמיצה שש:
   * ‏`export { X } from './y'`, ‏`export * from './y'`,
   * ‏`export type { X } from './y'`, ‏`import('./y')` דינמי (עם `await`
   * ובלעדיו), ‏`import{x}from'./y'` בלי רווחים, ו-import שנשבר לשתי שורות.
   *
   * החמצה ראשונה היא החמורה: ‏`export { Bytes } from './types'` יוצר בדיוק
   * את קשת התלות שהכלל קיים כדי למנוע, ו-`import()` דינמי הוא בדיוק הדרך
   * שבה מישהו „יפתור” את שרשרת ה-ESM שלמעלה.
   *
   * לכן הענף האמצעי אינו מחפש `import`/`export` אלא את **מפרט המודול**
   * עצמו (`from '…'`), שקיים בכל צורת ייבוא וייצוא־מחדש ובכל פיצול שורות;
   * הראשון סוגר ייבוא-לוואי שאין בו `from`, והשלישי את שתי צורות הקריאה.
   */
  const DEPENDENCY = /^\s*import\b|\bfrom\s+['"]|\b(?:require|import)\s*\(/;

  it('docx-parts.ts אינו מייבא דבר', () => {
    const parts = sources.find(({ path }) => path.split(sep).join('/') === 'engine/docx-parts.ts');
    expect(parts, 'engine/docx-parts.ts').toBeDefined();
    const imports = (parts?.text ?? '')
      .split('\n')
      .flatMap((line, index) => (DEPENDENCY.test(line) ? [`docx-parts.ts:${index + 1}`] : []));
    expect(imports).toEqual([]);
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

/**
 * החריגה השלישית: טווח ה-pm של השורה המצוירת.
 *
 * מה שמצדיק אותה: שלוש תקלות סמן שהמנוע אינו חושף דרכן שום API. ‏`End` בשורה
 * עברית מעביר את הסמן לתחילתה במקום לסופה, החצים האופקיים זזים לכיוון ההפוך
 * מהמקש בכל פסקה שיש בה רשימה, טאב, ספרות או לטינית, והחצים האנכיים מחליפים
 * את שני קצות השורה שמתחת כשעמודת המטרה נופלת מחוצה לה. לשלושתן צריך לדעת
 * **איפה בדיוק צויר** כל דבר — שורה, ריצה, תו — וזו תוצאה של פריסה:
 * `readMountedLayoutData` מתאר שורות ב-`fromRun`/`toChar` ואינו נותן את אורכי
 * הריצות, כלומר אי אפשר להמיר אותם להיסט בלי לנחש. המדידות וההצדקה בהערות
 * הפתיחה של `engine/rtl-line-end.ts`, `engine/rtl-caret.ts` ו-
 * `engine/rtl-vertical-caret.ts`.
 *
 * ומה שהכלל הזה **אוכף**, וזה העיקר: הקריאה עצמה נמצאת במקום **אחד**.
 * ‏`engine/painted-lines.ts` הוא הקורא, ושלושת המודולים שמעליו מחזיקים
 * מדיניות בלבד — איזה מקש עושה מה עם מה שנקרא. שלושה קוראים עצמאיים היו
 * שלושה עותקים של כללים שכל אחד מהם נמדד מול המנוע, ומתגלים שוב דרך באג
 * כשמנחשים אותם מחדש.
 *
 * ‏`engine/rtl-line-end.ts` נשאר קורא שני, בכוונה: הוא קדם לשניים האחרים,
 * מחזיק קורא מצומצם משלו, יציב ומגודר בשער — והסבה שלו היא סיכון בלי תמורה.
 *
 * מה שהחריגה **אינה** מתירה, וזה מה שנמדד כאן: מקום שלישי שנוגע באותם
 * עיגונים, וכל דבר שאינו קריאה. הזזת הסמן עצמה נעשית דרך ה-API של המנוע
 * (`authoring.setSelectionTarget`) ולא דרך ה-DOM.
 */
describe('טווח ה-pm של השורה המצוירת', () => {
  const READERS = ['engine/painted-lines.ts', 'engine/rtl-line-end.ts'];

  function normalize(path: string): string {
    return path.split(sep).join('/');
  }

  it('רק painted-lines.ts ו-rtl-line-end.ts נוגעים בעיגונים של השורה', () => {
    const offenders = hits(/data-pm-(?:start|end)|data-source-node-id/).filter(
      (hit) => !READERS.some((reader) => normalize(hit).startsWith(reader))
        && !normalize(hit).startsWith('engine/justify-last-line.ts:'),
    );

    expect(offenders).toEqual([]);
  });

  // #4136 needs paragraph identity to verify justification through the public
  // API. This narrowly scoped painter writes only its own word-spacing.
  it('last-line painting reads identity but never writes document content or pm anchors', () => {
    const source = sources.find((file) => normalize(file.path) === 'engine/justify-last-line.ts')!.text;
    expect(source).not.toMatch(/data-pm-(?:start|end)|innerHTML|textContent\s*=|appendChild|removeChild/);
    expect(source).not.toMatch(/\.style\.(?!wordSpacing\b)\w+\s*=/);
  });

  it.each(READERS)('%s קורא בלבד — אינו בונה, מוחק או כותב', (path) => {
    const reader = sources.find((file) => normalize(file.path) === path);
    expect(reader, path).toBeDefined();

    const source = reader?.text ?? '';
    expect(source).toMatch(/getAttribute/);
    expect(
      /innerHTML|textContent|setAttribute|insertAdjacent|appendChild|removeChild|createElement/.test(
        source,
      ),
    ).toBe(false);
  });
});
