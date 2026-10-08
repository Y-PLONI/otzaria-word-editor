/**
 * שער היגיינה על מקור ה-CSS.
 *
 * למה בדיקה ולא עין: הקוד המת כאן נצבר פעמיים בלי שאיש שם לב. b2f0635 החליף
 * צרכנים של טוקנים ב---color-* ישירים והשאיר חמישה טוקני --word-* בלי צרכן,
 * וכפתור ה-launcher שהועבר להערה ב-RibbonGroup.vue השאיר אחריו שני בלוקים
 * ב-ribbon.css שסלקטור שלהם לא מתאים לשום אלמנט. שני הדברים אינם נראים
 * בדפדפן — הם נראים רק בספירה.
 *
 * מה נמדד:
 *   1. כל custom property שמוגדר במקור, ואינו חלק מה-palette של ה-SDK, יש לו
 *      לפחות צרכן אחד.
 *   2. כל מחלקה שמופיעה בסלקטור בגלובלים (src/styles/*.css) קיימת באמת
 *      בקומפוננטה או בקוד.
 *   3. כל var(--word-*) שנצרך אכן מוגדר — טוקן שלא הוגדר פשוט לא צובע כלום.
 *   4. אין `:has()` בסלקטור שלא אושר במפורש. ראו „:has() בסלקטורים” למטה —
 *      זה אינו סגנון אלא ביצועים שנמדדו.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';

/** vitest רץ משורש המאגר, ולכן src/ נמצא ביחס ל-cwd. */
const SRC = join(process.cwd(), 'src');

function walk(dir: string, out: string[] = []): string[] {
  for (const entry of readdirSync(dir)) {
    const path = join(dir, entry);
    if (statSync(path).isDirectory()) walk(path, out);
    else out.push(path);
  }
  return out;
}

const ALL_FILES = walk(SRC);
const STYLE_SHEETS = ALL_FILES.filter((f) => f.endsWith('.css'));
const CODE_FILES = ALL_FILES.filter((f) => f.endsWith('.vue') || f.endsWith('.ts'));

/** קובץ → תוכן, פעם אחת: הבדיקות למטה סורקות את אותם קבצים. */
const CONTENT = new Map(ALL_FILES.map((f) => [f, readFileSync(f, 'utf8')]));
const ALL_SOURCE = [...CONTENT.values()].join('\n');

function short(path: string): string {
  return path.slice(SRC.length - 3);
}

/**
 * ההגדרות בלבד: `--x:` בתחילת שורה. העיגון לתחילת השורה הוא מה שמפריד הגדרה
 * מאזכור בתוך var(--x, fallback) ומקריאת setProperty ב-TypeScript.
 */
function definitions(source: string): string[] {
  return [...source.matchAll(/^\s*(--[\w-]+)\s*:/gm)].map((m) => m[1]);
}

function references(source: string): string[] {
  return [...source.matchAll(/var\(\s*(--[\w-]+)/g)].map((m) => m[1]);
}

describe('טוקני CSS', () => {
  it('לכל טוקן שהמצאנו יש צרכן', () => {
    // --color-* יוצאים מהכלל: הם ה-palette המתועד של ה-SDK, host/theme.ts כותב
    // אותם, והם החוזה מול אוצריא — גם תפקיד שאיננו צובעים בו כלום כרגע נשאר.
    const used = new Set(references(ALL_SOURCE));
    const orphans: string[] = [];

    for (const file of ALL_FILES) {
      for (const token of definitions(CONTENT.get(file) ?? '')) {
        if (token.startsWith('--color-')) continue;
        if (!used.has(token)) orphans.push(`${token} (${short(file)})`);
      }
    }

    expect(orphans).toEqual([]);
  });

  it('כל טוקן --word-* שנצרך גם מוגדר', () => {
    const defined = new Set(definitions(ALL_SOURCE));
    const undeclared = [
      ...new Set(references(ALL_SOURCE).filter((t) => t.startsWith('--word-'))),
    ].filter((token) => !defined.has(token));

    expect(undeclared).toEqual([]);
  });
});

describe('סלקטורים בגלובלים', () => {
  /**
   * שם מחלקה שנבנה בזמן ריצה מתבנית — `btn-${variant}` ב-RibbonButton.vue —
   * אינו מופיע במקור כמחרוזת שלמה. מחפשים את התחילית שלפני החלק הדינמי.
   */
  function isComposed(name: string): boolean {
    const parts = name.split('-');
    for (let i = parts.length - 1; i > 0; i -= 1) {
      if (ALL_SOURCE.includes(`\`${parts.slice(0, i).join('-')}-\${`)) return true;
    }
    return false;
  }

  /**
   * מחלקות שה-DOM שלהן שייך למנוע ולא לנו. `.superdoc` הוא ה-wrapper שהמנוע
   * מרנדר, והכלל היחיד שנוגע בו — מרכוז העמוד ב-shell.css — מתועד שם במלואו.
   * `.sd-v2-local-selection-caret` הוא סמן העריכה שהמנוע מייצר, ואנימציית
   * ההבהוב שלו ב-shell.css והסתרתו ב-print.css מתועדות בחוזה caret-blink.
   * ההחרגה מפורשת ולא „עוברת בטעות” מפני ש-'superdoc' מופיע גם כשם החבילה
   * ב-import: בלעדיה הבדיקה הייתה מאשרת אותו מסיבה לא נכונה.
   */
  const ENGINE_OWNED = new Set([
    'superdoc',
    'sd-v2-local-selection-caret',
  ]);

  it('כל מחלקה בסלקטור קיימת בקומפוננטה או בקוד', () => {
    // רק הגלובלים: סגנונות scoped בתוך .vue הם של הקומפוננטה עצמה, ושם
    // הסלקטור והתבנית יושבים באותו קובץ ומתוחזקים יחד.
    const componentSource = CODE_FILES.map((f) => CONTENT.get(f) ?? '').join('\n');
    const dead: string[] = [];

    for (const sheet of STYLE_SHEETS) {
      // הערות מוסרות תחילה: הן יושבות לפני '{' ולכן נקראות כחלק מהסלקטור.
      // הערה שמזכירה מחלקה של המנוע (`superdoc__layers`) דיווחה עליה כמחלקה
      // מדומה, וזו הייתה בדיקה שנכשלת על תיעוד.
      const css = (CONTENT.get(sheet) ?? '').replace(/\/\*[\s\S]*?\*\//g, ' ');
      for (const block of css.matchAll(/(?:^|\})([^{}]*)\{/g)) {
        for (const cls of block[1].matchAll(/\.(-?[_a-zA-Z][\w-]*)/g)) {
          const name = cls[1];
          if (ENGINE_OWNED.has(name) || isComposed(name)) continue;
          if (componentSource.includes(name)) continue;
          dead.push(`${name} (${short(sheet)})`);
        }
      }
    }

    expect([...new Set(dead)]).toEqual([]);
  });
});

/* ------------------------------------------------------------------ */
/* :has() בסלקטורים                                                    */
/* ------------------------------------------------------------------ */

/**
 * `:has()` שהעוגן שלו יכול לשבת בתוך עץ המסמך של המנוע (`.superdoc`) הוא
 * מוקש ביצועים שנמדד ולא הונח: כלל אחד כזה ב-engine-chrome.css —
 * `.superdoc__mutation-status:has([data-superdoc-v2-edit-rejected])` — גרם
 * ל-Blink לסמן את `.superdoc` כמושפע מ-`:has()`, ומרגע זה כל תו שנקלד תזמן
 * חישוב סגנון מחדש של כל עץ המסמך: 42 חישובים של 2,748 אלמנטים (60–115ms
 * כל אחד) ב-50 הקשות, ו-53 long tasks. מחיקת הכלל לבדה הורידה את שניהם
 * ל-0 ול-3; מחיקת כלל אחר לא שינתה דבר. הסיפור המלא בהערה על באנר
 * `edit-rejected` ב-engine-chrome.css (הכלל עצמו נמחק כשהמנוע הפסיק לצייר את
 * הבאנר), והמדידה החוזרת ב-scripts/qa/typing-style-recalc-qa.mjs.
 *
 * הכלל כאן גס בכוונה: **כל** `:has()` בגיליונות ובקומפוננטות חייב להופיע
 * ברשימה למטה עם הנימוק למה העוגן שלו אינו יכול להיות אב של צומת במסמך.
 * מי שמוסיף `:has()` חדש נדרש להחליט, לא לגלוש.
 */
describe(':has() בסלקטורים', () => {
  /**
   * `.word-ribbon-group:has(+ .word-ribbon-group--end)` — העוגן הוא קבוצה
   * ברצועה, מחוץ ל-`.editor-stack`, ולכן שינוי DOM במסמך אינו עובר דרכו.
   * נמדד באותו ניסוי: מחיקת הכלל הזה בזמן ריצה לא שינתה את מספר החישובים
   * המלאים (44 מול 42 בבסיס), בעוד מחיקת כלל הבאנר הורידה אותם ל-0.
   *
   * `.shell-top :deep(.word-ribbon-container:has(> .word-ribbon-body[style*='display: none']) .word-tab-bar)`
   * — העוגן הוא `.word-ribbon-container` שנמצא ב-`.shell-top`, מחוץ לחלוטין
   * ל-`.editor-stack`. הכלל בודק אם גוף הרצועה מוסתר (v-show), ומעגל את פינות
   * שורת הלשוניות בהתאם — שינוי DOM בתוך המסמך אינו עובר דרכו.
   */
  const ALLOWED = new Set([
    '.word-ribbon-group:has(+ .word-ribbon-group--end)',
    ".shell-top :deep(.word-ribbon-container:has(> .word-ribbon-body[style*='display: none']) .word-tab-bar)",
  ]);

  /** פיצול רשימת סלקטורים בפסיקים שמחוץ לסוגריים — `:has(a, b)` נשאר שלם. */
  function splitSelectorList(selector: string): string[] {
    const parts: string[] = [];
    let depth = 0;
    let current = '';
    for (const ch of selector) {
      if (ch === '(') depth += 1;
      if (ch === ')') depth -= 1;
      if (ch === ',' && depth === 0) {
        parts.push(current);
        current = '';
        continue;
      }
      current += ch;
    }
    parts.push(current);
    return parts.map((part) => part.replace(/\s+/g, ' ').trim()).filter(Boolean);
  }

  /** כל הסלקטורים עם `:has(` בגיליון או בבלוקי `<style>` של קומפוננטה. */
  function hasSelectors(source: string): string[] {
    // בקומפוננטה רק בלוקי <style> הם CSS; התבנית מכילה '{' של Vue ואינה נסרקת.
    const css = (
      source.includes('<style')
        ? [...source.matchAll(/<style[^>]*>([\s\S]*?)<\/style>/g)].map((m) => m[1]).join('\n')
        : source
    ).replace(/\/\*[\s\S]*?\*\//g, ' ');
    const out: string[] = [];
    for (const block of css.matchAll(/(?:^|\})([^{}]*)\{/g)) {
      const selector = block[1].trim();
      if (!selector.includes(':has(')) continue;
      for (const single of splitSelectorList(selector)) {
        if (single.includes(':has(')) out.push(single);
      }
    }
    return out;
  }

  it('הסורק אכן מזהה :has() (הגנה מפני regex שהפסיק להתאים)', () => {
    expect(hasSelectors('.a:has(> .b) > .c { color: red; }\n.d { }')).toEqual(['.a:has(> .b) > .c']);
    expect(hasSelectors('/* .x:has(.y) בהערה */ .z { }')).toEqual([]);
  });

  it('כל :has() בקוד מאושר במפורש', () => {
    const found: string[] = [];
    for (const file of [...STYLE_SHEETS, ...CODE_FILES.filter((f) => f.endsWith('.vue'))]) {
      for (const selector of hasSelectors(CONTENT.get(file) ?? '')) {
        if (!ALLOWED.has(selector)) found.push(`${selector} (${short(file)})`);
      }
    }
    expect(found).toEqual([]);
  });

  it('הרשימה המאושרת אינה מתה — כל כלל בה עוד קיים', () => {
    const present = new Set<string>();
    for (const file of [...STYLE_SHEETS, ...CODE_FILES]) {
      for (const selector of hasSelectors(CONTENT.get(file) ?? '')) present.add(selector);
    }
    expect([...ALLOWED].filter((selector) => !present.has(selector))).toEqual([]);
  });
});

/* ------------------------------------------------------------------ */
/* שלושת מצבי הפקד                                                    */
/* ------------------------------------------------------------------ */

/**
 * ההגדרות של tokens.css כמפה, כדי שאפשר יהיה לפתור שרשרת של `var()`.
 * `--word-btn-active: var(--color-primary-subtle)` ו-
 * `--word-btn-active: rgba(21, 101, 192, 0.12)` הם אותו צבע בפועל, ושער
 * שמשווה מחרוזות גולמיות היה מפספס את אחת משתי הצורות.
 */
function tokenTable(): Map<string, string> {
  const css = CONTENT.get(join(SRC, 'styles/tokens.css')) ?? '';
  const table = new Map<string, string>();
  for (const m of css.matchAll(/^\s*(--[\w-]+)\s*:\s*([^;]+);/gm)) table.set(m[1], m[2].trim());
  return table;
}

/** פותרת `var(--x)` עד לערך ממשי. מחזירה null על טוקן שאינו מוגדר או על מעגל. */
function resolveToken(name: string, table: Map<string, string>): string | null {
  const seen = new Set<string>();
  let value = table.get(name) ?? null;
  while (value !== null) {
    const ref = /^var\(\s*(--[\w-]+)\s*\)$/.exec(value);
    if (!ref) return value;
    if (seen.has(ref[1])) return null;
    seen.add(ref[1]);
    value = table.get(ref[1]) ?? null;
  }
  return null;
}

describe('שלושת מצבי הפקד ברצועה', () => {
  /**
   * hover, דלוק, ודלוק+עכבר. זו ההבחנה המרכזית ברצועה של Word, וכששתיים מהן
   * נגזרות לאותו ערך הפקד הדלוק נראה בדיוק כמו פקד שהעכבר עובר מעליו —
   * הבאג המקורי כאן. מוטציה שהחזירה אותו (`--word-btn-active` שקיבל את הערך
   * של `--word-btn-hover`) עברה 146 בדיקות בירוק: הבדיקות שהיו מדדו טוקן
   * מיותם וסלקטור מת, ולא שוויון ערכים.
   *
   * הגזירה המקבילה בזמן ריצה נמדדת ב-theme.test.ts — הטוקנים כאן הם ברירת
   * המחדל שלפני `plugin.boot`, ו-`applyTheme` דורס אותם. שני המקומות חייבים
   * שער, כי כל אחד מהם לבדו הוא מסלול שלם שהמשתמש רואה.
   */
  const STATES = ['--word-btn-hover', '--word-btn-active', '--word-btn-active-hover'] as const;

  it('כל אחד מהשלושה מוגדר ונפתר לערך ממשי', () => {
    // בלי זה השער למטה היה עובר בירוק על שלוש פעמים null.
    const table = tokenTable();
    const unresolved = STATES.filter((name) => resolveToken(name, table) === null);
    expect(unresolved).toEqual([]);
  });

  it('שלושת המצבים נבדלים זה מזה בערך שנפתר', () => {
    const table = tokenTable();
    const resolved = STATES.map((name) => `${name} = ${resolveToken(name, table)}`);
    const values = STATES.map((name) => resolveToken(name, table));
    expect(new Set(values).size, resolved.join(' | ')).toBe(STATES.length);
  });
});

/* ------------------------------------------------------------------ */
/* צבע קשיח                                                           */
/* ------------------------------------------------------------------ */

/**
 * ולידציית העיצוב של אוצריא פוסלת צבע hex, `rgb()` או שם צבע באנגלית בתוך
 * CSS ודורשת `var(--color-*)`, והסיבה אינה פורמלית: המשתמש מחליף ערכת צבעים
 * ומצב כהה/בהיר בכל רגע, וצבע שאינו זז איתם נשאר על הרקע הלא נכון. זה נמדד
 * שלוש פעמים באותה סריקה — `#ffffff` על כפתור בצבע primary במצב כהה,
 * `rgba(176, 0, 32, .1)` קפוא לצד `color: var(--color-error)` דינמי באותו
 * כלל, ו-`#ffffff` על `--color-outline` במצב בהיר.
 *
 * הסריקה על **ערכי הצהרות** ולא על הטקסט: `white-space: nowrap` הוא שם
 * מאפיין, ו-`var(--word-blue)` הוא שם טוקן — שניהם היו נתפסים בחיפוש מחרוזת
 * ומכשילים את השער מהסיבה הלא נכונה.
 */
const NAMED_COLORS = [
  'white', 'black', 'red', 'green', 'blue', 'gray', 'grey', 'silver', 'orange', 'yellow',
  'purple', 'navy', 'teal', 'aqua', 'lime', 'maroon', 'olive', 'fuchsia', 'brown', 'pink',
  'gold', 'ivory', 'beige', 'tan', 'cyan', 'magenta', 'rebeccapurple', 'darkgray', 'lightgray',
];

interface Declaration {
  file: string;
  property: string;
  value: string;
}

/** ההצהרות של גלובל או של בלוק `<style>` — אחרי הסרת הערות. */
function declarations(file: string): Declaration[] {
  const source = CONTENT.get(file) ?? '';
  const css = file.endsWith('.css')
    ? source
    : [...source.matchAll(/<style[^>]*>([\s\S]*?)<\/style>/g)].map((m) => m[1]).join('\n');
  const clean = css.replace(/\/\*[\s\S]*?\*\//g, ' ');
  return [...clean.matchAll(/([-\w]+)\s*:\s*([^;{}]+)[;}]/g)].map((m) => ({
    file,
    property: m[1],
    value: m[2].trim(),
  }));
}

/** הצבעים הקשיחים בערך הצהרה. שמות טוקנים מוסרים לפני חיפוש שם צבע. */
function hardcodedColors(value: string): string[] {
  const withoutTokens = value.replace(/--[\w-]+/g, '');
  const found = [
    ...withoutTokens.matchAll(/#[0-9a-fA-F]{3,8}\b/g),
    ...withoutTokens.matchAll(/\b(?:rgb|rgba|hsl|hsla)\([^)]*\)/g),
  ].map((m) => m[0]);
  for (const name of NAMED_COLORS) {
    if (new RegExp(`\\b${name}\\b`, 'i').test(withoutTokens)) found.push(name);
  }
  return found;
}

/**
 * שתי החרגות, ושתיהן מנוסחות כחוק ולא כרשימת קבצים — רשימה הייתה גדלה בשקט:
 *
 * 1. **הגדרות הטוקנים ב-tokens.css.** מדריך העיצוב מחייב ברירות מחדל ל-color
 *    roles לפני `plugin.boot`, ובלעדיהן התוסף חשוף כמה מילישניות בלי צבע
 *    בכלל. זה המקום היחיד בפרויקט שבו hex הוא הדבר הנכון, ולכן ההחרגה מוגבלת
 *    לו ולהגדרות של custom properties בתוכו.
 * 2. **צל בשחור שקוף.** `box-shadow: ... rgba(0, 0, 0, α)` הוא הדפוס של
 *    המדריך עצמו (ראו „פאנל Overlay צף” ו„פופאובר ממוקם”), והוא ניטרלי: שחור
 *    בשקיפות נמוכה מכהה את מה שמתחתיו בשני המצבים. צל בכל צבע **אחר** אינו
 *    מוחרג — הוא צבע מותג שקפא בצל.
 *
 * הפלטה של בורר הצבעים אינה נזכרת כאן במכוון: `THEME_COLUMNS`/
 * `STANDARD_COLORS` הם נתונים ב-`<script>` ולא CSS, ולכן הסריקה אינה נוגעת
 * בהם מלכתחילה. הם פלטת Office — הצבע *הוא* המידע, ואין לו לזוז עם הערכה.
 */
const TOKENS_CSS = join(SRC, 'styles/tokens.css');
const NEUTRAL_SHADOW = /^rgba\(\s*0\s*,\s*0\s*,\s*0\s*,\s*(?:0?\.\d+|0|1)\s*\)$/;

function isAllowed(declaration: Declaration, color: string): boolean {
  if (declaration.file === TOKENS_CSS && declaration.property.startsWith('--')) return true;
  const isShadow = declaration.property === 'box-shadow' || declaration.property === 'text-shadow';
  return isShadow && NEUTRAL_SHADOW.test(color);
}

describe('צבע קשיח ב-CSS', () => {
  const SCANNED = [...STYLE_SHEETS, ...ALL_FILES.filter((f) => f.endsWith('.vue'))];

  it('הסריקה אכן קוראת הצהרות (הגנה מפני regex שהפסיק להתאים)', () => {
    // שער שאינו מודד כלום עובר בירוק על כל מוטציה, וזו בדיוק התקלה שהשער הזה
    // נבנה בגללה.
    const all = SCANNED.flatMap(declarations);
    expect(all.length).toBeGreaterThan(400);
    expect(all.some((d) => d.property === 'background' && d.value.includes('var(--color-'))).toBe(
      true,
    );
    // ההחרגות עצמן חייבות להיות בשימוש, אחרת הן טענה על העבר.
    const tokenDefaults = declarations(TOKENS_CSS).filter(
      (d) => d.property.startsWith('--') && hardcodedColors(d.value).length > 0,
    );
    expect(tokenDefaults.length).toBeGreaterThan(10);
  });

  it('אין צבע קשיח מחוץ לשתי ההחרגות', () => {
    const offenders: string[] = [];
    for (const file of SCANNED) {
      for (const declaration of declarations(file)) {
        for (const color of hardcodedColors(declaration.value)) {
          if (isAllowed(declaration, color)) continue;
          offenders.push(`${short(file)} — ${declaration.property}: ${declaration.value}`);
        }
      }
    }
    expect([...new Set(offenders)]).toEqual([]);
  });

  it('הגלאי מזהה את שלושת הצבעים שנמדדו בפועל', () => {
    // בלי זה אין הוכחה שהשער היה אדום על מה שהוא נבנה בגללו.
    expect(hardcodedColors('#ffffff')).toEqual(['#ffffff']);
    expect(hardcodedColors('rgba(176, 0, 32, 0.1)')).toEqual(['rgba(176, 0, 32, 0.1)']);
    expect(hardcodedColors('#e67e22')).toEqual(['#e67e22']);
    // ומה שאסור לו להיות אדום: שם מאפיין, שם טוקן, ומילת מפתח.
    expect(hardcodedColors('nowrap')).toEqual([]);
    expect(hardcodedColors('var(--word-blue)')).toEqual([]);
    expect(hardcodedColors('1px solid var(--color-outline)')).toEqual([]);
    expect(hardcodedColors('transparent')).toEqual([]);
    expect(hardcodedColors('currentColor')).toEqual([]);
  });

  it('צל שאינו שחור ניטרלי אינו מוחרג', () => {
    const shadow = { file: 'x.css', property: 'box-shadow', value: '0 1px 4px #1565c0' };
    expect(isAllowed(shadow, '#1565c0')).toBe(false);
    expect(isAllowed({ ...shadow, value: '0 1px 4px rgba(0, 0, 0, 0.08)' }, 'rgba(0, 0, 0, 0.08)')).toBe(
      true,
    );
  });
});
