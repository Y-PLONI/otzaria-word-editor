/**
 * ערכת הנושא של אוצריא → CSS variables.
 *
 * שמות המשתנים והמיפוי הם אלה שמדריך העיצוב של ה-SDK מחייב
 * (docs/plugin-sdk/DESIGN_GUIDE.md): ולידציית העיצוב של הוולידטור פוסלת צבע
 * hex, rgb() או שם צבע באנגלית בתוך CSS, ודורשת var(--color-*). כל צבע ב-UI
 * מגיע מכאן — המשתמש יכול להחליף ערכת צבעים ומצב כהה/בהיר בכל רגע, וה-UI
 * חייב להשתנות איתו בלי רענון.
 */
import type { ColorScheme, ThemePayload } from '../types/otzaria_plugin';

const COLOR_VARS: ReadonlyArray<readonly [string, keyof ColorScheme]> = [
  ['--color-primary', 'primary'],
  ['--color-on-primary', 'onPrimary'],
  ['--color-secondary', 'secondary'],
  ['--color-on-secondary', 'onSecondary'],
  ['--color-secondary-container', 'secondaryContainer'],
  ['--color-on-secondary-container', 'onSecondaryContainer'],
  ['--color-surface', 'surface'],
  ['--color-on-surface', 'onSurface'],
  ['--color-on-surface-variant', 'onSurfaceVariant'],
  ['--color-surface-container-high', 'surfaceContainerHigh'],
  ['--color-surface-container-highest', 'surfaceContainerHighest'],
  ['--color-error', 'error'],
  ['--color-on-error', 'onError'],
  ['--color-outline', 'outline'],
];

/** `#rrggbb` או `#rgb` → שלושת הערוצים, או null על כל דבר אחר. */
function parseHex(hex: string): [number, number, number] | null {
  const value = hex.trim().replace(/^#/, '');
  const full =
    value.length === 3
      ? value
          .split('')
          .map((c) => c + c)
          .join('')
      : value;

  if (!/^[0-9a-f]{6}$/i.test(full)) return null;

  return [
    parseInt(full.slice(0, 2), 16),
    parseInt(full.slice(2, 4), 16),
    parseInt(full.slice(4, 6), 16),
  ];
}

/** `#rrggbb` או `#rgb` → rgba(). מחזירה null על כל דבר אחר. */
export function hexToRgba(hex: string, alpha: number): string | null {
  const rgb = parseHex(hex);
  if (!rgb) return null;
  return `rgba(${rgb[0]}, ${rgb[1]}, ${rgb[2]}, ${alpha})`;
}

/**
 * מיזוג שני צבעי hex לצבע אטום — `base` שנדחף אל `overlay` בשיעור `ratio`.
 *
 * למה אטום ולא rgba: הצריכה היחידה היא רקע ה-hover של כפתור ממולא, שכבר צבוע
 * ב-primary. שכבה שקופה מעליו מתמזגת עם ה-primary שמתחת ובקושי נראית, ולכן כאן
 * צריך צבע מחושב. הדחיפה היא אל צבע הטקסט של המשטח: במצב בהיר זה כהה יותר,
 * במצב כהה בהיר יותר — כלומר ההיענות ל-hover נשארת נראית בשני המצבים.
 */
export function blendHex(base: string, overlay: string, ratio: number): string | null {
  const a = parseHex(base);
  const b = parseHex(overlay);
  if (!a || !b) return null;

  const mix = (i: number): number => Math.round(a[i] * (1 - ratio) + b[i] * ratio);
  return `rgb(${mix(0)}, ${mix(1)}, ${mix(2)})`;
}

export function applyTheme(theme: ThemePayload): void {
  const root = document.documentElement;
  const colors: Partial<ColorScheme> = theme.colorScheme ?? {};

  for (const [name, key] of COLOR_VARS) {
    const value = colors[key];
    // ערך חסר משאיר את ברירת המחדל שב-tokens.css במקום לאפס את המשתנה.
    if (typeof value === 'string' && value !== '') root.style.setProperty(name, value);
  }

  // surfaceContainerHigh הוא הרקע שהמדריך מחייב לפס העליון, אבל הוא נוסף
  // ב-SDK 1.1.0 ואופציונלי בטיפוס. בגרסה שלא מחזירה אותו נופלים ל-highest,
  // כדי שהפס לא ייראה כמו גוף המסמך.
  if (!colors.surfaceContainerHigh && colors.surfaceContainerHighest) {
    root.style.setProperty('--color-surface-container-high', colors.surfaceContainerHighest);
  }

  // סולם הדרגות השקופות. שלוש דרגות מ-primary ולא אחת: hover ומצב דלוק שנגזרו
  // לאותו ערך הם בדיוק הרגרסיה שהייתה כאן — ברצועה של Word ההבחנה בין „הכפתור
  // דלוק” ל„העכבר עובר מעליו” היא ההבחנה המרכזית, ומסגרת לבדה אינה מספיקה.
  // 8%/12% הם הערכים של DESIGN_GUIDE.md; 20% הוא דלוק+עכבר, שחייב להעמיק
  // מעל הדלוק. הגזירה כאן ולא ב-color-mix — ל-WebView2 שאוצריא מריצה לא
  // מובטחת תמיכה, וגזירה ב-JS היא הדפוס שכל הצבעים כאן כבר עוברים בו.
  const ALPHA_SHADES: ReadonlyArray<readonly [string, keyof ColorScheme, number]> = [
    ['--color-primary-hover', 'primary', 0.08],
    ['--color-primary-subtle', 'primary', 0.12],
    ['--color-primary-selected-hover', 'primary', 0.2],
    ['--color-secondary-subtle', 'secondary', 0.12],
    ['--color-error-subtle', 'error', 0.12],
  ];

  for (const [name, key, alpha] of ALPHA_SHADES) {
    const base = colors[key];
    const shade = typeof base === 'string' ? hexToRgba(base, alpha) : null;
    if (shade) root.style.setProperty(name, shade);
  }

  // ה-hover של כפתור ממולא — ראו blendHex.
  if (typeof colors.primary === 'string' && typeof colors.onSurface === 'string') {
    const filledHover = blendHex(colors.primary, colors.onSurface, 0.2);
    if (filledHover) root.style.setProperty('--color-primary-filled-hover', filledHover);
  }

  // גופן הקריאה של אוצריא (`typography.fontFamily`) אינו מוחל על הממשק. שתי
  // סיבות נצפו יחד בכפתורי הרצועה: אוצריא מזריקה למשפחה `@font-face` אחד בלי
  // דסקריפטור `font-weight` כלל, ולכן כל בולד בממשק עבר סינתוז של הדפדפן —
  // ציור מרוח ולא ציור האות הבולד (מתוקן בצד אוצריא ב-otzaria#986); ובנוסף
  // גופן ספרים בן 12px מצייר את קוויו הדקים דקים מפיקסל, והמסך צובע אותם אפור.
  // `--font-main` נשאר Assistant מ-tokens.css: נארז איתנו בארבעה משקלים
  // אמיתיים, ולכן גם הבולד שלו אמיתי, והוא sans שנשאר חד בקטן.
  const typography = theme.typography;
  if (typeof typography?.fontSize === 'number' && typography.fontSize > 0) {
    root.style.setProperty('--font-size-base', `${typography.fontSize}px`);
  }
  if (typeof typography?.lineHeight === 'number' && typography.lineHeight > 0) {
    root.style.setProperty('--line-height', String(typography.lineHeight));
  }

  if (theme.mode) root.dataset.theme = theme.mode;
  document.body?.classList.toggle('dark-mode', theme.mode === 'dark');
}
