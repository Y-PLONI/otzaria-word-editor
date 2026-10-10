/**
 * צבע הבד — המשטח שסביב הדף.
 *
 * ## למה טוקן משלו ולא דריסה של `--color-surface-container-highest`
 *
 * הצבע שהבד נצבע בו עד עכשיו הוא טוקן של ערכת הנושא, ואותו טוקן בדיוק צובע
 * עוד ארבעה דברים: שני הסרגלים (DocumentRuler.vue, VerticalRuler.vue), הפינה
 * שביניהם וכפתור היציאה ממצב מיקוד. דריסה שלו הייתה מקבלת „צבע רקע” שמשנה
 * גם את הסרגלים — כלומר פקד שעושה יותר ממה ששמו אומר. `--word-canvas-bg`
 * (styles/tokens.css) יושב בין השניים, והצרכן היחיד שלו הוא `.editor-stack` —
 * הבד עצמו.
 *
 * ## ברירת המחדל: גוון של אוצריא, בהיר משורת טאבי המסמכים
 *
 * בלי העדפה הבד נצבע ב-`--word-shell-band-bg` — אותו צבע כמו פס הכותרת
 * ושורת הלשוניות של הרצועה: `surfaceContainerHigh` של אוצריא (הצבע של שורת
 * טאבי המסמכים) מעורבב עם `surface`, כלומר בהיר ממנה ונגזר מערכת הנושא בשני
 * המצבים. הפס בבורר מראה את אותו צבע — `canvasDefaultColor` למטה.
 *
 * ## למה סגנון inline על שורש המסמך
 *
 * זה בדיוק מה ש-`applyTheme` כבר עושה (host/theme.ts): צבעי אוצריא נכתבים
 * כ-`style.setProperty` על `documentElement`, וסגנון inline מנצח את הכלל
 * `:root` שב-tokens.css. משמע ההעדפה דורסת את ברירת המחדל, ו-`removeProperty`
 * מחזירה אליה. „אין העדפה” הוא היעדר ההצהרה, ולא צבע שנשמר ב-`storage` —
 * כך שינוי עתידי של ברירת המחדל מגיע גם למי שמעולם לא בחר צבע.
 *
 * ## מה זה אינו
 *
 * העדפה של התוכנה, לא תכונה של המסמך — בדיוק כמו הסרגל (host/settings.ts):
 * היא אינה נכתבת ל-DOCX, היא אינה נוסעת עם הקובץ, והיא אינה מודפסת
 * (`styles/print.css` מאפס `.editor-stack` ל-`background: none`).
 */
import { ref } from 'vue';
import { saveCanvasColor } from '../host/settings';

/** הטוקן שהבד נצבע ממנו. ההגדרה וברירת המחדל ב-styles/tokens.css. */
export const CANVAS_COLOR_VAR = '--word-canvas-bg';

/**
 * החלק של `surfaceContainerHigh` בתערובת — `color-mix(... 35%, surface)` של
 * `--word-shell-band-bg` ב-tokens.css. כתוב פעמיים, כאן בשביל הפס שבבורר ושם
 * בשביל הבד, ו-tests/unit/canvas-color.test.ts מחזיק את השניים זהים.
 */
export const BAND_MIX = 0.35;

/** ערכי ברירת המחדל של שני הטוקנים ב-tokens.css, לפני שאוצריא שולחת ערכת נושא. */
export const FALLBACK_SURFACE = '#f8f9fa';
export const FALLBACK_SURFACE_CONTAINER_HIGH = '#f3f2f1';

/** `color-mix(in srgb, high BAND_MIX, surface)` — מה שהדפדפן מצייר על הבד. */
export function bandColor(high: string, surface: string): string {
  const channels = (hex: string) => hex.slice(1).match(/../g)!.map((pair) => parseInt(pair, 16));
  const h = channels(high);
  const s = channels(surface);
  return `#${h
    .map((c, i) => Math.round(c * BAND_MIX + s[i] * (1 - BAND_MIX)).toString(16).padStart(2, '0'))
    .join('')}`;
}

/** הצבע שהבד נצבע בו בלי העדפה, לפני שאוצריא שולחת ערכת נושא. */
export const DEFAULT_CANVAS_COLOR = bandColor(FALLBACK_SURFACE_CONTAINER_HIGH, FALLBACK_SURFACE);

/** ברירת המחדל כפי שהיא נראית עכשיו, לפי ערכת הנושא — בשביל הפס בבורר. */
export const canvasDefaultColor = ref(DEFAULT_CANVAS_COLOR);

/**
 * מעדכנת את `canvasDefaultColor` אחרי החלת ערכת נושא. אותה נפילה כמו
 * ב-host/theme.ts: בלי `surfaceContainerHigh` — `surfaceContainerHighest`.
 * ערך שאינו `#rrggbb` משאיר את ברירת המחדל של tokens.css.
 */
export function setCanvasDefaultColor(colors: {
  surface?: string;
  surfaceContainerHigh?: string;
  surfaceContainerHighest?: string;
}): void {
  const surface = normalizeCanvasColor(colors.surface) ?? FALLBACK_SURFACE;
  const high = normalizeCanvasColor(colors.surfaceContainerHigh)
    ?? normalizeCanvasColor(colors.surfaceContainerHighest)
    ?? FALLBACK_SURFACE_CONTAINER_HIGH;
  canvasDefaultColor.value = bandColor(high, surface);
}

/**
 * `#rrggbb` באותיות קטנות, או `null` על כל דבר אחר.
 *
 * שני מקורות מזינים את זה, ושניהם כבר מייצרים בדיוק את הצורה הזאת: הפלטה
 * של `ColorPickerPopover` כתובה כך במקור, ו-`input[type=color]` מוגדר להחזיר
 * „simple color” — שבעה תווים, אותיות קטנות. הבדיקה כאן אינה בשבילם אלא
 * בשביל המסלול השלישי: מה שחוזר מה-`storage` של אוצריא הוא JSON שנכתב
 * בהפעלה קודמת, וגרסה קודמת, ערך שנקטע או קריאה שנכשלה מגיעים לאותה נקודה.
 * צבע פגום שמגיע עד `setProperty` פשוט אינו צובע — כלומר בד ללא רקע כלל —
 * ולכן הדחייה כאן ולא שם.
 */
export function normalizeCanvasColor(raw: unknown): string | null {
  if (typeof raw !== 'string') return null;
  const value = raw.trim().toLowerCase();
  return /^#[0-9a-f]{6}$/.test(value) ? value : null;
}

/**
 * ההעדפה כפי שהיא כרגע. `null` = אין העדפה, והבד ב-`canvasDefaultColor`.
 *
 * מודול ולא `provide`: הבד נצבע בעלייה (App.vue), והפקד שמשנה אותו יושב
 * בלשונית „תצוגה” — שהיא `v-else-if` ב-Ribbon.vue, כלומר מורכבת רק כשהיא
 * הפעילה. מצב שחי בקומפוננטה היה נולד מחדש בכל מעבר לשונית.
 */
export const canvasColor = ref<string | null>(null);

/**
 * מחילה צבע על הבד בלי לשמור אותו.
 *
 * זהו המסלול של העלייה: App.vue קורא את ההעדפה במקביל לשאר ההגדרות ומוסר
 * אותה לכאן. הפרדה מ-`setCanvasColor` כדי שהעלייה לא תכתוב חזרה ל-`storage`
 * את מה שהרגע קראה ממנו.
 */
export function applyCanvasColor(color: string | null): void {
  canvasColor.value = color;
  const root = document.documentElement;
  if (color) root.style.setProperty(CANVAS_COLOR_VAR, color);
  else root.style.removeProperty(CANVAS_COLOR_VAR);
}

/**
 * הבחירה של המשתמש: מחילה מיד ושומרת. `null` = חזרה לברירת המחדל.
 *
 * צבע שאינו עובר את `normalizeCanvasColor` נקרא כ„אין העדפה”, ולא נשמר כפי
 * שהוא: הערך היחיד שאפשר לתת למחרוזת שאינה צבע הוא היעדר צבע, ושמירה שלה
 * הייתה מחזירה את אותה שאלה בהפעלה הבאה.
 *
 * הכתיבה ל-`storage` שקטה (`tryCall` ב-host/settings.ts) ולא ממתינה לפני
 * הצביעה: הבד כבר צבוע כשהסבב מול אוצריא מתחיל, וכשל שלו מאבד את הזיכרון
 * להפעלה הבאה בלבד.
 */
export async function setCanvasColor(color: string | null): Promise<void> {
  const normalized = normalizeCanvasColor(color);
  applyCanvasColor(normalized);
  await saveCanvasColor(normalized);
}
