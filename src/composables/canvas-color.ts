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
 * ## ברירת המחדל: המשטח הנמוך ביותר של ערכת הנושא
 *
 * בלי העדפה הבד עוקב אחרי `--color-surface-container-lowest`, עם נפילה
 * ל-`--color-surface`. שכבת הכהיה נפרדת מוחלת רק במצב הזה; בחירת משתמש
 * דורסת את הטוקן inline ואינה מוחשכת.
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
 * צבע fallback לפיתוח ולפני קבלת ערכת נושא. בזמן ריצה
 * `canvasDefaultColor` מתעדכן מערך הנושא ומשמש גם את פס הצבע בבורר.
 */
export const DEFAULT_CANVAS_COLOR = '#f8f9fa';

/** ברירת המחדל הנוכחית, מסונכרנת עם ערכת הנושא להצגה בבורר. */
export const canvasDefaultColor = ref(DEFAULT_CANVAS_COLOR);

/** מעדכנת את ברירת המחדל המוצגת בבורר לאחר החלת ערכת נושא. */
export function setCanvasDefaultColor(raw: unknown): void {
  canvasDefaultColor.value = normalizeCanvasColor(raw) ?? DEFAULT_CANVAS_COLOR;
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
 * ההעדפה כפי שהיא כרגע. `null` = אין העדפה, והבד ב-`DEFAULT_CANVAS_COLOR`.
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
