/**
 * גופן שהמשתמש **בחר** ושהדפדפן אינו מצייר — ומאיפה משיגים אותו.
 *
 * ## הבעיה, כפי שנמדדה
 *
 * `fonts.listInstalled` מונה דרך GDI, כלומר משפחות של ארבעה סגנונות. וריאנט
 * משקל — `Light`, `Semilight`, `Demi` — או שמקופל לתוך משפחת הבסיס, או שמדווח
 * כמשפחה נפרדת שכרום אינו פותר: הוא מקלף אסימון סגנון מסוף השם ומחפש את
 * השאר, וזה עובד ל-`Calibri Light` ונכשל ל-`Segoe UI Semilight`.
 *
 * נמדד ב-Windows 10, מנייה אמיתית של GDI מול מדידת רוחב בכרום: מ-287 משפחות
 * שהמארח מדווח, **43 אינן נפתרות לפי השם**, וחמש מהן עבריות —
 * `Guttman Kav-Light`, `Guttman Yad-Light`, `Guttman Haim-Condensed`,
 * `Segoe UI Semilight`, `Malgun Gothic Semilight`. כולן מותקנות במכונה, כולן
 * מופיעות ברשימת הגופנים של אוצריא עצמה, ואף אחת מהן לא הייתה נגישה בבורר.
 *
 * ## מה שהקובץ הזה עושה, ומה שהוא בכוונה אינו עושה
 *
 * `fonts.resolveFamilies` מחזירה `@font-face` עם הבייטים של גופן מותקן, בשם
 * שביקשנו. זה בדיוק החסר: השם הופך לשם שהדפדפן פותר, הדגימה בבורר מציגה את
 * הגופן עצמו, והטקסט במסמך מצויר בו ולא ב-fallback.
 *
 * **לפי דרישה, ולא בעלייה.** גופן ב-base64 שוקל מאות קילובייטים, ו-43 מהם הם
 * מגה-בייטים לתוך ה-WebView בשביל רשימה שהמשתמש כנראה לא יגלול עד סופה. לכן
 * הטריגר הוא הבחירה: הגופן שהמשתמש באמת רצה, ורק הוא.
 *
 * **בלי תחליפים.** `substitutes` נשלח עם השם עצמו בלבד. המסלול של המסמך
 * (docx-fonts.ts) כן שולח שרשרת תחליפים, ושם זה נכון — גופן שהמסמך מבקש ואינו
 * קיים חייב להיראות כמו משהו. כאן ההפך: המשתמש בחר שם מפורש, והגשת גופן אחר
 * תחתיו הייתה מציירת `Segoe UI` ומכריזה `Segoe UI Semilight`. אם אין בייטים —
 * השורה נשארת מסומנת, וזו האמת.
 *
 * ## הסגנון נפרד, וזה לא פרט
 *
 * `installDocumentFontAliases` **דורס** את `FONT_ALIAS_STYLE_ID` בכל פתיחת
 * מסמך. הזרקה לתוכו הייתה נעלמת ברגע שהמשתמש פותח מסמך — כלומר גופן שעבד
 * מפסיק לעבוד, בלי שדבר בבורר משתנה. לכן יש כאן אלמנט משלנו, והוא **נצבר**
 * ולא נדרס: כל בחירה מוסיפה לו כלל.
 *
 * מטמון הזמינות נשכח רק למשפחה שהכללים שלה השתנו. הכללים שלנו נשארים
 * ב-DOM גם בפתיחת מסמך, ולכן אין צורך לסרוק שוב את כל גופני הבורר.
 */
import { tryCall } from '../host/otzaria-client';
import type { ResolveFontFamiliesResult } from '../types/otzaria_plugin';
import { forgetFamilyAvailability, isFamilyAvailable } from './docx-fonts';

/** ה-id של הסגנון שמחזיק את מה שנשלף לבורר. נפרד מזה של המסמך — ראו למעלה. */
export const PICKER_FONT_STYLE_ID = 'otzaria-picker-font-faces';

/** מחרוזת המדידה לטעינה. אותה מחרוזת שהמסלול של המסמך טוען איתה. */
const LOAD_PROBE_TEXT = 'אבגדהוזחטי ABCDEFGHIJ';

/** מה שאפשר להחליף בבדיקה. ברירות המחדל הן המציאות. */
export interface PickerFontDeps {
  call?: typeof tryCall;
  available?: (name: string) => boolean;
  forget?: (name: string) => void;
  /** ההמתנה לטעינת ה-`@font-face` שהוזרק. ראו `loadInjectedFace`. */
  load?: (name: string) => Promise<void>;
  style?: () => HTMLStyleElement | null;
}

/**
 * שמות שהמארח אינו יודע להגיש, כדי לא לשאול אותו שוב על כל בחירה.
 *
 * התשובה הזאת אינה משתנה תוך ההפעלה, וזה מה שמצדיק לזכור אותה: היא נגזרת ממה
 * שמותקן בדיסק ומהמפה שאוצריא בנתה ממנו בעלייה — לא מ-`@font-face` שהוזרק,
 * ולכן פתיחת מסמך אינה הופכת אותה.
 *
 * שם שהמארח **כן** הגיש אינו נשמר כאן: `isFamilyAvailable` הוא כבר התשובה
 * לשאלה הזאת, והוא ממותת ממילא. שתי רשימות לאותו דבר היו נפרדות ביום שאחד
 * מהמטמונים מתנקה.
 */
const refused = new Set<string>();

/** בקשות שבאוויר, לפי מפתח — שתי בחירות רצופות אינן שתי בקשות. */
const inFlight = new Map<string, Promise<boolean>>();

/** מי שרוצה לדעת שמשפחה חדשה הפכה ניתנת לציור. ראו `onPickerFontsChanged`. */
const listeners = new Set<(family: string) => void>();

/**
 * הודעה שמשפחה נוספת הפכה ניתנת לציור — ולכן יש מה להרכיב מחדש בבורר.
 *
 * מנוי ולא `ref` של Vue: הקבצים ב-`engine/` אינם מכירים את המסגרת, וזה הגבול
 * היחיד שהופך אותם לניתנים למדידה בלי DOM ובלי הרכבה. מי שמחזיק את המצב
 * (App.vue) הוא זה שממיר את זה לתגובתיות.
 *
 * מחזירה מבטל מנוי, כדי שאתר הקריאה לא יצטרך לזכור את הפונקציה שהעביר.
 */
export function onPickerFontsChanged(listener: (family: string) => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

function notifyChanged(family: string): void {
  for (const listener of [...listeners]) {
    try {
      listener(family);
    } catch (error) {
      // מנוי שנכשל אינו סיבה לבטל את שאר המנויים, ובוודאי לא את ההזרקה
      // שכבר הצליחה.
      console.warn('[otzaria-word] מנוי גופני הבורר נכשל', error);
    }
  }
}

/**
 * דואגת שהדפדפן יצייר את `name` — ומחזירה אם הוא מצייר אותו בסוף.
 *
 * אינה זורקת לעולם: כל כשל בדרך מחזיר `false`, כלומר „נשארים עם מה שיש”.
 * בחירת גופן אינה אמורה להיכשל בגלל שהמארח לא ענה.
 *
 * הסדר חשוב, וכל צעד בו חוסך את הבא: מדידה (ממותתת) לפני מטמון הסירובים, והוא
 * לפני בקשה שבאוויר, והיא לפני בקשה חדשה.
 */
export async function ensureFamilyDrawable(
  name: string,
  deps: PickerFontDeps = {},
): Promise<boolean> {
  const family = typeof name === 'string' ? name.trim() : '';
  if (family === '') return false;

  const { available = isFamilyAvailable } = deps;
  // הדפדפן כבר פותר את השם — אין מה לעשות, וזה המסלול של כמעט כל בחירה.
  if (available(family)) return true;

  const key = family.toLowerCase();
  if (refused.has(key)) return false;

  const pending = inFlight.get(key);
  if (pending) return pending;

  const request = fetchAndInject(family, key, deps).finally(() => {
    inFlight.delete(key);
  });
  inFlight.set(key, request);
  return request;
}

async function fetchAndInject(
  family: string,
  key: string,
  deps: PickerFontDeps,
): Promise<boolean> {
  const {
    call = tryCall,
    available = isFamilyAvailable,
    forget = forgetFamilyAvailability,
    load = loadInjectedFace,
    style = pickerStyleElement,
  } = deps;

  try {
    const res = await call<ResolveFontFamiliesResult>('fonts.resolveFamilies', {
      // השם כתחליף של עצמו — ראו „בלי תחליפים” בהערת הפתיחה.
      families: [{ name: family, substitutes: [family] }],
    });
    const css = typeof res?.css === 'string' ? res.css : '';
    if (css === '') {
      refused.add(key);
      return false;
    }

    const element = style();
    if (!element) return false;
    // נצבר ולא נדרס: בחירה שנייה אינה אמורה למחוק את הגופן של הראשונה.
    element.textContent = `${element.textContent ?? ''}\n${css}`;

    await load(family);
    // אחרי הטעינה ולא לפניה: `url(data:…)` אינו מוכן ברגע שה-CSS נכתב, ומדידה
    // באותו חלון הייתה נכנסת למטמון כ-`false` ונשארת שם.
    forget(family);

    if (!available(family)) {
      // הבייטים הגיעו והדפדפן לא קיבל אותם — פורמט שאינו נתמך, למשל. אין טעם
      // לבקש שוב, וגם אין מה להודיע: שום שורה בבורר לא השתנתה.
      refused.add(key);
      return false;
    }

    notifyChanged(family);
    return true;
  } catch (error) {
    console.warn(`[otzaria-word] שליפת הגופן ${family} מאוצריא נכשלה`, error);
    return false;
  }
}

/**
 * ממתינה לטעינת ה-`@font-face` שהוזרק, ובולעת כשל.
 *
 * `FontFaceSet.load` ולא `ready`: הוא מבקש את המשפחה במפורש, ולכן גם כלל שטרם
 * שימש בטקסט מתחיל להיטען. `ready` יכול להיפתר לפני שמשפחה כזאת בכלל נכנסה
 * ל-`loading` — אותה מכניקה בדיוק שבמסלול של המסמך.
 */
async function loadInjectedFace(family: string): Promise<void> {
  const fontSet = typeof document === 'undefined' ? undefined : document.fonts;
  if (!fontSet || typeof fontSet.load !== 'function') return;
  try {
    await fontSet.load(`72px ${JSON.stringify(family)}`, LOAD_PROBE_TEXT);
  } catch {
    // הכלל כבר ב-DOM; אם הטעינה נדחתה, המדידה שאחריה תגיד את זה בעצמה.
  }
}

/** הסגנון שמחזיק את הכללים, נוצר בפעם הראשונה. `null` בלי DOM. */
function pickerStyleElement(): HTMLStyleElement | null {
  if (typeof document === 'undefined') return null;

  const existing = document.getElementById(PICKER_FONT_STYLE_ID);
  if (existing instanceof HTMLStyleElement) return existing;

  const style = document.createElement('style');
  style.id = PICKER_FONT_STYLE_ID;
  document.head.appendChild(style);
  return style;
}

/**
 * מאפסת את הזיכרון שבין הבחירות. **לבדיקות בלבד** — מטמון הסירובים והמנויים
 * הם מצב מודול, ובדיקה שנייה הייתה יורשת את התשובות של הראשונה.
 */
export function resetPickerFonts(): void {
  refused.clear();
  inFlight.clear();
  listeners.clear();
}
