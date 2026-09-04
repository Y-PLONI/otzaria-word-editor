/**
 * ניתוב לחיצה על קישור `otzaria://` שבתוך המסמך אל הקורא של אוצריא.
 *
 * בלי זה קישור עומק שנוצר במסמך אינו עושה דבר: התנהגות ברירת המחדל של המנוע
 * היא לפתוח את ה-href בדפדפן, וה-WebView חוסם `window.open`.
 *
 * `onActivate` חייב להחזיר סינכרונית (Promises אינם נתמכים), ולכן הניווט
 * נשלח ואינו מומתן, והתשובה היא `suppress` מיד. קישור שאינו `otzaria://`
 * מוחזר ל-`default` ואינו נוגע בנו.
 */
import { parseOtzariaLink, type OtzariaLinkTarget } from './otzaria-link';

/** צורת ההקשר שמעניינת אותנו מתוך `HyperlinkActivationContext` של SuperDoc. */
export interface LinkActivationContextLike {
  href?: string | null;
}

export type LinkActivationResult = { type: 'default' } | { type: 'suppress' };

export interface LinkActivationOptions {
  /** מבצע את הניווט בפועל. מוזרק כדי שהבדיקות לא ידרשו גשר. */
  navigate: (target: OtzariaLinkTarget) => void;
  /** דיווח כשל למשתמש. */
  onStatus?: (message: string, isError: boolean) => void;
}

/** בונה את ה-handler שנמסר ל-`new SuperDoc({ hyperlinks: { onActivate } })`. */
export function createOtzariaLinkActivation(
  options: LinkActivationOptions,
): (context: LinkActivationContextLike) => LinkActivationResult {
  return (context) => {
    const target = parseOtzariaLink(context?.href ?? '');
    if (!target) return { type: 'default' };
    try {
      options.navigate(target);
    } catch (error) {
      console.warn('[otzaria-word] פתיחת הקישור נכשלה', error);
      options.onStatus?.('פתיחת הקישור נכשלה', true);
    }
    return { type: 'suppress' };
  };
}
