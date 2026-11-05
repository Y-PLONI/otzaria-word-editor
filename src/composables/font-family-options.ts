/**
 * רשימת הגופנים כפי ש**בורר** רואה אותה — פעם אחת, ולכל מי שמציג בורר גופן.
 *
 * ## למה זה יצא מ-`use-font-controls.ts`
 *
 * המיפוי מ-`FontFamilyChoice` (מה שהמיזוג ב-engine/font-options.ts מייצר) אל
 * `PickerOption` (מה שהפקד מצייר) אינו העתקת שדות: הוא נושא שלוש הכרעות שכל
 * אחת מהן נמדדה ותוקנה אחרי באג — היפוך הדגל ל-`unavailable`, הוספת הגופן
 * שבמסמך כשאינו ברשימה (`withCurrent`), והשאלה מי בכלל רשאי לשאול את הדפדפן
 * אם השם נפתר. כל עוד המיפוי חי בתוך בורר הרצועה, בורר **שני** היה מקבל או
 * העתק שיתיישן, או — וזה מה שהיה — תיבת טקסט חופשי במקום בורר.
 *
 * ושתי תיבות כאלה היו: „גופן מורכב” ב„גופן מתקדם”, ו„גופן ברירת מחדל”
 * ב„ברירות מחדל למסמך”. שתיהן ביקשו מהמשתמש להקליד שם גופן מהזיכרון — בעורך
 * שהרשימה שלו יודעת בדיוק מה מותקן במכונה, מה יש במסמך, ומה מכסה עברית.
 *
 * ## „ללא שינוי” כשורה ברשימה, ולא כתיבה ריקה
 *
 * שני הדיאלוגים הם patch לפי מפתח: שדה שלא נגעו בו אינו נשלח. בתיבת טקסט
 * המצב הזה הוא „ריק”, וריק הוא גם מה שנראה כשמוחקים — כלומר אותה תצוגה לשני
 * דברים. בבורר אין „ריק” שאפשר להקליד: `commitValue` אינו מחיל מחרוזת ריקה,
 * ולכן בלי שורה מפורשת לא הייתה שום דרך **לחזור** מבחירת גופן ל„אל תיגע”.
 * השורה הזאת היא הדרך, והיא גם מה שאומר את המצב במילים במקום בהיעדר טקסט.
 */
import { computed, type ComputedRef } from 'vue';
import { cachedFamilyAvailability } from '../engine/docx-fonts';
import type { FontFamilyChoice } from '../engine/font-options';
import { withCurrent, type PickerOption } from './picker-value';
import { useFontOptions } from './useFontOptions';

/** הערך של „אל תיגע בגופן”. ריק, מפני שזה גם מה שלא נשלח ב-patch. */
export const FONT_FAMILY_UNCHANGED = '';

/**
 * המיפוי עצמו — טהור, ולכן נמדד בלי Vue ובלי DOM.
 *
 * `current` הוא הגופן שהבורר עומד עליו: הוא נוסף לרשימה כשאינו בה, וזה מה
 * שמונע בורר שמציג את השורה הראשונה במקום את מה שבאמת מוחל.
 */
export function familyPickerOptions(
  families: readonly FontFamilyChoice[],
  current: string,
): readonly PickerOption[] {
  return withCurrent(
    families.map((option) => ({
      value: option.value,
      label: option.label,
      preview: option.previewFamily,
      // הקיבוץ וכיסוי העברית נקבעים במיזוג ולא כאן — engine/font-options.ts.
      group: option.group,
      hebrew: option.hebrew,
      /*
       * הדגל הפוך לזה שבמיזוג (`available`) בכוונה: הפקד מסמן **חריגה**, ו-
       * `unavailable` דלוק רק בשורה שיש עליה מה לומר. `available` היה מחייב
       * את הפקד לצייר סימון על היעדר דגל — כלומר גם על כל בורר שאינו גופנים,
       * שאין לו את השדה בכלל.
       */
      unavailable: option.available === false,
      measured: option.measured,
      pending: option.available === undefined,
      // מותקן שאינו מצוייר — הפקד אומר עליו משהו אחר מ„אינו מותקן”. ראו
      // `installedNotDrawable` ב-engine/font-options.ts.
      installedNotDrawable: option.installedNotDrawable,
    })),
    current,
    /*
     * גופן שאינו ברשימה מוצג בגופן עצמו, כמו כל שאר השורות — **אם** הדפדפן
     * פותר אותו. `available` הוא מה שמונע מהמסלול הזה להיות היצרן השני של
     * השורה המשקרת: הוא נגיש בהקלדת שם חופשי, ואז השורה והתיבה הסגורה היו
     * מציירות שם שאינו קיים ב-fallback.
     */
    { preview: true, available: cachedFamilyAvailability },
  );
}

/**
 * אותה רשימה, כ-`computed`, למי שמציג בורר גופן בקומפוננטה.
 *
 * `unchangedLabel` הוא מה שמפריד בין שני הקוראים: הרצועה מחילה בכל בחירה ואין
 * לה מצב „לא נגעתי” בכלל, ודיאלוג שמרכיב patch חייב אותו. מי שאינו מוסר אותו
 * מקבל בדיוק את הרשימה שהייתה.
 */
export function useFamilyPicker(
  current: () => string,
  unchangedLabel?: string,
): ComputedRef<readonly PickerOption[]> {
  const { families } = useFontOptions();
  return computed(() => {
    const options = familyPickerOptions(families.value, current());
    if (unchangedLabel === undefined) return options;
    return [{ value: FONT_FAMILY_UNCHANGED, label: unchangedLabel }, ...options];
  });
}
