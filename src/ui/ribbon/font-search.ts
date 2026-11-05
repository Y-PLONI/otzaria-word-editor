/**
 * החיפוש בבורר הגופן, במנותק מהקומפוננטה.
 *
 * למה מודול נפרד, ומאותו טעם כמו `aria.ts`: אלה ההכרעות שנוטים לשבור בלי
 * לשים לב — מה נחשב התאמה, מה מדורג לפני מה, ולאן חץ מתקדם — וכפונקציות
 * טהורות אפשר למדוד אותן ישירות. הקומפוננטה נשארת חיווט.
 *
 * ## ההכרעה המרכזית: חיפוש מבטל את הקיבוץ
 *
 * בלי שאילתה הרשימה מקובצת („עברית”, „כל הגופנים”) — הקיבוץ הוא מה שהופך
 * מאות שמות לרשימה שאפשר לגלול בה. **עם** שאילתה הקיבוץ הופך למכשול: מי
 * שהקליד „ari” יודע מה הוא מחפש, והתאמה מדויקת שיושבת בקבוצה השלישית מתחת
 * לכותרת היא בדיוק מה שמסתיר אותה. לכן חיפוש מחזיר רשימה שטוחה **מדורגת**,
 * והדירוג הוא מה שמציב את ההתאמה הנכונה ראשונה.
 *
 * ## ולמה הדירוג ולא סדר הרשימה
 *
 * „ari” מופיע ב-`Arial`, ב-`Arial Black` וב-`Bahnschrift Ari…`. סדר הרשימה
 * המקורי היה מציב את מה שבמקרה הגיע ראשון; הדירוג מציב את מי שהשם שלו
 * **מתחיל** בשאילתה. זה ההבדל בין להקליד שלוש אותיות ולחצוץ Enter לבין
 * להקליד שלוש אותיות ולחפש בעיניים.
 */

/** אפשרות אחת בבורר, בצורה שהקומפוננטה מציגה. */
export interface ComboOption {
  value: string;
  label: string;
  /** `font-family` של CSS — כך כל שם מוצג בגופן שלו. */
  preview?: string;
  /** כותרת הקבוצה. ריק או חסר = בראש, בלי כותרת. */
  group?: string;
  /**
   * הגופן מכסה עברית — ואז נוספת לשורה דגימה של אותיות עבריות בגופן עצמו.
   * הדירוג והניווט אינם מבחינים בו: הוא מראה בלבד. ראו `RibbonCombo`.
   */
  hebrew?: boolean;
  /**
   * הדפדפן אינו פותר את שם הגופן, ולכן אין עליו דגימה אמיתית להציג.
   *
   * הדירוג והניווט אינם מבחינים בו, ובכוונה: השורה נשארת בחירה חוקית — מסמך
   * נודד, והגופן עשוי להיות מותקן במכונה שבה המסמך ייפתח. מה שנופל הוא הציור
   * בלבד. נקבע ב-engine/font-options.ts (`available`).
   */
  unavailable?: boolean;
  /** `unavailable` נמדד ולא הוכרז. ראו `PickerOption` ב-composables/picker-value.ts. */
  measured?: boolean;
  /** הזמינות טרם נבדקה; אין דגימה עד שתגיע תשובה. */
  pending?: boolean;
  /**
   * הגופן מותקן במכונה והדפדפן בכל זאת אינו מצייר אותו — ואז ההסבר על השורה
   * הוא אחר. ראו `PickerOption` ב-composables/picker-value.ts.
   */
  installedNotDrawable?: boolean;
}

/** שורה ברשימה הנפתחת: כותרת קבוצה, או אפשרות עם המספר שלה לניווט מקלדת. */
export type ComboRow =
  | { type: 'group'; label: string }
  | { type: 'option'; option: ComboOption; index: number };

/** מה שהרשימה הנפתחת מציגה, ובכמה אפשרויות אפשר לנווט. */
export interface ComboRows {
  rows: readonly ComboRow[];
  count: number;
}

/** השוואה חסרת רגישות לאותיות ולרווחים בקצוות. */
function fold(text: string): string {
  return text.trim().toLowerCase();
}

/**
 * דירוג ההתאמה: קטן יותר = מוקדם יותר. `null` = אינו מתאים כלל.
 *
 * הערך נבדק ולא רק התווית מפני שהשניים נבדלים: „David” היא התווית של
 * `TaameyDavidCLM`, ומי שמקליד „taamey” מתכוון בדיוק לגופן הזה.
 */
export function matchRank(option: ComboOption, query: string): number | null {
  const q = fold(query);
  if (q === '') return 0;

  const label = fold(option.label ?? '');
  const value = fold(option.value ?? '');

  if (label.startsWith(q)) return 0;
  if (value.startsWith(q)) return 1;
  if (label.includes(q)) return 2;
  if (value.includes(q)) return 3;
  return null;
}

/**
 * הרשימה שתוצג: מקובצת כשאין שאילתה, שטוחה ומדורגת כשיש.
 *
 * המיון יציב (`index` כמפתח שני) — שני שמות באותו דירוג שומרים על סדר הרשימה
 * המקורי, שכבר הכריע מה חשוב יותר.
 */
export function buildComboRows(
  options: readonly ComboOption[],
  query: string,
): ComboRows {
  const q = fold(query);

  if (q === '') {
    const rows: ComboRow[] = [];
    let index = 0;
    let openGroup: string | null = null;

    for (const option of options) {
      const group = option.group ?? '';
      if (group !== '' && group !== openGroup) {
        rows.push({ type: 'group', label: group });
      }
      // גם קבוצה ריקה נרשמת: אפשרות חשופה שבאה **אחרי** קבוצה עם כותרת אינה
      // אמורה להיבלע בה, וכותרת חוזרת עדיפה על שיוך שקרי.
      openGroup = group === '' ? null : group;
      rows.push({ type: 'option', option, index: index++ });
    }
    return { rows, count: index };
  }

  const ranked = options
    .map((option, order) => ({ option, order, rank: matchRank(option, q) }))
    .filter((entry): entry is { option: ComboOption; order: number; rank: number } => entry.rank !== null)
    .sort((a, b) => a.rank - b.rank || a.order - b.order);

  return {
    rows: ranked.map(({ option }, index) => ({ type: 'option', option, index })),
    count: ranked.length,
  };
}

/**
 * לאן מקש ניווט מזיז את הסימון. `null` כשהמקש אינו מקש ניווט.
 *
 * אין כאן היפוך RTL, בניגוד ל-`nextTabIndex`: הרשימה אנכית, ו-WAI-ARIA קובע
 * שהחצים נעים לפי הכיוון **החזותי** — למטה הוא למטה בשתי השפות.
 *
 * עטיפה מסוף לתחילה, ו-`current` שלילי (אין סימון עדיין) נכנס לראש הרשימה
 * ב-ArrowDown ולסופה ב-ArrowUp — בדיוק כמו בורר נייטיב.
 */
export function nextOptionIndex(key: string, current: number, count: number): number | null {
  if (count <= 0) return null;

  switch (key) {
    case 'ArrowDown':
      return current < 0 ? 0 : (current + 1) % count;
    case 'ArrowUp':
      return current < 0 ? count - 1 : (current - 1 + count) % count;
    case 'Home':
      return 0;
    case 'End':
      return count - 1;
    default:
      return null;
  }
}

/**
 * מה ש-Enter מחיל.
 *
 * שלושה מצבים, ורק השלישי אינו מובן מאליו:
 * 1. יש סימון — הוא נבחר.
 * 2. אין סימון ואין שאילתה — לא קורה דבר.
 * 3. אין אף התאמה, אבל הוקלד טקסט — הטקסט **עצמו** מוחל.
 *
 * השלישי הוא מה ש-Word עושה, וזה לא קפריזה: גופן שמותקן ואינו ברשימת
 * המועמדים (ובלי מארח שמונה — זה קורה) לא היה נגיש בשום דרך אחרת. המשתמש
 * יודע מה שמו, והבורר אינו אמור להתווכח איתו.
 *
 * ## `normalize` — כשהרשימה היא הצעה ולא מלאי
 *
 * בבורר הגודל הטקסט שהוקלד הוא **תמיד** ערך לגיטימי: 13pt אינו „גודל שאינו
 * קיים”, הוא פשוט אינו בסולם שהרשימה מציעה. פקד שמקבל `normalize` מוסר את
 * הכרעת המצב השלישי לקורא — שם יושבים גם ההידוק לטווח והעיגול לחצי נקודה
 * (`parseFontSizeInput`) — ומחיל את מה שהוקלד גם כשיש התאמות ברשימה.
 *
 * בלי זה „13” היה מוחל כ„10”: החיפוש מדרג כל אפשרות ש**מכילה** את הספרה,
 * ו-Enter היה בוחר את ראש הדירוג במקום את המספר שהוקלד.
 */
export function commitValue(
  rows: ComboRows,
  activeIndex: number,
  query: string,
  normalize?: (typed: string) => string | null,
): string | null {
  for (const row of rows.rows) {
    if (row.type === 'option' && row.index === activeIndex) return row.option.value;
  }
  const typed = query.trim();
  if (typed === '') return null;
  if (normalize) return normalize(typed);
  return rows.count === 0 ? typed : null;
}
