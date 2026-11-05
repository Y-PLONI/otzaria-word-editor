/**
 * שני הכללים שכל בורר ברצועה חי לפיהם — במקום אחד, ולא בכל לשונית מחדש.
 *
 * שניהם נכתבו במקור בתוך `HomeTab.vue`, ושניהם תוקנו שם אחרי באג אמיתי. הם
 * יצאו לכאן ברגע שאותם בוררים הופיעו במקום שני (תפריט הלחצן הימני): כלל
 * שמועתק הוא כלל שיתוקן פעם אחת בלבד.
 */
import type { Ref } from 'vue';
import type { CommandOutcome } from '../engine/command-adapter';

/**
 * אפשרות אחת בבורר, בצורה שמשותפת ל-`RibbonSelect` ול-`RibbonCombo`.
 *
 * `preview` ו-`group` אופציונליים מפני שהבורר הנייטיב אינו מכיר קיבוץ, ובורר
 * הגודל אינו מציג תצוגה מקדימה — אבל שני הפקדים מקבלים את אותה צורה, ולכן
 * `withCurrent` אחד מספיק לשלושת הבוררים.
 */
export interface PickerOption {
  value: string;
  label: string;
  /** `font-family` של CSS — כך כל שם מוצג בגופן עצמו. */
  preview?: string;
  /** כותרת הקבוצה בבורר החיפוש. חסר = בראש, בלי כותרת. */
  group?: string;
  /**
   * הגופן מכסה עברית — ואז בבורר החיפוש מופיעה לפני שמו דגימה של אותיות
   * עבריות. נקבע ב-engine/font-options.ts; ראו שם למה זו שאלה של כיסוי ולא של
   * ייעוד. חסר = בלי דגימה, וזו גם ברירת המחדל של בורר הגודל ושל האפשרות
   * שנוספת ב-`withCurrent`.
   */
  hebrew?: boolean;
  /**
   * הדפדפן אינו פותר את שם הגופן, ולכן אין דגימה אמיתית להציג עליו.
   *
   * השורה נשארת בחירה חוקית — מסמך נודד, והגופן עשוי להיות מותקן במכונה שבה
   * המסמך ייפתח. מה שנופל הוא **הציור**: `preview` חסר, והשורה מסומנת כך
   * שהמשתמש לא יחשוב שמה שהוא רואה הוא הגופן. נקבע ב-engine/font-options.ts
   * (`available`), ורק למקורות שאין להם מזריק — ראו שם.
   */
  unavailable?: boolean;
  /**
   * `unavailable` נגזר ממדידה ולא מהכרזה.
   *
   * מה שזה קונה: שער QA יכול לבדוק רק את השורות שיש עליהן מה לבדוק. גופן
   * שהוכרז זמין בלי מדידה (ראו `available` ב-engine/font-options.ts) אינו
   * נפתר בכרום נקי, ושער שדורש הסכמה על כל שורה היה נכשל עליו תמיד.
   */
  measured?: boolean;
  /** הזמינות טרם נבדקה — אין דגימה ואין טענה שהגופן חסר. */
  pending?: boolean;
  /**
   * הגופן **מותקן** במכונה, ובכל זאת הדפדפן אינו מצייר אותו.
   *
   * מפריד בין שני מצבים שנראים זהים בשורה ואינם: `Aptos` שאינו במכונה, לעומת
   * `Guttman Kav-Light` שמותקן ושהמנייה מדווחת בשם GDI שהדפדפן אינו פותר.
   * הפקד בוחר לפיו את ההסבר, כי „הגופן אינו מותקן במכונה” הוא שקר על השני —
   * וזו בדיוק התלונה שהגיעה: המשתמש התקין את הגופנים שנעדרו.
   *
   * מקורו `installedNotDrawable` ב-engine/font-options.ts. חסר בכל בורר שאינו
   * גופנים, ובכל שורה שנפתרת.
   */
  installedNotDrawable?: boolean;
}

/**
 * צורת הרשימה שהערך נוסף אליה.
 *
 * פרמטר מפורש ולא ניחוש טיפוס מהמחרוזת: „20.5” הוא גודל, „Arial” הוא שם, ו-
 * `Number(current)` היה מכריע את זה נכון בדיוק עד לגופן שהשם שלו מתחיל בספרה.
 * הקורא יודע איזה בורר הוא מרכיב, וזה סוף הדיון.
 */
export interface CurrentShape {
  /**
   * האפשרות שנוספת נושאת `preview`, כלומר הערך הוא שם גופן.
   *
   * ברירת המחדל `false` היא מה שמונע `style="font-family: 13"` בבורר הגודל:
   * `preview` הוא `font-family` של CSS, ובורר מספרי שמכריז עליו מבקש
   * מהדפדפן גופן בשם „13”. ברצועה זה לא נראה כל עוד הגודל היה `<select>`
   * (`RibbonSelect` אינו מצייר `preview` בכלל), ונראה מיד כשהוא הפך
   * ל-`RibbonCombo`.
   */
  readonly preview?: boolean;
  /**
   * הרשימה מספרית, ולכן הערך נכנס **במקומו לפי הסדר** ולא בראשה.
   *
   * גם זה נראה רק מרגע שבורר הגודל הפך לרשימה נפתחת: 13pt בראש הרשימה נתן
   * „13, 8, 9, 10, 11, 12…”, כלומר סולם מספרים שאינו מסודר — ובורר שאי אפשר
   * לאמוד בו מרחק הוא בורר שצריך לקרוא כל שורה בו.
   */
  readonly numeric?: boolean;
  /**
   * האם הדפדפן פותר את שם הגופן — ומה שקורה כשלא.
   *
   * בלי זה האפשרות שנוספת כאן הייתה **המסלול השני** שמייצר בדיוק את השורה
   * שהתיקון ב-engine/font-options.ts בא לבטל: `preview: current` הוא השם
   * החשוף, ולכן שם שאינו נפתר צויר ב-fallback תחת השם שלו. והוא נגיש:
   * `commitValue` מחיל טקסט חופשי כמו שהוא כשאין התאמה, כלומר די להקליד שם
   * גופן שאינו מותקן וללחוץ Enter.
   *
   * הסתם-פונקציה ולא ייבוא של `isFamilyAvailable`: הקובץ הזה משרת גם את בורר
   * הגודל, ומדידת גופנים אינה עניינו. הקורא שיודע שהוא מרכיב בורר גופן הוא
   * זה שמעביר — כמו `preview` ו-`numeric` שמעליו.
   *
   * חסר = לא נשאל, והאפשרות נחשבת זמינה. זו ההתנהגות שהייתה, ולכן בורר
   * הגודל אינו משתנה.
   */
  readonly available?: (value: string) => boolean | undefined;
}

/**
 * הערך הנוכחי חייב להיות אחת האפשרויות, אחרת `<select>` מציג את הראשונה
 * ומשקר. גופן או גודל שאינם ברשימה (מסמך שנכתב בגופן שהמנוע לא הציע, טקסט
 * ב-20.5pt) מתווספים אליה — בדיוק מה ש-Word עושה.
 */
export function withCurrent<T extends PickerOption>(
  options: readonly T[],
  current: string,
  shape: CurrentShape = {},
): readonly (T | PickerOption)[] {
  /*
   * השוואה **מדויקת**, ובכוונה — ניסיון לקפל רישיות כאן נמדד כגרוע יותר מהפגם
   * שהוא פתר. `RibbonCombo` משווה מדויק בארבעה מקומות (`chosen`,
   * `aria-selected`, `previewStyle`, `indexOfCurrent`), ולכן ערך שנבלע כאן
   * בשל קיפול משאיר את הרשימה **בלי שום שורה מסומנת**: לא רואים איזה גופן
   * מוחל, והרשימה נפתחת מהראש במקום עליו.
   *
   * ההתאמה בין „APTOS” שבמסמך לבין „Aptos” שברשימה נעשית לכן במקור — הקורא
   * מנרמל את הערך לאיות הקנוני לפני שהוא מגיע לכאן (ראו `use-font-controls`).
   */
  if (current === '' || options.some((option) => option.value === current)) return options;

  const available = shape.available ? shape.available(current) : true;
  const measured = shape.available !== undefined && available !== undefined;
  const added: PickerOption = shape.preview
    ? available === undefined
      ? { value: current, label: current, pending: true, measured: false }
      : available
      ? { value: current, label: current, preview: current, measured }
      : { value: current, label: current, unavailable: true, measured }
    : { value: current, label: current };

  const value = Number(current);
  // ערך שאינו מספר בבורר מספרי אינו „שגיאה” כאן — פשוט אין לו מקום בסולם,
  // ולכן הוא חוזר להתנהגות הרגילה: בראש, שם רואים אותו.
  if (!shape.numeric || !Number.isFinite(value)) return [added, ...options];

  const at = options.findIndex((option) => {
    const other = Number(option.value);
    return Number.isFinite(other) && other > value;
  });
  return at === -1
    ? [...options, added]
    : [...options.slice(0, at), added, ...options.slice(at)];
}

/**
 * שולחת את הבחירה ומחזיקה אותה על המסך עד לתשובה: בהצלחה היא נשמרת כ„אחרון
 * שידענו”, ובכשל היא נעלמת — כלומר מה שלא קרה במסמך אינו מוצג.
 *
 * למה אופטימי ולא „להמתין לתשובה”: הבורר חייב להגיב מיד, והמנוע אינו מדווח
 * ערך בכלל על בחירה מעורבת — תיבה שממתינה לו הייתה נראית קפואה גם בהצלחה
 * מלאה.
 *
 * הבדיקה `pending.value !== next` לפני העדכון: אם המשתמש בחר שוב בזמן
 * ההמתנה, הבקשה שבאוויר אינה שלנו יותר, ותשובה מאוחרת אינה אמורה למחוק בחירה
 * טרייה.
 */
export async function applyOptimistically<T>(
  pending: Ref<T | null>,
  memo: Ref<T>,
  next: T,
  run: () => Promise<CommandOutcome>,
): Promise<void> {
  pending.value = next;
  const outcome = await run();
  if (pending.value !== next) return;
  if (outcome.ok) memo.value = next;
  pending.value = null;
}
