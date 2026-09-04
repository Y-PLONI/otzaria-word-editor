/**
 * הלוגיקה הטהורה של אזכור „@”: מתי הקלדה היא טריגר, ומה נכתב למסמך כשההצעה
 * מתקבלת. אין כאן DOM ואין קריאות מנוע — הכול נבדק ב-vitest.
 *
 * ## למה הטריגר נקרא מהטקסט ולא מאירוע המקלדת
 *
 * `event.key === '@'` היה מפספס הדבקה, ותלוי בפריסת המקלדת. קריאת הטקסט
 * שלפני הסמן עובדת בשני המקרים, ומאותו מקור שממנו נבנית ההחלפה — כך שאין
 * דרך שהטווח שיוחלף יסטה ממה שזוהה.
 *
 * ## למה לא לפתור את ההפניה כאן
 *
 * ההתאמה עצמה היא של מנוע `find_ref` באוצריא (`library.resolveRef`): טבלת
 * ראשי-התיבות שלו מונה עשרות אלפי מונחים, והנרמול שלו מוצמד לזה שבנה את
 * אינדקס ההפניות. פורט מקומי היה סוטה ממנו בשקט.
 */

import type { ResolvedRefHit } from '../types/otzaria_plugin';

export type { ResolvedRefHit };

export interface AtTrigger {
  /** מה שהוקלד אחרי ה-„@”, בלי ה-„@” עצמו. */
  query: string;
  /** ההיסט של תו ה-„@” בתוך המחרוזת שנמסרה. */
  atIndex: number;
}

/**
 * אורך מרבי להפניה. „שולחן ערוך אורח חיים סימן תרנא” הוא 32 תווים, ולכן 48
 * מכסה בנוחות גם הפניות ארוכות. מעבר לזה המשתמש כבר ממשיך לכתוב משפט, ולא
 * ממתין להצעה.
 */
const MAX_QUERY_LENGTH = 48;

/** אותו נימוק, במונחי מילים. „תלמוד ירושלמי עירובין פרק ו הלכה ז” = 7. */
const MAX_QUERY_WORDS = 7;

/**
 * סף התחלת החיפוש. זהה לסף של `library.resolveRef` עצמו ושל מסך „איתור
 * מקורות” — תו בודד מתאים לחצי הספרייה.
 */
export const MIN_QUERY_LENGTH = 2;

/**
 * תווים שאחריהם „@” עדיין פותח אזכור. בלי הבדיקה הזו כתובת דוא״ל
 * (`dev@example.com`) הייתה נראית כטריגר בכל הקלדה.
 */
const OPENS_MENTION = /[\s(\[{"'׳״“”‘’]/u;

/**
 * אותיות השימוש (בכל״ם, ו׳ החיבור, ה׳ הידיעה, ש׳ הזיקה) נצמדות למילה שאחריהן,
 * ולכן „ראה ב@פסחים לד” היא הצורה הטבעית לכתוב את זה בעברית — ופסילתה יחד עם
 * כתובות הדוא״ל הייתה חוסמת בדיוק את מה שהפיצ׳ר נועד לו. אות בודדת כזו נחשבת
 * פותחת רק כשהיא עצמה עומדת אחרי רווח: ב-`dev@example.com` התו שלפני ה-„@”
 * אינו עברי, וב„סתם@” אין אות בודדת.
 */
const PREFIX_LETTER = /[בהוכלמש]/u;

/** האם ה-„@” שב-`atIndex` פותח אזכור, לפי מה שקדם לו. */
function opensMention(text: string, atIndex: number): boolean {
  if (atIndex === 0) return true;
  const previous = text[atIndex - 1]!;
  if (OPENS_MENTION.test(previous)) return true;
  if (!PREFIX_LETTER.test(previous)) return false;
  // אות שימוש בודדת: מה שלפניה חייב להיות רווח, אחרת זו סיומת מילה שלמה.
  return atIndex === 1 || OPENS_MENTION.test(text[atIndex - 2]!);
}

/**
 * מאתרת אזכור פתוח בטקסט שלפני הסמן.
 *
 * `beforeCaret` הוא הטקסט מתחילת חלון הקריאה ועד הסמן; ההיסט המוחזר יחסי
 * לאותה מחרוזת, והקורא מוסיף לו את בסיס החלון.
 */
export function parseAtTrigger(beforeCaret: string): AtTrigger | null {
  const atIndex = beforeCaret.lastIndexOf('@');
  if (atIndex < 0) return null;

  if (!opensMention(beforeCaret, atIndex)) return null;

  const query = beforeCaret.slice(atIndex + 1);
  if (query.length > MAX_QUERY_LENGTH) return null;
  // שבירת שורה או טאב סוגרות את האזכור — הוא אינו חוצה פסקה.
  if (/[\n\r\t]/u.test(query)) return null;
  // רווח כפול הוא הסימן שהמשתמש המשיך במשפט ולא בהפניה.
  if (query.includes('  ')) return null;
  if (query.includes('@')) return null;

  const words = query.trim().split(/\s+/u).filter(Boolean);
  if (words.length > MAX_QUERY_WORDS) return null;

  return { query, atIndex };
}

/** האם הטריגר בשל לשליחה לאוצריא. */
export function isQueryable(trigger: AtTrigger): boolean {
  return trigger.query.trim().length >= MIN_QUERY_LENGTH;
}

/**
 * בונה את היעד שאליו הקישור יצביע.
 *
 * כשיש `id` מספרי מוחזר קישור עומק ישיר, שעובד גם במסמך שנפתח מחוץ לאוצריא.
 * בהיעדרו — ספר אישי או PDF ממערכת הקבצים — מוחזר קישור „איתור מקורות” עם
 * ההפניה כלשונה: `user_books.db` מקצה מזהים באותו טווח כמו ספריית הבסיס,
 * ולכן `otzaria://open/book/<id>` אליו היה נפתר לפי מי שקדם ברשימה. העברת
 * הפתירה לאוצריא עצמה מגיעה ליעד הנכון גם שם.
 */
export function buildRefHref(hit: ResolvedRefHit, rawRef: string): string {
  if (hit.id != null && !hit.isUserBook) {
    if (hit.isPdf) {
      // ב-PDF ה-index הוא מספר עמוד 1-based, וזה גם מה שהראוטר דורש.
      const page = Math.max(1, Math.trunc(hit.index));
      return `otzaria://open/pdf/${hit.id}?index=${page}`;
    }
    const line = Math.max(0, Math.trunc(hit.index));
    return `otzaria://open/book/${hit.id}?index=${line}`;
  }
  return `otzaria://open/detection?q=${encodeURIComponent(rawRef.trim())}`;
}

/**
 * הטקסט שייכתב במסמך במקום ה-„@”.
 *
 * ההפניה שנפתרה מועדפת על מה שהוקלד: היא הצורה הקנונית של אותו מיקום, ואם
 * הקישור מצביע ל„בראשית פרק א” אין טעם שהטקסט הנראה יגיד משהו אחר.
 */
export function buildLinkText(hit: ResolvedRefHit, query: string): string {
  return hit.reference.trim() || hit.title.trim() || query.trim();
}

/** השורה המשנית בהצעה: מאיפה בספרייה ההתאמה, וכמה היא מדויקת. */
export function suggestionSubtitle(hit: ResolvedRefHit): string {
  const parts: string[] = [];
  if (hit.bookPath.trim()) parts.push(hit.bookPath.trim());
  if (hit.isUserBook) parts.push('ספר אישי');
  else if (hit.isPdf) parts.push('PDF');
  // רמת TOC בלבד — הקישור יגיע לתחילת הפרק, לא לשורה עצמה.
  if (!hit.isSourceLine && !hit.isPdf) parts.push('לרמת הפרק');
  return parts.join(' · ');
}
