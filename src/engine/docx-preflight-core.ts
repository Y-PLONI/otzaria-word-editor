/**
 * תיקון מה שהמנוע אינו קורא נכון, לפני שהוא רואה אותו.
 *
 * שני תיקונים יושבים כאן, ומשני טעמים שונים: הראשון מפני שאין שום מקום אחר
 * שיכול לתפוס אותו (קיפאון על החוט הראשי), והשני מפני שאין שום מקום אחר שבו
 * הוא נכון (הדגשה שצריכה להיכנס לחישוב הפריסה). מה שמשותף לשניהם, וזה מה
 * שמצדיק שכן אחד: שניהם חייבים לקרות **לפני** שהבייטים נמסרים.
 *
 * שלישי הצטרף אליהם מאותו טעם: איחוד ריצות זהות, שבלעדיו מסמך Word שנערך
 * לאורך זמן מקליד לאט עד כדי לא שמיש (engine/docx-run-merge.ts).
 *
 * ## התיקון הראשון: ערך שמקפיא את המנוע
 *
 * `<w:defaultTabStop w:val="0"/>` ב-`word/settings.xml` שולח את בונה עצירות
 * הטאב של המנוע ללולאה שאינה נגמרת: הוא מקדם מיקום בצעדים של הערך הזה, וצעד
 * של אפס אינו מקדם. נמדד ב-superdoc 2.8.0 וגם ב-2.10.0 — מסמך Word ריק ותקין
 * שמכריחים בו את המאפיין הזה לאפס מקפיא את הדפדפן לצמיתות.
 *
 * הלולאה רצה על **החוט הראשי**, וזאת הנקודה שקובעת את הצורה של הקובץ הזה:
 * `OPEN_TIMEOUT_MS` ב-create-editor.ts אינו יכול להציל כאן, מפני שהטיימר שלו
 * צריך בדיוק את החוט שחסום. אין שום שעון-שמירה שיכול לתפוס את זה מבחוץ. הדרך
 * היחידה היא לא למסור למנוע את הערך מלכתחילה — כלומר כאן, לפני הפתיחה.
 *
 * ## התיקון השני: הדגשה של כתב מורכב שאינה מגיעה למסך
 *
 * `<w:bCs/>` היא „מודגש לכתב מורכב”, וזה מה ש-Word העברי כותב כשמדגישים בחירה
 * שכולה עברית: `w:bCs` לבדה, בלי `w:b`. Word מרנדר הדגשה של ריצה עברית מ-
 * `bCs`; המנוע מרנדר אותה מ-`w:b` בלבד. התוצאה היא מסמך שכותרותיו מודגשות
 * ב-Word ומגיעות אלינו דקות.
 *
 * נמדד ב-superdoc 2.10.0 (engine 0.9.0), על הקובץ שדווח, ב-Chrome אמיתי:
 * עשר הכותרות נצבעו `font-weight: 400`; אותו קובץ בדיוק עם `<w:b/>` שהושלם
 * לצד כל `<w:bCs/>` — כולן 700. ראו scripts/qa/bold-cs-qa.mjs.
 *
 * ### למה כאן, ולא בגיליון סגנונות
 *
 * קיפאון אין לתפוס אחר כך; הדגשה, לכאורה, כן. אלא שאין במה: ה-DOM שהמנוע
 * מצייר אינו נושא שום סימן ל-`bCs` — `.superdoc-text-run` מקבל `styleid` ו-
 * `style` מחושב עם `font-weight` סופי, ולא את מאפייני הריצה — ולכן אין למה
 * להיתלות בסלקטור. וגם אילו היה: המשקל נכנס לחישוב שבירת השורות ולעימוד,
 * והדגשה שנצבעת אחרי הפריסה מזיזה טקסט מתחת לפריסה שכבר חושבה. המנוע חייב
 * לדעת, ולכן זה נכתב לתוך ה-XML.
 *
 * ### מה זה עושה למסמך, במפורש
 *
 * `w:b` נכתב לצד `w:bCs`, וממילא גם ייוצא. על ריצה עברית אין לזה משמעות
 * ב-Word — הוא קורא `bCs` בכל מקרה — אבל על טקסט **לטיני שיושב באותו `rPr`**
 * זה שינוי אמיתי: הוא יוצג מודגש גם ב-Word, ולא היה. זה המחיר, והוא נבחר
 * ביודעין: מסמך עברי שכותרותיו אינן מודגשות הוא באג שהמשתמש רואה בכל עמוד,
 * ואנגלית בתוך כותרת של מסמך עברי היא מקרה קצה.
 *
 * `<w:b w:val="0"/>` **אינו** נהפך. היעדר `w:b` הוא מה ש-Word השאיר מאחור;
 * `w:b` מכובה הוא אמירה מפורשת של מי שכתב את הקובץ, ואין לנו רשות להפוך אותה.
 *
 * וההיקף הרחב ביותר של אותו מחיר: `bCs` על `w:rPrDefault` ב-`styles.xml`
 * מקבלת `w:b` בדיוק כמו כל `rPr` אחרת, ואז **כל** ריצה לטינית במסמך תהיה
 * מודגשת ב-Word אחרי שמירה — ברירת מחדל של המסמך, לא כותרת. זה הכלל עצמו
 * ולא חריג ממנו, וזה נאמר כאן כי הסקלה שונה. בקורפוס של 81 מסמכי Word אמיתיים
 * אין אף אחד שנושא `bCs` דולקת ב-`rPrDefault`.
 *
 * ## הכלל: לתקן, ולא לחסום
 *
 * כל כשל בדרך — zip שלא נקרא, חלק חסר, דחיסה שלא נתמכת — מחזיר את המקור כמות
 * שהוא. שלב מקדים שנכשל אינו אמור למנוע פתיחה של מסמך שהיה נפתח בלעדיו: הוא
 * תיקון, לא שער. לכן אין כאן אף מסלול שזורק.
 *
 * ## למה zip מלא ולא חיפוש-והחלפה על הבייטים
 *
 * החלקים דחוסים בתוך הארכיון, ולכן אי אפשר לגעת בהם בלי לפרוס אותם. כל חלק
 * שלא תוקן מועתק בייט-בבייט בדיוק כפי שהיה, וכך הקובץ שנמסר למנוע זהה למקור
 * בכל מה שאינו התיקון עצמו.
 *
 * ### החלק שתוקן נכתב דחוס, ורק אחרי שהדחיסה אומתה
 *
 * הגרסה הראשונה כתבה את החלק המתוקן כרשומה לא-דחוסה (`STORED`), מהטעם ש„דוחס
 * שמתנהג אחרת מהצפוי הוא באג שקט”. שני דברים שנמדדו הפכו את זה:
 *
 * 1. **המחיר לא היה „פי 2 עד פי 4”.** על 81 מסמכים אמיתיים, 34 תוקנו, וסך
 *    ה-Blob שנמסר למנוע גדל מ-68.8MB ל-359.7MB — **פי 5.2**, והגרוע פי 15.3:
 *    מסמך של 4.3MB הגיע למנוע כ-35.9MB. חמישה מסמכים חצו 17MB. וזה בדיוק על
 *    מסמכי הספרים שהם קהל היעד.
 * 2. **הבאג „השקט” אינו שקט.** ה-CRC נמדד על הבייטים הלא-דחוסים, והמנוע פורס
 *    את החלק מיד בפתיחה — deflate שגוי הוא כשל רועש בפתיחה, לא שינוי שקט
 *    במסמך. ובכל זאת אין כאן הסתמכות על זה: הפלט נפרס בחזרה ומושווה למקור
 *    (`deflateVerified`), ורק אז נכתב. דוחס שגוי נופל ל-`STORED`, לא למסמך.
 *
 * `CompressionStream('deflate-raw')` הוא אותו API שהקובץ הזה כבר משתמש בהופכי
 * שלו, ועם אותה נפילה-חזרה כשהוא חסר או נכשל: `STORED`, שהוא רשומת ZIP חוקית
 * לגמרי. אותם 59 `document.xml`: 337.1MB גלוי מול 25.9MB דחוס, ב-~41ms למסמך.
 *
 * זה זיכרון ולא קובץ — הייצוא דוחס מחדש מהמנוע, ולכן שום רשומה שנכתבה כאן
 * אינה מגיעה לדיסק. ורק על מסמכים שיש בהם מה לתקן: מסמך שאין בו `bCs` אינו
 * נכתב מחדש בכלל.
 */
import {
  CONTENT_PARTS,
  SKIPPED_SPANS,
  TOKEN_SOURCE,
  VAL_ATTRIBUTE,
  isOn,
  readEntryText,
  readZip,
  readZipComment,
  rewriteEntry,
  valueOf,
  writeZip,
  type Bytes,
  type ZipEntry,
} from './docx-parts';
import { mergeFragmentedRuns } from './docx-run-merge';
import { readDocumentVba, type DocumentVba } from './vba-import';

/** החלק שבו יושבות הגדרות המסמך. */
export const SETTINGS_PART = 'word/settings.xml';

/**
 * החלק שמתאר את הגופנים שהמסמך משתמש בהם — כולל כאלה שאינם מותקנים.
 *
 * אינו מתוקן כאן ואינו נוגע לקיפאון; הוא נקרא ונמסר החוצה, מפני שזה המקום
 * היחיד שכבר פותח את הארכיון. מה שנעשה איתו נמצא ב-engine/docx-fonts.ts.
 */
export const FONT_TABLE_PART = 'word/fontTable.xml';

/**
 * ברירת המחדל של Word לעצירת טאב, ב-twips (720 = חצי אינץ' = 1.27 ס"מ).
 *
 * זה גם מה ש-OOXML מגדיר כערך כשהמאפיין נעדר לגמרי, ולכן כתיבתו במפורש שקולה
 * למחיקת המאפיין — ומפורשת יותר למי שיפתח את הקובץ אחר כך.
 */
export const DEFAULT_TAB_STOP_TWIPS = 720;

/** שם האלמנט שהערך שבו מקפיא את המנוע. */
const DEFAULT_TAB_STOP_ELEMENT = 'defaultTabStop';

/**
 * מתקנת את `settings.xml`. `null` = אין מה לתקן.
 *
 * הבדיקה היא על **הערך** ולא על המחרוזת `"0"`: גם `-1` וגם ערך שאינו מספר
 * מגיעים למנוע כאותו צעד-אפס.
 *
 * ## למה זה סורק, ולא רגקס אחד
 *
 * הגרסה הראשונה כאן הייתה `/<w:defaultTabStop\b[^>]*\/>/`, וסקירה יריבה מצאה
 * בה **שלוש** דרכים שבהן הקיפאון פשוט אינו נמנע. לא „תיקון פחות מדויק” — הוא
 * לא קורה, והמשתמש נשאר עם אוצריא תקועה שאין ממנה יציאה:
 *
 * | הקלט | מה קרה |
 * |---|---|
 * | `<!-- …val="720"… --><w:defaultTabStop w:val="0"/>` | ההתאמה הראשונה היא זו שבהערה, ערכה תקין, ולכן `null` |
 * | `<w:defaultTabStop w:val="0"></w:defaultTabStop>` | XML חוקי לגמרי, אבל אינו נגמר ב-`/>` ולכן לא הותאם |
 * | `<w:defaultTabStop w:foo="a>b" w:val="0"/>` | `[^>]*` נעצר על ה-`>` שבתוך הערך |
 *
 * ובכיוון ההפוך, `<!-- <w:defaultTabStop w:val="0"/> -->` **תוקן בתוך ההערה**
 * — כלומר עריכה של טקסט שאינו עיצוב.
 *
 * לכן כאן אותו סורק של התיקון השני: מודע למרכאות, מדלג על הערות, CDATA
 * והוראות עיבוד, ואינו נעול על הקידומת `w`.
 *
 * ## ולמה נכתב רק **הערך**
 *
 * במקום להחליף את האלמנט כולו ב-`<w:defaultTabStop w:val="720"/>`. כך צורת
 * האלמנט נשמרת — סוגר-עצמו נשאר סוגר-עצמו, וזוג פתיחה-סגירה נשאר זוג — ואין
 * מסלול שבו נכתב תג סוגר-עצמו וסגירתו הקודמת נשארת תלושה. הקידומת ומרכאות
 * הערך נשמרות אף הן, מאותו טעם: כתיבה חוזרת של מה שלא היה צריך להשתנות היא
 * בדיוק מה שהמודול הזה מבטיח לא לעשות.
 */
export function repairSettings(xml: string): string | null {
  const token = new RegExp(TOKEN_SOURCE.source, 'g');
  for (let match = token.exec(xml); match; match = token.exec(xml)) {
    const closer = SKIPPED_SPANS.get(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      // הערה שאינה נסגרת. שום התאמה שאחריה אינה אמינה, ותיקון במקום הלא נכון
      // גרוע מאי-תיקון: הקיפאון לפחות משוחזר ונראה.
      if (end < 0) return null;
      token.lastIndex = end + closer.length;
      continue;
    }

    const [, closing, prefix, name, attributes] = match;
    if (closing || name !== DEFAULT_TAB_STOP_ELEMENT) continue;

    const current = valueOf(attributes);
    const twips = current === null ? Number.NaN : Number(current.trim());
    if (Number.isFinite(twips) && twips > 0) return null;

    const attributesAt = match.index + 1 + closing.length + prefix.length + 1 + name.length;
    if (current === null) {
      // אין `w:val` בכלל. הוא נדרש ב-`CT_TwipsMeasure`, ולכן זה מסמך שגוי —
      // אבל הכתיבה כאן היא בדיוק מה שהערך הנעדר אומר, ולכן היא גם התיקון.
      // נכנס לפני ה-`/` של תג סוגר-עצמו, ובסוף המאפיינים בכל צורה אחרת — ואחרי
      // שהרווח שלפני ה-`/` (אם יש) נדחף אל אחרי הערך, כדי לא לכתוב שניים.
      const before = attributes.slice(0, attributes.length - (attributes.endsWith('/') ? 1 : 0));
      const insertAt = attributesAt + before.trimEnd().length;
      return (
        xml.slice(0, insertAt) +
        ` ${prefix}:val="${DEFAULT_TAB_STOP_TWIPS}"` +
        xml.slice(insertAt)
      );
    }

    // מקום הערך עצמו: ההתאמה נגמרת במרכאת הסיום, ולכן הערך הוא האורך שלו
    // אחורה ממנה. כך המאפיין, הקידומת וסוג המרכאות נשארים תו-בתו.
    const inside = VAL_ATTRIBUTE.exec(attributes);
    if (!inside) return null;
    const valueEnd = attributesAt + inside.index + inside[0].length - 1;
    const valueStart = valueEnd - current.length;
    return (
      xml.slice(0, valueStart) + String(DEFAULT_TAB_STOP_TWIPS) + xml.slice(valueEnd)
    );
  }

  return null;
}

/** הכנסה אחת: המקום, והקידומת שבה לכתוב. */
interface BoldInsert {
  at: number;
  prefix: string;
}

/** מה שנאסף על ה-`rPr` החיה בזמן הסריקה. */
interface RunPropsScope {
  /** האם `b` מופיעה בה — בכל קידומת ובכל צורה, גם מכובה. */
  hasBold: boolean;
  /** מקום התג `bCs`, או `null` כשאינו שם. */
  boldCsAt: number | null;
  /** האם ה-`bCs` שנמצא דולק. */
  boldCsOn: boolean;
  /** הקידומת של ה-`bCs` שנמצא. */
  prefix: string;
}

/**
 * משלימה `<w:b/>` לצד כל `<w:bCs/>` דולקת שאין לצדה `w:b`. `null` = אין מה
 * לתקן.
 *
 * ## למה סורק ולא רגקס על `<w:rPr>…</w:rPr>`
 *
 * מפני ש-`rPr` **מקננת**: `<w:rPr>…<w:rPrChange><w:rPr>…</w:rPr></w:rPrChange>
 * </w:rPr>`. רגקס לא-להוט היה קושר את הפתיחה החיצונית לסגירה הפנימית, כלומר
 * מודד את התכונות של השינוי המסומן במקום של הריצה — ובדיוק במסמכים שיש בהם
 * מעקב שינויים, שהם המסמכים שאין רשות לשבור.
 *
 * ## מה מוגן, ואיך — הכלל שקובע
 *
 * **`rPr` שיושבת בתוך `rPr` אחרת אינה עיצוב חי, ומדולגת.** ב-`CT_RPr` האיבר
 * היחיד שמכיל `rPr` הוא `w:rPrChange`, ולכן קינון **הוא** ההגדרה של „זו
 * היסטוריה”. אין כאן זיהוי של `rPrChange` בשמו, ובכוונה: הניסיון הראשון ספר
 * `<w:rPrChange>` פתוחות מול סגורות, ומונה כזה נשבר משלוש דרכים שנמדדו — שם
 * מחבר שיש בו `>`, קידומת שאינה `w`, וסגירה תלושה שהורידה אותו מתחת לאפס ואז
 * כתבה **לתוך** ההיסטוריה הבאה. הגרועה מכולן: מונה שנתקע על 1 מכבה את התיקון
 * מאותו בייט ועד סוף החלק, בשקט.
 *
 * המחיר של הכלל הצר: `rPr` שיושבת בתוך `w:pPrChange` אינה מוגנת. `CT_PPrBase`
 * — ה-`pPr` שבתוך `pPrChange` — אינו מכיל `rPr` כלל, ולכן Word אינו יכול
 * לכתוב שם אחת. זה מה שהכלל מכסה ומה שאינו, בלי להבטיח יותר.
 *
 * ## סדר האלמנטים
 *
 * `<w:b/>` נכתב **מיד לפני** `<w:bCs/>`. `CT_RPr` היא רצף (`xsd:sequence`) ובו
 * `b` באה לפני `bCs`; נמדד גם על קורפוס: ב-10,855 `rPr` שנכתבו בידי Word ויש
 * בהן את שתיהן, `b` קדמה ב-100%, ו-0 בסדר ההפוך. סכימת OOXML לא נמצאה במכונה
 * (נסרקו ה-node_modules, הרפו, ושתי התקנות Office), ולכן זו עדות ולא הוכחה.
 */
export function repairComplexScriptBold(xml: string): string | null {
  // חיפוש מחרוזת אחד לפני הסריקה. `document.xml` של ספר הוא מגה-בייטים, ורובם
  // המכריע של המסמכים אינם נוגעים לזה בכלל.
  if (!xml.includes('bCs')) return null;

  const inserts: BoldInsert[] = [];
  /** עומק ה-`rPr`. 1 = העיצוב החי; 2 ומעלה = היסטוריה של שינוי מסומן. */
  let depth = 0;
  let scope: RunPropsScope | null = null;

  // רגקס חדש לכל קריאה, ולא אחד גלובלי שמאפסים לו `lastIndex`: `lastIndex` הוא
  // מצב, ומצב משותף בין קריאות הוא מלכודת שמחכה למי שיקרא מכאן פעמיים. המחיר
  // זניח — נקרא לכל היותר פעם אחת לחלק.
  const token = new RegExp(TOKEN_SOURCE.source, 'g');
  for (let match = token.exec(xml); match; match = token.exec(xml)) {
    const closer = SKIPPED_SPANS.get(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      // הערה שאינה נסגרת: אין יותר תגים שאפשר לסמוך עליהם, וחצי סריקה גרועה
      // מלא-סריקה. מה שנאסף עד כאן מוחל, וזה בטוח — הוא כולו מלפני ההערה.
      if (end < 0) break;
      token.lastIndex = end + closer.length;
      continue;
    }

    const [, closing, prefix, name, attributes] = match;
    const selfClosing = attributes.endsWith('/');

    if (name === 'rPr') {
      if (selfClosing) continue;
      if (!closing) {
        depth += 1;
        if (depth === 1) {
          scope = { hasBold: false, boldCsAt: null, boldCsOn: false, prefix };
        }
        continue;
      }
      // סגירה בלי פתיחה — מסמך קטוע. מתעלמים, ולא יורדים מתחת לאפס.
      if (depth === 0) continue;
      depth -= 1;
      if (depth === 0 && scope) {
        if (scope.boldCsOn && scope.boldCsAt !== null && !scope.hasBold) {
          inserts.push({ at: scope.boldCsAt, prefix: scope.prefix });
        }
        scope = null;
      }
      continue;
    }

    if (depth !== 1 || !scope || closing) continue;
    if (name === 'b') scope.hasBold = true;
    else if (name === 'bCs') {
      // המקום מה**ראשונה**, המצב מה**אחרונה** — ושני חצאי ההחלטה שונים
      // בכוונה. `rPr` עם שתי `bCs` פסולה מלכתחילה, אבל התוצאה אסור לה להיות
      // פסולה **יותר**: `CT_RPr` היא `xsd:sequence` שבו `b` בא לפני `bCs`,
      // ולכן הוספה לפני ה-`bCs` השנייה מציבה `b` **אחרי** `bCs` — כלומר חלק
      // שהסכימה של Word פוסלת. המצב נלקח מהאחרונה, כי בכפילות Word מכריע
      // „האחרון קובע”.
      if (scope.boldCsAt === null) scope.boldCsAt = match.index;
      scope.boldCsOn = isOn(attributes);
      scope.prefix = prefix;
    }
  }

  if (inserts.length === 0) return null;

  // הפיצול מניח סדר עולה, וזה מובטח מבנית: הכנסה נרשמת רק כשה-`rPr` **החיצונית**
  // נסגרת, והבאה אחריה נפתחת אחריה. הבדיקה „אינה משנה דבר מלבד ה-b שהוסיפה”
  // היא מה ששומר על ההנחה הזאת.
  //
  // מערך ו-`join` **לא בגלל מהירות** — נמדד, ואין הפרש. `+=` נראה מהיר פי
  // עשרה בהרכבה (V8 בונה cons-string ודוחה את השיטוח), אבל השיטוח נפרע
  // ב-`TextEncoder().encode` שקורא את התווים, ואז השניים שווים בתוך הרעש:
  // על 28,400 הכנסות ב-3.5MB נמדדו 31 מול 27ms בהרצה אחת ו-48 מול 50ms
  // בשנייה. מדידה של ההרכבה בלבד כאן היא מדידה של שום דבר.
  //
  // מה שכן: הצורה הזאת אומרת „חתוך, הכנס, חתוך” במקום לתחזק מחרוזת מצטברת.
  const parts: string[] = [];
  let at = 0;
  for (const insert of inserts) {
    parts.push(xml.slice(at, insert.at), `<${insert.prefix}:b/>`);
    at = insert.at;
  }
  parts.push(xml.slice(at));
  return parts.join('');
}

/**
 * מה שנאמר למשתמש כשהדגשת כתב מורכב הושלמה.
 *
 * מנוסח כמצב ולא כשגיאה, ואומר גם את מה שיקרה אחר כך — השמירה תכתוב את
 * ההשלמה לקובץ, וזה החלק שהמשתמש אינו יכול לנחש.
 */
export const COMPLEX_SCRIPT_BOLD_NOTICE =
  'הדגשה שמוצגת ב-Word ולא הוצגה כאן הושלמה במסמך — שמירה תכתוב אותה לקובץ';

/** תיקון אחד: על אילו חלקים הוא חל, מה הוא עושה, ומה נרשם ביומן כשעשה. */
interface PartRepair {
  matches(name: string): boolean;
  /** `null` = אין מה לתקן בחלק הזה. */
  repair(xml: string): string | null;
  note(name: string): string;
  /**
   * מה שנאמר למשתמש בשורת המצב כשהתיקון הוחל, או `null` כשאין מה לומר לו.
   *
   * שדה מפורש ולא גזירה מנוסח ה-`note`: שורת המצב שהמשתמש רואה אינה אמורה
   * להיות תלויה במחרוזת שנכתבת ליומן, ושינוי נוסח ביומן אינו אמור להעלים
   * אותה בשקט.
   */
  notice: string | null;
}

/**
 * התיקונים, בסדר שבו הם מוחלים. השלישי — איחוד ריצות — אינו תיקון של קריאה
 * שגויה אלא של מחיר: ראו engine/docx-run-merge.ts.
 */
const REPAIRS: PartRepair[] = [
  {
    // ללא תלות ברישיות, כמו `CONTENT_PARTS`. שני התיקונים ניגשים לאותו
    // ארכיון, ו„`Word/Settings.xml` מקבל אחד ולא את השני” הוא חוסר עקביות
    // שאין לו טעם — גם אם אין כלי שכותב כך.
    matches: (name) => name.toLowerCase() === SETTINGS_PART,
    repair: repairSettings,
    note: () =>
      `${SETTINGS_PART}: defaultTabStop מתוקן ל-${DEFAULT_TAB_STOP_TWIPS} — הערך שהיה מקפיא את המנוע`,
    // `defaultTabStop` אינו נראה למשתמש בשום צורה — הוא רק ההפרש בין מסמך
    // שנפתח למסמך שקופא — ושורת מצב שמדווחת על כל דבר היא שורת מצב שאיש אינו
    // קורא.
    notice: null,
  },
  {
    matches: (name) => CONTENT_PARTS.test(name),
    repair: repairComplexScriptBold,
    note: (name) => `${name}: <w:b/> הושלם לצד <w:bCs/> — בלעדיו הדגשת עברית אינה מגיעה למסך`,
    notice: COMPLEX_SCRIPT_BOLD_NOTICE,
  },
  {
    // אחרון, ובכוונה: ההשוואה בין ריצות היא על ה-`rPr` כמחרוזת, ולכן היא
    // צריכה לראות את ה-`rPr` אחרי שהתיקון הקודם השלים בו `w:b`.
    matches: (name) => CONTENT_PARTS.test(name),
    repair: mergeFragmentedRuns,
    note: (name) => `${name}: ריצות סמוכות בעיצוב זהה אוחדו — המנוע משלם על כל ריצה בכל הקשה`,
    // אינו נראה: אותו טקסט, אותו עיצוב. ראו engine/docx-run-merge.ts.
    notice: null,
  },
];

/** מה שהשלב המקדים מוציא: המסמך שיימסר למנוע, ומה שנקרא עליו בדרך. */
export interface PreflightPackage {
  /** null = אין שינוי; ה-Blob המקורי נשאר אצל הקורא, בלי להעביר אותו בחזרה. */
  bytes: Bytes | null;
  notes: string[];
  /** תוכן `FONT_TABLE_PART`, או `null` כשאין או שלא נקרא. */
  fontTable: string | null;
  /**
   * המאקרו שבמסמך — לקריאה בלבד (engine/vba-import.ts).
   *
   * כאן מאותו טעם כמו טבלת הגופנים: הבייטים כבר נקראו והארכיון כבר נפתח, וזה
   * גם **הרגע** הנכון — הידיעה שבמסמך יש מאקרו שWord מריץ בפתיחה שייכת לזמן
   * הפתיחה, לא לזמן שבו המשתמש יחשוב לחפש אותה.
   */
  vba: DocumentVba;
  /**
   * הודעה למשתמש כשהמסמך תוקן, או `null` כשלא נגענו בו.
   *
   * **למה זה יוצא מכאן ולא נשאר ב-`console`:** התיקון השני נכתב לתוך המסמך,
   * וממילא ייצא איתו בשמירה. משתמש שהמסמך שלו שונה זכאי לדעת שהוא שונה, וזה
   * אינו מידע שאפשר להשאיר ביומן שאיש אינו פותח. אותו מסלול בדיוק שבו
   * `vba.status` מגיע לשורת המצב (App.vue).
   *
   * `notes` המפורטים נשארים ביומן: שורת מצב אינה המקום לשמות חלקים.
   */
  notice: string | null;
}

/**
 * עיבוד החבילה בלבד: ללא DOM, רשת, Blob או חיבור לעורך. זהו אותו קוד ב-worker
 * ובמסלול ההתאוששות המקומי. טבלת הגופנים והתיקונים חולקים ספריית ZIP אחת.
 */
export async function preflightPackage(bytes: Bytes): Promise<PreflightPackage> {
  // **ספרייה מרכזית אחת לשני הקוראים.** קודם `readDocxPart` פתח את הארכיון
  // בשביל טבלת הגופנים ו-`preflightDocx` פתח אותו שוב בשביל התיקונים; אותה
  // ספרייה, אותם בייטים, פעמיים. הרשומות הן תצוגות (`subarray`) אל אותם
  // בייטים ואף אחד מהשניים אינו כותב לתוכן — `preflightDocx` בונה מערך חדש
  // — ולכן שיתוף כאן אינו קושר ביניהם.
  //
  // המאקרו אינו שותף: `extractVbaFromDocx` הוא של superdoc-macros, עם קורא
  // ZIP משלו ובלי API שמקבל רשומות שנקראו. הוא גם כבר עושה את הזול קודם —
  // מסמך בלי מאקרו מקבל תשובה מיד אחרי חיפוש החלק, בלי לפרוס דבר — ולכן
  // „לבדוק תחילה אם קיים חלק VBA” היה פותח את החבילה **פעם נוספת**, לא
  // חוסך פתיחה.
  const entries = readZip(bytes);

  const fontTable = await partText(entries, FONT_TABLE_PART);
  // על בייטי המקור ולא על המתוקנים: התיקונים נוגעים לחלקי ה-XML של הגוף
  // והסגנונות, ואין טעם לקרוא את המאקרו מעותק שנכתב מחדש.
  const vba = await readDocumentVba(bytes);

  let repaired: DocxRepair | null;
  try {
    repaired = await repairEntries(entries, readZipComment(bytes));
  } catch (error) {
    // „לתקן, ולא לחסום” גם כאן: הזריקה היחידה שנשארה בפנים היא הקצאה של
    // ארכיון גדול מדי, ומסמך שהיה נפתח בלי השלב הזה ייפתח בלעדיו.
    console.warn('[otzaria-word] הבדיקה המקדימה נכשלה, והמסמך נמסר כמות שהוא', error);
    return { bytes: null, notes: [], fontTable, vba, notice: null };
  }
  if (!repaired) return { bytes: null, notes: [], fontTable, vba, notice: null };

  return {
    bytes: repaired.bytes,
    notes: repaired.notes,
    fontTable,
    vba,
    notice: repaired.notice,
  };
}

/**
 * תוכן חלק מתוך ה-DOCX כטקסט. `null` כשאינו קיים, כשהארכיון אינו נקרא, או
 * כשהדחיסה אינה נתמכת — שלושה מקרים שבהם פשוט אין לנו מה לומר עליו.
 */
export async function readDocxPart(bytes: Bytes, name: string): Promise<string | null> {
  return partText(readZip(bytes), name);
}

/** אותו דבר, על ספרייה שכבר נקראה. ראו `preflightPackage`. */
async function partText(
  entries: readonly ZipEntry[] | null,
  name: string,
): Promise<string | null> {
  const entry = entries?.find((candidate) => candidate.name === name);
  return entry ? readEntryText(entry) : null;
}

/** מסמך שתוקן: הבייטים שיימסרו למנוע, ומה שנעשה בהם. */
export interface DocxRepair {
  bytes: Bytes;
  /** שורה לכל תיקון שהוחל, לפי הסדר. ריק אינו אפשרי — בלי תיקון אין תוצאה. */
  notes: string[];
  /**
   * מה שיש לומר למשתמש, או `null` כשכל מה שתוקן אינו נראה לו. רק התיקון
   * שנכתב לתוך המסמך מגיע לשורת המצב — ראו `PartRepair.notice`.
   */
  notice: string | null;
}

/**
 * בייטי DOCX מתוקנים, או `null` כשאין מה לתקן ואין מה לדווח.
 *
 * מיוצאת בנפרד מ-`preflightPackage` כדי שאפשר יהיה לבדוק אותה בלי רשת ובלי
 * `Blob` — היא כל הלוגיקה שיש כאן.
 *
 * חלק שאינו נקרא (דחיסה שאינה נתמכת, נתונים שאינם נפרסים) מדולג ואינו מפיל את
 * שאר התיקונים: אין קשר בין החלקים, ומסמך שאחד מהם אינו נקרא אינו סיבה למסור
 * את השני שבור.
 */
export async function preflightDocx(bytes: Bytes): Promise<DocxRepair | null> {
  return repairEntries(readZip(bytes), readZipComment(bytes));
}

/** התיקונים עצמם, על ספרייה שכבר נקראה. ראו `preflightPackage`. */
async function repairEntries(
  entries: readonly ZipEntry[] | null,
  archiveComment: Bytes,
): Promise<DocxRepair | null> {
  if (!entries) return null;

  const notes: string[] = [];
  let notice: string | null = null;
  const patched = new Map<ZipEntry, ZipEntry>();

  for (const entry of entries) {
    const repairs = REPAIRS.filter((candidate) => candidate.matches(entry.name));
    if (repairs.length === 0) continue;

    const original = await readEntryText(entry);
    if (original === null) continue;

    let xml = original;
    for (const repair of repairs) {
      const next = repair.repair(xml);
      if (next === null) continue;
      xml = next;
      notes.push(repair.note(entry.name));
      notice ??= repair.notice;
    }
    if (xml !== original) {
      patched.set(entry, await rewriteEntry(entry, new TextEncoder().encode(xml)));
    }
  }

  if (patched.size === 0) return null;

  const rewritten = entries.map((entry) => patched.get(entry) ?? entry);
  return { bytes: writeZip(rewritten, archiveComment), notes, notice };
}
