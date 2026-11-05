/**
 * חצים אופקיים בשורה עברית — „לוחץ ימין והסמן הולך שמאלה”.
 *
 * ## מה שבור במנוע
 *
 * המנוע בוחר מודל כיוון **לכל פסקה בנפרד**: בפסקה עברית נקייה החצים חזותיים
 * (ימין=ימינה, כמו Word), ואילו בפריט רשימה, בפסקה עם טאב, ובכל שורה שיש בה
 * ספרות או לטינית הם לוגיים — כלומר ימין מזיז את הסמן שמאלה על המסך. אותו
 * מקש, שתי התנהגויות הפוכות, באותו מסמך. ‏`Shift+חץ` יורש את הסתירה, ובגבול
 * בין פסקה „חזותית” ל„לוגית” הסמן מקפץ בין השתיים ואינו חולף.
 *
 * וכך גם בחירה: בפסקה „לוגית” `Shift+שמאל` הרחיב ימינה, וחץ רגיל על בחירה
 * קיימת קרס לקצה ההפוך. ההסבר ב„בחירה” שלמטה.
 *
 * מדוד ומתועד ב-`docs/engine-gaps.md` („חצים אופקיים בעברית”), ודווח למעלה
 * כ-[#3996](https://github.com/superdoc/docx-editor/issues/3996). הקוד שמזיז
 * את הסמן אינו בחלקים הפתוחים של SuperDoc, ולכן אין PR — רק יירוט.
 *
 * ## למה אפשר לתקן את זה מבחוץ, אף שנכתב כאן קודם שאי אפשר
 *
 * הטענה הקודמת הייתה ש„היעד הנכון תלוי במודל הדו-כיווני שהמנוע מחזיק ואיננו
 * רואים”. זה לא נכון, ונמדד כלא נכון: הטקסט מצויר ב-DOM אמיתי, ולכן **המיקום
 * המצויר של כל תו ניתן למדידה מכאן**. אין צורך במודל של המנוע — צריך רק לדעת
 * מי התו שמצויר מימין לסמן, וזו שאלה על מלבנים.
 *
 * הגשש `scripts/qa/rtl-caret-visual-feasibility.mjs` הוכיח את זה מול האמת:
 * לכל היסט בשורה נכתבה בחירה, נמדד ה-x שהמנוע צייר בו את הסמן, והושווה למה
 * שהחישוב כאן גוזר מה-DOM. התוצאה (superdoc 2.15.0, Chrome, ה-dist הארוז):
 *
 *   | השורה                        | התאמה     |
 *   |------------------------------|-----------|
 *   | עברית נקייה                  | **20/20** |
 *   | פריט ברשימה ממוספרת          | **18/18** |
 *   | עברית עם ספרות               | **12/12** |
 *   | בקרה: שורה לטינית            | **17/17** |
 *   | שורה שנגמרת בספרה            | **13/13** |
 *   | שורה שמתחילה באי לטיני       | 12/13     |
 *   | עברית עם מילה באנגלית        | 12/14     |
 *
 * זהה עד עשירית פיקסל, וכל אי-התאמה היא **היסט תפר** — בדיוק אלה שמדולגים
 * כאן. שתי השורות האחרונות שבטבלה נוספו ב-17.9.2026, כשהוטל ספק בחריץ הראשון
 * ובחריץ האחרון של שורה; ראו `caretSlots`.
 *
 * ## המחיר — מה נמדד, ומה לא
 *
 * מה שנכתב כאן קודם („חציון 0.1–0.5ms, הגרוע 4.8ms”) היה מדידה של **רכיב אחד**
 * מתוך המסלול — חילוץ התיבות של שורה אחת — והוצג כמחיר ההקשה. נמדד מחדש
 * ב-17.9.2026 בשתי רמות, 40 חזרות לכל שורה:
 *
 *   | מה נמדד                                   | חציון      | הגרוע  |
 *   |-------------------------------------------|------------|--------|
 *   | חילוץ התיבות של שורה (`charBoxes`)        | 0–0.8ms    | 1.1ms  |
 *   | הקשה שלמה: `keydown` אמיתי דרך כל המסלול  | 0.1–7.7ms  | 17.8ms |
 *
 * ההקשה השלמה כוללת את הסמן המצויר ומלבנו, שלוש הסריקות על המארח,
 * `lineDistance` לכל מועמדת, `readLineChars` (ובקצה — פעמיים), ואת החלק
 * הסינכרוני של `setSelectionTarget`. היא **אינה** כוללת את הציור עצמו, שהוא
 * א-סינכרוני. שתי ריצות על אותו build נתנו לשורה הכבדה (88 תווים, רוחב A4
 * מלא) חציון 4.7ms והגרוע 11.8ms, ואז חציון 7.7ms והגרוע 17.8ms — כלומר
 * פריים (16ms) הוא סדר הגודל ולא גבול בטוח, על מכונה עמוסה.
 *
 * וזה רץ **רק בהקשת חץ**; הקלדה אינה נוגעת בו. אבל חץ הוא המקש שמחזיקים,
 * ולכן „רק בהקשה” אינו גבול — ולכן נמדד גם מה עולה חץ **מוחזק**, עם
 * ‏`PerformanceObserver` על `event` ועל `longtask` ומונה פריימים:
 *
 *   | מה                                   | פסקה עברית | פריט רשימה |
 *   |--------------------------------------|------------|------------|
 *   | הקשות בשנייה                         | 51         | 46         |
 *   | עיבוד ה-`keydown` השלם, טיפוסי       | 5–8ms      | 5–6ms      |
 *   | הגרוע באותה ריצה                     | 17ms       | 11ms       |
 *   | ‏`longtask` (מעל 50ms)                | 0          | 0          |
 *   | פריימים ב-1.2 שניות (במנוחה: 74)     | **74**     | **74**     |
 *
 * כלומר בהחזקה רצופה, מעל קצב ה-autorepeat של Windows, **לא נפל אף פריים**
 * ולא נרשמה אף משימה ארוכה. ‎~30% מהחוט הראשי בזמן ההחזקה — וזה **גזירה**
 * מהמספרים שבטבלה, לא מדידה נפרדת.
 *
 * (אותה מדידה רצה קודם, בטעות, על superdoc 2.14.0-next.5 שנשאר ב-
 * `node_modules` של העץ הראשי, ונתנה את אותן התאמות ואת אותם ערכי x בדיוק.)
 *
 * ## שלושה דברים שהמדידה מצאה, והם מה שמעצב את הקוד כאן
 *
 * 1. **כיוון הוא תכונה של מה שמצויר, לא של מה שהוצהר.** „פרק 12 בספר” הוא
 *    ריצת OOXML **אחת** שבתוכה אי לטיני; חלוקה לפי אלמנטים או לפי `w:rtl`
 *    מחמיצה אותו. לכן הכיוון כאן נגזר מ**צמידות**: שני תווים עוקבים שהתיבות
 *    שלהם נוגעות זו בזו שייכים לאותו קטע, והצד שבו הן נוגעות הוא הכיוון.
 *    ראו `bidi-declaration-is-not-direction` בהיסטוריה.
 *
 * 2. **לתפר יש שני בתים, ואין ידית לבחור ביניהם.** בגבול בין עברית ללטינית
 *    להיסט לוגי אחד יש **שני** מיקומים חוקיים על המסך, והמנוע בוחר לפי איך
 *    הגעת. נמדד ב„מילה ABC מילה”: היסט 56 מצויר ב-980 אחרי לחיצה בעכבר,
 *    וב-1012.9 אחרי כתיבה דרך `setSelectionTarget` — שמקבל מספר בלי צד.
 *    נמדדו חמש צורות כתיבה — עוגן מכל צד, `collapse` לשני הכיוונים — וכולן
 *    החזירו את **אותו** x בדיוק. אין ידית.
 *
 *    ומכאן ההכרעה: **היסט שיושב על תפר אינו יעד.** אילו כיוונו אליו, המנוע
 *    היה מצייר אותו בבית שהוא בחר — ובאי לטיני זה הבית שבצד השני, כלומר
 *    „שמאלה” היה מזיז ימינה. זו בדיוק התקלה שבאנו לתקן. המחיר הוא עצירה אחת
 *    בכל קצה של אי לועזי (עדיין נגישה בלחיצת עכבר), והתמורה היא
 *    **מונוטוניות**: ימין לעולם אינו מזיז שמאלה.
 *
 *    ‏**שני קצות השורה אינם תפר כזה** — זה נמדד בנפרד, ראו `caretSlots`:
 *    בשורה שמתחילה או נגמרת באי לועזי המנוע מצייר את ההיסט הראשון והאחרון
 *    בבית של **התו**, בדיוק במקום שמחושב כאן.
 *
 * 3. **כל אלמנט מצויר נושא את טווח ה-pm שלו.** ‏`SPAN.superdoc-text-run` וגם
 *    `SPAN.superdoc-tab` — והטאב הוא יחידת pm אחת בלי טקסט (נמדד:
 *    `pm=5..6`, רוחב 24, אפס תווים). לכן ההיסט של כל תו נקרא מהאלמנט שלו
 *    ואינו נספר באינדקסים, וסמן המספור של פריט רשימה — שאינו נושא טווח —
 *    נופל מעצמו.
 *
 * ## קצוות
 *
 * בקצה השורה אין חריץ בכיוון המבוקש, ואז היעד הוא **השכן הלוגי** בשורה או
 * בפסקה הסמוכה: ימין בשורה עברית הוא אחורה לוגית, ולכן הוא מוביל לסוף השורה
 * שמעליה — שהוא קצה שמאל שלה. בלי זה ההקשה הייתה חוזרת למנוע, והוא מזיז שם
 * קדימה לוגית, כלומר **שמאלה על המסך** — נמדד בשער.
 *
 * מה שנשאר למנוע: תנועה **בתוך** תא טבלה ובכותרת עליונה, פסקה שכנה שאינה
 * צמודה במסמך, עמוד שעוד לא צויר, שורה שאינה מוכרזת `dir="rtl"` — שם הוא
 * נמדד תקין —
 * ושורה שטווחי ה-pm המצוירים בה אינם רציפים. בכל אלה ההקשה עוברת אליו כפי
 * שהיא. שער: `npm run check:arrows`.
 *
 * וגם **`Ctrl/Alt/Meta+חץ`** — קפיצת מילה — נשאר לו, ובמודע: היא דורשת גבולות
 * מילה, ומפת התיבות כאן אינה יודעת עליהן דבר. המחיר הוא סתירה חדשה: באותו
 * פריט רשימה `ArrowRight` יזוז ימינה בעוד `Ctrl+ArrowRight` יקפוץ מילה
 * שמאלה (לפני המודול הזה שניהם זזו שמאלה). מתועד ב-`docs/engine-gaps.md`.
 *
 * ## בחירה
 *
 * נמדד (4.11.2026, superdoc 2.23.0, `scripts/qa/rtl-selection-arrow-probe.mjs`)
 * מהיסט 9 ב„פרק 12 בספר הזה”: `Shift+שמאל` נתן 8..9, כלומר **ימינה** על המסך,
 * ו-`Shift+ימין` נתן 9..10; ובחירה 6..9 שנגמרה ב`חץ ימין` קרסה ל-9, הקצה
 * השמאלי. בפסקה עברית נקייה כל אלה תקינים — אותה סתירה בדיוק כמו בסמן.
 *
 * מה שעמד בדרך היה הקריאה: התצלום הסינכרוני מחזיר `selectionTarget: null` לכל
 * טווח, ו-`doc.selection.current()` מחזיר `start`/`end` בסדר המסמך, בלי כיוון.
 * בכתיבה דרך `setSelectionTarget`, לעומת זאת, start הוא העוגן ו-end הוא הראש;
 * שומרים על הסדר הזה גם בבחירה לאחור, כדי שההקשה הבאה של המנוע תמשיך מאותו ראש.
 *
 *   • **בחירה שאנחנו בונים — זוכרים.** העוגן והראש מוחזקים ב-`HeldSelection`,
 *     וההקשה הבאה ממשיכה מהראש שנזכר בלי לשאול את המנוע. „נשכח” בכל מקש אחר,
 *     בכל לחיצת עכבר, בהחלפת עורך ובסגירת ההתקנה. גם קריאה שבדרך מתבטלת.
 *   • **גרירה באותה שורה.** את הקצוות קוראים פעם אחת בסיום הגרירה, דרך
 *     `doc.selection.current()`. רק אם שניהם ניתנים למיפוי באותה שורה,
 *     כיוון הגרירה מול מיקומם המצויר מזהה את העוגן — גם בתוך אי לטיני.
 *     אין קריאה או המתנה בתוך keydown: מקש שהגיע לפני התשובה עובר למנוע,
 *     והתשובה המאוחרת נפסלת. אין תור שיכול לבלוע הקשה בזמן קריסה.
 *   • **טווח שכיוונו לא ידוע** (למשל אחרי `Shift+Home/End`, `Ctrl+Shift+חץ`
 *     או פקודה) נשאר של המנוע ב-`Shift+חץ`. ה-API הציבורי אינו מגלה מי הראש,
 *     ואין מניחים שהוא סוף הטווח. גם חץ רגיל על טווח חיצוני שאינו מוכר
 *     נשאר של המנוע, במקום לבלוע אותו לפני שידוע אם אפשר להגיע לקצה.
 *
 * מה שנשאר למנוע: `Shift+חץ` בשורה לטינית כשהבחירה לא התחילה אצלנו, טווח
 * שכיוונו לא ידוע, גרירה מרובת שורות או עם Shift, קצה לא מצויר או תפר BiDi,
 * וכל מה שלמעלה כ„נשאר למנוע”. קריסה בטווח מוכר באותה שורה משווה את ה-x
 * של הקצוות; בין שורות הולכים לפי סדר המסמך וכיוון השורה.
 *
 * ## פסקה ריקה
 *
 * אין בה תו מצויר, ולכן אין ממה לגזור חריץ — אבל יש בה בדיוק מקום סמן אחד,
 * וההקשה שייכת לנו. נמדד בשער שמסירה למנוע כאן אינה ניטרלית: מתחילת הפסקה
 * שאחרי הריקה `ArrowRight` הזיז אותו **בתוך** אותה פסקה מהיסט 0 להיסט 1,
 * כלומר שמאלה על המסך, והריקה דולגה. הצורה המצוירת נמדדה: fragment שטווחו
 * באורך אפס (‏`14..14`) ובתוכו שורה באותו טווח בלי אף נושא היסט — וכן, פסקה
 * ריקה תופסת מקום pm נוסף אחריה (‏`1..13`, `14..14`, `16..28`).
 *
 * ## תחילת המסמך וסופו — ההקשה נבלעת
 *
 * שם אין לאן לזוז, ומסירה למנוע אינה ניטרלית: בפסקה שהמנוע מזיז בה לוגית הוא
 * זז לכיוון ההפוך, ההקשה הבאה שלנו מחזירה, והסמן קפץ בין שני היסטים בלי סוף
 * (נמדד בשתי הקצוות). הקצה מזוהה מהעמודים, ולא מהיעדר פסקה ב-DOM: המנוע
 * מצייר רק עמודים קרובים (נמדד: 3 מתוך 25), ולכן „אין פסקה אחרי” לבדו אינו
 * סוף המסמך.
 *
 * ## פסקה שנחצית בין עמודים
 *
 * כל חלק שלה הוא fragment משלו, וההמשך מתחיל ב-pm שבו הקודם נגמר (נמדד:
 * 1214..2112 ו-2112..2383, עם `data-continues-from-prev`). ההיסט של הסמן
 * יחסי לתחילת **הפסקה**, ולא לתחילת ה-fragment, ולכן הבסיס נלקח מהחלק
 * הראשון; בגבול עצמו — pm אחד ששייך לשניהם — ה-y של הסמן מכריע, כמו בין
 * שורות גולשות. בלי זה החלק השני נמסר כולו למנוע, והחץ בגבול קפץ לתחילת
 * השורה האחרונה בעמוד הקודם (נמדד: 898 → 811).
 *
 * ## טבלה שכנה — נכנסים לתא
 *
 * טבלה מצוירת כ-fragment בלי מזהה מקור, והפסקאות שבתאים שלה הן אלמנטים עם
 * מזהה אבל בלי טווח משלהם — הטווח על השורות (נמדד: הטבלה pm 115..122, ובתוכה
 * פסקת התא עם שורה 115..122; הפסקה שלפני נגמרת ב-114 וזו שאחרי מתחילה
 * ב-123). מסירה למנוע אינה עובדת כאן בפסקה לוגית: הוא זז קדימה בתוך הפסקה
 * ולא נכנס לטבלה (נמדד). לכן היעד הוא קצה הפסקה הראשונה בטבלה (קדימה) או
 * האחרונה (אחורה). התנועה **בתוך** התא עצמו נשארת למנוע.
 *
 * ## איפה הקריאה עצמה
 *
 * כל מה שקורא את מה שהמנוע צייר — מבנה העמודים והשורות, טווחי ה-pm, התיבה של
 * כל גרפמה, הכיוון שלה והחריצים שנגזרים ממנה — יושב ב-`painted-lines.ts`,
 * ומשותף עם `rtl-vertical-caret.ts` (החצים האנכיים). כאן נשארה ה**מדיניות**
 * של החצים האופקיים בלבד: מה „ימינה” אומר בשורה עברית, מתי היעד הוא השכן,
 * ומתי ההקשה נמסרת למנוע. ההסברים והמדידות של הקריאה עברו איתה, צמודים
 * לפונקציות שהם מתארים.
 */

import {
  CARET_SELECTOR,
  atDocumentEdge,
  blockIdOf,
  blockStart,
  blocksInside,
  caretSlots,
  contiguousWith,
  emptyLineSlot,
  findFragment,
  findPaintedLine,
  isEmptyRange,
  lineAt,
  lineDirectionUnder,
  linesOf,
  neighbourOf,
  readLineChars,
  type CaretSlot,
} from './painted-lines';
import { readDocSelection, type SelectionDocumentApi } from './doc-selection';

/**
 * החריץ הקרוב ביותר בכיוון המבוקש **על המסך**, או `null` כשאין כזה בשורה.
 *
 * נקודת המוצא היא ה-x שהסמן **מצויר** בו ולא החריץ של ההיסט שלו: בתפר השניים
 * אינם זהים, וה-x המצויר הוא האמת. כך טעות בצעד אחד אינה נצברת — כל הקשה
 * מודדת מחדש מאיפה שהסמן באמת נמצא.
 */
export function visualTarget(
  slots: readonly CaretSlot[],
  caretX: number,
  toRight: boolean,
): CaretSlot | null {
  let best: CaretSlot | null = null;
  for (const slot of slots) {
    if (toRight ? slot.x <= caretX + 0.5 : slot.x >= caretX - 0.5) continue;
    if (!best || (toRight ? slot.x < best.x : slot.x > best.x)) best = slot;
  }
  return best;
}

/**
 * הכניסה לשורה או לפסקה הסמוכה: הקצה הלוגי שלה מהצד שממנו באים.
 *
 * `caretPm` מוחרג מפני שגבול בין שתי שורות גולשות הוא **היסט אחד** ששייך
 * לשתיהן — בלי ההחרגה הצעד היה מחזיר את הסמן לעצמו.
 */
export function edgeSlot(
  slots: readonly CaretSlot[],
  forward: boolean,
  caretPm: number,
): CaretSlot | null {
  let best: CaretSlot | null = null;
  for (const slot of slots) {
    if (slot.pm === caretPm) continue;
    if (!best || (forward ? slot.pm < best.pm : slot.pm > best.pm)) best = slot;
  }
  return best;
}

/**
 * האם התנועה המבוקשת היא **קדימה לוגית**.
 *
 * בשורה עברית „ימינה על המסך” הוא אחורה בטקסט, ובשורה לטינית להפך. זה כל מה
 * שצריך כדי לדעת לאיזו שורה או פסקה לעבור בקצה.
 */
export function movesForward(toRight: boolean, lineRtl: boolean): boolean {
  return toRight ? !lineRtl : lineRtl;
}

/**
 * האם ההקשה היא חץ אופקי — בלי Ctrl/Meta/Alt, עם Shift או בלעדיו.
 *
 * ‏**מה שנכתב כאן קודם היה שגוי, ותוקן ממדידה** — פעמיים. הנימוק הראשון
 * להשארת `Shift+חץ` למנוע היה „בחירה שאינה מכווצת אינה מצוירת כלל אחרי כתיבה
 * דרך `setSelectionTarget`”. הוא נסמך על חיפוש **אלמנט** ששמו מכיל `selection`,
 * וזו אינה שאלה על פיקסלים: אותו חיפוש מחזיר ריק גם כשהמנוע עצמו מצייר את
 * הבחירה. נמדד מחדש ב-17.9.2026 (superdoc 2.15.0) בצילום מלבן הפסקה והשוואת
 * בתים —
 *
 *   | מה נעשה                              | התמונה |
 *   |--------------------------------------|--------|
 *   | בקרת יציבות: אותו מצב פעמיים         | זהה    |
 *   | ‏`Shift+חץ` של המנוע                  | השתנתה |
 *   | ‏`Shift+End` — כתיבה דרך ה-API        | השתנתה |
 *   | כתיבה ידנית של טווח 2..8 דרך ה-API   | השתנתה |
 *
 * כלומר טווח שנכתב דרך ה-API **כן** מצויר. והנימוק השני, „אין מאיפה לקרוא את
 * ראש הבחירה בהקשה הבאה”, נכון לגבי התצלום אבל לא חוסם — ראו „בחירה” בראש
 * הקובץ.
 */
export function isHorizontalArrow(
  event: Pick<KeyboardEvent, 'key' | 'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey'> &
    Partial<Pick<KeyboardEvent, 'isComposing' | 'keyCode'>>,
): boolean {
  if (event.key !== 'ArrowRight' && event.key !== 'ArrowLeft') return false;
  // הרכבה במקלדת (IME): החץ שייך לחלונית ההרכבה, ולא לסמן. `swallow()` כאן
  // היה בולע אותו לגמרי (הוא עוצר גם את המאזינים האחרים) ומתרגם אותו לתנועת
  // סמן — כלומר שובר את ההרכבה. התקן של הריפו ב-`src/ui/shortcuts/match.ts`:
  // דפדפן שאינו מציב `isComposing` מדווח `keyCode === 229`.
  if (event.isComposing === true || event.keyCode === 229) return false;
  return !event.ctrlKey && !event.metaKey && !event.altKey;
}

/* ------------------------------------------------------------------ */
/* קריאת מה שמצויר                                                      */
/* ------------------------------------------------------------------ */

/** נקודת קצה של בחירה, כפי שהמנוע מחזיר אותה. */
interface TextPoint {
  kind?: string;
  blockId?: string;
  offset?: number;
  story?: unknown;
}

interface LiveSelectionSnapshot {
  selectionTarget?: {
    kind?: string;
    start?: TextPoint | null;
    end?: TextPoint | null;
    story?: unknown;
    coordinateSpace?: string;
  } | null;
}

/** מה שנצרך מ-`activeEditor`. מוגדר כאן ולא מיובא — ההסבר ב-document-api.ts. */
export interface RtlCaretEditor {
  host?: {
    readLiveSelectionSyncSnapshot?: () => LiveSelectionSnapshot | null | undefined;
  } | null;
  doc?: SelectionDocumentApi | null;
  authoring?: {
    setSelectionTarget?: (input: {
      target: {
        kind: 'selection';
        start: TextPoint;
        end: TextPoint;
        story: unknown;
        coordinateSpace?: string;
      };
      focus?: boolean;
    }) => unknown;
  } | null;
}

export interface RtlCaretHost {
  activeEditor?: RtlCaretEditor | null;
}

export interface RtlCaretOptions {
  /** ה-div שהמנוע מצייר בתוכו — `paintedHost(editor.ui)`. */
  host: HTMLElement | null;
  /** ה-SuperDoc של אותו מסמך. */
  superdoc: RtlCaretHost | null | undefined;
}

export interface RtlCaretHandle {
  dispose(): void;
}

/** ההקשה שלנו, אבל אין לאן לזוז — נבלעת. */
export const STAY = 'stay' as const;

/** היעד, והשורה וה-fragment שהוא יושב בהם. */
interface Target {
  fragment: Element;
  line: Element;
  slot: CaretSlot;
}

/** מקום במסמך: הכתובת של המנוע, ה-pm להשוואת סדר, והיכן הוא מצויר. */
interface Spot {
  blockId: string;
  offset: number;
  pm: number;
  x: number;
  y: number;
  rtl: boolean;
  line: Element;
}

/** בחירה שאנחנו מחזיקים — ההסבר ב„בחירה” בראש הקובץ. */
interface HeldSelection {
  anchor: Pick<Spot, 'blockId' | 'offset' | 'pm' | 'line'>;
  head: Spot;
  story: unknown;
  coordinateSpace: string | undefined;
}

/** המקשים שאינם פעולה בפני עצמם: `Shift` שנלחץ שוב אינו סוף הבחירה. */
const MODIFIER_KEYS = new Set(['Shift', 'Control', 'Alt', 'Meta']);

/** פחות מזה בין לחיצה לשחרור — אותה שורה. */
const SAME_LINE_PX = 8;

const middleOf = (el: Element): number => {
  const rect = el.getBoundingClientRect();
  return (rect.top + rect.bottom) / 2;
};

const isTextPoint = (
  point: TextPoint | null | undefined,
): point is TextPoint & { blockId: string; offset: number } =>
  point?.kind === 'text' && typeof point.blockId === 'string' && typeof point.offset === 'number';

/** היכן מצויר היסט. לתפר BiDi אין מיקום יחיד שאפשר לכתוב — אין מנחשים. */
function paintedSpot(
  host: HTMLElement,
  blockId: string,
  offset: number,
  y: number,
  known?: Pick<Spot, 'pm' | 'line'>,
): Spot | null {
  const home = known ? findPaintedLine(host, blockId, known.pm, known.line) : findFragment(host, blockId, offset, y);
  if (!home) return null;
  const lines = linesOf(home.fragment);
  const line = 'line' in home ? home.line as Element : lines[lineAt(lines, home.caretPm, y)];
  if (!line) return null;
  const rtl = line.getAttribute('dir') === 'rtl';
  const chars = readLineChars(line);
  const slots = chars ? caretSlots(chars, rtl) : [emptyLineSlot(line, rtl)];
  const slot = slots.find((candidate) => candidate?.pm === home.caretPm);
  return slot ? { blockId, offset, pm: home.caretPm, x: slot.x, y: middleOf(line), rtl, line } : null;
}

/**
 * השכן החזותי של הסמן: בתוך השורה, ואם אין — בשורה או בפסקה הסמוכה.
 *
 * `null` פירושו „לא שלנו”, וההקשה ממשיכה למנוע. `ownLatin` — בחירה שכבר
 * מוחזקת כאן ממשיכה מאותו ראש גם דרך שורה לטינית.
 */
export function findTarget(
  host: HTMLElement,
  fragment: Element,
  caretPm: number,
  caretX: number,
  caretY: number,
  toRight: boolean,
  ownLatin = false,
  knownLine?: Element,
): Target | typeof STAY | null {
  const lines = linesOf(fragment);
  const index = knownLine ? lines.indexOf(knownLine) : lineAt(lines, caretPm, caretY);
  if (index < 0) return null;

  const line = lines[index]!;
  // שורה לטינית — המנוע נמדד תקין בה (17/17), ואין מה לתקן.
  const lineRtl = line.getAttribute('dir') === 'rtl';
  if (!lineRtl && !ownLatin) return null;

  /*
   * פסקה ריקה אינה „אין מיפוי” אלא שורה בלי חריצים פנימיים: אין בה תו לעבור,
   * והיעד הוא תמיד השכנה. מסירה למנוע כאן אינה ניטרלית — נמדד בשער שההקשה
   * שנמסרה הזיזה **בתוך** הפסקה שממנה באנו, כלומר לכיוון ההפוך (ימין
   * מ-0 הגיע ל-1, ושמאל מ-12 הגיע ל-11).
   */
  const chars = readLineChars(line);
  if (!chars && !isEmptyRange(line)) return null;

  const slots = chars ? caretSlots(chars, lineRtl) : [];
  if (knownLine && chars) {
    const current = slots.find((slot) => slot.pm === caretPm);
    if (!current) return null;
    // הכתובת נשמרת, הקואורדינטה נמדדת מחדש באותה סריקת תווים של ההקשה.
    caretX = current.x;
  }
  const inside = visualTarget(slots, caretX, toRight);
  if (inside) return { fragment, line, slot: inside };

  /* קצה השורה. „ימינה” בשורה עברית הוא אחורה לוגית, ולכן הוא מוביל לשורה
     שמעליה — ואל **סופה**, שהוא קצה שמאל שלה. */
  const forward = movesForward(toRight, lineRtl);
  const neighbour = lines[index + (forward ? 1 : -1)];
  if (neighbour) {
    const nextChars = readLineChars(neighbour);
    if (!nextChars) return null;
    const rtl = neighbour.getAttribute('dir') === 'rtl';
    const slot = edgeSlot(caretSlots(nextChars, rtl), forward, caretPm);
    return slot ? { fragment, line: neighbour, slot } : null;
  }

  /* קצה הפסקה. השכן הוא הבלוק הבא **באותו מכל ציור** — כותרת עליונה או תא
     טבלה אינם שכנים של הסמן; העמוד הבא בגוף כן. */
  const sibling = neighbourOf(host, fragment, forward);
  if (!sibling) return atDocumentEdge(host, fragment, forward) ? STAY : null;
  if (!contiguousWith(fragment, sibling, forward)) return null;

  const siblingId = blockIdOf(sibling);

  /* בלוק בלי מזהה — טבלה: היעד הוא הפסקה הראשונה בה, או האחרונה. */
  if (siblingId === null) {
    const inner = blocksInside(sibling);
    const cell = inner[forward ? 0 : inner.length - 1];
    if (!cell) return null;
    const cellLines = linesOf(cell);
    const cellLine = forward ? cellLines[0] : cellLines[cellLines.length - 1];
    if (!cellLine) return null;
    const cellChars = readLineChars(cellLine);
    if (!cellChars) return null;
    const slot = edgeSlot(caretSlots(cellChars, cellLine.getAttribute('dir') === 'rtl'), forward, caretPm);
    return slot ? { fragment: cell, line: cellLine, slot } : null;
  }

  const siblingLines = linesOf(sibling);
  const edgeLine = forward ? siblingLines[0] : siblingLines[siblingLines.length - 1];
  if (!edgeLine) return null;
  const edgeChars = readLineChars(edgeLine);
  const rtl = edgeLine.getAttribute('dir') === 'rtl';
  // שכנה ריקה: החריץ היחיד שלה הוא ההיסט שלה עצמה.
  const slot = edgeChars
    ? edgeSlot(caretSlots(edgeChars, rtl), forward, caretPm)
    : emptyLineSlot(edgeLine, rtl);
  return slot ? { fragment: sibling, line: edgeLine, slot } : null;
}

/* ------------------------------------------------------------------ */
/* היירוט                                                               */
/* ------------------------------------------------------------------ */

/** מתקין את היירוט. המצב שלו הוא הבחירה המוחזקת, והוא נפרק עם המסמך. */
export function installRtlVisualArrows({ host, superdoc }: RtlCaretOptions): RtlCaretHandle {
  if (!host) return { dispose() {} };

  let held: HeldSelection | null = null;
  let pressedAt: { x: number; y: number; pointerId: number } | null = null;
  /** כל „שכחה” מבטלת קריאה שבדרך: התשובה שלה מתארת בחירה שכבר אינה. */
  let generation = 0;
  let stateEditor = superdoc?.activeEditor;

  const forget = (): void => {
    held = null;
    pressedAt = null;
    generation += 1;
  };

  /* ה-`story` מגיע מהתצלום — נמדד שהוא קיים גם על ה-target וגם על הקצה.
     ברירת מחדל מומצאת הייתה נכתבת למנוע בלי שנמדד אי פעם שהיא הנכונה.

     בלי try/catch ובלי `.catch`: נמדד מה הכתובת הפסולה ביותר עושה — כתיבה
     ל-`blockId` שאינו קיים **אינה זורקת ואינה דוחה**, אלא מחזירה הבטחה
     שנפתרת ב-`{ok:false, reason:'target-not-mounted'}`. זה ערוץ הכשל של
     המנוע, ואין לו מה לעשות כאן: ההקשה כבר נבלעה, והסמן פשוט נשאר. */
  const write = (
    start: TextPoint,
    end: TextPoint,
    story: unknown,
    coordinateSpace: string | undefined,
  ): void => {
    const editor = superdoc?.activeEditor;
    void editor?.authoring?.setSelectionTarget?.call(editor.authoring, {
      target: {
        kind: 'selection',
        start,
        end,
        story,
        ...(typeof coordinateSpace === 'string' ? { coordinateSpace } : {}),
      },
      focus: true,
    });
  };

  /** השכן החזותי של מקום, ככתובת שאפשר לכתוב. `null` כשאינו שלנו. */
  const aim = (
    from: Pick<Spot, 'blockId' | 'offset' | 'x' | 'y'> & Partial<Pick<Spot, 'pm' | 'line'>>,
    toRight: boolean,
    ownLatin: boolean,
  ): Spot | typeof STAY | null => {
    const known = from.line && typeof from.pm === 'number';
    const home = known
      ? findPaintedLine(host, from.blockId, from.pm!, from.line!)
      : findFragment(host, from.blockId, from.offset, from.y);
    if (!home) return null;
    const line = 'line' in home ? home.line as Element : undefined;
    const target = findTarget(host, home.fragment, home.caretPm, from.x,
      line ? middleOf(line) : from.y, toRight, ownLatin, line);
    if (!target || target === STAY) return target;

    // צמצום טיפוס: כל ארבעת המסלולים ב-`findTarget` מחזירים fragment שנבחר
    // **לפי** התכונה הזאת, ולכן היא קיימת; `getAttribute` עדיין מחזיר null.
    const blockId = blockIdOf(target.fragment);
    if (!blockId) return null;

    /* ‏`NaN` כשהחלק הראשון של הפסקה אינו מצויר — היעד הוא המשך של פסקה
       שהתחילה בעמוד שהמנוע לא צייר (נמדד: 3 עמודים מתוך 25), ואז אין מיפוי
       מ-pm להיסט. כמו בכל „אין מיפוי” כאן, ההקשה נמסרת למנוע. */
    const offset = target.slot.pm - blockStart(host, blockId);
    if (!Number.isFinite(offset)) return null;

    return {
      blockId,
      offset,
      pm: target.slot.pm,
      x: target.slot.x,
      y: middleOf(target.line),
      rtl: target.line.getAttribute('dir') === 'rtl',
      line: target.line,
    };
  };

  const pointOf = (spot: Pick<Spot, 'blockId' | 'offset'>, story: unknown): TextPoint => ({
    kind: 'text',
    blockId: spot.blockId,
    offset: spot.offset,
    story,
  });

  /** start הוא העוגן ו-end הוא הראש; רק קריאת selection.current מנרמלת את הסדר. */
  const writeHeld = (selection: HeldSelection): void => {
    const { anchor, head, story, coordinateSpace } = selection;
    const a = pointOf(anchor, story);
    const b = pointOf(head, story);
    write(a, b, story, coordinateSpace);
  };

  /** הקשה על בחירה מוחזקת. `false` — ההקשה אינה שלנו. */
  const stepHeld = (selection: HeldSelection, toRight: boolean, extending: boolean): boolean => {
    const { anchor, head, story, coordinateSpace } = selection;

    if (!extending && anchor.pm !== head.pm) {
      const start = paintedSpot(host, anchor.blockId, anchor.offset, head.y, anchor);
      const end = paintedSpot(host, head.blockId, head.offset, head.y, head);
      if (!start || !end) return false;
      // באותה שורה משווים את שני הקצוות המצוירים, גם בתוך אי לטיני/ספרות.
      // בין שורות הכיוון הלוגי של השורה קובע לאיזה קצה המסמך קורס.
      const edge = start.line === end.line
        ? (toRight === (start.x > end.x) ? start : end)
        : (movesForward(toRight, end.rtl) === (head.pm < anchor.pm) ? start : end);
      const point = pointOf(edge, story);
      write(point, point, story, coordinateSpace);
      held = null;
      return true;
    }

    const next = aim(head, toRight, true);
    if (!next) return false;
    if (next === STAY) return true;
    selection.head = next;
    if (!extending) selection.anchor = next;
    writeHeld(selection);
    return true;
  };

  /** קוראים פעם אחת אחרי גרירה, בלי לעכב או לבלוע אף הקשת מקלדת. */
  const rememberDrag = async (
    from: { x: number; y: number },
    to: { x: number; y: number },
    mine: number,
    owner: RtlCaretEditor | null | undefined,
  ): Promise<void> => {
    // ה-pointerup של המנוע מסיים את הגרירה לפני הקריאה.
    await Promise.resolve();
    if (mine !== generation || owner !== superdoc?.activeEditor) return;
    const { selectionTarget } = await readDocSelection(superdoc);
    if (mine !== generation || owner !== superdoc?.activeEditor) return;

    const range = selectionTarget as NonNullable<LiveSelectionSnapshot['selectionTarget']> | null;
    const start = range?.start;
    const end = range?.end;
    if (!isTextPoint(start) || !isTextPoint(end) || start.blockId !== end.blockId) return;
    const story = range?.story ?? end.story;
    const coordinateSpace = range?.coordinateSpace;

    const a = paintedSpot(host, start.blockId, start.offset, from.y);
    const b = paintedSpot(host, end.blockId, end.offset, to.y);
    // כיוון מגרירה אמין כאן רק כששני הקצוות באותה שורה מצוירת. גם גלילה
    // בזמן הגרירה, ראש מחוץ למסך ותפר דו-כיווני נשארים למנוע, לפני keydown.
    if (!a || !b || !a.rtl || !b.rtl || a.line !== b.line
      || Math.abs(a.y - from.y) >= SAME_LINE_PX || Math.abs(b.y - to.y) >= SAME_LINE_PX
      || Math.abs(a.x - b.x) < 1) return;
    const backward = (to.x - from.x) * (b.x - a.x) < 0;
    held = { anchor: backward ? b : a, head: backward ? a : b, story, coordinateSpace };
  };

  const onKeyDown = (event: KeyboardEvent): void => {
    if (!isHorizontalArrow(event)) return;

    const editor = superdoc?.activeEditor;
    if (editor !== stateEditor) {
      forget();
      stateEditor = editor;
    }
    const setSelectionTarget = editor?.authoring?.setSelectionTarget;
    const readSnapshot = editor?.host?.readLiveSelectionSyncSnapshot;
    if (typeof setSelectionTarget !== 'function' || typeof readSnapshot !== 'function') return;

    const toRight = event.key === 'ArrowRight';
    const extending = event.shiftKey;
    const swallow = (): void => {
      // מכאן ואילך זה שלנו: המנוע לא יראה את ההקשה.
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
    };

    // תשובה שהגיעה אחרי שהמנוע כבר קיבל מקש מתארת בחירה קודמת.
    generation += 1;
    pressedAt = null;

    /* עוגן וראש באותו מקום הם סמן: חץ רגיל אינו „קריסה” אלא תנועה רגילה. */
    if (held && !extending && held.anchor.pm === held.head.pm) held = null;
    if (held) {
      if (stepHeld(held, toRight, extending)) swallow();
      else held = null;
      return;
    }

    /* בלי try/catch: נמדד (17.9.2026, superdoc 2.15.0) שהקריאה אינה זורקת —
       מאות קריאות בסמן מכווץ, בבחירה שאינה מכווצת ואחרי כתיבה. מה שהיא כן
       עושה הוא להחזיר `selectionTarget: null`, וזה מטופל למטה. */
    const snapshot: LiveSelectionSnapshot | null | undefined = readSnapshot.call(editor?.host);

    const selection = snapshot?.selectionTarget;
    if (!selection) {
      // ה-API הציבורי מחזיר טווח מסודר, בלי כיוון. אחרי Shift+Home/End,
      // בחירה בפקודה וכדומה איננו יודעים מי הראש; המנוע שומר אותו בעצמו.
      return;
    }

    const head = selection.end;
    if (!isTextPoint(head)) return;

    const caretRect = host.querySelector(CARET_SELECTOR)?.getBoundingClientRect();
    if (!caretRect) return;
    const caretY = (caretRect.top + caretRect.bottom) / 2;

    let anchor: HeldSelection['anchor'] | null = null;
    if (extending) {
      const home = findFragment(host, head.blockId, head.offset, caretY);
      const lines = home ? linesOf(home.fragment) : [];
      const line = home ? lines[lineAt(lines, home.caretPm, caretY)] : null;
      if (!home || !line) return;
      anchor = { blockId: head.blockId, offset: head.offset, pm: home.caretPm, line };
    }

    const target = aim(
      { ...(anchor ?? { blockId: head.blockId, offset: head.offset }), x: caretRect.left, y: caretY },
      toRight,
      false,
    );
    if (!target) return;

    if (target === STAY) {
      swallow();
      return;
    }

    /*
     * היעד הוא ההיסט שהסמן כבר עליו — ואז בולעים ולא מוסרים.
     *
     * שתי נקודות המוצא של החישוב מתעדכנות בקצב שונה: ההיסט מגיע מהתצלום
     * הסינכרוני, וה-x מהסמן ה**מצויר**. כשהציור מפגר אחרי התצלום (פער מתועד
     * — „ציור שמאחר אחרי לחיצה”) החישוב יוצא מ-x ישן ומחזיר את היעד שכבר
     * נכתב. מסירה למנוע כאן אינה ניטרלית: בפסקה שהוא מזיז בה לוגית הוא זז
     * לכיוון ההפוך — כלומר החזקת חץ ברשימה הייתה מייצרת קפיצות אחורה
     * מזדמנות. אותה הכרעה בדיוק ב-`rtl-line-end.ts` („הסמן כבר שם”), שם
     * ביטול האירוע קודם לבדיקה.
     *
     * ## והחשש ההפוך — שהבליעה תאכל הקשות בהחזקה — נמדד ונשלל
     *
     * הטענה שנבדקה: תחת autorepeat כל הקשה שמקדימה את הציור תיבלע, והמרחק
     * שהסמן עובר ייחסם בקצב הציור ולא בקצב המקשים. נמדד ב-17.9.2026 על
     * ה-dist הארוז, ב-`Input.dispatchKeyEvent` עם `autoRepeat`, בשתי פסקאות
     * — עברית רגילה ופריט רשימה (שבו המנוע זז לוגית):
     *
     *   • **אין חלון.** דגימה בכל פריים אחרי הקשה בודדת: ההיסט וה-x משתנים
     *     ב**אותה** דגימה (‏48@355.3 → 47@363 אחרי 55ms; ברשימה
     *     ‏40@349.7 → 39@356.6 אחרי 52ms). לא נצפתה ולו דגימה אחת שבה ההיסט
     *     כבר חדש וה-x עדיין ישן.
     *   • **אין בליעה.** 25 הקשות מוחזקות ב-19 וב-32 הקשות/שנייה נתנו
     *     ‏**25 קריאות `setSelectionTarget`** בכל אחת מארבע הריצות; הענף הזה
     *     לא ירה אף פעם. במדידה נפרדת 20 הקשות הזיזו את ההיסט ב-20 בדיוק,
     *     גם ב-39 הקשות/שנייה.
     *
     * כלומר הענף נשאר מה שהוא: גידור לפיגור ציור שנמדד בלחיצת עכבר, ולא
     * חסם על מהירות התנועה.
     */
    if (target.blockId === head.blockId && target.offset === head.offset) {
      swallow();
      return;
    }

    const story = selection.story ?? head.story;
    const coordinateSpace = selection.coordinateSpace;

    if (anchor) {
      swallow();
      held = {
        anchor,
        head: target,
        story,
        coordinateSpace,
      };
      writeHeld(held);
      return;
    }

    swallow();
    const point = pointOf(target, story);
    write(point, point, story, coordinateSpace);
  };

  /* על ה-`window` מאותה סיבה בדיוק כמו ב-`rtl-vertical-caret.ts`: כדי לא להיות
     תלויים בסדר ההתקנה, ומפני שגם מקש בדיאלוג ולחיצה ברצועה מזיזים בחירה. */
  const onAnyKeyDown = (event: KeyboardEvent): void => {
    if (!isHorizontalArrow(event) && !MODIFIER_KEYS.has(event.key)) forget();
  };
  const onPointerDown = (event: PointerEvent): void => {
    forget();
    if (!host.contains(event.target as Node) || event.button !== 0 || event.isPrimary === false
      || event.shiftKey || event.ctrlKey || event.altKey || event.metaKey) return;
    pressedAt = { x: event.clientX, y: event.clientY, pointerId: event.pointerId };
  };
  const onPointerUp = (event: PointerEvent): void => {
    const from = pressedAt;
    pressedAt = null;
    if (!from || from.pointerId !== event.pointerId || event.shiftKey || event.ctrlKey
      || event.altKey || event.metaKey || Math.abs(event.clientX - from.x) < 1
      || Math.abs(event.clientY - from.y) >= SAME_LINE_PX) return;
    const to = { x: event.clientX, y: event.clientY };
    if (lineDirectionUnder(host, { left: to.x, right: to.x, top: to.y, bottom: to.y }) !== true) return;
    void rememberDrag(from, to, generation, superdoc?.activeEditor);
  };

  host.addEventListener('keydown', onKeyDown, true);
  window.addEventListener('keydown', onAnyKeyDown, true);
  window.addEventListener('pointerdown', onPointerDown, true);
  window.addEventListener('pointerup', onPointerUp, true);
  window.addEventListener('pointercancel', forget, true);

  return {
    dispose() {
      forget();
      host.removeEventListener('keydown', onKeyDown, true);
      window.removeEventListener('keydown', onAnyKeyDown, true);
      window.removeEventListener('pointerdown', onPointerDown, true);
      window.removeEventListener('pointerup', onPointerUp, true);
      window.removeEventListener('pointercancel', forget, true);
    },
  };
}
