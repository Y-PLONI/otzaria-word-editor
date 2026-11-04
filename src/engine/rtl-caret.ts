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
 * טווח, והמנוע מנרמל `start`/`end` לסדר המסמך. שתי הידיות שפותרות את זה:
 *
 *   • **בחירה שאנחנו בונים — זוכרים.** העוגן והראש מוחזקים ב-`HeldSelection`,
 *     וההקשה הבאה ממשיכה מהראש שנזכר בלי לשאול את המנוע. „נשכח” בכל מקש אחר
 *     ובכל לחיצת עכבר, בכל מקום בדף — אלה המסלולים שמשנים בחירה.
 *   • **בחירה שלא בנינו** (עכבר, `Shift+End`, `Ctrl+A`). `getSelectionAnchorRect()`
 *     סינכרוני ומחזיר את מלבן הבחירה גם כשהתצלום `null`, ולכן ההכרעה אם ההקשה
 *     שלנו נופלת בתוך ה-keydown; את הקצוות קוראים פעם אחת מ-`doc.selection
 *     .current()`, והקשות שמגיעות בינתיים ממתינות בתור. הראש הוא סוף הבחירה,
 *     אלא אם העכבר נגרר אחורה (`dragWentBackward`).
 *     תמונה אינה מגיעה לכאן: לחיצה עליה נותנת סמן מכווץ ולא בחירה (נמדד).
 *
 * מה שנשאר למנוע: `Shift+חץ` בשורה לטינית כשהבחירה לא התחילה אצלנו, וכל מה
 * שלמעלה כ„נשאר למנוע”. שני גבולות שנשארו: אחרי `Shift+Home` או
 * ‏`Ctrl+Shift+חץ` הראש נחשב סוף הבחירה גם כשהוא בתחילתה, כי אלה אינם משאירים
 * עקבות; וקריסה בחץ רגיל הולכת לפי הכיוון הלוגי של השורה, ולכן קצה שיושב בתוך
 * אי ספרות אינו הקצה המצויר.
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
  isEmptyRange,
  lineAt,
  lineDirectionUnder,
  linesOf,
  neighbourOf,
  readLineChars,
  type CaretSlot,
  type SelectionRect,
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
    getSelectionAnchorRect?: () => SelectionRect | null | undefined;
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
}

/** בחירה שאנחנו מחזיקים — ההסבר ב„בחירה” בראש הקובץ. */
interface HeldSelection {
  anchor: Pick<Spot, 'blockId' | 'offset' | 'pm'>;
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

/**
 * האם גרירת העכבר הלכה נגד סדר המסמך, ולכן ראש הבחירה בתחילתה.
 *
 * המנוע מנרמל את הקצוות, ולכן זה הדבר היחיד שנשאר מהכיוון. שורה אחרת מוכרעת
 * לפי y, ובאותה שורה אחורה בעברית הוא ימינה.
 */
export function dragWentBackward(dx: number, dy: number, lineRtl: boolean): boolean {
  if (Math.abs(dy) > SAME_LINE_PX) return dy < 0;
  return lineRtl ? dx > 0 : dx < 0;
}

/** היכן מצויר היסט. בתפר אין לו חריץ, והקרוב ביותר בסדר המסמך מייצג אותו. */
function paintedSpot(host: HTMLElement, blockId: string, offset: number, y: number): Spot | null {
  const home = findFragment(host, blockId, offset, y);
  if (!home) return null;
  const lines = linesOf(home.fragment);
  const line = lines[lineAt(lines, home.caretPm, y)];
  if (!line) return null;
  const rtl = line.getAttribute('dir') === 'rtl';
  const chars = readLineChars(line);
  const slots = chars ? caretSlots(chars, rtl) : [emptyLineSlot(line, rtl)];
  let nearest: CaretSlot | null = null;
  for (const slot of slots) {
    if (!slot) continue;
    if (!nearest || Math.abs(slot.pm - home.caretPm) < Math.abs(nearest.pm - home.caretPm)) nearest = slot;
  }
  return nearest ? { blockId, offset, pm: home.caretPm, x: nearest.x, y: middleOf(line), rtl } : null;
}

/**
 * השכן החזותי של הסמן: בתוך השורה, ואם אין — בשורה או בפסקה הסמוכה.
 *
 * `null` פירושו „לא שלנו”, וההקשה ממשיכה למנוע. `ownLatin` — בחירה שכבר
 * מוחזקת כאן ממשיכה דרך שורה לטינית: המנוע אינו יודע איזה קצה הוא הראש.
 */
export function findTarget(
  host: HTMLElement,
  fragment: Element,
  caretPm: number,
  caretX: number,
  caretY: number,
  toRight: boolean,
  ownLatin = false,
): Target | typeof STAY | null {
  const lines = linesOf(fragment);
  const index = lineAt(lines, caretPm, caretY);
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

  const inside = chars ? visualTarget(caretSlots(chars, lineRtl), caretX, toRight) : null;
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
  let pressedAt: { x: number; y: number } | null = null;
  let drag: { dx: number; dy: number } | null = null;
  /** קריאת הקצוות שבדרך, והקשות שממתינות לה. */
  let reading: Promise<void> | null = null;
  /** כל „שכחה” מבטלת קריאה שבדרך: התשובה שלה מתארת בחירה שכבר אינה. */
  let generation = 0;

  const forget = (): void => {
    held = null;
    drag = null;
    reading = null;
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
    from: Pick<Spot, 'blockId' | 'offset' | 'x' | 'y'>,
    toRight: boolean,
    ownLatin: boolean,
  ): Spot | typeof STAY | null => {
    const home = findFragment(host, from.blockId, from.offset, from.y);
    if (!home) return null;

    const target = findTarget(host, home.fragment, home.caretPm, from.x, from.y, toRight, ownLatin);
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
    };
  };

  const pointOf = (spot: Pick<Spot, 'blockId' | 'offset'>, story: unknown): TextPoint => ({
    kind: 'text',
    blockId: spot.blockId,
    offset: spot.offset,
    story,
  });

  /** כותב בחירה בסדר המסמך: המנוע מנרמל ממילא, ואין סיבה לסמוך על זה בין פסקאות. */
  const writeHeld = (selection: HeldSelection): void => {
    const { anchor, head, story, coordinateSpace } = selection;
    const a = pointOf(anchor, story);
    const b = pointOf(head, story);
    if (head.pm < anchor.pm) write(b, a, story, coordinateSpace);
    else write(a, b, story, coordinateSpace);
  };

  /** הקשה על בחירה מוחזקת. `false` — ההקשה אינה שלנו. */
  const stepHeld = (selection: HeldSelection, toRight: boolean, extending: boolean): boolean => {
    const { anchor, head, story, coordinateSpace } = selection;

    if (!extending) {
      const toEnd = movesForward(toRight, head.rtl);
      const point = pointOf(toEnd === head.pm < anchor.pm ? anchor : head, story);
      write(point, point, story, coordinateSpace);
      held = null;
      return true;
    }

    const next = aim(head, toRight, true);
    if (!next) return false;
    if (next === STAY) return true;
    selection.head = next;
    writeHeld(selection);
    return true;
  };

  /** הבחירה לא נבנתה כאן: קוראים את הקצוות, ומשם זו אותה הקשה. */
  const readThenStep = async (
    rect: SelectionRect,
    toRight: boolean,
    extending: boolean,
  ): Promise<void> => {
    const mine = generation;
    const gesture = drag;
    const { selectionTarget } = await readDocSelection(superdoc);
    if (mine !== generation) return;

    const range = selectionTarget as NonNullable<LiveSelectionSnapshot['selectionTarget']> | null;
    const start = range?.start;
    const end = range?.end;
    if (!start || !end) return;
    const story = range?.story ?? end.story;
    const coordinateSpace = range?.coordinateSpace;

    if (!extending) {
      const point = movesForward(toRight, true) ? end : start;
      write(point, point, story, coordinateSpace);
      return;
    }
    if (!isTextPoint(start) || !isTextPoint(end)) return;

    const backward = gesture ? dragWentBackward(gesture.dx, gesture.dy, true) : false;
    const [anchor, tip] = backward ? [end, start] : [start, end];
    const head = paintedSpot(host, tip.blockId, tip.offset, backward ? rect.top : rect.bottom);
    if (!head) return;
    /* עוגן שאינו מצויר הוא רחוק, ואז מספיק לדעת באיזה צד הוא. */
    const base = blockStart(host, anchor.blockId);
    const pm = Number.isFinite(base) ? base + anchor.offset : backward ? Infinity : -Infinity;
    held = {
      anchor: { blockId: anchor.blockId, offset: anchor.offset, pm },
      head,
      story,
      coordinateSpace,
    };
    stepHeld(held, toRight, true);
  };

  const onKeyDown = (event: KeyboardEvent): void => {
    if (!isHorizontalArrow(event)) return;

    const editor = superdoc?.activeEditor;
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

    if (reading) {
      swallow();
      const mine = generation;
      reading = reading.then(() => {
        if (mine === generation && held) stepHeld(held, toRight, extending);
      });
      return;
    }

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
      const rect = editor?.host?.getSelectionAnchorRect?.call(editor.host);
      if (!rect || lineDirectionUnder(host, rect) !== true) return;
      swallow();
      const mine = generation;
      reading = readThenStep(rect, toRight, extending).finally(() => {
        if (mine === generation) reading = null;
      });
      return;
    }

    const head = selection.end;
    if (!isTextPoint(head)) return;

    const caretRect = host.querySelector(CARET_SELECTOR)?.getBoundingClientRect();
    if (!caretRect) return;
    const caretY = (caretRect.top + caretRect.bottom) / 2;

    const target = aim(
      { blockId: head.blockId, offset: head.offset, x: caretRect.left, y: caretY },
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

    swallow();

    const story = selection.story ?? head.story;
    const coordinateSpace = selection.coordinateSpace;

    if (extending) {
      const pm = blockStart(host, head.blockId) + head.offset;
      held = {
        anchor: { blockId: head.blockId, offset: head.offset, pm },
        head: target,
        story,
        coordinateSpace,
      };
      writeHeld(held);
      return;
    }

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
    pressedAt = { x: event.clientX, y: event.clientY };
  };
  const onPointerUp = (event: PointerEvent): void => {
    if (pressedAt) drag = { dx: event.clientX - pressedAt.x, dy: event.clientY - pressedAt.y };
    pressedAt = null;
  };

  host.addEventListener('keydown', onKeyDown, true);
  window.addEventListener('keydown', onAnyKeyDown, true);
  window.addEventListener('pointerdown', onPointerDown, true);
  window.addEventListener('pointerup', onPointerUp, true);

  return {
    dispose() {
      host.removeEventListener('keydown', onKeyDown, true);
      window.removeEventListener('keydown', onAnyKeyDown, true);
      window.removeEventListener('pointerdown', onPointerDown, true);
      window.removeEventListener('pointerup', onPointerUp, true);
    },
  };
}
