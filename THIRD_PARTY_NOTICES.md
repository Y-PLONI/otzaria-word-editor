# הודעות רישוי של רכיבי צד שלישי

התוסף עצמו מופץ תחת [AGPL-3.0](../LICENSE). המסמך הזה מפרט את הרכיבים
שנארזים לתוך ה־`.otzplugin` ואת חובות הרישוי שלהם. העתקים של נוסחי הרישיון
נארזים תחת `third-party/` בתוך החבילה עצמה, לא רק במאגר.

## superdoc 2.12.0 — AGPL-3.0

- מקור: <https://github.com/superdoc/docx-editor>
- רישיון: AGPL-3.0 (`third-party/SUPERDOC-LICENSE.txt`)
- הודעה: `third-party/SUPERDOC-NOTICE.txt`

זו הסיבה שהתוסף כולו AGPL-3.0, שהמקור מפורסם, ושהמקור המפורסם זהה לבינארי
המופץ.

## @superdoc/docx-engine 0.11.0 — קנייני

מנוע ה־DOCX אינו קוד פתוח. הוא נמשך כתלות של `superdoc` ונארז לתוך החבילה
(כולל קוד ה־Workers שמוטמע ב־`assets/engine-workers.js`).

- רישיון: DOCX Engine Proprietary License Agreement
  (`third-party/DOCX-ENGINE-LICENSE.md`, גרסה 2026-07-14)
- הודעה: `third-party/DOCX-ENGINE-NOTICE.md`
- Copyright © 2026 Harbour Enterprises, Inc., d/b/a SuperDoc

סעיף 3.1(d) ברישיון אוסר redistribution. מפתחי SuperDoc אישרו במפורש
ב־[issue #3927](https://github.com/superdoc/docx-editor/issues/3927#issuecomment-5383145303)
שתוסף קוד פתוח תחת AGPLv3 רשאי לארוז ולהפיץ את המנוע, ה־Workers ונכסי ה־runtime
שלו בתוך חבילה אופליין, ללא רישיון מסחרי; האיסור מכוון להפצת המנוע כחבילה
עצמאית.

חובות מעשיות שהקוד מחויב להן:

- מייבאים `superdoc` בלבד. אין import ישיר ל־`@superdoc/docx-engine` ואין
  שימוש בנתיב פנימי שאינו export ציבורי של החבילה.
- אין לשנות, לפרק, לעשות deobfuscate או reverse engineering למנוע — כולל
  בעזרת כלי AI. אין לקרוא את קוד המנוע כדי להסיק ממנו מימוש.
- סעיף 3.1(c): אין להסיר או להסתיר הודעות רישוי, banners או markers. באנר
  הרישוי של המנוע חייב לשרוד את הבנייה; `npm run check:dist` מאמת זאת.
- אין להשתמש במנוע כדי לפתח מוצר מתחרה או מימוש חלופי.
- אין להעלות את חבילת המנוע למערכות AI של צד שלישי.

## Selawik 1.01 — SIL OFL 1.1

נארז תחת `fonts/` (3 קבצים, 129KB) ומוצהר כ־`@font-face` ב־`src/styles/fonts.ts`.

- מקור: <https://github.com/microsoft/Selawik> (release 1.01)
- רישיון: SIL Open Font License 1.1 (`third-party/SELAWIK-LICENSE.txt`)
- Copyright © 2015 Microsoft Corporation, with Reserved Font Name **Selawik**
- `fsType = 0` — Installable Embedding, בלי הגבלת הטמעה או הפצה

למה הוא נארז: מסמכי DOCX שנכתבו ב־Word קוראים לגופנים של Word, ו־`Segoe UI` אינו
קיים ב־macOS ובלינוקס. Selawik הוא הגופן ש־Microsoft שחררה בעצמה כתחליף
**מטרית־תואם** ל־Segoe UI, בדיוק בשביל השימוש הזה. מדריך העיצוב של ה־SDK אומר
שאין צורך לארוז גופנים, אבל מה שאוצריא מזריקה הוא גופן הקריאה שנבחר בהגדרות
בלבד, לא גופני מסמכים.

חובות מעשיות שהקוד מחויב להן:

- **סעיף 2 ב־OFL:** נוסח הרישיון מופץ עם הגופן. `third-party/SELAWIK-LICENSE.txt`
  נארז לתוך החבילה, ו־`npm run check:dist` מאמת שהוא שם.
- **Reserved Font Name:** אין לשנות את קובצי הגופן ולהמשיך לקרוא להם „Selawik”.
  הקבצים נארזים כפי שהם, בלי subsetting ובלי המרה.
- הגופן מוצהר בשני שמות: `Selawik` (שמו) ו־`Segoe UI` (שם התאמה, כדי שמסמך
  יקבל את המטריקות הנכונות). „Segoe UI” הוא סימן מסחרי של Microsoft ומופיע
  כשם התאמה בלבד — אותה החלפה שעושים fontconfig ו־LibreOffice. הגופן עצמו
  אינו מוצג בשום מקום כ־Segoe UI כלפי המשתמש.

מה שנמדד בקבצים ומגדיר את הגבול של הפתרון:

- **אין עברית.** 348 מתווים, אפס בבלוק העברי. Selawik פותר את הטקסט הלטיני ואת
  המטריקות; טקסט עברי — כלומר כמעט כל מה שייכתב בתוסף הזה — נופל ל־`David`
  ולגופן המערכת, כמו לפני האריזה.
- **אין פנים נטויה** בריליס. הדפדפן מטה את הרגילה סינתטית.

> גרסה קודמת של התוסף ארזה את **Segoe UI** עצמו (3.3MB, © Microsoft,
> `fsType = 8`). לא היה לזה היתר הפצה, והוא הוחלף. הקבצים ההם נוקו גם
> מהיסטוריית ה־git.

## Fluent System Icons 1.1.338 — MIT

אייקונים ב־`src/ui/icons/icons.ts` שהם ה־path data המקורי של Microsoft, ולא
ציור בבית. כל הווריאנטים הם `*_20_regular`.

**לשונית „קובץ”**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `newDoc` | מסמך חדש | `document_add` |
| `folder` | פתח קובץ | `folder_open` |
| `save` | שמור (וגם סרגל הגישה המהירה) | `save` |
| `saveAs` | שמור בשם... | `save_edit` |
| `export` | ייצוא ל־Word | `arrow_export_rtl` |
| `exportPdf` | ייצוא ל־PDF | `document_pdf` |
| `print` | הדפסה | `print` |

**„פתח מסמך” — העץ והרשימה**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `history` | אחרונים | `history` |
| `library` | ספריית אוצריא | `library` |
| `folderClosed` | מדף, תיקייה | `folder` |
| `folderAdd` | הוסף תיקייה | `folder_add` |
| `docFile` | ספר או קובץ ברשימה | `document_text` |
| `info` | אודות | `info` |

**לשונית „בית” — לוח**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `paste` | הדבק | `clipboard_paste` |
| `pasteMerge` | הדבקה עם מיזוג עיצוב | `clipboard_text_edit` |
| `pasteText` | הדבקת טקסט בלבד | `clipboard_text_rtl` |
| `cut` | גזור | `cut` |
| `copy` | העתק | `copy` |
| `formatPainter` | מברשת עיצוב | `paint_brush` |

**לשונית „בית” — גופן**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `bold` | מודגש | `text_bold` |
| `italic` | נטוי | `text_italic` |
| `underline` | קו תחתון | `text_underline` |
| `strikethrough` | קו חוצה | `text_strikethrough` |
| `subscript` | כתב תחתי | `text_subscript` |
| `superscript` | כתב עילי | `text_superscript` |
| `highlight` | הדגשת טקסט | `highlight` |
| `clearFormatting` | נקה עיצוב | `text_clear_formatting` |
| `growFont` | הגדל גופן | `font_increase` |
| `shrinkFont` | הקטן גופן | `font_decrease` |

**לשונית „בית” — פיסקה**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `alignRight` | יישור לימין | `text_align_right` |
| `alignCenter` | מרכז | `text_align_center` |
| `alignLeft` | יישור לשמאל | `text_align_left` |
| `alignJustify` | יישור לשני הצדדים | `text_align_justify` |
| `bulletList` | רשימת תבליטים | `text_bullet_list_rtl` |
| `numberList` | רשימה ממוספרת | `text_number_list_rtl` |
| `indentIncrease` | הגדל כניסה | `text_indent_increase_rtl` |
| `indentDecrease` | הקטן כניסה | `text_indent_decrease_rtl` |
| `dirRtl` | כיוון פסקה מימין לשמאל | `text_paragraph_direction_left` |
| `dirLtr` | כיוון פסקה משמאל לימין | `text_paragraph_direction_right` |
| `pilcrow` | הצג/הסתר סימני עיצוב | `text_paragraph` |
| `paragraphOptions` | תפריט פסקה | `options` |
| `lineSpacing` | מרווח שורות | `text_line_spacing` |
| `lineSpacingIncrease` | הגדל מרווח שורות | `text_expand` |
| `lineSpacingDecrease` | הקטן מרווח שורות | `text_collapse` |
| `borders` | גבולות | `border_all` |
| `shading` | צביעה | `paint_bucket` |

**לשונית „בית” — עריכה**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `search` | חיפוש (וגם סרגל הגישה המהירה) | `search` |
| `mic` | הקלטה (הכפתור הצף ולשונית אוצריא) | `mic` |
| `replace` | החלפה | `arrow_swap` |
| `select` | בחר הכל | `select_all_on` |

**לשונית „הוספה”**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `table` | טבלה | `table` |
| `image` | תמונה | `image` |
| `link` | קישור | `link` |
| `pageBreak` | מעבר עמוד | `document_page_break` |
| `bookmark` | סימנייה (וגם „הפניות”) | `bookmark` |
| `header` | כותרת עליונה | `document_header` |
| `footer` | כותרת תחתונה | `document_footer` |
| `pageNumber` | מספר עמוד (וגם „פריסה”) | `document_page_number` |
| `dateTime` | תאריך ושעה | `calendar_clock` |
| `updateFields` | עדכן שדות, וכל „עדכן…” ב„הפניות” | `arrow_clockwise` |

**לשונית „פריסה”**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `margins` | שוליים | `document_margins` |
| `orientation` | כיוון | `document_landscape` |
| `paperSize` | גודל | `slide_size` |
| `columns` | עמודות | `text_column_two` |

**לשונית „הפניות”**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `footnote` | הערת שוליים | `text_footnote` |

**לשונית „סקירה”**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `trackChanges` | עקוב אחר שינויים | `document_edit` |
| `accept` | אשר שינוי | `checkmark_circle` |
| `reject` | דחה שינוי | `dismiss_circle` |
| `comment` | הערה | `comment` |
| `proofing` | בדיקת איות | `text_proofing_tools` |

**לשונית „תצוגה”**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `ruler` | סרגל | `ruler` |
| `zoom` | זום | `zoom_in` |
| `fitWidth` | התאם לרוחב | `arrow_autofit_width` |
| `focusMode` | מצב מיקוד | `full_screen_maximize` |
| `navigationPane` | חלונית ניווט | `text_bullet_list_tree` |

**לשונית „אוצריא” והמעטפת**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `book` | ציטוט מהקורא | `book` |
| `undo` | בטל | `arrow_undo` |
| `redo` | בצע שוב | `arrow_redo` |
| `chevronDown` | פתיחת תפריט וכיווץ הרצועה | `chevron_down` |
| `chevronUp` | כיווץ הרצועה | `chevron_up` |
| `chevronLeft` | גלילת גלריית הסגנונות | `chevron_left` |
| `chevronRight` | גלילת גלריית הסגנונות | `chevron_right` |
| `keyboard` | ניהול קיצורים | `keyboard` |
| `bug` | דיווח על באג | `bug` |

- מקור: <https://github.com/microsoft/fluentui-system-icons>
- חבילה: `@fluentui/svg-icons@1.1.338` ב־npm, וריאנט `*_20_regular`
- רישיון: MIT, Copyright © 2020 Microsoft Corporation
- הודעה: באנר `@license MIT` בראש `src/ui/icons/icons.ts`

חובות מעשיות שהקוד מחויב להן:

- **סעיף היחיד ב־MIT:** נוסח הרישיון והקרדיט מופצים עם כל עותק. הם אינם קובץ
  נפרד תחת `third-party/` אלא באנר legal comment בראש `icons.ts`, שנאסף
  לסוף `assets/app.js` דרך `esbuild.legalComments: 'eof'`. `npm run check:dist`
  מאמת שהבאנר שרד את המינימיזציה, בדיוק כמו באנר מנוע ה־DOCX.
- **אין תלות חדשה.** `@fluentui/svg-icons` אינו ב־`package.json` ואינו נארז.
  הועתק ה־path data בלבד; מעטפת ה־`<svg>`, השמות ב־`ICONS` והמנגנון סביבם הם
  קוד של התוסף. `npm run check:icons` מושך את החבילה ל־`npm pack` בתיקייה זמנית
  לצורך ההשוואה בלבד, מוחק אותה בסוף, ואינו נוגע ב־`package.json`, ב־`node_modules`
  או ב־`dist`.
- **גרסה נעוצה.** ההשוואה היא מול `1.1.338` ולא מול „האחרונה”. גרסה חדשה יותר
  עשויה לצייר מחדש גליף, ואז ה־path data שכאן היה נפסל כ„סטייה” בלי שהשתנה כלום.
  שדרוג הוא החלטה: מעדכנים את הגרסה בשער, מריצים אותו, ומעדכנים כאן.
- **אין אייקוני מיתוג.** ה־MIT מכסה את האייקונים, אך סמלי לוגו ומוצר של
  Microsoft הם סימני מסחר. נלקחו אייקוני ממשק גנריים בלבד — דף, תיקייה,
  דיסקט, מדפסת, חץ, עיגול מידע. `word` ו־`otzaria` נשארו ציור בבית.

למה דווקא Fluent System Icons ולא Fluent MDL2, שדומה יותר ל־Ribbon של Word:
MDL2 היא הספרייה של Office UI Fabric שיצאה משימוש, System Icons היא הסט
הפעיל של Fluent 2 תחת MIT, יש לה גריד 20 שמתאים בדיוק ל־`viewBox` של הסט
הקיים, ויש לה וריאנטי RTL מוצהרים — `arrow_export_rtl` מול `arrow_export_ltr`
— שזה בדיוק מה שנדרש בממשק עברי.

תשע בחירות שאין להן מקבילה מדויקת אצל Microsoft, ולכן הן מתועדות כאן:

- `replace` הוא `arrow_swap` — אין ב־System Icons אייקון „מצא והחלף”.
- `formatPainter` הוא `paint_brush` ולא `clipboard_brush`, כי זה מה ש־Word
  מציג, ולוח כבר מופיע בשלושת השכנים שלו באותה קבוצה.
- `orientation` הוא `document_landscape` ולא `orientation` — האייקון שנקרא כך
  אצל Microsoft הוא סיבוב **מכשיר**, לא כיוון דף.
- `paperSize` הוא `slide_size` — אין „גודל נייר”, וזה הגליף שאומר גודל בלי
  להיות דף נוסף ליד `margins`.
- `accept`/`reject` הם `checkmark_circle`/`dismiss_circle` ולא הגרסאות
  העירומות: `dismiss` יוצא 60% ונופל מהמינימום של 70%, ואז הוא נראה קטן
  מה־`checkmark` שלצידו. שני העיגולים 80%×80%.
- `dirRtl` הוא `text_paragraph_direction_left`. בשתי הגרסאות סימן הפיסקה יושב
  בימין והחץ בשמאל; מה שמבדיל הוא כיוון החץ, כלומר כיוון הזרימה.
- `exportPdf` הוא `document_pdf` ולא `document_arrow_down`. השני חולק עם
  `document_add` — שכנו „מסמך חדש” באותה לשונית — את **אותו** קו מתאר של דף,
  ונבדל ממנו רק בגליף שבתג; ב־20px זה קורא כאותו אייקון פעמיים. ב־`document_pdf`
  הפורמט כתוב באותיות ואין מה לפענח.
- `dateTime` הוא `calendar_clock` ולא `calendar_ltr`/`calendar_rtl`. הפקד מכניס
  תאריך **ושעה**, והשעון הוא ההבדל שנקרא; מה שמבדיל בין וריאנטי ה־LTR ל־RTL הוא
  סדר הימים, שאינו נראה בגריד 20.
- `updateFields` הוא `arrow_clockwise` ולא `document_sync`. השם הזה משרת גם את
  „עדכן טבלה”, „עדכן הפניות”, „עדכן מפתח” ו„עדכן ביבליוגרפיה” בלשונית „הפניות”,
  ואף אחד מהם אינו עדכון של דף — הסימן צריך להיות רענון גנרי. `arrow_sync`,
  המועמד השני, יוצא 93.5% לגובה ונחתך.

> **84 מתוך 90 האייקונים** ב־`icons.ts` הם Fluent System Icons, וזו אינה
> הצהרת כוונות: `npm run check:icons` מוריד את `@fluentui/svg-icons@1.1.338`
> מ־npm ומשווה כל `d=` בקובץ מול הגליף המקורי — byte-for-byte — וגם מאמת שהטבלה
> למעלה נוקבת בשם הגליף הנכון. השער נכשל על סטייה, על אייקון מצויר בבית שאינו
> ברשימת החריגים כאן, ועל חריג שכבר יש לו מקבילה. הוא **מדלג** כשאין רשת, כמו
> `check:sdk`.
>
> **שניים מהשבעים ושישה אינם העתק אלא נגזרת, ולא מאותו סוג: `toc` סודר
> מחדש, ו־`fontColor` נלקח ממנו חלק בלבד.**
>
> **`toc`** (תוכן עניינים, וגם
> „הפניות”), מ־`document_bullet_list`. אצל Microsoft התבליטים יושבים משמאל
> לשורות וקיפול הפינה בימין־למעלה — דף שנקרא משמאל לימין — ובממשק עברי זה
> הפוך. ב־1.1.338 יש 19 גליפי `_rtl` בגריד 20, ואף אחד מהם אינו `document_*`;
> הקרוב ביותר, `text_bullet_list_rtl`, הוא **שיקוף אופקי מלא** של
> `text_bullet_list`, וזה מה שנעשה כאן: הגליף כולו משוקף סביב `x=10` — דף,
> קיפול ותבליטים כאחד.
>
> ה־path של Microsoft **אינו משתנה בתו אחד**: הוא כולו, כמות שהוא, בתוך
> `<g transform>`. השיקוף הוא קוד של התוסף; הקווים הם של Microsoft.
> `check:icons` מאמת בכל הרצה את שתי הטענות — שה־`d` זהה למקור byte־for־byte,
> **ושה־`transform` הוא בדיוק המחרוזת שנרשמה**: בלי transform זהו העתק ומקומו
> בטבלה ולא ברשימה שמרפה את ההשוואה, ועם transform אחר זו נגזרת אחרת מזו
> שההצהרה מתארת. ה־MIT מתיר שינוי ממילא; ההפרדה כאן היא כדי שההצהרה תישאר
> מדויקת לגבי **מה** נלקח.
>
> **`fontColor`** (צבע גופן) הוא `text_color`, ומה שנלקח ממנו הוא **האות
> בלבד**. בגליף המקורי יושב מתחת לאות מלבן חלול — המקום שבו Word מראה את
> הצבע — אבל הפקד כאן כבר מצייר פס כזה בעצמו, מתחת לאייקון, בצבע שהלחיצה
> תחיל (`ColorPickerPopover.vue`). התוצאה הייתה שני מלבנים זה מעל זה,
> שהעליון מהם ריק תמיד. המלבן החלול הושמט, והאות הוגדלה למרכז הגריד כדי
> שלא תישאר קטנה בפינה — `translate(-7.5 -1) scale(1.75)`.
>
> גם כאן **אף קו לא צויר מחדש**: מה שנשאר הוא מחרוזת מתוך ה־`d` של
> Microsoft, מילה במילה. `check:icons` מאמת זאת בטענה המדויקת —
> **שרשור מה שנשאר ומה שהושמט (`omitted` ב־DERIVED) מחזיר את הגליף
> byte־for־byte** — ולכן ההצהרה אינה נשענת על עין אלא על מדידה. השמטה,
> כמו שיקוף, מותרת ב־MIT ממילא; מה שהיא משנה הוא **מה** נלקח, וזה מה
> שנרשם כאן.
>
> ששת היוצאים נשארים ציור בבית, ולא מאותה סיבה — וההבחנה חשובה כאן, כי הצהרת
> רישוי שמונה אייקון של הפרויקט בין המכוסים ב־MIT מרחיבה את מה שמוצהר מעבר למה
> שנלקח:
>
> - `word` — תג האפליקציה בשורת הכותרת וב„אודות”. ה־MIT של Microsoft מכסה
>   אייקוני ממשק ולא סמלי מוצר. ממולא ולא קווי, כמקובל בתגי אפליקציה (כך גם
>   ב־Word עצמו), והוא אינו יושב בשורה אחת עם אייקוני פקודה.
> - `otzaria` — הלוגו של אוצריא, ספר פתוח, על הכפתור „פתח ספרייה”. הוא אינו של
>   Microsoft כלל, ולכן אין לו מקבילה בספרייה.
> - `exit` — „יציאה”, ואייקון פקודה לכל דבר. הוא נשאר ציור בבית מסיבה אחרת
>   לגמרי: הוא נדרש עם חץ **שמאלה**, כי הממשק עברי, ובספרייה אין אייקון יציאה
>   בכיוון הזה. `arrow_exit` ו־`sign_out` שניהם חץ ימינה, ולאף אחד מהם אין
>   וריאנט `_rtl` כמו ש־`arrow_export` דווקא כן מציע. `door_arrow_left` הקיים
>   אינו הווריאנט המשוקף אלא ההפך במשמעות: אצל Microsoft הדלת יושבת בשמאל בשתי
>   הגרסאות, ולכן חץ שמאלה מצביע **פנימה** אליה — כניסה. להציג „כניסה” על כפתור
>   יציאה גרוע מציור בבית.
> - `macro` — סמל ההקלטה המוסכם, טבעת עם דיסק במרכזה. ב־1.1.338 אין גליף
>   `record`, ומה שקרוב אליו (`circle`, `presence_available`) הוא עיגול יחיד ולא
>   טבעת עם דיסק. הוא הורכב בבית מטבעת בכיווני winding הפוכים, כמו זו של `info`.
> - `firstPageHeader` ו־`oddEvenPages` — „שונה בעמוד ראשון” ו„שונה בעמודים
>   זוגיים ואי־זוגיים”. שניהם „שני דפים שנבדלים זה מזה”, ולזה אין גליף:
>   `document_header`, `document_footer` ו־`document_header_footer` כולם על דף
>   **יחיד**, ואין ב־1.1.338 שום גליף שמעמיד שני דפים זה מול זה — לא תחת
>   `parity`, לא `odd`, לא `even`, ולא `mirror` שאינו של מסך כפול. מה שאין לו
>   מקבילה הוא ההבחנה עצמה, ולא הכותרת; „שונה” היא טענה על הפרש בין שניים ואי
>   אפשר לצייר אותה על עותק יחיד. הם מצוירים בבית בגריד ובמשקל של השאר.
>
> קודם היו כאן אחד־עשר חריגים, ושישה מהם ירדו מהרשימה כשהחבילה נמשכה מ־npm
> ונקרא ממנה ה־path data עצמו: `header` ו־`footer` קיימים כ־`document_header`
> ו־`document_footer` — ההנמקה הקודמת, „אין ב־Fluent גליף לכותרת עליונה או
> תחתונה בכלל”, פשוט לא הייתה נכונה — ו־`bookmark`, `pageNumber`, `dateTime`
> ו־`updateFields` נחסמו רק בשל היעדר גישה ל־path data המקורי, מחסום שהמשיכה
> הסירה. שביעי, `exportPdf`, לא היה ברשימה מעולם: הוא נוסף אחרי כתיבת המסמך הזה
> ולא נרשם בו כלל — לא בטבלה ולא בין החריגים — ואף בדיקה לא התלוננה. `check:icons`
> קיים בדיוק בשביל שני הכיוונים האלה: הנמקה שהתיישנה, ואייקון שנוסף בשקט.

## רכיבי MIT שנארזים

נכנסים לחבילה דרך התלויות של superdoc ושל הממשק. הודעות הרישוי שלהם נאספות
אוטומטית לסוף `assets/app.js` בבנייה (`esbuild.legalComments: 'eof'`):

| רכיב | גרסה שנמדדה | רישיון |
|---|---|---|
| vue (ו-`@vue/*`) | 3.5.41 | MIT |
| pinia | 3.0.4 | MIT |
| superdoc-macros | 0.9.0 | MIT |

הרשימה נמדדת מהפלט, לא מהצהרה: `grep '@license' dist/assets/app.js` מציג את
מה שנארז בפועל. אם תיווסף תלות עם רישיון שאינו MIT/BSD/ISC — יש לתעד אותה כאן
לפני פרסום.

## המילון התורני — `src/data/torah-dictionary.txt`

102,465 ערכים: מונחים תורניים, ראשי תיבות וצורות כתיב מקובלות. נארז לתוך
`assets/torah-dictionary.js` ונטען רק כשבדיקת האיות נדלקת.

- מקור: נתרם למאגר ב־[issue #25](../../issues/25) בידי YairDaniel11, שבנה
  אותו מטקסט תורני והצהיר שם שהוא חופשי לשימוש בפרויקט.
- הרשימה עצמה היא **מילים בעברית** — עובדות לשוניות, ולא ביטוי יצירתי — ולכן
  אינה נושאת זכויות יוצרים עצמאיות. אין בה טקסט רץ מתוך שום מקור.
- מנגנון ההתאמה סביבה (נרמול, תחיליות בשני הכיוונים) נכתב מחדש
  ב־`src/engine/spellcheck.ts`; מהענף שהוצע בעקבות ה־issue נלקחו ההחלטות
  שנמדדו שם, לא הקוד.

## רשימות השלמה סטטיות — `src/data/talmudic-phrases.json`, `src/data/authors.json`

שתי רשימות של מחרוזות קצרות, שנארזות יחד ל־`assets/static-completion.js`
(32KB) ונטענות רק אחרי שההשלמה מהספר לא מצאה התאמה.

- **`talmudic-phrases.json` (325 ערכים)** — כותרות הערכים בקטגוריית ויקימילון
  „[ניבים, ביטויים ופתגמים מהתלמוד ומהמשנה](https://he.wiktionary.org/wiki/קטגוריה:ניבים,_ביטויים_ופתגמים_מהתלמוד_ומהמשנה)”,
  רישיון **CC BY-SA + GFDL** (ויקימדיה). מה שנלקח הוא **כותרות בלבד** — ניב
  תלמודי הוא עובדה לשונית ולא ביטוי יצירתי עצמאי, אותו נימוק שחל על
  `torah-dictionary.txt` למעלה. לא נלקחה שום פרוזה הסברית מגוף הערכים.
- **`authors.json` (651 ערכים)** — טבלת `author` ב־`seforim.db` של אוצריא.
  נתוני מטא-דאטה של אוצריא עצמה, ולא תוכן ספר. אינה בהכרח הרשימה המלאה של
  ספריית ברירת המחדל.

**מה במכוון לא נכנס:** ל־PR ‏#46 היה גם
`talmudic-phrases-community-wikis.json` (2,189 ערכים) מוויקישיבה
(**CC BY-NC 3.0**) ומוויקיסוגיה (**רישיון לא ידוע**). מבחינה טכנית הוא זול
(~1.2MB heap, כחצי מילישנייה לסריקה), אבל שני הפרופילים האלה שונים מהותית
מהמסגרת של המסמך הזה, והתאמתם ל־AGPL-3.0 ולהפצה דרך אוצריא לא הוכרעה. הוא
נשאר מחוץ להפצה עד שתוכרע — כאן החסם המשפטי גדול מהחסם הביצועי.

## מילון ראשי-תיבות — `src/data/acronyms.json`

25,363 ראשי-תיבות ופירושיהם, משני מקורות ממוזגים. נארז כנכס מקומי נפרד
(`assets/acronyms.js`) ונטען רק כשהמשתמש הקליד מילה בצורת ראשי-תיבות; הוא
אינו נכלל ב־`assets/app.js`.

- **מילון אוצריא הרשמי** (`Acronyms.json`, 13,105 ערכים) — נתון ספרייה של
  אוצריא, ולא תוכן של ספר חיצוני.
- **[KleiKodeshProject](https://github.com/KleiKodesh/KleiKodeshProject)**
  („פרוייקט כלי קודש לוורד — ארגז כלים לעורך התורני”), 12,258 ערכים נוספים —
  רישיון **Apache License 2.0**. הפרויקט עצמו מציין כמקורותיו את Wikipedia,
  ויקי ספרי יהדות, לעזי רש״י ואוצריא.

במיזוג הופעלה עדיפות: ערך שקיים בשני המקורות — הפירוש של אוצריא נשאר ראשון
ברשימה, ופירושי KleiKodesh מתווספים אחריו. הנכס נושא את הפירוש הראשון בלבד
(`scripts/acronyms-asset.ts`), וקובץ המקור נשאר שלם.

## קוד שהועתק ממאגרים אחרים

ה־path data של 72 אייקוני הרצועה — ראו [Fluent System Icons](#fluent-system-icons-11338--mit)
למעלה. הקובץ המושפע הוא `src/ui/icons/icons.ts` בלבד.

מלבד זה הממשק נכתב מאפס. מקורות שהיוו השראה חזותית בלבד (ללא העתקת קוד):

| מקור וקומיט שנבדק | מה מאמצים | מה לא מעתיקים |
|---|---|---|
| [LocalOffice `60bd8cef`](https://github.com/Anon5T4R/LocalOffice/commit/60bd8cef8f135a9cc9183f9a8217a0b8e8d84528) | פירוק Ribbon ללשוניות/קבוצות/כפתורים, mount של הלשונית הפעילה בלבד, dirty indicator, שורת מצב, שמירת format painter במעבר לשונית | פקודות TipTap, מודל המסמך וקוד React |
| [SuperDoc `examples/custom-ui` ב־`b0ff2221`](https://github.com/superdoc/docx-editor/commit/b0ff2221645f79b7094e1c037723fe2a435ffd3c) | `ui: false`, שימוש ב־`superdoc.ui`, observe למצב פקודה, `mousedown.preventDefault()`, `executeAsync`, teardown | אין לסטות אל API פנימי; מוצמד לגרסת 2.8.0 המקורית ולא ל־main משתנה |
| [Herramienta_Optimizacion_PBM `437d79d2`](https://github.com/T0m4s1n/Herramienta_Optimizacion_PBM/commit/437d79d203db44af384861fe588ea5a0dd57724f) | בורר טבלה, overflow אופקי, aria, Escape/outside-click וקיצורי מקלדת | חיבור SuperDoc v1 ו־`headless-toolbar`; אין רישיון מזוהה ולכן אין העתקת קוד |
| [canvas-editor `03a481bb`](https://github.com/Hufe921/canvas-editor/commit/03a481bbd012f2dcb4044cd34471477db921fe52) | רעיונות לסרגל, עמודים, תפריטים הקשריים לטבלה/תמונה/קישור וחלוניות | import/export DOCX ומודל ה־Canvas אינם מקור אמת |
| [ONLYOFFICE/web-apps `9c0ca538`](https://github.com/ONLYOFFICE/web-apps/commit/9c0ca538c3b211052347df09d2a4d6781f023403) | compact/full Ribbon, לשוניות הקשריות, state מרכזי לנעילת פקדים, ניווט ושורת מצב | אין לחלץ את מסגרת ה־frontend הגדולה או controllers שלה |

כל העתקה נוספת מתועדת כאן עם קישור, גרסה, רישיון ורשימת הקבצים המושפעים.
