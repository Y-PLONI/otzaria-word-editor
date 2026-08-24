# הודעות רישוי של רכיבי צד שלישי

התוסף עצמו מופץ תחת [AGPL-3.0](../LICENSE). המסמך הזה מפרט את הרכיבים
שנארזים לתוך ה־`.otzplugin` ואת חובות הרישוי שלהם. העתקים של נוסחי הרישיון
נארזים תחת `third-party/` בתוך החבילה עצמה, לא רק במאגר.

## superdoc 2.8.0 — AGPL-3.0

- מקור: <https://github.com/superdoc/docx-editor>
- רישיון: AGPL-3.0 (`third-party/SUPERDOC-LICENSE.txt`)
- הודעה: `third-party/SUPERDOC-NOTICE.txt`

זו הסיבה שהתוסף כולו AGPL-3.0, שהמקור מפורסם, ושהמקור המפורסם זהה לבינארי
המופץ.

## @superdoc/docx-engine 0.7.0 — קנייני

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
| `print` | הדפסה | `print` |
| `info` | אודות | `info` |

**לשונית „בית” — לוח**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `paste` | הדבק | `clipboard_paste` |
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
| `fontColor` | צבע גופן | `text_color` |
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
| `lineSpacing` | מרווח שורות | `text_line_spacing` |
| `borders` | גבולות | `border_all` |
| `shading` | צביעה | `paint_bucket` |

**לשונית „בית” — עריכה**

| שם ב־`ICONS` | פקד | אייקון מקורי |
|---|---|---|
| `search` | חיפוש (וגם סרגל הגישה המהירה) | `search` |
| `replace` | החלפה | `arrow_swap` |
| `select` | בחר הכל | `select_all_on` |

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
  קוד של התוסף.
- **אין אייקוני מיתוג.** ה־MIT מכסה את האייקונים, אך סמלי לוגו ומוצר של
  Microsoft הם סימני מסחר. נלקחו אייקוני ממשק גנריים בלבד — דף, תיקייה,
  דיסקט, מדפסת, חץ, עיגול מידע. `word` ו־`otzaria` נשארו ציור בבית.

למה דווקא Fluent System Icons ולא Fluent MDL2, שדומה יותר ל־Ribbon של Word:
MDL2 היא הספרייה של Office UI Fabric שיצאה משימוש, System Icons היא הסט
הפעיל של Fluent 2 תחת MIT, יש לה גריד 20 שמתאים בדיוק ל־`viewBox` של הסט
הקיים, ויש לה וריאנטי RTL מוצהרים — `arrow_export_rtl` מול `arrow_export_ltr`
— שזה בדיוק מה שנדרש בממשק עברי.

שתי בחירות שאין להן מקבילה מדויקת אצל Microsoft, ולכן הן מתועדות כאן:

- `replace` הוא `arrow_swap` — אין ב־System Icons אייקון „מצא והחלף”.
- `formatPainter` הוא `paint_brush` ולא `clipboard_brush`, כי זה מה ש־Word
  מציג, ולוח כבר מופיע בשלושת השכנים שלו באותה קבוצה.

> המיגרציה נעשית לשונית־לשונית. עברו „קובץ” ו„בית”; נשארו הוספה, פריסה,
> הפניות, סקירה, תצוגה, אוצריא, ה־chevron־ים, `undo`/`redo`, ו־`word`/`otzaria`
> שהם מיתוג ויישארו ציור בבית. הטבלאות כאן מתעדכנות עם כל לשונית שעוברת.

## רכיבי MIT שנארזים

נכנסים לחבילה דרך התלויות של superdoc ושל הממשק. הודעות הרישוי שלהם נאספות
אוטומטית לסוף `assets/app.js` בבנייה (`esbuild.legalComments: 'eof'`):

| רכיב | גרסה שנמדדה | רישיון |
|---|---|---|
| vue (ו-`@vue/*`) | 3.5.41 | MIT |
| pinia | 3.0.4 | MIT |

הרשימה נמדדת מהפלט, לא מהצהרה: `grep '@license' v2/dist/assets/app.js` מציג את
מה שנארז בפועל. אם תיווסף תלות עם רישיון שאינו MIT/BSD/ISC — יש לתעד אותה כאן
לפני פרסום.

## קוד שהועתק ממאגרים אחרים

ה־path data של אייקוני לשונית „קובץ” — ראו [Fluent System Icons](#fluent-system-icons-11338--mit)
למעלה. הקובץ המושפע הוא `src/ui/icons/icons.ts` בלבד.

מלבד זה הממשק נכתב מאפס. מקורות שהיוו השראה חזותית בלבד מתועדים ב־
[../docs/word-plugin-implementation-plan.md](../docs/word-plugin-implementation-plan.md) §3.3.
כל העתקה נוספת מתועדת כאן עם קישור, גרסה, רישיון ורשימת הקבצים המושפעים.
