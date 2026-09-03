# שולחן העורך — מסמך מסירה מלא (handoff) לסשן חדש

**מטרה:** לתת לסשן חדש את כל מה שסשן החקירה של 3.9.2026 ידע, כדי שיוכל להמשיך
בלי לחזור על שום חקירה. קראו קודם את המסמך הזה, אחר כך את
[`shulchan-gap-map.md`](shulchan-gap-map.md) (מפת הפערים המלאה), ורק אז את הקוד.

---

## 1. איפה עומדים (מצב הענף)

- ענף עבודה: `feat/shulchan-fidelity`, נוצר מ-`main` אחרי `git pull` (120 קומיטים fast-forward).
- **שום קובץ מקור ב-`src/` לא שונה עדיין.** נוספו רק מסמכים תחת `docs/`.
- `origin` = `Y-PLONI/otzaria-word-editor` (upstream). הפורק של המשתמש = `palmoni5/otzaria-word-editor`
  (remote בשם `fork`). המשתמש הוא palmoni5, ובעל גם `superdoc-macros` וגם אוצריא.
- `npm ci` הורץ; node_modules עכשיו superdoc 2.11.0 ו-superdoc-macros 0.9.0 (לפני כן היה מיושן:
  2.10.0 / 0.7.1). אם node_modules נעלם — להריץ `npm ci` שוב לפני `npm run typecheck` / `npm test`.

## 2. מה נמצא בתיקיית המקור [`shulchan-source/`](shulchan-source/)

| קובץ / תיקייה | תוכן |
|---|---|
| `vba-main/` | 37 מודולי VBA שחולצו (olevba) מ-`שולחן העורך.dotm` גירסה 4.0 |
| `vba-para/` | 19 מודולים מ-`שולחן העורך - עיצוב פיסקה.dotm` |
| `ribbon-main.xml`, `ribbon-para.xml` | customUI של שתי הרצועות: אילו כפתורים קיימים ולאיזה Sub הם קוראים |
| `engine-facts.md` | עובדות מנוע שנמדדו נגד ה-dist. **לא לחזור על ה-grep** |
| `engine-issues/` | טקסט ה-issues שנפתחו ב-SuperDoc (ראו §4) |

מודולים מרכזיים לפי כלי: `Typos.bas` (FixHebrewPunctuation בשורות ~140-155), `UnclosedParentheses.bas`,
`TextAlternating.bas`, `BracketsAndFootnotes.bas`, `EditingErrors.bas` (אחידות), `FormatFirstWord.bas`,
`HangingFirstWord.bas`, `LastLineBalance.bas`, `CenterLastLine.bas`, `AlignPages.bas`, `AlignColumns.bas`,
`DocReduction.bas`, `PageMarking.bas`, `CropMarks.bas`, `SplitDocument.bas`, `ContinuousFootnotes.bas`,
`Footnotes.bas`, `NoteNumbering.bas`, `DecorationsTitles.bas`, `LineCommentBox.bas`, `TableContents.bas`,
`SettingsHelper.bas` (זכירת הגדרות ב-INI), `RibbonControl.bas` (המתאם „החל עיצובים” בתבנית השנייה).

## 3. עובדות מנוע (SuperDoc 2.11.0) — נמדדו, לקבל כנתון

ראו [`shulchan-source/engine-facts.md`](shulchan-source/engine-facts.md). התמצית:

- **אין API ציבורי לפריסה** (שורות בפסקה, גבולות עמודים, מיקום טווח). קיים פנימי ב-`dist/layout-engine`
  (`ResolvedTextLineItem`, paginator) אך לא מיוצא. `page-ruler.ts` מודד `[data-page-index]` + `getClientRects`
  עם `SETTLE_DELAYS_MS` (היוריסטי; **לא** `pageMetrics.getSnapshot()` כפי שנטען בטעות באחד הדוחות).
- `PARAGRAPH_ALIGNMENTS = ['left','center','right','justify']` — אין כתיבת `distribute`; קריאה `'distributed'` קיימת.
- `text.rewrite where:{by:'target'}` מוצהר ב-`StepWhere` אך נדחה בריצה:
  `"v2 text.rewrite currently requires a ref produced by query.match/find or a single text selector"`.
  `by:'block'` משטח עיצוב ריצות (`onNonUniform:'majority'`). לכן הכלים עושים N קריאות `doc.replace` = N צעדי undo.
  מתועד גם ב-`src/engine/search.ts` (~שורות 496-532).
- אין `create.textbox/shape`; `create.image` מחזיר `OPERATION_UNAVAILABLE`.
- סגנונות: `styles.apply` + `styles.getCatalog` בלבד; אין create/remove; החלת סגנון רק דרך פקודת `linked-style` על הבחירה.
- `footnotes.insert` בלי `customMark`; `footnotes.configure` דורס `footnotePr`; אין `getConfig`.
- **`defaultTabStop=0` מקפיא את המנוע** — `docx-preflight.ts` מתקן ל-720. לכן את שיטת CenterLastLine
  של Word (`DefaultTabStop=0`+`jc=distribute`+טאב) **אסור** לממש מקומית.
- צורות מודל: `SDParagraphProps.alignment?: 'left'|'center'|'right'|'justify'|'start'|'end'|'distributed'|...`;
  צומת פסקה נושא `props` ו-`resolved?: Partial<SDParagraphProps>`; פריטי `blocks.list` נושאים `styleId`, `nodeType`, `text`.

## 4. Issues שנפתחו ב-SuperDoc (superdoc/docx-editor)

| # | נושא | טקסט |
|---|---|---|
| [3969](https://github.com/superdoc/docx-editor/issues/3969) | `text.rewrite where:{by:'target'}` מוצהר אך נדחה — חוסם undo יחיד | `engine-issues/3969-text-rewrite-by-target.md` |
| [3970](https://github.com/superdoc/docx-editor/issues/3970) | API קריא-בלבד לפריסה (`layout.lines/pages/settled`) — חוסם חלון, איזון, יישור עמודים/טורים, סימון עמודים | `engine-issues/3970-layout-read-api.md` |
| [3971](https://github.com/superdoc/docx-editor/issues/3971) | כתיבת `jc=distribute` — חוסם מירכוז שורה אחרונה | `engine-issues/3971-jc-distribute.md` |

עוד לא נפתחו (מופיעים בטבלת המנוע בסוף gap-map): `footnotes.insert({customMark})` + `getConfig`,
`create.image` פעיל + anchor, `create.textbox` מקושרות, `styles.create/remove` + סגנון תו,
`sections.setColumns({widths[]})`, TOC `\t` + PAGEREF. לפתוח לפי הצורך, באותו פורמט.

## 5. ההנחיה האחרונה של המשתמש ומה נגזר ממנה

> „פתח אישוז. ובמקביל החל לעבוד על מה שבטוח לגמרי, לא מה שבקירוב”

„בטוח לגמרי” = שינויי טקסט/אלגוריתם טהורים בתבנית הכלים הקיימת, בלי מדידת DOM.
**לא לממש עכשיו** (קירובים): עיצוב חלון, איזון שורה אחרונה, יישור עמודים/טורים, סימון עמודים, צמצום מסמך,
סימני חיתוך, פירוק מסמך, מירכוז שורה אחרונה.

### 5.1 הרשימה הבטוחה — מפרט לכל פריט

1. **Typos — מרכאות מסולסלות** (`src/engine/shulchan/typos.ts`):
   `doubleApostrophes` היום `/''/g` — לתפוס גם U+2019 ותערובת (`/['’]{2}/g` → `"`);
   `shiftedHebrewAfterQuote` היום `/"([A-Z<>])/g` — להרחיב ל-`/["“”„]([A-Z<>])/g` עם `SHIFTED_HEBREW_MAP`.
   לעדכן `tests/unit/shulchan-typos.test.ts`.
2. **כלי חדש „תיקון העתקה מתוכנות”** (`FixHebrewPunctuation`): NBSP (` `) → רווח רגיל, **פרט** כאשר התו הקודם הוא
   מעבר שורה ידני (`` / `\n`). תחום: הבחירה (`scopedBlocks(host,'selection',...)`). כפתור נפרד ברצועה.
   התבנית השנייה של המקור (נרמול כיווניות פיסוק „הפוך” ע"י self-replace `^&`) — אין שקילות במנוע; לתעד ולדלג.
3. **סוגריים לא סגורים — „כל המסמך כאחד”** (`unclosed-parens.ts`): להוסיף `scanDocumentAsOne(blocks)` שמעביר מחסנית
   בין בלוקים ומדווח פתוח-בלי-סגור רק בסוף המסמך; בורר מצב ב-`ShulchanUnclosedDialog.vue`.
4. **טקסט מתחלף — סט תווים** (`text-alternating.ts`, `ShulchanAlternatingDialog.vue`): `startChar/endChar` יהפכו לסטים
   (כל תו בסט תוחם); להסיר `length!==1 → []`, להסיר `maxlength="1"` ואת דרישת `startChar !== endChar` ב-`canSubmit`.
5. **הערות ⟵ סוגריים לפי הבחירה** (`brackets-notes.ts`): `convertFootnotesToBrackets` מעבד היום את **כל** `listFootnoteRefs`;
   לסנן לפי blockIds+ranges של הבחירה. `scopedBlocks` ב-`shulchan-doc.ts` מחזיר ids בלבד — להרחיב כדי להחזיר גם
   `range:{start,end}` מ-`selection.current()`. לעדכן תיאור הכלי ב-`tools-registration.ts` („תוכן כל הערות השוליים”).
6. **אחידות טורים** (`sections-uniform.ts`): `readColumnsProfiles`/`applyColumnsProfile` — לעבוד רק על `count === 2`
   (כמו המקור) ו**לא** לכתוב `count` בהחלה (היום `count<2 → skip` ו-`count: profile.count` דורס מקטע 3 טורים).
7. **מילה ראשונה** (`first-word.ts`, `ShulchanFirstWordDialog.vue`): `skipHeadings` בודק רק `nodeType!=='paragraph'` —
   להוסיף דילוג על `alignment==='center'` (לקרוא props/resolved של הפסקה בתבנית `findParagraphProps({body},blockId)` כמו ב-`line-spacing.ts`);
   סינון אופציונלי לפי `styleId` (רשימה מ-`styles.getCatalog`); לתקן תווית מצב `fixed` (במקור „גודל קבוע”=גודל הגוף, אצלנו נק' מוחלטות).
8. **זכירת הגדרות לכל הדיאלוגים**: `src/host/settings.ts` מחזיק היום helpers לפי מפתח מעל `tryCall('storage.get'|'storage.set',{key,value})`
   בלי `loadSetting` גנרי. להוסיף `loadSetting<T>(key, fallback)` / `saveSetting(key,value)` וקומפוזבל (למשל `useRememberedDialog`)
   שכל דיאלוג Shulchan משתמש בו; לשמור ב-submit ובביטול. היום `ShulchanTyposDialog`/`ShulchanFirstWordDialog` מאפסים ל-`default*Options()` בכל פתיחה.
9. (נדחה, אפשרי) מתאם „החל עיצובים” בסדר קבוע מרווח שורות ⟵ מילה ראשונה — רק אחרי 1-8.

### 5.2 כללי מבחנים שחייבים להישמר

- `tests/unit/tab-controls.test.ts`: **כל** `<RibbonButton|RibbonMenuButton ... />` ב-`ShulchanTab.vue` חייב `:disabled=` (כפתור חדש → `:disabled="!ready"`).
- `tests/component/ribbon-tabs.test.ts` לוחץ על כל כפתור ומצפה לאפקט נצפה (דיאלוג נפתח / כלי רץ).
- מבחני יחידה עם `fakeShulchanHost({blocks, selected, runs, spacing, notes, refs, sections})` מ-`tests/unit/shulchan-fake.ts`
  (מחזיר `{host, calls, textOf, blocks}`). מבחני הגדרות ממקמים `window.Otzaria = { call }` שמחזיר `{success:true,data}`.
- לפני סיום: `npm test` ו-`npm run typecheck`. לעדכן `docs/shulchan-haorech.md` לכל כלי שמשתנה.

### 5.3 תבנית הכלים (איך כלי Shulchan כתוב)

`src/engine/shulchan/*`: `readShulchanBlocks` / `scopedBlocks(host,'selection'|'document',failedAction)` → חישוב JS טהור →
כתיבות נקודתיות `doc.replace` (`replaceRange`, `textTarget(blockId,start,end)`), `applyInline`/`doc.format.apply`,
`doc.format.paragraph.*`, `doc.sections.*`, `doc.footnotes.*`; `revealRange`, `unavailableOutcome`. אין נגיעה ב-DOM של המנוע.
רישום ב-`tools-registration.ts` דרך `kit.registerTool` (superdoc-macros 0.9.0); UI ב-`src/ui/ribbon/tabs/ShulchanTab.vue`
(`runTool(commandId, action, summary)`, קבוצות: הגהה / הערות שוליים / עיצוב פסקה / אחידות מסמך).

## 6. מה תלוי עדיין במנוע (סיכום)

| יכולת שחסרה במנוע | כלים שנחסמים | סטטוס |
|---|---|---|
| layout קריא-בלבד | חלון, איזון, יישור עמודים/טורים, סימון עמודים מדויק, פירוש תחת המילים, ספי מספר-שורות | issue 3970 |
| `by:'target'` ב-`text.rewrite` | undo יחיד לכל הכלים | issue 3969 |
| `setAlignment('distribute')` | מירכוז שורה אחרונה | issue 3971 |
| `footnotes.insert({customMark, body})`, `getConfig`, `'hebrew1'` | הערה ללא מספר, מספור עברי, פירוק הערות סיום | לא נפתח |
| `create.image` פעיל + anchor behindText | עיטורים (שני הסוגים) | לא נפתח |
| `create.textbox` מקושרות + `create.shape` בפוטר | פירוש תחת המילים, הערות ברצף, סימני חיתוך בקובץ | לא נפתח |
| `styles.create/remove`, סגנון תו לפי target | „עם סגנון” במילה ראשונה, החלפת/מחיקת סגנונות | לא נפתח |
| `sections.setColumns({widths[]})` | אחידות טורים מלאה | לא נפתח |
| TOC `customStyles` (`\t`) + PAGEREF | תוכן עניינים לפי סגנונות חופשיים / פורמט רציף | לא נפתח (למדוד קודם אם `\t` נשמר ב-create) |

מה **לא** דורש מנוע ולא אוצריא: כל §5.1, וגם הקירובים שנדחו (צמצום, סימון עמודים, סימני חיתוך בהדפסה, פירוק שוליים).
אוצריא כבר מספקת כל הנדרש (`storage.*`, `fs.pickUserFile`, „שמור בשם” בלי `targetToken`, `ui.exportPdf({pageSize})`).

## 7. תקלות שנפגשו בסשן (כדי לא לבזבז זמן)

- סוכני-משנה נפלו עם HTTP 418 / „JSON Parse error: Unexpected EOF” — **לא** חסימת נטפרי (המשתמש אישר), תקלה זמנית.
  פתרון: סוכנים קטנים יותר עם `engine-facts.md` משותף.
- כתיבת קבצים דרך Bash heredoc נכשלה על מרכאות מקוננות ועל תווי בקרה (NBSP, VT) בתוך הפקודה — להשתמש בכלי Write.
