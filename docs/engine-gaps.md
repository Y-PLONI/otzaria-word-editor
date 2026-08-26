# פערים שנמדדו במנוע — ומה לא נשלח בגללם

המסמך הזה קיים כדי שאיש לא יחקור שוב את מה שכבר נמדד. כל שורה כאן היא
תוצאה של הרצה ב-Chrome אמיתי מול ה-`dist` הארוז, ולא קריאה של טיפוסים.

**הכלל שנגזר מכל אלה:** `available: true` בקטלוג היכולות ו-`success: true`
בקבלה **אינם הוכחה שהפעולה עובדת**. פעולה שכותבת קוד שדה של Word חייבת
אימות מול תיעוד Word, ורצוי גם מול ה-docx המיוצא.

המתכון למדידה נמצא בסוף `docs/spike.md` ובקוד: `scripts/cdp.mjs`.

## פעולות שמדווחות הצלחה וכותבות מסמך שבור

### `crossRefs.insert` — הפניה מקושרת
נמדד פעמיים, בשני סבבים בלתי תלויים, על 9 סוגי תצוגה ו-6 סוגי יעד. קוד
השדה שנכתב:

    REF SDXREF kind=bookmark;value=%7B%22kind%22%3A%22bookmark%22...%7D;display=pageNumber

האסימון שאחרי `REF` הוא `SDXREF` ולא שם הסימנייה, ולכן Word יציג „שגיאה!
מקור ההפניה לא נמצא”. גם המנוע עצמו אינו פותר אותו: `resolvedText` נשאר
ריק אחרי `rebuild` בכל הצירופים. הפניה לסימנייה שאינה קיימת מחזירה
`success: true`.

**לא נשלח.** מה שכן נשלח: `crossRefs.list` תופס שדות `REF` שנוצרו ב-Word,
ו-`rebuild` עליהם מחשב באמת.

### `authorities.entries.insert` — סימון ציטוט לטבלת מקורות
ה-`instruction` שנכתב:

    TA "בראשית א, א" \s "בר׳ א א" \c 1

ל-`TA` של Word **אין ארגומנט כללי**; התחביר הוא `{ TA [switches] }`,
והציטוט הארוך מגיע רק מ-`\l`. כלומר Word יקרא שדה בלי ציטוט ארוך, והערך
יופיע ריק בטבלה. אומת בשלוש שכבות: `entries.get`, ה-docx המיוצא עצמו
(`<w:fldSimple w:instr="TA &quot;בראשית א, א&quot; ...">`), והבנאי במנוע —
תבנית קשיחה שאין בה מסלול שפולט `\l`.

בנוסף אין שום בריחה של גרשיים: `longCitation` שמכיל `"` נכתב כמות שהוא
ומייצר גרשיים מקוננים, עם `success: true`.

### `index.entries.insert` — השדה `subEntry`
`{text:'אבות', subEntry:'יצחק'}` כותב `XE "אבות" \s "יצחק"`, ו-`\s` אינו
מתג של `XE` ב-Word (המתגים: `\b \f \i \r \t \y`).

**נעקף:** הצורה הקנונית `XE "אבות:יצחק"` עובדת, והמנוע מפרק אותה נכון
בחזרה. המודול שולח תמיד אותה ולעולם לא את `subEntry`.

### `bibliography.configure` — סגנון הביבליוגרפיה
הקריאה עובדת בצד אחד ושבורה בצד שני, ושניהם נמדדו באותו קובץ מיוצא.
הסגנון **כן** מגיע למקום הנכון: `configure({style:'Chicago'})` כתב
`<b:Sources SelectedStyle="/CHICAGO.XSL" StyleName="Chicago" Version="16">`,
ואחד עשר השמות הקנוניים ממופים נכון. אבל אותה קריאה כותבת גם ל-instruction:

    BIBLIOGRAPHY \sdStyle "Chicago"

`\sdStyle` אינו מתג של Word — המתגים המתועדים לשדה `BIBLIOGRAPHY` הם `\l`
ו-`\f` — ואין דרך לבקש את הראשון בלי השני.

גם המסלול השני כותב אותו: `bibliography.insert({style:'Chicago'})` מייצר
את אותו `BIBLIOGRAPHY \sdStyle "Chicago"` (נמדד). כלומר אין קריאה שמכניסה
ביבליוגרפיה עם סגנון ובלי המתג הלא-מתועד.

**לא נשלח פקד סגנון.** בלעדיו כל קוד שדה שנכתב הוא קנוני, והסגנון נשאר
ברירת המחדל שגם Word מתחיל בה (APA).

### `citations.insert` עם יותר ממקור אחד
שני מקורות כותבים `CITATION src-a;src-b` (נמדד גם ב-docx). תחביר ריבוי
המקורות של Word הוא המתג `\m`: `{ CITATION Tag1 \m Tag2 }`. אסימון אחד
שמחבר שני תגים בנקודה ופסיק אינו tag קיים.

**נעקף:** המודול שולח תמיד מקור אחד, וגם הממשק מאפשר רק אחד.

### `citations.sources.remove` על מקור מצוטט
מחזיר `success: true`, מוחק את המקור, ומשאיר את שדה ה-`CITATION` מצביע
לתג שכבר אינו קיים — כלומר בדיוק המסמך השבור של `crossRefs`, רק שכאן
אנחנו אלה שיוצרים אותו.

**נעקף:** `removeCitationSource` סופר את הציטוטים דרך `citations.list`
ומסרב, ומדווח כמה מהם מחזיקים במקור.

### `captions.update` — עריכת טקסט של כיתוב
הפעולה **אינה מחליפה את הטקסט אלא מוסיפה עליו**. שלושה צעדים רצופים על
אותו כיתוב:

    insert 'אלף'   → get 'אלף'
    update 'בית'   → get 'אלף: בית'
    update 'גימל'  → get 'אלף: בית: גימל'

אומת ב-docx ולא רק בקבלה: הריצה שאחרי השדה מכילה
`<w:t xml:space="preserve">: שרטוט המשכן: </w:t>` — הישן, המפריד והחדש.
`patch: { text: '' }` אינו מוחק אלא מוסיף מפריד ריק, ו-`patch` בלי `text`
(או עם שדה שאינו בחוזה) מוחזר `NO_OP`.

**נעקף:** עריכה היא `remove` ואז `insert` באותו מקום, והעוגן נקרא **לפני**
ההסרה. נמדד שהתוצאה זהה לכיתוב שנוצר מאפס, שהמיקום נשמר ושהמספור מתעדכן.

בחירת העוגן היא שני צעדים, ולא אחד, מפני ש-`captions.insert` מקבל פסקה
בלבד (ראו הסעיף הבא):

1. הבלוק **שלפני** הכיתוב, עם `position: 'below'`.
2. אם הוא אינו פסקה — הבלוק **שאחרי** הכיתוב, עם `position: 'above'`.

שני העוגנים מצביעים על אותו רווח בדיוק, ולכן הנפילה-לאחור אינה מזיזה את
הכיתוב. נמדד על `פסקה │ tbl │ כיתוב │ פסקה`: העוגן הראשון
(`tbl:41964672`) הוחזר `TARGET_NOT_FOUND`, השני התקבל, וסדר הבלוקים אחרי
העריכה זהה תו-בתו לסדר שלפניה — הכיתוב בין הטבלה לפסקה, ו-`captions.list`
מדווח את אותו מספר (`טבלה 1`). זה חשוב מפני שכיתוב מתחת ללוח הוא הצורה
השכיחה ביותר, וכל docx מיובא מלא בהם.

הסירוב נשאר לשני מצבים, ובשניהם **לפני** שנגעו במסמך: כיתוב שהוא הבלוק
היחיד, וכיתוב ששני שכניו אינם פסקאות (טבלה מכאן וטבלה מכאן). ומעליהם רשת
ביטחון: הוספה שנכשלה אחרי הסרה שהצליחה מנסה להחזיר את התוכן הישן, ורק
כשגם השחזור נכשל ההודעה מודה שהכיתוב הוסר ומפנה ל-Ctrl+Z.

### `captions.insert` — עוגן שאינו פסקה
`adjacentTo` מקבל כתובת שהיא `nodeType: 'paragraph'` בלבד, ובלוק שאינו
פסקה מוחזר `TARGET_NOT_FOUND` („target paragraph tbl:… was not found”).
נמדד על טבלה שנוצרה ב-`create.table`: `tbl:41964672` כעוגן נדחה, בעוד
שפסקה רגילה באותו מסמך התקבלה. כלומר הכיתוב הטיפוסי של Word — זה שיושב
**מתחת ללוח** — הוא בדיוק המקרה שאין לו עוגן קביל בבלוק שלפניו.

זה מה שהופך את מסלול ה-`remove`+`insert` של העריכה למסוכן: אם העוגן נבחר
מהבלוק שלפני הכיתוב בלי לבדוק את סוגו, ההסרה מצליחה וההוספה נכשלת —
כלומר כיתוב שנמחק.

**נעקף:** סוג הבלוק נקרא מ-`blocks.list` (`nodeType`), והעריכה נופלת
לאחור אל הפסקה שאחרי הכיתוב עם `position: 'above'` — אותו רווח בדיוק.
סירוב רק כששני השכנים אינם פסקאות, ולפני ההסרה. „נקי” בסעיף הכיתובים
שלמטה מתייחס לקוד השדה שנכתב, לא לקבילות העוגן.

### `images.insertCaption` — כיתוב לתמונה
התווית קשיחה ובאנגלית. אין ב-`InsertCaptionInput` שדה `label` בכלל
(`{ imageId, text }`), והמימוש כותב `SEQ Figure` וטקסט `Figure <n> <text>`
— בלי נקודתיים ובלי `\* ARABIC`. זו בדיוק הבעיה שהפילה את טבלת המקורות,
רק שכאן אין אפילו פרמטר לנסות דרכו.

בנוסף אין לתוסף דרך להביא תמונה למסמך מלכתחילה: `create.image` אינה
זמינה, ו-`doc.insert` של HTML עם `<img src="data:…">` נדחה
(`INVALID_PAYLOAD`, „HTML produced no safe canonical content”).

**לא נשלח.** `captions.insert` עושה את אותו דבר טוב יותר — הוא מקבל תווית
עברית, והעוגן שלו הוא הפסקה, כולל הפסקה שהתמונה יושבת בה.

### `captions.configure` — מספור הכיתובים
אינרטית לגמרי. `configure({label:'איור', format:'upperRoman'})` חזר
`success: true`, והכיתוב הבא נכתב `SEQ איור \* ARABIC` — ה-`format` אינו
מגיע לשום מקום, גם לא כשהוא `'zigzag'`. `includeChapter` הוא היחיד שאומר
את האמת: `CAPABILITY_UNAVAILABLE / caption-include-chapter-unsupported`.

**לא נשלח פקד מספור.** המספור נשאר ערבי, מה שגם Word מתחיל בו.

### `footnotes` — כתובת אחת לשתי הערות שונות
`FootnoteAddress` הוא `{ kind:'entity', entityType:'footnote', noteId }`,
ו-**`entityType` הוא `'footnote'` גם עבור הערת סיום**. שני הרצפים מתחילים
מ-1 בנפרד, ולכן במסמך שיש בו הערת שוליים 1 והערת סיום 1 שתיהן נושאות את
אותה כתובת בדיוק, ואת אותו `handle.ref` (`footnote:1`) ואת אותו `id`
ב-`list`. `type` הוא ההבדל היחיד, והוא **אינו** בכתובת.

מה שנמדד: `get`/`update`/`remove` על הכתובת הזאת פוגעים תמיד ב**הערת
השוליים**; אחרי שהיא הוסרה, אותה כתובת פוגעת בהערת הסיום;
ו-`entityType: 'endnote'` נזרק („target must be a FootnoteAddress …
entityType 'footnote'”). כלומר „הסר” על הערת סיום מוחק הערת שוליים אחרת,
עם `success: true`.

**נעקף:** לפני כל עריכה והסרה נקרא `footnotes.get` על הכתובת, והפעולה
מסרבת כשהסוג שחזר אינו הסוג שהמשתמש בחר — לפני שנגעו במסמך. `get` ולא
השוואה מול `list`, כי הוא מודד את אותו מסלול שהמוטציה תלך בו.

**וההגנה עצמה היא TOCTOU, ודורשת נעילה בצד הממשק.** בין ה-`get` ובין
המוטציה יש חלון: `get` נפתר מעבר לגבול macrotask (נמדד ~10ms במסמך ריק
וקר, וגדל עם גודל המסמך), ובזמן שהוא באוויר לחיצה על „הערת שוליים”
ברצועה נקלטת ומוסיפה הערה — כלומר משנה את מה שהכתובת נפתרת אליו. מה
שנמדד: „הסר” על הערת סיום 1, שאושר מפני שלא הייתה אז הערת שוליים 1, מחק
את הערת השוליים **החדשה**, הערת הסיום נשארה, והמשתמש קיבל „בוצע”
(שוחזר במנוע האמיתי, פעם אחת בשש חזרות).

`get` נוסף אינו פותר — הוא רק מקצר את החלון. מה שסוגר אותו הוא שההוספה
לא תיקלט כל עוד פעולה על הערה באוויר: `inFlight` ב-`ReferencesTab.vue`
שנדלק כשהפעולה יוצאת ונכבה ב-`finally`, ומנטרל בזמן הזה את שני כפתורי
ההוספה שברצועה ואת כפתורי הדיאלוג. **זו מלכודת שתחזור בכל פעולה עתידית
על הערות**: כל אימות שנעשה מול המנוע לפני מוטציה על כתובת שאינה מזהה
את מושאה באופן חד-משמעי הוא בדיקה של מצב שעלול להתחלף עד המוטציה, וכל
מסלול שיכול לשנות את המצב הזה חייב להיות נעול בזמן שהיא באוויר.

### `footnotes.configure` — כותב קנונית, ובכל זאת לא נשלח
זה ה-`configure` הראשון מאז גל 3 שבאמת מגיע לקובץ. `settings.xml` קיבל
`<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="4"/>
<w:numRestart w:val="eachPage"/><w:pos w:val="beneathText"/></w:footnotePr>`,
ובמסמך נקי אין `w:footnotePr` כלל עד לקריאה הראשונה. ובכל זאת אין לו פקד,
משלוש סיבות שנמדדו:

1. **אין קריאה.** אין בכל ה-API דרך לקרוא את ההגדרות שבמסמך — לא
   ב-`footnotes`, לא ב-`sections` ולא ב-`info`. דיאלוג היה מציג ערכים
   שאינם של המסמך שעל המסך, וזה בדיוק מה שנאסר בדיאלוג של תוכן העניינים.
2. **כל קריאה מחליפה את האלמנט כולו, ואינה מטליאה אותו.** `configure`
   מלא ואחריו `configure({ numbering: { start: 9 } })` משאיר
   `<w:footnotePr><w:numStart w:val="9"/></w:footnotePr>` בלבד, ו-
   `numbering: {}` משאיר `<w:footnotePr></w:footnotePr>` ריק. כלומר אישור
   אחד בטופס שאינו יודע מה היה במסמך מוחק את מה שהוגדר ב-Word.
3. **שלושה ערכים שכן בחוזה נכתבים כאסימונים שאינם של Word:**
   `restartPolicy:'eachSection'` → `eachSection` (התקן: `eachSect`),
   `format:'symbol'` → `symbol` (התקן: `chicago`), ומיקום הערת סיום →
   `sectionEnd`/`documentEnd` (התקן: `sectEnd`/`docEnd`). מיקום הערת
   שוליים (`pageBottom`/`beneathText`) וחמשת פורמטי המספור הלטיניים כן
   תקניים.

**מספור עברי אפשרי, ואינו נשלח.** `numFmt` נכתב גולמית, ולכן
`format: 'hebrew1'` מייצר `<w:numFmt w:val="hebrew1"/>` — אסימון תקני של
Word ובדיוק המספור שספר תורני רוצה. `'hebrew1'` אינו ב-union של
`FootnoteNumberingConfig`, כלומר זו הישענות על ערך שאינו בטיפוסים
הציבוריים, והבריף אוסר אותה. הממצא מדווח למפקח.

`scope: { kind: 'section' }` הוא היחיד שאומר את האמת:
`CAPABILITY_UNAVAILABLE` („section-scoped note configuration is not
supported by v2 yet”).

## `doc.sections` — ה-namespace היחיד שכן מאמת קלט

נמדד בגל 10 (פריסת עמוד מתקדמת), Chrome headless על ה-dist הארוז, כולל פירוק
ה-zip של `export.toDocx`. זו ההפך מכל תשעת הגלים שקדמו: ערך שאינו ב-union
**נזרק** בזמן ריצה ואינו נבלע.

    setLineNumbering.restart: 'zigzag'  → „must be one of: continuous, newPage, newSection.”
    setLineNumbering.countBy: 0/-3/2.5  → „must be a positive integer.”
    setLineNumbering.enabled חסר        → „must be a boolean.”
    setVerticalAlign.value: 'zigzag'    → „must be one of: top, center, bottom, both.”
    setBreakType.breakType:'nextColumn' → „must be one of: continuous, nextPage, evenPage, oddPage.”
    setPageNumbering.format: 'hebrew1'  → „must be one of: decimal, lowerLetter, upperLetter,
                                            lowerRoman, upperRoman, numberInDash.”
    setPageBorders.borders.display      → „must be one of: allPages, firstPage, notFirstPage.”
    sections.list limit: -3             → „limit must be a positive integer.”
    target מומצא                        → קבלה `TARGET_NOT_FOUND` („Section 'section-99' was not found.”)

וכל אסימון שכן ב-union **הוא אסימון Word תקני**: שלושת ערכי
`ST_LineNumberRestart`, ארבעת ערכי `ST_VerticalJc`, ארבעת ערכי `ST_SectionMark`
וששת ערכי `ST_NumberFormat` — אין כאן `eachSection` מול `eachSect`. זו
הקבוצה הראשונה שאין בה אף אסימון פנימי.

### מה שנכתב, ונמדד ב-docx

    <w:type w:val="oddPage"/>
    <w:pgBorders w:display="allPages" w:offsetFrom="text" w:zOrder="front">
      <w:top w:val="double" w:sz="12" w:space="24" w:color="FF0000" w:shadow="0" w:frame="0"/>…
    <w:lnNumType w:countBy="5" w:start="1" w:distance="360" w:restart="newPage"/>
    <w:pgNumType w:start="3" w:fmt="upperRoman"/>
    <w:vAlign w:val="center"/>
    <w:pgMar … w:header="1008" w:footer="864"/>

הכול קנוני ובסדר האלמנטים של `CT_SectPr`, ובאותה יחידה כמו השוליים:
**אינצ'ים** ב-API, twips ב-XML. גרשיים בתוך `style` מוברחים ל-`&quot;` —
אין הזרקת XML.

### מספור עמודים עברי — **אינו אפשרי**, בניגוד להערות השוליים

זה ההבדל המדויק מול `footnotes.configure`: שם `numFmt` נכתב גולמית ו-
`'hebrew1'` היה מייצר `<w:numFmt w:val="hebrew1"/>` תקני (ולא נשלח רק מפני
שהערך אינו בטיפוסים הציבוריים). כאן ה-union **נאכף בזמן ריצה**, ואין דרך
ציבורית לכתוב `<w:pgNumType w:fmt="hebrew1"/>`. `SectionPageNumberingFormat`
הוא שש אפשרויות בלבד, וכולן לטיניות.

`readPageLayoutState` בכל זאת מסנן פורמט שאינו ב-union: מסמך שנוצר ב-Word
עם מספור עברי **כן** מחזיר `format: 'hebrew1'` מ-`sections.list`, וטופס
שהיה מציג אותו כערך נבחר היה נזרק באישור.

### ההפיכות נמדדה

`setLineNumbering({enabled:false})` מוריד את `<w:lnNumType>` כולו, ו-
`clearPageBorders` מוריד את `<w:pgBorders>` (קריאה שנייה → `NO_OP`). לעומתם
**`<w:pgNumType>` אינו ניתן להסרה**: `setPageNumbering` דורש לפחות שדה אחד
ואין לו `clear`, כלומר „המשך מהמקטע הקודם” של Word אינו ניתן להשגה מהתוסף.
זה כתוב בדיאלוג לפני האישור.

### כל קריאה מחליפה את האלמנט כולו

`setLineNumbering({enabled:true,countBy:1,distance:0})` אחרי קריאה מלאה
השאיר `<w:lnNumType w:countBy="1" w:distance="0"/>` — `start` ו-`restart`
ירדו. זו בדיוק הסיבה ש-`footnotes.configure` לא נשלח בגל 9, ומה שמבדיל כאן
הוא ש**יש קריאה**: `sections.list` מחזיר `lineNumbering`, `pageNumbering`,
`headerFooterMargins`, `verticalAlign` ו-`pageBorders` מלאים, ולכן הפקד משמר
את `countBy`/`start`/`distance` שנקבעו ב-Word במקום למחוק אותם. ההשלמה
נקראת מאותו `sections.list` שהכתובת נלקחה ממנו — ולכן אין חלון TOCTOU בין
הקריאה למוטציה.

### `setBreakType` — עובד, ולא נשלח

כותב `<w:type w:val="oddPage"/>` קנונית ומשנה מקטע קיים (אין צורך ב-
`create.sectionBreak`). מה שאין לו הוא פקד שאפשר להציג: פעולות המקטע חלות
על **כל** המקטעים, ובמסמך בעל מקטע יחיד ה-`w:type` היחיד מתאר איך המסמך
מתחיל — כלומר אינו עושה דבר; ובמסמך מרובה מקטעים הוא הופך פקד של מקטע אחד
לסריקה שמשכתבת את כל מעברי המקטע שנקבעו ב-Word.

### `chapterStyle` ו-`chapterSeparator` נבלעים לגמרי

`setPageNumbering({chapterStyle:1,chapterSeparator:'colon'})` החזיר
`success: true` וכתב `<w:pgNumType/>` **ריק**. שני השדות אינם מגיעים לשום
מקום, ו-`sections.get` אינו מדווח אותם. אין להם פקד.

## פעולות שבולעות קלט בשקט

בכל אלה המנוע מחזיר `success: true` על ערך שאינו בחוזה, אינו חוקי, או
אינו נכתב כלל. **כל ולידציה חייבת לשבת אצלנו, לפני הקריאה.**

| פעולה | מה נבלע |
|---|---|
| `toc.configure` | `tabLeader` (גם `'zigzag'`), `rightAlignPageNumbers`, `includePageNumbers` |
| `index.configure` | `columns` של 0 / 1- / 2.5, שדה שאינו בחוזה, `letterRange:{from:'zigzag'}` → `\p "zigzag-9"` |
| `authorities.configure` | `tabLeader:'zigzag'` → `\l "zigzag"`, שדה שאינו בחוזה |
| `authorities.entries.insert` | `category` של `99`, `0`, `2.5`, `'zigzag'` ואפילו `'פסוקים'` — כולם נכתבים גולמית ל-`\c` |
| `index.insert` | `\c 99` — מעל התקרה של Word (4) |
| `toc.markEntry` | `\l 12` — מעל התקרה של Word (9) |
| `fields.insert` | `DATE \* HEBREW` — מתג לוח השנה נבלע לגמרי |
| `citations.sources.insert` | `fields: {}`, `title: ''`, `title: '   '`, `type: 'zigzag'`, ושדה שאינו בחוזה — כולם `success: true` ונכתבים לקובץ |
| `citations.bibliography.configure` | `style: 'zigzag'` → `SelectedStyle="/zigzag.XSL"`, גיליון סגנון שאינו קיים |
| `captions.insert` | `label: '   '` → `SEQ "   " \* ARABIC`; ירידת שורה בתווית נכתבת **גולמית לתוך קוד השדה**; `text: '   '` נכתב `: ` ואז רווחים בעוד ש-`get` מחזיר `''`; ירידת שורה בטקסט נכתבת גולמית לתוך `<w:t>`; `text: 5` חוזר `success: true` והטקסט נעלם בלי זכר |
| `captions.configure` | `format: 'upperRoman'`, `format: 'zigzag'`, ושדה שאינו בחוזה — הכול `success: true` ואינו נכתב |
| `captions.list` | `limit: -3` מחזיר `total` נכון ורשימה ריקה |
| `footnotes.update` | `patch: {}` ו-`patch` עם שדה שאינו בחוזה — `success: true` בלי `NO_OP` ובלי שינוי |
| `footnotes.configure` | `format:'zigzag'` → `<w:numFmt w:val="zigzag"/>`; `start` של `0`, `-5`, `2.5` ו-`'א'` — כולם נכתבים גולמית ל-`w:numStart`; `position:'zigzag'` → `<w:pos w:val="zigzag"/>`; `restartPolicy:'eachZigzag'` נכתב אף הוא |
| `footnotes.list` | `limit: -3` מחזיר `total` נכון ורשימה ריקה |
| `sections.setPageBorders` | `style` הוא `string` חופשי: `'zigzag'` → `<w:top w:val="zigzag"/>`, ו-`style: ''` מייצר `<w:top w:sz="8"/>` — גבול **בלי `w:val`**, שהיא תכונה נדרשת ב-`CT_Border`. `size: 999` → `w:sz="999"` (התקרה 96), `size: 2.5` → `w:sz="2.5"` בתכונה שהיא מספר שלם, `space: 999` → `w:space="999"` (התקרה 31), `color: '#FF0000'` ו-`color: 'zigzag'` נכתבים כמות שהם ואינם `ST_HexColor` |
| `sections.setHeaderFooterMargins` | `header: 99` → `w:header="142560"`, כלומר כותרת במרחק 2.5 מטר מקצה הדף. אין תקרה |
| `sections.setLineNumbering` | `distance: 999` → `w:distance="1438560"`, ו-`countBy`/`start` של מיליארד נכתבים כמות שהם. `enabled: true` בלי שדה נוסף מייצר `<w:lnNumType/>` ריק |
| `sections.setPageNumbering` | `chapterStyle` ו-`chapterSeparator` — `success: true`, ואינם נכתבים כלל; `{chapterStyle:1}` לבדו מייצר `<w:pgNumType/>` ריק |

## מתג שכן עובד

`fields.insert` עם `DATE \@ "dd/MM/yyyy"` — מתג תמונת-הפורמט **מפורש**
כהלכה, ומתקן גם היסט של יום שיש ב-`DATE` העירום (הוא ISO ב-UTC).

## מה שכן נכתב נכון — ציטוטים

זו הקבוצה הראשונה מאז הסימניות שעברה גם את שכבת ה-docx בלי הסתייגות, ולכן
היא רשומה כאן במפורש: לא כל מה שהמנוע כותב שבור.

- המקורות יושבים ב-`customXml/item1.xml` כ-`<b:Sources>` בסכימת OOXML,
  עם `itemProps1.xml` שמצהיר על ה-`schemaRef`, רלציה מ-`document.xml.rels`
  ו-`Override` ב-`[Content_Types].xml`. זה בדיוק המקום של „נהל מקורות”
  ב-Word.
- `<w:fldSimple w:instr="CITATION src-…">` וה-`<b:Tag>` שלצידו **זהים**.
  כלומר Word יפתור את הציטוט — ההפך מ-`REF SDXREF` ומ-`TA` בלי `\l`.
- העברית עוברת שלמה בכל השדות: `שו״ת הרמב״ם`, `בן מימון`, `תתקצ״ה`,
  `מוסד הרב קוק`, `ירושלים`.
- הביבליוגרפיה נבנית **מלאה כבר ביצירה**, ו-`rebuild` באמת אוסף מקור
  שנוסף אחריה (`sourceCount` 2 → 3) וגם עריכה של מקור קיים.
- `citations.insert` על `sourceId` שאינו קיים מוחזר `TARGET_NOT_FOUND`,
  ו-`bibliography.rebuild` על מזהה של פסקה רגילה גם הוא. הכתובות מאומתות
  ואינן נבלעות.
- `bibliography.remove` מפיל את הבלוק כולו ואינו משאיר פסקה — ההפך מתוכן
  העניינים.

## מה שכן נכתב נכון — כיתובים

הקבוצה השנייה (אחרי הציטוטים) שעברה את שכבת ה-docx בלי הסתייגות על קוד
השדה:

- הפסקה היא `<w:pStyle w:val="Caption"/>` — סגנון הכיתוב האמיתי של Word —
  ובתוכה `<w:fldSimple w:instr="SEQ איור \* ARABIC"><w:r><w:t>1</w:t></w:r></w:fldSimple>`
  עם התוצאה ה-cached לצידה, בדיוק הצורה ש-Word עצמו כותב.
- **התווית העברית נכנסת אל תוך קוד השדה כמות שהיא.** `SEQ איור`,
  `SEQ טבלה` — לא `Figure`, לא רשימה סגורה, ולא תרגום. `label` הוא
  `string` חופשי, וכל מחרוזת מתקבלת.
- **המספור אמיתי ולפי סדר המסמך.** כיתוב שני באותה תווית מקבל 2, כיתוב
  שנוסף לפניו דוחף אותו ל-3, והסרה מורידה את השאר — גם ב-`list` וגם בערך
  שבתוך `fldSimple`. כל תווית מנהלת רצף משלה.
- **גרשיים ולוכסן מוברחים כהלכה.** `א"ב` → `SEQ "א\"ב" \* ARABIC`,
  ו-`איור \* MERGEFORMAT` → `SEQ "איור \\* MERGEFORMAT"` — כלומר אי אפשר
  להזריק מתג דרך התווית. ההפך מ-`TA`.
- **ההסרה נקייה.** הפסקה כולה יורדת מ-`blocks.list`, ואין שיירים. הסרה
  חוזרת על אותה כתובת, כתובת של פסקה רגילה, ומזהה מומצא — כולם
  `TARGET_NOT_FOUND`.
- **הכתובות ייחודיות.** שני כיתובים זהים תו-בתו מקבלים שני `nodeId`
  שונים — ההפך מתוכן העניינים, שבו ה-hash נגזר מה-`instruction`.

## מה שכן נכתב נכון — הערות שוליים והערות סיום

הקבוצה השלישית (אחרי הציטוטים והכיתובים) שעברה את שכבת ה-docx בלי
הסתייגות על מה שנכתב לקובץ:

- `document.xml` מקבל
  `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>`,
  ו-`footnotes.xml` נבנה שלם: `separator` ו-`continuationSeparator`
  במקומם, וכל הערה `<w:footnote w:id="…"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/>`.
  הערות סיום מקבלות את אותו טיפול ב-`endnotes.xml` עם `EndnoteReference`
  ו-`EndnoteText`. זו הצורה ש-Word עצמו כותב.
- **`footnotes.update` מחליף ואינו מוסיף** — ההפך מ-`captions.update`.
  שלושה צעדים רצופים על אותה הערה החזירו `'ראשונה'` → `'שנייה'` →
  `'שלישית'`, ולא שרשור. `content: ''` **מוחק** את התוכן. כלומר עריכה כאן
  היא קריאה אחת, **בלי** `remove`+`insert` — הדרך שהכיתובים נאלצו ללכת בה
  מפני ש-`captions.update` מוסיף — ולכן גם בלי בחירת עוגן ובלי רשת שחזור.
- **העברית עוברת שלמה**, כולל ניקוד וגרשיים:
  `רַשִׁ״י בְּרֵאשִׁית א׳ א׳, ועיין ב"שו״ת הרמב״ם" סי׳ ק״י` חזר תו-בתו
  מ-`get` ונכתב תו-בתו ל-`<w:t>`.
- **הקלט של `insert` מאומת ואינו נבלע:** `content: 5` נזרק („requires a
  content string”), ו-`type: 'zigzag'` נזרק אף הוא. ירידת שורה בתוכן
  **כן** נכתבת גולמית לתוך `<w:t>`, כמו בכיתובים.
- **`remove` מוריד את ההערה** מ-`footnotes.xml` ומ-`list`, והסרה חוזרת על
  אותה כתובת מוחזרת `TARGET_NOT_FOUND`. מה שהוא כן משאיר בגוף המסמך הוא
  ריצה ריקה בסגנון `FootnoteReference` בלי `<w:footnoteReference>` בתוכה
  — שארית בלתי נראית שאין דרך ציבורית לנקות, ואינה פוגמת בתקינות הקובץ.

## מלכודות מבניות

- **כתובות אינן בהכרח ייחודיות.** שתי טבלאות תוכן עניינים עם אותו
  `instruction` מקבלות את **אותו** `nodeId`, גם ב-`toc.list` וגם
  ב-`blocks.list` — ה-hash נגזר מה-`instruction`. `toc.remove` על הפריט
  השני מוחק את הראשון. במפתח ובטבלת מקורות אין כפילות כזאת.
- **תוכן עניינים אינו בלוק אחד.** הראשון `tableOfContents` והשאר פסקאות
  בסגנון `TOC1…TOC9`; `remove` מוחק את הראשון בלבד ומשאיר את השאר על
  המסך עם `success: true`. במפתח ובטבלת מקורות ההסרה נקייה.
- **אין מיון ואין מספרי עמודים במפתח.** הערכים מופיעים בסדר הסימון, בלי
  כותרות אותיות, למרות `\h "A"`. Word ממיין וממספר בפתיחה.
- **אין דרך למצוא ביבליוגרפיה דרך `blocks`.** ל-`citations.bibliography`
  אין `list`, ו-`blocks.list` מציג את הבלוק כ-`nodeType: 'paragraph'` רגילה.
  הדרך היחידה היא `fields.list`, שמחזיר `fieldType: 'BIBLIOGRAPHY'` ואת
  `address.blockId` — וכתובת שנבנתה ממנו מניעה `get`/`rebuild`/`remove`.
  זה גם מה שמאפשר לעבוד על ביבליוגרפיה שנוצרה ב-Word.
- **`citations.insert` דורש יעד מכווץ.** טווח חוזר `INVALID_TARGET`
  („requires a collapsed text target”), וסמן **בתוך** שדה קיים חוזר
  `CAPABILITY_UNAVAILABLE` („text-range-in-field”).
- **`citations.sources.update` הוא `Partial` אמיתי.** נמדד בשני הכיוונים:
  patch **בלי** `year` השאיר את `תש״ף` שבמסמך כמו שהיה, ו-patch עם
  `year: ''` מחק אותו. כלומר השמטה משמרת ומחרוזת ריקה מוחקת — וטופס עריכה
  שמשמיט שדה שהמשתמש רוקן מייצר „הצלחה מדומה”: `{ok:true}` בלי הודעה,
  והערך חוזר ברענון הבא. מי ששולח patch חייב להחליט לכל שדה מה משמעות
  הריקון אצלו.
- **תצוגת הציטוט אינה מתרעננת אחרי עריכת המקור.** כותרת שהשתנתה מתעדכנת
  בביבליוגרפיה אחרי `rebuild`, אבל הטקסט שבתוך שדה ה-`CITATION` נשאר
  הישן עד `citations.update` על אותו ציטוט. Word מחשב מחדש בפתיחה.
- **מחבר בלי `last` מפיל את המנוע** ב-`TypeError` גולמי ולא בקבלה:
  „Cannot read properties of undefined (reading 'trim')”.
- **אין API להזזת הסמן בין stories.** `doc.selection` הוא קריאה בלבד, ולכן
  אי אפשר להעביר את הסמן אל גוף הכותרת העליונה או אל הכותרת התחתונה.
- **`selection.current` אינו מדווח מקטע**, ואין מיפוי ציבורי סמן→מקטע.
  לכן פעולות מקטע חלות על כל המקטעים.
- **כתובת מקטע היא `section-<index>`, ו-`refStability` שלה `'ephemeral'` —
  ובכל זאת היא יציבה.** `create.sectionBreak` תומך ב-v2 **רק** ב-
  `documentEnd` (`at: {kind:'before', …}` מוחזר `INVALID_TARGET`:
  „supports body documentEnd targets on v2”), ולכן מקטע חדש מקבל תמיד את
  האינדקס הבא ואינו מזיז את הקיימים. נמדד: תצלום של הכתובות שרד מוטציה על
  מקטע אחר וגם הוספת מקטע — `snap[1]` המשיך להיפתר לאותו מקטע עם אותם
  ערכים. כלומר החלון שהיה הופך את הלוגו של „החל על כל המקטעים” ל-TOCTOU
  סגור מצד המנוע, ולא מצידנו.
- **`sections.list` הוא `DiscoveryOutput` עם `limit` של 250, ו-`applyToSections`
  קורא עמוד אחד — פער **קיים**, לא פער של גל 10.** `list()` מחזיר
  `page: { limit: 250, offset: 0, returned: N }` (נמדד), `SectionsListQuery`
  **כן** חושף `offset` (`{limit:1,offset:1}` החזיר את המקטע השני עם `total`
  מלא), ו-`limit: -3` **זורק** ואינו מוחזר ריק כמו ב-`footnotes.list`.
  `applyToSections` שב-`page-setup.ts` קורא `list()` בלי ארגומנטים, **אינו
  משווה `items.length` ל-`total`, ואינו מדפדף** — כלומר במסמך של 251 מקטעים
  ומעלה הפעולה חלה על 250 בלבד והמשתמש מקבל „בוצע”.

  **הסיווג חשוב:** זו התנהגות של `applyToSections` מגל 1, והיא חלה באותה
  מידה על ארבעת הפקדים שקדמו — שוליים, כיוון, גודל נייר ועמודות. חמשת
  הפקדים של גל 10 יורשים אותה ואינם מקורה. **הפער לא נסגר כאן בכוונה**,
  והוכרע לגל נפרד: התיקון משנה את התנהגותו של המודול כולו, ודורש הכרעה
  שאינה טכנית — מה מדווחים כשמקטע נכשל באמצע דפדוף, אחרי שעמודים שלמים
  כבר שונו ואין `rollback`.
- **פסקת כיתוב נכתבת בלי `<w:bidi/>`.** `captions.insert` כותב
  `<w:pPr><w:pStyle w:val="Caption"/></w:pPr>` ותו לא, בעוד שהפסקה הרגילה
  שלצידה במסמך העברי כן נושאת אותו — כלומר הכיתוב ייפתח ב-Word משמאל
  לימין. `paragraphs.setDirection({direction:'rtl'})` מתקן: נמדד שהוא
  מוסיף `<w:bidi/>` לאותה `pPr` ומשאיר את סגנון ה-`Caption` על מקומו
  (`alignmentPolicy: 'matchDirection'` מוסיף גם `<w:jc w:val="right"/>`,
  ולכן אינו נשלח). על פסקה שכבר ימין-לשמאל התשובה היא `NO_OP`.
- **`paragraphs.setDirection` אינו בקטלוג היכולות.**
  `capabilities.get().operations['paragraphs.setDirection']` הוא
  `undefined`, אף שהפעולה עצמה עובדת וכותבת `<w:bidi/>` (נמדד). כלומר
  הקטלוג אינו רשימה מלאה של הפעולות הקיימות, וקוד שנועל פקד על „הפעולה
  בקטלוג” היה מנטרל כאן פעולה תקינה. הבדיקה היחידה שאפשר לסמוך עליה היא
  `typeof doc.paragraphs?.setDirection === 'function'`.
- **„טבלת איורים” אפשרית, אבל ריקה.** `create.tableOfContents` מקבל
  `instruction` גולמי, ו-`TOC \c "איור" \h \z` נכתב קנונית לתוך ה-sdt של
  תוכן העניינים; `toc.list` מחזיר אותו עם
  `preserved.seqFieldIdentifier: 'איור'`. אבל המנוע **אינו** אוסף את
  הכיתובים אליה: `entryCount` נשאר 0 גם אחרי `toc.update`. Word ימלא
  אותה בפתיחה, בדיוק כמו את המפתח. (`toc.configure` אינו יכול להגדיר את
  `\c` — הוא ב-`TocPreservedSwitches`, שמוגדר „round-tripped but not
  configurable”.)
- **`displayNumber` של הערה אינו המספר ש-Word יציג.** הוא זהה ל-`noteId`,
  כלומר לסדר ה**יצירה**: הערה שנוספה בהמשך המסמך וקיבלה 1 נשארת 1, והערה
  שנוספה אחריה במקום מוקדם יותר בטקסט מקבלת 2 — בעוד ש-Word ממספר לפי סדר
  ההופעה (נמדד ב-docx: הרפרנסים יושבים בגוף לפי המיקום). בנוסף, `remove`
  **אינו** ממספר מחדש את השאר: הסרה של 2 השאירה 1 ו-3, גם ב-`list`.
  `footnotes.list` גם מחזיר את הפריטים בסדר היצירה ולא בסדר המסמך, ואין
  בכל ה-API דרך לדעת היכן הערה יושבת. כל ממשק שמציג את המספר הזה חייב
  לומר שהוא סדר היצירה.
- **אימות לפני מוטציה על הערה הוא חלון פתוח.** הכתובת אינה מבדילה בין
  הערת שוליים להערת סיום, ולכן כל עריכה והסרה מאמתות ב-`get` — אבל ה-`get`
  חוצה גבול macrotask, וכל הוספה שנקלטת בזמן הזה משנה את מה שהכתובת
  נפתרת אליו. ההגנה היחידה היא נעילה בצד הממשק בזמן שהפעולה באוויר; ראו
  „כתובת אחת לשתי הערות שונות” למעלה.
- **`footnotes.insert` בלי `at` דורש בחירה חיה.** ב-headless (שם
  `activeEditor.view` הוא `null`) הוא מוחזר
  `PRECONDITION_FAILED / live-selection-unavailable`. זה גם מה שקובע
  שהוספת הערה שייכת לכפתור ברצועה ולא לדיאלוג: מרגע שדיאלוג נפתח, הסמן
  אינו בעורך.
- **`blocks.list` מחזיר מעטפה אחרת מכל שאר ה-discovery**: `blocks` ולא
  `items`. קוד שקורא `items` מקבל `undefined` על מסמך שיש בו פסקה.
- **`doc.insert` פשוט מוסיף לפסקה האחרונה** ואינו יוצר פסקה חדשה, ולכן
  אי אפשר לזרוע שתי פסקאות בשתי קריאות.
- **`activeEditor.view` הוא `null` ב-headless** — אי אפשר למדוד שם שום דבר
  שדורש מיקוד בעורך.

## פעולות שהמנוע מסמן כלא-זמינות

`create.image`, `images.delete`, `images.replaceSource`, `hyperlinks.patch`
— `OPERATION_UNAVAILABLE`. אין לבנות עליהן פקד פעיל.

## `format.paragraph.*` — גל 11, מה שנמדד לפני המימוש

Chrome headless על ה-dist הארוז; כל סבב מלווה בפירוק ה-zip של
`export.toDocx`. ההנמקות ב-engine/paragraph-format.ts, וזה הפער:

- **היחידות הן twips גולמיים.** `setIndentation({left:720})` כתב
  `<w:ind w:left="720"/>`, `setSpacing({before:240})` כתב `w:before="240"` —
  אחד לאחר, בלי המרה. **שונה** מ-`sections.*`, שם ה-API מקבל אינצ'ים וכותב
  `Math.round(v*1440)`. מי שמעביר ערכי UI ישירות כותב שוליים במידות מטר.
- **כל קריאה מחליפה את האלמנט כולו.** `setIndentation({left:-500})` אחרי
  `setIndentation({left,right,firstLine})` השאיר `<w:ind w:left="-500"/>`
  בלבד. אותו דין ל-`w:spacing`. אין patch; כל ממשק חייב לשלוח מצב מלא.
- **מה שהמנוע מאמת וזורק** (`INVALID_INPUT`, ולא קבלה): ערך שאינו מספר
  שלם (`hanging:0.5` → „must be a non-negative integer”), שלילי בריווח
  (`before:-240`), וערכי enum (`lineRule:'zigzag'`, alignment `'zigzag'`
  בטאב). הזריקות מחייבות catch אצל כל קורא.
- **מה שעובר בשקט ונאסר אצלנו:**
  - `setIndentation({left:-500})` → `success:true` ו-`w:left="-500"`.
    חוקי ב-OOXML (`ST_SignedTwipsMeasure`) אך לא מוצע ב-Word — ולא אצלנו.
  - `setTabStop({position:-100})` → `success:true`. `w:pos` שלילי אינו
    חוקי ב-ECMA-376; השער יושב במודול שלנו (מיקום חייב להיות שלם > 0).
- **NO_OP:** קריאה חוזרת עם ערכים זהים מחזירה
  `success:false / code:'NO_OP'` — „produced no changes”. זו הצלחה
  מבחינת המשתמש, כמו בכל המרחבים האחרים.
- **טאבים הם רשימה:** `setTabStop` **מוסיף** ואינו נוגע באחרות (שתי קריאות
  השאירו `<w:tab w:val="center" w:pos="1440" w:leader="dot"/>` ו-
  `<w:tab w:val="right" w:pos="2880"/>` יחד); `clearTabStop({position})`
  מוריד יעד יחיד; `clearAllTabStops` מוריד את `<w:tabs>` כולו.
- **קריאת המצב בנקודות.** `doc.get()` מחזיר SDM/1 שבו ה-indentation,
  ה-spacing וה-tabs הם ב**נקודות** (והטאבים נושאים `kind:'set'|'clear'`) —
  פי 20 מה-API של הכתיבה. הקורא (`readParagraphFormat`) הוא המקום היחיד
  שמכיר את שתי המערכות.
- **keep options:** `keepNext`/`keepLines` נכתבים `<w:keepNext/>`,
  `<w:keepLines/>`; `widowControl:false` נכתב `<w:widowControl w:val="0"/>`.
  סדר הילדים ב-`pPr` יוצא מהמנוע קנוני ועובר round-trip.

## `format.apply` ומשפחת `format.<inlineKey>` — גל 12, מה שנמדד

Chrome headless על ה-dist הארוז; פירוק zip לכל סבב. ההנמקות ב-
engine/font-advanced.ts:

- **היחידות — ה-API בנקודות:** `letterSpacing: 2` → `w:spacing="40"` (×20,
  twips); `position: 3` → `w:position="6"`, `kerning: 12` → `w:kern="24"`,
  `fontSizeCs: 12.5` → `szCs="25"` (×2, חצאי-נקודות — **חצאי נקודות
  מקובלות**, כמו fontSize); `charScale` אחוזים כמות-שהוא.
- **letterSpacing ו-position חתומים:** שלילי עובר (`-20` → `-400`) וזה
  חוקי — „מכווץ"/„מונמך". לעומתם `charScale` ו-`kerning` דורשים שער:
  - `charScale: 9999` → `success:true` ו-`w:w="9999"` — Word תחום
    1..600 אחוז.
  - `kerning: -5` → `success:true` ו-`w:kern="-10"` — ST_HpsMeasure
    אינו חתום.
- **הליבה העברית נכתבת קנונית:** `rtl/cs/bCs/iCs` → `<w:rtl/>` וכו';
  `fontSizeCs` → `szCs`; `lang {bidi:'he-IL'}` → `<w:lang w:bidi=...>`;
  `rFonts {cs:'David'}` → `<w:rFonts w:cs="David"/>`.
- **פער במנוע — `bold` על עברית כותב `w:b` בלבד, בלי `w:bCs`.** Word מציג
  הדגשה של טקסט מורכב מ-`bCs`; ריצה עברית עם `b` לבד אינה תוצג מודגשת.
  פקד ה-bold (דרך הפקודה) וגם `format.bold` מתנהגים כך. עקיפה אצלנו:
  הדיאלוג מציע „מודגש (מורכב)" דרך `bCs`. הפער עצמו מדווח כאן.
- **NO_OP מופיע רק ב-`format.apply`:** ה-alias הבודדים החזירו
  `success:true` גם על חזרה זהה; `apply` מחזיר NO_OP ("produced no
  change") כשה-patch לא משנה דבר.
- **`format.<key>` דורש SelectionTarget** ואינו מקבל TextTarget
  („target must be a SelectionTarget object") — וב-headless `view` הוא
  null ואין בחירה חיה, כלומר היעד מגיע תמיד מהממשק (vert-align.ts).
- **`vanish:false` כותב `w:vanish w:val="0"`** — הסרה מפורשת בדיוק כמו
  Word, ולא הסרת האלמנט; round-trip תקין.
- **`rStyle` לא נשלח** — נוגע בסגנונות תו; גל 13 (סגנונות) הוא בעליו,
  ושני מסלולים לאותה כתיבה הם באג. גם `webHidden` דילג: אין לו משמעות
  ממשק מחוץ לתשתית ההסתרה של Word.

## `styles.*` — גל 13, מה שנמדד

- **`format.paragraph.setStyle/setStyleRef/clearStyle` אינן נתמכות
  בדפדפן** למרות ש-`capabilities.get()` מדווח עליהן `available:true`:
  הריצה מחזירה `CAPABILITY_UNAVAILABLE` ("not a supported v2 browser
  Document API operation"). **הסתירה החדה ביותר במאגר** של הכלל „available
  אינו הוכחה". הפקודה `linked-style` נשארת המסלול היחיד להחלת סגנון על
  תוכן — אין להוסיף מסלול Document API מקביל.
- **`resetDirectFormatting` כן עובד** (success), וכבר מיוצג בפקודה
  „נקה עיצוב".
- **docDefaults (ערוץ run) ביחידות גולמיות:** `patch {fontSize: 14}` →
  `<w:sz w:val="14"/>` = **חצאי-נקודות**, ולא נקודות כמו `format.fontSize`
  (שם 24 → sz 48). `fontFamily` record `{ascii,hAnsi,cs}` נכתב ישירות
  ל-`w:rFonts`. ההמרה pt→×2 יושבת ב-engine/doc-style-defaults.ts.
- **`dryRun` הוא קריאת המצב היחידה:** אין `styles.get`;
  `apply(...,{dryRun:true})` מחזיר `before/after` בלי לשנות. במסמך הריק
  `before.fontSize = 24` (=12pt).
- **חזרה זהה אינה NO_OP אלא `success:true, changed:false`** — שונה מכל
  מרחב אחר; מבחינת המשתמש זו הצלחה.
- **`getCatalog({view:'quickGallery'})` עובד:** 7 פריטים במסמך ריק,
  `sourceStatus.styles:'present'`, שמות כפי שהם במסמך (`heading 1`
  באות קטנה). הגלריה כבר צורכת אותו דרך `ui.styles` מגל קודם — לא
  שוכפלו כאן שני מסלולי קריאה.

## `lists.*` — גל 14א, התשובה לשאלת המספור העברי

- **`setLevelNumberStyle` עם `'hebrew1'` עובד.** `numberStyle` הוא
  **string חופשי** בחוזה (לא union) — ההפך מ-`sections.setPageNumbering`.
  נמדד: `<w:numFmt w:val="hebrew1"/>` נכתב ל-numbering.xml. מספור
  א׳ ב׳ ג׳ אפשרי, ומיושם בתפריט „רשימה" ב„בית". השער שלנו: רשימת
  numFmt תקניים של ECMA-376; מחוץ לה — נדחה.
- **`restartAt({startAt})`** עובד; **`continuePrevious`** מחזיר קבלת
  כשל `INVALID_CONTEXT / NO_PREVIOUS_LIST` כשאין קודם (לא זורק);
  **`canContinuePrevious` בוליאני = TOCTOU** — לא נשלח לפני פעולה,
  הקבלה עצמה מדווחת.
- **`convertToText({includeMarker:true})`** מעתיק את סמן הרשימה
  ('a. ') לתוך הטקסט והפריט הופך לפסקה — בלתי-הפיך למעשה; הפקד דורש
  אישור דו-לחיצה.
- **כתובת פריט:** `{kind:'block', nodeType:'listItem', nodeId}`; פסקה
  רגילה מקבלת `target.nodeType must be 'listItem'`. היעד נפתר אצלנו
  מהבחירה + `blocks.list`.
- **`lists.create mode:'fromParagraphs'`** מקבל BlockAddress בודד או
  BlockRange `{from,to}` — **לא מערך**.

## גלים 17–25 — ממצאי מדידה והכרעות

נמדדו בסבב אחד על ה-dist הארוז (Chrome headless, CDP):

### גל 17 — תמונות: דחייה מנומקת
`images.list` עובד (ריק), אך `create.image` לא-זמין ו-HTML `<img data:>`
נדחה `INVALID_PAYLOAD` — אין דרך לקבל תמונה במסמך, ולכן אין מה לבדוק
ואין מה לבנות. כפי שהמדריך קבע: „הגל הוא תיעוד ולא קוד".

### גל 18 — metadata: חוסך בין החוזה לתיאור
`metadata.*` בחוזה הוא anchored metadata (JSON מוסתר ב-SDT + Custom XML
Storage Part על טווח טקסט) — ולא „מאפייני מסמך" (כותרת/מחבר). נמדד
עובד: attach/list עם payload עברי (`customXml/item1.xml`). אין API
שכותב docProps ב-2.8.0, ולכן לא נבנה UI „מאפיינים" על מצג שווא.
התשתית בעלת ערך עתידי (למשל קישור מקור→ציטוט).

### גל 19 — הגנת מסמך: מיושם
מסלול הביטול נמדד לפני ההפעלה ועובד ללא סיסמה: set(readOnly) →
enforced:true; capabilities.get אחריה: **4 פעולות נפלו ל-false**
(התוסף עצמו מוגבל!); clear → enforced:false. מיושם כמתג עם אישור
דו-לחיצה ב„סקירה" (engine/protection.ts).

### גל 20 — diff: דחייה עד לפתרון ה-host
`diff.capture` עובד (`sd-diff-snapshot/v2`, fingerprint sha256).
`diff.compare` דורש **מסמך שני** — כלומר `fs.pickUserFile` של אוצריא;
בלעדיו אין מסלול משתמש. `diff.apply` הרסני ולא נבדק בנפרד.

### גל 21 — תגובות: דחייה על חוסם זהות
`comments.create` דורש מחבר; לאוצריא אין מודל זהות (`app.getInfo`
אינו מחזיר משתמש) — §13.1 דורש הגדרת משתמש מקומית, שעדיין אינה
קיימת. הפקד „תגובה חדשה" נשאר מנוטרל עם הסבר. `trackChanges.decide`
מתנגש בשש הפקודות הקיימות ולא נשלח; `history.undo/redo` מתנגשות
בפקודות undo/redo ולא נשלחות.

### גל 22 — hyperlinks: מיושם (ראו למעלה)
wrap/remove דורשים TextAddress; wrap דורש מפרט `{link:{destination}}`;
patch לא זמין → עריכה = remove+wrap.

### גל 23 — blocks/create: רובו דחוי — אין פקד Word
`create.paragraph({at:{target,placement:'after'}})` עובד.
`blocks.delete` על פסקה עם `w:bidi` זרק
`paragraph-tracked-wrapper-unsupported` — האי-עקביות שהמדריך ציין
**אושרה**. כל הפעולות הרסניות ואין להן מקבילה ברצועת Word — דילוג
לפי השאלה שהמדריך מציב.

### גל 24 — plan.execute: חוסם חתימת input ל-bookmarks.insert
`plan.execute` רץ אך נכשל בשלב הראשון: "Cannot use 'in' operator to
search for 'story' in undefined" — bookmarks.insert דרך plan דורש
input שונה מה-direct call. נדרש מיפוי מלא של חתימות ה-plan לכל
operationId לפני שאפשר לשלוח מאקרו. mutations.preview/apply לא
נבדקו עד אז.

### גל 25 — contentControls: דילוג לפי ההמלצה
`d.contentControls` ו-`d.customXml` קיימים. 54 פעולות לקהל שאינו
קהל התוסף — ההשקעה גרועה, כפי שהמדריך קובע.

## הסרגל — שלושה ממצאים, ומה נגזר מהם

נמדד ב-Chrome headless על ה-`dist` הארוז, עם קליק אמיתי בתוך הפסקה
(`Input.dispatchMouseEvent`) והקלדה אמיתית (`Input.insertText`), ואחר כך
מדידה של המלבנים שהמנוע צייר בפועל.

### `w:ind` — `left`/`right` ממופים לוגית, `firstLine`/`hanging` לא

בפסקה עברית (`bidi: true`), שוליים של 2.54 ס"מ:

    setIndentation({left: 1440})   → הקצה הימני של הטקסט מ-96px ל-192px  ✔ צד ההתחלה
    setIndentation({right: 1440})  → הקצה השמאלי נכנס פנימה               ✔ צד הסוף
    setIndentation({firstLine: 1440}) → הקצה הימני של השורה הראשונה מ-96px
                                        ל-**0** — כלומר החוצה, אל תוך השוליים
    setIndentation({left:1440, hanging:720}) → השורה הראשונה **עמוק יותר**
                                        מהשאר (240px מול 192px)

כלומר `left`/`right` מתנהגים כמו `w:start`/`w:end` של OOXML — הצד הלוגי —
ואילו שני האחרים מצוירים בסמנטיקה פיזית, כלומר הפוכה. „כניסת שורה ראשונה”
של Word בפסקה עברית אינה ניתנת להשגה: `firstLine` מצייר החוצה, ולערך שלילי
המנוע עונה „must be a non-negative integer”.

**מה נגזר:** הסרגל מציג שני סמני כניסה — התחלה וסוף — ואינו מציג את סמן
השורה הראשונה ואת הסמן התלוי. סמן שגורר ערך שמצויר הפוך גרוע מסמן שאינו
קיים. ראו engine/page-ruler.ts.

### `doc.get()` אינו מחזיר את עצירות הטאב

`format.paragraph.setTabStop({position:2880, alignment:'right', leader:'dot'})`
מחזיר `success: true`, אבל תכונות הפסקה שחוזרות מ-`doc.get()` הן
`{ indent: {...}, spacing: {...}, bidi: true }` — בלי `tabs`, וגם בלי
`keepWithNext`/`keepLines`/`widowControl` אחרי `setKeepOptions` מוצלח.

**מה נגזר:** אין לסרגל דרך לצייר את העצירות הקיימות, ולכן אין בו עצירות
טאב בכלל. סרגל שמראה רק את מה שנוסף בו עצמו, ומעלים את מה שהגיע מקובץ
Word, מטעה יותר משהוא עוזר. אותו ממצא הוא גם הסיבה ש-`readParagraphFormat`
מחזיר `tabs: []` ו-`keepNext: false` על מסמך שיש בו את שניהם.

### מודל הפסקה: `paragraphIds.paraId` ו-`indent`

הצומת שחוזר מ-`doc.get()` הוא

    { kind: 'paragraph', paragraphIds: { paraId: '41964671' },
      paragraph: { inlines: [...], props: { indent: {...}, bidi: true } } }

— **בלי `id`**, ועם `indent` ולא `indentation`. הקוד שחיפש `node.id` ואת
`props.indentation` החזיר אפסים על כל מסמך, ודיאלוג „פסקה” שנפתח עליהם
ואושר מחק כניסות שהגיעו מ-Word (`setIndentation` מחליף את `<w:ind>` כולו).
תוקן בגל הזה; `paraId` הוא בדיוק ה-`blockId` שהבחירה מחזירה.

### מלבן העמוד אינו ניתן לחישוב מבחוץ

זום מיושם ב-`width: 100/zoom%` + `transform: scale(zoom)` עם
`transform-origin: top left`, ולכן העמוד ממורכז בתוך **תיבת הפריסה של
ה-wrapper** ולא בתוך מיכל הגלילה. „רוחב עמוד כפול זום, ממורכז במיכל” נמדד
כשגוי בכל זום שאינו 100% (ב-50% העמוד נמצא ב-‎-625px, והנוסחה נותנת 176px).
אין API ציבורי שמחזיר את המלבן.

**מה נגזר:** הסרגל מודד את המלבן דרך `ui.viewport.getHost()` ותכונת
`data-page-index` שהמנוע מסמן בה עמוד. זו חריגה מתועדת מגבול ה-DOM, והיא
נשמרת צרה בשני שערים: tests/unit/engine-boundaries.test.ts מוודא שרק
engine/page-ruler.ts נוגע בעיגון ושהוא קורא בלבד, ו-
tests/contract/engine-page-hooks.test.ts מוודא שהעיגון עדיין קיים באריזה.

## בחירה בעכבר — מה שנמדד

שלוש התנהגויות של עכבר נמדדו מול המנוע ב-Chrome אמיתי, גם על המנוע לבדו
וגם על התוסף הארוז מ-`file://` עם DOCX עברי מנוקד. שתיים מהן תקינות, ואחת
שבורה בדיוק במקום שכואב לאוצריא.

### לחיצה כפולה — נשברת על ניקוד וטעמים

אותו משפט, בארבע צורות, בלחיצה כפולה על אותה מילה:

| הטקסט | מה שנבחר |
|---|---|
| `שלום עולם גדול` | `עולם` — תקין |
| `שלום, עולם! (גדול)` | `עולם` — תקין |
| `בראשית ברא אלהים` **מנוקד** | `ר` — אות בודדת |
| אותו טקסט **מנוקד ומוטעם** | `ר` — אות בודדת |

כלומר סימני הניקוד והטעמים נספרים אצל המנוע כמפרידי מילה. במסמכי אוצריא
זה אומר שהלחיצה הכפולה כמעט לעולם אינה בוחרת מילה. גם `רמב״ם` נשבר
לפני הגרשיים ומחזיר `רמב`.

**נעקף בשכבה מבחוץ,** `src/engine/word-selection.ts`, בלי לגעת במנוע: הזרע
הוא מה שהמנוע בחר (`doc.selection.current`), החלון סביבו נקרא ב-
`doc.ranges.resolve`, גבול המילה מחושב אצלנו, והבחירה נקבעת ב-
`ui.selection.apply` — שלוש פעולות ציבוריות. נמדד ב-8ms על פסקה בת 3,899
תווים.

שני גבולות של `ranges.resolve` שנמדדו ומכתיבים את המימוש:
`preview.text` נחתך ב-200 תווים ומסמן `truncated`, ולכן החלון הוא ±90 תווים
ולא הפסקה כולה; והיסט שחורג מאורך הבלוק **נחתך** לאורכו, וזה מה שמעיד
שהחלון הגיע לסוף הבלוק.

`query.match` עם regex בתוך הבלוק היה מסלול חלופי שעובד (מחזיר את המילים
עם הטווחים), אך הוא דורש `within` עם `nodeType` — `{kind:'block', nodeId}`
לבדו נדחה ב-`within-node-type-unsupported` — כלומר קריאה נוספת ל-
`getNodeById`, ו-150ms על אותה פסקה. לא נבחר.

### שלוש לחיצות — תקין

בוחר את הפסקה כולה (נמדד: 0..519 על פסקה בת 519 תווים), גם עם שכבת המילה
מותקנת וגם בלעדיה. אין מה לתקן.

### גרירה שיורדת מתחת לשטח הטקסט מאפסת את הבחירה — **פתוח**

גרירה כלפי מטה, בצעדים של 30px, על פסקה אחת ארוכה שכל מילה בה ייחודית:

    y=255  → 881 תווים נבחרו
    y=405  → 1,763
    y=555  → 2,644
    y=705  → 3,636
    y=855  → 91      ← קריסה
    y=1005 → 91      ← ונשאר שם עד השחרור

הבחירה קורסת לטווח קטן שצמוד לעוגן, ונתקעת שם לשארית הגרירה — כלומר
המשתמש משחרר ומגלה שכמעט לא נבחר כלום. נמדד גם בחלון גבוה שבו כל הטקסט
נראה ואין גלילה בכלל, ולכן **גלילה אוטומטית אינה הגורם**.

מה שכן עובד, ונמדד בנפרד: גרירה בתוך שטח הטקסט, גרירה שחורגת מקצה השורה
(ב-RTL — משמאלה), גרירה בין פסקאות, גרירה אל תוך הרצועה ומחוץ לעמוד,
וגרירה מתחת לפסקה האחרונה במסמך קצר. במסמך של 31 פסקאות נפרדות הגרירה
הגיעה עד סוף המסמך ונשארה שם.

**לא נעקף.** תיקון אפשרי דורש לשמור את הבחירה האחרונה שהייתה תקינה ולהחזיר
אותה כשהמנוע קורס — היוריסטיקה שאינה מבדילה בוודאות בין קריסה ובין משתמש
שגורר לאחור, ושמחזירה בחירה עד המקום שבו המנוע עוד תפקד ולא עד המקום שבו
המשתמש שחרר. לפני שמכניסים אותה יש לצמצם את התנאי המדויק לקריסה.
