# פערים שנמדדו במנוע — ומה לא נשלח בגללם

המסמך הזה קיים כדי שאיש לא יחקור שוב את מה שכבר נמדד. כל שורה כאן היא
תוצאה של הרצה ב-Chrome אמיתי מול ה-`dist` הארוז, ולא קריאה של טיפוסים.

**הכלל שנגזר מכל אלה:** `available: true` בקטלוג היכולות ו-`success: true`
בקבלה **אינם הוכחה שהפעולה עובדת**. פעולה שכותבת קוד שדה של Word חייבת
אימות מול תיעוד Word, ורצוי גם מול ה-docx המיוצא.

המתכון למדידה נמצא בסוף `docs/spike.md` ובקוד: `scripts/cdp.mjs`.

## מסמך שמקפיא את המנוע לצמיתות

### `w:defaultTabStop` של אפס — לולאה שאין ממנה יציאה

`<w:defaultTabStop w:val="0"/>` ב-`word/settings.xml` מקפיא את החוט הראשי
לתמיד: לא `onReady`, לא `onException`, ו-100% מעבד. נמדד מעל עשר דקות לפני
שוויתרנו. נמדד ב-superdoc 2.8.0 (engine 0.7.0) **וגם** ב-2.10.0 (engine 0.9.0),
ובאותה מידה מ-`file://` ומ-`http://127.0.0.1` עם workers אמיתיים.

זה הפער היחיד כאן שאי אפשר להתגונן מפניו אחרי שהמסמך נמסר: `OPEN_TIMEOUT_MS`
צריך בדיוק את החוט שחסום, ולכן אינו יורה. שום שעון-שמירה אינו יכול לתפוס את
זה. לכן ההגנה היא **לפני** המסירה — `src/engine/docx-preflight.ts`.

מה שנמדד בבידוד, כדי שלא ייחקר שוב:

| מה נבדק | תוצאה |
|---|---|
| הקובץ שדווח, כמות שהוא | קופא |
| אותו קובץ, `0` הוחלף ב-`720`, `document.xml` זהה בייט-בבייט | נפתח |
| מסמך Word ריק ותקין, `defaultTabStop` הוכרח ל-`0` | קופא |
| אותו קובץ עם גוף שכולו `<w:p/>` ריקה | נפתח |
| אותו קובץ עם פסקה אחת שיש בה טקסט | קופא |

השורות האחרונות הן מה שקובע את היקף ההגנה: אין צורך ב-`<w:tab/>` ואין צורך
בעצירות טאב מפורשות. די בכך שיש טקסט לפרוס.

דווח למעלה: https://github.com/superdoc/docx-editor/issues/3944

**וההגנה עצמה הייתה חלקית, ואינה עוד.** הרגקס שכתב את הערך
(`/<w:defaultTabStop\b[^>]*\/>/`) לא הותאם בשלוש צורות חוקיות לגמרי, ובכל אחת
מהן הקיפאון פשוט לא נמנע: הערה שיש בה ערך תקין **לפני** האלמנט החי (ההתאמה
הראשונה היא זו שבהערה), אלמנט שאינו סוגר את עצמו
(`<w:defaultTabStop w:val="0"></w:defaultTabStop>`), ו-`>` בתוך ערך של מאפיין
אחר. בכיוון ההפוך, הערך **בתוך** הערה תוקן. שלושתן מכוסות עכשיו, וההגנה עוברת
דרך הסורק המודע-למרכאות של הסעיף הבא. אין שער CDP על המסלול הזה — רק בדיקות
יחידה — מפני שמסמך שמפעיל אותו מקפיא את הדפדפן ולא נותן לשער דרך לדווח.

## מסמך עברי שכותרותיו מודגשות ב-Word ומגיעות אלינו דקות

### `w:bCs` אינו מרונדר — הדגשת כתב מורכב נבלעת

`<w:bCs/>` היא „מודגש לכתב מורכב”, וזה מה ש-Word העברי כותב כשמדגישים בחירה
שכולה עברית: `w:bCs` לבדה, בלי `w:b`. Word מרנדר הדגשה של ריצה עברית מ-`bCs`;
המנוע מרנדר אותה מ-`w:b` **בלבד**.

נמדד ב-superdoc 2.10.0 (engine 0.9.0) **וגם** ב-2.11.0 (engine 0.10.0), Chrome
אמיתי על ה-dist הארוז, על הקובץ שדווח (סגנונות „כותרת 2/3/4/5” נושאים
`<w:bCs/>` ואין בהם `<w:b/>`):

| מה נבדק | `font-weight` שנצבע |
|---|---|
| הקובץ שדווח, כמות שהוא | 400 בכל עשר הכותרות |
| אותו קובץ, `<w:b/>` הושלם לצד כל `<w:bCs/>` | 700 בכולן |
| `bCs` על **הריצה** (עיצוב ישיר), בלי `b` | 400 |
| `bCs` על **`w:rPrDefault`** ב-`styles.xml`, בלי `b` | 400 |

זה **הכיוון ההפוך** לפער שמתועד למטה על `format.apply` (שם המנוע כותב `w:b`
בלי `w:bCs`, ולכן Word אינו מציג את מה שאנחנו מציגים). שני הפערים הם אותה
עובדה אחת: המנוע מכיר `b` ואינו מכיר `bCs`.

**מה נעשה כאן:** ההשלמה נעשית ב-`src/engine/docx-preflight.ts`, לפני שהמנוע רואה את
הבייטים. לא בגיליון סגנונות, ומשתי סיבות שנמדדו: ה-DOM שהמנוע מצייר אינו נושא
שום סימן ל-`bCs` (`.superdoc-text-run` מקבל `styleid` ו-`font-weight` מחושב
ב-`style`, ולא את מאפייני הריצה), והמשקל נכנס לחישוב שבירת השורות — הדגשה
שנצבעת אחרי הפריסה מזיזה טקסט מתחת לפריסה שכבר חושבה.

השער: `scripts/qa/bold-cs-qa.mjs`. הוא אינו ריק — עם התיקון מושבת בבאנדל
נופלות שתי שורות המסך שנבנות מהייצוא (`bCs` על הריצה, `bCs` על ברירות המחדל)
וגם „הודעה למשתמש”, ובקרה של „טקסט בלי `bCs` נשאר דק” עוברת בשני המצבים.
השורה של הקובץ שדווח נמדדת רק כשנמסר לו נתיב — ב-`verify:qa` היא „לא נבדק”,
מפני שהקובץ אינו במאגר.

`<w:iCs/>` (נטוי לכתב מורכב) הוא אותו מנגנון בדיוק, וכך גם `w:szCs` ו-
`w:rFonts/@cs`: המנוע מונה את ארבעתם כ„מחסנית הכתב המורכב” ואינו מיישם אף אחד
מהם. כאן תוקנה הדגשה בלבד — זה מה שדווח.

**מה השלב המקדים אינו מכסה, ובמפורש:** הוא רץ **בפתיחה** בלבד. „מודגש (מורכב)”
בדיאלוג הגופן המתקדם כותב `bCs` על ריצה חיה (ראו `engine/font-advanced.ts`,
והסעיף על `format.apply` למטה), וההדגשה הזאת לא תיראה על המסך עד שהקובץ ייפתח
מחדש. זו נגזרת של אותה עובדה שנמדדה — המנוע מרנדר מ-`w:b` בלבד — ולא מדידה
נפרדת. מה שסוגר את זה הוא התיקון במעלה הזרם, שיושב בשכבת הרינדור ולכן חל גם
על עריכה חיה.

### התיקון במעלה הזרם

**הפער הוא בקוד ה-OSS של המנוע, ולא בחלק הקנייני.** הפירוק של `w:bCs` אכן
קורה — `RunProperties.boldCs` מגיע מלא — אבל
`packages/layout-engine/style-engine/src/normalize/run-attrs.ts` (הפונקציה
`normalizeRunAttrsFromOoxml`) קורא רק את הווריאנטים הלטיניים, ומשם והלאה חוזה
הפריסה נושא `bold` בודד. אימות שהפונקציה הזאת אכן על מסלול הרינדור: הצורה
המהודרת שלה נמצאת בתוך `@superdoc/docx-engine/dist/docx-engine.es.js` שורה
אחר שורה.

המנוע עצמו מתעד את זה כעבודה שלא נעשתה —
`layout-engine/contracts/src/direction-context.ts`: „preservation-only in Wave
1a. **Wave 1b** implements the stack-selection logic (`resolveRunScriptContext`
…)”. `resolveRunScriptContext` אינו ממומש בשום קובץ — השם מופיע רק בהערה
הזאת (גם ב-`.d.ts` שבחבילה), ולא בשום `.js` — וגם לא ב-0.10.0 שיצא מאז.

דווח למעלה: https://github.com/superdoc/docx-editor/issues/3959
תיקון נשלח: https://github.com/superdoc/docx-editor/pull/3958 — בחירת מחסנית
לפי `w:rtl`/`w:cs` עם נפילה חזרה לדגל הלטיני, ובדיקות.

**ושני התיקונים אינם כפילות — נמדד.** התיקון של המנוע הוחל על הבאנדל הארוז
בצורתו המהודרת, עם התיקון המקומי **מושבת**, ואותו שער רץ עליו:

| מה נמדד | המנוע לבדו | התיקון המקומי לבדו |
|---|---|---|
| הקובץ שדווח (ריצות עם `<w:rtl/>`) | 10/10 מודגשות | 10/10 מודגשות |
| `bCs` על ריצה **בלי** `w:rtl`/`w:cs` | 400 — לא מודגש | 700 |
| `bCs` על `w:rPrDefault`, בלי `w:rtl` | 400 — לא מודגש | 700 |

וזה בדיוק הגבול התיאורטי: לפי Annex I, בהיעדר `w:rtl` ו-`w:cs` המחסנית נבחרת
לפי **הכתב של הטקסט**, וזה מידע שאינו מגיע לשכבה שבה התיקון של המנוע יושב. Word
העברי כותב `w:rtl` על ריצות עבריות, ולכן המנוע יכסה את המסמכים המציאותיים —
והתיקון כאן נשאר הרחב מהשניים.

### מה מסיר את השלב המקדים, ומתי

השלב המקדים **משנה את המסמך** (`w:b` שהוזרק נשמר לקובץ), ולכן הוא אינו אמור
להישאר מטעמי הרגל אחרי שהמנוע יתוקן. תנאי ההסרה, שניהם יחד:

1. superdoc/docx-editor#3958 (או שקול לו) נכלל בגרסת `@superdoc/docx-engine`
   ש-`package.json` מצביע עליה.
2. `scripts/qa/bold-cs-qa.mjs` רץ עם `repairComplexScriptBold` **מושבת**, ושלוש
   שורות המסך עוברות: `bCs` על הריצה, `bCs` על ברירות המחדל, והקובץ שדווח —
   שדורש להעביר לשער את הנתיב שלו, כי ב-`verify:qa` הוא מדולג.

ומה שנשאר פתוח גם אז, ומוכרע רק כשמגיעים לשם: שתי השורות של „בלי `w:rtl`”
בטבלה שלמעלה יישארו דקות במנוע גם אחרי #3958, וזו בחירה בין השארת ההזרקה
כרשת רחבה יותר (עם המחיר של שינוי המסמך) לבין הסתפקות בהתנהגות של Word — שגם
הוא, על ריצה בלי `w:rtl`/`w:cs`, מכריע לפי הכתב ולא לפי הדגל.

עד אז אין שער גרסה ואין מתג: השלב רץ על כל מסמך, ומסמכים שנשמרו בינתיים
נושאים את `w:b` לתמיד. זה מקובל בכוונה — `w:b` לצד `w:bCs` הוא מה ש-Word עצמו
כותב כשמדגישים בחירה מעורבת, ואינו מסמך „מקולקל”.

**זיכרון בפתיחה.** החלק המתוקן נכתב דחוס (`CompressionStream('deflate-raw')`,
מאומת בפריסה חוזרת והשוואת CRC, ונופל ל-`STORED` כשהדחיסה חסרה או נכשלת).
הגרסה הראשונה כתבה אותו גלוי, ועל 81 מסמכים אמיתיים זה היה פי 5.2 בסך הכול
ופי 15.3 בגרוע — 35.9MB למנוע מקובץ של 4.3MB. הפרטים בכותרת
`src/engine/docx-preflight.ts`.

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
                                            ← נכון ל-2.8.0 בלבד. ב-2.10.0 ה-union
                                              כולל hebrew1/hebrew2, והקריאה כותבת
                                              <w:pgNumType w:fmt="hebrew1"/>.
                                              ר' docs/superdoc-2.10-review.md
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

### מספור עמודים עברי — היה בלתי אפשרי, ונסגר ב-2.10.0

הפער נמדד על 2.8.0: ה-union **נאכף בזמן ריצה**, ולא הייתה דרך ציבורית לכתוב
`<w:pgNumType w:fmt="hebrew1"/>` — `SectionPageNumberingFormat` היה שש
אפשרויות לטיניות בלבד. זה ההבדל שהיה אז מול `footnotes.configure`, שבו
`numFmt` נכתב גולמית.

**ב-superdoc@2.10.0 ה-union כולל `hebrew1` ו-`hebrew2`**, ואותה קריאה בדיוק
נמדדה כמצליחה: `<w:pgNumType w:start="1" w:fmt="hebrew1"/>` ב-docx המיוצא.
שני הפורמטים מוצעים בדיאלוג, נשלחים, וחוזרים מ-`sections.list` — ר'
`PAGE_NUMBER_FORMATS` ב-`src/engine/page-setup.ts` ואת שתי הבדיקות
(`tests/unit/page-setup.test.ts`, `tests/component/layout-group.test.ts`).

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

## הסרגל האנכי — שני ממצאים נוספים

נמדד באותה שיטה: ה-`dist` הארוז ב-Chrome, קליק והקלדה אמיתיים, ואז סריקת
ערכים דרך `sections.setPageMargins` ומדידת מה שצויר.

### שולי הטקסט מורמים בשקט כשיש כותרת עליונה

A4, כותרת עליונה **ריקה**, `headerDistance` של חצי אינץ':

| `top` שנשלח | ראש הטקסט שצויר |
|---|---|
| 0.75" | 72px |
| 0.7"  | 67.2px |
| 0.6"  | 66.4px |
| 0.5"  | 66.4px |
| 0.25" | 66.4px |
| 0"    | 66.4px |

כלומר מתחת ל-66.4px הטקסט **אינו זז**, וכל גרירה נוספת אינה עושה דבר.
66.4 = 48 (מרחק הכותרת) + 18.4 (גובה שורה ריקה אחת). עם `headerDistance`
של אינץ' הרצפה עוברת ל-114.4px — כלומר הנוסחה היא
`max(topMargin, headerDistance + גובה הכותרת)`, וזו גם הנוסחה שמופיעה
במקור של מנוע הפריסה (`layout-engine/src/section-breaks.d.ts`, ובאריזה
`X=(io,ia)=>O>0?Math.max(ia,io+O):ia`). ההתנהגות עצמה תואמת ל-Word.

**מה שאינו תואם ל-Word** הוא שקריאת המסמך אינה מגלה את זה: `sections.list()`
מחזיר את `w:top` כפי שנכתב, ולכן פקד שנשען עליו מצייר שוליים שאין מתחתיהם
טקסט. מה שכן מגלה: **`activeEditor.pageMetrics.getSnapshot()`**, שמחזיר
ב-`pages[0].base.marginTopPx` את הערך **אחרי** ההרמה (66.4 גם כשבמסמך כתוב 0).

**מה נגזר:** הסרגל האנכי מצייר ומגביל לפי `pageMetrics` ולא לפי המסמך, והידית
נעצרת על הרצפה עם הסבר ב-`title`. המימוש ב-`readEffectiveMargins`
(engine/page-setup.ts), והשער באריזה ב-tests/contract/engine-page-hooks.test.ts —
`pageMetrics` אינו מופיע ב-`.d.ts` של `superdoc`, בדיוק כמו `data-page-index`.

### שוליים בלי מקום לטקסט מוחקים את המסמך מהמסך — **בלי דרך חזרה**

A4 (גובה 1122.53px), שוליים תחתונים של אינץ':

| `top` | גובה הטקסט שנשאר | מה שצויר |
|---|---|---|
| 10.29" | 38.7px | 8 עמודים |
| 10.60" | 8.9px  | 4 עמודים |
| 10.68" | 1.3px  | 4 עמודים |
| **10.70"** | **‎-0.7px** | **אפס עמודים** |

ומרגע הקריסה זה נשאר כך: `setPageMargins({top: 1})` מחזיר `success: true`,
`setPageMargins({bottom: 0.9})` אחריו גם — והמסך נשאר ריק, `pageMetrics` קפוא
על התצלום האחרון, ו-`[data-page-index]` אינו קיים. כלומר המסמך נעלם ואינו חוזר
עד טעינה מחדש. אין קבלת כשל, אין אזהרה, ואין הודעה בקונסולה.

**מה נגזר:** `applyPageMargins` דוחה שוליים שאינם משאירים חצי אינץ' לטקסט,
לפני שהוא נוגע במנוע — בשכבת המנוע ולא בפקד, כדי שכל כותב עתידי יעבור דרך
החסם. הסרגל מגביל את הגרירה לאותו ערך. חצי אינץ' ולא ס"מ: ס"מ (37.8px) עבר
במדידה, אבל מרווח של פחות מפיקסל מקריסה בלתי הפיכה אינו חסם אלא צירוף מקרים.
מסמך שכבר מגיע חנוק אינו ננעל — שינוי שמשפר את המצב מותר, אחרת החסם היה מונע
דווקא את התיקון.

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

### שלוש לחיצות — תקין במנוע, ובכל זאת נשלח גם מאיתנו

במנוע לבדו שלוש לחיצות בוחרות את הפסקה כולה (נמדד: 0..519 על פסקה בת 519
תווים), גם עם שכבת המילה מותקנת וגם בלעדיה. אף על פי כן `installWordSelection`
סופר לחיצות בעצמו ושולח בחירת פסקה בלחיצה השלישית — ראו הסעיף הבא.

### ספירת הלחיצות אינה בידי הדף, והיא שונה בין הסביבות

זה הפער היחיד כאן שאינו של מנוע ה-DOCX אלא של המארח, והוא מסביר למה בדיקה
ב-Chrome אינה עדות למה שקורה באוצריא.

אוצריא ב-Windows מריצה את התוסף ב-WebView2 במצב **visual hosting**, ושם כל
אירוע עכבר מועבר ידנית מ-Flutter דרך `SendMouseInput`. מה שנשלח בפועל
(`flutter_inappwebview_windows`, הפורק של אוצריא, `resolved-ref c0c43c42d`):

    MOVE · LEFT/RIGHT/MIDDLE_BUTTON_DOWN · ..._UP · LEAVE · WHEEL

`COREWEBVIEW2_MOUSE_EVENT_KIND_LEFT_BUTTON_DOUBLE_CLICK` **אינו קיים בחבילה
כלל** — לא ב-C++ (`windows/in_app_webview/in_app_webview.cpp`,
`setPointerButtonState`) ולא ב-Dart (`InAppWebViewPointerEventKind` מכיל
`activate, down, enter, leave, up, update, cancel`). ערכי ה-enum של WebView2
הם הודעות Win32, ו-`DOUBLE_CLICK` הוא `WM_LBUTTONDBLCLK`; מה שמציל את המצב
הוא ש-Chromium מחשב `clickCount` בעצמו מזמן וממרחק, ולכן לחיצה כפולה כן
מגיעה לדף. אבל זה חישוב של המארח, לא חוזה של הדף.

**מה נגזר:** השכבה אינה מאזינה ל-`dblclick` אלא סופרת `mousedown`/`click`
בעצמה, באותם ספים (500ms — ברירת המחדל של Windows — ו-5px). כשהמנוע צודק
התוצאה זהה ולא נשלחת בחירה שנייה; כשהוא לא, התוצאה עדיין נכונה.

### גרירה שיורדת מתחת לשורה האחרונה מאפסת את הבחירה — **פתוח, והגורם נמדד**

גרירה כלפי מטה, בצעדים של 30px, על פסקה אחת ארוכה (4,199 תווים, מילים
ייחודיות) בחלון 1400x1200. בכל צעד נדגמו ארבעה דברים באותה נקודה: מה
`document.elementFromPoint` מחזיר, לאן `document.caretRangeFromPoint` ממפה,
מה `getSelection()` מחזיק, ומה `doc.selection.current()` מדווח.

    y=284..1004  → 0 → 3,960 תווים, גדל ב-180 לכל 30px   תחת הסמן: SPAN.superdoc-text-run
    y=1034       → 4,139                                  תחת הסמן: DIV.superdoc-line
    y=1064       → 1        ← קריסה                       תחת הסמן: DIV.superdoc-page
    y=1094..1154 → 0–1                                    תחת הסמן: DIV.superdoc-page
    y=1184       → 1                                      תחת הסמן: FOOTER (מחוץ לעמוד)

**הגורם:** הקריסה מתרחשת בדיוק בצעד שבו הסמן חדל להיות מעל תיבת שורה. עמוד
הוא מלבן A4 בגודל קבוע, והטקסט כמעט אף פעם אינו ממלא אותו — לכן בין השורה
האחרונה לתחתית העמוד יש רצועה מתה שהיא `DIV.superdoc-page` ותו לא. מיפוי
הקואורדינטה למיקום במסמך של המנוע אינו מהדק (clamp) את הנקודה לשורה הקרובה
ברצועה הזאת, ולכן הגרירה מקבלת מיקום ליד תחילת הבלוק והבחירה מתאפסת.

זו אינה תקלה של הדפדפן: `document.caretRangeFromPoint` באותן קואורדינטות
בדיוק המשיך להחזיר את המיקום הנכון והמהודק — צומת הטקסט האחרון, היסט 59 —
בכל אחד מהצעדים שאחרי הקריסה. המנוע גם אינו מניח את הבחירה על ה-DOM כלל:
`getSelection()` היה ריק לאורך כל הגרירה, כלומר הוא מצייר שכבת בחירה משלו
ומחזיק את המיקום בעצמו, ולכן אינו יורש את ההידוק של הדפדפן בחינם.

**בקרה שמאשרת:** אותה מדידה בדיוק, עם פסקה קצרה שנוספה אחרי הארוכה. הקריסה
זזה יחד עם הטקסט — ב-y=1064 וב-y=1094 הבחירה החזיקה (4,175 תווים) מפני
שהסמן היה מעל שורות הפסקה השנייה, והקריסה הופיעה רק ב-y=1124, מתחת לשורה
האחרונה שלה. זה גם מסביר למה במסמך של 31 פסקאות קצרות הגרירה הגיעה עד סוף
המסמך: הטקסט מילא את העמוד, ולא נותרה רצועה מתה.

**ו-„91" מהמדידה הראשונה אינו קבוע קסם:** הבחירה גדלה ב-180 תווים לכל 30px
בגובה שורה של 17px, כלומר שורה בקובץ הזה היא כ-90 תווים. מה שנשאר אחרי
הקריסה הוא שווה-ערך לשורה אחת.

**נקודה שכדאי לדעת:** ברגע שהסמן יוצא מהעמוד לגמרי (y≥1184, מעל שורת המצב)
המנוע מפסיק לעדכן והבחירה נשארת בערך האחרון התקין. כלומר גרירה *רחוקה*
מטה בטוחה; מה שהורס היא דווקא הרצועה המתה שבתוך העמוד.

**וגורם שני, במארח, שאינו זהה לזה:** `custom_platform_view.dart` שולח
`LEAVE` מתוך `MouseRegion.onExit` **בלי לבדוק אם כפתור לחוץ**, ועם
`virtualKeys = NONE`. כלומר ברגע שהסמן יוצא מגבולות ה-WebView באמצע גרירה,
Chromium מקבל „העכבר עזב, בלי כפתורים" ומסיים את בחירת הגרירה. זה מסלול נפרד
מהקריסה שנמדדה למעלה. תוקן ב-Otzaria/flutter_inappwebview#11.

**נעקף ב-`engine/pointer-snap.ts`, ולא בחסימה אלא בהצמדה.** התנאי המדויק —
„הסמן בתוך העמוד ולא מעל `.superdoc-line`" — הוא אותו תנאי בדיוק של הסעיף
הבא (לחיצה מחוץ לגליפים), והעקיפה משותפת: בזמן שהכפתור לחוץ, כל אירוע תנועה
ושחרור מקבל את הקואורדינטות של הנקודה הקרובה על שורת הגוף הקרובה, ולכן
הגרירה ברצועה המתה **ממשיכה עד סוף הטקסט** כמו ב-Word, במקום להיעצר על
הערך האחרון. מה שנמדד לפני כן ונדרש כאן: המנוע מאזין ל-`pointerdown`
ו-`mousedown` על ה-host שלו, ול-`pointermove`/`pointerup` על `window` ב-capture
שהוא רושם ברגע הלחיצה; הוא קורא `clientX`/`clientY` מהאירוע, ולכן מאזין
שנרשם על `window` **לפני בניית המופע** ומחליף אותם על האירוע עצמו קודם לו.

### לחיצה מחוץ לגליפים — הסמן קופץ לתחילת הטקסט — **נעקף**

נמדד ב-`scripts/qa/click-snap-qa.mjs` (superdoc 2.11.0, docx-engine 0.10.0):
מסמך של שלוש פסקאות — עברית ארוכה, ריקה, עברית קצרה — ולחיצה (press, 40ms,
release) בנקודות שאינן על גליף:

| הלחיצה | המנוע | Word |
|---|---|---|
| משמאל לסוף השורה העברית, על השורה | היסט 0 — **תחילת** השורה | סוף השורה |
| בשוליים הימניים, ליד השורה | היסט 0 | תחילת השורה (במקרה זהה) |
| מתחת לשורה האחרונה, בכל x | **הפסקה הראשונה**, היסט 0 | סוף השורה האחרונה |
| בשוליים העליונים | הפסקה הראשונה, היסט 0 | השורה הראשונה, באותו x |

**הגורם**, בקוד המנוע (`clickToPositionDom`): המיפוי עובר ב-`elementsFromPoint`.
כשנמצאה שורה (`.superdoc-line` — והיא רחבה כרוחב הטור, לא כרוחב הטקסט),
ה-x מומר להיסט; מחוץ לגליפים ההמרה מחזירה 0 במקום להדק לקצה. כשלא נמצא
פרגמנט בכלל (הרצועה המתה, השוליים), הנפילה-לאחור היא **הפרגמנט הראשון של
העמוד** — לא הקרוב. אותו מסלול בדיוק הוא שמאפס את הגרירה שבסעיף הקודם.

**העקיפה** — `engine/pointer-snap.ts`, והגיאומטריה ב-`measurePageGlyphs` של
`engine/page-ruler.ts`: מאזין על `window` ב-capture, שנרשם לפני בניית המופע,
מחליף על האירוע עצמו (`Object.defineProperty` על `clientX`/`clientY`, שמאפיל
על ה-getter של `MouseEvent.prototype`) את הנקודה בנקודה הקרובה על השורה
הקרובה — לפי מלבני הגליפים של צומתי הטקסט, בהיקף שנגזר ממטרת האירוע (השורה,
התא, או עמוד שלם ברצועה המתה, ואז שורות הגוף בלבד). המנוע, וגם
`word-selection.ts` ו-`format-painter.ts` שקוראים אחריו, רואים את אותה נקודה;
אין אירוע מסונתז ואין `stopPropagation`. תמונות וידיות אינן מוצמדות — לחיצה
עליהן היא בחירת אובייקט.

**מה שלא נמדד:** לחיצה כפולה באזור כותרת עליונה שפתוחה לעריכה. ההצמדה אינה
נוגעת ב-`dblclick`, והמנוע נכנס לכותרת מהאירוע הזה בקואורדינטות המקוריות.

### הלחצן הימני — המנוע מתעלם ממנו

נמדד ב-`scripts/context-menu-probe.mjs`, Chrome אמיתי על ה-dist הארוז, מסמך של
24 פסקאות.

| מה נבדק | התוצאה |
|---|---|
| `contextmenu` נורה בדף | כן — גם ב-capture וגם ב-bubble, ואינו מגיע מבוטל |
| לחיצה ימנית **בתוך** בחירה קיימת | הבחירה נשמרת בשלמותה |
| לחיצה ימנית **מחוץ** לבחירה | **הסמן אינו זז** |
| בקרה: לחיצה שמאלית באותה נקודה | הסמן זז (בלוק אחר, סמן מכווץ) |
| `preventDefault` על `pointerdown` ימני | `contextmenu` עדיין נורה |

הבקרה היא מה שהופך את זה לממצא. במדידה הראשונה גם הלחיצה השמאלית לא הזיזה —
מפני שהנקודה שנבחרה כלל לא הייתה על טקסט. בלי בקרה הייתה נרשמת כאן מסקנה על
המנוע שהיא בעצם באג במדידה.

**מה שמאפשר לעקוף מבחוץ**, ושניהם נמדדו באותה ריצה:

- `ui.selection` חושף `getRects()` ו-`getAnchorRect()` — המלבנים שהבחירה
  **מצוירת** בהם. כלומר „האם הנקודה שנלחצה בתוך הבחירה” נענית בלי שום hit-test
  על ה-DOM של המנוע, וגם ל-`Shift+F10` יש עוגן אמיתי.
- לחיצה שמאלית **מסונתזת** (`pointerdown`/`mousedown`/`pointerup`/`mouseup`/
  `click` על `elementFromPoint`) כן מזיזה את הסמן.

`composables/use-context-menu.ts` מרכיב את השניים להתנהגות של Word: לחיצה ימנית
בתוך הבחירה משאירה אותה, ומחוצה לה מזיזה את הסמן. הלחיצה המסונתזת מסומנת
ב-`markSyntheticPointer` כדי שמונה הלחיצות של `word-selection.ts` לא יספור
אותה — אחרת שתי לחיצות ימניות באותה נקודה היו נחשבות ללחיצה כפולה.

**שני דברים נוספים שנמדדו שם, ואינם על הלחצן הימני:**

- **הבחירה שורדת מעבר מיקוד** לפקד של הממשק. זה מה שמתיר לתפריט ההקשר לקבל
  מיקוד — כלומר ניווט מקלדת אמיתי — במקום להעתיק את `@pointerdown.prevent`
  של הרצועה, שהיה הופך אותו לתפריט שאי אפשר לנווט בו.
- **`story` הוא אובייקט קריא**: `{"kind":"story","storyType":"body"}`.
- **`hyperlinks.list` מחזיר מיקום** — `paraId`, `hyperlinkNodeId`, `address`,
  `text`, `externalTarget`, `targetKind`. העטיפה ב-`hyperlinks-manage.ts`
  משטחת את זה ל-`{id,href,anchor}`; זיהוי „קישור תחת הסמן” אפשרי, והוא דורש
  להרחיב את העטיפה ולא לעקוף אותה.

**מה שאינו נמדד כאן, ובמכוון:** האם `preventDefault` מבטל את תפריט ברירת
המחדל של WebView2, והאם `contextmenu` מסונתז שם בכלל. השער רץ ב-Chrome
headless, שאין בו תפריט נייטיב להדגים עליו; אוצריא מעבירה כל אירוע עכבר ידנית
דרך `SendMouseInput`, ולכן זה נבדק באוצריא עצמה עם DevTools.

## שני טורים במקטע עברי — סדר המילוי מתעלם מ-`w:bidi`

ECMA-376 §17.6.1 קובע ש-`w:bidi` ב-`sectPr` הוא שמכריע את **הכרום של המקטע**:
מספרי עמודים, שוליים פנימיים, ו**סדר הטורים**. כלומר במקטע עברי הטור הראשון
שייך לצד **ימין**. `w:bidi` של הפסקה (§17.3.1.6) הוא ציר נפרד ואינו קובע כאן.

המנוע ממלא את הטורים **שמאל→ימין** בכל מקרה. נמדד ב-Chrome אמיתי מול ה-`dist`
הארוז, `scripts/qa/column-selection-probe.mjs`:

| מה נמדד | תוצאה |
|---|---|
| „עמודות ← שתיים” מודיעה למשתמש | **עובד** — ההודעה מגיעה לשורת המצב דרך התפריט האמיתי |
| הייצוא — `sectPr` אחרי „עמודות ← שתיים” | **נכון**: `w:bidi` קיים, ולצדו `<w:cols w:space="720" w:num="2" w:equalWidth="1"/>` |
| שורה 01 של המסמך | נוחתת בטור ה**שמאלי** (x=538, אמצע הדף 700) — בוורד היא מתחילה מימין |
| גרירה בסדר המסמך (שמאל ואז ימינה) | **עובד** — רציפה, וחוצה את הגבול; נבחרו 6..30 (25 שורות) |
| גרירה בסדר ה**קריאה** (ימין ואז שמאלה) | הבחירה **מתהפכת סביב העוגן** ואינה נמחקת: `contiguous: true`; הטור הימני הגיע ל-8 מלבנים וירד ל-5 אחרי המעבר, ובסוף 20..30 (11) במקום להמשיך מ-30 |
| סמן ברוחב הטור הימני מתחת לשורה האחרונה שבו | הראש אינו מהודק לסוף הטור — 6 מלבנים משמאל ו-5 מימין |
| `Shift+חץ מטה` בתחתית הטור | לא רק שאינו חוצה — הוא **גולש לתחילת המסמך**: סמן שאומת בהקשה בודדת בשורה 23, ושש הקשות **מטה** נתנו 4..23, כלומר הראש עלה 19 שורות אחורה במקום להגיע לראש הטור השני |

שלוש השורות שעובדות אינן קוסמטיקה בדוח: בלעדיהן אי אפשר לדעת אם „הבחירה
שבורה” או „סדר הטורים הפוך”, וזה בדיוק ההבדל שקובע איפה התיקון.

**זה מסביר את הבאג שדווח.** הטור הימני הוא ה**שני** בסדר המסמך, ולכן גרירה
שמתחילה בו וממשיכה שמאלה הולכת **אחורה** בסדר המסמך. וזה נמדד במדויק:
הבחירה נשארת **רצף אחד** (`contiguous: true`), כלומר הראש עבר למקום מוקדם
מהעוגן — זהו **היפוך** של הבחירה סביב העוגן, ולא מחיקה של מה שסומן. ההבחנה
אינה קוסמטית: „נמחק” היה שולח לחפש את הבאג בשמירת הבחירה, וההיפוך מצביע על
סדר המילוי, שהוא המקום שבו הוא באמת יושב.

**הנזק הוא בתצוגה ובאינטראקציה בלבד.** הקובץ נשמר נכון ונפתח נכון ב-Word,
וזה נמדד ולא הונח. לכן הפעולה אינה נחסמת — חסימה הייתה מונעת מהמשתמש לייצר
מסמך תקין בגלל באג בציור — אלא מלווה בהודעה בפס המצב (`rtlColumnNote`
ב-`src/engine/page-setup.ts`).

**למה אין תיקון בצד התוסף.** ה-Document API אינו חושף שום ידית לסדר המילוי.
זה כל מה ש-`sections.setColumns` מקבל (`superdoc@2.11.0`,
`dist/document-api/src/sections/sections.types.d.ts`):

```ts
export interface SectionsSetColumnsInput extends SectionTargetInput {
    count?: number;  gap?: number;  equalWidth?: boolean;
}
```

שלושתם נכתבים ל-`w:cols` כמות שהם, ואף אחד מהם אינו נוגע בשאלה איזה טור נמלא
ראשון. `SectionColumns` שנקרא חזרה מ-`sections.list()` נושא בדיוק את אותם
שלושה שדות. **אין `widths`** — לא בכתיבה ולא בקריאה; השם הזה קיים ב-API רק
בהקשר של טבלאות. גם היפוך גיאומטרי של שני ה-DIV-ים הממוקמים-מוחלט שמחזיקים
את הטורים (הם נמדדו — ראו `docs/superdoc-2.10-review.md`) אינו פתרון: הוא
היה מזיז את הציור בלבד, ומשאיר את פגיעת-העכבר ואת סדר הבחירה במקומם — כלומר
הופך „הפוך אבל עקבי” ל„הפוך **ולא** עקבי”.

**מיקום התיקון:** מנוע הפריסה, `packages/layout-engine` במונורפו של superdoc —
קוד פתוח (AGPL v3), ולכן ניתן ל-PR. החלק שקורא את `w:sectPr/w:bidi` ומעביר
אותו ל-`ColumnLayout` נמצא במנוע ה-DOCX הסגור, ולכן PR לצד הפריסה בלבד אינו
מספיק כדי שהתיקון יגיע למשתמשים.

PR פתוח: https://github.com/superdoc/docx-editor/pull/3953 · דווח: issue #3952

**מה ההודעה לא מכסה, במפורש.** היא תלויה ב-`reportCommand`, כלומר יוצאת רק
מהפעולה בתפריט. שני מסלולים אחרים מגיעים לאותה תצוגה הפוכה בשקט:

  1. **מסמך שנפתח כשכבר יש בו שני טורים** — ואצל משתמש אוצריא זה המסלול
     השכיח מבין השניים. הודעה כאן הייתה צריכה להיתלות בפתיחה ולא בפקודה,
     ולהיאמר על מסמך שהמשתמש לא עשה בו דבר; זו החלטה נפרדת, ולא נעשתה.
  2. **„שולחן ערוך ← אחידות טורים”** (`applyColumnsProfile`
     ב-`src/engine/shulchan/sections-uniform.ts`) — משנה מספר טורים במקטעים
     שכבר מרובי-טורים, ומדווח דרך `STATUS_NOTIFIER` („בוצעו N תיקונים”),
     שהוא כותב אחר לאותו פס. שתי הודעות על אותה שורה היו דורסות זו את זו.

**להסרת ההודעה** כשהתיקון ייכנס לגרסה שהתוסף נועל: `rtlColumnNote`
ב-`src/engine/page-setup.ts`, והבדיקות שלה ב-`tests/unit/page-setup.test.ts`
וב-`tests/component/app-shell.test.ts`.
## תוכן עניינים — אין עצירת טאב, ולכן אין נקודות מוביל

### `toc.insert` אינו כותב `w:tabs`, ואף סגנון `TOC*` אינו מוגדר

שורת תוכן עניינים עברית צריכה להיקרא „כותרת......5”: הכותרת בימין, נקודות
מוביל, ומספר העמוד נדחף לשולי השורה. נמדד על superdoc 2.11.0 / docx-engine
0.10.0 — `scripts/qa/toc-direction-probe.mjs`:

- **הכיוון תקין, ומקור אחר ממה ששוער:** אין `<w:bidi/>` ב-`pPr` של פסקאות
  ה-TOC, אבל `sectPr` הוא `bidi` והכיוון נורש ממנו. הכותרת יושבת בימין
  (x=535) ומספר העמוד בשמאלה (x=497). `setDirection` על פסקת TOC כן כותב
  `<w:bidi/>`, ואינו משנה את הפריסה — כלומר הסדר שהיה נכון נשמר.
- **מה שחסר הוא עצירת הטאב:** `TOC1: tabs=[אין]`, הסגנונות `TOC1`/`TOC2`
  **אינם מוגדרים ב-`styles.xml`** כלל, ובגוף הריצות יש `<w:tab/>` בלי
  `w:leader` ובלי `w:pos`. שלוש השכבות שיכולות לשאת עצירה עם
  `w:leader="dot"` — `pPr`, סגנון, `w:defaultTabStop` — וכולן ריקות ממנה.
- **מה שרואים על המסך:** מספר העמוד יושב 38px מתחילת הכותרת ולא בשולי השורה
  ההפוכים, ובלי נקודות בדרך. תו טאב בלי עצירה נופל על `w:defaultTabStop` —
  קפיצה אחת, ובלי מוביל.

`toc.tabLeader` ו-`toc.rightAlignPageNumbers` מחזירים `success: true` ואינם
עושים דבר (מתועד ב-`src/engine/toc.ts`), ולכן הדרך היחידה היא לכתוב `w:tabs`
בעצמנו — או בסגנון `TOC*` שאנחנו מגדירים, או על הפסקאות. טרם נעשה.

## מיקום הגלילה — snapshot יחיד לכל המנועים

### שורש הגלילה נזכר במקום אחד, ושני מסמכים חיים דורסים זה את זה

נמדד ב-superdoc 2.11.0 (docx-engine 0.10.0), Chrome אמיתי על ה-dist הארוז, שני
טאבים פתוחים עם מנוע חי בכל אחד. הדיווח שהוביל לכאן: „הוא זוכר איפה אני, וברגע
שאני מתחיל לגלול הוא חוזר לראש”.

התופעה **אינה** אובדן של `scrollTop`. הוא משוחזר נכון בכל נקודות הזמן שנמדדו
אחרי מעבר טאב, ואז גלגלת אחת מאפסת אותו:

| מתי נמדד `scrollTop` אחרי חזרה לטאב | הערך |
|---|---|
| מיד, באותה משימה | 720 |
| מיקרו-משימה | 720 |
| `requestAnimationFrame` | 720 |
| rAF שני | 720 |
| 150ms | 720 |
| **אחרי גלגלת אחת** | **0** |

הכתיבה היא **JS של המנוע** ולא של הדפדפן: `dist/assets/app.js:2046`, שמקורו
ב-`@superdoc/docx-engine`. הוא מחזיק snapshot **יחיד** של שורש הגלילה, ומזרים
אליו בגלילה הראשונה את הערך שנשמר בו. נמדד ישירות על ה-setter:
`{"ev":"set","asked":0,"was":720,"got":0}` — כלומר „בקשו 0, היה 720, יצא 0”.

**האפס אינו שרירותי:** הוא נקרא בזמן שהפאנל של הטאב האחר היה מוסתר. עם טאב
אחד אין איפוס בכלל, ולכן החתימה של הפער היא בדיוק „שני מנועים חיים חולקים
snapshot אחד”.

**ומה שהופרך באותה מדידה:** ההנחה שהייתה כתובה אצלנו — ש-`display: none` הורס
את קופסת הפריסה ולכן הדפדפן שוכח את הגלילה — אינה נכונה. Chrome **משחזר
בעצמו** אחרי מחזור `display:none`: מיכל שנגלל ל-840 קרא 840 בחזרה, באותה משימה
סינכרונית (840→0→840).

**לא נעקף מהמנוע, ואי אפשר לעקוף בהשמה.** האיפוס קורה מתוך הגלגלת של המשתמש,
כלומר אחרי כל השמה שאנחנו יכולים לעשות בזמן מעבר הטאב. העקיפה בצד שלנו היא
`guardPaneScroll` (`src/sessions/pane-scroll.ts`): מאזין חד-פעמי לאירוע הגלילה
הראשון שאחרי הפעלת טאב, שמזהה את החתימה („המיכל על ראש המסמך ואנחנו זוכרים
אחרת”) ומחזיר את המיקום. הוא מתפרק אחרי שרץ פעם אחת, או ברגע שהמשתמש גלל
בעצמו — שומר שנשאר דרוך היה הופך תיקון חד-פעמי לנעילה.

**מה שנשלל במפורש כפתרון:** להחליף את `display: none` של טאב לא-פעיל
ב-`visibility: hidden`. נכונות ההדפסה נשענת בדיוק על כך שה-`display: none`
ה-inline שורד — `src/styles/print.css` אינו מצהיר `display` ל-`.document-pane`
— ולכן החלפה כזאת הייתה מדפיסה את **שני** הטאבים.

## תצוגה חיה של גופן — שלושה מסלולים, ואף אחד מהם אינו נאמן במסמך

נמדד ב-superdoc 2.11.0 (docx-engine 0.10.0), Chrome 152.0.7977.75, על ה-dist
הארוז. השאלה: המשתמש מרחף על שם גופן ברשימה, ואמור לראות מיד איך הטקסט הנבחר
ייראה. המימוש הראשון עשה זאת בכתיבה למסמך, וזה נשלל במדידה; שני המסלולים
האחרים אינם נוגעים במסמך כלל, ובכל זאת נשללו — מסיבה אחרת לגמרי.

### נקודת הפתיחה: הכתיבה למסמך, ומה היא עולה

`doc.format.apply({ target, inline: { fontFamily } })` מדווח `success: true`,
ובאותה מדידה:

| | לפני | אחרי |
|---|---|---|
| `undoDepth` | 7 | **8** |
| `markDirty` | 39 | **40** |
| `word/document.xml` | — | **שונה** |
| `w:rFonts` בריצה | `[]` | `<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>` |

כלומר ריחוף אחד — בלי לחיצה — מסמן את המסמך כנערך, מתזמן שמירה, וקונה דרגת
undo. `Ctrl+Z` הראשון של המשתמש מבטל „גופן שהוא לא בחר”, ולא את מה שהקליד.
זה מה שהוביל ל-`FONT_PREVIEW_ENABLED = false` ב-`src/engine/font-preview.ts`.

### מה שהופרך: אין ProseMirror, ולכן אין `addToHistory: false`

התיקון המתבקש לכתיבה שמלכלכת היסטוריה הוא טרנזקציה מסומנת
`tr.setMeta('addToHistory', false)`. הוא אינו קיים כאן:

| מה נבדק | מופעים |
|---|---|
| `prosemirror` ב-`node_modules/@superdoc/docx-engine/dist/docx-engine.es.js` (11MB) | **0** |
| `setMeta` באותו באנדל | **0** |
| `addToHistory` בכל עץ המקור `C:/tmp/superdoc-src` | **0** |
| `prosemirror` ב-`THIRD_PARTY_NOTICES` של המנוע | **0** |

וברמת המופע החי: `activeEditor.view`, `.state` ו-`.commands` הם `null`;
`.chain`, `.extensionManager`, `.schema`, `.plugins` ו-`.registerPlugin` הם
`undefined`. כל כיוון שמניח ProseMirror מתחת ל-superdoc 2 מבוזבז מראש.

### מה שכן קיים, ועובד: ספק דקורציות ציבורי, render-only

החוזה קיים ומתועד ב-2.11.0 המותקן:

- `node_modules/superdoc/dist/superdoc/src/public/index.d.ts:180` — `defineSuperDocExtension`
- `node_modules/superdoc/dist/superdoc/src/public/index.d.ts:198` — `SuperDocDecorationProvider`
- `node_modules/superdoc/dist/superdoc/src/core/extensions/types.d.ts:172`, `:344`
- `node_modules/superdoc/dist/superdoc/src/core/types/index.d.ts:4099` — `Config.extensions?: SuperDocExtension[]`
- בעץ המקור: `packages/superdoc/src/core/extensions/types.ts:161`, `:357`, `:480`, `:521`, ודוגמה ב-`define.ts:37`

הרישום **חייב להיות בבנאי** — אין רישום מאוחר. `activate` נקרא פעם אחת לכל
מסמך, לפני `onReady`. `ctx.capabilities` שנמדד:
`{ canRender: true, canUseShortcuts: false, canMutate: false }`, ו-`ctx.visuals`
הוא `{ highlight, decorate, inlineBox }`. ידית `highlight` מדביקה class על
הטקסט, ואת הגופן מחליפים במשתנה CSS.

**השער שהרג את המסלול הראשון עובר כאן במלואו:**

```
הקלדה     → undoDepth 1
צביעה     → undoDepth 1     ← לא זז
Ctrl+Z אחד → ביטל את ההקלדה
```

בכל הגששים: `undoDepth` 0→0, `onEditorUpdate` 0→0, `onTransaction` 0→0,
`diagnostics.getSnapshot()` = `[]`, ו-**כל 12 החלקים ב-docx זהים בייט-בבייט**
לפני, בזמן ואחרי. אין `w:rFonts` כלל. ההחזרה מדויקת לפיקסל (רוחב 192 → 169 →
192). והביצועים אינם השיקול: הצביעה הראשונה 17–35ms, וריחוף לגופן הבא הוא
כתיבת משתנה CSS אחת — 200 החלפות ב-0ms, בלי קריאה למנוע. (המשמעות הנגזרת:
השהיית ה-300ms ב-`src/composables/font-preview.ts` מיותרת לכל צביעה שאינה
הראשונה בסבב.)

**עמידות שנמדדה:** ה-class שורד עימוד מחדש בזום ושורד הקלדה בפסקה אחרת בזמן
שהצביעה חיה, ו-`handle.dispose()` מותיר אפס class ואפס attribute. זה **מפל**,
לא כתיבה לצומת — צמתים שנולדים מעימוד מחדש מקבלים את הכלל מעצמם. לשם השוואה,
`Range.surroundContents` כן עובד (רוחב 473.83 → 494.75) אך **נמחק בהקלדה
הראשונה באותה פסקה** (`wrappersLeft` 1 → 0).

### ובכל זאת נשלל: המנוע אינו זורם מחדש, והתצוגה משקרת

זה הפער האמיתי, והוא **משותף לכל מסלול חזותי** — הוא נובע מ-`position:
absolute` ברוחב מחושב מראש, לא מהטכניקה שמחליפה את הגופן. החלפת גופן משנה רוחב
תווים; המנוע אינו מודד מחדש; הריצה גדלה במקום.

| מה נמדד | Arial (מקור) | בזמן הצביעה |
|---|---|---|
| רוחב הריצה | 389px | **518px** |
| חריגה מעמודת הטקסט | 0 | **122px** |
| `scrollWidth` של העמוד (מול `clientWidth: 794`) | 794 | **824** |
| מספר שורות / עמודים | 3 / 1 | 3 / 1 — **ללא שינוי** |

ומכיוון שהעמוד הוא `overflow: hidden`, 30 הפיקסלים האלה **נחתכים בפועל**. מה
שהמשתמש רואה, מילה במילה מן הצילום:

```
לפני:      The quick brown fox jumps over the lazy dog while the BOLDRU…
בריחוף:      e quick brown fox jumps over the lazy dog while the BOLDRU…
```

תחילת המשפט נעלמה מעבר לקצה.

**המכניקה, ולא מה שנראה במבט ראשון:** העלים כן נמצאים בזרימת inline רגילה,
ולכן ריצה שהתרחבה **דוחפת** את שכנתה ואינה נוחתת עליה — נמדד: השכנה זזה 1px עם
Courier ו-‎29px עם Impact, וחפיפה אנכית **אפס בכל המקרים** (גובה השורה מקובע).
מה שאין הוא **שבירת שורות מחדש**: השורה כולה מתארכת מעבר לטור, ומה שלא נכנס
נחתך בקצה. על שורה שמילאה את הטור נמדד גרוע פי שניים מהטבלה למעלה — ריצה
517 → **759px**, גלישה **237px** מתיבת הטקסט, **141px** מחוץ לעמוד ו-
`scrollWidth` 794 → **939**, כלומר 145px גלילה אופקית שלא הייתה. במסלול החוק
הגלובלי נמדד אותו דבר: ריצה 601 → 797px, **195px** מעבר לשוליים ו-**99px מעבר
לקצה הנייר** — הטקסט על האפור שמסביב לדף.

גופן **צר** נקי לגמרי (Impact: אפס גלישה, `scrollWidth` ללא שינוי). גופן רחב על
שורה מלאה — מכוער, ובלי גידור אפשרי: איננו יודעים מראש אם הטור יספיק.

**ופגם שני, שאין לו תיקון בצד שלנו:** פס הבחירה נשאר במקומו בזמן שהטקסט זז
מתחתיו. נמדד בזרימה האמיתית (בוחרים, ואז מרחפים) — הפס ב-`462..993`, הטקסט עבר
ל-`302..993`; פער של 160px ו-198px. הסימון הכחול והאותיות נפרדים על המסך.

זרימה מחדש הייתה פותרת את שניהם, והיא מחייבת מדידה מחדש של הטקסט במנוע —
כלומר בדיוק את מה ש-`canMutate: false` שולל, ואת מה שהצביעה החזותית נמנעת ממנו
בכוונה. **התצוגה במסמך אינה יכולה להיות נאמנה, ותצוגה שמראה משהו אחר ממה
שיקרה בפועל גרועה מאין תצוגה.**

### פער שלישי: `scope: 'text'` אינו מדויק לתו, בניגוד לחוזה שלו

`SuperDocVisualOptions.scope` מתועד כ„`'text'` paints exact visible text
ranges”. נמדד אחרת: דקורציה על `BBBB` בפסקה `AAAA |BBBB| CCCC` צבעה את **כל
עלה `BBBB`** — לא את הבלוק, אך גם לא את התווים המדויקים. האנקור עצמו כן פותר
נכון (`segments()` = `[{ blockId, range: { start: 5, end: 9 } }]`), ולכן חוסר
הדיוק הוא בצייר.

בבאנדל הסגור (`docx-engine.es.js`, פונקציות `Wxn`/`Z2o`/`eBo`/`tBo`) יש שני
מסלולים — מוכוון-טווח כשה-host מספק callback לצביעת טווח, ונפילה לכל הבלוק
בלעדיו. שרץ הראשון, ולכן התוצאה היא עלה ולא בלוק.

**וזה מה שמכריע את הפיצ'ר, יותר מן החיתוך.** נמדד מקצה לקצה, ריצה אחת:

| | הטקסט |
|---|---|
| הריצה במסמך | `one two three four five` (23 תווים) |
| מה שהמשתמש סימן | `‎ two thre` (9 תווים) |
| מה שהריחוף צבע | `one two three four five` — **כל 23** |
| מה שהלחיצה שינתה בפועל | `‎ two thre` בלבד — מאושר ב-OOXML |

כלומר התצוגה מראה שינוי רחב פי שניים וחצי ממה שיקרה. הכתובת מדויקת והצביעה גסה,
וזו הטעיה ולא אי-דיוק.

הסתירה לחוזה מנוסחת בעץ הפתוח עצמו —
`packages/superdoc/src/core/extensions/types.ts:497`: „paints exact visible text
ranges … never silently paints the whole block”. מכיוון שהצייר סגור, מה שפתוח
כאן הוא **דיווח באג**, לא PR.

הצייר של `highlight` (התכונות `data-superdoc-extension-decoration` ו-`-classes`)
נמצא **רק** בבאנדל הסגור — אפס מופעים בכל עץ המקור הפתוח, ולכן אין כאן PR.

### שלוש חסימות נוספות שנמדדו, לחוסך למי שינסה שוב

**`::highlight()` בכרום 152 אינו נושא `font-family`.** עם
`background: yellow; color: red; font-family: "Courier New"; font-weight: 900`
הצבעים כן נצבעו, אבל `Range.getBoundingClientRect().width` נשאר **46.31 →
46.31**. גם עם `font-family` ו-`font-size: 32px` בלבד — רוחב 46.31, גובה 17.
ה-Custom Highlight API הוא קבוצת תכונות צביעה, ואינו יכול להחליף גופן.

**`window.getSelection()` אינו רואה את הבחירה.** עם בחירה חיה ומאומתת
(`{ status: 'ready', empty: false }`, פס כחול על המסך): `rangeCount: 1`,
`collapsed: true`, `textLen: 0`, `clientRects: 0`. ה-`activeElement` הוא
`<textarea data-v2-ime-host>` בגודל 1px ובאטימות 0. המנוע אינו מניח את הבחירה
ב-DOM, ולכן כל מסלול שנשען על `Range` מן הדפדפן חסום מהיסוד. גם למפות
מודל→DOM בעצמנו אינו פתוח: ה-API נותן `{ blockId, offset }` וה-DOM נושא
`data-pm-start`/`data-pm-end` במספור אחר (גלובלי מול מקומי, סטייה של אחד).

**`!important` הוא חובה, לא סגנון.** המנוע כותב `font-family` **inline** על כל
`span.superdoc-text-run`, ולכן כלל CSS על class — כולל class שידית דקורציה
מדביקה — לא יתפוס בלעדיו. נמדד: כלל ירושה על המיכל נכשל לחלוטין
(`computedFF` נשאר `Arial, __superdoc_core_symbols__, sans-serif`), ואותו כלל
עם `!important` הצליח. מימוש בלי `!important` נכשל **בשקט**.

### `paintInlineBoxes` מדלג על RTL לחלוטין — וזה פתוח ל-PR

`C:/tmp/superdoc-src/packages/layout-engine/painters/dom/src/runs/inline-box.ts`
פותח את `paintInlineBoxes` ב-`if (isRtl) return;`. החבילה
(`@superdoc/painter-dom`) היא **קוד פתוח**, בשונה מצייר ה-`highlight`. אינו
נדרש לתצוגת הגופן, ונרשם כאן מפני שהוא הפתח היחיד שנמצא במסלול הזה לתיקון
במעלה הזרם.

### מה שנשלח במקומו

תצוגה שאינה במסמך: **פס דגימה בתחתית רשימת הגופנים**, שמציג את הטקסט הנבחר של
המשתמש בגופן שמרחפים עליו. הוא נאמן — אין בו עימוד לשקר עליו — ואינו נוגע
במסמך, בהיסטוריה או בשמירה. הטקסט מגיע מ-`doc.selection.current({ includeText:
true })`, שנמדדה כקריאה טהורה: 1ms, ו-`history.get()` זהה לפניה ואחריה.

**מלכודת מדידה שיש לרשום:** `doc.history.get()` מחזיר **Promise**. בלי `await`
הוא נראה כאובייקט ריק `{}`, וכל שער undo שנשען עליו „עובר” בלי למדוד דבר. שתי
מדידות ראשונות בסבב הזה נפלו בדיוק כך.

**ומה שנשלל במפורש כפתרון:** להשתמש ב-`.superdoc-text-run` כסלקטור. היא מחלקה
פנימית של המנוע, מאותה קטגוריה ש-`tests/unit/engine-boundaries.test.ts` אוסר
ושנדחתה כבר ב-`src/engine/line-number-layer.ts`. נמדד ש-`.editor-stack__host *`
— מחלקה שלנו (`src/sessions/editor-swap.ts`) — נותנת תוצאה זהה בדיוק, ולכן גם
אילו המסלול החזותי היה נשלח, לא היה בו צורך בשם פנימי.

## גל 23 — אזכור „@” וקישורי עומק

### `hyperlinks.wrap` נכשל על טווח שנכתב זה עתה

`wrap` מחזירה `INVALID_TARGET` (`"hyperlinks.wrap could not resolve the target
range."`) על טווח שנוצר בקריאת `doc.insert` שקדמה לה — גם כאשר
`doc.ranges.resolve` על **אותו טווח בדיוק** מחזירה את הטקסט הנכון ואת אותם
היסטים. נמדד ב-`scripts/qa/wrap-after-insert-probe.mjs`:

```
resolved: { text: "פסחים דף לד", start: 4, end: 15 }
wrap:     { success: false, code: "INVALID_TARGET" }
```

**שלוש הסתעפויות שנשללו במדידה:**

- **תזמון** — ניסיון שני אחרי `requestAnimationFrame` + 60ms נכשל זהה.
- **`blockId` שהתחלף בהחלפה** — ה-`blockId` נקרא מהבחירה שאחרי ההכנסה והוא
  זהה לזה שלפניה (`41964671`).
- **סינון הסכימה** — `wrap` עם `href` של `https://example.com/x` על אותו טווח
  נכשל באותה הודעה. כלומר ההודעה על „הטווח” אינה מכסה על דחיית `otzaria:`.

`wrap` נמדדה עובדת רק על הטווח שהמשתמש סימן בעצמו — כך היא נקראת
ב-`src/engine/hyperlinks-manage.ts`, וזה המסלול שממשיך לעבוד.

### מה שנשלח במקומו

מחיקת הטווח ואז הוספה: `doc.insert({ value: '' })` על האזכור, ואז
`doc.hyperlinks.insert({ target: נקודה, text, link })`. שתי הפעולות נמדדו
מצליחות ברצף, והתוצר נבדק ב-OOXML ולא בקבלה — `<w:hyperlink r:id>` בתוך
`document.xml` יחד עם ה-`Relationship` התואם ב-`document.xml.rels`.

`hyperlinks.insert` לבדה אינה מספיקה: נמדד שהיא **מוסיפה ואינה מחליפה**, וללא
המחיקה שלפניה נשאר „@פסחים לד” לצד הקישור שנוצר.

### `otzaria://` נחסם לציור, ולכן הקישור אינו לחיץ בתוך העורך

ה-DomPainter מריץ כל `href` דרך סניטציה שמשווה את הסכימה מול רשימה קבועה —
`http`, `https`, `mailto`, `tel`, `sms` (ו-`ftp`/`sftp`/`irc` כאופציונליות) —
וחוסם את השאר עם `[DomPainter] Blocked potentially unsafe URL`. הקריאה
לסניטציה נעשית **בלי להעביר אופציות**, כך ש-`allowedProtocols` הקיים בחתימה
אינו נגיש מנקודת השימוש, ואין קונפיגורציה בבנאי.

**מה זה כן שובר ומה לא:** הקישור נכתב ל-DOCX תקין ועובד בכל תוכנה שפותחת את
הקובץ. מה שאינו עובד הוא לחיצה עליו **בתוך העורך שלנו**:
`hyperlinks.onActivate` (`src/engine/otzaria-link-activation.ts`) אינו נקרא, כי
המנוע חוסם את ה-href לפניו. ה-handler נשאר מחווט — הוא נכון, זול, ויתחיל
לעבוד ברגע שהסכימה תעבור.

השער `scripts/qa/at-mention-qa.mjs` מודד את הפער הזה בצעד נפרד ומדווח אותו
כ„חלקי”, כך שסגירתו במעלה הזרם תיראה מיד.

### `hyperlink-nested-unsupported`

הוספת קישור כשהסמן כבר בתוך קישור קיים נכשלת בקוד הזה, עם הודעה באנגלית מן
המנוע. `at-mention-overlay.ts` מתרגם אותו להודעה בעברית ואינו מנסה לעקוף.
