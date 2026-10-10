# שערי ה-QA של הפקדים

המסגרת כאן פותחת את ה-`dist` הארוז ב-Chrome אמיתי, מזריקה דמה של מאחז אוצריא,
ולוחצת על הפקדים ברצועה **בלחיצות עכבר אמיתיות** דרך CDP. זה המקום היחיד שבו
„הכפתור עובד” נמדד ולא מונח: הבדיקות ב-`tests/` רצות ב-jsdom, ושם אין מנוע DOCX,
אין workers ואין עימוד — כלומר אין מסמך שאפשר לבדוק אם משהו נכתב אליו.

## איך מריצים

```
npm run build          # חובה — השערים רצים על dist, לא על המקור
node scripts/qa/<שם>-qa.mjs
```

`npm run check:shell-colors` מודד צבעים מחושבים וניגודיות בנושא בהיר וכהה,
כולל החלפת נושא דרך אירוע SDK, בחירת צבע בד ואיפוסו דרך הבורר. הוא בודק גם
שצבע הבד אינו משנה את פקדי המעטפת, שפופאוברים נשארים קריאים ושמעבר מסמך
אינו משנה את רוחבי הטאבים. השער נכלל ב-`verify` וב-`verify:qa`.

את כל הרשימה מריץ `npm run verify:qa`, ודרך `run-all.mjs` ולא שרשרת `&&`:
שרשרת עוצרת בכשל הראשון, וכל מה שאחריו אינו רץ בכלל — בפלט שנראה כמו ריצה
מלאה. הרַץ מריץ עד הסוף, מסכם, ומחזיר קוד יציאה 1 אם מישהו נכשל.

כל שער חייב **יציאה משלו** (`port`). שני שערים על אותה יציאה מדברים עם אותו
דפדפן ומשחיתים זה את מדידתו.

שני משתני סביבה, ולא אחד: השערים כאן קוראים את יציאת ה-CDP מ-`QA_PORT`
(`harness.mjs`, וברירת מחדל אחרת לכל שער), והכלים שמעל — `scripts/cdp.mjs`
ומה שמשתמש בו — מ-`CDP_PORT` (ברירת מחדל 9333). מי שמריץ שער ובודק
שמעל `scripts/` במקביל, ומגדיר רק אחד מהשניים, מקבל „CDP לא נפתח” על שער
שתפוס — כשל תשתית שנראה בדיוק כמו תקלת מוצר. הגדירו את שניהם, בערכים שונים.

## השלד של שער

```js
import { openApp, createReport } from './harness.mjs';

// `strict: true` — שער שמפיל את הריצה על שורה שבורה. בלעדיו זהו שער **סקר**:
// הוא מודד ומדווח, מכריז על עצמו בסוף שקוד היציאה אינו נגזר מהשורות, ואינו
// מפיל. בחרו במפורש; ברירת המחדל היא סקר.
const report = createReport('שם הקבוצה', { strict: true });
const app = await openApp({ name: 'layout', port: 9361 });
try {
  await app.caret(0);            // בלי סמן כל פקד מדווח „יש למקם את הסמן”
  await app.tab('פריסה');
  await app.openMenu('שוליים');
  await app.clickMenu('רחב');
  const files = await app.docx(); // ההוכחה: מה נכתב ל-OOXML
  /w:left="2880"/.test(files['word/document.xml'])
    ? report.pass('שוליים רחבים') : report.fail('שוליים רחבים', 'ה-pgMar לא השתנה');
} finally { app.close(); }
process.exit(report.print() > 0 ? 1 : 0);
```

## מה יש ב-`app`

**ניווט ולחיצה**
- `tab(label)` — מעבר לשונית (`קובץ`, `בית`, `הוספה`, `פריסה`, `הפניות`, `סקירה`, `תצוגה`, `✦ אוצריא`)
- `click(name)` — לחיצה על פקד לפי שמו (הכותרת בטולטיפ, ובהיעדרה ה-`title` בלי הקיצור). מחזירה `false` כשלא נמצא
- `clickAt(x, y)`, `clickSel(selector, index)`, `escape()`
- `openMenu(name)` → פריטי התפריט, `menuItems()`, `clickMenu(label)`, `menuOpen()`
- `clickGallery(label)`, `galleryItems()`
- `clickPalette(index)`, `paletteSwatches()`, `paletteOpen()`
- `clickTableCell(row, col)`
- `selectValue(name, value)` / `options(name)` — ל-`<select>` (בורר גופן, גודל)

**דיאלוגים** — `dialog()` (שם + כל הפקדים), `dialogFill(idאוName, value)`, `clickDialog(name)`

**מסמך וקלט**
- `caret(lineIndex)` — לחיצה בשורת טקסט; `selectLine(i)`, `extendSelection(n)`.
  **`lineIndex` אינו אינדקס פסקה**: הסלקטור שמאחוריו הוא
  `'.superdoc-line, .superdoc-fragment'`, וה-line מקונן ב-fragment — כלומר כל
  פסקה תופסת שני אינדקסים, ו-`caret(1)` הוא עוד הפסקה הראשונה. `caret(0)`
  תמיד הפסקה הראשונה, ולכן רוב הקוראים אינם נפגעים
- `caretPara(indexאוText)` — סמן בפסקה, **מאומת מול המנוע** (`data-source-node-id`
  של ה-fragment מול `doc.selection.current()`) וזורק כשהלחיצה נחתה על אחרת.
  ‏`paraCount()`, `caretBlock()`. זה מה שצריך כשהמדידה תלויה **באיזו** פסקה
- `type(text)`, `press(key, code, vk, modifiers, text)`
- `screenText()`, `lineCount()`, `selection()`
- `docx()` → מפה של `נתיב → מחרוזת` מתוך ה-docx המיוצא. **זו ההוכחה.** `word/document.xml`,
  `word/styles.xml`, `word/settings.xml`, `word/numbering.xml`, `word/footnotes.xml` וכו'

**מצב ודיווח**
- `state(name)` → `{found, disabled, active, pressed, value, visible, rect}`
- `controls(scope)` → כל הפקדים בלשונית הפעילה
- `cmd(id)` → מצב הפקודה כפי שהמנוע מדווח (`ui.commands`)
- `status()` → שורת המצב (`{text, error}`) — שם מופיעה הודעת הכשל של הפקודה
- `messages()` → מה שהוצג דרך המאחז (`ui.showError` וכו'), `hostCalls()` → כל הקריאות למאחז
- `log()` → אזהרות, שגיאות ו-rejections שנתפסו בדף
- `reset()` — מנקה `messages`, `hostCalls` ו-`log` לפני הצעד הבא

## הכלל של המדידה

`success: true` **אינו** הוכחה. פקד שנחשב עובד הוא פקד ש:
1. קיים, מוצג, ואינו מושבת בהקשר שבו הוא אמור לפעול;
2. **שינה את המסמך** בצורה שנקראת ב-`docx()` — או שינה מצב מנוע שאפשר לקרוא
   (`cmd(id).active`, זום, נראות סרגל);
3. לא הותיר שגיאה ב-`status()`, ב-`messages()` או ב-`log()`.

פקד שנכשל באחד השלושה — `report.fail`, עם הפירוט שנמדד.

לפני שקובעים ש„הפקד שבור” — לקרוא את המודול שלו ב-`src/engine/`. חלק מהפקדים
מוגבלים **בכוונה** בגלל פערים שנמדדו במנוע, והם מתועדים ב-`docs/engine-gaps.md`.
