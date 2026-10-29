/**
 * שפת התפריטים (הרצועה) לפי שפת המשתמש.
 *
 * אוצריא מדווחת את שפת המשתמש ב־`plugin.boot`: `app.language` הוא קוד השפה
 * בלבד (`'he'` / `'en'`) ו־`app.locale` התג המלא (`'he-IL'`) — ראו
 * docs/plugin-sdk (BootPayload.app). `setMenuLocale` נקראת מ־main.ts ברגע
 * שה־payload מגיע, ומאז כל מחרוזת שעוברת דרך `menuString` מוצגת באנגלית
 * כששפת המשתמש אנגלית, ובעברית בכל יתר המקרים.
 *
 * שתי החלטות עיצוב:
 *
 * 1. **התרגום בנקודת התצוגה, לא במקור.** לשוניות הרצועה ממשיכות להכיל את
 *    המחרוזות העבריות המקוריות, והתרגום קורה ברכיבי הבסיס שמציגים אותן
 *    (RibbonButton, RibbonMenuButton, RibbonGroup, RibbonSelect,
 *    StyleGallery, TablePicker, ColorPickerPopover ו־Ribbon.vue). כך עבודת
 *    התפריטים והמקור עצמם אינם משתנים — רק מה שמוצג. ברירת המחדל `'he'`
 *    מחזירה כל מחרוזת כמו שהיא, ולכן בהרכבות הבדיקות (שפת stub הפיתוח
 *    היא `'he'`) לא משתנה דבר.
 *
 * 2. **מילון מהעברית לאנגלית, עם נפילה חזרה למקור.** מחרוזת שאינה במילון
 *    מוצגת בעברית — תרגום חסר אינו יכול לשבור כלום, ולכן גם אינו נראה:
 *    tests/unit/menu-locale-coverage.test.ts סורק את מחרוזות הרצועה ונכשל על
 *    כל אחת שאין לה כאן ערך, ועל כל מפתח שאינו במקור.
 *
 * הגבול: הרצועה בלבד. שורת הכותרת, שורת המצב, הסרגלים, תפריט ההקשר, רשימת
 * הקיצורים והדיאלוגים אינם עוברים כאן — הם עברית בכל שפה.
 */

import { ref } from 'vue';

export type MenuLocale = 'he' | 'en';

/**
 * השפה גם על שורש ה-HTML, ולא רק ב-`ref`.
 *
 * הצריכה היחידה מבחוץ היא styles/engine-chrome.css: העברות של שכבת הכותרות
 * שהמנוע מצייר חייבת להיות CSS (ראו ההסבר שם), ו-CSS אינו יכול לקרוא `ref`.
 * שכבה שמתורגמת בחצי CSS וחצי JS מחייבת שהשניים ייכבו יחד, ותכונה על השורש
 * היא האות המשותפת. `engine/hf-chrome.ts` קורא את אותו סימן.
 *
 * היעדר התכונה הוא עברית — שפת המקור — ולכן לפני האתחול ובבדיקות שאינן קוראות
 * ל-`setMenuLocale` הכללים חלים, כמו קודם.
 */
export const MENU_LOCALE_ATTRIBUTE = 'data-menu-locale';

const locale = ref<MenuLocale>('he');

/**
 * קביעת שפת התפריטים מקוד שפה או מתג מלא.
 *
 * רק `'en…'` מתרגם לאנגלית; כל יתר הערכים (כולל `'he-IL'`, `undefined`,
 * וכשל אתחול) משאירים עברית — שפת המקור של התוסף.
 */
export function setMenuLocale(language: string | null | undefined): void {
  const code = (language ?? '').toLowerCase().split(/[-_]/)[0];
  locale.value = code === 'en' ? 'en' : 'he';
  if (typeof document !== 'undefined') {
    document.documentElement.setAttribute(MENU_LOCALE_ATTRIBUTE, locale.value);
  }
}

/** השפה שנקבעה. נקראת ממי שאינו תבנית — ראו `engine/hf-chrome.ts`. */
export function menuLocale(): MenuLocale {
  return locale.value;
}

/** התרגומים לאנגלית, לפי מחרוזת המקור העברית. */
const EN: Readonly<Record<string, string>> = {
  // ── לשוניות (Ribbon.vue) ───────────────────────────────────────────
  'קובץ': 'File',
  'בית': 'Home',
  'הוספה': 'Insert',
  'פריסה': 'Layout',
  'הפניות': 'References',
  'סקירה': 'Review',
  'תצוגה': 'View',
  'מפתחים': 'Developer',
  '✦ אוצריא': '✦ Otzaria',
  'לשוניות הרצועה': 'Ribbon tabs',
  'הצג את הרצועה': 'Show the ribbon',
  'כווץ את הרצועה': 'Collapse the ribbon',
  'אין מקום לקבוצה ברוחב הזה — לחיצה פותחת את הפקדים שלה':
    'No room for this group at this width — click to open its controls',

  // ── לשונית „קובץ” ──────────────────────────────────────────────────
  'קובץ ומסמך': 'File & Document',
  'מסמך חדש': 'New Document',
  'יצירת מסמך Word ריק חדש': 'Create a new blank Word document',
  'פתח קובץ': 'Open File',
  'פתיחת מסמך Word (.docx) מהמחשב': 'Open a Word (.docx) document from the computer',
  'שמירה': 'Save',
  'שמור': 'Save',
  'שמירת שינויים במסמך': 'Save changes to the document',
  'שמור בשם...': 'Save As...',
  'שמירת המסמך כקובץ חדש': 'Save the document as a new file',
  'ייצוא והדפסה': 'Export & Print',
  'הדפסה': 'Print',
  'הדפסת המסמך': 'Print the document',
  'ייצוא ל-PDF': 'Export to PDF',
  'שמירת המסמך כקובץ PDF': 'Save the document as a PDF file',
  'ייצוא לאוצריא': 'Export to Otzaria',
  'שמירת המסמך כספר בפורמט של אוצריא (טקסט עם רמות כותרות), לקליטה בספרייה':
    'Save the document as a book in Otzaria format (text with heading levels), ready for the library',
  'ייצוא ל-PDF דורש גרסה עדכנית יותר של אוצריא':
    'Export to PDF requires a newer version of Otzaria',
  'יציאה': 'Exit',
  'סגירת המסמך וחזרה למסך הספרייה של אוצריא':
    'Close the document and return to the Otzaria library screen',
  'אין מסמך פתוח': 'No open document',
  'השמירה רצה כרגע — רגע אחד': 'Saving is in progress — one moment',
  'פתיחת מסמך רצה כרגע': 'Opening a document is in progress',
  'סגירת המסמכים רצה כרגע — רגע אחד': 'Closing documents is in progress — one moment',
  'מידע': 'Info',
  'אודות': 'About',
  'אודות עורך Word לאוצריא': 'About the Word Editor for Otzaria',
  'קיצורים': 'Shortcuts',
  'רשימת קיצורי המקלדת': 'List of keyboard shortcuts',
  'ניהול קיצורים': 'Manage shortcuts',
  'דיווח על באג': 'Report a bug',
  'פתיחת דיווח על תקלה במאגר התוסף ב-GitHub': 'Open a bug report in the plugin repository on GitHub',
  'הצמדת צירוף מקשים לערכת עיצוב': 'Bind a key combination to a formatting preset',

  // ── לשונית „בית” ───────────────────────────────────────────────────
  'לוח': 'Clipboard',
  'הדבק': 'Paste',
  'הדבקת תוכן מהלוח': 'Paste content from the clipboard',
  'אפשרויות הדבקה': 'Paste Options',
  'שמירה על עיצוב המקור, מיזוג עיצוב או טקסט בלבד — ובחירת ההתנהגות של Ctrl+V':
    'Keep source formatting, merge formatting or keep text only — and choose what Ctrl+V does',
  'שמירה על עיצוב המקור': 'Keep Source Formatting',
  'הגופן, הגודל, הצבע והעיצוב כמו במקום שממנו הועתק':
    'The font, size, color and formatting of the place it was copied from',
  'Ctrl+V: שמירה על עיצוב המקור': 'Ctrl+V: Keep Source Formatting',
  'לחיצה: שמירה על עיצוב המקור': 'Click: Keep Source Formatting',
  'מיזוג עיצוב': 'Merge Formatting',
  'העיצוב של המקום שאליו מדביקים, עם המודגש, הנטוי והקו התחתון של המקור':
    'The formatting of the place you paste into, keeping bold, italic and underline from the source',
  'Ctrl+V: מיזוג עיצוב': 'Ctrl+V: Merge Formatting',
  'לחיצה: מיזוג עיצוב': 'Click: Merge Formatting',
  'שמירת טקסט בלבד': 'Keep Text Only',
  'הטקסט בלבד, בעיצוב של המקום שאליו מדביקים': 'The text only, in the formatting of the place you paste into',
  'Ctrl+V: שמירת טקסט בלבד': 'Ctrl+V: Keep Text Only',
  'לחיצה: שמירת טקסט בלבד': 'Click: Keep Text Only',
  'ההדבקה הרגילה — נשמרת להפעלות הבאות': 'The regular paste — remembered for next time',
  'גזור': 'Cut',
  'גזירת הבחירה ללוח': 'Cut the selection to the clipboard',
  'העתק': 'Copy',
  'העתקת הבחירה ללוח': 'Copy the selection to the clipboard',
  'מברשת עיצוב': 'Format Painter',
  'העתק עיצוב ממקום אחד והחל במקום אחר':
    'Copy formatting from one place and apply it elsewhere',
  'גופן': 'Font',
  // כותרות הקבוצות בבורר הגופן (engine/font-options.ts).
  'גופנים אחרונים': 'Recent Fonts',
  'עברית': 'Hebrew',
  'כל הגופנים': 'All Fonts',
  'אין גופן בשם הזה': 'No matching font',
  'Enter מחיל את הגודל שהוקלד': 'Enter applies the size you typed',
  'פתח את הרשימה': 'Open the list',
  'גודל גופן': 'Font Size',
  'הגדל גופן': 'Increase Font Size',
  'מגדיל את הטקסט המסומן בדרגה אחת בכל לחיצה':
    'Grows the selected text by one step per click',
  'הקטן גופן': 'Decrease Font Size',
  'מקטין את הטקסט המסומן בדרגה אחת בכל לחיצה':
    'Shrinks the selected text by one step per click',
  'נקה את כל העיצוב': 'Clear All Formatting',
  'מחזיר את הטקסט המסומן לעיצוב הרגיל, והתוכן נשאר':
    'Returns the selected text to plain formatting; the content stays',
  'מתקדם': 'Advanced',
  'ריווח תווים, מיקום ביחס לשורה, אפקטים וגופן מורכב':
    'Character spacing, position relative to the line, effects and complex-script font',
  'מודגש': 'Bold',
  'מעבה את הטקסט המסומן': 'Makes the selected text bold',
  'נטוי': 'Italic',
  'מטה את הטקסט המסומן': 'Slants the selected text',
  'קו תחתון': 'Underline',
  'מוסיף קו מתחת לטקסט המסומן': 'Adds a line under the selected text',
  'קו חוצה': 'Strikethrough',
  'מעביר קו באמצע הטקסט המסומן': 'Draws a line through the middle of the selected text',
  'צבע סימון טקסט': 'Text Highlight Color',
  'צבע גופן': 'Font Color',
  'כתב עליון': 'Superscript',
  'כתב תחתי': 'Subscript',
  'פיסקה': 'Paragraph',
  'תבליטים': 'Bullets',
  'הופך את הפסקאות המסומנות לרשימה מסומנת בנקודות':
    'Turns the selected paragraphs into a bulleted list',
  'פעולות תבליטים': 'Bullet actions',
  'הפוך רשימה ממוספרת לתבליטים': 'Convert a numbered list to bullets',
  'החלפת רשימה ממוספרת לתבליטים, והמרה לטקסט':
    'Switch a numbered list to bullets, and convert to text',
  'מספור': 'Numbering',
  'הופך את הפסקאות המסומנות לרשימה ממוספרת':
    'Turns the selected paragraphs into a numbered list',
  'פעולות מספור': 'Numbering actions',
  'סגנון מספור (כולל עברי), התחלה מחדש והמרה לטקסט':
    'Numbering style (including Hebrew), restart, and convert to text',
  // תוויות סגנון המספור (engine/lists.ts). דוגמאות האותיות הן הפלט עצמו
  // ונשארות עבריות; רק שם השיטה מתורגם.
  'א, ב, ג … יא, יב (גימטריה)': 'א, ב, ג … יא, יב (Gematria)',
  'א, ב, ג … כ, ל (אלף־בית)': 'א, ב, ג … כ, ל (alphabetical)',
  'הקטן הזחה': 'Decrease Indent',
  'מקרב את הפסקה לשולי הדף': 'Moves the paragraph closer to the page margin',
  'הגדל הזחה': 'Increase Indent',
  'מרחיק את הפסקה משולי הדף': 'Moves the paragraph away from the page margin',
  'רשימה': 'List',
  'התחל מחדש מ-1': 'Restart at 1',
  'המשך מספור קודם': 'Continue previous numbering',
  'המר לטקסט…': 'Convert to text…',
  'לחץ שוב לאישור — הפעולה בלתי-הפיכה':
    'Click again to confirm — this action cannot be undone',
  'זיהוי רשימות בהקלדה: פעיל': 'List autoformat while typing: on',
  'זיהוי רשימות בהקלדה: כבוי': 'List autoformat while typing: off',
  'כיוון פסקה מימין לשמאל': 'Paragraph direction right-to-left',
  'מסדר את הפסקה לקריאה בעברית: ההזחה והיישור בצד ימין':
    'Sets the paragraph up for Hebrew reading: indent and alignment on the right',
  'כיוון פסקה משמאל לימין': 'Paragraph direction left-to-right',
  'מסדר את הפסקה לקריאה בלטינית: ההזחה והיישור בצד שמאל':
    'Sets the paragraph up for Latin reading: indent and alignment on the left',
  'הצג/הסתר סימני עיצוב': 'Show/Hide formatting marks',
  'מציג סימני פסקה, טאבים ורווחים על המסך. הם אינם מודפסים':
    'Shows paragraph marks, tabs and spaces on screen. They are not printed',
  'יישור לימין': 'Align Right',
  'מיישר את הפסקה לשוליים הימניים': 'Aligns the paragraph to the right margin',
  'מרכז': 'Center',
  'ממרכז את הפסקה בין שני השוליים': 'Centers the paragraph between both margins',
  'יישור לשמאל': 'Align Left',
  'מיישר את הפסקה לשוליים השמאליים': 'Aligns the paragraph to the left margin',
  'יישור לשני הצדדים': 'Justify',
  'מותח את השורות עד שני השוליים, מלבד השורה האחרונה':
    'Stretches the lines to both margins, except the last line',
  'מרווח בין שורות': 'Line Spacing',
  'מרווח שורות וריווח': 'Line and Paragraph Spacing',
  'אפשרויות מרווח שורות…': 'Line Spacing Options…',
  'הוסף רווח לפני הפסקה': 'Add Space Before Paragraph',
  'הסר רווח לפני הפסקה': 'Remove Space Before Paragraph',
  'הוסף רווח אחרי הפסקה': 'Add Space After Paragraph',
  'הסר רווח אחרי הפסקה': 'Remove Space After Paragraph',
  // כיתוב פס התצוגה שבתחתית אותו תפריט. „נק׳” הוא הגרש העברי ולא אפוסטרוף —
  // אותה מחרוזת בדיוק נסרקת בשער הכיסוי, ושתי הצורות אינן אותו מפתח.
  'שורה בודדת': 'Single',
  'שורה כפולה': 'Double',
  'מרובה': 'Multiple',
  'בדיוק': 'Exactly',
  'לפחות': 'At least',
  'לפני': 'Before',
  'אחרי': 'After',
  'נק׳': 'pt',
  'התצוגה תופיע כשהמסמך ייטען': 'The preview appears once the document loads',
  'הגדל מרווח שורות': 'Increase Line Spacing',
  'הקטן מרווח שורות': 'Decrease Line Spacing',
  'תפריט פסקה': 'Paragraph dialog',
  'כניסות, ריווח בין פסקאות, מרווח שורות ועצירות טאב':
    'Indents, paragraph spacing, line spacing and tab stops',
  'סגנונות': 'Styles',
  'אפשרויות סגנונות': 'Style options',
  'החל סגנון': 'Apply style',
  'עדכן כך שיתאים לבחירה': 'Update to match selection',
  'שנה סגנון…': 'Modify…',
  'שנה שם…': 'Rename…',
  'צור סגנון…': 'Create a style…',
  'שנה את הסגנון הנוכחי…': 'Modify current style…',
  'עריכה': 'Editing',
  'חפש': 'Find',
  'חיפוש טקסט במסמך': 'Search text in the document',
  'החלפה': 'Replace',
  'החלפת טקסט במסמך': 'Replace text in the document',
  'בחר הכל': 'Select All',
  'בחירת כל הטקסט במסמך': 'Select all text in the document',
  'הרמת הטקסט המסומן לכתב עליון; לחיצה נוספת מחזירה אותו לשורה':
    'Raise the selected text to superscript; clicking again restores it inline',
  'הנמכת הטקסט המסומן לכתב תחתי; לחיצה נוספת מחזירה אותו לשורה':
    'Lower the selected text to subscript; clicking again restores it inline',

  // ── לשונית „הוספה” ─────────────────────────────────────────────────
  'טבלאות': 'Tables',
  'טבלה': 'Table',
  'הוספת טבלה': 'Insert table',
  'הוסף טבלה': 'Add table',
  'בחירת מידות הטבלה': 'Choose table dimensions',
  'עמודה אחת': 'one column',
  'עמודות': 'columns',
  'שורה אחת': 'one row',
  'שורות': 'rows',
  'על': 'by',
  'איורים': 'Illustrations',
  'תמונות': 'Pictures',
  'הוספת תמונה מקובץ (PNG או JPEG)': 'Insert a picture from a file (PNG or JPEG)',
  'התמונה נוספת למסמך…': 'Picture added to the document…',
  'בחירת התמונה נכשלה': 'Picture selection failed',
  'לא ניתן להטמיע את התמונה הזאת במסמך':
    'This picture cannot be embedded in the document',
  'קישורים': 'Links',
  'קישור': 'Link',
  'הוספת היפר-קישור לכתובת אינטרנט או לדואר':
    'Insert a hyperlink to a web address or e-mail',
  'הסר קישור': 'Remove Link',
  'הסרת ההיפר-קישור מהטקסט המסומן (הטקסט נשמר)':
    'Remove the hyperlink from the selected text (the text is kept)',
  'כותרת עליונה': 'Header',
  'כותרת תחתונה': 'Footer',
  'עריכת כותרת עליונה': 'Edit header',
  'עריכת כותרת תחתונה': 'Edit footer',
  'הסרת כותרת עליונה': 'Remove header',
  'הסרת כותרת תחתונה': 'Remove footer',
  'יוצר כותרת ריקה אם עדיין אין': 'Creates an empty one if there isn’t yet',
  'מוחק את הכותרת מכל המסמך': 'Deletes it from the entire document',
  'יצירת כותרת עליונה ריקה. לעריכתה — לחיצה כפולה על אזור הכותרת':
    'Creates an empty header. Double-click the header area to edit it',
  'יצירת כותרת תחתונה ריקה. לעריכתה — לחיצה כפולה על אזור הכותרת':
    'Creates an empty footer. Double-click the footer area to edit it',
  'למסמך יש כותרת עליונה. לחיצה כפולה על אזור הכותרת פותחת אותה לעריכה':
    'The document has a header. Double-click the header area to edit it',
  'למסמך יש כותרת תחתונה. לחיצה כפולה על אזור הכותרת פותחת אותה לעריכה':
    'The document has a footer. Double-click the footer area to edit it',
  'טקסט': 'Text',
  'מספר עמוד': 'Page Number',
  'הכנסת שדה מספר עמוד במקום הסמן': 'Insert a page number field at the cursor',
  'מספר העמודים במסמך': 'Number of pages in the document',
  'מתעדכן לפי העמוד שהשדה נמצא בו': 'Updates according to the page the field is on',
  'לצירוף „עמוד X מתוך Y” יש להקליד את המילים ולהוסיף את שני השדות':
    'For “Page X of Y”, type the words and insert both fields',
  'תאריך ושעה': 'Date & Time',
  'הכנסת שדה תאריך ושעה שמתעדכן, בפורמט יום/חודש/שנה שעה:דקות':
    'Insert an auto-updating date and time field, in day/month/year hours:minutes format',
  'יום/חודש/שנה ושעה': 'Day/Month/Year and time',
  'תאריך בלבד': 'Date Only',
  'יום/חודש/שנה, בלי שעה': 'Day/Month/Year, without time',
  'שעה בלבד': 'Time Only',
  'שעה בלבד, בלי תאריך': 'Time only, without date',
  'עדכן שדות': 'Update Fields',
  'חישוב מחדש של כל השדות במסמך, כמו F9 ב-Word':
    'Recalculate all fields in the document, like F9 in Word',
  'אין במסמך שדות לעדכן': 'There are no fields in the document to update',
  'סימנייה': 'Bookmark',
  'סימון הפסקה שבה הסמן בשם, לניווט ולהפניות מתוך Word':
    'Mark the paragraph where the cursor is with a name, for navigation and cross-references from Word',
  'התחל בעמוד חדש': 'Page Break',
  'הפסקה שבה הסמן תתחיל בראש עמוד חדש':
    'The paragraph where the cursor will start at the top of a new page',
  'הפסקה הזאת כבר מתחילה בעמוד חדש. לחיצה תבטל זאת':
    'This paragraph already starts on a new page. Clicking undoes that',

  // ── לשונית „פריסה” ─────────────────────────────────────────────────
  'הגדרת עמוד': 'Page Setup',
  'שוליים': 'Margins',
  'הגדרת שולי הדף (רגיל, צר, רחב)': 'Set the page margins (normal, narrow, wide)',
  // פריטי התפריטים עצמם (engine/page-setup.ts): תווית ורמז לכל ברירה.
  'צר': 'Narrow',
  'רחב': 'Wide',
  '2.54 ס"מ מכל צד': '2.54 cm on every side',
  '1.27 ס"מ מכל צד': '1.27 cm on every side',
  '2.54 ס"מ למעלה ולמטה, 5.08 בצדדים': '2.54 cm top and bottom, 5.08 at the sides',
  'כיוון': 'Orientation',
  'כיוון הדף: לאורך או לרוחב': 'Page orientation: portrait or landscape',
  'לאורך': 'Portrait',
  'הדף גבוה מרוחבו': 'The page is taller than it is wide',
  'לרוחב': 'Landscape',
  'הדף רחב מגובהו': 'The page is wider than it is tall',
  'גודל': 'Size',
  'בחירת גודל נייר (A4, Letter)': 'Choose paper size (A4, Letter)',
  '21 × 29.7 ס"מ': '21 × 29.7 cm',
  '21.6 × 27.9 ס"מ': '21.6 × 27.9 cm',
  '14.8 × 21 ס"מ': '14.8 × 21 cm',
  'פיצול הטקסט לשתי עמודות או יותר': 'Split the text into two or more columns',
  'אחת': 'One',
  'שתיים': 'Two',
  'שתי עמודות שוות': 'Two equal columns',
  'שלוש': 'Three',
  'שלוש עמודות שוות': 'Three equal columns',
  'מקטע': 'Section',
  'קשר לקודם': 'Link to Previous',
  'אין במסמך מקטע נוסף — הקישור נוגע רק במקטעים שאחרי הראשון':
    'There is no additional section — linking only affects sections after the first',
  'הכותרות של המקטעים הבאים יהיו זהות לאלה של המקטע שלפניהם':
    'The headers of the following sections will match those of the section before them',
  'גבולות עמוד': 'Page Borders',
  'מסגרת סביב העמוד': 'A border around the page',
  'ללא גבול': 'No border',
  'הסרת הגבול מהעמוד': 'Remove the border from the page',
  'קו יחיד': 'Single line',
  'חצי נקודה': 'Half a point',
  'קו עבה': 'Thick line',
  'שלוש נקודות': 'Three points',
  'קו כפול': 'Double line',
  'שני קווים דקים': 'Two thin lines',
  'מקווקו': 'Dashed',
  'קו מקווקו דק': 'A thin dashed line',
  'מנוקד': 'Dotted',
  'קו מנוקד דק': 'A thin dotted line',
  'מספרי שורות': 'Line Numbers',
  'מספור השורות בשולי הדף': 'Number the lines in the page margin',
  'ללא': 'None',
  'בלי מספרי שורות': 'No line numbers',
  'רציף': 'Continuous',
  'המספור נמשך לאורך כל המסמך': 'The numbering runs through the whole document',
  'התחל מחדש בכל עמוד': 'Restart on every page',
  'כל עמוד מתחיל מ-1': 'Every page starts at 1',
  'התחל מחדש בכל מקטע': 'Restart in every section',
  'כל מקטע מתחיל מ-1': 'Every section starts at 1',
  'מספור עמודים': 'Page Numbering',
  'תבנית מספרי העמודים ומספר ההתחלה': 'The page number format and starting number',
  'יישור אנכי': 'Vertical Alignment',
  'מיקום הטקסט בגובה העמוד': 'Positioning of the text along the page height',
  'למעלה': 'Top',
  'הטקסט מתחיל בראש העמוד': 'The text starts at the top of the page',
  'הטקסט ממורכז בין הכותרות': 'The text is centered between the headers',
  'מיושר': 'Justified',
  'הפסקאות נפרשות על גובה העמוד': 'The paragraphs spread over the page height',
  'למטה': 'Bottom',
  'הטקסט צמוד לתחתית העמוד': 'The text sits at the bottom of the page',
  'ברירות מחדל': 'Defaults',
  'גופן וגודל ברירת המחדל של המסמך כולו': 'Default font and size for the whole document',
  'עמודים': 'Pages',
  'כותרת עליונה ותחתונה': 'Header & Footer',
  'מרחק הכותרת': 'Header Distance',
  'מרחק הכותרת העליונה והתחתונה מקצה הדף':
    'Distance of the header and footer from the edge of the page',
  'שונה בעמוד ראשון': 'Different First Page',
  'לעמוד הראשון תהיה כותרת משלו': 'The first page will have its own header',
  'שונה בעמודים זוגיים ואי-זוגיים': 'Different Odd & Even Pages',
  'כותרת אחת לעמודים הזוגיים ואחרת לאי-זוגיים':
    'One header for even pages and another for odd pages',

  // ── לשונית „הפניות” ────────────────────────────────────────────────
  'תוכן עניינים': 'Table of Contents',
  'יצירת תוכן עניינים אוטומטי': 'Create an automatic table of contents',
  'הוספת תוכן עניינים למסמך': 'Insert a table of contents into the document',
  'סמן ערך': 'Mark Entry',
  'סימון טקסט שייכנס לתוכן העניינים': 'Mark text to be included in the table of contents',
  'עדכן טבלה': 'Update Table',
  'בניית תוכן העניינים מחדש מהכותרות שבמסמך':
    'Rebuild the table of contents from the headings in the document',
  'אין במסמך תוכן עניינים לעדכן': 'There is no table of contents in the document to update',
  'התאמה אישית': 'Customize',
  'רמות הכותרות שייכללו, והאם הערכים יהיו קישורים':
    'Which heading levels to include, and whether entries will be links',
  'אין במסמך תוכן עניינים להתאים':
    'There is no table of contents in the document to customize',
  'הסר': 'Remove',
  'מחיקת תוכן העניינים מהמסמך': 'Delete the table of contents from the document',
  'אין במסמך תוכן עניינים להסיר': 'There is no table of contents in the document to remove',
  'הערות שוליים': 'Footnotes',
  'הערת שוליים': 'Footnote',
  'הוספת הערת שוליים בתחתית העמוד': 'Insert a footnote at the bottom of the page',
  'הערת סיום': 'Endnote',
  'הוספת הערת סיום בסוף המסמך': 'Insert an endnote at the end of the document',
  'נהל הערות': 'Manage Notes',
  'עריכה והסרה של הערות השוליים והערות הסיום שבמסמך':
    'Edit and remove the footnotes and endnotes in the document',
  'אין במסמך הערות לנהל': 'There are no notes in the document to manage',
  'פעולה על הערה עדיין בעבודה — ההוספה תיפתח כשהיא תסתיים':
    'A note operation is still running — management opens when it finishes',
  'הפניות מקושרות': 'Cross-references',
  'עדכן הפניות': 'Update References',
  'חישוב מחדש של ההפניות המקושרות במסמך':
    'Recalculate the linked references in the document',
  'אין במסמך הפניות מקושרות לעדכן':
    'There are no linked references in the document to update',
  'מפתח': 'Index',
  'סמן ערך למפתח': 'Mark Index Entry',
  'סימון הטקסט שנבחר כערך במפתח': 'Mark the selected text as an index entry',
  'הוסף מפתח': 'Insert Index',
  'הוספת מפתח הערכים בסוף המסמך': 'Insert the index at the end of the document',
  'עדכן מפתח': 'Update Index',
  'בניית המפתח מחדש מהערכים שסומנו במסמך':
    'Rebuild the index from the entries marked in the document',
  'אין במסמך מפתח לעדכן': 'There is no index in the document to update',
  'הגדרות מפתח': 'Index Settings',
  'מספר הטורים של המפתח, והאם תת-הערכים רצופים':
    'The index column count, and whether sub-entries run on',
  'אין במסמך מפתח להתאים': 'There is no index in the document to customize',
  'הסר מפתח': 'Remove Index',
  'מחיקת המפתח מהמסמך. הערכים שסומנו נשארים':
    'Delete the index from the document. Marked entries remain',
  'אין במסמך מפתח להסיר': 'There is no index in the document to remove',
  'ציטוטים וביבליוגרפיה': 'Citations & Bibliography',
  'ציטוט מהקורא': 'Citation from the Reader',
  'הכנסת הקטע המסומן בקורא של אוצריא, עם המקור, במיקום הסמן':
    'Insert the passage selected in the Otzaria reader, with its source, at the cursor',
  'יש לפתוח מסמך שאפשר לכתוב בו': 'Open a writable document first',
  'הוסף ציטוט': 'Add Citation',
  'הוספת ציטוט למקור במקום הסמן': 'Insert a citation to a source at the cursor',
  'נהל מקורות': 'Manage Sources',
  'הוספה, עריכה ומחיקה של המקורות שבמסמך':
    'Add, edit, and delete the sources in the document',
  'ביבליוגרפיה': 'Bibliography',
  'הוספת רשימת המקורות בסוף המסמך': 'Insert the list of sources at the end of the document',
  'עדכן ביבליוגרפיה': 'Update Bibliography',
  'בניית הביבליוגרפיה מחדש מהמקורות שבמסמך':
    'Rebuild the bibliography from the sources in the document',
  'אין במסמך ביבליוגרפיה לעדכן': 'There is no bibliography in the document to update',
  'הסר ביבליוגרפיה': 'Remove Bibliography',
  'מחיקת הביבליוגרפיה מהמסמך. המקורות עצמם נשארים':
    'Delete the bibliography from the document. The sources themselves remain',
  'אין במסמך ביבליוגרפיה להסיר': 'There is no bibliography in the document to remove',
  'כיתובים': 'Captions',
  'הוסף כיתוב': 'Insert Caption',
  'הוספת כיתוב ממוספר לתמונה, לטבלה או לתרשים':
    'Insert a numbered caption for a picture, table, or chart',
  'המסמך עדיין נטען': 'The document is still loading',
  // ── לשונית „סקירה” ─────────────────────────────────────────────────
  'הגהה': 'Proofing',
  'תיקון אוטומטי': 'AutoCorrect',
  'אפשרויות התיקון האוטומטי — דלוק כרגע': 'AutoCorrect options — currently on',
  'אפשרויות התיקון האוטומטי — כבוי כרגע': 'AutoCorrect options — currently off',
  'בדיקת איות': 'Spelling Check',
  'טוען את המילון התורני…': 'Loading the Torah dictionary…',
  'כיבוי בדיקת האיות התורנית': 'Turn off Torah spelling check',
  'סימון מילים שאינן במילון התורני. לחיצה ימנית על מילה מסומנת מאפשרת להוסיף אותה למילון':
    'Mark words that are not in the Torah dictionary. Right-click a marked word to add it to the dictionary',
  'תגובה חדשה': 'New Comment',
  'הוספת תגובה — תתווסף בשלב הבא, יחד עם זהות המחבר ופאנל התגובות':
    'Adding comments — coming in the next phase, together with author identity and the comments panel',
  'תגובות': 'Comments',
  'מעקב אחר שינויים': 'Tracking',
  'עקוב אחר שינויים': 'Track Changes',
  'כיבוי מצב מעקב אחר שינויים': 'Turn off track changes',
  'הפעלת מצב מעקב אחר שינויים במסמך': 'Turn on track changes for the document',
  'שינויים': 'Changes',
  'קבל שינוי': 'Accept Change',
  'קבלת השינוי הנוכחי': 'Accept the current change',
  'דחה שינוי': 'Reject Change',
  'דחיית השינוי הנוכחי': 'Reject the current change',
  'קבל את כל השינויים': 'Accept All Changes',
  'קבלת כל השינויים במסמך': 'Accept all changes in the document',
  'דחה את כל השינויים': 'Reject All Changes',
  'דחיית כל השינויים במסמך': 'Reject all changes in the document',
  'הגנה': 'Protection',
  'הגבל עריכה': 'Restrict Editing',
  'הפעולה מתבצעת…': 'The operation is in progress…',
  'לחץ שוב לאישור: המסמך יינעל לקריאה בלבד (ניתן לביטול מכאן)':
    'Click again to confirm: the document will be locked to read-only (can be undone here)',
  'ביטול ההגבלה — המסמך יחזור לעריכה מלאה':
    'Remove restriction — the document returns to full editing',
  'ביטול ההגבלה — המסמך יחזור למצב מעקב אחר שינויים':
    'Remove restriction — the document returns to track-changes mode',
  'הצג את המסמך במצב „קריאה בלבד". ניתן לבטל מכאן בכל עת.':
    'View the document in “read-only” mode. Can be undone here at any time.',

  // ── לשונית „תצוגה” ─────────────────────────────────────────────────
  'תצוגות': 'Views',
  'הצג': 'Show',
  'מצב מיקוד': 'Focus Mode',
  'מצב קריאה ומיקוד ללא הסחות דעת': 'Reading and focus mode without distractions',
  'סרגל': 'Ruler',
  'הצג או הסתר את סרגל המידות': 'Show or hide the ruler',
  'סימני עיצוב': 'Formatting Marks',
  'הצג סימני פסקאות ותווים נסתרים': 'Show paragraph marks and hidden characters',
  'חלונית ניווט': 'Navigation Pane',
  'הצג את כותרות המסמך בצד, ועבור מכותרת לכותרת בלחיצה':
    'Show the document headings on the side, and jump between them with a click',
  'גודל אמיתי': 'Actual Size',
  'הצג את המסמך בגודלו האמיתי (100%)': 'View the document at actual size (100%)',
  'שינוי גודל תצוגה': 'Zoom',
  'התאם את תצוגת העמוד לרוחב החלון': 'Fit the page view to the window width',
  'רוחב עמוד': 'Page Width',
  'רקע': 'Background',
  'צבע רקע': 'Background Color',
  'צבע רקע העורך': 'Editor Background Color',
  // הפריט המנקה בבורר צבע הבד. „ללא צבע” של הפלטה אינו נכון שם — הניקוי
  // מחזיר את צבע ערכת הנושא, ולא מסיר צבע. ראו `clearLabel` ב-ColorPickerPopover.
  'ברירת מחדל': 'Default',

  // ── לשונית „אוצריא” ────────────────────────────────────────────────
  'אוצריא': 'Otzaria',
  'פתח ספרייה': 'Open Library',
  'פתיחת ספריית הספרים של אוצריא': 'Open the Otzaria book library',
  'זמין רק כשהעורך פועל בתוך אוצריא': 'Available only when the editor runs inside Otzaria',
  'חיפוש באוצריא': 'Search in Otzaria',
  'חיפוש הטקסט המסומן במסמך בכל ספריות אוצריא':
    'Search the selected text in all Otzaria libraries',
  'יש לפתוח מסמך ולסמן בו את הטקסט לחיפוש':
    'Open a document and select the text to search',
  'השלמה מהספר': 'Complete from Book',
  'הקלטה': 'Dictate',
  'דיבור': 'Voice',
  'להורדה': 'Download',
  'ביטול': 'Cancel',
  'להקלדה בדיבור צריך את „שירות הדיבור של אוצריא”, תוכנה נפרדת שמתקינים פעם אחת. הזיהוי נעשה במחשב, בלי אינטרנט.':
    'Voice typing needs the Otzaria voice service, a separate program installed once. Recognition runs on this computer, offline.',
  'המסמך פתוח לקריאה בלבד': 'The document is open read-only',
  'עצור דיבור': 'Stop Dictation',
  'בטל הקלטה': 'Cancel Dictation',
  'ביטול התחלת ההקלטה': 'Cancel starting dictation',
  'יש לפתוח מסמך ולהציב בו את הסמן': 'Open a document and place the cursor in it',
  'עצירת ההקלטה אחרי המשפט האחרון': 'Stop dictating after the last sentence',
  'הקלדה בדיבור במיקום הסמן. הזיהוי נעשה במחשב, בלי אינטרנט':
    'Type by voice at the cursor. Recognition runs on this computer, offline',
  'בזמן הקלדה, אם הטקסט תואם את הספר הפתוח בקורא — Tab משלים 5 מילים מהמקור':
    'While typing, if the text matches the book open in the reader — Tab completes five words from the source',
  'יש לפתוח מסמך תחילה': 'Open a document first',

  // ── לשונית „מפתחים” ────────────────────────────────────────────────
  'מאקרו': 'Macros',
  'ניהול מאקרו': 'Manage Macros',
  'רשימת המאקרו, קטעי הטקסט והסקריפטים — הרצה, עריכה ושיתוף':
    'The list of macros, text snippets and scripts — run, edit and share',
  'הקלט מאקרו': 'Record Macro',
  'עצור הקלטה': 'Stop Recording',
  'מקליט את הפעולות במסמך — הקלדה, עיצוב, רשימות — לניגון חוזר':
    'Records the actions in the document — typing, formatting, lists — for replay',
  'נגן אחרון': 'Play Last',
  'מריץ את המאקרו האחרון שהוקלט, מהמקום שבו הסמן עומד':
    'Runs the last recorded macro from wherever the cursor stands',

  // ── בורר הצבעים (ColorPickerPopover) ───────────────────────────────
  'בחירת צבע': 'Pick a color',
  'ללא צבע': 'No color',
  'צבעי ערכת נושא': 'Theme Colors',
  'צבעים רגילים': 'Standard Colors',
  'צבעים נוספים...': 'More colors...',
  // „כחול, גוון 3”: `shadeName` מרכיב את השם משם המשפחה ומהמילה „גוון”, וכל
  // חלק מתורגם בנפרד — מחרוזת מורכבת לא הייתה מתאימה לשום מפתח.
  'גוון': 'Shade',
  'לבן ואפור בהיר': 'White and light gray',
  'שחור ואפור כהה': 'Black and dark gray',
  'חום בהיר': 'Light brown',
  'כחול כהה': 'Dark blue',
  'כחול': 'Blue',
  'אדום': 'Red',
  'ירוק זית': 'Olive green',
  'סגול': 'Purple',
  'טורקיז': 'Turquoise',
  'כתום': 'Orange',
  'אדום כהה': 'Dark red',
  'צהוב': 'Yellow',
  'ירוק בהיר': 'Light green',
  'ירוק': 'Green',
  'תכלת': 'Light blue',

  // ── גלריית הסגנונות (StyleGallery / engine/style-gallery) ─────────
  'רגיל': 'Normal',
  'ללא מרווח': 'No Spacing',
  'גוף טקסט': 'Body Text',
  'כותרת': 'Title',
  'כותרת משנה': 'Subtitle',
  'כותרת 1': 'Heading 1',
  'כותרת 2': 'Heading 2',
  'כותרת 3': 'Heading 3',
  'כותרת 4': 'Heading 4',
  'כותרת 5': 'Heading 5',
  'כותרת 6': 'Heading 6',
  'כותרת 7': 'Heading 7',
  'כותרת 8': 'Heading 8',
  'כותרת 9': 'Heading 9',
  'ציטוט': 'Quote',
  'ציטוט מודגש': 'Intense Quote',
  'פסקת רשימה': 'List Paragraph',
  'כתובית': 'Caption',
  'חזק': 'Strong',
  'הדגשה': 'Emphasis',
  'הדגשה עדינה': 'Subtle Emphasis',
  'הדגשה מודגשת': 'Intense Emphasis',
  'הפניה עדינה': 'Subtle Reference',
  'הפניה מודגשת': 'Intense Reference',
  'שם ספר': 'Book Title',
  'טקסט הערת שוליים': 'Footnote Text',
  'טקסט הערת סיום': 'Endnote Text',
  'טקסט הערה': 'Comment Text',
  'היפר-קישור': 'Hyperlink',
  'הסגנונות הקודמים': 'Previous styles',
  'הסגנונות הבאים': 'Next styles',

  /* לשונית „שולחן העורך” */
  'שולחן העורך': "Editor's Desk",
  'שגיאות מצויות': 'Common Mistakes',
  'סוגריים לא סגורים': 'Unclosed Brackets',
  'טקסט מתחלף': 'Alternating Text',
  'סוגריים ⟵ הערות': 'Brackets → Footnotes',
  'הערות ⟵ סוגריים': 'Footnotes → Brackets',
  'עיצוב פסקה': 'Paragraph Design',
  'מילה ראשונה': 'First Word',
  'מרווח שורות אחיד': 'Uniform Line Spacing',
  'בטל מרווח אחיד': 'Undo Uniform Spacing',
  'אחידות מסמך': 'Document Consistency',
  'גודל עמוד ושוליים': 'Page Size & Margins',
  'רוחב טורים': 'Column Widths',
  'תיקון שגיאות הקלדה נפוצות בכל המסמך — רווחים כפולים, פיסוק, סוגריים':
    'Fix common typing errors throughout the document — double spaces, punctuation, brackets',
  'סריקת המסמך אחר סוגריים עגולים ומרובעים שאינם מאוזנים':
    'Scan the document for unbalanced round and square brackets',
  'הדגשת דיבור-המתחיל בפסקאות המסומנות — מתחילת הפסקה עד תו הסיום, ובין תו ההתחלה לתו הסיום':
    'Bold the opening phrase in the selected paragraphs — from the start of the paragraph to the end character, and between the start and end characters',
  'כל קטע בסוגריים בפסקאות המסומנות הופך להערת שוליים במקומו':
    'Each bracketed passage in the selected paragraphs becomes a footnote in its place',
  'תוכן הערות השוליים שבבחירה חוזר לגוף הטקסט בסוגריים, במקום ההפניה':
    'The content of the footnotes in the selection returns to the body in brackets, where the reference stood',
  'תיקון העתקה': 'Paste Cleanup',
  'רווחים קשיחים שהגיעו מהדבקה מתוכנות אחרות הופכים לרווחים רגילים, בפסקאות המסומנות':
    'Non-breaking spaces pasted from other programs become regular spaces, in the selected paragraphs',
  'סוג הסוגריים להמרה':
    'Bracket type to convert',
  'הגדלה והדגשה של המילה הראשונה בכל פסקה מסומנת, פרופורציונלית לגוף הפסקה':
    'Enlarge and bold the first word of each selected paragraph, in proportion to the paragraph body',
  'קיבוע מרווח „בדיוק” בגובה שורה של גופן הגוף — שהמילה המוגדלת לא תמתח את השורה הראשונה. מומלץ להריץ לפני עיצוב המילה הראשונה':
    'Fix line spacing to "exact" at the body font line height, so an enlarged first word does not stretch the first line. Best run before First Word',
  'החזרת מרווח „בדיוק” למרווח „מרובה” שקול בפסקאות המסומנות':
    'Convert "exact" line spacing back to an equivalent "multiple" in the selected paragraphs',
  'איתור מקטעים שסטו מגודל העמוד או מהשוליים של שאר המסמך, והשוואתם לפרופיל אחד':
    'Find sections that deviate from the page size or margins of the rest of the document, and apply a single profile',
  'איתור מקטעים מרובי-טורים שרוחב הטורים או המרווח ביניהם שונה, והשוואתם':
    'Find multi-column sections whose column width or gap differs, and make them uniform',
  'אחידות גודל עמוד ושוליים':
    'Page Size & Margin Consistency',
  'אחידות רוחב טורים':
    'Column Width Consistency',
  'נמצאו במסמך כמה פרופילים של גודל עמוד ושוליים. יש לבחור את הנכון — והוא יוחל על כל המקטעים:':
    'Several page size and margin profiles were found in the document. Choose the correct one — it will be applied to all sections:',
  'נמצאו במסמך כמה פרופילים של טורים. יש לבחור את הנכון — והוא יוחל על כל המקטעים מרובי-הטורים:':
    'Several column profiles were found in the document. Choose the correct one — it will be applied to all multi-column sections:',
  'גודל העמוד והשוליים הוחלו על כל המקטעים':
    'Page size and margins applied to all sections',
  'רוחב הטורים הוחל על כל המקטעים מרובי-הטורים':
    'Column widths applied to all multi-column sections',
  'הפעולה נכשלה':
    'The action failed',
  'עמודים ודפוס': 'Pages & Print',
  'צמצום מסמך': 'Shrink Document',
  'צמצום המסמך ליעד עמודים — שוליים, ריווח פסקאות, מרווח שורות וגופן, בסבבים של 10%':
    'Shrink the document to a target page count — margins, paragraph spacing, line spacing and font, in 10% rounds',
  'סמן עמודים': 'Mark Pages',
  'סימון המילה הראשונה והאחרונה של כל עמוד, ושמירת תצלום של שבירות העמודים לבדיקה מאוחרת':
    'Mark the first and last word of every page, and save a snapshot of the page breaks for later checking',
  'בדוק עמודים': 'Check Pages',
  'בדיקה אילו עמודים נפתחים היום במקום אחר מאשר בסימון האחרון':
    'Check which pages now start somewhere else than in the last marking',
  'הסר סימון': 'Unmark',
  'כיבוי סימון העמודים ומחיקת התצלום השמור':
    'Turn page marking off and delete the saved snapshot',
  'סימני חיתוך': 'Crop Marks',
  'הגדלת הדף והשוליים במילימטרים לבחירתך והוספת סימני חיתוך בפינות — להדפסה ול-PDF':
    'Enlarge the page and margins by your chosen millimetres and add corner crop marks — for print and PDF',
  'פירוק מסמך': 'Split Document',
  'העברת כל הערות השוליים למסמך חדש בטאב נפרד; בגוף נשאר מספר ההערה בכתב עילי':
    'Move all footnotes to a new document in a separate tab; the note number stays in the body as superscript',
  'סימון העמודים נכשל: אין מסמך פתוח, או שהמסמך אינו תומך בפעולה':
    'Page marking failed: no document is open, or the document does not support the action',
  'בדיקת העמודים נכשלה: אין מסמך פתוח, או שהמסמך אינו תומך בפעולה':
    'Page check failed: no document is open, or the document does not support the action',
  'בדיקת העמודים נכשלה: אין סימון שמור למסמך זה — יש לסמן תחילה':
    'Page check failed: no saved marking for this document — mark it first',
  'סימון העמודים הוסר': 'Page marking removed',
  'מתחיל…': 'Starting…',
  'הגופן אינו מותקן במכונה — אין דגימה להציג':
    'The font is not installed on this machine — no sample to show',
  'הגופן מותקן אך הדפדפן אינו מצייר אותו — אין דגימה להציג':
    'The font is installed but the browser cannot render it — no sample to show',
};

/**
 * המחרוזת שתוצג: התרגום לאנגלית כששפת המשתמש אנגלית, ואחרת המקור העברי.
 *
 * נקראת מהתבניות ומה־computed של רכיבי הרצועה; הקריאה קוראת את `locale`
 * בתוך ה-render, ולכן הממשק מתעדכן גם אם השפה תיקבע אחרי ההרכבה הראשונה.
 */
export function menuString(hebrew: string): string {
  if (locale.value !== 'en') return hebrew;
  return EN[hebrew] ?? hebrew;
}
