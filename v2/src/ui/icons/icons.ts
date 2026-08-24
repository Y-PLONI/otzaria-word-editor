/**
 * ספריית אייקוני ה-SVG של הסרגל, בסגנון Microsoft Word / Fluent UI.
 *
 * המיגרציה ל-Fluent System Icons (Fluent 2) נעשית לשונית-לשונית. עד כה
 * הועברה לשונית "קובץ" — newDoc, folder, save, saveAs, export, print, info —
 * ובה ה-path data הוא של Microsoft, בגרסת regular בגריד 20. כל השאר עדיין
 * מצויר בבית, במשקל קו כבד יותר; לכן `save` שיושב גם בסרגל הגישה המהירה
 * ייראה קל יותר מ-undo/redo/search שלצידו עד שגם הם יעברו.
 *
 * שלוש מגבלות שמחייבות את כל האייקונים כאן, ו-tests/unit/icons.test.ts אוכף
 * אותן על הנתונים האמיתיים:
 *
 * 1. `viewBox="0 0 20 20"` בדיוק, בלי יוצא מן הכלל. ה-SvgIcon נותן לאייקון
 *    גודל בפיקסלים, ולכן גריד שונה בין אייקונים = עובי קווים שונה באותה שורה
 *    בסרגל (נמדד: `paste` על גריד 24 נראה דק מהשכנים שלו בקבוצת "לוח").
 * 2. אף נקודה על ה-path אינה חורגת מה-viewBox. חריגה אינה נראית כשגיאה —
 *    היא פשוט נחתכת בשקט (נמדד: `dirRtl`/`dirLtr` חרגו 1.1 יחידות מעל הגבול
 *    העליון, ו-`cut` 1.3 מתחת לתחתון).
 * 3. יחס מילוי אחיד: המידה הדומיננטית של ה-bounding box תופסת 70%–85%
 *    מה-viewBox. בלי זה מתקבלת שורה שבה שכנים נראים בגדלים שונים (נמדד:
 *    `reject` תפס 47%x47% ליד `footnote` שתפס 90%x75%). התקרה היא 85% ולא
 *    84% כי זה הגריד של Fluent עצמו: גליף מלבני כמו `info` יושב 2–18
 *    (80%), אבל דף עם תג כמו `document_add` יושב 2–19 לגובה ו-1–16 לרוחב.
 *
 * כל האייקונים בנויים ממילוי (`fill="currentColor"`) ולא מקווי stroke, כדי
 * שהמשקל האופטי יהיה נשלט במידות ה-path עצמו — וכדי שבדיקת ה-bounding box
 * תהיה מדויקת ולא תצטרך להעריך את התרחבות הקו.
 *
 * אייקונים ללא צרכן: `lineSpacing`, `borders` ו-`shading` הם אייקוני
 * Word-parity לקבוצת "פיסקה", ו-`chevronLeft`/`chevronRight`/`chevronUp`
 * נדרשים לגלילת גלריית הסגנונות ולכיווץ הסרגל. הם מוגדרים כאן ויחוברו
 * לפקדים בקומיטים נפרדים; הבדיקה מחריגה אותם דרך רשימת PLANNED_ICONS
 * כדי שהחרגה לא תתרחב בשקט לאייקונים שנשכחו.
 */

/*!
 * Fluent System Icons — MIT
 *
 * The SVG path data of the icons marked "Fluent System Icons" in this file
 * is taken from Fluent System Icons by Microsoft
 * (https://github.com/microsoft/fluentui-system-icons), as published in the
 * npm package @fluentui/svg-icons@1.1.338. Only path data was copied; the
 * <svg> wrapper, the icon names and the rest of this file are the plugin's
 * own code.
 *
 * MIT License
 *
 * Copyright (c) 2020 Microsoft Corporation
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 * @license MIT
 */

/** אייקונים שמוגדרים לפני שיש להם צרכן. ראו הערת הפתיחה. */
export const PLANNED_ICONS = ['lineSpacing', 'borders', 'shading'] as const;

export const ICONS: Record<string, string> = {
  // מותג ואפליקציה
  // ה-W הוא תת-path בכיוון הפוך למסגרת, ולכן הוא חור (nonzero winding) ולא
  // צורה מלאה. היפוך כיוון של אחד השניים "יסתום" את הלוגו.
  word: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M16.5 2.5h-13A1.5 1.5 0 0 0 2 4v12a1.5 1.5 0 0 0 1.5 1.5h13a1.5 1.5 0 0 0 1.5-1.5V4a1.5 1.5 0 0 0-1.5-1.5zM13.5 14.5h-1.9L9.5 8.1l-2 6.4H5.7L4 5.5h2.1l1.1 5.6L9.1 5.5h1.8l1.9 5.6 1.1-5.6H16l-1.7 9z"/></svg>`,
  // הלוגו של אוצריא הוא **ספר פתוח**: שתי דפים משופעים שנפגשים בשדרה, ובתוכם
  // שורות טקסט (assets/icon/iconnew.png, web/icons/Icon-192.png). מה שהיה כאן
  // הוא משולש חלול עם מקף אנכי ונקודה מתחתיו — כלומר משולש + סימן קריאה, שהוא
  // **אייקון האזהרה המוסכם** בכל סט אייקונים, על כפתור „פתח ספרייה”. לכן
  // tests/unit/icons.test.ts אוסר משולש שממלא את ה-viewBox על כל אייקון בסט:
  // בדיקה גאומטרית אינה יכולה לאמת „נראה כמו ספר”, אבל היא כן יכולה לאמת
  // „אינו סימן אזהרה”.
  //
  // הבנייה: לכל דף טבעת (מסגרת חיצונית בכיוון אחד, פנימית בכיוון ההפוך =
  // חור ב-nonzero winding), ושורות הטקסט בכיוון החיצוני כדי שיתמלאו בתוך
  // החור. היפוך כיוון של אחת הטבעות „יסתום” את הדף לכתם מלא.
  otzaria: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2.2 5L10 7v9l-7.8-2zM3.5 6.6v6.4l5.5 1.4V8zM10 7l7.8-2v9l-7.8 2zM16.5 6.6L11 8v6.4l5.5-1.4zM4.5 9.2h3.4v1H4.5zM4.5 11.4h3.4v1H4.5zM12.1 9.2h3.4v1h-3.4zM12.1 11.4h3.4v1h-3.4z"/></svg>`,

  // קובץ: מסמך חדש, פתיחה, שמירה, ייצוא. שבעת האייקונים של לשונית "קובץ"
  // — כולל print ו-info שיושבים למטה — הם Fluent System Icons מקוריים,
  // גרסת regular בגריד 20. ראו את באנר הרישוי בראש הקובץ.
  //
  // `document_add`: דף עם קיפול בפינה הימנית-העליונה, ותג "+" בעיגול
  // בפינה השמאלית-התחתונה. ל-Microsoft אין גרסת RTL לאייקון הזה, והוא
  // נשאר לא משוקף — כמו אייקוני המסמך ב-Word בעברית.
  newDoc: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M6 2a2 2 0 0 0-2 2v5.2q.49-.13 1-.18V4a1 1 0 0 1 1-1h4v3.5c0 .83.67 1.5 1.5 1.5H15v8a1 1 0 0 1-1 1h-3.6q-.27.54-.66 1H14a2 2 0 0 0 2-2V7.41c0-.4-.16-.78-.44-1.06l-3.91-3.91A1.5 1.5 0 0 0 10.59 2zm8.8 5h-3.3a.5.5 0 0 1-.5-.5V3.2zM10 14.5a4.5 4.5 0 1 1-9 0 4.5 4.5 0 0 1 9 0m-4-2a.5.5 0 0 0-1 0V14H3.5a.5.5 0 0 0 0 1H5v1.5a.5.5 0 0 0 1 0V15h1.5a.5.5 0 0 0 0-1H6z"/></svg>`,
  // `folder_open`: המעטפת נמדדת 85% רוחב, אך הדיו נעצר ב-x=18 — ההפרש הוא
  // נקודות בקרה של בזייה, שבדיקת המעטפת סופרת בכוונה כגבול שמרני.
  folder: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M3 5.5v6.6l1.5-2.6A3 3 0 0 1 7.1 8H15v-.5c0-.83-.67-1.5-1.5-1.5h-4a.5.5 0 0 1-.35-.15l-1.71-1.7A.5.5 0 0 0 7.09 4H4.5C3.67 4 3 4.67 3 5.5m1.28 10.48.22.02h9.4a2 2 0 0 0 1.73-1l2.17-3.75A1.5 1.5 0 0 0 16.5 9H7.1a2 2 0 0 0-1.73 1L3.2 13.75a1.5 1.5 0 0 0 1.08 2.23M2 14.46V5.5A2.5 2.5 0 0 1 4.5 3h2.59c.4 0 .78.16 1.06.44L9.7 5h3.79A2.5 2.5 0 0 1 16 7.5V8h.5a2.5 2.5 0 0 1 2.16 3.75L16.5 15.5a3 3 0 0 1-2.6 1.5H4.5a2.5 2.5 0 0 1-1.62-.6A2.5 2.5 0 0 1 2 14.46"/></svg>`,
  save: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M3 5c0-1.1.9-2 2-2h8.38a2 2 0 0 1 1.41.59l1.62 1.62A2 2 0 0 1 17 6.62V15a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2zm2-1a1 1 0 0 0-1 1v10a1 1 0 0 0 1 1v-4.5c0-.83.67-1.5 1.5-1.5h7c.83 0 1.5.67 1.5 1.5V16a1 1 0 0 0 1-1V6.62a1 1 0 0 0-.3-.7L14.1 4.28a1 1 0 0 0-.71-.29H13v2.5c0 .83-.67 1.5-1.5 1.5h-4A1.5 1.5 0 0 1 6 6.5V4zm2 0v2.5c0 .28.22.5.5.5h4a.5.5 0 0 0 .5-.5V4zm7 12v-4.5a.5.5 0 0 0-.5-.5h-7a.5.5 0 0 0-.5.5V16z"/></svg>`,
  // "שמור בשם" = `save_edit`, דיסקט עם עיפרון — אותה כוונה שהייתה לאייקון
  // המצויר שקדם לו, עכשיו בגריד ובמשקל הקו של Fluent.
  saveAs: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M5 3a2 2 0 0 0-2 2v10c0 1.1.9 2 2 2h3l.06-.35.16-.65H6v-4.5c0-.28.22-.5.5-.5h5.44l1-1H6.5c-.83 0-1.5.67-1.5 1.5V16a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1h1v2.5C6 7.33 6.67 8 7.5 8h4c.83 0 1.5-.67 1.5-1.5V4h.38a1 1 0 0 1 .7.3l1.63 1.61a1 1 0 0 1 .29.71V8q.52-.02 1 .13v-1.5a2 2 0 0 0-.59-1.42L14.8 3.59A2 2 0 0 0 13.38 3zm2 3.5V4h5v2.5a.5.5 0 0 1-.5.5h-4a.5.5 0 0 1-.5-.5m7.8 3.05-4.82 4.83a2 2 0 0 0-.58 1.02l-.37 1.5a.9.9 0 0 0 1.08 1.07l1.5-.37q.58-.16 1.01-.58l4.83-4.83a1.87 1.87 0 0 0-2.64-2.64"/></svg>`,
  // `arrow_export_rtl` ולא `arrow_export_ltr`: הרצועה RTL, ולכן "החוצה"
  // הוא שמאלה. שתי הגרסאות קיימות אצל Microsoft בדיוק בשביל ההבחנה הזאת.
  export: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M17.5 4a.5.5 0 0 0-.5.5v11a.5.5 0 0 0 1 0v-11a.5.5 0 0 0-.5-.5M15 10a.5.5 0 0 0-.5-.5H3.7l3.15-3.15a.5.5 0 0 0-.7-.7l-4 4a.5.5 0 0 0 0 .7l4 4a.5.5 0 0 0 .7-.7L3.71 10.5H14.5a.5.5 0 0 0 .5-.5"/></svg>`,

  // סרגל גישה מהירה
  undo: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M10.5 6.5A6.5 6.5 0 0 0 5 9.1L2.5 6.5v7h7l-2.8-2.8A4.8 4.8 0 0 1 15 13.5l1.6-1.2A6.5 6.5 0 0 0 10.5 6.5z"/></svg>`,
  redo: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M9.5 6.5A6.5 6.5 0 0 1 15 9.1l2.5-2.6v7h-7l2.8-2.8A4.8 4.8 0 0 0 5 13.5L3.4 12.3A6.5 6.5 0 0 1 9.5 6.5z"/></svg>`,
  search: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M8.5 3a5.5 5.5 0 0 1 4.3 8.9l4.4 4.4-1.3 1.3-4.4-4.4A5.5 5.5 0 1 1 8.5 3zm0 1.8a3.7 3.7 0 1 0 0 7.4 3.7 3.7 0 0 0 0-7.4z"/></svg>`,
  chevronDown: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M5.3 7.3L10 12l4.7-4.7 1.3 1.4-6 6-6-6 1.3-1.4z"/></svg>`,
  chevronUp: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M5.3 12.7L10 8l4.7 4.7 1.3-1.4-6-6-6 6 1.3 1.4z"/></svg>`,
  // chevron אופקי מצויר ולא מסובב ב-CSS: transform: rotate על ה-span משנה גם
  // את תיבת המיקוד ואת מרכז הסיבוב מול ריפוד לא סימטרי, וזה נראה קפוץ בגלילה.
  chevronLeft: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M12.7 5.3L8 10l4.7 4.7-1.4 1.3-6-6 6-6 1.4 1.3z"/></svg>`,
  chevronRight: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M7.3 5.3L12 10l-4.7 4.7 1.4 1.3 6-6-6-6-1.4 1.3z"/></svg>`,

  // קבוצת לוח (Clipboard)
  paste: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M7.4 6.1V3.9h1.2a1.4 1.4 0 0 1 2.8 0h1.2v2.2zM3.5 4.5h13v1.6h-13zM3.5 15.9h13v1.6h-13zM3.5 6.1h1.6v9.8H3.5zM14.9 6.1h1.6v9.8h-1.6z"/></svg>`,
  cut: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M7 2.5L5.6 3.6l8 10.2 1.4-1.1zM13 2.5l1.4 1.1-8 10.2-1.4-1.1zM5.7 12.4a2.4 2.4 0 1 1 0 4.8 2.4 2.4 0 0 1 0-4.8zM5.7 13.6a1.2 1.2 0 1 0 0 2.4 1.2 1.2 0 0 0 0-2.4zM14.3 12.4a2.4 2.4 0 1 1 0 4.8 2.4 2.4 0 0 1 0-4.8zM14.3 13.6a1.2 1.2 0 1 0 0 2.4 1.2 1.2 0 0 0 0-2.4z"/></svg>`,
  copy: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M13 2.5H4A1.6 1.6 0 0 0 2.4 4.1V13H4V4.1H13V2.5zM5.5 6h12v1.6h-12zM5.5 15.9h12v1.6h-12zM5.5 7.6h1.6v8.3H5.5zM15.9 7.6h1.6v8.3h-1.6z"/></svg>`,
  formatPainter: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M15.5 2.5h-10a1.5 1.5 0 0 0-1.5 1.5v3.5a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V4a1.5 1.5 0 0 0-1.5-1.5zM5.8 4.2h8.4v1.6H5.8V4.2zM7.5 10h1.8v6.5a1 1 0 0 0 1 1h.4a1 1 0 0 0 1-1V10h1.8V8.5H7.5V10z"/></svg>`,

  // קבוצת גופן (Font)
  bold: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M4.4 3h5.6a3.55 3.55 0 0 1 2.48 6.03A3.98 3.98 0 0 1 10.65 17H4.4V3zm2.58 2.37v3.87h3.02a1.51 1.51 0 0 0 1.51-1.51 1.51 1.51 0 0 0-1.51-1.51H6.98v-.85zm0 6.24v3.87h3.66a1.94 1.94 0 0 0 1.94-1.94 1.94 1.94 0 0 0-1.94-1.94H6.98z"/></svg>`,
  italic: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M7.5 3h8v1.9h-2.7l-3.6 10.4h2.7v1.9H4.5v-1.9h2.7L10.8 4.9H7.5V3z"/></svg>`,
  underline: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M10 13.5a4 4 0 0 0 4-4V3.5h-2v6a2 2 0 0 1-4 0v-6H6v6a4 4 0 0 0 4 4zm-6 2.5h12v1.6H4V16z"/></svg>`,
  strikethrough: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2.5 10.5h15v1.6h-15v-1.6zm4.8-1.7a2.8 2.8 0 0 1 2.5-1.5c1.4 0 2.2.8 2.2 2.2v4.8h-1.6v-1.2A2.7 2.7 0 0 1 8 14.5c-1.5 0-2.5-.9-2.5-2.2 0-1.5 1.2-2.3 3.6-2.5l1.1-.1v-.3c0-.6-.4-1-1.2-1a1.5 1.5 0 0 0-1.4.9l-1.3-.5zm4.4 3.7v-.9l-1.3.1c-1.2.1-1.7.5-1.7 1.2 0 .6.5 1 1.2 1 .9 0 1.8-.5 1.8-1.4z"/></svg>`,
  subscript: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M4.5 4.5l2.2 3.8 2.2-3.8h2.1l-3.3 5 3.5 5.5H9.1l-2.4-4-2.4 4H2.2l3.5-5.5-3.3-5h2.1zm11 7h2v1l-1.7 1.7c-.2.2-.3.4-.3.7 0 .4.3.6.6.6h1.8v1.2h-2.1c-.8 0-1.3-.5-1.3-1.2 0-.4.2-.7.4-1l1.6-1.6H15.5v-1.4z"/></svg>`,
  superscript: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M4.5 7.5l2.2 3.8 2.2-3.8h2.1l-3.3 5 3.5 5.5H9.1l-2.4-4-2.4 4H2.2l3.5-5.5-3.3-5h2.1zm11-4h2v1l-1.7 1.7c-.2.2-.3.4-.3.7 0 .4.3.6.6.6h1.8v1.2h-2.1c-.8 0-1.3-.5-1.3-1.2 0-.4.2-.7.4-1l1.6-1.6H15.5V3.5z"/></svg>`,
  // הפס התחתון של fontColor, highlight ו-clearFormatting יושב על אותו y
  // (16–17.6): שלושתם באותה שורה בקבוצת "גופן", וקו בסיס שונה נראה כמו
  // אייקון שזז.
  fontColor: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M8.5 3h3l4.5 12h-2.3l-1.1-3.2H7.4L6.3 15H4L8.5 3zm2.3 7L9.9 5.8 9 10h1.8zM3 16h14v1.6H3z"/></svg>`,
  highlight: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M4.7 14.9l.9-2.7 1.8 1.8zM5.6 12.2l1.8 1.8 7.9-7.9-1.8-1.8zM13.5 4.3l1.8 1.8 1.2-1.2-1.8-1.8zM3 16h14v1.6H3z"/></svg>`,
  clearFormatting: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M6 3.5h1.8L11.5 15h-1.9l-.8-2.4H5.2L4.4 15H2.5L6 3.5zm1.7 7.2L6.9 6.2 6 10.7h1.7zm6.8-5.7l3 3-5.5 5.5-3-3 5.5-5.5zm1.5-1.5l1.5 1.5-1 1-1.5-1.5 1-1zM11 16h7v1.6h-7z"/></svg>`,
  growFont: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M6 3.5h1.8L11.5 15h-1.9l-.8-2.4H5.2L4.4 15H2.5L6 3.5zm1.7 7.2L6.9 6.2 6 10.7h1.7zM14.5 3.5l3 3.5h-2v7h-2V7h-2l3-3.5z"/></svg>`,
  shrinkFont: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M6 3.5h1.8L11.5 15h-1.9l-.8-2.4H5.2L4.4 15H2.5L6 3.5zm1.7 7.2L6.9 6.2 6 10.7h1.7zM14.5 14l3-3.5h-2V3.5h-2v7h-2l3 3.5z"/></svg>`,

  // קבוצת פיסקה (Paragraph)
  alignRight: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2 3.5h16v1.6H2V3.5zm5 3.8h11v1.6H7V7.3zm-5 3.8h16v1.6H2v-1.6zm5 3.8h11v1.6H7v-1.6z"/></svg>`,
  alignCenter: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2 3.5h16v1.6H2V3.5zm2.5 3.8h11v1.6h-11V7.3zm-2.5 3.8h16v1.6H2v-1.6zm2.5 3.8h11v1.6h-11v-1.6z"/></svg>`,
  alignLeft: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2 3.5h16v1.6H2V3.5zm0 3.8h11v1.6H2V7.3zm0 3.8h16v1.6H2v-1.6zm0 3.8h11v1.6H2v-1.6z"/></svg>`,
  alignJustify: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2 3.5h16v1.6H2V3.5zm0 3.8h16v1.6H2V7.3zm0 3.8h16v1.6H2v-1.6zm0 3.8h16v1.6H2v-1.6z"/></svg>`,
  bulletList: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2.5 4.2a1.5 1.5 0 1 1 3 0 1.5 1.5 0 0 1-3 0zM2.5 10a1.5 1.5 0 1 1 3 0 1.5 1.5 0 0 1-3 0zM2.5 15.8a1.5 1.5 0 1 1 3 0 1.5 1.5 0 0 1-3 0zM7.5 3.4h10v1.7h-10zM7.5 9.2h10v1.7h-10zM7.5 15h10v1.7h-10z"/></svg>`,
  numberList: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M4.5 3h1.2v4H4.5V4.2H3.8V3.4l.7-.4zm-1 6.5a1.3 1.3 0 0 1 1.4-1.3c.8 0 1.3.5 1.3 1.2 0 .5-.3.9-.8 1.3l-.7.6h1.6v.9H2.8v-.8l1.3-1.1c.3-.3.4-.5.4-.7 0-.3-.2-.5-.5-.5-.3 0-.5.2-.5.5H3.5zm0 4.8c0-.7.6-1.1 1.3-1.1s1.3.4 1.3 1c0 .4-.2.7-.6.9.5.2.7.5.7 1 0 .7-.6 1.1-1.4 1.1-.8 0-1.4-.4-1.4-1.1h1c0 .3.2.4.4.4.3 0 .4-.2.4-.4 0-.3-.2-.4-.5-.4h-.4v-.8h.4c.3 0 .4-.1.4-.4 0-.2-.1-.3-.3-.3-.2 0-.4.1-.4.3h-.9zm4.5-9.6h9v1.6h-9V4.7zm0 4.8h9v1.6h-9V9.5zm0 4.8h9v1.6h-9v-1.6z"/></svg>`,
  indentIncrease: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2 3.5h16v1.6H2V3.5zm6 3.8h10v1.6H8V7.3zm0 3.8h10v1.6H8v-1.6zm-6 3.8h16v1.6H2v-1.6zM2 7l4 3.5L2 14V7z"/></svg>`,
  indentDecrease: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2 3.5h16v1.6H2V3.5zm6 3.8h10v1.6H8V7.3zm0 3.8h10v1.6H8v-1.6zm-6 3.8h16v1.6H2v-1.6zm4-3.5L6 7v7l-4-3.5z"/></svg>`,
  // הקערה של סימן הפיסקה נבנית מחצי-מעגל מדויק (מרחק המיתר = 2r). הגרסה
  // הקודמת השתמשה ברדיוס גדול מדי לגובה שהוקצה לה, וטופס הקערה יצא 1.1
  // יחידות מעל ה-viewBox — כלומר נחתך.
  dirRtl: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M10.5 2.5H4v1.6h1.7v8.4h1.9V4.1h1.5v8.4h2V8.2h.6a2.85 2.85 0 0 0 0-5.7zM6.2 14.95H15v1.6H6.2zM3 15.75l3.6-1.9v3.8z"/></svg>`,
  dirLtr: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M9.5 2.5h6.5v1.6h-1.7v8.4h-1.9V4.1h-1.5v8.4H8.9V8.2h-.6a2.85 2.85 0 0 1 0-5.7zM5 14.95h8.8v1.6H5zM17 15.75l-3.6-1.9v3.8z"/></svg>`,
  pilcrow: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M9.5 2.5h6.5v1.9h-1.7v13.1h-1.9V4.4h-1.5v13.1H8.9V9.5h-.6a3.5 3.5 0 0 1 0-7z"/></svg>`,
  lineSpacing: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M8.5 4.5h9v1.6h-9V4.5zm0 5h9v1.6h-9V9.5zm0 5h9v1.6h-9v-1.6zM4.5 3l-2.5 3h1.8v8H2l2.5 3 2.5-3H5.2V6H7L4.5 3z"/></svg>`,
  borders: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M16.5 2.5h-13a1 1 0 0 0-1 1v13a1 1 0 0 0 1 1h13a1 1 0 0 0 1-1v-13a1 1 0 0 0-1-1zm-7 1.8v5h-5v-5h5zm-5 6.8h5v5h-5v-5zm6.8 5v-5h5v5h-5zm5-6.8h-5v-5h5v5z"/></svg>`,
  shading: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M16 9.44a2.5 2.17 0 0 1-2.5 2.17 2.5 2.17 0 0 1-2.5-2.17c0-1.3 2.5-3.9 2.5-3.9s2.5 2.6 2.5 3.9zM4.5 9l4-3.47 4 3.47h-8zM14 8.14l-6.5-5.64-1.2 1.04 2 1.73-4.5 3.9a1.5 1.3 0 0 0 0 1.82l4.5 3.9a1.5 1.3 0 0 0 2.1 0L14 11l-2-1.73 2-1.13zM2.5 15.94h15v1.56h-15z"/></svg>`,

  // קבוצת עריכה (Editing)
  replace: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2.6 5.75h10.6v1.7H2.6zM12.6 3.6L17.4 6.6 12.6 9.6zM6.8 12.55h10.6v1.7H6.8zM7.4 10.4L2.6 13.4l4.8 3z"/></svg>`,
  select: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M4.8 2.5l10.4 8.5-4.9.4 2.6 5.6-1.9.8-2.6-5.7-3.6 3.4V2.5z"/></svg>`,

  // לשוניות הוספה, פריסה, הפניות
  table: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M16.18 2.5H3.82a1.32 1.5 0 0 0-1.32 1.5v12a1.32 1.5 0 0 0 1.32 1.5h12.35a1.32 1.5 0 0 0 1.32-1.5V4a1.32 1.5 0 0 0-1.32-1.5zm-7.06 1.8v3.5H4V4.3H9.12zm0 5.2v3.5H4V9.5H9.12zm0 5.2v3H4v-3H9.12zm6.88 3H10.71v-3h5.29v3zm0-4.8H10.71V9.5h5.29v3.5zm0-5.2H10.71V4.3h5.29v3.5z"/></svg>`,
  image: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2.5 3.8h15v1.5h-15zM2.5 14.7h15v1.5h-15zM2.5 5.3H4v9.4H2.5zM16 5.3h1.5v9.4H16zM4 14.7l3.2-4.2 2.3 2.9 2.4-3.1 4.1 4.4zM6.6 6.5a1.3 1.3 0 1 1 0 2.6 1.3 1.3 0 0 1 0-2.6z"/></svg>`,
  link: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M8.27 11.73a4.04 4.04 0 0 1 0-5.77l2.31-2.31a4.04 4.04 0 0 1 5.77 5.77l-1.15 1.15-1.5-1.5 1.15-1.15a1.96 1.96 0 0 0-2.77-2.77l-2.31 2.31a1.96 1.96 0 0 0 2.77 2.77l.69-.69 1.5 1.5-.69.69a4.04 4.04 0 0 1-5.77 0zm3.46-3.46a4.04 4.04 0 0 1 0 5.77l-2.31 2.31a4.04 4.04 0 0 1-5.77-5.77l1.15-1.15 1.5 1.5-1.15 1.15a1.96 1.96 0 0 0 2.77 2.77l2.31-2.31a1.96 1.96 0 0 0-2.77-2.77l-.69.69-1.5-1.5.69-.69a4.04 4.04 0 0 1 5.77 0z"/></svg>`,
  pageBreak: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M3.5 2.5h13v3.5h-13V2.5zm0 11.5h13v3.5h-13V14zM2 10.8h4v-1.6H2v1.6zm6 0h4v-1.6H8v1.6zm6 0h4v-1.6h-4v1.6z"/></svg>`,
  toc: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2.5 3.5h15v1.6h-15V3.5zm0 4.2h11v1.6h-11V7.7zm0 4.2h15v1.6h-15v-1.6zm0 4.2h11v1.6h-11v-1.6z"/></svg>`,
  margins: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M3 2.5v15h14v-15H3zm12.2 13.2H4.8V4.3h10.4v11.4zM6.5 6h1.8v8H6.5V6zm5.2 0h1.8v8h-1.8V6z"/></svg>`,
  orientation: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M16.5 3.5h-13a1 1 0 0 0-1 1v11a1 1 0 0 0 1 1h13a1 1 0 0 0 1-1v-11a1 1 0 0 0-1-1zm-13 11.2V5.2h13v9.5h-13z"/></svg>`,
  paperSize: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M12 2.5H4.5a1.5 1.5 0 0 0-1.5 1.5v12A1.5 1.5 0 0 0 4.5 17.5h11a1.5 1.5 0 0 0 1.5-1.5V7l-5-4.5zm3.5 13.2h-11V4.3h6.3v3.8h4.7v7.6z"/></svg>`,
  columns: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2.5 2.5h6.4V4H2.5zM2.5 16h6.4v1.5H2.5zM2.5 4H4v12H2.5zM7.4 4h1.5v12H7.4zM11.1 2.5h6.4V4h-6.4zM11.1 16h6.4v1.5h-6.4zM11.1 4h1.5v12h-1.5zM16 4h1.5v12H16z"/></svg>`,
  footnote: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M16.25 2.5h-12.5a1.25 1.5 0 0 0-1.25 1.5v12a1.25 1.5 0 0 0 1.25 1.5h12.5a1.25 1.5 0 0 0 1.25-1.5V4a1.25 1.5 0 0 0-1.25-1.5zm-12.5 13.2V4.3h12.5v11.4h-12.5zM5.42 6h2.92v1.8H5.42V6zm0 3.5h9.17v1.8h-9.17V9.5zm0 3.5h6.67v1.8h-6.67V13z"/></svg>`,

  // לשונית סקירה. info/proofing/trackChanges היו דיסקים שחורים מלאים בזמן
  // שכל שאר הסט קווי, ולכן "צעקו" מהמסך; כאן הם טבעת וגליפים.
  // trackChanges היה בדיוק אותו path של info — כלומר "עקוב אחר שינויים"
  // הציג אייקון מידע.
  trackChanges: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2.5 3.4h10V5h-10zM2.5 7.3H10v1.6H2.5zM2.5 11.2H7v1.6H2.5zM8 17.3l.9-2.6 1.7 1.7zM8.9 14.7l1.7 1.7 5.7-5.7-1.7-1.7zM14.6 9l1.7 1.7 1.2-1.2-1.7-1.7z"/></svg>`,
  accept: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M7.3 15.2L3 11l1.6-1.6 2.7 2.7 8.4-8.4 1.6 1.6z"/></svg>`,
  reject: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M15.6 2.8L17.2 4.4 11.6 10l5.6 5.6-1.6 1.6L10 11.6l-5.6 5.6-1.6-1.6L8.4 10 2.8 4.4l1.6-1.6L10 8.4z"/></svg>`,
  comment: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M17.5 3.5h-13.64a1.36 1.5 0 0 0-1.36 1.5v9a1.36 1.5 0 0 0 1.36 1.5h10l3.18 3V5a1.36 1.5 0 0 0-1.36-1.5zm0 9.2h-10l-2 2V5.2h12v7.5z"/></svg>`,
  // "ABC" עם וי, כמו בדיקת האיות של Word. הקערות של ה-B הן תת-paths בכיוון
  // הפוך ולכן הן חורים; ה-C הוא טבעת פתוחה משתי קשתות קונצנטריות.
  proofing: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M4 2.5h1.2l-1.4 5.6H2.5zM4 2.5h1.2l1.5 5.6H5.4zM3.9 5.9h2.1V7H3.9zM7.7 2.5h2.4a1.5 1.5 0 0 1 1.1 2.6 1.6 1.6 0 0 1-.7 3H7.7zM9 3.8v1.3h1a.65 .65 0 0 0 0-1.3zM9 6.4v1.4h1.1a.7 .7 0 0 0 0-1.4zM16.6 3.3a2.8 2.8 0 1 0 0 4l-.9-.9a1.5 1.5 0 1 1 0-2.2zM12 16.7L9.5 14.2l.9-.9 1.6 1.5 4.8-4.8.9.9z"/></svg>`,

  // לשונית תצוגה
  ruler: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M2.5 4.5v11h15v-11h-15zm13.2 9.2H4.3V6.3h1.8v2.5h1.8V6.3h1.8v1.8h1.8V6.3h1.8v2.5h1.8V6.3h1.1v7.4z"/></svg>`,
  zoom: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M8.5 3a5.5 5.5 0 0 1 4.3 8.9l4.4 4.4-1.3 1.3-4.4-4.4A5.5 5.5 0 1 1 8.5 3zm0 1.8a3.7 3.7 0 1 0 0 7.4 3.7 3.7 0 0 0 0-7.4zm.9 1.8h-1.8v1.8H5.8v1.8h1.8v1.8h1.8v-1.8h1.8V8.4H9.4V6.6z"/></svg>`,
  fitWidth: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M3.4 5.5h13.2V7H3.4zM3.4 13h13.2v1.5H3.4zM3.4 7h1.5v6H3.4zM15.1 7h1.5v6h-1.5zM2 5.5h1.2v9H2zM16.8 5.5H18v9h-1.2z"/></svg>`,
  focusMode: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M4.5 4.5h3.5V3H3v5h1.5V4.5zm8-1.5v1.5H16V8h1.5V3h-5zm3.5 12.5h-3.5V17H17v-5h-1.5v3.5zM8 15.5H4.5V12H3v5h5v-1.5z"/></svg>`,
  // print ו-info הם פקדים של לשונית "קובץ" שיושבים בקבוצה הזאת היסטורית.
  // `info` הוא טבעת עם גליף ולא דיסק מלא, ולכן הוא אינו "צועק" ליד סט קווי.
  print: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M5 4.5C5 3.67 5.67 3 6.5 3h7c.83 0 1.5.67 1.5 1.5V5h.5A2.5 2.5 0 0 1 18 7.5v5c0 .83-.67 1.5-1.5 1.5H15v1.5c0 .83-.67 1.5-1.5 1.5h-7A1.5 1.5 0 0 1 5 15.5V14H3.5A1.5 1.5 0 0 1 2 12.5v-5A2.5 2.5 0 0 1 4.5 5H5zM6 5h8v-.5a.5.5 0 0 0-.5-.5h-7a.5.5 0 0 0-.5.5zm-1 8v-1.5c0-.83.67-1.5 1.5-1.5h7c.83 0 1.5.67 1.5 1.5V13h1.5a.5.5 0 0 0 .5-.5v-5c0-.83-.67-1.5-1.5-1.5h-11C3.67 6 3 6.67 3 7.5v5c0 .28.22.5.5.5zm1.5-2a.5.5 0 0 0-.5.5v4c0 .28.22.5.5.5h7a.5.5 0 0 0 .5-.5v-4a.5.5 0 0 0-.5-.5z"/></svg>`,
  info: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M10.5 8.91a.5.5 0 0 0-1 .09v4.6a.5.5 0 0 0 1-.1V8.91m.3-2.16a.75.75 0 1 0-1.5 0 .75.75 0 0 0 1.5 0M18 10a8 8 0 1 0-16 0 8 8 0 0 0 16 0M3 10a7 7 0 1 1 14 0 7 7 0 0 1-14 0"/></svg>`,
  book: `<svg viewBox="0 0 20 20" fill="currentColor"><path d="M15.5 2.5H5a1.5 1.5 0 0 0-1.5 1.5v12a1.5 1.5 0 0 0 1.5 1.5h10.5v-15zM5.2 4.3h4v6.5L7.2 9.5 5.2 10.8V4.3z"/></svg>`,
};
