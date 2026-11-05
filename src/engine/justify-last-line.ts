/**
 * שורה אחרונה בפסקה מיושרת לשני הצדדים שבורחת מהשוליים — עקיפה עד שהמנוע יתוקן.
 *
 * הפער, הגורם בקוד המנוע ומה שנמדד: docs/engine-gaps.md, „השורה האחרונה בפסקה
 * מיושרת חורגת מהשוליים”. בקצרה: ביישור לשני הצדדים המודד מכניס לשורה מילה
 * שחורגת מעט, על סמך כיווץ הרווחים בציור — כמו Word במצב 2013 ואילך. הצייר
 * מכווץ רק שורה שהוא מיישר, והשורה האחרונה בפסקה אינה מיושרת; לכן היא נשארת
 * ברוחבה הטבעי, והמילה האחרונה יוצאת מהשוליים (נמדד: 2.5–17.6 פיקסלים).
 *
 * Word מכווץ את השורה הזאת עד שהיא נוגעת בשוליים בדיוק, בלי למתוח אותה
 * (נמדד ב-PDF). זה מה שנעשה כאן, על ה-DOM שהמנוע צייר: שורה שהטקסט שלה יוצא
 * מהקופסה שלה, ושהמנוע לא נתן לה `word-spacing`, מקבלת `word-spacing` שלילי
 * שמחזיר אותה בדיוק לקופסה. שורה שהמנוע כבר מכווץ או מותח נושאת `word-spacing`
 * משלו ואינה נוגעת — ולכן כשהמנוע יתוקן, העקיפה מפסיקה לפעול מעצמה.
 *
 * הסמן והלחיצה עוקבים אחרי הכיווץ: המנוע ממפה לחיצה דרך `elementsFromPoint`
 * (ראו pointer-snap.ts), כלומר לפי מה שמצויר.
 */

/** שורה מצוירת של המנוע. */
export const LINE_SELECTOR = '.superdoc-line';

/**
 * חריגה שמתחתיה לא נוגעים — עיגול תת-פיקסל. המודד עצמו מקבל חצי פיקסל
 * (`WIDTH_FUDGE_PX`) בשורה שאינה מיושרת, וזו אינה חריגה שנראית.
 */
const MIN_SPILL_PX = 0.5;

/**
 * מסמן שה-`word-spacing` של השורה נכתב כאן, ובאיזה ערך. בלעדיו כיווץ שלנו
 * שנשאר על שורה שהמנוע עדכן במקומה נראה כמו כיווץ של המנוע, ולא היה נמדד שוב.
 */
const OURS = 'oweSpillSpacing';

/** הרווחים ש-`word-spacing` חל עליהם, וגם המנוע סופר כנקודות כיווץ. */
const SPACES = /[  ]/g;

/**
 * בכמה פיקסלי CSS הטקסט של השורה יוצא מהקופסה שלה, בצד שבו הוא יוצא; 0 אם
 * אינו יוצא.
 *
 * נמדד על המסך ומומר חזרה ל-CSS: הזום של העורך הוא `transform` על עוטף
 * העמודים, ו-`word-spacing` נכתב ביחידות שלפני ה-transform.
 */
export function lineSpill(line: HTMLElement): number {
  const box = line.getBoundingClientRect();
  if (box.width === 0 || line.offsetWidth === 0) return 0;
  const range = line.ownerDocument.createRange();
  range.selectNodeContents(line);
  const rects = Array.from(range.getClientRects()).filter((r) => r.width > 0);
  if (rects.length === 0) return 0;
  const inkLeft = Math.min(...rects.map((r) => r.left));
  const inkRight = Math.max(...rects.map((r) => r.right));
  const spill = Math.max(box.left - inkLeft, inkRight - box.right, 0);
  return spill * (line.offsetWidth / box.width);
}

/**
 * מכווץ שורה אחת שחורגת, אם צריך. מחזיר אם נכתב משהו.
 *
 * שורה בלי רווחים (מילה אחת ארוכה מהטור) נשארת כפי שהיא — אין מה לכווץ, וגם
 * Word שובר אותה ולא מכווץ.
 */
export function compressSpillingLine(line: HTMLElement): boolean {
  // כיווץ שלנו מסבב קודם — השורה עודכנה במקומה, ומודדים אותה מחדש בלעדיו.
  if (line.dataset[OURS] !== undefined && line.dataset[OURS] === line.style.wordSpacing) {
    line.style.wordSpacing = '';
    delete line.dataset[OURS];
  }
  if (line.style.wordSpacing !== '') return false;
  const spill = lineSpill(line);
  if (spill <= MIN_SPILL_PX) return false;
  const spaces = line.textContent?.match(SPACES)?.length ?? 0;
  if (spaces === 0) return false;
  line.style.wordSpacing = `${-spill / spaces}px`;
  line.dataset[OURS] = line.style.wordSpacing;
  return true;
}

export interface JustifyLastLineHandle {
  dispose(): void;
}

/**
 * מתקין את העקיפה על ה-container של המנוע.
 *
 * ‏`MutationObserver` על הוספת צמתים בלבד: שורה שהמנוע צייר מחדש מגיעה כאן
 * כהוספה — של השורה עצמה, או של ריצה בתוכה. ה-`word-spacing` שאנחנו כותבים הוא שינוי
 * תכונה, ואינו מזמן מעבר נוסף. המדידה נדחית לפריים הבא — היא קוראת פריסה,
 * ובתוך ה-callback המנוע עוד באמצע כתיבה.
 */
export function installJustifyLastLine(root: HTMLElement): JustifyLastLineHandle {
  const view = root.ownerDocument.defaultView ?? window;
  const pending = new Set<HTMLElement>();
  let frame = 0;

  const flush = () => {
    frame = 0;
    for (const line of pending) if (line.isConnected) compressSpillingLine(line);
    pending.clear();
  };

  const collect = (node: Node) => {
    const el = node instanceof view.HTMLElement ? node : node.parentElement;
    if (!el) return;
    const owner = el.closest<HTMLElement>(LINE_SELECTOR);
    if (owner) pending.add(owner);
    else el.querySelectorAll<HTMLElement>(LINE_SELECTOR).forEach((line) => pending.add(line));
  };

  collect(root);
  const observer = new view.MutationObserver((records) => {
    for (const record of records) record.addedNodes.forEach(collect);
    if (pending.size > 0 && frame === 0) frame = view.requestAnimationFrame(flush);
  });
  observer.observe(root, { childList: true, subtree: true });
  if (pending.size > 0) frame = view.requestAnimationFrame(flush);

  return {
    dispose() {
      observer.disconnect();
      if (frame !== 0) view.cancelAnimationFrame(frame);
      pending.clear();
    },
  };
}
