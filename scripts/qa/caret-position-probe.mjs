/**
 * שער מיקום הסמן: האם הקו המהבהב מצויר במקום שבו הסמן באמת נמצא.
 *
 * ## התלונה שזה נולד ממנה
 *
 * „הסמן ממוקם מקום אחד לפני המיקום האמיתי” — מדווח על טקסט עברי, ומאומת
 * ויזואלית בצילום מסך שבו הטקסט מעורב: עברית עם לוכסנים בתוכה.
 *
 * ## מה נמדד, ולמה דווקא כך
 *
 * שלוש כמויות, ולא אחת — כי „הסמן במקום הלא נכון” יכול לנבוע משני מקומות
 * שונים לגמרי, והתיקון שונה בכל אחד:
 *
 *   1. **ההיסט שהמנוע מדווח** (`doc.selection.current()` → `offset`). אחרי
 *      הקלדת N תווים לפסקה ריקה הוא חייב להיות N.
 *   2. **איפה הדפדפן מציב את גבול ההיסט הזה** — `Range` מכווץ על צומת הטקסט
 *      המצויר, ו-`getClientRects()`. זו האמת הגיאומטרית: זה בדיוק המקום שבו
 *      התו הבא ייכנס.
 *   3. **איפה הקו מצויר בפועל** — `sd-v2-local-selection-caret`, שנמצא לפי
 *      צורה (אלמנט צר וגבוה) ולא לפי שם מחלקה שאינו בחוזה.
 *
 * ההפרש בין 2 ל-3 הוא באג ציור: ההיסט נכון, הקו לא שם. ההפרש בין 1 לאורך
 * הטקסט הוא באג מודל: הסמן באמת יושב במקום הלא נכון, והקו רק מספר את האמת.
 * שער שמודד רק את הקו לא היה יודע להבדיל.
 *
 * ## למה דווקא הטקסטים האלה
 *
 * עברית נקייה כבר נמדדה תקינה (סטייה 0.0px). המקרים שנשארו הם **גבולות
 * דו-כיווניים** — לוכסן, ספרה, אות לטינית בתוך עברית — ושם הסמן במקום אחד
 * לוגי יכול להיות מצויר בשני מקומות פיזיים, וזה בדיוק המצב שבו „תו אחד לפני”
 * נולד. הצילום שהתלונה הגיעה איתו הכיל לוכסנים, ולכן הם הראשונים ברשימה.
 *
 * **למה `getClientRects()` ולא `getBoundingClientRect()`:** על גבול דו-כיווני
 * טווח מכווץ מחזיר **שני** מלבנים — אחד לכל כיוון — וה-bounding שלהם הוא
 * האיחוד, כלומר קופסה רחבה שאינה אף אחד משני המיקומים. המדידה קוראת את
 * המלבנים עצמם, ומקבלת את הקרוב שבהם: סמן שמצויר על אחד משני הצדדים
 * החוקיים של הגבול אינו הבאג — הבאג הוא סמן שאינו על אף אחד מהם.
 *
 * מונע דרך CDP ולא ב-jsdom: אין ל-jsdom פריסה, ובלי פריסה אין x לסמן ואין מה
 * למדוד. זו גם הסיבה שזה שער ולא טסט יחידה.
 *
 *   npm run build && node scripts/qa/caret-position-probe.mjs
 */
import { openApp, createReport } from './harness.mjs';

/**
 * המקרים. `back` = כמה פעמים ללחוץ „חץ אחורה” אחרי ההקלדה, כדי למדוד גם סמן
 * שיושב **בתוך** הטקסט ולא רק בסופו: בסוף הפסקה יש רק כיוון אחד אפשרי, ובאמצע
 * יש שניים — ושם הבאג היה מתחבא.
 */
const CASES = [
  { name: 'עברית נקייה', text: 'שלום עולם', back: 0 },
  { name: 'עברית עם לוכסנים', text: 'רעש /סמן/ פכלך', back: 0 },
  { name: 'עברית עם לוכסנים — סמן באמצע', text: 'רעש /סמן/ פכלך', back: 4 },
  { name: 'עברית עם ספרות', text: 'פרק 12 משנה 3', back: 0 },
  { name: 'עברית עם לטינית', text: 'ספר Word גדול', back: 0 },
  // ניקוד: סימני צירוף ברוחב אפס. הצילום שהתלונה הגיעה איתו הכיל אותם, וזה
  // גם המקרה שמנוע ה-DOCX כבר נמדד כשוגה בו במקום אחר (engine/word-selection.ts).
  { name: 'עברית מנוקדת', text: 'בְּרֵאשִׁית בָּרָא', back: 0 },
];

/**
 * מסלולי החצים. כל מסלול נמדד **צעד-צעד** ולא בסך הכול: „הסמן לא הגיע ליעד”
 * אינו אומר איפה הוא נתקע, ובלי זה אי אפשר לדעת אם הבעיה היא גבול אחד או כל
 * המסלול. הביקורת (`עברית נקייה`) היא מה שהופך „החצים שבורים” ל„החצים שבורים
 * דווקא כאן”.
 */
const ARROW_WALKS = [
  { name: 'חצים — עברית נקייה (ביקורת)', text: 'שלום עולם רחב' },
  { name: 'חצים — עברית עם ספרות', text: 'פרק 12 משנה 3' },
  { name: 'חצים — עברית עם לטינית', text: 'ספר Word גדול' },
  { name: 'חצים — עברית מנוקדת', text: 'בְּרֵאשִׁית בָּרָא' },
];

/**
 * סטייה מותרת בין הגבול הגיאומטרי לקו המצויר, בפיקסלים.
 *
 * 2px, ולא אפס: הקו עצמו רחב 1–2px והמנוע ממקם אותו במרכזו, ועיגול לתת-פיקסל
 * בזום 100% מוסיף עוד חצי. רוחב תו עברי בגופן ברירת המחדל נמדד 4–11px, ולכן
 * הסף מפריד היטב בין „אותו מקום” לבין „תו אחד משם”.
 */
const TOLERANCE_PX = 2;

/** הטקסט שנלחצים עליו, והתווים שנבדקים בו — תחילה, אמצע וסוף. */
const CLICK_TEXT = 'אבגדהוזחטיכלמנסע';
const CLICK_AT = [2, 5, 8, 11, 14];

/**
 * קורא את שלוש הכמויות בבת אחת.
 *
 * חייב להיות מדידה **אחת** ולא שלוש קריאות נפרדות: כל הלוך-ושוב של CDP הוא
 * הזדמנות לפריסה מחדש, ושלוש כמויות שנמדדו בשלושה רגעים אינן ניתנות להשוואה.
 */
const measure = () => `(async () => {
  const out = { text: null, offset: null, blockLen: null, boundary: null, caret: null, chars: null, error: null };
  try {
    // ההיסט מגיע מהמנוע ולא מ-window.getSelection(): המנוע מחזיק textarea
    // מוסתר ומצייר סמן משלו, ולבחירה של הדפדפן אין כאן צומת טקסט כלל.
    const doc = window.__qa.doc();
    if (!doc) { out.error = 'אין doc'; return JSON.stringify(out); }
    const cur = await doc.selection.current({ includeText: true });
    const t = cur && cur.selectionTarget;
    if (!t || !t.start) { out.error = 'אין selectionTarget: ' + JSON.stringify(cur); return JSON.stringify(out); }
    out.offset = t.start.offset;
    out.engineTarget = { blockId: t.start.blockId, start: t.start.offset, end: t.end && t.end.offset };

    // אורך הבלוק, מהמנוע: \`ranges.resolve\` חותך היסט שחורג ומחזיר את האורך
    // האמיתי. זו הדרך הציבורית היחידה לשאול „איפה הבלוק נגמר” — ראו
    // engine/word-selection.ts, BLOCK_LENGTH_PROBE. בלי זה אי אפשר לדעת אם
    // היסט גדול הוא באג או פשוט בלוק ארוך.
    try {
      const probe = await doc.ranges.resolve({
        start: { kind: 'point', point: { kind: 'text', blockId: t.start.blockId, offset: 0 } },
        end: { kind: 'point', point: { kind: 'text', blockId: t.start.blockId, offset: 1000000 } },
      });
      out.blockLen = probe && probe.target && probe.target.end ? probe.target.end.offset : null;
      out.blockText = probe && probe.preview ? probe.preview.text : null;
    } catch (e) { out.blockLen = 'resolve נכשל: ' + String(e && e.message); }

    // הקו המצויר.
    //
    // **חייב להיות ה-DIV של שכבת הסמן, ולא „האלמנט הצר הראשון”.** איתור לפי
    // צורה בלבד תופס גם את ה-textarea המוסתר של המנוע, והוא הראשון ב-DOM.
    // שניהם ברוחב 1px ובאותה קואורדינטה בדיוק, ולכן מדידה שלוקחת את הראשון
    // נראית נכונה — ומודדת את הדבר הלא נכון. נמדד: הזזת ה-DIV ב-12px (≈1.7
    // תווים) השאירה את השער על „סטייה 0.0px” בכל 36 השורות, כלומר עיוורון
    // מוחלט לבאג שהשער הזה נכתב כדי לתפוס.
    //
    // לכן ה-textarea נאסף בנפרד (shadow) ואינו מועמד: הוא מדווח לצורך
    // אבחון, ולעולם אינו מה שנמדד.
    const cands = [];
    const shadow = [];
    document.querySelectorAll('.editor-stack *').forEach((el) => {
      const b = el.getBoundingClientRect();
      if (!(b.width > 0 && b.width <= 4 && b.height >= 8 && b.height <= 60)) return;
      const cls = el.className && el.className.baseVal !== undefined ? el.className.baseVal : String(el.className || '');
      const row = { cls, tag: el.tagName, left: Math.round(b.left * 10) / 10, top: Math.round(b.top * 10) / 10,
                    w: Math.round(b.width * 10) / 10, h: Math.round(b.height * 10) / 10 };
      if (el.tagName === 'TEXTAREA' || !/caret/i.test(cls)) shadow.push(row);
      else cands.push(row);
    });
    out.caret = cands;
    out.shadow = shadow;
    if (cands.length === 0) { out.error = 'לא אותר אלמנט סמן מצויר (מחלקה שמכילה caret). מועמדים שנפסלו: ' + JSON.stringify(shadow); return JSON.stringify(out); }

    // צומת הטקסט המצויר — נבחר לפי **חפיפה אנכית עם הסמן**, ולא לפי תוכן:
    // שתי פסקאות יכולות להכיל את אותו טקסט, וחיפוש לפי תחילית היה מודד פסקה
    // אחת מול סמן שיושב באחרת.
    const bar = cands[0];
    const midY = bar ? bar.top + bar.h / 2 : null;
    let node = null;
    const lines = document.querySelectorAll('.superdoc-line, .superdoc-fragment');
    for (const el of lines) {
      const r0 = el.getBoundingClientRect();
      if (midY === null || midY < r0.top - 2 || midY > r0.bottom + 2) continue;
      const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      let n;
      while ((n = w.nextNode())) { if (n.data && n.data.trim()) { node = n; break; } }
      if (node) break;
    }
    if (!node) { out.error = 'לא נמצאה שורה מצוירת שהסמן בתוכה'; return JSON.stringify(out); }
    out.text = node.data;

    // הגבול הגיאומטרי של ההיסט. שני מלבנים על גבול דו-כיווני — ראו ההערה.
    const at = Math.max(0, Math.min(out.offset, node.data.length));
    out.clamped = at !== out.offset;
    const r = document.createRange();
    r.setStart(node, at);
    r.setEnd(node, at);
    out.boundary = Array.from(r.getClientRects()).map((b) => Math.round(b.left * 10) / 10);

    // רוחב כל תו, כדי לתרגם „כמה פיקסלים” ל„כמה תווים”.
    out.chars = [];
    for (let i = 0; i < node.data.length; i++) {
      const c = document.createRange();
      c.setStart(node, i); c.setEnd(node, i + 1);
      const b = c.getBoundingClientRect();
      out.chars.push({ ch: node.data[i], left: Math.round(b.left * 10) / 10, w: Math.round(b.width * 10) / 10 });
    }
  } catch (e) { out.error = String(e && e.message); }
  return JSON.stringify(out);
})()`;

const report = createReport('מיקום הסמן', { strict: true });
const app = await openApp({ name: 'caret-position', port: Number(process.env.QA_PORT ?? 9371) });

/** מלבן הגבול הקרוב ביותר לקו — ראו ההנמקה על שני המלבנים. */
function nearest(bars, boundary) {
  return bars
    .map((b) => ({ bar: b, gap: Math.min(...boundary.map((x) => Math.abs(b.left - x))) }))
    .sort((a, b) => a.gap - b.gap)[0];
}

/** כל השער: מקרי ההקלדה, ואחריהם סבבי הקליק ב-100% ובזום. */
async function run() {
  await app.caret(0);

  for (const [i, c] of CASES.entries()) {
    // כל מקרה בפסקה משלו, ו-Enter נלחץ רק כשהסמן בסוף. זה חובה ולא נימוס:
    // מקרה עם `back` משאיר את הסמן באמצע, ו-Enter שם **מפצל** את הפסקה —
    // השארית נדבקת לתחילת המקרה הבא ומזייפת לו את אורך הבלוק (נמדד: בלוק בן
    // 17 תווים על טקסט בן 13, ואחר כך 35). ההחזרה היא באותם חצים שהוציאו
    // אותו משם, ולא ב-`End`: נמדד ש-`End` אינו מחזיר את הסמן לסוף כאן.
    if (i > 0) await app.press('Enter', 'Enter', 13, 0, '\r');
    await app.type(c.text);
    for (let k = 0; k < c.back; k++) await app.press('ArrowRight', 'ArrowRight', 39, 0);
    await app.sleep(500);

    const m = JSON.parse(await app.js(measure()));
    console.log(`\n--- ${c.name} ---\n${JSON.stringify(m)}`);

    if (m.error) {
      report.fail(c.name, m.error);
      continue;
    }

    const expected = c.text.length - c.back;
    if (m.offset !== expected) {
      report.fail(
        c.name + ' — ההיסט',
        `offset=${m.offset}, מצופה ${expected} (${c.text.length} תווים פחות ${c.back} חצים); אורך הבלוק לפי המנוע ${m.blockLen}, טקסט מצויר ${JSON.stringify(m.text)}`,
      );
    } else {
      report.pass(c.name + ' — ההיסט', String(m.offset));
    }

    const bars = m.caret;
    if (bars.length === 0 || m.boundary.length === 0) {
      report.skip(c.name + ' — הקו', `bars=${bars.length}, boundary=${m.boundary.length}`);
      continue;
    }
    // קיצוץ ההיסט לאורך הצומת המצויר פירושו שההיסט והטקסט אינם מאותו בלוק,
    // וכל השוואה אחריו חסרת משמעות. נכשל במקום להשוות מספרים לא קשורים.
    if (m.clamped) {
      report.fail(c.name + ' — הקו', `ההיסט ${m.offset} חורג מהצומת המצויר (${m.text.length} תווים) — המדידה אינה על אותו בלוק`);
      continue;
    }
    const near = nearest(bars, m.boundary);
    const avgChar = m.chars.length ? m.chars.reduce((s, ch) => s + ch.w, 0) / m.chars.length : 0;
    if (m.boundary.length > 1) {
      // גבול דו-כיווני: הטווח המכווץ החזיר שני מלבנים, ו-`nearest` בוחר את
      // הקרוב. זו בדיוק ההגדרה שמוציאה את תלונת „תו אחד לפני” מתחום השער —
      // סמן שצויר בצד הלא נכון מקבל „0.0px”. לכן זה מדווח ואינו נספר כעובר.
      report.partial(
        c.name + ' — הקו',
        `גבול דו-כיווני: שני מלבנים ${JSON.stringify(m.boundary)}, הקו ב-${near.bar.left}. השער אינו יודע לומר איזה מהם הנכון.`,
      );
    } else if (near.gap <= TOLERANCE_PX) {
      report.pass(c.name + ' — הקו', `סטייה ${near.gap.toFixed(1)}px`);
    } else {
      report.fail(
        c.name + ' — הקו',
        `הקו ב-${near.bar.left}, הגבול ב-${JSON.stringify(m.boundary)} — סטייה ${near.gap.toFixed(1)}px ≈ ${(near.gap / (avgChar || 1)).toFixed(2)} תווים`,
      );
    }

    // חזרה לסוף הפסקה, כדי שה-Enter של המקרה הבא לא יפצל אותה. ראו למעלה.
    for (let k = 0; k < c.back; k++) await app.press('ArrowLeft', 'ArrowLeft', 37, 0);
  }
  /* ---------------------------------------------------------------- *
   * מיקום בקליק — הפעולה שהתלונה נמדדה עליה                            *
   * ---------------------------------------------------------------- *
   * הקלדה מוכיחה שהסמן מצויר על הגבול שהמנוע מדווח. היא **אינה** מוכיחה
   * שהגבול הזה הוא זה שנלחץ עליו: שם, ולא בציור, יושבת התלונה „לחצתי, והוא
   * קפץ תו אחד לפני”. לכן כאן נמדד הפער בין ה-x שנלחץ ל-x שהסמן יצא בו.
   *
   * הסף הוא **חצי רוחב תו** ולא אפס: קליק נופל תמיד על הגבול הקרוב, ולכן
   * קליק במרכז תו רחוק בהגדרה חצי תו מכל אחד משני גבולותיו. פער גדול מזה
   * פירושו שהמנוע בחר את הגבול הרחוק — כלומר תו שלם משם.  */
  await app.press('End', 'End', 35, 0);
  await app.press('Enter', 'Enter', 13, 0, '\r');
  await app.type(CLICK_TEXT);
  await app.sleep(600);

  /* ---------------------------------------------------------------- *
   * מסלולי החצים                                                       *
   * ---------------------------------------------------------------- */
  for (const walk of ARROW_WALKS) {
    await app.press('Enter', 'Enter', 13, 0, '\r');
    await app.type(walk.text);
    await app.sleep(500);

    // נקודת המוצא נמדדת ואינה מונחת: אם ה-Enter לא יצר פסקה נקייה, ההנחה
    // „ההיסט ההתחלתי הוא אורך הטקסט” הופכת כל צעד אחריה לרעש. נמדד: הנחה כזו
    // ייצרה מסלול [14…26] על טקסט בן 13 תווים, כלומר מדידה של בלוק אחר.
    const first = JSON.parse(await app.js(measure()));
    const start = first.offset;
    const steps = Math.min(walk.text.length, typeof first.blockLen === 'number' ? first.blockLen : walk.text.length);

    const trail = [start];
    for (let k = 0; k < steps; k++) {
      await app.press('ArrowRight', 'ArrowRight', 39, 0);
      await app.sleep(120);
      const m = JSON.parse(await app.js(measure()));
      trail.push(m.offset);
      if (m.offset === 0) break;
    }
    console.log(`\n--- ${walk.name} --- בלוק ${first.blockLen}, טקסט ${JSON.stringify(first.blockText)}, מסלול ${JSON.stringify(trail)}`);

    // כל לחיצה חייבת להזיז. שוויון בין שני צעדים עוקבים = הסמן נתקע — וזה
    // נכון בשני הכיוונים, כי כיוון החץ תלוי בכיווניות הפסקה שנמדדת ולא מונחת.
    const stuckAt = trail.findIndex((v, idx) => idx > 0 && v === trail[idx - 1]);
    if (stuckAt >= 0) {
      report.fail(walk.name, `נתקע בצעד ${stuckAt} בהיסט ${trail[stuckAt]}. בלוק בן ${first.blockLen}, מסלול ${JSON.stringify(trail)}`);
    } else {
      report.pass(walk.name, `${trail.length - 1} צעדים, בלי היתקעות: ${JSON.stringify(trail)}`);
    }

    // חזרה לנקודת המוצא, כדי שה-Enter הבא לא יפצל את הפסקה.
    for (let k = 1; k < trail.length; k++) await app.press('ArrowLeft', 'ArrowLeft', 37, 0);
    await app.sleep(300);
  }

  await clickPhase('100%');

  /* ואותו דבר בזום שאינו 100%.
   *
   * למה זה שלב נפרד ולא עוד מקרה: המנוע מצייר את הדף דרך `scale()` על מיכל
   * העימוד (ראו styles/shell.css), ושכבת הסמן היא שכבה **אחרת**. שתי שכבות
   * שמתמרחות בשני חישובים שונים נפרדות זו מזו ככל שמתרחקים מפינת העיגון —
   * כלומר בדיוק ההתנהגות של „הסמן תו אחד לפני”, שגדלה לאורך השורה. ב-100%
   * ה-scale הוא 1 והבדיקה שלמעלה עיוורת לזה לגמרי. */
  const zoomed = await stepZoomUp(4);
  if (zoomed === null) {
    report.skip('קליק בזום', 'לא נמצא כפתור הגדלת תצוגה');
  } else {
    await app.sleep(900);
    await clickPhase(`זום ${zoomed}`);
  }
}

/** מעלה את הזום בכמה צעדים ומחזירה את מה ששורת המצב מציגה, או null. */
async function stepZoomUp(steps) {
  for (let i = 0; i < steps; i++) {
    const ok = await app.clickSel('.zoom-step-btn', 1, { after: 350 });
    if (!ok) return null;
  }
  return app.js(`(() => { const b = document.querySelector('.zoom-pct-btn'); return b ? b.textContent.trim() : null; })()`);
}

/**
 * ממקמת סמן בשורה שמכילה `needle`, ומגלגלת אותה למסך קודם.
 *
 * הגלילה חובה אחרי הגדלת זום: שורה שיצאה מהמסך מחזירה מלבן שלילי, וקליק שם
 * הוא קליק בשום מקום — נמדד כ-`selectionTarget: null`.
 */
async function anchorOn(needle) {
  const rect = JSON.parse(await app.js(`(() => {
    const lines = document.querySelectorAll('.superdoc-line, .superdoc-fragment');
    for (const el of lines) {
      if ((el.textContent || '').indexOf(${JSON.stringify(needle)}) < 0) continue;
      el.scrollIntoView({ block: 'center' });
      const r = el.getBoundingClientRect();
      return JSON.stringify({ x: Math.round(r.x + r.width - 8), y: Math.round(r.y + r.height / 2) });
    }
    return 'null';
  })()`));
  if (!rect) return false;
  await app.clickAt(rect.x, rect.y);
  await app.sleep(600);
  return true;
}

/** סבב הקליקים: לוחץ במרכז תווים נבחרים, ומודד איפה הסמן יצא ולאן נכנסה אות. */
async function clickPhase(label) {
  // עוגן: לחיצה על השורה שנבדקת, לפי תוכנה. בלי זה המדידה הראשונה נעשית על
  // השורה שהסמן במקרה נשאר בה — ואחרי שינוי זום הוא כבר לא שם.
  const anchored = await anchorOn(CLICK_TEXT.slice(0, 4));
  if (!anchored) {
    report.fail(`קליק ${label} — עוגן`, `לא נמצאה שורה שמכילה ${JSON.stringify(CLICK_TEXT.slice(0, 4))}`);
    return;
  }
  const laid = JSON.parse(await app.js(measure()));
  if (laid.error || !laid.chars) {
    report.fail(`קליק ${label} — פריסת הטקסט`, laid.error ?? 'אין chars');
    return;
  }
  {
    const y = laid.caret[0]?.top;

    /* הטקסט חייב להיות RTL טהור, אחרת החישוב שמתחת שגוי בשקט. הבדיקה היא
     * שהתווים יורדים ב-x — כלומר תו k יושב מימין לתו k+1. */
    const rtl = laid.chars.length > 1 && laid.chars.every((c, i) => i === 0 || c.left < laid.chars[i - 1].left);
    if (!rtl) {
      report.fail(`קליק ${label} — כיווניות`, `הטקסט אינו RTL טהור, החישוב אינו תקף: ${JSON.stringify(laid.chars.map((c) => c.left))}`);
      return;
    }

    for (const k of CLICK_AT) {
      const ch = laid.chars[k];
      if (!ch || ch.w === 0 || y === undefined) {
        report.skip(`קליק ${label} על תו ${k}`, 'אין גיאומטריה לתו');
        continue;
      }

      /* **לא מרכז התו.**
       *
       * תו k משתרע על `[left, left+w]`, ושני גבולותיו רחוקים ממרכזו בדיוק
       * `w/2`. לכן קליק במרכז עם סף `w/2 + 2` מקבל את **שני** הגבולות, וסטייה
       * של תו אחד אינה ניתנת לזיהוי בהגדרה. נמדד בריצה „נקייה”: שלושה מתוך
       * עשרה קליקים נחתו על הגבול השני ונרשמו כעוברים.
       *
       * לכן הקליק ב-80% לתוך התו מצד הגבול הצפוי. ב-RTL הגבול „לפני תו k”
       * (היסט k) הוא הקצה **הימני** שלו, `left + w`. קליק שם חד-משמעי: הגבול
       * הנכון במרחק 0.2w, השגוי במרחק 0.8w, והסף מפריד ביניהם. */
      const expectedX = ch.left + ch.w;
      const x = ch.left + ch.w * 0.8;
      await app.clickAt(Math.round(x), Math.round(y + 8));
      await app.sleep(450);

      const m = JSON.parse(await app.js(measure()));
      const bar = m.caret?.[0];
      if (m.error || !bar) {
        report.fail(`קליק ${label} על תו ${k} (${ch.ch})`, m.error ?? 'לא אותר סמן');
        continue;
      }
      const gap = Math.abs(bar.left - expectedX);
      const limit = ch.w * 0.5;
      const name = `קליק ${label} על תו ${k} (${ch.ch})`;
      const detail = `נלחץ ב-${x.toFixed(1)}, הגבול הצפוי ${expectedX.toFixed(1)} (היסט ${k}), הסמן ב-${bar.left} — פער ${gap.toFixed(1)}px, offset=${m.offset}`;
      if (m.offset !== k) {
        report.fail(name, `${detail}; המנוע החזיר היסט ${m.offset} במקום ${k} — סטייה של ${m.offset - k} תווים`);
      } else if (gap > limit) {
        report.fail(name, `${detail}; הסף ${limit.toFixed(1)}px — ההיסט נכון אבל הקו מצויר במקום אחר`);
      } else {
        report.pass(name, detail);
      }

      /* המבחן המכריע: התו הבא נכנס למקום שבו הקו עמד?
       *
       * שני השערים שלמעלה מודדים **הסכמה** — שהמנוע, הדפדפן והקו מספרים אותו
       * סיפור. הם לא היו תופסים מצב שבו כולם מסכימים ופשוט טועים. מה שהמשתמש
       * חווה הוא זה: הוא רואה קו במקום אחד, מקליד, והאות נכנסת במקום אחר.
       * לכן כאן נמדדת התוצאה ולא הכוונה — התו מוקלד, והטקסט שיצא נקרא.
       *
       * ההשוואה היא ל-`k` — הגבול שנלחץ עליו — ולא ל-`m.offset`. השוואה
       * ל-`m.offset` היא עקביות פנימית של המנוע מול עצמו: מנוע שמסכים עם
       * עצמו וטועה היה עובר אותה. */
      await app.type('ת');
      await app.sleep(400);
      const after = JSON.parse(await app.js(measure()));
      const got = typeof after.blockText === 'string' ? after.blockText.indexOf('ת') : -1;
      if (got === k) {
        report.pass(`הקלדה ${label} אחרי קליק על תו ${k}`, `נכנס בהיסט ${got}, בדיוק בגבול שנלחץ`);
      } else {
        report.fail(
          `הקלדה ${label} אחרי קליק על תו ${k}`,
          `נלחץ על הגבול ${k}, האות נכנסה בהיסט ${got} — הפרש ${got - k}. הטקסט שיצא: ${JSON.stringify(after.blockText)}`,
        );
      }
      // מחיקת התו שהוקלד, כדי שהמקרה הבא יימדד על אותו טקסט בדיוק.
      // Backspace ולא Ctrl+Z: נמדד ש-undo כאן מיזג פסקאות והשאיר את המסמך
      // במצב אחר לגמרי („אבגדהוזחטיכלמנסעספר Word גדולשנה 3”), וכל מדידה
      // אחריו הייתה על טקסט שלא נבדק.
      await app.press('Backspace', 'Backspace', 8, 0);
      await app.sleep(400);
    }
  }
}

try {
  await run();
} catch (error) {
  report.fail('השער עצמו', String(error && error.message));
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
