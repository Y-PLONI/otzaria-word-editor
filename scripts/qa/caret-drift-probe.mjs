/**
 * שער סחיפת הסמן: האם הקו נשאר על ההיסט בזמן הקלדה **מתמשכת**.
 *
 * ## התלונה שזה נולד ממנה
 *
 * „הסמן ממוקם מקום אחד לפני המיקום האמיתי — לא כשמתחילים להקליד, רק אחרי
 * 10–15 תווים.”
 *
 * ## למה זה שער נפרד מ-caret-position-probe
 *
 * השער האחר מודד **מצב**: מקליד, עוצר, ומודד פעם אחת. הוא ירוק, והוכח שהוא
 * מסוגל להאדים — ובכל זאת אינו יכול לתפוס את זה, כי „אחרי 10–15 תווים” אינו
 * גבול של תווים אלא של **זמן**: 10–15 תווים בהקלדה אנושית הם 2.5–4 שניות, וזה
 * בדיוק `AUTOSAVE_DELAY_MS = 2500` (sessions/save-coordinator.ts). מדידה אחרי
 * עצירה מפספסת את החלון שבו הבעיה חיה.
 *
 * לכן כאן ההפך: הקלדה **רציפה** בקצב אנושי, ודגימה אחרי **כל תו**. מה שנמדד
 * הוא לא „האם הסמן נכון בסוף” אלא „האם היה רגע שבו הוא לא היה”.
 *
 * ## למה שני סבבים, ולא אחד
 *
 * סבב אחד שמראה סחיפה מוכיח שיש בעיה, לא **מה** גורם לה. שני הסבבים זהים
 * בכל פרט חוץ ממתג „שמירה אוטומטית”, ולכן ההפרש ביניהם הוא בדיוק תרומת
 * השמירה — ואם אין הפרש, ההשערה נופלת ואנחנו יודעים לחפש במקום אחר. זה מה
 * שהופך את השער הזה לניסוי ולא לתצפית.
 *
 * ## הקצב
 *
 * 220ms בין תווים, ולא 45ms כמו בשער האחר. הקלדה מהירה של מכונה מסיימת 16
 * תווים ב-0.7 שניות — לפני שכל דבר מתוזמן בקוד הזה בכלל מתעורר, ולכן היא
 * עיוורת לכל באג שנולד מהשהיה. 220ms הוא קצב הקלדה אנושי סביר, ו-40 תווים
 * בו הם ~9 שניות — כלומר שלושה חלונות אוטוסייב, לא אחד על הגבול.
 *
 *   npm run build && node scripts/qa/caret-drift-probe.mjs
 */
import { openApp, createReport } from './harness.mjs';

/** מה שמוקלד. ארוך מספיק כדי לחצות את חלון האוטוסייב שלוש פעמים. */
const TEXT = 'אבגדהוזחטיכלמנסעפצקרשתאבגדהוזחטיכלמנסעפצ';

/**
 * קצב אנושי. ההנמקה למעלה.
 *
 * ניתן לעקיפה ב-`CARET_GAP_MS`: „אחרי 10–15 תווים” יכול להיות גם גבול של
 * **תור** ולא של שעון — הקלדה מהירה מספיק כדי שהעבודה תצטבר מאחור. שני
 * הקצבים נמדדים באותו שער, כי הם שתי השערות שונות על אותו תסמין.
 */
const GAP_MS = Number(process.env.CARET_GAP_MS ?? 220);

/** סטייה מותרת בין הקו לגבול, בפיקסלים. ראו caret-position-probe.mjs. */
const TOLERANCE_PX = 2;

/**
 * דוגם **בתוך הדף**, ב-16ms, ולא דרך CDP.
 *
 * ## למה זה ההבדל בין לתפוס את הבאג לבין לפספס אותו
 *
 * המדידה הראשונה כאן דגמה דרך CDP אחרי כל תו — כלומר המתינה 220ms **ועוד**
 * הלוך-ושוב של הפרוטוקול, ורק אז שאלה. היא יצאה נקייה לחלוטין (40/40, עם
 * שמירה אוטומטית ובלעדיה). אבל היא אינה יכולה לראות **פיגור**: אם הקו נגרר
 * תו אחד אחורה בזמן ההקלדה ומדביק את הפער כשעוצרים, כל דגימה שנעשית אחרי
 * המתנה תמצא אותו במקום הנכון. „הסמן תו אחד לפני” הוא בדיוק תיאור של פיגור.
 *
 * לכן הדגימה כאן היא גיאומטרית טהורה ורצה בדף: אין קריאה א-סינכרונית למנוע,
 * ולכן אין למה להמתין. זו אותה שיטה, ומאותה סיבה, כמו scripts/readout-probe.mjs.
 *
 * ## מה נמדד בלי לשאול את המנוע
 *
 * כל עוד מקלידים בסוף השורה, המקום הנכון של הסמן ידוע גיאומטרית: הגבול שאחרי
 * התו האחרון. ב-RTL זה הקצה **השמאלי** של הטקסט. לכן הדגימה משווה את ה-x של
 * הקו המצויר ל-x של `Range` מכווץ בסוף הצומת — שני מספרים שנקראים מה-DOM
 * באותו פריים, בלי לצאת מהדף.
 *
 * ה-textarea המוסתר נפסל במפורש: הוא ברוחב 1px ובאותה קואורדינטה כמו הסמן,
 * ואיתור „האלמנט הצר הראשון” מחזיר אותו — הטעות שהשאירה שער קודם עיוור
 * להזזה של 12px בסמן עצמו.
 */
const INSTALL_SAMPLER = `(() => {
  if (window.__drift && window.__drift.timer) clearInterval(window.__drift.timer);
  const D = window.__drift = { samples: [], timer: null };
  D.timer = setInterval(() => {
    try {
      let bar = null;
      const all = document.querySelectorAll('.editor-stack *');
      for (const el of all) {
        const cls = el.className && el.className.baseVal !== undefined ? el.className.baseVal : String(el.className || '');
        if (el.tagName === 'TEXTAREA' || !/caret/i.test(cls)) continue;
        const b = el.getBoundingClientRect();
        if (b.width > 0 && b.width <= 4 && b.height >= 8) { bar = b; break; }
      }
      if (!bar) return;
      const midY = bar.top + bar.height / 2;
      let node = null;
      for (const el of document.querySelectorAll('.superdoc-line, .superdoc-fragment')) {
        const r0 = el.getBoundingClientRect();
        if (midY < r0.top - 2 || midY > r0.bottom + 2) continue;
        const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
        let n;
        while ((n = w.nextNode())) { if (n.data && n.data.trim()) { node = n; break; } }
        if (node) break;
      }
      if (!node) return;
      const r = document.createRange();
      r.setStart(node, node.data.length); r.setEnd(node, node.data.length);
      const rects = Array.from(r.getClientRects());
      if (!rects.length) return;
      const endX = Math.round(rects[0].left * 10) / 10;
      const caretX = Math.round(bar.left * 10) / 10;
      const last = D.samples[D.samples.length - 1];
      // רק שינויים נרשמים: 16ms × 10 שניות הם 600 דגימות זהות ברובן, והרצף
      // המעניין הוא מתי הזוג השתנה — לא כמה פעמים הוא לא.
      if (!last || last.caretX !== caretX || last.endX !== endX || last.len !== node.data.length) {
        D.samples.push({ t: Math.round(performance.now()), caretX, endX, len: node.data.length });
      }
    } catch (e) { /* פריים אחד שנפל אינו סיבה להרוג את הדוגם */ }
  }, 16);
  return true;
})()`;

const READ_SAMPLER = `(() => {
  const D = window.__drift;
  if (!D) return 'null';
  clearInterval(D.timer);
  D.timer = null;
  return JSON.stringify(D.samples);
})()`;

/**
 * דגימה אחת דרך CDP: ההיסט מהמנוע, הגבול הגיאומטרי שלו, ו-x של הקו המצויר.
 * משמשת לאימות מצב סופי, לא לתפיסת פיגור.
 */
const SAMPLE = `(async () => {
  const out = { offset: null, boundary: null, caretX: null, text: null, error: null };
  try {
    const doc = window.__qa.doc();
    if (!doc) { out.error = 'אין doc'; return JSON.stringify(out); }
    const cur = await doc.selection.current();
    const t = cur && cur.selectionTarget;
    if (!t || !t.start) { out.error = 'אין selectionTarget'; return JSON.stringify(out); }
    out.offset = t.start.offset;

    let bar = null;
    document.querySelectorAll('.editor-stack *').forEach((el) => {
      if (bar) return;
      const cls = el.className && el.className.baseVal !== undefined ? el.className.baseVal : String(el.className || '');
      if (el.tagName === 'TEXTAREA' || !/caret/i.test(cls)) return;
      const b = el.getBoundingClientRect();
      if (b.width > 0 && b.width <= 4 && b.height >= 8) bar = b;
    });
    if (!bar) { out.error = 'לא אותר סמן מצויר'; return JSON.stringify(out); }
    out.caretX = Math.round(bar.left * 10) / 10;

    // צומת הטקסט שהסמן יושב בו, לפי חפיפה אנכית.
    const midY = bar.top + bar.height / 2;
    let node = null;
    for (const el of document.querySelectorAll('.superdoc-line, .superdoc-fragment')) {
      const r0 = el.getBoundingClientRect();
      if (midY < r0.top - 2 || midY > r0.bottom + 2) continue;
      const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      let n;
      while ((n = w.nextNode())) { if (n.data && n.data.trim()) { node = n; break; } }
      if (node) break;
    }
    if (!node) { out.error = 'לא נמצאה שורה מצוירת'; return JSON.stringify(out); }
    out.text = node.data;
    if (out.offset > node.data.length) { out.error = 'ההיסט חורג מהצומת'; return JSON.stringify(out); }

    const r = document.createRange();
    r.setStart(node, out.offset); r.setEnd(node, out.offset);
    out.boundary = Array.from(r.getClientRects()).map((b) => Math.round(b.left * 10) / 10);
  } catch (e) { out.error = String(e && e.message); }
  return JSON.stringify(out);
})()`;

const report = createReport('סחיפת הסמן בהקלדה מתמשכת', { strict: true });
const app = await openApp({ name: 'caret-drift', port: Number(process.env.QA_PORT ?? 9373) });

/** האם מתג „שמירה אוטומטית” דלוק. */
const autosaveOn = () =>
  app.js(`(() => { const b = document.querySelector('.autosave-toggle'); return b ? b.getAttribute('aria-checked') === 'true' : null; })()`);

/**
 * מקליד את `TEXT` בקצב אנושי ודוגם אחרי כל תו.
 *
 * הדגימה נעשית **בין** התווים ולא בסופם: זה מה שהופך את המדידה לחלון ולא
 * לצילום, וזה כל ההבדל מול השער האחר.
 */
async function typeAndSample(label) {
  const anchored = await app.caret(0);
  if (!anchored) {
    report.fail(`${label} — עוגן`, 'לא ניתן למקם סמן');
    return null;
  }

  // הדוגם נדלק לפני התו הראשון ונקרא אחרי האחרון. בין לבין אין שום קריאת
  // CDP: כל הלוך-ושוב היה נותן למנוע להתייצב, וזה בדיוק מה שמסתיר פיגור.
  await app.js(INSTALL_SAMPLER);
  for (let i = 0; i < TEXT.length; i++) {
    await app.type(TEXT[i], 0);
    await app.sleep(GAP_MS);
  }
  await app.sleep(600);
  const samples = JSON.parse(await app.js(READ_SAMPLER)) ?? [];

  /* חריגה = הקו אינו בסוף הטקסט המצויר.
   *
   * `len` נלקח לצד המרחק כדי להבדיל בין שני הסברים: פיגור אמיתי (הקו מאחור
   * בזמן ש-`len` כבר גדל) לעומת פריים שבו הטקסט צויר והסמן טרם — שני מצבים
   * שנראים אותו דבר במספר אחד. */
  const bad = samples.filter((s) => Math.abs(s.caretX - s.endX) > TOLERANCE_PX);

  /* כמה זמן הקו היה מחוץ למקום. דגימה שנרשמה ב-t והבאה ב-t' — משך החריגה הוא
   * ההפרש. פיגור של פריים אחד (16–32ms) אינו נראה לעין; פיגור שנמשך מאות
   * מילישניות הוא בדיוק מה שהמשתמש מתאר. */
  let worstMs = 0;
  for (let i = 0; i < samples.length; i++) {
    if (Math.abs(samples[i].caretX - samples[i].endX) <= TOLERANCE_PX) continue;
    const until = samples[i + 1] ? samples[i + 1].t : samples[i].t;
    worstMs = Math.max(worstMs, until - samples[i].t);
  }

  console.log(`\n--- ${label} --- ${samples.length} שינויים, ${bad.length} חורגות, החריגה הארוכה ${worstMs}ms`);
  if (bad.length) console.log('החורגות:', JSON.stringify(bad.slice(0, 15)));

  if (bad.length === 0) {
    report.pass(`${label} — הקו בסוף הטקסט`, `${samples.length} שינויים נדגמו ב-16ms, אף אחד לא חרג`);
  } else {
    report.fail(
      `${label} — הקו בסוף הטקסט`,
      `${bad.length} מתוך ${samples.length} שינויים חורגים, החריגה הארוכה ${worstMs}ms. הראשונה: ${JSON.stringify(bad[0])}`,
    );
  }

  // אימות מצב סופי דרך המנוע — שההיסט עצמו הגיע לאן שצריך.
  const fin = JSON.parse(await app.js(SAMPLE));
  if (fin.error) report.fail(`${label} — מצב סופי`, fin.error);
  else if (fin.offset !== TEXT.length) report.fail(`${label} — מצב סופי`, `היסט ${fin.offset} אחרי ${TEXT.length} תווים`);
  else report.pass(`${label} — מצב סופי`, `היסט ${fin.offset}`);

  return { samples, bad, worstMs };
}

/** מנקה את הפסקה: בחירת הכול במקלדת ומחיקה. */
async function clearLine() {
  for (let i = 0; i < TEXT.length + 5; i++) await app.press('Backspace', 'Backspace', 8, 0);
  await app.sleep(500);
}

try {
  const initial = await autosaveOn();
  console.log('מתג שמירה אוטומטית בהתחלה:', initial);

  /* סבב א׳ — כפי שהמצב הגיע. */
  const withAutosave = await typeAndSample(initial ? 'עם שמירה אוטומטית' : 'בלי שמירה אוטומטית');
  await clearLine();

  /* סבב ב׳ — אותו דבר בדיוק, עם המתג הפוך. ההפרש הוא התשובה. */
  const flipped = await app.clickSel('.autosave-toggle', 0, { after: 700 });
  const now = await autosaveOn();
  if (!flipped || now === initial) {
    report.skip('הסבב השני', `המתג לא התהפך (לפני=${initial}, אחרי=${now}) — אין השוואה`);
  } else {
    const without = await typeAndSample(now ? 'עם שמירה אוטומטית' : 'בלי שמירה אוטומטית');

    /* המסקנה: האם ההפרש הוא השמירה. */
    if (withAutosave && without) {
      const a = withAutosave.bad.length;
      const b = without.bad.length;
      const verdict = `${initial ? 'דלוק' : 'כבוי'}=${a} חורגות, ${now ? 'דלוק' : 'כבוי'}=${b} חורגות`;
      if (a === 0 && b === 0) report.pass('השוואת שני הסבבים', `אין סחיפה בשניהם — ${verdict}. ההשערה על השמירה האוטומטית נופלת.`);
      else if (a > 0 !== b > 0) report.fail('השוואת שני הסבבים', `הסחיפה תלויה במתג — ${verdict}. זהו הגורם.`);
      else report.fail('השוואת שני הסבבים', `סחיפה בשני הסבבים — ${verdict}. המתג אינו הגורם, אבל יש סחיפה.`);
    }
  }
} catch (error) {
  report.fail('השער עצמו', String(error && error.message));
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
