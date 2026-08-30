/**
 * שער הסמן ב-WebKit: משחזר את „הסמן תו אחד לפני, ורק אחרי רווח”.
 *
 * ## למה שער נפרד, ולמה WebKit
 *
 * שני השערים האחרים רצים ב-Chrome, ושניהם ירוקים — כולל אחרי שהוכח שהם
 * מסוגלים להאדים. הם ירוקים בצדק: **הבאג אינו קיים ב-Chrome.** אומת על ידי
 * המשתמש, ואומת כאן.
 *
 * התוסף אצל המשתמש אינו רץ ב-Chrome. הוא רץ בתוך אוצריא, וב-macOS זה
 * `WKWebView` — כלומר WebKit. לכן כל שער שרץ ב-Chrome בלבד אינו יכול לראות
 * את זה, ולא משנה כמה יסודי הוא.
 *
 * ## מה שכבר נמדד, ומה שעדיין לא
 *
 * בדף מינימלי (בלי המנוע, בלי האפליקציה) נמדד הפרש חד בין שני המנועים:
 * `Range` מכווץ בסוף צומת טקסט שנגמר ברווח, ב-`white-space: pre-wrap`,
 * מחזיר ב-Chromium מלבן אחד תקין — וב-WebKit **רשימה ריקה**. אותו טווח בדיוק
 * בלי הרווח הסופי מחזיר מלבן תקין בשני המנועים.
 *
 * מה שזה עדיין **לא** מוכיח: שזה מה שקורה באפליקציה. המנוע מצייר את הסמן
 * כ-DIV שהוא ממקם בעצמו, ואיננו יודעים מאיזה API הוא לוקח את המספר. השער
 * הזה נועד לסגור את הפער: הוא מקליד באפליקציה האמיתית, ב-WebKit, ומודד איפה
 * ה-DIV באמת יושב — לפני הרווח ואחריו.
 *
 * ## מה נמדד
 *
 * שני מצבים על אותה פסקה, ובכוונה בסדר הזה:
 *   1. אחרי מילה — כאן הסמן אמור להיות תקין, וזו הביקורת.
 *   2. אחרי רווח — כאן התלונה.
 * אם (1) תקין ו-(2) שבור, הרווח הוא הגורם, ולא „הסמן שבור ב-WebKit”.
 *
 *   npm run build && node scripts/qa/caret-webkit-probe.mjs
 */
import { readFileSync, writeFileSync, existsSync, rmSync } from 'node:fs';
import { join, dirname, extname, normalize } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { createServer } from 'node:http';
import { createReport } from './harness.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = join(HERE, '..', '..');
const DIST = join(ROOT, 'dist');

/**
 * Playwright אינו תלות של הפרויקט, והוא לא ייהפך לכזו בשביל שער אחד: הוא
 * גורר בניות דפדפן במאות מגה-בייט לכל מי שיריץ `npm install`. לכן הוא נטען
 * מהמטמון של `npx` אם הוא שם, והשער מדלג בהודעה מפורשת אם לא — במקום ליפול
 * ולהיראות כמו רגרסיה.
 */
function loadPlaywright() {
  const require = createRequire(import.meta.url);
  const paths = [
    ...(process.env.PLAYWRIGHT_PATH ? [process.env.PLAYWRIGHT_PATH] : []),
    'playwright',
  ];
  for (const p of paths) {
    try {
      return require(p);
    } catch {
      /* הבא בתור */
    }
  }
  return null;
}

const HOST_STUB = readFileSync(join(HERE, 'host-stub.js'), 'utf8');
const QA_API = readFileSync(join(HERE, 'qa-api.js'), 'utf8');

/** אותה הזרקה של harness.openApp — אחרי ה-latch, לפני הבאנדל. */
function writeProbe() {
  const index = join(DIST, 'index.html');
  if (!existsSync(index)) throw new Error('dist/index.html אינו קיים — הריצו npm run build תחילה');
  const html = readFileSync(index, 'utf8');
  const cut = html.indexOf('</script>') + '</script>'.length;
  if (cut <= '</script>'.length) throw new Error('לא נמצא ה-latch ב-dist/index.html');
  const probe = join(DIST, '__qa-webkit-caret.html');
  writeFileSync(probe, html.slice(0, cut) + `\n<script>${HOST_STUB}</script>\n<script>${QA_API}</script>\n` + html.slice(cut));
  return probe;
}

/**
 * שרת סטטי מעל `dist`.
 *
 * **חובה, ולא נוחות.** ה-QA האחר מגיש את הדף מ-`file://`, וב-Chrome זה עובד.
 * ב-WebKit זה נכשל: הוא מתייחס לכל `file://` כאל origin ייחודי וחוסם ממנו
 * Web Workers, ומנוע ה-DOCX אינו קם בלעדיהם — נמדד, `__qa.ready()` לא נעשה
 * אמת ב-90 שניות. HTTP נותן origin אמיתי, וזה גם מה שקורה אצל המשתמש.
 */
function serveDist() {
  const TYPES = {
    '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.mjs': 'text/javascript',
    '.css': 'text/css', '.json': 'application/json', '.wasm': 'application/wasm',
    '.woff': 'font/woff', '.woff2': 'font/woff2', '.ttf': 'font/ttf',
    '.png': 'image/png', '.jpg': 'image/jpeg', '.svg': 'image/svg+xml', '.map': 'application/json',
  };
  const server = createServer((req, res) => {
    const rel = decodeURIComponent((req.url ?? '/').split('?')[0]);
    // חסימת יציאה מהתיקייה: השרת הזה חי רק בשער, אבל נתיב שמטפס החוצה הוא
    // באג גם כשאין תוקף.
    const file = normalize(join(DIST, rel));
    if (!file.startsWith(DIST) || !existsSync(file)) {
      res.writeHead(404).end('לא נמצא');
      return;
    }
    res.writeHead(200, { 'Content-Type': TYPES[extname(file)] ?? 'application/octet-stream' });
    res.end(readFileSync(file));
  });
  return new Promise((resolve) => {
    server.listen(0, '127.0.0.1', () => resolve({ server, port: server.address().port }));
  });
}

/**
 * המדידה, בתוך הדף.
 *
 * מחזירה את ה-x של ה-DIV של הסמן ואת שני הגבולות הגיאומטריים שסביב התו
 * האחרון. ה-textarea המוסתר נפסל: הוא ברוחב 1px ובאותה קואורדינטה, ואיתור
 * „האלמנט הצר הראשון” מחזיר אותו במקום את הסמן.
 */
const MEASURE = `(() => {
  const out = { caretX: null, text: null, endBoundary: null, prevBoundary: null, error: null };
  try {
    let bar = null;
    for (const el of document.querySelectorAll('.editor-stack *')) {
      const cls = el.className && el.className.baseVal !== undefined ? el.className.baseVal : String(el.className || '');
      if (el.tagName === 'TEXTAREA' || !/caret/i.test(cls)) continue;
      const b = el.getBoundingClientRect();
      if (b.width > 0 && b.width <= 4 && b.height >= 8) { bar = b; break; }
    }
    if (!bar) { out.error = 'לא אותר סמן מצויר'; return out; }
    out.caretX = Math.round(bar.left * 10) / 10;

    const midY = bar.top + bar.height / 2;
    let node = null;
    for (const el of document.querySelectorAll('.superdoc-line, .superdoc-fragment')) {
      const r0 = el.getBoundingClientRect();
      if (midY < r0.top - 3 || midY > r0.bottom + 3) continue;
      const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      let n;
      while ((n = w.nextNode())) { if (n.data && n.data.length) { node = n; break; } }
      if (node) break;
    }
    if (!node) { out.error = 'לא נמצאה שורה מצוירת'; return out; }
    out.text = node.data;
    out.ws = getComputedStyle(node.parentElement).whiteSpace;

    const rectsAt = (i) => {
      const r = document.createRange(); r.setStart(node, i); r.setEnd(node, i);
      return Array.from(r.getClientRects()).map((b) => Math.round(b.left * 10) / 10);
    };
    out.endBoundary = rectsAt(node.data.length);
    out.prevBoundary = node.data.length > 0 ? rectsAt(node.data.length - 1) : [];
  } catch (e) { out.error = String(e && e.message); }
  return out;
})()`;

const report = createReport('הסמן ב-WebKit', { strict: true });
const pw = loadPlaywright();

if (!pw) {
  report.skip(
    'WebKit',
    'playwright לא נמצא. התקינו אותו זמנית והצביעו עליו: npx playwright install webkit, ואז PLAYWRIGHT_PATH=<נתיב ל-playwright> node scripts/qa/caret-webkit-probe.mjs',
  );
  report.print();
  process.exit(0);
}

let probe;
let host;
try {
  probe = writeProbe();
  host = await serveDist();
} catch (error) {
  report.fail('הכנת הדף', String(error && error.message));
  report.print();
  process.exit(1);
}

/**
 * מריץ את אותו תסריט בדיוק בשני המנועים. ההשוואה היא הראיה: „שבור ב-WebKit”
 * לבדו יכול להיות גם „השער לא יודע למדוד ב-WebKit”, ורק ההפרש מול Chromium
 * על אותו קוד מבדיל בין השניים.
 */
async function run(engine, name, baseUrl) {
  const browser = await engine.launch();
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  page.on('console', (m) => { if (m.type() === 'error') console.log(`[${name} console] ${m.text().slice(0, 200)}`); });
  page.on('pageerror', (e) => console.log(`[${name} pageerror] ${String(e).slice(0, 200)}`));
  try {
    await page.goto(baseUrl + '/__qa-webkit-caret.html');
    await page.waitForFunction('!!window.__qa && window.__qa.ready()', null, { timeout: 90_000 });
    await page.waitForTimeout(3_000);

    // מיקום סמן בשורה הראשונה.
    const rect = await page.evaluate('window.__qa.lineRect(0)');
    if (!rect) throw new Error('אין שורת טקסט');
    await page.mouse.click(rect.x, rect.y);
    await page.waitForTimeout(600);

    await page.keyboard.type('שלום', { delay: 60 });
    await page.waitForTimeout(700);
    const afterWord = await page.evaluate(MEASURE);

    await page.keyboard.type(' ', { delay: 60 });
    await page.waitForTimeout(700);
    const afterSpace = await page.evaluate(MEASURE);

    return { afterWord, afterSpace };
  } finally {
    await browser.close();
  }
}

/**
 * המבחן: הסמן על הגבול שאחרי התו האחרון.
 *
 * `prevBoundary` נמדד לצדו כדי להבחין בין „הסמן במקום אחר כלשהו” לבין
 * „הסמן בדיוק תו אחד אחורה” — שני מצבים ששולחים לחפש במקומות שונים.
 */
function judge(label, m) {
  if (m.error) {
    report.fail(label, m.error);
    return;
  }
  const detail = `טקסט=${JSON.stringify(m.text)} ws=${m.ws} סמן=${m.caretX} גבול-סוף=${JSON.stringify(m.endBoundary)} גבול-קודם=${JSON.stringify(m.prevBoundary)}`;
  if (!m.endBoundary || m.endBoundary.length === 0) {
    const atPrev = m.prevBoundary.length ? Math.min(...m.prevBoundary.map((x) => Math.abs(m.caretX - x))) <= 2 : false;
    report.fail(
      label,
      `${detail} — הדפדפן לא החזיר מלבן לגבול שאחרי התו האחרון${atPrev ? ', והסמן יושב על הגבול הקודם: תו אחד לפני' : ''}`,
    );
    return;
  }
  const gap = Math.min(...m.endBoundary.map((x) => Math.abs(m.caretX - x)));
  if (gap <= 2) report.pass(label, `${detail} — סטייה ${gap.toFixed(1)}px`);
  else report.fail(label, `${detail} — סטייה ${gap.toFixed(1)}px`);
}

try {
  for (const [name, engine] of [['chromium', pw.chromium], ['webkit', pw.webkit]]) {
    let out;
    try {
      out = await run(engine, name, `http://127.0.0.1:${host.port}`);
    } catch (error) {
      report.fail(`${name} — הרצה`, String(error && error.message));
      continue;
    }
    console.log(`\n--- ${name} ---\nאחרי מילה: ${JSON.stringify(out.afterWord)}\nאחרי רווח: ${JSON.stringify(out.afterSpace)}`);
    judge(`${name} — הסמן אחרי מילה (ביקורת)`, out.afterWord);
    judge(`${name} — הסמן אחרי רווח`, out.afterSpace);
  }
} finally {
  rmSync(probe, { force: true });
  host.server.close();
}

process.exit(report.print() > 0 ? 1 : 0);
