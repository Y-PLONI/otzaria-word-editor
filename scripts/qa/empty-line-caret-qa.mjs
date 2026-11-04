/**
 * שער רגרסיה ל„הסמן בשורה חדשה אינו מיושר עם השורה שמעליה”.
 *
 * מה שדווח: אחרי Enter בעברית הסמן עומד מעט שמאלה מתחילת השורה שמעליו. השורש
 * שנמדד (superdoc 2.23.0): בשורה ריקה המנוע מצייר ריצת NBSP ומעמיד את הסמן
 * בקצה השמאלי שלה, גם כשתחילת השורה היא הקצה הימני. העקיפה ב-styles/shell.css.
 *
 * מה שהשער מודד, על ה-dist הארוז ובהקשות אמיתיות:
 *
 *   1. המסלול שדווח: הקלדה ו-Enter — הסמן על הקצה הימני של השורה שמעל.
 *   2. התו הראשון שמוקלד בשורה מתחיל באותו קצה.
 *   3. במרכוז הסמן על מרכז השורה, בשורה לטינית על הקצה השמאלי.
 *   4. גופן גדול ו-`w:spacing` על סימן הפסקה — אותו דבר.
 *   5. בריווח כפול הסמן בגובה הטקסט ולא בגובה השורה: זה מה שמבדיל את התיקון
 *      מ-`inline-block` ברוחב אפס.
 *
 * שימוש:  node scripts/qa/empty-line-caret-qa.mjs   (QA_PORT דורס 9736)
 */
import { openApp, createReport, sleep } from './harness.mjs';
import { buildDocx, p } from './docx-fixtures.mjs';

const report = createReport('סמן בשורה ריקה — יישור עם השורה שמעליה', { strict: true });
const app = await openApp({ name: 'empty-line-caret', port: Number(process.env.QA_PORT ?? 9736) });

/** סטייה מותרת: מיקומי המנוע שברי פיקסל, והתקלה שנמדדה היא 4px ומעלה. */
const TOLERANCE = 0.6;
const TEXT = 'שורה עם טקסט';
const BIDI = '<w:bidi/>';
const CENTER = '<w:bidi/><w:jc w:val="center"/>';
const DOUBLE = '<w:bidi/><w:spacing w:line="480" w:lineRule="auto"/>';
const SIZE = '<w:rPr><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr>';
const TRACKED_MARK = '<w:bidi/><w:rPr><w:spacing w:val="40"/></w:rPr>';

const emptyPara = (pPr) => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}</w:p>`;
const bigPara = `<w:p><w:pPr>${BIDI}</w:pPr><w:r>${SIZE}<w:t>${TEXT}</w:t></w:r></w:p>`;

/** כל שורה ריקה יושבת מיד אחרי שורת טקסט באותה צורה, ושתיהן נמדדות יחד. */
const CASES = [
  { name: 'ימין', body: p(TEXT, BIDI) + emptyPara(BIDI), expect: (a) => a.ink[1] },
  { name: 'מרכז', body: p(TEXT, CENTER) + emptyPara(CENTER), expect: (a) => (a.box[0] + a.box[1]) / 2 },
  { name: 'ריווח כפול', body: p(TEXT, DOUBLE) + emptyPara(DOUBLE), expect: (a) => a.ink[1], textHeight: true },
  { name: '24pt', body: bigPara + emptyPara(BIDI + SIZE), expect: (a) => a.ink[1] },
  { name: 'ריווח תווים על סימן הפסקה', body: p(TEXT, BIDI) + emptyPara(TRACKED_MARK), expect: (a) => a.ink[1] },
  { name: 'לטינית', body: p('latin line') + emptyPara(''), expect: (a) => a.ink[0] },
];

const caretRect = () =>
  app
    .js(
      `(function(){
        var el = document.querySelector('[data-v2-local-selection-caret]');
        if (!el) return 'null';
        var r = el.getBoundingClientRect();
        return JSON.stringify({ left: r.left, top: r.top, height: r.height });
      })()`,
    )
    .then(JSON.parse);

/** השורות של גוף המסמך, כפי שצוירו: תיבה, קצוות הטקסט, והאם ריקה. */
const paintedLines = () =>
  app
    .js(
      `JSON.stringify(Array.prototype.map.call(document.querySelectorAll('[data-source-node-id] .superdoc-line'), function(line){
        var range = document.createRange();
        range.selectNodeContents(line);
        var ink = range.getBoundingClientRect(), box = line.getBoundingClientRect();
        return { empty: line.getAttribute('data-pm-start') === line.getAttribute('data-pm-end'),
          box: [box.left, box.right], ink: [ink.left, ink.right], top: box.top, bottom: box.bottom };
      }))`,
    )
    .then(JSON.parse);

const near = (a, b) => Math.abs(a - b) <= TOLERANCE;
const fixed = (n) => (typeof n === 'number' ? n.toFixed(2) : String(n));

try {
  const docx = buildDocx({ body: CASES.map((c) => c.body).join('') + emptyPara(BIDI) });
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    Buffer.from(docx).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-empty',url:${JSON.stringify(dataUrl)},name:'empty-line.docx',size:${docx.length},access:'readwrite'}})}`,
  );
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 2500 });
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.clickSel('.open-browse', 0, { after: 9000 });

  for (const [index, c] of CASES.entries()) {
    const row = index * 2 + 1;
    if (index === 0) {
      const line = (await paintedLines())[row];
      await app.clickAt((line.box[0] + line.box[1]) / 2, (line.top + line.bottom) / 2);
    } else {
      // לחיצה שנייה על שורה ריקה נמדדה כלא מזיזה את הסמן; החצים כן.
      await app.press('ArrowDown', 'ArrowDown', 40);
      await sleep(350);
      await app.press('ArrowDown', 'ArrowDown', 40);
    }
    await sleep(1000);

    const lines = await paintedLines();
    const caret = await caretRect();
    const line = lines[row];
    const textLine = lines[row - 1];
    if (!caret || !line?.empty || caret.top < line.top - 1 || caret.top > line.bottom) {
      report.skip(`${c.name}: הסמן בשורה הריקה`, 'הסמן לא נחת על השורה הנמדדת');
      continue;
    }
    const expected = c.expect(textLine);
    if (near(caret.left, expected)) report.pass(`${c.name}: מיקום הסמן`, fixed(caret.left));
    else report.fail(`${c.name}: מיקום הסמן`, `סמן ${fixed(caret.left)}, צפוי ${fixed(expected)}`);

    if (c.textHeight) {
      const lineHeight = line.bottom - line.top;
      if (caret.height < lineHeight * 0.6) report.pass(`${c.name}: גובה הסמן`, fixed(caret.height));
      else report.fail(`${c.name}: גובה הסמן`, `${fixed(caret.height)} מתוך שורה של ${fixed(lineHeight)}`);
    }
  }

  /* אחרון ובאותו מסמך: מסמך שני שנפתח אחרי הקלדה משאיר את הראשון ב-DOM. */
  await app.press('End', 'End', 35, 2);
  await sleep(800);
  await app.type('שלום עולם טוב');
  await sleep(800);
  await app.press('Enter', 'Enter', 13, 0, '\r');
  await sleep(1000);

  const afterEnter = await caretRect();
  const typedLines = await paintedLines();
  const above = typedLines[typedLines.length - 2];
  if (!afterEnter || !above) {
    report.skip('הסמן אחרי Enter', 'לא נמצאו סמן או שורה מעליו');
  } else if (near(afterEnter.left, above.ink[1])) {
    report.pass('הסמן אחרי Enter על קצה השורה שמעליו', fixed(afterEnter.left));
  } else {
    report.fail('הסמן אחרי Enter על קצה השורה שמעליו', `סמן ${fixed(afterEnter.left)}, קצה ${fixed(above.ink[1])}`);
  }

  await app.type('א');
  await sleep(800);
  const withChar = (await paintedLines()).at(-1);
  if (above && near(withChar.ink[1], above.ink[1])) {
    report.pass('התו הראשון מתחיל באותו קצה', fixed(withChar.ink[1]));
  } else {
    report.fail('התו הראשון מתחיל באותו קצה', `תו ${fixed(withChar.ink[1])}, קצה ${fixed(above?.ink[1])}`);
  }
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
