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
 *   6. כל שורה ריקה נבחרת בלחיצה ממשית, גם ברשימות ובטבלאות. בתא טבלה
 *      ההקלדה נבדקת במודל המסמך: התו נכנס לפסקה הריקה, ולא לזו שמעליה.
 *
 * שימוש:  node scripts/qa/empty-line-caret-qa.mjs   (QA_PORT דורס 9736)
 */
import { openApp, createReport, sleep } from './harness.mjs';
import { buildDocx, p, numPr, table } from './docx-fixtures.mjs';

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
const LIST_LTR = numPr(1) + '<w:jc w:val="left"/>';
const LIST_RTL = numPr(1) + BIDI + '<w:jc w:val="left"/>';
const LIST_NESTED = numPr(1, 1) + BIDI + '<w:jc w:val="left"/>';

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
  { name: 'רשימה LTR', body: p('latin list', LIST_LTR) + emptyPara(LIST_LTR), expect: (a) => a.ink[0] },
  { name: 'רשימה RTL', body: p(TEXT, LIST_RTL) + emptyPara(LIST_RTL), expect: (a) => a.ink[1] },
  { name: 'רשימה מקוננת RTL', body: p(TEXT, LIST_NESTED) + emptyPara(LIST_NESTED), expect: (a) => a.ink[1] },
  { name: 'טבלה LTR', body: table(p('latin table') + emptyPara('')), expect: (a) => a.ink[0], typing: true },
  { name: 'טבלה RTL', body: table(p(TEXT, BIDI) + emptyPara(BIDI)), expect: (a) => a.ink[1], typing: true },
  { name: 'טבלה ממורכזת RTL', body: table(p(TEXT, CENTER) + emptyPara(CENTER)), expect: (a) => (a.box[0] + a.box[1]) / 2, typing: true },
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

/** השורות של גוף המסמך: תיבה, קצוות טקסט התוכן (בלי מספור הרשימה), והאם ריקה. */
const paintedLines = () =>
  app
    .js(
      `JSON.stringify(Array.prototype.map.call(document.querySelectorAll('[data-source-node-id] .superdoc-line'), function(line){
        var range = document.createRange(), rects = [];
        var walker = document.createTreeWalker(line, NodeFilter.SHOW_TEXT);
        for (var node = walker.nextNode(); node; node = walker.nextNode()) {
          if (node.parentElement?.closest('.superdoc-list-marker, .superdoc-marker-suffix-space, .superdoc-marker-suffix-tab')) continue;
          range.selectNodeContents(node);
          rects.push(...range.getClientRects());
        }
        var ink = { left: Math.min(...rects.map(r => r.left)), right: Math.max(...rects.map(r => r.right)) };
        var box = line.getBoundingClientRect();
        return { id: line.closest('[data-source-node-id]').getAttribute('data-source-node-id'),
          empty: line.getAttribute('data-pm-start') === line.getAttribute('data-pm-end'),
          box: [box.left, box.right], ink: [ink.left, ink.right], top: box.top, bottom: box.bottom };
      }))`,
    )
    .then(JSON.parse);

const selection = () => app.js('window.__otzariaEditor.superdoc.activeEditor.doc.selection.current()');
const paragraphTexts = (ids) => app.js(`(async function(){
  var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
  return Promise.all(${JSON.stringify(ids)}.map(async function(id){
    var result = await doc.getNodeById({ nodeId: id });
    if (result.node?.kind !== 'paragraph') throw new Error('Missing paragraph ' + id);
    return result.node.paragraph.inlines.map(function(inline){ return inline.run?.text || ''; }).join('');
  }));
})()`);

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

  const initialLines = await paintedLines();
  const resetId = initialLines.at(-1)?.id;
  for (const [index, c] of CASES.entries()) {
    const row = index * 2 + 1;
    const id = initialLines[row]?.id;
    if (!id || !resetId || !initialLines[row]?.empty) {
      report.fail(`${c.name}: הכנת שורה ריקה`, 'הפסקה הצפויה לא צוירה');
      continue;
    }
    // Seed the selection elsewhere, then let only the real click select the tested paragraph.
    await app.js(`(async function(){
      var away = { kind: 'text', blockId: ${JSON.stringify(resetId)}, offset: 0 };
      await window.__otzariaEditor.superdoc.ui.selection.apply({ kind: 'selection', start: away, end: away });
      document.querySelector('[data-source-node-id="${id}"] .superdoc-line').scrollIntoView({ block: 'center' });
    })()`);
    await sleep(350);
    const clickLine = (await paintedLines())[row];
    await app.clickAt((clickLine.box[0] + clickLine.box[1]) / 2, (clickLine.top + clickLine.bottom) / 2);
    await sleep(1000);

    const lines = await paintedLines();
    const caret = await caretRect();
    const line = lines[row];
    const textLine = lines[row - 1];
    const selected = await selection();
    const start = selected.selectionTarget?.start;
    const end = selected.selectionTarget?.end;
    if (selected.empty && start?.blockId === id && end?.blockId === id && start.offset === 0 && end.offset === 0) {
      report.pass(`${c.name}: הלחיצה בחרה את הפסקה הריקה`, `${id}:0`);
    } else {
      report.fail(`${c.name}: הלחיצה בחרה את הפסקה הריקה`, JSON.stringify(selected.selectionTarget));
    }
    if (!caret || !line?.empty || caret.top < line.top - 1 || caret.top > line.bottom) {
      report.fail(`${c.name}: הסמן בשורה הריקה`, 'הסמן לא נחת על השורה הנמדדת');
    } else {
      const expected = c.expect(textLine);
      if (near(caret.left, expected)) report.pass(`${c.name}: מיקום הסמן`, fixed(caret.left));
      else report.fail(`${c.name}: מיקום הסמן`, `סמן ${fixed(caret.left)}, צפוי ${fixed(expected)}`);

      if (c.textHeight) {
        const lineHeight = line.bottom - line.top;
        if (caret.height < lineHeight * 0.6) report.pass(`${c.name}: גובה הסמן`, fixed(caret.height));
        else report.fail(`${c.name}: גובה הסמן`, `${fixed(caret.height)} מתוך שורה של ${fixed(lineHeight)}`);
      }
    }

    if (c.typing) {
      const ids = [textLine.id, id];
      const before = await paragraphTexts(ids);
      await app.type('X');
      await sleep(1200);
      const after = await paragraphTexts(ids);
      if (before[1] === '' && after[0] === before[0] && after[1] === 'X') {
        report.pass(`${c.name}: ההקלדה נכנסה לפסקה הריקה`, JSON.stringify(after));
      } else {
        report.fail(`${c.name}: ההקלדה נכנסה לפסקה הריקה`, JSON.stringify({ before, after }));
      }
      await app.press('z', 'KeyZ', 90, 2);
      await sleep(900);
      const restored = await paragraphTexts(ids);
      if (JSON.stringify(restored) !== JSON.stringify(before)) {
        report.fail(`${c.name}: ביטול ההקלדה`, JSON.stringify({ before, restored }));
      }
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
    report.fail('הסמן אחרי Enter', 'לא נמצאו סמן או שורה מעליו');
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
