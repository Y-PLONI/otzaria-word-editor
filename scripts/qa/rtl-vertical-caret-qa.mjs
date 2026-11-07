/**
 * שער: חצים אנכיים בשורה עברית נוחתים בקצה הנכון.
 *
 * הגשש `rtl-vertical-caret-probe.mjs` מדד את התקלה; זה השער שקובע שהיא
 * מתוקנת ונשארת מתוקנת. שניהם על אותו מסמך, וההבדל הוא מה נדרש: הגשש מתאר,
 * והשער **מפיל את הריצה**.
 *
 * ## מה נדרש בכל מקרה
 *
 * לכל שורת יעד נמדדות חמש עמודות התחלה לאורך השורה שמעליה (או שמתחתיה), וכל
 * אחת חייבת לנחות במקום **הקרוב ביותר בשורת היעד** לעמודה שיצאה ממנה:
 *
 *   • עמודה מעבר לסוף השורה  → הסמן בסוף השורה
 *   • עמודה לפני תחילת השורה → הסמן בתחילת השורה
 *   • עמודה בתוך השורה       → הסמן באותה עמודה, עד רוחב תו
 *
 * וזה נדרש **גם** מהבקרות הלועזיות: שם המנוע נמדד תקין, ולכן השער תופס גם
 * עקיפה שדולפת אל מה שלא היה שבור.
 *
 * ## ועמודת המטרה
 *
 * ארוכה → קצרה → ארוכה חייבת לחזור לעמודה המקורית. זו הדרישה שהעקיפה עלולה
 * לשבור דווקא מפני שהיא כותבת בחירה: נמדד ש-`setSelectionTarget` דורס את
 * עמודת המטרה של המנוע. בלי השורה הזאת בשער, תיקון שמתקן הקשה אחת ושובר את
 * הבאה אחריה היה עובר.
 *
 *   npm run check:vertical-arrows          (QA_PORT עוקף 9719)
 */
import { openApp, createReport } from './harness.mjs';
import { zipStored } from './docx-fixtures.mjs';
import { runDocumentEndCaretQa } from './rtl-document-end-caret-qa.mjs';

const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const RTL = '<w:bidi/>';

const p1 = (text, pPr = '', rPr = '<w:rtl/>') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}` +
  `<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const stylesXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>`;

/**
 * רמת מספור אחת, עם מפריד הטאב שהוא ברירת המחדל של Word.
 *
 * פריט רשימה נמצא כאן מפני שהוא הפסקה שהתנהגה **אחרת** מכל השאר בחצים
 * האופקיים (‏`rtl-caret.ts`: „חזותי בפסקה, לוגי ברשימה”), והוא גם היחיד
 * שמצייר בשורה משהו שאינו טקסט המסמך — סמן המספור וטאב-הסיומת שלו, שאין להם
 * טווח pm. אם הקצה של השורה ייקח אותם בחשבון, עמודה שנופלת על אזור המספור
 * תיחשב „בתוך השורה” ותימסר למנוע.
 */
const numPr = (numId) => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`;

const numberingXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W}>
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

function buildDocx(bodyXml) {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>${bodyXml}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:bidi/></w:sectPr></w:body></w:document>`;
  return zipStored({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    'word/document.xml': documentXml,
    'word/numbering.xml': numberingXml(),
    'word/styles.xml': stylesXml(),
  });
}

const HE_WRAP =
  'זהו משפט ארוך בעברית שנועד למלא את רוחב השורה כולו ולגלוש אל השורה הבאה כדי ' +
  'שאפשר יהיה למדוד לאן בדיוק החץ מוריד את הסמן כאשר השורה שמתחת קצרה יותר. סוף';
const EN_WRAP =
  'This is a long English sentence meant to fill the whole width of the line and ' +
  'then wrap onto the next one so the landing spot can be measured. End';

const BODY = [
  p1(HE_WRAP, RTL),
  p1('שורה עברית ארוכה שממלאת כמעט את כל רוחב העמוד מצד אחד אל הצד השני ואינה גולשת', RTL),
  p1('שורה קצרה', RTL),
  p1('שורה עברית נוספת שגם היא ממלאת כמעט את כל רוחב העמוד ואינה גולשת אל הבאה', RTL),
  p1('עוד שורה עברית ארוכה שממלאת כמעט את כל רוחב העמוד ואינה גולשת אל הבאה', RTL),
  p1('קצרה וממורכזת', RTL + '<w:jc w:val="center"/>'),
  p1(EN_WRAP, '', ''),
  p1('A long English line that fills nearly the entire width of the page without wrapping', '', ''),
  p1('Short line', '', ''),
  p1('Another long English line that also fills nearly the entire page width here', '', ''),
  p1('Yet another long English line filling nearly the whole width of the page', '', ''),
  p1('Short centered', '<w:jc w:val="center"/>', ''),
  p1('שורה עברית ארוכה אחרונה שממלאת כמעט את כל רוחב העמוד ואינה גולשת הלאה', RTL),
  p1('פריט קצר', RTL + numPr(1)),
  p1('שורת מקור לבדיקה שומרת עמודת מטרה דרך פסקה ריקה', RTL),
  p1('יעד קצר לפני ריקה', RTL),
  `<w:p><w:pPr><w:bidi/></w:pPr></w:p>`,
  p1('יעד קצר אחרי ריקה', RTL),
].join('');

const docx = buildDocx(BODY);
const report = createReport('חצים אנכיים: הנחיתה בקצה הנכון', { strict: true });
const app = await openApp({ name: 'rtl-vert-qa', port: Number(process.env.QA_PORT ?? 9719) });

async function openFixture() {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    Buffer.from(docx).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-vq',url:${JSON.stringify(dataUrl)},name:'rtl-vertical-qa.docx',size:${docx.length},access:'readwrite'}})}`,
  );
  await app.tab('קובץ');
  if (!(await app.click('פתח קובץ', { after: 500 }))) throw new Error('„פתח קובץ” לא נמצא');
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.click('עיון בקבצים…', { after: 8000 });
  for (let waited = 0; waited < 40_000; waited += 250) {
    await app.sleep(250);
    if (!(await app.exists('.status-load'))) break;
  }
  await app.sleep(2500);
  for (let attempt = 0; attempt < 8; attempt += 1) {
    try {
      await app.tab('בית');
      break;
    } catch {
      await app.sleep(1500);
    }
  }
  const paras = await app.paraCount();
  console.log(`פתיחה: paraCount=${paras}`);
  return paras >= 14;
}

/**
 * כל השורות המצוירות, מקובצות לפי ה-y של הריצות שבהן.
 *
 * הבורר והסינון כאן חייבים להיות **אותם** של `lineExtent` במודול: אותו
 * `PM_CARRIER_SELECTOR`, ורק מה שנושא טווח pm. סמן המספור של פריט רשימה
 * מצויר בשורה ואין לו טווח; אילו השער היה סופר אותו והמודול לא, המקרה של
 * פריט הרשימה היה עובר או נכשל לפי מזל גאומטרי ולא לפי התנהגות.
 */
const allLines = () =>
  app
    .js(`(function(){
      var out = [];
      var frags = document.querySelectorAll('.superdoc-fragment');
      for (var i = 0; i < frags.length; i++) {
        var runs = frags[i].querySelectorAll('.superdoc-text-run, .superdoc-tab');
        var lines = [];
        for (var j = 0; j < runs.length; j++) {
          if (!runs[j].hasAttribute('data-pm-start')) continue;
          var r = runs[j].getBoundingClientRect();
          if (!r.width || !r.height) continue;
          var hit = null;
          for (var k = 0; k < lines.length; k++) {
            if (Math.abs(lines[k].top - r.top) < 4) { hit = lines[k]; break; }
          }
          if (!hit) { hit = { top: r.top, bottom: r.bottom, left: r.left, right: r.right, text: '' }; lines.push(hit); }
          hit.left = Math.min(hit.left, r.left);
          hit.right = Math.max(hit.right, r.right);
          hit.bottom = Math.max(hit.bottom, r.bottom);
          hit.text += runs[j].textContent || '';
        }
        for (var m = 0; m < lines.length; m++) {
          lines[m].left = Math.round(lines[m].left * 10) / 10;
          lines[m].right = Math.round(lines[m].right * 10) / 10;
          lines[m].top = Math.round(lines[m].top * 10) / 10;
          lines[m].bottom = Math.round(lines[m].bottom * 10) / 10;
          out.push(lines[m]);
        }
      }
      out.sort(function(a, b) { return a.top - b.top; });
      return JSON.stringify(out);
    })()`)
    .then(JSON.parse);

const caretRect = () =>
  app
    .js(`(function(){
      var el = document.querySelector('.sd-v2-local-selection-caret');
      if (!el) return JSON.stringify({ x: null, y: null });
      var r = el.getBoundingClientRect();
      return JSON.stringify({ x: Math.round(r.left*10)/10, y: Math.round((r.top + r.bottom)/2*10)/10 });
    })()`)
    .then(JSON.parse);

function lineAt(lines, y) {
  if (typeof y !== 'number') return null;
  for (const line of lines) if (y >= line.top - 6 && y <= line.bottom + 6) return line;
  return null;
}

/** הבחירה כפי שהמנוע מדווח אותה — בלוק והיסט, ולא פיקסלים. */
const selectionNow = () =>
  app.js(`(async function(){
    try {
      var ed = window.__otzariaEditor.superdoc.activeEditor;
      var info = await ed.doc.selection.current();
      var seg = info && info.target && info.target.segments && info.target.segments[0];
      return seg ? (seg.blockId + '@' + seg.range.start + '..' + seg.range.end) : 'אין';
    } catch (e) { return 'שגיאה ' + String(e); }
  })()`);

/** הטקסט של הפסקה שמכילה את המחרוזת — מה שבאמת נמצא במסמך. */
const paragraphText = (find) =>
  app.js(`(function(){
    var frags = document.querySelectorAll('[data-source-node-id]');
    for (var i = 0; i < frags.length; i++) {
      var text = frags[i].textContent || '';
      if (text.indexOf(${JSON.stringify(find)}) >= 0) return text;
    }
    return '';
  })()`);

/**
 * מה שנדרש מהנחיתה: המקום הקרוב ביותר בשורת היעד לעמודה שיצאנו ממנה.
 *
 * הסבילות היא **רוחב תו**, מפני שהסמן נוחת על חריץ ולא על פיקסל שרירותי. היא
 * נגזרת מהשורה עצמה (רוחב חלקי מספר התווים) ולא מונחת כקבוע, כדי שהשער לא
 * יתרופף בגופן אחר.
 */
function expectedX(target, goalX, rtl) {
  const startX = rtl ? target.right : target.left;
  const endX = rtl ? target.left : target.right;
  if (rtl ? goalX > startX : goalX < startX) return startX;
  if (rtl ? goalX < endX : goalX > endX) return endX;
  return goalX;
}

async function run({ label, sourceText, sourceLineIndex, key, vk, targetText, rtl }) {
  const lines = await allLines();
  const matches = lines.filter((l) => l.text.indexOf(sourceText) >= 0);
  const sourceLine = sourceLineIndex === 'last' ? matches[matches.length - 1] : matches[sourceLineIndex ?? 0];
  if (!sourceLine) throw new Error(`אין שורת מקור ל„${sourceText.slice(0, 20)}”`);
  const targetProbe = lines.find((l) => l.text.indexOf(targetText) >= 0);
  if (!targetProbe) throw new Error(`אין שורת יעד ל„${targetText.slice(0, 20)}”`);

  const width = sourceLine.right - sourceLine.left;
  const rows = [];
  for (const frac of [0.05, 0.25, 0.5, 0.75, 0.95]) {
    const x = Math.round(sourceLine.left + width * frac);
    const y = Math.round((sourceLine.top + sourceLine.bottom) / 2);
    await app.clickAt(x, y);
    await app.sleep(500);
    const before = await caretRect();

    await app.press(key, key, vk);
    await app.sleep(350);
    const after = await caretRect();

    // הגאומטריה נלכדת מחדש: הקשה עלולה לגלול, והמלבנים הם מלבני מסך.
    const now = await allLines();
    const target = now.find((l) => l.text.indexOf(targetText) >= 0);
    const landedOn = lineAt(now, after.y);
    const inTarget = !!(target && landedOn && landedOn.text.indexOf(targetText) >= 0);
    // רוחב תו ממוצע בשורת היעד, עם רצפה של 4px לשורות קצרות מאוד.
    const charWidth = target ? Math.max((target.right - target.left) / Math.max(target.text.length, 1), 4) : 0;
    const want = target ? expectedX(target, before.x, rtl) : null;
    const gap = inTarget && typeof after.x === 'number' ? Math.abs(after.x - want) : Infinity;
    rows.push({ frac, from: before.x, to: after.x, want, gap, inTarget, tolerance: charWidth * 1.5 });
    console.log(
      `   ${String(Math.round(frac * 100)).padStart(3)}% : ${before.x} → ${after.x}` +
        ` (נדרש ${want === null ? '?' : Math.round(want * 10) / 10}, סטייה ${gap === Infinity ? '∞' : Math.round(gap * 10) / 10})`,
    );
  }

  const missed = rows.filter((r) => !r.inTarget || r.gap > r.tolerance);
  const detail = missed.length
    ? missed.map((r) => `${Math.round(r.frac * 100)}%: ${r.to} במקום ${Math.round(r.want * 10) / 10}`).join('; ')
    : `חמש עמודות, כולן במקום`;
  if (missed.length) report.fail(label, detail);
  else report.pass(label, detail);
}

const CASES = [
  { label: 'עברית: חץ למטה בתוך פסקה גולשת', sourceText: 'זהו משפט ארוך בעברית', sourceLineIndex: 0, key: 'ArrowDown', vk: 40, targetText: 'סוף', rtl: true },
  { label: 'עברית: חץ למטה לפסקה קצרה', sourceText: 'שורה עברית ארוכה שממלאת', sourceLineIndex: 0, key: 'ArrowDown', vk: 40, targetText: 'שורה קצרה', rtl: true },
  { label: 'עברית: חץ למעלה לפסקה קצרה', sourceText: 'שורה עברית נוספת שגם היא', sourceLineIndex: 0, key: 'ArrowUp', vk: 38, targetText: 'שורה קצרה', rtl: true },
  { label: 'עברית: חץ למטה לשורה ממורכזת', sourceText: 'עוד שורה עברית ארוכה', sourceLineIndex: 0, key: 'ArrowDown', vk: 40, targetText: 'קצרה וממורכזת', rtl: true },
  { label: 'בקרה LTR: חץ למטה בתוך פסקה גולשת', sourceText: 'This is a long English', sourceLineIndex: 0, key: 'ArrowDown', vk: 40, targetText: 'End', rtl: false },
  { label: 'בקרה LTR: חץ למטה לפסקה קצרה', sourceText: 'A long English line that fills', sourceLineIndex: 0, key: 'ArrowDown', vk: 40, targetText: 'Short line', rtl: false },
  { label: 'בקרה LTR: חץ למעלה לפסקה קצרה', sourceText: 'Another long English line', sourceLineIndex: 0, key: 'ArrowUp', vk: 38, targetText: 'Short line', rtl: false },
  { label: 'בקרה LTR: חץ למטה לשורה ממורכזת', sourceText: 'Yet another long English line', sourceLineIndex: 0, key: 'ArrowDown', vk: 40, targetText: 'Short centered', rtl: false },
  { label: 'עברית: חץ למטה לפריט רשימה קצר', sourceText: 'שורה עברית ארוכה אחרונה', sourceLineIndex: 0, key: 'ArrowDown', vk: 40, targetText: 'פריט קצר', rtl: true },
];

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1500, height: 1100, deviceScaleFactor: 1, mobile: false });
  await app.sleep(400);
  if (!(await openFixture())) {
    report.fail('פתיחת הקובץ', 'המסמך לא נפתח');
    throw new Error('no doc');
  }

  for (const c of CASES) {
    console.log(`\n== ${c.label}`);
    try {
      await run(c);
    } catch (e) {
      report.stuck(c.label, String(e.message || e).slice(0, 100));
    }
  }

  /* -------- עמודת המטרה שורדת את השורה הקצרה -------- */
  console.log('\n== עמודת המטרה');
  for (const c of [
    { label: 'עברית: ארוכה → קצרה → ארוכה', source: 'שורה עברית ארוכה שממלאת', third: 'שורה עברית נוספת שגם היא', rtl: true },
    { label: 'בקרה LTR: ארוכה → קצרה → ארוכה', source: 'A long English line that fills', third: 'Another long English line', rtl: false },
  ]) {
    try {
      const lines = await allLines();
      const sourceLine = lines.find((l) => l.text.indexOf(c.source) >= 0);
      if (!sourceLine) throw new Error('אין שורת מקור');
      const width = sourceLine.right - sourceLine.left;
      const x = Math.round(c.rtl ? sourceLine.right - width * 0.75 : sourceLine.left + width * 0.75);
      await app.clickAt(x, Math.round((sourceLine.top + sourceLine.bottom) / 2));
      await app.sleep(500);
      const start = await caretRect();
      await app.press('ArrowDown', 'ArrowDown', 40);
      await app.sleep(350);
      const mid = await caretRect();
      await app.press('ArrowDown', 'ArrowDown', 40);
      await app.sleep(350);
      const end = await caretRect();

      const now = await allLines();
      const landedOn = lineAt(now, end.y);
      const back = !!(landedOn && landedOn.text.indexOf(c.third) >= 0);
      const charWidth = landedOn ? Math.max((landedOn.right - landedOn.left) / Math.max(landedOn.text.length, 1), 4) : 8;
      const kept = typeof end.x === 'number' && Math.abs(end.x - start.x) <= charWidth * 1.5;
      const detail = `${start.x} → ${mid.x} → ${end.x}`;
      console.log(`   ${c.label}: ${detail}`);
      if (!back) report.fail(`${c.label} — עמודת המטרה`, `לא הגיע לשורה השלישית (${detail})`);
      else if (!kept) report.fail(`${c.label} — עמודת המטרה אינה נשמרת`, detail);
      else report.pass(`${c.label} — עמודת המטרה נשמרת`, detail);
    } catch (e) {
      report.stuck(`${c.label} — עמודת המטרה`, String(e.message || e).slice(0, 100));
    }
  }

  /* -------- חץ מוחזק: התצלום מפגר, והעקיפה חייבת לא להיתקע -------- */
  /*
   * העקיפה קוראת את הסמן מהתצלום הסינכרוני וכותבת בחירה א-סינכרונית. אם
   * ההקשה הבאה מגיעה לפני שהכתיבה נקלטה, היא מחושבת מהמקום **הישן** — וכותבת
   * שוב את אותו יעד, כלומר הסמן נתקע. זה בדיוק מה שקורה כשמחזיקים את החץ,
   * וזה המקרה שאי אפשר לתפוס בבדיקת יחידה: שם הדמה עונה מיד.
   *
   * המקבילה בחצים האופקיים נמדדה ולא נתקעה (25 הקשות → 25 כתיבות). כאן זה
   * נבדק מחדש, מפני שהמדיניות שונה — יש מצב, ומרגע שהתערבנו כל הקשה היא
   * שלנו.
   */
  console.log('');
  console.log('== חץ מוחזק');
  try {
    const lines = await allLines();
    const start = lines.find((l) => l.text.indexOf('זהו משפט ארוך בעברית') >= 0);
    if (!start) throw new Error('אין שורת פתיחה');
    const width = start.right - start.left;
    await app.clickAt(Math.round(start.left + width * 0.1), Math.round((start.top + start.bottom) / 2));
    await app.sleep(600);
    const before = await caretRect();

    const seen = [before.y];
    for (let i = 0; i < 6; i += 1) {
      await app.press('ArrowDown', 'ArrowDown', 40);
      await app.sleep(70);
      seen.push((await caretRect()).y);
    }
    await app.sleep(500);
    const settled = await caretRect();

    /*
     * שש הקשות = **שש** שורות, לא „לפחות שתיים”. הנוסח הקודם כאן בדק
     * `settled.y > before.y + 40` ומספר ערכי y שונים, ושניהם אינם הדבר: קפיצה
     * כפולה מייצרת יותר ערכים שונים ולא פחות, ו„ירד שתי שורות” עבר ירוק גם
     * כשארבע הקשות נבלעו. הצעד נמדד מהגאומטריה עצמה ולא כקבוע, כדי שהשער לא
     * ייפול על מטריקות גופן אחרות.
     */
    const step = lines.length > 1 ? lines[1].top - lines[0].top : 18.4;
    const steps = [];
    for (let i = 1; i < seen.length; i += 1) {
      const a = seen[i - 1];
      const b = seen[i];
      steps.push(typeof a === 'number' && typeof b === 'number' ? Math.round((b - a) * 10) / 10 : null);
    }
    const good = steps.filter((d) => d !== null && d > step * 0.6 && d < step * 1.6).length;
    const detail =
      `y: ${seen.join(' → ')} (ואחרי המתנה ${settled.y}); ` +
      `צעדים ${steps.join(', ')} מול שורה ${Math.round(step * 10) / 10}`;
    console.log(`   ${detail}`);
    if (good === steps.length) report.pass('חץ מוחזק — שש הקשות, שש שורות', detail);
    else if (steps.some((d) => d === 0)) report.fail('חץ מוחזק — הסמן נתקע', detail);
    else report.fail('חץ מוחזק — צעד שאינו שורה אחת', detail);
  } catch (e) {
    report.stuck('חץ מוחזק', String(e.message || e).slice(0, 100));
  }
  /* -------- עמודת המטרה שורדת מסירה למנוע אחרי שכבר התערבנו -------- */
  console.log('');
  console.log('== מסירה למנוע דרך פסקה ריקה אחרי התערבות');
  try {
    const lines = await allLines();
    const source = lines.find((l) => l.text.indexOf('שורת מקור לבדיקה') >= 0);
    const target = lines.find((l) => l.text.indexOf('יעד קצר אחרי ריקה') >= 0);
    if (!source || !target) throw new Error('לא נמצאו שורות המקור והיעד');

    const x = Math.round(source.left + (source.right - source.left) * 0.15);
    await app.clickAt(x, Math.round((source.top + source.bottom) / 2));
    await app.sleep(500);

    const advance = async (before) => {
      await app.press('ArrowDown', 'ArrowDown', 40);
      let after = before;
      for (let waited = 0; waited < 5_000 && after === before; waited += 100) {
        await app.sleep(100);
        after = await selectionNow();
      }
      if (after === before) throw new Error(`הבחירה לא זזה אחרי החץ (${before})`);
      return after;
    };

    const start = await selectionNow();
    const onShort = await advance(start);
    const onEmpty = await advance(onShort);
    const afterEmpty = await advance(onEmpty);
    const caret = await caretRect();
    const now = await allLines();
    const landedOn = lineAt(now, caret.y);
    const charWidth = Math.max((target.right - target.left) / Math.max(target.text.length, 1), 4);
    const onExpectedLine = !!landedOn && landedOn.text.indexOf('יעד קצר אחרי ריקה') >= 0;
    const atExpectedColumn = Math.abs(caret.x - target.left) <= charWidth * 1.5;
    const detail = `${start} → ${onShort} → ${onEmpty} → ${afterEmpty}; x=${caret.x}, יעד=${target.left}`;
    console.log(`   ${detail}`);

    if (afterEmpty === onEmpty) report.fail('הבחירה לא יצאה מהפסקה הריקה', detail);
    else if (!onExpectedLine) report.fail('החץ שאחרי הפסקה הריקה לא הגיע ליעד', detail);
    else if (!atExpectedColumn) report.fail('עמודת המטרה לא שרדה את המסירה למנוע', detail);
    else report.pass('עמודת המטרה שורדת מסירה למנוע דרך פסקה ריקה', detail);
  } catch (e) {
    report.fail('עמודת המטרה שורדת מסירה למנוע דרך פסקה ריקה', String(e.message || e).slice(0, 120));
  }

  /* -------- המסלול שדווח, מקצה לקצה: מה שנכנס למסמך -------- */
  /*
   * כל מה שלמעלה מודד **פיקסלים**. הנזק שבגללו המודול קיים הוא במסמך: מי
   * שיורד בחץ וממשיך להקליד מקליד בראש הפסקה, ובפסקה הראשונה — בתחילת המסמך.
   * שער שמודד רק את הסמן המצויר עובר ירוק גם על תיקון שמצייר נכון ומשאיר את
   * נקודת ההקלדה של המנוע במקום אחר. לכן המקרה הזה אחרון: הוא **משנה** את
   * המסמך, ולכן הגאומטריה שאחריו כבר אינה זו שנמדדה.
   */
  console.log('');
  console.log('== המסלול שדווח — מה שנכנס למסמך');
  try {
    const lines = await allLines();
    const source = lines.find((l) => l.text.indexOf('שורה עברית ארוכה שממלאת') >= 0);
    const target = lines.find((l) => l.text.indexOf('שורה קצרה') >= 0);
    if (!source || !target) throw new Error('לא נמצאו שתי הפסקאות');

    // עמוק בתוך השורה הארוכה, משמאל לסוף השורה הקצרה — בדיוק הדיווח.
    const width = source.right - source.left;
    await app.clickAt(Math.round(source.left + width * 0.25), Math.round((source.top + source.bottom) / 2));
    await app.sleep(600);
    /*
     * ההמתנה כאן היא **על הבחירה** ולא על השעון, וזה לא ייפור: `sleep(400)`
     * קבוע נמדד מהבהב — באחת משתי ריצות על אותו build בדיוק, הנקודה הוקלדה
     * לפני שהמנוע קלט את הכתיבה שלנו, ולא נכנסה כלל. שער `strict` שמהבהב שובר
     * CI ומלמד להתעלם ממנו, ולכן הוא ממתין לעובדה: הבחירה זזה מהמקום שממנו
     * יצאנו.
     */
    const before = await selectionNow();
    await app.press('ArrowDown', 'ArrowDown', 40);
    let landed = before;
    for (let waited = 0; waited < 5_000 && landed === before; waited += 100) {
      await app.sleep(100);
      landed = await selectionNow();
    }
    if (landed === before) throw new Error(`הבחירה לא זזה אחרי החץ (${before})`);
    console.log(`   אחרי החץ: ${before} → ${landed}`);

    await app.type('.');
    let typed = landed;
    for (let waited = 0; waited < 5_000 && typed === landed; waited += 100) {
      await app.sleep(100);
      typed = await selectionNow();
    }
    console.log(`   אחרי ההקלדה: ${typed}`);

    const text = (await paragraphText('שורה קצרה')).trim();
    console.log(`   הפסקה הקצרה אחרי ההקלדה: ${JSON.stringify(text)}`);
    if (text.startsWith('.')) report.fail('הנקודה נכנסה בראש הפסקה', JSON.stringify(text));
    else if (text.endsWith('.')) report.pass('הנקודה נכנסה בסוף השורה', JSON.stringify(text));
    else report.fail('הנקודה לא נכנסה בסוף השורה', JSON.stringify(text));
  } catch (e) {
    report.stuck('המסלול שדווח', String(e.message || e).slice(0, 100));
  }

  await runDocumentEndCaretQa(app, report);

  /*
   * ‏„תקוע” אינו נספר כשבור ב-`createReport` — ובצדק, מפני שהוא נועד לקפיאת
   * מדידה ולא לכשל פקד. כאן לעומת זאת כל `stuck` מקורו ב„לא נמצאה שורת
   * מקור/יעד”, כלומר הפיקסטורה לא צוירה כמצופה — וזה כשל שחייב להפיל, אחרת
   * שער שאף מקרה שלו לא רץ יוצא ירוק בדיוק כשהתיקון חסר.
   */
  const notMeasured = report.rows.filter((row) => row.verdict === 'תקוע');
  if (notMeasured.length) {
    report.fail('מקרים שלא נמדדו', notMeasured.map((row) => row.name).join('; '));
  }
} finally {
  app.close();
}

report.print();
