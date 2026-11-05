/**
 * שער: חץ ימני אינו מזיז את הסמן שמאלה.
 *
 * זו התכונה היחידה שהמשתמש מרגיש, וזו שהייתה שבורה: המנוע בוחר מודל כיוון
 * לכל פסקה בנפרד, ולכן אותו מקש עשה שני דברים הפוכים בשתי פסקאות סמוכות.
 * העקיפה יושבת ב-`src/engine/rtl-caret.ts`, וההסבר המלא שם.
 *
 * ## מה נבדק, ולמה דווקא כך
 *
 * לא „ההיסט ירד ב-1” אלא **ה-x המצויר של הסמן**. ההיסט הוא מודל פנימי; מה
 * שהמשתמש רואה הוא לאן הסמן זז על המסך, ובשורה דו-כיוונית השניים אינם אותו
 * דבר. לכן האינוריאנטה נמדדת על הציור:
 *
 *   `ArrowRight` לעולם אינו מקטין את x, ו-`ArrowLeft` לעולם אינו מגדיל אותו.
 *
 * והיא נכונה גם בעברית וגם בלטינית — ולכן פסקה לטינית היא בקרה אמיתית כאן
 * ולא מקרה נפרד.
 *
 * צעד שעובר שורה — לבלוק אחר, או לשורה אחרת באותה פסקה גולשת — אינו נמדד
 * ככיוון: שם x קופץ לקצה השני מעצם הפריסה. „מלכודת” — חזרה לאותו (בלוק,
 * היסט) אחרי ביקור בשני — נמדדת בנפרד, כי זה ה„נתקע” שדווח.
 *
 * בחירה נמדדת באותה שיטה: אין בה סמן מצויר, ולכן המדד הוא **מלבן הבחירה**
 * (`getSelectionAnchorRect`) מול ה-x של העוגן. `Shift+שמאל` מרחיב משמאל לעוגן,
 * ו`שמאל` על בחירה קורס לקצה השמאלי שלה. ההיסטים הם מודל פנימי גם כאן.
 *
 * ושני דברים שהאינוריאנטה שלמעלה אינה רואה, ולכן נבדקים בנפרד:
 *   - **תקיעה בתוך הבלוק.** ליד ספרה או אות לטינית בודדת הסמן קפץ בין שני
 *     היסטים שמצוירים באותו x (הפרש 0.1px) — שום צעד אינו „לכיוון ההפוך”,
 *     ובכל זאת הסמן אינו מתקדם. A→B→A→B באותו בלוק הוא כשל.
 *   - **טבלה בין פסקאות.** החץ קפץ מעל הטבלה לפסקה שאחריה. הכניסה לתא היא
 *     של המנוע, והשער מוודא שהיא לא נגזלה.
 *   - **קצות המסמך.** בפסקה ראשונה ואחרונה שיש בהן ספרה (המנוע זז בהן לוגית)
 *     ההקשה שנמסרה למנוע הזיזה לכיוון ההפוך, והבאה החזירה — תנודה בלי סוף.
 *     שם אין לאן לזוז, והסמן חייב להישאר במקומו.
 *
 *   npm run check:arrows            (QA_PORT עוקף 9714)
 */
import { openApp, createReport } from './harness.mjs';
import { table, zipStored } from './docx-fixtures.mjs';

const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const RTL = '<w:bidi/>';
const numPr = (numId, ilvl = 0) =>
  `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;

const pRuns = (runs, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs
    .map(
      (r) =>
        `<w:r>${r.rPr ? `<w:rPr>${r.rPr}</w:rPr>` : ''}${
          r.tab ? '<w:tab/>' : `<w:t xml:space="preserve">${r.t}</w:t>`
        }</w:r>`,
    )
    .join('')}</w:p>`;

const p1 = (text, pPr = '', rPr = '<w:rtl/>') => pRuns([{ t: text, rPr }], pPr);

const lvl = (fmt, text) =>
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>` +
  `<w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>` +
  `<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>`;

const numberingXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W}>
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lvl('decimal', '%1.')}</w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvl('bullet', '•')}</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

const stylesXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
</w:styles>`;

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

const BODY = [
  p1('פתיחה 7 כאן', RTL),
  p1('אבגדהוז ראשון גדול', RTL),
  p1('מספור ראשון ברשימה', RTL + numPr(1)),
  p1('תבליט עגול בפריט', RTL + numPr(2)),
  pRuns([{ t: 'קודם', rPr: '<w:rtl/>' }, { tab: true }, { t: 'אחרון', rPr: '<w:rtl/>' }], RTL),
  p1('פרק 12 בספר הזה', RTL),
  pRuns(
    [
      { t: 'מילה ', rPr: '<w:rtl/>' },
      { t: 'ABC', rPr: '' },
      { t: ' מילה אחרת', rPr: '<w:rtl/>' },
    ],
    RTL,
  ),
  p1('הזחה בלי מספור כאן', RTL + '<w:pStyle w:val="ListParagraph"/>'),
  p1(Array.from({ length: 40 }, (_, i) => `מילה${i}`).join(' '), RTL),
  p1('plain latin paragraph', '', ''),
  p1('סעיף 3 בחוק הזה', RTL),
  pRuns([{ t: 'אות ', rPr: '<w:rtl/>' }, { t: 'a', rPr: '' }, { t: ' אחת בלבד', rPr: '<w:rtl/>' }], RTL),
  p1('שלוש × ארבע שווה', RTL),
  p1('חמש ÷ שתיים', RTL),
  // אות מנוקדת עם טעם בתוך שורה שיש בה גם אי לטיני: „שָׁ” הוא שלוש יחידות
  // UTF-16 ותיבה אחת, והתפר שמפיל את הבסיס משאיר דווקא את ההיסטים שבתוך
  // האשכול. זה הטקסט של העורך הזה — פירוש מנוקד — ולא מקרה קצה.
  pRuns([
    { t: 'מילה ', rPr: '<w:rtl/>' },
    { t: 'abc', rPr: '' },
    { t: 'שָׁלום כאן', rPr: '<w:rtl/>' },
  ], RTL),
  // שורה שמ**תחילה** באי לטיני: התפר שבין כיוון הפסקה לריצה הלטינית יושב
  // על היסט 0, ולכן הוא לא נראה בשורה שמתחילה בעברית.
  pRuns([{ t: 'ABC', rPr: '' }, { t: ' מילה כאן', rPr: '<w:rtl/>' }], RTL),
  // פסקה ריקה בין שתי עבריות: אין לה אף תו מצויר, ולכן אין ממה לקרוא חריץ.
  p1('לפני הריקה 4', RTL),
  pRuns([], RTL),
  p1('אחרי הריקה 5', RTL),
  p1('לפני הטבלה 4', RTL),
  table(p1('תוך התא', RTL)),
  p1('אחרי הטבלה 5', RTL),
].join('');

const CASES = [
  { label: 'פסקה עברית רגילה', find: 'אבגדהוז ראשון גדול' },
  { label: 'רשימה ממוספרת', find: 'מספור ראשון ברשימה' },
  { label: 'רשימת תבליטים', find: 'תבליט עגול בפריט' },
  { label: 'פסקה עם טאב', find: 'קודם' },
  { label: 'עברית עם ספרות', find: 'פרק 12 בספר הזה' },
  { label: 'עברית עם לטינית', find: 'מילה ABC מילה אחרת' },
  { label: 'הזחה בלי מספור', find: 'הזחה בלי מספור כאן' },
  { label: 'פסקה גולשת לשתי שורות', find: 'מילה0 מילה1' },
  { label: 'בקרה: פסקה לטינית', find: 'plain latin paragraph' },
  // מהקצוות ולא מהאמצע, ובעשרה צעדים: כדי לעבור את התו הבודד מכל צד.
  { label: 'ספרה בודדת', find: 'סעיף 3 בחוק הזה', edges: true },
  { label: 'אות לטינית בודדת', find: 'אות a אחת בלבד', edges: true },
  /*
   * שני קצות השורה, ושניהם מהקצה: התפר שבין כיוון הפסקה לאי לועזי שמתחיל או
   * מסיים את השורה נראה **רק** כשמתחילים בקצה ומקישים לשני הכיוונים.
   * ‏„לפני הטבלה 4” כבר היה ב-fixture ולא נבדק — שם החץ הימני נמסר למנוע.
   */
  { label: 'שורה שמתחילה באי לטיני', find: 'ABC מילה כאן', edges: true },
  { label: 'שורה שנגמרת בספרה', find: 'לפני הטבלה 4', edges: true },
  { label: 'אות מנוקדת אחרי אי לטיני', find: 'שָׁלום כאן', edges: true },
];

const docx = buildDocx(BODY);
const report = createReport('חצים אופקיים: ימין אינו מזיז שמאלה', { strict: true });
const app = await openApp({ name: 'rtl-caret', port: Number(process.env.QA_PORT ?? 9714) });

/** „פתח קובץ” פותח דיאלוג; הבורר נקרא רק מ„עיון בקבצים…” שבתוכו. */
async function openFixture() {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    Buffer.from(docx).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-arrow',url:${JSON.stringify(dataUrl)},name:'rtl-caret.docx',size:${docx.length},access:'readwrite'}})}`,
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
  return paras >= 9;
}

/** הסמן: היכן המנוע מצייר אותו, ובאיזה בלוק והיסט הוא לדעתו. */
const caretState = () =>
  app
    .js(`(function(){
      var ed = window.__otzariaEditor && window.__otzariaEditor.superdoc && window.__otzariaEditor.superdoc.activeEditor;
      var snap = ed && ed.host && ed.host.readLiveSelectionSyncSnapshot && ed.host.readLiveSelectionSyncSnapshot();
      var sel = snap && snap.selectionTarget;
      var el = document.querySelector('.sd-v2-local-selection-caret');
      var r = el ? el.getBoundingClientRect() : null;
      return JSON.stringify({
        block: sel && sel.end ? sel.end.blockId : null,
        off: sel && sel.end ? sel.end.offset : null,
        x: r ? Math.round(r.left * 10) / 10 : null,
        y: r ? Math.round(r.top) : null
      });
    })()`)
    .then(JSON.parse);

/** הבחירה דרך ה-API הא-סינכרוני — היחיד שמתאר טווחים (נמדד). */
const asyncSelection = () =>
  app
    .js(`(async function(){
      try {
        var ed = window.__otzariaEditor.superdoc.activeEditor;
        var info = await ed.doc.selection.current();
        var seg = info && info.target && info.target.segments && info.target.segments[0];
        return JSON.stringify({
          empty: info ? info.empty : null,
          start: seg && seg.range ? seg.range.start : null,
          end: seg && seg.range ? seg.range.end : null,
          blocks: info && info.target && info.target.segments ? info.target.segments.length : 0
        });
      } catch (e) { return JSON.stringify({ error: String(e).slice(0, 60) }); }
    })()`)
    .then(JSON.parse);

/**
 * מניחה סמן בקצה של הטקסט: `'left'` או `'right'`.
 */
async function caretEdge(find, side) {
  const rect = JSON.parse(
    await app.js(`(function(){
      var frags = document.querySelectorAll('.superdoc-fragment');
      for (var i = 0; i < frags.length; i++) {
        if ((frags[i].textContent || '').indexOf(${JSON.stringify(find)}) < 0) continue;
        var runs = frags[i].querySelectorAll('.superdoc-text-run');
        var left = Infinity, right = -Infinity, top = 0, bottom = 0;
        for (var j = 0; j < runs.length; j++) {
          var r = runs[j].getBoundingClientRect();
          if (!r.width) continue;
          left = Math.min(left, r.left); right = Math.max(right, r.right);
          top = r.top; bottom = r.bottom;
        }
        if (right < 0) return JSON.stringify(null);
        return JSON.stringify({
          x: Math.round(${JSON.stringify(side)} === 'left' ? left + 1 : right - 1),
          y: Math.round((top + bottom) / 2),
          nodeId: frags[i].getAttribute('data-source-node-id')
        });
      }
      return JSON.stringify(null);
    })()`),
  );
  if (!rect) throw new Error(`אין טקסט מצויר לפסקה „${find}”`);
  await app.clickAt(rect.x, rect.y);
  await app.sleep(600);
  const state = await caretState();
  if (state.block !== rect.nodeId) {
    throw new Error(`הלחיצה נחתה על בלוק ${state.block} ולא על ${rect.nodeId}`);
  }
  return { ...state, nodeId: rect.nodeId };
}

/**
 * מניחה סמן ב**אמצע הטקסט** של הפסקה ולא במלבן השורה.
 *
 * קצה השורה של פריט רשימה הוא אזור ההזחה, ולחיצה שם נוחתת על היסט 0 בכל
 * הפריטים — מה שהיה מסתיר את מה שנמדד כאן. נמדד בגשש הסקר.
 */
async function caretMid(find) {
  const rect = JSON.parse(
    await app.js(`(function(){
      var frags = document.querySelectorAll('.superdoc-fragment');
      for (var i = 0; i < frags.length; i++) {
        if ((frags[i].textContent || '').indexOf(${JSON.stringify(find)}) < 0) continue;
        var runs = frags[i].querySelectorAll('.superdoc-text-run');
        var left = Infinity, right = -Infinity, top = 0, bottom = 0;
        for (var j = 0; j < runs.length; j++) {
          var r = runs[j].getBoundingClientRect();
          if (!r.width) continue;
          left = Math.min(left, r.left); right = Math.max(right, r.right);
          top = r.top; bottom = r.bottom;
        }
        if (right < 0) return JSON.stringify(null);
        return JSON.stringify({
          x: Math.round((left + right) / 2), y: Math.round((top + bottom) / 2),
          nodeId: frags[i].getAttribute('data-source-node-id')
        });
      }
      return JSON.stringify(null);
    })()`),
  );
  if (!rect) throw new Error(`אין טקסט מצויר לפסקה „${find}”`);
  await app.clickAt(rect.x, rect.y);
  await app.sleep(600);
  const state = await caretState();
  if (state.block !== rect.nodeId) {
    throw new Error(`הלחיצה נחתה על בלוק ${state.block} ולא על ${rect.nodeId}`);
  }
  return state;
}

/** מלבן הבחירה כפי שהמנוע מודד אותו, ו-`null` כשאין טווח. */
const selectionRect = async () =>
  JSON.parse(
    await app.js(`(function(){
      var ed = window.__otzariaEditor.superdoc.activeEditor;
      var r = ed.host.getSelectionAnchorRect();
      if (!r || !r.width) return 'null';
      return JSON.stringify({ left: Math.round(r.left * 10) / 10, right: Math.round(r.right * 10) / 10 });
    })()`),
  );

async function walk(key, times, { shift = false } = {}) {
  const vk = key === 'ArrowRight' ? 39 : 37;
  const steps = [];
  for (let i = 0; i < times; i += 1) {
    await app.press(key, key, vk, shift ? 8 : 0);
    await app.sleep(220);
    steps.push(await caretState());
  }
  return steps;
}

const trace = (steps) =>
  steps.map((s) => `${String(s.block ?? '?').slice(-2)}:${s.off ?? '?'}@${s.x ?? '?'}`).join(' → ');

/**
 * ההפרות: צעדים **בתוך אותה שורה מצוירת** שבהם הסמן זז לכיוון ההפוך מהמקש.
 * מעבר בלוק או שורה אינו נמדד — שם x קופץ לקצה השני מעצם הפריסה.
 */
function violations(steps, toRight) {
  const bad = [];
  let moved = 0;
  for (let i = 1; i < steps.length; i += 1) {
    const a = steps[i - 1];
    const b = steps[i];
    if (a.block !== b.block) continue;
    if (typeof a.x !== 'number' || typeof b.x !== 'number') continue;
    if (typeof a.y === 'number' && typeof b.y === 'number' && Math.abs(a.y - b.y) > 2) continue;
    const dx = b.x - a.x;
    if (Math.abs(dx) <= 0.5) continue;
    moved += 1;
    if (toRight ? dx < 0 : dx > 0) bad.push(`${a.off}@${a.x} → ${b.off}@${b.x}`);
  }
  return { bad, moved };
}

/** תקיעה בתוך בלוק: A→B→A→B, כשכל הקשה מחליפה היסט ואינה מתקדמת. */
function oscillations(steps) {
  let hits = 0;
  for (let i = 3; i < steps.length; i += 1) {
    const [a, b, c, d] = [steps[i - 3], steps[i - 2], steps[i - 1], steps[i]];
    if (a.block === b.block && b.block === c.block && c.block === d.block &&
        a.off === c.off && b.off === d.off && a.off !== b.off) hits += 1;
  }
  return hits;
}

/** „נתקע”: חזרה לאותו (בלוק, היסט) אחרי ביקור בבלוק אחר. */
function trapped(steps) {
  let loops = 0;
  for (let i = 2; i < steps.length; i += 1) {
    const a = steps[i - 2];
    const b = steps[i - 1];
    const c = steps[i];
    if (a.block !== b.block && b.block !== c.block && a.block === c.block && a.off === c.off) {
      loops += 1;
    }
  }
  return loops;
}

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1500, height: 1000, deviceScaleFactor: 1, mobile: false });
  await app.sleep(400);
  if (!(await openFixture())) {
    report.fail('פתיחת הקובץ', 'המסמך לא נפתח');
    throw new Error('no doc');
  }

  for (const c of CASES) {
    let right;
    let left;
    try {
      if (c.edges) {
        right = [await caretEdge(c.find, 'left'), ...(await walk('ArrowRight', 10))];
        left = [await caretEdge(c.find, 'right'), ...(await walk('ArrowLeft', 10))];
      } else {
        const start = await caretMid(c.find);
        right = [start, ...(await walk('ArrowRight', 6))];
        await caretMid(c.find);
        left = [start, ...(await walk('ArrowLeft', 6))];
      }
    } catch (error) {
      report.stuck(c.label, String(error.message || error).slice(0, 90));
      continue;
    }

    console.log(`\n== ${c.label}`);
    console.log(`   ArrowRight: ${trace(right)}`);
    console.log(`   ArrowLeft : ${trace(left)}`);

    const r = violations(right, true);
    const l = violations(left, false);
    const loops = trapped(right) + trapped(left);
    const shaking = oscillations(right) + oscillations(left);

    if (shaking) {
      report.fail(c.label, `הסמן נתקע: ${shaking} תנודות בין שני היסטים באותו בלוק`);
    } else if (r.bad.length || l.bad.length) {
      report.fail(
        c.label,
        `ימין הזיז שמאלה ב-${r.bad.length} צעדים (${r.bad.slice(0, 2).join('; ')}), ` +
          `שמאל הזיז ימינה ב-${l.bad.length}`,
      );
    } else if (!r.moved || !l.moved) {
      report.fail(c.label, `הסמן לא זז כלל (ימין ${r.moved} צעדים, שמאל ${l.moved})`);
    } else if (loops) {
      report.fail(c.label, `מלכודת גבול: ${loops} חזרות לאותו היסט`);
    } else {
      report.pass(c.label, `${r.moved} צעדים ימינה, ${l.moved} שמאלה, אף אחד לכיוון ההפוך`);
    }
  }

  /* -------- סימן ניטרלי: החץ עובר דרכו תו-תו, בלי לדלג -------- */
  /*
   * `×` ו-`÷` הם ON ב-UBA — הדפדפן מצייר אותם בכיוון הפסקה, ולכן שני
   * התפרים שלהם נמצאים ב-x שונה ושניהם יעדים. זו אינה הנקודה העיוורת של אי
   * של תו אחד, וקיפול שלהם מוחק מקום נגיש: כל הצעדים חייבים להיות של היסט אחד.
   */
  for (const [label, find] of [
    ['סימן כפל', 'שלוש × ארבע שווה'],
    ['סימן חלוק', 'חמש ÷ שתיים'],
  ]) {
    try {
      const steps = [await caretEdge(find, 'right'), ...(await walk('ArrowLeft', 8))];
      console.log(`\n== ${label}: ${trace(steps)}`);
      const jumps = [];
      for (let i = 1; i < steps.length; i += 1) {
        if (steps[i].block !== steps[i - 1].block) continue;
        const delta = steps[i].off - steps[i - 1].off;
        if (delta !== 1) jumps.push(`${steps[i - 1].off}→${steps[i].off}`);
      }
      jumps.length
        ? report.fail(`${label} — החץ דילג`, `צעדים שאינם היסט אחד: ${jumps.join(', ')}`)
        : report.pass(`${label} — תו-תו`, `${steps.length - 1} הקשות, כל אחת היסט אחד`);
    } catch (error) {
      report.stuck(label, String(error.message || error).slice(0, 90));
    }
  }

  /* -------- אות מנוקדת: הסמן אינו נכנס בין האות לניקוד שלה -------- */
  /*
   * זו האינוריאנטה שהמונוטוניות שלמעלה **אינה** רואה: „שָׁ” הוא אשכול אחד
   * ושלוש יחידות pm, וחריץ בהיסט 9 או 10 יושב בין ש לקמץ שלה — הקשה שם
   * מכניסה תו לתוך האות. הגבולות נגזרים כאן מ-`Intl.Segmenter` על הטקסט
   * שהמנוע צייר, ולא מרשימה קשיחה.
   */
  try {
    const find = 'שָׁלום כאן';
    const legal = JSON.parse(
      await app.js(`(function(){
        var frags = document.querySelectorAll('[data-source-node-id][data-pm-start]');
        for (var i = 0; i < frags.length; i++) {
          var text = frags[i].textContent || '';
          if (text.indexOf(${JSON.stringify(find)}) < 0) continue;
          var seg = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
          var out = [];
          var it = seg.segment(text)[Symbol.iterator]();
          for (var r = it.next(); !r.done; r = it.next()) out.push(r.value.index);
          out.push(text.length);
          return JSON.stringify({ nodeId: frags[i].getAttribute('data-source-node-id'), text: text, bounds: out });
        }
        return JSON.stringify(null);
      })()`),
    );
    if (!legal) throw new Error('הפסקה המנוקדת אינה מצוירת');
    const steps = [
      await caretEdge(find, 'right'),
      ...(await walk('ArrowLeft', 12)),
      await caretEdge(find, 'left'),
      ...(await walk('ArrowRight', 12)),
    ];
    console.log(`
== אות מנוקדת: ${trace(steps)}`);
    console.log(`   גבולות אשכול: ${legal.bounds.join(',')}`);
    const inside = steps
      .filter((t) => t.block === legal.nodeId && typeof t.off === 'number' && !legal.bounds.includes(t.off))
      .map((t) => t.off);
    inside.length
      ? report.fail(
          'אות מנוקדת — הסמן בתוך אשכול',
          `היסטים שאינם גבול אשכול: ${[...new Set(inside)].join(', ')} (גבולות: ${legal.bounds.join(',')})`,
        )
      : report.pass('אות מנוקדת — כל עצירה על גבול אשכול', `${steps.length} עצירות, ${legal.bounds.length} גבולות`);
  } catch (error) {
    report.stuck('אות מנוקדת', String(error.message || error).slice(0, 90));
  }

  /* -------- טבלה בין פסקאות: הכניסה לתא נשארת של המנוע -------- */
  try {
    const before = await caretEdge('לפני הטבלה', 'left');
    const after = await caretEdge('אחרי הטבלה', 'right');
    await caretEdge('לפני הטבלה', 'left');
    await app.press('End', 'End', 35);
    await app.sleep(300);
    await app.press('ArrowLeft', 'ArrowLeft', 37);
    await app.sleep(500);
    const down = await caretState();
    await caretEdge('אחרי הטבלה', 'right');
    await app.press('Home', 'Home', 36);
    await app.sleep(300);
    await app.press('ArrowRight', 'ArrowRight', 39);
    await app.sleep(500);
    const up = await caretState();
    console.log(`\n== טבלה: שמאל מסוף „לפני” → ${down.block}; ימין מתחילת „אחרי” → ${up.block}`);
    const skipped = [];
    if (down.block === after.nodeId || down.block === before.nodeId) skipped.push(`שמאל נחת ב-${down.block}`);
    if (up.block === before.nodeId || up.block === after.nodeId) skipped.push(`ימין נחת ב-${up.block}`);
    skipped.length
      ? report.fail('טבלה בין פסקאות', `החץ לא נכנס לתא: ${skipped.join('; ')}`)
      : report.pass('טבלה בין פסקאות', 'שני הכיוונים נכנסים לתא');
  } catch (error) {
    report.stuck('טבלה בין פסקאות', String(error.message || error).slice(0, 90));
  }

  /* -------- פסקה ריקה: יש לה מקום סמן אחד, והחץ חייב לעצור בו -------- */
  /*
   * לפסקה בלי תווים אין אף אלמנט שנושא טווח pm, ולכן אין ממה לגזור חריץ —
   * והקוד החזיר „לא שלנו”. מסירה למנוע אינה ניטרלית כאן: בפסקה לוגית הוא זז
   * קדימה, כלומר החץ הימני מדלג **מעל** הריקה ויורד שורה.
   */
  try {
    const empty = JSON.parse(
      await app.js(`(function(){
        var frags = document.querySelectorAll('[data-source-node-id][data-pm-start]');
        for (var i = 0; i < frags.length; i++) {
          if ((frags[i].textContent || '').indexOf('אחרי הריקה 5') < 0) continue;
          var prev = frags[i].previousElementSibling;
          var kids = [];
          if (prev) for (var k = 0; k < prev.children.length; k++) {
            var kid = prev.children[k];
            kids.push({
              tag: kid.tagName,
              cls: String(kid.className).slice(0, 40),
              dir: kid.getAttribute('dir'),
              pm: kid.getAttribute('data-pm-start') + '..' + kid.getAttribute('data-pm-end'),
              carriers: kid.querySelectorAll('.superdoc-text-run, .superdoc-tab').length,
              text: JSON.stringify(kid.textContent || '').slice(0, 24)
            });
          }
          return JSON.stringify({
            id: prev ? prev.getAttribute('data-source-node-id') : null,
            pm: prev ? prev.getAttribute('data-pm-start') + '..' + prev.getAttribute('data-pm-end') : null,
            kids: kids,
            next: frags[i].getAttribute('data-source-node-id')
          });
        }
        return JSON.stringify(null);
      })()`),
    );
    if (!empty || !empty.id) throw new Error('הפסקה הריקה לא נמצאה ב-DOM');
    console.log(
      `\n== פסקה ריקה: id=${empty.id} pm=${empty.pm}\n` +
        (empty.kids ?? [])
          .map((k) => `   <${k.tag}.${k.cls}> dir=${k.dir} pm=${k.pm} נושאי-היסט=${k.carriers} ${k.text}`)
          .join('\n'),
    );

    await caretEdge('אחרי הריקה 5', 'right');
    await app.press('Home', 'Home', 36);
    await app.sleep(300);
    await app.press('ArrowRight', 'ArrowRight', 39);
    await app.sleep(500);
    const up = await caretState();

    await caretEdge('לפני הריקה 4', 'left');
    await app.press('End', 'End', 35);
    await app.sleep(300);
    await app.press('ArrowLeft', 'ArrowLeft', 37);
    await app.sleep(500);
    const down = await caretState();

    console.log(`   ימין מתחילת „אחרי” → ${up.block}:${up.off}; שמאל מסוף „לפני” → ${down.block}:${down.off}`);
    const missed = [];
    if (up.block !== empty.id) missed.push(`ימין נחת ב-${up.block} ולא בריקה`);
    if (down.block !== empty.id) missed.push(`שמאל נחת ב-${down.block} ולא בריקה`);
    missed.length
      ? report.fail('פסקה ריקה בין שתי עבריות', missed.join('; '))
      : report.pass('פסקה ריקה בין שתי עבריות', 'שני הכיוונים עוצרים בה');
  } catch (error) {
    report.stuck('פסקה ריקה בין שתי עבריות', String(error.message || error).slice(0, 90));
  }

  /* -------- קצות המסמך: אין לאן לזוז, והסמן נשאר -------- */
  for (const [label, find, side, homeKey, key] of [
    ['תחילת המסמך', 'פתיחה 7 כאן', 'right', ['Home', 36], ['ArrowRight', 39]],
    ['סוף המסמך', 'אחרי הטבלה 5', 'left', ['End', 35], ['ArrowLeft', 37]],
  ]) {
    try {
      await caretEdge(find, side);
      await app.press(homeKey[0], homeKey[0], homeKey[1]);
      await app.sleep(300);
      const edge = [await caretState()];
      for (let i = 0; i < 5; i += 1) {
        await app.press(key[0], key[0], key[1]);
        await app.sleep(220);
        edge.push(await caretState());
      }
      console.log(`\n== ${label}: ${trace(edge)}`);
      const moved = edge.filter((s) => s.block !== edge[0].block || s.off !== edge[0].off).length;
      moved === 0
        ? report.pass(`${label} — הסמן נשאר במקומו`, `${edge[0].off}, ${edge.length - 1} הקשות`)
        : report.fail(`${label} — הסמן זז`, `${moved} מתוך ${edge.length - 1} הקשות: ${trace(edge)}`);
    } catch (error) {
      report.stuck(label, String(error.message || error).slice(0, 90));
    }
  }

  /* -------- בחירה: לאיזה צד של העוגן היא גדלה -------- */
  const near = (a, b) => Math.abs(a - b) < 2.5;
  const shiftLeft = (n) => walk('ArrowLeft', n, { shift: true });
  const shiftRight = (n) => walk('ArrowRight', n, { shift: true });
  /* טאב אינו כאן: `getSelectionAnchorRect` מחזיר בשורה עם טאב מלבן שאינו על
     השורה (נמדד: 51..385 לשורה ב-578..653), אף שההיסטים נכונים. הוא נבדק למטה
     בהיסטים. */
  const SELECTION_CASES = CASES.filter((x) =>
    /רגילה|ממוספרת|תבליטים|ספרות|עם לטינית|הזחה/.test(x.label),
  );

  /* עוגן על תפר מצויר בבית שהמנוע בחר, וכתיבתו חזרה מציירת אותו בבית השני
     (ראו „שני דברים שהמדידה מצאה” ב-`rtl-caret.ts`). באמצע „עברית עם לטינית”
     הלחיצה נוחתת על תפר כזה, ולכן הולכים ממנו לעברית לפני שמודדים. */
  const anchorAt = async (find) => {
    await caretMid(find);
    await walk('ArrowLeft', 2);
    return caretState();
  };

  for (const c of SELECTION_CASES) {
    try {
      const { x: anchor } = await anchorAt(c.find);
      await shiftLeft(3);
      const left = await selectionRect();
      console.log(`\n== ${c.label} — עוגן ${anchor}, Shift+Left×3: ${JSON.stringify(left)}`);
      if (left && left.left < anchor - 1 && near(left.right, anchor)) {
        report.pass(`${c.label} — Shift+שמאל מרחיב שמאלה`);
      } else {
        report.fail(`${c.label} — Shift+שמאל מרחיב שמאלה`, `עוגן ${anchor}, בחירה ${JSON.stringify(left)}`);
      }

      await app.press('ArrowRight', 'ArrowRight', 39);
      await app.sleep(300);
      const right = await caretState();
      if (left && near(right.x, left.right)) {
        report.pass(`${c.label} — ימין קורס לקצה הימני`);
      } else {
        report.fail(`${c.label} — ימין קורס לקצה הימני`, `x=${right.x}, בחירה ${JSON.stringify(left)}`);
      }

      await anchorAt(c.find);
      await shiftRight(3);
      const grown = await selectionRect();
      console.log(`   Shift+Right×3: ${JSON.stringify(grown)}`);
      if (grown && grown.right > anchor + 1 && near(grown.left, anchor)) {
        report.pass(`${c.label} — Shift+ימין מרחיב ימינה`);
      } else {
        report.fail(`${c.label} — Shift+ימין מרחיב ימינה`, `עוגן ${anchor}, בחירה ${JSON.stringify(grown)}`);
      }

      await app.press('ArrowLeft', 'ArrowLeft', 37);
      await app.sleep(300);
      const leftEdge = await caretState();
      if (grown && near(leftEdge.x, grown.left)) {
        report.pass(`${c.label} — שמאל קורס לקצה השמאלי`);
      } else {
        report.fail(`${c.label} — שמאל קורס לקצה השמאלי`, `x=${leftEdge.x}, בחירה ${JSON.stringify(grown)}`);
      }

      /* חזרה דרך העוגן: הראש חולף אותו, והעוגן עצמו אינו זז. */
      await anchorAt(c.find);
      await shiftLeft(2);
      await shiftRight(4);
      const passed = await selectionRect();
      if (passed && passed.right > anchor + 1 && near(passed.left, anchor)) {
        report.pass(`${c.label} — הראש חולף את העוגן`);
      } else {
        report.fail(`${c.label} — הראש חולף את העוגן`, `עוגן ${anchor}, בחירה ${JSON.stringify(passed)}`);
      }
    } catch (error) {
      report.stuck(`${c.label} — בחירה`, String(error.message || error).slice(0, 90));
    }
  }

  /* -------- בחירה בעכבר, ואז מקלדת -------- */
  /*
   * התצלום הסינכרוני אינו מוסר אותה. הקצוות נקראים בסיום הגרירה, והכיוון
   * נגזר ממיקומם מול תנועת העכבר; המקש עצמו פועל רק על בחירה שכבר נמדדה.
   */
  const LIST = 'מספור ראשון ברשימה';
  const textSpan = async (find) => JSON.parse(await app.js(`(function(){
    var frags = document.querySelectorAll('.superdoc-fragment');
    for (var i = 0; i < frags.length; i++) {
      if ((frags[i].textContent || '').indexOf(${JSON.stringify(find)}) < 0) continue;
      var r = frags[i].querySelector('.superdoc-text-run').getBoundingClientRect();
      return JSON.stringify({ left: r.left, right: r.right, y: (r.top + r.bottom) / 2 });
    }
    return 'null';
  })()`));

  for (const [label, fromRight] of [['עם סדר המסמך', true], ['נגד סדר המסמך', false]]) {
    try {
      const span = await textSpan(LIST);
      const third = (span.right - span.left) / 3;
      const [from, to] = fromRight
        ? [span.right - third, span.left + third]
        : [span.left + third, span.right - third];
      await app.drag(from, span.y, to, span.y);
      const before = await selectionRect();
      await shiftLeft(1);
      const after = await selectionRect();
      console.log(`\n== גרירה ${label}: ${JSON.stringify(before)} → ${JSON.stringify(after)}`);
      /* ‏`Shift+שמאל` מזיז את הראש שמאלה: בגרירה עם הסדר הראש בשמאל והבחירה
         גדלה, ובגרירה נגדו הראש בימין והיא קטנה. הקצה הרחוק נשאר בשני המקרים. */
      const ok = before && after && (fromRight
        ? after.left < before.left - 1 && near(after.right, before.right)
        : after.right < before.right - 1 && near(after.left, before.left));
      if (ok) report.pass(`גרירה ${label} — Shift+שמאל מזיז את הראש`);
      else report.fail(`גרירה ${label} — Shift+שמאל מזיז את הראש`, `${JSON.stringify(before)} → ${JSON.stringify(after)}`);

      await app.drag(from, span.y, to, span.y);
      const dragged = await selectionRect();
      await app.press('ArrowRight', 'ArrowRight', 39);
      await app.sleep(300);
      const edge = await caretState();
      if (dragged && near(edge.x, dragged.right)) report.pass(`גרירה ${label} — ימין קורס לקצה הימני`);
      else report.fail(`גרירה ${label} — ימין קורס לקצה הימני`, `x=${edge.x}, ${JSON.stringify(dragged)}`);
    } catch (error) {
      report.stuck(`גרירה ${label}`, String(error.message || error).slice(0, 90));
    }
  }

  /* -------- מעבר פסקאות, החזקה, והחלפת הבחירה -------- */
  try {
    await caretEdge('מספור ראשון ברשימה', 'right');
    await shiftRight(2);
    const across = await asyncSelection();
    console.log(`\n== Shift+ימין מעבר לפסקה: ${JSON.stringify(across)}`);
    if (across.blocks === 2) report.pass('Shift+ימין עובר לפסקה הקודמת');
    else report.fail('Shift+ימין עובר לפסקה הקודמת', JSON.stringify(across));

    /* עברית נקייה: שם כל הקשה היא היסט אחד, ואין תפרים שמדולגים. */
    await caretMid('אבגדהוז ראשון גדול');
    for (let i = 0; i < 6; i += 1) await app.press('ArrowLeft', 'ArrowLeft', 37, 8);
    await app.sleep(500);
    const heldKey = await asyncSelection();
    console.log(`== שש Shift+שמאל ברצף: ${JSON.stringify(heldKey)}`);
    if (heldKey.end - heldKey.start === 6) report.pass('הקשות רצופות אינן הולכות לאיבוד');
    else report.fail('הקשות רצופות אינן הולכות לאיבוד', JSON.stringify(heldKey));

    /* טאב: לפי היסטים — ימינה בשורה עברית הוא אחורה. */
    await caretMid('קודם');
    const tabFrom = await caretState();
    await shiftRight(2);
    const tabbed = await asyncSelection();
    console.log(`== טאב: מהיסט ${tabFrom.off}, ${JSON.stringify(tabbed)}`);
    if (tabbed.end === tabFrom.off && tabbed.start < tabFrom.off) report.pass('פסקה עם טאב — Shift+ימין נע אחורה');
    else report.fail('פסקה עם טאב — Shift+ימין נע אחורה', `${tabFrom.off} ${JSON.stringify(tabbed)}`);

    await caretMid('סעיף 3 בחוק הזה');
    await shiftLeft(3);
    await app.type('ש');
    await app.sleep(600);
    const typed = await app.js(`(function(){
      var f = Array.from(document.querySelectorAll('.superdoc-fragment')).find(function (el) { return (el.textContent || '').indexOf('סעיף') >= 0; });
      return f ? f.textContent : '';
    })()`);
    console.log(`== הקלדה על בחירה: ${typed}`);
    if (typed.length === 'סעיף 3 בחוק הזה'.length - 2) report.pass('הקלדה מחליפה בחירה שנכתבה כאן');
    else report.fail('הקלדה מחליפה בחירה שנכתבה כאן', typed);
  } catch (error) {
    report.stuck('מעבר פסקאות והחזקה', String(error.message || error).slice(0, 90));
  }

  /* -------- מה שנשאר למנוע -------- */
  try {
    await caretMid('plain latin paragraph');
    const before = await caretState();
    await shiftRight(3);
    const latin = await selectionRect();
    console.log(`\n== לטינית: עוגן ${before.x}, ${JSON.stringify(latin)}`);
    if (latin && latin.right > before.x + 1) report.pass('שורה לטינית — Shift+ימין מרחיב ימינה');
    else report.fail('שורה לטינית — Shift+ימין מרחיב ימינה', JSON.stringify(latin));

    const cell = await textSpan('תוך התא');
    await app.drag(cell.right - 4, cell.y, cell.left + 4, cell.y);
    const selected = await asyncSelection();
    await app.press('ArrowLeft', 'ArrowLeft', 37);
    await app.sleep(300);
    const collapsed = await asyncSelection();
    console.log(`== תא טבלה: ${JSON.stringify(selected)} → ${JSON.stringify(collapsed)}`);
    if (selected.empty === false && collapsed.empty === true) report.pass('תא טבלה — חץ על בחירה אינו אובד');
    else report.fail('תא טבלה — חץ על בחירה אינו אובד', `${JSON.stringify(selected)} → ${JSON.stringify(collapsed)}`);
  } catch (error) {
    report.stuck('מה שנשאר למנוע', String(error.message || error).slice(0, 90));
  }
} catch (error) {
  console.error(error);
  report.fail('הריצה', String(error.message || error).slice(0, 120));
} finally {
  report.print();
  app.close();
}
