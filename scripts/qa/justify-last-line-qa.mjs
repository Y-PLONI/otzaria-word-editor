/** Real-engine regression gate for #4136 and the list/tab regressions in #158.
 * Uses the packaged app, real Range/caret geometry and exported DOCX text.
 * npm run check:justify-last-line (QA_PORT overrides 9742).
 */
import { openApp, createReport } from './harness.mjs';
import { zipStored, table } from './docx-fixtures.mjs';

// Assistant is bundled with the app, including Hebrew glyphs. A calibrated
// text box keeps the overflow cases independent of installed system fonts.
// Set JUSTIFY_QA_FONT='Times New Roman' to repeat the original Word fixture.
const FONT = process.env.JUSTIFY_QA_FONT ?? 'Assistant';
const N = FONT === 'Times New Roman' ? 18 : 24;
let pageWidth = 11906;
let expectedDocumentText = '';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const run = (text) => `<w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/><w:rtl/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const words = (n, k) => [...Array(n).fill('אבגד'), 'ה'.repeat(k)].join(' ');
const para = (text, props = '<w:jc w:val="both"/>') => `<w:p><w:pPr><w:bidi/><w:spacing w:after="120" w:line="240" w:lineRule="auto"/>${props}</w:pPr>${run(text)}</w:p>`;
const numPr = (id) => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr><w:jc w:val="both"/>`;
const LONG = `${Array(3).fill(words(N, 4)).join(' ')} ${words(N, 6)}`;
const CASES = [
  { name: 'שורה טבעית', text: words(N, 5), mode: 'natural' },
  { name: 'כיווץ קטן', text: words(N, 6), mode: 'shrink' },
  { name: 'כיווץ גדול', text: words(N, 8), mode: 'shrink' },
  { name: 'רצף מילים חלופי', text: words(N + 1, 2), mode: 'shrink' },
  { name: 'שבירה לשורה חדשה', text: words(N + 1, 3), mode: 'wrap' },
  { name: 'פסקה ארוכה', text: LONG, mode: 'long', xml: `<w:p><w:pPr><w:bidi/><w:jc w:val="both"/></w:pPr>${run(words(N, 4))}<w:r><w:br/></w:r>${run(words(N, 6))}</w:p>` },
  { name: 'יישור בירושת סגנון', text: words(N, 8), mode: 'shrink', props: '<w:pStyle w:val="Inherited"/>' },
  { name: 'יישור מסגנון ברירת המחדל', text: words(N, 8), mode: 'shrink', props: '' },
  { name: 'יישור לימין', text: words(N, 8), mode: 'untouched', props: '<w:jc w:val="right"/>' },
  { name: 'יישור למרכז', text: words(N, 8), mode: 'untouched', props: '<w:jc w:val="center"/>' },
  { name: 'מספר מחוץ לתיבת הטקסט', text: 'פריט רשימה קצר ותקין', mode: 'untouched', props: numPr(1) },
  { name: 'תבליט מחוץ לתיבת הטקסט', text: 'פריט תבליט קצר ותקין', mode: 'untouched', props: numPr(2) },
  { name: 'רשימה עם שורה ארוכה', text: words(N - 1, 8), mode: 'bounded', props: numPr(1) },
  { name: 'טאב בעל מיקום קבוע', text: 'טקסט לפני הטאב', mode: 'untouched', xml: `<w:p><w:pPr><w:bidi/><w:jc w:val="both"/><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>${run('טקסט לפני הטאב')}<w:r><w:tab/></w:r>${run(words(N - 2, 8))}</w:p>` },
  { name: 'רווחים בסוף', text: `${words(N, 8)}  `, mode: 'bounded' },
  { name: 'רווחים בהתחלה', text: `  ${words(N, 6)}`, mode: 'bounded' },
  { name: 'מספר LTR מחוץ לתיבה', text: 'short list text fits', mode: 'untouched', xml: `<w:p><w:pPr>${numPr(1)}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:t>short list text fits</w:t></w:r></w:p>` },
  { name: 'הזחה תלויה מכוונת', text: 'שורה עם הזחה תלויה מכוונת', mode: 'untouched', props: '<w:jc w:val="both"/><w:ind w:right="0" w:hanging="360"/>' },
];
function fixture() {
  const styles = `<w:styles ${W}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Inherited"><w:name w:val="Inherited"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="table" w:styleId="Conditional"><w:name w:val="Conditional"/><w:tblStylePr w:type="firstRow"><w:pPr><w:jc w:val="right"/></w:pPr></w:tblStylePr></w:style></w:styles>`;
  const level = (fmt, marker) => `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${marker}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>`;
  const numbering = `<w:numbering ${W}><w:abstractNum w:abstractNumId="0">${level('decimal', '%1.')}</w:abstractNum><w:abstractNum w:abstractNumId="1">${level('bullet', '•')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  const conditional = table(para('בקרת טבלה עם ירושה', '')).replace('<w:tblPr>', '<w:tblPr><w:tblStyle w:val="Conditional"/><w:tblLook w:firstRow="1"/>');
  const body = CASES.map((c) => c.xml ?? para(c.text, c.props)).join('') + conditional;
  expectedDocumentText = [...body.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]).join('');
  const documentXml = `<w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="${pageWidth}" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/></w:sectPr></w:body></w:document>`;
  return zipStored({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + ['document', 'styles', 'settings', 'numbering'].map((part) => `<Override PartName="/word/${part}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${part === 'document' ? 'document.main' : part}+xml"/>`).join('') + '</Types>',
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ['styles', 'settings', 'numbering'].map((part, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${part}" Target="${part}.xml"/>`).join('') + '</Relationships>',
    'word/settings.xml': `<w:settings ${W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
    'word/document.xml': documentXml, 'word/styles.xml': styles, 'word/numbering.xml': numbering,
  });
}

// Independent measurement of the actual painted text. Marker/tab boxes have
// intentional widths and overhangs, so they are not text overflow.
const inspect = () => {
  const paragraphs = new Map();
  for (const line of document.querySelectorAll('.superdoc-line')) {
    const fragment = line.closest('.superdoc-fragment');
    const block = fragment?.getAttribute('data-source-node-id');
    const box = line.getBoundingClientRect();
    if (!block || !box.width) continue;
    const runs = [...line.querySelectorAll('.superdoc-text-run')];
    const text = runs.map((r) => r.textContent).join('');
    const first = text.length - text.trimStart().length, last = text.trimEnd().length;
    let offset = 0; const rects = [];
    for (const run of runs) {
      const walker = document.createTreeWalker(run, NodeFilter.SHOW_TEXT); let node;
      while ((node = walker.nextNode())) {
        const from = Math.max(0, first - offset), to = Math.min(node.length, last - offset);
        if (to > from) { const range = document.createRange(); range.setStart(node, from); range.setEnd(node, to); rects.push(...range.getClientRects()); }
        offset += node.length;
      }
    }
    const ink = rects.filter((r) => r.width > 0);
    const left = Math.min(...ink.map((r) => r.left)), right = Math.max(...ink.map((r) => r.right));
    const rtl = getComputedStyle(line).direction === 'rtl';
    const scale = line.offsetWidth / box.width;
    const gap = (rtl ? left - box.left : box.right - right) * scale;
    const row = { text, block, gap, ours: line.dataset.oweSpillSpacing !== undefined, spacing: line.style.wordSpacing,
      marker: !!line.querySelector('.superdoc-list-marker'), tab: !!line.querySelector('.superdoc-tab:not(.superdoc-marker-suffix-tab)'),
      left, right, top: box.top, bottom: box.bottom };
    if (!paragraphs.has(block)) paragraphs.set(block, []);
    paragraphs.get(block).push(row);
  }
  return JSON.stringify([...paragraphs.values()]);
};

const report = createReport('יישור שורה אחרונה — פריסה ועריכה', { strict: true });
const app = await openApp({ name: 'justify-last-line', port: Number(process.env.QA_PORT ?? 9742) });
const read = async () => JSON.parse(await app.js(`(${inspect})()`));
const check = (name, passed, detail = '') => passed ? report.pass(name, detail) : report.fail(name, detail);
let blocks;
async function place(index, start, end = start) {
  const blockId = blocks[index][0].block;
  return app.js(`(async()=>{const sd=window.__qa.sd(),seed=sd.ui.selection.getSnapshot().selectionTarget;const p={...seed.start,kind:'text',blockId:${JSON.stringify(blockId)},offset:${start},story:{kind:'story',storyType:'body'}};const target={...seed,kind:'selection',start:p,end:{...p,offset:${end}},story:{kind:'story',storyType:'body'}};const r=await sd.ui.selection.apply(target);if(!r.ok)throw new Error(JSON.stringify(r));return true})()`);
}
// A known UTF-16 offset in our single-run paragraph, measured after transforms.
const geometry = async (index, start, end = start + 1) => {
  const block = blocks[index][0].block;
  return JSON.parse(await app.js(`(()=>{const line=[...document.querySelectorAll('.superdoc-line')].find(l=>l.closest('.superdoc-fragment')?.getAttribute('data-source-node-id')===${JSON.stringify(block)});const node=line.querySelector('.superdoc-text-run').firstChild;const r=document.createRange();r.setStart(node,${start});r.setEnd(node,${end});const ink=r.getBoundingClientRect();const sd=window.__qa.sd(),anchor=sd.activeEditor.host.getSelectionAnchorRect(),caret=document.querySelector('.sd-v2-local-selection-caret')?.getBoundingClientRect();return JSON.stringify({lineLeft:line.getBoundingClientRect().left,ink:{left:ink.left,right:ink.right,top:ink.top,bottom:ink.bottom},anchor:anchor?{left:anchor.left,right:anchor.right}:null,caret:caret?.left})})()`));
};

try {
  const calibration = JSON.parse(await app.js(`(async()=>{await document.fonts.load('16px '+${JSON.stringify(FONT)});const ctx=document.createElement('canvas').getContext('2d');ctx.font='16px "'+${JSON.stringify(FONT)}+'"';ctx.fontKerning='none';const width=ctx.measureText(${JSON.stringify(words(N, 6))}).width-2.5;const candidates=${JSON.stringify(Array.from({length:10},(_,i)=>({k:i+1,text:words(N+1,i+1)})))};const alternative=candidates.find(c=>{const d=ctx.measureText(c.text).width-width;return d>1&&d<8});return JSON.stringify({width,k:alternative?.k})})()`));
  pageWidth = Math.round(calibration.width * 15) + 2880;
  CASES[3].text = calibration.k ? words(N + 1, calibration.k) : words(N, 6);
  CASES[4].text = words(N + 1, 12);
  const docx = fixture();
  const dataUrl = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + docx.toString('base64');
  await app.js(`window.__qaHost.replies['fs.pickUserFile']=()=>Promise.resolve({success:true,error:null,data:{token:'tok-justify',url:${JSON.stringify(dataUrl)},name:'justify.docx',size:${docx.length},access:'readwrite'}})`);
  await app.tab('קובץ'); await app.click('פתח קובץ', { after: 500 });
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.click('עיון בקבצים…', { after: 8000 });
  for (let waited = 0; waited < 40_000 && await app.exists('.status-load'); waited += 250) await app.sleep(250);
  await app.sleep(1500); blocks = await read();
  check('כל פסקאות הבקרה נטענו', blocks.length === CASES.length, `${blocks.length}/${CASES.length}`);
  for (const [index, scenario] of CASES.entries()) {
    const lines = blocks[index] ?? [], last = lines.at(-1);
    let pass = !!last;
    if (scenario.mode === 'shrink') pass &&= lines.length === 1 && last.ours && Math.abs(last.gap) <= 0.6;
    if (scenario.mode === 'natural' || scenario.mode === 'untouched') pass &&= lines.every((l) => !l.ours);
    if (scenario.mode === 'wrap') pass &&= lines.length === 2 && lines.every((l) => !l.ours);
    if (scenario.mode === 'long') pass &&= lines.length >= 2 && lines.slice(0, -1).every((l) => !l.ours) && last.ours;
    if (scenario.mode === 'bounded') pass &&= lines.every((l) => l.gap >= -0.6) && lines.filter((l) => l.ours).every((l) => Math.abs(l.gap) <= 0.6);
    check(scenario.name, pass, lines.map((l) => `gap=${l.gap.toFixed(2)}, ours=${l.ours}, spacing=${l.spacing || '—'}`).join(' | '));
  }
  const cell = JSON.parse(await app.js(`JSON.stringify([...document.querySelectorAll('.superdoc-line')].filter(l=>l.textContent.includes('בקרת טבלה')).map(l=>({text:l.textContent,ours:l.dataset.oweSpillSpacing!==undefined})))`));
  check('ירושה בטבלה נשארת למנוע', cell.length === 1 && !cell[0].ours, JSON.stringify(cell));
  const docxText = (xml) => [...xml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]).join('');
  const initialExport = await app.docx();
  check('העקיפה אינה משנה את טקסט ה-DOCX', docxText(initialExport['word/document.xml']) === expectedDocumentText);
  for (const alignment of ['right', 'justify', null]) {
    const result = await app.js(`(async()=>{const doc=window.__qa.sd().activeEditor.doc,target={kind:'block',nodeType:'paragraph',nodeId:${JSON.stringify(blocks[2][0].block)}};const receipt=await doc.paragraphs.${alignment === null ? 'clearAlignment' : 'setAlignment'}({target${alignment === null ? '' : ',alignment:' + JSON.stringify(alignment)}});return JSON.stringify(receipt)})()`);
    await app.sleep(600); const changed = (await read())[2];
    check('שינוי יישור חי: '+(alignment ?? 'ירושה'), JSON.parse(result)?.success === true && (alignment === 'right' ? changed.every(l=>!l.ours) : changed.length===1 && changed[0].ours && Math.abs(changed[0].gap)<=0.6), result);
  }
  const offset = CASES[2].text.lastIndexOf(' ') + 3;
  for (const zoom of [1, 0.7, 1.5]) {
    await app.js(`window.__otzariaEditor.ui.zoom.set(${zoom * 100})`); await app.sleep(350);
    await place(2, offset); await app.sleep(300); const caret = await geometry(2, offset);
    check(`סמן בכיווץ, זום ${zoom * 100}%`, Math.abs(caret.caret - caret.ink.right) <= 0.8, JSON.stringify(caret));
    await place(2, offset - 2, offset + 6); await app.sleep(300); const selection = await geometry(2, offset - 2, offset + 6);
    check(`בחירה בכיווץ, זום ${zoom * 100}%`, !!selection.anchor && Math.abs(selection.anchor.left - selection.ink.left) <= 0.8 && Math.abs(selection.anchor.right - selection.ink.right) <= 0.8, JSON.stringify(selection));
    const layout = await read(); check(`פריסה בזום ${zoom * 100}%`, [1, 2, 3, 6, 7].every((i) => layout[i].length === 1 && layout[i][0].ours && Math.abs(layout[i][0].gap) <= 0.6));
  }
  await app.js('window.__otzariaEditor.ui.zoom.set(100)'); await app.sleep(250);
  await place(2, offset); await app.sleep(250);
  let previous = await geometry(2, offset);
  for (let step = 1; step <= 3; step++) {
    await app.press('ArrowLeft', 'ArrowLeft', 37); await app.sleep(180);
    const state = JSON.parse(await app.js(`JSON.stringify(window.__qa.sd().ui.selection.getSnapshot().selectionTarget)`));
    const current = await geometry(2, state.end.offset);
    check(`חץ שמאלה ${step}`, state.end.offset === offset + step && current.caret - current.lineLeft < previous.caret - previous.lineLeft && Math.abs(current.caret - current.ink.right) <= 0.8, JSON.stringify({ offset: state.end.offset, ...current }));
    previous = current;
  }
  // Real pointer hit followed by input; export verifies document insertion,
  // not only the painted glyphs. Scroll the character into the viewport first.
  await app.js(`(()=>{const id=${JSON.stringify(blocks[2][0].block)};[...document.querySelectorAll('.superdoc-fragment')].find(f=>f.getAttribute('data-source-node-id')===id)?.scrollIntoView({block:'center'})})()`);
  await app.sleep(250); const point = await geometry(2, offset);
  await app.clickAt(point.ink.right - 1, (point.ink.top + point.ink.bottom) / 2); await app.sleep(300);
  await app.type('ש'); await app.sleep(800);
  const entries = await app.docx();
  const xml = entries['word/document.xml'];
  const text = docxText(xml);
  const expected = CASES[2].text.slice(0, offset) + 'ש' + CASES[2].text.slice(offset);
  check('לחיצה והקלדה מוסיפות במקום הנכון ב-DOCX', text === expectedDocumentText.replace(CASES[2].text, expected));
  await app.press('Backspace', 'Backspace', 8); await app.sleep(700);
  const erased = await app.docx(); check('מחיקה משחזרת את כל טקסט ה-DOCX', docxText(erased['word/document.xml']) === expectedDocumentText);
  let edited = await read();
  for (let waited = 0; waited < 3000 && !(edited[2].length === 1 && edited[2][0].ours && Math.abs(edited[2][0].gap) <= 0.6); waited += 100) { await app.sleep(100); edited = await read(); }
  check('הקלדה ומחיקה משאירות את הכיווץ תקין', edited[2].length === 1 && edited[2][0].ours && Math.abs(edited[2][0].gap) <= 0.6, JSON.stringify(edited[2]));
} finally { app.close(); }
process.exit(report.print() > 0 ? 1 : 0);
