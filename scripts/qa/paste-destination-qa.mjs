/** Focused real-engine regressions for destination formatting and async paste notifications. */
import assert from 'node:assert/strict';
import { openApp } from './harness.mjs';
const app = await openApp({ name: 'paste-destination-qa', port: Number(process.env.QA_PORT ?? 9499) });
// Clipboard promises and native editing commands can stall under headless focus
// contention. Surface a failed measurement instead of hanging the QA entrypoint.
const send = app.cdp.send.bind(app.cdp);
app.cdp.send = async (method, params) => {
  let timer;
  try {
    return await Promise.race([send(method, params), new Promise((_, reject) => {
      timer = setTimeout(() => reject(new Error(`QA transport deadline: ${method}`)), 45_000);
    })]);
  } finally { clearTimeout(timer); }
};
const text = (xml) => [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((item) => item[1]).join('');
async function menu(label) {
  assert(await app.openMenu('אפשרויות הדבקה'));
  await app.js(`(()=>{const label=${JSON.stringify(label)};const button=[...document.querySelectorAll('.ribbon-menu__item')].find(e=>e.querySelector('.ribbon-menu__item-label')?.textContent.trim()===label);button?.scrollIntoView({block:'nearest'});})()`);
  await app.sleep(150);
  assert(await app.clickMenu(label));
}
async function clipboard(value) { assert.equal(await app.js(`navigator.clipboard.writeText(${JSON.stringify(value)}).then(()=> 'ok')`), 'ok'); }
async function ctrlV() {
  const modifiers = /Mac|iPhone|iPad/.test(await app.js('navigator.platform')) ? 4 : 2;
  await app.cdp.send('Input.dispatchKeyEvent', { type: 'rawKeyDown', key: 'v', code: 'KeyV', windowsVirtualKeyCode: 86, nativeVirtualKeyCode: modifiers === 4 ? 9 : 86, modifiers, commands: ['paste'] });
  await app.cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'v', code: 'KeyV', windowsVirtualKeyCode: 86, modifiers });
  await app.sleep(1800);
}
try {
  await app.cdp.send('Page.bringToFront');
  await app.cdp.send('Browser.grantPermissions', { permissions: ['clipboardReadWrite', 'clipboardSanitizedWrite'] });
  await app.caret(0);
  await app.js(`window.__qaHost.replies['library.resolveRef'] = () => Promise.resolve({ success:true,data:[{id:42,bookId:'פסחים',bookUid:'id:42',type:'text',title:'פסחים',reference:'פסחים דף לד',index:1234,isPdf:false,isSourceLine:true,isUserBook:false,bookPath:'שס'}],error:null });
    window.__nativeEvents=[]; document.addEventListener('keydown',e=>window.__nativeEvents.push({type:e.type,key:e.key,code:e.code,meta:e.metaKey,ctrl:e.ctrlKey,target:e.target.tagName,inEditor:window.__qa.sd().element.contains(e.target)}),true); document.addEventListener('paste',e=>window.__nativeEvents.push({type:e.type,types:[...e.clipboardData?.types??[]],trusted:e.isTrusted}),true);
    window.__pasteEvents=[]; document.addEventListener('otzaria:paste-started',()=>window.__pasteEvents.push('start')); document.addEventListener('otzaria:paste-finished',()=>window.__pasteEvents.push('finish'));`);
  await menu('Ctrl+V: שמירת טקסט בלבד');
  await app.openMenu('אפשרויות הדבקה');
  const radio = await app.js(`([...document.querySelectorAll('[role=menuitemradio]')].find(e=>e.textContent.trim().startsWith('Ctrl+V: שמירת טקסט בלבד')))?.getAttribute('aria-checked')`);
  assert.equal(radio, 'true', 'the new default is actually selected');
  await app.escape();
  await app.caret(0);
  await app.js(`window.__qa.sd().ui.commands.executeAsync('font-family', 'David')`);
  await app.js(`window.__qa.sd().ui.commands.executeAsync('font-size', 18)`);
  await app.sleep(1500);
  await clipboard('@פסחים לד');
  await ctrlV();
  const blankFiles = await app.docx();
  const blankXml = blankFiles['word/document.xml'];

  assert.equal(text(blankXml), '@פסחים לד', 'blank Normal paste inserts once');
  assert.match(blankXml, /w:ascii="David"/);
  assert.match(blankXml, /w:cs="David"/);
  assert.match(blankXml, /<w:sz w:val="36"/);
  assert.match(blankXml, /<w:szCs w:val="36"/);
  assert.equal(await app.js(`JSON.stringify(window.__pasteEvents)`), '["start","finish"]');
  assert.equal(await app.js(`window.__nativeEvents.some(e=>e.type==='paste'&&e.trusted&&e.types.includes('text/plain'))`), true);
  assert.equal(await app.js(`!!document.querySelector('.otzaria-at-mention [role="option"]')`), true, 'new @ mention opens after async insertion');
  console.log('PASS blank Normal + pending David18 + trusted browser paste + bubbling completion + @ suggestions');

  await app.js(`(async()=>{const sd=window.__qa.sd(),d=sd.activeEditor.doc;await d.clearContent({});await d.styles.apply({target:{scope:'docDefaults',channel:'run'},patch:{bold:true,boldCs:true,color:{val:'FF0000'}}});})()`);
  await app.caret(0);
  await app.js(`(async()=>{const sd=window.__qa.sd();let r=await sd.activeEditor.doc.clipboard.insert({target:sd.ui.selection.getSnapshot().selectionTarget,fragment:{kind:'superdoc.clipboard.fragment',version:'v2.1',blocks:[{kind:'paragraph',runs:[{text:'DEST',marks:{fontFamily:'Arial',fontSizePt:12,fontSizeCsPt:12,bold:false,boldCs:false,italic:false,italicCs:false,underline:false,strike:false,color:'#000000'}}]}]}});let t=r.effects.insertedText[0].selectionTarget;t.start={...t.end};await sd.ui.selection.apply(t)})()`);
  await app.sleep(1000);
  await clipboard('PASTED');
  await menu('שמירת טקסט בלבד');
  const xml = (await app.docx())['word/document.xml'];
  const pasted = [...xml.matchAll(/<w:r>(?:<w:rPr>(.*?)<\/w:rPr>)?<w:t(?: [^>]*)?>([^<]*)<\/w:t><\/w:r>/g)].find((run) => run[2].includes('PASTED'));
  assert(pasted, 'text was pasted');
  assert.match(pasted[1], /<w:color w:val="000000"/);
  assert.match(pasted[1], /<w:b w:val="0"/);
  assert.match(pasted[1], /<w:bCs w:val="0"/);
  console.log('PASS explicit black + Latin/CS bold-off over red/bold defaults');

  // Direct uniform fonts equal to docDefaults must override the themed Heading1 font.
  await app.js(`(async()=>{const sd=window.__qa.sd(),d=sd.activeEditor.doc;let r=await d.clipboard.insert({target:sd.ui.selection.getSnapshot().selectionTarget,fragment:{kind:'superdoc.clipboard.fragment',version:'v2.1',blocks:[{kind:'paragraph',styleRef:'Heading1',runs:[{text:'HEAD',marks:{fontFamily:'Arial',fontSizePt:12,fontSizeCsPt:12,bold:false,boldCs:false,color:'#000000'}}]}]}});let t=r.effects.insertedText[0].selectionTarget;t.start={...t.end};await sd.ui.selection.apply(t)})()`);
  await app.sleep(1000);
  await clipboard('INHEADING');
  await menu('שמירת טקסט בלבד');
  const headingXml = (await app.docx())['word/document.xml'];
  const headingRun = [...headingXml.matchAll(/<w:r>(?:<w:rPr>(.*?)<\/w:rPr>)?<w:t(?: [^>]*)?>([^<]*)<\/w:t><\/w:r>/g)].find((run) => run[2].includes('INHEADING'));
  assert(headingRun);
  assert.match(headingRun[1], /w:ascii="Arial"/);
  assert.match(headingRun[1], /w:cs="Arial"/);
  assert.match(headingRun[1], /<w:sz w:val="24"/);
  console.log('PASS direct Arial12 equal to defaults in differently styled Heading1');
  await app.js(`navigator.clipboard.write([new ClipboardItem({ 'text/html':new Blob(['<p>רגיל <b>מודגש</b> <i>נטוי</i></p>'],{type:'text/html'}), 'text/plain':new Blob(['רגיל מודגש נטוי'],{type:'text/plain'}) })])`);
  await menu('Ctrl+V: מיזוג עיצוב');
  await app.press('End', 'End', 35);
  await app.sleep(1200);
  await ctrlV();
  const emphasisXml = (await app.docx())['word/document.xml'];
  const emphasisRuns = [...emphasisXml.matchAll(/<w:r>(?:<w:rPr>(.*?)<\/w:rPr>)?<w:t(?: [^>]*)?>([^<]*)<\/w:t><\/w:r>/g)];
  const boldRun = emphasisRuns.find((run) => run[2].includes('מודגש'));
  const italicRun = emphasisRuns.find((run) => run[2].includes('נטוי'));
  const plainRun = emphasisRuns.find((run) => run[2].includes('רגיל'));
  assert(boldRun && italicRun && plainRun);
  assert.match(boldRun[1], /<w:bCs(?:\s*\/| w:val="1")/);
  assert.match(italicRun[1], /<w:iCs(?:\s*\/| w:val="1")/);
  assert.match(plainRun[1], /<w:bCs w:val="0"/);
  assert.match(plainRun[1], /<w:iCs w:val="0"/);
  console.log('PASS HTML local emphasis retained in Latin and Hebrew, plain neighbors stay off');
  // Keep-source uses the existing transport/plan and keeps source appearance.
  await app.js(`navigator.clipboard.write([new ClipboardItem({ 'text/html':new Blob(['<p><span style="color:#ff0000;font-size:20pt;font-family:Arial">BASELINE</span></p>'],{type:'text/html'}), 'text/plain':new Blob(['BASELINE'],{type:'text/plain'}) })])`);
  await menu('שמירה על עיצוב המקור');
  const baselineXml = (await app.docx())['word/document.xml'];
  const baselineRun = [...baselineXml.matchAll(/<w:r>(?:<w:rPr>(.*?)<\/w:rPr>)?<w:t(?: [^>]*)?>([^<]*)<\/w:t><\/w:r>/g)].find((run) => run[2].includes('BASELINE'));
  assert(baselineRun);
  assert.match(baselineRun[1], /<w:color w:val="FF0000"/);
  assert.match(baselineRun[1], /<w:sz w:val="40"/);
  console.log('PASS keep-source preserves the original rich source formatting');

  const rangeSeed = await app.js(`(async()=>{const sd=window.__qa.sd(),t=sd.ui.selection.getSnapshot().selectionTarget;window.__guardCaret=t;const r={...t,start:{...t.start,offset:t.start.offset-1}};window.__guardRange=r;await sd.ui.selection.apply(r);return (await sd.activeEditor.doc.info()).revision})()`);
  await app.sleep(1200);
  await menu('מיזוג עיצוב');
  assert.equal(await app.js(`(async()=> (await window.__qa.sd().activeEditor.doc.info()).revision)()`), rangeSeed);
  assert.match(await app.js(`window.__qa.status().text`), /ההדבקה לא בוצעה/);
  console.log('PASS selected range is refused before mutation');

  const splitSeed = await app.js(`(async()=>{const sd=window.__qa.sd(),d=sd.activeEditor.doc;const receipt=await d.format.apply({target:window.__guardRange,inline:{rFonts:{ascii:'Arial',hAnsi:'Arial',cs:'David'}}});if(receipt.success!==true)throw new Error(JSON.stringify(receipt));await sd.ui.selection.apply(window.__guardCaret);return (await d.info()).revision})()`);
  await app.sleep(1200);
  await menu('שמירת טקסט בלבד');
  assert.notEqual(await app.js(`(async()=> (await window.__qa.sd().activeEditor.doc.info()).revision)()`), splitSeed);
  assert.doesNotMatch(await app.js(`window.__qa.status().text`) ?? '', /ההדבקה לא בוצעה/);
  const splitRun = [...(await app.docx())['word/document.xml'].matchAll(/<w:r>(?:<w:rPr>(.*?)<\/w:rPr>)?<w:t(?: [^>]*)?>([^<]*)<\/w:t><\/w:r>/g)].filter((run) => run[2].includes('BASELINE')).pop();
  assert.match(splitRun[1], /w:ascii="Arial"/, 'Latin text takes the Latin channel of a split destination');
  console.log('PASS split Latin/CS destination: Latin text pasted with the Latin font');

  async function splitDestination(seed = true) {
    const capture = await app.js(`(async()=>{const sd=window.__qa.sd(),d=sd.activeEditor.doc;if(${seed}){const r=await d.format.apply({target:window.__guardRange,inline:{rFonts:{ascii:'Arial',hAnsi:'Arial',cs:'David'},fontSize:12,fontSizeCs:14,color:'000000'}});if(r.success!==true)throw new Error(JSON.stringify(r));}const selection=await sd.ui.selection.apply(window.__guardCaret);if(selection.ok!==true)throw new Error(JSON.stringify(selection));return (await d.selection.extractOoxml({at:window.__guardRange})).fragment.xml})()`);
    assert.match(capture, /<w:sz w:val="24"/);
    assert.match(capture, /<w:szCs w:val="28"/);
    await app.sleep(1200);
  }
  await splitDestination();
  await app.js(`(async()=>{const c=window.__qa.sd().ui.commands;await c.executeAsync('font-family','David');await c.executeAsync('font-size',14)})()`);
  const pending = JSON.parse(await app.js(`JSON.stringify({font:window.__qa.sd().ui.commands.get('font-family').getState(),size:window.__qa.sd().ui.commands.get('font-size').getState()})`));
  assert.equal(pending.font.value, 'David');
  assert.equal(pending.font.active, true);
  assert.equal(String(pending.size.value), '14');
  assert.equal(pending.size.active, true);
  await clipboard('PENDING-PICK');
  await menu('שמירת טקסט בלבד');
  const pendingRun = [...(await app.docx())['word/document.xml'].matchAll(/<w:r>(?:<w:rPr>(.*?)<\/w:rPr>)?<w:t(?: [^>]*)?>([^<]*)<\/w:t><\/w:r>/g)].find((run) => run[2].includes('PENDING-PICK'));
  assert(pendingRun);
  assert.match(pendingRun[1], /w:ascii="David"/);
  assert.match(pendingRun[1], /<w:sz w:val="28"/);
  console.log('PASS pending David14 equal to CS channel overrides Latin Arial12');

  await splitDestination(false);
  await app.js(`navigator.clipboard.write([new ClipboardItem({'text/html':new Blob(['<p>abc <span style="color:#00ff00">9876</span></p>'],{type:'text/html'}),'text/plain':new Blob(['abc 9876'],{type:'text/plain'})})])`);
  await menu('מיזוג עיצוב');
  const numberRun = [...(await app.docx())['word/document.xml'].matchAll(/<w:r>(?:<w:rPr>(.*?)<\/w:rPr>)?<w:t(?: [^>]*)?>([^<]*)<\/w:t><\/w:r>/g)].find((run) => run[2].includes('9876'));
  assert(numberRun);
  assert.match(numberRun[1], /w:ascii="Arial"/);
  assert.match(numberRun[1], /<w:sz w:val="24"/);
  assert.doesNotMatch(numberRun[1], /<w:color w:val="00FF00"/);
  console.log('PASS numeric HTML span inherits preceding Latin script after source color is dropped');
} catch (error) {
  console.error('QA STATE', await app.js(`JSON.stringify({selection:window.__qa.sd().ui.selection.getSnapshot(),events:window.__pasteEvents,native:window.__nativeEvents,focus:document.hasFocus(),active:document.activeElement?.outerHTML.slice(0,400),viewport:{w:innerWidth,h:innerHeight},status:window.__qa.status()})`));
  console.error('QA ERRORS', await app.log());
  throw error;
} finally { app.close(); }
