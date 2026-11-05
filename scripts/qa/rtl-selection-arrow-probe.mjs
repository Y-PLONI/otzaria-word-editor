/**
 * גשש: לאן זז הסמן כשיש בחירה — `Shift+חץ`, וחץ רגיל על בחירה קיימת.
 *
 * המדד הוא המסך ולא ההיסט: בכל הקשה נקרא טווח הבחירה ומלבני הצבע המצוירים,
 * ומהם ניכר לאיזה צד הבחירה גדלה. בנוסף — מה זמין **סינכרונית** בזמן בחירה,
 * כי זו השאלה שמכריעה אם אפשר ליירט.
 */
import { openApp } from './harness.mjs';
import { zipStored } from './docx-fixtures.mjs';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const p = (t, pPr = '<w:bidi/>') => `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const body = [
  p('אבגדהוז ראשון גדול'),
  p('פרק 12 בספר הזה'),
  p('מספור ראשון ברשימה', '<w:bidi/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
  `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>קודם</w:t></w:r><w:r><w:tab/></w:r><w:r><w:rPr><w:rtl/></w:rPr><w:t>אחרון</w:t></w:r></w:p>`,
].join('');
const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:bidi/></w:sectPr></w:body></w:document>`;
const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
const docx = zipStored({
  '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
  '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`,
  'word/document.xml': documentXml,
  'word/numbering.xml': numbering,
});

const app = await openApp({ name: 'sel-arrow', port: Number(process.env.QA_PORT ?? 9741) });

const state = () =>
  app.js(`(async function(){
    var ed = window.__otzariaEditor.superdoc.activeEditor;
    var info = await ed.doc.selection.current();
    var seg = info && info.target && info.target.segments && info.target.segments[0];
    var snap = ed.host.readLiveSelectionSyncSnapshot();
    var t = snap && snap.selectionTarget;
    var caret = document.querySelector('.sd-v2-local-selection-caret');
    var cr = caret && caret.getBoundingClientRect();
    var rects = [];
    document.querySelectorAll('[class*="sd-v2-local-selection"]').forEach(function(el){
      if (el === caret) return;
      var r = el.getBoundingClientRect();
      if (r.width) rects.push(Math.round(r.left*10)/10 + '..' + Math.round(r.right*10)/10);
    });
    return JSON.stringify({ r: seg ? seg.range : null, n: info && info.target && info.target.segments && info.target.segments.length,
      sync: t ? [t.start && t.start.offset, t.end && t.end.offset] : null,
      caretX: cr ? Math.round(cr.left*10)/10 : null, rects: rects, snapKeys: snap ? Object.keys(snap) : null });
  })()`);

const shapes = () =>
  app.js(`(function(){
    var ed = window.__otzariaEditor.superdoc.activeEditor;
    var sel = ed.ui && ed.ui.selection;
    var protoKeys = function (o) { var k = new Set(); for (var c = o; c && c !== Object.prototype; c = Object.getPrototypeOf(c)) Object.getOwnPropertyNames(c).forEach(function (n) { k.add(n); }); return Array.from(k).filter(function (n) { return /sel|caret|anchor|head|range/i.test(n); }); };
    return JSON.stringify({ host: protoKeys(ed.host), authoring: protoKeys(ed.authoring), ed: protoKeys(ed), ui: sel ? protoKeys(sel) : null });
  })()`);

const L = ['ArrowLeft', 'ArrowLeft', 37];
const R = ['ArrowRight', 'ArrowRight', 39];
const shift = async (k) => { await app.press(k[0], k[1], k[2], 8); await app.sleep(250); };
const plain = async (k) => { await app.press(k[0], k[1], k[2]); await app.sleep(250); };

try {
  const dataUrl = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + Buffer.from(docx).toString('base64');
  await app.js(`window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,data:{token:'tok-sel',url:${JSON.stringify(dataUrl)},name:'sel.docx',size:${docx.length},access:'readwrite'}})}`);
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 500 });
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.click('עיון בקבצים…', { after: 8000 });
  for (let waited = 0; waited < 40_000; waited += 250) {
    await app.sleep(250);
    if (!(await app.exists('.status-load'))) break;
  }
  await app.sleep(2500);
  await app.tab('בית');
  console.log(await shapes());

  for (const text of ['אבגדהוז ראשון גדול', 'פרק 12 בספר הזה', 'מספור ראשון ברשימה']) {
    console.log(`\n=== ${text}`);
    await app.caretPara(text);
    // אמצע השורה בחצים רגילים, כדי שיהיה מקום לשני הכיוונים.
    for (let i = 0; i < 5; i++) await plain(L);
    console.log('מוצא          ', await state());
    await shift(L); console.log('Shift+Left    ', await state());
    await shift(L); console.log('Shift+Left    ', await state());
    await shift(R); console.log('Shift+Right   ', await state());
    await shift(R); console.log('Shift+Right   ', await state());
    await shift(R); console.log('Shift+Right   ', await state());
    await plain(L); console.log('Left (קריסה)  ', await state());

    await app.caretPara(text);
    for (let i = 0; i < 5; i++) await plain(L);
    await shift(L); await shift(L); await shift(L); console.log('בחירה         ', await state());
    await plain(R); console.log('Right (קריסה) ', await state());
  }
} finally {
  await app.close();
}
