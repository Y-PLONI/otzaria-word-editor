/**
 * רגרסיות בחירה: הכיוון נשמר כשעוברים בין העקיפה ובין המנוע, ואין בליעת
 * מקשים בזמן קריאה א-סינכרונית. התחלות הבחירה והמקשים הם פעולות UI אמיתיות.
 */
import { openApp, createReport } from './harness.mjs';
import { zipStored } from './docx-fixtures.mjs';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const p = (t, pPr = '<w:bidi/>') => `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const body = [
  p('אבגדהוז ראשון גדול'),
  p('תחילת השורה הבאה לבדיקה'),
  p('פרק 12345 בספר'),
  p('מספור ראשון ברשימה', '<w:bidi/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
  `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>קודם</w:t></w:r><w:r><w:tab/></w:r><w:r><w:rPr><w:rtl/></w:rPr><w:t>אחרון</w:t></w:r></w:p>`,
  ...Array.from({ length: 5 }, () => '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    + p('עמוד נוסף לבדיקת ראש שאינו מצויר')),
  p('סוף המסמך לבדיקת בחירה'),
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


const app = await openApp({ name: 'rtl-selection-regressions', port: Number(process.env.QA_PORT ?? 9717) });
const report = createReport('רגרסיות בחירה בחצים', { strict: true });
const selection = () => app.js(`(async () => {
  const value = await window.__otzariaEditor.superdoc.activeEditor.doc.selection.current();
  return value.selectionTarget;
})()`);
const caret = () => app.js(`(() => {
  const ed = window.__otzariaEditor.superdoc.activeEditor;
  const target = ed.host.readLiveSelectionSyncSnapshot()?.selectionTarget;
  const r = document.querySelector('.sd-v2-local-selection-caret')?.getBoundingClientRect();
  return { offset: target?.end?.offset, blockId: target?.end?.blockId,
    x: r?.left, y: r ? (r.top + r.bottom) / 2 : null };
})()`);
const key = async (name, modifiers = 0) => {
  const vk = { Home: 36, End: 35, ArrowLeft: 37, ArrowRight: 39, ArrowUp: 38, ArrowDown: 40 }[name];
  await app.press(name, name, vk, modifiers);
  await app.sleep(180);
};
const at = async (text, offset) => {
  await app.caretPara(text);
  await key('Home');
  for (let i = 0; i < text.length + 2; i += 1) {
    const point = await caret();
    if (point.offset === offset) return point;
    await key('ArrowLeft');
  }
  throw new Error(`לא הגיע להיסט ${offset}: ${text}`);
};
const expectRange = async (label, start, end) => {
  const range = await selection();
  const ok = range?.start?.offset === start && range?.end?.offset === end
    && range.start.blockId === range.end.blockId;
  if (ok) report.pass(label, `${start}..${end}`);
  else report.fail(label, JSON.stringify(range));
};
const first = 'אבגדהוז ראשון גדול';
const second = 'תחילת השורה הבאה לבדיקה';
const digits = 'פרק 12345 בספר';
try {
  const dataUrl = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,'
    + Buffer.from(docx).toString('base64');
  await app.js(`window.__qaHost.replies['fs.pickUserFile'] = () => Promise.resolve({ success: true, error: null,
    data: { token: 'tok-selection-regressions', url: ${JSON.stringify(dataUrl)}, name: 'selection.docx',
      size: ${docx.length}, access: 'readwrite' } })`);
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 500 });
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({block:'center'})");
  await app.click('עיון בקבצים…', { after: 8000 });
  for (let waited = 0; waited < 40000; waited += 250) {
    if (!(await app.exists('.status-load'))) break;
    await app.sleep(250);
  }
  await app.sleep(2000);
  await app.tab('בית');

  await at(first, 5);
  await key('Home', 8);
  await expectRange('Shift+Home מתחיל בטווח 0..5', 0, 5);
  await key('ArrowLeft', 8);
  await expectRange('Shift+שמאל אחרי Home שומר עוגן 5 ומכווץ', 1, 5);

  const anchor = await at(second, 9);
  await key('ArrowRight', 8);
  await key('ArrowRight', 8);
  await expectRange('בחירה אופקית לאחור היא 7..9', 7, 9);
  await key('ArrowUp', 8);
  const vertical = await selection();
  if (vertical?.end?.blockId === anchor.blockId && vertical.end.offset === 9
    && vertical.start.blockId !== anchor.blockId) {
    report.pass('Shift+מעלה ממשיך מהראש ושומר את העוגן המקורי');
  } else report.fail('Shift+מעלה ממשיך מהראש ושומר את העוגן המקורי', JSON.stringify(vertical));

  // התשובה באה מהמנוע עצמו; רק מועד ההחזרה מושהה באופן דטרמיניסטי.
  await at(first, 5);
  await key('Home', 8);
  await app.js(`(() => {
    const api = window.__otzariaEditor.superdoc.activeEditor.doc.selection;
    const original = api.current;
    const gate = new Promise(resolve => { window.__releaseSelectionRead = resolve; });
    api.current = function (...args) { return Promise.resolve(original.apply(this, args)).then(value => gate.then(() => value)); };
    window.__restoreSelectionRead = () => { api.current = original; };
  })()`);
  try {
    await key('ArrowLeft');
    await key('ArrowLeft');
  } finally {
    await app.js('window.__restoreSelectionRead(); window.__releaseSelectionRead();');
  }
  await app.sleep(400);
  await expectRange('שני חצים בזמן קריאה מושהית קורסים ואז מתקדמים', 6, 6);

  for (const [arrow, expected] of [['ArrowRight', 6], ['ArrowLeft', 5]]) {
    await at(digits, 5);
    await key('ArrowRight', 8);
    await expectRange(`בחירה פנימית באי ספרות לפני ${arrow}`, 5, 6);
    await key(arrow);
    await expectRange(`${arrow} קורס לצד החזותי בתוך ספרות`, expected, expected);
  }

  const a = await at(digits, 5);
  await key('ArrowRight');
  const b = await caret();
  for (const [from, to, arrow, expected] of [
    [a, b, 'ArrowLeft', 5], [b, a, 'ArrowRight', 6],
  ]) {
    await app.drag(from.x, from.y, to.x, to.y);
    await expectRange(`גרירה בתוך ספרות לפני ${arrow}`, 5, 6);
    await key(arrow, 8);
    await expectRange(`כיוון גרירה בתוך ספרות שומר עוגן ${expected}`, expected, expected);
  }

  const origin = await at(first, 5);
  await key('Home');
  const home = await caret();
  await at(first, 5);
  await app.cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', button: 'left', buttons: 1,
    clickCount: 1, x: home.x, y: home.y, modifiers: 8 });
  await app.cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', button: 'left', buttons: 0,
    clickCount: 1, x: home.x, y: home.y, modifiers: 8 });
  await app.sleep(300);
  await expectRange('Shift+click יוצר בחירה לאחור', 0, origin.offset);
  await key('ArrowLeft', 8);
  await expectRange('Shift+חץ אחרי Shift+click שומר עוגן', 1, origin.offset);

  for (const [zoom, shift, expected] of [[125, true, 7], [75, false, 5]]) {
    await at(first, 5);
    await key('ArrowLeft', 8);
    await app.js(`window.__otzariaEditor.ui.zoom.set(${zoom})`);
    await app.sleep(500);
    await key(shift ? 'ArrowLeft' : 'ArrowRight', shift ? 8 : 0);
    await expectRange(`בחירה אחרי זום ${zoom}% ${shift ? 'ממשיכה תו אחד' : 'קורסת לקצה הנכון'}`, 5, expected);
    await app.js('window.__otzariaEditor.ui.zoom.set(100)');
    await app.sleep(500);
  }

  await at(first, 5);
  await app.press('a', 'KeyA', 65, 2);
  await app.sleep(400);
  const all = await selection();
  const headMounted = await app.js(`Array.from(document.querySelectorAll('[data-source-node-id]'))
    .some(el => el.getAttribute('data-source-node-id') === ${JSON.stringify(all?.end?.blockId)})`);
  if (all?.start?.blockId !== all?.end?.blockId) report.pass('Ctrl+A בוחר מסמך רב־עמודי', `ראש מצויר: ${headMounted}`);
  else report.fail('Ctrl+A בוחר מסמך רב־עמודי', JSON.stringify({ headMounted, all }));
  await key('ArrowRight', 8);
  const narrowed = await selection();
  if (narrowed?.start?.blockId === all.start.blockId && narrowed.start.offset === all.start.offset
    && narrowed?.end?.blockId === all.end.blockId && narrowed.end.offset === all.end.offset - 1) {
    report.pass('Shift+ימין אחרי Ctrl+A מצמצם תו אחד');
  } else report.fail('Shift+ימין אחרי Ctrl+A מצמצם תו אחד', JSON.stringify({ all, narrowed }));
} catch (error) {
  report.fail('הריצה', String(error?.stack ?? error));
} finally {
  report.print();
  app.close();
}
