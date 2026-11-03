/**
 * גשש: „מיזוג עיצוב” על יעד שנראה כמו מסמך Word עברי.
 *
 * כל פסקה במסמך היא מקרה אחד שכיח בקבצים שנוצרו ב-Word: ברירות מחדל עם גופן
 * לטיני שונה מהעברי, ריצה עם `w:rtl`, `w:hint="cs"`, סימן פסקה עם `w:rPr`.
 * לכל אחת: הסמן באמצע הטקסט, הדבקה במיזוג, והגשש מדפיס אם ההדבקה נדחתה ובאיזו
 * סיבה (`console.error` של readDestination), ואם עברה — את ה-rPr של מה שהודבק.
 *
 * הרצה:  node scripts/qa/paste-merge-hebrew-probe.mjs   (QA_PORT דורס 9497)
 */
import { openApp, unzip } from './harness.mjs';
import { buildDocx, stylesXml } from './docx-fixtures.mjs';

const PASTED = 'שלום עולם abc';

const styles = stylesXml().replace(
  '<w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>',
  '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="David"/>' +
    '<w:sz w:val="24"/><w:szCs w:val="28"/><w:lang w:val="en-US" w:bidi="he-IL"/></w:rPr></w:rPrDefault>',
);

const CASES = [
  ['פסקה רגילה של Word: rtl בריצה ובסימן הפסקה', 'אאאאאא אאאאאא',
    '<w:pPr><w:bidi/><w:rPr><w:rtl/></w:rPr></w:pPr>', '<w:rPr><w:rtl/></w:rPr>'],
  ['ריצה עם hint=cs ו-rtl', 'בבבבבב בבבבבב',
    '<w:pPr><w:bidi/></w:pPr>', '<w:rPr><w:rFonts w:hint="cs"/><w:rtl/></w:rPr>'],
  ['ריצה עם rtl בלבד, בלי rPr בסימן הפסקה', 'גגגגגג גגגגגג',
    '<w:pPr><w:bidi/></w:pPr>', '<w:rPr><w:rtl/></w:rPr>'],
  ['בלי rtl — רק ברירות המחדל המפוצלות', 'דדדדדד דדדדדד',
    '<w:pPr><w:bidi/></w:pPr>', ''],
  ['ביקורת: David ישיר בשלושת הערוצים, גודל אחיד', 'הההההה הההההה',
    '<w:pPr><w:bidi/></w:pPr>', '<w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'],
];

const body = CASES.map(([, text, pPr, rPr]) =>
  `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('');

const app = await openApp({ name: 'paste-merge-hebrew', port: Number(process.env.QA_PORT ?? 9497) });

async function openDocx(buffer, name) {
  const dataUrl = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + Buffer.from(buffer).toString('base64');
  await app.js(`window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,data:{token:'tok-${name}',url:${JSON.stringify(dataUrl)},name:${JSON.stringify(name + '.docx')},size:${buffer.length},access:'readwrite'}})}`);
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 2500 });
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.clickSel('.open-browse', 0, { after: 12000 });
  return app.js("document.querySelector('.doc-title-input')?.value");
}

/** לחיצה באמצע השורה שמכילה `needle` — הסמן נוחת בתוך הטקסט. */
async function caretIn(needle) {
  const spot = await app.js(`(function(){var f=document.querySelectorAll('[data-source-node-id][data-pm-start]');for(var i=0;i<f.length;i++){if((f[i].textContent||'').indexOf(${JSON.stringify(needle)})<0)continue;var l=f[i].children[0];if(!l)return '';var r=l.getBoundingClientRect();return JSON.stringify({x:Math.round((r.left+r.right)/2),y:Math.round((r.top+r.bottom)/2)});}return '';})()`);
  if (!spot) return false;
  const { x, y } = JSON.parse(spot);
  await app.clickAt(x, y);
  await app.sleep(800);
  // הלחיצה רק ממקדת; הסמן עצמו מוצב דרך המנוע, בפסקה שהטקסט שלה מתחיל ב-needle.
  const placed = await app.js(`(async()=>{const sd=window.__qa.sd(),d=sd.activeEditor.doc;const l=await d.blocks.list({limit:50});const b=(l.blocks||[]).find(x=>JSON.stringify(x).includes(${JSON.stringify(needle)}));if(!b)return 'no-block:'+JSON.stringify(l.blocks?.[0]).slice(0,300);const t=sd.ui.selection.getSnapshot().selectionTarget;const p={...t.start,blockId:b.nodeId,offset:3};const r=await sd.ui.selection.apply({...t,start:p,end:p});return JSON.stringify(sd.ui.selection.getSnapshot().selectionTarget.start)})()`);
  await app.sleep(600);
  if (process.env.DUMP) console.log('caret', placed);
  return true;
}

async function merge() {
  await app.tab('בית');
  if (!(await app.openMenu('אפשרויות הדבקה'))) throw new Error('תפריט ההדבקה לא נפתח');
  await app.js(`(()=>{const b=[...document.querySelectorAll('.ribbon-menu__item')].find(e=>e.querySelector('.ribbon-menu__item-label')?.textContent.trim()==='מיזוג עיצוב');b?.scrollIntoView({block:'nearest'});})()`);
  await app.sleep(150);
  if (!(await app.clickMenu('מיזוג עיצוב'))) throw new Error('„מיזוג עיצוב” לא נמצא');
  await app.sleep(2500);
}

try {
  await app.cdp.send('Page.bringToFront');
  await app.cdp.send('Browser.grantPermissions', { permissions: ['clipboardReadWrite', 'clipboardSanitizedWrite'] });
  console.log('נפתח:', await openDocx(buildDocx({ body, styles }), 'merge-hebrew'));
  await app.js(`window.__errs=[];(function(){const o=console.error.bind(console);console.error=(...a)=>{window.__errs.push(a.map(String).join(' '));o(...a);};})()`);

  for (const [label, text] of CASES) {
    const marker = text.slice(0, 3);
    await app.js(`window.__errs=[]`);
    await app.js(`navigator.clipboard.writeText(${JSON.stringify(PASTED)})`);
    if (!(await caretIn(marker))) { console.log(`✗ ${label}: השורה לא נמצאה`); continue; }
    if (process.env.DUMP) console.log(await app.js(`(async()=>{const sd=window.__qa.sd(),d=sd.activeEditor.doc;const t=sd.ui.selection.getSnapshot().selectionTarget;const o=t.start.offset;const at={...t,start:{...t.start,offset:Math.max(0,o-1)},end:{...t.start,offset:o>0?o:1}};const c=await d.selection.extractOoxml({at});return JSON.stringify({frag:c.fragment.xml,ctx:c.context.xml,style:c.source.paragraphStyleId,deps:c.dependencies.map(x=>x.target)})})()`));
    const before = await app.js(`(async()=>(await window.__qa.sd().activeEditor.doc.info()).revision)()`);
    await merge();
    const after = await app.js(`(async()=>(await window.__qa.sd().activeEditor.doc.info()).revision)()`);
    const errs = JSON.parse(await app.js(`JSON.stringify(window.__errs)`));
    const reason = errs.find((e) => e.includes('מיזוג עיצוב נדחה'));
    if (before === after) {
      console.log(`נדחה   ${label}\n       ${reason ?? '(אין סיבה רשומה) ' + JSON.stringify(await app.status())}`);
      continue;
    }
    const xml = (await app.docx())['word/document.xml'];
    const paragraph = (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []).find((p) => p.includes(marker));
    const runs = (paragraph?.match(/<w:r>[\s\S]*?<\/w:r>/g) ?? [])
      .filter((r) => /שלום|עולם|abc/.test(r))
      .map((r) => `${(r.match(/<w:t[^>]*>([^<]*)/) ?? [])[1]}  ⇐ ${(r.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/) ?? [])[1] ?? '(אין rPr)'}`);
    console.log(`הודבק  ${label}\n       ${runs.join('\n       ')}`);
  }
} catch (error) {
  console.error('QA ERRORS', await app.log());
  throw error;
} finally { app.close(); }
