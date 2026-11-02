/** SuperDoc2.23: second link and its text form one Undo/Redo transaction. */
import assert from 'node:assert/strict';
import {buildSync} from 'esbuild';
import {openApp} from './harness.mjs';
const bundled=buildSync({entryPoints:['src/engine/hyperlinks-manage.ts'],bundle:true,platform:'browser',format:'iife',globalName:'__links',write:false}).outputFiles[0].text;
const app=await openApp({name:'hyperlink-transaction',port:Number(process.env.QA_PORT??9736),extra:`<script>${bundled.replaceAll('</script','<\\/script')}</script>`});
try {
 const ask=async code=>JSON.parse(await app.js(`(async()=>JSON.stringify(await (${code})))()`));
 const d='window.__qa.doc()';
 const xml=async()=>(await app.docx())['word/document.xml'];
 const history=()=>ask(`${d}.history.get()`);
 await app.caretPara(0); await app.type('before ');
 const target=await ask(`${d}.selection.current().then(s=>s.selectionTarget)`);
 const at=offset=>({...target,start:{...target.start,offset},end:{...target.end,offset}});
 const link=(text,offset)=>ask(`__links.insertHyperlinkAtCaret(window.__qa.sd(),${JSON.stringify({href:'https://example.org/'+text,text,at:at(offset)})})`);
 assert.deepEqual(await link('first',7),{ok:true});
 await ask(`${d}.insert({value:' gap ',type:'text',target:${JSON.stringify(at(12))}})`);
 const before=await xml(), depth=(await history()).undoDepth;
 assert.deepEqual(await link('second',17),{ok:true});
 assert.equal((await history()).undoDepth,depth+1);
 assert.equal(((await xml()).match(/<w:hyperlink /g)||[]).length,2);
 await ask(`${d}.history.undo()`); assert.equal(await xml(),before);
 await ask(`${d}.history.redo()`); assert.equal(((await xml()).match(/<w:hyperlink /g)||[]).length,2);
 console.log('PASS second-link single Undo/Redo');
 // Existing link nesting must fail without text or history changes.
 const nestedBefore=await xml(), nestedDepth=(await history()).undoDepth;
 const nested=await link('nested',9);

 assert.equal(nested.ok,false);
 assert.equal(await xml(),nestedBefore); assert.equal((await history()).undoDepth,nestedDepth);
 console.log('PASS nested rejection atomic');
 // Direct authored formatting must survive the inline insertion.
 await ask(`${d}.format.apply({target:${JSON.stringify({...at(17),end:{...target.end,offset:23}})},inline:{bold:true,fontSize:24,fontFamily:'Arial'}})`);
 await ask(`${d}.insert({type:'text',value:' plain ',target:${JSON.stringify(at(23))}})`);
 await ask(`${d}.format.apply({target:${JSON.stringify({...at(23),end:{...target.end,offset:30}})},inline:{bold:true,fontSize:24,fontFamily:'Arial'}})`);
 assert.deepEqual(await link('formatted',26),{ok:true});
 const formattedXml=await xml();
 const formattedLink=formattedXml.slice(0,formattedXml.indexOf('formatted</w:t>')).split('<w:hyperlink ').at(-1);
 assert.match(formattedLink,/<w:b\b/); assert.match(formattedLink,/<w:sz w:val="48"/); assert.match(formattedLink,/w:ascii="Arial"/);
 console.log('PASS authored run formatting');
 // Suggesting mode must author a tracked insertion in a single history entry.
 await app.js(`window.__qa.ui().commands.executeAsync('document-mode',{mode:'suggesting'})`);
 const beforeTracked=(await history()).undoDepth;
 assert.deepEqual(await link('tracked',39),{ok:true});
 assert.equal((await history()).undoDepth,beforeTracked+1);
 assert.match(await xml(),/<w:ins\b[^>]*>[\s\S]*?tracked/);
 console.log('PASS tracked insertion');
 await app.js(`window.__qa.ui().commands.executeAsync('document-mode',{mode:'editing'})`);
 await ask(`${d}.insert({type:'text',value:'',target:${JSON.stringify({...at(7),end:{...target.end,offset:12}})}})`);
 const beforeBlank=(await history()).undoDepth;

 assert.deepEqual(await link('blankreplacement',7),{ok:true});
 assert.equal((await history()).undoDepth,beforeBlank+1);
 console.log('PASS empty hyperlink remnant');
 // Keep the application's supported custom and mail links, not only HTTPS.
 for (const href of ['mailto:reader@example.org']) {
   const end=await ask(`${d}.getText({}).then(t=>t.length)`);
   assert.deepEqual(await ask(`__links.insertHyperlinkAtCaret(window.__qa.sd(),${JSON.stringify({href,text:'scheme',at:at(end)})})`),{ok:true});
 }
 assert.equal(await ask(`__links.insertHyperlinkAtCaret(window.__qa.sd(),${JSON.stringify({href:'otzaria://open/book/42?index=7',text:'custom',at:at(0)})})`),null);
 console.log('PASS mailto and native custom-scheme fallback');
} finally {app.close();}
