/** Real SuperDoc regression checks: normal note editing, cancellation on body edits, and stale offsets. */
import assert from 'node:assert/strict';
import { openApp } from './harness.mjs';

const port = Number(process.env.QA_PORT ?? 9818);

function assertRunFormat(xml, text, size, family) {
  const runs = (xml.match(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g) ?? []).map((xml) => ({
    xml,
    text: [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((match) => match[1]).join(''),
  }));
  const plain = runs.map((run) => run.text).join('');
  const start = plain.indexOf(text);
  assert.ok(start >= 0, `missing ${text} in ${plain}`);
  let offset = 0;
  for (const run of runs) {
    const end = offset + run.text.length;
    if (end > start && offset < start + text.length) {
      assert.ok(run.xml.includes(`w:val="${size * 2}"`) && run.xml.includes(`w:ascii="${family}"`), run.xml);
    }
    offset = end;
  }
}

async function scenario(name, run) {
  const app = await openApp({ name: `note-format-${name}`, port, settle: 1500 });
  const ask = async (code) => JSON.parse(await app.js(`(async()=>JSON.stringify(await (${code})))()`));
  const move = async (offset) => {
    await ask(`(async()=>{
      const point={kind:'text',blockId:window.__noteBlock,offset:${offset},story:{kind:'story',storyType:'body'}};
      return window.__qa.sd().activeEditor.authoring.setSelectionTarget({
        target:{kind:'selection',start:point,end:point,story:point.story},focus:true});
    })()`);
    await app.sleep(400);
  };
  const insert = async (type = 'footnote') => {
    await move(3);
    await app.tab('הפניות');
    await app.click(type === 'footnote' ? 'הערת שוליים' : 'הערת סיום', { exact: true });
    await app.sleep(400);
    // Observe actual correction commands without changing their behavior.
    await app.js(`(()=>{
      window.__noteFontCalls=[];
      const commands=window.__qa.ui().commands;
      const execute=commands.executeAsync.bind(commands);
      commands.executeAsync=(id,payload)=>{
        if(id==='font-family'||id==='font-size')window.__noteFontCalls.push([id,payload]);
        return execute(id,payload);
      };
    })()`);
  };
  const calls = () => ask('window.__noteFontCalls');
  try {
    await app.caret(0);
    await app.js(`(async()=>{
      await window.__qa.ui().commands.executeAsync('font-family','TaameyDavidCLM');
      await window.__qa.ui().commands.executeAsync('font-size',16);
    })()`);
    await app.type('abcdefghij', 10);
    await app.sleep(400);
    await ask(`(async()=>{
      window.__noteBlock=(await window.__qa.doc().selection.current()).selectionTarget.start.blockId;
      return true;
    })()`);
    await run({ app, ask, move, insert, calls });
    console.log(`PASS ${name}`);
  } finally {
    app.close();
  }
}

for (const type of ['footnote', 'endnote']) {
  await scenario(`${type}-normal`, async ({ app, move, insert, calls }) => {
    await insert(type);
    await app.type('note content', 10);
    await app.sleep(400);
    await move(4);
    assert.deepEqual(await calls(), [['font-family', 'TaameyDavidCLM'], ['font-size', 16]]);
    await app.type('GOOD', 10);
    await app.sleep(500);
    const xml = (await app.docx())['word/document.xml'];
    assertRunFormat(xml, 'GOOD', 16, 'TaameyDavidCLM');
  });
}

await scenario('body-edit-moves-live-reference', async ({ ask, move, insert, calls }) => {
  await insert();
  await ask(`(async()=>{
    const story={kind:'story',storyType:'body'};
    const point={kind:'text',blockId:window.__noteBlock,offset:0,story};
    return window.__qa.doc().insert({type:'text',value:'x',
      target:{kind:'selection',start:point,end:point,story}});
  })()`);
  const notes = await ask('window.__qa.doc().footnotes.list()');
  assert.equal(notes.total, 1, 'the moved reference must remain live');
  await move(4);
  assert.deepEqual(await calls(), [], 'a live reference shifted by a body edit must cancel the stale offset');
});

await scenario('deleted-before-first-body-return', async ({ ask, move, insert, calls }) => {
  await insert();
  await ask(`(async()=>{
    const page=await window.__qa.doc().footnotes.list();
    return window.__qa.doc().footnotes.remove({target:page.items[0].address});
  })()`);
  await move(4);
  assert.deepEqual(await calls(), [], 'deleted reference must cancel the pending correction');
});

await scenario('return-elsewhere-and-change-font', async ({ app, ask, move, insert, calls }) => {
  await insert();
  await move(0);
  await ask(`(async()=>{
    const page=await window.__qa.doc().footnotes.list();
    return window.__qa.doc().footnotes.remove({target:page.items[0].address});
  })()`);
  await app.press('a', 'KeyA', 65, 2);
  await app.js(`(async()=>{
    await window.__qa.ui().commands.executeAsync('font-family','Arial');
    await window.__qa.ui().commands.executeAsync('font-size',24);
    window.__noteFontCalls=[];
  })()`);
  await move(4);
  assert.deepEqual(await calls(), []);
  await app.type('ZZ', 10);
  await app.sleep(500);
  const xml = (await app.docx())['word/document.xml'];
  assertRunFormat(xml, 'ZZ', 24, 'Arial');
});
