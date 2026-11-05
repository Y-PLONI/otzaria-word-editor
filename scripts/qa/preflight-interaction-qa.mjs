/** A controlled opening must never discard keystrokes accepted by the old editor. */
import { readFileSync, writeFileSync, rmSync, mkdirSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
import { buildDocx, p } from './docx-fixtures.mjs';
import { unzip } from './harness.mjs';
import { openPage, onInterrupt, sleep } from '../cdp.mjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const { values: args } = parseArgs({ options: {
  dist: { type: 'string', default: join(root, 'dist') }, 'baseline-dist': { type: 'string' },
  out: { type: 'string', default: '/tmp/preflight-interaction.json' },
  port: { type: 'string', default: '9795' }, observe: { type: 'boolean' },
  // Diagnostic only: reproduce global shortcuts with an inert old container
  // on a pre-fix build, without claiming the product implements that guard.
  'probe-inert': { type: 'boolean' },
} });
assert(!args['probe-inert'] || args.observe, '--probe-inert requires --observe');
const port = Number(args.port);
assert(Number.isInteger(port) && port > 1024 && port < 65000);
const marker = 'מסמך יעד חדש';
const fixture = buildDocx({ body: p(marker) });
const token = 'LostQz', afterToken = 'KeepQz';
const host = readFileSync(join(root, 'scripts/qa/host-stub.js'), 'utf8');
const qa = readFileSync(join(root, 'scripts/qa/qa-api.js'), 'utf8');
const results = [];
let browser = null, probe = null;
function cleanup() { browser?.close(); browser = null; if (probe) rmSync(probe, { force: true }); probe = null; }
const forget = onInterrupt(cleanup);

function barrier(stage, size) {
  const gate = window.__openingBarrier = { stage, armed: false, pending: false, used: false, terminated: false, release: null };
  const originalRead = Blob.prototype.arrayBuffer;
  Blob.prototype.arrayBuffer = function () {
    const read = originalRead.apply(this, arguments);
    if (stage !== 'blob' || !gate.armed || gate.used || this.size !== size) return read;
    gate.used = true;
    return read.then(bytes => new Promise(resolve => {
      gate.pending = true; gate.release = () => { gate.pending = false; resolve(bytes); };
    }));
  };
  const Original = window.Worker;
  window.Worker = new Proxy(Original, { construct(Target, args) {
    const worker = Reflect.construct(Target, args);
    if (stage !== 'worker' || !args[1]?.name?.startsWith('otzaria-docx-')) return worker;
    const post = worker.postMessage, terminate = worker.terminate;
    worker.postMessage = function (request, transfer) {
      if (!gate.armed || gate.used) return post.call(this, request, transfer);
      gate.used = true; gate.pending = true;
      gate.release = () => { gate.pending = false; if (!gate.terminated) post.call(worker, request, transfer); gate.release = null; };
    };
    worker.terminate = function () { gate.terminated = true; gate.pending = false; gate.release = null; return terminate.call(this); };
    return worker;
  } });
}

async function until(expression, timeout = 60000) {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) { if (await browser.cdp.evaluate(expression)) return; await sleep(50); }
  throw new Error(`Timed out: ${expression}`);
}
async function click(rect) {
  assert(rect, 'A visible click target exists');
  for (const type of ['mousePressed', 'mouseReleased']) await browser.cdp.send('Input.dispatchMouseEvent', {
    type, x: rect.x, y: rect.y, button: 'left', buttons: type === 'mousePressed' ? 1 : 0, clickCount: 1,
  });
}
async function type(text) {
  for (const key of text) {
    await browser.cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key, text: key });
    await browser.cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key });
    await sleep(30);
  }
  await sleep(100);
}
async function chord(key, code, windowsVirtualKeyCode) {
  for (const type of ['keyDown', 'keyUp']) await browser.cdp.send('Input.dispatchKeyEvent', {
    type, key, code, windowsVirtualKeyCode, modifiers: 2,
  });
  await sleep(300);
}
async function pageBreaks() {
  const base64 = await browser.cdp.evaluate('window.__qa.exportBase64()');
  assert(base64, 'The old document can be exported');
  const xml = unzip(Buffer.from(base64, 'base64'))['word/document.xml'];
  assert(xml, 'The exported old document contains its body XML');
  return (xml.match(/<w:pageBreakBefore\b[^>]*>/g) ?? []).filter(tag => {
    const value = tag.match(/\bw:val\s*=\s*(['"])(.*?)\1/)?.[2]?.toLowerCase();
    return !['0', 'false', 'off'].includes(value ?? 'true');
  }).length;
}
const nodeText = `node.node.paragraph.inlines.map(inline => inline.run?.text ?? '').join('')`;

async function scenario(dist, stage, cancel, baseline) {
  const label = `${baseline ? 'baseline' : 'candidate'}-${cancel ? 'cancel' : 'finish'}`;
  const appSha256 = createHash('sha256').update(readFileSync(join(dist, 'assets/app.js'))).digest('hex');
  const html = readFileSync(join(dist, 'index.html'), 'utf8');
  const cut = html.indexOf('</script>') + '</script>'.length;
  assert(cut > '</script>'.length, 'Boot latch is present');
  probe = join(dist, `__qa-preflight-interaction-${process.pid}-${label}.html`);
  const inject = `<script>${host}</script><script>${qa}</script><script>(${barrier.toString()})(${JSON.stringify(stage)}, ${fixture.length});</script>`;
  writeFileSync(probe, html.slice(0, cut) + inject + html.slice(cut));
  browser = await openPage(pathToFileURL(probe).href, { port, label: `preflight-${label}` });
  await until('window.__qa?.ready() && !document.querySelector(".status-load__skip") && window.__qa.lineRect(0)');
  await click(await browser.cdp.evaluate('window.__qa.lineRect(0)'));
  await until(`window.__qa.caretBlock(2000).then(caret => caret.answered && caret.blockId && window.__otzariaEditor.container.contains(document.activeElement))`);
  const original = await browser.cdp.evaluate(`(async () => {
    const editor = window.__otzariaEditor, doc = window.__qa.doc();
    const selection = await doc.selection.current(), blockId = selection.target?.segments?.[0]?.blockId;
    if (!selection.empty || !blockId || !editor.container.contains(document.activeElement)) throw new Error('The old blank editor must have the input focus');
    const owner = [...window.__otzariaEditors.values()].find(session => session.swap.current === editor);
    if (!owner || owner.save.snapshot.isDirty || owner.save.snapshot.targetToken !== null) throw new Error('Expected a clean reusable placeholder');
    const node = await doc.getNodeById({ nodeId: blockId });
    window.__oldOpeningEditor = { editor, doc, blockId };
    return { text: ${nodeText}, tabs: document.querySelectorAll('.word-doctab').length };
  })()`);
  assert.equal(original.tabs, 1);
  await browser.cdp.evaluate(`window.__openingBarrier.armed = true; window.__qaHost.emit('plugin.page_opened', { param: { name: 'target.docx', dataUrl: ${JSON.stringify(`data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${fixture.toString('base64')}`)} } });`);
  await until('window.__openingBarrier.pending', 10000);
  if (args['probe-inert']) await browser.cdp.evaluate(`(() => {
    const container = window.__oldOpeningEditor.editor.container;
    window.__probePreviousInert = container.inert;
    container.inert = true;
  })()`);
  await browser.cdp.evaluate('new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))');
  await type(token);
  const pending = await browser.cdp.evaluate(`(async () => {
    const old = window.__oldOpeningEditor, node = await old.doc.getNodeById({ nodeId: old.blockId });
    return { text: ${nodeText}, inert: old.editor.container.inert, oldFocused: old.editor.container.contains(document.activeElement) };
  })()`);
  const shortcutBefore = { bold: await browser.cdp.evaluate('window.__qa.cmd("bold").active'), pageBreaks: await pageBreaks() };
  assert.equal(shortcutBefore.pageBreaks, 0, 'The clean placeholder initially has no enabled page break');
  await chord('b', 'KeyB', 66);
  await chord('Enter', 'Enter', 13);
  const shortcutAfter = { bold: await browser.cdp.evaluate('window.__qa.cmd("bold").active'), pageBreaks: await pageBreaks() };
  const shortcutsChanged = JSON.stringify(shortcutBefore) !== JSON.stringify(shortcutAfter);
  if (cancel) {
    await click(await browser.cdp.evaluate('window.__qa.rectOf(document.querySelector(".status-load__skip"))'));
    await until('!document.querySelector(".status-load__skip") && window.__otzariaEditor === window.__oldOpeningEditor.editor');
    // Blob reads are not abortable. Let the baseline's already-cancelled read
    // finish too, so the observation includes a late result rather than a hang.
    await browser.cdp.evaluate('window.__openingBarrier.release?.()');
    await sleep(250);
  } else {
    await browser.cdp.evaluate('window.__openingBarrier.release()');
    await until(`!document.querySelector('.status-load__skip') && window.__qa.screenText()?.includes(${JSON.stringify(marker)})`);
  }
  if (args['probe-inert']) await browser.cdp.evaluate('window.__oldOpeningEditor.editor.container.inert = window.__probePreviousInert');
  const final = await browser.cdp.evaluate(`(async () => {
    const old = window.__oldOpeningEditor;
    const oldRetained = [...window.__otzariaEditors.values()].some(session => session.swap.current === old.editor);
    const doc = window.__qa.doc(), list = await doc.blocks.list();
    const node = await doc.getNodeById({ nodeId: list.blocks[0].nodeId });
    return { text: ${nodeText}, oldRetained, activeIsOld: window.__otzariaEditor === old.editor,
      inert: window.__otzariaEditor.container.inert,
      focused: window.__otzariaEditor.container.contains(document.activeElement),
      terminated: window.__openingBarrier.terminated, tabs: document.querySelectorAll('.word-doctab').length };
  })()`);
  // Cancellation keeps the old editor; clicking it must work again. The new
  // candidate must receive its automatic opening focus without that click.
  if (cancel) {
    await browser.cdp.evaluate('window.__otzariaEditor.container.querySelector(".superdoc-line")?.scrollIntoView({ block: "center" })');
    await click(await browser.cdp.evaluate('window.__qa.lineRect(0)'));
  } else assert(final.focused, 'The new candidate receives automatic opening focus');
  await type(afterToken);
  const editable = await browser.cdp.evaluate(`(async () => {
    const doc = window.__qa.doc(), list = await doc.blocks.list();
    const node = await doc.getNodeById({ nodeId: list.blocks[0].nodeId });
    return (${nodeText}).includes(${JSON.stringify(afterToken)});
  })()`);
  const acceptedWhileLoading = pending.text !== original.text;
  const lost = acceptedWhileLoading && !final.oldRetained && !final.text.includes(token);
  const result = {
    label, baseline, stage, cancel, appSha256, syntheticInert: !!args['probe-inert'],
    original, pending, shortcutBefore, shortcutAfter, shortcutsChanged, final, editable, acceptedWhileLoading, lost,
  };
  results.push(result);
  assert(editable, 'The surviving/new editor accepts keyboard input');
  if (!baseline && !args.observe) assert(!lost, 'Text accepted by the old editor was discarded during opening');
  if (!baseline && !args.observe) assert(!shortcutsChanged, 'Global shortcuts changed the old editor while opening');
  if (cancel) assert(final.activeIsOld, 'Cancellation retains the original editor');
  cleanup();
}

try {
  if (args['baseline-dist']) for (const cancel of [false, true]) await scenario(resolve(args['baseline-dist']), 'blob', cancel, true);
  for (const cancel of [false, true]) await scenario(resolve(args.dist), 'worker', cancel, false);
} catch (error) {
  results.push({ error: error.stack ?? String(error) }); process.exitCode = 1;
} finally {
  forget(); cleanup();
  mkdirSync(dirname(resolve(args.out)), { recursive: true });
  writeFileSync(resolve(args.out), JSON.stringify(results, null, 2) + '\n');
  console.log(JSON.stringify(results, null, 2));
}
