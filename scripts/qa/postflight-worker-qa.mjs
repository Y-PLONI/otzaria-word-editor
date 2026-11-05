/** file:// package equivalence and real typing during a product save. */
import { build } from 'vite';
import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir, cpus, platform, totalmem } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { execFileSync } from 'node:child_process';
import { parseArgs } from 'node:util';
import { zipDeflated, buildDocx, p, stylesXml } from './docx-fixtures.mjs';
import { openApp } from './harness.mjs';
import { openPage, onInterrupt, requireChrome, sleep } from '../cdp.mjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const { values: args } = parseArgs({ options: {
  out: { type: 'string', default: join(tmpdir(), 'otzaria-postflight-worker.json') },
  runs: { type: 'string', default: '3' }, paragraphs: { type: 'string', default: '5000' },
  keys: { type: 'string', default: '12' }, cpu: { type: 'string', default: '1' },
  port: { type: 'string', default: '9792' }, 'package-only': { type: 'boolean' },
} });
function integer(value, min, max) {
  const number = Number(value);
  assert(Number.isSafeInteger(number) && number >= min && number <= max, `Invalid value: ${value}`);
  return number;
}
const runs = integer(args.runs, 1, 20), paragraphs = integer(args.paragraphs, 1, 40000);
const keys = integer(args.keys, 1, 50), cpu = integer(args.cpu, 1, 20), port = integer(args.port, 1024, 65000);
const temporary = mkdtempSync(join(tmpdir(), 'otzaria-postflight-qa-'));
let browser, app, binding;
let stopping = false;
const cleanup = () => { stopping = true; binding?.close(); browser?.close(); app?.close(); rmSync(temporary, { recursive: true, force: true }); };
const forget = onInterrupt(cleanup);

/** Observe only our package workers, and provide a local save destination. */
function instrument(disabled = false) {
  const now = () => performance.timeOrigin + performance.now();
  const trace = window.__postflightTrace = {
    enabled: false, phases: [], uploads: [], longTasks: [],
    workers: { created: 0, transferred: 0, responses: 0, terminated: 0, failed: 0, blocked: 0 },
  };
  const Original = window.Worker;
  window.Worker = new Proxy(Original, { construct(Target, args) {
    if (args[1]?.name !== 'otzaria-docx-postflight') return Reflect.construct(Target, args);
    const previous = trace.phases[trace.phases.length - 1];
    if (trace.enabled && (!previous || previous.endEpochMs !== null)) {
      const phase = { startEpochMs: now(), endEpochMs: null };
      trace.phases.push(phase);
      // Notify Node before processing. The gate verifies that actual native key
      // dispatch times overlap this phase, including a busy renderer's queue.
      window.__otzariaPostflightStage?.(JSON.stringify(phase));
    }
    if (disabled) { trace.workers.blocked++; throw new Error('QA forced local postflight fallback'); }
    const worker = Reflect.construct(Target, args); trace.workers.created++;
    const post = worker.postMessage;
    let request;
    worker.postMessage = function (message, transfer) {
      request = { version: message.version, id: message.id, operation: message.operation };
      const bytes = message.bytes;
      const result = Reflect.apply(post, this, arguments);
      if (transfer?.includes(bytes) && bytes.byteLength === 0) trace.workers.transferred++;
      return result;
    };
    const terminate = worker.terminate;
    worker.terminate = function () { trace.workers.terminated++; return Reflect.apply(terminate, this, arguments); };
    worker.addEventListener('message', ({ data }) => {
      trace.workers.responses++;
      if (data?.ok !== true || !request || data.version !== request.version || data.id !== request.id
        || data.operation !== request.operation || !(data.result?.bytes === null || data.result?.bytes instanceof ArrayBuffer)) {
        trace.workers.failed++;
      }
    });
    worker.addEventListener('error', () => trace.workers.failed++);
    worker.addEventListener('messageerror', () => trace.workers.failed++);
    return worker;
  } });
  if (PerformanceObserver.supportedEntryTypes.includes('longtask')) {
    new PerformanceObserver(list => trace.longTasks.push(...list.getEntries().map(task => ({
      startEpochMs: performance.timeOrigin + task.startTime, durationMs: task.duration,
    })))).observe({ type: 'longtask' });
  }
  const host = window.__qaHost;
  if (!host) return;
  let ticket = 0;
  host.replies['fs.beginBinaryWrite'] = () => {
    const phase = trace.phases[trace.phases.length - 1];
    if (trace.enabled && phase && phase.endEpochMs === null) phase.endEpochMs = now();
    ticket++;
    return Promise.resolve({ success: true, data: { writeToken: `postflight-${ticket}`, uploadUrl: `https://qa.invalid/postflight/${ticket}` }, error: null });
  };
  host.replies['fs.commitUserFileWrite'] = payload => Promise.resolve({ success: true, data: {
    cancelled: false, token: payload.targetToken ?? 'postflight-qa', name: 'postflight.docx',
  }, error: null });
  host.replies['fs.abortBinaryWrite'] = () => Promise.resolve({ success: true, data: {}, error: null });
  const fetch = window.fetch;
  window.fetch = function (input, options) {
    const url = typeof input === 'string' ? input : input?.url;
    if (url?.startsWith('https://qa.invalid/postflight/')) {
      trace.uploads.push(options.body);
      return Promise.resolve(new Response('', { status: 200 }));
    }
    return Reflect.apply(fetch, this, arguments);
  };
}

/** Event-only CDP connection, so renderer polling cannot delay the typing start. */
async function phaseBinding(name) {
  const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
  const target = targets.find(item => item.url?.endsWith(`/__qa-${name}.html`));
  assert(target?.webSocketDebuggerUrl, 'The probe tab must be uniquely identified');
  const socket = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((accept, reject) => {
    socket.addEventListener('open', accept, { once: true });
    socket.addEventListener('error', reject, { once: true });
  });
  let resolveStart, rejectStart, resolveSetup, rejectSetup;
  const started = new Promise((accept, reject) => { resolveStart = accept; rejectStart = reject; });
  const setup = new Promise((accept, reject) => { resolveSetup = accept; rejectSetup = reject; });
  const timer = setTimeout(() => rejectStart(new Error('No postflight worker stage was reached')), 30000);
  socket.addEventListener('message', event => {
    const message = JSON.parse(event.data);
    if (message.id === 2) message.error ? rejectSetup(new Error(JSON.stringify(message.error))) : resolveSetup();
    if (message.method === 'Runtime.bindingCalled' && message.params.name === '__otzariaPostflightStage') {
      clearTimeout(timer); resolveStart(JSON.parse(message.params.payload));
    }
  });
  socket.send(JSON.stringify({ id: 1, method: 'Runtime.enable' }));
  socket.send(JSON.stringify({ id: 2, method: 'Runtime.addBinding', params: { name: '__otzariaPostflightStage' } }));
  try { await setup; }
  catch (error) { clearTimeout(timer); socket.close(); throw error; }
  return { started, close() { clearTimeout(timer); socket.close(); } };
}

async function typeBurst() {
  const pending = [], samples = [];
  for (let index = 0; index < keys && !stopping; index++) {
    const sentEpochMs = Date.now();
    const sample = { sentEpochMs, commandRoundTripMs: null }; samples.push(sample);
    pending.push(Promise.all([
      app.cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'x', code: 'KeyX', text: 'x', unmodifiedText: 'x', windowsVirtualKeyCode: 88, nativeVirtualKeyCode: 88, timestamp: sentEpochMs / 1000 }),
      app.cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'x', code: 'KeyX', windowsVirtualKeyCode: 88, nativeVirtualKeyCode: 88 }),
    ]).then(replies => {
      sample.commandRoundTripMs = Date.now() - sentEpochMs;
      sample.error = replies.find(reply => reply.error)?.error ?? null;
    }));
    await sleep(25);
  }
  await Promise.all(pending);
  assert(samples.every(sample => !sample.error), 'Every native key command must succeed');
  return samples;
}

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const body = count => p('בדיקת שמירה', '<w:bidi/>') + Array.from({ length: count }, (_, i) =>
  `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>פסקה ${i} עם טקסט mixed 123.</w:t></w:r></w:p>`).join('');
const fixtures = {
  unchanged: zipDeflated({ 'word/document.xml': `<w:document xmlns:w="${ns}"><w:body>${p('plain text')}</w:body></w:document>` }),
  large: zipDeflated({
    'word/document.xml': `<w:document xmlns:w="${ns}"><w:body>${body(paragraphs)}</w:body></w:document>`,
    'word/styles.xml': stylesXml(), 'word/media/unchanged.bin': new Uint8Array([3, 5, 7, 11]),
  }),
  corrupt: Buffer.from('invalid ZIP'),
  officeDocm: readFileSync(join(root, 'tests/fixtures/preflight/office-macros.docm')),
};
const report = {
  revision: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim(),
  sourceDirty: !!execFileSync('git', ['status', '--porcelain'], { cwd: root, encoding: 'utf8' }).trim(),
  environment: { platform: platform(), cpu: cpus()[0]?.model, logicalCpus: cpus().length, ramBytes: totalmem(), rendererCpuRate: cpu },
  note: 'Headless Chromium, file://, synthetic typing document. Main fallback and worker use this same build. Dispatch-to-handler measures input queue delay; two-rAF after keydown is only a scheduling opportunity, not per-key engine acceptance or native paint. Saved snapshot, final canonical and painted text, body caret, active document and dirty state are validated separately. No user document text is recorded.',
  paragraphs, typingParagraphs: Math.min(paragraphs, 1000) + 1, package: [], typing: [],
  typingPolicy: { autosaveEnabled: false, explicitSaves: 2, recoveryDraftsEnabled: true },
};

try {
  requireChrome();
  await build({ configFile: false, root, logLevel: 'warn', worker: { format: 'iife' }, build: {
    outDir: temporary, emptyOutDir: false, target: 'es2020',
    rollupOptions: { input: join(root, 'scripts/qa/postflight-worker-browser.js'), output: { format: 'iife', inlineDynamicImports: true, entryFileNames: 'probe.js' } },
  } });
  const entry = join(temporary, 'index.html');
  writeFileSync(entry, `<meta charset="utf-8"><script>(${instrument.toString()})();</script><script src="./probe.js"></script>`);
  browser = await openPage(pathToFileURL(entry).href, { port, label: 'postflight-package' });
  assert.equal((await browser.cdp.send('Emulation.setCPUThrottlingRate', { rate: cpu })).error, undefined, 'Package CPU throttling must be accepted');
  report.environment.userAgent = await browser.cdp.evaluate('navigator.userAgent');
  for (const [name, bytes] of Object.entries(fixtures)) {
    for (let run = 1; run <= runs; run++) {
      const samples = {};
      for (const mode of run % 2 ? ['main', 'worker'] : ['worker', 'main']) {
        samples[mode] = await browser.cdp.evaluate(`window.__postflightQA.measure(${JSON.stringify(bytes.toString('base64'))}, ${JSON.stringify(mode)})`);
      }
      for (const key of ['sha256', 'unchanged', 'outputBytes']) assert.deepEqual(samples.worker[key], samples.main[key], `${name}: ${key}`);
      for (const key of ['created', 'transferred', 'responses', 'terminated']) assert.equal(samples.worker.workers[key], 1, `${name}: ${key}`);
      assert.equal(samples.worker.workers.failed, 0);
      report.package.push({ name, run, inputBytes: bytes.length, ...samples });
    }
  }
  browser.close(); browser = null;

  if (!args['package-only']) {
    // Keep engine pagination bounded; package-only comparisons above use the full size.
    const docx = buildDocx({ body: body(Math.min(paragraphs, 1000)) });
    for (let run = 1; run <= runs; run++) for (const disabled of run % 2 ? [true, false] : [false, true]) {
      const mode = disabled ? 'main-fallback' : 'worker';
      const name = `postflight-typing-${mode}-${run}`;
      app = await openApp({ name, port, extra: `<script>(${instrument.toString()})(${disabled});</script><script src="${pathToFileURL(join(temporary, 'probe.js')).href}"></script>` });
      await app.js(`window.__qaHost.emit('plugin.page_opened', ${JSON.stringify({ param: {
        name: 'postflight.docx', dataUrl: 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + docx.toString('base64'),
      } })})`);
      let loaded = false;
      for (let attempt = 0; attempt < 200; attempt++) {
        loaded = await app.js(`document.querySelector('.superdoc-page')?.textContent?.includes('בדיקת שמירה') && !document.querySelector('.status-load__skip')`);
        if (loaded) break;
        await sleep(100);
      }
      assert(loaded, 'The requested typing fixture must finish opening');
      await app.sleep(1000);
      assert.equal((await app.cdp.send('Emulation.setCPUThrottlingRate', { rate: cpu })).error, undefined, 'Typing CPU throttling must be accepted');
      await app.js('window.__postflightQA.prepareTyping()');
      binding = await phaseBinding(name);
      await app.js('window.__postflightQA.startSave()');
      await binding.started;
      const dispatched = await typeBurst();
      const result = await app.js(`window.__postflightQA.finishTyping(${keys})`);
      binding.close(); binding = null;
      assert.equal(result.outcome.status, 'saved');
      assert.equal(result.packagePhases, 1); assert(result.phase?.endEpochMs > result.phase.startEpochMs);
      for (const key of ['activeDocumentMatches', 'targetMatches', 'typedTextMatches', 'paintedTextMatches', 'uploadedOldSnapshot', 'dirtyAfterOldCommit', 'nextSnapshotMatches']) assert.equal(result[key], true, `${mode}: ${key}`);
      assert.equal(result.nextOutcome.status, 'saved'); assert.equal(result.dirtyAfterNextCommit, false);
      // A recovered worker failure must not be reported as a successful worker
      // timing sample. The fallback comparison is deliberately labelled instead.
      for (const key of ['created', 'transferred', 'responses', 'terminated']) {
        assert.equal(result.workers[key], disabled ? 0 : 1, `${mode}: ${key}`);
      }
      // Vite may retry a rejected Blob worker with a data URL before returning
      // control to our one local fallback. Both attempts remain in one phase.
      if (disabled) assert(result.workers.blocked >= 1, `${mode}: forced fallback`);
      else assert.equal(result.workers.blocked, 0, `${mode}: forced fallback`);
      assert.equal(result.workers.failed, 0, `${mode}: unexpected worker failure`);
      assert.equal(result.keys.length, keys); assert(result.keys.every(key => key.trusted && key.targetInsideEditor && key.sameDocument));
      const measured = result.keys.map((key, index) => {
        const sent = dispatched[index];
        assert(Math.abs(key.eventEpochMs - sent.sentEpochMs) < 20, 'Browser event timestamps must match the supplied CDP timestamps');
        return {
          dispatchedDuringPostflight: sent.sentEpochMs >= result.phase.startEpochMs && sent.sentEpochMs <= result.phase.endEpochMs,
          handledDuringPostflight: key.receivedEpochMs >= result.phase.startEpochMs && key.receivedEpochMs <= result.phase.endEpochMs,
          dispatchToHandlerMs: key.receivedEpochMs - sent.sentEpochMs,
          dispatchToTwoFramesOpportunityMs: key.twoFramesEpochMs - sent.sentEpochMs,
          commandRoundTripMs: sent.commandRoundTripMs,
        };
      });
      assert(measured.some(key => key.dispatchedDuringPostflight), 'At least one native key must be dispatched during postflight');
      assert(measured.every(key => key.dispatchToHandlerMs >= 0 && key.dispatchToTwoFramesOpportunityMs >= key.dispatchToHandlerMs));
      const longTasks = await app.js('window.__postflightTrace.longTasks');
      report.typing.push({ mode, run, fixtureBytes: docx.length, totalSaveMs: result.totalSaveMs,
        postflightMs: result.phase.endEpochMs - result.phase.startEpochMs,
        validatedKeys: keys, keys: measured, workers: result.workers,
        oldSnapshotVerified: true, activeDocumentTextCaretAndPaintVerified: true, dirtyRevisionVerified: true, nextSaveVerified: true,
        longTasks: longTasks.filter(task => task.startEpochMs < result.phase.endEpochMs && task.startEpochMs + task.durationMs > result.phase.startEpochMs)
          .map(task => ({ durationMs: task.durationMs })),
      });
      app.close(); app = null;
    }
  }
} catch (error) {
  report.error = error.stack ?? String(error); process.exitCode = 1;
} finally {
  forget(); cleanup();
  const output = resolve(args.out); mkdirSync(dirname(output), { recursive: true });
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify({ output, packageComparisons: report.package.length, typing: report.typing, error: report.error }, null, 2));
}
