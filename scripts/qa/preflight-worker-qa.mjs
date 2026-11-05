/** Real file:// transport, equivalence and separate wall/main-thread measurements. */
import { build } from 'vite';
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync, rmSync, existsSync, statSync } from 'node:fs';
import { join, dirname, resolve } from 'node:path';
import { tmpdir, cpus, platform, totalmem } from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
import { zipDeflated, buildDocx, p } from './docx-fixtures.mjs';
import { openPage, onInterrupt, requireChrome, sleep } from '../cdp.mjs';
import { openApp } from './harness.mjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const { values: args } = parseArgs({ options: {
  out: { type: 'string', default: join(tmpdir(), 'otzaria-preflight-worker.json') },
  runs: { type: 'string', default: '3' }, paragraphs: { type: 'string', default: '5000' },
  cpu: { type: 'string', default: '1' }, port: { type: 'string', default: '9794' },
  'package-only': { type: 'boolean' },
  'baseline-dist': { type: 'string' },
} });
function integer(value, min, max) {
  const number = Number(value);
  assert(Number.isSafeInteger(number) && number >= min && number <= max, `Invalid value ${value}`);
  return number;
}
const runs = integer(args.runs, 1, 20), paragraphs = integer(args.paragraphs, 1, 40000);
const cpu = integer(args.cpu, 1, 20), port = integer(args.port, 1024, 65000);
const temporary = mkdtempSync(join(tmpdir(), 'otzaria-package-qa-'));
let browser, app;
const cleanup = () => { browser?.close(); app?.close(); rmSync(temporary, { force: true, recursive: true }); };
const forget = onInterrupt(cleanup);

/** Runs before product code, observes only our package workers, never engine workers. */
function observeWorkers(disabled = false) {
  const stats = window.__packageWorkers = { created: 0, transferred: 0, responses: 0, terminated: 0, failed: 0, blocked: 0, blobWorkers: 0, moduleWorkers: 0, sourceBytes: 0, constructorMs: 0, firstResponseMs: 0 };
  const sizes = new Map(), createURL = URL.createObjectURL;
  URL.createObjectURL = function (blob) { const url = Reflect.apply(createURL, this, arguments); sizes.set(url, blob.size); return url; };
  const Original = window.Worker;
  window.Worker = new Proxy(Original, { construct(Target, args) {
    if (!args[1]?.name?.startsWith('otzaria-docx-')) return Reflect.construct(Target, args);
    if (disabled) { stats.blocked++; throw new Error('QA forced local package processing'); }
    const start = performance.now();
    const worker = Reflect.construct(Target, args); stats.created++;
    if (String(args[0]).startsWith('blob:')) stats.blobWorkers++;
    if (args[1]?.type === 'module') stats.moduleWorkers++;
    stats.constructorMs += performance.now() - start;
    stats.sourceBytes += sizes.get(args[0]) ?? 0;
    const post = worker.postMessage;
    worker.postMessage = function (message, transfer) {
      const buffer = message.bytes;
      const result = Reflect.apply(post, this, arguments);
      if (transfer?.includes(buffer) && buffer.byteLength === 0) stats.transferred++;
      return result;
    };
    const terminate = worker.terminate;
    worker.terminate = function () { stats.terminated++; return Reflect.apply(terminate, this, arguments); };
    worker.addEventListener('message', event => { stats.responses++; if (event.data?.ok !== true) stats.failed++; });
    worker.addEventListener('message', () => { stats.firstResponseMs += performance.now() - start; }, { once: true });
    worker.addEventListener('error', () => stats.failed++);
    return worker;
  } });
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const settings = zero => `<w:settings ${W}><w:defaultTabStop w:val="${zero ? 0 : 720}"/></w:settings>`;
const fontTable = `<w:fonts ${W}><w:font w:name="David"/></w:fonts>`;
const unchanged = zipDeflated({ 'word/settings.xml': settings(false), 'word/fontTable.xml': fontTable });
const large = zipDeflated({
  'word/settings.xml': settings(true), 'word/fontTable.xml': fontTable,
  'word/document.xml': `<w:document ${W}><w:body>${Array.from({ length: paragraphs }, (_, i) =>
    `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:bCs/></w:rPr><w:t>בדיקה ${i}</w:t></w:r><w:r><w:rPr><w:bCs/></w:rPr><w:t> המשך</w:t></w:r></w:p>`).join('')}</w:body></w:document>`,
  'word/media/untouched.bin': new Uint8Array([3, 7, 11, 19]),
});
const macro = zipDeflated({
  'word/settings.xml': settings(true), 'word/fontTable.xml': fontTable,
  '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
  'word/vbaProject.bin': new Uint8Array(600),
});
const fixtures = {
  empty: Buffer.alloc(0), unchanged, large, corrupt: Buffer.from('invalid ZIP'), macro,
  officeMacro: readFileSync(join(root, 'tests/fixtures/preflight/office-macros.docm')),
  officeMacroRepair: readFileSync(join(root, 'tests/fixtures/preflight/office-macros-repair.docm')),
  autoRunMacro: readFileSync(join(root, 'tests/fixtures/preflight/auto-open-macros.docm')),
};
const artifact = path => existsSync(path)
  ? { bytes: statSync(path).size, sha256: createHash('sha256').update(readFileSync(path)).digest('hex') } : null;
const appArtifact = artifact(join(root, 'dist/assets/app.js'));
const baselineArtifact = args['baseline-dist'] ? artifact(join(resolve(args['baseline-dist']), 'assets/app.js')) : null;
const report = {
  revision: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim(),
  sourceDirty: !!execFileSync('git', ['status', '--porcelain'], { cwd: root, encoding: 'utf8' }).trim(),
  environment: { platform: platform(), cpu: cpus()[0]?.model, logicalCpus: cpus().length, ramBytes: totalmem(), rendererCpuRate: cpu },
  note: 'Headless Chromium; synthetic packages plus an Office-produced VBA project. Package wall time is separate from main-thread long tasks. App baseline is the same build with its local fallback, not a native release comparison.',
  paragraphs, package: [], cancellation: null, opening: [],
  artifacts: { app: appArtifact, baseline: baselineArtifact, appDeltaBytes: appArtifact && baselineArtifact ? appArtifact.bytes - baselineArtifact.bytes : null },
};

try {
  requireChrome();
  await build({ configFile: false, root, logLevel: 'warn', worker: { format: 'iife' }, build: {
    outDir: temporary, emptyOutDir: false, target: 'es2020',
    rollupOptions: { input: join(root, 'scripts/qa/preflight-worker-browser.js'), output: {
      format: 'iife', inlineDynamicImports: true, entryFileNames: 'probe.js',
    } },
  } });
  const page = join(temporary, 'index.html');
  writeFileSync(page, `<meta charset="utf-8"><script>(${observeWorkers.toString()})();</script><script src="./probe.js"></script>`);
  browser = await openPage(pathToFileURL(page).href, { port, label: 'preflight-package' });
  await browser.cdp.send('Emulation.setCPUThrottlingRate', { rate: cpu });
  report.environment.userAgent = await browser.cdp.evaluate('navigator.userAgent');
  for (const [name, bytes] of Object.entries(fixtures)) {
    for (let run = 1; run <= runs; run++) {
      const samples = {};
      // Alternate order to reduce warm-cache/order bias; no performance thresholds.
      for (const mode of run % 2 ? ['main', 'worker'] : ['worker', 'main']) {
        samples[mode] = await browser.cdp.evaluate(`window.__preflightQA.measure(${JSON.stringify(bytes.toString('base64'))}, ${JSON.stringify(mode)})`);
      }
      for (const key of ['sha256', 'unchanged', 'fontTable', 'vba', 'notice', 'vbaBytesPreserved']) assert.deepEqual(samples.worker[key], samples.main[key], `${name}: ${key}`);
      if (name === 'officeMacro' || name === 'officeMacroRepair') {
        assert(samples.worker.vba.modules.length > 0, 'Actual Office modules were decoded');
        assert(samples.worker.vba.modules.some(module => module.source.includes('say_hello')), 'Actual source text survived data cloning');
        assert.equal(samples.worker.vba.unreadable, false);
      }
      if (name === 'officeMacroRepair' || name === 'autoRunMacro') {
        assert.equal(samples.worker.unchanged, false, `${name}: XML repair was exercised`);
        assert(samples.worker.notice, `${name}: repair notice survived data cloning`);
      }
      if (name === 'autoRunMacro') {
        assert.deepEqual(samples.worker.vba.autoRun, ['Module1.AutoOpen']);
        assert(samples.worker.vba.warnings.length > 0, 'Auto-run warnings survived data cloning');
        assert(samples.worker.vba.modules.some(module => module.source.includes('בדיקת תעבורה בלבד')), 'Windows-1255 source text survived data cloning');
        assert.equal(samples.worker.vba.unreadable, false);
      }
      if (['macro', 'officeMacro', 'officeMacroRepair', 'autoRunMacro'].includes(name)) {
        assert.equal(samples.worker.vbaBytesPreserved, true, `${name}: original VBA bytes are unchanged`);
      }
      assert.equal(samples.worker.worker.transferred, 1, `${name}: transferable input`);
      assert.equal(samples.worker.worker.responses, 1, `${name}: real worker result`);
      assert.equal(samples.worker.worker.terminated, 1, `${name}: no live worker`);
      assert.equal(samples.worker.worker.failed, 0, `${name}: no local fallback`);
      assert.equal(samples.worker.worker.blobWorkers, 1, `${name}: Blob URL`);
      assert.equal(samples.worker.worker.moduleWorkers, 0, `${name}: classic worker`);
      assert(samples.worker.worker.sourceBytes > 0, `${name}: measured worker payload size`);
      report.package.push({ name, run, inputBytes: bytes.length, ...samples });
    }
  }
  report.cancellation = await browser.cdp.evaluate(`window.__preflightQA.cancel(${JSON.stringify(large.toString('base64'))})`);
  assert.deepEqual(report.cancellation, { cancelled: true, detached: true, terminated: 1, created: 1 });
  browser.close(); browser = null;

  if (!args['package-only']) {
    // Modest real DOCX for the engine gate. The larger fixture above measures only
    // package work, so an engine's page count does not dominate that comparison.
    const text = 'בדיקת פתיחה של מסמך';
    const docx = buildDocx({ body: Array.from({ length: 100 }, (_, i) => p(`${text} ${i}`)).join('') });
    for (const disabled of [true, false]) {
      const setup = `<script>(${observeWorkers.toString()})(${disabled});
        const bytes = Uint8Array.from(atob(${JSON.stringify(docx.toString('base64'))}), c => c.charCodeAt(0));
        const url = URL.createObjectURL(new Blob([bytes]));
        const H = window.__qaHost;
        H.replies['fs.resolveFileUrl'] = { url, name: 'package-qa.docx', size: bytes.length };
        H.storage.session = { version: 2, activeId: 'package-qa', documents: [{ id: 'package-qa', document: { token: 'package-qa', name: 'package-qa.docx', writable: false }, draft: null, previousDraft: null, caret: null }], view: { zoom: 100, focusMode: false, ribbonTab: 'home', ribbonCollapsed: false } };
        const fetchSource = window.fetch;
        const timing = window.__packageOpening = { start: null, settled: null, tasks: [] };
        window.fetch = function (input) { if (input === url) timing.start = performance.now(); return Reflect.apply(fetchSource, this, arguments); };
        new PerformanceObserver(list => timing.tasks.push(...list.getEntries().map(e => ({ start: e.startTime, duration: e.duration })))).observe({ type: 'longtask' });
        (function check() {
          if (timing.start !== null && window.__qa.ready() && !document.querySelector('.status-load__skip') && window.__qa.screenText()?.includes(${JSON.stringify(text)})) timing.settled = performance.now();
          else requestAnimationFrame(check);
        })();
      </script>`;
      app = await openApp({ name: `preflight-opening-${disabled ? 'local' : 'worker'}`, port, extra: setup, settle: 0 });
      for (let i = 0; i < 200 && !(await app.js('window.__packageOpening.settled !== null')); i++) await sleep(50);
      const result = await app.js(`(() => {
        const t = window.__packageOpening;
        if (t.start === null || t.settled === null) throw new Error('The requested fixture did not finish opening');
        return { mode: ${JSON.stringify(disabled ? 'local-fallback' : 'worker')}, rendererCpuRate: 1, sourceToSettledDocumentMs: t.settled - t.start,
          longTasks: t.tasks.filter(task => task.start >= t.start && task.start < t.settled), worker: window.__packageWorkers };
      })()`);
      if (disabled) {
        // Vite may try both Blob and data URLs after a constructor failure.
        assert(result.worker.blocked >= 1);
        assert.equal(result.worker.created, 0);
        assert.equal(result.worker.responses, 0);
      } else {
        assert.equal(result.worker.responses, 1);
        assert.equal(result.worker.terminated, 1);
        assert.equal(result.worker.failed, 0);
      }
      // Functional editing proof is separate from the opening timestamp above.
      await app.caret(0);
      const before = await app.js(`(async () => {
        const doc = window.__qa.doc(), selection = await doc.selection.current();
        const target = selection.target, segment = target?.segments?.[0];
        if (!selection.empty || !segment || (target.story && target.story.storyType !== 'body') || target.segments.length !== 1) throw new Error('Expected a body caret');
        const node = await doc.getNodeById({ nodeId: segment.blockId });
        return { blockId: segment.blockId, offset: segment.range.start, text: node.node.paragraph.inlines.map(i => i.run?.text ?? '').join('') };
      })()`);
      const token = 'Qz';
      assert(!before.text.includes(token));
      await app.type(token);
      const expected = before.text.slice(0, before.offset) + token + before.text.slice(before.offset);
      const readTyping = () => app.js(`(async () => {
        const node = await window.__qa.doc().getNodeById({ nodeId: ${JSON.stringify(before.blockId)} });
        const model = node.node.paragraph.inlines.map(i => i.run?.text ?? '').join('');
        const fragment = window.__otzariaEditor.container.querySelector('.superdoc-fragment[data-source-node-id="' + CSS.escape(${JSON.stringify(before.blockId)}) + '"]');
        const visible = fragment && [...fragment.querySelectorAll('.superdoc-text-run')].map(n => n.textContent).join('');
        return { model, visible };
      })()`);
      // The engine model can reply before its asynchronous layout reaches the
      // DOM. This is a functional gate, so wait for the exact result rather than
      // treating an arbitrary two frames as an editing latency requirement.
      const typingDeadline = Date.now() + 10000;
      let typed;
      do {
        typed = await readTyping();
        if (typed.model === expected && typed.visible === expected) break;
        await sleep(30);
      } while (Date.now() < typingDeadline);
      await app.js('new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))');
      typed = await readTyping();
      assert.equal(typed.model, expected, 'Canonical model contains the keyboard input');
      assert.equal(typed.visible, expected, 'Rendered paragraph contains the keyboard input');
      result.typingValidated = true;
      report.opening.push(result);
      app.close(); app = null;
    }
  }
} catch (error) {
  report.error = error.stack ?? String(error);
  process.exitCode = 1;
} finally {
  forget(); cleanup();
  const output = resolve(args.out); mkdirSync(dirname(output), { recursive: true });
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify({ output, packageComparisons: report.package.length, cancellation: report.cancellation, opening: report.opening, error: report.error }, null, 2));
}
