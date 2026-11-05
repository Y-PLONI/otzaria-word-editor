/** Validated, opt-in startup/typing benchmark. See docs/performance-measurements.md. */
import { readFileSync, writeFileSync, mkdirSync, rmSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { cpus, totalmem, platform, release, arch } from 'node:os';
import { createHash, randomUUID } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { parseArgs } from 'node:util';
import { connect, openPage, requireChrome, sleep, onInterrupt } from './cdp.mjs';
import { installPerformanceRecorder } from './performance-recorder.mjs';
import { summarize } from './performance-report.mjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const { values: args } = parseArgs({ options: {
  out: { type: 'string' }, dist: { type: 'string', default: join(root, 'dist') },
  docx: { type: 'string' }, runs: { type: 'string' }, keys: { type: 'string', default: '60' },
  gap: { type: 'string', default: '40' }, pause: { type: 'string', default: '3000' },
  cache: { type: 'string', default: 'fresh-profile' }, port: { type: 'string', default: '9790' },
  'api-bytes': { type: 'boolean' }, smoke: { type: 'boolean' }, autosave: { type: 'boolean' },
  attach: { type: 'string' }, target: { type: 'string' }, reload: { type: 'boolean' },
  'edit-copy': { type: 'boolean' }, help: { type: 'boolean' },
} });
if (args.help) {
  console.log('node scripts/performance-probe.mjs --out report.json [--docx fixture.docx] [--runs 5] [--cache fresh-profile|reload] [--smoke] [--api-bytes] [--autosave]\nNative: --attach PORT --target CDP_PAGE_ID --edit-copy [--reload]. Uses a disposable copy; typing changes it. See docs/performance-measurements.md.');
  process.exit(0);
}
const integer = (value, name, min, max) => {
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number < min || number > max) throw new Error(`Invalid --${name}`);
  return number;
};
if (!args.out) throw new Error('--out is required');
if (!['fresh-profile', 'reload'].includes(args.cache)) throw new Error('Invalid --cache');
if (args.attach && (!args.target || !args['edit-copy'])) throw new Error('Native attachment requires --target and --edit-copy');
if (args.attach && args.docx) throw new Error('--docx belongs to the headless fixture run');
if (args.reload && !args.attach) throw new Error('Use --cache reload for headless runs');
const count = integer(args.runs ?? (args.smoke || (args.attach && !args.reload) ? '1' : '5'), 'runs', 1, 50);
if (!args.smoke && !(args.attach && !args.reload) && count < 5) throw new Error('Use at least five startup runs, or --smoke');
if (args.attach && !args.reload && count !== 1) throw new Error('Late attachment supports one run; use --reload for startup repetitions');
const keyCount = integer(args.keys, 'keys', args.smoke ? 2 : 40, 200);
const gap = integer(args.gap, 'gap', 0, 1000);
const pause = integer(args.pause, 'pause', 1000, 30_000);
const port = integer(args.attach ?? args.port, 'port', 1024, 65535);
const hash = buffer => createHash('sha256').update(buffer).digest('hex');
const limitFile = name => {
  try { return readFileSync('/sys/fs/cgroup/' + name, 'utf8').trim(); } catch { return null; }
};
const memoryLimit = limitFile('memory.max');
const cpuLimit = limitFile('cpu.max');
const fixture = args.docx ? readFileSync(resolve(args.docx)) : null;
const dist = resolve(args.dist);
const runId = randomUUID();
const injection = `(${installPerformanceRecorder.toString()})(${JSON.stringify({ ownerId: runId, apiBytes: Boolean(args['api-bytes']), boundary: args.attach && !args.reload ? 'late-attach' : 'plugin-navigation-start' })});`;
const owned = body => `(() => { const recorder = window.__otzariaPerformance; if (recorder?.ownerId !== ${JSON.stringify(runId)}) throw new Error('Recording ownership changed'); ${body} })()`;
const output = resolve(args.out);
const pagePath = join(dist, `__performance-${runId}.html`);
let browser = null, nativeScript = null, cdp = null;
let revision = null, sourceDirty = null;
try {
  revision = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim();
  sourceDirty = execFileSync('git', ['status', '--porcelain'], { cwd: root, encoding: 'utf8' }).trim() !== '';
} catch { /* exported source */ }
const report = {
  schemaVersion: 1, recordedAt: new Date().toISOString(), revision, sourceDirty,
  tools: ['performance-probe.mjs', 'performance-recorder.mjs', 'performance-report.mjs']
    .map(script => ({ script, sha256: hash(readFileSync(join(root, 'scripts', script))) })),
  environment: { kind: args.attach ? 'attached-CDP' : 'headless-file', platform: platform(), release: release(), arch: arch(),
    cpu: cpus()[0]?.model ?? null, logicalCpus: cpus().length, ramBytes: totalmem(), node: process.version,
    cgroupMemoryMax: memoryLimit, cgroupCpuMax: cpuLimit,
    browser: null, nativeReleaseVerified: false, osCacheControlled: false,
    browserFlags: args.attach ? 'host-controlled' : 'scripts/cdp.mjs; headless; GPU and background throttling disabled' },
  scenario: { document: fixture ? 'restored-docx' : args.attach ? 'host-current-document' : 'new-blank',
    fixtureSha256: fixture ? hash(fixture) : null, fixtureBytes: fixture?.length ?? null,
    cache: args.attach ? (args.reload ? 'host-page-reload' : 'late-attach') : args.cache,
    runs: count, keysPerBurst: keyCount, gapMs: gap, pauseMs: pause, apiBytes: Boolean(args['api-bytes']),
    autosaveRequested: Boolean(args.autosave), io: args.attach ? 'existing-host-unmodified' : 'in-memory-host-stub',
    readouts: args.attach ? 'host preferences; record the comparison matrix separately' : 'default preferences, home ribbon',
    smoke: Boolean(args.smoke) },
  artifacts: [], runs: [], summary: null,
};

function fixtureScript(base64, autosave) {
  const host = window.__qaHost;
  host.storage['autosave-enabled'] = autosave;
  if (base64) {
    const bytes = Uint8Array.from(atob(base64), char => char.charCodeAt(0));
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
    host.replies['fs.resolveFileUrl'] = { url, name: 'benchmark.docx', size: bytes.length };
    host.storage.session = { version: 2, activeId: 'benchmark',
      documents: [{ id: 'benchmark', document: { token: 'benchmark', name: 'benchmark.docx', writable: autosave }, caret: null, draft: null, previousDraft: null }],
      view: { zoom: 100, focusMode: false, ribbonTab: 'home', ribbonCollapsed: false } };
  }
  // The engine still performs its real export. Only host disk/network IO is
  // replaced; this scenario must not be reported as native save performance.
  host.replies['fs.beginBinaryWrite'] = { writeToken: 'benchmark-write', uploadUrl: 'https://benchmark.invalid/upload', maxBytes: 0 };
  host.replies['fs.commitUserFileWrite'] = { cancelled: false, token: 'benchmark', name: 'benchmark.docx' };
  const originalFetch = window.fetch;
  window.fetch = function (input, options) {
    if (input === 'https://benchmark.invalid/upload' && options?.method === 'PUT') return Promise.resolve(new Response('', { status: 200 }));
    return Reflect.apply(originalFetch, this, arguments);
  };
}

async function until(expression, timeout = 60_000, interval = 25) {
  const end = Date.now() + timeout;
  while (Date.now() < end) {
    const result = await cdp.evaluate(expression);
    if (result) return result;
    await sleep(interval);
  }
  throw new Error('Timed out waiting for the measured condition');
}
async function key(character) {
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: character, text: character });
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: character });
}
async function typeBurst(phase, text, { waitForSave = false, documentId = null } = {}) {
  const prepared = await cdp.evaluate(owned(`return recorder.prepare(${JSON.stringify(phase)}, ${JSON.stringify(text)});`));
  if (documentId !== null && documentId !== prepared.documentId) throw new Error('The active document changed between typing phases');
  if (phase === 'first-edit' && !args.attach) {
    const recording = await cdp.evaluate(owned('return recorder.snapshot();'));
    const opening = recording.events.find(event => event.name === 'document.open.start' && event.span === prepared.openingId);
    const opened = recording.events.some(event => event.name === 'document.open.end' && event.span === prepared.openingId && event.status === 'opened');
    if (!opened || opening?.kind !== (fixture ? 'file' : 'new')) {
      throw new Error('The requested fixture/opening did not become the measured document');
    }
  }
  if (waitForSave) await until(`(() => { for (const s of window.__otzariaEditors.values()) if (s.swap.current === window.__otzariaEditor) return s.save.snapshot.isSaving; return false; })()`, 12_000, 10);
  for (const character of text) { await key(character); if (gap) await sleep(gap); }
  await until(owned(`const p = recorder.pending(); return p && p.received === ${text.length} && p.displayed === p.received;`), 20_000);
  const result = await cdp.evaluate(owned('return recorder.finishBurst();'));
  if (!result.valid) throw new Error('The typed sequence did not validate against the intended document');
  return { ...result, documentId: prepared.documentId, openingId: prepared.openingId };
}
async function capture(index) {
  const run = { index, valid: false, failureStage: null, startupHooksObserved: false, recording: null, resources: [], resourceTimingAvailable: false, host: null, bursts: [] };
  report.runs.push(run);
  let stage = 'ready';
  try {
    await until(`(() => { const e = window.__otzariaEditor; return !!(window.__otzariaPerformance?.ownerId === ${JSON.stringify(runId)} && e?.ui?.selection?.getSnapshot()?.selectionTarget && !document.getElementById('otzaria-splash')); })()`);
    await cdp.evaluate(owned("recorder.mark('driver.ready-observed');"));
    report.environment.browser = (await cdp.send('Browser.getVersion')).result;
    run.host = await cdp.evaluate(`(() => { const b = window.__otzariaBoot?.payload; return b ? { appVersion: b.app?.version ?? null, platform: b.app?.platform ?? null, pluginVersion: b.plugin?.version ?? null } : null; })()`);
    stage = 'first-edit';
    run.bursts.push(await typeBurst('first-edit', 'א'));
    const target = { documentId: run.bursts[0].documentId };
    if (args.autosave && !args.attach && !fixture) await cdp.evaluate(`(() => { for (const s of window.__otzariaEditors.values()) if (s.swap.current === window.__otzariaEditor) s.save.adoptTarget({ token: 'benchmark', name: 'benchmark.docx' }); })()`);
    const letters = 'אבגדהוזחטיכלמנסעפצקרשת';
    const text = Array.from({ length: keyCount }, (_, i) => letters[i % letters.length]).join('');
    stage = 'continuous';
    run.bursts.push(await typeBurst('continuous', text, target));
    stage = 'after-pause';
    await sleep(pause);
    run.bursts.push(await typeBurst('after-pause', text, target));
    if (args.autosave) {
      stage = 'autosave-overlap';
      run.bursts.push(await typeBurst('autosave-overlap', text, { ...target, waitForSave: true }));
    }
    run.recording = await cdp.evaluate(owned("recorder.mark('measurement.finished'); return recorder.snapshot();"));
    run.startupHooksObserved = run.recording.events.some(event => event.name === 'document.open.end' && event.status === 'opened');
    if (run.recording.dropped || !run.recording.firstEdit || run.recording.keys.some(sample => !sample.validated)) throw new Error('Incomplete recording');
    if (!args.attach && !run.startupHooksObserved) throw new Error('Build lacks startup instrumentation');
    if (args.autosave && !run.recording.keys.some(sample => sample.phase === 'autosave-overlap' && ['exporting', 'uploading', 'committing'].includes(sample.saveState))) throw new Error('No keystroke overlapped a save');
    run.resources = await cdp.evaluate(String.raw`performance.getEntriesByType('resource').filter(r => /\/(app|engine-workers)\.js(?:[?#]|$)/.test(r.name)).map(r => ({ asset: r.name.includes('engine-workers') ? 'engine-workers.js' : 'app.js', start: r.startTime, duration: r.duration, fetchEnd: r.responseEnd, transferBytes: r.transferSize > 0 ? r.transferSize : null, encodedBytes: r.encodedBodySize > 0 ? r.encodedBodySize : null }))`);
    run.resourceTimingAvailable = run.resources.length > 0;
    run.valid = true;
  } catch (error) {
    run.failureStage = stage;
    try { run.recording = await cdp.evaluate(owned('return recorder.snapshot();')); } catch { /* page unavailable or owned by a different recording */ }
    console.error(`Run ${index} failed at ${stage}: ${error.message}`);
  }
  console.log(`Run ${index}/${count}: ${run.valid ? 'validated' : 'FAILED'}`);
}

const cleanup = () => { browser?.close(); if (!args.attach) rmSync(pagePath, { force: true }); };
let nativeCleanup = null;
const closeNative = () => nativeCleanup ??= (async () => {
  if (!args.attach || !cdp) return;
  try {
    try { await cdp.evaluate(`if (window.__otzariaPerformance?.ownerId === ${JSON.stringify(runId)}) window.__otzariaPerformance.stop();`); } catch { /* host closed */ }
    if (nativeScript) await cdp.send('Page.removeScriptToEvaluateOnNewDocument', { identifier: nativeScript });
  } finally { cdp.close(); }
})();
const forget = onInterrupt(async () => { await closeNative(); cleanup(); });
try {
  if (args.attach) {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`, { signal: AbortSignal.timeout(5000) })).json();
    const target = targets.find(item => item.id === args.target && item.type === 'page');
    if (!target) throw new Error('The requested CDP page target was not found');
    cdp = await connect(target.webSocketDebuggerUrl);
    await cdp.send('Page.enable');
    if (await cdp.evaluate('!!window.__otzariaPerformance')) throw new Error('A performance recording is already installed');
    if (args.reload) {
      const response = await cdp.send('Page.addScriptToEvaluateOnNewDocument', { source: injection });
      nativeScript = response.result?.identifier;
      if (!nativeScript) throw new Error('Could not install the navigation recorder');
    } else await cdp.evaluate(injection);
  } else {
    requireChrome();
    let html = readFileSync(join(dist, 'index.html'), 'utf8');
    const firstScriptEnd = html.indexOf('</script>');
    if (firstScriptEnd < 0 || !html.slice(0, firstScriptEnd).includes('__otzariaBoot')) throw new Error('Expected production boot latch as the first script');
    const host = readFileSync(join(root, 'scripts/qa/host-stub.js'), 'utf8');
    const stub = `<script>${host}\n(${fixtureScript.toString()})(${JSON.stringify(fixture?.toString('base64') ?? null)}, ${Boolean(args.autosave)});</script>`;
    html = html.slice(0, firstScriptEnd + 9) + stub + html.slice(firstScriptEnd + 9);
    html = html.replace(/<head[^>]*>/, match => match + `<script>${injection}</script>`);
    writeFileSync(pagePath, html);
    for (const asset of ['index.html', 'assets/app.js', 'assets/engine-workers.js']) {
      const content = readFileSync(join(dist, asset));
      report.artifacts.push({ asset, bytes: content.length, sha256: hash(content) });
    }
  }
  for (let index = 1; index <= count; index += 1) {
    if (!args.attach && (!browser || args.cache === 'fresh-profile')) {
      browser?.close();
      browser = await openPage(pathToFileURL(pagePath).href, { port, label: `performance-${runId}-${index}`, waitForLoad: false });
      cdp = browser.cdp;
    } else if (index > 1 || args.reload) {
      await cdp.evaluate(`(() => { const recorder = window.__otzariaPerformance; if (recorder && recorder.ownerId !== ${JSON.stringify(runId)}) throw new Error('Recording ownership changed'); recorder?.stop(); })()`);
      await cdp.send('Page.reload', { ignoreCache: false });
    }
    await capture(index);
    // Failure is a structural signal, not a slow sample to omit and retry.
    if (!report.runs.at(-1).valid) break;
  }
} finally {
  await closeNative();
  forget(); cleanup();
  report.summary = summarize(report.runs);
  mkdirSync(dirname(output), { recursive: true });
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
}
console.log(JSON.stringify(report.summary, null, 2));
if (report.runs.length !== count || report.runs.some(run => !run.valid)) process.exitCode = 1;
