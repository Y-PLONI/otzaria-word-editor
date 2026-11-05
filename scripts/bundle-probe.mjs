/**
 * Built file:// gate: all startup scripts, lazy panels, failed-load retry,
 * and real PDF/EMF rendering. --startup-only --root <checkout> also measures
 * the old IIFE build without modifying its source. Each run uses a fresh Chrome
 * profile. CHROME and QA_PORT follow the existing QA harness conventions.
 */
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { emfDocx, rectanglePdf } from './qa/bundle-fixtures.mjs';
import { onInterrupt } from './cdp.mjs';
import { openApp } from './qa/harness.mjs';

const args = process.argv.slice(2);
const option = (name, fallback) => args.includes(name) ? args[args.indexOf(name) + 1] : fallback;
const root = resolve(option('--root', fileURLToPath(new URL('..', import.meta.url))));
const out = option('--out', null);
const startupOnly = args.includes('--startup-only');
const cpuRate = Number(option('--cpu', '1'));
const dist = join(root, 'dist');
const bundle = existsSync(join(dist, 'bundle-report.json')) ? JSON.parse(readFileSync(join(dist, 'bundle-report.json'), 'utf8')) : null;

const INSTRUMENT = `<script>
(function () {
  var probe = window.__bundleProbe = { scripts: [], shellReadyAt: null, paints: [], phase: 'startup', errorEvents: [] };
  window.addEventListener('error', function (event) {
    probe.errorEvents.push({ phase: probe.phase, at: performance.now(), message: event.message, filename: event.filename, lineno: event.lineno, colno: event.colno, error: event.error == null ? null : String(event.error) });
  });
  new PerformanceObserver(function (list) {
    list.getEntries().forEach(function (entry) { probe.paints.push({ name: entry.name, at: entry.startTime }); });
  }).observe({ type: 'paint', buffered: true });
  new MutationObserver(function (records) {
    records.forEach(function (record) {
      record.addedNodes.forEach(function (node) {
        if (node.tagName !== 'SCRIPT' || !node.src) return;
        var url = new URL(node.src);
        var item = { file: url.pathname.split('/').slice(-2).join('/'), requestedAt: performance.now(), loadedAt: null, failedAt: null };
        probe.scripts.push(item);
        node.addEventListener('load', function () { item.loadedAt = performance.now(); });
        node.addEventListener('error', function () { item.failedAt = performance.now(); });
      });
    });
    if (probe.shellReadyAt === null && window.__qa && window.__qa.ready()) probe.shellReadyAt = performance.now();
  }).observe(document.documentElement, { childList: true, subtree: true });
})();
</script>`;

const app = await openApp({ name: 'bundle-probe', port: Number(process.env.QA_PORT ?? 9385), extra: INSTRUMENT, settle: 700, buildDirectory: dist, cpuRate });
const result = { format: bundle?.format ?? 'classic-iife', cpuRate, startup: null, checks: {} };
let heldFile;
const restoreHeld = () => {
  if (heldFile && existsSync(`${heldFile}.held`)) renameSync(`${heldFile}.held`, heldFile);
  heldFile = undefined;
};
const forgetHeldCleanup = onInterrupt(restoreHeld);
const json = async (expression) => JSON.parse(await app.js(`JSON.stringify(${expression})`));
const wait = async (expression, label) => {
  for (let attempt = 0; attempt < 150; attempt++) {
    if (await app.js(expression)) return;
    await app.sleep(100);
  }
  throw new Error(`Timed out: ${label}`);
};
const requests = (file) => json(`window.__bundleProbe.scripts.filter(item => item.file === ${JSON.stringify(file)})`);
const chunkFor = (pattern) => {
  const found = Object.entries(bundle.chunks).find(([, chunk]) => chunk.modules.some((module) => pattern.test(module.id)));
  assert.ok(found, `chunk for ${pattern}`);
  return found[0];
};
const close = async () => { await app.press('Escape', 'Escape', 27); await app.sleep(100); };

try {
  result.browser = await app.cdp.send('Browser.getVersion');
  result.testPage = await app.js('location.href');
  assert.equal(await app.js('location.protocol'), 'file:');
  result.startup = await json('window.__bundleProbe');
  result.startup.factoryMarks = await json('performance.getEntriesByType("mark").filter(item => item.name.startsWith("otzaria:module:")).map(item => ({ name: item.name, at: item.startTime }))');
  result.startup.resources = await json('performance.getEntriesByType("resource").filter(item => item.name.startsWith("file:")).map(item => ({ file: new URL(item.name).pathname.split("/").slice(-2).join("/"), at: item.startTime, duration: item.duration }))');
  result.startup.files = Object.fromEntries([...new Set(result.startup.scripts.map((item) => item.file))].map((file) => {
    const bytes = readFileSync(join(dist, file));
    return [file, { bytes: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex') }];
  }));
  result.startup.scriptBytes = Object.values(result.startup.files).reduce((sum, item) => sum + item.bytes, 0);
  const index = readFileSync(join(dist, 'index.html'));
  result.startup.index = { bytes: index.length, sha256: createHash('sha256').update(index).digest('hex') };
  assert.ok(result.startup.shellReadyAt > 0);
  assert.ok(result.startup.scripts.every((item) => item.loadedAt !== null && item.failedAt === null));
  assert.deepEqual(await json('window.__qa.log.filter(item => /uncaught:|rejected:/.test(item))'), []);

  if (!startupOnly) {
    assert.ok(bundle, 'split build report');
    // Static graph separation alone is insufficient: ordinary document opening
    // could eagerly call a dynamic import. Assert the actual browser requests.
    for (const pattern of [
      /pdfjs-dist\/build\/pdf\.mjs$/,
      /superdoc\/dist\/chunks\/pdf\.worker[^/]+\.js$/,
      /src\/ui\/panels\/(?:MacrosDialog|OpenDocumentDialog|ShortcutManagerDialog)\.vue$/,
    ]) {
      for (const [file, chunk] of Object.entries(bundle.chunks)) {
        if (chunk.modules.some((module) => pattern.test(module.id))) {
          assert.deepEqual(await requests(file), [], `${file} was requested at startup`);
        }
      }
    }
    for (const file of bundle.startupFiles) assert.equal((await requests(file)).length, 1, `${file} startup request count`);
    result.checks.ordinaryDocumentDefersChunks = true;
    const appFactoryReady = result.startup.factoryMarks.find(item => item.name === 'otzaria:module:assets/app:ready');
    assert.ok(appFactoryReady, 'entry factory timing exists');
    assert.ok(bundle.preloadFiles?.length, 'mandatory DOCX preload plan exists');
    for (const file of bundle.preloadFiles) {
      const [request] = await requests(file);
      assert.ok(request.requestedAt < appFactoryReady.at, `${file} fetch must start before the app factory completes`);
    }
    result.checks.docxPreloadBeforeAppFactory = true;
    await app.js("window.__bundleProbe.phase = 'typing'");
    const blockId = await app.js('(async () => (await window.__qa.doc().selection.current()).target?.segments?.[0]?.blockId)()');
    assert.equal(typeof blockId, 'string', 'an editable body caret exists');
    await app.type('z', 0);
    const visible = `Array.from(document.querySelectorAll('.superdoc-fragment[data-source-node-id="' + ${JSON.stringify(blockId)} + '"] .superdoc-text-run')).some(node => node.textContent === 'z' && node.getBoundingClientRect().width > 0)`;
    await wait(visible, 'typed character is visible');
    assert.equal(await app.js(`(async () => {
      const result = await window.__qa.doc().getNodeById({ nodeId: ${JSON.stringify(blockId)} });
      return result.address.nodeId === ${JSON.stringify(blockId)} && result.node.paragraph.inlines.some(inline => inline.kind === 'run' && inline.run.text === 'z');
    })()`), true, 'the same character reached the document API');
    await app.js('new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))');
    result.checks.typingModelAndVisibleParagraph = true;
    await app.press('z', 'KeyZ', 90, 2);
    result.offline = { enable: await app.cdp.send('Network.enable') };
    assert.equal(result.offline.enable.error, undefined, 'CDP Network.enable succeeds');
    result.offline.emulation = await app.cdp.send('Network.emulateNetworkConditions', { offline: true, latency: 0, downloadThroughput: -1, uploadThroughput: -1 });
    assert.equal(result.offline.emulation.error, undefined, 'CDP offline emulation succeeds');
    result.offline.navigatorOnLine = await app.js('navigator.onLine');
    assert.equal(result.offline.navigatorOnLine, false, 'the page is offline before loading deferred chunks');

    await app.js("window.__bundleProbe.phase = 'open-dialog'");
    const openFile = chunkFor(/src\/ui\/panels\/OpenDocumentDialog\.vue$/);
    await app.js('window.__bundleFocusOrigin = document.activeElement');
    heldFile = join(dist, openFile);
    renameSync(heldFile, `${heldFile}.held`);
    await app.press('o', 'KeyO', 79, 2);
    await wait('!!document.querySelector(".panel-load-state [role=alert]")', 'missing chunk is recoverable');
    renameSync(`${heldFile}.held`, heldFile);
    heldFile = undefined;
    await app.js('document.querySelector(".panel-load-state button").focus()');
    assert.equal(await app.click('נסה שוב'), true);
    await wait('!!document.querySelector(".open-dialog")', 'open dialog retry');
    assert.equal(await app.js('!!document.activeElement?.closest(".open-dialog")'), true, 'first-open focus');
    await close();
    await wait('!document.querySelector(".open-dialog")', 'open dialog close');
    assert.equal(await app.js('document.activeElement === window.__bundleFocusOrigin'), true, 'focus returns after Retry and close');
    await app.press('o', 'KeyO', 79, 2);
    await wait('!!document.querySelector(".open-dialog")', 'open dialog repeat');
    await close();
    await wait('!document.querySelector(".open-dialog")', 'open dialog close again');
    const openRequests = await requests(openFile);
    assert.equal(openRequests.length, 2);
    assert.ok(openRequests[0].failedAt !== null && openRequests[1].loadedAt !== null);
    result.checks.openDialogFailureRetryFocusAndRepeat = true;

    await app.js("window.__bundleProbe.phase = 'shortcut-manager'");
    const shortcutFile = chunkFor(/src\/ui\/panels\/ShortcutManagerDialog\.vue$/);
    heldFile = join(dist, shortcutFile);
    renameSync(heldFile, `${heldFile}.held`);
    await app.tab('קובץ');
    assert.equal(await app.click('ניהול קיצורים'), true);
    await wait('!!document.querySelector(".panel-load-state [role=alert]")', 'shortcut manager missing chunk');
    await close();
    await wait('!document.querySelector(".panel-load-state")', 'Escape dismisses a failed lazy manager');
    renameSync(`${heldFile}.held`, heldFile);
    heldFile = undefined;
    for (let attempt = 0; attempt < 2; attempt++) {
      assert.equal(await app.click('ניהול קיצורים'), true);
      await wait('!!document.querySelector(".shortmgr-dialog")', 'shortcut manager');
      await close();
      await wait('!document.querySelector(".shortmgr-dialog")', 'shortcut manager close');
    }
    const shortcutRequests = await requests(shortcutFile);
    assert.equal(shortcutRequests.length, 2);
    assert.ok(shortcutRequests[0].failedAt !== null && shortcutRequests[1].loadedAt !== null);
    result.checks.shortcutsDismissFailureFirstAndRepeat = true;

    await app.js("window.__bundleProbe.phase = 'macros'");
    const macrosFile = chunkFor(/src\/ui\/panels\/MacrosDialog\.vue$/);
    await app.tab('מפתחים');
    for (let attempt = 0; attempt < 2; attempt++) {
      assert.equal(await app.click('ניהול מאקרו'), true);
      await wait('!!document.querySelector(".macros-dialog")', 'macro manager');
      assert.equal(await app.click('סגור', { scope: '.macros-dialog', selector: '.md-close-btn', exact: true }), true);
      await wait('!document.querySelector(".macros-dialog")', 'macro manager close');
    }
    assert.equal((await requests(macrosFile)).length, 1);
    result.checks.macrosFirstAndRepeat = true;

    const emf = emfDocx();
    const url = `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${emf.toString('base64')}`;
    for (let attempt = 0; attempt < 2; attempt++) {
      await app.js(`window.__bundleProbe.phase = 'emf-${attempt}'`);
      await app.js(`window.__qaHost.replies['fs.pickUserFile'] = ${JSON.stringify({ token: `emf-${attempt}`, url, name: `bundle-emf-${attempt}.docx`, size: emf.length, access: 'readwrite' })}`);
      await app.press('o', 'KeyO', 79, 2);
      await wait('!!document.querySelector(".open-dialog")', 'EMF open dialog');
      assert.equal(await app.click('עיון בקבצים'), true);
      await wait(`document.querySelector('.doc-title-input')?.value === 'bundle-emf-${attempt}'`, 'EMF document title');
      await wait('Array.from(document.images).some(img => img.naturalWidth >= 50 && img.getBoundingClientRect().width > 0)', 'EMF rendered image');
      const pixel = await app.js(`(function () {
        var img = Array.from(document.images).find(img => img.naturalWidth >= 50 && img.getBoundingClientRect().width > 0);
        var canvas = document.createElement('canvas'); canvas.width = img.naturalWidth; canvas.height = img.naturalHeight;
        var ctx = canvas.getContext('2d'); ctx.drawImage(img, 0, 0);
        return JSON.stringify(Array.from(ctx.getImageData(Math.floor(canvas.width / 2), Math.floor(canvas.height / 2), 1, 1).data));
      })()`);
      const rgba = JSON.parse(pixel);
      assert.ok(rgba[2] > 200 && rgba[0] < 50 && rgba[1] < 50 && rgba[3] === 255, `EMF blue center: ${rgba}`);
    }
    result.checks.emfFirstAndRepeat = true;

    await app.js("window.__bundleProbe.phase = 'pdf'");
    const pdfFile = chunkFor(/pdfjs-dist\/build\/pdf\.mjs$/);
    const workerFile = chunkFor(/superdoc\/dist\/chunks\/pdf\.worker[^/]+\.js$/);
    const pdf = rectanglePdf().toString('base64');
    for (let attempt = 0; attempt < 2; attempt++) {
      const pixel = await app.js(`(async function () {
        try {
        var modules = window.__otzariaModules;
        var pair = await Promise.all([modules.load(${JSON.stringify(pdfFile.replace(/\.js$/, ''))}), modules.load(${JSON.stringify(workerFile.replace(/\.js$/, ''))})]);
        var pdfjs = pair[0]; pdfjs.GlobalWorkerOptions.workerSrc = pair[1].default;
        var task = pdfjs.getDocument({ data: Uint8Array.from(atob(${JSON.stringify(pdf)}), ch => ch.charCodeAt(0)) });
        var doc = await task.promise; var page = await doc.getPage(1); var viewport = page.getViewport({ scale: 1 });
        var canvas = document.createElement('canvas'); canvas.width = viewport.width; canvas.height = viewport.height;
        var ctx = canvas.getContext('2d'); await page.render({ canvasContext: ctx, viewport: viewport }).promise;
        var pixel = Array.from(ctx.getImageData(50, 50, 1, 1).data); await doc.destroy(); return JSON.stringify(pixel);
        } catch (error) {
          return JSON.stringify({ name: error && error.name, message: String(error && error.message || error), stack: error && error.stack });
        }
      })()`);
      assert.deepEqual(JSON.parse(pixel), [0, 0, 255, 255]);
    }
    assert.equal((await requests(pdfFile)).length, 1);
    assert.equal((await requests(workerFile)).length, 1);
    result.checks.pdfFirstAndRepeat = true;
    result.allScripts = await json('window.__bundleProbe.scripts');
    await app.js('new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))');
    const errorCountBeforeQuiet = await app.js('window.__bundleProbe.errorEvents.length');
    await app.sleep(500);
    result.errorEvents = await json('window.__bundleProbe.errorEvents');
    result.errorQuietWindow = { durationMs: 500, beforeCount: errorCountBeforeQuiet, afterCount: result.errorEvents.length };
    result.unhandled = await json('window.__qa.log.filter(item => /uncaught:|rejected:/.test(item))');
    // Chrome 157 emits this native delivery notice on the unchanged baseline
    // too. Retain it in the report; do not excuse a JS Error/Promise rejection,
    // a similarly worded error, or an observer loop that continues after settle.
    const resizeMessage = 'ResizeObserver loop completed with undelivered notifications.';
    result.resizeObserverWarnings = result.errorEvents.filter(event => event.message === resizeMessage &&
      event.lineno === 0 && event.error === null && event.filename === result.testPage);
    let notices = result.resizeObserverWarnings.length;
    result.unexpectedErrors = result.unhandled.filter(message => {
      if (message === `uncaught: ${resizeMessage}` && notices > 0) { notices--; return false; }
      return true;
    });
    assert.equal(result.errorEvents.length, errorCountBeforeQuiet, 'no recurring browser errors after settle');
    assert.deepEqual(result.unexpectedErrors, []);
  }
  console.log(JSON.stringify({ format: result.format, cpuRate, shellReadyMs: result.startup.shellReadyAt, scriptBytes: result.startup.scriptBytes, checks: result.checks }));
} catch (error) {
  result.failure = String(error);
  throw error;
} finally {
  restoreHeld();
  forgetHeldCleanup();
  app.close();
  if (out) writeFileSync(out, `${JSON.stringify(result, null, 2)}\n`);
}
