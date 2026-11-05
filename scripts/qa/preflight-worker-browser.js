// Test-only entry, bundled with the same IIFE worker format as the product.
import { preflightSource, preflightPackage } from '../../src/engine/docx-preflight';
import { runPackageTask } from '../../src/engine/docx-package-client';
import { readZip } from '../../src/engine/docx-parts';

const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
const decode = base64 => Uint8Array.from(atob(base64), char => char.charCodeAt(0));
const digest = async bytes => Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)))
  .map(byte => byte.toString(16).padStart(2, '0')).join('');

window.__preflightQA = {
  async measure(base64, mode) {
    const input = decode(base64), original = new Blob([input]);
    const tasks = [];
    const supported = PerformanceObserver.supportedEntryTypes.includes('longtask');
    const observer = supported ? new PerformanceObserver(list => tasks.push(...list.getEntries())) : null;
    const before = { ...window.__packageWorkers };
    // Keep fixture decoding outside the measured event-loop task.
    await delay(0);
    observer?.observe({ type: 'longtask' });
    const start = performance.now();
    let result;
    if (mode === 'main') {
      const local = await preflightPackage(new Uint8Array(await original.arrayBuffer()));
      result = { ...local, source: local.bytes ? new Blob([local.bytes]) : original };
    } else result = await preflightSource(original);
    const end = performance.now();
    // Long-task entries are delivered after their task has completed.
    await delay(60);
    if (observer) { tasks.push(...observer.takeRecords()); observer.disconnect(); }
    const during = tasks.filter(task => task.startTime < end && task.startTime + task.duration > start);
    const bytes = await result.source.arrayBuffer();
    const sourceVba = readZip(input)?.find(entry => entry.name === 'word/vbaProject.bin');
    const resultVba = readZip(new Uint8Array(bytes))?.find(entry => entry.name === 'word/vbaProject.bin');
    const vbaBytesPreserved = sourceVba
      ? !!resultVba && (await digest(sourceVba.data)) === (await digest(resultVba.data)) : null;
    const after = window.__packageWorkers;
    return {
      mode, wallMs: end - start,
      longTasks: supported ? during.map(task => ({ startMs: task.startTime - start, durationMs: task.duration })) : null,
      longestTaskMs: supported ? Math.max(0, ...during.map(task => task.duration)) : null,
      mainBlockingMs: supported ? during.reduce((total, task) => total + Math.max(0, task.duration - 50), 0) : null,
      unchanged: result.source === original,
      outputBytes: bytes.byteLength, sha256: await digest(bytes),
      vbaBytesPreserved,
      fontTable: result.fontTable, vba: result.vba, notice: result.notice,
      worker: Object.fromEntries(Object.entries(after).map(([key, value]) => [key, value - before[key]])),
    };
  },
  async cancel(base64) {
    const bytes = decode(base64).buffer;
    const controller = new AbortController();
    const before = { ...window.__packageWorkers };
    const pending = runPackageTask('preflight', bytes, controller.signal);
    controller.abort();
    let cancelled = false;
    try { await pending; } catch (error) { cancelled = error.name === 'AbortError'; }
    await delay(30);
    return {
      cancelled, detached: bytes.byteLength === 0,
      terminated: window.__packageWorkers.terminated - before.terminated,
      created: window.__packageWorkers.created - before.created,
    };
  },
};
