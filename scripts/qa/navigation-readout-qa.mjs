/** Real SuperDoc 2.23 / file:// contract for selective navigation and ¶ reads. */
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { openApp, ROOT } from './harness.mjs';
import { buildDocx, p, table } from './docx-fixtures.mjs';

const compiled = await build({
  stdin: { contents: `
    export { readNavOutline } from './src/engine/navigation-pane';
    export { readFormattingMarksBlocks } from './src/engine/formatting-marks';
    export { measureFormattingMarksWindow } from './src/engine/formatting-marks-viewport';
    export { buildMappedPilcrowMarks } from './src/engine/formatting-marks-layer';
  `, resolveDir: ROOT, loader: 'ts' },
  bundle: true, format: 'iife', globalName: 'NavigationPerf', write: false,
});
const long = 'LONG ' + 'מילים חוזרות לבדיקת פסקה חוצה עמודים '.repeat(360) + ' END';
const bytes = buildDocx({ body: p('NAVIGATION PROBE', '<w:pStyle w:val="Heading1"/>') +
  p('טקסט זהה', '<w:bidi/>') + p('') + p('טקסט זהה', '<w:bidi/>') + table(p('בתוך טבלה')) +
  p(long, '<w:bidi/>') + p('כותרת אחרונה', '<w:pStyle w:val="Heading1"/>') +
  Array.from({ length: 120 }, (_, i) => p(i % 13 === 0 ? '' : 'פסקה רגילה עם טקסט זהה', '<w:bidi/>')).join('') });
const app = await openApp({
  name: 'navigation-readout', port: Number(process.env.QA_PORT ?? 9790),
  extra: '<script>' + compiled.outputFiles[0].text + '</script>',
});
async function until(predicate) {
  for (let i = 0; i < 120; i++) { if (await predicate()) return; await app.sleep(250); }
  throw new Error('Timed out waiting for the requested document/UI state');
}
try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1400, height: 900, deviceScaleFactor: 1, mobile: false });
  await app.js(`window.__qaHost.emit('plugin.page_opened', ${JSON.stringify({ param: {
    name: 'navigation.docx', dataUrl: 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + bytes.toString('base64'),
  } })})`);
  await until(() => app.js(`document.querySelector('.superdoc-page')?.textContent?.includes('NAVIGATION PROBE') && !document.querySelector('.status-load__skip')`));
  await app.sleep(1200);
  const initial = JSON.parse(await app.js(`(async () => {
    const actual = window.__qa.doc();
    const calls = [];
    const blocks = new Proxy(actual.blocks, { get(target, name) {
      return name === 'list' ? (...args) => { calls.push(args[0]); return target.list(...args); } : target[name];
    }});
    const doc = new Proxy(actual, { get(target, name) { return name === 'blocks' ? blocks : target[name]; }});
    const host = { activeEditor: { doc } };
    const outline = await NavigationPerf.readNavOutline(host);
    const before = calls.length;
    const same = await NavigationPerf.readNavOutline(host);
    const nav = { headings: outline?.headings.map(h => ({ level: h.level, ordinal: h.ordinal })), ordinals: outline?.ordinals.size,
      calls: calls.slice(), cacheHit: same === outline, repeatedCalls: calls.length - before };
    const viewport = window.__qa.ui().viewport.getHost();
    const reference = document.querySelector('.pilcrow-layer');
    window.__readMarksProbe = async () => {
      const before = calls.length;
      const painted = NavigationPerf.measureFormattingMarksWindow(viewport, reference);
      const ids = painted.mode === 'mapped' ? [...new Set(painted.endings.map(end => end.nodeId))] : null;
      const selected = await NavigationPerf.readFormattingMarksBlocks(host, ids);
      const marks = painted.mode === 'mapped' ? NavigationPerf.buildMappedPilcrowMarks(selected ?? [], painted.endings) : [];
      return { mode: painted.mode, ids, selected: selected?.map(b => ({ nodeId: b.nodeId, type: b.nodeType, length: b.text.length, empty: b.isEmpty })),
        endings: painted.mode === 'mapped' ? painted.endings.map(end => ({ nodeId: end.nodeId, continuesFromPrevious: end.continuesFromPrevious })) : [],
        marks: marks.map(m => ({ nodeId: m.nodeId, approximate: m.approximate })), calls: calls.slice(before) };
    };
    return JSON.stringify({ nav, top: await window.__readMarksProbe() });
  })()`));
  assert.equal(initial.nav.ordinals, 127);
  assert.deepEqual(initial.nav.headings, [{ level: 1, ordinal: 0 }, { level: 1, ordinal: 6 }]);
  assert.equal(initial.nav.calls.length, 2);
  assert.equal(initial.nav.calls[0].includeText, false);
  assert.equal(initial.nav.calls[1].nodeIds.length, 2);
  assert.equal(initial.nav.repeatedCalls, 0);
  assert.equal(initial.top.mode, 'mapped');
  assert.equal(initial.top.marks.length, 3);
  assert.equal(initial.top.selected.filter(b => b.empty).length, 1);
  assert.ok(initial.top.selected.every(b => b.type === 'paragraph'));
  assert.ok(initial.top.calls.every(call => call.nodeIds?.length && call.nodeTypes?.[0] === 'paragraph'));

  // This fixture's final fragment of LONG is on page 4. Scroll the real host;
  // checking the resulting offset prevents a no-op scroll from passing.
  await app.js('window.__qa.ui().viewport.getHost().scrollTop = 3500');
  await until(() => app.js('window.__qa.ui().viewport.getHost().scrollTop === 3500'));
  await app.sleep(1200);
  const bottom = JSON.parse(await app.js('(async () => JSON.stringify(await window.__readMarksProbe()))()'));
  assert.equal(bottom.mode, 'mapped');
  const continued = bottom.endings.find(end => end.continuesFromPrevious);
  assert.ok(continued, 'fixture must exercise a paragraph ending after a page split');
  assert.ok(bottom.marks.some(mark => mark.nodeId === continued.nodeId));
  assert.ok(bottom.ids.length < 30, 'text requests stay in the current reading window');

  // The actual product wiring must render the same scoped marks as the tested
  // reader, and must remove them when disabled.
  await app.tab('תצוגה');
  assert.ok(await app.click('סימני עיצוב', { after: 900 }));
  // Ribbon focus restoration can reveal the current caret at the document's
  // start. Restore the exercised reading window before comparing its marks.
  await app.js('window.__qa.ui().viewport.getHost().scrollTop = 3500');
  await until(() => app.js('window.__qa.ui().viewport.getHost().scrollTop === 3500'));
  await until(async () => (await app.js('document.querySelectorAll(".pilcrow-layer__mark").length')) === bottom.marks.length);
  assert.ok(await app.click('סימני עיצוב', { after: 300 }));
  assert.equal(await app.js('document.querySelectorAll(".pilcrow-layer__mark").length'), 0);
  assert.ok(await app.click('חלונית ניווט', { after: 700 }));
  await until(async () => (await app.js('document.querySelectorAll(".nav-pane [role=treeitem]").length')) === 2);
  console.log(JSON.stringify({ protocol: 'file:', nav: initial.nav, topParagraphs: initial.top.marks.length,
    visibleParagraphs: bottom.marks.length, continuedParagraphVerified: true, productToggleVerified: true }, null, 2));
} finally { app.close(); }
