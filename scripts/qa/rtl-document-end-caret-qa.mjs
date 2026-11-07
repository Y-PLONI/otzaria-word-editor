/**
 * Last-document-line ArrowDown, on the packaged engine with real input.
 * Imported by check:vertical-arrows. Expected offsets come from fixture text,
 * never from the painted PM attributes used by the implementation.
 */
import assert from 'node:assert/strict';
import { buildDocx, p } from './docx-fixtures.mjs';
import { unzip } from './harness.mjs';

const RTL_SOURCE = 'שורת מקור רחבה מספיק כדי לבדוק חזרה מדויקת אל העמודה המקורית';
const LTR_SOURCE = 'A reference line with enough width to preserve the starting column';
const MARKER = 'QZ161Q';
const CASES = [
  { id: 'rtl', label: 'סוף המסמך RTL', rtl: true, text: 'השורה האחרונה במסמך לבדיקה' },
  { id: 'ltr', label: 'סוף המסמך LTR', rtl: false, text: 'Final English line for this document' },
  {
    id: 'wrapped', label: 'סוף המסמך בפסקה גולשת', rtl: true, wrapped: true,
    text: 'פסקה אחרונה זו ארוכה דיה כדי להישבר לשורות אחדות ברוחב העמוד הרגיל והבדיקה מתחילה באמצע השורה החזותית האחרונה ואז מבקשת להגיע אל סוף הטקסט הלוגי ולחזור אל העמודה המקורית',
  },
  { id: 'mixed', label: 'סוף המסמך בטקסט מעורב ומנוקד', rtl: true, text: 'שורת סיום עם שָׁלוֹם ועם Latin 123' },
];

export async function runDocumentEndCaretQa(app, report) {
  const key = (name) => app.press(name, name, name === 'ArrowDown' ? 40 : 38);

  async function waitFor(label, read, accept, timeout = 5000) {
    const deadline = Date.now() + timeout;
    let value;
    do {
      value = await read();
      if (accept(value)) return value;
      await app.sleep(80);
    } while (Date.now() < deadline);
    throw new Error(`${label}: ${JSON.stringify(value)}`);
  }

  // Bound Document API/export reads too: a stalled worker must fail the gate,
  // rather than leave CDP's awaitPromise pending forever.
  const bounded = (expression, timeout = 3000) => app.js(`(async () => {
    let timer;
    try {
      return await Promise.race([
        Promise.resolve().then(() => (${expression})),
        new Promise((_, reject) => { timer = setTimeout(() => reject(new Error('engine read timed out')), ${timeout}); }),
      ]);
    } finally { clearTimeout(timer); }
  })()`);

  const layout = () => app.js(`(() => {
    const paragraphs = [];
    for (const fragment of document.querySelectorAll('.superdoc-fragment[data-source-node-id]')) {
      const rect = fragment.getBoundingClientRect();
      if (!rect.width || !rect.height) continue;
      const blockId = fragment.getAttribute('data-source-node-id');
      const lines = [];
      for (const line of fragment.querySelectorAll('.superdoc-line')) {
        const runs = [...line.querySelectorAll('.superdoc-text-run')];
        const boxes = runs.map(run => run.getBoundingClientRect()).filter(box => box.width && box.height);
        if (!boxes.length) continue;
        lines.push({ blockId, text: runs.map(run => run.textContent || '').join(''),
          left: Math.min(...boxes.map(box => box.left)), right: Math.max(...boxes.map(box => box.right)),
          top: Math.min(...boxes.map(box => box.top)), bottom: Math.max(...boxes.map(box => box.bottom)) });
      }
      paragraphs.push({ blockId,
        text: [...fragment.querySelectorAll('.superdoc-text-run')].map(run => run.textContent || '').join(''), lines });
    }
    return { loading: !!document.querySelector('.status-load'), paragraphs };
  })()`);

  const selection = () => bounded(`(async () => {
    const info = await window.__otzariaEditor.superdoc.activeEditor.doc.selection.current();
    const target = info?.selectionTarget;
    const caret = [...document.querySelectorAll('.sd-v2-local-selection-caret')]
      .map(node => node.getBoundingClientRect()).find(rect => rect.height > 0 && Number.isFinite(rect.left));
    return { start: target?.start, end: target?.end,
      x: caret?.left, y: caret ? (caret.top + caret.bottom) / 2 : null };
  })()`);

  const collapsedAt = (state, blockId, offset) => Number.isInteger(offset) && state.start?.blockId === blockId
    && state.end?.blockId === blockId && state.start.offset === offset && state.end.offset === offset;
  const onLine = (state, line) => state.start?.blockId === line.blockId
    && Number.isFinite(state.y) && state.y >= line.top - 2 && state.y <= line.bottom + 2;

  async function openFixture(test, source) {
    const properties = test.rtl ? '<w:bidi/><w:jc w:val="right"/>' : '<w:bidi w:val="0"/><w:jc w:val="left"/>';
    const bytes = buildDocx({ body: p(source, properties) + p(test.text, properties) });
    const name = `vertical-document-end-${test.id}`;
    const url = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,'
      + Buffer.from(bytes).toString('base64');
    // Different tokens are essential: reopening the same token activates its
    // existing (possibly edited) tab instead of loading the next fixture.
    await app.js(`window.__qaHost.replies['fs.pickUserFile'] = () => Promise.resolve({ success: true, error: null,
      data: { token: ${JSON.stringify(`tok-${name}`)}, url: ${JSON.stringify(url)},
        name: ${JSON.stringify(`${name}.docx`)}, size: ${bytes.length}, access: 'readwrite' } })`);
    await app.tab('קובץ');
    assert(await app.click('פתח קובץ', { after: 250 }), 'open-file button missing');
    await waitFor('file picker', () => app.exists('.open-browse'), Boolean, 10000);
    await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
    assert(await app.clickSel('.open-browse', 0, { after: 250 }), 'browse button missing');
    await waitFor('fresh fixture ready', layout, value => !value.loading && value.paragraphs.length === 2
      && value.paragraphs[0].text === source && value.paragraphs[1].text === test.text, 20000);
    await app.tab('בית');
    return layout();
  }

  async function exportedParagraphs() {
    const base64 = await bounded('window.__qa.exportBase64()', 15000);
    assert(base64, 'DOCX export missing');
    const xml = unzip(Buffer.from(base64, 'base64'))['word/document.xml'];
    assert(xml, 'word/document.xml missing');
    return app.js(`(() => {
      const xml = new DOMParser().parseFromString(${JSON.stringify(xml)}, 'application/xml');
      if (xml.querySelector('parsererror')) throw new Error('invalid exported XML');
      const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const body = xml.getElementsByTagNameNS(ns, 'body')[0];
      return [...body.children].filter(node => node.namespaceURI === ns && node.localName === 'p')
        .map(node => [...node.getElementsByTagNameNS(ns, 't')].map(text => text.textContent).join(''));
    })()`);
  }

  for (const test of CASES) {
    console.log(`\n== ${test.label}`);
    try {
      const source = test.rtl ? RTL_SOURCE : LTR_SOURCE;
      const initial = await openFixture(test, source);
      const finalParagraph = initial.paragraphs[1];
      const last = finalParagraph.lines.at(-1);
      assert(last, 'last visual line missing');
      if (test.wrapped) assert(finalParagraph.lines.length > 1, 'fixture must actually soft-wrap');
      const previousParagraph = finalParagraph.lines.length > 1 ? 1 : 0;
      const previousIndex = initial.paragraphs[previousParagraph].lines.length - (previousParagraph === 1 ? 2 : 1);
      const previous = initial.paragraphs[previousParagraph].lines[previousIndex];
      assert(previous, 'preceding visual line missing');
      const left = Math.max(last.left, previous.left);
      const right = Math.min(last.right, previous.right);
      assert(right - left > 16, 'fixture needs a shared interior goal column');
      await app.clickAt((left + right) / 2, (last.top + last.bottom) / 2);
      const before = await waitFor('caret inside final visual line', selection, state =>
        collapsedAt(state, finalParagraph.blockId, state.start?.offset)
        && state.start.offset > 0 && state.start.offset < test.text.length && onLine(state, last));

      const atEnd = state => collapsedAt(state, finalParagraph.blockId, test.text.length);
      await key('ArrowDown');
      await waitFor('ArrowDown at document end', selection, atEnd);
      report.pass(`${test.label} — למטה מגיע לסוף הלוגי`, `${before.start.offset} → ${test.text.length}`);

      for (let i = 0; i < 3; i += 1) {
        await key('ArrowDown');
        await app.sleep(150);
        const repeated = await selection();
        assert(atEnd(repeated), `repeated ArrowDown moved selection: ${JSON.stringify(repeated)}`);
      }
      report.pass(`${test.label} — הקשות חוזרות נשארות בסוף`, '3 repetitions');

      await key('ArrowUp');
      const restored = await waitFor('ArrowUp restores original goal column', async () => {
        const state = await selection();
        const current = (await layout()).paragraphs[previousParagraph]?.lines[previousIndex];
        const goal = current ? before.x + current.left - previous.left : NaN;
        const tolerance = current ? Math.max(4, (current.right - current.left) / current.text.length) * 1.5 : 0;
        return { state, current, gap: Math.abs(state.x - goal), tolerance };
      }, value => value.current && onLine(value.state, value.current)
        && collapsedAt(value.state, value.current.blockId, value.state.start?.offset) && value.gap <= value.tolerance);
      report.pass(`${test.label} — למעלה שומר את העמודה`, `x=${before.x.toFixed(1)} → ${restored.state.x.toFixed(1)}`);

      // Return by real keys, wait for the last line, then exercise the new
      // branch again before typing. No programmatic selection writes.
      await key('ArrowDown');
      await waitFor('back on final line', async () => ({ state: await selection(), view: await layout() }),
        value => {
          const line = value.view.paragraphs[1]?.lines.at(-1);
          return line && onLine(value.state, line) && value.state.start.offset < test.text.length;
        });
      await key('ArrowDown');
      await waitFor('ready to append', selection, atEnd);
      await app.type(MARKER);
      await waitFor('marker inserted', selection, state => collapsedAt(state, finalParagraph.blockId, test.text.length + MARKER.length));
      assert.deepEqual(await exportedParagraphs(), [source, test.text + MARKER], 'typed marker must append in the saved DOCX');
      report.pass(`${test.label} — הקלדה נוספת נשמרת בסוף`, MARKER);
    } catch (error) {
      report.fail(test.label, String(error?.stack || error));
    }
  }
}
