// Test-only entry: the same package code/inline worker as the product build.
import { exportDocx } from '../../src/engine/export';
import { postflightDocx } from '../../src/engine/docx-postflight';
import { readDocxPart } from '../../src/engine/docx-preflight-core';

const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
const epoch = () => performance.timeOrigin + performance.now();
const decode = base64 => Uint8Array.from(atob(base64), char => char.charCodeAt(0));
const digest = async bytes => Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)))
  .map(byte => byte.toString(16).padStart(2, '0')).join('');
async function until(predicate) {
  const deadline = performance.now() + 30000;
  while (performance.now() < deadline) { if (await predicate()) return; await delay(50); }
  throw new Error('Postflight QA state did not settle');
}
async function firstExportedParagraph(blob) {
  if (!(blob instanceof Blob)) throw new Error('No uploaded snapshot');
  const xml = await readDocxPart(new Uint8Array(await blob.arrayBuffer()), 'word/document.xml');
  if (!xml) throw new Error('Uploaded snapshot has no document.xml');
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  const body = document.getElementsByTagNameNS(ns, 'body')[0];
  const paragraph = Array.from(body?.children ?? []).find(node => node.localName === 'p');
  if (!paragraph) throw new Error('Uploaded snapshot has no first paragraph');
  return Array.from(paragraph.getElementsByTagNameNS(ns, 't')).map(node => node.textContent).join('');
}

function visible(element) {
  if (!element?.isConnected || document.visibilityState !== 'visible') return false;
  const rect = element.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0 || rect.top < 0 || rect.left < 0 || rect.bottom > innerHeight || rect.right > innerWidth) return false;
  for (let node = element; node instanceof Element; node = node.parentElement) {
    const style = getComputedStyle(node);
    if (style.display === 'none' || style.visibility === 'hidden' || Number(style.opacity) === 0) return false;
    if (node !== element) {
      const bounds = node.getBoundingClientRect();
      const left = bounds.left + node.clientLeft, top = bounds.top + node.clientTop;
      if (/^(auto|scroll|hidden|clip)$/.test(style.overflowX) && (rect.left < left || rect.right > left + node.clientWidth)) return false;
      if (/^(auto|scroll|hidden|clip)$/.test(style.overflowY) && (rect.top < top || rect.bottom > top + node.clientHeight)) return false;
    }
  }
  return true;
}
function renderedParagraph(editor, blockId) {
  const selector = `.superdoc-fragment[data-source-node-id="${CSS.escape(blockId)}"]`;
  const fragments = [...editor.container.querySelectorAll(selector)].filter(visible);
  // This fixture uses one short body paragraph. Refuse a split/duplicated
  // fragment instead of guessing how its painted text maps to the block.
  if (fragments.length !== 1) return null;
  return [...fragments[0].querySelectorAll('.superdoc-text-run')].map(node => node.textContent ?? '').join('');
}
function bodyCaret(selection, blockId, offset) {
  const target = selection?.target, segments = target?.segments;
  if (target?.kind !== 'text' || segments?.length !== 1) return false;
  if (![target.story, selection.selectionTarget?.story, selection.selectionTarget?.start?.story, selection.selectionTarget?.end?.story]
    .every(story => !story || story.storyType === 'body')) return false;
  const segment = segments[0];
  return segment.blockId === blockId && segment.range?.start === offset && segment.range?.end === offset;
}

let typing;
function sameTypingDocument() {
  return window.__otzariaEditor === typing.editor && typing.session.swap.current === typing.editor
    && typing.editor.superdoc === typing.host && window.__qa.sd() === typing.host;
}
window.__postflightQA = {
  async measure(base64, mode) {
    const original = new Blob([decode(base64)]);
    const tasks = [];
    const supported = PerformanceObserver.supportedEntryTypes.includes('longtask');
    const observer = supported ? new PerformanceObserver(list => tasks.push(...list.getEntries())) : null;
    const before = { ...window.__postflightTrace.workers };
    await delay(0);
    observer?.observe({ type: 'longtask' });
    const start = performance.now();
    let output;
    if (mode === 'main') {
      const bytes = await postflightDocx(new Uint8Array(await original.arrayBuffer()));
      output = bytes ? new Blob([bytes]) : original;
    } else output = await exportDocx({ export: async () => original });
    const end = performance.now();
    await delay(60);
    if (observer) { tasks.push(...observer.takeRecords()); observer.disconnect(); }
    const during = tasks.filter(task => task.startTime < end && task.startTime + task.duration > start);
    const bytes = await output.arrayBuffer();
    return {
      mode, wallMs: end - start, outputBytes: bytes.byteLength,
      sha256: await digest(bytes), unchanged: output === original,
      longestTaskMs: supported ? Math.max(0, ...during.map(task => task.duration)) : null,
      mainBlockingMs: supported ? during.reduce((total, task) => total + Math.max(0, task.duration - 50), 0) : null,
      workers: Object.fromEntries(Object.entries(window.__postflightTrace.workers).map(([key, value]) => [key, value - before[key]])),
    };
  },

  async prepareTyping() {
    const host = window.__qa.sd();
    const session = [...window.__otzariaEditors.values()].find(item => item.swap.current?.superdoc === host);
    if (!session) throw new Error('Active document session was not found');
    // Measure exactly the two explicit saves below. A delayed renderer must not
    // let the ordinary autosave timer add an unlabelled third save to the sample.
    session.save.setAutosaveEnabled(false);
    const first = (await host.activeEditor.doc.blocks.list({ offset: 0, limit: 1, includeText: true })).blocks[0];
    if (!first || typeof first.text !== 'string') throw new Error('Missing first paragraph');
    const point = { kind: 'text', blockId: first.nodeId, offset: first.text.length };
    await host.ui.selection.apply({ kind: 'selection', start: point, end: point });
    host.focus();
    typing = { host, session, editor: session.swap.current, blockId: first.nodeId, before: first.text, keys: [], outcome: null, end: null };
    await until(async () => sameTypingDocument() && bodyCaret(await host.activeEditor.doc.selection.current(), first.nodeId, first.text.length)
      && renderedParagraph(typing.editor, first.nodeId) === first.text);
    typing.listener = event => {
      if (event.code !== 'KeyX' || event.key !== 'x') return;
      const receivedEpochMs = epoch();
      const sample = {
        trusted: event.isTrusted, targetInsideEditor: typing.editor.container.contains(event.target), sameDocument: sameTypingDocument(),
        eventEpochMs: performance.timeOrigin + event.timeStamp, receivedEpochMs, twoFramesEpochMs: null,
      };
      typing.keys.push(sample);
      // Scheduling opportunity only: keydown can precede engine application.
      // Canonical and painted acceptance are checked separately after the burst.
      requestAnimationFrame(() => requestAnimationFrame(() => { sample.twoFramesEpochMs = epoch(); }));
    };
    window.addEventListener('keydown', typing.listener, true);
    return { blockId: first.nodeId, textLength: first.text.length };
  },

  startSave() {
    const trace = window.__postflightTrace;
    trace.enabled = true;
    typing.start = epoch();
    // Force one explicit save even though the freshly opened fixture is clean.
    typing.promise = typing.session.save.saveNow({ forceSaveAs: true }).then(outcome => {
      typing.outcome = outcome; typing.end = epoch(); return outcome;
    });
  },

  async finishTyping(count) {
    await typing.promise;
    const expected = typing.before + 'x'.repeat(count);
    let selection, current;
    await until(async () => {
      if (!sameTypingDocument()) throw new Error('The active document changed during the typing sample');
      current = (await typing.host.activeEditor.doc.blocks.list({ nodeIds: [typing.blockId], includeText: true, offset: 0, limit: 1 })).blocks[0];
      selection = await typing.host.activeEditor.doc.selection.current();
      return sameTypingDocument() && current?.nodeId === typing.blockId && current.text === expected
        && bodyCaret(selection, typing.blockId, expected.length) && renderedParagraph(typing.editor, typing.blockId) === expected;
    });
    await until(() => typing.keys.length === count && typing.keys.every(key => key.twoFramesEpochMs !== null));
    window.removeEventListener('keydown', typing.listener, true);
    const trace = window.__postflightTrace;
    trace.enabled = false;
    const dirtyAfterOldCommit = typing.session.save.snapshot.isDirty;
    const uploaded = await firstExportedParagraph(trace.uploads[0]);
    const result = {
      outcome: typing.outcome, totalSaveMs: typing.end - typing.start,
      phase: trace.phases[0], packagePhases: trace.phases.length,
      beforeLength: typing.before.length, afterLength: current.text.length,
      activeDocumentMatches: sameTypingDocument(), targetMatches: bodyCaret(selection, typing.blockId, expected.length),
      typedTextMatches: current.text === expected, uploadedOldSnapshot: uploaded === typing.before,
      paintedTextMatches: renderedParagraph(typing.editor, typing.blockId) === expected,
      dirtyAfterOldCommit,
      keys: typing.keys, workers: { ...trace.workers },
    };
    // The next save must commit the new snapshot and clear only that revision.
    result.nextOutcome = await typing.session.save.saveNow();
    result.nextSnapshotMatches = await firstExportedParagraph(trace.uploads[1]) === expected;
    result.dirtyAfterNextCommit = typing.session.save.snapshot.isDirty;
    return result;
  },
};
