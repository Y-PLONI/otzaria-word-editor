/**
 * Opt-in browser recorder. Kept outside the shipped bundle and injected by the
 * benchmark before navigation. This function is deliberately self-contained.
 * No document text, file name, token, URL or API payload enters snapshot().
 */
export function installPerformanceRecorder({ apiBytes = false, limit = 4000, boundary = 'plugin-navigation-start', ownerId = null } = {}) {
  if (window.__otzariaPerformance) throw new Error('A performance recording is already installed');
  if (!Number.isSafeInteger(limit) || limit < 1 || limit > 50_000) throw new Error('Invalid recording limit');
  if (!['plugin-navigation-start', 'late-attach'].includes(boundary)) throw new Error('Invalid recording boundary');
  const now = () => performance.now();
  const events = [], longTasks = [], keys = [], api = new Map(), coverage = [];
  const editors = new WeakMap(), disposers = new Set();
  let spanId = 0, documentId = 0, dropped = 0, stopped = false, burst = null;
  let longObserver = null, firstEdit = null;
  const push = (list, value) => { if (list.length < limit) list.push(value); else dropped += 1; };
  const mark = (name, details = {}) => {
    if (!stopped) push(events, { name, at: now(), ...details });
  };
  const bytes = value => {
    if (!apiBytes) return null;
    try {
      const encoded = JSON.stringify(value);
      return encoded === undefined ? 0 : new TextEncoder().encode(encoded).byteLength;
    } catch { return null; }
  };
  function measureApi(scope, method, request) {
    if (stopped) return () => {};
    const key = scope + ':' + method;
    let row = api.get(key);
    if (!row) {
      row = { scope, method, calls: 0, completed: 0, failed: 0, totalMs: 0,
        requestBytes: apiBytes ? 0 : null, responseBytes: apiBytes ? 0 : null, unknownBytes: 0 };
      api.set(key, row);
    }
    const started = now();
    row.calls += 1;
    const size = bytes(request);
    if (size !== null) row.requestBytes += size;
    else if (apiBytes) row.unknownBytes += 1;
    let ended = false;
    return (response, status) => {
      if (ended || stopped) return;
      ended = true;
      row.completed += 1;
      row.failed += status === 'error' ? 1 : 0;
      row.totalMs += now() - started;
      const length = bytes(response);
      if (length !== null) row.responseBytes += length;
      else if (apiBytes) row.unknownBytes += 1;
    };
  }

  // Read-only public methods used by the app's expensive readouts. These are
  // counts at this facade, not all internal engine RPCs or wire byte counts.
  const METHODS = ['get', 'getText', 'getNode', 'getNodeById', 'blocks.list', 'sections.list', 'selection.current'];
  function bind(editor, openingId = null) {
    if (stopped) return;
    if (editors.has(editor)) return editors.get(editor);
    const item = { id: ++documentId, openingId, originals: new Map(), disposed: false };
    editors.set(editor, item);
    const doc = editor.superdoc?.activeEditor?.doc;
    const undo = [];
    const dispose = () => {
      if (item.disposed) return;
      item.disposed = true;
      if (burst?.item === item) clearBurst();
      for (const restore of undo) restore();
      disposers.delete(dispose);
    };
    disposers.add(dispose);
    try { editor.onDispose?.(dispose); } catch { /* stop() still restores the facade */ }
    if (item.disposed) return item;
    for (const path of METHODS) {
      const parts = path.split('.');
      let owner = doc;
      for (const part of parts.slice(0, -1)) owner = owner?.[part];
      const name = parts.at(-1);
      const descriptor = owner && Object.getOwnPropertyDescriptor(owner, name);
      const fn = descriptor?.value;
      const status = typeof fn !== 'function' ? 'unavailable' : !descriptor.writable ? 'readonly' : 'wrapped';
      const covered = { documentId: item.id, method: path, status };
      push(coverage, covered);
      if (typeof fn !== 'function') continue;
      item.originals.set(path, { owner, fn });
      if (status !== 'wrapped') continue;
      const wrapped = function (...args) {
        const end = measureApi('document', path, args);
        try {
          const result = Reflect.apply(fn, this, args);
          if (result && typeof result.then === 'function') {
            result.then(value => end(value, 'ok'), () => end(undefined, 'error'));
          } else end(result, 'ok');
          return result; // Preserve the original promise and receiver.
        } catch (error) { end(undefined, 'error'); throw error; }
      };
      try { Object.defineProperty(owner, name, { ...descriptor, value: wrapped }); }
      catch { covered.status = 'unavailable'; continue; }
      undo.push(() => {
        try { if (owner[name] === wrapped) Object.defineProperty(owner, name, descriptor); }
        catch { /* an engine which sealed its facade must still be disposable */ }
      });
    }
    return item;
  }
  function begin(name, details = {}) {
    const id = ++spanId;
    mark(name + '.start', { span: id, ...details });
    let ended = false;
    return {
      mark(stage) { if (!ended) mark(stage, { span: id }); },
      bind(editor) { bind(editor, id); },
      end(status) {
        if (!ended) mark(name + '.end', { span: id, status });
        ended = true;
      },
    };
  }
  function visible(element, complete = false) {
    if (!element?.isConnected || document.visibilityState !== 'visible') return false;
    const rect = element.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0 || rect.bottom <= 0 || rect.top >= innerHeight ||
        rect.right <= 0 || rect.left >= innerWidth) return false;
    if (complete && (rect.top < 0 || rect.left < 0 || rect.bottom > innerHeight || rect.right > innerWidth)) return false;
    for (let node = element; node instanceof Element; node = node.parentElement) {
      const style = getComputedStyle(node);
      if (style.display === 'none' || style.visibility === 'hidden' || Number(style.opacity) === 0) return false;
      if (complete && node !== element) {
        const clipsX = /^(auto|scroll|hidden|clip)$/.test(style.overflowX);
        const clipsY = /^(auto|scroll|hidden|clip)$/.test(style.overflowY);
        if (clipsX || clipsY) {
          const bounds = node.getBoundingClientRect();
          const left = bounds.left + node.clientLeft, top = bounds.top + node.clientTop;
          if (clipsX && (rect.left < left || rect.right > left + node.clientWidth)) return false;
          if (clipsY && (rect.top < top || rect.bottom > top + node.clientHeight)) return false;
        }
      }
    }
    return true;
  }
  function renderedText(editor, blockId) {
    const selector = '.superdoc-fragment[data-source-node-id="' + CSS.escape(blockId) + '"]';
    const fragments = [...editor.container.querySelectorAll(selector)].filter(node => visible(node, true));
    // A split/virtualized paragraph needs a different fixture location. Do not
    // concatenate duplicate fragments or invent a logical text offset from them.
    if (fragments.length !== 1) return null;
    return [...fragments[0].querySelectorAll('.superdoc-text-run')]
      .map(node => node.textContent ?? '').join('');
  }
  function textOf(result, blockId) {
    if (result?.address?.nodeId !== blockId) throw new Error('The node address does not match the measured block');
    if (result.address.story && result.address.story.storyType !== 'body') throw new Error('The measured node is outside the body story');
    const node = result.node;
    const inlines = node?.paragraph?.inlines ?? node?.heading?.inlines;
    if (!Array.isArray(inlines)) throw new Error('Choose a plain body paragraph for the typing sample');
    const flatten = nodes => nodes.map(inline => {
      if (inline.kind === 'run' && typeof inline.run?.text === 'string') return inline.run.text;
      if (inline.kind === 'hyperlink') return flatten(inline.hyperlink.inlines);
      if (inline.kind === 'tab') return '\t';
      if (inline.kind === 'lineBreak') return '\n';
      throw new Error('The sampled paragraph contains unsupported inline content');
    }).join('');
    return flatten(inlines);
  }
  function callProbe(item, path, ...args) {
    const original = item.originals.get(path);
    if (!original) throw new Error('Public document API unavailable: ' + path);
    // Probe reads are counted separately; they must not look like app readouts.
    const end = measureApi('probe', path, args);
    return Promise.resolve().then(() => Reflect.apply(original.fn, original.owner, args)).then(
      value => { end(value, 'ok'); return value; },
      error => { end(undefined, 'error'); throw error; },
    );
  }
  function saveState(editor) {
    for (const session of window.__otzariaEditors?.values?.() ?? []) {
      if (session.swap?.current === editor) return session.save.snapshot.state;
    }
    return 'unknown';
  }
  function sameDocument(current) {
    return window.__otzariaEditor === current.editor && !current.item.disposed && visible(current.editor.container);
  }
  function bodyTarget(target) {
    return [target?.story, target?.start?.story, target?.end?.story]
      .every(story => !story || story.storyType === 'body');
  }
  function expectedText(current, count) {
    return current.before.slice(0, current.offset) + current.text.slice(0, count) + current.before.slice(current.offset);
  }
  function renderOpportunity(current) {
    if (current.frame !== null || burst !== current) return;
    const candidates = current.samples.filter(sample => sample.visibleDomAt !== null && sample.renderOpportunityAt === null);
    if (!candidates.length) return;
    current.frame = requestAnimationFrame(() => {
      current.frame = requestAnimationFrame(() => {
        current.frame = null;
        if (burst !== current || !sameDocument(current)) return;
        if (renderedText(current.editor, current.blockId) === expectedText(current, current.seen)) {
          const presented = now();
          // Only these keys had visible DOM before BOTH animation frames.
          // Keys arriving between the callbacks need their own opportunity.
          for (const sample of candidates) sample.renderOpportunityAt = presented;
        } else inspect(current);
        renderOpportunity(current);
      });
    });
  }
  function inspect(current) {
    const at = now();
    if (burst !== current || !sameDocument(current)) return;
    const text = renderedText(current.editor, current.blockId);
    let count = current.received;
    while (count > current.seen && text !== expectedText(current, count)) count -= 1;
    if (count <= current.seen) return;
    const grouped = count - current.seen > 1;
    for (let i = current.seen; i < count; i += 1) {
      current.samples[i].visibleDomAt = at;
      current.samples[i].coalesced = grouped;
    }
    current.seen = count;
    renderOpportunity(current);
  }
  function keydown(event) {
    const current = burst;
    if (!current || !current.editor.container.contains(event.target) || event.ctrlKey || event.altKey || event.metaKey) return;
    if (event.isComposing || event.key.length !== 1) return;
    const downAt = now();
    if (!sameDocument(current) || event.key !== current.text[current.received]) {
      current.invalid = true;
      return;
    }
    if (keys.length >= limit) { dropped += 1; current.invalid = true; return; }
    const sample = { sequence: keys.length, phase: current.phase, documentId: current.item.id,
      downAt, visibleDomAt: null, renderOpportunityAt: null, coalesced: false,
      saveState: saveState(current.editor), validated: false };
    keys.push(sample);
    current.samples.push(sample);
    current.received += 1;
  }
  function clearBurst() {
    if (!burst) return;
    burst.observer.disconnect();
    if (burst.frame !== null) cancelAnimationFrame(burst.frame);
    burst = null;
  }
  async function prepare(phase, text) {
    if (stopped || burst) throw new Error('Recording stopped or previous burst not validated');
    if (typeof text !== 'string' || !/^[a-z\u05d0-\u05ea]+$/u.test(text) || text.length > 500) {
      throw new Error('Use 1–500 lowercase Latin or Hebrew probe letters');
    }
    const editor = window.__otzariaEditor;
    if (!editor || !visible(editor.container)) throw new Error('No visible active editor');
    const item = bind(editor);
    const target = editor.ui.selection.getSnapshot().selectionTarget;
    const start = target?.start, end = target?.end;
    if (!bodyTarget(target)) throw new Error('Select the body story');
    if (start?.kind !== 'text' || end?.kind !== 'text' || typeof start.blockId !== 'string' || start.blockId !== end.blockId ||
        start.offset !== end.offset || !Number.isSafeInteger(start.offset)) throw new Error('A collapsed text caret is required');
    const before = textOf(await callProbe(item, 'getNodeById', { nodeId: start.blockId }), start.blockId);
    const latest = editor.ui.selection.getSnapshot().selectionTarget;
    if (window.__otzariaEditor !== editor || item.disposed || !visible(editor.container) ||
        !bodyTarget(latest) || latest?.start?.kind !== 'text' || latest?.end?.kind !== 'text' ||
        latest?.start?.blockId !== start.blockId || latest?.end?.blockId !== end.blockId ||
        latest?.start?.offset !== start.offset || latest?.end?.offset !== end.offset) {
      throw new Error('The active document or caret changed while preparing the sample');
    }
    if (start.offset < 0 || start.offset > before.length || renderedText(editor, start.blockId) !== before) {
      throw new Error('The full sampled paragraph must be visible and match its document API text');
    }
    const current = { phase, text, before, editor, item, blockId: start.blockId, offset: start.offset,
      received: 0, seen: 0, samples: [], frame: null, invalid: false, observer: null };
    current.observer = new MutationObserver(() => inspect(current));
    current.observer.observe(editor.container, { subtree: true, childList: true, characterData: true });
    burst = current;
    mark('probe.armed', { documentId: item.id, phase });
    return { documentId: item.id, openingId: item.openingId, inputCount: text.length };
  }
  async function finishBurst() {
    const current = burst;
    if (!current) throw new Error('No armed burst');
    try {
      const model = textOf(await callProbe(current.item, 'getNodeById', { nodeId: current.blockId }), current.blockId);
      const selection = await callProbe(current.item, 'selection.current');
      const segment = selection?.target?.segments?.[0];
      const valid = !current.invalid && sameDocument(current) && current.received === current.text.length &&
        current.samples.every(sample => sample.renderOpportunityAt !== null) &&
        model === expectedText(current, current.text.length) && selection?.empty === true &&
        bodyTarget(selection?.target) && bodyTarget(selection?.selectionTarget) &&
        selection?.target?.segments?.length === 1 && segment?.blockId === current.blockId &&
        segment.range?.start === current.offset + current.text.length && segment.range?.end === segment.range?.start;
      for (const sample of current.samples) sample.validated = valid;
      if (valid && !firstEdit && current.samples.length) {
        const first = current.samples[0];
        firstEdit = { documentId: current.item.id, openingId: current.item.openingId,
          inputAt: first.downAt, visibleDomAt: first.visibleDomAt,
          renderOpportunityAt: first.renderOpportunityAt, documentApiValidated: true };
        mark('first-edit.validated', { documentId: current.item.id });
      }
      return { valid, received: current.received, expected: current.text.length,
        displayed: current.samples.filter(sample => sample.renderOpportunityAt !== null).length };
    } finally { clearBurst(); }
  }
  const recordLong = list => {
    for (const entry of list) push(longTasks, { start: entry.startTime, duration: entry.duration });
  };
  try {
    longObserver = new PerformanceObserver(list => recordLong(list.getEntries()));
    longObserver.observe({ type: 'longtask', buffered: true });
  } catch { longObserver = null; }
  const visibility = () => { if (burst) burst.invalid = true; mark('visibility', { visible: document.visibilityState === 'visible' }); };
  document.addEventListener('keydown', keydown, true);
  document.addEventListener('visibilitychange', visibility);
  function snapshot() {
    if (longObserver) recordLong(longObserver.takeRecords());
    return JSON.parse(JSON.stringify({ schemaVersion: 1, boundary,
      hostClickIncluded: false, inputSource: 'CDP keyboard', clock: 'performance.now',
      timeOrigin: performance.timeOrigin, firstEdit, events, keys, longTasks,
      longTasksSupported: longObserver !== null, api: [...api.values()], apiCoverage: coverage,
      apiBytes: apiBytes ? 'JSON UTF-8 estimate; not transport/wire bytes' : 'disabled',
      renderMetric: 'visible target text followed by two animation frames; not native pixel presentation',
      dropped, stopped }));
  }
  function stop() {
    if (stopped) return;
    clearBurst();
    if (longObserver) { recordLong(longObserver.takeRecords()); longObserver.disconnect(); }
    stopped = true;
    document.removeEventListener('keydown', keydown, true);
    document.removeEventListener('visibilitychange', visibility);
    for (const dispose of disposers) dispose();
    if (window.__otzariaPerformance === recorder) delete window.__otzariaPerformance;
  }
  const recorder = { ownerId, mark, begin, api: measureApi, bind, prepare, finishBurst, snapshot, stop,
    pending: () => burst ? { received: burst.received, displayed: burst.samples.filter(x => x.renderOpportunityAt !== null).length } : null };
  window.__otzariaPerformance = recorder;
  mark('recorder.installed');
  return recorder;
}

/** Validate the earliest selected intent and opening, never a later fallback. */
export function validateStartupOpening(recording, openingId, expectedKind = null) {
  if (recording.boundary === 'late-attach') return false;
  const events = recording.events;
  const intentAt = events.findIndex(event => event.name === 'document.intent.start');
  const openingAt = events.findIndex(event => event.name === 'document.open.start');
  const intent = events[intentAt], opening = events[openingAt];
  if (!intent || !opening || !Number.isSafeInteger(intent.span) || !Number.isSafeInteger(opening.span)) {
    throw new Error('Build lacks startup instrumentation');
  }
  const intentEndAt = events.findIndex(event => event.name === 'document.intent.end' && event.span === intent.span);
  const openingEndAt = events.findIndex(event => event.name === 'document.open.end' && event.span === opening.span);
  const sourceUnavailable = events.slice(intentAt + 1, intentEndAt)
    .some(event => event.name === 'document.intent.source-unavailable');
  if (openingId !== opening.span || intentAt >= openingAt || openingEndAt <= openingAt || intentEndAt <= openingEndAt ||
      events[intentEndAt]?.status !== 'opened' || events[openingEndAt]?.status !== 'opened' || sourceUnavailable ||
      (expectedKind !== null && opening.kind !== expectedKind)) {
    throw new Error('The intended startup opening did not become the measured document');
  }
  return true;
}
