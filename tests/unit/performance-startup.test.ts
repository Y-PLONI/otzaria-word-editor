import { readFileSync } from 'node:fs';
import { createContext, runInContext } from 'node:vm';
import ts from 'typescript';
import { describe, expect, it } from 'vitest';
import { installPerformanceRecorder, validateStartupOpening } from '../../scripts/performance-recorder.mjs';
import { beginPerformanceSpan, markPerformance } from '../../src/diagnostics/performance';
import { activeEntry, decideDraftRecovery, documentViewFor, draftPathFor } from '../../src/sessions/session-state';

// Execute the real coordinator/restoration functions, without mounting the
// unrelated App UI or loading SuperDoc. Only host IO and editor creation are
// controlled doubles; source decisions, ordering and diagnostics stay real.
function sourceFunctions(path: string, names: string[]): string {
  const source = readFileSync(new URL(path, import.meta.url), 'utf8');
  const script = source.includes('<script setup')
    ? source.split(/<script setup[^>]*>/)[1]!.split('</script>')[0]!
    : source;
  const ast = ts.createSourceFile(path, script, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
  return names.map(name => {
    const declaration = ast.statements.find(node => ts.isFunctionDeclaration(node) && node.name?.text === name);
    if (!declaration) throw new Error(`Missing startup function: ${name}`);
    return declaration.getText(ast).replace(/^export\s+/, '');
  }).join('\n');
}

const source = sourceFunctions('../../src/App.vue', [
  'openStartupDocument', 'openPendingTab', 'reopenPreviousSession',
  'resolveRememberedFile', 'recoverDraft', 'readDraftBytes',
]) + '\n' + sourceFunctions('../../src/host/files.ts', ['resolveFileUrl']);
const executable = ts.transpileModule(source, {
  compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS },
}).outputText;

interface Scenario {
  name: string;
  file?: boolean;
  missingFile?: boolean;
  draft?: boolean;
  missingDraft?: boolean;
  useDisk?: boolean;
  failFirstOpening?: boolean;
  queue?: 'failure' | 'failure-then-success' | 'pending' | 'success';
  valid: boolean;
  kinds: string[];
  intentStatuses: string[];
  draftReads?: number;
}

async function replay(scenario: Scenario, enabled: boolean) {
  const recorder = enabled ? installPerformanceRecorder() : undefined;
  const effects: string[] = [];
  const intentBeforeWork: boolean[] = [];
  const entry = {
    id: 'PRIVATE_TAB',
    document: scenario.file ? { token: 'PRIVATE_TOKEN', name: 'PRIVATE_NAME', writable: true } : null,
    draft: scenario.draft ? {
      path: 'PRIVATE_PATH', documentToken: scenario.file ? 'PRIVATE_TOKEN' : null, sourceSize: 5,
    } : null,
    previousDraft: null, caret: null,
  };
  const restored = {
    id: entry.id, pendingRestore: true, slept: false, ui: { scroll: null },
    keeper: { state: { version: 2, activeId: entry.id, documents: [entry], view: { zoom: 100 } } },
    swap: { current: null as object | null }, destroy: async () => { effects.push('destroy'); },
  };
  const activeSession = { value: restored };
  const statusText = { value: '' }, isStatusError = { value: false };
  const requests: Array<() => Promise<boolean>> = [];
  const kinds: string[] = [];
  let openingId: number | null = null;
  const observeIntent = () => {
    const events = recorder?.snapshot().events ?? [];
    intentBeforeWork.push(events.some(event => event.name === 'document.intent.start' &&
      !events.some(end => end.name === 'document.intent.end' && end.span === event.span)));
  };
  const openDocument = async (file?: object, options: { draft?: Blob } = {}) => {
    const kind = options.draft ? 'draft' : file ? 'file' : 'new';
    kinds.push(kind); effects.push('open:' + kind);
    const opening = beginPerformanceSpan('document.open', { kind });
    const events = recorder?.snapshot().events ?? [];
    const id = events[events.length - 1]?.span;
    const opened = !scenario.failFirstOpening || kinds.length > 1;
    if (opened) {
      const editor = { superdoc: { activeEditor: { doc: { get: () => null } } }, onDispose() {} };
      restored.swap.current = editor;
      opening?.bind(editor);
      openingId = typeof id === 'number' ? id : null;
    }
    opening?.end(opened ? 'opened' : 'cancelled');
    return opened;
  };
  const context = createContext({
    Blob, console, activeEntry, decideDraftRecovery, documentViewFor, draftPathFor,
    beginPerformanceSpan, markPerformance, restored, activeSession, statusText, isStatusError,
    sessions: new Map([[restored.id, restored]]), recentTabs: [],
    collectingStartupRequests: true, startupOpenRequests: requests, startupSendEvents: [],
    startupAttemptedSessions: new Set(), pendingTabLoad: false,
    DOCX_MIME: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    documentId: entry.id,
    activateTab(session: typeof restored) { activeSession.value = session; effects.push('activate'); },
    restorePaneScroll() { effects.push('scroll'); },
    setStatus(message: string, error = false) {
      statusText.value = message; isStatusError.value = error; effects.push(error ? 'status:error' : 'status:ok');
    },
    restoredDraftMessage: () => 'restored',
    persistCombinedSession: async () => { effects.push('persist'); },
    runSendToDocument: async () => { effects.push('send'); },
    call: async (name: string) => {
      expect(name).toBe('fs.resolveFileUrl'); observeIntent(); effects.push('resolve');
      if (scenario.missingFile) throw new Error('host resolution failed');
      return { url: 'PRIVATE_URL', name: 'PRIVATE_NAME', size: scenario.useDisk ? 6 : 5 };
    },
    readWorkspaceBytes: async () => {
      observeIntent(); effects.push('draft.read');
      return scenario.missingDraft ? null : new Uint8Array([1]);
    },
    confirm: async () => { effects.push('confirm'); return !scenario.useDisk; },
    openDocument,
  });
  runInContext(executable, context);
  if (scenario.queue === 'failure' || scenario.queue === 'failure-then-success') {
    requests.push(async () => { observeIntent(); effects.push('request:failed'); return false; });
  }
  if (scenario.queue === 'success' || scenario.queue === 'failure-then-success') {
    requests.push(async () => {
      observeIntent(); effects.push('request:open');
      const opened = await openDocument({});
      if (opened) restored.pendingRestore = false;
      return opened;
    });
  }
  if (scenario.queue === 'pending') {
    requests.push(async () => {
      observeIntent(); effects.push('request:pending');
      return runInContext('openPendingTab(restored, { allowBlankFallback: false })', context);
    });
  }
  try {
    await runInContext('openStartupDocument(restored)', context);
    return {
      effects, kinds, openingId, intentBeforeWork, pendingRestore: restored.pendingRestore,
      pendingTabLoad: context.pendingTabLoad, recording: recorder?.snapshot(),
    };
  } finally {
    recorder?.stop();
  }
}

const scenarios: Scenario[] = [
  { name: 'genuine empty startup', valid: true, kinds: ['new'], intentStatuses: ['opened'] },
  { name: 'intended file restoration', file: true, valid: true, kinds: ['file'], intentStatuses: ['opened'] },
  { name: 'missing remembered token followed by blank fallback', file: true, missingFile: true,
    valid: false, kinds: ['new'], intentStatuses: ['failed'] },
  { name: 'missing token with successful intended orphan draft', file: true, missingFile: true, draft: true,
    valid: true, kinds: ['draft'], intentStatuses: ['opened'], draftReads: 1 },
  { name: 'selected draft unavailable followed by disk fallback', file: true, draft: true, missingDraft: true,
    valid: false, kinds: ['file'], intentStatuses: ['opened'], draftReads: 1 },
  { name: 'selected unsaved draft unavailable followed by blank fallback', draft: true, missingDraft: true,
    valid: false, kinds: ['new'], intentStatuses: ['opened'], draftReads: 1 },
  { name: 'intentional choice of disk file instead of draft', file: true, draft: true, useDisk: true,
    valid: true, kinds: ['file'], intentStatuses: ['opened'] },
  { name: 'failed queued request before successful restoration', queue: 'failure', file: true,
    valid: false, kinds: ['file'], intentStatuses: ['failed', 'opened'] },
  { name: 'failed queued request before a later successful request', queue: 'failure-then-success',
    valid: false, kinds: ['file'], intentStatuses: ['failed', 'opened'] },
  { name: 'successful explicit queued request', queue: 'success',
    valid: true, kinds: ['file'], intentStatuses: ['opened'] },
  { name: 'explicit request restoring a pending tab under one intent', queue: 'pending', file: true,
    valid: true, kinds: ['file'], intentStatuses: ['opened'] },
  { name: 'failed explicit pending tab followed by blank without retry', queue: 'pending', file: true, missingFile: true,
    valid: false, kinds: ['new'], intentStatuses: ['failed', 'failed'] },
  { name: 'failed intended file opening followed by blank fallback', file: true, failFirstOpening: true,
    valid: false, kinds: ['file', 'new'], intentStatuses: ['failed'] },
];

describe('startup intent through the actual App coordinator', () => {
  it.each(scenarios)('$name', async scenario => {
    const disabled = await replay(scenario, false);
    const measured = await replay(scenario, true);
    // Turning diagnostics on changes no host calls, opening choices, effects,
    // pending-tab state or cleanup in these concrete recovery paths.
    expect(measured.effects).toEqual(disabled.effects);
    expect(measured.pendingRestore).toBe(disabled.pendingRestore);
    expect(measured.pendingTabLoad).toBe(false);
    expect(measured.kinds).toEqual(scenario.kinds);
    expect(measured.effects.filter(effect => effect === 'draft.read')).toHaveLength(scenario.draftReads ?? 0);
    expect(measured.intentBeforeWork.every(Boolean)).toBe(true);
    const recording = measured.recording!;
    expect(recording.events.filter(event => event.name === 'document.intent.end').map(event => event.status))
      .toEqual(scenario.intentStatuses);
    const validate = () => validateStartupOpening(recording, measured.openingId);
    if (scenario.valid) expect(validate()).toBe(true);
    else expect(validate).toThrow('intended startup opening');
    expect(JSON.stringify(recording)).not.toContain('PRIVATE_');
    expect(window.__otzariaPerformance).toBeUndefined();
  });
});
