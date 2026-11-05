import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { rollup } from 'rollup';
import { installClassicModules } from '../../scripts/classic-modules';
import { classicChunks } from '../../scripts/classic-chunks';

type Factory = (...values: unknown[]) => unknown;
const target = window as Window & {
  __otzariaDefine?: (id: string, dependencies: string[] | Factory, factory?: Factory) => void;
};

function scriptFor(id: string): HTMLScriptElement {
  const script = [...document.scripts].find((item) => item.src.endsWith(`/${id}.js`));
  expect(script, `script ${id}`).toBeDefined();
  return script!;
}

function define(id: string, dependencies: string[], factory: Factory): void {
  const script = scriptFor(id);
  const current = vi.spyOn(document, 'currentScript', 'get').mockReturnValue(script);
  target.__otzariaDefine!(id, dependencies, factory);
  current.mockRestore();
  script.dispatchEvent(new Event('load'));
}

beforeEach(() => {
  document.head.innerHTML = '';
  vi.stubGlobal('performance', { mark: vi.fn(), now: () => Date.now() });
});
afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
  delete target.__otzariaDefine;
});

describe('classic module loader', () => {
  it('serializes without a build-time closure and does not claim global AMD names', async () => {
    const install = new Function(`return (${installClassicModules.toString()})`)() as typeof installClassicModules;
    const previousDefine = (window as unknown as { define?: unknown }).define;
    const modules = install(['assets/app.js']);
    const request = modules.load('assets/app');
    expect(scriptFor('assets/app').type).toBe('');
    expect(scriptFor('assets/app').async).toBe(false);
    define('assets/app', ['exports', 'module'], (exports, module) => {
      Object.assign(exports!, { uri: (module as { uri: string }).uri, value: 17 });
    });
    expect(await request).toEqual({ uri: scriptFor('assets/app').src, value: 17 });
    expect((window as unknown as { define?: unknown }).define).toBe(previousDefine);
  });

  it.each([false, true])('records physical App loading before pending dependency work (diagnostics throw: %s)', async (throws) => {
    const mark = vi.fn(() => { if (throws) throw new Error('diagnostic failure'); });
    vi.stubGlobal('__otzariaPerformance', { mark });
    // Exercise the serialized bootstrap, not a separate module-bound recorder.
    const install = new Function(`return (${installClassicModules.toString()})`)() as typeof installClassicModules;
    const modules = install(['assets/app.js', 'assets/engine.js']);
    const preload = modules.preload('assets/app');
    const app = modules.load('assets/app');
    expect(modules.load('assets/app')).toBe(app);
    expect(mark.mock.calls).toEqual([['script.requested', { asset: 'app.js' }]]);
    const appFactory = vi.fn((engine) => engine);
    define('assets/app', ['./engine'], appFactory);
    await preload;
    expect(mark.mock.calls).toEqual([
      ['script.requested', { asset: 'app.js' }],
      ['script.loaded', { asset: 'app.js' }],
    ]);
    expect(appFactory).not.toHaveBeenCalled();
    expect(performance.mark).not.toHaveBeenCalledWith('otzaria:module:assets/app:ready');

    await vi.waitFor(() => expect(document.scripts).toHaveLength(2));
    let release!: (value: number) => void;
    const engineFactory = vi.fn(() => new Promise<number>((resolve) => { release = resolve; }));
    define('assets/engine', [], engineFactory);
    await vi.waitFor(() => expect(engineFactory).toHaveBeenCalledTimes(1));
    expect(appFactory).not.toHaveBeenCalled();
    expect(performance.mark).not.toHaveBeenCalledWith('otzaria:module:assets/app:ready');
    expect(mark).toHaveBeenCalledTimes(2);
    release(42);
    expect(await app).toBe(42);
    expect(performance.mark).toHaveBeenCalledWith('otzaria:module:assets/app:ready');

    await modules.preload('assets/app');
    expect(await modules.load('assets/app')).toBe(42);
    expect(appFactory).toHaveBeenCalledTimes(1);
    expect(document.scripts).toHaveLength(2);
    expect(mark).toHaveBeenCalledTimes(2);
  });

  it('records a failed physical App attempt and allows retry even when diagnostics throw', async () => {
    vi.useFakeTimers();
    const mark = vi.fn(() => { throw new Error('diagnostic failure'); });
    vi.stubGlobal('__otzariaPerformance', { mark });
    const modules = installClassicModules(['assets/app.js']);
    const first = modules.load('assets/app');
    const rejected = expect(first).rejects.toThrow('Module load timed out: assets/app');
    await vi.advanceTimersByTimeAsync(30_000);
    await rejected;
    expect(document.scripts).toHaveLength(0);
    const retry = modules.load('assets/app');
    define('assets/app', [], () => 23);
    expect(await retry).toBe(23);
    expect(await modules.load('assets/app')).toBe(23);
    expect(mark.mock.calls).toEqual([
      ['script.requested', { asset: 'app.js' }],
      ['script.failed', { asset: 'app.js' }],
      ['script.requested', { asset: 'app.js' }],
      ['script.loaded', { asset: 'app.js' }],
    ]);
    expect(document.scripts).toHaveLength(1);
  });

  it('shares the pending load and evaluated exports across all callers', async () => {
    const modules = installClassicModules(['assets/shared.js']);
    const a = modules.load('assets/shared');
    const b = modules.load('assets/shared');
    expect(a).toBe(b);
    expect(document.scripts).toHaveLength(1);
    const factory = vi.fn(() => ({ answer: 42 }));
    define('assets/shared', [], factory);
    expect(await a).toBe(await b);
    expect(await modules.load('assets/shared')).toEqual({ answer: 42 });
    expect(factory).toHaveBeenCalledTimes(1);
    expect(document.scripts).toHaveLength(1);
  });

  it('fetches a DOCX definition beside the app without evaluating its factory early', async () => {
    const modules = installClassicModules(['assets/app.js', 'assets/engine.js', 'assets/pdf.js']);
    const app = modules.load('assets/app');
    const preload = modules.preload('assets/engine');
    expect([...document.scripts].map((script) => new URL(script.src).pathname.split('/').pop()))
      .toEqual(['app.js', 'engine.js']);
    const engineFactory = vi.fn((value) => ({ value }));
    define('assets/engine', ['./app'], engineFactory);
    await preload;
    expect(engineFactory).not.toHaveBeenCalled();
    define('assets/app', [], () => 17);
    await app;
    expect(engineFactory).not.toHaveBeenCalled();
    expect(await modules.load('assets/engine')).toEqual({ value: 17 });
    expect(engineFactory).toHaveBeenCalledTimes(1);
    expect(document.scripts).toHaveLength(2);
  });

  it('shares an unfinished preload with concurrent preload and load consumers', async () => {
    const modules = installClassicModules(['assets/engine.js']);
    const first = modules.preload('assets/engine');
    const second = modules.preload('assets/engine');
    const consumer = modules.load('assets/engine');
    expect(first).toBe(second);
    expect(document.scripts).toHaveLength(1);
    const factory = vi.fn(() => 42);
    define('assets/engine', [], factory);
    await Promise.all([first, second]);
    expect(await consumer).toBe(42);
    expect(await modules.load('assets/engine')).toBe(42);
    expect(factory).toHaveBeenCalledTimes(1);
    expect(document.scripts).toHaveLength(1);
  });

  it.each(['error', 'missing-definition', 'timeout'] as const)('retries a failed %s preload through the normal consumer', async (failure) => {
    vi.useFakeTimers();
    const modules = installClassicModules(['assets/engine.js']);
    const preload = modules.preload('assets/engine');
    const rejected = expect(preload).rejects.toThrow();
    const failed = scriptFor('assets/engine');
    if (failure === 'timeout') await vi.advanceTimersByTimeAsync(30_000);
    else failed.dispatchEvent(new Event(failure === 'error' ? 'error' : 'load'));
    await rejected;
    expect(document.scripts).toHaveLength(0);
    const consumer = modules.load('assets/engine');
    expect(scriptFor('assets/engine')).not.toBe(failed);
    define('assets/engine', [], () => 23);
    expect(await consumer).toBe(23);
    expect(document.scripts).toHaveLength(1);
  });

  it('loads static dependencies before their consumer and dynamic ones only on demand', async () => {
    const modules = installClassicModules(['assets/app.js', 'assets/shared.js', 'assets/lazy.js']);
    const request = modules.load('assets/app');
    const order: string[] = [];
    define('assets/app', ['require', './shared'], (require, shared) => {
      order.push('app');
      return {
        shared,
        lazy: () => new Promise((accept, reject) => {
          (require as (ids: string[], accept: (value: unknown) => void, reject: (error: unknown) => void) => void)(['./lazy'], accept, reject);
        }),
      };
    });
    await vi.waitFor(() => expect(document.scripts).toHaveLength(2));
    expect(order).toEqual([]);
    define('assets/shared', [], () => { order.push('shared'); return 9; });
    const app = await request as { shared: number; lazy(): Promise<unknown> };
    expect(app.shared).toBe(9);
    expect(order).toEqual(['shared', 'app']);
    expect(document.scripts).toHaveLength(2);
    const lazy = app.lazy();
    await vi.waitFor(() => expect(document.scripts).toHaveLength(3));
    define('assets/lazy', ['./shared'], (shared) => ({ shared }));
    expect(await lazy).toEqual({ shared: 9 });
    expect(await app.lazy()).toEqual({ shared: 9 });
    expect(document.scripts).toHaveLength(3);
  });

  it.each(['error', 'missing-definition', 'timeout'] as const)('allows retry after %s and removes the failed script', async (failure) => {
    vi.useFakeTimers();
    const modules = installClassicModules(['assets/lazy.js']);
    const request = modules.load('assets/lazy');
    const rejected = expect(request).rejects.toThrow();
    const failed = scriptFor('assets/lazy');
    if (failure === 'timeout') await vi.advanceTimersByTimeAsync(30_000);
    else failed.dispatchEvent(new Event(failure === 'error' ? 'error' : 'load'));
    await rejected;
    expect(document.scripts).toHaveLength(0);
    const retry = modules.load('assets/lazy');
    expect(scriptFor('assets/lazy')).not.toBe(failed);
    define('assets/lazy', [], () => 12);
    expect(await retry).toBe(12);
  });

  it('rejects an old timed-out script without corrupting its replacement', async () => {
    vi.useFakeTimers();
    const modules = installClassicModules(['assets/lazy.js']);
    const first = modules.load('assets/lazy');
    const rejected = expect(first).rejects.toThrow('timed out');
    const old = scriptFor('assets/lazy');
    await vi.advanceTimersByTimeAsync(30_000);
    await rejected;
    const retry = modules.load('assets/lazy');
    const current = vi.spyOn(document, 'currentScript', 'get').mockReturnValue(old);
    expect(() => target.__otzariaDefine!('assets/lazy', [], () => 'old')).toThrow('Unexpected module script');
    current.mockRestore();
    define('assets/lazy', [], () => 'new');
    expect(await retry).toBe('new');
  });

  it('retries a failed dependency and a failed factory without refetching successful code', async () => {
    const modules = installClassicModules(['assets/app.js', 'assets/dependency.js']);
    const initial = modules.load('assets/app');
    const rejected = expect(initial).rejects.toThrow();
    let attempts = 0;
    define('assets/app', ['./dependency'], (value) => {
      if (++attempts === 1) throw new Error('transient initialization');
      return value;
    });
    await vi.waitFor(() => expect(document.scripts).toHaveLength(2));
    scriptFor('assets/dependency').dispatchEvent(new Event('error'));
    await rejected;
    const retry = modules.load('assets/app');
    const secondFailure = expect(retry).rejects.toThrow('transient initialization');
    await vi.waitFor(() => expect(document.scripts).toHaveLength(2));
    define('assets/dependency', [], () => 23);
    await secondFailure;
    expect(await modules.load('assets/app')).toBe(23);
    expect(document.scripts).toHaveLength(2);
    expect(attempts).toBe(2);
  });

  it('refuses files outside the build manifest without injecting a script', async () => {
    const modules = installClassicModules(['assets/app.js']);
    await expect(modules.load('https://example.com/code')).rejects.toThrow('Unknown local module');
    await expect(modules.load('../foreign')).rejects.toThrow('Unknown local module');
    await expect(modules.preload('https://example.com/code')).rejects.toThrow('Unknown local module');
    await expect(modules.preload('../foreign')).rejects.toThrow('Unknown local module');
    expect(document.scripts).toHaveLength(0);
  });

  it('lets the real JSZip/Browserify stream probe fail inside its synchronous catch', async () => {
    const modules = installClassicModules(['assets/app.js']);
    const request = modules.load('assets/app');
    let caught = false;
    define('assets/app', ['require'], (require) => {
      // Browserify forwards the missing module with its second `true` flag;
      // JSZip catches that lookup to disable its optional Node stream support.
      let stream: unknown = 'not returned';
      try { stream = (require as (id: string, jumped: boolean) => unknown)('stream', true); }
      catch { caught = true; }
      return { nodeStream: caught ? false : stream };
    });
    expect(await request).toEqual({ nodeStream: false });
    expect(caught).toBe(true);
    expect(document.scripts).toHaveLength(1);
  });

  it('returns only an evaluated value for synchronous require and never fetches for it', async () => {
    const modules = installClassicModules(['assets/app.js', 'assets/optional.js']);
    const request = modules.load('assets/app');
    define('assets/app', ['require'], (require) => require);
    const require = await request as (id: string) => unknown;
    expect(() => require('unknown')).toThrow('Unknown local module');
    expect(() => require('./optional')).toThrow('not ready for synchronous require');
    expect(document.scripts).toHaveLength(1);
    const optional = modules.load('assets/optional');
    expect(() => require('./optional')).toThrow('not ready for synchronous require');
    const value = { ready: true };
    define('assets/optional', [], () => value);
    // Registration is not evaluation, even after the script's load event.
    expect(() => require('./optional')).toThrow('not ready for synchronous require');
    await optional;
    expect(require('./optional')).toBe(value);
    expect(document.scripts).toHaveLength(2);
  });

  it('runs actual Rollup AMD chunks, including lazy code and module.uri', async () => {
    const sources: Record<string, string> = {
      entry: `export const uri = import.meta.url; export const later = () => import('optional');`,
      optional: `export const answer = 42;`,
    };
    const build = await rollup({ input: 'entry', plugins: [{
      name: 'fixture', resolveId: (id) => id,
      load: (id) => sources[id],
    }, classicChunks()] });
    try {
      const { output } = await build.generate({
        format: 'amd', amd: { autoId: true, define: 'window.__otzariaDefine' }, exports: 'named',
        entryFileNames: 'assets/app.js', chunkFileNames: 'assets/[name].js',
      });
      const chunks = output.filter((item) => item.type === 'chunk');
      const modules = installClassicModules(chunks.map((item) => item.fileName));
      const run = (file: string) => {
        const script = scriptFor(file.replace(/\.js$/, ''));
        const current = vi.spyOn(document, 'currentScript', 'get').mockReturnValue(script);
        new Function(chunks.find((chunk) => chunk.fileName === file)!.code)();
        current.mockRestore();
        script.dispatchEvent(new Event('load'));
      };
      const first = modules.load('assets/app');
      run('assets/app.js');
      const app = await first as { uri: string; later(): Promise<{ answer: number }> };
      expect(app.uri).toBe(scriptFor('assets/app').src);
      expect(document.scripts).toHaveLength(1);
      const deferred = app.later();
      await vi.waitFor(() => expect(document.scripts).toHaveLength(2));
      run('assets/optional.js');
      expect(await deferred).toEqual({ answer: 42 });
      expect((await app.later()).answer).toBe(42);
    } finally { await build.close(); }
  });

  it.each(['require', 'r'])('rejects unsupported Rollup file URLs with AMD receiver %s', async (receiver) => {
    let assetRef = '';
    const build = await rollup({ input: 'entry', plugins: [{
      name: 'file-url-fixture', resolveId: (id) => id,
      buildStart() { assetRef = this.emitFile({ type: 'asset', fileName: 'assets/fixture.txt', source: 'fixture' }); },
      load: () => `export const asset = import.meta.ROLLUP_FILE_URL_${assetRef};`,
      // Rename the actual emitted parameter/call, preserving the 'require'
      // dependency ID: the build contract must not depend on an identifier.
      renderChunk(code) { return { code: code.replace(/\brequire(?=\s*[.,)])/g, receiver), map: null }; },
    }, classicChunks()] });
    try {
      await expect(build.generate({ format: 'amd', entryFileNames: 'assets/app.js' }))
        .rejects.toThrow('Unsupported AMD asset URL helper');
    } finally { await build.close(); }
  });

  it('builds the DOCX preload closure from module metadata while keeping PDF and panels lazy', async () => {
    const engineId = '/fixture/node_modules/@superdoc/docx-engine/dist/docx-engine.es.js';
    const sources: Record<string, string> = {
      entry: `export const value = 'app'; export const engine = () => import('engine'); export const pdf = () => import('pdf'); export const panel = () => import('panel');`,
      [engineId]: `import { value } from 'entry'; import { tag } from 'helper'; export const result = value + tag;`,
      helper: `export const tag = 'engine';`,
      pdf: `export const pdf = 1;`,
      panel: `export const panel = 2;`,
    };
    const build = await rollup({ input: 'entry', plugins: [{
      name: 'fixture', resolveId: (id) => id === 'engine' ? engineId : id, load: (id) => sources[id],
    }, classicChunks({ preloadDocumentEngine: true })] });
    try {
      const { output } = await build.generate({
        format: 'amd', amd: { autoId: true }, exports: 'named', entryFileNames: 'assets/app.js',
        chunkFileNames: 'assets/[name]-[hash].js', manualChunks: { 'engine-helper': ['helper'] },
      });
      const chunks = output.filter((item) => item.type === 'chunk');
      const fileFor = (id: string) => chunks.find((chunk) => Object.keys(chunk.modules).includes(id))!.fileName;
      const report = output.find((item) => item.type === 'asset' && item.fileName === 'bundle-report.json');
      if (!report || report.type !== 'asset') throw new Error('Missing bundle report');
      const plan = JSON.parse(String(report.source));
      expect(plan.entryFiles).toEqual(['assets/app.js']);
      expect(new Set(plan.preloadFiles)).toEqual(new Set([fileFor('helper'), fileFor(engineId)]));
      expect(new Set(plan.startupFiles)).toEqual(new Set(['assets/engine-workers.js', 'assets/app.js', ...plan.preloadFiles]));
      expect(plan.deferredFiles).toContain(fileFor('pdf'));
      expect(plan.deferredFiles).toContain(fileFor('panel'));
      expect(plan.deferredFiles).not.toContain(fileFor(engineId));
    } finally { await build.close(); }
  });

  it('fails the production build if the required DOCX engine metadata is absent', async () => {
    const build = await rollup({ input: 'entry', plugins: [{
      name: 'fixture', resolveId: (id) => id, load: () => 'export const app = 1;',
    }, classicChunks({ preloadDocumentEngine: true })] });
    try {
      await expect(build.generate({ format: 'amd', entryFileNames: 'assets/app.js' }))
        .rejects.toThrow('Expected one DOCX engine chunk, found 0');
    } finally { await build.close(); }
  });

  it('preserves PDF.js native import of an embedded data URL', async () => {
    const build = await rollup({ input: 'entry', plugins: [{
      name: 'fixture', resolveId: (id) => id,
      load: () => `export const worker = (url) => import(url);`,
    }, classicChunks()] });
    try {
      const { output } = await build.generate({
        format: 'amd', amd: { autoId: true, define: 'window.__otzariaDefine' }, exports: 'named', entryFileNames: 'assets/app.js',
      });
      const code = output.find((item) => item.type === 'chunk')!;
      expect(code.type).toBe('chunk');
      if (code.type === 'chunk') {
        expect(code.code).toContain('import(url)');
        expect(code.code).not.toContain("require([url]");
      }
    } finally { await build.close(); }
  });

  it('rejects static chunk cycles at build time instead of leaving the loader waiting forever', async () => {
    const sources: Record<string, string> = {
      entry: `export { a } from 'a';`,
      a: `import { b } from 'b'; export function a() { return b; }`,
      b: `import { a } from 'a'; export function b() { return a; }`,
    };
    const build = await rollup({ input: 'entry', onwarn() {}, plugins: [{
      name: 'fixture', resolveId: (id) => id, load: (id) => sources[id],
    }, classicChunks()] });
    try {
      await expect(build.generate({
        format: 'amd', amd: { autoId: true }, entryFileNames: 'assets/app.js',
        chunkFileNames: 'assets/[name].js', manualChunks: { a: ['a'], b: ['b'] },
      })).rejects.toThrow('Static classic-script chunk cycle');
    } finally { await build.close(); }
  });

});
