export interface ClassicModules {
  load(id: string): Promise<unknown>;
  /** Fetch/register only; dependency and factory evaluation still waits for load. */
  preload(id: string): Promise<void>;
}

/**
 * The small AMD subset emitted by Rollup, loaded with ordinary script elements.
 * This function is serialized into index.html: keep every runtime dependency
 * inside it. Named modules and an allowlist avoid a global RequireJS/define API.
 * The build rejects static chunk cycles; dynamic imports may return to any
 * already evaluated chunk. Failed loads/factories are retryable.
 */
export function installClassicModules(files: readonly string[]): ClassicModules {
  type Factory = (...dependencies: unknown[]) => unknown;
  interface Entry {
    file: string;
    exports: Record<string, unknown>;
    definition?: { dependencies: string[]; factory: Factory };
    script?: HTMLScriptElement;
    fetching?: Promise<void>;
    evaluating?: Promise<unknown>;
    evaluated?: { value: unknown };
  }

  const base = new URL('.', document.baseURI);
  const entries = new Map<string, Entry>();
  for (const file of files) entries.set(file.replace(/\.js$/, ''), { file, exports: {} });

  function entryFor(id: string): Entry {
    const entry = entries.get(id);
    if (!entry) throw new Error(`Unknown local module: ${id}`);
    return entry;
  }

  function resolve(request: string, parent: string): string {
    let id = request;
    if (request.startsWith('.')) {
      const parts = parent.split('/');
      parts.pop();
      for (const part of request.split('/')) {
        if (part === '..') parts.pop();
        else if (part && part !== '.') parts.push(part);
      }
      id = parts.join('/');
    }
    id = id.replace(/\.js$/, '');
    entryFor(id);
    return id;
  }

  const target = window as Window & {
    __otzariaDefine?: (id: string, dependencies: string[] | Factory, factory?: Factory) => void;
  };
  target.__otzariaDefine = (id, dependencies, factory) => {
    const entry = entryFor(id);
    // A timed-out/removed script must not replace a newer attempt's definition.
    if (document.currentScript !== entry.script) throw new Error(`Unexpected module script: ${id}`);
    const body = typeof dependencies === 'function' ? dependencies : factory;
    const imports = typeof dependencies === 'function' ? [] : dependencies;
    if (!Array.isArray(imports) || typeof body !== 'function' || entry.definition) {
      throw new Error(`Invalid module definition: ${id}`);
    }
    entry.definition = { dependencies: imports, factory: body };
  };

  function fetchDefinition(id: string, entry: Entry): Promise<void> {
    if (entry.definition) return Promise.resolve();
    if (entry.fetching) return entry.fetching;
    // Match the boot probe's physical entry-script events. A module can remain
    // pending on dependencies/factories after this script has registered.
    function measured(event: string): void {
      if (id !== 'assets/app') return;
      try { target.__otzariaPerformance?.mark(event, { asset: 'app.js' }); }
      catch { /* Diagnostics must not change loading or retry behavior. */ }
    }
    entry.fetching = new Promise<void>((accept, reject) => {
      const script = document.createElement('script');
      entry.script = script;
      // Preserve the initial worker-sources → app order, while fetching both.
      script.async = false;
      script.src = new URL(entry.file, base).href;
      let finished = false;
      const timeout = setTimeout(() => finish(new Error(`Module load timed out: ${id}`)), 30_000);
      function finish(error?: Error): void {
        if (finished) return;
        finished = true;
        clearTimeout(timeout);
        script.onload = script.onerror = null;
        if (error) {
          entry.script = undefined;
          entry.definition = undefined;
          script.remove();
          measured('script.failed');
          reject(error);
        } else {
          performance.mark(`otzaria:script:${id}:loaded`);
          measured('script.loaded');
          accept();
        }
      }
      script.onload = () => finish(entry.definition ? undefined : new Error(`Module did not register: ${id}`));
      script.onerror = () => finish(new Error(`Cannot load local module: ${id}`));
      performance.mark(`otzaria:script:${id}:start`);
      measured('script.requested');
      document.head.appendChild(script);
    }).catch((error: unknown) => {
      entry.fetching = undefined;
      throw error;
    });
    return entry.fetching;
  }

  function load(id: string): Promise<unknown> {
    let entry: Entry;
    try { entry = entryFor(id); } catch (error) { return Promise.reject(error); }
    if (entry.evaluating) return entry.evaluating;
    entry.evaluating = fetchDefinition(id, entry).then(async () => {
      const definition = entry.definition!;
      const module = { id, uri: new URL(entry.file, base).href, exports: entry.exports };
      const require = (requests: string | string[], accept?: unknown, reject?: unknown): unknown => {
        if (typeof requests === 'string') {
          // Browserify/JSZip probes require('stream', true) inside try/catch.
          // A synchronous request must either return a completed local module
          // or throw here. Starting an async load would escape that catch.
          const dependency = entryFor(resolve(requests, id));
          if (!dependency.evaluated) throw new Error(`Module is not ready for synchronous require: ${requests}`);
          return dependency.evaluated.value;
        }
        if (!Array.isArray(requests) || typeof accept !== 'function' || typeof reject !== 'function') {
          throw new Error('Invalid asynchronous module request');
        }
        // A rejected dependency reaches Rollup's dynamic-import promise; it is
        // not left as an unhandled rejection or cached as a permanent failure.
        void Promise.resolve().then(() => Promise.all(requests.map((request) => load(resolve(request, id)))))
          .then((values) => accept(...values), (error: unknown) => reject(error));
      };
      const values = await Promise.all(definition.dependencies.map((request) => {
        if (request === 'exports') return entry.exports;
        if (request === 'module') return module;
        if (request === 'require') return require;
        return load(resolve(request, id));
      }));
      const result = await definition.factory(...values);
      const value = result === undefined ? module.exports : result;
      entry.evaluated = { value };
      performance.mark(`otzaria:module:${id}:ready`);
      return value;
    }).catch((error: unknown) => {
      entry.evaluating = undefined;
      entry.evaluated = undefined;
      entry.exports = {};
      throw error;
    });
    return entry.evaluating;
  }

  function preload(id: string): Promise<void> {
    try { return fetchDefinition(id, entryFor(id)); }
    catch (error) { return Promise.reject(error); }
  }

  return { load, preload };
}
