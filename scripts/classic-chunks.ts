import type { OutputBundle, OutputChunk } from 'rollup';
import type { Plugin } from 'vite';

export function classicChunkFiles(bundle: OutputBundle): string[] {
  return Object.values(bundle).filter((item): item is OutputChunk => item.type === 'chunk').map((item) => item.fileName).sort();
}

interface ClassicChunkOptions {
  /** The editor needs the public package's dynamic engine to open a DOCX. */
  preloadDocumentEngine?: boolean;
}

export function classicStartupPlan(bundle: OutputBundle, { preloadDocumentEngine = false }: ClassicChunkOptions = {}) {
  const visiting = new Set<string>();
  const visited = new Set<string>();
  const walk = (file: string): void => {
    if (visiting.has(file)) throw new Error(`Static classic-script chunk cycle: ${file}`);
    if (visited.has(file)) return;
    const chunk = bundle[file];
    if (!chunk || chunk.type !== 'chunk') throw new Error(`Missing local chunk: ${file}`);
    visiting.add(file);
    chunk.imports.forEach(walk);
    visiting.delete(file);
    visited.add(file);
  };
  walk('assets/app.js');
  const entryFiles = [...visited];
  if (preloadDocumentEngine) {
    // Identify the chunk produced by the public SuperDoc import, without a
    // private source import or a filename/hash that changes between builds.
    const engines = Object.values(bundle).filter((item): item is OutputChunk => item.type === 'chunk' &&
      Object.entries(item.modules).some(([id, module]) => module.renderedLength > 0 &&
        id.replace(/\\/g, '/').endsWith('/node_modules/@superdoc/docx-engine/dist/docx-engine.es.js')));
    if (engines.length !== 1) throw new Error(`Expected one DOCX engine chunk, found ${engines.length}`);
    walk(engines[0]!.fileName);
  }
  const initialFiles = [...visited];
  // Validate every static graph, including chunks that remain fully lazy.
  classicChunkFiles(bundle).forEach(walk);
  return {
    entryFiles,
    preloadFiles: initialFiles.filter((file) => file !== 'assets/app.js'),
    startupFiles: ['assets/engine-workers.js', ...initialFiles],
  };
}

/** Emit a reviewable size/dependency report and enforce the loader's contract. */
export function classicChunks(options: ClassicChunkOptions = {}): Plugin {
  const pageChunks = new Set<string>();
  return {
    name: 'otzaria-classic-chunks',
    apply: 'build',
    config() {
      return { experimental: { renderBuiltUrl(filename, { hostId, hostType }) {
        // Vite inserts asset URLs after Rollup has chosen AMD parameters. Its
        // require.toUrl can therefore refer to an undeclared global `require`.
        // Only this page's chunks use the HTML root. Vite inherits this callback
        // in worker builds too; leave their self.location-based URLs untouched.
        return hostType === 'js' && pageChunks.has(hostId)
          ? { runtime: `new URL(${JSON.stringify(filename)}, document.baseURI).href` }
          : { relative: true };
      } } };
    },
    buildStart() { pageChunks.clear(); },
    renderChunk: {
      order: 'pre',
      handler(_code, chunk) {
        // Runs before Vite's asset render hooks. Main plugins are not installed
        // in worker builds, so their chunks never enter this set.
        pageChunks.add(chunk.fileName);
        return null;
      },
    },
    // PDF.js's public fallback imports an embedded data: worker URL. Keep that
    // native expression (valid in a classic script), as before this split.
    // Resolved imports are emitted as local AMD dependencies by Rollup.
    renderDynamicImport({ targetModuleId }) {
      return targetModuleId === null ? { left: 'import(', right: ')' } : null;
    },
    generateBundle(_options, bundle) {
      const files = classicChunkFiles(bundle);
      const { entryFiles, preloadFiles, startupFiles } = classicStartupPlan(bundle, options);
      const chunks = Object.fromEntries(files.map((file) => {
        const chunk = bundle[file] as OutputChunk;
        // Page/worker asset URLs use renderBuiltUrl above. The minimal AMD
        // loader intentionally omits require.toUrl; reject any return of this
        // helper in final output, even if its receiver has been renamed.
        if (/\.\s*toUrl\b/.test(chunk.code)) {
          this.error(`Unsupported AMD asset URL helper in ${file}: use the Vite page/worker URL paths`);
        }
        const modules = Object.entries(chunk.modules).map(([id, module]) => ({
          id: id.replace(/\\/g, '/').replace(/^.*\/node_modules\//, 'node_modules/').replace(`${process.cwd().replace(/\\/g, '/')}/`, ''),
          renderedLength: module.renderedLength,
        }));
        return [file, {
          bytes: Buffer.byteLength(chunk.code),
          imports: chunk.imports,
          dynamicImports: chunk.dynamicImports,
          engineLicense: chunk.code.includes('DOCX Engine Proprietary License Agreement'),
          iconsLicense: chunk.code.includes('Fluent System Icons — MIT'),
          modules,
        }];
      }));
      this.emitFile({
        type: 'asset',
        fileName: 'bundle-report.json',
        source: JSON.stringify({
          format: 'classic-amd',
          entry: 'assets/app.js',
          entryFiles,
          preloadFiles,
          startupFiles,
          deferredFiles: files.filter((file) => !startupFiles.includes(file)),
          chunks,
        }, null, 2),
      });
    },
  };
}
