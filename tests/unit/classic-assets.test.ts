// @vitest-environment node
// Keep Vite/esbuild in Node's realm; the DOM loader is tested separately.
import { expect, it, vi } from 'vitest';
import { build as viteBuild } from 'vite';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { runInNewContext } from 'node:vm';
import { classicChunks } from '../../scripts/classic-chunks';

it('runs Vite page assets without a global require and preserves worker asset URLs', async () => {
  const root = mkdtempSync(join(tmpdir(), 'otzaria-classic-assets-'));
  try {
    writeFileSync(join(root, 'asset.txt'), 'local asset');
    writeFileSync(join(root, 'worker-asset.txt'), 'worker asset');
    writeFileSync(join(root, 'worker.js'), 'import url from "./worker-asset.txt?url"; self.postMessage(url);');
    writeFileSync(join(root, 'entry.js'), 'import url from "./asset.txt?url"; window.__assetTestUrl = url; window.__assetTestWorker = () => new Worker(new URL("./worker.js", import.meta.url));');
    const built = await viteBuild({
      configFile: false, root, logLevel: 'silent', base: './',
      plugins: [classicChunks()],
      worker: { format: 'iife' },
      build: { write: false, minify: false, assetsInlineLimit: 0, modulePreload: false, rollupOptions: {
        input: join(root, 'entry.js'),
        output: { format: 'amd', amd: { autoId: true, define: 'window.__otzariaDefine' }, entryFileNames: 'assets/app.js', assetFileNames: 'assets/[name][extname]' },
      } },
    });
    if (Array.isArray(built) || !('output' in built)) throw new Error('Expected one Rollup output');
    const chunk = built.output.find((item) => item.type === 'chunk');
    if (!chunk || chunk.type !== 'chunk') throw new Error('Missing fixture chunk');
    expect(chunk.code).not.toContain('require.toUrl');
    const baseURI = 'file:///plugin/index.html';
    type Factory = (...values: unknown[]) => unknown;
    const target: {
      __assetTestUrl?: string;
      __assetTestWorker?: () => void;
      __otzariaDefine?: (id: string, dependencies: string[] | Factory, factory?: Factory) => void;
    } = {};
    const exports = {};
    const module = { uri: new URL(chunk.fileName, baseURI).href, exports };
    target.__otzariaDefine = (id, dependencies, factory) => {
      expect(id).toBe('assets/app');
      if (typeof dependencies === 'function') { dependencies(); return; }
      factory!(...dependencies.map((dependency) => {
        if (dependency === 'module') return module;
        if (dependency === 'exports') return exports;
        throw new Error(`Unexpected fixture dependency: ${dependency}`);
      }));
    };
    const Worker = vi.fn(function (_url: URL | string) {});
    runInNewContext(chunk.code, { window: target, document: { baseURI }, URL, Worker });
    expect(target.__assetTestUrl).toBe(new URL('assets/asset.txt', baseURI).href);

    const worker = built.output.find((item) => item.type === 'asset' && /\/worker-[^.]+\.js$/.test(item.fileName));
    if (!worker || worker.type !== 'asset') throw new Error('Missing fixture worker');
    target.__assetTestWorker!();
    const workerUri = new URL(worker.fileName, baseURI).href;
    expect(String(Worker.mock.calls[0]?.[0])).toBe(workerUri);
    const workerCode = String(worker.source);
    expect(workerCode).not.toContain('document.baseURI');
    expect(workerCode).toContain('self.location.href');
    const postMessage = vi.fn();
    new Function('self', workerCode)({ location: { href: workerUri }, postMessage });
    const asset = built.output.find((item) => item.type === 'asset' && /\/worker-asset(?:-[^.]+)?\.txt$/.test(item.fileName));
    expect(asset).toBeDefined();
    expect(postMessage).toHaveBeenCalledWith(new URL(asset!.fileName, baseURI).href);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
