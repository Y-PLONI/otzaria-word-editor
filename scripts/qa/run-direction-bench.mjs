/**
 * Measure long paragraphs with neutral runs, which previously did a quadratic
 * search for each run's nearest directional neighbour. No machine-dependent
 * timing gate: the fixture checks content/mark counts and prints scaling.
 *
 * node scripts/qa/run-direction-bench.mjs [maximum runs=64000] [repeats=3]
 */
import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';
import { performance } from 'node:perf_hooks';

const maximum = Number(process.argv[2] ?? 64_000);
const repeats = Number(process.argv[3] ?? 3);
if (!Number.isSafeInteger(maximum) || maximum < 8 || !Number.isSafeInteger(repeats) || repeats < 1) {
  throw new Error('Expected maximum runs >= 8 and repeats >= 1');
}
const bundled = await build({
  stdin: {
    contents: "export { markRtlRuns } from './src/engine/docx-run-direction'; export { readStyleSheet } from './src/engine/docx-style-sheet';",
    resolveDir: fileURLToPath(new URL('../../', import.meta.url)),
    loader: 'ts',
  },
  bundle: true, format: 'esm', platform: 'neutral', write: false,
});
const { markRtlRuns, readStyleSheet } = await import(`data:text/javascript;base64,${Buffer.from(bundled.outputFiles[0].text).toString('base64')}`);
const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const sheet = readStyleSheet(`<w:styles ${ns}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:cs="David"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`);
const run = '<w:r><w:t>1</w:t></w:r>';
for (const placement of ['before', 'after']) {
  for (const count of [maximum / 8, maximum / 4, maximum / 2, maximum].map(Math.floor)) {
    const strong = '<w:r><w:t>א</w:t></w:r>';
    const text = placement === 'before' ? run.repeat(count) + strong : strong + run.repeat(count);
    const xml = `<w:document ${ns}><w:body><w:p><w:pPr><w:bidi/></w:pPr>${text}</w:p></w:body></w:document>`;
    const timings = [];
    for (let i = 0; i < repeats; i += 1) {
      const start = performance.now();
      const output = markRtlRuns(xml, sheet);
      timings.push(performance.now() - start);
      if ((output?.match(/<w:rtl\/>/g) ?? []).length !== count + 1 ||
          (output?.match(/<w:t>1<\/w:t>/g) ?? []).length !== count) {
        throw new Error('Direction marking lost text or omitted neutral runs');
      }
    }
    timings.sort((a, b) => a - b);
    console.log(JSON.stringify({ placement, runs: count, bytes: Buffer.byteLength(xml), minMs: +timings[0].toFixed(1), medianMs: +timings[Math.floor(repeats / 2)].toFixed(1) }));
  }
}
