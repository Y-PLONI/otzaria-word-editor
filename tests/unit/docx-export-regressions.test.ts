import { describe, expect, it } from 'vitest';
import { hasDeclaredRtlRuns, markRtlRuns } from '../../src/engine/docx-run-direction';
import { readStyleSheet } from '../../src/engine/docx-style-sheet';
import { postflightDocx } from '../../src/engine/docx-postflight';
import { crc32, readEntryText, readZip, writeZip, type Bytes, type ZipEntry } from '../../src/engine/docx-parts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const NS = `xmlns:w="${W}" xmlns:mc="${MC}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
const doc = (body: string) => `<w:document ${NS}><w:body>${body}</w:body></w:document>`;
const run = (text: string, props = '') => `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t>${text}</w:t></w:r>`;
const para = (runs: string, props = '<w:bidi/>') => `<w:p><w:pPr>${props}</w:pPr>${runs}</w:p>`;
const styles = (defaults: string, extra = '') =>
  `<w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr>${defaults}</w:rPr></w:rPrDefault></w:docDefaults>${extra}</w:styles>`;
const tableStyle = (id: string, props: string, extra = '', attrs = '') =>
  `<w:style w:type="table" w:styleId="${id}" ${attrs}>${extra}<w:rPr>${props}</w:rPr></w:style>`;
const table = (body: string, id: string | null, props = '') =>
  `<w:tbl><w:tblPr>${id === null ? '' : `<w:tblStyle w:val="${id}"/>`}${props}</w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc>${body}</w:tc></w:tr></w:tbl>`;
const aware = readStyleSheet(styles('<w:rFonts w:ascii="Courier New" w:cs="David"/><w:szCs w:val="32"/>'));
const latinDefaults = '<w:rFonts w:ascii="Arial"/><w:sz w:val="20"/>';
const xmlTree = (xml: string) => {
  const parsed = new DOMParser().parseFromString(xml, 'application/xml');
  expect(parsed.getElementsByTagName('parsererror')).toHaveLength(0);
  return parsed;
};
const textOf = (xml: string) => Array.from(xmlTree(xml).getElementsByTagNameNS(W, 't'), (t) => t.textContent).join('');
const propsOf = (xml: string) => Array.from(xmlTree(xml).getElementsByTagNameNS(W, 'r'), (r) =>
  (r.getElementsByTagNameNS(W, 'rPr')[0]?.outerHTML ?? '').replace(/\sxmlns(?::[\w.-]+)?="[^"]*"/g, ''));

const enc = new TextEncoder();
function entry(name: string, text: string): ZipEntry {
  const data = enc.encode(text) as Bytes;
  return {
    name, nameBytes: enc.encode(name) as Bytes, data, uncompressedSize: data.length, crc: crc32(data),
    versionMadeBy: 20, versionNeeded: 10, flags: 0, method: 0, modTime: 0, modDate: 0x21,
    internalAttrs: 0, externalAttrs: 0,
  };
}
async function exported(styleXml: string, documentXml: string): Promise<Record<string, string>> {
  const bytes = writeZip([entry('word/styles.xml', styleXml), entry('word/document.xml', documentXml)]);
  const result = (await postflightDocx(bytes)) ?? bytes;
  const parts: Record<string, string> = {};
  for (const part of readZip(result)!) parts[part.name] = (await readEntryText(part))!;
  return parts;
}

const alternatives = '<mc:AlternateContent><mc:Choice Requires="w14"><w:t>שלום abc</w:t></mc:Choice>' +
  '<mc:Fallback><w:t>שלום abc</w:t></mc:Fallback></mc:AlternateContent>';

describe('export preserves XML text and foreign run content', () => {
  it('does not flatten Choice and Fallback into duplicated visible text', async () => {
    const opaqueRun = `<w:r>${alternatives}</w:r>`;
    const xml = doc(para(opaqueRun) + para(run('עברית')));
    const parts = await exported(styles(latinDefaults), xml);
    const output = parts['word/document.xml']!;
    expect(output).toContain(opaqueRun);
    expect(xmlTree(output).getElementsByTagNameNS(MC, 'AlternateContent')).toHaveLength(1);
    expect(output).toContain('<w:rtl/>'); // Other paragraphs still get converted.
  });

  it.each(['שלום abc', 'שלום'])('leaves a foreign wrapper intact for %s', (text) => {
    const xml = doc(para(`<w:r><x:wrapper xmlns:x="urn:test"><w:t>${text}</w:t></x:wrapper></w:r>`));
    expect(markRtlRuns(xml, aware)).toBeNull();
  });

  it.each([
    '<![CDATA[עולם xyz]]>',
    '<![CDATA[שלום]]>',
    'שלום <!--עברית-->',
    '<![CDATA[שלום &amp; abc]]>',
    'שלום <!--editor note--> abc',
    'שלום <?editor note?> abc',
  ])('does not turn opaque text markup into escaped visible text: %s', async (text) => {
    const xml = doc(para(run(text)));
    const parts = await exported(styles(latinDefaults), xml);
    expect(parts['word/document.xml']).toContain(`<w:t>${text}</w:t>`);
    expect(textOf(parts['word/document.xml']!)).toBe(textOf(xml));
  });

  it('comments with Latin letters do not classify actual Hebrew text as mixed', () => {
    const xml = doc(para(run('שלום <!--abc--> עולם')));
    const output = markRtlRuns(xml, aware)!;
    expect(output).toContain('<w:rtl/>');
    expect(output).toContain('<w:t>שלום <!--abc--> עולם</w:t>');
    expect(textOf(output)).toBe('שלום  עולם');
  });

  it('does not split a run with both rtl and cs enabled', () => {
    expect(markRtlRuns(doc(para(run('שלום abc', '<w:rtl/><w:cs/>'))), aware)).toBeNull();
  });

  it('does not split a mixed run that inherits its complex-script declaration', () => {
    const inherited = readStyleSheet(styles('<w:rtl/>'));
    expect(markRtlRuns(doc(para(run('שלום abc'))), inherited)).toBeNull();
  });
});

describe('live RTL detection and style alignment', () => {
  const engine = styles('<w:rFonts w:ascii="Arial" w:cs="Times New Roman (Body CS)"/><w:szCs w:val="24"/>');
  const history = '<w:rPrChange w:id="1" w:author="a"><w:rPr><w:rtl/></w:rPr></w:rPrChange>';

  it('does not count a historical run declaration, and aligns the engine font', async () => {
    const xml = doc(para(run('שלום', history)));
    expect(hasDeclaredRtlRuns(xml)).toBe(false);
    const parts = await exported(engine, xml);
    expect(parts['word/styles.xml']).toContain('w:cs="Arial"');
    expect(parts['word/document.xml']).toContain(`<w:rtl/>${history}`);
  });

  it('a live declaration still counts alongside a historical disabled declaration', () => {
    const offHistory = history.replace('<w:rtl/>', '<w:rtl w:val="0"/>');
    expect(hasDeclaredRtlRuns(doc(para(run('שלום', `<w:rtl/>${offHistory}`))))).toBe(true);
  });

  it('cs enabled still counts when rtl is explicitly disabled', async () => {
    const xml = doc(para(run('שלום', '<w:rtl w:val="0"/><w:cs/>')));
    expect(hasDeclaredRtlRuns(xml)).toBe(true);
    expect((await exported(engine, xml))['word/styles.xml']).toBe(engine);
  });

  it('does not count paragraph-mark properties, comments, or a foreign math run', () => {
    expect(hasDeclaredRtlRuns(doc(para(run('שלום'), '<w:rPr><w:rtl/></w:rPr>')))).toBe(false);
    expect(hasDeclaredRtlRuns(doc(`<!--${para(run('שלום', '<w:rtl/>'))}-->`))).toBe(false);
    expect(hasDeclaredRtlRuns(doc('<w:p><m:r xmlns:m="urn:math"><m:rPr><m:rtl/></m:rPr><m:t>שלום</m:t></m:r></w:p>'))).toBe(false);
  });

  it('does not count Hebrew contained only in an XML comment', () => {
    expect(hasDeclaredRtlRuns(doc(para(run('abc <!--שלום-->', '<w:rtl/>'))))).toBe(false);
  });

  it('counts actual CDATA text without decoding entities inside it', () => {
    expect(hasDeclaredRtlRuns(doc(para(run('<![CDATA[שלום abc]]>', '<w:rtl/>'))))).toBe(true);
    expect(hasDeclaredRtlRuns(doc(para(run('<![CDATA[&#1513;]]>', '<w:rtl/>'))))).toBe(false);
  });

  it('detects and marks Hebrew encoded entirely as numeric character references', () => {
    const text = '&#1513;&#x5dc;&#1493;&#1501;';
    const xml = doc(para(run(text)));
    expect(markRtlRuns(xml, aware)).toContain(`<w:rtl/></w:rPr><w:t>${text}</w:t>`);
    expect(hasDeclaredRtlRuns(doc(para(run(text, '<w:rtl/>'))))).toBe(true);
  });

  it('supports the declared Word namespace prefix instead of matching local names', () => {
    const xml = doc(para(run('שלום', '<w:rtl/>'))).replace(/w:/g, 'word:').replace('xmlns:w=', 'xmlns:word=');
    expect(hasDeclaredRtlRuns(xml)).toBe(true);
  });

  it('does not align styles whose conditional table formatting declares complex script', async () => {
    const conditional = tableStyle('Table', '', '<w:tblStylePr w:type="firstRow"><w:rPr><w:rtl/></w:rPr></w:tblStylePr>');
    const stylesheet = engine.replace('</w:styles>', `${conditional}</w:styles>`);
    const xml = doc(table(para(run('שלום')), 'Table'));
    expect((await exported(stylesheet, xml))['word/styles.xml']).toBe(stylesheet);
  });
});

describe('table style inheritance during run marking', () => {
  it('does not override a table style that declares Hebrew font and size', async () => {
    const sheet = styles(latinDefaults, tableStyle('Big', '<w:rFonts w:ascii="Courier New" w:cs="David"/><w:sz w:val="32"/><w:szCs w:val="36"/>'));
    const xml = doc(table(para(run('שלום')), 'Big'));
    const output = (await exported(sheet, xml))['word/document.xml']!;
    expect(propsOf(output)).toEqual(['<w:rPr><w:rtl/></w:rPr>']);
    expect((await exported(sheet, output))['word/document.xml']).toBe(output);
  });

  it('inherits a table style through basedOn', () => {
    const sheet = readStyleSheet(styles(latinDefaults,
      tableStyle('Base', '<w:rFonts w:cs="David"/><w:szCs w:val="36"/>') +
      tableStyle('Child', '', '<w:basedOn w:val="Base"/>')));
    const output = markRtlRuns(doc(table(para(run('שלום')), 'Child')), sheet)!;
    expect(propsOf(output)).toEqual(['<w:rPr><w:rtl/></w:rPr>']);
  });

  it('falls back to table Latin values when the entire chain has no complex-script values', () => {
    const sheet = readStyleSheet(styles(latinDefaults, tableStyle('Latin', '<w:rFonts w:ascii="Courier New"/><w:sz w:val="36"/><w:b/>')));
    const output = markRtlRuns(doc(table(para(run('שלום')), 'Latin')), sheet)!;
    expect(output).toContain('<w:rFonts w:cs="Courier New"/><w:bCs/><w:szCs w:val="36"/><w:rtl/>');
  });

  it('uses the default table style only inside a table, with distinct cached results', () => {
    const sheet = readStyleSheet(styles(latinDefaults,
      tableStyle('Default', '<w:rFonts w:cs="David"/><w:szCs w:val="36"/>', '', 'w:default="1"')));
    const output = markRtlRuns(doc(table(para(run('בתא')), null) + para(run('בחוץ'))), sheet)!;
    const props = propsOf(output);
    expect(props[0]).toBe('<w:rPr><w:rtl/></w:rPr>');
    expect(props[1]).toContain('w:cs="Arial"');
    expect(props[1]).toContain('w:val="20"');
  });

  it('restores the enclosing style after leaving a nested table', () => {
    const sheet = readStyleSheet(styles(latinDefaults,
      tableStyle('Outer', '<w:rFonts w:ascii="Courier New"/>') + tableStyle('Inner', '<w:rFonts w:ascii="David"/>')));
    const xml = doc(table(para(run('חוץ')) + table(para(run('פנים')), 'Inner') + para(run('חוץשוב')), 'Outer'));
    const props = propsOf(markRtlRuns(xml, sheet)!);
    expect(props.map((p) => /w:cs="([^"]+)"/.exec(p)?.[1])).toEqual(['Courier New', 'David', 'Courier New']);
  });

  it('ignores historical table style changes', () => {
    const sheet = readStyleSheet(styles(latinDefaults, tableStyle('Live', '<w:rFonts w:cs="David"/><w:szCs w:val="36"/>')));
    const change = '<w:tblPrChange w:id="1" w:author="a"><w:tblPr><w:tblStyle w:val="Old"/></w:tblPr></w:tblPrChange>';
    const output = markRtlRuns(doc(table(para(run('שלום')), 'Live', change)), sheet)!;
    expect(propsOf(output)).toEqual(['<w:rPr><w:rtl/></w:rPr>']);
  });

  it('preserves conditional table formatting, including inherited conditions', () => {
    const conditional = '<w:tblStylePr w:type="firstRow"><w:rPr><w:rFonts w:cs="David"/><w:szCs w:val="36"/></w:rPr></w:tblStylePr>';
    const sheet = readStyleSheet(styles(latinDefaults,
      tableStyle('Base', '', conditional) + tableStyle('Child', '', '<w:basedOn w:val="Base"/>')));
    expect(markRtlRuns(doc(table(para(run('שלום abc')), 'Child')), sheet)).toBeNull();
  });

  it('does not skip conversion for conditional cell borders that do not affect run or paragraph formatting', () => {
    const sheet = readStyleSheet(styles(latinDefaults,
      tableStyle('Border', '', '<w:tblStylePr w:type="firstRow"><w:tcPr><w:tcBorders/></w:tcPr></w:tblStylePr>')));
    expect(markRtlRuns(doc(table(para(run('שלום')), 'Border')), sheet)).toContain('<w:rtl/>');
  });

  it('respects complex-script properties inherited from the table style', () => {
    const sheet = readStyleSheet(styles(latinDefaults, tableStyle('Complex', '<w:rtl/><w:rFonts w:cs="David"/>')));
    expect(markRtlRuns(doc(table(para(run('שלום abc')), 'Complex')), sheet)).toBeNull();
  });
});

describe('neutral-run donor lookup', () => {
  it('keeps backward priority and uses a following donor only when none precedes it', () => {
    const output = markRtlRuns(doc(para(run('1') + run('שלום') + run('.') + run('abc') + run('?'))), aware)!;
    expect(propsOf(output).map((p) => p.includes('<w:rtl/>'))).toEqual([true, true, true, false, false]);
  });

  it('an existing declaration blocks donation across a neutral chain', () => {
    const xml = doc(para(run('שלום', '<w:rtl/>') + run('1') + run('2') + run('עולם')));
    const output = markRtlRuns(xml, aware)!;
    expect(propsOf(output).map((p) => p.includes('<w:rtl/>'))).toEqual([true, false, false, true]);
  });

  it('marks a large neutral chain without overflowing the edit-array call stack', () => {
    const count = 128_000;
    const xml = doc(para(run('א') + run('1').repeat(count)));
    const output = markRtlRuns(xml, aware)!;
    expect(output.match(/<w:rtl\/>/g)).toHaveLength(count + 1);
    expect(output.match(/<w:t[^>]*>1<\/w:t>/g)).toHaveLength(count);
    expect(markRtlRuns(output, aware)).toBeNull();
  });
});
