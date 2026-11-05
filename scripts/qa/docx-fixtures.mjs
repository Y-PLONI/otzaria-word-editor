/**
 * Minimal docx builder (stored ZIP, CRC32) — no dependencies — and the
 * scenario documents of issue #14 (c) that list-resolver-qa.mjs opens.
 * `node scripts/qa/docx-fixtures.mjs <dir>` writes them as files for a manual look.
 */
import { writeFileSync } from 'node:fs';
import { deflateRawSync } from 'node:zlib';

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

export function zipStored(entries) {
  return zip(entries, false);
}

/** Compressed incoming archives for package-processing benchmarks. */
export function zipDeflated(entries) {
  return zip(entries, true);
}

function zip(entries, compressed) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  let count = 0;
  for (const [name, content] of Object.entries(entries)) {
    const nameBuf = Buffer.from(name, 'utf8');
    const raw = Buffer.from(content, 'utf8');
    const data = compressed ? deflateRawSync(raw) : raw;
    const crc = crc32(raw);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(compressed ? 8 : 0, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0x21, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(compressed ? 8 : 0, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0x21, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(raw.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    locals.push(local, nameBuf, data);
    centrals.push(central, nameBuf);
    offset += local.length + nameBuf.length + data.length;
    count += 1;
  }
  const cd = Buffer.concat(centrals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(count, 8);
  eocd.writeUInt16LE(count, 10);
  eocd.writeUInt32LE(cd.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);
  return Buffer.concat([...locals, cd, eocd]);
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

export function p(text, pPr = '') {
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}
export const numPr = (numId, ilvl = 0) => `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;

export function numberingXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W}>
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;
}

export function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/>${numPr(1, 1)}<w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="MyNumbered"><w:name w:val="My Numbered"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr>${numPr(1)}<w:ind w:left="720" w:hanging="360"/></w:pPr></w:style>
</w:styles>`;
}

export function buildDocx({ body, numbering = numberingXml(), styles = stylesXml() }) {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  return zipStored({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    'word/document.xml': documentXml,
    'word/numbering.xml': numbering,
    'word/styles.xml': styles,
  });
}

export const table = (cellParas) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${cellParas}</w:tc></w:tr></w:tbl>`;

const rtl = '<w:bidi/>';

/** Scenario docs keyed by name: { desc, target (text of the paragraph to put the caret in), docx }. */
export function scenarios() {
  return {
    styleNum: {
      desc: 'numbering from style pPr only (w:pStyle=MyNumbered whose pPr carries numPr), no direct numPr',
      target: 'סעיף שני מסגנון',
      docx: buildDocx({
        body:
          p('כותרת רגילה') +
          p('סעיף ראשון מסגנון', '<w:pStyle w:val="MyNumbered"/>' + rtl) +
          p('סעיף שני מסגנון', '<w:pStyle w:val="MyNumbered"/>' + rtl) +
          p('פסקה אחרונה'),
      }),
    },
    directNum: {
      desc: 'control: ListParagraph + direct numPr (typical Word output)',
      target: 'פריט שני ישיר',
      docx: buildDocx({
        body:
          p('כותרת רגילה') +
          p('פריט ראשון ישיר', '<w:pStyle w:val="ListParagraph"/>' + numPr(1) + rtl) +
          p('פריט שני ישיר', '<w:pStyle w:val="ListParagraph"/>' + numPr(1) + rtl) +
          p('פסקה אחרונה'),
      }),
    },
    inTable: {
      desc: 'list item inside a table cell',
      target: 'פריט בתא',
      docx: buildDocx({
        body:
          p('לפני הטבלה') +
          table(p('פריט בתא', '<w:pStyle w:val="ListParagraph"/>' + numPr(1) + rtl) + p('פריט שני בתא', numPr(1))) +
          p('אחרי הטבלה'),
      }),
    },
    numberedHeading: {
      desc: 'Word numbered heading: w:pStyle=Heading1 + direct numPr (blocks.list resolves it as heading, not listItem)',
      target: 'כותרת ממוספרת שנייה',
      docx: buildDocx({
        body:
          p('פתיח') +
          p('כותרת ממוספרת ראשונה', '<w:pStyle w:val="Heading1"/>' + numPr(1) + rtl) +
          p('גוף') +
          p('כותרת ממוספרת שנייה', '<w:pStyle w:val="Heading1"/>' + numPr(1) + rtl) +
          p('סיום'),
      }),
    },
    headingStyleNum: {
      desc: 'heading whose numbering comes from the style only (Heading2 style pPr carries numPr ilvl=1)',
      target: 'כותרת משנה מסגנון',
      docx: buildDocx({
        body: p('פתיח') + p('כותרת משנה מסגנון', '<w:pStyle w:val="Heading2"/>' + rtl) + p('סיום'),
      }),
    },
    numIdZero: {
      desc: 'direct numPr with numId=0',
      target: 'פריט אפס',
      docx: buildDocx({ body: p('לפני') + p('פריט אפס', numPr(0) + rtl) + p('אחרי') }),
    },
    numIdMissing: {
      desc: 'direct numPr with numId=7 that is absent from numbering.xml',
      target: 'פריט יתום',
      docx: buildDocx({ body: p('לפני') + p('פריט יתום', numPr(7) + rtl) + p('אחרי') }),
    },
  };
}

if (process.argv[1] && process.argv[1].replace(/\\/g, '/').endsWith('docx-fixtures.mjs')) {
  const out = process.argv[2] || '.';
  for (const [name, s] of Object.entries(scenarios())) writeFileSync(`${out}/${name}.docx`, s.docx);
  console.log('written', Object.keys(scenarios()).join(','));
}
