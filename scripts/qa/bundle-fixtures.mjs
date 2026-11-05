import { zipStored, stylesXml } from './docx-fixtures.mjs';

/** A blue rectangle, with no fonts, images, network assets or PDF library. */
export function rectanglePdf() {
  const stream = '0 0 1 rg 10 10 80 80 re f\n';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Resources << >> /Contents 4 0 R >>',
    `<< /Length ${stream.length} >>\nstream\n${stream}endstream`,
  ];
  let source = '%PDF-1.4\n';
  const offsets = [0];
  for (let index = 0; index < objects.length; index++) {
    offsets.push(Buffer.byteLength(source));
    source += `${index + 1} 0 obj\n${objects[index]}\nendobj\n`;
  }
  const xref = Buffer.byteLength(source);
  source += `xref\n0 5\n0000000000 65535 f \n${offsets.slice(1).map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`).join('')}`;
  source += `trailer\n<< /Size 5 /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(source);
}

/**
 * EMF Header + CREATEBRUSHINDIRECT + SELECTOBJECT + RECTANGLE + EOF.
 * MS-EMF header and record definitions:
 * https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/e4a35c41-e8e3-43f9-bc07-a18e99bb866d
 * https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/3c471238-0a02-4992-90a2-bfd2afd98f2a
 */
function rectangleEmf() {
  const record = (values) => {
    const bytes = Buffer.alloc(values.length * 4);
    values.forEach((value, index) => bytes.writeUInt32LE(value, index * 4));
    return bytes;
  };
  const records = [
    record([39, 24, 1, 0, 0x00ff0000, 0]),
    record([37, 12, 1]),
    record([43, 24, 10, 10, 90, 90]),
    record([14, 20, 0, 0, 20]),
  ];
  const header = record([
    1, 88, 0, 0, 100, 100, 0, 0, 2646, 2646,
    0x464d4520, 0x00010000, 88 + records.reduce((sum, item) => sum + item.length, 0),
    5, 2, 0, 0, 0, 100, 100, 26, 26,
  ]);
  return Buffer.concat([header, ...records]);
}

/** Open through the real host/document path to exercise SuperDoc's EMF fallback. */
export function emfDocx() {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const drawing = `<w:p><w:r><w:drawing><wp:inline>
    <wp:extent cx="952500" cy="952500"/><wp:docPr id="1" name="bundle-emf"/>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>
      <pic:nvPicPr><pic:cNvPr id="1" name="bundle-emf"/><pic:cNvPicPr/></pic:nvPicPr>
      <pic:blipFill><a:blip r:embed="rImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
      <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
    </pic:pic></a:graphicData></a:graphic>
  </wp:inline></w:drawing></w:r></w:p>`;
  return zipStored({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/><Default Extension="emf" ContentType="image/x-emf"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rImage" Type="${r}/image" Target="media/rectangle.emf"/><Relationship Id="rStyles" Type="${r}/styles" Target="styles.xml"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${w}" xmlns:r="${r}"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>
      <w:p><w:r><w:t>Bundle EMF fixture</w:t></w:r></w:p>${drawing}
      <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
    </w:body></w:document>`,
    'word/styles.xml': stylesXml(),
    'word/media/rectangle.emf': rectangleEmf(),
  });
}
