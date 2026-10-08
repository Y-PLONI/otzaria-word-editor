/**
 * הדרך החוצה מקצה לקצה, על חבילה שלמה: הסדר בין היישור, הסימון והמראה, וההכרעה
 * מתי מיישרים — שתלויה בכל חלקי המסמך ולא בחלק אחד.
 */
import { describe, it, expect } from 'vitest';
import { postflightDocx } from '../../src/engine/docx-postflight';
import { crc32, readEntryText, readZip, writeZip, type Bytes, type ZipEntry } from '../../src/engine/docx-parts';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const enc = new TextEncoder();

function entry(name: string, text: string): ZipEntry {
  const data = enc.encode(text) as Bytes;
  return {
    name,
    nameBytes: enc.encode(name) as Bytes,
    versionMadeBy: 20,
    versionNeeded: 10,
    flags: 0,
    method: 0,
    modTime: 0,
    modDate: 0x21,
    crc: crc32(data),
    internalAttrs: 0,
    externalAttrs: 0,
    data,
    uncompressedSize: data.byteLength,
  };
}

async function part(bytes: Bytes, name: string): Promise<string> {
  const found = readZip(bytes)!.find((e) => e.name === name)!;
  return (await readEntryText(found))!;
}

const run = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const body = (...runs: string[]) =>
  `<w:document ${W}><w:body>${runs.map((r) => `<w:p><w:pPr><w:bidi/></w:pPr>${r}</w:p>`).join('')}</w:body></w:document>`;

/** מסמך Word עברי: הגופן העברי David 16, הלטיני Courier 10. */
const WORD_STYLES =
  `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="David"/>` +
  `<w:sz w:val="20"/><w:szCs w:val="32"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`;

/** החתימה של תבנית המנוע — מסמך שנוצר בתוסף. */
const ENGINE_STYLES =
  `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Times New Roman (Body CS)"/>` +
  `<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="h1"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/></w:rPr></w:style></w:styles>`;

const pack = (styles: string, document: string) =>
  writeZip([entry('word/styles.xml', styles), entry('word/document.xml', document)]);

describe('postflightDocx — הטקסט שנכתב בתוסף יוצא כעברית', () => {
  it('מסמך Word: הריצה החדשה מוצהרת ויורשת את הגופן העברי; הסגנונות אינם נגעים', async () => {
    const out = (await postflightDocx(pack(WORD_STYLES, body(run('ישן', '<w:rtl/>'), run('חדש')))))!;
    const doc = await part(out, 'word/document.xml');
    expect(doc).toContain('<w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">חדש</w:t>');
    expect(doc).not.toContain('w:cs="Courier');
    expect(await part(out, 'word/styles.xml')).toBe(WORD_STYLES);
  });

  it('הדגשה שהמנוע כתב בצד הלטיני — התאום נכתב גם על ריצה שזה עתה סומנה', async () => {
    const out = (await postflightDocx(pack(WORD_STYLES, body(run('חדש', '<w:b/><w:sz w:val="28"/>')))))!;
    expect(await part(out, 'word/document.xml')).toContain(
      '<w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/><w:rtl/></w:rPr>',
    );
  });

  it('מסמך שנוצר בתוסף: הסגנונות מיושרים, והריצה רק מוצהרת', async () => {
    const out = (await postflightDocx(pack(ENGINE_STYLES, body(run('שלום')))))!;
    const styles = await part(out, 'word/styles.xml');
    expect(styles).toContain('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>');
    expect(styles).toContain('w:cstheme="majorHAnsi"');
    expect(await part(out, 'word/document.xml')).toContain('<w:rPr><w:rtl/></w:rPr>');
  });

  it('מסמך של התוסף שנערך גם ב-Word: אין יישור — העברית שנכתבה שם כבר נראית לפי הצד המורכב', async () => {
    const out = (await postflightDocx(pack(ENGINE_STYLES, body(run('ישן', '<w:rtl/>'), run('חדש')))))!;
    expect(await part(out, 'word/styles.xml')).toBe(ENGINE_STYLES);
    expect(await part(out, 'word/document.xml')).toContain('<w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">חדש');
  });

  it('עברית מוצהרת בכותרת עליונה בלבד חוסמת גם היא את היישור', async () => {
    const out = (await postflightDocx(
      writeZip([
        entry('word/styles.xml', ENGINE_STYLES),
        entry('word/document.xml', body(run('חדש'))),
        entry('word/header1.xml', `<w:hdr ${W}><w:p>${run('כותרת', '<w:rtl/>')}</w:p></w:hdr>`),
      ]),
    ))!;
    expect(await part(out, 'word/styles.xml')).toBe(ENGINE_STYLES);
  });

  it('מסמך לטיני אינו נגע', async () => {
    expect(await postflightDocx(pack(WORD_STYLES, body(run('hello'))))).toBeNull();
  });
});
