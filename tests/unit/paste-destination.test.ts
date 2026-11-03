import { describe, expect, it } from 'vitest';
import { marksFromDestinationCapture } from '../../src/engine/paste-destination';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const capture = (props: string, style = '') => ({
  fragment: { xml: `<w:r><w:rPr>${props}</w:rPr><w:t>x</w:t></w:r>` },
  source: { paragraphStyleId: style ? 'Destination' : null },
  dependencies: [{ target: '/word/styles.xml', dataBase64: btoa(`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>${style ? `<w:style w:type="paragraph" w:styleId="Destination"><w:rPr>${style}</w:rPr></w:style>` : ''}</w:styles>`) }],
});
describe('destination font capture', () => {
  it('keeps a directly authored family/size equal to defaults over a different paragraph style', () => {
    expect(marksFromDestinationCapture(capture('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/>', '<w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/><w:sz w:val="36"/><w:szCs w:val="36"/>')))
      .toMatchObject({ fontFamily: 'Arial', fontSizePt: 12, fontSizeCsPt: 12 });
  });
  it('resolves the paragraph style rather than falsely classifying inherited size as direct', () => {
    expect(marksFromDestinationCapture(capture('', '<w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/><w:sz w:val="36"/><w:szCs w:val="36"/>')))
      .toMatchObject({ fontFamily: 'David', fontSizePt: 18, fontSizeCsPt: 18 });
  });
  it('captures split Latin/Hebrew fonts as two channels instead of forcing one onto the other', () => {
    expect(marksFromDestinationCapture(capture('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="David"/>')))
      .toMatchObject({ fontFamily: 'Arial', fontFamilyCs: 'David' });
  });
  it('still rejects two different Latin channels', () => {
    expect(() => marksFromDestinationCapture(capture('<w:rFonts w:ascii="Arial" w:hAnsi="Times New Roman" w:cs="David"/>'))).toThrow();
  });
  it('accepts what Hebrew Word runs carry: rtl, hint, kern, proofing, and the paragraph mark rPr', () => {
    for (const props of ['<w:rtl/>', '<w:rFonts w:hint="cs"/>', '<w:kern w:val="2"/>', '<w:noProof/>', '<w:rFonts w:eastAsia="MS Mincho"/>']) {
      expect(marksFromDestinationCapture(capture(props))).toMatchObject({ fontFamily: 'Arial', fontSizePt: 12 });
    }
    expect(marksFromDestinationCapture({ ...capture('<w:rtl/>'), context: { xml: ['<w:pPr><w:bidi/><w:rPr><w:rtl/></w:rPr></w:pPr>'] } }))
      .toMatchObject({ fontFamily: 'Arial' });
  });
  it('rejects direct theme fonts, character styles, and appearance merge cannot carry', () => {
    for (const props of ['<w:cs/>', '<w:highlight w:val="none"/>', '<w:shd w:fill="FFFF00"/>', '<w:spacing w:val="20"/>', '<w:rFonts w:asciiTheme="majorHAnsi"/>', '<w:rStyle w:val="Emphasis"/>']) {
      expect(() => marksFromDestinationCapture(capture(props))).toThrow();
    }
    expect(() => marksFromDestinationCapture({ ...capture(''), context: { xml: ['<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>'] } })).toThrow();
  });
  it('retains explicit CS false and separate authored size channels', () => {
    expect(marksFromDestinationCapture(capture('<w:bCs w:val="0"/><w:iCs w:val="false"/><w:sz w:val="22"/><w:szCs w:val="32"/>')))
      .toMatchObject({ boldCs: false, italicCs: false, fontSizePt: 11, fontSizeCsPt: 16 });
  });
});
