/**
 * תבנית „מסמך חדש”. הבדיקה הראשית רצה על המסמך הריק של המנוע **המותקן**, ולכן
 * שדרוג superdoc שמשנה את המסמך הריק נופל כאן — ולא כמסמך חדש בלי RTL.
 */
import { describe, expect, it } from 'vitest';
import {
  blankDocumentBlob,
  ENGINE_BLANK_PAGE_SIZE,
  patchBlankDocumentXml,
  patchBlankStylesXml,
} from '../../src/engine/blank-document';
import { PAPER_SIZES } from '../../src/engine/page-setup';
import {
  deriveHebrewBlankDocx,
  DOCUMENT_PART,
  engineBlankDocx,
  readZip,
  STYLES_PART,
} from '../../scripts/blank-docx';

const a4 = PAPER_SIZES.find((size) => size.id === 'a4')!;
const A4 = `<w:pgSz w:w="${a4.widthTwips}" w:h="${a4.heightTwips}" w:code="${a4.code}"/>`;

const DOCUMENT = `<w:document><w:body><w:p w14:paraId="1"/><w:sectPr>${ENGINE_BLANK_PAGE_SIZE}<w:pgMar w:top="1440"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>`;
const STYLES = `<w:styles><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults></w:styles>`;

describe('patchBlankDocumentXml', () => {
  it('A4, bidi במקטע לפני docGrid, ו-bidi בפסקה הראשונה', () => {
    expect(patchBlankDocumentXml(DOCUMENT)).toBe(
      `<w:document><w:body><w:p w14:paraId="1"><w:pPr><w:bidi/></w:pPr></w:p><w:sectPr>${A4}<w:pgMar w:top="1440"/><w:cols w:space="720"/><w:bidi/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>`,
    );
  });

  it('זורקת כשהמסמך הריק אינו Letter, או כשכבר יש בו bidi', () => {
    expect(() => patchBlankDocumentXml(DOCUMENT.replace(ENGINE_BLANK_PAGE_SIZE, A4))).toThrow(/pgSz/);
    expect(() => patchBlankDocumentXml(DOCUMENT.replace('<w:cols', '<w:bidi/><w:cols'))).toThrow(/bidi/);
  });
});

describe('patchBlankStylesXml', () => {
  it('bidi בברירת המחדל של הפסקאות, ושפת bidi עברית', () => {
    expect(patchBlankStylesXml(STYLES)).toBe(
      `<w:styles><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="he-IL"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:bidi/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    );
  });

  it('זורקת כשאין pPrDefault ריק', () => {
    expect(() => patchBlankStylesXml(STYLES.replace('<w:pPrDefault/>', '<w:pPrDefault><w:pPr/></w:pPrDefault>'))).toThrow(/pPrDefault/);
  });
});

describe('deriveHebrewBlankDocx — על המסמך הריק של המנוע המותקן', () => {
  const derived = readZip(
    deriveHebrewBlankDocx({ patchDocument: patchBlankDocumentXml, patchStyles: patchBlankStylesXml }),
  );
  const original = readZip(engineBlankDocx());
  const part = (entries: typeof derived, name: string) => entries.find((entry) => entry.name === name)!.data.toString('utf8');

  it('document.xml נושא A4 ו-bidi במקטע ובפסקה', () => {
    const xml = part(derived, DOCUMENT_PART);
    expect(xml).toContain(A4);
    expect(xml).toMatch(/<w:sectPr[^>]*>[\s\S]*<w:bidi\/><w:docGrid/);
    expect(xml).toMatch(/<w:p\b[^>]*><w:pPr><w:bidi\/><\/w:pPr><\/w:p>/);
  });

  it('styles.xml נושא bidi בברירת המחדל של הפסקאות', () => {
    expect(part(derived, STYLES_PART)).toContain('<w:pPrDefault><w:pPr><w:bidi/></w:pPr></w:pPrDefault>');
  });

  it('styles.xml: הגופן העברי של ברירות המחדל הוא הלטיני — Arial, ולא מחרוזת התפריט של Word ל-Mac', () => {
    const xml = part(derived, STYLES_PART);
    expect(part(original, STYLES_PART)).toContain('w:cs="Times New Roman (Body CS)"');
    expect(xml).not.toContain('Body CS');
    expect(xml).toMatch(/<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial"[^>]*w:cs="Arial"/);
    // כותרת 1: ערכת הנושא העברית (Times New Roman ב-Word) הוחלפה בלטינית.
    expect(xml).toMatch(/w:styleId="Heading1"[\s\S]*?w:cstheme="majorHAnsi"/);
  });

  it('שאר החלקים זהים בייט-בבייט, ובאותו סדר', () => {
    expect(derived.map((entry) => entry.name)).toEqual(original.map((entry) => entry.name));
    for (const entry of original) {
      if (entry.name === DOCUMENT_PART || entry.name === STYLES_PART) continue;
      expect(part(derived, entry.name)).toBe(entry.data.toString('utf8'));
    }
  });

  it('הקידוד ל-Blob מחזיר את אותם בייטים', async () => {
    const docx = deriveHebrewBlankDocx({ patchDocument: patchBlankDocumentXml, patchStyles: patchBlankStylesXml });
    const blob = blankDocumentBlob(docx.toString('base64'))!;
    // FileReader ולא blob.arrayBuffer(): ה-Blob של jsdom אינו מממש אותו.
    const bytes = await new Promise<Uint8Array>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(new Uint8Array(reader.result as ArrayBuffer));
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(blob);
    });
    expect(bytes).toEqual(new Uint8Array(docx));
    expect(blankDocumentBlob('')).toBeUndefined();
  });
});
