/**
 * היישור של הצד המורכב לצד הלטיני ב-`styles.xml` — במסמך שנוצר בתוסף.
 *
 * החובה: אחרי היישור, **כל** ריצה שתסומן `w:rtl` נפתרת ב-Word בדיוק כמו
 * שנפתרה בצד הלטיני. לכן הבדיקה המרכזית אינה על המחרוזת אלא על הפתרון: לכל
 * צירוף של סגנון פסקה וסגנון תו בתבנית האמיתית של המנוע, השרשרת המורכבת
 * מצהירה בדיוק מה שהלטינית מצהירה.
 */
import { describe, it, expect } from 'vitest';
import {
  ENGINE_TEMPLATE_CS_FONT,
  alignComplexScriptStyles,
  hasEngineTemplateSignature,
} from '../../src/engine/docx-cs-align';
import { readStyleSheet, type LevelProps } from '../../src/engine/docx-style-sheet';
import { engineBlankDocx, readZip } from '../../scripts/blank-docx';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const sheet = (rPrDefault: string, extra = ''): string =>
  `<w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr>${rPrDefault}</w:rPr></w:rPrDefault></w:docDefaults>${extra}</w:styles>`;
const style = (id: string, rPr: string, type = 'paragraph'): string =>
  `<w:style w:type="${type}" w:styleId="${id}"><w:name w:val="${id}"/><w:rPr>${rPr}</w:rPr></w:style>`;

function engineStyles(): string {
  const entry = readZip(engineBlankDocx()).find((e) => e.name === 'word/styles.xml');
  return entry!.data.toString('utf8');
}

/** הצד הלטיני והמורכב של רמה, באותה צורה — להשוואה. */
function sides(level: LevelProps) {
  const font = (f: LevelProps['font']) => (f ? (f.theme !== null ? `theme:${f.theme}` : `name:${f.name}`) : null);
  return {
    latin: [level.bold ?? null, level.italic ?? null, level.size ?? null, font(level.font)],
    cs: [level.boldCs ?? null, level.italicCs ?? null, level.sizeCs ?? null, font(level.fontCs)],
  };
}

describe('alignComplexScriptStyles', () => {
  it('תבנית המנוע נושאת את החתימה — והיא נעלמת אחרי היישור', () => {
    const xml = engineStyles();
    expect(xml).toContain(ENGINE_TEMPLATE_CS_FONT);
    expect(hasEngineTemplateSignature(xml)).toBe(true);
    expect(hasEngineTemplateSignature(alignComplexScriptStyles(xml))).toBe(false);
  });

  it('בתבנית האמיתית, כל רמה מצהירה בצד המורכב בדיוק מה שבלטיני', () => {
    const aligned = readStyleSheet(alignComplexScriptStyles(engineStyles()));
    const levels = [aligned.defaults, ...[...aligned.styles.values()].map((s) => s.run)];
    expect(levels.length).toBeGreaterThan(20);
    for (const level of levels) {
      const { latin, cs } = sides(level);
      expect(cs).toEqual(latin);
    }
  });

  it('גופן בשם: cs באותו שם, ו-cstheme יורד', () => {
    const out = alignComplexScriptStyles(sheet('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Times New Roman (Body CS)"/>'));
    expect(out).toContain('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>');
    const theme = alignComplexScriptStyles(sheet('<w:rFonts w:ascii="Arial" w:cstheme="minorBidi"/>'));
    expect(theme).toContain('<w:rFonts w:ascii="Arial" w:cs="Arial"/>');
  });

  it('גופן ערכת נושא: cstheme באותה ערכה', () => {
    const out = alignComplexScriptStyles(
      sheet('', style('H', '<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>')),
    );
    expect(out).toContain('w:hAnsiTheme="majorHAnsi" w:cstheme="majorHAnsi"/>');
  });

  it('ערכת הנושא העברית של התבנית ברמה בלי לטיני — יורדת, והרמה יורשת כמו הלטיני', () => {
    const out = alignComplexScriptStyles(sheet('', style('H3', '<w:rFonts w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/>')))!;
    expect(out).toContain('<w:rFonts w:eastAsiaTheme="majorEastAsia"/>');
  });

  it('הגדרה עברית שמשתמש קבע ב-Word אינה נמחקת (נמצא ב-QA)', () => {
    // מסמך של התוסף, ובו Normal שהמשתמש הגדיר ב-Word: גופן עברי David 16, בלי לטיני.
    const xml = sheet(
      '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Times New Roman (Body CS)"/>',
      style('Normal', '<w:rFonts w:cs="David"/><w:szCs w:val="32"/>'),
    );
    const out = alignComplexScriptStyles(xml)!;
    expect(out).toContain('<w:rPr><w:rFonts w:cs="David"/><w:szCs w:val="32"/></w:rPr>');
    expect(out).toContain('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>');
  });

  it('גופן עברי שמשתמש בחר לצד לטיני — אינו נדרס', () => {
    expect(alignComplexScriptStyles(sheet('<w:rFonts w:ascii="Arial" w:cs="David"/>'))).toBeNull();
    expect(alignComplexScriptStyles(sheet('<w:rFonts w:asciiTheme="minorHAnsi" w:cs="David"/>'))).toBeNull();
    // ערכת נושא שאינה העברית של התבנית — גם היא בחירה של מישהו.
    expect(alignComplexScriptStyles(sheet('<w:rFonts w:ascii="Arial" w:cstheme="minorHAnsi"/>'))).toBeNull();
    expect(alignComplexScriptStyles(sheet('<w:rFonts w:asciiTheme="majorHAnsi" w:cstheme="minorHAnsi"/>'))).toBeNull();
  });

  it('דגל חסר: התאום נכתב עם ערך השותף — גם כבוי', () => {
    const out = alignComplexScriptStyles(sheet('', style('S', '<w:b w:val="0"/><w:i/>')))!;
    expect(out).toContain('<w:b w:val="0"/><w:bCs w:val="0"/><w:i/><w:iCs/>');
  });

  it('גודל או דגל מורכב שקיים ושונה — אינו נגע: בתבנית אין זוג כזה', () => {
    expect(alignComplexScriptStyles(sheet('<w:sz w:val="22"/><w:szCs w:val="32"/><w:b w:val="0"/><w:bCs/>'))).toBeNull();
  });

  it('איבר זוגי נשאר זוג, ותאום נכנס אחרי הסגירה שלו', () => {
    const out = alignComplexScriptStyles(sheet('<w:b></w:b><w:rFonts w:ascii="A"></w:rFonts>'))!;
    expect(out).toContain('<w:b></w:b><w:bCs/><w:rFonts w:ascii="A" w:cs="A"></w:rFonts>');
  });

  it('rPrChange ו-pPr/rPr אינם רמות', () => {
    const xml = sheet(
      '',
      `<w:style w:type="paragraph" w:styleId="P"><w:pPr><w:rPr><w:b/></w:rPr></w:pPr>` +
        `<w:rPr><w:i/><w:iCs/><w:rPrChange w:id="1"><w:rPr><w:sz w:val="40"/></w:rPr></w:rPrChange></w:rPr></w:style>`,
    );
    expect(alignComplexScriptStyles(xml)).toBeNull();
  });

  it('מה שכבר מיושר אינו נכתב שוב', () => {
    const once = alignComplexScriptStyles(engineStyles())!;
    expect(alignComplexScriptStyles(once)).toBeNull();
  });
});
