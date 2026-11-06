/**
 * הצגת הגופן והגודל העבריים של הסגנונות בחריץ הלטיני — והחזרתם בשמירה.
 *
 * שתי חובות, ושתיהן נבדקות על הפתרון ולא רק על המחרוזת:
 * 1. אחרי ההצגה, השרשרת **הלטינית** של כל צירוף סגנונות נפתרת למה שהשרשרת
 *    **המורכבת** נפתרה לפני כן — זה מה שהמנוע יצייר וזה מה ש-Word מצייר.
 * 2. ההחזרה של מה שלא שונה מחזירה את `styles.xml` בייט-בבייט.
 */
import { describe, it, expect } from 'vitest';
import {
  DISPLAY_NS,
  NO_THEME_FONTS,
  readThemeFonts,
  restoreComplexScriptStyles,
  showComplexScriptRuns,
  showComplexScriptStyles,
  themeFontName,
  type ThemeFonts,
} from '../../src/engine/docx-cs-display';
import { readStyleSheet, resolveRun, type FontRef } from '../../src/engine/docx-style-sheet';
import { postflightDocx } from '../../src/engine/docx-postflight';
import { preflightDocx } from '../../src/engine/docx-preflight-core';
import { crc32, readZip, writeZip, readEntryText, decodeXml, type Bytes, type ZipEntry } from '../../src/engine/docx-parts';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const sheet = (rPrDefault: string, extra = '', root = NS): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles ${root}><w:docDefaults><w:rPrDefault><w:rPr>${rPrDefault}</w:rPr></w:rPrDefault></w:docDefaults>${extra}</w:styles>`;
const style = (id: string, rPr: string, { type = 'paragraph', basedOn = '', isDefault = false } = {}): string =>
  `<w:style w:type="${type}"${isDefault ? ' w:default="1"' : ''} w:styleId="${id}"><w:name w:val="${id}"/>` +
  `${basedOn ? `<w:basedOn w:val="${basedOn}"/>` : ''}<w:rPr>${rPr}</w:rPr></w:style>`;

/** ערכת נושא כמו של Office: `a:cs` ריק, והגופן העברי בטבלת הכתבים. */
const THEME = `<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Office">
<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Arab" typeface="Times New Roman"/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hebr" typeface="Arial"/><a:font script="Arab" typeface="Arial"/></a:minorFont>
</a:fontScheme></a:themeElements></a:theme>`;
const SETTINGS = `<w:settings ${NS}><w:themeFontLang w:val="en-US" w:bidi="he-IL"/></w:settings>`;
const FONTS: ThemeFonts = readThemeFonts(THEME, SETTINGS);

/** הסגנון הרגיל של המסמך שדווח: לטיני „+גוף”, עברי גופן מותקן. */
const REPORTED = sheet(
  '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:szCs w:val="22"/>',
  style('Normal', '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cs="Guttman Keren"/><w:sz w:val="24"/><w:szCs w:val="28"/>', { isDefault: true }) +
    style('Heading1', '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="36"/>', { basedOn: 'Normal' }) +
    style('Quote', '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="20"/>', { basedOn: 'Normal' }) +
    style('Emph', '<w:rFonts w:cs="David"/><w:szCs w:val="30"/>', { type: 'character' }) +
    style('Code', '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>', { type: 'character', basedOn: 'Emph' }),
  `${NS} ${MC} mc:Ignorable="w14"`,
);

const nameOf = (ref: FontRef | null): string | null =>
  ref === null ? null : ref.theme !== null ? themeFontName(ref.theme, FONTS) : ref.name;

describe('readThemeFonts', () => {
  it('*Bidi נפתר לפי הכתב של שפת ה-bidi, כש-a:cs ריק', () => {
    expect(themeFontName('minorBidi', FONTS)).toBe('Arial');
    expect(themeFontName('majorBidi', FONTS)).toBe('Times New Roman');
    expect(themeFontName('minorHAnsi', FONTS)).toBe('Calibri');
    expect(themeFontName('majorAscii', FONTS)).toBe('Calibri Light');
  });

  it('בלי שפת bidi ובלי a:cs — אין שם, ואין מה להציג', () => {
    const fonts = readThemeFonts(THEME, null);
    expect(themeFontName('minorBidi', fonts)).toBeNull();
    expect(themeFontName('minorBidi', NO_THEME_FONTS)).toBeNull();
  });

  it('a:cs מלא נבחר כשאין גופן לכתב', () => {
    const theme = THEME.replace('<a:cs typeface=""/><a:font script="Hebr" typeface="Arial"/>', '<a:cs typeface="Tahoma"/>');
    expect(themeFontName('minorBidi', readThemeFonts(theme, SETTINGS))).toBe('Tahoma');
  });
});

describe('showComplexScriptStyles', () => {
  it('המקרה שדווח: Normal עם לטיני „+גוף” ועברי מותקן — הלטיני מציג את העברי', () => {
    const shown = showComplexScriptStyles(REPORTED, FONTS)!;
    const run = resolveRun(readStyleSheet(shown), 'Normal', null);
    expect(run.latin.font).toEqual({ name: 'Guttman Keren', theme: null });
    expect(run.latin.size).toBe('28');
  });

  it('השרשרת הלטינית אחרי ההצגה = השרשרת המורכבת לפני, בכל צירוף', () => {
    const before = readStyleSheet(REPORTED);
    const after = readStyleSheet(showComplexScriptStyles(REPORTED, FONTS)!);
    for (const pStyle of ['Normal', 'Heading1', 'Quote', null]) {
      for (const rStyle of ['Emph', 'Code', null]) {
        const cs = resolveRun(before, pStyle, rStyle).cs;
        const latin = resolveRun(after, pStyle, rStyle).latin;
        expect({ pStyle, rStyle, font: nameOf(latin.font) }).toEqual({ pStyle, rStyle, font: nameOf(cs.font) });
        expect({ pStyle, rStyle, size: latin.size }).toEqual({ pStyle, rStyle, size: cs.size });
      }
    }
  });

  it('המורכב של הרמות אינו נגע — Word ממשיך לקרוא אותו כמו שהיה', () => {
    const before = readStyleSheet(REPORTED);
    const after = readStyleSheet(showComplexScriptStyles(REPORTED, FONTS)!);
    expect(after.defaults.fontCs).toEqual(before.defaults.fontCs);
    for (const [id, record] of before.styles) {
      expect(after.styles.get(id)!.run.fontCs).toEqual(record.run.fontCs);
      expect(after.styles.get(id)!.run.sizeCs).toEqual(record.run.sizeCs);
    }
  });

  it('לטיני בלבד, בלי מורכב מעליו — נשאר (השמירה תכתוב אותו לריצה)', () => {
    const xml = sheet('<w:sz w:val="24"/>', style('Normal', '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>', { isDefault: true }));
    expect(showComplexScriptStyles(xml, FONTS)).toBeNull();
  });

  it('גופן ערכת נושא שאינו נפתר — הרמה אינה נגעת', () => {
    const xml = sheet('<w:rFonts w:asciiTheme="minorHAnsi" w:cstheme="minorBidi"/>');
    expect(showComplexScriptStyles(xml, NO_THEME_FONTS)).toBeNull();
  });

  it('החריצים כבר שווים — אין מה להציג', () => {
    const xml = sheet('<w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/><w:sz w:val="24"/><w:szCs w:val="24"/>');
    expect(showComplexScriptStyles(xml, FONTS)).toBeNull();
  });

  it('מכריז את מרחב השמות כ-Ignorable — גם כשאין mc בשורש', () => {
    const shown = showComplexScriptStyles(sheet('<w:rFonts w:cs="David"/>'), FONTS)!;
    expect(shown).toContain(`xmlns:owcs="${DISPLAY_NS}"`);
    expect(shown).toMatch(/xmlns:mc="[^"]+" mc:Ignorable="owcs"/);
    const withMc = showComplexScriptStyles(REPORTED, FONTS)!;
    expect(withMc).toContain('mc:Ignorable="w14 owcs"');
  });

  it('טיוטה שנפתחת שוב — המקור הרשום אינו נדרס', () => {
    const once = showComplexScriptStyles(REPORTED, FONTS)!;
    expect(showComplexScriptStyles(once, FONTS)).toBeNull();
    expect(restoreComplexScriptStyles(once)).toBe(REPORTED);
  });

  it('טיוטה שבה שונה רק העברי — נפתחת עם העברי החדש, וההחזרה עדיין למקור', () => {
    const xml = sheet('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="David"/>');
    const draft = showComplexScriptStyles(xml, FONTS)!.replace('w:cs="David"', 'w:cs="Miriam"');
    const reopened = showComplexScriptStyles(draft, FONTS)!;
    expect(readStyleSheet(reopened).defaults.font).toEqual({ name: 'Miriam', theme: null });
    // שאר התג השתנה (cs), ולכן רק הלטיני חוזר — הערך, לא בהכרח סדר המאפיינים.
    const saved = readStyleSheet(restoreComplexScriptStyles(reopened)).defaults;
    expect(saved.font).toEqual({ name: 'Arial', theme: null });
    expect(saved.fontCs).toEqual({ name: 'Miriam', theme: null });
  });

  it('שם גופן עם מרכאה ממקור במרכאות בודדות — פתיחה חוזרת אינה מזריקה שוב', () => {
    const xml = sheet(`<w:rFonts w:ascii="X" w:hAnsi="X" w:cs='Quote"Font'/>`);
    const once = showComplexScriptStyles(xml, FONTS)!;
    expect(showComplexScriptStyles(once, FONTS)).toBeNull();
    expect(restoreComplexScriptStyles(once)).toBe(xml);
  });

  it('סגנון תו עם לטיני בלבד, כשהעברי מוצהר בסגנון הפסקה ברירת המחדל — יורש אותו', () => {
    const xml = sheet(
      '<w:sz w:val="22"/>',
      style('P', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="David"/>', { isDefault: true }) +
        style('Code', '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>', { type: 'character' }),
    );
    const shown = readStyleSheet(showComplexScriptStyles(xml, FONTS)!);
    expect(nameOf(resolveRun(shown, 'P', 'Code').latin.font)).toBe('David');
  });
});

describe('השורש: mc:Ignorable', () => {
  const MCURI = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  const roots = [
    `${NS} xmlns:mc="${MCURI}" mc:Ignorable='w14'`,
    `${NS} xmlns:mc='${MCURI}' mc:Ignorable="w14"`,
    `${NS} xmlns:mc="${MCURI}" mc:Ignorable=""`,
    `${NS} xmlns:mc="${MCURI}" mc:Ignorable="w14  w15"`,
    `${NS} xmlns:mc="${MCURI}"`,
    `${NS} xmlns:ve="${MCURI}" ve:Ignorable="w14"`,
  ];
  it.each(roots)('תקין אחרי ההצגה, וחוזר בייט-בבייט: %s', (root) => {
    const xml = sheet('<w:rFonts w:cs="David"/>', '', root);
    const shown = showComplexScriptStyles(xml, FONTS)!;
    const head = shown.slice(shown.indexOf('<w:styles'), shown.indexOf('>', shown.indexOf('<w:styles')));
    const names = [...head.matchAll(/\s([\w.:-]+)\s*=/g)].map((m) => m[1]);
    expect(new Set(names).size).toBe(names.length);
    expect(head).toMatch(/Ignorable=(["'])[^"']*\bowcs\1/);
    expect(restoreComplexScriptStyles(shown)).toBe(xml);
  });
});

describe('restoreComplexScriptStyles', () => {
  it('מה שלא שונה חוזר בייט-בבייט', () => {
    for (const xml of [REPORTED, sheet('<w:rFonts w:cs="David"/><w:szCs w:val="40"/>')]) {
      expect(restoreComplexScriptStyles(showComplexScriptStyles(xml, FONTS)!)).toBe(xml);
    }
  });

  it('גופן שהמשתמש בחר בדיאלוג (שלושת החריצים) — נשמר; הגודל שלא שונה — חוזר', () => {
    const shown = showComplexScriptStyles(REPORTED, FONTS)!;
    const edited = shown.replace(
      '<w:rFonts w:cs="Guttman Keren" w:ascii="Guttman Keren" w:hAnsi="Guttman Keren"/>',
      '<w:rFonts w:ascii="Frank Ruehl" w:hAnsi="Frank Ruehl" w:cs="Frank Ruehl"/>',
    );
    expect(edited).not.toBe(shown);
    const normal = readStyleSheet(restoreComplexScriptStyles(edited)).styles.get('Normal')!.run;
    expect(normal.font).toEqual({ name: 'Frank Ruehl', theme: null });
    expect(normal.size).toBe('24');
  });

  it('גודל שהמשתמש שינה — נשמר; הגופן שלא שונה — חוזר לערכת הנושא', () => {
    const shown = showComplexScriptStyles(REPORTED, FONTS)!;
    const edited = shown.replace(/(w:styleId="Normal"[\s\S]*?)<w:sz w:val="28"\/><w:szCs w:val="28"\/>/, '$1<w:sz w:val="40"/><w:szCs w:val="40"/>');
    expect(edited).not.toBe(shown);
    const normal = readStyleSheet(restoreComplexScriptStyles(edited)).styles.get('Normal')!.run;
    expect(normal.size).toBe('40');
    expect(normal.font).toEqual({ name: null, theme: 'minorHAnsi' });
  });

  it('לטיני שהוסר להצגה ולא שונה — חוזר; גודל שהוסר — חוזר למקומו בסכמה', () => {
    const xml = sheet(
      '<w:rFonts w:cs="David"/><w:szCs w:val="28"/>',
      style('Quote', '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:b/><w:sz w:val="20"/><w:lang w:val="en-US"/>'),
    );
    const shown = showComplexScriptStyles(xml, FONTS)!;
    const quote = readStyleSheet(shown).styles.get('Quote')!.run;
    expect(quote.font).toBeUndefined();
    expect(quote.size).toBeUndefined();
    expect(restoreComplexScriptStyles(shown)).toBe(xml);
  });

  it('מסמך בלי סימון — אינו נגע', () => {
    expect(restoreComplexScriptStyles(REPORTED)).toBeNull();
  });
});

describe('showComplexScriptRuns', () => {
  const doc = (body: string) => `<w:document ${NS}><w:body>${body}</w:body></w:document>`;
  const para = (rPr: string, text: string) => `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;

  it('ריצה עברית מוצהרת עם גופן וגודל עבריים ישירים — הלטיני מקבל אותם', () => {
    const out = showComplexScriptRuns(doc(para('<w:rFonts w:hint="cs" w:cs="Guttman Keren"/><w:szCs w:val="32"/><w:rtl/>', 'שלום')), FONTS)!;
    expect(out).toContain('<w:rFonts w:hint="cs" w:cs="Guttman Keren" w:ascii="Guttman Keren" w:hAnsi="Guttman Keren"/>');
    expect(out).toContain('<w:sz w:val="32"/><w:szCs w:val="32"/>');
  });

  it('w:cs במקום w:rtl — גם: זו אותה הצהרה', () => {
    expect(showComplexScriptRuns(doc(para('<w:rFonts w:cs="David"/><w:cs/>', 'שלום')), FONTS)).toContain('w:ascii="David"');
  });

  it('לא נוגעת: כל ריצה שאינה מוצהרת (גם עברית), rtl כבוי, ולטיני בלבד', () => {
    // עברית לא מוצהרת: מילה לטינית שתוקלד בה תפוצל בשמירה ותיקרא מהחריץ הלטיני.
    for (const [rPr, text] of [
      ['<w:rFonts w:cs="David"/>', 'שלום'],
      ['<w:rFonts w:cs="David"/>', 'שלום world'],
      ['<w:rFonts w:cs="David"/>', 'hello'],
      ['<w:rFonts w:cs="David"/>', '123'],
      ['<w:rFonts w:cs="David"/><w:rtl w:val="0"/>', 'שלום'],
      ['<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:rtl/>', 'שלום'],
    ]) {
      expect(showComplexScriptRuns(doc(para(rPr!, text!)), FONTS)).toBeNull();
    }
  });

  it('ריצה מעורבת מוצהרת — כן: Word קורא את כולה מהמורכב', () => {
    expect(showComplexScriptRuns(doc(para('<w:rFonts w:cs="David"/><w:rtl/>', 'שלום world')), FONTS)).toContain('w:ascii="David"');
  });

  it('rPrChange אינו העיצוב החי', () => {
    const rPr = '<w:rtl/><w:rPrChange w:id="1" w:author="a"><w:rPr><w:rFonts w:cs="David"/></w:rPr></w:rPrChange>';
    expect(showComplexScriptRuns(doc(para(rPr, 'שלום')), FONTS)).toBeNull();
  });
});

/* ------------------------------------------------------------------ */
/* מקצה לקצה: פתיחה → (המנוע) → שמירה                                   */
/* ------------------------------------------------------------------ */

const encoder = new TextEncoder();
function entry(name: string, text: string): ZipEntry {
  const data = encoder.encode(text) as Bytes;
  return {
    name,
    nameBytes: encoder.encode(name) as Bytes,
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

async function packageOf(parts: Record<string, string>): Promise<Bytes> {
  return writeZip(Object.entries(parts).map(([name, text]) => entry(name, text)));
}

async function partsOf(bytes: Bytes): Promise<Record<string, string>> {
  const out: Record<string, string> = {};
  for (const e of readZip(bytes)!) out[e.name] = (await readEntryText(e)) ?? decodeXml(e.data);
  return out;
}

const RELS = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>`;
const DOCUMENT = `<w:document ${NS}><w:body><w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>שלום</w:t></w:r></w:p></w:body></w:document>`;

describe('preflight → postflight', () => {
  it('הסגנונות מוצגים בפתיחה וחוזרים בשמירה כמו שהיו', async () => {
    const source = await packageOf({
      'word/document.xml': DOCUMENT,
      'word/styles.xml': REPORTED,
      'word/settings.xml': SETTINGS,
      'word/theme/theme1.xml': THEME,
      'word/_rels/document.xml.rels': RELS,
    });
    const opened = await preflightDocx(source);
    const inEngine = (await partsOf(opened!.bytes))['word/styles.xml']!;
    expect(resolveRun(readStyleSheet(inEngine), 'Normal', null).latin.font).toEqual({ name: 'Guttman Keren', theme: null });

    const saved = await postflightDocx(opened!.bytes);
    const out = (await partsOf(saved ?? opened!.bytes))['word/styles.xml']!;
    expect(out).toBe(REPORTED);
  });

  it('ערכת הנושא נמצאת דרך ההקשרים — „+גוף בכתב מורכב” בברירות המחדל', async () => {
    const styles = sheet('<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>');
    const opened = await preflightDocx(
      await packageOf({
        'word/document.xml': DOCUMENT,
        'word/styles.xml': styles,
        'word/settings.xml': SETTINGS,
        'word/theme/theme1.xml': THEME,
        'word/_rels/document.xml.rels': RELS,
      }),
    );
    const inEngine = (await partsOf(opened!.bytes))['word/styles.xml']!;
    expect(readStyleSheet(inEngine).defaults.font).toEqual({ name: 'Arial', theme: null });
  });

  it('מסמך התבנית בלי עברית מוצהרת — אינו מוצג (השמירה תיישר ללטיני)', async () => {
    const styles = sheet('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Times New Roman (Body CS)"/>');
    const plain = `<w:document ${NS}><w:body><w:p><w:r><w:t>שלום</w:t></w:r></w:p></w:body></w:document>`;
    const opened = await preflightDocx(await packageOf({ 'word/document.xml': plain, 'word/styles.xml': styles }));
    const inEngine = opened ? (await partsOf(opened.bytes))['word/styles.xml'] : styles;
    expect(inEngine).toBe(styles);
  });
});
