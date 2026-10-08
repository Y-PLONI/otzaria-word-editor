/**
 * הבאג שהבדיקות כאן שומרות עליו: „הטקסט שנכתב בתוכנה מוגדר ב-Word כטקסט
 * לטיני, ולכן אינו מושפע מהגדרת הגופן העברי”. המנוע כותב ריצות עבריות בלי
 * `<w:rtl/>`, ו-Word — שבוחר מחסנית לפי ההצהרה ולא לפי התווים — מצייר אותן
 * בגופן ובגודל הלטיניים של המסמך.
 *
 * שלוש חובות:
 *
 * 1. **הריצה העברית מוצהרת** — וגם הניטרלית שלידה, והקטע העברי של ריצה
 *    מעורבת (המילה הלועזית נשארת לטינית).
 * 2. **מה שהמסמך מצהיר לעברית חל עליה** — אין העתקה של הלטיני כשהמסמך מצהיר
 *    את הצד המורכב. זה מה שדווח.
 * 3. **רשת הביטחון** — כשאף רמה אינה מצהירה את הצד המורכב, התאום נכתב מהערך
 *    הלטיני האפקטיבי. בלעדיה Word נופל ל-Times New Roman 10 (נמדד).
 */
import { describe, it, expect } from 'vitest';
import { hasDeclaredRtlRuns, kindOf, markRtlRuns } from '../../src/engine/docx-run-direction';
import { classOf } from '../../src/engine/docx-neutral-mark';
import { mirrorComplexScript } from '../../src/engine/docx-cs-mirror';
import { readStyleSheet } from '../../src/engine/docx-style-sheet';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const doc = (body: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`;
const para = (runs: string, pPr = '<w:bidi/>'): string => `<w:p><w:pPr>${pPr}</w:pPr>${runs}</w:p>`;
const run = (text: string, rPr = ''): string =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const rtlCount = (xml: string | null): number => (xml?.match(/<w:rtl\s*\/>/g) ?? []).length;

/** גיליון סגנונות: ברירות מחדל וסגנונות נוספים. */
const styles = (rPrDefault: string, extra = '', pPrDefault = ''): string =>
  `<w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr>${rPrDefault}</w:rPr></w:rPrDefault>` +
  `<w:pPrDefault>${pPrDefault ? `<w:pPr>${pPrDefault}</w:pPr>` : ''}</w:pPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${extra}</w:styles>`;

/** מסמך Word עברי טיפוסי: לטינית Courier 10, עברית David 16 — הפיקסטורה שנמדדה ב-Word. */
const AWARE = readStyleSheet(
  styles('<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="David"/><w:sz w:val="20"/><w:szCs w:val="32"/>'),
);

/** מסמך שאינו מצהיר את הצד המורכב בשום רמה. */
const UNAWARE = readStyleSheet(
  styles(
    '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="32"/>',
    '<w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="h"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:rFonts w:asciiTheme="majorHAnsi"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Strong"><w:name w:val="s"/><w:rPr><w:i/></w:rPr></w:style>',
  ),
);

describe('סימון ריצות עבריות', () => {
  it('ריצה עברית בלי rPr מקבלת rPr עם ההצהרה', () => {
    expect(markRtlRuns(doc(para(run('שלום'))), AWARE)).toContain('<w:r><w:rPr><w:rtl/></w:rPr><w:t');
  });

  it('ההצהרה נכתבת במקומה בסכמה — לפני lang, ואחרי color', () => {
    const out = markRtlRuns(doc(para(run('שלום', '<w:color w:val="FF0000"/><w:lang w:bidi="he-IL"/>'))), AWARE);
    expect(out).toContain('<w:rPr><w:color w:val="FF0000"/><w:rtl/><w:lang w:bidi="he-IL"/></w:rPr>');
  });

  it('ההצהרה נכתבת לפני איבר ממרחב שמות אחר', () => {
    const xml = doc(para(run('שלום', '<w:sz w:val="24"/><w14:ligatures w14:val="standard"/>')));
    expect(markRtlRuns(xml, AWARE)).toContain('<w:sz w:val="24"/><w:rtl/><w14:ligatures');
  });

  it('ההצהרה נכתבת לפני rPrChange, ולא לתוך ההיסטוריה', () => {
    const rPr = '<w:b/><w:rPrChange w:id="1" w:author="a"><w:rPr><w:i/></w:rPr></w:rPrChange>';
    expect(markRtlRuns(doc(para(run('שלום', rPr))), AWARE)).toContain('<w:b/><w:rtl/><w:rPrChange');
  });

  it('נקודה בריצה משלה יורשת מהשכנה העברית', () => {
    expect(rtlCount(markRtlRuns(doc(para(run('שלום') + run('.'))), AWARE))).toBe(2);
  });

  it('ספרות ורווח בין ריצות עבריות מסומנים — כמו ש-Word מסמן הקלדה במקלדת עברית', () => {
    expect(rtlCount(markRtlRuns(doc(para(run('שנת ') + run('2024') + run(' היא'))), AWARE))).toBe(3);
  });

  it('ריצה לטינית ונקודה שאחריה אינן נוגעות', () => {
    const out = markRtlRuns(doc(para(run('hello') + run('.')) + para(run('שלום'))), AWARE);
    expect(out).toContain('<w:r><w:t xml:space="preserve">hello</w:t></w:r><w:r><w:t xml:space="preserve">.</w:t></w:r>');
  });

  it('הצהרה קיימת — דלוקה או כבויה, rtl או cs — אינה נגעת', () => {
    for (const rPr of ['<w:rtl/>', '<w:rtl w:val="0"/>', '<w:cs/>']) {
      expect(markRtlRuns(doc(para(run('שלום', rPr))), AWARE)).toBeNull();
    }
  });

  it('קוד שדה אינו מסומן', () => {
    const body = para(`<w:r><w:instrText>PAGE</w:instrText></w:r>` + run('שלום'));
    expect(rtlCount(markRtlRuns(doc(body), AWARE))).toBe(1);
  });

  it('rPr של סימן הפסקה אינה מקבלת הצהרה', () => {
    const body = para(run('שלום'), '<w:bidi/><w:rPr><w:b/></w:rPr>');
    expect(markRtlRuns(doc(body), AWARE)).toContain('<w:pPr><w:bidi/><w:rPr><w:b/></w:rPr></w:pPr>');
  });

  it('תוכן בתוך הערה אינו נגע', () => {
    const xml = doc(`<!-- ${para(run('שלום'))} -->` + para(run('hello')));
    expect(markRtlRuns(xml, AWARE)).toBeNull();
  });

  it('נוסחה (m:r) אינה ריצת Word', () => {
    const xml = doc(`<w:p><m:oMath xmlns:m="m"><m:r><m:t>א</m:t></m:r></m:oMath></w:p>`);
    expect(markRtlRuns(xml, AWARE)).toBeNull();
  });

  it('הרצה שנייה על הפלט אינה משנה דבר', () => {
    const once = markRtlRuns(doc(para(run('מילה Word בעברית.') + run('.'))), AWARE)!;
    expect(markRtlRuns(once, AWARE)).toBeNull();
  });

  it('מסמך בלי עברית אינו נוגע כלל', () => {
    expect(markRtlRuns(doc(para(run('hello world'))), AWARE)).toBeNull();
  });
});

describe('ריצה מעורבת מפוצלת', () => {
  it('„מילה Word בעברית.” — שלוש ריצות, והלועזית נשארת לטינית', () => {
    const out = markRtlRuns(doc(para(run('מילה Word בעברית.', '<w:color w:val="00FF00"/>'))), AWARE)!;
    const runs = out.match(/<w:r>[\s\S]*?<\/w:r>/g)!;
    expect(runs).toHaveLength(3);
    expect(runs[0]).toContain('<w:rtl/>');
    expect(runs[1]).not.toContain('rtl');
    expect(runs[1]).toContain('>Word<');
    expect(runs[2]).toContain('<w:rtl/>');
    for (const r of runs) expect(r).toContain('<w:color w:val="00FF00"/>');
  });

  it('בפסקה בלי bidi ישיר, הכיוון יורש מ-docDefaults — הרווח בין הכיוונים הולך איתו', () => {
    const sheet = readStyleSheet(styles('', '', '<w:bidi/>'));
    const out = markRtlRuns(doc(para(run('שלום world'), '')), sheet)!;
    expect(out).toContain('<w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">שלום </w:t>');
  });

  it('ריצה עם טאב אינה מפוצלת', () => {
    const body = para(`<w:r><w:t>שלום</w:t><w:tab/><w:t>world</w:t></w:r>`);
    expect(markRtlRuns(doc(body), AWARE)).toBeNull();
  });

  it('ניקוד נשאר עם האות שלפניו', () => {
    const out = markRtlRuns(doc(para(run('שָׁלוֹם abc'))), AWARE)!;
    expect(out).toContain('>שָׁלוֹם </w:t>');
  });
});

describe('מה שהמסמך מצהיר לעברית — חל עליה', () => {
  it('מסמך שמצהיר את הצד המורכב: רק ההצהרה, בלי העתקה של הלטיני', () => {
    expect(markRtlRuns(doc(para(run('שלום'))), AWARE)).toContain('<w:r><w:rPr><w:rtl/></w:rPr><w:t');
  });

  it('עיצוב ישיר שהמנוע כתב בצד הלטיני — המראה משלימה את התאום אחרי הסימון', () => {
    const marked = markRtlRuns(doc(para(run('שלום', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="40"/>'))), AWARE)!;
    const out = mirrorComplexScript(marked)!;
    expect(out).toContain('w:cs="Arial"');
    expect(out).toContain('<w:b/><w:bCs/>');
    expect(out).toContain('<w:sz w:val="40"/><w:szCs w:val="40"/>');
  });
});

describe('רשת הביטחון — כשאף רמה אינה מצהירה את הצד המורכב', () => {
  it('הגופן והגודל הלטיניים של docDefaults נכתבים כתאום', () => {
    const out = markRtlRuns(doc(para(run('שלום'))), UNAWARE);
    expect(out).toContain('<w:rPr><w:rFonts w:cs="Courier New"/><w:szCs w:val="32"/><w:rtl/></w:rPr>');
  });

  it('הדגשה מסגנון הפסקה, וגופן ערכת נושא — כערכת נושא', () => {
    const out = markRtlRuns(doc(para(run('שלום'), '<w:pStyle w:val="Heading"/><w:bidi/>')), UNAWARE)!;
    expect(out).toContain('<w:rFonts w:cstheme="majorHAnsi"/><w:bCs/><w:szCs w:val="32"/><w:rtl/>');
  });

  it('נטייה מסגנון תו', () => {
    const out = markRtlRuns(doc(para(run('שלום', '<w:rStyle w:val="Strong"/>'))), UNAWARE)!;
    expect(out).toContain('<w:rStyle w:val="Strong"/><w:rFonts w:cs="Courier New"/><w:iCs/><w:szCs w:val="32"/><w:rtl/>');
  });

  it('rFonts ישיר בלי גופן לטיני מקבל את המאפיין, ולא תג שני', () => {
    const out = markRtlRuns(doc(para(run('שלום', '<w:rFonts w:eastAsia="MS Mincho"/>'))), UNAWARE)!;
    expect(out).toContain('<w:rFonts w:eastAsia="MS Mincho" w:cs="Courier New"/>');
    expect(out.match(/<w:rFonts/g)).toHaveLength(1);
  });

  it('עיצוב ישיר אינו מטופל כאן — זה של המראה', () => {
    const out = markRtlRuns(doc(para(run('שלום', '<w:sz w:val="28"/>'))), UNAWARE)!;
    expect(out).not.toContain('szCs');
  });

  it('גיליון ריק: רק ההצהרה', () => {
    expect(markRtlRuns(doc(para(run('שלום'))))).toContain('<w:rPr><w:rtl/></w:rPr>');
  });
});

describe('hasDeclaredRtlRuns', () => {
  it('ריצה עברית מוצהרת — כן; לא מוצהרת, מכובה, או לטינית מוצהרת — לא', () => {
    expect(hasDeclaredRtlRuns(doc(para(run('שלום', '<w:rtl/>'))))).toBe(true);
    expect(hasDeclaredRtlRuns(doc(para(run('שלום', '<w:cs/>'))))).toBe(true);
    expect(hasDeclaredRtlRuns(doc(para(run('שלום'))))).toBe(false);
    expect(hasDeclaredRtlRuns(doc(para(run('שלום', '<w:rtl w:val="0"/>'))))).toBe(false);
    expect(hasDeclaredRtlRuns(doc(para(run('hello', '<w:rtl/>')) + para(run('שלום'))))).toBe(false);
  });
});

describe('עברית בירושה', () => {
  it('w:rtl ב-docDefaults — הריצה כבר עברית, ואין מה לשנות (נמדד בקורפוס)', () => {
    const sheet = readStyleSheet(styles('<w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/><w:rtl/>'));
    expect(markRtlRuns(doc(para(run('שלום', '<w:bCs/>'))), sheet)).toBeNull();
  });

  it('w:rtl בסגנון תו', () => {
    const sheet = readStyleSheet(styles('', '<w:style w:type="character" w:styleId="He"><w:name w:val="he"/><w:rPr><w:rtl/></w:rPr></w:style>'));
    expect(markRtlRuns(doc(para(run('שלום', '<w:rStyle w:val="He"/>'))), sheet)).toBeNull();
  });
});

describe('ממצאי QA', () => {
  const WORD = AWARE;

  it('לועזית שהמנוע צירף לריצה מוצהרת — מפוצלת, והקטע הלטיני יוצא בלי ההצהרה', () => {
    const out = markRtlRuns(doc(para(run('2026 סוף xyz', '<w:rtl/>'))), WORD)!;
    const runs = out.match(/<w:r>[\s\S]*?<\/w:r>/g)!;
    expect(runs.map((r) => /<w:rtl\/>/.test(r))).toEqual([true, false]);
    expect(runs[1]).toContain('>xyz<');
    expect(runs[1]).toContain('<w:rPr></w:rPr>');
  });

  it('ריצה מוצהרת דרך w:cs, או עם rPrChange — אינה מפוצלת', () => {
    expect(markRtlRuns(doc(para(run('סוף xyz', '<w:cs/>'))), WORD)).toBeNull();
    const change = '<w:rtl/><w:rPrChange w:id="1" w:author="a"><w:rPr/></w:rPrChange>';
    expect(markRtlRuns(doc(para(run('סוף xyz', change))), WORD)).toBeNull();
  });

  it('ריצה מוצהרת שכולה לועזית — אינה נגעת', () => {
    expect(markRtlRuns(doc(para(run('xyz', '<w:rtl/>'))), WORD)).toBeNull();
  });

  it('ספרות ש-Word כתב בין שתי ריצות מוצהרות — אינן יורשות', () => {
    const body = para(run('שלום', '<w:rtl/>') + run(' 2026 ') + run('עולם', '<w:rtl/>'));
    expect(markRtlRuns(doc(body), WORD)).toBeNull();
  });

  it('אבל נקודה אחרי ריצה שמסומנת עכשיו — יורשת', () => {
    expect(rtlCount(markRtlRuns(doc(para(run('שלום', '<w:rtl/>') + run('עולם') + run('.'))), WORD))).toBe(3);
  });

  it('קוד שדה אינו שכן לטיני — הנקודה שאחרי PAGE יורשת מהעברית', () => {
    const field =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run('1') + '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const out = markRtlRuns(doc(para(run('עמוד ') + field + run('.'))), WORD)!;
    expect(out).toContain('<w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">1</w:t>');
    expect(out).toContain('<w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">.</w:t>');
    expect(out).not.toMatch(/<w:rtl\/><\/w:rPr><w:instrText/);
  });

  it('רשת הביטחון אינה מאחדת חצי מכל צד: שרשרת שמצהירה szCs — אין bCs מסגנון', () => {
    const sheet = readStyleSheet(
      styles(
        '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="David"/><w:sz w:val="20"/><w:szCs w:val="32"/>',
        '<w:style w:type="character" w:styleId="Big"><w:name w:val="b"/><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>',
      ),
    );
    expect(markRtlRuns(doc(para(run('שלום', '<w:rStyle w:val="Big"/>'))), sheet)).toContain(
      '<w:rStyle w:val="Big"/><w:rtl/></w:rPr>',
    );
  });

  it('מעגל basedOn — Word מתעלם מהסגנונות, ולכן אין הדגשה לרשת', () => {
    const sheet = readStyleSheet(
      styles(
        '<w:rFonts w:ascii="Courier New"/>',
        '<w:style w:type="paragraph" w:styleId="A"><w:name w:val="a"/><w:basedOn w:val="B"/><w:rPr><w:b/></w:rPr></w:style>' +
          '<w:style w:type="paragraph" w:styleId="B"><w:name w:val="b"/><w:basedOn w:val="A"/></w:style>',
      ),
    );
    const out = markRtlRuns(doc(para(run('שלום'), '<w:pStyle w:val="A"/><w:bidi/>')), sheet)!;
    expect(out).not.toContain('bCs');
  });

  it('<w:rPr/> שסוגר את עצמו מוחלף, ולא נוספת rPr שנייה', () => {
    const out = markRtlRuns(doc(para('<w:r><w:rPr/><w:t>שלום</w:t></w:r>')), WORD)!;
    expect(out).toContain('<w:r><w:rPr><w:rtl/></w:rPr><w:t>שלום</w:t></w:r>');
    expect(out.match(/<w:rPr/g)).toHaveLength(1);
  });
});

describe('kindOf', () => {
  it('זהה ל-classOf על כל ה-BMP', () => {
    for (let code = 0; code < 0x10000; code += 1) {
      if (code >= 0xd800 && code <= 0xdfff) continue;
      const ch = String.fromCharCode(code);
      if (kindOf(ch) !== classOf(ch)) throw new Error(`U+${code.toString(16)}: ${kindOf(ch)} ≠ ${classOf(ch)}`);
    }
  });
});
