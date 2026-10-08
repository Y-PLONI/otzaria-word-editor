import { describe, expect, it } from 'vitest';
import { mirrorComplexScript } from '../../src/engine/docx-cs-mirror';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const MATH_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

function doc(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS} ${MATH_NS}><w:body>${body}</w:body></w:document>`;
}

/** ריצה אחת עם ה-`rPr` שנמסרת. */
function run(rPr: string, text = 'שלום'): string {
  return `<w:p><w:pPr/><w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;
}

/** ה-`rPr` שיצאה, או `null`. */
function propsOf(xml: string): string | null {
  const out = mirrorComplexScript(xml);
  if (out === null) return null;
  return (out.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) ?? [])[0] ?? '';
}

const RTL = '<w:rtl/>';

describe('mirrorComplexScript — מה שממורה', () => {
  it('הדגשה', () => {
    expect(propsOf(doc(run(`<w:b/>${RTL}`)))).toBe(`<w:rPr><w:b/><w:bCs/>${RTL}</w:rPr>`);
  });

  it('נטייה', () => {
    expect(propsOf(doc(run(`<w:i/>${RTL}`)))).toBe(`<w:rPr><w:i/><w:iCs/>${RTL}</w:rPr>`);
  });

  it('גודל — הערך מועתק', () => {
    expect(propsOf(doc(run(`<w:sz w:val="36"/>${RTL}`)))).toBe(
      `<w:rPr><w:sz w:val="36"/><w:szCs w:val="36"/>${RTL}</w:rPr>`,
    );
  });

  it('גופן — כמאפיין על rFonts הקיים', () => {
    expect(propsOf(doc(run(`<w:rFonts w:ascii="David" w:hAnsi="David"/>${RTL}`)))).toBe(
      `<w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>${RTL}</w:rPr>`,
    );
  });

  it('גופן שמוצהר ב-hAnsi בלבד', () => {
    expect(propsOf(doc(run(`<w:rFonts w:hAnsi="David"/>${RTL}`)))).toBe(
      `<w:rPr><w:rFonts w:hAnsi="David" w:cs="David"/>${RTL}</w:rPr>`,
    );
  });

  it('‏rFonts שאינו סוגר את עצמו', () => {
    expect(propsOf(doc(run(`<w:rFonts w:ascii="David"></w:rFonts>${RTL}`)))).toBe(
      `<w:rPr><w:rFonts w:ascii="David" w:cs="David"></w:rFonts>${RTL}</w:rPr>`,
    );
  });

  /** הצורה שנמדדה בפועל מהמנוע: `b` לפני `rFonts`, ו-`i` לפני `b`. */
  it('הסדר שהמנוע כותב — כל תאום נצמד לשותפו', () => {
    const engine = `<w:i/><w:b/><w:rFonts w:ascii="David" w:hAnsi="David"/>${RTL}`;
    expect(propsOf(doc(run(engine)))).toBe(
      `<w:rPr><w:i/><w:iCs/><w:b/><w:bCs/><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>${RTL}</w:rPr>`,
    );
  });

  it('ארבעתם יחד', () => {
    const all = `<w:rFonts w:ascii="David"/><w:b/><w:i/><w:sz w:val="28"/>${RTL}`;
    expect(propsOf(doc(run(all)))).toBe(
      `<w:rPr><w:rFonts w:ascii="David" w:cs="David"/><w:b/><w:bCs/><w:i/><w:iCs/>` +
        `<w:sz w:val="28"/><w:szCs w:val="28"/>${RTL}</w:rPr>`,
    );
  });

  it('‏`w:rtl` בלי ערך הוא דלוק', () => {
    expect(propsOf(doc(run(`<w:b/><w:rtl w:val="1"/>`)))).toBe(
      '<w:rPr><w:b/><w:bCs/><w:rtl w:val="1"/></w:rPr>',
    );
  });

  it('שתי ריצות, כל אחת לעצמה', () => {
    const out = mirrorComplexScript(doc(run(`<w:b/>${RTL}`) + run(`<w:i/>${RTL}`, 'עוד')));
    expect(out).not.toBeNull();
    expect((out!.match(/<w:bCs\/>/g) ?? []).length).toBe(1);
    expect((out!.match(/<w:iCs\/>/g) ?? []).length).toBe(1);
  });
});

/**
 * ‏`w:b` ו-`w:i` הם `ST_OnOff`, ולכן מה שנמרה אינו נוכחות הדגל אלא הערך שהוא
 * נושא. תאום שנכתב דלוק ליד שותף מכובה הופך „לא מודגש” מפורש למודגש — היפוך,
 * ולא השלמה. נמדד שזה מה ש-Word עצמו כותב: מכל 115 ה-`rPr` שבקורפוס שיש בהן
 * `b` מכובה לצד `bCs`, ב-115 גם ה-`bCs` מכובה.
 */
describe('mirrorComplexScript — מצב ההדלקה עובר אל התאום', () => {
  it('‏`w:b` מכובה מוליד `bCs` מכובה', () => {
    expect(propsOf(doc(run(`<w:b w:val="0"/>${RTL}`)))).toBe(
      `<w:rPr><w:b w:val="0"/><w:bCs w:val="0"/>${RTL}</w:rPr>`,
    );
  });

  it('שלוש הצורות של „כבוי” — `0`, `false`, `off`', () => {
    for (const off of ['0', 'false', 'off']) {
      expect(propsOf(doc(run(`<w:b w:val="${off}"/>${RTL}`)))).toBe(
        `<w:rPr><w:b w:val="${off}"/><w:bCs w:val="${off}"/>${RTL}</w:rPr>`,
      );
    }
  });

  it('נטייה מכובה', () => {
    expect(propsOf(doc(run(`<w:i w:val="false"/>${RTL}`)))).toBe(
      `<w:rPr><w:i w:val="false"/><w:iCs w:val="false"/>${RTL}</w:rPr>`,
    );
  });

  it('ערך דלוק מפורש עובר גם הוא כמות שהוא', () => {
    expect(propsOf(doc(run(`<w:b w:val="1"/>${RTL}`)))).toBe(
      `<w:rPr><w:b w:val="1"/><w:bCs w:val="1"/>${RTL}</w:rPr>`,
    );
  });

  it('מרכאות בודדות במקור — התאום נכתב בכפולות', () => {
    expect(propsOf(doc(run(`<w:b w:val='0'/>${RTL}`)))).toBe(
      `<w:rPr><w:b w:val='0'/><w:bCs w:val="0"/>${RTL}</w:rPr>`,
    );
  });

  /** הצורה שנמדדה בשלושה מסמכים אמיתיים: ביטול הדגשה על ריצה עברית מוצהרת. */
  it('הצורה שנמדדה בקורפוס — ביטול הדגשה לצד גודל וגופן', () => {
    const measured = `<w:b w:val="0"/><w:rFonts w:cs="Carizma"/><w:sz w:val="28"/><w:szCs w:val="28"/>${RTL}`;
    expect(propsOf(doc(run(measured)))).toBe(
      `<w:rPr><w:b w:val="0"/><w:bCs w:val="0"/><w:rFonts w:cs="Carizma"/>` +
        `<w:sz w:val="28"/><w:szCs w:val="28"/>${RTL}</w:rPr>`,
    );
  });

  it('הרצה שנייה על מכובה אינה מוסיפה דבר', () => {
    const once = mirrorComplexScript(doc(run(`<w:b w:val="0"/><w:i w:val="off"/>${RTL}`)));
    expect(once).not.toBeNull();
    expect(mirrorComplexScript(once!)).toBeNull();
  });
});

describe('mirrorComplexScript — מה שאינו נגע', () => {
  it('ריצה בלי הצהרת rtl', () => {
    expect(mirrorComplexScript(doc(run('<w:b/>')))).toBeNull();
  });

  it('‏`w:rtl` מוצהר מכובה — הכותב אמר „זו אינה עברית”', () => {
    expect(mirrorComplexScript(doc(run('<w:b/><w:rtl w:val="0"/>')))).toBeNull();
    expect(mirrorComplexScript(doc(run('<w:b/><w:rtl w:val="false"/>')))).toBeNull();
  });

  it('התאום כבר שם', () => {
    expect(mirrorComplexScript(doc(run(`<w:b/><w:bCs/>${RTL}`)))).toBeNull();
    expect(mirrorComplexScript(doc(run(`<w:sz w:val="36"/><w:szCs w:val="20"/>${RTL}`)))).toBeNull();
    expect(
      mirrorComplexScript(doc(run(`<w:rFonts w:ascii="David" w:cs="Arial"/>${RTL}`))),
    ).toBeNull();
  });

  it('אין צד לטיני למרות', () => {
    expect(mirrorComplexScript(doc(run(RTL)))).toBeNull();
    expect(mirrorComplexScript(doc(run(`<w:rFonts w:cs="David"/>${RTL}`)))).toBeNull();
  });

  it('‏`rPrChange` — היסטוריה של שינוי מסומן, לא העיצוב החי', () => {
    const withChange =
      `<w:p><w:pPr/><w:r><w:rPr><w:b/>${RTL}` +
      `<w:rPrChange w:id="1" w:author="a" w:date="2020-01-01T00:00:00Z"><w:rPr><w:b/><w:rtl/></w:rPr></w:rPrChange>` +
      `</w:rPr><w:t>שלום</w:t></w:r></w:p>`;
    expect(mirrorComplexScript(doc(withChange))).toBeNull();
  });

  it('ה-`rPr` שבתוך `rPrChange` אינה נקראת בפני עצמה', () => {
    const only =
      `<w:p><w:pPr/><w:r><w:rPr>${RTL}` +
      `<w:rPrChange w:id="1" w:author="a" w:date="2020-01-01T00:00:00Z"><w:rPr><w:b/><w:rtl/></w:rPr></w:rPrChange>` +
      `</w:rPr><w:t>שלום</w:t></w:r></w:p>`;
    expect(mirrorComplexScript(doc(only))).toBeNull();
  });

  it('‏`rPr` של סימן הפסקה', () => {
    const mark = `<w:p><w:pPr><w:rPr><w:b/>${RTL}</w:rPr></w:pPr><w:r><w:t>שלום</w:t></w:r></w:p>`;
    expect(mirrorComplexScript(doc(mark))).toBeNull();
  });

  it('ריצה של נוסחה — שם מקומי זהה, מרחב שמות אחר', () => {
    const math = `<w:p><w:pPr/><m:oMath><m:r><m:rPr><m:b/><m:rtl/></m:rPr><m:t>שלום</m:t></m:r></m:oMath></w:p>`;
    expect(mirrorComplexScript(doc(math))).toBeNull();
  });

  it('מה שבתוך הערה אינו ריצה', () => {
    expect(mirrorComplexScript(doc(`<!-- ${run(`<w:b/>${RTL}`)} -->`))).toBeNull();
  });

  it('חלק בלי שום הצהרת rtl — יציאה מוקדמת', () => {
    expect(mirrorComplexScript(doc(run('<w:b/><w:i/>')))).toBeNull();
  });

  it('אין הצהרת מרחב שמות', () => {
    expect(mirrorComplexScript('<document><body><p><r><rPr><b/><rtl/></rPr></r></p></body></document>')).toBeNull();
  });
});

describe('mirrorComplexScript — הערובות', () => {
  it('paired property tags receive sibling mirrors, not nested children', () => {
    const input = doc(run(`<w:b w:val="0"></w:b><w:i></w:i><w:sz w:val="28"></w:sz>${RTL}`));
    const expected = input
      .replace('</w:b>', '</w:b><w:bCs w:val="0"/>')
      .replace('</w:i>', '</w:i><w:iCs/>')
      .replace('</w:sz>', '</w:sz><w:szCs w:val="28"/>');
    expect(mirrorComplexScript(input)).toBe(expected);
    expect(mirrorComplexScript(expected)).toBeNull();
  });

  it('הפלט הוא הקלט ועוד האיברים שנוספו — שום בייט לא נמחק ולא זז', () => {
    const input = doc(
      run(`<w:rFonts w:ascii="David"/><w:b/><w:i/><w:sz w:val="28"/>${RTL}`) + run('<w:b/>', 'לטינית'),
    );
    const out = mirrorComplexScript(input)!;
    expect(out).not.toBeNull();
    const added = ['<w:bCs/>', '<w:iCs/>', '<w:szCs w:val="28"/>', ' w:cs="David"'];
    let stripped = out;
    for (const piece of added) stripped = stripped.replace(piece, '');
    expect(stripped).toBe(input);
  });

  it('הרצה שנייה אינה מוסיפה דבר', () => {
    const once = mirrorComplexScript(doc(run(`<w:b/><w:i/><w:sz w:val="36"/>${RTL}`)));
    expect(once).not.toBeNull();
    expect(mirrorComplexScript(once!)).toBeNull();
  });

  it('ערך עם מרכאות מוגן', () => {
    const out = propsOf(doc(run(`<w:rFonts w:ascii='Da"vid'/>${RTL}`)));
    expect(out).toContain('w:cs="Da&quot;vid"');
  });

  it('‏`rPr` שלא נסגרה אינה מייצרת פלט פגום', () => {
    const broken = doc(`<w:p><w:pPr/><w:r><w:rPr><w:b/>${RTL}<w:t>שלום</w:t></w:r></w:p>`);
    const out = mirrorComplexScript(broken);
    if (out !== null) expect(out.replace('<w:bCs/>', '')).toBe(broken);
  });
});

describe('mirrorComplexScript — גופן ערכת נושא', () => {
  it('asciiTheme ישיר עובר כ-cstheme באותה ערכה (נמצא ב-QA)', () => {
    expect(propsOf(doc(run(`<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>${RTL}`)))).toBe(
      `<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:cstheme="majorHAnsi"/>${RTL}</w:rPr>`,
    );
  });

  it('ערכת נושא גוברת על שם באותו תג', () => {
    expect(propsOf(doc(run(`<w:rFonts w:ascii="Arial" w:asciiTheme="minorHAnsi"/>${RTL}`)))).toContain('w:cstheme="minorHAnsi"');
  });

  it('cstheme קיים — אין תאום נוסף', () => {
    expect(mirrorComplexScript(doc(run(`<w:rFonts w:ascii="Arial" w:cstheme="minorBidi"/>${RTL}`)))).toBeNull();
  });
});
