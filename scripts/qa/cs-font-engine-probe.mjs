/**
 * גשש: האם המנוע עצמו — בלי התוסף ובלי השלב המקדים — מצייר ריצה עברית בגופן
 * ובגודל המורכבים (superdoc/docx-editor#3994). זה מה שיכריע מתי למחוק את
 * `src/engine/docx-cs-display.ts`.
 *
 * רץ מול מאחז vite נקי של הגרסה הנבדקת: `npm i superdoc@<ver> vite@6` בתיקייה
 * נפרדת, דף שחושף `window.openDoc(base64)` ו-`window.lines()`, ו-
 * `optimizeDeps.exclude: ['superdoc', '@superdoc/docx-engine']`.
 *
 * הרצה: ENGINE_URL=http://localhost:5329/ BLANK_URL=file:///…/blank.html node scripts/qa/cs-font-engine-probe.mjs
 *
 * נמדד על 2.29.0: שלושת המקרים הראשונים מצוירים מהחריץ הלטיני (Courier New /
 * Calibri, 16px). התיקון במנוע = Tahoma בשלושתם, וגודל 26.67px בראשון ובשלישי.
 */
import { openPage, sleep } from '../cdp.mjs';
import { buildDocx } from './docx-fixtures.mjs';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const heb = (text, rPr = '<w:rtl/>') =>
  `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const styles = ({ defaults = '<w:sz w:val="24"/>', normal = '' }) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr>${defaults}</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr>${normal}</w:rPr></w:style></w:styles>`;

const CASES = {
  normalSplit: buildDocx({
    body: heb('שורה עברית מוצהרת'),
    styles: styles({ normal: '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Tahoma"/><w:sz w:val="24"/><w:szCs w:val="40"/>' }),
  }),
  themeLatinCsName: buildDocx({
    body: heb('לטיני ערכת נושא ועברי בשם'),
    styles: styles({ normal: '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cs="Tahoma"/>' }),
  }),
  runCsOnly: buildDocx({
    body: heb('ריצה עם cs בלבד', '<w:rFonts w:hint="cs" w:cs="Tahoma"/><w:szCs w:val="40"/><w:rtl/>'),
    styles: styles({ normal: '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>' }),
  }),
  control: buildDocx({
    body: heb('בקרה לטינית'),
    styles: styles({ normal: '<w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma" w:cs="Tahoma"/>' }),
  }),
};

const page = await openPage(process.env.BLANK_URL, { port: Number(process.env.CDP_PORT ?? 9394), label: 'csf-engine' });
try {
  await page.cdp.send('Page.navigate', { url: process.env.ENGINE_URL });
  for (let i = 0; i < 240 && !(await page.cdp.evaluate('window.ready === true').catch(() => false)); i++) await sleep(500);
  console.log('engine:', await page.cdp.evaluate("document.title"));
  for (const [name, buffer] of Object.entries(CASES)) {
    await page.cdp.evaluate(`window.openDoc(${JSON.stringify(Buffer.from(buffer).toString('base64'))})`);
    console.log(name, JSON.stringify(await page.cdp.evaluate('window.lines()')));
  }
} finally {
  page.close();
}
