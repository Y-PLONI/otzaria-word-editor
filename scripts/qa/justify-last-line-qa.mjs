/**
 * שער: השורה האחרונה בפסקה מיושרת לשני הצדדים אינה בורחת מהשוליים.
 *
 * העקיפה יושבת ב-`src/engine/justify-last-line.ts`, והפער ב-docs/engine-gaps.md.
 *
 * ## הפיקסטורה
 *
 * פסקאות עבריות של שורה אחת, ‏Times New Roman ‏12pt, אזור טקסט של 9026 twips.
 * כל פסקה היא N פעמים „אבגד” ומילה אחרונה של K פעמים „ה”. הצירופים נבחרו מול
 * Word 2013+ (‏PDF): ב-„כיווץ” Word משאיר את המילה האחרונה בשורה ומכווץ את
 * הרווחים עד השוליים; ב-„טבעי” השורה נכנסת בלי כיווץ; ב-„שבירה” Word מוריד את
 * המילה לשורה שנייה. ופסקה של כמה שורות שהאחרונה בהן חורגת, כדי שהשורות
 * המיושרות שלפניה (שהמנוע נותן להן `word-spacing` משלו) יימדדו כשלא נגעו בהן.
 *
 * ## מה נבדק
 *
 *   1. אף שורה מצוירת אינה יוצאת מהקופסה שלה.
 *   2. השורות שבמקרי „כיווץ” מכווצות — בדיוק עד השוליים — והשורות הטבעיות
 *      והמיושרות אינן מקבלות `word-spacing` מכאן.
 *   3. לחיצה על השורה המכווצת נוחתת היכן שהטקסט מצויר: לחיצה על האות הראשונה
 *      של המילה האחרונה והקלדה מכניסות את התו לפני המילה הזאת.
 *   4. אחרי הקלדה ומחיקה בשורה המכווצת — עדיין אין שורה שבורחת.
 *
 *   npm run check:justify-last-line        (QA_PORT עוקף 9742)
 */
import { openApp, createReport } from './harness.mjs';
import { zipStored } from './docx-fixtures.mjs';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const FONT = 'Times New Roman';
const run = (text) =>
  `<w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/><w:rtl/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (text) =>
  `<w:p><w:pPr><w:bidi/><w:spacing w:after="120" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>${run(text)}</w:p>`;
const words = (n, k) => [...Array(n).fill('אבגד'), 'ה'.repeat(k)].join(' ');

/** הצירופים, וההכרעה של Word עליהם. */
const CASES = [
  { n: 18, k: 5, word: 'טבעי' },
  { n: 18, k: 6, word: 'כיווץ' },
  { n: 18, k: 8, word: 'כיווץ' },
  { n: 19, k: 2, word: 'כיווץ' },
  { n: 19, k: 3, word: 'שבירה' },
];
/** פסקה של שלוש שורות מלאות ושורה אחרונה שחורגת. */
const LONG = `${Array(3).fill(words(18, 4)).join(' ')} ${words(18, 6)}`;

function fixture() {
  const body = [...CASES.map((c) => para(words(c.n, c.k))), para(LONG)].join('');
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  return zipStored({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`,
    'word/settings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
    'word/document.xml': documentXml,
  });
}

/** כל השורות המצוירות: הטקסט, הפסקה, החריגה מהקופסה, ה-`word-spacing` ומי כתב אותו. */
const readLines = () => {
  const out = [];
  for (const line of document.querySelectorAll('.superdoc-line')) {
    const box = line.getBoundingClientRect();
    if (box.width === 0) continue;
    const range = document.createRange();
    range.selectNodeContents(line);
    const rects = [...range.getClientRects()].filter((r) => r.width > 0);
    if (rects.length === 0) continue;
    const inkLeft = Math.min(...rects.map((r) => r.left));
    const inkRight = Math.max(...rects.map((r) => r.right));
    const frag = line.closest('[data-block-id]');
    out.push({
      text: line.textContent,
      block: frag?.getAttribute('data-block-id') ?? '',
      spill: +Math.max(box.left - inkLeft, inkRight - box.right).toFixed(2),
      gapLeft: +(inkLeft - box.left).toFixed(2),
      wordSpacing: line.style.wordSpacing,
      ours: line.dataset.oweSpillSpacing !== undefined,
    });
  }
  return JSON.stringify(out);
};

/**
 * נקודה על הקצה הימני של האות ה-`into` במילה האחרונה של השורה שמסתיימת
 * ב-`" " + word` — כלומר הסמן אמור לנחות אחרי `into` אותיות של המילה. בתוך
 * המילה ולא על הרווח: על גבול של רווח שני ההיסטים מצוירים באותו x.
 */
const pointInLastWord = (word, into) => {
  for (const line of document.querySelectorAll('.superdoc-line')) {
    if (!line.textContent.endsWith(` ${word}`)) continue;
    const walker = document.createTreeWalker(line, NodeFilter.SHOW_TEXT);
    let node;
    let last = null;
    while ((node = walker.nextNode())) if (node.data.includes(' ')) last = node;
    if (!last) return null;
    const at = last.data.lastIndexOf(' ') + 1 + into;
    const range = document.createRange();
    range.setStart(last, at);
    range.setEnd(last, at + 1);
    const r = range.getBoundingClientRect();
    return JSON.stringify({ x: r.right - 1, y: r.top + r.height / 2 });
  }
  return null;
};

const paraTextOf = (word) => {
  for (const frag of document.querySelectorAll('[data-block-id]')) {
    const text = [...frag.querySelectorAll('.superdoc-line')].map((l) => l.textContent).join('');
    if (text.endsWith(` ${word}`)) return text;
  }
  return null;
};

const report = createReport('השורה האחרונה בפסקה מיושרת', { strict: true });
const app = await openApp({ name: 'justify-last-line', port: Number(process.env.QA_PORT ?? 9742) });
const lines = async () => JSON.parse(await app.js(`(${readLines})()`));

try {
  const docx = fixture();
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + docx.toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-justify',url:${JSON.stringify(dataUrl)},name:'justify.docx',size:${docx.length},access:'readwrite'}})}`,
  );
  await app.tab('קובץ');
  if (!(await app.click('פתח קובץ', { after: 500 }))) throw new Error('„פתח קובץ” לא נמצא');
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.click('עיון בקבצים…', { after: 8000 });
  for (let waited = 0; waited < 40_000; waited += 250) {
    await app.sleep(250);
    if (!(await app.exists('.status-load'))) break;
  }
  await app.sleep(2500);

  // 1–2. אחרי הפתיחה.
  const opened = await lines();
  const spilling = opened.filter((l) => l.spill > 0.6);
  if (spilling.length === 0) report.pass('אין שורה שיוצאת מהקופסה', `${opened.length} שורות`);
  else report.fail('אין שורה שיוצאת מהקופסה', spilling.map((l) => `${l.spill}px «${l.text.slice(-12)}»`).join(', '));

  const bare = (text) => text.replace(/\s/g, '');
  const blockLines = (text) => {
    const blocks = [...new Set(opened.map((l) => l.block))];
    const block = blocks.find((b) => bare(opened.filter((l) => l.block === b).map((l) => l.text).join('')) === bare(text));
    return opened.filter((l) => l.block === block);
  };

  for (const c of CASES) {
    const own = blockLines(words(c.n, c.k));
    const name = `${c.n}×אבגד + ${'ה'.repeat(c.k)} (Word: ${c.word})`;
    if (own.length === 0) {
      report.fail(name, 'הפסקה לא נמצאה');
      continue;
    }
    const last = own.at(-1);
    if (c.word === 'כיווץ') {
      if (own.length === 1 && last.ours && Math.abs(last.gapLeft) <= 0.6)
        report.pass(name, `שורה אחת, word-spacing ${last.wordSpacing}, נוגעת בשוליים`);
      else report.fail(name, `${own.length} שורות, ours=${last.ours} ws=«${last.wordSpacing}» gap=${last.gapLeft}`);
    } else if (own.every((l) => !l.ours) && last.wordSpacing === '') {
      report.pass(name, `לא נגעו בה — ${own.length} שורות, רווח ${last.gapLeft}px בשורה האחרונה`);
    } else {
      report.fail(name, own.map((l) => `ours=${l.ours} ws=«${l.wordSpacing}»`).join(' | '));
    }
  }

  const longLines = blockLines(LONG);
  const justified = longLines.slice(0, -1);
  const lastLong = longLines.at(-1);
  if (longLines.length >= 2 && justified.every((l) => !l.ours) && lastLong?.ours && Math.abs(lastLong.gapLeft) <= 0.6)
    report.pass('פסקה ארוכה', `${justified.length} שורות מיושרות של המנוע, האחרונה מכווצת עד השוליים`);
  else
    report.fail(
      'פסקה ארוכה',
      longLines.map((l) => `ours=${l.ours} ws=«${l.wordSpacing}» gap=${l.gapLeft}`).join(' | ') || 'הפסקה לא נמצאה',
    );

  // 3. לחיצה והקלדה: בשורה המכווצת, ובבקרה — שורה טבעית שלא נגעו בה.
  const clickAndType = async (name, word) => {
    const before = await app.js(`(${paraTextOf})(${JSON.stringify(word)})`);
    const point = JSON.parse((await app.js(`(${pointInLastWord})(${JSON.stringify(word)}, 2)`)) ?? 'null');
    if (!before || !point) {
      report.fail(name, 'המילה האחרונה לא נמצאה');
      return false;
    }
    await app.clickAt(point.x, point.y);
    await app.sleep(400);
    await app.type('ש');
    await app.sleep(1500);
    const expected = `${before.slice(0, -word.length)}${word.slice(0, 2)}ש${word.slice(2)}`;
    const after = await app.js(`(${paraTextOf})(${JSON.stringify(word.slice(2))})`);
    // בלי רווחים: התו הנוסף מוריד את המילה לשורה חדשה, והצייר משמיט את הרווח
    // שבנקודת השבירה — כלומר הטקסט המצויר של הפסקה מאבד רווח שבמסמך קיים.
    const found = await app.js(
      `[...document.querySelectorAll('[data-block-id]')].map((f) => [...f.querySelectorAll('.superdoc-line')].map((l) => l.textContent).join('').replace(/\\s/g, '')).includes(${JSON.stringify(expected.replace(/\s/g, ''))})`,
    );
    if (found) report.pass(name, 'התו נכנס אחרי שתי האותיות שנלחצו');
    else {
      const where = await app.js(
        `JSON.stringify([...document.querySelectorAll('[data-block-id]')].map((f) => [...f.querySelectorAll('.superdoc-line')].map((l) => l.textContent).join('')).filter((t) => t.includes('ש')).map((t) => t.slice(t.indexOf('ש') - 8, t.indexOf('ש') + 9)))`,
      );
      report.fail(name, `ציפינו ל«${expected.slice(-14)}»; ה-ש נחת ב: ${where}`);
    }
    return true;
  };

  await clickAndType('לחיצה בשורה טבעית (בקרה)', 'ה'.repeat(5));
  await app.press('Backspace', 'Backspace', 8);
  await app.sleep(1000);
  if (await clickAndType('לחיצה בשורה המכווצת', 'ה'.repeat(8))) {
    // 4. אחרי הקלדה — ואחרי מחיקה.
    const typed = (await lines()).filter((l) => l.spill > 0.6);
    await app.press('Backspace', 'Backspace', 8);
    await app.sleep(1500);
    const erased = (await lines()).filter((l) => l.spill > 0.6);
    if (typed.length === 0 && erased.length === 0) report.pass('הקלדה ומחיקה', 'אין שורה שבורחת');
    else report.fail('הקלדה ומחיקה', `אחרי הקלדה ${typed.length}, אחרי מחיקה ${erased.length}`);
  }
} finally {
  app.close();
}
process.exit(report.print() > 0 ? 1 : 0);
