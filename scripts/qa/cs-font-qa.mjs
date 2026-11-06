/**
 * שער: הגופן והגודל העבריים של הסגנונות מוצגים בעורך, ונשמרים כמו שהיו.
 *
 * מה שדווח: Normal עם לטיני „+גוף” (ערכת נושא) ועברי גופן מותקן — נפתח בתוסף
 * בגופן אחר. נמדד על 2.23.0: המנוע מצייר ריצה עברית מהחריץ הלטיני בלבד
 * (superdoc/docx-editor#3994). התיקון: engine/docx-cs-display.ts.
 *
 * מה נבדק, בדפדפן אמיתי ועל ה-dist:
 *   1. המסמך שדווח: העברית מצוירת בגופן ובגודל העבריים של Normal.
 *   2. שמירה (Ctrl+S, הבייטים שנכתבים לקובץ): `styles.xml` זהה למקור בייט-בבייט.
 *   3. טיוטת השחזור (הבייטים של המנוע, בלי השלב היוצא) נפתחת שוב — אותו ציור,
 *      ושמירה ממנה עדיין מחזירה את המקור.
 *   4. שינוי גודל של Normal בעורך — נשמר; הגופן שלא שונה חוזר לערכת הנושא.
 *   5. שינוי גופן של Normal בעורך — נשמר בשלושת החריצים; הגודל חוזר.
 *   6. ריצה עברית עם גופן וגודל עבריים ישירים — מצוירת בהם.
 *   7. בקרה: מסמך שהחריצים בו שווים — אינו נגע, וגופן לטיני נשאר לטיני.
 *
 * הרצה (מ-PowerShell):  node scripts/qa/cs-font-qa.mjs     (QA_PORT דורס 9392)
 */
import { buildSync } from 'esbuild';
import { createReport, openApp, sleep, unzip } from './harness.mjs';
import { buildDocx, zipStored } from './docx-fixtures.mjs';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"';
const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

/** גופן „מותקן” שאינו ברירת מחדל של המנוע, וגופן לטיני שקל להבדיל ממנו. */
const HEBREW = 'Tahoma';
const THEME_LATIN = 'Courier New';

const THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="QA"><a:themeElements><a:clrScheme name="QA"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="000000"/></a:dk2><a:lt2><a:srgbClr val="FFFFFF"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="QA"><a:majorFont><a:latin typeface="${THEME_LATIN}"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hebr" typeface="Times New Roman"/></a:majorFont><a:minorFont><a:latin typeface="${THEME_LATIN}"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hebr" typeface="Arial"/></a:minorFont></a:fontScheme>
<a:fmtScheme name="QA"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

const SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings ${W}><w:defaultTabStop w:val="720"/><w:themeFontLang w:val="en-US" w:bidi="he-IL"/></w:settings>`;

/** כמו ש-Word העברי כותב: ברירות מחדל בערכת נושא, Normal עם עברי מותקן. */
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W} ${W14} ${MC}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="he-IL"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:bidi/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cs="${HEBREW}"/><w:sz w:val="24"/><w:szCs w:val="40"/></w:rPr></w:style>
</w:styles>`;

const heb = (text, rPr = '<w:rtl/>') =>
  `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** חבילה עם ערכת נושא, הגדרות והקשר לערכת הנושא — מה ש-`buildDocx` אינו כותב. */
function wordPackage({ body, styles = STYLES, theme = true }) {
  const parts = unzip(buildDocx({ body, styles }));
  if (theme) {
    parts['word/theme/theme1.xml'] = THEME;
    parts['word/settings.xml'] = SETTINGS;
    parts['word/_rels/document.xml.rels'] = parts['word/_rels/document.xml.rels'].replace(
      '</Relationships>',
      '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
    );
    parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(
      '</Types>',
      '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    );
  }
  return zipStored(parts);
}

const LINES = `(function(){var out=[];var lines=document.querySelectorAll('.superdoc-line');
for(var i=0;i<lines.length;i++){if(lines[i].offsetParent===null)continue;var t=(lines[i].textContent||'').trim();if(!t)continue;
var w=document.createTreeWalker(lines[i],NodeFilter.SHOW_TEXT);var n=w.nextNode();var h=n&&n.parentElement?n.parentElement:lines[i];
var s=getComputedStyle(h);out.push({text:t.slice(0,30),family:s.fontFamily,size:s.fontSize});}return JSON.stringify(out);})()`;

const live = buildSync({
  entryPoints: ['src/engine/live-style-edit.ts'],
  bundle: true,
  platform: 'browser',
  format: 'iife',
  globalName: '__styleLive',
  write: false,
}).outputFiles[0].text;

const defaults = buildSync({
  entryPoints: ['src/engine/doc-style-defaults.ts'],
  bundle: true,
  platform: 'browser',
  format: 'iife',
  globalName: '__docDefaults',
  write: false,
}).outputFiles[0].text;

const report = createReport('גופן וגודל עבריים של סגנונות', { strict: true });
/** פירוט רק לשורה שנכשלה; ההדפסה המיידית — כי דפדפן שנהרג באמצע אינו מגיע ל-`print`. */
const check = (name, ok, detail = '') => {
  console.log(`${ok ? '✓' : '✗'} ${name}`);
  (ok ? report.pass : report.fail)(name, ok ? '' : String(detail ?? ''));
};
const app = await openApp({
  name: 'cs-font-qa',
  port: Number(process.env.QA_PORT ?? 9392),
  // ‏`QA_DIST` — בנייה אחרת, למשל מוטציה שמכבה את ההחזרה כדי לראות את השער אדום.
  ...(process.env.QA_DIST ? { buildDirectory: process.env.QA_DIST } : {}),
  extra: [live, defaults].map((code) => `<script>${code.replaceAll('</script', '<\\/script')}</script>`).join(''),
});

const titleNow = () => app.js("document.querySelector('.doc-title-input')?.value ?? null");
const ask = async (code) => JSON.parse(await app.js(`(async()=>JSON.stringify(await (${code})))()`));

async function open(buffer, name) {
  const dataUrl = `data:${DOCX_MIME};base64,` + Buffer.from(buffer).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:${JSON.stringify(`qa-csf-${name}`)},url:${JSON.stringify(dataUrl)},name:${JSON.stringify(`${name}.docx`)},size:${buffer.length},access:'readwrite'}})}`,
  );
  await app.tab('קובץ');
  if (!(await app.click('פתח קובץ', { after: 600 }))) throw new Error('„פתח קובץ” לא נמצא');
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  if (!(await app.click('עיון בקבצים…', { after: 300 }))) throw new Error('„עיון בקבצים…” לא נמצא');
  for (let waited = 0; waited < 40_000; waited += 250) {
    await sleep(250);
    if (!(await app.exists('.status-load')) && (await titleNow()) === name) {
      await sleep(2_500);
      return;
    }
  }
  throw new Error(`„${name}” לא נפתח בתוך 40 שניות`);
}

async function captureSaves() {
  await app.js(
    `(function(){if(window.__origFetch)return;window.__origFetch=window.fetch.bind(window);` +
      `window.fetch=function(url,opts){if(opts&&opts.method==='PUT'){var body=opts.body;` +
      `var read=body&&body.arrayBuffer?body.arrayBuffer():Promise.resolve(null);` +
      `return read.then(function(buf){if(buf){var b=new Uint8Array(buf),s='';for(var i=0;i<b.length;i++)s+=String.fromCharCode(b[i]);` +
      `window.__savedDocx=btoa(s);}return new Response('',{status:200});});}return window.__origFetch(url,opts);};})()`,
  );
  await app.js(
    `window.__qaHost.replies['fs.beginBinaryWrite']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{writeToken:'wt-csf',uploadUrl:'https://qa-upload.local/csf',maxBytes:999999999}})}`,
  );
  await app.js(
    `window.__qaHost.replies['fs.commitUserFileWrite']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{cancelled:false,token:'tok-csf',name:'csf.docx',size:1234}})}`,
  );
}

/** Ctrl+S, והחלקים שנכתבו לקובץ. */
async function save() {
  // הקיצור נתפס במסמך: אחרי הפתיחה המיקוד נשאר ברצועה, ו-Ctrl+S משם אינו שומר.
  const spot = await app.js(
    `(function(){var l=[...document.querySelectorAll('.superdoc-line')].find(function(x){return x.offsetParent!==null&&(x.textContent||'').trim()});` +
      `if(!l)return '';var r=l.getBoundingClientRect();return JSON.stringify({x:Math.round((r.left+r.right)/2),y:Math.round((r.top+r.bottom)/2)});})()`,
  );
  if (!spot) return null;
  const { x, y } = JSON.parse(spot);
  // מאופס לפני הכול: עריכה מלכלכת את המסמך, והשמירה האוטומטית יכולה לכתוב
  // לפני Ctrl+S — ואז הקיצור מוצא מסמך נקי. כל כתיבה מכאן והלאה היא של המצב הנוכחי.
  await app.js('window.__savedDocx = ""');
  for (let attempt = 0; attempt < 3; attempt += 1) {
    await app.clickAt(x, y);
    await sleep(400);
    // מסמך נקי שיש לו יעד אינו נשמר (`saveNow` מחזיר `clean`, כמו ב-Word). תו
    // שמוקלד ונמחק מסמן אותו כמלוכלך בלי לגעת בתוכן או בסגנונות.
    await app.type('x');
    await sleep(300);
    await app.press('Backspace', 'Backspace', 8);
    await sleep(600);
    await app.press('s', 'KeyS', 83, 2, 's');
    for (let waited = 0; waited < 10_000; waited += 250) {
      await sleep(250);
      const saved = await app.js('window.__savedDocx || ""');
      if (saved) return unzip(Buffer.from(saved, 'base64'));
    }
  }
  return null;
}

const firstLine = async () => JSON.parse(await app.js(LINES))[0] ?? null;
const startsWith = (family, name) => family.replace(/"/g, '').startsWith(name);
const styleNode = (xml, id) => xml.match(new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`))?.[0] ?? '';

/** עריכת סגנון חיה — אותו מסלול שהדיאלוג עובר בו (App.vue, `applyStyleEdit`). */
async function editStyle(format) {
  const xml = unzip(Buffer.from(await app.js('window.__qa.exportBase64()'), 'base64'))['word/styles.xml'];
  const { revision } = await ask('window.__qa.doc().info({})');
  const result = await ask(
    `__styleLive.applyLiveStyleEdit(window.__qa.sd(),${JSON.stringify(xml)},${JSON.stringify({ kind: 'modify', id: 'Normal', format })},${JSON.stringify(revision)})`,
  );
  await sleep(1_500);
  return result.outcome?.ok === true;
}

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1600, height: 1000, deviceScaleFactor: 1, mobile: false });
  await sleep(1_500);
  await captureSaves();

  // 1. המסמך שדווח.
  const reported = wordPackage({ body: heb('שורה עברית בסגנון רגיל') });
  await open(reported, 'reported');
  const line = await firstLine();
  check('העברית מצוירת בגופן העברי של Normal', startsWith(line?.family ?? '', HEBREW), line?.family);
  check('ובגודל העברי של Normal (szCs 40 = 26.67px)', line?.size === '26.6667px', line?.size);

  // 2. שמירה.
  let saved = await save();
  check('שמירה: styles.xml זהה למקור בייט-בבייט', saved?.['word/styles.xml'] === STYLES, saved ? saved['word/styles.xml'] : 'לא נשמר');

  // 3. טיוטה.
  const draft = Buffer.from(await app.js('window.__qa.exportBase64()'), 'base64');
  check('הטיוטה נושאת את המקור', (unzip(draft)['word/styles.xml'] ?? '').includes('owcs:rFonts='));
  await open(draft, 'draft');
  const again = await firstLine();
  check('טיוטה שנפתחה שוב: אותו ציור', startsWith(again?.family ?? '', HEBREW) && again?.size === '26.6667px', `${again?.family} ${again?.size}`);
  saved = await save();
  check('שמירה מהטיוטה: styles.xml זהה למקור', saved?.['word/styles.xml'] === STYLES, saved ? saved['word/styles.xml'] : 'לא נשמר');

  // 4. שינוי גודל בעורך.
  await open(reported, 'size-edit');
  check('עריכת גודל ל-Normal התקבלה', await editStyle({ fontSizePt: 15 }));
  saved = await save();
  const sized = styleNode(saved?.['word/styles.xml'] ?? '', 'Normal');
  check('הגודל החדש נשמר (sz/szCs 30)', sized.includes('<w:sz w:val="30"/>') && sized.includes('<w:szCs w:val="30"/>'), sized);
  check('הגופן שלא שונה חזר ל„+גוף”', sized.includes(`<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cs="${HEBREW}"/>`), sized);
  check('אין סימון בקובץ', !(saved?.['word/styles.xml'] ?? 'owcs').includes('owcs'));

  // 5. שינוי גופן בעורך.
  await open(reported, 'font-edit');
  check('עריכת גופן ל-Normal התקבלה', await editStyle({ fontFamily: 'Arial' }));
  saved = await save();
  const fonted = styleNode(saved?.['word/styles.xml'] ?? '', 'Normal');
  check('הגופן החדש נשמר בשלושת החריצים', /w:ascii="Arial"/.test(fonted) && /w:hAnsi="Arial"/.test(fonted) && /w:cs="Arial"/.test(fonted), fonted);
  check('הגודל שלא שונה חזר (sz 24, szCs 40)', fonted.includes('<w:sz w:val="24"/>') && fonted.includes('<w:szCs w:val="40"/>'), fonted);

  // 6. עיצוב ישיר על ריצה.
  const direct = wordPackage({
    body: heb('ריצה עם גופן עברי ישיר', `<w:rFonts w:hint="cs" w:cs="${HEBREW}"/><w:szCs w:val="36"/><w:rtl/>`),
    styles: STYLES.replace(`w:cs="${HEBREW}"`, 'w:cs="Times New Roman"').replace('<w:szCs w:val="40"/>', '<w:szCs w:val="24"/>'),
  });
  await open(direct, 'direct');
  const run = await firstLine();
  check('ריצה עם cs ישיר מצוירת בו', startsWith(run?.family ?? '', HEBREW), run?.family);
  check('ובגודל szCs הישיר (36 = 24px)', run?.size === '24px', run?.size);

  // 8. ברירות המחדל של המסמך — דרך המנוע (`styles.apply`), המסלול של הדיאלוג.
  // השאלה: האם המנוע שומר את הסימון על `rPr` של `docDefaults` כשהוא כותב אותו.
  const splitDefaults = STYLES.replace('<w:szCs w:val="22"/>', '<w:szCs w:val="32"/>');
  const rPrDefault = (xml) => xml.match(/<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/)?.[0] ?? '';
  const applyDefaults = async (patch) =>
    (await ask(`__docDefaults.applyDocStyleDefaults(window.__qa.sd(),${JSON.stringify(patch)})`))?.ok === true;

  await open(wordPackage({ body: heb('ברירות מחדל'), styles: splitDefaults }), 'defaults-font');
  check('ברירות מחדל: שינוי גופן התקבל', await applyDefaults({ fontFamily: 'Times New Roman' }));
  saved = await save();
  let level = saved ? rPrDefault(saved['word/styles.xml'] ?? '') : 'לא נשמר';
  check('ברירות מחדל, גופן: הגופן החדש בשלושת החריצים', /w:ascii="Times New Roman"/.test(level) && /w:cs="Times New Roman"/.test(level), level);
  check('ברירות מחדל, גופן: הגודל שלא שונה חזר (sz 22, szCs 32)', level.includes('<w:sz w:val="22"/>') && level.includes('<w:szCs w:val="32"/>'), level);
  check('ברירות מחדל, גופן: אין סימון בקובץ', !(saved?.['word/styles.xml'] ?? 'owcs').includes('owcs'));

  await open(wordPackage({ body: heb('ברירות מחדל'), styles: splitDefaults }), 'defaults-size');
  check('ברירות מחדל: שינוי גודל התקבל', await applyDefaults({ fontSizePt: 14 }));
  saved = await save();
  level = saved ? rPrDefault(saved['word/styles.xml'] ?? '') : 'לא נשמר';
  check('ברירות מחדל, גודל: הגודל החדש נשמר (28/28)', level.includes('<w:sz w:val="28"/>') && level.includes('<w:szCs w:val="28"/>'), level);
  check('ברירות מחדל, גודל: הגופן שלא שונה חזר לערכת הנושא', /w:asciiTheme="minorHAnsi"/.test(level) && !/w:ascii="/.test(level), level);
  check('ברירות מחדל, גודל: אין סימון בקובץ', !(saved?.['word/styles.xml'] ?? 'owcs').includes('owcs'));

  // 7. בקרה.
  const sameStyles = STYLES.replace(
    `<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cs="${HEBREW}"/><w:sz w:val="24"/><w:szCs w:val="40"/>`,
    `<w:rFonts w:ascii="${THEME_LATIN}" w:hAnsi="${THEME_LATIN}" w:cs="${THEME_LATIN}"/><w:sz w:val="24"/><w:szCs w:val="24"/>`,
  );
  await open(wordPackage({ body: heb('בקרה') + '<w:p><w:r><w:t>Latin control</w:t></w:r></w:p>', styles: sameStyles }), 'control');
  const control = JSON.parse(await app.js(LINES));
  check('בקרה: החריצים שווים — מצויר כמו קודם', control.every((row) => startsWith(row.family, THEME_LATIN) && row.size === '16px'), JSON.stringify(control));
  saved = await save();
  check('בקרה: styles.xml נשמר כמו שהוא', saved?.['word/styles.xml'] === sameStyles);
} finally {
  await app.close();
}
report.print();
