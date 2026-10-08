/**
 * שער: עברית שנכתבה בתוסף יוצאת לקובץ כעברית — `<w:rtl/>` — ומקבלת ב-Word את
 * ההגדרות העבריות של המסמך.
 *
 * מה שדווח: „בפתיחת קובץ Word קיים דרך התוסף, הטקסט שנכתב בתוכנה מוגדר
 * בהגדרות הגופן של Word כטקסט לטיני, ולכן אינו מושפע מהגדרת הגופן העברי.”
 * ראו engine/docx-run-direction.ts — שם המדידה ב-Word — ו-issue 4011 למעלה.
 *
 * ## שני תרחישים
 *
 * 1. **מסמך חדש של התוסף** — ההקלדה יוצאת מוצהרת, וגיליון הסגנונות נושא גופן
 *    עברי זהה ללטיני (Arial), ולא את „Times New Roman (Body CS)” של תבנית
 *    המנוע. אחרת מה שמוצג Arial בעורך יוצא Times New Roman ב-Word (נמדד).
 * 2. **מסמך Word קיים** שבו הלטיני Courier New 10 והעברי David 16, עם פסקה
 *    שנכתבה ב-Word (מוצהרת). פסקה שמוקלדת בעורך חייבת לצאת מוצהרת **ובלי**
 *    לכפות את הלטיני — כלומר לרשת את David 16 כמו הפסקה של Word.
 *
 * ## למה השער מיירט את ההעלאה ואינו משתמש ב-`app.docx()`
 *
 * ‏`window.__qa.exportBase64` קורא ל-`superdoc.export` **ישירות**, ולכן מודד
 * את פלט המנוע ולא את מה שנכתב לקובץ. התיקון יושב ב-`engine/export.ts`. לכן
 * השער מיירט את `fetch(uploadUrl, {method:'PUT'})` וקורא את הבייטים שנשלחו —
 * אותם בייטים בדיוק שהמשתמש מקבל. וזו גם הבקרה: אותו מסמך דרך `app.docx()`
 * חייב לצאת **בלי** ההצהרה על הפסקה שהוקלדה. שער ששתי הקריאות בו נותנות אותו
 * דבר אינו מודד את השלב שלנו אלא את המנוע.
 *
 * ‏`QA_SAVE_DIR` — כששמור, שני הקבצים שנשמרו נכתבים לשם, לפתיחה ב-Word.
 *
 * הרצה:  node scripts/qa/rtl-run-export-qa.mjs   (QA_PORT דורס 9387)
 */
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { openApp, createReport, unzip } from './harness.mjs';
import { buildDocx, numberingXml } from './docx-fixtures.mjs';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const RTL = '<w:bidi/>';
const RLM = '‏';

/** הגדרות Word עבריות: הלטיני Courier New 10, העברי David 16. */
const WORD_STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr>` +
  `<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="David"/><w:sz w:val="20"/><w:szCs w:val="32"/>` +
  `<w:lang w:val="en-US" w:bidi="he-IL"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>${RTL}</w:pPr></w:pPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

/** פסקה שנכתבה ב-Word: כל ריצה עברית מוצהרת, כולל הנקודה. */
const WORD_PARA =
  `<w:p><w:pPr>${RTL}</w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>נכתב בוורד</w:t></w:r>` +
  `<w:r><w:rPr><w:rtl/></w:rPr><w:t>.</w:t></w:r></w:p>`;

/** פסקה עברית שהנקודה שלה ריצה בפני עצמה — הצורה שהמנוע כותב. */
const ENGINE_PARA =
  `<w:p><w:pPr>${RTL}</w:pPr><w:r><w:t xml:space="preserve">שלום עולם</w:t></w:r>` +
  `<w:r><w:t xml:space="preserve">.</w:t></w:r></w:p>`;

/** הדגשה וגודל שהמנוע כותב בצד הלטיני בלבד. */
const BOLD_PARA =
  `<w:p><w:pPr>${RTL}</w:pPr><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr>` +
  `<w:t xml:space="preserve">כותרת מודגשת.</w:t></w:r></w:p>`;

/** משפט עברי עם מילה לועזית, כריצה אחת — כמו שהעורך כותב אותו. */
const MIXED_PARA =
  `<w:p><w:pPr>${RTL}</w:pPr><w:r><w:t xml:space="preserve">מילה Word בעברית.</w:t></w:r></w:p>`;

const LATIN_PARA =
  `<w:p><w:pPr><w:bidi w:val="0"/></w:pPr><w:r><w:t xml:space="preserve">hello world</w:t></w:r>` +
  `<w:r><w:t xml:space="preserve">.</w:t></w:r></w:p>`;

/** הגדרת מספור עם `w:nsid`, כמו בתבנית של המנוע — בשביל בדיקת הייחודיות. */
const NUMBERING = numberingXml().replace(
  '<w:multiLevelType w:val="hybridMultilevel"/>',
  '<w:nsid w:val="587013BA"/><w:multiLevelType w:val="hybridMultilevel"/>',
);

const TYPED = 'טקסט חדש מהתוסף';

const report = createReport('עברית מהתוסף יוצאת מוצהרת', { strict: true });
const app = await openApp({ name: 'rtl-run-export', port: Number(process.env.QA_PORT ?? 9387) });

async function captureUpload() {
  await app.js(
    `(function(){if(window.__origFetch)return;window.__origFetch=window.fetch.bind(window);` +
      `window.fetch=function(url,opts){` +
      `if(opts&&opts.method==='PUT'){` +
      `var body=opts.body;` +
      `var read=body&&body.arrayBuffer?body.arrayBuffer():Promise.resolve(null);` +
      `return read.then(function(buf){` +
      `if(buf){var b=new Uint8Array(buf),s='';for(var i=0;i<b.length;i++)s+=String.fromCharCode(b[i]);` +
      `window.__savedDocx=btoa(s);}` +
      `return new Response('',{status:200});});}` +
      `return window.__origFetch(url,opts);};})()`,
  );
  await app.js(
    `window.__qaHost.replies['fs.beginBinaryWrite']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{writeToken:'wt-rtl',uploadUrl:'https://qa-upload.local/rtl',maxBytes:999999999}})}`,
  );
  await app.js(
    `window.__qaHost.replies['fs.commitUserFileWrite']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{cancelled:false,token:'tok-rtl',name:'rtl.docx',size:1234}})}`,
  );
}

/** Ctrl+S, וההעלאה שנתפסה — או `null`. */
async function save(label) {
  await app.js('window.__savedDocx = ""');
  await app.press('s', 'KeyS', 83, 2, 's');
  for (let i = 0; i < 20; i += 1) {
    await app.sleep(500);
    const saved = await app.js('window.__savedDocx || ""');
    if (saved) {
      const bytes = Buffer.from(saved, 'base64');
      if (process.env.QA_SAVE_DIR) writeFileSync(join(process.env.QA_SAVE_DIR, `${label}.docx`), bytes);
      return unzip(bytes);
    }
  }
  return null;
}

async function openDocx(buffer, name) {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    Buffer.from(buffer).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-${name}',url:${JSON.stringify(dataUrl)},name:${JSON.stringify(name + '.docx')},size:${buffer.length},access:'readwrite'}})}`,
  );
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 2500 });
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.clickSel('.open-browse', 0, { after: 12000 });
  return app.js("document.querySelector('.doc-title-input')?.value");
}

/** הפסקאות של `document.xml`, וכל מה שהשער שואל עליהן. */
function paragraphsOf(xml) {
  const body = xml.slice(xml.indexOf('<w:body'));
  return (body.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []).map((para) => {
    const pPr = (para.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) ?? [''])[0];
    const runs = (para.replace(pPr, '').match(/<w:r[ >][\s\S]*?<\/w:r>/g) ?? []).map((run) => ({
      text: (run.match(/<w:t[^>]*>[\s\S]*?<\/w:t>/g) ?? []).map((t) => t.replace(/<[^>]+>/g, '')).join(''),
      rtl: /<w:rtl\s*\/>/.test(run),
      rPr: (run.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) ?? [''])[0],
    }));
    const text = runs.map((run) => run.text).join('');
    return { text, plain: text.split(RLM).join(''), runs };
  });
}

/** ריצות שיש בהן עברית ואינן מוצהרות. */
const undeclaredHebrew = (para) => para.runs.filter((run) => /[֐-׿]/.test(run.text) && !run.rtl);

try {
  await captureUpload();

  /* ── תרחיש 1: מסמך חדש של התוסף ── */
  await app.caretPara(0);
  await app.type('מסמך חדש', 90);
  await app.sleep(1200);
  const fresh = await save('new-document');
  if (!fresh) {
    report.fail('מסמך חדש — שמירה', 'לא נתפסה שום העלאה');
  } else {
    const styles = fresh['word/styles.xml'] ?? '';
    const defaults = (styles.match(/<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/) ?? [''])[0];
    const csFont = /w:cs="([^"]*)"/.exec(defaults)?.[1];
    const latinFont = /w:ascii="([^"]*)"/.exec(defaults)?.[1];
    if (csFont && csFont === latinFont) report.pass('מסמך חדש — הגופן העברי זהה ללטיני', `${latinFont}`);
    else report.fail('מסמך חדש — הגופן העברי זהה ללטיני', `לטיני=${latinFont}, עברי=${csFont}`);

    const typed = paragraphsOf(fresh['word/document.xml'] ?? '').find((para) => para.plain.includes('מסמך חדש'));
    const engineTyped = paragraphsOf((await app.docx())['word/document.xml'] ?? '').find((para) => para.plain.includes('מסמך חדש'));
    if (engineTyped && undeclaredHebrew(engineTyped).length > 0) report.pass('מסמך חדש — בקרה: המנוע עצמו אינו מצהיר', 'הריצה יוצאת מהמנוע בלי w:rtl');
    else report.fail('מסמך חדש — בקרה: המנוע עצמו אינו מצהיר', 'כבר מוצהרת בפלט המנוע — השער אינו מודד את השלב שלנו');
    if (!typed) report.fail('מסמך חדש — ההקלדה מוצהרת', 'הפסקה שהוקלדה לא נמצאה');
    else if (undeclaredHebrew(typed).length === 0) report.pass('מסמך חדש — ההקלדה מוצהרת', `${typed.runs.length} ריצות`);
    else report.fail('מסמך חדש — ההקלדה מוצהרת', JSON.stringify(typed.runs));
  }

  /* ── תרחיש 2: מסמך Word קיים ── */
  const opened = await openDocx(
    buildDocx({
      body: WORD_PARA + ENGINE_PARA + BOLD_PARA + MIXED_PARA + LATIN_PARA,
      numbering: NUMBERING,
      styles: WORD_STYLES,
    }),
    'rtl-run-export',
  );
  if (opened !== 'rtl-run-export') {
    report.fail('פתיחת המסמך', `שם המסמך אחרי הפתיחה: ${opened}`);
  } else {
    await app.caretPara('שלום עולם');
    await app.press('End', 'End', 35);
    await app.sleep(200);
    // סוף הפסקה שהמנוע כתב ← Enter ← פסקה חדשה. נמדד שכאן ההקלדה יוצאת
    // מהמנוע **בלי** הצהרה: היא יורשת מהשכנה, ולשכנה אין. אחרי פסקה ש-Word
    // כתב המנוע מעביר את `w:rtl` בעצמו — ולכן הבקרה נבנית כאן ולא שם. וזה
    // גם המקור של הדיווח: מסמך מ-15.9 שנכתב בתוסף ונערך ב-Word נושא 1,680
    // ריצות עבריות לא מוצהרות לצד 31 ש-Word כתב.
    await app.press('Enter', 'Enter', 13, 0, '\r');
    await app.sleep(400);
    await app.type(TYPED, 90);
    await app.sleep(600);
    // ורשימה עברית מזיהוי ההקלדה — בשביל בדיקת ה-nsid.
    await app.press('Enter', 'Enter', 13, 0, '\r');
    await app.sleep(400);
    await app.type('א) פריט', 90);
    await app.sleep(1500);

    /* בקרה: פלט המנוע עצמו, לפני השלב שלנו. */
    const engine = paragraphsOf((await app.docx())['word/document.xml'] ?? '');
    const engineTyped = engine.find((para) => para.plain.includes(TYPED));
    if (engineTyped && undeclaredHebrew(engineTyped).length > 0) {
      report.pass('בקרה — המנוע עצמו אינו מצהיר', `${undeclaredHebrew(engineTyped).length} ריצות לא מוצהרות`);
    } else {
      report.fail('בקרה — המנוע עצמו אינו מצהיר', 'הפסקה שהוקלדה כבר מוצהרת בפלט המנוע — השער אינו מודד את השלב שלנו');
    }

    const parts = await save('word-document');
    if (!parts) {
      report.fail('שמירה', 'לא נתפסה שום העלאה — המסמך לא נשמר');
    } else {
      const paragraphs = paragraphsOf(parts['word/document.xml'] ?? '');
      const find = (needle) => paragraphs.find((para) => para.plain.includes(needle));

      /* הדיווח עצמו. */
      const typed = find(TYPED);
      if (!typed) report.fail('הפסקה שהוקלדה מוצהרת עברית', 'לא נמצאה בקובץ');
      else if (undeclaredHebrew(typed).length === 0) report.pass('הפסקה שהוקלדה מוצהרת עברית', `${typed.runs.length} ריצות`);
      else report.fail('הפסקה שהוקלדה מוצהרת עברית', JSON.stringify(undeclaredHebrew(typed)));

      // ולא כופה את הלטיני: בלי `cs`/`szCs` ברמת הריצה היא יורשת David 16.
      if (typed && typed.runs.every((run) => !/w:cs=|<w:szCs/.test(run.rPr)))
        report.pass('הפסקה שהוקלדה יורשת את הגופן העברי של המסמך', 'אין cs/szCs ברמת הריצה');
      else if (typed) report.fail('הפסקה שהוקלדה יורשת את הגופן העברי של המסמך', JSON.stringify(typed.runs.map((r) => r.rPr)));

      const styles = parts['word/styles.xml'] ?? '';
      if (styles.includes('w:cs="David"') && styles.includes('<w:szCs w:val="32"/>'))
        report.pass('גיליון הסגנונות של Word לא נגע', 'David 16 נשאר');
      else report.fail('גיליון הסגנונות של Word לא נגע', 'ההגדרה העברית השתנתה');

      const engineDot = find('שלום עולם');
      if (engineDot && undeclaredHebrew(engineDot).length === 0 && engineDot.runs.every((run) => run.rtl))
        report.pass('ריצה ניטרלית יורשת — הנקודה מוצהרת', JSON.stringify(engineDot.text));
      else report.fail('ריצה ניטרלית יורשת — הנקודה מוצהרת', JSON.stringify(engineDot?.runs));

      const bold = find('כותרת מודגשת');
      const boldPr = bold?.runs[0]?.rPr ?? '';
      if (bold && /<w:bCs\s*\/>/.test(boldPr) && /<w:szCs w:val="36"\/>/.test(boldPr) && bold.runs[0].rtl)
        report.pass('עיצוב ישיר נוסע עם ההצהרה', 'bCs, szCs 36');
      else report.fail('עיצוב ישיר נוסע עם ההצהרה', boldPr);

      const mixed = find('Word');
      const latinRun = mixed?.runs.find((run) => run.text === 'Word');
      if (mixed && mixed.runs.length === 3 && latinRun && !latinRun.rtl && undeclaredHebrew(mixed).length === 0)
        report.pass('ריצה מעורבת מפוצלת — הלועזית נשארת לטינית', mixed.runs.map((r) => `${r.rtl ? 'R' : 'L'}:${r.text}`).join(' | '));
      else report.fail('ריצה מעורבת מפוצלת — הלועזית נשארת לטינית', JSON.stringify(mixed?.runs));

      const latin = find('hello');
      if (latin && latin.runs.every((run) => !run.rtl) && !latin.text.includes(RLM))
        report.pass('הפסקה הלטינית לא נגעה', `${latin.runs.length} ריצות`);
      else report.fail('הפסקה הלטינית לא נגעה', JSON.stringify(latin?.runs));

      const numbering = parts['word/numbering.xml'] ?? '';
      const nsids = [...numbering.matchAll(/<w:abstractNum\b[\s\S]*?<w:nsid w:val="([^"]+)"/g)].map((m) => m[1].toUpperCase());
      if (nsids.length < 2) report.fail('nsid ייחודי', `פחות משתי הגדרות עם nsid (${nsids.length})`);
      else if (new Set(nsids).size === nsids.length) report.pass('nsid ייחודי לכל הגדרת מספור', nsids.join(','));
      else report.fail('nsid ייחודי לכל הגדרת מספור', `כפולים: ${nsids.join(',')}`);
    }
  }
} finally {
  report.print();
  await app.close();
}
