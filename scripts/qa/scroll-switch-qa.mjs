/**
 * שער רגרסיה למיקום הגלילה כששני מסמכים פתוחים באותו חלון.
 *
 * ## מה נמדד כאן, ולמה השער נולד דווקא עכשיו
 *
 * הדיווח המקורי היה „הוא זוכר איפה אני, וברגע שאני מתחיל לגלול הוא חוזר
 * לראש”. המדידה (Chrome אמיתי, שני מסמכים חיים) הראתה שהחזרה עצמה תקינה —
 * `scrollTop` נכון בכל נקודות הזמן אחרי מעבר טאב — ושהקפיצה מגיעה **בגלגלת
 * הראשונה שאחרי החזרה**: `@superdoc/docx-engine` 2.11.0 החזיק snapshot יחיד של
 * שורש הגלילה לכל המנועים, והזרים אליו ערך שנקרא בזמן שהפאנל של הטאב האחר היה
 * מוסתר. נגדנו את זה ב-`guardPaneScroll`.
 *
 * הכתיבה הזאת נעלמה מהמנוע ב-2.13.0, ונמדדה שוב כנעדרת ב-2.15.0-next.15 —
 * ולכן השומר הוסר (ראו `src/sessions/pane-scroll.ts`). ההסרה הותירה חור: אחרי
 * שהשומר ירד, שום בדיקה בריפו לא נגעה יותר במיקום הגלילה, וגרסת מנוע שתחזיר
 * את האיפוס הייתה נכנסת בלי שאף שער יאדים. השער הזה הוא הכיסוי הזה.
 *
 * ## מה הוא עושה
 *
 * שני מסמכים בני 160 פסקאות בשני טאבים, על ה-dist הארוז ובגלגלות אמיתיות:
 *
 *   1. כל טאב נגלל למיקום משלו (7 גלגלות ו-4 גלגלות), ונרשם מה נזכר.
 *   2. שני מחזורים של מעבר הלוך-חזור; אחרי כל חזרה נקרא `scrollTop` — הוא
 *      חייב להיות בדיוק המיקום שנזכר.
 *   3. מיד אחרי כל חזרה נשלחת **גלגלת אחת** של 100, והמיכל חייב לזוז בדיוק
 *      ב-100. זו הנקודה שבה האיפוס של המנוע נראה, ולא בחזרה עצמה.
 *   4. כל כתיבה ל-`scrollTop` וכל `scrollTo`/`scroll` על מיכלי העורך נרשמות
 *      (מה נתבקש / מה היה / מה יצא), כך שכתיבה של 0 מזוהה בשמה גם אם משהו
 *      מתקן אותה אחר כך. בקרה חיובית מוודאת שהמאזין עצמו חי.
 *
 * כישלון בשורת „גלגלת” בלי כישלון בשורת „חזרה” = המנוע מאפס שוב, כלומר
 * הרגרסיה שבגללה נכתב `guardPaneScroll` חזרה.
 *
 * שימוש:  CHROME=<path> node scripts/qa/scroll-switch-qa.mjs   (QA_PORT דורס 9397)
 */
import { openApp, createReport, sleep } from './harness.mjs';
import { buildDocx } from './docx-fixtures.mjs';

/** פסקה עברית אחת; 160 מהן נותנות מסמך ארוך דיו לגלילה חופשית בשני הכיוונים. */
const para = (i) =>
  `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr>` +
  `<w:t xml:space="preserve">פסקה ${i} — טקסט לגלילה בין שני מסמכים פתוחים באותו חלון</w:t></w:r></w:p>`;
const DOCX = buildDocx({ body: Array.from({ length: 160 }, (_, i) => para(i + 1)).join('') });
const DATA_URL =
  'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
  Buffer.from(DOCX).toString('base64');

/** כמה גלגלות לכל טאב בהתחלה, ומה שווה גלגלת אחת. */
const WHEEL = 120;
const START_WHEELS = { 0: 7, 1: 4 };
/** הגלגלת שנשלחת אחרי כל חזרה — ערך שאינו כפולה של ההתחלה, כדי שסכום מקרי לא יתחזה לתקין. */
const STEP = 100;
const CYCLES = 2;

const report = createReport('מיקום הגלילה של שני מסמכים פתוחים', { strict: true });
const app = await openApp({ name: 'scroll-switch', port: Number(process.env.QA_PORT ?? 9397) });

/** המיכל הגולל של המסמך הפעיל: מיכל עורך מצויר שאינו ממתין ואינו מוסתר. */
const HOST = `Array.prototype.filter.call(document.querySelectorAll('.editor-stack__host:not(.editor-stack__host--pending)'), function(h){ return h.offsetParent !== null; })[0]`;

const hostTop = () => app.js(`(function(){ var h = ${HOST}; return h ? Math.round(h.scrollTop) : null; })()`);
const logLen = () => app.js('window.__scrollLog.length');
const logSince = (mark) => app.js(`JSON.stringify(window.__scrollLog.slice(${mark}))`).then(JSON.parse);
/** כתיבה שמאפסת מיכל שלא היה על אפס — החתימה של „המיקום נמחק”. */
const resets = (entries) => entries.filter((e) => (e.asked === 0 || e.got === 0) && e.was > 0);
const brief = (entries) =>
  entries
    .slice(0, 3)
    .map((e) => `${e.m ?? 'scrollTop'}: ביקש ${e.asked ?? e.arg}, היה ${e.was}, יצא ${e.got}`)
    .join('; ');

/** פתיחת מסמך דרך המסלול האמיתי: „קובץ” → „פתח קובץ” → „עיון בקבצים…”. */
async function openDoc(tag) {
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-${tag}',url:${JSON.stringify(DATA_URL)},name:${JSON.stringify(tag + '.docx')},size:${DOCX.length},access:'readwrite'}})}`,
  );
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 2500 });
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.clickSel('.open-browse', 0, { after: 12000 });
  await app.tab('בית');
}

/** מעבר טאב בלחיצה אמיתית על הלשונית, כמו משתמש. */
async function switchTo(index) {
  const rect = await app
    .js(
      `(function(){ var el = document.querySelectorAll('.word-doctab')[${index}]; if (!el) return 'null';` +
        ` var r = el.getBoundingClientRect(); return JSON.stringify({ x: r.left + r.width/2, y: r.top + r.height/2 }); })()`,
    )
    .then(JSON.parse);
  if (!rect) throw new Error(`טאב ${index} לא נמצא`);
  await app.clickAt(rect.x, rect.y);
  await sleep(1500);
}

/** גלגלות אמיתיות מעל מרכז המיכל הפעיל. */
async function wheel(times, delta = WHEEL) {
  const at = await app
    .js(
      `(function(){ var h = ${HOST}; var r = h.getBoundingClientRect();` +
        ` return JSON.stringify({ x: Math.round(r.left + r.width/2), y: Math.round(r.top + r.height/2) }); })()`,
    )
    .then(JSON.parse);
  await app.cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: at.x, y: at.y, button: 'none', buttons: 0 });
  for (let i = 0; i < times; i += 1) {
    await app.cdp.send('Input.dispatchMouseEvent', {
      type: 'mouseWheel',
      x: at.x,
      y: at.y,
      deltaX: 0,
      deltaY: delta,
      button: 'none',
      buttons: 0,
    });
    await sleep(120);
  }
  await sleep(600);
}

try {
  // חלון קבוע: המיקומים שנמדדים כאן הם מספרים, ולא „בערך”.
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1400,
    height: 1000,
    deviceScaleFactor: 1,
    mobile: false,
  });

  await openDoc('scroll-a');
  await app.js('document.querySelector(".word-doctabs-new")?.click()');
  // עד שהמסמך הריק של „+” מוכן, ולא זמן קבוע: „פתח קובץ” בזמן שהוא עוד נטען
  // נבלע, ו„מסמך ב'” נשאר המסמך הריק — עמוד אחד, שבו 403 הוא כבר התחתית
  // והגלגלת „לא זזה”. נמדד: כשל בכמחצית הריצות על מכונה עמוסה, ו-0 עם ההמתנה.
  for (let waited = 0; waited < 30_000; waited += 250) {
    const ready = await app.js(
      `(window.__otzariaEditors ? window.__otzariaEditors.size : 0) === 2 && !document.querySelector('.editor-stack__host--pending')`,
    );
    if (ready) break;
    await sleep(250);
  }
  await sleep(1000);
  await openDoc('scroll-b');
  await switchTo(0);

  const tabs = await app.js('document.querySelectorAll(".word-doctab").length');
  const editors = await app.js('window.__otzariaEditors ? window.__otzariaEditors.size : -1');
  if (tabs === 2 && editors === 2) report.pass('שני מסמכים פתוחים בשני טאבים', 'שני מנועים חיים');
  else report.fail('שני מסמכים פתוחים בשני טאבים', `${tabs} טאבים, ${editors} מנועים`);

  /*
   * מאזין על כל מיכלי העורך: `scrollTop` הופך ל-accessor שמדווח, ו-`scrollTo`
   * /`scroll` נעטפים. בלעדיו כישלון נראה כמו „המספר לא זז” ולא כמו „מישהו
   * כתב 0” — וההבדל הזה הוא כל ההבדל בין רגרסיית מנוע לבאג אצלנו.
   */
  await app.js(`(function(){
    var log = window.__scrollLog = [];
    var d = Object.getOwnPropertyDescriptor(Element.prototype, 'scrollTop');
    Array.prototype.forEach.call(document.querySelectorAll('.editor-stack__host'), function (h, idx) {
      Object.defineProperty(h, 'scrollTop', { configurable: true,
        get: function () { return d.get.call(this); },
        set: function (v) {
          var was = d.get.call(this); d.set.call(this, v);
          log.push({ host: idx, asked: Math.round(v), was: Math.round(was), got: Math.round(d.get.call(this)) });
        } });
      ['scrollTo', 'scroll'].forEach(function (m) {
        var orig = h[m];
        h[m] = function (a, b) {
          var was = d.get.call(this); var r = orig.apply(this, arguments);
          log.push({ host: idx, m: m, arg: typeof a === 'object' ? JSON.stringify(a) : a + ',' + b,
            was: Math.round(was), got: Math.round(d.get.call(this)) });
          return r;
        };
      });
    });
    return true; })()`);

  // בקרה חיובית: בלעדיה „אפס כתיבות” עשוי להיות מאזין מת ולא מנוע שקט.
  const control = await app.js(
    `(function(){ var h = ${HOST}; var n = window.__scrollLog.length; h.scrollTop = 5; h.scrollTop = 0;` +
      ` var seen = window.__scrollLog.length - n; window.__scrollLog.length = n; return seen; })()`,
  );
  if (control === 2) report.pass('המאזין לכתיבות גלילה חי', 'שתי כתיבות מהדף נרשמו');
  else report.fail('המאזין לכתיבות גלילה חי', `נרשמו ${control} כתיבות במקום 2`);

  /* מיקום פתיחה נפרד לכל מסמך. */
  const remembered = {};
  await wheel(START_WHEELS[0]);
  remembered[0] = await hostTop();
  await switchTo(1);
  const secondStart = await hostTop();
  await wheel(START_WHEELS[1]);
  remembered[1] = await hostTop();

  if (remembered[0] > 0 && remembered[1] > 0 && remembered[0] !== remembered[1] && secondStart === 0)
    report.pass('לכל מסמך מיקום גלילה משלו', `${remembered[0]} מול ${remembered[1]}`);
  else
    report.fail(
      'לכל מסמך מיקום גלילה משלו',
      `מסמך א' ${remembered[0]}, מסמך ב' ${remembered[1]}, מסמך ב' נפתח ב-${secondStart}`,
    );

  /* מעברים חוזרים: חזרה מדויקת, ואז גלגלת אחת שחייבת לזוז בדיוק ב-STEP. */
  const allResets = [];
  for (let cycle = 1; cycle <= CYCLES; cycle += 1) {
    for (const tab of [0, 1]) {
      const doc = tab === 0 ? "א'" : "ב'";
      const markSwitch = await logLen();
      await switchTo(tab);
      const back = await hostTop();
      const markWheel = await logLen();
      await wheel(1, STEP);
      const after = await hostTop();

      const duringSwitch = (await logSince(markSwitch)).slice(0, markWheel - markSwitch);
      const afterWheel = await logSince(markWheel);
      allResets.push(...resets([...duringSwitch, ...afterWheel]));

      if (back === remembered[tab])
        report.pass(`מחזור ${cycle}, מסמך ${doc}: החזרה נוחתת על המיקום שנזכר`, `${back}`);
      else
        report.fail(
          `מחזור ${cycle}, מסמך ${doc}: החזרה נוחתת על המיקום שנזכר`,
          `${back} במקום ${remembered[tab]}` +
            (duringSwitch.length ? ` — ${brief(duringSwitch)}` : ' — בלי שום כתיבה'),
        );

      if (after === back + STEP)
        report.pass(`מחזור ${cycle}, מסמך ${doc}: גלגלת אחת אחרי החזרה מזיזה ב-${STEP}`, `${back} → ${after}`);
      else
        report.fail(
          `מחזור ${cycle}, מסמך ${doc}: גלגלת אחת אחרי החזרה מזיזה ב-${STEP}`,
          `${back} → ${after}, ציפינו ל-${back + STEP}` +
            (afterWheel.length ? ` — ${brief(afterWheel)}` : ' — בלי שום כתיבה'),
        );

      remembered[tab] = after;
    }
  }

  if (allResets.length === 0)
    report.pass('אף אחד לא מאפס את מיכל הגלילה', `${await logLen()} כתיבות בסך הכול`);
  else report.fail('אף אחד לא מאפס את מיכל הגלילה', `${allResets.length} כתיבות אפס — ${brief(allResets)}`);
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
