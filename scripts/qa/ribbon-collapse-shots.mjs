/**
 * גשש: צילום הרצועה בכל לשונית בכמה רוחבים, עם מדידה של כל קבוצה — גובה
 * הקבוצה, גובה הצ'יפ, ומיקום האייקון שלו מול האייקונים של הכפתורים הגדולים.
 *
 *   npm run build && node scripts/qa/ribbon-collapse-shots.mjs [out-dir]
 */
import { writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { openApp, sleep, ROOT } from './harness.mjs';

const OUT = process.argv[2] ?? join(ROOT, 'tmp', 'ribbon-shots');
mkdirSync(OUT, { recursive: true });

const WIDTHS = (process.env.WIDTHS ?? '1400,1100,900,700,560').split(',').map(Number);
const TABS = (process.env.TABS ?? 'קובץ,בית,הוספה,פריסה,הפניות,סקירה,תצוגה,מפתחים,שולחן העורך,✦ אוצריא').split(',');

const MEASURE = `JSON.stringify((function () {
  var body = document.querySelector('.word-ribbon-body');
  var b = body.getBoundingClientRect();
  var out = { W: innerWidth, bodyTop: Math.round(b.top), bodyH: Math.round(b.height), overflow: body.scrollWidth - body.clientWidth, groups: [] };
  body.querySelectorAll('.word-ribbon-group').forEach(function (g) {
    var r = g.getBoundingClientRect();
    var chip = g.querySelector('.word-group-chip');
    var row = { t: ((g.querySelector('.word-group-title') || chip || {}).textContent || '').trim(), x: Math.round(r.left), w: Math.round(r.width), h: Math.round(r.height), top: Math.round(r.top - b.top), s: Number(g.dataset.scale || 0) };
    if (chip) {
      var c = chip.getBoundingClientRect();
      var ic = chip.querySelector('svg').getBoundingClientRect();
      row.chip = { h: Math.round(c.height), top: Math.round(c.top - b.top), iconTop: Math.round(ic.top - b.top), iconH: Math.round(ic.height) };
    } else {
      var big = g.querySelector('.btn-large svg');
      if (big) { var bi = big.getBoundingClientRect(); row.bigIconTop = Math.round(bi.top - b.top); row.bigIconH = Math.round(bi.height); }
    }
    out.groups.push(row);
  });
  return out;
})())`;

const app = await openApp({ name: 'ribbon-shots', port: Number(process.env.QA_PORT ?? 9661) });
const log = [];

async function resize(width) {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width, height: 700, deviceScaleFactor: 1, mobile: false });
  await sleep(450);
}

async function shoot(name) {
  const clip = JSON.parse(await app.js(`JSON.stringify((function(){var r=document.querySelector('.word-ribbon-container').getBoundingClientRect();return {x:0,y:0,width:innerWidth,height:Math.ceil(r.bottom)+4,scale:1};})())`));
  const shot = await app.cdp.send('Page.captureScreenshot', { format: 'png', clip });
  if (shot?.result?.data) writeFileSync(join(OUT, `${name}.png`), Buffer.from(shot.result.data, 'base64'));
}

try {
  for (const tab of TABS) {
    await resize(1400);
    await app.tab(tab);
    for (const width of WIDTHS) {
      await resize(width);
      const state = JSON.parse(await app.js(MEASURE));
      log.push({ tab, ...state });
      await shoot(`${tab.replace(/[^\p{L}\p{N}]+/gu, '')}-${width}`);
    }
  }
  if (process.env.POPOVER) {
    await resize(1400);
    await app.tab('בית');
    await resize(Number(process.env.POPOVER));
    const chips = JSON.parse(await app.js(`JSON.stringify(Array.from(document.querySelectorAll('.word-group-chip')).map(function(c){var r=c.getBoundingClientRect();return {t:c.textContent.trim(),x:r.left+r.width/2,y:r.top+r.height/2};}))`));
    for (const [i, chip] of chips.entries()) {
      await app.clickAt(chip.x, chip.y);
      await sleep(350);
      const size = await app.js(`(function(){var p=document.querySelector('.word-ribbon-group--collapsed.is-open .word-group-panel');if(!p)return 'closed';var r=p.getBoundingClientRect();return Math.round(r.width)+'x'+Math.round(r.height)+' @'+Math.round(r.top);})()`);
      log.push({ popover: chip.t, size });
      const shot = await app.cdp.send('Page.captureScreenshot', { format: 'png', clip: { x: 0, y: 0, width: Number(process.env.POPOVER), height: 320, scale: 1 } });
      if (shot?.result?.data) writeFileSync(join(OUT, `popover-${i}.png`), Buffer.from(shot.result.data, 'base64'));
      await app.clickAt(chip.x, chip.y);
      await sleep(250);
    }
  }
} finally {
  app.close();
}

writeFileSync(join(OUT, 'measure.json'), JSON.stringify(log, null, 1));
console.log(`נכתב ל-${OUT}`);
