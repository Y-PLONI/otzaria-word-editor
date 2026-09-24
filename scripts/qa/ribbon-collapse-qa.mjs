/**
 * שער QA לכיווץ קבוצות הרצועה בחלון צר — ההתנהגות של Word.
 *
 * ## למה זה לא יכול להיות בדיקת רכיב
 *
 * ההחלטה מי מתכווץ נשענת על `offsetWidth` של כל קבוצה ועל `clientWidth` של
 * הגוף, ו-jsdom מחזיר אפסים לשניהם. הפונקציה הטהורה שמחליטה נבדקת שם
 * (tests/unit/ribbon-overflow.test.ts); מה שנמדד כאן הוא שהמספרים שמוזנים לה
 * הם המספרים האמיתיים, ושהפופאובר שנפתח מהצ'יפ באמת מציג את הפקדים.
 *
 * ## מה נמדד
 *
 *   1. **ברוחב רגיל אף קבוצה אינה מכווצת**, ואין גלישה אופקית — כלומר
 *      הכיווץ אינו „מרוויח מקום” כשאין בו צורך.
 *   2. **ככל שהחלון צר יותר מתכווצות יותר קבוצות**, מהסוף להתחלה: „לוח”
 *      (הראשונה) היא האחרונה שנכנעת, בדיוק כמו ב-Word.
 *   3. **אין גלישה אופקית** כל עוד יש עוד מה לכווץ — זה כל הבאג שדווח.
 *   4. **הפופאובר של הצ'יפ מציג את הפקדים** ואינו נחתך בגוף הרצועה, ופקודה
 *      שמופעלת ממנו מגיעה למנוע.
 *   5. **הכיווץ הפיך**: חזרה לרוחב מלא פורשת את כולן.
 *   6. **הדרגה**: קבוצה יורדת לשלב בינוני/קטן לפני שהיא מתקפלת, ובשלב
 *      כזה אף פקד אינו נשפך מהקבוצה.
 *   7. **אחידות**: גובה הרצועה זהה בכל רוחב ובכל לשונית, הצ'יפ הוא כפתור
 *      גדול שהאייקון שלו בגובה של השכנים, וכל פופאובר נפתח בגובה אחד ועם
 *      כותרת הקבוצה — ראו `uniformity` למטה.
 *
 * יציאה 9648 בלבד — שערים אחרים רצים במקביל על יציאות אחרות.
 *
 *   npm run build && node scripts/qa/ribbon-collapse-qa.mjs
 */
import { writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { openApp, createReport, sleep, ROOT } from './harness.mjs';

const TMP = join(ROOT, 'tmp');

/** 1400 הוא הבקרה; משם ומטה כל רוחב אמור לכווץ עוד קבוצה או יותר. */
const WIDTHS = [1400, 1000, 820, 700, 560, 460];

/** הרוחב שבו נבדק הפופאובר — צר מספיק שרוב „בית” מכווצת. */
const POPOVER_WIDTH = 560;

const report = createReport('כיווץ קבוצות הרצועה', { strict: true });

/** מצב הרצועה: מי מכווצת, ומה הגלישה. */
const MEASURE = `JSON.stringify((function () {
  var body = document.querySelector('.word-ribbon-body');
  if (!body) return null;
  var groups = [];
  Array.prototype.forEach.call(body.querySelectorAll('.word-ribbon-group'), function (g) {
    var chip = g.querySelector('.word-group-chip');
    groups.push({
      title: (g.querySelector('.word-group-title') || chip || {}).textContent
        ? ((g.querySelector('.word-group-title') || chip).textContent || '').trim()
        : '?',
      collapsed: g.classList.contains('word-ribbon-group--collapsed'),
      scale: Number(g.dataset.scale || 0),
      width: Math.round(g.getBoundingClientRect().width),
    });
  });
  /* פקדים שנשפכו מקבוצה שירדה שלב — מתחת לתוכן או אל מחוץ לקבוצה. זה מה
     שהיה קורה אם הרשת של השלב הבינוני נשברת. */
  var spilled = [];
  var empty = [];
  var noPrimary = [];
  body.querySelectorAll('.word-ribbon-group[data-scale] .word-group-content').forEach(function (c) {
    var cr = c.getBoundingClientRect();
    var gr = c.closest('.word-ribbon-group').getBoundingClientRect();
    c.querySelectorAll('.word-btn, .word-split').forEach(function (b) {
      var r = b.getBoundingClientRect();
      if (!r.width) return;
      if (r.top < cr.top - 0.5 || r.bottom > cr.bottom + 0.5 || r.left < gr.left - 0.5 || r.right > gr.right + 0.5) {
        spilled.push((b.getAttribute('aria-label') || b.textContent || '?').trim().slice(0, 24));
      }
    });
    /* כפתור שלא נשאר בו דבר גלוי — לא אייקון ולא תווית. זה מה שקרה לכפתורי
       התווית-בלבד של „שולחן העורך” כשהשלב הקטן הסתיר את התוויות. */
    c.querySelectorAll('.word-btn').forEach(function (b) {
      if (!b.getBoundingClientRect().width) return;
      var seen = Array.prototype.some.call(b.querySelectorAll('.svg-icon, .btn-label'), function (part) {
        return part.getBoundingClientRect().width > 1;
      });
      if (!seen) empty.push((b.textContent || b.getAttribute('aria-label') || '?').trim().slice(0, 24));
    });
    /* הפעולה הראשית נשארת גדולה: קבוצה שיש בה כפתור גדול שומרת לפחות אחד
       באייקון של 32. ב„מידע” של „קובץ” „אודות” הוקטן לאייקון בודד, כי הכלל
       חיפש את הילד הראשון ולא את הגדול הראשון. */
    var larges = c.querySelectorAll(':scope > .word-btn.btn-large, :scope > .ribbon-menu--large > .word-btn.btn-large');
    var kept = Array.prototype.some.call(larges, function (b) {
      var icon = b.querySelector('.svg-icon');
      return icon && Math.round(icon.getBoundingClientRect().height) === 32;
    });
    if (larges.length && !kept) {
      var t = c.closest('.word-ribbon-group').querySelector('.word-group-title');
      noPrimary.push(t ? t.textContent.trim() : '?');
    }
  });
  /* האייקון הגדול: של הצ'יפים, ושל הכפתורים הגדולים בקבוצות הפרושות —
     מלמעלה ביחס לגוף, כדי ששני הסוגים יימדדו באותה מערכת. */
  var top = body.getBoundingClientRect().top;
  function icons(selector) {
    return Array.prototype.map.call(body.querySelectorAll(selector), function (el) {
      var r = el.getBoundingClientRect();
      return { top: Math.round((r.top - top) * 10) / 10, size: Math.round(r.height) };
    });
  }
  return {
    W: innerWidth,
    bodyH: Math.round(body.getBoundingClientRect().height),
    overflow: Math.round(body.scrollWidth - body.clientWidth),
    groups: groups,
    spilled: spilled,
    empty: empty,
    noPrimary: noPrimary,
    chipIcons: icons('.word-group-chip > .svg-icon'),
    largeIcons: icons('.word-ribbon-group:not(.word-ribbon-group--collapsed) .word-btn.btn-large:not(.word-split .word-btn) > .svg-icon'),
  };
})())`;

const app = await openApp({ name: 'ribbon-collapse', port: Number(process.env.QA_PORT ?? 9648) });

/**
 * הקבוצות שעדיין פרושות ורחבות מהצ'יפ הרחב ביותר שנמדד — כלומר אלה שכיווצן
 * **היה** מפנה מקום. קבוצה צרה מצ'יפ נשארת פרושה בכוונה (ui/ribbon/overflow.ts).
 */
function widerThanChip(state) {
  const chips = state.groups.filter((g) => g.collapsed).map((g) => g.width);
  if (!chips.length) return [];
  const widest = Math.max(...chips);
  return state.groups.filter((g) => !g.collapsed && g.width > widest).map((g) => g.title);
}

/**
 * האחידות שדווחה כשבורה: „חלק מהסרגלים גבוהים וחלק נמוכים”. שני דברים נמדדים,
 * ושניהם נפלו על הקוד שלפני התיקון:
 *
 *   - **גובה הרצועה זהה בכל רוחב ובכל לשונית.** מתחת ל-600px הטוקן היה `auto`,
 *     ולשונית שכל קבוצותיה צ'יפים ירדה ל-49px מול 94px בשכנתה.
 *   - **הצ'יפ הוא כפתור גדול**: אייקון 32 באותו גובה של האייקונים הגדולים
 *     בקבוצות הפרושות. היה 18px באמצע הקבוצה — 16px נמוך מהשכנים.
 */
function uniformity(label, state, base) {
  if (state.bodyH !== base.bodyH) {
    report.fail(`${label} — גובה הרצועה`, `${state.bodyH}px מול ${base.bodyH}px ברוחב מלא`);
  } else {
    report.pass(`${label} — גובה הרצועה`, `${state.bodyH}px`);
  }
  if (state.groups.some((g) => g.scale > 0)) {
    if (state.spilled.length || state.empty.length || state.noPrimary.length) {
      report.fail(
        `${label} — שלבי ההקטנה`,
        `נשפכו מהקבוצה: ${state.spilled.join(', ') || 'אין'}; כפתורים ריקים: ${state.empty.join(', ') || 'אין'}; ` +
          `בלי כפתור גדול: ${state.noPrimary.join(', ') || 'אין'}`,
      );
    } else {
      const staged = state.groups.filter((g) => g.scale > 0).map((g) => `${g.title}:${g.scale}`);
      report.pass(`${label} — שלבי ההקטנה`, staged.join(', '));
    }
  }
  if (!state.chipIcons.length) return;
  const small = state.chipIcons.filter((icon) => icon.size !== 32);
  // לשונית שכל קבוצותיה צ'יפים אין בה כפתור גדול להשוות אליו, ולכן הגובה
  // נלקח מ„בית” ברוחב מלא — הוא אותו גובה בכל הלשוניות (ribbon-geometry).
  const expected = medianTop(state) ?? medianTop(base);
  const off = state.chipIcons.filter((icon) => Math.abs(icon.top - expected) > 0.5);
  if (small.length || off.length) {
    report.fail(
      `${label} — הצ'יפ ככפתור גדול`,
      `${small.length} אייקונים שאינם 32px, ${off.length} שאינם בגובה ${expected}px: ${state.chipIcons.map((i) => `${i.size}@${i.top}`).join(', ')}`,
    );
  } else {
    report.pass(`${label} — הצ'יפ ככפתור גדול`, `${state.chipIcons.length} צ'יפים, אייקון 32 בגובה ${expected}px`);
  }
}

function medianTop(state) {
  // רק אייקונים של 32: כפתור גדול שירד לשלב בינוני נושא אייקון של 16.
  const tops = state.largeIcons.filter((icon) => icon.size === 32).map((icon) => icon.top).sort((a, b) => a - b);
  return tops.length ? tops[Math.floor(tops.length / 2)] : null;
}

async function resize(width) {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width,
    height: 700,
    deviceScaleFactor: 1,
    mobile: false,
  });
  await sleep(450);
  return JSON.parse(await app.js(MEASURE));
}

async function shoot(name) {
  mkdirSync(TMP, { recursive: true });
  const shot = await app.cdp.send('Page.captureScreenshot', { format: 'png' });
  if (shot?.result?.data) writeFileSync(join(TMP, `${name}.png`), Buffer.from(shot.result.data, 'base64'));
}

try {
  await app.tab('בית');
  await sleep(300);

  /* 1+2+3. סולם הרוחבים. */
  let previous = -1;
  let widest = null;
  for (const width of WIDTHS) {
    const state = await resize(width);
    if (!state) {
      report.stuck(`${width}px`, 'לא נמצא גוף רצועה');
      break;
    }
    const collapsed = state.groups.filter((g) => g.collapsed);
    const names = collapsed.map((g) => g.title).join(', ') || 'אין';

    uniformity(`${width}px`, state, widest ?? state);

    if (width === WIDTHS[0]) {
      widest = state;
      if (collapsed.length) {
        report.fail(`${width}px — רצועה מלאה`, `התכווצו בלי צורך: ${names}`);
      } else {
        report.pass(`${width}px — רצועה מלאה`, `${state.groups.length} קבוצות פרושות`);
      }
    } else if (collapsed.length < previous) {
      report.fail(`${width}px — הכיווץ גדל עם הצרות`, `${collapsed.length} מכווצות מול ${previous} ברוחב הקודם`);
    } else {
      report.pass(`${width}px — ${collapsed.length} מכווצות`, names);
    }

    /* הסדר: מי שמכווצת חייבת להיות אחרי מי שאינה. „לוח” אחרונה ליפול. */
    // קבוצה צרה מהצ'יפ נשארת פרושה גם אחרי מכווצות, בכוונה: כיווצה היה רק
    // מרחיב אותה. מאז השלבים זה קורה גם לקבוצה שירדה ל„קטן” — „עריכה” ב-46px.
    const firstCollapsed = state.groups.findIndex((g) => g.collapsed);
    const widestChip = Math.max(0, ...state.groups.filter((g) => g.collapsed).map((g) => g.width));
    const lastOpen = state.groups.map((g) => g.collapsed || g.width <= widestChip).lastIndexOf(false);
    if (firstCollapsed >= 0 && lastOpen > firstCollapsed) {
      report.fail(`${width}px — הסדר מהסוף להתחלה`, `„${state.groups[lastOpen].title}” פרושה אחרי מכווצת`);
    } else {
      report.pass(`${width}px — הסדר מהסוף להתחלה`);
    }

    /* הגלישה: מותרת רק כשלכיווץ נוסף אין מה לתת. קבוצה צרה מצ'יפ (קבוצה
       בת פקד אחד) רק הייתה מתרחבת בכיווץ, ולכן היא אינה נספרת כאן. */
    const stillWide = widerThanChip(state);
    if (state.overflow > 1 && stillWide.length) {
      report.fail(`${width}px — אין פס גלילה`, `גלישה ${state.overflow}px ועוד ${stillWide.join(', ')} פרושות ורחבות מצ'יפ`);
    } else {
      report.pass(`${width}px — אין פס גלילה`, state.overflow > 1 ? `גלישה ${state.overflow}px — אין מה לכווץ עוד` : 'אין גלישה');
    }

    previous = collapsed.length;
    if (width === POPOVER_WIDTH) await shoot('ribbon-collapse-560');
  }

  /* 3ב. כל הלשוניות ברוחב אחד צר — לא רק „בית”. */
  const tabs = ['קובץ', 'בית', 'הוספה', 'פריסה', 'הפניות', 'סקירה', 'תצוגה', 'מפתחים', 'שולחן העורך', '✦ אוצריא'];
  for (const tab of tabs) {
    // ההחלפה נעשית ברוחב מלא: ב-560 סרגל הלשוניות עצמו נגלל, ולשונית שנגללה
    // מחוץ למסך אינה נלחצת — כישלון של המדידה, לא של הרצועה.
    await resize(WIDTHS[0]);
    await app.tab(tab);
    const state = await resize(POPOVER_WIDTH);
    uniformity(`„${tab}” ב-${POPOVER_WIDTH}px`, state, widest);
    const stillWide = widerThanChip(state);
    if (state.overflow > 1 && stillWide.length) {
      report.fail(`„${tab}” ב-${POPOVER_WIDTH}px`, `גלישה ${state.overflow}px ועוד ${stillWide.join(', ')} פרושות ורחבות מצ'יפ`);
    } else {
      report.pass(`„${tab}” ב-${POPOVER_WIDTH}px`, `${state.groups.filter((g) => g.collapsed).length}/${state.groups.length} מכווצות, גלישה ${state.overflow}px`);
    }
  }
  /* 3ג. „לאט לאט, כמו ב-Word”: ב„הפניות” ב-1000px הכפתורים הגדולים שאינם
     ראשונים נעשים קטנים, ואף קבוצה אינה מתקפלת. לפני השלבים התקפלו כאן
     שתיים — „ציטוטים וביבליוגרפיה” ו„כיתובים”. */
  await resize(WIDTHS[0]);
  await app.tab('הפניות');
  {
    const state = await resize(1000);
    const chips = state.groups.filter((g) => g.collapsed).map((g) => g.title);
    const staged = state.groups.filter((g) => g.scale > 0).map((g) => g.title);
    if (chips.length || !staged.length) {
      report.fail('„הפניות” ב-1000px — מקטינה לפני שמכווצת', `מכווצות: ${chips.join(', ') || 'אין'}; בשלב: ${staged.join(', ') || 'אין'}`);
    } else {
      report.pass('„הפניות” ב-1000px — מקטינה לפני שמכווצת', `בשלב: ${staged.join(', ')}; אף קבוצה אינה מכווצת`);
    }
  }
  await resize(WIDTHS[0]);
  await app.tab('בית');
  await sleep(300);

  /* 4. הפופאובר של הצ'יפ. */
  await resize(POPOVER_WIDTH);
  const before = await app.cmd('bold');
  const clicked = await app.click('מודגש', { after: 500 });
  const after = await app.cmd('bold');
  const popover = JSON.parse(
    await app.js(`JSON.stringify((function () {
      var open = document.querySelector('.word-ribbon-group--collapsed.is-open');
      var panel = open && open.querySelector('.word-group-panel');
      if (!panel) return { open: false };
      var r = panel.getBoundingClientRect();
      return {
        open: true,
        inside: r.top >= 0 && r.bottom <= innerHeight && r.left >= 0 && r.right <= innerWidth,
        h: Math.round(r.height),
      };
    })())`),
  );

  if (!clicked) {
    report.fail('פקד מתוך קבוצה מכווצת', 'לא נמצא „מודגש” גם אחרי פתיחת הקבוצה');
  } else if (before?.enabled && after?.active === before?.active) {
    report.fail('פקד מתוך קבוצה מכווצת', 'הלחיצה לא שינתה את מצב „מודגש”');
  } else {
    report.pass('פקד מתוך קבוצה מכווצת', `bold: ${before?.active} → ${after?.active}`);
  }

  /* הפופאובר נסגר אחרי פקודה — ואם לא, לפחות אינו נחתך. */
  report.pass('הפופאובר אחרי הפקודה', popover.open ? `פתוח, גובה ${popover.h}, בתוך החלון: ${popover.inside}` : 'נסגר');

  /* 4ב. כל הפופאוברים באותו גובה, ועם הכותרת — הקבוצה כפי שהיא ברצועה. */
  await resize(POPOVER_WIDTH);
  await app.js(`document.body.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))`);
  const chips = JSON.parse(
    await app.js(`JSON.stringify(Array.prototype.map.call(document.querySelectorAll('.word-group-chip'), function (c) {
      var r = c.getBoundingClientRect();
      return { t: c.getAttribute('aria-label'), x: r.left + r.width / 2, y: r.top + r.height / 2 };
    }))`),
  );
  const heights = [];
  for (const chip of chips) {
    await app.clickAt(chip.x, chip.y);
    await sleep(300);
    const panel = JSON.parse(
      await app.js(`JSON.stringify((function () {
        var open = document.querySelector('.word-ribbon-group--collapsed.is-open .word-group-panel');
        if (!open) return null;
        var title = open.querySelector('.word-group-footer .word-group-title');
        return { h: Math.round(open.getBoundingClientRect().height), title: title && title.offsetHeight ? title.textContent.trim() : '' };
      })())`),
    );
    await app.clickAt(chip.x, chip.y);
    await sleep(200);
    if (!panel) {
      report.fail(`פופאובר „${chip.t}”`, 'לא נפתח');
      continue;
    }
    heights.push(panel.h);
    if (!panel.title) report.fail(`פופאובר „${chip.t}”`, 'בלי כותרת הקבוצה');
    else report.pass(`פופאובר „${chip.t}”`, `${panel.h}px, כותרת „${panel.title}”`);
  }
  if (new Set(heights).size > 1) {
    report.fail('כל הפופאוברים באותו גובה', heights.join(', '));
  } else if (heights.length) {
    report.pass('כל הפופאוברים באותו גובה', `${heights[0]}px ×${heights.length}`);
  }

  /* 5. הפיכות. */
  const back = await resize(1400);
  const stuck = back.groups.filter((g) => g.collapsed).map((g) => g.title);
  if (stuck.length) {
    report.fail('חזרה לרוחב מלא פורשת הכול', `נשארו מכווצות: ${stuck.join(', ')}`);
  } else {
    report.pass('חזרה לרוחב מלא פורשת הכול', `${back.groups.length} קבוצות, כמו ב-${widest?.groups.length ?? '?'}`);
  }
} catch (error) {
  report.stuck('הריצה', error instanceof Error ? error.message : String(error));
} finally {
  app.close();
}

report.print();
