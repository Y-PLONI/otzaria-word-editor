/**
 * שער ה-QA של אזכור „@”.
 *
 * הכלל היחיד, כמו בשאר השערים: `success: true` אינו הוכחה — מה שנכתב ל-OOXML
 * הוא ההוכחה. קישור ב-DOCX הוא שני דברים שחייבים להסכים: `<w:hyperlink r:id>`
 * ב-document.xml, ו-Relationship עם ה-Target ב-document.xml.rels. אחד בלי
 * השני נראה כמו הצלחה ואינו קישור.
 *
 * השער גם מודד את מה שלא ידענו בכתיבה: האם `hyperlinks.insert` האטומית קיימת
 * ועובדת, או שהמסלול שרץ בפועל הוא `insert` + `wrap` (ראו הערת המודול
 * ב-engine/at-mention-overlay.ts).
 *
 * הרצה:  node scripts/qa/at-mention-qa.mjs
 * היציאה 9610 שמורה לשער הזה בלבד.
 */
import { openApp, createReport } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9610);
const report = createReport('אזכור „@”', { strict: true });

const log = (...a) => console.log(...a);

/** ההתאמות שהמאחז יחזיר במקום אוצריא. */
const HITS = [
  {
    id: 42,
    bookId: 'פסחים',
    bookUid: 'id:42',
    type: 'text',
    title: 'פסחים',
    reference: 'פסחים דף לד',
    index: 1234,
    isPdf: false,
    isSourceLine: true,
    isUserBook: false,
    bookPath: 'ש"ס, בבלי',
  },
  {
    id: 43,
    bookId: 'פסחים',
    bookUid: 'id:43',
    type: 'text',
    title: 'פסחים',
    reference: 'פסחים דף לה',
    index: 1300,
    isPdf: false,
    isSourceLine: false,
    isUserBook: false,
    bookPath: 'ש"ס, בבלי',
  },
];

async function step(name, fn) {
  log(`\n──────── ${name} ────────`);
  try {
    await fn();
  } catch (error) {
    log('!! זרק:', error?.message);
    report.fail(name, `הצעד זרק: ${error?.message}`);
  }
}

/** תמונת מצב של המסמך. */
async function snap(app) {
  const files = (await app.docx()) ?? {};
  return {
    files,
    doc: files['word/document.xml'] ?? '',
    rels: files['word/_rels/document.xml.rels'] ?? '',
  };
}

/** מצב הרשימה הצפה, נקרא מה-DOM האמיתי. */
async function popup(app) {
  return JSON.parse(
    await app.js(`(function () {
      var el = document.querySelector('.otzaria-at-mention');
      if (!el) return JSON.stringify({ open: false });
      var rows = [].slice.call(el.querySelectorAll('[role="option"]'));
      var box = el.getBoundingClientRect();
      return JSON.stringify({
        open: true,
        count: rows.length,
        active: rows.findIndex(function (r) { return r.getAttribute('aria-selected') === 'true'; }),
        texts: rows.map(function (r) { return r.textContent; }),
        rect: { left: box.left, top: box.top, right: box.right, bottom: box.bottom },
        role: el.getAttribute('role'),
        activedescendant: el.getAttribute('aria-activedescendant'),
        inEngineTree: !!el.closest('[class*="superdoc"]'),
      });
    })()`),
  );
}

/** כמה קישורים יש במסמך. */
function countLinks(doc) {
  return (doc.match(/<w:hyperlink\b/g) ?? []).length;
}

/**
 * מנקה את המסמך בין תרחישים.
 *
 * `Meta` ולא `Ctrl`: השער רץ ב-Chrome על macOS, ושם Ctrl+A אינו „בחר הכול”.
 * ניקוי שאינו מנקה הוא הדבר הגרוע ביותר בשער — התרחיש הבא מודד את השאריות
 * של קודמו ונראה שבור בלי שיהיה.
 */
const SELECT_ALL_MODIFIER = process.platform === 'darwin' ? 4 : 2;

async function clearDoc(app) {
  await app.press('a', 'KeyA', 65, SELECT_ALL_MODIFIER);
  await app.sleep(150);
  await app.press('Backspace', 'Backspace', 8, 0);
  await app.sleep(500);
  const { doc } = await snap(app);
  const left = doc.replace(/<[^>]+>/g, '').trim();
  if (left) throw new Error(`הניקוי לא רוקן את המסמך: ${left.slice(0, 80)}`);
  // מחיקת הטקסט משאירה `<w:hyperlink>` ריק, והסמן נשאר בתוכו — הקלדה שם
  // נכשלת ב-hyperlink-nested-unsupported. פסקה חדשה מוציאה אותו החוצה.
  if (countLinks(doc) > 0) {
    await app.press('Enter', 'Enter', 13, 0);
    await app.sleep(400);
  }
}

/** מקליד אזכור וממתין ל-debounce ולתשובה. */
async function typeMention(app, text) {
  await app.type(text);
  await app.sleep(700);
}

const app = await openApp({ name: 'at-mention', port: PORT });

try {
  await app.js(`window.__qaHost.replies['library.resolveRef'] = function (payload) {
    window.__qaResolveCalls = (window.__qaResolveCalls || []).concat([payload]);
    return Promise.resolve({ success: true, data: ${JSON.stringify(HITS)}, error: null });
  };`);

  await step('הרשימה נפתחת על אזכור, ונשלחת ההפניה בלבד', async () => {
    await app.caretPara(0);
    await typeMention(app, 'ראה @פסחים לד');

    const state = await popup(app);
    if (!state.open) return report.fail('פתיחת הרשימה', 'הרשימה לא נפתחה אחרי הקלדת אזכור');
    report.pass('פתיחת הרשימה');

    const calls = JSON.parse(await app.js('JSON.stringify(window.__qaResolveCalls || [])'));
    const last = calls[calls.length - 1];
    if (last?.ref !== 'פסחים לד') {
      report.fail('ההפניה שנשלחה', `נשלח ${JSON.stringify(last?.ref)} במקום „פסחים לד”`);
    } else {
      report.pass('ההפניה שנשלחה');
    }

    if (state.count !== HITS.length) {
      report.fail('מספר ההצעות', `${state.count} שורות במקום ${HITS.length}`);
    } else {
      report.pass('מספר ההצעות');
    }

    // הרשימה חייבת לשבת מחוץ לעץ ה-DOM של המנוע — ראו engine-boundaries.
    if (state.inEngineTree) {
      report.fail('הרשימה מחוץ למנוע', 'הרשימה נמצאת בתוך עץ ה-DOM של SuperDoc');
    } else {
      report.pass('הרשימה מחוץ למנוע');
    }

    if (state.role !== 'listbox' || !state.activedescendant) {
      report.fail('ARIA', `role=${state.role} activedescendant=${state.activedescendant}`);
    } else {
      report.pass('ARIA');
    }

    const { rect } = state;
    const onScreen =
      rect.left >= 0 && rect.top >= 0 && rect.right <= 4000 && rect.bottom <= 4000 && rect.right > rect.left;
    if (!onScreen) {
      report.fail('מיקום הרשימה', `הרשימה מחוץ למסך: ${JSON.stringify(rect)}`);
    } else {
      report.pass('מיקום הרשימה');
    }
  });

  await step('Enter כותב קישור אמיתי ל-OOXML', async () => {
    await app.press('Enter', 'Enter', 13, 0);
    await app.sleep(900);

    const after = await popup(app);
    if (after.open) report.fail('סגירה אחרי בחירה', 'הרשימה נשארה פתוחה');
    else report.pass('סגירה אחרי בחירה');

    const { doc, rels } = await snap(app);

    if (!/<w:hyperlink[^>]*r:id="([^"]+)"/.test(doc)) {
      return report.fail('הקישור ב-document.xml', 'אין <w:hyperlink r:id> במסמך');
    }
    report.pass('הקישור ב-document.xml');

    const relId = doc.match(/<w:hyperlink[^>]*r:id="([^"]+)"/)[1];
    const rel = new RegExp(`Id="${relId}"[^>]*Target="([^"]+)"`).exec(rels);
    if (!rel) {
      return report.fail('ה-Relationship', `אין Relationship עבור ${relId} ב-document.xml.rels`);
    }
    // ה-Target עובר escaping של XML; & הוא &amp; בקובץ.
    const target = rel[1].replace(/&amp;/g, '&');
    if (target !== 'otzaria://open/book/42?index=1234') {
      report.fail('יעד הקישור', `Target=${target}`);
    } else {
      report.pass('יעד הקישור');
    }

    // הטקסט הנראה הוא ההפניה שנפתרה, וה-@ נעלם.
    const text = doc.replace(/<[^>]+>/g, '');
    if (!text.includes('פסחים דף לד')) {
      report.fail('טקסט הקישור', `לא נמצא „פסחים דף לד” ב: ${text.slice(0, 200)}`);
    } else {
      report.pass('טקסט הקישור');
    }
    if (text.includes('@')) {
      report.fail('ה-@ הוחלף', `נשאר „@” בטקסט: ${text.slice(0, 200)}`);
    } else {
      report.pass('ה-@ הוחלף');
    }
    // המילה שלפני האזכור נשארת במקומה.
    if (!text.includes('ראה')) {
      report.fail('הטקסט שלפני נשמר', `„ראה” נעלם: ${text.slice(0, 200)}`);
    } else {
      report.pass('הטקסט שלפני נשמר');
    }
  });

  await step('חצים בוחרים הצעה אחרת', async () => {
    await clearDoc(app);
    await typeMention(app, '@פסחים');

    const before = await popup(app);
    if (!before.open) return report.fail('פתיחה שנייה', 'הרשימה לא נפתחה');

    await app.press('ArrowDown', 'ArrowDown', 40, 0);
    await app.sleep(200);
    const moved = await popup(app);
    if (moved.active !== 1) {
      return report.fail('חץ למטה', `הפריט הפעיל הוא ${moved.active} במקום 1`);
    }
    report.pass('חץ למטה');

    await app.press('Tab', 'Tab', 9, 0);
    await app.sleep(900);

    const { doc, rels } = await snap(app);
    // האחרון ולא הראשון: קישורים ריקים משאריות התרחישים הקודמים נשארים במסמך.
    const ids = [...doc.matchAll(/<w:hyperlink[^>]*r:id="([^"]+)"/g)].map((m) => m[1]);
    if (!ids.length) return report.fail('כתיבה אחרי חץ', 'לא נכתב קישור');
    const rel = new RegExp(`Id="${ids[ids.length - 1]}"[^>]*Target="([^"]+)"`).exec(rels);
    const target = (rel?.[1] ?? '').replace(/&amp;/g, '&');
    if (target !== 'otzaria://open/book/43?index=1300') {
      report.fail('ההצעה השנייה נבחרה', `Target=${target}`);
    } else {
      report.pass('ההצעה השנייה נבחרה');
    }
  });

  await step('Escape סוגר בלי לכתוב', async () => {
    await clearDoc(app);
    const linksBefore = countLinks((await snap(app)).doc);
    await typeMention(app, '@פסחים');
    if (!(await popup(app)).open) return report.fail('פתיחה שלישית', 'הרשימה לא נפתחה');

    await app.press('Escape', 'Escape', 27, 0);
    await app.sleep(300);
    if ((await popup(app)).open) {
      report.fail('Escape סוגר', 'הרשימה נשארה פתוחה');
    } else {
      report.pass('Escape סוגר');
    }

    // ספירה ולא נוכחות: מחיקת הטקסט משאירה `<w:hyperlink>` ריק במסמך, ולכן
    // „יש קישור” אינו אומר „נכתב קישור עכשיו”.
    const { doc } = await snap(app);
    if (countLinks(doc) > linksBefore) {
      report.fail('Escape אינו כותב', 'נוסף קישור למרות Escape');
    } else {
      report.pass('Escape אינו כותב');
    }
    // הטקסט שהוקלד נשאר כפי שהוא — Escape מוותר על ההצעה, לא על ההקלדה.
    const text = doc.replace(/<[^>]+>/g, '');
    if (!text.includes('@פסחים')) {
      report.fail('הטקסט נשאר אחרי Escape', `הטקסט: ${text.slice(0, 120)}`);
    } else {
      report.pass('הטקסט נשאר אחרי Escape');
    }
  });

  await step('כתובת דוא"ל אינה פותחת רשימה', async () => {
    await clearDoc(app);
    await typeMention(app, 'dev@example');
    if ((await popup(app)).open) {
      report.fail('דוא"ל אינו טריגר', 'הרשימה נפתחה על כתובת דוא"ל');
    } else {
      report.pass('דוא"ל אינו טריגר');
    }
  });

  await step('אות שימוש כן פותחת רשימה', async () => {
    await clearDoc(app);
    await typeMention(app, 'כמובא ב@פסחים');
    const state = await popup(app);
    if (!state.open) {
      report.fail('אות שימוש היא טריגר', '„ב@” לא פתח רשימה');
    } else {
      report.pass('אות שימוש היא טריגר');
    }
    await app.press('Escape', 'Escape', 27, 0);
    await app.sleep(200);
  });

  await step('מדידה: איזה מסלול כתיבה קיים במנוע', async () => {
    const shape = JSON.parse(
      await app.js(`(function () {
        var doc = window.__otzariaEditor && window.__otzariaEditor.superdoc
          && window.__otzariaEditor.superdoc.activeEditor
          && window.__otzariaEditor.superdoc.activeEditor.doc;
        if (!doc) return JSON.stringify({ found: false });
        var h = doc.hyperlinks || {};
        return JSON.stringify({
          found: true,
          insert: typeof h.insert,
          wrap: typeof h.wrap,
          patch: typeof h.patch,
          remove: typeof h.remove,
        });
      })()`),
    );
    if (!shape.found) {
      report.fail('צורת ה-API', 'לא נמצאה ידית עורך');
    } else {
      log(`   hyperlinks: insert=${shape.insert} wrap=${shape.wrap} patch=${shape.patch}`);
      // אחד משני המסלולים חייב להיות זמין, אחרת הפיצ'ר אינו יכול לכתוב.
      const canWrite = shape.insert === 'function' || shape.wrap === 'function';
      if (!canWrite) report.fail('צורת ה-API', 'אין insert ואין wrap');
      else report.pass('צורת ה-API');
    }
  });

  await step('מדידה: הקישור נחסם לציור במנוע', async () => {
    // ה-DomPainter מסנן href מול רשימת סכימות קבועה (http/https/mailto/tel/sms)
    // ואינו מקבל קונפיגורציה בנקודת הקריאה, ולכן `otzaria://` נחסם. הקישור
    // עצמו נכתב ל-DOCX תקין — הפער הוא בלחיצה בתוך העורך בלבד.
    const pageLog = (await app.log()) ?? [];
    const blocked = pageLog.filter((l) => /Blocked potentially unsafe URL/.test(l));
    if (blocked.length) {
      report.partial(
        'לחיצה בתוך העורך',
        'המנוע חוסם otzaria:// לציור; הקישור ב-DOCX תקין. ראו docs/engine-gaps.md',
      );
    } else {
      // אם המנוע הפסיק לחסום — הפער נסגר, ואפשר להסיר את ההסתייגות מהתיעוד.
      report.pass('לחיצה בתוך העורך');
    }
  });

  await step('לא הצטבר רעש', async () => {
    const [status, messages, pageLog] = await Promise.all([app.status(), app.messages(), app.log()]);
    const bad = [];
    if (status?.error) bad.push(`status=${status.text}`);
    const errs = (messages ?? []).filter((m) => m.method === 'ui.showError');
    if (errs.length) bad.push(`showError=${errs.map((m) => m.text).join(' | ')}`);
    const noisy = (pageLog ?? []).filter(
      (l) =>
        !/DevTools|Download the Vue/i.test(l) &&
        // אזהרות של המאחז עצמו: הדמה אינו מממש את שתי המתודות האלה.
        !/reader\.addContextMenuItem|fonts\.listInstalled/.test(l) &&
        // פער מדוד במנוע, לא כשל של הפיצ'ר — נבדק בצעד משלו למטה.
        !/Blocked potentially unsafe URL/.test(l) &&
        // פער ותיק ומתועד של ה-projection במסמך רב-פסקאות (ראו
        // engine/text-search.ts ו-docs/superdoc-2.10-review.md). הוא צץ כאן
        // רק בגלל הפסקאות שהניקוי מוסיף, ואינו נוגע לאזכור.
        !/projection-incomplete|typing-mutation: engine pass failed/.test(l),
    );
    if (noisy.length) bad.push(`log=${noisy.join(' | ')}`);
    if (bad.length) report.fail('ללא רעש', bad.join('; '));
    else report.pass('ללא רעש');
  });
} finally {
  app.close();
}

report.print();
process.exit(process.exitCode ?? 0);
