/**
 * שער החזרה: „סגרתי את התוסף וחזרתי אליו — הכול כמו שהיה”.
 *
 * ## מה נמדד, ולמה דווקא כך
 *
 * זו הבדיקה היחידה שיכולה לענות על השאלה באמת, מפני שהיא מריצה את **המסמך
 * חוזר דרך הדיסק**: מה שנכתב בהפעלה אחת מיוצא ל-DOCX, עובר base64, נשמר,
 * ונטען בהפעלה שנייה למנוע אמיתי. כל שלב בשרשרת הזאת דורש workers אמיתיים
 * ומנוע DOCX אמיתי, ולכן שום בדיקת jsdom אינה יכולה להחליף אותה — היא הייתה
 * מאשרת בירוק שרשרת שאף פעם לא רצה.
 *
 * המסמך שנבדק הוא **מסמך חדש שמעולם לא נשמר**: אין לו קובץ ואין לו token,
 * ולכן הטיוטה היא הדבר היחיד בעולם שמחזיק את מה שנכתב. זה גם המקרה שבו
 * הסגירה הייתה מוחקת עבודה עד היום.
 *
 * ## איך „סגירה” מדומה
 *
 * שתי טעינות דף נפרדות, ולא רענון: מה ששורד ביניהן הוא **רק** מה שהתוסף
 * כתב דרך ה-SDK. הכפיל מקליט כל `storage.set` וכל `fs.writeFile` לאובייקט
 * אחד, הדף השני נבנה עם האובייקט הזה כזרע — וכל מה שלא נכתב דרך ה-SDK פשוט
 * אינו קיים שם. כך אי אפשר לעבור את השער „במקרה”, דרך זיכרון שנשאר.
 *
 * שש השאלות שהשער עונה עליהן:
 *   1. הטקסט שנכתב — חזר?
 *   2. מקום הסמן — חזר?
 *   3. גודל התצוגה — חזר?
 *   4. הלשונית ברצועה — חזרה?
 *   5. **כל** הטאבים שהיו פתוחים — חזרו לרצועה?
 *   6. הטאב שלא נטען בעלייה — נפתח עם התוכן שלו כשעוברים אליו?
 *
 * שתי האחרונות נוספו אחרי שהתגלה שסגירת התוסף עם שני מסמכים פתוחים החזירה
 * אחד: הרשומה כן נשמרה במלואה, אבל העלייה קראה ממנה את הטאב הפעיל בלבד.
 * שתיהן חייבות להיות כאן ולא בבדיקת jsdom: ה„טאב השני” נשען על טיוטה שעברה
 * ייצוא DOCX אמיתי ועל מנוע שני שנפתח בפועל.
 *
 *   npm run build && npm run check:session
 */
import { readFileSync, writeFileSync, existsSync, rmSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage, requireChrome, sleep } from './cdp.mjs';

const DIST = join(dirname(fileURLToPath(import.meta.url)), '..', 'dist');
const INDEX = join(DIST, 'index.html');
const PAGE_ONE = join(DIST, '__session-probe-1.html');
const PAGE_TWO = join(DIST, '__session-probe-2.html');

/** כמה להמתין למסמך לפני שמוותרים. נדיב: המנוע פורס 16MB בדרך. */
const OPEN_MS = 60_000;

/**
 * מתי ה-boot מגיע, מתחילת הדף. באוצריא נמדד ~1.2 שניות, אחרי שהמעטפת כבר
 * עלתה; כאן מעט יותר, כדי שגם מכונה מהירה תגיע לקריאות לפני ה-boot.
 */
const BOOT_DELAY_MS = 2_000;

/** הטקסט שנכתב בהפעלה הראשונה. עברית בכוונה — היא מה שעובר את השרשרת. */
const TYPED = 'שחזור';

/** הטקסט שנכתב בטאב השני. שונה מהראשון — כך „איזה מסמך חזר” אינו ניחוש. */
const TYPED_SECOND = 'שני';

/** גודל התצוגה שנקבע בהפעלה הראשונה. */
const ZOOM = 150;

/** הלשונית שנבחרת בהפעלה הראשונה. */
const TAB = 'הפניות';

if (!existsSync(INDEX)) {
  console.error('dist/index.html אינו קיים — הריצו npm run build תחילה');
  process.exit(1);
}
requireChrome();

/**
 * הכפיל של אוצריא. `seed` הוא מה ששרד מההפעלה הקודמת; בהפעלה הראשונה הוא ריק.
 *
 * הכפיל מממש בדיוק את מה שהתוסף קורא לו: `storage.*` ו-`fs.*` של המרחב
 * הפרטי. `fs.resolveFileUrl` נכשל בכוונה — אין קובץ בתרחיש הזה, וכשל שלו הוא
 * מה שאוצריא הייתה מחזירה.
 */
function instrument(seed) {
  return `
<script>
(function () {
  var BOOT = {
    plugin: { id: 'session-probe', version: '0' },
    app: { version: '9.9.9', platform: 'session-probe', language: 'he' },
    theme: { mode: 'light', colorScheme: {}, typography: {} },
    connectivity: { isOnline: false }, permissions: []
  };

  var STATE = ${JSON.stringify(seed)};
  STATE.storage = STATE.storage || {};
  STATE.files = STATE.files || {};

  var P = window.__probe = { state: STATE, log: [], calls: [] };
  ['error', 'warn'].forEach(function (level) {
    var original = console[level];
    console[level] = function () {
      try { P.log.push(level + ': ' + Array.prototype.map.call(arguments, String).join(' ').slice(0, 300)); } catch (e) {}
      return original.apply(console, arguments);
    };
  });

  function ok(data) { return Promise.resolve({ success: true, data: data, error: null }); }
  function fail(code) {
    return Promise.resolve({ success: false, data: null, error: { code: code, message: code } });
  }

  var REAL = {
    call: function (method, payload) {
      payload = payload || {};
      P.calls.push(method);
      switch (method) {
        case 'app.getInfo': return ok(BOOT.app);
        case 'app.getTheme': return ok(BOOT.theme);
        case 'storage.get':
          return ok(Object.prototype.hasOwnProperty.call(STATE.storage, payload.key)
            ? JSON.parse(STATE.storage[payload.key]) : null);
        case 'storage.set':
          STATE.storage[payload.key] = JSON.stringify(payload.value === undefined ? null : payload.value);
          return ok(true);
        case 'storage.remove':
          delete STATE.storage[payload.key];
          return ok(true);
        case 'fs.writeFile':
          STATE.files[payload.path] = String(payload.content == null ? '' : payload.content);
          return ok({ path: payload.path, size: STATE.files[payload.path].length });
        case 'fs.readFile':
          if (!Object.prototype.hasOwnProperty.call(STATE.files, payload.path)) return fail('error.not_found');
          return ok({ path: payload.path, encoding: 'base64', content: STATE.files[payload.path] });
        case 'fs.deleteEntry':
          delete STATE.files[payload.path];
          return ok(true);
        case 'fs.stat':
          return ok(Object.prototype.hasOwnProperty.call(STATE.files, payload.path)
            ? { exists: true, path: payload.path, type: 'file', size: STATE.files[payload.path].length }
            : { exists: false });
        default:
          return fail('error.not_supported');
      }
    },
    on: function () {}, off: function () {}
  };

  /*
   * כמו אוצריא (plugin_sdk_scripts.dart): עד _boot ה-SDK הוא stub שדוחה כל
   * קריאה ב-„not ready yet”, וה-boot מגיע אחרי שהבאנדל כבר רץ. כפיל שמאתחל
   * מיד עבר ירוק על 7.0, שקרא את הרשומה לפני ה-boot ודרס אותה ברשומה ריקה.
   */
  var booted = false;
  window.Otzaria = {
    call: function (method, payload) {
      if (booted) return REAL.call(method, payload);
      P.calls.push('not-ready:' + method);
      return Promise.reject(new Error('Otzaria SDK not ready yet'));
    },
    on: REAL.on, off: REAL.off,
    _boot: function () {
      booted = true;
      window.Otzaria._booted = true;
      window.dispatchEvent(new CustomEvent('plugin.boot', { detail: BOOT }));
    }
  };
  setTimeout(function () { window.Otzaria._boot(); }, ${BOOT_DELAY_MS});

  /**
   * מה שהתוסף הצליח לשחזר, כפי שנראה מבחוץ.
   *
   * הטקסט נקרא מרשימת הפסקאות ולא מ-ui.document.getText: האחרון מוגדר כמי
   * שמחזיר null "כשאינו זמין", ובפועל הוא החזיר מחרוזת ריקה על מסמך שהטקסט
   * בו נראה על המסך — כלומר קריאה שאינה יכולה לשמש שער.
   */
  P.readback = async function () {
    var session = window.__otzariaEditor;
    var ui = session && session.ui;
    var text = '';
    try {
      var listed = await session.superdoc.activeEditor.doc.blocks.list({ includeText: true });
      var parts = (listed.blocks || []).map(function (b) { return b.text || b.textPreview || ''; });
      text = parts.join(' / ');
    } catch (e) { text = 'שגיאה: ' + e; }

    var zoom = null;
    try { zoom = ui && ui.zoom.getSnapshot().value; } catch (e) {}

    var caret = null;
    try {
      var slice = ui && ui.selection.getSnapshot();
      var target = slice && slice.selectionTarget;
      if (target && target.start && target.start.kind === 'text') {
        caret = { blockId: target.start.blockId, offset: target.start.offset };
      }
    } catch (e) {}

    var tab = '';
    var active = document.querySelector('.word-tab-btn.active');
    if (active) tab = active.textContent.trim();

    return JSON.stringify({ text: text, zoom: zoom, caret: caret, tab: tab });
  };

  /**
   * שורת הטקסט הראשונה **הנראית**. עם כמה טאבים פתוחים יש כמה פאנלים ב-DOM
   * ורק אחד מהם גלוי (display:none על השאר, App.vue), ולכן querySelector
   * סתמי היה מחזיר שורה של טאב מוסתר — מלבן אפס, כלומר לחיצה לשום מקום.
   *
   * בלי גרש-אחורי בהערה הזאת בכוונה: הכל נמצא בתוך template literal.
   */
  P.visibleLine = function () {
    var lines = document.querySelectorAll('.superdoc-line, .superdoc-fragment');
    for (var i = 0; i < lines.length; i++) {
      var rect = lines[i].getBoundingClientRect();
      if (rect.width > 0 && rect.height > 0) {
        return JSON.stringify({
          x: Math.round(rect.x + Math.min(20, rect.width / 2)),
          y: Math.round(rect.y + rect.height / 2),
        });
      }
    }
    return 'none';
  };

  /** רצועת הטאבים כפי שהיא נראית: כותרת לכל טאב, ומי מהם פעיל. */
  P.tabs = function () {
    var out = [];
    var tabs = document.querySelectorAll('.word-doctab');
    for (var i = 0; i < tabs.length; i++) {
      var title = tabs[i].querySelector('.word-doctab-title');
      out.push({
        title: title ? title.textContent.trim() : '',
        active: tabs[i].classList.contains('active'),
      });
    }
    return JSON.stringify(out);
  };

  /** „+” ברצועת הטאבים. */
  P.newTab = function () {
    var button = document.querySelector('.word-doctabs-new');
    if (!button) return false;
    button.click();
    return true;
  };

  /** מעבר לטאב לפי מיקומו. */
  P.selectDocumentTab = function (index) {
    var tabs = document.querySelectorAll('.word-doctab');
    if (!tabs[index]) return false;
    tabs[index].click();
    return true;
  };

  /** בוחרת לשונית ברצועה לפי התווית שלה. */
  P.selectTab = function (label) {
    var tabs = document.querySelectorAll('.word-tab-btn');
    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].textContent.trim() === label) { tabs[i].click(); return true; }
    }
    return false;
  };
})();
</script>
`;
}

/** בונה עמוד בדיקה מה-dist, עם הכפיל מוזרק מיד אחרי ה-latch של ה-boot. */
function buildPage(path, seed) {
  const html = readFileSync(INDEX, 'utf8');
  const afterLatch = html.indexOf('</script>') + '</script>'.length;
  if (afterLatch <= '</script>'.length) {
    console.error('לא נמצא ה-latch ב-dist/index.html — אין לאן להזריק את הכפיל');
    process.exit(1);
  }
  writeFileSync(path, html.slice(0, afterLatch) + instrument(seed) + html.slice(afterLatch));
  return `file:///${path.split('\\').join('/')}`;
}

/** ממתינה למסמך פתוח, או נכשלת. */
async function waitForDocument(cdp) {
  for (let waited = 0; waited < OPEN_MS; waited += 500) {
    await sleep(500);
    const ready = await cdp.evaluate(
      '!!window.__otzariaEditor && !document.getElementById("otzaria-splash")',
    );
    if (ready) return;
  }
  throw new Error('המסמך לא נפתח בזמן סביר');
}

/**
 * מציבה סמן בשורת טקסט. לחיצה על העמוד עצמו אינה מספיקה: העורך קולט הקלדה
 * ב-textarea מוסתר, ולחיצה על שטח ריק ממקדת אותו בלי לפתור יעד בחירה.
 */
async function clickIntoText(cdp) {
  const line = await cdp.evaluate('window.__probe.visibleLine()');
  if (line === 'none') throw new Error('לא נמצאה שורת טקסט במסמך — אין לאן למקם סמן');
  const at = JSON.parse(line);

  await cdp.send('Input.dispatchMouseEvent', {
    type: 'mousePressed', x: at.x, y: at.y, button: 'left', buttons: 1, clickCount: 1,
  });
  await sleep(60);
  await cdp.send('Input.dispatchMouseEvent', {
    type: 'mouseReleased', x: at.x, y: at.y, button: 'left', buttons: 0, clickCount: 1,
  });
  await sleep(800);
}

async function typeText(cdp, text) {
  for (const ch of text) {
    await cdp.send('Input.dispatchKeyEvent', {
      type: 'keyDown', key: ch, text: ch, unmodifiedText: ch,
    });
    await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: ch });
    await sleep(90);
  }
}

const failures = [];
const notes = [];
let saved = null;
let restored = null;
let firstLog = [];
let secondLog = [];

// ---------------------------------------------------------------------------
// הפעלה ראשונה: כותבים, מכוונים, ו„עוזבים”.
// ---------------------------------------------------------------------------
try {
  const { cdp, close } = await openPage(buildPage(PAGE_ONE, {}), { label: 'session-1' });
  try {
    await waitForDocument(cdp);
    await sleep(2_500);

    await clickIntoText(cdp);
    await typeText(cdp, TYPED);
    await sleep(1_500);

    // גודל התצוגה דרך המנוע — בדיוק כמו „התאמה לרוחב החלון” שאינה באה
    // מאיתנו. זה מה שההאזנה ב-App.vue אמורה לתפוס.
    await cdp.evaluate(`window.__otzariaEditor.ui.zoom.set(${ZOOM})`);
    await sleep(500);

    if (!(await cdp.evaluate(`window.__probe.selectTab(${JSON.stringify(TAB)})`))) {
      failures.push(`הלשונית „${TAB}” לא נמצאה ברצועה`);
    }
    await sleep(500);

    const before = JSON.parse(await cdp.evaluate('window.__probe.readback()'));
    if (!before.text.includes(TYPED)) {
      failures.push(`ההקלדה לא נכנסה למסמך בהפעלה הראשונה (התקבל ${JSON.stringify(before.text.slice(0, 40))})`);
    }
    notes.push(`הפעלה 1: טקסט ${JSON.stringify(before.text.trim().slice(0, 30))}, זום ${before.zoom}, לשונית ${before.tab}, סמן ${JSON.stringify(before.caret)}`);

    // ---------------------------------------------------------------------
    // טאב שני, ובו טקסט משלו. מכאן והלאה השאלה היא „כמה מסמכים חוזרים”.
    // ---------------------------------------------------------------------
    if (!(await cdp.evaluate('window.__probe.newTab()'))) {
      failures.push('לא נמצא הכפתור „+” ברצועת הטאבים');
    } else {
      await sleep(3_000);
      await clickIntoText(cdp);
      await typeText(cdp, TYPED_SECOND);
      await sleep(1_500);

      const secondBack = JSON.parse(await cdp.evaluate('window.__probe.readback()'));
      if (!secondBack.text.includes(TYPED_SECOND)) {
        failures.push(
          `ההקלדה לא נכנסה לטאב השני (התקבל ${JSON.stringify(secondBack.text.slice(0, 40))})`,
        );
      }
      const openTabs = JSON.parse(await cdp.evaluate('window.__probe.tabs()'));
      if (openTabs.length !== 2) failures.push(`ציפינו לשני טאבים פתוחים, נמדדו ${openTabs.length}`);
      notes.push(`הפעלה 1: ${openTabs.length} טאבים, טקסט בשני ${JSON.stringify(secondBack.text.trim().slice(0, 30))}`);

      /**
       * חזרה לטאב הראשון לפני היציאה, ובכוונה.
       *
       * ארבע השאלות המקוריות — טקסט, סמן, זום, לשונית — נשאלות בהפעלה 2 על
       * הטאב **הפעיל**, וגודל התצוגה הוא תכונה של המסמך (`documentViewFor`).
       * בלי החזרה הזאת הן היו נמדדות על המסמך השני, שלא כוונו בו לא זום ולא
       * סמן, ו„הזום לא חזר” היה מדווח על מסמך שמעולם לא היה בו זום אחר.
       */
      if (!(await cdp.evaluate('window.__probe.selectDocumentTab(0)'))) {
        failures.push('המעבר חזרה לטאב הראשון נכשל');
      }
      await sleep(1_500);
    }

    // „המשתמש עזב”. אותו אירוע בדיוק שאוצריא מייצרת כשהיא מפרקת את הדף.
    await cdp.evaluate('window.dispatchEvent(new Event("pagehide"))');

    // הכתיבה אסינכרונית — ייצוא, base64, גשר. ממתינים עד שהטיוטה על „הדיסק”.
    for (let waited = 0; waited < 20_000; waited += 250) {
      await sleep(250);
      const done = await cdp.evaluate(
        'Object.keys(window.__probe.state.files).length > 0 && !!window.__probe.state.storage.session',
      );
      if (done) break;
    }

    saved = JSON.parse(await cdp.evaluate('JSON.stringify(window.__probe.state)'));
    firstLog = JSON.parse(await cdp.evaluate('JSON.stringify(window.__probe.log)'));
  } finally {
    close();
  }
} catch (error) {
  failures.push(`ההפעלה הראשונה נכשלה: ${error.message}`);
}

/**
 * רשומת המסמך הפעיל מתוך רשומה שמורה. הצורה (`version: 2`, session-state.ts)
 * היא אוסף `documents` עם `activeId` — לא `document`/`caret` שטוחים כמו
 * לפני ריבוי המסמכים. השער הזה בודק מסמך יחיד, ולכן הפעיל הוא תמיד היחיד.
 */
function activeDocumentEntry(record) {
  return record.documents?.find((entry) => entry.id === record.activeId) ?? null;
}

if (saved) {
  const files = Object.keys(saved.files ?? {});
  if (files.length === 0) failures.push('לא נכתבה טיוטה כלל — אין מה לשחזר');
  if (!saved.storage?.session) failures.push('לא נכתבה רשומת הפעלה');
  else {
    const record = JSON.parse(saved.storage.session);
    const entry = activeDocumentEntry(record);
    notes.push(
      `נשמר: טיוטה ${files.join(', ') || '—'}; זום ${record.view?.zoom}; ` +
        `לשונית ${record.view?.ribbonTab}; סמן ${JSON.stringify(entry?.caret?.start ?? null)}`,
    );

    // הרשומה תמיד שמרה את כל הטאבים; מה שנשבר היה הקריאה שלה בעלייה. בלי
    // הבדיקה כאן, כשל בשאלה „שני טאבים חזרו” היה מעורפל בין השניים.
    if ((record.documents?.length ?? 0) !== 2) {
      failures.push(`הרשומה נשמרה עם ${record.documents?.length ?? 0} מסמכים במקום שניים`);
    }
    if (Object.keys(saved.files ?? {}).length !== 2) {
      failures.push(`ציפינו לשתי טיוטות — אחת לכל טאב — ונכתבו ${Object.keys(saved.files ?? {}).length}`);
    }
  }
}

// ---------------------------------------------------------------------------
// הפעלה שנייה: רק מה שנכתב דרך ה-SDK עובר לכאן.
// ---------------------------------------------------------------------------
/** רצועת הטאבים כפי שחזרה, ומה שנקרא מהטאב שלא היה פעיל. */
let restoredTabs = [];
let otherTabText = null;

if (saved && failures.length === 0) {
  try {
    const { cdp, close } = await openPage(buildPage(PAGE_TWO, saved), { label: 'session-2' });
    try {
      await waitForDocument(cdp);
      await sleep(3_000);
      restored = JSON.parse(await cdp.evaluate('window.__probe.readback()'));
      restoredTabs = JSON.parse(await cdp.evaluate('window.__probe.tabs()'));

      /**
       * הטאב שלא נטען בעלייה — טעינה עצלה, App.vue restoreTabs. הוא קיים
       * ברצועה מהרגע הראשון, והמסמך שלו נפתח כאן, ברגע שעוברים אליו. זו
       * השאלה השישית, והיא המבדילה בין טאב מצויר לבין טאב שבאמת חזר.
       */
      const otherIndex = restoredTabs.findIndex((tab) => !tab.active);
      if (otherIndex < 0) {
        failures.push('אין טאב שני ברצועה בהפעלה השנייה');
      } else {
        await cdp.evaluate(`window.__probe.selectDocumentTab(${otherIndex})`);
        // פתיחה מלאה של מנוע שני: קריאת הטיוטה, preflight, ובניית המסמך.
        for (let waited = 0; waited < OPEN_MS; waited += 500) {
          await sleep(500);
          const ready = await cdp.evaluate(
            '(function(){ var s = window.__otzariaEditor; return !!(s && s.superdoc); })()',
          );
          if (ready) break;
        }
        await sleep(2_500);
        otherTabText = JSON.parse(await cdp.evaluate('window.__probe.readback()')).text;
      }

      secondLog = JSON.parse(await cdp.evaluate('JSON.stringify(window.__probe.log)'));
    } finally {
      close();
    }
  } catch (error) {
    failures.push(`ההפעלה השנייה נכשלה: ${error.message}`);
  }
}

// ---------------------------------------------------------------------------
// ארבע השאלות.
// ---------------------------------------------------------------------------
if (restored) {
  const record = JSON.parse(saved.storage.session);
  const entry = activeDocumentEntry(record);

  if (!restored.text.includes(TYPED)) {
    failures.push(
      `הטקסט לא חזר: ציפינו ל-${JSON.stringify(TYPED)}, התקבל ${JSON.stringify(restored.text.slice(0, 60))}`,
    );
  }

  if (restored.zoom !== ZOOM) {
    failures.push(`גודל התצוגה לא חזר: ציפינו ל-${ZOOM}, התקבל ${restored.zoom}`);
  }

  if (restored.tab !== TAB) {
    failures.push(`הלשונית לא חזרה: ציפינו ל-${TAB}, התקבלה ${JSON.stringify(restored.tab)}`);
  }

  const savedCaret = entry?.caret?.start ?? null;
  if (!savedCaret) {
    failures.push('מקום הסמן לא נשמר בכלל');
  } else if (!restored.caret) {
    failures.push('הסמן לא הוחזר למסמך');
  } else if (restored.caret.offset !== savedCaret.offset) {
    failures.push(
      `הסמן חזר למקום אחר: ציפינו להיסט ${savedCaret.offset}, התקבל ${restored.caret.offset}`,
    );
  } else {
    // אינו שער אלא מדידה: היא שאומרת איזה משני העוגנים עשה את העבודה, וזו
    // התשובה היחידה שיש לשאלה „האם מזהי הפסקאות שורדים סבב ייצוא”.
    notes.push(
      restored.caret.blockId === savedCaret.blockId
        ? 'העוגן: מזהה הפסקה שרד את סבב הייצוא'
        : `העוגן: המזהה לא שרד (${savedCaret.blockId} → ${restored.caret.blockId}); הסדר הוא שהחזיר את הסמן`,
    );
  }

  notes.push(
    `הפעלה 2: טקסט ${JSON.stringify(restored.text.trim().slice(0, 30))}, זום ${restored.zoom}, ` +
      `לשונית ${restored.tab}, סמן ${JSON.stringify(restored.caret)}`,
  );

  // 5. כל הטאבים שהיו פתוחים — ולא רק הפעיל.
  if (restoredTabs.length !== 2) {
    failures.push(`חזרו ${restoredTabs.length} טאבים במקום שניים`);
  }

  // 6. והטאב שלא נטען בעלייה — יש בו את מה שנכתב בו, ולא מסמך ריק.
  if (otherTabText !== null && !otherTabText.includes(TYPED_SECOND)) {
    failures.push(
      `הטאב השני חזר בלי התוכן שלו: ציפינו ל-${JSON.stringify(TYPED_SECOND)}, ` +
        `התקבל ${JSON.stringify(otherTabText.slice(0, 60))}`,
    );
  }
  notes.push(
    `הפעלה 2: ${restoredTabs.length} טאבים; הטאב השני ` +
      JSON.stringify((otherTabText ?? '').trim().slice(0, 30)),
  );
}

for (const path of [PAGE_ONE, PAGE_TWO]) {
  try {
    rmSync(path, { force: true });
  } catch {
    /* קובץ בדיקה שנשאר ב-dist אינו סיבה להכשיל את השער */
  }
}

for (const note of notes) console.log(`  ${note}`);

const engineErrors = [...firstLog, ...secondLog].filter((line) => line.startsWith('error:'));
if (engineErrors.length > 0) {
  console.log('  שגיאות שנרשמו בדף:');
  for (const line of engineErrors.slice(0, 5)) console.log(`    ${line}`);
}

if (failures.length > 0) {
  console.error('\nשער החזרה נכשל:');
  for (const failure of failures) console.error(`  • ${failure}`);
  process.exit(1);
}

console.log(
  '\nשער החזרה עבר: הטקסט, הסמן, גודל התצוגה, הלשונית — ושני הטאבים — חזרו.',
);
