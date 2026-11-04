/**
 * המסגרת המשותפת לשערי ה-QA של הפקדים: פותחת את ה-dist הארוז ב-Chrome אמיתי,
 * מזריקה דמה של המאחז ו-API בדיקה לתוך הדף, ומחזירה כלים ללחוץ, להקליד,
 * ולקרוא את המסמך שיצא.
 *
 * למה על ה-dist ולא ב-jsdom: מנוע ה-DOCX דורש workers אמיתיים ו-canvas אמיתי,
 * ובלעדיו אין מסמך — כלומר אין מה לבדוק. השערים האלה הם המקום היחיד שבו
 * „הכפתור עובד” נמדד ולא מונח.
 *
 * ריצה מקבילה: כל שער חייב `port` משלו. שני שערים על אותה יציאה מדברים עם
 * אותו דפדפן ומשחיתים זה את מדידתו.
 */
import { readFileSync, writeFileSync, existsSync, readdirSync, rmSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { inflateRawSync } from 'node:zlib';
import { onInterrupt, openPage, requireChrome, sleep } from '../cdp.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
export const ROOT = join(HERE, '..', '..');
export const DIST = join(ROOT, 'dist');

export { sleep };

/** כמה להמתין למסמך הראשון. נדיב: המנוע פורס 16MB בדרך. */
export const OPEN_MS = 60_000;

const HOST_STUB = readFileSync(join(HERE, 'host-stub.js'), 'utf8');
const QA_API = readFileSync(join(HERE, 'qa-api.js'), 'utf8');

/**
 * מרכיבה עותק של `dist/index.html` עם הדמה וה-API בתוכו.
 *
 * ההזרקה היא **אחרי ה-latch** ולפני שהבאנדל מוזרק: `plugin.boot` חייב להיות
 * זמין לפני שהאפליקציה קמה, בדיוק כמו אצל אוצריא.
 */
function writeProbe(name, extra = '') {
  const index = join(DIST, 'index.html');
  if (!existsSync(index)) {
    console.error('dist/index.html אינו קיים — הריצו npm run build תחילה');
    process.exit(1);
  }
  const html = readFileSync(index, 'utf8');
  const cut = html.indexOf('</script>') + '</script>'.length;
  if (cut <= '</script>'.length) {
    console.error('לא נמצא ה-latch ב-dist/index.html — אין לאן להזריק את הבדיקה');
    process.exit(1);
  }
  /*
   * דף בדיקה שנשאר משער שמת לפני `close()` נארז לתוך התוסף, ולכן `check:dist`
   * נכשל עליו — גם כשהריצה הנוכחית נקייה. נמדד: `dist/__qa-fp-p5.html` שנשכח
   * הפיל את השער הבא, וההרצה שאחריו עברה בלי ששום דבר השתנה. מנקים את מה
   * שנשאר **לפני** שכותבים, ולא סומכים על ה-`finally` של הריצה שמתה.
   */
  for (const stale of readdirSync(DIST)) {
    if (stale.startsWith('__qa-') && stale.endsWith('.html')) rmSync(join(DIST, stale), { force: true });
  }
  const probe = join(DIST, `__qa-${name}.html`);
  const inject = `\n<script>${HOST_STUB}</script>\n<script>${QA_API}</script>\n${extra}\n`;
  writeFileSync(probe, html.slice(0, cut) + inject + html.slice(cut));
  return probe;
}

/**
 * פותחת את התוסף, ממתינה למסמך, ומחזירה את הכלים.
 *
 * `port` חובה בפועל בכל שער שרץ לצד אחרים.
 */
export async function openApp({ name = 'qa', port = Number(process.env.QA_PORT ?? 9350), extra = '', settle = 3_000 } = {}) {
  requireChrome();
  const probe = writeProbe(name, extra);
  // Ctrl+C באמצע: הדפדפן נסגר דרך `cdp.mjs`, אבל דף הבדיקה שייך לכאן.
  const forgetProbe = onInterrupt(() => rmSync(probe, { force: true }));
  const removeProbe = () => {
    forgetProbe();
    rmSync(probe, { force: true });
  };
  const { cdp, close } = await openPage(`file:///${probe.split('\\').join('/')}`, { port, label: name });

  const api = makeApi(cdp);

  try {
    let ready = false;
    for (let waited = 0; waited < OPEN_MS; waited += 500) {
      await sleep(500);
      if (await cdp.evaluate('!!window.__qa && window.__qa.ready()')) {
        ready = true;
        break;
      }
    }
    if (!ready) throw new Error('התוסף לא הגיע למצב מוכן בזמן');
    // הקטלוגים (גופנים, סגנונות, יכולות) נפתרים אחרי הפתיחה.
    await sleep(settle);
  } catch (error) {
    close();
    removeProbe();
    throw error;
  }

  return {
    cdp,
    ...api,
    close() {
      try {
        cdp.close();
      } catch {
        /* הדפדפן ממילא נהרג מיד אחרי */
      }
      close();
      removeProbe();
    },
  };
}

/** הכלים שמעל ה-CDP: לחיצה, הקלדה, קריאה מהדף. */
function makeApi(cdp) {
  const js = (expr) => cdp.evaluate(expr);
  const call = (fn, ...args) =>
    cdp.evaluate(`(async () => JSON.stringify(await window.__qa.${fn}(${args.map((a) => JSON.stringify(a)).join(',')}) ?? null))()`).then((raw) =>
      raw === undefined ? null : JSON.parse(raw),
    );

  async function mouse(type, x, y, extra = {}) {
    await cdp.send('Input.dispatchMouseEvent', {
      type,
      x,
      y,
      button: 'left',
      buttons: type === 'mousePressed' ? 1 : 0,
      clickCount: 1,
      ...extra,
    });
  }

  async function clickAt(x, y, { moveFirst = true } = {}) {
    if (moveFirst) await mouse('mouseMoved', x, y, { button: 'none', buttons: 0 });
    await mouse('mousePressed', x, y);
    await sleep(40);
    await mouse('mouseReleased', x, y);
  }

  async function press(key, code, vk, modifiers = 0, text) {
    await cdp.send('Input.dispatchKeyEvent', {
      type: text ? 'keyDown' : 'rawKeyDown',
      key,
      code,
      text,
      unmodifiedText: text,
      windowsVirtualKeyCode: vk,
      nativeVirtualKeyCode: vk,
      modifiers,
    });
    await cdp.send('Input.dispatchKeyEvent', {
      type: 'keyUp',
      key,
      code,
      windowsVirtualKeyCode: vk,
      nativeVirtualKeyCode: vk,
      modifiers,
    });
  }

  return {
    js,
    call,
    clickAt,
    press,
    sleep,

    /** הקלדה בעורך. הפער בין תווים מדמה הקלדה אנושית ולא הזרקה. */
    async type(text, gapMs = 45) {
      for (const ch of text) {
        const upper = ch.toUpperCase();
        await press(ch, /^[a-zA-Z]$/.test(ch) ? `Key${upper}` : 'Space', upper.charCodeAt(0), 0, ch);
        await sleep(gapMs);
      }
    },

    /** מעבר ללשונית ברצועה, בלחיצה אמיתית. */
    async tab(label) {
      const rect = JSON.parse(await js(`JSON.stringify(window.__qa.tabRect(${JSON.stringify(label)}))`));
      if (!rect) throw new Error(`לשונית „${label}” לא נמצאה`);
      await clickAt(rect.x, rect.y);
      await sleep(300);
      const active = await js('window.__qa.activeTab()');
      if (active !== label) throw new Error(`לחיצה על „${label}” לא החליפה לשונית (פעילה: ${active})`);
    },

    /**
     * לחיצה על פקד לפי שמו. מחזירה false כשהוא לא נמצא או אינו מוצג.
     *
     * קבוצה שהתכווצה (חלון צר — ראו ui/ribbon/overflow.ts) מחביאה את פקדיה
     * מאחורי צ'יפ אחד, ולכן הפתיחה קודמת למדידה. זה בדיוק מה שהמשתמש עושה.
     */
    async click(name, opts = {}) {
      const args = `${JSON.stringify(name)}, ${JSON.stringify(opts)}`;
      let rect = JSON.parse(await js(`JSON.stringify(window.__qa.hit(${args}))`));
      // `pending` = הייתה קבוצה מכווצת והיא נפתחה; הפקד מצויר רק אחרי ש-Vue
      // רינדר את הפאנל, ולכן המדידה השנייה.
      if (rect?.pending) {
        await sleep(150);
        rect = JSON.parse(await js(`JSON.stringify(window.__qa.rect(${args}))`));
      }
      if (!rect) return false;
      await clickAt(rect.x, rect.y);
      await sleep(opts.after ?? 400);
      return true;
    },

    state: (name, opts = {}) =>
      js(`JSON.stringify(window.__qa.state(${JSON.stringify(name)}, ${JSON.stringify(opts)}))`).then(JSON.parse),

    controls: (scope) => js(`JSON.stringify(window.__qa.controls(${JSON.stringify(scope ?? null)}))`).then(JSON.parse),
    shadowed: () => js('JSON.stringify(window.__qa.shadowed())').then(JSON.parse),
    tabs: () => js('JSON.stringify(window.__qa.tabs())').then(JSON.parse),

    status: () => js('JSON.stringify(window.__qa.status())').then(JSON.parse),
    messages: () => js('JSON.stringify(window.__qa.messages())').then(JSON.parse),
    hostCalls: () => js('JSON.stringify(window.__qa.hostCalls())').then(JSON.parse),
    log: () => js('JSON.stringify(window.__qa.log)').then(JSON.parse),
    reset: () => js('window.__qa.reset()'),
    cmd: (id) => js(`JSON.stringify(window.__qa.cmd(${JSON.stringify(id)}))`).then(JSON.parse),
    selection: () => js('JSON.stringify(window.__qa.selection())').then(JSON.parse),
    screenText: () => js('window.__qa.screenText()'),
    lineCount: () => js('window.__qa.lineCount()'),
    /*
      `Promise.resolve(...)` עוטף את שתי אלה מפני שבורר החיפוש (RibbonCombo)
      אינו יכול לענות סינכרונית: הרשימה שלו קיימת ב-DOM רק כשהוא פתוח, ו-Vue
      מרנדר במיקרו-משימה. `Runtime.evaluate` נשלח עם `awaitPromise`, ולכן
      הביטוי רשאי להחזיר Promise — ובורר `<select>` נייטיב, שכן עונה מיד,
      עובר דרך אותה עטיפה בלי שינוי בהתנהגות.
    */
    selectValue: (name, value) =>
      js(`Promise.resolve(window.__qa.selectValue(${JSON.stringify(name)}, ${JSON.stringify(value)}))`),
    options: (name) =>
      js(`Promise.resolve(window.__qa.options(${JSON.stringify(name)})).then(function (r) { return JSON.stringify(r); })`).then(JSON.parse),


    /* -------- תפריטים, פופאוברים ודיאלוגים -------- */

    menuItems: () => js('JSON.stringify(window.__qa.menuItems())').then(JSON.parse),
    menuOpen: () => js('window.__qa.menuOpen()'),

    /** לוחצת על פריט בתפריט פתוח. false כשהפריט אינו שם. */
    async clickMenu(label, { after = 700 } = {}) {
      const rect = JSON.parse(await js(`JSON.stringify(window.__qa.menuRect(${JSON.stringify(label)}))`));
      if (!rect) return false;
      await clickAt(rect.x, rect.y);
      await sleep(after);
      return true;
    },

    /** פותחת תפריט של RibbonMenuButton ומחזירה את הפריטים שבו. */
    async openMenu(buttonName) {
      const opened = await this.click(buttonName, { after: 400 });
      if (!opened) return null;
      return JSON.parse(await js('JSON.stringify(window.__qa.menuItems())'));
    },

    paletteOpen: () => js('window.__qa.paletteOpen()'),
    paletteSwatches: () => js('JSON.stringify(window.__qa.paletteSwatches())').then(JSON.parse),
    async clickPalette(index, { after = 600 } = {}) {
      const rect = JSON.parse(await js(`JSON.stringify(window.__qa.paletteRect(${index}))`));
      if (!rect) return false;
      await clickAt(rect.x, rect.y);
      await sleep(after);
      return true;
    },

    galleryItems: () => js('JSON.stringify(window.__qa.galleryItems())').then(JSON.parse),
    async clickGallery(label, { after = 800 } = {}) {
      const rect = JSON.parse(await js(`JSON.stringify(window.__qa.galleryRect(${JSON.stringify(label)}))`));
      if (!rect) return false;
      await clickAt(rect.x, rect.y);
      await sleep(after);
      return true;
    },

    async clickTableCell(row, col, { after = 900 } = {}) {
      const rect = JSON.parse(await js(`JSON.stringify(window.__qa.tableCellRect(${row}, ${col}))`));
      if (!rect) return false;
      // הגריד מסתמך על hover לפני הלחיצה — בלעדיו המידות אינן נבחרות.
      await mouse('mouseMoved', rect.x, rect.y, { button: 'none', buttons: 0 });
      await sleep(120);
      await clickAt(rect.x, rect.y, { moveFirst: false });
      await sleep(after);
      return true;
    },

    dialog: () => js('JSON.stringify(window.__qa.dialog())').then(JSON.parse),
    dialogFill: (idOrName, value) =>
      js(`window.__qa.dialogFill(${JSON.stringify(idOrName)}, ${JSON.stringify(value)})`),
    async clickDialog(name, { after = 700 } = {}) {
      const rect = JSON.parse(await js(`JSON.stringify(window.__qa.dialogRect(${JSON.stringify(name)}))`));
      if (!rect) return false;
      await clickAt(rect.x, rect.y);
      await sleep(after);
      return true;
    },

    exists: (selector) => js(`window.__qa.exists(${JSON.stringify(selector)})`),
    async clickSel(selector, index = 0, { after = 500 } = {}) {
      const rect = JSON.parse(await js(`JSON.stringify(window.__qa.rectSel(${JSON.stringify(selector)}, ${index}))`));
      if (!rect) return false;
      await clickAt(rect.x, rect.y);
      await sleep(after);
      return true;
    },

    /** Escape — לסגור תפריט או דיאלוג. */
    async escape() {
      await press('Escape', 'Escape', 27);
      await sleep(300);
    },

    /**
     * ממקמת סמן בשורת טקסט. בלי זה אין בחירה, וכל פקד מדווח „אין סמן”.
     *
     * `lineIndex` **אינו אינדקס פסקה** — ראו `Q.lineRect`. למי שמתכוון לפסקה
     * מסוימת יש `caretPara`.
     */
    async caret(lineIndex = 0) {
      const rect = JSON.parse(await js(`JSON.stringify(window.__qa.lineRect(${lineIndex}))`));
      if (!rect) throw new Error('אין שורת טקסט במסמך — אין לאן למקם סמן');
      await clickAt(rect.x, rect.y);
      await sleep(600);
      return rect;
    },

    paraCount: () => js('window.__qa.paraCount()').then(Number),
    caretBlock: (timeoutMs = null) => call('caretBlock', timeoutMs),

    /**
     * ממקמת סמן בפסקה — לפי אינדקס פסקה או לפי הטקסט שבה — **ומאמתת מול
     * המנוע** שהסמן אכן שם, כי לחיצה שנחתה על השכנה עוברת אחרת בשקט.
     */
    async caretPara(target, { attempts = 4, after = 600, verifyMs = 20_000 } = {}) {
      const where =
        typeof target === 'number'
          ? `window.__qa.paraRect(${target})`
          : `window.__qa.paraRectByText(${JSON.stringify(target)})`;
      const name = typeof target === 'number' ? `פסקה ${target}` : `הפסקה „${target}”`;
      /*
        חימום לפני הלחיצה: `blocks.list` הראשון אחרי מוטציה נמדד ב-390ms מול
        ~1ms בהמשך, ובלעדיו הקריאה הקרה היא זו שעומדת מול הפעמון.
      */
      const warm = await this.caretBlock(verifyMs);
      if (!warm.answered) throw new Error(`האימות של ${name} לא התבצע — ${warm.why}`);

      let landed = null;
      for (let attempt = 0; attempt < attempts; attempt += 1) {
        // מחדש בכל סבב: reflow מזיז את הגאומטריה בין הלחיצות.
        const rect = JSON.parse(await js(`JSON.stringify(${where})`));
        if (!rect) throw new Error(`אין פסקה ${JSON.stringify(target)} במסמך`);
        await clickAt(rect.x, rect.y);
        await sleep(after);
        landed = await this.caretBlock(verifyMs);
        // „לא ענה” אינו „נחת במקום אחר”, ולכן אינו נספר כנחיתה שגויה.
        if (!landed.answered) throw new Error(`האימות של ${name} לא התבצע — ${landed.why}`);
        if (landed.blockId === rect.nodeId) return { ...rect, block: landed };
        await sleep(400);
      }
      const at = landed.text ? `„${landed.text}”` : `בלוק ${landed.blockId}`;
      throw new Error(`הסמן לא הגיע ל${name} — נחת על ${at}`);
    },

    /** גרירת עכבר אמיתית בין שתי נקודות. */
    async drag(fromX, fromY, toX, toY) {
      await mouse('mouseMoved', fromX, fromY, { button: 'none', buttons: 0 });
      await mouse('mousePressed', fromX, fromY);
      await sleep(60);
      await mouse('mouseMoved', toX, toY, { buttons: 1 });
      await sleep(60);
      await mouse('mouseReleased', toX, toY);
      await sleep(500);
    },

    /** בוחרת את השורה כולה: לחיצה בתחילתה וגרירה לסופה. */
    async selectLine(lineIndex = 0) {
      const rect = JSON.parse(await js(`JSON.stringify(window.__qa.lineRect(${lineIndex}))`));
      if (!rect) throw new Error('אין שורת טקסט לבחירה');
      await mouse('mouseMoved', rect.x, rect.y, { button: 'none', buttons: 0 });
      await mouse('mousePressed', rect.x, rect.y);
      await sleep(60);
      await mouse('mouseMoved', rect.right, rect.y, { buttons: 1 });
      await sleep(60);
      await mouse('mouseReleased', rect.right, rect.y);
      await sleep(500);
      return rect;
    },

    /** בחירת טווח במקלדת: Shift+חץ, כשהסמן כבר במקום. */
    async extendSelection(chars = 5) {
      for (let i = 0; i < chars; i++) {
        await press('ArrowRight', 'ArrowRight', 39, 8);
        await sleep(60);
      }
      await sleep(300);
    },

    /** מייצאת docx ומחזירה את הקבצים שבתוכו כמחרוזות. ההוכחה האמיתית. */
    async docx() {
      const base64 = await cdp.evaluate('window.__qa.exportBase64()');
      if (!base64) return null;
      return unzip(Buffer.from(base64, 'base64'));
    },
  };
}

/* ------------------------------------------------------------------ */
/* קורא ZIP מינימלי — כדי לקרוא את ה-docx בלי להוסיף תלות                */
/* ------------------------------------------------------------------ */

/**
 * מחלצת את כל הערכים מ-ZIP. מטפלת ב-stored (0) וב-deflate (8) — שני הסוגים
 * שמייצא docx מייצר. מחזירה מפה של נתיב → מחרוזת UTF-8.
 */
export function unzip(buffer) {
  const files = {};
  // סוף הספרייה המרכזית: חותם 0x06054b50, מהסוף אחורה (יש הערה בסוף).
  let eocd = -1;
  for (let i = buffer.length - 22; i >= 0 && i > buffer.length - 22 - 65_536; i--) {
    if (buffer.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('הקובץ אינו ZIP תקין — לא נמצא EOCD');

  const count = buffer.readUInt16LE(eocd + 10);
  let ptr = buffer.readUInt32LE(eocd + 16);

  for (let i = 0; i < count; i++) {
    if (buffer.readUInt32LE(ptr) !== 0x02014b50) throw new Error('רשומת ספרייה פגומה');
    const method = buffer.readUInt16LE(ptr + 10);
    const compressedSize = buffer.readUInt32LE(ptr + 20);
    const nameLen = buffer.readUInt16LE(ptr + 28);
    const extraLen = buffer.readUInt16LE(ptr + 30);
    const commentLen = buffer.readUInt16LE(ptr + 32);
    const localOffset = buffer.readUInt32LE(ptr + 42);
    const name = buffer.toString('utf8', ptr + 46, ptr + 46 + nameLen);

    // הכותרת המקומית: שדות ה-extra שלה עשויים להיות באורך אחר מזה שבספרייה.
    const localNameLen = buffer.readUInt16LE(localOffset + 26);
    const localExtraLen = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLen + localExtraLen;
    const raw = buffer.subarray(dataStart, dataStart + compressedSize);

    try {
      const bytes = method === 0 ? raw : inflateRawSync(raw);
      files[name] = bytes.toString('utf8');
    } catch {
      files[name] = '';
    }

    ptr += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}

/* ------------------------------------------------------------------ */
/* דיווח                                                               */
/* ------------------------------------------------------------------ */

/**
 * אוסף תוצאות של שער אחד. כל בדיקה היא שורה בדוח.
 *
 * ## `strict`, ולמה הוא קיים
 *
 * `strict: true` — השער מפיל את הריצה כשיש שורה שבורה. `false` (ברירת המחדל)
 * — שער **סקר**: הוא מודד ומדווח, והתמונה שהוא מצייר היא הערך שלו, לא פסק
 * דין. שערי הסקר סוקרים לשוניות שלמות ובהן פקדים שידוע שאינם מלאים, ולכן
 * „שבור” אצלם הוא מידע ולא רגרסיה.
 *
 * ההערה כאן הבטיחה את `strict` הזה עוד לפני שהוא נכתב, והחתימה קיבלה `title`
 * בלבד. התוצאה הייתה שבעה שערים שקוראים `report.print()` ומשליכים את הערך
 * המוחזר — כלומר מדפיסים „✗ … שבור” ויוצאים עם קוד 0. מי שהריץ אותם ראה פלט
 * שאומר „נכשל” לצד ריצה שהצליחה, וזה גרוע משתי האפשרויות: שער שנכשל ביושר
 * מתקן, ושער שעובר ביושר מרגיע.
 *
 * לכן שתי ההתנהגויות מפורשות מעתה, ואף אחת מהן אינה שתיקה: `strict` מציב
 * `process.exitCode`, ובלעדיו נדפסת שורה שאומרת במפורש שקוד היציאה אינו נגזר
 * מהשורות שלמעלה.
 *
 * `process.exitCode` ולא `process.exit()`, משתי סיבות שאינן „לא לקטוע את
 * `finally`” — כל השערים קוראים `print()` **אחרי** ה-`finally`, ולכן
 * `app.close()` ממילא כבר רץ. הסיבות האמיתיות: השורה שאחרי `print()` בשער
 * צריכה להספיק להידפס, ושער שישכח לקרוא `process.exit()` עדיין ייכשל.
 *
 * ואכן, בשערים שכבר כותבים `process.exit(report.print() > 0 ? 1 : 0)` הערך
 * שנקבע כאן נדרס מיד אחר כך — באותו ערך בדיוק. זה מכוון: המנגנון כאן הוא
 * רשת ביטחון, לא המסלול הראשי, ומה שהשתנה בפועל אצל שערי הסקר הוא ההצהרה
 * המפורשת שמתחת.
 */
export function createReport(title, { strict = false } = {}) {
  const rows = [];
  return {
    rows,
    strict,
    pass(name, detail = '') {
      rows.push({ name, verdict: 'עובד', detail });
    },
    fail(name, detail = '') {
      rows.push({ name, verdict: 'שבור', detail });
    },
    partial(name, detail = '') {
      rows.push({ name, verdict: 'חלקי', detail });
    },
    skip(name, detail = '') {
      rows.push({ name, verdict: 'לא נבדק', detail });
    },
    /**
     * הצעד לא הסתיים — הדף קפא, לא הפקד נכשל.
     *
     * ורדיקט נפרד מ-`fail`, כי אלה שני דברים שונים שנספרו יחד: בשער אחד נמדדו
     * „20 שבורים”, ומתוכם **16 היו קפיאות headless** וארבע כשלים אמיתיים.
     * מספר אחד שמערבב „הפקד שבור” עם „השער לא הצליח למדוד” אינו אומר דבר על
     * אף אחד מהשניים, ומי שקורא אותו לומד את הדבר הלא נכון בכל כיוון.
     *
     * ואינו נספר כשבור גם ב-strict, מאותה סיבה: קפיאה היא כשל של סביבת
     * המדידה, וגרירת השער לאדום בגללה הייתה מלמדת להתעלם ממנו.
     */
    stuck(name, detail = '') {
      rows.push({ name, verdict: 'תקוע', detail });
    },
    print() {
      console.log(`\n=== ${title} ===`);
      for (const row of rows) {
        const mark = { עובד: '✓', שבור: '✗', חלקי: '~', תקוע: '⧗', 'לא נבדק': '·' }[row.verdict];
        console.log(`${mark} ${row.name}${row.detail ? ' — ' + row.detail : ''}`);
      }
      const count = (verdict) => rows.filter((r) => r.verdict === verdict).length;
      const broken = count('שבור');
      const partial = count('חלקי');
      const stuck = count('תקוע');
      const skipped = count('לא נבדק');
      // „עובדים” נספר במפורש ולא כשארית: החישוב הקודם היה
      // `rows.length - broken - partial`, ולכן שורות „לא נבדק” נספרו כעובדות.
      console.log(
        `\nסה"כ ${rows.length}: ${count('עובד')} עובדים, ${partial} חלקיים, ${broken} שבורים` +
          (stuck ? `, ${stuck} תקועים (קפיאת מדידה, לא כשל פקד)` : '') +
          (skipped ? `, ${skipped} לא נבדקו` : ''),
      );
      if (strict) {
        if (broken > 0) process.exitCode = 1;
      } else {
        console.log(
          broken > 0
            ? `שער סקר: ${broken} שורות שבורות **אינן** מפילות את הריצה — קוד היציאה כאן אינו עדות למצב שלמעלה.`
            : 'שער סקר: קוד היציאה אינו נגזר מהשורות שלמעלה.',
        );
      }
      console.log('JSON:' + JSON.stringify({ title, strict, rows }));
      return broken;
    },
  };
}
