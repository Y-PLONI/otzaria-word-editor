/**
 * כפתור ההקלטה הצף (src/ui/shell/DictationFab.vue) מול שירות דיבור מדומה.
 *
 * השירות האמיתי (127.0.0.1:47811) אינו זמין כאן, ולכן `network.fetchStream`
 * נענה מהדמה: `/status` מוכן, ו-`/dictate` שולח מילה אחת מיד ומילה שנייה
 * אחרי `SECOND_WORD_MS`. כך אפשר לראות לאן נכתבת המילה השנייה כשהמשתמש עבר
 * למסמך אחר באמצע.
 *
 * מה נמדד: צבע ומיקום; הקלטה וטקסט; מעבר טאב באמצע הקלטה (המילה השנייה אינה
 * נכתבת לשום מסמך); ביטול בזמן ההתחלה; מסמך לקריאה בלבד (הכפתור הצף והכפתור
 * בלשונית „אוצריא”); מצב מיקוד.
 * הרצה: npm run build && node scripts/qa/dictation-fab-qa.mjs.
 */
import { openApp, sleep } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9747);
const SECOND_WORD_MS = 2500;
let failed = false;

function check(label, ok, detail = '') {
  console.log(`${ok ? '✓' : '✗'} ${label}${detail ? ` — ${detail}` : ''}`);
  if (!ok) failed = true;
}

/**
 * הדמה בתוך הדף. `__dict.gate` עוצר את `/status` כשהוא מוגדר, כדי לבדוק ביטול
 * בזמן ההתחלה; `__dict.paths` רושם כל בקשה.
 */
const INSTALL_SERVICE = `(() => {
  const D = (window.__dict = { paths: [], gate: null, release: null });
  const line = (o) => ({ type: 'data', body: JSON.stringify(o) + '\\n' });
  window.__qaHost.replies['network.fetchStream'] = (payload) => {
    const path = String(payload.url).replace('http://127.0.0.1:47811', '');
    D.paths.push(path);
    return (async function* () {
      yield { sequence: 0, type: 'response', status: 200, ok: true, headers: {} };
      if (path === '/status') {
        if (D.gate) await D.gate;
        yield line({ service: 'otzaria-voice', pack: true, runtime: true, microphone: true });
        return;
      }
      if (path === '/dictate') {
        yield line({ type: 'ready' });
        yield line({ type: 'update', segment: 0, text: 'מילהראשונה', final: true });
        await new Promise((r) => setTimeout(r, ${SECOND_WORD_MS}));
        yield line({ type: 'update', segment: 1, text: 'מילהשנייה', final: true });
        await new Promise((r) => { D.release = r; });
        yield line({ type: 'end', reason: 'stopped' });
        return;
      }
      if (D.release) D.release();
      yield line({});
    })();
  };
  return true;
})()`;

async function main() {
  const page = await openApp({ name: 'dictation-fab', port: PORT });
  try {
    const js = (expr) => page.js(expr);
    const json = async (expr) => JSON.parse(await js(`JSON.stringify(${expr})`));
    const fab = () =>
      json(`(() => {
        const b = document.querySelector('.dictation-fab');
        if (!b) return null;
        const r = b.getBoundingClientRect();
        const cs = getComputedStyle(b);
        return { cls: b.className, disabled: b.disabled, desc: b.dataset.tipDesc, bg: cs.backgroundColor,
          left: r.left, right: r.right, top: r.top, bottom: r.bottom };
      })()`);
    const color = (token) =>
      json(`(() => { const e = document.createElement('div'); e.style.color = 'var(${token})';
        document.body.appendChild(e); const c = getComputedStyle(e).color; e.remove(); return c; })()`);
    const paneTexts = () =>
      json(`[...document.querySelectorAll('.document-pane')].map((p) => (p.textContent || '').replace(/\\s+/g, ' '))`);
    const caretInDoc = async () => {
      const r = await json(`(() => { const l = [...document.querySelectorAll('.document-pane')]
        .find((p) => getComputedStyle(p).display !== 'none')?.querySelector('.superdoc-line');
        const b = l && l.getBoundingClientRect(); return b && { x: b.right - 4, y: b.top + b.height / 2 }; })()`);
      if (!r) throw new Error('אין שורה במסמך הפעיל');
      await page.clickAt(r.x, r.y);
      await sleep(400);
    };
    const clickFab = async () => {
      if (!await page.clickSel('.dictation-fab', 0, { after: 0 })) throw new Error('הכפתור הצף אינו נגיש');
    };
    const waitFor = async (pred, ms = 5000) => {
      for (let t = 0; t < ms; t += 100) {
        if (await pred()) return true;
        await sleep(100);
      }
      return false;
    };

    await js(INSTALL_SERVICE);

    // --- מראה ומיקום
    const primary = await color('--color-primary');
    const error = await color('--color-error');
    let f = await fab();
    check('הכפתור הצף קיים ופעיל', !!f && !f.disabled);
    check('בהמתנה: צבע הנושא', f?.bg === primary, `${f?.bg} מול ${primary}`);
    const statusTop = await json(`document.querySelector('.word-statusbar').getBoundingClientRect().top`);
    check('32 פיקסלים מעל שורת המצב', Math.round(statusTop - f.bottom) === 32, `${Math.round(statusTop - f.bottom)}px`);

    // --- הקלטה. קודם נפתח מסמך שני וחוזרים לראשון: המעבר שנבדק למטה הוא אל
    // טאב **קיים**. לשני טאבים שנפתחו פעם אחת יש אותו `documentGeneration`, ולכן
    // רק מעבר כזה חושף הקלטה שממשיכה לכתוב למסמך שהוסתר; טאב חדש מקבל מונה אחר.
    if (!await page.clickSel('.word-doctabs-new')) throw new Error('„+” אינו נגיש');
    await sleep(1500);
    if (!await page.clickSel('.word-doctab', 0)) throw new Error('הטאב הראשון אינו נגיש');
    await sleep(1500);
    await caretInDoc();
    await clickFab();
    check('בזמן הקלטה: אדום', await waitFor(async () => (await fab()).bg === error), `${(await fab()).bg} מול ${error}`);
    check('המילה הראשונה נכתבה', await waitFor(async () => (await paneTexts()).some((t) => t.includes('מילהראשונה'))));

    // --- מעבר לטאב הקיים באמצע הקלטה
    if (!await page.clickSel('.word-doctab', 1)) throw new Error('הטאב השני אינו נגיש');
    check('מעבר למסמך אחר עוצר את ההקלטה', await waitFor(async () => !(await fab()).cls.includes('--on')));
    await sleep(SECOND_WORD_MS + 1000);
    const texts = await paneTexts();
    check('המילה השנייה לא נכתבה לשום מסמך', texts.every((t) => !t.includes('מילהשנייה')), JSON.stringify(texts.map((t) => t.slice(0, 40))));

    // --- ביטול בזמן ההתחלה
    await js(`(() => { window.__dict.paths.length = 0; window.__dict.gate = new Promise((r) => { window.__dict.openGate = r; }); })()`);
    await caretInDoc();
    await clickFab();
    check('בזמן ההתחלה: לחיץ (לביטול)', await waitFor(async () => {
      const s = await fab();
      return s.cls.includes('--starting') && !s.disabled;
    }));
    await clickFab();
    await js(`window.__dict.openGate()`);
    await sleep(800);
    const paths = await json(`window.__dict.paths`);
    check('ביטול בזמן ההתחלה: המיקרופון לא נפתח', !paths.includes('/dictate'), JSON.stringify(paths));
    check('ביטול בזמן ההתחלה: חזר להמתנה', await waitFor(async () => {
      const s = await fab();
      return !s.cls.includes('--starting') && !s.cls.includes('--on');
    }));
    await js(`window.__dict.gate = null`);

    // --- קריאה בלבד
    const lock = async () => {
      await js(`[...document.querySelectorAll('[role=tab], button')].find((e) => e.textContent.trim() === 'סקירה')?.click()`);
      await sleep(500);
      const clicked = await js(`(() => { const b = [...document.querySelectorAll('[data-tip-title]')].find((e) => e.dataset.tipTitle === 'הגבל עריכה'); b?.click(); return !!b; })()`);
      if (!clicked) throw new Error('„הגבל עריכה” לא נמצא');
      await sleep(300);
    };
    await lock();
    await lock(); // לחיצה שנייה מאשרת
    check('קריאה בלבד: הכפתור הצף מנוטרל', await waitFor(async () => (await fab()).disabled));
    check('קריאה בלבד: הטולטיפ מסביר', (await fab()).desc === 'המסמך פתוח לקריאה בלבד', (await fab()).desc);
    await js(`[...document.querySelectorAll('[role=tab], button')].find((e) => e.textContent.trim().includes('אוצריא'))?.click()`);
    await sleep(500);
    const ribbonMic = await json(`(() => { const b = [...document.querySelectorAll('button')].find((e) => e.textContent.trim() === 'הקלטה'); return b ? { disabled: b.disabled || b.getAttribute('aria-disabled') === 'true' } : null; })()`);
    check('קריאה בלבד: גם הכפתור בלשונית „אוצריא” מנוטרל', ribbonMic?.disabled === true, JSON.stringify(ribbonMic));
    await lock(); // ביטול הנעילה
    check('שחרור הנעילה מחזיר את הכפתור', await waitFor(async () => !(await fab()).disabled));

    // --- מצב מיקוד
    await js(`document.querySelector('.view-mode-buttons .sb-icon-btn')?.click()`);
    await sleep(800);
    const overlap = await json(`(() => { const a = document.querySelector('.dictation-fab').getBoundingClientRect();
      const b = document.querySelector('.focus-exit')?.getBoundingClientRect(); if (!b) return null;
      return !(a.right < b.left || a.left > b.right || a.bottom < b.top || a.top > b.bottom); })()`);
    check('מצב מיקוד: אינו מכסה את „יציאה ממצב מיקוד”', overlap === false, String(overlap));
    const shellTop = await json(`document.querySelector('.word-app-shell').getBoundingClientRect().top`);
    check('מצב מיקוד: המעטפת לא זזה', shellTop === 0, `${shellTop}`);
  } finally {
    page.close();
  }

  if (failed) {
    console.error('\nהשער נכשל — ראו ✗ למעלה.');
    process.exit(1);
  }
  console.log('\nהשער עבר: הכפתור הצף, ההקלטה, מעבר טאב, ביטול, קריאה בלבד ומצב מיקוד.');
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
