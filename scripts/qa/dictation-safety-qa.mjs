/** Dictation regressions against the real document engine and a streamed voice-service stub. */
import assert from 'node:assert/strict';
import { openApp, sleep } from './harness.mjs';

const app = await openApp({ name: 'dictation-safety', port: Number(process.env.QA_PORT ?? 9751) });
const json = (expression) => app.js(`(async () => JSON.stringify(await (${expression})))()`).then(JSON.parse);
const selection = 'window.__qa.sd().ui.selection';
let checks = 0;
function check(label, condition) {
  assert.ok(condition, label);
  checks++;
  console.log(`✓ ${label}`);
}
async function until(predicate) {
  const deadline = Date.now() + 5000;
  do {
    if (await predicate()) return;
    await sleep(50);
  } while (Date.now() < deadline);
  throw new Error('Timed out waiting for dictation');
}
const idle = () => json(`document.querySelector('.dictation-fab').getAttribute('aria-busy') === 'false' &&
  document.querySelector('.dictation-fab').getAttribute('aria-pressed') === 'false'`);
async function start() {
  assert.equal(await idle(), true, 'previous recording must be finished');
  assert.equal(await app.clickSel('.dictation-fab', 0, { after: 0 }), true);
  await until(() => json(`document.querySelector('.dictation-fab').getAttribute('aria-pressed') === 'true'`));
}
const send = (event) => app.js(`window.__dictationSafety.send(${JSON.stringify(event)})`);
const text = async () => {
  const xml = (await app.docx())['word/document.xml'];
  return [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((match) => match[1]).join('');
};
async function draft(expected) {
  // Export only after the recording has settled. Exporting on every poll can
  // race the engine's selection restoration against an in-flight caret move.
  await until(async () => {
    const live = await json(`({
      text: [...document.querySelectorAll('.document-pane .superdoc-text-run')].map((run) => run.textContent).join(''),
      target: ${selection}.getSnapshot().selectionTarget,
    })`);
    return live.text === expected && live.target?.end.offset === expected.length;
  });
}

try {
  await app.caret(0);
  await app.type('ראשית');
  const styled = await json(`(async () => {
    const nodeId = ${selection}.getSnapshot().selectionTarget.start.blockId;
    return window.__qa.doc().paragraphs.setStyle({ target: { kind: 'block', nodeType: 'paragraph', nodeId }, styleId: 'Heading1' });
  })()`);
  assert.equal(styled.success, true);
  await app.js(`(() => {
    const D = window.__dictationSafety = { contexts: [], session: null, failure: null };
    D.send = (event) => {
      const session = D.session;
      if (!session || session.closed) return;
      session.events.push(event);
      session.wake?.();
      session.wake = null;
    };
    window.__qaHost.replies['network.fetchStream'] = (request) => (async function* () {
      if (D.failure) throw new Error(D.failure); // SDK stream errors contain message, without code.
      yield { type: 'response', sequence: 0, status: 200, ok: true, headers: {} };
      const line = (event) => ({ type: 'data', sequence: 1, body: JSON.stringify(event) + '\\n' });
      if (request.url.endsWith('/status')) {
        yield line({ service: 'otzaria-voice', pack: true, runtime: true });
      } else if (request.url.endsWith('/dictate')) {
        D.contexts.push(JSON.parse(request.body).context);
        const session = D.session = { events: [], wake: null, closed: false };
        try {
          yield line({ type: 'ready' });
          while (true) {
            if (!session.events.length) await new Promise((resolve) => { session.wake = resolve; });
            const event = session.events.shift();
            yield line(event);
            if (event.type === 'end') break;
          }
        } finally { session.closed = true; }
      } else {
        D.send({ type: 'end', reason: 'stopped' });
        yield line({});
      }
    })();
  })()`);

  await start();
  await send({ type: 'update', segment: 0, text: 'המשך', final: true });
  await draft('ראשית המשך');
  await send({ type: 'end', reason: 'stopped' });
  await until(idle);
  check('הכתבה אחרי כותרת מוסיפה רווח', (await text()) === 'ראשית המשך');
  check('טקסט הכותרת נשלח כהקשר', (await json('window.__dictationSafety.contexts'))[0].join(' ') === 'ראשית');

  // An API edit leaves the caret at the same offset and emits no DOM input:
  // the expectedText guard must protect it independently of input listeners.
  await start();
  await send({ type: 'update', segment: 0, text: 'אבג' });
  await draft('ראשית המשך אבג');
  const edited = await json(`(async () => {
    const target = ${selection}.getSnapshot().selectionTarget;
    const end = target.end.offset;
    const result = await window.__qa.doc().replace({
      target: { ...target, start: { ...target.start, offset: end - 3 } }, text: 'דהו'
    });
    ${selection}.apply(target);
    return result;
  })()`);
  assert.equal(edited.success, true);
  await send({ type: 'update', segment: 0, text: 'אבגד', final: true });
  await until(idle);
  check('החלפת טיוטה בודקת את הטקסט הקודם', (await text()) === 'ראשית המשך דהו');

  await start();
  await send({ type: 'update', segment: 0, text: 'זחט' });
  await draft('ראשית המשך דהו זחט');
  await app.js(`(() => {
    const target = ${selection}.getSnapshot().selectionTarget;
    ${selection}.apply({ ...target, start: { ...target.start, offset: target.end.offset - 3 } });
  })()`);
  await app.type('יכל');
  await until(idle);
  await send({ type: 'update', segment: 0, text: 'זחטי', final: true });
  check('קלט משתמש עוצר את ההקלטה ושומר את התיקון', (await text()) === 'ראשית המשך דהו יכל');
  check('המיקרופון נעצר דרך בקשת ביטול', await json('window.__dictationSafety.session.closed'));

  const message = 'לתוסף אין הרשאת גישה לשירותים מקומיים (network.localhost). ניתן להפעיל את ההרשאה בהגדרות.';
  await app.js(`window.__dictationSafety.failure = ${JSON.stringify(message)}`);
  await app.clickSel('.dictation-fab', 0, { after: 0 });
  await until(async () => (await app.status()).text === message);
  check('שגיאת הרשאה נשמרת ללא הפניה להתקנת שירות', (await app.status()).text === message);
  check('שגיאת הרשאה מחזירה את הכפתור להמתנה', await idle());

  const errors = (await app.log()).filter((entry) => /^(uncaught|rejected):/.test(entry));
  check('אין שגיאות JavaScript לא מטופלות', errors.length === 0);
  console.log(`\n${checks} בדיקות בטיחות הכתבה עברו.`);
} catch (error) {
  console.error({
    status: await app.status(),
    text: await text(),
    selection: await json(`${selection}.getSnapshot()`),
    contexts: await json('window.__dictationSafety?.contexts'),
    log: await app.log(),
  });
  throw error;
} finally {
  app.close();
}
