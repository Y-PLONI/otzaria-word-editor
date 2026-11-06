/**
 * שער השלמת ראשי-התיבות, על ה-dist הארוז ב-Chrome אמיתי.
 *
 * זהו השער שחסר: הבדיקה היחידה שהייתה מסרה לטוען `JSON.stringify(...)` — מחרוזת
 * שהבנייה אינה מייצרת — ולכן נשארה ירוקה בעוד הנכס האמיתי מציב **אובייקט**,
 * הטוען מחפש מחרוזת, וההשלמה לא עבדה מעולם. כאן רץ הנכס שנארז, בדף שנארז.
 *
 * ארבע טענות, וכל אחת היא דבר שנשבר בשקט:
 *
 *   1. **עצלות** — הנכס אינו נמשך בעלייה.
 *   2. **השער** — מילה עברית רגילה **אינה** מזריקה אותו. בלי הבדיקה הזאת די
 *      בכך שההשלמה מהספר לא מצאה דבר כדי למשוך 0.86MB אחרי כל מילה.
 *   3. **הפירוש** — ר"ת שהוקלד מציג ghost, ו-Tab כותב אותו למסמך.
 *   4. **ניקיון** — תגית ה-`<script>` אינה נשארת ב-DOM.
 *
 * הרצה:  node scripts/qa/acronym-completion-qa.mjs
 * היציאה 9611 שמורה לשער הזה בלבד.
 */
import { openApp, createReport } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9611);

/** מה שהנכס אמור להכיל בדיוק — ר' scripts/acronyms-asset.ts. */
const EXPECTED_KEYS = 17_717;

/** ר"ת שאומת מול src/data/acronyms.json: הפירוש שלו הוא זה. */
const ACRONYM = 'חז"ל';
const EXPANSION = 'חכמינו זכרונם לברכה';
/** מילים עבריות רגילות. אף אחת אינה בצורת ר"ת, ולכן אינן אמורות לגעת בנכס. */
const PLAIN_WORDS = 'ואמרו רבותינו ';

const report = createReport('שער השלמת ראשי-תיבות', { strict: true });
const app = await openApp({ name: 'acronyms', port: PORT });

const log = (...a) => console.log(...a);

/** מצב הנכס: הגלובל, ומספר תגיות ה-script שנשארו ב-DOM. */
const assetState = () =>
  app
    .js(
      "JSON.stringify({ global: typeof window.__OTZARIA_ACRONYMS__," +
        " keys: window.__OTZARIA_ACRONYMS__ ? Object.keys(window.__OTZARIA_ACRONYMS__).length : 0," +
        " tags: document.querySelectorAll('script[src*=\"acronyms\"]').length })",
    )
    .then(JSON.parse);

/** ה-ghost של ההשלמה, אם הוא מוצג. */
const ghostText = () =>
  app.js(
    "(document.querySelector('.otzaria-book-completion-ghost') || {}).textContent || ''",
  );

/** הטקסט שבמסמך, בלי תגיות. */
async function docText() {
  const files = (await app.docx()) ?? {};
  return (files['word/document.xml'] ?? '').replace(/<[^>]+>/g, '');
}

async function step(name, fn) {
  log(`\n──────── ${name} ────────`);
  try {
    await fn();
  } catch (error) {
    log('!! זרק:', error?.message);
    report.fail(name, `הצעד זרק: ${error?.message}`);
  }
}

try {
  /* 1. עצלות -------------------------------------------------------- */
  await step('הנכס אינו נמשך בעלייה', async () => {
    const state = await assetState();
    log('בעלייה:', JSON.stringify(state));
    if (state.global === 'undefined' && state.tags === 0) report.pass('עצלות בעלייה');
    else report.fail('עצלות בעלייה', JSON.stringify(state));
  });

  /* ההדלקה ---------------------------------------------------------- */
  await step('הדלקת „השלמה מהספר”', async () => {
    await app.tab('✦ אוצריא');
    const clicked = await app.click('השלמה מהספר', { after: 700 });
    if (!clicked) return report.fail('הדלקת ההשלמה', 'הכפתור „השלמה מהספר” לא נמצא');
    const state = await app.state('השלמה מהספר');
    if (state?.active !== true) return report.fail('הדלקת ההשלמה', `הכפתור לא נדלק: ${JSON.stringify(state)}`);
    report.pass('הדלקת ההשלמה');
  });

  /* 2. השער --------------------------------------------------------- */
  await step('מילה רגילה אינה מזריקה את הנכס', async () => {
    await app.caretPara(0);
    await app.type(PLAIN_WORDS);
    // מעל ה-debounce של 150ms, ועוד זמן להזרקה אילו הייתה קורית.
    await app.sleep(1_200);

    const state = await assetState();
    log('אחרי מילים רגילות:', JSON.stringify(state));
    if (state.global === 'undefined') report.pass('השער לפני הטעינה');
    else report.fail('השער לפני הטעינה', `הנכס נטען אחרי „${PLAIN_WORDS.trim()}”: ${JSON.stringify(state)}`);
  });

  /* 3. הפירוש ------------------------------------------------------- */
  await step('ר"ת מציג את הפירוש', async () => {
    await app.type(`${ACRONYM} `);
    await app.sleep(1_500);

    const typed = await docText();
    if (!typed.includes(ACRONYM)) {
      return report.fail('ההקלדה', `הר"ת לא נכנס למסמך כפי שהוקלד: ${JSON.stringify(typed.slice(-40))}`);
    }

    const state = await assetState();
    log('אחרי הר"ת:', JSON.stringify(state));
    // ספירה מדויקת ולא „לפחות”: סינון שיישבר ויוריד אלפי מפתחות היה מדפיס
    // מספר אחר ונשאר ירוק. אותו מספר נאכף ב-tests/unit/acronyms.test.ts.
    if (state.global !== 'object' || state.keys !== EXPECTED_KEYS) {
      return report.fail('טעינת הנכס', `נטענו ${state.keys} ערכים במקום ${EXPECTED_KEYS}: ${JSON.stringify(state)}`);
    }
    report.pass('טעינת הנכס', `${state.keys} ערכים`);

    const ghost = await ghostText();
    log('ghost:', JSON.stringify(ghost));
    if (ghost.trim() === EXPANSION) report.pass('ה-ghost מציג את הפירוש');
    else report.fail('ה-ghost מציג את הפירוש', `הוצג ${JSON.stringify(ghost)} במקום „${EXPANSION}”`);
  });

  await step('Tab כותב את הפירוש למסמך', async () => {
    await app.press('Tab', 'Tab', 9, 0);
    await app.sleep(900);

    const text = await docText();
    log('המסמך:', JSON.stringify(text.slice(-80)));
    if (text.includes(EXPANSION)) report.pass('הפירוש נכתב ל-OOXML');
    else report.fail('הפירוש נכתב ל-OOXML', `הפירוש אינו במסמך: ${JSON.stringify(text.slice(-80))}`);

    const ghost = await ghostText();
    if (ghost === '') report.pass('ה-ghost נעלם אחרי הקבלה');
    else report.fail('ה-ghost נעלם אחרי הקבלה', `נשאר ${JSON.stringify(ghost)}`);

    await app.type('א');
    await app.sleep(250);
    const typedAfter = await docText();
    if (typedAfter === `${text}א`) report.pass('ההקלדה אחרי Tab נמצאת אחרי הפירוש');
    else report.fail('ההקלדה אחרי Tab נמצאת אחרי הפירוש', `המסמך: ${JSON.stringify(typedAfter.slice(-100))}`);
  });

  /* 4. ניקיון ------------------------------------------------------- */
  await step('תגית ה-script אינה נשארת', async () => {
    const state = await assetState();
    if (state.tags === 0) report.pass('אין תגיות שנשארו ב-DOM');
    else report.fail('אין תגיות שנשארו ב-DOM', `${state.tags} תגיות`);
  });

  await step('ללא רעש', async () => {
    const lines = await app.log();
    const noisy = (lines ?? []).filter(
      (l) =>
        /error|שגיאה|failed|נכשל/i.test(l) &&
        !/DevTools|Download the Vue/i.test(l) &&
        // הדמה אינו מממש את שתי המתודות האלה.
        !/reader\.addContextMenuItem|fonts\.listInstalled/.test(l) &&
        // פער ותיק ומתועד של ה-projection (ראו engine/text-search.ts).
        !/projection-incomplete|typing-mutation: engine pass failed/.test(l),
    );
    if (noisy.length) report.fail('ללא רעש', noisy.join(' | '));
    else report.pass('ללא רעש');
  });
} finally {
  app.close();
}

report.print();
process.exit(process.exitCode ?? 0);
