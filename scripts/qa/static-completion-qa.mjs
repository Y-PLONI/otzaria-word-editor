/**
 * שער ההשלמה מהרשימות הסטטיות (ביטויים תלמודיים, שמות מחברים), על ה-dist
 * הארוז ב-Chrome אמיתי.
 *
 * הטענה שהשער הזה קיים בשבילה היא **גבול הרשומה**. המסלול של PR ‏#46 חיבר את
 * הרשומות ב-`\n` והריץ עליהן את מנוע ההתאמה של הספר, שאינו יודע איפה רשומה
 * נגמרת — ו-`tests/unit/static-completion.test.ts` מודד את שני המסלולים באותה
 * לולאה: החדש 0 חציות, הקודם מאות. שם נמדד האלגוריתם; כאן נמדד מה שנכתב
 * בפועל למסמך.
 *
 * חמש טענות:
 *
 *   1. **עצלות** — הנכס אינו נמשך בעלייה.
 *   2. **הצעה** — ביטוי חלקי מציג ghost עם הביטוי השלם.
 *   3. **גבול** — ההצעה אינה כוללת מילה מהרשומה הבאה.
 *   4. **Tab** — מה שנכתב ל-OOXML הוא הביטוי, ולא רצף שחצה רשומות.
 *   5. **מחבר** — הרשימה השנייה היא fallback של הראשונה, והפיסוק בשם נשמר.
 *
 * הרצה:  node scripts/qa/static-completion-qa.mjs
 * היציאה 9614 שמורה לשער הזה בלבד.
 */
import { openApp, createReport } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9614);

/**
 * שתי רשומות עוקבות בקובץ הביטויים שחולקות את המילה הראשונה — בדיוק הצורה
 * שהמסלול הקודם היה מדביק זו לזו.
 */
const CONTEXT = 'אבן';
const PARTIAL = 'הש';
const PHRASE = 'אבן השתיה';
const NEXT_ENTRY_WORD = 'שאין';

/** מהרשימה השנייה, ועם פסיק — כדי שגם ה-fallback וגם הפיסוק ייבדקו. */
const AUTHOR_PARTIAL = 'גב';
const AUTHOR = 'גבאי, מאיר בן יחזקאל';

const report = createReport('שער השלמה מרשימות סטטיות', { strict: true });
const app = await openApp({ name: 'static-completion', port: PORT });

const log = (...a) => console.log(...a);

const assetState = () =>
  app
    .js(
      "JSON.stringify({ global: typeof window.__OTZARIA_STATIC_COMPLETION__," +
        " phrases: (window.__OTZARIA_STATIC_COMPLETION__ || {}).phrases ? window.__OTZARIA_STATIC_COMPLETION__.phrases.length : 0," +
        " authors: (window.__OTZARIA_STATIC_COMPLETION__ || {}).authors ? window.__OTZARIA_STATIC_COMPLETION__.authors.length : 0," +
        " tags: document.querySelectorAll('script[src*=\"static-completion\"]').length })",
    )
    .then(JSON.parse);

const ghostText = () =>
  app.js("(document.querySelector('.otzaria-book-completion-ghost') || {}).textContent || ''");

async function docText() {
  const files = (await app.docx()) ?? {};
  return (files['word/document.xml'] ?? '').replace(/<[^>]+>/g, '');
}

/** מנקה את המסמך בין הצעדים, כדי שכל צעד יתחיל מפסקה ריקה. */
async function clearDoc() {
  await app.press('a', 'KeyA', 65, process.platform === 'darwin' ? 4 : 2);
  await app.sleep(150);
  await app.press('Backspace', 'Backspace', 8, 0);
  await app.sleep(500);
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
  await step('הנכס אינו נמשך בעלייה', async () => {
    const state = await assetState();
    log('בעלייה:', JSON.stringify(state));
    if (state.global === 'undefined' && state.tags === 0) report.pass('עצלות בעלייה');
    else report.fail('עצלות בעלייה', JSON.stringify(state));
  });

  await step('הדלקת „השלמה מהספר”', async () => {
    await app.tab('✦ אוצריא');
    const clicked = await app.click('השלמה מהספר', { after: 700 });
    if (!clicked) return report.fail('הדלקת ההשלמה', 'הכפתור „השלמה מהספר” לא נמצא');
    const state = await app.state('השלמה מהספר');
    if (state?.active !== true) return report.fail('הדלקת ההשלמה', `הכפתור לא נדלק: ${JSON.stringify(state)}`);
    report.pass('הדלקת ההשלמה');
  });

  await step('ביטוי חלקי מציג את הביטוי השלם, ולא חוצה לרשומה הבאה', async () => {
    await app.caretPara(0);
    await app.type(`${CONTEXT} ${PARTIAL}`);
    await app.sleep(1_500);

    const state = await assetState();
    log('אחרי ההקלדה:', JSON.stringify(state));
    if (state.global !== 'object' || state.phrases !== 325 || state.authors !== 651) {
      return report.fail('טעינת הנכס', `הנכס לא נטען כמצופה: ${JSON.stringify(state)}`);
    }
    report.pass('טעינת הנכס', `${state.phrases} ביטויים, ${state.authors} מחברים`);

    const ghost = await ghostText();
    log('ghost:', JSON.stringify(ghost));
    if (ghost.trim() === 'השתיה') report.pass('ה-ghost מציג את ההשלמה');
    else report.fail('ה-ghost מציג את ההשלמה', `הוצג ${JSON.stringify(ghost)} במקום „השתיה”`);

    /*
     * ובלי `return` שלפניה: הטענה על החציה חייבת להיבדק גם — ובעיקר — כשמה
     * שהוצג אינו מה שצפינו. הצעה שגויה שגם חוצה רשומה היא שני ממצאים, ושער
     * שיוצא בראשון מדווח על אחד.
     */
    if (ghost.includes(NEXT_ENTRY_WORD)) {
      report.fail('ההצעה אינה חוצה רשומה', `ההצעה כוללת „${NEXT_ENTRY_WORD}” מהרשומה הבאה: ${JSON.stringify(ghost)}`);
    } else {
      report.pass('ההצעה אינה חוצה רשומה');
    }
  });

  await step('Tab כותב את הביטוי ל-OOXML', async () => {
    await app.press('Tab', 'Tab', 9, 0);
    await app.sleep(900);

    const text = await docText();
    log('המסמך:', JSON.stringify(text.slice(-60)));
    if (!text.includes(PHRASE)) {
      report.fail('הביטוי נכתב ל-OOXML', `הביטוי אינו במסמך: ${JSON.stringify(text.slice(-60))}`);
    } else {
      report.pass('הביטוי נכתב ל-OOXML');
    }

    if (text.includes(NEXT_ENTRY_WORD)) {
      report.fail('מה שנכתב אינו חוצה רשומה', `נכתב גם „${NEXT_ENTRY_WORD}”: ${JSON.stringify(text.slice(-60))}`);
    } else {
      report.pass('מה שנכתב אינו חוצה רשומה');
    }

    await app.type('א');
    await app.sleep(250);
    const typedAfter = await docText();
    if (typedAfter === `${text}א`) report.pass('ההקלדה אחרי Tab נמצאת אחרי הביטוי');
    else report.fail('ההקלדה אחרי Tab נמצאת אחרי הביטוי', `המסמך: ${JSON.stringify(typedAfter.slice(-80))}`);
  });

  await step('שם מחבר: ה-fallback לרשימה השנייה, והפיסוק', async () => {
    await clearDoc();
    await app.caretPara(0);
    await app.type(`${CONTEXT} ${AUTHOR_PARTIAL}`);
    await app.sleep(1_200);

    const ghost = await ghostText();
    log('ghost:', JSON.stringify(ghost));
    if (ghost.trim() === AUTHOR) report.pass('ההשלמה מרשימת המחברים', 'כולל הפסיק');
    else report.fail('ההשלמה מרשימת המחברים', `הוצג ${JSON.stringify(ghost)} במקום „${AUTHOR}”`);
  });

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
