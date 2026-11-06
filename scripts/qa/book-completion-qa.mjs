/**
 * שער „השלמה מהספר” — המסלול הוותיק, שעד עכשיו לא היה מכוסה בכלל.
 *
 * הפער הזה אינו תיאורטי: הוא איפשר רגרסיה שנמדדה בפועל — אחרי ה-Tab הראשון
 * ה-ghost נעלם, אבל ההצעה נשארה **חיה**, ולכן ה-Tab הבא נתפס
 * (`preventDefault`) והכניס חמש מילים מהספר בלי ששום דבר הוצג. שני השערים
 * האחרים לא יכלו לתפוס זאת: המאחז אינו מממש את הקורא, ולכן אצלם `cache`
 * נשאר null וההשלמה מהספר כלל אינה נכנסת למסלול.
 *
 * לכן השער הזה מדמה את הקורא של אוצריא (`reader.getCurrentState`,
 * `reader.getSectionTextMap`), ומודד את מה שנשבר: **מה שמוצג שווה למה
 * שנכתב**, גם בסבב ההמשך.
 *
 * הרצה:  node scripts/qa/book-completion-qa.mjs
 * היציאה 9615 שמורה לשער הזה בלבד.
 */
import { openApp, createReport } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9615);

/** ה"ספר" שהמאחז מחזיר. שורה אחת, בלי ניקוד — הנרמול נבדק במקום אחר. */
const LINE = 'בראשית ברא אלהים את השמים ואת הארץ והארץ היתה תהו ובהו וחשך על פני תהום';
/** מה שמוקלד, ומה שאמור להיות מוצע: חמש המילים שאחרי „בראשית”. */
const TYPED = 'בראשית ב';
const FIRST_GHOST = 'ברא אלהים את השמים ואת';

const report = createReport('שער השלמה מהספר', { strict: true });
const app = await openApp({ name: 'book-completion', port: PORT });

const log = (...a) => console.log(...a);

const ghostText = () =>
  app.js("(document.querySelector('.otzaria-book-completion-ghost') || {}).textContent || ''");

async function docText() {
  const files = (await app.docx()) ?? {};
  return (files['word/document.xml'] ?? '').replace(/<[^>]+>/g, '').replace(/\r?\n/g, '');
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
  await step('הקמת הקורא המדומה והדלקת ההשלמה', async () => {
    await app.js(`
      window.__qaHost.replies['reader.getCurrentState'] = function () {
        return Promise.resolve({ success: true, data: { currentBook: 'בראשית', currentBookId: 'b1', currentRef: 'בראשית א', currentIndex: 0 }, error: null });
      };
      window.__qaHost.replies['reader.getSectionTextMap'] = function (payload) {
        var text = payload && payload.sectionIndex === 0 ? ${JSON.stringify(LINE)} : '';
        return Promise.resolve({ success: true, data: { renderedText: text, sourceText: text, hasMore: false, nextCursor: null }, error: null });
      };
    `);

    await app.tab('✦ אוצריא');
    const clicked = await app.click('השלמה מהספר', { after: 900 });
    if (!clicked) return report.fail('הדלקת ההשלמה', 'הכפתור „השלמה מהספר” לא נמצא');
    const state = await app.state('השלמה מהספר');
    if (state?.active !== true) return report.fail('הדלקת ההשלמה', `הכפתור לא נדלק: ${JSON.stringify(state)}`);
    report.pass('הדלקת ההשלמה');
  });

  /** מה שההצעה הראשונה הציגה — נשמר כדי להשוות אל מה שנכתב. */
  let firstGhost = '';

  await step('הקלדה מהספר מציגה את המשך הפסוק', async () => {
    await app.caretPara(0);
    await app.type(TYPED);
    await app.sleep(1_600);

    firstGhost = (await ghostText()).trim();
    log('ghost:', JSON.stringify(firstGhost));
    if (firstGhost === FIRST_GHOST) report.pass('ה-ghost מציג את המשך הספר');
    else report.fail('ה-ghost מציג את המשך הספר', `הוצג ${JSON.stringify(firstGhost)} במקום „${FIRST_GHOST}”`);
  });

  await step('Tab כותב בדיוק את מה שהוצג', async () => {
    await app.press('Tab', 'Tab', 9, 0);
    await app.sleep(1_000);

    const text = await docText();
    log('המסמך:', JSON.stringify(text));
    if (firstGhost !== '' && text.endsWith(firstGhost)) report.pass('מה שנכתב שווה למה שהוצג');
    else report.fail('מה שנכתב שווה למה שהוצג', `במסמך ${JSON.stringify(text)}, הוצג ${JSON.stringify(firstGhost)}`);
  });

  /*
   * הצעד שתופס את הרגרסיה. שתי טענות שחייבות להתקיים יחד, ודווקא הצירוף הוא
   * העיקר: הצעה חיה בלי ghost פירושה Tab שכותב למסמך בלי ששום דבר הוצג.
   */
  let continuation = '';

  await step('סבב ההמשך מוצג, ולא רק חי', async () => {
    continuation = (await ghostText()).trim();
    log('ghost ההמשך:', JSON.stringify(continuation));
    if (continuation !== '') report.pass('ה-ghost של ההמשך מוצג');
    else report.fail('ה-ghost של ההמשך מוצג', 'אין ghost אחרי ה-Tab הראשון');
  });

  await step('ה-Tab הבא כותב את מה שהוצג, ורק אם הוצג', async () => {
    const before = await docText();
    await app.press('Tab', 'Tab', 9, 0);
    await app.sleep(1_000);
    const after = await docText();
    log('לפני:', JSON.stringify(before), '| אחרי:', JSON.stringify(after));

    /*
     * הטענה הזאת היא שהייתה אדומה ברגרסיה: המסמך גדל אחרי Tab בעוד שלא הוצג
     * שום ghost — כלומר Tab נתפס וכתב חמש מילים בלי ששום דבר נראה.
     */
    if (after.length > before.length && continuation === '') {
      report.fail('Tab כותב רק כשהוצג', 'המסמך גדל אחרי Tab בלי ששום ghost הוצג');
    } else {
      report.pass('Tab כותב רק כשהוצג');
    }

    /*
     * ולא „מה שנוסף שווה ל-ghost”: ה-ghost מציג את הטקסט מהמילה שתאמה, כולל
     * מילת ההקשר שכבר הוקלדה, וההכנסה מחליפה אותה כדי שלא יישאר משפט
     * חצי-מנוקד (ר' book-completion.ts). לכן הטענה היא על **סוף המסמך**.
     */
    if (continuation !== '' && !after.endsWith(continuation)) {
      report.fail('סוף המסמך שווה למה שהוצג', `במסמך ${JSON.stringify(after)}, הוצג ${JSON.stringify(continuation)}`);
    } else {
      report.pass('סוף המסמך שווה למה שהוצג');
    }
  });

  /*
   * המילים שעלו זו על זו: ה-ghost נשאר בעמדת הסמן של ההערכה הקודמת בזמן
   * שההקלדה כבר מציירת אותיות חדשות בדיוק שם. לפני התיקון 11 מתוך 18 דגימות
   * באמצע הקלדה היו כאלה. הטענה: בכל רגע, ghost שמוצג יושב בסמן.
   */
  await step('ghost אינו נשאר מעל טקסט שמוקלד', async () => {
    const sample = () =>
      app.js(`(function(){
        var g = document.querySelector('.otzaria-book-completion-ghost');
        var sd = window.__otzariaEditor && window.__otzariaEditor.superdoc;
        var c = sd && sd.ui && sd.ui.selection && sd.ui.selection.getAnchorRect({placement:'end'});
        if (!g || !c) return null;
        var r = g.getBoundingClientRect();
        return JSON.stringify({ dx: Math.round(r.right - c.left), dy: Math.round(r.top - c.top) });
      })()`).then((raw) => (raw ? JSON.parse(raw) : null));

    const stale = [];
    for (const ch of 'הארץ') {
      await app.type(ch, 0);
      for (let i = 0; i < 3; i += 1) {
        await app.sleep(40);
        const s = await sample();
        if (s && (s.dx > 2 || Math.abs(s.dy) > 2)) stale.push(`${ch}:${s.dx},${s.dy}`);
      }
    }
    if (stale.length === 0) report.pass('ghost אינו נשאר מעל טקסט שמוקלד');
    else report.fail('ghost אינו נשאר מעל טקסט שמוקלד', `ghost מחוץ לסמן: ${stale.join(' ')}`);
  });

  /*
   * מכאן הקורא המדומה מחזיק כמה ספרים, ואפשר לשנות את הטקסט „המוצג” שלהם
   * באמצע. `Shift` לבדו מעורר הערכה (keyup) בלי לפסול הצעה ובלי לכתוב, ושתי
   * הערכות במרחק של יותר מ-STATE_TTL_MS הן מה שמריץ רענון של מצב הקורא
   * ואחריו הערכה שרואה את התוצאה שלו.
   */
  const nudge = async (wait = 600) => {
    await app.press('Shift', 'ShiftLeft', 16, 0);
    await app.sleep(wait);
  };
  const refreshReader = async () => {
    await nudge(1_800);
    await nudge(1_800);
  };
  const setReader = (state, books) =>
    app.js(`
      window.__qaReader = { state: ${JSON.stringify(state)}, books: ${JSON.stringify(books)} };
      window.__qaHost.replies['reader.getCurrentState'] = function () {
        return Promise.resolve({ success: true, data: window.__qaReader.state, error: null });
      };
      window.__qaHost.replies['reader.getSectionTextMap'] = function (payload) {
        var book = window.__qaReader.books[(payload && (payload.bookUid || payload.bookId)) || ''];
        if (book === undefined) {
          return Promise.resolve({ success: false, data: null, error: { code: 'error.not_found', message: 'text book was not found' } });
        }
        var text = payload.sectionIndex === 0 ? book : '';
        return Promise.resolve({ success: true, data: { renderedText: text, sourceText: text, hasMore: false, nextCursor: null }, error: null });
      };
    `);

  /** אותו פסוק, כפי שהקורא מציג אותו כשהניקוד גלוי. */
  const NIKUD_LINE = 'בְּרֵאשִׁית בָּרָא אֱלֹהִים אֵת הַשָּׁמַיִם וְאֵת הָאָרֶץ וְהָאָרֶץ הָיְתָה תֹהוּ וָבֹהוּ';
  const NIKUD_GHOST = 'בָּרָא אֱלֹהִים אֵת הַשָּׁמַיִם וְאֵת';
  const GENESIS = { currentBook: 'בראשית', currentBookId: 'b1', bookUid: null, currentRef: 'בראשית א', currentIndex: 0 };

  /*
   * הדיווח: „מעתיק עם ניקוד ולא לפי איך שזה מוצג”. המטמון נשמר לפי ספר
   * ושורה, והסתרת ניקוד אינה משנה אף אחד מהם — ולכן ההשלמה המשיכה להציע את
   * מה שנטען קודם, עד גלילה. כאן שני הכיוונים, בלי לזוז מהשורה.
   */
  await step('שינוי ניקוד בקורא מגיע להשלמה בלי גלילה', async () => {
    await setReader(GENESIS, { b1: NIKUD_LINE });
    await refreshReader();
    await app.press('Enter', 'Enter', 13, 0);
    await app.type('בראשית ב');
    await app.sleep(1_600);
    const shown = (await ghostText()).trim();
    log('ניקוד גלוי:', JSON.stringify(shown));
    if (shown === NIKUD_GHOST) report.pass('הצגת ניקוד בקורא — ההשלמה מנוקדת');
    else report.fail('הצגת ניקוד בקורא — ההשלמה מנוקדת', `הוצג ${JSON.stringify(shown)} במקום „${NIKUD_GHOST}”`);

    await setReader(GENESIS, { b1: LINE });
    await refreshReader();
    const hidden = (await ghostText()).trim();
    log('ניקוד מוסתר:', JSON.stringify(hidden));
    if (hidden === FIRST_GHOST) report.pass('הסתרת ניקוד בקורא — ההשלמה בלי ניקוד');
    else report.fail('הסתרת ניקוד בקורא — ההשלמה בלי ניקוד', `הוצג ${JSON.stringify(hidden)} במקום „${FIRST_GHOST}”`);
  });

  /*
   * הדיווח: „כשיש הרבה ספרים פתוחים אין קשר בין מה שהוא מציע לנכתב”. הספר
   * שנקרא אחרון (שמות) אינו הספר שמעתיקים ממנו (בראשית, בטאב של מהדורה
   * חלופית — הכותרת אינה שם הספר, ולכן הוא נמצא רק לפי `bookUid`).
   */
  await step('ספר שמעתיקים ממנו גובר על הספר שנקרא אחרון', async () => {
    const EXODUS_LINE = 'ואלה שמות בני ישראל הבאים מצרימה את יעקב איש וביתו באו';
    await setReader(
      {
        currentBook: 'שמות',
        currentBookId: 'שמות',
        bookUid: 'u2',
        currentType: 'text',
        currentRef: 'שמות א',
        currentIndex: 0,
        openTabs: [
          { toolId: null, isSelf: false, bookId: 'שמות', bookUid: 'u2', type: 'text', index: 0 },
          { toolId: null, isSelf: false, bookId: 'בראשית (מהדורה)', bookUid: 'u1', type: 'text', index: 0 },
          { toolId: 'otzaria.word', isSelf: true, bookId: 'וורד', bookUid: null, type: null, index: 0 },
        ],
      },
      { u2: EXODUS_LINE, u1: LINE },
    );
    await app.press('Enter', 'Enter', 13, 0);
    await refreshReader();

    await app.type('בראשית ב');
    await app.sleep(1_600);
    const withContext = (await ghostText()).trim();
    log('עם הקשר:', JSON.stringify(withContext));
    if (withContext === FIRST_GHOST) report.pass('הקשר תואם בספר רקע מציע ממנו');
    else report.fail('הקשר תואם בספר רקע מציע ממנו', `הוצג ${JSON.stringify(withContext)} במקום „${FIRST_GHOST}”`);

    // „וחש” מופיע רק בבראשית, ובלי שום מילת הקשר: זה בדיוק הקטע האקראי מספר
    // אחר שדווח. מותר רק מהספר שנקרא אחרון. הטענה על המשך **הספר** ולא על
    // „אין ghost”: הרשימות הסטטיות רצות אחרי הספר ומציעות לאותן אותיות משלהן.
    const BACKGROUND_CONTINUATION = 'וחשך על פני תהום';
    await app.press('Enter', 'Enter', 13, 0);
    await app.type('וחש');
    await app.sleep(1_600);
    const standalone = (await ghostText()).trim();
    log('בלי הקשר:', JSON.stringify(standalone));
    if (standalone !== BACKGROUND_CONTINUATION) report.pass('בלי הקשר אין הצעה מספר רקע');
    else report.fail('בלי הקשר אין הצעה מספר רקע', `הוצג ${JSON.stringify(standalone)} מספר הרקע`);
  });

  /*
   * אחרי ההשלמה הראשונה ארבע מילות ההקשר משותפות לשני ספרים. התאמה חדשה
   * ב-keyup של Tab הייתה מעבירה את ההמשך לספר הראשי. `app.press` שולח גם
   * keydown וגם keyup דרך CDP; ההמתנה שאחריהם מכסה גם את ה-debounce.
   */
  await step('רצף Tab נשאר בספר הרקע גם כשמילות ההקשר משותפות', async () => {
    const primary = 'commonone commontwo commonthree commonfour primarytail primaryend';
    const background = 'distinct seed commonone commontwo commonthree commonfour backgroundtail backgroundend';
    const first = 'seed commonone commontwo commonthree commonfour';
    const next = 'backgroundtail backgroundend';
    await setReader(
      {
        currentBook: 'Primary continuation book', currentBookId: 'continuation-primary',
        bookUid: 'continuation-primary', currentType: 'text', currentIndex: 0,
        currentRef: null,
        openTabs: [
          { toolId: null, isSelf: false, bookId: 'Background continuation book', bookUid: 'continuation-background', type: 'text', index: 0 },
        ],
      },
      { 'continuation-primary': primary, 'continuation-background': background },
    );
    await app.press('Enter', 'Enter', 13, 0);
    await refreshReader();
    await app.type('distinct s');
    await app.sleep(1_600);
    const initial = (await ghostText()).trim();
    if (initial !== first) {
      return report.fail('התאמה ייחודית מתחילה בספר הרקע', `הוצג ${JSON.stringify(initial)} במקום ${JSON.stringify(first)}`);
    }
    report.pass('התאמה ייחודית מתחילה בספר הרקע');

    await app.press('Tab', 'Tab', 9, 0);
    await app.sleep(1_000);
    const continued = (await ghostText()).trim();
    const afterFirst = await docText();
    log('רקע אחרי Tab ו-keyup:', JSON.stringify(continued));
    if (continued !== next || !afterFirst.endsWith(`distinct ${first}`)) {
      return report.fail('keyup של Tab שומר את מקור ההמשך', `ghost: ${JSON.stringify(continued)}, מסמך: ${JSON.stringify(afterFirst)}`);
    }
    report.pass('keyup של Tab שומר את מקור ההמשך');

    await app.press('Tab', 'Tab', 9, 0);
    await app.sleep(1_000);
    const afterSecond = await docText();
    if (afterSecond.endsWith(background)) report.pass('Tab חוזר מכניס את המשך ספר הרקע');
    else report.fail('Tab חוזר מכניס את המשך ספר הרקע', `סוף המסמך אינו ספר הרקע: ${JSON.stringify(afterSecond)}`);
    const exhausted = (await ghostText()).trim();
    if (exhausted === '') report.pass('סוף ספר הרקע אינו מתחיל התאמה חדשה');
    else report.fail('סוף ספר הרקע אינו מתחיל התאמה חדשה', `נשאר ghost: ${JSON.stringify(exhausted)}`);
  });

  /*
   * שינוי מקור בין ההצגה לבין קבלת ההצעה, בלי אירוע מהקורא ובלי keyup
   * מקדים שירענן אותו. Tab חייב לאמת את המקור לפני כתיבה למסמך האמיתי.
   */
  await step('Tab דוחה טקסט מקור שהשתנה אחרי הצגת ההצעה', async () => {
    const state = {
      currentBook: 'Mutable source book', currentBookId: 'mutable-source',
      bookUid: 'mutable-source', currentType: 'text', currentIndex: 0,
      currentRef: null, openTabs: [],
    };
    const original = 'unique seed originalone originaltwo originalthree originalfour originalend';
    const replacement = 'unique seed replacementone replacementtwo replacementthree replacementfour replacementend';
    const originalGhost = 'seed originalone originaltwo originalthree originalfour';
    const replacementGhost = 'seed replacementone replacementtwo replacementthree replacementfour';
    await setReader(state, { 'mutable-source': original });
    await app.press('Enter', 'Enter', 13, 0);
    await refreshReader();
    await app.type('unique s');
    await app.sleep(1_600);
    const initial = (await ghostText()).trim();
    if (initial !== originalGhost) {
      return report.fail('הצעת המקור לפני השינוי', `הוצג ${JSON.stringify(initial)} במקום ${JSON.stringify(originalGhost)}`);
    }
    report.pass('הצעת המקור לפני השינוי');
    const before = await docText();

    await setReader(state, { 'mutable-source': replacement });
    await app.press('Tab', 'Tab', 9, 0);
    await app.sleep(1_000);
    const after = await docText();
    if (after === before) report.pass('הצעה מיושנת אינה נכתבת למסמך');
    else report.fail('הצעה מיושנת אינה נכתבת למסמך', `לפני: ${JSON.stringify(before)}, אחרי: ${JSON.stringify(after)}`);

    const updated = (await ghostText()).trim();
    if (updated !== replacementGhost) {
      return report.fail('המקור העדכני מוצג לאחר דחיית ההצעה', `הוצג ${JSON.stringify(updated)} במקום ${JSON.stringify(replacementGhost)}`);
    }
    report.pass('המקור העדכני מוצג לאחר דחיית ההצעה');

    await app.press('Tab', 'Tab', 9, 0);
    await app.sleep(1_000);
    const accepted = await docText();
    if (accepted.endsWith(`unique ${replacementGhost}`)) report.pass('Tab נוסף מקבל את הטקסט העדכני בלבד');
    else report.fail('Tab נוסף מקבל את הטקסט העדכני בלבד', `סוף המסמך אינו ההצעה העדכנית: ${JSON.stringify(accepted)}`);
  });

  await step('כיבוי מסיר את ה-ghost', async () => {
    await app.click('השלמה מהספר', { after: 700 });
    const ghost = await ghostText();
    if (ghost === '') report.pass('הכיבוי מנקה את ה-ghost');
    else report.fail('הכיבוי מנקה את ה-ghost', `נשאר ${JSON.stringify(ghost)}`);
  });

  await step('ללא רעש', async () => {
    const lines = await app.log();
    const noisy = (lines ?? []).filter(
      (l) =>
        /error|שגיאה|failed|נכשל/i.test(l) &&
        !/DevTools|Download the Vue/i.test(l) &&
        !/reader\.addContextMenuItem|reader\.addToolbarItem|fonts\.listInstalled/.test(l) &&
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
