/**
 * פתיחת טאב שני, מעבר בין מסמכים וסגירת מסמך עם שינויים לא שמורים.
 * הביטול והאישור נלחצים בדיאלוג הפנימי; ui.showConfirm של ה-SDK אינו
 * משתתף במסלול הזה. הרצה: node scripts/qa/multi-doc-qa.mjs.
 */
import { openApp, sleep } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9652);
let failed = false;

function check(label, ok, detail = '') {
  console.log(`${ok ? '✓' : '✗'} ${label}${detail ? ` — ${detail}` : ''}`);
  if (!ok) failed = true;
}

async function main() {
  const page = await openApp({ name: 'multi-doc', port: PORT });
  try {
    const tabsCount = () => page.js('document.querySelectorAll(".word-doctab").length');
    const editorsSize = () => page.js('window.__otzariaEditors.size');
    const activeId = () => page.js('document.querySelector(".word-doctab.active")?.id ?? null');
    const click = async (selector, index = 0) => {
      if (!await page.clickSel(selector, index)) throw new Error(`פקד אינו נגיש: ${selector}`);
    };

    check('טאב יחיד בעליה', (await tabsCount()) === 1);
    check('window.__otzariaEditors קיים עם רשומה אחת', (await editorsSize()) === 1);
    await click('.word-doctabs-new');
    await sleep(1500);
    check('אחרי „+”: שני טאבים', (await tabsCount()) === 2);
    check('window.__otzariaEditors: שתי רשומות', (await editorsSize()) === 2);
    const ids = JSON.parse(await page.js('JSON.stringify([...document.querySelectorAll(".word-doctab")].map(e => e.id))'));
    check('המסמך החדש הוא הטאב הפעיל', (await activeId()) === ids[1]);

    // סימון דרך מנהל השמירה של המסמך החדש, ולא שינוי מלאכותי במחלקת CSS.
    await page.js(`(() => {
      const sessions = [...window.__otzariaEditors.values()];
      sessions[1].save.markDirty();
    })()`);
    await sleep(300);
    check('הטאב השני מציג שינויים לא שמורים', await page.exists('.word-doctab.active .word-doctab-dirty'));

    await click('.word-doctab', 0);
    check('מעבר לטאב הראשון החליף את המסמך הפעיל', (await activeId()) === ids[0]);
    await click('.word-doctab', 1);
    check('חזרה לטאב השני החליפה את המסמך הפעיל', (await activeId()) === ids[1]);

    await click('.word-doctab-close', 1);
    check('סגירת מסמך מלוכלך פותחת דיאלוג פנימי', await page.exists('.unsaved-dialog'));
    await click('.unsaved-dialog [data-choice="cancel"]');
    check('ביטול סוגר את הדיאלוג', !await page.exists('.unsaved-dialog'));
    check('ביטול משאיר שני טאבים ואת אותו מסמך פעיל',
      (await tabsCount()) === 2 && (await activeId()) === ids[1]);
    check('ביטול שומר את שני העורכים ואת סימון השינויים',
      (await editorsSize()) === 2 && await page.exists('.word-doctab.active .word-doctab-dirty'));

    await click('.word-doctab-close', 1);
    await click('.unsaved-dialog [data-choice="discard"]');
    // הפירוק ממתין לגיבוי המסמך; מדידה מיידית עלולה לתפוס אותו באמצע.
    for (let waited = 0; waited < 5000 && (await editorsSize()) !== 1; waited += 100) await sleep(100);
    check('„לא לשמור” מחזיר לטאב יחיד', (await tabsCount()) === 1);
    check('window.__otzariaEditors: רשומה אחת אחרי סגירה', (await editorsSize()) === 1);
    check('הטאב שנותר חוזר להיות פעיל', (await activeId()) === ids[0]);
    check('הדיאלוג נסגר אחרי אישור', !await page.exists('.unsaved-dialog'));
  } finally {
    page.close();
  }

  if (failed) {
    console.error('\nהשער נכשל — ראו ✗ למעלה.');
    process.exit(1);
  }
  console.log('\nהשער עבר: פתיחת טאב, מעבר, וסגירה (ביטול ואישור) עובדים כצפוי.');
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
