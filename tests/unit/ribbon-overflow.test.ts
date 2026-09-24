/**
 * ההחלטה מי מתכווץ ברצועה צרה, ובאיזה שלב (ui/ribbon/overflow.ts).
 *
 * המדידה עצמה היא DOM, אבל ההחלטה אינה: היא פונקציה מרשימת רוחבים ורוחב פנוי
 * לרשימת שלבים. זה מה שנבדק כאן — jsdom מחזיר אפס מכל `offsetWidth`, ולכן
 * ההתנהגות בדפדפן אמיתי נמדדת בשער (scripts/qa/ribbon-collapse-qa.mjs).
 */
import { describe, expect, it } from 'vitest';
import { planScale, type GroupPlan, type GroupWidth } from '../../src/ui/ribbon/overflow';

/** קבוצות בלי שלבי ביניים — רק פרוש או מכווץ, כמו לפני ההדרגה. */
function flat(...naturals: number[]): GroupWidth[] {
  return naturals.map((natural) => ({ steps: [natural, natural, natural], chip: 70 }));
}

/** קבוצה עם שלושה שלבים: פרוש, בינוני, קטן. */
function staged(natural: number, medium: number, small: number, chip = 70): GroupWidth {
  return { steps: [natural, medium, small], chip };
}

const collapsed = (plan: GroupPlan[]): boolean[] => plan.map((group) => group.collapsed);
const scales = (plan: GroupPlan[]): number[] => plan.map((group) => group.scale);

describe('planScale — הכיווץ לצ׳יפ', () => {
  it('רצועה שנכנסת אינה מכווצת דבר', () => {
    expect(collapsed(planScale(flat(200, 200, 200), 700))).toEqual([false, false, false]);
  });

  it('מכווצת מהסוף להתחלה, וכמה שצריך בלבד', () => {
    // 600 טבעי מול 500 פנוי: כיווץ אחד (‎-130) מספיק.
    expect(collapsed(planScale(flat(200, 200, 200), 500))).toEqual([false, false, true]);
    // 340 פנוי: גם השני חייב ליפול.
    expect(collapsed(planScale(flat(200, 200, 200), 340))).toEqual([false, true, true]);
  });

  it('בקצה מכווצת את כולן, ולא נכנסת ללולאה', () => {
    expect(collapsed(planScale(flat(200, 200, 200), 100))).toEqual([true, true, true]);
  });

  it('מדלגת על קבוצה שאינה רחבה מהצ׳יפ שלה', () => {
    // האמצעית היא קבוצה בת פקד אחד: כיווצה רק היה מרחיב אותה.
    const mixed: GroupWidth[] = [
      { steps: [300, 300, 300], chip: 70 },
      { steps: [50, 50, 50], chip: 70 },
      { steps: [300, 300, 300], chip: 70 },
    ];
    expect(collapsed(planScale(mixed, 400))).toEqual([true, false, true]);
  });

  it('גלישה שאי אפשר לרפא משאירה את הצרות פרושות', () => {
    const narrow: GroupWidth[] = [
      { steps: [50, 50, 50], chip: 70 },
      { steps: [50, 50, 50], chip: 70 },
      { steps: [50, 50, 50], chip: 70 },
    ];
    expect(collapsed(planScale(narrow, 100))).toEqual([false, false, false]);
  });

  it('רוחב שאינו ידוע עדיין אינו מכווץ כלום', () => {
    // `clientWidth` אפס — רצועה מכווצת או שטרם נפרסה.
    expect(planScale(flat(200, 200), 0)).toEqual([
      { scale: 0, collapsed: false },
      { scale: 0, collapsed: false },
    ]);
    expect(planScale([], 500)).toEqual([]);
  });

  it('אותו רוחב מחזיר תמיד את אותה תוכנית', () => {
    // אין כאן זיכרון של המצב הקודם, ולכן אין הבהוב בין שתי תוכניות.
    const widths = [staged(300, 220, 160), staged(250, 180, 120), staged(200, 150, 100)];
    for (let available = 100; available <= 800; available += 37) {
      expect(planScale(widths, available)).toEqual(planScale(widths, available));
    }
  });
});

describe('planScale — ההקטנה ההדרגתית', () => {
  const three = [staged(300, 200, 150), staged(300, 200, 150), staged(300, 200, 150)];

  it('מקטינה לפני שהיא מכווצת', () => {
    // 900 מול 800: הבינוני של האחרונה (‎-100) מספיק, ואף אחת אינה מתקפלת.
    const plan = planScale(three, 800);
    expect(scales(plan)).toEqual([0, 0, 1]);
    expect(collapsed(plan)).toEqual([false, false, false]);
  });

  it('כל הקבוצות עוברות שלב, מהסוף להתחלה, לפני שמישהי יורדת לשלב הבא', () => {
    // 900 מול 650: שלוש לבינוני (‎-300) הן 600 — השלישית נחוצה, ואף אחת
    // אינה יורדת ל„קטן” כל עוד יש לאחרת עוד בינוני לתת.
    expect(scales(planScale(three, 650))).toEqual([1, 1, 1]);
    // 580: עכשיו גם „קטן” — שוב מהסוף.
    expect(scales(planScale(three, 580))).toEqual([1, 1, 2]);
  });

  it('הכיווץ לצ׳יפ הוא השלב האחרון, אחרי שכולן כבר קטנות', () => {
    // שלוש קטנות הן 450. 400 מחייב צ׳יפ אחד — רק האחרונה.
    const plan = planScale(three, 400);
    expect(collapsed(plan)).toEqual([false, false, true]);
    expect(scales(plan).slice(0, 2)).toEqual([2, 2]);
  });

  it('שלב שאינו חוסך רוחב — מדלגים עליו, ולא על השלבים שאחריו', () => {
    // הבינוני של האחרונה רחב מהפרוש (תווית ארוכה לצד אייקון): אסור לבחור
    // בו. אבל ה„קטן” שלה צר — ולכן היא יורדת ישר אליו.
    const groups = [staged(300, 200, 150), staged(200, 240, 120)];
    // 500 מול 480: הבינוני של האחרונה נדחה, והראשונה היא שנותנת.
    expect(scales(planScale(groups, 480))).toEqual([1, 0]);
    // 330: בשלב השני האחרונה יורדת ישר מפרוש ל„קטן” (320), וזה מספיק.
    expect(scales(planScale(groups, 330))).toEqual([1, 2]);
  });

  it('רוחב שלא נמדד לשלב אינו נחשב כחיסכון', () => {
    const unmeasured: GroupWidth[] = [{ steps: [300], chip: 70 }];
    const plan = planScale(unmeasured, 200);
    expect(plan).toEqual([{ scale: 0, collapsed: true }]);
  });
});
