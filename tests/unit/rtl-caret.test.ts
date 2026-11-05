/**
 * חצים אופקיים בשורה עברית — התיקון.
 *
 * **המספרים כאן אינם מומצאים.** כולם נמדדו ב-Chrome אמיתי על ה-`dist` הארוז
 * בגשש `scripts/qa/rtl-caret-visual-feasibility.mjs`, שהציב את הסמן בכל היסט
 * בתורו וקרא את ה-x שהמנוע צייר בו — על superdoc 2.14.0-next.5 ושוב על 2.15.0,
 * עם אותם ערכים בדיוק. תיבות התווים כאן נגזרות מאותן מדידות, וה„אמת” שבכל
 * בדיקה היא מה שהמנוע עצמו עשה.
 *
 * שלוש השורות שנמדדו וחוזרות כאן:
 *   „שלום עולם פשוט מאוד” — עברית נקייה, התאמה 20/20
 *   „פרק 12 בספר”        — אי של ספרות בתוך ריצה עברית אחת, 12/12
 *   „מילה ABC מילה”      — אי לטיני מוצהר, 12/14 (שני התפרים)
 *
 * ושתי שורות עם אי של **תו אחד** („סעיף 3 בחוק הזה”, „אות a אחת בלבד”), שנמדדו
 * באותה שיטה כשהתברר שהחץ נתקע בהן — ראו ההערה על `charDirections`.
 *
 * וגם צורת הציור של טאב, שנמדדה ב-`scripts/qa/tab-paint-probe.mjs`:
 * ‏`SPAN.superdoc-tab` עם `data-pm-start`/`data-pm-end` משלו, בלי טקסט.
 */
import { describe, it, expect, vi, afterEach } from 'vitest';
/*
 * הקריאה של מה שצויר עברה ל-`painted-lines.ts`, והמדיניות נשארה
 * ב-`rtl-caret.ts`. הבדיקות של שתיהן נשארו יחד כאן **בכוונה**: הן חולקות את
 * אותן תיבות שנמדדו (`LATIN_ISLAND`, `DIGIT_ISLAND`, השורות מהגשש), ופיצול
 * הקובץ היה מכפיל אותן — כלומר מייצר שני מקורות אמת למדידה אחת. מה שמפריד
 * בין השניים הוא שם הייבוא שלמטה.
 */
import {
  edgeSlot,
  installRtlVisualArrows,
  isHorizontalArrow,
  movesForward,
  visualTarget,
  type RtlCaretHost,
} from '../../src/engine/rtl-caret';
import {
  caretSlots,
  charDirections,
  readLineChars,
  type CaretSlot,
  type PaintedChar,
} from '../../src/engine/painted-lines';

/**
 * תיבות מרשימת קצוות, עם היסטים רצופים מ-`pm`.
 * `rtl` הופך את הקצוות — ברשימה של שורה עברית הם יורדים.
 */
function boxes(pm: number, edges: number[], rtl: boolean): PaintedChar[] {
  return edges.slice(0, -1).map((edge, i) => {
    const next = edges[i + 1]!;
    return { pm: pm + i, left: rtl ? next : edge, right: rtl ? edge : next };
  });
}

/**
 * „מילה ABC מילה”, ‏pm 51..64. שלוש ריצות: עברית, אי לטיני, עברית.
 * האי מצויר ב-[980 … 1012.9], והמילה העברית הראשונה מימין עד 1043.4.
 */
const LATIN_ISLAND: PaintedChar[] = [
  ...boxes(51, [1043.4, 1035.6, 1031.3, 1024.4, 1016.9, 1012.9], true),
  ...boxes(56, [980, 991.5, 1002.2, 1012.9], false),
  ...boxes(59, [980, 976, 968.2, 963.9, 957, 949.5], true),
];

/** „פרק 12 בספר”, ‏pm 39..50 — האי הוא ספרות בתוך ריצה עברית **אחת**. */
const DIGIT_ISLAND: PaintedChar[] = [
  ...boxes(39, [1043.4, 1036.3, 1028.8, 1021.2, 1017.2], true),
  ...boxes(43, [1001.2, 1009.2, 1017.2], false),
  ...boxes(45, [1001.2, 997.2, 989.9, 982.3, 975.3, 967.8], true),
];

/** „שלום עולם פשוט מאוד”, ‏pm 1..20 — עברית נקייה, 19 תווים. */
const CLEAN: PaintedChar[] = boxes(
  1,
  [
    1043.4, 1033.6, 1026.7, 1022.2, 1014.1, 1010.1, 1002.8, 998.4, 991.5, 983.4, 979.4, 972.3,
    962.5, 958.1, 950.5, 946.5, 938.7, 931, 926.5, 919.8,
  ],
  true,
);

/** תיבות עם התו עצמו — כך ש-`readLineChars` מחזירה אותן. */
function lettered(pm: number, text: string, edges: [number, number][]): PaintedChar[] {
  return [...text].map((ch, i) => ({ pm: pm + i, left: edges[i]![0], right: edges[i]![1], ch }));
}

/**
 * „סעיף 3 בחוק הזה”, ‏pm 1..15 — ספרה **בודדת** בין שני רווחים. נמדד ב-Chrome
 * על ה-dist (superdoc 2.15.0): התיבות, ובנפרד ה-x שהמנוע צייר בו את הסמן בכל
 * היסט. שני התפרים של הספרה (היסטים 5 ו-6) מצוירים שניהם בקצה השמאלי שלה.
 */
const LONE_DIGIT = lettered(1, 'סעיף 3 בחוק הזה', [
  [1035.8, 1043.4], [1028.5, 1035.8], [1024.2, 1028.5], [1017, 1024.2], [1013, 1017],
  [1005, 1013], [1000.9, 1005], [993.7, 1001], [986, 993.7], [981.6, 986],
  [974, 981.6], [970, 974], [962.5, 970], [957.9, 962.5], [950.3, 957.9],
]);
const LONE_DIGIT_PAINTED = [
  1043.3, 1035.8, 1028.5, 1024.2, 1017, 1005, 1004.9, 1000.9, 993.7, 986, 981.6, 974, 970, 962.5, 957.9,
  950.3,
];

/**
 * „אות a אחת בלבד”, ‏pm 36..49 — אות לטינית **בודדת**. כאן שני התפרים (היסטים
 * 4 ו-5) מצוירים שניהם בקצה הימני של האות.
 */
const LONE_LETTER = lettered(36, 'אות a אחת בלבד', [
  [1035.6, 1043.4], [1031.2, 1035.7], [1023.2, 1031.2], [1019.2, 1023.2], [1012.1, 1019.2],
  [1008.1, 1012.1], [1000.4, 1008.1], [992.8, 1000.4], [984.8, 992.8], [980.8, 984.8],
  [973.5, 980.8], [966.6, 973.5], [959.3, 966.6], [952.6, 959.3],
]);
const LONE_LETTER_PAINTED = [
  1043.3, 1035.6, 1031.2, 1023.2, 1019.2, 1019.2, 1008.1, 1000.4, 992.8, 984.8, 980.8, 973.5, 966.6, 959.3,
  952.6,
];

/**
 * „שלוש × ארבע שווה”, ‏pm 1..17, ו„חמש ÷ שתיים”, ‏pm 18..29 — סימן **בודד**
 * בין שני רווחים, כמו האי של תו אחד שלמעלה, אבל מסוג אחר: `×` ו-`÷` הם
 * ניטרליים (ON) ב-UBA, ולכן הדפדפן מצייר אותם בכיוון הפסקה. התיבה של הסימן
 * נוגעת בשכנותיה **בצד הימני** משני הצדדים, שני התפרים שלו מצוירים ב-x שונה
 * (617.8 ו-608.8), ואין כאן שום נקודה עיוורת שצריך לעקוף.
 *
 * נמדד ב-Chrome על ה-dist (superdoc 2.15.0) ב-`scratchpad/measure-neutral.mjs`,
 * באותה שיטה: התיבות מ-`Range.getBoundingClientRect`, וה-x המצויר מהצבת
 * `setSelectionTarget` בכל היסט בתורו.
 */
const TIMES_SIGN = lettered(1, 'שלוש × ארבע שווה', [
  [642.9, 652.7], [636, 643], [631.6, 636], [621.8, 631.6], [617.8, 621.8],
  [608.8, 617.8], [604.8, 608.8], [597.1, 604.8], [589.6, 597.1], [582.3, 589.6],
  [575, 582.3], [571, 575], [561.3, 571], [556.8, 561.3], [552.3, 556.8], [544.8, 552.4],
]);
const TIMES_SIGN_PAINTED = [
  652.7, 642.9, 636, 631.6, 621.8, 617.8, 608.8, 604.8, 597.1, 589.6, 582.3, 575, 571, 561.3,
  556.8, 552.3, 544.8,
];

const DIVISION_SIGN = lettered(18, 'חמש ÷ שתיים', [
  [645.1, 652.7], [637.3, 645.1], [627.5, 637.3], [623.5, 627.6], [614.8, 623.6],
  [610.8, 614.8], [601, 610.8], [593, 601], [588.7, 593], [584.5, 588.8], [576.4, 584.5],
]);
const DIVISION_SIGN_PAINTED = [
  652.7, 645.1, 637.3, 627.5, 623.5, 614.8, 610.8, 601, 593, 588.7, 584.5, 576.4,
];

/**
 * „ABC מילה כאן” בפסקה `w:bidi`, ‏pm 1..13 — שורה שמ**תחילה** באי לטיני.
 *
 * נמדד בשער (`check:arrows`, superdoc 2.15.0, 17.9.2026): ה-x שהמנוע צייר בו
 * את הסמן בכל היסט, בהליכה מלאה לשני הכיוונים. האי יושב בקצה הימני של השורה
 * ‏[1010.5 … 1043.3], והעברית ממשיכה שמאלה ממנו.
 */
const STARTS_LATIN: PaintedChar[] = [
  ...lettered(1, 'ABC', [[1010.5, 1022], [1022, 1032.7], [1032.7, 1043.3]]),
  ...lettered(4, ' מילה כאן', [
    [1006.4, 1010.5], [998.7, 1006.4], [994.4, 998.7], [987.5, 994.4], [980, 987.5],
    [976, 980], [969, 976], [961.3, 969], [957, 961.3],
  ]),
];
/** ה-x המצויר לכל היסט 0..12 של אותה שורה, מאותה ריצה. */
const STARTS_LATIN_PAINTED = [
  1010.5, 1022, 1032.7, 1043.3, 1006.4, 998.7, 994.4, 987.5, 980, 976, 969, 961.3, 957,
];

/**
 * „שלום😀עולם” — אימוג'י הוא נקודת קוד אחת ו**שתי** יחידות UTF-16.
 *
 * התיבות נמדדו ב-Chrome (16pt serif, `dir=rtl`): חיתוך `Range` לכל **יחידה**
 * בנפרד מחזיר לשני חצאי הזוג את **אותה תיבה מלאה** ‏[531.7, 561] ברוחב 29.3 —
 * ולא מלבן ריק. כלומר בלי איטרציה לפי אשכול נוצר חריץ שני באותו x, שההיסט
 * שלו יושב בתוך נקודת הקוד.
 *
 * המדידה חיה ב-`scripts/qa/rtl-caret-cluster-probe.mjs` וניתן להריץ אותה
 * שוב (‏`node scripts/qa/rtl-caret-cluster-probe.mjs`); קודם היא צוטטה מקובץ
 * ‏`scratchpad/surrogate.html` שאינו במאגר, ולכן לא ניתן היה לאמת אותה.
 */
const EMOJI_BOXES: [number, number][] = [
  [587, 600], [577.8, 587], [571.8, 577.8], [561, 571.8],
  [531.7, 561], [531.7, 561],
  [522, 531.7], [516.1, 522], [506.9, 516.1], [496, 506.9],
];

/**
 * „מילה abcשָׁלום” — אות עברית עם קמץ ושין-ימנית, בשורה שיש בה גם אי לטיני.
 *
 * התיבות נמדדו ב-Chrome ב-`scripts/qa/rtl-caret-cluster-probe.mjs`. כל שלוש
 * היחידות של „שָׁ” (ש, U+05B8, U+05C1) מחזירות את **אותה** תיבה
 * ‏[637.5, 647.3] — וזו הצורה שבה הדפדפן עונה על חיתוך בתוך אשכול. הרשימה
 * כאן היא לפי **יחידת UTF-16**, כלומר בדיוק מה ש-`stubRects` יגיש לכל
 * ‏`startOffset`, וממנה `readLineChars` אמורה להרכיב 12 אשכולות ולא 14.
 */
const NIQQUD_BOXES: [number, number][] = [
  [692.2, 700], [688, 692.2], [681, 688], [673.5, 681.1], [669.5, 673.5],
  [647.3, 654.4], [654.4, 662.4], [662.4, 669.5],
  [637.5, 647.3], [637.5, 647.3], [637.5, 647.3],
  [630.6, 637.5], [626.2, 630.6], [618.1, 626.2],
];
const NIQQUD_TEXT = 'מילה abcשָׁלום';

/**
 * „מילה éa מילה” — אות לטינית מפורקת (e ואחריה U+0301) בתוך שורה עברית.
 * שתי היחידות שלה מחזירות את אותה תיבה [655.3, 662.4], מאותה מדידה.
 */
const DECOMPOSED_BOXES: [number, number][] = [
  [692.2, 700], [688, 692.2], [681, 688], [673.5, 681.1], [669.5, 673.5],
  [655.3, 662.4], [655.3, 662.4], [662.4, 669.5],
  [651.3, 655.3], [643.5, 651.3], [639.3, 643.5], [632.4, 639.3], [624.8, 632.4],
];
const DECOMPOSED_TEXT = 'מילה éa מילה';

/**
 * ההיסטים החוקיים לסמן בשתי השורות: תחילת כל אשכול, ועוד ההיסט שאחרי האחרון.
 *
 * הרשימות **אינן** נגזרות מ-`Intl.Segmenter` בזמן הבדיקה, ובכוונה: המימוש
 * נשען עליו, ובדיקה שתשאל אותו הייתה מאשרת את עצמה. הן הועתקו מפלט הגשש
 * ‏`scripts/qa/rtl-caret-cluster-probe.mjs` שרץ ב-Chrome — כלומר מהמדידה.
 * ‏„שָׁ” הוא אשכול אחד בהיסט 8 ולכן 9 ו-10 חסרים; ‏`e`+U+0301 הוא אשכול אחד
 * בהיסט 5 ולכן 6 חסר.
 */
const NIQQUD_LEGAL = [0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 13, 14];
const DECOMPOSED_LEGAL = [0, 1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12, 13];

const pmsOf = (slots: readonly CaretSlot[]) => slots.map((s) => s.pm);
const xOf = (slots: readonly CaretSlot[], pm: number) => slots.find((s) => s.pm === pm)?.x;

/**
 * הקשה אחר הקשה, מול מה שהמנוע באמת מצייר: נקודת המוצא של כל הקשה היא ה-x
 * המצויר של ההיסט הנוכחי, בדיוק כמו ב-`installRtlVisualArrows`.
 */
function hold(
  chars: readonly PaintedChar[],
  painted: readonly number[],
  base: number,
  from: number,
  toRight: boolean,
): number[] {
  const slots = caretSlots(chars, true);
  const visited = [from];
  let offset = from;
  for (let step = 0; step < chars.length + 2; step += 1) {
    const target = visualTarget(slots, painted[offset]!, toRight);
    if (!target) break;
    const next = target.pm - base;
    if (next === offset) break;
    visited.push(next);
    offset = next;
  }
  return visited;
}

/** הליכה מכל היסט, בשני הכיוונים: בלי ביקור חוזר, וכל צעד מתקדם על המסך. */
function expectWalksForward(
  chars: readonly PaintedChar[],
  painted: readonly number[],
  base: number,
): void {
  for (let from = 0; from < painted.length; from += 1) {
    for (const toRight of [true, false]) {
      const visited = hold(chars, painted, base, from, toRight);
      expect(new Set(visited).size, `ביקור חוזר מ-${from}`).toBe(visited.length);
      for (let i = 1; i < visited.length; i += 1) {
        const dx = painted[visited[i]!]! - painted[visited[i - 1]!]!;
        expect(toRight ? dx : -dx, `צעד ${visited[i - 1]}→${visited[i]}`).toBeGreaterThan(0.5);
      }
    }
  }
}

describe('charDirections', () => {
  it('עברית נקייה — כל התווים ימין-לשמאל', () => {
    expect(charDirections(CLEAN, true).every((rtl) => rtl)).toBe(true);
  });

  it('אי לטיני מוצהר — שלושת תוויו נגזרים כשמאל-לימין', () => {
    expect(charDirections(LATIN_ISLAND, true)).toEqual([
      true, true, true, true, true, false, false, false, true, true, true, true, true,
    ]);
  });

  it('אי של ספרות בתוך ריצה עברית אחת — אותה הכרעה, בלי שום הצהרה שתעזור', () => {
    // זה המקרה שחלוקה לפי אלמנט DOM או לפי `w:rtl` מחמיצה: הריצה אחת, ורק
    // הגאומטריה מפרידה. ראו `bidi-declaration-is-not-direction`.
    expect(charDirections(DIGIT_ISLAND, true)).toEqual([
      true, true, true, true, false, false, true, true, true, true, true,
    ]);
  });

  it('שורה לטינית — כל התווים שמאל-לימין', () => {
    expect(charDirections(boxes(1, [441.6, 449.6, 454.1, 461.2, 465.6], false), false)).toEqual([
      false, false, false, false,
    ]);
  });
});

describe('caretSlots', () => {
  it('עברית נקייה: חריץ לכל היסט, ובדיוק ב-x שהמנוע צייר בו', () => {
    const slots = caretSlots(CLEAN, true);
    expect(slots).toHaveLength(CLEAN.length + 1);
    // שלושת אלה נמדדו מול המנוע, אחד לאחד.
    expect(xOf(slots, 1)).toBe(1043.4);
    expect(xOf(slots, 9)).toBe(991.5);
    expect(xOf(slots, 20)).toBe(919.8);
  });

  it('אי לטיני: שני התפרים מושמטים, ופנים האי נשאר', () => {
    const slots = caretSlots(LATIN_ISLAND, true);
    // 56 ו-59 הם ההיסטים שהמנוע צייר את שניהם באותו x (1012.9) — שני בתים
    // לאותו מקום, ובלי ידית לבחור. כיוון אליהם היה מזיז לכיוון ההפוך.
    expect(pmsOf(slots)).not.toContain(56);
    expect(pmsOf(slots)).not.toContain(59);
    expect(xOf(slots, 57), 'פנים האי — נמדד 991.5 במנוע').toBe(991.5);
    expect(xOf(slots, 58), 'פנים האי — נמדד 1002.2 במנוע').toBe(1002.2);
    expect(xOf(slots, 60), 'העברית שמשמאל לאי').toBe(976);
  });

  it('אי של ספרות: אותם שני תפרים מושמטים', () => {
    const slots = caretSlots(DIGIT_ISLAND, true);
    expect(pmsOf(slots)).not.toContain(43);
    expect(pmsOf(slots)).not.toContain(45);
    expect(xOf(slots, 44), 'בין שתי הספרות').toBe(1009.2);
    expect(xOf(slots, 42)).toBe(1021.2);
  });

  it('שורה ריקה — אין חריצים', () => {
    expect(caretSlots([], true)).toEqual([]);
  });

  it('אי לטיני שפותח את השורה: ההיסט הראשון הוא חריץ, ובדיוק ב-x שהמנוע צייר', () => {
    // הקצה **אינו** תפר, וזה נמדד: אילו דילגנו עליו (כאילו הכיוון משתנה מול
    // הפסקה) היינו מוחקים מקום שהמנוע מצייר נכון, והחץ היה קופץ מעל „A”.
    const slots = caretSlots(STARTS_LATIN, true);
    expect(xOf(slots, 1), 'היסט 0 — נמדד 1010.5 במנוע').toBe(1010.5);
    expect(xOf(slots, 2)).toBe(1022);
    expect(xOf(slots, 3)).toBe(1032.7);
    // התפר האמיתי היחיד בשורה הזאת: המעבר מהאי לעברית (היסט 3).
    expect(pmsOf(slots)).not.toContain(4);
  });

  it('אי לטיני שסוגר את השורה: חריץ הסיום הוא הקצה הנגרר של הספרה', () => {
    // אותה שאלה בקצה השני. נמדד על „לפני הטבלה 4”: היסט הסיום מצויר ב-975.4,
    // שהוא הקצה הימני של הספרה — ולא בקצה השמאלי של השורה.
    const chars = lettered(1, 'לפני 4', [
      [1036.5, 1043.4], [1029.4, 1036.5], [1024.5, 1029.4], [1020.3, 1024.5],
      [975.4, 979.4], [967.8, 975.4],
    ]);
    const slots = caretSlots(chars, true);
    expect(pmsOf(slots), 'התפר שלפני הספרה מדולג').not.toContain(6);
    expect(xOf(slots, 7), 'חריץ הסיום — הקצה הימני של הספרה').toBe(975.4);
  });

  it('הליכה מלאה על שורה שמתחילה באי לטיני — כל צעד מתקדם על המסך', () => {
    expectWalksForward(STARTS_LATIN, STARTS_LATIN_PAINTED, 1);
  });
});

describe('אי של תו אחד', () => {
  const cases = [
    { name: 'ספרה', chars: LONE_DIGIT, painted: LONE_DIGIT_PAINTED, base: 1, island: 5 },
    { name: 'אות לטינית', chars: LONE_LETTER, painted: LONE_LETTER_PAINTED, base: 36, island: 4 },
  ];

  for (const c of cases) {
    it(`${c.name} בודדת מסווגת כשמאל-לימין, ושכניה נשארים עבריים`, () => {
      const dirs = charDirections(c.chars, true);
      expect(dirs[c.island]).toBe(false);
      expect(dirs[c.island - 1]).toBe(true);
      expect(dirs[c.island + 1]).toBe(true);
    });

    it(`${c.name} בודדת: שני התפרים שלה אינם יעד`, () => {
      const pms = pmsOf(caretSlots(c.chars, true));
      expect(pms).not.toContain(c.base + c.island);
      expect(pms).not.toContain(c.base + c.island + 1);
    });

    it(`${c.name} בודדת: החזקת חץ מכל היסט מתקדמת על המסך בכל הקשה, בלי לחזור`, () => {
      // זה ה„נתקע” שנמדד: שני התפרים מצוירים באותו x, והקשה אחת החזירה את
      // הסמן לתפר השני במקום לעבור את התו.
      expectWalksForward(c.chars, c.painted, c.base);
    });
  }
});

describe('סימן ניטרלי אינו אי לועזי', () => {
  const cases = [
    { name: '×', chars: TIMES_SIGN, painted: TIMES_SIGN_PAINTED, base: 1, sign: 5 },
    { name: '÷', chars: DIVISION_SIGN, painted: DIVISION_SIGN_PAINTED, base: 18, sign: 4 },
  ];

  for (const c of cases) {
    it(`${c.name} נשאר בכיוון השורה — הגאומטריה כבר הכריעה`, () => {
      expect(charDirections(c.chars, true).every((rtl) => rtl), 'כל השורה ימנית').toBe(true);
    });

    it(`שני התפרים של ${c.name} נשארים יעד, כל אחד ב-x שלו`, () => {
      const slots = caretSlots(c.chars, true);
      const before = c.base + c.sign;
      expect(pmsOf(slots)).toContain(before);
      expect(pmsOf(slots)).toContain(before + 1);
      // עיגול תת-פיקסל מפריד את קצה התיבה מה-x המצויר ב-0.1, כמו בקיבועים שלמעלה.
      expect(xOf(slots, before)!).toBeCloseTo(c.painted[c.sign]!, 0);
      expect(xOf(slots, before + 1)!).toBeCloseTo(c.painted[c.sign + 1]!, 0);
      expect(
        Math.abs(xOf(slots, before)! - xOf(slots, before + 1)!),
        'שני תפרים נפרדים, ולא שניהם באותו x כמו באי של תו אחד',
      ).toBeGreaterThan(0.5);
    });

    it(`החזקת חץ עוברת דרך ${c.name} תו-תו, בלי לדלג`, () => {
      expectWalksForward(c.chars, c.painted, c.base);
      // דילוג היה נראה כצעד יחיד שחוצה שני היסטים; כאן כל היסט הוא תחנה.
      expect(hold(c.chars, c.painted, c.base, c.sign - 1, false).slice(0, 3)).toEqual([
        c.sign - 1, c.sign, c.sign + 1,
      ]);
    });
  }
});

describe('visualTarget', () => {
  it('עברית נקייה: ימין מזיז ימינה על המסך, וההיסט יורד — כמו Word', () => {
    const slots = caretSlots(CLEAN, true);
    const from = xOf(slots, 9)!;
    const target = visualTarget(slots, from, true)!;
    expect(target.pm, 'היסט אחד אחורה').toBe(8);
    expect(target.x).toBeGreaterThan(from);
  });

  it('עברית נקייה: שמאל מזיז שמאלה, וההיסט עולה', () => {
    const slots = caretSlots(CLEAN, true);
    const from = xOf(slots, 9)!;
    const target = visualTarget(slots, from, false)!;
    expect(target.pm).toBe(10);
    expect(target.x).toBeLessThan(from);
  });

  it('הליכה שלמה על השורה: ימין לעולם אינו מזיז שמאלה', () => {
    // זו התכונה שהמשתמש מרגיש. שבירתה היא הבאג עצמו.
    for (const chars of [CLEAN, LATIN_ISLAND, DIGIT_ISLAND]) {
      const slots = caretSlots(chars, true);
      let x = Math.min(...slots.map((s) => s.x));
      let steps = 0;
      for (;;) {
        const next = visualTarget(slots, x, true);
        if (!next) break;
        expect(next.x, 'כל צעד ימינה גדל ב-x').toBeGreaterThan(x);
        x = next.x;
        steps += 1;
      }
      expect(steps, 'ההליכה עוברת את כל החריצים').toBe(slots.length - 1);
    }
  });

  it('כניסה לאי הלטיני מצד שמאל נוחתת בתוכו, ולא בצדו השני', () => {
    // 60@976 → 57@991.5. זה בדיוק מה שהמנוע מצייר ל-57 (נמדד), ולכן הצעד
    // אמיתי. הגבול ב-980 מדולג — זה המחיר המתועד של דילוג התפר.
    const slots = caretSlots(LATIN_ISLAND, true);
    expect(visualTarget(slots, 976, true)).toMatchObject({ pm: 57, x: 991.5 });
  });

  it('אין חריץ בכיוון המבוקש — `null`, והקצה מטופל בנפרד', () => {
    const slots = caretSlots(CLEAN, true);
    const rightmost = Math.max(...slots.map((s) => s.x));
    expect(visualTarget(slots, rightmost, true)).toBeNull();
  });
});

describe('movesForward', () => {
  it('בשורה עברית ימינה הוא אחורה בטקסט, ובלטינית להפך', () => {
    expect(movesForward(true, true), 'ימין בעברית').toBe(false);
    expect(movesForward(false, true), 'שמאל בעברית').toBe(true);
    expect(movesForward(true, false), 'ימין בלטינית').toBe(true);
    expect(movesForward(false, false), 'שמאל בלטינית').toBe(false);
  });
});

describe('edgeSlot', () => {
  const slots = caretSlots(CLEAN, true);

  it('אחורה: הקצה הגבוה של השורה שמעל — כלומר סופה', () => {
    expect(edgeSlot(slots, false, 999)).toMatchObject({ pm: 20 });
  });

  it('קדימה: תחילתה', () => {
    expect(edgeSlot(slots, true, 999)).toMatchObject({ pm: 1 });
  });

  it('ההיסט שהסמן כבר עליו מוחרג — גבול בין שתי שורות גולשות הוא היסט אחד', () => {
    // בלי ההחרגה הצעד היה מחזיר את הסמן בדיוק לאותו מקום, והמשתמש היה
    // רואה מקש שאינו עושה דבר.
    expect(edgeSlot(slots, false, 20)).toMatchObject({ pm: 19 });
  });

  it('שורה שכל חריציה הם ההיסט הנוכחי — `null`', () => {
    expect(edgeSlot([{ pm: 7, x: 10 }], false, 7)).toBeNull();
  });
});

describe('isHorizontalArrow', () => {
  const key = (over: Partial<KeyboardEvent> = {}) =>
    ({ key: 'ArrowRight', ctrlKey: false, metaKey: false, altKey: false, ...over }) as KeyboardEvent;

  it('חץ ימני ושמאלי', () => {
    expect(isHorizontalArrow(key())).toBe(true);
    expect(isHorizontalArrow(key({ key: 'ArrowLeft' }))).toBe(true);
  });

  it('Shift+חץ הוא שלנו — גם הוא יורש את הסתירה של המנוע', () => {
    expect(isHorizontalArrow(key({ shiftKey: true }))).toBe(true);
  });

  it('Ctrl/Alt/Meta הם קפיצת מילה — לא שלנו', () => {
    expect(isHorizontalArrow(key({ ctrlKey: true }))).toBe(false);
    expect(isHorizontalArrow(key({ altKey: true }))).toBe(false);
    expect(isHorizontalArrow(key({ metaKey: true }))).toBe(false);
  });

  it('חצים אנכיים אינם נוגעים', () => {
    expect(isHorizontalArrow(key({ key: 'ArrowUp' }))).toBe(false);
    expect(isHorizontalArrow(key({ key: 'ArrowDown' }))).toBe(false);
  });
});

/* ------------------------------------------------------------------ */
/* קריאת ה-DOM וההתקנה                                                 */
/* ------------------------------------------------------------------ */

/**
 * jsdom אינו פורס, ולכן `getBoundingClientRect` שלו מחזיר אפסים — גם ל-
 * `Range`. הדריסה כאן מזינה את המדידות האמיתיות במקומן: התיבות נשמרות על
 * אלמנט הריצה, ו-Range קורא מהן לפי ההיסט שלו.
 */
const originalRangeRect = Range.prototype.getBoundingClientRect;

function stubRects(): void {
  Range.prototype.getBoundingClientRect = function (this: Range) {
    const parent = (this.startContainer as Text).parentElement as
      | (HTMLElement & { __chars?: PaintedChar[] })
      | null;
    const box = parent?.__chars?.[this.startOffset];
    return { left: box?.left ?? 0, right: box?.right ?? 0 } as DOMRect;
  };
}

afterEach(() => {
  Range.prototype.getBoundingClientRect = originalRangeRect;
  document.body.innerHTML = '';
  vi.restoreAllMocks();
});

interface RunSpec {
  text?: string;
  chars: PaintedChar[];
  /** ריצה שאינה טקסט המסמך — סמן המספור של פריט רשימה. */
  marker?: boolean;
  /** טאב: אלמנט בלי טקסט, שהתיבה שלו היא שלו עצמו. */
  tab?: boolean;
}

/** בונה שורה מצוירת אחת בצורה שנמדדה. */
function buildLine(
  runs: RunSpec[],
  pmStart: number,
  pmEnd: number,
  rtl: boolean,
  top = 0,
): HTMLElement {
  const line = document.createElement('div');
  line.setAttribute('data-pm-start', String(pmStart));
  line.setAttribute('data-pm-end', String(pmEnd));
  if (rtl) line.setAttribute('dir', 'rtl');
  line.getBoundingClientRect = () =>
    ({ top, bottom: top + 20, left: 51, right: 1043.4 }) as DOMRect;

  for (const run of runs) {
    const span = document.createElement('span') as HTMLElement & { __chars?: PaintedChar[] };
    if (run.tab) span.className = 'superdoc-tab';
    else if (!run.marker) span.className = 'superdoc-text-run';
    if (!run.marker && run.chars.length) {
      span.setAttribute('data-pm-start', String(run.chars[0]!.pm));
      span.setAttribute('data-pm-end', String(run.chars[run.chars.length - 1]!.pm + 1));
    }
    if (run.tab) {
      const box = run.chars[0]!;
      span.getBoundingClientRect = () => ({ left: box.left, right: box.right }) as DOMRect;
    } else {
      span.textContent = run.text ?? '';
      span.__chars = run.chars;
    }
    line.append(span);
  }
  return line;
}

/** מה ש-`readLineChars` קוראת מהשורה המעורבת: אותן תיבות, עם התו עצמו. */
const LATIN_ISLAND_READ: PaintedChar[] = LATIN_ISLAND.map((box, i) => ({ ...box, ch: [...'מילה ABC מילה'][i] }));

/** אותה שורה מעורבת שנמדדה, מפורקת לשלוש ריצות כפי שהיא מצוירת. */
const mixedRuns = (): RunSpec[] => [
  { text: 'מילה ', chars: LATIN_ISLAND.slice(0, 5) },
  { text: 'ABC', chars: LATIN_ISLAND.slice(5, 8) },
  { text: ' מילה', chars: LATIN_ISLAND.slice(8) },
];

describe('readLineChars', () => {
  it('מחזירה תיבה והיסט לכל תו, לפי הטווח של הריצה עצמה', () => {
    stubRects();
    const line = buildLine(mixedRuns(), 51, 64, true);
    document.body.append(line);
    expect(readLineChars(line)).toEqual(LATIN_ISLAND_READ);
  });

  it('טאב הוא יחידת pm אחת, והתיבה היא של האלמנט עצמו', () => {
    // נמדד: <SPAN.superdoc-tab> pm=5..6, רוחב 24, אפס תווים. בלי הטיפול הזה
    // השורה נפסלת כולה, וההקשה חוזרת למנוע — כלומר הבאג נשאר בפסקה עם טאב.
    stubRects();
    const line = buildLine(
      [
        { text: 'קודם', chars: boxes(1, [1043, 1035, 1027, 1019, 1011], true) },
        { tab: true, chars: [{ pm: 5, left: 987, right: 1011 }] },
        { text: 'אחרון', chars: boxes(6, [987, 979, 971, 963, 955, 947], true) },
      ],
      1,
      11,
      true,
    );
    document.body.append(line);
    const chars = readLineChars(line);
    expect(chars).toHaveLength(10);
    expect(chars![4]).toEqual({ pm: 5, left: 987, right: 1011 });
  });

  it('סמן המספור של פריט רשימה אינו נספר — אין לו טווח pm', () => {
    stubRects();
    const line = buildLine(
      [{ marker: true, chars: [] }, ...mixedRuns()],
      51,
      64,
      true,
    );
    document.body.append(line);
    expect(readLineChars(line)).toEqual(LATIN_ISLAND_READ);
  });

  it('טאב-הסיומת של סמן הרשימה אינו נושא טווח, ואינו נקרא כטווח 0..0', () => {
    // נמדד: <SPAN.superdoc-tab.superdoc-marker-suffix-tab> עם רווח אחד ובלי
    // data-pm. `Number(null)` הוא 0, ולכן בלי בדיקת קיום התכונה הוא היה
    // מתחזה לטווח 0..0 ופוסל את כל השורה — כלומר מחזיר את הבאג לכל רשימה.
    stubRects();
    const line = buildLine(mixedRuns(), 51, 64, true);
    const suffix = document.createElement('span');
    suffix.className = 'superdoc-tab superdoc-marker-suffix-tab';
    suffix.textContent = ' ';
    line.prepend(suffix);
    document.body.append(line);
    expect(readLineChars(line)).toEqual(LATIN_ISLAND_READ);
  });

  it("אימוג'י הוא תו אחד ושתי יחידות pm — ולא שני חריצים באותו x", () => {
    // נמדד ב-Chrome: שני חצאי הזוג מחזירים את אותה תיבה מלאה [531.7, 561].
    // בלי איטרציה לפי נקודת קוד נוצר כאן חריץ שני, שההיסט שלו (6) יושב
    // **בתוך** נקודת הקוד — ובאי לועזי, שבו הראשון נופל כתפר, הוא זה שהיה
    // נכתב למנוע.
    stubRects();
    const run = EMOJI_BOXES.map(([left, right], i) => ({ pm: 1 + i, left, right }));
    const line = buildLine([{ text: 'שלום😀עולם', chars: run }], 1, 11, true);
    document.body.append(line);

    const chars = readLineChars(line)!;
    expect(chars).toHaveLength(9);
    expect(chars[4]).toEqual({ pm: 5, left: 531.7, right: 561, ch: '😀', units: 2 });
    expect(chars[5]!.pm, 'ההיסט הבא מדלג שתי יחידות').toBe(7);

    const slots = caretSlots(chars, true);
    expect(pmsOf(slots)).toEqual([1, 2, 3, 4, 5, 7, 8, 9, 10, 11]);
    expect(charDirections(chars, true).every((rtl) => rtl), "האימוג'י אינו אי לועזי").toBe(true);
  });

  it('אות עברית עם ניקוד וטעם היא אשכול אחד — ואין חריץ בין האות לניקודה', () => {
    /*
     * זה הדבר שמשתמש מרגיש: פירוש עברי נכתב מנוקד, ואילו נשאר חריץ בין „ש”
     * לקמץ שלה — הקשת חץ הייתה מניחה שם את הסמן, והתו הבא היה נכנס **בין**
     * האות לניקוד.
     *
     * נמדד ב-Chrome (`scripts/qa/rtl-caret-cluster-probe.mjs`): שלוש היחידות
     * של „שָׁ” מחזירות את אותה תיבה. איטרציה לפי נקודת קוד הייתה יוצרת שלושה
     * חריצים ב-x זהה, ובשורה הזאת — שיש בה אי לטיני, ולכן הבסיס (8) נופל
     * כתפר — **9 ו-10 שרדו**, וההליכה בשני הכיוונים נחתה על 9.
     */
    stubRects();
    const run = NIQQUD_BOXES.map(([left, right], i) => ({ pm: i, left, right }));
    const line = buildLine([{ text: NIQQUD_TEXT, chars: run }], 0, 14, true);
    document.body.append(line);

    const chars = readLineChars(line)!;
    expect(chars, '12 אשכולות ולא 14 יחידות').toHaveLength(12);
    expect(chars[8]).toEqual({ pm: 8, left: 637.5, right: 647.3, ch: 'שָׁ', units: 3 });
    expect(chars[9]!.pm, 'ההיסט הבא מדלג שלוש יחידות').toBe(11);

    const slots = caretSlots(chars, true);
    expect(pmsOf(slots)).toEqual([0, 1, 2, 3, 4, 6, 7, 11, 12, 13, 14]);
    // ‏`expectWalksForward` מוודא מונוטוניות על המסך ואינו רואה חוקיות של
    // אשכול; זו הטענה שחסרה לו, וכל יעד ש-`visualTarget` מחזירה הוא אחד מאלה.
    for (const slot of slots) expect(NIQQUD_LEGAL, `חריץ ${slot.pm}`).toContain(slot.pm);
  });

  it('אות לטינית מפורקת בשורה עברית — אותו כלל, ובלי חריץ בין האות לסימן', () => {
    // ‏`e` ואחריו U+0301, בתוך אי לטיני. כאן החריץ הפסול (6) לא רק שרד —
    // ההליכה נחתה עליו בשני הכיוונים, כי הוא אינו חולק x עם אף חריץ אחר.
    stubRects();
    const run = DECOMPOSED_BOXES.map(([left, right], i) => ({ pm: i, left, right }));
    const line = buildLine([{ text: DECOMPOSED_TEXT, chars: run }], 0, 13, true);
    document.body.append(line);

    const chars = readLineChars(line)!;
    expect(chars).toHaveLength(12);
    expect(chars[5]).toEqual({ pm: 5, left: 655.3, right: 662.4, ch: 'é', units: 2 });

    const slots = caretSlots(chars, true);
    expect(pmsOf(slots)).toEqual([0, 1, 2, 3, 4, 7, 9, 10, 11, 12, 13]);
    for (const slot of slots) expect(DECOMPOSED_LEGAL, `חריץ ${slot.pm}`).toContain(slot.pm);
  });

  it('טווחים שאינם מכסים את השורה ברצף — `null`', () => {
    stubRects();
    const line = buildLine(mixedRuns(), 51, 70, true);
    document.body.append(line);
    expect(readLineChars(line)).toBeNull();
  });
});

interface Setup {
  host: HTMLElement;
  line: HTMLElement;
  superdoc: RtlCaretHost;
  setSelectionTarget: ReturnType<typeof vi.fn>;
  /** כמה פעמים היירוט קרא את התצלום — כלומר ראה את ההקשה. */
  seen: () => number;
}

/**
 * „המודול חי”, ולא רק „לא הפריע”.
 *
 * בדיקת „נמסר למנוע” שמסתפקת ב-`defaultPrevented === false` ובמרגל שלא נקרא
 * עוברת ירוק גם על מודול מת: החלפת `installRtlVisualArrows` ב-no-op השאירה
 * 12 מ-24 בדיקות היירוט ירוקות (נמדד). לכן כל אחת מהן קובעת גם **שהיירוט
 * רץ** — או שקרא את התצלום, או שהקשה נקייה באותה התקנה כן נבלעת.
 */
function expectAlive(setup: Setup, target: HTMLElement): void {
  const event = press(target, 'ArrowRight');
  expect(event.defaultPrevented, 'הקשה נקייה באותה התקנה נבלעת — המודול חי').toBe(true);
  setup.setSelectionTarget.mockClear();
}

function setup({
  runs = mixedRuns(),
  pmStart = 51,
  pmEnd = 64,
  caretOffset = 9,
  caretX = 976,
  caretY = 10,
  rtl = true,
  extraLine,
  nextFragment,
}: {
  runs?: RunSpec[];
  pmStart?: number;
  pmEnd?: number;
  caretOffset?: number;
  caretX?: number;
  caretY?: number;
  rtl?: boolean;
  extraLine?: { runs: RunSpec[]; pmStart: number; pmEnd: number };
  nextFragment?: { runs: RunSpec[]; pmStart: number; pmEnd: number };
} = {}): Setup {
  stubRects();

  const host = document.createElement('div');
  const page = document.createElement('div');
  const fragment = document.createElement('div');
  fragment.setAttribute('data-source-node-id', 'p1');
  fragment.setAttribute('data-pm-start', String(pmStart));
  fragment.setAttribute('data-pm-end', String(extraLine ? extraLine.pmEnd : pmEnd));

  const line = buildLine(runs, pmStart, pmEnd, rtl);
  fragment.append(line);
  if (extraLine) fragment.append(buildLine(extraLine.runs, extraLine.pmStart, extraLine.pmEnd, rtl, 20));
  page.append(fragment);

  if (nextFragment) {
    const second = document.createElement('div');
    second.setAttribute('data-source-node-id', 'p2');
    second.setAttribute('data-pm-start', String(nextFragment.pmStart));
    second.setAttribute('data-pm-end', String(nextFragment.pmEnd));
    second.append(buildLine(nextFragment.runs, nextFragment.pmStart, nextFragment.pmEnd, true));
    page.append(second);
  }

  host.append(page);

  const caret = document.createElement('div');
  caret.className = 'sd-v2-local-selection-caret';
  caret.getBoundingClientRect = () =>
    ({ left: caretX, right: caretX, top: caretY, bottom: caretY }) as DOMRect;
  host.append(caret);
  document.body.append(host);

  const setSelectionTarget = vi.fn();
  const story = { kind: 'story', storyType: 'body' };
  let seen = 0;
  const superdoc: RtlCaretHost = {
    activeEditor: {
      host: {
        readLiveSelectionSyncSnapshot: () => {
          seen += 1;
          return {
            selectionTarget: {
              kind: 'selection',
              start: { kind: 'text', blockId: 'p1', offset: caretOffset, story },
              end: { kind: 'text', blockId: 'p1', offset: caretOffset, story },
              story,
            },
          };
        },
      },
      authoring: { setSelectionTarget },
    },
  };

  return { host, line, superdoc, setSelectionTarget, seen: () => seen };
}

const press = (target: HTMLElement, key: string, init: KeyboardEventInit = {}): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...init });
  target.dispatchEvent(event);
  return event;
};

/** מה שנכתב למנוע בקריאה הראשונה. */
const written = (spy: ReturnType<typeof vi.fn>) =>
  spy.mock.calls[0]![0] as {
    target: { start: { offset: number }; end: { offset: number; blockId: string } };
  };

describe('installRtlVisualArrows', () => {
  it('חץ ימני בשורה עברית: ההקשה נבלעת, והסמן נקבע לשכן החזותי', () => {
    const { host, line, superdoc, setSelectionTarget } = setup();
    const handle = installRtlVisualArrows({ host, superdoc });

    const event = press(line, 'ArrowRight');

    expect(event.defaultPrevented, 'המנוע לא יראה את ההקשה').toBe(true);
    expect(setSelectionTarget).toHaveBeenCalledTimes(1);
    // pm 57 בתוך האי, כלומר היסט 6 בתוך הבלוק.
    expect(written(setSelectionTarget).target.end.offset).toBe(6);
    expect(written(setSelectionTarget).target.start.offset, 'בלי Shift הבחירה מתכווצת').toBe(6);
    handle.dispose();
  });

  it('היעד הוא ההיסט הנוכחי — ההקשה נבלעת ואינה נמסרת למנוע', () => {
    // הציור מפגר אחרי התצלום (פער מתועד): התצלום כבר על היסט 9, והסמן עדיין
    // מצויר ב-x של 10. אז החישוב מחזיר את היסט 9 עצמו — ואילו מסרנו את
    // ההקשה למנוע הוא היה מזיז לוגית, כלומר שמאלה על המסך.
    const state = setup({ caretOffset: 9, caretX: 968.2 });
    const { host, line, superdoc, setSelectionTarget } = state;
    const handle = installRtlVisualArrows({ host, superdoc });

    const event = press(line, 'ArrowRight');

    expect(event.defaultPrevented, 'אין לאן לזוז — אבל ההקשה שלנו').toBe(true);
    expect(setSelectionTarget, 'אין מה לכתוב').not.toHaveBeenCalled();
    handle.dispose();
  });

  it('Shift+חץ: העוגן נשאר, והראש נע לשכן החזותי', () => {
    const { host, line, superdoc, setSelectionTarget } = setup();
    const handle = installRtlVisualArrows({ host, superdoc });

    expect(press(line, 'ArrowRight', { shiftKey: true }).defaultPrevented).toBe(true);

    // הראש לפני העוגן, אבל הכתיבה שומרת את start כעוגן גם עבור המקש הנייטיבי הבא.
    expect(written(setSelectionTarget).target.start.offset).toBe(9);
    expect(written(setSelectionTarget).target.end.offset).toBe(6);
    handle.dispose();
  });

  it('Shift+חץ מעוגן על תפר BiDi נשאר למנוע כשאין לו מיקום יחיד', () => {
    const state = setup({ caretOffset: 5, caretX: 980 });
    const handle = installRtlVisualArrows(state);
    expect(press(state.line, 'ArrowRight', { shiftKey: true }).defaultPrevented).toBe(false);
    expect(state.setSelectionTarget).not.toHaveBeenCalled();
    expectAlive(state, state.line);
    handle.dispose();
  });

  it('Shift+חץ שני ממשיך מהראש שנזכר, גם כשהתצלום כבר אינו מוסר בחירה', () => {
    const { host, line, superdoc, setSelectionTarget } = setup();
    const handle = installRtlVisualArrows({ host, superdoc });
    press(line, 'ArrowRight', { shiftKey: true });
    superdoc.activeEditor!.host!.readLiveSelectionSyncSnapshot = () => ({ selectionTarget: null });

    press(line, 'ArrowRight', { shiftKey: true });

    const second = setSelectionTarget.mock.calls[1]![0];
    expect(second.target.start.offset, 'העוגן לא זז').toBe(9);
    // היסט 6 בתוך האי הלטיני, ושם ימינה הוא קדימה.
    expect(second.target.end.offset, 'הראש התקדם').toBe(7);
    handle.dispose();
  });

  it('חץ רגיל על בחירה שנבנתה כאן קורס לקצה שבכיוון המקש', () => {
    const { host, line, superdoc, setSelectionTarget } = setup();
    const handle = installRtlVisualArrows({ host, superdoc });
    press(line, 'ArrowLeft', { shiftKey: true });
    setSelectionTarget.mockClear();

    // הראש משמאל לעוגן; ימין הוא חזרה אל העוגן, היסט 9.
    press(line, 'ArrowRight');

    expect(written(setSelectionTarget).target.start.offset).toBe(9);
    expect(written(setSelectionTarget).target.end.offset).toBe(9);
    handle.dispose();
  });

  it.each([
    { key: 'ArrowRight', expected: 7 },
    { key: 'ArrowLeft', expected: 6 },
  ])('קריסה בתוך אי לטיני ב-$key נבחרת לפי הצד המצויר', ({ key, expected }) => {
    // היסטים 6 ו-7 יושבים בתוך ABC ומתקדמים משמאל לימין, למרות כיוון הפסקה.
    const state = setup({ caretOffset: 6, caretX: 991.5 });
    const handle = installRtlVisualArrows(state);
    press(state.line, 'ArrowRight', { shiftKey: true });
    state.setSelectionTarget.mockClear();
    expect(press(state.line, key).defaultPrevented).toBe(true);
    expect(written(state.setSelectionTarget).target.start.offset).toBe(expected);
    expect(written(state.setSelectionTarget).target.end.offset).toBe(expected);
    handle.dispose();
  });

  it.each([
    { shiftKey: true, key: 'ArrowRight', expected: 4 },
    { shiftKey: false, key: 'ArrowRight', expected: 7 },
    { shiftKey: false, key: 'ArrowLeft', expected: 6 },
  ])('בחירה זזה עם הפריסה: $key עם Shift=$shiftKey מודד קצוות מחדש', ({ shiftKey, key, expected }) => {
    const state = setup({ caretOffset: 6, caretX: 991.5 });
    const handle = installRtlVisualArrows(state);
    press(state.line, 'ArrowRight', { shiftKey: true });
    state.setSelectionTarget.mockClear();
    // אותו טקסט אחרי הזזה אופקית ואנכית של העמוד, בלי שום שינוי בבחירה.
    for (const span of Array.from(state.line.children) as (HTMLElement & { __chars: PaintedChar[] })[]) {
      span.__chars = span.__chars.map((c) => ({ ...c, left: c.left + 100, right: c.right + 100 }));
    }
    state.line.getBoundingClientRect = () => ({ left: 151, right: 1143.4, top: 40, bottom: 60 }) as DOMRect;
    expect(press(state.line, key, { shiftKey }).defaultPrevented).toBe(true);
    expect(written(state.setSelectionTarget).target.end.offset).toBe(expected);
    expect(written(state.setSelectionTarget).target.start.offset).toBe(shiftKey ? 6 : expected);
    handle.dispose();
  });

  it('פריסה שהחליפה את השורה משחזרת היסט יחיד בלי להשתמש בקואורדינטה ישנה', () => {
    const state = setup({ caretOffset: 6, caretX: 991.5 });
    const handle = installRtlVisualArrows(state);
    press(state.line, 'ArrowRight', { shiftKey: true });
    const moved = mixedRuns().map((run) => ({ ...run,
      chars: run.chars.map((c) => ({ ...c, left: c.left + 100, right: c.right + 100 })),
    }));
    const replacement = buildLine(moved, 51, 64, true, 40);
    state.line.replaceWith(replacement);
    state.setSelectionTarget.mockClear();
    expect(press(replacement, 'ArrowRight', { shiftKey: true }).defaultPrevented).toBe(true);
    expect(written(state.setSelectionTarget).target.start.offset).toBe(6);
    expect(written(state.setSelectionTarget).target.end.offset).toBe(4);
    handle.dispose();
  });

  it('auto-scroll בגבול שורות שומר את השורה של הראש וממשיך תו אחד', () => {
    const state = setup({
      runs: [{ text: 'אחת', chars: boxes(51, [1043, 1035, 1027, 1019], true) }],
      pmStart: 51, pmEnd: 54, caretOffset: 2, caretX: 1027,
      extraLine: { runs: [{ text: 'שנייה', chars: boxes(54, [1043, 1035, 1027, 1019, 1011, 1003], true) }],
        pmStart: 54, pmEnd: 59 },
    });
    const handle = installRtlVisualArrows(state);
    press(state.line, 'ArrowLeft', { shiftKey: true });
    expect(written(state.setSelectionTarget).target.end.offset).toBe(3);
    state.line.getBoundingClientRect = () => ({ top: -100, bottom: -80, left: 51, right: 1043 }) as DOMRect;
    state.line.nextElementSibling!.getBoundingClientRect = () =>
      ({ top: -80, bottom: -60, left: 51, right: 1043 }) as DOMRect;
    state.setSelectionTarget.mockClear();
    expect(press(state.line, 'ArrowLeft', { shiftKey: true }).defaultPrevented).toBe(true);
    expect(written(state.setSelectionTarget).target.start.offset).toBe(2);
    expect(written(state.setSelectionTarget).target.end.offset).toBe(4);
    handle.dispose();
  });

  it('שורה שהוחלפה והפכה את הראש לתפר בין שורות מוסרת את המקש בלי לנחש', () => {
    const state = setup({ caretOffset: 6, caretX: 991.5 });
    const handle = installRtlVisualArrows(state);
    press(state.line, 'ArrowRight', { shiftKey: true });
    const first = buildLine([{ text: 'אבגדהוז', chars: boxes(51, [1000, 990, 980, 970, 960, 950, 940, 930], true) }], 51, 58, true);
    const second = buildLine([{ text: 'חטיכלמ', chars: boxes(58, [1000, 990, 980, 970, 960, 950, 940], true) }], 58, 64, true, 20);
    state.line.replaceWith(first, second);
    state.setSelectionTarget.mockClear();
    expect(press(second, 'ArrowRight', { shiftKey: true }).defaultPrevented).toBe(false);
    expect(state.setSelectionTarget).not.toHaveBeenCalled();
    handle.dispose();
  });

  it('מקש אחר ולחיצת עכבר שוכחים את הבחירה — אבל Shift לבדו לא', () => {
    const { host, line, superdoc, setSelectionTarget } = setup();
    const handle = installRtlVisualArrows({ host, superdoc });

    press(line, 'ArrowLeft', { shiftKey: true });
    press(line, 'Shift', { shiftKey: true });
    press(line, 'ArrowLeft', { shiftKey: true });
    expect(setSelectionTarget.mock.calls[1]![0].target.start.offset, 'העוגן נזכר').toBe(9);

    setSelectionTarget.mockClear();
    press(line, 'End');
    press(line, 'ArrowRight');
    expect(written(setSelectionTarget).target.start.offset, 'סמן מהתצלום ולא קריסה').toBe(6);
    handle.dispose();
  });

  it('בחירת עכבר בעברית נקראת בסיום הגרירה, ואז שתי הקשות ממשיכות מהראש', async () => {
    const { host, line, superdoc, setSelectionTarget } = setup();
    const story = { kind: 'story', storyType: 'body' };
    const point = (offset: number) => ({ kind: 'text', blockId: 'p1', offset, story });
    line.getBoundingClientRect = () =>
      ({ left: 900, right: 1100, top: 0, bottom: 20 }) as DOMRect;
    const editor = superdoc.activeEditor!;
    editor.host!.readLiveSelectionSyncSnapshot = () => ({ selectionTarget: null });
    editor.doc = {
      selection: {
        current: async () => ({
          empty: false,
          target: { segments: [{ blockId: 'p1', range: { start: 4, end: 9 } }] },
          selectionTarget: { kind: 'selection', start: point(4), end: point(9), story },
        }),
      },
    };
    const handle = installRtlVisualArrows({ host, superdoc });

    line.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, clientX: 1000, clientY: 10 }));
    line.dispatchEvent(new MouseEvent('pointerup', { bubbles: true, clientX: 960, clientY: 10 }));
    await new Promise((resolve) => setTimeout(resolve, 0));
    const first = press(line, 'ArrowLeft', { shiftKey: true });
    const queued = press(line, 'ArrowLeft', { shiftKey: true });
    expect(first.defaultPrevented, 'ההכרעה נופלת סינכרונית').toBe(true);
    expect(queued.defaultPrevented).toBe(true);

    await vi.waitFor(() => expect(setSelectionTarget).toHaveBeenCalledTimes(2));
    const last = setSelectionTarget.mock.calls[1]![0];
    expect(last.target.start.offset, 'העוגן הוא תחילת הבחירה').toBe(4);
    expect(last.target.end.offset, 'שתי ההקשות נספרו').toBeGreaterThan(9);
    handle.dispose();
  });

  /** טווח אמיתי בצורת ה-API, עם תיבות עבריות מדודות וקריאה שניתנת לעיכוב. */
  const externalRange = () => {
    const state = setup({
      runs: [{ text: 'שלום עולם פשוט מאוד', chars: CLEAN }],
      pmStart: 1,
      pmEnd: 20,
      caretOffset: 5,
      caretX: 1010.1,
    });
    const story = { kind: 'story', storyType: 'body' };
    const point = (offset: number) => ({ kind: 'text', blockId: 'p1', offset, story });
    const range = {
      empty: false,
      target: { segments: [{ blockId: 'p1', range: { start: 0, end: 5 } }] },
      selectionTarget: { kind: 'selection', start: point(0), end: point(5), story },
    };
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const current = vi.fn(async () => { await gate; return range; });
    const editor = state.superdoc.activeEditor!;
    editor.host!.readLiveSelectionSyncSnapshot = () => ({ selectionTarget: null });
    state.line.getBoundingClientRect = () =>
      ({ left: 900, right: 1100, top: 0, bottom: 20 }) as DOMRect;
    editor.doc = { selection: { current } };
    return { ...state, current, release };
  };

  it('טווח מכיוון לא ידוע נשאר למנוע ב-Shift+חץ בלי להחליף את הראש בעוגן', () => {
    const state = externalRange();
    const handle = installRtlVisualArrows(state);
    // selection.current מחזיר 0..5 גם כאשר Shift+Home יצר עוגן 5 וראש 0.
    for (const key of ['ArrowLeft', 'ArrowRight']) {
      expect(press(state.line, key, { shiftKey: true }).defaultPrevented).toBe(false);
    }
    expect(state.current).not.toHaveBeenCalled();
    expect(state.setSelectionTarget).not.toHaveBeenCalled();
    expect(press(state.line, 'ArrowLeft').defaultPrevented).toBe(false);
    // ההתקנה עצמה חיה; חץ על סמן רגיל עדיין מתוקן.
    state.superdoc.activeEditor!.host!.readLiveSelectionSyncSnapshot = () => ({
      selectionTarget: { end: { kind: 'text', blockId: 'p1', offset: 5 } },
    });
    expectAlive(state, state.line);
    handle.dispose();
    state.release();
  });

  const dragRange = (state: Setup, from = 1043.4, to = 1010.1, modifiers: MouseEventInit = {}) => {
    state.line.dispatchEvent(new MouseEvent('pointerdown', {
      bubbles: true, clientX: from, clientY: 10, ...modifiers,
    }));
    state.line.dispatchEvent(new MouseEvent('pointerup', {
      bubbles: true, clientX: to, clientY: 10, ...modifiers,
    }));
  };

  it('קריאה מושהית אינה בולעת חצים, ואינה מחזירה בחירה שכבר זזה במנוע', async () => {
    const state = externalRange();
    const handle = installRtlVisualArrows(state);
    dragRange(state);
    await vi.waitFor(() => expect(state.current).toHaveBeenCalledTimes(1));
    expect(press(state.line, 'ArrowLeft').defaultPrevented).toBe(false);
    expect(press(state.line, 'ArrowLeft').defaultPrevented).toBe(false);
    state.release();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(state.setSelectionTarget).not.toHaveBeenCalled();
    expect(press(state.line, 'ArrowLeft', { shiftKey: true }).defaultPrevented).toBe(false);
    handle.dispose();
  });

  it.each(['key', 'pointer', 'pointercancel', 'dispose', 'editor'] as const)(
    'קריאה שהתיישנה אינה כותבת או מחזירה מיקוד אחרי %s', async (cancel) => {
      const state = externalRange();
      const handle = installRtlVisualArrows(state);
      dragRange(state);
      await vi.waitFor(() => expect(state.current).toHaveBeenCalledTimes(1));
      if (cancel === 'key') press(state.line, 'Home');
      if (cancel === 'pointer') state.line.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }));
      if (cancel === 'pointercancel') state.line.dispatchEvent(new Event('pointercancel', { bubbles: true }));
      if (cancel === 'dispose') handle.dispose();
      if (cancel === 'editor') state.superdoc.activeEditor = { ...state.superdoc.activeEditor };
      state.release();
      await state.current.mock.results[0]!.value;
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(state.setSelectionTarget).not.toHaveBeenCalled();
      expect(press(state.line, 'ArrowLeft', { shiftKey: true }).defaultPrevented).toBe(false);
      handle.dispose();
    },
  );

  it.each([
    { label: 'Shift+click', from: 1043.4, to: 1010.1, modifiers: { shiftKey: true } },
    { label: 'לחיצה כפולה בלי גרירה', from: 1010.1, to: 1010.1, modifiers: {} },
  ])('$label אינו מנחש עוגן לפי תנועת העכבר', ({ from, to, modifiers }) => {
    const state = externalRange();
    const handle = installRtlVisualArrows(state);
    dragRange(state, from, to, modifiers);
    expect(press(state.line, 'ArrowLeft', { shiftKey: true }).defaultPrevented).toBe(false);
    expect(state.current).not.toHaveBeenCalled();
    expect(state.setSelectionTarget).not.toHaveBeenCalled();
    handle.dispose();
    state.release();
  });

  it('ראש בחירה שאינו מצויר נשאר למנוע, בלי בליעת המקש', async () => {
    const state = externalRange();
    const handle = installRtlVisualArrows(state);
    dragRange(state);
    await vi.waitFor(() => expect(state.current).toHaveBeenCalledTimes(1));
    state.line.remove();
    state.release();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(press(state.host, 'ArrowLeft', { shiftKey: true }).defaultPrevented).toBe(false);
    expect(state.setSelectionTarget).not.toHaveBeenCalled();
    handle.dispose();
  });

  it.each([
    { from: 991.5, to: 1002.2, anchor: 6, head: 6 },
    { from: 1002.2, to: 991.5, anchor: 7, head: 9 },
  ])('גרירה בתוך אי לטיני שומרת את העוגן $anchor לפי המיקום המצויר', async ({ from, to, anchor, head }) => {
    const state = setup();
    const editor = state.superdoc.activeEditor!;
    editor.host!.readLiveSelectionSyncSnapshot = () => ({ selectionTarget: null });
    state.line.getBoundingClientRect = () => ({ left: 900, right: 1100, top: 0, bottom: 20 }) as DOMRect;
    editor.doc = { selection: { current: async () => ({
      empty: false,
      selectionTarget: {
        kind: 'selection',
        start: { kind: 'text', blockId: 'p1', offset: 6 },
        end: { kind: 'text', blockId: 'p1', offset: 7 },
      },
    }) } };
    const handle = installRtlVisualArrows(state);
    dragRange(state, from, to);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(press(state.line, 'ArrowLeft', { shiftKey: true }).defaultPrevented).toBe(true);
    expect(written(state.setSelectionTarget).target.start.offset).toBe(anchor);
    expect(written(state.setSelectionTarget).target.end.offset).toBe(head);
    handle.dispose();
  });

  it('בחירה חיצונית אינה מתחילה קריאה או בולעת את המקש', () => {
    const { host, line, superdoc, setSelectionTarget } = setup();
    const editor = superdoc.activeEditor!;
    editor.host!.readLiveSelectionSyncSnapshot = () => ({ selectionTarget: null });
    const handle = installRtlVisualArrows({ host, superdoc });

    expect(press(line, 'ArrowRight').defaultPrevented).toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    handle.dispose();
  });

  it('חץ באמצע הרכבת IME אינו נגזל — לא בדגל ולא ב-keyCode 229', () => {
    // ‏`swallow()` עוצר גם את שאר המאזינים (`stopImmediatePropagation`), ולכן
    // יירוט כאן היה מוחק מקש שחלונית ההרכבה מחכה לו. התקן של הריפו:
    // `src/ui/shortcuts/match.ts`.
    const state = setup();
    const { host, line, superdoc, setSelectionTarget } = state;
    const handle = installRtlVisualArrows({ host, superdoc });

    expect(press(line, 'ArrowRight', { isComposing: true }).defaultPrevented).toBe(false);
    expect(press(line, 'ArrowRight', { keyCode: 229 }).defaultPrevented).toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    expect(state.seen(), 'היירוט לא טרח אפילו לקרוא את התצלום').toBe(0);
    expectAlive(state, line);
    handle.dispose();
  });

  it('בקצה הימני של שורה גולשת: היעד הוא סוף השורה שמעליה', () => {
    // בלי זה ההקשה הייתה חוזרת למנוע, והוא מזיז שם קדימה לוגית — כלומר
    // שמאלה על המסך. נמדד בשער.
    const { host, superdoc, setSelectionTarget } = setup({
      runs: [{ text: 'שורה', chars: boxes(51, [1043, 1035, 1027, 1019, 1011], true) }],
      pmStart: 51,
      pmEnd: 55,
      caretOffset: 4, // pm 55 — התפר, ששייך לשתי השורות
      caretX: 1043,
      caretY: 30, // השורה השנייה, ולא הראשונה
      extraLine: {
        runs: [{ text: 'שנייה', chars: boxes(55, [1043, 1035, 1027, 1019, 1011, 1003], true) }],
        pmStart: 55,
        pmEnd: 60,
      },
    });
    const handle = installRtlVisualArrows({ host, superdoc });

    // הסמן בתחילת השורה השנייה, בקצה הימני שלה; ימינה = אחורה לוגית.
    press(host, 'ArrowRight');

    expect(setSelectionTarget).toHaveBeenCalledTimes(1);
    // סוף השורה הראשונה הוא pm 54, כלומר היסט 3 — ולא 55, שהוא הסמן עצמו.
    expect(written(setSelectionTarget).target.end.offset).toBe(3);
    handle.dispose();
  });

  it('ימינה בתחילת הפסקה: אין פסקה קודמת, וההקשה נמסרת למנוע', () => {
    const state = setup({
      caretOffset: 0, // pm 51 — תחילת הבלוק, הקצה הימני בעברית
      caretX: 1043.4,
      nextFragment: {
        runs: [{ text: 'אחרי', chars: boxes(65, [900, 892, 884, 876, 868], true) }],
        pmStart: 65,
        pmEnd: 69,
      },
    });
    const { host, line, superdoc, setSelectionTarget } = state;
    const handle = installRtlVisualArrows({ host, superdoc });

    // ימינה בתחילת פסקה = אחורה לוגית → הפסקה הקודמת. אין כזו, ולכן ההקשה
    // חוזרת למנוע; הכיוון ההפוך — שמאלה — נבדק בבדיקה שאחרי זו.
    const event = press(line, 'ArrowRight');
    expect(event.defaultPrevented, 'אין פסקה קודמת').toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    expect(state.seen(), 'היירוט רץ וקרא את התצלום').toBeGreaterThan(0);
    handle.dispose();
  });

  it('שמאלה בסוף הפסקה עובר לתחילת הפסקה הבאה', () => {
    const { host, line, superdoc, setSelectionTarget } = setup({
      caretOffset: 13, // pm 64 — סוף הבלוק, הקצה השמאלי
      caretX: 949.5,
      nextFragment: {
        runs: [{ text: 'אחרי', chars: boxes(65, [900, 892, 884, 876, 868], true) }],
        pmStart: 65,
        pmEnd: 69,
      },
    });
    const handle = installRtlVisualArrows({ host, superdoc });

    press(line, 'ArrowLeft');

    expect(setSelectionTarget).toHaveBeenCalledTimes(1);
    expect(written(setSelectionTarget).target.end.blockId).toBe('p2');
    expect(written(setSelectionTarget).target.end.offset, 'תחילת הפסקה הבאה').toBe(0);
    handle.dispose();
  });

  it('טבלה בין הפסקאות: פער בטווחים, וההקשה נמסרת למנוע ולא מדלגת מעליה', () => {
    // נמדד: „לפני” pm 104..114, הטבלה 115..122 (fragment בלי מזהה מקור, ולכן
    // מחוץ לבורר), „אחרי” 123..133. בלי בדיקת הרציפות החץ קפץ לפסקה שמעבר
    // לטבלה; המנוע עצמו נכנס לתא.
    const state = setup({
      caretOffset: 13, // pm 64 — סוף הבלוק, הקצה השמאלי
      caretX: 949.5,
      nextFragment: {
        runs: [{ text: 'אחרי', chars: boxes(73, [900, 892, 884, 876, 868], true) }],
        pmStart: 73,
        pmEnd: 77,
      },
    });
    const { host, line, superdoc, setSelectionTarget } = state;
    const handle = installRtlVisualArrows({ host, superdoc });

    const event = press(line, 'ArrowLeft');

    expect(event.defaultPrevented, 'המנוע הוא שנכנס לטבלה').toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    expect(state.seen(), 'היירוט רץ וקרא את התצלום').toBeGreaterThan(0);
    handle.dispose();
  });

  it('פסקה עם טאב מטופלת ואינה נופלת חזרה למנוע', () => {
    const { host, line, superdoc, setSelectionTarget } = setup({
      runs: [
        { text: 'קודם', chars: boxes(51, [1043, 1035, 1027, 1019, 1011], true) },
        { tab: true, chars: [{ pm: 55, left: 987, right: 1011 }] },
        { text: 'אחרון', chars: boxes(56, [987, 979, 971, 963, 955, 947], true) },
      ],
      pmStart: 51,
      pmEnd: 61,
      caretOffset: 7, // pm 58, בתוך „אחרון”
      caretX: 971,
    });
    const handle = installRtlVisualArrows({ host, superdoc });

    const event = press(line, 'ArrowRight');

    expect(event.defaultPrevented).toBe(true);
    expect(written(setSelectionTarget).target.end.offset, 'היסט אחד ימינה').toBe(6);
    handle.dispose();
  });

  it('שורה לטינית: ההקשה ממשיכה למנוע כפי שהיא', () => {
    const state = setup({ rtl: false });
    const { host, line, superdoc, setSelectionTarget } = state;
    const handle = installRtlVisualArrows({ host, superdoc });

    const event = press(line, 'ArrowRight');

    expect(event.defaultPrevented).toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    expect(state.seen(), 'היירוט רץ וקרא את התצלום').toBeGreaterThan(0);
    handle.dispose();
  });

  it('טווחים שאינם רציפים — אין מיפוי, ואין יעד', () => {
    const state = setup({ pmEnd: 70 });
    const { host, line, superdoc, setSelectionTarget } = state;
    const handle = installRtlVisualArrows({ host, superdoc });

    const event = press(line, 'ArrowRight');

    expect(event.defaultPrevented).toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    expect(state.seen(), 'היירוט רץ וקרא את התצלום').toBeGreaterThan(0);
    handle.dispose();
  });

  it('Ctrl+חץ הוא קפיצת מילה של המנוע, ואינו נגזל', () => {
    // ומכאן גם הסתירה שנשארת: באותו פריט רשימה `ArrowRight` זז ימינה בעוד
    // `Ctrl+ArrowRight` ממשיך לקפוץ מילה שמאלה. מתועד ב-`docs/engine-gaps.md`.
    const state = setup();
    const { host, line, superdoc, setSelectionTarget } = state;
    const handle = installRtlVisualArrows({ host, superdoc });

    const event = press(line, 'ArrowRight', { ctrlKey: true });

    expect(event.defaultPrevented).toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    expectAlive(state, line);
    handle.dispose();
  });

  it('בלי סמן מצויר אין נקודת מוצא, וההקשה נמסרת למנוע', () => {
    const state = setup();
    const { host, line, superdoc, setSelectionTarget } = state;
    host.querySelector('.sd-v2-local-selection-caret')?.remove();
    const handle = installRtlVisualArrows({ host, superdoc });

    const event = press(line, 'ArrowRight');

    expect(event.defaultPrevented).toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    expect(state.seen(), 'היירוט רץ וקרא את התצלום — ורק אז לא מצא סמן').toBeGreaterThan(0);
    handle.dispose();
  });

  it('אחרי dispose המאזין אינו נוגע עוד', () => {
    const state = setup();
    const { host, line, superdoc, setSelectionTarget } = state;
    const handle = installRtlVisualArrows({ host, superdoc });

    // לפני שההסרה נבדקת — שההתקנה עצמה תוכיח שהיא עשתה משהו.
    expectAlive(state, line);
    handle.dispose();

    const event = press(line, 'ArrowRight');

    expect(event.defaultPrevented).toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
  });
});

/* ------------------------------------------------------------------ */
/* עמודים: קצות המסמך, ופסקה שנחצית                                    */
/* ------------------------------------------------------------------ */

interface FragmentSpec {
  id: string;
  pmStart: number;
  pmEnd: number;
  continuesFromPrev?: boolean;
  continuesOnNext?: boolean;
  lines: Array<{ runs: RunSpec[]; pmStart: number; pmEnd: number; top: number }>;
}

/**
 * גוף עם עמודים (`.superdoc-page`) — כמו שהמנוע מצייר: ה-fragments ילדים
 * ישירים של העמוד, ופסקה חצויה היא שני fragments עם אותו מזהה.
 */
function pagesSetup(pages: FragmentSpec[][], caret: { blockId: string; offset: number; x: number; y: number }) {
  stubRects();
  const host = document.createElement('div');
  for (const fragments of pages) {
    const page = document.createElement('div');
    page.className = 'superdoc-page';
    for (const spec of fragments) {
      const fragment = document.createElement('div');
      fragment.className = 'superdoc-fragment';
      fragment.setAttribute('data-source-node-id', spec.id);
      fragment.setAttribute('data-pm-start', String(spec.pmStart));
      fragment.setAttribute('data-pm-end', String(spec.pmEnd));
      if (spec.continuesFromPrev) fragment.setAttribute('data-continues-from-prev', 'true');
      if (spec.continuesOnNext) fragment.setAttribute('data-continues-on-next', 'true');
      for (const line of spec.lines) fragment.append(buildLine(line.runs, line.pmStart, line.pmEnd, true, line.top));
      page.append(fragment);
    }
    host.append(page);
  }
  const caretEl = document.createElement('div');
  caretEl.className = 'sd-v2-local-selection-caret';
  caretEl.getBoundingClientRect = () =>
    ({ left: caret.x, right: caret.x, top: caret.y, bottom: caret.y }) as DOMRect;
  host.append(caretEl);
  document.body.append(host);

  const setSelectionTarget = vi.fn();
  const story = { kind: 'story', storyType: 'body' };
  const point = { kind: 'text', blockId: caret.blockId, offset: caret.offset, story };
  let seen = 0;
  const superdoc: RtlCaretHost = {
    activeEditor: {
      host: {
        readLiveSelectionSyncSnapshot: () => {
          seen += 1;
          return { selectionTarget: { kind: 'selection', start: point, end: point, story } };
        },
      },
      authoring: { setSelectionTarget },
    },
  };
  const handle = installRtlVisualArrows({ host, superdoc });
  const line = host.querySelector('[dir="rtl"]') as HTMLElement;
  return { host, line, setSelectionTarget, handle, seen: () => seen };
}

/** פסקה של שורה אחת, „סעיף 3 בחוק הזה” — הלוגית שבה נמדדה המלכודת בקצוות. */
const loneDigitFragment = (id: string, top = 0): FragmentSpec => ({
  id,
  pmStart: 1,
  pmEnd: 16,
  lines: [{ runs: [{ text: 'סעיף 3 בחוק הזה', chars: LONE_DIGIT }], pmStart: 1, pmEnd: 16, top }],
});

describe('installRtlVisualArrows — קצות המסמך', () => {
  it('ימינה בתחילת המסמך — נבלע, ואינו נמסר למנוע', () => {
    // נמדד: המנוע זז שם לוגית, כלומר שמאלה, וההקשה הבאה החזירה — 0↔1 בלי סוף.
    const { line, setSelectionTarget, handle } = pagesSetup([[loneDigitFragment('p1')]], {
      blockId: 'p1', offset: 0, x: 1043.3, y: 10,
    });
    const event = press(line, 'ArrowRight');
    expect(event.defaultPrevented).toBe(true);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    handle.dispose();
  });

  it('שמאלה בסוף המסמך — נבלע', () => {
    const { line, setSelectionTarget, handle } = pagesSetup([[loneDigitFragment('p1')]], {
      blockId: 'p1', offset: 15, x: 950.3, y: 10,
    });
    const event = press(line, 'ArrowLeft');
    expect(event.defaultPrevented).toBe(true);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    handle.dispose();
  });

  it('ראשון בעמוד שאינו הראשון — עמוד קודם שלא צויר, וההקשה של המנוע', () => {
    // המנוע מצייר רק עמודים קרובים (נמדד: 3 מתוך 25) — „אין פסקה לפני” אינו
    // תחילת המסמך.
    const { line, setSelectionTarget, handle, seen } = pagesSetup([[], [loneDigitFragment('p9')]], {
      blockId: 'p9', offset: 0, x: 1043.3, y: 10,
    });
    const event = press(line, 'ArrowRight');
    expect(event.defaultPrevented).toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    expect(seen(), 'היירוט רץ וקרא את התצלום').toBeGreaterThan(0);
    handle.dispose();
  });

  it('אחרון בעמוד שאינו האחרון — ההקשה של המנוע', () => {
    const { line, handle, seen } = pagesSetup([[loneDigitFragment('p1')], []], {
      blockId: 'p1', offset: 15, x: 950.3, y: 10,
    });
    expect(press(line, 'ArrowLeft').defaultPrevented).toBe(false);
    expect(seen(), 'היירוט רץ וקרא את התצלום').toBeGreaterThan(0);
    handle.dispose();
  });

  it('פסקה שממשיכה בעמוד הבא אינה סוף המסמך', () => {
    const spec = { ...loneDigitFragment('p1'), continuesOnNext: true };
    const { line, handle, seen } = pagesSetup([[spec]], { blockId: 'p1', offset: 15, x: 950.3, y: 10 });
    expect(press(line, 'ArrowLeft').defaultPrevented).toBe(false);
    expect(seen(), 'היירוט רץ וקרא את התצלום').toBeGreaterThan(0);
    handle.dispose();
  });
});

describe('installRtlVisualArrows — פסקה שנחצית בין עמודים', () => {
  // „שורה” בעמוד אחד (pm 51..55), והמשך „שנייה” בעמוד הבא (pm 55..60).
  const first: FragmentSpec = {
    id: 'p1',
    pmStart: 51,
    pmEnd: 55,
    continuesOnNext: true,
    lines: [{ runs: [{ text: 'שורה', chars: boxes(51, [1043, 1035, 1027, 1019, 1011], true) }], pmStart: 51, pmEnd: 55, top: 0 }],
  };
  const rest: FragmentSpec = {
    id: 'p1',
    pmStart: 55,
    pmEnd: 60,
    continuesFromPrev: true,
    lines: [{ runs: [{ text: 'שנייה', chars: boxes(55, [1043, 1035, 1027, 1019, 1011, 1003], true) }], pmStart: 55, pmEnd: 60, top: 500 }],
  };

  it('בתוך ההמשך: ההיסט יחסי לתחילת הפסקה, ולא ל-fragment', () => {
    // היסט 6 → pm 57, בשורה שבעמוד השני (x=1027). ימינה → pm 56 → היסט 5.
    const { host, setSelectionTarget, handle } = pagesSetup([[first], [rest]], {
      blockId: 'p1', offset: 6, x: 1027, y: 510,
    });
    press(host.querySelectorAll('[dir="rtl"]')[1] as HTMLElement, 'ArrowRight');
    expect(setSelectionTarget).toHaveBeenCalledTimes(1);
    expect(written(setSelectionTarget).target.end).toMatchObject({ blockId: 'p1', offset: 5 });
    handle.dispose();
  });

  it('בגבול, כשהסמן מצויר בעמוד הבא — ה-fragment שלו הוא ההמשך', () => {
    // pm 55 שייך לשניהם. הסמן בתחילת השורה של העמוד השני (x=1043, y=510):
    // שמאלה → pm 56 → היסט 5. בלי ה-y היה נבחר העמוד הראשון, והסמן קפץ
    // לתחילת השורה שם (נמדד: 898 → 811).
    const { host, setSelectionTarget, handle } = pagesSetup([[first], [rest]], {
      blockId: 'p1', offset: 4, x: 1043, y: 510,
    });
    press(host.querySelectorAll('[dir="rtl"]')[1] as HTMLElement, 'ArrowLeft');
    expect(written(setSelectionTarget).target.end).toMatchObject({ blockId: 'p1', offset: 5 });
    handle.dispose();
  });

  it('מסוף העמוד — שמאלה עובר להמשך בעמוד הבא', () => {
    // הסמן בסוף השורה של העמוד הראשון (pm 55, x=1011, y=10). אין חריץ משמאל,
    // אין שורה נוספת ב-fragment — וההמשך בעמוד הבא צמוד (55 = 55). היעד הוא
    // pm 56 → היסט 5 (55 עצמו מוחרג, כמו בין שורות גולשות).
    const { host, setSelectionTarget, handle } = pagesSetup([[first], [rest]], {
      blockId: 'p1', offset: 4, x: 1011, y: 10,
    });
    const event = press(host.querySelector('[dir="rtl"]') as HTMLElement, 'ArrowLeft');
    expect(event.defaultPrevented).toBe(true);
    expect(written(setSelectionTarget).target.end).toMatchObject({ blockId: 'p1', offset: 5 });
    handle.dispose();
  });

  it('מתחילת ההמשך — ימינה חוזר לעמוד הקודם', () => {
    const { host, setSelectionTarget, handle } = pagesSetup([[first], [rest]], {
      blockId: 'p1', offset: 4, x: 1043, y: 510,
    });
    press(host.querySelectorAll('[dir="rtl"]')[1] as HTMLElement, 'ArrowRight');
    expect(written(setSelectionTarget).target.end).toMatchObject({ blockId: 'p1', offset: 3 });
    handle.dispose();
  });

  it('החלק הראשון אינו מצויר — אין מיפוי, וההקשה של המנוע', () => {
    // היסט 2 מתחילת הפסקה. אילו הבסיס היה תחילת ההמשך (55), הוא היה ממופה
    // ל-pm 57 שבתוכו — מיפוי שגוי שנראה תקין.
    const { host, setSelectionTarget, handle, seen } = pagesSetup([[], [rest]], {
      blockId: 'p1', offset: 2, x: 1027, y: 510,
    });
    const event = press(host.querySelector('[dir="rtl"]') as HTMLElement, 'ArrowRight');
    expect(event.defaultPrevented).toBe(false);
    expect(setSelectionTarget).not.toHaveBeenCalled();
    expect(seen(), 'היירוט רץ וקרא את התצלום').toBeGreaterThan(0);
    handle.dispose();
  });
});

describe('installRtlVisualArrows — טבלה שכנה', () => {
  /**
   * כמו שהמנוע מצייר (נמדד): fragment של טבלה בלי מזהה, ובתוכו פסקת התא —
   * עם מזהה, בלי טווח משלה, והטווח על השורה שלה. „לפני” 51..64, הטבלה 65..70,
   * „אחרי” 71..75.
   */
  function tableSetup(caret: { blockId: string; offset: number; x: number }) {
    stubRects();
    const host = document.createElement('div');
    const page = document.createElement('div');
    page.className = 'superdoc-page';

    const before = document.createElement('div');
    before.setAttribute('data-source-node-id', 'p1');
    before.setAttribute('data-pm-start', '51');
    before.setAttribute('data-pm-end', '64');
    before.append(buildLine(mixedRuns(), 51, 64, true, 0));

    const tableEl = document.createElement('div');
    tableEl.setAttribute('data-pm-start', '65');
    tableEl.setAttribute('data-pm-end', '70');
    const row = document.createElement('div');
    const cell = document.createElement('div');
    cell.setAttribute('data-source-node-id', 'c1');
    cell.append(buildLine([{ text: 'בתוכה', chars: boxes(65, [900, 892, 884, 876, 868, 860], true) }], 65, 70, true, 30));
    row.append(cell);
    tableEl.append(row);

    const after = document.createElement('div');
    after.setAttribute('data-source-node-id', 'p2');
    after.setAttribute('data-pm-start', '71');
    after.setAttribute('data-pm-end', '75');
    after.append(buildLine([{ text: 'אחרי', chars: boxes(71, [1043, 1035, 1027, 1019, 1011], true) }], 71, 75, true, 60));

    page.append(before, tableEl, after);
    host.append(page);
    const caretEl = document.createElement('div');
    caretEl.className = 'sd-v2-local-selection-caret';
    const y = caret.blockId === 'p1' ? 10 : 70;
    caretEl.getBoundingClientRect = () => ({ left: caret.x, right: caret.x, top: y, bottom: y }) as DOMRect;
    host.append(caretEl);
    document.body.append(host);

    const setSelectionTarget = vi.fn();
    const story = { kind: 'story', storyType: 'body' };
    const point = { kind: 'text', blockId: caret.blockId, offset: caret.offset, story };
    const superdoc: RtlCaretHost = {
      activeEditor: {
        host: { readLiveSelectionSyncSnapshot: () => ({ selectionTarget: { kind: 'selection', start: point, end: point, story } }) },
        authoring: { setSelectionTarget },
      },
    };
    const handle = installRtlVisualArrows({ host, superdoc });
    return { before, after, setSelectionTarget, handle };
  }

  it('שמאלה מסוף הפסקה שלפני — לתחילת הפסקה שבתא', () => {
    // נמדד: בפסקה לוגית המנוע זז קדימה בתוך הפסקה ולא נכנס לטבלה.
    const { before, setSelectionTarget, handle } = tableSetup({ blockId: 'p1', offset: 13, x: 949.5 });
    const event = press(before.firstElementChild as HTMLElement, 'ArrowLeft');
    expect(event.defaultPrevented).toBe(true);
    expect(written(setSelectionTarget).target.end).toMatchObject({ blockId: 'c1', offset: 0 });
    handle.dispose();
  });

  it('ימינה מתחילת הפסקה שאחרי — לסוף הפסקה שבתא', () => {
    const { after, setSelectionTarget, handle } = tableSetup({ blockId: 'p2', offset: 0, x: 1043 });
    const event = press(after.firstElementChild as HTMLElement, 'ArrowRight');
    expect(event.defaultPrevented).toBe(true);
    // סוף הפסקה שבתא — pm 70, היסט 5 מתחילתה (65).
    expect(written(setSelectionTarget).target.end).toMatchObject({ blockId: 'c1', offset: 5 });
    handle.dispose();
  });
});

/* ------------------------------------------------------------------ */
/* פסקה ריקה בין שתי פסקאות                                            */
/* ------------------------------------------------------------------ */

describe('installRtlVisualArrows — פסקה ריקה', () => {
  /**
   * הצורה נמדדה ב-Chrome על ה-`dist` (superdoc 2.15.0, 17.9.2026): פסקה ריקה
   * היא fragment שטווחו באורך אפס (14..14), ובתוכו `DIV.superdoc-line` באותו
   * טווח, עם `dir="rtl"` ובלי אף נושא היסט. ובאותה מדידה — הפרשי ה-pm:
   * „לפני” 1..13, הריקה 14..14, „אחרי” **16**..28, כלומר 2 אחרי ריקה.
   *
   * כאן הטווחים מוזזים כדי להשתמש בתיבות שנמדדו („סעיף 3 בחוק הזה”), והמבנה
   * — אורך אפס, והפרש 1 ואז 2 — נשמר בדיוק.
   */
  const shifted = (by: number): PaintedChar[] =>
    LONE_DIGIT.map((c) => ({ ...c, pm: c.pm + by }));

  const before: FragmentSpec = {
    id: 'p1',
    pmStart: 1,
    pmEnd: 16,
    lines: [{ runs: [{ text: 'סעיף 3 בחוק הזה', chars: LONE_DIGIT }], pmStart: 1, pmEnd: 16, top: 0 }],
  };
  const empty: FragmentSpec = {
    id: 'p2',
    pmStart: 17,
    pmEnd: 17,
    lines: [{ runs: [], pmStart: 17, pmEnd: 17, top: 30 }],
  };
  const after: FragmentSpec = {
    id: 'p3',
    pmStart: 19,
    pmEnd: 34,
    lines: [{ runs: [{ text: 'סעיף 3 בחוק הזה', chars: shifted(18) }], pmStart: 19, pmEnd: 34, top: 60 }],
  };

  it('ימינה בתחילת הפסקה שאחריה — היעד הוא הריקה, ולא צעד בתוך הפסקה', () => {
    // נמדד בשער בלי התיקון: ההקשה נמסרה למנוע, והוא הזיז מהיסט 0 להיסט 1 —
    // כלומר החץ הימני הזיז את הסמן שמאלה, בדיוק הבאג שהמודול נועד להרוג.
    const { host, setSelectionTarget, handle } = pagesSetup([[before, empty, after]], {
      blockId: 'p3', offset: 0, x: 1043.3, y: 70,
    });
    const event = press(host.querySelectorAll('[dir="rtl"]')[2] as HTMLElement, 'ArrowRight');
    expect(event.defaultPrevented).toBe(true);
    expect(written(setSelectionTarget).target.end).toMatchObject({ blockId: 'p2', offset: 0 });
    handle.dispose();
  });

  it('שמאלה בסוף הפסקה שלפניה — היעד הוא הריקה', () => {
    const { host, setSelectionTarget, handle } = pagesSetup([[before, empty, after]], {
      blockId: 'p1', offset: 15, x: 950.3, y: 10,
    });
    const event = press(host.querySelector('[dir="rtl"]') as HTMLElement, 'ArrowLeft');
    expect(event.defaultPrevented).toBe(true);
    expect(written(setSelectionTarget).target.end).toMatchObject({ blockId: 'p2', offset: 0 });
    handle.dispose();
  });

  it('מתוך הריקה: ימינה חוזר לסוף הקודמת, ושמאלה לתחילת הבאה', () => {
    const first = pagesSetup([[before, empty, after]], {
      blockId: 'p2', offset: 0, x: 1043.3, y: 40,
    });
    const emptyLine = first.host.querySelectorAll('[dir="rtl"]')[1] as HTMLElement;
    expect(press(emptyLine, 'ArrowRight').defaultPrevented).toBe(true);
    expect(written(first.setSelectionTarget).target.end).toMatchObject({ blockId: 'p1', offset: 15 });
    first.handle.dispose();
    document.body.innerHTML = '';

    const second = pagesSetup([[before, empty, after]], {
      blockId: 'p2', offset: 0, x: 1043.3, y: 40,
    });
    const emptyLine2 = second.host.querySelectorAll('[dir="rtl"]')[1] as HTMLElement;
    expect(press(emptyLine2, 'ArrowLeft').defaultPrevented).toBe(true);
    expect(written(second.setSelectionTarget).target.end).toMatchObject({ blockId: 'p3', offset: 0 });
    second.handle.dispose();
  });
});
