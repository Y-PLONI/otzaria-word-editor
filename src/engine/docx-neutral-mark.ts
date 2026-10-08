/**
 * תו ניטרלי בסוף פסקה עברית — בדרך **החוצה**, אחרי שהמנוע כתב את המסמך.
 *
 * ## מה שדווח
 *
 * „הנקודה שבסוף הפסקה מוצגת בתחילת השורה”. המשתמש כתב משפט עברי בעורך, שמר,
 * ופתח ב-Word — והנקודה הסוגרת מצוירת צמודה לשוליים הימניים, כלומר לפני המילה
 * הראשונה.
 *
 * ## השורש
 *
 * תו **ניטרלי** — נקודה, פסיק, סוגר — אינו נושא כיוון משלו. באלגוריתם הדו-כיווני
 * הוא נפתר מהשכנים ומרמת ההטמעה, ובריצה שהמנוע כתב בלי `<w:rtl/>` ‏Word פותר
 * אותו כשמאל-לימין ומציב אותו בקצה ההתחלה של הפסקה. ‏`<w:bidi/>` ברמת הפסקה
 * אינו מכסה את זה: הוא קובע את כיוון הפסקה, לא את רמת ההטמעה של הריצה.
 *
 * דווח למעלה: <https://github.com/superdoc/docx-editor/issues/4011>.
 *
 * ## למה **לא** `<w:rtl/>`, אף שהוא מתקן
 *
 * גרסה קודמת של הקובץ הזה סימנה ריצות עבריות ב-`<w:rtl/>`. זה עובד, וזה נמדד
 * שעובד — אבל **הוא מחליף את כל מודל פתירת העיצוב של Word**, וזה נמדד גם כן.
 * ‏A/B ב-Word 365 (COM נסתר → SaveAs PDF → קריאת הגופן והגודל מהזרם), פסקה
 * עברית שכל ההבדל בין גרסאותיה הוא מקור העיצוב ונוכחות הדגל:
 *
 * | מקור העיצוב | `w:rtl` | מה Word צייר |
 * |---|---|---|
 * | ישיר: `w:b`, `sz 36`, `rFonts ascii="David"` | ✗ | David-Bold 18, מודגש |
 * | אותו דבר | ✓ | **Arial 12, לא מודגש** |
 * | מסגנון פסקה, אותן תכונות | ✗ | David-Bold 18, מודגש |
 * | מאותו סגנון | ✓ | **Arial 12, לא מודגש** |
 * | ישיר + מראה מלאה (`bCs`, `szCs`, `rFonts cs`) | ✓ | David-Bold 18, מודגש |
 *
 * כלומר Word אינו בוחר את מחסנית הכתב המורכב לפי הכתב של התווים; **הדגל עצמו**
 * מעביר אותו לשם. ומכאן שסימון `w:rtl` מחייב למלא את המחסנית המקבילה
 * (`bCs`, `iCs`, `szCs`, `rFonts@cs`, וכל מה שדומה להם) **בכל תכונה ובכל רמה**
 * של היררכיית הסגנונות — עיצוב ישיר, סגנון תו, סגנון פסקה, סגנון שיורש סגנון,
 * ‏`docDefaults`, סגנון טבלה, עיצוב מותנה, גופני ערכת נושא. השורה הרביעית
 * בטבלה היא ההוכחה שזה אינו ניתן לסגירה ברמת הריצה: לעיצוב שמגיע מסגנון אין
 * `rPr` על הריצה שאפשר למרות.
 *
 * זה שטח שאינו נגמר, ותיקון **חלקי** שלו גרוע מכלום: הוא מחליף באג שרואים
 * בנקודה אחת בבאג שמוחק הדגשה, גודל וגופן מכל כותרת עברית במסמך.
 *
 * **עדכון (7.10.2026):** ‏`w:rtl` חזר, ב-`docx-run-direction.ts`, אחרי
 * שהשטח נסגר במדידה ולא בהנחה: מה שיורש מהסגנונות **אמור** לעבור למחסנית
 * המורכבת (זו ההגדרה העברית של המסמך, וזה מה שדווח שחסר), רשת ביטחון מכסה
 * שרשרת שאינה מצהירה אותה כלל, ו-`docx-cs-align.ts` מיישר את המסמכים שבהם
 * הצד המורכב הגיע שבור מתבנית המנוע. ה-RLM נשאר: הוא אינו נוגע בעיצוב, והוא
 * מכסה ריצה שהסימון מדלג עליה (הצהרה כבויה, ריצה מורכבת שאינה מפוצלת).
 *
 * ## מה שנעשה במקום, ולמה הוא חסום
 *
 * `RLM` ‏(U+200F, RIGHT-TO-LEFT MARK) הוא תו חזק ימין-לשמאל חסר רוחב. הוא פותר
 * את התו הניטרלי **ברמת הטקסט**, דרך אותו אלגוריתם דו-כיווני, ו**אינו נוגע
 * באף תכונת עיצוב** — ולכן אין לו את השטח הבלתי-חסום שלמעלה. נמדד באותה דרך,
 * כשטווח השורה הוא 370 (הקצה השמאלי, סוף שורה עברית) עד 524 (הימני, תחילתה):
 *
 * | הפסקה | התיקון | x של הנקודה |
 * |---|---|---|
 * | `<w:bidi/>` על הפסקה | — | 518.7 — בתחילת השורה |
 * | `<w:bidi/>` על הפסקה | RLM אחרי הנקודה | **369.9 — תקין** |
 * | bidi **יורש מ-`docDefaults`** | — | 518.7 |
 * | bidi יורש | `w:rtl` | 370.0 |
 * | bidi יורש | RLM אחרי הנקודה | **369.9 — תקין** |
 * | נקודה **בתוך** ריצת הטקסט, bidi יורש | — | 521.1 |
 * | נקודה בתוך ריצת הטקסט, bidi יורש | RLM | **372.3 — תקין** |
 * | — | RLM **לפני** הנקודה | 521.1 — שבור |
 *
 * שלוש מסקנות: הוא עובד גם כשכיוון הפסקה **יורש** — וזו הצורה שבה נולד מסמך
 * חדש של התוסף (`blank-document.ts` כותב `<w:bidi/>` ב-`pPrDefault`) וגם זו של
 * קובץ משתמש אמיתי (51 מתוך 53 פסקאות ב-`tmp/sp3946/case.docx`), כלומר בדיוק
 * המקום שבו הגישה הקודמת החמיצה; הוא עובד בשתי הצורות של הבאג; והוא חייב לבוא
 * **אחרי** התו הניטרלי.
 *
 * ## איפה בדיוק הוא נכנס: לפני הרווח הסופי
 *
 * ‏`<w:t>abc. </w:t>` בלי `xml:space="preserve"` — Word **מוחק** את הרווח הסופי.
 * הצמדת ה-RLM בסוף הייתה הופכת אותו לפנימי, כלומר מחיה רווח שהיה נמחק. נמדד,
 * פעמיים על אותו קובץ — `Range.Text` של Word (אילו תווים שרדו) וה-x מה-PDF:
 *
 * | מה נכתב | הטקסט ש-Word רואה | x של הנקודה |
 * |---|---|---|
 * | `בנקודה. ` (בסיס) | `בנקודה.` — הרווח נמחק | 521.1 — שבור |
 * | `בנקודה.<RLM> ` | `בנקודה.<RLM>` — הרווח נמחק | **372.3 — תקין** |
 * | `בנקודה. <RLM>` ‏+ `preserve` | `בנקודה.·<RLM>` — **הרווח שרד** | 369.9 |
 * | `בנקודה. <RLM>` בלי `preserve` | `בנקודה.·<RLM>` — **הרווח שרד** | 369.9 |
 *
 * שתי הצורות האחרונות מתקנות את הנקודה ו**משנות את הטקסט**. לכן ה-RLM נכנס
 * אחרי התו הניטרלי האחרון ולפני הרווח הסופי: הרווח נשאר סופי ונמחק כמו קודם,
 * והטקסט הנראה יוצא זהה בדיוק.
 *
 * ## הכלל, ולמה הוא צר בכוונה
 *
 * פסקה שהטקסט שלה — בלי רווחים סופיים — נגמר בתו ניטרלי, והתו החזק האחרון
 * שלפניו הוא אות ימנית, מקבלת RLM אחרי הניטרלי האחרון. זהו.
 *
 * מה שנשאר **בחוץ בכוונה**: ספרה או אות לטינית אחרי האות העברית האחרונה
 * („שלום 12.”, „שלום abc.”). שם הפתרון הנכון תלוי במה שביניהם, וכל הרחבה
 * חייבת מדידה משלה. הרחבה בלי מדידה היא בדיוק מה שהפיל את הגרסה הקודמת.
 *
 * ## הערובה: הפלט הוא הקלט ועוד תווים חסרי-רוחב, או שאין פלט
 *
 * כל שינוי כאן הוא **הכנסה** של תו אחד בתוך תוכן של `<w:t>`; שום בייט של הקלט
 * אינו נמחק ואינו מוזז. זה נאכף ולא מונח: ההיסטים חייבים להיות ממוינים ובתוך
 * גבולות המחרוזת, והפלט נבנה מחיתוכים רציפים שאורכם הכולל הוא אורך הקלט. כל
 * הפרה מחזירה `null`, כלומר הבייטים המקוריים יוצאים כמות שהם.
 *
 * זה מה שהופך כל באג עתידי כאן לאי-פעולה במקום להשחתה — וזה בדיוק מה שחסר
 * לגרסה הקודמת, שמצב הכשל שלה לא היה זריקה אלא פלט שגוי בשקט.
 *
 * ## למה על הבייטים ולא במודל
 *
 * `insert` דרך ה-API עובד, אבל הוא **מוטציה במסמך**: צעד היסטוריה, דגל „לא
 * נשמר”, וציור מחדש, בכל שמירה ושמירה. תיקון על הבייטים שיוצאים אינו נוגע במה
 * שהמשתמש רואה ואינו מזיז לו את ה-undo. זו גם אותה נקודה שבה `docx-preflight.ts`
 * מתקן את הכיוון הנכנס, ואותו קורא zip משרת את שניהם.
 */
import { skippedCloser, TOKEN_SOURCE, applyXmlInserts, wordPrefix, type XmlInsert } from './docx-parts';

/** ‏U+200F. חסר רוחב, ולכן אינו נראה לא בעורך ולא ב-Word. */
const RLM = '‏';

/**
 * מערכות כתב שנקראות מימין לשמאל. אות שלהן היא R או AL באלגוריתם הדו-כיווני.
 *
 * זו אותה רשימה שב-`rtl-caret.ts`, ומאותו טעם שנמדד שם: המבחן הוא **תכונת
 * התו** ולא טווח קודים. טווח בולע את מה שיושב בתוכו ואינו אות — ניקוד (NSM),
 * ספרות ערביות-הודיות (AN), פסיק ערבי (CS) ו-`U+FEFF` (BN) — וכל אחד מהם
 * היה משנה את ההכרעה כאן.
 */
const RTL_SCRIPT =
  /[\p{Script=Hebrew}\p{Script=Arabic}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Adlam}]/u;

/** אות — כולל סימני ניקוד שמצטרפים לאות (Other_Alphabetic). */
const LETTER = /\p{Alphabetic}/u;

/** ספרה. אינה חזקה, אבל גם אינה ניטרלית — ראו „הכלל” בהערת הפתיחה. */
const DIGIT = /\p{Nd}/u;

/** סימן מצטרף: ניקוד, טעם, ותג. לוקח את כיוון התו שלפניו, ולכן מדולג. */
const MARK = /[\p{Mn}\p{Me}]/u;

/**
 * הסיווג שהכלל כאן צריך, ולא יותר ממנו.
 *
 * ‏`R` אינו „אות עברית” אלא **כל תו ממערכת כתב ימנית**, וזו אינה קפדנות: מקף
 * (U+05BE), פסק (U+05C0), סוף פסוק (U+05C3), גרש (U+05F3) וגרשיים (U+05F4)
 * הם סימני פיסוק עבריים שמחלקת ה-bidi שלהם היא **R** — חזקים ימניים, ולא
 * ניטרליים. פסקה שנגמרת ב„ר' יוסי ז׳” כבר נפתרת נכון, ו-RLM אחריה היה בייט
 * מיותר בתוך טקסט תורני. מבחן של `\p{Alphabetic}` לבדו היה מסווג אותם
 * כניטרליים ומסמן.
 */
export type CharClass = 'R' | 'L' | 'D' | 'M' | 'N';

export function classOf(ch: string): CharClass {
  if (MARK.test(ch)) return 'M';
  if (DIGIT.test(ch)) return 'D';
  if (ch === '\u200e') return 'L';
  if (ch === RLM || ch === '\u061c') return 'R';
  if (RTL_SCRIPT.test(ch)) return 'R';
  if (LETTER.test(ch)) return 'L';
  return 'N';
}

/** תו אחד שנקרא מ-`<w:t>`, וההיסט ב-XML המקורי מיד אחרי הייצוג שלו. */
interface DecodedChar {
  ch: string;
  end: number;
  /**
   * הטווח של ה-`<w:t>` שהתו נקרא ממנה. נישא על התו ואינו מחושב מחדש: חיפוש
   * הטווח בזמן ההחלה היה סריקה על כל הטווחים לכל הכנסה, כלומר ריבועי במספר
   * הפסקאות — נמדד 2,229ms על 40,000 פסקאות מול 387ms בלעדיו.
   */
  span: { from: number; to: number };
}

const NAMED_ENTITIES: Readonly<Record<string, string>> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

const ENTITY = /&(?:#(\d+)|#x([0-9a-fA-F]+)|(amp|lt|gt|quot|apos));/y;

/**
 * הטקסט של `<w:t>` תו-תו, עם ההיסט של כל תו ב-XML.
 *
 * ההיסטים נדרשים כדי שההכנסה תיפול **בין** תווים ולא באמצע ישות: ‏`&amp;` הוא
 * תו אחד שתופס חמישה בייטים, ותו על-בסיסי הוא תו אחד ששני יחידות UTF-16.
 * שניהם מטופלים כאן, ולכן `[...text]` במקום לולאת אינדקסים.
 */
function readTextTail(paragraph: ParagraphRecord, raw: string, base: number, to: number): void {
  const span = { from: base, to };
  let at = raw.length;
  let foundTail = false;
  // Only the final non-space character and the last directional character are
  // needed. Read each text node backwards; do not allocate one object per
  // character or spread an arbitrarily long paragraph into Array.push().
  while (at > 0) {
    const end = base + at;
    let start = at - 1;
    let ch = raw[start]!;
    if (ch === ';') {
      let amp = start - 1;
      while (amp >= 0 && /[a-zA-Z0-9#]/.test(raw[amp]!)) amp -= 1;
      ENTITY.lastIndex = amp;
      const match = amp >= 0 && raw[amp] === '&' ? ENTITY.exec(raw) : null;
      if (match && ENTITY.lastIndex === at) {
        const [whole, dec, hex, named] = match;
        if (named) {
          ch = NAMED_ENTITIES[named] ?? whole;
        } else {
          const code = dec ? Number.parseInt(dec, 10) : Number.parseInt(hex!, 16);
          // סרוגייט בודד אינו תו, ותו שמעבר לטווח אינו קיים. ישות כזאת נשארת
          // כמות שהיא, ונספרת כתו אחד — די בכך שההיסט יהיה נכון.
          ch = Number.isFinite(code) && code <= 0x10ffff && (code < 0xd800 || code > 0xdfff)
            ? String.fromCodePoint(code)
            : whole;
        }
        start = amp;
      }
    } else if (start > 0 && /[\uDC00-\uDFFF]/.test(ch) && /[\uD800-\uDBFF]/.test(raw[start - 1]!)) {
      start -= 1;
      ch = raw.slice(start, at);
    }
    at = start;
    if (!foundTail && !/^\s$/u.test(ch)) {
      paragraph.tail = { ch, end, span };
      foundTail = true;
    }
    const kind = classOf(ch);
    if (kind !== 'N' && kind !== 'M') {
      paragraph.lastDirectional = kind;
      return;
    }
  }
}

/** פסקה אחת: התווים שנאספו מ-`<w:t>` שלה, בסדר. */
interface ParagraphRecord {
  tail: DecodedChar | null;
  lastDirectional: CharClass | null;
}

/**
 * ההיסט שבו פסקה אחת צריכה RLM, או `null` כשאינה צריכה.
 *
 * הכלל בהערת הפתיחה, כאן בקוד: מורידים רווחים סופיים (Word מוחק אותם ממילא),
 * דורשים שמה שנשאר ייגמר בתו ניטרלי, ושהתו המכריע האחרון לפניו יהיה תו ממערכת
 * כתב ימנית.
 */
function insertionFor(paragraph: ParagraphRecord): XmlInsert | null {
  const tail = paragraph.tail;
  if (!tail) return null;
  // כבר סומן. ‏RLM הוא עצמו חזק ימני, ולכן בלי הבדיקה הזאת פסקה שיש לה רווח
  // סופי הייתה מקבלת RLM נוסף בכל שמירה.
  if (tail.ch === RLM) return null;
  // חייב להיגמר בניטרלי. סימן מצטרף בסוף יוצא מהתחום הצר: הוא לוקח את כיוון
  // שכנו, ופסקה שנגמרת בו אינה המקרה שדווח.
  if (classOf(tail.ch) !== 'N') return null;

  return paragraph.lastDirectional === 'R' ? { at: tail.end, text: RLM, span: tail.span } : null;
}

/**
 * מוסיפה RLM אחרי תו ניטרלי שסוגר פסקה עברית. ‏`null` = אין מה לשנות, או שיש
 * סיבה לא לגעת.
 */
export function markNeutralParagraphEnds(xml: string): string | null {
  // חיפוש אחד לפני הסריקה: חלק שאין בו אף אות ימנית אינו נוגע לזה בכלל.
  if (!RTL_SCRIPT.test(xml) && !/&#(?:x[0-9a-fA-F]+|\d+);/.test(xml)) return null;
  const prefix = wordPrefix(xml);
  if (prefix === null) return null;

  /** ההכנסה, והטווח של ה-`<w:t>` שהיא נופלת בתוכה. ראו `applyInserts`. */
  const inserts: XmlInsert[] = [];
  /** מחסנית הפסקאות — פסקה בתוך תיבת טקסט היא פסקה בתוך ריצה. */
  const paragraphs: ParagraphRecord[] = [];
  /**
   * עומק הריצות. מחוץ לריצה אין `<w:t>` שמעניין אותנו — וזה גם מה שמייתר
   * מעקב אחרי `pPr`: ל-`<w:pPr>` אין בן `<w:r>` בסכמה, ולכן ה-`<w:rPr>` של
   * סימן הפסקה שיושבת בתוכה אינה יכולה להגיע לכאן ממילא.
   */
  let runDepth = 0;
  /** מיקום פתיחת `<w:t>` שנפתחה ועדיין לא נסגרה. */
  let textFrom: number | null = null;

  const token = new RegExp(TOKEN_SOURCE.source, 'g');
  for (let match = token.exec(xml); match; match = token.exec(xml)) {
    const closer = skippedCloser(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      // הערה שאינה נסגרת: אין יותר תגים שאפשר לסמוך עליהם. מה שנאסף עד כאן
      // שייך לפסקאות שכבר נסגרו, ולכן תקף.
      if (end < 0) break;
      token.lastIndex = end + closer.length;
      continue;
    }

    const closing = match[1]!;
    const tagPrefix = match[2]!;
    const name = match[3]!;
    const attributes = match[4]!;
    if (tagPrefix !== prefix) continue;
    const selfClosing = attributes.endsWith('/');
    if (selfClosing) continue;

    if (name === 'p') {
      if (closing) {
        const done = paragraphs.pop();
        const at = done ? insertionFor(done) : null;
        if (at !== null) inserts.push(at);
      } else {
        paragraphs.push({ tail: null, lastDirectional: null });
      }
      continue;
    }

    if (name === 'r') {
      if (closing) {
        runDepth = Math.max(0, runDepth - 1);
        // ‏`<w:t>` שלא נסגרה עד סוף הריצה אינה טקסט שאפשר לסמוך עליו.
        textFrom = null;
      } else {
        runDepth += 1;
      }
      continue;
    }

    if (name === 't' && runDepth > 0) {
      const paragraph = paragraphs[paragraphs.length - 1];
      if (closing) {
        if (textFrom !== null && paragraph) {
          readTextTail(paragraph, xml.slice(textFrom, match.index), textFrom, match.index);
        }
        textFrom = null;
      } else {
        textFrom = token.lastIndex;
      }
    }
  }

  if (inserts.length === 0) return null;
  return applyXmlInserts(xml, inserts);
}
