/**
 * מראת הכתב המורכב — בדרך **החוצה**, על ריצות שכבר מצהירות `<w:rtl/>`.
 *
 * ## מה שדווח
 *
 * „הדגשתי כותרת עברית, וב-Word היא לא מודגשת.” הקובץ הגיע מ-Word, ההדגשה
 * נראית בעורך, ואחרי שמירה היא נעלמת — ואיתה, אם שונו, גם הגודל והגופן.
 *
 * ## השורש
 *
 * לריצה ב-OOXML יש **שתי** מחסניות עיצוב מקבילות: הצד הלטיני (`w:b`, `w:i`,
 * `w:sz`, `w:rFonts@ascii`) והצד של הכתב המורכב (`w:bCs`, `w:iCs`, `w:szCs`,
 * `w:rFonts@cs`). ‏Word בוחר ביניהן לפי ההצהרה `w:rtl` (ECMA-376 §17.3.2.30)
 * או `w:cs` (§17.3.2.7) — לא לפי הכתב של התווים. נמדד, פסקה עברית שכל ההבדל
 * בין גרסאותיה הוא הדגל:
 *
 * | ה-`rPr` | מה Word צייר |
 * |---|---|
 * | `rFonts ascii="David"`, `b`, `sz 36` — **בלי** `w:rtl` | David-Bold 18pt, מודגש |
 * | אותו דבר **+ `w:rtl`** | **Arial 12pt, לא מודגש** |
 * | אותו דבר + `w:rtl` + `bCs`, `szCs`, `rFonts cs` | David-Bold 18pt, מודגש |
 *
 * שתי השורות הראשונות נבדלות בדגל בלבד, והשלישית מחזירה הכול. וה-12pt Arial
 * אינו נפילה שרירותית: ה-`docDefaults` של אותו מסמך נושא `szCs w:val="24"`
 * ו-`cstheme="minorBidi"`, כלומר Word פתר את המחסנית המורכבת ומצא אותה ריקה.
 *
 * ‏**וריצה שמצהירה `w:rtl` היא המקרה הרגיל ולא קצה:** ‏Word העברי מפצל טקסט
 * לפי כתב ומסמן כל ריצה עברית. נמדד על משפט שהוקלד ב-Word עצמו — שש ריצות,
 * ו-`<w:rtl/>` על כולן פרט למילה הלטינית, כולל על שני הרווחים ועל הנקודה.
 *
 * ## ומאז: גם על ריצות שזה עתה סומנו
 *
 * ‏`docx-run-direction.ts` מסמן היום כל ריצה עברית, ורץ **לפני** המודול הזה
 * (`docx-postflight.ts`). לכן הדגשה שהמנוע כתב בצד הלטיני על ריצה שנכתבה
 * בתוסף מקבלת את התאום כאן, בדיוק כמו על ריצה ש-Word סימן. מה שנכתב למטה על
 * „הוספת `w:rtl` נפסלה” מתאר את ההחלטה של אז; היא נסגרה שם.
 *
 * ## למה זה חסום, בשונה מסימון `w:rtl` שנפסל
 *
 * זו ההבחנה שכל המודול הזה תלוי בה. `docx-neutral-mark.ts` מתעד למה **הוספת**
 * `w:rtl` נפסלה: היא מעבירה את Word למחסנית שאיש אינו ממלא, ולכן היא מחייבת
 * מראה מלאה בכל תכונה **ובכל רמה של היררכיית הסגנונות** — כולל `styles.xml`,
 * שאין בו `<w:r>` שאפשר למרות. שטח שאינו נגמר.
 *
 * כאן ההצהרה **כבר בקובץ**. בחירת המחסנית נעשתה בידי מי שכתב אותו, ואנחנו
 * איננו מעבירים דבר — רק ממלאים את המחסנית ש-Word ממילא קורא ומוצא ריקה.
 * ומכאן שלושה דברים:
 *
 * 1. **אין צורך בהיררכיית הסגנונות.** ערך שמגיע מסגנון ואינו ממורה שם כבר
 *    נופל היום, לפני הקובץ הזה ובלי קשר אליו. המראה ברמת הריצה אינה מתקנת
 *    אותו ואינה מחמירה אותו.
 * 2. **התכונות סופיות:** ארבע.
 * 3. **הפעולה מוסיפה בלבד, והתאום נושא את ערך השותף.** היא כותבת רק לאיבר
 *    שאינו קיים, רק כשהתאום הלטיני שלו כן קיים, ועם אותו `w:val` בדיוק.
 *    הגרוע שיכול לקרות הוא שמה שהמסמך אומר במפורש ברמת הריצה **ייראה** — גם
 *    „מודגש” וגם „לא מודגש”.
 *
 * ## הדגל נושא ערך, ולא רק נוכחות
 *
 * ‏`w:b` ו-`w:i` הם `ST_OnOff`: נוכחות בלי `w:val` היא „דלוק”, ו-`w:val="0"`
 * (או `false`, או `off`) היא „כבוי” — אמירה מפורשת שמבטלת עיצוב שיורש מסגנון.
 * לכן התאום נכתב עם **אותו `w:val` של השותף**, בדיוק כפי ש-`szCs` נושא את ערך
 * ה-`sz`: ‏`<w:b/>` מוליד `<w:bCs/>`, ו-`<w:b w:val="0"/>` מוליד
 * `<w:bCs w:val="0"/>`.
 *
 * תאום ש**מתעלם** מהערך אינו „פחות תיקון” אלא היפוך: הוא מדליק את המחסנית
 * שהמסמך כיבה. נמדד על 132 המסמכים שעל המכונה: 24 ריצות בשלושה מהם מצהירות
 * `w:rtl` דלוק, נושאות `<w:b w:val="0"/>` ואין בהן `bCs` — כלומר נופלות בדיוק
 * כאן. ובאותו קורפוס, מכל 115 ה-`rPr` שיש בהן `b` מכובה לצד `bCs`, ב-**115**
 * גם ה-`bCs` מכובה ובאף אחת אינה דלוקה: העתקת הערך היא מה ש-Word עצמו כותב.
 *
 * ‏`sz` ו-`rFonts` מעולם לא סבלו מזה, והם אינם צריכים תיקון: הם נושאים ערך ולא
 * דגל, והקוד העתיק אותו מתחילתו. ל-`ST_HpsMeasure` אין „כבוי” שאפשר לאבד,
 * ו-`<w:sz/>` בלי `w:val` אינה נגעת בכלל, כי אין מה להעתיק.
 *
 * ## מה שנשאר בחוץ, ולמה
 *
 * ריצה ש-`w:rtl` שלה מוצהר **מכובה** (`w:val="0"`) אינה נגעת: מי שכתב את
 * הקובץ אמר „זו אינה עברית”, ו-Word קורא שם את הצד הלטיני ממילא.
 *
 * ‏**אבל לא ריצה ש-`w:b` שלה מכובה.** ההבחנה אינה „דגל מכובה” אלא איזו מחסנית
 * ‏Word קורא: ‏`w:rtl` מכובה מפנה אותו לצד הלטיני, ולכן אין שם מה למלא; ‏`w:b`
 * מכובה משאיר אותו קורא את המחסנית המורכבת, והשארתה ריקה פירושה שההדגשה
 * תיפתר מהסגנון — כלומר דווקא אי-כתיבה היא שמתעלמת מהאמירה המפורשת. וזה עולה
 * בקנה אחד עם מה שהמודול האח כבר קבע בדרך פנימה: ‏`docx-preflight.ts` אינו
 * הופך `<w:b w:val="0"/>` ל„מודגש”, מפני שדגל מכובה הוא אמירה של מי שכתב את
 * הקובץ. אותה אמירה עצמה, בדרך החוצה, היא שנכתבת כאן לתאום.
 *
 * ‏`rPrChange` — ההיסטוריה של שינוי מסומן — אינה העיצוב החי, ואינה נגעת.
 *
 * וכל שאר התכונות שיש להן תאום מורכב (`w:lang w:bidi`, `w:em`) אינן כאן, כי
 * לא נמדדו. הרחבה בלי מדידה היא בדיוק מה שהפיל את הגישה הקודמת.
 *
 * ## סדר הסכמה, ומה שהוא **אינו** מבטיח כאן
 *
 * ‏`CT_RPr` (ECMA-376 §17.3.2.28) קובע את הרצף `rFonts, b, bCs, i, iCs, …,
 * sz, szCs, …`. לכן כל תאום נכתב **מיד אחרי** הלטיני שלו, ו-`w:cs` נכנס
 * כמאפיין על `w:rFonts` הקיים — כלומר הסדר ה**יחסי** בין תאום לשותפו נכון.
 *
 * אבל אין בזה הבטחה שה-`rPr` כולה תקינה, מפני שהיא אינה תקינה כשהיא מגיעה.
 * נמדד על 2.15.0, דרך מסלול השמירה: הדגשה על ריצה מצהירה יצאה
 * `<w:b/><w:rFonts …/><w:rtl/>` — ‏`b` **לפני** `rFonts` — ואחריה נטייה יצאה
 * `<w:i/><w:b/>…`. שני הסדרים מפרים את הרצף. המראה אינה מתקנת את זה ואינה
 * מחמירה אותו: היא מצמידה כל תאום לשותפו, ולכן מוסיפה בדיוק במקום שבו
 * הסכמה הייתה מצפה לו **ביחס לשותף**.
 *
 * ## הערובה
 *
 * כל שינוי הוא **הכנסה** בתוך ה-`rPr` החיה של ריצה, ולא מחיקה ולא הזזה. זה
 * נאכף ב-`applyXmlInserts` (`docx-parts.ts`), שדורש מכל היסט ליפול בתוך
 * הטווח שההכנסה הצהירה עליו. הפרה מחזירה `null`, והבייטים המקוריים יוצאים
 * כמות שהם.
 */
import {
  skippedCloser,
  TOKEN_SOURCE,
  applyXmlInserts,
  attribute,
  isOn,
  quoteAttribute,
  valueOf,
  wordPrefix,
  type XmlInsert,
} from './docx-parts';

/** מה שנאסף על ה-`rPr` החיה של ריצה אחת. מיקומים הם היסטים ב-XML המקורי. */
interface RunProps {
  /** הטווח של ה-`rPr`, לאכיפת הערובה. */
  span: { from: number; to: number };
  /** מה ש-`w:rtl` מצהירה, או `null` כשאינה שם. */
  declaredRtl: boolean | null;
  /** יש בה `rPrChange` — היסטוריה של שינוי מסומן. */
  hasChange: boolean;
  /** סוף התג של `b`, כדי ש-`bCs` ייכתב מיד אחריו לפי סדר הסכמה. */
  boldEnd: number | null;
  /** ה-`w:val` של `b` כפי שנקרא, או `null` כשאינו שם — כלומר „דלוק”. */
  boldVal: string | null;
  hasBoldCs: boolean;
  italicEnd: number | null;
  italicVal: string | null;
  hasItalicCs: boolean;
  sizeEnd: number | null;
  sizeVal: string | null;
  hasSizeCs: boolean;
  /** המקום שלפני סוגר התג של `rFonts` — לשם נכנס המאפיין `w:cs`. */
  fontsAttrsEnd: number | null;
  fontsAscii: string | null;
  /** ‏`asciiTheme`/`hAnsiTheme` — גובר על השם, כמו ב-Word. */
  fontsTheme: string | null;
  /** יש `cs` או `cstheme` — אחד מהם מספיק כדי שהתאום כבר יהיה שם. */
  hasFontCs: boolean;
}

function newProps(from: number): RunProps {
  return {
    span: { from, to: from },
    declaredRtl: null,
    hasChange: false,
    boldEnd: null,
    boldVal: null,
    hasBoldCs: false,
    italicEnd: null,
    italicVal: null,
    hasItalicCs: false,
    sizeEnd: null,
    sizeVal: null,
    hasSizeCs: false,
    fontsAttrsEnd: null,
    fontsAscii: null,
    fontsTheme: null,
    hasFontCs: false,
  };
}

/**
 * ה-`w:val` שהתאום נושא: זה של השותף, מילה במילה. השותף בלי `w:val` — ‏`<w:b/>`
 * — מוליד תאום בלי `w:val`, ושניהם „דלוק” לפי `ST_OnOff`.
 */
function sameValue(raw: string | null, prefix: string): string {
  return raw === null ? '' : ` ${prefix}:val="${quoteAttribute(raw)}"`;
}

/**
 * ההכנסות שריצה אחת דורשת. ריק = אין מה לעשות בה.
 *
 * רק ריצה שמצהירה `w:rtl` **דלוק**, ורק איבר שהתאום הלטיני שלו קיים והוא
 * עצמו חסר. מצב ההדלקה של הדגל אינו תנאי אלא **תוכן**: הוא עובר אל התאום.
 */
function mirrorFor(props: RunProps, prefix: string): XmlInsert[] {
  if (props.declaredRtl !== true || props.hasChange) return [];
  const out: XmlInsert[] = [];
  const add = (at: number, text: string) => out.push({ at, text, span: props.span });

  if (props.boldEnd !== null && !props.hasBoldCs) {
    add(props.boldEnd, `<${prefix}:bCs${sameValue(props.boldVal, prefix)}/>`);
  }
  if (props.italicEnd !== null && !props.hasItalicCs) {
    add(props.italicEnd, `<${prefix}:iCs${sameValue(props.italicVal, prefix)}/>`);
  }
  if (props.sizeEnd !== null && props.sizeVal !== null && !props.hasSizeCs) {
    add(props.sizeEnd, `<${prefix}:szCs ${prefix}:val="${quoteAttribute(props.sizeVal)}"/>`);
  }
  // גופן ערכת נושא עובר כערכת נושא: נמדד ש-`cstheme="majorHAnsi"` מצייר בדיוק
  // מה שהצד הלטיני צייר. (נמצא ב-QA: בלי זה `asciiTheme` ישיר נפל לגופן העברי
  // של המסמך, בעוד `ascii` ישיר נשמר.)
  if (props.fontsAttrsEnd !== null && !props.hasFontCs) {
    if (props.fontsTheme !== null) add(props.fontsAttrsEnd, ` ${prefix}:cstheme="${quoteAttribute(props.fontsTheme)}"`);
    else if (props.fontsAscii !== null) add(props.fontsAttrsEnd, ` ${prefix}:cs="${quoteAttribute(props.fontsAscii)}"`);
  }
  return out;
}

/**
 * משלימה את מחסנית הכתב המורכב על ריצות שמצהירות `<w:rtl/>`. ‏`null` = אין מה
 * לשנות.
 *
 * ## למה סורק ולא רגקס על `<w:rPr>…</w:rPr>`
 *
 * שתי סיבות שכל אחת מהן לבדה מספיקה, ושתיהן נמדדות במסמכים אמיתיים:
 *
 * 1. **`rPr` מקננת.** `<w:rPr>…<w:rPrChange><w:rPr>…</w:rPr></w:rPrChange></w:rPr>`
 *    — רגקס לא-להוט קושר את הפתיחה החיצונית לסגירה הפנימית, כלומר קורא את
 *    ההיסטוריה של שינוי מסומן כאילו הייתה העיצוב החי.
 * 2. **`rPr` של סימן הפסקה.** `<w:pPr><w:rPr>` אינה של שום ריצה, ואין לה
 *    `w:rtl` משלה שמעניין אותנו. היא מזוהה לפי היותה בתוך `pPr`.
 *
 * הערות ו-CDATA נבלעות שלמות, כמו בשני המודולים האחים.
 */
export function mirrorComplexScript(xml: string): string | null {
  // חיפוש אחד לפני הסריקה: חלק שאין בו אף הצהרה אינו נוגע לזה בכלל.
  if (!/<[\w.-]+:rtl[\s/>]/.test(xml)) return null;
  const prefix = wordPrefix(xml);
  if (prefix === null) return null;

  const inserts: XmlInsert[] = [];
  /** עומק ה-`rPr` בריצה. 1 = העיצוב החי; 2 ומעלה = `rPrChange`. */
  let depth = 0;
  let props: RunProps | null = null;
  /** עומק `pPr`. כל עוד הוא מעל אפס, מה שנסרק אינו של ריצה. */
  let paraProps = 0;

  const token = new RegExp(TOKEN_SOURCE.source, 'g');
  for (let match = token.exec(xml); match; match = token.exec(xml)) {
    const closer = skippedCloser(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      // הערה שאינה נסגרת: מה שנאסף עד כאן שייך ל-`rPr` שכבר נסגרו, ולכן תקף.
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

    if (name === 'pPr' && !selfClosing) {
      paraProps = Math.max(0, paraProps + (closing ? -1 : 1));
      continue;
    }
    if (paraProps > 0) continue;

    if (name === 'rPr' && !selfClosing) {
      if (closing) {
        if (depth === 0) continue;
        depth -= 1;
        if (depth === 0 && props) {
          props.span.to = match.index;
          inserts.push(...mirrorFor(props, prefix));
          props = null;
        }
      } else {
        depth += 1;
        if (depth === 1) props = newProps(token.lastIndex);
      }
      continue;
    }

    // רק הבנים הישרים של ה-`rPr` החיה. `rPrChange` מעלה את העומק, ולכן מה
    // שבתוכה נספר כאן כעומק 2 ואינו נקרא.
    if (depth !== 1 || !props) continue;
    const tagEnd = match.index + match[0].length;
    // Empty properties can also use paired XML tags. Insert siblings after
    // the closing tag, never inside <w:b> or <w:sz>.
    if (closing) {
      if (name === 'b') props.boldEnd = tagEnd;
      else if (name === 'i') props.italicEnd = tagEnd;
      else if (name === 'sz') props.sizeEnd = tagEnd;
      continue;
    }

    if (name === 'b') {
      props.boldEnd = selfClosing ? tagEnd : null;
      props.boldVal = valueOf(attributes);
    } else if (name === 'bCs') props.hasBoldCs = true;
    else if (name === 'i') {
      props.italicEnd = selfClosing ? tagEnd : null;
      props.italicVal = valueOf(attributes);
    } else if (name === 'iCs') props.hasItalicCs = true;
    else if (name === 'sz') {
      props.sizeEnd = selfClosing ? tagEnd : null;
      props.sizeVal = valueOf(attributes);
    } else if (name === 'szCs') props.hasSizeCs = true;
    else if (name === 'rFonts') {
      // המקום שלפני סוגר התג: `/>` כשהוא סוגר את עצמו, ו-`>` כשלא.
      props.fontsAttrsEnd = tagEnd - (selfClosing ? 2 : 1);
      props.fontsAscii = attribute(attributes, 'ascii') ?? attribute(attributes, 'hAnsi');
      props.fontsTheme = attribute(attributes, 'asciiTheme') ?? attribute(attributes, 'hAnsiTheme');
      props.hasFontCs = attribute(attributes, 'cs') !== null || attribute(attributes, 'cstheme') !== null;
    } else if (name === 'rtl') props.declaredRtl = isOn(attributes);
    else if (name === 'rPrChange') props.hasChange = true;
  }

  if (inserts.length === 0) return null;
  return applyXmlInserts(xml, inserts);
}
