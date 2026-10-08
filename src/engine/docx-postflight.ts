/**
 * התיקונים שהבייטים עוברים בדרך **החוצה**, אחרי שהמנוע כתב את המסמך ולפני
 * שהם נכתבים לקובץ. הצד הנכנס הוא `docx-preflight.ts`, ושניהם יושבים על אותו
 * קורא וכותב zip (`docx-parts.ts`).
 *
 * חמישה תיקונים, בסדר הזה:
 *   - `docx-cs-align.ts` — במסמך שנוצר בתוסף ושאין בו עברית מוצהרת, הצד
 *     המורכב של הסגנונות מיושר לצד הלטיני, ב-`styles.xml`.
 *   - `docx-neutral-mark.ts` — תו ניטרלי שסוגר פסקה עברית מקבל RLM אחריו.
 *   - `docx-run-direction.ts` — ריצה עברית מקבלת `w:rtl`, וריצה מעורבת
 *     מפוצלת. הוא קורא את גיליון הסגנונות **אחרי** היישור.
 *   - `docx-cs-mirror.ts` — ריצה שמצהירה `w:rtl` מקבלת את מחסנית הכתב המורכב
 *     שהמנוע אינו כותב (`bCs`, `iCs`, `szCs`, `rFonts@cs`). רץ אחרי הסימון,
 *     ולכן חל גם על מה שזה עתה סומן.
 *   - `docx-numbering-ids.ts` — לכל הגדרת מספור nsid משלה, ב-`numbering.xml`.
 *
 * ולמה הסדר: הסימון מעביר את Word למחסנית המורכבת, והמראה והיישור הם מה
 * שמבטיח שהיא מחזיקה את מה שהמשתמש ראה.
 */
import { CONTENT_PARTS, rewriteDocxXmlParts, type Bytes } from './docx-parts';
import { uniqueNumberingIds } from './docx-numbering-ids';
import { markNeutralParagraphEnds } from './docx-neutral-mark';
import { mirrorComplexScript } from './docx-cs-mirror';
import { alignComplexScriptStyles, hasEngineTemplateSignature } from './docx-cs-align';
import { hasDeclaredRtlRuns, markRtlRuns } from './docx-run-direction';
import { readStyleSheet, type StyleSheet } from './docx-style-sheet';

/** הספרה אופציונלית, בדיוק כמו ב-`CONTENT_PARTS` — זה אותו שם חלק. */
const NUMBERING_PART = /^word\/numbering\d*\.xml$/i;

/** גיליון הסגנונות הראשי. ‏`stylesWithEffects.xml` הוא עותק ל-Word 2010. */
const STYLES_PART = 'word/styles.xml';

/** החלקים שיש בהם פסקאות — ושם בלבד יש ריצות לסמן. */
const PARAGRAPH_PARTS = /^word\/(?:document|footnotes|endnotes|comments|header|footer)\d*\.xml$/i;

interface Context {
  sheet: StyleSheet;
  /** ‏`styles.xml` אחרי היישור, או `null` כשאין יישור. */
  alignedStyles: string | null;
}

/**
 * מה שהתיקונים צריכים לדעת על המסמך כולו, לפני שחלק אחד נכתב: גיליון
 * הסגנונות, והאם ליישר אותו — מה שתלוי בכל חלקי התוכן ולא רק באחד.
 */
async function prepare(read: (name: string) => Promise<string | null>, names: readonly string[]): Promise<Context> {
  const styles = await read(STYLES_PART);
  let alignedStyles: string | null = null;
  if (hasEngineTemplateSignature(styles)) {
    // ‏`w:rtl` בסגנון כלשהו — עברית שיורשת אותו כבר נקראת מהצד המורכב.
    const original = readStyleSheet(styles);
    let declared = [original.defaults, ...[...original.styles.values()].map((s) => s.run)].some((l) => l.complex);
    for (const name of names) {
      if (!PARAGRAPH_PARTS.test(name)) continue;
      const xml = await read(name);
      if (xml && hasDeclaredRtlRuns(xml)) {
        declared = true;
        break;
      }
    }
    if (!declared && styles) alignedStyles = alignComplexScriptStyles(styles);
  }
  return { sheet: readStyleSheet(alignedStyles ?? styles), alignedStyles };
}

/** הבייטים המתוקנים, או `null` כשאין מה לתקן. */
export function postflightDocx(bytes: Bytes): Promise<Bytes | null> {
  return rewriteDocxXmlParts<Context>(
    bytes,
    (name) => CONTENT_PARTS.test(name),
    (xml, name, context) => {
      let next = name === STYLES_PART ? (context.alignedStyles ?? xml) : xml;
      if (PARAGRAPH_PARTS.test(name)) {
        next = markNeutralParagraphEnds(next) ?? next;
        next = markRtlRuns(next, context.sheet) ?? next;
      }
      next = mirrorComplexScript(next) ?? next;
      if (NUMBERING_PART.test(name)) next = uniqueNumberingIds(next) ?? next;
      return next === xml ? null : next;
    },
    prepare,
  );
}
