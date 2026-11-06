/**
 * התיקונים שהבייטים עוברים בדרך **החוצה**, אחרי שהמנוע כתב את המסמך ולפני
 * שהם נכתבים לקובץ. הצד הנכנס הוא `docx-preflight.ts`, ושניהם יושבים על אותו
 * קורא וכותב zip (`docx-parts.ts`).
 *
 * שישה תיקונים, בסדר הזה:
 *   - `docx-cs-display.ts` — מה שהוצג בפתיחה בחריץ הלטיני של הסגנונות חוזר
 *     למקור, ב-`styles.xml`. ראשון: כל מה שאחריו קורא את הגיליון המקורי.
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
import { PARAGRAPH_PARTS, alignComplexScriptStyles, templateAlignmentApplies } from './docx-cs-align';
import { restoreComplexScriptStyles } from './docx-cs-display';
import { markRtlRuns } from './docx-run-direction';
import { readStyleSheet, type StyleSheet } from './docx-style-sheet';

/** הספרה אופציונלית, בדיוק כמו ב-`CONTENT_PARTS` — זה אותו שם חלק. */
const NUMBERING_PART = /^word\/numbering\d*\.xml$/i;

/** גיליון הסגנונות הראשי. ‏`stylesWithEffects.xml` הוא עותק ל-Word 2010. */
const STYLES_PART = 'word/styles.xml';

interface Context {
  sheet: StyleSheet;
  /** ‏`styles.xml` אחרי ההחזרה והיישור, או `null` כששניהם לא נגעו בו. */
  styles: string | null;
}

/**
 * מה שהתיקונים צריכים לדעת על המסמך כולו, לפני שחלק אחד נכתב: גיליון
 * הסגנונות, והאם ליישר אותו — מה שתלוי בכל חלקי התוכן ולא רק באחד.
 */
async function prepare(read: (name: string) => Promise<string | null>, names: readonly string[]): Promise<Context> {
  const original = await read(STYLES_PART);
  const restored = original ? restoreComplexScriptStyles(original) : null;
  const styles = restored ?? original;
  const sheet = readStyleSheet(styles);
  const aligned = styles && (await templateAlignmentApplies(sheet, names, read)) ? alignComplexScriptStyles(styles) : null;
  return { sheet: aligned ? readStyleSheet(aligned) : sheet, styles: aligned ?? restored };
}

/** הבייטים המתוקנים, או `null` כשאין מה לתקן. */
export function postflightDocx(bytes: Bytes): Promise<Bytes | null> {
  return rewriteDocxXmlParts<Context>(
    bytes,
    (name) => CONTENT_PARTS.test(name),
    (xml, name, context) => {
      let next = name === STYLES_PART ? (context.styles ?? xml) : xml;
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
