/**
 * יישור מחסנית הכתב המורכב לצד הלטיני ב-`styles.xml` — למסמך שבו העברית
 * תמיד צוירה מהצד הלטיני.
 *
 * ## למה
 *
 * ‏`docx-run-direction.ts` מסמן ריצות עבריות ב-`w:rtl`, ומאותו רגע Word קורא
 * אותן מהצד המורכב של כל הסגנונות. במסמך Word זה בדיוק המבוקש: הצד המורכב
 * שם הוא ההגדרה העברית שהמשתמש בחר. אבל במסמך **שנוצר בתוסף** הצד המורכב
 * הגיע מתבנית המנוע ואיש לא בחר בו — ושם הוא שבור. נמדד ב-Word 365:
 *
 * | התבנית של המנוע | בלי `w:rtl` | עם `w:rtl` |
 * |---|---|---|
 * | גוף: `ascii="Arial"`, ‏`cs="Times New Roman (Body CS)"` | Arial | **Times New Roman** |
 * | כותרת 1: `asciiTheme="majorHAnsi"`, ‏`cstheme="majorBidi"` | Arial | **Times New Roman** |
 *
 * ‏„Times New Roman (Body CS)” אינו שם של גופן אלא האופן שבו Word ל-Mac מציג
 * את גופן ערכת הנושא בתפריט, שנכתב לקובץ כמחרוזת. בעורך ובכל שמירה עד היום
 * הטקסט הזה הוצג Arial; סימון בלי יישור היה מחליף אותו ב-Times New Roman.
 *
 * ## מה נעשה
 *
 * בכל רמה — `docDefaults`, כל סגנון, וכל `tblStylePr` — התאום המורכב מיושר
 * ללטיני של אותה רמה. כש-שתי השרשראות זהות בכל רמה, הפתרון שלהן זהה לכל
 * צירוף של סגנון פסקה, סגנון תו ותכונות מתחלפות — ולכן Word מצייר את הריצה
 * המסומנת בדיוק כמו שצייר אותה בלי הדגל. גופן ערכת נושא עובר כערכת נושא
 * (`cstheme="majorHAnsi"`), ונמדד שזה מצייר מה שהצד הלטיני צייר.
 *
 * **אבל רק ערך של התבנית נדרס.** תאום שחסר נכתב; תאום שקיים נדרס או נמחק רק
 * כשהוא ערך שהתבנית הביאה — המחרוזת השבורה, או ערכת הנושא העברית
 * (`majorBidi`/`minorBidi`). כל ערך אחר נכתב בידי מישהו: QA מצא מסמך של
 * התוסף שבו המשתמש קבע ב-Word לסגנון Normal גופן עברי David וגודל 16 — בלי
 * להקליד עברית שם — והגרסה הקודמת מחקה את שניהם. בתבנית עצמה כל זוג גודל
 * והדגשה כבר שווה, ולכן היישור שלה אינו צריך לדרוס אף אחד מהם.
 *
 * ## מתי, וזה צר בכוונה
 *
 * רק כשהמסמך נושא את חתימת התבנית — ההצהרה השבורה ב-`docDefaults` — **ואין בו
 * אף ריצה עברית מוצהרת**. התנאי השני הוא ההוכחה שהצד המורכב מעולם לא צויר:
 * במסמך שנערך גם ב-Word יש עברית שנכתבה בו, והיא כבר נראית לפי הצד המורכב —
 * יישור היה משנה אותה. אחרי השמירה הראשונה החתימה נעלמת (ההצהרה יושרה),
 * והמסמך מתנהג כמו כל מסמך Word.
 *
 * מסמך חדש של התוסף אינו מגיע לכאן: התבנית עצמה מיושרת בזמן הבנייה
 * (`blank-document.ts`), באותה פונקציה.
 */
import {
  SKIPPED_SPANS,
  TOKEN_SOURCE,
  applyXmlEdits,
  attribute,
  quoteAttribute,
  valueOf,
  wordPrefix,
  type XmlEdit,
} from './docx-parts';
import { readStyleSheet } from './docx-style-sheet';

/** ההצהרה השבורה שבתבנית המנוע — ראו הערת הפתיחה. */
export const ENGINE_TEMPLATE_CS_FONT = 'Times New Roman (Body CS)';

/** האם `styles.xml` נושא את חתימת התבנית של המנוע. */
export function hasEngineTemplateSignature(stylesXml: string | null): boolean {
  return readStyleSheet(stylesXml).defaults.fontCs?.name === ENGINE_TEMPLATE_CS_FONT;
}

/** איבר אחד בתוך `rPr`: הטווח שלו, כולל תג סגירה כשהוא זוג. */
interface Child {
  from: number;
  /** סוף תג הפתיחה. */
  openEnd: number;
  /** סוף האיבר — תג הסגירה כשהוא זוג. */
  to: number;
  attributes: string;
}

interface Level {
  span: { from: number; to: number };
  children: Map<string, Child>;
}

/** הנתיבים של `rPr` שהם רמה בהיררכיה. */
const LEVEL_PATHS = new Set(['styles/docDefaults/rPrDefault/rPr', 'styles/style/rPr', 'styles/style/tblStylePr/rPr']);

const PAIRS = [
  ['b', 'bCs'],
  ['i', 'iCs'],
  ['sz', 'szCs'],
] as const;

/** ה-`w:val` של איבר, כפי שייכתב לתאום — `''` כשאין. */
function sameValue(attributes: string, prefix: string): string {
  const value = valueOf(attributes);
  return value === null ? '' : ` ${prefix}:val="${quoteAttribute(value)}"`;
}

/**
 * גודל, הדגשה ונטייה: התאום נכתב רק כשהוא חסר. תאום קיים — שונה או לא — אינו
 * נגע: בתבנית אין זוג כזה שאינו שווה, ולכן כל הבדל נכתב בידי מישהו.
 */
function alignPair(level: Level, latin: string, cs: string, prefix: string, out: XmlEdit[]): void {
  const source = level.children.get(latin);
  if (!source || level.children.has(cs)) return;
  const wanted = `<${prefix}:${cs}${sameValue(source.attributes, prefix)}/>`;
  out.push({ at: source.to, end: source.to, text: wanted, span: level.span });
}

/** ערכת הנושא העברית — מה שהתבנית הביאה לכותרות. */
const TEMPLATE_CS_THEMES = new Set(['majorBidi', 'minorBidi']);

/** מסיר מאפיינים לפי שם מקומי, עם הרווח שלפניהם. */
function without(attributes: string, names: readonly string[]): string {
  const pattern = new RegExp(`\\s(?:[\\w.-]+:)?(?:${names.join('|')})\\s*=\\s*(?:"[^"]*"|'[^']*')`, 'g');
  return attributes.replace(pattern, '');
}

/**
 * ‏`rFonts`: ערכת נושא לטינית ← `cstheme` זהה (ו-`cs` אינו נגע — ערכת הנושא
 * גוברת עליו); שם לטיני ← `cs` זהה ובלי `cstheme`; בלי לטיני ← בלי מורכב.
 * וכל זה רק כשהמורכב שבתג חסר או שהוא ערך של התבנית.
 */
function alignFonts(level: Level, prefix: string, out: XmlEdit[]): void {
  const fonts = level.children.get('rFonts');
  if (!fonts) return;
  const attrs = fonts.attributes;
  const theme = attribute(attrs, 'asciiTheme') ?? attribute(attrs, 'hAnsiTheme');
  const name = attribute(attrs, 'ascii') ?? attribute(attrs, 'hAnsi');
  const cs = attribute(attrs, 'cs');
  const csTheme = attribute(attrs, 'cstheme');
  const templateName = cs === null || cs === ENGINE_TEMPLATE_CS_FONT;
  const templateTheme = csTheme === null || TEMPLATE_CS_THEMES.has(csTheme);

  const selfClosing = attrs.endsWith('/');
  const body = selfClosing ? attrs.slice(0, -1) : attrs;
  let next: string;
  if (theme !== null) {
    if (csTheme === theme || !templateTheme || (csTheme === null && !templateName)) return;
    next = `${without(body, ['cstheme']).trimEnd()} ${prefix}:cstheme="${quoteAttribute(theme)}"`;
  } else if (name !== null) {
    if ((cs === name && csTheme === null) || !templateName || !templateTheme) return;
    next = `${without(body, ['cs', 'cstheme']).trimEnd()} ${prefix}:cs="${quoteAttribute(name)}"`;
  } else {
    if ((cs === null && csTheme === null) || !templateName || !templateTheme) return;
    next = without(body, ['cs', 'cstheme']).trimEnd();
  }
  // רק תג הפתיחה מוחלף; זוג פתיחה-סגירה נשאר זוג.
  const text = `<${prefix}:rFonts${next}${selfClosing ? '/' : ''}>`;
  out.push({ at: fonts.from, end: fonts.openEnd, text, span: level.span });
}

/**
 * מיישרת את הצד המורכב של כל רמה לצד הלטיני שלה. ‏`null` = אין מה לשנות.
 * אינה בודקת את התנאי „מתי” — זה של הקורא (`docx-postflight.ts`, הבנייה).
 */
export function alignComplexScriptStyles(xml: string): string | null {
  const prefix = wordPrefix(xml);
  if (prefix === null) return null;

  const edits: XmlEdit[] = [];
  const stack: string[] = [];
  let level: Level | null = null;
  /** איבר זוגי שנפתח ברמה ועוד לא נסגר. */
  let pending: { name: string; child: Child } | null = null;

  const token = new RegExp(TOKEN_SOURCE.source, 'g');
  for (let match = token.exec(xml); match; match = token.exec(xml)) {
    const closer = SKIPPED_SPANS.get(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      if (end < 0) return null;
      token.lastIndex = end + closer.length;
      continue;
    }
    const [, closing, tagPrefix, name, attributes] = match;
    if (tagPrefix !== prefix) continue;
    const selfClosing = attributes.endsWith('/');

    if (closing) {
      if (pending && pending.name === name && stack[stack.length - 1] === name) {
        pending.child.to = token.lastIndex;
        pending = null;
      }
      const at = stack.lastIndexOf(name);
      if (at >= 0) stack.length = at;
      if (name === 'rPr' && level && LEVEL_PATHS.has([...stack, 'rPr'].join('/'))) {
        level.span.to = match.index;
        for (const [latin, cs] of PAIRS) alignPair(level, latin, cs, prefix, edits);
        alignFonts(level, prefix, edits);
        level = null;
      }
      continue;
    }

    const path = stack.join('/');
    if (name === 'rPr' && !selfClosing && LEVEL_PATHS.has(`${path}/rPr`)) {
      level = { span: { from: token.lastIndex, to: token.lastIndex }, children: new Map() };
    } else if (level && LEVEL_PATHS.has(path)) {
      const child: Child = { from: match.index, openEnd: token.lastIndex, to: token.lastIndex, attributes };
      if (!level.children.has(name)) level.children.set(name, child);
      if (!selfClosing) pending = { name, child };
    }
    if (!selfClosing) stack.push(name);
  }

  if (edits.length === 0) return null;
  return applyXmlEdits(xml, edits);
}
