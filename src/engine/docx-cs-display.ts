/**
 * הגופן והגודל העבריים של הסגנונות — מוצגים בעורך, ומוחזרים בשמירה.
 *
 * ## מה שדווח
 *
 * „יש לי קובץ וורד עם גופן מסוים לסגנון הרגיל, אך כשאני פותח אותו דרך התוסף
 * הגופן משתנה לגופן אחר.” בקובץ: הגופן הלטיני של Normal הוא „+גוף” (גופן
 * ערכת הנושא) והעברי גופן שהמשתמש התקין. כשהלטיני שונה לאותו גופן — הוצג נכון.
 *
 * ## השורש
 *
 * לכל רמה ב-`styles.xml` שני חריצים: הלטיני (`rFonts@ascii`, `sz`) והמורכב
 * (`rFonts@cs`, `szCs`). ‏Word מצייר ריצה עברית — כל ריצה שמצהירה `w:rtl`, וזו
 * כל ריצה עברית אחרי שמירה אצלנו (`docx-run-direction.ts`) — מהחריץ המורכב.
 * המנוע מצייר כל ריצה מהחריץ הלטיני, ונופל ל-`cs` רק כשכל המסלול הלטיני ריק
 * (superdoc/docx-editor#3994, פתוח). נמדד על 2.23.0 (`scripts/qa/cs-font-qa.mjs`):
 * ‏Normal עם לטיני Courier New ומורכב Tahoma מצויר Courier New, ו-`szCs 40`
 * לצד `sz 24` מצויר 16px.
 *
 * ## מה נעשה
 *
 * **בפתיחה**, בכל רמה — `docDefaults`, כל סגנון, כל `tblStylePr` — החריץ
 * הלטיני מקבל את מה שהמורכב של אותה רמה אומר, כך שהשרשרת הלטינית שהמנוע פותר
 * זהה לשרשרת המורכבת ש-Word פותר:
 *
 * - רמה שמצהירה `cs`/`cstheme` — הלטיני נכתב כשם הגופן המורכב. ערכת נושא
 *   נפתרת לשם מתוך `theme1.xml` (המנוע אינו מכיר את ערכות הנושא המורכבות).
 * - רמה שמצהירה לטיני בלבד, כשרמה **מעליה** מצהירה מורכב — הלטיני מוסר. ‏Word
 *   יורש שם את המורכב מלמעלה, והמנוע ממזג `rFonts` לפי מאפיין (נמדד: רמה עם
 *   `cs` בלבד אינה דורסת את הלטיני של ההורה), ולכן הוא יורש את מה שהוזרק שם.
 *   בלי מורכב מעליה — נשאר: ‏`docx-run-direction.ts` כותב לריצה את הלטיני
 *   האפקטיבי כשאף רמה אינה מצהירה מורכב, ו-Word יצייר אותו.
 * - ‏`szCs`/`sz` — באותו אופן.
 *
 * **בשמירה** (`docx-postflight.ts`) — כל רמה חוזרת למה שהיה בה, **אלא אם
 * המשתמש שינה אותה בינתיים**. ולשם כך המקור נשמר בתוך הקובץ, על ה-`rPr` של
 * אותה רמה, במרחב שמות משלנו — ולא בזיכרון: טיוטת השחזור נכתבת מהמנוע בלי
 * השלב היוצא (`ExportDocxOptions.postflight`) ונפתחת שוב דרך השלב הנכנס, והיא
 * חייבת לשאת את המקור איתה. נמדד שהמנוע משמר מאפיין זר על `rPr` בייצוא.
 *
 * „שינה” = החריץ הלטיני אינו עוד מה שהוזרק. עריכת סגנון אצלנו כותבת תמיד את
 * שלושת החריצים יחד (`docx-style-edit.ts`), ולכן גופן שנבחר בדיאלוג נשמר כמות
 * שהוא, ואילו שינוי של הגודל בלבד מחזיר את הגופן הלטיני המקורי.
 *
 * ## המחיר, במפורש
 *
 * טקסט לטיני **בעורך** מוצג בגופן ובגודל העבריים של הסגנון — זה הצד השני של
 * חריץ אחד. בקובץ הוא נשמר כמו שהיה, ו-Word מציג אותו כמו קודם. וממנו נגזרים:
 *
 * - **PDF והדפסה** מציירים את המסך, ולכן גם בהם הלטיני בגופן העברי.
 * - **בחירה מפורשת בערך שכבר מוצג** — למשל גודל ברירת המחדל שהדיאלוג מציג,
 *   מוקלד שוב כדי להחיל אותו גם על הלטיני — אינה ניתנת להבחנה מ„לא נגעו”,
 *   וההחזרה מחזירה את הלטיני למקור. ערך אחר מזה שמוצג נשמר כרגיל.
 * - **מי שקורא את הגיליון החי כדי לכתוב ממנו** חייב לקרוא אותו דרך
 *   `restoreComplexScriptStyles`: הדבקה במיזוג עיצוב (`paste-destination.ts`)
 *   ו„גם למסמכים חדשים” בדיאלוג הסגנון (`use-style-manager.ts`) עושים זאת.
 *
 * ## מה שאינו ניתן לכיסוי ברמת הגיליון
 *
 * רמה אחת מוצגת בערך אחד, ויש רמות שהמורכב שמעליהן תלוי בהקשר: סגנון תו יורש
 * מסגנון הפסקה שבה הריצה, וסגנון פסקה בתוך טבלה — גם מסגנון הטבלה. כשרק ההקשר
 * מצהיר מורכב (ולא `docDefaults`), נבחרת שרשרת סגנון הפסקה ברירת המחדל עבור
 * סגנון תו, וסגנון פסקה נשאר כמו שהוא; בתוך טבלה כזאת הוא יוצג בלטיני שלו.
 * ‏Word העברי כותב `cstheme` ב-`docDefaults` תמיד, ושם השאלה אינה עולה.
 *
 * ## מתי לא
 *
 * מסמך שנושא את חתימת התבנית של המנוע ואין בו עברית מוצהרת — `docx-cs-align.ts`
 * יישר בשמירה את המורכב ללטיני, כלומר Word יצייר את הלטיני. העורך כבר מצייר
 * אותו, ואין מה להציג.
 */
import { RPR_ORDER, liveRunProps } from './docx-run-direction';
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
import { readRunChild, readStyleSheet, type FontRef, type LevelProps, type StyleSheet } from './docx-style-sheet';

/* ------------------------------------------------------------------ */
/* גופני ערכת הנושא                                                    */
/* ------------------------------------------------------------------ */

/** מה ש-`theme1.xml` נותן לכל אחת משתי הקבוצות. ‏`null` = ריק או חסר. */
interface ThemeFontSet {
  latin: string | null;
  ea: string | null;
  cs: string | null;
  /** ‏`<a:font script="Hebr" typeface="…"/>`, לפי הכתב. */
  scripts: Map<string, string>;
}

export interface ThemeFonts {
  major: ThemeFontSet | null;
  minor: ThemeFontSet | null;
  /** הכתב שלפיו נבחר הגופן המורכב — מ-`w:themeFontLang@w:bidi`. */
  bidiScript: string | null;
}

export const NO_THEME_FONTS: ThemeFonts = { major: null, minor: null, bidiScript: null };

/** שפת ה-bidi → הכתב בטבלת `a:font`. מה שאינו כאן — אין בחירה לפי כתב. */
const BIDI_SCRIPTS: ReadonlyMap<string, string> = new Map([
  ['he', 'Hebr'],
  ['yi', 'Hebr'],
  ['ar', 'Arab'],
  ['fa', 'Arab'],
  ['ur', 'Arab'],
  ['ps', 'Arab'],
  ['sd', 'Arab'],
  ['ug', 'Arab'],
  ['ckb', 'Arab'],
  ['syr', 'Syrc'],
  ['dv', 'Thaa'],
]);

function typefaceOf(block: string, name: string): string | null {
  const match = new RegExp(`<(?:[\\w.-]+:)?${name}\\b([^>]*)>`).exec(block);
  const face = match ? attribute(match[1]!, 'typeface') : null;
  return face ? face : null;
}

function readFontSet(themeXml: string, name: 'majorFont' | 'minorFont'): ThemeFontSet | null {
  const match = new RegExp(`<((?:[\\w.-]+:)?)${name}\\b[^>]*>([\\s\\S]*?)</\\1${name}>`).exec(themeXml);
  if (!match) return null;
  const block = match[2]!;
  const scripts = new Map<string, string>();
  for (const font of block.matchAll(/<(?:[\w.-]+:)?font\b([^>]*)>/g)) {
    const script = attribute(font[1]!, 'script');
    const face = attribute(font[1]!, 'typeface');
    if (script && face) scripts.set(script, face);
  }
  return { latin: typefaceOf(block, 'latin'), ea: typefaceOf(block, 'ea'), cs: typefaceOf(block, 'cs'), scripts };
}

/** גופני ערכת הנושא ושפת ה-bidi של המסמך. חלק חסר — ‏`null` בשדה שלו. */
export function readThemeFonts(themeXml: string | null, settingsXml: string | null): ThemeFonts {
  const lang = settingsXml ? /<(?:[\w.-]+:)?themeFontLang\b([^>]*)>/.exec(settingsXml) : null;
  const bidi = lang ? attribute(lang[1]!, 'bidi') : null;
  return {
    major: themeXml ? readFontSet(themeXml, 'majorFont') : null,
    minor: themeXml ? readFontSet(themeXml, 'minorFont') : null,
    bidiScript: bidi ? (BIDI_SCRIPTS.get(bidi.toLowerCase().split('-')[0]!) ?? null) : null,
  };
}

/**
 * השם שגופן של ערכת נושא נפתר אליו, או `null` כשאי אפשר לדעת. ‏`*Bidi` הוא
 * הגופן המורכב: קודם לפי הכתב של שפת ה-bidi, ואז `a:cs`.
 */
export function themeFontName(theme: string, fonts: ThemeFonts): string | null {
  const set = theme.startsWith('major') ? fonts.major : theme.startsWith('minor') ? fonts.minor : null;
  if (!set) return null;
  const slot = theme.slice(5);
  if (slot === 'Bidi') return (fonts.bidiScript ? set.scripts.get(fonts.bidiScript) : undefined) ?? set.cs;
  if (slot === 'HAnsi' || slot === 'Ascii') return set.latin;
  if (slot === 'EastAsia') return set.ea;
  return null;
}

/** הגופן שרמה מצהירה, כשם. ערכת נושא גוברת על שם באותה רמה — כמו ב-Word. */
export function fontName(ref: FontRef, fonts: ThemeFonts): string | null {
  return ref.theme !== null ? themeFontName(ref.theme, fonts) : ref.name;
}

/* ------------------------------------------------------------------ */
/* הסימון                                                              */
/* ------------------------------------------------------------------ */

/** מרחב השמות של הסימון. ‏`mc:Ignorable` — Word מתעלם ממנו גם אם ידלוף. */
export const DISPLAY_NS = 'urn:otzaria-word:cs-display';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
/** הקידומת שלנו. ייחודית מספיק כדי לא להתנגש בקידומת של מסמך. */
const P = 'owcs';

/** החריצים הלטיניים של `rFonts`, בסדר שבו הם נכתבים. */
const LATIN = ['ascii', 'hAnsi', 'asciiTheme', 'hAnsiTheme'] as const;

/**
 * על `rPr` של רמה שהוצגה:
 * - `owcs:font` — מה שהוזרק לחריץ הלטיני (`""` = הוסר). נוכחותו = הגופן נגע.
 * - `owcs:rFonts` — המאפיינים של תג ה-`rFonts` המקורי, כפי שנכתבו.
 * - `owcs:sz` — הגודל שהוזרק (`""` = הוסר). נוכחותו = הגודל נגע.
 * - `owcs:szTag` — איבר ה-`sz` המקורי, רק כשהיה.
 *
 * המקור נשמר **כמחרוזת** ולא כערכים, כדי שההחזרה תהיה בייט-בבייט: אותו סדר
 * מאפיינים, אותם רווחים. הוא מקודד (`encodeURIComponent`) כי הוא עצמו XML.
 */
const MARKS = ['font', 'rFonts', 'sz', 'szTag'] as const;

const encode = (raw: string): string => encodeURIComponent(raw);
function decode(value: string): string | null {
  try {
    return decodeURIComponent(value);
  } catch {
    return null;
  }
}

/** איבר אחד בתוך `rPr`. */
interface Child {
  name: string;
  from: number;
  /** סוף תג הפתיחה. */
  openEnd: number;
  /** סוף האיבר — תג הסגירה כשהוא זוג. */
  to: number;
  attributes: string;
}

/** רמה אחת: ה-`rPr` שלה, ומי מעליה בשרשרת. */
interface Level {
  /** תג הפתיחה של `rPr`: מ-`<` עד אחרי `>`. */
  open: { from: number; to: number; attributes: string };
  /** תחילת התוכן ועד `</w:rPr>`. */
  span: { from: number; to: number };
  children: Child[];
  /** ‏`null` = `docDefaults`; אחרת — הסגנון שה-`rPr` שלו (או `tblStylePr` שלו). */
  style: string | null;
  /** ‏`tblStylePr` — הסגנון עצמו הוא „מעליה”. */
  conditional: boolean;
}

const LEVEL_PATHS = new Set(['styles/docDefaults/rPrDefault/rPr', 'styles/style/rPr', 'styles/style/tblStylePr/rPr']);
const ORDER = new Map(RPR_ORDER.map((name, index) => [name, index]));

/** כל רמה ב-`styles.xml`, עם הבנים הישירים של ה-`rPr` שלה. */
function readLevels(xml: string, prefix: string): Level[] | null {
  const levels: Level[] = [];
  const stack: string[] = [];
  let style: string | null = null;
  let level: Level | null = null;
  let pending: Child | null = null;

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
    const selfClosing = attributes!.endsWith('/');

    if (closing) {
      if (pending && pending.name === name && stack[stack.length - 1] === name) {
        pending.to = token.lastIndex;
        pending = null;
      }
      const at = stack.lastIndexOf(name!);
      if (at >= 0) stack.length = at;
      if (name === 'rPr' && level && LEVEL_PATHS.has([...stack, 'rPr'].join('/'))) {
        level.span.to = match.index;
        levels.push(level);
        level = null;
      }
      if (name === 'style') style = null;
      continue;
    }

    const path = stack.join('/');
    if (path === 'styles' && name === 'style') style = attribute(attributes!, 'styleId');
    if (name === 'rPr' && !selfClosing && LEVEL_PATHS.has(`${path}/rPr`)) {
      level = {
        open: { from: match.index, to: token.lastIndex, attributes: attributes! },
        span: { from: token.lastIndex, to: token.lastIndex },
        children: [],
        style: path === 'styles/docDefaults/rPrDefault' ? null : style,
        conditional: path === 'styles/style/tblStylePr',
      };
    } else if (level && LEVEL_PATHS.has(path)) {
      const child: Child = { name: name!, from: match.index, openEnd: token.lastIndex, to: token.lastIndex, attributes: attributes! };
      level.children.push(child);
      if (!selfClosing) pending = child;
    }
    if (!selfClosing) stack.push(name!);
  }
  return levels;
}

function childOf(level: Level, name: string): Child | undefined {
  return level.children.find((child) => child.name === name);
}

/** המקום שבו איבר חדש בשם `name` נכתב, לפי סדר הסכמה. */
function slotFor(level: Level, name: string): number {
  const order = ORDER.get(name) ?? RPR_ORDER.length;
  for (const child of level.children) if ((ORDER.get(child.name) ?? RPR_ORDER.length) > order) return child.from;
  return level.span.to;
}

/** החריצים הלטיניים כפי שהם בתג — רק מה שקיים. */
function latinOf(attributes: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const name of LATIN) {
    const value = attribute(attributes, name);
    if (value !== null) out.set(name, value);
  }
  return out;
}

/** מסיר מאפיינים לפי שם מקומי (בכל קידומת), עם הרווח שלפניהם. */
function without(attributes: string, names: readonly string[]): string {
  const pattern = new RegExp(`\\s(?:[\\w.-]+:)?(?:${names.join('|')})\\s*=\\s*(?:"[^"]*"|'[^']*')`, 'g');
  return attributes.replace(pattern, '');
}

/** מאפייני הסימון שלנו בלבד — בלי לגעת ב-`w:sz` או ב-`w:ascii`. */
function withoutMarks(attributes: string): string {
  const pattern = new RegExp(`\\s${P}:(?:${MARKS.join('|')})\\s*=\\s*(?:"[^"]*"|'[^']*')`, 'g');
  return attributes.replace(pattern, '');
}

function mark(attributes: string, name: string): string | null {
  const match = new RegExp(`\\s${P}:${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(attributes);
  return match ? (match[1] ?? match[2]!) : null;
}

/** תג פתיחה חדש לאיבר, עם המאפיינים שנבנו. */
function openTag(prefix: string, name: string, attributes: string, selfClosing: boolean): string {
  return `<${prefix}:${name}${attributes.trimEnd()}${selfClosing ? '/' : ''}>`;
}

/** ‏`rFonts` שהלטיני שלו הוחלף ב-`latin` (מאפיין → ערך גולמי). */
function fontsTag(fonts: Child, latin: ReadonlyMap<string, string>, prefix: string): string {
  const selfClosing = fonts.attributes.endsWith('/');
  const body = without(selfClosing ? fonts.attributes.slice(0, -1) : fonts.attributes, LATIN);
  let next = body.trimEnd();
  for (const [name, value] of latin) next += ` ${prefix}:${name}="${quoteAttribute(value)}"`;
  return openTag(prefix, 'rFonts', next, selfClosing);
}

/* ------------------------------------------------------------------ */
/* הכיוון הנכנס                                                        */
/* ------------------------------------------------------------------ */

/** האם סגנון כלשהו בשרשרת שמתחילה ב-`id` (כולל) מצהיר את התכונה. */
function chainDeclares(sheet: StyleSheet, id: string | null, key: 'fontCs' | 'sizeCs'): boolean {
  const seen = new Set<string>();
  for (let current = id; current !== null && !seen.has(current); ) {
    seen.add(current);
    const record = sheet.styles.get(current);
    if (!record) break;
    if (record.run[key] !== undefined) return true;
    current = record.basedOn;
  }
  return false;
}

/**
 * האם רמה כלשהי **מעל** הרמה הזאת מצהירה את התכונה המורכבת.
 *
 * מעל סגנון תו יושבת גם שרשרת הפסקה, ואותה אי אפשר לדעת מהגיליון לבדו — היא
 * של הפסקה שבה הריצה. נלקחת שרשרת סגנון הפסקה ברירת המחדל: שם יושב רוב הטקסט.
 */
function declaredAbove(sheet: StyleSheet, level: Level, key: 'fontCs' | 'sizeCs'): boolean {
  if (level.style === null) return false;
  if (sheet.defaults[key] !== undefined) return true;
  const record = sheet.styles.get(level.style);
  if (chainDeclares(sheet, level.conditional ? level.style : (record?.basedOn ?? null), key)) return true;
  return record?.type === 'character' && chainDeclares(sheet, sheet.defaultParagraph, key);
}

/** מה שרמה אחת מצהירה, בדיוק כמו ש-`docx-style-sheet.ts` קורא אותו. */
function propsOf(level: Level): LevelProps {
  const props: LevelProps = {};
  for (const child of level.children) readRunChild(props, child.name, child.attributes);
  return props;
}

/**
 * ערך גולמי של מאפיין, בצורה אחת: מה שנכתב בידינו עובר `quoteAttribute`
 * (`"` → `&quot;`), ומה שנקרא ממקור במרכאות בודדות נושא `"` כמות שהוא.
 */
const sameRaw = (value: string) => value.replace(/&quot;/g, '"');

function sameLatin(current: ReadonlyMap<string, string>, wanted: ReadonlyMap<string, string>): boolean {
  if (current.size !== wanted.size) return false;
  for (const [name, value] of wanted) {
    const other = current.get(name);
    if (other === undefined || sameRaw(other) !== sameRaw(value)) return false;
  }
  return true;
}

function shownLatin(name: string): Map<string, string> {
  return new Map([['ascii', name], ['hAnsi', name]]);
}

/**
 * העריכות של רמה אחת בכיוון הנכנס, והסימון שנכתב על ה-`rPr` שלה.
 *
 * רמה שכבר מסומנת — טיוטה שנפתחת שוב — מוצגת מחדש לפי המורכב **של עכשיו**,
 * אבל רק כשהחריץ הלטיני שלה הוא עדיין מה שהוזרק (אחרת המשתמש שינה אותו,
 * וההחזרה תשמור אותו). המקור שרשום עליה אינו נדרס: הוא מה שהיה בקובץ.
 */
function showLevel(xml: string, level: Level, sheet: StyleSheet, fonts: ThemeFonts, prefix: string, out: XmlEdit[]): boolean {
  const attrs = level.open.attributes;
  const props = propsOf(level);
  /** סימון → ערך, כפי שייכתב. */
  const marks = new Map<string, string>();

  const rFonts = childOf(level, 'rFonts');
  if (rFonts) {
    const current = latinOf(rFonts.attributes);
    const shown = mark(attrs, 'font');
    const untouched = shown === null || sameLatin(current, shown ? shownLatin(shown) : new Map());
    let wanted: Map<string, string> | null = null;
    if (props.fontCs) {
      const name = fontName(props.fontCs, fonts);
      if (name !== null) wanted = shownLatin(name);
    } else if (current.size > 0 && declaredAbove(sheet, level, 'fontCs')) {
      wanted = new Map();
    }
    if (untouched && wanted && !sameLatin(current, wanted)) {
      out.push({ at: rFonts.from, end: rFonts.openEnd, text: fontsTag(rFonts, wanted, prefix), span: level.span });
      marks.set('font', quoteAttribute(wanted.get('ascii') ?? ''));
      if (shown === null) marks.set('rFonts', encode(rFonts.attributes));
    }
  }

  const sz = childOf(level, 'sz');
  const size = sz ? valueOf(sz.attributes) : null;
  const shownSize = mark(attrs, 'sz');
  const sizeUntouched = shownSize === null || ((size ?? '') === shownSize && (sz !== undefined) === (shownSize !== ''));
  /** ‏`undefined` = אין מה לעשות; `null` = להסיר. */
  let wantedSize: string | null | undefined;
  if (props.sizeCs !== undefined) wantedSize = props.sizeCs === size ? undefined : props.sizeCs;
  else if (sz && props.size !== undefined && declaredAbove(sheet, level, 'sizeCs')) wantedSize = null;
  if (sizeUntouched && wantedSize !== undefined) {
    const text = wantedSize === null ? '' : `<${prefix}:sz ${prefix}:val="${quoteAttribute(wantedSize)}"/>`;
    if (sz) out.push({ at: sz.from, end: sz.to, text, span: level.span });
    else out.push({ at: slotFor(level, 'sz'), end: slotFor(level, 'sz'), text, span: level.span });
    marks.set('sz', wantedSize === null ? '' : quoteAttribute(wantedSize));
    if (shownSize === null && sz) marks.set('szTag', encode(xml.slice(sz.from, sz.to)));
  }

  if (marks.size === 0) return false;
  const selfClosing = attrs.endsWith('/');
  let body = selfClosing ? attrs.slice(0, -1) : attrs;
  body = body.replace(new RegExp(`\\s${P}:(?:${[...marks.keys()].join('|')})\\s*=\\s*(?:"[^"]*"|'[^']*')`, 'g'), '');
  for (const [name, value] of marks) body = `${body.trimEnd()} ${P}:${name}="${value}"`;
  out.push({
    at: level.open.from,
    end: level.open.to,
    text: openTag(prefix, 'rPr', body, selfClosing),
    span: { from: level.open.from, to: level.open.to },
  });
  return true;
}

/** תג השורש: מיקום ומאפיינים. */
function rootOf(xml: string): { from: number; to: number; attributes: string; name: string; prefix: string } | null {
  const token = new RegExp(TOKEN_SOURCE.source, 'g');
  for (let match = token.exec(xml); match; match = token.exec(xml)) {
    const closer = SKIPPED_SPANS.get(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      if (end < 0) return null;
      token.lastIndex = end + closer.length;
      continue;
    }
    if (match[1]) return null;
    return { from: match.index, to: token.lastIndex, attributes: match[4]!, name: match[3]!, prefix: match[2]! };
  }
  return null;
}

/**
 * ‏`styles.xml` שבו החריץ הלטיני של כל רמה מציג את המורכב שלה, עם המקור רשום
 * עליה; או `null` כשאין מה להציג.
 */
export function showComplexScriptStyles(xml: string, fonts: ThemeFonts): string | null {
  const prefix = wordPrefix(xml);
  if (prefix === null) return null;
  const levels = readLevels(xml, prefix);
  const root = rootOf(xml);
  if (!levels || !root) return null;
  const sheet = readStyleSheet(xml);

  const edits: XmlEdit[] = [];
  let shown = false;
  for (const level of levels) shown = showLevel(xml, level, sheet, fonts, prefix, edits) || shown;
  if (!shown) return null;

  if (!new RegExp(`\\sxmlns:${P}\\s*=`).test(root.attributes)) {
    let attrs = root.attributes.endsWith('/') ? root.attributes.slice(0, -1) : root.attributes;
    attrs = `${attrs.trimEnd()} xmlns:${P}="${DISPLAY_NS}"`;
    const mc = mcPrefix(root.attributes);
    if (mc === null) attrs += ` xmlns:mc="${MC_NS}" mc:Ignorable="${P}" ${P}:mc="added"`;
    else {
      const ignorable = ignorableOf(mc).exec(attrs);
      if (ignorable) {
        const [whole, head, quote, list] = ignorable;
        attrs = attrs.replace(whole, `${head}${quote}${list ? `${list} ` : ''}${P}${quote}`);
      } else attrs += ` ${mc}:Ignorable="${P}" ${P}:mc="ignorable"`;
    }
    edits.push({ at: root.from, end: root.to, text: `<${root.prefix}:${root.name}${attrs}>`, span: { from: root.from, to: root.to } });
  }
  return applyXmlEdits(xml, edits);
}

/** הקידומת של markup-compatibility בשורש, או `null`. שני סוגי המרכאות — כמו `attribute`. */
function mcPrefix(attributes: string): string | null {
  const match = /\sxmlns:([\w.-]+)\s*=\s*(["'])http:\/\/schemas\.openxmlformats\.org\/markup-compatibility\/2006\2/.exec(attributes);
  return match ? match[1]! : null;
}

/** ‏`mc:Ignorable` בשורש: הרווח והשם, המרכאה, והרשימה. */
function ignorableOf(mc: string): RegExp {
  return new RegExp(`(\\s${mc}:Ignorable\\s*=\\s*)(["'])([^"']*)\\2`);
}

/** מאפיין בשם מלא, בכל סוג מרכאות, עם הרווח שלפניו. */
const anyQuoted = (name: string) => new RegExp(`\\s${name}\\s*=\\s*(?:"[^"]*"|'[^']*')`);

/* ------------------------------------------------------------------ */
/* הכיוון היוצא                                                        */
/* ------------------------------------------------------------------ */

/** העריכות שמחזירות רמה אחת למקור — רק מה שלא שונה מאז. */
function restoreLevel(level: Level, prefix: string, out: XmlEdit[]): boolean {
  const attrs = level.open.attributes;
  const font = mark(attrs, 'font');
  const size = mark(attrs, 'sz');
  if (font === null && size === null && withoutMarks(attrs) === attrs) return false;

  if (font !== null) {
    const rFonts = childOf(level, 'rFonts');
    const current = rFonts ? latinOf(rFonts.attributes) : new Map<string, string>();
    const original = decode(mark(attrs, 'rFonts') ?? '');
    if (original !== null && sameLatin(current, font ? shownLatin(font) : new Map())) {
      if (!rFonts) {
        if (latinOf(original).size > 0) {
          const at = slotFor(level, 'rFonts');
          const closed = original.endsWith('/') ? original : `${original}/`;
          out.push({ at, end: at, text: `<${prefix}:rFonts${closed}>`, span: level.span });
        }
      } else {
        // שאר התג לא השתנה — המקור כמות שהוא. השתנה (`cs` אחר, `hint`) — רק
        // החריץ הלטיני חוזר, והשאר נשאר כמו שהוא עכשיו.
        const rest = (text: string) => without(text.endsWith('/') ? text.slice(0, -1) : text, LATIN).trim();
        const same = rest(original) === rest(rFonts.attributes) && original.endsWith('/') === rFonts.attributes.endsWith('/');
        const text = same ? `<${prefix}:rFonts${original}>` : fontsTag(rFonts, latinOf(original), prefix);
        out.push({ at: rFonts.from, end: rFonts.openEnd, text, span: level.span });
      }
    }
  }

  if (size !== null) {
    const sz = childOf(level, 'sz');
    const current = sz ? valueOf(sz.attributes) : null;
    const original = decode(mark(attrs, 'szTag') ?? '');
    if (original !== null && (current ?? '') === size && (sz !== undefined) === (size !== '')) {
      if (sz) out.push({ at: sz.from, end: sz.to, text: original, span: level.span });
      else if (original) out.push({ at: slotFor(level, 'sz'), end: slotFor(level, 'sz'), text: original, span: level.span });
    }
  }

  const selfClosing = attrs.endsWith('/');
  const body = withoutMarks(selfClosing ? attrs.slice(0, -1) : attrs);
  out.push({
    at: level.open.from,
    end: level.open.to,
    text: openTag(prefix, 'rPr', body, selfClosing),
    span: { from: level.open.from, to: level.open.to },
  });
  return true;
}

/**
 * ‏`styles.xml` שבו כל רמה שהוצגה חזרה למקור — חוץ ממה שהמשתמש שינה — והסימון
 * וההכרזה שלו הוסרו; או `null` כשאין סימון.
 */
export function restoreComplexScriptStyles(xml: string): string | null {
  if (!xml.includes(DISPLAY_NS)) return null;
  const prefix = wordPrefix(xml);
  if (prefix === null) return null;
  const levels = readLevels(xml, prefix);
  const root = rootOf(xml);
  if (!levels || !root) return null;

  const edits: XmlEdit[] = [];
  for (const level of levels) restoreLevel(level, prefix, edits);

  let attrs = root.attributes;
  // ‏`added` — הכרזנו גם את mc; `ignorable` — mc היה, `Ignorable` לא; בלי סימן —
  // הוספנו `owcs` לסוף רשימה קיימת, ואותו בדיוק מסירים, כדי שהרשימה תחזור כפי שנכתבה.
  const added = mark(attrs, 'mc');
  attrs = attrs.replace(anyQuoted(`xmlns:${P}`), '').replace(anyQuoted(`${P}:mc`), '');
  const mc = mcPrefix(attrs);
  if (mc !== null) {
    if (added === 'added') attrs = attrs.replace(anyQuoted(`xmlns:${mc}`), '');
    if (added !== null) attrs = attrs.replace(anyQuoted(`${mc}:Ignorable`), '');
    else {
      attrs = attrs.replace(ignorableOf(mc), (_whole, head: string, quote: string, list: string) => {
        const rest = list === P ? '' : list.endsWith(` ${P}`) ? list.slice(0, -P.length - 1)
          : list.split(/\s+/).filter((item) => item && item !== P).join(' ');
        return `${head}${quote}${rest}${quote}`;
      });
    }
  }
  edits.push({ at: root.from, end: root.to, text: `<${root.prefix}:${root.name}${attrs}>`, span: { from: root.from, to: root.to } });
  return applyXmlEdits(xml, edits);
}

/* ------------------------------------------------------------------ */
/* ריצות                                                               */
/* ------------------------------------------------------------------ */

/** הבנים הישירים של `rPr` שמתחיל ב-`from` ונגמר ב-`to`. */
function runLevel(xml: string, prefix: string, from: number, to: number): Level | null {
  const token = new RegExp(TOKEN_SOURCE.source, 'g');
  token.lastIndex = from;
  const head = token.exec(xml);
  if (!head || head.index !== from || head[1] || head[4]!.endsWith('/')) return null;
  const level: Level = {
    open: { from, to: token.lastIndex, attributes: head[4]! },
    span: { from: token.lastIndex, to },
    children: [],
    style: null,
    conditional: false,
  };
  let depth = 0;
  let pending: Child | null = null;
  for (let match = token.exec(xml); match && match.index < to; match = token.exec(xml)) {
    const closer = SKIPPED_SPANS.get(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      if (end < 0) return null;
      token.lastIndex = end + closer.length;
      continue;
    }
    const [, closing, tagPrefix, name, attributes] = match;
    const selfClosing = attributes!.endsWith('/');
    if (closing) {
      if (depth === 0) {
        level.span.to = match.index;
        break;
      }
      depth -= 1;
      if (depth === 0 && pending) {
        pending.to = token.lastIndex;
        pending = null;
      }
      continue;
    }
    if (depth === 0 && tagPrefix === prefix) {
      const child: Child = { name: name!, from: match.index, openEnd: token.lastIndex, to: token.lastIndex, attributes: attributes! };
      level.children.push(child);
      if (!selfClosing) pending = child;
    }
    if (!selfClosing) depth += 1;
  }
  return level;
}

/** בדיקה זולה לפני הסריקה: גופן או גודל מורכבים מפורשים בכלל בחלק. */
const DIRECT_COMPLEX = /:(?:cs|cstheme)\s*=|:szCs\b/;

/**
 * ריצות ש-Word יקרא מהמחסנית המורכבת, ושמצהירות **ישירות** גופן או גודל
 * מורכב: החריץ הלטיני שלהן מקבל את אותו ערך. ‏Word העברי כותב כך כל שינוי של
 * גופן או גודל על בחירה עברית (`<w:rFonts w:hint="cs" w:cs="…"/>`, `w:szCs`).
 *
 * **ובניגוד לסגנונות — אין סימון ואין החזרה**, ולכן רק ריצה ש**כבר מצהירה**
 * `w:rtl` או `w:cs`: ‏Word קורא אותה כולה מהחריץ המורכב, ו-`docx-run-direction.ts`
 * אינו נוגע בריצה מוצהרת — כלומר החריץ הלטיני שלה לא ייקרא לעולם, גם אם
 * יוקלד בה טקסט לטיני. ריצה עברית שאינה מוצהרת אינה נגעת: אילו הוקלדה בה
 * מילה לטינית, השמירה הייתה מפצלת אותה, והחלק הלטיני היה נקרא מהחריץ הלטיני
 * — שכבר היה מחזיק את הגופן העברי.
 *
 * ריצה שמצהירה לטיני בלבד אינה נגעת: ‏`docx-cs-mirror.ts` ממלא בשמירה את
 * המורכב מהלטיני, כלומר העורך כבר מציג את מה ש-Word יציג.
 */
export function showComplexScriptRuns(xml: string, fonts: ThemeFonts): string | null {
  if (!DIRECT_COMPLEX.test(xml)) return null;
  const prefix = wordPrefix(xml);
  if (prefix === null) return null;
  const edits: XmlEdit[] = [];
  for (const run of liveRunProps(xml)) {
    if (run.declared !== true) continue;
    const level = runLevel(xml, prefix, run.from, run.to);
    if (!level) continue;
    const props = propsOf(level);

    const rFonts = childOf(level, 'rFonts');
    if (rFonts && props.fontCs) {
      const name = fontName(props.fontCs, fonts);
      const wanted = name === null ? null : shownLatin(name);
      if (wanted && !sameLatin(latinOf(rFonts.attributes), wanted)) {
        edits.push({ at: rFonts.from, end: rFonts.openEnd, text: fontsTag(rFonts, wanted, prefix), span: level.span });
      }
    }
    if (props.sizeCs !== undefined && props.sizeCs !== props.size) {
      const sz = childOf(level, 'sz');
      const text = `<${prefix}:sz ${prefix}:val="${quoteAttribute(props.sizeCs)}"/>`;
      const at = sz ? sz.from : slotFor(level, 'sz');
      edits.push({ at, end: sz ? sz.to : at, text, span: level.span });
    }
  }
  return edits.length ? applyXmlEdits(xml, edits) : null;
}
