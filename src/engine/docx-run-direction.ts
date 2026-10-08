/**
 * ריצות עבריות מוצהרות כעברית — `<w:rtl/>` — בדרך **החוצה**.
 *
 * ## מה שדווח
 *
 * „בפתיחת קובץ Word קיים דרך התוסף, הטקסט שנכתב בתוכנה מוגדר בהגדרות הגופן
 * של Word כטקסט לטיני, ולכן הוא אינו מושפע מהגדרת הגופן העברי.” כלומר: מה
 * שהמשתמש הקליד אצלנו נראה ב-Word אחרת מהעברית שסביבו, ושינוי הגופן העברי
 * של המסמך אינו חל עליו.
 *
 * ## השורש
 *
 * ‏Word בוחר את מחסנית העיצוב לפי ההצהרה `w:rtl` (ECMA-376 §17.3.2.30), ולא
 * לפי הכתב של התווים. Word עצמו מסמן כל ריצה עברית שמוקלדת בו; המנוע אינו
 * מסמן אף אחת (superdoc/docx-editor#4011). נמדד ב-Word 365 (COM → PDF), מסמך
 * שבו הגופן הלטיני Courier New 10 והעברי David 16:
 *
 * | הריצה | מה Word צייר |
 * |---|---|
 * | בלי `w:rtl` — כמו שהמנוע כותב | Courier New 10 |
 * | עם `w:rtl` — כמו ש-Word כותב | David 16 |
 * | עם `w:cs` במקום `w:rtl` | David 16 |
 *
 * והשורה הראשונה היא גם מה שדיאלוג הגופן של Word מראה: הגופן הלטיני
 * (`Font.Name`) הוא Courier New, והעברי (`Font.NameBi`) הוא David שאינו חל.
 *
 * ## מה נעשה
 *
 * ריצה שיש בה אות ממערכת כתב ימנית, ואין בה הצהרה (`w:rtl` או `w:cs`), מקבלת
 * `<w:rtl/>`. ריצה ניטרלית — נקודה, רווח, ספרות — יורשת את הכיוון של השכנה
 * החזקה הקרובה, כמו ש-Word מסמן את מה שהוקלד במקלדת עברית. ריצה מעורבת
 * מפוצלת לפי כיוון, וזה מה ש-Word כותב כשמקלידים אותו משפט: נמדד שריצה מעורבת
 * מסומנת כולה מציירת גם את המילה הלועזית בגופן העברי (David במקום Courier).
 *
 * ריצה שכבר מצהירה — דלוק או כבוי — אינה נגעת: ההצהרה של מי שכתב את הקובץ
 * גוברת. ריצת קוד שדה (`instrText`) לעולם אינה מסומנת.
 *
 * ## מה נוסע יחד עם הדגל
 *
 * הדגל מעביר את Word למחסנית המורכבת **בכל הרמות**, ומכאן שלושה מקרים:
 *
 * 1. **עיצוב ישיר** שהמנוע כתב בצד הלטיני (Ctrl+B כותב `w:b` בלבד) —
 *    ‏`docx-cs-mirror.ts` משלים את התאום, והוא רץ אחרי המודול הזה.
 * 2. **עיצוב שיורש**, כשהמסמך **מצהיר** את הצד המורכב ברמה כלשהי — נשאר
 *    כמות שהוא. זה בדיוק מה שדווח: הריצה יורשת את הגופן והגודל העבריים של
 *    המסמך, כמו העברית ש-Word עצמו כתב בו, ושינוי שלהם ב-Word חל עליה.
 * 3. **עיצוב שיורש, כשאף רמה אינה מצהירה את הצד המורכב** — כאן Word נופל
 *    ל-Times New Roman 10 (נמדד). לכן הריצה מקבלת את התאום ברמת הריצה, מהערך
 *    הלטיני האפקטיבי — מה שהמנוע מצייר, ומה ש-Word צייר עד עכשיו.
 *
 * גופן מערכת נושא עובר כערכת נושא: `asciiTheme="majorHAnsi"` נכתב
 * `cstheme="majorHAnsi"`. נמדד שזה מצייר בדיוק מה שהצד הלטיני צייר (Arial
 * לעברית בגופן Aptos Display, עם הדגל ובלעדיו), בעוד `majorBidi` — ערכת
 * הנושא העברית — מצייר Times New Roman.
 *
 * ## מה שנפסל בעבר, ולמה זה שונה
 *
 * גרסה קודמת של הקובץ הזה סימנה ריצות ומירתה רק עיצוב ישיר, והוסרה (ראו
 * `docx-neutral-mark.ts`): במסמך שהתוסף יצר, ההצהרה העברית בברירות המחדל היא
 * שם גופן שבור מתבנית המנוע, והכותרות יורשות את ערכת הנושא העברית — וסימון
 * עיוור החליף שם Arial ב-Times New Roman. שני דברים סוגרים את זה עכשיו: רשת
 * הביטחון של מקרה 3, ו-`docx-cs-align.ts`, שמיישר את הצד המורכב של הסגנונות
 * לצד הלטיני במסמך כזה **לפני** שהסימון רץ.
 *
 * ## למה על הבייטים ולא במודל
 *
 * ‏`format.apply({ inline: { rtl: true } })` עובד, אבל הוא מוטציה במסמך: צעד
 * היסטוריה, דגל „לא נשמר” וציור מחדש בכל שמירה. התיקון על הבייטים שיוצאים
 * אינו נוגע במה שהמשתמש רואה. כל העריכות נאכפות ב-`applyXmlEdits` בתוך גבולות
 * הריצה שהן שייכות לה; הפרה מחזירה `null`, והבייטים יוצאים כמות שהם.
 */
import {
  skippedCloser,
  TOKEN_SOURCE,
  applyXmlEdits,
  isOn,
  quoteAttribute,
  valueOf,
  wordPrefix,
  type XmlEdit,
} from './docx-parts';
import { classOf } from './docx-neutral-mark';
import {
  emptyStyleSheet,
  inheritedBidi,
  readRunChild,
  resolveRun,
  type LevelProps,
  type StyleSheet,
} from './docx-style-sheet';

/** בדיקה אחת לפני הסריקה: חלק שאין בו תו ימני אינו נוגע לזה. */
const ANY_RTL =
  /[\p{Script=Hebrew}\p{Script=Arabic}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Adlam}\u200f\u061c]/u;

/**
 * סדר האיברים ב-`CT_RPr` (ECMA-376 §17.3.2.28). כל איבר חדש נכתב לפני הראשון
 * שבא אחריו בסדר הזה. איבר ממרחב שמות אחר (`w14:ligatures`) בא אחרי כולם.
 */
const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
  'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
  'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
  'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange',
];
const ORDER = new Map(RPR_ORDER.map((name, index) => [name, index]));
const FOREIGN_ORDER = RPR_ORDER.length;

type RunDirection = 'rtl' | 'ltr' | 'mixed' | 'neutral';
/** סיווג תו: חזק ימני, חזק שמאלי, ספרה, סימן מצטרף, ניטרלי. */
type CharKind = 'R' | 'L' | 'D' | 'M' | 'N';

/** ה-`rPr` החיה של ריצה. מיקומים הם היסטים ב-XML המקורי. */
interface RunProps {
  openAt: number;
  /** מיקום `</w:rPr>`. */
  closeAt: number;
  closeEnd: number;
  /** הבנים הישירים: השם וסדרו בסכמה, ומיקום תחילתם. */
  children: { order: number; at: number }[];
  /** ‏`null` — אין הצהרה; אחרת — מה ש-`w:rtl`/`w:cs` מצהירים. */
  declared: boolean | null;
  hasChange: boolean;
  rStyle: string | null;
  /**
   * הבנים הישירים כפי שנקראו — שם ומאפיינים. הפירוק שלהם לרמה (`levelOf`)
   * נדחה עד שהריצה באמת מסומנת: ברוב הריצות במסמך זה לעולם אינו קורה.
   */
  tags: [string, string][];
  level: LevelProps | null;
  /** המקום שלפני סוגר התג של `rFonts`, כשהוא קיים. */
  fontsAttrsEnd: number | null;
  /** הטווח של `<w:rtl/>` דלוק שסוגר את עצמו — מה שיוצא מהקטע הלטיני בפיצול. */
  rtlTag: { from: number; to: number } | null;
  /** ‏`<w:rPr/>` שסוגר את עצמו: אין בו מקום להכניס, ולכן הוא מוחלף כולו. */
  empty: boolean;
}

interface RunRecord {
  start: number;
  afterOpen: number;
  /** מיד אחרי `</w:r>`, או `-1` כל עוד הריצה פתוחה. */
  end: number;
  /** הטקסט כפי שהוא ב-XML — ישויות עדיין מקודדות. */
  raw: string;
  /** הטקסט המפוענח, פעם אחת לריצה. */
  decoded: string | null;
  fieldCode: boolean;
  props: RunProps | null;
  /** עומק ה-`rPr`. 1 = העיצוב החי; 2 ומעלה = `rPrChange`. */
  propsDepth: number;
  textFrom: number | null;
  childDepth: number;
  /** רק `rPr` וטקסט — ריצה שמותר לפצל. */
  simple: boolean;
}

interface ParagraphRecord {
  runs: RunRecord[];
  /** ‏`w:bidi` ישיר, או `null` — ואז הוא יורש מהסגנון. */
  bidi: boolean | null;
  pStyle: string | null;
}

const NAMED_ENTITIES: Readonly<Record<string, string>> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

/** הטקסט כפי שהוא נקרא — בלי זה `&amp;` נספר כשלוש אותיות לטיניות. */
function decodeText(raw: string): string {
  return raw.replace(/&(?:#(\d+)|#x([0-9a-fA-F]+)|(amp|lt|gt|quot|apos));/g, (whole, dec, hex, named) => {
    if (named) return NAMED_ENTITIES[named as string] ?? whole;
    const code = dec ? Number.parseInt(dec as string, 10) : Number.parseInt(hex as string, 16);
    return Number.isFinite(code) && code <= 0x10ffff && (code < 0xd800 || code > 0xdfff)
      ? String.fromCodePoint(code)
      : whole;
  });
}

function encodeText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * ‏`classOf` עם מסלול מהיר לשני הטווחים שהם רוב הטקסט: ASCII ואותיות עבריות.
 * התוצאה זהה (נבדק על כל ה-BMP), והבדיקה של כמה מחלקות Unicode לכל תו היא
 * החלק היקר בסריקה של מסמך גדול.
 */
export function kindOf(ch: string): CharKind {
  const code = ch.charCodeAt(0);
  if (code < 0x80) {
    if ((code >= 0x41 && code <= 0x5a) || (code >= 0x61 && code <= 0x7a)) return 'L';
    return code >= 0x30 && code <= 0x39 ? 'D' : 'N';
  }
  if (code >= 0x05d0 && code <= 0x05ea) return 'R';
  return classOf(ch);
}

/** הטקסט של הריצה כפי שהוא נקרא. קוד שדה — ריק: אינו טקסט ואינו נותן כיוון. */
function textOf(run: RunRecord): string {
  if (run.fieldCode) return '';
  if (run.decoded === null) run.decoded = run.raw.includes('&') ? decodeText(run.raw) : run.raw;
  return run.decoded;
}

function classify(run: RunRecord): RunDirection {
  let rtl = false;
  let ltr = false;
  for (const ch of textOf(run)) {
    const kind = kindOf(ch);
    if (kind === 'R') rtl = true;
    else if (kind === 'L') ltr = true;
    if (rtl && ltr) return 'mixed';
  }
  return rtl ? 'rtl' : ltr ? 'ltr' : 'neutral';
}

/**
 * הכיוון של כל תו ברצף, לפי UBA (W7, N1, N2) — `true` = ימין-לשמאל.
 *
 * ספרה אחרי חזק לטיני היא לטינית, וכל ספרה אחרת פועלת כחזקה ימנית על
 * הניטרליים שלצידה. ניטרלי בין שני חזקים זהים מקבל את כיוונם; בין שונים, או
 * בקצה, — את כיוון הפסקה. סימן מצטרף הולך עם התו שלפניו.
 */
function resolveChars(kinds: readonly CharKind[], paragraphRtl: boolean): boolean[] {
  const edge: 'R' | 'L' = paragraphRtl ? 'R' : 'L';
  const strong: ('R' | 'L' | null)[] = [];
  let last: 'R' | 'L' = edge;
  for (const kind of kinds) {
    if (kind === 'R' || kind === 'L') {
      last = kind;
      strong.push(kind);
    } else if (kind === 'D') {
      strong.push(last === 'L' ? 'L' : 'R');
    } else {
      strong.push(null);
    }
  }
  const before: ('R' | 'L')[] = [];
  let previous: 'R' | 'L' = edge;
  for (const value of strong) {
    before.push(previous);
    if (value) previous = value;
  }
  const result: boolean[] = new Array(kinds.length);
  let next: 'R' | 'L' = edge;
  for (let i = kinds.length - 1; i >= 0; i -= 1) {
    const own = strong[i];
    result[i] = (own ?? (before[i] === next ? next : edge)) === 'R';
    if (own) next = own;
  }
  for (let i = 1; i < kinds.length; i += 1) if (kinds[i] === 'M') result[i] = result[i - 1]!;
  return result;
}

/** מה שהעיצוב הישיר מצהיר, מארבעת הזוגות — מפורק בפעם הראשונה שנדרש. */
function levelOf(props: RunProps): LevelProps {
  if (!props.level) {
    props.level = {};
    for (const [name, attributes] of props.tags) readRunChild(props.level, name, attributes);
  }
  return props.level;
}

/** המקום שבו איבר חדש בשם `name` נכתב, לפי סדר הסכמה. */
function slotFor(props: RunProps, name: string): number {
  const order = ORDER.get(name) ?? FOREIGN_ORDER;
  for (const child of props.children) if (child.order > order) return child.at;
  return props.closeAt;
}

/**
 * התאומים שריצה מסומנת צריכה **ברמת הריצה**: רק תכונה שאין לה עיצוב ישיר
 * (את הישיר ממרה `docx-cs-mirror.ts`), ושאף רמה בשרשרת אינה מצהירה את הצד
 * המורכב שלה — מקרה 3 בהערת הפתיחה.
 */
function fallbackTwins(
  direct: LevelProps,
  sheet: StyleSheet,
  pStyle: string | null,
  rStyle: string | null,
): { elements: { name: string; text: (p: string) => string }[]; font: string | null } {
  const { latin, csDeclared } = resolveRun(sheet, pStyle, rStyle);
  // הדגשה ונטייה: „לא מוצהר” הוא „כבוי”, ערך של ממש ולא נפילה לברירת מחדל
  // שרירותית. לכן הן נכתבות רק בשרשרת שאינה מצהירה **שום** תאום מורכב — מסמך
  // שאינו מכיר את הצד המורכב. בשרשרת שמכירה אותו, תאום אחד כן ואחד לא היה
  // מצייר חצי מכל צד: לא מה שהעורך מציג, ולא מה ש-Word מצייר לעברית שלו.
  const unaware = !csDeclared.bold && !csDeclared.italic && !csDeclared.size && !csDeclared.font;
  const elements: { name: string; text: (p: string) => string }[] = [];
  if (direct.bold === undefined && direct.boldCs === undefined && unaware && latin.bold) {
    elements.push({ name: 'bCs', text: (p) => `<${p}:bCs/>` });
  }
  if (direct.italic === undefined && direct.italicCs === undefined && unaware && latin.italic) {
    elements.push({ name: 'iCs', text: (p) => `<${p}:iCs/>` });
  }
  if (direct.size === undefined && direct.sizeCs === undefined && !csDeclared.size && latin.size !== null) {
    const size = latin.size;
    elements.push({ name: 'szCs', text: (p) => `<${p}:szCs ${p}:val="${quoteAttribute(size)}"/>` });
  }
  let font: string | null = null;
  if (direct.font === undefined && direct.fontCs === undefined && !csDeclared.font && latin.font) {
    const { theme, name } = latin.font;
    if (theme !== null) font = `cstheme="${quoteAttribute(theme)}"`;
    else if (name !== null) font = `cs="${quoteAttribute(name)}"`;
  }
  return { elements, font };
}

/** העריכות שמסמנות ריצה אחת. ריק = אין מה לעשות בה. */
function markEdits(run: RunRecord, paragraph: ParagraphRecord, sheet: StyleSheet, prefix: string): XmlEdit[] {
  const span = { from: run.start, to: run.end };
  const props = run.props;
  const p = prefix;

  // עברית בירושה — Word כבר קורא אותה מהצד המורכב, ואין מה לשנות.
  if (resolveRun(sheet, paragraph.pStyle, props?.rStyle ?? null).complex) return [];

  if (!props || props.empty) {
    const { elements, font } = fallbackTwins({}, sheet, paragraph.pStyle, null);
    const fonts = font ? `<${p}:rFonts ${p}:${font}/>` : '';
    const text = `<${p}:rPr>${fonts}${elements.map((e) => e.text(p)).join('')}<${p}:rtl/></${p}:rPr>`;
    return props
      ? [{ at: props.openAt, end: props.closeEnd, text, span }]
      : [{ at: run.afterOpen, end: run.afterOpen, text, span }];
  }
  if (props.declared !== null || props.closeAt < 0) return [];

  const edits: XmlEdit[] = [];
  const insert = (at: number, text: string) => edits.push({ at, end: at, text, span });
  const { elements, font } = fallbackTwins(levelOf(props), sheet, paragraph.pStyle, props.rStyle);
  if (font) {
    if (props.fontsAttrsEnd !== null) insert(props.fontsAttrsEnd, ` ${p}:${font}`);
    else insert(slotFor(props, 'rFonts'), `<${p}:rFonts ${p}:${font}/>`);
  }
  for (const element of elements) insert(slotFor(props, element.name), element.text(p));
  insert(slotFor(props, 'rtl'), `<${p}:rtl/>`);
  return edits;
}

/** מחילה עריכות שנבנו על ה-XML המלא על קטע ממנו שמתחיל ב-`offset`. */
function applyLocal(source: string, edits: readonly XmlEdit[], offset: number): string {
  const sorted = [...edits].sort((first, second) => first.at - second.at);
  const parts: string[] = [];
  let at = 0;
  for (const edit of sorted) {
    parts.push(source.slice(at, edit.at - offset), edit.text);
    at = edit.end - offset;
  }
  parts.push(source.slice(at));
  return parts.join('');
}

/**
 * האם מותר לפצל. ריצה בלי הצהרה — כן. ריצה **מוצהרת** — רק כשההצהרה היא
 * `<w:rtl/>` דלוק שסוגר את עצמו: אז הקטע הלטיני יוצא בלעדיו.
 *
 * למה גם מוצהרת: המנוע מצרף הקלדה לריצה שלידה. מילה לועזית שמוקלדת בשמירה
 * השנייה, ליד טקסט שסומן בשמירה הראשונה, נכנסת לריצה המוצהרת — ו-Word מצייר
 * אותה בגופן העברי (נמצא ב-QA). ‏Word עצמו כמעט אינו כותב כך: בקורפוס, 17
 * מתוך 568,390 ריצות מוצהרות מכילות גם עברית וגם לטינית.
 */
function canSplit(run: RunRecord): boolean {
  if (!run.simple || run.fieldCode || run.end < 0 || run.textFrom !== null) return false;
  const props = run.props;
  if (!props) return true;
  if (props.closeEnd < 0 || props.hasChange) return false;
  return props.declared === null || (props.declared === true && props.rtlTag !== null);
}

/**
 * ריצה מעורבת — כמה ריצות, אחת לכל קטע כיוון. הקטעים הימניים מסומנים כמו
 * ריצה עברית שלמה; הלטיניים יוצאים עם ה-`rPr` המקורית.
 */
function splitEdit(
  xml: string,
  run: RunRecord,
  text: string,
  dirs: readonly boolean[],
  paragraph: ParagraphRecord,
  sheet: StyleSheet,
  prefix: string,
): XmlEdit | null {
  const chars = [...text];
  if (chars.length !== dirs.length || !canSplit(run)) return null;
  const segments: { text: string; rtl: boolean }[] = [];
  chars.forEach((ch, i) => {
    const last = segments[segments.length - 1];
    if (last && last.rtl === dirs[i]) last.text += ch;
    else segments.push({ text: ch, rtl: dirs[i]! });
  });
  if (segments.length < 2) return null;

  const p = prefix;
  const open = xml.slice(run.start, run.afterOpen);
  const props = run.props;
  const plain = props ? xml.slice(props.openAt, props.closeEnd) : '';
  let marked: string;
  let latin = plain;
  if (props?.declared === true && props.rtlTag) {
    // מוצהרת: הקטע העברי נשאר כמות שהוא, והלטיני יוצא בלי ההצהרה.
    marked = plain;
    const { from, to } = props.rtlTag;
    latin = applyLocal(plain, [{ at: from, end: to, text: '', span: { from, to } }], props.openAt);
  } else {
    const markOnly = markEdits(run, paragraph, sheet, prefix);
    marked = props ? applyLocal(plain, markOnly, props.openAt) : markOnly[0]?.text ?? '';
  }
  const replacement = segments
    .map((segment) => {
      const rPr = segment.rtl ? marked : latin;
      return `${open}${rPr}<${p}:t xml:space="preserve">${encodeText(segment.text)}</${p}:t></${p}:r>`;
    })
    .join('');
  return { at: run.start, end: run.end, text: replacement, span: { from: run.start, to: run.end } };
}

/**
 * העריכות שפסקה אחת דורשת. ריצה ניטרלית יורשת מהשכן החזק הקרוב ביותר —
 * קודם אחורה, ואז קדימה; שכן מעורב מוסר את כיוון התו שבקצה הקרוב שלו. פסקה
 * שכולה ניטרלית אינה מסומנת: אין בה עדות לכיוון.
 */
function paragraphEdits(xml: string, paragraph: ParagraphRecord, sheet: StyleSheet, prefix: string): XmlEdit[] {
  const { runs } = paragraph;
  const kinds = runs.map(classify);
  if (!kinds.includes('rtl') && !kinds.includes('mixed')) return [];
  const texts = runs.map(textOf);

  let charDirs: boolean[][] = [];
  if (kinds.includes('mixed')) {
    const classes = texts.flatMap((text) => [...text].map(kindOf));
    const paragraphRtl = paragraph.bidi ?? inheritedBidi(sheet, paragraph.pStyle);
    const resolved = resolveChars(classes, paragraphRtl);
    let cursor = 0;
    charDirs = texts.map((text) => {
      const length = [...text].length;
      const part = resolved.slice(cursor, cursor + length);
      cursor += length;
      return part;
    });
  }

  /*
   * הכיוון שריצה חזקה מוסרת לשכנה ניטרלית. ריצה שכבר מוצהרת — בעצמה או
   * בירושה — מוסרת „לא”: השכנה שלה נכתבה לצידה בידי מי שכתב את ההצהרה, והוא
   * בחר לא להצהיר אותה. נמצא ב-QA: ספרות שהוקלדו ב-Word בין שתי ריצות מוצהרות
   * עברו מ-Courier 10 ל-David 16 בשמירה בלי עריכה.
   */
  const donor = (index: number, fromEnd: boolean): boolean | null => {
    const kind = kinds[index];
    if (kind === 'neutral') return null;
    const props = runs[index]!.props;
    if (props?.declared != null || resolveRun(sheet, paragraph.pStyle, props?.rStyle ?? null).complex) return false;
    if (kind !== 'mixed') return kind === 'rtl';
    const dirs = charDirs[index] ?? [];
    return dirs.length ? dirs[fromEnd ? dirs.length - 1 : 0]! : null;
  };

  const edits: XmlEdit[] = [];
  runs.forEach((run, index) => {
    const kind = kinds[index];
    if (run.end < 0) return;
    if (kind === 'mixed') {
      const split = splitEdit(xml, run, texts[index]!, charDirs[index] ?? [], paragraph, sheet, prefix);
      if (split) edits.push(split);
      return;
    }
    let rtl = kind === 'rtl';
    if (kind === 'neutral' && texts[index] !== '') {
      let inherited: boolean | null = null;
      for (let back = index - 1; back >= 0 && inherited === null; back -= 1) inherited = donor(back, true);
      for (let ahead = index + 1; ahead < runs.length && inherited === null; ahead += 1) {
        inherited = donor(ahead, false);
      }
      rtl = inherited === true;
    }
    if (rtl) edits.push(...markEdits(run, paragraph, sheet, prefix));
  });
  return edits;
}

function newRun(start: number, afterOpen: number): RunRecord {
  return {
    start,
    afterOpen,
    end: -1,
    raw: '',
    decoded: null,
    fieldCode: false,
    props: null,
    propsDepth: 0,
    textFrom: null,
    childDepth: 0,
    simple: true,
  };
}

function newProps(openAt: number): RunProps {
  return {
    openAt,
    closeAt: -1,
    closeEnd: -1,
    children: [],
    declared: null,
    hasChange: false,
    rStyle: null,
    tags: [],
    level: null,
    fontsAttrsEnd: null,
    rtlTag: null,
    empty: false,
  };
}

/**
 * מסמנת ריצות עבריות בחלק אחד של המסמך. ‏`null` = אין מה לשנות.
 *
 * ‏`sheet` הוא גיליון הסגנונות של אותו מסמך, **אחרי** היישור של
 * `docx-cs-align.ts` — רשת הביטחון נשענת על מה שהוא מצהיר.
 *
 * סורק ולא רגקס, משלוש סיבות שכל אחת מספיקה: `rPr` מקננת (`rPrChange`), ריצות
 * מקננות (תיבת טקסט היא פסקה בתוך ריצה), ו-`rPr` של סימן הפסקה שאינה של
 * שום ריצה. הערות, CDATA והוראות עיבוד נבלעות שלמות.
 */
export function markRtlRuns(xml: string, sheet: StyleSheet = emptyStyleSheet()): string | null {
  if (!ANY_RTL.test(xml)) return null;
  const prefix = wordPrefix(xml);
  if (prefix === null) return null;

  const edits: XmlEdit[] = [];
  const paragraphs: ParagraphRecord[] = [];
  const runs: RunRecord[] = [];
  let paraPropsDepth = 0;
  /** עומק האיברים בתוך `pPr` — `bidi` ו-`pStyle` נקראים רק כבנים ישירים. */
  let paraPropsInner = 0;

  const token = new RegExp(TOKEN_SOURCE.source, 'g');
  for (let match = token.exec(xml); match; match = token.exec(xml)) {
    const closer = skippedCloser(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      if (end < 0) break;
      token.lastIndex = end + closer.length;
      continue;
    }

    const closing = match[1]!;
    const tagPrefix = match[2]!;
    const name = match[3]!;
    const attributes = match[4]!;
    const run = runs[runs.length - 1];
    const selfClosing = attributes.endsWith('/');

    // איבר ממרחב שמות אחר בתוך ה-`rPr` החיה — נרשם רק לסדר הכתיבה.
    if (tagPrefix !== prefix) {
      if (run && run.propsDepth === 1 && run.props && !closing && paraPropsDepth === 0) {
        run.props.children.push({ order: FOREIGN_ORDER, at: match.index });
      }
      continue;
    }
    const paragraph = paragraphs[paragraphs.length - 1];

    if (name === 'pPr') {
      if (!selfClosing) paraPropsDepth += closing ? -1 : 1;
      if (paraPropsDepth < 0) paraPropsDepth = 0;
      if (paraPropsDepth <= 1 && !closing) paraPropsInner = 0;
      continue;
    }
    if (paraPropsDepth > 0) {
      if (!closing && paraPropsDepth === 1 && paraPropsInner === 0 && paragraph) {
        if (name === 'bidi') paragraph.bidi = isOn(attributes);
        else if (name === 'pStyle') paragraph.pStyle = valueOf(attributes);
      }
      if (!selfClosing) paraPropsInner = Math.max(0, paraPropsInner + (closing ? -1 : 1));
      continue;
    }

    if (name === 'p' && !selfClosing) {
      if (closing) {
        const record = paragraphs.pop();
        if (record) edits.push(...paragraphEdits(xml, record, sheet, prefix));
      } else {
        paragraphs.push({ runs: [], bidi: null, pStyle: null });
      }
      continue;
    }

    if (name === 'r' && !selfClosing) {
      if (closing) {
        if (run) run.end = token.lastIndex;
        runs.pop();
      } else {
        const record = newRun(match.index, token.lastIndex);
        runs.push(record);
        paragraph?.runs.push(record);
      }
      continue;
    }

    if (!run) continue;

    if (name === 'rPr' && selfClosing && run.propsDepth === 0 && !run.props) {
      run.props = newProps(match.index);
      run.props.closeAt = match.index;
      run.props.closeEnd = token.lastIndex;
      run.props.empty = true;
      continue;
    }
    if (name === 'rPr' && !selfClosing) {
      if (closing) {
        if (run.propsDepth === 0) continue;
        run.propsDepth -= 1;
        if (run.propsDepth === 0 && run.props) {
          run.props.closeAt = match.index;
          run.props.closeEnd = token.lastIndex;
        }
      } else {
        run.propsDepth += 1;
        if (run.propsDepth === 1) run.props = newProps(match.index);
      }
      continue;
    }

    if (run.propsDepth > 0) {
      const props = run.props;
      if (!props || run.propsDepth !== 1 || closing) continue;
      props.children.push({ order: ORDER.get(name) ?? FOREIGN_ORDER, at: match.index });
      props.tags.push([name, attributes]);
      if (name === 'rtl' || name === 'cs') {
        // ‏`w:cs` מעביר את Word לאותה מחסנית; שתי ההצהרות נחשבות הצהרה.
        props.declared = (props.declared ?? false) || isOn(attributes);
        if (name === 'rtl' && selfClosing && isOn(attributes)) {
          props.rtlTag = { from: match.index, to: token.lastIndex };
        }
      } else if (name === 'rStyle') {
        props.rStyle = valueOf(attributes);
      } else if (name === 'rFonts') {
        props.fontsAttrsEnd = match.index + match[0].length - (selfClosing ? 2 : 1);
      } else if (name === 'rPrChange') {
        props.hasChange = true;
      }
      continue;
    }

    if (name === 't') {
      if (selfClosing) continue;
      if (closing) {
        if (run.textFrom !== null) run.raw += xml.slice(run.textFrom, match.index);
        run.textFrom = null;
      } else {
        if (run.childDepth !== 0) run.simple = false;
        run.textFrom = token.lastIndex;
      }
      continue;
    }

    if (name === 'instrText') run.fieldCode = true;
    if (closing) {
      if (run.childDepth > 0) run.childDepth -= 1;
    } else {
      // כל בן ישיר שאינו `rPr` או `t` — טאב, שבירה, שדה, ציור — הופך את הריצה
      // לכזו שאין לפצל.
      if (run.childDepth === 0) run.simple = false;
      if (!selfClosing) run.childDepth += 1;
    }
  }

  while (paragraphs.length > 0) {
    const record = paragraphs.pop();
    if (record) edits.push(...paragraphEdits(xml, record, sheet, prefix));
  }

  if (edits.length === 0) return null;
  return applyXmlEdits(xml, edits);
}

/** ריצה עברית שכבר מוצהרת — `w:rtl` או `w:cs` דלוקים, ואות ימנית בטקסט. */
export function hasDeclaredRtlRuns(xml: string): boolean {
  if (!ANY_RTL.test(xml)) return false;
  const runs = xml.match(/<([\w.-]+):r[\s>][\s\S]*?<\/\1:r>/g) ?? [];
  return runs.some((run) => {
    const declared = /<[\w.-]+:(?:rtl|cs)((?:[^>"']|"[^"]*"|'[^']*')*)\/?>/.exec(run);
    if (!declared || !isOn(declared[1] ?? '')) return false;
    const text = (run.match(/<[\w.-]+:t(?:\s[^>]*)?>([^<]*)</g) ?? []).join('');
    return ANY_RTL.test(decodeText(text));
  });
}
