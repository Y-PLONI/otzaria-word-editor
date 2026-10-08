/**
 * קריאה של `word/styles.xml` במידה שהסימון של ריצות עבריות צריך — ולא יותר.
 *
 * ## למה צריך את זה בכלל
 *
 * ‏Word בוחר בין שתי מחסניות עיצוב לפי ההצהרה `w:rtl` על הריצה, ולא לפי הכתב
 * של התווים (נמדד, ראו `docx-run-direction.ts`). ריצה שמקבלת את ההצהרה עוברת
 * מהצד הלטיני (`b`, `i`, `sz`, `rFonts@ascii`) לצד של הכתב המורכב (`bCs`,
 * `iCs`, `szCs`, `rFonts@cs`) — **בכל רמות ההיררכיה**: ברירות המחדל של המסמך,
 * שרשרת סגנון הפסקה, שרשרת סגנון התו, והעיצוב הישיר.
 *
 * כשאף רמה אינה מצהירה את הצד המורכב, Word נופל לברירת המחדל המובנית שלו —
 * ‏Times New Roman בגודל 10 (נמדד: מסמך שמצהיר רק `rFonts ascii` ו-`sz 32`
 * מצויר Courier 16 בלי הדגל ו-TNR 10 איתו). כדי לדעת מתי זה קורה צריך לקרוא
 * את השרשרת, וזה מה שהמודול הזה עושה.
 *
 * ## מה נקרא
 *
 * ארבעה זוגות בלבד — הדגשה, נטייה, גודל וגופן — בכל רמה: `docDefaults`
 * ו-`rPr` שהוא בן ישיר של `w:style`, כולל סגנון טבלה. ‏`rPrChange` הוא
 * היסטוריה, ו-`pPr/rPr` שייך לסימן הפסקה, ולכן שניהם אינם נקראים. ‏`tblStylePr`
 * דורש הקשר שורה/תא: ההצהרות שלו נקראות כדי למנוע יישור של עברית קיימת,
 * והפסקאות בטבלה עם עיצוב מותנה אינן מומרות. ובנוסף `w:bidi` של הפסקה,
 * לכיוון הבסיס.
 */
import { SKIPPED_SPANS, TOKEN_SOURCE, attribute, isOn, valueOf, wordPrefix } from './docx-parts';

/**
 * גופן כפי שרמה אחת מצהירה עליו. ‏`theme` גובר על `name` כששניהם באותה רמה
 * — כך Word פותר `rFonts`, וכך גם המנוע.
 */
export interface FontRef {
  name: string | null;
  theme: string | null;
}

/** מה שרמה אחת מצהירה. `undefined` = לא הוצהר ברמה הזאת. */
export interface LevelProps {
  bold?: boolean;
  boldCs?: boolean;
  italic?: boolean;
  italicCs?: boolean;
  size?: string;
  sizeCs?: string;
  font?: FontRef;
  fontCs?: FontRef;
  /** ‏`w:rtl` או `w:cs` — ההצהרה שמעבירה את Word למחסנית המורכבת. */
  complex?: boolean;
}

interface StyleRecord {
  id: string;
  type: string;
  basedOn: string | null;
  run: LevelProps;
  bidi?: boolean;
  /** Conditional table formatting requires row/cell context; preserve it during export. */
  conditional: boolean;
  conditionalRuns: LevelProps[];
}

export interface StyleSheet {
  defaults: LevelProps;
  defaultsBidi?: boolean;
  styles: Map<string, StyleRecord>;
  /** סגנון הפסקה שחל על פסקה בלי `pStyle` (`w:default="1"`). */
  defaultParagraph: string | null;
  defaultTable: string | null;
  /** זיכרון של `resolveRun`, לפי סגנונות הפסקה, התו והטבלה. */
  cache: Map<string, ResolvedRun>;
}

/** גיליון ריק: מסמך בלי `styles.xml`, או כזה שלא נקרא. */
export function emptyStyleSheet(): StyleSheet {
  return { defaults: {}, styles: new Map(), defaultParagraph: null, defaultTable: null, cache: new Map() };
}

/** מה שאיבר אחד בתוך `rPr` תורם לרמה. איבר שאינו מארבעת הזוגות — כלום. */
export function readRunChild(level: LevelProps, name: string, attributes: string): void {
  switch (name) {
    case 'b':
      level.bold = isOn(attributes);
      break;
    case 'bCs':
      level.boldCs = isOn(attributes);
      break;
    case 'i':
      level.italic = isOn(attributes);
      break;
    case 'iCs':
      level.italicCs = isOn(attributes);
      break;
    case 'sz': {
      const value = valueOf(attributes);
      if (value !== null) level.size = value;
      break;
    }
    case 'szCs': {
      const value = valueOf(attributes);
      if (value !== null) level.sizeCs = value;
      break;
    }
    case 'rtl':
    case 'cs':
      level.complex = (level.complex ?? false) || isOn(attributes);
      break;
    case 'rFonts': {
      const theme = attribute(attributes, 'asciiTheme') ?? attribute(attributes, 'hAnsiTheme');
      const fontName = attribute(attributes, 'ascii') ?? attribute(attributes, 'hAnsi');
      if (theme !== null || fontName !== null) level.font = { name: fontName, theme };
      const csTheme = attribute(attributes, 'cstheme');
      const csName = attribute(attributes, 'cs');
      if (csTheme !== null || csName !== null) level.fontCs = { name: csName, theme: csTheme };
      break;
    }
  }
}

/**
 * קורא את גיליון הסגנונות. קלט שאינו נקרא — ‏`null`, בלי הצהרת מרחב שמות —
 * מחזיר גיליון ריק, כלומר „שום רמה אינה מצהירה דבר”.
 */
export function readStyleSheet(xml: string | null): StyleSheet {
  const sheet = emptyStyleSheet();
  if (!xml) return sheet;
  const prefix = wordPrefix(xml);
  if (prefix === null) return sheet;

  /** שמות האיברים הפתוחים, מהשורש. */
  const stack: string[] = [];
  let style: StyleRecord | null = null;
  let conditionalRun: LevelProps | null = null;

  const token = new RegExp(TOKEN_SOURCE.source, 'g');
  for (let match = token.exec(xml); match; match = token.exec(xml)) {
    const closer = SKIPPED_SPANS.get(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      if (end < 0) break;
      token.lastIndex = end + closer.length;
      continue;
    }
    const [, closing, tagPrefix, name, attributes] = match;
    if (tagPrefix !== prefix) continue;
    const selfClosing = attributes.endsWith('/');

    if (closing) {
      // סגירה של מה שנפתח אחרון; XML פגום אינו מפיל את הקריאה, רק מקצר אותה.
      const at = stack.lastIndexOf(name);
      if (at >= 0) stack.length = at;
      if (name === 'style') style = null;
      if (name === 'rPr') conditionalRun = null;
      continue;
    }

    const path = stack.join('/');
    if (path === 'styles/docDefaults/rPrDefault/rPr') {
      readRunChild(sheet.defaults, name, attributes);
    } else if (path === 'styles/docDefaults/pPrDefault/pPr' && name === 'bidi') {
      sheet.defaultsBidi = isOn(attributes);
    } else if (path === 'styles' && name === 'style') {
      const id = attribute(attributes, 'styleId');
      if (id !== null) {
        style = { id, type: attribute(attributes, 'type') ?? 'paragraph', basedOn: null, run: {}, conditional: false, conditionalRuns: [] };
        sheet.styles.set(id, style);
        const isDefault = attribute(attributes, 'default');
        if (style.type === 'paragraph' && isDefault !== null && isOn(` val="${isDefault}"`)) {
          sheet.defaultParagraph = id;
        }
        if (style.type === 'table' && isDefault !== null && isOn(` val="${isDefault}"`)) {
          sheet.defaultTable = id;
        }
      }
    } else if (style && path === 'styles/style') {
      if (name === 'basedOn') style.basedOn = valueOf(attributes);
    } else if (style && path === 'styles/style/rPr') {
      readRunChild(style.run, name, attributes);
    } else if (style && path === 'styles/style/tblStylePr' && (name === 'rPr' || name === 'pPr')) {
      style.conditional = true;
      if (name === 'rPr' && !selfClosing) {
        conditionalRun = {};
        style.conditionalRuns.push(conditionalRun);
      }
    } else if (conditionalRun && path === 'styles/style/tblStylePr/rPr') {
      readRunChild(conditionalRun, name, attributes);
    } else if (style && path === 'styles/style/pPr' && name === 'bidi') {
      style.bidi = isOn(attributes);
    }

    if (!selfClosing) stack.push(name);
  }
  return sheet;
}

/**
 * השרשרת של סגנון, מהשורש אל העלה. שרשרת שיש בה מעגל `basedOn` — ריקה: נמדד
 * ב-QA ש-Word מתעלם מסגנונות כאלה (מעגל A↔B שמצהיר `b` מצויר לא מודגש).
 */
function chainOf(sheet: StyleSheet, id: string | null): StyleRecord[] {
  const chain: StyleRecord[] = [];
  const seen = new Set<string>();
  for (let current = id; current !== null; ) {
    if (seen.has(current)) return [];
    seen.add(current);
    const record = sheet.styles.get(current);
    if (!record) break;
    chain.unshift(record);
    current = record.basedOn;
  }
  return chain;
}

/** מה ש-Word רואה בריצה בלי עיצוב ישיר, בפסקה ובסגנון תו נתונים. */
export interface ResolvedRun {
  /** הערכים הלטיניים האפקטיביים — מה שהמנוע מצייר, ומה ש-Word מצייר בלי הדגל. */
  latin: { bold: boolean; italic: boolean; size: string | null; font: FontRef | null };
  /** הערכים המורכבים האפקטיביים — מה ש-Word מצייר עם הדגל, כשהשרשרת מצהירה אותם. */
  cs: { bold: boolean; italic: boolean; size: string | null; font: FontRef | null };
  /**
   * הריצה כבר עברית בירושה: רמה בשרשרת מצהירה `w:rtl`/`w:cs`. נמדד שזה
   * קיים — מסמכים שנוצרו בכלי אחר נושאים `<w:rtl/>` ב-`docDefaults`.
   */
  complex: boolean;
  /** Row/cell-dependent formatting has not been resolved, so do not convert this paragraph. */
  conditionalTable: boolean;
  /** האם רמה כלשהי בשרשרת מצהירה את התאום המורכב. */
  csDeclared: { bold: boolean; italic: boolean; size: boolean; font: boolean };
}

function pickStyle(sheet: StyleSheet, id: string | null, type: string): string | null {
  if (id === null) return null;
  const record = sheet.styles.get(id);
  return record && record.type === type ? id : null;
}

/**
 * פותר את השרשרת של ריצה: `docDefaults` ← סגנון הטבלה ← סגנון הפסקה ← סגנון התו.
 *
 * גודל וגופן: הרמה הספציפית ביותר שמצהירה גוברת. הדגשה ונטייה הן תכונות
 * **מתחלפות** (ECMA-376 §17.7.3): הערך משרשרת הפסקה ומשרשרת התו מצטרפים
 * ב-XOR, וכל שרשרת נותנת את ההצהרה הקרובה ביותר לעלה שלה.
 */
export function resolveRun(
  sheet: StyleSheet,
  pStyle: string | null,
  rStyle: string | null,
  tableStyle?: string | null,
): ResolvedRun {
  // undefined = outside a table; null = inside a table with the default table style.
  const key = JSON.stringify([pStyle, rStyle, tableStyle ?? null, tableStyle !== undefined]);
  const cached = sheet.cache.get(key);
  if (cached) return cached;

  const table = tableStyle === undefined ? [] : chainOf(sheet, pickStyle(sheet, tableStyle, 'table') ?? sheet.defaultTable);
  const paragraph = chainOf(sheet, pickStyle(sheet, pStyle, 'paragraph') ?? sheet.defaultParagraph);
  const character = chainOf(sheet, pickStyle(sheet, rStyle, 'character'));
  const levels: LevelProps[] = [
    sheet.defaults,
    ...table.map((s) => s.run),
    ...paragraph.map((s) => s.run),
    ...character.map((s) => s.run),
  ];

  const last = <K extends keyof LevelProps>(key: K): LevelProps[K] | undefined => {
    for (let i = levels.length - 1; i >= 0; i -= 1) {
      const value = levels[i]![key];
      if (value !== undefined) return value;
    }
    return undefined;
  };
  const nearest = (chain: StyleRecord[], key: 'bold' | 'italic' | 'boldCs' | 'italicCs'): boolean | undefined => {
    for (let i = chain.length - 1; i >= 0; i -= 1) {
      const value = chain[i]!.run[key];
      if (value !== undefined) return value;
    }
    return undefined;
  };
  const toggle = (key: 'bold' | 'italic' | 'boldCs' | 'italicCs'): boolean =>
    (nearest(paragraph, key) ?? nearest(table, key) ?? sheet.defaults[key] ?? false) !== (nearest(character, key) ?? false);
  const declared = (key: keyof LevelProps): boolean => levels.some((level) => level[key] !== undefined);

  const resolved: ResolvedRun = {
    latin: {
      bold: toggle('bold'),
      italic: toggle('italic'),
      size: last('size') ?? null,
      font: last('font') ?? null,
    },
    cs: {
      bold: toggle('boldCs'),
      italic: toggle('italicCs'),
      size: last('sizeCs') ?? null,
      font: last('fontCs') ?? null,
    },
    complex: last('complex') ?? false,
    conditionalTable: table.some((style) => style.conditional),
    csDeclared: {
      bold: declared('boldCs'),
      italic: declared('italicCs'),
      size: declared('sizeCs'),
      font: declared('fontCs'),
    },
  };
  sheet.cache.set(key, resolved);
  return resolved;
}

/** כיוון הבסיס של פסקה בלי `w:bidi` ישיר: שרשרת הסגנון, ואז `docDefaults`. */
export function inheritedBidi(sheet: StyleSheet, pStyle: string | null, tableStyle?: string | null): boolean {
  const chain = chainOf(sheet, pickStyle(sheet, pStyle, 'paragraph') ?? sheet.defaultParagraph);
  for (let i = chain.length - 1; i >= 0; i -= 1) {
    const value = chain[i]!.bidi;
    if (value !== undefined) return value;
  }
  if (tableStyle !== undefined) {
    const table = chainOf(sheet, pickStyle(sheet, tableStyle, 'table') ?? sheet.defaultTable);
    for (let i = table.length - 1; i >= 0; i -= 1) {
      if (table[i]!.bidi !== undefined) return table[i]!.bidi!;
    }
  }
  return sheet.defaultsBidi ?? false;
}
