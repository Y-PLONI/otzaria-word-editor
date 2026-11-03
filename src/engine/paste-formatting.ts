/** Three paste modes. Merge keeps local emphasis from the source, while both
 * merge and text-only adopt destination run formatting. Script-specific fonts
 * and sizes come from the authored destination OOXML (paste-destination.ts),
 * never by comparing a toolbar value against document defaults.
 */

/* ------------------------------------------------------------------ */
/* צורת ה-fragment כפי שהמודול צורך אותה                               */
/* ------------------------------------------------------------------ */

/**
 * הטיפוסים מוגדרים כאן ולא מיובאים: הם יושבים תחת
 * `superdoc/dist/document-api/...`, ו-import מנתיב פנימי אסור
 * (tests/unit/engine-boundaries). רק השדות שהמודול נוגע בהם מוצהרים; כל השאר
 * עובר הלאה כמות שהוא, ב-spread.
 */
export interface PasteRunMarks {
  bold?: boolean;
  boldCs?: boolean;
  italic?: boolean;
  italicCs?: boolean;
  underline?: boolean;
  strike?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  color?: string;
  highlight?: string;
  fontSizePt?: number;
  fontSizeCsPt?: number;
  fontFamily?: string;
  /**
   * גופן הכתב המורכב של היעד, כשהוא שונה מ-`fontFamily`. פנימי בלבד: לחוזה
   * של המנוע יש גופן אחד לריצה, ולכן הריצה נחתכת לפי כתב (`scriptSegments`)
   * והשדה אינו יוצא מהמודול.
   */
  fontFamilyCs?: string;
  verticalAlign?: 'baseline' | 'superscript' | 'subscript';
}

interface PasteText {
  text: string;
  marks?: PasteRunMarks;
  [key: string]: unknown;
}

type PasteInline =
  | ({ kind: 'text' } & PasteText)
  | { kind: 'hyperlink' | 'sdt'; inlines: readonly PasteInline[]; [key: string]: unknown }
  | { kind: string; [key: string]: unknown };

export interface PasteParagraph {
  kind: 'paragraph' | 'heading';
  runs: readonly PasteText[];
  inlines?: readonly PasteInline[];
  [key: string]: unknown;
}

interface PasteTableCell {
  blocks: readonly PasteParagraph[];
  [key: string]: unknown;
}

interface PasteTable {
  kind: 'table';
  rows: readonly { cells: readonly PasteTableCell[]; [key: string]: unknown }[];
  [key: string]: unknown;
}

export type PasteBlock = PasteParagraph | PasteTable | { kind: string; [key: string]: unknown };

export interface PasteFragment {
  blocks: readonly PasteBlock[];
  [key: string]: unknown;
}

/** שלוש האפשרויות, בשמות של Word. */
export type PasteMode = 'keep-source' | 'merge' | 'text-only';

export const PASTE_MODES: readonly PasteMode[] = ['keep-source', 'merge', 'text-only'];

export function isPasteMode(value: unknown): value is PasteMode {
  return typeof value === 'string' && (PASTE_MODES as readonly string[]).includes(value);
}

/* ------------------------------------------------------------------ */
/* ההדגשות                                                            */
/* ------------------------------------------------------------------ */

/**
 * מה ש„מזג עיצוב” משאיר מהמקור.
 *
 * מודגש, נטוי וקו תחתון הם הרשימה ש-Word עצמו נוקב בה („formatting that is
 * considered emphasis”). `boldCs`/`italicCs` הם אותה הדגשה במחסנית הכתב
 * המורכב — ריצה עברית שהגיעה מ-Word מודגשת דרכם. כתב עילי ותחתי נשמרים כי
 * הם משמעות ולא מראה: `m²` שמאבד את ה-`²` הוא טקסט אחר.
 *
 * קו חוצה, אותיות רישיות, צבע, סימון, גופן וגודל — אינם. אלה המראה של המקור,
 * וזה בדיוק מה שהמשתמש ביקש להשאיר מאחור.
 */
const EMPHASIS = ['bold', 'boldCs', 'italic', 'italicCs', 'underline'] as const;
type Emphasis = (typeof EMPHASIS)[number];

/** HTML bold/italic applies to Hebrew and Latin alike. The SDK parses HTML
 * into Latin flags only; private DOCX fragments may intentionally split scripts.
 */
function emphasisValue(marks: PasteRunMarks | undefined, key: Emphasis, html: boolean): boolean | undefined {
  if (html && key === 'boldCs' && marks?.boldCs === undefined) return marks?.bold;
  if (html && key === 'italicCs' && marks?.italicCs === undefined) return marks?.italic;
  return marks?.[key];
}

function emphasisOf(marks: PasteRunMarks | undefined, dropped: ReadonlySet<Emphasis>, html: boolean): PasteRunMarks {
  const kept: PasteRunMarks = {};
  if (!marks) return kept;
  for (const key of EMPHASIS) {
    if (emphasisValue(marks, key, html) === true && !dropped.has(key)) kept[key] = true;
  }
  if (marks.verticalAlign === 'superscript' || marks.verticalAlign === 'subscript') {
    kept.verticalAlign = marks.verticalAlign;
  }
  return kept;
}

/** כל ריצות הטקסט ב-fragment, כולל בתוך קישורים ותאי טבלה. */
function* textRuns(blocks: readonly PasteBlock[]): Generator<PasteText> {
  for (const block of blocks) {
    if (block.kind === 'table') {
      for (const row of (block as PasteTable).rows ?? []) {
        for (const cell of row.cells ?? []) yield* textRuns(cell.blocks ?? []);
      }
      continue;
    }
    if (block.kind !== 'paragraph' && block.kind !== 'heading') continue;
    const paragraph = block as PasteParagraph;
    // `inlines` הוא הייצוג המלא כשהוא קיים, ו-`runs` הוא השטוח שלו. שניהם
    // נושאים את אותו טקסט, ולכן נספרים מאחד מהם בלבד.
    if (Array.isArray(paragraph.inlines)) {
      for (const inline of paragraph.inlines) {
        if (inline.kind === 'text') yield inline as PasteText;
        else if (inline.kind === 'hyperlink' || inline.kind === 'sdt') {
          yield* textRuns([{ kind: 'paragraph', runs: [], inlines: (inline as { inlines: readonly PasteInline[] }).inlines }]);
        }
      }
    } else {
      yield* paragraph.runs ?? [];
    }
  }
}

/**
 * ההדגשות שחלות על **כל** הטקסט שהודבק, ולכן נזרקות.
 *
 * זו ההבחנה של Word: הדגשה נשמרת „when it is applied to only a portion of the
 * selection”. כותרת מודגשת שהועתקה כולה אינה הדגשה בתוך טקסט אלא המראה של
 * הפסקה במקור — וזה מה שהמיזוג מוותר עליו. מילה מודגשת בתוך משפט היא הדגשה,
 * והיא נשארת.
 *
 * ריצות של רווחים בלבד אינן מכריעות: רווח בין שתי מילים מודגשות אינו מודגש
 * לרוב, וזה לא הופך את המשפט ל„מודגש בחלקו”.
 */
function uniformEmphasis(blocks: readonly PasteBlock[], html: boolean): Set<Emphasis> {
  const uniform = new Set<Emphasis>(EMPHASIS);
  let any = false;
  for (const run of textRuns(blocks)) {
    if (typeof run.text !== 'string' || run.text.trim() === '') continue;
    any = true;
    for (const key of EMPHASIS) {
      if (emphasisValue(run.marks, key, html) !== true) uniform.delete(key);
    }
  }
  return any ? uniform : new Set();
}

/* ------------------------------------------------------------------ */
/* בניית ה-fragment                                                   */
/* ------------------------------------------------------------------ */

/** שדות של בלוק פסקה שהם עיצוב המקור, ולא תוכן. */
const PARAGRAPH_FORMATTING = ['styleRef', 'semanticStyleRole', 'paragraphProps', 'pPrXml'] as const;

/** שדות של רשימה. נשמרים במיזוג ונזרקים ב„טקסט בלבד”. */
const LIST_FIELDS = ['numbering', 'listIntent', 'listItemContinuationLevel'] as const;

interface Rewrite {
  marks: (source: PasteRunMarks | undefined) => PasteRunMarks;
  keepLists: boolean;
}

function withoutKeys<T extends Record<string, unknown>>(object: T, keys: readonly string[]): T {
  const copy: Record<string, unknown> = { ...object };
  for (const key of keys) delete copy[key];
  return copy as T;
}

/* ------------------------------------------------------------------ */
/* יעד שהגופן או הגודל שלו שונים בין עברית ללטינית                       */
/* ------------------------------------------------------------------ */

/** אותיות שנכתבות במחסנית הכתב המורכב: עברית, ערבית וצורות ההצגה שלהן. */
const COMPLEX_LETTER = /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/u;
const LETTER = /\p{L}/u;

/**
 * הטקסט בקטעים לפי כתב. תו ניטרלי (רווח, פיסוק, ספרה) מצטרף לקטע שלפניו —
 * „שלום, world” הוא „שלום, ” ו-„world” — ובתחילת הריצה לקטע הראשון שאחריו.
 * ריצה בלי אף אות נחשבת עברית: ב-RTL ספרות ופיסוק נקראים מהמחסנית המורכבת.
 */
export function scriptSegments(text: string): { text: string; complex: boolean }[] {
  const segments: { text: string; complex: boolean | null }[] = [];
  for (const char of text) {
    const complex = COMPLEX_LETTER.test(char) ? true : LETTER.test(char) ? false : null;
    const last = segments[segments.length - 1];
    if (last && (complex === null || last.complex === complex)) last.text += char;
    else if (last && last.complex === null) { last.text += char; last.complex = complex; }
    else segments.push({ text: char, complex });
  }
  return segments.map((segment) => ({ text: segment.text, complex: segment.complex !== false }));
}

/** הסימנים של קטע אחד: גופן וגודל אחידים מהערוץ של הכתב שלו. */
function scriptMarks(marks: PasteRunMarks, complex: boolean): PasteRunMarks {
  const { fontFamilyCs, ...rest } = marks;
  const family = complex ? fontFamilyCs ?? marks.fontFamily : marks.fontFamily;
  const size = complex ? marks.fontSizeCsPt ?? marks.fontSizePt : marks.fontSizePt ?? marks.fontSizeCsPt;
  return {
    ...rest,
    ...(family !== undefined ? { fontFamily: family } : {}),
    ...(size !== undefined ? { fontSizePt: size, fontSizeCsPt: size } : {}),
  };
}

function rewriteText<T extends PasteText>(run: T, rewrite: Rewrite): T[] {
  const marks = rewrite.marks(run.marks);
  const { marks: _source, ...rest } = run;
  const split = (marks.fontFamilyCs !== undefined && marks.fontFamilyCs !== marks.fontFamily) ||
    (marks.fontSizeCsPt !== undefined && marks.fontSizePt !== undefined && marks.fontSizeCsPt !== marks.fontSizePt);
  if (!split || typeof run.text !== 'string') {
    const { fontFamilyCs: _cs, ...uniform } = marks;
    return [(Object.keys(uniform).length > 0 ? { ...rest, marks: uniform } : rest) as T];
  }
  return scriptSegments(run.text).map((segment) => ({ ...rest, text: segment.text, marks: scriptMarks(marks, segment.complex) }) as T);
}

function rewriteInline(inline: PasteInline, rewrite: Rewrite): PasteInline[] {
  if (inline.kind === 'text') return rewriteText(inline as PasteText & { kind: 'text' }, rewrite);
  if (inline.kind === 'hyperlink' || inline.kind === 'sdt') {
    const link = inline as { kind: 'hyperlink' | 'sdt'; inlines: readonly PasteInline[] };
    return [{ ...link, inlines: (link.inlines ?? []).flatMap((inner) => rewriteInline(inner, rewrite)) }];
  }
  return [inline];
}

/**
 * פסקה, בלי עיצוב המקור שלה.
 *
 * כותרת הופכת לפסקה: „כותרת 1” של המקור היא סגנון, והמיזוג מאמץ את סגנון
 * היעד. `level` נזרק יחד עם `kind` — הוא שדה של כותרת בלבד.
 */
function rewriteParagraph(block: PasteParagraph, rewrite: Rewrite): PasteParagraph {
  const dropped = [...PARAGRAPH_FORMATTING, ...(rewrite.keepLists ? [] : LIST_FIELDS), 'level'];
  const base = withoutKeys(block, dropped);
  return {
    ...base,
    kind: 'paragraph',
    runs: (block.runs ?? []).flatMap((run) => rewriteText(run, rewrite)),
    ...(Array.isArray(block.inlines)
      ? { inlines: block.inlines.flatMap((inline) => rewriteInline(inline, rewrite)) }
      : {}),
  };
}

/**
 * טבלה נשארת טבלה — המבנה הוא תוכן. מה שנזרק הוא עיצוב התאים (צבע מילוי,
 * ריפוד) ועיצוב הפסקאות שבתוכם; `gridSpan` הוא מבנה ונשאר.
 */
function rewriteTable(block: PasteTable, rewrite: Rewrite): PasteTable {
  return {
    ...block,
    rows: (block.rows ?? []).map((row) => ({
      ...row,
      cells: (row.cells ?? []).map((cell) => ({
        ...withoutKeys(cell, ['shadingColor', 'padding']),
        blocks: (cell.blocks ?? []).map((paragraph) => rewriteParagraph(paragraph, rewrite)),
      })),
    })),
  };
}

function rewriteFragment<T extends PasteFragment>(fragment: T, rewrite: Rewrite): T {
  const blocks = (fragment.blocks ?? []).map((block): PasteBlock => {
    if (block.kind === 'paragraph' || block.kind === 'heading') {
      return rewriteParagraph(block as PasteParagraph, rewrite);
    }
    if (block.kind === 'table') return rewriteTable(block as PasteTable, rewrite);
    // קו אופקי ושאר הבלוקים שאין בהם עיצוב תווים — כמות שהם.
    return block;
  });
  return { ...fragment, blocks };
}

/**
 * „מזג עיצוב”: עיצוב היעד, וההדגשות של המקור מעליו.
 *
 * הסדר הוא ההכרעה: הדגשה של המקור **מוסיפה** על היעד ואינה מוחקת ממנו. טקסט
 * רגיל שהודבק לתוך משפט מודגש יוצא מודגש — כמו הקלדה באותו מקום.
 */
export function mergeFragment<T extends PasteFragment>(fragment: T, destination: PasteRunMarks, { mirrorHtmlEmphasis = false }: { mirrorHtmlEmphasis?: boolean } = {}): T {
  const dropped = uniformEmphasis(fragment.blocks ?? [], mirrorHtmlEmphasis);
  return rewriteFragment(fragment, {
    marks: (source) => ({ ...destination, ...emphasisOf(source, dropped, mirrorHtmlEmphasis) }),
    keepLists: true,
  });
}

/**
 * „שמור טקסט בלבד”: עיצוב היעד בלבד. הרשימות נזרקות — ב-Word הן הופכות
 * לפסקאות רגילות.
 *
 * תמונות נזרקות כבר לפני כן, בבחירת הפריטים מהלוח (`textOnlyPayload` ב-
 * clipboard.ts), ולכן fragment שנבנה מ-`text/plain` אינו נושא אותן.
 */
export function textOnlyFragment<T extends PasteFragment>(fragment: T, destination: PasteRunMarks): T {
  return rewriteFragment(fragment, { marks: () => ({ ...destination }), keepLists: false });
}

/* ------------------------------------------------------------------ */
/* עיצוב היעד                                                         */
/* ------------------------------------------------------------------ */

/** מה שנקרא בסמן — הערכים שהרצועה מציגה. `undefined` = לא ידוע. */
export interface CaretFormat {
  fontFamily?: string | null;
  fontSizePt?: number | null;
  color?: string | null;
  highlight?: string | null;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
}

/**
 * שש-עשרה צבעי הסימון של Word (`ST_HighlightColor`), לפי ה-hex שהרצועה מדווחת.
 *
 * נמדד: `highlight: 'yellow'` ב-fragment נכתב כ-`<w:highlight w:val="yellow"/>`,
 * ו-`'#FFFF00'` — כ-`<w:shd w:fill="FFFF00"/>`, כלומר **הצללה**. על המסך הם
 * זהים, אבל ב-Word „צבע סימון טקסט” אינו מזהה הצללה, ו„ללא צבע” אינו מסיר
 * אותה. לכן סימון שהוא אחד מהצבעים האלה נשלח בשמו — כמו שה-`parse` של המנוע
 * עצמו שולח אותו. צבע אחר נשאר hex: הצללה באותו צבע היא הקרוב ביותר שיש.
 */
const HIGHLIGHT_NAMES: Record<string, string> = {
  '#FFFF00': 'yellow',
  '#00FF00': 'green',
  '#00FFFF': 'cyan',
  '#FF00FF': 'magenta',
  '#0000FF': 'blue',
  '#FF0000': 'red',
  '#000080': 'darkBlue',
  '#008080': 'darkCyan',
  '#008000': 'darkGreen',
  '#800080': 'darkMagenta',
  '#800000': 'darkRed',
  '#808000': 'darkYellow',
  '#808080': 'darkGray',
  '#C0C0C0': 'lightGray',
  '#000000': 'black',
  '#FFFFFF': 'white',
};

/** Scalar states explicitly include black and false: omission would re-enable
 * a paragraph/document default that the destination run has turned off. Fonts
 * and sizes are supplied independently from the destination OOXML capture.
 */
export function destinationMarks(
  caret: CaretFormat,
): PasteRunMarks {
  const marks: PasteRunMarks = {};

  const color = caret.color?.toUpperCase();
  if (color && /^#[0-9A-F]{6}$/.test(color)) marks.color = color;

  const highlight = caret.highlight?.toUpperCase();
  if (highlight && /^#[0-9A-F]{6}$/.test(highlight)) marks.highlight = HIGHLIGHT_NAMES[highlight] ?? highlight;

  if (typeof caret.bold === 'boolean') marks.bold = caret.bold;
  if (typeof caret.italic === 'boolean') marks.italic = caret.italic;
  if (typeof caret.underline === 'boolean') marks.underline = caret.underline;
  if (typeof caret.strikethrough === 'boolean') marks.strike = caret.strikethrough;

  return marks;
}
