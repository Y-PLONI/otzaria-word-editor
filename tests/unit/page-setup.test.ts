/**
 * לשונית „פריסה”: מה **בדיוק** נשלח למנוע.
 *
 * הכפיל כאן אינו מקליט קריאות — הוא **מאמת** אותן ואז ממיר אותן ל-OOXML כפי
 * שהמנוע עושה. זו לא הקפדה מיותרת: `tests/unit/ribbon-commands.test.ts` הוא
 * הדוגמה ההפוכה — mock שמחזיר `true` לכל קריאה, וכך אישר בירוק payloads
 * שהמנוע דוחה. שני הכללים שהכפיל אוכף הם אלה שנמדדו במימוש
 * (`@superdoc/docx-engine`):
 *
 *   1. **הוולידציה**: `top/right/bottom/left/gutter/width/height/gap` חייבים
 *      להיות מספר סופי אי-שלילי, `count` מספר שלם חיובי, `paperSize` מחרוזת
 *      לא ריקה, וחייב להגיע לפחות שדה אחד. קלט פסול **זורק**, ואינו מחזיר
 *      קבלה — וזו הסיבה שהמודול עוטף כל קריאה ב-try.
 *   2. **ההמרה**: המנוע כותב `String(Math.round(value * 1440))`. כלומר ה-API
 *      מקבל אינצ'ים, וה-XML נמדד ב-twips. הבדיקות למטה משוות את ה-twips
 *      שנכתבו למספרים שנמדדו ב-`word/document.xml` של המסמך הריק של המנוע:
 *      `w:pgMar w:top="1440"` ו-`w:pgSz w:w="12240" w:h="15840"`.
 *
 * ## גל 10 — הכפיל חייב לדעת לייצר מסמך שבור
 *
 * חמש הפעולות של „פריסת עמוד מתקדמת” נוספו כאן, ועם הבדיקות שלהן נוספו שני
 * דברים שאין בכפיל הקודם:
 *
 * 1. **הכפיל מרנדר `sectPr` אמיתי** (`sectPrXml`), ולא רק זוכר מספרים. זה מה
 *    שמאפשר לשאול „האם מה שנכתב לקובץ חוקי” ולא „האם נשלח מה שהתכוונו”.
 * 2. **הכפיל בולע בדיוק את מה שהמנוע בולע.** נמדד: `style: 'zigzag'` נכתב
 *    `w:val="zigzag"`, `style: ''` מייצר גבול **בלי `w:val`**, `size: 2.5`
 *    נכתב `w:sz="2.5"`, `space: 999` ו-`color: '#FF0000'` נכתבים כמות שהם,
 *    `header: 99` הופך ל-`w:header="142560"`, ו-`chapterStyle` נעלם בשקט.
 *    כפיל שהיה מאמת את הערכים האלה בעצמו היה מחזיר `ok: false` על מוטציה
 *    ומאשר בירוק שהמודול „מוגן” — בעוד שבמנוע האמיתי אותה מוטציה מחזירה
 *    `success: true` וכותבת קובץ ש-Word אינו פותח.
 *
 * ולכן `assertWordLegal` הוא שער נפרד על ה-XML שיצא: הוא בודק את מה שהתקן
 * דורש (`w:val` נדרש ב-`CT_Border`, `w:sz` שלם 2–96, `w:space` שלם 0–31,
 * `ST_HexColor` שש ספרות או `auto`, `w:fmt` מתוך `ST_NumberFormat`), ולא את
 * מה ששלחנו. מוטציה במודול נתפסת שם, גם כשהקבלה חוזרת מוצלחת.
 */
import { describe, expect, it } from 'vitest';
import {
  COLUMN_GAP_TWIPS,
  HEADER_DISTANCE_DEFAULT_CM,
  HEADER_DISTANCE_MAX_CM,
  LINE_COUNT_BY_MAX,
  LINE_NUMBER_CHOICES,
  MARGIN_PRESETS,
  NUMBER_START_MAX,
  PAGE_BORDER_PRESETS,
  PAGE_NUMBER_FORMATS,
  PAPER_SIZES,
  TWIPS_PER_INCH,
  VERTICAL_ALIGNS,
  applyColumns,
  applyHeaderDistance,
  applyLineNumbering,
  applyMarginPreset,
  applyOrientation,
  applyPageBorders,
  applyPageMargins,
  applyPageNumbering,
  applyPaperSize,
  readPageMargins,
  applyVerticalAlign,
  cmToInches,
  normalizeHeaderDistanceCm,
  normalizePageNumberStart,
  readPageLayoutState,
  type PageSetupDocumentApi,
  type PageSetupHost,
} from '../../src/engine/page-setup';

/** אותה בדיקה שהמנוע עושה על כל שדה מידה. */
function assertMeasure(value: unknown, field: string): void {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) {
    throw new Error(`${field} must be a non-negative number.`);
  }
}

function assertAnyOf(input: Record<string, unknown>, fields: string[], op: string): void {
  if (!fields.some((field) => input[field] !== undefined)) {
    throw new Error(`${op} requires at least one field.`);
  }
}

/** twips מהערך שנשלח, בדיוק כפי שהמנוע כותב אותם ל-XML. */
function toTwips(inches: number): number {
  return Math.round(inches * TWIPS_PER_INCH);
}

interface BorderXml {
  /** `w:val`. `undefined` = התכונה לא נכתבה בכלל — מה שקורה על `style: ''`. */
  val?: string;
  sz?: number;
  space?: number;
  color?: string;
}

interface PgBordersXml {
  display?: string;
  offsetFrom?: string;
  top?: BorderXml;
  right?: BorderXml;
  bottom?: BorderXml;
  left?: BorderXml;
}

interface SectionXml {
  sectionId: string;
  pgMar: Record<string, number>;
  pgSz: { w?: number; h?: number; orient?: string; code?: string };
  cols: { num?: number; space?: number; equalWidth?: boolean };
  /**
   * `w:header`/`w:footer` — הם תכונות של `w:pgMar` ב-OOXML, ומוחזקים כאן
   * בנפרד ממנו מפני ש-`setPageMargins` ו-`setHeaderFooterMargins` הן שתי
   * פעולות שונות, והבדיקות של השוליים משוות את `pgMar` כולו.
   */
  hfMar: { header?: number; footer?: number };
  /** `<w:lnNumType>`; `null` = האלמנט אינו במסמך. */
  lnNumType: { countBy?: number; start?: number; distance?: number; restart?: string } | null;
  pgNumType: { start?: number; fmt?: string } | null;
  pgBorders: PgBordersXml | null;
  vAlign?: string;
}

interface FakeOptions {
  /** מזהי המקטעים במסמך. */
  sectionIds?: string[];
  /** כיוון המקטע כפי ש-`sections.list()` מדווח אותו. נמדד: `'rtl'` במסמך עברי. */
  sectionDirection?: 'rtl' | 'ltr';
  /** מידות התחלה לכל מקטע, ב-twips. ברירת המחדל: Letter לאורך, כמו המסמך הריק. */
  startWidth?: number;
  startHeight?: number;
  /** קבלה חלופית — לבדיקת כשל, NO_OP, או הבטחה. */
  receipt?: () => unknown;
  /**
   * מה ש-`pageMetrics` מדווח, בפיקסלים — כלומר מה שהמנוע **צייר**. כותרת
   * עליונה מרימה את `marginTopPx` מעל מה שכתוב במסמך; ראו readEffectiveMargins.
   */
  effectiveMarginsPx?: { top?: number; bottom?: number };
  /** `pageMetrics` שזורק — גרסת מנוע שהחוזה שלה שונה. */
  metricsThrows?: boolean;
  /** להסיר פעולה מהחוזה, כדי לדמות גרסה שאינה מכירה אותה. */
  omit?: Array<
    | 'list'
    | 'setPageMargins'
    | 'setPageSetup'
    | 'setColumns'
    | 'setLineNumbering'
    | 'setVerticalAlign'
    | 'setHeaderFooterMargins'
    | 'setPageNumbering'
    | 'setPageBorders'
    | 'clearPageBorders'
  >;
  /** `list` שזורקת. */
  throwOnList?: boolean;
  /**
   * מה שהמקטע **כבר** נושא, כפי שהמנוע מחזיר אותו: אינצ'ים ל-`distance`,
   * מספרים שלמים ל-`countBy`/`start`. זה מה שהופך את „מספרי שורות” לבדיקה
   * אמיתית — בלעדיו אין מה לשמר, וכל מימוש עובר.
   */
  lineNumbering?: {
    enabled?: boolean;
    countBy?: number;
    start?: number;
    distance?: number;
    restart?: string;
  };
  headerFooterMargins?: { header?: number; footer?: number };
  pageNumbering?: { start?: number; format?: string };
}

/** `ST_Border` — הסגנונות ש-Word מכיר בגבול עמוד. */
const BORDER_STYLES = new Set([
  'nil', 'none', 'single', 'thick', 'double', 'dotted', 'dashed', 'dotDash',
  'dotDotDash', 'triple', 'thinThickSmallGap', 'thickThinSmallGap',
  'thinThickThinSmallGap', 'thinThickMediumGap', 'thickThinMediumGap',
  'thinThickThinMediumGap', 'thinThickLargeGap', 'thickThinLargeGap',
  'thinThickThinLargeGap', 'wave', 'doubleWave', 'dashSmallGap',
  'dashDotStroked', 'threeDEmboss', 'threeDEngrave', 'outset', 'inset',
]);

/** `ST_NumberFormat` בחלק שרלוונטי ל-`w:pgNumType/@w:fmt`, כולל המספור העברי. */
const NUMBER_FORMATS = new Set([
  'decimal', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter',
  'numberInDash', 'hebrew1', 'hebrew2', 'ordinal', 'cardinalText',
]);

const LINE_RESTARTS = new Set(['continuous', 'newPage', 'newSection']);
const VERTICAL_JC = new Set(['top', 'center', 'both', 'bottom']);

function attr(name: string, value: unknown): string {
  return value === undefined ? '' : ` w:${name}="${String(value)}"`;
}

function borderXml(tag: string, border: BorderXml | undefined): string {
  if (!border) return '';
  return `<w:${tag}${attr('val', border.val)}${attr('sz', border.sz)}${attr('space', border.space)}${attr('color', border.color)}/>`;
}

/**
 * ה-`sectPr` שהיה נכתב ל-`word/document.xml`, בסדר האלמנטים של `CT_SectPr`.
 *
 * הרינדור הוא מה שמאפשר לבדוק את **הקובץ** ולא את הקריאה. סדר האלמנטים אינו
 * קוסמטיקה: `CT_SectPr` הוא sequence, ו-Word דוחה קובץ שבו `w:vAlign` בא
 * לפני `w:pgSz`.
 */
function sectPrXml(section: SectionXml): string {
  const { pgSz, pgMar, hfMar, cols, lnNumType, pgNumType, pgBorders, vAlign } = section;
  const borders = pgBorders
    ? `<w:pgBorders${attr('display', pgBorders.display)}${attr('offsetFrom', pgBorders.offsetFrom)}>` +
      borderXml('top', pgBorders.top) +
      borderXml('right', pgBorders.right) +
      borderXml('bottom', pgBorders.bottom) +
      borderXml('left', pgBorders.left) +
      '</w:pgBorders>'
    : '';
  return (
    '<w:sectPr>' +
    `<w:pgSz${attr('w', pgSz.w)}${attr('h', pgSz.h)}${attr('orient', pgSz.orient)}${attr('code', pgSz.code)}/>` +
    `<w:pgMar${attr('top', pgMar.top)}${attr('right', pgMar.right)}${attr('bottom', pgMar.bottom)}${attr('left', pgMar.left)}${attr('header', hfMar.header)}${attr('footer', hfMar.footer)}/>` +
    borders +
    (lnNumType
      ? `<w:lnNumType${attr('countBy', lnNumType.countBy)}${attr('start', lnNumType.start)}${attr('distance', lnNumType.distance)}${attr('restart', lnNumType.restart)}/>`
      : '') +
    (pgNumType ? `<w:pgNumType${attr('start', pgNumType.start)}${attr('fmt', pgNumType.fmt)}/>` : '') +
    `<w:cols${attr('num', cols.num)}${attr('space', cols.space)}/>` +
    (vAlign === undefined ? '' : `<w:vAlign w:val="${vAlign}"/>`) +
    '</w:sectPr>'
  );
}

/* ------------------------------------------------------------------ */
/* המספרים של התקן — קשיחים, ובכוונה                                   */
/* ------------------------------------------------------------------ */

/**
 * **כאן, ורק כאן, מספר קשוח הוא הדבר הנכון.**
 *
 * השער מייצג את ECMA-376, לא את המימוש שלנו. כשהוא ייבא את `NUMBER_START_MAX`
 * מ-page-setup.ts הוא לא שוטר על המודול אלא שיקף אותו: מוטציה שהחליפה שם את
 * התקרה ל-מיליארד הזיזה גם את התקרה שהשער בודק, והשער עבר בירוק על
 * `w:pgNumType w:start="1000000000"` — כלומר על קובץ ש-Word אינו פותח. שער
 * שהקבועים שלו זזים עם הקוד שהוא בודק אינו שער.
 *
 * ולכן המספרים כתובים כאן במלואם, עם המקור שלהם:
 *   • 32767 — התקרה של Word ל-`w:pgNumType/@w:start` ול-`w:lnNumType/@w:start`.
 *   • 100 — התקרה של „ספור לפי” (`w:lnNumType/@w:countBy`).
 *   • 31680 twips — 22 אינץ', התקרה של Word למרחק הכותרת מקצה הדף ולמרחק
 *     מספרי השורות מהטקסט. גם `TWIPS_PER_INCH` אינו מיובא לכאן מאותה סיבה.
 * (2..96 ל-`w:sz`, 0..31 ל-`w:space` וששת התווים של `ST_HexColor` היו קשיחים
 * כאן מהיום הראשון — אלה שלושה מספרים שהצטרפו אליהם, לא חריג חדש.)
 */
const STD_NUMBER_START_MAX = 32767;
const STD_LINE_COUNT_BY_MAX = 100;
const STD_DISTANCE_MAX_TWIPS = 31680;

/**
 * שער התקן על ה-XML שיצא.
 *
 * בודק את מה ש-ECMA-376 דורש, ולא את מה ששלחנו: זו הבדיקה שתופסת מוטציה
 * במודול גם כשהמנוע החזיר `success: true`. כל כשל כאן הוא קובץ ש-Word יסרב
 * לפתוח או יתקן בשקט.
 */
function assertWordLegal(section: SectionXml): void {
  const { hfMar, lnNumType, pgNumType, pgBorders, vAlign } = section;

  for (const field of ['header', 'footer'] as const) {
    const value = hfMar[field];
    if (value === undefined) continue;
    if (!Number.isInteger(value) || value < 0 || value > STD_DISTANCE_MAX_TWIPS) {
      throw new Error(`w:${field}="${value}" אינו מרחק חוקי`);
    }
  }

  if (lnNumType) {
    const { countBy, start, distance, restart } = lnNumType;
    if (countBy !== undefined && (!Number.isInteger(countBy) || countBy < 1 || countBy > STD_LINE_COUNT_BY_MAX)) {
      throw new Error(`w:countBy="${countBy}" מחוץ לטווח`);
    }
    if (start !== undefined && (!Number.isInteger(start) || start < 1 || start > STD_NUMBER_START_MAX)) {
      throw new Error(`w:start="${start}" מחוץ לטווח`);
    }
    if (distance !== undefined && (!Number.isInteger(distance) || distance < 0 || distance > STD_DISTANCE_MAX_TWIPS)) {
      throw new Error(`w:distance="${distance}" מחוץ לטווח`);
    }
    if (restart !== undefined && !LINE_RESTARTS.has(restart)) {
      throw new Error(`w:restart="${restart}" אינו ST_LineNumberRestart`);
    }
    // אותה בדיקה בדיוק שיש ל-`<w:pgNumType/>` ריק, ומאותה סיבה: נמדד
    // ש-`{ enabled: true }` לבד מייצר `<w:lnNumType/>` בלי אף תכונה —
    // אלמנט שאינו אומר דבר, ואי-סימטריה בשער היא חור בשער.
    if (
      countBy === undefined &&
      start === undefined &&
      distance === undefined &&
      restart === undefined
    ) {
      throw new Error('<w:lnNumType/> ריק — נכתב אלמנט שאינו אומר דבר');
    }
  }

  if (pgNumType) {
    if (pgNumType.fmt !== undefined && !NUMBER_FORMATS.has(pgNumType.fmt)) {
      throw new Error(`w:fmt="${pgNumType.fmt}" אינו ST_NumberFormat`);
    }
    if (
      pgNumType.start !== undefined &&
      (!Number.isInteger(pgNumType.start) ||
        pgNumType.start < 1 ||
        pgNumType.start > STD_NUMBER_START_MAX)
    ) {
      throw new Error(`w:pgNumType/@w:start="${pgNumType.start}" מחוץ לטווח`);
    }
    if (pgNumType.start === undefined && pgNumType.fmt === undefined) {
      throw new Error('<w:pgNumType/> ריק — נכתב אלמנט שאינו אומר דבר');
    }
  }

  if (pgBorders) {
    if (pgBorders.display !== undefined && !['allPages', 'firstPage', 'notFirstPage'].includes(pgBorders.display)) {
      throw new Error(`w:display="${pgBorders.display}" אינו חוקי`);
    }
    if (pgBorders.offsetFrom !== undefined && !['page', 'text'].includes(pgBorders.offsetFrom)) {
      throw new Error(`w:offsetFrom="${pgBorders.offsetFrom}" אינו חוקי`);
    }
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const border = pgBorders[side];
      if (!border) continue;
      // `w:val` הוא **תכונה נדרשת** ב-CT_Border. זה בדיוק מה ש-`style: ''`
      // מייצר במנוע האמיתי, ולכן זו הבדיקה הראשונה.
      if (border.val === undefined) throw new Error(`<w:${side}/> בלי w:val`);
      if (!BORDER_STYLES.has(border.val)) throw new Error(`w:val="${border.val}" אינו ST_Border`);
      if (border.sz !== undefined && (!Number.isInteger(border.sz) || border.sz < 2 || border.sz > 96)) {
        throw new Error(`w:sz="${border.sz}" מחוץ ל-2..96`);
      }
      if (border.space !== undefined && (!Number.isInteger(border.space) || border.space < 0 || border.space > 31)) {
        throw new Error(`w:space="${border.space}" מחוץ ל-0..31`);
      }
      if (border.color !== undefined && !/^([0-9A-Fa-f]{6}|auto)$/.test(border.color)) {
        throw new Error(`w:color="${border.color}" אינו ST_HexColor`);
      }
    }
  }

  if (vAlign !== undefined && !VERTICAL_JC.has(vAlign)) {
    throw new Error(`w:vAlign="${vAlign}" אינו ST_VerticalJc`);
  }
}

function fakeEngine(options: FakeOptions = {}) {
  const ids = options.sectionIds ?? ['s0'];
  const omit = new Set(options.omit ?? []);
  const calls: Array<{ op: string; input: Record<string, unknown> }> = [];

  const xml = new Map<string, SectionXml>(
    ids.map((sectionId) => [
      sectionId,
      {
        sectionId,
        pgMar: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        // 720 twips = חצי אינץ', וזה מה שהמסמך הריק של המנוע נושא.
        hfMar: { header: 720, footer: 720 },
        pgSz: { w: options.startWidth ?? 12240, h: options.startHeight ?? 15840 },
        cols: { space: 720 },
        lnNumType: null,
        pgNumType: null,
        pgBorders: null,
      },
    ]),
  );

  function sectionOf(input: Record<string, unknown>): SectionXml {
    const target = input.target as { kind?: string; sectionId?: string } | undefined;
    if (!target || target.kind !== 'section' || typeof target.sectionId !== 'string') {
      throw new Error('target must be a section address.');
    }
    const found = xml.get(target.sectionId);
    if (!found) throw new Error('INVALID_TARGET');
    return found;
  }

  const receipt = options.receipt;

  const sections: NonNullable<PageSetupDocumentApi['sections']> = {};

  if (!omit.has('list')) {
    sections.list = () => {
      if (options.throwOnList) throw new Error('boom');
      return Promise.resolve({
        items: ids.map((sectionId, index) => {
          const current = xml.get(sectionId)!;
          return {
            address: { kind: 'section', sectionId },
            index,
            // המנוע מחזיר את המידות ב-Document API; היחס הוא מה שמשמש לזיהוי „לרוחב”.
            pageSetup: {
              width: (current.pgSz.w ?? 0) / TWIPS_PER_INCH,
              height: (current.pgSz.h ?? 0) / TWIPS_PER_INCH,
            },
            // ההגדרות שהמנוע מחזיר: אינצ'ים ל-`distance`, ולא twips.
            // `options` ולא `current` — זה מה שהמסמך נפתח איתו, כלומר מה
            // שהגיע מ-Word ומה שהמודול אמור לשמר.
            lineNumbering: options.lineNumbering,
            headerFooterMargins: options.headerFooterMargins ?? {
              header: (current.hfMar.header ?? 0) / TWIPS_PER_INCH,
              footer: (current.hfMar.footer ?? 0) / TWIPS_PER_INCH,
            },
            pageNumbering: options.pageNumbering,
            // נמדד על המנוע: השוליים חוזרים באינצ'ים, ולצדם `sectionDirection`.
            margins: {
              top: (current.pgMar.top ?? 0) / TWIPS_PER_INCH,
              right: (current.pgMar.right ?? 0) / TWIPS_PER_INCH,
              bottom: (current.pgMar.bottom ?? 0) / TWIPS_PER_INCH,
              left: (current.pgMar.left ?? 0) / TWIPS_PER_INCH,
              gutter: 0,
            },
            sectionDirection: options.sectionDirection ?? 'rtl',
          };
        }),
      });
    };
  }

  if (!omit.has('setPageMargins')) {
    sections.setPageMargins = (input) => {
      const raw = input as unknown as Record<string, unknown>;
      calls.push({ op: 'sections.setPageMargins', input: raw });
      const section = sectionOf(raw);
      assertAnyOf(raw, ['top', 'right', 'bottom', 'left', 'gutter'], 'sections.setPageMargins');
      for (const field of ['top', 'right', 'bottom', 'left', 'gutter']) {
        if (raw[field] !== undefined) {
          assertMeasure(raw[field], `sections.setPageMargins.${field}`);
          section.pgMar[field] = toTwips(raw[field] as number);
        }
      }
      return (receipt?.() ?? { success: true, section: raw.target }) as never;
    };
  }

  if (!omit.has('setPageSetup')) {
    sections.setPageSetup = (input) => {
      const raw = input as unknown as Record<string, unknown>;
      calls.push({ op: 'sections.setPageSetup', input: raw });
      const section = sectionOf(raw);
      assertAnyOf(raw, ['width', 'height', 'orientation', 'paperSize'], 'sections.setPageSetup');
      if (raw.width !== undefined) {
        assertMeasure(raw.width, 'sections.setPageSetup.width');
        section.pgSz.w = toTwips(raw.width as number);
      }
      if (raw.height !== undefined) {
        assertMeasure(raw.height, 'sections.setPageSetup.height');
        section.pgSz.h = toTwips(raw.height as number);
      }
      if (raw.paperSize !== undefined) {
        if (typeof raw.paperSize !== 'string' || raw.paperSize.trim() === '') {
          throw new Error('sections.setPageSetup.paperSize must be a non-empty string.');
        }
        section.pgSz.code = raw.paperSize;
      }
      if (raw.orientation !== undefined) {
        if (raw.orientation !== 'portrait' && raw.orientation !== 'landscape') {
          throw new Error('sections.setPageSetup.orientation must be portrait or landscape.');
        }
        section.pgSz.orient = raw.orientation;
        // ההחלפה שהמנוע עושה בעצמו כשהיחס אינו מתאים לכיוון המבוקש.
        const { w, h } = section.pgSz;
        if (typeof w === 'number' && typeof h === 'number') {
          if ((raw.orientation === 'landscape' && w <= h) || (raw.orientation === 'portrait' && w > h)) {
            section.pgSz.w = h;
            section.pgSz.h = w;
          }
        }
      }
      return (receipt?.() ?? { success: true, section: raw.target }) as never;
    };
  }

  if (!omit.has('setColumns')) {
    sections.setColumns = (input) => {
      const raw = input as unknown as Record<string, unknown>;
      calls.push({ op: 'sections.setColumns', input: raw });
      const section = sectionOf(raw);
      assertAnyOf(raw, ['count', 'gap', 'equalWidth'], 'sections.setColumns');
      if (raw.count !== undefined) {
        if (!Number.isInteger(raw.count) || (raw.count as number) <= 0) {
          throw new Error('sections.setColumns.count must be a positive integer.');
        }
        section.cols.num = raw.count as number;
      }
      if (raw.gap !== undefined) {
        assertMeasure(raw.gap, 'sections.setColumns.gap');
        section.cols.space = toTwips(raw.gap as number);
      }
      if (raw.equalWidth !== undefined) section.cols.equalWidth = raw.equalWidth as boolean;
      return (receipt?.() ?? { success: true, section: raw.target }) as never;
    };
  }

  /**
   * חמש הפעולות של גל 10. כל אחת מהן מדגמת **שתי** התנהגויות שנמדדו:
   * הוולידציה שהמנוע כן עושה (זורק, ולא מחזיר קבלה), והבליעה שהוא עושה על
   * כל מה שעבר אותה. ראו הערת הפתיחה.
   */
  if (!omit.has('setLineNumbering')) {
    sections.setLineNumbering = (input) => {
      const raw = input as unknown as Record<string, unknown>;
      calls.push({ op: 'sections.setLineNumbering', input: raw });
      const section = sectionOf(raw);
      if (typeof raw.enabled !== 'boolean') {
        throw new Error('sections.setLineNumbering.enabled must be a boolean.');
      }
      for (const field of ['countBy', 'start'] as const) {
        if (raw[field] === undefined) continue;
        if (!Number.isInteger(raw[field]) || (raw[field] as number) <= 0) {
          throw new Error(`sections.setLineNumbering.${field} must be a positive integer.`);
        }
      }
      if (raw.distance !== undefined) assertMeasure(raw.distance, 'sections.setLineNumbering.distance');
      if (raw.restart !== undefined && !LINE_RESTARTS.has(raw.restart as string)) {
        throw new Error(
          'sections.setLineNumbering.restart must be one of: continuous, newPage, newSection.',
        );
      }
      // כיבוי מוריד את האלמנט כולו; הדלקה **מחליפה** אותו ואינה מטליאה.
      section.lnNumType = raw.enabled
        ? {
            countBy: raw.countBy as number | undefined,
            start: raw.start as number | undefined,
            distance: raw.distance === undefined ? undefined : toTwips(raw.distance as number),
            restart: raw.restart as string | undefined,
          }
        : null;
      return (receipt?.() ?? { success: true, section: raw.target }) as never;
    };
  }

  if (!omit.has('setVerticalAlign')) {
    sections.setVerticalAlign = (input) => {
      const raw = input as unknown as Record<string, unknown>;
      calls.push({ op: 'sections.setVerticalAlign', input: raw });
      const section = sectionOf(raw);
      if (!VERTICAL_JC.has(raw.value as string)) {
        throw new Error('sections.setVerticalAlign.value must be one of: top, center, bottom, both.');
      }
      section.vAlign = raw.value as string;
      return (receipt?.() ?? { success: true, section: raw.target }) as never;
    };
  }

  if (!omit.has('setHeaderFooterMargins')) {
    sections.setHeaderFooterMargins = (input) => {
      const raw = input as unknown as Record<string, unknown>;
      calls.push({ op: 'sections.setHeaderFooterMargins', input: raw });
      const section = sectionOf(raw);
      assertAnyOf(raw, ['header', 'footer'], 'sections.setHeaderFooterMargins');
      for (const field of ['header', 'footer'] as const) {
        if (raw[field] === undefined) continue;
        assertMeasure(raw[field], `sections.setHeaderFooterMargins.${field}`);
        // **אין תקרה.** `99` נכתב 142560, וזה מה שנמדד.
        section.hfMar[field] = toTwips(raw[field] as number);
      }
      return (receipt?.() ?? { success: true, section: raw.target }) as never;
    };
  }

  if (!omit.has('setPageNumbering')) {
    sections.setPageNumbering = (input) => {
      const raw = input as unknown as Record<string, unknown>;
      calls.push({ op: 'sections.setPageNumbering', input: raw });
      const section = sectionOf(raw);
      assertAnyOf(
        raw,
        ['start', 'format', 'chapterStyle', 'chapterSeparator'],
        'sections.setPageNumbering',
      );
      if (raw.start !== undefined && (!Number.isInteger(raw.start) || (raw.start as number) <= 0)) {
        throw new Error('sections.setPageNumbering.start must be a positive integer.');
      }
      if (
        raw.format !== undefined &&
        !PAGE_NUMBER_FORMATS.some((item) => item.id === raw.format)
      ) {
        throw new Error(
          'sections.setPageNumbering.format must be one of: decimal, lowerLetter, upperLetter, lowerRoman, upperRoman, numberInDash.',
        );
      }
      // `chapterStyle` ו-`chapterSeparator` **נבלעים**: הם מתקבלים, מרוצים
      // את דרישת „לפחות שדה אחד”, ואינם נכתבים. זה מה שהופך
      // `{chapterStyle:1}` ל-`<w:pgNumType/>` ריק — ובדיוק מה ש-
      // `assertWordLegal` פוסל.
      section.pgNumType = {
        start: raw.start as number | undefined,
        fmt: raw.format as string | undefined,
      };
      return (receipt?.() ?? { success: true, section: raw.target }) as never;
    };
  }

  if (!omit.has('setPageBorders')) {
    sections.setPageBorders = (input) => {
      const raw = input as unknown as Record<string, unknown>;
      calls.push({ op: 'sections.setPageBorders', input: raw });
      const section = sectionOf(raw);
      const borders = (raw.borders ?? {}) as Record<string, unknown>;
      assertAnyOf(
        borders,
        ['display', 'offsetFrom', 'zOrder', 'top', 'right', 'bottom', 'left'],
        'sections.setPageBorders.borders',
      );
      if (
        borders.display !== undefined &&
        !['allPages', 'firstPage', 'notFirstPage'].includes(borders.display as string)
      ) {
        throw new Error(
          'sections.setPageBorders.borders.display must be one of: allPages, firstPage, notFirstPage.',
        );
      }
      if (borders.offsetFrom !== undefined && !['page', 'text'].includes(borders.offsetFrom as string)) {
        throw new Error('sections.setPageBorders.borders.offsetFrom must be one of: page, text.');
      }
      const rendered: PgBordersXml = {
        display: borders.display as string | undefined,
        offsetFrom: borders.offsetFrom as string | undefined,
      };
      for (const side of ['top', 'right', 'bottom', 'left'] as const) {
        const spec = borders[side] as Record<string, unknown> | undefined;
        if (!spec) continue;
        if (spec.size !== undefined) {
          assertMeasure(spec.size, `sections.setPageBorders.borders.${side}.size`);
        }
        // **`style` אינו מאומת בכלל**, ומחרוזת ריקה משמיטה את `w:val`.
        // `size`, `space` ו-`color` נכתבים גולמית — כולל שבר, 999 ו-'#FF0000'.
        rendered[side] = {
          val: spec.style === '' || spec.style === undefined ? undefined : String(spec.style),
          sz: spec.size as number | undefined,
          space: spec.space as number | undefined,
          color: spec.color as string | undefined,
        };
      }
      section.pgBorders = rendered;
      return (receipt?.() ?? { success: true, section: raw.target }) as never;
    };
  }

  if (!omit.has('clearPageBorders')) {
    sections.clearPageBorders = (input) => {
      const raw = input as unknown as Record<string, unknown>;
      calls.push({ op: 'sections.clearPageBorders', input: raw });
      const section = sectionOf(raw);
      // נמדד: הסרה על מקטע שאין בו גבול מוחזרת NO_OP, ולא כשל.
      if (section.pgBorders === null) {
        return (receipt?.() ?? {
          success: false,
          failure: { code: 'NO_OP', message: 'sections.clearPageBorders did not produce a section change.' },
        }) as never;
      }
      section.pgBorders = null;
      return (receipt?.() ?? { success: true, section: raw.target }) as never;
    };
  }

  const pageMetrics =
    options.effectiveMarginsPx || options.metricsThrows
      ? {
          getSnapshot: () => {
            if (options.metricsThrows) throw new Error('boom');
            const first = xml.get(ids[0]!)!;
            return {
              pages: [
                {
                  base: {
                    widthPx: (first.pgSz.w ?? 0) / 15,
                    heightPx: (first.pgSz.h ?? 0) / 15,
                    marginTopPx: options.effectiveMarginsPx?.top ?? (first.pgMar.top ?? 0) / 15,
                    marginBottomPx:
                      options.effectiveMarginsPx?.bottom ?? (first.pgMar.bottom ?? 0) / 15,
                  },
                },
              ],
            };
          },
        }
      : undefined;

  const host: PageSetupHost = { activeEditor: { doc: { sections }, pageMetrics } };
  return { host, calls, xml };
}

describe('applyMarginPreset', () => {
  it('„רגיל” כותב 1440 twips בכל ארבעת הצדדים', async () => {
    // 1440 twips = אינץ' = 2.54 ס"מ, וזה בדיוק מה שנמדד ב-w:pgMar של המסמך
    // הריק של המנוע.
    const { host, xml } = fakeEngine();

    expect(await applyMarginPreset(host, 'normal')).toEqual({ ok: true });
    expect(xml.get('s0')!.pgMar).toEqual({ top: 1440, right: 1440, bottom: 1440, left: 1440 });
  });

  it('„צר” כותב 720, ו„רחב” 1440 לאורך ו-2880 בצדדים', async () => {
    const narrow = fakeEngine();
    expect(await applyMarginPreset(narrow.host, 'narrow')).toEqual({ ok: true });
    expect(narrow.xml.get('s0')!.pgMar).toEqual({ top: 720, right: 720, bottom: 720, left: 720 });

    const wide = fakeEngine();
    expect(await applyMarginPreset(wide.host, 'wide')).toEqual({ ok: true });
    expect(wide.xml.get('s0')!.pgMar).toEqual({ top: 1440, right: 2880, bottom: 1440, left: 2880 });
  });

  it('שולחת אינצ\'ים ולא twips', async () => {
    // זו הטעות שהבדיקה הזאת קיימת בשבילה: שליחת 1440 הייתה מייצרת
    // w:top="2073600" — שולי דף בגובה 36 מטר, בלי שום שגיאה מהמנוע.
    const { host, calls } = fakeEngine();

    await applyMarginPreset(host, 'normal');

    expect(calls[0]!.input).toEqual({
      target: { kind: 'section', sectionId: 's0' },
      top: 1,
      right: 1,
      bottom: 1,
      left: 1,
    });
  });

  it('כל preset בטבלה עובר את הוולידציה של המנוע', async () => {
    for (const preset of MARGIN_PRESETS) {
      const { host } = fakeEngine();
      expect(await applyMarginPreset(host, preset.id), preset.id).toEqual({ ok: true });
    }
  });

  it('preset שאינו קיים אינו נוגע במנוע', async () => {
    const { host, calls } = fakeEngine();

    const outcome = await applyMarginPreset(host, 'huge');

    expect(outcome.ok).toBe(false);
    expect(calls).toEqual([]);
  });
});

describe('readPageMargins', () => {
  it('מחזירה רוחב עמוד ושוליים ב-twips, וכיוון מהמקטע', async () => {
    // A4 לאורך, 2.54 ס"מ מכל צד — המסמך שהתוסף פותח.
    const { host } = fakeEngine({ startWidth: 11906, startHeight: 16838 });

    expect(await readPageMargins(host)).toEqual({
      pageWidthTwips: 11906,
      pageHeightTwips: 16838,
      leftTwips: 1440,
      rightTwips: 1440,
      topTwips: 1440,
      bottomTwips: 1440,
      effectiveTopTwips: 1440,
      effectiveBottomTwips: 1440,
      direction: 'rtl',
    });
  });

  it('בלי `pageMetrics` הערכים האפקטיביים הם מה שכתוב במסמך', async () => {
    // גרסת מנוע שאינה חושפת מדידות: הסרגל חייב להמשיך לעבוד כמו קודם.
    const { host } = fakeEngine();
    expect(await readPageMargins(host)).toMatchObject({
      topTwips: 1440,
      effectiveTopTwips: 1440,
    });
  });

  it('כותרת עליונה מרימה את השוליים האפקטיביים, והמסמך נשאר כשהיה', async () => {
    // נמדד: כותרת ריקה במרחק חצי אינץ' מרימה את ראש הטקסט ל-66.4px.
    const { host } = fakeEngine({ effectiveMarginsPx: { top: 66.4 } });
    await applyPageMargins(host, { topTwips: 720 });

    const state = await readPageMargins(host);
    expect(state?.topTwips).toBe(720); // מה שכתוב במסמך
    expect(state?.effectiveTopTwips).toBe(996); // 66.4px * 15 = מה שצויר
  });

  it('מדידה נמוכה מהמסמך אינה מורידה — היא יכולה להיות רק תצלום ישן', async () => {
    const { host } = fakeEngine({ effectiveMarginsPx: { top: 10, bottom: 10 } });
    const state = await readPageMargins(host);
    expect(state?.effectiveTopTwips).toBe(1440);
    expect(state?.effectiveBottomTwips).toBe(1440);
  });

  it('`pageMetrics` שזורק אינו מפיל את הקריאה', async () => {
    const { host } = fakeEngine({ metricsThrows: true });
    expect(await readPageMargins(host)).toMatchObject({ effectiveTopTwips: 1440 });
  });

  it('מסמך לועזי מדווח `ltr`, ולא כיוון הממשק', async () => {
    const { host } = fakeEngine({ sectionDirection: 'ltr' });
    expect((await readPageMargins(host))?.direction).toBe('ltr');
  });

  it('שוליים שהשתנו נקראים חזרה — הסרגל והגלריה קוראים את אותו מספר', async () => {
    const { host } = fakeEngine({ startWidth: 11906 });
    await applyMarginPreset(host, 'narrow');

    expect(await readPageMargins(host)).toMatchObject({ leftTwips: 720, rightTwips: 720 });
  });

  it('מסמך שעדיין נטען מוחזר כ-`null`, ולא כשגיאה', async () => {
    for (const host of [null, undefined, {}, { activeEditor: null }, { activeEditor: { doc: null } }]) {
      expect(await readPageMargins(host as never)).toBeNull();
    }
  });

  it('`list` שזורקת אינה מפילה את הסרגל', async () => {
    const { host } = fakeEngine({ throwOnList: true });
    expect(await readPageMargins(host)).toBeNull();
  });

  it('מקטע בלי מידות אינו מייצר סרגל של חצי אמת', async () => {
    // רוחב בלי שוליים היה מצייר סרגל שנראה נכון עם אזור טקסט מומצא.
    const sections = {
      list: () => Promise.resolve({ items: [{ pageSetup: { width: 8.5 } }] }),
    };
    expect(await readPageMargins({ activeEditor: { doc: { sections } } } as never)).toBeNull();
  });
});

/**
 * הצוק שהחסם הזה מגן עליו נמדד על ה-`dist` הארוז: שוליים שאינם משאירים גובה
 * טקסט חיובי מפילים את הפריסה לאפס עמודים, וכל `setPageMargins` שאחריו
 * מחזיר `success: true` בלי שהמסמך חוזר. כלומר בלי החסם אין דרך חזרה בתוך
 * ההפעלה. ראו docs/engine-gaps.md.
 */
describe('applyPaperSize — מקום לטקסט', () => {
  it('דף קטן יותר שהשוליים הקיימים לא נכנסים בו נדחה', async () => {
    // A4 גבוה מ-Letter בכמעט אינץ': שוליים שהיו חוקיים ב-A4 חוצים את הצוק.
    const { host, calls, xml } = fakeEngine({ startWidth: 11906, startHeight: 16838 });
    xml.get('s0')!.pgMar.top = 16838 - 1440 - 720;

    const outcome = await applyPaperSize(host, 'letter');

    expect(outcome).toMatchObject({
      ok: false,
      message: expect.stringContaining('גדולים מדי לגובה של Letter'),
    });
    expect(calls).toEqual([]);
  });

  it('מסמך רגיל עובר גודל דף כרגיל', async () => {
    const { host } = fakeEngine({ startWidth: 11906, startHeight: 16838 });
    expect(await applyPaperSize(host, 'letter')).toEqual({ ok: true });
  });
});

describe('applyPageMargins — מקום לטקסט', () => {
  /** Letter לאורך: 12240 × 15840 twips. */
  it('שוליים שחונקים את גובה העמוד נדחים, והמנוע אינו נקרא', async () => {
    const { host, calls } = fakeEngine();

    const outcome = await applyPageMargins(host, { topTwips: 15840 - 1440 - 719 });

    expect(outcome).toMatchObject({
      ok: false,
      message: expect.stringContaining('לא יישאר מקום לטקסט לגובה העמוד'),
    });
    expect(calls).toEqual([]);
  });

  it('שוליים שחונקים את רוחב העמוד נדחים גם הם', async () => {
    const { host, calls } = fakeEngine();

    const outcome = await applyPageMargins(host, { leftTwips: 12240 - 1440 - 719 });

    expect(outcome).toMatchObject({
      ok: false,
      message: expect.stringContaining('לא יישאר מקום לטקסט לרוחב העמוד'),
    });
    expect(calls).toEqual([]);
  });

  it('בדיוק על החסם — עובר', async () => {
    const { host } = fakeEngine();

    expect(await applyPageMargins(host, { topTwips: 15840 - 1440 - 720 })).toEqual({ ok: true });
  });

  it('הצד שלא נשלח נלקח מהמסמך — הצוק הוא בסכום', async () => {
    const { host } = fakeEngine();
    // 7000 לבד חוקי; 7000 + 7000 כבר לא, וזה מה שקורה בגרירה שנייה.
    expect(await applyPageMargins(host, { topTwips: 7000 })).toEqual({ ok: true });

    const outcome = await applyPageMargins(host, { bottomTwips: 9000 });
    expect(outcome.ok).toBe(false);
  });

  it('מסמך שכבר חנוק אינו ננעל — שינוי שמשפר מותר', async () => {
    // קובץ Word יכול להגיע כך. סירוב גורף היה מונע דווקא את התיקון.
    const { host, xml } = fakeEngine();
    xml.get('s0')!.pgMar.top = 15500;

    const outcome = await applyPageMargins(host, { topTwips: 15000 });

    expect(outcome.ok).toBe(true);
    expect(xml.get('s0')!.pgMar.top).toBe(15000);
  });

  it('אבל החמרה על מסמך חנוק נדחית', async () => {
    const { host, xml } = fakeEngine();
    xml.get('s0')!.pgMar.top = 15000;

    expect((await applyPageMargins(host, { topTwips: 15100 })).ok).toBe(false);
  });

  it('כשאין גודל עמוד אין על מה להתלונן', async () => {
    // מקטע בלי `pageSetup` — אין דרך לדעת אם יש מקום, והסירוב היה שרירותי.
    const { host } = fakeEngine({ startWidth: 0, startHeight: 0 });

    expect((await applyPageMargins(host, { topTwips: 14000 })).ok).toBe(true);
  });
});

describe('applyPageMargins', () => {
  it("שולחת אינצ'ים, ורק את שני הצדדים שהסרגל גורר", async () => {
    // נמדד על המנוע: `setPageMargins({left, right})` משאיר את top/bottom כפי
    // שהם. שליחת ארבעתם הייתה משכתבת שוליים שהמשתמש לא נגע בהם.
    const { host, calls, xml } = fakeEngine();

    expect(await applyPageMargins(host, { leftTwips: 2880, rightTwips: 720 })).toEqual({ ok: true });

    expect(calls[0]!.input).toEqual({
      target: { kind: 'section', sectionId: 's0' },
      left: 2,
      right: 0.5,
    });
    expect(xml.get('s0')!.pgMar).toEqual({ top: 1440, right: 720, bottom: 1440, left: 2880 });
  });

  it('ערך שאינו שלם אי-שלילי נעצר לפני המנוע', async () => {
    const { host, calls } = fakeEngine();

    const outcome = await applyPageMargins(host, { leftTwips: -100, rightTwips: 720 });

    expect(outcome.ok).toBe(false);
    expect(calls).toEqual([]);
  });

  it('הסרגל האנכי שולח רק מעלה או מטה — הצדדים האופקיים אינם נוגעים', async () => {
    const { host, calls, xml } = fakeEngine();

    expect(await applyPageMargins(host, { topTwips: 2880 })).toEqual({ ok: true });

    expect(calls[0]!.input).toEqual({ target: { kind: 'section', sectionId: 's0' }, top: 2 });
    expect(xml.get('s0')!.pgMar).toEqual({ top: 2880, right: 1440, bottom: 1440, left: 1440 });
  });

  it('קריאה בלי אף צד נעצרת אצלנו — המנוע היה דוחה אותה', async () => {
    const { host, calls } = fakeEngine();

    expect((await applyPageMargins(host, {})).ok).toBe(false);
    expect(calls).toEqual([]);
  });

  it('חלה על כל המקטעים, כמו הגלריה „שוליים”', async () => {
    const { host, xml } = fakeEngine({ sectionIds: ['s0', 's1'] });

    await applyPageMargins(host, { leftTwips: 720, rightTwips: 720 });

    expect(xml.get('s0')!.pgMar.left).toBe(720);
    expect(xml.get('s1')!.pgMar.left).toBe(720);
  });

  it('גרסה בלי `setPageMargins` מדווחת ואינה מנחשת', async () => {
    const { host } = fakeEngine({ omit: ['setPageMargins'] });
    const outcome = await applyPageMargins(host, { leftTwips: 720, rightTwips: 720 });
    expect(outcome.ok).toBe(false);
  });
});

describe('applyOrientation', () => {
  it('שולחת orientation בלבד — המנוע מחליף את המידות בעצמו', async () => {
    const { host, calls, xml } = fakeEngine();

    expect(await applyOrientation(host, 'landscape')).toEqual({ ok: true });
    expect(calls[0]!.input).toEqual({
      target: { kind: 'section', sectionId: 's0' },
      orientation: 'landscape',
    });
    // 12240×15840 (Letter לאורך) התהפך.
    expect(xml.get('s0')!.pgSz).toEqual({ w: 15840, h: 12240, orient: 'landscape' });
  });

  it('חזרה ל„לאורך” מחזירה את היחס', async () => {
    const { host, xml } = fakeEngine({ startWidth: 15840, startHeight: 12240 });

    await applyOrientation(host, 'portrait');

    expect(xml.get('s0')!.pgSz).toEqual({ w: 12240, h: 15840, orient: 'portrait' });
  });
});

describe('applyPaperSize', () => {
  it('A4 = 11906 × 16838 twips עם קוד נייר 9', async () => {
    const { host, xml } = fakeEngine();

    expect(await applyPaperSize(host, 'a4')).toEqual({ ok: true });
    expect(xml.get('s0')!.pgSz).toEqual({ w: 11906, h: 16838, code: '9' });
  });

  it('Letter = 12240 × 15840 twips עם קוד נייר 1', async () => {
    const { host, xml } = fakeEngine({ startWidth: 11906, startHeight: 16838 });

    expect(await applyPaperSize(host, 'letter')).toEqual({ ok: true });
    expect(xml.get('s0')!.pgSz).toEqual({ w: 12240, h: 15840, code: '1' });
  });

  it('המידות המדויקות נשמרות אף שהן נשלחות כשבר של אינץ\'', async () => {
    // 11906/1440 אינו מספר עגול; העיגול במנוע חייב להחזיר את ה-twips המקורי.
    const { calls, host } = fakeEngine();

    await applyPaperSize(host, 'a4');

    const input = calls[0]!.input;
    expect(Math.round((input.width as number) * TWIPS_PER_INCH)).toBe(11906);
    expect(Math.round((input.height as number) * TWIPS_PER_INCH)).toBe(16838);
  });

  it('במקטע שהוא לרוחב המידות מוחלפות, ואין סתירה בין orient למידות', async () => {
    // בלי ההחלפה היה נשאר w:orient="landscape" על דף שמידותיו לאורך.
    const { host, xml } = fakeEngine({ startWidth: 15840, startHeight: 12240 });

    await applyPaperSize(host, 'a4');

    expect(xml.get('s0')!.pgSz).toEqual({ w: 16838, h: 11906, code: '9' });
  });

  it('כל גודל בטבלה עובר את הוולידציה של המנוע', async () => {
    for (const size of PAPER_SIZES) {
      const { host } = fakeEngine();
      expect(await applyPaperSize(host, size.id), size.id).toEqual({ ok: true });
    }
  });

  it('גודל שאינו קיים אינו נוגע במנוע', async () => {
    const { host, calls } = fakeEngine();

    expect((await applyPaperSize(host, 'a3')).ok).toBe(false);
    expect(calls).toEqual([]);
  });
});

describe('applyColumns', () => {
  it('שולחת count שלם, equalWidth ורווח של חצי אינץ\'', async () => {
    const { host, calls, xml } = fakeEngine();

    expect(await applyColumns(host, 2)).toEqual({ ok: true });
    expect(calls[0]!.input).toEqual({
      target: { kind: 'section', sectionId: 's0' },
      count: 2,
      gap: 0.5,
      equalWidth: true,
    });
    // 720 twips — הרווח שהמסמך הריק נושא ושWord קובע ב-presets.
    expect(xml.get('s0')!.cols).toEqual({ num: 2, space: COLUMN_GAP_TWIPS, equalWidth: true });
  });

  it('מספר עמודות שאינו שלם חיובי נעצר לפני המנוע', async () => {
    // הוולידציה במנוע **זורקת** על ערך כזה, ולא מחזירה קבלה.
    for (const count of [0, -1, 1.5, Number.NaN]) {
      const { host, calls } = fakeEngine();

      const outcome = await applyColumns(host, count);

      expect(outcome.ok, String(count)).toBe(false);
      expect(calls).toEqual([]);
    }
  });
});

describe('פתרון המקטע והדיווח', () => {
  it('מוחלת על כל מקטעי המסמך — כמו „החל על: כל המסמך” ב-Word', async () => {
    const { host, xml, calls } = fakeEngine({ sectionIds: ['s0', 's1', 's2'] });

    expect(await applyMarginPreset(host, 'narrow')).toEqual({ ok: true });
    expect(calls).toHaveLength(3);
    for (const id of ['s0', 's1', 's2']) {
      expect(xml.get(id)!.pgMar.top, id).toBe(720);
    }
  });

  it('מסמך בלי מקטעים מדווח ולא קורא לפעולה', async () => {
    const { host, calls } = fakeEngine({ sectionIds: [] });

    const outcome = await applyMarginPreset(host, 'normal');

    expect(outcome).toEqual({
      ok: false,
      message: 'שינוי השוליים ל„רגיל” נכשל: לא נמצא מקטע במסמך',
      reason: 'target-unresolved',
    });
    expect(calls).toEqual([]);
  });

  it('אין Document API — הודעה בעברית, לא חריגה', async () => {
    for (const host of [null, undefined, {}, { activeEditor: null }, { activeEditor: { doc: null } }]) {
      const outcome = await applyColumns(host as PageSetupHost, 2);

      expect(outcome).toEqual({
        ok: false,
        message: 'שינוי מספר העמודות ל-2 נכשל: המסמך עדיין נטען',
        reason: 'document-api-unavailable',
      });
    }
  });

  it('גרסה שאין בה את הפעולה מדווחת „אינה נתמכת”', async () => {
    const { host } = fakeEngine({ omit: ['setColumns'] });

    expect(await applyColumns(host, 3)).toEqual({
      ok: false,
      message: 'שינוי מספר העמודות ל-3 נכשל: הפעולה אינה נתמכת בגרסה הזאת של המנוע',
      reason: 'command-unsupported',
    });
  });

  it('גרסה שאין בה `sections.list` מדווחת ולא מנחשת מקטע', async () => {
    const { host } = fakeEngine({ omit: ['list'] });

    expect((await applyMarginPreset(host, 'normal')).ok).toBe(false);
  });

  it('קבלה שנכשלה מתורגמת לעברית עם ההקשר של הפעולה', async () => {
    const { host } = fakeEngine({
      receipt: () => ({ success: false, failure: { code: 'DOCUMENT_READONLY' } }),
    });

    expect(await applyMarginPreset(host, 'wide')).toEqual({
      ok: false,
      message: 'שינוי השוליים ל„רחב” נכשל: המסמך פתוח לקריאה בלבד',
      reason: 'DOCUMENT_READONLY',
    });
  });

  it('קוד כשל שאין לו תרגום מוצג עם ההסבר והקוד של המנוע', async () => {
    const { host } = fakeEngine({
      receipt: () => ({ success: false, failure: { code: 'WEIRD_CODE', message: 'nope' } }),
    });

    const outcome = await applyOrientation(host, 'landscape');

    expect(outcome).toEqual({
      ok: false,
      message: 'שינוי כיוון הדף ל„לרוחב” נכשל: nope (WEIRD_CODE)',
      reason: 'WEIRD_CODE',
    });
  });

  it('NO_OP אינה שגיאה — הערכים כבר מוגדרים', async () => {
    const { host } = fakeEngine({ receipt: () => ({ success: false, failure: { code: 'NO_OP' } }) });

    expect(await applyMarginPreset(host, 'normal')).toEqual({ ok: true });
  });

  it('פעולה שזורקת מדווחת ואינה מפילה את התוסף', async () => {
    const { host } = fakeEngine({
      receipt: () => {
        throw new Error('INVALID_INPUT: nope');
      },
    });

    const outcome = await applyColumns(host, 2);

    expect(outcome.ok).toBe(false);
    expect(outcome).toMatchObject({
      message: 'שינוי מספר העמודות ל-2 נכשל: INVALID_INPUT: nope',
      reason: 'threw',
    });
  });

  it('`sections.list` שזורקת מדווחת ואינה מפילה את התוסף', async () => {
    const { host } = fakeEngine({ throwOnList: true });

    expect(await applyMarginPreset(host, 'normal')).toMatchObject({
      message: 'שינוי השוליים ל„רגיל” נכשל: boom',
      reason: 'threw',
    });
  });

  it('סובלת קבלה סינכרונית וקבלה כהבטחה', async () => {
    const sync = fakeEngine({ receipt: () => ({ success: true }) });
    expect(await applyMarginPreset(sync.host, 'normal')).toEqual({ ok: true });

    const async = fakeEngine({ receipt: () => Promise.resolve({ success: true }) });
    expect(await applyMarginPreset(async.host, 'normal')).toEqual({ ok: true });
  });

  it('כשל במקטע אחד עוצר ומדווח, ואינו נבלע', async () => {
    let call = 0;
    const { host, calls } = fakeEngine({
      sectionIds: ['s0', 's1', 's2'],
      receipt: () => (++call === 2 ? { success: false, failure: { code: 'LOCK_VIOLATION' } } : { success: true }),
    });

    const outcome = await applyMarginPreset(host, 'normal');

    expect(outcome).toMatchObject({ ok: false, reason: 'LOCK_VIOLATION' });
    expect(calls).toHaveLength(2);
  });
});

/* ------------------------------------------------------------------ */
/* גל 10 — פריסת עמוד מתקדמת                                          */
/* ------------------------------------------------------------------ */

/** ה-`sectPr` של המקטע היחיד, אחרי שהוא נבדק מול התקן. */
function legalSectPr(xml: Map<string, SectionXml>, sectionId = 's0'): string {
  const section = xml.get(sectionId)!;
  assertWordLegal(section);
  return sectPrXml(section);
}

describe('applyLineNumbering', () => {
  it('„רציף” כותב lnNumType קנוני עם countBy ו-start של Word', async () => {
    const { host, xml, calls } = fakeEngine();

    expect(await applyLineNumbering(host, 'continuous')).toEqual({ ok: true });
    expect(calls[0]!.input).toEqual({
      target: { kind: 'section', sectionId: 's0' },
      enabled: true,
      restart: 'continuous',
      countBy: 1,
      start: 1,
      distance: undefined,
    });
    expect(legalSectPr(xml)).toContain('<w:lnNumType w:countBy="1" w:start="1" w:restart="continuous"/>');
  });

  it('שלושת אסימוני האיפוס הם אסימוני Word, ולא שמות פנימיים', async () => {
    // זו הבדיקה שקיימת בגלל `eachSection` מול `eachSect` של הערות השוליים:
    // ערך שכן בחוזה ואינו של Word עובר את המנוע ומגיע לקובץ.
    for (const choice of LINE_NUMBER_CHOICES) {
      if (choice.restart === null) continue;
      const { host, xml } = fakeEngine();
      expect(await applyLineNumbering(host, choice.id), choice.id).toEqual({ ok: true });
      expect(legalSectPr(xml)).toContain(`w:restart="${choice.restart}"`);
    }
  });

  it('„ללא” מוריד את lnNumType כולו, ואינו משאיר אלמנט ריק', async () => {
    const { host, xml, calls } = fakeEngine({ lineNumbering: { enabled: true, countBy: 5 } });

    await applyLineNumbering(host, 'newPage');
    expect(legalSectPr(xml)).toContain('<w:lnNumType');

    expect(await applyLineNumbering(host, 'none')).toEqual({ ok: true });
    expect(legalSectPr(xml)).not.toContain('lnNumType');
    expect(calls[calls.length - 1]!.input).toEqual({
      target: { kind: 'section', sectionId: 's0' },
      enabled: false,
    });
  });

  it('משמר countBy, start ו„מהטקסט” שנקבעו ב-Word', async () => {
    // כל קריאה מחליפה את `<w:lnNumType>` כולו (נמדד), ולכן שליחה בלי
    // הערכים הקיימים הייתה מוחקת אותם — הבאג המדויק ש-`footnotes.configure`
    // לא נשלח בגללו.
    const { host, xml } = fakeEngine({
      lineNumbering: { enabled: true, countBy: 5, start: 7, distance: 0.25, restart: 'continuous' },
    });

    expect(await applyLineNumbering(host, 'newPage')).toEqual({ ok: true });

    expect(legalSectPr(xml)).toContain(
      '<w:lnNumType w:countBy="5" w:start="7" w:distance="360" w:restart="newPage"/>',
    );
  });

  it('ערך פסול שהגיע מהמסמך אינו מוחזר למנוע אלא נופל לברירת המחדל', async () => {
    // מסמך שנוצר בכלי אחר עשוי לשאת countBy של מיליארד; החזרה שלו הייתה
    // הופכת אותנו לכותבי הערך הפסול, והמנוע היה מקבל אותו בשקט.
    for (const countBy of [0, -3, 2.5, 1_000_000_000, Number.NaN, 'zigzag']) {
      const { host, xml } = fakeEngine({
        lineNumbering: { enabled: true, countBy: countBy as number },
      });

      expect(await applyLineNumbering(host, 'continuous'), String(countBy)).toEqual({ ok: true });
      expect(legalSectPr(xml)).toContain('w:countBy="1"');
    }
  });

  it('start ו„מהטקסט” פסולים מטופלים באותה מידה', async () => {
    const { host, xml } = fakeEngine({
      lineNumbering: { enabled: true, start: 0, distance: 999 },
    });

    await applyLineNumbering(host, 'continuous');

    const sectPr = legalSectPr(xml);
    expect(sectPr).toContain('w:start="1"');
    // מרחק פסול אינו נשלח כלל; בהיעדרו Word מחשב אותו בעצמו.
    expect(sectPr).not.toContain('w:distance');
  });

  it('countBy בגבול העליון של Word עובר כמות שהוא', async () => {
    const { host, xml } = fakeEngine({ lineNumbering: { enabled: true, countBy: LINE_COUNT_BY_MAX } });

    await applyLineNumbering(host, 'continuous');

    expect(legalSectPr(xml)).toContain(`w:countBy="${LINE_COUNT_BY_MAX}"`);
  });

  it('אפשרות שאינה קיימת אינה נוגעת במנוע', async () => {
    const { host, calls } = fakeEngine();

    const outcome = await applyLineNumbering(host, 'zigzag');

    expect(outcome).toEqual({
      ok: false,
      message: 'שינוי מספרי השורות נכשל: אין אפשרות בשם zigzag',
      reason: 'unknown-line-numbering',
    });
    expect(calls).toEqual([]);
  });

  it('גרסה שאין בה את הפעולה מדווחת „אינה נתמכת”', async () => {
    const { host } = fakeEngine({ omit: ['setLineNumbering'] });

    expect(await applyLineNumbering(host, 'continuous')).toEqual({
      ok: false,
      message: 'שינוי מספרי השורות ל„רציף” נכשל: הפעולה אינה נתמכת בגרסה הזאת של המנוע',
      reason: 'command-unsupported',
    });
  });

  it('כשל של קבלה וחריגה — שניהם מדווחים בעברית עם הטיית הכשל', async () => {
    const failing = fakeEngine({
      receipt: () => ({ success: false, failure: { code: 'DOCUMENT_READONLY' } }),
    });
    expect(await applyLineNumbering(failing.host, 'none')).toEqual({
      ok: false,
      message: 'ביטול מספרי השורות נכשל: המסמך פתוח לקריאה בלבד',
      reason: 'DOCUMENT_READONLY',
    });

    const throwing = fakeEngine({
      receipt: () => {
        throw new Error('boom');
      },
    });
    expect(await applyLineNumbering(throwing.host, 'newSection')).toMatchObject({
      ok: false,
      message: 'שינוי מספרי השורות ל„התחל מחדש בכל מקטע” נכשל: boom',
      reason: 'threw',
    });
  });
});

describe('applyPageBorders', () => {
  it('כל preset בגלריה מייצר pgBorders חוקי בארבעה צדדים', async () => {
    for (const preset of PAGE_BORDER_PRESETS) {
      if (preset.style === null) continue;
      // גבול קיים כדי ש„ללא” לא ייפול ל-NO_OP; כאן זה לא רלוונטי, אבל
      // ההתחלה זהה לכל preset.
      const { host, xml } = fakeEngine();

      expect(await applyPageBorders(host, preset.id), preset.id).toEqual({ ok: true });

      const sectPr = legalSectPr(xml);
      expect(sectPr, preset.id).toContain('<w:pgBorders w:display="allPages" w:offsetFrom="page">');
      for (const side of ['top', 'right', 'bottom', 'left']) {
        expect(sectPr, `${preset.id}/${side}`).toContain(
          `<w:${side} w:val="${preset.style}" w:sz="${preset.size}" w:space="24" w:color="auto"/>`,
        );
      }
    }
  });

  it('„ללא גבול” מוריד את pgBorders, ו-NO_OP על מסמך בלי גבול אינו שגיאה', async () => {
    const { host, xml } = fakeEngine();

    // הסרה על מסמך נקי — הכפיל מחזיר NO_OP, בדיוק כמו המנוע.
    expect(await applyPageBorders(host, 'none')).toEqual({ ok: true });

    await applyPageBorders(host, 'double');
    expect(legalSectPr(xml)).toContain('pgBorders');

    expect(await applyPageBorders(host, 'none')).toEqual({ ok: true });
    expect(legalSectPr(xml)).not.toContain('pgBorders');
  });

  it('סגנון שאינו בגלריה אינו נוגע במנוע', async () => {
    const { host, calls } = fakeEngine();

    expect(await applyPageBorders(host, 'zigzag')).toEqual({
      ok: false,
      message: 'שינוי גבול העמוד נכשל: אין סגנון בשם zigzag',
      reason: 'unknown-page-border',
    });
    expect(calls).toEqual([]);
  });

  it('גרסה בלי `clearPageBorders` מדווחת, ואינה שולחת גבול ריק במקום', async () => {
    const { host, calls } = fakeEngine({ omit: ['clearPageBorders'] });

    expect(await applyPageBorders(host, 'none')).toEqual({
      ok: false,
      message: 'הסרת גבול העמוד נכשלה: הפעולה אינה נתמכת בגרסה הזאת של המנוע',
      reason: 'command-unsupported',
    });
    expect(calls).toEqual([]);
  });

  it('כשל של קבלה מדווח עם שם הסגנון', async () => {
    const { host } = fakeEngine({
      receipt: () => ({ success: false, failure: { code: 'DOCUMENT_READONLY' } }),
    });

    expect(await applyPageBorders(host, 'thick')).toEqual({
      ok: false,
      message: 'הוספת גבול העמוד „קו עבה” נכשלה: המסמך פתוח לקריאה בלבד',
      reason: 'DOCUMENT_READONLY',
    });
  });
});

describe('applyVerticalAlign', () => {
  it('כל ארבעת היישורים הם אסימוני ST_VerticalJc', async () => {
    for (const item of VERTICAL_ALIGNS) {
      const { host, xml, calls } = fakeEngine();

      expect(await applyVerticalAlign(host, item.id), item.id).toEqual({ ok: true });
      expect(calls[0]!.input).toEqual({
        target: { kind: 'section', sectionId: 's0' },
        value: item.id,
      });
      expect(legalSectPr(xml)).toContain(`<w:vAlign w:val="${item.id}"/>`);
    }
  });

  it('יישור שאינו קיים אינו נוגע במנוע', async () => {
    const { host, calls } = fakeEngine();

    expect(await applyVerticalAlign(host, 'middle')).toEqual({
      ok: false,
      message: 'שינוי היישור האנכי נכשל: אין יישור בשם middle',
      reason: 'unknown-vertical-align',
    });
    expect(calls).toEqual([]);
  });

  it('גרסה שאין בה את הפעולה, וקבלה שנכשלה', async () => {
    const { host } = fakeEngine({ omit: ['setVerticalAlign'] });
    expect(await applyVerticalAlign(host, 'center')).toEqual({
      ok: false,
      message: 'שינוי היישור האנכי ל„מרכז” נכשל: הפעולה אינה נתמכת בגרסה הזאת של המנוע',
      reason: 'command-unsupported',
    });

    const failing = fakeEngine({ receipt: () => ({ success: false, failure: { code: 'NO_OP' } }) });
    expect(await applyVerticalAlign(failing.host, 'both')).toEqual({ ok: true });
  });
});

describe('applyPageNumbering', () => {
  it('כל ששת הפורמטים הם אסימוני ST_NumberFormat, ונכתבים ל-w:fmt', async () => {
    for (const item of PAGE_NUMBER_FORMATS) {
      const { host, xml } = fakeEngine();

      expect(await applyPageNumbering(host, { format: item.id, start: null }), item.id).toEqual({
        ok: true,
      });
      expect(legalSectPr(xml)).toContain(`<w:pgNumType w:fmt="${item.id}"/>`);
    }
  });

  it('מספר התחלה נשלח רק כשהתבקש, ואינו נכתב כשהוא null', async () => {
    const withStart = fakeEngine();
    expect(await applyPageNumbering(withStart.host, { format: 'decimal', start: 3 })).toEqual({
      ok: true,
    });
    expect(withStart.calls[0]!.input).toEqual({
      target: { kind: 'section', sectionId: 's0' },
      format: 'decimal',
      start: 3,
    });
    expect(legalSectPr(withStart.xml)).toContain('<w:pgNumType w:start="3" w:fmt="decimal"/>');

    const withoutStart = fakeEngine();
    await applyPageNumbering(withoutStart.host, { format: 'upperRoman', start: null });
    expect(withoutStart.calls[0]!.input).toEqual({
      target: { kind: 'section', sectionId: 's0' },
      format: 'upperRoman',
    });
    expect(legalSectPr(withoutStart.xml)).not.toContain('w:start');
  });

  it('מספור עברי אינו נשלח — הוא אינו ב-union והמנוע זורק עליו', async () => {
    // הממצא המרכזי של הגל: `hebrew1` הוא אסימון Word תקני, ובכל זאת אין דרך
    // ציבורית לכתוב אותו. הפקד אינו מציע אותו, ומי שינסה בכל זאת נעצר אצלנו
    // לפני שהמנוע זורק.
    const { host, calls } = fakeEngine();

    const outcome = await applyPageNumbering(host, {
      format: 'hebrew1' as never,
      start: null,
    });

    expect(outcome).toEqual({
      ok: false,
      message: 'שינוי מספור העמודים נכשל: אין פורמט מספור בשם hebrew1',
      reason: 'unknown-page-number-format',
    });
    expect(calls).toEqual([]);
    expect(PAGE_NUMBER_FORMATS.map((item) => item.id)).not.toContain('hebrew1');
  });

  it('מספר התחלה פסול נעצר לפני המנוע', async () => {
    for (const start of [0, -5, 2.5, NUMBER_START_MAX + 1, Number.NaN]) {
      const { host, calls } = fakeEngine();

      const outcome = await applyPageNumbering(host, { format: 'decimal', start });

      expect(outcome.ok, String(start)).toBe(false);
      expect(outcome, String(start)).toMatchObject({ reason: 'invalid-page-number-start' });
      expect(calls, String(start)).toEqual([]);
    }
  });

  it('התקרה היא 32767 — המספר של Word, כתוב במלואו', async () => {
    // המספר קשוח כאן ולא `NUMBER_START_MAX` בכוונה: זו הבדיקה שנשארת אמת
    // כשמישהו מזיז את הקבוע במודול. `32768` מגיע מהתקן, ולכן מוטציה על
    // התקרה שבמודול הופכת אותו ל„נשלח” — והשער על ה-XML פוסל את מה שנכתב.
    const above = fakeEngine();
    const rejected = await applyPageNumbering(above.host, { format: 'decimal', start: 32768 });
    expect(rejected).toMatchObject({ ok: false, reason: 'invalid-page-number-start' });
    expect(above.calls).toEqual([]);

    const atMax = fakeEngine();
    expect(await applyPageNumbering(atMax.host, { format: 'decimal', start: 32767 })).toEqual({
      ok: true,
    });
    expect(legalSectPr(atMax.xml)).toContain('w:start="32767"');
  });

  it('`normalizePageNumberStart` מקבל גם מחרוזת מהטופס', () => {
    expect(normalizePageNumberStart('7')).toBe(7);
    expect(normalizePageNumberStart(' 12 ')).toBe(12);
    expect(normalizePageNumberStart('')).toBe(null);
    expect(normalizePageNumberStart('abc')).toBe(null);
    expect(normalizePageNumberStart('0')).toBe(null);
    expect(normalizePageNumberStart(NUMBER_START_MAX)).toBe(NUMBER_START_MAX);
  });

  it('גרסה שאין בה את הפעולה מדווחת „אינה נתמכת”', async () => {
    const { host } = fakeEngine({ omit: ['setPageNumbering'] });

    expect(await applyPageNumbering(host, { format: 'decimal', start: null })).toEqual({
      ok: false,
      message: 'שינוי מספור העמודים נכשל: הפעולה אינה נתמכת בגרסה הזאת של המנוע',
      reason: 'command-unsupported',
    });
  });
});

describe('applyHeaderDistance', () => {
  it('סנטימטרים הופכים לאינצ\'ים, וה-XML מקבל twips', async () => {
    // 1.25 ס"מ = 0.492 אינץ' = 708 twips. הבדיקה כאן היא ששתי ההמרות
    // מצטרפות נכון: cm→inch אצלנו, inch→twips במנוע.
    const { host, xml, calls } = fakeEngine();

    expect(
      await applyHeaderDistance(host, { headerCm: HEADER_DISTANCE_DEFAULT_CM, footerCm: 2.5 }),
    ).toEqual({ ok: true });

    expect(calls[0]!.input).toEqual({
      target: { kind: 'section', sectionId: 's0' },
      header: cmToInches(HEADER_DISTANCE_DEFAULT_CM),
      footer: cmToInches(2.5),
    });
    const sectPr = legalSectPr(xml);
    expect(sectPr).toContain(`w:header="${Math.round((1.25 / 2.54) * TWIPS_PER_INCH)}"`);
    expect(sectPr).toContain(`w:footer="${Math.round((2.5 / 2.54) * TWIPS_PER_INCH)}"`);
  });

  it('פסיק עשרוני מהטופס מגיע שלם — הדיאלוג שולח את הטקסט שהוא אימת', async () => {
    // הממצא שסגר את הפער בין הדיאלוג ובין המודול: `normalizeHeaderDistanceCm`
    // מקבל פסיק עשרוני ורווחים, ולכן `'1,5'` הוא ערך שהטופס מסמן כתקין —
    // ואילו `Number('1,5')` הוא `NaN`. מרגע שהדיאלוג שולח את הטקסט כמות
    // שהוא, זו הפונקציה היחידה שמכריעה, וזה מה שנמדד כאן.
    const { host, xml, calls } = fakeEngine();

    expect(await applyHeaderDistance(host, { headerCm: '1,5', footerCm: ' 1,25 ' })).toEqual({
      ok: true,
    });

    expect(calls[0]!.input).toEqual({
      target: { kind: 'section', sectionId: 's0' },
      header: cmToInches(1.5),
      footer: cmToInches(1.25),
    });
    expect(legalSectPr(xml)).toContain(`w:header="${Math.round((1.5 / 2.54) * TWIPS_PER_INCH)}"`);
  });

  it('אפס הוא ערך חוקי — Word מתיר כותרת בקצה הדף', async () => {
    const { host, xml } = fakeEngine();

    expect(await applyHeaderDistance(host, { headerCm: 0, footerCm: 0 })).toEqual({ ok: true });
    expect(legalSectPr(xml)).toContain('w:header="0"');
  });

  it('מרחק מעל התקרה של Word נעצר לפני המנוע — שם הוא נבלע', async () => {
    // נמדד: `header: 99` מוחזר `success: true` ונכתב `w:header="142560"`,
    // כלומר כותרת שני מטרים וחצי מקצה הדף. אין תקרה במנוע.
    for (const cm of [HEADER_DISTANCE_MAX_CM + 0.01, 251.46, -1, Number.NaN, 'zigzag']) {
      const { host, calls } = fakeEngine();

      const outcome = await applyHeaderDistance(host, {
        headerCm: cm as number,
        footerCm: 1.25,
      });

      expect(outcome.ok, String(cm)).toBe(false);
      expect(outcome, String(cm)).toMatchObject({ reason: 'invalid-header-distance' });
      expect(calls, String(cm)).toEqual([]);
    }
  });

  it('הגבול העליון עצמו עובר, והתחתון גם', async () => {
    const { host, xml } = fakeEngine();

    expect(
      await applyHeaderDistance(host, { headerCm: HEADER_DISTANCE_MAX_CM, footerCm: 0 }),
    ).toEqual({ ok: true });
    // 22 אינץ' = 31680 twips, בדיוק על התקרה שהתקן שלנו אוכף.
    expect(legalSectPr(xml)).toContain('w:header="31680"');
  });

  it('`normalizeHeaderDistanceCm` מקבל פסיק עשרוני, כמו במקלדת עברית', () => {
    expect(normalizeHeaderDistanceCm('1,25')).toBeCloseTo(cmToInches(1.25), 10);
    expect(normalizeHeaderDistanceCm(' 2.5 ')).toBeCloseTo(cmToInches(2.5), 10);
    expect(normalizeHeaderDistanceCm('')).toBe(null);
    expect(normalizeHeaderDistanceCm('-0.5')).toBe(null);
  });

  it('גרסה שאין בה את הפעולה, וקבלה שנכשלה', async () => {
    const { host } = fakeEngine({ omit: ['setHeaderFooterMargins'] });
    expect(await applyHeaderDistance(host, { headerCm: 1, footerCm: 1 })).toEqual({
      ok: false,
      message: 'שינוי מרחק הכותרת מקצה הדף נכשל: הפעולה אינה נתמכת בגרסה הזאת של המנוע',
      reason: 'command-unsupported',
    });

    const failing = fakeEngine({
      receipt: () => ({ success: false, failure: { code: 'DOCUMENT_READONLY' } }),
    });
    expect(await applyHeaderDistance(failing.host, { headerCm: 1, footerCm: 1 })).toEqual({
      ok: false,
      message: 'שינוי מרחק הכותרת מקצה הדף נכשל: המסמך פתוח לקריאה בלבד',
      reason: 'DOCUMENT_READONLY',
    });
  });
});

describe('readPageLayoutState', () => {
  it('מחזירה את מרחק הכותרת בסנטימטרים ואת מספור העמודים שבמסמך', async () => {
    const { host } = fakeEngine({
      headerFooterMargins: { header: 0.7, footer: 0.6 },
      pageNumbering: { start: 3, format: 'upperRoman' },
    });

    expect(await readPageLayoutState(host)).toEqual({
      headerDistanceCm: { header: 0.7 * 2.54, footer: 0.6 * 2.54 },
      pageNumberFormat: 'upperRoman',
      pageNumberStart: 3,
    });
  });

  it('מסמך בלי pgNumType מחזיר null, וזה המצב הרגיל ולא תקלה', async () => {
    const { host } = fakeEngine();

    const state = await readPageLayoutState(host);

    expect(state.pageNumberFormat).toBe(null);
    expect(state.pageNumberStart).toBe(null);
    // המסמך הריק כן נושא header/footer של חצי אינץ'.
    expect(state.headerDistanceCm).toEqual({ header: 0.5 * 2.54, footer: 0.5 * 2.54 });
  });

  it('פורמט שאינו ב-union אינו מוצג — הוא היה נשלח בחזרה ונזרק', async () => {
    // מסמך שנוצר ב-Word עם מספור עברי הוא בדיוק המקרה: `hebrew1` יגיע
    // מהקריאה, והטופס אינו יכול להציע אותו כי אישור עליו ייכשל.
    const { host } = fakeEngine({ pageNumbering: { format: 'hebrew1', start: 1 } });

    const state = await readPageLayoutState(host);

    expect(state.pageNumberFormat).toBe(null);
    expect(state.pageNumberStart).toBe(1);
  });

  it('ערך מרחק פסול שהגיע מהמסמך אינו מוצג', async () => {
    const { host } = fakeEngine({ headerFooterMargins: { header: 99, footer: 0.5 } });

    expect((await readPageLayoutState(host)).headerDistanceCm).toBe(null);
  });

  it('בלי Document API, בלי `list`, ו-`list` שזורקת — הכול „אין”, בלי חריגה', async () => {
    expect(await readPageLayoutState(null)).toEqual({
      headerDistanceCm: null,
      pageNumberFormat: null,
      pageNumberStart: null,
    });

    const noList = fakeEngine({ omit: ['list'] });
    expect((await readPageLayoutState(noList.host)).pageNumberFormat).toBe(null);

    const throwing = fakeEngine({ throwOnList: true });
    expect((await readPageLayoutState(throwing.host)).headerDistanceCm).toBe(null);
  });

  it('מסמך בלי מקטעים מחזיר „אין”', async () => {
    const { host } = fakeEngine({ sectionIds: [] });

    expect((await readPageLayoutState(host)).headerDistanceCm).toBe(null);
  });
});

describe('כל הפעולות החדשות חלות על כל המקטעים', () => {
  it('שלושה מקטעים — שלוש קריאות, ו-XML חוקי בכל אחד', async () => {
    const runs: Array<[string, (host: PageSetupHost) => Promise<unknown>]> = [
      ['lineNumbering', (host) => applyLineNumbering(host, 'newPage')],
      ['pageBorders', (host) => applyPageBorders(host, 'double')],
      ['verticalAlign', (host) => applyVerticalAlign(host, 'center')],
      ['pageNumbering', (host) => applyPageNumbering(host, { format: 'decimal', start: 2 })],
      ['headerDistance', (host) => applyHeaderDistance(host, { headerCm: 1.5, footerCm: 1.5 })],
    ];

    for (const [name, run] of runs) {
      const { host, xml, calls } = fakeEngine({ sectionIds: ['s0', 's1', 's2'] });

      expect(await run(host), name).toEqual({ ok: true });

      expect(calls, name).toHaveLength(3);
      for (const id of ['s0', 's1', 's2']) legalSectPr(xml, id);
    }
  });
});

describe('שער התקן עצמו — שהוא באמת תופס', () => {
  /**
   * בדיקה על הבדיקה. `assertWordLegal` הוא מה שמפריד בין „נשלח מה שהתכוונו”
   * ובין „הקובץ חוקי”, ושער שאינו נכשל על מסמך שבור אינו שער. כל אחת מהצורות
   * כאן היא בדיוק מה שהמנוע האמיתי כותב על מוטציה שנמדדה.
   */
  const broken: Array<[string, Partial<SectionXml>]> = [
    ['גבול בלי w:val', { pgBorders: { top: { sz: 8 } } }],
    ['סגנון שאינו ST_Border', { pgBorders: { top: { val: 'zigzag', sz: 8 } } }],
    ['w:sz שאינו שלם', { pgBorders: { top: { val: 'single', sz: 2.5 } } }],
    ['w:sz מעל 96', { pgBorders: { top: { val: 'single', sz: 999 } } }],
    ['w:space מעל 31', { pgBorders: { top: { val: 'single', space: 999 } } }],
    ['צבע עם סולמית', { pgBorders: { top: { val: 'single', color: '#FF0000' } } }],
    ['צבע שאינו הקסה', { pgBorders: { top: { val: 'single', color: 'zigzag' } } }],
    ['w:restart שאינו של Word', { lnNumType: { restart: 'eachSection' } }],
    ['countBy של מיליארד', { lnNumType: { countBy: 1_000_000_000 } }],
    // `{ enabled: true }` לבד — נמדד שהוא מייצר בדיוק את זה.
    ['lnNumType ריק', { lnNumType: {} }],
    ['pgNumType ריק', { pgNumType: {} }],
    ['w:fmt שאינו ST_NumberFormat', { pgNumType: { fmt: 'zigzag' } }],
    // 32768 ולא `NUMBER_START_MAX + 1`: קלט שנגזר מהמודול היה זז יחד עם
    // מוטציה במודול, ואז הבדיקה הייתה מאשרת כל תקרה שהמודול יטען.
    ['w:pgNumType/@w:start מעל 32767', { pgNumType: { start: 32768 } }],
    // שני אלה נגזרים מ-`ST_PageBorderDisplay` ומ-`ST_PageBorderOffset`, והם
    // הסעיפים היחידים בשער שלא היה להם קלט שבור כאן. `top` תקין נוסף כדי
    // שהכשל יגיע מהתכונה הנבדקת ולא מגבול חסר `w:val`.
    ['w:display שאינו של Word', { pgBorders: { display: 'everyPage', top: { val: 'single' } } }],
    ['w:offsetFrom שאינו של Word', { pgBorders: { offsetFrom: 'margin', top: { val: 'single' } } }],
    ['vAlign שאינו ST_VerticalJc', { vAlign: 'middle' }],
  ];

  for (const [name, patch] of broken) {
    it(`נכשל על ${name}`, () => {
      const { xml } = fakeEngine();
      Object.assign(xml.get('s0')!, patch);

      expect(() => assertWordLegal(xml.get('s0')!)).toThrow();
    });
  }

  it('עובר על המסמך הריק כמות שהוא', () => {
    const { xml } = fakeEngine();

    expect(() => assertWordLegal(xml.get('s0')!)).not.toThrow();
  });

  it('מרחק כותרת מעל 22 אינץ\' נתפס גם הוא', () => {
    const { xml } = fakeEngine();
    xml.get('s0')!.hfMar.header = 142560;

    expect(() => assertWordLegal(xml.get('s0')!)).toThrow();
  });

  it('התקרות בשער אינן זזות עם התקרות שבמודול', () => {
    // הבדיקה על החולשה העקרונית של השער, ולא על צורה שבורה נוספת: כל עוד
    // הוא ייבא את הקבועים מ-page-setup.ts, מוטציה על אותם קבועים הזיזה גם
    // את מה שהוא בודק. השוואה מול המספרים של ECMA-376 היא מה שמוכיח
    // שהמספרים בשער הם של התקן — ומוטציה במודול נשארת גלויה.
    expect(STD_NUMBER_START_MAX).toBe(32767);
    expect(STD_LINE_COUNT_BY_MAX).toBe(100);
    expect(STD_DISTANCE_MAX_TWIPS).toBe(31680);

    const { xml } = fakeEngine();
    Object.assign(xml.get('s0')!, { pgNumType: { start: STD_NUMBER_START_MAX + 1 } });
    expect(() => assertWordLegal(xml.get('s0')!)).toThrow();
  });
});
