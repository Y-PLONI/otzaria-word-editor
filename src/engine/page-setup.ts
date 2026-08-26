/**
 * לשונית „פריסה”: שוליים, כיוון, גודל נייר ועמודות — דרך `doc.sections`.
 *
 * ## היחידות (נמדד, לא הונח)
 *
 * המסמך הריק של המנוע ארוז ב-base64 בתוך
 * `node_modules/superdoc/dist/chunks/blank-docx-*.es.js`. חילוץ ה-ZIP וקריאת
 * `word/document.xml` נותנים את ה-`sectPr` הזה:
 *
 *     <w:pgSz w:w="12240" w:h="15840"/>
 *     <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
 *              w:header="720" w:footer="720" w:gutter="0"/>
 *
 * כלומר ה-XML הוא twips, 1440 לאינץ' — `1440` = אינץ' אחד = 2.54 ס"מ, וגודל
 * המסמך הריק הוא Letter (12240×15840) ולא A4.
 *
 * **אבל ה-API אינו מקבל twips אלא אינצ'ים.** זה נמדד במימוש עצמו
 * (`@superdoc/docx-engine`, הפונקציות שכותבות את ה-`sectPr`): כל אחת מהן
 * כותבת `String(Math.round(value * 1440))`. כלומר `setPageMargins({top: 1})`
 * מייצר `w:top="1440"`, ו-`setPageMargins({top: 1440})` היה מייצר
 * `w:top="2073600"` — שולי דף בגובה 36 מטר. אותה המרה חלה על
 * `setPageSetup.width/height` ועל `setColumns.gap`.
 *
 * לכן הקבועים כאן נשמרים ב-twips — אלה המספרים שנמדדו ושמופיעים ב-OOXML —
 * והחלוקה ב-`TWIPS_PER_INCH` נעשית ברגע הקריאה למנוע. שתי היחידות גלויות
 * במקום שבו הן נפגשות, ולא מוסתרות בקבוע אחד ששמו לא אומר מה הוא.
 *
 * ## כיוון הדף
 *
 * אין צורך להחליף `width`/`height` בעצמנו: המימוש בודק את היחס הקיים ומחליף
 * לבדו כשהוא אינו מתאים לכיוון המבוקש (`landscape && w<=h` או
 * `portrait && w>h`). לכן `applyOrientation` שולח `orientation` בלבד. לעומת
 * זאת **החלפת גודל נייר כן חייבת להתחשב בכיוון**: המימוש מחליף רק כשנשלח
 * `orientation`, ושליחת מידות A4 לאורך למסמך שהוא לרוחב הייתה משאירה
 * `w:orient="landscape"` על דף שמידותיו לאורך. לכן `applyPaperSize` קורא את
 * הכיוון הנוכחי של המקטע ומחליף את המידות בעצמו.
 *
 * ## על מה זה מוחל
 *
 * על **כל** המקטעים במסמך, ולא על המקטע שבו הסמן. זה מה ש-Word עושה: בדיאלוג
 * „הגדרת עמוד” ברירת המחדל של „החל על” היא „כל המסמך”, וכך גם הגלריות בסרגל.
 * זה גם חוסך תלות ב-`selection` — מיפוי הסמן למקטע דורש את אינדקס הפסקה שלו,
 * וה-selection API אינו חושף אותו ישירות. במסמך רגיל יש מקטע אחד.
 *
 * ## NO_OP אינה שגיאה
 *
 * המנוע מחזיר `success: false, failure.code: 'NO_OP'` כשהערכים המבוקשים כבר
 * מוגדרים. מבחינת המשתמש זו הצלחה — הוא בחר „רגיל” והשוליים רגילים. הודעת
 * שגיאה במצב הזה הייתה גורמת לו לחשוב שהפקד שבור.
 *
 * ## גל 10 — פריסת עמוד מתקדמת, ומה נמדד לפניה
 *
 * Chrome headless על ה-dist הארוז, חמישה סבבים, וכל סבב פורק את ה-zip של
 * `export.toDocx` וקורא את ה-`sectPr` עצמו. הממצא הראשון הפוך מכל תשעת הגלים
 * שקדמו: **`doc.sections` הוא ה-namespace שכן מאמת קלט.** ערך שאינו ב-union
 * **נזרק** ואינו נבלע:
 *
 *     setLineNumbering.restart: 'zigzag'  → „must be one of: continuous, newPage, newSection.”
 *     setLineNumbering.countBy: 0 / -3 / 2.5 → „must be a positive integer.”
 *     setVerticalAlign.value: 'zigzag'    → „must be one of: top, center, bottom, both.”
 *     setBreakType.breakType: 'nextColumn'→ „must be one of: continuous, nextPage, evenPage, oddPage.”
 *     setPageNumbering.format: 'hebrew1'  → „must be one of: decimal, lowerLetter, upperLetter,
 *                                            lowerRoman, upperRoman, numberInDash.”
 *
 * ולכן `applyToSections` שעוטף כל קריאה ב-try אינו זהירות יתר כאן אלא המסלול
 * הרגיל: קלט פסול הוא חריגה, לא קבלה.
 *
 * **מספור עמודים עברי אינו אפשרי.** זו ההפך מ-`footnotes.configure`, שבו
 * `numFmt` נכתב גולמית ו-`'hebrew1'` היה מייצר OOXML תקני (ולא נשלח רק מפני
 * שהערך אינו בטיפוסים). כאן ה-union נאכף בזמן ריצה, ואין דרך ציבורית לכתוב
 * `<w:pgNumType w:fmt="hebrew1"/>`. הממצא מדווח ב-docs/engine-gaps.md.
 *
 * ### מה שנכתב, ונמדד ב-docx
 *
 *     <w:type w:val="oddPage"/>
 *     <w:pgBorders w:display="allPages" w:offsetFrom="text" w:zOrder="front">
 *       <w:top w:val="double" w:sz="12" w:space="24" w:color="FF0000"/>…</w:pgBorders>
 *     <w:lnNumType w:countBy="5" w:start="1" w:distance="360" w:restart="newPage"/>
 *     <w:pgNumType w:start="3" w:fmt="upperRoman"/>
 *     <w:vAlign w:val="center"/>
 *     <w:pgMar … w:header="1008" w:footer="864"/>
 *
 * הכול קנוני, בסדר האלמנטים ש-`CT_SectPr` דורש, ובאותה יחידה: **אינצ'ים**
 * ב-API, twips ב-XML (`0.7` → `1008`, `0.25` → `360`) — אותה המרה בדיוק
 * שמתועדת למעלה על השוליים.
 *
 * ### ומה שכן נבלע — כלומר איפה הוולידציה שלנו
 *
 * האכיפה עוצרת ב-enum. מה שעבר אותו נכתב כמות שהוא:
 *
 * - `setPageBorders.borders.*.style` הוא `string` חופשי בחוזה, ו-`'zigzag'`
 *   נכתב `<w:top w:val="zigzag"/>`; `style: ''` מייצר `<w:top w:sz="8"/>` —
 *   כלומר גבול **בלי `w:val`**, ו-`w:val` הוא תכונה נדרשת ב-`CT_Border`.
 * - `size: 999` → `w:sz="999"`, `size: 2.5` → `w:sz="2.5"` (ערך שאינו שלם
 *   בתכונה שהיא `ST_EighthPointMeasure`), `space: 999` → `w:space="999"`
 *   בעוד שהתקרה של Word היא 31.
 * - `color: '#FF0000'` → `w:color="#FF0000"`, ו-`'zigzag'` נכתב אף הוא —
 *   `ST_HexColor` הוא שש ספרות הקסה או `auto` בלבד.
 * - `setHeaderFooterMargins.header: 99` → `w:header="142560"`, כלומר כותרת
 *   במרחק 2.5 מטר מקצה הדף.
 * - `setLineNumbering.distance: 999` → `w:distance="1438560"`, ו-`countBy`
 *   של מיליארד נכתב כמות שהוא.
 * - `setPageNumbering.chapterStyle` ו-`chapterSeparator` **נבלעים לגמרי**:
 *   `{chapterStyle:1,chapterSeparator:'colon'}` החזיר `success: true` וכתב
 *   `<w:pgNumType/>` ריק. אין להם פקד.
 *
 * מה ש**כן** מוברח: מחרוזת עם גרשיים בתוך `style` נכתבת `&quot;` ואינה
 * מזריקה XML. כלומר הסיכון הוא ערך לא חוקי, לא הזרקה.
 *
 * ### ההפיכות נמדדה, ולא הונחה
 *
 * `setLineNumbering({enabled:false})` **מוריד את `<w:lnNumType>` כולו**,
 * ו-`clearPageBorders` מוריד את `<w:pgBorders>` (וקריאה שנייה מוחזרת
 * `NO_OP`). לעומתם `<w:pgNumType>` אינו ניתן להסרה: `setPageNumbering` דורש
 * לפחות שדה אחד ואין לו `clear`, כלומר „המשך מהמקטע הקודם” של Word אינו
 * ניתן להשגה מכאן. זה כתוב בדיאלוג עצמו, ולא נבלע.
 *
 * ### כל קריאה מחליפה את האלמנט כולו
 *
 * `setLineNumbering({enabled:true,countBy:1,distance:0})` השאיר
 * `<w:lnNumType w:countBy="1" w:distance="0"/>` — `start` ו-`restart` ירדו.
 * זו בדיוק הסיבה ש-`footnotes.configure` לא נשלח בגל 9. כאן הוא כן נשלח,
 * מפני שכאן **יש קריאה**: `sections.list()` מחזיר `lineNumbering` מלא, ולכן
 * הבחירה „רציף” משמרת את `countBy`/`start`/`distance` שהמשתמש קבע ב-Word
 * במקום למחוק אותם.
 *
 * וההשלמה נקראת מאותו תצלום שהכתובת נלקחה ממנו — `sections.list()` אחד,
 * בלי קריאה שנייה. פעולה שקוראת מצב ואז משנה אותו היא TOCTOU, וכאן החלון
 * נסגר במבנה ולא בבדיקה נוספת. הנעילה שמונעת פעולה שנייה בזמן שהראשונה
 * באוויר יושבת ב-LayoutTab.vue.
 *
 * ### `setBreakType` — לא נשלח
 *
 * הפעולה עובדת וכותבת `<w:type w:val="oddPage"/>` קנונית. מה שאין לה הוא
 * פקד שאפשר להציג: כל פעולות המודול הזה חלות על **כל** המקטעים (ראו „על מה
 * זה מוחל”), ובמסמך בעל מקטע יחיד — היחיד שהתוסף עצמו יודע לייצר — ה-`w:type`
 * היחיד שנכתב מתאר איך **המסמך** מתחיל, כלומר אינו עושה דבר. במסמך מרובה
 * מקטעים הוא הופך פקד של מקטע אחד לסריקה שמשכתבת את כל מעברי המקטע שהמשתמש
 * קבע ב-Word. שתי ההתנהגויות שגויות, ולכן דווח ולא נשלח.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';

/** 1440 twips לאינץ'. נמדד ב-`w:pgMar w:top="1440"` של המסמך הריק. ראו הערת הפתיחה. */
export const TWIPS_PER_INCH = 1440;

export interface MarginPreset {
  id: string;
  label: string;
  /** מה שמוצג בתפריט מתחת לשם, בסנטימטרים — היחידה שהמשתמש חושב בה. */
  hint: string;
  /** twips. ההמרה לאינצ'ים נעשית בקריאה למנוע. */
  top: number;
  right: number;
  bottom: number;
  left: number;
}

/**
 * ה-presets של Word, בערכים שלו בדיוק. „רחב” אינו 5.08 מכל צד: ב-Word הוא
 * 2.54 למעלה ולמטה ו-5.08 בצדדים, וזו הפריסה שנראית כמו מסמך רחב-שוליים ולא
 * כמו טקסט שנדחס לריבוע.
 */
export const MARGIN_PRESETS: readonly MarginPreset[] = [
  { id: 'normal', label: 'רגיל', hint: '2.54 ס"מ מכל צד', top: 1440, right: 1440, bottom: 1440, left: 1440 },
  { id: 'narrow', label: 'צר', hint: '1.27 ס"מ מכל צד', top: 720, right: 720, bottom: 720, left: 720 },
  { id: 'wide', label: 'רחב', hint: '2.54 ס"מ למעלה ולמטה, 5.08 בצדדים', top: 1440, right: 2880, bottom: 1440, left: 2880 },
];

export interface PaperSize {
  id: string;
  label: string;
  hint: string;
  /** twips, לאורך. ראו הערת הפתיחה. */
  widthTwips: number;
  heightTwips: number;
  /**
   * `w:pgSz/@w:code` — קוד גודל הנייר של Windows (DMPAPER). Word כותב אותו
   * לצד המידות, וקוד שאינו מתאים למידות מבלבל את דיאלוג ההדפסה.
   */
  code: string;
}

/** A4 ראשון: זו ברירת המחדל בעברית, והמסמך הריק של המנוע דווקא נפתח ב-Letter. */
export const PAPER_SIZES: readonly PaperSize[] = [
  { id: 'a4', label: 'A4', hint: '21 × 29.7 ס"מ', widthTwips: 11906, heightTwips: 16838, code: '9' },
  { id: 'letter', label: 'Letter', hint: '21.6 × 27.9 ס"מ', widthTwips: 12240, heightTwips: 15840, code: '1' },
];

export type PageOrientation = 'portrait' | 'landscape';

export const ORIENTATIONS: readonly { id: PageOrientation; label: string; hint: string }[] = [
  { id: 'portrait', label: 'לאורך', hint: 'הדף גבוה מרוחבו' },
  { id: 'landscape', label: 'לרוחב', hint: 'הדף רחב מגובהו' },
];

export const COLUMN_CHOICES: readonly { count: number; label: string; hint: string }[] = [
  { count: 1, label: 'אחת', hint: 'עמודה אחת' },
  { count: 2, label: 'שתיים', hint: 'שתי עמודות שוות' },
  { count: 3, label: 'שלוש', hint: 'שלוש עמודות שוות' },
];

/** 720 twips = חצי אינץ'. זה הרווח שהמסמך הריק נושא (`w:cols w:space="720"`) וזה שWord קובע ב-presets. */
export const COLUMN_GAP_TWIPS = 720;

/**
 * מה שנקרא מ-`sections.list()`.
 *
 * `pageSetup` נדרש כדי לזהות מקטע שהוא לרוחב, ו-`lineNumbering` כדי לשמר את
 * מה שהמשתמש קבע ב-Word: כל קריאה ל-`setLineNumbering` מחליפה את
 * `<w:lnNumType>` כולו (נמדד), ולכן בחירת „רציף” בלי הערכים הקיימים הייתה
 * מוחקת `countBy` ו-`start` שאיש לא ביקש למחוק. שאר השדות נקראים לצורך
 * הדיאלוגים בלבד — ראו `readPageLayoutState`.
 */
interface SectionItem {
  address?: unknown;
  pageSetup?: { width?: number; height?: number; orientation?: string };
  /** שולי הדף, באינצ'ים — אותה יחידה שבה `setPageMargins` מקבל אותם. */
  margins?: { top?: number; right?: number; bottom?: number; left?: number };
  /** `'rtl'` במסמך עברי. נכתב על ידי `<w:bidi>` ב-`sectPr`. */
  sectionDirection?: string;
  lineNumbering?: {
    enabled?: boolean;
    countBy?: number;
    start?: number;
    distance?: number;
    restart?: string;
  };
  headerFooterMargins?: { header?: number; footer?: number };
  pageNumbering?: { start?: number; format?: string };
  verticalAlign?: string;
}

export interface PageSetupDocumentApi {
  sections?: {
    list?: () => MaybePromise<{ items?: readonly SectionItem[] } | undefined>;
    setPageMargins?: (input: {
      target: unknown;
      top?: number;
      right?: number;
      bottom?: number;
      left?: number;
    }) => MaybePromise<DocReceipt>;
    setPageSetup?: (input: {
      target: unknown;
      width?: number;
      height?: number;
      orientation?: PageOrientation;
      paperSize?: string;
    }) => MaybePromise<DocReceipt>;
    setColumns?: (input: {
      target: unknown;
      count?: number;
      gap?: number;
      equalWidth?: boolean;
    }) => MaybePromise<DocReceipt>;
    setLineNumbering?: (input: {
      target: unknown;
      enabled: boolean;
      countBy?: number;
      start?: number;
      distance?: number;
      restart?: LineNumberRestart;
    }) => MaybePromise<DocReceipt>;
    setVerticalAlign?: (input: { target: unknown; value: VerticalAlign }) => MaybePromise<DocReceipt>;
    setHeaderFooterMargins?: (input: {
      target: unknown;
      header?: number;
      footer?: number;
    }) => MaybePromise<DocReceipt>;
    setPageNumbering?: (input: {
      target: unknown;
      start?: number;
      format?: PageNumberFormat;
    }) => MaybePromise<DocReceipt>;
    setPageBorders?: (input: {
      target: unknown;
      borders: PageBordersInput;
    }) => MaybePromise<DocReceipt>;
    clearPageBorders?: (input: { target: unknown }) => MaybePromise<DocReceipt>;
  };
}

/** צד אחד של `<w:pgBorders>`, בשמות של החוזה. `style` הוא `string` חופשי בו — ראו הערת הפתיחה. */
export interface PageBorderSide {
  style: string;
  size: number;
  space: number;
  color: string;
}

export interface PageBordersInput {
  display: 'allPages' | 'firstPage' | 'notFirstPage';
  offsetFrom: 'page' | 'text';
  top: PageBorderSide;
  right: PageBorderSide;
  bottom: PageBorderSide;
  left: PageBorderSide;
}

/** מה שנדרש מ-SuperDoc: רק הפאסדה של המסמך. ראו document-defaults.ts. */
export interface PageSetupHost {
  activeEditor?: { doc?: PageSetupDocumentApi | null } | null;
}

/**
 * ה-union הוא מה שמאפשר להעביר גם את המופע האמיתי וגם כפיל בבדיקות.
 *
 * בלעדיו TypeScript משווה מבנית את `BrowserDocumentApi` המלא מול הצורה
 * המוצהרת כאן, ונכשל על `target: unknown` מול `SectionAddress` — כלומר החוזה
 * המצומצם היה מחייב לשכפל את כל טיפוסי הכתובות של המנוע. אותה תבנית בדיוק
 * כמו `applyHebrewDocumentDefaults`.
 */
export type PageSetupTarget = SuperDoc | PageSetupHost | null | undefined;

type Sections = NonNullable<PageSetupDocumentApi['sections']>;

/** אינצ'ים מ-twips, מעוגל לשש ספרות: `round(x * 1440)` במנוע מחזיר את ה-twips המדויקים. */
export function twipsToInches(twips: number): number {
  return twips / TWIPS_PER_INCH;
}

function unavailable(failedAction: string, detail: string, reason: string): CommandOutcome {
  return { ok: false, message: `${failedAction}: ${detail}`, reason };
}

/**
 * מריצה מוטציה על כל מקטעי המסמך ומחזירה תוצאה אחת.
 *
 * לעולם אינה זורקת: פעולות ה-Document API זורקות `INVALID_INPUT` על קלט פסול
 * במקום להחזיר קבלה, וחריגה מפקד ב-Ribbon מפילה את הרינדור של הרצועה כולה.
 */
async function applyToSections(
  host: PageSetupTarget,
  failedAction: string,
  pick: (sections: Sections) => ((section: SectionItem, target: unknown) => MaybePromise<DocReceipt>) | null,
): Promise<CommandOutcome> {
  const doc = (host as PageSetupHost | null | undefined)?.activeEditor?.doc;
  if (!doc) return unavailable(failedAction, 'המסמך עדיין נטען', 'document-api-unavailable');

  const sections = doc.sections;
  const mutate = sections ? pick(sections) : null;
  if (!sections?.list || !mutate) {
    return unavailable(failedAction, 'הפעולה אינה נתמכת בגרסה הזאת של המנוע', 'command-unsupported');
  }

  let items: readonly SectionItem[];
  try {
    items = (await sections.list())?.items ?? [];
  } catch (error) {
    return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
  }

  const targets = items.filter((item) => item.address !== undefined && item.address !== null);
  if (targets.length === 0) {
    return unavailable(failedAction, 'לא נמצא מקטע במסמך', 'target-unresolved');
  }

  for (const section of targets) {
    let receipt: DocReceipt;
    try {
      receipt = await mutate(section, section.address);
    } catch (error) {
      return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
    }

    const code = receipt?.failure?.code;
    // NO_OP = הערכים כבר מוגדרים. ראו הערת הפתיחה.
    if (receipt?.success === false && code !== 'NO_OP') {
      return { ok: false, message: receiptFailureText(failedAction, receipt), reason: code };
    }
  }

  return { ok: true };
}

export function findMarginPreset(id: string): MarginPreset | undefined {
  return MARGIN_PRESETS.find((preset) => preset.id === id);
}

export function findPaperSize(id: string): PaperSize | undefined {
  return PAPER_SIZES.find((size) => size.id === id);
}

export function applyMarginPreset(
  host: PageSetupTarget,
  presetId: string,
): Promise<CommandOutcome> {
  const preset = findMarginPreset(presetId);
  if (!preset) {
    // באג בקוד שלנו, לא מצב של המסמך: הפקד לא היה צריך להציע את הערך הזה.
    return Promise.resolve({
      ok: false,
      message: `שינוי השוליים נכשל: אין preset בשם ${presetId}`,
      reason: 'unknown-preset',
    });
  }

  return applyToSections(host, `שינוי השוליים ל„${preset.label}” נכשל`, (sections) => {
    const setPageMargins = sections.setPageMargins;
    if (!setPageMargins) return null;
    return (_section, target) =>
      setPageMargins({
        target,
        top: twipsToInches(preset.top),
        right: twipsToInches(preset.right),
        bottom: twipsToInches(preset.bottom),
        left: twipsToInches(preset.left),
      });
  });
}

/** גיאומטריית הדף שהסרגל מצייר. הכול ב-twips, ההמרה מאינצ'ים כאן. */
export interface PageMarginsState {
  pageWidthTwips: number;
  /** גובה הדף — מה שהסרגל האנכי מצייר עליו. */
  pageHeightTwips: number;
  leftTwips: number;
  rightTwips: number;
  topTwips: number;
  bottomTwips: number;
  /** כיוון המקטע. קובע איזה צד של העמוד הוא „ההתחלה” בסרגל. */
  direction: 'rtl' | 'ltr';
}

/** אינצ'ים מ-`sections.list()` → twips שלמים. ערך פסול מוחזר כ-`null`. */
function inchesToTwips(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) return null;
  return Math.round(value * TWIPS_PER_INCH);
}

/**
 * רוחב הדף, השוליים והכיוון של המקטע הראשון — מה שהסרגל צריך כדי לצייר.
 *
 * המקטע הראשון ולא „המקטע שהסמן בו”, מאותה סיבה שכל המודול הזה פועל על כל
 * המקטעים: מיפוי הסמן למקטע דורש את אינדקס הפסקה שלו, וה-selection API אינו
 * חושף אותו (ראו „על מה זה מוחל” בהערת הפתיחה). במסמך רגיל יש מקטע אחד.
 *
 * `null` הוא מצב רגיל ולא כשל: כך זה נראה כשהמסמך עדיין נטען. הסרגל פשוט
 * אינו מצייר — אין לו על מה להתלונן.
 */
export async function readPageMargins(host: PageSetupTarget): Promise<PageMarginsState | null> {
  const list = (host as PageSetupHost | null | undefined)?.activeEditor?.doc?.sections?.list;
  if (typeof list !== 'function') return null;

  let first: SectionItem | undefined;
  try {
    first = (await list())?.items?.[0];
  } catch {
    return null;
  }
  if (!first) return null;

  const pageWidthTwips = inchesToTwips(first.pageSetup?.width);
  const pageHeightTwips = inchesToTwips(first.pageSetup?.height);
  const leftTwips = inchesToTwips(first.margins?.left);
  const rightTwips = inchesToTwips(first.margins?.right);
  const topTwips = inchesToTwips(first.margins?.top);
  const bottomTwips = inchesToTwips(first.margins?.bottom);
  // הכול או כלום: סרגל שמצייר מידת דף אמיתית לצד שוליים משוערים הוא סרגל
  // שמראה מספרים שאינם המסמך.
  if (
    pageWidthTwips === null ||
    pageHeightTwips === null ||
    leftTwips === null ||
    rightTwips === null ||
    topTwips === null ||
    bottomTwips === null
  ) {
    return null;
  }
  if (pageWidthTwips <= 0 || pageHeightTwips <= 0) return null;

  return {
    pageWidthTwips,
    pageHeightTwips,
    leftTwips,
    rightTwips,
    topTwips,
    bottomTwips,
    direction: first.sectionDirection === 'rtl' ? 'rtl' : 'ltr',
  };
}

/** מה שגרירה בסרגל משנה. מה שלא נשלח — לא נגעו בו. */
export interface PageMarginsInput {
  leftTwips?: number;
  rightTwips?: number;
  topTwips?: number;
  bottomTwips?: number;
}

/**
 * שוליים מגרירה בסרגל — האופקי (ימין/שמאל) או האנכי (מעלה/מטה).
 *
 * **רק הצדדים שנגררו נשלחים**, וזה נמדד ולא הונח: `setPageMargins({left, right})`
 * החזיר `success: true` והשאיר את `w:top`/`w:bottom` כפי שהיו (הם נשארו 96px
 * בתצלום המדדים של המנוע). כלומר הפעולה הזאת **אינה** מהמשפחה שמחליפה את
 * האלמנט כולו, בשונה מ-`setIndentation` — ולכן אין צורך לשלוח מצב מלא, ואין
 * סיכון למחוק צד שהמשתמש לא נגע בו.
 *
 * מוחל על כל המקטעים, כמו כל שאר המודול — ראו „על מה זה מוחל”. גרירה בסרגל
 * ו„שוליים ← רגיל” ברצועה חייבות להתנהג אותו דבר; שתי התנהגויות לאותה פעולה
 * הן באג בפני עצמו.
 */
export function applyPageMargins(
  host: PageSetupTarget,
  input: PageMarginsInput,
): Promise<CommandOutcome> {
  const failedAction = 'שינוי השוליים נכשל';
  const sides = [
    ['left', input.leftTwips],
    ['right', input.rightTwips],
    ['top', input.topTwips],
    ['bottom', input.bottomTwips],
  ] as const;

  const payload: Record<string, number> = {};
  for (const [side, twips] of sides) {
    if (twips === undefined) continue;
    if (!Number.isInteger(twips) || twips < 0) {
      return Promise.resolve({
        ok: false,
        message: `${failedAction}: השוליים חייבים להיות מספרים שלמים לא-שליליים`,
        reason: 'invalid-input',
      });
    }
    payload[side] = twipsToInches(twips);
  }

  if (Object.keys(payload).length === 0) {
    // באג בקוד שלנו: ידית שנגררה חייבת לשלוח צד. המנוע היה דוחה קריאה ריקה.
    return Promise.resolve({
      ok: false,
      message: `${failedAction}: לא נמסר אף צד לשינוי`,
      reason: 'invalid-input',
    });
  }

  return applyToSections(host, failedAction, (sections) => {
    const setPageMargins = sections.setPageMargins;
    if (!setPageMargins) return null;
    return (_section, target) => setPageMargins({ target, ...payload });
  });
}

export function applyOrientation(
  host: PageSetupTarget,
  orientation: PageOrientation,
): Promise<CommandOutcome> {
  const label = orientation === 'landscape' ? 'לרוחב' : 'לאורך';
  return applyToSections(host, `שינוי כיוון הדף ל„${label}” נכשל`, (sections) => {
    const setPageSetup = sections.setPageSetup;
    if (!setPageSetup) return null;
    // בלי width/height: המנוע מחליף אותם בעצמו כשהם אינם מתאימים לכיוון.
    return (_section, target) => setPageSetup({ target, orientation });
  });
}

/** האם המקטע כרגע לרוחב. היחס `width > height` אינו תלוי ביחידה שבה נקרא. */
function isLandscape(section: SectionItem): boolean {
  const setup = section.pageSetup;
  if (setup?.orientation === 'landscape') return true;
  if (setup?.orientation === 'portrait') return false;
  const { width, height } = setup ?? {};
  return typeof width === 'number' && typeof height === 'number' && width > height;
}

export function applyPaperSize(
  host: PageSetupTarget,
  sizeId: string,
): Promise<CommandOutcome> {
  const size = findPaperSize(sizeId);
  if (!size) {
    return Promise.resolve({
      ok: false,
      message: `שינוי גודל הדף נכשל: אין גודל בשם ${sizeId}`,
      reason: 'unknown-paper-size',
    });
  }

  return applyToSections(host, `שינוי גודל הדף ל-${size.label} נכשל`, (sections) => {
    const setPageSetup = sections.setPageSetup;
    if (!setPageSetup) return null;
    return (section, target) => {
      const landscape = isLandscape(section);
      // המידות בטבלה הן לאורך; במקטע שהוא לרוחב מחליפים אותן, אחרת נשארת
      // סתירה בין `w:orient` ובין המידות בפועל.
      const width = landscape ? size.heightTwips : size.widthTwips;
      const height = landscape ? size.widthTwips : size.heightTwips;
      return setPageSetup({
        target,
        width: twipsToInches(width),
        height: twipsToInches(height),
        paperSize: size.code,
      });
    };
  });
}

export function applyColumns(
  host: PageSetupTarget,
  count: number,
): Promise<CommandOutcome> {
  if (!Number.isInteger(count) || count < 1) {
    return Promise.resolve({
      ok: false,
      message: `שינוי מספר העמודות נכשל: ${count} אינו מספר עמודות חוקי`,
      reason: 'invalid-column-count',
    });
  }

  return applyToSections(host, `שינוי מספר העמודות ל-${count} נכשל`, (sections) => {
    const setColumns = sections.setColumns;
    if (!setColumns) return null;
    return (_section, target) =>
      setColumns({
        target,
        count,
        gap: twipsToInches(COLUMN_GAP_TWIPS),
        equalWidth: true,
      });
  });
}

/* ------------------------------------------------------------------ */
/* מספרי שורות                                                        */
/* ------------------------------------------------------------------ */

/** `ST_LineNumberRestart` — שלושת האסימונים של Word, ואותם שלושה שה-union אוכף. */
export type LineNumberRestart = 'continuous' | 'newPage' | 'newSection';

export interface LineNumberChoice {
  id: string;
  label: string;
  hint: string;
  /** `null` = „ללא”, כלומר `enabled: false`. */
  restart: LineNumberRestart | null;
}

/**
 * ארבע האפשרויות שבתפריט „מספרי שורות” של Word העברי, באותו סדר.
 *
 * „אפשרויות מספור שורות” של Word אינו כאן: הוא פותח את דיאלוג „הגדרת עמוד”
 * ומאפשר לקבוע `countBy`, `start` ו„מהטקסט”. שלושת אלה **נשמרים** כאן ואינם
 * ניתנים לעריכה — ראו `preservedLineNumbering`. פקד שמאפשר לקבוע אותם היה
 * טופס רביעי בלשונית, ומה שהוא נותן מעבר לזה הוא נישה; מה שהוא **לוקח** —
 * מחיקה בשוגג של הגדרות שהגיעו מהמסמך — כבר קרה, וזה מה שנמנע כאן.
 */
export const LINE_NUMBER_CHOICES: readonly LineNumberChoice[] = [
  { id: 'none', label: 'ללא', hint: 'בלי מספרי שורות', restart: null },
  { id: 'continuous', label: 'רציף', hint: 'המספור נמשך לאורך כל המסמך', restart: 'continuous' },
  { id: 'newPage', label: 'התחל מחדש בכל עמוד', hint: 'כל עמוד מתחיל מ-1', restart: 'newPage' },
  { id: 'newSection', label: 'התחל מחדש בכל מקטע', hint: 'כל מקטע מתחיל מ-1', restart: 'newSection' },
];

/** התקרה של Word ל„ספור לפי”. מעליה המנוע כותב את הערך כמות שהוא (נמדד: מיליארד). */
export const LINE_COUNT_BY_MAX = 100;

/** התקרה של Word למספר התחלה, וגם ל„התחל ב” של מספור העמודים. */
export const NUMBER_START_MAX = 32767;

/** התקרה של Word למרחק מהטקסט ולמרחק הכותרת מקצה הדף: 22 אינץ' = 55.88 ס"מ. */
export const DISTANCE_MAX_INCHES = 22;

export const CM_PER_INCH = 2.54;

export function inchesToCm(inches: number): number {
  return inches * CM_PER_INCH;
}

export function cmToInches(cm: number): number {
  return cm / CM_PER_INCH;
}

export function findLineNumberChoice(id: string): LineNumberChoice | undefined {
  return LINE_NUMBER_CHOICES.find((choice) => choice.id === id);
}

/** מספר שלם בטווח, או `undefined` — כלומר השדה לא יישלח כלל. */
function boundedInteger(raw: unknown, min: number, max: number): number | undefined {
  return Number.isInteger(raw) && (raw as number) >= min && (raw as number) <= max
    ? (raw as number)
    : undefined;
}

/** מידה סופית אי-שלילית בטווח, או `undefined`. */
function boundedMeasure(raw: unknown, max: number): number | undefined {
  return typeof raw === 'number' && Number.isFinite(raw) && raw >= 0 && raw <= max
    ? raw
    : undefined;
}

/**
 * הערכים שיישלחו חזרה למנוע יחד עם ה-`restart` החדש.
 *
 * מה שנקרא מהמסמך **עובר את אותה ולידציה** שערך שהמשתמש היה מקליד עובר, ולא
 * מוחזר בעיוורון: מסמך שהגיע מכלי אחר עשוי לשאת `w:countBy="1000000000"`, ו-
 * החזרה שלו הייתה הופכת אותנו לכותבי הערך הפסול. ערך שאינו קביל נופל לברירת
 * המחדל של Word (`countBy: 1`, `start: 1`), ו„מהטקסט” שאינו קביל פשוט אינו
 * נשלח — בהיעדרו Word מחשב את המרחק בעצמו, וזו ברירת המחדל שלו.
 */
function preservedLineNumbering(section: SectionItem): {
  countBy: number;
  start: number;
  distance?: number;
} {
  const current = section.lineNumbering;
  return {
    countBy: boundedInteger(current?.countBy, 1, LINE_COUNT_BY_MAX) ?? 1,
    start: boundedInteger(current?.start, 1, NUMBER_START_MAX) ?? 1,
    distance: boundedMeasure(current?.distance, DISTANCE_MAX_INCHES),
  };
}

/**
 * „מספרי שורות” — מדליקה, מכבה או משנה את מדיניות האיפוס.
 *
 * הכיבוי מוריד את `<w:lnNumType>` כולו (נמדד), ולכן „ללא” הוא באמת ללא ולא
 * „מספור שקוף”.
 */
export function applyLineNumbering(host: PageSetupTarget, choiceId: string): Promise<CommandOutcome> {
  const choice = findLineNumberChoice(choiceId);
  if (!choice) {
    // באג בקוד שלנו, לא מצב של המסמך: הפקד לא היה צריך להציע את הערך הזה.
    return Promise.resolve({
      ok: false,
      message: `שינוי מספרי השורות נכשל: אין אפשרות בשם ${choiceId}`,
      reason: 'unknown-line-numbering',
    });
  }

  const failedAction =
    choice.restart === null
      ? 'ביטול מספרי השורות נכשל'
      : `שינוי מספרי השורות ל„${choice.label}” נכשל`;

  return applyToSections(host, failedAction, (sections) => {
    const setLineNumbering = sections.setLineNumbering;
    if (!setLineNumbering) return null;
    if (choice.restart === null) {
      return (_section, target) => setLineNumbering({ target, enabled: false });
    }
    const restart = choice.restart;
    return (section, target) =>
      setLineNumbering({ target, enabled: true, restart, ...preservedLineNumbering(section) });
  });
}

/* ------------------------------------------------------------------ */
/* גבולות עמוד                                                        */
/* ------------------------------------------------------------------ */

export interface PageBorderPreset {
  id: string;
  label: string;
  hint: string;
  /** `null` = „ללא גבול”, כלומר `clearPageBorders`. */
  style: string | null;
  /** `w:sz` — שמיניות נקודה. `4` = חצי נקודה, ברירת המחדל של Word. */
  size: number;
}

/**
 * חמישה סגנונות `ST_Border` שקיימים ב-Word, ועוד „ללא”.
 *
 * רשימה סגורה ולא בורר סגנונות חופשי, מפני ש-`style` הוא `string` בחוזה וכל
 * מחרוזת נכתבת ל-`w:val` כמות שהיא (נמדד: `'zigzag'`, וגם `''` שמייצר גבול
 * בלי `w:val` בכלל). מחוץ לרשימה הזאת אין מה שיעצור ערך שיפיל את הקובץ.
 */
export const PAGE_BORDER_PRESETS: readonly PageBorderPreset[] = [
  { id: 'none', label: 'ללא גבול', hint: 'הסרת הגבול מהעמוד', style: null, size: 0 },
  { id: 'single', label: 'קו יחיד', hint: 'חצי נקודה', style: 'single', size: 4 },
  { id: 'thick', label: 'קו עבה', hint: 'שלוש נקודות', style: 'single', size: 24 },
  { id: 'double', label: 'קו כפול', hint: 'שני קווים דקים', style: 'double', size: 6 },
  { id: 'dashed', label: 'מקווקו', hint: 'קו מקווקו דק', style: 'dashed', size: 4 },
  { id: 'dotted', label: 'מנוקד', hint: 'קו מנוקד דק', style: 'dotted', size: 4 },
];

/** `w:space` — נקודות מקצה הדף. 24 היא ברירת המחדל של Word, והתקרה שלו היא 31. */
export const PAGE_BORDER_SPACE_POINTS = 24;

/** `w:color="auto"` — צבע הטקסט של הנושא, וזה מה ש-Word כותב בגבול ברירת המחדל. */
export const PAGE_BORDER_COLOR = 'auto';

export function findPageBorderPreset(id: string): PageBorderPreset | undefined {
  return PAGE_BORDER_PRESETS.find((preset) => preset.id === id);
}

/**
 * „גבולות עמוד” — מקיפה את הדף בארבעה צדדים, או מסירה את הגבול.
 *
 * ארבעת הצדדים ולא צד אחד: „גבולות עמוד” ב-Word הוא מסגרת, והפקד הזה הוא
 * גלריית מסגרות. `offsetFrom: 'page'` הוא ברירת המחדל של Word — המסגרת נמדדת
 * מקצה הנייר ולא מהטקסט, וזה מה שמייצר את המסגרת שרואים בשער של ספר.
 */
export function applyPageBorders(host: PageSetupTarget, presetId: string): Promise<CommandOutcome> {
  const preset = findPageBorderPreset(presetId);
  if (!preset) {
    return Promise.resolve({
      ok: false,
      message: `שינוי גבול העמוד נכשל: אין סגנון בשם ${presetId}`,
      reason: 'unknown-page-border',
    });
  }

  if (preset.style === null) {
    return applyToSections(host, 'הסרת גבול העמוד נכשלה', (sections) => {
      const clearPageBorders = sections.clearPageBorders;
      if (!clearPageBorders) return null;
      return (_section, target) => clearPageBorders({ target });
    });
  }

  const side: PageBorderSide = {
    style: preset.style,
    size: preset.size,
    space: PAGE_BORDER_SPACE_POINTS,
    color: PAGE_BORDER_COLOR,
  };

  return applyToSections(host, `הוספת גבול העמוד „${preset.label}” נכשלה`, (sections) => {
    const setPageBorders = sections.setPageBorders;
    if (!setPageBorders) return null;
    return (_section, target) =>
      setPageBorders({
        target,
        borders: {
          display: 'allPages',
          offsetFrom: 'page',
          top: side,
          right: side,
          bottom: side,
          left: side,
        },
      });
  });
}

/* ------------------------------------------------------------------ */
/* יישור אנכי                                                          */
/* ------------------------------------------------------------------ */

/** `ST_VerticalJc`. ארבעת האסימונים, וה-union אוכף אותם בזמן ריצה. */
export type VerticalAlign = 'top' | 'center' | 'both' | 'bottom';

export const VERTICAL_ALIGNS: readonly { id: VerticalAlign; label: string; hint: string }[] = [
  { id: 'top', label: 'למעלה', hint: 'הטקסט מתחיל בראש העמוד' },
  { id: 'center', label: 'מרכז', hint: 'הטקסט ממורכז בין הכותרות' },
  { id: 'both', label: 'מיושר', hint: 'הפסקאות נפרשות על גובה העמוד' },
  { id: 'bottom', label: 'למטה', hint: 'הטקסט צמוד לתחתית העמוד' },
];

export function findVerticalAlign(id: string): { id: VerticalAlign; label: string } | undefined {
  return VERTICAL_ALIGNS.find((item) => item.id === id);
}

/**
 * „יישור אנכי” — איפה יושב הטקסט בגובה העמוד.
 *
 * זה הפקד ששער של ספר תורני נשען עליו: „מרכז” על מקטע בעל עמוד אחד הוא בדיוק
 * הדרך שבה Word ממרכז דף שער.
 */
export function applyVerticalAlign(host: PageSetupTarget, value: string): Promise<CommandOutcome> {
  const choice = findVerticalAlign(value);
  if (!choice) {
    return Promise.resolve({
      ok: false,
      message: `שינוי היישור האנכי נכשל: אין יישור בשם ${value}`,
      reason: 'unknown-vertical-align',
    });
  }

  return applyToSections(host, `שינוי היישור האנכי ל„${choice.label}” נכשל`, (sections) => {
    const setVerticalAlign = sections.setVerticalAlign;
    if (!setVerticalAlign) return null;
    return (_section, target) => setVerticalAlign({ target, value: choice.id });
  });
}

/* ------------------------------------------------------------------ */
/* מספור עמודים                                                        */
/* ------------------------------------------------------------------ */

export type PageNumberFormat =
  | 'decimal'
  | 'lowerLetter'
  | 'upperLetter'
  | 'lowerRoman'
  | 'upperRoman'
  | 'numberInDash';

/**
 * ששת הפורמטים שה-union מתיר, וכולם אסימוני `ST_NumberFormat` תקניים.
 *
 * **אין כאן מספור עברי.** `format: 'hebrew1'` — המספור שספר תורני רוצה —
 * נזרק בזמן ריצה עם רשימת ה-union, ואין דרך ציבורית אחרת לכתוב את התכונה.
 * זו הבחנה מול הערות השוליים, שבהן אותו ערך דווקא מגיע ל-XML; ההנמקה בהערת
 * הפתיחה ובמסמך הפערים.
 */
export const PAGE_NUMBER_FORMATS: readonly { id: PageNumberFormat; label: string }[] = [
  { id: 'decimal', label: '1, 2, 3' },
  { id: 'upperLetter', label: 'A, B, C' },
  { id: 'lowerLetter', label: 'a, b, c' },
  { id: 'upperRoman', label: 'I, II, III' },
  { id: 'lowerRoman', label: 'i, ii, iii' },
  { id: 'numberInDash', label: '- 1 -, - 2 -, - 3 -' },
];

export const PAGE_NUMBER_START_HINT = `מספר ההתחלה חייב להיות מספר שלם בין 1 ל-${NUMBER_START_MAX}`;

export function isPageNumberFormat(value: unknown): value is PageNumberFormat {
  return PAGE_NUMBER_FORMATS.some((item) => item.id === value);
}

/**
 * מספר ההתחלה כפי שיישלח, או `null` כשהוא פסול.
 *
 * `0` פסול כאן מפני שהמנוע דוחה אותו („must be a positive integer”), ולא
 * מפני שהחלטנו כך — ומספר שאינו שלם, שלילי או מעל התקרה של Word נדחה אצלנו
 * לפני שהוא מגיע לשם, כדי שההודעה תהיה בעברית ותאמר מה הטווח.
 */
export function normalizePageNumberStart(raw: unknown): number | null {
  // ריק אינו מספר, ו-`Number('')` הוא `0`. כאן זה נדחה גם בלי הבדיקה
  // המקדימה (0 אינו בטווח), והבדיקה קיימת כדי שהכוונה תהיה כתובה.
  if (typeof raw === 'string' && raw.trim() === '') return null;
  const value = typeof raw === 'string' ? Number(raw.trim()) : raw;
  return boundedInteger(value, 1, NUMBER_START_MAX) ?? null;
}

export interface PageNumberingSettings {
  format: PageNumberFormat;
  /** `null` = לא לגעת במספר ההתחלה. */
  start: number | null;
}

/**
 * „עיצוב מספרי עמודים” — הפורמט, ואם התבקש גם המספר שממנו מתחילים.
 *
 * „המשך מהמקטע הקודם” של Word אינו כאן, ולא מפני ששכחנו: `setPageNumbering`
 * דורש לפחות שדה אחד ואין לו פעולת ניקוי, ולכן אין דרך למחוק `w:start` שכבר
 * נכתב. הדיאלוג אומר את זה למשתמש לפני שהוא מאשר.
 */
export function applyPageNumbering(
  host: PageSetupTarget,
  settings: PageNumberingSettings,
): Promise<CommandOutcome> {
  const failedAction = 'שינוי מספור העמודים נכשל';

  if (!isPageNumberFormat(settings?.format)) {
    return Promise.resolve({
      ok: false,
      message: `${failedAction}: אין פורמט מספור בשם ${settings?.format}`,
      reason: 'unknown-page-number-format',
    });
  }

  // `null` הוא „אל תיגע”, ולכן הוא אינו נבדק. כל ערך אחר חייב לעבור.
  if (settings.start !== null && normalizePageNumberStart(settings.start) === null) {
    return Promise.resolve({
      ok: false,
      message: `${failedAction}: ${PAGE_NUMBER_START_HINT}`,
      reason: 'invalid-page-number-start',
    });
  }

  const { format, start } = settings;
  return applyToSections(host, failedAction, (sections) => {
    const setPageNumbering = sections.setPageNumbering;
    if (!setPageNumbering) return null;
    return (_section, target) =>
      setPageNumbering(start === null ? { target, format } : { target, format, start });
  });
}

/* ------------------------------------------------------------------ */
/* מרחק הכותרת מקצה הדף                                                */
/* ------------------------------------------------------------------ */

export const HEADER_DISTANCE_MAX_CM = Math.round(inchesToCm(DISTANCE_MAX_INCHES) * 100) / 100;

export const HEADER_DISTANCE_HINT = `המרחק חייב להיות מספר בין 0 ל-${HEADER_DISTANCE_MAX_CM} ס"מ`;

/** ברירת המחדל של Word בעברית: 1.25 ס"מ, שהם 720 twips בדיוק כמו במסמך הריק. */
export const HEADER_DISTANCE_DEFAULT_CM = 1.25;

/**
 * המרחק כפי שיישלח (באינצ'ים), או `null` כשהוא פסול.
 *
 * הקלט בסנטימטרים מפני שזו היחידה שהמשתמש העברי מקליד, וההמרה נעשית כאן ולא
 * בממשק: המנוע בולע `header: 99` וכותב `w:header="142560"` — כותרת במרחק
 * 2.5 מטר מקצה הדף, עם `success: true`. התקרה כאן היא זו של Word.
 */
export function normalizeHeaderDistanceCm(raw: unknown): number | null {
  // `Number('')` הוא `0`, ולכן שדה שרוקן היה נשלח כ„אפס” — כלומר כותרת
  // שנדחפת לקצה הדף במסלול שהמשתמש קורא לו „לא נגעתי”. ריק אינו מידה.
  if (typeof raw === 'string' && raw.trim() === '') return null;
  const cm = typeof raw === 'string' ? Number(raw.trim().replace(',', '.')) : raw;
  if (typeof cm !== 'number' || !Number.isFinite(cm) || cm < 0 || cm > HEADER_DISTANCE_MAX_CM) {
    return null;
  }
  return cmToInches(cm);
}

export interface HeaderDistanceSettings {
  /**
   * סנטימטרים, כפי שהוקלדו — **גם כמחרוזת, ובכוונה**.
   *
   * `normalizeHeaderDistanceCm` מקבל מחרוזת עם רווחים ועם פסיק עשרוני, ולכן
   * הטופס יכול לאשר קלט ש-`Number()` יהפוך ל-`NaN`. דיאלוג שהיה ממיר בעצמו
   * לפני השליחה היה שולח `NaN` על ערך שהוא עצמו סימן כתקין — כפתור פעיל
   * שמסתיים בהודעת שגיאה. הטיפוס מתיר את הטקסט הגולמי כדי שהפונקציה שהכריעה
   * בטופס תהיה גם הפונקציה שממירה כאן, ולא שתי נוסחאות לאותה שאלה.
   */
  headerCm: number | string;
  footerCm: number | string;
}

/** „מרחק מקצה הדף” — הכותרת העליונה והתחתונה, בדיאלוג „הגדרת עמוד → פריסה”. */
export function applyHeaderDistance(
  host: PageSetupTarget,
  settings: HeaderDistanceSettings,
): Promise<CommandOutcome> {
  const failedAction = 'שינוי מרחק הכותרת מקצה הדף נכשל';

  const header = normalizeHeaderDistanceCm(settings?.headerCm);
  const footer = normalizeHeaderDistanceCm(settings?.footerCm);
  if (header === null || footer === null) {
    return Promise.resolve({
      ok: false,
      message: `${failedAction}: ${HEADER_DISTANCE_HINT}`,
      reason: 'invalid-header-distance',
    });
  }

  return applyToSections(host, failedAction, (sections) => {
    const setHeaderFooterMargins = sections.setHeaderFooterMargins;
    if (!setHeaderFooterMargins) return null;
    return (_section, target) => setHeaderFooterMargins({ target, header, footer });
  });
}

/* ------------------------------------------------------------------ */
/* קריאת המצב, בשביל הדיאלוגים                                        */
/* ------------------------------------------------------------------ */

/**
 * מה שהדיאלוגים נפתחים עליו.
 *
 * **מהמקטע הראשון**, מפני שזה מה ש„החל על כל המסמך” אומר: הפעולות כאן חלות
 * על כל המקטעים, ולכן הערך שיוצג הוא הערך שיידרס. במסמך רגיל יש מקטע אחד.
 * כל שדה עשוי להיות `null` — מסמך שאין בו `w:pgNumType` הוא המצב הרגיל,
 * ולא תקלה.
 */
export interface PageLayoutState {
  headerDistanceCm: { header: number; footer: number } | null;
  pageNumberFormat: PageNumberFormat | null;
  pageNumberStart: number | null;
}

export function emptyPageLayoutState(): PageLayoutState {
  return { headerDistanceCm: null, pageNumberFormat: null, pageNumberStart: null };
}

/**
 * קוראת את מצב המקטע הראשון. לעולם אינה זורקת: כשל של קריאה מחזיר „אין”,
 * כלומר הדיאלוג ייפתח על ברירות המחדל של Word — ולא ימציא ערכים.
 */
export async function readPageLayoutState(host: PageSetupTarget): Promise<PageLayoutState> {
  const list = (host as PageSetupHost | null | undefined)?.activeEditor?.doc?.sections?.list;
  if (typeof list !== 'function') return emptyPageLayoutState();

  let first: SectionItem | undefined;
  try {
    first = (await list())?.items?.[0];
  } catch {
    return emptyPageLayoutState();
  }
  if (!first) return emptyPageLayoutState();

  const header = boundedMeasure(first.headerFooterMargins?.header, DISTANCE_MAX_INCHES);
  const footer = boundedMeasure(first.headerFooterMargins?.footer, DISTANCE_MAX_INCHES);
  const start = boundedInteger(first.pageNumbering?.start, 1, NUMBER_START_MAX);

  return {
    // שני הערכים יחד או אף אחד: הדיאלוג מציג שתי שורות, ואי אפשר להציג אחת
    // מהמסמך ואחת מברירת מחדל בלי לומר איזו היא איזו.
    headerDistanceCm:
      header === undefined || footer === undefined
        ? null
        : { header: inchesToCm(header), footer: inchesToCm(footer) },
    // פורמט שאינו ב-union אינו מוצג: הוא היה נבחר בטופס ונשלח בחזרה, ואז
    // נזרק על ידי המנוע — כלומר דיאלוג שאי אפשר לאשר.
    pageNumberFormat: isPageNumberFormat(first.pageNumbering?.format)
      ? first.pageNumbering.format
      : null,
    pageNumberStart: start ?? null,
  };
}
