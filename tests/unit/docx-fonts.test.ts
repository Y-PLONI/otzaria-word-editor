/**
 * גופן חסר אינו עניין של מראה: `lineRule="auto"` גוזר את גובה השורה ממדדי
 * הגופן שנבחר בפועל, ולכן החלפה שגויה — או החלפה מיותרת — משנה את פריסת כל
 * המסמך. מכאן שלוש החובות שהבדיקות כאן שומרות עליהן:
 *
 * 1. **מחליפים רק מה שחסר.** גופן שנפתר במכונה חייב לעבור בלי שייגעו בו.
 * 2. **מחליפים רק עברית.** ה-fallback הלטיני של הדפדפן סביר, והתערבות בו היא
 *    שינוי בלי סיבה.
 * 3. **בייטים מהמארח רק כשאין ברירה.** גופן ב-base64 שוקל מאות קילובייטים,
 *    ובקשה שלו כשיש תחליף מותקן היא ניפוח של ה-WebView.
 *
 * ההחלטה נבדקת דרך `available` מוזרק ולא דרך canvas אמיתי: בדיקה שנשענת על
 * הגופנים שמותקנים במכונה שמריצה אותה בודקת את המכונה, לא את הקוד.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  FONT_ALIAS_STYLE_ID,
  installDocumentFontAliases,
  coversHebrew,
  isFamilyAvailable,
  parseFontTable,
  planFontAliases,
  shapeOf,
  substitutesFor,
  type DeclaredFont,
} from '../../src/engine/docx-fonts';

const { tryCallMock } = vi.hoisted(() => ({ tryCallMock: vi.fn() }));
vi.mock('../../src/host/otzaria-client', () => ({ tryCall: tryCallMock }));

/** טבלת הגופנים של המסמך שהתלונה נולדה ממנו, מקוצרת לשלוש הרשומות שקובעות. */
const REAL_FONT_TABLE =
  '<w:fonts xmlns:w="ns">' +
  '<w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/><w:charset w:val="00"/>' +
  '<w:family w:val="swiss"/><w:pitch w:val="variable"/>' +
  '<w:sig w:usb0="E0002EFF" w:usb1="C000785B" w:csb0="000001FF" w:csb1="00000000"/></w:font>' +
  '<w:font w:name="FrankRuehl DP"><w:panose1 w:val="02000600000000000000"/><w:charset w:val="00"/>' +
  '<w:family w:val="auto"/><w:pitch w:val="variable"/>' +
  '<w:sig w:usb0="00000803" w:usb1="40000000" w:csb0="00000021" w:csb1="00000000"/></w:font>' +
  '<w:font w:name="Guttman Drogolin"><w:panose1 w:val="02010401010101010101"/><w:charset w:val="B1"/>' +
  '<w:family w:val="auto"/><w:pitch w:val="variable"/>' +
  '<w:sig w:usb0="00000801" w:usb1="40000000" w:csb0="00000020" w:csb1="00000000"/></w:font>' +
  '</w:fonts>';

function hebrewFont(overrides: Partial<DeclaredFont> = {}): DeclaredFont {
  return { name: 'חסר', family: 'auto', charset: '00', panose: null, hebrew: true, ...overrides };
}

afterEach(() => {
  tryCallMock.mockReset();
  document.getElementById(FONT_ALIAS_STYLE_ID)?.remove();
});

describe('parseFontTable', () => {
  it('קורא שם, משפחה, charset ו-panose', () => {
    const fonts = parseFontTable(REAL_FONT_TABLE);
    expect(fonts.map((font) => font.name)).toEqual(['Arial', 'FrankRuehl DP', 'Guttman Drogolin']);

    const frank = fonts[1];
    expect(frank.family).toBe('auto');
    expect(frank.panose).toBe('02000600000000000000');
    expect(frank.charset).toBe('00');
  });

  it('מזהה עברית משלושת הסימנים, וכל אחד מהם לבדו מספיק', () => {
    // csb0 ביט 5, usb0 ביט 11, ו-charset B1 — גופנים עבריים ותיקים ממלאים
    // רק חלק מהם, ולכן אף אחד מהשלושה אינו יכול להיות התנאי היחיד.
    const byCsb = parseFontTable('<w:font w:name="a"><w:sig w:csb0="00000020"/></w:font>');
    const byUsb = parseFontTable('<w:font w:name="b"><w:sig w:usb0="00000800"/></w:font>');
    const byCharset = parseFontTable('<w:font w:name="c"><w:charset w:val="B1"/></w:font>');
    expect([byCsb[0].hebrew, byUsb[0].hebrew, byCharset[0].hebrew]).toEqual([true, true, true]);
  });

  it('csb0 ריק אינו מצהיר כלום — נופלים ל-usb0 ולא נעצרים עליו', () => {
    // גופנים מודרניים רבים לא מכריזים דף קוד ותיק כלל (csb0 אפס) ומסתמכים רק
    // על טווח היוניקוד. csb0="00000000" קיים בתג, ולכן "אין csb0 כלל" לבדו לא
    // תופס אותו — חייבים להתייחס אליו כאילו לא הוצהר.
    const font = parseFontTable(
      '<w:font w:name="a"><w:sig w:csb0="00000000" w:usb0="00000800"/></w:font>',
    );
    expect(font[0].hebrew).toBe(true);
  });

  it('גופן לטיני אינו מסומן כעברי', () => {
    const fonts = parseFontTable(REAL_FONT_TABLE);
    expect(fonts[0].hebrew).toBe(false);
    expect(fonts[1].hebrew).toBe(true);
    expect(fonts[2].hebrew).toBe(true);
  });

  it('טבלה פגומה אינה מפילה — היא רק מחזירה פחות', () => {
    expect(parseFontTable('')).toEqual([]);
    expect(parseFontTable('<w:fonts><w:font/></w:fonts>')).toEqual([]);
    expect(parseFontTable('לא XML בכלל')).toEqual([]);
    expect(parseFontTable('<w:font w:name="יחיד"/>')).toHaveLength(1);
  });
});

describe('shapeOf', () => {
  it('`w:family` מפורש קובע', () => {
    expect(shapeOf(hebrewFont({ family: 'roman' }))).toBe('serif');
    expect(shapeOf(hebrewFont({ family: 'swiss' }))).toBe('sans');
    expect(shapeOf(hebrewFont({ family: 'modern' }))).toBe('sans');
  });

  it('בלי `w:family` נופלים ל-PANOSE', () => {
    expect(shapeOf(hebrewFont({ panose: '0203000000000000000' }))).toBe('serif');
    // serif-style 11 ("Normal Sans") נכתב הקסה כ-"0B", לא כ-"11".
    expect(shapeOf(hebrewFont({ panose: '020B000000000000000' }))).toBe('sans');
  });

  it('ספרת serif-style הקסדצימלית (A–F) נקראת נכון', () => {
    // "020B0604020202020204" הוא ה-PANOSE האמיתי של Arial: family=2,
    // serif-style=11 ("Normal Sans", נכתב "0B"). serif-style בטווח 10-15
    // חייב הקסה כדי להיקרא נכון — parseInt במסד 10 היה עוצר ב-"0" ומחזיר
    // 'serif' בטעות.
    expect(shapeOf(hebrewFont({ family: null, panose: '020B0604020202020204' }))).toBe('sans');
    expect(shapeOf(hebrewFont({ family: null, panose: '020F000000000000000' }))).toBe('sans');
  });

  it('בלי מידע — סריפי, כי זה מה שגופני ספרים עבריים הם', () => {
    // בדיוק המקרה של FrankRuehl DP: family=auto, ו-serif-style אפס.
    expect(shapeOf(hebrewFont({ family: 'auto', panose: '02000600000000000000' }))).toBe('serif');
    expect(shapeOf(hebrewFont({ family: null, panose: null }))).toBe('serif');
  });
});

describe('planFontAliases', () => {
  const installed = (names: string[]) => (name: string) => names.includes(name);

  it('גופן שנפתר אינו מוחלף', () => {
    const plan = planFontAliases(parseFontTable(REAL_FONT_TABLE), () => true);
    expect(plan.css).toBe('');
    expect(plan.missing).toEqual([]);
    expect(plan.needBytes).toEqual([]);
  });

  it('גופן לטיני חסר אינו מוחלף', () => {
    // Arial חסר כאן, ובכל זאת אין לו כלל: ה-fallback הלטיני סביר, והחלפה
    // שלו הייתה שינוי מראה בלי סיבה.
    const plan = planFontAliases(parseFontTable(REAL_FONT_TABLE), installed(['FrankRuehl']));
    expect(plan.css).not.toContain('Arial');
    expect(plan.missing.map((font) => font.name)).toEqual(['FrankRuehl DP', 'Guttman Drogolin']);
  });

  it('גופן עברי חסר מקבל שרשרת `local()` של מה שמותקן בלבד', () => {
    const plan = planFontAliases(parseFontTable(REAL_FONT_TABLE), installed(['FrankRuehl', 'David']));
    expect(plan.css).toContain('@font-face{font-family:"FrankRuehl DP";');
    expect(plan.css).toContain('local("FrankRuehl")');
    expect(plan.css).toContain('local("David")');
    // Narkisim אינו מותקן — אין טעם להציע אותו.
    expect(plan.css).not.toContain('Narkisim');
    expect(plan.needBytes).toEqual([]);
  });

  it('בלי אף תחליף מותקן — מבקשים בייטים, ולא כותבים כלל ריק', () => {
    const plan = planFontAliases(parseFontTable(REAL_FONT_TABLE), installed([]));
    expect(plan.css).toBe('');
    expect(plan.needBytes.map((font) => font.name)).toEqual(['FrankRuehl DP', 'Guttman Drogolin']);
  });

  it('אותה משפחה פעמיים בטבלה מקבלת כלל אחד', () => {
    const twice = `<w:fonts>${'<w:font w:name="כפול"><w:charset w:val="B1"/></w:font>'.repeat(2)}</w:fonts>`;
    const plan = planFontAliases(parseFontTable(twice), installed(['FrankRuehl']));
    expect(plan.css.match(/@font-face/g)).toHaveLength(1);
  });

  it('שם עם גרשיים אינו שובר את ה-CSS', () => {
    const odd = '<w:font w:name="גופן &quot;מוזר&quot;"><w:charset w:val="B1"/></w:font>';
    const plan = planFontAliases(parseFontTable(odd), installed(['FrankRuehl']));
    expect(plan.css).toContain('\\"');
  });
});

describe('substitutesFor', () => {
  it('סריפי מקבל שרשרת עברית סריפית, וחלק שרשרת חלקה', () => {
    expect(substitutesFor(hebrewFont({ family: 'roman' }))[0]).toBe('FrankRuehl');
    expect(substitutesFor(hebrewFont({ family: 'swiss' }))[0]).toBe('Arial');
    // FrankRuhlCLM הוא הגופן שאוצריא אורזת, וזה מה שהופך אותו לתחליף שאפשר
    // לקבל את הבייטים שלו גם במכונה ריקה.
    expect(substitutesFor(hebrewFont({ family: 'roman' }))).toContain('FrankRuhlCLM');
  });
});

describe('coversHebrew', () => {
  it('בלי canvas אינו מבטיח כיסוי — ההפך מ-isFamilyAvailable, ובכוונה', () => {
    // שם שהדפדפן פותר אינו בהכרח שם שיש בו אות עברית. הבורר מצייר דגימה של
    // אותיות עבריות לפי התשובה הזאת, ודגימה שלא נמדדה היא דגימה שעלולה
    // להיות של גופן אחר לגמרי — לכן כאן „בלי מדידה” פירושו „לא מבטיחים”.
    expect(coversHebrew('David')).toBe(false);
    expect(coversHebrew('גופן שאינו קיים בשום מקום')).toBe(false);
  });
});

/* ------------------------------------------------------------------ */
/* coversHebrew — המדידה עצמה, על canvas מדומה                        */
/* ------------------------------------------------------------------ */

/**
 * למה canvas מדומה, ולא jsdom ולא הדפדפן.
 *
 * jsdom אינו מודד כלום, ולכן `coversHebrew` שם מחזירה `false` לפני שהיא
 * מגיעה למדידה — כלומר אי אפשר לבדוק שם את הלוגיקה בכלל. canvas אמיתי היה
 * בודק את הגופנים שמותקנים במכונה שמריצה את הבדיקה, ולא את הקוד.
 *
 * מה שהמדמה מחקה הוא בדיוק מה שנמדד בכרום ב-Windows, ושתי העובדות שהוא
 * מחזיק הן מה שהכריע את התיקון:
 *
 *   1. התאמת הגופן ב-CSS היא לכל תו בנפרד: שרשרת `"F", B` על אות עברית
 *      שאין ב-F נופלת ל-B, ולכן מודדת בדיוק כמו `B` לבדו.
 *   2. ברירת המחדל לעברית של `serif` **ושל** `monospace` היא Times New
 *      Roman — של `serif` מפני שכך הוגדר בכרום, ושל `monospace` מפני
 *      ש-Courier New אינו מכסה עברית ונופל בעצמו ל-fallback המערכתי. ורק
 *      של `sans-serif` היא Arial.
 *
 * מכאן שגופן שהוא עצמו ברירת מחדל של בסיס — Times New Roman — נמדד זהה מול
 * שני הבסיסים הראשונים, ורק השלישי מבדיל אותו.
 */
const HEBREW_DEFAULT_OF_BASE: Record<string, string> = {
  monospace: 'Times New Roman',
  serif: 'Times New Roman',
  'sans-serif': 'Arial',
};

/** רוחב מדומה לכל גופן. השונות היא כל מה שהמדידה קוראת. */
const PROBE_WIDTH: Record<string, number> = {
  'Times New Roman': 100,
  Arial: 111,
  Narkisim: 122,
  FrankRuhlCLM: 133,
};

interface FontProbe {
  covers: (name: string) => boolean;
  /** „האם הדפדפן פותר את השם” — שאלה אחרת מכיסוי, ובעלת מטמון משלה. */
  available: (name: string) => boolean;
  /** המיזוג מהמודול **הזה**, כדי שהמדידה שהוא עושה תהיה זו שנספרת כאן. */
  merge: typeof import('../../src/engine/font-options').mergeFontFamilies;
  compose: typeof import('../../src/engine/font-options').composeFontOptions;
  cached: (name: string) => boolean | undefined;
  cachedCoverage: (name: string) => boolean | undefined;
  install: (fontTableXml: string | null) => Promise<string[]>;
  /** כמה מדידות רוחב נעשו בפועל — כך נראה מה נשמר במפה ומה נמדד מחדש. */
  measurements: () => number;
  /** מה שהדפדפן „מכיר” כרגע. ההזרקה של המסמך משנה את זה. */
  hebrew: Set<string>;
  restore: () => void;
}

/**
 * מודול טרי עם canvas מדומה. טרי בכל בדיקה, כי גם ה-context וגם מפת הכיסוי
 * הם מצב ברמת המודול.
 */
async function fontProbe(hebrew: readonly string[]): Promise<FontProbe> {
  const known = new Set(hebrew);
  let measurements = 0;

  const context = {
    font: '',
    measureText(text: string) {
      measurements += 1;
      const parsed = /^72px (?:"(.*)", )?([a-z-]+)$/.exec(context.font);
      if (!parsed || text.length === 0) return { width: 0 };
      const [, name, base] = parsed;
      const drawn = name !== undefined && known.has(name) ? name : HEBREW_DEFAULT_OF_BASE[base];
      return { width: PROBE_WIDTH[drawn] ?? 100 };
    },
  };

  const realCreate = document.createElement.bind(document);
  const stub = vi.spyOn(document, 'createElement').mockImplementation(((tag: string) => {
    const el = realCreate(tag);
    if (tag === 'canvas') (el as unknown as { getContext: () => unknown }).getContext = () => context;
    return el;
  }) as typeof document.createElement);

  vi.resetModules();
  const fonts = await import('../../src/engine/docx-fonts');
  // אחרי ה-reset ובאותו גרף: כך `mergeFontFamilies` קורא ל-`isFamilyAvailable`
  // של המודול המדומה, ולא לזה של גרף אחר שאינו סופר מדידות.
  const options = await import('../../src/engine/font-options');

  return {
    covers: fonts.coversHebrew,
    available: fonts.isFamilyAvailable,
    merge: options.mergeFontFamilies,
    compose: options.composeFontOptions,
    cached: fonts.cachedFamilyAvailability,
    cachedCoverage: fonts.cachedHebrewCoverage,
    install: fonts.installDocumentFontAliases,
    measurements: () => measurements,
    hebrew: known,
    restore: () => stub.mockRestore(),
  };
}

describe('coversHebrew — שלושת הבסיסים', () => {
  it('`Times New Roman` מכסה עברית — ומול שני בסיסים בלבד הוא נמדד `false`', async () => {
    // הדוגמה שהתיעוד עצמו מביא („Arial ו-Times New Roman מכסים עברית”)
    // נמדדה `false` בכרום. הסיבה אינה באג בקוד אלא בבסיסים: `serif`
    // ב-Windows הוא Times New Roman, ו-`monospace` נופל אליו בעצמו.
    const probe = await fontProbe(['Times New Roman', 'Arial', 'Narkisim']);
    try {
      expect(probe.covers('Times New Roman')).toBe(true);
    } finally {
      probe.restore();
    }
  });

  it('וגם `Arial` — ברירת המחדל של הבסיס השלישי — נשאר `true`', async () => {
    // התיקון אינו מחליף פינה אחת באחרת: Arial מזוהה דרך שני הבסיסים
    // הראשונים, בדיוק כמו קודם.
    const probe = await fontProbe(['Times New Roman', 'Arial', 'Narkisim']);
    try {
      expect(probe.covers('Arial')).toBe(true);
      expect(probe.covers('Narkisim')).toBe(true);
    } finally {
      probe.restore();
    }
  });

  it('גופן בלי עברית נשאר `false` — הבסיס השלישי אינו מייצר חיובי-שגוי', async () => {
    // Wingdings, Webdings, Symbol, Marlett, Cambria Math והגופנים
    // המזרח-אסיאתיים נמדדו שליליים בכרום, וכך הם נשארים: שרשרת `"F", B` על
    // אות עברית שאין ב-F נופלת ל-B ומודדת בדיוק כמוהו — בכל אחד משלושת
    // הבסיסים.
    const probe = await fontProbe(['Times New Roman', 'Arial']);
    try {
      for (const name of [
        'Wingdings',
        'Webdings',
        'Symbol',
        'Marlett',
        'Cambria Math',
        'MS Gothic',
        'גופן שאינו מותקן',
      ]) {
        expect(probe.covers(name)).toBe(false);
      }
    } finally {
      probe.restore();
    }
  });
});

describe('coversHebrew — המפתח שהתשובה נשמרת תחתיו', () => {
  it('רישיות אינה מייצרת רשומה שנייה ומדידה שנייה', async () => {
    // `familyKey` ב-font-options.ts ממותת, ולכן `Arial`/`ARIAL`/`arial` הם
    // אפשרות אחת בבורר — בזמן שהמפה כאן מדדה אותם שלוש פעמים.
    const probe = await fontProbe(['Arial']);
    try {
      expect(probe.covers('Arial')).toBe(true);
      const afterFirst = probe.measurements();
      expect(afterFirst).toBeGreaterThan(0);

      expect(probe.covers('ARIAL')).toBe(true);
      expect(probe.covers('  arial  ')).toBe(true);
      expect(probe.measurements()).toBe(afterFirst);
    } finally {
      probe.restore();
    }
  });

  it('הוספה והסרה של אליאס מבטלות כיסוי רק למשפחה שהשתנתה', async () => {
    const probe = await fontProbe(['Arial']);
    try {
      tryCallMock.mockResolvedValue({
        css: '@font-face{font-family:"Narkisim";src:url(data:font/woff2;base64,AA==)}',
        resolved: ['Narkisim'],
      });
      expect(probe.covers('Narkisim')).toBe(false);
      expect(probe.covers('Arial')).toBe(true);
      await probe.install('<w:font w:name="Narkisim"><w:charset w:val="B1"/></w:font>');
      probe.hebrew.add('Narkisim');
      expect(probe.cachedCoverage('Narkisim')).toBeUndefined();
      expect(probe.covers('Narkisim')).toBe(true);
      const measured = probe.measurements();
      expect(probe.covers('Arial')).toBe(true);
      expect(probe.measurements()).toBe(measured);
      await probe.install(null);
      probe.hebrew.delete('Narkisim');
      expect(probe.cachedCoverage('Narkisim')).toBeUndefined();
      expect(probe.covers('Narkisim')).toBe(false);
      expect(probe.cachedCoverage('Arial')).toBe(true);
    } finally { probe.restore(); }
  });
});

describe('isFamilyAvailable', () => {
  it('בלי canvas מחזיר „זמין” — לא מחליפים על סמך ניחוש', () => {
    // jsdom אינו מממש getContext('2d'). זו בדיוק הסביבה שבה אסור לנו להחליט:
    // החלפה שגויה גרועה מלא לגעת.
    expect(isFamilyAvailable('גופן שאינו קיים בשום מקום')).toBe(true);
  });
});

describe('installDocumentFontAliases', () => {
  beforeEach(() => {
    tryCallMock.mockResolvedValue(null);
  });

  it('בלי טבלה — סגנון ריק, ובלי פנייה למארח', async () => {
    await expect(installDocumentFontAliases(null)).resolves.toEqual([]);
    expect(document.getElementById(FONT_ALIAS_STYLE_ID)?.textContent).toBe('');
    expect(tryCallMock).not.toHaveBeenCalled();
  });

  it('קריאה חוזרת דורסת את הסגנון ואינה מוסיפה עוד אחד', async () => {
    await installDocumentFontAliases(REAL_FONT_TABLE);
    await installDocumentFontAliases(REAL_FONT_TABLE);
    expect(document.querySelectorAll(`#${FONT_ALIAS_STYLE_ID}`)).toHaveLength(1);
  });

  it('ממתינה לגופני מארח לפני שמבטלים את תשובת המטמון', async () => {
    /*
     * גופן שמגיע ב-`url(data:...)` נטען אחרי שה-`@font-face` נכתב. ביטול
     * המטמון לפני ה-load מאפשר למדידה ביניים לשמור `false` מיושן.
     * המדמה מחזיק את הטעינה באמצע כדי להגן על סדר ההמתנה והביטול.
     */
    const probe = await fontProbe([]);
    const previousFonts = Object.getOwnPropertyDescriptor(document, 'fonts');
    let release!: () => void;
    const loading = new Promise<void>((resolve) => { release = resolve; });
    const load = vi.fn(async () => {
      await loading;
      return [];
    });

    Object.defineProperty(document, 'fonts', { configurable: true, value: { load } });
    try {
      tryCallMock.mockResolvedValue({
        css: '@font-face{font-family:"Narkisim";src:url(data:font/woff2;base64,AA==)}',
        resolved: ['Narkisim'],
      });
      const fontTable = '<w:font w:name="Narkisim"><w:charset w:val="B1"/></w:font>';
      expect(probe.available('Narkisim')).toBe(false);

      const installing = probe.install(fontTable);
      await vi.waitFor(() => expect(tryCallMock).toHaveBeenCalled());
      await vi.waitFor(() => expect(load).toHaveBeenCalled());

      let settled = false;
      void installing.then(() => { settled = true; });
      await Promise.resolve();
      expect(settled).toBe(false);

      // מרגע שהדפדפן סיים לטעון, אותה משפחה באמת נפתרת דרך האליאס.
      probe.hebrew.add('Narkisim');
      release();
      await installing;
      expect(probe.available('Narkisim')).toBe(true);
    } finally {
      if (previousFonts) Object.defineProperty(document, 'fonts', previousFonts);
      else Reflect.deleteProperty(document, 'fonts');
      probe.restore();
    }
  });

  it('בסביבה בלי canvas שום גופן אינו נחשב חסר, ולכן אין מה להתקין', async () => {
    // הנפילה הבטוחה, מקצה לקצה: אין מדידה → אין החלפה → אין פנייה למארח.
    await expect(installDocumentFontAliases(REAL_FONT_TABLE)).resolves.toEqual([]);
    expect(tryCallMock).not.toHaveBeenCalled();
  });

  it('תשובה חלקית טוענת ומבטלת מטמון רק למה שסופק, ומנסה שוב עד תשובה מלאה', async () => {
    const probe = await fontProbe([]);
    const previousFonts = Object.getOwnPropertyDescriptor(document, 'fonts');
    const load = vi.fn(async (_font: string, _text?: string) => []);
    Object.defineProperty(document, 'fonts', { configurable: true, value: { load } });
    try {
      const names = ['Narkisim', 'FrankRuhlCLM'];
      const xml = names.map((name) => `<w:font w:name="${name}"><w:charset w:val="B1"/></w:font>`).join('');
      const rules = names.map((name) => `@font-face{font-family:"${name}";src:url(data:font/woff2;base64,AA==)}`);
      tryCallMock.mockResolvedValueOnce({ css: rules[0], resolved: ['Narkisim'] });
      tryCallMock.mockResolvedValueOnce({ css: rules.join('\n'), resolved: names });
      expect(probe.available('Narkisim')).toBe(false);
      expect(probe.available('FrankRuhlCLM')).toBe(false);

      await probe.install(xml);
      expect(load).toHaveBeenCalledTimes(1);
      expect(load.mock.calls[0][0]).toBe('72px "Narkisim"');
      expect(probe.cached('Narkisim')).toBeUndefined();
      // משפחה שלא קיבלה כלל אינה בבעלות הסגנון, והמטמון שלה נשאר תקף.
      expect(probe.cached('FrankRuhlCLM')).toBe(false);
      probe.hebrew.add('Narkisim');
      expect(probe.available('Narkisim')).toBe(true);

      await probe.install(xml);
      expect(tryCallMock).toHaveBeenCalledTimes(2);
      expect(load).toHaveBeenCalledTimes(3);
      expect(load.mock.calls.slice(1).map(([font]) => font)).toEqual(['72px "Narkisim"', '72px "FrankRuhlCLM"']);
      probe.hebrew.add('FrankRuhlCLM');
      expect(probe.available('FrankRuhlCLM')).toBe(true);
      const style = document.getElementById(FONT_ALIAS_STYLE_ID)!;
      expect(style.textContent).toBe(rules.join('\n'));
      const write = vi.spyOn(style, 'textContent', 'set');

      // אחרי תשובה מלאה נשמרת יעילות הפתיחה החוזרת.
      await probe.install(xml);
      expect(tryCallMock).toHaveBeenCalledTimes(2);
      expect(load).toHaveBeenCalledTimes(3);
      expect(write).not.toHaveBeenCalled();
      expect(probe.cached('FrankRuhlCLM')).toBe(true);
      write.mockRestore();
    } finally {
      if (previousFonts) Object.defineProperty(document, 'fonts', previousFonts);
      else Reflect.deleteProperty(document, 'fonts');
      probe.restore();
    }
  });

  it('CSS בלי רשימת שמות תקינה אינו מוזרק ואינו חוסם ניסיון חוזר', async () => {
    const probe = await fontProbe([]);
    try {
      const xml = '<w:font w:name="Narkisim"><w:charset w:val="B1"/></w:font>';
      const css = '@font-face{font-family:"Narkisim";src:url(data:font/woff2;base64,AA==)}';
      tryCallMock.mockResolvedValueOnce({ css });
      tryCallMock.mockResolvedValueOnce({ css, resolved: ['Narkisim'] });
      await probe.install(xml);
      expect(document.getElementById(FONT_ALIAS_STYLE_ID)?.textContent).toBe('');
      expect(probe.cached('Narkisim')).toBe(false);
      await probe.install(xml);
      expect(tryCallMock).toHaveBeenCalledTimes(2);
      expect(document.getElementById(FONT_ALIAS_STYLE_ID)?.textContent).toBe(css);
      expect(probe.cached('Narkisim')).toBeUndefined();
    } finally { probe.restore(); }
  });
});

/**
 * `isFamilyAvailable` — השאלה שכל תצוגת הגופן נשענת עליה.
 *
 * הבדיקה היחידה שהייתה כאן קודם טענה שהיא מחזירה `true` בלי canvas, כלומר
 * מדדה את **הסביבה** ולא את הקוד. עם ה-canvas המדומה אפשר סוף־סוף לראות גם
 * `false`, וזה הצד שהתיקון כולו חי ממנו.
 */
describe('isFamilyAvailable — מדידה, מטמון, והתיישנות', () => {
  it('שם שהדפדפן פותר נמדד זמין, ושם שאינו — לא', async () => {
    const probe = await fontProbe(['Narkisim']);
    try {
      expect(probe.available('Narkisim')).toBe(true);
      expect(probe.available('Aptos')).toBe(false);
    } finally {
      probe.restore();
    }
  });

  it('התשובה נשמרת ואינה נמדדת שוב, ובלי רגישות לאותיות ולרווחים', async () => {
    // בלי מיתות היו כאן שלוש רשומות ל-Arial, ו-`keepAvailable` שואל על כל שם
    // שהמכונה מדווחת — נמדד 287 משפחות.
    const probe = await fontProbe(['Narkisim']);
    try {
      probe.available('Narkisim');
      const afterFirst = probe.measurements();
      expect(afterFirst).toBeGreaterThan(0);

      probe.available('Narkisim');
      probe.available('NARKISIM');
      probe.available('  narkisim  ');

      expect(probe.measurements()).toBe(afterFirst);
    } finally {
      probe.restore();
    }
  });

  it('אליאס חוזר אינו מוחק תשובות או מסיר את עצמו', async () => {
    const probe = await fontProbe(['FrankRuhlCLM', 'Arial']);
    try {
      const xml = '<w:font w:name="Narkisim"><w:charset w:val="B1"/></w:font>';
      expect(probe.available('Narkisim')).toBe(false);
      expect(probe.available('Arial')).toBe(true);
      await probe.install(xml);
      probe.hebrew.add('Narkisim');
      expect(probe.available('Narkisim')).toBe(true);
      expect(probe.covers('Narkisim')).toBe(true);
      const style = document.getElementById(FONT_ALIAS_STYLE_ID)!;
      const css = style.textContent;
      const write = vi.spyOn(style, 'textContent', 'set');
      const measured = probe.measurements();
      await probe.install(xml);
      expect(style.textContent).toBe(css);
      expect(css).toContain('font-family:"Narkisim"');
      expect(write).not.toHaveBeenCalled();
      expect(probe.cached('Narkisim')).toBe(true);
      expect(probe.cachedCoverage('Narkisim')).toBe(true);
      expect(probe.available('Arial')).toBe(true);
      expect(probe.measurements()).toBe(measured);
      await probe.install(null);
      expect(probe.cached('Narkisim')).toBeUndefined();
      expect(probe.cached('Arial')).toBe(true);
      write.mockRestore();
    } finally { probe.restore(); }
  });

  it('הרכבת 300 שמות אינה מודדת canvas וקוראת תוצאות רק אחרי בדיקה', async () => {
    const probe = await fontProbe(['Arial']);
    try {
      const families = Array.from({ length: 300 }, (_, i) => ({ value: `Family ${i}`, label: `Family ${i}` }));
      const installed = { families, hebrew: new Set<string>(), unresolved: new Set<string>(), source: 'host' as const };
      const slice = { options: [{ value: 'Aptos Display', label: 'Aptos Display' }] };
      const merged = probe.compose(slice, installed).families;
      expect(probe.measurements()).toBe(0);
      expect(merged.find((font) => font.value === 'Arial')?.available).toBeUndefined();
      expect(merged.find((font) => font.value === 'Family 299')?.hebrew).toBeUndefined();
      probe.available('Arial');
      probe.covers('Arial');
      probe.available('Aptos Display');
      const measured = probe.measurements();
      const refreshed = probe.compose(slice, installed).families;
      expect(refreshed.find((font) => font.value === 'Arial')).toMatchObject({ available: true, hebrew: true, measured: true });
      expect(refreshed.find((font) => font.value === 'Aptos Display')?.available).toBe(false);
      expect(probe.measurements()).toBe(measured);
    } finally { probe.restore(); }
  });
});
