/**
 * אפשרויות הגופן. מה שנבדק כאן הוא הדבר שהיה שבור: הרשימה הייתה קשיחה, ולכן
 * גופן שהמסמך משתמש בו ולא ניחשנו מראש לא היה בבורר בכלל.
 *
 * וגם הכיוון ההפוך — הגופן הארוז שלנו וגופני העברית של אוצריא אינם מוכרים
 * למנוע (הם מותקנים אחרי שהוא בנה את הרשימה שלו), ולכן מיזוג שנשען על המנוע
 * לבדו היה מוחק אותם.
 */
import { describe, it, expect, vi } from 'vitest';
import type { FontFamilyOption } from 'superdoc/ui';
import {
  FONT_GROUP_ALL,
  FONT_GROUP_HEBREW,
  FONT_GROUP_RECENT,
  FONT_GROUP_TOP,
  RECENT_FONT_LIMIT,
  LATIN_FONT_FAMILIES,
  OTZARIA_FONT_FAMILIES,
  composeFontOptions,
  fallbackFontOptions,
  mergeFontFamilies,
  mergeFontSizes,
  observeFontSlice,
  readFontSlice,
  type FontFamilyChoice,
  type FontOptionsSource,
  type FontsSliceLike,
} from '../../src/engine/font-options';
import type { InstalledFontsSnapshot } from '../../src/engine/system-fonts';

const values = (options: readonly { value: string }[]) => options.map((option) => option.value);

/**
 * מנייה מזויפת. `hebrew` הוא מה שקובע לאיזו קבוצה שם נופל, ו-`unresolved` הוא
 * מה שהמנייה סימנה כמותקן שהדפדפן אינו פותר — כלומר מי שהמיזוג ימדוד מחדש.
 */
function installed(
  names: readonly string[],
  hebrew: readonly string[] = [],
  unresolved: readonly string[] = [],
): InstalledFontsSnapshot {
  return {
    families: names.map((name) => ({ value: name, label: name, previewFamily: `"${name}", sans-serif` })),
    hebrew: new Set(hebrew.map((name) => name.toLowerCase())),
    unresolved: new Set(unresolved.map((name) => name.toLowerCase())),
    source: 'host',
  };
}

/** הקבוצה שאפשרות נפלה אליה. */
const groupOf = (options: readonly { value: string; group: string }[], value: string) =>
  options.find((option) => option.value === value)?.group;

describe('mergeFontFamilies', () => {
  it('הגופנים שלנו ראשונים, גם כשהמנוע מציע אחרים', () => {
    const merged = mergeFontFamilies([{ value: 'Verdana', label: 'Verdana' }]);
    expect(values(merged).slice(0, OTZARIA_FONT_FAMILIES.length)).toEqual(
      values(OTZARIA_FONT_FAMILIES),
    );
    expect(values(merged)).toContain('Verdana');
  });

  it('גופני המסמך מגיעים מהמנוע, בקבוצה משלהם אחרי הקבועים', () => {
    // הקבועים — שלנו וזנב הלטינית — הם „מה שתמיד שם”, ולכן הם קבוצה אחת בראש
    // ולא שתיים שהמנוע חוצץ ביניהן. גופני המסמך מקבלים כותרת משלהם.
    const merged = mergeFontFamilies([{ value: 'Cambria', label: 'Cambria' }]);
    const names = values(merged);
    expect(names.indexOf('Cambria')).toBeGreaterThan(names.indexOf('Arial'));
    expect(groupOf(merged, 'Cambria')).toBe(FONT_GROUP_RECENT);
  });

  it('אין כפילות, וההשוואה חסרת רגישות לאותיות כמו במנוע', () => {
    const merged = values(
      mergeFontFamilies([
        { value: 'assistant', label: 'assistant' },
        { value: 'ARIAL', label: 'ARIAL' },
      ]),
    );
    expect(merged.filter((value) => value.toLowerCase() === 'assistant')).toHaveLength(1);
    expect(merged.filter((value) => value.toLowerCase() === 'arial')).toHaveLength(1);
  });

  it('התווית שלנו מנצחת — „David” ולא „TaameyDavidCLM”', () => {
    const merged = mergeFontFamilies([{ value: 'TaameyDavidCLM', label: 'TaameyDavidCLM' }]);
    expect(merged.find((option) => option.value === 'TaameyDavidCLM')?.label).toBe('David');
  });

  it('דגימה מוצעת רק כאשר הזמינות ידועה', () => {
    const merged = mergeFontFamilies([{ value: 'Verdana', label: 'Verdana' }]);
    expect(merged.filter((option) => option.available === true).every((option) => Boolean(option.previewFamily))).toBe(true);
    expect(merged.find((option) => option.value === 'Verdana')).toMatchObject({ available: undefined, measured: false });
    expect(merged.find((option) => option.value === 'Verdana')?.previewFamily).toBeUndefined();
  });

  it('שורה פגומה מהמנוע נדחית ואינה מפילה את הרשימה', () => {
    const broken = [
      { value: '', label: 'ריק' },
      { value: '   ', label: 'רווחים' },
      null,
    ] as unknown as readonly FontFamilyOption[];
    expect(values(mergeFontFamilies(broken))).toEqual(
      values([...OTZARIA_FONT_FAMILIES, ...LATIN_FONT_FAMILIES]),
    );
  });
});

describe('mergeFontFamilies — מה שמותקן במכונה', () => {
  it('גופני המכונה מגיעים אחרי הרשימה הקבועה ואחרי גופני המסמך', () => {
    // הסדר הוא כל העניין: מי שפותח מסמך עברי צריך למצוא את Frank Ruhl בשורה
    // הראשונה, לא אחרי 300 שמות שהמכונה מדווחת עליהם.
    const merged = values(
      mergeFontFamilies([{ value: 'Cambria', label: 'Cambria' }], installed(['Narkisim'], ['Narkisim'])),
    );
    expect(merged.indexOf('Narkisim')).toBeGreaterThan(merged.indexOf('Cambria'));
    expect(merged.indexOf('Narkisim')).toBeGreaterThan(merged.indexOf('Arial'));
  });

  it('עברית לפני השאר, וכל אחת בקבוצה שלה', () => {
    const merged = mergeFontFamilies(undefined, installed(['Bahnschrift', 'Narkisim'], ['Narkisim']));
    const names = values(merged);
    expect(names.indexOf('Narkisim')).toBeLessThan(names.indexOf('Bahnschrift'));
    expect(groupOf(merged, 'Narkisim')).toBe(FONT_GROUP_HEBREW);
    expect(groupOf(merged, 'Bahnschrift')).toBe(FONT_GROUP_ALL);
  });

  it('הרשימה הקבועה כולה בקבוצה העליונה, בלי כותרת', () => {
    const merged = mergeFontFamilies([{ value: 'Cambria', label: 'Cambria' }], installed(['Narkisim']));
    for (const name of ['Assistant', 'Arial']) {
      expect(groupOf(merged, name)).toBe(FONT_GROUP_TOP);
    }
  });

  it('גופן שכבר בראש אינו חוזר תחת „כל הגופנים”', () => {
    // Arial מותקן כמעט תמיד, והמארח ידווח עליו. הופעה שנייה למטה הייתה גורמת
    // למשתמש לבחור בו מהמקום הלא נכון ולראות ערך זהה פעמיים בבורר.
    const merged = mergeFontFamilies(undefined, installed(['Arial', 'Assistant'], ['Arial', 'Assistant']));
    expect(values(merged).filter((value) => value === 'Arial')).toHaveLength(1);
    expect(groupOf(merged, 'Arial')).toBe(FONT_GROUP_TOP);
    expect(groupOf(merged, 'Assistant')).toBe(FONT_GROUP_TOP);
  });

  it('בלי מנייה הרשימה זהה למה שהייתה לפניה', () => {
    // רגרסיה: המנייה נוחתת אחרי האתחול, וכל עוד היא לא נחתה הבורר חייב להיראות
    // בדיוק כמו קודם.
    expect(values(mergeFontFamilies([{ value: 'Cambria', label: 'Cambria' }]))).toEqual(
      values([...OTZARIA_FONT_FAMILIES, ...LATIN_FONT_FAMILIES, { value: 'Cambria' }]),
    );
  });
});

/**
 * גופן שהמכונה מדווחת עליו ושהדפדפן אינו מצייר.
 *
 * זה הצירוף שהסתיר 43 מ-287 גופנים **מותקנים** במכונה שנמדדה, חמישה מהם
 * עבריים: המנייה מחקה אותם, והבורר לא ידע עליהם. עכשיו הם מגיעים — ומה שנבדק
 * כאן הוא שהם מגיעים **מסומנים** ולא מוכרזים זמינים, כי שורה שמכריזה שם
 * ומציירת גופן אחר היא בדיוק מה שהדגל קיים כדי למנוע.
 */
describe('mergeFontFamilies — מותקן שאינו מצוייר', () => {
  const find = (options: readonly FontFamilyChoice[], value: string) =>
    options.find((option) => option.value === value);

  it('שם שהמנייה סימנה נמדד, ומגיע לא-זמין ובלי דגימה', () => {
    const merged = mergeFontFamilies(
      undefined,
      installed(['David', 'Guttman Kav-Light'], ['David', 'Guttman Kav-Light'], [
        'Guttman Kav-Light',
      ]),
      () => true,
      // המכונה פותרת את David ולא את השני — בדיוק מה שנמדד בכרום.
      (family) => family === 'David',
    );

    const flagged = find(merged, 'Guttman Kav-Light');
    expect(flagged?.available).toBe(false);
    expect(flagged?.measured).toBe(true);
    expect(flagged?.previewFamily).toBeUndefined();
    // והוא בקבוצת „עברית”, כלומר נגיש — לא מוסתר בזנב ולא נעלם.
    expect(flagged?.group).toBe(FONT_GROUP_HEBREW);
  });

  it('שמות מהמארח נשארים לא בדוקים עד שיש תוצאת מטמון', () => {
    const merged = mergeFontFamilies(undefined, installed(['David', 'Narkisim'], ['David']));
    expect(find(merged, 'David')).toMatchObject({ measured: false, available: undefined, hebrew: true });
    expect(find(merged, 'Narkisim')).toMatchObject({ measured: false, available: undefined, hebrew: undefined });
  });

  it('„מותקן ואינו מצוייר” מסומן בנפרד מ„אינו מותקן”', () => {
    // שתי השורות לא-זמינות, ושתיהן צריכות לומר למשתמש דברים שונים: `Aptos`
    // באמת אינו במכונה, ו-`Guttman Kav-Light` מותקן והדפדפן אינו פותר אותו.
    const merged = mergeFontFamilies(
      undefined,
      installed(['Guttman Kav-Light'], ['Guttman Kav-Light'], ['Guttman Kav-Light']),
      () => true,
      () => false,
    );

    expect(find(merged, 'Guttman Kav-Light')?.installedNotDrawable).toBe(true);
    expect(find(merged, 'Aptos')?.available).toBe(false);
    expect(find(merged, 'Aptos')?.installedNotDrawable).toBeUndefined();
  });

  it('אחרי שהבייטים הוזרקו — המדידה מחזירה זמין, והדגימה חוזרת', () => {
    // זו הדרך שבה השורה מתעוררת אחרי `ensureFamilyDrawable`: אותה מנייה
    // בדיוק, ומכונה שכבר פותרת את השם.
    const snapshot = installed(['Guttman Kav-Light'], ['Guttman Kav-Light'], [
      'Guttman Kav-Light',
    ]);
    const before = mergeFontFamilies(undefined, snapshot, () => true, () => false);
    const after = mergeFontFamilies(undefined, snapshot, () => true, () => true);

    expect(find(before, 'Guttman Kav-Light')?.available).toBe(false);
    expect(find(after, 'Guttman Kav-Light')?.available).toBe(true);
    expect(find(after, 'Guttman Kav-Light')?.previewFamily).toBeDefined();
  });
});

/**
 * המכסה על „אחרונים”. מה שנבדק כאן הוא הצד שקל לשבור בלי לשים לב: לא שהקבוצה
 * מתקצרת — זה הקל — אלא שמה שנחתך ממנה **עדיין נבחר**. גופן שהמסמך נכתב בו
 * ואינו מותקן במכונה קיים בבורר רק דרך רשימת המנוע, ומכסה שמוחקת אותו מוחקת
 * את היכולת לבחור בו בכלל.
 */
describe('mergeFontFamilies — המכסה על „אחרונים”', () => {
  const many = (count: number) =>
    Array.from({ length: count }, (_, index) => ({ value: `Doc${index}`, label: `Doc${index}` }));

  const inGroup = (options: readonly { value: string; group: string }[], group: string) =>
    options.filter((option) => option.group === group).map((option) => option.value);

  it('הקבוצה מציגה עד המכסה, ולא עשרים שם', () => {
    const merged = mergeFontFamilies(many(20));
    expect(inGroup(merged, FONT_GROUP_RECENT)).toHaveLength(RECENT_FONT_LIMIT);
  });

  it('מה שנחתך אינו נעלם — הוא נבחר מ„כל הגופנים”', () => {
    const merged = mergeFontFamilies(many(20));
    expect(values(merged)).toContain('Doc19');
    expect(groupOf(merged, 'Doc19')).toBe(FONT_GROUP_ALL);
  });

  it('עודף שמכסה עברית נוחת ב„עברית”, ולא ב„כל הגופנים”', () => {
    // הדגל והדגימה היו נכונים גם קודם; הקבוצה הכריזה את ההפך ממה שהשורה
    // מראה. המקור האחרון היה `[engine, FONT_GROUP_ALL]` בלי פיצול לפי כיסוי,
    // בשונה מהמותקנים.
    const engine = [...many(RECENT_FONT_LIMIT), { value: 'Narkisim', label: 'Narkisim' }];
    const merged = mergeFontFamilies(engine, installed([]), (family) => family === 'Narkisim');

    expect(groupOf(merged, 'Narkisim')).toBe(FONT_GROUP_HEBREW);
    expect(inGroup(merged, FONT_GROUP_RECENT)).not.toContain('Narkisim');
  });

  it('ועודף שאינו מכסה נשאר ב„כל הגופנים”', () => {
    const engine = [...many(RECENT_FONT_LIMIT), { value: 'Cambria', label: 'Cambria' }];
    const merged = mergeFontFamilies(engine, installed([]), () => false);
    expect(groupOf(merged, 'Cambria')).toBe(FONT_GROUP_ALL);
  });

  it('העודף העברי צמוד למותקנים העבריים — כותרת „עברית” אחת ולא שתיים', () => {
    // `buildComboRows` פותח כותרת בכל **החלפה** של קבוצה, ולכן מקור עברי
    // שנכנס אחרי „כל הגופנים” היה מייצר כותרת שנייה באמצע הרשימה.
    const engine = [...many(RECENT_FONT_LIMIT), { value: 'Narkisim', label: 'Narkisim' }];
    const merged = mergeFontFamilies(
      engine,
      installed(['Gisha', 'Bahnschrift'], ['Gisha']),
      (family) => family === 'Narkisim',
    );

    const groups = merged.map((option) => option.group);
    const starts = groups.filter((group, index) => group !== groups[index - 1]);
    expect(starts.filter((group) => group === FONT_GROUP_HEBREW)).toHaveLength(1);
  });

  it('המכסה סופרת מה שנוסף, לא מה שנסרק', () => {
    // מסמך שכתוב בגופנים שלנו: חמשת הראשונים ברשימת המנוע כבר בראש הבורר,
    // ומכסה שסופרת אותם הייתה מציגה קבוצה כמעט ריקה.
    const engine = [
      ...OTZARIA_FONT_FAMILIES.map((font) => ({ value: font.value, label: font.label })),
      ...many(RECENT_FONT_LIMIT),
    ];
    expect(inGroup(mergeFontFamilies(engine), FONT_GROUP_RECENT)).toHaveLength(RECENT_FONT_LIMIT);
  });
});

/**
 * הדגל שהבורר מצייר לפיו דגימה של אותיות עבריות. מה שנבדק כאן הוא בדיוק מה
 * שאי אפשר לגזור מהקבוצה: שש המשפחות שלנו יושבות ב-`FONT_GROUP_TOP` יחד עם
 * הלטינית, וגופן שמכסה עברית יכול לשבת שם גם הוא (Arial).
 */
describe('mergeFontFamilies — כיסוי עברית', () => {
  const hebrewOf = (options: readonly { value: string; hebrew: boolean | undefined }[], value: string) =>
    options.find((option) => option.value === value)?.hebrew;

  it('שש המשפחות שלנו מסומנות עבריות גם בלי מנייה', () => {
    // הן מוזרקות אחרי שהמנייה רצה, ודגימה שנעלמת לשנייה ואז חוזרת גרועה משתי
    // האפשרויות — לכן הן אינן נשאלות בכלל.
    const merged = mergeFontFamilies(undefined);
    for (const font of OTZARIA_FONT_FAMILIES) {
      expect(hebrewOf(merged, font.value)).toBe(true);
    }
  });

  it('מה שבקבוצת „עברית” מסומן, ומה שבזנב אינו', () => {
    const merged = mergeFontFamilies(undefined, installed(['Bahnschrift', 'Narkisim'], ['Narkisim']));
    expect(hebrewOf(merged, 'Narkisim')).toBe(true);
    expect(hebrewOf(merged, 'Bahnschrift')).toBeUndefined();
  });

  it('גופן בראש הרשימה מסומן לפי המנייה ולא לפי הקבוצה', () => {
    // Arial בזנב הלטינית ובקבוצה העליונה, אבל הוא באמת מכסה עברית — ודגימה בו
    // אמיתית בדיוק כמו ב-Frank Ruhl. Aptos, שאינו מכסה, נשאר בלי.
    const merged = mergeFontFamilies(undefined, installed(['Aptos', 'Arial'], ['Arial']));
    expect(groupOf(merged, 'Arial')).toBe(FONT_GROUP_TOP);
    expect(hebrewOf(merged, 'Arial')).toBe(true);
    expect(hebrewOf(merged, 'Aptos')).toBeUndefined();
  });

  it('גופן שהמסמך משתמש בו מסומן לפי המנייה', () => {
    const engine = [{ value: 'Narkisim', label: 'Narkisim' }, { value: 'Cambria', label: 'Cambria' }];
    const merged = mergeFontFamilies(engine, installed(['Narkisim'], ['Narkisim']));
    expect(hebrewOf(merged, 'Narkisim')).toBe(true);
    expect(hebrewOf(merged, 'Cambria')).toBeUndefined();
  });

  it('בלי מנייה ובלי מדידה גופן לטיני אינו מסומן — דגימה שנופלת ל-fallback היא שקר', () => {
    const merged = mergeFontFamilies([{ value: 'Cambria', label: 'Cambria' }]);
    expect(hebrewOf(merged, 'Cambria')).toBeUndefined();
    expect(hebrewOf(merged, 'Aptos')).toBeUndefined();
  });

  it('מדידה מסמנת גם גופן שהמנייה אינה מכירה', () => {
    // התקלה שדווחה: „יש גופנים עם עברית שאינם מוצגים ככאלה”. המנייה יודעת רק
    // על מה שכתוב ברשימת המועמדים שלה; המדידה עונה על כל שם.
    const merged = mergeFontFamilies(
      [{ value: 'Gisha', label: 'Gisha' }],
      installed([]),
      (family) => family === 'Gisha',
    );
    expect(hebrewOf(merged, 'Gisha')).toBe(true);
  });

  it('מדידה קובעת גם את הקבוצה, לא רק את הדגימה', () => {
    const merged = mergeFontFamilies(undefined, installed(['Gisha', 'Bahnschrift']), (family) => family === 'Gisha');
    expect(groupOf(merged, 'Gisha')).toBe(FONT_GROUP_HEBREW);
    expect(groupOf(merged, 'Bahnschrift')).toBe(FONT_GROUP_ALL);
  });
});

describe('coverage cache states', () => {
  it('distinguishes a measured negative from unknown host metadata', () => {
    const options = mergeFontFamilies(undefined, installed(['Unknown', 'Latin', 'Hebrew']),
      (name) => name === 'Unknown' ? undefined : name === 'Hebrew');
    expect(options.find((font) => font.value === 'Unknown')?.hebrew).toBeUndefined();
    expect(options.find((font) => font.value === 'Latin')?.hebrew).toBe(false);
    expect(options.find((font) => font.value === 'Hebrew')?.hebrew).toBe(true);
  });
});

describe('mergeFontSizes', () => {
  it('גדלים ממוינים כמספרים ולא כמחרוזות', () => {
    // מיון לקסיקוגרפי היה מציב את 8 אחרי 72.
    const merged = values(mergeFontSizes([{ value: '72', label: '72' }])).map(Number);
    expect(merged).toEqual([...merged].sort((a, b) => a - b));
    expect(merged[0]).toBe(8);
  });

  it('גודל שהמנוע מציע ואינו בסולם Word מתווסף במקומו', () => {
    const merged = values(mergeFontSizes([{ value: '10.5', label: '10.5' }]));
    expect(merged).toContain('10.5');
    expect(merged.indexOf('10.5')).toBe(merged.indexOf('11') - 1);
  });

  it('אין כפילות של אותו מספר בכתיב אחר', () => {
    const merged = values(mergeFontSizes([{ value: '12.0', label: '12.0' }]));
    expect(merged.filter((value) => Number.parseFloat(value) === 12)).toHaveLength(1);
  });
});

describe('readFontSlice', () => {
  it('קורא מ-ui.fonts', () => {
    const ui: FontOptionsSource = {
      fonts: {
        getFamilyOptions: () => [{ value: 'Cambria', label: 'Cambria' }],
        getSizeOptions: () => [{ value: '13', label: '13' }],
      },
    };
    const options = composeFontOptions(readFontSlice(ui));
    expect(values(options.families)).toContain('Cambria');
    expect(values(options.sizes)).toContain('13');
  });

  it('מנוע בלי fonts נופל לרשימה שלנו ולא לרשימה ריקה', () => {
    expect(composeFontOptions(readFontSlice({}))).toEqual(fallbackFontOptions());
    expect(composeFontOptions(readFontSlice(null))).toEqual(fallbackFontOptions());
    expect(fallbackFontOptions().families.length).toBeGreaterThan(0);
  });

  it('קריאה שזורקת אינה מפילה את הבורר', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const ui: FontOptionsSource = {
      fonts: {
        getFamilyOptions: () => {
          throw new Error('המנוע לא מוכן');
        },
      },
    };
    expect(values(composeFontOptions(readFontSlice(ui)).families)).toEqual(
      values(fallbackFontOptions().families),
    );
    expect(warn).toHaveBeenCalled();
    warn.mockRestore();
  });
});

describe('observeFontSlice', () => {
  it('כל דיווח של המנוע מגיע כמות שהוא, גם השני', () => {
    // מחזיק ולא משתנה מקומי: TS מצמצם `let` שמשויך רק בתוך callback ל-null.
    const captured: { emit: ((slice: FontsSliceLike) => void) | null } = { emit: null };
    const dispose = vi.fn();
    const ui: FontOptionsSource = {
      fonts: {
        observe: (listener) => {
          captured.emit = listener;
          listener({ options: [{ value: 'Aptos', label: 'Aptos' }], sizeOptions: [] });
          return dispose;
        },
      },
    };

    const seen: string[][] = [];
    const stop = observeFontSlice(ui, (slice) =>
      seen.push([...values(composeFontOptions(slice).families)]),
    );

    // המנוע פותר את גופני המסמך אחרי הפתיחה; בלי ההאזנה הבורר היה קופא על
    // הרשימה של הרגע הראשון.
    captured.emit?.({ options: [{ value: 'Cambria', label: 'Cambria' }], sizeOptions: [] });

    expect(seen).toHaveLength(2);
    expect(seen[0]).toContain('Aptos');
    expect(seen[1]).toContain('Cambria');
    expect(seen[1]).toContain('Assistant');

    stop();
    expect(dispose).toHaveBeenCalled();
  });

  it('מנוע בלי observe מדווח פעם אחת ומחזיר disposer', () => {
    const seen: number[] = [];
    const stop = observeFontSlice({}, (slice) =>
      seen.push(composeFontOptions(slice).families.length),
    );
    expect(seen).toHaveLength(1);
    expect(seen[0]).toBeGreaterThan(0);
    expect(() => stop()).not.toThrow();
  });
});

describe('composeFontOptions — שני מקורות שנוחתים בזמנים שונים', () => {
  it('מנייה שנוחתת אחרי שהמסמך נפתח אינה מוחקת את גופני המסמך', () => {
    // זה כל הטעם שבגללו App.vue מחזיק שני refs ומרכיב, במקום לדחוף לאחד:
    // המקור שנחת שני היה מוחק את מה שהראשון הביא.
    const slice = { options: [{ value: 'Cambria', label: 'Cambria' }], sizeOptions: [] };
    const after = composeFontOptions(slice, installed(['Narkisim'], ['Narkisim']));
    expect(values(after.families)).toContain('Cambria');
    expect(values(after.families)).toContain('Narkisim');
  });

  it('מסמך שנפתח אחרי שהמנייה נחתה אינו מוחק את גופני המכונה', () => {
    const snapshot = installed(['Narkisim'], ['Narkisim']);
    const before = composeFontOptions(null, snapshot);
    const after = composeFontOptions({ options: [{ value: 'Cambria', label: 'Cambria' }] }, snapshot);
    expect(values(before.families)).toContain('Narkisim');
    expect(values(after.families)).toContain('Narkisim');
    expect(values(after.families)).toContain('Cambria');
  });

  it('שני סדרי הנחיתה מגיעים לאותה תוצאה בדיוק', () => {
    const slice = { options: [{ value: 'Cambria', label: 'Cambria' }] };
    const snapshot = installed(['Narkisim', 'Verdana'], ['Narkisim']);
    expect(composeFontOptions(slice, snapshot)).toEqual(composeFontOptions(slice, snapshot));
  });
});

describe('mergeFontFamilies — תוויות כפולות', () => {
  it('„David” פעמיים מובהר — זה מה שנצפה בבורר', () => {
    // TaameyDavidCLM נושא את התווית „David”, והמכונה מדווחת גם על David של
    // Windows. שתיהן לגיטימיות, ולמשתמש הן נראו כתקלה.
    const merged = mergeFontFamilies(undefined, installed(['David'], ['David']));
    const labels = merged.map((option) => option.label);
    expect(labels.filter((label) => label === 'David')).toHaveLength(1);
    expect(labels).toContain('David (TaameyDavidCLM)');
  });

  it('שם מערכתי נשאר נקי — לא „David (David)”', () => {
    const merged = mergeFontFamilies(undefined, installed(['David'], ['David']));
    expect(merged.find((option) => option.value === 'David')?.label).toBe('David');
  });

  it('בלי התנגשות התוויות אינן משתנות', () => {
    const merged = mergeFontFamilies(undefined, installed(['Narkisim'], ['Narkisim']));
    expect(merged.find((option) => option.value === 'TaameyDavidCLM')?.label).toBe('David');
  });
});

describe('observeFontSlice — דיווח חוזר וזהה', () => {
  it('המנוע מדווח על כל תזוזת סמן; רק שינוי אמיתי עובר', () => {
    // ה-slice נבנה מחדש בכל `recompute`, כולל `host-selection`. בלי השער
    // הבורר היה נבנה מחדש בכל הקשה — עם מאות משפחות ברשימה.
    const captured: { emit: ((slice: FontsSliceLike) => void) | null } = { emit: null };
    const ui: FontOptionsSource = {
      fonts: {
        observe: (listener) => {
          captured.emit = listener;
          return () => {};
        },
      },
    };

    let reports = 0;
    observeFontSlice(ui, () => { reports += 1; });

    const same = () => ({ options: [{ value: 'Arial', label: 'Arial' }], sizeOptions: [{ value: '12', label: '12' }] });
    captured.emit?.(same());
    expect(reports).toBe(1);

    // אובייקטים חדשים לגמרי, אותם ערכים — בדיוק מה שהמנוע מייצר.
    captured.emit?.(same());
    captured.emit?.(same());
    expect(reports).toBe(1);

    captured.emit?.({ options: [{ value: 'Cambria', label: 'Cambria' }], sizeOptions: [] });
    expect(reports).toBe(2);
  });

  it('שינוי בתווית בלבד עובר — המנוע פותר את תוויות המסמך אחרי הפתיחה', () => {
    // חתימה על הערך לבדו הייתה בולעת אותו, והבורר היה נשאר על התווית הגולמית.
    const captured: { emit: ((slice: FontsSliceLike) => void) | null } = { emit: null };
    const ui: FontOptionsSource = {
      fonts: {
        observe: (listener) => {
          captured.emit = listener;
          return () => {};
        },
      },
    };

    let reports = 0;
    observeFontSlice(ui, () => { reports += 1; });

    captured.emit?.({ options: [{ value: 'TaameyDavidCLM', label: 'TaameyDavidCLM' }] });
    expect(reports).toBe(1);

    captured.emit?.({ options: [{ value: 'TaameyDavidCLM', label: 'David' }] });
    expect(reports).toBe(2);
  });
});

/**
 * זמינות הגופן — הבאג שנמדד, והגבול שנשמר.
 *
 * הטריגר אינו תיאורטי: `scripts/qa/font-availability-qa.mjs` מנה 67 שורות
 * שהבורר מציג בכרום אמיתי, ומצא ש-`Aptos` — ברירת המחדל של Word 365 — אינה
 * נפתרת במכונה בלי Office. השורה הכריזה „Aptos” וציירה Arial, ופס הדגימה
 * הראה את הטקסט של המשתמש בגודל שבמסמך תחת שם של גופן אחר.
 *
 * שני חצאים נבדקים כאן, והשני חשוב כמו הראשון:
 *
 * 1. **שהמדידה קורית** למי שאין לו מזריק, ושהיא מפילה את הציור ולא את השורה.
 * 2. **שהיא אינה קורית** לכל השאר. השש שלנו מוזרקות אחרי המנייה, והמנייה
 *    עצמה כבר סוננה — מדידה שם הייתה מבהבת או עבודה כפולה על מאות שמות.
 *
 * ב-jsdom אין canvas, ולכן `isFamilyAvailable` האמיתית מחזירה `true` תמיד
 * („בלי canvas אין לנו מה לומר”). כלומר בלי הכפיל המסלול הזה אינו נבדק בכלל,
 * וזו הסיבה שכל בדיקה כאן מזריקה `resolves`.
 */
describe('mergeFontFamilies — זמינות', () => {
  /** כפיל: כל שם נפתר, למעט מי שנמסר כשבור. */
  const resolvesAllBut = (broken: readonly string[]) => {
    const set = new Set(broken.map((name) => name.toLowerCase()));
    return (family: string) => !set.has(family.trim().toLowerCase());
  };

  const find = (options: readonly FontFamilyChoice[], value: string) =>
    options.find((option) => option.value === value);

  it('גופן לטיני שאינו נפתר מסומן לא-זמין ומאבד את הציור', () => {
    const merged = mergeFontFamilies(undefined, undefined, undefined, resolvesAllBut(['Aptos']));
    const aptos = find(merged, 'Aptos');

    expect(aptos?.available).toBe(false);
    // וזה הלב: `previewFamily` נופל, ולכן שלושת מקומות הציור בפקד — השורה,
    // התיבה ופס הדגימה — מפסיקים לצייר בו־זמנית.
    expect(aptos?.previewFamily).toBeUndefined();
  });

  it('והשורה נשארת — מסמך נודד, והגופן עשוי להיות מותקן במכונה אחרת', () => {
    const merged = mergeFontFamilies(undefined, undefined, undefined, resolvesAllBut(['Aptos']));
    expect(values(merged)).toContain('Aptos');
    expect(groupOf(merged, 'Aptos')).toBe(FONT_GROUP_TOP);
  });

  it('גופן לטיני שכן נפתר שומר על הציור', () => {
    const merged = mergeFontFamilies(undefined, undefined, undefined, resolvesAllBut(['Aptos']));
    const arial = find(merged, 'Arial');

    expect(arial?.available).toBe(true);
    expect(arial?.previewFamily).toBe('Arial, sans-serif');
  });

  it('ששת הגופנים שלנו אינם נמדדים — הם מוזרקות אחרי המנייה', () => {
    // הכפיל אומר שאף אחד מהם אינו קיים. זה בדיוק המצב בשנייה הראשונה של
    // ההפעלה (נמדד: `FrankRuhlCLM=false`), ואסור שהוא יפיל את הדגימה —
    // הגופן יהיה שם בעוד רגע, ודגימה שנעלמת וחוזרת גרועה משתי האפשרויות.
    const names = values(OTZARIA_FONT_FAMILIES);
    const merged = mergeFontFamilies(undefined, undefined, undefined, resolvesAllBut(names));

    for (const option of OTZARIA_FONT_FAMILIES) {
      const row = find(merged, option.value);
      expect(row?.available, option.value).toBe(true);
      expect(row?.previewFamily, option.value).toBe(option.previewFamily);
    }
  });

  it('תשובת מטמון שלילית חלה גם על שם שדווח כמותקן', () => {
    const merged = mergeFontFamilies(undefined, installed(['Narkisim'], ['Narkisim']),
      () => true, resolvesAllBut(['Narkisim']));
    expect(find(merged, 'Narkisim')).toMatchObject({ available: false, measured: true, installedNotDrawable: true });
    expect(find(merged, 'Narkisim')?.previewFamily).toBeUndefined();
  });

  it('גופן מסמך שאינו נפתר מסומן — זה המקור שהבאג באמת חי בו', () => {
    // הגשש שגילה את הבאג לא פתח מסמך, ולכן 67 השורות שהוא מנה לא כללו אפילו
    // גופן מסמך אחד — כלומר המקור המסוכן ביותר לא נמדד. מסמך Word 365 נושא
    // `Aptos` **ו-`Aptos Display`**, והשנייה מגיעה מכאן בלבד.
    const merged = mergeFontFamilies(
      [{ value: 'Aptos Display', label: 'Aptos Display' }],
      undefined,
      () => false,
      resolvesAllBut(['Aptos Display']),
    );
    const row = find(merged, 'Aptos Display');

    expect(row?.available).toBe(false);
    expect(row?.previewFamily).toBeUndefined();
    expect(values(merged)).toContain('Aptos Display');
  });

  it('וגופן מסמך שכן נפתר — למשל דרך אליאס עברי — נשאר מצויר', () => {
    // `planFontAliases` מזריק `@font-face` לגופן עברי חסר, ואז הוא **נפתר**
    // דרך התחליף. הדגימה בו נכונה ואינה פשרה: גוף המסמך מצויר באותו תחליף.
    const merged = mergeFontFamilies(
      [{ value: 'Guttman Yad', label: 'Guttman Yad' }],
      undefined,
      () => true,
      () => true,
    );
    const row = find(merged, 'Guttman Yad');

    expect(row?.available).toBe(true);
    expect(row?.previewFamily).toBe('Guttman Yad');
  });

  it('Aptos מגיע גם מהמנוע — והשורה נמדדת פעם אחת, לא משתי דרכים', () => {
    /*
     * התרחיש שנמדד בפועל: מסמך שברירת המחדל שלו Aptos, כלומר המנוע **מדווח**
     * עליו. השורה נתפסת ברשימת הלטינית שיושבת לפני המנוע ב-`sources`, ולכן
     * לפני שגם המנוע נמדד ההגנה כאן הייתה סדר הרשומות בלבד — והערת התיעוד של
     * `LATIN_FONT_FAMILIES` דווקא מזמינה להזיז אותה („נשארים כזנב אחרי
     * אפשרויות המנוע”). הזזה כזאת הייתה מחזירה את הבאג ומשאירה הכול ירוק.
     */
    const merged = mergeFontFamilies(
      [{ value: 'Aptos', label: 'Aptos', previewFamily: "'Aptos', sans-serif" }],
      undefined,
      undefined,
      resolvesAllBut(['Aptos']),
    );
    const row = find(merged, 'Aptos');

    expect(row?.available).toBe(false);
    expect(row?.previewFamily).toBeUndefined();
    expect(values(merged).filter((value) => value === 'Aptos')).toHaveLength(1);
  });

  it('קורא מטמון עבור כל מקור, בלי להפוך תשובה חסרה לשלילית', () => {
    const cached = vi.fn((_family: string) => undefined);
    const many = Array.from({ length: 300 }, (_, index) => `Family ${index}`);
    const merged = composeFontOptions({ options: [{ value: 'Cambria', label: 'Cambria' }] },
      installed(many), cached, cached).families;
    expect(merged).toHaveLength(311);
    expect(merged.filter((option) => option.available === undefined)).toHaveLength(305);
    expect(merged.filter((option) => option.hebrew === undefined)).toHaveLength(305);
  });

  it('בלי canvas אין טענה שגופן מותקן או חסר', () => {
    const merged = mergeFontFamilies([{ value: 'Cambria', label: 'Cambria' }]);
    const cambria = find(merged, 'Cambria');
    expect(cambria?.available).toBeUndefined();
    expect(cambria?.previewFamily).toBeUndefined();
    expect(cambria?.measured).toBe(false);
  });
});
