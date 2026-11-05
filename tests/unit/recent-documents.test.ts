/**
 * „המסמכים האחרונים” — הרשימה שמחליטה מה נעלם מהמסך.
 *
 * שלוש השאלות שהבדיקות כאן מקבעות, וכולן שאלות על היעלמות שקטה:
 *
 * 1. **מה נופל בתקרה.** התקרה נספרת על הלא-מוצמדים בלבד; מוצמד אינו נזרק
 *    ואינו גונב מקום מהרשימה הרגילה.
 * 2. **מה קורה לקובץ שנפתח פעמיים.** שורה אחת שמתעדכנת, לא שתיים — ובלי
 *    שההצמדה תיפול בדרך.
 * 3. **מה שורד רשומה פגומה.** שורה שנכתבה חלקית נשמטת לבדה ואינה מפילה את
 *    שאר הרשימה.
 *
 * ולצדן: כל הפונקציות טהורות. הרשימה מוחזקת ב-state של המעטפת, ומוטציה בה
 * הייתה משנה מסך בלי שאיש ביקש לצייר מחדש.
 */
import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  MAX_RECENT_DOCUMENTS,
  filterRecents,
  forgetRecent,
  normalizeRecents,
  rememberRecent,
  setRecentPinned,
  sortedRecents,
  type RecentDocument,
} from '../../src/sessions/recent-documents';
import { loadRecentDocuments, saveRecentDocuments } from '../../src/host/settings';

function doc(patch: Partial<RecentDocument> = {}): RecentDocument {
  return { token: 'tok', name: 'א.docx', size: 1_000, openedAt: 5_000, writable: true, pinned: false, ...patch };
}

/** רשומה של `count` פתיחות, מהישנה לחדשה, מעל רשימה נתונה. */
function afterOpens(count: number, list: readonly RecentDocument[] = []): RecentDocument[] {
  let out = [...list];
  for (let i = 1; i <= count; i += 1) {
    out = rememberRecent(out, {
      token: `t${i}`, name: `${i}.docx`, size: i, openedAt: 10_000 + i, writable: true,
    });
  }
  return out;
}

/** עותק עמוק להשוואה אחרי קריאה — מוטציה על הקלט נתפסת בהשוואה אליו. */
function frozenCopy(list: readonly RecentDocument[]): RecentDocument[] {
  return JSON.parse(JSON.stringify(list)) as RecentDocument[];
}

describe('normalizeRecents', () => {
  /**
   * שורה שנכתבה לפני שהשדה קיים. נקראת כקריאה-בלבד ולא ככתיבה, וזו הנפילה
   * הבטוחה: „שמור בשם” על קובץ שאפשר לכתוב אליו הוא הטרדה, כתיבה ל-token
   * שאין עליו הרשאה היא כשל.
   */
  it('רשומה בלי writable נקראת כקריאה-בלבד', () => {
    const [item] = normalizeRecents([{ token: 't', name: 'א.docx', size: 10, openedAt: 1 }]);
    expect(item!.writable).toBe(false);
  });

  it('רק true מפורש הוא כתיב', () => {
    const list = normalizeRecents([
      { token: 'a', writable: true },
      { token: 'b', writable: 'yes' },
      { token: 'c', writable: 1 },
    ]);
    expect(list.map((item) => item.writable)).toEqual([true, false, false]);
  });

  it('קורא רשימה מלאה כפי שנשמרה', () => {
    const stored = [doc({ token: 'a' }), doc({ token: 'b', pinned: true, openedAt: 9 })];

    expect(normalizeRecents(JSON.parse(JSON.stringify(stored)))).toEqual(stored);
  });

  it('ערך שאינו מערך מחזיר רשימה ריקה', () => {
    expect(normalizeRecents(null)).toEqual([]);
    expect(normalizeRecents(undefined)).toEqual([]);
    expect(normalizeRecents({ tokens: ['a'] })).toEqual([]);
    expect(normalizeRecents('אחרונים')).toEqual([]);
  });

  it('שורה בלי token נשמטת ואינה פוסלת את השאר', () => {
    // שורה אחת שנכתבה חלקית אינה סיבה שמשתמש יאבד את כל הרשימה שלו.
    const read = normalizeRecents([
      { name: 'בלי-token.docx', size: 5, openedAt: 1 },
      null,
      'לא-אובייקט',
      { token: '', name: 'ריק' },
      doc({ token: 'שרד' }),
    ]);

    expect(read.map((item) => item.token)).toEqual(['שרד']);
  });

  it('token חוזר נשמט, והראשון נשמר', () => {
    // ה-token הוא הזהות: שתי שורות עליו היו נותנות שתי „הסרות” לאותו קובץ,
    // ואת אותה הצמדה בשני מצבים סותרים.
    const read = normalizeRecents([
      doc({ token: 'a', name: 'ראשון.docx' }),
      doc({ token: 'a', name: 'שני.docx', pinned: true }),
      doc({ token: 'b' }),
    ]);

    expect(read.map((item) => item.token)).toEqual(['a', 'b']);
    expect(read[0]?.name).toBe('ראשון.docx');
    expect(read[0]?.pinned).toBe(false);
  });

  it('שדות פגומים נופלים לברירת מחדל שפויה', () => {
    const read = normalizeRecents([
      { token: 'a', name: '', size: -7, openedAt: Number.NaN, pinned: 'כן' },
      { token: 'b', name: 42, size: 'הרבה', openedAt: -1, pinned: 1 },
    ]);

    expect(read).toEqual([
      { token: 'a', name: 'מסמך', size: 0, openedAt: 0, writable: false, pinned: false },
      { token: 'b', name: 'מסמך', size: 0, openedAt: 0, writable: false, pinned: false },
    ]);
  });

  it('הקריאה אינה מקצצת לתקרה', () => {
    // קריאה אינה המקום לאבד שורות: הפתיחה הבאה מקצצת ממילא, ושם נכנס פריט
    // חדש ונפילת הישן היא מה שמצופה.
    const stored = Array.from({ length: MAX_RECENT_DOCUMENTS + 5 }, (_, i) => doc({ token: `t${i}` }));

    expect(normalizeRecents(stored)).toHaveLength(MAX_RECENT_DOCUMENTS + 5);
  });
});

describe('rememberRecent', () => {
  it('מוסיף מסמך שלא היה', () => {
    const list = rememberRecent([], { token: 'a', name: 'א.docx', size: 12, openedAt: 700, writable: true });

    expect(list).toEqual([
      { token: 'a', name: 'א.docx', size: 12, openedAt: 700, writable: true, pinned: false },
    ]);
  });

  it('token שכבר ברשימה מתעדכן במקום ואינו מוכפל', () => {
    const before = [doc({ token: 'a', name: 'ישן.docx', size: 10, openedAt: 100 }), doc({ token: 'b' })];

    const after = rememberRecent(before, { token: 'a', name: 'חדש.docx', size: 90, openedAt: 900, writable: true });

    expect(after).toHaveLength(2);
    // השם והגודל נדרסים: הקובץ יכול היה להשתנות או להיות משונה שם, וה-token
    // הוא שמצביע עליו — השם הישן היה שקר על מה שייפתח בלחיצה.
    expect(after[0]).toEqual({
      token: 'a', name: 'חדש.docx', size: 90, openedAt: 900, writable: true, pinned: false,
    });
  });

  it('פתיחה חוזרת אינה מבטלת הצמדה', () => {
    // ההצמדה היא החלטה על הקובץ, לא נתון של הפתיחה.
    const before = [doc({ token: 'a', pinned: true, openedAt: 100 })];

    const after = rememberRecent(before, { token: 'a', name: 'א.docx', size: 1, openedAt: 900, writable: true });

    expect(after[0]?.pinned).toBe(true);
    expect(after[0]?.openedAt).toBe(900);
  });

  it('הרשומה שנפתחה עולה לראש הרשימה', () => {
    const before = [doc({ token: 'a' }), doc({ token: 'b' }), doc({ token: 'c' })];

    const after = rememberRecent(before, {
      token: 'c', name: 'ג.docx', size: 1, openedAt: 9_000, writable: true,
    });

    expect(after.map((item) => item.token)).toEqual(['c', 'a', 'b']);
  });

  it('התקרה מפילה את הישן ביותר, ורק אותו', () => {
    const list = afterOpens(MAX_RECENT_DOCUMENTS + 5);

    expect(list).toHaveLength(MAX_RECENT_DOCUMENTS);
    expect(list.map((item) => item.token)).toContain('t25');
    expect(list.map((item) => item.token), 'חמש הפתיחות הראשונות נפלו').not.toContain('t5');
    expect(list.map((item) => item.token)).toContain('t6');
  });

  it('מוצמד אינו נזרק בתקרה ואינו גונב מקום מהרשימה הרגילה', () => {
    // המוצמדים כאן הם הישנים ביותר בכל הרשימה (openedAt=1): אילו נספרו
    // בתקרה, הם היו הראשונים ליפול — וגם היו מקצרים את „האחרונים” בשלוש.
    const pins = [
      doc({ token: 'p1', pinned: true, openedAt: 1 }),
      doc({ token: 'p2', pinned: true, openedAt: 1 }),
      doc({ token: 'p3', pinned: true, openedAt: 1 }),
    ];

    const list = afterOpens(MAX_RECENT_DOCUMENTS + 5, pins);

    expect(list).toHaveLength(MAX_RECENT_DOCUMENTS + pins.length);
    expect(list.filter((item) => item.pinned).map((item) => item.token)).toEqual(['p1', 'p2', 'p3']);
    expect(list.filter((item) => !item.pinned)).toHaveLength(MAX_RECENT_DOCUMENTS);
  });

  it('רשומה בלי token תקין אינה נכנסת לרשימה', () => {
    // שורה שאי אפשר לפתוח אינה שווה מקום — וזה בדיוק מה שהקריאה עושה איתה.
    const before = [doc({ token: 'a' })];

    expect(
      rememberRecent(before, { token: '', name: 'א.docx', size: 1, openedAt: 9, writable: true }),
    ).toEqual(before);
  });

  it('מנרמל את מה שנרשם, כדי שהרשימה בזיכרון תהיה זו שתיקרא אחרי הפעלה מחדש', () => {
    const [added] = rememberRecent([], { token: 'a', name: '', size: -3, openedAt: 0, writable: false });

    expect(added).toEqual({ token: 'a', name: 'מסמך', size: 0, openedAt: 0, writable: false, pinned: false });
  });

  it('אינו משנה את הרשימה שנמסרה', () => {
    const before = [doc({ token: 'a', pinned: true }), doc({ token: 'b' })];
    const copy = frozenCopy(before);

    rememberRecent(before, { token: 'a', name: 'אחר.docx', size: 5, openedAt: 9_999, writable: true });
    rememberRecent(before, { token: 'ג', name: 'ג.docx', size: 5, openedAt: 9_999, writable: true });

    expect(before).toEqual(copy);
  });
});

describe('forgetRecent', () => {
  it('מסיר את המסמך המבוקש בלבד', () => {
    const before = [doc({ token: 'a' }), doc({ token: 'b' }), doc({ token: 'c' })];

    expect(forgetRecent(before, 'b').map((item) => item.token)).toEqual(['a', 'c']);
  });

  it('מסיר גם מוצמד', () => {
    // ההצמדה מגנה מפני שכחה אוטומטית, לא מפני בקשה מפורשת להסיר.
    const before = [doc({ token: 'a', pinned: true })];

    expect(forgetRecent(before, 'a')).toEqual([]);
  });

  it('token שאינו ברשימה אינו משנה דבר', () => {
    const before = [doc({ token: 'a' })];

    expect(forgetRecent(before, 'zzz')).toEqual(before);
  });

  it('אינו משנה את הרשימה שנמסרה', () => {
    const before = [doc({ token: 'a' }), doc({ token: 'b' })];
    const copy = frozenCopy(before);

    forgetRecent(before, 'a');

    expect(before).toEqual(copy);
  });
});

describe('setRecentPinned', () => {
  it('מצמיד ומשחרר', () => {
    const before = [doc({ token: 'a' }), doc({ token: 'b' })];

    const pinned = setRecentPinned(before, 'a', true);
    expect(pinned.map((item) => item.pinned)).toEqual([true, false]);

    expect(setRecentPinned(pinned, 'a', false).map((item) => item.pinned)).toEqual([false, false]);
  });

  it('token שאינו ברשימה אינו יוצר שורה', () => {
    // אין ממה לבנות אותה: שם, גודל וזמן אינם ידועים כאן.
    const before = [doc({ token: 'a' })];

    expect(setRecentPinned(before, 'zzz', true)).toEqual(before);
  });

  it('שחרור הצמדה אינו מוחק שורה אחרת מהמסך', () => {
    // אחרי השחרור יש `MAX+1` לא-מוצמדים. קיצוץ כאן היה מעלים שורה שהלחיצה
    // לא ביקשה לגעת בה; הקיצוץ שייך לפתיחה הבאה.
    const full = afterOpens(MAX_RECENT_DOCUMENTS, [doc({ token: 'p', pinned: true, openedAt: 1 })]);
    expect(full).toHaveLength(MAX_RECENT_DOCUMENTS + 1);

    const released = setRecentPinned(full, 'p', false);

    expect(released).toHaveLength(MAX_RECENT_DOCUMENTS + 1);
    expect(released.every((item) => !item.pinned)).toBe(true);
  });

  it('אינו משנה את הרשימה שנמסרה', () => {
    const before = [doc({ token: 'a' })];
    const copy = frozenCopy(before);

    setRecentPinned(before, 'a', true);

    expect(before).toEqual(copy);
  });
});

describe('sortedRecents', () => {
  it('מוצמדים ראשונים, ובכל קבוצה האחרון שנפתח בראש', () => {
    const list = [
      doc({ token: 'a', openedAt: 100 }),
      doc({ token: 'p-ישן', openedAt: 1, pinned: true }),
      doc({ token: 'b', openedAt: 300 }),
      doc({ token: 'p-חדש', openedAt: 50, pinned: true }),
    ];

    expect(sortedRecents(list).map((item) => item.token)).toEqual(['p-חדש', 'p-ישן', 'b', 'a']);
  });

  it('מיון יציב: זמן שווה שומר על סדר הרשימה', () => {
    const list = [
      doc({ token: 'a', openedAt: 7 }),
      doc({ token: 'b', openedAt: 7 }),
      doc({ token: 'c', openedAt: 7 }),
    ];

    expect(sortedRecents(list).map((item) => item.token)).toEqual(['a', 'b', 'c']);
  });

  it('אינו משנה את הרשימה שנמסרה', () => {
    const list = [doc({ token: 'a', openedAt: 1 }), doc({ token: 'b', openedAt: 9 })];
    const copy = frozenCopy(list);

    sortedRecents(list);

    expect(list).toEqual(copy);
  });
});

describe('filterRecents', () => {
  const list = [
    doc({ token: 'a', name: 'שולחן ערוך.docx' }),
    doc({ token: 'b', name: 'Report FINAL.docx' }),
    doc({ token: 'c', name: 'שו"ת הרמב"ם.docx' }),
  ];

  it('שאילתה ריקה מחזירה את הרשימה כמות שהיא', () => {
    // „לא הקלדתי כלום” אינו „לא נמצא דבר”.
    expect(filterRecents(list, '')).toEqual(list);
    expect(filterRecents(list, '   ')).toEqual(list);
  });

  it('חיפוש אינו רגיש לרישיות ומתעלם מרווחים בקצוות', () => {
    expect(filterRecents(list, '  final  ').map((item) => item.token)).toEqual(['b']);
  });

  it('מוצא מילה באמצע השם', () => {
    // שמות קבצים מתחילים לעתים בתאריך או במספר סימן, והמילה שזוכרים באמצע.
    expect(filterRecents(list, 'ערוך').map((item) => item.token)).toEqual(['a']);
  });

  it('גרשיים עבריים וישרים הם אותו חיפוש', () => {
    // מקלדת עברית מייצרת ״ ו-׳, ומערכת הקבצים מלאה ב-" וב-'. בלי האיחוד הזה
    // חיפוש ראשי תיבות — הנפוץ ביותר בספרייה תורנית — לא היה מוצא.
    expect(filterRecents(list, 'שו״ת').map((item) => item.token)).toEqual(['c']);
    expect(filterRecents(list, 'הרמב"ם').map((item) => item.token)).toEqual(['c']);

    const hebrewName = [doc({ token: 'd', name: 'רש״י על התורה.docx' })];
    expect(filterRecents(hebrewName, 'רש"י')).toHaveLength(1);
    expect(filterRecents(hebrewName, 'רש״י')).toHaveLength(1);
  });

  it('אין התאמה — רשימה ריקה', () => {
    expect(filterRecents(list, 'משנה ברורה')).toEqual([]);
  });

  it('שומר על סדר הרשימה ואינו משנה אותה', () => {
    const copy = frozenCopy(list);

    expect(filterRecents(list, '.docx').map((item) => item.token)).toEqual(['a', 'b', 'c']);
    expect(list).toEqual(copy);
  });
});

describe('הגשר ל-storage', () => {
  function hostReturns(data: unknown): ReturnType<typeof vi.fn> {
    const call = vi.fn(async () => ({ success: true, data, error: null }));
    window.Otzaria = { call } as never;
    return call;
  }

  afterEach(() => {
    delete (window as Partial<Window>).Otzaria;
  });

  it('הקריאה מחזירה גולמי — הפירוש אינו בגשר', async () => {
    const call = hostReturns([{ token: 'a' }, 'זבל']);

    await expect(loadRecentDocuments()).resolves.toEqual([{ token: 'a' }, 'זבל']);
    expect(call).toHaveBeenCalledWith('storage.get', { key: 'recent-documents' });
  });

  it('מחוץ לאוצריא אין היסטוריה שמורה', async () => {
    await expect(loadRecentDocuments()).resolves.toBeNull();
  });

  it('הכתיבה שולחת עותק של הרשימה', async () => {
    const call = hostReturns(true);
    const list = [doc({ token: 'a' })];

    await saveRecentDocuments(list);

    expect(call).toHaveBeenCalledWith('storage.set', { key: 'recent-documents', value: list });
    const sent = call.mock.calls[0]?.[1] as { value: RecentDocument[] };
    expect(sent.value, 'לא אותו מערך — מה שיוצא לגשר אינו ה-state של המעטפת').not.toBe(list);
  });
});
