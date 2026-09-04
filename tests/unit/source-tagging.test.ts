/**
 * מנוע תיוג המקורות (engine/source-tagging.ts): פונקציות טהורות. המקרה
 * שנבדק הכי הרבה הוא השלילי — "פסחים יג" בלי סימון עמוד לא הופך לקישור —
 * כי זו ההגבלה שבשבילה כל המודול הזה נכתב.
 */
import { describe, it, expect } from 'vitest';
import {
  findOpenTagToken,
  matchBookNames,
  findBookNameInQuery,
  parsePageMarker,
  canonicalDafText,
  buildTocIndex,
  resolveSourceTag,
  buildSourceTagHref,
  type BookNameMatch,
} from '../../src/engine/source-tagging';

describe('findOpenTagToken', () => {
  it('מאתרת @ פתוח בסוף הטקסט', () => {
    expect(findOpenTagToken('שלום @יומא יג')).toEqual({ atOffset: 5, query: 'יומא יג' });
  });

  it('null כשאין @ בכלל', () => {
    expect(findOpenTagToken('שלום עולם')).toBeNull();
  });

  it('רווח בודד בתוך ה-query (בין שם הספר לסימון העמוד) לא סוגר את ה-token', () => {
    expect(findOpenTagToken('@יומא יג:')).toEqual({ atOffset: 0, query: 'יומא יג:' });
  });

  it('null על רווח כפול — המשתמש כבר המשיך בפרוזה רגילה', () => {
    expect(findOpenTagToken('@יומא  יג')).toBeNull();
  });

  it('null כשיש שבר שורה אחרי ה-@', () => {
    expect(findOpenTagToken('@יומא\nיג')).toBeNull();
  });

  it('@ בתחילת הטקסט', () => {
    expect(findOpenTagToken('@')).toEqual({ atOffset: 0, query: '' });
  });
});

describe('matchBookNames', () => {
  const books = [
    { bookId: 'יומא', title: 'יומא' },
    { bookId: 'יבמות', title: 'יבמות' },
    { bookId: 'בראשית', title: 'בראשית' },
  ];

  it('מתאימה לפי תחילית, בלי תלות בניקוד', () => {
    const matches = matchBookNames(books, 'יומ');
    expect(matches).toHaveLength(1);
    expect(matches[0]?.bookId).toBe('יומא');
    expect(matches[0]?.restOfTitle).toBe('א');
  });

  it('כמה התאמות ממוינות לפי אורך כותרת', () => {
    const matches = matchBookNames(books, 'י');
    expect(matches.map((m) => m.bookId)).toEqual(['יומא', 'יבמות']);
  });

  it('מערך ריק כש-query ריק', () => {
    expect(matchBookNames(books, '')).toEqual([]);
  });
});

describe('findBookNameInQuery', () => {
  const books = [
    { bookId: 'יומא', title: 'יומא' },
    { bookId: 'בבא קמא', title: 'בבא קמא' },
    { bookId: 'בבא', title: 'בבא' },
  ];

  it('מוצאת שם ספר כתחילית של query ארוך יותר', () => {
    const match = findBookNameInQuery(books, 'יומא יג:');
    expect(match?.bookId).toBe('יומא');
  });

  it('בוחרת את הכותרת הארוכה יותר כששתיהן תואמות ("בבא קמא" על פני "בבא")', () => {
    const match = findBookNameInQuery(books, 'בבא קמא יג:');
    expect(match?.bookId).toBe('בבא קמא');
  });

  it('null כשה-query הוא בדיוק שם הספר, בלי המשך', () => {
    expect(findBookNameInQuery(books, 'יומא')).toBeNull();
  });

  it('null כשאין שם ספר תואם בכלל', () => {
    expect(findBookNameInQuery(books, 'זזזזז יג:')).toBeNull();
  });
});

describe('parsePageMarker', () => {
  it.each([
    ['יג.', { letters: 'יג', amud: 'a' }],
    ['יג:', { letters: 'יג', amud: 'b' }],
    ["יג א'", { letters: 'יג', amud: 'a' }],
    ["יג ב'", { letters: 'יג', amud: 'b' }],
    ['יג ע"א', { letters: 'יג', amud: 'a' }],
    ['יג ע"ב', { letters: 'יג', amud: 'b' }],
  ] as const)('%s מזוהה כסימון עמוד תקין', (raw, expected) => {
    const marker = parsePageMarker(raw);
    expect(marker?.letters).toBe(expected.letters);
    expect(marker?.amud).toBe(expected.amud);
  });

  it('null בלי שום סימון עמוד — זו ההגבלה המרכזית', () => {
    expect(parsePageMarker('יג')).toBeNull();
  });

  it('null על מספר ערבי בלבד', () => {
    expect(parsePageMarker('13:')).toBeNull();
  });
});

describe('canonicalDafText', () => {
  it('תואם בדיוק לצורת ה-tocEntry שנמדדה באוצריא', () => {
    expect(canonicalDafText('יג', 'a')).toBe('דף יג.');
    expect(canonicalDafText('יג', 'b')).toBe('דף יג:');
  });
});

describe('resolveSourceTag', () => {
  const yomaMatch: BookNameMatch = {
    bookId: 'יומא',
    title: 'יומא',
    matchedPrefix: 'יומא',
    restOfTitle: '',
  };
  const toc = buildTocIndex([
    { text: 'יומא', index: 0, level: 0 },
    { text: 'דף יג.', index: 24, level: 1 },
    { text: 'דף יג:', index: 25, level: 1 },
  ]);

  it('"יומא יג:" נפתר לאינדקס ה-toc הנכון', () => {
    const resolved = resolveSourceTag('יומא יג:', yomaMatch, toc);
    expect(resolved?.tocIndex).toBe(25);
  });

  it('"יומא יג" (בלי סימון) — אין קישור', () => {
    expect(resolveSourceTag('יומא יג', yomaMatch, toc)).toBeNull();
  });

  it('"יומאיג:" (בלי רווח מפריד) — אין קישור', () => {
    expect(resolveSourceTag('יומאיג:', yomaMatch, toc)).toBeNull();
  });

  it('דף שלא קיים ב-toc — אין קישור', () => {
    expect(resolveSourceTag('יומא ק:', yomaMatch, toc)).toBeNull();
  });

  it("סימון עם ע\"א/ע\"ב נפתר לאותו index כמו הנקודה/נקודתיים", () => {
    const resolved = resolveSourceTag('יומא יג ע"ב', yomaMatch, toc);
    expect(resolved?.tocIndex).toBe(25);
  });
});

describe('buildSourceTagHref', () => {
  it('בונה URL בפורמט otzaria:// הרשמי', () => {
    expect(buildSourceTagHref(1234, 57)).toBe('otzaria://open/book/1234?index=57');
  });
});
