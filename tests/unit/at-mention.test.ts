import { describe, expect, it } from 'vitest';
import {
  buildLinkText,
  buildRefHref,
  isQueryable,
  parseAtTrigger,
  suggestionSubtitle,
  type ResolvedRefHit,
} from '../../src/engine/at-mention';

function hit(overrides: Partial<ResolvedRefHit> = {}): ResolvedRefHit {
  return {
    id: 42,
    bookId: 'פסחים',
    bookUid: 'id:42',
    title: 'פסחים',
    reference: 'פסחים דף לד',
    index: 1234,
    isPdf: false,
    isSourceLine: true,
    isUserBook: false,
    bookPath: 'ש"ס, בבלי',
    ...overrides,
  };
}

describe('parseAtTrigger', () => {
  it('תופס אזכור בתחילת פסקה', () => {
    expect(parseAtTrigger('@פסחים לד')).toEqual({ query: 'פסחים לד', atIndex: 0 });
  });

  it('תופס אזכור אחרי רווח, ומחזיר את היסט ה-@ עצמו', () => {
    const trigger = parseAtTrigger('כמו שכתוב ב@פסחים לד');
    expect(trigger).toEqual({ query: 'פסחים לד', atIndex: 11 });
  });

  it('תופס אזכור אחרי אות שימוש — הצורה הטבעית בעברית', () => {
    // "ראה ב@פסחים לד" הוא איך שכותבים את זה בפועל; אות השימוש נשארת במסמך
    // ורק ה-@ מוחלף.
    expect(parseAtTrigger('ראה ב@פסחים לד')).toEqual({ query: 'פסחים לד', atIndex: 5 });
    expect(parseAtTrigger('כמובא ל@רמב"ם')?.query).toBe('רמב"ם');
    expect(parseAtTrigger('ו@ברכות')?.query).toBe('ברכות');
  });

  it('אות שימוש נספרת רק כשהיא עומדת לבדה', () => {
    // "כתוב@" — ב' היא סופה של מילה שלמה, לא אות שימוש.
    expect(parseAtTrigger('כתוב@פסחים')).toBeNull();
  });

  it('אינו תופס כתובת דוא"ל', () => {
    // הבעיה שהבדיקה הזו מגנה עליה: כל הקלדה של כתובת הייתה פותחת רשימת הצעות.
    expect(parseAtTrigger('שלחו אלי dev@example.com')).toBeNull();
    expect(parseAtTrigger('dev@exa')).toBeNull();
  });

  it('תופס אזכור אחרי סוגר או מירכאה פותחת', () => {
    expect(parseAtTrigger('(@פסחים')?.query).toBe('פסחים');
    expect(parseAtTrigger('"@פסחים')?.query).toBe('פסחים');
  });

  it('לוקח את ה-@ האחרון', () => {
    expect(parseAtTrigger('@ברכות ב, @פסחים')).toEqual({ query: 'פסחים', atIndex: 10 });
  });

  it('@ בלבד הוא טריגר פתוח, אך אינו בשל לשאילתה', () => {
    const trigger = parseAtTrigger('@');
    expect(trigger).toEqual({ query: '', atIndex: 0 });
    expect(isQueryable(trigger!)).toBe(false);
  });

  it('נסגר על שבירת שורה', () => {
    expect(parseAtTrigger('@פסחים\nלד')).toBeNull();
  });

  it('נסגר על רווח כפול — המשפט המשיך', () => {
    expect(parseAtTrigger('@פסחים  לד')).toBeNull();
  });

  it('נסגר כשההפניה ארוכה מדי', () => {
    expect(parseAtTrigger(`@${'א'.repeat(49)}`)).toBeNull();
    expect(parseAtTrigger(`@${'א'.repeat(48)}`)).not.toBeNull();
  });

  it('נסגר על יותר מדי מילים', () => {
    expect(parseAtTrigger('@תלמוד ירושלמי עירובין פרק ו הלכה ז')).not.toBeNull();
    expect(parseAtTrigger('@א ב ג ד ה ו ז ח')).toBeNull();
  });

  it('רווח נגרר נשמר בטריגר אך אינו מונע שאילתה', () => {
    const trigger = parseAtTrigger('@פסחים ');
    expect(trigger?.query).toBe('פסחים ');
    expect(isQueryable(trigger!)).toBe(true);
  });

  it('סף השאילתה הוא שני תווים', () => {
    expect(isQueryable({ query: 'פ', atIndex: 0 })).toBe(false);
    expect(isQueryable({ query: 'פס', atIndex: 0 })).toBe(true);
    expect(isQueryable({ query: ' פ ', atIndex: 0 })).toBe(false);
  });
});

describe('buildRefHref', () => {
  it('בונה קישור עומק לספר טקסט', () => {
    expect(buildRefHref(hit(), 'פסחים לד')).toBe('otzaria://open/book/42?index=1234');
  });

  it('בונה קישור PDF עם מספר עמוד', () => {
    expect(buildRefHref(hit({ isPdf: true, index: 17 }), 'x')).toBe(
      'otzaria://open/pdf/42?index=17',
    );
  });

  it('עמוד PDF לעולם אינו קטן מ-1 — הראוטר דוחה 0', () => {
    expect(buildRefHref(hit({ isPdf: true, index: 0 }), 'x')).toBe(
      'otzaria://open/pdf/42?index=1',
    );
  });

  it('ספר אישי נופל לאיתור מקורות ולא לקישור עומק', () => {
    // user_books.db מקצה מזהים באותו טווח כמו ספריית הבסיס: קישור לפי id
    // היה נפתר לספר אחר לגמרי.
    expect(buildRefHref(hit({ isUserBook: true }), 'הערות שלי')).toBe(
      'otzaria://open/detection?q=%D7%94%D7%A2%D7%A8%D7%95%D7%AA%20%D7%A9%D7%9C%D7%99',
    );
  });

  it('היעדר id נופל לאיתור מקורות', () => {
    expect(buildRefHref(hit({ id: null }), 'ספר סרוק ג')).toContain('open/detection?q=');
    expect(buildRefHref(hit({ id: undefined }), 'ספר סרוק ג')).toContain('open/detection?q=');
  });

  it('הפניה מקודדת ב-UTF-8, כפי שהראוטר מצפה', () => {
    const href = buildRefHref(hit({ id: null }), '  בראשית א  ');
    expect(href).toBe('otzaria://open/detection?q=%D7%91%D7%A8%D7%90%D7%A9%D7%99%D7%AA%20%D7%90');
    expect(decodeURIComponent(new URL(href).searchParams.get('q') ?? '')).toBe('בראשית א');
  });
});

describe('buildLinkText', () => {
  it('מעדיף את ההפניה שנפתרה על מה שהוקלד', () => {
    expect(buildLinkText(hit(), 'פסחים לד')).toBe('פסחים דף לד');
  });

  it('נופל לכותרת ואז למה שהוקלד', () => {
    expect(buildLinkText(hit({ reference: '   ' }), 'פסחים לד')).toBe('פסחים');
    expect(buildLinkText(hit({ reference: '', title: '' }), ' פסחים לד ')).toBe('פסחים לד');
  });
});

describe('suggestionSubtitle', () => {
  it('מציג את נתיב הקטגוריה', () => {
    expect(suggestionSubtitle(hit())).toBe('ש"ס, בבלי');
  });

  it('מסמן התאמה ברמת פרק בלבד', () => {
    expect(suggestionSubtitle(hit({ isSourceLine: false }))).toBe('ש"ס, בבלי · לרמת הפרק');
  });

  it('מסמן ספר אישי ו-PDF', () => {
    expect(suggestionSubtitle(hit({ isUserBook: true }))).toContain('ספר אישי');
    expect(suggestionSubtitle(hit({ isPdf: true, isSourceLine: false }))).toBe('ש"ס, בבלי · PDF');
  });

  it('ריק כשאין מה לומר', () => {
    expect(suggestionSubtitle(hit({ bookPath: '' }))).toBe('');
  });
});
