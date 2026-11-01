import { describe, expect, it, vi } from 'vitest';
import { parseOtzariaLink } from '../../src/engine/otzaria-link';
import { createOtzariaLinkActivation } from '../../src/engine/otzaria-link-activation';
import { buildRefHref, type ResolvedRefHit } from '../../src/engine/at-mention';

describe('parseOtzariaLink', () => {
  it('מפענח קישור לספר טקסט', () => {
    expect(parseOtzariaLink('otzaria://open/book/42?index=1234')).toEqual({
      kind: 'book',
      id: 42,
      index: 1234,
    });
  });

  it('מפענח קישור ל-PDF', () => {
    expect(parseOtzariaLink('otzaria://open/pdf/8?index=17')).toEqual({
      kind: 'pdf',
      id: 8,
      index: 17,
    });
  });

  it('index חסר נופל לברירת המחדל של הראוטר', () => {
    expect(parseOtzariaLink('otzaria://open/book/42')).toMatchObject({ index: 0 });
    // ב-PDF העמוד הוא 1-based, ו-0 נדחה בראוטר.
    expect(parseOtzariaLink('otzaria://open/pdf/8')).toMatchObject({ index: 1 });
    expect(parseOtzariaLink('otzaria://open/pdf/8?index=0')).toMatchObject({ index: 1 });
    expect(parseOtzariaLink('otzaria://open/book/42?index=abc')).toMatchObject({ index: 0 });
    expect(parseOtzariaLink('otzaria://open/book/42?index=-3')).toMatchObject({ index: 0 });
  });

  it('מפענח איתור מקורות ומפרק את הקידוד', () => {
    expect(
      parseOtzariaLink('otzaria://open/detection?q=%D7%A4%D7%A1%D7%97%D7%99%D7%9D%20%D7%9C%D7%93'),
    ).toEqual({ kind: 'detection', query: 'פסחים לד' });
  });

  /**
   * הצורות ש-book_link_builder.dart של אוצריא כותב ב„העתק קישור” של תפריט
   * הקורא. בלעדיהן הקישור נפתח במקטע הנכון אבל בלי ההדגשה (Otzaria#3822).
   */
  it('קישור עם הדגשת טקסט: m נקרא ומפוענח', () => {
    expect(
      parseOtzariaLink('otzaria://open/book/42?index=7&m=%D7%95%D7%99%D7%90%D7%9E%D7%A8%20%D7%94%27'),
    ).toEqual({ kind: 'book', id: 42, index: 7, markText: "ויאמר ה'" });
  });

  it('קישור עם הדגשת המקטע: mark הוא דגל גם בלי ערך', () => {
    expect(parseOtzariaLink('otzaria://open/book/42?index=7&mark')).toEqual({
      kind: 'book',
      id: 42,
      index: 7,
      markSection: true,
    });
    expect(parseOtzariaLink('otzaria://open/book/42?index=7&mark=')).toMatchObject({
      markSection: true,
    });
  });

  it('m ריק או רווחים בלבד אינו הדגשה', () => {
    expect(parseOtzariaLink('otzaria://open/book/42?index=7&m=%20%20')).toEqual({
      kind: 'book',
      id: 42,
      index: 7,
    });
  });

  it('source של הקישור מתורגם לערכי חוזה התוספים', () => {
    expect(parseOtzariaLink('otzaria://open/book/42?source=user&index=7&mark')).toEqual({
      kind: 'book',
      id: 42,
      index: 7,
      source: 'user',
      markSection: true,
    });
    expect(
      parseOtzariaLink('otzaria://open/book/42?source=db%3Amy-lib&index=7&m=%D7%90'),
    ).toEqual({ kind: 'book', id: 42, index: 7, source: 'attached', markText: 'א' });
    expect(parseOtzariaLink('otzaria://open/book/42?source=official')).toMatchObject({
      source: 'library',
    });
    expect(parseOtzariaLink('otzaria://open/pdf/8?source=user&index=3')).toEqual({
      kind: 'pdf',
      id: 8,
      index: 3,
      source: 'user',
    });
  });

  it('בלי source, או עם ערך שהראוטר אינו מכיר, השדה אינו נוצר', () => {
    expect(parseOtzariaLink('otzaria://open/book/42?index=7')).not.toHaveProperty('source');
    expect(parseOtzariaLink('otzaria://open/book/42?source=weird')).not.toHaveProperty('source');
    expect(parseOtzariaLink('otzaria://open/book/42?source=db%3A')).not.toHaveProperty('source');
  });

  it('בקישור PDF הדגשה אינה נקראת — כמו בראוטר', () => {
    expect(parseOtzariaLink('otzaria://open/pdf/8?index=3&m=%D7%90&mark')).toEqual({
      kind: 'pdf',
      id: 8,
      index: 3,
    });
  });

  it('פוסל מזהה שאינו מספר חיובי', () => {
    expect(parseOtzariaLink('otzaria://open/book/0')).toBeNull();
    expect(parseOtzariaLink('otzaria://open/book/-1')).toBeNull();
    expect(parseOtzariaLink('otzaria://open/book/abc')).toBeNull();
  });

  it('פוסל איתור מקורות בלי שאילתה', () => {
    expect(parseOtzariaLink('otzaria://open/detection')).toBeNull();
    expect(parseOtzariaLink('otzaria://open/detection?q=%20')).toBeNull();
  });

  it('פוסל מה שאינו קישור אוצריא', () => {
    expect(parseOtzariaLink('https://example.com')).toBeNull();
    expect(parseOtzariaLink('mailto:a@b.c')).toBeNull();
    expect(parseOtzariaLink('')).toBeNull();
    expect(parseOtzariaLink('לא כתובת בכלל')).toBeNull();
  });

  it('פוסל יעד של אוצריא שאיננו מכירים — הוא באחריותה', () => {
    expect(parseOtzariaLink('otzaria://library/reindex')).toBeNull();
    expect(parseOtzariaLink('otzaria://open/settings')).toBeNull();
  });

  it('סוגר מעגל מול buildRefHref — מה שנכתב הוא מה שנקרא', () => {
    const hit: ResolvedRefHit = {
      id: 42,
      bookId: 'פסחים',
      title: 'פסחים',
      reference: 'פסחים דף לד',
      index: 1234,
      isPdf: false,
      isSourceLine: true,
      isUserBook: false,
      bookPath: '',
    };
    expect(parseOtzariaLink(buildRefHref(hit, 'פסחים לד'))).toEqual({
      kind: 'book',
      id: 42,
      index: 1234,
    });
    expect(parseOtzariaLink(buildRefHref({ ...hit, isPdf: true, index: 3 }, 'x'))).toEqual({
      kind: 'pdf',
      id: 42,
      index: 3,
    });
    // ההפניה שנפתרה ולא מה שהוקלד — ראו at-mention.ts:buildRefHref.
    expect(parseOtzariaLink(buildRefHref({ ...hit, id: null }, 'פסחים לד'))).toEqual({
      kind: 'detection',
      query: 'פסחים דף לד',
    });
  });

  it('קורא את המזהה היציב, וסוגר עליו מעגל מול buildRefHref', () => {
    const hit: ResolvedRefHit = {
      id: 42,
      bookId: 'פסחים',
      bookUid: 'id:42',
      title: 'פסחים',
      reference: 'פסחים דף לד',
      index: 1234,
      isPdf: false,
      isSourceLine: true,
      isUserBook: false,
      bookPath: '',
    };
    expect(parseOtzariaLink(buildRefHref(hit, 'פסחים לד'))).toEqual({
      kind: 'book',
      id: 42,
      index: 1234,
      uid: 'id:42',
    });
  });

  it('uid עם תווים שדורשים קידוד חוזר כפי שנכתב', () => {
    expect(parseOtzariaLink('otzaria://open/book/5?index=0&uid=uid%3A5')).toEqual({
      kind: 'book',
      id: 5,
      index: 0,
      uid: 'uid:5',
    });
  });

  it('קישור ישן בלי uid נשאר תקף — השדה אינו תנאי', () => {
    expect(parseOtzariaLink('otzaria://open/book/42?index=7')).toEqual({
      kind: 'book',
      id: 42,
      index: 7,
    });
  });

  it('uid ריק אינו נספר', () => {
    expect(parseOtzariaLink('otzaria://open/pdf/3?index=2&uid=%20%20')).toEqual({
      kind: 'pdf',
      id: 3,
      index: 2,
    });
  });
});

describe('createOtzariaLinkActivation', () => {
  it('מנווט ובולע את הלחיצה', () => {
    const navigate = vi.fn();
    const activate = createOtzariaLinkActivation({ navigate });

    expect(activate({ href: 'otzaria://open/book/42?index=7' })).toEqual({ type: 'suppress' });
    expect(navigate).toHaveBeenCalledWith({ kind: 'book', id: 42, index: 7 });
  });

  it('קישור אחר ממשיך להתנהג כרגיל', () => {
    const navigate = vi.fn();
    const activate = createOtzariaLinkActivation({ navigate });

    expect(activate({ href: 'https://example.com' })).toEqual({ type: 'default' });
    expect(activate({ href: null })).toEqual({ type: 'default' });
    expect(navigate).not.toHaveBeenCalled();
  });

  it('כשל בניווט אינו מחזיר את הקישור לדפדפן', () => {
    // `suppress` גם בכשל: ברירת המחדל הייתה מנסה לפתוח otzaria:// בדפדפן,
    // וזה נכשל בשקט ב-WebView.
    const onStatus = vi.fn();
    const activate = createOtzariaLinkActivation({
      navigate: () => {
        throw new Error('boom');
      },
      onStatus,
    });

    expect(activate({ href: 'otzaria://open/book/42' })).toEqual({ type: 'suppress' });
    expect(onStatus).toHaveBeenCalledWith('פתיחת הקישור נכשלה', true);
  });
});
