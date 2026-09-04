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
    expect(parseOtzariaLink(buildRefHref({ ...hit, id: null }, 'פסחים לד'))).toEqual({
      kind: 'detection',
      query: 'פסחים לד',
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
