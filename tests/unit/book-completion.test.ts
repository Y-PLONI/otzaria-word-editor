/**
 * מנוע ההתאמה של "השלמה מהספר" (engine/book-completion.ts): פונקציות טהורות,
 * בלי DOM ובלי RPC. הטקסט כאן מנוקד — בדיוק כמו טקסט אמיתי מהספר — וההקשר
 * המוקלד תמיד בלי ניקוד, כמו שמשתמש מקליד בפועל.
 */
import { describe, it, expect } from 'vitest';
import {
  buildSectionCache,
  matchAtCursor,
  matchBookTitle,
  sliceWords,
  stripDiacritics,
} from '../../src/engine/book-completion';

const VERSE =
  'בְּרֵאשִׁ֖ית בָּרָ֣א אֱלֹהִ֑ים אֵ֥ת הַשָּׁמַ֖יִם וְאֵ֥ת הָאָֽרֶץ';

describe('buildSectionCache', () => {
  it('מנרמלת ניקוד וטעמים ומפרקת למילים', () => {
    const cache = buildSectionCache(VERSE);
    expect(cache.normalizedText).toBe(stripDiacritics(VERSE));
    expect(cache.words).toHaveLength(7);
  });
});

describe('sliceWords', () => {
  it('גוזרת מהמקור, כולל ניקוד, ומחזירה את המילה הבאה בתור', () => {
    const cache = buildSectionCache(VERSE);
    const slice = sliceWords(cache, 0, 2);
    expect(slice?.text).toBe('בְּרֵאשִׁ֖ית בָּרָ֣א');
    expect(slice?.nextWordIndex).toBe(2);
  });

  it('null כשאין עוד מילים', () => {
    const cache = buildSectionCache(VERSE);
    expect(sliceWords(cache, 7, 5)).toBeNull();
  });
});

describe('matchAtCursor', () => {
  it('משלימה מילה חלקית בלי הקשר, מהתחלת ה-section', () => {
    const cache = buildSectionCache(VERSE);
    const match = matchAtCursor(cache, { precedingWords: [], partialWord: 'ברא' });

    expect(match?.text).toBe('בְּרֵאשִׁ֖ית בָּרָ֣א אֱלֹהִ֑ים אֵ֥ת הַשָּׁמַ֖יִם');
    expect(match?.matchedWordIndex).toBe(0);
    expect(match?.nextWordIndex).toBe(5);
  });

  /**
   * הרעש שנמדד בשטח: „ח” אחרי „קנה די” (שאינו בספר) התאים למילה הראשונה
   * בעמוד שמתחילה ב-ח, וההצעה שהוצגה הייתה טקסט אקראי.
   */
  it('אות בודדת בלי הקשר תואם אינה מספיקה כדי להציע', () => {
    const cache = buildSectionCache(VERSE);
    expect(matchAtCursor(cache, { precedingWords: [], partialWord: 'א' })).toBeNull();
    expect(matchAtCursor(cache, { precedingWords: ['לא', 'בספר'], partialWord: 'א' })).toBeNull();
  });

  it('הקשר אחד שתואם מספיק גם לאות בודדת', () => {
    const cache = buildSectionCache(VERSE);
    const match = matchAtCursor(cache, { precedingWords: ['בראשית'], partialWord: 'ב' });

    expect(match?.matchedWordIndex).toBe(1);
  });

  it('משתמשת בהקשר כדי לאתר את המיקום הנכון באמצע המשפט', () => {
    const cache = buildSectionCache(VERSE);
    const match = matchAtCursor(cache, {
      precedingWords: ['בראשית', 'ברא'],
      partialWord: 'אלה',
    });

    expect(match?.matchedWordIndex).toBe(2);
    expect(match?.text).toBe('אֱלֹהִ֑ים אֵ֥ת הַשָּׁמַ֖יִם וְאֵ֥ת הָאָֽרֶץ');
    expect(match?.nextWordIndex).toBe(7);
  });

  /**
   * `contextWordsUsed` הוא מה שמאפשר ל-overlay להחליף גם את מילות ההקשר
   * שהמשתמש הקליד בלי ניקוד — אחרת רק המילה האחרונה יוצאת מנוקדת.
   */
  it('מדווחת כמה מילות הקשר נוצלו בפועל', () => {
    const cache = buildSectionCache(VERSE);
    const two = matchAtCursor(cache, { precedingWords: ['בראשית', 'ברא'], partialWord: 'אלה' });
    expect(two?.contextWordsUsed).toBe(2);

    const none = matchAtCursor(cache, { precedingWords: [], partialWord: 'ברא' });
    expect(none?.contextWordsUsed).toBe(0);
  });

  it('בגבול מילה (בלי partial) דורשת הקשר מלא ומציעה את מה שנשאר', () => {
    const cache = buildSectionCache(VERSE);
    const match = matchAtCursor(cache, {
      precedingWords: ['בראשית', 'ברא', 'אלהים'],
      partialWord: '',
    });

    expect(match?.matchedWordIndex).toBe(3);
    expect(match?.text).toBe('אֵ֥ת הַשָּׁמַ֖יִם וְאֵ֥ת הָאָֽרֶץ');
  });

  it('null כשההקלדה לא נמצאת בטקסט', () => {
    const cache = buildSectionCache(VERSE);
    expect(matchAtCursor(cache, { precedingWords: ['לא', 'קיים'], partialWord: 'כלל' })).toBeNull();
  });

  it('null בגבול מילה בלי שום הקשר', () => {
    const cache = buildSectionCache(VERSE);
    expect(matchAtCursor(cache, { precedingWords: [], partialWord: '' })).toBeNull();
  });
});

describe('matchBookTitle', () => {
  const TITLE = 'משנה ברורה';
  const REF = 'סימן א';

  it('משלימה את שארית השם ואת המיקום', () => {
    const match = matchBookTitle(TITLE, REF, { precedingWords: [], partialWord: 'מש' });
    expect(match?.completionText).toBe('נה ברורה סימן א');
  });

  it('ממשיכה ממילה שנייה חלקית', () => {
    const match = matchBookTitle(TITLE, REF, { precedingWords: ['משנה'], partialWord: 'ברו' });
    expect(match?.completionText).toBe('רה סימן א');
  });

  it('שם מלא מוקלד — מוסיפה רק את המיקום', () => {
    const match = matchBookTitle(TITLE, REF, { precedingWords: ['משנה', 'ברורה'], partialWord: '' });
    expect(match?.completionText).toBe(' סימן א');
  });

  it('null כשאין מיקום ושם הספר כבר הושלם', () => {
    const match = matchBookTitle(TITLE, null, { precedingWords: ['משנה', 'ברורה'], partialWord: '' });
    expect(match).toBeNull();
  });

  it('null כשההקלדה לא תואמת את שם הספר', () => {
    expect(matchBookTitle(TITLE, REF, { precedingWords: ['לא'], partialWord: 'קשור' })).toBeNull();
  });
});
