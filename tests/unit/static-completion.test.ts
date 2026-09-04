/**
 * השלמה ממילונים סטטיים (engine/static-completion.ts): הטעינה העצלה (אותה
 * תבנית כמו spellcheck-dictionary.test.ts) וההתאמה עצמה — שימוש חוזר במנוע
 * ההתאמה הקיים, מקור אחר בלבד.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { buildSectionCache } from '../../src/engine/book-completion';
import {
  matchStaticCompletion,
  loadStaticSources,
  resetStaticSources,
} from '../../src/engine/static-completion';

const PACKED = {
  phrases: ['קם ליה בדרבה מיניה', 'אין אדם נתפס בשעת צערו'],
  authors: ["ר' אלחנן ווסרמן"],
};

describe('loadStaticSources', () => {
  beforeEach(() => resetStaticSources());

  it('נטען פעם אחת בלבד, גם בקריאות מקבילות', async () => {
    let calls = 0;
    const loader = async () => {
      calls += 1;
      return PACKED;
    };

    const [first, second] = await Promise.all([loadStaticSources(loader), loadStaticSources(loader)]);
    expect(calls).toBe(1);
    expect(first).toBe(second);
    expect(first?.map((s) => s.name)).toEqual(['talmudic-phrases', 'authors']);
  });

  it('כשל מחזיר null ואינו נזכר — הניסיון הבא טוען מחדש', async () => {
    let calls = 0;
    const failThenSucceed = async () => {
      calls += 1;
      return calls === 1 ? null : PACKED;
    };

    expect(await loadStaticSources(failThenSucceed)).toBeNull();
    expect(await loadStaticSources(failThenSucceed)).not.toBeNull();
    expect(calls).toBe(2);
  });
});

describe('matchStaticCompletion', () => {
  const sources = [
    { name: 'phrases', cache: buildSectionCache('קם ליה בדרבה מיניה\nאין אדם נתפס בשעת צערו') },
    { name: 'authors', cache: buildSectionCache('ר\' אלחנן ווסרמן') },
  ];

  it('משלימה מתוך המקור הראשון שיש בו התאמה', () => {
    const match = matchStaticCompletion({ precedingWords: ['קם', 'ליה'], partialWord: 'בדר' }, sources);
    expect(match?.source).toBe('phrases');
    expect(match?.text.startsWith('בדרבה')).toBe(true);
  });

  it('נופלת למקור השני כשאין התאמה בראשון', () => {
    const match = matchStaticCompletion({ precedingWords: [], partialWord: 'אלחנן' }, sources);
    expect(match?.source).toBe('authors');
  });

  it('null כשאין התאמה באף מקור', () => {
    expect(matchStaticCompletion({ precedingWords: [], partialWord: 'זזזזז' }, sources)).toBeNull();
  });
});
