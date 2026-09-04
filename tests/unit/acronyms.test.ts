/**
 * מילון ראשי-התיבות: פענוח (engine/acronyms.ts) וטעינה עצלה
 * (engine/acronym-dictionary.ts) — אותה תבנית כמו spellcheck-dictionary.test.ts.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { createAcronymDictionary } from '../../src/engine/acronyms';
import { loadAcronymDictionary, resetAcronymDictionary } from '../../src/engine/acronym-dictionary';

const PACKED = JSON.stringify({ 'א"א': ['אי אפשר', 'אמר אברהם'], 'רמב"ם': ['רבי משה בן מימון'] });

describe('createAcronymDictionary', () => {
  const dictionary = createAcronymDictionary(PACKED);

  it('מחזירה את הפירוש הראשון לר"ת מוכר', () => {
    expect(dictionary.lookup('א"א')).toBe('אי אפשר');
    expect(dictionary.lookup('רמב"ם')).toBe('רבי משה בן מימון');
  });

  it('null לר"ת שאינו ברשימה', () => {
    expect(dictionary.lookup('זזזזז')).toBeNull();
  });
});

describe('loadAcronymDictionary', () => {
  beforeEach(() => resetAcronymDictionary());

  it('נטען פעם אחת בלבד, גם בקריאות מקבילות', async () => {
    let calls = 0;
    const loader = async () => {
      calls += 1;
      return PACKED;
    };

    const [first, second] = await Promise.all([loadAcronymDictionary(loader), loadAcronymDictionary(loader)]);
    expect(calls).toBe(1);
    expect(first).toBe(second);
  });

  it('כשל מחזיר null ואינו נזכר — הניסיון הבא טוען מחדש', async () => {
    let calls = 0;
    const failThenSucceed = async () => {
      calls += 1;
      return calls === 1 ? null : PACKED;
    };

    const failed = await loadAcronymDictionary(failThenSucceed);
    expect(failed).toBeNull();

    const succeeded = await loadAcronymDictionary(failThenSucceed);
    expect(succeeded?.lookup('א"א')).toBe('אי אפשר');
    expect(calls).toBe(2);
  });
});
