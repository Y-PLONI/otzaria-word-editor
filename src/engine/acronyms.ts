/**
 * ראשי תיבות: מיפוי `"א"א" → ["אי אפשר", "אמר אברהם", ...]`. המקור —
 * `Acronyms.json` הרשמי של אוצריא (13,105 ערכים) — ר' docs/smart-source-completion-plan.md.
 *
 * בניגוד ל-book-completion/static-completion, כאן אין "השלמת המשך" — הראשי
 * תיבות כבר הוקלד **במלואו** (עם הגרשיים), וההשלמה היא הפירוש שלו. לכן
 * ההתאמה היא שווה-ערך מדויק על המילה האחרונה שהוקלדה, לא prefix.
 *
 * פירוש יחיד לכל ר"ת בשלב זה: הראשון ברשימה (סדר המקור, לא מנוין מחדש).
 * מחזור בין פירושים נוספים — לא בסקופ הנוכחי (ר' המסמך).
 */
export interface AcronymDictionary {
  /** הפירוש הראשון של `acronym`, או `null` אם אינו ר"ת מוכר. */
  lookup(acronym: string): string | null;
}

export function createAcronymDictionary(packedJson: string): AcronymDictionary {
  const parsed = JSON.parse(packedJson) as Record<string, string[]>;

  return {
    lookup(acronym: string): string | null {
      const expansions = parsed[acronym];
      return Array.isArray(expansions) && expansions.length > 0 ? (expansions[0] ?? null) : null;
    },
  };
}
