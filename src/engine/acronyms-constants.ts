/**
 * שם הנכס והמפתח הגלובלי של מילון ראשי-התיבות, בדיוק כמו
 * `spellcheck.ts` (`TORAH_DICTIONARY_FILE`/`_GLOBAL`) — מודול בלי import כדי
 * ש-`vite.config.ts` (שאסור לו לתלות בשרשרת המנוע) יוכל לייבא אותו.
 */
export const ACRONYMS_FILE = 'assets/acronyms.js';
export const ACRONYMS_GLOBAL = '__OTZARIA_ACRONYMS__';
