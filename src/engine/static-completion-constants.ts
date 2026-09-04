/**
 * שם הנכס והמפתח הגלובלי של מילוני ההשלמה הסטטיים (ביטויים תלמודיים +
 * מחברים) — מודול בלי import, כמו `spellcheck.ts`/`acronyms-constants.ts,
 * כדי ש-`vite.config.ts` יוכל לייבא אותו.
 *
 * שני המקורות באסט אחד ולא שניים: הם תמיד נטענים יחד (fallback אחד אחרי
 * השני לאותה הקשה), ושניהם קטנים בהרבה מהמילון התורני — אין תועלת בפיצול.
 */
export const STATIC_COMPLETION_FILE = 'assets/static-completion.js';
export const STATIC_COMPLETION_GLOBAL = '__OTZARIA_STATIC_COMPLETION__';
