/**
 * מה שהתוסף זוכר בין הפעלות.
 *
 * ה-`storage` של אוצריא הוא KV של JSON, ולכן נשמרים בו **רק** metadata
 * ו-tokens אטומים — לא bytes ולא Blob של מסמך. ה-token הוא מה ששורד: ה-URL
 * שאוצריא מחזירה תקף לריצה אחת בלבד, כי הפורט של שרת ה-loopback מתחלף.
 */
import { call, tryCall } from './otzaria-client';
import type { RecentDocument } from '../sessions/recent-documents';
import type { DiscardedDocument } from '../sessions/discard-backup';
import type { DocFolder, LibraryCache } from '../sessions/open-sources';

/**
 * המסמך שהיה פתוח לאחרונה.
 *
 * **מפתח מדור קודם.** מאז ש„חזרה בדיוק למה שהיה” כוללת גם את הסמן, את התצוגה
 * ואת מה שלא נשמר, זהות המסמך היא שדה אחד ברשומת ההפעלה (`SESSION_KEY`)
 * ואינה נכתבת לכאן יותר. הקריאה נשארה בשביל מסלול אחד בלבד: משתמש שמעדכן
 * מגרסה קודמת ויש לו רק את המפתח הזה. אחרי שהוא נקרא פעם אחת הוא נמחק, כדי
 * שלא יישאר מקור שני לאותה שאלה.
 */
export interface LastDocument {
  token: string;
  name: string;
  /** האם ה-token ניתן לכתיבה — כלומר „שמור” לא יפתח דיאלוג. */
  writable: boolean;
}

const LAST_DOCUMENT_KEY = 'last-document';

/**
 * קוראת את המסמך האחרון. כשל או ערך פגום מוחזרים כ-`null`: זיכרון של מסמך
 * אחרון אינו סיבה להיכשל בעלייה.
 */
export async function loadLastDocument(): Promise<LastDocument | null> {
  const raw = await tryCall<unknown>('storage.get', { key: LAST_DOCUMENT_KEY });
  if (!raw || typeof raw !== 'object') return null;

  const value = raw as Partial<LastDocument>;
  if (typeof value.token !== 'string' || value.token === '') return null;
  return {
    token: value.token,
    name: typeof value.name === 'string' && value.name !== '' ? value.name : 'מסמך',
    writable: value.writable === true,
  };
}

/** נמחק אחרי שהומר לרשומת ההפעלה — ראו ההערה על `LastDocument`. */
export async function forgetLastDocument(): Promise<void> {
  await tryCall('storage.remove', { key: LAST_DOCUMENT_KEY });
}

const SESSION_KEY = 'session';

/**
 * הרשומה הגולמית של מצב ההפעלה. הפירוש שלה — כולל אימות הגרסה והמסלול
 * ממשתמש שיש לו רק `last-document` מגרסה קודמת — יושב ב-sessions/session-state.ts,
 * כדי שההחלטות יהיו נבדקות בלי לזייף את הגשר.
 */
export async function loadSessionRecord(): Promise<unknown> {
  return tryCall<unknown>('storage.get', { key: SESSION_KEY });
}

export async function saveSessionRecord(value: unknown): Promise<void> {
  await tryCall('storage.set', { key: SESSION_KEY, value });
}

const AUTOSAVE_KEY = 'autosave-enabled';

/**
 * מתג „שמירה אוטומטית”. ברירת המחדל היא **דלוק**, וכל מה שאינו `false` מפורש
 * נקרא כדלוק: כשל קריאה או ערך פגום אינם סיבה להשאיר מסמך בלי שמירה
 * אוטומטית — הכיוון הבטוח כאן הוא לשמור יותר, לא פחות.
 */
export async function loadAutosaveEnabled(): Promise<boolean> {
  const raw = await tryCall<unknown>('storage.get', { key: AUTOSAVE_KEY });
  return raw !== false;
}

export async function saveAutosaveEnabled(enabled: boolean): Promise<void> {
  await tryCall('storage.set', { key: AUTOSAVE_KEY, value: enabled });
}

const RULER_KEY = 'ruler-visible';

/**
 * האם סרגל המידות מוצג. ברירת המחדל **כבויה**, כמו ב-Word מ-2013 ואילך.
 *
 * למה זה נשמר בכלל: מצב הסרגל יושב על מופע המנוע (`config.rulers`), ומופע חדש
 * נולד בכל פתיחת מסמך — כלומר בלי הזיכרון הזה הסרגל היה נכבה בכל פעם שמסמך
 * נפתח. ב-Word זו העדפה של התוכנה ולא תכונה של המסמך, וכך גם כאן.
 */
export async function loadRulerVisible(): Promise<boolean> {
  return (await tryCall<unknown>('storage.get', { key: RULER_KEY })) === true;
}

export async function saveRulerVisible(visible: boolean): Promise<void> {
  await tryCall('storage.set', { key: RULER_KEY, value: visible });
}

/** נשמר בנפרד מ-tryCall כדי שכשל בכתיבה לא ייעלם בשקט בקריאה מפורשת. */
export async function setSetting(key: string, value: unknown): Promise<void> {
  await call('storage.set', { key, value });
}

/**
 * הגדרה גנרית לפי מפתח — לזיכרון של דיאלוגים (composables/useRememberedOptions.ts).
 *
 * `fallback` חוזר גם כשאין ערך, גם בכשל קריאה וגם מחוץ לאוצריא: זיכרון של
 * דיאלוג אינו סיבה שהדיאלוג לא ייפתח. האימות של הצורה נעשה אצל הקורא —
 * כאן רק ההבחנה בין „יש משהו” ל„אין”.
 */
export async function loadSetting<T>(
  key: string,
  fallback: T,
  parse: (raw: unknown) => T | null,
): Promise<T> {
  const raw = await tryCall<unknown>('storage.get', { key });
  if (raw === null || raw === undefined) return fallback;
  return parse(raw) ?? fallback;
}

/** כתיבה שקטה — כשל בשמירת זיכרון של דיאלוג אינו כשל של הכלי. */
export async function saveSetting(key: string, value: unknown): Promise<void> {
  await tryCall('storage.set', { key, value });
}

const SPELLCHECK_KEY = 'spellcheck-enabled';

/**
 * האם בדיקת האיות התורנית דלוקה. ברירת המחדל **כבויה**, ולא במקרה: הדלקה
 * מושכת את המילון (1.3MB בנכס נפרד, engine/spellcheck-dictionary.ts), ומשתמש
 * שלא ביקש אותה לא אמור לשלם עליו.
 *
 * העדפה של התוכנה ולא תכונה של המסמך — כמו הסרגל שמעל, ומאותה סיבה: המצב
 * חי על שכבת התצוגה שלנו, ומסמך חדש היה מכבה אותה בכל פתיחה.
 */
export async function loadSpellcheckEnabled(): Promise<boolean> {
  return (await tryCall<unknown>('storage.get', { key: SPELLCHECK_KEY })) === true;
}

export async function saveSpellcheckEnabled(enabled: boolean): Promise<void> {
  await tryCall('storage.set', { key: SPELLCHECK_KEY, value: enabled });
}

const LIST_AUTOFORMAT_KEY = 'list-autoformat-enabled';

/**
 * זיהוי רשימות בהקלדה — „‏`א)` ורווח” הופך את הפסקה לרשימה. ברירת המחדל
 * **דלוקה**, וכל מה שאינו `false` מפורש נקרא כדלוק: זו התנהגות ש-Word נותן
 * מאז ומעולם, ומשתמש שלא הביע דעה מצפה לה.
 *
 * ובכל זאת יש מתג, בשונה מאזכור „@”: כאן המנגנון **משנה טקסט שהמשתמש
 * הקליד**, ומי שכותב „1. ” בכוונה כטקסט חייב דרך לכבות אותו.
 */
export async function loadListAutoformatEnabled(): Promise<boolean> {
  return (await tryCall<unknown>('storage.get', { key: LIST_AUTOFORMAT_KEY })) !== false;
}

export async function saveListAutoformatEnabled(enabled: boolean): Promise<void> {
  await tryCall('storage.set', { key: LIST_AUTOFORMAT_KEY, value: enabled });
}

const CANVAS_COLOR_KEY = 'canvas-color';

/**
 * צבע הבד — המשטח שסביב הדף. `null` = אין העדפה, והבד מקבל את ברירת
 * המחדל שב-composables/canvas-color.ts וב-styles/tokens.css.
 *
 * מוחזר גולמי ולא כצבע: מה שמגיע מ-`storage` הוא JSON שנכתב בהפעלה קודמת,
 * והאימות שלו יושב אצל הקורא (composables/canvas-color.ts) — שם גם מוגדר
 * מה נחשב צבע, ושם נכתב מה ש*נשמר*. שתי הגדרות לאותה שאלה הן בדיוק המקום
 * שבו ערך פגום עובר באחת מהן ונופל בשנייה.
 *
 * העדפה של התוכנה ולא תכונה של המסמך — כמו הסרגל ובדיקת האיות שמעל, ומאותה
 * סיבה: הבד הוא שכבת התצוגה שלנו, והוא נבנה מחדש בכל פתיחת מסמך.
 */
export async function loadCanvasColor(): Promise<unknown> {
  return tryCall<unknown>('storage.get', { key: CANVAS_COLOR_KEY });
}

export async function saveCanvasColor(color: string | null): Promise<void> {
  await tryCall('storage.set', { key: CANVAS_COLOR_KEY, value: color });
}

const SPELLCHECK_WORDS_KEY = 'spellcheck-user-words';

/**
 * מילון המשתמש — מה שנוסף דרך „הוסף למילון” בתפריט ההקשר.
 *
 * נשמר לבד, ולא כעותק של המילון הקבוע: 102,465 הערכים אינם עוברים ב-`storage`
 * לעולם. ערך פגום נקרא כרשימה ריקה — מילון משתמש שאבד אינו סיבה להיכשל
 * בעלייה, והמשתמש פשוט יוסיף שוב.
 */
export async function loadSpellcheckWords(): Promise<string[]> {
  const raw = await tryCall<unknown>('storage.get', { key: SPELLCHECK_WORDS_KEY });
  if (!Array.isArray(raw)) return [];
  return raw.filter((word): word is string => typeof word === 'string' && word.length > 0);
}

export async function saveSpellcheckWords(words: readonly string[]): Promise<void> {
  await tryCall('storage.set', { key: SPELLCHECK_WORDS_KEY, value: [...words] });
}

const CUSTOM_SHORTCUTS_KEY = 'custom-shortcuts';

/**
 * הקיצורים שהמשתמש הגדיר בעצמו, גולמיים.
 *
 * הפירוש — השמטת רשומות פגומות, השמטת צירוף שתפוס ברג'יסטרי והתקרה — יושב
 * ב-ui/shortcuts/custom-shortcuts.ts (`normalizeCustomShortcuts`), כמו ההפרדה
 * בין `loadSessionRecord` ל-`normalizeSession`: ההחלטות נבדקות בלי לזייף את
 * הגשר.
 *
 * ב-`storage` של אוצריא ולא ב-localStorage, בשונה ממערכת המאקרו: קיצור אישי
 * הוא העדפה של המשתמש כמו הסרגל ומילון האיות, והוא אמור לשרוד ניקוי של אחסון
 * הדפדפן שהתוסף יושב בו.
 */
export async function loadCustomShortcuts(): Promise<unknown> {
  return tryCall<unknown>('storage.get', { key: CUSTOM_SHORTCUTS_KEY });
}

/** כשל כתיבה נבלע כמו בכל שאר ההעדפות כאן — ומדווח למשתמש באתר הקריאה. */
export async function saveCustomShortcuts(list: readonly unknown[]): Promise<void> {
  await tryCall('storage.set', { key: CUSTOM_SHORTCUTS_KEY, value: [...list] });
}

const RECENT_DOCUMENTS_KEY = 'recent-documents';

/**
 * הרשימה הגולמית של „המסמכים האחרונים”. הפירוש — השמטת שורות פגומות, השמטת
 * token חוזר וברירות המחדל — יושב ב-sessions/recent-documents.ts
 * (`normalizeRecents`), בדיוק כמו ההפרדה בין `loadSessionRecord`
 * ל-`normalizeSession` שמעל, ומאותו טעם: ההחלטות נבדקות בלי לזייף את הגשר.
 */
export async function loadRecentDocuments(): Promise<unknown> {
  return tryCall<unknown>('storage.get', { key: RECENT_DOCUMENTS_KEY });
}

/**
 * כותבת את הרשימה. `[...list]` ולא `list` עצמה: מה שנמסר עשוי להיות מערך
 * לקריאה בלבד שנשמר גם ב-state של המעטפת, וה-JSON שיוצא לגשר לא אמור להיות
 * אותו אובייקט.
 *
 * כשל כתיבה נבלע (`tryCall`) כמו בכל שאר ההעדפות כאן: רשימת אחרונים שלא
 * נשמרה אינה סיבה להפיל פתיחת מסמך.
 */
export async function saveRecentDocuments(list: readonly RecentDocument[]): Promise<void> {
  await tryCall('storage.set', { key: RECENT_DOCUMENTS_KEY, value: [...list] });
}

const DISCARD_BACKUPS_KEY = 'discard-backups';

/**
 * רשומת „חמשת האחרונים שנסגרו בלי לשמור” — metadata בלבד: משבצת, שם, גודל,
 * זמן ו-token. הבייטים עצמם יושבים במרחב הפרטי (`host/workspace.ts`), בדיוק
 * כמו הטיוטה, ומאותה סיבה: ב-`storage` נשמר JSON, לא מסמך.
 *
 * הפירוש — השמטת שורות פגומות, משבצת כפולה, והסדר מהחדש לישן — יושב
 * ב-sessions/discard-backup.ts (`normalizeBackups`), כמו כל שאר ההפרדות כאן.
 */
export async function loadDiscardBackups(): Promise<unknown> {
  return tryCall<unknown>('storage.get', { key: DISCARD_BACKUPS_KEY });
}

/**
 * כותבת את הרשומה. כשל נבלע — גיבוי שלא נרשם אינו סיבה להפיל את הסגירה
 * שהמשתמש ביקש; מה שהוא כן עושה הוא להשאיר קובץ שהרשומה אינה מכירה, ואת זה
 * המשבצות כבר פותרות (ראו את ראש discard-backup.ts).
 */
export async function saveDiscardBackups(list: readonly DiscardedDocument[]): Promise<void> {
  await tryCall('storage.set', { key: DISCARD_BACKUPS_KEY, value: [...list] });
}

const DOC_FOLDERS_KEY = 'open-dialog-folders';

/**
 * התיקיות שהמשתמש הוסיף ל„פתח מסמך”. גולמי — הפירוש ב-sessions/open-sources.ts
 * (`normalizeFolders`), כמו כל שאר ההפרדות כאן. נשמרים token, שם ונתיב
 * לתצוגה בלבד: ה-token הוא ההרשאה, והנתיב אינו פותח דבר בלעדיו.
 */
export async function loadDocFolders(): Promise<unknown> {
  return tryCall<unknown>('storage.get', { key: DOC_FOLDERS_KEY });
}

export async function saveDocFolders(list: readonly DocFolder[]): Promise<void> {
  await tryCall('storage.set', { key: DOC_FOLDERS_KEY, value: [...list] });
}

const LIBRARY_CACHE_KEY = 'open-dialog-library';

/**
 * העץ הגזום של ספרי ה-Word מהפעם הקודמת — מה שמאפשר לדיאלוג להיפתח מלא
 * מיד, לפני שאוצריא ענתה. גולמי; הפענוח ב-`normalizeLibraryCache`.
 */
export async function loadLibraryCache(): Promise<unknown> {
  return tryCall<unknown>('storage.get', { key: LIBRARY_CACHE_KEY });
}

export async function saveLibraryCache(cache: LibraryCache): Promise<void> {
  await tryCall('storage.set', { key: LIBRARY_CACHE_KEY, value: cache });
}

const OPEN_PLACE_KEY = 'open-dialog-place';

/** הצומת שהיה נבחר בעץ בפעם הקודמת — מחרוזת אטומה של הדיאלוג. */
export async function loadOpenDialogPlace(): Promise<unknown> {
  return tryCall<unknown>('storage.get', { key: OPEN_PLACE_KEY });
}

export async function saveOpenDialogPlace(place: string): Promise<void> {
  await tryCall('storage.set', { key: OPEN_PLACE_KEY, value: place });
}
