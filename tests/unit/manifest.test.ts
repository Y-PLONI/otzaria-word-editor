/**
 * המניפסט הוא ההצהרה היחידה שאוצריא קוראת לפני שהיא מעניקה הרשאה. קריאת RPC
 * להרשאה שלא הוצהרה נכשלת ב-`error.permission_denied` — כלומר כפתור שנראה
 * עובד ואינו עובד, והכשל מתגלה אצל המשתמש ולא כאן.
 *
 * הבדיקה נשענת על `READER_PERMISSIONS` (host/otzaria-reader.ts) ולא על רשימה
 * שנכתבה כאן: רשימה שנייה הייתה מקור אמת שני, ומתודה שתתווסף בעתיד לא הייתה
 * מגיעה אליה.
 *
 * `public/manifest.json` הוא המקור. `dist/manifest.json` נוצר ממנו ב-build
 * (תיקיית `public` של Vite מועתקת כמו שהיא, והיא אינה במעקב git) ולכן אין
 * לערוך אותו ואין לבדוק אותו.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { READER_PERMISSIONS, SEND_TO_DOCUMENT_ITEM } from '../../src/host/otzaria-reader';
import { LIBRARY_PERMISSIONS } from '../../src/host/otzaria-library';

/** vitest רץ משורש המאגר, ולכן public/ נמצא ביחס ל-cwd. */
const manifest = JSON.parse(
  readFileSync(join(process.cwd(), 'public/manifest.json'), 'utf8'),
) as {
  permissions?: string[];
  minAppVersion?: string;
  contributes?: { startup?: { contextMenuItems?: unknown[] } };
};

describe('public/manifest.json', () => {
  it('מצהיר על כל ההרשאות שהקוד צורך מהקורא ומהניווט', () => {
    const declared = new Set(manifest.permissions ?? []);
    const missing = [...new Set(Object.values(READER_PERMISSIONS))].filter(
      (permission) => !declared.has(permission),
    );

    expect(missing).toEqual([]);
  });

  it('מצהיר על כל ההרשאות שתיוג המקורות (source-tagging) צורך מהספרייה', () => {
    const declared = new Set(manifest.permissions ?? []);
    const missing = [...new Set(Object.values(LIBRARY_PERMISSIONS))].filter(
      (permission) => !declared.has(permission),
    );

    expect(missing).toEqual([]);
  });

  it('מצהיר על הרשאות קבצי המשתמש שמסלול הפתיחה והשמירה צורך', () => {
    // host/files.ts קורא ל-fs.pickUserFile ול-fs.commitUserFileWrite, ושתי
    // ההרשאות האלה היו היחידות במניפסט לפני שהתווספו הרשאות הקורא.
    expect(manifest.permissions).toContain('fs.user_files.read');
    expect(manifest.permissions).toContain('fs.user_files.write');
  });

  it('אין הרשאה מוצהרת פעמיים', () => {
    const permissions = manifest.permissions ?? [];
    expect(permissions.length).toBe(new Set(permissions).size);
  });

  it('minAppVersion מכסה את ה-API החדש ביותר שהתוסף קורא לו', () => {
    // `reader.getSelection`, `reader.openSearchTab` ו-`navigation.goTo` קיימים
    // מ-0.9.89; מסלול הכתיבה (`fs.commitUserFileWrite`, `access: 'readwrite'`)
    // דורש 0.9.97 והוא הגבוה מביניהם. סקריפט האריזה של אוצריא חוסם אריזה
    // כשההצהרה נמוכה ממה שהקוד קורא לו, ולכן זו אינה בדיקה תיאורטית.
    expect(manifest.minAppVersion).toBe('0.9.97');
  });

  /**
   * הפריט „שלח למסמך” מוצהר במניפסט ולא נרשם מ-JS בלבד, ובכוונה: רישום מ-JS
   * חי רק כל עוד מופע התוסף חי, ולכן הפריט לא היה קיים בדיוק בתרחיש שלו —
   * משתמש שקורא בספרייה ועוד לא פתח את התוסף. אוצריא מיישמת את ההצהרה
   * ב-Dart בעליית האפליקציה, והיא שורדת סגירה של לשונית התוסף.
   */
  it('מצהיר את „שלח למסמך” כתרומת עלייה, עם ההרשאה שההצהרה דורשת', () => {
    expect(manifest.permissions).toContain('app.startup_contributions');
    expect(manifest.contributes?.startup?.contextMenuItems).toEqual([
      { ...SEND_TO_DOCUMENT_ITEM, contexts: [...SEND_TO_DOCUMENT_ITEM.contexts] },
    ]);
  });
});
