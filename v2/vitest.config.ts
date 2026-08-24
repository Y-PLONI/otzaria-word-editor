import { defineConfig } from 'vitest/config';
import vue from '@vitejs/plugin-vue';

/**
 * הבדיקות אינן מרימות את מנוע ה-DOCX: הוא דורש workers ו-canvas אמיתיים,
 * ולכן ריצה חיה נבדקת בשערי Windows (docs/spike-windows.md) ולא ב-jsdom.
 * מה שכן נבדק כאן: חוזה ה-API של superdoc/ui, ה-registry של הפקודות,
 * האדפטרים שלנו מול כפילים — ומאז שיש @vue/test-utils גם הקומפוננטות עצמן
 * (tests/component), שמורכבות ב-jsdom ונלחצות באמת.
 *
 * ה-plugin של Vue נדרש בשביל אותן בדיקות: בלעדיו `import` של קובץ `.vue`
 * מגיע ל-esbuild כ-JavaScript ונופל על התבנית.
 */
export default defineConfig({
  plugins: [vue()],
  test: {
    environment: 'jsdom',
    include: ['tests/**/*.test.ts'],
    /**
     * ברירת המחדל של vitest היא 5 שניות, וזה מעט מדי לסריקת הכפתורים
     * ב-tests/component/ribbon-tabs.test.ts: היא מרכיבה את הלשונית **מחדש
     * לכל כפתור** כדי שכל לחיצה תימדד על מצב נקי, ולשונית „בית” היא כשלושים
     * הרכבות של קומפוננטה בת 24KB ב-jsdom.
     *
     * נמדד: 3.6 שניות כשהקובץ רץ לבד, ומעל 5 כשהוא רץ במקביל לשאר החבילה.
     * כלומר התקרה הקודמת הייתה תלוית-עומס — הבדיקה נפלה על המכונה ולא על
     * הקוד, וזה שער מתעתע ולא שער. 20 שניות נותנות מרווח פי חמישה מהמדידה
     * הבודדת, ובלי לבטל את ההגנה מפני בדיקה שנתקעת באמת.
     */
    testTimeout: 20_000,
  },
});
