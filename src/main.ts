/**
 * נקודת הכניסה הראשית — עורך Word לאוצריא (SuperDoc v2 + Vue 3).
 */
import { createApp } from 'vue';
import './styles/tokens.css';
import './styles/shell.css';
import './styles/ribbon.css';
import './styles/print.css';
import './styles/crop-marks.css';
import App from './App.vue';
import { installBundledFonts } from './styles/fonts';
import { onThemeChanged, resolveBoot } from './host/otzaria-client';
import { setHostAppVersion } from './host/host-capabilities';
import { splashFail, splashStage, SPLASH_STAGES } from './host/splash';
import { applyTheme } from './host/theme';
import { setMenuLocale } from './ui/ribbon/i18n';
import { setCanvasDefaultColor } from './composables/canvas-color';
import type { ThemePayload } from './types/otzaria_plugin';

/** ב-build הסקריפט קלאסי, כלומר הוא עשוי לרוץ לפני שה-body נפרס. */
function domReady(): Promise<void> {
  if (document.readyState !== 'loading') return Promise.resolve();
  return new Promise((resolve) => {
    document.addEventListener('DOMContentLoaded', () => resolve(), { once: true });
  });
}

function applyOtzariaTheme(theme: ThemePayload): void {
  applyTheme(theme);
  setCanvasDefaultColor(theme.colorScheme.surface);
}

async function main(): Promise<void> {
  // התקנת גופנים ארוזים ראשונה
  installBundledFonts();

  if (import.meta.env.DEV) {
    const { installDevStub } = await import('./host/dev-stub');
    installDevStub();
  }

  // שחזור או קבלת אירוע boot
  const bootPromise = resolveBoot();

  await domReady();

  // הרכבת אפליקציית Vue
  const app = createApp(App);
  app.mount('#app');
  splashStage(SPLASH_STAGES.shellMounted, 'מכין את סביבת העריכה…');

  // תוצאת האתחול, על שורש ה-HTML: 'event' — האירוע נתפס ב-latch; 'recovered' —
  // האירוע אבד והמצב שוחזר ב-RPC; 'failed' — שניהם כשלו, והממשק עולה עם ערכת
  // הנושא של ברירת המחדל.
  //
  // זה לא קוסמטי ולא רק אבחון: זהו הסימן היחיד מבחוץ למה שקרה באתחול. הממשק
  // עולה בכל שלושת המצבים (כשל אתחול אינו מקפיא אותו יותר), ולכן „הכפתור
  // נפתח” הפסיק להיות עדות — ושער `check:boot` נשען מעכשיו על התכונה הזאת.
  const root = document.documentElement;
  try {
    const info = await bootPromise;
    applyOtzariaTheme(info.theme);
    // שפת התפריטים לפי שפת המשתמש (`app.language` — 'he' / 'en'; ראו
    // docs/plugin-sdk). נקבעת גם במסלול „recovered", כי `app.getInfo`
    // מחזיר את אותו מידע. כשל אתחול משאיר עברית — שפת ברירת המחדל.
    setMenuLocale(info.app.language);
    // גרסת אוצריא קובעת אילו קריאות Host קיימות — ר' host/host-capabilities.ts.
    // נקבעת גם במסלול „recovered", מאותה סיבה ומאותו מקור כמו השפה.
    setHostAppVersion(info.app.version);
    onThemeChanged(applyOtzariaTheme);
    root.dataset.boot = info.source === 'recovered' ? 'recovered' : 'event';

    if (info.source === 'recovered') {
      console.warn('[otzaria-word] plugin.boot אבד; מצב האתחול שוחזר ב-RPC');
    }
  } catch (error) {
    root.dataset.boot = 'failed';
    console.error('[otzaria-word] כשל באתחול ערכת הנושא של אוצריא:', error);
  }
}

// כשל לפני שהממשק הורכב אינו מגיע לשום מקום שאפשר לראות אותו בו — הממשק
// עצמו הוא מה שלא עלה. מסך הטעינה הוא המשטח היחיד שנשאר, ולכן הוא מדווח.
void main().catch((error: unknown) => {
  console.error('[otzaria-word] כשל בעליית התוסף:', error);
  splashFail('התוסף לא הצליח לעלות', error instanceof Error ? error.message : String(error));
});
