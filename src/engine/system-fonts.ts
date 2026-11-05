/**
 * הגופנים שמותקנים במכונה — ומה עושים כשאי אפשר לשאול.
 *
 * ## למה זה לא טריוויאלי
 *
 * הבורר יכול להציג רק שמות שהוא יודע עליהם מראש, ולדף אין דרך פשוטה למנות
 * גופני מערכת. שתי הדרכים המתבקשות, ומה שידוע על כל אחת:
 *
 * - **`fs.listDir`** — סגור. לא הנחה: ה-d.ts מצהיר במפורש שכל נתיב יחסי לשורש
 *   התוסף ושנתיב מוחלט נדחה ב-`error.forbidden`
 *   (types/otzaria_plugin.d.ts, „The plugin's private file workspace”).
 *
 * - **`queryLocalFonts()`** — **לא נמדד באוצריא, וזו ההנחה שהכריעה כאן.**
 *   מה שכן נמדד, ב-Chrome headless על ה-dist מ-`file://`
 *   (scripts/qa/installed-fonts-qa.mjs): הפונקציה **קיימת**
 *   ו-`isSecureContext === true` — כלומר שער ה-origin, שהיה החשד הראשון, אינו
 *   חוסם. מה שנשאר פתוח הוא ההרשאה עצמה: `local-fonts` מגיעה ב-WebView2 דרך
 *   `PermissionRequested`, ואוצריא ב-Windows רצה ב-visual hosting
 *   (docs/engine-gaps.md) — שם אין חלון להציג בו prompt. Chrome headless אינו
 *   WebView2, ולכן המדידה שלמעלה **אינה** עונה על השאלה הזאת.
 *
 * המסקנה המעשית אינה תלויה בהכרעה: גם אילו `queryLocalFonts` היה עובד, הוא
 * אינו מחזיר את מה שהמארח כן מחזיר — **אילו ערכות תווים כל גופן מכסה**, ובעורך
 * עברי זה מה שמפריד בין רשימה שטוחה של מאות שמות לבין בורר עם קבוצת „עברית”
 * בראשו. לכן `fonts.listInstalled` (docs/otzaria-fonts-list-request.md) היא
 * הבקשה שהוגשה, ו-`queryLocalFonts` נשאר שכבה רביעית אפשרית אם יימדד שהוא עובד.
 *
 * ## שלוש שכבות, בסדר הזה
 *
 * 1. **`fonts.listInstalled`** — הרשימה האמיתית. `tryCall` מחזיר `null` למארח
 *    שאינו מכיר את המתודה, ולכן אין כאן שום בדיקת גרסה: אוצריא ישנה פשוט
 *    נופלת לשכבה הבאה.
 *
 *    **קיימת מ-0.9.97.** מומשה ב-`installed_fonts.dart` (GDI,
 *    `EnumFontFamiliesExW`) ומוזגה ל-`upstream/dev` ול-`upstream/pre-0.9.97`;
 *    החוזה ב-docs/otzaria-fonts-list-request.md. נמדד על Windows דרך הקוד של
 *    אוצריא עצמו: 287 משפחות · 57 עם `hebrew` · 22 `monospace`.
 *
 *    **בנייה מלפני כן נופלת לשכבה 2**, ואין כאן שום בדיקת גרסה שתאמר זאת:
 *    `tryCall` מחזיר `null` למארח שאינו מכיר את המתודה, וזה כל מה שנדרש.
 *    נמדד על הפורטבל 0.9.96 — ה-snapshot שלו אינו מכיל את שם המתודה כלל.
 *
 *    אין כאן הצהרת הרשאה: `app.info.read`, שהמתודה יושבת תחתיה, היא **הרשאת
 *    בסיס** שאוצריא מעניקה אוטומטית, והוולידטור שלה ממליץ להסיר הצהרות
 *    מיותרות ממניפסטים (ראו host/otzaria-client.ts).
 * 2. **מדידה של רשימת מועמדים** — קירוב. מכסה כל מכונת Windows סבירה, ואינו
 *    דורש דבר מהמארח.
 * 3. **כלום** — כשאין canvas אין לנו מה לומר, והבורר נשאר על הרשימה הקבועה.
 *
 * ## גם מה שהמארח אומר נמדד — אבל אינו נמחק
 *
 * המארח מונה דרך GDI, והדפדפן מרנדר דרך DirectWrite. השניים כמעט זהים אבל לא
 * לגמרי, ושם שהמארח מכיר והדפדפן אינו פותר הוא בדיוק התקלה שהפרויקט הזה כבר
 * מכיר: `lineRule="auto"` גוזר את גובה השורה מהגופן שנבחר **בפועל**, וגופן
 * שנפל ל-fallback מותח כל שורה במסמך (ראו docx-fonts.ts). לכן כל שם שחוזר
 * מהמארח עובר את אותה מדידה.
 *
 * מה שהמדידה **עושה** בתוצאה שלה שונה בין שני המקורות, וזה תוקן אחרי שנמדד:
 * ברשימת המארח היא מסמנת (`unresolved`), וברשימת המועמדים היא מוחקת. שם שהמארח
 * דיווח עליו מותקן — אי-פתירה שלו היא פער GDI/DirectWrite ולא היעדר — ומחיקתו
 * הסתירה 43 מ-287 גופנים מותקנים במכונה שנמדדה. ראו `measureUnresolved`.
 *
 * אם אין במה למדוד — סומכים על המארח, כי ידיעה חלקית עדיפה על היעדר ידיעה;
 * אבל **ניחוש** בלי מדידה אינו שווה כלום, ולכן שכבה 2 אינה רצה בלי canvas.
 */
import type { FontFamilyOption } from 'superdoc/ui';
import { tryCall } from '../host/otzaria-client';
import { cachedFamilyAvailability, cachedHebrewCoverage, canMeasureFonts, isFamilyAvailable } from './docx-fonts';

/**
 * משפחה אחת, בצורה שהיא גם מה שהמארח מחזיר וגם מה שרשימת המועמדים כתובה בה.
 *
 * מוגדרת כאן ולא ב-`types/otzaria_plugin.d.ts` אף שזה חוזה של המארח: הקובץ ההוא
 * הוא **העתק verbatim** של ה-d.ts הרשמי (ראו host/otzaria-client.ts), ותוספת
 * ידנית בו נמחקת בהעתקה הבאה — ואיתה גם הייבוא מכאן. החוזה עצמו חי ב-
 * docs/otzaria-fonts-list-request.md.
 */
export interface InstalledFont {
  name: string;
  /** `hebrew` הוא היחיד שנצרך כאן; השאר נשמרים כפי שהם. */
  scripts?: readonly string[];
  monospace?: boolean;
}

/** מה ש-`fonts.listInstalled` מחזירה. הכול אופציונלי — מגיע מחוץ לתוסף. */
export interface ListInstalledFontsResult {
  families?: readonly InstalledFont[];
  /** `windows` | `android` | `linux` | `macos`. נקרא לדיווח בלבד. */
  platform?: string;
}

/** מה שהבורר מקבל בסוף. */
export interface InstalledFontsSnapshot {
  /** ממוין, בלי כפילות. ריק כשאין ידיעה. */
  families: readonly FontFamilyOption[];
  /** המפתחות (lowercase) של המשפחות שמכסות עברית. */
  hebrew: ReadonlySet<string>;
  /**
   * המפתחות (lowercase) של המשפחות שהמארח דיווח עליהן ו**הדפדפן אינו פותר**.
   *
   * הן נשארות ברשימה — ראו „מה שהמנייה מדווחת אינו מסונן” ב-`loadInstalledFonts`
   * — והקבוצה הזאת היא מה שמאפשר למיזוג לסמן אותן בלי למדוד מחדש את כולן.
   *
   * ריקה במסלול המדידה: שם שלא נפתר שם כלל אינו נכנס לרשימה.
   */
  unresolved: ReadonlySet<string>;
  /** מאיפה הגיע. לדיווח ולבדיקה — הבורר עצמו אינו מבחין. */
  source: 'host' | 'measured' | 'none';
}

/** אין ידיעה. מה שמסופק לפני שהמנייה נחתה, וגם כשאין במה למדוד. */
export function emptyInstalledFonts(): InstalledFontsSnapshot {
  return {
    families: [],
    hebrew: new Set<string>(),
    unresolved: new Set<string>(),
    source: 'none',
  };
}

/* ------------------------------------------------------------------ */
/* רשימת המועמדים — שכבה 2 בלבד                                        */
/* ------------------------------------------------------------------ */

/**
 * גופנים עבריים שסביר שמותקנים: מה ש-Windows מביא, מה ש-Office מוסיף, ומה
 * שאוצריא מזריקה. גם `Arial` ו-`Segoe UI` כאן — הם באמת מכסים עברית, וזה
 * בדיוק מה שהמארח היה מדווח עליהם.
 *
 * **הסיווג הזה אינו נמדד — זו רשימה מתוחזקת ביד.** המדידה בשכבה 2 עונה על
 * שאלה אחת בלבד, „האם הדפדפן פותר את השם”, ואינה יודעת דבר על כיסוי תווים.
 * לכן שם שנוסף כאן בטעות ייכנס לקבוצת „עברית” גם אם אין בו אות עברית אחת.
 * במסלול המארח הסיווג אמיתי — GDI מדווח `lfCharSet` לכל משפחה.
 */
const HEBREW_CANDIDATES: readonly string[] = [
  'Aharoni', 'Alef', 'Arial', 'Assistant', 'Courier New', 'David', 'David CLM',
  'Drugulin CLM', 'Ellinia CLM', 'FrankRuehl', 'FrankRuhlCLM', 'Gisha',
  'Hadasim CLM', 'Heebo', 'Keter YG', 'Levenim MT', 'Miriam', 'Miriam CLM',
  'Miriam Fixed', 'Narkisim', 'NotoRashiHebrew', 'Rod', 'Rubik', 'Segoe UI',
  'Shofar', 'Taamey Frank CLM', 'TaameyDavidCLM', 'Tahoma', 'Times New Roman',
];

/** גופני הלטינית של Windows ושל Office. */
const LATIN_CANDIDATES: readonly string[] = [
  'Aptos', 'Aptos Display', 'Aptos Mono', 'Aptos Narrow', 'Arial Black',
  'Arial Narrow', 'Bahnschrift', 'Book Antiqua', 'Bookman Old Style',
  'Calibri', 'Calibri Light', 'Cambria', 'Candara', 'Cascadia Code',
  'Cascadia Mono', 'Century Gothic', 'Comic Sans MS', 'Consolas',
  'Constantia', 'Corbel', 'Ebrima', 'Franklin Gothic Book',
  'Franklin Gothic Medium', 'Gabriola', 'Gadugi', 'Garamond', 'Georgia',
  'Impact', 'Ink Free', 'Lucida Console', 'Lucida Sans Unicode',
  'Microsoft Sans Serif', 'MS Gothic', 'MV Boli', 'Nirmala UI',
  'Palatino Linotype', 'Perpetua', 'Rockwell', 'Segoe Print', 'Segoe Script',
  'Segoe UI Emoji', 'Segoe UI Symbol', 'Sitka Text', 'Sylfaen', 'Symbol',
  'Trebuchet MS', 'Verdana', 'Webdings', 'Wingdings', 'Yu Gothic',
];

const MONOSPACE_CANDIDATES: ReadonlySet<string> = new Set([
  'Aptos Mono', 'Cascadia Code', 'Cascadia Mono', 'Consolas', 'Courier New',
  'Lucida Console', 'Miriam Fixed', 'MS Gothic',
]);

function candidate(name: string, scripts: readonly string[]): InstalledFont {
  return { name, scripts, monospace: MONOSPACE_CANDIDATES.has(name) };
}

/** רשימת המועמדים שנמדדת כשאין מארח שיודע לענות. */
export const MEASURED_CANDIDATES: readonly InstalledFont[] = [
  ...HEBREW_CANDIDATES.map((name) => candidate(name, ['hebrew', 'latin'])),
  ...LATIN_CANDIDATES.map((name) => candidate(name, ['latin'])),
];

/** תוצאת הביניים של סריקת המועמדים; אין כאן מדידה סינכרונית. */
export function cachedInstalledFonts(): InstalledFontsSnapshot {
  return toSnapshot(MEASURED_CANDIDATES
    .filter((font) => cachedFamilyAvailability(font.name) === true)
    .map((font) => ({ ...font, scripts: cachedHebrewCoverage(font.name) === true ? ['hebrew'] : [] })),
  'measured');
}

/* ------------------------------------------------------------------ */
/* המנייה                                                              */
/* ------------------------------------------------------------------ */

/** מה שאפשר להחליף בבדיקה. ברירות המחדל הן המציאות. */
export interface InstalledFontsDeps {
  call?: typeof tryCall;
  available?: (name: string) => boolean;
  canMeasure?: () => boolean;
  /** ההמתנה לגופנים המוזרקים. ראו `waitForWebFonts`. */
  fontsReady?: () => Promise<unknown>;
}

/**
 * ההמתנה שבלעדיה המדידה מודדת את הרגע הלא נכון.
 *
 * `isFamilyAvailable` משווה רוחב מול גופן בסיס, וגופן `@font-face` שטרם נטען
 * מודד **בדיוק כמו הבסיס** — כלומר „אינו קיים”. זה אינו תרחיש תיאורטי: הגופן
 * הארוז מוצהר `font-display: swap` (styles/fonts.ts), וגופני העברית של אוצריא
 * מוזרקים אחרי עליית הדף (ראו font-options.ts). מדידה מוקדמת מדי הייתה מוחקת
 * בדיוק את הגופנים העבריים שהמנייה קיימת בשבילם — Alef, Heebo, ו-CLM למיניהם.
 *
 * נכשלת בשקט: דפדפן בלי `document.fonts` ימדוד מוקדם, וזה עדיין טוב מלא למדוד.
 */
function waitForWebFonts(): Promise<unknown> {
  try {
    return document.fonts?.ready ?? Promise.resolve();
  } catch {
    return Promise.resolve();
  }
}

/**
 * מה שהמארח אומר, ובלי למדוד דבר — **שלב א' מתוך שניים.**
 *
 * ## למה זה נפרד
 *
 * המדידה יקרה, ונמדדה: 6.46 שניות ל-287 משפחות, במקטעים שהארוך בהם 68ms גם
 * אחרי הוויתור על החוט (ראו `YIELD_BUDGET_MS`). היא אינה חוסמת את פתיחת
 * המסמך — הקוראת מפעילה בלי `await` — אבל היא כן מתחרה עם ההקלדה על ה-main
 * thread בדיוק בשניות הראשונות שאחרי הפתיחה, וזו התלונה שהובילה לכאן.
 *
 * שלב א' הוא קריאת RPC אחת ומיון. הבורר מתמלא ממנו מיד, וכל מה שחסר בו הוא
 * הסימון של השמות שהדפדפן אינו פותר (`unresolved`). עד לבדיקה הזמינות
 * אינה ידועה ואין דגימה. App.vue בודקת בתור idle הניתן להשהיה;
 * `verifyInstalledFonts` נשארת כניסה נפרדת למי שממתין לתוצאה מלאה.
 *
 * מארח שאינו מכיר את המתודה אינו מקבל כאן כלום — שכבה 2 היא **כולה** מדידה,
 * ולכן היא שייכת לשלב ב'. ראו `loadInstalledFonts` להסבר על השכבות.
 *
 * אינה זורקת לעולם — כל כשל מחזיר פחות ידיעה, לא חריגה.
 */
export async function loadReportedFonts(
  deps: InstalledFontsDeps = {},
): Promise<InstalledFontsSnapshot> {
  const { call = tryCall } = deps;

  const hosted = await call<ListInstalledFontsResult>('fonts.listInstalled').catch(() => null);
  const reported: readonly InstalledFont[] =
    hosted && Array.isArray(hosted.families) ? hosted.families : [];

  return reported.length > 0 ? toSnapshot(reported, 'host') : emptyInstalledFonts();
}

/**
 * המדידה — **שלב ב'.** מקבלת את מה ש-`loadReportedFonts` החזירה ומחזירה את
 * אותה ידיעה, מאומתת.
 *
 * שני מסלולים, לפי מה שהשלב הראשון מצא:
 *
 * - **`source: 'host'`** — הרשימה נשארת שלמה, והמדידה רק **מסמנת** את מי
 *   שהדפדפן אינו פותר. ראו `measureUnresolved` לכל ההנמקה.
 * - **`source: 'none'`** — המארח לא ענה, וכאן נופלים לרשימת המועמדים. זו
 *   שכבה 2, והיא כולה מדידה: בלי canvas אין לה מה לומר.
 *
 * אינה זורקת לעולם, ומחזירה את מה שנכנס כשאין מה לאמת.
 */
export async function verifyInstalledFonts(
  snapshot: InstalledFontsSnapshot,
  deps: InstalledFontsDeps = {},
): Promise<InstalledFontsSnapshot> {
  const {
    call = tryCall,
    available = isFamilyAvailable,
    canMeasure = canMeasureFonts,
    fontsReady = waitForWebFonts,
  } = deps;

  if (snapshot.source === 'host') {
    // בלי canvas סומכים על המארח כמות שהוא: הוא ספר את המכונה, ואנחנו לא.
    if (!canMeasure()) return snapshot;
    await fontsReady().catch(() => {});
    const reported = snapshot.families.map((family) => ({ name: family.value }));
    return { ...snapshot, unresolved: await measureUnresolved(reported, available) };
  }

  // הנפילה לשכבה 2 שקטה לחלוטין מבחינת המשתמש — הרשימה פשוט מתכווצת. שורה
  // אחת בקונסולה היא ההבדל בין חמש דקות לשעה של חיפוש.
  void explainHostGap(call);

  // ניחוש בלי מדידה אינו שווה כלום — עשרות שמות שאיש אינו יודע אם קיימים.
  if (!canMeasure()) return snapshot;
  await fontsReady().catch(() => {});
  return toSnapshot(await keepAvailable(MEASURED_CANDIDATES, available), 'measured');
}

/**
 * שני השלבים יחד. אינה זורקת לעולם — כל כשל מחזיר פחות ידיעה, לא חריגה.
 *
 * מי שרוצה את הרשימה מיד ואת המדידה בזמן שנוח לו קורא לשני השלבים בנפרד —
 * App.vue משתמשת בתור המדידות הניתן להשהיה שב-font-probes.ts. כאן נשמרת
 * נקודת הכניסה החד־פעמית, למי שצריך להמתין לתשובה המלאה.
 */
export async function loadInstalledFonts(
  deps: InstalledFontsDeps = {},
): Promise<InstalledFontsSnapshot> {
  return verifyInstalledFonts(await loadReportedFonts(deps), deps);
}

/**
 * ההרשאה ש-`fonts.listInstalled` תשב תחתיה. **הרשאת בסיס** — אוצריא מעניקה
 * אותה לכל תוסף בלי הצהרה במניפסט, ולכן אין כאן מה להצהיר.
 */
export const FONTS_PERMISSION = 'app.info.read';

/**
 * למה המארח לא ענה — לקונסולה, פעם אחת.
 *
 * הסיבה הצפויה היא אחת: אין באוצריא מתודה בשם הזה עדיין. הבדיקה בכל זאת שואלת
 * מה אושר, מפני ש-`tryCall` מחזיר `null` גם על סירוב הרשאה — ואם המתודה תיכנס
 * יום אחד תחת הרשאה שאינה בסיס, זה ההבדל בין „לשדרג את אוצריא” לבין „לאשר”.
 *
 * `info` ולא `warn` על הסיבה הצפויה, ו-`warn` על מה שאינו צפוי: כל עוד המתודה
 * אינה קיימת באוצריא זה המצב של **כל** הפעלה, ואזהרה על מה שתמיד קורה היא מה
 * שמאמן אנשים להפסיק לקרוא את הקונסולה.
 *
 * אינה חוסמת דבר ואינה מפילה דבר — היא רצה בלי `await` ובולעת כל כשל. אבחון
 * שיכול להפיל מנייה גרוע מהיעדר אבחון.
 */
async function explainHostGap(call: typeof tryCall): Promise<void> {
  try {
    const info = await call<{ permissions?: unknown }>('app.getGrantedPermissions');
    const granted = Array.isArray(info?.permissions) ? (info.permissions as unknown[]) : null;
    const has = granted?.some((name) => name === FONTS_PERMISSION) ?? false;

    const head = '[otzaria-word] fonts.listInstalled לא החזירה רשימה — הבורר נופל למדידה מקורבת.';
    if (granted !== null && has) {
      console.info(`${head} ההרשאה ${FONTS_PERMISSION} מאושרת, ולכן זו גרסת אוצריא שאינה מכירה את המתודה — הצפוי כרגע.`);
      return;
    }
    console.warn(
      `${head} ` +
        (granted === null
          ? 'לא ניתן לקרוא את ההרשאות שאושרו.'
          : `ההרשאה ${FONTS_PERMISSION} אינה מאושרת. אושרו: ${granted.join(', ')}`),
    );
  } catch {
    /* אבחון שנכשל אינו סיבה לרעש נוסף */
  }
}

/**
 * כמה **מילי-שניות** של עבודה רצופה מותרות לפני ויתור על החוט.
 *
 * לפי זמן ולא לפי מספר שמות, וזה תוקן אחרי מדידה. מדידה של שם שאינו מותקן
 * זולה — הדפדפן אינו פותר אותו ונופל לבסיס מיד; מדידה של משפחה **מותקנת**
 * דורשת טעינת הגופן מהדיסק אל ה-renderer, ונמדדה ~15ms. פי שש, ולכן מספר שמות
 * קבוע אינו מייצג זמן קבוע: `YIELD_EVERY = 40` שהיה כאן נכתב על 79 מועמדים
 * שרובם החטאות, וברשימה של 287 משפחות מהמארח — שכולן קיימות — הוא הפך לשמונה
 * מקטעים של סביב שנייה.
 *
 * נמדד על 287 משפחות ב-renderer קר (`renderer` חדש בכל סבב, שאחרת הגופנים
 * כבר טעונים והכול זניח):
 *
 * | ויתור | מקטעים | הארוך | חציון | סה"כ |
 * |---|---|---|---|---|
 * | כל 40 שמות | 8 | **1357ms** | 1180ms | 7866ms |
 * | כל 8 שמות | 36 | 316ms | 156ms | 6218ms |
 * | **כל 8ms** | 227 | **68ms** | 22ms | 6458ms |
 * | כל 4ms | 244 | 174ms | 20ms | 6529ms |
 *
 * הסך אינו משתנה — הוא טעינת הגופנים, והוא נגבה כך או כך. מה שמשתנה הוא
 * הבלימה הארוכה ביותר, בסדר גודל.
 *
 * ולמה 8 ולא 4: משפחה מותקנת אחת עולה ~15ms, כלומר היא לבדה חורגת מכל תקציב
 * קטן ממנה — 4ms לא קנה מקטע קצר יותר, רק יותר ויתורים. 8ms הוא בערך פריים
 * אחד ב-120Hz, וזה הגבול שמתחתיו אין מה להרוויח.
 */
const YIELD_BUDGET_MS = 8;

/**
 * מסננת למה שהדפדפן באמת פותר — **על רשימת המועמדים בלבד**.
 *
 * מוותרת על החוט כל `YIELD_BUDGET_MS` של עבודה: מדידה של מאות משפחות היא
 * סינכרונית מטבעה, וללא הוויתור היא הייתה בולעת פריים שלם בזמן שהמשתמש מקליד.
 *
 * למה זה חל על שכבה 2 ולא על רשימת המארח — ראו `measureUnresolved`.
 */
async function keepAvailable(
  candidates: readonly InstalledFont[],
  available: (name: string) => boolean,
): Promise<InstalledFont[]> {
  const kept: InstalledFont[] = [];
  await forEachMeasured(candidates, (font) => {
    if (typeof font?.name === 'string' && available(font.name)) kept.push(font);
  });
  return kept;
}

/**
 * אילו מהמשפחות שהמארח דיווח עליהן הדפדפן **אינו** פותר — בלי למחוק אף אחת.
 *
 * ## מה שהיה כאן, ולמה זו הייתה טעות
 *
 * רשימת המארח עברה ב-`keepAvailable`, כלומר בדיוק באותה מסננת שמסננת את
 * רשימת המועמדים. שם היא נכונה; על רשימת המארח היא מחקה גופנים ש**מותקנים**.
 *
 * `fonts.listInstalled` מונה דרך GDI — משפחות של ארבעה סגנונות. וריאנט משקל
 * (`Light`, `Semilight`, `Demi`) או מקופל לתוך משפחת הבסיס, או מדווח כמשפחה
 * נפרדת שכרום אינו פותר: הוא מקלף אסימון סגנון מסוף השם ומחפש את השאר, וזה
 * עובד ל-`Calibri Light` ונכשל ל-`Segoe UI Semilight`.
 *
 * **נמדד** ב-Windows 10, מנייה אמיתית של GDI מול מדידת רוחב בכרום: מ-287
 * משפחות שהמארח מדווח, **43 אינן נפתרות לפי השם**, וחמש מהן עבריות —
 * `Guttman Kav-Light`, `Guttman Yad-Light`, `Guttman Haim-Condensed`,
 * `Segoe UI Semilight`, `Malgun Gothic Semilight`. כולן מותקנות במכונה, כולן
 * מופיעות ברשימת הגופנים של אוצריא עצמה, וכולן נעדרו מהבורר. זו התלונה
 * שהגיעה מהשטח, ושתי הדוגמאות שבה — `HadasaNew Light` ו-`Rashi` — הן מאותה
 * צורה בדיוק.
 *
 * זו גם הייתה חריגה מהתכנון של המיזוג עצמו: כל מקור אחר **שומר** שם שאינו
 * נפתר ורק מסמן אותו (`available` ב-font-options.ts — „השורה אינה נעלמת
 * בכוונה”). רק רשימת המארח נמחקה, ולכן `Aptos` שאינו מותקן בכלל הוצג
 * ו-`Guttman Kav-Light` שכן מותקן — לא.
 *
 * ## למה שכבה 2 כן ממשיכה להיות מסוננת
 *
 * ההבדל אינו טעם אלא מה שכל מקור **יודע**. המארח ספר את המכונה: שם שהוא מחזיר
 * מותקן, ואי-פתירה שלו היא פער בין GDI ל-DirectWrite — ומשהו שאפשר לתקן,
 * בבייטים מאוצריא (engine/picker-fonts.ts). רשימת המועמדים היא רשימה
 * מתוחזקת ביד של ניחושים, ושם אי-פתירה היא התשובה היחידה שיש לשאלה „האם זה
 * קיים כאן”.
 *
 * ## ולמה המדידה נשארת כאן ולא עוברת למיזוג
 *
 * המיזוג הוא `computed` סינכרוני. מדידה של משפחה **מותקנת** נמדדה ~15ms —
 * טעינת הגופן מהדיסק אל ה-renderer — ולכן 287 שמות שם היו בולמים סביב ארבע
 * שניות בכל פתיחת מסמך, מפני ש-`installDocumentFontAliases` מקדם את דור
 * המדידה ומרוקן את המטמון. כאן זה קורה פעם אחת, אסינכרונית, עם ויתור על החוט.
 *
 * המיזוג כן מודד מחדש — אבל רק את מי שהקבוצה הזאת מסמנת: 43 שמות ולא 287, וכל
 * אחד מהם **זול** (הדפדפן אינו פותר אותו ונופל לבסיס מיד, בלי טעינה מהדיסק).
 * זה מה שמאפשר לשורה להתעורר אחרי שהבייטים הוזרקו.
 */
async function measureUnresolved(
  reported: readonly InstalledFont[],
  available: (name: string) => boolean,
): Promise<Set<string>> {
  const unresolved = new Set<string>();
  await forEachMeasured(reported, (font) => {
    const name = typeof font?.name === 'string' ? font.name.trim() : '';
    if (name !== '' && !available(name)) unresolved.add(name.toLowerCase());
  });
  return unresolved;
}

/**
 * הלולאה שמוותרת על החוט, לשני הקוראים שמעליה.
 *
 * הוצאה מ-`keepAvailable` כשנוסף קורא שני: תקציב הוויתור הוא מדידה
 * (`YIELD_BUDGET_MS`), ושתי לולאות שמעתיקות אותו היו נפרדות ביום שהוא ישתנה.
 */
async function forEachMeasured(
  fonts: readonly InstalledFont[],
  visit: (font: InstalledFont) => void,
): Promise<void> {
  let blockStart = performance.now();
  for (let i = 0; i < fonts.length; i++) {
    visit(fonts[i]);
    // אחרי המדידה ולא לפניה: השם הראשון חייב להימדד לפני שיש מה למדוד עליו
    // זמן, ובדיקה לפני הלולאה הייתה מוותרת מיד על מקטע ריק.
    if (performance.now() - blockStart >= YIELD_BUDGET_MS && i < fonts.length - 1) {
      await yieldToBrowser();
      blockStart = performance.now();
    }
  }
}

function yieldToBrowser(): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, 0);
  });
}

/**
 * מהרשימה הגולמית לצורה שהבורר צורך.
 *
 * שני סינונים שנראים מיותרים ואינם: שם ריק, ושם שמתחיל ב-`@` — הווריאנטים
 * האנכיים של CJK ש-GDI מונה. המפרט מבקש מהמארח לדלג עליהם, אבל רשימה שמגיעה
 * מבחוץ אינה מקום לסמוך על מפרט.
 *
 * המיון ב-`localeCompare` עם `'he'`: מיון בינארי היה מפזר את השמות העבריים
 * אחרי כל הלטיניים ובסדר שאינו א״ב.
 */
function toSnapshot(
  fonts: readonly InstalledFont[],
  source: 'host' | 'measured',
  unresolved: ReadonlySet<string> = new Set<string>(),
): InstalledFontsSnapshot {
  const families: FontFamilyOption[] = [];
  const hebrew = new Set<string>();
  const seen = new Set<string>();

  for (const font of fonts) {
    const name = typeof font?.name === 'string' ? font.name.trim() : '';
    if (name === '' || name.startsWith('@')) continue;

    const key = name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);

    if (Array.isArray(font.scripts) && font.scripts.includes('hebrew')) hebrew.add(key);

    // הגיבוי בערימה נגזר מ-`monospace` ולא קבוע: שם שהדפדפן לא יפתור בכל זאת
    // ייראה בבורר קרוב יותר למה שהוא באמת.
    const fallback = font.monospace === true ? 'monospace' : 'sans-serif';
    families.push({
      value: name,
      label: name,
      previewFamily: `${JSON.stringify(name)}, ${fallback}`,
    });
  }

  families.sort((a, b) => a.label.localeCompare(b.label, 'he'));
  return { families, hebrew, unresolved, source };
}
