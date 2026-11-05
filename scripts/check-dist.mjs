/**
 * בדיקת שפיות על תיקיית dist, לפני אריזה.
 *
 * שני האילוצים שהתוסף חייב לעמוד בהם (docs/word-plugin-implementation-plan.md §2, §18):
 * הפלט הוא סקריפטים קלאסיים בלבד — WebView2 אינו טוען <script type="module">
 * מ-file:// — והכול מקומי, בלי רשת. הבדיקה נכשלת על הפרה של אלה, ומדפיסה
 * גדלים כעדות לשער B.
 */
import { readFileSync, writeFileSync, existsSync, statSync, readdirSync, rmSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import { Script } from 'node:vm';

const DIST = join(dirname(fileURLToPath(import.meta.url)), '..', 'dist');
const errors = [];
const warnings = [];

if (!existsSync(DIST)) {
  console.error('dist אינו קיים — הריצו npm run build תחילה');
  process.exit(1);
}

const indexPath = join(DIST, 'index.html');
if (!existsSync(indexPath)) errors.push('חסר dist/index.html');

const html = existsSync(indexPath) ? readFileSync(indexPath, 'utf8') : '';

if (/<script[^>]*\btype=("|')module\1/.test(html)) {
  errors.push('dist/index.html מכיל <script type="module"> — WebView2 לא יטען אותו מ-file://');
}
if (/\bcrossorigin\b/.test(html)) {
  errors.push('dist/index.html מכיל crossorigin — מיותר ב-file:// ומעורר בקשה חוצת-מקור');
}
if (/<link[^>]*rel=("|')modulepreload\1/.test(html)) {
  errors.push('dist/index.html מכיל modulepreload — ה-chunks נטענים כסקריפטים קלאסיים');
}

let bundleReport;
try {
  bundleReport = JSON.parse(readFileSync(join(DIST, 'bundle-report.json'), 'utf8'));
  if (bundleReport.format !== 'classic-amd' || bundleReport.entry !== 'assets/app.js') throw new Error('unexpected format');
} catch (error) {
  errors.push(`דוח ה-chunks חסר או לא תקין: ${error.message}`);
}
const chunks = bundleReport?.chunks ?? {};
const startupFiles = bundleReport?.startupFiles ?? ['assets/engine-workers.js', 'assets/app.js'];
const preloadFiles = bundleReport?.preloadFiles ?? [];
if (!preloadFiles.length) errors.push('חסרה תוכנית preload למנוע DOCX — העלייה תוסיף שרשרת טעינת סקריפטים');
for (const file of preloadFiles) {
  if (!chunks[file] || !startupFiles.includes(file)) errors.push(`preload חסר בסגור העלייה: ${file}`);
}
if (!html.includes('modules.preload(')) errors.push('הטוען אינו מתחיל את ה־preload שבדוח');
const startupSources = startupFiles.filter((file) => existsSync(join(DIST, file)))
  .map((file) => readFileSync(join(DIST, file), 'utf8')).join('\n');

for (const [file, chunk] of Object.entries(chunks)) {
  const full = join(DIST, file);
  if (!existsSync(full)) { errors.push(`חסר chunk: ${file}`); continue; }
  if (statSync(full).size !== chunk.bytes) errors.push(`גודל ${file} אינו תואם לדוח ה-build`);
  if (!html.includes(JSON.stringify(file))) errors.push(`${file} חסר במפת הטוען הקלאסי`);
  for (const dependency of [...chunk.imports, ...chunk.dynamicImports]) {
    if (!chunks[dependency]) errors.push(`${file} מפנה ל-chunk חסר/חיצוני: ${dependency}`);
  }
}

// These public engine modules must still be packaged, but outside the startup
// closure. Disabling the PDF build path would pass a size-only assertion while
// silently removing the EMF/PDF fallback.
const deferredModules = [
  /pdfjs-dist\/build\/pdf\.mjs$/,
  /superdoc\/dist\/chunks\/pdf\.worker[^/]+\.js$/,
  /src\/ui\/panels\/MacrosDialog\.vue$/,
  /src\/ui\/panels\/OpenDocumentDialog\.vue$/,
  /src\/ui\/panels\/ShortcutManagerDialog\.vue$/,
];
for (const expected of deferredModules) {
  const matching = Object.entries(chunks).filter(([, chunk]) => chunk.modules.some((module) => expected.test(module.id)));
  if (!matching.length) errors.push(`רכיב לפי צורך חסר מהאריזה: ${expected}`);
  for (const [file] of matching) {
    if (startupFiles.includes(file)) errors.push(`${file} נטען בעלייה אף שהוא מכיל ${expected}`);
  }
}

// כל href/src ב-HTML חייב להיות נכס מקומי שקיים בפועל.
for (const match of html.matchAll(/\b(src|href)=("|')([^"']+)\2/g)) {
  const attr = match[1];
  const url = match[3];
  if (/^(https?:)?\/\//i.test(url)) {
    errors.push(`dist/index.html מפנה לכתובת חיצונית ב-${attr}: ${url}`);
    continue;
  }
  if (url.startsWith('data:') || url.startsWith('#')) continue;
  const local = join(DIST, url.replace(/^\.?\//, '').split('?')[0]);
  if (!existsSync(local)) errors.push(`נכס חסר ב-dist: ${url}`);
}

/**
 * שרשרת הטעינה. שני הבאנדלים אינם `<script src>` ב-HTML יותר, אלא מוזרקים
 * מטוען inline אחרי הצביעה הראשונה (`deferredEntry` ב-vite.config.ts): סקריפט
 * חוסם ב-`<head>` עוצר את פריסת ה-HTML, כלומר את ה-`<body>` ובתוכו את מסך
 * הטעינה עצמו — מסך לבן מ-file://. המספר ב-`scripts/startup-probe.mjs`.
 *
 * מכאן שבדיקת הנכסים שמעל אינה מכסה אותם: הם מחרוזות בתוך JS ולא תכונות src.
 * מי שישבור חוליה יקבל תוסף שעולה, מצייר מסך טעינה — ונתקע עליו בלי שגיאה.
 * לכן כל חוליה נבדקת בשמה, וגם הסדר: engine-workers.js מציב את
 * `__SUPERDOC_WORKER_SOURCES__`, ו-engineWorkerUrls() נצרך בהקמת המנוע.
 */
const workersAt = html.indexOf('./assets/engine-workers.js');
const appAt = html.indexOf('./assets/app.js');
if (workersAt === -1) {
  errors.push('dist/index.html אינו מזריק את ./assets/engine-workers.js — המנוע יקום בלי workers');
}
if (appAt === -1) {
  errors.push('dist/index.html אינו מזריק את ./assets/app.js — התוסף לא ייטען כלל');
}
if (workersAt !== -1 && appAt !== -1 && workersAt > appAt) {
  errors.push('engine-workers.js מוזרק אחרי app.js — המנוע יקום בלי ה-workers');
}

const headEnd = html.indexOf('</head>');
if (/<script[^>]*\bsrc=/.test(headEnd === -1 ? html : html.slice(0, headEnd))) {
  errors.push('יש <script src> ב-<head> של dist/index.html — הוא חוסם את ציור מסך הטעינה');
}

/**
 * מסך הטעינה. הוא כל מה שהמשתמש רואה בשנייה-שתיים הראשונות, והוא inline
 * ב-HTML בדיוק כדי שלא יהיה תלוי בבאנדל שהוא מכסה עליו. אם הוא ייעלם, כל
 * הקריאות ב-`src/host/splash.ts` יהפכו ל-no-op **בשקט** — הממשק יעלה כרגיל,
 * והפתיחה תחזור להיות מסך לבן בלי ששום דבר ייכשל.
 */
if (!html.includes('__otzariaSplash')) {
  errors.push('ה-API של מסך הטעינה אינו ב-dist/index.html — הפתיחה תחזור להיות מסך לבן');
}
if (!html.includes('id="otzaria-splash"')) {
  errors.push('חסרה תגית מסך הטעינה ב-dist/index.html');
}

/**
 * ה-latch של plugin.boot חייב להיות הסקריפט הראשון בדף. אוצריא משגרת את
 * האירוע פעם אחת ואינה משחזרת אותו, וכל סקריפט שקודם ל-latch הוא חלון שבו
 * הוא יכול ללכת לאיבוד — וזה בדיוק הכשל שנצפה: „אוצריא לא סיימה לאתחל”
 * בטעינה ראשונה, ותוסף שעולה רק אחרי רענון.
 */
const latchAt = html.indexOf('__otzariaBoot');
if (latchAt === -1) {
  errors.push('ה-latch של plugin.boot אינו ב-dist/index.html — התוסף ייתקע על אתחול');
} else if (workersAt !== -1 && latchAt > workersAt) {
  errors.push('ה-latch של plugin.boot בא אחרי engine-workers.js — 5MB לפני ההרשמה לאירוע');
} else if (appAt !== -1 && latchAt > appAt) {
  errors.push('ה-latch של plugin.boot בא אחרי app.js');
}

// CDN-ים שמנועי צד-שלישי נוטים ליפול אליהם כברירת מחדל. אינם נכשלים
// אוטומטית: מחרוזת בבאנדל אינה בקשה. הן נרשמות כדי שייבדקו ידנית בשער A.
const CDN_HINTS = ['cdn.jsdelivr.net', 'unpkg.com', 'cdnjs.cloudflare.com', 'fonts.googleapis.com'];

/**
 * סעיף 3.1(c) ברישיון מנוע ה-DOCX אוסר להסיר או להסתיר הודעות רישוי. המינימיזציה
 * מוחקת הערות כברירת מחדל, ולכן זו בדיקה חוסמת ולא אזהרה: הקובץ שמכיל את המנוע
 * חייב לשאת את הבאנר שלו.
 */
const ENGINE_LICENSE_MARK = 'DOCX Engine Proprietary License Agreement';
const ENGINE_BEARING_FILES = ['assets/engine-workers.js', ...Object.entries(chunks)
  .filter(([, chunk]) => chunk.modules.some((module) => module.renderedLength > 0 && /@superdoc\/docx-engine\//.test(module.id)))
  .map(([file]) => file)];
if (ENGINE_BEARING_FILES.length < 2) errors.push('לא נמצא chunk של מנוע המסמכים בדוח ה-build');

/**
 * ה-path data של אייקוני Fluent System Icons מוטמע ב-src/ui/icons/icons.ts,
 * ורישיון ה-MIT מחייב שהודעת הרישוי תופץ עם כל עותק. ההודעה אינה קובץ נפרד
 * אלא באנר legal comment בראש הקובץ, שנשען על esbuild.legalComments: 'eof'
 * כדי לשרוד את המינימיזציה. אם ההגדרה תיפול, הבאנר ייעלם בשקט — הבנייה
 * תצליח, וההפצה תהיה בהפרה. לכן שער ולא הערה.
 */
const ICONS_LICENSE_MARK = 'Fluent System Icons — MIT';

/**
 * המילון התורני חייב להיות **נכס נפרד**, ו`assets/app.js` חייב לא להכיל
 * אותו. זו לא בדיקת שפיות אלא השער היחיד שתופס את הנסיגה: מי שיחזיר את
 * `src/data/torah-dictionary.txt` ל-`import` יקבל אריזה שעובדת בדיוק כמו
 * קודם — `inlineDynamicImports` בולע גם `await import()` — רק ש-1.3MB
 * נכנסים לבאנדל הראשי שכל משתמש פורס בעלייה, כולל מי שהתכונה כבויה אצלו.
 * שער שבודק רק „האם המילון נטען כשהדליקו” היה נשאר ירוק על זה.
 *
 * הסמן הוא ערך אמיתי מהמילון שאין שום סיבה אחרת שיופיע בקוד.
 */
const DICTIONARY_ASSET = 'assets/torah-dictionary.js';
const DICTIONARY_MARKER = 'אאוגרייהו';
const dictionaryPath = join(DIST, DICTIONARY_ASSET);

if (!existsSync(dictionaryPath)) {
  errors.push(`חסר ${DICTIONARY_ASSET} — בדיקת האיות לא תמצא מילון`);
} else if (!readFileSync(dictionaryPath, 'utf8').includes(DICTIONARY_MARKER)) {
  errors.push(`${DICTIONARY_ASSET} אינו מכיל את המילון — התוסף שנארז חסר את הנתונים`);
}

if (startupSources.includes(DICTIONARY_MARKER)) {
  errors.push(
    'קובצי העלייה מכילים את המילון התורני — הוא נבלע בעלייה במקום להישאר נכס נפרד. ' +
      'כל משתמש פורס עכשיו 1.3MB בעלייה בשביל תכונה שברירת המחדל שלה כבויה.',
  );
}

/**
 * אותו שער בדיוק לשני נכסי ההשלמה, ומאותה סיבה: הם 0.75MB יחד, התכונה שלהם
 * כבויה כברירת מחדל, ו-`import` ישיר של קובץ ה-JSON היה מכניס אותם לבאנדל
 * הראשי בלי ששום דבר ייכשל. כל סמן הוא ערך אמיתי מהנתונים שאין סיבה אחרת
 * שיופיע בקוד, ואינו קיים באף אחד מהנכסים האחרים.
 */
const LAZY_ASSETS = [
  { file: 'assets/acronyms.js', markers: ['רבי משה בן מימון'], what: 'מילון ראשי-התיבות', size: '0.7MB' },
  {
    file: 'assets/static-completion.js',
    // סמן לכל קובץ נתונים ולא לכל נכס: השניים נארזים יחד, ו-`import` ישיר של
    // `authors.json` בלבד (20KB) היה עובר בשקט תחת סמן מהביטויים.
    markers: ['אבן שואבת', 'אבן גבאי, מאיר בן יחזקאל'],
    what: 'רשימות ההשלמה הסטטיות',
    size: '32KB',
  },
];

for (const { file, markers, what, size } of LAZY_ASSETS) {
  const path = join(DIST, file);
  const asset = existsSync(path) ? readFileSync(path, 'utf8') : null;
  if (asset === null) {
    errors.push(`חסר ${file} — ${what} לא ייטענו`);
  }
  for (const marker of markers) {
    if (asset !== null && !asset.includes(marker)) {
      errors.push(`${file} אינו מכיל את „${marker}” — התוסף שנארז חסר חלק מ${what}`);
    }
    if (startupSources.includes(marker)) {
      errors.push(
        `קובצי העלייה מכילים את ${what} — הם נבלעו בעלייה במקום להישאר נכס נפרד. ` +
          `כל משתמש פורס עכשיו ${size} בעלייה בשביל תכונה שברירת המחדל שלה כבויה.`,
      );
    }
  }
}

const files = [];
function walk(dir, prefix = '') {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const rel = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isDirectory()) walk(join(dir, entry.name), rel);
    else files.push(rel);
  }
}
walk(DIST);

console.log('גדלים ב-dist:');
let total = 0;
for (const rel of files.sort()) {
  const size = statSync(join(DIST, rel)).size;
  total += size;
  console.log(`  ${rel.padEnd(34)} ${(size / 1024 / 1024).toFixed(2)} MB`);
}
console.log(`  ${'סה"כ'.padEnd(34)} ${(total / 1024 / 1024).toFixed(2)} MB`);

for (const rel of files) {
  if (!rel.endsWith('.js')) continue;
  const full = join(DIST, rel);
  try {
    // Parse as a classic script explicitly; node --check under type:module
    // would otherwise accept an accidental ESM chunk.
    new Script(readFileSync(full, 'utf8'), { filename: rel });
  } catch {
    errors.push(`${rel} אינו סקריפט קלאסי תקין`);
  }
  const text = readFileSync(full, 'utf8');
  // A previous hashed chunk must not survive into an otzplugin archive, even
  // when the current loader cannot request it. A dirty output directory would
  // otherwise silently ship both engine versions and inflate the package.
  if (/window\.__otzariaDefine\s*\(/.test(text) && !chunks[rel]) {
    errors.push(`${rel} הוא chunk שאינו במפת ה-build — יש לנקות פלט ישן לפני אריזה`);
  }
  for (const hint of CDN_HINTS) {
    if (text.includes(hint)) warnings.push(`${rel} מכיל את המחרוזת ${hint}`);
  }
}

for (const rel of ENGINE_BEARING_FILES) {
  const full = join(DIST, rel);
  if (!existsSync(full)) {
    errors.push(`${rel} חסר ב-dist`);
    continue;
  }
  if (!readFileSync(full, 'utf8').includes(ENGINE_LICENSE_MARK)) {
    errors.push(
      `${rel} אינו נושא את באנר הרישוי של מנוע ה-DOCX — ` +
        "בדקו את esbuild.legalComments ב-vite.config.ts",
    );
  }
}

const iconFiles = Object.entries(chunks)
  .filter(([, chunk]) => chunk.modules.some((module) => /src\/ui\/icons\/icons\.ts$/.test(module.id)))
  .map(([file]) => file);
if (!iconFiles.length) errors.push('לא נמצא chunk האייקונים בדוח ה-build');
for (const rel of iconFiles) {
  const full = join(DIST, rel);
  if (existsSync(full) && !readFileSync(full, 'utf8').includes(ICONS_LICENSE_MARK)) {
    errors.push(
      `${rel} אינו נושא את באנר ה-MIT של Fluent System Icons — ` +
        "בדקו את esbuild.legalComments ב-vite.config.ts",
    );
  }
}

/**
 * הגופן הארוז (src/styles/fonts.ts). ה-@font-face נבנה כמחרוזת ומוזרק בזמן
 * ריצה — דווקא כדי לא לעבור דרך פותר הנכסים של Vite — ולכן שינוי שם קובץ או
 * נתיב אינו מפיל את הבנייה. הרשימה כאן חוזרת על עצמה בכוונה: שער צריך להצהיר
 * את הציפייה בעצמו, אחרת הוא מאמת את הקוד מול הקוד.
 *
 * רישיון ה-OFL מחייב שנוסח הרישיון יופץ עם הגופן, ולכן גם הוא נבדק.
 */
const FONT_FILES = [
  'Assistant-Regular.ttf',
  'Assistant-Medium.ttf',
  'Assistant-SemiBold.ttf',
  'Assistant-Bold.ttf',
];
for (const file of FONT_FILES) {
  if (!existsSync(join(DIST, 'fonts', file))) errors.push(`חסר גופן ב-dist: fonts/${file}`);
  if (!startupSources.includes(`./fonts/${file}`)) {
    errors.push(`קובצי העלייה אינם מפנים ל-./fonts/${file} — ההצהרה והנכס יצאו מסינכרון`);
  }
}

if (!existsSync(join(DIST, 'third-party/ASSISTANT-LICENSE.txt'))) {
  errors.push('חסר third-party/ASSISTANT-LICENSE.txt — ה-OFL מחייב להפיץ את הרישיון עם הגופן');
}

// שערי הבדיקה כותבים דפי HTML זמניים לתוך dist (scripts/font-check.html
// מועתק לשם ביד, scripts/boot-check.mjs כותב ומוחק). דף שנשאר שם אחרי קריסה
// נארז לתוך התוסף, ולכן כל HTML שאינו index.html הוא שגיאה.
for (const rel of files) {
  if (rel.endsWith('.html') && rel !== 'index.html') {
    errors.push(`dist/${rel} אינו חלק מהתוסף — דף בדיקה שנשכח; יש למחוק לפני אריזה`);
  }
}

/**
 * קוד ה-workers יושב בתוך מחרוזות JSON, ולכן `node --check` על הקובץ העוטף
 * אינו נוגע בו כלל — הוא בודק שורת השמה אחת. אלה 4.9MB שנטענים בפועל
 * כ-workers קלאסיים, וכשל שלהם פירושו תוסף שלא פותח מסמכים; לכן הם נפרסים
 * ונבדקים בנפרד, ומול אותן חתימות ESM שה-build אוכף.
 */
const ESM_SIGNATURES = [
  /^\s*import\s+[\w{*'"]/m,
  /^\s*export\s+(?:default|const|let|var|function|class|\{)/m,
  /\bimport\.meta\b/,
  /^\s*import\s*\(/m,
];

const workersFile = join(DIST, 'assets/engine-workers.js');
if (existsSync(workersFile)) {
  const wrapper = readFileSync(workersFile, 'utf8');
  // הצורה היא `JSON.parse("…")` ולא אובייקט ליטרלי, ובכוונה: אלה 5MB שהמנתח
  // של JavaScript היה פורס כתחביר. השכבה החיצונית היא מחרוזת JS, ולכן הפירוק
  // כאן הוא בשני שלבים — JSON.parse על הליטרל, ואז על מה שהוא החזיר.
  const json = wrapper.match(
    /^window\.__SUPERDOC_WORKER_SOURCES__ = JSON\.parse\(([\s\S]*)\);\s*$/,
  );

  if (!json) {
    errors.push('assets/engine-workers.js אינו בצורה המצופה — לא ניתן לבדוק את קוד ה-workers');
  } else {
    let sources;
    try {
      sources = JSON.parse(JSON.parse(json[1]));
    } catch (error) {
      errors.push(`קוד ה-workers אינו JSON תקין: ${error.message}`);
    }

    const roles = sources ? Object.keys(sources) : [];
    for (const role of ['document', 'reviewIndex']) {
      if (!roles.includes(role)) errors.push(`חסר קוד worker לתפקיד ${role}`);
    }

    for (const [role, code] of Object.entries(sources ?? {})) {
      const tmp = join(DIST, `.worker-check-${role}.js`);
      writeFileSync(tmp, code);
      try {
        execFileSync(process.execPath, ['--check', tmp], { stdio: 'pipe' });
      } catch {
        errors.push(`קוד ה-worker "${role}" אינו סקריפט קלאסי תקין (node --check נכשל)`);
      } finally {
        rmSync(tmp, { force: true });
      }

      const esm = ESM_SIGNATURES.filter((pattern) => pattern.test(code));
      if (esm.length) {
        errors.push(
          `קוד ה-worker "${role}" מכיל תחביר ESM ולכן לא ייטען כ-worker קלאסי, ` +
            'וב-file:// אין חלופה',
        );
      }

      console.log(`  worker ${role}: ${(code.length / 1024 / 1024).toFixed(2)} MB, קלאסי`);
    }
  }
}

for (const w of new Set(warnings)) console.warn(`אזהרה: ${w}`);
if (errors.length) {
  for (const e of errors) console.error(`שגיאה: ${e}`);
  process.exit(1);
}
console.log('dist תקין: סקריפטים קלאסיים, כל הנכסים מקומיים.');
