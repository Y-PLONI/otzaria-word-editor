import { defineConfig, type Plugin } from 'vite';
import vue from '@vitejs/plugin-vue';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { TORAH_DICTIONARY_FILE, TORAH_DICTIONARY_GLOBAL } from './src/engine/spellcheck';
import { ACRONYMS_FILE } from './src/engine/acronyms-constants';
import { buildAcronymsAsset } from './scripts/acronyms-asset';
import { STATIC_COMPLETION_FILE } from './src/engine/static-completion-constants';
import { buildStaticCompletionAsset } from './scripts/static-completion-asset';
import { patchBlankDocumentXml, patchBlankStylesXml } from './src/engine/blank-document';
import { deriveHebrewBlankDocx } from './scripts/blank-docx';
import { superdocSelectionIdle } from './scripts/superdoc-selection-idle';
import { classicChunks, classicChunkFiles, classicStartupPlan } from './scripts/classic-chunks';
import { installClassicModules } from './scripts/classic-modules';

/**
 * ל-WebView2 של Windows אין תמיכה ב-<script type="module"> מ-file:// ,
 * ואוצריא טוענת תוסף ארוז בדיוק משם. ה-build מפיק chunks של AMD, שנרשמים
 * ומורצים דרך סקריפטים קלאסיים מקומיים. במצב dev התגית נשארת module כדי
 * שה-HMR של Vite ימשיך לעבוד.
 *
 * ומעבר לכך: Vite מזריק את תגית הכניסה ל-`<head>`, וסקריפט קלאסי שם חוסם את
 * פריסת ה-HTML — כלומר ה-`<body>`, ובתוכו מסך הטעינה, אינו נפרס עד ששני
 * הבאנדלים (‏17MB יחד) נפרסו והורצו. זה בדיוק המסך הלבן שנמדד
 * ב-`scripts/startup-probe.mjs` — שם המספר, ורק שם. שום דבר לא נכשל; פשוט
 * לא היה מה לראות.
 *
 * לכן שתי התגיות מוסרות מה-HTML, ובמקומן נכנס טוען inline שמזריק אותן אחרי
 * הצביעה הראשונה — ומדווח למסך הטעינה בין השלבים. הצביעה ירדה ל-50ms.
 */
function deferredEntry(): Plugin {
  const WORKERS_SRC = './assets/engine-workers.js';
  const ENTRY = /[ \t]*<script\s+src="(\.\/assets\/app\.js)"><\/script>\n?/;

  return {
    name: 'otzaria-deferred-entry',
    apply: 'build',
    // אחרי inlineEngineWorkers: התגית שלו כבר בדף, וכאן היא מוסרת יחד עם
    // תגית הכניסה ומוחלפת בטוען.
    enforce: 'post',

    transformIndexHtml(html, context) {
      if (!context.bundle) throw new Error('Classic chunk manifest is unavailable');
      const moduleFiles = classicChunkFiles(context.bundle);
      const { preloadFiles } = classicStartupPlan(context.bundle, { preloadDocumentEngine: true });
      const classic = html.replace(/\s+type="module"/g, '').replace(/\s+crossorigin/g, '');

      const match = classic.match(ENTRY);
      if (!match) {
        // בלי התגית אין מה לדחות, ותוסף בלי app.js הוא מסך טעינה לנצח.
        throw new Error(
          'לא נמצאה תגית הכניסה assets/app.js ב-index.html — ' +
            'ייתכן ש-entryFileNames או צורת ההזרקה של Vite השתנו.',
        );
      }

      const withoutTags = classic
        .replace(ENTRY, '')
        .replace(new RegExp(`[ \\t]*<script src="${WORKERS_SRC.replace(/[./]/g, '\\$&')}"></script>\\n?`), '');

      const loader = `    <script>
      /* טוען הכניסה.

         שני פריימים ואז הזרקה: הראשון מתזמן ציור, השני רץ אחרי שהוא הושלם —
         כלומר מסך הטעינה כבר על המסך כשהבאנדלים מתחילים להיפרס.

         ה-setTimeout אינו חגורה כפולה מיותרת: אוצריא עשויה להקים את ה-WebView
         של התוסף כשהוא עדיין אינו נראה, וב-Chromium requestAnimationFrame
         בדף מוסתר אינו נורה כלל. בלי השעון הזה תוסף שנפתח ברקע לא היה נטען
         לעולם. מי שמגיע ראשון מנצח; השני נבלע.

         „async = false” על אלמנט שמוזרק ב-JS הוא מה שמחזיק את סדר ההרצה:
         engine-workers.js מציב את __SUPERDOC_WORKER_SOURCES__, ו-app.js צורך
         אותו בהקמת המנוע. בלעדיו הדפדפן מריץ לפי סדר ההגעה — ואלה שני קבצים
         בגדלים שונים מאוד. ההורדה עצמה נשארת מקבילה, כי שתי התגיות נכנסות
         באותו tick. */
      (function () {
        var started = false;
        function load() {
          if (started) return;
          started = true;
          var splash = window.__otzariaSplash;
          var modules = window.__otzariaModules = (${installClassicModules.toString()})(${JSON.stringify(moduleFiles)});
          function say(stage) {
            if (splash && stage) splash.set(stage.at, stage.text);
          }
          /* התחנה מדווחת על השלב ש**מתחיל**, לא על זה שנגמר.

             דיווח ב-onload של app.js נעלם בבנייה הישנה: ה-IIFE הריץ את
             main.ts לפני onload, והתחנה 68 הקדימה את התחנה 55. נמדד ברצף שהוצג
             בפועל: מתחיל · טוען את מנוע המסמכים… · מכין את סביבת העריכה… ·
             פותח את המסמך… · מוכן — התחנה שבאמצע לא הופיעה כלל.

             גם ב-AMD התחנה 55 נשארת לפני הרכבת הממשק. load של chunk מודד
             רישום factory בלבד; מדד module:ready נפרד מסמן את הרצת התלויות
             וה-factory, ו-main.ts ממשיך לדווח על הרכבת המעטפת. */
          /* חימום worker המסמך, ברגע ש-engine-workers.js הציב את המקורות —
             כלומר בזמן שקובצי הממשק עוד נפרסים. ‏Worker זמני על ה-blob URL
             מקמפל את ~4.6MB קוד ה-worker במקביל לפריסה הזאת, ומושלך ברגע
             שאיתת שעלה; כשהמנוע יקים Worker על **אותו URL** הקומפילציה כבר
             חמה (נמדד: ‏~210ms במקום ~325ms). לכן ה-URL נשמר על window,
             ו-engine/workers.ts מאמץ אותו במקום לבנות URL חדש — blob אחר,
             גם עם אותו תוכן, אינו פוגע ב-cache.

             חימום שנכשל אינו נוגע בעלייה: ה-catch בולע, ו-workers.ts בונה
             את ה-URL בעצמו כשהגלובל חסר. השם __otzariaDocWorkerUrl משותף
             עם src/engine/workers.ts. */
          function warm() {
            try {
              var sources = window.__SUPERDOC_WORKER_SOURCES__;
              if (!sources || !sources.document || typeof Worker !== 'function') return;
              var url = URL.createObjectURL(new Blob([sources.document], { type: 'text/javascript' }));
              window.__otzariaDocWorkerUrl = url;
              // Worker קלאסי, ובמכוון — אל תוסיפו כאן type: 'module'.
              //
              // ה-code cache ממופתח גם בסוג הסקריפט, ולכן probe מסוג אחר
              // מזה שהמנוע יוצר לא היה מחמם דבר. מה המנוע יוצר בפועל נמדד
              // (עטיפת window.Worker לפני העלייה, על ה-dist הארוז מ-file://):
              // Worker **קלאסי** בשם superdoc-v2-edit, על אותו URL בדיוק.
              // ובאותה מדידה: Worker מסוג module על blob URL טרי נכשל כאן
              // מיד (‏~10ms, גם על סקריפט של 14 בתים) — module workers מ-blob
              // חסומים ב-origin האטום של file://. כלומר probe כזה היה מת
              // בשקט, ומבטל את החימום כולו.
              var probe = new Worker(url);
              function drop() { try { probe.terminate(); } catch (ignored) {} }
              probe.addEventListener('message', drop, { once: true });
              probe.addEventListener('error', drop, { once: true });
              // רשת ביטחון בלבד: ‏ה-worker מודיע מיוזמתו תוך ~70ms (נמדד),
              // ואז הוא נזרק שם. הקומפילציה נשמרת ב-cache גם אחרי שהוא מת.
              setTimeout(drop, 10000);
            } catch (ignored) { /* חימום הוא קיצור, לא תנאי */ }
          }
          say({ at: 22, text: 'טוען את מנוע המסמכים…' });
          [
            { src: '${WORKERS_SRC}', next: { at: 55, text: 'מרכיב את הממשק…' }, warm: true },
            { src: '${match[1]}', next: null }
          ].forEach(function (step) {
            if (step.src === '${match[1]}') {
              modules.load('assets/app').catch(function (error) {
                console.error('[otzaria-word] Local module load failed:', error);
                if (splash) splash.fail('טעינת קוד התוסף נכשלה');
              });
              return;
            }
            var script = document.createElement('script');
            script.async = false;
            script.src = step.src;
            function measured(event) {
              try { window.__otzariaPerformance?.mark(event, { asset: step.warm ? 'engine-workers.js' : 'app.js' }); } catch (ignored) {}
            }
            script.addEventListener('load', function () {
              measured('script.loaded');
              say(step.next);
              if (step.warm) warm();
            });
            script.addEventListener('error', function () {
              measured('script.failed');
              if (splash) splash.fail('טעינת קוד התוסף נכשלה');
            });
            measured('script.requested');
            document.head.appendChild(script);
          });
          // Fetch/register the mandatory DOCX engine alongside the entry. Its
          // factory still waits for SuperDoc's normal dynamic load. Insertion
          // stays workers → app → preloads; a speculative failure is handled,
          // and the eventual consumer can retry through the same loader.
          ${JSON.stringify(preloadFiles)}.forEach(function (file) {
            modules.preload(file.slice(0, -3)).catch(function () {});
          });
        }
        requestAnimationFrame(function () {
          requestAnimationFrame(load);
        });
        setTimeout(load, 120);
      })();
    </script>
`;

      return withoutTags.replace('</body>', `${loader}  </body>`);
    },
  };
}

/** איזה worker של המנוע ממופה לאיזה שדה ב-config.workerUrls של SuperDoc. */
const WORKER_ROLES: Array<{ match: string; role: 'document' | 'reviewIndex' | 'drop' }> = [
  { match: 'browser-worker-entry', role: 'document' },
  { match: 'review-index-worker-entry', role: 'reviewIndex' },
  // worker השיתופיות אינו נארז: התוסף עובד אופליין וללא הרשאת רשת, ולכן
  // הוא לעולם לא נטען — ו-5MB זה מחיר שאין סיבה לשלם.
  { match: 'collaboration-worker-entry', role: 'drop' },
];

/**
 * ה-workers של המנוע מוטמעים באריזה כמחרוזות, ובזמן ריצה נבנה מהן blob: URL
 * שנמסר ל-`config.workerUrls`.
 *
 * ההטמעה נשארת גם אחרי פיצול ה-chunks: module workers אינם נטענים מ-file://.
 * מקורות ה-workers נארזים כ-IIFE והכתובות המקומיות נמסרות ב-workerUrls,
 * כדי ש-engine/workers.ts יוכל להפעיל אותם כ-workers קלאסיים מ-blob.
 * המדידות המלאות, כולל למה blob ולא data:, ב-docs/spike.md §שער A.
 *
 * הפלט הוא `JSON.parse('…')` ולא אובייקט ליטרלי, וזה אינו סגנון: אלה 5MB
 * שהמנתח של JavaScript היה פורס כתחביר — ליטרל אחד ענק עם escaping — מול
 * מנתח JSON ייעודי שמקבל מחרוזת אחת. נמדד ב-scripts/startup-probe.mjs:
 * זמן ההרצה של הקובץ ירד בערך למחצית.
 */
function inlineEngineWorkers(): Plugin {
  return {
    name: 'otzaria-inline-engine-workers',
    apply: 'build',
    enforce: 'post',

    generateBundle(_options, bundle) {
      const sources: Record<string, string> = {};

      for (const [fileName, output] of Object.entries(bundle)) {
        const spec = WORKER_ROLES.find((w) => fileName.includes(w.match));
        if (!spec) continue;

        if (spec.role !== 'drop') {
          sources[spec.role] = output.type === 'chunk' ? output.code : String(output.source);
        }
        delete bundle[fileName];
      }

      const missing = WORKER_ROLES.filter((w) => w.role !== 'drop' && !(w.role in sources));
      if (missing.length) {
        // כשל שקט כאן פירושו תוסף ארוז שלא פותח מסמכים — עדיף להפיל את ה-build.
        this.error(
          `לא נמצאו קובצי worker של מנוע ה-DOCX: ${missing.map((m) => m.match).join(', ')}. ` +
            'ייתכן ששמות הנכסים השתנו בגרסת superdoc חדשה — יש לעדכן את WORKER_ROLES.',
        );
      }

      // מחרוזת JSON בתוך ליטרל JS: JSON.stringify פעמיים — הפנימי בונה את
      // ה-JSON, החיצוני הופך אותו למחרוזת JS חוקית עם כל ה-escaping.
      const payload = JSON.stringify(JSON.stringify(sources));
      this.emitFile({
        type: 'asset',
        fileName: 'assets/engine-workers.js',
        source: `window.__SUPERDOC_WORKER_SOURCES__ = JSON.parse(${payload});\n`,
      });
    },

    transformIndexHtml(html) {
      // התגית מוזרקת כאן, ו-deferredEntry (שרץ אחריו) מחליף אותה ואת תגית
      // הכניסה בטוען אחד. ההזרקה היא לפני הסקריפט הראשון שיש לו src, ולא לפני
      // ה-`<script` הראשון: ה-latch של plugin.boot הוא סקריפט inline ב-head
      // וחייב להישאר ראשון.
      return html.replace(
        /<script([^>]*\bsrc=)/,
        '<script src="./assets/engine-workers.js"></script>\n    <script$1',
      );
    },
  };
}

/**
 * המילון התורני יוצא כנכס נפרד — לא לתוך `assets/app.js`.
 *
 * לפני פיצול ה-build, `inlineDynamicImports: true` בלע גם ייבוא דינמי: 1.3MB שכל
 * משתמש פורס בעלייה בשביל תכונה שברירת המחדל שלה כבויה. הנכס הנפרד נשאר ונטען
 * בהזרקת `<script>` בזמן ריצה, ורק כשהמשתמש מדליק את הבדיקה — ראו
 * `src/engine/spellcheck-dictionary.ts`, שם גם ההסבר למה `<script>` ולא
 * `fetch` (‏`file://` עם origin opaque).
 *
 * הפלט הוא ליטרל תבנית (backtick) ולא `JSON.parse`: המילון הוא עברית, גרש
 * וגרשיים בלבד — 29 תווים שאף אחד מהם אינו דורש escaping בליטרל תבנית —
 * ולכן `\n` נשאר תו אחד במקום ארבעה בתים של `\\n`. ההפרש נמדד: 1.31MB מול
 * 1.55MB לאותם נתונים בדיוק. `build` מוודא שההנחה מחזיקה, כי תו בודד שיברח
 * (backtick, `$`, לוכסן הפוך) היה מייצר קובץ JS פגום — כלומר תוסף שנפרס
 * ונשבר, ולא בנייה שנכשלת — ובאותה הזדמנות גם את המיון ואת אורך הערכים,
 * שני דברים ש-`createDictionary` מניח ואיש אינו בודק.
 */
function torahDictionaryAsset(): Plugin {
  const SOURCE = fileURLToPath(new URL('./src/data/torah-dictionary.txt', import.meta.url));
  /** עברית, גרש וגרשיים ישרים, ומפריד השורות. שום דבר אחר. */
  const PACKABLE = /^[\u05D0-\u05EA"'\n]+$/;
  /** הערך הקצר ביותר שמותר. ערך בן תו אחד היה מכשיר כמעט כל מחרוזת. */
  const MIN_LENGTH = 2;

  /**
   * שלוש טענות על קובץ הנתונים, ולא רק זו שנוגעת לאריזה.
   *
   * ה-`PACKABLE` בלבד שומר על **בטיחות ליטרל התבנית** — שתו בורח לא ייצר
   * קובץ JS פגום. אבל `createDictionary` מניח שתי הנחות נוספות שאיש לא בדק:
   * שהרשימה **ממוינת לפי יחידות UTF-16** (בלי זה החיפוש הבינארי פשוט לא
   * מוצא חלק מהערכים, בשקט), ושאין בה שורות ריקות או ערכים בני תו אחד.
   * הקובץ הוא נתונים שנערכים ביד — מילה שתתווסף במקום הלא נכון היא בדיוק
   * הדבר שיישאר ירוק בכל בדיקה אחרת.
   */
  function build(): string {
    const packed = readFileSync(SOURCE, 'utf8').replace(/\r\n?/g, '\n').trim();

    if (!PACKABLE.test(packed)) {
      throw new Error(
        `${TORAH_DICTIONARY_FILE}: המילון מכיל תו שאינו עברית/גרש/גרשיים — ליטרל תבנית אינו בטוח עבורו. ` +
          'בדקו את src/data/torah-dictionary.txt.',
      );
    }

    const words = packed.split('\n');
    for (let i = 0; i < words.length; i++) {
      const word = words[i]!;
      if (word.length < MIN_LENGTH) {
        throw new Error(
          `${TORAH_DICTIONARY_FILE}: שורה ${i + 1} קצרה מ-${MIN_LENGTH} תווים (${JSON.stringify(word)}). ` +
            'ערך בן תו אחד או שורה ריקה מכשירים כמעט כל מחרוזת.',
        );
      }
      if (i > 0 && !(words[i - 1]! < word)) {
        throw new Error(
          `${TORAH_DICTIONARY_FILE}: שורות ${i} ו-${i + 1} אינן בסדר עולה ` +
            `(${JSON.stringify(words[i - 1])} ואז ${JSON.stringify(word)}). ` +
            'החיפוש הבינארי מניח מיון לפי יחידות UTF-16, וללא מיון הוא מפספס ערכים בשקט.',
        );
      }
    }

    return `window.${TORAH_DICTIONARY_GLOBAL} = \`${packed}\`;\n`;
  }

  return {
    name: 'otzaria-torah-dictionary',

    // בפיתוח אין `emitFile`, ולכן אותו תוכן בדיוק מוגש מהזיכרון — כדי
    // שהמסלול שנבדק ידנית יהיה המסלול שנארז.
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        if (!req.url || req.url.split('?')[0] !== `/${TORAH_DICTIONARY_FILE}`) return next();
        res.setHeader('Content-Type', 'text/javascript; charset=utf-8');
        res.end(build());
      });
    },

    // `generateBundle` רץ ב-build בלבד; ב-dev ה-middleware שמעל הוא המסלול.
    generateBundle() {
      this.emitFile({ type: 'asset', fileName: TORAH_DICTIONARY_FILE, source: build() });
    },
  };
}

/** מילון ראשי-תיבות, כנכס עצל: ב-file:// אין fetch, ולכן מוזרק script קלאסי. */
function acronymsAsset(): Plugin {
  return {
    name: 'otzaria-acronyms',
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        if (!req.url || req.url.split('?')[0] !== `/${ACRONYMS_FILE}`) return next();
        res.setHeader('Content-Type', 'text/javascript; charset=utf-8');
        res.end(buildAcronymsAsset());
      });
    },
    generateBundle() {
      this.emitFile({ type: 'asset', fileName: ACRONYMS_FILE, source: buildAcronymsAsset() });
    },
  };
}

/** ביטויים תלמודיים ושמות מחברים, באותה תבנית נכס עצל בדיוק. */
function staticCompletionAsset(): Plugin {
  return {
    name: 'otzaria-static-completion',
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        if (!req.url || req.url.split('?')[0] !== `/${STATIC_COMPLETION_FILE}`) return next();
        res.setHeader('Content-Type', 'text/javascript; charset=utf-8');
        res.end(buildStaticCompletionAsset());
      });
    },
    generateBundle() {
      this.emitFile({
        type: 'asset',
        fileName: STATIC_COMPLETION_FILE,
        source: buildStaticCompletionAsset(),
      });
    },
  };
}

const BLANK_DOCX_MODULE = 'virtual:otzaria-blank-docx';

/**
 * תבנית „מסמך חדש” עברית, כ-base64 במודול וירטואלי. נגזרת מהמסמך הריק של
 * המנוע בכל בנייה — ראו src/engine/blank-document.ts. נכס נפרד לא היה עובד:
 * `fetch` מ-file:// חסום (origin opaque), כמו במילון.
 */
function hebrewBlankDocx(): Plugin {
  const resolved = `\0${BLANK_DOCX_MODULE}`;
  return {
    name: 'otzaria-hebrew-blank-docx',
    resolveId(id) {
      return id === BLANK_DOCX_MODULE ? resolved : undefined;
    },
    load(id) {
      if (id !== resolved) return undefined;
      const docx = deriveHebrewBlankDocx({ patchDocument: patchBlankDocumentXml, patchStyles: patchBlankStylesXml });
      return `export default ${JSON.stringify(docx.toString('base64'))};\n`;
    },
  };
}

export default defineConfig({
  base: './',
  plugins: [
    superdocSelectionIdle(),
    vue(),
    hebrewBlankDocx(),
    torahDictionaryAsset(),
    acronymsAsset(),
    staticCompletionAsset(),
    inlineEngineWorkers(),
    classicChunks({ preloadDocumentEngine: true }),
    deferredEntry(),
  ],
  worker: { format: 'iife' },

  // ברירת המחדל של Vite ב-build היא legalComments: 'none', והיא מוחקת את באנר
  // הרישוי של מנוע ה-DOCX. סעיף 3.1(c) ברישיון המנוע אוסר להסיר הודעות רישוי,
  // ולכן ההודעות נאספות לסוף הקובץ. check-dist.mjs מאמת שהן שם.
  esbuild: { legalComments: 'eof' },

  // בשרת הפיתוח אין הטמעת workers, והמנוע בונה את ה-URL שלהם יחסית למודול
  // שלו. אם ה-dep optimizer של Vite אורז את המנוע מחדש ל-node_modules/.vite/deps,
  // ה-URL היחסי מצביע לשם — ושם אין קובץ worker, כלומר המסמך לא נפתח בפיתוח.
  // החרגה מה-optimizer משאירה את המנוע במקומו, ואת ה-URL נפתר.
  // Keep superdoc/ui in the transform pipeline for the selection-idle fix.
  optimizeDeps: { exclude: ['superdoc', '@superdoc/docx-engine'] },

  // PORT מכובד כשהוא מוגדר: כלי תצוגה (וכל סביבת עבודה עם כמה שרתי פיתוח
  // במקביל) מקצים פורט דרך משתנה הסביבה, ו-Vite מעצמו קורא רק --port.
  // בלי זה שרת שני נופל ל-5174 בעוד הכלי מצביע על הפורט שהקצה — דף ריק.
  // .vs/ is a Visual Studio index directory that locks files on Windows; exclude it from watching.
  server: {
    ...(process.env.PORT ? { port: Number(process.env.PORT), strictPort: true } : {}),
    watch: {
      ignored: ['**/.vs/**'],
    },
  },

  build: {
    target: 'es2020',
    assetsDir: 'assets',
    sourcemap: false,
    assetsInlineLimit: 4096,
    // The classic loader owns dependencies. For AMD, Vite injects each chunk's
    // CSS inside its factory, so lazy panel styles need no CORS/fetch/preloads.
    modulePreload: false,
    cssCodeSplit: true,
    chunkSizeWarningLimit: 12_000,
    rollupOptions: {
      output: {
        format: 'amd',
        amd: { autoId: true, define: 'window.__otzariaDefine' },
        exports: 'named',
        entryFileNames: 'assets/app.js',
        chunkFileNames: 'assets/[name]-[hash].js',
        assetFileNames: 'assets/[name][extname]',
      },
    },
  },
});
