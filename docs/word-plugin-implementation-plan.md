# תוכנית ביצוע מאומתת — „וורד לאוצריא” על SuperDoc v2

- תאריך בדיקה: 23.8.2026
- בסיס התוסף הישן: `main` בקומיט `b3b95d0`
- קומיט ה־spike שנבדק: `4c7a88a3634bfe01ea5d6f0ded0998af0ea46888`
- גרסת SuperDoc שנבדקה: `superdoc@2.8.0`

מסמך זה הוא תוכנית העבודה הסמכותית והיחידה. הוא **מחליף** את
`docs/superdoc-v2-plan.md` ואת `superdoc_word_ui_plan_he.md`, שנשארו בענף ה־spike
הנטוש `feat/superdoc-v2-editor` ואינם חלק מן הענף הזה; אין לבצע מהם משימות.

---

## 1. התוצאה המבוקשת

תוסף אופליין לאוצריא, בעברית וב־RTL, שמרגיש כמו עורך Word ושומר מסמכי DOCX
אמיתיים. SuperDoc v2 הוא מקור האמת היחיד למסמך; הממשק, ה־Ribbon וחיבורי אוצריא
הם קוד התוסף.

הגרסה הראשונה נחשבת מוכנה רק כאשר משתמש יכול:

1. לפתוח מסמך DOCX דרך בורר הקבצים של אוצריא, או ליצור מסמך ריק.
2. לערוך עברית, אנגלית וטקסט מעורב בעימוד עמודים.
3. להשתמש בלשונית „בית”, חיפוש (ללא החלפה — ראו §11), זום ושורת מצב.
4. לבצע „שמור” ו„שמור בשם” דרך API בטוח של אוצריא.
5. לסגור ולפתוח מחדש בלי לאבד מסמך שמור או לדרוס קובץ חלקי.
6. לייצא DOCX שנפתח שוב ב־Microsoft Word וב־SuperDoc.
7. לקבל טקסט מסומן מן הקורא דרך „שלח לוורד”.
8. לעבוד כתוסף ארוז ב־Windows WebView2, לא רק בשרת הפיתוח.

גרסת ההשקה היא דסקטופ תחילה. Windows הוא שער חובה; macOS ו־Linux עוברים
בדיקות עשן לפני פרסום. Android/iOS אינם מובטחים עד שיעברו שער נפרד של ביצועים,
זיכרון, בורר קבצים וממשק מותאם למסך קטן.

### אבני דרך מוצריות

| אבן דרך | תכולה | אופן הפצה |
|---|---|---|
| 2.0 experimental | שער Windows, מסמך יחיד, פתיחה/שמירה בטוחה, „בית”, חיפוש, זום/שורת מצב ו„שלח לוורד” | מזהה התוסף החדש `com.otzaria_word_editor.superdoc`, לצד 1.3.6 |
| 2.1 authoring | הוספה, פריסה, סקירה, הערות, הגהה מקומית, חיפוש באוצריא, והחלפה אם ה־capability gate שלה נפתח | עדיין לצד התוסף הישן; אוספים תאימות ו־round-trip |
| 2.2 replacement candidate | ריבוי מסמכים, drafts, ארכיון, תבניות והפיצ'רים המתקדמים שאושרו | מועמד להחלפת 1.3.6 רק אחרי מטריצת התאימות המלאה |

מספרי הגרסאות הם אבני דרך בתוסף, לא תחליף ל־`minAppVersion` של אוצריא.

---

## 2. החלטות סופיות

- משתמשים ב־`superdoc@2.8.0` בגרסה נעולה, ומעדכנים רק במסגרת משימת שדרוג נפרדת.
- מייבאים `superdoc` ו־`superdoc/ui` בלבד. אין לייבא את
  `@superdoc/docx-engine` ישירות ואין לשנות את המנוע או את ה־workers שלו.
- מריצים `new SuperDoc({ ui: false, telemetry: { enabled: false } })`.
- משתמשים ב־Vue 3 + TypeScript לממשק. אין להוסיף React.
- הבנייה הסופית היא סקריפטים קלאסיים בלבד, ללא `<script type="module">`, CDN או
  תלות ברשת.
- רישוי SuperDoc **סגור בחיוב** לפי
  [תשובת המפתחים ב־GitHub](https://github.com/superdoc/docx-editor/issues/3927#issuecomment-5383145303).
  הוא אינו שער עבודה.
  התוסף נשאר AGPL-3.0, שומר הודעות רישוי ומפרסם את המקור הזהה לבינארי.
- `src/` הישן נשאר תקין ולא משתנה עד להחלפה הסופית. כל העבודה החדשה בקוד החדש.
- אין `contentEditable`, `document.execCommand`, מודל HTML מקביל, Mammoth או
  יצואן OOXML ידני בקוד החדש.
- אין שמירת Blob/base64 ב־`storage`. ה־storage של אוצריא משמש רק להגדרות,
  metadata ו־tokens אטומים.

---

## 3. מה נמצא בביקורת ומה חייבים לתקן

### 3.1 קומיט `4c7a88a`

הקומיט הוא spike שימושי: ה־typecheck וה־build עוברים, מתקבל פלט קלאסי, וה־workers
נארזים מקומית. גודל ה־dist שנמדד הוא בערך 15.8MB לפני ZIP. עם זאת, הוא עדיין לא
הוכיח ריצה כתוסף ארוז ב־Windows.

לפני שמפתחים UI יש לתקן את הנקודות האלה:

| בקומיט | הבעיה | התיקון המחייב |
|---|---|---|
| `createSuperDocUI({ superdoc })` אחרי `new SuperDoc` | נוצר controller שני. לפי ה־API והדוגמה הרשמית, המופע כבר מחזיק controller ב־`superdoc.ui` | בתוך `onReady` מקבלים את מופע SuperDoc המוכן ושואלים ממנו את `readySuperDoc.ui`. אין לקרוא `createSuperDocUI` |
| `ui.destroy()` ואז `superdoc.destroy()` | ה־UI הוא משאב מושאל שבבעלות SuperDoc | בפירוק קוראים רק `superdoc.destroy()` ומבטלים subscriptions של הקוד שלנו |
| `<input type="file">` | מתאים ל־spike בדפדפן, לא למסלול production של אוצריא | להשתמש ב־`fs.pickUserFile`, לשמור token, ובטעינה חוזרת להשתמש ב־`fs.resolveFileUrl` |
| הורדה דרך `<a download>` | אינה נותנת „שמור” אמין לקובץ הקיים ואינה בסיס ל־autosave | לממש תחילה את הרחבת ה־SDK שבסעיף 7 |
| תוכנית autosave ל־`storage` | DOCX הוא Blob בינארי; `storage` הוא JSON KV. base64 יכביד על הגשר ועל הזיכרון | bytes עוברים ב־HTTP loopback אל ה־Host; ב־storage נשמרים רק token, שם, revision והגדרות |
| פס עליון בגובה 48px וצבע `surface-high` | אינו תואם את מדריך העיצוב הנוכחי | 56px, או 44px במצב compact; `surfaceContainerHigh` ומשתני ה־SDK |
| רשימת „50 פקודות” כממשק קבוע | `BUILT_IN_COMMAND_IDS` הציבורי אינו מבטיח את כל הקטלוג | registry קטן שלנו עם מזהים שבדקנו, `ui.commands.has`, בדיקת חוזה ואפשרות capability fallback |

### 3.2 התוסף הקיים ב־`main`

ממנו מעבירים רק את שפת המוצר והרעיונות:

- שמות הלשוניות והפקדים בעברית.
- מבנה „קובץ / בית / הוספה / פריסה / הפניות / תצוגה / סקירה / אוצריא”.
- חיבורי קורא, ספרייה וחיפוש של אוצריא, לאחר התאמה ל־API הנוכחי.
- `src/js/dictionary.js` כנתוני מילון מקומיים, לאחר המרה למודול ולאחר בדיקת
  זכויות/מקור הנתונים.
- רעיונות כגון שורת מצב, חלונית ניווט, כותרות ספר, TXT לאוצריא, תבניות ותגיות.

לא מעבירים קוד שמבצע עריכת DOM, חלוקת עמודים, import/export ידני, comments או
track changes עצמאיים. גם `search.fullText` הישן אינו מועתק; משתמשים ב־
`search.query` הזורם, או פותחים טאב חיפוש מובנה כאשר זה מתאים יותר ל־UX.

### 3.3 מקורות העיצוב שנבדקו

| מקור וקומיט שנבדק | מה מאמצים | מה לא מעתיקים |
|---|---|---|
| [LocalOffice `60bd8cef`](https://github.com/Anon5T4R/LocalOffice/commit/60bd8cef8f135a9cc9183f9a8217a0b8e8d84528) | פירוק Ribbon ללשוניות/קבוצות/כפתורים, mount של הלשונית הפעילה בלבד, dirty indicator, שורת מצב, שמירת format painter במעבר לשונית | פקודות TipTap, מודל המסמך וקוד React |
| [SuperDoc `examples/custom-ui` ב־`b0ff2221`](https://github.com/superdoc/docx-editor/commit/b0ff2221645f79b7094e1c037723fe2a435ffd3c) | `ui: false`, שימוש ב־`superdoc.ui`, observe למצב פקודה, `mousedown.preventDefault()`, `executeAsync`, teardown | אין לסטות אל API פנימי; מצמידים את הדוגמה לגרסת 2.8.0 ולא ל־main משתנה |
| [Herramienta_Optimizacion_PBM `437d79d2`](https://github.com/T0m4s1n/Herramienta_Optimizacion_PBM/commit/437d79d203db44af384861fe588ea5a0dd57724f) | בורר טבלה, overflow אופקי, aria, Escape/outside-click וקיצורי מקלדת | חיבור SuperDoc v1 ו־`headless-toolbar`; אין רישיון מזוהה ולכן אין העתקת קוד |
| [canvas-editor `03a481bb`](https://github.com/Hufe921/canvas-editor/commit/03a481bbd012f2dcb4044cd34471477db921fe52) | רעיונות לסרגל, עמודים, תפריטים הקשריים לטבלה/תמונה/קישור וחלוניות | import/export DOCX ומודל ה־Canvas אינם מקור אמת |
| [ONLYOFFICE/web-apps `9c0ca538`](https://github.com/ONLYOFFICE/web-apps/commit/9c0ca538c3b211052347df09d2a4d6781f023403) | compact/full Ribbon, לשוניות הקשריות, state מרכזי לנעילת פקדים, ניווט ושורת מצב | אין לחלץ את מסגרת ה־frontend הגדולה או controllers שלה |

אם מועתק קוד ממשי ממקור AGPL/MIT, יש להוסיף אותו ל־`THIRD_PARTY_NOTICES.md` עם
קישור, קומיט, רישיון ורשימת קבצים. השראה חזותית בלבד מתועדת בלי העתקה.

---

## 4. ארכיטקטורה מחייבת

```text
Vue UI בעברית
  AppShell / Ribbon / Panels / StatusBar / Dialogs
                │
                ├── CommandAdapter ── superdoc.ui (מושאל מהמופע)
                │                      commands, state, search, zoom,
                │                      styles, comments, trackChanges
                │
                ├── DocumentAdapter ─ activeEditor.doc (API ציבורי)
                │                      sections, tables, footnotes,
                │                      headers/footers, TOC, images
                │
                ├── SessionStore ──── metadata + dirty revision + file tokens
                │
                └── OtzariaClient ─── theme, fs, storage, reader,
                                       library, search, notifications
```

כללי מימוש:

1. פעולת Ribbon קיימת כפקודה? משתמשים ב־`superdoc.ui.commands`.
2. יכולת מבנית שאינה פקודת Ribbon? משתמשים רק ב־`activeEditor.doc` הציבורי.
3. אין selector אל DOM פנימי של SuperDoc ואין import מנתיב שאינו export ציבורי בחבילה.
4. כל mutation דרך Document API בודק receipt ומתרגם כשל להודעה בעברית.
5. כל subscription מחזיר disposer ונרשם ב־`DisposableBag` של session.
6. כפתור Ribbon מבטל `pointerdown/mousedown` כדי לא לאבד selection לפני הפקודה.
7. רק session פעיל מחזיק מופע SuperDoc. ריבוי מסמכים יתווסף לאחר MVP; במעבר
   מסמך מייצאים snapshot, מפרקים את המופע ומקימים מופע למסמך החדש.

### Registry התחלתי לפקודות ב־2.8.0

זו רשימת ההתחלה ל־`capabilities.ts`; ה־contract test של שלב 0 הוא הסמכות בזמן
הבנייה. אין לייבא את `COMMAND_CATALOG` מנתיב פנימי של החבילה.

| תחום | command IDs צפויים |
|---|---|
| היסטוריה | `undo`, `redo` |
| תו | `bold`, `italic`, `underline`, `strikethrough`, `clear-formatting`, `copy-format` |
| גופן | `font-family`, `font-size` |
| צבע | `text-color`, `highlight-color` |
| פסקה | `text-align`, `line-height`, `linked-style`, `direction-rtl`, `direction-ltr` |
| רשימות | `bullet-list`, `numbered-list`, `indent-increase`, `indent-decrease` |
| הוספה | `link`, `image`, `table-of-contents-insert`, `table-insert` |
| טבלה בהקשר | `table-add-row-before`, `table-add-row-after`, `table-delete-row`, `table-add-column-before`, `table-add-column-after`, `table-delete-column`, `table-delete`, `table-merge-cells`, `table-split-cell`, `table-remove-borders` |
| סקירה | `acceptChange`, `rejectChange`, `acceptAllChanges`, `rejectAllChanges`, `track-changes-accept-selection`, `track-changes-reject-selection` |
| תצוגה | `zoom`, `zoom-fit-width`, `ruler`, `formatting-marks`, `document-mode`, `measurement-unit` |

`BUILT_IN_COMMAND_IDS` הציבורי מכסה רק את קבוצת המזהים הקנוניים, לא את כל
המזהים בטבלה. לפני הצגת control קוראים `ui.commands.has(id)` ומאזינים ל־handle
שלו. פעולות table משתמשות גם ב־`ui.tables.getContext()`; בלי תא/טבלה חוקיים הן
disabled. `table-fix` אינו נתמך ואינו מוצג.

### מבנה קבצים יעד

```text
.
├── docs/
│   ├── spike.md
│   ├── spike-windows.md
│   ├── compatibility-matrix.md
│   └── fixtures/README.md
├── public/
│   ├── manifest.json
│   └── third-party/                 # notices בלבד, לא npm packages מלאים
├── src/
│   ├── main.ts
│   ├── App.vue
│   ├── engine/
│   │   ├── create-editor.ts
│   │   ├── command-adapter.ts
│   │   ├── document-adapter.ts
│   │   ├── capabilities.ts
│   │   ├── proofing-provider.ts
│   │   └── workers.ts
│   ├── host/
│   │   ├── otzaria-client.ts
│   │   ├── files.ts
│   │   ├── reader.ts
│   │   ├── search.ts
│   │   ├── settings.ts
│   │   └── theme.ts
│   ├── sessions/
│   │   ├── document-session.ts
│   │   ├── save-coordinator.ts
│   │   └── session-store.ts
│   ├── ui/
│   │   ├── shell/
│   │   ├── ribbon/common/
│   │   ├── ribbon/tabs/
│   │   ├── panels/
│   │   ├── dialogs/
│   │   └── status/
│   ├── composables/
│   │   ├── useCommand.ts
│   │   ├── useDisposable.ts
│   │   └── useOtzariaTheme.ts
│   ├── styles/
│   │   ├── tokens.css
│   │   ├── shell.css
│   │   └── rtl.css
│   └── types/otzaria_plugin.d.ts
├── tests/
│   ├── unit/
│   ├── integration/
│   └── contract/
├── THIRD_PARTY_NOTICES.md
├── package.json
└── vite.config.ts
```

---

## 5. שיטת עבודה למפתח ג׳וניור

כל שלב להלן הוא PR/קומיט עצמאי. לא מערבבים Host SDK, מנוע, UI ופיצ'רים באותו
קומיט. בכל משימה עובדים בסדר הזה:

1. קוראים את סעיף ה־API המצוין ואת ה־`.d.ts`; לא מנחשים חתימה.
2. כותבים test שמדגים את החוזה או את הכשל.
3. מממשים את השינוי הקטן ביותר.
4. מריצים typecheck, unit tests ו־build.
5. בודקים ידנית בעברית וב־RTL.
6. מעדכנים את מסמך תוצאות השלב ורק אז מסמנים checkbox.

אין להתקדם מעבר לשער שנכשל. כשל בשער נרשם עם הודעת השגיאה, גרסת אוצריא,
מערכת הפעלה וצעדי שחזור.

---

## 6. שלב 0 — ליישר ולסגור את ה־spike

מטרה: להוכיח שהמנוע וה־workers פועלים מתוסף ארוז לפני בניית Ribbon.

> **בוצע חלקית.** התוצאות המלאות: [`docs/spike.md`](spike.md).
> 0.1 ו־0.2 עברו, ו־0.4 נמדד ב־Chrome/macOS. **0.3 — שער A — לא עבר:** אין בו
> חוסם ידוע, אבל הוא לא הורץ על תוסף ארוז ב־Windows/WebView2.

### 0.1 תיקון חוזה SuperDoc

- [x] להסיר `createSuperDocUI` ואת הטיפוס הבעלים `SuperDocUI` מ־
  `src/engine/create-editor.ts`.
- [x] ב־`onReady`, להשתמש במופע המוכן וב־`readySuperDoc.ui`.
      החתימה בפועל: `onReady?: (params: SuperDocReadyPayload) => void`, כלומר
      `({ superdoc }) => superdoc.ui`.
- [x] לשנות את `EditorSession.ui` לטיפוס המושאל הציבורי המתאים —
      `BorrowedSuperDocUI`, שהוא `Omit<SuperDocUI, 'destroy'>`.
- [x] `destroy()` מבטל subscriptions של התוסף וקורא רק `superdoc.destroy()`.
- [x] להוסיף test שמוודא שלא נקראת פונקציית יצירת UI נוספת. נוספה גם בדיקת
      גבולות על המקור: אין import למנוע, אין נתיב פנימי, אין `execCommand`.

קבלה: `npm run typecheck`, unit tests ו־build עוברים; Bold ו־RTL ממשיכים לעבוד.

**נמצא ותוקן מעבר לרשום כאן:** ההרשמה ל־`plugin.boot` הייתה אחרי `await`.
האירוע נורה פעם אחת, אוצריא אינה שומרת את ה־payload ואין `getBootInfo` — כלומר
מסך שנשאר תלוי. ה־latch עבר לזמן טעינת המודול. בנוסף: `minAppVersion` היה
`0.9.97`, גבוה מגרסת אוצריא בפועל (0.9.96), מה שהיה מונע התקנה.

### 0.2 בדיקת חוזה פקודות

- [x] ליצור `tests/contract/superdoc-commands.test.ts`.
- [x] לרשום רק command IDs שהממשק באמת משתמש בהם — 47 מזהים ב־
      `src/engine/capabilities.ts`, מחולקים לפי תחום.
- [x] לכל מזהה לבדוק `ui.commands.has(id)`. **תיקון לתכנית:** `has` לבדו אינו
      מספיק — הוא מחזיר `true` גם ל־`table-fix`. ההבחנה היא `state.source`:
      `'unsupported'` מול `'builtin'`. וזה גם מייתר את ההמתנה ל־`onReady`:
      לפני שהמנוע מוכן `supported` הוא `false` לכל פקודה, אבל `source` נכון —
      ולכן הבדיקה רצה על controller שנבנה מעל host מבני ריק, בלי דפדפן ובלי
      מנוע DOCX.
- [x] פקודה חסרה גורמת לכשל test, לא לכפתור מת בשקט.
- [x] `table-fix` אינו נכנס לרשימה, והבדיקה מקבעת שהמנוע אכן מסמן אותו
      unsupported — כך שגרסה שתתמוך בו תתריע.

קבלה: ה־registry וה־UI נבנים מאותה רשימה; אין רשימה נוספת במסמך או בקומפוננטה.

הערה לשלב הבנייה: `BUILT_IN_COMMAND_IDS` הציבורי מכיל 16 מזהים (לא 14 כפי
שכתוב ב־JSDoc שלו), `COMMAND_CATALOG` אינו export ציבורי כלל, והגילוי הציבורי
היחיד הוא `ui.commands.ids` ו־`has`.

### 0.3 בדיקת Windows ארוזה — שער A

**מצב: השער לא עבר. אין חוסם ידוע; לא הורץ על Windows.** כל המדידות להלן הן
ב־Google Chrome headless על macOS, על `dist` מוגשת ישירות — לא בתוסף ארוז ולא
ב־WebView2. התיעוד: [`docs/spike-windows.md`](spike-windows.md).

מנוע ה־DOCX יוצר את ה־worker שלו כ־`new Worker(url, { type: 'module' })` בכל
מקומות הקריאה. השילובים נמדדו על האריזה האמיתית ב־Chromium:

| origin של הדף | צורת ה־worker | תוצאה |
|---|---|---|
| `file://` | blob:, **קלאסי** | **עובד**, `onReady` תוך 485ms |
| `http://127.0.0.1` | blob:, קלאסי | עובד, 470–477ms |
| `file://` | blob:, module | נכשל: `module-load-failed` |
| `file://` | data:, module | עובד עקרונית, חסום בגודל (~2MB) |
| `http://127.0.0.1` | ה־URL היחסי של המנוע | נכשל: `module-load-failed` |

שלוש מסקנות:

1. `workerUrls` הוא חובה ולא אופטימיזציה: ה־build הוא IIFE, ובו
   `import.meta.url` אינו מצביע לקובץ ה־JS, ולכן ה־URL שהמנוע בונה לבד אינו
   נפתר גם מ־origin תקין.
2. מ־`file://` (origin opaque) module worker נחסם, ו־data חסום בגודל.
3. worker **קלאסי** מ־blob עובד. Vite מפיק את ה־workers כ־IIFE, ובקוד המוטמע
   אין `import`/`export` ואין `import.meta` — כלומר הוא תואם־קלאסי. מה שמנע את
   הטעינה היה האופציה, לא הקוד. `src/engine/workers.ts` עוטף את בנאי ה־`Worker`
   ומסיר `type: 'module'` — רק ל־blob URLs שאנחנו בנינו. ה־build ו־`check:dist`
   נופלים אם הקוד המוטמע יהפוך ל־ESM, כי אז אין חלופה.

מה שנשאר לשער: הרצה ב־WebView2 על Windows, אריזה עם `otzaria pack-plugin`
ומסמכים אמיתיים. שער A עובר רק אם התוסף **הארוז** פועל; בדיקה מ־localhost לבדה
אינה מספיקה, והיא משמשת רק להפרדת כשל מנוע מכשל טעינה.

> **תיקון לגרסה קודמת של המסמך.** כאן נכתב שהשער חסום ושנדרש שינוי בצד אוצריא
> (הגשת תיקיית התוסף מ־`http://127.0.0.1`). זה היה שגוי — ההנחה הייתה שקוד
> ה־worker חייב להיטען כ־module. הדרישה בוטלה, ואיתה §7.0 שנוסף בעקבותיה.

### 0.4 גודל וביצועים — שער B

נמדד על macOS ב־Chromium (מלא: `docs/spike.md` §4):

| מה | ערך |
|---|---|
| `dist/assets/app.js` | 10.20MB (3.13MB gzip) |
| `dist/assets/engine-workers.js` | 4.90MB |
| סה"כ `dist/` | 15.16MB |
| ZIP של `dist/` | 4.31MB |
| boot עד `onReady`, מסמך ריק | 470–477ms |
| `superdoc.export()` על מסמך ריק | ~200ms, מחזיר Blob |

worker השיתופיות מושמט מהאריזה. שני הנשארים, כפי ש-`check:dist` מדפיס אותם:
המסמך 4.45MB ו-review-index 0.31MB.

אריזה עם ה־CLI האמיתי הורצה ועברה: `exit 0`, 8 קבצים, **4.32MB**, כולל
ולידציית העיצוב. (`dart run` על הסקריפט קורס בקומפילציה של חבילת אוצריא במכונה
שנבדקה — באג SDK שאינו קשור לתוסף — ולכן ה־CLI הורץ דרך `flutter test`.)

עוד לא נמדד, ודורש Windows: זמן פתיחת 50 עמודים, peak memory, ומגבלת החנות
כמספר.

קריטריון ראשוני: האריזה מתקבלת על ידי validator והחנות; boot קר אינו נתקע; אין
קריסה במסמך 50 עמודים. אם קיימת מגבלת חנות, היא נכתבת כמספר מדויק ולא כ־TODO.

---

## 7. שלב Host — כתיבת DOCX בטוחה ב־SDK של אוצריא

זהו תנאי מוקדם ל„שמור”, autosave וארכיון. ה־SDK הנוכחי יודע לתת URL לקריאה אך
אין בו כתיבת bytes, ואין הרשאה `fs.user_files.write` — היא אינה קיימת בקוד
אוצריא. אסור להעביר DOCX כ־base64 ב־JSON-RPC.

השינוי נעשה במאגר אוצריא, ב־PR נפרד, תוך תאימות לאחור. הוא חוסם **שמירה**, לא
פיתוח ולא הפצה של גרסה לקריאה ולייצוא.

> גרסה קודמת של המסמך הכילה כאן §7.0 — הגשת תיקיית התוסף מ־origin loopback —
> כתנאי מוקדם לשער A. הוא נמחק: האריזה עובדת מ־`file://` כמו שהיא (§0.3), ואין
> צורך לגעת בשרת ה־loopback, ב־mime types או במסלול הטעינה של ה־WebView.

### 7.1 חוזה API מוצע

להוסיף הרשאה רגישה חדשה: `fs.user_files.write`.

להרחיב באופן תואם־לאחור גם את הפתיחה:

```ts
Otzaria.call('fs.pickUserFile', {
  extensions: ['docx'],
  access: 'readwrite' // ברירת המחדל נשארת 'read'
})
// => { cancelled, token, url, name, size, access }
```

בקשת `readwrite` דורשת גם `fs.user_files.read` וגם `fs.user_files.write`. מאחר
שהמשתמש בחר את הקובץ בדיאלוג והעניק לתוסף הרשאת כתיבה בהתקנה, ה־token יכול לשמש
ל־Save לאותו קובץ. קריאה ישנה בלי `access` ממשיכה לקבל token לקריאה בלבד.

```ts
// שלב 1: פתיחת upload זמני. targetToken קיים רק ב-Save חוזר.
Otzaria.call('fs.beginBinaryWrite', {
  purpose: 'user-file' | 'plugin-file',
  targetToken?: string,
  expectedSize?: number
})
// => { writeToken, uploadUrl, expiresAt, maxBytes }

// שלב 2: הבייטים אינם עוברים בגשר.
fetch(uploadUrl, {
  method: 'PUT',
  headers: { 'Content-Type': DOCX_MIME },
  body: blob
})

// שלב 3א: שמירה לקובץ משתמש.
Otzaria.call('fs.commitUserFileWrite', {
  writeToken,
  suggestedName,
  extension: 'docx',
  title?: string
})
// => { cancelled, token?, name?, size? }

// שלב 3ב, מאוחר יותר: snapshot פרטי לשחזור מסמך שלא נשמר.
Otzaria.call('fs.commitPluginFileWrite', {
  writeToken,
  name
})
// => { token, name, size }

Otzaria.call('fs.resolvePluginFileUrl', { token })
// => { token, url, name, size }

Otzaria.call('fs.revokePluginFile', { token })
// => true
```

משמעות `targetToken`:

- token שנוצר מ־`pickUserFile` ישן הוא read-only. השמירה הראשונה שלו פותחת
  „שמור בשם” ומחזירה token חדש עם `readwrite`.
- token שנוצר מ־`pickUserFile({ access: 'readwrite' })` הוא יעד חוקי ל־Save חוזר.
- token שחזר מ־`commitUserFileWrite` הוא readwrite. „שמור” הבא כותב לאותו יעד
  בלי דיאלוג.
- token פרטי פונה רק לאחסון התוסף ואינו יכול להצביע לקובץ משתמש.
- `purpose` נאכף כבר ב־begin: `user-file` דורש `fs.user_files.write`, ואילו
  `plugin-file` דורש `plugin.storage.write`. resolve/revoke של private draft
  דורשים בהתאמה `plugin.storage.read`/`plugin.storage.write`.

### 7.2 כללי אבטחה ושרידות

- [ ] שרת ה־loopback מאזין רק ל־`127.0.0.1` ובפורט אקראי.
- [ ] `writeToken` הוא 256-bit אקראי, משויך ל־pluginId, חד־פעמי ופג תוך 2 דקות.
- [ ] upload מקבל `PUT` יחיד בלבד, עם `Content-Length` ומגבלת 100MB בשלב הראשון.
- [ ] לכל תוסף מותרות לכל היותר שתי העלאות פעילות.
- [ ] לא מתקבל path מן ה־JavaScript בשום שלב.
- [ ] קובץ עולה ל־temp; commit מבצע flush ואז atomic replace/rename ככל שהפלטפורמה
  מאפשרת. בכשל, הקובץ המקורי נשאר שלם.
- [ ] ביטול Save As מוחק את ה־temp ואינו משנה grant.
- [ ] קבצים זמניים שפג תוקפם נמחקים בעלייה וב־dispose.
- [ ] כותרות CORS מתירות רק את המתודות/headers הדרושים; GET/HEAD הישנים ממשיכים.
- [ ] grant הנשמר ב־KV משתנה מ־`token -> path` ל־
  `token -> { path, access }`; migration קוראת string ישן כ־`access: read`.
- [ ] `revokeFile` מבטל גם grant כתיבה ו־sessions הקשורים אליו.
- [ ] אחסון פרטי מקבל quota נפרד (התחלה: 250MB לתוסף), LRU וניקוי בהסרה.

### 7.3 קבצים לעדכון במאגר אוצריא

- `lib/plugins/services/plugin_file_server.dart` — PUT sessions, מגבלות ו־cleanup,
  או service חדש שמופרד מן ה־GET server.
- `lib/plugins/bridge/plugin_bridge_adapter.dart` — פעולות ה־RPC החדשות ו־grant migration.
- `lib/plugins/bridge/plugin_bridge_handler.dart` — הרשאות, timeout וניתוב.
- `lib/plugins/models/plugin_valid_permissions.dart` — permission mapping.
- `lib/plugins/models/plugin_permission_labels.dart` — הסבר ברור למשתמש.
- `lib/plugins/services/plugin_extended_validator.dart` — APIs מוכרים וגרסת מינימום.
- `docs/plugin-sdk/README.md`, `API_REFERENCE.md`, `COOKBOOK.md` ו־
  `otzaria_plugin.d.ts` — תיעוד ודוגמה מלאה.
- tests של file server, adapter, handler, permissions ו־validator.

### 7.4 בדיקות Host שחייבות לעבור

- [ ] `pickUserFile` ללא `access` נשאר read-only; בקשת `readwrite` ללא הרשאת
  הכתיבה נכשלת; עם שתי ההרשאות היא מחזירה grant מתאים.
- [ ] upload תקין, גדול מדי, ללא length, token שגוי, token פג ושימוש שני.
- [ ] plugin אחר אינו יכול להשתמש ב־writeToken או file token.
- [ ] Save As: הצלחה וביטול.
- [ ] overwrite אטומי: כשל באמצע אינו פוגע בקובץ המקורי.
- [ ] migration של grant ישן.
- [ ] revoke בזמן upload.
- [ ] ניקוי temp אחרי crash מדומה.
- [ ] Windows, macOS ו־Linux; mobile לפני הכרזת תמיכה mobile.
- [ ] `flutter analyze` וכל בדיקות plugin SDK עוברות במאגר אוצריא.

לאחר המיזוג, מעדכנים את `src/types/otzaria_plugin.d.ts` מהמקור ואת
`minAppVersion` לגרסת אוצריא הראשונה שמכילה את ה־API. אין לנחש מספר גרסה.

---

## 8. שלב 1 — מעטפת אפליקציה ו־Host client

### 8.1 Boot ו־lifecycle

- [ ] להרכיב shell מינימלי לפני כל קריאת SDK, אך לא לקרוא ל־SDK לפני
  `plugin.boot`.
- [ ] `OtzariaClient` מספק `call`, `on`, `off` עם טיפוסים ושגיאות מנורמלות.
- [ ] `plugin.suspended`: לעצור timers, להמתין לשמירה פעילה רק במסגרת timeout קצר,
  ולסמן pending save. אין להתחיל export כבד בזמן suspension.
- [ ] `plugin.resumed`: לרענן URL מ־token אם צריך, לחדש subscriptions ולטפל
  ב־pending save.
- [ ] כל listener מתבטל ב־unmount.

### 8.2 Theme ועיצוב בסיסי

- [ ] פס עליון קבוע בגובה 56px; 44px ב־compact.
- [ ] צבעים רק ממשתני אוצריא כגון surface, surfaceContainerHigh, primary,
  onSurface ו־outline.
- [ ] `theme.changed` מעדכן בלי reload.
- [ ] `dir="rtl"`, logical CSS properties ו־focus rings גלויים.
- [ ] אין hex בצד ה־UI פרט ל־fallbacks מרוכזים ב־`tokens.css`.

### 8.3 Manifest מינימלי

ב־MVP מבקשים רק:

- `fs.user_files.read`
- `fs.user_files.write`
- `reader.context_menu`
- `reader.open` — זו ההרשאה הנוכחית של `reader.getSelection`
- הרשאות library/search רק בשלב אוצריא, לא לפני כן

`storage`, theme ו־feedback בסיסי אינם מוכפלים אם הם baseline בגרסת ה־SDK.

קבלה: light/dark מתעדכנים, suspension/resume אינם משאירים listeners, וה־manifest
עובר validator בלי הרשאות שאינן בשימוש.

---

## 9. שלב 2 — מסמך יחיד, פתיחה ושמירה ללא אובדן נתונים

### 9.1 מודל session

```ts
interface DocumentSession {
  id: string;
  title: string;
  sourceToken: string | null;
  writableToken: string | null;
  dirtyRevision: number;
  savedRevision: number;
  saveState: 'idle' | 'exporting' | 'uploading' | 'committing' | 'error';
}
```

אין bytes בתוך האובייקט המתמיד. בזמן ריצה בלבד מותר לשמור Blob זמני.

> מומש ב־`src/sessions/save-coordinator.ts` ובמעטפת. מה שנשאר פתוח מסומן
> להלן, ובדיקות ידניות ב־`docs/spike-windows.md`.

### 9.2 פתיחה

- [x] „פתח” קורא
  `fs.pickUserFile({ extensions: ['docx'], access: 'readwrite' })`. בלי הרשאת
  כתיבה נופלים ל־`read` — מסמך שנפתח ואינו נשמר עדיף על מסמך שלא נפתח.
- [x] cancellation אינו שגיאה ואינו מפרק את המסמך הנוכחי.
- [x] URL מוחזר נמסר ישירות ל־`Config.document`.
- [x] token/name נשמרים רק לאחר `onReady` מוצלח (ההחלפה עצמה אטומית —
  `sessions/editor-swap.ts`).
- [x] בעלייה חוזרת קוראים `fs.resolveFileUrl` עם ה־token ששמור ב־`storage`.
  קובץ שהוזז/נמחק, או שנפתר ולא נפתח (פגום), נשכח ונפתח מסמך ריק עם הודעה —
  לא לולאת שגיאה ולא תוסף בלי מסמך.
- [x] פתיחה חדשה כאשר יש dirty document דורשת Save / Discard / Cancel.
  `ui.showConfirm` דו־כפתורי, ולכן שלושת המצבים נבנים משתי שאלות.

### 9.3 SaveCoordinator

- [x] `onEditorUpdate` מעלה `dirtyRevision`.
- [x] Ctrl/Cmd+S קורא `saveNow()`; Ctrl/Cmd+Shift+S הוא „שמור בשם”, ושניהם
  מתעלמים בזמן שמירה (אחרת הם מצטרפים לסבב שרץ ומאבדים את המשמעות).
- [x] אין שתי שמירות במקביל. שינוי בזמן שמירה מריץ סבב נוסף.
- [x] כל סבב מצלם revision, מייצא, מעלה דרך loopback ועושה commit.
- [x] רק אחרי commit מוצלח: `savedRevision = exportedRevision`.
- [x] אם writable token חסר, commit פותח Save As. ביטול משאיר dirty.
- [x] שגיאת export/upload/commit משאירה dirty; אין הודעת „נשמר”.
- [x] סבב שנעצר אחרי ההעלאה משחרר אותה מיד דרך `fs.abortBinaryWrite`, ואינו
  משאיר קובץ זמני וסלוט במכסה תפוסים עד שה־token פג.
- [x] autosave מתחיל רק אחרי שלמסמך יש writable token, עם debounce של 2.5 שניות.
- [x] ב־before close/switch משתמשים ב־Save / Discard / Cancel; אין הסתמכות על
  `beforeunload`.
- [x] **סבב שייך למסמך שפתח אותו.** `reset` מעלה epoch, וסבב שמסתיים אחרי מעבר
  מסמך מוחזר כ־`stale`: אינו מאמץ יעד, אינו מזיז `savedRevision` ואינו מפרסם
  מצב. בלי זה שמירה של א' שהסתיימה אחרי פתיחת ב' הייתה מפנה את autosave של ב'
  לקובץ של א'.

למסמך חדש שלא נבחר לו יעד אין autosave לדיסק ב־MVP. לאחר מימוש
`commitPluginFileWrite`, שומרים snapshot פרטי משוחזר ומוחקים אותו רק אחרי Save
מוצלח או Discard מפורש.

### 9.4 קבלה

- פתיחה, עריכה, Save As, עריכה נוספת ו־Save חוזר עובדים.
- ביטול Save As אינו מאבד תוכן.
- כשל upload מדומה אינו דורס את הקובץ הישן.
- שינוי בזמן save יוצר save שני ולא מסומן בטעות כשמור.
- restart פותח token שמור; מסמך private-draft משוחזר כאשר H2 קיים.

---

## 10. שלב 3 — תשתית UI משותפת ו־Ribbon „בית”

### 10.1 רכיבים משותפים

ליצור לפני הלשוניות:

- `Ribbon.vue` — tabs, groups ו־overflow; רק הלשונית הפעילה mounted.
- `RibbonButton.vue` — icon, label, tooltip, active, disabled, aria-pressed.
- `RibbonSplitButton.vue`, `RibbonSelect.vue`, `ColorPicker.vue`.
- `useCommand(id)` — subscribe בעת mount, unsubscribe בעת unmount, `run(payload)`.
- `useCapability(name)` — מסתיר או משבית יכולת שאינה קיימת בגרסה הנעולה.

בכל control:

- `pointerdown.preventDefault()` לפני click כדי לשמור selection.
- Tab navigation, Enter/Space, Escape וסגירת popover בלחיצה חיצונית.
- disabled מגיע ממצב SuperDoc, לא מחישוב DOM.
- אין swallow לשגיאה; reason/receipt מתורגמים פעם אחת ב־adapter.

### 10.2 לשונית בית — סדר מימוש

כל שורה היא קומיט קטן עם test רכיב ובדיקה ידנית:

1. Undo / Redo.
2. Bold / Italic / Underline / Strikethrough / Clear formatting.
3. Font family / size; רשימת הגופנים מגיעה מ־`superdoc.ui.fonts` ונשמרת memoized.
4. Text color / highlight.
5. RTL / LTR וכיווניות פסקה.
6. יישור, line height והזחות.
7. bullet/numbered lists.
8. linked styles / style picker.
9. format painter, רק אם command contract test עובר; ה־state נשמר ברמת Ribbon.

קבלה: בהזזת הסמן בין שני קטעים בעלי עיצוב שונה, active/value/disabled מתעדכנים
ללא polling. הפעלת control אינה מאבדת את הבחירה.

---

## 11. שלב 4 — חיפוש, תצוגה, שורת מצב ונגישות

- [ ] חיפוש דרך `superdoc.ui.search`, לא `window.find` ולא סריקת DOM.
- [ ] **החלפה היא capability gate, לא תכולה של 2.0.** מה שנמדד ב־2.8.0 על מסמך
  חי: `search.available` הוא `true` ו־`canReplace` הוא `true`, אבל `replace`
  ו־`replaceAll` החזירו `{ ok: false, reason: 'operation-unavailable' }`.
  המדידה נעשתה על מסמך ריק ובלי התאמה פעילה, ולכן היא **אינה** מפרידה בין
  „החלפה לא מומשה” ל„אין מה להחליף”. במקביל, אוצר ה־reasons של החבילה מתעד
  `replace-unsupported` עם ההסבר ש־replace ו־replace-all נכשלים סגור „until
  replace ships”. בשתי הקריאות אין להבטיח החלפה.
  לפני מימוש יש להריץ: מסמך עם טקסט, `search(query)` שמחזיר `total > 0`,
  ואז `replace`. אם התוצאה עדיין כושלת — הפקד אינו מוצג, או מוצג disabled עם
  הודעה „אינו זמין בגרסה זו”, ולא מסתיר את הכשל.
- [ ] בכל מקרה: אם `canReplace` הוא `false` — אין להציג את הפקד.
- [ ] Zoom, fit width וערך zoom דרך `superdoc.ui.zoom`/פקודות ציבוריות.
- [ ] מספר עמודים מ־`onPaginationUpdate`; עמוד פעיל רק אם קיים מקור ציבורי אמין.
- [ ] ספירת מילים/תווים דרך Document API (`getText` או info), עם debounce ולא בכל
  keystroke. לא לספור HTML.
- [x] ruler ו־formatting marks דרך controllers/commands ציבוריים. הפקודה `ruler`
  מנותבת ל־`SuperDoc.toggleRuler()` והמצב שלה הוא `config.rulers`, אבל הסרגל
  המובנה **מושתק** ב־`ui: false` — ולכן הציור הוא שלנו
  (`src/ui/shell/DocumentRuler.vue`, `VerticalRuler.vue`), מעל מדידה אחת של
  העמוד המצויר. ראו „הסרגל” ב־[engine-gaps.md](engine-gaps.md).
- [ ] Ctrl/Cmd+F, Ctrl/Cmd+H, Ctrl/Cmd+S, Ctrl/Cmd+Z/Y; לא לדרוס קיצור כשהפוקוס
  בתוך input והפעולה אינה עריכת מסמך.
- [ ] Focus mode ו־compact Ribbon הם state של ה־shell.
- [ ] בדיקת keyboard-only, screen-reader labels, contrast ו־200% zoom פנימי.

קבלה: אין focus trap; Escape סוגר surface עליון; כל control מקבל שם נגיש; חיפוש
עובד בעברית, כולל טקסט עם ניקוד. החלפה אינה תנאי קבלה — ראו ה־capability gate
שלמעלה.

---

## 12. שלב 5 — „הוספה” ו„פריסה” דרך Document API

אין לבנות את כל הלשונית מראש. לכל feature מוסיפים adapter קטן, test fixture
ו־round-trip test.

### הוספה

1. קישור.
2. תמונה מקובץ מקומי, כולל alt text אם ה־API מאפשר.
3. טבלה עם בורר 1×1 עד 10×10 וניהול שורות/עמודות.
4. page break.
5. header/footer ומספר עמוד.
6. footnote/endnote.
7. TOC.
8. סימנים מיוחדים ותאריך — כהכנסת טקסט דרך Document API, לא DOM.

### פריסה

1. margins presets + custom.
2. portrait/landscape.
3. paper size.
4. columns.
5. section direction ו־section breaks.
6. page borders רק אם ה־API הציבורי וה־round-trip תומכים בו ב־2.8.0.

לכל פעולה:

- לבדוק capability בעת boot.
- להשתמש ב־target/selection ציבורי.
- לבדוק `receipt.success` ולתעד failure code.
- לייצא, לפתוח מחדש ולהשוות את החלק הרלוונטי במסמך.

פקד שאין לו API ציבורי אמין מסומן „לא זמין בגרסה זו”; לא מממשים אותו דרך XML
ידני או DOM פנימי.

---

## 13. שלב 6 — סקירה, הערות והגהה

### 13.1 Track Changes ותגובות

- [ ] document mode: editing/suggesting/viewing דרך `superdoc.ui.document`.
- [ ] accept/reject selection/all דרך commands שנבדקו.
- [ ] רשימת שינויים דרך `superdoc.ui.trackChanges`.
- [ ] יצירה/עריכה/מחיקה של תגובות דרך `superdoc.ui.comments` או Document API
  ציבורי, בלי מערכת תגובות מקבילה.
- [ ] author identity קבוע מתוך הגדרת משתמש מקומית; לא לשלוח מידע לרשת.

### 13.2 מילון תורני מקומי

להמיר את `src/js/dictionary.js` מקובץ שמציב global למודול נתונים נטען מקומית.
אין לטעון את כל המילון שוב בכל בדיקה.

לממש provider לפי `Config['proofing']`:

- `id: 'otzaria-torah-dictionary'`.
- capabilities: spelling, suggestions לפי מה שנתמך באמת, ללא רשת.
- `check(request)` עובר על `request.segments`, מחזיר offsets ב־UTF-16, מכבד
  `request.signal` ו־`maxSuggestions`.
- נרמול Unicode נבדק במיוחד לניקוד, טעמים, גרש/גרשיים ומקף עברי.
- המרה לאותיות בלי ניקוד אינה משנה offsets; אם בודקים ייצוג מנורמל, שומרים
  מפת אינדקסים אל הטקסט המקורי.
- batching, debounce ו־timeout מוגדרים ב־SuperDoc, לא ב־timer נפרד לכל מילה.

קבלה: אין false positive על מילים במילון; issue מצביע בדיוק למילה המקורית גם
בניקוד; abort מפסיק עבודה; מסמך גדול אינו קופא.

---

## 14. שלב 7 — חיבורי אוצריא

### 14.1 „שלח לוורד” מהקורא

להצהיר ב־manifest, לא לרשום בכל boot:

```json
{
  "contributes": {
    "startup": {
      "contextMenuItems": [
        {
          "id": "send-to-word",
          "title": "שלח לוורד",
          "contexts": ["reader-selection"],
          "openPlugin": true,
          "param": { "action": "insert-selection" }
        }
      ]
    }
  }
}
```

- [ ] להצהיר `app.startup_contributions` ו־`reader.context_menu`; אין צורך
  ב־keepAlive או בתהליך רקע קבוע.
- [ ] מיד אחרי `plugin.boot`, ולפני יצירת העורך או כל await ארוך, להירשם ל־
  `contextMenu.itemClicked`. ה־Host ממתין לסיום ה־boot; אין להוסיף השהיה משלנו.
- [ ] לקבל את `selection`/`selectedText` מן event, לא לבצע קריאת DOM לקורא.
- [ ] אם אין מסמך פתוח, ליצור חדש; אם יש dirty session פעיל, להוסיף במיקום הסמן.
- [ ] להציע שלושה מצבים: טקסט בלבד, טקסט + מקור, וציטוט מעוצב.
- [ ] הכנסת הטקסט נעשית דרך Document API ומכבדת RTL.

### 14.2 חיפוש ופתיחת מקורות

- [ ] עטיפה ל־`search.query` כ־AsyncIterable עם cancel ו־limit.
- [ ] להציג תוצאות בדפים; לא לשמור את כל corpus בזיכרון.
- [ ] פתיחת מקור באוצריא דרך פעולת reader הציבורית המתאימה לגרסת ה־SDK.
- [ ] הכנסת ציטוט למסמך שומרת metadata מינימלי באובייקט הקישור/הערה רק אם
  SuperDoc round-trips אותו; אחרת המקור נכנס כטקסט גלוי.
- [ ] אין `search.fullText` חדש ואין קריאות מעל 100 לשנייה.

קבלה: סימון בקורא → „שלח לוורד” → התוסף נפתח והטקסט נכנס פעם אחת בלבד; חיפוש
ניתן לביטול; פתיחת תוצאה חוזרת לקורא במיקום הנכון.

---

## 15. שלב 8 — ריבוי מסמכים, drafts, ארכיון ותבניות

שלב זה מתחיל רק לאחר שמסמך יחיד ושמירה יציבים.

- [ ] `SessionManager` מחזיק metadata לכל tab, אך מופע SuperDoc אחד פעיל בלבד.
- [ ] לפני מעבר tab: export snapshot, commit כ־plugin-private draft, שמירת selection
  אם יש API ציבורי, ואז destroy.
- [ ] לאחר מעבר: resolve draft URL, יצירת SuperDoc חדש, שחזור zoom/selection כאשר אפשר.
- [ ] dirty dot ו־save state לכל tab.
- [ ] סגירת tab: Save / Discard / Cancel.
- [ ] ארכיון הוא רשימת private tokens + metadata; מחיקה מבטלת token ומפנה quota.
- [ ] תבנית היא DOCX private read-only שממנה נוצרת טיוטה חדשה.
- [ ] macro הוא רצף פעולות ברמת command/document adapters עם schema וגרסה; אין
  הקלטת clicks, selectors או פעולות DOM.
- [ ] להגביל מספר tabs פתוחים ולמדוד memory בכל מעבר של 20 פעמים.

קבלה: מעבר בין חמישה מסמכים אינו מאבד שינוי, אינו משאיר workers ישנים ואינו
מגדיל זיכרון ללא גבול.

---

## 16. שלב 9 — פיצ'רים מתקדמים מן התוסף הישן

להכניס אחד־אחד, ורק כאשר קיים API ציבורי ו־round-trip fixture:

- חלונית ניווט לפי headings.
- bibliography/citations/cross-references.
- הגנת מסמך/content controls.
- כותרות ספר ותבניות תורניות.
- TXT בפורמט אוצריא, הנגזר מ־Document API ולא מ־HTML.
- תגיות, פתקים וקליפים — metadata צדדי; לא להכניס XML פרטי למסמך בלי חוזה.
- הדפסה ו־PDF רק אחרי יציבות DOCX; כל worker/asset חייב להיות מקומי.

פיצ'ר שאינו שורד export→open מסומן experimental ואינו מופעל כברירת מחדל.

---

## 17. בדיקות תאימות ו־round-trip

### 17.1 Fixtures

ליצור 10–15 מסמכים סינתטיים או מורשים ב־`docs/fixtures/`; אין להכניס מסמכים
פרטיים או ספרים מוגנים. כל fixture מתעד מה הוא בודק ומה צפוי להישמר.

הכיסוי המינימלי:

- עברית RTL; עברית+אנגלית; ניקוד וטעמים.
- fonts/styles ורשימות רב־רמתיות.
- טבלאות מורכבות ותמונות.
- headers/footers ומספרי עמודים.
- footnotes/endnotes.
- page/section breaks, margins, orientation ו־columns.
- TOC ושדות Word.
- comments ו־track changes.
- מסמך 50–100 עמודים ומסמך קרוב ל־100MB limit.

### 17.2 תרחיש קבוע לכל fixture

1. פתח דרך `fs.pickUserFile`.
2. המתן ל־`onReady` ולסיום pagination.
3. בצע שינוי קטן ומזוהה.
4. שמור דרך Host API.
5. פתח את הפלט מחדש ב־SuperDoc.
6. פתח ב־Microsoft Word ב־Windows.
7. בדוק חזותית את העמודים הרלוונטיים.
8. השווה OOXML רק ככלי בדיקה אוטומטי; הקוד המוצרי לעולם אינו עורך OOXML ידנית.

### 17.3 מטריצת פלטפורמות

| סביבה | לפני merge | לפני release |
|---|---:|---:|
| Chrome/Vite dev | unit + integration | כן |
| Windows WebView2, תוסף ארוז | smoke של השלב | כל fixtures |
| Microsoft Word Windows | fixtures שהשתנו | כל fixtures |
| macOS אוצריא | smoke | fixtures מרכזיים |
| Linux אוצריא | smoke | fixtures מרכזיים |
| Android/iOS | לא חוסם desktop | שער נפרד לפני הכרזת תמיכה |

---

## 18. CI, אריזה והפצה

### בכל PR

- `npm ci`
- typecheck
- unit tests
- component/integration tests
- contract tests מול `superdoc@2.8.0`
- build production
- בדיקה שאין `<script type="module">`
- בדיקה שאין URL חיצוני או asset חסר ב־dist
- `otzaria pack-plugin` validator
- `git diff --check`

### לפני release

- [ ] כל שערי Windows ו־round-trip עברו.
- [ ] `THIRD_PARTY_NOTICES.md`, LICENSE והודעות SuperDoc בתוך החבילה.
- [ ] source tag זהה לבינארי המופץ.
- [ ] `minAppVersion` תואם ל־SDK write API.
- [ ] permissions מינימליות ומוסברות למשתמש.
- [ ] אין source maps עם נתיבים מקומיים או fixtures פרטיים.
- [ ] release workflow בונה מהשורש ומצרף checksum.
- [ ] רק עכשיו מחליפים את נתיב הבנייה הישן; `src/` נשמר tag אחד לפחות ל־rollback.

---

## 19. סדר ה־PRs המומלץ

1. ~~`spike: use borrowed superdoc.ui and add contract tests`~~ — בוצע.
2. ~~`spike: load engine workers as classic workers`~~ — בוצע; החוסם הוסר,
   השער עצמו לא הורץ.
3. `spike: run gate A on packaged Windows and record limits` — נותר להריץ.
4. במאגר אוצריא: `sdk: add streamed atomic user-file writes`
5. במאגר אוצריא: `sdk: add quota-bound private binary drafts`
6. `v2: typed host client, lifecycle and Otzaria theme shell`
7. `v2: single-document sessions and safe save coordinator`
8. `v2: shared ribbon controls and Home tab`
9. `v2: search, zoom, status and accessibility`
10. `v2: Insert features, one capability per commit`
11. `v2: Layout features, one capability per commit`
12. `v2: Review and local Torah proofing`
13. `v2: declarative reader integration and library search`
14. `v2: multi-document drafts, archive and templates`
15. `v2: compatibility suite, packaging and release cutover`

אין להתחיל PR 7 לפני ש־PR 4 נמצא לפחות בגרסת אוצריא מקומית ניתנת לבדיקה. PR 5
חוסם drafts וריבוי מסמכים, אך אינו חוסם MVP של מסמך שמור יחיד. PR 3 אינו חוסם
פיתוח — שלבים 1–8 נבנים ונבדקים גם משרת הפיתוח — אבל הוא חוסם שחרור.

---

## 20. Definition of Done לגרסה 2.0 experimental

- [ ] שער workers בתוסף Windows ארוז עבר.
- [ ] כתיבה דרך SDK עברה בדיקות כשל ואבטחה, ומובטח שכשל אינו הורס את הקובץ
  הקיים (אטומיות ההחלפה תלויה במערכת הקבצים — ראו §7).
- [ ] מסמך יחיד: new/open/save/save-as/reopen ללא אובדן נתונים.
- [ ] Home, חיפוש (ללא החלפה), zoom/status ו־RTL עובדים דרך API ציבורי בלבד.
- [ ] „שלח לוורד” עובד דקלרטיבית מהקורא.
- [ ] אין רשת, telemetry, CDN, module script או import ישיר למנוע.
- [ ] אין שימוש ב־DOM הפנימי של SuperDoc או במנוע הישן.
- [ ] בדיקות keyboard, theme, suspension ו־memory עברו.
- [ ] האריזה כוללת רישיונות ומקורה פורסם תחת AGPL-3.0.
- [ ] קיים rollback מתועד לגרסה 1.3.6.

כאשר כל הסעיפים האלה מסומנים אפשר לפרסם את 2.0 כתוסף experimental נפרד; עדיין
לא מכריזים עליו כמחליף של התוסף הישן. ההחלפה מותרת רק ב־2.2, לאחר שגם שלבים
12–16, private drafts, ריבוי מסמכים וכל מטריצת התאימות של סעיף 17 עברו, ותועד
rollback לגרסה 1.3.6.
