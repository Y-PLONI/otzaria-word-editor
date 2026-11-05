# מדידת פתיחה והקלדה

`bench:performance` מפיק JSON ומכשיל ריצה שלא הוכיחה שהקלט הגיע למסמך המיועד.
הכלי הוא כלי אבחון; בשימוש רגיל אין recorder, מאזיני הקלדה, MutationObserver או
סריקות מסמך שנוספו לצורך המדידה. ה־recorder אינו חלק מהבאנדל הנשלח למשתמש.

## הרצה

```sh
npm run build
# CHROME: נתיב Chrome/Chromium. ברירת המחדל נמצאת ב־scripts/cdp.mjs.
npm run bench:performance -- --out /tmp/blank-fresh.json
npm run bench:performance -- --out /tmp/blank-reload.json --cache reload
npm run bench:performance -- --docx /path/to/disposable-fixture.docx --out /tmp/restore.json
npm run bench:performance -- --autosave --out /tmp/autosave.json
# בדיקת מבנה קצרה; אינה מדגם להשוואת ביצועים.
npm run bench:performance -- --smoke --keys 8 --out /tmp/smoke.json
```

ברירת מחדל: 5 פתיחות; בכל פתיחה תו ראשון, 60 תווים ברצף, הפסקה של 3 שניות
ועוד 60 תווים. `--gap 40`, `--keys 60`, `--pause 3000`, `--runs 5` ניתנים לשינוי.
`--api-bytes` מוסיף אומדן נפח UTF-8 של JSON; לספירה ולסריאליזציה יש מחיר,
ולכן השוו ריצות עם אותו מצב. לריצות latency רגילות העדיפו ספירת קריאות בלבד.
גם אימות ה־DOM מוסיף עבודה בזמן מדידה; השוו שתי בניות עם אותו גשש ותנאים.

`--autosave` דורש הקלדה שבאמת חפפה ל־exporting/uploading/committing; אם השמירה
הסתיימה לפני הקלט, הריצה נכשלת ולא מדווחת כאילו נמדדה חפיפה. ב־headless הייצוא
של המנוע אמיתי, אך ה־SDK וכתיבת הקובץ מדומים בזיכרון. זו אינה מדידת דיסק או רשת.

## מה מוכח ומה נמדד

| שדה/שלב | משמעות וגבול |
|---|---|
| `plugin.boot`, `boot.resolve` | אירוע המארח ופתרון אירוע/שחזור; תחילת השעון היא ניווט התוסף, לא לחיצת המשתמש במארח |
| `script.requested` → `script.loaded` | זמן בקשת סקריפט עד אירוע load, כולל טעינה והרצה; לא זמן הורדה בלבד |
| `resources` | Resource Timing כשהדפדפן חושף אותו; רשימה חסרה מסומנת `resourceTimingAvailable: false`, ונפח לא ידוע הוא `null`. ב־file:// ייתכן שלא ייחשף כלל |
| `app.entry` | הכניסה למודול הראשי אחרי הערכת התלויות הסטטיות |
| `bundled-fonts` | רישום גופנים ארוזים; אינו מבטיח שכל הגופנים נטענו או אומתו |
| `preferences`, `preflight`, `document-fonts`, `engine`, `arrange`, `view-restore` | שלבים נפרדים במסלול הפתיחה; `span` מקשר אותם לניסיון הפתיחה הנכון |
| `document.open.end` | סוף מסלול הפתיחה והשחזור; אינו סמן editor-ready בפני עצמו |
| `firstEdit`, `keys` | keydown בעורך הפעיל → טקסט צפוי בבלוק הגלוי → שתי הזדמנויות animation frame; בסוף הרצף מאומתים טקסט המודל ומיקום הסמן ב־Document API |
| `api`, `apiCoverage` | ספירות מעטפת SDK וקריאות readout ציבוריות שניתן לעטוף; אינן כל ה־RPC הפנימיים במנוע. קריאות הבדיקה מסומנות `probe` בנפרד |
| `longTasks` | long tasks בחוט הראשי, כשהדפדפן תומך; אינו כולל עבודת worker |

אין API כאן שמוכיח הצגת פיקסלים על מסך Word/native. המדד נקרא
`keyToRenderOpportunityMs` במפורש, ולא key-to-paint מדויק. FCP אינו מדד מוכנות.
זמן הניווט עד העריכה הראשונה כולל גם חיבור CDP והכנת הבדיקה; האירועים
`driver.ready-observed` ו־`probe.armed` מציגים את החלק הזה. הוא תצפית על עריכה
שהצליחה, ולא הערכה לרגע המוקדם ביותר שבו היה אפשר לערוך.

הבדיקה דורשת סמן מכווץ בפסקת גוף פשוטה, גלויה במלואה ובעלת fragment יחיד.
פסקה מפוצלת/וירטואלית, inline לא נתמך, החלפת מסמך, טאב מוסתר, טקסט שונה במודל,
הקשות חסרות או חיתוך הלוג מכשילים את המדידה. כל ההקשות נשמרות; עדכוני DOM
שאספו כמה הקשות מסומנים `coalesced`. JSON אינו מכיל טקסט מסמך, שמות קבצים,
tokens, כתובות קובץ או payload של API; fixture מיוצג ב־SHA-256 וגודל בלבד.

## אוצריא Release והשוואה לפני/אחרי

פתחו עותק בדיקה שניתן לשנות, חשפו CDP במארח ובחרו את מזהה הדף המדויק מתוך
`http://127.0.0.1:PORT/json/list`:

```sh
# מודד הקלדה בלבד; מדידות פתיחה לא מיוחסות להתקנה מאוחרת.
node scripts/performance-probe.mjs --attach PORT --target PAGE_ID --edit-copy --out native-typing.json
# מוסיף recorder לפני סקריפטי הדף ומבצע 5 רענונים של התוסף.
node scripts/performance-probe.mjs --attach PORT --target PAGE_ID --edit-copy --reload --out native-reload.json
```

הרענון יכול להפעיל שחזור/שמירה של המארח; השתמשו בעותק בדיקה. הכלי אינו מפעיל
מחדש את אוצריא ואינו מאשר שזו בניית Release. `nativeReleaseVerified: false`
נשאר בדוח עד שאימות חיצוני מצורף. חומרה, גרסת דפדפן/WebView וגרסת המארח זמינות
בדוח; לחיצה במארח ו־native pixel presentation דורשות tracing חיצוני.

לכל שינוי, שמרו את שני קובצי ה־JSON עם SHA של הבניות, באותו מחשב, fixture,
WebView, רצועה/בוררים וקצב הקלדה. ה־recorder וה־hooks צריכים להימצא בשתי הבניות.
בצעו כל תא במטריצה לפני ואחרי: ריק, רגיל, וגדול עם RTL/הערות/טורים; פתיחה
חדשה ושחזור; תכונות readout כבויות ופעילות היכן שיש מתג נתמך. תעדו את המתגים
בדוח המצורף — הכלי אינו ממציא מצב כבוי לתכונה שאין לה מתג.

`fresh-profile` פותח תהליך ופרופיל Chrome חדש; `reload` מרענן באותו תהליך.
מטמון מערכת ההפעלה אינו נשלט. מדידת פתיחה קרה של אוצריא דורשת הפעלה מחדש
של המארח ופרוטוקול מטמון מתועד; אל תקראו לרענון „פתיחה קרה”.

הסיכום משתמש ב־nearest-rank p50/p95, מפריד שלבי הקלדה ומציג מספר דגימות.
ריצה כושלת נרשמת, מפסיקה את הסדרה ומחזירה exit code 1; אין השמטת כשלים שקטה.
חמש פתיחות מספיקות לסיכום בסיסי, אך p95 של חמש הוא המקסימום שלהן.

## שערים

- `npm test -- tests/unit/performance-recorder.test.ts --maxWorkers=1`
- `npm run bench:performance -- --smoke --keys 8 --out /tmp/performance-smoke.json`
- `npm run check:typing-recalc` נשמר ללא שינוי: שער מבני נגד סריקת CSS של כל העץ.
- `npm run check:startup` ממשיך לבדוק splash וצביעה ראשונה, בלי לטעון שהמסמך מוכן.

אין ספי latency קשיחים למכונה אקראית. דוחות headless מאמתים את המדידה ואת
מסלול המנוע; מטריצת native Release ונתוני לפני/אחרי של שיפור עדיין מחייבים הרצה
בסביבה המתוארת. ה־README מבחין בין גדלי הבנייה הנוכחיים למדידות ותקלות עבר.
