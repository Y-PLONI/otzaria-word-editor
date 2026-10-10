/**
 * צבעי המעטפת נמדדים ב-Chrome על ה-dist: מחרוזת תקינה ב-tokens.css אינה
 * מוכיחה איזה כלל ניצח במפל הסגנונות. מכסה החלפת נושא דרך אירוע SDK,
 * בחירת בד ושחזור ברירת המחדל דרך הבורר, קריאות פקדים ופופאוברים,
 * וגאומטריה יציבה של טאבים. הרצה: npm run check:shell-colors.
 */
import { openApp, createReport, sleep } from './harness.mjs';

const report = createReport('צבעי מעטפת ובד', { strict: true });
const THEMES = [
  {
    mode: 'light', colorScheme: {
      primary: '#1565c0', onPrimary: '#ffffff', surface: '#f8f9fa',
      onSurface: '#1a1a2e', onSurfaceVariant: '#49454f',
      surfaceContainerHigh: '#f3f2f1', surfaceContainerHighest: '#edebe9', outline: '#cbd5e1',
    },
  },
  {
    // onPrimary לבן תקין מעל primary כחול גם בנושא כהה. שימוש בו כרקע
    // לפקדי onSurface היה הופך אותם לבהירים על לבן.
    mode: 'dark', colorScheme: {
      primary: '#1565c0', onPrimary: '#ffffff', surface: '#101014',
      onSurface: '#e6e6e6', onSurfaceVariant: '#c9c5d0',
      surfaceContainerHigh: '#2b2930', surfaceContainerHighest: '#36343b', outline: '#938f99',
    },
  },
];

function check(name, ok, detail = '') {
  report[ok ? 'pass' : 'fail'](name, detail);
}

function rgb(css) {
  const parts = css.match(/[\d.]+/g)?.map(Number);
  if (css.startsWith('color(srgb ') && parts?.length >= 3) return parts.slice(0, 3).map((c) => c * 255);
  if (css.startsWith('rgb') && parts?.length >= 3) return parts.slice(0, 3);
  throw new Error(`צבע מחושב לא מוכר: ${css}`);
}

function hexRgb(hex) {
  return `rgb(${hex.slice(1).match(/../g).map((c) => parseInt(c, 16)).join(', ')})`;
}

// ±1 לערוץ: `color-mix` מחושב בדפדפן ומעוגל אחרת מהחישוב כאן.
function near(css, expected) {
  const a = rgb(css);
  return rgb(expected).every((c, i) => Math.abs(c - a[i]) <= 1);
}

/**
 * הפס העליון וברירת המחדל של הבד: 35% שורת טאבי המסמכים
 * (`surfaceContainerHigh`) ו-65% משטח — בהיר מהשורה ונגזר מהנושא.
 */
function expectedBand(theme) {
  const high = rgb(hexRgb(theme.colorScheme.surfaceContainerHigh));
  const surface = rgb(hexRgb(theme.colorScheme.surface));
  return `rgb(${high.map((c, i) => Math.round(c * 0.35 + surface[i] * 0.65)).join(', ')})`;
}

function compositeBlackOverlay(background, image) {
  if (image === 'none') return background;
  const match = image.match(/rgba\(0,\s*0,\s*0,\s*([\d.]+)\)/);
  if (!match) throw new Error(`שכבת רקע לא מוכרת: ${image}`);
  const alpha = Number(match[1]);
  return `rgb(${rgb(background).map((channel) => Math.round(channel * (1 - alpha))).join(', ')})`;
}

function contrast(fg, bg, highlight = 0) {
  const luminance = (channels) => channels.map((c) => {
    const v = c / 255;
    return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
  }).reduce((sum, v, i) => sum + v * [0.2126, 0.7152, 0.0722][i], 0);
  const a = luminance(rgb(fg));
  const b = luminance(rgb(bg).map((c) => c * (1 - highlight) + 255 * highlight));
  return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
}

const app = await openApp({
  name: 'shell-colors', port: Number(process.env.QA_PORT ?? 9651),
  extra: '<script>window.__qaHost.storage["ruler-visible"] = true;</script>',
});

async function measure() {
  return JSON.parse(await app.js(`JSON.stringify((() => {
    const style = (selector) => {
      const el = document.querySelector(selector);
      if (!el) throw new Error('חסר ' + selector);
      const c = getComputedStyle(el);
      return { background: c.backgroundColor, color: c.color, image: c.backgroundImage };
    };
    return {
      theme: document.documentElement.dataset.theme,
      canvas: style('.editor-stack'),
      titlebar: style('.word-titlebar'),
      doctabsBar: style('.word-doctabs-bar'),
      swatch: style('[data-tip-title="צבע רקע העורך"] .color-indicator-bar').background,
      stored: window.__qaHost.storage['canvas-color'] ?? null,
      chrome: {
        ribbon: style('.word-ribbon-body'),
        ribbonButton: style('.word-ribbon-body .word-btn'),
        status: style('.word-statusbar'),
        active: style('.word-doctab.active'),
        ruler: style('.doc-ruler'), verticalRuler: style('.doc-vruler'),
      },
    };
  })())`));
}

async function palette() {
  if (!await app.click('צבע רקע העורך', { after: 100 })) throw new Error('בורר הרקע לא נפתח');
}

async function checkScrollbarMatchesStatusBar(theme) {
  const point = JSON.parse(await app.js(`JSON.stringify((() => {
    const r = document.querySelector('.editor-stack__host').getBoundingClientRect();
    return { x: r.left + 50, y: r.top + 50 };
  })())`));
  await app.cdp.send('Input.dispatchMouseEvent', {
    type: 'mouseMoved', x: point.x, y: point.y,
  });
  await sleep(150);
  const state = JSON.parse(await app.js(`JSON.stringify((() => {
    // רק scrollbarColor: getComputedStyle על '::-webkit-scrollbar-track'
    // מחזיר את הכלל גם כש-scrollbar-color מבטל אותו ושום דבר לא מצויר.
    const host = document.querySelector('.editor-stack__host');
    return {
      standard: getComputedStyle(host).scrollbarColor,
      statusColor: getComputedStyle(document.querySelector('.word-statusbar')).backgroundColor,
    };
  })())`));
  const expectedStandardTrack = hexRgb(theme.colorScheme.surface);
  check(
    `${theme.mode} — רקע פס הגלילה תואם למשטח שורת המצב`,
    state.standard.endsWith(expectedStandardTrack) && state.statusColor === expectedStandardTrack,
    `תקני=${state.standard}, שורת מצב=${state.statusColor}`,
  );
  await app.cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: 600, y: 15 });
  await sleep(150);
}

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1400, height: 900, deviceScaleFactor: 1, mobile: false });
  await app.tab('תצוגה');

  for (const theme of THEMES) {
    await app.js(`window.__qaHost.emit('theme.changed', ${JSON.stringify(theme)})`);
    await sleep(100);
    const base = await measure();
    const band = expectedBand(theme);
    const expectedRibbonBg = theme.mode === 'light'
      ? 'rgb(255, 255, 255)'
      : hexRgb(theme.colorScheme.surfaceContainerHigh);
    check(`${theme.mode} — הנושא הוחל דרך ה-SDK`, base.theme === theme.mode);
    check(
      `${theme.mode} — רקע רצועת הכלים`,
      base.chrome.ribbon.background === expectedRibbonBg,
      `רצועה=${base.chrome.ribbon.background}, צפוי=${expectedRibbonBg}`,
    );
    check(`${theme.mode} — הבד בגוון הפס, בלי שכבה, והפס בבורר מראה אותו`,
      near(base.canvas.background, band) && base.canvas.image === 'none' && near(base.swatch, band),
      `בד=${base.canvas.background} ${base.canvas.image}, פס=${base.swatch}, צפוי=${band}`);
    check(`${theme.mode} — שורת טאבי המסמכים בצבע הנושא, והפס שונה ממנה`,
      near(base.doctabsBar.background, hexRgb(theme.colorScheme.surfaceContainerHigh))
        && !near(base.titlebar.background, base.doctabsBar.background),
      `פס=${base.titlebar.background}, שורת טאבים=${base.doctabsBar.background}`);
    check(`${theme.mode} — הפס העליון והטאב הפעיל בצבע הפס, נפרדים מהרצועה`,
      near(base.titlebar.background, band) && base.titlebar.image === 'none'
        && near(base.chrome.active.background, band)
        && !near(base.chrome.ribbon.background, band),
      `כותרת=${base.titlebar.background} ${base.titlebar.image}, טאב=${base.chrome.active.background}, צפוי=${band}, רצועה=${base.chrome.ribbon.background}`);

    const pairs = [
      ['רצועה', base.chrome.ribbonButton.color, base.chrome.ribbon.background, 0],
      ['שורת מצב', base.chrome.status.color, base.chrome.status.background, 0.04],
      ['טאב פעיל', base.chrome.active.color, base.chrome.active.background, 0],
    ];
    for (const [name, fg, bg, highlight] of pairs) {
      const ratio = contrast(fg, bg, highlight);
      check(`${theme.mode} — ${name} קריא`, ratio >= 4.5, `${ratio.toFixed(2)}:1`);
    }

    for (const color of ['#ffffff', '#000000', '#3366cc']) {
      await palette();
      // אותה פעולת change שהבורר המקורי של הדפדפן פולט; כל מסלול Vue,
      // applyCanvasColor והאחסון נשאר אמיתי.
      await app.js(`(() => {
        const input = document.querySelector('.custom-color-input');
        if (!input) throw new Error('חסר input של צבע');
        input.value = ${JSON.stringify(color)};
        input.dispatchEvent(new Event('change', { bubbles: true }));
      })()`);
      await sleep(100);
      const state = await measure();
      const expected = hexRgb(color);
      check(`${theme.mode} — ${color} צובע את הבד ומשתקף בבורר ובאחסון`,
        state.canvas.background === expected && state.swatch === expected && state.stored === color && state.canvas.image === 'none',
        `בד=${state.canvas.background}, פס=${state.swatch}, נשמר=${state.stored}`);
      check(`${theme.mode} — ${color} אינו צובע את פקדי המעטפת והסרגלים`,
        JSON.stringify(state.chrome) === JSON.stringify(base.chrome)
          && JSON.stringify(state.titlebar) === JSON.stringify(base.titlebar));
    }

    await palette();
    if (!await app.clickSel('.palette-clear-btn', 0, { after: 100 })) throw new Error('כפתור ברירת המחדל אינו נגיש');
    const reset = await measure();
    check(`${theme.mode} — איפוס מחזיר את הבד ואת הפס לברירת המחדל`,
      near(reset.canvas.background, band) && near(reset.swatch, band) && reset.stored === null,
      `בד=${reset.canvas.background}, פס=${reset.swatch}, צפוי=${band}`);
    await app.escape();

    // שינוי בנושא מבלי לפתוח מחדש את המסמך חייב לצבוע גם פופאובר אמיתי.
    await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 560, height: 900, deviceScaleFactor: 1, mobile: false });
    await app.tab('בית');
    if (!await app.clickSel('.word-ribbon-group--collapsed .word-group-chip', 0, { after: 100 })) throw new Error('לא נמצאה קבוצה מכווצת');
    const popup = JSON.parse(await app.js(`JSON.stringify((() => {
      const p = document.querySelector('.word-ribbon-group--collapsed.is-open .word-group-panel');
      if (!p) throw new Error('הפופאובר אינו פתוח');
      return { background: getComputedStyle(p).backgroundColor, color: getComputedStyle(p.querySelector('.word-btn')).color };
    })())`));
    const popupRatio = contrast(popup.color, popup.background);
    check(`${theme.mode} — פופאובר הקבוצה קריא`, popupRatio >= 4.5, `${popupRatio.toFixed(2)}:1`);
    await app.escape();
    await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1400, height: 900, deviceScaleFactor: 1, mobile: false });
    await app.tab('תצוגה');
    await checkScrollbarMatchesStatusBar(theme);
  }

  const hostRect = JSON.parse(await app.js(`JSON.stringify((() => {
    const r = document.querySelector('.editor-stack__host').getBoundingClientRect();
    return { x: r.left + 50, y: r.top + 50 };
  })())`));
  const scrollbar = () => app.js(`getComputedStyle(document.querySelector('.editor-stack__host')).scrollbarColor`);
  await app.cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: hostRect.x, y: hostRect.y });
  await sleep(150);
  const visible = await scrollbar();
  const canvasStyle = JSON.parse(await app.js(`JSON.stringify((() => {
    const style = getComputedStyle(document.querySelector('.editor-stack'));
    return { background: style.backgroundColor, image: style.backgroundImage };
  })())`));
  const effectiveCanvas = compositeBlackOverlay(canvasStyle.background, canvasStyle.image);
  await app.cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: 600, y: 15 });
  await sleep(150);
  const hidden = await scrollbar();
  check('פס הגלילה נחשף רק בריחוף על אזור המסמך', visible !== hidden && hidden.includes('rgba(0, 0, 0, 0)'), `${visible} → ${hidden}`);
  const railTrack = hexRgb(THEMES[1].colorScheme.surface);
  check(
    'מסילת הגלילה נשארת בגוון נפרד מהקנבס',
    visible.includes(railTrack) && railTrack !== effectiveCanvas,
    `מסילה=${visible}, קנבס=${effectiveCanvas}`,
  );

  if (!await app.clickSel('.word-doctabs-new')) throw new Error('כפתור מסמך חדש אינו נגיש');
  await sleep(1500);
  const widths = () => app.js(`JSON.stringify([...document.querySelectorAll('.word-doctab')].map(e => e.getBoundingClientRect().width))`);
  const before = JSON.parse(await widths());
  if (!await app.clickSel('.word-doctab', 0, { after: 100 })) throw new Error('הטאב הראשון אינו נגיש');
  const after = JSON.parse(await widths());
  check('מעבר מסמך אינו משנה רוחבי טאבים', before.length === 2 && JSON.stringify(before) === JSON.stringify(after), `${before} → ${after}`);
} catch (error) {
  report.fail('הריצה', error.message);
} finally {
  app.close();
}
process.exit(report.print() > 0 ? 1 : 0);
