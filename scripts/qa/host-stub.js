/**
 * דמה של מאחז אוצריא, בשביל שערי ה-QA שרצים על ה-dist הארוז ב-file://.
 *
 * זה אינו `src/host/dev-stub.ts`: הדמה ההוא נטען בתוך הבאנדל של הפיתוח, וכאן
 * צריך משהו שיושב **לפני** הבאנדל, בדף עצמו, כמו שאוצריא מזריקה בפועל. הוא גם
 * מקליט: כל קריאה נשמרת ב-`window.__qaHost.calls`, וכך שער יכול לשאול „האם
 * הכפתור הזה באמת פנה למאחז”.
 */
(function () {
  var BOOT = {
    plugin: { id: 'otzaria-word-qa', version: '0' },
    app: { version: '9.9.9', platform: 'qa', language: 'he' },
    theme: { mode: 'light', colorScheme: {}, typography: {} },
    connectivity: { isOnline: false },
    // כמו במארח, app.info.read היא הרשאת בסיס ולכן גם מופיעה בצילום ההרשאות.
    permissions: ['app.info.read', 'storage', 'clipboard.read', 'fs.read', 'fs.write'],
  };

  var H = (window.__qaHost = {
    calls: [],
    storage: {},
    /** מה שהמאחז יענה על מתודה מסוימת. שער יכול לדרוס לפני שהוא לוחץ. */
    replies: {},
    /** מה שהוצג למשתמש דרך המאחז — showError / showMessage / showConfirm. */
    messages: [],
    confirmAnswer: true,
    /** מדמה אירוע SDK אמיתי, למשל החלפת נושא אחרי העלייה. */
    emit: function (event, payload) {
      (listeners[event] || []).slice().forEach(function (handler) { handler(payload); });
    },
    /** מה שנשלח ל-`ui.exportPdf`, ומה שהמאחז יענה עליו. */
    exportPdfCalls: [],
    exportPdfReply: null,
    reset: function () {
      H.calls.length = 0;
      H.messages.length = 0;
      H.exportPdfCalls.length = 0;
      H.exportPdfReply = null;
    },
  });

  function ok(data) {
    return Promise.resolve({ success: true, data: data === undefined ? null : data, error: null });
  }
  function fail(message, code) {
    return Promise.resolve({
      success: false,
      data: null,
      error: { message: message, code: code || 'error.not_supported' },
    });
  }

  var listeners = {};

  window.Otzaria = {
    call: function (method, payload) {
      H.calls.push({ method: method, payload: payload, t: Math.round(performance.now()) });

      if (Object.prototype.hasOwnProperty.call(H.replies, method)) {
        var custom = H.replies[method];
        return typeof custom === 'function' ? custom(payload) : ok(custom);
      }

      switch (method) {
        case 'app.getInfo':
          return ok(BOOT.app);
        case 'app.getTheme':
          return ok(BOOT.theme);
        case 'app.getGrantedPermissions':
          return ok({ permissions: BOOT.permissions });
        /*
         * הערך **עצמו**, לא `{ value }`: `call()` מוסר את `res.data` כמו שהוא
         * (host/otzaria-client.ts), וכל הצרכנים קוראים ישירות —
         * `loadAutosaveEnabled` בודק `raw !== false`, `loadRulerVisible` בודק
         * `=== true`, `loadSpellcheckWords` בודק `Array.isArray`. עטיפה כאן
         * אינה נכשלת אלא **שותקת**: כל העדפה שנזרעה נדחית בנרמול ונופלת
         * לברירת המחדל, ולכן שער שזורע העדפה ומצפה לראות אותה מודד בפועל
         * את המסלול הריק — ועובר בירוק. `host/dev-stub.ts` ו-
         * `scripts/session-probe.mjs` מחזירים לא עטוף, וכך גם המאחז האמיתי.
         *
         * מפתח שאינו קיים חוזר כ-`null` (דרך `ok`, שממיר `undefined`) ולא
         * כאובייקט ריק — אחרת השומר של „אין ערך” ב-`loadSetting` אינו נורה.
         *
         * `settings-restore-qa.mjs` הוא מה שמחזיק את זה במקום.
         */
        case 'storage.get':
          return ok(H.storage[payload && payload.key]);
        case 'storage.set':
          H.storage[payload.key] = payload.value;
          return ok({});
        case 'storage.remove':
          delete H.storage[payload.key];
          return ok({});
        case 'ui.showError':
        case 'ui.showMessage':
        case 'ui.showSuccess':
          H.messages.push({ method: method, text: payload && payload.message });
          return ok({});
        case 'ui.showConfirm':
          H.messages.push({ method: method, text: payload && (payload.content || payload.title) });
          return ok({ confirmed: H.confirmAnswer });
        // ייצוא ל-PDF: הדמה רושם את מה שנשלח ומחזיר תשובה שהשער קובע
        // (`H.exportPdfReply`), כדי שאפשר יהיה למדוד גם „נשמר”, גם „בוטל”
        // וגם כשל — בלי דיאלוג מערכת.
        case 'ui.exportPdf':
          H.exportPdfCalls.push(payload || {});
          if (H.exportPdfReply && H.exportPdfReply.__throw) {
            return fail(H.exportPdfReply.message || 'נכשל', H.exportPdfReply.code);
          }
          return ok(H.exportPdfReply || { saved: true, name: 'מסמך.pdf' });
        default:
          return fail('אין תמיכה בדמה: ' + method);
      }
    },
    on: function (event, handler) {
      (listeners[event] = listeners[event] || []).push(handler);
    },
    off: function (event, handler) {
      var list = listeners[event] || [];
      var i = list.indexOf(handler);
      if (i >= 0) list.splice(i, 1);
    },
  };

  window.dispatchEvent(new CustomEvent('plugin.boot', { detail: BOOT }));
})();
