/**
 * plugin.boot נורה פעם אחת ואוצריא אינה משחזרת אותו: אין getBootInfo, ו-`on`
 * של ה-SDK הוא window.addEventListener בלי replay. הבדיקות כאן מקבעות את שלוש
 * שכבות ההגנה, כי הרגרסיה שלהן היא בדיוק מה שנצפה בפועל — „אוצריא לא סיימה
 * לאתחל את התוסף” בטעינה ראשונה, ותוסף שעולה רק אחרי רענון:
 *   1. ה-latch ב-<head> של index.html, שרץ לפני הבאנדל;
 *   2. קריאתו מתוך otzaria-client בזמן טעינת המודול;
 *   3. שחזור ב-RPC אם האירוע אבד בכל זאת.
 */
import { describe, it, expect, vi, afterEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { BootPayload } from '../../src/types/otzaria_plugin';

const BOOT = {
  plugin: { id: 'test', version: '1' },
  app: { version: '0.9.97', platform: 'macos' },
  theme: { mode: 'light' },
} as unknown as BootPayload;

async function freshClient(): Promise<typeof import('../../src/host/otzaria-client')> {
  vi.resetModules();
  return import('../../src/host/otzaria-client');
}

afterEach(() => {
  vi.useRealTimers();
  delete (window as Partial<Window>).Otzaria;
});

describe('waitForBoot', () => {
  it('מחזירה את ה-payload גם כשהאירוע נורה לפני ההמתנה', async () => {
    const client = await freshClient();

    window.dispatchEvent(new CustomEvent('plugin.boot', { detail: BOOT }));

    await expect(client.waitForBoot()).resolves.toBe(BOOT);
  });

  it('מחזירה את אותו payload לכל קורא', async () => {
    const client = await freshClient();
    window.dispatchEvent(new CustomEvent('plugin.boot', { detail: BOOT }));

    const [first, second] = await Promise.all([client.waitForBoot(), client.waitForBoot()]);

    expect(first).toBe(BOOT);
    expect(second).toBe(BOOT);
  });

  it('מתעלמת מאירוע boot שני', async () => {
    const client = await freshClient();
    window.dispatchEvent(new CustomEvent('plugin.boot', { detail: BOOT }));
    window.dispatchEvent(new CustomEvent('plugin.boot', { detail: { plugin: { id: 'other' } } }));

    await expect(client.waitForBoot()).resolves.toBe(BOOT);
  });

  it('נכשלת בזמן קצוב במקום להישאר תלויה', async () => {
    vi.useFakeTimers();
    const client = await freshClient();

    const pending = client.waitForBoot(100);
    const assertion = expect(pending).rejects.toThrow('אוצריא לא סיימה לאתחל את התוסף');
    await vi.advanceTimersByTimeAsync(100);

    await assertion;
  });

  it('מאזין שנרשם אחרי הירייה מפספס — הסיבה שה-latch קיים', async () => {
    await freshClient();
    window.dispatchEvent(new CustomEvent('plugin.boot', { detail: BOOT }));

    const late = vi.fn();
    window.addEventListener('plugin.boot', late);

    expect(late).not.toHaveBeenCalled();
  });
});

describe('call', () => {
  it('זורקת הודעה בעברית כשה-SDK אינו קיים', async () => {
    const client = await freshClient();

    await expect(client.call('app.getInfo')).rejects.toThrow('ה-SDK של אוצריא אינו זמין');
    expect(client.isAvailable()).toBe(false);
  });

  it('מחזירה את data כשהקריאה הצליחה', async () => {
    const client = await freshClient();
    window.Otzaria = {
      call: vi.fn(async () => ({ success: true, data: { version: '0.9.96' }, error: null })),
    } as never;

    await expect(client.call('app.getInfo')).resolves.toEqual({ version: '0.9.96' });
  });

  it('מתרגמת כשל של ה-Host לשגיאה עם ההודעה שלו', async () => {
    const client = await freshClient();
    window.Otzaria = {
      call: vi.fn(async () => ({
        success: false,
        data: null,
        error: { code: 'error.denied', message: 'ההרשאה נדחתה' },
      })),
    } as never;

    await expect(client.call('fs.pickUserFile')).rejects.toThrow('ההרשאה נדחתה');
  });

  it('tryCall מחזירה null במקום לזרוק', async () => {
    const client = await freshClient();

    await expect(client.tryCall('ui.showMessage')).resolves.toBeNull();
  });

  it('נושאת את הקוד של אוצריא ולא רק את ההודעה', async () => {
    // ההודעה היא טקסט חופשי ואינה מבטיחה להזכיר את הקוד; זיהוי „ההרשאה
    // חסרה” לפי חיפוש מחרוזת בה תלוי בנוסח שאוצריא בחרה.
    const client = await freshClient();
    window.Otzaria = {
      call: vi.fn(async () => ({
        success: false,
        data: null,
        error: { code: 'error.permission_denied', message: 'no' },
      })),
    } as never;

    const error = await client.call('navigation.goTo').catch((e: unknown) => e);

    expect(client.hostErrorCode(error)).toBe('error.permission_denied');
    expect(client.isPermissionDenied(error)).toBe(true);
    expect((error as { method?: string }).method).toBe('navigation.goTo');
  });

  it('שגיאה שאינה מאוצריא אינה מקבלת קוד', async () => {
    const client = await freshClient();

    expect(client.hostErrorCode(new Error('משהו אחר'))).toBeNull();
    expect(client.isPermissionDenied(new Error('משהו אחר'))).toBe(false);
    // הגיבוי לפי ההודעה נשמר, כי host/files.ts נשען עליו.
    expect(client.isPermissionDenied(new Error('error.permission_denied: no'))).toBe(true);
  });
});

describe('on', () => {
  it('מחזירה ביטול שקורא ל-off עם אותה הפניה', async () => {
    const client = await freshClient();
    const off = vi.fn();
    const onFn = vi.fn();
    window.Otzaria = { on: onFn, off } as never;

    const listener = vi.fn();
    const stop = client.on('theme.changed', listener);
    stop();

    expect(onFn).toHaveBeenCalledWith('theme.changed', listener);
    expect(off).toHaveBeenCalledWith('theme.changed', listener);
  });
});

describe('ה-latch ב-index.html', () => {
  // הנרמול הוא מה שמפעיל את הבדיקה הראשונה בכלל: index.html נשמר ב-CRLF,
  // וההשוואה שם מצפה למפריד שורה יחיד בתוך קריאת ה-addEventListener. בלעדיו
  // הבדיקה נכשלה תמיד — כלומר שער הסדר של ה-latch לא היה שמור בפועל.
  const html = readFileSync(join(process.cwd(), 'index.html'), 'utf8').replace(/\r\n/g, '\n');

  it('נרשם ל-plugin.boot בתוך ה-head, לפני סקריפט התוסף', () => {
    const head = html.slice(0, html.indexOf('</head>'));

    expect(head).toContain("addEventListener(\n          'plugin.boot'");
    // הסדר הוא כל העניין: ה-latch חייב לקדום לבאנדל.
    expect(html.indexOf('plugin.boot')).toBeLessThan(html.indexOf('src/main.ts'));
  });

  it('כותב לאותו שם שהקוד קורא ממנו', async () => {
    const { BOOT_LATCH_KEY } = await freshClient();

    expect(html).toContain(`window.${BOOT_LATCH_KEY} =`);
  });

  it('מודול שנטען אחרי הירייה מקבל את ה-payload מה-latch', async () => {
    // התרחיש שנצפה: האירוע נורה, ורק אחר כך הבאנדל שלנו נטען ומריץ את המודול.
    // בלי קריאת ה-latch, ההמתנה כאן לא הייתה נגמרת לעולם.
    const client0 = await freshClient();
    void client0;
    (window as unknown as Record<string, unknown>).__otzariaBoot = {
      payload: BOOT,
      at: 12,
    };

    const client = await freshClient();

    await expect(client.waitForBoot()).resolves.toBe(BOOT);
  });
});

describe('resolveBoot', () => {
  afterEach(() => {
    delete (window as unknown as Record<string, unknown>).__otzariaBoot;
  });

  it('מחזירה את מה שהאירוע הביא, בלי לגעת ב-RPC', async () => {
    const client = await freshClient();
    const call = vi.fn();
    (window as Partial<Window>).Otzaria = { call } as never;

    window.dispatchEvent(new CustomEvent('plugin.boot', { detail: BOOT }));
    const info = await client.resolveBoot();

    expect(info.source).toBe('event');
    expect(info.theme).toBe(BOOT.theme);
    expect(call).not.toHaveBeenCalled();
  });

  it('משחזרת ב-RPC כשהאירוע אבד', async () => {
    vi.useFakeTimers();
    const client = await freshClient();
    (window as Partial<Window>).Otzaria = {
      call: vi.fn(async (method: string) => ({
        success: true,
        data: method === 'app.getInfo' ? BOOT.app : BOOT.theme,
        error: null,
      })),
    } as never;

    const pending = client.resolveBoot({ graceMs: 50, pollMs: 10, timeoutMs: 1000 });
    await vi.advanceTimersByTimeAsync(60);

    await expect(pending).resolves.toMatchObject({ source: 'recovered', app: BOOT.app });
  });

  it('אינה נוגעת ב-RPC לפני חלון החסד', async () => {
    vi.useFakeTimers();
    const client = await freshClient();
    const call = vi.fn(async () => ({ success: true, data: {}, error: null }));
    (window as Partial<Window>).Otzaria = { call } as never;

    void client.resolveBoot({ graceMs: 500, pollMs: 10, timeoutMs: 5000 });
    await vi.advanceTimersByTimeAsync(400);

    expect(call).not.toHaveBeenCalled();
  });

  it('חוזרת ומנסה כל עוד ה-SDK אינו מוכן', async () => {
    vi.useFakeTimers();
    const client = await freshClient();
    let ready = false;
    const call = vi.fn(async (method: string) => {
      if (!ready) return { success: false, data: null, error: { message: 'not ready yet' } };
      return {
        success: true,
        data: method === 'app.getInfo' ? BOOT.app : BOOT.theme,
        error: null,
      };
    });
    (window as Partial<Window>).Otzaria = { call } as never;

    const pending = client.resolveBoot({ graceMs: 10, pollMs: 100, timeoutMs: 5000 });
    await vi.advanceTimersByTimeAsync(250);
    expect(call).toHaveBeenCalled();

    ready = true;
    await vi.advanceTimersByTimeAsync(200);

    await expect(pending).resolves.toMatchObject({ source: 'recovered' });
  });

  it('אירוע שמגיע באיחור גובר על שחזור שנכשל', async () => {
    vi.useFakeTimers();
    const client = await freshClient();
    (window as Partial<Window>).Otzaria = {
      call: vi.fn(async () => ({ success: false, data: null, error: { message: 'not ready' } })),
    } as never;

    const pending = client.resolveBoot({ graceMs: 10, pollMs: 50, timeoutMs: 5000 });
    await vi.advanceTimersByTimeAsync(200);
    window.dispatchEvent(new CustomEvent('plugin.boot', { detail: BOOT }));

    await expect(pending).resolves.toMatchObject({ source: 'event' });
  });

  it('מפסיקה לנסות אחרי שנפתרה', async () => {
    vi.useFakeTimers();
    const client = await freshClient();
    const call = vi.fn(async () => ({ success: false, data: null, error: { message: 'no' } }));
    (window as Partial<Window>).Otzaria = { call } as never;

    const pending = client.resolveBoot({ graceMs: 10, pollMs: 50, timeoutMs: 5000 });
    await vi.advanceTimersByTimeAsync(200);
    window.dispatchEvent(new CustomEvent('plugin.boot', { detail: BOOT }));
    await pending;

    const afterResolve = call.mock.calls.length;
    await vi.advanceTimersByTimeAsync(2000);

    // סבב polling שנשאר חי היה מציף את ה-Host בקריאות לנצח. שתי הגנות עוצרות
    // אותו — ניקוי הטיימרים ובדיקת settled בתוך attempt — וכל אחת מספיקה
    // לבדה (נבדק: הסרת אחת מהן אינה מפילה כלום, הסרת שתיהן מפילה את הבדיקה
    // הזאת). הן נשארות שתיהן בכוונה.
    expect(call.mock.calls.length).toBe(afterResolve);
  });

  it('נכשלת בזמן קצוב עם הודעה שאומרת מה לעשות', async () => {
    vi.useFakeTimers();
    const client = await freshClient();
    (window as Partial<Window>).Otzaria = {
      call: vi.fn(async () => ({ success: false, data: null, error: { message: 'no' } })),
    } as never;

    const pending = client.resolveBoot({ graceMs: 10, pollMs: 50, timeoutMs: 300 });
    const assertion = expect(pending).rejects.toThrow('נסו לטעון את הלשונית מחדש');
    await vi.advanceTimersByTimeAsync(300);

    await assertion;
  });
});
