/**
 * ה-URLs של ה-workers הם התלות היחידה שלנו במשהו שה-build מזריק ל-window.
 * הבדיקות מקבעות את שתי ההתנהגויות שאין להן ביטוי בטיפוסים: בלי הזרקה מחזירים
 * undefined (כדי ש-SuperDoc ייפול חזרה ל-worker המובנה שלו במקום לקבל URL
 * ריק), ותפקיד חסר אינו הופך ל-blob ריק.
 */
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { engineWorkerUrls, resetEngineWorkerUrlsCache } from '../../src/engine/workers';

beforeEach(() => {
  resetEngineWorkerUrlsCache();
  delete window.__SUPERDOC_WORKER_SOURCES__;
  // jsdom אינו מממש createObjectURL.
  URL.createObjectURL = vi.fn((blob: Blob) => `blob:mock/${blob.size}`);
  URL.revokeObjectURL = vi.fn();
});

afterEach(() => {
  resetEngineWorkerUrlsCache();
});

describe('טעינה קלאסית של ה-worker', () => {
  class FakeWorker {
    static calls: Array<{ url: string; options?: WorkerOptions }> = [];
    constructor(url: string | URL, options?: WorkerOptions) {
      FakeWorker.calls.push({ url: String(url), options });
    }
  }

  beforeEach(() => {
    FakeWorker.calls.length = 0;
    (window as unknown as { Worker: unknown }).Worker = FakeWorker;
  });

  it('מסיר type: module מ-URL שאנחנו בנינו', () => {
    window.__SUPERDOC_WORKER_SOURCES__ = { document: 'a' };
    const urls = engineWorkerUrls()!;

    new window.Worker(urls.document!, { type: 'module', name: 'superdoc-v2-edit' });

    expect(FakeWorker.calls).toEqual([
      { url: urls.document, options: { name: 'superdoc-v2-edit' } },
    ]);
  });

  it('אינו נוגע ב-URL שאינו שלנו', () => {
    window.__SUPERDOC_WORKER_SOURCES__ = { document: 'a' };
    engineWorkerUrls();

    new window.Worker('https://example.test/w.js', { type: 'module' });

    expect(FakeWorker.calls).toEqual([
      { url: 'https://example.test/w.js', options: { type: 'module' } },
    ]);
  });

  it('אינו נוגע בבנייה שאין בה type: module', () => {
    window.__SUPERDOC_WORKER_SOURCES__ = { document: 'a' };
    const urls = engineWorkerUrls()!;

    new window.Worker(urls.document!, { name: 'x' });

    expect(FakeWorker.calls).toEqual([{ url: urls.document, options: { name: 'x' } }]);
  });
});

describe('engineWorkerUrls', () => {
  it('בלי קוד מוטמע מחזיר undefined', () => {
    expect(engineWorkerUrls()).toBeUndefined();
  });

  it('בונה blob לכל תפקיד שיש לו קוד', () => {
    window.__SUPERDOC_WORKER_SOURCES__ = { document: 'a', reviewIndex: 'bb' };

    expect(engineWorkerUrls()).toEqual({ document: 'blob:mock/1', reviewIndex: 'blob:mock/2' });
  });

  it('מדלג על תפקיד בלי קוד במקום לבנות blob ריק', () => {
    window.__SUPERDOC_WORKER_SOURCES__ = { document: 'a', reviewIndex: '' };

    expect(engineWorkerUrls()).toEqual({ document: 'blob:mock/1' });
  });

  it('מתעלם ממפתחות שאינם תפקיד מוכר', () => {
    window.__SUPERDOC_WORKER_SOURCES__ = { document: 'a', collaboration: 'x' };

    expect(engineWorkerUrls()).toEqual({ document: 'blob:mock/1' });
  });

  it('בונה את ה-URLs פעם אחת בלבד', () => {
    window.__SUPERDOC_WORKER_SOURCES__ = { document: 'a' };

    const first = engineWorkerUrls();

    expect(engineWorkerUrls()).toBe(first);
    expect(URL.createObjectURL).toHaveBeenCalledTimes(1);
  });

  it('משחרר את מקורות ה-workers ושומר את הכתובות לעורכים נוספים', () => {
    window.__SUPERDOC_WORKER_SOURCES__ = { document: 'a', reviewIndex: 'bb' };

    const first = engineWorkerUrls();

    expect(window.__SUPERDOC_WORKER_SOURCES__).toBeUndefined();
    expect(engineWorkerUrls()).toBe(first);
    expect(engineWorkerUrls()).toEqual({ document: 'blob:mock/1', reviewIndex: 'blob:mock/2' });
    expect(URL.createObjectURL).toHaveBeenCalledTimes(2);
    expect(URL.revokeObjectURL).not.toHaveBeenCalled();
  });

  it('מאמץ את כתובת החימום בלי לבנות או לבטל אותה', () => {
    window.__SUPERDOC_WORKER_SOURCES__ = { document: 'a', reviewIndex: 'bb' };
    window.__otzariaDocWorkerUrl = 'blob:warm/document';

    const first = engineWorkerUrls();

    expect(first).toEqual({ document: 'blob:warm/document', reviewIndex: 'blob:mock/2' });
    expect(engineWorkerUrls()).toBe(first);
    expect(window.__SUPERDOC_WORKER_SOURCES__).toBeUndefined();
    expect(URL.createObjectURL).toHaveBeenCalledTimes(1);
    expect(URL.revokeObjectURL).not.toHaveBeenCalled();
  });

  it('משאיר מקורות לניסיון חוזר ומבטל רק כתובות חדשות אחרי כשל חלקי', () => {
    const sources = { document: 'a', reviewIndex: 'bb' };
    window.__SUPERDOC_WORKER_SOURCES__ = sources;
    vi.mocked(URL.createObjectURL)
      .mockReturnValueOnce('blob:first/document')
      .mockImplementationOnce(() => { throw new Error('blob creation failed'); })
      .mockReturnValueOnce('blob:retry/document')
      .mockReturnValueOnce('blob:retry/review');

    expect(() => engineWorkerUrls()).toThrow('blob creation failed');
    expect(window.__SUPERDOC_WORKER_SOURCES__).toBe(sources);
    expect(URL.revokeObjectURL).toHaveBeenCalledExactlyOnceWith('blob:first/document');

    expect(engineWorkerUrls()).toEqual({ document: 'blob:retry/document', reviewIndex: 'blob:retry/review' });
    expect(window.__SUPERDOC_WORKER_SOURCES__).toBeUndefined();
    expect(URL.revokeObjectURL).toHaveBeenCalledTimes(1);
    expect(URL.createObjectURL).toHaveBeenCalledTimes(4);
  });

  it('אינו מבטל את כתובת החימום כאשר יצירת כתובת הביקורת נכשלת', () => {
    const sources = { document: 'a', reviewIndex: 'bb' };
    window.__SUPERDOC_WORKER_SOURCES__ = sources;
    window.__otzariaDocWorkerUrl = 'blob:warm/document';
    vi.mocked(URL.createObjectURL)
      .mockImplementationOnce(() => { throw new Error('review failed'); })
      .mockReturnValueOnce('blob:retry/review');

    expect(() => engineWorkerUrls()).toThrow('review failed');
    expect(window.__SUPERDOC_WORKER_SOURCES__).toBe(sources);
    expect(window.__otzariaDocWorkerUrl).toBe('blob:warm/document');
    expect(URL.revokeObjectURL).not.toHaveBeenCalled();

    expect(engineWorkerUrls()).toEqual({ document: 'blob:warm/document', reviewIndex: 'blob:retry/review' });
    expect(window.__SUPERDOC_WORKER_SOURCES__).toBeUndefined();
    expect(URL.revokeObjectURL).not.toHaveBeenCalled();
  });
});
