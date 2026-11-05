import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  forgetDocumentReads,
  invalidateDocumentReads,
  observeDocumentReadInvalidation,
} from '../../src/engine/document-read-state';
import type { DocReceipt } from '../../src/engine/document-api';
import {
  applyLineNumbering,
  applyPageMargins,
  createLineNumberingModel,
  createPageBorderModel,
  readLineNumbering,
  readPageBorders,
  readPageMargins,
  type PageSections,
  type PageSetupDocumentApi,
  type PageSetupHost,
} from '../../src/engine/page-setup';
import { createSectionSnapshot } from '../../src/engine/section-snapshot';

type Section = NonNullable<PageSections['items']>[number];

function section(overrides: Partial<Section> = {}): Section {
  const side = { style: 'single', size: 4, space: 12, color: 'auto' };
  return {
    address: { sectionId: 'section-1' },
    pageSetup: { width: 8.5, height: 11 },
    margins: { top: 1, bottom: 1, left: 1, right: 1 },
    sectionDirection: 'rtl',
    lineNumbering: { enabled: true, countBy: 1, start: 1, restart: 'continuous' },
    pageBorders: { top: side, bottom: side, left: side, right: side },
    ...overrides,
  };
}

function fixture(initial: PageSections = { items: [section()] }) {
  const list = vi.fn(async (): Promise<PageSections | undefined> => initial);
  const doc: PageSetupDocumentApi = { sections: { list } };
  const host = { activeEditor: { doc } } satisfies PageSetupHost;
  return { host, doc, list };
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: Error) => void;
  const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

afterEach(() => vi.useRealTimers());

describe('createSectionSnapshot', () => {
  it('הסרגל, הסמן, הגבולות והמספור חולקים בקשה ממתינה וגם תשובה שכבר התקבלה', async () => {
    const { host, list } = fixture();
    const pending = deferred<PageSections>();
    list.mockReturnValueOnce(pending.promise);
    const snapshot = createSectionSnapshot(host);
    const readers = () => Promise.all([
      readPageMargins(host, snapshot.read),
      readPageMargins(host, snapshot.read),
      readPageBorders(host, snapshot.read),
      readLineNumbering(host, snapshot.read),
    ]);

    const first = readers();
    await Promise.resolve();
    expect(list).toHaveBeenCalledTimes(1);
    pending.resolve({ items: [section()] });
    const [ruler, cursor, borders, lines] = await first;
    expect(ruler?.leftTwips).toBe(1440);
    expect(cursor).toEqual(ruler);
    expect(borders?.top.style).toBe('single');
    expect(lines?.page).toEqual(ruler);

    await readers();
    expect(list).toHaveBeenCalledTimes(1);
    snapshot.dispose();
  });

  it('שומר נתונים גולמיים בלבד: שינוי פריסה מעדכן שוליים אפקטיביים בלי list נוסף', async () => {
    const { host, list } = fixture();
    let topPx = 96;
    let bottomPx = 96;
    const measuredHost: PageSetupHost = {
      activeEditor: {
        ...host.activeEditor,
        pageMetrics: { getSnapshot: () => ({ pages: [{ base: { marginTopPx: topPx, marginBottomPx: bottomPx } }] }) },
      },
    };
    const snapshot = createSectionSnapshot(measuredHost);
    expect((await readPageMargins(measuredHost, snapshot.read))?.effectiveTopTwips).toBe(1440);

    topPx = 192;
    bottomPx = 144;
    const [page, lines] = await Promise.all([
      readPageMargins(measuredHost, snapshot.read),
      readLineNumbering(measuredHost, snapshot.read),
    ]);
    expect(page).toMatchObject({ topTwips: 1440, bottomTwips: 1440, effectiveTopTwips: 2880, effectiveBottomTwips: 2160 });
    expect(lines?.page).toEqual(page);
    expect(list).toHaveBeenCalledTimes(1);
    snapshot.dispose();
  });

  it('מספור שורות קורא מקטעים פעם אחת גם ללא מטמון', async () => {
    const { host, list } = fixture();
    expect(await readLineNumbering(host)).toMatchObject({ countBy: 1, page: { leftTwips: 1440 } });
    expect(list).toHaveBeenCalledTimes(1);
  });

  it('משמר את receiver של sections.list גם כשהמימוש סינכרוני', async () => {
    const raw = { items: [section()] };
    const sections = {
      list: vi.fn(function (this: object) {
        expect(this).toBe(sections);
        return raw;
      }),
    };
    const snapshot = createSectionSnapshot({ activeEditor: { doc: { sections } } });
    expect(await Promise.all([snapshot.read(), snapshot.read()])).toEqual([raw, raw]);
    expect(sections.list).toHaveBeenCalledTimes(1);
    snapshot.dispose();
  });

  it.each(['success', 'failure'] as const)('תשובה מגרסה קודמת (%s) אינה מתפרסמת או דורסת תשובה חדשה', async (outcome) => {
    const { host, doc, list } = fixture();
    const old = deferred<PageSections>();
    const fresh = { items: [section({ margins: { top: 2, bottom: 1, left: 1, right: 1 } })] };
    list.mockReturnValueOnce(old.promise).mockResolvedValueOnce(fresh);
    const snapshot = createSectionSnapshot(host);
    const first = snapshot.read();
    await Promise.resolve();
    invalidateDocumentReads(doc);
    expect(await snapshot.read()).toBe(fresh);
    if (outcome === 'success') old.resolve({ items: [section()] });
    else old.reject(new Error('old request failed'));
    expect(await first).toBeNull();
    expect(await snapshot.read()).toBe(fresh);
    expect(list).toHaveBeenCalledTimes(2);
    snapshot.dispose();
  });

  it('פסילה חלה גם על קריאה שפגעה במטמון שכבר נפתר', async () => {
    const { host, doc, list } = fixture();
    const snapshot = createSectionSnapshot(host);
    await snapshot.read();
    const cached = snapshot.read();
    invalidateDocumentReads(doc);
    expect(await cached).toBeNull();
    expect(await snapshot.read()).not.toBeNull();
    expect(list).toHaveBeenCalledTimes(2);
    snapshot.dispose();
  });

  it('החלפת doc באותו host אינה מעבירה נתונים או בקשות בין מסמכים', async () => {
    const firstDoc = fixture();
    const secondDoc = fixture({ items: [section({ sectionDirection: 'ltr' })] });
    const old = deferred<PageSections>();
    firstDoc.list.mockReturnValueOnce(old.promise);
    const snapshot = createSectionSnapshot(firstDoc.host);
    const first = snapshot.read();
    await Promise.resolve();

    firstDoc.host.activeEditor.doc = secondDoc.doc;
    expect((await readPageMargins(firstDoc.host, snapshot.read))?.direction).toBe('ltr');
    old.resolve({ items: [section()] });
    expect(await first).toBeNull();
    expect((await readPageMargins(firstDoc.host, snapshot.read))?.direction).toBe('ltr');
    expect(firstDoc.list).toHaveBeenCalledTimes(1);
    expect(secondDoc.list).toHaveBeenCalledTimes(1);
    snapshot.dispose();
  });

  it('שני sessions נשארים עצמאיים גם כשהפסילה והקריאות חופפות', async () => {
    const a = fixture();
    const b = fixture({ items: [section({ sectionDirection: 'ltr' })] });
    const first = createSectionSnapshot(a.host);
    const second = createSectionSnapshot(b.host);
    await Promise.all([first.read(), second.read()]);
    invalidateDocumentReads(a.doc);
    const [rtl, ltr] = await Promise.all([
      readPageMargins(a.host, first.read),
      readPageMargins(b.host, second.read),
    ]);
    expect(rtl?.direction).toBe('rtl');
    expect(ltr?.direction).toBe('ltr');
    expect(a.list).toHaveBeenCalledTimes(2);
    expect(b.list).toHaveBeenCalledTimes(1);
    first.dispose();
    second.dispose();
  });

  it.each(['dispose', 'forget'] as const)('%s מבטל תוצאה ממתינה', async (action) => {
    const { host, doc, list } = fixture();
    const pending = deferred<PageSections>();
    list.mockReturnValueOnce(pending.promise);
    const snapshot = createSectionSnapshot(host);
    const result = snapshot.read();
    await Promise.resolve();
    if (action === 'dispose') snapshot.dispose();
    else forgetDocumentReads(doc);
    pending.resolve({ items: [section()] });
    expect(await result).toBeNull();
    snapshot.dispose();
    expect(await snapshot.read()).toBeNull();
    expect(list).toHaveBeenCalledTimes(1);
  });

  it('פירוק לפני תחילת הבקשה אינו פונה למנוע', async () => {
    const { host, list } = fixture();
    const snapshot = createSectionSnapshot(host);
    const result = snapshot.read();
    snapshot.dispose();
    expect(await result).toBeNull();
    expect(list).not.toHaveBeenCalled();
  });

  it.each(['rejection', 'undefined'] as const)('כשל זמני (%s) משותף אך אינו נשמר לנצח', async (failure) => {
    const { host, list } = fixture();
    if (failure === 'rejection') list.mockRejectedValueOnce(new Error('not ready'));
    else list.mockResolvedValueOnce(undefined);
    const snapshot = createSectionSnapshot(host);
    expect(await Promise.all([snapshot.read(), snapshot.read()])).toEqual([null, null]);
    expect(list).toHaveBeenCalledTimes(1);
    expect(await snapshot.read()).not.toBeNull();
    expect(await snapshot.read()).not.toBeNull();
    expect(list).toHaveBeenCalledTimes(2);
    snapshot.dispose();
  });

  it.each([{}, { items: [] }] satisfies PageSections[])(
    'תשובה ללא מקטעים (%j) משותפת במקביל אך ניתנת לרענון בלי invalidation',
    async (empty) => {
      const { host, list } = fixture();
      list.mockResolvedValueOnce(empty);
      const snapshot = createSectionSnapshot(host);
      expect(await Promise.all([snapshot.read(), snapshot.read()])).toEqual([empty, empty]);
      expect(list).toHaveBeenCalledTimes(1);
      expect((await snapshot.read())?.items).toHaveLength(1);
      expect((await snapshot.read())?.items).toHaveLength(1);
      expect(list).toHaveBeenCalledTimes(2);
      snapshot.dispose();
    },
  );

  it('API חסר מוחזר כ-null', async () => {
    const unavailable = createSectionSnapshot({ activeEditor: { doc: {} } });
    expect(await unavailable.read()).toBeNull();
    unavailable.dispose();
  });
});

describe('כתיבות מול snapshot של תצוגה', () => {
  it('כתיבה קוראת כתובות והגדרות טריות ופוסלת את snapshot התצוגה', async () => {
    const { host, doc, list } = fixture();
    const snapshot = createSectionSnapshot(host);
    await snapshot.read();
    const latest = section({
      address: { sectionId: 'replacement-section' },
      lineNumbering: { enabled: true, countBy: 5, start: 12, distance: 0.5, restart: 'continuous' },
    });
    list.mockResolvedValue({ items: [latest] });
    const setLineNumbering = vi.fn(async () => ({ success: true }));
    doc.sections!.setLineNumbering = setLineNumbering;

    expect((await applyLineNumbering(host, 'newPage')).ok).toBe(true);
    expect(setLineNumbering).toHaveBeenCalledWith({
      target: latest.address, enabled: true, countBy: 5, start: 12, distance: 0.5, restart: 'newPage',
    });
    expect((await readLineNumbering(host, snapshot.read))?.countBy).toBe(5);
    expect(list).toHaveBeenCalledTimes(3);
    snapshot.dispose();
  });

  it.each(['throw', 'receipt'] as const)('כשל חלקי (%s) פוסל את המסמך המקורי גם אחרי מעבר טאב', async (failure) => {
    const a = fixture();
    const b = fixture();
    const commandHost: PageSetupHost = { activeEditor: { doc: a.doc } };
    const snapshot = createSectionSnapshot(a.host);
    const otherSnapshot = createSectionSnapshot(b.host);
    await Promise.all([snapshot.read(), otherSnapshot.read()]);
    const started = deferred<void>();
    const mutation = deferred<DocReceipt>();
    a.doc.sections!.setPageMargins = () => { started.resolve(); return mutation.promise; };
    const outcome = applyPageMargins(commandHost, { topTwips: 2880 });
    await started.promise;

    a.list.mockResolvedValue({ items: [section({ margins: { top: 2, bottom: 1, left: 1, right: 1 } })] });
    commandHost.activeEditor = { doc: b.doc };
    if (failure === 'throw') mutation.reject(new Error('partial write'));
    else mutation.resolve({ success: false, failure: { code: 'FAILED', message: 'partial write' } });
    expect((await outcome).ok).toBe(false);

    expect((await readPageMargins(a.host, snapshot.read))?.topTwips).toBe(2880);
    await otherSnapshot.read();
    expect(a.list).toHaveBeenCalledTimes(3);
    expect(b.list).toHaveBeenCalledTimes(1);
    snapshot.dispose();
    otherSnapshot.dispose();
  });

  it('פקודת עיצוב בלי onUpdate מרעננת גבולות ומספור דרך אותה רשימה', async () => {
    vi.useFakeTimers();
    const { host, doc, list } = fixture();
    const snapshot = createSectionSnapshot(host);
    const borders = createPageBorderModel({ read: () => readPageBorders(host, snapshot.read), onChange: vi.fn() });
    const lines = createLineNumberingModel({ read: () => readLineNumbering(host, snapshot.read), onChange: vi.fn() });
    const stop = observeDocumentReadInvalidation(doc, () => {
      borders.noteDocumentChanged();
      lines.noteDocumentChanged();
    });
    borders.refreshNow();
    lines.refreshNow();
    await vi.advanceTimersByTimeAsync(0);
    expect(borders.getState()).not.toBeNull();
    expect(lines.getState()).not.toBeNull();

    list.mockResolvedValue({ items: [section({ pageBorders: undefined, lineNumbering: { enabled: false } })] });
    invalidateDocumentReads(doc);
    await vi.advanceTimersByTimeAsync(300);
    expect(borders.getState()).toBeNull();
    expect(lines.getState()).toBeNull();
    expect(list).toHaveBeenCalledTimes(2);
    stop();
    borders.dispose();
    lines.dispose();
    snapshot.dispose();
  });
});
