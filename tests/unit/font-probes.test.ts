import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createFontProbeQueue } from '../../src/engine/font-probes';
import { emptyInstalledFonts, type InstalledFontsSnapshot } from '../../src/engine/system-fonts';

function harness(reportedFonts?: () => InstalledFontsSnapshot) {
  const callbacks = new Map<number, IdleRequestCallback>();
  let serial = 0;
  let now = 0;
  let paused = false;
  const available = new Map<string, boolean>();
  const hebrew = new Map<string, boolean>();
  const measured: string[] = [];
  const idle = vi.fn((callback: IdleRequestCallback) => {
    callbacks.set(++serial, callback);
    return serial;
  });
  vi.stubGlobal('requestIdleCallback', idle);
  vi.stubGlobal('cancelIdleCallback', (id: number) => callbacks.delete(id));
  vi.spyOn(performance, 'now').mockImplementation(() => now);
  const onMeasured = vi.fn();
  const key = (name: string) => name.toLowerCase();
  const queue = createFontProbeQueue({
    isPaused: () => paused,
    reportedFonts,
    canMeasure: () => true,
    availability: (name) => available.get(key(name)),
    coverage: (name) => hebrew.get(key(name)),
    measureAvailability: (name) => {
      measured.push(name);
      now += 8; // גופן בודד עשוי לצרוך את כל הפרוסה.
      available.set(key(name), true);
      return true;
    },
    measureCoverage: (name) => {
      measured.push(`${name}:hebrew`);
      now += 8;
      hebrew.set(key(name), true);
      return true;
    },
    onMeasured,
  });
  return {
    queue, measured, onMeasured, idle, callbacks, available, hebrew,
    pause(value: boolean) { paused = value; },
    advance(ms: number) { now += ms; vi.advanceTimersByTime(ms); },
    runIdle(budget = 50) {
      const entry = callbacks.entries().next().value;
      expect(entry).toBeDefined();
      const [id, callback] = entry!;
      callbacks.delete(id);
      callback({ didTimeout: false, timeRemaining: () => budget });
    },
  };
}

beforeEach(() => vi.useFakeTimers());
afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe('optional font probes', () => {
  it('rechecks Hebrew coverage after invalidation even while a measured snapshot still lists it', () => {
    let snapshot: InstalledFontsSnapshot = {
      ...emptyInstalledFonts(), source: 'measured', hebrew: new Set(['narkisim']),
    };
    const h = harness(() => snapshot);
    h.available.set('narkisim', true);
    h.hebrew.set('narkisim', true);
    h.queue.enqueue(['Narkisim']);
    expect(h.callbacks.size).toBe(0);

    // הזרקת אליאס מבטלת את המטמון; צילום הרשימה עדיין מחזיק בתוצאה הישנה.
    h.available.delete('narkisim');
    h.hebrew.delete('narkisim');
    h.onMeasured.mockImplementation(() => {
      snapshot = { ...snapshot, hebrew: new Set([...h.hebrew].filter(([, value]) => value).map(([name]) => name)) };
    });
    h.queue.enqueue(['Narkisim']);
    h.runIdle(); // בדיקת הזמינות צורכת את הפרוסה; הכיסוי צריך להישאר בתור.
    expect(h.measured).toEqual(['Narkisim']);
    expect(snapshot.hebrew.has('narkisim')).toBe(false);
    expect(h.callbacks.size).toBe(1);
    h.runIdle();
    expect(h.measured).toEqual(['Narkisim', 'Narkisim:hebrew']);
    expect(snapshot.hebrew.has('narkisim')).toBe(true);
    expect(h.callbacks.size).toBe(0);
    h.queue.dispose();
  });

  it('accepts authoritative Hebrew metadata from the host without probing its coverage', () => {
    const h = harness(() => ({ ...emptyInstalledFonts(), source: 'host', hebrew: new Set(['narkisim']) }));
    h.queue.enqueue(['Narkisim']);
    h.runIdle();
    expect(h.measured).toEqual(['Narkisim']);
    expect(h.callbacks.size).toBe(0);
    h.queue.dispose();
  });

  it('does no synchronous probes and requests a fresh idle period after each slice', () => {
    const h = harness();
    h.queue.enqueue(Array.from({ length: 300 }, (_, i) => `Family ${i}`));
    expect(h.measured).toEqual([]);
    expect(h.idle).toHaveBeenCalledTimes(1);
    expect(h.idle.mock.calls[0]).toHaveLength(1); // אין timeout שיכפה עבודה.
    h.runIdle();
    expect(h.measured).toEqual(['Family 0']);
    expect(h.callbacks.size).toBe(1);
    expect(vi.getTimerCount()).toBe(0);
    h.queue.dispose();
  });

  it('waits for a useful budget and promotes selected/document/visible names without duplicates', () => {
    const h = harness();
    h.queue.enqueue(['Background', 'Visible', 'Document']);
    h.queue.enqueue(['Document', 'Selected', ' visible '], true);
    h.runIdle(0);
    expect(h.measured).toEqual([]);
    h.runIdle();
    h.runIdle();
    h.runIdle();
    expect(h.measured).toEqual(['Document', 'Selected', 'visible']);
    for (let i = 0; i < 5; i += 1) h.runIdle();
    expect(h.measured.filter((name) => !name.endsWith(':hebrew')))
      .toEqual(['Document', 'Selected', 'visible', 'Background']);
    expect(h.callbacks.size).toBe(0);
    h.queue.enqueue(['DOCUMENT', 'Background']);
    expect(h.callbacks.size).toBe(0);
    h.queue.dispose();
  });

  it('pauses an already scheduled probe during opening, typing and IME composition', () => {
    const h = harness();
    h.queue.enqueue(['Document']);
    h.pause(true);
    h.runIdle();
    h.advance(2000);
    expect(h.measured).toEqual([]);
    h.pause(false);
    h.advance(100);
    expect(h.callbacks.size).toBe(1);
    for (let i = 0; i < 10; i += 1) {
      h.queue.noteActivity();
      h.advance(200);
      expect(h.callbacks.size).toBe(0);
    }
    h.queue.setComposing(true);
    h.advance(2000);
    expect(h.callbacks.size).toBe(0);
    h.queue.setComposing(false);
    h.advance(500);
    h.runIdle();
    expect(h.measured).toEqual(['Document']);
    h.queue.dispose();
  });

  it('stops pending callbacks and ignores new work after disposal', () => {
    const h = harness();
    h.queue.enqueue(['A']);
    const stale = [...h.callbacks.values()][0];
    h.queue.dispose();
    stale({ didTimeout: false, timeRemaining: () => 50 });
    h.queue.enqueue(['B'], true);
    h.advance(2000);
    expect(h.callbacks.size).toBe(0);
    expect(h.measured).toEqual([]);
    expect(h.onMeasured).not.toHaveBeenCalled();
  });

  it('keeps unknown state without canvas and uses a quiet timer fallback without idle support', () => {
    const h = harness();
    const unavailableCanvas = createFontProbeQueue({
      isPaused: () => false, onMeasured: h.onMeasured, canMeasure: () => false,
      measureAvailability: () => { throw new Error('must not measure'); },
    });
    unavailableCanvas.enqueue(['No canvas']);
    h.runIdle();
    expect(h.onMeasured).not.toHaveBeenCalled();
    expect(h.callbacks.size).toBe(0);
    unavailableCanvas.dispose();
    vi.stubGlobal('requestIdleCallback', undefined);
    h.queue.enqueue(['Fallback']);
    h.queue.noteActivity();
    h.advance(400);
    expect(h.measured).toEqual([]);
    h.advance(200);
    expect(h.measured).toEqual(['Fallback']);
    h.queue.dispose();
  });
});
