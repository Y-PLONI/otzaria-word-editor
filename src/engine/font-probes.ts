import {
  cachedFamilyAvailability,
  cachedHebrewCoverage,
  canMeasureFonts,
  coversHebrew,
  isFamilyAvailable,
} from './docx-fonts';
import type { InstalledFontsSnapshot } from './system-fonts';

const QUIET_MS = 500;
const SLICE_MS = 8;
const RETRY_MS = 100;

interface ProbeDeadline {
  timeRemaining(): number;
}

export interface FontProbeQueueOptions {
  isPaused(): boolean;
  onMeasured(): void;
  /** משפחות שהיישום מספק בעצמו אינן מועמדות לסריקה אופציונלית. */
  skip?: (family: string) => boolean;
  reportedFonts?: () => InstalledFontsSnapshot;
  canMeasure?: () => boolean;
  availability?: (family: string) => boolean | undefined;
  coverage?: (family: string) => boolean | undefined;
  measureAvailability?: (family: string) => boolean;
  measureCoverage?: (family: string) => boolean;
}

/**
 * סריקת הבורר אינה חלק מפתיחת מסמך. כל פרוסה מחכה לזמן פנוי מחדש, בלי
 * timeout שמאלץ מדידה בזמן הקלדה. מדידה בודדת היא סינכרונית ואינה ניתנת
 * לפיצול; התקציב מגביל מתי מתחילים את הבאה, לא את זמן טעינת הגופן הבודד.
 */
export function createFontProbeQueue({
  isPaused,
  onMeasured,
  skip = () => false,
  reportedFonts,
  canMeasure = canMeasureFonts,
  availability = cachedFamilyAvailability,
  coverage = cachedHebrewCoverage,
  measureAvailability = isFamilyAvailable,
  measureCoverage = coversHebrew,
}: FontProbeQueueOptions) {
  const pending = new Map<string, string>();
  const priority = new Map<string, string>();
  let idle: number | undefined;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let quietUntil = 0;
  let disposed = false;
  let composing = false;

  const hasWork = () => priority.size + pending.size > 0;
  const knowsHebrew = (name: string) => {
    const snapshot = reportedFonts?.();
    // רק דיווח המארח אינו תלוי במטמון. צילום מדידות ישן עלול לשרוד הזרקת
    // אליאס שביטלה את הכיסוי, ואסור שהוא ימנע את הבדיקה החוזרת.
    return snapshot?.source === 'host' && snapshot.hebrew.has(name.toLowerCase());
  };
  const incomplete = (name: string) => availability(name) === undefined || (
    availability(name) !== false && !knowsHebrew(name) && coverage(name) === undefined
  );
  const paused = () => composing || isPaused() || performance.now() < quietUntil;

  function cancelScheduled(): void {
    if (idle !== undefined) window.cancelIdleCallback(idle);
    if (timer !== undefined) clearTimeout(timer);
    idle = undefined;
    timer = undefined;
  }

  function schedule(): void {
    if (disposed || !hasWork() || idle !== undefined || timer !== undefined) return;
    if (paused()) {
      timer = setTimeout(() => { timer = undefined; schedule(); }, RETRY_MS);
    } else if (typeof window.requestIdleCallback === 'function') {
      idle = window.requestIdleCallback(run);
    } else {
      // WebViews בלי idle callback מקבלים פרוסה קצרה, ועדיין מכבדים הקלדה
      // ופתיחה. אין שרשרת setTimeout(0) שממלאת כל הזדמנות שהדפדפן נותן.
      timer = setTimeout(() => {
        timer = undefined;
        const started = performance.now();
        run({ timeRemaining: () => Math.max(0, SLICE_MS - (performance.now() - started)) });
      }, RETRY_MS);
    }
  }

  function run(deadline: ProbeDeadline): void {
    idle = undefined;
    if (disposed) return;
    if (paused()) { schedule(); return; }
    if (!canMeasure()) {
      pending.clear();
      priority.clear();
      return; // בלי canvas המצב נשאר לא ידוע, לא הופך ל״חסר״.
    }
    const started = performance.now();
    let measured = false;
    while (hasWork() && !paused() && deadline.timeRemaining() > 1 &&
      performance.now() - started < SLICE_MS) {
      const queue = priority.size > 0 ? priority : pending;
      const [key, name] = queue.entries().next().value!;
      queue.delete(key);
      if (!incomplete(name)) continue;
      const available = availability(name) ?? measureAvailability(name);
      measured = true;
      if (available && !knowsHebrew(name) && coverage(name) === undefined) {
        // גם זמינות וכיסוי יכולים לטעון גופן: אם נגמר הזמן, הכיסוי יחכה
        // לפרוסה הבאה באותה עדיפות.
        if (deadline.timeRemaining() <= 1 || performance.now() - started >= SLICE_MS) {
          queue.set(key, name);
          break;
        }
        measureCoverage(name);
      }
    }
    if (measured) onMeasured();
    schedule();
  }

  function enqueue(families: readonly string[], urgent = false): void {
    if (disposed) return;
    for (const family of families) {
      const name = family.trim();
      if (!name || skip(name) || !incomplete(name)) continue;
      const key = name.toLowerCase();
      if (urgent) { pending.delete(key); priority.set(key, name); }
      else if (!priority.has(key)) pending.set(key, name);
    }
    schedule();
  }

  function noteActivity(): void {
    quietUntil = performance.now() + QUIET_MS;
    cancelScheduled();
    schedule();
  }

  return {
    enqueue,
    noteActivity,
    setComposing(value: boolean) { composing = value; noteActivity(); },
    dispose() {
      disposed = true;
      cancelScheduled();
      pending.clear();
      priority.clear();
    },
  };
}
