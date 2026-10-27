/**
 * הקלדה בדיבור לתוך המסמך: כפתור אחד מתחיל הקלטה, והוא עצמו עוצר אותה.
 *
 * השירות מזרים טקסט **טיוטה** תוך כדי דיבור, והמילים האחרונות מתוקנות כשהמשפט
 * מסתיים. לכן הטקסט נכתב במקום קבוע במסמך ("הטיוטה"), וכל עדכון מחליף את
 * הטיוטה הקודמת במקומה במקום להוסיף אחריה.
 *
 * הכתיבה עוברת דרך ה-Document API, עם בדיקת הטקסט הקודם לפני החלפת הטיוטה.
 */
import type { SuperDoc } from 'superdoc';
import { applySpokenPunctuation, isPunctuation, splitLineCommands } from './spoken-punctuation';
import {
  explainError,
  explainNotReady,
  startDictation,
  voiceStatus,
  VoiceServiceError,
  type DictationHandle,
  type FetchStream,
} from '../host/voice-service';

/** נקודה בטקסט המסמך: פסקה והיסט בתוכה (כמו ב-Document API). */
export interface DictationPoint {
  blockId: string;
  offset: number;
  story: unknown;
}

export interface DictationMutationGuard {
  /** ביטול בזמן המתנה אינו מאפשר להמשיך לכתוב או להזיז את הסמן. */
  isActive(): boolean;
  /** החלפה מותרת רק אם הטווח עדיין מכיל את הטיוטה הקודמת. */
  expectedText?: string;
}

/** מה שהכתיבה צריכה מהעורך. נפרד כדי שאפשר יהיה לבדוק בלי SuperDoc. */
export interface DictationTarget {
  /** הסמן (בלי בחירה), או `null` כשאין סמן בטקסט או שיש בחירה. סינכרוני. */
  caret(): DictationPoint | null;
  /** הטקסט שלפני הנקודה באותה פסקה, או מחרוזת ריקה כשאי אפשר לקרוא. */
  textBefore(at: DictationPoint): Promise<string>;
  /** מחליפה את הטווח [from, to) של הפסקה בטקסט ומציבה את הסמן אחריו. */
  replace(at: DictationPoint, from: number, to: number, text: string, guard?: DictationMutationGuard): Promise<boolean>;
  /**
   * ירידת שורה כמו Enter בנקודה `offset` של הפסקה. מחזירה את תחילת הפסקה
   * החדשה (והסמן שם), או `null` כשנכשלה.
   */
  breakLine(at: DictationPoint, offset: number, guard?: DictationMutationGuard): Promise<DictationPoint | null>;
  /** קלט משתמש מבטל הכתבה עוד לפני שהקלט משנה את המסמך או את הסמן. */
  onUserInteraction?(listener: () => void): () => void;
}

interface SelectionTargetLike {
  kind?: string;
  coordinateSpace?: string;
  start: { blockId: string; offset: number; story?: unknown };
  end: { blockId: string; offset: number };
  story?: unknown;
}

interface DocLike {
  insert(input: Record<string, unknown>): Promise<{ success?: boolean }>;
  blocks?: {
    split?(input: Record<string, unknown>): Promise<{ success?: boolean; inserted?: { nodeId?: string } }>;
  };
  replace(input: Record<string, unknown>): Promise<{ success?: boolean }>;
  ranges?: {
    resolve(input: Record<string, unknown>): Promise<{
      preview?: { text?: string; truncated?: boolean };
    }>;
  };
}

interface SuperDocLike {
  ui?: {
    selection?: {
      getSnapshot?: () => { selectionTarget?: SelectionTargetLike } | null;
      apply?: (target: unknown) => { ok?: boolean } | undefined;
    };
  };
  activeEditor?: { doc?: DocLike };
}

function collapsedTarget(at: DictationPoint, from: number, to: number): Record<string, unknown> {
  return {
    kind: 'selection',
    start: { kind: 'text', blockId: at.blockId, offset: from, story: at.story },
    end: { kind: 'text', blockId: at.blockId, offset: to, story: at.story },
    story: at.story,
  };
}

/**
 * העורך הפעיל של המסמך, או `null` כשאין מסמך שאפשר לכתוב בו.
 *
 * הכתיבה עוברת דרך ה-Document API (`doc.insert` / `doc.replace` עם יעד של
 * פסקה והיסט), כמו שאר התוסף: ל-SuperDoc 2 אין חשיפה של ProseMirror. אחרי כל
 * כתיבה הסמן מוזז לסוף הטקסט שנכתב (`ui.selection.apply`) — ה-API עצמו משאיר
 * אותו במקומו — וכך הסמן עוקב אחרי הדיבור, וזז רק אם המשתמש הזיז אותו.
 */
export function superdocTarget(host: SuperDoc | null, container?: HTMLElement | null): DictationTarget | null {
  const sd = host as unknown as SuperDocLike | null;
  const doc = sd?.activeEditor?.doc;
  const selection = sd?.ui?.selection;
  if (!doc || typeof doc.insert !== 'function' || typeof doc.replace !== 'function') return null;
  if (typeof selection?.getSnapshot !== 'function') return null;

  async function placeCaret(at: DictationPoint, offset: number, guard?: DictationMutationGuard): Promise<boolean> {
    // פריסת הטקסט עשויה להגיע אחרי קבלת אישור הכתיבה. גם הכנסה רגילה, ולא
    // רק פיצול פסקה, יכולה להחזיר זמנית ok:false בעת הצבת הסמן החדש.
    for (let attempt = 0; attempt < 10; attempt++) {
      if (guard && !guard.isActive()) return false;
      if (selection?.apply?.(collapsedTarget(at, offset, offset))?.ok !== false) return true;
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
    return false;
  }

  return {
    caret() {
      let target: SelectionTargetLike | undefined;
      try {
        target = selection.getSnapshot?.()?.selectionTarget;
      } catch {
        return null;
      }
      if (!target || target.kind !== 'selection') return null;
      if (target.coordinateSpace && target.coordinateSpace !== 'visible') return null;
      if (typeof target.start?.blockId !== 'string' || target.start.blockId !== target.end?.blockId) return null;
      if (typeof target.start.offset !== 'number') return null;
      if (target.start.offset !== target.end.offset) return null;
      return { blockId: target.start.blockId, offset: target.start.offset, story: target.story ?? target.start.story };
    },
    async textBefore(at) {
      try {
        if (at.offset === 0) return '';
        const from = Math.max(0, at.offset - 40);
        // חלון קטן במודל הטקסט של המנוע: כולל כותרות, קישורים ותווים מיוחדים,
        // בלי להוריד פסקה שלמה ובלי לשחזר ידנית את סרגל ההיסטים שלה.
        const target = collapsedTarget(at, from, at.offset);
        const resolved = await doc.ranges?.resolve({
          start: { kind: 'point', point: target.start },
          end: { kind: 'point', point: target.end },
          ...(at.story ? { in: at.story } : {}),
        });
        const text = resolved?.preview?.text;
        return typeof text === 'string' && !resolved?.preview?.truncated && text.length === at.offset - from
          ? text
          : '';
      } catch {
        return '';
      }
    },
    async replace(at, from, to, text, guard) {
      if (guard && !guard.isActive()) return false;
      const target = collapsedTarget(at, from, to);
      const receipt =
        from === to
          ? await doc.insert({ value: text, type: 'text', target })
          : await doc.replace({ target, text, ...(guard ? { expectedText: guard.expectedText } : {}) });
      if (guard && !guard.isActive()) return false;
      // NO_OP: הטקסט זהה למה שכבר כתוב (עדכון שחזר על עצמו) — לא כשל.
      const failure = (receipt as { failure?: { code?: string } } | undefined)?.failure;
      if (receipt?.success === false && failure?.code !== 'NO_OP') {
        console.warn('[otzaria-word] כתיבת הטקסט המוכתב נדחתה', failure);
        return false;
      }
      return placeCaret(at, from + text.length, guard);
    },
    async breakLine(at, offset, guard) {
      if (guard && !guard.isActive()) return null;
      const split = doc.blocks?.split;
      if (typeof split !== 'function') return null;
      const result = await split.call(doc.blocks, {
        target: { kind: 'block', nodeType: 'paragraph', nodeId: at.blockId },
        offset,
      });
      if (guard && !guard.isActive()) return null;
      const nodeId = result?.success === false ? undefined : result?.inserted?.nodeId;
      if (!nodeId) return null;
      const next: DictationPoint = { blockId: nodeId, offset: 0, story: at.story };
      return await placeCaret(next, 0, guard) ? next : null;
    },
    onUserInteraction(listener) {
      // אירועי קלט בדפדפן; כתיבה דרך ה-API ו-selection.apply אינן
      // מבטלות את עצמן. capture מקדים את הטיפול של המנוע בקלט המשתמש.
      const events = ['beforeinput', 'keydown', 'pointerdown'] as const;
      for (const event of events) container?.addEventListener(event, listener, true);
      return () => {
        for (const event of events) container?.removeEventListener(event, listener, true);
      };
    },
  };
}

export type DictationState = 'idle' | 'starting' | 'listening';

export interface VoiceDictationDeps {
  fetchStream: () => FetchStream;
  target: () => DictationTarget | null;
  /** הודעה למשתמש (שורת המצב). */
  say(text: string, isError?: boolean): void;
  onState(state: DictationState): void;
  /**
   * `false` כשהמסמך לקריאה בלבד: ה-Document API דוחה בו כתיבה, ופתיחת
   * המיקרופון הייתה נגמרת ב„הכתיבה למסמך נכשלה”. בלעדיו — תמיד אפשר.
   */
  canWrite?: () => boolean;
  /**
   * השירות לא ענה בבדיקה שלפני ההקלטה, כלומר כנראה אינו מותקן. כשהוא קיים,
   * הוא מקבל את המקרה במקום הודעת שורת המצב (במעטפת: חלון עם קישור להורדה).
   */
  onServiceMissing?: () => void;
}

export interface VoiceDictation {
  state(): DictationState;
  /**
   * לחיצה על הכפתור: מתחילה; בזמן ההתחלה — מבטלת; בזמן הקלטה — עוצרת אחרי
   * המשפט האחרון.
   */
  toggle(): Promise<void>;
  /** עוצרת מיד (סגירת מסמך, מעבר למסמך אחר, יציאה), גם בזמן ההתחלה. */
  cancel(): void;
}

interface Pending {
  text: string;
  final: boolean;
}

function firstPending(map: Map<number, Pending>): [number, Pending] | undefined {
  const key = Math.min(...map.keys());
  const item = map.get(key);
  return item === undefined ? undefined : [key, item];
}

export function createVoiceDictation(deps: VoiceDictationDeps): VoiceDictation {
  let state: DictationState = 'idle';
  let handle: DictationHandle | null = null;
  let stopping = false;
  // ביטול שהגיע לפני שההקלטה עצמה התחילה — בזמן בדיקת השירות או קריאת
  // הטקסט שלפני הסמן, כשעוד אין `handle` לבטל.
  let aborted = false;
  let cancelSession: (() => void) | null = null;
  let removeInteraction: (() => void) | null = null;

  function cancel(): void {
    aborted = true;
    cancelSession?.();
    removeInteraction?.();
    removeInteraction = null;
    handle?.cancel();
  }

  function setState(next: DictationState): void {
    if (next === 'idle') {
      removeInteraction?.();
      removeInteraction = null;
    }
    state = next;
    deps.onState(next);
  }

  async function start(): Promise<void> {
    const target = deps.target();
    if (!target) {
      deps.say('פתחו מסמך והציבו בו את הסמן לפני ההקלטה', true);
      return;
    }
    if (deps.canWrite && !deps.canWrite()) {
      deps.say('המסמך פתוח לקריאה בלבד, ולכן אי אפשר להקליט לתוכו', true);
      return;
    }
    const start0 = target.caret();
    if (start0 === null) {
      deps.say('יש טקסט מסומן, או שאין סמן. לחצו במסמך כדי להציב סמן, ואז התחילו להקליט', true);
      return;
    }
    aborted = false;
    setState('starting');
    removeInteraction = target.onUserInteraction?.(cancel) ?? null;
    let fetchStream: FetchStream;
    try {
      fetchStream = deps.fetchStream();
      const status = await voiceStatus(fetchStream);
      if (aborted) {
        setState('idle');
        return;
      }
      const problem = explainNotReady(status);
      if (problem) {
        deps.say(problem, true);
        setState('idle');
        return;
      }
    } catch (error) {
      // המשתמש כבר ביטל: שירות שנכשל אחר כך אינו חדשה בשבילו.
      if (!aborted) {
        const missing = error instanceof VoiceServiceError && error.code === 'unreachable';
        if (missing && deps.onServiceMissing) deps.onServiceMissing();
        else deps.say(explainError(error), true);
      }
      setState('idle');
      return;
    }

    // ההקלטה נמשכת עד שהמשתמש עוצר אותה. השירות מקליט עד 100 שניות בכל פעם,
    // ולכן בסיומן ממשיכים בהקלטה חדשה במקום שבו הסמן עומד.
    stopping = false;
    let origin = start0;
    try {
      while (true) {
        const result = await runSession(fetchStream, target, origin);
        if (result.reason === 'no_speech') deps.say('לא נשמע דיבור');
        if (result.reason !== 'time_limit' || stopping || aborted || result.detached) break;
        const next = target.caret();
        if (!next) break;
        origin = next;
      }
    } catch (error) {
      if (!aborted) deps.say(explainError(error), true);
    } finally {
      handle = null;
      setState('idle');
    }
  }

  /**
   * הקלטה אחת מול השירות. הטיוטה מתחילה בסמן, ואורכה הנוכחי. הטקסט מהשירות
   * מצטבר מתחילת ההקלטה, ולכן כשהסמן זז (המשתמש לחץ במקום אחר) ההקלטה
   * מסתיימת: מה שנכתב נשאר כמות שהוא, ושום דבר לא נכתב במקום שהמשתמש לא ביקש.
   */
  async function runSession(
    fetchStream: FetchStream,
    target: DictationTarget,
    origin: DictationPoint,
  ): Promise<{ reason: string; detached: boolean }> {
    const before = await target.textBefore(origin);
    // ביטול, או עצירה שנלחצה בדיוק במעבר בין שתי הקלטות של 100 שניות — אז
    // ה-`stop` הגיע להקלטה שכבר הסתיימה, וההקלטה הזו עוד לא התחילה.
    if (aborted || stopping) return { reason: 'cancelled', detached: true };
    let lead = before !== '' && !/\s$/.test(before) ? ' ' : '';
    // מקטעים שנסגרו אינם משתנים עוד: רק המקטע הנוכחי מוחלף בכל עדכון, ולכן
    // עלות הכתיבה אינה גדלה עם אורך ההכתבה. `frozen` הוא אורך מה שנסגר בפסקה
    // הנוכחית; ירידת שורה פותחת פסקה חדשה, ואז הספירה מתחילה מחדש.
    let here = origin;
    let frozen = 0;
    let current = -1;
    let currentLength = 0;
    let lastWritten = '';
    let detached = false;
    const pending = new Map<number, Pending>();
    const broken = new Set<number>(); // מקטעים שירידות השורה שלהם כבר בוצעו
    let inflight: Promise<void> | null = null;
    const isActive = (): boolean => !aborted && !detached;
    cancelSession = () => {
      detached = true;
      pending.clear();
    };

    /** כותבת טקסט במקום המקטע הנוכחי. `false` כשהכתיבה נכשלה. */
    const writeCurrent = async (text: string): Promise<boolean> => {
      if (!isActive()) return false;
      const shaped = applySpokenPunctuation(text);
      // סימן פיסוק נצמד למה שלפניו, גם כשהוא פותח מקטע חדש
      const gap = shaped !== '' && isPunctuation(shaped[0]!) ? '' : frozen > 0 ? ' ' : lead;
      const written = shaped ? gap + shaped : '';
      if (currentLength > 0 && lastWritten === written) return true;
      if (currentLength === 0 && written === '') return true;
      const at = here.offset + frozen;
      let ok = false;
      try {
        ok = await target.replace(here, at, at + currentLength, written, { expectedText: lastWritten, isActive });
      } catch (error) {
        console.warn('[otzaria-word] כתיבת הטקסט המוכתב נכשלה', error);
      }
      if (!isActive() || !ok) return false;
      currentLength = written.length;
      lastWritten = written;
      return true;
    };

    const fail = (message: string): void => {
      if (!isActive()) return;
      detached = true;
      deps.say(message, true);
      void handle?.stop();
    };

    /** כותבת את העדכונים שהצטברו, מקטע אחרי מקטע, לפי הסדר. */
    const pump = async (): Promise<void> => {
      while (pending.size > 0 && isActive()) {
        const segment = Math.min(...pending.keys());
        const { text, final } = pending.get(segment) as Pending;
        pending.delete(segment);
        if (segment < current || broken.has(segment)) continue; // מקטע שכבר נסגר
        if (segment > current) {
          frozen += currentLength;
          current = segment;
          currentLength = 0;
          lastWritten = '';
        }
        const caret = target.caret();
        if (!caret || caret.blockId !== here.blockId || caret.offset !== here.offset + frozen + currentLength) {
          detached = true;
          void handle?.stop();
          return;
        }
        const pieces = splitLineCommands(text);
        // בטיוטה אין ירידת שורה: הפקודה נעלמת, והירידה מתבצעת כשהמקטע נסגר
        if (!final || pieces.length === 1) {
          if (!(await writeCurrent(pieces.join(' ')))) return fail('הכתיבה למסמך נכשלה, ההקלטה הופסקה');
          continue;
        }
        broken.add(segment);
        for (let i = 0; i < pieces.length; i++) {
          if (!isActive()) return;
          if (i > 0) {
            const next = await target.breakLine(here, here.offset + frozen + currentLength, { isActive });
            if (!isActive()) return;
            if (!next) return fail('ירידת השורה במסמך נכשלה, ההקלטה הופסקה');
            here = next;
            frozen = 0;
            currentLength = 0;
            lastWritten = '';
            lead = '';
          }
          if (!(await writeCurrent(pieces[i]!))) return fail('הכתיבה למסמך נכשלה, ההקלטה הופסקה');
        }
      }
    };

    const onSegment = (segment: number, text: string, final: boolean): void => {
      if (!isActive()) return;
      pending.set(segment, { text, final });
      if (!inflight) {
        inflight = pump().catch(() => fail('הכתיבה למסמך נכשלה, ההקלטה הופסקה')).finally(() => {
          inflight = null;
          if (pending.size > 0 && isActive()) {
            const next = firstPending(pending);
            if (next) onSegment(next[0], next[1].text, next[1].final);
          }
        });
      }
    };

    handle = startDictation(
      fetchStream,
      { before, continuous: true },
      {
        onReady: () => { if (isActive()) setState('listening'); },
        onSegment,
        onError: (message) => { if (isActive()) deps.say('זיהוי הדיבור נכשל: ' + message, true); },
      },
    );
    try {
      const result = await handle.done;
      while (inflight) await inflight; // העדכון האחרון נכתב לפני שמסיימים
      return { reason: result.reason, detached };
    } finally {
      // גם זרם שנכשל אינו משאיר כותב ישן שימשיך לתוך ההקלטה הבאה.
      cancelSession();
      while (inflight) await inflight;
      cancelSession = null;
    }
  }

  return {
    state: () => state,
    async toggle() {
      if (state === 'idle') {
        await start();
      } else if (state === 'starting') {
        // המיקרופון עוד לא נפתח: אין מה לסיים, רק לבטל.
        cancel();
      } else if (handle) {
        stopping = true;
        await handle.stop();
      }
    },
    cancel,
  };
}
