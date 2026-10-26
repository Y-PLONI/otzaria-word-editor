/**
 * הקלדה בדיבור לתוך המסמך: כפתור אחד מתחיל הקלטה, והוא עצמו עוצר אותה.
 *
 * השירות מזרים טקסט **טיוטה** תוך כדי דיבור, והמילים האחרונות מתוקנות כשהמשפט
 * מסתיים. לכן הטקסט נכתב במקום קבוע במסמך ("הטיוטה"), וכל עדכון מחליף את
 * הטיוטה הקודמת במקומה במקום להוסיף אחריה.
 *
 * הכתיבה עוברת ישירות לעורך (ProseMirror) ולא דרך `doc.insert`: ל-Document API
 * אין החלפה של טווח שהוכנס זה עתה, והחלפה היא כל העניין כאן.
 */
import type { SuperDoc } from 'superdoc';
import { applySpokenPunctuation, isPunctuation, splitLineCommands } from './spoken-punctuation';
import {
  explainError,
  explainNotReady,
  startDictation,
  voiceStatus,
  type DictationHandle,
  type FetchStream,
} from '../host/voice-service';

/** נקודה בטקסט המסמך: פסקה והיסט בתוכה (כמו ב-Document API). */
export interface DictationPoint {
  blockId: string;
  offset: number;
  story: unknown;
}

/** מה שהכתיבה צריכה מהעורך. נפרד כדי שאפשר יהיה לבדוק בלי SuperDoc. */
export interface DictationTarget {
  /** הסמן (בלי בחירה), או `null` כשאין סמן בטקסט או שיש בחירה. סינכרוני. */
  caret(): DictationPoint | null;
  /** הטקסט שלפני הנקודה באותה פסקה, או מחרוזת ריקה כשאי אפשר לקרוא. */
  textBefore(at: DictationPoint): Promise<string>;
  /** מחליפה את הטווח [from, to) של הפסקה בטקסט ומציבה את הסמן אחריו. */
  replace(at: DictationPoint, from: number, to: number, text: string): Promise<boolean>;
  /**
   * ירידת שורה כמו Enter בנקודה `offset` של הפסקה. מחזירה את תחילת הפסקה
   * החדשה (והסמן שם), או `null` כשנכשלה.
   */
  breakLine(at: DictationPoint, offset: number): Promise<DictationPoint | null>;
}

interface SelectionTargetLike {
  start: { blockId: string; offset: number; story?: unknown };
  end: { blockId: string; offset: number };
}

interface DocLike {
  insert(input: Record<string, unknown>): Promise<{ success?: boolean }>;
  blocks?: {
    split?(input: Record<string, unknown>): Promise<{ success?: boolean; inserted?: { nodeId?: string } }>;
  };
  replace(input: Record<string, unknown>): Promise<{ success?: boolean }>;
  getNodeById(input: { nodeId: string }): Promise<{
    node?: { paragraph?: { inlines?: Array<{ run?: { text?: string } }> } };
  }>;
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
export function superdocTarget(host: SuperDoc | null): DictationTarget | null {
  const sd = host as unknown as SuperDocLike | null;
  const doc = sd?.activeEditor?.doc;
  const selection = sd?.ui?.selection;
  if (!doc || typeof doc.insert !== 'function' || typeof doc.replace !== 'function') return null;
  if (typeof selection?.getSnapshot !== 'function') return null;

  return {
    caret() {
      let target: SelectionTargetLike | undefined;
      try {
        target = selection.getSnapshot?.()?.selectionTarget;
      } catch {
        return null;
      }
      if (!target || target.start.blockId !== target.end.blockId) return null;
      if (target.start.offset !== target.end.offset) return null;
      return { blockId: target.start.blockId, offset: target.start.offset, story: target.start.story };
    },
    async textBefore(at) {
      try {
        const found = await doc.getNodeById({ nodeId: at.blockId });
        const text = (found?.node?.paragraph?.inlines ?? []).map((i) => i.run?.text ?? '').join('');
        return text.slice(Math.max(0, at.offset - 40), at.offset);
      } catch {
        return '';
      }
    },
    async replace(at, from, to, text) {
      const target = collapsedTarget(at, from, to);
      const receipt =
        from === to
          ? await doc.insert({ value: text, type: 'text', target })
          : await doc.replace({ target, text });
      // NO_OP: הטקסט זהה למה שכבר כתוב (עדכון שחזר על עצמו) — לא כשל.
      const failure = (receipt as { failure?: { code?: string } } | undefined)?.failure;
      if (receipt?.success === false && failure?.code !== 'NO_OP') {
        console.warn('[otzaria-word] כתיבת הטקסט המוכתב נדחתה', failure);
        return false;
      }
      selection.apply?.(collapsedTarget(at, from + text.length, from + text.length));
      return true;
    },
    async breakLine(at, offset) {
      const split = doc.blocks?.split;
      if (typeof split !== 'function') return null;
      const result = await split.call(doc.blocks, {
        target: { kind: 'block', nodeType: 'paragraph', nodeId: at.blockId },
        offset,
      });
      const nodeId = result?.success === false ? undefined : result?.inserted?.nodeId;
      if (!nodeId) return null;
      const next: DictationPoint = { blockId: nodeId, offset: 0, story: at.story };
      // הפסקה החדשה עוד לא פתורה בשכבת התצוגה ברגע הפיצול: מנסים שוב קצרות.
      for (let attempt = 0; attempt < 10; attempt++) {
        if (selection.apply?.(collapsedTarget(next, 0, 0))?.ok !== false) return next;
        await new Promise((resolve) => setTimeout(resolve, 50));
      }
      return null;
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
}

export interface VoiceDictation {
  state(): DictationState;
  /** לחיצה על הכפתור: מתחילה, או עוצרת אם כבר מקליטים. */
  toggle(): Promise<void>;
  /** עוצרת מיד (סגירת מסמך, יציאה). */
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

  function setState(next: DictationState): void {
    state = next;
    deps.onState(next);
  }

  async function start(): Promise<void> {
    const target = deps.target();
    if (!target) {
      deps.say('פתחו מסמך והציבו בו את הסמן לפני ההקלטה', true);
      return;
    }
    const start0 = target.caret();
    if (start0 === null) {
      deps.say('יש טקסט מסומן, או שאין סמן. לחצו במסמך כדי להציב סמן, ואז התחילו להקליט', true);
      return;
    }
    setState('starting');
    let fetchStream: FetchStream;
    try {
      fetchStream = deps.fetchStream();
      const status = await voiceStatus(fetchStream);
      const problem = explainNotReady(status);
      if (problem) {
        deps.say(problem, true);
        setState('idle');
        return;
      }
    } catch (error) {
      deps.say(explainError(error), true);
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
        if (result.reason !== 'time_limit' || stopping || result.detached) break;
        const next = target.caret();
        if (!next) break;
        origin = next;
      }
    } catch (error) {
      deps.say(explainError(error), true);
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

    /** כותבת טקסט במקום המקטע הנוכחי. `false` כשהכתיבה נכשלה. */
    const writeCurrent = async (text: string): Promise<boolean> => {
      const shaped = applySpokenPunctuation(text);
      // סימן פיסוק נצמד למה שלפניו, גם כשהוא פותח מקטע חדש
      const gap = shaped !== '' && isPunctuation(shaped[0]!) ? '' : frozen > 0 ? ' ' : lead;
      const written = shaped ? gap + shaped : '';
      if (currentLength > 0 && lastWritten === written) return true;
      if (currentLength === 0 && written === '') return true;
      const at = here.offset + frozen;
      let ok = false;
      try {
        ok = await target.replace(here, at, at + currentLength, written);
      } catch (error) {
        console.warn('[otzaria-word] כתיבת הטקסט המוכתב נכשלה', error);
      }
      if (!ok) return false;
      currentLength = written.length;
      lastWritten = written;
      return true;
    };

    const fail = (message: string): void => {
      detached = true;
      deps.say(message, true);
      void handle?.stop();
    };

    /** כותבת את העדכונים שהצטברו, מקטע אחרי מקטע, לפי הסדר. */
    const pump = async (): Promise<void> => {
      while (pending.size > 0 && !detached) {
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
          if (i > 0) {
            const next = await target.breakLine(here, here.offset + frozen + currentLength);
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
      if (detached) return;
      pending.set(segment, { text, final });
      if (!inflight) {
        inflight = pump().finally(() => {
          inflight = null;
          if (pending.size > 0 && !detached) {
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
        onReady: () => setState('listening'),
        onSegment,
        onError: (message) => deps.say('זיהוי הדיבור נכשל: ' + message, true),
      },
    );
    const result = await handle.done;
    while (inflight) await inflight; // העדכון האחרון נכתב לפני שמסיימים
    return { reason: result.reason, detached };
  }

  return {
    state: () => state,
    async toggle() {
      if (state === 'idle') {
        await start();
      } else if (handle) {
        stopping = true;
        await handle.stop();
      }
    },
    cancel() {
      handle?.cancel();
    },
  };
}
