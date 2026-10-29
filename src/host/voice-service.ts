/**
 * הלקוח של שירות ההקלטה והזיהוי של אוצריא ("אוצריא דיבור").
 *
 * השירות מאזין ב-`127.0.0.1:47811` בלבד, מקליט מהמיקרופון ומזרים חזרה טקסט
 * בזמן שהמשתמש מדבר. הוא עונה רק לתוספים של אוצריא: בקשה שמגיעה מדף אינטרנט
 * נושאת `Origin` ונדחית, והבקשה כאן עוברת דרך `network.fetchStream` של ה-SDK
 * (ההרשאה `network.localhost` והכתובת ב-`network.allowlist` של המניפסט).
 *
 * המודול הזה לא נוגע במסמך — הוא רק מתרגם בין השירות לאירועים. הכתיבה למסמך
 * נעשית ב-`engine/voice-dictation.ts`.
 */
import type { NetworkFetchStreamChunk } from '../types/otzaria_plugin';

export const VOICE_SERVICE_URL = 'http://127.0.0.1:47811';

/** דף ההורדה של השירות: בוחרים בו את המתקין למערכת שלכם. */
export const VOICE_SERVICE_DOWNLOAD_URL = 'https://github.com/yakov100/otzaria-voice/releases/latest';

const HEADERS = { 'X-Otzaria-Voice': '1', 'Content-Type': 'application/json' };

/** הדיבור מסתיים מעצמו אחרי משפט שנגמר ושקט קצר, או כשלא נאמר דבר. */
const END_AFTER_SILENCE_MS = 1500;
const NO_SPEECH_MS = 8000;

/** מה שהשירות מדווח על עצמו ב-`GET /status` — רק מה שהתוסף צריך. */
export interface VoiceStatus {
  service: string;
  pack: boolean;
  packJob?: { state?: string } | null;
  runtime: boolean;
  microphone?: boolean;
}

/** אירוע בשורה אחת בגוף התשובה של `/dictate`. */
export type DictationEvent =
  | { type: 'ready' }
  | { type: 'update'; segment: number; text: string; final?: boolean }
  | { type: 'error'; message: string }
  | { type: 'end'; reason: string };

/** סיבת הסיום כפי שהשירות מדווח, או `error` כשהשיחה כשלה. */
export type DictationEndReason =
  | 'stopped'
  | 'silence'
  | 'no_speech'
  | 'time_limit'
  | 'cancelled'
  | 'error'
  | string;

/** כשל בשיחה עם השירות, עם קוד שמאפשר להסביר למשתמש מה לעשות. */
export class VoiceServiceError extends Error {
  constructor(
    readonly code: string,
    message: string,
  ) {
    super(message);
    this.name = 'VoiceServiceError';
  }
}

/** מה שהלקוח צריך מהגשר — מוזרק כדי שאפשר יהיה לבדוק בלי אוצריא. */
export type FetchStream = (request: {
  url: string;
  method: string;
  headers: Record<string, string>;
  body: string;
  timeoutMs: number;
}) => AsyncIterable<NetworkFetchStreamChunk>;

/** ה-SDK מחזיר את הזרם ישירות מ-`call`, בלי להמתין לו. */
export function otzariaFetchStream(): FetchStream {
  const sdk = (window as Partial<Window>).Otzaria;
  if (!sdk) throw new VoiceServiceError('no_host', 'ההקלטה זמינה רק כשהעורך פועל בתוך אוצריא');
  return (request) =>
    (sdk.call as unknown as (m: string, p: unknown) => AsyncIterable<NetworkFetchStreamChunk>)(
      'network.fetchStream',
      request,
    );
}

/**
 * כשל מהגשר של אוצריא. הרשאה חסרה או כתובת שנחסמה אינם "השירות לא פועל" —
 * אוצריא מסבירה בעצמה מה לעשות (למשל להפעיל את ההרשאה בהגדרות התוספים),
 * וההודעה שלה היא שמגיעה למשתמש. ה-SDK עשוי להעביר רק message, בלי code:
 * כשל לא מוכר נשמר כפי שהוא; רק שגיאת חיבור מזוהה מפנה להתקנת השירות.
 */
function hostFailure(error: unknown): VoiceServiceError {
  const message = String((error as Error)?.message ?? error);
  const suppliedCode = (error as { code?: unknown } | null)?.code;
  const code = typeof suppliedCode === 'string'
    ? suppliedCode.replace(/^error\./, '')
    : /^(?:Exception: )?error\.([a-z_]+):/.exec(message)?.[1];
  if (code) return new VoiceServiceError(code, message);
  const connectionFailure = /connection (?:refused|reset|closed|aborted)|failed host lookup|socketexception|econnrefused|econnreset|enetunreach|ehostunreach|failed to fetch|networkerror/i.test(message);
  return new VoiceServiceError(connectionFailure ? 'unreachable' : 'host_error', message);
}

/**
 * שיחה אחת עם השירות: מחזירה את השורות של גוף התשובה כאובייקטי JSON, ככל
 * שהן מגיעות. גבולות המקטעים אינם גבולות שורה, ולכן שורה חלקית נשמרת לסבב הבא.
 */
async function* serviceLines(
  fetchStream: FetchStream,
  path: string,
  method: string,
  body: string,
  timeoutMs: number,
  onIterator?: (it: AsyncIterator<NetworkFetchStreamChunk>) => void,
): AsyncGenerator<Record<string, unknown>> {
  let it: AsyncIterator<NetworkFetchStreamChunk>;
  try {
    it = fetchStream({
      url: VOICE_SERVICE_URL + path,
      method,
      headers: HEADERS,
      body,
      timeoutMs,
    })[Symbol.asyncIterator]();
  } catch (error) {
    throw hostFailure(error);
  }
  onIterator?.(it);

  let pending = '';
  let status = 200;
  let errorBody = '';
  while (true) {
    let step: IteratorResult<NetworkFetchStreamChunk>;
    try {
      step = await it.next();
    } catch (error) {
      throw hostFailure(error);
    }
    if (step.done) break;
    const chunk = step.value;
    if (chunk.type === 'response') {
      status = chunk.status;
      continue;
    }
    if (status !== 200) {
      errorBody += chunk.body;
      continue;
    }
    pending += chunk.body;
    const lines = pending.split('\n');
    pending = lines.pop() ?? '';
    for (const line of lines) if (line.trim()) yield JSON.parse(line) as Record<string, unknown>;
  }
  if (status !== 200) {
    let err: { error?: string; message?: string } = {};
    try {
      err = JSON.parse(errorBody) as typeof err;
    } catch {
      // לא JSON — נשארים עם קוד ה-HTTP
    }
    throw new VoiceServiceError(err.error ?? `http_${status}`, err.message ?? `HTTP ${status}`);
  }
  if (pending.trim()) yield JSON.parse(pending) as Record<string, unknown>;
}

async function lastOf(
  fetchStream: FetchStream,
  path: string,
  method: string,
): Promise<Record<string, unknown> | null> {
  let last: Record<string, unknown> | null = null;
  for await (const obj of serviceLines(fetchStream, path, method, '', 5000)) last = obj;
  return last;
}

export async function voiceStatus(fetchStream: FetchStream): Promise<VoiceStatus> {
  const st = (await lastOf(fetchStream, '/status', 'GET')) as VoiceStatus | null;
  if (!st || st.service !== 'otzaria-voice') {
    throw new VoiceServiceError('unreachable', 'שירות ההקלטה לא ענה');
  }
  return st;
}

/** הסבר בעברית למה אי אפשר להתחיל — או `null` כשהכול מוכן. */
export function explainNotReady(st: VoiceStatus): string | null {
  if (!st.runtime) return 'שירות ההקלטה לא מצא את רכיב ONNX Runtime של אוצריא';
  const busy = st.packJob?.state;
  if (!st.pack || busy === 'downloading' || busy === 'verifying' || busy === 'extracting') {
    return 'חבילת הדיבור עדיין לא מותקנת. התקינו אותה מכפתור „חיפוש בדיבור” בסרגל של הקורא';
  }
  if (st.microphone === false) return 'אין גישה למיקרופון';
  return null;
}

/** ההסבר למשתמש על כשל שנזרק מכאן. */
export function explainError(error: unknown): string {
  if (error instanceof VoiceServiceError) {
    if (error.code === 'unreachable') {
      // הפרטים נשארים בהודעה: "לא פועל" הוא ניחוש, וההודעה המקורית היא העובדה.
      return (
        'שירות הדיבור של אוצריא אינו פועל. זו תוכנה נפרדת שמתקינים פעם אחת, והתוסף „חיפוש בדיבור” מפנה להורדה שלה' +
        (error.message ? ` (פרטים: ${error.message})` : '')
      );
    }
    if (error.code === 'permission_denied' || error.code === 'forbidden') {
      return error.message.replace(/^error\.[a-z_]+:\s*/, '');
    }
    if (error.code === 'busy') return 'ההקלטה כבר פועלת, בחלון אחר';
    return error.message;
  }
  return String((error as Error)?.message ?? error);
}

/** ההקשר שעוזר למודל השפה להמשיך: שתי המילים האחרונות שלפני הסמן. */
export function contextWords(before: string): string[] {
  return before
    .split(/\s+/)
    .filter((w) => w)
    .slice(-2);
}

/** הטקסט של כל המקטעים לפי הסדר. */
export function composeSegments(segments: Record<number, string>): string {
  return Object.keys(segments)
    .map(Number)
    .sort((a, b) => a - b)
    .map((k) => segments[k])
    .filter((t) => t)
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export interface DictationHooks {
  onReady?: () => void;
  /** הטקסט המצטבר של ההקלטה עד כה; `final` כשהמקטע האחרון נסגר. */
  onText?: (text: string, final: boolean) => void;
  /** העדכון של מקטע אחד (לפני הצירוף); מקטע קודם אינו משתנה אחרי שהתחיל מקטע חדש. */
  onSegment?: (segment: number, text: string, final: boolean) => void;
  onError?: (message: string) => void;
}

export interface DictationHandle {
  /** מסיימת אחרי שהמשפט האחרון זוהה. */
  stop(): Promise<void>;
  /** מפסיקה מיד, בלי להמתין לטקסט. */
  cancel(): void;
  /** מסתיימת עם הטקסט הסופי וסיבת הסיום; נכשלת כשלא הצליחה להתחיל. */
  done: Promise<{ text: string; reason: DictationEndReason }>;
}

/** מתחילה הקלטה אחת. השירות מקליט הקלטה אחת בכל פעם. */
export function startDictation(
  fetchStream: FetchStream,
  options: { before?: string; continuous?: boolean },
  hooks: DictationHooks,
): DictationHandle {
  const segments: Record<number, string> = {};
  let cancelled = false;
  let iterator: AsyncIterator<NetworkFetchStreamChunk> | null = null;

  const body = JSON.stringify({
    context: contextWords(options.before ?? ''),
    // continuous: ההקלטה נגמרת רק בעצירה (או אחרי 100 שניות); 0 = בלי סיום אוטומטי
    endAfterSilenceMs: options.continuous ? 0 : END_AFTER_SILENCE_MS,
    noSpeechMs: options.continuous ? 30000 : NO_SPEECH_MS,
  });

  const done = (async () => {
    let reason: DictationEndReason = 'error';
    for await (const ev of serviceLines(fetchStream, '/dictate', 'POST', body, 120000, (it) => {
      iterator = it;
    })) {
      if (cancelled) break;
      const event = ev as unknown as DictationEvent;
      if (event.type === 'ready') hooks.onReady?.();
      else if (event.type === 'update') {
        segments[event.segment] = event.text;
        hooks.onSegment?.(event.segment, event.text, event.final === true);
        hooks.onText?.(composeSegments(segments), event.final === true);
      } else if (event.type === 'error') hooks.onError?.(event.message);
      else if (event.type === 'end') reason = event.reason;
    }
    if (cancelled) reason = 'cancelled';
    return { text: composeSegments(segments), reason };
  })();

  return {
    done,
    async stop() {
      try {
        await lastOf(fetchStream, '/stop', 'POST');
      } catch {
        // זרם ההקלטה מדווח על התוצאה בכל מקרה
      }
    },
    cancel() {
      cancelled = true;
      void iterator?.return?.();
      lastOf(fetchStream, '/cancel', 'POST').catch(() => {});
    },
  };
}
