import { describe, expect, it } from 'vitest';
import type { NetworkFetchStreamChunk } from '../../src/types/otzaria_plugin';
import {
  composeSegments,
  contextWords,
  explainError,
  explainNotReady,
  startDictation,
  voiceStatus,
  type FetchStream,
} from '../../src/host/voice-service';
import { createVoiceDictation, superdocTarget, type DictationTarget } from '../../src/engine/voice-dictation';

/** תשובה מוכנה: מקטעי גוף שגבולותיהם אינם גבולות שורה. */
function respond(status: number, parts: string[]): AsyncIterable<NetworkFetchStreamChunk> {
  return (async function* () {
    yield { sequence: 0, type: 'response', status, ok: status === 200, headers: {} };
    let n = 1;
    for (const body of parts) yield { sequence: n++, type: 'data', body };
  })();
}

const READY = { service: 'otzaria-voice', pack: true, runtime: true, microphone: true };

/** שירות מזויף: /status, /dictate (אירועים מתוסרטים), /stop, /cancel. */
function fakeService(events: object[], calls: string[] = []): FetchStream {
  return (req) => {
    const path = req.url.replace('http://127.0.0.1:47811', '');
    calls.push(req.method + ' ' + path);
    expect(req.headers['X-Otzaria-Voice']).toBe('1');
    if (path === '/status') return respond(200, [JSON.stringify(READY) + '\n']);
    if (path === '/dictate') {
      const text = events.map((e) => JSON.stringify(e) + '\n').join('');
      // חיתוך באמצע שורה, כמו ברשת
      return respond(200, [text.slice(0, 17), text.slice(17)]);
    }
    return respond(200, ['{"stopped":true}\n']);
  };
}

type FakeTarget = DictationTarget & { text: () => string; moveCaret(p: number): void };

function fakeTarget(initial = ''): FakeTarget {
  const paras = [initial];
  let para = 0;
  let caret = initial.length;
  const idx = (at: { blockId: string }) => Number(at.blockId.slice(1));
  return {
    caret: () => ({ blockId: 'b' + para, offset: caret, story: null }),
    textBefore: async (at) => paras[idx(at)]!.slice(Math.max(0, at.offset - 40), at.offset),
    async replace(at, from, to, text) {
      const i = idx(at);
      paras[i] = paras[i]!.slice(0, from) + text + paras[i]!.slice(to);
      para = i;
      caret = from + text.length;
      return true;
    },
    async breakLine(at, offset) {
      const i = idx(at);
      const tail = paras[i]!.slice(offset);
      paras[i] = paras[i]!.slice(0, offset);
      paras.splice(i + 1, 0, tail);
      para = i + 1;
      caret = 0;
      return { blockId: 'b' + (i + 1), offset: 0, story: null };
    },
    text: () => paras.join('\n'),
    moveCaret(p) {
      caret = p;
    },
  };
}

const END = { type: 'end', reason: 'stopped' };

describe('voice-service', () => {
  it('מחבר מקטעים לפי הסדר ומדלג על ריקים', () => {
    expect(composeSegments({ 2: 'עולם', 0: 'שלום', 1: '' })).toBe('שלום עולם');
  });

  it('ההקשר הוא שתי המילים האחרונות', () => {
    expect(contextWords('א ב ג ד')).toEqual(['ג', 'ד']);
    expect(contextWords('')).toEqual([]);
  });

  it('מסביר מה חסר לפני שמתחילים', () => {
    expect(explainNotReady(READY)).toBeNull();
    expect(explainNotReady({ ...READY, pack: false })).toContain('חבילת הדיבור');
    expect(explainNotReady({ ...READY, runtime: false })).toContain('ONNX');
    expect(explainNotReady({ ...READY, microphone: false })).toContain('מיקרופון');
  });

  it('מזהה שירות אחר או שירות שלא עונה', async () => {
    await expect(voiceStatus(() => respond(200, ['{"service":"other"}\n']))).rejects.toMatchObject({
      code: 'unreachable',
    });
    await expect(
      voiceStatus(() => {
        throw new Error('connection refused');
      }),
    ).rejects.toMatchObject({ code: 'unreachable' });
  });

  it('הרשאה חסרה באוצריא מוסברת בהודעה של אוצריא, לא כ"שירות לא פועל"', async () => {
    const denied: FetchStream = () => {
      throw new Error('error.permission_denied: לתוסף אין הרשאת גישה לשירותים מקומיים (localhost).');
    };
    const err = await voiceStatus(denied).catch((e) => e);
    expect(err.code).toBe('permission_denied');
    expect(explainError(err)).toBe('לתוסף אין הרשאת גישה לשירותים מקומיים (localhost).');
  });

  it('שגיאת HTTP מגיעה עם הקוד שהשירות נתן', async () => {
    const h = startDictation(
      () => respond(503, ['{"error":"pack_installing","message":"x"}']),
      {},
      {},
    );
    await expect(h.done).rejects.toMatchObject({ code: 'pack_installing' });
  });

  it('קורא אירועים משורות חצויות ומחזיר טקסט וסיבה', async () => {
    const seen: string[] = [];
    const h = startDictation(
      fakeService([
        { type: 'ready' },
        { type: 'update', segment: 0, text: 'שלו' },
        { type: 'update', segment: 0, text: 'שלום', final: true },
        { type: 'update', segment: 1, text: 'עולם' },
        { type: 'end', reason: 'silence' },
      ]),
      {},
      { onText: (t) => seen.push(t) },
    );
    expect(await h.done).toEqual({ text: 'שלום עולם', reason: 'silence' });
    expect(seen).toEqual(['שלו', 'שלום', 'שלום עולם']);
  });
});

describe('voice-dictation', () => {
  function setup(events: object[], target: DictationTarget | null, calls: string[] = []) {
    const said: Array<[string, boolean]> = [];
    const states: string[] = [];
    const d = createVoiceDictation({
      fetchStream: () => fakeService(events, calls),
      target: () => target,
      say: (t, e) => said.push([t, e === true]),
      onState: (s) => states.push(s),
    });
    return { d, said, states };
  }

  it('כותב טיוטה ומחליף אותה במקומה בכל עדכון', async () => {
    const t = fakeTarget('בראשית');
    const { d, states } = setup(
      [
        { type: 'ready' },
        { type: 'update', segment: 0, text: 'ברא' },
        { type: 'update', segment: 0, text: 'ברא אלהים', final: true },
        END,
      ],
      t,
    );
    await d.toggle();
    expect(t.text()).toBe('בראשית ברא אלהים');
    expect(states).toEqual(['starting', 'listening', 'idle']);
    expect(d.state()).toBe('idle');
  });

  it('מקטע שנסגר אינו נכתב מחדש: מוחלף רק המקטע הנוכחי', async () => {
    const t = fakeTarget('');
    const writes: Array<[number, number]> = [];
    const spied: DictationTarget = {
      ...t,
      replace: async (at, from, to, text) => {
        writes.push([from, to]);
        return t.replace(at, from, to, text);
      },
    };
    const { d } = setup(
      [
        { type: 'ready' },
        { type: 'update', segment: 1, text: 'אבג' },
        { type: 'update', segment: 1, text: 'אבג דהו', final: true },
        { type: 'update', segment: 2, text: 'זח' },
        { type: 'update', segment: 2, text: 'זחט' },
        END,
      ],
      spied,
    );
    await d.toggle();
    expect(t.text()).toBe('אבג דהו זחט');
    // המקטע הראשון נכתב בטווח 0..; המקטע השני מתחיל אחרי "אבג דהו" (7 תווים)
    expect(writes.filter(([from]) => from >= 7).length).toBeGreaterThan(0);
    // אחרי שהמקטע השני התחיל, שום כתיבה לא חוזרת לטווח של הראשון
    const firstOfSecond = writes.findIndex(([from]) => from >= 7);
    expect(writes.slice(firstOfSecond).every(([from]) => from >= 7)).toBe(true);
  });

  it('לא מוסיף רווח בתחילת מסמך או אחרי רווח', async () => {
    const empty = fakeTarget('');
    await setup([{ type: 'ready' }, { type: 'update', segment: 0, text: 'שלום' }, END], empty).d.toggle();
    expect(empty.text()).toBe('שלום');

    const spaced = fakeTarget('א ');
    await setup([{ type: 'ready' }, { type: 'update', segment: 0, text: 'ב' }, END], spaced).d.toggle();
    expect(spaced.text()).toBe('א ב');
  });

  it('כשהסמן זז באמצע ההקלטה - עוצר ולא כותב במקום אחר', async () => {
    const t = fakeTarget('א');
    const service: FetchStream = (req) => {
      if (req.url.endsWith('/status')) return respond(200, [JSON.stringify(READY) + '\n']);
      if (req.url.endsWith('/stop')) return respond(200, ['{"stopped":true}\n']);
      return (async function* () {
        yield { sequence: 0, type: 'response', status: 200, ok: true, headers: {} } as const;
        yield { sequence: 1, type: 'data', body: '{"type":"update","segment":0,"text":"ב"}\n' } as const;
        t.moveCaret(0); // המשתמש לחץ בתחילת המסמך
        yield { sequence: 2, type: 'data', body: '{"type":"update","segment":0,"text":"ב ג"}\n' } as const;
        yield { sequence: 3, type: 'data', body: '{"type":"end","reason":"stopped"}\n' } as const;
      })();
    };
    const d = createVoiceDictation({
      fetchStream: () => service,
      target: () => t,
      say: () => {},
      onState: () => {},
    });
    await d.toggle();
    expect(t.text()).toBe('א ב');
  });

  it('לא מסתיימת בשקט, וממשיכה בהקלטה חדשה אחרי מגבלת הזמן', async () => {
    const t = fakeTarget('');
    const bodies: string[] = [];
    let sessions = 0;
    const service: FetchStream = (req) => {
      if (req.url.endsWith('/status')) return respond(200, [JSON.stringify(READY) + '\n']);
      bodies.push(req.body);
      sessions += 1;
      const events =
        sessions === 1
          ? [{ type: 'update', segment: 0, text: 'ראשון' }, { type: 'end', reason: 'time_limit' }]
          : [{ type: 'update', segment: 0, text: 'שני' }, END];
      return respond(200, [events.map((e) => JSON.stringify(e) + '\n').join('')]);
    };
    const d = createVoiceDictation({
      fetchStream: () => service,
      target: () => t,
      say: () => {},
      onState: () => {},
    });
    await d.toggle();
    expect(t.text()).toBe('ראשון שני');
    expect(sessions).toBe(2);
    expect(JSON.parse(bodies[0]!).endAfterSilenceMs).toBe(0);
  });

  it('בלי מסמך או עם בחירה - מסביר ולא פונה לשירות', async () => {
    const calls: string[] = [];
    const none = setup([], null, calls);
    await none.d.toggle();
    expect(none.said[0]?.[1]).toBe(true);
    expect(calls).toEqual([]);

    const selecting: DictationTarget = { ...fakeTarget('אבג'), caret: () => null };
    const sel = setup([], selecting, calls);
    await sel.d.toggle();
    expect(sel.said[0]?.[0]).toContain('מסומן');
    expect(calls).toEqual([]);
  });

  it('מסמך לקריאה בלבד - מסביר ולא פותח את המיקרופון', async () => {
    const calls: string[] = [];
    const said: Array<[string, boolean]> = [];
    const states: string[] = [];
    const d = createVoiceDictation({
      fetchStream: () => fakeService([{ type: 'ready' }, END], calls),
      target: () => fakeTarget('אבג'),
      say: (text, isError) => said.push([text, isError === true]),
      onState: (s) => states.push(s),
      canWrite: () => false,
    });
    await d.toggle();
    expect(said).toEqual([['המסמך פתוח לקריאה בלבד, ולכן אי אפשר להקליט לתוכו', true]]);
    expect(calls).toEqual([]);
    expect(states).toEqual([]);
  });

  it('ביטול בזמן ההתחלה (בדיקת השירות) - המיקרופון אינו נפתח', async () => {
    for (const how of ['cancel', 'toggle'] as const) {
      const calls: string[] = [];
      const states: string[] = [];
      let answer!: () => void;
      const statusGate = new Promise<void>((resolve) => {
        answer = resolve;
      });
      const service = fakeService([{ type: 'ready' }, END], calls);
      const slowStatus: FetchStream = (req) => {
        if (!req.url.endsWith('/status')) return service(req);
        calls.push(req.method + ' /status');
        return (async function* () {
          await statusGate;
          yield* respond(200, [JSON.stringify(READY) + '\n']);
        })();
      };
      const t = fakeTarget('אבג');
      const d = createVoiceDictation({
        fetchStream: () => slowStatus,
        target: () => t,
        say: () => {},
        onState: (s) => states.push(s),
      });
      const running = d.toggle();
      await Promise.resolve();
      expect(d.state()).toBe('starting');
      if (how === 'cancel') d.cancel();
      else await d.toggle();
      answer();
      await running;
      expect(calls).toEqual(['GET /status']);
      expect(states).toEqual(['starting', 'idle']);
      expect(t.text()).toBe('אבג');
    }
  });

  it('שירות שלא פועל מוסבר בעברית וחוזר למצב רגיל', async () => {
    const said: string[] = [];
    const d = createVoiceDictation({
      fetchStream: () => () => {
        throw new Error('refused');
      },
      target: () => fakeTarget('x'),
      say: (t) => said.push(t),
      onState: () => {},
    });
    await d.toggle();
    expect(said[0]).toContain('שירות הדיבור של אוצריא אינו פועל');
    expect(d.state()).toBe('idle');
  });

  it('מחוץ לאוצריא (אין SDK) - הודעה ברורה, בלי קריסה', async () => {
    const said: string[] = [];
    const d = createVoiceDictation({
      fetchStream: () => {
        throw new VoiceServiceErrorLike();
      },
      target: () => fakeTarget('x'),
      say: (t) => said.push(t),
      onState: () => {},
    });
    await d.toggle();
    expect(said[0]).toContain('אוצריא');
    expect(d.state()).toBe('idle');
  });
});

import { VoiceServiceError } from '../../src/host/voice-service';

class VoiceServiceErrorLike extends VoiceServiceError {
  constructor() {
    super('no_host', 'ההקלטה זמינה רק כשהעורך פועל בתוך אוצריא');
  }
}

describe('superdocTarget', () => {
  function fakeSuperdoc(replaceReceipt: object) {
    const calls: string[] = [];
    const sd = {
      activeEditor: {
        doc: {
          insert: async () => ({ success: true }),
          replace: async () => replaceReceipt,
          getNodeById: async () => ({}),
        },
      },
      ui: {
        selection: {
          getSnapshot: () => null,
          apply: () => {
            calls.push('apply');
            return { ok: true };
          },
        },
      },
    };
    return { sd, calls };
  }
  const at = { blockId: 'b', offset: 0, story: null };

  it('עדכון זהה למה שכבר כתוב (NO_OP) אינו כשל', async () => {
    const { sd, calls } = fakeSuperdoc({ success: false, failure: { code: 'NO_OP' } });
    const target = superdocTarget(sd as never);
    expect(await target!.replace(at, 0, 3, 'abc')).toBe(true);
    expect(calls).toEqual(['apply']);
  });

  it('כשל אמיתי של ה-API נשאר כשל', async () => {
    const { sd } = fakeSuperdoc({ success: false, failure: { code: 'INVALID_TARGET' } });
    expect(await superdocTarget(sd as never)!.replace(at, 0, 3, 'abc')).toBe(false);
  });
});

import { applySpokenPunctuation, splitLineCommands } from '../../src/engine/spoken-punctuation';

describe('סימני פיסוק בדיבור', () => {
  it('פסיק, נקודה וסימן שאלה הופכים לסימנים הנצמדים למילה שלפניהם', () => {
    expect(applySpokenPunctuation('שלום פסיק עולם נקודה')).toBe('שלום, עולם.');
    expect(applySpokenPunctuation('מה שלומך סימן שאלה')).toBe('מה שלומך?');
    expect(applySpokenPunctuation('אמר נקודתיים כך')).toBe('אמר: כך');
  });

  it('פסיק רישיה הוא מונח ולא פקודה', () => {
    expect(applySpokenPunctuation('דבר שהוא פסיק רישיה')).toBe('דבר שהוא פסיק רישיה');
  });

  it('סימן בתחילת הטקסט נשאר לבדו', () => {
    expect(applySpokenPunctuation('פסיק ואחר כך')).toBe(', ואחר כך');
    expect(applySpokenPunctuation('סימן')).toBe('סימן');
  });

  it('בכתיבה: סימן שפותח מקטע חדש נצמד לטקסט שלפניו', async () => {
    const t = fakeTarget('');
    const { d } = setupFor(
      [
        { type: 'ready' },
        { type: 'update', segment: 1, text: 'שלום', final: true },
        { type: 'update', segment: 2, text: 'פסיק', final: true },
        { type: 'update', segment: 3, text: 'עולם נקודה', final: true },
        END,
      ],
      t,
    );
    await d.toggle();
    expect(t.text()).toBe('שלום, עולם.');
  });
});

function setupFor(events: object[], target: DictationTarget) {
  const d = createVoiceDictation({
    fetchStream: () => fakeService(events),
    target: () => target,
    say: () => {},
    onState: () => {},
  });
  return { d };
}

describe('ירידת שורה בדיבור', () => {
  it('מפצל לפי פקודות ירידת שורה', () => {
    expect(splitLineCommands('א ב')).toEqual(['א ב']);
    expect(splitLineCommands('א תרד שורה ב')).toEqual(['א', 'ב']);
    expect(splitLineCommands('א שורה חדשה')).toEqual(['א', '']);
    expect(splitLineCommands('תרד שורה')).toEqual(['', '']);
  });

  it('"תרד שורה" יורד שורה כמו Enter, והמשך הדיבור נכתב בפסקה החדשה', async () => {
    const t = fakeTarget('');
    const { d } = setupFor(
      [
        { type: 'ready' },
        { type: 'update', segment: 1, text: 'שלום עולם תרד', final: false },
        { type: 'update', segment: 1, text: 'שלום עולם תרד שורה', final: true },
        { type: 'update', segment: 2, text: 'פסקה שניה', final: true },
        END,
      ],
      t,
    );
    await d.toggle();
    expect(t.text()).toBe('שלום עולם\nפסקה שניה');
  });

  it('פקודה באמצע מקטע מפצלת אותו, ופקודה בסוף משאירה פסקה ריקה לדיבור הבא', async () => {
    const t = fakeTarget('');
    const { d } = setupFor(
      [
        { type: 'ready' },
        { type: 'update', segment: 1, text: 'א ב תרד שורה ג ד', final: true },
        { type: 'update', segment: 2, text: 'ה תרד שורה', final: true },
        { type: 'update', segment: 3, text: 'ו', final: true },
        END,
      ],
      t,
    );
    await d.toggle();
    expect(t.text()).toBe('א ב\nג ד ה\nו');
  });

  it('בטיוטה הפקודה לא יוצרת פסקה', async () => {
    const t = fakeTarget('');
    const { d } = setupFor(
      [{ type: 'ready' }, { type: 'update', segment: 1, text: 'שלום תרד שורה', final: false }, END],
      t,
    );
    await d.toggle();
    expect(t.text()).toBe('שלום');
  });
});
