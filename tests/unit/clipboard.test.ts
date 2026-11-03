/**
 * שכבת הלוח.
 *
 * שני דברים נמדדים כאן, ושניהם היו התקלה:
 *
 * 1. **מה נשלח למנוע.** הכפיל אינו „מקליט ומחזיר true”. הוא **מאמת** — target
 *    שאינו `SelectionTarget`, `serializeSelection` בלי `includeHtml`, מחיקה
 *    בלי target — כל אחד מאלה מפיל את הבדיקה. הגרסה הקודמת של
 *    tests/unit/ribbon-commands.test.ts החזירה `true` לכל קריאה ואישרה בירוק
 *    payloads שהמנוע דוחה; כפיל כזה בודק ש-JavaScript מעביר ארגומנטים.
 *
 * 2. **שהכשל מגיע למשתמש.** לוח מערכת שנחסם, `serializeSelection` שזורק,
 *    קבלה עם `success: false` — כל אחד מהם חייב לחזור כ-`ok: false` עם הודעה
 *    בעברית. „העתקה שקטה שלא קרתה” היא בדיוק מה שהיה כאן קודם.
 *
 * `navigator.clipboard` אינו קיים ב-jsdom, ולכן הוא נשתל לכל מקרה בנפרד. זה
 * גם מה שמאפשר לכפול את המצב הנפוץ באמת ב-`file://`: `NotAllowedError`.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  copySelection,
  cutSelection,
  internalClipboard,
  pasteFromClipboard,
  selectWholeDocument,
  type ClipboardDocumentApi,
  type ClipboardHost,
  type ClipboardPayload,
  type ClipboardPayloadItem,
  type ClipboardPlan,
  type SelectionSnapshot,
} from '../../src/engine/clipboard';

/* ------------------------------------------------------------------ */
/* כפיל הבחירה                                                        */
/* ------------------------------------------------------------------ */

/** בחירה חוקית: טווח בין שתי נקודות טקסט. זו הצורה שפעולות הכתיבה מקבלות. */
const RANGE = {
  kind: 'selection',
  start: { kind: 'text', blockId: 'p1', offset: 0 },
  end: { kind: 'text', blockId: 'p1', offset: 5 },
} as const;

/** סמן: אותה צורה בלי טווח. זה המצב הרגיל של הדבקה, והוא פסול להעתקה. */
const CARET = {
  kind: 'selection',
  start: { kind: 'text', blockId: 'p2', offset: 3 },
  end: { kind: 'text', blockId: 'p2', offset: 3 },
} as const;

/**
 * הטווח שהמנוע מדווח עליו בקבלה כ„מה שנוצר” (`effects.insertedText`). בלוק
 * אחר מזה שהסמן היה בו, כדי שהזזת הסמן לא תוכל להצליח במקרה. הקבלה האמיתית
 * נושאת שם גם `kind` ו-`text`; מה שנצרך כאן הוא ה-`selectionTarget` בלבד.
 */
const INSERTED = {
  kind: 'selection',
  start: { kind: 'text', blockId: 'p9', offset: 4 },
  end: { kind: 'text', blockId: 'p9', offset: 9 },
} as const;

/** אותו קצה, מכווץ לנקודה — זה מה ש-`apply` אמור לקבל אחרי הדבקה. */
const CARET_AFTER_PASTE = {
  kind: 'selection',
  start: { kind: 'text', blockId: 'p9', offset: 9 },
  end: { kind: 'text', blockId: 'p9', offset: 9 },
} as const;

/**
 * הסמן שאחרי „גזור”: תחילת הטווח שנגזר, מכווצת לנקודה. זה מה ש-`apply` אמור
 * לקבל — וזה גם מה ש-Word עושה, הסמן נשאר במקום שממנו הטקסט נעלם.
 */
const CARET_AFTER_CUT = {
  kind: 'selection',
  start: { kind: 'text', blockId: 'p1', offset: 0 },
  end: { kind: 'text', blockId: 'p1', offset: 0 },
} as const;

/** האם זה באמת `SelectionTarget`, ולא סתם אובייקט שהועבר הלאה. */
function isSelectionTarget(value: unknown): boolean {
  const target = value as { kind?: string; start?: { kind?: string }; end?: { kind?: string } } | null;
  if (!target || target.kind !== 'selection') return false;
  const points = [target.start, target.end];
  return points.every((point) => point?.kind === 'text' || point?.kind === 'nodeEdge');
}

function selectionSurface(snapshot: SelectionSnapshot | undefined) {
  return { getSnapshot: () => snapshot };
}

const READY_SELECTION: SelectionSnapshot = {
  status: 'ready',
  empty: false,
  selectionTarget: RANGE as never,
};

/* ------------------------------------------------------------------ */
/* כפיל ה-Document API                                                */
/* ------------------------------------------------------------------ */

/** ה-payload שהמנוע האמיתי מחזיר: text/plain, text/html ו-fragment פרטי. */
function enginePayload(text = 'שלום'): ClipboardPayload {
  return {
    source: 'api',
    items: [
      { type: 'text/plain', kind: 'string', data: text },
      { type: 'text/html', kind: 'string', data: `<p>${text}</p>` },
      { type: 'application/x-superdoc-v2-fragment', kind: 'string', data: '{"blocks":[]}' },
    ],
  };
}

interface FakeCall {
  op: string;
  input: unknown;
}

/**
 * הוולידטור של `clipboard.parse`/`clipboard.insert` כפי שהוא במנוע: `items`
 * מערך, `type` מחרוזת לא ריקה, `kind` אחד משניים, ו-`data` **מהטיפוס** שהוא
 * מכריז עליו. זה מה שהופך את הכפיל לבדיקה ולא להקלטה.
 */
function assertPayload(payload: ClipboardPayload | undefined, op: string): void {
  if (!payload || typeof payload !== 'object' || !Array.isArray(payload.items)) {
    throw new Error(`${op} requires a ClipboardPayload with items`);
  }
  payload.items.forEach((item: ClipboardPayloadItem, index: number) => {
    if (!item || typeof item !== 'object') throw new Error(`${op} item ${index} must be an object`);
    if (typeof item.type !== 'string' || item.type.length === 0) {
      throw new Error(`${op} item ${index} requires a MIME type`);
    }
    if (item.kind !== 'string' && item.kind !== 'bytes') {
      throw new Error(`${op} item ${index} has an invalid kind`);
    }
    if (item.kind === 'string' && typeof item.data !== 'string') {
      throw new Error(`${op} item ${index} requires string data`);
    }
    if (item.kind === 'bytes' && !(item.data instanceof Uint8Array)) {
      throw new Error(`${op} item ${index} requires Uint8Array data`);
    }
  });
}

/** אותו דבר ל-`plan`: `fragment.blocks` מערך, ו-`diagnostics` מערך. */
function assertPlan(plan: ClipboardPlan | undefined, op: string): void {
  if (!plan || typeof plan !== 'object') throw new Error(`${op} plan must be an object`);
  if (!plan.fragment || typeof plan.fragment !== 'object' || !Array.isArray(plan.fragment.blocks)) {
    throw new Error(`${op} fragment.blocks must be an array`);
  }
  if (!Array.isArray(plan.diagnostics)) throw new Error(`${op} plan.diagnostics must be an array`);
}

/** התכנית שהמנוע מחזיר מ-`parse`. */
function enginePlan(): ClipboardPlan {
  return { fragment: { blocks: [{ kind: 'paragraph', runs: [] }] }, diagnostics: [] };
}

function fakeDoc(overrides: Partial<ClipboardDocumentApi> = {}) {
  const calls: FakeCall[] = [];

  const doc: ClipboardDocumentApi = {
    clipboard: {
      serializeSelection: (input) => {
        // הוולידטור של המנוע: קלט שאינו אובייקט זורק INVALID_INPUT, ו-target
        // שאינו SelectionTarget נדחה. כפיל שמקבל הכול אינו מודד כלום.
        if (input !== undefined && typeof input !== 'object') {
          throw new Error('serializeSelection input must be an object when provided');
        }
        if (input?.target !== undefined && !isSelectionTarget(input.target)) {
          throw new Error('serializeSelection target must be a SelectionTarget');
        }
        calls.push({ op: 'clipboard.serializeSelection', input });
        return { payload: enginePayload() };
      },
      parse: (payload) => {
        assertPayload(payload, 'clipboard.parse');
        calls.push({ op: 'clipboard.parse', input: payload });
        return { success: true, plan: enginePlan() };
      },
      insert: (input) => {
        // „exactly one of payload, plan, or fragment” — זה החוזה, והוא זורק
        // INVALID_INPUT על שניים או על אפס.
        const provided = [input?.payload, input?.plan, input?.fragment].filter((value) => value !== undefined);
        if (provided.length !== 1) {
          throw new Error('clipboard.insert requires exactly one of payload, plan, or fragment');
        }
        if (input.payload !== undefined) assertPayload(input.payload, 'clipboard.insert');
        if (input.plan !== undefined) assertPlan(input.plan, 'clipboard.insert');
        // `target` אופציונלי בחוזה — אבל כשהוא נשלח הוא נבדק, ו-target פסול
        // נדחה ב-INVALID_INPUT במקום להידחף בסוף המסמך.
        if (input.target !== undefined && !isSelectionTarget(input.target)) {
          throw new Error('clipboard.insert target must be a ClipboardTarget');
        }
        calls.push({ op: 'clipboard.insert', input });
        // המנוע מדווח על הטווח שנוצר, ולא רק על „הצליח” — זה מה שמאפשר
        // להזיז את הסמן לסופו בלי לספור תווים.
        return { success: true, effects: { insertedText: [{ selectionTarget: INSERTED as never }] } };
      },
    },
    delete: (input) => {
      if (!isSelectionTarget(input?.target)) {
        throw new Error('delete requires a SelectionTarget');
      }
      calls.push({ op: 'delete', input });
      return { success: true };
    },
    ...overrides,
  };

  const host: ClipboardHost = {
    activeEditor: { doc },
    ui: { selection: selectionSurface(READY_SELECTION) },
  };

  return { doc, host, calls };
}

/* ------------------------------------------------------------------ */
/* כפיל לוח המערכת                                                    */
/* ------------------------------------------------------------------ */

const NO_CLIPBOARD = Symbol('no-clipboard');

function setSystemClipboard(value: unknown): void {
  if (value === NO_CLIPBOARD) {
    Reflect.deleteProperty(navigator, 'clipboard');
    return;
  }
  Object.defineProperty(navigator, 'clipboard', { value, configurable: true, writable: true });
}

/**
 * `ClipboardItem` אינו קיים ב-jsdom, אבל הוא כן קיים ב-WebView2 — ולכן הוא
 * נשתל כברירת מחדל: המסלול הנפוץ במציאות הוא זה שכותב HTML, וסביבת בדיקה
 * שאין בה את הבנאי הייתה בודקת רק את מסלול הנפילה.
 */
class FakeClipboardItem {
  readonly types: string[];
  private readonly parts: Record<string, Blob>;
  constructor(parts: Record<string, Blob>) {
    this.parts = parts;
    this.types = Object.keys(parts);
  }
  async getType(type: string): Promise<Blob> {
    return this.parts[type]!;
  }
}

function setClipboardItem(available: boolean): void {
  if (!available) {
    Reflect.deleteProperty(globalThis, 'ClipboardItem');
    return;
  }
  Object.defineProperty(globalThis, 'ClipboardItem', {
    value: FakeClipboardItem,
    configurable: true,
    writable: true,
  });
}

/** `Blob.text()` אינו קיים ב-jsdom שהבדיקות רצות בו; `FileReader` כן. */
function readBlob(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsText(blob);
  });
}

/** לוח שעובד. שומר את מה שנכתב, כדי שנוכל לבדוק שנכתבו גם HTML וגם טקסט. */
function workingClipboard() {
  const written: Array<Record<string, string>> = [];
  return {
    written,
    api: {
      write: vi.fn(async (items: ClipboardItem[]) => {
        const entry: Record<string, string> = {};
        for (const item of items) {
          for (const type of item.types) entry[type] = await readBlob(await item.getType(type));
        }
        written.push(entry);
      }),
      writeText: vi.fn(async (text: string) => {
        written.push({ 'text/plain': text });
      }),
    },
  };
}

/** המצב הנפוץ ב-`file://`: ההרשאה נדחית. */
function deniedClipboard() {
  const denied = (): Promise<never> =>
    Promise.reject(Object.assign(new Error('denied'), { name: 'NotAllowedError' }));
  return { write: vi.fn(denied), writeText: vi.fn(denied), read: vi.fn(denied), readText: vi.fn(denied) };
}

/**
 * פריט קריאה מלוח המערכת.
 *
 * ה-Blob כאן הוא כפיל ולא `new Blob(...)`, ובכוונה: ל-Blob של הדפדפן **יש**
 * `text()` ו-`arrayBuffer()`, ולזה של jsdom אין. כפיל עם המתודות האמיתיות
 * מתאר את המציאות יותר טוב מהאובייקט החסר שהסביבה מספקת.
 */
function readableItem(parts: Record<string, string | Uint8Array>) {
  return {
    types: Object.keys(parts),
    getType: async (type: string) => {
      const value = parts[type]!;
      return {
        type,
        text: async () => (typeof value === 'string' ? value : new TextDecoder().decode(value)),
        arrayBuffer: async () =>
          typeof value === 'string'
            ? new TextEncoder().encode(value).buffer
            : (value.buffer as ArrayBuffer),
      } as unknown as Blob;
    },
  } as unknown as ClipboardItem;
}

/** לוח שאפשר לקרוא ממנו. */
function readableClipboard(parts: Record<string, string | Uint8Array>) {
  return {
    read: vi.fn(async () => [readableItem(parts)]),
    readText: vi.fn(async () => String(parts['text/plain'] ?? '')),
  };
}

beforeEach(() => {
  internalClipboard.clear();
  setSystemClipboard(NO_CLIPBOARD);
  setClipboardItem(true);
});

afterEach(() => {
  setSystemClipboard(NO_CLIPBOARD);
  setClipboardItem(false);
  vi.restoreAllMocks();
});

/* ------------------------------------------------------------------ */

describe('copySelection', () => {
  it('מסדרת את הבחירה עם ה-target שנקרא, ומבקשת גם HTML', async () => {
    const { host, calls } = fakeDoc();
    setSystemClipboard(workingClipboard().api);

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(true);
    expect(calls).toEqual([
      { op: 'clipboard.serializeSelection', input: { target: RANGE, includeHtml: true } },
    ]);
  });

  it('כותבת ללוח המערכת גם text/html וגם text/plain', async () => {
    // זה מה שמאפשר להדביק ל-Word אמיתי בלי לאבד עיצוב. `writeText` לבדו היה
    // משטח את הכול.
    const { host } = fakeDoc();
    const clipboard = workingClipboard();
    setSystemClipboard(clipboard.api);

    await copySelection(host);

    expect(clipboard.api.write).toHaveBeenCalledTimes(1);
    expect(clipboard.written[0]).toEqual({ 'text/html': '<p>שלום</p>', 'text/plain': 'שלום' });
  });

  it('אינה שולחת את ה-MIME הפרטי של המנוע ללוח המערכת', async () => {
    // לוח המערכת מסנן טיפוסים שאינם ברשימה המותרת, וניסיון כזה מפיל את כל
    // הקריאה — כלומר היה מבטל גם את ההעתקה של הטקסט.
    const { host } = fakeDoc();
    const clipboard = workingClipboard();
    setSystemClipboard(clipboard.api);

    await copySelection(host);

    expect(Object.keys(clipboard.written[0]!)).not.toContain('application/x-superdoc-v2-fragment');
  });

  it('נופלת ל-writeText כשאין ClipboardItem בסביבה', async () => {
    const { host } = fakeDoc();
    const clipboard = workingClipboard();
    setSystemClipboard(clipboard.api);
    setClipboardItem(false);

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(true);
    expect(clipboard.api.write).not.toHaveBeenCalled();
    expect(clipboard.api.writeText).toHaveBeenCalledWith('שלום');
  });

  it('נופלת ל-writeText כש-`write` נדחה אבל `writeText` עובר', async () => {
    // המצב שנמדד ב-WebView2: הכתיבה העשירה נחסמת, והמצומצמת כן עוברת.
    const { host } = fakeDoc();
    const writeText = vi.fn(async () => {});
    setSystemClipboard({
      write: vi.fn(() => Promise.reject(new Error('denied'))),
      writeText,
    });

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(true);
    expect(writeText).toHaveBeenCalledWith('שלום');
  });

  it('הרשאה שנדחתה: התוכן בלוח הפנימי, וההודעה אומרת את זה', async () => {
    const { host } = fakeDoc();
    setSystemClipboard(deniedClipboard());

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(false);
    if (outcome.ok) return;
    expect(outcome.reason).toBe('system-clipboard-blocked');
    expect(outcome.message).toContain('לוח המערכת');
    expect(outcome.message).toContain('Ctrl+C');
    // לא שקר על הצלחה — אבל גם לא אובדן: התוכן קיים.
    expect(internalClipboard.read()).toEqual(enginePayload());
  });

  it('אין `navigator.clipboard` בכלל — אותה הודעה, בלי חריגה', async () => {
    const { host } = fakeDoc();

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.reason).toBe('system-clipboard-blocked');
    expect(internalClipboard.read()).not.toBeNull();
  });

  it('בחירה ריקה נעצרת לפני המנוע, עם „יש לסמן טקסט תחילה”', async () => {
    const { host, calls } = fakeDoc();
    host.ui = { selection: selectionSurface({ status: 'ready', empty: true, selectionTarget: null }) };

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.message).toContain('יש לסמן טקסט תחילה');
    expect(calls).toEqual([]);
  });

  it('קריאת בחירה שטרם הסתיימה אינה עוצרת העתקה', async () => {
    // `status: 'pending'` אינו „לא סימנת”: המנוע כן יודע מה הבחירה החיה, גם
    // אם ה-slice של ה-UI עוד לא התעדכן. רק בחירה **ריקה** ודאית עוצרת.
    const { host, calls } = fakeDoc();
    host.ui = { selection: selectionSurface({ status: 'pending', selectionTarget: null }) };
    setSystemClipboard(workingClipboard().api);

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(true);
    expect(calls[0]!.input).toEqual({ includeHtml: true });
  });

  it('תצלום stale שמתייצב — „העתק” ממתין, ולא נכשל על בחירה שקיימת', async () => {
    // המרוץ שהתגלה: Shift+חץ יוצר טווח אמיתי במסמך, אבל `getSnapshot()` ממשיך
    // לדווח `empty: true` יחד עם `status` שאינו `ready` במשך עד כ-820ms. שני
    // התצלומים הראשונים כאן חוזרים stale (כמו שנמדד), והשלישי — אחרי שהבחירה
    // כבר התייצבה. בלי הבדיקה בכל הענפים, „העתק” היה נכשל מיד עם „יש לסמן
    // טקסט תחילה” על בחירה שקיימת.
    const { host, calls } = fakeDoc();
    let reads = 0;
    host.ui = {
      selection: {
        getSnapshot: () => {
          reads += 1;
          return reads <= 2
            ? { status: 'stale', empty: true, selectionTarget: RANGE as never }
            : { status: 'ready', empty: false, selectionTarget: RANGE as never };
        },
      },
    };
    setSystemClipboard(workingClipboard().api);

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(true);
    expect(reads).toBeGreaterThan(2);
    // אחרי שהתייצב יש target אמיתי — ולא הקריאה בלי target (שנכשלת במנוע האמיתי).
    expect(calls[0]!.input).toEqual({ target: RANGE, includeHtml: true });
  });

  it('אין משטח בחירה — מסדרת את הבחירה החיה בלי target', async () => {
    // החוזה קובע ש-serializeSelection מסדר "the current or supplied model
    // selection", ולכן „העתק” אינו נופל רק מפני שלא הצלחנו לקרוא את הבחירה.
    const { host, calls } = fakeDoc();
    host.ui = null;
    setSystemClipboard(workingClipboard().api);

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(true);
    expect(calls[0]!.input).toEqual({ includeHtml: true });
  });

  it('אין `doc.clipboard` — מנוטרל עם „אינו זמין בגרסה זו”', async () => {
    const { host } = fakeDoc({ clipboard: undefined });

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('command-unsupported');
      expect(outcome.message).toBe('ההעתקה נכשלה: אינו זמין בגרסה זו');
    }
  });

  it('אין Document API בכלל — אותו כשל סגור', async () => {
    for (const host of [null, undefined, {}, { activeEditor: null }, { activeEditor: { doc: null } }]) {
      const outcome = await copySelection(host as ClipboardHost);

      expect(outcome.ok).toBe(false);
      if (!outcome.ok) expect(outcome.reason).toBe('command-unsupported');
    }
  });

  it('סובלת גם קריאה שמחזירה הבטחה — הפאסדה בדפדפן א-סינכרונית', async () => {
    const { host } = fakeDoc({
      clipboard: { serializeSelection: () => Promise.resolve({ payload: enginePayload('א') }) },
    });
    setSystemClipboard(workingClipboard().api);

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(true);
    expect(internalClipboard.read()).toEqual(enginePayload('א'));
  });

  it('סדרוּר שזורק אינו מפיל את הרצועה', async () => {
    const { host } = fakeDoc({
      clipboard: {
        serializeSelection: () => {
          throw new Error('boom');
        },
      },
    });

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('threw');
      expect(outcome.message).toBe('ההעתקה נכשלה: boom');
    }
  });

  it('הבטחה שנדחית מטופלת כמו זריקה', async () => {
    const { host } = fakeDoc({
      clipboard: { serializeSelection: () => Promise.reject(new Error('boom')) },
    });

    const outcome = await copySelection(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.reason).toBe('threw');
  });

  it('payload ריק או חסר אינו „העתקה שהצליחה”', async () => {
    for (const payload of [undefined, { items: [] }, { items: null }]) {
      internalClipboard.clear();
      const { host } = fakeDoc({
        clipboard: { serializeSelection: () => ({ payload: payload as never }) },
      });

      const outcome = await copySelection(host);

      expect(outcome.ok, JSON.stringify(payload)).toBe(false);
      if (!outcome.ok) expect(outcome.message).toContain('יש לסמן טקסט תחילה');
      expect(internalClipboard.read()).toBeNull();
    }
  });
});

describe('cutSelection', () => {
  it('מסדרת ואז מוחקת — באותו target', async () => {
    const { host, calls } = fakeDoc();
    setSystemClipboard(workingClipboard().api);

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(true);
    expect(calls.map((call) => call.op)).toEqual(['clipboard.serializeSelection', 'delete']);
    expect(calls[1]!.input).toEqual({ target: RANGE });
  });

  it('סדרוּר שנכשל אינו מוחק כלום', async () => {
    // זה כל הטעם בסדר הפעולות: מחיקה לפני שיש עותק היא טקסט שנעלם.
    const { calls, host } = fakeDoc();
    const doc = host.activeEditor!.doc!;
    doc.clipboard = {
      serializeSelection: () => {
        throw new Error('boom');
      },
    };

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(false);
    expect(calls.map((call) => call.op)).not.toContain('delete');
  });

  it('לוח מערכת שנחסם כן ממשיך למחיקה, וההודעה אומרת מה קרה', async () => {
    // התוכן קיים בלוח הפנימי, ולכן הגזירה אינה אובדן. הודעה שקטה כאן הייתה
    // משאירה את המשתמש עם טקסט שנמחק ולוח מערכת ריק.
    const { host, calls } = fakeDoc();
    setSystemClipboard(deniedClipboard());

    const outcome = await cutSelection(host);

    expect(calls.map((call) => call.op)).toEqual(['clipboard.serializeSelection', 'delete']);
    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('system-clipboard-blocked');
      expect(outcome.message).toContain('Ctrl+X');
    }
    expect(internalClipboard.read()).toEqual(enginePayload());
  });

  it('אין מסלול לקרוא את הבחירה — נכשלת, ולא מנחשת target', async () => {
    const { host, calls } = fakeDoc();
    host.ui = null;

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.reason).toBe('host-capability-unavailable');
    expect(calls).toEqual([]);
  });

  it('בחירה שטרם נקראה — „המסמך עדיין נטען”, ולא „לא סימנת”', async () => {
    // שני המצבים מגיעים כ-`selectionTarget: null`, וההבחנה ביניהם היא ההבדל
    // בין הודעה שאפשר לפעול לפיה לבין הודעה מבלבלת.
    const { host, calls } = fakeDoc();
    host.ui = { selection: selectionSurface({ status: 'pending', selectionTarget: null }) };

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('not-ready');
      expect(outcome.message).toBe('הגזירה נכשלה: המסמך עדיין נטען');
    }
    expect(calls).toEqual([]);
  });

  it('בחירה ריקה — „יש לסמן טקסט תחילה”', async () => {
    const { host, calls } = fakeDoc();
    host.ui = { selection: selectionSurface({ status: 'ready', empty: true, selectionTarget: null }) };

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.message).toBe('הגזירה נכשלה: יש לסמן טקסט תחילה');
    expect(calls).toEqual([]);
  });

  it('תצלום stale שמתייצב — „גזור” ממתין, ולא נכשל על בחירה שקיימת', async () => {
    // אותו מרוץ כמו ב„העתק”, אבל „גזור” אינו יכול ליפול על „המנוע יסדר את
    // הבחירה החיה בעצמו”: `doc.delete` דורש target מפורש. לכן כאן הפתרון הוא
    // המתנה להתייצבות, ולא מקור חלופי. שני התצלומים הראשונים חוזרים כ-stale
    // (בדיוק כמו שנמדד ב-Shift+חץ), והשלישי — אחרי שהבחירה כבר התייצבה.
    const { host, calls } = fakeDoc();
    let reads = 0;
    host.ui = {
      selection: {
        getSnapshot: () => {
          reads += 1;
          return reads <= 2
            ? { status: 'stale', empty: true, selectionTarget: RANGE as never }
            : { status: 'ready', empty: false, selectionTarget: RANGE as never };
        },
      },
    };
    setSystemClipboard(workingClipboard().api);

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(true);
    expect(reads).toBeGreaterThan(2);
    const deleteCall = calls.find((call) => call.op === 'delete');
    expect(deleteCall?.input).toEqual({ target: RANGE });
  });

  it('`getSnapshot` שזורק אינו מפיל את הכפתור', async () => {
    const { host } = fakeDoc();
    host.ui = {
      selection: {
        getSnapshot: () => {
          throw new Error('boom');
        },
      },
    };

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.message).toContain('המסמך עדיין נטען');
  });

  it('אין `doc.delete` — „גזור” מנוטרל, בלי לסדר כלום', async () => {
    const { host, calls } = fakeDoc({ delete: undefined });

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('command-unsupported');
      expect(outcome.message).toBe('הגזירה נכשלה: אינו זמין בגרסה זו');
    }
    expect(calls).toEqual([]);
  });

  it('קבלת מחיקה שנכשלה מגיעה למשתמש עם הקוד של המנוע', async () => {
    const { host } = fakeDoc({
      delete: () => ({ success: false, failure: { code: 'DOCUMENT_READONLY' } }),
    });
    setSystemClipboard(workingClipboard().api);

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('DOCUMENT_READONLY');
      expect(outcome.message).toBe('הגזירה נכשלה: המסמך פתוח לקריאה בלבד');
    }
  });

  it('מחיקה שזורקת אינה מפילה את הרצועה', async () => {
    const { host } = fakeDoc({
      delete: () => {
        throw new Error('boom');
      },
    });
    setSystemClipboard(workingClipboard().api);

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('threw');
      expect(outcome.message).toBe('הגזירה נכשלה: boom');
    }
  });

  it('הסמן חוזר לתחילת מה שנגזר', async () => {
    // אחרי `doc.delete` המנוע נשאר בלי בחירה בכלל — נמדד מול המנוע האמיתי —
    // וההקלדה הבאה אינה נכנסת לשום מקום. ההנמקה ב-`caretAfterCut`.
    const { host } = fakeDoc();
    const apply = vi.fn(() => ({ ok: true }));
    host.ui = { selection: { getSnapshot: () => READY_SELECTION, apply } };
    setSystemClipboard(workingClipboard().api);

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(true);
    expect(apply).toHaveBeenCalledTimes(1);
    expect(apply).toHaveBeenCalledWith(CARET_AFTER_CUT);
  });

  it('גם כשלוח המערכת נחסם — הטקסט נגזר, ולכן הסמן חוזר', async () => {
    const { host } = fakeDoc();
    const apply = vi.fn(() => ({ ok: true }));
    host.ui = { selection: { getSnapshot: () => READY_SELECTION, apply } };
    setSystemClipboard(deniedClipboard());

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(false);
    expect(apply).toHaveBeenCalledWith(CARET_AFTER_CUT);
  });

  it('מחיקה שנכשלה אינה מזיזה את הסמן — אין מה לגזור, ואין לאן לחזור', async () => {
    const { host } = fakeDoc({
      delete: () => ({ success: false, failure: { code: 'DOCUMENT_READONLY' } }),
    });
    const apply = vi.fn(() => ({ ok: true }));
    host.ui = { selection: { getSnapshot: () => READY_SELECTION, apply } };
    setSystemClipboard(workingClipboard().api);

    await cutSelection(host);

    expect(apply).not.toHaveBeenCalled();
  });

  it('`apply` שנכשל סגור אינו הופך גזירה שהצליחה לכישלון, ואינו חוזר על עצמו', async () => {
    // להבדיל מההדבקה: היעד כאן הוא בלוק שכבר היה במסמך, ולכן אין ממה להמתין.
    const { host } = fakeDoc();
    const apply = vi.fn(() => ({ ok: false, reason: 'target-unresolved' }));
    host.ui = { selection: { getSnapshot: () => READY_SELECTION, apply } };
    setSystemClipboard(workingClipboard().api);

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(true);
    expect(apply).toHaveBeenCalledTimes(1);
  });

  it('`apply` שזורק אינו מפיל את הגזירה', async () => {
    const { host } = fakeDoc();
    const apply = vi.fn(() => {
      throw new Error('boom');
    });
    host.ui = { selection: { getSnapshot: () => READY_SELECTION, apply } };
    setSystemClipboard(workingClipboard().api);

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(true);
  });

  it('משטח בחירה בלי `apply` — הגזירה עדיין מצליחה', async () => {
    const { host, calls } = fakeDoc();
    host.ui = { selection: selectionSurface(READY_SELECTION) };
    setSystemClipboard(workingClipboard().api);

    const outcome = await cutSelection(host);

    expect(outcome.ok).toBe(true);
    expect(calls.map((call) => call.op)).toEqual(['clipboard.serializeSelection', 'delete']);
  });
});

describe('pasteFromClipboard', () => {
  it('קוראת מלוח המערכת, מפרקת, ומדביקה את התכנית שהתקבלה', async () => {
    const { host, calls } = fakeDoc();
    setSystemClipboard(readableClipboard({ 'text/html': '<p>טקסט</p>', 'text/plain': 'טקסט' }));

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(true);
    expect(calls.map((call) => call.op)).toEqual(['clipboard.parse', 'clipboard.insert']);
    // text/html לפני text/plain: הראשון הוא זה שנושא עיצוב, והמנוע בוחר
    // את הייצוג הנאמן ביותר מבין הפריטים.
    expect(calls[0]!.input).toEqual({
      source: 'browser',
      items: [
        { type: 'text/html', kind: 'string', data: '<p>טקסט</p>' },
        { type: 'text/plain', kind: 'string', data: 'טקסט' },
      ],
    });
    // ה-target הוא הבחירה החיה, ולא נשמט: `insert` בלי יעד מדביק בסוף המסמך.
    expect(calls[1]!.input).toEqual({ plan: enginePlan(), target: RANGE });
  });

  it('סמן בלי טווח הוא מקום הדבקה חוקי, ומגיע ל-`insert` כמו שהוא', async () => {
    // המצב הרגיל של הדבקה: המשתמש לחץ במסמך ולא סימן כלום. `empty: true` עם
    // `selectionTarget` הוא סמן — ולא „אין בחירה”.
    const { host, calls } = fakeDoc();
    host.ui = {
      selection: selectionSurface({ status: 'ready', empty: true, selectionTarget: CARET as never }),
    };
    setSystemClipboard(readableClipboard({ 'text/plain': 'טקסט' }));

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(true);
    const insertCall = calls.find((call) => call.op === 'clipboard.insert');
    expect((insertCall!.input as { target?: unknown }).target).toEqual(CARET);
  });

  it('אין סמן במסמך — נכשלת, ואינה מדביקה בסוף המסמך', async () => {
    // הכשל שהתיקון נולד ממנו: `insert` בלי `target` מחזיר `success: true`
    // ומכניס את התוכן אחרי הפסקה האחרונה. „נכשל בגלוי” הוא ההתנהגות הנכונה —
    // הדבקה שקטה במקום אחר היא הפתעה שאין ממנה חזרה חוץ מ„בטל”.
    const { host, calls } = fakeDoc();
    host.ui = { selection: selectionSurface({ status: 'ready', empty: true, selectionTarget: null }) };
    const clipboard = readableClipboard({ 'text/plain': 'טקסט' });
    setSystemClipboard(clipboard);

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('selection-required');
      expect(outcome.message).toBe('ההדבקה נכשלה: יש למקם את הסמן במסמך');
    }
    expect(calls).toEqual([]);
    // הבחירה נקראת לפני הלוח, ולכן הרשאת הלוח לא נדרשת בכלל במצב הזה.
    expect(clipboard.read).not.toHaveBeenCalled();
  });

  it('בחירה שטרם נקראה — „המסמך עדיין נטען”, ולא „לא מיקמת סמן”', async () => {
    const { host, calls } = fakeDoc();
    host.ui = { selection: selectionSurface({ status: 'pending', selectionTarget: null }) };
    setSystemClipboard(readableClipboard({ 'text/plain': 'טקסט' }));

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('not-ready');
      expect(outcome.message).toContain('המסמך עדיין נטען');
    }
    expect(calls).toEqual([]);
  });

  it('אין משטח בחירה — נכשלת, ולא נותנת למנוע לבחור את המקום', async () => {
    const { host, calls } = fakeDoc();
    host.ui = undefined;
    setSystemClipboard(readableClipboard({ 'text/plain': 'טקסט' }));

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.reason).toBe('host-capability-unavailable');
    expect(calls).toEqual([]);
  });

  it('תוכן שכבר הגיע ללוח המערכת אינו מודבק מהלוח הפנימי כשהקריאה נחסמה', async () => {
    // ההעתקה הצליחה, כלומר לוח המערכת מחזיק את התוכן — ומאותו רגע הוא המקור.
    // המשתמש עשוי להעתיק שם משהו אחר, או לבחור פריט אחר מהיסטוריית Win+V,
    // ואנחנו לא נדע. הדבקת הראי הפנימי כאן הייתה מחזירה תוכן ישן בלי לספר.
    const { host, calls } = fakeDoc();
    setSystemClipboard(workingClipboard().api);
    expect((await copySelection(host)).ok).toBe(true);

    setSystemClipboard(deniedClipboard());
    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('system-clipboard-blocked');
      expect(outcome.message).toContain('Ctrl+V');
    }
    expect(calls.some((call) => call.op === 'clipboard.insert')).toBe(false);
  });

  it('תמונה מהלוח נקראת כ-bytes, כפי שהוולידטור דורש', async () => {
    // `data instanceof Uint8Array` הוא תנאי מפורש במנוע; מחרוזת base64 כאן
    // הייתה נדחית ב-INVALID_INPUT.
    const png = new Uint8Array([137, 80, 78, 71]);
    const { host, calls } = fakeDoc();
    setSystemClipboard(readableClipboard({ 'image/png': png }));

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(true);
    const payload = calls[0]!.input as ClipboardPayload;
    expect(payload.items[0]!.kind).toBe('bytes');
    expect(payload.items[0]!.data).toBeInstanceOf(Uint8Array);
  });

  it('`read` שנדחה נופל ל-`readText`', async () => {
    const { host, calls } = fakeDoc();
    setSystemClipboard({
      read: vi.fn(() => Promise.reject(Object.assign(new Error('denied'), { name: 'NotAllowedError' }))),
      readText: vi.fn(async () => 'רק טקסט'),
    });

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(true);
    expect(calls[0]!.input).toEqual({
      source: 'browser',
      items: [{ type: 'text/plain', kind: 'string', data: 'רק טקסט' }],
    });
  });

  it('הרשאה שנדחתה לגמרי — נופלת ללוח הפנימי', async () => {
    // זה המסלול שמאפשר להעתיק ולהדביק בתוך התוסף גם כשלוח המערכת חסום לגמרי.
    const { host, calls } = fakeDoc();
    internalClipboard.write(enginePayload('מהתוסף'));
    setSystemClipboard(deniedClipboard());

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(true);
    expect(calls[0]!.input).toEqual(enginePayload('מהתוסף'));
  });

  it('העתק ואז הדבק דרך הלוח הפנימי מכניס בדיוק את מה שהועתק', async () => {
    // מסלול ה-round-trip: `file://` חוסם גם כתיבה וגם קריאה, ובלי הלוח הפנימי
    // שני הכפתורים היו חסרי תוחלת.
    const { host, calls } = fakeDoc();
    setSystemClipboard(deniedClipboard());

    const copied = await copySelection(host);
    const pasted = await pasteFromClipboard(host);

    expect(copied.ok).toBe(false); // ההודעה הכנה: לא הגיע ללוח המערכת
    expect(pasted.ok).toBe(true);
    const parseCall = calls.find((call) => call.op === 'clipboard.parse');
    expect(parseCall!.input).toEqual(enginePayload());
  });

  it('אין הרשאה ואין לוח פנימי — הודעה שמפנה ל-Ctrl+V', async () => {
    const { host, calls } = fakeDoc();
    setSystemClipboard(deniedClipboard());

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('system-clipboard-blocked');
      expect(outcome.message).toContain('Ctrl+V');
    }
    expect(calls).toEqual([]);
  });

  it('לוח שנקרא והיה ריק אינו מדביק תוכן ישן מהלוח הפנימי', async () => {
    // ההבחנה בין „נחסם” ל„ריק”: הדבקה של משהו שהועתק לפני חצי שעה בתוך
    // התוסף, בזמן שהמשתמש רואה לוח מערכת ריק, היא הפתעה ולא נוחות.
    const { host, calls } = fakeDoc();
    internalClipboard.write(enginePayload('ישן'));
    setSystemClipboard({ read: vi.fn(async () => []), readText: vi.fn(async () => '') });

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.message).toContain('לוח המערכת ריק');
    expect(calls).toEqual([]);
  });

  it('`parse` שנכשל אינו נוגע במסמך, וההסבר הוא הסיבה המדויקת', async () => {
    const { host, calls } = fakeDoc();
    const clipboard = host.activeEditor!.doc!.clipboard!;
    clipboard.parse = () => ({
      success: false,
      failure: {
        code: 'CAPABILITY_UNSUPPORTED',
        message: 'no faithful representation',
        details: { unsupportedReason: 'paste-no-faithful-representation' },
      },
    });
    setSystemClipboard(readableClipboard({ 'text/plain': 'א' }));

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.message).toBe('ההדבקה נכשלה: אין דרך להדביק את התוכן הזה בלי לשנות אותו');
      expect(outcome.reason).toBe('CAPABILITY_UNSUPPORTED');
    }
    expect(calls.map((call) => call.op)).not.toContain('clipboard.insert');
  });

  it('`parse` שנכשל בלי unsupportedReason מציג את קוד המנוע', async () => {
    const { host } = fakeDoc();
    host.activeEditor!.doc!.clipboard!.parse = () => ({
      success: false,
      failure: { code: 'EMPTY_FRAGMENT' },
    });
    setSystemClipboard(readableClipboard({ 'text/plain': 'א' }));

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.message).toBe('ההדבקה נכשלה: אין תוכן להדבקה');
  });

  it('`insert` עם success:false מגיע למשתמש עם הקוד', async () => {
    const { host } = fakeDoc();
    host.activeEditor!.doc!.clipboard!.insert = () => ({
      success: false,
      failure: { code: 'DOCUMENT_READONLY' },
    });
    setSystemClipboard(readableClipboard({ 'text/plain': 'א' }));

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('DOCUMENT_READONLY');
      expect(outcome.message).toBe('ההדבקה נכשלה: המסמך פתוח לקריאה בלבד');
    }
  });

  it('`insert` שזורק אינו מפיל את הרצועה', async () => {
    const { host } = fakeDoc();
    host.activeEditor!.doc!.clipboard!.insert = () => {
      throw new Error('boom');
    };
    setSystemClipboard(readableClipboard({ 'text/plain': 'א' }));

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('threw');
      expect(outcome.message).toBe('ההדבקה נכשלה: boom');
    }
  });

  it('גרסה שאין בה `parse` שולחת payload גולמי ל-`insert`', async () => {
    // החוזה מקבל בדיוק אחד מ-payload/plan/fragment, ולכן שליחת שניהם או
    // אפס הייתה זורקת. הכפיל מאמת את זה.
    const { host, calls } = fakeDoc();
    host.activeEditor!.doc!.clipboard!.parse = undefined;
    setSystemClipboard(readableClipboard({ 'text/plain': 'א' }));

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(true);
    expect(calls.map((call) => call.op)).toEqual(['clipboard.insert']);
    expect(calls[0]!.input).toEqual({
      payload: { source: 'browser', items: [{ type: 'text/plain', kind: 'string', data: 'א' }] },
      target: RANGE,
    });
  });

  it('אין `doc.clipboard.insert` — מנוטרל עם „אינו זמין בגרסה זו”, בלי לקרוא את הלוח', async () => {
    const { host } = fakeDoc({ clipboard: {} });
    const clipboard = readableClipboard({ 'text/plain': 'א' });
    setSystemClipboard(clipboard);

    const outcome = await pasteFromClipboard(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('command-unsupported');
      expect(outcome.message).toBe('ההדבקה נכשלה: אינו זמין בגרסה זו');
    }
    expect(clipboard.read).not.toHaveBeenCalled();
  });

  it('סובלת גם קריאות שמחזירות הבטחה', async () => {
    const { host } = fakeDoc();
    const clipboard = host.activeEditor!.doc!.clipboard!;
    clipboard.parse = () => Promise.resolve({ success: true, plan: enginePlan() });
    clipboard.insert = () => Promise.resolve({ success: true });
    setSystemClipboard(readableClipboard({ 'text/plain': 'א' }));

    expect((await pasteFromClipboard(host)).ok).toBe(true);
  });

  it('הסמן עובר לסוף מה שהודבק', async () => {
    // הבאג שזה נולד ממנו: המנוע אינו מזיז את הבחירה החיה אחרי `insert`,
    // וההדבקה הבאה קוראת את אותו סמן — „משה” ואז „כהן” יצא „כהןמשה”.
    const { host } = fakeDoc();
    const apply = vi.fn(() => ({ ok: true }));
    host.ui = { selection: { getSnapshot: () => READY_SELECTION, apply } };
    setSystemClipboard(readableClipboard({ 'text/plain': 'משה' }));

    expect((await pasteFromClipboard(host)).ok).toBe(true);
    // מכווץ לקצה, ולא הטווח שנוצר: אחרי הדבקה יש סמן, לא טקסט מסומן.
    expect(apply).toHaveBeenCalledWith(CARET_AFTER_PASTE);
  });

  it('קבלה בלי `effects` — הסמן נשאר במקומו במקום לזוז לניחוש', async () => {
    const { host } = fakeDoc();
    host.activeEditor!.doc!.clipboard!.insert = () => ({ success: true });
    const apply = vi.fn(() => ({ ok: true }));
    host.ui = { selection: { getSnapshot: () => READY_SELECTION, apply } };
    setSystemClipboard(readableClipboard({ 'text/plain': 'א' }));

    expect((await pasteFromClipboard(host)).ok).toBe(true);
    expect(apply).not.toHaveBeenCalled();
  });

  it('`apply` שנדחה כי הבלוק החדש טרם אותר — ניסיון חוזר, ואותו יעד', async () => {
    // הדבקה שיוצרת פסקה חדשה: הבלוק שבקבלה אינו ניתן לאיתור ברגע שההכנסה
    // חזרה, והמנוע מחזיר `target-unresolved`. נמדד — ההנמקה ב-clipboard.ts.
    const { host } = fakeDoc();
    const apply = vi
      .fn<(target: unknown) => { ok: boolean; reason?: string }>()
      .mockReturnValueOnce({ ok: false, reason: 'target-unresolved' })
      .mockReturnValue({ ok: true });
    host.ui = { selection: { getSnapshot: () => READY_SELECTION, apply } };
    setSystemClipboard(readableClipboard({ 'text/plain': 'א' }));

    expect((await pasteFromClipboard(host)).ok).toBe(true);
    expect(apply).toHaveBeenCalledTimes(2);
    expect(apply).toHaveBeenLastCalledWith(CARET_AFTER_PASTE);
  });

  it('`apply` שנכשל סגור אינו הופך הדבקה שהצליחה לכישלון, ואינו מנסה לנצח', async () => {
    // התוכן כבר במסמך. „ההדבקה נכשלה” בגלל סמן שלא זז הוא שקר למשתמש.
    const { host } = fakeDoc();
    const apply = vi.fn(() => ({ ok: false, reason: 'document-readonly' }));
    host.ui = { selection: { getSnapshot: () => READY_SELECTION, apply } };
    setSystemClipboard(readableClipboard({ 'text/plain': 'א' }));

    expect((await pasteFromClipboard(host)).ok).toBe(true);
    // ניסיון אחד ועוד שלושה חוזרים. חסם, ולא לולאה שממתינה לנס.
    expect(apply).toHaveBeenCalledTimes(4);
  });

  it('`apply` שזורק אינו מפיל את ההדבקה, ואינו חוזר על עצמו', async () => {
    // חריגה אינה „הבלוק עדיין לא נקלט” אלא מסלול שבור — אין מה לנסות שוב.
    const { host } = fakeDoc();
    const apply = vi.fn(() => {
      throw new Error('boom');
    });
    host.ui = { selection: { getSnapshot: () => READY_SELECTION, apply } };
    setSystemClipboard(readableClipboard({ 'text/plain': 'א' }));

    expect((await pasteFromClipboard(host)).ok).toBe(true);
    expect(apply).toHaveBeenCalledTimes(1);
  });

  it('משטח בחירה בלי `apply` — ההדבקה עדיין מצליחה', async () => {
    // גרסת מנוע שאינה חושפת את ה-helper: אין הזזת סמן, ואין כשל.
    const { host } = fakeDoc();
    setSystemClipboard(readableClipboard({ 'text/plain': 'א' }));

    expect((await pasteFromClipboard(host)).ok).toBe(true);
  });
});

const destinationDependencies = [{ target: '/word/styles.xml', dataBase64: btoa('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>') }];

describe('pasteFromClipboard — אפשרויות ההדבקה', () => {
  /** תכנית עשירה, כמו שה-`parse` בונה מ-HTML מעוצב. */
  function richPlan(): ClipboardPlan {
    return {
      fragment: {
        blocks: [
          {
            kind: 'heading',
            level: 1,
            styleRef: 'Heading1',
            runs: [
              { text: 'מודגש', marks: { bold: true, fontFamily: 'Arial', color: '#FF0000' } },
              { text: ' רגיל', marks: { fontFamily: 'Arial', color: '#FF0000' } },
            ],
          },
        ],
      },
      diagnostics: [],
    };
  }

  /** הכפיל, עם `parse` שמחזיר תכנית עשירה, פקודות שמדווחות David 18, ו-`docDefaults` של Arial 12. */
  function richDoc(caret: Record<string, { value?: unknown; active?: boolean }> = {}) {
    const fake = fakeDoc();
    fake.doc.clipboard!.parse = (payload) => {
      assertPayload(payload, 'clipboard.parse');
      fake.calls.push({ op: 'clipboard.parse', input: payload });
      return { success: true, plan: richPlan() };
    };
    const states: Record<string, { value?: unknown; active?: boolean }> = {
      'font-family': { value: 'David' },
      'font-size': { value: '18' },
      'text-color': { value: '#000000' },
      ...caret,
    };
    (fake.doc as Record<string, unknown>).styles = {
      apply: (_input: unknown, options?: { dryRun?: boolean }) => {
        // dryRun בלבד: „מזג עיצוב” קורא את ברירות המחדל, ולעולם אינו כותב אותן.
        if (options?.dryRun !== true) throw new Error('docDefaults must not be written by paste');
        return {
          success: true,
          before: { fontFamily: { ascii: 'Arial', hAnsi: 'Arial', cs: 'Arial' }, fontSize: 24, fontSizeCs: 24 },
        };
      },
    };
    fake.doc.info = () => ({ revision: '1' });
    fake.doc.selection = { extractOoxml: () => ({
      dependencies: destinationDependencies, fragment: { xml: '<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr><w:t>x</w:t></w:r>' },
    }) };
    fake.host.ui = {
      ...fake.host.ui,
      selection: selectionSurface({ status: 'ready', empty: true, selectionTarget: CARET as never }),
      commands: {
        has: (id: string) => id in states,
        get: (id: string) => ({ getState: () => states[id] }),
      },
    };
    return fake;
  }

  const insertedPlan = (calls: FakeCall[]) =>
    (calls.find((call) => call.op === 'clipboard.insert')!.input as { plan: ClipboardPlan }).plan;

  it('„שמירה על עיצוב המקור” מעבירה את התכנית כמו שהיא', async () => {
    const { host, calls } = richDoc();
    setSystemClipboard(readableClipboard({ 'text/html': '<h1>x</h1>', 'text/plain': 'x' }));

    expect((await pasteFromClipboard(host, { mode: 'keep-source' })).ok).toBe(true);
    expect(insertedPlan(calls)).toEqual(richPlan());
  });

  it('„מזג עיצוב”: עיצוב היעד (David 18), ההדגשה של המקור — ולא הכותרת, הגופן או הצבע שלו', async () => {
    const { host, calls } = richDoc();
    setSystemClipboard(readableClipboard({ 'text/html': '<h1>x</h1>', 'text/plain': 'x' }));

    expect((await pasteFromClipboard(host, { mode: 'merge' })).ok).toBe(true);
    const target = { fontFamily: 'David', fontSizePt: 18, fontSizeCsPt: 18, color: '#000000', bold: false, boldCs: false, italic: false, italicCs: false, underline: false, strike: false };
    expect(insertedPlan(calls).fragment.blocks).toEqual([
      {
        kind: 'paragraph',
        runs: [
          { text: 'מודגש', marks: { ...target, bold: true } },
          { text: ' רגיל', marks: target },
        ],
      },
    ]);
  });

  it('„מזג עיצוב” בטקסט שלא עוצב: ברירות המחדל אינן נכתבות כעיצוב ישיר', async () => {
    const { host, calls } = richDoc({ 'font-family': { value: 'Arial' }, 'font-size': { value: '12' } });
    setSystemClipboard(readableClipboard({ 'text/html': '<h1>x</h1>', 'text/plain': 'x' }));

    await pasteFromClipboard(host, { mode: 'merge' });
    const blocks = insertedPlan(calls).fragment.blocks as Array<{ runs: Array<{ marks?: unknown }> }>;
    const destination = { fontFamily: 'Arial', fontSizePt: 12, fontSizeCsPt: 12, color: '#000000', bold: false, boldCs: false, italic: false, italicCs: false, underline: false, strike: false };
    expect(blocks[0]!.runs.map((run) => run.marks)).toEqual([{ ...destination, bold: true }, destination]);
  });

  it('„שמירת טקסט בלבד” שולחת למנוע `text/plain` בלבד', async () => {
    const { host, calls } = richDoc();
    setSystemClipboard(readableClipboard({ 'text/html': '<h1>x</h1>', 'text/plain': 'x', 'image/png': new Uint8Array([1]) }));

    expect((await pasteFromClipboard(host, { mode: 'text-only' })).ok).toBe(true);
    expect(calls[0]!.input).toEqual({ source: 'browser', items: [{ type: 'text/plain', kind: 'string', data: 'x' }] });
    const blocks = insertedPlan(calls).fragment.blocks as Array<{ runs: Array<{ marks?: unknown }> }>;
    // גם ההדגשה של המקור נזרקת; נשאר רק עיצוב היעד.
    for (const run of blocks[0]!.runs) expect(run.marks).toEqual({ fontFamily: 'David', fontSizePt: 18, fontSizeCsPt: 18, color: '#000000', bold: false, boldCs: false, italic: false, italicCs: false, underline: false, strike: false });
  });

  it('תכנית `plain-text` משוכתבת נשלחת כ-`fragment` ולא כ-`plan`', async () => {
    // superdoc 2.23: `insert({plan})` של תכנית plain-text מתעלם מה-marks
    // ומעתיק את הריצה שבסמן. נמדד: David 18 נכתב ב-2.20 ולא ב-2.23, ו-
    // `insert({fragment})` כותב אותו בשתיהן.
    const { host, calls } = richDoc();
    host.activeEditor!.doc!.clipboard!.parse = (payload) => {
      calls.push({ op: 'clipboard.parse', input: payload });
      return {
        success: true,
        plan: {
          fragment: { blocks: [{ kind: 'paragraph', runs: [{ text: 'x' }] }] },
          diagnostics: [],
          summary: { sourceKind: 'plain-text', plainFallback: false },
        },
      };
    };
    setSystemClipboard(readableClipboard({ 'text/plain': 'x' }));

    expect((await pasteFromClipboard(host, { mode: 'text-only' })).ok).toBe(true);
    const sent = calls.find((call) => call.op === 'clipboard.insert')!.input as {
      plan?: unknown;
      fragment?: { blocks: Array<{ runs: Array<{ marks?: unknown }> }> };
    };
    expect(sent.plan).toBeUndefined();
    expect(sent.fragment!.blocks[0]!.runs[0]!.marks).toMatchObject({ fontFamily: 'David', fontSizePt: 18, fontSizeCsPt: 18 });
  });

  it('„שמירת טקסט בלבד” על לוח בלי טקסט — נכשלת בלי לגעת במסמך', async () => {
    const { host, calls } = richDoc();
    setSystemClipboard(readableClipboard({ 'image/png': new Uint8Array([1, 2]) }));

    const outcome = await pasteFromClipboard(host, { mode: 'text-only' });
    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.message).toBe('ההדבקה נכשלה: אין בלוח טקסט להדבקה');
    expect(calls).toEqual([]);
  });

  it('payload שנמסר (מאירוע Ctrl+V) מודבק בלי לקרוא את לוח המערכת', async () => {
    const { host, calls } = richDoc();
    const clipboard = readableClipboard({ 'text/plain': 'אחר' });
    setSystemClipboard(clipboard);
    const payload: ClipboardPayload = { source: 'browser', items: [{ type: 'text/plain', kind: 'string', data: 'מהאירוע' }] };

    expect((await pasteFromClipboard(host, { mode: 'merge', payload })).ok).toBe(true);
    expect(clipboard.read).not.toHaveBeenCalled();
    expect(calls[0]!.input).toEqual(payload);
  });
});

describe('paste destination safety', () => {
  function transientHost({ changeTarget = false, changeRevision = false } = {}) {
    const { host, doc, calls } = fakeDoc();
    let reads = 0;
    let revision = '1';
    host.ui = { selection: { getSnapshot: () => {
      reads += 1;
      if (changeRevision && reads > 2) revision = '2';
      return { status: reads <= 2 ? 'stale' : 'ready', empty: true, selectionTarget: (changeTarget && reads > 2 ? { ...CARET, start: { ...CARET.start, offset: 4 }, end: { ...CARET.end, offset: 4 } } : CARET) as never };
    } } };
    doc.info = () => ({ revision });
    doc.selection = { extractOoxml: () => ({ dependencies: destinationDependencies, fragment: { xml: '<w:r><w:t>x</w:t></w:r>' } }) };
    return { host, calls };
  }
  it('waits for transient stale state at the same caret, then inserts once', async () => {
    const { host, calls } = transientHost();
    expect((await pasteFromClipboard(host, { mode: 'merge', payload: enginePayload('x') })).ok).toBe(true);
    expect(calls.filter((call) => call.op === 'clipboard.insert')).toHaveLength(1);
  });
  it('refuses before insertion if settling reveals a different caret or revision', async () => {
    for (const options of [{ changeTarget: true }, { changeRevision: true }]) {
      const { host, calls } = transientHost(options);
      expect((await pasteFromClipboard(host, { mode: 'merge', payload: enginePayload('x') })).ok).toBe(false);
      expect(calls.some((call) => call.op === 'clipboard.insert')).toBe(false);
    }
  });
  it('refuses a caret change while the final revision read is pending', async () => {
    const { host, doc, calls } = fakeDoc();
    let target: SelectionSnapshot['selectionTarget'] = CARET as never;
    host.ui = { selection: { getSnapshot: () => ({ status: 'ready', empty: true, selectionTarget: target }) } };
    let reads = 0;
    let resolveFinal!: (info: { revision: string }) => void;
    let requested!: () => void;
    const finalRequested = new Promise<void>((resolve) => { requested = resolve; });
    doc.info = () => {
      reads += 1;
      if (reads !== 2) return { revision: '1' };
      const pending = new Promise<{ revision: string }>((resolve) => { resolveFinal = resolve; });
      requested();
      return pending;
    };
    doc.selection = { extractOoxml: () => ({ dependencies: destinationDependencies, fragment: { xml: '<w:r><w:t>x</w:t></w:r>' } }) };
    const pending = pasteFromClipboard(host, { mode: 'merge', payload: enginePayload('x') });
    await finalRequested;
    target = { ...CARET, start: { ...CARET.start, offset: 4 }, end: { ...CARET.end, offset: 4 } } as never;
    resolveFinal({ revision: '1' });
    expect((await pending).ok).toBe(false);
    expect(calls.some((call) => call.op === 'clipboard.insert')).toBe(false);
  });

  it('refuses a caret change while the initial settled revision read is pending', async () => {
    const { host, doc, calls } = fakeDoc();
    let target: SelectionSnapshot['selectionTarget'] = CARET as never;
    let snapshots = 0;
    host.ui = { selection: { getSnapshot: () => ({ status: ++snapshots === 1 ? 'stale' : 'ready', empty: true, selectionTarget: target }) } };
    let reads = 0;
    let resolveAfter!: (info: { revision: string }) => void;
    let requested!: () => void;
    const afterRequested = new Promise<void>((resolve) => { requested = resolve; });
    doc.info = () => {
      if (++reads !== 2) return { revision: '1' };
      const pending = new Promise<{ revision: string }>((resolve) => { resolveAfter = resolve; });
      requested();
      return pending;
    };
    doc.selection = { extractOoxml: vi.fn() };
    const pending = pasteFromClipboard(host, { mode: 'merge', payload: enginePayload('x') });
    await afterRequested;
    target = { ...CARET, start: { ...CARET.start, offset: 4 }, end: { ...CARET.end, offset: 4 } } as never;
    resolveAfter({ revision: '1' });
    expect((await pending).ok).toBe(false);
    expect(doc.selection.extractOoxml).not.toHaveBeenCalled();
    expect(calls.some((call) => call.op === 'clipboard.insert')).toBe(false);
  });

  it('refuses mixed or selected ranges without even reading the clipboard', async () => {
    const { host, calls } = fakeDoc();
    const clipboard = readableClipboard({ 'text/plain': 'x' });
    setSystemClipboard(clipboard);
    const outcome = await pasteFromClipboard(host, { mode: 'merge' });
    expect(outcome.ok).toBe(false);
    expect(clipboard.read).not.toHaveBeenCalled();
    expect(calls).toEqual([]);
  });
  it('gives Hebrew and Latin their own destination font and size when the two differ', async () => {
    const { host, doc, calls } = fakeDoc();
    host.ui = { selection: selectionSurface({ status: 'ready', empty: true, selectionTarget: CARET as never }) };
    doc.info = () => ({ revision: '1' });
    doc.selection = { extractOoxml: () => ({ dependencies: destinationDependencies, fragment: { xml: '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="David"/><w:szCs w:val="28"/><w:rtl/></w:rPr><w:t>x</w:t></w:r>' } }) };
    doc.clipboard!.parse = () => ({ success: true, plan: { fragment: { blocks: [{ kind: 'paragraph', runs: [{ text: 'שלום abc' }] }] }, diagnostics: [] } });
    const outcome = await pasteFromClipboard(host, { mode: 'merge', payload: enginePayload('שלום abc') });
    expect(outcome.ok).toBe(true);
    const insert = calls.find((call) => call.op === 'clipboard.insert')?.input as { plan: { fragment: { blocks: { runs: { text: string; marks: Record<string, unknown> }[] }[] } } };
    expect(insert.plan.fragment.blocks[0]!.runs.map((run) => [run.text, run.marks.fontFamily, run.marks.fontSizePt, run.marks.fontSizeCsPt])).toEqual([
      ['שלום ', 'David', 14, 14],
      ['abc', 'Arial', 12, 12],
    ]);
    expect(JSON.stringify(insert)).not.toContain('fontFamilyCs');
  });
  it('refuses unsupported script fonts before inserting, and logs why', async () => {
    const { host, doc, calls } = fakeDoc();
    host.ui = { selection: selectionSurface({ status: 'ready', empty: true, selectionTarget: CARET as never }) };
    doc.info = () => ({ revision: '1' });
    doc.selection = { extractOoxml: () => ({ dependencies: destinationDependencies, fragment: { xml: '<w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsi="Arial" w:cs="David"/></w:rPr><w:t>x</w:t></w:r>' } }) };
    const errors = vi.spyOn(console, 'error').mockImplementation(() => {});
    const outcome = await pasteFromClipboard(host, { mode: 'merge', payload: enginePayload('x') });
    expect(outcome.ok).toBe(false);
    expect(calls.some((call) => call.op === 'clipboard.insert')).toBe(false);
    expect(errors).toHaveBeenCalledWith(expect.stringContaining('Direct theme font unsupported'));
    errors.mockRestore();
  });
  it('refuses an async paste if the document changes while parsing with the same selection', async () => {
    const { host, doc, calls } = fakeDoc();
    host.ui = { selection: selectionSurface({ status: 'ready', empty: true, selectionTarget: CARET as never }) };
    let revision = '1';
    doc.info = () => ({ revision });
    doc.selection = { extractOoxml: () => ({ dependencies: destinationDependencies, fragment: { xml: '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr><w:t>x</w:t></w:r>' } }) };
    doc.clipboard!.parse = async () => { revision = '2'; return { success: true, plan: enginePlan() }; };
    const outcome = await pasteFromClipboard(host, { mode: 'merge', payload: enginePayload('x') });
    expect(outcome.ok).toBe(false);
    expect(calls.some((call) => call.op === 'clipboard.insert')).toBe(false);
  });
});

describe('selectWholeDocument', () => {
  /** `ranges.resolve` שמאמת את העוגנים ומחזיר טווח שמכסה את הגוף. */
  function rangesHost(overrides: Partial<ClipboardDocumentApi> = {}) {
    const calls: FakeCall[] = [];
    const applied: unknown[] = [];

    const doc: ClipboardDocumentApi = {
      ranges: {
        resolve: (input) => {
          // הוולידטור של המנוע דורש עוגן עם `kind` מוכר. `{kind:'document'}`
          // בלי `edge`, או `edge` שאינו start/end, נדחים.
          for (const anchor of [input?.start, input?.end]) {
            if (anchor?.kind !== 'document' || (anchor.edge !== 'start' && anchor.edge !== 'end')) {
              throw new Error('ranges.resolve requires a valid RangeAnchor');
            }
          }
          calls.push({ op: 'ranges.resolve', input });
          return { target: RANGE as never };
        },
      },
      ...overrides,
    };

    const host: ClipboardHost = {
      activeEditor: { doc },
      ui: {
        selection: {
          getSnapshot: () => READY_SELECTION,
          apply: (target) => {
            if (!isSelectionTarget(target)) throw new Error('apply requires a SelectionTarget');
            applied.push(target);
            return { ok: true };
          },
        },
      },
    };

    return { host, calls, applied };
  }

  it('פותרת את גבולות המסמך ומחילה את הטווח על העורך', async () => {
    const { host, calls, applied } = rangesHost();

    const outcome = await selectWholeDocument(host);

    expect(outcome.ok).toBe(true);
    expect(calls[0]!.input).toEqual({
      start: { kind: 'document', edge: 'start' },
      end: { kind: 'document', edge: 'end' },
    });
    expect(applied).toEqual([RANGE]);
  });

  it('אינה נוגעת ב-DOM: אין קריאה ל-window.getSelection', async () => {
    // זו הרגרסיה עצמה — `selectAllChildren(document.body)` סימן את הרצועה
    // ואת שורת המצב, ולא את המסמך.
    const { host } = rangesHost();
    const spy = vi.spyOn(window, 'getSelection');

    await selectWholeDocument(host);

    expect(spy).not.toHaveBeenCalled();
  });

  it('אין `ranges.resolve` — מנוטרל עם „אינו זמין בגרסה זו”', async () => {
    const { host } = rangesHost({ ranges: undefined });

    const outcome = await selectWholeDocument(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('command-unsupported');
      expect(outcome.message).toBe('בחירת כל המסמך נכשלה: אינו זמין בגרסה זו');
    }
  });

  it('אין `ui.selection.apply` — נכשלת לפני שפותרת טווח', async () => {
    const { host, calls } = rangesHost();
    host.ui = { selection: { getSnapshot: () => READY_SELECTION } };

    const outcome = await selectWholeDocument(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('host-capability-unavailable');
      expect(outcome.message).toBe('בחירת כל המסמך נכשלה: היכולת הדרושה לפעולה אינה זמינה');
    }
    expect(calls).toEqual([]);
  });

  it('`apply` שנכשל סגור מדווח את ה-reason של המנוע', async () => {
    const { host } = rangesHost();
    host.ui = {
      selection: {
        getSnapshot: () => READY_SELECTION,
        apply: () => ({ ok: false, reason: 'document-readonly' }),
      },
    };

    const outcome = await selectWholeDocument(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.message).toBe('בחירת כל המסמך נכשלה: המסמך פתוח לקריאה בלבד');
  });

  it('טווח שלא נפתר אינו נשלח ל-apply', async () => {
    for (const target of [undefined, null, 'body']) {
      const { host, applied } = rangesHost({
        ranges: { resolve: () => ({ target: target as never }) },
      });

      const outcome = await selectWholeDocument(host);

      expect(outcome.ok, String(target)).toBe(false);
      if (!outcome.ok) expect(outcome.reason).toBe('target-unresolved');
      expect(applied).toEqual([]);
    }
  });

  it('`resolve` שזורק אינו מפיל את הרצועה', async () => {
    const { host } = rangesHost({
      ranges: {
        resolve: () => {
          throw new Error('boom');
        },
      },
    });

    const outcome = await selectWholeDocument(host);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.reason).toBe('threw');
      expect(outcome.message).toBe('בחירת כל המסמך נכשלה: boom');
    }
  });

  it('סובלת גם קריאות שמחזירות הבטחה', async () => {
    const { host } = rangesHost({
      ranges: { resolve: () => Promise.resolve({ target: RANGE as never }) },
    });
    const apply = vi.fn(() => Promise.resolve({ ok: true }));
    host.ui = { selection: { getSnapshot: () => READY_SELECTION, apply } };

    expect((await selectWholeDocument(host)).ok).toBe(true);
    expect(apply).toHaveBeenCalledWith(RANGE);
  });

  it('אין Document API — כשל סגור, בלי חריגה', async () => {
    for (const host of [null, undefined, {}, { activeEditor: { doc: null } }]) {
      const outcome = await selectWholeDocument(host as ClipboardHost);

      expect(outcome.ok).toBe(false);
      if (!outcome.ok) expect(outcome.reason).toBe('command-unsupported');
    }
  });
});

/**
 * שער על המקור עצמו, בנוסף לבדיקות ההתנהגות.
 *
 * ההתנהגות נבדקת מול כפילים, אבל מה שהיה ב-HomeTab היה **היעדר** התנהגות:
 * פונקציה ריקה, ובחירה שמסמנת את הממשק. שני אלה הם JavaScript תקין לחלוטין
 * ואף typecheck אינו מתלונן עליהם, ולכן הם נבדקים כאן — כמו ב-
 * tests/unit/tab-controls.test.ts.
 */
describe('קבוצת „לוח” ב-HomeTab', () => {
  const source = readFileSync(join(process.cwd(), 'src/ui/ribbon/tabs/HomeTab.vue'), 'utf8');

  it('אין בחירה דרך ה-DOM', () => {
    // `selectAllChildren(document.body)` סימן את הרצועה ואת שורת המצב.
    expect(source).not.toContain('selectAllChildren');
    expect(source).not.toContain('getSelection');
  });

  it('אין קריאה ישירה ל-navigator.clipboard מהקומפוננטה', () => {
    // הלוח עובר בשכבה שאפשר לבדוק, ולא כ-`void ...readText()` שזורק את התוצאה.
    expect(source).not.toContain('navigator.clipboard');
  });

  it('לכל אחד מארבעת הפקדים יש מטפל וגם חיווט של disabled', () => {
    for (const label of ['גזור', 'העתק', 'בחר הכל']) {
      const control = source.match(new RegExp(`<RibbonButton\\b[^>]*?label="${label}"[\\s\\S]*?/>`));
      expect(control, label).not.toBeNull();
      expect(control![0], label).toMatch(/@click=/);
      expect(control![0], label).toMatch(/:disabled=/);
    }
    // „הדבק” מפוצל, כמו ב-Word: הגוף מדביק (`@action`) והחץ פותח את
    // אפשרויות ההדבקה (`@select`).
    const paste = source.match(/<RibbonMenuButton\b[^>]*?label="הדבק"[\s\S]*?\/>/);
    expect(paste).not.toBeNull();
    expect(paste![0]).toMatch(/@action=/);
    expect(paste![0]).toMatch(/@select=/);
    expect(paste![0]).toMatch(/:disabled=/);
  });

  it('הפקד עצמו מונע מהלחיצה לגזול את המיקוד מהעורך', () => {
    // בלי `@pointerdown.prevent` ב-RibbonButton, לחיצה על „העתק” מעבירה את
    // המיקוד לכפתור — והבחירה במסמך, שהיא כל מה שיש להעתיק, אובדת.
    const button = readFileSync(
      join(process.cwd(), 'src/ui/ribbon/common/RibbonButton.vue'),
      'utf8',
    );
    expect(button).toContain('@pointerdown.prevent');
  });
});
