/**
 * הגנת מסמך (גל 19). הבדיקה: מסלול ההפעלה שולח mode:'readOnly', הביטול
 * נקי, NO_OP הצלחה, וכשלים מתורגמים — הסדר שהמדריך מחייב (ביטול לפני
 * הפעלה) נמדד בדפדפן.
 *
 * ## הבאג שהתגלה אחרי המדידה בדפדפן: „הגבל עריכה” אינו נועל כלום
 *
 * `setEditingRestriction` כותב `w:documentProtection` נכון, אבל המנוע שוער
 * חסימת קלט ופקודות (`bold` ודומיו) **אך ורק** לפי `document-mode ===
 * 'viewing'` — לא לפי `editingRestriction.enforced`. לכן `enableReadOnlyProtection`
 * ו-`disableProtection` מקבלות `CommandAdapter` ומעבירות את `document-mode`
 * יחד עם כתיבת ה-XML. הבדיקות בקבוצה „חסימה בפועל” למטה מוכיחות את זה ישירות
 * מול כפיל אדפטר, בלי דפדפן.
 */
import { describe, expect, it } from 'vitest';
import type { CommandAdapter, CommandOutcome } from '../../src/engine/command-adapter';
import type { EditorSession } from '../../src/engine/create-editor';
import { createEditorInteraction } from '../../src/engine/editor-interaction';
import {
  disableProtection,
  enableReadOnlyProtection,
  readProtectionState,
  syncProtectionRuntime,
} from '../../src/engine/protection';

function fakeDoc(options: {
  get?: unknown;
  set?: () => unknown;
  clear?: () => unknown;
} = {}) {
  const calls = new Map<string, unknown[]>();
  const make = (name: string, fallback: () => unknown, override?: () => unknown) => (
    input: unknown,
  ) => {
    calls.set(name, [...(calls.get(name) ?? []), input]);
    return (override ?? fallback)() as never;
  };

  const doc = {
    protection: {
      get: make('get', () => options.get ?? { editingRestriction: { mode: 'none', enforced: false } }),
      setEditingRestriction: make('set', () => ({ success: true }), options.set),
      clearEditingRestriction: make('clear', () => ({ success: true }), options.clear),
    },
  } as never;

  return { doc, calls, host: { activeEditor: { doc } } };
}

/**
 * כפיל `CommandAdapter` צר: מחזיק רק את `document-mode`, ומדווח את מה
 * ש-`run` קבע — בדיוק ה"מנוע" ש-`enable/disableProtection` צריכות כדי
 * לבצע את החסימה בפועל. `mode()` חושף את המצב הנוכחי לבדיקה.
 */
function fakeCommandAdapter(initialMode: string): { adapter: CommandAdapter; mode: () => string; setMode: (mode: string) => void; payloads: unknown[] } {
  let mode = initialMode;
  const payloads: unknown[] = [];

  const adapter: CommandAdapter = {
    has: (id) => id === 'document-mode',
    getState: (id) => ({
      supported: true,
      enabled: true,
      active: false,
      value: id === 'document-mode' ? mode : undefined,
    }),
    observe: () => () => {},
    async run(id, payload): Promise<CommandOutcome> {
      payloads.push({ id, payload });
      if (id === 'document-mode') {
        mode = (payload as { mode: string }).mode;
      }
      return { ok: true };
    },
  };

  return { adapter, mode: () => mode, setMode: value => { mode = value; }, payloads };
}

describe('enableReadOnlyProtection', () => {
  it('שולח mode readOnly', async () => {
    const { host, calls } = fakeDoc();

    await expect(enableReadOnlyProtection(host, null)).resolves.toMatchObject({ ok: true });
    expect(calls.get('set')?.[0]).toEqual({ mode: 'readOnly' });
  });

  it('NO_OP = כבר מוגן = הצלחה', async () => {
    const { host } = fakeDoc({
      set: () => ({ success: false, failure: { code: 'NO_OP' } }),
    });

    await expect(enableReadOnlyProtection(host, null)).resolves.toMatchObject({ ok: true });
  });

  it('אין set במנוע — הנוסח של §12', async () => {
    const host = { activeEditor: { doc: {} } };
    const outcome = await enableReadOnlyProtection(host as never, null);

    expect(outcome).toMatchObject({ ok: false, reason: 'command-unsupported' });
  });
});

describe('disableProtection', () => {
  it('clear נקרא ללא קלט', async () => {
    const { host, calls } = fakeDoc();

    await expect(disableProtection(host, null)).resolves.toEqual({ ok: true });
    expect(calls.get('clear')).toHaveLength(1);
  });

  it('כשל בביטול מתורגם — ולא נבלע', async () => {
    const { host } = fakeDoc({
      clear: () => ({ success: false, failure: { code: 'DOCUMENT_READONLY', message: 'locked' } }),
    });

    const outcome = await disableProtection(host, null);

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) expect(outcome.reason).toBe('DOCUMENT_READONLY');
  });
});

describe('readProtectionState', () => {
  it('מצב מוחזר מהמנוע מנורמל, כולל runtimeEnforced', async () => {
    const { host } = fakeDoc({
      get: { editingRestriction: { mode: 'readOnly', enforced: true, runtimeEnforced: true } },
    });

    const state = await readProtectionState(host);

    expect(state).toEqual({ mode: 'readOnly', enforced: true, runtimeEnforced: true });
  });

  it('get שזורק → null ולא זריקה', async () => {
    const throwingHost = {
      activeEditor: {
        doc: {
          protection: {
            get: () => {
              throw new Error('boom');
            },
          },
        },
      },
    };

    await expect(readProtectionState(throwingHost as never)).resolves.toBeNull();
  });

  it('הדגל נקרא מהמסמך בכל קריאה, לא מדגל מקומי שנשמר בצד', async () => {
    // המנוע (כאן: המסמך של הכפיל) הוא היחיד שמחזיק את התשובה. אם protection.ts
    // היה זוכר תשובה קודמת, הקריאה השנייה הייתה מחזירה את הראשונה למרות
    // שהמסמך עצמו השתנה.
    let enforced = false;
    const doc = {
      protection: {
        get: () => ({
          editingRestriction: { mode: enforced ? 'readOnly' : 'none', enforced, runtimeEnforced: enforced },
        }),
      },
    } as never;
    const host = { activeEditor: { doc } };

    expect(await readProtectionState(host as never)).toEqual({
      mode: 'none',
      enforced: false,
      runtimeEnforced: false,
    });

    enforced = true;
    expect(await readProtectionState(host as never)).toEqual({
      mode: 'readOnly',
      enforced: true,
      runtimeEnforced: true,
    });
  });
});

describe('חסימה בפועל: document-mode יחד עם ה-XML', () => {
  it('serializes a delayed protection write with opening and refuses unprotect while leased', async () => {
    let finishSet!: (receipt: { success: boolean }) => void;
    const pendingSet = new Promise<{ success: boolean }>(resolve => { finishSet = resolve; });
    const { host, calls } = fakeDoc({ set: () => pendingSet });
    const commands = fakeCommandAdapter('suggesting');
    const interaction = createEditorInteraction({
      superdoc: { setDocumentMode: commands.setMode }, container: document.createElement('div'), onDispose: () => {},
    } as unknown as EditorSession, commands.adapter);

    const enabling = enableReadOnlyProtection(host, interaction.commands);
    expect(interaction.pendingModeChanges.value).toBe(1);
    expect(() => interaction.block()).toThrow('שינוי הגנת המסמך עדיין מתבצע');
    expect(commands.mode()).toBe('suggesting');
    finishSet({ success: true });
    expect(await enabling).toEqual({ ok: true, previousMode: 'suggesting' });
    expect(interaction.pendingModeChanges.value).toBe(0);
    expect(commands.mode()).toBe('viewing');

    const release = interaction.block();
    expect(await disableProtection(host, interaction.commands, 'suggesting')).toMatchObject({ ok: false, reason: 'not-ready' });
    expect(await enableReadOnlyProtection(host, interaction.commands)).toMatchObject({ ok: false, reason: 'not-ready' });
    expect(calls.get('clear')).toBeUndefined();
    expect(calls.get('set')).toHaveLength(1);
    release();
    expect(commands.mode()).toBe('viewing');
    expect(await disableProtection(host, interaction.commands, 'suggesting')).toEqual({ ok: true });
    expect(commands.mode()).toBe('suggesting');
  });

  it('הפעלה מעבירה document-mode ל-viewing — זה מה שחוסם קלט ופקודות בפועל', async () => {
    const { host } = fakeDoc();
    const commands = fakeCommandAdapter('editing');

    const outcome = await enableReadOnlyProtection(host, commands.adapter);

    expect(outcome).toEqual({ ok: true, previousMode: 'editing' });
    expect(commands.mode()).toBe('viewing');
    expect(commands.payloads).toEqual([{ id: 'document-mode', payload: { mode: 'viewing' } }]);
  });

  it('ביטול משחזר את המצב שנשמר — לא "editing" גורף', async () => {
    const { host } = fakeDoc();
    const commands = fakeCommandAdapter('viewing');

    const outcome = await disableProtection(host, commands.adapter, 'suggesting');

    expect(outcome).toEqual({ ok: true });
    expect(commands.mode()).toBe('suggesting');
  });

  it('סבב מלא: מסמך שהיה במעקב שינויים (suggesting) חוזר למעקב, לא ל-editing', async () => {
    const { host } = fakeDoc();
    const commands = fakeCommandAdapter('suggesting');

    const enableOutcome = await enableReadOnlyProtection(host, commands.adapter);
    expect(enableOutcome).toEqual({ ok: true, previousMode: 'suggesting' });
    expect(commands.mode()).toBe('viewing');

    if (!enableOutcome.ok) throw new Error('unreachable');
    const disableOutcome = await disableProtection(host, commands.adapter, enableOutcome.previousMode);

    expect(disableOutcome).toEqual({ ok: true });
    expect(commands.mode()).toBe('suggesting');
  });

  it('כשאין CommandAdapter (commands: null) — רק ה-XML נכתב, אין נגיעה במצב', async () => {
    const { host } = fakeDoc();

    const outcome = await enableReadOnlyProtection(host, null);

    expect(outcome).toEqual({ ok: true, previousMode: 'editing' });
  });
});

describe('syncProtectionRuntime', () => {
  it('reserves a delayed protection read so a temporary viewing lease cannot mask permanent protection', async () => {
    let finishRead!: (value: unknown) => void;
    const reply = new Promise(resolve => { finishRead = resolve; });
    const { host, calls } = fakeDoc({ get: reply });
    const commands = fakeCommandAdapter('editing');
    const interaction = createEditorInteraction({
      superdoc: { setDocumentMode: commands.setMode }, container: document.createElement('div'), onDispose: () => {},
    } as unknown as EditorSession, commands.adapter);
    const syncing = syncProtectionRuntime(host, interaction.commands);
    expect(interaction.pendingModeChanges.value).toBe(1);
    expect(() => interaction.block()).toThrow('שינוי הגנת המסמך עדיין מתבצע');
    expect(commands.mode()).toBe('editing');
    finishRead({ editingRestriction: { mode: 'readOnly', enforced: true } });
    expect(await syncing).toMatchObject({ enforced: true });
    expect(interaction.pendingModeChanges.value).toBe(0);
    expect(commands.mode()).toBe('viewing');
    const release = interaction.block();
    expect(await syncProtectionRuntime(host, interaction.commands)).toBeNull();
    expect(calls.get('get')).toHaveLength(1);
    release();
    expect(commands.mode()).toBe('viewing');
  });

  it('מסמך שנטען כשהוא כבר מוגן: כופה viewing גם בלי הפעלה בסשן הזה', async () => {
    const { host } = fakeDoc({
      get: { editingRestriction: { mode: 'readOnly', enforced: true } },
    });
    const commands = fakeCommandAdapter('editing');

    const state = await syncProtectionRuntime(host, commands.adapter);

    expect(state).toMatchObject({ enforced: true });
    expect(commands.mode()).toBe('viewing');
  });

  it('מסמך שאינו מוגן: לא נוגעת ב-document-mode', async () => {
    const { host } = fakeDoc();
    const commands = fakeCommandAdapter('suggesting');

    await syncProtectionRuntime(host, commands.adapter);

    expect(commands.mode()).toBe('suggesting');
    expect(commands.payloads).toEqual([]);
  });
});
