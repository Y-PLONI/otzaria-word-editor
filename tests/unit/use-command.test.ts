/**
 * ה-composable שכל פקד ב-Ribbon עובר דרכו.
 *
 * מה שנבדק כאן הוא הדבר שהיה שבור: `run()` החזירה תוצאה עם הודעה בעברית, וכל
 * 38 אתרי הקריאה עשו `void cmd.run()` וזרקו אותה. לכן שלוש טבלאות התרגום
 * ב-command-adapter.ts היו קוד מת, וכשל פקודה נראה למשתמש כמו כפתור שבור.
 * מעכשיו הדיווח הוא חלק מהחוזה של `run()` ולא באחריות אתר הקריאה.
 *
 * ההרכבה היא `createApp` על div ב-jsdom ולא @vue/test-utils: `inject`
 * ו-`onUnmounted` דורשים הקשר קומפוננטה אמיתי, וזה מספיק כדי לקבל אותו בלי
 * תלות חדשה.
 */
import { describe, expect, it, vi } from 'vitest';
import { createApp, defineComponent, h, ref, type App, type Ref } from 'vue';
import { useCommand } from '../../src/composables/useCommand';
import { COMMAND_ADAPTER, COMMAND_REPORTER, DOCUMENT_INTERACTION_BLOCKED, type CommandReporter } from '../../src/composables/keys';
import type { CommandAdapter, CommandOutcome } from '../../src/engine/command-adapter';

type Api = ReturnType<typeof useCommand>;

function fakeAdapter(outcome: CommandOutcome, calls: Array<{ id: string; payload?: unknown }> = []) {
  const adapter: CommandAdapter = {
    has: () => true,
    getState: () => ({ supported: true, enabled: true, active: false, value: undefined }),
    observe: () => () => {},
    run: async (id, payload) => {
      calls.push({ id, payload });
      return outcome;
    },
  };
  return { adapter, calls };
}

/** מרכיבה קומפוננטה שמשתמשת ב-composable ומחזירה את ה-API שלה. */
function mount(
  commandId: string,
  options: { adapter?: CommandAdapter | null; reporter?: CommandReporter; blocked?: Ref<boolean> } = {},
): { api: Api; app: App } {
  let api: Api | null = null;

  const Component = defineComponent({
    setup() {
      api = useCommand(commandId);
      return () => h('div');
    },
  });

  const app = createApp(Component);
  if (options.adapter !== undefined) app.provide(COMMAND_ADAPTER, ref(options.adapter));
  if (options.reporter) app.provide(COMMAND_REPORTER, options.reporter);
  if (options.blocked) app.provide(DOCUMENT_INTERACTION_BLOCKED, options.blocked);
  app.mount(document.createElement('div'));

  if (!api) throw new Error('הקומפוננטה לא הורכבה');
  return { api, app };
}

describe('useCommand', () => {
  it('blocks UI dispatch during candidate initialization without replacing the adapter or losing its readout', async () => {
    const blocked = ref(false);
    const { adapter, calls } = fakeAdapter({ ok: true });
    adapter.getState = () => ({ supported: true, enabled: true, active: false, value: 18 });
    const { api, app } = mount('font-size', { adapter, blocked, reporter: () => {} });
    expect(api.state.value.value).toBe(18);
    blocked.value = true;
    expect(api.enabled.value).toBe(false);
    expect(api.state.value.value).toBe(18);
    expect(await api.run(24)).toMatchObject({ ok: false, reason: 'not-ready' });
    expect(calls).toEqual([]);
    blocked.value = false;
    expect(api.enabled.value).toBe(true);
    expect(api.state.value.value).toBe(18);
    await api.run(24);
    expect(calls).toEqual([{ id: 'font-size', payload: 24 }]);
    app.unmount();
  });

  it('כשל פקודה מדווח למי שיודע להציג אותו, עם ההודעה בעברית של האדפטר', async () => {
    const reporter = vi.fn();
    const { adapter } = fakeAdapter({
      ok: false,
      message: 'יש למקם את הסמן במסמך',
      reason: 'selection-required',
    });

    const { api, app } = mount('text-align', { adapter, reporter });
    await api.run({ alignment: 'right' });
    app.unmount();

    expect(reporter).toHaveBeenCalledTimes(1);
    expect(reporter).toHaveBeenCalledWith(
      { ok: false, message: 'יש למקם את הסמן במסמך', reason: 'selection-required' },
      'text-align',
    );
  });

  it('הצלחה מדווחת גם היא — כדי שאפשר יהיה לנקות שגיאה קודמת מהמסך', async () => {
    const reporter = vi.fn();
    const { adapter } = fakeAdapter({ ok: true });

    const { api, app } = mount('bold', { adapter, reporter });
    const outcome = await api.run();
    app.unmount();

    expect(outcome).toEqual({ ok: true });
    expect(reporter).toHaveBeenCalledWith({ ok: true }, 'bold');
  });

  it('בלי מנוע — הכשל מדווח ואינו נעלם', async () => {
    // זה המסלול שהיה השקט מכולם: לחיצה לפני שהמסמך נטען לא עשתה כלום.
    const reporter = vi.fn();

    const { api, app } = mount('bold', { adapter: null, reporter });
    const outcome = await api.run();
    app.unmount();

    expect(outcome).toEqual({ ok: false, message: 'המנוע אינו מוכן', reason: 'not-ready' });
    expect(reporter).toHaveBeenCalledWith(outcome, 'bold');
  });

  it('בלי מדווח מוזרק הכשל מגיע לקונסולה ואינו זורק', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const { adapter } = fakeAdapter({ ok: false, message: 'המסמך פתוח לקריאה בלבד' });

    const { api, app } = mount('bold', { adapter });
    await expect(api.run()).resolves.toEqual({ ok: false, message: 'המסמך פתוח לקריאה בלבד' });
    app.unmount();

    expect(warn).toHaveBeenCalledWith('[otzaria-word] bold: המסמך פתוח לקריאה בלבד');
    warn.mockRestore();
  });

  it('ה-payload מועבר לאדפטר כפי שהוא', async () => {
    const { adapter, calls } = fakeAdapter({ ok: true });

    const { api, app } = mount('font-size', { adapter });
    await api.run(16);
    app.unmount();

    expect(calls).toEqual([{ id: 'font-size', payload: 16 }]);
  });

  it('פקודה שהמנוע אינו מכיר נשארת לא נתמכת ולא מדווחת מצב שקרי', () => {
    const adapter: CommandAdapter = {
      has: () => false,
      getState: () => {
        throw new Error('אין לקרוא מצב של פקודה שאינה מוכרת');
      },
      observe: () => () => {},
      run: async () => ({ ok: true }),
    };

    const { api, app } = mount('no-such-command', { adapter });

    expect(api.state.value.supported).toBe(false);
    expect(api.enabled.value).toBe(false);
    app.unmount();
  });

  it('פירוק הקומפוננטה מבטל את ההאזנה למצב הפקודה', () => {
    const unsubscribe = vi.fn();
    const adapter: CommandAdapter = {
      has: () => true,
      getState: () => ({ supported: true, enabled: true, active: false, value: undefined }),
      observe: () => unsubscribe,
      run: async () => ({ ok: true }),
    };

    const { app } = mount('bold', { adapter });
    expect(unsubscribe).not.toHaveBeenCalled();

    app.unmount();
    expect(unsubscribe).toHaveBeenCalledTimes(1);
  });
});
