import { ref, watch, onUnmounted, inject, computed, shallowRef } from 'vue';
import type { CommandAdapter, CommandOutcome } from '../engine/command-adapter';
import type { CommandState } from 'superdoc/ui';
import { COMMAND_ADAPTER, COMMAND_REPORTER, DOCUMENT_INTERACTION_BLOCKED, READOUT_SELECTION, type CommandReporter } from './keys';
import { UNSETTLED_SELECTION, heldCommandState, type ReadoutSelection } from '../engine/readout-hold';

/**
 * ברירת המחדל כשאין מדווח: הכשל לקונסולה ולא לשום מקום. משמש בבדיקות ובכל
 * הרכבה חלקית של קומפוננטה — ולעולם לא במעטפת עצמה, שמזריקה מדווח אמיתי.
 */
const consoleReporter: CommandReporter = (outcome, commandId) => {
  if (!outcome.ok) console.warn(`[otzaria-word] ${commandId}: ${outcome.message}`);
};

const EMPTY_STATE: CommandState = {
  supported: false,
  enabled: false,
  active: false,
  value: undefined,
};

export function useCommand(commandId: string) {
  const adapterRef = inject(COMMAND_ADAPTER, ref(null));
  const blocked = inject(DOCUMENT_INTERACTION_BLOCKED, ref(false));
  const report = inject(COMMAND_REPORTER, consoleReporter);
  /**
   * ברירת המחדל היא „לא התיישב” ולא „התיישב”: קומפוננטה שמורכבת בלי המעטפת
   * (בדיקה, או רצועה שעולה לפני שנפתח מסמך) אינה יודעת דבר על הבחירה, ובמצב
   * כזה עדיף להחזיק את הקריאה האחרונה מאשר לרוקן אותה. ראו readout-hold.ts.
   */
  const selection = inject(READOUT_SELECTION, shallowRef<ReadoutSelection>(UNSETTLED_SELECTION), true);

  const state = ref<CommandState>({ ...EMPTY_STATE });

  /**
   * הערך האחרון שהמנוע **כן** דיווח.
   *
   * למה הוא נדרש, וההנמקה המלאה ב-engine/readout-hold.ts: המנוע מקפל „מעורב”
   * ו„עוד לא נפתר” לאותו `undefined`, ובכל תו שנקלד הקריאות שלו מתאפסות. בלי
   * הזיכרון הזה כפתור „יישור לימין” נכבה ונדלק 34 פעמים ב-40 שניות הקלדה
   * (נמדד — scripts/ribbon-typing-probe.mjs).
   *
   * `shallowRef` ולא `ref`: הערך הוא תשובה של המנוע (מחרוזת, מספר, ולפעמים
   * אובייקט), ואין שום סיבה לעטוף אותו ב-proxy עמוק.
   */
  const held = shallowRef<unknown>(undefined);

  let unsubscribe: (() => void) | null = null;

  function cleanup(): void {
    if (unsubscribe) {
      unsubscribe();
      unsubscribe = null;
    }
  }

  function setup(adapter: CommandAdapter | null): void {
    cleanup();
    // מסמך חדש מתחיל בלי זיכרון: החזקת ערך של המסמך הקודם הייתה מציגה את
    // היישור שלו על המסמך שנפתח אחריו.
    held.value = undefined;
    if (!adapter || !adapter.has(commandId)) {
      state.value = { ...EMPTY_STATE };
      return;
    }
    state.value = adapter.getState(commandId);
    unsubscribe = adapter.observe(commandId, (newState) => {
      state.value = newState;
    });
  }

  watch(
    adapterRef,
    (newAdapter) => {
      setup(newAdapter);
    },
    { immediate: true }
  );

  // הזיכרון נכתב רק מדיווח שיש בו ערך, ולפני החישוב המוצג — כך „מעורב” אמיתי
  // אינו מוחק אותו, והוא נשאר זמין לרגע שבו הבחירה תחזור להיות סמן.
  watch(
    () => state.value.value,
    (value) => {
      if (value !== undefined) held.value = value;
    },
    { immediate: true }
  );

  /** מה שהפקד מציג: הכול טרי מהמנוע, מלבד הערך שעבר דרך ההחזקה. */
  const displayed = computed<CommandState>(() => {
    const current = heldCommandState(state.value, held.value, selection.value);
    return blocked.value ? { ...current, enabled: false, reason: 'not-ready' } : current;
  });

  onUnmounted(() => {
    cleanup();
  });

  /**
   * מריצה את הפקודה **ומדווחת** על התוצאה. הדיווח כאן ולא באתר הקריאה בכוונה:
   * הפקדים ב-Ribbon קוראים `void cmd.run()`, וכל דרך אחרת הייתה מחייבת 38
   * אתרי קריאה לזכור לטפל בכשל — וזה בדיוק מה שלא קרה עד עכשיו.
   */
  async function run(payload?: unknown): Promise<CommandOutcome> {
    const outcome = adapterRef.value && !blocked.value
      ? await adapterRef.value.run(commandId, payload)
      : ({ ok: false, message: 'המנוע אינו מוכן', reason: 'not-ready' } as CommandOutcome);

    report(outcome, commandId);
    return outcome;
  }

  return {
    /**
     * המצב כפי שהפקד מציג אותו — כולל ההחזקה. הצרכנים קוראים ממנו
     * `state.value.value`, ולכן ההחזקה מגיעה גם למי שאינו עובר דרך `value`
     * שלמטה (כפתורי היישור משווים `state.value.value === 'right'`).
     */
    state: displayed,
    /** המצב הגולמי מהמנוע, בלי החזקה. למי שצריך למדוד את המנוע עצמו. */
    engineState: state,
    enabled: computed(() => displayed.value.enabled),
    active: computed(() => state.value.active),
    value: computed(() => displayed.value.value),
    run,
  };
}
