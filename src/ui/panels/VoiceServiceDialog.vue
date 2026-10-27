<template>
  <div
    v-if="open"
    class="modal-backdrop"
    @click.self="$emit('close')"
  >
    <div
      ref="dialogRef"
      class="voice-dialog"
      role="dialog"
      aria-modal="true"
      :aria-labelledby="TEXT_ID"
      tabindex="-1"
      @keydown.esc.stop="$emit('close')"
      @keydown.tab="onTab"
    >
      <p
        :id="TEXT_ID"
        class="voice-text"
      >
        {{ menuString('להקלדה בדיבור צריך את „שירות הדיבור של אוצריא”, תוכנה נפרדת שמתקינים פעם אחת. הזיהוי נעשה במחשב, בלי אינטרנט.') }}
      </p>

      <div class="voice-footer">
        <button
          ref="downloadRef"
          type="button"
          class="voice-btn voice-btn-primary"
          @pointerdown.prevent
          @click="$emit('download')"
        >
          {{ menuString('להורדה') }}
        </button>
        <button
          type="button"
          class="voice-btn"
          @pointerdown.prevent
          @click="$emit('close')"
        >
          {{ menuString('ביטול') }}
        </button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
/**
 * „שירות הדיבור אינו מותקן” — נפתח כשלחיצה על „הקלטה” לא מצאה את השירות
 * (engine/voice-dictation.ts, `onServiceMissing`). שורה אחת ושני כפתורים,
 * באותו מבנה של UnsavedChangesDialog: בלי כותרת ובלי „✕”, מיקוד על הכפתור
 * הממולא, ומלכודת מיקוד.
 */
import { nextTick, ref, watch } from 'vue';
import { menuString } from '../ribbon/i18n';

const props = withDefaults(defineProps<{ open?: boolean }>(), { open: false });

defineEmits<{
  (e: 'download'): void;
  (e: 'close'): void;
}>();

const TEXT_ID = 'voice-service-dialog-text';

const dialogRef = ref<HTMLElement | null>(null);
const downloadRef = ref<HTMLButtonElement | null>(null);

/** לאן המיקוד חוזר בסגירה — בלעדיו הוא נופל ל-`body`. */
let focusOnClose: HTMLElement | null = null;

watch(
  () => props.open,
  (open) => {
    if (open) {
      focusOnClose = document.activeElement instanceof HTMLElement ? document.activeElement : null;
      void nextTick(() => (downloadRef.value ?? dialogRef.value)?.focus());
      return;
    }
    const target = focusOnClose;
    focusOnClose = null;
    if (target && document.contains(target)) target.focus();
  },
  { immediate: true },
);

/** מלכודת המיקוד — זהה ל-UnsavedChangesDialog. */
function onTab(event: KeyboardEvent): void {
  const items = [...(dialogRef.value?.querySelectorAll<HTMLElement>('button') ?? [])];
  if (items.length === 0) return;
  const first = items[0]!;
  const last = items[items.length - 1]!;
  const active = document.activeElement;
  if (event.shiftKey && (active === first || active === dialogRef.value)) {
    event.preventDefault();
    last.focus();
    return;
  }
  if (!event.shiftKey && active === last) {
    event.preventDefault();
    first.focus();
  }
}
</script>

<style scoped>
.modal-backdrop {
  position: fixed;
  inset: 0;
  background: var(--color-scrim);
  z-index: 2000;
  display: flex;
  align-items: center;
  justify-content: center;
}

.voice-dialog {
  background: var(--color-surface);
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-md);
  box-shadow: 0 12px 36px rgba(0, 0, 0, 0.24);
  width: 360px;
  max-width: 90vw;
  font-family: var(--font-main);
  overflow: hidden;
}

.voice-text {
  margin: 0;
  padding: 20px 20px 16px;
  font-size: 14px;
  line-height: 1.5;
  color: var(--color-on-surface);
}

.voice-footer {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 8px;
  padding: 0 20px 20px;
}

.voice-btn {
  padding: 5px 14px;
  font-size: 12px;
  font-family: var(--font-main);
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  cursor: pointer;
  transition: all 0.08s;
}

.voice-btn:hover {
  background: var(--word-btn-hover);
  border-color: var(--word-blue);
}

.voice-btn-primary {
  background: var(--word-blue);
  color: var(--color-on-primary);
  border-color: var(--word-blue);
  font-weight: 600;
}

.voice-btn-primary:hover {
  background: var(--word-blue-dark);
  border-color: var(--word-blue-dark);
}
</style>
