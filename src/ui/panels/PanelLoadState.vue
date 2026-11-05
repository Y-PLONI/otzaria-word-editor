<template>
  <div class="panel-load-backdrop" @click.self="$emit('close')">
    <div
      ref="root"
      class="panel-load-state"
      role="dialog"
      aria-modal="true"
      aria-label="פתיחת חלון"
      tabindex="-1"
      @keydown.esc.stop="$emit('close')"
      @keydown.tab="onTab"
    >
      <span :role="failed ? 'alert' : 'status'">{{ failed ? 'לא ניתן לפתוח את החלון. אפשר לנסות שוב.' : 'טוען את החלון…' }}</span>
      <button v-if="failed" ref="retry" type="button" @click="$emit('retry')">נסה שוב</button>
      <button ref="cancel" type="button" @click="$emit('close')">ביטול</button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { onMounted, ref, watch } from 'vue';

const props = defineProps<{ failed: boolean }>();
defineEmits<{ retry: []; close: [] }>();
const root = ref<HTMLElement | null>(null);
const retry = ref<HTMLButtonElement | null>(null);
const cancel = ref<HTMLButtonElement | null>(null);

onMounted(() => { cancel.value?.focus(); });
watch(() => props.failed, (failed) => {
  // Retry disappears while loading. Move focus before Vue removes that button;
  // the wrapper restores the outside target only when closing or handing off.
  if (!failed && root.value?.contains(document.activeElement)) cancel.value?.focus();
});

function onTab(event: KeyboardEvent): void {
  const first = retry.value ?? cancel.value;
  const last = cancel.value;
  if (!first || !last) return;
  const active = document.activeElement;
  if (event.shiftKey && (active === first || active === root.value)) {
    event.preventDefault();
    last.focus();
  } else if (!event.shiftKey && active === last) {
    event.preventDefault();
    first.focus();
  }
}

defineExpose({ contains: (node: Node | null) => root.value?.contains(node) ?? false });
</script>

<style scoped>
.panel-load-backdrop {
  position: fixed;
  inset: 0;
  z-index: 3000;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--color-scrim);
}
.panel-load-state {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 0.75rem;
  max-inline-size: calc(100vw - 4rem);
  padding: 0.75rem 1rem;
  background: var(--color-surface);
  color: var(--color-on-surface);
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-sm);
  box-shadow: var(--ribbon-shadow-bottom);
}
.panel-load-state button {
  font: inherit;
  color: inherit;
  padding: 0.25rem 0.5rem;
  white-space: nowrap;
}
</style>
