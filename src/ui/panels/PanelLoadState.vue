<template>
  <div ref="root" class="panel-load-state" :role="failed ? 'alert' : 'status'">
    <span>{{ failed ? 'לא ניתן לפתוח את החלון. אפשר לנסות שוב.' : 'טוען את החלון…' }}</span>
    <button v-if="failed" type="button" @click="$emit('retry')">נסה שוב</button>
    <button type="button" @click="$emit('close')">ביטול</button>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue';

defineProps<{ failed: boolean }>();
defineEmits<{ retry: []; close: [] }>();
const root = ref<HTMLElement | null>(null);
defineExpose({ contains: (node: Node | null) => root.value?.contains(node) ?? false });
</script>

<style scoped>
.panel-load-state {
  position: fixed;
  inset-block-start: 9rem;
  inset-inline-end: 2rem;
  z-index: 2100;
  display: flex;
  align-items: center;
  gap: 0.75rem;
  max-inline-size: calc(100vw - 4rem);
  padding: 0.75rem 1rem;
  background: var(--color-surface);
  color: var(--color-on-surface);
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-sm);
  box-shadow: 0 0.25rem 1rem rgb(0 0 0 / 15%);
}
.panel-load-state button {
  font: inherit;
  color: inherit;
  padding: 0.25rem 0.5rem;
  white-space: nowrap;
}
</style>
