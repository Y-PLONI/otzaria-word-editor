<template>
  <!-- רצועת טאבים אופקית מעל פס הכותרת — אחד לכל מסמך פתוח. -->
  <div class="word-doctabs-bar">
    <div
      class="word-doctabs-strip"
      role="tablist"
      aria-orientation="horizontal"
      aria-label="מסמכים פתוחים"
      @keydown="onTabKeydown"
      @wheel="handleWheelScroll"
    >
      <button
        v-for="(tab, index) in tabs"
        :id="docTabId(tab.id)"
        :key="tab.id"
        :ref="(el) => registerTabRef(el, index)"
        type="button"
        role="tab"
        class="word-doctab"
        :class="{ active: tab.id === activeId }"
        :aria-selected="tab.id === activeId ? 'true' : 'false'"
        :tabindex="tab.id === activeId ? 0 : -1"
        @click="$emit('select-tab', tab.id)"
      >
        <span
          v-if="tab.isDirty"
          class="word-doctab-dirty"
          aria-hidden="true"
        >•</span>
        <span class="word-doctab-title">{{ tab.title || 'מסמך ללא שם' }}</span>
        <span
          class="word-doctab-close"
          role="button"
          :aria-label="`סגור את ${tab.title || 'מסמך ללא שם'}`"
          @click.stop="$emit('close-tab', tab.id)"
        >×</span>
      </button>
    </div>

    <button
      type="button"
      class="word-doctabs-new"
      aria-label="פתח מסמך נוסף"
      data-tip-title="פתח מסמך נוסף"
      @click="$emit('new-tab')"
    >
      +
    </button>
  </div>
</template>

<script setup lang="ts">
import type { ComponentPublicInstance } from 'vue';
import { nextTabIndex } from '../ribbon/aria';
import { handleWheelScroll } from '../../composables/wheel-scroll';

export interface DocumentTabItem {
  id: string;
  title: string;
  isDirty: boolean;
}

const props = withDefaults(
  defineProps<{
    tabs?: DocumentTabItem[];
    activeId?: string | null;
  }>(),
  { tabs: () => [], activeId: null },
);

const emit = defineEmits<{
  (e: 'select-tab', id: string): void;
  (e: 'close-tab', id: string): void;
  (e: 'new-tab'): void;
}>();

function docTabId(id: string): string {
  return `word-doctab-${id}`;
}

/** רק הטאב הפעיל נמצא ב-tab order, ולכן החצים צריכים להזיז מיקוד בעצמם. */
const tabButtons: Array<HTMLButtonElement | null> = [];

function registerTabRef(el: Element | ComponentPublicInstance | null, index: number): void {
  tabButtons[index] = el instanceof HTMLButtonElement ? el : null;
}

/** אותה לוגיקת ניווט כמו לשוניות הרצועה — ראו ui/ribbon/aria.ts. */
function onTabKeydown(event: KeyboardEvent): void {
  const current = props.tabs.findIndex((tab) => tab.id === props.activeId);
  const next = nextTabIndex(event.key, current, props.tabs.length, 'rtl');
  if (next === null) return;

  event.preventDefault();
  const target = props.tabs[next];
  if (!target) return;
  emit('select-tab', target.id);
  tabButtons[next]?.focus();
}
</script>

<style scoped>
.word-doctabs-bar {
  display: flex;
  align-items: center;
  gap: 4px;
  flex-shrink: 0;
  height: var(--doctabs-height);
  padding-inline: 8px;
  background: var(--color-surface-container-high);
  border-bottom: 2px solid var(--color-outline-variant);
  user-select: none;
}

.word-doctabs-strip {
  display: flex;
  align-items: stretch;
  gap: 2px;
  min-width: 0;
  height: calc(100% + 2px);
  margin-bottom: -2px;
  overflow-x: auto;
  padding-inline: 10px;
  scrollbar-width: none;
}

.word-doctabs-strip::-webkit-scrollbar {
  display: none;
}

.word-doctab {
  display: flex;
  align-items: center;
  gap: 6px;
  background: transparent;
  border: 2px solid transparent;
  border-radius: var(--radius-sm);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: var(--font-size-tab);
  padding: 2px 10px;
  cursor: pointer;
  white-space: nowrap;
  max-width: 220px;
}

.word-doctab:hover:not(.active) {
  background: var(--word-btn-hover);
}

.word-doctab.active {
  background: var(--word-doctab-active-bg);
  border-color: var(--color-outline-variant);
  border-bottom-color: transparent;
  color: var(--color-on-surface);
  box-shadow: none;
  border-top-left-radius: var(--radius-md);
  border-top-right-radius: var(--radius-md);
  border-bottom-left-radius: 0;
  border-bottom-right-radius: 0;
  margin-top: 2px;
  margin-bottom: -2px;
  position: relative;
  z-index: 6;
}

.word-doctab.active::before,
.word-doctab.active::after {
  content: '';
  position: absolute;
  bottom: 0;
  width: 10px;
  height: 10px;
  background: transparent;
  pointer-events: none;
}

.word-doctab.active::before {
  left: -12px;
  border-bottom-right-radius: 10px;
  border-right: 2px solid var(--color-outline-variant);
  border-bottom: 2px solid var(--color-outline-variant);
  box-shadow: 5px 5px 0 5px var(--word-doctab-active-bg);
}

.word-doctab.active::after {
  right: -12px;
  border-bottom-left-radius: 10px;
  border-bottom: 2px solid var(--color-outline-variant);
  border-left: 2px solid var(--color-outline-variant);
  box-shadow: -5px 5px 0 5px var(--word-doctab-active-bg);
}

.word-doctab-title {
  overflow: hidden;
  text-overflow: ellipsis;
}

.word-doctab-dirty {
  color: var(--color-secondary);
  line-height: 1;
  flex-shrink: 0;
}

.word-doctab-close {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 14px;
  height: 14px;
  border-radius: var(--radius-sm);
  color: var(--color-on-surface-variant);
  flex-shrink: 0;
  line-height: 1;
}

.word-doctab-close:hover {
  background: var(--word-btn-active);
  color: var(--color-on-surface);
}

.word-doctabs-new {
  flex-shrink: 0;
  background: transparent;
  border: 1px solid transparent;
  border-radius: var(--radius-sm);
  color: var(--color-on-surface-variant);
  font-family: var(--font-main);
  font-size: var(--font-size-tab);
  line-height: 1;
  width: 20px;
  height: 20px;
  padding: 0;
  cursor: pointer;
}

.word-doctabs-new:hover {
  background: var(--word-btn-hover);
  color: var(--color-on-surface);
}
</style>
