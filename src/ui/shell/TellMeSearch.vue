<template>
  <div
    ref="rootRef"
    class="tell-me-container"
    :class="{ 'is-open': isOpen, 'is-focused': isFocused }"
  >
    <!-- שורת החיפוש בפס הכותרת -->
    <div
      class="tell-me-box search-box"
      @click="focusInput"
    >
      <SvgIcon
        name="search"
        :size="14"
        class="tell-me-search-icon"
      />
      <input
        ref="inputRef"
        v-model="query"
        type="text"
        class="tell-me-input"
        role="combobox"
        aria-autocomplete="list"
        :aria-expanded="isOpen"
        aria-controls="tell-me-listbox"
        :aria-activedescendant="activeItemId"
        aria-label="חיפוש אפשרויות, פקודות ועזרה"
        :placeholder="`חפש אפשרויות, פקודות ועזרה (${TELL_ME_LABEL})`"
        @focus="onFocus"
        @blur="onBlur"
        @input="onInput"
        @keydown="onKeydown"
      >
      <button
        v-if="query"
        type="button"
        class="tell-me-clear-btn"
        aria-label="נקה חיפוש"
        tabindex="-1"
        @click.stop="clearQuery"
      >
        ✕
      </button>
      <span
        v-else
        class="tell-me-shortcut-hint"
      >{{ TELL_ME_LABEL }}</span>
    </div>

    <!-- תפריט תוצאות נפתח (Dropdown) -->
    <div
      v-show="isOpen"
      id="tell-me-listbox"
      ref="dropdownRef"
      class="tell-me-dropdown"
      role="listbox"
      aria-label="תוצאות חיפוש פקודות"
      @mousedown.prevent
    >
      <div class="tell-me-section-title">
        {{ query.trim() ? 'פקודות ואפשרויות' : 'פעולות מוצעות' }}
      </div>

      <!-- רשימת הפקודות -->
      <div
        v-if="filteredActions.length > 0"
        class="tell-me-list"
      >
        <div
          v-for="(action, index) in filteredActions"
          :id="`tell-me-item-${action.id}`"
          :key="action.id"
          role="option"
          class="tell-me-item"
          :class="{ active: index === activeIndex }"
          :aria-selected="index === activeIndex"
          @mouseenter="activeIndex = index"
          @click="executeAction(action)"
        >
          <div class="tell-me-item-icon">
            <SvgIcon
              :name="action.icon"
              :size="16"
            />
          </div>
          <div class="tell-me-item-content">
            <div class="tell-me-item-title">
              {{ action.title }}
            </div>
            <div class="tell-me-item-category">
              {{ action.category }}
            </div>
          </div>
          <div
            v-if="action.shortcut"
            class="tell-me-item-shortcut"
          >
            <kbd class="shortcut-pill">{{ action.shortcut }}</kbd>
          </div>
        </div>
      </div>

      <div
        v-else-if="query.trim()"
        class="tell-me-empty"
      >
        לא נמצאו פקודות מתאימות ל-"{{ query }}"
      </div>

      <!-- שורת חיפוש בתוכן המסמך (מופיעה כשיש שאילתה) -->
      <div
        v-if="query.trim()"
        class="tell-me-divider"
      />
      <div
        v-if="query.trim()"
        id="tell-me-item-doc-search"
        role="option"
        class="tell-me-item tell-me-item-find"
        :class="{ active: activeIndex === filteredActions.length }"
        :aria-selected="activeIndex === filteredActions.length"
        @mouseenter="activeIndex = filteredActions.length"
        @click="executeDocumentSearch"
      >
        <div class="tell-me-item-icon">
          <SvgIcon
            name="search"
            :size="16"
          />
        </div>
        <div class="tell-me-item-content">
          <div class="tell-me-item-title">
            חפש במסמך: <strong>"{{ query }}"</strong>
          </div>
          <div class="tell-me-item-category">
            חיפוש והחלפה בתוכן
          </div>
        </div>
        <div class="tell-me-item-shortcut">
          <kbd class="shortcut-pill">{{ FIND_LABEL }}</kbd>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted } from 'vue';
import SvgIcon from '../icons/SvgIcon.vue';
import {
  searchTellMeActions,
  type TellMeAction,
  type TellMeCustomAction,
} from './tell-me-actions';
import type { RibbonTabId } from '../ribbon/tabs';
import { shortcutLabel, type ShellAction } from '../shortcuts/registry';

/*
 * שלוש התוויות שבתבנית באות מהרג'יסטרי ואינן נכתבות כאן.
 *
 * הן היו כתובות ביד — `Alt+Q` פעמיים ו-`Ctrl+F` — והיו נכונות. מה שהיה שבור
 * הוא ההגנה: בדיקת החוזה סורקת `title`/`tooltip`/`aria-label` בלבד, ולכן
 * `placeholder` וצומת טקסט חמקו ממנה. זה בדיוק המצב שהרג'יסטרי נבנה כדי
 * למנוע — תווית שמבטיחה למשתמש צירוף, בלי שום דבר שיצליב אותה מול המאזין.
 * הבדיקה הורחבה יחד עם השינוי הזה.
 */
const TELL_ME_LABEL = shortcutLabel('tell-me');
const FIND_LABEL = shortcutLabel('find');

const emit = defineEmits<{
  (e: 'run-command', id: string, payload?: unknown): void;
  (e: 'run-action', action: ShellAction): void;
  (e: 'custom-action', action: TellMeCustomAction): void;
  (e: 'open-ribbon-tab', tab: RibbonTabId): void;
  (e: 'open-find', initialQuery?: string): void;
}>();

const rootRef = ref<HTMLElement | null>(null);
const inputRef = ref<HTMLInputElement | null>(null);
const dropdownRef = ref<HTMLElement | null>(null);

const query = ref('');
const isOpen = ref(false);
const isFocused = ref(false);
const activeIndex = ref(0);

const filteredActions = computed(() => {
  return searchTellMeActions(query.value);
});

// סך כל הפריטים האינטראקטיביים בתפריט
const totalItemsCount = computed(() => {
  return filteredActions.value.length + (query.value.trim() ? 1 : 0);
});

const activeItemId = computed(() => {
  if (activeIndex.value < filteredActions.value.length) {
    const action = filteredActions.value[activeIndex.value];
    return action ? `tell-me-item-${action.id}` : undefined;
  }
  if (query.value.trim() && activeIndex.value === filteredActions.value.length) {
    return 'tell-me-item-doc-search';
  }
  return undefined;
});

function focusInput(): void {
  inputRef.value?.focus();
}

function onFocus(): void {
  isFocused.value = true;
  isOpen.value = true;
  activeIndex.value = 0;
}

function onBlur(): void {
  isFocused.value = false;
}

function onInput(): void {
  if (query.value) {
    isOpen.value = true;
  }
  activeIndex.value = 0;
}

function clearQuery(): void {
  query.value = '';
  activeIndex.value = 0;
  inputRef.value?.focus();
}

function closeDropdown(): void {
  isOpen.value = false;
  activeIndex.value = 0;
}

function executeAction(action: TellMeAction): void {
  closeDropdown();
  query.value = '';
  inputRef.value?.blur();

  if (action.command) {
    emit('run-command', action.command.id, action.command.payload);
  } else if (action.shellAction) {
    emit('run-action', action.shellAction);
  } else if (action.customAction) {
    emit('custom-action', action.customAction);
  } else if (action.ribbonTab) {
    emit('open-ribbon-tab', action.ribbonTab);
  }
}

function executeDocumentSearch(): void {
  const searchQuery = query.value.trim();
  closeDropdown();
  query.value = '';
  inputRef.value?.blur();
  emit('open-find', searchQuery);
}

function onKeydown(e: KeyboardEvent): void {
  if (!isOpen.value) {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'Enter') {
      isOpen.value = true;
      e.preventDefault();
      return;
    }
  }

  if (e.key === 'ArrowDown') {
    e.preventDefault();
    if (totalItemsCount.value > 0) {
      activeIndex.value = (activeIndex.value + 1) % totalItemsCount.value;
    }
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    if (totalItemsCount.value > 0) {
      activeIndex.value = (activeIndex.value - 1 + totalItemsCount.value) % totalItemsCount.value;
    }
  } else if (e.key === 'Enter') {
    e.preventDefault();
    if (activeIndex.value >= 0 && activeIndex.value < filteredActions.value.length) {
      executeAction(filteredActions.value[activeIndex.value]);
    } else if (query.value.trim() && activeIndex.value === filteredActions.value.length) {
      executeDocumentSearch();
    } else if (filteredActions.value.length > 0) {
      executeAction(filteredActions.value[0]);
    } else if (query.value.trim()) {
      executeDocumentSearch();
    }
  } else if (e.key === 'Escape' || e.key === 'Esc' || e.code === 'Escape') {
    e.preventDefault();
    closeDropdown();
    query.value = '';
    inputRef.value?.blur();
  } else if (e.key === 'Tab') {
    closeDropdown();
  }
}

function onDocumentClick(e: MouseEvent): void {
  if (rootRef.value && !rootRef.value.contains(e.target as Node)) {
    closeDropdown();
  }
}

onMounted(() => {
  document.addEventListener('click', onDocumentClick, true);
});

onUnmounted(() => {
  document.removeEventListener('click', onDocumentClick, true);
});

defineExpose({
  focus: () => {
    inputRef.value?.focus();
    inputRef.value?.select();
    isOpen.value = true;
  },
  clear: () => {
    query.value = '';
    closeDropdown();
  },
});
</script>

<style scoped>
.tell-me-container {
  position: relative;
  width: 100%;
  max-width: 320px;
  display: flex;
  justify-content: center;
}

.tell-me-box {
  display: flex;
  align-items: center;
  gap: 8px;
  background: var(--color-surface);
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-sm);
  padding: 3px 10px;
  width: 100%;
  cursor: text;
  color: var(--color-on-surface-variant);
  font-family: var(--font-main);
  font-size: 12px;
  transition: border-color 0.12s, box-shadow 0.12s;
  box-sizing: border-box;
}

.tell-me-box:hover {
  border-color: var(--word-blue);
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
}

.tell-me-container.is-focused .tell-me-box {
  border-color: var(--word-blue);
  box-shadow: 0 0 0 1px var(--word-blue);
  background: var(--color-surface);
}

.tell-me-search-icon {
  color: var(--color-on-surface-variant);
  flex-shrink: 0;
}

.tell-me-input {
  flex: 1 1 auto;
  border: none;
  background: transparent;
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: 12px;
  outline: none;
  min-width: 0;
  padding: 0;
}

.tell-me-input::placeholder {
  color: var(--color-on-surface-variant);
  opacity: 0.85;
}

.tell-me-shortcut-hint {
  font-size: 11px;
  color: var(--color-on-surface-variant);
  opacity: 0.65;
  flex-shrink: 0;
}

.tell-me-clear-btn {
  background: none;
  border: none;
  color: var(--color-on-surface-variant);
  cursor: pointer;
  font-size: 12px;
  padding: 0 2px;
  display: flex;
  align-items: center;
  justify-content: center;
  opacity: 0.7;
}

.tell-me-clear-btn:hover {
  opacity: 1;
  color: var(--color-error);
}

/* התפריט הצף */
.tell-me-dropdown {
  position: absolute;
  top: calc(100% + 4px);
  right: 50%;
  transform: translateX(50%);
  width: 380px;
  max-width: 90vw;
  max-height: 420px;
  overflow-y: auto;
  background: var(--color-surface);
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-md);
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.15);
  z-index: 1000;
  padding: 6px 0;
  box-sizing: border-box;
}

.tell-me-section-title {
  font-size: 11px;
  font-weight: 600;
  color: var(--color-on-surface-variant);
  padding: 4px 12px 6px;
  border-block-end: 1px solid var(--color-outline-variant);
  letter-spacing: 0.2px;
}

.tell-me-list {
  display: flex;
  flex-direction: column;
}

.tell-me-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 6px 12px;
  cursor: pointer;
  transition: background 0.08s;
  user-select: none;
}

.tell-me-item:hover,
.tell-me-item.active {
  background: var(--word-btn-hover);
}

.tell-me-item-icon {
  display: flex;
  align-items: center;
  justify-content: center;
  color: var(--word-blue);
  flex-shrink: 0;
  width: 20px;
}

.tell-me-item-content {
  flex: 1 1 auto;
  min-width: 0;
  display: flex;
  flex-direction: column;
  gap: 1px;
}

.tell-me-item-title {
  font-size: 12px;
  font-weight: 500;
  color: var(--color-on-surface);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.tell-me-item-category {
  font-size: 10px;
  color: var(--color-on-surface-variant);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.tell-me-item-shortcut {
  flex-shrink: 0;
}

.shortcut-pill {
  font-family: inherit;
  font-size: 10px;
  padding: 2px 6px;
  border-radius: var(--radius-sm);
  background: var(--color-surface-container-high);
  border: 1px solid var(--color-outline-variant);
  color: var(--color-on-surface-variant);
}

.tell-me-divider {
  height: 1px;
  background: var(--color-outline-variant);
  margin: 4px 0;
}

.tell-me-item-find {
  border-radius: 0;
}

.tell-me-empty {
  padding: 12px;
  font-size: 12px;
  color: var(--color-on-surface-variant);
  text-align: center;
}

@media (max-width: 560px) {
  .tell-me-shortcut-hint {
    display: none;
  }

  .tell-me-input::placeholder {
    content: "חפש";
  }
}
</style>
