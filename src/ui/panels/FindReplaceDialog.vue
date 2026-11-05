<template>
  <div
    v-if="isOpen"
    class="find-replace-dialog"
    :style="dragStyle"
    role="dialog"
    aria-label="חיפוש והחלפה"
    @keydown.esc.stop="$emit('close')"
  >
    <div
      class="fr-header dialog-drag-handle"
      @pointerdown="startDialogDrag"
    >
      <div
        class="fr-tabs"
        role="tablist"
        aria-label="מצב חיפוש"
      >
        <button
          type="button"
          class="fr-tab"
          role="tab"
          :aria-selected="activeMode === 'find'"
          :class="{ active: activeMode === 'find' }"
          @click="mode = 'find'"
        >
          חפש
        </button>
        <!-- capability gate: `canReplace: false` = המנוע לא ישנה את המסמך, ולכן
             אין לשונית להחלפה. ההסבר מוצג במקומה ולא מסתיר את הכשל. -->
        <button
          v-if="canReplace"
          type="button"
          class="fr-tab"
          role="tab"
          :aria-selected="activeMode === 'replace'"
          :class="{ active: activeMode === 'replace' }"
          @click="mode = 'replace'"
        >
          החלף
        </button>
      </div>
      <button
        type="button"
        class="fr-close-btn"
        data-tip-title="סגור"
        data-tip-shortcut="Esc"
        aria-label="סגור את חיפוש והחלפה"
        @click="$emit('close')"
      >
        ✕
      </button>
    </div>

    <div class="fr-body">
      <!-- שורת חיפוש -->
      <div class="fr-input-row">
        <label
          for="fr-search-input"
          class="fr-label"
        >חפש את:</label>
        <div class="fr-input-wrapper">
          <input
            id="fr-search-input"
            ref="searchInputRef"
            v-model="searchQuery"
            type="text"
            class="fr-input"
            placeholder="הזן מילת חיפוש..."
            aria-label="טקסט לחיפוש"
            @keydown.enter.exact="findNext"
            @keydown.enter.shift="findPrev"
            @input="onSearchInput"
          >
          <span
            v-if="resultText"
            class="fr-counter"
            role="status"
            aria-live="polite"
          >{{ resultText }}</span>
        </div>
      </div>

      <!-- שורת החלפה (אם מצב החלפה פעיל) -->
      <div
        v-if="activeMode === 'replace'"
        class="fr-input-row"
      >
        <label
          for="fr-replace-input"
          class="fr-label"
        >החלף ב:</label>
        <div class="fr-input-wrapper">
          <input
            id="fr-replace-input"
            v-model="replaceQuery"
            type="text"
            class="fr-input"
            placeholder="טקסט חלופי..."
            aria-label="טקסט חלופי"
            :disabled="isReplacing || readOnly"
            @keydown.enter="doReplace"
          >
        </div>
      </div>

      <!-- הסבר ולא הסתרה בשקט: המשתמש שביקש Ctrl+H צריך לדעת למה אין החלפה. -->
      <p
        v-if="!canReplace"
        class="fr-note"
        role="note"
      >
        {{ REPLACE_UNAVAILABLE_TEXT }}
      </p>
      <p v-if="readOnly" class="fr-note" role="note">
        החלפה אינה זמינה בזמן פתיחת מסמך
      </p>
    </div>

    <!-- כפתורי פעולה -->
    <div class="fr-footer">
      <button
        type="button"
        class="fr-btn fr-btn-primary"
        :disabled="!searchQuery"
        @click="findNext"
      >
        מצא הבא
      </button>
      <button
        type="button"
        class="fr-btn"
        :disabled="!searchQuery"
        @click="findPrev"
      >
        מצא קודם
      </button>
      <button
        v-if="activeMode === 'replace'"
        type="button"
        class="fr-btn"
        :disabled="!searchQuery || isReplacing || readOnly"
        @click="doReplace"
      >
        החלף
      </button>
      <button
        v-if="activeMode === 'replace'"
        type="button"
        class="fr-btn"
        :disabled="!searchQuery || isReplacing || readOnly"
        @click="doReplaceAll"
      >
        החלף הכל
      </button>
      <button
        type="button"
        class="fr-btn fr-btn-secondary"
        @click="$emit('close')"
      >
        ביטול
      </button>
    </div>
  </div>
</template>

<script setup lang="ts">
/**
 * הדיאלוג מציג בלבד: הוא אינו קורא למנוע ואינו מחזיק מצב חיפוש. מונה
 * התוצאות, זמינות ההחלפה וההשקטה של חיפוש-בזמן-הקלדה מגיעים מ-engine/search.ts
 * דרך App.vue, כי שם הם נבדקים — ולא בקומפוננטה, שאין לה כאן תשתית בדיקה.
 *
 * מה שהיה כאן קודם: `resultText` שהוא ref מקומי שרק נמחק ואף פעם לא נכתב,
 * כלומר מונה תוצאות שהוא קוד מת, אף שהמנוע מספק `total` ו-`activeIndex`.
 */
import { ref, computed, watch, nextTick } from 'vue';
import { REPLACE_UNAVAILABLE_TEXT } from '../../engine/search';
import { useDialogDrag } from '../../composables/dialog-drag';

/* הדיאלוג נגרר בכותרת שלו — composables/dialog-drag.ts. */
const { dragStyle, startDialogDrag } = useDialogDrag();

const props = withDefaults(
  defineProps<{
    isOpen?: boolean;
    initialMode?: 'find' | 'replace';
    /** „3 מתוך 12” / „אין תוצאות”, מחושב מהמצב של המנוע. */
    resultText?: string;
    /** האם המנוע יכול להחליט כרגע. `false` = אין להציג את פקדי ההחלפה. */
    canReplace?: boolean;
    /** שאילתה שהגיעה מבחוץ — „חפש במסמך” של Tell Me. ריקה = אין. */
    initialQuery?: string;
    /** החלפה שנשלחה למנוע וטרם הסתיימה. */
    isReplacing?: boolean;
    /** פתיחת מסמך משהה כתיבה בלי למחוק את השאילתה או לסגור את הלוח. */
    readOnly?: boolean;
  }>(),
  {
    isOpen: false,
    initialMode: 'find',
    resultText: '',
    canReplace: false,
    isReplacing: false,
    readOnly: false,
    initialQuery: '',
  }
);

const emit = defineEmits<{
  (e: 'close'): void;
  (e: 'find', query: string, direction: 'next' | 'prev'): void;
  /** הקלדה בשדה החיפוש. ההשקטה באדפטר ולא כאן. */
  (e: 'query-change', query: string): void;
  /** טקסט ההחלפה בלבד — `SearchHandle.replace(replacement)` מקבל ארגומנט אחד. */
  (e: 'replace', replacement: string): void;
  (e: 'replace-all', replacement: string): void;
}>();

const mode = ref<'find' | 'replace'>(props.initialMode);
const searchQuery = ref(props.initialQuery);
const replaceQuery = ref('');
const searchInputRef = ref<HTMLInputElement | null>(null);

/**
 * המצב שמוצג בפועל. Ctrl+H מבקש „החלף” גם כשהמנוע אינו מאפשר החלפה, ואז
 * הדיאלוג נשאר במצב חיפוש עם ההסבר — ולא מציג פקדים שיכשלו.
 */
const activeMode = computed<'find' | 'replace'>(() =>
  props.canReplace ? mode.value : 'find'
);

watch(
  () => props.initialMode,
  (newMode) => {
    mode.value = newMode;
  }
);

/**
 * השאילתה מבחוץ נכנסת לשדה בפתיחה, וגם כשהיא משתנה בדיאלוג שכבר פתוח.
 * שאילתה ריקה אינה מוחקת את מה שהמשתמש חיפש קודם: `Ctrl+F` רגיל נפתח על
 * החיפוש האחרון, כמו ב-Word.
 */
watch([() => props.isOpen, () => props.initialQuery], () => {
  if (!props.isOpen || props.initialQuery === '') return;
  searchQuery.value = props.initialQuery;
  emit('query-change', props.initialQuery);
});

watch(
  () => props.isOpen,
  (open) => {
    if (!open) return;
    nextTick(() => {
      searchInputRef.value?.focus();
      searchInputRef.value?.select();
    });
  }
);

function onSearchInput(): void {
  emit('query-change', searchQuery.value);
}

function findNext(): void {
  if (!searchQuery.value) return;
  emit('find', searchQuery.value, 'next');
}

function findPrev(): void {
  if (!searchQuery.value) return;
  emit('find', searchQuery.value, 'prev');
}

function doReplace(): void {
  if (!searchQuery.value || !props.canReplace || props.isReplacing || props.readOnly) return;
  emit('replace', replaceQuery.value);
}

function doReplaceAll(): void {
  if (!searchQuery.value || !props.canReplace || props.isReplacing || props.readOnly) return;
  emit('replace-all', replaceQuery.value);
}
</script>

<style scoped>
.find-replace-dialog {
  position: absolute;
  top: 140px;
  inset-inline-start: 40px;
  z-index: 2000;
  background: var(--color-surface);
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.16);
  width: 380px;
  font-family: var(--font-main);
  user-select: none;
}

.fr-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  border-block-end: 1px solid var(--color-outline-variant);
  padding-inline: 8px;
  background: var(--color-surface-container-high);
  border-radius: var(--radius-sm) var(--radius-sm) 0 0;
}

.fr-tabs {
  display: flex;
  gap: 2px;
}

.fr-tab {
  background: none;
  border: none;
  padding: 6px 12px;
  font-size: 12px;
  color: var(--color-on-surface-variant);
  cursor: pointer;
  border-block-end: 2px solid transparent;
}

.fr-tab.active {
  color: var(--word-blue);
  font-weight: 600;
  border-block-end-color: var(--word-blue);
}

.fr-close-btn {
  background: none;
  border: none;
  font-size: 14px;
  color: var(--color-on-surface-variant);
  cursor: pointer;
  padding: 4px 6px;
  border-radius: var(--radius-xs);
}

.fr-close-btn:hover {
  background: var(--word-btn-hover);
  color: var(--color-error);
}

.fr-body {
  padding: 12px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.fr-input-row {
  display: flex;
  align-items: center;
  gap: 8px;
}

.fr-label {
  width: 65px;
  font-size: 11px;
  color: var(--color-on-surface);
  flex-shrink: 0;
}

.fr-input-wrapper {
  flex: 1;
  position: relative;
  display: flex;
  align-items: center;
}

.fr-input {
  width: 100%;
  padding: 4px 8px;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: 12px;
  outline: none;
}

.fr-input:focus {
  border-color: var(--word-blue);
  box-shadow: 0 0 0 1px var(--word-blue);
}

.fr-note {
  margin: 0;
  font-size: 11px;
  line-height: 1.4;
  color: var(--color-on-surface-variant);
}

.fr-counter {
  position: absolute;
  inset-inline-end: 6px;
  font-size: 10px;
  color: var(--color-on-surface-variant);
  pointer-events: none;
}

.fr-footer {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 6px;
  padding: 8px 12px;
  border-block-start: 1px solid var(--color-outline-variant);
  background: var(--color-surface-container-high);
  border-radius: 0 0 var(--radius-sm) var(--radius-sm);
  flex-wrap: wrap;
}

.fr-btn {
  padding: 4px 10px;
  font-size: 11px;
  font-family: var(--font-main);
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  cursor: pointer;
  transition: all 0.08s;
}

.fr-btn:hover:not(:disabled) {
  background: var(--word-btn-hover);
  border-color: var(--word-blue);
}

/* הטקסט על הכפתור הממולא הוא `--color-on-primary` ולא לבן קבוע: זה בדיוק
   התפקיד שלו ב-M3 („טקסט/אייקון בתוך אלמנטים בצבע primary”), והוא מגיע
   מהערכה. `#ffffff` שהיה כאן נמדד במצב כהה כלבן על כחול-בהיר — כמעט בלתי
   קריא, כי שם ה-primary עצמו בהיר וה-onPrimary שלו כהה. */
.fr-btn-primary {
  background: var(--word-blue);
  color: var(--color-on-primary);
  border-color: var(--word-blue);
}

.fr-btn-primary:hover:not(:disabled) {
  background: var(--word-blue-dark);
}

.fr-btn:disabled {
  opacity: 0.4;
  cursor: default;
}
</style>
