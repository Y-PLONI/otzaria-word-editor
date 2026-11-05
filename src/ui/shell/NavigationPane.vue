<template>
  <aside
    class="nav-pane"
    aria-label="ניווט"
  >
    <header class="nav-head">
      <span class="nav-title">ניווט</span>
      <button
        type="button"
        class="nav-close"
        aria-label="סגור את חלונית הניווט"
        data-tip-title="סגור"
        @pointerdown.prevent
        @click="$emit('close')"
      >
        ×
      </button>
    </header>

    <ul
      v-if="rows.length > 0"
      ref="treeRef"
      class="nav-tree"
      role="tree"
      aria-label="כותרות"
    >
      <template
        v-for="row in rows"
        :key="row.heading.blockId"
      >
        <li
          v-if="!row.hidden"
          class="nav-item"
          :class="{ 'nav-item--current': row.heading.blockId === currentId }"
          role="treeitem"
          :tabindex="row.heading.blockId === tabStopId ? 0 : -1"
          @focus="focusedId = row.heading.blockId"
          @keydown="onTreeKey($event, row)"
          :aria-level="row.depth + 1"
          :aria-expanded="row.expandable ? !row.collapsed : undefined"
          :aria-current="row.heading.blockId === currentId ? 'location' : undefined"
          :style="{ '--nav-depth': row.depth }"
          :data-tip-title="row.heading.text"
          @pointerdown.prevent
          @click="go(row.heading.blockId)"
        >
          <span
            class="nav-twisty"
            :class="{ 'nav-twisty--open': row.expandable && !row.collapsed }"
            aria-hidden="true"
            @click.stop="row.expandable && toggle(row.heading.blockId)"
          >
            <SvgIcon
              v-if="row.expandable"
              name="chevronLeft"
              :size="12"
            />
          </span>
          <span class="nav-label">{{ row.heading.text }}</span>
        </li>
      </template>
    </ul>

    <p
      v-else
      class="nav-empty"
    >
      {{ emptyText }}
    </p>
    <p v-if="outline?.partial" class="nav-notice" role="status">
      במסמך ארוך מוצג ניווט חלקי מתחילת המסמך. כותרות בהמשך אינן מוצגות, והפרק הנוכחי אינו מסומן שם.
    </p>
  </aside>
</template>

<script setup lang="ts">
/**
 * „חלונית ניווט” — הכותרות של המסמך הפתוח כעץ, ולחיצה על כותרת מעבירה אליה.
 * הקריאה, העץ והקפיצה ב-engine/navigation-pane.ts; כאן רק התזמון והציור.
 *
 * ## מתי נקרא מחדש
 *
 * - **מסמך אחר** (`host`/`generation`) — מיד, והקיפולים מתאפסים: הם נרשמו
 *   על מזהי בלוקים של המסמך הקודם.
 * - **עריכה** (`docRevision`) — אחרי השקטה. קריאת כל הבלוקים בכל תו הייתה
 *   עבודה מיותרת, וקריאות המנוע ממילא אינן חוזרות באמצע הקלדה רציפה.
 * - **תזוזת סמן** (`caretRevision`) — רק הבחירה נקראת, לא המסמך: הפרק
 *   הנוכחי נגזר ממיקום הבלוק במפה שכבר נקראה.
 *
 * הרכיב קיים רק כשהחלונית פתוחה (`v-if` ב-App.vue), ולכן חלונית סגורה אינה
 * קוראת כלום.
 */
import { computed, nextTick, onBeforeUnmount, ref, shallowRef, watch } from 'vue';
import type { SuperDoc } from 'superdoc';
import SvgIcon from '../icons/SvgIcon.vue';
import {
  goToNavHeading,
  invalidateNavOutline,
  navHasChildren,
  navHidden,
  navParents,
  navSectionIndex,
  readCaretBlockId,
  readNavOutline,
  type NavOutline,
} from '../../engine/navigation-pane';

const props = defineProps<{
  host: SuperDoc | null;
  generation: number;
  docRevision: number;
  caretRevision: number;
}>();

const emit = defineEmits<{
  (e: 'close'): void;
  /** הסמן הועבר לכותרת — האב מחזיר את המיקוד למסמך. */
  (e: 'navigated'): void;
}>();

/** השקטה אחרי עריכה — אותו סדר גודל כמו „סימני עיצוב” (300ms). */
const DOC_DEBOUNCE_MS = 400;
/** תזוזת סמן זולה יותר לקרוא, אבל חצים מוחזקים יורים בכל שורה. */
const CARET_DEBOUNCE_MS = 150;

const outline = shallowRef<NavOutline | null>(null);
const loaded = ref(false);
const collapsed = ref<ReadonlySet<string>>(new Set());
const caretOrdinal = ref<number | null>(null);
const treeRef = ref<HTMLElement | null>(null);
const focusedId = ref<string | null>(null);

let docTimer: ReturnType<typeof setTimeout> | undefined;
let caretTimer: ReturnType<typeof setTimeout> | undefined;
let generationToken = 0;
let caretToken = 0;
let disposed = false;

async function readOutline(): Promise<void> {
  const mine = ++generationToken;
  const host = props.host;
  const generation = props.generation;
  const next = await readNavOutline(host, () => !disposed && mine === generationToken && host === props.host && generation === props.generation);
  if (disposed || mine !== generationToken || host !== props.host || generation !== props.generation) return;
  // `null` היא קריאה שנכשלה: משאירים את מה שהיה, ולא „אין כותרות”.
  if (next !== null || !loaded.value) outline.value = next;
  loaded.value = true;
  await readCaret();
}

async function readCaret(): Promise<void> {
  const mine = ++caretToken;
  const host = props.host;
  const generation = props.generation;
  const blockId = await readCaretBlockId(host);
  if (disposed || mine !== caretToken || host !== props.host || generation !== props.generation) return;
  const ordinal = blockId === null ? undefined : outline.value?.ordinals.get(blockId);
  // Unknown blocks include newly inserted text and the unscanned document tail.
  // Do not keep an old section highlighted as though it describes this caret.
  caretOrdinal.value = ordinal ?? null;
}

watch(
  () => [props.host, props.generation] as const,
  (_, previous) => {
    if (previous) invalidateNavOutline(props.host);
    clearTimeout(docTimer);
    clearTimeout(caretTimer);
    caretToken += 1;
    focusedId.value = null;
    outline.value = null;
    loaded.value = false;
    collapsed.value = new Set();
    caretOrdinal.value = null;
    void readOutline();
  },
  { immediate: true },
);

watch(
  () => props.docRevision,
  () => {
    invalidateNavOutline(props.host);
    generationToken += 1;
    caretToken += 1;
    clearTimeout(docTimer);
    docTimer = setTimeout(() => void readOutline(), DOC_DEBOUNCE_MS);
  },
);

watch(
  () => props.caretRevision,
  () => {
    caretToken += 1;
    clearTimeout(caretTimer);
    caretTimer = setTimeout(() => void readCaret(), CARET_DEBOUNCE_MS);
  },
);

onBeforeUnmount(() => {
  disposed = true;
  generationToken += 1;
  caretToken += 1;
  clearTimeout(docTimer);
  clearTimeout(caretTimer);
});

const structure = computed(() => {
  const headings = outline.value?.headings ?? [];
  const parents = navParents(headings);
  const hasChildren = navHasChildren(parents);

  const depths: number[] = [];
  return headings.map((heading, index) => {
    const parent = parents[index];
    depths.push(parent >= 0 ? depths[parent] + 1 : 0);
    return {
      heading,
      depth: depths[index],
      parent,
      expandable: hasChildren[index],
    };
  });
});

// Tree structure changes only after an outline read, not on every caret move.
const rows = computed(() => {
  const items = structure.value;
  const hidden = navHidden(items.map((row) => row.parent), collapsed.value, items.map((row) => row.heading.blockId));
  return items.map((row, index) => ({ ...row, hidden: hidden[index], collapsed: collapsed.value.has(row.heading.blockId) }));
});
const visibleRows = computed(() => rows.value.filter((row) => !row.hidden));
const tabStopId = computed(() => visibleRows.value.find((row) => row.heading.blockId === focusedId.value)?.heading.blockId ?? visibleRows.value[0]?.heading.blockId);
const currentId = computed(() => {
  const headings = outline.value?.headings ?? [];
  const index = caretOrdinal.value === null ? -1 : navSectionIndex(headings, caretOrdinal.value);
  return headings[index]?.blockId;
});
type TreeRow = (typeof rows.value)[number];

async function focusRow(row: TreeRow | undefined): Promise<void> {
  if (!row) return;
  focusedId.value = row.heading.blockId;
  await nextTick();
  treeRef.value?.querySelector<HTMLElement>('[tabindex="0"]')?.focus();
}

function onTreeKey(event: KeyboardEvent, row: TreeRow): void {
  if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return;
  const visible = visibleRows.value;
  const index = visible.indexOf(row);
  switch (event.key) {
    case 'ArrowDown': void focusRow(visible[Math.min(index + 1, visible.length - 1)]); break;
    case 'ArrowUp': void focusRow(visible[Math.max(index - 1, 0)]); break;
    case 'Home': void focusRow(visible[0]); break;
    case 'End': void focusRow(visible[visible.length - 1]); break;
    // The tree is RTL: left opens a branch and right returns to its parent.
    case 'ArrowLeft':
      if (row.expandable && row.collapsed) toggle(row.heading.blockId);
      else if (row.expandable) void focusRow(visible[index + 1]);
      break;
    case 'ArrowRight':
      if (row.expandable && !row.collapsed) toggle(row.heading.blockId);
      else if (row.parent >= 0) void focusRow(rows.value[row.parent]);
      break;
    case 'Enter': case ' ': void go(row.heading.blockId); break;
    default: return;
  }
  event.preventDefault();
  event.stopPropagation();
}

const emptyText = computed(() => {
  if (!props.host) return 'אין מסמך פתוח';
  if (!loaded.value) return 'קורא את הכותרות…';
  if (outline.value === null) return 'לא ניתן לקרוא את הכותרות';
  if (outline.value.partial) return 'לא נמצאו כותרות בחלק המסמך שנקרא.';
  return 'אין כותרות במסמך. החילו סגנון „כותרת 1”, „כותרת 2” וכו׳ כדי שיופיעו כאן.';
});

function toggle(blockId: string): void {
  const next = new Set(collapsed.value);
  if (!next.delete(blockId)) next.add(blockId);
  collapsed.value = next;
}

async function go(blockId: string): Promise<void> {
  const ordinal = outline.value?.ordinals.get(blockId);
  if (ordinal !== undefined) caretOrdinal.value = ordinal;
  const host = props.host;
  const generation = props.generation;
  if (await goToNavHeading(host, blockId) && !disposed && host === props.host && generation === props.generation) emit('navigated');
}
</script>

<style scoped>
.nav-pane {
  display: flex;
  flex-direction: column;
  flex: 0 0 auto;
  width: 240px;
  min-height: 0;
  background: var(--color-surface-container-high);
  border-inline-end: 1px solid var(--color-outline-variant);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: 13px;
}

.nav-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex: 0 0 auto;
  padding: 6px 10px 6px 6px;
  border-block-end: 1px solid var(--color-outline-variant);
}

.nav-title {
  font-weight: 600;
}

.nav-close {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 24px;
  height: 24px;
  padding: 0;
  background: none;
  border: none;
  border-radius: var(--radius-sm);
  color: var(--color-on-surface-variant);
  font-size: 16px;
  line-height: 1;
  cursor: pointer;
}

.nav-close:hover {
  background: var(--word-btn-hover);
}

/* פס הגלילה מימין, כמו ב-`.src-tree` (OpenSourcesNav.vue). */
.nav-tree {
  list-style: none;
  margin: 0;
  padding: 4px;
  flex: 1 1 auto;
  min-height: 0;
  overflow-y: auto;
  direction: ltr;
}

/* הזחה לפי עומק, כמו ב-`.src-item`: הצמתים שטוחים ב-DOM. */
.nav-item {
  direction: rtl;
  display: flex;
  align-items: center;
  gap: 4px;
  min-height: 26px;
  padding-inline: calc(2px + var(--nav-depth, 0) * 14px) 8px;
  border-radius: var(--radius-sm);
  cursor: pointer;
  user-select: none;
  transition: background-color 0.08s ease;
}

.nav-item:focus-visible {
  outline: 2px solid var(--word-blue);
  outline-offset: -2px;
}

.nav-item:hover {
  background: var(--word-btn-hover);
}

/* הפרק שהסמן בו: אותה שפה כמו הפריט הנבחר ב-`.src-item--selected`. */
.nav-item--current {
  background: var(--color-primary-subtle);
  box-shadow: inset -3px 0 0 var(--word-blue);
  font-weight: 600;
}

.nav-item--current:hover {
  background: var(--color-primary-selected-hover);
}

.nav-twisty {
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  justify-content: center;
  width: 16px;
  height: 16px;
  color: var(--color-on-surface-variant);
  transition: transform 0.12s ease;
}

/* ‹ פונה שמאלה ב-RTL (מקופל); פתוח הוא ‹ מסובב כלפי מטה. */
.nav-twisty--open {
  transform: rotate(-90deg);
}

.nav-label {
  flex: 1 1 auto;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.nav-notice {
  margin: 0;
  padding: 8px 12px;
  border-block-start: 1px solid var(--color-outline-variant);
  line-height: 1.5;
}

.nav-empty {
  margin: 0;
  padding: 12px;
  color: var(--color-on-surface-variant);
  line-height: 1.5;
}

@media (prefers-reduced-motion: reduce) {
  .nav-item,
  .nav-twisty {
    transition: none;
  }
}
</style>
