<template>
  <div
    v-show="visible"
    ref="rootRef"
    class="doc-vruler"
    role="group"
    aria-label="סרגל אנכי"
  >
    <div
      v-if="geometry"
      class="doc-vruler__page"
      :style="{ top: `${geometry.pageTopPx}px`, height: `${geometry.pageHeightPx}px` }"
    >
      <!-- אזור הטקסט, בין השוליים העליונים לתחתונים -->
      <div
        class="doc-vruler__text-area"
        :style="{ top: `${geometry.textTopPx}px`, height: `${geometry.textHeightPx}px` }"
      />

      <div
        v-for="(tick, index) in ticks"
        :key="`t${index}`"
        class="doc-vruler__tick"
        :class="`doc-vruler__tick--${tick.kind}`"
        :style="{ top: `${tick.px}px` }"
      >
        <span
          v-if="tick.label"
          class="doc-vruler__number"
        >{{ tick.label }}</span>
      </div>

      <div
        v-for="handle in handles"
        :key="handle.id"
        class="doc-vruler__handle"
        :class="{ 'is-dragging': dragValue?.id === handle.id, 'is-disabled': !canEdit }"
        role="slider"
        aria-orientation="vertical"
        :tabindex="canEdit ? 0 : -1"
        :aria-label="handle.label"
        :aria-disabled="canEdit ? undefined : 'true'"
        :aria-valuemin="0"
        :aria-valuemax="round2(handle.maxValueTwips / unitTwips)"
        :aria-valuenow="round2(handle.valueTwips / unitTwips)"
        :aria-valuetext="measureLabel(handle.valueTwips, unit)"
        :title="`${handle.label}: ${measureLabel(handle.valueTwips, unit)}`"
        :style="{ top: `${handle.px}px` }"
        @pointerdown="onPointerDown(handle, $event)"
        @keydown="onKeyDown(handle, $event)"
      />

      <div
        v-if="draggedHandle"
        class="doc-vruler__readout"
        :style="readoutStyle"
      >
        {{ draggedHandle.label }}: {{ measureLabel(draggedHandle.valueTwips, unit) }}
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
/**
 * הסרגל האנכי — שוליים עליונים ותחתונים, כמו ב-Word.
 *
 * ## מה שונה כאן מהסרגל האופקי
 *
 * שלושה דברים, וכולם מפשטים:
 *
 *   1. **אין כיווניות.** הציר האנכי נמדד מלמעלה למטה בכל שפה, ולכן אין כאן
 *      „צד התחלה”: המרחק נמדד מקצה הדף העליון, תמיד. `pixelOffset` נקרא עם
 *      `'ltr'` קבוע, וזה אינו קיצור דרך אלא הכיוון הנכון.
 *   2. **אין כניסות.** כניסת פסקה היא מידה אופקית; ב-Word גם הסרגל האנכי אינו
 *      מציג אותה. שתי ידיות בלבד — שוליים עליונים ותחתונים.
 *   3. **העמוד שנמדד הוא זה שנראה**, ולא הראשון: המשתמש גלל לעמוד השני, והמידה
 *      שהוא צריך היא של מה שמולו. ההכרעה עצמה ב-`measurePageRect` עם `axis: 'y'`.
 *
 * כל השאר — הצמדה, חסמים, כתיבה בשחרור, מקלדת ו-`role="slider"` — זהה
 * ל-DocumentRuler.vue ומגיע מאותם מודולים, כדי ששני הסרגלים לא יתחילו
 * להתנהג שונה.
 */
import { computed, inject, ref, shallowRef, watch, type Ref } from 'vue';
import type { SuperDoc } from 'superdoc';
import { ACTIVE_SUPERDOC } from '../../engine/document-api';
import { COMMAND_REPORTER, type CommandReporter } from '../../composables/keys';
import { applyPageMargins } from '../../engine/page-setup';
import {
  watchPageRect,
  type PageRect,
  type RulerReading,
  type ViewportSource,
} from '../../engine/page-ruler';
import {
  clampMargin,
  measureLabel,
  pixelOffset,
  rulerTicks,
  snapStepUnits,
  snapTwips,
  twipsFromPixel,
  twipsPerUnit,
  type RulerTick,
  type RulerUnit,
} from '../../engine/ruler-geometry';

const props = withDefaults(
  defineProps<{
    visible?: boolean;
    reading?: RulerReading | null;
    host?: HTMLElement | null;
    viewportSource?: ViewportSource | null;
    unit?: RulerUnit;
    /**
     * גודל התצוגה באחוזים. אינו משמש לחישוב — המלבן נמדד ממילא — אלא כטריגר:
     * `viewport.observe` של המנוע **אינו** מדווח על שינוי זום (נמדד על ה-dist
     * הארוז: זום 70% הקטין את העמוד, והסרגל נשאר במקומו), ואין אירוע DOM שקורה
     * אז. זה הסימן היחיד שיש לנו.
     */
    zoom?: number;
    editable?: boolean;
  }>(),
  {
    visible: false,
    reading: null,
    host: null,
    viewportSource: null,
    unit: 'cm',
    zoom: 100,
    editable: true,
  },
);

const emit = defineEmits<{ (e: 'changed'): void }>();

const fallbackReporter: CommandReporter = (outcome, id) => {
  if (!outcome.ok) console.warn(`[otzaria-word] ${id}: ${outcome.message}`);
};

const superdoc = inject(ACTIVE_SUPERDOC, shallowRef<SuperDoc | null>(null)) as Ref<SuperDoc | null>;
const report = inject(COMMAND_REPORTER, fallbackReporter);

const rootRef = ref<HTMLElement | null>(null);
const pageRect = shallowRef<PageRect | null>(null);
const inFlight = ref(false);

const unit = computed<RulerUnit>(() => props.unit);
const unitTwips = computed(() => twipsPerUnit(unit.value));
const canEdit = computed(() => props.editable && !inFlight.value && props.reading !== null);

/* ------------------------------------------------------------------ */

let watcher: ReturnType<typeof watchPageRect> | null = null;

watch(
  [() => props.host, () => props.visible, rootRef],
  ([host, visible, root]) => {
    watcher?.dispose();
    watcher = null;
    pageRect.value = null;
    if (!visible || !host || !root) return;
    watcher = watchPageRect({
      host,
      reference: root,
      ui: props.viewportSource,
      axis: 'y',
      onChange: (rect) => {
        pageRect.value = rect;
      },
    });
  },
  { immediate: true, flush: 'post' },
);

watch([() => props.reading, () => props.zoom], () => watcher?.measure());

/* ------------------------------------------------------------------ */

interface Geometry {
  pageTopPx: number;
  pageHeightPx: number;
  pageHeightTwips: number;
  topMarginTwips: number;
  bottomMarginTwips: number;
  textTopPx: number;
  textHeightPx: number;
  pxPerUnit: number;
}

const geometry = computed<Geometry | null>(() => {
  const reading = props.reading;
  const rect = pageRect.value;
  if (!reading || !rect || !(rect.heightPx > 0) || !(reading.page.pageHeightTwips > 0)) return null;

  const drag = dragValue.value;
  const pageHeightTwips = reading.page.pageHeightTwips;
  const topMarginTwips = drag?.id === 'margin-top' ? drag.valueTwips : reading.page.topTwips;
  const bottomMarginTwips =
    drag?.id === 'margin-bottom' ? drag.valueTwips : reading.page.bottomTwips;
  const pxPerTwip = rect.heightPx / pageHeightTwips;

  return {
    pageTopPx: rect.topPx,
    pageHeightPx: rect.heightPx,
    pageHeightTwips,
    topMarginTwips,
    bottomMarginTwips,
    textTopPx: topMarginTwips * pxPerTwip,
    textHeightPx: Math.max(0, (pageHeightTwips - topMarginTwips - bottomMarginTwips) * pxPerTwip),
    pxPerUnit: pxPerTwip * twipsPerUnit(unit.value),
  };
});

/** מרחק מקצה הדף העליון → פיקסלים בתוך מלבן העמוד. */
function toPx(twipsFromTop: number): number {
  const geo = geometry.value;
  if (!geo) return 0;
  return pixelOffset(
    twipsFromTop,
    { leftPx: 0, widthPx: geo.pageHeightPx, widthTwips: geo.pageHeightTwips },
    'ltr',
  );
}

function toTwips(pxInPage: number): number {
  const geo = geometry.value;
  if (!geo) return 0;
  return twipsFromPixel(
    pxInPage,
    { leftPx: 0, widthPx: geo.pageHeightPx, widthTwips: geo.pageHeightTwips },
    'ltr',
  );
}

const ticks = computed<Array<RulerTick & { px: number }>>(() => {
  const geo = geometry.value;
  if (!geo) return [];
  return rulerTicks({
    pageWidthTwips: geo.pageHeightTwips,
    textStartTwips: geo.topMarginTwips,
    unit: unit.value,
    pxPerUnit: geo.pxPerUnit,
  }).map((tick) => ({ ...tick, px: toPx(tick.twips) }));
});

/* ------------------------------------------------------------------ */

type HandleId = 'margin-top' | 'margin-bottom';

interface Handle {
  id: HandleId;
  label: string;
  posTwips: number;
  valueTwips: number;
  maxValueTwips: number;
  px: number;
}

const handles = computed<Handle[]>(() => {
  const geo = geometry.value;
  if (!geo) return [];
  return [
    {
      id: 'margin-top',
      label: 'שוליים עליונים',
      posTwips: geo.topMarginTwips,
      valueTwips: geo.topMarginTwips,
      maxValueTwips: geo.pageHeightTwips - geo.bottomMarginTwips,
      px: toPx(geo.topMarginTwips),
    },
    {
      id: 'margin-bottom',
      label: 'שוליים תחתונים',
      posTwips: geo.pageHeightTwips - geo.bottomMarginTwips,
      valueTwips: geo.bottomMarginTwips,
      maxValueTwips: geo.pageHeightTwips - geo.topMarginTwips,
      px: toPx(geo.pageHeightTwips - geo.bottomMarginTwips),
    },
  ];
});

function valueAtPosition(id: HandleId, posTwips: number): number {
  const geo = geometry.value;
  if (!geo) return 0;
  return id === 'margin-top' ? posTwips : geo.pageHeightTwips - posTwips;
}

function clampValue(id: HandleId, valueTwips: number): number {
  const geo = geometry.value;
  if (!geo) return 0;
  return clampMargin(valueTwips, {
    pageWidthTwips: geo.pageHeightTwips,
    otherMarginTwips: id === 'margin-top' ? geo.bottomMarginTwips : geo.topMarginTwips,
  });
}

/* ------------------------------------------------------------------ */

interface DragValue {
  id: HandleId;
  valueTwips: number;
}

const dragValue = shallowRef<DragValue | null>(null);
let dragPointerId: number | null = null;

const draggedHandle = computed(() => {
  const drag = dragValue.value;
  if (!drag) return null;
  return handles.value.find((handle) => handle.id === drag.id) ?? null;
});

/**
 * התווית יוצאת אל מחוץ לעמודה, ולכן היא `fixed` ולא `absolute`: עמודה ברוחב
 * 22px חייבת `overflow: hidden` (אחרת השנתות של עמוד שגלול חציו מציירות מעל
 * הרצועה ושורת המצב), ו-`hidden` חותך גם את התווית. קואורדינטות ה-viewport
 * הן בדיוק מה שיש כאן ממילא.
 */
const readoutStyle = computed(() => {
  const root = rootRef.value;
  const geo = geometry.value;
  const handle = draggedHandle.value;
  if (!root || !geo || !handle) return {};
  const box = root.getBoundingClientRect();
  return {
    position: 'fixed' as const,
    top: `${box.top + geo.pageTopPx + handle.px}px`,
    left: `${box.left + box.width + 4}px`,
  };
});

function pageOriginY(): number {
  const root = rootRef.value;
  const geo = geometry.value;
  if (!root || !geo) return 0;
  return root.getBoundingClientRect().top + geo.pageTopPx;
}

function onPointerDown(handle: Handle, event: PointerEvent): void {
  if (!canEdit.value || event.button !== 0) return;
  const element = event.currentTarget;
  if (!(element instanceof HTMLElement)) return;

  event.preventDefault();
  try {
    element.setPointerCapture?.(event.pointerId);
  } catch {
    /* גרירה עובדת גם בלי תפיסה */
  }
  element.focus();
  dragPointerId = event.pointerId;
  dragValue.value = { id: handle.id, valueTwips: handle.valueTwips };

  const move = (moveEvent: PointerEvent): void => {
    if (moveEvent.pointerId !== dragPointerId) return;
    const drag = dragValue.value;
    if (!drag) return;
    const raw = valueAtPosition(drag.id, toTwips(moveEvent.clientY - pageOriginY()));
    dragValue.value = {
      id: drag.id,
      valueTwips: clampValue(drag.id, snapTwips(raw, unit.value, moveEvent.altKey)),
    };
  };

  const stop = (): void => {
    window.removeEventListener('pointermove', move);
    window.removeEventListener('pointerup', finish);
    window.removeEventListener('pointercancel', cancel);
    dragPointerId = null;
  };

  const finish = (upEvent: PointerEvent): void => {
    if (upEvent.pointerId !== dragPointerId) return;
    const drag = dragValue.value;
    stop();
    dragValue.value = null;
    if (drag) void commit(drag.id, drag.valueTwips);
  };

  const cancel = (): void => {
    stop();
    dragValue.value = null;
  };

  window.addEventListener('pointermove', move);
  window.addEventListener('pointerup', finish);
  window.addEventListener('pointercancel', cancel);
}

/** למעלה ולמטה, בלי היפוך: הציר האנכי אינו מתהפך בין שפות. */
function onKeyDown(handle: Handle, event: KeyboardEvent): void {
  if (!canEdit.value) return;
  const geo = geometry.value;
  if (!geo) return;

  const step = snapStepUnits(unit.value) * unitTwips.value * (event.shiftKey ? 0.2 : 1);
  let posTwips: number | null = null;
  if (event.key === 'ArrowUp') posTwips = handle.posTwips - step;
  else if (event.key === 'ArrowDown') posTwips = handle.posTwips + step;
  else if (event.key === 'Home') posTwips = 0;
  else if (event.key === 'End') posTwips = geo.pageHeightTwips;
  if (posTwips === null) return;

  event.preventDefault();
  void commit(handle.id, clampValue(handle.id, valueAtPosition(handle.id, Math.round(posTwips))));
}

async function commit(id: HandleId, valueTwips: number): Promise<void> {
  const geo = geometry.value;
  const host = superdoc.value;
  if (!geo || !host || inFlight.value) return;

  inFlight.value = true;
  try {
    const outcome = await applyPageMargins(host, {
      ...(id === 'margin-top' ? { topTwips: valueTwips } : { bottomTwips: valueTwips }),
    });
    report(outcome, 'ruler-margins-vertical');
    if (outcome.ok) emit('changed');
  } finally {
    inFlight.value = false;
  }
}

function round2(value: number): number {
  return Math.round(value * 100) / 100;
}
</script>

<style scoped>
/* עמודה צרה לצד אזור המסמך. הרוחב זהה לגובה הסרגל האופקי, כדי ששני הסרגלים
   ייראו כמו זוג ולא כמו שני פקדים שהגיעו ממקומות שונים. */
.doc-vruler {
  position: relative;
  flex-shrink: 0;
  width: 22px;
  overflow: hidden;
  background: var(--color-surface-container-highest);
  border-inline-end: 1px solid var(--color-outline-variant);
  user-select: none;
  touch-action: none;
}

/* ראו את ההסבר על `box-shadow` ב-DocumentRuler.vue: גבול היה מזיז את
   התוכן פיקסל אחד מהעמוד. */
.doc-vruler__page {
  position: absolute;
  inset-inline: 0;
  background: var(--color-surface-container-high);
  box-shadow:
    inset 0 1px var(--color-outline-variant),
    inset 0 -1px var(--color-outline-variant);
}

.doc-vruler__text-area {
  position: absolute;
  inset-inline: 3px;
  background: var(--color-surface);
  border: 1px solid var(--color-outline-variant);
  border-radius: 1px;
}

.doc-vruler__tick {
  position: absolute;
  height: 1px;
  background: var(--color-on-surface-variant);
  opacity: 0.55;
}

/* אותם דירוגים כמו בסרגל האופקי, מסובבים ב-90 מעלות — וכולם נצמדים לקצה
   אחד של העמודה. הצד השני שמור למספרים: ברצועה של 22px, שנתות ומספרים
   שחולקים מקום פירושם שאחד מהם נחתך. */
.doc-vruler__tick--major {
  inset-inline-end: 2px;
  width: 6px;
  opacity: 0.75;
}

.doc-vruler__tick--mid {
  inset-inline-end: 2px;
  width: 4px;
}

.doc-vruler__tick--minor {
  inset-inline-end: 2px;
  width: 2px;
  opacity: 0.4;
}

.doc-vruler__number {
  position: absolute;
  inset-inline-start: 1px;
  top: 50%;
  transform: translateY(-50%);
  font-family: var(--font-main);
  font-size: 9px;
  line-height: 1;
  color: var(--color-on-surface-variant);
  pointer-events: none;
}

.doc-vruler__handle {
  position: absolute;
  inset-inline: 0;
  height: 14px;
  margin-block-start: -7px;
  cursor: row-resize;
}

.doc-vruler__handle.is-disabled {
  cursor: default;
  opacity: 0.5;
}

.doc-vruler__handle:focus-visible {
  outline: 2px solid var(--color-primary);
  outline-offset: 1px;
  border-radius: var(--radius-xs);
}

.doc-vruler__handle::before {
  content: '';
  position: absolute;
  inset-inline: 2px;
  top: 50%;
  height: 3px;
  transform: translateY(-50%);
  border-radius: var(--radius-pill);
  background: var(--color-on-surface-variant);
  opacity: 0.55;
}

.doc-vruler__handle:hover::before,
.doc-vruler__handle.is-dragging::before {
  background: var(--word-blue);
  opacity: 1;
}

/* התווית יוצאת אל תוך אזור המסמך — ברצועה של 22px אין מקום למספר.
   המיקום עצמו מגיע מ-`readoutStyle` (position: fixed); ראו את ההסבר שם. */
.doc-vruler__readout {
  transform: translateY(-50%);
  padding: 0 5px;
  border-radius: var(--radius-xs);
  background: var(--color-primary);
  color: var(--color-on-primary);
  font-family: var(--font-main);
  font-size: 10px;
  line-height: 16px;
  white-space: nowrap;
  pointer-events: none;
  z-index: 2;
}
</style>
