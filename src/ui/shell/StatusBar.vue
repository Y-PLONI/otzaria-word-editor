<template>
  <footer
    id="status"
    class="status word-statusbar"
    role="status"
    :class="{ 'status--error': isError }"
  >
    <!-- צד ימין (RTL): נתוני מסמך. כל פריט מוצג רק אם נמדד — ראו shell-format. -->
    <div class="statusbar-start">
      <div
        v-if="pageText"
        class="status-item"
        data-tip-title="עמודי המסמך"
      >
        <span>{{ pageText }}</span>
      </div>
      <div
        v-if="pageText && wordText"
        class="status-divider"
      />
      <div
        v-if="wordText"
        class="status-item"
        data-tip-title="מספר מילים במסמך"
      >
        <span>{{ wordText }}</span>
      </div>
      <div
        v-if="(pageText || wordText) && statusText"
        class="status-divider"
      />
      <!-- הטולטיפ נושא את ההודעה המלאה, מפני שבחלון צר היא נחתכת בשלוש נקודות
           (ראו `.status-message`) וזו הדרך היחידה לקרוא את סופה. הוא מוצג גם
           כשההודעה שלמה: „חתוכה” ו„שלמה” משתנות עם רוחב החלון ועם אורך ההודעה,
           ותכונה שמופיעה ונעלמת לפי מדידה הייתה עולה ב-ResizeObserver על פס
           שמתעדכן בכל הקלדה. קוראי מסך אינם תלויים בה ממילא: ל-`footer` יש
           `role="status"`, והוא מכריז את הטקסט המלא בלי קשר לחיתוך. -->
      <div
        v-if="statusText"
        class="status-item status-message"
        :class="{ error: isError }"
        :data-tip-title="statusText"
      >
        <span>{{ statusText }}</span>
      </div>

      <!-- מחוון הטעינה. מוצג רק בזמן שמסמך נפתח — ראו sessions/document-load.ts,
           שם גם ההסבר למה „דלג” נעלם ברגע שהפתיחה הצליחה. -->
      <div
        v-if="load.active"
        class="status-divider"
      />
      <div
        v-if="load.active"
        class="status-load"
      >
        <span class="status-load__text">
          <span class="status-load__name">{{ load.name }}</span>
          <span
            v-if="load.stage"
            class="status-load__stage"
          >— {{ load.stage }}</span>
        </span>
        <div
          class="status-load__bar"
          :class="{ 'status-load__bar--full': load.percent >= 100 }"
          role="progressbar"
          aria-label="התקדמות פתיחת המסמך"
          aria-valuemin="0"
          aria-valuemax="100"
          :aria-valuenow="load.percent"
        >
          <div
            class="status-load__fill"
            :style="{ inlineSize: `${load.percent}%` }"
          />
        </div>
        <button
          v-if="load.cancellable"
          type="button"
          class="status-load__skip"
          data-tip-title="דלג"
          data-tip-desc="הפסקת הפתיחה. המסמך שפתוח עכשיו יישאר פתוח"
          @pointerdown.prevent
          @click="$emit('skip-load')"
        >
          דלג
        </button>
      </div>
    </div>

    <!-- צד שמאל (RTL): מצב תצוגה ובקרת זום -->
    <div class="statusbar-end">
      <div class="view-mode-buttons">
        <button
          type="button"
          class="sb-icon-btn"
          :class="{ active: isFocusMode }"
          :aria-pressed="isFocusMode"
          data-tip-title="מצב מיקוד"
          aria-label="מצב מיקוד"
          @pointerdown.prevent
          @click="$emit('toggle-focus')"
        >
          <SvgIcon
            name="focusMode"
            :size="13"
          />
        </button>
      </div>

      <div class="status-divider" />

      <!-- בקרת זום. הגבולות מגיעים מהמנוע ואינם מקודדים כאן. -->
      <div class="zoom-controls">
        <button
          type="button"
          class="zoom-step-btn"
          data-tip-title="הקטן תצוגה"
          aria-label="הקטן תצוגה"
          :disabled="zoomLevel <= zoomMin"
          @pointerdown.prevent
          @click="stepZoom(-ZOOM_STEP)"
        >
          -
        </button>
        <input
          type="range"
          :min="zoomMin"
          :max="zoomMax"
          :step="ZOOM_STEP"
          :value="zoomLevel"
          class="zoom-slider"
          data-tip-title="שינוי גודל תצוגה"
          aria-label="גודל תצוגה באחוזים"
          @input="onZoomSliderChange"
        >
        <button
          type="button"
          class="zoom-step-btn"
          data-tip-title="הגדל תצוגה"
          aria-label="הגדל תצוגה"
          :disabled="zoomLevel >= zoomMax"
          @pointerdown.prevent
          @click="stepZoom(ZOOM_STEP)"
        >
          +
        </button>
        <button
          type="button"
          class="zoom-pct-btn"
          data-tip-title="אפס ל-100%"
          aria-label="אפס ל-100%"
          @pointerdown.prevent
          @click="resetZoom"
        >
          {{ zoomLevel }}%
        </button>
      </div>
    </div>
  </footer>
</template>

<script setup lang="ts">
import { computed } from 'vue';
import SvgIcon from '../icons/SvgIcon.vue';
import { pageLabel, wordCountLabel } from '../../composables/shell-format';
import { clampZoom, FALLBACK_ZOOM } from '../../engine/zoom';
import { idleLoadSnapshot, type LoadSnapshot } from '../../sessions/document-load';

/** צעד הזום של לחצני ± ושל הסרגל, באחוזים. */
const ZOOM_STEP = 10;

const props = withDefaults(
  defineProps<{
    /** עמוד הסמן, מבוסס 1. `null` = אין מקור אמין, ואז מוצג מספר העמודים בלבד. */
    currentPage?: number | null;
    /** מספר עמודי הפריסה. `null` = העימוד טרם דיווח. */
    totalPages?: number | null;
    /** מספר המילים. `null` = טרם נמדד; 0 הוא מסמך ריק, וזה לא אותו דבר. */
    wordCount?: number | null;
    statusText?: string;
    isError?: boolean;
    isFocusMode?: boolean;
    zoomLevel?: number;
    zoomMin?: number;
    zoomMax?: number;
    /**
     * מצב פתיחת המסמך. מגיע כאובייקט אחד ולא כחמישה props, מפני שהוא נמדד
     * כיחידה אחת: אחוז בלי `active` הוא פס שנשאר על המסך אחרי שהפתיחה
     * נגמרה, ו-`cancellable` בלי האחוז הוא כפתור „דלג” בלי מה לדלג עליו.
     */
    load?: LoadSnapshot;
  }>(),
  {
    currentPage: null,
    totalPages: null,
    wordCount: null,
    statusText: '',
    isError: false,
    isFocusMode: false,
    zoomLevel: FALLBACK_ZOOM.value,
    zoomMin: FALLBACK_ZOOM.min,
    zoomMax: FALLBACK_ZOOM.max,
    load: () => idleLoadSnapshot(),
  }
);

const emit = defineEmits<{
  (e: 'update:zoomLevel', zoom: number): void;
  (e: 'toggle-focus'): void;
  (e: 'skip-load'): void;
}>();

const pageText = computed(() => pageLabel(props.currentPage, props.totalPages));
const wordText = computed(() => wordCountLabel(props.wordCount));

function emitZoom(value: number): void {
  emit('update:zoomLevel', clampZoom(value, props.zoomMin, props.zoomMax));
}

function onZoomSliderChange(event: Event): void {
  emitZoom(Number.parseInt((event.target as HTMLInputElement).value, 10));
}

function stepZoom(delta: number): void {
  emitZoom(props.zoomLevel + delta);
}

function resetZoom(): void {
  emitZoom(FALLBACK_ZOOM.value);
}
</script>

<style scoped>
.word-statusbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  height: var(--statusbar-height);
  padding-inline: 10px;
  /* שכבת הבהרה עדינה על משטח הנושא; צבע הבד אינו צובע פקדי ממשק. */
  background: var(--statusbar-bg);
  border-block-start: 1px solid var(--color-outline-variant);
  color: var(--color-on-surface-variant);
  font-size: 11px;
  user-select: none;
  flex-shrink: 0;
  box-shadow: var(--ribbon-shadow-top), var(--ribbon-shadow-bottom);
}

.statusbar-start,
.statusbar-end {
  display: flex;
  align-items: center;
  gap: 4px;
}

/* בלעדיו שם המסמך הנטען דוחף את בקרת הזום מהשורה במקום להיחתך בשלוש נקודות:
   פריט flex אינו מצטמצם מתחת לתוכן שלו בלי זה. */
.statusbar-start {
  min-width: 0;
  /* גם אחרי שההודעה התכווצה עד לריפוד שלה נשאר רוחב מזערי — הריפוד של שני
     פריטי המדידה, שני הקווים והנקודות. נמדד: בחלון של 340 פיקסלים הוא עוד
     חורג ב-6 מעל בקרת הזום. חיתוך הוא צורת הכישלון הנכונה כאן: פס מצב שאיבד
     את סוף הטקסט עדיין קריא, פס שנצבע על גבי הזום אינו. */
  overflow: hidden;
}

/* בקרת הזום ומצב המיקוד אינם נסוגים. `min-width: 0` על `.statusbar-start`
   מצמצם את **התיבה** שלו, אבל טקסט ב-`nowrap` בתוכו גולש ממנה החוצה וממשיך
   להיצבע — ומה שהוא נצבע מעליו הוא הפס הזה. נמדד: בחלון של 560 פיקסלים
   ההודעה „סמנו במסמך את הטקסט לחיפוש…” כיסתה 28 פיקסלים מסרגל הזום, ובחלון
   של 420 — ‏168, כלומר את הסרגל ואת „100%” יחד. */
.statusbar-end {
  flex-shrink: 0;
}

.status-item {
  padding: 1px 6px;
  border-radius: var(--radius-xs);
  cursor: default;
  white-space: nowrap;
}

/* מספר העמודים ומספר המילים קצרים וברוחב כמעט קבוע, והקווים המפרידים הם
   פיקסל. מה שמתקצר הוא ההודעה בלבד — היא היחידה שאורכה אינו חסום. */
.status-item:not(.status-message),
.status-divider {
  flex-shrink: 0;
}

.status-item:hover {
  background: var(--word-btn-hover);
}

.status-divider {
  width: 1px;
  height: 12px;
  background: var(--color-outline-variant);
}

/* ההודעה נחתכת בשלוש נקודות ולא נדחפת מהשורה. `min-width: 0` הוא מה שמאפשר
   את זה בכלל: ברירת המחדל `auto` מקבעת את הרוחב המזערי על רוחב התוכן, ופריט
   שאינו יכול לרדת מתחתיו אינו מצטמצם — הוא גולש. הטקסט המלא בטולטיפ. */
.status-message {
  color: var(--color-on-surface);
  font-weight: 500;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
}

.status-message.error {
  color: var(--color-error);
}

/* -------------------------------------------------------------------- */
/* מחוון הטעינה                                                          */
/* -------------------------------------------------------------------- */

.status-load {
  display: flex;
  align-items: center;
  gap: 6px;
  padding-inline: 6px;
  min-width: 0;
}

.status-load__text {
  display: flex;
  align-items: center;
  gap: 4px;
  min-width: 0;
  color: var(--color-on-surface);
}

/* שם המסמך הוא מה שמתקצר בשלוש נקודות, והשלב הוא מה שנשאר שלם: „ספר בראשית
   מבואר.docx — קורא…” אומר פחות מ„ספר בראש… — קורא את הקובץ”. שם המסמך נמצא
   כאן כי פס הכותרת מציג בזמן פתיחה עדיין את המסמך הקודם. */
.status-load__name {
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  max-inline-size: 22ch;
}

.status-load__stage {
  white-space: nowrap;
  flex-shrink: 0;
  color: var(--color-on-surface-variant);
}

.status-load__bar {
  position: relative;
  inline-size: 96px;
  block-size: 4px;
  flex-shrink: 0;
  border-radius: var(--radius-pill);
  background: var(--color-outline-variant);
  overflow: hidden;
}

.status-load__fill {
  position: absolute;
  inset-block: 0;
  inset-inline-start: 0;
  border-radius: inherit;
  background: var(--word-blue);
  transition: inline-size 180ms ease-out;
}

/* נצנוץ שרץ על הפס בלי קשר להתקדמות — אותה הכרעה כמו במסך הטעינה
   שב-index.html: השלב שבו המנוע בונה את המסמך משאיר את הפס כמעט במקום,
   ובלי תנועה עצמאית המשתמש קורא את זה כקריסה ולא כאיטיות. */
.status-load__bar::after {
  content: '';
  position: absolute;
  inset-block: 0;
  inline-size: 38%;
  border-radius: inherit;
  background: linear-gradient(to left, transparent, var(--word-blue), transparent);
  /* שקיפות ב-opacity ולא ב-color-mix: ל-WebView2 שאוצריא מריצה אין תמיכה
     מובטחת ב-color-mix, וכלל שאינו נתמך היה משאיר נצנוץ שקוף לגמרי. */
  opacity: 0.45;
  animation: status-load-sweep 1.4s ease-in-out infinite;
}

/* הפתיחה נגמרה: הפס מלא ועומד. נצנוץ על פס מלא נראה כמו עוד המתנה. */
.status-load__bar--full::after {
  animation: none;
  opacity: 0;
}

@keyframes status-load-sweep {
  from {
    transform: translateX(120%);
  }
  to {
    transform: translateX(-320%);
  }
}

.status-load__skip {
  background: none;
  border: 1px solid var(--color-outline-variant);
  padding: 0 6px;
  border-radius: var(--radius-xs);
  font: inherit;
  color: var(--color-on-surface-variant);
  cursor: pointer;
  flex-shrink: 0;
}

.status-load__skip:hover {
  background: var(--word-btn-hover);
  border-color: var(--word-btn-active-border);
  color: var(--word-blue);
}

@media (prefers-reduced-motion: reduce) {
  .status-load__fill {
    transition: none;
  }

  .status-load__bar::after {
    animation: none;
  }
}

.view-mode-buttons {
  display: flex;
  align-items: center;
  gap: 2px;
}

.sb-icon-btn {
  background: none;
  border: none;
  padding: 2px 4px;
  border-radius: var(--radius-xs);
  color: var(--color-on-surface-variant);
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
}

.sb-icon-btn:hover {
  background: var(--word-btn-hover);
  color: var(--color-on-surface);
}

.sb-icon-btn.active {
  background: var(--word-btn-active);
  color: var(--word-blue);
}

.zoom-controls {
  display: flex;
  align-items: center;
  gap: 4px;
}

.zoom-step-btn {
  background: none;
  border: none;
  padding: 0 4px;
  font-size: 13px;
  font-weight: 600;
  color: var(--color-on-surface-variant);
  cursor: pointer;
}

.zoom-step-btn:hover:not(:disabled) {
  color: var(--color-on-surface);
}

.zoom-step-btn:disabled {
  opacity: 0.35;
  cursor: default;
}

.zoom-slider {
  -webkit-appearance: none;
  width: 72px;
  height: 3px;
  background: var(--color-outline);
  border-radius: 2px;
  outline: none;
  cursor: pointer;
}

.zoom-slider::-webkit-slider-thumb {
  -webkit-appearance: none;
  width: 10px;
  height: 10px;
  border-radius: 50%;
  background: var(--word-blue);
  cursor: pointer;
}

.zoom-pct-btn {
  background: none;
  border: none;
  padding: 1px 4px;
  border-radius: var(--radius-xs);
  font-size: 11px;
  color: var(--color-on-surface);
  cursor: pointer;
  min-width: 36px;
  text-align: center;
}

.zoom-pct-btn:hover {
  background: var(--word-btn-hover);
}
</style>
