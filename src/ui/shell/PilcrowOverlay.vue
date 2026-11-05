<template>
  <div
    ref="rootRef"
    class="pilcrow-layer"
    aria-hidden="true"
  >
    <span
      v-for="mark in marks"
      :key="mark.nodeId"
      class="pilcrow-layer__mark"
      :class="{ 'pilcrow-layer__mark--rtl': mark.direction === 'rtl' }"
      :style="{
        left: `${mark.anchorXPx}px`,
        top: `${mark.topPx}px`,
        height: `${mark.heightPx}px`,
        lineHeight: `${mark.heightPx}px`,
      }"
      >¶</span
    >
  </div>
</template>

<script setup lang="ts">
/**
 * שכבת „סימני עיצוב" — מציירת ¶ בסוף כל פסקה, כשהפקד „הצג/הסתר סימני
 * עיצוב" דלוק.
 *
 * ## למה השכבה הזאת קיימת בכלל
 *
 * `docs/superdoc-2.10-review.md` ("סימני עיצוב (¶/·/→) — נחקר לעומק") מתעד:
 * הפקודה `formatting-marks` מתהפכת כמו שצריך במנוע, אבל נתיב הציור המובנה
 * שלו מת מבנית ב-2.10.0 (`getDocumentRuntime()` מוקלד `() => null`, ואף
 * קורא ל-`setDocumentRuntime` לא קיים בכל ה-bundle). בדיוק כמו „גבולות
 * עמוד"/„מספרי שורות" — זו לא באג בקוד שלנו, והשכבה הזאת היא ציור **שלנו**,
 * מעל מה שהמנוע כבר מסמן כדלוק, ואינה נוגעת בשום דבר בשכבת הכתיבה/ייצוא.
 *
 * ## למה ¶ בלבד, ולא גם · (רווח) ו-→ (טאב)
 *
 * ראו `docs/superdoc-2.10-review.md`: טאבים אינם קיימים כתו בשום Text node,
 * וסמן רשימה מזריק טקסט שאינו קיים בטקסט הקנוני של `blocks.list()` — שני
 * חסמים שפוסלים מיקום מדויק. ¶ עוקף אותם כי הוא נוגע רק בסוף הפסקה. ההיקף
 * המצומצם הזה (וגם: רק `nodeType === 'paragraph'`, לא כותרות/פריטי-רשימה)
 * מתועד ונבדק — לא ניחוש חלקי. ראו engine/formatting-marks-layer.ts.
 *
 * ## שני מקורות
 *
 *   1. **מה לצייר** — `blocks` (`FormattingMarksBlock[]`, engine/formatting-marks.ts),
 *      נקרא מ-`doc.blocks.list` לפי מזהי סופי הפסקאות שבחלון. `null`/ריק
 *      כשאין Document API — אין מה לצייר.
 *   2. **איפה לצייר** — חלון של סופי פסקאות מזוהים, דרך
 *      `watchFormattingMarksWindow`. עיגוני העמוד נשארים ב-page-ruler.ts,
 *      ועיגוני זהות הפסקה ב-painted-lines.ts. כאן אין selectors של מנוע.
 *
 * ההרכבה לפי מזהה היא engine/formatting-marks-layer.ts. בלי מיפוי תקין
 * נשמר fallback עם שני הרצפים המלאים; אין סינון של צד אחד בלבד.
 *
 * ## `visible` עוצר את המדידה, לא רק את הציור
 *
 * בשונה מ-`PageBorderOverlay`/`LineNumberOverlay` (שתמיד עוקבים אחרי
 * הגיאומטריה, ורק `reading:null` מצייר ריק), כאן `visible:false` **מפרקת**
 * את המעקב אחרי ריצות הטקסט לגמרי: המדידה כאן (TreeWalker+Range על כל צומת,
 * ועוד `getComputedStyle` לכיוון) יקרה יותר מזו של שתי השכבות האחרות, וסימני
 * עיצוב כבויים הם המצב השכיח (ברירת המחדל של הפקד — כבוי). אין טעם למדוד
 * מה שאיש לא רואה.
 *
 * ## עדכון חי
 *
 * הצגה/הסתרה מגיעה מ-`adapter.observe('formatting-marks', …)` ב-App.vue —
 * בשונה מ„גבולות עמוד"/„מספרי שורות", הפקודה הזו **כן** מפעילה עדכון מצב
 * תקין (נמדד: `active` מתהפך כראוי), ולכן אין צורך ברענון מפורש על הפקודה
 * עצמה; App.vue קורא `formattingMarksModel.setEnabled(active)` על כל שינוי.
 * עריכת טקסט/reflow מגיעים דרך `noteDocumentChanged()` (App.vue, `onUpdate`)
 * וגם דרך `watchFormattingMarksWindow` (גלילה, שינוי גודל, `viewport.observe`).
 */
import { computed, onBeforeUnmount, shallowRef, watch } from 'vue';
import {
  type PageRectWatch,
  type ViewportSource,
} from '../../engine/page-ruler';
import { watchFormattingMarksWindow, type FormattingMarksWindow } from '../../engine/formatting-marks-viewport';
import { buildMappedPilcrowMarks, buildPilcrowMarks, type FormattingMarksBlock, type PilcrowMark } from '../../engine/formatting-marks-layer';

const props = withDefaults(
  defineProps<{
    /** ה-host המצויר של המסמך הפתוח, מ-`paintedHost(ui)`. */
    host?: HTMLElement | null;
    /** ה-controller, בשביל `viewport.observe` — אותו prop כמו שתי השכבות האחרות. */
    viewportSource?: ViewportSource | null;
    /** בלוקי המסמך הנוכחיים, או `null` כשאין (מסמך סגור, או עדיין לא נקרא). */
    blocks?: readonly FormattingMarksBlock[] | null;
    /** מצב הפקד „הצג/הסתר סימני עיצוב". `false` = לא מוצג ולא נמדד. */
    visible?: boolean;
  }>(),
  {
    host: null,
    viewportSource: null,
    blocks: null,
    visible: false,
  },
);

const emit = defineEmits<{
  /** null means the complete legacy sequence, never a one-sided window. */
  (e: 'scope-change', nodeIds: readonly string[] | null): void;
}>();

const rootRef = shallowRef<HTMLElement | null>(null);
const painted = shallowRef<FormattingMarksWindow | null>(null);

let watcher: PageRectWatch | null = null;

function stopWatching(): void {
  watcher?.dispose();
  watcher = null;
}

watch(
  [() => props.host, rootRef, () => props.visible, () => props.viewportSource],
  ([host, root, visible]) => {
    stopWatching();
    painted.value = null;
    if (!host || !root || !visible) { emit('scope-change', []); return; }
    watcher = watchFormattingMarksWindow({
      host,
      reference: root,
      ui: props.viewportSource,
      onChange: (next) => {
        painted.value = next;
        emit('scope-change', next.mode === 'mapped' ? [...new Set(next.endings.map((end) => end.nodeId))] : null);
      },
    });
  },
  { immediate: true, flush: 'post' },
);

// בלוקים חדשים (פתיחת מסמך, עריכה שכבר עודכנה ב-Document API) — למדוד שוב
// את הגיאומטריה גם אם אף אחד מגלילה/שינוי-גודל/`viewport.observe` לא ירה:
// רשת ביטחון לאותה מלכודת שכבר נתפסה ב-DocumentRuler/PageBorderOverlay.
watch(() => props.blocks, () => watcher?.measure());
onBeforeUnmount(stopWatching);

const marks = computed<PilcrowMark[]>(() => {
  if (!props.visible || !props.blocks || !painted.value) return [];
  return painted.value.mode === 'mapped'
    ? buildMappedPilcrowMarks(props.blocks, painted.value.endings)
    : buildPilcrowMarks(props.blocks, painted.value.runs);
});
</script>

<style scoped>
.pilcrow-layer {
  position: absolute;
  inset: 0;
  overflow: hidden;
  pointer-events: none;
  /* מעל תוכן המסמך שהמנוע מצייר, מתחת לכל דיאלוג/תפריט — אותו z-index כמו
     שתי השכבות האחרות. */
  z-index: 1;
}

.pilcrow-layer__mark {
  position: absolute;
  box-sizing: border-box;
  font-family: var(--font-main);
  color: var(--color-on-surface-variant);
  white-space: nowrap;
  user-select: none;
}

/* RTL: הקצה שנמדד (`anchorXPx`) הוא ה-`left` של התו האחרון, כלומר ה-¶ צריך
   להסתיים שם ולא להתחיל שם — הזזה מלאה של רוחב עצמו שמאלה. */
.pilcrow-layer__mark--rtl {
  transform: translateX(-100%);
}
</style>
