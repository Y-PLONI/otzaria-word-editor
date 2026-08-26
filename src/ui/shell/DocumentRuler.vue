<template>
  <div
    v-show="visible"
    ref="rootRef"
    class="doc-ruler"
    role="group"
    aria-label="סרגל המידות"
  >
    <div
      v-if="geometry"
      class="doc-ruler__page"
      :style="{ left: `${geometry.pageLeftPx}px`, width: `${geometry.pageWidthPx}px` }"
    >
      <!-- אזור הטקסט. השוליים הם הרקע, והם כהים ממנו — כמו ב-Word. -->
      <div
        class="doc-ruler__text-area"
        :style="{ left: `${geometry.textLeftPx}px`, width: `${geometry.textWidthPx}px` }"
      />

      <!-- שנתות ומספרים -->
      <div
        v-for="(tick, index) in ticks"
        :key="`t${index}`"
        class="doc-ruler__tick"
        :class="`doc-ruler__tick--${tick.kind}`"
        :style="{ left: `${tick.px}px` }"
      >
        <span
          v-if="tick.label"
          class="doc-ruler__number"
        >{{ tick.label }}</span>
      </div>

      <!-- ידיות -->
      <div
        v-for="handle in handles"
        :key="handle.id"
        class="doc-ruler__handle"
        :class="[
          `doc-ruler__handle--${handle.shape}`,
          { 'is-dragging': dragValue?.id === handle.id, 'is-disabled': !canEdit },
        ]"
        role="slider"
        :tabindex="canEdit ? 0 : -1"
        :aria-label="handle.label"
        :aria-disabled="canEdit ? undefined : 'true'"
        :aria-valuemin="0"
        :aria-valuemax="round2(handle.maxValueTwips / unitTwips)"
        :aria-valuenow="round2(handle.valueTwips / unitTwips)"
        :aria-valuetext="measureLabel(handle.valueTwips, unit)"
        :title="`${handle.label}: ${measureLabel(handle.valueTwips, unit)}`"
        :style="{ left: `${handle.px}px` }"
        @pointerdown="onPointerDown(handle, $event)"
        @keydown="onKeyDown(handle, $event)"
      />

      <!-- תווית הגרירה. מופיעה רק בזמן גרירה, ומראה את המידה שתיכתב. -->
      <div
        v-if="draggedHandle"
        class="doc-ruler__readout"
        :style="{ left: `${draggedHandle.px}px` }"
      >
        {{ draggedHandle.label }}: {{ measureLabel(draggedHandle.valueTwips, unit) }}
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
/**
 * סרגל המידות — הפקד שלשונית „תצוגה” מדליקה.
 *
 * ## מה הוא מראה, ומאיפה כל מספר מגיע
 *
 * שנתות בסנטימטרים (או באינצ'ים, לפי `measurementUnit` של המנוע), אזור טקסט
 * בהיר בין שני אזורי שוליים כהים, וארבע ידיות: שולי ההתחלה, שולי הסוף,
 * וכניסת הפסקה בשני הצדדים. כל המספרים מגיעים מהמסמך דרך engine/page-ruler.ts
 * — אין כאן ולו קבוע אחד שמתאר את הדף.
 *
 * ## הכיוון
 *
 * הכול נמדד „מקצה ההתחלה של העמוד”, וההיפוך לפיקסלים קורה במקום אחד בלבד —
 * `pixelOffset` ב-engine/ruler-geometry.ts. במסמך עברי ההתחלה היא הקצה הימני,
 * ולכן ה-0 של הסרגל שם והמספרים גדלים שמאלה. הכיוון מגיע מהמקטע
 * (`sectionDirection`) ולא מכיוון הממשק: מסמך לועזי שנפתח בתוסף עברי חייב
 * לקבל סרגל לועזי.
 *
 * ## למה `left` ולא `inset-inline-start` בסגנונות המחושבים
 *
 * המעטפת כולה `dir="rtl"`, ולכן `inset-inline-start` היה נמדד מימין — ואילו
 * כל המספרים כאן הם פיקסלים **פיזיים** ביחס למלבן העמוד המצויר, כפי שהמנוע
 * ציייר אותו. ערבוב של השניים הוא בדיוק איך שסרגל מגיע לצד הלא נכון של המסך.
 *
 * ## הכתיבה קורית בשחרור, לא בגרירה
 *
 * `setPageMargins` ו-`setIndentation` מפעילים עימוד מחדש של המסמך כולו.
 * כתיבה בכל `pointermove` הייתה מקפיאה את הממשק על מסמך ארוך. לכן הגרירה
 * מצוירת מקומית (`dragValue`), והמנוע מקבל ערך אחד בשחרור — וזו גם ההתנהגות
 * של Word.
 */
import { computed, inject, ref, shallowRef, watch, type Ref } from 'vue';
import type { SuperDoc } from 'superdoc';
import { ACTIVE_SUPERDOC } from '../../engine/document-api';
import { COMMAND_REPORTER, type CommandReporter } from '../../composables/keys';
import type { CommandOutcome } from '../../engine/command-adapter';
import { applyPageMargins } from '../../engine/page-setup';
import {
  applyRulerIndents,
  watchPageRect,
  type PageRect,
  type RulerReading,
  type ViewportSource,
} from '../../engine/page-ruler';
import {
  MIN_TEXT_AREA_TWIPS,
  clampIndent,
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
    /** האם הסרגל מוצג. מגיע ממצב הפקודה `ruler` של המנוע. */
    visible?: boolean;
    /** מצב המסמך כפי שנקרא ב-App.vue. `null` = אין מה לצייר. */
    reading?: RulerReading | null;
    /** ה-host שהמנוע מצייר בתוכו, מ-`paintedHost(ui)`. */
    host?: HTMLElement | null;
    /** ה-controller, בשביל `viewport.observe`. */
    viewportSource?: ViewportSource | null;
    unit?: RulerUnit;
    /**
     * גודל התצוגה באחוזים. אינו משמש לחישוב — המלבן נמדד ממילא — אלא כטריגר:
     * `viewport.observe` של המנוע **אינו** מדווח על שינוי זום (נמדד על ה-dist
     * הארוז: זום 70% הקטין את העמוד, והסרגל נשאר במקומו), ואין אירוע DOM שקורה
     * אז. זה הסימן היחיד שיש לנו.
     */
    zoom?: number;
    /** `false` במסמך שפתוח לקריאה בלבד — אז אין גרירה. */
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

const emit = defineEmits<{
  /** אחרי כתיבה מוצלחת — כדי שהמצב ייקרא מחדש בלי להמתין להשקטה. */
  (e: 'changed'): void;
}>();

const fallbackReporter: CommandReporter = (outcome, id) => {
  if (!outcome.ok) console.warn(`[otzaria-word] ${id}: ${outcome.message}`);
};

const superdoc = inject(ACTIVE_SUPERDOC, shallowRef<SuperDoc | null>(null)) as Ref<SuperDoc | null>;
const report = inject(COMMAND_REPORTER, fallbackReporter);

const rootRef = ref<HTMLElement | null>(null);
const pageRect = shallowRef<PageRect | null>(null);

/** פעולה על המסמך באוויר. שתי גרירות בו-זמנית היו כותבות זו על זו. */
const inFlight = ref(false);

const unit = computed<RulerUnit>(() => props.unit);
const unitTwips = computed(() => twipsPerUnit(unit.value));
const canEdit = computed(() => props.editable && !inFlight.value && props.reading !== null);

/* ------------------------------------------------------------------ */
/* מעקב אחרי המלבן של העמוד                                            */
/* ------------------------------------------------------------------ */

let watcher: ReturnType<typeof watchPageRect> | null = null;

function stopWatching(): void {
  watcher?.dispose();
  watcher = null;
}

/**
 * המעקב נבנה מחדש כשמשתנה ה-host (מסמך שהוחלף), אלמנט הייחוס, או הצגת
 * הסרגל. סרגל מוסתר אינו מודד: `getBoundingClientRect` בכל גלילה של מסמך
 * שאיש אינו רואה הוא עבודה מיותרת.
 */
watch(
  [() => props.host, () => props.visible, rootRef],
  ([host, visible, root]) => {
    stopWatching();
    pageRect.value = null;
    if (!visible || !host || !root) return;
    watcher = watchPageRect({
      host,
      reference: root,
      ui: props.viewportSource,
      onChange: (rect) => {
        pageRect.value = rect;
      },
    });
  },
  { immediate: true, flush: 'post' },
);

// מסמך שהשתנה או פסקה שהתחלפה יכולים לשנות את מלבן העמוד (שוליים, עימוד),
// והמדידה חינמית כשהיא מותנית בשינוי אמיתי.
watch([() => props.reading, () => props.zoom], () => watcher?.measure());

/* ------------------------------------------------------------------ */
/* גיאומטריה                                                           */
/* ------------------------------------------------------------------ */

interface Geometry {
  pageLeftPx: number;
  pageWidthPx: number;
  pageWidthTwips: number;
  direction: 'rtl' | 'ltr';
  /** שולי ההתחלה והסוף, לוגית. */
  startMarginTwips: number;
  endMarginTwips: number;
  /** אזור הטקסט בפיקסלים פיזיים, בתוך מלבן העמוד. */
  textLeftPx: number;
  textWidthPx: number;
  pxPerUnit: number;
}

/**
 * הגיאומטריה, אחרי החלת הגרירה שבאוויר.
 *
 * הגרירה נכנסת **כאן** ולא רק במיקום הידית, ובכוונה: ב-Word אזור הטקסט
 * והמספרים זזים יחד עם הידית, וזה מה שמראה למשתמש מה יקרה עוד לפני שהוא
 * משחרר. מכיוון שהכול נגזר מהמקום הזה, גם הידית עצמה עוקבת אחרי הסמן בלי
 * מסלול ציור שני שיכול להיפרד ממנו.
 */
const geometry = computed<Geometry | null>(() => {
  const reading = props.reading;
  const rect = pageRect.value;
  if (!reading || !rect || !(rect.widthPx > 0)) return null;

  const { pageWidthTwips, direction } = reading.page;
  const drag = dragValue.value;
  const committedStart = direction === 'rtl' ? reading.page.rightTwips : reading.page.leftTwips;
  const committedEnd = direction === 'rtl' ? reading.page.leftTwips : reading.page.rightTwips;

  const startMarginTwips = drag?.id === 'margin-start' ? drag.valueTwips : committedStart;
  const endMarginTwips = drag?.id === 'margin-end' ? drag.valueTwips : committedEnd;
  const rtl = direction === 'rtl';
  const leftTwips = rtl ? endMarginTwips : startMarginTwips;
  const rightTwips = rtl ? startMarginTwips : endMarginTwips;
  const pxPerTwip = rect.widthPx / pageWidthTwips;

  return {
    pageLeftPx: rect.leftPx,
    pageWidthPx: rect.widthPx,
    pageWidthTwips,
    direction,
    startMarginTwips,
    endMarginTwips,
    textLeftPx: leftTwips * pxPerTwip,
    textWidthPx: Math.max(0, (pageWidthTwips - leftTwips - rightTwips) * pxPerTwip),
    pxPerUnit: pxPerTwip * twipsPerUnit(unit.value),
  };
});

/** מיקום פיזי בתוך מלבן העמוד, מערך „מתחילת העמוד”. */
function toPx(twipsFromStart: number): number {
  const geo = geometry.value;
  if (!geo) return 0;
  return pixelOffset(
    twipsFromStart,
    { leftPx: 0, widthPx: geo.pageWidthPx, widthTwips: geo.pageWidthTwips },
    geo.direction,
  );
}

/** ההמרה ההפוכה, לגרירה: פיקסל בתוך מלבן העמוד → מרחק מתחילת העמוד. */
function toTwips(pxInPage: number): number {
  const geo = geometry.value;
  if (!geo) return 0;
  return twipsFromPixel(
    pxInPage,
    { leftPx: 0, widthPx: geo.pageWidthPx, widthTwips: geo.pageWidthTwips },
    geo.direction,
  );
}

const ticks = computed<Array<RulerTick & { px: number }>>(() => {
  const geo = geometry.value;
  if (!geo) return [];
  return rulerTicks({
    pageWidthTwips: geo.pageWidthTwips,
    textStartTwips: geo.startMarginTwips,
    unit: unit.value,
    pxPerUnit: geo.pxPerUnit,
  }).map((tick) => ({ ...tick, px: toPx(tick.twips) }));
});

/* ------------------------------------------------------------------ */
/* ידיות                                                               */
/* ------------------------------------------------------------------ */

type HandleId = 'margin-start' | 'margin-end' | 'indent-start' | 'indent-end';

interface Handle {
  id: HandleId;
  label: string;
  shape: 'margin' | 'indent-start' | 'indent-end';
  /** המיקום על הסרגל, „מתחילת העמוד”. */
  posTwips: number;
  /** המידה שהמשתמש חושב עליה — גודל השוליים או גודל הכניסה. */
  valueTwips: number;
  /** התקרה של אותה מידה, ל-`aria-valuemax`. */
  maxValueTwips: number;
  px: number;
}

const handles = computed<Handle[]>(() => {
  const geo = geometry.value;
  const reading = props.reading;
  if (!geo || !reading) return [];

  const startLabel = geo.direction === 'rtl' ? 'שוליים ימניים' : 'שוליים שמאליים';
  const endLabel = geo.direction === 'rtl' ? 'שוליים שמאליים' : 'שוליים ימניים';
  const textWidthTwips = geo.pageWidthTwips - geo.startMarginTwips - geo.endMarginTwips;

  const list: Handle[] = [
    {
      id: 'margin-start',
      label: startLabel,
      shape: 'margin',
      posTwips: geo.startMarginTwips,
      valueTwips: geo.startMarginTwips,
      maxValueTwips: Math.max(0, geo.pageWidthTwips - geo.endMarginTwips - MIN_TEXT_AREA_TWIPS),
      px: toPx(geo.startMarginTwips),
    },
    {
      id: 'margin-end',
      label: endLabel,
      shape: 'margin',
      posTwips: geo.pageWidthTwips - geo.endMarginTwips,
      valueTwips: geo.endMarginTwips,
      maxValueTwips: Math.max(0, geo.pageWidthTwips - geo.startMarginTwips - MIN_TEXT_AREA_TWIPS),
      px: toPx(geo.pageWidthTwips - geo.endMarginTwips),
    },
  ];

  // סמני הכניסה קיימים רק כשיש פסקה — כלומר כשהסמן במסמך.
  const committed = reading.indents;
  if (committed) {
    const drag = dragValue.value;
    const indents = {
      leftTwips: drag?.id === 'indent-start' ? drag.valueTwips : committed.leftTwips,
      rightTwips: drag?.id === 'indent-end' ? drag.valueTwips : committed.rightTwips,
    };
    const startPos = geo.startMarginTwips + indents.leftTwips;
    const endPos = geo.pageWidthTwips - geo.endMarginTwips - indents.rightTwips;
    list.push(
      {
        id: 'indent-start',
        label: 'כניסה מצד ההתחלה',
        shape: 'indent-start',
        posTwips: startPos,
        valueTwips: indents.leftTwips,
        maxValueTwips: Math.max(0, textWidthTwips),
        px: toPx(startPos),
      },
      {
        id: 'indent-end',
        label: 'כניסה מצד הסוף',
        shape: 'indent-end',
        posTwips: endPos,
        valueTwips: indents.rightTwips,
        maxValueTwips: Math.max(0, textWidthTwips),
        px: toPx(endPos),
      },
    );
  }

  return list;
});

/**
 * המידה שידית מייצגת, מתוך מיקום על הסרגל. בלי חסמים — אלה ב-`clampValue`.
 *
 * ההפרדה בין השתיים אינה קוסמטית: ההצמדה חייבת לפעול על **המידה** ולא על
 * המיקום. סרגל שמצמיד מיקומים לרשת שמעוגנת בקצה הדף מייצר כניסה של „0.96
 * ס\"מ” כשהמשתמש כיוון ל-1 — כי שולי הדף עצמם אינם על הרשת. מה שהמשתמש קורא
 * בדיאלוג הוא המידה, ולכן היא זו שמתעגלת.
 */
function valueAtPosition(id: HandleId, posTwips: number): number {
  const geo = geometry.value;
  if (!geo) return 0;
  if (id === 'margin-start') return posTwips;
  if (id === 'margin-end') return geo.pageWidthTwips - posTwips;
  if (id === 'indent-start') return posTwips - geo.startMarginTwips;
  return geo.pageWidthTwips - geo.endMarginTwips - posTwips;
}

/** החסמים של אותה מידה: לא שלילית, ולא על חשבון עמודת הטקסט. */
function clampValue(id: HandleId, valueTwips: number): number {
  const geo = geometry.value;
  const reading = props.reading;
  if (!geo || !reading) return 0;
  const textWidthTwips = geo.pageWidthTwips - geo.startMarginTwips - geo.endMarginTwips;

  if (id === 'margin-start') {
    return clampMargin(valueTwips, {
      pageWidthTwips: geo.pageWidthTwips,
      otherMarginTwips: geo.endMarginTwips,
    });
  }
  if (id === 'margin-end') {
    return clampMargin(valueTwips, {
      pageWidthTwips: geo.pageWidthTwips,
      otherMarginTwips: geo.startMarginTwips,
    });
  }
  if (id === 'indent-start') {
    return clampIndent(valueTwips, {
      columnWidthTwips: textWidthTwips,
      otherIndentTwips: reading.indents?.rightTwips ?? 0,
    });
  }
  return clampIndent(valueTwips, {
    columnWidthTwips: textWidthTwips,
    otherIndentTwips: reading.indents?.leftTwips ?? 0,
  });
}

/* ------------------------------------------------------------------ */
/* גרירה                                                               */
/* ------------------------------------------------------------------ */

/** מה שנגרר כרגע. הערך היחיד שנשמר — כל השאר נגזר ממנו ב-`geometry`. */
interface DragValue {
  id: HandleId;
  valueTwips: number;
}

const dragValue = shallowRef<DragValue | null>(null);
/** מזהה המצביע שתפס את הידית. אצבע שנייה על מסך מגע אינה ממשיכה את הגרירה. */
let dragPointerId: number | null = null;

/** הידית שנגררת, עם המיקום שכבר עודכן — לתווית ולסימון החזותי. */
const draggedHandle = computed(() => {
  const drag = dragValue.value;
  if (!drag) return null;
  return handles.value.find((handle) => handle.id === drag.id) ?? null;
});

function pageOriginX(): number {
  const root = rootRef.value;
  const geo = geometry.value;
  if (!root || !geo) return 0;
  return root.getBoundingClientRect().left + geo.pageLeftPx;
}

function onPointerDown(handle: Handle, event: PointerEvent): void {
  if (!canEdit.value || event.button !== 0) return;
  const element = event.currentTarget;
  if (!(element instanceof HTMLElement)) return;

  event.preventDefault();
  // תפיסת המצביע היא שיפור ולא תנאי: הדפדפן זורק `NotFoundError` על מזהה
  // מצביע שכבר אינו פעיל, וההאזנה ב-`window` מכסה את הגרירה בלעדיה ממילא.
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
    // Alt מבטל את ההצמדה — כמו ב-Word, שבו הוא מאפשר מיקום חופשי.
    const raw = valueAtPosition(drag.id, toTwips(moveEvent.clientX - pageOriginX()));
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

  /** Escape או ביטול של המערכת: הערך הישן חוזר, ולמסמך לא נכתב דבר. */
  const cancel = (): void => {
    stop();
    dragValue.value = null;
  };

  window.addEventListener('pointermove', move);
  window.addEventListener('pointerup', finish);
  window.addEventListener('pointercancel', cancel);
}

/* ------------------------------------------------------------------ */
/* מקלדת                                                               */
/* ------------------------------------------------------------------ */

/**
 * החצים מזיזים את הידית **פיזית**: ArrowLeft שמאלה על המסך, ArrowRight ימינה,
 * בשני הכיוונים. זו ההתנהגות שאינה מפתיעה — המשתמש רואה סמן וזז אותו — ולכן
 * ההמרה לערך לוגי נעשית דרך `valueFromPosition`, בדיוק כמו בגרירה.
 *
 * Shift מקטין את הצעד לחמישית, ל-כוונון עדין; Home/End מקפיצים לקצוות.
 */
function onKeyDown(handle: Handle, event: KeyboardEvent): void {
  if (!canEdit.value) return;
  const geo = geometry.value;
  if (!geo) return;

  const step = snapStepUnits(unit.value) * unitTwips.value * (event.shiftKey ? 0.2 : 1);
  const physicalToLogical = geo.direction === 'rtl' ? -1 : 1;
  let posTwips: number | null = null;

  if (event.key === 'ArrowLeft') posTwips = handle.posTwips - step * physicalToLogical;
  else if (event.key === 'ArrowRight') posTwips = handle.posTwips + step * physicalToLogical;
  else if (event.key === 'Home') posTwips = 0;
  else if (event.key === 'End') posTwips = geo.pageWidthTwips;
  if (posTwips === null) return;

  event.preventDefault();
  void commit(handle.id, clampValue(handle.id, valueAtPosition(handle.id, Math.round(posTwips))));
}

/* ------------------------------------------------------------------ */
/* כתיבה למסמך                                                         */
/* ------------------------------------------------------------------ */

async function commit(id: HandleId, valueTwips: number): Promise<void> {
  const reading = props.reading;
  const geo = geometry.value;
  const host = superdoc.value;
  if (!reading || !geo || !host || inFlight.value) return;

  inFlight.value = true;
  try {
    let outcome: CommandOutcome;

    if (id === 'margin-start' || id === 'margin-end') {
      const startMargin = id === 'margin-start' ? valueTwips : geo.startMarginTwips;
      const endMargin = id === 'margin-end' ? valueTwips : geo.endMarginTwips;
      // חזרה מהלוגי לפיזי: זה מה ש-`sections.setPageMargins` מקבל (נמדד).
      const rtl = geo.direction === 'rtl';
      outcome = await applyPageMargins(host, {
        leftTwips: rtl ? endMargin : startMargin,
        rightTwips: rtl ? startMargin : endMargin,
      });
    } else {
      const indents = reading.indents;
      const target = reading.target;
      if (!indents || !target) return;
      outcome = await applyRulerIndents(host, target, indents, {
        startTwips: id === 'indent-start' ? valueTwips : indents.leftTwips,
        endTwips: id === 'indent-end' ? valueTwips : indents.rightTwips,
      });
    }

    report(outcome, id === 'margin-start' || id === 'margin-end' ? 'ruler-margins' : 'ruler-indents');
    if (outcome.ok) emit('changed');
  } finally {
    inFlight.value = false;
  }
}

/** עיגול לתצוגה ב-`aria-valuenow`, שדורש מספר ולא מחרוזת. */
function round2(value: number): number {
  return Math.round(value * 100) / 100;
}
</script>

<style scoped>
/* הרצועה עצמה. גובה קבוע ב-px ולא ביחידות גופן: הסרגל הוא סרגל מדידה, והוא
   לא אמור לגדול עם גודל הגופן של הספרייה — אותה הכרעה כמו בפס הכותרת. */
.doc-ruler {
  position: relative;
  /* פריט בשורת הסרגל: תופס את כל מה שנשאר אחרי הפינה. */
  flex: 1 1 auto;
  min-width: 0;
  height: 22px;
  overflow: hidden;
  background: var(--color-surface-container-highest);
  border-block-end: 1px solid var(--color-outline-variant);
  user-select: none;
  touch-action: none;
}

/* מלבן העמוד. `left` פיזי — ראו הערת הפתיחה של הקומפוננטה. */
/* `box-shadow` ולא `border`: ילדים ממוקמים ביחס לתיבת הריפוד, וגבול של פיקסל
   היה מזיז את כל השנתות והידיות פיקסל אחד מהטקסט. נמדד על ה-dist הארוז —
   אזור הטקסט בסרגל יצא ב-97px מול 96px של הטקסט המצויר. */
.doc-ruler__page {
  position: absolute;
  inset-block: 0;
  background: var(--color-surface-container-high);
  box-shadow:
    inset 1px 0 var(--color-outline-variant),
    inset -1px 0 var(--color-outline-variant);
}

.doc-ruler__text-area {
  position: absolute;
  inset-block: 3px;
  background: var(--color-surface);
  border: 1px solid var(--color-outline-variant);
  border-radius: 1px;
}

/* שנתה. הקו נמשך כלפי מטה מהמרכז, וגובהו הוא הדירוג שלה. */
.doc-ruler__tick {
  position: absolute;
  width: 1px;
  background: var(--color-on-surface-variant);
  opacity: 0.55;
}

.doc-ruler__tick--major {
  top: 4px;
  height: 5px;
  opacity: 0.75;
}

.doc-ruler__tick--mid {
  top: 8px;
  height: 5px;
}

.doc-ruler__tick--minor {
  top: 10px;
  height: 3px;
  opacity: 0.4;
}

/* המספר יושב מעל השנתה הראשית, ממורכז עליה. */
.doc-ruler__number {
  position: absolute;
  bottom: 100%;
  left: 50%;
  transform: translateX(-50%);
  font-family: var(--font-main);
  font-size: 9px;
  line-height: 1;
  color: var(--color-on-surface-variant);
  pointer-events: none;
}

/* ידית. אזור התפיסה רחב מהסמן המצויר: 3px של משולש אינם יעד עכבר. */
.doc-ruler__handle {
  position: absolute;
  width: 14px;
  margin-inline-start: -7px;
  cursor: col-resize;
}

.doc-ruler__handle.is-disabled {
  cursor: default;
  opacity: 0.5;
}

.doc-ruler__handle:focus-visible {
  outline: 2px solid var(--color-primary);
  outline-offset: 1px;
  border-radius: var(--radius-xs);
}

/* ידית שוליים: קו על גבול השוליים, בגובה כל הסרגל. */
.doc-ruler__handle--margin {
  inset-block: 0;
}

.doc-ruler__handle--margin::before {
  content: '';
  position: absolute;
  inset-block: 2px;
  left: 50%;
  width: 3px;
  transform: translateX(-50%);
  border-radius: var(--radius-pill);
  background: var(--color-on-surface-variant);
  opacity: 0.55;
}

.doc-ruler__handle--margin:hover::before,
.doc-ruler__handle--margin.is-dragging::before {
  background: var(--word-blue);
  opacity: 1;
}

/* סמני הכניסה: משולשים, כמו ב-Word. `clip-path` ולא גבולות — הצורה נשארת
   חדה בכל DPI, והצבע מגיע מטוקן אחד ולא מארבעה גבולות. */
.doc-ruler__handle--indent-start,
.doc-ruler__handle--indent-end {
  bottom: 1px;
  height: 8px;
}

.doc-ruler__handle--indent-start::before,
.doc-ruler__handle--indent-end::before {
  content: '';
  position: absolute;
  bottom: 0;
  left: 50%;
  width: 9px;
  height: 8px;
  transform: translateX(-50%);
  background: var(--word-blue);
  clip-path: polygon(50% 0, 100% 100%, 0 100%);
}

.doc-ruler__handle--indent-end::before {
  /* הצד השני — משולש הפוך, כדי שאפשר יהיה להבדיל ביניהם במבט. */
  clip-path: polygon(0 0, 100% 0, 50% 100%);
}

.doc-ruler__handle:hover::before,
.doc-ruler__handle.is-dragging::before {
  filter: brightness(0.85);
}

/* תווית הגרירה. מעל הסרגל ולא בתוכו — 22px אין בהם מקום למספר קריא. */
.doc-ruler__readout {
  position: absolute;
  top: 2px;
  transform: translateX(-50%);
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
