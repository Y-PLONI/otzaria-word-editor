<template>
  <div
    ref="rootRef"
    class="word-ribbon-group"
    :class="{
      'word-ribbon-group--end': end,
      'word-ribbon-group--collapsed': collapsed,
      'is-open': isOpen,
    }"
  >
    <!-- הצ'יפ שמחליף את הקבוצה כשאין לה מקום. שמו הוא שם הקבוצה, וזה מתנגש
         בשם של פקד („תוכן עניינים” הוא גם קבוצה וגם כפתור) — לכן שערי ה-QA
         מדלגים עליו לפי המחלקה ולא לפי השם (scripts/qa/qa-api.js). -->
    <button
      v-if="collapsed"
      type="button"
      class="word-group-chip"
      :data-tip-title="menuString(title)"
      :data-tip-desc="menuString('אין מקום לקבוצה ברוחב הזה — לחיצה פותחת את הפקדים שלה')"
      :aria-label="menuString(title)"
      aria-haspopup="true"
      :aria-expanded="isOpen ? 'true' : 'false'"
      @pointerdown.prevent
      @click="toggle"
    >
      <!-- 32 ולא 18: הצ'יפ הוא כפתור גדול לכל דבר, כמו ב-Word, והאייקון שלו
           יושב באותו גובה בדיוק של האייקונים הגדולים בקבוצות השכנות. -->
      <SvgIcon
        :name="icon"
        :size="32"
      />
      <!-- U+2060 (word joiner) בין השם לחץ: אלמנט מוטבע הוא נקודת שבירה, ובלעדיו
           החץ יכול לרדת לבדו לשורה שנייה. -->
      <span class="word-group-chip__label">{{ menuString(title) }}&#8288;<SvgIcon
        class="word-group-chip__arrow"
        name="chevronDown"
        :size="8"
      /></span>
    </button>

    <!--
      אותו אלמנט בשני המצבים, ולא שני עצים: העברת התוכן בין מכל מוטבע למכל
      צף הייתה בונה את כל הפקדים מחדש בכל שינוי רוחב — כלומר סוגרת בורר פתוח
      ומאבדת מצב. פרוש הוא `display: contents` (המכל נעלם מהפריסה), ומכווץ
      הוא פופאובר `position: fixed` — ראו composables/popover-position.ts.
    -->
    <div
      v-show="!collapsed || isOpen"
      ref="panelRef"
      class="word-group-panel"
      :style="collapsed ? popoverStyle : undefined"
      @keydown.escape="onEscape"
      @click="onPanelClick"
    >
      <div
        class="word-group-content"
        :class="{ 'column-flow': columnFlow }"
      >
        <slot />
      </div>
      <!-- הכותרת נשארת גם בפופאובר: הקבוצה שנפתחת מהצ'יפ נראית בדיוק כמו
           ברצועה, באותו גובה — כמו ב-Word — ולא כפס פקדים בלי שם. -->
      <div class="word-group-footer">
        <span class="word-group-title">{{ menuString(title) }}</span>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, onMounted, onUnmounted, ref } from 'vue';
import { menuString } from '../i18n';
import { useRibbonOverflow } from '../overflow';
import { usePopoverPosition } from '../../../composables/popover-position';
import SvgIcon from '../../icons/SvgIcon.vue';

/* אין כאן `launcher`: כפתור ההרחבה בפינת הקבוצה הוסר מהעיצוב (b2f0635),
   ה-HTML שלו הוסר (e66dc8f), וה-prop עצמו נשאר אחריו בלי צרכן. הוא הוסר ביחד
   עם כל אתרי הקריאה שהעבירו אותו — prop שאינו מוצהר נוזל ל-`$attrs` ומרונדר
   כתכונת DOM על ה-div של הקבוצה. */
withDefaults(
  defineProps<{
    title: string;
    /**
     * האייקון שמייצג את הקבוצה כשהיא מכווצת. חובה, ולא ברירת מחדל: קבוצה בלי
     * אייקון מזוהה הייתה מתקפלת לצ'יפ שאי אפשר לזהות בשורה של צ'יפים.
     */
    icon: string;
    columnFlow?: boolean;
    /**
     * הצמדה לקצה הרצועה — בעברית, שמאל. מה ש-Word עושה ל„עזרה”: קבוצה שאינה
     * חלק מהזרימה של הפעולות, ולכן יושבת בקצה קבוע ולא בתור.
     *
     * הכלל עצמו (`margin-inline-start: auto`) ב-ribbon.css, וכתוב שם גם התנאי
     * שהוא דורש: `width: 100%` על `.ribbon-tab-pane` של הלשונית. בלי מרווח
     * פנוי אין מה ש-`auto` יבלע, וההצמדה פשוט אינה קורית.
     */
    end?: boolean;
  }>(),
  {
    columnFlow: false,
    end: false,
  }
);

const rootRef = ref<HTMLElement | null>(null);
const panelRef = ref<HTMLElement | null>(null);
const collapsed = ref(false);
const open = ref(false);

/** הפופאובר קיים רק כשהקבוצה מכווצת — פרושה היא פשוט מוטבעת. */
const isOpen = computed(() => collapsed.value && open.value);

const { popoverStyle } = usePopoverPosition(rootRef, panelRef, isOpen);

const overflow = useRibbonOverflow();

function close(): void {
  open.value = false;
}

function toggle(): void {
  open.value = !open.value;
}

function onEscape(event: KeyboardEvent): void {
  if (!isOpen.value) return;
  event.stopPropagation();
  close();
  rootRef.value?.querySelector<HTMLElement>('.word-group-chip')?.focus();
}

/** לחיצה מחוץ לקבוצה סוגרת — כמו בכל פופאובר של הרצועה. */
function onPointerDown(event: PointerEvent): void {
  if (!isOpen.value) return;
  if (rootRef.value?.contains(event.target as Node)) return;
  close();
}

/**
 * פקודה שהופעלה סוגרת את הפופאובר.
 *
 * ההכרעה נדחית בפריים אחד ונבדקת על התוצאה ולא על הפקד: לחיצה שפתחה בורר,
 * תפריט או פלטה משאירה אותם פתוחים בתוך הפאנל, וסגירת הפאנל הייתה מסתירה את
 * מה שהיא בדיוק פתחה — הפופאוברים האלה `position: fixed` **בתוכו**.
 */
const OPEN_INSIDE = '[aria-expanded="true"], .color-btn-wrapper.active';

function onPanelClick(event: MouseEvent): void {
  if (!isOpen.value) return;
  const target = (event.target as Element | null)?.closest(
    'button, [role="option"], [role="menuitem"]',
  );
  if (!target) return;
  requestAnimationFrame(() => {
    if (!isOpen.value || panelRef.value?.querySelector(OPEN_INSIDE)) return;
    close();
  });
}

let unregister: (() => void) | null = null;

onMounted(() => {
  const el = rootRef.value;
  if (el && overflow) unregister = overflow.register({ el, collapsed, close });
  document.addEventListener('pointerdown', onPointerDown);
});

onUnmounted(() => {
  unregister?.();
  document.removeEventListener('pointerdown', onPointerDown);
});
</script>
