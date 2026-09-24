<template>
  <button
    type="button"
    class="word-btn"
    :class="[
      `btn-${variant}`,
      { active: active, 'word-btn--icon': !!icon }
    ]"
    :disabled="disabled"
    :aria-label="accessibleName"
    :data-tip-title="tip.title"
    :data-tip-shortcut="tip.shortcut || undefined"
    :data-tip-desc="tip.description || undefined"
    :aria-pressed="ariaPressed"
    @pointerdown.prevent
    @click="$emit('click', $event)"
  >
    <SvgIcon
      v-if="icon"
      :name="icon"
      :size="iconSize"
    />
    <span
      v-if="label && variant !== 'icon-only'"
      class="btn-label"
    >{{ menuString(label) }}</span>
    <slot />
  </button>
</template>

<script setup lang="ts">
import { computed, getCurrentInstance } from 'vue';
import { isToggleButton } from '../aria';
import { menuString } from '../i18n';
import SvgIcon from '../../icons/SvgIcon.vue';
import { shortcutLabel, type ShortcutId } from '../../shortcuts/registry';
import { tipParts } from '../../tooltip/tooltip-content';

const props = withDefaults(
  defineProps<{
    icon?: string;
    label?: string;
    tooltip?: string;
    /**
     * שורת ההסבר בטולטיפ — מה שהפקד *עושה*, מתחת לשמו ולצירוף.
     *
     * כשהיא חסרה היא נגזרת: `tooltip` שאינו זהה ל-`label` הוא ההסבר, וה-`label`
     * הוא הכותרת. הכלל ולמה הוא כזה — ב-ui/tooltip/tooltip-content.ts.
     */
    description?: string;
    /**
     * מזהה מהרג'יסטרי של הקיצורים, לא מחרוזת חופשית. כך אי אפשר להבטיח
     * למשתמש „Ctrl+B” שאין לו מאזין: מזהה שאינו ברג'יסטרי נופל בבנייה.
     */
    shortcutId?: ShortcutId;
    variant?: 'large' | 'small' | 'icon-only';
    active?: boolean;
    disabled?: boolean;
  }>(),
  {
    variant: 'icon-only',
    active: false,
    disabled: false,
  }
);

defineEmits<{
  (e: 'click', event: MouseEvent): void;
}>();

// נמדד פעם אחת ב-setup, לא בכל render: אתר קריאה אינו מוסיף או מסיר את הקישור
// ל-active בזמן ריצה. ההסבר למה vnode.props ולא props — ב-isToggleButton.
const isToggle = isToggleButton(getCurrentInstance()?.vnode.props);

/** undefined מסיר את התכונה: כפתור פעולה אינו מתג, ואינו מדווח מצוב. */
const ariaPressed = computed<'true' | 'false' | undefined>(() => {
  if (!isToggle) return undefined;
  return props.active ? 'true' : 'false';
});

const iconSize = computed(() => {
  if (props.variant === 'large') return 32;
  if (props.variant === 'small') return 16;
  return 18;
});

/**
 * השם שקורא מסך מכריז — ורק כשאין תווית גלויה.
 *
 * כאן היה `title`, וזאת הייתה שגיאה כפולה. ראשית `title` הוא מה שמצייר את
 * המלבן האפור של מערכת ההפעלה מעל הכרטיס המעוצב (ההסבר המלא ב-tooltip-content.ts).
 * שנית הוא ערבב שני דברים: מה שהעכבר מגלה, ומה שקורא מסך מכריז.
 *
 * ההפרדה משפרת גם את הנגישות עצמה. `undefined` בווריאנטים שיש בהם `<span
 * class="btn-label">` הוא בכוונה: לכפתור כזה כבר יש שם מהתוכן שלו, ו-`aria-label`
 * היה **דורס** אותו — כלומר גם מוסיף לו את הצירוף כרעש („מודגש Ctrl+B”), וגם
 * שובר שליטה קולית, שמצפה שהשם יהיה מה שכתוב על הכפתור. `title` לא עשה זאת, כי
 * הוא נדחק לסוף התור של חישוב השם; `aria-label` נמצא בראשו.
 *
 * בכפתור אייקון אין תוכן, ולכן שם ה-`aria-label` הוא כל מה שיש — כולל הצירוף,
 * שהוא מידע ולא רעש כשאין דרך אחרת לדעת אותו.
 */
const accessibleName = computed<string | undefined>(() => {
  if (props.label && props.variant !== 'icon-only') return undefined;

  const base = menuString(props.tooltip || props.label || '');
  if (!props.shortcutId) return base || undefined;
  return `${base} (${shortcutLabel(props.shortcutId)})`;
});

/**
 * שלושת שדות הטולטיפ המעוצב, כתכונות `data-tip-*` שהשכבה קוראת
 * (ui/tooltip/TooltipLayer.vue). הן, ולא `title`, מה שהופך את הכפתור לעוגן.
 */
const tip = computed(() =>
  tipParts({
    label: menuString(props.label || ''),
    tooltip: menuString(props.tooltip || ''),
    description: menuString(props.description || ''),
    shortcut: props.shortcutId ? shortcutLabel(props.shortcutId) : '',
  }),
);
</script>
