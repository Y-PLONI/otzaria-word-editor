<template>
  <!--
    כפתור ההקלטה הצף: עיגול בפינה התחתונה של אזור המסמך, בצד שאין בו
    חלונית ניווט וסרגל אנכי. ממתין בצבע הנושא של אוצריא, ומקליט באדום —
    כך רואים מכל מקום במסך שהמיקרופון פתוח, גם כשהרצועה מכווצת או בלשונית
    אחרת. הכפתור בלשונית „אוצריא” נשאר, וזה אותו מתג בדיוק.
  -->
  <button
    type="button"
    class="dictation-fab"
    :class="{
      'dictation-fab--on': state === 'listening',
      'dictation-fab--starting': state === 'starting',
    }"
    :data-tip-title="title"
    :data-tip-desc="description"
    :aria-label="title"
    :aria-pressed="state === 'listening'"
    :aria-busy="state === 'starting'"
    :disabled="!available && state === 'idle'"
    @pointerdown.prevent
    @click="$emit('toggle')"
  >
    <SvgIcon
      name="mic"
      :size="22"
    />
  </button>
</template>

<script setup lang="ts">
import { computed } from 'vue';
import SvgIcon from '../icons/SvgIcon.vue';
import { menuString } from '../ribbon/i18n';

const props = withDefaults(
  defineProps<{
    state?: 'idle' | 'starting' | 'listening';
    /**
     * בתוך אוצריא, ויש מסמך פתוח לעריכה. מחוץ לאוצריא אין שירות דיבור.
     * הקלטה שכבר רצה, או שמתחילה, נשארת ניתנת לעצירה גם כשזה `false`.
     */
    available?: boolean;
    /** הטולטיפ כשהכפתור מנוטרל. */
    unavailableReason?: string;
  }>(),
  { state: 'idle', available: false, unavailableReason: 'זמין רק בתוך אוצריא, כשמסמך פתוח' },
);

defineEmits<{ (e: 'toggle'): void }>();

const title = computed(() => menuString(props.state === 'idle' ? 'הקלטה' : 'עצור דיבור'));

const description = computed(() => {
  if (props.state === 'listening') return menuString('עצירת ההקלטה אחרי המשפט האחרון');
  if (props.state === 'starting') return menuString('פותח את המיקרופון. לחיצה מבטלת');
  if (!props.available) return menuString(props.unavailableReason);
  return menuString('הקלדה בדיבור במיקום הסמן. הזיהוי נעשה במחשב, בלי אינטרנט');
});
</script>

<style scoped>
/*
 * המיקום: 32 פיקסלים מעל שורת המצב ו-64 מהקצה — בשוליים האפורים שליד
 * העמוד, ולא צמוד לפינה, שם הוא היה נוגע בפס הגלילה בממשק משמאל לימין.
 * `inset-inline-end` ולא `left`: בממשק בעברית זה השמאל, ובאנגלית — הימין,
 * כמו שאר המעטפת.
 *
 * `z-index: 1` — מעל שכבות המסמך (PageBorderOverlay ואחיותיה, שגם הן 1, והוא
 * אחריהן ב-DOM) ומתחת ללוחות שנחשפים במצב מיקוד (2).
 */
.dictation-fab {
  position: absolute;
  inset-block-end: 32px;
  inset-inline-end: 64px;
  z-index: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  inline-size: 44px;
  block-size: 44px;
  padding: 0;
  border: none;
  border-radius: 50%;
  background: var(--color-primary);
  color: var(--color-on-primary);
  cursor: pointer;
  box-shadow: 0 2px 8px var(--color-scrim);
  transition:
    background 0.15s ease,
    box-shadow 0.15s ease,
    opacity 0.15s ease;
}

.dictation-fab:hover:not(:disabled) {
  background: var(--color-primary-filled-hover);
  box-shadow: 0 4px 14px var(--color-scrim);
}

.dictation-fab:focus-visible {
  outline: 2px solid var(--color-primary);
  outline-offset: 3px;
}

/* מקליט: אדום, עם טבעת שמתרחבת — סימן שהמיקרופון פתוח ולא רק „לחוץ”. */
.dictation-fab--on,
.dictation-fab--on:hover:not(:disabled) {
  background: var(--color-error);
  color: var(--color-on-error);
}

.dictation-fab--on:focus-visible {
  outline-color: var(--color-error);
}

.dictation-fab--on::after {
  content: '';
  position: absolute;
  inset: 0;
  border-radius: 50%;
  border: 2px solid var(--color-error);
  pointer-events: none;
  animation: dictation-fab-pulse 1.6s ease-out infinite;
}

@keyframes dictation-fab-pulse {
  from {
    transform: scale(1);
    opacity: 0.7;
  }
  to {
    transform: scale(1.5);
    opacity: 0;
  }
}

/* בודק את השירות ופותח את המיקרופון: עדיין בצבע הנושא, עמום מעט. לחיצה
   בשלב הזה מבטלת: המיקרופון לא ייפתח. הכפתור חוזר להמתנה כשבדיקת השירות
   נגמרת — מיד בשירות תקין, ועד 5 שניות בשירות שאינו עונה. */
.dictation-fab--starting {
  opacity: 0.75;
}

.dictation-fab:disabled {
  opacity: 0.4;
  cursor: default;
  box-shadow: none;
}

/* במצב מיקוד הכפתור עולה מעל „יציאה ממצב מיקוד”, שיושב באותה פינה. הכלל
   ב-App.vue (`.focus-mode :deep(.dictation-fab)`) ולא כאן: `:global()` בסגנון
   scoped של Vue מחליף את **כל** הסלקטור בתוכן שלו, והכלל היה חל על המעטפת
   עצמה ומזיז אותה. */

/* בחלון צר העמוד מגיע עד הפינה. נמדד בשרת הפיתוח, בזום 100%: ב-1366
   פיקסלים הכפתור (במרחק 64) בשוליים האפורים, אבל ב-1000 הוא כבר על העמוד.
   במרחק 16: ב-1000 מחוץ לעמוד, וב-800 וב-740 על שולי העמוד בלבד — הטקסט
   מתחיל ב-92 וב-69, והכפתור נגמר ב-60. בחלון צר מזה הוא עלול לכסות את קצה
   השורות התחתונות. */
@media (max-width: 1100px) {
  .dictation-fab {
    inset-inline-end: 16px;
  }
}

@media (prefers-reduced-motion: reduce) {
  .dictation-fab {
    transition: none;
  }

  .dictation-fab--on::after {
    animation: none;
    opacity: 0;
  }
}
</style>
