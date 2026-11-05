<template>
  <div class="ribbon-tab-pane developer-tab">
    <!-- מאקרו: הקלטה, ניגון וניהול. הפעולות עצמן ב-engine/macros.ts, דרך App.vue. -->
    <RibbonGroup
      title="מאקרו"
      icon="macro"
    >
      <RibbonButton
        icon="macro"
        label="ניהול מאקרו"
        shortcut-id="macro-manage"
        variant="large"
        :tooltip="macrosTooltip('רשימת המאקרו, קטעי הטקסט והסקריפטים — הרצה, עריכה ושיתוף')"
        :disabled="!macrosAvailable"
        @click="$emit('manage-macros')"
      />
      <RibbonStack>
        <RibbonButton
          :label="isRecording ? 'עצור הקלטה' : 'הקלט מאקרו'"
          shortcut-id="macro-record"
          variant="small"
          :tooltip="macrosTooltip('מקליט את הפעולות במסמך — הקלדה, עיצוב, רשימות — לניגון חוזר')"
          :active="isRecording"
          :disabled="!macrosAvailable"
          @click="$emit('macro-record')"
        />
        <RibbonButton
          label="נגן אחרון"
          shortcut-id="macro-play"
          variant="small"
          :tooltip="macrosTooltip('מריץ את המאקרו האחרון שהוקלט, מהמקום שבו הסמן עומד')"
          :disabled="!macrosAvailable"
          @click="$emit('macro-play')"
        />
      </RibbonStack>
    </RibbonGroup>
  </div>
</template>

<script setup lang="ts">
/**
 * „מפתחים” — לשונית המאקרו, בשם שיש לה ב-Word.
 *
 * שלושת הפקדים ישבו עד עכשיו בלשונית „אוצריא”, ליד הציטוט מהקורא והחיפוש
 * בספרייה, ולא היה להם שם דבר: מערכת המאקרו רצה כולה בעורך ואינה נוגעת
 * ב-SDK של אוצריא — זו בדיוק הסיבה שהיא נשארת חיה גם מחוץ לאוצריא, בזמן
 * ששאר פקדי הלשונית ההיא מנוטרלים. מי שמחפש מאקרו מחפש „מפתחים”, ושם הם
 * מעכשיו — אחרי „תצוגה”, במקום שבו Word מציב את הלשונית הזאת.
 *
 * הזמינות נקבעת כאן ולא ב-`App.vue`, מאותו טעם כמו בשאר הלשוניות: כפתור
 * שאינו יכול לעבוד צריך להיראות כך לפני הלחיצה, לא אחריה.
 */
import { computed, inject, shallowRef } from 'vue';
import RibbonGroup from '../common/RibbonGroup.vue';
import RibbonStack from '../common/RibbonStack.vue';
import RibbonButton from '../common/RibbonButton.vue';
import { ACTIVE_MACROS, type MacrosHandle } from '../../../engine/macros';
import { DOCUMENT_INTERACTION_BLOCKED } from '../../../composables/keys';

defineEmits<{
  (e: 'manage-macros'): void;
  (e: 'macro-record'): void;
  (e: 'macro-play'): void;
}>();

/**
 * מערכת המאקרו של המסמך הפתוח. `null` בלי מסמך — ואז הכפתורים מנוטרלים:
 * כפתור שאינו יכול לעבוד נראה כך לפני הלחיצה.
 */
const macros = inject(ACTIVE_MACROS, shallowRef<MacrosHandle | null>(null));
const blocked = inject(DOCUMENT_INTERACTION_BLOCKED, shallowRef(false));
const macrosAvailable = computed(() => macros.value !== null && !blocked.value);

/**
 * מצב ההקלטה, מתוך ה-ref שה-handle חושף: הכפתור צריך להתחלף ל„עצור הקלטה”
 * גם כשההקלטה התחילה מהמקלדת (Ctrl+Alt+R), לא רק מלחיצה עליו.
 */
const isRecording = computed(() => macros.value?.recording.value ?? false);

function macrosTooltip(available: string): string {
  return macrosAvailable.value ? available : 'יש לפתוח מסמך תחילה';
}
</script>

<style scoped>
.ribbon-tab-pane {
  display: flex;
  align-items: stretch;
  gap: 0;
  height: 100%;
}
</style>
