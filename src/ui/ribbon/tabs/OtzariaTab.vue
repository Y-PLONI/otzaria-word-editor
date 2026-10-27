<template>
  <div class="ribbon-tab-pane otzaria-tab">
    <!-- הקלדה בדיבור: קבוצה משלה, הראשונה ולכן בקצה הימני. היא אינה פעולה
         מול הקורא כמו שאר הכפתורים כאן, אלא מול שירות מקומי נפרד. -->
    <RibbonGroup
      title="דיבור"
      icon="mic"
    >
      <RibbonButton
        class="ribbon-dictation"
        icon="mic"
        :label="dictationLabel"
        variant="large"
        :active="dictationState === 'listening'"
        :aria-busy="dictationState === 'starting' ? 'true' : 'false'"
        :tooltip="dictationTooltip"
        :disabled="!canDictate"
        @click="$emit('toggle-dictation')"
      />
    </RibbonGroup>
    <!-- שילוב עם אוצריא -->
    <RibbonGroup
      title="אוצריא"
      icon="otzaria"
    >
      <RibbonButton
        icon="book"
        label="ציטוט מהקורא"
        shortcut-id="insert-citation"
        variant="large"
        :tooltip="citationTooltip"
        :disabled="!canInsertCitation"
        @click="$emit('insert-citation')"
      />
      <RibbonButton
        icon="search"
        label="חיפוש באוצריא"
        shortcut-id="search-otzaria"
        variant="large"
        :tooltip="searchTooltip"
        :disabled="!canSearch"
        @click="$emit('search-otzaria')"
      />
      <RibbonButton
        icon="otzaria"
        label="פתח ספרייה"
        shortcut-id="open-library"
        variant="large"
        :tooltip="sdkAvailable ? 'פתיחת ספריית הספרים של אוצריא' : OUTSIDE_OTZARIA"
        :disabled="!sdkAvailable"
        @click="$emit('open-library')"
      />
      <!-- הייצוא יושב כאן ולא ב„קובץ”: הוא פעולה מול אוצריא — הקובץ נכתב
           בפורמט הספרים שלה ונקלט לספרייה — ומי שמחפש אותו מחפש „אוצריא”. -->
      <RibbonButton
        icon="export"
        label="ייצוא לאוצריא"
        variant="large"
        :tooltip="exportBookTooltip"
        :disabled="!hasDocument"
        @click="$emit('export-otzaria')"
      />
      <RibbonButton
        icon="highlight"
        label="השלמה מהספר"
        variant="large"
        :active="bookCompletionEnabled"
        :tooltip="bookCompletionTooltip"
        :disabled="!sdkAvailable"
        @click="$emit('toggle-book-completion')"
      />
    </RibbonGroup>
  </div>
</template>

<script setup lang="ts">
/**
 * „אוצריא” — הלשונית שמחברת את העורך לקורא.
 *
 * שלושת הכפתורים כאן פלטו event, ו-`App.vue` ענה עליו בהודעת סטטוס שמתארת
 * פעולה שלא קרתה („פותח את ספריית אוצריא...”). הפעולה עצמה עוברת עכשיו
 * ב-host/otzaria-reader.ts, וההודעה מגיעה רק כשיש מה לדווח.
 *
 * הזמינות נקבעת כאן ולא ב-`App.vue`, מאותו טעם כמו ב-ReferencesTab: כפתור
 * שאינו יכול לעבוד צריך להיראות כך לפני הלחיצה, לא אחריה.
 */
import { computed, inject, shallowRef, watch } from 'vue';
import type { SuperDoc } from 'superdoc';
import RibbonGroup from '../common/RibbonGroup.vue';
import RibbonButton from '../common/RibbonButton.vue';
import { ACTIVE_SUPERDOC } from '../../../engine/document-api';
import { isAvailable } from '../../../host/otzaria-client';
import { canInsertText } from '../../../host/otzaria-reader';

const props = withDefaults(
  defineProps<{
    bookCompletionEnabled?: boolean;
    dictationState?: 'idle' | 'starting' | 'listening';
    /** המסמך לקריאה בלבד — אין לאן לכתוב, ולכן ההקלטה אינה מתחילה. */
    dictationReadOnly?: boolean;
  }>(),
  { bookCompletionEnabled: false, dictationState: 'idle', dictationReadOnly: false },
);

defineEmits<{
  (e: 'insert-citation'): void;
  (e: 'search-otzaria'): void;
  (e: 'open-library'): void;
  (e: 'export-otzaria'): void;
  (e: 'toggle-book-completion'): void;
  (e: 'toggle-dictation'): void;
}>();

const superdoc = inject(ACTIVE_SUPERDOC, shallowRef<SuperDoc | null>(null));

const OUTSIDE_OTZARIA = 'זמין רק כשהעורך פועל בתוך אוצריא';

/**
 * האם ה-SDK של אוצריא קיים. נקרא פעם אחת ב-setup ולא כערך reactive: הרצועה
 * היא „mount on active” (ראו Ribbon.vue) — הלשונית נבנית רק כשהמשתמש לוחץ
 * עליה, כלומר הרבה אחרי ה-boot, ובאותו רגע התשובה סופית. מחוץ לאוצריא
 * הכפתורים מנוטרלים במקום להיכשל בלחיצה.
 */
const sdkAvailable = isAvailable();

/**
 * השאילתה של „חיפוש באוצריא” היא הטקסט המסומן במסמך, ולכן בלי מסמך פתוח אין
 * מה לחפש. הבחירה עצמה נקראת ברגע הלחיצה (ראו `onSearchOtzaria`) — מנוי על
 * כל שינוי בחירה בשביל צביעת כפתור אינו שווה את המחיר.
 */
const canSearch = computed(() => sdkAvailable && superdoc.value !== null);

/**
 * הציטוט נכנס למסמך דרך `doc.insert`, ולכן השאלה היא האם ה-Document API של
 * המסמך הפתוח חושף אותו ומדווח אותו כזמין — ולא האם יש מסמך. התשובה נשאלת
 * במודול (`canInsertText`), ששואל את מרחב השאלות המשותף.
 *
 * `shallowRef` ולא `computed`, כי הקריאה למנוע א-סינכרונית. ראו ReferencesTab:
 * `generation` הוא מה שמונע מתשובה של מסמך קודם לדרוס את התשובה של המסמך
 * הנוכחי, ו-`false` בזמן ההמתנה הוא הכשל הסגור — כפתור שנראה זמין לפני
 * שהתשובה חזרה הוא בדיוק הכפתור המת.
 */
const canInsertCitation = shallowRef(false);

let generation = 0;

watch(
  superdoc,
  async (host) => {
    const mine = ++generation;
    canInsertCitation.value = false;
    const allowed = sdkAvailable && (await canInsertText(host));
    if (mine === generation) canInsertCitation.value = allowed;
  },
  { immediate: true }
);

const citationTooltip = computed(() => {
  if (!sdkAvailable) return OUTSIDE_OTZARIA;
  if (!canInsertCitation.value) return 'יש לפתוח מסמך שאפשר לכתוב בו';
  return 'הכנסת הקטע המסומן בקורא של אוצריא, עם המקור, במיקום הסמן';
});

/**
 * הייצוא אינו דורש את ה-SDK: מחוץ לאוצריא (`host/dev-stub.ts`) מסלול השמירה
 * ממומש בכפיל, וכך הוא נבדק בדפדפן. מה שהוא כן דורש הוא מסמך פתוח.
 */
const hasDocument = computed(() => superdoc.value !== null);

const exportBookTooltip = computed(() =>
  hasDocument.value
    ? 'שמירת המסמך כספר בפורמט של אוצריא (טקסט עם רמות כותרות), לקליטה בספרייה'
    : 'יש לפתוח מסמך תחילה',
);

const searchTooltip = computed(() => {
  if (!sdkAvailable) return OUTSIDE_OTZARIA;
  if (superdoc.value === null) return 'יש לפתוח מסמך ולסמן בו את הטקסט לחיפוש';
  return 'חיפוש הטקסט המסומן במסמך בכל ספריות אוצריא';
});

/**
 * הקלטה: הכפתור אחד, והוא מתחיל ועוצר. בזמן ההקלטה הוא לחוץ והתווית מציעה
 * לעצור; בזמן ההתחלה (בדיקת השירות) הוא מנוטרל כדי שלא יתחילו פעמיים.
 * מחוץ לאוצריא אין שירות, ולכן הוא מנוטרל כמו שאר הפקדים כאן.
 */
const canDictate = computed(
  () =>
    sdkAvailable &&
    superdoc.value !== null &&
    props.dictationState !== 'starting' &&
    (!props.dictationReadOnly || props.dictationState === 'listening'),
);

const dictationLabel = computed(() =>
  props.dictationState === 'listening' ? 'עצור דיבור' : 'הקלטה',
);

const dictationTooltip = computed(() => {
  if (!sdkAvailable) return OUTSIDE_OTZARIA;
  if (superdoc.value === null) return 'יש לפתוח מסמך ולהציב בו את הסמן';
  if (props.dictationState === 'listening') return 'עצירת ההקלטה אחרי המשפט האחרון';
  if (props.dictationReadOnly) return 'המסמך פתוח לקריאה בלבד';
  return 'הקלדה בדיבור במיקום הסמן. הזיהוי נעשה במחשב, בלי אינטרנט';
});

/**
 * הטולטיפ הוא המקום היחיד שאומר למשתמש איך לקבל את ההצעה — "לחיצה על Tab
 * משלימה" חייב להופיע כאן במפורש, ראו engine/book-completion-overlay.ts.
 */
const bookCompletionTooltip = computed(() => {
  if (!sdkAvailable) return OUTSIDE_OTZARIA;
  return 'בזמן הקלדה, אם הטקסט תואם את הספר הפתוח בקורא — Tab משלים 5 מילים מהמקור';
});
</script>

<style scoped>
.ribbon-tab-pane {
  display: flex;
  align-items: stretch;
  gap: 0;
  height: 100%;
}
</style>
