<template>
  <div
    ref="containerRef"
    class="color-picker-container"
    :class="{ 'color-picker-container--large': isLarge }"
  >
    <div :class="wrapperClass">
      <!--
        שם הצבע בתיאור ובשם הנגיש, ולא רק בפס: הפס הוא הסימן היחיד למה
        שהלחיצה תעשה, והוא ויזואלי בלבד. הכותרת נשארת „צבע גופן” כדי שהחיפוש
        בטולטיפ ובבדיקות ימשיך למצוא את הפקד.
      -->
      <button
        type="button"
        :class="mainClass"
        :data-tip-title="menuString(title)"
        :data-tip-desc="activeColorName"
        :aria-label="`${menuString(title)}, ${activeColorName}`"
        :disabled="disabled"
        @pointerdown.prevent
        @click="applyOnClick ? applyCurrentColor() : toggleDropdown()"
      >
        <!--
          16px ולא 18px: הכפתור גבוה `--ribbon-row-h` (22px), וב-22px נכנסים
          גבול 1 + אייקון 16 + רווח 1 + פס 3 + גבול 1 — ראו ההערה ב-ribbon.css
          שקובעת שאייקון גדול מ-16 בכפתור קטן גולש. עם 18px התוכן היה 24px
          בתוך 20px, והפס — היחיד כאן שמתכווץ — נדחס ל-0 ונעלם מהמסך.
        -->
        <SvgIcon
          :name="icon"
          :size="isLarge ? LARGE_ICON_PX : 16"
        />
        <div
          class="color-indicator-bar"
          :class="{ 'is-none': activeColor === null }"
          :style="{ backgroundColor: activeColor ?? 'transparent' }"
        />
        <span
          v-if="isLarge && label"
          class="btn-label"
        >{{ menuString(label) }}</span>
      </button>
      <button
        type="button"
        :class="arrowClass"
        :disabled="disabled"
        :data-tip-title="menuString('בחירת צבע')"
        :aria-label="menuString('בחירת צבע')"
        @pointerdown.prevent
        @click="toggleDropdown"
      >
        <SvgIcon
          name="chevronDown"
          :size="8"
        />
      </button>
    </div>

    <!--
      פופאובר פלטת הצבעים של Office. `:style` ולא מיקום ב-CSS: `.word-ribbon-body`
      חותך אנכית, ולכן הפופאובר `position: fixed` בקואורדינטות שנמדדות —
      composables/popover-position.ts.
    -->
    <div
      v-if="isOpen"
      ref="popoverRef"
      class="color-palette-popover"
      :style="popoverStyle"
      @pointerdown.prevent.stop
    >
      <!--
        `.prevent` על המעטפת ולא רק על הדוגמיות: לחיצה על כותרת פלטה או
        ברווח שביניהן גזלה את הפוקוס מהמסמך. שדה הצבע המותאם נפתח ב-`click()`
        תכנותי, שביטול `pointerdown` אינו נוגע בו.
      -->
      <div
        v-if="allowClear"
        class="palette-section"
      >
        <button
          type="button"
          class="palette-clear-btn"
          @pointerdown.prevent
          @click="selectColor(null)"
        >
          <span class="clear-icon" />
          {{ menuString(clearLabel) }}
        </button>
      </div>

      <!-- צבעי ערכת נושא (Theme Colors) -->
      <div class="palette-section">
        <div class="palette-title">
          {{ menuString('צבעי ערכת נושא') }}
        </div>
        <div class="theme-colors-grid">
          <div
            v-for="(col, colIndex) in THEME_COLUMNS"
            :key="colIndex"
            class="theme-column"
          >
            <button
              v-for="(hex, rowIndex) in col.shades"
              :key="rowIndex"
              type="button"
              class="color-swatch"
              :class="{ selected: modelValue?.toLowerCase() === hex.toLowerCase() }"
              :style="{ backgroundColor: hex }"
              :data-tip-title="shadeName(col.family, rowIndex)"
              :data-tip-desc="hex"
              :aria-label="shadeName(col.family, rowIndex)"
              @pointerdown.prevent
              @click="selectColor(hex)"
            />
          </div>
        </div>
      </div>

      <!-- צבעים סטנדרטיים (Standard Colors) -->
      <div class="palette-section">
        <div class="palette-title">
          {{ menuString('צבעים רגילים') }}
        </div>
        <div class="standard-colors-row">
          <button
            v-for="color in STANDARD_COLORS"
            :key="color.hex"
            type="button"
            class="color-swatch"
            :class="{ selected: modelValue?.toLowerCase() === color.hex.toLowerCase() }"
            :style="{ backgroundColor: color.hex }"
            :data-tip-title="menuString(color.name)"
            :data-tip-desc="color.hex"
            :aria-label="menuString(color.name)"
            @pointerdown.prevent
            @click="selectColor(color.hex)"
          />
        </div>
      </div>

      <!-- צבע מותאם אישית -->
      <div class="palette-section custom-color-section">
        <label
          class="custom-color-label"
          @pointerdown.prevent="openCustomColorPicker"
        >
          <span>{{ menuString('צבעים נוספים...') }}</span>
          <input
            ref="customColorRef"
            type="color"
            :value="activeColor || defaultColor"
            class="custom-color-input"
            @change="selectColor(($event.target as HTMLInputElement).value)"
          >
        </label>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, inject, ref, shallowRef, onMounted, onUnmounted } from 'vue';
import type { SuperDoc } from 'superdoc';
import { ACTIVE_SUPERDOC } from '../../../engine/document-api';
import { focusDocument } from '../../../engine/focus';
import SvgIcon from '../../icons/SvgIcon.vue';
import { menuString } from '../i18n';
import { usePopoverPosition } from '../../../composables/popover-position';

/**
 * שם הצבע, ולא רק הקוד שלו.
 *
 * `#1f497d` הוא מה שקורא מסך היה מכריז עד עכשיו — כלומר שמונה תווים באיות.
 * שם המשפחה כבר היה כאן כהערה, ומכאן הוא נתון: הוא נכנס לכותרת הטולטיפ
 * ולשם הנגיש, והקוד יורד לשורת ההסבר (הוא עדיין המידע שמעצב מחפש).
 *
 * „גוון 3” ולא „בהיר יותר 60%”: הסדר בעמודה אכן קבוע, אבל האחוזים הם מה
 * ש-Word מחשב מהערכה, ואין כאן חישוב כזה — מספר שאינו נמדד היה דיוק מדומה.
 */
const THEME_COLUMNS = [
  { family: 'לבן ואפור בהיר', shades: ['#ffffff', '#f2f2f2', '#d9d9d9', '#bfbfbf', '#a6a6a6', '#7f7f7f'] },
  { family: 'שחור ואפור כהה', shades: ['#000000', '#7f7f7f', '#595959', '#3f3f3f', '#262626', '#0c0c0c'] },
  { family: 'חום בהיר', shades: ['#eeece1', '#ddd9c3', '#c4bd97', '#948a54', '#494429', '#1d1b10'] },
  { family: 'כחול כהה', shades: ['#1f497d', '#c6d9f1', '#8db3e2', '#548dd4', '#366092', '#17365d'] },
  { family: 'כחול', shades: ['#4f81bd', '#dce6f1', '#b8cce4', '#95b3d7', '#376092', '#254061'] },
  { family: 'אדום', shades: ['#c0504d', '#f2dcdb', '#e6b8b7', '#da9694', '#963634', '#632423'] },
  { family: 'ירוק זית', shades: ['#9bbb59', '#ebf1dd', '#d7e3bc', '#c3d69b', '#76933c', '#4f6228'] },
  { family: 'סגול', shades: ['#8064a2', '#e5e0ec', '#ccc1da', '#b2a2c7', '#604a7b', '#403151'] },
  { family: 'טורקיז', shades: ['#4bacc6', '#dbeef3', '#b7dde8', '#92cddc', '#31859b', '#215967'] },
  { family: 'כתום', shades: ['#f79646', '#fdeada', '#fbd5b5', '#fac08f', '#e36c09', '#974806'] },
];

/** שמות הצבעים הסטנדרטיים, בסדר שבו הם מוצגים — כמו ב-Word. */
const STANDARD_COLORS = [
  { hex: '#c00000', name: 'אדום כהה' },
  { hex: '#ff0000', name: 'אדום' },
  { hex: '#ffc000', name: 'כתום' },
  { hex: '#ffff00', name: 'צהוב' },
  { hex: '#92d050', name: 'ירוק בהיר' },
  { hex: '#00b050', name: 'ירוק' },
  { hex: '#00b0f0', name: 'תכלת' },
  { hex: '#0070c0', name: 'כחול' },
  { hex: '#002060', name: 'כחול כהה' },
  { hex: '#7030a0', name: 'סגול' },
];

/**
 * „כחול, גוון 3”. הבסיס הוא הראשון בעמודה, ולכן הוא בשם המשפחה בלבד.
 *
 * התרגום כאן ולא בקורא: השם מורכב משם המשפחה ומהמילה „גוון”, ומחרוזת מורכבת
 * לא הייתה מתאימה לשום מפתח במילון. `menuString` נקראת מתוך ה-render של
 * התבנית, ולכן הקריאה כאן עדיין מגיבה לשינוי שפה.
 */
function shadeName(family: string, index: number): string {
  const name = menuString(family);
  return index === 0 ? name : `${name}, ${menuString('גוון')} ${index + 1}`;
}

/**
 * שם הצבע מתוך שתי הפלטות, ובנפילה חזרה הקוד עצמו.
 *
 * צבע מותאם אישית אינו נמצא באף אחת מהן, ואין לו שם — הקוד הוא מה שיש. שתי
 * הפלטות כתובות באותיות קטנות, והערך שמגיע לכאן לא בהכרח (`#FFFF00` הוא
 * ברירת המחדל של הסימון), ולכן ההשוואה על גרסה קטנה.
 */
function colorName(hex: string): string {
  const lower = hex.toLowerCase();
  const standard = STANDARD_COLORS.find((color) => color.hex === lower);
  if (standard) return menuString(standard.name);
  for (const col of THEME_COLUMNS) {
    const index = col.shades.indexOf(lower);
    if (index !== -1) return shadeName(col.family, index);
  }
  return hex.toUpperCase();
}

/**
 * כברירת מחדל `modelValue` הוא **צבע המסמך** — מה שהמנוע מדווח על הבחירה — ולא הצבע של
 * הפקד. הוא מסמן את המשבצת בפלטה („הטקסט המסומן כבר אדום כהה”), וזה כל
 * תפקידו. מה שהכפתור הראשי מחיל, ומה שהפס מתחת לאייקון מראה, הוא `activeColor`
 * שלמטה.
 */
const props = withDefaults(
  defineProps<{
    modelValue?: string;
    icon: string;
    title: string;
    defaultColor?: string;
    /** בבורר של משטח יחיד הפס עוקב אחרי המודל, כולל חזרה לברירת המחדל.
     * בבוררי טקסט נשמרת הבחירה האחרונה בנפרד מצבע הטקסט שליד הסמן. */
    followModelValue?: boolean;
    allowClear?: boolean;
    /**
     * מה שהפריט המנקה אומר, ומה שקורא מסך מכריז כשלא נבחר צבע.
     *
     * „ללא צבע” נכון לטקסט — שם הניקוי באמת מסיר צבע — ואינו נכון לכל צרכן:
     * בבד (רקע העורך, `ViewTab.vue`) הניקוי מחזיר את צבע ברירת המחדל, ומשטח
     * חסר צבע אינו קיים שם כלל. תווית שמבטיחה „ללא צבע” ומחזירה צבע היא
     * תיאור שגוי של מה שהלחיצה עושה, ולכן היא נתונה של הצרכן.
     */
    clearLabel?: string;
    /**
     * האם החצי הראשי **מחיל** את הצבע שהפס מראה, או פותח את הפלטה כמו החץ.
     *
     * הפיצול הוא הדפוס הנכון לצבע של **בחירה**: הפס מראה צבע שהטקסט המסומן
     * עדיין אינו צבוע בו, ולכן ללחיצה יש מה לעשות. הוא אינו נכון לצבע של
     * משטח יחיד — בבד (רקע העורך, `ViewTab.vue`) הפס מראה את הצבע שהבד כבר
     * צבוע בו, ולחיצה עליו מחילה את מה שכבר קיים. נמדד: שער „אין כפתור מת”
     * ב-tests/component/ribbon-tabs.test.ts דיווח על החצי הזה ככפתור שנלחץ
     * ולא קרה דבר.
     *
     * כשהוא כבוי הפקד נשאר בדיוק אותו פקד — הפס עדיין מדווח מה הצבע הנוכחי,
     * וזה המידע שהוא נועד לתת — ושתי חציו פותחים את הפלטה.
     */
    applyOnClick?: boolean;
    /**
     * `'large'` הוא הדפוס של „תאריך ושעה” (RibbonMenuButton עם `split` ו-
     * `variant="large"`): אייקון ותווית בעמודה, ורצועת חץ במלוא הרוחב מתחתיהם.
     * הוא נועד לפקד שיושב **לבדו** בקבוצה, שם פקד בן 22px נראה אבוד בקבוצה
     * בת 70px.
     *
     * הגאומטריה כולה מגיעה מ-`.word-split--large` ב-styles/ribbon.css, ולא
     * מכללים מקבילים כאן — שני דפוסים שנראים אותו דבר ומחושבים בשני מקומות
     * הם בדיוק מה שמפריד ביניהם בשינוי הבא.
     */
    variant?: 'small' | 'large';
    /** התווית שמתחת לאייקון. ב-`'small'` אין לה מקום, והיא אינה מוצגת. */
    label?: string;
    disabled?: boolean;
  }>(),
  {
    modelValue: '',
    defaultColor: '#000000',
    followModelValue: false,
    allowClear: true,
    clearLabel: 'ללא צבע',
    applyOnClick: true,
    variant: 'small',
    label: '',
    disabled: false,
  }
);

const isLarge = computed(() => props.variant === 'large');

/**
 * 26 ולא 32 כמו ב-`RibbonButton` הגדול, וזה מספר **שנמדד** ולא הוקטן לזהירות.
 *
 * לכפתור מפוצל גדול אין שום מרווח: נמדד על „תאריך ושעה” בדפדפן — המעטפת היא
 * 70px (‎`--ribbon-content-h`), גוף הכפתור לוקח 55.2 מהם, ומתוכם 47.2 תוכן
 * (אייקון 32 + רווח 2 + תווית 13.2 — בדיוק). רצועת החץ אינה מקבלת את 22px
 * שהיא מבקשת אלא את מה שנשאר, 12.8, כי היא הפריט היחיד שמתכווץ.
 *
 * כלומר כל פיקסל שנוסיף כאן יורד **מהחץ**. הפס הוא 3px ועוד 2px רווח, ולכן
 * אייקון של 32 היה מוריד את החץ ל-7.8px — צ'יפ בן 8px בקופסה שקטנה ממנו,
 * ורצועה שאינה נראית כמו זו של „תאריך ושעה”. ב-26: 26+2+3+2+13.2 = 46.2,
 * והחץ נשאר ~13.8 — אותה רצועה בדיוק.
 */
const LARGE_ICON_PX = 26;

/**
 * `null` = „ללא צבע”, ולא מחרוזת ריקה. זה החוזה של המנוע: `format.color` /
 * `format.highlight` מתעדים `if (value === null) return { target, value: null }`
 * כמסלול הניקוי, ומחרוזת ריקה נדחית שם במפורש (`value.trim() === ''` → `null`
 * → הפקודה נכשלת סגור). ראו engine/payloads.ts.
 */
const emit = defineEmits<{
  (e: 'update:modelValue', color: string): void;
  (e: 'change', color: string | null): void;
}>();

const containerRef = ref<HTMLElement | null>(null);
const popoverRef = ref<HTMLElement | null>(null);
const customColorRef = ref<HTMLInputElement | null>(null);
const isOpen = ref(false);
const superdoc = inject(ACTIVE_SUPERDOC, shallowRef<SuperDoc | null>(null));

/**
 * שלוש קבוצות המחלקות — המעטפת, הגוף והחץ.
 *
 * בגדול הן הגלובליות של הרצועה (`.word-split*`, `.word-btn.btn-large`),
 * ובקטן המקומיות של הפקד. ולא ערבוב של השתיים: `.color-btn-wrapper` קובע
 * `height: var(--ribbon-row-h)` והוא `scoped`, כלומר ספציפי יותר
 * מ-`.word-split--large` — הוא היה כובש את הגובה בחזרה ל-22px.
 */
const wrapperClass = computed(() =>
  isLarge.value
    ? ['word-split', 'word-split--large', { 'word-split--open': isOpen.value }]
    : ['color-btn-wrapper', { active: isOpen.value }],
);
const mainClass = computed(() => (isLarge.value ? ['word-btn', 'btn-large'] : ['color-main-btn']));
const arrowClass = computed(() => (isLarge.value ? ['word-split__arrow'] : ['color-arrow-btn']));

/**
 * הבחירה האחרונה בפקד הזה. `undefined` = טרם נבחר בו דבר.
 *
 * שלושה מצבים ולא שניים, מפני ש-„ללא צבע” הוא בחירה: `null` פירושו שהלחיצה
 * על הכפתור הראשי **מנקה**, ואינו אותו דבר כמו „טרם נבחר”, שבו הלחיצה מחילה
 * את ברירת המחדל.
 */
const chosen = ref<string | null | undefined>(undefined);

/**
 * הצבע שהלחיצה על הכפתור הראשי תחיל — והצבע שהפס מתחת לאייקון מראה. אלה
 * חייבים להיות אותו ערך: הפס הוא ההבטחה של הכפתור.
 *
 * למה לא צבע המסמך, שהיה כאן קודם: פס שמשקף את הטקסט שהסמן עומד עליו הופך את
 * הכפתור הראשי לחסר משמעות — הוא מחיל על הטקסט את הצבע שכבר יש לו. בסימון זה
 * גם היה בלתי-נראה (ברירת המחדל צהובה בכל מקרה), ובצבע הגופן זה איבד את הצבע
 * שהמשתמש בחר ברגע שהסמן עבר לטקסט שחור. ב-Word הצבע נדבק לכפתור עד הבחירה
 * הבאה, וזה מה שקורה כאן.
 */
const activeColor = computed<string | null>(() =>
  props.followModelValue
    ? props.modelValue || props.defaultColor
    : chosen.value === undefined ? props.defaultColor : chosen.value,
);

/** אותו ערך במילים — לטולטיפ ולקורא מסך, שאינם רואים את הפס. */
const activeColorName = computed(() =>
  activeColor.value === null ? menuString(props.clearLabel) : colorName(activeColor.value),
);

const { popoverStyle } = usePopoverPosition(containerRef, popoverRef, isOpen);

function toggleDropdown(): void {
  if (props.disabled) return;
  isOpen.value = !isOpen.value;
}

/**
 * גם הבחירה מהדו-שיח המקורי מגיעה לכאן, אבל ב-`change` ולא ב-`input`:
 * `input` יורה על כל תזוזה בתוך הדו-שיח, וכל ירייה סגרה את הפופאובר, שלחה
 * פקודת צבע למנוע והחזירה מיקוד למסמך — עוד לפני שהמשתמש סיים לבחור.
 */
function selectColor(hex: string | null): void {
  // הבחירה נדבקת לכפתור: מכאן והלאה זה מה שהפס מראה ומה שלחיצה תחיל, עד
  // הבחירה הבאה. גם „ללא צבע” נדבק — ראו `chosen`.
  chosen.value = hex;
  // `modelValue` נשאר מחרוזת. רק ה-`change`, כלומר מה שהופך ל-payload, נושא
  // את ההבחנה בין „ללא צבע” לבין מחרוזת ריקה.
  emit('update:modelValue', hex ?? '');
  emit('change', hex);
  isOpen.value = false;
  focusDocument(superdoc.value);
}

/**
 * `@pointerdown.prevent` על התווית מונע מהלחיצה לגזול את המיקוד מהעורך — בלעדיו
 * הבחירה במסמך אובדת והצבע לא מוחל על שום דבר. אבל הוא גם מבטל את ההתנהגות
 * המובנית של `label`, שפותחת את ה-`input[type=color]`, ולכן הפתיחה נעשית כאן.
 */
function openCustomColorPicker(): void {
  customColorRef.value?.click();
}

function applyCurrentColor(): void {
  if (props.disabled) return;
  emit('change', activeColor.value);
}

function handleClickOutside(event: MouseEvent): void {
  if (containerRef.value && !containerRef.value.contains(event.target as Node)) {
    isOpen.value = false;
  }
}

onMounted(() => {
  document.addEventListener('pointerdown', handleClickOutside);
});

onUnmounted(() => {
  document.removeEventListener('pointerdown', handleClickOutside);
});
</script>

<style scoped>
.color-picker-container {
  position: relative;
  display: inline-flex;
  flex-shrink: 0;
}

.color-btn-wrapper {
  display: inline-flex;
  align-items: stretch;
  border: 1px solid transparent;
  border-radius: var(--radius-sm);
  transition: background 0.08s, border-color 0.08s;
  /* אותה שורה כמו `.btn-icon-only` ו-`RibbonSelect`: הפקד הזה יושב איתם
     באותה `.word-group-row` ב„גופן”, ו-24px קשיח היה משאיר אותו גבוה מהם
     בשתי נקודות. ראו tokens.css. */
  height: var(--ribbon-row-h);
}

.color-btn-wrapper:hover {
  background: var(--word-btn-hover);
  border-color: var(--color-outline-variant);
}

.color-btn-wrapper.active {
  background: var(--word-btn-active);
  border-color: var(--word-btn-active-border);
}

/* ריפוד אנכי 0, ולא 1px: הגובה של המעטפת קבוע (22px), וכל פיקסל שנלקח כאן
   נלקח מהפס. החשבון המלא — 16 (אייקון) + 1 (רווח) + 3 (פס) = 20, כלומר בדיוק
   מה שנשאר אחרי שני הגבולות של המעטפת. */
.color-main-btn {
  background: none;
  border: none;
  padding: 0 4px;
  cursor: pointer;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  color: var(--color-on-surface);
  border-radius: var(--radius-sm) 0 0 var(--radius-sm);
}

/* הפס שמתחת לאייקון — הצבע שהלחיצה תחיל, ולא צבע הטקסט שהסמן עומד עליו.
   ראו `activeColor`.

   `flex-shrink: 0` אינו הגנה משוערת אלא תיקון של תקלה שנמדדה: ל-`.svg-icon`
   יש `flex-shrink: 0` משלו, ולכן כשהתוכן גלש מגובה הכפתור **הפס** ספג את כל
   הדחיסה — `getBoundingClientRect().height` שלו היה 0 בשני הפקדים, כלומר
   הסימן היחיד לצבע שהלחיצה תחיל לא נראה כלל. הגאומטריה מעליו מדויקת עכשיו,
   וזה שומר שהפס לא יהיה שוב הראשון להיעלם אם אייקון או ריפוד ישתנו. */
/* בווריאנט הגדול הפס רחב כמו האייקון שמעליו (`LARGE_ICON_PX`), ו-`margin-top`
   יורד: ל-`.btn-large` יש `gap: 2px` משלו, ומרווח שני על אותו תפר היה נלקח
   מרצועת החץ — היחידה שמתכווצת שם. ראו ההערה על `LARGE_ICON_PX`.

   כל שאר הגאומטריה של הווריאנט הזה אינה כאן אלא ב-`.word-split--large`
   (styles/ribbon.css), במכוון: זה אותו דפוס של „תאריך ושעה”. */
.color-picker-container--large .color-indicator-bar {
  width: 26px;
  margin-top: 0;
}

.color-indicator-bar {
  width: 16px;
  height: 3px;
  flex-shrink: 0;
  border-radius: 1px;
  margin-top: 1px;
  box-shadow: 0 0 1px rgba(0, 0, 0, 0.4);
}

/* „ללא צבע” הוא היעדר, ולכן הפס ריק — קו היקפי בלבד, כמו `.clear-icon` בפלטה.
   `inset` ולא `border`: הגובה נשאר 3px בדיוק כמו בכל צבע אחר, בלי קפיצה. */
.color-indicator-bar.is-none {
  box-shadow: inset 0 0 0 1px var(--color-outline);
}

.color-arrow-btn {
  background: none;
  border: none;
  padding: 0 2px;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  color: var(--color-on-surface-variant);
  border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
}

/* פופאובר פלטת הצבעים. `top` / `left` / `max-height` מגיעים מ-`:style` — ראו
   popover-position.ts. `overflow-y: auto` הוא הצד השני של אותה החלטה: כשאין
   מקום לגובה המלא הפופאובר נגלל בתוך עצמו, ולא נחתך. */
.color-palette-popover {
  position: fixed;
  z-index: 1000;
  background: var(--color-surface);
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.16);
  padding: 8px;
  min-width: 176px;
  overflow-y: auto;
}

.palette-section {
  margin-bottom: 8px;
}

.palette-section:last-child {
  margin-bottom: 0;
}

.palette-title {
  font-size: 10px;
  font-weight: 600;
  color: var(--color-on-surface-variant);
  margin-bottom: 4px;
}

.palette-clear-btn {
  display: flex;
  align-items: center;
  gap: 6px;
  width: 100%;
  background: none;
  border: 1px solid transparent;
  border-radius: var(--radius-xs);
  padding: 3px 6px;
  font-size: 11px;
  color: var(--color-on-surface);
  cursor: pointer;
  text-align: start;
}

.palette-clear-btn:hover {
  background: var(--word-btn-hover);
  border-color: var(--color-outline-variant);
}

.clear-icon {
  width: 12px;
  height: 12px;
  border: 1px dashed var(--color-outline);
  border-radius: 2px;
}

.theme-colors-grid {
  display: flex;
  gap: 2px;
}

.theme-column {
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.standard-colors-row {
  display: flex;
  gap: 2px;
}

/* המסגרת של המשבצת ב-`--color-outline` ולא בשחור שקוף: משבצת כהה מהפלטה
   (`#000000`, `#1d1b10`) בלעה מסגרת של שחור-12% ונראתה בלי גבול בכלל, ובמצב
   כהה גם המשבצות הבהירות איבדו אותה מול הרקע. הצבעים של המשבצות עצמן הם
   פלטת Office והם **נתון ולא עיצוב** — ראו THEME_COLUMNS/STANDARD_COLORS. */
.color-swatch {
  width: 14px;
  height: 14px;
  border: 1px solid var(--color-outline);
  border-radius: 1px;
  cursor: pointer;
  padding: 0;
  transition: transform 0.08s;
}

.color-swatch:hover {
  transform: scale(1.2);
  z-index: 2;
  box-shadow: 0 0 4px rgba(0, 0, 0, 0.3);
}

.color-swatch.selected {
  outline: 2px solid var(--word-blue);
  outline-offset: 1px;
}

.custom-color-label {
  display: flex;
  align-items: center;
  justify-content: space-between;
  font-size: 11px;
  color: var(--color-on-surface);
  cursor: pointer;
  padding: 2px 4px;
  border-radius: var(--radius-xs);
}

.custom-color-label:hover {
  background: var(--word-btn-hover);
}

.custom-color-input {
  opacity: 0;
  width: 0;
  height: 0;
  position: absolute;
}
</style>
