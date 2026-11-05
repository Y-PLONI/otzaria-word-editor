<template>
  <div
    ref="rootRef"
    class="ribbon-combo"
    :style="{ width }"
  >
    <input
      ref="inputRef"
      class="ribbon-combo-input"
      type="text"
      role="combobox"
      autocomplete="off"
      spellcheck="false"
      :value="shown"
      :disabled="disabled"
      :placeholder="menuString(placeholder)"
      :data-tip-title="menuString(title)"
      :aria-label="menuString(title)"
      :aria-expanded="open"
      :aria-controls="listId"
      :aria-activedescendant="open && activeIndex >= 0 ? optionId(activeIndex) : undefined"
      aria-autocomplete="list"
      aria-haspopup="listbox"
      :style="previewStyle"
      @focus="onFocus"
      @blur="onBlur"
      @input="onInput"
      @keydown="onKeydown"
    >
    <!--
      `pointerdown.prevent` ולא `click`: בלי מניעת ברירת המחדל הלחיצה מוציאה את
      הפוקוס מהשדה, `blur` סוגר את הרשימה, והפתיחה מיד אחריה נראתה כהבהוב.
      `pointerdown` ולא `mousedown`, כמו בשאר פקדי הרצועה: ביטול `pointerdown`
      מדכא גם את `mousedown` התואם (Pointer Events §11), ולכן מאזין שני על
      `mousedown` היה רץ פעמיים רק ב-jsdom ואף פעם בדפדפן. `click` נוסף עליו
      בשביל הפעלה שאינה מעכבר — ראו `onArrowClick`.
    -->
    <button
      type="button"
      class="ribbon-combo-arrow"
      tabindex="-1"
      :disabled="disabled"
      :aria-label="menuString('פתח את הרשימה')"
      @pointerdown.prevent="toggle"
      @click="onArrowClick"
    >
      <SvgIcon
        name="chevronDown"
        :size="10"
      />
    </button>

    <ul
      v-if="open"
      :id="listId"
      ref="listRef"
      class="ribbon-combo-list"
      :class="{ 'has-sample': showSample, 'has-glyphs': hasGlyphs, specimen: sampleSpecimen }"
      role="listbox"
      :style="[popoverStyle, { minWidth: listMinWidth }]"
      @pointerdown.prevent.stop
    >
      <!--
        גם על הרשימה עצמה ולא רק על הפריטים: לחיצה על כותרת קבוצה או ברווח
        שביניהם הייתה מוציאה את הפוקוס מהשדה, ו-`blur` סוגר את הרשימה.
      -->
      <template
        v-for="(row, i) in built.rows"
        :key="row.type === 'group' ? `g:${row.label}:${i}` : `o:${row.option.value}`"
      >
        <li
          v-if="row.type === 'group'"
          class="ribbon-combo-group"
          role="presentation"
        >
          {{ menuString(row.label) }}
        </li>
        <li
          v-else
          :id="optionId(row.index)"
          class="ribbon-combo-option"
          :class="{
            active: row.index === activeIndex,
            chosen: row.option.value === modelValue,
            hebrew: row.option.hebrew === true,
            unavailable: row.option.unavailable === true,
            pending: row.option.pending === true,
          }"
          role="option"
          :aria-selected="row.option.value === modelValue"
          :data-value="row.option.value"
          :data-group="row.option.group ?? ''"
          :data-availability="availabilityOf(row.option)"
          :data-tip-title="
            row.option.unavailable === true ? menuString(unavailableReason(row.option)) : undefined
          "
          :aria-label="rowAria(row.option)"
          :style="rowStyle(row.option)"
          @pointerdown.prevent.stop="choose(row.option.value)"
          @mousemove="activeIndex = row.index"
        >
          {{ menuString(row.option.label) }}
        </li>
      </template>
      <li
        v-if="built.count === 0"
        class="ribbon-combo-empty"
        role="presentation"
      >
        {{ menuString(emptyText) }}
      </li>
      <!--
        פס הדגימה. `role="presentation"` ו-`aria-hidden` כמו `ribbon-combo-group`
        ו-`ribbon-combo-empty` שמעליו: הוא חוזר על טקסט שהמשתמש עצמו סימן, ואין
        בו מה להכריז. `dir="auto"` ולא ירושה — ראו ההסבר ב-CSS.

        ה-`<span>` הפנימי הוא מה שמאפשר את שלוש השורות במצב הדגימה: הקטיעה שם
        היא `-webkit-line-clamp`, שדורש `display: -webkit-box` על האלמנט
        שנקטע — ו-`text-overflow: ellipsis` של מצב הבחירה חל על מכל **בלוק**,
        כלומר לא היה שורד על אותו אלמנט. במצב הבחירה ה-`<span>` נשאר inline
        ואינו משנה דבר. צומת אחד לכל הרשימה, ולא לכל שורה בה — זה ההבדל
        מהדגימה שבשורות עצמן, שנשארה `::before` בדיוק מהטעם הזה.
      -->
      <li
        v-if="showSample"
        class="ribbon-combo-sample"
        :class="{ unavailable: sampleUnavailable, specimen: sampleSpecimen }"
        role="presentation"
        aria-hidden="true"
        dir="auto"
        :style="sampleStyle"
      >
        <span>{{ samplePending ? menuString('הדגימה טרם נבדקה') : sample }}</span>
      </li>
    </ul>
  </div>
</template>

<script setup lang="ts">
/**
 * בורר עם חיפוש — הפקד של בורר הגופן.
 *
 * ## למה לא `<select>` נייטיב, שהיה כאן קודם
 *
 * מרגע שהמכונה נמנית (engine/system-fonts.ts) הבורר מציג מאות משפחות.
 * ב-`<select>` נייטיב אין חיפוש: יש „קפיצה לאות” שמתאפסת אחרי שנייה, ובעברית
 * היא כמעט חסרת ערך. מי שמחפש את „Narkisim” בין 300 שמות היה גולל.
 *
 * ## ומה שאבד בדרך, ולמה זה בסדר
 *
 * `<select>` נייטיב מצייר את הרשימה שלו **מחוץ** לחלון הדפדפן — היא אינה
 * נחתכת על ידי גבולות הרצועה, והמערכת מטפלת בגלילה. הרשימה כאן היא DOM רגיל,
 * ולכן היא זקוקה ל-`z-index` ול-`max-height` משלה. זה המחיר של חיפוש, והוא
 * שווה אותו: אין דרך לשים שדה קלט בתוך `<select>`.
 *
 * ההכרעות שאינן חיווט — מה נחשב התאמה, מה מדורג לפני מה, ומה Enter מחיל —
 * יושבות ב-`../font-search.ts` ונבדקות שם ישירות.
 *
 * ## שני מצבים: „בחר מהרשימה” ו„הקלד ערך”
 *
 * בבורר הגופן הרשימה היא **המלאי**: שם שאינו בה הוא כמעט תמיד שגיאת הקלדה,
 * ולכן Enter מחיל את ההתאמה המדורגת ראשונה. בבורר הגודל הרשימה היא **הצעה**:
 * 13pt הוא גודל לגיטימי לגמרי שפשוט אינו בסולם של Word, ומי שהקליד אותו
 * מתכוון אליו ולא ל-10 (ההתאמה הראשונה ל„1”).
 *
 * `normalize` הוא מה שמפריד ביניהם — פקד שמקבל אותו מחיל את מה שהוקלד, אחרי
 * שהקורא נרמל אותו (טווח, עיגול). בלעדיו ההתנהגות היא בורר השמות שהייתה כאן
 * מלכתחילה.
 *
 * ## הדגימה העברית, ולמה היא פסאודו-אלמנט
 *
 * שם לטיני של גופן עברי אינו מראה דבר: „Narkisim” ב-Narkisim ו„Gisha” ב-Gisha
 * נראים שניהם כמו אנגלית, ומי שמחפש כתב לספר עברי בוחר לפי שם. לכן שורה של
 * גופן שמכסה עברית (`option.hebrew`, נקבע ב-engine/font-options.ts) מקבלת לפני
 * השם כמה אותיות עבריות **בגופן של השורה עצמה**.
 *
 * הדגימה היא `::before` בגיליון ולא `<span>` בתבנית, וזה עניין של מחיר: הרשימה
 * מגיעה למאות שורות, ו-`<span>` היה מוסיף לכל אחת מהן צומת DOM, קישור מחרוזת
 * ו-patch של Vue. כאן אין אף אחד מהשלושה — התבנית מוסיפה מחלקה אחת, וכל השאר
 * הוא כלל CSS יחיד עם טקסט קבוע.
 *
 * ומה שאין בה גם כן במחיר: הגופן של השורה **כבר** נפתר בשביל השם, ולכן ארבע
 * אותיות נוספות ממנו אינן טעינה חדשה אלא רסטור של ארבעה גליפים. זו גם הסיבה
 * שהדגימה מותנית בכיסוי: בגופן שאין בו עברית היא הייתה נופלת ל-fallback —
 * כלומר מציגה את האותיות של גופן אחר תחת השם הזה, וזה גרוע מכלום.
 *
 * `aria-label` על השורה העברית: Chrome מכליל תוכן של `::before` בשם הנגיש,
 * וקורא מסך היה מכריז „אבגד Narkisim”. הדגימה היא מראה בלבד, והשם הוא השם.
 */
import { computed, inject, nextTick, onUnmounted, ref, shallowRef, watch } from 'vue';
import type { SuperDoc } from 'superdoc';
import { ACTIVE_SUPERDOC } from '../../../engine/document-api';
import { focusDocument } from '../../../engine/focus';
import { usePopoverPosition } from '../../../composables/popover-position';
import { FONT_PROBE_PRIORITIES } from '../../../composables/keys';
import SvgIcon from '../../icons/SvgIcon.vue';
import { menuString } from '../i18n';
import {
  buildComboRows,
  commitValue,
  nextOptionIndex,
  type ComboOption,
} from '../font-search';

const props = withDefaults(
  defineProps<{
    modelValue?: string;
    /** `readonly` — האפשרויות מגיעות מהמנוע ומהמנייה, ואין לפקד רשות לשנותן. */
    options: readonly ComboOption[];
    width?: string;
    disabled?: boolean;
    title?: string;
    /**
     * הופך את הפקד לתיבת **ערך**: מה שהוקלד מוחל אחרי שהפונקציה נרמלה אותו,
     * ו-`null` ממנה = אין מה להחיל, התיבה חוזרת לערך שהיה. ראו `commitValue`.
     */
    normalize?: (typed: string) => string | null;
    /** מה שמוצג כשלשאילתה אין אף התאמה. */
    emptyText?: string;
    /**
     * רוחב מזערי לרשימה. ברירת המחדל היא רוחב שם גופן („Franklin Gothic
     * Medium” אינו נכנס ברוחב התיבה); רשימת מספרים אינה זקוקה לו.
     */
    listMinWidth?: string;
    /**
     * טקסט לפס הדגימה בתחתית הרשימה, בגופן שהסימון עומד עליו. `''` = אין פס.
     *
     * הפקד אינו יודע מאין הטקסט בא — בבורר הגופן זה הטקסט שהמשתמש סימן
     * במסמך (`use-font-controls`), ובבורר הגודל אין פס כלל, מפני שאין מה
     * להראות: „13” אינו נראה אחרת ב-13 נקודות.
     */
    sample?: string;
    /**
     * הגודל שבו פס הדגימה מצייר את הטקסט — מחרוזת CSS מוכנה (`'18px'`).
     *
     * זה מה שהופך את הפס מ„הנה האותיות” ל„הנה איך הטקסט שלי ייראה”: הוא
     * מצייר את הבחירה בגופן שמרחפים עליו **ובגודל שבמסמך**, ולא בגודל קבוע
     * של הרשימה. המרת נקודות→פיקסלים והגבלת הטווח יושבות אצל הקורא
     * (`use-font-controls`), מפני שהן יודעות מהו גודל הבחירה; הפקד רק מצייר.
     * `undefined` = הגודל הקבוע של הפס (בורר הגודל, שאין לו פס בכלל).
     */
    sampleSize?: string;
    /**
     * מה שמוצג בתיבה ריקה. `''` = אין.
     *
     * ברצועה התיבה לעולם אינה ריקה — תמיד יש גופן מוחל — אבל בדיאלוג שמרכיב
     * patch היא **פותחת** ריקה, וריק בלי מילים אינו אומר „לא ייגע”.
     */
    placeholder?: string;
    /**
     * לאן חוזר המיקוד אחרי בחירה, Enter או Escape.
     *
     * `'document'` הוא הרצועה: מי שבחר גופן רוצה להמשיך להקליד בטקסט, וזו
     * הייתה ההתנהגות היחידה כאן. `'stay'` הוא דיאלוג — שם המסמך אינו היעד
     * הבא, ו-`focus()` עליו היה מוציא את המיקוד מדיאלוג שעדיין פתוח: ה-Escape
     * שלו נשען על מיקוד בתוכו, ולכן הדיאלוג היה מפסיק להיסגר במקלדת.
     */
    focusReturn?: 'document' | 'stay';
  }>(),
  {
    modelValue: '',
    width: 'auto',
    disabled: false,
    title: '',
    normalize: undefined,
    emptyText: 'אין גופן בשם הזה',
    listMinWidth: '150px',
    sample: '',
    sampleSize: undefined,
    placeholder: '',
    focusReturn: 'document',
  },
);

const emit = defineEmits<{
  (e: 'update:modelValue', val: string): void;
  /**
   * הסימון ברשימה עבר לאפשרות הזאת. `null` = אין סימון, או שהרשימה נסגרה.
   *
   * הפקד אינו יודע מה עושים עם זה — בבורר הגופן זו התצוגה החיה על המסמך
   * (composables/font-preview.ts), ובבורר הגודל אין מאזין בכלל. פקד שהיה מחיל
   * בעצמו היה גם צריך לדעת מה להחזיר, וזה כבר לא חיווט.
   */
  (e: 'preview', val: string | null): void;
  /** הרשימה נסגרה. `true` = נבחרה אפשרות, ולכן אין מה להחזיר. */
  (e: 'previewEnd', committed: boolean): void;
  /**
   * המשתמש סיים עם הפקד — אישר ערך, או ויתר ב-Escape.
   *
   * למה זה קיים: הפקד הזה הוא `input`, ולכן בניגוד לכפתורי הרצועה (שמונעים
   * את ברירת המחדל ב-`pointerdown` ולעולם אינם לוקחים מיקוד) הוא **חייב**
   * לקחת אותו. מי שמרכיב אותו צריך להחזיר אותו למסמך; בלי זה ההקלדה הבאה
   * נכנסת לתיבה שברצועה ולא לטקסט — „הסמן לא כותב” שדווח ב-Y-PLONI#14 סעיף א.
   *
   * נפלט **לפני** `update:modelValue`, כדי שההחלה תרוץ על מסמך שכבר ממוקד —
   * אותו סדר של `use-context-menu.ts` (`run()`). מה שנמדד, ובניגוד למה
   * שהיה כתוב כאן: הסדר עצמו **אינו** מה שמציל את העיצוב. תיבת הגודל שבתפריט
   * ההקשר מחילה לפני שהמיקוד חוזר, והעיצוב שורד גם שם
   * (`scripts/qa/context-font-focus-probe.mjs`). הקריטי הוא שהמיקוד יחזור
   * **בכלל**: בלעדיו החזרה למסמך היא לחיצת העכבר של המשתמש, ולחיצה קובעת
   * בחירה חדשה — וזו מה שמוחקת עיצוב שהוחל על סמן מכווץ.
   */
  (e: 'done'): void;
}>();

const rootRef = ref<HTMLElement | null>(null);
const inputRef = ref<HTMLInputElement | null>(null);
const listRef = ref<HTMLElement | null>(null);
const open = ref(false);
const activeIndex = ref(-1);
const superdoc = inject(ACTIVE_SUPERDOC, shallowRef<SuperDoc | null>(null));

/**
 * המיקום נמדד ואינו CSS, ומאותו טעם בדיוק שבו שאר הפופאוברים של הרצועה
 * נמדדים: `.word-ribbon-body` מוגדר `overflow-x: auto; overflow-y: hidden`
 * (styles/ribbon.css), והוא חותך אנכית בגובה הרצועה. רשימה של מאות גופנים
 * ב-`position: absolute` נראתה כשלוש שורות עם פס גלילה — בדיוק מה שדווח.
 *
 * ראו composables/popover-position.ts, כולל למה `overflow-y: visible` על
 * הרצועה אינו פתרון.
 */
const { popoverStyle } = usePopoverPosition(rootRef, listRef, open);

/**
 * מה שהוקלד, או `null` כשלא הוקלד דבר.
 *
 * `null` ולא מחרוזת ריקה, וזו ההבחנה שמחזיקה את הפקד: מחרוזת ריקה היא שאילתה
 * לגיטימית („מחקתי את התיבה”) שאמורה להציג את כל הרשימה, ואילו `null` פירושו
 * „לא נגעתי” — ואז התיבה מציגה את הגופן הנוכחי. בלי ההפרדה הזאת התיבה הייתה
 * מתרוקנת בכל פתיחה.
 */
const query = ref<string | null>(null);

/** מזהי DOM ייחודיים למופע — `aria-activedescendant` דורש מזהה אמיתי. */
const uid = `combo-${Math.random().toString(36).slice(2, 9)}`;
const listId = `${uid}-list`;
const optionId = (index: number) => `${uid}-opt-${index}`;

const shown = computed(() => query.value ?? props.modelValue);

/**
 * הסדר והקבוצות כפי שהיו ברגע הפתיחה; תשובות מדידה מתעדכנות גם כשהיא פתוחה.
 *
 * זו רגרסיה שנצפתה, ולא זהירות: המנוע מרכיב את קבוצת „אחרונים” מהגופנים
 * שהמסמך משתמש בהם, והתצוגה החיה משנה בדיוק את זה. כלומר כל שורה שרוחפים
 * מעליה קופצת אחרי רגע לראש הרשימה — והשורה שמתחת לעכבר מתחלפת באמצע התנועה.
 * מי שרצה את השורה השלישית לחץ על משהו אחר לגמרי.
 *
 * ההקפאה נכונה גם בלי התצוגה החיה: רשימה נפתחת שמשנה סדר תחת היד היא רשימה
 * שאי אפשר לכוון אליה. מנייה שנוחתת בזמן שהיא פתוחה תיראה בפתיחה הבאה.
 */
const frozen = shallowRef<readonly ComboOption[] | null>(null);
const built = computed(() => {
  if (!frozen.value) return buildComboRows(props.options, query.value ?? '');
  // סדר וקבוצות נשארים קפואים, אבל תשובות המדידה מגיעות גם כשהבורר פתוח.
  const live = new Map(props.options.map((option) => [option.value, option]));
  const options = frozen.value.map((option) => ({
    ...(live.get(option.value) ?? option), group: option.group, label: option.label,
  }));
  return buildComboRows(options, query.value ?? '');
});

/**
 * מה שהסימון עומד עליו, כל עוד הרשימה פתוחה.
 *
 * `computed` ולא פליטה מכל אתר שמזיז את הסימון, ואלה חמישה: פתיחה, חץ, עכבר,
 * הקלדה וסגירה. חמש פליטות היו חמש הזדמנויות לשכוח אחת.
 */
const highlighted = computed<string | null>(() => {
  if (!open.value) return null;
  for (const row of built.value.rows) {
    if (row.type === 'option' && row.index === activeIndex.value) return row.option.value;
  }
  return null;
});

watch(highlighted, (value) => emit('preview', value));


const prioritizeFonts = inject(FONT_PROBE_PRIORITIES, undefined);
const isFontPicker = computed(() => props.options.some((option) => option.measured !== undefined || option.pending !== undefined));
watch([() => props.modelValue, isFontPicker], ([value, fontPicker]) => {
  if (fontPicker && value) prioritizeFonts?.([value]);
}, { immediate: true });
watch(highlighted, (value) => {
  if (isFontPicker.value && value) prioritizeFonts?.([value]);
});
let visibleFonts: IntersectionObserver | undefined;
let visibilityGeneration = 0;
watch([open, query], async () => {
  const generation = ++visibilityGeneration;
  visibleFonts?.disconnect();
  visibleFonts = undefined;
  if (!open.value || !prioritizeFonts || !isFontPicker.value) return;
  await nextTick();
  if (generation !== visibilityGeneration || !open.value || !listRef.value) return;
  const rows = listRef.value.querySelectorAll<HTMLElement>('[role="option"]');
  if (typeof IntersectionObserver === 'undefined') {
    prioritizeFonts([...rows].slice(0, 12).map((row) => row.dataset.value!));
    return;
  }
  visibleFonts = new IntersectionObserver((entries) => {
    if (generation !== visibilityGeneration || !open.value) return;
    prioritizeFonts(entries.filter((entry) => entry.isIntersecting)
      .map((entry) => (entry.target as HTMLElement).dataset.value!));
  }, { root: listRef.value });
  for (const row of rows) visibleFonts.observe(row);
});
onUnmounted(() => {
  visibilityGeneration += 1;
  visibleFonts?.disconnect();
});

/**
 * מה שנאמר על שורה שהדפדפן אינו פותר את הגופן שלה.
 *
 * טולטיפ ולא טקסט בשורה: השורה היא שם גופן, וטקסט נוסף בתוכה היה דוחף את
 * העמודה ומשנה את גובה הרשימה בדיוק לפי אילו גופנים מותקנים במכונה. הרמז
 * הנראה הוא העמעום (ראו ה-CSS), וההסבר הוא למי שעוצר על השורה.
 *
 * `data-tip-title` ולא `title` מולד — זה חוזה של התוכנה ולא טעם, והוא נאכף
 * בשער `tests/unit/native-title.test.ts`: Blink קורא `title` **בתזוזת העכבר**
 * ולכן המלבן האפור מצויר גם אחרי שהתכונה ירדה, והוא גם מטפס להורים. שכבת
 * הטולטיפ (`ui/tooltip/tooltip-content.ts`) היא המסלול היחיד.
 */
const NOT_INSTALLED = 'הגופן אינו מותקן במכונה — אין דגימה להציג';

/**
 * מה שנאמר על גופן ש**כן** מותקן ושהדפדפן בכל זאת אינו מצייר.
 *
 * שתי מחרוזות ולא אחת, מפני ששני המצבים אומרים למשתמש לעשות דברים שונים:
 * `Aptos` שאינו במכונה יסתדר במכונה אחרת, ו-`Guttman Kav-Light` שמותקן כאן
 * לא — הוא פער בין מנייה שמדווחת שמות GDI לבין הדפדפן שאינו פותר אותם.
 * המשתמש שדיווח על התקלה **התקין** את הגופנים החסרים, ו„אינו מותקן במכונה”
 * היה השקר שהופך את השורה מחסרת-תועלת למבלבלת.
 *
 * מי מקבל איזו — `installedNotDrawable` שבמיזוג (engine/font-options.ts),
 * שדלוק בדיוק על הצירוף הזה.
 */
const INSTALLED_NOT_DRAWABLE = 'הגופן מותקן אך הדפדפן אינו מצייר אותו — אין דגימה להציג';

/** ההסבר שמתאים לשורה — ראו `INSTALLED_NOT_DRAWABLE`. */
const unavailableReason = (option: ComboOption): string =>
  option.installedNotDrawable === true ? INSTALLED_NOT_DRAWABLE : NOT_INSTALLED;

/**
 * ה-`aria-label` של שורה, וזה שדרוש **שני** דברים ולכן אינו ביטוי בתבנית.
 *
 * 1. **דגימת הגליפים.** שורה עברית נושאת `::before` עם „אבגד”, וקורא מסך היה
 *    מכריז „אבגד Narkisim”. התווית המפורשת היא מה שמנקה את זה, וזה היה הכלל
 *    היחיד כאן קודם.
 * 2. **גופן שאינו מותקן.** `data-tip-*` הוא ציור, ואינו נגיש: קורא מסך אינו
 *    רואה אותו, והעמעום גם הוא חזותי בלבד. בלי התווית הזאת העובדה שאין דגימה
 *    לא הייתה מגיעה בכלל למי שאינו רואה — כלומר דווקא למי שהעמעום אינו עוזר לו.
 *
 * `undefined` לשורה רגילה, כדי שהתוכן שלה יוכרז כמו שהוא.
 */
/**
 * מה השורה יודעת על הזמינות שלה — ומי רשאי לבדוק אותה.
 *
 * `pending` טרם נבדק. שלושת המצבים האחרים מאפשרים לשער QA למדוד את המוצר ואת
 * הסביבה: `declared` היא שורה שהוכרזה זמינה בלי מדידה (חמשת הגופנים שאוצריא
 * מזריקה), והיא באמת אינה נפתרת בכרום נקי שאין בו מארח. שער שדרש הסכמה בין
 * ההכרזה לדפדפן על כל שורה נכשל עליהן תמיד — נמדד.
 *
 * הרשימה נגזרת מהמוצר ולא מטבלה בשער, ובכוונה: טבלה כזאת הייתה נשברת בכל
 * גופן חדש, והתשובה הייתה „להרחיב את הטבלה”.
 */
const availabilityOf = (option: ComboOption): 'pending' | 'declared' | 'measured' | 'missing' => {
  if (option.pending === true) return 'pending';
  if (option.measured !== true) return 'declared';
  return option.unavailable === true ? 'missing' : 'measured';
};

const rowAria = (option: ComboOption): string | undefined => {
  if (option.pending === true) return `${menuString(option.label)} — ${menuString('הדגימה טרם נבדקה')}`;
  const unavailable = option.unavailable === true;
  if (!unavailable && option.hebrew !== true) return undefined;
  const name = menuString(option.label);
  return unavailable ? `${name} — ${menuString(unavailableReason(option))}` : name;
};

/**
 * הגופן שפס הדגימה מצייר בו: מה שהסימון עומד עליו, ובהיעדר סימון — הגופן
 * הנבחר, כדי שהפס לא יופיע ריק ברגע הפתיחה.
 *
 * `undefined` כשלאפשרות אין `preview` (כלומר בכל בורר שאינו גופנים) — ואז הפס
 * פשוט מצייר בגופן הרגיל, ולא ב-`font-family: 13`.
 */
const highlightedOption = computed(() => {
  const value = highlighted.value ?? props.modelValue;
  return props.options.find((option) => option.value === value);
});

/**
 * הגופן שמותר לצייר בו אפשרות — ואיפה נאכפת האינווריאנטה „לא זמין ⇒ לא מצויר”.
 *
 * היא נאכפת **פעמיים** בכוונה, וזו אינה הגנה ללא טריגר. `previewFamily` נופל
 * במיזוג (engine/font-options.ts) וזה מכבה את הציור — אבל המיזוג אינו היצרן
 * היחיד של אפשרויות: `withCurrent` (composables/picker-value.ts) מוסיף את
 * הגופן שבמסמך כשאינו ברשימה, והוא נמדד כמסלול שמייצר בדיוק את השורה
 * שהתיקון הזה בא לבטל. לכן התנאי חי גם כאן, במקום שדרכו עובר **כל** ציור.
 *
 * שלושת הקוראים — השורה, פס הדגימה והתיבה הסגורה — מקבלים אותו כלל, כדי
 * שהתשובה תהיה אחת ולא שלוש.
 */
const paintFamily = (option: ComboOption | undefined): string | undefined =>
  option && option.unavailable !== true && option.pending !== true ? option.preview : undefined;

/** מה שהשורה מצוירת בו. `undefined` — ואז אין `style` כלל, ולא `style` ריק. */
const rowStyle = (option: ComboOption): Record<string, string> | undefined => {
  const family = paintFamily(option);
  return family ? { fontFamily: family } : undefined;
};

const sampleFamily = computed<string | undefined>(() => paintFamily(highlightedOption.value));

/**
 * הפס מדגים גופן שאינו מותקן — כלומר אין לו מה להדגים.
 *
 * זה החצי שבלעדיו התיקון היה חסר ערך: הורדת הציור לבדה מותירה את הפס מצייר
 * את הטקסט של המשתמש בגופן הממשק, ב-24px, ונראה בדיוק כמו דגימה אמיתית. הוא
 * מעומעם בדיוק כמו השורה, ומאותו טוקן, כדי שהשניים יאמרו אותו דבר.
 */
const samplePending = computed(() => highlightedOption.value?.pending === true);
const sampleUnavailable = computed(() => highlightedOption.value?.unavailable === true || samplePending.value);

/**
 * הסגנון של פס הדגימה: הגופן שהסימון עומד עליו, ובגודל שבמסמך (`sampleSize`).
 *
 * שני המפתחות מותנים בנפרד: `fontFamily` נופל כשלאפשרות אין `preview` (בורר
 * שאינו גופנים), ו-`fontSize` נופל כשלא הועבר `sampleSize` — ואז הפס נשאר
 * בגודלו הקבוע, כפי שהיה לפני שהגודל האמיתי חובר. `undefined` מלא כששניהם
 * נופלים, כדי לא לכתוב `style` ריק.
 */
const sampleStyle = computed<Record<string, string> | undefined>(() => {
  const style: Record<string, string> = {};
  if (sampleFamily.value) style.fontFamily = sampleFamily.value;
  if (props.sampleSize) style.fontSize = props.sampleSize;
  return Object.keys(style).length > 0 ? style : undefined;
});

/**
 * הפס מוצג רק כשיש גם טקסט וגם רשימה. על „אין גופן בשם הזה” אין לו מה לעשות:
 * אין אפשרות מסומנת, ולכן אין גופן שהוא מדגים.
 */
const showSample = computed(() => props.sample !== '' && built.value.count > 0);

/**
 * האם לרשימה יש עמודת דגימה — כלומר האם היא בורר גופן. השמות נצמדים לקצה
 * הנגדי רק כאן; ברשימת גדלים „12” ו„14” יישארו זה מתחת לזה בצד ההתחלה.
 *
 * `props.options` ולא `built.rows`: החיפוש מסנן, ו„Arial” שהוקלד היה מבטל את
 * העמודה באמצע ההקלדה — כלומר מזיז את כל השמות בשורה שהמשתמש מסתכל עליה.
 */
const hasGlyphs = computed(() => props.options.some((option) => option.hebrew === true));

/**
 * מצב הדגימה: הפס מציג משפט קבוע ולא טקסט של המשתמש, ולכן שלוש שורות בגודל
 * הרשימה במקום שורה אחת בגודל שבמסמך.
 *
 * הסימן הוא היעדר `sampleSize`, ולא דגל נוסף — ראו `use-font-controls`: הגודל
 * שבמסמך **הוא** ההבטחה „הטקסט שלך ייראה כך”, וכשאין בחירה אין מה להבטיח.
 * שתי הידיעות הן אחת, ושני props לאותה ידיעה היו יכולים לסתור זה את זה.
 */
const sampleSpecimen = computed(() => showSample.value && !props.sampleSize);

/** התיבה מציגה את הגופן הנבחר בגופן שלו — כמו ב-Word. לא בזמן הקלדה. */
const previewStyle = computed(() => {
  if (query.value !== null) return undefined;
  const current = props.options.find((option) => option.value === props.modelValue);
  const family = paintFamily(current);
  return family ? { fontFamily: family } : undefined;
});

/** מיקום הערך הנוכחי ברשימה, או -1. */
function indexOfCurrent(): number {
  for (const row of built.value.rows) {
    if (row.type === 'option' && row.option.value === props.modelValue) return row.index;
  }
  return -1;
}

function openList(): void {
  if (props.disabled) return;
  frozen.value = props.options;
  open.value = true;
  // נפתח על הגופן הנוכחי ולא על ראש הרשימה: זה מה שמאפשר לפתוח, ללחוץ חץ
  // פעם אחת ולקבל את השכן — במקום לקפוץ ל-Assistant מכל מקום.
  activeIndex.value = indexOfCurrent();
}

/**
 * `committed` הוא ההבדל בין „בחרתי” לבין „יצאתי”: הראשון משאיר את מה שהוחל,
 * והשני מחזיר את מה שהיה. ברירת המחדל היא היציאה — Esc, `blur`, ולחיצה על
 * החץ — מפני שהיא הרוב, ומפני שהיא גם הצד שבו טעות עולה במסמך.
 */
function closeList(committed = false): void {
  open.value = false;
  activeIndex.value = -1;
  query.value = null;
  frozen.value = null;
  emit('previewEnd', committed);
}

/**
 * יציאה מהשדה — והשאלה היחידה היא לאן המיקוד הולך.
 *
 * שלושת המסלולים שמסיימים עריכה (בחירה, Enter, Escape) עשו את אותם שני
 * הצעדים בשלושה העתקים. הם כאן פעם אחת מפני שנוסף להם תנאי: בדיאלוג אין
 * לאן לצאת — ראו `focusReturn`.
 */
function leaveField(): void {
  if (props.focusReturn === 'stay') return;
  inputRef.value?.blur();
  focusDocument(superdoc.value);
}

function toggle(): void {
  if (open.value) {
    closeList();
    return;
  }
  openList();
  inputRef.value?.focus();
}

/**
 * הפעלת החץ שאינה מעכבר: מקלדת, או `click()` תכנותי.
 *
 * `detail === 0` הוא מה שמפריד ביניהן ללחיצת עכבר אמיתית, וההבחנה נדרשת מפני
 * ש-`pointerdown` כבר טיפל בזו: בלעדיה לחיצת עכבר הייתה פותחת ב-`pointerdown`
 * וסוגרת מיד ב-`click` שאחריו.
 */
function onArrowClick(event: MouseEvent): void {
  if (event.detail !== 0) return;
  toggle();
}

function choose(value: string): void {
  closeList(true);
  emit('done');
  if (value !== props.modelValue) emit('update:modelValue', value);
  leaveField();
}

function onFocus(): void {
  openList();
  // בחירת כל הטקסט: הקלדה מחליפה את השם הקיים במקום להיצמד אליו — התנהגות
  // תיבת הגופן של Word, ומה שהופך „הקלד שלוש אותיות ו-Enter” לזרימה אחת.
  inputRef.value?.select();
}

function onBlur(): void {
  // בלי החלה: יציאה מהשדה אינה בחירה. מי שהקליד ולא אישר חוזר לגופן שהיה,
  // וזה עדיף על להחיל בטעות גופן על טקסט מסומן.
  closeList();
}

function onInput(event: Event): void {
  query.value = (event.target as HTMLInputElement).value;
  // `??=` ולא השמה: הרשימה עשויה להיות פתוחה ומוקפאת כבר, ואיפוס ההקפאה
  // באמצע הקלדה היה מחזיר בדיוק את הרגרסיה שההקפאה נוספה בשבילה.
  //
  // ולמה בכלל יש כאן מסלול פתיחה שאינו `openList`: מיקוד בתיבה פותח ומקפיא,
  // לחיצה על החץ סוגרת (`closeList` מאפס `frozen`) והמיקוד **נשאר** בתיבה.
  // התו הבא נכנס לתיבה סגורה, ובלי השורה הזאת הרשימה נפתחה בלי הקפאה —
  // נמדד: 47 שורות שזזות תחת העכבר. `openList` אינו מתאים כאן, כי הוא מזיז
  // את הסימון לערך הנוכחי ואילו הקלדה מסמנת את ראש התוצאות.
  frozen.value ??= props.options;
  open.value = true;
  // ראש התוצאות ולא „אין סימון”: אחרי הקלדה ההתאמה הראשונה היא מה שמדורג
  // הכי גבוה, ו-Enter אמור להחיל אותה בלי חץ נוסף.
  //
  // בתיבת ערך (`normalize`) זה הפוך בדיוק: סימון אוטומטי היה גורם ל-Enter
  // להחיל את „10” על מי שהקליד „13”. שם הרשימה נשארת הצעה בלבד עד שבוחרים
  // בה בחץ או בעכבר.
  activeIndex.value = props.normalize ? -1 : built.value.count > 0 ? 0 : -1;
}

function onKeydown(event: KeyboardEvent): void {
  if (event.key === 'Escape') {
    if (!open.value) return;
    // עוצר את ההתפשטות: Esc בתוך הרשימה סוגר אותה, ואינו אמור להגיע למי
    // שסוגר דיאלוגים מעל.
    event.stopPropagation();
    event.preventDefault();
    closeList();
    // גם ויתור הוא סיום: מי שלחץ Escape רוצה לחזור לכתוב, לא להישאר בתיבה.
    emit('done');
    leaveField();
    return;
  }

  if (event.key === 'Enter') {
    if (!open.value) return;
    event.preventDefault();
    /*
     * ועוצר את ההתפשטות, בדיוק כמו ה-Escape שמעליו ומאותו טעם: ה-Enter הזה
     * סוגר את הרשימה, והוא כבר נצרך. דיאלוג נותן ל-Enter להפעיל את הכפתור
     * הראשי (composables/dialog-default-action.ts), ולכן בלי העצירה בחירת
     * גופן מהרשימה הייתה גם **מאשרת את הדיאלוג** — כלומר לחיצה אחת שמחילה
     * את מה שהמשתמש רק התחיל לבחור. רשימה סגורה אינה נכנסת לכאן בכלל
     * (`if (!open.value) return`), ולכן Enter „רגיל” בתיבה מאשר כמו תמיד.
     */
    event.stopPropagation();
    const value = commitValue(built.value, activeIndex.value, query.value ?? '', props.normalize);
    if (value !== null) {
      choose(value);
    } else {
      closeList();
      emit('done');
      leaveField();
    }
    return;
  }

  const moved = nextOptionIndex(event.key, activeIndex.value, built.value.count);
  if (moved === null) return;

  event.preventDefault();
  if (!open.value) openList();
  activeIndex.value = moved;
}

/**
 * גלילת האפשרות המסומנת לתוך התצוגה.
 *
 * `block: 'nearest'` ולא `'center'`: האחרון מזיז את הרשימה בכל חץ גם כשהיעד
 * כבר נראה, והתחושה היא של רשימה שקופצת מתחת לאצבע.
 */
watch(activeIndex, async (index) => {
  if (!open.value || index < 0) return;
  await nextTick();
  // `getElementById` ולא בורר CSS: המזהה נבנה כאן מאותיות, ספרות ומקפים
  // (`optionId`), ולכן אין מה לברוח ממנו — ו-`CSS.escape` אינו קיים ב-jsdom,
  // כלומר בורר עם בריחה היה מפיל את בדיקות הקומפוננטה בלי לשפר דבר.
  // `?.scrollIntoView?.` — jsdom אינו מממש אותה כלל, וגלילה שאינה זמינה אינה
  // סיבה להפיל את הפקד.
  document.getElementById(optionId(index))?.scrollIntoView?.({ block: 'nearest' });
});
</script>

<style scoped>
.ribbon-combo {
  position: relative;
  display: inline-flex;
  align-items: center;
}

.ribbon-combo-input {
  background: var(--color-surface);
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: 11px;
  height: 22px;
  padding-inline-start: 6px;
  padding-inline-end: 18px;
  width: 100%;
  outline: none;
  transition: border-color 0.1s;
}

.ribbon-combo-input:hover:not(:disabled) {
  border-color: var(--word-blue);
}

.ribbon-combo-input:focus {
  border-color: var(--word-blue);
  box-shadow: 0 0 0 1px var(--word-blue);
}

.ribbon-combo-input:disabled {
  opacity: 0.4;
  cursor: default;
}

.ribbon-combo-arrow {
  position: absolute;
  inset-inline-end: 2px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 14px;
  height: 18px;
  padding: 0;
  border: none;
  background: none;
  color: var(--color-on-surface-variant);
  cursor: pointer;
}

.ribbon-combo-arrow:disabled {
  opacity: 0.4;
  cursor: default;
}

/*
  אותה שכבה של ColorPickerPopover: הרשימה חייבת לצוף מעל גוף הרצועה ומעל
  המסמך, ו-`<select>` נייטיב קיבל את זה מהמערכת בחינם.
*/
.ribbon-combo-list {
  /*
    `top`/`left`/`max-height` מגיעים מ-`popoverStyle` — ראו ההסבר ב-script.
    מה שנשאר כאן הוא מה שאינו תלוי במדידה.
  */
  z-index: 1000;
  margin: 0;
  padding: 4px 0;
  /* רחב מהתיבה בכוונה: „Franklin Gothic Medium” אינו נכנס ב-130 פיקסלים. */
  width: max-content;
  /* `min-width` מגיע מ-`listMinWidth` — רשימת גדלים צרה מרשימת שמות. */
  max-width: 260px;
  overflow-y: auto;
  list-style: none;
  background: var(--color-surface);
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.16);
}

.ribbon-combo-group {
  padding: 4px 8px 2px;
  font-family: var(--font-main);
  font-size: 10px;
  font-weight: 600;
  color: var(--word-group-label-color);
  border-block-start: 1px solid var(--word-group-border);
}

.ribbon-combo-group:first-child {
  border-block-start: none;
}

/*
  גובה קבוע, ו-flex — שניהם בשביל הדגימה.

  הדגימה גדולה מהשם, ובזרימת שורה רגילה הגובה נגזר מהגליפים: הבסיס של קופסה
  פנימית גבוהה נדחף כלפי מטה, וכל שורה יוצאת בגובה מעט אחר לפי הגופן שלה
  (נמדד: 28, 29 ו-30 פיקסלים באותה רשימה). רשימה שנעה בפיקסל בין שורה לשורה
  נראית רועדת בגלילה.

  `align-items: baseline` הוא מה ששומר על ההיגיון החזותי: הדגימה והשם יושבים
  על אותו קו בסיס, בדיוק כמו בזרימת טקסט — רק בלי שהגובה תלוי בהם.

  `box-sizing: border-box` גלובלי (styles/shell.css), ולכן הגובה בולע את הריפוד.
*/
/*
  ## `content-visibility` — למה שורה אחת כאן שווה 1.5 שניות

  כל שורה בבורר הגופן מצוירת **בגופן של עצמה** (`previewFamily`), וזה כל
  הרעיון: „Narkisim” באנגלית אינו מראה דבר על צורת האותיות. המחיר הוא שהדפדפן
  חייב לטעון **כל** גופן ברשימה כדי לפרוס אותה — גם שורות שמחוץ למסך.

  נמדד באפליקציה על ה-dist הארוז, 251 שורות (המכונה מנתה 287 משפחות),
  דפדפן טרי לכל סבב, עם בקרת רעש:

  | | פתיחה ראשונה | גלילה (חציון/גרוע) |
  |---|---|---|
  | בסיס | 1869ms | 33/35ms |
  | `content-visibility: auto` | **138ms** | 54/128ms |
  | בסיס שוב (בקרה) | 1508ms | 33/34ms |

  כלומר פי עשר בפתיחה, במחיר גלילה שמתייקרת בפעם הראשונה שעוברים על הרשימה
  (אחריה הגופנים טעונים והיא חוזרת לעצמה). הפתיחה היא המתנה ודאית בכל פתיחה;
  הגלילה היא הסתה חד-פעמית של מי שגולל 25 מסכים. זו העסקה.

  `contain-intrinsic-size` הוא מה שמונע קריסת גובה: שורה שדולגה תופסת 28px
  כמו שורה מצוירת, ולכן פס הגלילה אינו קופץ. `auto` שומר את הגובה שנמדד בפעם
  שהשורה כן צוירה.

  ## שתי השערות שנבדקו ונשללו — כדי שלא ייבדקו שוב

  1. **`width: max-content` על הרשימה הוא הגורם.** נשמע מובן מאליו: כדי למצוא
     את השורה הרחבה ביותר צריך למדוד את כולן. בדף מבודד זה אכן נראה כך
     (10401ms מול 2292ms ברוחב קבוע), אבל **באפליקציה זה אינו נכון** — A/B על
     אותו dist נתן 1687 מול 1515 בסבב אחד ו-1316 מול 1698 בסבב השני, כלומר
     ההפרש קטן מהשונות. וגם אינו תנאי ל-`content-visibility`: 138ms נמדדו
     דווקא **עם** `max-content`, והרוחב נשאר יציב תוך גלילה בשני המקרים.
  2. **`align-items: baseline` דורש מטריקות ולכן מכריח טעינה.** נמדד עם
     `center` במקומו: 2926 מול 2292, בתוך הרעש. אינו גורם.

  מה שכן אומת כמנגנון: אותן 251 שורות, כשכולן באותו גופן אחד, נפתחות ב-107ms.
  הגופן הפרטי לכל שורה הוא העלות — לא הרוחב ולא היישור.
*/
.ribbon-combo-option {
  display: flex;
  align-items: baseline;
  height: 28px;
  padding: 3px 8px;
  font-size: 12px;
  line-height: 22px;
  color: var(--color-on-surface);
  cursor: pointer;
  white-space: nowrap;
  content-visibility: auto;
  contain-intrinsic-size: auto 28px;
}

/*
  הגופן שמוחל כרגע — בצבע, ולא במשקל.

  `font-weight: 600` היה כאן, וזה מה שדווח: התצוגה החיה משנה את הגופן שהמנוע
  מדווח, כלומר הסימון הזה נודד לשורה שרוחפים מעליה — ומשקל שמתחלף משנה רוחב,
  כלומר השורות קופצות תחת העכבר. צבע נושא בדיוק את אותה ידיעה בלי לגעת
  בפריסה.
*/
.ribbon-combo-option.chosen {
  color: var(--word-blue);
}

/*
  הדגימה העברית — ההסבר למה היא כאן ולא בתבנית בראש הקומפוננטה.

  התוכן במרכאות כפולות ובכוונה: מחרוזת עברית בגרשיים בתוך `src/ui/ribbon`
  נסרקת בידי tests/unit/menu-locale-coverage.test.ts כמחרוזת תפריט שחייבת
  תרגום, וזו אינה כזו — דגימת גליפים עבריים נשארת עברית גם בממשק האנגלי, שכן
  מי שבוחר גופן לספר עברי צריך לראות את האותיות ולא את שפת הכפתורים.

  ## הגודל

  18px, ולא גודל השם (12px). דגימה בגודל השם נמדדה כבלתי קריאה — ובגופן עברי
  ההבדל בין David לבין Frank Ruhl הוא בעובי המשיכה ובזווית הראש, וזה בדיוק מה
  שנעלם בגליף של 12 פיקסלים. הדגימה **היא** התוכן כאן; השם הוא התווית.

  הגובה אינו קופץ מזה: `line-height` של השורה קבוע (למעלה), ולכן שורה עם דגימה
  ושורה בלעדיה שוות גובה בדיוק.

  `min-width`: השמות מתחילים באותו קו לאורך הקבוצה, במקום לרקוד לפי רוחב
  הדגימה בכל גופן. `em` ולא פיקסלים — היחידה נגזרת מ-18px של הדגימה עצמה,
  ולכן שינוי הגודל אינו מחייב לחשב מחדש את העמודה.

  `flex-shrink: 0`: השורה היא flex, ורשימה צרה הייתה מכווצת דווקא את הדגימה —
  כלומר מוחקת את מה שהיא קיימת בשבילו.
*/
.ribbon-combo-option.hebrew::before {
  content: "אבגד";
  min-width: 3em;
  margin-inline-end: 8px;
  flex-shrink: 0;
  font-size: 18px;
  line-height: inherit;
}

/*
  שתי עמודות, ולא שתיים צמודות.

  הדגימה נצמדת לצד ההתחלה (ימין בעברית) והשם לצד הסיום (שמאל), והרווח שנובע
  מהפרשי אורך השמות נופל **באמצע** במקום להצטבר בקצה אחד. „Franklin Gothic
  Medium” ו„Arial” מתחילים מאותו קו משני הכיוונים, וזה מה שהופך את הרשימה
  לשתי עמודות אמיתיות: אותיות עבריות מול שמות לטיניים.

  ## ולמה `::before` ריק גם בשורה שאין בה עברית

  `space-between` מניח **פריט יחיד** בצד ההתחלה, ולכן שורה בלי דגימה
  („Arial Black”, וכל גופן לטיני) הייתה נצמדת לימין בזמן שכל שכניה נצמדים
  לשמאל. הפריט הריק הוא מה שמעמיד גם אותה על אותו קו — אותה תחבולה בדיוק
  שכבר משמשת למעלה לשורה שאינה זמינה, ומאותו טעם: העמודה נשארת, גם כשאין בה
  אותיות.

  ורק ברשימה שיש בה בכלל מה להדגים (`has-glyphs`): בבורר הגודל אין דגימה, ושם
  „12” ו„14” נשארים בצד ההתחלה כמו ברשימה רגילה.
*/
.ribbon-combo-list.has-glyphs .ribbon-combo-option {
  justify-content: space-between;
}

.ribbon-combo-list.has-glyphs .ribbon-combo-option:not(.hebrew)::before {
  content: "";
  min-width: 3em;
  margin-inline-end: 8px;
  flex-shrink: 0;
  font-size: 18px;
  line-height: inherit;
}

/*
  שורה שהדפדפן אינו פותר את הגופן שלה, ופס הדגימה כשהוא עומד עליה.

  ## למה עמעום, ולא הסתרה ולא כיבוי

  **לא הסתרה:** השורה היא בחירה חוקית. מסמך נודד, וגופן שאינו כאן עשוי להיות
  מותקן במכונה שבה המסמך ייפתח — וזו בדיוק הסיבה ש-`Aptos` (ברירת המחדל של
  Word 365) נכתבה ברשימה מלכתחילה. הסתרה הייתה מוחקת את הגופן של המסמך מהבורר.

  **לא `disabled`:** אפשר לבחור אותה, וסמן „כבוי” על דבר שנבחר הוא שקר אחר.

  **וכן עמעום**, כלומר `--color-on-surface-variant` — אותו טוקן של
  `.ribbon-combo-empty`, השורה השנייה בפקד הזה שאומרת „אין כאן מה לראות”.
  הצבע הוא כל ההבדל; אין כאן טקסט נוסף, ולכן גובה השורה ורוחב העמודה אינם
  משתנים לפי אילו גופנים מותקנים במכונה. ההסבר עצמו יושב ב-`title`.

  שתי הבוררות באותו כלל בכוונה: השורה והפס חייבים לאמר את אותו דבר. הפס הוא
  החצי שבלעדיו התיקון היה חסר ערך — בלי עמעום הוא מצייר את הטקסט של המשתמש
  בגופן הממשק, בגודל שבמסמך, ונראה כמו דגימה אמיתית לכל דבר.
*/
.ribbon-combo-option.unavailable,
.ribbon-combo-sample.unavailable {
  color: var(--color-on-surface-variant);
}

/*
  וכשהגופן שאינו מותקן הוא **הגופן המוחל** — הכחול מנצח את העמעום.

  זו רגרסיה שנמדדה, לא זהירות: שני הבוררים הם שתי מחלקות (0,2,0), והעמעום
  נכתב אחרי `.chosen`, ולכן הוא היה גובר. התוצאה, בדיוק בתרחיש שהתיקון נבנה
  בשבילו — מסמך Word 365 ב-Aptos שנפתח במכונה בלי Office — הייתה שאף שורה
  ברשימה אינה כחולה, כלומר המשתמש אינו רואה איזה גופן מוחל.

  ולמה הכחול הוא זה שנשאר: הערת ה-CSS מעל `.chosen` מסבירה שהצבע הוא הערוץ
  **היחיד** שנותר לסימון הזה — `font-weight` הוסר משם מפני שהוא הזיז שורות
  תחת העכבר. לעמעום, לעומת זאת, יש שני ערוצים נוספים על אותה שורה: הטולטיפ
  וה-`aria-label`. ידיעה עם שלושה ערוצים מוותרת לידיעה שיש לה אחד.
*/
.ribbon-combo-option.chosen.unavailable {
  color: var(--word-blue);
}

/*
  ומקום הציור הרביעי — דגימת „אבגד” — מכובה על שורה לא-זמינה.

  הוא היחיד שאינו נגזר מ-`previewFamily`, ולכן הפלת השדה במיזוג אינה מגיעה
  אליו: הוא `::before` על מחלקת `hebrew` לבדה. על שורה שאינה נפתרת הוא היה
  מצייר את האותיות של **גופן אחר** תחת השם — כלומר בדיוק השקר שהדגל קיים כדי
  למנוע, ובצורה החמורה יותר, שכן הדגימה היא התוכן שם והשם רק התווית.

  `content: ""` ולא `display: none`: ה-`min-width: 3em` נשאר, ולכן השמות
  ממשיכים להתחיל באותו קו לאורך הקבוצה. הסתרה הייתה מזיזה שורה אחת שמאלה
  בדיוק לפי אילו גופנים מותקנים במכונה.
*/
.ribbon-combo-option.unavailable.hebrew::before,
.ribbon-combo-option.pending.hebrew::before {
  content: "";
}

.ribbon-combo-option.active {
  background: var(--word-btn-active);
}

.ribbon-combo-empty {
  padding: 6px 8px;
  font-family: var(--font-main);
  font-size: 11px;
  color: var(--color-on-surface-variant);
}

/*
  פס הדגימה — הטקסט שהמשתמש סימן במסמך, בגופן שהסימון עומד עליו.

  ## למה `sticky` ולא שורה אחרונה

  הרשימה נגללת (`overflow-y: auto`) ובתפריט ההקשר היא מגיעה עד קצה החלון. שורה
  בסוף רשימה נגללת היא שורה שהמשתמש לא רואה מעולם — ודגימה שאינה נראית אינה
  דגימה. `sticky` משאיר אותה במסך בלי מעטפת חדשה: ה-`<ul>` הזה הוא **עצמו**
  אלמנט הפופאובר ש-`usePopoverPosition` ממקם ומודד, ומעטפת הייתה מחייבת להזיז
  את ה-`ref` ואת המידות יחד איתה.

  ## למה `dir="auto"` על האלמנט, ולא ירושה מהמעטפת

  הדגימה היא הטקסט של המשתמש, והעורך הזה מחזיק גם ציטוטים לטיניים בתוך מסמך
  עברי. במעטפת RTL, מחרוזת כמו „Hello world.” מקבלת את הנקודה **בקצה השמאלי**
  של המחרוזת — כלומר הפיסוק בצד הלא נכון. `dir="auto"` פותר כל דגימה לפי
  התו הראשון בעל כיווניות חזקה, ומשאיר עברית `rtl` כפי שהיא.

  הקטיעה בשלוש נקודות היא הגיבוי לתקרת האורך שב-`use-font-controls`: היא חוסמת
  פרק שלם, וזו חוסמת שורה אחת ארוכה.

  ## שני מצבים, ולמה הם שונים

  הפס עונה על שתי שאלות שונות, ולכן הוא נראה אחרת בכל אחת:

  **בחירה** — „איך **הטקסט שלי** ייראה”. שורה אחת, בגודל שבמסמך, נקטעת בשלוש
  נקודות. זה מה שכתוב מכאן ולמטה, ולא נגענו בו.

  **בלי בחירה** (`specimen`) — „אילו אותיות יש בגופן הזה”. אין טקסט של המשתמש
  ואין גודל להדגים, ולכן הפס מציג את הפסוק שב-`composables/font-sample.ts`
  בגודל הקבוע של הרשימה, ובשלוש שורות. ראו למטה.
*/
.ribbon-combo-sample {
  position: sticky;
  inset-block-end: 0;
  /*
    גובה מוצהר, ולא „מה שייצא”: הוא חייב להיות שווה ל-`scroll-padding-block-end`
    למטה, ושני מספרים שנגזרים מגופן ומ-padding אינם שווים בשום מקום שאפשר
    לבדוק. `nowrap` מבטיח שורה אחת, ולכן הגובה קבוע בפועל ולא רק בהצהרה.

    הגובה עלה מ-34 ל-44 כשהפס התחיל לצייר בגודל **שבמסמך** ולא בגודל קבוע
    (`sampleSize`, מ-`use-font-controls`): הגודל שם חסום ל-24px לכל היותר,
    ו-`line-height: 30px` מכיל אותו. 44 ולא 42, כדי שיישאר מרווח אמיתי מעל
    תיבת התוכן (31px אחרי ה-padding וה-border) — ב-42 ה-`scrollHeight` יצא
    שווה בדיוק ל-`clientHeight`, כלומר על גבול החיתוך ובלי שום שגיאת עיגול
    להרשות לעצמנו.

    ## ולמה `line-height` ולא `display: flex`

    המירכוז האנכי המתבקש כאן הוא `flex` עם `align-items: center`, והוא **שובר
    את הקטיעה**: `text-overflow: ellipsis` חל על מכל בלוק בלבד, ובמכל flex
    הטקסט נעטף בפריט אנונימי שאינו מקבל אותו — כלומר שורה ארוכה הייתה נחתכת
    בלי שלוש הנקודות, בשקט. `line-height` השווה לגובה תיבת התוכן ממרכז את
    השורה היחידה (`nowrap` מבטיח אחת) ומשאיר את הפס מכל בלוק.
  */
  box-sizing: border-box;
  height: var(--combo-sample-h);
  padding: 6px 8px;
  line-height: 30px;
  font-size: 14px;
  color: var(--color-on-surface);
  background: var(--color-surface-container-high);
  border-block-start: 1px solid var(--word-group-border);
  text-align: start;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

/*
  מצב הדגימה — הפסוק, בשלוש שורות.

  ## למה שלוש, וזו מדידה ולא טעם

  לשורה בפס יש **227 פיקסלים**: 260 של הרשימה, פחות מסגרת, ריפוד ופס גלילה
  (נמדד באפליקציה — `scripts/qa/font-list-layout-qa.mjs`). וזה פחות ממה שכל
  מועמד דורש; נמדד בכרום, ב-14px:

  | טקסט | Arial | David | Segoe UI | Guttman Yad-Brush |
  |---|---|---|---|---|
  | הפנגרם הקודם, 36 תווים | 231px | 218px | 254px | 303px |
  | הפסוק, 63 תווים | 394px | 373px | 434px | 510px |

  כלומר שורה אחת מעולם לא הספיקה: הפנגרם הקודם נחתך **גם ב-Arial**, ופס דגימה
  שנחתך הוא בדיוק כישלון של מה שהוא קיים בשבילו — האותיות שאינן נראות הן
  האותיות שלא נבדקו. שתי שורות כיסו את הפסוק בגופנים פרופורציונליים, אבל
  `Courier New` — גופן רוחב-קבוע נפוץ — נקטע בהן. שלוש שורות נותנות 681px,
  ולכן מציגות את כל 63 התווים גם בו.

  ## ולמה `line-height: 20px` דווקא

  30px הוא של מצב הבחירה, שם הוא ממרכז שורה אחת בגודל שמגיע ל-24px. שלוש שורות
  בו היו 90px, ולכן `20px` מכיל את 14px של הדגימה בכל גופן שנמדד (תיבת השורה
  הטבעית הגיעה ל-19px ב-Segoe UI וב-Levenim MT). שלוש שורות יוצאות 60px —
  בתוך 63 של תיבת התוכן, עם אותו מרווח אמיתי שגובה 44 משאיר לשורה אחת.

  ‏`line-clamp` הוא מה שמונע שורה רביעית לגלוש מתחת לגבול, ומוסיף את שלוש
  הנקודות בעצמו — `text-overflow` של מצב הבחירה אינו חל על יותר משורה אחת.
*/
.ribbon-combo-sample.specimen {
  line-height: 20px;
  white-space: normal;
  /* ממורכז ולא `start` כמו מצב הבחירה: זה משפט קבוע ולא הטקסט של המשתמש,
     והשורה השלישית שלו קצרה מהשתיים שמעליה בכל גופן. */
  text-align: center;
}

.ribbon-combo-sample.specimen > span {
  display: -webkit-box;
  -webkit-box-orient: vertical;
  -webkit-line-clamp: 3;
  overflow: hidden;
}

/*
  שני מספרים, ושני פגמים שנמדדו בכרום — לא הוסקו.

  ## `padding-block-end: 0` — רצועת ארבעת הפיקסלים

  ל-`.ribbon-combo-list` יש `padding: 4px 0`, ו-`sticky; inset-block-end: 0`
  נצמד לתחתית **תיבת התוכן** ולא לתחתית ה-scrollport. נמדד: הפס נגמר ב-937
  בזמן שהתחתית הפנימית היא 941, ובאותם 4 פיקסלים **הציצה שורה נגללת**
  (‏`Arial Narrow`). איפוס המרווח כשיש פס מזיז את התחתית אליו, ומשאיר את
  המרווח העליון — שם אין פס — כפי שהיה.

  ## `scroll-padding-block-end` — השורה הפעילה מאחורי הפס

  ה-`scrollIntoView({ block: 'nearest' })` שרודף אחרי הסימון (ה-`watch` על
  `activeIndex`) מחנה את השורה הפעילה מתחת לפס: מבחינת הדפדפן היא „נראית”,
  ומבחינת המשתמש היא מוסתרת. נמדד עם 30px — אחרי 31 חצים תחתית השורה הפעילה
  ב-911 והפס מתחיל ב-903, כלומר 8 פיקסלים מאחוריו. הערך חייב להיות **גובה הפס
  המלא**, ומכאן ההצהרה המשותפת.
*/
.ribbon-combo-list.has-sample {
  /*
    גובה הפס — ראו ההערה למעלה. משתנה אחד ולא שני מספרים שצריך לשמור שווים:
    הפס קורא אותו ל-`height`, והרשימה ל-`scroll-padding`, ושניהם משתנים יחד
    כשהפס עובר למצב הדגימה.
  */
  --combo-sample-h: 44px;

  padding-block-end: 0;
  scroll-padding-block-end: var(--combo-sample-h);
}

/* שלוש שורות של 20px, ריפוד ומסגרת — ועוד 3 של מרווח, כמו במצב הבחירה. */
.ribbon-combo-list.has-sample.specimen {
  --combo-sample-h: 76px;
}
</style>
