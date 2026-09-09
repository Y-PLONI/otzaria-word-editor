<template>
  <div
    v-if="isOpen"
    class="modal-backdrop"
    @click.self="$emit('close')"
  >
    <div
      ref="dialogRef"
      class="about-dialog"
      role="dialog"
      aria-modal="true"
      :aria-labelledby="TITLE_ID"
      tabindex="-1"
      @keydown.esc.stop="$emit('close')"
      @keydown.tab="onTab"
    >
      <div class="about-header">
        <div class="about-header-title">
          <SvgIcon
            name="word"
            :size="24"
            class="about-icon"
          />
          <span :id="TITLE_ID">וורד לאוצריא</span>
        </div>
        <button
          type="button"
          class="about-close-btn"
          data-tip-title="סגור"
          data-tip-shortcut="Esc"
          aria-label="סגור את חלון האודות"
          @click="$emit('close')"
        >
          ✕
        </button>
      </div>

      <div class="about-body">
        <p class="about-version">
          גרסה {{ APP_VERSION }} (SuperDoc v2)
        </p>
        <p class="about-desc">
          עורך מסמכי Word (.docx) מתקדם, מעוצב ומותאם במיוחד עבור אפליקציית אוצריא.
        </p>

        <div class="about-details">
          <div class="detail-row">
            <span class="detail-key">מנוע מסמכים:</span>
            <span class="detail-val">SuperDoc {{ ENGINE_VERSION }}</span>
          </div>
          <div class="detail-row">
            <span class="detail-key">תמיכה מלאה:</span>
            <span class="detail-val">עברית, RTL ועימוד עמודים</span>
          </div>
          <div class="detail-row">
            <span class="detail-key">רישיון:</span>
            <span class="detail-val">AGPL-3.0</span>
          </div>
        </div>
      </div>

      <div class="about-footer">
        <button
          ref="primaryRef"
          type="button"
          class="about-btn"
          @click="$emit('close')"
        >
          סגור
        </button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
/**
 * חלון „אודות”.
 *
 * **מה שהיה שבור:** הוא הכריז `role="dialog" aria-modal="true"` והתנהג כמו
 * שכבה מצוירת — בלי שם נגיש, בלי מיקוד ראשוני, ובלי מלכודת מיקוד. כלומר
 * ההצהרה „מודאלי” אמרה לקורא מסך שאין מה לקרוא מחוץ לחלון, בזמן ש-Tab הוציא
 * את המשתמש ישר לרצועה שמאחוריו. `FindReplaceDialog` ו-`LinkDialog` הם
 * התקדים במאגר, ושניהם עושים את שלושת הדברים.
 *
 * Escape מטופל גם גלובלית (`App.vue`), וההאזנה כאן היא בכל זאת נכונה מאותו
 * טעם ששני הדיאלוגים האחרים עושים אותה: הקומפוננטה שמכריזה על עצמה כמודאל
 * אחראית להתנהגות שלה, ולא נשענת על מי שהרכיב אותה. `.stop` כדי שהאירוע לא
 * יגיע פעמיים לאותה סגירה.
 *
 * **הגרסאות:** שני מספרים שקודם היו כתובים בתוך הטקסט. הם עדיין קשיחים — אין
 * ל-`import` של package.json מסלול נקי כאן (`resolveJsonModule` ב-tsconfig
 * ו-`define` ב-vite.config.ts הם שינויי תשתית שנוגעים בשערי הבנייה, ואינם
 * שווים את זה בשביל שני מספרים) — אבל הם קבועים בשם, ומקור האמת מתועד:
 * `version` ו-`dependencies.superdoc` ב-package.json. הדריפט עצמו נתפס
 * בבדיקה: tests/component/dialogs.test.ts משווה את המוצג לקבצי החבילה, ולכן
 * שדרוג גרסה שישכח את הדיאלוג ייפול אדום.
 */
import { nextTick, ref, watch } from 'vue';
import SvgIcon from '../icons/SvgIcon.vue';

/** ראו ההסבר למעלה: מקור האמת הוא package.json. */
const APP_VERSION = '2.0.0';
const ENGINE_VERSION = '2.13.0';

/** מקשר את החלון לכותרת שלו — השם הנגיש שחסר לו. */
const TITLE_ID = 'about-dialog-title';

const props = withDefaults(defineProps<{ isOpen?: boolean }>(), { isOpen: false });

defineEmits<{
  (e: 'close'): void;
}>();

const dialogRef = ref<HTMLElement | null>(null);
const primaryRef = ref<HTMLButtonElement | null>(null);

/** לאן המיקוד חוזר בסגירה — בלעדיו הוא נופל ל-`body` ו-Tab מתחיל מהתחלה. */
let focusOnClose: HTMLElement | null = null;

watch(
  () => props.isOpen,
  (open) => {
    if (open) {
      focusOnClose = document.activeElement instanceof HTMLElement ? document.activeElement : null;
      // אחרי הרינדור: לפניו אין למה למקד.
      void nextTick(() => {
        (primaryRef.value ?? dialogRef.value)?.focus();
      });
      return;
    }

    const target = focusOnClose;
    focusOnClose = null;
    // רק אם הוא עוד במסמך: הרצועה היא „mount on active”, ולשונית שהתחלפה
    // בזמן שהחלון היה פתוח לקחה איתה את הכפתור שנלחץ.
    if (target && document.contains(target)) target.focus();
  },
);

/** הפקדים שאפשר למקד בתוך החלון, בסדר ה-DOM. */
function focusables(): HTMLElement[] {
  const root = dialogRef.value;
  if (!root) return [];
  const selector = 'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';
  return [...root.querySelectorAll<HTMLElement>(selector)].filter(
    (element) => !element.hasAttribute('disabled'),
  );
}

/**
 * מלכודת המיקוד: Tab מהאחרון חוזר לראשון, ו-Shift+Tab מהראשון קופץ לאחרון.
 *
 * מדוע נדרש בכלל: `aria-modal="true"` הוא הצהרה שכל מה שמחוץ לחלון אינו זמין,
 * וקורא מסך מסתיר אותו בהתאם. בלי המלכודת המיקוד יוצא אל תוכן שהוכרז כלא
 * קיים — כלומר הצהרה שקרית, ולא רק אי-נוחות.
 */
function onTab(event: KeyboardEvent): void {
  const items = focusables();
  if (items.length === 0) return;

  const first = items[0];
  const last = items[items.length - 1];
  const active = document.activeElement;

  if (event.shiftKey && (active === first || active === dialogRef.value)) {
    event.preventDefault();
    last.focus();
    return;
  }
  if (!event.shiftKey && active === last) {
    event.preventDefault();
    first.focus();
  }
}
</script>

<style scoped>
.modal-backdrop {
  position: fixed;
  inset: 0;
  background: var(--color-scrim);
  z-index: 3000;
  display: flex;
  align-items: center;
  justify-content: center;
}

.about-dialog {
  background: var(--color-surface);
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-md);
  box-shadow: 0 12px 36px rgba(0, 0, 0, 0.24);
  width: 360px;
  max-width: 90vw;
  font-family: var(--font-main);
  overflow: hidden;
}

.about-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 12px 16px;
  background: var(--color-surface-container-high);
  border-block-end: 1px solid var(--color-outline-variant);
}

.about-header-title {
  display: flex;
  align-items: center;
  gap: 8px;
  font-weight: 600;
  font-size: 14px;
  color: var(--color-on-surface);
}

.about-icon {
  color: var(--word-blue);
}

.about-close-btn {
  background: none;
  border: none;
  font-size: 14px;
  color: var(--color-on-surface-variant);
  cursor: pointer;
  padding: 2px 6px;
  border-radius: var(--radius-xs);
}

.about-close-btn:hover {
  background: var(--word-btn-hover);
  color: var(--color-error);
}

.about-body {
  padding: 16px;
}

.about-version {
  font-size: 12px;
  font-weight: 600;
  color: var(--word-blue);
  margin-bottom: 6px;
}

.about-desc {
  font-size: 12px;
  color: var(--color-on-surface-variant);
  line-height: 1.5;
  margin-bottom: 14px;
}

.about-details {
  background: var(--color-surface-container-high);
  border-radius: var(--radius-sm);
  padding: 8px 12px;
  display: flex;
  flex-direction: column;
  gap: 6px;
  font-size: 11px;
}

.detail-row {
  display: flex;
  justify-content: space-between;
}

.detail-key {
  color: var(--color-on-surface-variant);
}

.detail-val {
  color: var(--color-on-surface);
  font-weight: 500;
}

.about-footer {
  padding: 10px 16px;
  border-block-start: 1px solid var(--color-outline-variant);
  background: var(--color-surface-container-high);
  display: flex;
  justify-content: flex-end;
}

/* הטקסט על הכפתור הממולא הוא `--color-on-primary` ולא לבן קבוע: זה בדיוק
   התפקיד שלו ב-M3 („טקסט/אייקון בתוך אלמנטים בצבע primary”), והוא מגיע
   מהערכה. `#ffffff` שהיה כאן נמדד במצב כהה כלבן על כחול-בהיר — כמעט בלתי
   קריא, כי שם ה-primary עצמו בהיר וה-onPrimary שלו כהה. */
.about-btn {
  padding: 4px 16px;
  font-size: 12px;
  font-family: var(--font-main);
  background: var(--word-blue);
  color: var(--color-on-primary);
  border: none;
  border-radius: var(--radius-xs);
  cursor: pointer;
}

.about-btn:hover {
  background: var(--word-blue-dark);
}
</style>
