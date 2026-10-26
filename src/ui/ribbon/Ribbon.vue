<template>
  <div class="word-ribbon-container">
    <!-- סרגל הלשוניות -->
    <div class="word-tab-bar">
      <!-- ה-tablist עוטף את הלשוניות בלבד: הצאצאים של role="tablist" חייבים
           להיות role="tab", וכפתור הכיווץ אינו לשונית -->
      <div
        class="word-tab-strip"
        role="tablist"
        aria-orientation="horizontal"
        :aria-label="menuString('לשוניות הרצועה')"
        @keydown="onTabKeydown"
        @wheel="handleWheelScroll"
      >
        <button
          v-for="(tab, index) in translatedTabs"
          :id="ribbonTabId(tab.id)"
          :key="tab.id"
          :ref="(el) => registerTabRef(el, index)"
          type="button"
          role="tab"
          class="word-tab-btn"
          :class="[
            { active: currentTabId === tab.id },
            tab.className || ''
          ]"
          :aria-selected="currentTabId === tab.id ? 'true' : 'false'"
          :aria-controls="RIBBON_PANEL_ID"
          :tabindex="currentTabId === tab.id ? 0 : -1"
          @click="selectTab(tab.id)"
          @dblclick="toggleCollapsed"
        >
          {{ tab.label }}
        </button>
      </div>

      <button
        type="button"
        class="word-ribbon-toggle"
        :data-tip-title="isCollapsed ? menuString('הצג את הרצועה') : menuString('כווץ את הרצועה')"
        :aria-label="isCollapsed ? menuString('הצג את הרצועה') : menuString('כווץ את הרצועה')"
        :aria-expanded="!isCollapsed"
        :aria-controls="RIBBON_PANEL_ID"
        @click="toggleCollapsed"
      >
        <SvgIcon
          :name="isCollapsed ? 'chevronDown' : 'chevronUp'"
          :size="14"
        />
      </button>
    </div>

    <!-- תוכן הלשונית הפעילה בלבד (Mount on active). פאנל אחד שמתחלף ולא שמונה
         פאנלים, ולכן aria-labelledby מצביע על הלשונית הפעילה כרגע -->
    <div
      v-show="!isCollapsed"
      :id="RIBBON_PANEL_ID"
      ref="bodyRef"
      class="word-ribbon-body"
      role="tabpanel"
      :aria-labelledby="ribbonTabId(currentTabId)"
      @wheel="handleWheelScroll"
    >
      <FileTab
        v-if="currentTabId === 'file'"
        :has-document="hasDocument"
        :has-pdf-export="hasPdfExport"
        :is-saving="isSaving"
        :is-opening="isOpening"
        :is-exiting="isExiting"
        @new-doc="$emit('new-doc')"
        @open-doc="$emit('open-doc')"
        @save-doc="$emit('save-doc')"
        @save-as-doc="$emit('save-as-doc')"
        @print-doc="$emit('print-doc')"
        @export-pdf="$emit('export-pdf')"
        @about="$emit('about')"
        @report-bug="$emit('report-bug')"
        @shortcuts-help="$emit('shortcuts-help')"
        @shortcuts-manage="$emit('shortcuts-manage')"
        @exit-app="$emit('exit-app')"
      />
      <HomeTab
        v-else-if="currentTabId === 'home'"
        :list-autoformat-enabled="listAutoformatEnabled"
        @open-find="$emit('open-find')"
        @open-replace="$emit('open-replace')"
        @toggle-list-autoformat="$emit('toggle-list-autoformat')"
      />
      <InsertTab
        v-else-if="currentTabId === 'insert'"
        @open-link="$emit('open-link')"
      />
      <LayoutTab v-else-if="currentTabId === 'layout'" />
      <ReferencesTab v-else-if="currentTabId === 'references'" />
      <ReviewTab v-else-if="currentTabId === 'review'" />
      <ViewTab
        v-else-if="currentTabId === 'view'"
        :navigation-pane-visible="navigationPaneVisible"
        @toggle-focus-mode="$emit('toggle-focus-mode')"
        @toggle-navigation-pane="$emit('toggle-navigation-pane')"
      />
      <DeveloperTab
        v-else-if="currentTabId === 'developer'"
        @manage-macros="$emit('manage-macros')"
        @macro-record="$emit('macro-record')"
        @macro-play="$emit('macro-play')"
      />
      <ShulchanTab v-else-if="currentTabId === 'shulchan'" />
      <OtzariaTab
        v-else-if="currentTabId === 'otzaria'"
        :book-completion-enabled="bookCompletionEnabled"
        :dictation-state="dictationState"
        @insert-citation="$emit('insert-citation')"
        @search-otzaria="$emit('search-otzaria')"
        @open-library="$emit('open-library')"
        @export-otzaria="$emit('export-otzaria')"
        @toggle-book-completion="$emit('toggle-book-completion')"
        @toggle-dictation="$emit('toggle-dictation')"
      />
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, type ComponentPublicInstance } from 'vue';
import { RIBBON_PANEL_ID, nextTabIndex, ribbonTabId } from './aria';
import { handleWheelScroll } from '../../composables/wheel-scroll';
import { provideRibbonOverflow } from './overflow';
import { menuString } from './i18n';
import { RIBBON_TABS } from './tabs';
import SvgIcon from '../icons/SvgIcon.vue';
import HomeTab from './tabs/HomeTab.vue';
import FileTab from './tabs/FileTab.vue';
import InsertTab from './tabs/InsertTab.vue';
import LayoutTab from './tabs/LayoutTab.vue';
import ReferencesTab from './tabs/ReferencesTab.vue';
import ReviewTab from './tabs/ReviewTab.vue';
import ViewTab from './tabs/ViewTab.vue';
import DeveloperTab from './tabs/DeveloperTab.vue';
import ShulchanTab from './tabs/ShulchanTab.vue';
import OtzariaTab from './tabs/OtzariaTab.vue';

const TABS = RIBBON_TABS;

/**
 * הכיווץ של הקבוצות כשאין מקום, כמו ב-Word — במקום פס גלילה. הבקר יושב כאן
 * מפני שגוף הרצועה הוא המידה שכל הקבוצות מתחלקות בה (ui/ribbon/overflow.ts).
 */
const bodyRef = ref<HTMLElement | null>(null);
provideRibbonOverflow(bodyRef);

/**
 * התוויות שמוצגות למשתמש, מתורגמות לפי שפת המשתמש (ui/ribbon/i18n).
 * `TABS` עצמו נשאר בעברית והוא מקור האמת למזהים ולסדר; רק התווית מתורגמת.
 */
const translatedTabs = computed(() => TABS.map((tab) => ({ ...tab, label: menuString(tab.label) })));

/**
 * מצב המעטפת, לפקדי לשונית „קובץ”.
 *
 * הרצועה קיבלה עד עכשיו אפס props, וזה היה נכון: כל שאר הלשוניות שואבות את
 * המצב שלהן מהמנוע (`useCommand`, `ACTIVE_SUPERDOC`) ואינן צריכות דבר מהאב.
 * „קובץ” היא היחידה שפקדיה הם פעולות מעטפת — מסמך פתוח, שמירה שרצה, פתיחה
 * שרצה — ואת המצב הזה רק App.vue מחזיק. ההסבר המלא, כולל למה props ולא מפתח
 * הזרקה חדש, ב-FileTab.vue.
 *
 * הרצועה עצמה אינה קוראת אותם: היא צינור, בדיוק כמו שהיא צינור ל-events
 * בכיוון ההפוך.
 */
withDefaults(
  defineProps<{
    hasDocument?: boolean;
    /** האם ה-Host תומך ב-`ui.exportPdf` (אוצריא 0.9.97 ומעלה). */
    hasPdfExport?: boolean;
    isSaving?: boolean;
    isOpening?: boolean;
    isExiting?: boolean;
    bookCompletionEnabled?: boolean;
    dictationState?: 'idle' | 'starting' | 'listening';
    listAutoformatEnabled?: boolean;
    navigationPaneVisible?: boolean;
  }>(),
  {
    hasDocument: false,
    hasPdfExport: false,
    isSaving: false,
    isOpening: false,
    isExiting: false,
    bookCompletionEnabled: false,
    dictationState: 'idle',
    listAutoformatEnabled: true,
    navigationPaneVisible: false,
  },
);

/**
 * הלשונית הפעילה ומצב הכיווץ — `defineModel` ולא `ref` מקומי.
 *
 * למה השתנה: המשתמש שסגר את התוסף בלשונית „הפניות” וחזר אליו מצא את עצמו
 * ב„בית”. הבחירה הזאת שייכת למי שיושב מול המסך ולא למסמך, ולכן היא צריכה
 * לשרוד הפעלות — וזוכר-ההפעלה יושב ב-App.vue (sessions/session-keeper.ts).
 *
 * `defineModel` ולא props+emit ידניים, מפני שהוא זה שמשאיר את הרצועה עובדת
 * **גם בלי אב שקושר אליה**: בלי `v-model` הערך הוא ref מקומי רגיל, וזה בדיוק
 * מה שכל בדיקות הרצועה מרכיבות. אין כאן שני מצבי הפעלה לתחזק.
 */
const activeTabId = defineModel<string>('activeTab', { default: 'home' });
const isCollapsed = defineModel<boolean>('collapsed', { default: false });

/**
 * הלשונית שמוצגת בפועל.
 *
 * מזהה שאינו מוכר — רשומה ישנה מ-storage, או לשונית שהוסרה מהתוסף — היה
 * משאיר רצועה עם גוף ריק ובלי אף לשונית דלוקה. הנפילה ל„בית” היא מה שמונע
 * את זה, והיא נגזרת ואינה כתיבה חזרה: כתיבה כאן הייתה יוצרת לולאת עדכון עם
 * האב בדיוק ברגע העלייה.
 */
const currentTabId = computed(() =>
  TABS.some((tab) => tab.id === activeTabId.value) ? activeTabId.value : 'home',
);

defineEmits<{
  (e: 'new-doc'): void;
  (e: 'open-doc'): void;
  (e: 'save-doc'): void;
  (e: 'save-as-doc'): void;
  (e: 'print-doc'): void;
  (e: 'export-pdf'): void;
  (e: 'export-otzaria'): void;
  (e: 'about'): void;
  (e: 'report-bug'): void;
  (e: 'shortcuts-help'): void;
  (e: 'shortcuts-manage'): void;
  (e: 'exit-app'): void;
  (e: 'open-find'): void;
  (e: 'open-replace'): void;
  (e: 'open-link'): void;
  (e: 'toggle-focus-mode'): void;
  (e: 'insert-citation'): void;
  (e: 'search-otzaria'): void;
  (e: 'open-library'): void;
  (e: 'manage-macros'): void;
  (e: 'macro-record'): void;
  (e: 'macro-play'): void;
  (e: 'toggle-book-completion'): void;
  (e: 'toggle-dictation'): void;
  (e: 'toggle-list-autoformat'): void;
  (e: 'toggle-navigation-pane'): void;
}>();

/** רק הלשונית הפעילה נמצאת ב-tab order, ולכן החצים צריכים להזיז מיקוד בעצמם. */
const tabButtons = ref<Array<HTMLButtonElement | null>>([]);

function registerTabRef(el: Element | ComponentPublicInstance | null, index: number): void {
  tabButtons.value[index] = el instanceof HTMLButtonElement ? el : null;
}

function selectTab(id: string): void {
  activeTabId.value = id;
  if (isCollapsed.value) {
    isCollapsed.value = false;
  }
}

/** הפעלה אוטומטית: החץ מעביר מיקוד **ומחליף** לשונית, כמו ברצועה של Word. */
function onTabKeydown(event: KeyboardEvent): void {
  const current = TABS.findIndex((tab) => tab.id === currentTabId.value);
  // 'rtl' קבוע: המעטפת של התוסף היא dir="rtl" (index.html), והפונקציה תומכת
  // בשני הכיוונים כדי שאפשר יהיה למדוד את שניהם.
  const next = nextTabIndex(event.key, current, TABS.length, 'rtl');
  if (next === null) return;

  // בלי זה החצים גם גוללים את סרגל הלשוניות שגלילתו auto.
  event.preventDefault();
  selectTab(TABS[next].id);
  tabButtons.value[next]?.focus();
}

function toggleCollapsed(): void {
  isCollapsed.value = !isCollapsed.value;
}
</script>
