<template>
  <!-- Teleport מאותו טעם כמו בשאר הדיאלוגים. ראו BookmarkDialog.vue. -->
  <Teleport to="body">
    <div
      v-if="isOpen"
      class="macros-dialog"
      :style="dragStyle"
      role="dialog"
      aria-modal="true"
      :aria-label="DIALOG_TITLE"
      @keydown.esc.stop="$emit('close')"
    >
      <div
        class="md-header dialog-drag-handle"
        @pointerdown="startDialogDrag"
      >
        <span class="md-title">{{ DIALOG_TITLE }}</span>
        <button
          type="button"
          class="md-close-btn"
          data-tip-title="סגור"
          data-tip-shortcut="Esc"
          aria-label="סגור את ניהול המאקרו"
          @click="$emit('close')"
        >
          ✕
        </button>
      </div>

      <div
        v-if="!handle"
        class="md-body"
      >
        <p
          class="md-note"
          role="note"
        >
          יש לפתוח מסמך כדי לנהל מאקרו
        </p>
      </div>

      <template v-else>
        <div
          class="md-tabs"
          role="tablist"
          aria-label="סוגי המאקרו"
        >
          <button
            v-for="tab in TABS"
            :key="tab.id"
            type="button"
            class="md-tab"
            :class="{ 'md-tab--active': section === tab.id }"
            role="tab"
            :aria-selected="section === tab.id"
            @pointerdown.prevent
            @click="switchSection(tab.id)"
          >
            {{ tab.title }}
          </button>
        </div>

        <!-- ═══ הקלטות ═══ -->
        <div
          v-if="section === 'recordings'"
          class="md-body"
        >
          <div
            v-if="recordings.length > 0"
            class="md-list"
            role="listbox"
            aria-label="המאקרו המוקלטים"
          >
            <button
              v-for="item in recordings"
              :key="item.id"
              type="button"
              class="md-list-item"
              :class="{ 'md-list-item--selected': item.id === selectedId }"
              role="option"
              :aria-selected="item.id === selectedId"
              @pointerdown.prevent
              @click="selectRecording(item.id)"
            >
              <span class="md-item-name">{{ item.name }}</span>
              <span class="md-item-meta">{{ recordingMeta(item) }}</span>
            </button>
          </div>
          <p
            v-else
            class="md-note"
            role="note"
          >
            אין עדיין הקלטות. Ctrl+Alt+R מתחיל הקלטה: עובדים רגיל במסמך —
            הקלדה, עיצוב, רשימות — ואותו צירוף עוצר ושומר.
          </p>

          <template v-if="selectedId && section === 'recordings'">
            <div class="md-input-row">
              <label
                for="md-rec-name"
                class="md-label"
              >שם:</label>
              <input
                id="md-rec-name"
                v-model="itemName"
                type="text"
                class="md-input"
              >
            </div>
            <div class="md-input-row">
              <label
                for="md-rec-shortcut"
                class="md-label"
              >קיצור מקלדת:</label>
              <input
                id="md-rec-shortcut"
                v-model="itemShortcut"
                type="text"
                class="md-input"
                placeholder="למשל: Ctrl+Alt+1 (רשות)"
                :aria-invalid="shortcutInvalid"
              >
            </div>
            <p
              v-if="shortcutInvalid"
              class="md-error"
              role="alert"
            >
              {{ shortcutError }}
            </p>
          </template>
        </div>

        <!-- ═══ קטעי טקסט ═══ -->
        <div
          v-else-if="section === 'snippets'"
          class="md-body"
        >
          <div
            v-if="snippets.length > 0"
            class="md-list"
            role="listbox"
            aria-label="קטעי הטקסט"
          >
            <button
              v-for="item in snippets"
              :key="item.id"
              type="button"
              class="md-list-item"
              :class="{ 'md-list-item--selected': item.id === selectedId }"
              role="option"
              :aria-selected="item.id === selectedId"
              @pointerdown.prevent
              @click="selectSnippet(item.id)"
            >
              <span class="md-item-name">{{ item.name }}</span>
              <span class="md-item-meta">{{ snippetMeta(item) }}</span>
            </button>
          </div>

          <div class="md-input-row">
            <label
              for="md-snip-name"
              class="md-label"
            >שם:</label>
            <input
              id="md-snip-name"
              v-model="itemName"
              type="text"
              class="md-input"
              placeholder="למשל: חתימה"
            >
          </div>
          <div class="md-input-row md-input-row--top">
            <label
              for="md-snip-text"
              class="md-label"
            >תוכן:</label>
            <textarea
              id="md-snip-text"
              v-model="snippetText"
              class="md-textarea"
              rows="3"
              :placeholder="SNIPPET_TEXT_PLACEHOLDER"
            />
          </div>
          <div class="md-input-row">
            <label
              for="md-snip-trigger"
              class="md-label"
            >מילת הפעלה:</label>
            <input
              id="md-snip-trigger"
              v-model="snippetTrigger"
              type="text"
              class="md-input"
              placeholder="הקלדת המילה + רווח מרחיבה (רשות)"
            >
          </div>
          <div class="md-input-row">
            <label
              for="md-snip-shortcut"
              class="md-label"
            >קיצור מקלדת:</label>
            <input
              id="md-snip-shortcut"
              v-model="itemShortcut"
              type="text"
              class="md-input"
              placeholder="למשל: Ctrl+Alt+2 (רשות)"
              :aria-invalid="shortcutInvalid"
            >
          </div>
          <p
            v-if="shortcutInvalid"
            class="md-error"
            role="alert"
          >
            {{ shortcutError }}
          </p>
          <p
            class="md-note"
            role="note"
          >
            משתנים בתוכן: <bdi dir="ltr">{{ '\{\{date\}\}' }}</bdi> תאריך,
            <bdi dir="ltr">{{ '\{\{time\}\}' }}</bdi> שעה,
            <bdi dir="ltr">{{ '\{\{selection\}\}' }}</bdi> הטקסט המסומן.
          </p>
        </div>

        <!-- ═══ סקריפטים ═══ -->
        <div
          v-else-if="section === 'scripts'"
          class="md-body"
        >
          <div
            v-if="scripts.length > 0"
            class="md-list"
            role="listbox"
            aria-label="המאקרו הכתובים"
          >
            <button
              v-for="item in scripts"
              :key="item.id"
              type="button"
              class="md-list-item"
              :class="{ 'md-list-item--selected': item.id === selectedId }"
              role="option"
              :aria-selected="item.id === selectedId"
              @pointerdown.prevent
              @click="selectScript(item.id)"
            >
              <span class="md-item-name">{{ item.name }}</span>
              <span
                v-if="item.shortcut"
                class="md-item-meta"
              >{{ item.shortcut }}</span>
            </button>
          </div>

          <div class="md-input-row">
            <label
              for="md-scr-name"
              class="md-label"
            >שם:</label>
            <input
              id="md-scr-name"
              v-model="itemName"
              type="text"
              class="md-input"
              placeholder="למשל: כותרת דבר תורה"
            >
          </div>
          <div class="md-input-row md-input-row--top">
            <label
              for="md-scr-source"
              class="md-label"
            >קוד:</label>
            <textarea
              id="md-scr-source"
              v-model="scriptSource"
              class="md-textarea md-textarea--code"
              rows="6"
              dir="ltr"
              spellcheck="false"
              :placeholder="SCRIPT_PLACEHOLDER"
            />
          </div>
          <div class="md-input-row">
            <label
              for="md-scr-shortcut"
              class="md-label"
            >קיצור מקלדת:</label>
            <input
              id="md-scr-shortcut"
              v-model="itemShortcut"
              type="text"
              class="md-input"
              placeholder="למשל: Ctrl+Alt+3 (רשות)"
              :aria-invalid="shortcutInvalid"
            >
          </div>
          <p
            v-if="shortcutInvalid"
            class="md-error"
            role="alert"
          >
            {{ shortcutError }}
          </p>
          <p
            v-if="scriptResult"
            class="md-note"
            :class="{ 'md-error': scriptResultIsError }"
            :role="scriptResultIsError ? 'alert' : 'note'"
          >
            {{ scriptResult }}
          </p>
          <p
            class="md-note"
            role="note"
          >
            הקוד רץ בסביבה מבודדת מול <bdi dir="ltr">api</bdi> של המסמך —
            למשל <bdi dir="ltr">await api.insertText('…')</bdi>,
            <bdi dir="ltr">api.bold()</bdi>,
            <bdi dir="ltr">api.replaceAll(a, b)</bdi>.
          </p>
        </div>

        <!-- ═══ מאקרו VBA שבמסמך ═══ -->
        <div
          v-else-if="section === 'vba'"
          class="md-body"
        >
          <p
            class="md-note md-note--warn"
            role="note"
          >
            אלה מאקרו ה-VBA שכבר נמצאים במסמך. הם <strong>אינם מורצים כאן</strong>
            — אין מנוע VBA בדפדפן — ונשמרים בקובץ כמות שהם. הקוד מוצג לעיון, כדי
            שאפשר יהיה להעביר אותו למאקרו של העורך.
          </p>

          <p
            v-for="(warning, index) in documentVba.warnings"
            :key="index"
            class="md-note md-note--warn"
            role="alert"
          >
            {{ warning }}
          </p>

          <div
            v-if="vbaModules.length > 0"
            class="md-input-row"
          >
            <label
              for="md-vba-module"
              class="md-label"
            >מודול:</label>
            <select
              id="md-vba-module"
              v-model="vbaModuleName"
              class="md-input"
            >
              <option
                v-for="module in vbaModules"
                :key="module.name"
                :value="module.name"
              >
                {{ module.name }} — {{ moduleKindLabel(module.kind) }}
              </option>
            </select>
          </div>

          <div
            v-if="selectedVbaModule"
            class="md-input-row md-input-row--top"
          >
            <label
              for="md-vba-source"
              class="md-label"
            >קוד:</label>
            <textarea
              id="md-vba-source"
              class="md-textarea md-textarea--code"
              rows="10"
              dir="ltr"
              readonly
              spellcheck="false"
              :value="selectedVbaModule.source"
            />
          </div>

          <p
            v-else-if="documentVba.hasMacroPart"
            class="md-note"
            role="note"
          >
            לא נקרא אף מודול מפרויקט המאקרו של המסמך.
          </p>
          <p
            v-else
            class="md-note"
            role="note"
          >
            במסמך הזה אין מאקרו VBA.
          </p>
        </div>

        <!-- ═══ כלים מובנים ═══ -->
        <div
          v-else-if="section === 'tools'"
          class="md-body"
        >
          <div
            v-if="tools.length > 0"
            class="md-list"
            role="listbox"
            aria-label="הכלים המובנים"
          >
            <button
              v-for="item in tools"
              :key="item.id"
              type="button"
              class="md-list-item"
              :class="{ 'md-list-item--selected': item.id === selectedId }"
              role="option"
              :aria-selected="item.id === selectedId"
              @pointerdown.prevent
              @click="selectTool(item.id)"
            >
              <span class="md-item-name">{{ item.name }}</span>
              <span class="md-item-meta">{{ item.shortcut ?? '' }}</span>
            </button>
          </div>
          <p
            v-else
            class="md-note"
            role="note"
          >
            אין כלים מובנים במסמך הזה.
          </p>

          <template v-if="selectedId && section === 'tools'">
            <p
              v-if="selectedToolDescription"
              class="md-note"
              role="note"
            >
              {{ selectedToolDescription }}
            </p>
            <div class="md-input-row">
              <label
                for="md-tool-shortcut"
                class="md-label"
              >קיצור מקלדת:</label>
              <input
                id="md-tool-shortcut"
                v-model="itemShortcut"
                type="text"
                class="md-input"
                placeholder="למשל: Ctrl+Alt+3 (רשות)"
                :aria-invalid="shortcutInvalid"
              >
            </div>
            <p
              v-if="shortcutInvalid"
              class="md-error"
              role="alert"
            >
              {{ shortcutError }}
            </p>
          </template>
        </div>

        <!-- ═══ ייבוא/ייצוא ═══ -->
        <div
          v-else
          class="md-body"
        >
          <div class="md-input-row md-input-row--top">
            <label
              for="md-transfer"
              class="md-label"
            >JSON:</label>
            <textarea
              id="md-transfer"
              v-model="transferText"
              class="md-textarea md-textarea--code"
              rows="8"
              dir="ltr"
              spellcheck="false"
              placeholder="„ייצא לכאן” ממלא את התיבה; הדבקה + „ייבא מכאן” ממזגת"
            />
          </div>
          <p
            class="md-note"
            role="note"
          >
            הייצוא כולל את כל ההקלטות, הקטעים והסקריפטים — להעתקה, לגיבוי או
            לשיתוף. הייבוא ממזג: פריט עם אותו מזהה מוחלף, השאר מתווספים.
          </p>
        </div>

        <div class="md-footer">
          <template v-if="section === 'recordings'">
            <button
              type="button"
              class="md-btn md-btn-primary"
              :disabled="busy || !selectedId"
              @pointerdown.prevent
              @click="onReplay"
            >
              נגן
            </button>
            <button
              type="button"
              class="md-btn"
              :disabled="!selectedId || !canSaveItem"
              @pointerdown.prevent
              @click="onUpdateRecording"
            >
              עדכן
            </button>
          </template>

          <template v-else-if="section === 'snippets'">
            <button
              type="button"
              class="md-btn md-btn-primary"
              :disabled="!canSaveSnippet"
              @pointerdown.prevent
              @click="onSaveSnippet"
            >
              {{ selectedId ? 'עדכן' : 'הוסף' }}
            </button>
            <button
              type="button"
              class="md-btn"
              :disabled="busy || !selectedId"
              @pointerdown.prevent
              @click="onInsertSnippet"
            >
              הכנס למסמך
            </button>
          </template>

          <template v-else-if="section === 'scripts'">
            <button
              type="button"
              class="md-btn md-btn-primary"
              :disabled="busy || !canRunScript"
              @pointerdown.prevent
              @click="onRunScript"
            >
              הרץ
            </button>
            <button
              type="button"
              class="md-btn"
              :disabled="!canSaveScript"
              @pointerdown.prevent
              @click="onSaveScript"
            >
              {{ selectedId ? 'עדכן' : 'הוסף' }}
            </button>
          </template>

          <!--
            „VBA במסמך” אינה מוסיפה דבר ל-footer — „סגור” בלבד. במפורש ולא
            דרך ה-`v-else`: בלי הענף הזה הלשונית הייתה מקבלת את כפתורי
            הייבוא/ייצוא של „ייבוא וייצוא”, כלומר כפתור „ייבא” על לשונית
            שכל עניינה הוא שהקוד הזה אינו נכנס לשום מקום.
          -->
          <template v-else-if="section === 'tools'">
            <button
              type="button"
              class="md-btn md-btn-primary"
              :disabled="busy || !selectedId"
              @pointerdown.prevent
              @click="onRunTool"
            >
              הרץ
            </button>
            <button
              type="button"
              class="md-btn"
              :disabled="!selectedId || shortcutInvalid"
              @pointerdown.prevent
              @click="onSaveToolShortcut"
            >
              שמור קיצור
            </button>
          </template>

          <template v-else-if="section === 'vba'" />

          <template v-else>
            <button
              type="button"
              class="md-btn md-btn-primary"
              @pointerdown.prevent
              @click="onExport"
            >
              ייצא לכאן
            </button>
            <button
              type="button"
              class="md-btn"
              :disabled="transferText.trim() === ''"
              @pointerdown.prevent
              @click="onImport"
            >
              ייבא מכאן
            </button>
          </template>

          <!--
            `managesItems` ולא רשימת `!==`: כל לשונית חדשה שאינה מנהלת פריטים
            של ה-kit הייתה מקבלת „מחק” בשקט. על „VBA במסמך” הכפתור הזה גם היה
            שקר — אין שם מה למחוק, המאקרו שייכים לקובץ.
          -->
          <button
            v-if="managesItems"
            type="button"
            class="md-btn"
            :disabled="!selectedId"
            @pointerdown.prevent
            @click="onRemove"
          >
            מחק
          </button>
          <button
            v-if="managesItems && section !== 'recordings'"
            type="button"
            class="md-btn"
            @pointerdown.prevent
            @click="clearForm"
          >
            חדש
          </button>
          <button
            type="button"
            class="md-btn"
            @pointerdown.prevent
            @click="$emit('close')"
          >
            סגור
          </button>
        </div>
      </template>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
/**
 * דיאלוג ניהול המאקרו — המקבילה של Alt+F8 ב-Word, בתוספת קטעי הטקסט
 * (Building Blocks) והסקריפטים.
 *
 * הדיאלוג עובד ישירות מול `MacroKit` של ה-session, וזו סטייה מכוונת מתבנית
 * „מציג בלבד” של BookmarkDialog: הכלל שם נועד למנוע מסלול שני אל **המנוע**
 * — קריאה שמחזירה קבלות באנגלית ודורשת תרגום אחיד. ה-kit אינו המנוע: הוא
 * שכבת האדפטר של המאקרו (engine/macros.ts), כל התוצאות שלו כבר בעברית, והוא
 * המקור היחיד לרשימות — אין כאן שני קולות לאותו מידע. הרשימות מרועננות
 * מה-kit אחרי כל פעולה, כמו שהלשוניות קוראות מחדש אחרי פעולת מנוע.
 *
 * הדיאלוג **נשאר פתוח** אחרי כל פעולה, כמו דיאלוג הסימניות ומאותו טעם:
 * ניהול הוא רצף — מוסיפים קטע, מוחקים שניים, מריצים סקריפט — ולכן הכפתור
 * האחרון הוא „סגור” ולא „ביטול”.
 *
 * ולידציית הקיצורים היא `parseShortcut` של החבילה — אותה פונקציה שתקרא את
 * הקיצור בזמן הקשה. שני נוסחים לאותה שאלה היו מאפשרים לשמור קיצור שלעולם
 * לא יופעל (הלקח של BookmarkDialog מול `normalizeBookmarkName`).
 */
import { computed, ref, shallowRef, watch } from 'vue';
import type { BuiltinToolInfo, RecordedMacro, SavedScript, Snippet, VbaModule } from 'superdoc-macros';
import type { MacrosHandle } from '../../engine/macros';
import { MODULE_KIND_LABEL, NO_VBA, type DocumentVba } from '../../engine/vba-import';
import { useDialogDrag } from '../../composables/dialog-drag';

/* הדיאלוג נגרר בכותרת שלו — composables/dialog-drag.ts. */
const { dragStyle, startDialogDrag } = useDialogDrag();

const props = withDefaults(
  defineProps<{
    isOpen?: boolean;
    /** Disable document execution while preserving edits to the library form. */
    busy?: boolean;
    /** Read the active editor lock at click time, before a render can catch up. */
    isDocumentBlocked?: () => boolean;
    /** מערכת המאקרו של המסמך הפתוח, או `null` כשאין מסמך. */
    handle?: MacrosHandle | null;
    /**
     * מאקרו ה-VBA שכבר במסמך — לתצוגה בלבד. תכונה של המסמך ולא של ה-kit,
     * ולכן prop נפרד: אין שום מסלול שמריץ אותם או שומר אותם כפריט.
     */
    documentVba?: DocumentVba;
  }>(),
  { isOpen: false, busy: false, isDocumentBlocked: () => false, handle: null, documentVba: () => NO_VBA }
);

const emit = defineEmits<{
  (e: 'close'): void;
  /** דיווח לשורת המצב — הדיאלוג אינו מציג סטטוס גלובלי בעצמו. */
  (e: 'status', message: string, isError?: boolean): void;
}>();

const DIALOG_TITLE = 'ניהול מאקרו';
const SNIPPET_TEXT_PLACEHOLDER = 'הטקסט שיוכנס. אפשר לשלב {{date}}, {{time}}, {{selection}}';
const SCRIPT_PLACEHOLDER = `await api.insertText('...');`;

type Section = 'recordings' | 'snippets' | 'scripts' | 'tools' | 'vba' | 'transfer';

/**
 * לשונית הסקריפטים מוצגת רק כשהדגל דלוק (ראו SCRIPTS_FLAG_KEY ב-engine/macros.ts):
 * מאקרו כתובים נשארים feature flag עד שההקשחה תוכרע. הקלטות וקטעים — תמיד.
 *
 * לשונית „VBA במסמך” מוצגת רק כשיש מה להציג בה: היא מתארת את הקובץ שנפתח, ולא
 * יכולת של העורך, ולשונית ריקה על מסמך רגיל הייתה רק מבלבלת.
 */
const TABS = computed<ReadonlyArray<{ id: Section; title: string }>>(() => [
  { id: 'recordings', title: 'הקלטות' },
  { id: 'snippets', title: 'קטעי טקסט' },
  ...(props.handle?.scriptsEnabled ? [{ id: 'scripts' as const, title: 'סקריפטים' }] : []),
  // „כלים” מוצגת רק כשהמעטפת רשמה כלים מובנים — kit בלי כלים הוא לשונית ריקה.
  ...(tools.value.length > 0 ? [{ id: 'tools' as const, title: 'כלים' }] : []),
  ...(props.documentVba.hasMacroPart ? [{ id: 'vba' as const, title: 'VBA במסמך' }] : []),
  { id: 'transfer', title: 'ייבוא וייצוא' },
]);

/* ---------- VBA שבמסמך (קריאה בלבד) ---------- */

const vbaModules = computed<readonly VbaModule[]>(() => props.documentVba.modules);

/** המודול המוצג. `''` כשאין מודולים. */
const vbaModuleName = ref('');

function moduleKindLabel(kind: VbaModule['kind']): string {
  return MODULE_KIND_LABEL[kind];
}

const selectedVbaModule = computed<VbaModule | null>(
  () => vbaModules.value.find((module) => module.name === vbaModuleName.value) ?? null
);

/* מסמך אחר — מודולים אחרים. בלי האיפוס הבחירה הייתה מצביעה על מודול של המסמך
   הקודם, והתיבה הייתה מוצגת ריקה בלי הסבר. */
watch(
  vbaModules,
  (modules) => {
    if (!modules.some((module) => module.name === vbaModuleName.value)) {
      vbaModuleName.value = modules[0]?.name ?? '';
    }
  },
  { immediate: true }
);

const section = ref<Section>('recordings');

/**
 * האם הלשונית הנוכחית מנהלת פריטים של ה-kit — כלומר האם „מחק”/„חדש” נכונים
 * בה. „ייבוא וייצוא” ו„VBA במסמך” אינן: הראשונה עובדת על המצב כולו, והשנייה
 * מציגה את הקובץ ואין בה מה למחוק.
 */
const managesItems = computed(
  () => section.value === 'recordings' || section.value === 'snippets' || section.value === 'scripts'
);

/** הרשימות — תצלום מה-kit, מרוענן אחרי כל פעולה. ראו הערת הפתיחה. */
const recordings = shallowRef<readonly RecordedMacro[]>([]);
const snippets = shallowRef<readonly Snippet[]>([]);
const scripts = shallowRef<readonly SavedScript[]>([]);
/**
 * הכלים המובנים שהמעטפת רשמה על ה-kit (registerShulchanTools). `TABS` קורא
 * אותם אף שהוא מוגדר מעל: הוא `computed`, כלומר גופו רץ בקריאה הראשונה —
 * הרבה אחרי שכל ההשמות של ה-setup הסתיימו.
 */
const tools = shallowRef<readonly BuiltinToolInfo[]>([]);

/** הפריט הנבחר בלשונית הנוכחית, או `''`. */
const selectedId = ref('');

/* שדות הטופס. `itemName`/`itemShortcut` משותפים לשלוש הלשוניות — בכל רגע
   רק טופס אחד מוצג, והחלפת לשונית מנקה אותם (ראו switchSection). */
const itemName = ref('');
const itemShortcut = ref('');
const snippetText = ref('');
const snippetTrigger = ref('');
const scriptSource = ref('');
const scriptResult = ref('');
const scriptResultIsError = ref(false);
const transferText = ref('');

function refresh(): void {
  const kit = props.handle?.kit;
  recordings.value = kit ? [...kit.listRecordings()] : [];
  snippets.value = kit ? [...kit.listSnippets()] : [];
  scripts.value = kit ? [...kit.listScripts()] : [];
  tools.value = kit && typeof kit.listTools === 'function' ? [...kit.listTools()] : [];
}

watch(
  () => props.isOpen,
  (open) => {
    if (!open) return;
    // נקי בכל פתיחה, כמו BookmarkDialog: בחירה מפתיחה קודמת עשויה להצביע
    // על פריט של מסמך אחר או על פריט שנמחק.
    refresh();
    clearForm();
    scriptResult.value = '';
    // לשונית שהוסתרה (דגל הסקריפטים כבה בין הפעלות) אינה יכולה להישאר פעילה.
    if (!TABS.value.some((tab) => tab.id === section.value)) section.value = 'recordings';
  },
  // `immediate`: הדיאלוג יכול להיוולד כבר פתוח (בדיקות, שחזור מצב), ואז
  // אין מעבר false→true שירענן את הרשימות.
  { immediate: true }
);

function switchSection(next: Section): void {
  section.value = next;
  clearForm();
}

function clearForm(): void {
  selectedId.value = '';
  itemName.value = '';
  itemShortcut.value = '';
  snippetText.value = '';
  snippetTrigger.value = '';
  scriptSource.value = '';
}

/* ---------- ולידציה ---------- */

/**
 * הוולידציה היא `kit.validateShortcut` — אותה בדיקה שהשמירה אוכפת: קיצור
 * חייב modifier אמיתי (אות בודדת הייתה יורה על כל הקלדה), אסור לו להתנגש
 * בקיצור של העורך (רשימת הרג'יסטרי נמסרת ל-kit בהתקנה) או בפריט שמור אחר.
 * שני נוסחים לאותה שאלה היו מאפשרים לאשר קיצור שהשמירה תדחה.
 */
const shortcutValidation = computed(
  () => props.handle?.kit.validateShortcut(itemShortcut.value, selectedId.value || undefined) ?? { ok: true as const }
);
const shortcutInvalid = computed(() => !shortcutValidation.value.ok);
const shortcutError = computed(() =>
  shortcutValidation.value.ok ? '' : shortcutValidation.value.message
);

const canSaveItem = computed(() => itemName.value.trim() !== '' && !shortcutInvalid.value);
const canSaveSnippet = computed(() => canSaveItem.value && snippetText.value !== '');
const canSaveScript = computed(() => canSaveItem.value && scriptSource.value.trim() !== '');
const canRunScript = computed(() => scriptSource.value.trim() !== '');

/* ---------- בחירה ---------- */

function selectRecording(id: string): void {
  const item = recordings.value.find((entry) => entry.id === id);
  if (!item) return;
  selectedId.value = id;
  itemName.value = item.name;
  itemShortcut.value = item.shortcut ?? '';
}

function selectSnippet(id: string): void {
  const item = snippets.value.find((entry) => entry.id === id);
  if (!item) return;
  selectedId.value = id;
  itemName.value = item.name;
  itemShortcut.value = item.shortcut ?? '';
  snippetText.value = item.text;
  snippetTrigger.value = item.trigger ?? '';
}

function selectTool(id: string): void {
  const item = tools.value.find((entry) => entry.id === id);
  if (!item) return;
  selectedId.value = id;
  itemName.value = item.name;
  itemShortcut.value = item.shortcut ?? '';
}

const selectedToolDescription = computed(
  () => tools.value.find((entry) => entry.id === selectedId.value)?.description ?? ''
);

function selectScript(id: string): void {
  const item = scripts.value.find((entry) => entry.id === id);
  if (!item) return;
  selectedId.value = id;
  itemName.value = item.name;
  itemShortcut.value = item.shortcut ?? '';
  scriptSource.value = item.source;
  scriptResult.value = '';
}

/* ---------- תוויות ---------- */

function recordingMeta(item: RecordedMacro): string {
  const steps = `${item.steps.length} צעדים`;
  return item.shortcut ? `${steps} · ${item.shortcut}` : steps;
}

function snippetMeta(item: Snippet): string {
  const parts: string[] = [];
  if (item.trigger) parts.push(item.trigger);
  if (item.shortcut) parts.push(item.shortcut);
  return parts.join(' · ');
}

/* ---------- פעולות ---------- */

function onReplay(): void {
  if (props.busy || props.isDocumentBlocked()) return;
  const kit = props.handle?.kit;
  if (!kit || !selectedId.value) return;
  void kit.replayRecording(selectedId.value).then((result) => {
    if (result.ok) emit('status', 'המאקרו נוגן');
    else emit('status', `ניגון המאקרו נכשל: ${result.failures[0]?.message ?? 'כשל לא ידוע'}`, true);
  });
}

/**
 * השמירות עטופות: ה-kit אוכף את ולידציית הקיצור בזריקה, והדיאלוג אמנם בודק
 * את אותה שאלה לפני שהכפתור נדלק — אבל בין הבדיקה ללחיצה פריט אחר יכול היה
 * לתפוס את הקיצור (למשל עצירת הקלטה מהמקלדת). הכשל מגיע לשורת המצב במקום
 * להפיל את הדיאלוג.
 */
function guardedSave(action: () => void): void {
  try {
    action();
  } catch (error) {
    emit('status', error instanceof Error ? error.message : 'השמירה נכשלה', true);
  }
}

function onUpdateRecording(): void {
  const kit = props.handle?.kit;
  if (!kit || !selectedId.value || !canSaveItem.value) return;
  guardedSave(() => {
    kit.updateRecording({
      id: selectedId.value,
      name: itemName.value.trim(),
      shortcut: itemShortcut.value.trim(),
    });
  });
  refresh();
}

function onSaveSnippet(): void {
  const kit = props.handle?.kit;
  if (!kit || !canSaveSnippet.value) return;
  guardedSave(() => {
    const saved = kit.saveSnippet({
      ...(selectedId.value ? { id: selectedId.value } : {}),
      name: itemName.value.trim(),
      text: snippetText.value,
      trigger: snippetTrigger.value.trim() || undefined,
      shortcut: itemShortcut.value.trim() || undefined,
    });
    selectedId.value = saved.id;
  });
  refresh();
}

function onInsertSnippet(): void {
  if (props.busy || props.isDocumentBlocked()) return;
  const kit = props.handle?.kit;
  if (!kit || !selectedId.value) return;
  void kit.expandSnippet(selectedId.value).then((result) => {
    if (!result.ok) emit('status', result.message ?? 'הכנסת הקטע נכשלה', true);
  });
}

function onSaveScript(): void {
  const kit = props.handle?.kit;
  if (!kit || !canSaveScript.value) return;
  guardedSave(() => {
    const saved = kit.saveScript({
      ...(selectedId.value ? { id: selectedId.value } : {}),
      name: itemName.value.trim(),
      source: scriptSource.value,
      shortcut: itemShortcut.value.trim() || undefined,
    });
    selectedId.value = saved.id;
  });
  refresh();
}

/**
 * מריץ את מה שבעורך — לא את הגרסה השמורה: כך אפשר לנסות שינוי לפני „עדכן”,
 * שזה בדיוק מחזור העבודה של כתיבת מאקרו.
 */
function onRunScript(): void {
  if (props.busy || props.isDocumentBlocked()) return;
  const kit = props.handle?.kit;
  if (!kit || !canRunScript.value) return;
  scriptResult.value = 'מריץ…';
  scriptResultIsError.value = false;
  void kit.runSource(scriptSource.value).then((result) => {
    scriptResultIsError.value = !result.ok;
    scriptResult.value = result.ok ? 'המאקרו הסתיים בהצלחה' : result.message;
  });
}

function onRunTool(): void {
  if (props.busy || props.isDocumentBlocked()) return;
  const kit = props.handle?.kit;
  if (!kit || !selectedId.value) return;
  void kit.runTool(selectedId.value).then((outcome) => {
    // סיכום הצלחה מגיע לשורת המצב מהכלי עצמו (onSummary ברישום); כאן רק כשל.
    if (!outcome.ok) emit('status', outcome.message, true);
  });
}

function onSaveToolShortcut(): void {
  const kit = props.handle?.kit;
  if (!kit || !selectedId.value || shortcutInvalid.value) return;
  guardedSave(() => {
    kit.setToolShortcut(selectedId.value, itemShortcut.value.trim() || undefined);
  });
  refresh();
}

function onRemove(): void {
  const kit = props.handle?.kit;
  if (!kit || !selectedId.value) return;
  // guardedSave גם כאן: מחיקה היא commit לאחסון, והוא יכול לזרוק (quota).
  guardedSave(() => {
    if (section.value === 'recordings') kit.removeRecording(selectedId.value);
    else if (section.value === 'snippets') kit.removeSnippet(selectedId.value);
    else kit.removeScript(selectedId.value);
  });
  refresh();
  clearForm();
}

function onExport(): void {
  const kit = props.handle?.kit;
  if (!kit) return;
  transferText.value = kit.exportState();
}

function onImport(): void {
  const kit = props.handle?.kit;
  if (!kit) return;
  const outcome = kit.importState(transferText.value, { merge: true });
  if (outcome.ok) {
    refresh();
    emit('status', 'המאקרו יובאו ומוזגו');
  } else {
    emit('status', outcome.message ?? 'הייבוא נכשל', true);
  }
}
</script>

<style scoped>
.macros-dialog {
  position: fixed;
  top: 120px;
  inset-inline-start: 40px;
  z-index: 2000;
  background: var(--color-surface);
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.16);
  width: 460px;
  font-family: var(--font-main);
  user-select: none;
}

.md-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  border-block-end: 1px solid var(--color-outline-variant);
  padding: 6px 8px 6px 12px;
  background: var(--color-surface-container-high);
  border-radius: var(--radius-sm) var(--radius-sm) 0 0;
}

.md-title {
  font-size: 12px;
  font-weight: 600;
  color: var(--color-on-surface);
}

.md-close-btn {
  background: none;
  border: none;
  font-size: 14px;
  color: var(--color-on-surface-variant);
  cursor: pointer;
  padding: 4px 6px;
  border-radius: var(--radius-xs);
}

.md-close-btn:hover {
  background: var(--word-btn-hover);
  color: var(--color-error);
}

.md-tabs {
  display: flex;
  gap: 2px;
  padding: 6px 12px 0;
  border-block-end: 1px solid var(--color-outline-variant);
}

.md-tab {
  border: 0;
  background: transparent;
  color: var(--color-on-surface-variant);
  font-family: var(--font-main);
  font-size: 11px;
  padding: 5px 10px;
  cursor: pointer;
  border-radius: var(--radius-xs) var(--radius-xs) 0 0;
  border-block-end: 2px solid transparent;
}

.md-tab:hover {
  background: var(--word-btn-hover);
}

.md-tab--active {
  color: var(--color-on-surface);
  font-weight: 600;
  border-block-end-color: var(--word-blue);
}

.md-body {
  padding: 12px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.md-list {
  display: flex;
  flex-direction: column;
  gap: 1px;
  max-height: 130px;
  overflow-y: auto;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  padding: 2px;
  background: var(--color-surface);
}

.md-list-item {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 8px;
  border: 0;
  background: transparent;
  color: var(--color-on-surface);
  font: inherit;
  font-size: 12px;
  /* `start` ולא `right` — המעטפת RTL, וערך מוחלט היה שובר שם לועזי. */
  text-align: start;
  padding: 4px 8px;
  border-radius: var(--radius-xs);
  cursor: pointer;
}

.md-list-item:hover {
  background: var(--word-btn-hover);
}

.md-list-item--selected {
  background: var(--color-primary-subtle);
}

.md-item-name {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.md-item-meta {
  flex-shrink: 0;
  font-size: 10px;
  color: var(--color-on-surface-variant);
  direction: ltr;
}

.md-input-row {
  display: flex;
  align-items: center;
  gap: 8px;
}

.md-input-row--top {
  align-items: flex-start;
}

.md-label {
  width: 85px;
  font-size: 11px;
  color: var(--color-on-surface);
  flex-shrink: 0;
}

.md-input,
.md-textarea {
  flex: 1;
  min-width: 0;
  padding: 4px 8px;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: 12px;
  outline: none;
}

.md-textarea {
  resize: vertical;
  line-height: 1.5;
}

.md-textarea--code {
  font-family: Consolas, 'Courier New', monospace;
  font-size: 11px;
}

.md-input:focus,
.md-textarea:focus {
  border-color: var(--word-blue);
  box-shadow: 0 0 0 1px var(--word-blue);
}

.md-input[aria-invalid='true'] {
  border-color: var(--color-error);
}

.md-note {
  margin: 0;
  font-size: 11px;
  line-height: 1.4;
  color: var(--color-on-surface-variant);
}

/**
 * הערה שיש לקרוא, לא רק לראות: „הקוד הזה אינו רץ כאן” ואזהרות פרויקט המאקרו.
 * צבע הטקסט הרגיל ופס בקצה ההתחלה — לא אדום: אין כאן כשל ואין מה לתקן, יש
 * עובדה על הקובץ. אדום היה קורא למשתמש לחפש בעיה שאינה קיימת.
 */
.md-note--warn {
  padding-inline-start: 8px;
  border-inline-start: 2px solid var(--color-outline);
  color: var(--color-on-surface);
}

.md-error {
  margin: 0;
  font-size: 11px;
  line-height: 1.4;
  color: var(--color-error);
}

.md-footer {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 6px;
  padding: 8px 12px;
  border-block-start: 1px solid var(--color-outline-variant);
  background: var(--color-surface-container-high);
  border-radius: 0 0 var(--radius-sm) var(--radius-sm);
}

.md-btn {
  padding: 4px 10px;
  font-size: 11px;
  font-family: var(--font-main);
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  cursor: pointer;
  transition: all 0.08s;
}

.md-btn:hover:not(:disabled) {
  background: var(--word-btn-hover);
  border-color: var(--word-blue);
}

/* הטקסט על הכפתור הממולא הוא `--color-on-primary` — ראו LinkDialog.vue. */
.md-btn-primary {
  background: var(--word-blue);
  color: var(--color-on-primary);
  border-color: var(--word-blue);
}

.md-btn-primary:hover:not(:disabled) {
  background: var(--word-blue-dark);
}

.md-btn:disabled {
  opacity: 0.4;
  cursor: default;
}
</style>
