<template>
  <div
    ref="shellRef"
    class="word-app-shell"
    :class="[
      { 'focus-mode': isFocusMode },
      isFocusMode && revealed ? `reveal-${revealed}` : '',
    ]"
    @pointermove="onPointerMove"
    @pointerleave="revealed = null"
    @contextmenu="contextMenu.handleContextMenu"
  >
    <!-- פס עליון -->
    <TitleBar
      :title="title"
      :is-dirty="saveSnapshot.isDirty"
      :is-saving="saveSnapshot.isSaving"
      :is-save-error="saveSnapshot.state === 'error'"
      :save-state-text="saveStateMessage"
      :autosave-enabled="autosaveEnabled"
      :can-undo="canUndo"
      :can-redo="canRedo"
      @save="onSave(false)"
      @undo="onUndo"
      @redo="onRedo"
      @open-find="openFindDialog('find')"
      @toggle-autosave="toggleAutosave"
      @update-title="onTitleUpdate"
    />

    <!-- רצועת הכלים (Ribbon) -->
    <Ribbon
      v-model:active-tab="ribbonTab"
      v-model:collapsed="ribbonCollapsed"
      :has-document="hasDocument"
      :has-pdf-export="supportsPdfExport"
      :is-saving="saveSnapshot.isSaving"
      :is-opening="isOpening"
      :book-completion-enabled="bookCompletionEnabled"
      @new-doc="onNewDocument"
      @open-doc="onPickAndOpen"
      @save-doc="onSave(false)"
      @save-as-doc="onSave(true)"
      @export-doc="onExportDocx"
      @print-doc="onPrint"
      @export-pdf="onExportPdf"
      @about="isAboutOpen = true"
      @shortcuts-help="isShortcutsHelpOpen = true"
      @exit-app="onExit"
      @open-find="openFindDialog('find')"
      @open-replace="openFindDialog('replace')"
      @open-link="() => void linkDialog.open()"
      @toggle-focus-mode="toggleFocusMode"
      @insert-citation="onInsertCitation"
      @search-otzaria="onSearchOtzaria"
      @open-library="onOpenLibrary"
      @manage-macros="isMacrosOpen = true"
      @macro-record="onMacroRecord"
      @macro-play="onMacroPlay"
      @toggle-book-completion="onToggleBookCompletion"
    />

    <!-- שורת הסרגל האופקי. הפינה שלפניו רחבה כמו הסרגל האנכי, וכך שניהם
         מתחילים בדיוק במקום שבו אזור המסמך מתחיל — כמו ב-Word. -->
    <div class="ruler-row">
      <div
        v-show="isRulerVisible"
        class="ruler-corner"
      />
      <DocumentRuler
        :visible="isRulerVisible"
        :reading="rulerReading"
        :host="rulerHost"
        :viewport-source="rulerViewport"
        :unit="rulerUnit"
        :zoom="zoom.value"
        :editable="isDocumentEditable"
        @changed="onRulerChanged"
      />
    </div>

    <!-- אזור המסמך: הסרגל האנכי ולצדו ה-stack שהמנוע מצייר בתוכו -->
    <div class="editor-area">
      <VerticalRuler
        :visible="isRulerVisible"
        :reading="rulerReading"
        :host="rulerHost"
        :viewport-source="rulerViewport"
        :unit="rulerUnit"
        :zoom="zoom.value"
        :editable="isDocumentEditable"
        @changed="onRulerChanged"
      />
      <main
        ref="editorStackRef"
        class="editor-stack"
      />
      <PageBorderOverlay
        :host="rulerHost"
        :viewport-source="rulerViewport"
        :reading="pageBorders"
      />
      <LineNumberOverlay
        :host="rulerHost"
        :viewport-source="rulerViewport"
        :reading="lineNumbering"
      />
      <PilcrowOverlay
        :host="rulerHost"
        :viewport-source="rulerViewport"
        :blocks="formattingMarksBlocks"
        :visible="formattingMarksVisible"
      />
    </div>

    <!-- תפריט הלחצן הימני. אחרי אזור המסמך ולפני הדיאלוגים, כמו ה-z-index שלו. -->
    <ContextMenu
      :open="contextMenu.isOpen.value"
      :point="contextMenu.point.value"
      :sections="contextMenu.sections.value"
      @run="contextMenu.run"
      @close="closeContextMenu"
    />

    <LinkDialog
      :is-open="linkDialog.isOpen.value"
      :has-range="linkDialog.selection.value.hasRange"
      :selected-text="linkDialog.selection.value.text"
      @close="linkDialog.close()"
      @submit="linkDialog.submit"
    />

    <!-- שורת מצב תחתונה -->
    <StatusBar
      :current-page="docMetrics.currentPage"
      :total-pages="docMetrics.totalPages"
      :word-count="docMetrics.words"
      :status-text="statusText"
      :is-error="isStatusError"
      :is-focus-mode="isFocusMode"
      :zoom-level="zoom.value"
      :zoom-min="zoom.min"
      :zoom-max="zoom.max"
      @update:zoom-level="onZoomChange"
      @toggle-focus="toggleFocusMode"
    />

    <!-- דיאלוגים ופאנלים -->
    <FindReplaceDialog
      :is-open="isFindOpen"
      :initial-mode="findMode"
      :result-text="searchCounter"
      :can-replace="canShowReplace"
      :is-replacing="searchState.isReplacing"
      @close="closeFindDialog"
      @find="onFindText"
      @query-change="onFindQueryChange"
      @replace="onReplaceText"
      @replace-all="onReplaceAllText"
    />

    <AboutDialog
      :is-open="isAboutOpen"
      @close="isAboutOpen = false"
    />

    <ShortcutsDialog
      :is-open="isShortcutsHelpOpen"
      @close="isShortcutsHelpOpen = false"
    />

    <MacrosDialog
      :is-open="isMacrosOpen"
      :handle="activeMacros"
      @close="isMacrosOpen = false"
      @status="setStatus"
    />

    <!--
      הטולטיפ של כל התוכנה — מופע אחד, בסוף המעטפת. הוא מאזין במסירה על המסמך
      ולא נקשר לפקד מסוים, ולכן אין לו props: כל פקד שיש לו `title` או
      `data-tip-*` מקבל אותו. ההסבר המלא ב-ui/tooltip/TooltipLayer.vue.
    -->
    <TooltipLayer />
  </div>
</template>

<script setup lang="ts">
import { ref, provide, onMounted, onUnmounted, computed, shallowRef, watch } from 'vue';
import TitleBar from './ui/shell/TitleBar.vue';
import Ribbon from './ui/ribbon/Ribbon.vue';
import StatusBar from './ui/shell/StatusBar.vue';
import DocumentRuler from './ui/shell/DocumentRuler.vue';
import VerticalRuler from './ui/shell/VerticalRuler.vue';
import PageBorderOverlay from './ui/shell/PageBorderOverlay.vue';
import LineNumberOverlay from './ui/shell/LineNumberOverlay.vue';
import PilcrowOverlay from './ui/shell/PilcrowOverlay.vue';
import FindReplaceDialog from './ui/panels/FindReplaceDialog.vue';
import AboutDialog from './ui/panels/AboutDialog.vue';
import LinkDialog from './ui/panels/LinkDialog.vue';
import ShortcutsDialog from './ui/panels/ShortcutsDialog.vue';
import TooltipLayer from './ui/tooltip/TooltipLayer.vue';
import { createCommandAdapter, type CommandAdapter, type CommandOutcome } from './engine/command-adapter';
import type { CommandId } from './engine/capabilities';
import {
  COMMAND_ADAPTER,
  COMMAND_REPORTER,
  DOCUMENT_GENERATION,
  FONT_OPTIONS,
  READOUT_SELECTION,
  STYLE_GALLERY,
} from './composables/keys';
import { ACTIVE_SUPERDOC } from './engine/document-api';
import { readDocSelection } from './engine/doc-selection';
import {
  buildCitationText,
  getReaderSelection,
  insertCitation,
  normalizeSelectedText,
  openLibrary,
  openSearchTab,
  type ReaderResult,
} from './host/otzaria-reader';
import {
  fallbackStyleGallery,
  observeStyleGallery,
  type StyleGalleryState,
} from './engine/style-gallery';
import { fallbackFontOptions, observeFontOptions, type FontOptions } from './engine/font-options';
import {
  UNSETTLED_SELECTION,
  observeReadoutSelection,
  type ReadoutSelection,
} from './engine/readout-hold';
import { zoomPayload } from './engine/payloads';
import {
  createSearchAdapter,
  idleSearchState,
  replaceControlsVisible,
  searchCounterText,
  type SearchAdapter,
  type SearchOutcome,
  type SearchState,
} from './engine/search';
import { createEditorSwap, type EditorSwap } from './sessions/editor-swap';
import { createSaveCoordinator, type SaveCoordinator, type SaveSnapshot } from './sessions/save-coordinator';
import { createEditor, type EditorSession } from './engine/create-editor';
import { ACTIVE_MACROS, installMacros, type MacrosHandle } from './engine/macros';
import MacrosDialog from './ui/panels/MacrosDialog.vue';
import { installBookCompletion } from './engine/book-completion-overlay';
import { preflightSource } from './engine/docx-preflight';
import { installDocumentFontAliases } from './engine/docx-fonts';
import {
  anchorPageIndex,
  createDocMetrics,
  emptyDocMetrics,
  readDocumentInfo,
  type DocMetrics,
  type DocMetricsAdapter,
} from './engine/doc-metrics';
import { FALLBACK_ZOOM, observeZoom, type ZoomState } from './engine/zoom';
import { createZoomCenter, type ZoomCenter } from './engine/zoom-center';
import {
  createRulerModel,
  paintedHost,
  readRulerUnit,
  type RulerModel,
  type RulerReading,
  type ViewportSource,
} from './engine/page-ruler';
import {
  createLineNumberingModel,
  createPageBorderModel,
  readLineNumbering,
  readPageBorders,
  readPageMargins,
  type LineNumberingModel,
  type LineNumberingReading,
  type PageBorderModel,
  type PageBordersReading,
} from './engine/page-setup';
import {
  createFormattingMarksModel,
  readFormattingMarksBlocks,
  type FormattingMarksModel,
} from './engine/formatting-marks';
import type { FormattingMarksBlock } from './engine/formatting-marks-layer';
import { readParagraphIndents } from './engine/paragraph-format';
import type { RulerUnit } from './engine/ruler-geometry';
import {
  applyHebrewDocumentDefaults,
  applyHebrewPaperSize,
} from './engine/document-defaults';
import type { SuperDoc } from 'superdoc';
import { DOCX_MIME, exportDocx, docxFileName } from './engine/export';
import { exportPdfDocument, pdfSuggestedName, printDocument } from './engine/print';
import { downloadBlob } from './host/download';
import {
  beginBinaryWrite,
  uploadBytes,
  abortBinaryWrite,
  commitUserFileWrite,
  pickDocxFile,
  resolveFileUrl,
  type UserFile,
} from './host/files';
import { decideDocumentSwitch } from './sessions/open-flow';
import { call, confirm, notifyError } from './host/otzaria-client';
import { supportsPdfExport } from './host/host-capabilities';
import { splashDone } from './host/splash';
import {
  loadLastDocument,
  forgetLastDocument,
  loadAutosaveEnabled,
  saveAutosaveEnabled,
  loadRulerVisible,
  saveRulerVisible,
  loadSessionRecord,
  saveSessionRecord,
} from './host/settings';
import {
  DRAFT_PATH,
  decideDraftRecovery,
  documentViewFor,
  draftAgeLabel,
  normalizeSession,
  sessionFromLastDocument,
  type SessionState,
} from './sessions/session-state';
import { createSessionKeeper, type SessionKeeper } from './sessions/session-keeper';
import { applyCaretAnchor, readCaretAnchor, type CaretAnchor } from './engine/caret-anchor';
import {
  deleteWorkspaceEntry,
  readWorkspaceBytes,
  writeWorkspaceBytes,
} from './host/workspace';
import { onPluginHidden } from './host/lifecycle';
import { revealZone, type RevealZone } from './composables/focus-mode';
import { selectWholeDocument } from './engine/clipboard';
import {
  DEFAULT_FONT_SIZE_PT,
  fontSizePayload,
  grownFontSize,
  parseFontSizePt,
  shrunkFontSize,
} from './engine/payloads';
import { toggleVertAlign } from './engine/vert-align';
import { insertNote } from './engine/footnotes';
import { startParagraphOnNewPage, pageBreakTracker } from './engine/page-break';
import { createLinkDialog } from './composables/use-link-dialog';
import { createShellActionRunner } from './ui/shortcuts/actions';
import { useContextMenu } from './composables/use-context-menu';
import ContextMenu from './ui/menu/ContextMenu.vue';
import {
  createShortcutDispatcher,
  isTextEntryTarget,
  type ShortcutDispatcher,
} from './ui/shortcuts/dispatch';
import { createDirectionShortcut } from './ui/shortcuts/direction';
import { watchUndoRedoKeys, type UndoRedoWatcher } from './ui/shortcuts/undo-redo-watch';
import { createFocusRing } from './ui/shortcuts/focus-ring';
import { focusDocument } from './engine/focus';

const editorStackRef = ref<HTMLElement | null>(null);
const shellRef = ref<HTMLElement | null>(null);

const commandAdapter = shallowRef<CommandAdapter | null>(null);
provide(COMMAND_ADAPTER, commandAdapter);

/**
 * אפשרויות הגופן של המסמך הפתוח. מסופקות מכאן ולא נקראות בקומפוננטה, כי
 * `ui.fonts` הוא handle של ה-session — מסמך חדש מביא רשימה חדשה, ורק מי שמנהל
 * את ה-session יודע מתי. הקומפוננטה רואה מפתח צר (`FONT_OPTIONS`) ולא את `ui`.
 */
const fontOptions = shallowRef<FontOptions>(fallbackFontOptions());
provide(FONT_OPTIONS, fontOptions);

/**
 * גלריית הסגנונות של המסמך הפתוח. מאותו טעם כמו אפשרויות הגופן, וביתר שאת:
 * `ui.styles` פותר את הקטלוג **אסינכרונית** אחרי הפתיחה, ולכן קריאה חד-פעמית
 * מחזירה רשימה ריקה — רק מי שמנהל את ה-session יודע מתי להירשם.
 */
const styleGallery = shallowRef<StyleGalleryState>(fallbackStyleGallery());
provide(STYLE_GALLERY, styleGallery);

/**
 * מצב הבחירה, בשביל החזקת החיווי ברצועה (engine/readout-hold.ts).
 *
 * מסופק מכאן ולא נקרא בקומפוננטה, מאותו טעם כמו שני המפתחות שמעליו:
 * `ui.selection` הוא handle של ה-session, ורק מי שמנהל אותו יודע מתי להירשם
 * ומתי לשחרר. הזרקה אחת לכל הרצועה — כל 38 הפקדים שואלים את אותה שאלה.
 */
const readoutSelection = shallowRef<ReadoutSelection>(UNSETTLED_SELECTION);
provide(READOUT_SELECTION, readoutSelection);

/**
 * המופע הפתוח, בשביל הפקדים שאין להם פקודה ב-registry של ה-controller —
 * שוליים, כיוון דף, עמודות, הערות שוליים. המסלול הציבורי היחיד שלהם הוא
 * ה-Document API, והוא יושב על המופע ולא על ה-controller. ראו engine/document-api.ts.
 */
const activeSuperdoc = shallowRef<SuperDoc | null>(null);
provide(ACTIVE_SUPERDOC, activeSuperdoc);

/**
 * ה-container של המסמך הפתוח — installBookCompletion (engine/book-completion-
 * overlay.ts) מותקן עליו, לא על editorStackRef: זה ה-container הספציפי
 * שהמנוע מרנדר לתוכו, ראו create-editor.ts:EditorSession.container.
 */
const activeEditorContainer = shallowRef<HTMLElement | null>(null);

/**
 * מונה "מסמך אחר" — ראו ההסבר המלא ב-composables/keys.ts. מעודכן מ-`swap.
 * documentGeneration` באותו רגע בדיוק שבו `activeSuperdoc` מוחלף (openDocument),
 * כדי ששני העדכונים יגיעו לצרכנים באותו tick.
 */
const documentGeneration = shallowRef(0);
provide(DOCUMENT_GENERATION, documentGeneration);

/**
 * מערכת המאקרו של ה-session, בשביל כפתורי הרצועה ודיאלוג הניהול. אותו דפוס
 * כמו `activeSuperdoc`: נקבעת אחרי פתיחה מוצלחת ומתאפסת בפירוק. ראו
 * engine/macros.ts.
 */
const activeMacros = shallowRef<MacrosHandle | null>(null);
provide(ACTIVE_MACROS, activeMacros);

/** דיאלוג ניהול המאקרו (Alt+F8). */
const isMacrosOpen = ref(false);

/**
 * שני המטפלים מחזירים „האם טופל”, בשביל מסלול הקיצורים: בלי מסמך פתוח אין
 * מערכת מאקרו, והצירוף צריך להישאר של הדפדפן. הרצועה מתעלמת מערך ההחזרה —
 * הכפתורים שלה ממילא מנוטרלים בלי מסמך.
 */
function onMacroRecord(): boolean {
  const macros = activeMacros.value;
  if (!macros) return false;
  macros.toggleRecording();
  return true;
}

function onMacroPlay(): boolean {
  const macros = activeMacros.value;
  if (!macros) return false;
  macros.replayLast();
  return true;
}

/**
 * האם יש מסמך פתוח — מה שפקדי לשונית „קובץ” נשענים עליו.
 *
 * נגזר מ-`activeSuperdoc` ולא מ-`swap?.current`, שזו הבדיקה שהמטפלים עצמם
 * עושים: `swap` הוא משתנה רגיל ולא מצב reactive, ולכן פקד שהיה נשען עליו לא
 * היה מתעדכן כשמסמך נפתח או נסגר. שני הערכים עולים ונופלים יחד — `activeSuperdoc`
 * נקבע מיד אחרי פתיחה מוצלחת ומתאפס בפירוק ה-session.
 */
const hasDocument = computed(() => activeSuperdoc.value !== null);

const title = ref('מסמך חדש');
const isOpening = ref(false);
const autosaveEnabled = ref(true);
const statusText = ref('');
const isStatusError = ref(false);
const isFocusMode = ref(false);
const revealed = ref<RevealZone>(null);
const bookCompletionEnabled = ref(false);

/**
 * הלשונית ברצועה ומצב הכיווץ. הוחזקו עד עכשיו בתוך `Ribbon.vue` עצמו, ועלו
 * לכאן מסיבה אחת: הם שורדים הפעלות, ומי שזוכר יושב כאן. ההנמקה המלאה בראש
 * ההגדרה ב-Ribbon.vue.
 */
const ribbonTab = ref('home');
const ribbonCollapsed = ref(false);

// `watch` ולא מטפל על הרצועה: הרצועה מחליפה לשונית משלושה מקומות — קליק,
// חצים, ולחיצה כפולה שמכווצת — ומטפל היה צריך להיקרא בכל אחד מהם. השינוי
// עצמו הוא מה שמעניין, ולכן מאזינים לו ולא למי שגרם לו.
watch([ribbonTab, ribbonCollapsed], ([tab, collapsed]) => {
  keeper?.updateView({ ribbonTab: tab, ribbonCollapsed: collapsed });
});

const isFindOpen = ref(false);
const findMode = ref<'find' | 'replace'>('find');
const isAboutOpen = ref(false);
const isShortcutsHelpOpen = ref(false);

/**
 * מצב החיפוש כפי שהמנוע מדווח עליו. הדיאלוג נשען עליו למונה התוצאות ולשאלה
 * אם להציג פקדי החלפה — ולא על state מקומי משלו, שהיה יכול להראות „3 מתוך 12”
 * על מסמך שהחיפוש בו כלל לא רץ.
 */
const searchState = ref<SearchState>(idleSearchState());
const searchCounter = computed(() => searchCounterText(searchState.value));

/**
 * האם הדיאלוג מציג את פקדי ההחלפה. **לא** `searchState.canReplace`: הדגל ההוא
 * תלוי בקבוצת ההתאמות הנוכחית, ולכן חיבורו הישיר לכאן העלים את שדה ההחלפה
 * ברגע שהמשתמש הקליד מילה שאינה במסמך — ובמקומו הופיעה הודעה שהאשימה את גרסת
 * המנוע. ההכרעה עצמה ב-engine/search.ts, כדי שתהיה נבדקת.
 */
const canShowReplace = computed(() => replaceControlsVisible(searchState.value));

/**
 * מה ששורת המצב מציגה. שלושת הערכים היו `ref(1)`, `ref(1)` ו-`ref(0)` שלא
 * התעדכנו מעולם — „עמוד 1 מתוך 1” ו„0 מילים” על כל מסמך. עכשיו הם מדידה,
 * ו-`null` בהם פירושו „טרם נמדד” ולא מספר (ראו engine/doc-metrics.ts).
 */
const docMetrics = ref<DocMetrics>(emptyDocMetrics());

/** גודל התצוגה והגבולות שהמנוע מתיר. הסרגל לא מקודד אותם יותר. */
const zoom = ref<ZoomState>({ ...FALLBACK_ZOOM });

const canUndo = ref(false);
const canRedo = ref(false);

/**
 * סרגל המידות.
 *
 * ## מי בעל המצב
 *
 * **המנוע.** `ruler` היא פקודה ב-registry שלו, והיא מנותבת ל-
 * `SuperDoc.toggleRuler()` שמחליף את `config.rulers`; המצב `active` של הפקודה
 * נקרא מאותו דגל. נמדד על המנוע האמיתי שההרצה מודיעה למי שמאזין — כלומר
 * הכפתור בלשונית „תצוגה” נדלק, והסרגל כאן מופיע, מאותו מקור אחד. הפיתוי היה
 * להחזיק כאן `ref` משלנו ולקרוא ל-`toggleRuler` בצד; זה היה יוצר שני מצבים
 * שיכולים להיפרד, ומצב שני הוא בדיוק מה שהופך פקד ל„לפעמים לא עובד”.
 *
 * מה שהמנוע **אינו** עושה הוא לצייר סרגל: `ui: false` מכבה את הסרגל המובנה
 * שלו (הוא `suppressed` ולא רק כבוי), ולכן הכפתור עד עכשיו הדליק דגל שאיש לא
 * הסתכל עליו. הציור הוא שלנו — ui/shell/DocumentRuler.vue.
 */
const isRulerVisible = ref(false);
const rulerReading = shallowRef<RulerReading | null>(null);
/** ה-host המצויר של המסמך הפתוח, ומקור אירועי הגיאומטריה. */
const rulerHost = shallowRef<HTMLElement | null>(null);
const rulerViewport = shallowRef<ViewportSource | null>(null);
const rulerUnit = ref<RulerUnit>('cm');
/** `false` במסמך שפתוח לקריאה בלבד — אז הידיות אינן נגררות. */
const isDocumentEditable = ref(true);

/**
 * מצב „גבולות עמוד” של המסמך הפתוח, ל-ui/shell/PageBorderOverlay.vue.
 * `null` כשאין `<w:pgBorders>` — השכבה מציירת אפס גבולות, לא גבול ריק.
 */
const pageBorders = shallowRef<PageBordersReading | null>(null);
/**
 * מצב „מספרי שורות” של המסמך הפתוח, ל-ui/shell/LineNumberOverlay.vue.
 * `null` כשאין `<w:lnNumType>` — השכבה מציירת אפס מספרים, לא מספור ריק.
 */
const lineNumbering = shallowRef<LineNumberingReading | null>(null);
/**
 * בלוקי המסמך הנוכחיים ל-ui/shell/PilcrowOverlay.vue, או `null` כשסימני
 * העיצוב כבויים (ברירת המחדל) או כשאין Document API. בשונה מ-`pageBorders`/
 * `lineNumbering` (שמשקפים מה שקיים ב-docx), זה קלט לחישוב **גיאומטרי** —
 * הציור בפועל תלוי גם ב-DOM (engine/formatting-marks-layer.ts).
 */
const formattingMarksBlocks = shallowRef<readonly FormattingMarksBlock[] | null>(null);
/** מצב הפקד „הצג/הסתר סימני עיצוב", ל-`visible` של PilcrowOverlay.vue. */
const formattingMarksVisible = ref(false);
/** ההעדפה שנשמרת בין הפעלות. ראו host/settings.ts. */
let rulerPreference = false;

const saveSnapshot = ref<SaveSnapshot>({
  state: 'idle',
  isDirty: false,
  isSaving: false,
  targetToken: null,
  name: null,
  lastError: null,
});

let swap: EditorSwap | null = null;
let save: SaveCoordinator | null = null;
let searchAdapter: SearchAdapter | null = null;
/** מודד את המסמך הפתוח. מוחלף בכל מעבר מסמך, כמו אדפטר החיפוש. */
let metrics: DocMetricsAdapter | null = null;
/** קורא את מצב הסרגל של המסמך הפתוח. שייך ל-session, כמו המודד. */
let ruler: RulerModel | null = null;
/** קורא את מצב „גבולות עמוד” של המסמך הפתוח. שייך ל-session, כמו הסרגל. */
let pageBorderModel: PageBorderModel | null = null;
/** קורא את מצב „מספרי שורות” של המסמך הפתוח. שייך ל-session, כמו גבולות עמוד. */
let lineNumberModel: LineNumberingModel | null = null;
/** קורא את בלוקי המסמך עבור „סימני עיצוב”. שייך ל-session, כמו שני אלה שמעל. */
let formattingMarksModel: FormattingMarksModel | null = null;

/** מרכוז העמוד בזום. יחיד למאגס — אינו מוחלף בין מסמכים. */
let zoomCenter: ZoomCenter | null = null;

/**
 * זוכר ההפעלה. יחיד למעטפת ואינו מוחלף בין מסמכים — הוא מה שמחזיק את הרשומה
 * שעוברת מהפעלה להפעלה. ההנמקה המלאה ב-sessions/session-keeper.ts.
 */
let keeper: SessionKeeper | null = null;
/** מבטל את ההאזנה למעבר לרקע. */
let hiddenListener: (() => void) | null = null;

const saveStateMessage = computed(() => {
  const state = saveSnapshot.value.state;
  if (state === 'exporting') return 'מייצא…';
  if (state === 'uploading' || state === 'committing') return 'שומר…';
  if (state === 'error') return 'שגיאה בשמירה';
  if (saveSnapshot.value.isDirty) return 'שינויים לא שמורים';
  if (saveSnapshot.value.targetToken) return 'נשמר';
  return 'טרם נשמר';
});

function setStatus(text: string, isError = false): void {
  statusText.value = text;
  isStatusError.value = isError;
  if (isError) notifyError(text);
}

/**
 * כל פקד ב-Ribbon מדווח לכאן דרך useCommand. עד עכשיו הפקדים עשו
 * `void cmd.run()` וזרקו את התוצאה, ולכן „יש למקם את הסמן במסמך” או „הפעולה
 * אינה נתמכת בגרסה הזאת של המנוע” לא הגיעו למשתמש אף פעם — הכפתור פשוט נראה
 * שבור. כאן ההודעה נכנסת לשורת המצב, ובכשל גם ללוג של אוצריא.
 */
function reportCommand(outcome: CommandOutcome, commandId: string): void {
  if (!outcome.ok) {
    setStatus(outcome.message, true);
    console.warn(`[otzaria-word] ${commandId} נכשלה: ${outcome.message} (${outcome.reason ?? '—'})`);
    return;
  }
  // הצלחה מנקה שגיאה קודמת שנשארה על המסך, ולא דורסת הודעה תקינה.
  if (isStatusError.value) setStatus('');

  /**
   * „גבולות עמוד” — נמדד ב-QA: `applyPageBorders`/`clearPageBorders`
   * (engine/page-setup.ts, קריאת section-level) **אינן** מפעילות את
   * `onUpdate` של המנוע בעצמן — בשונה משוליים/כיוון/עמודות, ששינוי שלהם
   * כן מגיע דרך `onUpdate` (הערה ב-createEditor למעלה). בלי הרענון המפורש
   * הזה `pageBorders.value` נשאר ישן עד לעריכת טקסט הבאה, כלומר גבול
   * שנבחר לא מצטייר במשך שניות, וגבול שהוסר נשאר על המסך כ„גבול רפאים”.
   * `refreshNow` ולא `noteDocumentChanged`: כאן הפעולה כבר הצליחה, ואין
   * טעם בהשקטה שנועדה למנוע הצפת קריאות בזמן עריכה רציפה.
   */
  if (commandId === 'page-borders') pageBorderModel?.refreshNow();

  /**
   * „מספרי שורות” — אותה מלכודת בדיוק כמו „גבולות עמוד” שמעל, ואותו תיקון:
   * `applyLineNumbering` (engine/page-setup.ts, קריאת section-level) אינה
   * מפעילה `onUpdate` בעצמה, ובלעדי הרענון המפורש הזה `lineNumbering.value`
   * נשאר ישן עד לעריכת טקסט הבאה — בחירה בתפריט לא הייתה מצטיירת עד אז.
   */
  if (commandId === 'page-line-numbering') lineNumberModel?.refreshNow();
}

provide(COMMAND_REPORTER, reportCommand);

function initSaveCoordinator(): SaveCoordinator {
  return createSaveCoordinator({
    exportDocument: () => {
      const active = swap?.current;
      if (!active) throw new Error('אין מסמך פתוח');
      return exportDocx(active.superdoc);
    },
    beginWrite: (size) => beginBinaryWrite(size),
    upload: uploadBytes,
    abort: abortBinaryWrite,
    commit: (input) =>
      commitUserFileWrite({
        writeToken: input.writeToken,
        targetToken: input.targetToken,
        suggestedName: input.suggestedName ?? title.value,
        title: 'שמירת המסמך',
      }),
    onStateChange: (snapshot) => {
      saveSnapshot.value = snapshot;
    },
    /**
     * כאן, ולא באתר הקריאה ל-`saveNow`.
     *
     * שלושה מסלולים מגיעים לשמירה — „שמור” של המשתמש, „לשמור לפני שפותחים
     * אחר”, ושמירה אוטומטית — ורק הראשון עובר במעטפת. תלייה על `onSave`
     * בלבד השאירה את טיוטת השחזור חיה אחרי כל שמירה אוטומטית, ומכיוון
     * שהיא מפסיקה להתעדכן ברגע שהמסמך נקי, היא הייתה נפתחת בהפעלה הבאה
     * **מעל עבודה חדשה ממנה** — ואז נכתבת לקובץ. כאן זה מסלול אחד לכולם.
     */
    onSaved: (info) => {
      // „שמור בשם” מחליף את הקובץ שמאחורי המסמך; „שמור” רגיל אינו. רק
      // הראשון הוא זהות חדשה, ורק אז נכון לשכוח את מקום הסמן.
      if (keeper && keeper.state.document?.token !== info.token) {
        keeper.setDocument({ token: info.token, name: info.name, writable: true });
      }
      // הגודל הוא של מה שנכתב עכשיו, ולכן הוא הבסיס להשוואה הבאה מול
      // הדיסק — בלעדיו „הקובץ השתנה מבחוץ” היה נשאל אחרי כל שמירה רגילה.
      void keeper?.noteSaved(info.size);
    },
  });
}

/**
 * זוכר ההפעלה. כל התלויות שלו הן פונקציות של המעטפת ולא ייבוא ישיר, מאותו
 * טעם כמו בקואורדינטור השמירה: הקוד שמחליט מתי לכתוב את העבודה של המשתמש
 * צריך להיות נבדק בלי מנוע ובלי גשר.
 */
function initSessionKeeper(): SessionKeeper {
  return createSessionKeeper({
    persist: (state) => saveSessionRecord(state),
    exportDocument: () => {
      const active = swap?.current;
      if (!active) throw new Error('אין מסמך פתוח');
      return exportDocx(active.superdoc);
    },
    // ההמרה מ-`Blob` כאן ולא ב-host/workspace.ts: כאן יושב מי שמחזיק את
    // המנוע, ושם יושב מי שמדבר עם הגשר.
    writeDraft: async (content) =>
      writeWorkspaceBytes(DRAFT_PATH, new Uint8Array(await content.arrayBuffer())),
    removeDraft: () => deleteWorkspaceEntry(DRAFT_PATH),
    draftPath: DRAFT_PATH,
    readCaret: (previous) => {
      const active = swap?.current;
      if (!active) return Promise.resolve(null);
      return readCaretAnchor(active.ui, active.superdoc, previous);
    },
    isDirty: () => save?.snapshot.isDirty === true,
    isSaving: () => save?.snapshot.isSaving === true,
    settleSave: () => save?.settled() ?? Promise.resolve(),
    // שגיאה ולא הודעה רגילה: זו אינה התקדמות אלא רשת ביטחון שאינה פרושה,
    // והמשתמש צריך לדעת שעליו לשמור בעצמו.
    onDraftTooLarge: () => {
      setStatus(
        'המסמך גדול מכדי לשמור ממנו עותק לשחזור — שינויים שלא יישמרו לקובץ עלולים לאבוד',
        true,
      );
    },
  });
}

/**
 * כיווניות עברית למסמך חדש. ההחלה עצמה ב-engine/document-defaults.ts; כאן רק
 * הדיווח — כשל שקט הוא בדיוק מה שהחזיר מסמך חדש ל-LTR בלי שאף אחד ידע.
 *
 * `data-document-direction` על שורש ה-HTML הוא מה שאפשר לראות מבחוץ:
 * שער `check:rtl` נשען עליו, ובלוג של אוצריא הוא מפריד בין „לא הוחל” ל„הוחל
 * ולא נראה”.
 */
async function applyNewDocumentDirection(superdoc: SuperDoc): Promise<void> {
  const report = await applyHebrewDocumentDefaults(superdoc);

  if (report.failures.length === 0) {
    document.documentElement.dataset.documentDirection = 'rtl';
    return;
  }

  delete document.documentElement.dataset.documentDirection;
  console.warn('[otzaria-word] כיווניות המסמך החדש לא הוחלה במלואה:', report.failures.join('; '));
  setStatus(`המסמך נפתח, אך כיווניות עברית לא הוחלה: ${report.failures[0]}`, true);
}

/**
 * גודל הדף של מסמך חדש: A4 ולא ה-Letter שהמסמך הריק של המנוע נושא. ההחלה
 * ב-engine/document-defaults.ts; כאן רק הדיווח.
 *
 * דיווח נפרד מזה של הכיווניות, ובכוונה: `data-document-direction` ושער
 * `check:rtl` מודדים את שלוש שכבות הכיווניות, וכשל בגודל הדף אינו כשל
 * כיווניות. גם ההודעה כאן אינה מזכירה „כיווניות” — השער סורק את הלוג על המילה
 * הזאת, וכשל בגודל דף אסור לו להיראות שם ככשל כיווניות.
 */
async function applyNewDocumentPaperSize(superdoc: SuperDoc): Promise<void> {
  const report = await applyHebrewPaperSize(superdoc);
  if (report.applied) return;

  console.warn('[otzaria-word] גודל הדף של המסמך החדש לא הוגדר ל-A4:', report.failure);
  setStatus(`המסמך נפתח, אך גודל הדף לא הוגדר ל-A4: ${report.failure}`, true);
}

/** מה שאינו נגזר מהקובץ עצמו — המסלולים של „חזרה למה שהיה”. */
interface OpenOptions {
  /**
   * בייטים לפתוח **במקום** ה-URL של הקובץ: טיוטת השחזור. המסמך שנפתח כך
   * מסומן מיד כלא-שמור, כי זה בדיוק מה שהוא — עבודה שאינה בדיסק.
   */
  draft?: Blob;
  /** גודל התצוגה והסמן שיוחזרו אחרי הפתיחה. ראו restoreDocumentView. */
  restore?: { zoom: number | null; caret: CaretAnchor | null };
  /**
   * האם לרשום את מה שנפתח כמסמך של ההפעלה. ברירת המחדל: כן.
   *
   * `false` למסלול אחד — מסמך ריק שנפתח מפני שהמסמך האמיתי **לא הצליח**
   * להיפתח. רישום שלו היה מוחק מהרשומה את המסמך האחרון, ואיתו את הדרך
   * לנסות שוב בהפעלה הבאה: כשל בפתיחה עשוי להיות זמני, והטיוטה שמחזיקה את
   * העבודה מזוהה מול אותו token בדיוק.
   */
  remember?: boolean;
}

async function openDocument(file?: UserFile, options: OpenOptions = {}): Promise<boolean> {
  if (!swap) return false;
  isOpening.value = true;
  const startedAt = performance.now();
  setStatus(file ? `פותח את ${file.name}…` : 'פותח מסמך ריק…');

  // לפני המנוע ולא אחריו: ערך אחד ב-`word/settings.xml` שולח אותו ללולאה על
  // החוט הראשי, ומשם אין חזרה — גם `OPEN_TIMEOUT_MS` אינו יכול לירות. ראו
  // engine/docx-preflight.ts. השלב הזה אינו יכול להיכשל: הוא מחזיר את המקור
  // כמות שהוא בכל מקרה שאינו „מצאתי בדיוק את הערך הזה”.
  const { source, fontTable } = await preflightSource(options.draft ?? file?.url);

  // לפני שהמנוע מודד, ולא אחרי: `lineRule="auto"` גוזר את גובה השורה ממדדי
  // הגופן שנבחר בפועל, ולכן גופן חסר משנה את פריסת כל המסמך — לא רק את מראהו.
  // ראו engine/docx-fonts.ts. אינו יכול להיכשל, ואינו מעכב כשאין מה להחליף.
  await installDocumentFontAliases(fontTable);

  const outcome = await swap.open(source);
  isOpening.value = swap.isOpening;

  if (outcome.status === 'superseded') return false;

  if (outcome.status === 'failed') {
    const kept = swap.current ? ` ${title.value} נשאר פתוח.` : '';
    setStatus(`פתיחת המסמך נכשלה: ${outcome.error.message}.${kept}`, true);
    return false;
  }

  const editor = outcome.session;
  const adapter = createCommandAdapter(editor.ui);
  commandAdapter.value = adapter;

  // ה-`editor.superdoc` המקומי ולא `activeSuperdoc.value` בפירוק: אותה מלכודת
  // כמו באדפטר החיפוש — סגירת המסמך הקודם קורית אחרי שהחדש כבר נרשם.
  activeSuperdoc.value = editor.superdoc;
  activeEditorContainer.value = editor.container;
  // אותו tick בדיוק כמו ההשמה שמעל: מי שמשווה זהות `documentGeneration` בין
  // שתי קריאות (`PageBreakTracker.syncDocument`) חייב לראות את שתיהן יחד.
  documentGeneration.value = swap.documentGeneration;
  editor.onDispose(() => {
    if (activeSuperdoc.value === editor.superdoc) {
      activeSuperdoc.value = null;
      activeEditorContainer.value = null;
    }
    // בלי האיפוס הרצועה הייתה ממשיכה להחזיק את הקריאה של המסמך שנסגר.
    readoutSelection.value = UNSETTLED_SELECTION;
  });

  // החיפוש שייך ל-session: ה-handle הוא של ה-controller של המופע, ומסמך חדש
  // מקבל אדפטר חדש. ה-`session` המקומי ולא `searchAdapter` בפירוק — אחרת
  // סגירת המסמך הקודם הייתה מפרקת את האדפטר של המסמך שנפתח אחריו.
  // `observe` יורה מיד עם ה-snapshot ואז על כל שינוי: המנוע פותר את גופני
  // המסמך אחרי שהוא נפתח, ובלי האזנה הבורר היה קופא על הרשימה של הרגע הראשון.
  editor.onDispose(
    observeFontOptions(editor.ui, (options) => {
      fontOptions.value = options;
    })
  );

  // אותו טעם, ועוד יותר: `getQuickGallery()` מחזיר רשימה ריקה עד שהקטלוג
  // מתייצב, ולכן בלי ההרשמה הגלריה הייתה נשארת על רשת הביטחון לתמיד.
  editor.onDispose(
    observeStyleGallery(editor.ui, (state) => {
      styleGallery.value = state;
    })
  );

  // מערכת המאקרו (superdoc-macros) שייכת ל-session: ההקלטה עוטפת את
  // ה-controller של המופע הזה, וההקלדה נקלטת מה-host של המסמך הפתוח.
  // אותה תבנית פירוק כמו אדפטר החיפוש — ראו engine/macros.ts. ה-`macros`
  // המקומי ולא `activeMacros` בפירוק, מאותה מלכודת: סגירת המסמך הקודם קורית
  // אחרי שהחדש כבר נרשם.
  //
  // עטוף ב-try/catch, ובכוונה: המאקרו הם פיצ'ר אופציונלי, וכשל בהקמתו —
  // אחסון חסום, מנוע שהשתנה — אסור לו לעצור את פתיחת המסמך. ההתקנה עצמה
  // כבר נופלת לאחסון-זיכרון בכשל localStorage; זו רשת הביטחון למה שמעבר.
  if (editorStackRef.value) {
    try {
      const macros = installMacros(editor, editorStackRef.value, setStatus, {
        // אישור שמירה של הקלטה לא-שלמה — פעולה שאינה ניתנת להקלטה (למשל
        // הכנסת תמונה). הדיאלוג של אוצריא; מחוץ לאוצריא הוא מחזיר false,
        // וההקלטה מבוטלת — שמירה חלקית לא קורית בלי הסכמה.
        confirmIncomplete: (title, content) => confirm({ title, content }),
      });
      activeMacros.value = macros;
      editor.onDispose(() => {
        if (activeMacros.value === macros) activeMacros.value = null;
        macros.dispose();
      });
    } catch (error) {
      console.error('[otzaria-word] מערכת המאקרו לא הופעלה', error);
      setStatus('מערכת המאקרו לא הופעלה — המסמך נפתח בלעדיה', true);
    }
  }

  /**
   * מודד המסמך שייך ל-session: `doc` הוא של המופע הפתוח, ו-`getAnchorRect`
   * קורא את הגיאומטריה של ה-controller שלו. `sessionMetrics` המקומי ולא
   * `metrics` בפירוק — אחרת סגירת המסמך הקודם הייתה מפרקת את המודד של המסמך
   * שנפתח אחריו (אותה מלכודת כמו באדפטר החיפוש).
   */
  const sessionMetrics = createDocMetrics({
    readInfo: () => readDocumentInfo(editor.superdoc),
    readAnchorPageIndex: () => anchorPageIndex(editor.ui),
    onChange: (next) => {
      docMetrics.value = next;
    },
  });
  metrics = sessionMetrics;
  docMetrics.value = sessionMetrics.getState();
  editor.onDispose(() => {
    sessionMetrics.dispose();
    if (metrics === sessionMetrics) {
      metrics = null;
      docMetrics.value = emptyDocMetrics();
    }
  });

  /**
   * הסרגל של ה-session: הוא קורא את המקטע ואת הפסקה של **המסמך הזה**, ולכן
   * הוא נבנה ונפרק איתו — אותה תבנית כמו המודד ואדפטר החיפוש, כולל המשתנה
   * המקומי בפירוק (מסמך שנסגר אחרי שהבא נפתח אינו מפרק את הבא).
   */
  const sessionRuler = createRulerModel({
    readPage: () => readPageMargins(editor.superdoc),
    readIndents: () => readParagraphIndents(editor.superdoc),
    onChange: (next) => {
      rulerReading.value = next;
    },
  });
  ruler = sessionRuler;
  rulerHost.value = paintedHost(editor.ui);
  rulerViewport.value = editor.ui as ViewportSource;
  rulerUnit.value = readRulerUnit(editor.superdoc);
  editor.onDispose(() => {
    sessionRuler.dispose();
    if (ruler === sessionRuler) {
      ruler = null;
      rulerReading.value = null;
      rulerHost.value = null;
      rulerViewport.value = null;
    }
  });

  /**
   * „גבולות עמוד” של ה-session: אותה תבנית בדיוק כמו הסרגל, ומאותה סיבה —
   * הוא קורא את המקטע של **המסמך הזה**, ולכן הוא נבנה ונפרק איתו. קריאה
   * מיידית מיד אחרי היצירה (`refreshNow`) ולא רק בהמתנה ל-`onUpdate` הראשון:
   * מסמך שנפתח עם `<w:pgBorders>` כבר בתוכו (מ-Word) צריך לצייר אותו מיד,
   * לא רק אחרי העריכה הראשונה.
   */
  const sessionPageBorders = createPageBorderModel({
    read: () => readPageBorders(editor.superdoc),
    onChange: (next) => {
      pageBorders.value = next;
    },
  });
  pageBorderModel = sessionPageBorders;
  sessionPageBorders.refreshNow();
  editor.onDispose(() => {
    sessionPageBorders.dispose();
    if (pageBorderModel === sessionPageBorders) {
      pageBorderModel = null;
      pageBorders.value = null;
    }
  });

  /**
   * „מספרי שורות” של ה-session: אותה תבנית בדיוק כמו „גבולות עמוד” שמעל,
   * ומאותה סיבה — קריאה מיידית מיד אחרי היצירה כדי לצייר מסמך שהגיע עם
   * `<w:lnNumType>` כבר בתוכו (מ-Word) מיד, לא רק אחרי העריכה הראשונה.
   */
  const sessionLineNumbering = createLineNumberingModel({
    read: () => readLineNumbering(editor.superdoc),
    onChange: (next) => {
      lineNumbering.value = next;
    },
  });
  lineNumberModel = sessionLineNumbering;
  sessionLineNumbering.refreshNow();
  editor.onDispose(() => {
    sessionLineNumbering.dispose();
    if (lineNumberModel === sessionLineNumbering) {
      lineNumberModel = null;
      lineNumbering.value = null;
    }
  });

  /**
   * „סימני עיצוב” (¶) של ה-session: אותו רעיון כמו שני אלה שמעל, אבל המקור
   * שונה — לא `sections.list()` אלא `doc.blocks.list()` (engine/formatting-marks.ts),
   * ולכן `refreshNow()` **אינה** נקראת כאן: `setEnabled` למטה כבר קוראת
   * כשהפקד דלוק, וקריאה נוספת כאן הייתה סורקת מסמך שלם בכל פתיחה גם כשסימני
   * העיצוב כבויים (ברירת המחדל של הפקד).
   */
  const sessionFormattingMarks = createFormattingMarksModel({
    read: () => readFormattingMarksBlocks(editor.superdoc),
    onChange: (next) => {
      formattingMarksBlocks.value = next;
    },
  });
  formattingMarksModel = sessionFormattingMarks;
  editor.onDispose(() => {
    sessionFormattingMarks.dispose();
    if (formattingMarksModel === sessionFormattingMarks) {
      formattingMarksModel = null;
      formattingMarksBlocks.value = null;
    }
  });

  /**
   * הצגה/הסתרה של סימני העיצוב מגיעה מהמנוע, בדיוק כמו הסרגל — אבל בלי
   * העדפה נשמרת משלה (ברירת המחדל היא כבוי, בכל פתיחת מסמך, כמו ב-Word).
   * בשונה מ„גבולות עמוד”/„מספרי שורות” (שהקריאה של section-level אינה
   * מפעילה `onUpdate`, ולכן נזקקות לרענון מפורש ב-`reportCommand`), הפקודה
   * `formatting-marks` **כן** מדווחת שינוי מצב תקין — נמדד (`docs/superdoc-2.10-review.md`):
   * `active` מתהפך `false→true` על כל לחיצה — ולכן `adapter.observe` בלבד
   * מספיק כאן.
   */
  editor.onDispose(
    adapter.observe('formatting-marks', (state) => {
      if (formattingMarksVisible.value === state.active) return;
      formattingMarksVisible.value = state.active;
      sessionFormattingMarks.setEnabled(state.active);
    })
  );
  formattingMarksVisible.value = adapter.getState('formatting-marks').active;
  sessionFormattingMarks.setEnabled(formattingMarksVisible.value);

  /**
   * מצב הסרגל מגיע מהמנוע ולא ממתג שלנו — ראו `isRulerVisible`. ההרשמה כאן
   * ולא ב-DocumentRuler.vue מפני שהיא גם **כותבת**: העדפת המשתמש נשמרת, וכל
   * מסמך שנפתח מקבל אותה בחזרה (מופע חדש נולד עם `config.rulers: false`).
   *
   * שתי פונקציות ולא אחת, וההפרדה ביניהן היא הנקודה: „מה המנוע מראה” ו„מה
   * המשתמש ביקש” הם שני דברים שנפרדים בדיוק ברגע פתיחת המסמך, שבו המנוע
   * מראה `false` והמשתמש ביקש `true`. מיזוג שלהם לפונקציה אחת מוחק את
   * ההעדפה. את הציר השני — סרגל מוסתר שאינו קורא את המסמך — מחזיק
   * `setEnabled`, והוא יושב בשתיהן דרך `syncRulerVisible`.
   */

  /** שיקוף מצב המנוע, בלי לגעת בהעדפה. */
  const syncRulerVisible = (active: boolean): void => {
    isRulerVisible.value = active;
    sessionRuler.setEnabled(active);
  };

  /** ההחלפה הייתה בחירה של המשתמש, ולכן היא זו שנשמרת להפעלה הבאה. */
  const rememberRulerVisible = (active: boolean): void => {
    if (rulerPreference === active) return;
    rulerPreference = active;
    void saveRulerVisible(active);
  };

  editor.onDispose(
    adapter.observe('ruler', (state) => {
      if (isRulerVisible.value === state.active) return;
      syncRulerVisible(state.active);
      rememberRulerVisible(state.active);
    })
  );

  // ההעדפה נקראת **לפני** הסנכרון, ולא אחריו: הסנכרון מביא את מצב המנוע הטרי
  // (`rulers: false`), ומרגע שהוא רץ אין יותר דרך לדעת מה המשתמש ביקש בהפעלה
  // הקודמת.
  const wanted = rulerPreference;
  syncRulerVisible(adapter.getState('ruler').active);
  // ההחלה אחרי שה-observe רשום: `run` מחליף את הדגל במנוע, וההודעה חוזרת דרך
  // אותו מסלול שהכפתור ברצועה עובר בו.
  if (wanted && !isRulerVisible.value) void adapter.run('ruler');

  // מסמך שפתוח לקריאה בלבד — הידיות בסרגל אינן נגררות בו.
  editor.onDispose(
    adapter.observe('document-mode', (state) => {
      isDocumentEditable.value = state.value !== 'viewing';
    })
  );
  isDocumentEditable.value = adapter.getState('document-mode').value !== 'viewing';

  // עמוד הסמן מגיע מהבחירה, ולכן הוא נקרא כשהיא זזה. בלי ההאזנה המספר היה
  // נכון רק ברגע שהמסמך נפתח. אותה האזנה מזינה גם את הסרגל — סמני הכניסה הם
  // של הפסקה שהסמן בה — וגם את זוכר-ההפעלה: „איפה הסמן” הוא מה שהוא שומר,
  // והשאלה הזאת משתנה בדיוק כאן. הקריאה שלו מושהית — ראו session-keeper.ts.
  editor.onDispose(
    editor.ui.selection.observe(() => {
      sessionMetrics.noteSelectionChanged();
      sessionRuler.noteSelectionChanged();
      keeper?.noteChange();
    })
  );

  // אותה בחירה, שאלה אחרת: האם הקריאה התיישבה, והאם היא סמן או טווח. זה מה
  // שמונע מהחיווי ברצועה להיכבות ולהידלק בכל תו שנקלד — ההנמקה המלאה,
  // כולל המדידה, ב-engine/readout-hold.ts. מנוי נפרד ולא שדה נוסף במודד:
  // המודד שייך לשורת המצב, וזה שייך לרצועה.
  editor.onDispose(
    observeReadoutSelection(editor.ui, (state) => {
      readoutSelection.value = state;
    })
  );

  // גודל התצוגה: `observe` יורה מיד ואז על כל שינוי — כולל שינוי שלא בא
  // מאיתנו (התאמה לרוחב החלון), שאחרת היה משאיר את התווית על ערך שגוי.
  editor.onDispose(
    observeZoom(editor.ui, (state) => {
      zoom.value = state;
      // אותו דיווח מזין גם את המרכוז: הוא יורה על כל שינוי, כולל שינוי שלא
      // בא מאיתנו, ולכן העמוד אינו נשאר ממורכז לפי אחוז ישן. ראו
      // engine/zoom-center.ts.
      zoomCenter?.setZoom(state.value);
      // ואותו דיווח מזין גם את הזיכרון בין הפעלות. כאן ולא במטפל של הסרגל,
      // מאותו טעם בדיוק: גם שינוי שלא בא מאיתנו הוא גודל התצוגה שהמשתמש
      // רואה, והוא זה שצריך לחזור.
      keeper?.updateView({ zoom: state.value });
    })
  );

  // מדידה ראשונה, בלי להמתין לעריכה: מסמך שנפתח צריך להציג את מספר המילים
  // שלו. אם הפאסדה עוד לא מוכנה, הניסיון החוזר תלוי במעבר הפריסה הראשון.
  sessionMetrics.measureNow();
  // אותו טעם, ואותו רגע: סרגל שנפתח על מסמך חדש צריך את השוליים שלו מיד ולא
  // אחרי תזוזת הסמן הראשונה.
  sessionRuler.refreshNow();

  // `editor.superdoc` ולא `editor.ui`: החיפוש-והחלפה העצמאי צריך גם את
  // `activeEditor.doc` (Document API — קריאת בלוקים והחלפה), וגם את `ui`
  // (הדגשת המופע הפעיל) — שניהם חשופים על המופע עצמו, לא רק ה-controller.
  const sessionSearch = createSearchAdapter(editor.superdoc);
  searchAdapter = sessionSearch;
  searchState.value = sessionSearch.getState();
  editor.onDispose(
    sessionSearch.subscribe((state) => {
      searchState.value = state;
    })
  );
  editor.onDispose(() => {
    sessionSearch.dispose();
    if (searchAdapter === sessionSearch) {
      searchAdapter = null;
      searchState.value = idleSearchState();
    }
  });

  title.value = file ? file.name.replace(/\.docx$/i, '') : 'מסמך חדש';
  // זמן הטעינה הוא מדידת פיתוח ולא הודעה למשתמש: „נטען ב-473 מילישניות” תפס
  // את שורת המצב עד ההודעה הבאה. הוא נשמר — הוא מה שמסביר פתיחה איטית —
  // בלוג של אוצריא, במקום שבו מסתכלים על מדידות.
  console.info(
    `[otzaria-word] ${title.value} נטען ב-${Math.round(performance.now() - startedAt)} מילישניות`
  );
  setStatus('');

  if (file && file.access !== 'readwrite') {
    setStatus(`${title.value} — פתוח לקריאה; „שמור” יבקש מקום חדש`);
  }

  save?.reset(file && file.access === 'readwrite' ? { token: file.token, name: file.name } : null);

  if (options.remember !== false) {
    keeper?.setDocument(
      file ? { token: file.token, name: file.name, writable: file.access === 'readwrite' } : null,
      { sourceSize: file?.size ?? null, keepDraft: options.draft !== undefined },
    );
  }

  if (!file && !options.draft) {
    // גודל הדף לפני הכיווניות: `sections.setPageSetup` כותב את אותו `sectPr`
    // ש-`setSectionDirection` כותב אליו, וכך הכיווניות היא זו שנכתבת אחרונה.
    // גם הסדר של ההודעות נגזר מזה — כשל כיווניות הוא החמור, והוא זה שיישאר
    // בשורת המצב אם שניהם נכשלו.
    //
    // מסמך ששוחזר מטיוטה מדלג על שניהם בכוונה: ההגדרות האלה כבר בתוכו — הוא
    // היה מסמך פתוח שיוצא — והחלה חוזרת שלהן היא כתיבה למסמך של המשתמש
    // ברגע שהוא רק ביקש לחזור אליו.
    await applyNewDocumentPaperSize(editor.superdoc);
    await applyNewDocumentDirection(editor.superdoc);
  }

  if (options.draft) {
    // מה שנפתח אינו מה שבדיסק. בלי הסימון הזה „שמור” היה חושב שאין מה לשמור,
    // והפס העליון היה מציג „נשמר” על עבודה שאינה שמורה בשום מקום.
    save?.markDirty();
  }

  // האזנה למצב Undo/Redo
  editor.onDispose(
    adapter.observe('undo', (state) => {
      canUndo.value = state.enabled;
    })
  );
  editor.onDispose(
    adapter.observe('redo', (state) => {
      canRedo.value = state.enabled;
    })
  );

  // אחרון, ובכוונה: מסמך חדש עובר כאן דרך גודל דף וכיווניות, ומיקוד שקודם
  // להם היה מקבל סמן ואז פריסה שזזה תחתיו.
  focusOpenedDocument(editor.superdoc);

  // ואחריו המקום שהמשתמש היה בו. **אחרי** המיקוד ולא לפניו: `focus` עם
  // `restoreSelection` מציב סמן משלו, ומי שרץ אחרון הוא זה שקובע איפה הוא
  // יושב.
  await restoreDocumentView(editor, options.restore);

  return true;
}

/**
 * מחזירה למסמך שנפתח את גודל התצוגה ואת מקום הסמן שהיו בו בהפעלה הקודמת.
 *
 * הזום עובר דרך פקודת `zoom` של האדפטר ולא דרך `ui.zoom.set`, מאותו טעם
 * שמנוסח ב-engine/zoom.ts: יש מסלול כתיבה **אחד** לגודל התצוגה, וכל מי שמשנה
 * אותו — הסרגל, לחצני ±, וגם השחזור — עובר בו.
 *
 * כשל בשחזור אינו מגיע לשורת המצב: המשתמש ביקש לפתוח מסמך, לא לקפוץ למקום,
 * והודעת שגיאה על „לא מצאתי את השורה שהיית בה” היא רעש. הוא כן מגיע ללוג של
 * אוצריא, כי שם מודדים.
 */
async function restoreDocumentView(
  editor: EditorSession,
  restore: OpenOptions['restore'],
): Promise<void> {
  if (!restore) return;

  if (restore.zoom !== null) {
    const outcome = await commandAdapter.value?.run('zoom', zoomPayload(restore.zoom));
    if (outcome && !outcome.ok) {
      console.info(`[otzaria-word] גודל התצוגה השמור לא הוחזר: ${outcome.message}`);
    }
  }

  if (restore.caret && !(await applyCaretAnchor(editor.ui, editor.superdoc, restore.caret))) {
    console.info('[otzaria-word] מקום הסמן השמור לא נמצא במסמך שנפתח');
  }
}

async function onSave(forceSaveAs = false): Promise<void> {
  if (!swap?.current || !save) return;
  const outcome = await save.saveNow({ forceSaveAs, suggestedName: title.value });

  if (outcome.status === 'failed') {
    setStatus(outcome.message, true);
    return;
  }
  if (outcome.status === 'saved') {
    title.value = outcome.name.replace(/\.docx$/i, '') || title.value;
    // זוכר-ההפעלה אינו מעודכן כאן אלא ב-`onSaved` של הקואורדינטור: שם עוברות
    // גם השמירה האוטומטית וגם „לשמור לפני שפותחים אחר”, שאינן עוברות כאן.
    setStatus(`${title.value} נשמר`);
  }
}

async function onPickAndOpen(): Promise<void> {
  if (isOpening.value) return;
  try {
    const file = await pickDocxFile();
    if (!file) return;

    if (save && swap) {
      // נקרא **לפני** ההחלטה: אחרי „לשמור קודם” המסמך כבר נקי, ואז אי אפשר
      // להבחין בין „לא היה מה למחוק” לבין „המשתמש ביקש למחוק”.
      const hadUnsaved = save.snapshot.isDirty;
      const decision = await decideDocumentSwitch({
        isDirty: () => save!.snapshot.isDirty,
        isSaving: () => save!.snapshot.isSaving,
        confirm,
        documentName: () => title.value,
      });

      // „החלף” על מסמך שהיו בו שינויים פירושו שהמשתמש אישר את מחיקתם
      // במפורש — שתי שאלות, ראו open-flow.ts. זה המסלול היחיד שבו הטיוטה
      // נמחקת מלבד שמירה מוצלחת.
      if (hadUnsaved && decision.action === 'switch') await keeper?.discardDraft();

      if (decision.action === 'cancel') {
        setStatus(
          decision.reason === 'saving'
            ? 'השמירה עוד רצה — רגע אחד'
            : 'הפתיחה בוטלה, והמסמך נשאר פתוח'
        );
        return;
      }

      if (decision.action === 'save-first') {
        const outcome = await save.saveNow({ suggestedName: title.value });
        if (outcome.status !== 'saved') {
          if (outcome.status === 'failed') setStatus(outcome.message, true);
          else setStatus('הפתיחה נעצרה — המסמך לא נשמר');
          return;
        }
      }
    }

    await openDocument(file);
  } catch (error) {
    setStatus(error instanceof Error ? error.message : 'בחירת הקובץ נכשלה', true);
  }
}

async function onNewDocument(): Promise<void> {
  if (save && swap && save.snapshot.isDirty) {
    const decision = await decideDocumentSwitch({
      isDirty: () => save!.snapshot.isDirty,
      isSaving: () => save!.snapshot.isSaving,
      confirm,
      documentName: () => title.value,
    });
    // ראו onPickAndOpen: „החלף” כאן פירושו שהמחיקה אושרה במפורש.
    if (decision.action === 'switch') await keeper?.discardDraft();
    if (decision.action === 'cancel') return;
    if (decision.action === 'save-first') {
      const outcome = await save.saveNow({ suggestedName: title.value });
      if (outcome.status !== 'saved') return;
    }
  }
  await openDocument();
}

/**
 * „יציאה”.
 *
 * מה „יציאה” אומרת כאן, וזה אינו מובן מאליו: התוסף הוא לשונית בתוך אוצריא,
 * ולא אפליקציה שנסגרת. `navigation.goTo` מוציא את המשתמש מהמסך, ואוצריא
 * **משהה** את ה-WebView (`plugin.suspended`) במקום להרוס אותו — כלומר המסמך
 * ממתין כפי שהיה כשחוזרים. לכן היציאה אינה סוגרת את המסמך ואינה מוחקת דבר:
 * לסגור אותו היה מוחק עבודה שהמשתמש רק ביקש להתרחק ממנה, ודווקא בענף
 * „בלי לשמור” — שבו הוא אמר „אל תכתוב לדיסק”, ולא „תמחק לי את הטקסט”.
 *
 * מה שהכפתור כן קונה הוא השאלה: השמירה האוטומטית פועלת רק כשיש יעד כתיבה,
 * ומסמך חדש שטרם נשמר אין לו יעד — כלומר עד כאן הדרך היחידה לצאת ממנו הייתה
 * „פתח ספרייה” בלשונית „אוצריא”, שהוא כפתור ניווט ואינו שואל דבר. עכשיו יש
 * מסלול שמציע לשמור לפני שהולכים.
 *
 * ההחלטה עצמה היא `decideDocumentSwitch` עם `intent: 'exit'` — אותו קוד בדיוק
 * שמחליט על מעבר מסמך, כי „לצאת בלי לשמור” ו„לפתוח בלי לשמור” הם אותו סיכון.
 */
async function onExit(): Promise<void> {
  if (save && save.snapshot.isDirty) {
    const decision = await decideDocumentSwitch({
      isDirty: () => save!.snapshot.isDirty,
      isSaving: () => save!.snapshot.isSaving,
      confirm,
      documentName: () => title.value,
      intent: 'exit',
    });
    if (decision.action === 'cancel') {
      // בזמן שמירה אין לצאת: הסבב שרץ עוד לא כתב לדיסק. ההודעה היא זו של
      // מעבר מסמך, מאותו טעם ובאותו נוסח.
      if (decision.reason === 'saving') setStatus('השמירה עוד רצה — רגע אחד');
      return;
    }
    if (decision.action === 'save-first') {
      const outcome = await save.saveNow({ suggestedName: title.value });
      // שמירה שנכשלה או שבוטלה עוצרת את היציאה: המשתמש ביקש לשמור, וללכת
      // בכל זאת היה מתעלם ממה שביקש.
      if (outcome.status !== 'saved') {
        if (outcome.status === 'failed') setStatus(outcome.message, true);
        else setStatus('היציאה בוטלה — המסמך לא נשמר');
        return;
      }
    }
  }

  // אותו מסלול דיווח כמו „פתח ספרייה” בלשונית „אוצריא”: הודעה בעברית למשתמש
  // ושורה בלוג של אוצריא. כשל ניווט אינו מבטל את השמירה שכבר נעשתה.
  reportReader(await openLibrary());
}

async function onExportDocx(): Promise<void> {
  const active = swap?.current;
  if (!active) return;
  try {
    const blob = await exportDocx(active.superdoc);
    downloadBlob(blob, docxFileName(title.value));
    setStatus(`${title.value} יוצא ל-Word`);
  } catch (error) {
    setStatus(error instanceof Error ? error.message : 'הייצוא נכשל', true);
  }
}

/**
 * הדפסה. הכפתור קרא ל-`window.print()` בלבד, ולא היה בפרויקט אף `@media print`
 * — כלומר הוא הדפיס את הממשק (נמדד ב-CDP). הגלון ב-styles/print.css, וקביעת
 * `@page` לפי מידות הדף של המסמך ב-engine/print.ts; כאן רק הדיווח.
 *
 * גודל דף שלא נקרא אינו שגיאה: ההדפסה כן נפתחת, והמשתמש צריך לדעת שעליו לוודא
 * את גודל הנייר בדיאלוג. „הצלחה אינה מכריזה על עצמה” — התוצאה הנראית של
 * הדפסה היא דיאלוג ההדפסה עצמו.
 */
async function onPrint(): Promise<void> {
  if (!swap?.current) {
    setStatus('אין מסמך פתוח להדפסה', true);
    return;
  }

  const outcome = await printDocument(activeSuperdoc.value);
  if (!outcome.ok) {
    setStatus(outcome.message, true);
    return;
  }
  if (outcome.warning) {
    setStatus(outcome.warning);
    return;
  }
  if (isStatusError.value) setStatus('');
}

/**
 * ייצוא ל-PDF דרך `ui.exportPdf` של אוצריא (מ-0.9.97).
 *
 * ההכנה זהה להדפסה, ולא במקרה: אוצריא מייצרת את ה-PDF מדף התוסף עצמו, ולכן
 * הגלון של `styles/print.css` ומידות ה-`@page` הם שקובעים מה ייכנס לקובץ.
 * ההסבר המלא ב-engine/print.ts.
 *
 * ביטול בדיאלוג „שמור בשם” אינו שגיאה ואינו נכתב אדום: המשתמש נשאל ואמר לא,
 * וזו תשובה. אותו כלל כמו „אין התאמות” בהחלפה.
 */
async function onExportPdf(): Promise<void> {
  if (!swap?.current) {
    setStatus('אין מסמך פתוח לייצוא', true);
    return;
  }

  const outcome = await exportPdfDocument(
    activeSuperdoc.value,
    (input) => call('ui.exportPdf', input),
    { fileName: pdfSuggestedName(title.value), title: 'ייצוא ל-PDF' },
  );

  if (!outcome.ok) {
    setStatus(outcome.message, true);
    return;
  }
  if (!outcome.saved) {
    setStatus('הייצוא בוטל');
    return;
  }
  setStatus(outcome.warning ? `${outcome.name} נשמר — ${outcome.warning}` : `${outcome.name} נשמר`);
}

/**
 * „בטל”/„חזור” מפס הכותרת — דרך אותו מסלול כמו Ctrl+Z/Ctrl+Y.
 *
 * קודם עמד כאן `void commandAdapter.value?.run('undo')`, וה-`void` הוא הבאג:
 * `run` מחזיר `{ok, message, reason}`, וזריקת התוצאה פירושה שסירוב של המנוע
 * אינו מגיע לאף אחד. נמדד בשער המעטפת — לחיצה על „בטל”, הטקסט נשאר במסמך,
 * ולמשתמש לא נאמר דבר.
 *
 * מה שקרה שם: ברגע הלחיצה המנוע דיווח `{enabled: false, reason: 'not-ready'}`
 * בעוד הכפתור המצויר היה פעיל (החיווי מוחזק בכוונה, ראו engine/readout-hold.ts),
 * הפקודה נדחתה, וההודעה „המסמך עדיין נטען” — שכבר קיימת ב-`REASON_TEXT` —
 * נזרקה לפח. אותה פקודה בדיוק דרך המקלדת **כן** דיווחה, כי היא עוברת ב-
 * `runShortcutCommand`. שני מסלולים לאותה פעולה, ורק אחד מהם מדבר.
 *
 * `pageBreakTracker.forgetAllKeepingSnapshot`/`restoreSnapshot` כאן ולא רק
 * ב-`watchUndoRedoKeys`: לחיצה על הכפתורים כאן היא נתיב שני ל-Undo/Redo
 * שאינו עובר מקלדת בכלל — בלי הקריאה הישירה כאן, לחיצה על „בטל”/„חזור”
 * בפס הכותרת לא הייתה מנקה/משחזרת את המעקב אף פעם (רגרסיה שנחשפה תוך כדי
 * הוספת `watchUndoRedoKeys`: היא החליפה את הניקוי שישב קודם ב-
 * `runShortcutCommand`, ומעולם לא כיסתה את הכפתורים — הם אינם אירוע מקלדת).
 * בלי `isBlocked`: לחיצה מפורשת על כפתור תמיד עוסקת במסמך, לא במקום שהפוקוס
 * היה בו קודם. ראו ההסבר המלא ב-engine/page-break.ts, „QA עצמאי” → „Undo/Redo”.
 */
function onUndo(): void {
  pageBreakTracker.forgetAllKeepingSnapshot();
  void runShortcutCommand('undo');
}

function onRedo(): void {
  if (!pageBreakTracker.restoreSnapshot()) pageBreakTracker.forgetAll();
  void runShortcutCommand('redo');
}

function onTitleUpdate(newTitle: string): void {
  if (newTitle.trim()) {
    title.value = newTitle.trim();
    save?.markDirty();
    keeper?.noteChange();
  }
}

/**
 * המתג היה דקורטיבי: `autosaveEnabled` נכתב כאן ואיש לא קרא אותו, ו-
 * SaveCoordinator הריץ autosave על כל `markDirty` — כלומר כיבוי המתג לא כיבה
 * כלום. שתי השורות שנוספו הן מה שהופך אותו למתג: הבחירה מגיעה למי שמריץ את
 * ה-autosave, והיא שורדת הפעלות.
 */
function toggleAutosave(): void {
  autosaveEnabled.value = !autosaveEnabled.value;
  save?.setAutosaveEnabled(autosaveEnabled.value);
  void saveAutosaveEnabled(autosaveEnabled.value);
}

function toggleFocusMode(): void {
  isFocusMode.value = !isFocusMode.value;
  keeper?.updateView({ focusMode: isFocusMode.value });
  // יציאה ממצב מיקוד מאפסת את החשיפה: אחרת המחלקה נשארת והפסים מקבלים
  // opacity מיותר ברגע שחוזרים למצב הרגיל.
  if (!isFocusMode.value) revealed.value = null;
}

function onToggleBookCompletion(): void {
  bookCompletionEnabled.value = !bookCompletionEnabled.value;
}

/**
 * מתקינה/מפרקת את "השלמה מהספר" (engine/book-completion-overlay.ts) על
 * ה-container של המסמך הפתוח.
 *
 * `watch` על שלושתם ולא רק על הטוגל: מסמך שנפתח בזמן שהטוגל כבר דלוק צריך
 * גם הוא התקנה, ומסמך שנסגר (`activeEditorContainer` הופך `null`) צריך
 * פירוק — לא רק מעבר בין מסמכים, ולכן `documentGeneration` ולא `activeSuperdoc`
 * לבד: שני מסמכים עשויים לחלוק את אותו container לרגע (`editor-swap.ts`) והמונה
 * הוא הסימן החד-משמעי ל"מסמך אחר" (ראו doc על `documentGeneration`).
 */
let bookCompletion: ReturnType<typeof installBookCompletion> | null = null;
watch([activeEditorContainer, activeSuperdoc, bookCompletionEnabled, documentGeneration], () => {
  bookCompletion?.dispose();
  bookCompletion = null;
  if (!bookCompletionEnabled.value || !activeEditorContainer.value || !activeSuperdoc.value) return;
  bookCompletion = installBookCompletion(activeEditorContainer.value, activeSuperdoc.value, {
    onStatus: (message, isError) => setStatus(message, isError),
  });
});

/**
 * במצב מיקוד הפסים מוסתרים, ומתגלים כשהמצביע מתקרב לקצה. הקצה ולא כל המעטפת:
 * `:hover` על השורש החזיר את כולם בכל תנועה בחלון, כלומר המצב לא הסתיר כלום.
 * ההחלטה עצמה ב-composables/focus-mode.ts, כדי שתהיה נבדקת.
 */
function onPointerMove(event: PointerEvent): void {
  if (!isFocusMode.value) return;
  revealed.value = revealZone(event.clientY, window.innerHeight);
}

/**
 * הדיאלוג הוא שלנו ולא ה-surface המובנה של המנוע
 * (`modules: { surfaces: { findReplace: true } }`) — החלטה, לא שכחה: המנוע רץ
 * כאן ב-`ui: false`, הממשק כולו עברי ומימין לשמאל, ואילו ה-surface המובנה הוא
 * חלון באנגלית בעיצוב של SuperDoc שאין דרך ציבורית לתרגם או לעצב.
 *
 * הפעולות עצמן **אינן** עוברות דרך `ui.search` — הוא נמדד שאינו מכסה מסמך
 * רב-פסקאות (ראו הראש של engine/search.ts). המימוש שלנו עצמאי לגמרי:
 * `doc.blocks.list`/`doc.replace` של ה-Document API הציבורי.
 */
function openFindDialog(mode: 'find' | 'replace'): void {
  findMode.value = mode;
  isFindOpen.value = true;
  void reportSearch(searchAdapter?.open());
}

function closeFindDialog(): void {
  isFindOpen.value = false;
  // סגירה מנקה את ההדגשות במסמך. בלעדיה הן נשארות אחרי שהדיאלוג נעלם.
  searchAdapter?.close();
}

/**
 * התוצאה של כל פעולת חיפוש עוברת כאן: כשל לשורת המצב, הצלחה למונה.
 *
 * `async` כי `find()`/`open()` יכולים להגיע כ-Promise: קריאת `doc.blocks
 * .list()` מהמנוע היא א-סינכרונית (ראו engine/search.ts). `await` על ערך
 * שאינו Promise (כמו התוצאה הסינכרונית של `open()`) הוא no-op בטוח.
 */
async function reportSearch(outcome: SearchOutcome | Promise<SearchOutcome> | undefined): Promise<void> {
  const resolved = await outcome;
  if (!resolved) {
    setStatus('אין מסמך פתוח לחיפוש', true);
    return;
  }
  if (!resolved.ok) {
    setStatus(resolved.message, true);
    return;
  }
  searchState.value = resolved.snapshot;
}

function onFindText(query: string, direction: 'next' | 'prev'): void {
  void reportSearch(searchAdapter?.find(query, direction));
}

/** הקלדה בשדה החיפוש. ההשקטה עצמה באדפטר, כדי שתהיה נבדקת. */
function onFindQueryChange(query: string): void {
  searchAdapter?.findDebounced(query, reportSearch);
}

/**
 * החלפה היא capability gate ולא תכולה מובטחת: ב-superdoc@2.8.0 נמדד
 * ש-`replace`/`replaceAll` עשויים להחזיר `operation-unavailable`. לכן הכשל
 * מגיע לשורת המצב עם ההקשר שהוא כשל של החלפה — לא נבלע, ולא מתחפש להודעת
 * חיפוש.
 *
 * שני מצבים אינם כשל אלא תשובה, ולכן הם אינם אדומים ואינם נשלחים ללוג
 * השגיאות של אוצריא: „אין התאמות” ו„יש להזין טקסט לחיפוש”. שאילתה שלא נמצאה
 * היא מידע, ומי שכתב אותה אינו צריך התראת שגיאה עליה.
 */
const REPLACE_NOT_AN_ERROR = new Set(['no-matches', 'no-query']);

function reportReplace(outcome: SearchOutcome | undefined, success: string): void {
  if (!outcome) {
    setStatus('אין מסמך פתוח להחלפה', true);
    return;
  }
  if (!outcome.ok) {
    if (REPLACE_NOT_AN_ERROR.has(outcome.reason ?? '')) {
      setStatus(outcome.message);
      return;
    }
    setStatus(`ההחלפה לא בוצעה: ${outcome.message}`, true);
    return;
  }
  searchState.value = outcome.snapshot;
  setStatus(success);
}

async function onReplaceText(replacement: string): Promise<void> {
  reportReplace(await searchAdapter?.replace(replacement), 'המופע הוחלף');
}

async function onReplaceAllText(replacement: string): Promise<void> {
  // נקרא לפני הפעולה: אחריה קבוצת ההתאמות כבר התרוקנה.
  const matches = searchAdapter?.getState().total ?? 0;
  reportReplace(await searchAdapter?.replaceAll(replacement), `הוחלפו ${matches} מופעים`);
}

/**
 * חוזה הזום: `run('zoom', <אחוזים>)` — 100 הוא 100%, ו-`zoomPayload` בונה את
 * הצורה שהמנוע מקבל. `{ zoom: level / 100 }` שהיה כאן נדחה ב-
 * `instanceCommandPayloadIsValid` (הוא דורש `typeof payload === 'number'`
 * אחרי הנרמול) — התווית בשורת המצב התחדשה, והמסמך לא זז.
 *
 * הגבולות אינם קשיחים אלא `min`/`max` מ-`ui.zoom.getSnapshot()` דרך הנרמול
 * ב-engine/zoom.ts (כולל הרחבת התקרה להיקף Word — `ZOOM_PERCENT_MAX`),
 * וההגבלה נעשית ב-StatusBar לפי הגבולות האפקטיביים. הערך המוצג אינו נכתב כאן
 * אלא מגיע מ-`observeZoom`: כך התווית משקפת את מה שהמסמך באמת בו, גם כשהזום
 * השתנה ממקור אחר וגם כשהפקודה נדחתה.
 */
function onZoomChange(level: number): void {
  const payload = zoomPayload(level);
  if (payload === null) return;
  void commandAdapter.value?.run('zoom', payload);
}

/**
 * אחרי גרירה בסרגל שנכתבה למסמך.
 *
 * הקריאה המיידית ולא ההמתנה ל-`onUpdate`: זו מגיעה בהשקטה של חצי שנייה, ובזמן
 * הזה הידית הייתה קופצת בחזרה למקום הישן ואז שוב לחדש. הסמן אמור להישאר איפה
 * שהמשתמש עזב אותו.
 */
function onRulerChanged(): void {
  ruler?.refreshNow();
}

/**
 * דיווח לפקדי לשונית „אוצריא”.
 *
 * הצלחה אינה מכריזה על עצמה: התוצאה הנראית של „פתח ספרייה” ושל „חיפוש
 * באוצריא” היא מסך שמתחלף, והודעה שמתארת אותו היא בדיוק מה שהיה כאן קודם —
 * שלוש הודעות סטטוס („פותח חיפוש באוצריא...”) שתיארו פעולה שלא קרתה. מה
 * שההצלחה כן עושה הוא לנקות שגיאה קודמת שנשארה על המסך.
 */
function reportReader(outcome: ReaderResult<unknown>, success = ''): void {
  if (!outcome.ok) {
    setStatus(outcome.message, true);
    console.warn(`[otzaria-word] אוצריא: ${outcome.message} (${outcome.reason})`);
    return;
  }
  if (success || isStatusError.value) setStatus(success);
}

/**
 * ציטוט מהקורא: הבחירה בטאב הטקסט של אוצריא → מלל → הכנסה למסמך.
 *
 * „אין בחירה” אינו כשל אלא הוראה, ולכן `isError` כבוי: `reader.getSelection`
 * מחזיר `null` גם כשאין בחירה וגם כשהטאב הפעיל אינו טאב טקסט (PDF), ובשני
 * המקרים מה שהמשתמש צריך לשמוע זהה — לסמן קטע בספר.
 *
 * ההודעה על הצלחה אומרת **לאן** נכנס הציטוט: בלי סמן במסמך ה-Document API
 * מוסיף בסופו (זה החוזה), וזה בדיוק סוג הדבר שאין להשתיק.
 */
async function onInsertCitation(): Promise<void> {
  const selection = await getReaderSelection();
  if (!selection.ok) {
    reportReader(selection);
    return;
  }

  const text = buildCitationText(selection.value);
  if (!text) {
    setStatus('אין טקסט מסומן בקורא. סמנו קטע בספר הפתוח באוצריא, וחזרו לכאן');
    return;
  }

  const outcome = await insertCitation(activeSuperdoc.value, text);
  reportReader(
    outcome,
    outcome.ok && outcome.value === 'document-end'
      ? 'הציטוט נוסף בסוף המסמך — לא היה סמן במסמך'
      : 'הציטוט מאוצריא הוכנס במסמך',
  );
}

/**
 * השאילתה היא הטקסט המסומן במסמך — זה מה שהמשתמש רוצה לחפש כשהוא כותב חידוש
 * ומבקש את המקור. בלי בחירה אין שאילתה, ואוצריא דוחה `query` ריק; לכן ההודעה
 * מבקשת לסמן, ואינה שגיאה (`isError` כבוי — היא הוראה, לא כשל).
 */
async function onSearchOtzaria(): Promise<void> {
  const selection = await readDocSelection(activeSuperdoc.value, { includeText: true });
  const query = normalizeSelectedText(selection.text);
  if (!query) {
    setStatus('סמנו במסמך את הטקסט לחיפוש, ואז לחצו „חיפוש באוצריא”');
    return;
  }
  reportReader(await openSearchTab({ query }));
}

async function onOpenLibrary(): Promise<void> {
  reportReader(await openLibrary());
}

/**
 * קיצורי המקלדת. הרשימה עצמה ב-`ui/shortcuts/registry.ts`, ההכרעות (פוקוס,
 * דיאלוג פתוח, בליעת ברירת המחדל של הדפדפן) במנתב — וכאן נשארת ההרכבה בלבד.
 *
 * מה שהיה כאן קודם היה שרשרת `else if` שהשוותה `event.key` לאות. בפריסת מקלדת
 * עברית `Ctrl+S` מדווח `key: 'ד'`, ולכן כל הקיצורים מתו בדיוק כשהמשתמש עשה מה
 * שהתוסף נועד לו — כתב עברית. ההתאמה עברה ל-`event.code`, שאינו תלוי בפריסה.
 */
/**
 * תפריט הלחצן הימני. ההכרעות שלו — מה מוצג, איפה הוא נפתח, ומה קורה לסמן —
 * ב-composables/use-context-menu.ts; כאן נשארת ההרכבה.
 *
 * `runAction` נמסר כסגירה ולא כהפניה ישירה: `runShellAction` מוגדר מיד אחרי,
 * ואחת מהתלויות שלו (`openContextMenu`) היא של התפריט. שתי ההפניות נפתרות
 * בזמן ריצה, וכל סדר אחר היה מחייב לפצל אחת מהן לשני מקומות.
 */
const contextMenu = useContextMenu({
  superdoc: activeSuperdoc,
  shell: shellRef,
  isDocumentSurface,
  isFocusMode,
  isModalOpen: isModalDialogOpen,
  runAction: (action) => runShellAction(action),
  report: reportCommand,
});

/**
 * סגירת התפריט **והחזרת המיקוד למסמך**.
 *
 * הכרטיס מחזיק מיקוד אמיתי, ולכן סגירה שאינה מחזירה אותו משאירה את המיקוד על
 * `<body>`: הגלגלת סוגרת את התפריט, ומכאן ואילך ההקלדה אינה נכנסת לשום מקום
 * עד שהמשתמש לוחץ שוב במסמך. זו הסגירה הנפוצה ביותר, ולא מקרה קצה.
 */
function closeContextMenu(): void {
  contextMenu.close();
  focusDocument(activeSuperdoc.value);
}

const runShellAction = createShellActionRunner({
  isSaving: () => saveSnapshot.value.isSaving,
  save: (saveAs) => void onSave(saveAs),
  print: () => void onPrint(),
  openFind: (mode) => openFindDialog(mode),
  openLink: () => void linkDialog.open(),
  newDocument: () => void onNewDocument(),
  openDocument: () => void onPickAndOpen(),
  // שני אלה אינם פקודות של ה-controller אלא Document API ישיר, בדיוק כמו
  // הכפתורים המקבילים ברצועה — ולכן אותה פונקציה, ואותו דיווח.
  selectAll: () => void runSelectAll(),
  pageBreak: () => void runPageBreak(),
  growFont: () => void runFontStep(grownFontSize),
  shrinkFont: () => void runFontStep(shrunkFontSize),
  vertAlign: (kind) => void runVertAlign(kind),
  insertNote: (type) => void runInsertNote(type),
  toggleTrackChanges: () => void runToggleTrackChanges(),
  toggleFocusMode,
  findAgain,
  insertCitation: () => void onInsertCitation(),
  searchOtzaria: () => void onSearchOtzaria(),
  openLibrary: () => void onOpenLibrary(),
  toggleMacroRecording: onMacroRecord,
  replayLastMacro: onMacroPlay,
  toggleMacrosDialog: () => {
    if (isMacrosOpen.value) {
      isMacrosOpen.value = false;
      return true;
    }
    // אותה הכרעה כמו `toggleShortcutsHelp`: מעל דיאלוג אחר אין לפתוח שני.
    if (isModalDialogOpen()) return false;
    isMacrosOpen.value = true;
    return true;
  },
  openContextMenu: () => contextMenu.openAtCaret(),
  toggleShortcutsHelp: () => {
    if (isShortcutsHelpOpen.value) {
      isShortcutsHelpOpen.value = false;
      return true;
    }
    // מעל „אודות” או דיאלוג הקישור אין לפתוח חלון שני. הרשומה מסומנת
    // `inModal` כדי שתגיע לכאן בכלל — וההכרעה מי פתוח היא של המעטפת.
    if (isModalDialogOpen()) return false;
    isShortcutsHelpOpen.value = true;
    return true;
  },
  moveFocusRegion: (direction) => focusRing.move(direction) !== null,
  // „אודות” הוא `aria-modal`, ולכן הוא זה שנסגר כשהוא פתוח. החיפוש אינו מודאלי
  // ואפשר להמשיך לערוך מתחתיו, ולכן הוא נסגר רק כשאין חלון מעליו.
  closeTopmost: () => {
    if (isShortcutsHelpOpen.value) {
      isShortcutsHelpOpen.value = false;
      return true;
    }
    if (isMacrosOpen.value) {
      isMacrosOpen.value = false;
      return true;
    }
    if (isAboutOpen.value) {
      isAboutOpen.value = false;
      return true;
    }
    if (linkDialog.isOpen.value) {
      linkDialog.close();
      return true;
    }
    // התפריט **אחרי** החלונות המודאליים ולא לפניהם: הוא אמנם השכבה העליונה,
    // אבל מודאל פתוח חוסם את פתיחתו מלכתחילה — ולכן „תפריט פתוח מעל מודאל”
    // הוא מצב שלא אמור להתקיים. ענף ראשון היה מסכן בדיוק את מה שאינו אמור
    // לקרות: `Escape` שסוגר תפריט תקוע במקום את הדיאלוג שהמשתמש רואה.
    if (contextMenu.isOpen.value) {
      closeContextMenu();
      return true;
    }
    if (isFindOpen.value) {
      closeFindDialog();
      return true;
    }
    // מצב מיקוד הוא „חלון” גם הוא: הוא מסתיר את הרצועה ואת שורת המצב, ו-
    // `Escape` הוא המקש הראשון שכל משתמש מנסה כדי לצאת ממנו. בלי הענף הזה
    // היציאה היחידה הייתה למצוא שוב את F11 או לרחף מעל קצה המסך.
    if (isFocusMode.value) {
      toggleFocusMode();
      return true;
    }
    // אין מה לסגור: `Escape` מאחד מפסי המעטפת מחזיר את הפוקוס למסמך.
    // כשהוא כבר שם — `false`, והאירוע ממשיך למנוע ולדפדפן.
    return focusRing.toDocument();
  },
});

/**
 * דיאלוג הקישור. הוא יושב במעטפת ולא בלשונית „הוספה” מפני שלשונית שאינה
 * פעילה אינה מורכבת — ו-`Ctrl+K` חייב לעבוד מכל לשונית.
 */
const linkDialog = createLinkDialog({
  // אותו תנאי בדיוק שמנטרל את הכפתור ברצועה (`linkCmd.enabled`), ולא
  // „יש מסמך”. שני תנאים שונים לאותה פעולה פירושם ש-Ctrl+K פותח דיאלוג
  // שהאישור בו ייכשל — בעברית, אבל רק אחרי שהמשתמש כבר הקליד כתובת.
  canOpen: () => commandAdapter.value?.getState('link').enabled === true,
  readSelection: () => readDocSelection(activeSuperdoc.value, { includeText: true }),
  runLink: (payload) => void runShortcutCommand('link', payload),
  report: reportCommand,
});

/**
 * „הגדל/הקטן גופן”. הגודל נקרא **מהמנוע** ולא נשמר אצלנו: מונה מקומי היה
 * מטפס גם כשהמסמך דוחה את הפקודה, ואז הלחיצה הבאה הייתה מחשבת מגודל שאינו
 * במסמך. אותו כלל בדיוק שהכפתורים ברצועה עובדים לפיו.
 */
async function runFontStep(step: (current: number) => number): Promise<void> {
  const adapter = commandAdapter.value;
  if (!adapter) {
    setStatus('המסמך עדיין נטען', true);
    return;
  }

  const current = parseFontSizePt(adapter.getState('font-size').value) ?? DEFAULT_FONT_SIZE_PT;
  const payload = fontSizePayload(step(current));
  if (payload === null) return;

  reportCommand(await adapter.run('font-size', payload), 'font-size');
}

async function runInsertNote(type: 'footnote' | 'endnote'): Promise<void> {
  reportCommand(await insertNote(activeSuperdoc.value, type), `footnotes-insert-${type}`);
}

/**
 * מעקב שינויים. אין פקודה נפרדת: `document-mode` עם `'suggesting'` **הוא**
 * מצב המעקב, ולכן המצב הנוכחי נקרא מהמנוע — בדיוק כמו שהמתג ב„סקירה” עושה.
 */
async function runToggleTrackChanges(): Promise<void> {
  const adapter = commandAdapter.value;
  if (!adapter) {
    setStatus('המסמך עדיין נטען', true);
    return;
  }

  const suggesting = adapter.getState('document-mode').value === 'suggesting';
  const payload = { mode: suggesting ? 'editing' : 'suggesting' };
  reportCommand(await adapter.run('document-mode', payload), 'document-mode');
}

/**
 * „המופע הבא/הקודם”. בלי דיאלוג פתוח או בלי שאילתה — פותחים את החיפוש, כי
 * `F3` על מסמך שלא חיפשו בו הוא בקשה להתחיל לחפש ולא כשל.
 */
function findAgain(direction: 'next' | 'prev'): boolean {
  const query = searchState.value.query;
  if (!isFindOpen.value || query === '') {
    openFindDialog('find');
    return true;
  }

  void reportSearch(searchAdapter?.find(query, direction));
  return true;
}

async function runVertAlign(kind: 'superscript' | 'subscript'): Promise<void> {
  reportCommand(await toggleVertAlign(activeSuperdoc.value, kind), `vert-align-${kind}`);
}

async function runSelectAll(): Promise<void> {
  reportCommand(await selectWholeDocument(activeSuperdoc.value), 'select-all');
}

async function runPageBreak(): Promise<void> {
  reportCommand(await startParagraphOnNewPage(activeSuperdoc.value), 'page-break-before');
}

/**
 * פקודת מנוע שמגיעה מקיצור. אותו מסלול, ואותו דיווח, כמו לחיצת כפתור.
 *
 * **לא המקום שתופס Undo/Redo מהמקלדת** — זו הייתה ההנחה הראשונה וההיא הופרכה
 * במדידה: `createShortcutDispatcher` (dispatch.ts) מדלג בכוונה על אירוע
 * שכבר `defaultPrevented`, כדי לא להריץ קיצור שהמנוע כבר קשר בעצמו (Ctrl+B
 * וכדומה) פעמיים. `Ctrl+Z`/`Ctrl+Y` הם בדיוק המקרה הזה — הם ה-`history`
 * המובנה של ProseMirror, קשורים על אזור המסמך, ומבטלים את ברירת המחדל לפני
 * שהאירוע מגיע לכאן בכלל. נמדד: `runShortcutCommand('undo')` **לא רץ** על
 * Ctrl+Z אמיתי כשהפוקוס בתוך המסמך, גם שה-DOCX השתנה בפועל. `watchUndoRedoKeys`
 * (ui/shortcuts/undo-redo-watch.ts) הוא הפתרון — מאזין נפרד ב-capture, לפני
 * המנוע. ראו engine/page-break.ts, „QA עצמאי” → „Undo/Redo”.
 */
async function runShortcutCommand(id: CommandId, payload?: unknown): Promise<void> {
  const adapter = commandAdapter.value;
  if (!adapter) {
    setStatus('המסמך עדיין נטען', true);
    return;
  }
  reportCommand(await adapter.run(id, payload), id);
}

/**
 * האם היעד יושב בתוך אזור המסמך.
 *
 * זה מה שמפריד בין „שדה טקסט שלנו” ל„משטח ההקלדה של המנוע”. המנוע
 * מקבל הקשות דרך `<textarea>` נסתר ברוחב פיקסל אחד (נמדד בדפדפן: הוא
 * בנוי ל-IME), ולכן מרגע שהמשתמש מתחיל להקליד `event.target` של כל הקשה
 * הוא TEXTAREA. בלי ההצלבה הזאת כל הקיצורים נחסמים בדיוק כשהפוקוס במסמך —
 * כלומר במצב היחיד שבו הם נחוצים.
 *
 * הבדיקה היא הכלה באלמנט שאנחנו מחזיקים, ולא שאילתה על ה-DOM הפנימי של
 * המנוע — אותו גבול ש-tests/unit/engine-boundaries.test.ts שומר עליו.
 */
function isDocumentSurface(target: EventTarget | null): boolean {
  const host = editorStackRef.value;
  return host !== null && target instanceof Node && host.contains(target);
}

/**
 * דיאלוג שמכריז `aria-modal`. מה שמאחוריו אינו זמין — גם לא לקיצור.
 *
 * שלושה הרפים שהמעטפת מחזיקה בעצמה, ובנוסף — שאילתת DOM: `aria-modal="true"`
 * מוצהר גם ב-17 פאנלים שחיים בתוך לשוניות הרצועה עם מצב מקומי (פסקה, גופן,
 * הערה וכל השאר), וללא השאילתה קיצורי מקלדת וניווט חצים ממשיכים לפעול מתחת
 * להם — בדיוק ההצהרה שה-`aria-modal` שלהם מכחישה.
 */
function isModalDialogOpen(): boolean {
  if (isAboutOpen.value || linkDialog.isOpen.value || isShortcutsHelpOpen.value) return true;
  return document.querySelector('[aria-modal="true"]') !== null;
}

/**
 * הפוקוס אינו בעריכת המסמך: דיאלוג מודאלי פתוח, או שדה טקסט **של הממשק
 * שלנו** (לא של המסמך — `isDocumentSurface` היא ההצלבה). משמשת כל מי שצריך
 * לדעת שלחיצת מקלדת אינה אמורה לגעת במסמך: `createDirectionShortcut` (כיוון
 * פסקה) ו-`watchUndoRedoKeys` (Undo/Redo) — לשניהם יש מנגנון נפרד שאינו עובר
 * דרך `createShortcutDispatcher` הרגיל, ולכן אף אחד מהם לא מקבל את הבדיקה
 * הזאת בחינם ממנו.
 */
function isOutsideDocumentEditing(target: EventTarget | null): boolean {
  return isModalDialogOpen() || (isTextEntryTarget(target) && !isDocumentSurface(target));
}

/**
 * מיקוד המסמך ברגע שנפתח, כדי שאפשר יהיה להקליד בלי קליק מקדים.
 *
 * בלי זה כל פתיחה — בעלייה, ב„מסמך חדש” וב„פתח קובץ” — מגיעה בלי סמן: העורך
 * מוצג, המקלדת אינה שייכת לאיש, והמשתמש חייב ללחוץ עם העכבר בגוף הטקסט לפני
 * שיוכל לכתוב מילה.
 *
 * דרך המנוע ולא דרך ה-`<main>` שמארח אותו: מיקוד המארח מזיז את הפוקוס אבל
 * אינו מחזיר את הסמן לטקסט (ראו `engine/focus.ts`).
 *
 * שני שערים, ומאותו טעם: הפתיחה אסינכרונית ויכולה להימשך שניות, ובזמן הזה
 * המשתמש כבר עלול להיות במקום אחר.
 *
 *   * דיאלוג מודאלי פתוח — מה שמאחוריו אינו זמין, וחטיפת הפוקוס ממנו שוברת
 *     את מלכודת המיקוד שלו.
 *   * הפוקוס בשדה טקסט של הממשק (שורת החיפוש אינה מודאלית ונשארת פתוחה מעל
 *     המסמך) — שם המשתמש מקליד עכשיו, וקפיצה לגוף המסמך הייתה קוטעת אותו.
 */
function focusOpenedDocument(superdoc: SuperDoc): void {
  if (isModalDialogOpen()) return;

  const active = document.activeElement;
  if (isTextEntryTarget(active) && !isDocumentSurface(active)) return;

  focusDocument(superdoc);
}

/**
 * מעגל המיקוד של `F6`, בסדר של המסך: סרגל הכותרת, הרצועה, המסמך, שורת המצב.
 *
 * אזור המסמך ממוקד דרך המנוע ולא דרך ה-`<main>` שמארח אותו: מיקוד המארח מזיז
 * את הפוקוס אבל אינו מחזיר את הסמן לטקסט, כלומר המשתמש היה מקבל „חזרה למסמך”
 * שאי אפשר להקליד אחריה.
 *
 * שלושת פסי המעטפת מסומנים כלא-זמינים במצב מיקוד. הם עדיין בעץ — ההסתרה היא
 * `opacity: 0` — ולכן בלי הסימון `F6` היה ממקד פקד בלתי נראה. זה גם מה שנותן
 * לשדה שם המסמך דרך יציאה: `Escape` ממנו מזהה שהפוקוס בסרגל הכותרת ומחזיר
 * אותו למסמך.
 */
function shellRegion(selector: string): HTMLElement | null {
  return shellRef.value?.querySelector<HTMLElement>(selector) ?? null;
}

const outsideFocusMode = () => !isFocusMode.value;

const focusRing = createFocusRing({
  regions: [
    {
      id: 'titlebar',
      element: () => shellRegion('.word-titlebar'),
      isAvailable: outsideFocusMode,
    },
    {
      id: 'ribbon',
      element: () => shellRegion('.word-ribbon-container'),
      isAvailable: outsideFocusMode,
    },
    {
      id: 'document',
      element: () => editorStackRef.value,
      focus: () => focusDocument(activeSuperdoc.value),
    },
    {
      id: 'statusbar',
      element: () => shellRegion('.word-statusbar'),
      isAvailable: outsideFocusMode,
    },
  ],
});

let shortcuts: ShortcutDispatcher | null = null;
let directionShortcut: { dispose: () => void } | null = null;
let undoRedoWatcher: UndoRedoWatcher | null = null;

onMounted(async () => {
  shortcuts = createShortcutDispatcher({
    runCommand: (id, payload) => void runShortcutCommand(id, payload),
    runAction: runShellAction,
    // מודאלי = `aria-modal`, וזה מה שקובע. „אודות” ודיאלוג הקישור מכריזים
    // כך; דיאלוג החיפוש אינו מודאלי בכוונה, ומעליו עדיין מותר לערוך ולשמור.
    isModalOpen: () => isModalDialogOpen(),
    isDocumentSurface,
  });

  directionShortcut = createDirectionShortcut({
    runCommand: (id) => void runShortcutCommand(id),
    isBlocked: isOutsideDocumentEditing,
  });

  // Undo/Redo יכולים לשנות pageBreakBefore בלי לעבור דרך הכפתור ב-InsertTab.vue,
  // וגם בלי לעבור דרך `shortcuts` שמעל: ה-capture הנפרד כאן קיים בדיוק בגלל
  // זה — ראו ui/shortcuts/undo-redo-watch.ts. תצוגת „לא ידוע” (נופלת ל„כבוי”)
  // ב-InsertTab.vue עדיפה על „פעיל” כוזב שנשאר תקוע. `isBlocked` היא אותה
  // בדיקה בדיוק שלמעלה — QA מדד ש-Ctrl+Z בתוך שדה טקסט של הממשק (חיפוש,
  // למשל) ניקה את המעקב בלי שום קשר למסמך לפני שהיא נוספה.
  undoRedoWatcher = watchUndoRedoKeys({
    onUndo: () => pageBreakTracker.forgetAllKeepingSnapshot(),
    onRedo: () => {
      if (!pageBreakTracker.restoreSnapshot()) pageBreakTracker.forgetAll();
    },
    isBlocked: isOutsideDocumentEditing,
  });

  if (editorStackRef.value) {
    // לפני פתיחת המסמך הראשון: `observeZoom` יורה מיד עם ה-snapshot, וללא
    // הפקד הזה הדיווח הראשון היה הולך לאיבוד.
    zoomCenter = createZoomCenter(editorStackRef.value);

    save = initSaveCoordinator();
    keeper = initSessionKeeper();

    // הבחירה נטענת לפני שנפתח מסמך: העריכה הראשונה עלולה להתחיל סבב autosave,
    // ואם ההעדפה עוד לא הגיעה הוא היה רץ לפי ברירת המחדל ולא לפי מה שהמשתמש
    // בחר בהפעלה הקודמת.
    autosaveEnabled.value = await loadAutosaveEnabled();
    save.setAutosaveEnabled(autosaveEnabled.value);

    // גם ההעדפה של הסרגל, ומאותו טעם: היא חלה על המסמך שנפתח מיד אחרי כאן.
    rulerPreference = await loadRulerVisible();

    // ורשומת ההפעלה, מאותו טעם בדיוק: הפתיחה עצמה כותבת עליה (`setDocument`),
    // וקריאה אחריה הייתה קוראת את מה שהרגע דרסנו.
    const session = await loadPreviousSession();
    keeper.adopt(session);
    applyShellPreferences(session);

    // ההאזנה למעבר לרקע נרשמת כאן, אחרי שיש למי לדווח. שלושת המקורות
    // וההנמקה — ב-host/lifecycle.ts.
    hiddenListener = onPluginHidden(() => void keeper?.flush());

    swap = createEditorSwap(editorStackRef.value, (host, source) =>
      createEditor({
        container: host,
        source,
        onError: (err) => console.error('[otzaria-word] שגיאת מנוע:', err),
        onUpdate: () => {
          save?.markDirty();
          metrics?.noteDocumentChanged();
          // שוליים או כניסות יכולים להשתנות גם מפעולה ברצועה (גלריית
          // „שוליים”, דיאלוג הפסקה) ולא רק מהסרגל עצמו.
          ruler?.noteDocumentChanged();
          // וגם „גבולות עמוד” — אותה תחנה בדיוק, כדי שהתפריט ברצועה יעדכן
          // את השכבה בלי רענון.
          pageBorderModel?.noteDocumentChanged();
          // וגם „מספרי שורות” — עריכת טקסט משנה את מספר השורות ואת מיקומן.
          lineNumberModel?.noteDocumentChanged();
          // וגם „סימני עיצוב” — עריכת טקסט משנה את טקסט הבלוקים (ולכן את
          // מיקום ה-¶); ה-`setEnabled`-guard הפנימי שלה דואג שזה לא יקרא
          // כלום כשסימני העיצוב כבויים.
          formattingMarksModel?.noteDocumentChanged();
          // וגם זוכר-ההפעלה: זה הרגע שבו נולדת עבודה שאינה בדיסק.
          keeper?.noteChange();
        },
        // ה-callback נרשם פעם אחת, כאן, ולכן הוא מפנה למודד הנוכחי ולא
        // ל-session מסוים — בדיוק כמו `save?.markDirty()` שמעליו.
        onPaginationUpdate: (totalPages) => metrics?.notePaginationUpdate(totalPages),
      })
    );

    // טעינת מסמך אחרון או פתיחת מסמך ריק. תחנת „פותח את המסמך” מדווחת
    // מ-createEditor ולא מכאן: כאן היא הייתה מוקדמת בשנייה ויותר — הפתיחה
    // ממתינה קודם לחבילת המנוע, ומסך טעינה שאומר „פותח” בזמן שהוא מוריד
    // 9MB מתאר את השלב הלא נכון.
    try {
      await reopenPreviousSession(session);
    } finally {
      // גם פתיחה שנכשלה מסירה את מסך הטעינה, ולא רק כדי „לא להיתקע”: הודעת
      // הכשל יושבת בשורת המצב שמתחת, ומסך טעינה שנשאר פרוש מסתיר בדיוק את
      // מה שצריך להיקרא.
      splashDone();
    }
  } else {
    splashDone();
  }
});

onUnmounted(() => {
  bookCompletion?.dispose();
  bookCompletion = null;
  zoomCenter?.dispose();
  zoomCenter = null;
  shortcuts?.dispose();
  directionShortcut?.dispose();
  undoRedoWatcher?.dispose();
  undoRedoWatcher = null;
  // חיפוש-בזמן-הקלדה שממתין ירוץ אחרי הפירוק על handle של controller מפורק.
  searchAdapter?.dispose();
  hiddenListener?.();
  hiddenListener = null;
  keeper?.dispose();
  keeper = null;
});

/**
 * הרשומה של ההפעלה הקודמת, כולל המסלול ממשתמש שמעדכן מגרסה שלא הייתה בה
 * רשומה בכלל.
 *
 * `forgetLastDocument` אחרי ההמרה, ובכוונה: מרגע שהמפתח הישן נקרא, שני
 * מקורות לאותה שאלה הם מקור אחד יותר מדי — והישן הוא זה שכבר אינו מתעדכן.
 */
async function loadPreviousSession(): Promise<SessionState | null> {
  const stored = normalizeSession(await loadSessionRecord());
  if (stored) return stored;

  const migrated = sessionFromLastDocument(await loadLastDocument());
  if (migrated) void forgetLastDocument();
  return migrated;
}

/**
 * מצב המעטפת — מיקוד, לשונית, כיווץ — מוחל מיד ולא ממתין למסמך.
 *
 * זו העדפה של מי שיושב מול המסך ולא תכונה של המסמך (ראו `documentViewFor`),
 * ולכן היא נכונה גם אם המסמך האחרון לא ייפתח בכלל. החלה מוקדמת היא גם מה
 * שמונע הבהוב: הרצועה נפרסת פעם אחת בלשונית הנכונה, ולא קופצת אליה אחרי
 * שהמסמך נטען.
 */
function applyShellPreferences(session: SessionState | null): void {
  if (!session) return;
  isFocusMode.value = session.view.focusMode;
  if (session.view.ribbonTab) ribbonTab.value = session.view.ribbonTab;
  ribbonCollapsed.value = session.view.ribbonCollapsed;
}

/**
 * פותחת מחדש את מה שהיה — קובץ, טיוטה, או מסמך חדש.
 *
 * ## ארבעת המסלולים
 *
 * 1. **אין רשומה** — מסמך ריק, כמו תמיד.
 * 2. **יש קובץ, ואין טיוטה** — הקובץ נפתח מהדיסק, ועליו מוחזרים הזום והסמן.
 * 3. **יש טיוטה** — היא זו שנפתחת, כי היא מה שהיה על המסך. הקובץ עדיין הוא
 *    יעד השמירה, ולכן „שמור” יכתוב למקום הנכון.
 * 4. **ה-token לא נפתר** — הקובץ הוזז, נמחק, או שההרשאה בוטלה. נפתח מסמך
 *    חדש, והמשתמש מקבל הודעה במקום מסך ריק בלי הסבר. עבודה שלא נשמרה נפתחת
 *    לתוכו: היא אינה תלויה בקובץ שאבד.
 *
 * הטיוטה נבדקת גם כשאין קובץ כלל: מסמך חדש שמעולם לא נשמר הוא בדיוק המקרה
 * שבו אין שום דבר אחר לחזור אליו.
 */
async function reopenPreviousSession(session: SessionState | null): Promise<void> {
  const remembered = session?.document ?? null;
  const file = remembered ? await resolveRememberedFile(remembered) : null;

  if (remembered && !file) {
    // הקובץ אינו נגיש — הוזז, נמחק, או שההרשאה בוטלה. אבל עבודה שלא נשמרה
    // אינה תלויה בו: אין לה יעד כתיבה בכל מקרה („שמור” יפתח „שמור בשם”),
    // ולכן פתיחתה כמסמך חדש אינה יכולה לדרוס דבר — והיא הדרך היחידה שלא
    // לאבד אותה. הטיוטה עוברת לבעלות המסמך שנפתח ממנה (ראו `setDocument`).
    const orphan = session?.draft?.documentToken === remembered.token
      ? await readDraftBytes(session.draft.path)
      : null;

    await openDocument(undefined, { draft: orphan ?? undefined });
    setStatus(
      orphan
        ? `${remembered.name} לא נמצא — השינויים שלא נשמרו נפתחו כמסמך חדש`
        : `${remembered.name} לא נמצא — נפתח מסמך חדש`,
      !orphan,
    );
    return;
  }

  const draft = await recoverDraft(session, file);
  const restore = documentViewFor(session, file?.token ?? null);

  if (await openDocument(file ?? undefined, { draft: draft ?? undefined, restore })) {
    if (draft) setStatus(restoredDraftMessage(session));
    return;
  }

  // הפתיחה נכשלה. מסמך חדש עדיף על מסך ריק — אבל הרשומה **אינה** מתעדכנת
  // אליו: הכשל עשוי להיות זמני (worker שלא עלה, קובץ נעול), ואילו רישום
  // המסמך הריק היה מוחק את המסמך האחרון מהרשומה ומנתק אותה מהטיוטה שמחזיקה
  // את העבודה. בהפעלה הבאה מנסים שוב, בדיוק מאותה נקודה.
  await openDocument(undefined, { remember: false });
  setStatus('המסמך האחרון לא נפתח — נפתח מסמך חדש', true);
}

/**
 * ההודעה על שחזור, עם גיל הטיוטה כשהוא ידוע.
 *
 * הגיל אינו קישוט: הטיוטה נכתבת בקצב משלה ואינה בהכרח הרגע האחרון שלפני
 * הסגירה (ראו `draftAgeLabel`). המספר הוא מה שמאפשר למשתמש לדעת מיד אם
 * חסר לו משהו, במקום לגלות זאת מאוחר יותר.
 */
function restoredDraftMessage(session: SessionState | null): string {
  const base = 'שוחזרו שינויים שלא נשמרו מההפעלה הקודמת';
  const age = session?.draft ? draftAgeLabel(session.draft.savedAt, Date.now()) : null;
  return age ? `${base} (נשמרו ${age})` : base;
}

/** ה-URL העדכני של הקובץ שנזכר. `null` = אינו נגיש יותר. */
async function resolveRememberedFile(remembered: {
  token: string;
  name: string;
  writable: boolean;
}): Promise<UserFile | null> {
  const file = await resolveFileUrl(remembered.token);
  if (!file) return null;
  return {
    ...file,
    name: file.name || remembered.name,
    access: remembered.writable ? 'readwrite' : 'read',
  };
}

/**
 * הבייטים של הטיוטה, אם יש מה לשחזר וזה בטוח.
 *
 * ההחלטה עצמה ב-sessions/session-state.ts ולא כאן, מאותו טעם כמו
 * `decideDocumentSwitch`: היא קובעת אם עבודה של המשתמש נכתבת מעל קובץ, וקוד
 * כזה חייב להיבדק. מה שנשאר כאן הוא השאלה למשתמש במסלול היחיד שאין בו תשובה
 * נכונה אחת — הקובץ השתנה מבחוץ.
 */
async function recoverDraft(
  session: SessionState | null,
  file: UserFile | null,
): Promise<Blob | null> {
  const decision = decideDraftRecovery({
    draft: session?.draft ?? null,
    openingToken: file?.token ?? null,
    diskSize: file?.size ?? null,
  });

  if (decision.action === 'discard') {
    // טיוטה של מסמך אחר אינה נמחקת: היא עדיין העבודה של אותו מסמך, והוא
    // עשוי להיפתח שוב. מה שאינה — רלוונטית עכשיו.
    return null;
  }

  if (
    decision.action === 'ask' &&
    !(await confirm({
      title: `${file?.name ?? 'המסמך'} השתנה מחוץ לעורך`,
      content:
        'יש שינויים מההפעלה הקודמת שלא נשמרו, אבל הקובץ עצמו התעדכן בינתיים.' +
        ' לפתוח את השינויים שלא נשמרו? „לא” יפתח את הקובץ כפי שהוא בדיסק.',
    }))
  ) {
    return null;
  }

  return readDraftBytes(session?.draft?.path ?? DRAFT_PATH);
}

/** בייטי הטיוטה כ-Blob שאפשר למסור למנוע, או `null` כשאין. */
async function readDraftBytes(path: string): Promise<Blob | null> {
  const bytes = await readWorkspaceBytes(path);
  return bytes ? new Blob([bytes], { type: DOCX_MIME }) : null;
}
</script>

<style scoped>
.word-app-shell {
  display: flex;
  flex-direction: column;
  height: 100vh;
  width: 100vw;
  overflow: hidden;
  background: var(--color-surface);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  direction: rtl;
}

/* אזור המסמך: שורה של הסרגל האנכי וה-stack. `min-width: 0` על ה-stack הוא מה
   שמאפשר לו להצטמצם — פריט flex אינו יורד מתחת לרוחב התוכן שלו בלעדיו, ומיכל
   הגלילה של המנוע היה דוחף את הסרגל האנכי אל מחוץ למסך. */
.editor-area {
  display: flex;
  flex: 1 1 auto;
  min-height: 0;
  /* מסגרת ההתייחסות של PageBorderOverlay.vue ו-LineNumberOverlay.vue
     (`position: absolute; inset: 0`) — כך כל שכבה מכסה גם את .editor-stack
     וגם את הפינה של הסרגל האנכי, ואת עצמה היא ממקמת ביחס למלבן העמוד
     שנמדד, לא ביחס ל-CSS. */
  position: relative;
}

.editor-stack {
  position: relative;
  flex: 1 1 auto;
  min-width: 0;
  min-height: 0;
  background: var(--color-surface-container-highest);
  overflow: hidden;
}

/* שורת הסרגל האופקי */
.ruler-row {
  display: flex;
  flex-shrink: 0;
}

/* הפינה שבין שני הסרגלים. אותו רוחב כמו הסרגל האנכי, ואותו רקע — כך הפינה
   נראית כמו המשך שלהם ולא כמו חור. */
.ruler-corner {
  flex-shrink: 0;
  width: 22px;
  height: 22px;
  background: var(--color-surface-container-highest);
  border-block-end: 1px solid var(--color-outline-variant);
  border-inline-end: 1px solid var(--color-outline-variant);
}

/* מצב מיקוד. הסרגל מצטרף לפסים העליונים: הוא חלק מאותה קבוצה שנחשפת בקצה
   העליון, ומצב מיקוד שמשאיר סרגל מידות על המסך אינו מצב מיקוד. */
.word-app-shell.focus-mode :deep(.word-titlebar),
.word-app-shell.focus-mode :deep(.word-ribbon-container),
.word-app-shell.focus-mode :deep(.doc-ruler),
.word-app-shell.focus-mode :deep(.doc-vruler),
.word-app-shell.focus-mode .ruler-corner,
.word-app-shell.focus-mode :deep(.word-statusbar) {
  opacity: 0;
  pointer-events: none;
  transition: opacity 0.3s ease;
}

/* החשיפה לפי קצה, ולא `:hover` על השורש: השורש הוא כל החלון, ולכן כל תנועת
   עכבר החזירה את שלושת הפסים — ומצב המיקוד לא הסתיר כלום. */
.word-app-shell.focus-mode.reveal-top :deep(.word-titlebar),
.word-app-shell.focus-mode.reveal-top :deep(.word-ribbon-container),
.word-app-shell.focus-mode.reveal-top :deep(.doc-ruler),
.word-app-shell.focus-mode.reveal-top :deep(.doc-vruler),
.word-app-shell.focus-mode.reveal-top .ruler-corner,
.word-app-shell.focus-mode.reveal-bottom :deep(.word-statusbar) {
  opacity: 1;
  pointer-events: auto;
}
</style>
