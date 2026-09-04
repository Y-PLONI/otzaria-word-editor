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
    <!--
      הפסים העליונים כקבוצה אחת: כותרת, טאבי מסמכים, רצועה וסרגל אופקי.
      העטיפה אינה קוסמטית — היא מה שמאפשר למצב מיקוד להוציא את כולם
      מהזרימה במכה אחת ולהחזיר אותם כלוח צף אחד, בלי שהמסמך יזוז. ראו
      `.shell-top` ב-`<style>`.
    -->
    <div class="shell-top">
      <!-- פס עליון -->
      <TitleBar
        ref="titleBarRef"
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
        @open-find="(q) => openFindDialog('find', q)"
        @run-command="onRunCommandFromTellMe"
        @run-action="onRunActionFromTellMe"
        @custom-action="onCustomActionFromTellMe"
        @toggle-autosave="toggleAutosave"
        @update-title="onTitleUpdate"
      />

      <!-- רצועת טאבים — אחד ל-`DocumentSession` פתוח. ראו „ריבוי מסמכים” ליד `sessions` בסקריפט. -->
      <DocumentTabsBar
        :tabs="documentTabs"
        :active-id="documentIdView"
        @select-tab="onDocumentTabSelect"
        @close-tab="onDocumentTabClose"
        @new-tab="onDocumentTabNew"
      />

      <!-- רצועת הכלים (Ribbon) -->
      <Ribbon
        v-model:active-tab="ribbonTab"
        v-model:collapsed="ribbonCollapsed"
        :has-document="hasDocument"
        :has-pdf-export="supportsPdfExport"
        :is-saving="saveSnapshot.isSaving"
        :is-opening="isOpening"
        :is-exiting="isExiting"
        :book-completion-enabled="bookCompletionEnabled"
        @new-doc="openOpenDialog"
        @open-doc="openOpenDialog"
        @save-doc="onSave(false)"
        @save-as-doc="onSave(true)"
        @print-doc="onPrint"
        @export-pdf="onExportPdf"
        @export-otzaria="onExportOtzaria"
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
      <SpellingOverlay
        ref="spellingOverlayRef"
        :host="rulerHost"
        :viewport-source="rulerViewport"
        :dictionary="spellcheckDictionary"
        :revision="spellcheckRevision"
      />
    </div>

    <!--
      דרך יציאה שרואים. `Esc` ו-`F11` עובדים, אבל שניהם דורשים לדעת אותם —
      ובמצב מיקוד אין על המסך אף פקד שרומז עליהם. הכפתור יושב בפינה התחתונה
      של החלון — **מעל העמוד ולא לצדו**: העמוד מרוכז ורחב כמעט כמו החלון,
      ובחלון של 800 פיקסלים (וגם של 400) `elementFromPoint` מתחתיו מחזיר
      `.superdoc-page`. לכן הוא נשאר עמום עד ריחוף, וזה כל מה שמפריד אותו
      מהטקסט. המיקום, השכבה וההנמקה המלאה — ב-`.focus-exit` שב-`<style>`.
    -->
    <button
      v-if="isFocusMode"
      type="button"
      class="focus-exit"
      aria-label="יציאה ממצב מיקוד"
      data-tip-title="יציאה ממצב מיקוד"
      data-tip-shortcut="Esc"
      @pointerdown.prevent
      @click="toggleFocusMode"
    >
      <SvgIcon
        name="focusMode"
        :size="16"
      />
      <span>יציאה ממצב מיקוד</span>
    </button>

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
      :load="loadSnapshot"
      @update:zoom-level="onZoomChange"
      @toggle-focus="toggleFocusMode"
      @skip-load="onSkipLoad"
    />

    <!-- דיאלוגים ופאנלים -->
    <FindReplaceDialog
      :is-open="isFindOpen"
      :initial-mode="findMode"
      :initial-query="findInitialQuery"
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

    <!--
      „פתח מסמך”. תצוגה בלבד: הרשימה המלאה נכנסת כ-prop, והסינון והמיון
      לתצוגה נעשים בתוכו — ראו את ראש הקומפוננטה.
    -->
    <OpenDocumentDialog
      v-model:search-query="recentSearch"
      :is-open="isOpenDialogOpen"
      :templates="DOCUMENT_TEMPLATES"
      :recents="recentDocuments"
      :busy="isOpening || saveSnapshot.isSaving"
      :discarded-count="discardedBackups.length"
      @close="isOpenDialogOpen = false"
      @browse="onOpenDialogBrowse"
      @create-from-template="onOpenDialogCreate"
      @open-recent="onOpenDialogRecent"
      @toggle-pin="onOpenDialogTogglePin"
      @forget-recent="onOpenDialogForget"
      @show-discarded="onShowDiscarded"
    />

    <!--
      מסך השחזור של „לא לשמור”. נכנסים אליו מ„פתח מסמך” בלבד, ולכן הוא נפתח
      אחרי שזה נסגר — שני מודאלים פתוחים זה מעל זה הם שתי מלכודות מיקוד
      שמתחרות על אותו Tab.
    -->
    <DiscardedDocumentsDialog
      :is-open="isDiscardedOpen"
      :entries="discardedBackups"
      :busy="isDiscardedBusy"
      @close="isDiscardedOpen = false"
      @open="onOpenDiscarded"
      @forget="onForgetDiscarded"
    />

    <!--
      „המסמך לא נשמר”. הדיאלוג היחיד שתשובה בו מוחקת עבודה, ולכן הוא מודאלי
      אמיתי ולא לוח צף: `isModalDialogOpen` מזהה אותו לפי `aria-modal` וחוסמת
      בזמן שהוא פתוח את קיצורי המקלדת שמתחילים פתיחה או סגירה נוספת.
    -->
    <UnsavedChangesDialog
      :question="unsavedPrompt.question.value"
      @choose="unsavedPrompt.answer"
    />

    <MacrosDialog
      :is-open="isMacrosOpen"
      :handle="activeMacros"
      :document-vba="documentVba"
      @close="isMacrosOpen = false"
      @status="setStatus"
    />

    <!--
      הטולטיפ של כל התוכנה — מופע אחד, בסוף המעטפת. הוא מאזין במסירה על המסמך
      ולא נקשר לפקד מסוים, ולכן אין לו props: כל פקד שמצהיר על `data-tip-*`
      מקבל אותו. ההסבר המלא ב-ui/tooltip/TooltipLayer.vue.

      השורה הזאת אינה אופציונלית: `title` הוסר מכל התוכנה (הוא מה שצייר טולטיפ
      שני, אפור, מעל הכרטיס), ולכן בלי השכבה אין טולטיפ בכלל — לא נפילה לאחור
      למלבן של מערכת ההפעלה. השם הנגיש אינו תלוי בה (`aria-label`).
    -->
    <TooltipLayer />
  </div>
</template>

<script setup lang="ts">
import {
  ref,
  provide,
  nextTick,
  onMounted,
  onUnmounted,
  computed,
  shallowRef,
  shallowReactive,
  watch,
  watchEffect,
} from 'vue';
import TitleBar from './ui/shell/TitleBar.vue';
import DocumentTabsBar, { type DocumentTabItem } from './ui/shell/DocumentTabsBar.vue';
import Ribbon from './ui/ribbon/Ribbon.vue';
import StatusBar from './ui/shell/StatusBar.vue';
import DocumentRuler from './ui/shell/DocumentRuler.vue';
import VerticalRuler from './ui/shell/VerticalRuler.vue';
import PageBorderOverlay from './ui/shell/PageBorderOverlay.vue';
import LineNumberOverlay from './ui/shell/LineNumberOverlay.vue';
import PilcrowOverlay from './ui/shell/PilcrowOverlay.vue';
import SpellingOverlay from './ui/shell/SpellingOverlay.vue';
import FindReplaceDialog from './ui/panels/FindReplaceDialog.vue';
import AboutDialog from './ui/panels/AboutDialog.vue';
import LinkDialog from './ui/panels/LinkDialog.vue';
import ShortcutsDialog from './ui/panels/ShortcutsDialog.vue';
import OpenDocumentDialog from './ui/panels/OpenDocumentDialog.vue';
import UnsavedChangesDialog from './ui/panels/UnsavedChangesDialog.vue';
import DiscardedDocumentsDialog from './ui/panels/DiscardedDocumentsDialog.vue';
import TooltipLayer from './ui/tooltip/TooltipLayer.vue';
import { DOCUMENT_TEMPLATES, applyTemplate, type TemplateId } from './engine/templates';
import {
  normalizeRecents,
  rememberRecent,
  forgetRecent,
  setRecentPinned,
  sortedRecents,
  type RecentDocument,
} from './sessions/recent-documents';
import { createCommandAdapter, type CommandAdapter, type CommandOutcome } from './engine/command-adapter';
import type { CommandId } from './engine/capabilities';
import {
  COMMAND_ADAPTER,
  COMMAND_REPORTER,
  STATUS_NOTIFIER,
  DOCUMENT_GENERATION,
  FONT_MEMORY,
  FONT_OPTIONS,
  READOUT_SELECTION,
  SPELLCHECK,
  STYLE_GALLERY,
} from './composables/keys';
import { ACTIVE_SUPERDOC } from './engine/document-api';
import { readDocSelection } from './engine/doc-selection';
import type { Dictionary } from './engine/spellcheck';
import { loadTorahDictionary, rememberUserWord } from './engine/spellcheck-dictionary';
import {
  buildCitationText,
  getReaderSelection,
  insertCitation,
  normalizeSelectedText,
  openLibrary,
  openSearchTab,
  registerSendToDocumentItem,
  handleSendToDocument,
  takePendingContextMenuClicks,
  type SendToDocumentEvent,
  type ReaderResult,
} from './host/otzaria-reader';
import {
  fallbackStyleGallery,
  observeStyleGallery,
  type StyleGalleryState,
} from './engine/style-gallery';
import {
  composeFontOptions,
  fallbackFontOptions,
  observeFontSlice,
  type FontOptions,
  type FontsSliceLike,
} from './engine/font-options';
import {
  emptyInstalledFonts,
  loadInstalledFonts,
  type InstalledFontsSnapshot,
} from './engine/system-fonts';
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
import {
  createEditorSwap,
  HOST_CLASS,
  PENDING_CLASS,
  type EditorSwap,
  type OpenEditor,
} from './sessions/editor-swap';
import {
  applyPaneScroll,
  guardPaneScroll,
  readPaneScroll,
  repairPaneScroll,
  type PaneScroll,
} from './sessions/pane-scroll';
import {
  createDocumentSession,
  emptyUiSnapshot,
  type DocumentSession,
  type DocumentUiSnapshot,
} from './sessions/document-session';
import {
  createDocumentLoad,
  idleLoadSnapshot,
  LOAD_STAGES,
  type DocumentLoad,
  type LoadAttempt,
  type LoadSnapshot,
} from './sessions/document-load';
import { createSaveCoordinator, type SaveCoordinator, type SaveSnapshot } from './sessions/save-coordinator';
import { createEditor, type EditorSession } from './engine/create-editor';
import { ACTIVE_MACROS, installMacros, type MacrosHandle } from './engine/macros';
import { registerShulchanTools } from './engine/shulchan/tools-registration';
import MacrosDialog from './ui/panels/MacrosDialog.vue';
import { installBookCompletion } from './engine/book-completion-overlay';
import { installAtMention } from './engine/at-mention-overlay';
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
  verifyHebrewDocumentDefaults,
} from './engine/document-defaults';
import { blankDocumentSource } from './engine/blank-document-source';
import type { SuperDoc } from 'superdoc';
import {
  DOCX_MIME,
  documentFileName,
  exportDocx,
  resolveSaveExtension,
  retypeBlob,
  stripWordExtension,
  type WordExtension,
} from './engine/export';
import { NO_VBA, type DocumentVba } from './engine/vba-import';
import { exportPdfDocument, pdfSuggestedName, printDocument } from './engine/print';
import { buildOtzariaBook, otzariaBookFileName } from './engine/otzaria-book';
import {
  beginBinaryWrite,
  uploadBytes,
  abortBinaryWrite,
  commitUserFileWrite,
  pickDocxFile,
  resolveFileUrl,
  type UserFile,
  type WriteTicket,
} from './host/files';
import { decideDocumentSwitch, decidePendingTabClose } from './sessions/open-flow';
import { createUnsavedPrompt } from './composables/use-unsaved-prompt';
import {
  backupPathFor,
  forgetDiscard,
  nextBackupSlot,
  normalizeBackups,
  rememberDiscard,
  sortedBackups,
  type DiscardedDocument,
} from './sessions/discard-backup';
import { decideSleep, type SleepCandidate } from './sessions/sleep-policy';
import { call, confirm, isAvailable, notifyError, on, tryCall } from './host/otzaria-client';
import { supportsPdfExport } from './host/host-capabilities';
import { splashDone } from './host/splash';
import {
  loadLastDocument,
  forgetLastDocument,
  loadAutosaveEnabled,
  saveAutosaveEnabled,
  loadRulerVisible,
  saveRulerVisible,
  loadSpellcheckEnabled,
  saveSpellcheckEnabled,
  loadSessionRecord,
  saveSessionRecord,
  loadRecentDocuments,
  saveRecentDocuments,
  loadDiscardBackups,
  saveDiscardBackups,
} from './host/settings';
import {
  activeEntry,
  createDocumentSessionId,
  decideDraftRecovery,
  defaultView,
  documentViewFor,
  draftAgeLabel,
  draftPathFor,
  emptyDocumentEntry,
  normalizeSession,
  sessionForEntry,
  sessionFromLastDocument,
  SESSION_VERSION,
  type DocumentSessionId,
  type SessionDocumentEntry,
  type SessionState,
  type SessionView,
} from './sessions/session-state';
import { createSessionKeeper, type SessionKeeper } from './sessions/session-keeper';
import {
  applyCaretAnchor,
  applyDocumentStartCaret,
  hasTextCaret,
  readCaretAnchor,
  type CaretAnchor,
} from './engine/caret-anchor';
import { createTextCursorWatch } from './engine/text-cursor';
import {
  deleteWorkspaceEntry,
  readWorkspaceBytes,
  writeWorkspaceBytes,
} from './host/workspace';
import { onPluginHidden, onPluginShown } from './host/lifecycle';
import { revealZone, type RevealBounds, type RevealZone } from './composables/focus-mode';
import { enterFullscreen, exitFullscreen, isFullscreen, watchFullscreen } from './composables/window-fullscreen';
import SvgIcon from './ui/icons/SvgIcon.vue';
import { copySelection, cutSelection, pasteFromClipboard, selectWholeDocument } from './engine/clipboard';
import type { TellMeCustomAction } from './ui/shell/tell-me-actions';
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
import { createFontMemory } from './composables/use-font-controls';
import { createLinkDialog } from './composables/use-link-dialog';
import { createShellActionRunner } from './ui/shortcuts/actions';
import type { ShellAction } from './ui/shortcuts/registry';
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
 * הזיכרון של בוררי הגופן. מסופק מכאן ולא נוצר בפקד, מפני שאותם שני בוררים
 * מופיעים בשני מקומות — הרצועה ותפריט הלחצן הימני — וזיכרון פרטי לכל אחד מהם
 * פירושו שהם מציגים ערכים שונים באותו רגע. ראו composables/use-font-controls.ts.
 */
provide(FONT_MEMORY, createFontMemory());

/**
 * שני המקורות שמרכיבים את הבורר, ולמה הם נפרדים.
 *
 * המנוע מדווח על גופני **המסמך** בכל פתיחה; המנייה מדווחת על מה שמותקן
 * **במכונה**, פעם אחת, ונוחתת מתי שנוחתת. דחיפה ישירה משניהם ל-`fontOptions`
 * הייתה גורמת למי שנחת שני למחוק את מה שהראשון הביא — ולכן שניהם refs,
 * וההרכבה אחת. ראו engine/system-fonts.ts.
 */
const engineFontSlice = shallowRef<FontsSliceLike | null>(null);
const installedFonts = shallowRef<InstalledFontsSnapshot>(emptyInstalledFonts());
watchEffect(() => {
  fontOptions.value = composeFontOptions(engineFontSlice.value, installedFonts.value);
});

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
 * המאקרו של Word שכבר במסמך — לקריאה בלבד (engine/vba-import.ts).
 *
 * נקרא בשלב המקדים, שם הבייטים כבר בזיכרון, ומתאפס בכל פתיחה: זו תכונה של
 * **המסמך**, לא של ה-session של המאקרו, ולכן היא אינה יושבת ב-`activeMacros`
 * — שם מגיעים רק דברים שאפשר להריץ.
 */
const documentVba = shallowRef<DocumentVba>(NO_VBA);

/**
 * הסיומת שתחתה המסמך יישמר.
 *
 * נגזרת מסיומת המקור ומקיומו של חלק מאקרו בחבילה, ולא קבועה `docx`: קובץ עם
 * `vbaProject` שנשמר כ-`docx` הוא קובץ ש-Word מתלונן עליו. ראו
 * `resolveSaveExtension` ב-engine/export.ts.
 */
const saveExtension = shallowRef<WordExtension>('docx');

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
/** מונע כניסה שנייה לסגירה הא-סינכרונית של כל הטאבים. */
const isExiting = ref(false);
/**
 * מחוון הטעינה של שורת המצב. ההכרעות — התחנות, הזחילה, ומה „דלג” מבטל — ב-
 * sessions/document-load.ts; כאן נשארת ההרכבה בלבד.
 *
 * הפתיחה הראשונה אינה נראית כאן אלא במסך הטעינה שפרוש מעל הממשק כולו
 * (index.html), וזה בסדר: אין לו שורת מצב להציג בה, ואין בו מסמך פתוח שאפשר
 * לחזור אליו. המחוון הזה הוא של כל פתיחה שאחריה — קובץ שנבחר, מסמך חדש,
 * ומסמך שנפתח מהרשומה.
 */
const loadSnapshot = ref<LoadSnapshot>(idleLoadSnapshot());
const documentLoad: DocumentLoad = createDocumentLoad({
  onChange: (snapshot) => {
    loadSnapshot.value = snapshot;
  },
});
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
  updateShellView({ ribbonTab: tab, ribbonCollapsed: collapsed });
});

const isFindOpen = ref(false);
const findMode = ref<'find' | 'replace'>('find');
const findInitialQuery = ref('');
const titleBarRef = ref<InstanceType<typeof TitleBar> | null>(null);
const isAboutOpen = ref(false);
const isShortcutsHelpOpen = ref(false);

/**
 * „פתח מסמך” — הדיאלוג שהחליף את הקפיצה הישירה לבורר הקבצים.
 *
 * שלושת המצבים כאן שייכים למעטפת ולא לדיאלוג: הוא תצוגה בלבד (ראו את ראש
 * `OpenDocumentDialog.vue`), ואינו נוגע ב-storage. `recentDocuments` היא
 * הרשימה **המלאה**; הסינון והמיון לתצוגה נעשים בתוכו דרך אותן פונקציות
 * טהורות שנשמרות כאן — `sessions/recent-documents.ts`.
 */
const isOpenDialogOpen = ref(false);
const recentDocuments = ref<RecentDocument[]>([]);

/**
 * „מסמכים שנסגרו בלי לשמור” — חמשת הגיבויים האחרונים, ומסך השחזור שלהם.
 *
 * הרשימה נטענת פעם אחת בעלייה ומתעדכנת אחרי כל כתיבה
 * (`backupDiscardedDocument`), ולא נקראת מחדש בכל פתיחת דיאלוג: המונה מוצג
 * על כפתור בתוך „פתח מסמך”, וקריאת גשר בכל הצגה שלו היא סבב שלם על מידע
 * שאיש מלבדנו אינו כותב.
 */
const discardedBackups = ref<DiscardedDocument[]>([]);
const isDiscardedOpen = ref(false);
/** מונע שתי פתיחות של אותו עותק לפני שהקריאה מה-storage הסתיימה. */
const isDiscardedOpening = ref(false);
const isDiscardedBusy = computed(
  () => isOpening.value || saveSnapshot.value.isSaving || isDiscardedOpening.value,
);
const recentSearch = ref('');

/**
 * רושמת מסמך שנפתח או שנשמר בשם חדש לרשימת האחרונים.
 *
 * שני אתרי הקריאה הם בדיוק שני הרגעים שבהם ל-token יש משמעות חדשה: פתיחת
 * קובץ, ו„שמור בשם” שמחליף את הקובץ שמאחורי המסמך. שמירה רגילה אינה כאן —
 * היא אינה משנה זהות, והיא הייתה מקפיצה את המסמך לראש הרשימה בכל סבב
 * אוטומטי.
 *
 * הכתיבה ל-storage אינה ב-await אצל הקוראים: רשימת אחרונים שלא נשמרה אינה
 * סיבה לעכב פתיחת מסמך, ו-`saveRecentDocuments` בולעת כשל ממילא.
 */
function rememberRecentDocument(entry: {
  token: string;
  name: string;
  size: number;
  /**
   * חייב להירשם, ואי-אפשר לגזור אותו מחדש: `fs.resolveFileUrl` — המסלול
   * שפתיחה מהרשימה עוברת בו — אינו מחזיר `access`. ראו `RecentDocument`.
   */
  writable: boolean;
}): void {
  recentDocuments.value = rememberRecent(recentDocuments.value, { ...entry, openedAt: Date.now() });
  void saveRecentDocuments(recentDocuments.value);
}

/**
 * „המסמך לא נשמר” — שאלה אחת עם שלושה כפתורים, במקום שתי שאלות רצופות של
 * `ui.showConfirm` (שהוא דו-כפתורי). ההמתנה לתשובה ב-composables/use-unsaved-prompt.ts,
 * ההחלטה שנגזרת ממנה ב-sessions/open-flow.ts, והגיבוי שמאפשר שאלה אחת
 * ב-sessions/discard-backup.ts.
 */
const unsavedPrompt = createUnsavedPrompt();

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

/* ------------------------------------------------------------------ */
/* בדיקת איות תורנית                                                   */
/* ------------------------------------------------------------------ */

/**
 * המילון התורני, ל-ui/shell/SpellingOverlay.vue. `null` = הבדיקה כבויה, וזו
 * גם ברירת המחדל: המילון הוא נכס נפרד של 1.3MB שנמשך רק כשמדליקים
 * (engine/spellcheck-dictionary.ts).
 *
 * המילון אינו של הטאב אלא של התוסף — הוא נטען פעם אחת ומשותף לכל המסמכים,
 * ולכן הוא כאן ולא ב-`DocumentSession`.
 */
const spellcheckDictionary = shallowRef<Dictionary | null>(null);
const spellcheckBusy = ref(false);
/**
 * מונה שעולה אחרי עריכה. השכבה מודדת מחדש עליו — עריכה בתוך פסקה אינה מזיזה
 * שום מלבן עמוד, ולכן מעקב הגיאומטריה לבדו אינו יורה עליה.
 */
const spellcheckRevision = ref(0);
/** רק מה שנצרך מהשכבה — ראו `defineExpose` ב-SpellingOverlay.vue. */
const spellingOverlayRef = shallowRef<{ wordAt: (x: number, y: number) => string | null } | null>(null);

/**
 * השקטת המדידה החוזרת אחרי עריכה. ארוכה מזו של הסרגל (150) בכוונה: קו גלי
 * שמהבהב על כל הקשה מפריע לקריאה יותר משהוא עוזר, ו-Word עצמו מסמן מילה רק
 * אחרי שעזבו אותה.
 */
const SPELLCHECK_DEBOUNCE_MS = 400;
let spellcheckTimer: ReturnType<typeof setTimeout> | undefined;

function noteSpellcheckChanged(): void {
  if (!spellcheckDictionary.value) return;
  clearTimeout(spellcheckTimer);
  spellcheckTimer = setTimeout(() => {
    spellcheckRevision.value += 1;
  }, SPELLCHECK_DEBOUNCE_MS);
}

/**
 * הדלקה/כיבוי. ההדלקה הראשונה מושכת את המילון, וכשל בה מדווח למשתמש ומשאיר
 * את המתג כבוי — עורך שנופל בגלל בדיקת איות גרוע מעורך בלי בדיקת איות.
 *
 * ההעדפה נשמרת בשני הכיוונים, אבל **רק אחרי שההדלקה הצליחה**: משתמש שהמילון
 * לא נטען אצלו לא אמור לפגוש את אותו כשל בכל עלייה.
 */
async function toggleSpellcheck(): Promise<void> {
  if (spellcheckBusy.value) return;

  if (spellcheckDictionary.value) {
    spellcheckDictionary.value = null;
    clearTimeout(spellcheckTimer);
    await saveSpellcheckEnabled(false);
    return;
  }

  spellcheckBusy.value = true;
  try {
    const dictionary = await loadTorahDictionary();
    if (!dictionary) {
      setStatus('טעינת המילון התורני נכשלה — בדיקת האיות נשארה כבויה');
      return;
    }
    spellcheckDictionary.value = dictionary;
    setStatus(`בדיקת איות תורנית פעילה — ${dictionary.size.toLocaleString('he-IL')} ערכים`);
    await saveSpellcheckEnabled(true);
  } finally {
    spellcheckBusy.value = false;
  }
}

provide(SPELLCHECK, {
  enabled: computed(() => spellcheckDictionary.value !== null),
  busy: spellcheckBusy,
  toggle: () => void toggleSpellcheck(),
});

/**
 * „הוסף למילון” מתפריט ההקשר. המילה נשמרת ב-`storage` (רשימת המשתמש בלבד,
 * לא עותק של 102,465 הערכים), והסימון נמדד מחדש מיד — אחרת המילה שנוספה
 * נשארת מסומנת עד הגלילה הבאה.
 */
function addWordToDictionary(word: string): void {
  void (async () => {
    const dictionary = spellcheckDictionary.value;
    if (!(await rememberUserWord(dictionary, word))) return;
    spellcheckRevision.value += 1;
    setStatus(`„${word}” נוספה למילון`);
  })();
}
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

/**
 * כל הטאבים הפתוחים, לפי מזהה, ומי מהם פעיל.
 *
 * ## איך ריבוי המסמכים חי בתוך מודל מסמך-יחיד
 *
 * כל שאר הקובץ הזה — `swap`/`save`/`keeper`/`title`/`docMetrics`/`zoom`/וכו' —
 * ממשיך לייצג **את הטאב הפעיל בלבד**, בדיוק כמו לפני ריבוי המסמכים. `activateTab`
 * הוא מה שמגשר: לפני שהוא עוזב טאב הוא שומר את כל ה-refs האלה לתוך
 * `DocumentSession.ui` של הטאב היוצא (`stashActiveInto`), ואז טוען את אלה של
 * הטאב הנכנס (`restoreFromSession`). כך אף פונקציה אחרת בקובץ — `openDocumentInto`,
 * `onSave`, `reportCommand` וכל השאר — אינה צריכה לדעת שיש יותר מטאב אחד.
 *
 * הטאבים שברקע ממשיכים לרוץ באמת: ה-`pane` שלהם נשאר מורכב (`display: none`
 * בלבד), וה-`swap`/`save`/`keeper` שלהם הם אובייקטים עצמאיים לגמרי — שמירה
 * אוטומטית וכתיבת טיוטה ממשיכות לפעול עליהם גם כשהם אינם מוצגים.
 *
 * המקרה החריג היחיד: עדכון שמגיע **א-סינכרונית**, אחרי שהמסמך כבר ברקע —
 * עימוד שמתייצב בהשהיה, גופנים/סגנונות שנפתרים מאוחר. אלה עלולים לכתוב
 * ל-refs של הטאב הפעיל **האחר**. הפתרון: `guardIfActive` עוטף בדיוק את
 * הכתיבות האלה, ומדלג עליהן כשה-session שהן שייכות לו כבר אינו הפעיל. פעולות
 * שמגיעות מאינטראקציה ישירה (הקלדה, לחיצה) אינן זקוקות לכך — הן אינן יכולות
 * לקרות בטאב שאינו גלוי/ממוקד מלכתחילה.
 *
 * `shallowReactive` ולא `Map` רגיל: `documentTabs` (למטה) הוא `computed` שרץ
 * `Array.from(sessions.values()).map(...)` — כשהמפה ריקה (הרגע הראשון, לפני
 * `onMounted`) הלולאה לא רצה אף פעם, ואז ה-computed לא קורא שום `.value` ולא
 * נרשם לאף תלות. Vue "רדוד" עוקב אחרי הפעולות על המפה עצמה (`set`/`delete`/
 * גודל/איטרציה) גם כשהיא ריקה, ולכן ה-computed כן מתעדכן כש-`sessions.set`
 * הראשון קורה. "רדוד" ולא `reactive` רגיל: הערכים במפה (`DocumentSession`)
 * מחזיקים `HTMLElement`/`SuperDoc` אמיתיים שאסור לעטוף ב-Proxy.
 */
const sessions = shallowReactive(new Map<DocumentSessionId, DocumentSession>());
const activeSession = shallowRef<DocumentSession | null>(null);

/**
 * מה שמקדם את רצועת הטאבים כשמצב של טאב **ברקע** משתנה.
 *
 * `sessions` הוא `shallowReactive` בכוונה (ראו שם: הרשומות מחזיקות
 * `HTMLElement`/`SuperDoc` שאסור לעטוף ב-Proxy), ולכן כתיבה **לתוך**
 * `session.ui` אינה תלות ריאקטיבית של אף `computed`. לטאב הפעיל זה אינו
 * מורגש — `documentTabs` קורא עליו את הסינגלטונים (`title`, `saveSnapshot`)
 * שהם `ref` אמיתיים — אבל טאב ברקע נקרא מ-`session.ui`, ואיש לא מרנדר מחדש
 * כשהוא משתנה.
 *
 * זה נמדד ולא הונח: שמירה אוטומטית של טאב ברקע דיווחה `isDirty` דרך
 * `onStateChange` שלו, והנקודה על הטאב הופיעה רק במעבר הטאב הבא — כלומר
 * המשתמש רואה „נשמר” על מסמך שיש בו שינויים, עד שהוא במקרה יעבור.
 *
 * מונה ולא הפיכת `ui` ל-`reactive`: רצועת הטאבים קוראת ממנו בדיוק שני
 * שדות (שם ונקודה), והעטיפה הייתה נוגעת בשלושים אתרי כתיבה ובאובייקטים
 * שדווקא אסור לעטוף.
 */
const tabStripRevision = ref(0);

/** מסמנת שמה שרצועת הטאבים מציגה על טאב ברקע השתנה. */
function notifyTabStrip(): void {
  tabStripRevision.value += 1;
}

/**
 * כמה מסמכים מותר שיחזיקו מנוע חי בו-זמנית.
 *
 * ## למה יש תקרה בכלל
 *
 * כל מסמך פתוח הוא מופע SuperDoc מלא — חבילת מנוע, workers, ועותק מפורס של
 * המסמך. עשרה טאבים פתוחים בלי תקרה הם עשרה מופעים כאלה בזיכרון, גם אם
 * המשתמש נגע בשניים מהם בשעה האחרונה. מי שעובד על מכונה צנועה משלם על זה
 * בכל התוסף, לא רק בטאב שהוא שכח.
 *
 * ## למה דווקא שלושה
 *
 * שלושה הם „זה שאני עליו, וזה שאני מדלג אליו וחוזר, ועוד אחד” — התרחיש
 * שבגללו ריבוי המסמכים קיים (מצטטים ממאמר אחד לשני). כל טאב מעבר לזה כמעט
 * תמיד ממתין דקות, ופתיחתו מחדש עולה פחות מלהחזיק אותו. משתמש עם שלושה
 * טאבים או פחות — הרוב המוחלט — אינו פוגש את המנגנון הזה כלל.
 *
 * מה שקורה לטאב הרביעי מתואר ב-`sleepTab`, ובראשי תיבות: הוא הופך לבדיוק
 * אותו „ממתין לטעינה” של טאב ששוחזר מהפעלה קודמת — אותו מסלול, אותה רשומה,
 * אותה טיוטה — ונפתח מחדש כשחוזרים אליו.
 */
const MAX_LIVE_DOCUMENTS = 3;

/**
 * מזהי הטאבים לפי סדר השימוש, החדש בראש. זה מה שקובע מי נרדם ראשון, ולכן
 * הוא מתעדכן ב-`activateTab` — הרגע היחיד שבו „מתי השתמשו בו” משתנה.
 */
const recentTabs: DocumentSessionId[] = [];

/**
 * הטאבים שנסגרו בהפעלה הזאת, החדש בראש — מה ש-`Ctrl+Shift+T` פותח מחדש
 * (`reopenClosedTab`).
 *
 * **בזיכרון בלבד, ובכוונה.** מה ששורד הפעלה הוא רשימת „המסמכים האחרונים”
 * (`recentDocuments`), והיא כבר מוצגת בדיאלוג הפתיחה עם שם, גודל וזמן. שכבה
 * שנייה קבועה על אותם נתונים לא הייתה מוסיפה למשתמש דבר — ואילו „הטאב שסגרתי
 * בטעות לפני שנייה” הוא בדיוק שאלה של ההפעלה הנוכחית.
 *
 * רק טאב שיש לו קובץ נרשם: מסמך חדש שלא נשמר מעולם אין לו token, ואין ממה
 * לפתוח אותו מחדש. רישום שלו היה מייצר שורה שלחיצה עליה נכשלת תמיד.
 *
 * המזהה הוא ה-token ולא הנתיב, מאותו טעם שבגללו רשימת האחרונים בנויה עליו
 * (ראו sessions/recent-documents.ts): ה-URL של אוצריא תקף לריצה אחת בלבד.
 *
 * ## למה `writable` נשמר כאן, ולא נקרא מחדש בפתיחה
 *
 * `resolveFileUrl` מחזירה `{token, url, name, size}` — **בלי `access`**.
 * פתיחה דרכה נראית מוצלחת לגמרי ובשקט הופכת את המסמך לקריאה-בלבד: `save`
 * מאבד את יעד הכתיבה, שורת המצב מבטיחה „„שמור” יבקש מקום חדש”, והזוכר כותב
 * `writable: false` לרשומת ההפעלה — כלומר זה **נשמר**, ומלווה את המסמך גם
 * אחרי הפעלה מחדש. התרחיש: קובץ פתוח בכתיבה → `Ctrl+W` → `Ctrl+Shift+T` →
 * `Ctrl+S` פותח „שמור בשם”.
 *
 * לכן ההרשאה נלקחת ממה שהטאב **עצמו** ידע עליה (`SessionDocument.writable`)
 * ומורכבת בחזרה ב-`resolveRememberedFile`. זה בדיוק הדפוס של שחזור ההפעלה,
 * והוא היחיד שנמדד כעובד: אין שום עדות ש-`fs.resolveFileUrl` של אוצריא
 * מחזירה `access` בכלל — התיעוד תולה את השדה ב-`fs.pick`, לא בה.
 *
 * שאלת „רשומה ישנה בלי השדה” אינה קיימת כאן: המחסנית בזיכרון בלבד ונולדת
 * ריקה בכל הפעלה, ו-`SessionDocument.writable` אינו אופציונלי. ברשימת
 * „האחרונים”, ששורדת הפעלות, אותה שאלה כן קיימת ושם נבחרה נפילה סגורה.
 */
const closedTabs: { token: string; name: string; writable: boolean }[] = [];

/**
 * התקרה. עשרה כמו ברוב הדפדפנים; מעבר לזה מדובר בזיכרון ולא בפעולת „ביטול”,
 * וזה כבר תפקידה של רשימת האחרונים.
 */
const MAX_CLOSED_TABS = 10;

/** רושמת טאב שנסגר, אם יש לו קובץ לחזור אליו. ראו `closedTabs`. */
function rememberClosedTab(session: DocumentSession): void {
  const document = activeEntry(session.keeper.state)?.document ?? null;
  if (!document) return;

  // אותו קובץ פעמיים הוא שורה אחת, ולא שתי לחיצות שפותחות את אותו דבר.
  const at = closedTabs.findIndex((entry) => entry.token === document.token);
  if (at >= 0) closedTabs.splice(at, 1);

  closedTabs.unshift({
    token: document.token,
    name: sessionDisplayTitle(session),
    writable: document.writable,
  });
  closedTabs.splice(MAX_CLOSED_TABS);
}

/** סבב הרדמה אחד בכל רגע: הוא ממתין ל-`flush`, ושניים היו נכנסים זה בזה. */
let trimming = false;

/** ראו את ההסבר ליד `sessions`. עוטפת כתיבה שמקורה בטאב שעלול כבר להיות ברקע. */
function guardIfActive(session: DocumentSession, run: () => void): void {
  if (activeSession.value === session) run();
}

/**
 * מיכל הגלילה של הטאב — ה-host שה-swap יצר בתוך הפאנל שלו
 * (sessions/editor-swap.ts). `null` כשאין מסמך פתוח בטאב, למשל טאב שממתין
 * לטעינה או כזה שנרדם.
 */
function documentScrollHost(session: DocumentSession): HTMLElement | null {
  // `:not(PENDING_CLASS)` ולא ה-host הראשון: בזמן פתיחה יושבים בפאנל שניים —
  // הפעיל, והמועמד שעדיין נבנה (sessions/editor-swap.ts). מיקום גלילה שנקרא
  // מהמועמד הוא תמיד אפס, וכתיבה אליו נמחקת כשהוא מוסר.
  const host = session.pane.querySelector(`.${HOST_CLASS}:not(.${PENDING_CLASS})`);
  return host instanceof HTMLElement ? host : null;
}

/**
 * מפרק את השומר של הגלילה בטאב הפעיל. אחד בכל רגע — ראו `restorePaneScroll`.
 */
let paneScrollGuard: (() => void) | null = null;

/**
 * מחזירה לטאב שנכנס את מיקום הגלילה שנשמר בו.
 *
 * ## שלוש פעולות, ולא אחת
 *
 * **השמה, ועוד אחת בפריים הבא.** הפאנל בדיוק יצא מ-`display: none`, והדפדפן
 * מחשב את גובה התוכן שלו מחדש. השמה שקורית לפני שהחישוב הזה הסתיים נחתכת
 * לגובה שעדיין אינו נכון (`scrollTop` נצמד למקסימום האפשרי באותו רגע). שתיהן
 * אידמפוטנטיות — ראו `applyPaneScroll`.
 *
 * **ואחריהן שומר על אירוע הגלילה הראשון**, וזה מה שמתקן את הבאג שנשאר פתוח:
 * שתי ההשמות נמדדו כמצליחות (`scrollTop` הוא 720 בכל נקודות הזמן — מיד,
 * מיקרו-משימה, rAF, rAF שני ו-150ms), ואז **גלגלת אחת** החזירה אפס. הכתיבה
 * היא של המנוע ולא של הדפדפן, והיא קורית מתוך הגלגלת עצמה — כלומר אחרי כל
 * מה שאנחנו יכולים לעשות מכאן. ההנמקה המלאה והמדידה: `sessions/pane-scroll.ts`
 * ו-`docs/engine-gaps.md`.
 *
 * ## למה השומר נדרך בתוך ה-rAF ולא לפניו
 *
 * ההשמה הראשונה עלולה להיחתך (ראו למעלה), ואירוע הגלילה שהיא יורה היה נראה
 * לשומר בדיוק כמו „המשתמש גלל למקום אחר” — כלומר מכבה אותו לפני שהמנוע כתב
 * בכלל. אירועי גלילה נורים לפני קריאות ה-rAF של אותו פריים, ולכן דריכה בתוך
 * ה-rAF היא הרגע הראשון שבו כבר אין הד תלוי באוויר.
 */
function restorePaneScroll(session: DocumentSession, scroll: PaneScroll): void {
  const arm = (): void => {
    paneScrollGuard?.();
    paneScrollGuard = guardPaneScroll(documentScrollHost(session), scroll);
  };

  applyPaneScroll(documentScrollHost(session), scroll);
  if (typeof requestAnimationFrame !== 'function') {
    arm();
    return;
  }
  requestAnimationFrame(() => {
    // רק אם הוא עדיין הפעיל: מעבר טאב מהיר יותר מפריים היה מחזיר את הגלילה
    // של הטאב הקודם לתוך זה שנכנס אחריו.
    if (activeSession.value !== session) return;
    applyPaneScroll(documentScrollHost(session), scroll);
    arm();
  });
}

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
/**
 * מזהה „המסמך הפתוח” — היום תמיד יחיד, כמו כל שאר ה-state כאן. משמש לנתיב
 * הטיוטה (`draftPathFor`, sessions/session-state.ts): נקבע פעם אחת ב-onMounted,
 * לפני `initSessionKeeper` — מרשומת ההפעלה שנטענה אם יש כזאת, אחרת מזהה חדש.
 * נשאר קבוע לכל אורך ההפעלה, בדיוק כמו ש-`DRAFT_PATH` הקבוע היה קבוע קודם.
 */
let documentId: DocumentSessionId = createDocumentSessionId();
/**
 * מראה עבור רצועת הטאבים בלבד: `documentId` הוא משתנה רגיל ואינו reactive,
 * ומתעדכן בפועל פעם אחת ב-`onMounted` מרשומת ההפעלה — בלי המראה הזה הטאב
 * הראשון היה מוצג עם המזהה הזמני שקדם לטעינה.
 */
const documentIdView = ref<DocumentSessionId>(documentId);
/** מבטל את ההאזנה למעבר לרקע. */
let hiddenListener: (() => void) | null = null;
/** מבטל את ההאזנה לחזרה מהרקע. ראו `onPluginShown` ב-host/lifecycle.ts. */
let shownListener: (() => void) | null = null;
/**
 * טאב שממתין לטעינה נפתח כרגע. ראו `openPendingTab` — הוא מכסה את החלון
 * שלפני ש-`openDocument` מרימה את `isOpening` בעצמה.
 */
let pendingTabLoad = false;

/**
 * האם אסור להתחיל פתיחה או להחליף טאב כרגע.
 *
 * `openDocumentInto` כותב לטאב הפעיל לכל אורך ריצתו, ולכן כל מי שיכול להחליף
 * את הטאב הפעיל או להתחיל פתיחה שנייה חייב לשאול כאן — לא את `isOpening`
 * לבדו, שאינו מכסה את שלב ההכנה של `openPendingTab`.
 */
function isOpenBusy(): boolean {
  return isOpening.value || pendingTabLoad;
}
let contextMenuListener: (() => void) | null = null;
/** מבטל את ההאזנה ליציאה ממסך מלא שלא באה מאיתנו. */
let fullscreenListener: (() => void) | null = null;

const saveStateMessage = computed(() => {
  const state = saveSnapshot.value.state;
  if (state === 'exporting') return 'מייצא…';
  if (state === 'uploading' || state === 'committing') return 'שומר…';
  if (state === 'error') return 'שגיאה בשמירה';
  if (saveSnapshot.value.isDirty) return 'שינויים לא שמורים';
  if (saveSnapshot.value.targetToken) return 'נשמר';
  return 'טרם נשמר';
});

/**
 * ההודעה מ-`outcome.note` **כל עוד היא זו שעל המסך**, אחרת `null`.
 *
 * פקודה מוצלחת בלי הודעה מנקה את הקודמת: „עמודות ← אחת” אחרי „עמודות ←
 * שתיים” חייבת להסיר הודעה שהפכה לשקר. אבל הפס יש לו כותבים נוספים —
 * `STATUS_NOTIFIER` ומסלול פתיחת המסמך — ולכן הניקוי מותנה בכך שההודעה עדיין
 * שם. את התנאי הזה אוכף `setStatus`, שמפקיע את המשתנה בכל כתיבה אחרת לפס.
 */
let lastCommandNote: string | null = null;

function setStatus(text: string, isError = false): void {
  statusText.value = text;
  isStatusError.value = isError;
  // כל כתיבה אחרת לפס מפקיעה את ההודעה. בלי זה `lastCommandNote` היה שורד
  // גם ל-`setStatus('')` של פתיחת מסמך, ומצביע על טקסט שכבר איננו.
  if (text !== lastCommandNote) lastCommandNote = null;
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

  // הצלחה שיש עליה מה לומר — ראו `CommandOutcome.note` ב-engine/command-adapter.ts.
  if (outcome.note) {
    setStatus(outcome.note);
    lastCommandNote = outcome.note;
  } else if (lastCommandNote !== null) {
    // לא-null פירושו שההודעה עדיין על המסך — `setStatus` מפקיע אותו אחרת.
    setStatus('');
  }

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
// הודעות-מידע של כלי הלשוניות („בוצעו 3 תיקונים”) — ראו composables/keys.ts.
provide(STATUS_NOTIFIER, (text: string) => setStatus(text));

/**
 * קואורדינטור השמירה של טאב אחד. `session` הוא ה-`DocumentSession` שהוא
 * שייך אליו — כל התלויות קוראות **דרכו**, לא מהסינגלטונים ברמת המודול, כי
 * שמירה אוטומטית של טאב ברקע חייבת לפעול על המסמך שלו, לא על „הטאב הפעיל”
 * שמזדמן להיות מסמך אחר באותו רגע. `onStateChange`/`onSaved` הם המקום
 * היחיד שכן נוגע ב-refs המוצגים — ורק דרך `guardIfActive`. ראו „ריבוי
 * מסמכים” ליד `sessions` למעלה.
 */
/**
 * `getSession` ולא `session` ישירות: ב-`createNewDocumentSession` (למטה)
 * הקואורדינטור נבנה **לפני** שה-`DocumentSession` שמכיל אותו קיים — פרמטר
 * רגיל היה קופא על `undefined` (העתקה בערך, לא סגירה על המשתנה החיצוני);
 * getter הוא סגירה, ולכן הוא רואה את ההשמה שמגיעה מיד אחרי.
 */
function initSaveCoordinator(getSession: () => DocumentSession): SaveCoordinator {
  return createSaveCoordinator({
    exportDocument: async () => {
      const session = getSession();
      const active = session.swap.current;
      if (!active) throw new Error('אין מסמך פתוח');
      // ה-MIME הוא מה שנשלח כ-`Content-Type` בכתיבה, ו-`superdoc.export`
      // מסמן תמיד `docx`. מסמך עם מאקרו היה מוצהר לא נכון למי שקורא את
      // ההצהרה — ראו `retypeBlob`. הסיומת נקראת מהטאב הזה ולא מה-ref המוצג:
      // שמירה אוטומטית של טאב ברקע רצה כאן, על המסמך שלו.
      return retypeBlob(await exportDocx(active.superdoc), sessionSaveExtension(session));
    },
    beginWrite: (size) => beginBinaryWrite(size),
    upload: uploadBytes,
    abort: abortBinaryWrite,
    commit: (input) =>
      commitUserFileWrite({
        writeToken: input.writeToken,
        targetToken: input.targetToken,
        // הנפילה-לאחור נושאת סיומת גם היא: מסלול שלא העביר `suggestedName`
        // (שמירה אוטומטית) אינו אמור להציע לכתוב מסמך מאקרו בשם בלי סיומת.
        suggestedName:
          input.suggestedName ??
          documentFileName(sessionDisplayTitle(getSession()), sessionSaveExtension(getSession())),
        // הסיומת שהמאחז מצמיד לשם בדיאלוג „שמור בשם”, ושלפיה הוא מסנן בו.
        // בלי השדה הזה היא `docx` קבוע (ראו `CommitOptions` ב-host/files.ts),
        // ומסמך מאקרו היה מוצע לשמירה בשם `ספר.docm.docx`.
        extension: sessionSaveExtension(getSession()),
        title: 'שמירת המסמך',
      }),
    onStateChange: (snapshot) => {
      const session = getSession();
      session.ui.saveSnapshot = snapshot;
      // הקואורדינטור של טאב ברקע מדווח על **המסמך שלו**, וזה מה שמזיז את
      // הנקודה שברצועה. ראו `tabStripRevision`.
      notifyTabStrip();
      guardIfActive(session, () => {
        saveSnapshot.value = snapshot;
      });
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
      const session = getSession();
      // „שמור בשם” מחליף את הקובץ שמאחורי המסמך; „שמור” רגיל אינו. רק
      // הראשון הוא זהות חדשה, ורק אז נכון לשכוח את מקום הסמן.
      if (activeEntry(session.keeper.state)?.document?.token !== info.token) {
        session.keeper.setDocument({ token: info.token, name: info.name, writable: true });
        // „שמור בשם” הוא קובץ חדש שהמשתמש בחר בעצמו — בדיוק מה שרשימת
        // האחרונים אמורה להחזיק. שמירה רגילה אינה נכנסת לכאן (התנאי מעל),
        // ולכן היא אינה מקפיצה את המסמך לראש הרשימה בכל סבב אוטומטי.
        // „שמור בשם” כותב לקובץ שהמשתמש בחר עכשיו, ולכן הוא כתיב בהגדרה —
        // זה גם מה ש-`setDocument` שמעליו כותב לרשומת ההפעלה.
        rememberRecentDocument({
          token: info.token,
          name: info.name,
          size: info.size,
          writable: true,
        });
      }
      // הגודל הוא של מה שנכתב עכשיו, ולכן הוא הבסיס להשוואה הבאה מול
      // הדיסק — בלעדיו „הקובץ השתנה מבחוץ” היה נשאל אחרי כל שמירה רגילה.
      void session.keeper.noteSaved(info.size);
    },
  });
}

/**
 * זוכר ההפעלה של טאב אחד. אותו טעם בדיוק כמו `initSaveCoordinator` —
 * `getSession` ולא `session` — ומאותו טעם התלויות קוראות דרכו, לא
 * מהסינגלטונים: טיוטה של טאב ברקע חייבת להמשיך להיכתב לנתיב שלו, גם כשהוא
 * אינו מוצג.
 */
function initSessionKeeper(getSession: () => DocumentSession, id: DocumentSessionId): SessionKeeper {
  return createSessionKeeper({
    id,
    persist: (state) => persistCombinedSession(getSession(), state),
    exportDocument: () => {
      const active = getSession().swap.current;
      if (!active) throw new Error('אין מסמך פתוח');
      return exportDocx(active.superdoc);
    },
    // ההמרה מ-`Blob` כאן ולא ב-host/workspace.ts: כאן יושב מי שמחזיק את
    // המנוע, ושם יושב מי שמדבר עם הגשר.
    writeDraft: async (content) =>
      writeWorkspaceBytes(draftPathFor(id), new Uint8Array(await content.arrayBuffer())),
    removeDraft: () => deleteWorkspaceEntry(draftPathFor(id)),
    draftPath: draftPathFor(id),
    readCaret: (previous) => {
      const active = getSession().swap.current;
      if (!active) return Promise.resolve(null);
      return readCaretAnchor(active.ui, active.superdoc, previous);
    },
    isDirty: () => getSession().save.snapshot.isDirty === true,
    isSaving: () => getSession().save.snapshot.isSaving === true,
    settleSave: () => getSession().save.settled(),
    // שגיאה ולא הודעה רגילה: זו אינה התקדמות אלא רשת ביטחון שאינה פרושה,
    // והמשתמש צריך לדעת שעליו לשמור בעצמו. רק כשהטאב פעיל — אזהרה על טאב
    // ברקע שהמשתמש לא רואה כרגע היא רעש שאין לו הקשר.
    onDraftTooLarge: () => {
      const session = getSession();
      guardIfActive(session, () => {
        setStatus(
          'המסמך גדול מכדי לשמור ממנו עותק לשחזור — שינויים שלא יישמרו לקובץ עלולים לאבוד',
          true,
        );
      });
    },
  });
}

/**
 * הרשומה הנשמרת בין הפעלות היא של **כל** הטאבים הפתוחים, לא רק של מי שכתב —
 * כל `SessionKeeper` מחזיק את הרשומה שלו בלבד (`sessionForEntry`, ראו
 * session-state.ts), ולכן כאן מרכיבים מחדש את המערך המלא מכל הטאבים בכל פעם
 * שאחד מהם משנה משהו. זו גם הצורה שממנה `restoreTabs` בונה את הטאבים בעלייה
 * הבאה, וזו הסיבה שהיא חייבת להיות שלמה: מה שאינו כאן פשוט לא ייפתח.
 *
 * `activeId` ו-`view` נלקחים מהטאב הפעיל. `view` משותף לכל הטאבים (ראו
 * `updateShellView`), ולכן זו אינה בחירה שרירותית אלא קריאה של אותו ערך.
 */
function persistCombinedSession(session: DocumentSession, state: SessionState): Promise<void> {
  const documents: SessionState['documents'] = [];
  for (const other of sessions.values()) {
    // `activeEntry` ולא `.documents[0]`: זוכר שאימץ רשומה שמורה (`adopt`,
    // session-keeper.ts) עשוי להחזיק כמה מסמכים מהפעלה קודמת מרובת-טאבים —
    // האינדקס הראשון אינו בהכרח זה ששייך לטאב הזה. `activeEntry` מוצא לפי
    // מזהה, בדיוק כמו `emptySessionWithId` שמבטיחה שהמזהה הזה קיים תמיד.
    const entry = activeEntry(other === session ? state : other.keeper.state);
    if (entry) documents.push(entry);
  }
  const active = activeSession.value ?? session;
  const combined: SessionState = {
    version: SESSION_VERSION,
    documents,
    activeId: active.id,
    view: active.keeper.state.view,
  };
  return saveSessionRecord(combined);
}

/**
 * מצב המעטפת — הלשונית ברצועה, הכיווץ, מצב המיקוד — לכל הטאבים.
 *
 * ## למה לכולם ולא לפעיל בלבד
 *
 * ההפרדה שבראש sessions/session-state.ts אומרת ש-`view` שייך למי שיושב מול
 * המסך ואחת לכל ההפעלה, בשונה מהמסמך והסמן ששייכים לטאב. אבל הזוכר הוא
 * פר-טאב, וכל אחד מחזיק עותק משלו — ומה שנכתב ל-storage הוא של הטאב הפעיל
 * (`persistCombinedSession`). כתיבה לפעיל בלבד פירושה שהעדפה שנבחרה בטאב א'
 * נעלמת אם המשתמש סוגר את התוסף כשהוא עומד על טאב ב': ב' מעולם לא שמע עליה,
 * והוא זה שכותב. נמדד בשער `check:session` — הלשונית „הפניות” חזרה כ„בית”.
 *
 * `zoom` **אינו** עובר כאן, ובכוונה: „150%” הוא משפט על מסמך מסוים (ראו
 * `documentViewFor`), והוא מדווח פר-מסמך מהמנוע עצמו.
 */
function updateShellView(patch: Partial<SessionView>): void {
  for (const session of sessions.values()) session.keeper.updateView(patch);
}

/**
 * יוצרת את סגירת ה-`openEditor` של טאב אחד — ראו `OpenEditor` ב-editor-swap.ts.
 *
 * קוראת דרך `session.save`/`session.metrics`/... ולא מהסינגלטונים, מאותו טעם
 * כמו `initSaveCoordinator`: זו סגירה ארוכת-חיים שנקראת מהמנוע עצמו (עימוד,
 * עריכה) ועשויה להיקרא בזמן שהטאב הזה ברקע.
 */
function createOpenEditorForSession(session: DocumentSession): OpenEditor {
  return (host, source, signal) =>
    createEditor({
      container: host,
      source,
      signal,
      onError: (err) => console.error('[otzaria-word] שגיאת מנוע:', err),
      onUpdate: () => {
        session.save.markDirty();
        session.metrics?.noteDocumentChanged();
        session.ruler?.noteDocumentChanged();
        session.textCursor?.noteDocumentChanged();
        session.pageBorders?.noteDocumentChanged();
        session.lineNumbers?.noteDocumentChanged();
        session.formattingMarks?.noteDocumentChanged();
        // המילון משותף לכל הטאבים, ולכן זה מונה אחד ולא שדה של ה-session.
        // המחיר: עריכה בטאב **ברקע** מריצה מדידה על המסמך שעל המסך. היא לא
        // תצייר כלום (`sameTextSegments` מזהה שדבר לא זז), אבל היא כן נמדדת
        // — ‎~1.2ms לעמוד גלוי, בהשקטה של 400ms. זה זול מספיק מכדי להצדיק
        // מונה לכל טאב, וזה נאמר כאן כדי שלא ייקרא כאילו הוא חינם.
        noteSpellcheckChanged();
        session.keeper.noteChange();
      },
      onPaginationUpdate: (totalPages) => session.metrics?.notePaginationUpdate(totalPages),
    });
}

/**
 * תמונת המצב שטאב ששוחזר מציג **לפני** שהמסמך שלו נטען.
 *
 * שלושת השדות הם בדיוק מה שרצועת הטאבים קוראת (`sessionDisplayTitle`,
 * `sessionIsDirty`) ומה שהפס העליון יראה אם יעברו אליו לפני שהפתיחה הסתיימה.
 * „לא נשמר” נגזר מקיום הטיוטה ברשומה, וזו אמת ולא קישוט: טיוטה פירושה עבודה
 * שאינה בדיסק, בדיוק כמו הנקודה שמופיעה על טאב פתוח.
 */
function restoredUiSnapshot(entry: SessionDocumentEntry, fallbackTitle?: string): DocumentUiSnapshot {
  const base = emptyUiSnapshot();
  const remembered = entry.document;
  return {
    ...base,
    title: remembered ? stripWordExtension(remembered.name) : (fallbackTitle ?? base.title),
    saveSnapshot: {
      ...base.saveSnapshot,
      isDirty: entry.draft !== null,
      targetToken: remembered?.writable ? remembered.token : null,
      name: remembered?.name ?? null,
    },
  };
}

interface NewTabOptions {
  /** ברירת המחדל: מזהה חדש. נמסר בפירוש כשהטאב נוצר מרשומה שמורה. */
  id?: DocumentSessionId;
  /**
   * הרשומה שהטאב נוצר ממנה בשחזור ההפעלה, ומצב התצוגה המשותף שלצידה. נוכחותה
   * היא שהופכת אותו ל„ממתין לטעינה” — ראו „טעינה עצלה” ב-`restoreTabs`.
   */
  restore?: { entry: SessionDocumentEntry; view: SessionView };
}

/**
 * טאב חדש, ריק — עדיין בלי מסמך פתוח בתוכו. `pane` נוצר מוסתר: `activateTab`
 * הוא מי שחושף אותו, ולא היצירה עצמה — טאב שנוצר ברקע (למשל בזמן שחזור
 * ההפעלה) אסור לו להבליח לרגע לפני שההחלטה מי פעיל נופלת.
 */
function createNewDocumentSession(options: NewTabOptions = {}): DocumentSession {
  const restore = options.restore;
  const id = options.id ?? createDocumentSessionId();
  const pane = document.createElement('div');
  pane.className = 'document-pane';
  pane.style.display = 'none';
  editorStackRef.value?.appendChild(pane);

  // `session` נחוץ לסגירות (closures) של save/keeper/swap לפני שהוא עצמו קיים
  // — הן רק שומרות את המשתנה (נקרא בפועל מאוחר יותר, לא כאן), ולכן ההצהרה
  // המוקדמת עם `!` בטוחה: עד שמישהו קורא לתוכן שלהן, ההשמה למטה כבר קרתה.
  let session!: DocumentSession;
  const getSession = (): DocumentSession => session;
  session = createDocumentSession({
    id,
    pane,
    swap: createEditorSwap(pane, (host, source, signal) =>
      createOpenEditorForSession(session)(host, source, signal),
    ),
    save: initSaveCoordinator(getSession),
    keeper: initSessionKeeper(getSession, id),
    pendingRestore: restore !== undefined,
    ui: restore ? restoredUiSnapshot(restore.entry) : undefined,
  });
  // הזוכר מאמץ את הרשומה **של הטאב הזה בלבד** — ההנמקה ליד `sessionForEntry`
  // ב-session-state.ts. לפני `sessions.set`, כדי שהרשומה המשולבת שתיכתב
  // בעקבות טאב אחר לא תתפוס אותו רגע אחד עם רשומה ריקה במקומו.
  if (restore) session.keeper.adopt(sessionForEntry(restore.entry, restore.view));
  // טאב חדש נולד עם `defaultView()`, ומרגע שהוא הפעיל **הוא** זה שכותב את
  // `view` לרשומה (`persistCombinedSession`) — כלומר פתיחת טאב הייתה מוחקת
  // מה-storage את הלשונית, הכיווץ ומצב המיקוד שהמשתמש בחר. ראו `updateShellView`.
  else {
    session.keeper.updateView({
      ribbonTab: ribbonTab.value,
      ribbonCollapsed: ribbonCollapsed.value,
      focusMode: isFocusMode.value,
    });
  }
  session.save.setAutosaveEnabled(autosaveEnabled.value);
  sessions.set(session.id, session);
  // ידית QA: כל הטאבים הפתוחים, לפי מזהה — ראו התיוג הבודד ב-create-editor.ts
  // (`window.__otzariaEditor`, הטאב הפעיל בלבד). קוד האפליקציה אינו קורא כאן.
  (window as unknown as { __otzariaEditors?: Map<string, DocumentSession> }).__otzariaEditors = sessions;
  return session;
}

/**
 * שומרת את מה שהטאב הפעיל כרגע מציג, לתוך `ui` שלו. נקראת **לפני** שעוזבים
 * אותו — ראו „ריבוי מסמכים” ליד `sessions` למעלה.
 */
function stashActiveInto(session: DocumentSession): void {
  session.ui = {
    title: title.value,
    commandAdapter: commandAdapter.value,
    activeSuperdoc: activeSuperdoc.value,
    activeEditorContainer: activeEditorContainer.value,
    documentGeneration: documentGeneration.value,
    activeMacros: activeMacros.value,
    docMetrics: docMetrics.value,
    zoom: zoom.value,
    canUndo: canUndo.value,
    canRedo: canRedo.value,
    rulerReading: rulerReading.value,
    rulerHost: rulerHost.value,
    rulerViewport: rulerViewport.value,
    rulerUnit: rulerUnit.value,
    isDocumentEditable: isDocumentEditable.value,
    isRulerVisible: isRulerVisible.value,
    pageBorders: pageBorders.value,
    lineNumbering: lineNumbering.value,
    formattingMarksBlocks: formattingMarksBlocks.value,
    formattingMarksVisible: formattingMarksVisible.value,
    saveSnapshot: saveSnapshot.value,
    searchState: searchState.value,
    styleGallery: styleGallery.value,
    engineFontSlice: engineFontSlice.value,
    readoutSelection: readoutSelection.value,
    documentVba: documentVba.value,
    // נקרא **לפני** ש-`activateTab` מסתיר את הפאנל: ברגע שהוא
    // `display: none` המספר הזה כבר אבד. ראו sessions/pane-scroll.ts.
    scroll: readPaneScroll(documentScrollHost(session)),
    saveExtension: saveExtension.value,
  };
  swap = null;
  save = null;
  keeper = null;
  metrics = null;
  ruler = null;
  pageBorderModel = null;
  lineNumberModel = null;
  formattingMarksModel = null;
  searchAdapter = null;
}

/**
 * טוענת לתוך הסינגלטונים ברמת המודול את מה שהטאב הזה מציג — ההופכי המדויק
 * של `stashActiveInto`, בשדה אחד למעט: **מיקום הגלילה**.
 *
 * הוא נשמר ב-`stashActiveInto` אבל מוחזר ב-`activateTab` ולא כאן, ולא מטעמי
 * נוחות: החזרה שלו חייבת לקרות **אחרי** ש-`display` של הפאנל חוזר, כי לפני
 * כן אין לו קופסה שאפשר לגלול בה. ראו `restorePaneScroll`.
 */
function restoreFromSession(session: DocumentSession): void {
  const ui = session.ui;
  title.value = ui.title;
  commandAdapter.value = ui.commandAdapter;
  activeSuperdoc.value = ui.activeSuperdoc;
  activeEditorContainer.value = ui.activeEditorContainer;
  documentGeneration.value = ui.documentGeneration;
  activeMacros.value = ui.activeMacros;
  docMetrics.value = ui.docMetrics;
  zoom.value = ui.zoom;
  canUndo.value = ui.canUndo;
  canRedo.value = ui.canRedo;
  rulerReading.value = ui.rulerReading;
  rulerHost.value = ui.rulerHost;
  rulerViewport.value = ui.rulerViewport;
  rulerUnit.value = ui.rulerUnit;
  isDocumentEditable.value = ui.isDocumentEditable;
  isRulerVisible.value = ui.isRulerVisible;
  pageBorders.value = ui.pageBorders;
  lineNumbering.value = ui.lineNumbering;
  formattingMarksBlocks.value = ui.formattingMarksBlocks;
  formattingMarksVisible.value = ui.formattingMarksVisible;
  saveSnapshot.value = ui.saveSnapshot;
  searchState.value = ui.searchState;
  styleGallery.value = ui.styleGallery;
  engineFontSlice.value = ui.engineFontSlice;
  readoutSelection.value = ui.readoutSelection;
  documentVba.value = ui.documentVba;
  saveExtension.value = ui.saveExtension;

  swap = session.swap;
  save = session.save;
  keeper = session.keeper;
  metrics = session.metrics;
  ruler = session.ruler;
  pageBorderModel = session.pageBorders;
  lineNumberModel = session.lineNumbers;
  formattingMarksModel = session.formattingMarks;
  searchAdapter = session.searchAdapter;
  documentId = session.id;
  documentIdView.value = session.id;
  (window as unknown as { __otzariaEditor?: unknown }).__otzariaEditor = session.swap.current;
}

/**
 * מעבר טאב. אין מה לעשות כשהטאב המבוקש כבר פעיל — לא רק אופטימיזציה: הרצה
 * כפולה הייתה שוברת את `stashActiveInto`/`restoreFromSession` (שומרת אל
 * ומטעינה מאותו טאב, שזה no-op תקין, אבל גם מפסיקה ומתחילה מחדש שכבות
 * שאינן אמורות להבהב, כמו ה-`pane` שמוסתר ומוצג בחזרה).
 */
function activateTab(session: DocumentSession): void {
  if (activeSession.value === session) return;

  // השומר שייך לטאב שיוצא, והוא מאזין ל-host שלו. ראו `restorePaneScroll`.
  paneScrollGuard?.();
  paneScrollGuard = null;

  const previous = activeSession.value;
  if (previous) {
    stashActiveInto(previous);
    previous.pane.style.display = 'none';
  }

  session.pane.style.display = '';
  restoreFromSession(session);
  activeSession.value = session;
  // אחרי החשיפה ואחרי ההשמה ל-`activeSession`: הראשונה מחזירה לפאנל קופסה
  // שאפשר לגלול בה, והשנייה היא מה שהבדיקה שבפריים הבא נשענת עליה.
  restorePaneScroll(session, session.ui.scroll);
  // ובמצב מיקוד — גם הפוקוס. הלחיצה על הטאב באה מהלוח שנחשף, והפוקוס נשאר על
  // כפתור הטאב; ברגע שהלוח נסגר `visibility: hidden` מבריח אותו ל-`<body>`
  // (נמדד `active: BODY`), וההקלדה נעלמת בלי שום סימן על המסך. אותו טעם בדיוק
  // כמו ב-`toggleFocusMode`, ורק שם: מחוץ למצב מיקוד רצועת הטאבים נשארת
  // גלויה, ופוקוס שממשיך לשבת על הטאב שנלחץ הוא ההתנהגות הנכונה.
  const opened = activeSuperdoc.value;
  if (isFocusMode.value && opened) focusOpenedDocument(opened);
  noteTabUsed(session);
  void trimLiveDocuments();

  // „מי היה פעיל” הוא חלק מהרשומה, והוא משתנה בדיוק כאן. בלי הכתיבה הזאת
  // הוא נשמר רק כשמשהו אחר גרם לכתיבה (עריכה, יציאה מסודרת), ומעבר טאב
  // שלא לווה בעריכה היה חוזר בהפעלה הבאה לטאב הקודם.
  //
  // מיידית ולא מושהית, בשונה מ-`schedulePersist` של הזוכר: ההשהיה שם קיימת
  // כדי לא לכתוב על כל תו שנקלד, וכאן המקור הוא מחווה בודדת של המשתמש —
  // ואחריה עלול לבוא מיד המעבר שסוגר את התוסף. כתיבה אחת ל-storage, בלי
  // ייצוא ובלי טיוטה (ראו „שני קצבים” ב-session-keeper.ts).
  void persistCombinedSession(session, session.keeper.state).catch((error: unknown) => {
    console.warn('[otzaria-word] שמירת הטאב הפעיל נכשלה', error);
  });
}

/** מסמנת את הטאב כאחרון שהשתמשו בו. ראו `recentTabs`. */
function noteTabUsed(session: DocumentSession): void {
  const at = recentTabs.indexOf(session.id);
  if (at >= 0) recentTabs.splice(at, 1);
  recentTabs.unshift(session.id);
}

/** כמה טאבים מחזיקים כרגע מנוע חי. ראו `MAX_LIVE_DOCUMENTS`. */
function liveDocumentCount(): number {
  let live = 0;
  for (const session of sessions.values()) if (session.swap.current) live += 1;
  return live;
}

/**
 * מה שידוע על הטאב, בצורה שההחלטה מבינה. ההחלטה עצמה ב-sessions/sleep-policy.ts
 * ולא כאן — היא קובעת אם עבודה של המשתמש נמחקת מהזיכרון, וקוד כזה חייב
 * להיבדק.
 */
function sleepCandidate(session: DocumentSession, hasUnwrittenWork: boolean): SleepCandidate {
  return {
    isActive: session === activeSession.value,
    isPending: session.pendingRestore,
    hasEngine: session.swap.current !== null,
    isOpening: session.swap.isOpening,
    isSaving: session.save.snapshot.isSaving === true,
    hasFile: activeEntry(session.keeper.state)?.document != null,
    hasUnwrittenWork,
  };
}

/**
 * סינון מוקדם, לפני שמשקיעים `flush` שלם.
 *
 * `hasUnwrittenWork: false` כאן אינו טענה אלא „עוד לא ידוע”: לפני ה-`flush`
 * כל מסמך שנערך מדווח שיש בו עבודה שלא נכתבה, וזו בדיוק העבודה שה-`flush`
 * הולך לכתוב. השאלה האמיתית נשאלת אחריו, ב-`sleepTab`.
 */
function canSleepTab(session: DocumentSession): boolean {
  return decideSleep(sleepCandidate(session, false)).action === 'sleep';
}

/**
 * „טאב נרדם”: משחררת את המנוע של טאב ברקע, ומשאירה אותו ברצועה בדיוק כפי
 * שנראה טאב ששוחזר מהפעלה קודמת — שם, סימון „לא נשמר”, ורשומה מלאה.
 *
 * ## למה זה בטוח, ואיפה הגבול
 *
 * שחרור מנוע הוא מחיקה של כל מה שיש בו בזיכרון. מה שמאפשר אותה הוא שהכול
 * כבר נמצא במקום שאפשר לקרוא ממנו בחזרה: הקובץ בדיסק, העבודה שלא נשמרה
 * בטיוטה, והסמן והזום ברשומה. `flush()` הוא מה שמוודא שזה נכון **עכשיו**
 * ולא „בעוד עשר שניות”, ו-`decideSleep` הוא מי שבודק שהוא הצליח.
 *
 * שני דברים **אינם** חוזרים לטאב שנרדם, וזה כל המחיר: **היסטוריית ה-Undo**
 * ומיקום הגלילה המדויק — לשניהם אין ייצוג לא בקובץ ולא ברשומה (הסמן כן
 * חוזר, וממילא מחזיר את התצוגה לאזור הנכון). לכן התקרה אינה 1 ולכן הסדר
 * הוא לפי שימוש אחרון: הטאב שנרדם הוא זה שהמשתמש לא נגע בו הכי הרבה זמן,
 * ולא זה שהוא הרגע עזב.
 *
 * ההחלטה נשאלת שוב אחרי ה-`flush` ולא רק לפניו: הוא `await`, והמצב עשוי
 * להשתנות בתוכו — המשתמש עבר לטאב הזה, שמירה התחילה בו, או שהוא נסגר.
 */
async function sleepTab(session: DocumentSession): Promise<void> {
  await session.keeper.flush();

  const decision = decideSleep(sleepCandidate(session, session.keeper.hasUnwrittenWork));
  if (decision.action === 'keep') {
    // רק הסיבה הזאת ראויה לשורה בלוג: כל השאר הן „לא עכשיו”, וזו אומרת
    // שרשת הביטחון של הטיוטה אינה פרושה על המסמך הזה.
    if (decision.reason === 'unwritten-work') {
      console.info(`[otzaria-word] ${session.ui.title} נשאר בזיכרון — יש בו עבודה שלא נכתבה`);
    }
    return;
  }

  const entry = activeEntry(session.keeper.state);
  if (!entry) return;

  const title = session.ui.title;
  session.pendingRestore = true;
  session.slept = true;
  session.swap.close();
  // אחרי `close` ולא לפניו: הפירוק עצמו כותב לתמונת המצב של הטאב (ראו
  // `editor.onDispose` ב-openDocumentInto), וכאן היא נקבעת סופית — נקייה
  // ובלי הפניות למנוע שכבר אינו קיים.
  session.ui = restoredUiSnapshot(entry, title);
  // תמונת מצב חדשה לטאב שברקע — אותו טעם כמו ב-`onStateChange`.
  notifyTabStrip();
}

/**
 * מרדימה טאבים מעבר לתקרה, מהוותיק שבהם. ראו `MAX_LIVE_DOCUMENTS`.
 *
 * נקראת אחרי מעבר טאב ואחרי פתיחת מסמך — שתי הנקודות שבהן מספר המנועים
 * החיים או סדר השימוש משתנים. שקטה לחלוטין כשאין מה לעשות, וזה המצב אצל
 * כמעט כל משתמש.
 */
async function trimLiveDocuments(): Promise<void> {
  if (trimming) return;
  trimming = true;
  try {
    // עותק, ולא המערך עצמו: `sleepTab` ממתין, ומעבר טאב באמצע ההמתנה מזיז
    // איברים ב-`recentTabs` (`noteTabUsed`) — סריקה על המערך החי הייתה
    // מדלגת על מועמד או סורקת אותו פעמיים. מי שכבר אינו במפה מדולג ממילא.
    const candidates = [...recentTabs].reverse();
    for (const id of candidates) {
      if (liveDocumentCount() <= MAX_LIVE_DOCUMENTS) return;
      const candidate = sessions.get(id);
      if (candidate && canSleepTab(candidate)) await sleepTab(candidate);
    }
  } catch (error) {
    // שחרור זיכרון הוא שיפור, לא תנאי לשום דבר: כשל כאן משאיר מסמך פתוח.
    console.warn('[otzaria-word] שחרור מסמך שברקע נכשל', error);
  } finally {
    trimming = false;
  }
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
  /**
   * שם למסמך שנפתח מבייטים בלבד — כלומר בלי קובץ.
   *
   * למסלול אחד: פתיחת גיבוי מ„נסגרו בלי לשמור” (`onOpenDiscarded`). בלעדיו
   * המסמך היה נקרא „מסמך חדש”, ומי שפתח שלושה גיבויים היה מקבל שלושה טאבים
   * באותו שם — בלי שום דרך לדעת מה יש בכל אחד. אין לו יעד כתיבה בכל מקרה;
   * זהו שם לתצוגה, ו„שמור” עדיין פותח „שמור בשם”.
   */
  name?: string;
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

/**
 * פתיחת מסמך, עם המחוון שמלווה אותה.
 *
 * העטיפה קיימת בשביל ה-`finally`, והוא אינו הידור: הפתיחה ממשיכה לכתוב למצב
 * המעטפת גם אחרי שהמנוע הצליח — מודדים, סרגל, שחזור התצוגה — וחריגה באחד מהם
 * הייתה מדלגת על `finish()` ומשאירה פס תקוע על 96% לכל אורך ההפעלה. `fail`
 * הוא no-op על פתיחה שכבר הוכרעה, ולכן הוא רק סוגר את המסלול הזה.
 *
 * שם המסמך הנטען הוא ההודעה, והוא נכנס למחוון ולא לשורת המצב: שניהם יושבים
 * באותה שורה של 24 פיקסלים, ואין טעם שיאמרו את אותו דבר פעמיים.
 */
async function openDocument(file?: UserFile, options: OpenOptions = {}): Promise<boolean> {
  if (!swap) return false;
  const attempt = documentLoad.begin(
    file?.name ?? options.name ?? 'מסמך חדש',
    file || options.draft ? 'קורא את הקובץ…' : 'מכין מסמך ריק…',
  );
  try {
    return await openDocumentInto(attempt, file, options);
  } finally {
    attempt.fail();
    // וגם המצב שמשתק את הרצועה: פתיחה שזרקה השאירה אותה מנוטרלת לתמיד. הערך
    // נקרא מה-swap ולא נקבע ל-false, כדי שפתיחה חדשה יותר שכבר בדרך לא
    // תיראה כמי שהסתיימה.
    isOpening.value = swap?.isOpening ?? false;
    // מנוע חי נוסף נולד כאן — הנקודה השנייה שבה התקרה יכולה להישבר (הראשונה
    // היא מעבר טאב). ראו `MAX_LIVE_DOCUMENTS`.
    void trimLiveDocuments();
  }
}

async function openDocumentInto(
  attempt: LoadAttempt,
  file?: UserFile,
  options: OpenOptions = {},
): Promise<boolean> {
  if (!swap) return false;
  // ראו „ריבוי מסמכים” ליד `sessions`: פתיחה תמיד מתרחשת לתוך הטאב הפעיל —
  // `onPickAndOpen`/`onNewDocument`/`onDocumentTabNew` מפעילים טאב חדש (אם
  // צריך) **לפני** קריאה לכאן. `session` משמש רק למראות שצריכות לשרוד מעבר
  // טאב — ראו את השימושים למטה.
  const session = activeSession.value;
  isOpening.value = true;
  const startedAt = performance.now();
  setStatus('');

  // לפני המנוע ולא אחריו: ערך אחד ב-`word/settings.xml` שולח אותו ללולאה על
  // החוט הראשי, ומשם אין חזרה — גם `OPEN_TIMEOUT_MS` אינו יכול לירות. ראו
  // engine/docx-preflight.ts. השלב הזה אינו יכול להיכשל: הוא מחזיר את המקור
  // כמות שהוא בכל מקרה שאינו „מצאתי בדיוק את הערך הזה”.
  const { source, fontTable, vba, notice: preflightNotice } = await preflightSource(options.draft ?? file?.url);
  // „דלג” בזמן שהבייטים נקראו. הבדיקה כאן ולא רק לפני המנוע: שאר הפונקציה
  // כותבת למצב של המעטפת — כותרת, יעד שמירה, רשומת ההפעלה — ופתיחה שהמשתמש
  // נטש אינה אמורה לכתוב שם דבר.
  if (attempt.cancelled) return false;
  attempt.stage(LOAD_STAGES.fonts, 'מכין את הגופנים…');

  // לפני שהמנוע מודד, ולא אחרי: `lineRule="auto"` גוזר את גובה השורה ממדדי
  // הגופן שנבחר בפועל, ולכן גופן חסר משנה את פריסת כל המסמך — לא רק את מראהו.
  // ראו engine/docx-fonts.ts. אינו יכול להיכשל, ואינו מעכב כשאין מה להחליף.
  await installDocumentFontAliases(fontTable);
  if (attempt.cancelled) return false;
  attempt.stage(LOAD_STAGES.engine, 'בונה את המסמך…');

  // מסמך חדש נפתח מהתבנית העברית (A4, RTL) ולא מהמסמך הריק של המנוע — ראו
  // engine/blank-document.ts. `undefined` בבדיקות, ואז המנוע פותח את שלו.
  const blank = !file && !options.draft ? blankDocumentSource() : undefined;
  const outcome = await swap.open(source ?? blank);
  isOpening.value = swap.isOpening;

  // ה-swap מדווח `superseded` גם על פתיחה שבוטלה — הביטול מקדם שם את הדור
  // לפני שהוא מרים את האיתות (ראו EditorSwap.cancel). המחוון כבר סולק
  // על ידי מי שביטל, ולכן אין כאן מה לסגור.
  if (outcome.status === 'superseded') return false;

  if (outcome.status === 'failed') {
    attempt.fail();
    const kept = swap.current ? ` ${title.value} נשאר פתוח.` : '';
    setStatus(`פתיחת המסמך נכשלה: ${outcome.error.message}.${kept}`, true);
    return false;
  }

  attempt.stage(LOAD_STAGES.arranging, 'מסדר את התצוגה…');

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
    // בלי האיפוס הרצועה הייתה ממשיכה להחזיק את הקריאה של המסמך שנסגר. מותנה
    // בטאב: פירוק יכול לקרות גם על טאב שברקע (סגירת טאב, טאב שנרדם — ראו
    // „ריבוי מסמכים” ליד `sessions`), ואיפוס גורף היה מוחק את הקריאה של
    // המסמך שהמשתמש מסתכל עליו דווקא.
    if (session) session.ui.readoutSelection = UNSETTLED_SELECTION;
    if (!session || session === activeSession.value) readoutSelection.value = UNSETTLED_SELECTION;
  });

  // החיפוש שייך ל-session: ה-handle הוא של ה-controller של המופע, ומסמך חדש
  // מקבל אדפטר חדש. ה-`session` המקומי ולא `searchAdapter` בפירוק — אחרת
  // סגירת המסמך הקודם הייתה מפרקת את האדפטר של המסמך שנפתח אחריו.
  // `observe` יורה מיד עם ה-snapshot ואז על כל שינוי: המנוע פותר את גופני
  // המסמך אחרי שהוא נפתח, ובלי האזנה הבורר היה קופא על הרשימה של הרגע הראשון.
  editor.onDispose(
    observeFontSlice(editor.ui, (slice) => {
      // א-סינכרוני, ולכן ייתכן שיירה אחרי שהמשתמש עבר לטאב אחר — ראו
      // „ריבוי מסמכים” ליד `sessions`.
      if (session) {
        session.ui.engineFontSlice = slice;
        guardIfActive(session, () => {
          engineFontSlice.value = slice;
        });
      }
    })
  );

  // אותו טעם, ועוד יותר: `getQuickGallery()` מחזיר רשימה ריקה עד שהקטלוג
  // מתייצב, ולכן בלי ההרשמה הגלריה הייתה נשארת על רשת הביטחון לתמיד.
  editor.onDispose(
    observeStyleGallery(editor.ui, (state) => {
      if (session) {
        session.ui.styleGallery = state;
        guardIfActive(session, () => {
          styleGallery.value = state;
        });
      }
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
      // כלי „שולחן העורך” נרשמים על ה-kit של המסמך הזה: מופיעים בדיאלוג
      // ניהול המאקרו וניתנים לקיצור מקלדת. הרישום פר-התקנה — ה-kit נבנה
      // מחדש בכל פתיחה, ואין צורך בביטול.
      registerShulchanTools(macros.kit, () => editor.superdoc, (text) => setStatus(text));
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
      // `onPaginationUpdate` (createOpenEditorForSession) הוא א-סינכרוני
      // ועשוי לירות אחרי שהטאב עבר לרקע — ראו „ריבוי מסמכים” ליד `sessions`.
      if (session) session.ui.docMetrics = next;
      if (!session || session === activeSession.value) docMetrics.value = next;
    },
  });
  metrics = sessionMetrics;
  if (session) session.metrics = sessionMetrics;
  docMetrics.value = sessionMetrics.getState();
  editor.onDispose(() => {
    sessionMetrics.dispose();
    if (session && session.metrics === sessionMetrics) session.metrics = null;
    if (metrics === sessionMetrics) {
      metrics = null;
      if (!session || session === activeSession.value) docMetrics.value = emptyDocMetrics();
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
      // מושהה (`RULER_*_DEBOUNCE_MS`) ולכן עשוי לנחות אחרי מעבר טאב — ראו
      // „ריבוי מסמכים” ליד `sessions`. בלי ההגנה, הסרגל של הטאב שעל המסך
      // היה מציג את השוליים של מסמך אחר, וגרירת ידית הייתה מחשבת מהם.
      if (session) session.ui.rulerReading = next;
      if (!session || session === activeSession.value) rulerReading.value = next;
    },
  });
  ruler = sessionRuler;
  if (session) session.ruler = sessionRuler;
  rulerHost.value = paintedHost(editor.ui);
  rulerViewport.value = editor.ui as ViewportSource;
  rulerUnit.value = readRulerUnit(editor.superdoc);
  editor.onDispose(() => {
    sessionRuler.dispose();
    if (session && session.ruler === sessionRuler) session.ruler = null;
    // הגנה על טאב אחר: פירוק כאן יכול לקרות בזמן שהטאב הזה ברקע (סגירת
    // טאב, למשל) — ראו „ריבוי מסמכים” ליד `sessions`. בלי התנאי, סגירת טאב
    // ברקע הייתה מוחקת את הסרגל של הטאב הפעיל.
    if (ruler === sessionRuler && (!session || session === activeSession.value)) {
      ruler = null;
      rulerReading.value = null;
      rulerHost.value = null;
      rulerViewport.value = null;
    }
  });

  /**
   * סמן-הטקסט של העכבר (I-beam על כל עמודת הטקסט, כמו Word) — ראו
   * engine/text-cursor.ts. פר-session כי הוא מאזין על ה-host של **המסמך
   * הזה** וכותב עליו `style.cursor`; טאב ברקע ממשיך להחזיק מעקב משלו בלי
   * להפריע לפעיל. `refreshNow` מיד — מסמך חדש צריך את הסמן הנכון מהרגע
   * הראשון, לא אחרי העריכה הראשונה.
   */
  const sessionTextCursor = createTextCursorWatch({
    host: paintedHost(editor.ui),
    ui: editor.ui as ViewportSource,
    readMargins: () => readPageMargins(editor.superdoc),
  });
  if (session) session.textCursor = sessionTextCursor;
  sessionTextCursor.refreshNow();
  editor.onDispose(() => {
    sessionTextCursor.dispose();
    if (session && session.textCursor === sessionTextCursor) session.textCursor = null;
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
      // מושהה, כמו הסרגל — ומכאן אותה הגנה בדיוק. שכבת הציור היא מופע יחיד
      // מעל הפאנל הפעיל, ולכן קריאה של מסמך אחר נראית מיד על המסך.
      if (session) session.ui.pageBorders = next;
      if (!session || session === activeSession.value) pageBorders.value = next;
    },
  });
  pageBorderModel = sessionPageBorders;
  if (session) session.pageBorders = sessionPageBorders;
  sessionPageBorders.refreshNow();
  editor.onDispose(() => {
    sessionPageBorders.dispose();
    // גם השדה שב-session, ולא רק הסינגלטון: טאב שנרדם (`sleepTab`) היה
    // ממשיך להחזיק דרכו מודל מפורק שסגירתו לוכדת את ה-SuperDoc ההרוס —
    // כלומר בדיוק הזיכרון שהשחרור נועד לפנות.
    if (session && session.pageBorders === sessionPageBorders) session.pageBorders = null;
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
      if (session) session.ui.lineNumbering = next;
      if (!session || session === activeSession.value) lineNumbering.value = next;
    },
  });
  lineNumberModel = sessionLineNumbering;
  if (session) session.lineNumbers = sessionLineNumbering;
  sessionLineNumbering.refreshNow();
  editor.onDispose(() => {
    sessionLineNumbering.dispose();
    if (session && session.lineNumbers === sessionLineNumbering) session.lineNumbers = null;
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
      if (session) session.ui.formattingMarksBlocks = next;
      if (!session || session === activeSession.value) formattingMarksBlocks.value = next;
    },
  });
  formattingMarksModel = sessionFormattingMarks;
  if (session) session.formattingMarks = sessionFormattingMarks;
  editor.onDispose(() => {
    sessionFormattingMarks.dispose();
    if (session && session.formattingMarks === sessionFormattingMarks) session.formattingMarks = null;
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

  /**
   * ההחלפה הייתה בחירה של המשתמש, ולכן היא זו שנשמרת להפעלה הבאה.
   *
   * רק כשהטאב הזה על המסך: `void adapter.run('ruler')` שמתחת אינו מומתן,
   * והדיווח שלו עשוי לנחות אחרי שהמשתמש כבר עבר טאב. העדפה שנכתבת ממסמך
   * שאינו מוצג היא בדיוק „כיביתי את הסרגל והוא חזר לבד”.
   */
  const rememberRulerVisible = (active: boolean): void => {
    if (session && session !== activeSession.value) return;
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
      // הזוכר **של הטאב הזה**, בדיוק כמו ב-`observeZoom` שמתחת: הסינגלטון
      // מתאפס למשך מעבר טאב (`stashActiveInto`), ובטאב שברקע הוא כבר של
      // מסמך אחר — כלומר תזוזת הסמן הייתה נרשמת אצל השכן, או נבלעת.
      (session?.keeper ?? keeper)?.noteChange();
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
      // עשוי לירות מהתאמת רוחב חלון גם על טאב ברקע — ראו „ריבוי מסמכים” ליד
      // `sessions`. הכתיבה לתצוגה מוגנת; העדכון ל-session.keeper ולמרכוז
      // (שאינם משותפים, ראו שם) אינם צריכים הגנה.
      if (session) {
        session.ui.zoom = state;
        guardIfActive(session, () => {
          zoom.value = state;
          // אותו דיווח מזין גם את המרכוז: הוא יורה על כל שינוי, כולל שינוי
          // שלא בא מאיתנו, ולכן העמוד אינו נשאר ממורכז לפי אחוז ישן. ראו
          // engine/zoom-center.ts.
          zoomCenter?.setZoom(state.value);
        });
        // ואותו דיווח מזין גם את הזיכרון בין הפעלות. כאן ולא במטפל של הסרגל,
        // מאותו טעם בדיוק: גם שינוי שלא בא מאיתנו הוא גודל התצוגה שהמשתמש
        // רואה, והוא זה שצריך לחזור. `session.keeper` ולא `keeper` הסינגלטון —
        // האחרון מתאפס כשהטאב עובר לרקע.
        session.keeper.updateView({ zoom: state.value });
      }
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
  if (session) session.searchAdapter = sessionSearch;
  searchState.value = sessionSearch.getState();
  editor.onDispose(
    sessionSearch.subscribe((state) => {
      // „3 מתוך 12” של מסמך אחר הוא בדיוק מה שההערה ליד `searchState` מזהירה
      // ממנו: החיפוש-בזמן-הקלדה מושהה, והדיווח עשוי לנחות אחרי מעבר טאב.
      if (session) session.ui.searchState = state;
      if (!session || session === activeSession.value) searchState.value = state;
    })
  );
  editor.onDispose(() => {
    sessionSearch.dispose();
    if (session && session.searchAdapter === sessionSearch) session.searchAdapter = null;
    if (searchAdapter === sessionSearch) {
      searchAdapter = null;
      searchState.value = idleSearchState();
    }
  });

  title.value = stripWordExtension(file?.name ?? options.name ?? '') || 'מסמך חדש';
  // שני הדברים שהמסמך הזה מביא איתו, יחד: מה יש בו, ותחת איזו סיומת הוא
  // נשמר. הסיומת נגזרת מהשם **ומהחבילה** — מסמך `.docx` שנושא חלק מאקרו
  // (קורה) יישמר כ-`.docm`, אחרת Word יתלונן עליו.
  documentVba.value = vba;
  saveExtension.value = resolveSaveExtension(file?.name ?? options.name, vba.hasMacroPart);
  // זמן הטעינה הוא מדידת פיתוח ולא הודעה למשתמש: „נטען ב-473 מילישניות” תפס
  // את שורת המצב עד ההודעה הבאה. הוא נשמר — הוא מה שמסביר פתיחה איטית —
  // בלוג של אוצריא, במקום שבו מסתכלים על מדידות.
  console.info(
    `[otzaria-word] ${title.value} נטען ב-${Math.round(performance.now() - startedAt)} מילישניות`
  );
  setStatus('');

  /* שורת המצב מציגה הודעה אחת, ולכן יש סדר עדיפות בין אלה שיכולות להיאמר
     כאן. „יש במסמך מאקרו שWord מריץ בפתיחה” קודמת ל„פתוח לקריאה”: הראשונה היא
     מה שהמשתמש צריך לדעת על הקובץ שהוא פתח, והשנייה תתגלה לו בעצמה ברגע
     שילחץ „שמור”. ספירת מודולים בלבד — הפחות דחוף — נאמרת רק כשאין השתקה
     אחרת.

     הודעת השלב המקדים נכנסה **שנייה**, אחרי המאקרו ולפני „פתוח לקריאה”, ולא
     כי היא דחופה יותר מהן: „פתוח לקריאה” המשתמש יגלה בעצמו בשמירה הראשונה,
     ואת העובדה שהמסמך שלו **שונה** הוא אינו יכול לגלות בשום דרך — לא ברצועה,
     לא בביטול, ולא בקובץ. ראו COMPLEX_SCRIPT_BOLD_NOTICE. */
  const readOnlyNotice =
    file && file.access !== 'readwrite'
      ? `${title.value} — פתוח לקריאה; „שמור” יבקש מקום חדש`
      : null;
  if (vba.autoRun.length > 0 && vba.status) setStatus(vba.status);
  else if (preflightNotice) setStatus(preflightNotice);
  else if (readOnlyNotice) setStatus(readOnlyNotice);
  else if (vba.status) setStatus(vba.status);

  save?.reset(file && file.access === 'readwrite' ? { token: file.token, name: file.name } : null);

  if (options.remember !== false) {
    keeper?.setDocument(
      file ? { token: file.token, name: file.name, writable: file.access === 'readwrite' } : null,
      { sourceSize: file?.size ?? null, keepDraft: options.draft !== undefined },
    );
    // אותו תנאי בדיוק כמו הזוכר שמעל: `remember: false` הוא „אל תרשום את
    // הפתיחה הזאת בשום מקום”, וזה כולל את רשימת האחרונים. מסמך חדש (`!file`)
    // אין לו token, ואין מה לרשום.
    if (file) {
      rememberRecentDocument({
        token: file.token,
        name: file.name,
        size: file.size,
        writable: file.access === 'readwrite',
      });
    }
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
    //
    // כשהתבנית נפתחה ההגדרות כבר בקובץ, וקריאה אחת מאמתת זאת: כל כתיבה כאן
    // עולה בבנייה חוזרת של אינדקס הביקורת במנוע (~200ms לכל שכבה, נמדד).
    if (blank && (await verifyHebrewDocumentDefaults(editor.superdoc))) {
      document.documentElement.dataset.documentDirection = 'rtl';
    } else {
      await applyNewDocumentPaperSize(editor.superdoc);
      await applyNewDocumentDirection(editor.superdoc);
    }
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
  const caretRestored = await restoreDocumentView(editor, adapter, options.restore);

  /*
   * מסמך שנפתח בלי מקום שמור נפתח בלי סמן בכלל: `focus()` של המנוע ממקד את
   * המקלדת אבל אינו מציב סמן כשאין בחירה קודמת, וכל הקלדה נבלעת עד קליק
   * (ראו applyDocumentStartCaret). לכן — סמן בתחילת המסמך, כמו Word.
   *
   * אותם שערים כמו `focusOpenedDocument`, ומאותו טעם; ועוד אחד: בחירה
   * שהמשתמש כבר הספיק להציב בקליק בזמן שהפתיחה רצה אינה נדרסת.
   */
  if (
    !caretRestored &&
    !isOutsideDocumentEditing(document.activeElement) &&
    !hasTextCaret(editor.ui)
  ) {
    await applyDocumentStartCaret(editor.ui, editor.superdoc);
  }

  // אחרון, ולא ליד `outcome.status === 'opened'`: מרגע ש-100% על המסך „דלג”
  // נעלם, ואילו שחזור התצוגה הוא עוד המתנה שהמשתמש רואה.
  attempt.finish();
  return true;
}

/**
 * „דלג” בשורת המצב.
 *
 * שני הקוראים חייבים להיקרא, וכל אחד מהם לבד אינו מספיק: `documentLoad.cancel`
 * מסלק את המחוון ומסמן את הפתיחה כנטושה — כלומר `openDocument` לא תכתוב יותר
 * למצב המעטפת — ו-`swap.cancel` הוא זה שמפרק את המנוע החצי-בנוי ומשחרר את
 * ה-workers שלו. בלי השני „דלג” היה מסתיר את הפס בזמן שהמכונה ממשיכה לבנות
 * מסמך שאיש לא יראה.
 *
 * מה שאין לו כאן כפתור: חוט ראשי שנתקע בלולאה של המנוע. שם גם הלחיצה עצמה
 * אינה מגיעה, וההגנה היחידה היא לפני שהמנוע רואה את הבייטים
 * (engine/docx-preflight.ts).
 */
async function onSkipLoad(): Promise<void> {
  const name = loadSnapshot.value.name;
  // התנאי הוא של המודל ולא של המעטפת: הוא זה שיודע אם יש פתיחה שאפשר לנטוש,
  // וכפתור שנלחץ פעמיים אינו אמור לפתוח מסמך ריק שני.
  if (!documentLoad.cancel()) return;
  swap?.cancel();
  isOpening.value = false;

  if (hasDocument.value) {
    setStatus(`פתיחת ${name} הופסקה — ${title.value} נשאר פתוח`);
    return;
  }

  // אין לְמה לחזור. מסמך ריק עדיף על מסך ריק — אותה הכרעה בדיוק כמו בכשל
  // פתיחה ב-`reopenPreviousSession`, ומאותו טעם: המשתמש נשאר עם עורך שאפשר
  // לעבוד בו. `remember: false` הוא העיקר כאן — הרשומה ממשיכה להצביע על
  // המסמך שנטש, וההפעלה הבאה תנסה אותו שוב.
  await openDocument(undefined, { remember: false });
  setStatus(`פתיחת ${name} הופסקה — נפתח מסמך חדש`);
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
 *
 * מחזירה האם סמן הוצב — הקורא מציב סמן פתיחה בתחילת המסמך רק כשלא.
 */
async function restoreDocumentView(
  editor: EditorSession,
  adapter: CommandAdapter,
  restore: OpenOptions['restore'],
): Promise<boolean> {
  if (!restore) return false;

  if (restore.zoom !== null) {
    // `adapter` (של הטאב הזה, נקבע ב-openDocumentInto) ולא `commandAdapter.value`
    // הסינגלטון: הפתיחה כבר הפכה א-סינכרונית ב-`isOpening` (ראו „ריבוי
    // מסמכים” ליד `sessions`), ולכן ייתכן שהטאב עבר לרקע ממש כאן — קריאה
    // מהסינגלטון הייתה מריצה „זום” על הטאב הפעיל האחר במקום על זה שנפתח.
    const outcome = await adapter.run('zoom', zoomPayload(restore.zoom));
    if (outcome && !outcome.ok) {
      console.info(`[otzaria-word] גודל התצוגה השמור לא הוחזר: ${outcome.message}`);
    }
  }

  if (!restore.caret) return false;
  const caretApplied = await applyCaretAnchor(editor.ui, editor.superdoc, restore.caret);
  if (!caretApplied) {
    console.info('[otzaria-word] מקום הסמן השמור לא נמצא במסמך שנפתח');
  }
  return caretApplied;
}

async function onSave(forceSaveAs = false): Promise<void> {
  if (!swap?.current || !save) return;
  // כמו כל שאר מתחילי הפעולה בקובץ (`onPickAndOpen`, `onNewDocument`,
  // `onDocumentTabSelect`): בזמן פתיחה `swap.current` הוא עדיין המסמך
  // **הקודם**, ושמירה כאן הייתה מייצאת אותו — ועל מסמך חדש היא גם פותחת
  // „שמור בשם” על מה שעומד להימחק בעוד רגע. אין כאן „לחיצה שנבלעת”: מסך
  // הטעינה על המסך, והמשתמש רואה שהתוסף עסוק.
  if (isOpenBusy()) return;
  // ההצעה נושאת את הסיומת: זה מה שהמשתמש רואה בדיאלוג „שמור בשם”, וכאן נקבע
  // אם מסמך המאקרו שלו יישאר `.docm`.
  const outcome = await save.saveNow({
    forceSaveAs,
    suggestedName: documentFileName(title.value, saveExtension.value),
  });

  if (outcome.status === 'failed') {
    setStatus(outcome.message, true);
    return;
  }
  if (outcome.status === 'saved') {
    title.value = stripWordExtension(outcome.name) || title.value;
    // זוכר-ההפעלה אינו מעודכן כאן אלא ב-`onSaved` של הקואורדינטור: שם עוברות
    // גם השמירה האוטומטית וגם „לשמור לפני שפותחים אחר”, שאינן עוברות כאן.
    setStatus(`${title.value} נשמר`);
  }
}

/**
 * האם הטאב הפעיל „ריק במובן שאפשר להחליף” — בלי קובץ ובלי עריכה. פתיחת מסמך
 * עליו מחליפה אותו במקום, כמו ב-VSCode/דפדפנים; אחרת נפתח טאב נוסף.
 * `targetToken === null` ולא רק „לא dirty”: מסמך שכבר נשמר לקובץ ונקי כרגע
 * הוא עדיין מסמך אמיתי של המשתמש, ופתיחה נוספת לא אמורה לדרוס אותו במקום.
 */
function isActiveTabReplaceable(): boolean {
  return !saveSnapshot.value.isDirty && saveSnapshot.value.targetToken === null;
}

/**
 * מכינה טאב לפתיחה: אם הפעיל אינו „ריק”, פותחים טאב חדש ומפעילים אותו —
 * ואז ל-`decideDocumentSwitch` שבהמשך `onPickAndOpen`/`onNewDocument` אין
 * מה להחליט (הטאב החדש תמיד נקי), אבל הקריאה אליה עדיין רצה, בלי תנאי מיוחד.
 */
function ensureOpenTargetTab(): void {
  if (sessions.size > 0 && !isActiveTabReplaceable()) {
    activateTab(createNewDocumentSession());
  }
}

/**
 * „פתח מסמך”: מה שקורה סביב הדיאלוג.
 *
 * ## הדיאלוג אינו מסלול פתיחה נוסף
 *
 * שני הכפתורים הוותיקים נשארו בדיוק כפי שהיו, והדיאלוג רק החליף את הדרך
 * אליהם: „עיון בקבצים…” קורא ל-`onPickAndOpen`, וכרטיס תבנית קורא
 * ל-`onNewDocument`. כלומר „המסמך לא נשמר”, בחירת הטאב והשמירה-לפני-החלפה
 * ממשיכים לעבור בקוד אחד — ולא בעותק שני שיכול להתפצל בשקט.
 *
 * המסלול היחיד שהוא באמת חדש הוא פתיחה משורת „אחרונים”, ובו אין בורר קבצים
 * אלא token שנפתר. גם הוא אינו מוחק עבודה: `ensureOpenTargetTab` פותח **טאב
 * חדש** כשהפעיל אינו „ריק במובן שאפשר להחליף” (`isActiveTabReplaceable`),
 * ולכן טאב מלוכלך מקבל שכן ואינו נדרס — בדיוק כמו בדפדפן. אחריו הטאב שאליו
 * פותחים נקי תמיד, וזו הסיבה שאין שם שאלה שנייה לשאול.
 */
function openOpenDialog(): void {
  /*
   * שקט הוא לא תשובה. הכפתור ברצועה מנוטרל בזמן פתיחה ויש לו tooltip, אבל
   * `Ctrl+N`/`Ctrl+O` מגיעים לכאן דרך מפעיל הפעולות — שמדווח „טופל” ובולע
   * את ברירת המחדל של הדפדפן. כלומר מי שפותח docx גדול ולוחץ `Ctrl+O` ראה
   * מסך שאינו מגיב ושורת מצב שותקת, בניגוד להצהרה שבראש
   * ui/shortcuts/registry.ts: „קיצור שנכשל מדבר עברית בשורת המצב בדיוק כמו
   * כפתור שנכשל”. אותה הכרעה בדיוק כמו בענף השמירה שמתחת.
   */
  if (isOpenBusy()) {
    setStatus('הפתיחה עוד רצה — רגע אחד');
    return;
  }
  /*
   * שמירה שרצה עוצרת את הפתיחה **לפני** שהדיאלוג נפתח, ולא אחריה.
   *
   * ה-prop `busy` הוא `isOpening || saveSnapshot.isSaving`, ולכן דיאלוג
   * שנפתח בזמן שמירה נפתח כשכל הכרטיסים וכל השורות `disabled` — מסך שלם
   * שאי-אפשר ללחוץ בו על דבר, בלי שנאמר למה. גרוע מזה: המיקוד ההתחלתי הולך
   * לכרטיס הראשון, ומיקוד על כפתור מנוטרל הוא no-op — כלומר המיקוד נשאר
   * מאחורי המודאל ומלכודת ה-Tab אינה יכולה לתפוס אותו.
   *
   * ההודעה היא זו שכבר קיימת לאותו מצב בדיוק ב-`onPickAndOpen`.
   */
  if (saveSnapshot.value.isSaving) {
    setStatus('השמירה עוד רצה — רגע אחד');
    return;
  }
  recentSearch.value = '';
  isOpenDialogOpen.value = true;
}

async function onOpenDialogBrowse(): Promise<void> {
  isOpenDialogOpen.value = false;
  // המסלול המלא הקיים, על כל השאלות שבו — הדיאלוג רק החליף את הדרך אליו.
  await onPickAndOpen();
}

/**
 * „מסמך חדש מתבנית”.
 *
 * התבנית מוחלת **אחרי** שהמסמך נפתח, ולא לפניו: `applyTemplate` עובד על מופע
 * מנוע חי (`engine/templates.ts`), ואין מסמך לפני `openDocument()`.
 *
 * שלוש התוצאות מדווחות שונה, ובכוונה: כשל הוא שגיאה אדומה; `note` (למשל
 * „הטורים מצוירים הפוך”) הוא הודעה רגילה שאינה כשל; והצלחה שקטה אינה מדווחת
 * כלל — המסמך על המסך הוא הדיווח.
 */
async function onOpenDialogCreate(id: string): Promise<void> {
  isOpenDialogOpen.value = false;
  // המסלול המלא הקיים, ולא פתיחה מקבילה: `onNewDocument` הוא זה שמחזיק את
  // „המסמך לא נשמר” למקרה שבו אין טאב חדש לפתוח אליו.
  if (!(await onNewDocument())) return;

  const templateId = id as TemplateId;
  if (templateId === 'blank') return;

  const outcome = await applyTemplate(activeSuperdoc.value, templateId);
  if (!outcome.ok) {
    setStatus(outcome.message, true);
    return;
  }
  if (outcome.note) setStatus(outcome.note);
}

/**
 * פתיחת מסמך מרשימת האחרונים.
 *
 * ה-token שורד הפעלות אבל ה-URL לא (הפורט של שרת ה-loopback מתחלף), ולכן
 * הפתיחה עוברת דרך `resolveFileUrl`. `null` ממנו פירושו שהקובץ הוזז, נמחק או
 * שההרשאה בוטלה — ואז השורה **מוסרת מהרשימה**: שורה שנכשלת ונשארת היא שורה
 * שהמשתמש ילחץ עליה שוב.
 */
/**
 * האם פתיחה משורת „אחרונים” כבר בדרך.
 *
 * `isOpenBusy()` אינו מכסה את החלון הזה: הוא נדלק רק כש-`openDocument`
 * מתחיל, ולפניו יש `await resolveFileUrl` שלם — סבב IPC מלא מול אוצריא.
 * לחיצה על שורה שנייה בתוכו הייתה מוציאה פתיחה שנייה, ושתיהן היו כותבות
 * לאותו טאב.
 */
let recentOpenPending = false;

async function onOpenDialogRecent(token: string): Promise<void> {
  if (isOpenBusy() || recentOpenPending) return;
  const known = recentDocuments.value.find((item) => item.token === token);
  const name = known?.name ?? 'המסמך';

  recentOpenPending = true;
  let file: UserFile | null;
  try {
    file = await resolveFileUrl(token);
  } finally {
    recentOpenPending = false;
  }

  if (!file) {
    onOpenDialogForget(token);
    setStatus(`${name} לא נמצא — ייתכן שהקובץ הוזז או נמחק. הוא הוסר מהרשימה`, true);
    return;
  }

  // הדיאלוג נסגר בזמן ההמתנה — Escape, „סגור”, או לחיצה על הרקע. סגירה היא
  // ביטול, ופתיחה שממשיכה אחריה מחליפה למשתמש את המסמך אחרי שהוא כבר אמר
  // „לא”. אותו כלל בדיוק כמו `attempt.cancelled` ב-`openDocumentInto`.
  if (!isOpenDialogOpen.value) return;

  isOpenDialogOpen.value = false;
  ensureOpenTargetTab();
  /*
   * ה-`access` מורכב ממה שנשמר, ואינו מגיע מהגשר.
   *
   * `fs.resolveFileUrl` מחזירה `{token, url, name, size}` **בלי** `access`
   * (host/files.ts) — השדה מתועד על בורר הקבצים, לא עליה. בלי ההרכבה כאן
   * הקובץ נפתח כקריאה-בלבד: `save.reset` לא מקבל יעד כתיבה, המשתמש מקבל
   * „פתוח לקריאה” על קובץ שהוא כן יכול לכתוב אליו, ו-`keeper.setDocument`
   * **כותב את זה לרשומת ההפעלה** — כך שהמצב שורד הפעלות והופך קבוע.
   *
   * זה בדיוק מה ש-`resolveRememberedFile` עושה למסלול שחזור ההפעלה, ומאותה
   * סיבה; ההבדל היחיד הוא מאיפה מגיע הדגל — שם מרשומת ההפעלה, כאן מרשומת
   * „האחרונים”. שורה שנשמרה לפני שהשדה קיים נקראת כקריאה-בלבד, וזו הנפילה
   * הבטוחה: „שמור בשם” על קובץ כתיב מטריד, כתיבה ל-token בלי הרשאה נכשלת.
   */
  await openDocument({ ...file, access: known?.writable ? 'readwrite' : 'read' });
}

function onOpenDialogTogglePin(token: string, pinned: boolean): void {
  recentDocuments.value = setRecentPinned(recentDocuments.value, token, pinned);
  void saveRecentDocuments(recentDocuments.value);
}

/**
 * „נסגרו בלי לשמור” — המעבר ממסך הפתיחה למסך השחזור.
 *
 * הדיאלוג הראשון נסגר לפני שהשני נפתח, ולא נשאר מתחתיו: שניהם מכריזים
 * `aria-modal` ומחזיקים מלכודת מיקוד, ושתי מלכודות פתוחות בו-זמנית מתחרות על
 * אותו `Tab`. „סגור” במסך השחזור מחזיר את המשתמש לעורך, לא למסך הפתיחה —
 * מסך שנכנסים אליו כדי לפתוח משהו, וסגירתו פירושה שנמלכו בדעתם.
 */
function onShowDiscarded(): void {
  isOpenDialogOpen.value = false;
  isDiscardedOpen.value = true;
}

/**
 * פתיחת גיבוי — **כמסמך חדש**, ולא מעל הקובץ שהוא בא ממנו.
 *
 * הרשומה יודעת מאיזה `token` העבודה הזאת באה, והפיתוי היה לפתוח את הקובץ
 * ולהצמיד אליו את הבייטים — „שמור” אחד והכול חוזר למקומו. זה בדיוק מה שאסור:
 * המשתמש אמר על העבודה הזאת „לא לשמור”, ופתיחה שמצמידה אותה ליעד כתיבה הופכת
 * לחיצה אחת בשוגג („שמור” מתוך הרגל) לדריסת הקובץ שלו בגרסה שהוא דחה. בלי
 * יעד, „שמור” פותח „שמור בשם” — והמשתמש בוחר במפורש לאן.
 *
 * הרשומה **אינה** מוסרת אחרי פתיחה: מה שנפתח עדיין לא נשמר לשום מקום, ומחיקת
 * העותק ברגע הזה הייתה משאירה את העבודה שוב בלי רשת. „הסר” הוא פעולה מפורשת
 * (`onForgetDiscarded`), וזו ההזדמנות היחידה למחוק אותה.
 */
async function onOpenDiscarded(slot: number): Promise<void> {
  if (isOpenBusy() || isDiscardedOpening.value) return;
  const entry = discardedBackups.value.find((item) => item.slot === slot);
  if (!entry) return;

  isDiscardedOpening.value = true;
  try {
    const bytes = await readWorkspaceBytes(backupPathFor(slot));
    if (!bytes) {
      // הקובץ אינו במקומו — נמחק מבחוץ, או שהכתיבה שלו נכשלה מלכתחילה. שורה
      // שאי אפשר לפתוח היא שורה שהמשתמש ילחץ עליה שוב, ולכן היא יורדת.
      await forgetDiscarded(slot);
      setStatus(`העותק של ${entry.name} לא נמצא, והשורה הוסרה`, true);
      return;
    }

    // „סגור” מותר גם בזמן הקריאה. במקרה כזה אין לפתוח מסמך אחרי שהמשתמש
    // ביטל במפורש את הפעולה.
    if (!isDiscardedOpen.value) return;

    isDiscardedOpen.value = false;
    ensureOpenTargetTab();
    const opened = await openDocument(undefined, {
      draft: new Blob([bytes], { type: DOCX_MIME }),
      name: entry.name,
    });
    if (!opened) {
      // הרשומה **אינה** יורדת כאן, בשונה מ„העותק לא נמצא”: הבייטים קיימים,
      // והכשל עשוי להיות זמני (worker שלא עלה). מחיקת השורה הייתה מוחקת את
      // ההזדמנות היחידה לנסות שוב.
      setStatus(`${entry.name} לא נפתח מהעותק שנשמר בסגירה`, true);
      return;
    }

    const age = draftAgeLabel(entry.discardedAt, Date.now());
    setStatus(
      age
        ? `${entry.name} נפתח מהעותק שנשמר בסגירה (${age}) — טרם נשמר לקובץ`
        : `${entry.name} נפתח מהעותק שנשמר בסגירה — טרם נשמר לקובץ`,
    );
  } finally {
    isDiscardedOpening.value = false;
  }
}

/**
 * „הסר”. מוחקת גם את הקובץ ולא רק את השורה: המשבצת הייתה מתפנה בלאו הכי
 * והכתיבה הבאה דורסת אותה, אבל מי שביקש להסיר עותק של העבודה שלו התכוון
 * שהוא לא יישאר על הדיסק עד שמישהו יזדמן לדרוס אותו.
 */
async function onForgetDiscarded(slot: number): Promise<void> {
  if (isDiscardedOpening.value) return;
  await forgetDiscarded(slot);
}

function onOpenDialogForget(token: string): void {
  recentDocuments.value = forgetRecent(recentDocuments.value, token);
  void saveRecentDocuments(recentDocuments.value);
}

async function onPickAndOpen(): Promise<void> {
  if (isOpenBusy()) return;
  try {
    const file = await pickDocxFile();
    if (!file) return;

    ensureOpenTargetTab();

    if (save && swap) {
      // נקרא **לפני** ההחלטה: אחרי „לשמור קודם” המסמך כבר נקי, ואז אי אפשר
      // להבחין בין „לא היה מה למחוק” לבין „המשתמש ביקש למחוק”.
      const hadUnsaved = save.snapshot.isDirty;
      const decision = await decideDocumentSwitch({
        isDirty: () => save!.snapshot.isDirty,
        isSaving: () => save!.snapshot.isSaving,
        ask: unsavedPrompt.ask,
        documentName: () => title.value,
      });

      /*
       * שומר, ולא מסלול חי — וזה נאמר כאן כדי שהקורא הבא לא יחפש אותו.
       *
       * `ensureOpenTargetTab` שלמעלה רץ **לפני** ההחלטה, והוא פותח טאב חדש
       * בכל פעם שהפעיל אינו „ריק במובן שאפשר להחליף” — כלומר גם כשהוא
       * מלוכלך. מכאן והלאה הטאב שאליו פותחים נקי תמיד, `hadUnsaved` הוא
       * `false`, ו-`decideDocumentSwitch` מקצר על `!isDirty()` בלי לשאול
       * דבר. השאלה על עבודה שלא נשמרה חיה היום בסגירת טאב וביציאה בלבד.
       *
       * לא נמחק בכוונה: הוא הקוד שיציל את המסלול הזה אם „פותחים לטאב חדש”
       * ישתנה אי-פעם, והוא עולה שורת תנאי אחת.
       */
      if (hadUnsaved && decision.action === 'switch') {
        await discardWithBackup(activeSession.value);
      }

      if (decision.action === 'cancel') {
        setStatus(
          decision.reason === 'saving'
            ? 'השמירה עוד רצה — רגע אחד'
            : 'הפתיחה בוטלה, והמסמך נשאר פתוח'
        );
        return;
      }

      if (decision.action === 'save-first') {
        const outcome = await save.saveNow({ suggestedName: documentFileName(title.value, saveExtension.value) });
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

/**
 * „מסמך חדש”.
 *
 * מחזירה **האם מסמך אכן נפתח**, ולא `void`: `onOpenDialogCreate` צריך לדעת
 * את זה כדי להחליט אם להחיל תבנית — החלה על מסמך שלא נפתח (הפתיחה בוטלה,
 * השמירה נכשלה, „לא נשמר” נענה ב„בטל”) הייתה נוגעת במסמך **הקודם**, שנשאר
 * על המסך. זו הסיבה היחידה שהיא מחזירה ערך.
 */
async function onNewDocument(): Promise<boolean> {
  // כמו שאר מתחילי הפתיחה (onPickAndOpen/onDocumentTab*): הכפתור ברצועה
  // מנוטרל בזמן isOpening, אבל Ctrl+N מגיע דרך runShellAction ועוקף אותו —
  // בלי הבדיקה הזאת הוא היה פותח מסמך שני לתוך אותו טאב באמצע פתיחה ראשונה.
  if (isOpenBusy()) return false;
  ensureOpenTargetTab();
  if (save && swap && save.snapshot.isDirty) {
    const decision = await decideDocumentSwitch({
      isDirty: () => save!.snapshot.isDirty,
      isSaving: () => save!.snapshot.isSaving,
      ask: unsavedPrompt.ask,
      documentName: () => title.value,
    });
    // ראו onPickAndOpen: גם כאן זה שומר ולא מסלול חי — `ensureOpenTargetTab`
    // כבר העביר אותנו לטאב נקי, ולכן הבלוק הזה כולו אינו נכנס היום.
    if (decision.action === 'switch') await discardWithBackup(activeSession.value);
    if (decision.action === 'cancel') return false;
    if (decision.action === 'save-first') {
      const outcome = await save.saveNow({ suggestedName: documentFileName(title.value, saveExtension.value) });
      if (outcome.status !== 'saved') return false;
    }
  }
  return openDocument();
}

/**
 * „יציאה”: סוגרת את המסמכים הפתוחים, ומחזירה את המשתמש למסך הספרייה.
 *
 * מה „יציאה” אומרת כאן, וזה אינו מובן מאליו: התוסף הוא לשונית בתוך אוצריא,
 * ולא אפליקציה שנסגרת. `navigation.goTo` מוציא את המשתמש מהמסך, ואוצריא
 * **משהה** את ה-WebView (`plugin.suspended`) במקום להרוס אותו — כלומר בלי
 * סגירה מפורשת המסמך ממשיך לחכות פתוח כפי שהיה.
 *
 * זה מה שהכפתור עשה עד כאן, וזו הייתה סתירה שהמשתמש פוגש בלחיצה אחת: הוא
 * נשאל „לשמור לפני יציאה?”, ואם ענה „לא” גם „לצאת בלי לשמור? השינויים
 * יימחקו” — ואז דבר לא נסגר ודבר לא נמחק. כפתור ששאלותיו מתארות פעולה שאינה
 * קורית מלמד לא להאמין לשאלות שלו, וזו בדיוק השאלה שאסור שילמדו להתעלם ממנה.
 *
 * לכן היציאה סוגרת. „פתח ספרייה” בלשונית „אוצריא” נשאר מה שהיה — כפתור ניווט
 * שאינו סוגר ואינו שואל — ועכשיו ההבדל בין שני הכפתורים הוא בדיוק ההבדל
 * שהתוויות שלהם מבטיחות.
 *
 * השאלה אינה מחייבת לשמור: „לא לשמור” הוא אחד משלושת הכפתורים, והוא יוצא
 * ומשאיר עותק לשחזור — ראו sessions/open-flow.ts ו-sessions/discard-backup.ts.
 */
async function onExit(): Promise<void> {
  // בזמן פתיחה אין לסגור: `openDocumentInto` כותב לטאב הפעיל לכל אורך ריצתו
  // (ראו `isOpenBusy`), וסגירתו באמצע הייתה משאירה אותה כותבת לתוך מסמך
  // מפורק. הפקד עצמו מנוטרל אז; זה הגיבוי למסלול שיגיע שלא דרכו.
  if (isExiting.value || isOpenBusy()) return;
  isExiting.value = true;
  try {

  // הטאב הפעיל נשאל ראשון: הוא המסמך שעל המסך כשלוחצים „יציאה”, ושאלה על
  // מסמך שאינו נראה, לפניו, נקראת כשאלה עליו.
  const active = activeSession.value;
  const open = Array.from(sessions.values());
  const order = active ? [active, ...open.filter((session) => session !== active)] : open;

  /*
   * כל הטאבים נשאלים **לפני** שנסגר ולו אחד: „ביטול” על השלישי אחרי ששניים
   * כבר נסגרו הוא ביטול שאינו מבטל.
   *
   * מה ש„ביטול” כן משאיר מאחוריו, ונאמר כאן במפורש: **שמירות שכבר בוצעו**.
   * „שמור” על הטאב הראשון כותב לדיסק מיד, ו„ביטול” על השלישי אינו מחזיר את
   * הכתיבה הזאת — ואינו אמור: המשתמש ביקש לשמור את המסמך הראשון, וקיבל את
   * מה שביקש. מה שבוטל הוא היציאה, לא השמירה. „ביטול” על סגירה אף פעם אינו
   * הורס עבודה, וזו התכונה היחידה שחייבת להתקיים כאן.
   */
  const discarded = new Set<DocumentSessionId>();
  for (const session of order) {
    const resolution = await resolveUnsavedBeforeClose(session, 'exit');
    if (!resolution.closable) return;
    if (resolution.discarded) discarded.add(session.id);
  }

  await closeAllSessions(discarded);

  // אותו מסלול דיווח כמו „פתח ספרייה” בלשונית „אוצריא”: הודעה בעברית למשתמש
  // ושורה בלוג של אוצריא. כשל ניווט אינו מחזיר את המסמכים — כל אחד מהם כבר
  // נשמר או שמחיקתו אושרה במפורש, ומי שנשאר בתוסף נשאר עם עורך נקי והודעה
  // שאומרת מה נכשל.
    reportReader(await openLibrary());
  } finally {
    isExiting.value = false;
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
    (input) => call('ui.exportPdf', { ...input }),
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
 * ייצוא לפורמט ספר של אוצריא — טקסט עם רמות כותרות. הבנייה ב-
 * engine/otzaria-book.ts; כאן רק מסלול השמירה והדיווח.
 *
 * המסלול הוא מסלול השמירה הבינארית הרגיל (begin ← PUT ← commit עם דיאלוג
 * „שמור בשם”), עם סיומת `txt` — ולא הורדה דרך `<a download>`: תיקיית
 * ההורדות הייתה מפספסת את הנקודה, שהיא לשים את הקובץ בתיקייה אישית של
 * אוצריא. (זה גם מה שהוריד את „ייצוא ל-Word” מלשונית „קובץ” — ראו FileTab.vue.)
 *
 * אחרי שמירה מוצלחת נקרא `library.refreshUserBooks`: אם המשתמש שמר
 * לתיקייה אישית רשומה, הספר נקלט מיד; אחרת הסריקה פשוט לא תמצא דבר,
 * וההודעה אומרת מה אפשר לעשות. הרענון הוא best-effort (`tryCall`) —
 * היעדר ההרשאה או Host ישן אינם כשל של הייצוא, שכבר הצליח.
 *
 * **ולא בהמתנה.** הרענון סורק מחדש את כל התיקיות האישיות ומרענן את קטלוג
 * הספרייה, ולאוצריא יש עליו timeout של 15 דקות (`_userBooksRefreshTimeout`).
 * המתנה לו לפני ההודעה הייתה משאירה את המשתמש בלי שום אישור על קובץ שכבר
 * נכתב לדיסק — כמה זמן שהסריקה תימשך. לכן: „נשמר” מיד, והתוצאה מעדכנת את
 * השורה כשהיא חוזרת — ורק אם בינתיים לא נאמר שם משהו חדש יותר.
 */
async function onExportOtzaria(): Promise<void> {
  if (!swap?.current) {
    setStatus('אין מסמך פתוח לייצוא', true);
    return;
  }

  const built = await buildOtzariaBook(activeSuperdoc.value, title.value);
  if (!built.ok) {
    setStatus(built.message, true);
    return;
  }

  let ticket: WriteTicket | null = null;
  try {
    const blob = new Blob([built.text], { type: 'text/plain' });
    ticket = await beginBinaryWrite(blob.size);
    await uploadBytes(ticket.uploadUrl, blob);
    const result = await commitUserFileWrite({
      writeToken: ticket.writeToken,
      suggestedName: otzariaBookFileName(title.value),
      title: 'ייצוא לספר אוצריא',
      extension: 'txt',
    });
    // ה-commit צרך את ההעלאה — מכאן אין מה לבטל, גם אם הדיווח ייכשל.
    ticket = null;
    if (result.cancelled) {
      setStatus('הייצוא בוטל');
      return;
    }

    const name = result.name ?? 'הספר';
    const saved = `${name} נשמר — מרענן את ספריית אוצריא…`;
    setStatus(saved);

    void tryCall<{ addedBooks?: number; updatedBooks?: number }>(
      'library.refreshUserBooks',
      {},
    ).then((refreshed) => {
      // השורה מעודכנת רק אם היא עוד שלנו: בזמן הסריקה המשתמש המשיך לעבוד,
      // ודריסת „נשמר” או הודעת שגיאה חדשה בתוצאה של רענון היא בדיוק ההפתעה
      // שאסור לייצר.
      if (statusText.value !== saved) return;
      const landed = (refreshed?.addedBooks ?? 0) + (refreshed?.updatedBooks ?? 0) > 0;
      // הנוסח כשאין ספר חדש הוא מותנה, ולא הוראה: אפס נספרים פירושו גם
      // „לא בתיקייה רשומה” וגם „אותו ספר בדיוק כבר שם” (ייצוא חוזר שדורס
      // קובץ זהה), ואי אפשר להבחין ביניהם מהתשובה.
      setStatus(
        landed
          ? `${name} נשמר ונקלט בספריית אוצריא`
          : `${name} נשמר — אם אינו מופיע בספרייה, ודאו שהתיקייה רשומה בהגדרות אוצריא`,
      );
    });
  } catch (error) {
    if (ticket) void abortBinaryWrite(ticket.writeToken);
    setStatus(
      `הייצוא לספר אוצריא נכשל: ${error instanceof Error ? error.message : String(error)}`,
      true,
    );
  }
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

/** הכותרת המוצגת של טאב — מהתצוגה החיה אם הוא הפעיל, אחרת מתמונת המצב שלו. */
function sessionDisplayTitle(session: DocumentSession): string {
  return session === activeSession.value ? title.value : session.ui.title;
}

function sessionIsDirty(session: DocumentSession): boolean {
  return session === activeSession.value ? saveSnapshot.value.isDirty : session.ui.saveSnapshot.isDirty;
}

/**
 * הסיומת שתחתיה הטאב הזה נשמר — מהתצוגה החיה אם הוא הפעיל, אחרת מתמונת
 * המצב שלו. אותו דפוס כמו `sessionDisplayTitle`, ומאותו טעם: קואורדינטור
 * השמירה של טאב ברקע רץ על **המסמך שלו**, ו-`saveExtension.value` מתאר את
 * המסמך שמזדמן להיות מוצג באותו רגע.
 */
function sessionSaveExtension(session: DocumentSession): WordExtension {
  return session === activeSession.value ? saveExtension.value : session.ui.saveExtension;
}

/** רצועת הטאבים האמיתית — אחד לכל `DocumentSession` פתוח, לפי סדר הפתיחה. */
const documentTabs = computed<DocumentTabItem[]>(() => {
  // התלות המפורשת, ראו `tabStripRevision`: שני השדות שמתחת נקראים על טאב
  // ברקע מאובייקט פשוט, ובלי השורה הזאת שינוי בו אינו מרנדר מחדש.
  void tabStripRevision.value;
  return Array.from(sessions.values()).map((session) => ({
    id: session.id,
    title: sessionDisplayTitle(session),
    isDirty: sessionIsDirty(session),
  }));
});

/**
 * מעבר טאב. חסום בזמן פתיחה: `openDocumentInto` כותב סינכרונית לתוך הטאב
 * הפעיל לכל אורך ריצתה (ראו ההערה שם), ומעבר באמצע היה כותב לטאב הלא נכון.
 *
 * טאב ששוחזר ועדיין לא נטען נפתח כאן, ברגע שעוברים אליו — זו „הטעינה
 * העצלה” של `restoreTabs`, וזה הרגע היחיד שבה היא מתרחשת.
 */
function onDocumentTabSelect(id: DocumentSessionId): void {
  if (isOpenBusy()) return;
  const session = sessions.get(id);
  if (!session) return;
  activateTab(session);
  if (session.pendingRestore) {
    void openPendingTab(session).catch((error: unknown) => {
      console.warn('[otzaria-word] טעינת הטאב נכשלה', error);
    });
  }
}

/**
 * פותחת את המסמך של טאב ששוחזר, ברגע שעברו אליו.
 *
 * הדגל מכובה **לפני** ה-`await` הראשון ולא אחריו: הפתיחה מתחילה בפתירת
 * ה-token, ורק אחריה `openDocument` מרימה את `isOpening` שחוסמת מעבר נוסף.
 * בחלון שביניהן לחיצה חוזרת על אותו טאב הייתה מתחילה פתיחה שנייה של אותו
 * מסמך לתוך אותו טאב.
 *
 * `keeper.state` הוא רשומת הטאב הזה בלבד (`sessionForEntry`, אומצה ביצירה),
 * ולכן זו אותה פונקציה בדיוק שפותחת את הטאב הפעיל בעלייה — כולל שחזור
 * הטיוטה, הסמן והזום, וכולל המסלול של קובץ שאינו נגיש יותר.
 */
async function openPendingTab(session: DocumentSession): Promise<void> {
  session.pendingRestore = false;
  /**
   * הנעילה מורמת **לפני** ה-`await` הראשון, ולא נשענת על `isOpening` של
   * `openDocument`.
   *
   * הפתיחה כאן מתחילה בשני סבבי גשר — פתירת ה-token וקריאת בייטי הטיוטה —
   * ורק אחריהם `openDocument` מרימה את הדגל שלה. בחלון הזה שום דבר לא חסם
   * מעבר טאב, ומעבר כזה מחליף את הסינגלטונים ברמת המודול (`restoreFromSession`):
   * המסמך שנפתח היה נוחת בפאנל של הטאב **האחר**, מפרק את המסמך שהיה בו,
   * ומצמיד את יעד השמירה שלו לקובץ הלא נכון. נמדד בביקורת; זה המסלול היחיד
   * בקובץ שבו פתיחה מתחילה בלי שדבר מגן עליה.
   */
  pendingTabLoad = true;
  try {
    // `slept` ולא „ממתין”: טאב ששוחזר מהפעלה קודמת ונטען כאן לראשונה מקבל את
    // הנוסח על ההפעלה הקודמת, וטאב שנרדם באותה הפעלה — את זה שמדבר על עכשיו.
    await reopenPreviousSession(session.keeper.state, session.slept);
    session.slept = false;
  } finally {
    pendingTabLoad = false;
  }

  // והמיקום השמור, עכשיו. `activateTab` כבר קרא ל-`restorePaneScroll` — אבל
  // הוא רץ **לפני** כאן, כשלטאב עוד לא היה host כלל (טאב ששוחזר או שנרדם),
  // ולכן לא היה לו למה לכתוב. בלי הקריאה הזאת המיקום שנשמר בטאב שנרדם נזרק,
  // והוא נפתח בראש המסמך.
  //
  // עטופה בבדיקת „עוד הפעיל”: הפתיחה אורכת שניות, והמשתמש עשוי להיות בטאב
  // אחר מזמן — ואז זו הייתה כתיבה לתוך המסמך הלא נכון.
  if (activeSession.value === session) restorePaneScroll(session, session.ui.scroll);
}

/**
 * מה קורה למה שלא נשמר בטאב, לפני שהוא נסגר. מחזירה האם מותר לסגור: `false`
 * פירושו שהמשתמש ביטל או שהשמירה שביקש נכשלה, ובשני המקרים ההודעה כבר
 * בשורת המצב.
 *
 * משותפת לשני הסוגרים — „×” על טאב בודד ו„יציאה” שסוגרת את כולם — ואינה
 * משוכפלת ביניהם: זהו הקוד שקובע אם עבודה של המשתמש נמחקת, ועותק שני שלו הוא
 * עותק שיכול להתפצל בשקט. אותו טעם שבגללו ההחלטה עצמה יושבת ב-open-flow.ts.
 *
 * ## שתי שאלות שונות — לא זו אחר זו, אלא לפי מה שיש לטאב לאבד
 *
 * המסלול הרגיל שואל את המנוע („יש שינויים שלא נשמרו?”), אבל יש שני מצבים
 * שבהם אין מנוע שיענה, ובכל זאת **יש** עבודה: טאב ששוחזר ועוד לא נטען,
 * וטאב שהטעינה שלו **נכשלה** ונפל לתוכו מסמך ריק (`remember: false`,
 * `reopenPreviousSession`). השני הוא המסוכן: הוא נראה „נקי” לגמרי — המנוע
 * שבתוכו באמת ריק — בעוד הרשומה שלו עדיין מחזיקה טיוטה עם העבודה שלא
 * נשמרה, וסגירתו הייתה מוחקת אותה **בלי אף שאלה** (`destroy({removeDraft})`).
 *
 * לכן השאלה נגזרת מהרשומה ולא מהדגל: כל טאב שיש טיוטה ברשומה שלו והמנוע
 * שבו אינו מציג אותה (כלומר אינו „מלוכלך”) נשאל לפני שמוחקים.
 *
 * ## הטיוטה נמחקת בסגירה, לא כאן
 *
 * „לסגור בלי לשמור” הוא אישור מפורש למחיקת הטיוטה, אבל המחיקה עצמה היא
 * `destroy({ removeDraft: true })` של הסוגר — נקודה אחת ולא שתיים. זה משנה
 * ביציאה: היא שואלת על **כל** הטאבים לפני שהיא סוגרת ולו אחד, ומחיקה כאן
 * הייתה מוחקת את הטיוטה של הראשון גם כשהשאלה על השלישי מבטלת את הכול.
 */
interface UnsavedResolution {
  /** האם מותר לסגור. `false` = המשתמש ביטל, או שהשמירה שביקש נכשלה. */
  closable: boolean;
  /**
   * האם המשתמש בחר „לא לשמור” על טאב שהיה בו מה לאבד — כלומר יש מה לגבות.
   *
   * מדווח ואינו מטופל כאן, מפני ש„נשאל” ו„נסגר” אינם אותו רגע: היציאה שואלת
   * על **כל** הטאבים לפני שהיא סוגרת ולו אחד, וגיבוי שהיה נכתב כאן היה נשאר
   * גם כשהשאלה על הטאב השלישי מבטלת את הכול. הכתיבה עצמה יושבת ליד
   * `destroy({ removeDraft: true })` — ראו `backupDiscardedDocument`.
   */
  discarded: boolean;
}

async function resolveUnsavedBeforeClose(
  session: DocumentSession,
  intent: 'exit' | 'close-tab',
): Promise<UnsavedResolution> {
  const recordDraft = activeEntry(session.keeper.state)?.draft != null;
  const engineDirty = session.save.snapshot.isDirty;
  const askFromRecord = recordDraft && !engineDirty;

  // נקרא **לפני** ההחלטה: אחרי „שמור” המסמך כבר נקי, ואז אי אפשר להבחין בין
  // „לא היה מה לגבות” לבין „המשתמש ביקש למחוק”.
  const hadUnsaved = engineDirty || recordDraft;

  const decision =
    session.pendingRestore || askFromRecord
      ? await decidePendingTabClose({
          hasDraft: () => recordDraft,
          ask: unsavedPrompt.ask,
          documentName: () => sessionDisplayTitle(session),
        })
      : await decideDocumentSwitch({
          isDirty: () => session.save.snapshot.isDirty,
          isSaving: () => session.save.snapshot.isSaving,
          ask: unsavedPrompt.ask,
          documentName: () => sessionDisplayTitle(session),
        });

  /** הפעולה שבוטלה, בנוסח שלה. „היציאה בוטלה” ו„סגירת הטאב בוטלה”. */
  const what = intent === 'exit' ? 'היציאה' : 'סגירת הטאב';

  if (decision.action === 'cancel') {
    setStatus(decision.reason === 'saving' ? 'השמירה עוד רצה — רגע אחד' : `${what} בוטלה`);
    return { closable: false, discarded: false };
  }

  if (decision.action === 'save-first') {
    const outcome = await session.save.saveNow({
      suggestedName: documentFileName(sessionDisplayTitle(session), sessionSaveExtension(session)),
    });
    // שמירה שנכשלה או שבוטלה עוצרת את הסגירה: המשתמש ביקש לשמור, ולסגור בכל
    // זאת היה מתעלם ממה שביקש.
    if (outcome.status !== 'saved') {
      if (outcome.status === 'failed') setStatus(outcome.message, true);
      else setStatus(`${what} בוטלה — המסמך לא נשמר`);
      return { closable: false, discarded: false };
    }
  }

  // `switch` על טאב שהיה בו מה לאבד פירושו ש„לא לשמור” נבחר במפורש. הגיבוי
  // עצמו אינו כאן אלא אצל הסוגר — ראו `discarded` ב-`UnsavedResolution`.
  return { closable: true, discarded: hadUnsaved && decision.action === 'switch' };
}

/** סגירת טאב: השאלה על מה שלא נשמר, ואז הפירוק. */
async function onDocumentTabClose(id: DocumentSessionId): Promise<void> {
  if (isOpenBusy()) return;
  const session = sessions.get(id);
  if (!session) return;
  const resolution = await resolveUnsavedBeforeClose(session, 'close-tab');
  if (!resolution.closable) return;

  // לפני הפירוק: `destroy` משחרר את המנוע ומוחק את הטיוטה, ואחריו כבר אין
  // ממה לקרוא את המסמך שהיה בטאב. ראו `closedTabs`.
  rememberClosedTab(session);

  const wasActive = session === activeSession.value;
  sessions.delete(session.id);
  const recentAt = recentTabs.indexOf(session.id);
  if (recentAt >= 0) recentTabs.splice(recentAt, 1);
  // לפני הפירוק, ורק אחרי שברור שהטאב באמת נסגר: `destroy` מוחק את הטיוטה,
  // והמנוע שממנו מייצאים הולך איתו.
  if (resolution.discarded) await backupDiscardedDocument(session);
  await session.destroy({ removeDraft: true });

  // הרשומה עצמה אינה נכתבת כאן: `destroy({ removeDraft: true })` קורא
  // ל-`keeper.discardDraft`, וזה כותב דרך `persistCombinedSession` — שמרכיבה
  // את האוסף מהטאבים ש**נשארו** במפה (הטאב הזה כבר הוסר ממנה). כלומר טאב
  // שנסגר יורד מהרשומה מיד, ואינו חוזר בעלייה הבאה.
  if (!wasActive) return;

  await activateAfterClose();
}

/** כפתור „+”: תמיד טאב ריק חדש — לא בודק „שינויים לא שמורים”, כי אין מה לאבד. */
async function onDocumentTabNew(): Promise<void> {
  if (isOpenBusy()) return;
  activateTab(createNewDocumentSession());
  await openDocument();
}

/**
 * ## קיצורי הטאבים
 *
 * חמש הפעולות שמתחת הן מה שהרג'יסטרי קורא לו `tab-*`, והן **אינן** מסלול
 * שני: כל אחת מהן נכנסת לאותה פונקציה שהעכבר נכנס אליה — `onDocumentTabNew`,
 * `onDocumentTabClose`, `onDocumentTabSelect` — ולכן השאלה על „המסמך לא
 * נשמר”, החסימה בזמן פתיחה (`isOpenBusy`) והטעינה העצלה של טאב ששוחזר
 * מתרחשות בקיצור בדיוק כמו בלחיצה. זה אותו כלל שכבר קיים לכל הרצועה: קיצור
 * הוא הדרך השנייה לאותו כפתור, לא דרך עוקפת.
 *
 * הסדר שהמעבר נשען עליו הוא **סדר הרצועה** — `sessions` הוא `Map`, והאיטרציה
 * שלו היא סדר ההכנסה, כלומר בדיוק מה ש-`documentTabs` מרנדר. לא סדר השימוש
 * (`recentTabs`): הרצועה גלויה, והצירוף שמזיז בה סמן חייב להזיז אותו למקום
 * שהעין רואה. זו גם ההתנהגות של `Ctrl+Tab` בדפדפנים, ולא זו של `Ctrl+Tab`
 * ב-VSCode.
 *
 * ב-RTL „הבא” הוא שמאלה, מפני ש„הבא” הוא המקום הבא ברצועה — ורצועה עברית
 * מתחילה מימין. אותו כלל בדיוק כמו בחצי הלשוניות של הרצועה (`nextTabIndex`
 * ב-ui/ribbon/aria.ts), ומאותו טעם.
 */

/** מזהי הטאבים בסדר הרצועה. ראו „קיצורי הטאבים”. */
function tabOrder(): DocumentSessionId[] {
  return Array.from(sessions.keys());
}

/** `Ctrl+Tab` / `Ctrl+Page Down` / `Ctrl+F6` — והכיוון ההפוך עם Shift. */
function stepTab(direction: 'next' | 'prev'): void {
  const ids = tabOrder();
  // טאב יחיד: אין לאן. הצירוף עדיין נבלע — ראו „הטאבים” ב-ui/shortcuts/actions.ts.
  if (ids.length < 2) return;

  const at = activeSession.value ? ids.indexOf(activeSession.value.id) : -1;
  const delta = direction === 'next' ? 1 : -1;
  // גלישה מהסוף להתחלה, כמו בדפדפן. `at === -1` (מצב שאין לו מסלול ידוע)
  // נופל על הטאב הראשון ולא על חישוב שלילי.
  const target = ids[(((at + delta) % ids.length) + ids.length) % ids.length];
  if (target !== undefined) onDocumentTabSelect(target);
}

/**
 * `Alt+1`…`Alt+8` ו-`Alt+9`. מיקום שאין בו טאב אינו עושה דבר — בדיוק כמו
 * `Ctrl+7` בדפדפן עם שלוש לשוניות פתוחות.
 */
function goToTab(position: number | 'last'): void {
  const ids = tabOrder();
  const target = position === 'last' ? ids[ids.length - 1] : ids[position - 1];
  if (target !== undefined) onDocumentTabSelect(target);
}

/** `Ctrl+W` / `Ctrl+F4` — סוגר את הטאב שעל המסך. */
function closeActiveTab(): void {
  const session = activeSession.value;
  if (!session) return;
  void onDocumentTabClose(session.id).catch((error: unknown) => {
    console.warn('[otzaria-word] סגירת הטאב נכשלה', error);
  });
}

/**
 * `Ctrl+Shift+T` — הטאב האחרון שנסגר, בחזרה.
 *
 * מה שנפתח הוא **הקובץ**, לא מצב העריכה שהיה בטאב: הטאב פורק, המנוע שלו
 * שוחרר, והטיוטה נמחקה במסלול הסגירה (`destroy({ removeDraft: true })`).
 * זו גם ההתנהגות של דפדפן — לשונית שנפתחת מחדש טוענת את הדף, לא את מה
 * שהוקלד בטופס — וכל הבטחה אחרת כאן הייתה שקר, כי אין ממה לקיים אותה. מה
 * שכן שורד את „לסגור בלי לשמור” הוא הגיבוי של sessions/discard-backup.ts,
 * והוא מסלול אחר עם מסך משלו.
 *
 * הרשומה **נשלפת** מהמחסנית לפני הניסיון, גם כשהוא נכשל: token שאינו נפתר
 * לא ייפתר בלחיצה השנייה, ומחסנית שמחזירה את אותה שגיאה שוב ושוב היא מחסנית
 * תקועה. אותה הכרעה בדיוק כמו ב-`onOpenDialogRecent`, ששם היא מסירה את
 * השורה מרשימת האחרונים.
 */
async function reopenClosedTab(): Promise<void> {
  if (isOpenBusy()) return;
  const closed = closedTabs.shift();
  if (!closed) {
    setStatus('אין טאב שנסגר בהפעלה הזאת');
    return;
  }

  // `resolveRememberedFile` ולא `resolveFileUrl` הגולמית: היא מחזירה את
  // ההרשאה שהטאב ידע עליה, ובלעדיה הטאב שנפתח מחדש הוא קריאה-בלבד לתמיד.
  // ההנמקה המלאה ליד `closedTabs`.
  const file = await resolveRememberedFile(closed);
  if (!file) {
    setStatus(`${closed.name} לא נמצא — ייתכן שהקובץ הוזז או נמחק`, true);
    return;
  }

  ensureOpenTargetTab();
  await openDocument(file);
}

/**
 * הטאב שעובר אליו מי שסגר את הפעיל. `next()` על `Map.values()` הוא הטאב
 * הראשון שנשאר, לפי סדר הפתיחה — כמו ב-VSCode/דפדפנים.
 */
async function activateAfterClose(): Promise<void> {
  const next = sessions.values().next().value ?? createNewDocumentSession();
  activateTab(next);
  // שלושה מצבים, ורק אחד מהם כבר מוכן: טאב שהיה פתוח (יש בו מסמך), טאב
  // ששוחזר וטרם נטען (נטען עכשיו, בדיוק כמו מעבר רגיל אליו), וטאב חדש-לגמרי
  // שנוצר כאן מפני שהאחרון נסגר — צריך לפתוח בו מסמך, בדיוק כמו „+”.
  if (next.pendingRestore) await openPendingTab(next);
  else if (next.swap.current === null) await openDocument();
}

/**
 * סוגרת את כל הטאבים ומשאירה במקומם מסמך חדש-ריק. „יציאה” בלבד קוראת לה,
 * ורק אחרי שכל טאב נשאל (`resolveUnsavedBeforeClose`).
 *
 * הפירוק זהה לזה של סגירת טאב בודד, לכל טאב בתורו: הסרה מהמפה ומ-`recentTabs`,
 * ואז `destroy({ removeDraft: true })` — שמוחק גם את הטיוטה, כי לא נשאר טאב
 * שיפתח אותה (ראו `DocumentSession.destroy`).
 *
 * מה שנכתב לרשומה בסוף הוא האוסף שנשאר, כלומר הטאב החדש בלבד — וזו המשמעות
 * המעשית של „המסמך נסגר”: גם מי שיפתח את התוסף מחדש, ולא רק מי שחוזר אליו
 * מהספרייה, יקבל עורך נקי.
 *
 * הטאב החדש נוצר דרך `activateAfterClose`, בדיוק כמו בסגירת הטאב האחרון: אין
 * מצב „תוסף בלי טאבים” — רצועת הטאבים, הרצועה ושורת המצב כולן מניחות מסמך.
 */
async function closeAllSessions(discarded: ReadonlySet<DocumentSessionId> = new Set()): Promise<void> {
  for (const session of Array.from(sessions.values())) {
    sessions.delete(session.id);
    const recentAt = recentTabs.indexOf(session.id);
    if (recentAt >= 0) recentTabs.splice(recentAt, 1);
    // בדיוק הטאבים שנאמר עליהם „לא לשמור”, ולפני הפירוק שמוחק את הטיוטה.
    if (discarded.has(session.id)) await backupDiscardedDocument(session);
    await session.destroy({ removeDraft: true });
  }
  await activateAfterClose();
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

/**
 * מצב מיקוד — והמסך המלא שנלווה אליו.
 *
 * הפסים שלנו נעלמים ב-CSS (ראו „מצב מיקוד” ב-`<style>`), אבל סביבנו יושבת
 * אוצריא: פס הכותרת שלה, שורת הטאבים וסרגל הניווט. „מיקוד” שמשאיר אותם על
 * המסך מסתיר שליש ממה שמפריע. הבקשה למסך מלא היא הדרך היחידה שיש לתוסף
 * לבקש את החלון כולו — ההנמקה המלאה ב-composables/window-fullscreen.ts.
 *
 * הבקשה נשלחת **מכאן ולא מ-`watch`** על הדגל: הדפדפן מקבל מסך מלא רק מתוך
 * מחווה של המשתמש, וכאן זו תמיד לחיצה או `F11`. שחזור מצב מיקוד מהפעלה
 * קודמת (`applyShellPreferences`) אינו מחווה, והוא בכוונה אינו עובר כאן.
 *
 * כישלון אינו מבטל את מצב המיקוד: מאחז שאינו מרשה מסך מלא עדיין מקבל מעטפת
 * נקייה, וזה הרוב של מה שהמצב הזה נותן.
 */
function toggleFocusMode(): void {
  isFocusMode.value = !isFocusMode.value;
  updateShellView({ focusMode: isFocusMode.value });
  // יציאה ממצב מיקוד מאפסת את החשיפה: אחרת המחלקה `reveal-*` נשארת, והפסים
  // חוזרים למצב הרגיל כשהם עדיין מסומנים כחשופים.
  // הכניסה מתחילה **פתוחה**: מצב מיקוד שמתחיל בהעלמת כל הפקדים בבת אחת נראה
  // כמו תקלה, ולא כמו מצב שנבחר. הלוח נשאר עד תנועת העכבר הראשונה אל גוף
  // המסמך, ומשם והלאה החשיפה היא בקצה בלבד — כרגיל.
  revealed.value = isFocusMode.value ? 'top' : null;
  // הפוקוס היה יכול להישאר על הכפתור שהרגע נלחץ — והוא בתוך רצועה שברגע הזה
  // יצאה מהמסך. `visibility: hidden` מוציא אותו משם ממילא, אבל אל ה-`<body>`
  // ולא אל הטקסט: המשתמש היה מקבל מסך מיקוד שאי אפשר להקליד בו עד שילחץ.
  if (isFocusMode.value) focusRing.toDocument();
  void (isFocusMode.value ? enterFullscreen() : exitFullscreen());
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
 * אזכור „@” (engine/at-mention-overlay.ts), על אותו container ובאותו תנאי
 * `documentGeneration` כמו ההשלמה מהספר.
 *
 * בלי טוגל, בשונה ממנה: „@” כמעט אינו מופיע בטקסט עברי רץ, ולכן אין כאן רעש
 * להשתיק — הרשימה נפתחת רק על אזכור, וגם אז רק אחרי שני תווים.
 */
let atMention: ReturnType<typeof installAtMention> | null = null;
watch([activeEditorContainer, activeSuperdoc, documentGeneration], () => {
  atMention?.dispose();
  atMention = null;
  if (!activeEditorContainer.value || !activeSuperdoc.value) return;
  atMention = installAtMention(activeEditorContainer.value, activeSuperdoc.value, {
    onStatus: (message, isError) => setStatus(message, isError),
  });
});

/**
 * עד איפה מגיעים הפסים בפועל — הגובל שמחזיק את החשיפה פתוחה.
 *
 * נמדד ולא קבוע: הגובה תלוי במה שמוצג — רצועה מכונסת, סרגל מידות כבוי, שורת
 * מצב בגופן אחר.
 *
 * המדידה נקראת **רק כשמשהו חשוף** (ראו `onPointerMove`), וזה מה שהופך אותה
 * לנכונה גם אחרי שהפסים יצאו מהזרימה: לוח מוסתר מוזז ב-`translateY(-100%)`
 * ומלבנו יושב מעל הקצה העליון, כלומר `bottom` שלילי — והנפילה ל-0 מחזירה את
 * ההחלטה לרצועת הקצה בלבד, שזה בדיוק הסף הנכון כשאין מה להחזיק פתוח.
 */
function measureRevealBounds(): RevealBounds | null {
  const shell = shellRef.value;
  if (!shell) return null;

  const topRect = shell.querySelector('.shell-top')?.getBoundingClientRect();
  const top = topRect ? Math.max(0, topRect.bottom) : 0;

  const statusRect = shell.querySelector('.word-statusbar')?.getBoundingClientRect();
  const bottom = statusRect && statusRect.height > 0 ? statusRect.top : window.innerHeight;

  return { top, bottom };
}

/**
 * במצב מיקוד הפסים מוסתרים, ומתגלים כשהמצביע מתקרב לקצה. הקצה ולא כל המעטפת:
 * `:hover` על השורש החזיר את כולם בכל תנועה בחלון, כלומר המצב לא הסתיר כלום.
 * ההחלטה עצמה ב-composables/focus-mode.ts, כדי שתהיה נבדקת.
 *
 * המדידה נעשית רק כשמשהו כבר חשוף — כלומר רק כשהמצביע נמצא באזור הפסים.
 * בזמן הקלדה, כשהמצביע בגוף המסמך, אין כאן שום קריאת גאומטריה.
 */
function onPointerMove(event: PointerEvent): void {
  if (!isFocusMode.value) return;
  const current = revealed.value;
  revealed.value = revealZone(event.clientY, window.innerHeight, {
    current,
    bounds: current ? measureRevealBounds() : null,
  });
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
function openFindDialog(mode: 'find' | 'replace', initialQuery = ''): void {
  findMode.value = mode;
  findInitialQuery.value = initialQuery;
  isFindOpen.value = true;
  void reportSearch(searchAdapter?.open());
}

async function onRunCommandFromTellMe(id: string, payload?: unknown): Promise<void> {
  if (!commandAdapter.value) {
    setStatus('אין מסמך פתוח לביצוע הפעולה', true);
    return;
  }
  const outcome = await commandAdapter.value.run(id, payload);
  reportCommand(outcome, id);
}

function onRunActionFromTellMe(action: ShellAction): void {
  runShellAction(action);
}

/** פעולות Tell Me שאינן פקודת מנוע או פעולת מעטפת רגילה. */
function onCustomActionFromTellMe(action: TellMeCustomAction): void {
  switch (action) {
    case 'export-pdf':
      void onExportPdf();
      break;
    case 'export-otzaria':
      void onExportOtzaria();
      break;
    case 'about':
      isAboutOpen.value = true;
      break;
    case 'clipboard-copy':
      void copySelection(activeSuperdoc.value).then((outcome) => reportCommand(outcome, 'clipboard-copy'));
      break;
    case 'clipboard-cut':
      void cutSelection(activeSuperdoc.value).then((outcome) => reportCommand(outcome, 'clipboard-cut'));
      break;
    case 'clipboard-paste':
      void pasteFromClipboard(activeSuperdoc.value).then((outcome) => reportCommand(outcome, 'clipboard-paste'));
      break;
    case 'ribbon-shulchan':
      ribbonTab.value = 'shulchan';
      ribbonCollapsed.value = false;
      break;
  }
}

function closeFindDialog(): void {
  isFindOpen.value = false;
  // השאילתה שהגיעה מ-Tell Me שייכת לפתיחה הנוכחית בלבד.
  findInitialQuery.value = '';
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
  isDocumentSurface,
  isModalOpen: isModalDialogOpen,
  runAction: (action) => runShellAction(action),
  report: reportCommand,
  misspelledWordAt: (at) => spellingOverlayRef.value?.wordAt(at.x, at.y) ?? null,
  addToDictionary: addWordToDictionary,
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

/**
 * מחזירה את הפוקוס למסמך אם הוא התייתם — כלומר אם השכבה שנסגרה לקחה איתה את
 * האלמנט שהחזיק אותו, והדפדפן השאיר אותו על ה-`<body>`.
 *
 * הבדיקה על `body` היא מה שהופך את זה לבטוח: מי שסוגר שכבה ומחזיר את הפוקוס
 * בעצמו אינו נגרר לכאן.
 */
function returnOrphanedFocus(): void {
  // אחרי שה-DOM התעדכן, ולא באותה שורה: ברגע שהדגל כובה השדה עדיין בעץ ועדיין
  // מחזיק את הפוקוס, ו-`activeElement` היה מראה אותו במקום את ה-`<body>` —
  // כלומר הבדיקה הייתה תמיד יוצאת מוקדם.
  void nextTick(() => {
    const active = document.activeElement;
    if (active !== null && active !== document.body) return;
    const opened = activeSuperdoc.value;
    if (opened) focusOpenedDocument(opened);
  });
}

/**
 * מה ש-`Escape` סוגר, לפי שכבות. מחזירה `true` אם משהו נסגר בפועל.
 *
 * **מוצאת לפונקציה בשם ולא נשארת בתוך `closeTopmost` מפני שיש לה קורא שני:**
 * יציאה ממסך מלא שלא באה מאיתנו (`watchFullscreen` ב-`onMounted`). בתוך מסך
 * מלא ה-`Escape` שסוגר אותו **נבלע** ואינו מגיע לדף כלל — נמדד ב-Chrome
 * אמיתי `keys: []` מול `fullscreenchange: [false]` — ולכן אין למאזין ההוא שום
 * דרך אחרת לדעת שהמשתמש הקיש `Escape`. שני הקוראים חייבים לעבור באותן שכבות
 * בדיוק, אחרת אותו מקש עושה שני דברים שונים לפי מה שהמשתמש אינו יכול לראות.
 *
 * „אודות” הוא `aria-modal`, ולכן הוא זה שנסגר כשהוא פתוח. החיפוש אינו מודאלי
 * ואפשר להמשיך לערוך מתחתיו, ולכן הוא נסגר רק כשאין חלון מעליו.
 */
function closeTopmostLayer(): boolean {
  /**
   * „המסמך לא נשמר” ראשון, לפני כולם: הוא נפתח מעל כל שכבה אחרת (כולל „פתח
   * מסמך”, שממנו מגיעים אליו), ומאחוריו יושבת החלטה שממתינה לתשובה.
   *
   * הענף הזה אינו כפילות של ה-`Escape` שבקומפוננטה עצמה. שם הוא תלוי בכך
   * שהמיקוד בתוך החלון, וכאן זה נשבר: הכפתור שהיה ממוקד עשוי לרדת מה-DOM
   * (ראו מסך השחזור למטה), ואז ההקשה מגיעה ל-`window` — ומשם, בלי הענף,
   * היא הייתה נופלת עד `focusRing.toDocument()`, כלומר מיקוד לתוך המסמך
   * שמאחורי מודאל פתוח, בעוד ההבטחה תלויה לנצח.
   *
   * „ביטול” ולא „לא לשמור”: `Escape` הוא נסיגה, לא אישור למחיקה.
   */
  if (unsavedPrompt.question.value) {
    unsavedPrompt.answer('cancel');
    return true;
  }
  /**
   * מסך השחזור, מעל „פתח מסמך” שממנו נכנסים אליו.
   *
   * הוא **חייב** להיות כאן ולא להסתמך על ה-`Escape` שבקומפוננטה: „הסר” מוריד
   * את ה-`<li>` שהמיקוד עליו מה-DOM, המיקוד נופל ל-`body`, ומאותו רגע ההקשה
   * אינה מגיעה לחלון כלל.
   */
  if (isDiscardedOpen.value) {
    isDiscardedOpen.value = false;
    return true;
  }
  // „פתח מסמך”: הוא נפתח **מעל** כל השאר (מהרצועה או מקיצור), והוא
  // `aria-modal` — כלומר כל עוד הוא פתוח, שום שכבה מתחתיו אינה זמינה.
  if (isOpenDialogOpen.value) {
    isOpenDialogOpen.value = false;
    return true;
  }
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
    // השדה שהיה בדיאלוג הוסר מה-DOM, והדפדפן מבריח את הפוקוס ל-`<body>` —
    // נמדד `active: BODY`. במצב מיקוד זו הקלדה שנעלמת בלי שום סימן, כמו
    // במעבר טאב. רק כשהפוקוס באמת התייתם: שכבה שמחזירה אותו בעצמה (דיאלוג
    // הקיצורים, למשל) לא תיגרר מכאן.
    returnOrphanedFocus();
    return true;
  }
  // מצב מיקוד הוא „חלון” גם הוא: הוא מסתיר את הרצועה ואת שורת המצב, ו-
  // `Escape` הוא המקש הראשון שכל משתמש מנסה כדי לצאת ממנו. בלי הענף הזה
  // היציאה היחידה הייתה למצוא שוב את F11 או לרחף מעל קצה המסך.
  //
  // **אחרי** כל השכבות שמעליו, וזה עיקר העניין: שכבה פתוחה נסגרת קודם, ומצב
  // המיקוד נשאר. זה נכון גם כשהקורא הוא מאזין המסך המלא — מצב מיקוד בלי מסך
  // מלא הוא מצב נתמך ממילא (המאחז עשוי לסרב לבקשה, ראו `enterFullscreen`).
  if (isFocusMode.value) {
    toggleFocusMode();
    return true;
  }
  // אין מה לסגור: `Escape` מאחד מפסי המעטפת מחזיר את הפוקוס למסמך.
  // כשהוא כבר שם — `false`, והאירוע ממשיך למנוע ולדפדפן.
  return focusRing.toDocument();
}

const runShellAction = createShellActionRunner({
  isSaving: () => saveSnapshot.value.isSaving,
  save: (saveAs) => void onSave(saveAs),
  print: () => void onPrint(),
  openFind: (mode) => openFindDialog(mode),
  openLink: () => void linkDialog.open(),
  // Ctrl+N ו-Ctrl+O פותחים את אותו דיאלוג בדיוק כמו הכפתורים ברצועה: „מסמך
  // חדש” ו„פתח קובץ” הם שני חצאי אותו מסך, ולא שתי פעולות שונות. הפעולות
  // הישירות (`onNewDocument`/`onPickAndOpen`) נשארו בשימוש מתוך הדיאלוג עצמו.
  newDocument: () => openOpenDialog(),
  openDocument: () => openOpenDialog(),
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
  openTellMe: () => titleBarRef.value?.focusTellMe(),
  closeTopmost: closeTopmostLayer,
  // חמש פעולות הטאבים — ראו „קיצורי הטאבים” ליד `stepTab`. שלוש מהן הן
  // בדיוק המטפלים של רצועת הטאבים, ולכן הקיצור והעכבר אינם יכולים להתפצל.
  newTab: () => void onDocumentTabNew(),
  closeTab: closeActiveTab,
  reopenClosedTab: () =>
    void reopenClosedTab().catch((error: unknown) => {
      console.warn('[otzaria-word] פתיחת הטאב שנסגר נכשלה', error);
    }),
  stepTab,
  goToTab,
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
 * פסי המעטפת מסומנים כלא-זמינים במצב מיקוד **כל זמן שהם מחוץ למסך**. הם
 * עדיין בעץ — הם רק יצאו מהזרימה והוזזו החוצה — ולכן בלי הסימון `F6` היה
 * מעביר את המשתמש לפס שאינו נראה. זה גם מה שנותן לשדה שם המסמך דרך יציאה:
 * `Escape` ממנו מזהה שהפוקוס בסרגל הכותרת ומחזיר אותו למסמך.
 */
function shellRegion(selector: string): HTMLElement | null {
  return shellRef.value?.querySelector<HTMLElement>(selector) ?? null;
}

/**
 * האם פס המעטפת של האזור הזה זמין ל-`F6` עכשיו.
 *
 * מחוץ למצב מיקוד — תמיד. בתוכו — רק אם הוא **נחשף בפועל**, וזה מה שהתיקון
 * הזה הוסיף: הדגל הקודם היה `!isFocusMode` לבדו, ולכן רצועה פרושה על המסך
 * במלוא רוחבה (הכניסה למצב מיקוד מתחילה ב-`reveal-top`, בלי שום ריחוף) לא
 * הייתה נגישה למקלדת כלל — נמדד שהפוקוס נשאר במשטח ההקלדה של המנוע. זה גם
 * מה ש-`FocusRegion.isAvailable` מבטיח בתיעוד שלו.
 *
 * לפי אזור ולא דגל אחד: החשיפה היא של קבוצה אחת בכל רגע (`reveal-top` או
 * `reveal-bottom`), וסימון שניהם כזמינים היה מחזיר בדיוק את הבאג — הפעם
 * לשורת המצב שמחוץ למסך.
 */
function revealedBarAvailable(zone: Exclude<RevealZone, null>): () => boolean {
  return () => !isFocusMode.value || revealed.value === zone;
}

const focusRing = createFocusRing({
  regions: [
    {
      id: 'titlebar',
      element: () => shellRegion('.word-titlebar'),
      isAvailable: revealedBarAvailable('top'),
    },
    {
      id: 'ribbon',
      element: () => shellRegion('.word-ribbon-container'),
      isAvailable: revealedBarAvailable('top'),
    },
    {
      id: 'document',
      element: () => editorStackRef.value,
      focus: () => focusDocument(activeSuperdoc.value),
    },
    {
      /*
       * כפתור היציאה ממצב מיקוד. אזור, למרות שהוא פקד בודד: הוא הפקד היחיד
       * שנראה במצב מיקוד, והמנוע בולע `Tab` — כלומר בלי הכניסה הזאת אין אליו
       * שום דרך מהמקלדת (נמדד בדפדפן). מחוץ למצב מיקוד הוא אינו בעץ בכלל,
       * ואז `element()` הוא `null` והמעגל מדלג עליו מאליו.
       *
       * `focus` משלו ולא נפילה ל-`querySelector`: החיפוש הגנרי סורק **צאצאים**,
       * והאזור כאן הוא הכפתור עצמו.
       */
      id: 'focus-exit',
      element: () => shellRegion('.focus-exit'),
      focus: () => {
        const button = shellRegion('.focus-exit');
        if (!button) return false;
        button.focus();
        return true;
      },
    },
    {
      id: 'statusbar',
      element: () => shellRegion('.word-statusbar'),
      isAvailable: revealedBarAvailable('bottom'),
    },
  ],
});

let shortcuts: ShortcutDispatcher | null = null;
let directionShortcut: { dispose: () => void } | null = null;
let undoRedoWatcher: UndoRedoWatcher | null = null;

/** לחיצה אחת על „שלח למסמך” — מה-latch או מהמאזין החי, אותו מסלול בדיוק. */
function runSendToDocument(event: SendToDocumentEvent): void {
  void handleSendToDocument(event, {
    host: activeSuperdoc.value,
    resolveHost: () => activeSuperdoc.value,
  }).then((outcome) => {
    if (outcome.ok) {
      setStatus(outcome.value === 'at-cursor' ? 'הקטע נוסף במקום הסמן' : 'הקטע נוסף בסוף המסמך');
    } else if (outcome.message) {
      setStatus(outcome.message);
    }
  });
}

/**
 * „שלח למסמך” בתפריט ההקשר של הקורא.
 *
 * נקראת מוקדם בעלייה, לפני פתיחת המסמך: `openPlugin: true` מוסר את אירוע
 * הלחיצה מיד אחרי ה-boot, והמאזין צריך להיות שם קודם. מה שקרה לפני שהוא
 * הגיע נאסף ב-latch שב-`index.html` — ולכן הסדר כאן: קודם המאזין החי, ורק
 * אחריו `takePendingContextMenuClicks`, שמעביר את ה-latch למצב חי. שניהם
 * באותה משימה סינכרונית, ואין חלון שבו אירוע נופל בין השניים.
 *
 * ההמתנה למסמך היא בצד השני, ב-`handleSendToDocument`: כאן עוד אין מסמך.
 *
 * עטופה ב-try: `on` פונה ל-SDK דרך `bridge()`, שזורק כשאין `window.Otzaria`.
 * פריט בתפריט ההקשר אינו תנאי לעליית העורך, ולכן כשל כאן נרשם וזהו —
 * „ציטוט מהקורא” ברצועה עושה את אותו דבר מתוך העורך.
 */
function registerSendToDocument(): void {
  if (!isAvailable()) return;

  try {
    // הפריט עצמו מוצהר ב-`contributes.startup` שבמניפסט, ואוצריא רושמת אותו
    // בעלייה בלי JS; הרישום כאן הוא הגיבוי למי שביטל את ההרשאה להצהרה.
    void registerSendToDocumentItem().then((outcome) => {
      if (!outcome.ok) console.warn('[otzaria-word]', outcome.message);
    });

    contextMenuListener = on('contextMenu.itemClicked', (event) => runSendToDocument(event));
    for (const pending of takePendingContextMenuClicks()) runSendToDocument(pending);
  } catch (error) {
    console.warn('[otzaria-word] רישום „שלח למסמך” נכשל', error);
  }
}

onMounted(async () => {
  /**
   * המנייה של גופני המכונה — ראשונה, ובלי `await`.
   *
   * בלי `await` מפני שהיא אינה תנאי לשום דבר: עד שהיא נוחתת הבורר מציג את
   * הרשימה הקבועה, וברגע שהיא נוחתת `watchEffect` מרכיב מחדש והבורר מתמלא.
   * פתיחת המסמך הראשון אינה אמורה להמתין למנייה של מאות משפחות.
   */
  void loadInstalledFonts().then((snapshot) => {
    installedFonts.value = snapshot;
  });

  /*
   * יציאה ממסך מלא שלא באה מאיתנו — `Escape` או `F11` של הדפדפן.
   *
   * **האירוע הזה הוא ה-`Escape` עצמו, ואין דרך אחרת לדעת שהוא הוקש.** נמדד
   * ב-Chrome אמיתי: `Escape` שמשמש ליציאה ממסך מלא **נבלע** — הדף קיבל
   * `keys: []` ורק `fullscreenchange: [false]`. כלומר בתוך מסך מלא אין שום
   * `keydown` שאפשר לתלות בו את הסגירה, ומאזין ששולח את המקש הזה ישר לכיבוי
   * מצב המיקוד הפך את השכבות: `Ctrl+F` ואחריו `Escape` אחד נמדדו כ-
   * `{focus:false, fs:false, find:true}` — הדיאלוג נשאר פתוח ומצב המיקוד כבה,
   * בדיוק ההיפך ממה שהמקש מבטיח.
   *
   * לכן דרך `closeTopmostLayer` ולא ישר ל-`toggleFocusMode`: אותן שכבות
   * בדיוק שהמקלדת עוברת בהן. שכבה פתוחה נסגרת והמצב נשאר — מצב מיקוד בלי
   * מסך מלא נתמך ממילא, כי המאחז עשוי לסרב לבקשה מלכתחילה.
   *
   * `isFocusMode` בלבד ולא בשני הכיוונים: כניסה למסך מלא מכל סיבה אחרת אינה
   * אמורה להדליק מצב מיקוד. ויציאה שכן באה מאיתנו כבר כיבתה את הדגל לפני
   * הקריאה ל-`exitFullscreen`, ולכן היא נופלת כאן על התנאי ואינה סוגרת שכבה
   * בטעות.
   */
  fullscreenListener = watchFullscreen((fullscreen) => {
    if (fullscreen || !isFocusMode.value) return;
    closeTopmostLayer();
  });

  // „שלח למסמך” — לפני כל `await` שבהמשך. אוצריא מוסרת את אירוע הלחיצה מיד
  // אחרי ה-boot, וכל שלב שנכשל לפני הרישום היה מפיל אותו לבור: ה-latch היה
  // ממשיך לצבור ואיש לא היה מרוקן אותו. אינו תלוי בכלום כאן — קיומו של מסמך
  // נבדק בזמן הלחיצה, לא עכשיו.
  registerSendToDocument();

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

    // הבחירה נטענת לפני שנפתח מסמך: העריכה הראשונה עלולה להתחיל סבב autosave,
    // ואם ההעדפה עוד לא הגיעה הוא היה רץ לפי ברירת המחדל ולא לפי מה שהמשתמש
    // בחר בהפעלה הקודמת. ההעדפה של הסרגל — מאותו טעם: היא חלה על המסמך שנפתח
    // מיד אחרי כאן.
    //
    // כל הקריאות במקביל ולא בזו אחר זו: כל אחת היא סבב IPC מלא מול אוצריא,
    // הן קוראות מפתחות שונים ואינן תלויות זו בזו — והן עומדות בין המשתמש לבין
    // פתיחת המסמך הראשון.
    const [storedAutosave, storedRuler, stored, storedRecents, storedDiscarded, storedSpellcheck] =
      await Promise.all([
        loadAutosaveEnabled(),
        loadRulerVisible(),
        loadPreviousSession(),
        loadRecentDocuments(),
        loadDiscardBackups(),
        loadSpellcheckEnabled(),
      ]);
    autosaveEnabled.value = storedAutosave;
    rulerPreference = storedRuler;
    // ממוינת כבר כאן ולא רק בתצוגה: זו הרשימה שכל שאר הקוד רואה, ורשימה
    // שמגיעה מ-storage אין לה הבטחת סדר.
    recentDocuments.value = sortedRecents(normalizeRecents(storedRecents));
    // אותה הכרעה בדיוק כמו של „אחרונים”: מה שמגיע מ-storage אין לו הבטחת
    // סדר, והמיון הוא של הרשימה שכל שאר הקוד רואה — לא של התצוגה בלבד.
    discardedBackups.value = normalizeBackups(storedDiscarded);

    // בדיקת האיות — **לא** ב-await: משיכת המילון היא 1.3MB, והעלייה לא
    // תמתין לה. מי שהדליק בהפעלה הקודמת יקבל את הסימון כשהמילון יגיע.
    if (storedSpellcheck) {
      void loadTorahDictionary().then((dictionary) => {
        if (dictionary) spellcheckDictionary.value = dictionary;
      });
    }

    // הרשומה, לפני יצירת הטאבים: נתיב הטיוטה של כל זוכר תלוי במזהה הטאב שלו,
    // וזה חייב להיות המזהה **שברשומה** — לא מזהה חדש שאין לו קשר לטיוטה
    // הקודמת של אותו מסמך. `stored` עצמו נטען למעלה, במקביל לשאר ההעדפות.
    //
    // ניקוי הטיוטות ה„יתומות” שהיה כאן **נמחק בכוונה**: הוא מחק את הטיוטה של
    // כל טאב שלא שוחזר, ומהרגע שכל הטאבים חוזרים אין טאב כזה — הלולאה הזאת
    // הייתה מוחקת בדיוק את העבודה שלא נשמרה, של כל טאב שאינו הפעיל.
    const restored = restoreTabs(stored);
    applyShellPreferences(stored);

    // ההאזנה למעבר לרקע נרשמת כאן, אחרי שיש למי לדווח. שלושת המקורות
    // וההנמקה — ב-host/lifecycle.ts. `flush` על **כל** הטאבים הפתוחים —
    // לא רק הפעיל — כי לכולם יש עבודה שעלולה לא להיות בדיסק.
    hiddenListener = onPluginHidden(() => {
      // מיקום הגלילה נלכד כאן, בזמן שהמיכל עוד קיים ועוד יודע אותו. ראו
      // sessions/pane-scroll.ts — במסלולים מסוימים המעבר לרקע מאפס אותו.
      const active = activeSession.value;
      if (active) active.ui.scroll = readPaneScroll(documentScrollHost(active));
      void Promise.all(Array.from(sessions.values()).map((s) => s.keeper.flush()));
    });

    // והחזרה: תיקון בלבד, ורק למה שבאמת אבד — `repairPaneScroll` אינה נוגעת
    // במיכל ששרד. בפריים הבא, כי המנוע עשוי לעמד מחדש בחזרה למסך.
    shownListener = onPluginShown(() => {
      const active = activeSession.value;
      if (!active) return;
      const remembered = active.ui.scroll;
      const repair = (): void => {
        if (activeSession.value === active) repairPaneScroll(documentScrollHost(active), remembered);
      };
      repair();
      if (typeof requestAnimationFrame === 'function') requestAnimationFrame(repair);
    });

    // טעינת מסמך אחרון או פתיחת מסמך ריק. תחנת „פותח את המסמך” מדווחת
    // מ-createEditor ולא מכאן: כאן היא הייתה מוקדמת בשנייה ויותר — הפתיחה
    // ממתינה קודם לחבילת המנוע, ומסך טעינה שאומר „פותח” בזמן שהוא מוריד
    // 9MB מתאר את השלב הלא נכון.
    try {
      await reopenPreviousSession(restored.keeper.state);
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
  // מי שממתין לתשובה מקבל „ביטול” ואינו נשאר תלוי — ראו use-unsaved-prompt.ts.
  unsavedPrompt.dispose();
  bookCompletion?.dispose();
  bookCompletion = null;
  atMention?.dispose();
  atMention = null;
  zoomCenter?.dispose();
  zoomCenter = null;
  shortcuts?.dispose();
  directionShortcut?.dispose();
  undoRedoWatcher?.dispose();
  undoRedoWatcher = null;
  // חיפוש-בזמן-הקלדה שממתין ירוץ אחרי הפירוק על handle של controller מפורק.
  // בכל הטאבים, לא רק הפעיל — לכולם יש `searchAdapter`/`keeper` משלהם.
  for (const s of sessions.values()) {
    s.searchAdapter?.dispose();
  }
  hiddenListener?.();
  hiddenListener = null;
  shownListener?.();
  shownListener = null;
  contextMenuListener?.();
  contextMenuListener = null;
  fullscreenListener?.();
  fullscreenListener = null;
  paneScrollGuard?.();
  paneScrollGuard = null;
  // מעטפת שנפרקת בזמן מסך מלא הייתה משאירה את החלון מורחב בלי מי שיצא ממנו.
  if (isFullscreen()) void exitFullscreen();
  // הפריט עצמו אינו מוסר כאן: אוצריא מסירה את רישומי המופע בעצמה בפירוק,
  // וההסרה הידנית רק מסתכנת במחיקת העותק ההצהרתי — הפריט שאמור להישאר
  // בתפריט גם כשהתוסף סגור. ההנמקה המלאה ב-host/otzaria-reader.ts.
  for (const s of sessions.values()) {
    // המנוע וטיימרי השמירה של **כל** טאב, לא רק הפעיל: לפני ריבוי המסמכים
    // היה כאן אחד, ומאז כל טאב מחזיק מופע SuperDoc משלו עם ה-workers שלו.
    // `destroy()` של ה-session אינו מתאים כאן — הוא מוחק גם את הטיוטה, ופירוק
    // המעטפת אינו „המשתמש ביקש למחוק”.
    s.swap.destroy();
    s.save.dispose();
    s.keeper.dispose();
  }
  keeper = null;
  // ה-interval של הזחילה הוא הדבר היחיד כאן שממשיך לרוץ בלי בעלים.
  documentLoad.dispose();
});

/**
 * הנתיב השטוח שבו ישבה טיוטה יחידה לפני ריבוי המסמכים (`SESSION_VERSION` 1).
 * רשומת v1 נזרקת כולה ב-`normalizeSession` (גרסה לא תואמת), ואיתה נעלם כל
 * זכר לנתיב הזה — אבל הקובץ עצמו נשאר בדיסק. בלי ניקוי מפורש הוא יתום
 * לצמיתות: אין עוד קוד שכותב או קורא את השם הזה כדי לדרוס אותו.
 */
const LEGACY_DRAFT_PATH = 'session-draft.docx';

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

  // הרשומה שמצאנו אינה v2 (חסרה, פגומה, או v1 ישנה) — ניקוי חד-פעמי,
  // best-effort, של טיוטת ה-v1 השטוחה שאין עוד מי שיזכור אותה. ראו
  // `LEGACY_DRAFT_PATH`.
  void deleteWorkspaceEntry(LEGACY_DRAFT_PATH);

  const migrated = sessionFromLastDocument(await loadLastDocument());
  if (migrated) void forgetLastDocument();
  return migrated;
}

/**
 * בונה טאב לכל רשומת מסמך שהייתה פתוחה, ומחזירה את זה שיהיה פעיל.
 *
 * ## מה נפתח מיד ומה לא — „טעינה עצלה”
 *
 * עד כאן שוחזר טאב אחד בלבד — הפעיל — ושאר הטאבים שהמשתמש הותיר פתוחים פשוט
 * נעלמו, וטיוטותיהם נמחקו. עכשיו **כל** רשומה שבאוסף מקבלת טאב, בסדר שהיה,
 * עם השם שלה ועם הנקודה של „לא נשמר”; מה שנטען לתוך מנוע מיד הוא המסמך הפעיל
 * בלבד, וכל טאב אחר נטען ברגע שעוברים אליו (`openPendingTab`).
 *
 * ההחלטה הזאת אינה קיצור דרך אלא מה שנכון גם לכשעצמו. כל מסמך פתוח הוא מופע
 * SuperDoc מלא עם ה-workers שלו; פתיחת חמישה מסמכים בעלייה הייתה מכפילה פי
 * חמישה את הזמן עד שהמשתמש רואה את המסמך שהוא באמת עובד עליו, ואת הזיכרון —
 * בלי שביקש דבר. זה גם בדיוק מה ש-VSCode ודפדפנים עושים בשחזור לשוניות.
 * המסמך שממתין אינו „פחות פתוח”: הרשומה שלו, הסמן שלו והטיוטה שלו יושבים
 * בזוכר שלו מהרגע הראשון, נכתבים חזרה בסגירה, ואינם תלויים בכך שייטען.
 *
 * ## מה שאסור היה להישבר כאן
 *
 * המזהה של הטאב הוא המזהה שברשומה, ולא מזהה חדש: הוא מה שקושר את הטאב לקובץ
 * הטיוטה שלו (`draftPathFor`) ולרשומה שלו. רשומה שאין בה מסמכים כלל, או
 * שאין בה בכלל, נותנת טאב ריק אחד — בדיוק ההתנהגות של הפעלה ראשונה.
 */
function restoreTabs(stored: SessionState | null): DocumentSession {
  const view = stored?.view ?? defaultView();
  // רשומה ריקה (אין `documents`) אינה מצב חריג: זו כל הפעלה ראשונה, וגם
  // רשומה שנפסלה בקריאה. טאב אחד, חדש, בדיוק כמו קודם.
  const entries = stored && stored.documents.length > 0 ? stored.documents : [emptyDocumentEntry()];
  const activeId = entries.some((entry) => entry.id === stored?.activeId)
    ? stored!.activeId
    : entries[0]!.id;

  let active: DocumentSession | null = null;
  for (const entry of entries) {
    const tab = createNewDocumentSession({ id: entry.id, restore: { entry, view } });
    if (entry.id === activeId) active = tab;
  }

  // ה-`!` בטוח: `activeId` נבחר מתוך `entries`, וכל רשומה שבו יצרה טאב.
  const activeTab = active!;
  activateTab(activeTab);
  // הטאב הפעיל נפתח מיד (`reopenPreviousSession` ב-onMounted) ולכן אינו ממתין.
  activeTab.pendingRestore = false;

  // אין כאן ניקוי טיוטות יתומות, ובכוונה: כל רשומה שהגיעה לכאן קיבלה טאב.
  // רשומה שנשמטה בקריאה (מזהה כפול, ראו `readDocuments`) כבר אינה ב-`stored`
  // ואי אפשר לדעת עליה מכאן — וזו גם הסיבה שהיא נשמטת שם ולא כאן.
  return activeTab;
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
  // ואיתו הלוח העליון, בדיוק כמו בכניסה דרך `toggleFocusMode`: מצב מיקוד
  // שנפתח כשכל הפקדים מוסתרים נראה כמו תקלה ולא כמו מצב שנבחר. עלייה עם
  // ההעדפה דלוקה היא בדיוק הרגע הזה — ובלי השורה הזאת השחזור נתן את מה
  // שהכניסה הידנית הוגדרה כבר כתקלה.
  if (isFocusMode.value) revealed.value = 'top';
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
async function reopenPreviousSession(
  session: SessionState | null,
  fromThisRun = false,
): Promise<void> {
  const entry = activeEntry(session);
  const remembered = entry?.document ?? null;
  const file = remembered ? await resolveRememberedFile(remembered) : null;

  if (remembered && !file) {
    // הקובץ אינו נגיש — הוזז, נמחק, או שההרשאה בוטלה. אבל עבודה שלא נשמרה
    // אינה תלויה בו: אין לה יעד כתיבה בכל מקרה („שמור” יפתח „שמור בשם”),
    // ולכן פתיחתה כמסמך חדש אינה יכולה לדרוס דבר — והיא הדרך היחידה שלא
    // לאבד אותה. הטיוטה עוברת לבעלות המסמך שנפתח ממנה (ראו `setDocument`).
    const orphan = entry?.draft?.documentToken === remembered.token
      ? await readDraftBytes(entry.draft.path)
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
    if (draft) setStatus(restoredDraftMessage(session, fromThisRun));
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
function restoredDraftMessage(session: SessionState | null, fromThisRun = false): string {
  // „מההפעלה הקודמת” אינו נכון לטאב שנרדם ונפתח מחדש באותה הפעלה (`sleepTab`)
  // — שם השינויים הם של המשתמש מלפני דקה. משפט שאינו נכון בשורת המצב שוחק
  // את האמון בכל השאר שנאמר בה.
  const base = fromThisRun
    ? 'שוחזרו השינויים שלא נשמרו'
    : 'שוחזרו שינויים שלא נשמרו מההפעלה הקודמת';
  const draft = activeEntry(session)?.draft;
  const age = draft ? draftAgeLabel(draft.savedAt, Date.now()) : null;
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
  const entry = activeEntry(session);
  const decision = decideDraftRecovery({
    draft: entry?.draft ?? null,
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

  return readDraftBytes(entry?.draft?.path ?? draftPathFor(documentId));
}

/** בייטי הטיוטה כ-Blob שאפשר למסור למנוע, או `null` כשאין. */
async function readDraftBytes(path: string): Promise<Blob | null> {
  const bytes = await readWorkspaceBytes(path);
  return bytes ? new Blob([bytes], { type: DOCX_MIME }) : null;
}

/**
 * „לא לשמור” — עם עותק לפני שהוא נמחק.
 *
 * שני המסלולים שמוחקים טיוטה מיד („פתח מסמך אחר” ו„מסמך חדש”) עוברים כאן.
 * מסלול הסגירה אינו עובר כאן אלא קורא ל-`backupDiscardedDocument` בעצמו,
 * מפני ששם המחיקה היא של `destroy({ removeDraft: true })` ולא של
 * `discardDraft` — ראו `UnsavedResolution.discarded`.
 *
 * הסדר אינו סגנון: הגיבוי קורא את הבייטים, והמחיקה מוחקת את הקובץ שהם
 * נקראים ממנו במסלול הטאב שלא נטען.
 */
async function discardWithBackup(session: DocumentSession | null): Promise<void> {
  if (!session) return;
  await backupDiscardedDocument(session);
  await session.keeper.discardDraft();
}

/**
 * פעולות הגיבוי משתפות גם את רשומת המטא-נתונים וגם חמש משבצות קבועות.
 * ללא שרשור, שתי סגירות יכולות לבחור אותה משבצת מתוך אותה רשימה ישנה, ואז
 * כתיבה מאוחרת יותר תדרוס עותק ורשומה. אותו תור כולל גם „הסר”, כדי שלא
 * ידרוס כתיבה שזה עתה הושלמה את מחיקתו.
 */
let discardedStorageChain: Promise<void> = Promise.resolve();

function enqueueDiscardedStorageOperation(operation: () => Promise<void>): Promise<void> {
  const task = discardedStorageChain.then(operation);
  // פעולה שנכשלה מדווחת לקורא שלה, אך אינה מרעילה את התור עבור הפעולה הבאה.
  discardedStorageChain = task.catch(() => undefined);
  return task;
}

/**
 * כותבת את המסמך שנאמר עליו „לא לשמור” לאחת מחמש המשבצות של הגיבוי.
 *
 * ההחלטה מי נדרס — ולמה משבצות ולא מזהים — ב-sessions/discard-backup.ts.
 * כאן נשארת ההרכבה: לקרוא את הרשומה, לבחור משבצת, לכתוב את הבייטים, ורק אם
 * הכתיבה הצליחה לרשום אותה. הסדר הזה הוא מה שמונע רשומה שמצביעה לקובץ שלא
 * נכתב.
 *
 * **כשל נאמר ואינו בולע את הפעולה.** המשתמש ביקש לסגור בלי לשמור, וזה מה
 * שיקרה בכל מקרה; מה שהוא חייב לדעת הוא שההבטחה שהוצגה לו בדיאלוג
 * (`DISCARD_BACKUP_NOTE`) לא התקיימה הפעם. `setStatus(…, true)` מדווח גם
 * לשורת המצב וגם כשגיאה של אוצריא — ראו `setStatus`.
 */
async function backupDiscardedDocument(session: DocumentSession): Promise<void> {
  try {
    await enqueueDiscardedStorageOperation(() => backupDiscardedDocumentNow(session));
  } catch (error) {
    console.error('Could not save discarded document backup', error);
    setStatus('שמירת העותק לשחזור נכשלה', true);
  }
}

async function backupDiscardedDocumentNow(session: DocumentSession): Promise<void> {
  const bytes = await discardedBytes(session);
  if (!bytes) {
    setStatus('לא נשמר עותק לשחזור של השינויים שלא נשמרו', true);
    return;
  }

  const list = normalizeBackups(await loadDiscardBackups());
  const slot = nextBackupSlot(list);
  const written = await writeWorkspaceBytes(backupPathFor(slot), bytes);
  if (written !== 'written') {
    setStatus(
      written === 'too-large'
        ? 'המסמך גדול מכדי לשמור ממנו עותק לשחזור'
        : 'שמירת העותק לשחזור נכשלה',
      true,
    );
    return;
  }

  const next = rememberDiscard(list, {
    slot,
    name: sessionDisplayTitle(session),
    size: bytes.byteLength,
    discardedAt: Date.now(),
    // מי שישחזר צריך לדעת מאיזה קובץ זה בא. אחרי הסגירה הידיעה הזאת נעלמת
    // עם הטאב, ואי אפשר לאסוף אותה מאוחר יותר.
    token: activeEntry(session.keeper.state)?.document?.token ?? null,
  });
  await saveDiscardBackups(next);
  // הרשימה שבזיכרון היא מה שמזין את המונה על „נסגרו בלי לשמור” ואת מסך
  // השחזור. בלי השורה הזאת הכפתור היה מופיע רק בהפעלה הבאה — כלומר בדיוק לא
  // ברגע שבו המשתמש מחפש את מה שהרגע ויתר עליו.
  discardedBackups.value = sortedBackups(next);
}

async function forgetDiscarded(slot: number): Promise<void> {
  await enqueueDiscardedStorageOperation(async () => {
    const list = normalizeBackups(await loadDiscardBackups());
    const next = forgetDiscard(list, slot);
    await saveDiscardBackups(next);
    await deleteWorkspaceEntry(backupPathFor(slot));
    discardedBackups.value = sortedBackups(next);
    // מסך ריק אינו מסך: הכפתור שמוביל לכאן ממילא נעלם כשאין מה להציג.
    if (next.length === 0) isDiscardedOpen.value = false;
  });
}

/**
 * הבייטים שייכנסו לגיבוי, משני מקורות — ולא במקרה בסדר הזה.
 *
 * **המנוע קודם.** הטיוטה נכתבת בהשהיה (עד דקה, ראו `DRAFT_MAX_WAIT_MS`),
 * ולכן היא עשויה להיות ישנה בדקה שלמה מהמסמך שעל המסך. גיבוי שנועד להציל את
 * מה שהמשתמש הרגע ויתר עליו חייב להיות מה שהוא ראה.
 *
 * **הטיוטה כשאין מנוע.** טאב ששוחזר וטרם נטען, וטאב שהטעינה שלו נכשלה, הם
 * בדיוק המקרים שבהם אין ממה לייצא ובכל זאת יש עבודה — והיא נמצאת רק בקובץ
 * הטיוטה. זהו המסלול הרגיל שלהם, לא נפילה לאחור.
 */
async function discardedBytes(session: DocumentSession): Promise<Uint8Array | null> {
  const draftPath = activeEntry(session.keeper.state)?.draft?.path ?? null;

  /**
   * „יש מנוע” אינו „המנוע מחזיק את העבודה”, וזה ההבדל שבין גיבוי לבין מסמך
   * ריק בשם הנכון.
   *
   * המסלול ששבר את זה: פתיחה שנכשלה. `reopenPreviousSession` שנכשל פותח
   * `openDocument(undefined, { remember: false })` — מסמך **ריק** — ולטאב יש
   * מעכשיו `swap.current` אמיתי ש-`exportDocx` עליו **מצליח**. הרשומה שלו
   * עדיין מחזיקה את הטיוטה עם העבודה, `destroy({ removeDraft: true })` עומד
   * למחוק אותה, וגיבוי שמעדיף את המנוע היה כותב את הריקנות ומאבד את הכול —
   * בדיוק במסלול שהתכונה הזאת קיימת בשבילו.
   *
   * `isDirty` הוא ההבחנה, וזו אותה הבחנה בדיוק ש-`askFromRecord` עושה
   * ב-`resolveUnsavedBeforeClose` (שם היא קובעת **איזו שאלה** נשאלת, וכאן
   * **מה מגובה”): מנוע מלוכלך מחזיק עבודה שאינה בשום מקום אחר, ולכן הוא
   * המקור; מנוע נקי שיש לצדו טיוטה אינו יכול להיות המקור — טיוטה נמחקת
   * בכל שמירה מוצלחת (`noteSaved`), ולכן „נקי ויש טיוטה” פירושו תמיד
   * שהטיוטה היא העבודה.
   */
  if (!session.save.snapshot.isDirty && draftPath) {
    const fromDraft = await readWorkspaceBytes(draftPath);
    if (fromDraft) return fromDraft;
    // הטיוטה לא נקראה — הקובץ נעלם או שהגשר נפל. המנוע שלמטה עדיף על כלום.
  }

  const open = session.swap.current;
  if (open) {
    try {
      const exported = await exportDocx(open.superdoc);
      return new Uint8Array(await exported.arrayBuffer());
    } catch (error) {
      // אותה הכרעה כמו ב-`writeDraftNow`: ייצוא שנכשל אינו מפיל את הפעולה.
      // אם יש טיוטה, היא עדיין עדיפה על כלום.
      console.warn('[otzaria-word] ייצוא העותק לשחזור נכשל', error);
    }
  }

  return draftPath ? readWorkspaceBytes(draftPath) : null;
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
  /* מסגרת ההתייחסות של הפסים במצב מיקוד — שם הם יוצאים מהזרימה. */
  position: relative;
}

/* קבוצת הפסים העליונים: כותרת, טאבי מסמכים, רצועה ושורת הסרגל. מחוץ למצב
   מיקוד זו עמודה רגילה, והפריסה זהה למה שהיה כשארבעתם היו ילדים ישירים של
   המעטפת. הקבוצה קיימת בשביל מצב מיקוד — ראו למטה. */
.shell-top {
  display: flex;
  flex-direction: column;
  flex-shrink: 0;
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

/* פאנל הטאב — נוצר דינמית (`document.createElement`) ב-App.vue, ולכן מחוץ
   ל-scope של Vue; `:global()` כדי שהכלל עדיין יחול עליו. כל טאב מקבל אחד,
   ורק של הפעיל אין `display: none` (App.vue, `activateTab`). ה-host של
   `editor-swap.ts` נוצר בתוכו וממלא אותו באותו אופן בדיוק. */
:global(.document-pane) {
  position: absolute;
  inset: 0;
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

/**
 * ## מצב מיקוד
 *
 * ההסתרה הייתה `opacity: 0`, והיא לא הסתירה: הפסים שמרו על מקומם בפריסה,
 * כלומר אזור המסמך נשאר בדיוק באותו גובה ומעליו נמתחה רצועה לבנה (נמדד:
 * 194 פיקסלים בחלון של 429 — כמעט מחצית מהגובה). מצב מיקוד שמלבין את הפסים
 * במקום להעלים אותם אינו מצב מיקוד, אלא מסמך קטן עם שוליים ריקים.
 *
 * לכן הפסים **יוצאים מהזרימה**: `position: absolute` מוציא אותם מעמודת
 * ה-flex, ואזור המסמך גדל מיד לכל גובה החלון. החשיפה בקצה מחזירה אותם כלוח
 * צף מעל המסמך ולא כפס שדוחף אותו — ולכן ריחוף על הקצה אינו מזיז את הטקסט
 * ואינו מכריח את המנוע לפרוס את המסמך מחדש בכל תנועת עכבר.
 *
 * `visibility: hidden` ולא רק ההזזה: פס שהוזז אל מחוץ למסך עדיין נמצא בסדר
 * ה-Tab ובעץ הנגישות. ההשהיה עליו (`0s linear 0.24s`) היא מה שמשאיר אותו
 * נראה לכל אורך ההחלקה החוצה ומעלים אותו רק בסופה.
 *
 * `z-index` מפני ש-`.editor-area` בא אחרי הפסים ב-DOM וגם הוא ממוקם
 * (`position: relative`) — בלעדיו אזור המסמך היה נצבע **מעל** הלוח שנחשף.
 */
.word-app-shell.focus-mode .shell-top,
.word-app-shell.focus-mode :deep(.word-statusbar) {
  position: absolute;
  inset-inline: 0;
  z-index: 2;
  visibility: hidden;
  transition:
    transform 0.24s ease,
    visibility 0s linear 0.24s;
}

.word-app-shell.focus-mode .shell-top {
  top: 0;
  transform: translateY(-100%);
}

.word-app-shell.focus-mode :deep(.word-statusbar) {
  bottom: 0;
  transform: translateY(100%);
}

/* הסרגל האנכי הוא היחיד שאינו מצטרף לחשיפה: הוא עמודה לכל גובה המסמך, והחזרה
   שלו בכל ריחוף הייתה משנה את **רוחב** אזור העריכה — כלומר פורסת את המסמך
   מחדש בדיוק במצב שנועד להשקיט אותו. במצב מיקוד הוא פשוט אינו שם, ורוחבו
   נוסף למסמך. */
.word-app-shell.focus-mode :deep(.doc-vruler) {
  display: none;
}

/* החשיפה לפי קצה, ולא `:hover` על השורש: השורש הוא כל החלון, ולכן כל תנועת
   עכבר החזירה את הפסים — ומצב המיקוד לא הסתיר כלום. ההחלטה עצמה
   ב-composables/focus-mode.ts, כדי שתהיה נבדקת. */
.word-app-shell.focus-mode.reveal-top .shell-top,
.word-app-shell.focus-mode.reveal-bottom :deep(.word-statusbar) {
  transform: none;
  visibility: visible;
  transition:
    transform 0.24s ease,
    visibility 0s;
}

/* הצל הוא מה שאומר „זה צף מעל המסמך” ולא „זה חלק מהעמוד”. אותו ערך כמו כל
   משטח צף אחר בתוכנה (ContextMenu.vue, RibbonMenuButton.vue). */
.word-app-shell.focus-mode.reveal-top .shell-top {
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.16);
}

.word-app-shell.focus-mode.reveal-bottom :deep(.word-statusbar) {
  box-shadow: 0 -4px 16px rgba(0, 0, 0, 0.16);
}

/**
 * כפתור היציאה ממצב מיקוד.
 *
 * `Esc` ו-`F11` עובדים, אבל מצב מיקוד הוא בדיוק המצב שבו אין על המסך שום פקד
 * שרומז עליהם — ומשתמש שנכנס בטעות נשאר בפנים. הכפתור עמום ולא נעלם: תפקידו
 * להימצא כשמחפשים אותו, לא להתחרות בטקסט.
 *
 * המיקום אינו שרירותי: **בפינה התחתונה**, ו-32 פיקסלים מהתחתית. הקצה העליון
 * הוא המקום שהלוח נחשף בו — כפתור שם היה נעלם מתחתיו בדיוק כשנכנסים למצב
 * (הכניסה מתחילה פתוחה), ושורת המצב בגובה 24 פיקסלים היא הסיבה ל-32.
 *
 * **והוא יושב מעל העמוד, לא לצדו.** נמדד: `overlapsPage: true`, ו-
 * `elementFromPoint` מתחתיו מחזיר `DIV.superdoc-page` — גם בחלון של 800
 * פיקסלים וגם ב-400. אין שוליים אפורים לשבת עליהם: העמוד מרוכז ורחב כמעט
 * כמו החלון בכל רוחב סביר. זו הסיבה ל-`opacity: 0.5` ולגודל הקטן, וזה גם
 * הגבול שלהם — הכפתור אינו „מחוץ לתוכן”, הוא **עמום מעל התוכן**.
 *
 * `z-index: 1` — מעל שכבות המסמך (PageBorderOverlay ואחיותיה, שגם הן 1, והוא
 * אחריהן ב-DOM) ומתחת ללוח שנחשף בקצה (2).
 */
.focus-exit {
  position: absolute;
  bottom: 32px;
  inset-inline-end: 12px;
  z-index: 1;
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 5px 10px;
  border: 1px solid var(--color-outline-variant);
  border-radius: 999px;
  background: var(--color-surface-container-highest);
  color: var(--color-on-surface-variant);
  font-family: inherit;
  font-size: 12px;
  line-height: 1;
  cursor: pointer;
  opacity: 0.5;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.16);
  transition:
    opacity 0.15s ease,
    border-color 0.15s ease,
    color 0.15s ease;
}

/* `:focus-visible` אינו קוד מת למרות ה-`@pointerdown.prevent`: הכפתור הוא אזור
   במעגל ה-`F6` (ראו `focusRing`), וזו הדרך היחידה להגיע אליו מהמקלדת — המנוע
   בולע `Tab`. ה-`prevent` חוסם רק את המיקוד שבא מלחיצה. */
.focus-exit:hover,
.focus-exit:focus-visible {
  opacity: 1;
  border-color: var(--color-outline);
  color: var(--color-on-surface);
}

/* מי שביקש פחות תנועה מקבל את אותה חשיפה בלי ההחלקה — ולא חשיפה שנעלמת. */
@media (prefers-reduced-motion: reduce) {
  .word-app-shell.focus-mode .shell-top,
  .word-app-shell.focus-mode :deep(.word-statusbar),
  .word-app-shell.focus-mode.reveal-top .shell-top,
  .word-app-shell.focus-mode.reveal-bottom :deep(.word-statusbar),
  .focus-exit {
    transition: none;
  }
}
</style>
