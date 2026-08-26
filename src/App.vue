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
      :has-document="hasDocument"
      :is-saving="saveSnapshot.isSaving"
      :is-opening="isOpening"
      @new-doc="onNewDocument"
      @open-doc="onPickAndOpen"
      @save-doc="onSave(false)"
      @save-as-doc="onSave(true)"
      @export-doc="onExportDocx"
      @print-doc="onPrint"
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
    </div>

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
  </div>
</template>

<script setup lang="ts">
import { ref, provide, onMounted, onUnmounted, computed, shallowRef } from 'vue';
import TitleBar from './ui/shell/TitleBar.vue';
import Ribbon from './ui/ribbon/Ribbon.vue';
import StatusBar from './ui/shell/StatusBar.vue';
import DocumentRuler from './ui/shell/DocumentRuler.vue';
import VerticalRuler from './ui/shell/VerticalRuler.vue';
import FindReplaceDialog from './ui/panels/FindReplaceDialog.vue';
import AboutDialog from './ui/panels/AboutDialog.vue';
import LinkDialog from './ui/panels/LinkDialog.vue';
import ShortcutsDialog from './ui/panels/ShortcutsDialog.vue';

import { createCommandAdapter, type CommandAdapter, type CommandOutcome } from './engine/command-adapter';
import type { CommandId } from './engine/capabilities';
import { COMMAND_ADAPTER, COMMAND_REPORTER, FONT_OPTIONS, STYLE_GALLERY } from './composables/keys';
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
import { createEditor } from './engine/create-editor';
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
import { readPageMargins } from './engine/page-setup';
import { readParagraphIndents } from './engine/paragraph-format';
import type { RulerUnit } from './engine/ruler-geometry';
import {
  applyHebrewDocumentDefaults,
  applyHebrewPaperSize,
} from './engine/document-defaults';
import type { SuperDoc } from 'superdoc';
import { exportDocx, docxFileName } from './engine/export';
import { printDocument } from './engine/print';
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
import { confirm, notifyError } from './host/otzaria-client';
import {
  loadLastDocument,
  saveLastDocument,
  forgetLastDocument,
  loadAutosaveEnabled,
  saveAutosaveEnabled,
  loadRulerVisible,
  saveRulerVisible,
} from './host/settings';
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
import { startParagraphOnNewPage } from './engine/page-break';
import { createLinkDialog } from './composables/use-link-dialog';
import { createShellActionRunner } from './ui/shortcuts/actions';
import {
  createShortcutDispatcher,
  isTextEntryTarget,
  type ShortcutDispatcher,
} from './ui/shortcuts/dispatch';
import { createDirectionShortcut } from './ui/shortcuts/direction';
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
 * המופע הפתוח, בשביל הפקדים שאין להם פקודה ב-registry של ה-controller —
 * שוליים, כיוון דף, עמודות, הערות שוליים. המסלול הציבורי היחיד שלהם הוא
 * ה-Document API, והוא יושב על המופע ולא על ה-controller. ראו engine/document-api.ts.
 */
const activeSuperdoc = shallowRef<SuperDoc | null>(null);
provide(ACTIVE_SUPERDOC, activeSuperdoc);

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

/** מרכוז העמוד בזום. יחיד למאגס — אינו מוחלף בין מסמכים. */
let zoomCenter: ZoomCenter | null = null;

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

async function openDocument(file?: UserFile): Promise<boolean> {
  if (!swap) return false;
  isOpening.value = true;
  const startedAt = performance.now();
  setStatus(file ? `פותח את ${file.name}…` : 'פותח מסמך ריק…');

  const outcome = await swap.open(file?.url);
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
  editor.onDispose(() => {
    if (activeSuperdoc.value === editor.superdoc) activeSuperdoc.value = null;
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
   * מצב הסרגל מגיע מהמנוע ולא ממתג שלנו — ראו `isRulerVisible`. ההרשמה כאן
   * ולא ב-DocumentRuler.vue מפני שהיא גם **כותבת**: העדפת המשתמש נשמרת, וכל
   * מסמך שנפתח מקבל אותה בחזרה (מופע חדש נולד עם `config.rulers: false`).
   *
   * נקודה אחת ויחידה שמשנה את המצב, ולא שלוש: מצב שנקבע בהשמה ישירה במקום
   * אחד ובקריאה למודל במקום אחר הוא בדיוק איך שנוצר סרגל מוסתר שממשיך לקרוא
   * את המסמך על כל תזוזת סמן.
   */
  const applyRulerVisible = (active: boolean): void => {
    isRulerVisible.value = active;
    sessionRuler.setEnabled(active);
    if (rulerPreference === active) return;
    rulerPreference = active;
    void saveRulerVisible(active);
  };

  editor.onDispose(
    adapter.observe('ruler', (state) => {
      if (isRulerVisible.value !== state.active) applyRulerVisible(state.active);
    })
  );
  applyRulerVisible(adapter.getState('ruler').active);
  // ההעדפה חלה כאן, אחרי שה-observe רשום: `run` מחליף את הדגל במנוע, וההודעה
  // חוזרת דרך אותו מסלול שהכפתור ברצועה עובר בו. `rulerPreference` נקרא לפני
  // ההחלה, מפני שהיא עצמה כותבת אותו.
  const wanted = rulerPreference;
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
  // של הפסקה שהסמן בה.
  editor.onDispose(
    editor.ui.selection.observe(() => {
      sessionMetrics.noteSelectionChanged();
      sessionRuler.noteSelectionChanged();
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
    })
  );

  // מדידה ראשונה, בלי להמתין לעריכה: מסמך שנפתח צריך להציג את מספר המילים
  // שלו. אם הפאסדה עוד לא מוכנה, הניסיון החוזר תלוי במעבר הפריסה הראשון.
  sessionMetrics.measureNow();
  // אותו טעם, ואותו רגע: סרגל שנפתח על מסמך חדש צריך את השוליים שלו מיד ולא
  // אחרי תזוזת הסמן הראשונה.
  sessionRuler.refreshNow();

  const sessionSearch = createSearchAdapter(editor.ui);
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

  if (file) {
    void saveLastDocument({
      token: file.token,
      name: file.name,
      writable: file.access === 'readwrite',
    });
  } else {
    void forgetLastDocument();
    // גודל הדף לפני הכיווניות: `sections.setPageSetup` כותב את אותו `sectPr`
    // ש-`setSectionDirection` כותב אליו, וכך הכיווניות היא זו שנכתבת אחרונה.
    // גם הסדר של ההודעות נגזר מזה — כשל כיווניות הוא החמור, והוא זה שיישאר
    // בשורת המצב אם שניהם נכשלו.
    await applyNewDocumentPaperSize(editor.superdoc);
    await applyNewDocumentDirection(editor.superdoc);
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

  return true;
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
    void saveLastDocument({ token: outcome.token, name: outcome.name, writable: true });
    setStatus(`${title.value} נשמר`);
  }
}

async function onPickAndOpen(): Promise<void> {
  if (isOpening.value) return;
  try {
    const file = await pickDocxFile();
    if (!file) return;

    if (save && swap) {
      const decision = await decideDocumentSwitch({
        isDirty: () => save!.snapshot.isDirty,
        isSaving: () => save!.snapshot.isSaving,
        confirm,
        documentName: () => title.value,
      });

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

function onUndo(): void {
  void commandAdapter.value?.run('undo');
}

function onRedo(): void {
  void commandAdapter.value?.run('redo');
}

function onTitleUpdate(newTitle: string): void {
  if (newTitle.trim()) {
    title.value = newTitle.trim();
    save?.markDirty();
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
  // יציאה ממצב מיקוד מאפסת את החשיפה: אחרת המחלקה נשארת והפסים מקבלים
  // opacity מיותר ברגע שחוזרים למצב הרגיל.
  if (!isFocusMode.value) revealed.value = null;
}

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
 * חלון באנגלית בעיצוב של SuperDoc שאין דרך ציבורית לתרגם או לעצב. הפעולות
 * עצמן כן עוברות דרך `ui.search` — אותה שכבה שה-surface הזה נשען עליה — ולכן
 * אין כאן מימוש חיפוש מקביל.
 */
function openFindDialog(mode: 'find' | 'replace'): void {
  findMode.value = mode;
  isFindOpen.value = true;
  // פתיחת הדיאלוג פותחת session במנוע; בלעדיו `search` נכשל סגור.
  reportSearch(searchAdapter?.open());
}

function closeFindDialog(): void {
  isFindOpen.value = false;
  // סגירה מנקה את ההדגשות במסמך. בלעדיה הן נשארות אחרי שהדיאלוג נעלם.
  searchAdapter?.close();
}

/** התוצאה של כל פעולת חיפוש עוברת כאן: כשל לשורת המצב, הצלחה למונה. */
function reportSearch(outcome: SearchOutcome | undefined): void {
  if (!outcome) {
    setStatus('אין מסמך פתוח לחיפוש', true);
    return;
  }
  if (!outcome.ok) {
    setStatus(outcome.message, true);
    return;
  }
  searchState.value = outcome.snapshot;
}

function onFindText(query: string, direction: 'next' | 'prev'): void {
  reportSearch(searchAdapter?.find(query, direction));
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
    if (isAboutOpen.value) {
      isAboutOpen.value = false;
      return true;
    }
    if (linkDialog.isOpen.value) {
      linkDialog.close();
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

  reportSearch(searchAdapter?.find(query, direction));
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

/** פקודת מנוע שמגיעה מקיצור. אותו מסלול, ואותו דיווח, כמו לחיצת כפתור. */
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

/** דיאלוג שמכריז `aria-modal`. מה שמאחוריו אינו זמין — גם לא לקיצור. */
function isModalDialogOpen(): boolean {
  return isAboutOpen.value || linkDialog.isOpen.value || isShortcutsHelpOpen.value;
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
    isBlocked: (target) =>
      isModalDialogOpen() || (isTextEntryTarget(target) && !isDocumentSurface(target)),
  });

  if (editorStackRef.value) {
    // לפני פתיחת המסמך הראשון: `observeZoom` יורה מיד עם ה-snapshot, וללא
    // הפקד הזה הדיווח הראשון היה הולך לאיבוד.
    zoomCenter = createZoomCenter(editorStackRef.value);

    save = initSaveCoordinator();

    // הבחירה נטענת לפני שנפתח מסמך: העריכה הראשונה עלולה להתחיל סבב autosave,
    // ואם ההעדפה עוד לא הגיעה הוא היה רץ לפי ברירת המחדל ולא לפי מה שהמשתמש
    // בחר בהפעלה הקודמת.
    autosaveEnabled.value = await loadAutosaveEnabled();
    save.setAutosaveEnabled(autosaveEnabled.value);

    // גם ההעדפה של הסרגל, ומאותו טעם: היא חלה על המסמך שנפתח מיד אחרי כאן.
    rulerPreference = await loadRulerVisible();

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
        },
        // ה-callback נרשם פעם אחת, כאן, ולכן הוא מפנה למודד הנוכחי ולא
        // ל-session מסוים — בדיוק כמו `save?.markDirty()` שמעליו.
        onPaginationUpdate: (totalPages) => metrics?.notePaginationUpdate(totalPages),
      })
    );

    // טעינת מסמך אחרון או פתיחת מסמך ריק
    const last = await resolveLastDocument();
    if (!last) {
      await openDocument();
    } else if (!(await openDocument(last))) {
      void forgetLastDocument();
      await openDocument();
      setStatus('המסמך האחרון לא נפתח — נפתח מסמך חדש');
    }
  }
});

onUnmounted(() => {
  zoomCenter?.dispose();
  zoomCenter = null;
  shortcuts?.dispose();
  directionShortcut?.dispose();
  // חיפוש-בזמן-הקלדה שממתין ירוץ אחרי הפירוק על handle של controller מפורק.
  searchAdapter?.dispose();
});

async function resolveLastDocument(): Promise<UserFile | undefined> {
  const last = await loadLastDocument();
  if (!last) return undefined;

  const file = await resolveFileUrl(last.token);
  if (!file) {
    void forgetLastDocument();
    return undefined;
  }
  return { ...file, name: file.name || last.name, access: last.writable ? 'readwrite' : 'read' };
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
