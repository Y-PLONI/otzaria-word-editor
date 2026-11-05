/**
 * המעטפת עצמה — `App.vue` — מורכבת ונלחצת.
 *
 * ## למה הקובץ הזה קיים
 *
 * 915 שורות של חיווט מעטפת (שמירה, פתיחה, autosave, שם המסמך, דיאלוגים,
 * מטריקות) היו מאומתות **רק** ב-`readFileSync` + regex. סריקת מקור אינה יכולה
 * להבחין בין „הפונקציה קיימת” ל„הפונקציה מחוברת לפקד ומגיעה למי שאמור לענות
 * לה”, וזה נמדד: מוטציה שהסירה את `save?.setAutosaveEnabled(...)` מ-
 * `toggleAutosave` — כלומר החזירה את מתג השמירה האוטומטית להיות דקורטיבי,
 * הבאג המקורי בדיוק — עברה 203 בדיקות בירוק.
 *
 * ## הכפילים, ומה שנשאר אמיתי
 *
 * `onMounted` של המעטפת מקים מנוע SuperDoc אמיתי, ולכן מוחלפים בכפיל בדיוק
 * הדברים שאין להם קיום ב-jsdom או שהתשובה שלהם היא מה שנבדק:
 *
 *   * `engine/create-editor` — מייבא `superdoc` ו-workers. לא מגיעים אליו כאן
 *     בכלל (ה-swap מוחלף), אבל הייבוא הסטטי לבדו מפיל את ההרכבה.
 *   * `sessions/editor-swap` — מחזיר session מזויף שמצליח מיד.
 *   * `sessions/save-coordinator` — **זה מה שנמדד**: כל קריאה אליו מוקלטת, וגם
 *     ה-deps שלו נשמרים כדי שהבדיקה תוכל לדחוף snapshot ולראות מה הפס מציג.
 *   * `engine/command-adapter`, `engine/search`, `engine/doc-metrics`,
 *     `engine/document-defaults` — נשענים על handle של מנוע חי.
 *   * `host/settings` — כדי שהמתג יימדד גם על השאלה אם הבחירה נשמרה.
 *
 * מה שנשאר אמיתי: `TitleBar`, `Ribbon`, `StatusBar` והדיאלוגים, כל הזרימה של
 * `openDocument`, מטפל המקלדת, ו-`host/files`/`otzaria-client` (הם ניגשים
 * ל-`window.Otzaria` שאינו קיים ומחזירים כשל בשקט — בדיוק כמו בדפדפן).
 */
import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest';
import { mount } from '@vue/test-utils';
import { shallowRef } from 'vue';
import {
  autoUnmount,
  buttonByTip,
  createCommandDouble,
  createSuperdocDouble,
  settle,
  type CommandDouble,
  type SuperdocDouble,
} from './harness';
import type { SaveCoordinatorDeps, SaveSnapshot } from '../../src/sessions/save-coordinator';
import { COMPLEX_SCRIPT_BOLD_NOTICE } from '../../src/engine/docx-preflight';
import { NO_VBA, type DocumentVba } from '../../src/engine/vba-import';
import { COMMAND_REPORTER, STATUS_NOTIFIER, type CommandReporter } from '../../src/composables/keys';
import type { DocMetrics } from '../../src/engine/doc-metrics';
import type { MacrosHandle } from '../../src/engine/macros';
import FindReplaceDialog from '../../src/ui/panels/FindReplaceDialog.vue';
import TitleBar from '../../src/ui/shell/TitleBar.vue';
import Ribbon from '../../src/ui/ribbon/Ribbon.vue';
import DocumentRuler from '../../src/ui/shell/DocumentRuler.vue';
import VerticalRuler from '../../src/ui/shell/VerticalRuler.vue';

/**
 * המצב המשותף לכפילים. `vi.hoisted` נדרש: מפעלי ה-`vi.mock` מורמים אל מעל
 * הייבואים ורצים לפני גוף הקובץ, ולכן משתנה רגיל בהיקף המודול היה TDZ.
 */
const stub = vi.hoisted(() => ({
  /** מה שהקואורדינטור קיבל. `[]` אחרי מוטציה שמנתקת את המתג. */
  autosaveCalls: [] as boolean[],
  /** מה ש-`saveNow` קיבל, לפי הסדר. */
  saveNowCalls: [] as Array<{ forceSaveAs?: boolean; suggestedName?: string } | undefined>,
  markDirtyCalls: 0,
  /** האם „החלף הכל” של הכפיל זורק. */
  replaceAllFails: false,
  replaces: [] as string[],
  replaceAlls: [] as string[],
  searchAvailable: false,
  macros: null as MacrosHandle | null,
  /** כמה פעמים המעטפת סימנה „שוחזר מטיוטה” — ולא „נערך”. */
  markRestoredCalls: 0,
  /** הקפאות ושחרורים של פעולה כבדה, לפי הסדר. */
  holdCalls: 0,
  releaseCalls: 0,
  resetCalls: 0,
  /** מה שנשמר להפעלה הבאה. */
  persistedAutosave: [] as boolean[],
  /** מה שההעדפה השמורה מחזירה בעלייה. */
  storedAutosave: true,
  /** מצב הסרגל שנשמר בהפעלה הקודמת. */
  storedRuler: false,
  /** מה שנכתב לאחסון בשביל הסרגל, לפי הסדר. */
  persistedRuler: [] as boolean[],
  /** רשומת ההפעלה שה-storage מחזיר בעלייה. */
  storedSession: null as unknown,
  /** כל מה שנכתב לרשומת ההפעלה, לפי הסדר. */
  persistedSessions: [] as unknown[],
  searchOpens: 0,
  /** ה-deps שהמעטפת נתנה לקואורדינטור — דרך לדחוף snapshot כמו המנוע. */
  saveDeps: null as SaveCoordinatorDeps | null,
  /** ה-session שה-swap „פתח”. מוגדר בכל בדיקה מחדש. */
  session: null as unknown,
  /** כפיל המופע שבתוך ה-session, כדי לראות מה המעטפת ביקשה מהמנוע. */
  superdoc: null as SuperdocDouble | null,
  /** האדפטר שהמעטפת תזריק לרצועה. */
  adapter: null as unknown,
  /** מה שכל פתיחה קיבלה כמקור: URL, Blob, או undefined למסמך ריק. */
  openSources: [] as unknown[],
  /**
   * מה ש-`resolveFileUrl` מחזיר — `null` = הקובץ אינו נגיש יותר. פונקציה
   * כשהתרחיש מבחין בין tokens (ריבוי טאבים): היא מקבלת את ה-token המבוקש.
   */
  resolvedFile: null as unknown,
  /** בייטי הטיוטה שבמרחב הפרטי, או `null` כשאין. */
  draftBytes: null as Uint8Array | null,
  /** השהיה יזומה של קריאה מהמרחב הפרטי, לבדיקות תחרות פתיחה. */
  workspaceReadGate: null as Promise<void> | null,
  /**
   * טיוטות לפי נתיב — לתרחיש ריבוי הטאבים, שבו לכל טאב נתיב טיוטה משלו
   * (`draftPathFor`). נתיב שאינו כאן נופל ל-`draftBytes`.
   */
  draftsByPath: {} as Record<string, Uint8Array>,
  /** רשימת „המסמכים האחרונים” שה-storage מחזיר בעלייה. */
  storedRecents: null as unknown,
  /** כל מה שנכתב לרשימת האחרונים, לפי הסדר. */
  persistedRecents: [] as unknown[],
  /** כל כתיבה למרחב הפרטי — נתיב ובייטים. הטיוטה והגיבוי עוברים בו. */
  workspaceWrites: [] as Array<{ path: string; bytes: Uint8Array }>,
  /** רשומת הגיבוי של „לא לשמור” שה-storage מחזיר בעלייה. */
  storedDiscardBackups: null as unknown,
  discardHistoryFailure: false,
  /** כל מה שנכתב לרשומת הגיבוי, לפי הסדר. */
  persistedDiscardBackups: [] as unknown[],
  /** כמה פעמים נמחקה הטיוטה. */
  draftRemovals: 0,
  /** הנתיבים שנמחקו, לפי הסדר — „איזו טיוטה” ולא רק „כמה”. */
  removedDrafts: [] as string[],
  /** מה ש-`ui.selection.apply` קיבל — כלומר לאן הסמן הוחזר. */
  caretApplied: [] as unknown[],
  /** המפתח הישן, בשביל מסלול השדרוג. */
  lastDocument: null as unknown,
  /** האם המפתח הישן נמחק אחרי שהומר. */
  forgotLastDocument: false,
  /** כמה פתיחות הבאות ייכשלו. */
  openFailures: 0,
  /**
   * מה שהשלב המקדים מדווח על המסמך שנפתח. `notice` הוא ההודעה על תיקון שנכתב
   * לתוך המסמך; `vba` הוא המאקרו שנקרא ממנו. שניהם מתחרים על שורת המצב עם
   * „פתוח לקריאה”, וסדר העדיפות ביניהם הוא מה שנמדד למטה.
   */
  preflight: { notice: null as string | null, vba: null as DocumentVba | null },
  preflightGate: null as Promise<void> | null,
  preflightSignals: [] as AbortSignal[],
  /** כמה מנועים שוחררו כדי לעמוד בתקרת הזיכרון (`swap.close`). */
  closedEngines: 0,
  /**
   * כפילים **שונים לכל פתיחה** — לבדיקות הקישוריות בין טאבים.
   *
   * ברירת המחדל (`stub.session`, `stub.adapter`) היא אובייקט אחד לכל
   * הפתיחות, ולכן „הרצועה מדברת עם המנוע של הטאב הפעיל” אינו ניתן למדידה
   * איתה: מחיקת `commandAdapter.value = ui.commandAdapter` מהשחזור הייתה
   * עוברת בירוק, כי ה-ref היה מחזיק אותו אובייקט ממילא. כשה-factory מוגדר,
   * כל `swap.open` וכל `createCommandAdapter` מקבלים מופע חדש.
   */
  sessionFactory: null as (() => unknown) | null,
  adapterFactory: null as (() => unknown) | null,
  /** מטריקות לכל מודל מטריקות שנוצר — כדי ששני טאבים יראו מספרים שונים בפס. */
  metricsFactory: null as (() => DocMetrics) | null,
  /** ה-deps של **כל** קואורדינטור שנוצר, לפי סדר הטאבים — `saveDeps` מחזיק רק את האחרון. */
  saveDepsList: [] as SaveCoordinatorDeps[],
  /** סדר השחרור של מודלי המסמך, כפי שהמנוע המזויף הריץ אותם. */
  disposeOrder: [] as string[],
  /** קריאות `refreshNow()` על מודלי „גבולות עמוד” ו„מספרי שורות” — מה ש-`reportCommand` מפעיל. */
  refreshCalls: [] as string[],
}));

vi.mock('../../src/engine/create-editor', () => ({
  createEditor: vi.fn(),
  OPEN_TIMEOUT_MS: 1_000,
}));

/**
 * כפיל ה-swap. `current` הוא פר-מופע ולא גלובלי, בשונה מהעבר: התקרה על
 * מספר המנועים החיים (`MAX_LIVE_DOCUMENTS`) נמדדת בדיוק בשאלה „לכמה טאבים
 * יש `swap.current`”, וגטר משותף היה עונה עליה תמיד „לכולם”.
 */
vi.mock('../../src/sessions/editor-swap', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../src/sessions/editor-swap')>()),
  createEditorSwap: (container: HTMLElement) => {
    let current: unknown = null;
    /**
     * ה-host שהמנוע יושב בו הוא גם מיכל הגלילה של המסמך (shell.css), והוא
     * מה שהמעטפת קוראת ממנו את מיקום הגלילה. הכפיל יוצר אותו כמו האמיתי —
     * בלעדיו „הגלילה נשמרת במעבר טאב” לא היה ניתן למדידה בכלל.
     */
    let host: HTMLElement | null = null;
    const dropHost = (): void => {
      host?.remove();
      host = null;
    };
    return {
      get current() {
        return current;
      },
      get isOpening() {
        return false;
      },
      // These tests hold preflight before swap.open creates an engine candidate.
      cancel: () => false,
      open: async (source?: unknown) => {
        stub.openSources.push(source);
        if (stub.openFailures > 0) {
          stub.openFailures -= 1;
          return { status: 'failed', error: new Error('worker לא עלה') };
        }
        current = stub.sessionFactory ? stub.sessionFactory() : stub.session;
        if (!host) {
          host = document.createElement('div');
          host.className = 'editor-stack__host';
          container.appendChild(host);
        }
        return { status: 'opened', session: current };
      },
      // `destroy()` של ה-session המזויף לפני האיפוס — כמו ה-swap האמיתי, שקורא
      // ל-`session.destroy()` וזה מריץ את ה-disposers שהמעטפת רשמה (create-editor.ts).
      close: () => {
        stub.closedEngines += 1;
        (current as { destroy?: () => void } | null)?.destroy?.();
        current = null;
        dropHost();
      },
      destroy: () => {
        (current as { destroy?: () => void } | null)?.destroy?.();
        current = null;
        dropHost();
      },
    };
  },
}));

/**
 * השלב המקדים אינו יכול לרוץ כאן — הוא מושך את המסמך מ-URL ופותח את ה-zip —
 * ולכן הוא מוחלף בכפיל שמחזיר את המקור כמות שהוא, עם מה שהבדיקה קבעה. הקבוע
 * של ההודעה נשאר אמיתי: הבדיקה מודדת שהנוסח **הזה** מגיע לשורת המצב.
 */
vi.mock('../../src/engine/docx-preflight', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../src/engine/docx-preflight')>()),
  preflightSource: async (source: unknown, options?: { signal?: AbortSignal }) => {
    if (options?.signal) stub.preflightSignals.push(options.signal);
    if (stub.preflightGate) await stub.preflightGate;
    return {
      source,
      fontTable: null,
      vba: stub.preflight.vba ?? { hasMacroPart: false, modules: [], autoRun: [], warnings: [], status: null, unreadable: false },
      notice: stub.preflight.notice,
    };
  },
}));

vi.mock('../../src/host/files', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../src/host/files')>()),
  resolveFileUrl: async (token: string) =>
    typeof stub.resolvedFile === 'function'
      ? (stub.resolvedFile as (token: string) => unknown)(token)
      : stub.resolvedFile,
}));

vi.mock('../../src/host/workspace', () => ({
  MAX_PAYLOAD_BYTES: 10,
  MAX_CONTENT_BYTES: 7,
  readWorkspaceBytes: async (path: string) => {
    await stub.workspaceReadGate;
    return Object.prototype.hasOwnProperty.call(stub.draftsByPath, path)
      ? stub.draftsByPath[path]
      : stub.draftBytes;
  },
  /**
   * `'written'` ולא `true`: זה מה ש-`WorkspaceWrite` האמיתי מחזיר, ושני
   * הצרכנים — `SessionKeeper.writeDraftNow` והגיבוי של „לא לשמור” — משווים
   * אליו. כפיל שמחזיר `true` נקרא אצל שניהם ככשל, ולכן הוא היה מסתיר בדיוק
   * את המסלול שנבדק כאן.
   */
  writeWorkspaceBytes: async (path: string, bytes: Uint8Array) => {
    stub.workspaceWrites.push({ path, bytes });
    return 'written';
  },
  deleteWorkspaceEntry: async (path: string) => {
    stub.draftRemovals += 1;
    stub.removedDrafts.push(path);
    // מחיקה שמוחקת: בלעדיה „הגיבוי נכתב לפני שהטיוטה נמחקה” אינו ניתן
    // למדידה — קריאה אחרי המחיקה הייתה מחזירה את אותם בייטים.
    delete stub.draftsByPath[path];
  },
}));

vi.mock('../../src/sessions/save-coordinator', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../src/sessions/save-coordinator')>()),
  createSaveCoordinator: (deps: SaveCoordinatorDeps) => {
    stub.saveDeps = deps;
    stub.saveDepsList.push(deps);
    const snapshot: SaveSnapshot = {
      state: 'idle', isDirty: false, isSaving: false,
      targetToken: null, name: null, lastError: null,
    };
    return {
      get snapshot() { return snapshot; },
      markDirty: () => {
        stub.markDirtyCalls += 1;
      },
      markRestored: () => {
        stub.markRestoredCalls += 1;
      },
      hold: () => {
        stub.holdCalls += 1;
        return () => {
          stub.releaseCalls += 1;
        };
      },
      isHeld: false,
      setAutosaveEnabled: (enabled: boolean) => {
        stub.autosaveCalls.push(enabled);
      },
      adoptTarget: () => {},
      reset: (target?: { token: string; name: string } | null) => {
        stub.resetCalls += 1;
        snapshot.targetToken = target?.token ?? null;
        snapshot.name = target?.name ?? null;
        deps.onStateChange?.({ ...snapshot });
      },
      saveNow: async (options?: { forceSaveAs?: boolean; suggestedName?: string }) => {
        stub.saveNowCalls.push(options);
        // הכפיל מודיע על השמירה כמו הקואורדינטור האמיתי. בלי זה כל בדיקה
        // כאן הייתה מודדת מעטפת שאינה שומעת שמירה אוטומטית בכלל.
        await deps.exportDocument();
        deps.onSaved?.({ token: 'token-1', name: 'מסמך.docx', size: 4_096 });
        return { status: 'saved', token: 'token-1', name: 'מסמך.docx', size: 4_096 };
      },
      dispose: () => {},
    };
  },
}));

vi.mock('../../src/engine/command-adapter', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../src/engine/command-adapter')>()),
  createCommandAdapter: () => (stub.adapterFactory ? stub.adapterFactory() : stub.adapter),
}));

vi.mock('../../src/engine/macros', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../src/engine/macros')>();
  return {
    ...actual,
    installMacros: (...args: Parameters<typeof actual.installMacros>) => stub.macros ?? actual.installMacros(...args),
  };
});

vi.mock('../../src/engine/search', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../src/engine/search')>();
  const state = () => ({ ...actual.idleSearchState(), available: stub.searchAvailable, canReplace: stub.searchAvailable });
  return {
    ...actual,
    createSearchAdapter: () => ({
      getState: state,
      subscribe: () => () => {},
      open: () => {
        stub.searchOpens += 1;
        return { ok: true, snapshot: state() };
      },
      close: () => {},
      clear: () => {},
      find: () => ({ ok: true, snapshot: actual.idleSearchState() }),
      findDebounced: () => {},
      replace: async (replacement: string) => {
        stub.replaces.push(replacement);
        return { ok: true, snapshot: actual.idleSearchState() };
      },
      replaceAll: async (replacement: string) => {
        stub.replaceAlls.push(replacement);
        if (stub.replaceAllFails) throw new Error('ההחלפה נכשלה');
        return { ok: true, snapshot: actual.idleSearchState() };
      },
      dispose: () => {
        stub.disposeOrder.push('search');
      },
    }),
  };
});

vi.mock('../../src/engine/doc-metrics', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../src/engine/doc-metrics')>();
  return {
    ...actual,
    createDocMetrics: () => {
      // מצב לכל מודל ולא קריאה גלובלית: שני טאבים עם מספרים שונים הם מה
      // שמאפשר לראות בפס המצב איזה מסמך המעטפת מציגה באמת אחרי מעבר.
      const state = stub.metricsFactory ? stub.metricsFactory() : actual.emptyDocMetrics();
      return {
        getState: () => state,
        noteDocumentChanged: () => {},
        noteSelectionChanged: () => {},
        notePaginationUpdate: () => {},
        measureNow: () => {},
        dispose: () => {
          stub.disposeOrder.push('metrics');
        },
      };
    },
  };
});

/**
 * „גבולות עמוד” ו„מספרי שורות” — הכפיל קיים בשביל שני דברים שאין להם עדות
 * אחרת: `refreshNow()` ש-`reportCommand` מפעיל אחרי בחירה בתפריט (בלעדיו גבול
 * שנבחר אינו מצטייר עד העריכה הבאה — הערה ב-App.vue), וסדר השחרור בסגירת טאב.
 * אף בדיקה אחרת בעץ אינה מזכירה `page-borders`; הכיסוי היחיד היה גשש CDP יתום.
 */
vi.mock('../../src/engine/page-setup', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../src/engine/page-setup')>();
  const model = (name: string) => ({
    getState: () => null,
    refreshNow: () => {
      stub.refreshCalls.push(name);
    },
    noteDocumentChanged: () => {},
    dispose: () => {
      stub.disposeOrder.push(name);
    },
  });
  return {
    ...actual,
    createPageBorderModel: () => model('pageBorders'),
    createLineNumberingModel: () => model('lineNumbers'),
  };
});

vi.mock('../../src/engine/document-defaults', () => ({
  applyHebrewDocumentDefaults: async () => ({ failures: [] }),
  applyHebrewPaperSize: async () => ({ applied: true }),
  verifyHebrewDocumentDefaults: async () => false,
}));

vi.mock('../../src/host/settings', () => ({
  loadLastDocument: async () => stub.lastDocument,
  forgetLastDocument: async () => {
    stub.forgotLastDocument = true;
  },
  loadAutosaveEnabled: async () => stub.storedAutosave,
  saveAutosaveEnabled: async (enabled: boolean) => {
    stub.persistedAutosave.push(enabled);
  },
  loadRulerVisible: async () => stub.storedRuler,
  saveRulerVisible: async (visible: boolean) => {
    stub.persistedRuler.push(visible);
  },
  loadSessionRecord: async () => stub.storedSession,
  saveSessionRecord: async (value: unknown) => {
    stub.persistedSessions.push(value);
  },
  loadSpellcheckEnabled: async () => false,
  loadListAutoformatEnabled: async () => true,
  saveListAutoformatEnabled: async () => {},
  loadNavigationPaneVisible: async () => false,
  saveNavigationPaneVisible: async () => {},
  loadAutoCorrect: async () => null,
  loadNewDocumentDefaults: async () => null,
  saveNewDocumentDefaults: async () => {},
  saveAutoCorrect: async () => {},
  saveSpellcheckEnabled: async () => {},
  loadCanvasColor: async () => null,
  loadPasteDefault: async () => null,
  savePasteDefault: async () => {},
  saveCanvasColor: async () => {},
  loadSpellcheckWords: async () => [],
  saveSpellcheckWords: async () => {},
  loadRecentDocuments: async () => stub.storedRecents,
  saveRecentDocuments: async (list: unknown) => {
    stub.persistedRecents.push(list);
  },
  loadDiscardBackups: async () => {
    if (stub.discardHistoryFailure) throw new Error('offline');
    return stub.storedDiscardBackups;
  },
  saveDiscardBackups: async (list: unknown) => {
    stub.persistedDiscardBackups.push(list);
  },
  loadCustomShortcuts: async () => null,
  saveCustomShortcuts: async () => {},
  // העץ של „פתח מסמך” — `useOpenSources().init()` רץ בעלייה, ומוק בלי
  // הפונקציות האלה זורק מתוכו דחייה לא-מטופלת (כך נפל ה-CI של #78).
  loadDocFolders: async () => null,
  saveDocFolders: async () => {},
  loadLibraryCache: async () => null,
  saveLibraryCache: async () => {},
  loadOpenDialogPlace: async () => null,
  saveOpenDialogPlace: async () => {},
}));

// הייבוא **אחרי** ה-mocks במכוון (הם מורמים בכל מקרה, וזה הסדר שקורא נכון).
const { default: App } = await import('../../src/App.vue');

autoUnmount();

/** מרכיבה את המעטפת ומחזירה בקרה רק אחרי שכל זרימת ה-`onMounted` נרגעה. */
async function mountShell(): Promise<ReturnType<typeof mount>> {
  const wrapper = mount(App, { attachTo: document.body });
  // זרימת העלייה היא שרשרת של await-ים (העדפה, swap, openDocument, ברירות
  // מחדל של מסמך חדש), ולכן nextTick אחד אינו מספיק.
  await settle(12);
  return wrapper;
}

/** שני קבצים לתרחישי הטאבים — ראו „קיצורי הטאבים” בסוף הקובץ. */
const FIRST_DOC = { token: 'tok-1', name: 'ראשון.docx', writable: true };
const SECOND_DOC = { token: 'tok-2', name: 'שני.docx', writable: true };
const TAB_FILES: Record<string, unknown> = {
  'tok-1': { token: 'tok-1', url: 'loopback://first', name: 'ראשון.docx', size: 100 },
  'tok-2': { token: 'tok-2', url: 'loopback://second', name: 'שני.docx', size: 200 },
};

/** `Ctrl+<code>` על החלון — בדיוק כמו הקשה של המשתמש. */
function pressCtrl(code: string): void {
  window.dispatchEvent(
    new KeyboardEvent('keydown', { code, ctrlKey: true, cancelable: true, bubbles: true }),
  );
}

/** הרשומה האחרונה שנכתבה לגיבוי „לא לשמור”. */
function lastDiscardBackups(): Array<Record<string, unknown>> {
  const written = stub.persistedDiscardBackups;
  return (written[written.length - 1] ?? []) as Array<Record<string, unknown>>;
}

/** חלון „המסמך לא נשמר”, או `null` כשאינו פתוח. */
function unsavedDialog(): Element | null {
  return document.querySelector('.unsaved-dialog');
}

/**
 * עונה על „המסמך לא נשמר”. שאלה אחת ושלושה כפתורים — עד כאן היו כאן שתי
 * שאלות של אוצריא זו אחר זו, והתשובה נמסרה דרך `confirm` המזויף.
 */
async function answerUnsaved(choice: 'save' | 'discard' | 'cancel'): Promise<void> {
  const button = document.querySelector<HTMLButtonElement>(`[data-choice="${choice}"]`);
  if (!button) throw new Error(`„${choice}” אינו מוצג — הדיאלוג אינו פתוח, או שאין בו כפתור כזה`);
  button.click();
  await settle(12);
}

beforeEach(() => {
  stub.autosaveCalls.length = 0;
  stub.saveNowCalls.length = 0;
  stub.persistedAutosave.length = 0;
  stub.persistedSessions.length = 0;
  stub.storedSession = null;
  stub.openSources.length = 0;
  stub.resolvedFile = null;
  stub.draftBytes = null;
  stub.workspaceReadGate = null;
  for (const path of Object.keys(stub.draftsByPath)) delete stub.draftsByPath[path];
  stub.draftRemovals = 0;
  stub.replaceAllFails = false;
  stub.replaces.length = 0;
  stub.replaceAlls.length = 0;
  stub.searchAvailable = false;
  stub.macros = null;
  stub.markRestoredCalls = 0;
  stub.holdCalls = 0;
  stub.releaseCalls = 0;
  stub.removedDrafts.length = 0;
  stub.storedRecents = null;
  stub.persistedRecents.length = 0;
  stub.storedDiscardBackups = null;
  stub.discardHistoryFailure = false;
  stub.persistedDiscardBackups.length = 0;
  stub.workspaceWrites.length = 0;
  stub.caretApplied.length = 0;
  stub.lastDocument = null;
  stub.forgotLastDocument = false;
  stub.openFailures = 0;
  stub.preflight = { notice: null, vba: null };
  stub.preflightGate = null;
  stub.preflightSignals.length = 0;
  stub.closedEngines = 0;
  stub.markDirtyCalls = 0;
  stub.resetCalls = 0;
  stub.searchOpens = 0;
  stub.storedAutosave = true;
  stub.storedRuler = false;
  stub.persistedRuler.length = 0;
  stub.saveDeps = null;
  stub.saveDepsList.length = 0;
  stub.sessionFactory = null;
  stub.adapterFactory = null;
  stub.metricsFactory = null;
  stub.disposeOrder.length = 0;
  stub.refreshCalls.length = 0;
  stub.adapter = createCommandDouble({ states: { 'document-mode': { value: 'editing' } } });
  stub.superdoc = createSuperdocDouble();
  const adapter = stub.adapter as CommandDouble;
  stub.superdoc.host.setDocumentMode = vi.fn(mode => adapter.setState('document-mode', { value: mode }));
  // גם שמירה מזויפת עוברת בייצוא, כדי לתפוס את נקודת השחזור לפני commit.
  Object.assign(stub.superdoc.host, {
    export: async () => new Blob([new Uint8Array(4_096)]),
  });
  stub.session = {
    superdoc: stub.superdoc.host,
    container: document.createElement('div'),
    // ה-controller המזויף: רק מה שהמעטפת נוגעת בו ישירות. שאר הקוראים
    // (`observeZoom`, `observeFontOptions`, `observeStyleGallery`) מתוכננים
    // ליפול לברירת מחדל כשה-handle חסר, וזה מה שנמדד בבדיקות שלהם.
    ui: {
      selection: {
        observe: () => () => {},
        // מה ששחזור הסמן נשען עליו. `apply` מקליט את מה שהוא קיבל, ומצליח
        // תמיד — השאלה כאן היא אם המעטפת חיווטה אותו, לא אם המנוע יודע.
        apply: (target: unknown) => {
          stub.caretApplied.push(target);
          return { ok: true };
        },
      },
      viewport: { scrollIntoView: async () => ({ success: true }) },
    },
    onDispose: () => {},
    destroy: () => {},
  };
});

describe('הרכבת המעטפת', () => {
  it('העלייה פותחת מסמך ומחווטת את הפס, הרצועה ושורת המצב', async () => {
    // בלי זה כל הבדיקות למטה יכולות לעבור בירוק על מעטפת שלא סיימה לעלות.
    const wrapper = await mountShell();

    expect(wrapper.find('.word-titlebar').exists()).toBe(true);
    expect(wrapper.find('.word-statusbar').exists()).toBe(true);
    expect(wrapper.find('.editor-stack').exists()).toBe(true);
    expect(stub.saveDeps, 'הקואורדינטור הוקם').not.toBeNull();
    expect(stub.resetCalls, 'הפתיחה איפסה את מצב השמירה').toBe(1);
  });

  it('המסמך שנפתח מקבל את הסמן — אפשר להקליד בלי קליק מקדים', async () => {
    // הבאג שהתיקון בא לו: העורך נפתח, ההקלדה לא הגיעה לשום מקום, והמשתמש היה
    // צריך ללחוץ עם העכבר בגוף הטקסט לפני שיכול היה לכתוב מילה.
    await mountShell();

    expect(stub.superdoc?.ops(), 'הפתיחה ביקשה מהמנוע להחזיר את הסמן').toContain('focus');
  });

  it('פתיחה אינה חוטפת את הפוקוס משדה שמקלידים בו', async () => {
    // הפתיחה אסינכרונית ויכולה להימשך שניות. אם בינתיים המשתמש הקליד בשורת
    // החיפוש (שאינה מודאלית ונשארת פתוחה מעל המסמך), קפיצה לגוף המסמך הייתה
    // מוחקת לו את ההקלדה באמצע.
    const wrapper = await mountShell();
    await wrapper.find('.search-box').trigger('click');
    await settle();
    document.querySelector<HTMLInputElement>('#fr-search-input')?.focus();
    stub.superdoc?.reset();

    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'n', code: 'KeyN', ctrlKey: true }));
    await settle(12);
    // Ctrl+N פותח את „פתח מסמך”, ו„מסמך ריק” הוא הכרטיס שפותח בפועל.
    document.querySelector<HTMLButtonElement>('.tpl-card')?.click();
    await settle(12);

    expect(stub.resetCalls, 'המסמך החדש אכן נפתח').toBe(2);
    // הדיאלוג מחזיר את המיקוד למי שפתח אותו — שדה החיפוש — ולא לגוף המסמך.
    expect(stub.superdoc?.ops(), 'הפוקוס נשאר בשדה החיפוש').not.toContain('focus');
  });
});

describe('נעילת העורך הקודם בזמן פתיחה', () => {
  function containers(): { previous: HTMLElement; next: HTMLElement } {
    const previous = document.createElement('div'), next = document.createElement('div');
    previous.inert = next.inert = false;
    const base = stub.session as object;
    let opened = 0;
    stub.sessionFactory = () => ({ ...base, container: opened++ === 0 ? previous : next });
    return { previous, next };
  }

  function holdPreflight(): () => void {
    let release!: () => void;
    stub.preflightGate = new Promise<void>(resolve => { release = resolve; });
    return release;
  }

  async function beginBlank(beforeChoose?: () => void): Promise<AbortSignal> {
    const count = stub.preflightSignals.length;
    pressCtrl('KeyN');
    await settle(8);
    const card = document.querySelector<HTMLButtonElement>('.tpl-card');
    expect(card, 'בורר המסמך הריק פתוח').not.toBeNull();
    beforeChoose?.();
    card!.click();
    await vi.waitFor(() => expect(stub.preflightSignals).toHaveLength(count + 1));
    return stub.preflightSignals[count]!;
  }

  it('locks the real Ribbon, title actions, and rulers while preflight waits, then restores them on Skip', async () => {
    const { previous } = containers();
    const wrapper = await mountShell();
    const adapter = stub.adapter as CommandDouble;
    adapter.setState('document-mode', { value: 'suggesting' });
    const release = holdPreflight();
    const titleBar = wrapper.getComponent(TitleBar);
    const originalTitle = titleBar.props('title');
    try {
      expect((buttonByTip(wrapper, 'מודגש').element as HTMLButtonElement).disabled).toBe(false);
      await beginBlank();
      expect(previous.inert).toBe(true);
      expect(adapter.getState('document-mode').value).toBe('viewing');
      expect(wrapper.get('.shell-top').attributes('inert')).toBeDefined();
      expect(wrapper.getComponent(DocumentRuler).props('editable')).toBe(false);
      expect(wrapper.getComponent(VerticalRuler).props('editable')).toBe(false);
      const bold = buttonByTip(wrapper, 'מודגש');
      expect((bold.element as HTMLButtonElement).disabled).toBe(true);
      adapter.reset();
      await bold.trigger('click');
      titleBar.vm.$emit('undo');
      titleBar.vm.$emit('redo');
      titleBar.vm.$emit('update-title', 'must not rename the old document');
      await settle(8);
      expect(adapter.calls).toEqual([]);
      expect(titleBar.props('title')).toBe(originalTitle);
      const skip = wrapper.get('.status-load__skip');
      expect(skip.element.closest('[inert]'), 'Skip remains outside the locked shell').toBeNull();
      await skip.trigger('click');
      expect(adapter.getState('document-mode').value).toBe('suggesting');
      expect(previous.inert).toBe(false);
      expect(wrapper.getComponent(DocumentRuler).props('editable')).toBe(true);
      expect(wrapper.getComponent(VerticalRuler).props('editable')).toBe(true);
      expect((bold.element as HTMLButtonElement).disabled).toBe(false);
      await bold.trigger('click');
      expect(adapter.calls.map(call => call.id)).toEqual(['bold']);
    } finally { release(); await settle(12); }
  });

  it('Developer replay and record cannot bypass the opening lock through emitted events', async () => {
    containers();
    const replayLast = vi.fn(), toggleRecording = vi.fn();
    stub.macros = {
      kit: {
        listRecordings: () => [], listSnippets: () => [], listScripts: () => [], listTools: () => [], registerTool: () => {},
      } as unknown as MacrosHandle['kit'],
      recording: shallowRef(false), scriptsEnabled: false, replayLast, toggleRecording, dispose: () => {},
    };
    const wrapper = await mountShell();
    const ribbon = wrapper.getComponent(Ribbon);
    ribbon.vm.$emit('update:activeTab', 'developer');
    await settle(8);
    const play = wrapper.findAll('button').find(button => button.text() === 'נגן אחרון')!;
    expect(play.exists()).toBe(true);
    expect((play.element as HTMLButtonElement).disabled).toBe(false);
    const release = holdPreflight();
    try {
      await beginBlank();
      expect((play.element as HTMLButtonElement).disabled).toBe(true);
      ribbon.vm.$emit('macro-play');
      ribbon.vm.$emit('macro-record');
      await settle(8);
      expect(replayLast).not.toHaveBeenCalled();
      expect(toggleRecording).not.toHaveBeenCalled();
      await wrapper.get('.status-load__skip').trigger('click');
      expect((play.element as HTMLButtonElement).disabled).toBe(false);
      await play.trigger('click');
      expect(replayLast).toHaveBeenCalledOnce();
    } finally { release(); await settle(12); }
  });

  it('keeps suggesting→protected→Skip→unprotect state and will not open across a pending protection write', async () => {
    containers();
    const adapter = stub.adapter as CommandDouble;
    adapter.setState('document-mode', { value: 'suggesting' });
    const run = adapter.run.bind(adapter);
    adapter.run = async (id, payload) => {
      const outcome = await run(id, payload);
      if (outcome.ok && id === 'document-mode') adapter.setState(id, { value: (payload as { mode: string }).mode });
      return outcome;
    };
    let enforced = false;
    let finishSet!: () => void;
    const setReply = new Promise<void>(resolve => { finishSet = resolve; });
    const setProtection = vi.fn(async () => { await setReply; enforced = true; return { success: true }; });
    Object.assign(stub.superdoc!.host.activeEditor!.doc!.protection, {
      get: () => ({ editingRestriction: { mode: enforced ? 'readOnly' : 'none', enforced } }),
      setEditingRestriction: setProtection,
      clearEditingRestriction: () => { enforced = false; return { success: true }; },
    });
    const wrapper = await mountShell();
    wrapper.getComponent(Ribbon).vm.$emit('update:activeTab', 'review');
    await settle(8);
    const protect = wrapper.findAll('button').find(button => button.text() === 'הגבל עריכה')!;
    let release = (): void => {};
    try {
      await protect.trigger('click');
      await protect.trigger('click');
      expect(setProtection).toHaveBeenCalledOnce();
      const openings = stub.preflightSignals.length;
      pressCtrl('KeyN');
      await settle(8);
      expect(document.querySelector('.tpl-card'), 'opening waits for XML + runtime protection to settle').toBeNull();
      expect(stub.preflightSignals).toHaveLength(openings);
      expect(adapter.getState('document-mode').value).toBe('suggesting');
      finishSet();
      await settle(12);
      expect(adapter.getState('document-mode').value).toBe('viewing');
      expect(enforced).toBe(true);

      release = holdPreflight();
      await beginBlank();
      expect(adapter.getState('document-mode').value).toBe('viewing');
      await wrapper.get('.status-load__skip').trigger('click');
      await settle(12);
      expect(adapter.getState('document-mode').value).toBe('viewing');
      expect(enforced).toBe(true);
      await protect.trigger('click');
      await settle(12);
      expect(enforced).toBe(false);
      expect(adapter.getState('document-mode').value).toBe('suggesting');
    } finally { finishSet(); release(); await settle(12); }
  });

  it('an open nonmodal Replace panel preserves its query but cannot write until Skip', async () => {
    containers();
    stub.searchAvailable = true;
    const wrapper = await mountShell();
    wrapper.getComponent(Ribbon).vm.$emit('open-replace');
    await settle(8);
    const panel = wrapper.getComponent(FindReplaceDialog);
    const query = panel.get<HTMLInputElement>('#fr-search-input');
    await query.setValue('original query');
    await panel.get('#fr-replace-input').setValue('replacement');
    query.element.focus();
    const release = holdPreflight();
    try {
      await beginBlank();
      expect(panel.props('readOnly')).toBe(true);
      const replaceButtons = panel.findAll('.fr-btn').filter(button => button.text().startsWith('החלף'));
      expect(replaceButtons).toHaveLength(2);
      expect(replaceButtons.every(button => (button.element as HTMLButtonElement).disabled)).toBe(true);
      await panel.get('#fr-replace-input').trigger('keydown', { key: 'Enter' });
      panel.vm.$emit('replace', 'must not write');
      panel.vm.$emit('replace-all', 'must not write');
      await settle(8);
      expect(stub.replaces).toEqual([]);
      expect(stub.replaceAlls).toEqual([]);
      expect(stub.holdCalls).toBe(0);
      query.element.focus();
      await wrapper.get('.status-load__skip').trigger('click');
      expect(document.activeElement).toBe(query.element);
      expect(query.element.value).toBe('original query');
      expect(panel.get<HTMLInputElement>('#fr-replace-input').element.value).toBe('replacement');
      expect(panel.props('readOnly')).toBe(false);
      await replaceButtons[0]!.trigger('click');
      expect(stub.replaces).toEqual(['replacement']);
    } finally { release(); await settle(12); }
  });

  it('נועלת רק את הישן ומשאירה את המועמד החדש זמין למיקוד', async () => {
    const { previous, next } = containers();
    const candidate = createSuperdocDouble();
    candidate.host.setDocumentMode = vi.fn();
    const base = stub.session as object;
    let opens = 0, adapters = 0;
    stub.sessionFactory = () => ++opens === 1
      ? { ...base, container: previous }
      : { ...base, superdoc: candidate.host, container: next };
    stub.adapterFactory = () => ++adapters === 1 ? stub.adapter
      : createCommandDouble({ states: { 'document-mode': { value: 'editing' } } });
    const wrapper = await mountShell();
    const release = holdPreflight();
    try {
      await beginBlank();
      expect(previous.inert).toBe(true);
      expect(next.inert).toBe(false);
      expect((wrapper.get('.editor-stack').element as HTMLElement).inert).not.toBe(true);
      stub.superdoc!.reset();
      stub.caretApplied.length = 0;
      release();
      await vi.waitFor(() => expect(stub.resetCalls).toBe(2));
      await settle(12);
      expect(previous.inert).toBe(false);
      expect(next.inert).toBe(false);
      expect(candidate.ops()).toContain('focus');
      expect(candidate.host.setDocumentMode).not.toHaveBeenCalled();
      expect(stub.caretApplied, 'the unleased candidate also gets a usable initial caret').toEqual([{
        kind: 'selection',
        start: { kind: 'text', blockId: 'block-1', offset: 0 },
        end: { kind: 'text', blockId: 'block-1', offset: 0 },
      }]);
    } finally { release(); }
  });

  it.each([false, true])('ביטול משחזר מיד את inert המקורי (%s), לפני תשובה מאוחרת', async (prior) => {
    const { previous } = containers();
    const wrapper = await mountShell();
    const release = holdPreflight();
    try {
      // Ctrl+N itself is blocked on an inert editor. Preserve a lock that was
      // applied after the chooser opened, before the opening attempt starts.
      const signal = await beginBlank(() => { previous.inert = prior; });
      expect(previous.inert).toBe(true);
      await wrapper.get('.status-load__skip').trigger('click');
      expect(signal.aborted).toBe(true);
      expect(previous.inert).toBe(prior);
      release();
      await settle(12);
      expect(stub.openSources).toHaveLength(1);
      expect(previous.inert).toBe(prior);
    } finally { release(); }
  });

  it('כשל פתיחת המנוע משחרר את העורך שנשאר', async () => {
    const { previous } = containers();
    const wrapper = await mountShell();
    const release = holdPreflight();
    const adapter = stub.adapter as CommandDouble;
    adapter.setState('document-mode', { value: 'suggesting' });
    try {
      stub.openFailures = 1;
      await beginBlank();
      expect(previous.inert).toBe(true);
      (document.activeElement as HTMLElement | null)?.blur();
      stub.superdoc!.reset();
      release();
      await vi.waitFor(() => expect(wrapper.find('.word-statusbar').text()).toContain('פתיחת המסמך נכשלה'));
      await settle(8);
      expect(previous.inert).toBe(false);
      expect(adapter.getState('document-mode').value).toBe('suggesting');
      expect(stub.superdoc!.ops()).toContain('focus');
      expect(stub.resetCalls).toBe(1);
    } finally { release(); }
  });

  it('קיצורים מה-body וכיוון פסקה אינם משנים את הישן, בעוד F6 ו־Escape נשארים זמינים', async () => {
    const { previous } = containers();
    const wrapper = await mountShell();
    const release = holdPreflight();
    const adapter = stub.adapter as CommandDouble;
    const key = (init: KeyboardEventInit, target: HTMLElement = document.body): KeyboardEvent => {
      const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
      target.dispatchEvent(event);
      return event;
    };
    const direction = (): void => {
      const init = { code: 'ShiftLeft', key: 'Shift', ctrlKey: true, bubbles: true, cancelable: true };
      document.body.dispatchEvent(new KeyboardEvent('keydown', init));
      document.body.dispatchEvent(new KeyboardEvent('keyup', init));
    };
    try {
      await beginBlank();
      adapter.reset(); stub.superdoc!.reset();
      expect(key({ code: 'KeyB', ctrlKey: true }).defaultPrevented).toBe(true);
      expect(key({ code: 'Enter', ctrlKey: true }).defaultPrevented).toBe(true);
      direction();
      await settle(8);
      expect(adapter.calls).toEqual([]);
      expect(stub.superdoc!.ops()).not.toContain('format.paragraph.setFlowOptions');

      const ribbon = wrapper.get('.word-ribbon-container').element as HTMLElement;
      ribbon.tabIndex = -1; ribbon.focus();
      expect(key({ code: 'F6' }, ribbon).defaultPrevented).toBe(true);
      expect(wrapper.get('.word-statusbar').element.contains(document.activeElement)).toBe(true);
      expect(stub.superdoc!.ops(), 'F6 מדלג על העורך הנעול').not.toContain('focus');
      (document.activeElement as HTMLElement).blur();
      expect(key({ code: 'F6' }).defaultPrevented).toBe(true);
      expect(wrapper.get('.word-statusbar').element.contains(document.activeElement), 'F6 skips the inert titlebar and Ribbon from body too').toBe(true);

      key({ code: 'Slash', ctrlKey: true });
      await settle(8);
      expect(document.querySelector('[aria-modal="true"]')).not.toBeNull();
      expect(key({ code: 'Escape' }).defaultPrevented).toBe(true);
      await settle(8);
      expect(document.querySelector('[aria-modal="true"]')).toBeNull();
      expect(previous.inert).toBe(true);

      await wrapper.get('.status-load__skip').trigger('click');
      expect(previous.inert).toBe(false);
      key({ code: 'KeyB', ctrlKey: true });
      key({ code: 'Enter', ctrlKey: true });
      direction();
      await settle(12);
      expect(adapter.calls.map(call => call.id)).toEqual(['bold', 'direction-ltr']);
      expect(stub.superdoc!.ops()).toContain('format.paragraph.setFlowOptions');
    } finally { release(); }
  });

  it('פירוק משחרר מיד ומונע תוצאה מאוחרת מלפתוח מנוע', async () => {
    const { previous } = containers();
    const wrapper = await mountShell();
    const release = holdPreflight();
    try {
      const signal = await beginBlank();
      wrapper.unmount();
      expect(signal.aborted).toBe(true);
      expect(previous.inert).toBe(false);
      release();
      await settle(12);
      expect(stub.openSources).toHaveLength(1);
    } finally { release(); }
  });

  it('Tell Me שכבר פתוח אינו עוקף את הנעילה, ושדה החיפוש נשמר', async () => {
    containers();
    const wrapper = await mountShell();
    const release = holdPreflight();
    const adapter = stub.adapter as CommandDouble;
    const title = wrapper.getComponent(TitleBar);
    try {
      await beginBlank();
      await wrapper.get('.tell-me-input').setValue('בדיקת חיפוש');
      adapter.reset(); stub.superdoc!.reset();
      title.vm.$emit('run-command', 'bold');
      title.vm.$emit('run-action', 'page-break');
      title.vm.$emit('custom-action', 'about');
      await settle(8);
      expect(adapter.calls).toEqual([]);
      expect(stub.superdoc!.ops()).not.toContain('format.paragraph.setFlowOptions');
      expect(document.querySelector('[aria-modal="true"]')).toBeNull();

      await wrapper.get('.status-load__skip').trigger('click');
      expect((wrapper.get('.tell-me-input').element as HTMLInputElement).value).toBe('בדיקת חיפוש');
      title.vm.$emit('run-command', 'bold');
      title.vm.$emit('run-action', 'page-break');
      title.vm.$emit('custom-action', 'about');
      await settle(12);
      expect(adapter.calls.map(call => call.id)).toContain('bold');
      expect(stub.superdoc!.ops()).toContain('format.paragraph.setFlowOptions');
      expect(document.querySelector('[aria-modal="true"]')).not.toBeNull();
    } finally { release(); }
  });

  it('finally מאוחר של פתיחה שבוטלה אינו משחרר פתיחה חדשה על אותו container', async () => {
    const { previous } = containers();
    const wrapper = await mountShell();
    const releaseA = holdPreflight();
    let releaseB = (): void => {};
    try {
      await beginBlank();
      await wrapper.get('.status-load__skip').trigger('click');
      expect(previous.inert).toBe(false);
      releaseB = holdPreflight();
      const signalB = await beginBlank();
      expect(previous.inert).toBe(true);
      releaseA();
      await settle(12);
      expect(previous.inert).toBe(true);
      expect(signalB.aborted).toBe(false);
      expect(wrapper.find('.status-load__skip').exists()).toBe(true);
      releaseB();
      await vi.waitFor(() => expect(stub.resetCalls).toBe(2));
      await settle(12);
      expect(previous.inert).toBe(false);
      expect(stub.openSources).toHaveLength(2);
    } finally { releaseA(); releaseB(); }
  });
});

/**
 * הקינון שגלון ההדפסה נשען עליו.
 *
 * ## הרגרסיה שהבדיקה הזאת הייתה תופסת
 *
 * `styles/print.css` מסתיר את המעטפת **לפי מקום בעץ ולא לפי שם**: בכל רמה
 * בשרשרת אל המסמך שורד רק הילד שמוביל למסמך (`.editor-area > :not(.editor-stack)`
 * וכן הלאה). כלומר כל כלל שם הוא טענה על עץ ה-DOM שהקובץ הזה מרכיב.
 *
 * הכלל נכתב כש-`.editor-stack` היה ילד ישיר של `.word-app-shell`. הוספת
 * `.editor-area` ביניהם — עטיפה חדשה, שינוי תמים ב-App.vue — הפכה את הכלל
 * למי שמסתיר את **המסמך**, ו„הדפסה” הוציאה גיליון ריק. אותה זריקה בדיוק
 * נמדדה שוב מאוחר יותר: עטיפת `.editor-stack` בעוד `<div>` הפילה את הפלט
 * לגיליון אחד בלי שום זרם תוכן.
 *
 * ## ולמה בדיקת טקסט על ה-CSS אינה יכולה
 *
 * `tests/unit/print.test.ts` מאמת שהסלקטורים קיימים בקובץ, ו-`css-hygiene`
 * מאמת שהמחלקות שהם נוקבים בהן קיימות ב-`src`. שתיהן נשארות ירוקות אחרי
 * הזריקה: הסלקטור לא השתנה, המחלקה לא נעלמה — רק **מי מכיל את מי** השתנה,
 * וזה קיים אך ורק ב-DOM. שם היא נמדדת, כאן.
 *
 * הבדיקה גם אינה עוצרת ב„ההורה הנכון”: היא מאמתת שכל **שאר** הילדים באותה
 * רמה הם מה שהכלל מצפה לו. ילד חדש שיתווסף ל-`.editor-stack` או ל-
 * `.document-pane` לא ידפיס את עצמו ולא ייעלם בשקט — מישהו יצטרך להחליט.
 *
 * הרמה שאין לה DOM כאן היא `#app` (המעטפת מורכבת ישר ל-body), והיא מקובעת
 * מול index.html ו-main.ts ב-`tests/unit/print.test.ts`.
 */
describe('הקינון שגלון ההדפסה נשען עליו', () => {
  it('שרשרת ההורות אל המסמך היא בדיוק זו שכללי print.css נוקבים בה', async () => {
    await mountShell();

    const shell = document.querySelector('.word-app-shell');
    const area = document.querySelector('.editor-area');
    const stack = document.querySelector('.editor-stack');
    const pane = document.querySelector('.document-pane');
    const host = document.querySelector('.editor-stack__host');
    for (const [name, element] of [
      ['.word-app-shell', shell],
      ['.editor-area', area],
      ['.editor-stack', stack],
      ['.document-pane', pane],
      ['.editor-stack__host', host],
    ] as const) {
      expect(element, name).not.toBeNull();
    }

    // כל חוליה בנפרד, כי כל אחת מהן היא כלל אחר בגלון: עטיפה שתיכנס
    // באמצע תשבור בדיוק את זו שמעליה.
    expect(area!.parentElement, '.editor-area הוא ילד ישיר של המעטפת').toBe(shell);
    expect(stack!.parentElement, '.editor-stack הוא ילד ישיר של .editor-area').toBe(area);
    expect(pane!.parentElement, '.document-pane הוא ילד ישיר של .editor-stack').toBe(stack);
    expect(host!.parentElement, 'מיכל המנוע יושב בתוך פאנל הטאב').toBe(pane);

    // ומה ש-`:not()` מסתיר בשתי הרמות הפנימיות הוא באמת רק מה שהוא מתיים.
    for (const child of stack!.children) {
      expect(child.className, 'כל ילד של .editor-stack הוא פאנל טאב').toContain('document-pane');
    }
    for (const child of pane!.children) {
      expect(child.className, 'כל ילד של .document-pane הוא מיכל מנוע').toContain(
        'editor-stack__host',
      );
    }
  });
});

describe('מתג השמירה האוטומטית', () => {
  it('לחיצה מגיעה לקואורדינטור — לא רק לצבע של הפיל', async () => {
    // זו המוטציה שחמקה: הסרת `save?.setAutosaveEnabled(...)` השאירה מתג שמזיז
    // את הכפתור ואינו מכבה שום דבר, ו-203 בדיקות עברו.
    const wrapper = await mountShell();

    // העלייה טוענת את ההעדפה השמורה ומעבירה אותה לקואורדינטור.
    expect(stub.autosaveCalls).toEqual([true]);

    const toggle = wrapper.find('.autosave-toggle');
    expect(toggle.attributes('aria-checked')).toBe('true');

    await toggle.trigger('click');
    await settle();

    expect(stub.autosaveCalls, 'הכיבוי הגיע לקואורדינטור').toEqual([true, false]);
    expect(wrapper.find('.autosave-toggle').attributes('aria-checked')).toBe('false');
  });

  it('הבחירה נשמרת להפעלה הבאה', async () => {
    const wrapper = await mountShell();

    await wrapper.find('.autosave-toggle').trigger('click');
    await settle();

    expect(stub.persistedAutosave).toEqual([false]);
  });

  it('ההעדפה השמורה היא זו שנטענת, ולא ברירת המחדל', async () => {
    // כיבוי בהפעלה קודמת חייב להגיע לקואורדינטור **לפני** העריכה הראשונה,
    // אחרת סבב ה-autosave הראשון רץ לפי ברירת המחדל.
    stub.storedAutosave = false;

    const wrapper = await mountShell();

    expect(stub.autosaveCalls).toEqual([false]);
    expect(wrapper.find('.autosave-toggle').attributes('aria-checked')).toBe('false');
  });
});

describe('שמירה', () => {
  it('כפתור השמירה בסרגל המהיר מריץ שמירה על המסמך הפתוח', async () => {
    const wrapper = await mountShell();

    await wrapper.findAll('.qa-btn')[0]!.trigger('click');
    await settle();

    expect(stub.saveNowCalls).toHaveLength(1);
    expect(stub.saveNowCalls[0]).toMatchObject({ forceSaveAs: false });
  });

  it('Ctrl+S שומר, ו-Ctrl+Shift+S הוא „שמור בשם”', async () => {
    // המטפל יושב על `window`, ולכן זה מה שמעיד שהוא נרשם בפועל.
    await mountShell();

    window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', code: 'KeyS', ctrlKey: true }));
    await settle();
    window.dispatchEvent(
      new KeyboardEvent('keydown', { key: 's', code: 'KeyS', ctrlKey: true, shiftKey: true }),
    );
    await settle();

    expect(stub.saveNowCalls.map((call) => call?.forceSaveAs)).toEqual([false, true]);
  });

  it('Ctrl+S שומר גם בפריסת מקלדת עברית', async () => {
    // הרגרסיה שהתיקון בא לה: בפריסה עברית הדפדפן מדווח `key: 'ד'`, וההשוואה
    // הישנה (`event.key === 's'`) פשוט לא תפסה. בעורך לכתיבת חידושי תורה זה
    // אומר שהשמירה מתה בדיוק כשהמשתמש עשה את מה שהתוסף נועד לו.
    await mountShell();

    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'ד', code: 'KeyS', ctrlKey: true }));
    await settle();

    expect(stub.saveNowCalls).toHaveLength(1);
  });

  it('Ctrl+P מדפיס גם בפריסה עברית, ו-Ctrl+G אינו נבלע', async () => {
    await mountShell();

    // `cancelable` נדרש כדי ש-`defaultPrevented` יהיה מדיד. keydown אמיתי
    // בדפדפן הוא cancelable; אירוע מלאכותי בלי הדגל אינו, ו-preventDefault בו
    // הוא no-op שקט.
    const options = { ctrlKey: true, cancelable: true };
    const print = new KeyboardEvent('keydown', { key: 'פ', code: 'KeyP', ...options });
    const unknown = new KeyboardEvent('keydown', { key: 'ג', code: 'KeyG', ...options });
    window.dispatchEvent(print);
    window.dispatchEvent(unknown);
    await settle();

    expect(print.defaultPrevented).toBe(true);
    // צירוף שאינו שלנו נשאר של הדפדפן.
    expect(unknown.defaultPrevented).toBe(false);
  });

  it('שינוי שם המסמך מסמן אותו כלא-שמור', async () => {
    const wrapper = await mountShell();

    const input = wrapper.find('.doc-title-input');
    (input.element as HTMLInputElement).value = 'חידושי בבא קמא';
    await input.trigger('change');
    await settle();

    expect(stub.markDirtyCalls).toBe(1);
    expect((wrapper.find('.doc-title-input').element as HTMLInputElement).value).toBe(
      'חידושי בבא קמא',
    );
  });

  it('מצב השמירה שהקואורדינטור מדווח מגיע לפס הכותרת', async () => {
    // החיווט הזה (`:is-dirty`, `:save-state-text`) היה מאומת רק ב-regex, וכפיל
    // שמדווח „מלוכלך” ופס שאינו משתנה נראים בסריקת מקור זהים.
    const wrapper = await mountShell();
    expect(wrapper.find('.dirty-indicator').exists()).toBe(false);

    const dirty: SaveSnapshot = {
      state: 'idle',
      isDirty: true,
      isSaving: false,
      targetToken: null,
      name: null,
      lastError: null,
    };
    stub.saveDeps!.onStateChange!(dirty);
    await settle();

    expect(wrapper.find('.dirty-indicator').exists()).toBe(true);
    expect(wrapper.find('.save-state-pill').text()).toBe('שינויים לא שמורים');
  });
});

describe('חיפוש', () => {
  it('פתיחת חיפוש דרך Tell Me פותחת session במנוע ולא רק דיאלוג', async () => {
    // פתיחת הדיאלוג בלי `searchAdapter.open()` היא דיאלוג שכל חיפוש בו נכשל
    // סגור — ואת זה רואים רק ממעטפת מורכבת.
    const wrapper = await mountShell();

    const tellMeInput = wrapper.find('.tell-me-input');
    await tellMeInput.trigger('focus');
    await tellMeInput.setValue('בדיקה');
    await settle();

    await wrapper.find('#tell-me-item-doc-search').trigger('click');
    await settle();

    expect(stub.searchOpens).toBe(1);
  });

  it('קיצור Alt+Q מעביר מיקוד לתיבת Tell Me ופותח את התפריט', async () => {
    const wrapper = await mountShell();

    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'q', code: 'KeyQ', altKey: true }));
    await settle();

    const tellMeInput = wrapper.find('.tell-me-input');
    expect(tellMeInput.attributes('aria-expanded')).toBe('true');
  });
});

/**
 * ההעדפה של הסרגל, ולמה היא צריכה בדיקה משלה.
 *
 * מצב הסרגל שייך למנוע (`config.rulers`), ומופע מנוע חדש נולד כבוי — כלומר
 * בכל פתיחת מסמך המעטפת רואה `false` בדיוק ברגע שבו היא אמורה להחיל `true`
 * שנשמר בהפעלה הקודמת. הבאג שהיה כאן: הסנכרון ההתחלתי כתב את מה שראה אל תוך
 * ההעדפה, וכך **מחק** אותה לפני שהספיקה לחול — ואז לא היה מה להחיל.
 *
 * הבדיקות מקבעות את שני הכיוונים, מפני שתיקון של אחד מהם לבדו נראה נכון:
 * שההעדפה חלה, ושסנכרון לבדו אינו כותב אותה.
 */
describe('ההעדפה של סרגל המידות', () => {
  it('סרגל שנשמר דלוק מתבקש מהמנוע בפתיחה הבאה', async () => {
    stub.storedRuler = true;
    await mountShell();

    const adapter = stub.adapter as CommandDouble;
    expect(
      adapter.calls.map((call) => call.id),
      'המעטפת ביקשה מהמנוע להדליק את הסרגל',
    ).toContain('ruler');
  });

  it('הסנכרון עם מנוע שנולד כבוי אינו מוחק את ההעדפה', async () => {
    stub.storedRuler = true;
    await mountShell();

    // זה הלב: `getState('ruler').active` הוא `false` בפתיחה, ואסור שהוא
    // ייכתב לאחסון — אחרת ההפעלה הבאה כבר לא תדע שהמשתמש רצה סרגל.
    expect(stub.persistedRuler, 'ההעדפה לא נדרסה בסנכרון').not.toContain(false);
  });

  it('הסרגל מופיע כשהמנוע מאשר שהדגל התחלף', async () => {
    stub.storedRuler = true;
    const wrapper = await mountShell();

    // המנוע עונה על `run('ruler')` דרך אותו מסלול שהכפתור ברצועה עובר בו.
    (stub.adapter as CommandDouble).setState('ruler', { active: true });
    await settle();

    const ruler = wrapper.find('.doc-ruler').element as HTMLElement;
    expect(ruler.style.display, 'הרצועה מוצגת').not.toBe('none');
  });

  it('כיבוי יזום נשמר להפעלה הבאה', async () => {
    stub.storedRuler = true;
    await mountShell();
    const adapter = stub.adapter as CommandDouble;

    adapter.setState('ruler', { active: true });
    await settle();
    adapter.setState('ruler', { active: false });
    await settle();

    expect(stub.persistedRuler[stub.persistedRuler.length - 1], 'הכיבוי נשמר').toBe(false);
  });
});

/**
 * „חוזרים לתוסף והוא נפתח בדיוק כמו לפני הסגירה”.
 *
 * מה שנמדד כאן הוא החיווט, ולא ההחלטות: ההחלטות עצמן נבדקות ביחידה
 * (`sessions/session-state.ts`, `engine/caret-anchor.ts`), ומה שאף בדיקת
 * יחידה אינה יכולה לתפוס הוא „הכול נכון, ואף אחד לא קרא לזה”.
 */
describe('חזרה למה שהיה', () => {
  const REMEMBERED = { token: 'tok', name: 'חידושים.docx', writable: true };

  /**
   * רשומת הפעלה מלאה (v2 — `documents` + `activeId`), עם מה שהבדיקה רוצה
   * לשנות בה. `view` הוא ברמת הרשומה כולה; שאר המפתחות שייכים לרשומת
   * המסמך היחיד שבאוסף.
   */
  function storedSession(patch: {
    document?: unknown;
    caret?: unknown;
    draft?: unknown;
    view?: unknown;
  } = {}): unknown {
    const { view, ...entryPatch } = patch;
    const entry = {
      id: 'doc-1',
      document: REMEMBERED,
      caret: null,
      draft: null,
      ...entryPatch,
    };
    return {
      version: 2,
      documents: [entry],
      activeId: entry.id,
      view: view ?? { zoom: null, focusMode: false, ribbonTab: null, ribbonCollapsed: false },
    };
  }

  it('המסמך האחרון נפתח מה-URL שאוצריא נתנה עכשיו', async () => {
    // ה-URL של הריצה הקודמת מת — הפורט מתחלף — ולכן מה שנשמר הוא ה-token,
    // ובעלייה הוא נפתר מחדש.
    stub.storedSession = storedSession();
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };

    await mountShell();

    expect(stub.openSources).toEqual(['loopback://fresh']);
  });

  it('קובץ שאינו נגיש יותר נפתח כמסמך חדש, עם הודעה', async () => {
    stub.storedSession = storedSession();
    stub.resolvedFile = null;

    const wrapper = await mountShell();

    expect(stub.openSources, 'מסמך ריק, בלי מקור').toEqual([undefined]);
    expect(wrapper.find('.word-statusbar').text()).toContain('לא נמצא');
  });

  it('קובץ שאינו נגיש — העבודה שלא נשמרה נפתחת כמסמך חדש ולא נמחקת', async () => {
    // אין לה יעד כתיבה בכל מקרה („שמור” יפתח „שמור בשם”), ולכן פתיחתה כמסמך
    // חדש אינה יכולה לדרוס דבר — והיא הדרך היחידה שלא לאבד אותה.
    stub.storedSession = storedSession({
      draft: { path: 'session-draft.docx', savedAt: 1, documentToken: 'tok', sourceSize: 120 },
    });
    stub.resolvedFile = null;
    stub.draftBytes = new Uint8Array([80, 75, 3, 4]);

    const wrapper = await mountShell();

    expect(stub.openSources[0]).toBeInstanceOf(Blob);
    expect(wrapper.find('.word-statusbar').text()).toContain('נפתחו כמסמך חדש');
  });

  it('טיוטה שלא נשמרה נפתחת במקום הקובץ, והמסמך מסומן כלא-שמור', async () => {
    // זה המסלול שבו סגירת אוצריא הייתה מוחקת עבודה: מה שנכתב ולא נשמר.
    stub.storedSession = storedSession({
      draft: { path: 'session-draft.docx', savedAt: 1, documentToken: 'tok', sourceSize: 120 },
    });
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };
    stub.draftBytes = new Uint8Array([80, 75, 3, 4]);

    const wrapper = await mountShell();

    expect(stub.openSources[0], 'הבייטים של הטיוטה, לא ה-URL').toBeInstanceOf(Blob);
    // `markRestored` ולא `markDirty`: השני היה מתזמן autosave, ואז הטיוטה
    // ששוחזרה נכתבת לקובץ שתיים וחצי שניות אחרי העלייה — בכל עלייה.
    expect(stub.markRestoredCalls, 'עבודה שאינה בדיסק חייבת להיראות כך').toBeGreaterThan(0);
    expect(stub.markDirtyCalls, 'שחזור אינו עריכה').toBe(0);
    expect(wrapper.find('.word-statusbar').text()).toContain('שוחזרו שינויים');
  });

  it('טיוטה של מסמך אחר אינה נפתחת מעל המסמך הזה', async () => {
    // התרחיש היחיד שבו התכונה יכולה למחוק עבודה: תוכן של מסמך אחד שנפתח מעל
    // מסמך אחר, ואז נשמר לקובץ שלו.
    stub.storedSession = storedSession({
      draft: { path: 'session-draft.docx', savedAt: 1, documentToken: 'other', sourceSize: 1 },
    });
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };
    stub.draftBytes = new Uint8Array([80, 75, 3, 4]);

    await mountShell();

    expect(stub.openSources).toEqual(['loopback://fresh']);
  });

  it('מסמך חדש שמעולם לא נשמר חוזר מהטיוטה', async () => {
    // אין קובץ, אין token, ואין לאן לשמור — הטיוטה היא הדבר היחיד שמחזיק
    // את מה שנכתב.
    stub.storedSession = storedSession({
      document: null,
      draft: { path: 'session-draft.docx', savedAt: 1, documentToken: null, sourceSize: null },
    });
    stub.draftBytes = new Uint8Array([80, 75, 3, 4]);

    await mountShell();

    expect(stub.openSources[0]).toBeInstanceOf(Blob);
  });

  it('מצב המיקוד והלשונית ברצועה חוזרים', async () => {
    stub.storedSession = storedSession({
      view: { zoom: null, focusMode: true, ribbonTab: 'references', ribbonCollapsed: false },
    });

    const wrapper = await mountShell();

    expect(wrapper.find('.word-app-shell').classes()).toContain('focus-mode');
    const active = wrapper.findAll('.word-tab-btn').filter((tab) => tab.classes('active'));
    expect(active).toHaveLength(1);
    expect(active[0].text()).toBe('הפניות');
  });

  it('לשונית שאינה מוכרת נופלת ל„בית” ואינה משאירה רצועה ריקה', async () => {
    stub.storedSession = storedSession({
      view: { zoom: null, focusMode: false, ribbonTab: 'לשונית שנמחקה', ribbonCollapsed: false },
    });

    const wrapper = await mountShell();

    const active = wrapper.findAll('.word-tab-btn').filter((tab) => tab.classes('active'));
    expect(active[0].text()).toBe('בית');
  });

  it('הזום והסמן חוזרים למסמך שנפתח', async () => {
    stub.storedSession = storedSession({
      view: { zoom: 150, focusMode: false, ribbonTab: null, ribbonCollapsed: false },
      caret: { start: { blockId: 'b9', ordinal: 8, offset: 4 }, end: null },
    });
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };

    await mountShell();

    const zoomCall = (stub.adapter as { calls: Array<{ id: string; payload?: unknown }> }).calls.find(
      (call) => call.id === 'zoom',
    );
    // `zoomPayload` הוא אחוז ולא אובייקט — ראו engine/payloads.ts.
    expect(zoomCall?.payload).toBe(150);
    expect(stub.caretApplied).toEqual([
      {
        kind: 'selection',
        start: { kind: 'text', blockId: 'b9', offset: 4 },
        end: { kind: 'text', blockId: 'b9', offset: 4 },
      },
    ]);
  });

  it('זום וסמן של מסמך אחר אינם מוחלים על מה שנפתח', async () => {
    // ה-token לא נפתר, נפתח מסמך חדש — וקפיצה לאמצע מסמך אחר עליו היא
    // בדיוק מה שאסור. סמן **הפתיחה** (תחילת המסמך, applyDocumentStartCaret)
    // כן מוצב — הוא של המסמך שנפתח, לא של האחר.
    stub.storedSession = storedSession({
      view: { zoom: 150, focusMode: false, ribbonTab: null, ribbonCollapsed: false },
      caret: { start: { blockId: 'b9', ordinal: 8, offset: 4 }, end: null },
    });
    stub.resolvedFile = null;

    await mountShell();

    expect(stub.caretApplied).toEqual([
      {
        kind: 'selection',
        start: { kind: 'text', blockId: 'block-1', offset: 0 },
        end: { kind: 'text', blockId: 'block-1', offset: 0 },
      },
    ]);
  });

  it('הרשומה נכתבת בפועל, ולא רק „הייתה אמורה להיכתב”', async () => {
    stub.storedSession = storedSession();
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };

    await mountShell();

    expect(stub.persistedSessions.length, 'שום דבר לא נשמר להפעלה הבאה').toBeGreaterThan(0);
    const record = stub.persistedSessions[stub.persistedSessions.length - 1] as {
      documents?: Array<{ document?: { token?: string } }>;
    };
    expect(record.documents?.[0]?.document?.token).toBe('tok');
  });

  it('פתיחה שנכשלה אינה מוחקת את הטיוטה ואינה שוכחת את המסמך', async () => {
    // כשל בפתיחה עשוי להיות זמני — worker שלא עלה, קובץ נעול. מחיקת הטיוטה
    // או רישום המסמך הריק היו הופכים אותו לאובדן קבוע.
    stub.storedSession = storedSession({
      draft: { path: 'session-draft.docx', savedAt: 1, documentToken: 'tok', sourceSize: 120 },
    });
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };
    stub.draftBytes = new Uint8Array([80, 75, 3, 4]);
    stub.openFailures = 1;

    await mountShell();

    expect(stub.draftRemovals, 'הטיוטה היא הדבר היחיד שמחזיק את העבודה').toBe(0);

    // המסמך הריק שנפתח כגיבוי אינו נרשם. כל רשומה שנכתבה בכל זאת חייבת
    // עדיין לנקוב במסמך האחרון — אחרת ההפעלה הבאה לא תדע במה לנסות שוב,
    // וגם לא לְמי הטיוטה שייכת.
    const forgotten = (
      stub.persistedSessions as Array<{ documents?: Array<{ document?: unknown }> } | null>
    ).filter((record) => (record?.documents?.[0]?.document ?? null) == null);
    expect(forgotten, 'רשומה ששכחה את המסמך האחרון').toEqual([]);
  });

  it('שמירה אוטומטית מעבירה את הטיוטה לקודמת — לא רק שמירה ידנית', async () => {
    // הרגרסיה המקורית: הזוכר היה נתלה על `onSave` של המעטפת, ואילו ה-autosave
    // יורה מתוך הקואורדינטור ואינו עובר שם. התוצאה הייתה טיוטה שנשארת חיה
    // ומפסיקה להתעדכן — ואז נפתחת בהפעלה הבאה מעל עבודה חדשה ממנה.
    //
    // מה שהשתנה: היעד אינו מחיקה אלא העברה ל-`previousDraft`, כדי שיישאר
    // עותק של „לפני” גם אחרי שמירה אוטומטית. שני חצאי החוזה נבדקים כאן —
    // ש-`draft` התרוקן, ושהקובץ **לא** נמחק.
    stub.storedSession = storedSession({
      draft: { path: 'session-draft.docx', savedAt: 1, documentToken: 'tok', sourceSize: 120 },
    });
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };
    stub.draftBytes = new Uint8Array([80, 75, 3, 4]);

    await mountShell();
    expect(stub.draftRemovals, 'עוד לא נשמר דבר').toBe(0);

    await stub.saveDeps?.exportDocument();
    stub.saveDeps?.onSaved?.({ token: 'tok', name: 'חידושים.docx', size: 250 });
    await settle(6);

    const record = stub.persistedSessions[stub.persistedSessions.length - 1] as {
      documents?: Array<{ draft?: unknown; previousDraft?: { path?: string } | null }>;
    } | null;
    expect(record?.documents?.[0]?.draft, 'העבודה בדיסק — הטיוטה אינה „לא שמור” יותר').toBeNull();
    expect(record?.documents?.[0]?.previousDraft?.path, 'ויש לאן לחזור').toBe('session-draft.docx');
    expect(stub.draftRemovals, 'ההעברה היא החלפת מצביע, בלי קריאת גשר').toBe(0);
  });

  it('פתיחה רגילה אינה מוחקת שום טיוטה', async () => {
    stub.storedSession = storedSession({
      draft: { path: 'session-draft.docx', savedAt: 1, documentToken: 'other', sourceSize: 1 },
    });
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };
    stub.draftBytes = new Uint8Array([80, 75, 3, 4]);

    await mountShell();

    expect(stub.draftRemovals).toBe(0);
  });

  it('משתמש שמעדכן מגרסה קודמת אינו מאבד את המסמך שעבד עליו', async () => {
    // אין רשומת הפעלה, ויש רק את המפתח הישן.
    stub.storedSession = null;
    stub.lastDocument = REMEMBERED;
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };

    await mountShell();

    expect(stub.openSources).toEqual(['loopback://fresh']);
    expect(stub.forgotLastDocument, 'המפתח הישן נמחק כדי שלא יישאר מקור שני').toBe(true);
  });
});

/**
 * הצלחה שיש עליה מה לומר — `CommandOutcome.note`.
 *
 * הפקד ב-Ribbon אינו יודע לומר דבר: הוא מעביר את התוצאה ל-`COMMAND_REPORTER`
 * שהמעטפת מספקת, וכל ההחלטה מה יופיע בפס היא של `reportCommand`. לכן נמדד
 * כאן המדווח עצמו, בדיוק כפי שהפקד קורא לו.
 */
/**
 * שורת המצב מציגה הודעה אחת, ובפתיחה יש שלוש שיכולות להיאמר. הסדר נקבע
 * ב-App.vue ומנומק שם; כאן הוא נמדד, כדי ששינוי תמים ב-`else if` לא ידחוק
 * בשקט את ההודעה היחידה שאומרת למשתמש שהמסמך שלו **שונה**.
 */
describe('הודעת השלב המקדים בשורת המצב', () => {
  const statusText = (wrapper: ReturnType<typeof mount>) => wrapper.find('.word-statusbar').text();

  /**
   * מסמך זכור שנפתח בעלייה. „פתוח לקריאה” נגזרת מ-`writable` שברשומה
   * (ראו resolveRememberedFile), ולא ממה ש-`resolveFileUrl` מחזיר.
   */
  function openFile(access: 'read' | 'readwrite'): void {
    stub.storedSession = {
      version: 2,
      documents: [
        {
          id: 'doc-1',
          document: { token: 'tok', name: 'חידושים.docx', writable: access === 'readwrite' },
          caret: null,
          draft: null,
        },
      ],
      activeId: 'doc-1',
      view: { zoom: null, focusMode: false, ribbonTab: null, ribbonCollapsed: false },
    };
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };
  }

  it('מסמך שההדגשה בו הושלמה — ההודעה מגיעה לשורת המצב, ולא כשגיאה', async () => {
    openFile('readwrite');
    stub.preflight.notice = COMPLEX_SCRIPT_BOLD_NOTICE;

    const wrapper = await mountShell();

    expect(statusText(wrapper)).toContain(COMPLEX_SCRIPT_BOLD_NOTICE);
    expect(wrapper.find('.status-message').classes()).not.toContain('error');
  });

  it('קודמת ל„פתוח לקריאה” — את זה המשתמש יגלה בשמירה, ואת השינוי במסמך לא', async () => {
    openFile('read');
    stub.preflight.notice = COMPLEX_SCRIPT_BOLD_NOTICE;

    const wrapper = await mountShell();

    expect(statusText(wrapper)).toContain(COMPLEX_SCRIPT_BOLD_NOTICE);
    expect(statusText(wrapper)).not.toContain('פתוח לקריאה');
  });

  it('נדחית מפני מאקרו שWord מריץ בפתיחה', async () => {
    openFile('readwrite');
    stub.preflight.notice = COMPLEX_SCRIPT_BOLD_NOTICE;
    stub.preflight.vba = {
      ...NO_VBA,
      hasMacroPart: true,
      autoRun: ['Module1.AutoOpen'],
      status: 'יש במסמך מאקרו שWord מריץ בפתיחה',
    };

    const wrapper = await mountShell();

    expect(statusText(wrapper)).toContain('מאקרו שWord מריץ בפתיחה');
    expect(statusText(wrapper)).not.toContain(COMPLEX_SCRIPT_BOLD_NOTICE);
  });

  it('בלי תיקון — „פתוח לקריאה” נאמרת כרגיל, ומאקרו שאינו רץ בפתיחה נדחה מפניה', async () => {
    openFile('read');
    stub.preflight.vba = { ...NO_VBA, hasMacroPart: true, status: 'במסמך יש מאקרו' };

    const wrapper = await mountShell();

    expect(statusText(wrapper)).toContain('פתוח לקריאה');
    expect(statusText(wrapper)).not.toContain('במסמך יש מאקרו');
  });
});

describe('הודעת-מידע על פקודה שהצליחה', () => {
  /**
   * דוגמה, ולא ציטוט של פקודה קיימת. ההודעה שמילאה את התפקיד הזה — „העמודה
   * הראשונה מצוירת בצד שמאל…” — ירדה עם המעבר ל-superdoc 2.12.0, שמצייר את
   * הטור הראשון בימין. מה שנבדק כאן הוא **הערוץ** ולא ההודעה: `note` של
   * `CommandOutcome` (engine/command-adapter.ts) הוא הדרך שבה הצלחה שיש עליה
   * מה לומר מגיעה לפס המצב, והתנהגות המעטפת סביבו — הצגה, ניקוי, ואי-דריסה
   * של ערוץ אחר — היא חוזה שאסור שיישחק עד שיגיע המפיק הבא.
   */
  const NOTE = 'הפעולה בוצעה, ויש עליה מה לומר.';

  function reporterOf(wrapper: Awaited<ReturnType<typeof mountShell>>): CommandReporter {
    const provides = (wrapper.vm.$ as unknown as { provides: Record<symbol, unknown> }).provides;
    const report = provides[COMMAND_REPORTER as unknown as symbol];
    if (typeof report !== 'function') throw new Error('המעטפת לא סיפקה מדווח פקודות');
    return report as CommandReporter;
  }

  const statusOf = (wrapper: Awaited<ReturnType<typeof mountShell>>): string =>
    wrapper.find('.status-message').exists() ? wrapper.find('.status-message').text() : '';

  it('ההודעה מגיעה לפס המצב', async () => {
    const wrapper = await mountShell();

    reporterOf(wrapper)({ ok: true, note: NOTE }, 'page-columns');
    await settle();

    expect(statusOf(wrapper)).toContain(NOTE);
  });

  it('הצלחה בלי הודעה מנקה הודעה שהפכה לשקר', async () => {
    // פקודה שיש עליה מה לומר, ואחריה אותה פקודה בלי הודעה. בלי הניקוי הפס
    // היה ממשיך לתאר מצב שהפקודה השנייה כבר ביטלה.
    const wrapper = await mountShell();
    const report = reporterOf(wrapper);

    report({ ok: true, note: NOTE }, 'page-columns');
    await settle();
    report({ ok: true }, 'page-columns');
    await settle();

    expect(statusOf(wrapper)).toBe('');
  });

  it('אינה מוחקת הודעה של ערוץ אחר שנכתבה בינתיים', async () => {
    // הפס יש לו כותב שני: `STATUS_NOTIFIER`, שכלי הלשוניות מודיעים דרכו
    // („בוצעו 3 תיקונים”). הניקוי מותנה בכך שההודעה שעל המסך היא **זו**
    // שהוצגה, ובלי התנאי הזה כל פקודה מוצלחת הייתה מוחקת הודעה של הערוץ ההוא.
    const wrapper = await mountShell();
    const provides = (wrapper.vm.$ as unknown as { provides: Record<symbol, unknown> }).provides;
    const notify = provides[STATUS_NOTIFIER as unknown as symbol] as (text: string) => void;
    const report = reporterOf(wrapper);

    report({ ok: true, note: NOTE }, 'page-columns');
    await settle();
    notify('בוצעו 3 תיקונים');
    await settle();
    report({ ok: true }, 'page-columns');
    await settle();

    expect(statusOf(wrapper)).toContain('בוצעו 3 תיקונים');
  });

  it('כשל דוחה את ההודעה ומסמן שגיאה', async () => {
    const wrapper = await mountShell();
    const report = reporterOf(wrapper);

    report({ ok: true, note: NOTE }, 'page-columns');
    await settle();
    report({ ok: false, message: 'שינוי מספר העמודות ל-2 נכשל', reason: 'threw' }, 'page-columns');
    await settle();

    expect(statusOf(wrapper)).toContain('נכשל');
    expect(wrapper.find('.status-message').classes()).toContain('error');
  });
});

/**
 * שחזור **כל** הטאבים, ולא רק זה שהיה פעיל.
 *
 * ## הבאג שהבדיקות כאן שומרות עליו
 *
 * הרשומה תמיד שמרה רשומה לכל טאב פתוח, אבל בעלייה נקראה ממנה רשומה אחת —
 * הפעילה — ושאר הטאבים נעלמו; גרוע מזה, הטיוטות שלהם (עבודה שלא נשמרה
 * לקובץ!) נמחקו מהמרחב הפרטי כדי שלא יישארו יתומות. כלומר סגירת התוסף עם
 * שלושה מסמכים פתוחים החזירה מסמך אחד, ומחקה את מה שלא נשמר בשניים.
 *
 * ## מה נמדד, ולמה דווקא כך
 *
 * שלוש שאלות, וכל אחת מהן היא באג אחר אילו נשברה:
 *
 * 1. **כל הטאבים חוזרים** — ברצועה, עם השם שלהם, לפני שנגעו בהם.
 * 2. **רק הפעיל נטען** — זו הטעינה העצלה; `openSources` הוא מה שמבחין בין
 *    „הטאב קיים” ל„נפתח לתוכו מנוע”, ובלעדיו הבדיקה הייתה מאשרת בירוק גם
 *    פתיחה של חמישה מסמכים בעלייה.
 * 3. **מה שממתין אינו נמחק** — לא הטיוטה ולא הרשומה, וסגירתו שואלת קודם.
 */
describe('שחזור כל הטאבים', () => {
  const FIRST = { token: 'tok-1', name: 'ראשון.docx', writable: true };
  const SECOND = { token: 'tok-2', name: 'שני.docx', writable: true };
  const SECOND_DRAFT = {
    path: 'session-draft-doc-2.docx',
    savedAt: 1,
    documentToken: 'tok-2',
    sourceSize: 200,
  };

  /** רשומה של שני טאבים פתוחים, כשהראשון הוא הפעיל. */
  function twoTabs(secondDraft: unknown = null): unknown {
    return {
      version: 2,
      documents: [
        { id: 'doc-1', document: FIRST, caret: null, draft: null },
        { id: 'doc-2', document: SECOND, caret: null, draft: secondDraft },
      ],
      activeId: 'doc-1',
      view: { zoom: null, focusMode: false, ribbonTab: null, ribbonCollapsed: false },
    };
  }

  const RESOLVED: Record<string, unknown> = {
    'tok-1': { token: 'tok-1', url: 'loopback://first', name: 'ראשון.docx', size: 100 },
    'tok-2': { token: 'tok-2', url: 'loopback://second', name: 'שני.docx', size: 200 },
  };

  beforeEach(() => {
    // כל token נפתר לקובץ **שלו**: בלי זה „הטאב השני נפתח” היה עובר בירוק
    // גם אם נפתח בו המסמך של הראשון.
    stub.resolvedFile = (token: string) => RESOLVED[token] ?? null;
  });

  /** הרשומה האחרונה שנכתבה. `.at(-1)` אינו ב-lib של הפרויקט. */
  function lastPersistedSession(): unknown {
    return stub.persistedSessions[stub.persistedSessions.length - 1];
  }

  function tabTitles(wrapper: Awaited<ReturnType<typeof mountShell>>): string[] {
    return wrapper.findAll('.word-doctab-title').map((el) => el.text());
  }

  it('כל הטאבים שהיו פתוחים חוזרים לרצועה, עם השם שלהם', async () => {
    stub.storedSession = twoTabs();

    const wrapper = await mountShell();

    expect(tabTitles(wrapper)).toEqual(['ראשון', 'שני']);
    expect(wrapper.find('.word-doctab.active .word-doctab-title').text()).toBe('ראשון');
  });

  it('רק המסמך של הטאב הפעיל נטען בעלייה', async () => {
    // טעינה עצלה: כל מסמך הוא מופע מנוע מלא, וחמישה טאבים פתוחים אינם סיבה
    // להכפיל פי חמישה את הזמן עד שרואים את המסמך שעובדים עליו.
    //
    // הבדיקה הזאת שומרת על הכיוון ההפוך מכל השאר בקובץ הזה, ולכן היא הייתה
    // עוברת גם לפני השינוי (שם שוחזר טאב אחד בלבד): מה שהיא מונעת הוא
    // „תיקון” עתידי שיפתח את כל הטאבים בעלייה.
    stub.storedSession = twoTabs();

    await mountShell();

    expect(stub.openSources).toEqual(['loopback://first']);
  });

  it('מעבר לטאב שממתין פותח את הקובץ שלו — פעם אחת', async () => {
    stub.storedSession = twoTabs();
    const wrapper = await mountShell();

    await wrapper.findAll('.word-doctab')[1]!.trigger('click');
    await settle(12);

    expect(stub.openSources).toEqual(['loopback://first', 'loopback://second']);
    expect(wrapper.find('.word-doctab.active .word-doctab-title').text()).toBe('שני');

    // חזרה אליו אינה פותחת אותו שוב: הוא כבר טעון, בדיוק כמו כל טאב אחר.
    await wrapper.findAll('.word-doctab')[0]!.trigger('click');
    await settle(8);
    await wrapper.findAll('.word-doctab')[1]!.trigger('click');
    await settle(8);
    expect(stub.openSources).toHaveLength(2);
  });

  it('הרשומה שנכתבת מחזיקה את שני הטאבים גם לפני שהשני נטען', async () => {
    // זה מה שהופך את השחזור ליציב על פני כמה הפעלות: טאב שלא נגעו בו חייב
    // להיכתב חזרה, אחרת הוא נעלם בסגירה הבאה.
    stub.storedSession = twoTabs();

    await mountShell();

    const last = lastPersistedSession() as { documents: Array<{ id: string }>; activeId: string };
    expect(last.documents.map((entry) => entry.id)).toEqual(['doc-1', 'doc-2']);
    expect(last.activeId).toBe('doc-1');
  });

  it('הטיוטה של טאב שממתין אינה נמחקת בעלייה', async () => {
    // הבאג שהיה: עבודה שלא נשמרה בטאב שאינו הפעיל נמחקה מהמרחב הפרטי ברגע
    // העלייה, לפני שהמשתמש בכלל ראה שהטאב קיים.
    stub.storedSession = twoTabs(SECOND_DRAFT);
    stub.draftsByPath[SECOND_DRAFT.path] = new Uint8Array([80, 75, 3, 4]);

    const wrapper = await mountShell();

    expect(stub.removedDrafts).toEqual([]);
    expect(
      wrapper.findAll('.word-doctab')[1]!.find('.word-doctab-dirty').exists(),
      'טאב שיש בו עבודה שלא נשמרה מסומן ככזה עוד לפני שנטען',
    ).toBe(true);
  });

  it('מעבר לטאב שממתין פותח את הטיוטה שלו, ולא את הקובץ מהדיסק', async () => {
    stub.storedSession = twoTabs(SECOND_DRAFT);
    stub.draftsByPath[SECOND_DRAFT.path] = new Uint8Array([80, 75, 3, 4]);
    const wrapper = await mountShell();

    await wrapper.findAll('.word-doctab')[1]!.trigger('click');
    await settle(12);

    expect(stub.openSources[1], 'הבייטים של הטיוטה, לא ה-URL').toBeInstanceOf(Blob);
    expect(wrapper.find('.word-statusbar').text()).toContain('שוחזרו שינויים');
  });

  it('סגירת טאב שממתין ויש בו עבודה שלא נשמרה — שואלת, ו„ביטול” משאיר הכול', async () => {
    stub.storedSession = twoTabs(SECOND_DRAFT);
    stub.draftsByPath[SECOND_DRAFT.path] = new Uint8Array([80, 75, 3, 4]);
    const wrapper = await mountShell();

    await wrapper.findAll('.word-doctab')[1]!.find('.word-doctab-close').trigger('click');
    await settle(12);

    // שאלה אחת עם שני כפתורים — אין ממה לייצא בטאב שלא נטען, ולכן „שמור”
    // אינו מוצג כלל.
    expect(unsavedDialog(), 'הדיאלוג נפתח').not.toBeNull();
    expect(document.querySelector('[data-choice="save"]'), '„שמור” אינו מוצע').toBeNull();

    await answerUnsaved('cancel');

    expect(wrapper.findAll('.word-doctab')).toHaveLength(2);
    expect(stub.removedDrafts, 'עבודה שלא נשמרה אינה נמחקת בלי אישור').toEqual([]);
    expect(stub.persistedDiscardBackups, 'ומה שלא נמחק גם אינו מגובה').toEqual([]);
    expect(wrapper.find('.word-statusbar').text()).toContain('סגירת הטאב בוטלה');
  });

  it('„לא לשמור” על טאב שממתין — עותק לשחזור נכתב **לפני** שהטיוטה נמחקת', async () => {
    // זו ההבטחה שמופיעה בדיאלוג עצמו, וזה הסדר שקובע אם היא מתקיימת: מחיקה
    // קודם הייתה משאירה את הגיבוי בלי בייטים לקרוא מהם.
    stub.storedSession = twoTabs(SECOND_DRAFT);
    stub.draftsByPath[SECOND_DRAFT.path] = new Uint8Array([80, 75, 3, 4]);
    const wrapper = await mountShell();

    await wrapper.findAll('.word-doctab')[1]!.find('.word-doctab-close').trigger('click');
    await settle(12);
    await answerUnsaved('discard');

    expect(tabTitles(wrapper), 'הטאב נסגר').toEqual(['ראשון']);
    const backup = stub.workspaceWrites.find((write) => write.path.startsWith('discarded-'));
    expect(backup?.path, 'המשבצת הראשונה, כי הרשומה הייתה ריקה').toBe('discarded-0.docx');
    expect([...(backup?.bytes ?? [])], 'ובדיוק העבודה שלא נשמרה').toEqual([80, 75, 3, 4]);
    expect(stub.removedDrafts, 'ורק אז הטיוטה נמחקה').toContain(SECOND_DRAFT.path);
  });

  it('הרשומה של הגיבוי נושאת את השם והמשבצת, ונכתבת אחרי הכתיבה', async () => {
    stub.storedSession = twoTabs(SECOND_DRAFT);
    stub.draftsByPath[SECOND_DRAFT.path] = new Uint8Array([80, 75, 3, 4]);
    const wrapper = await mountShell();

    await wrapper.findAll('.word-doctab')[1]!.find('.word-doctab-close').trigger('click');
    await settle(12);
    await answerUnsaved('discard');

    const ledger = lastDiscardBackups();
    expect(ledger, 'שורה אחת').toHaveLength(1);
    expect(ledger[0]).toMatchObject({ slot: 0, name: 'שני' });
    expect(ledger[0]!.discardedAt, 'ומתי זה קרה').toBeGreaterThan(0);
  });

  it('המשבצת הבאה נבחרת לפי מה שכבר בגיבוי, ולא מאפס', async () => {
    // רשומה שנשארה מהפעלה קודמת: הכתיבה הבאה אינה דורסת את מה שיש בה כל עוד
    // יש משבצת פנויה.
    stub.storedDiscardBackups = [
      { slot: 0, name: 'ישן', size: 4, discardedAt: 1, token: null },
    ];
    stub.storedSession = twoTabs(SECOND_DRAFT);
    stub.draftsByPath[SECOND_DRAFT.path] = new Uint8Array([80, 75, 3, 4]);
    const wrapper = await mountShell();

    await wrapper.findAll('.word-doctab')[1]!.find('.word-doctab-close').trigger('click');
    await settle(12);
    await answerUnsaved('discard');

    const backup = stub.workspaceWrites.find((write) => write.path.startsWith('discarded-'));
    expect(backup?.path).toBe('discarded-1.docx');
    const ledger = lastDiscardBackups();
    expect(ledger.map((entry) => entry.slot), 'החדש ראשון, והישן נשאר').toEqual([1, 0]);
  });

  it('טעינת גיבויים מאוחרת אינה חוסמת פתיחה ואינה דורסת גיבוי שנוסף בינתיים', async () => {
    let release!: (value: unknown) => void;
    stub.storedDiscardBackups = new Promise(resolve => { release = resolve; });
    stub.storedSession = twoTabs(SECOND_DRAFT);
    stub.draftsByPath[SECOND_DRAFT.path] = new Uint8Array([80, 75, 3, 4]);
    const old = [{ slot: 0, name: 'ישן', size: 4, discardedAt: 1, token: null }];
    const wrapper = await mountShell();
    try {
      expect(stub.openSources).toHaveLength(1);
      await wrapper.findAll('.word-doctab')[1]!.find('.word-doctab-close').trigger('click');
      await settle(12);
      await answerUnsaved('discard');
      expect(stub.workspaceWrites.filter(write => write.path.startsWith('discarded-'))).toHaveLength(0);
      release(old);
      await settle(24);
      expect(lastDiscardBackups().map(item => item.slot)).toEqual([1, 0]);
      pressCtrl('KeyO');
      await settle(8);
      expect(wrapper.find('.open-discarded').text()).toContain('(2)');
    } finally {
      release(old);
      await settle(12);
      wrapper.unmount();
    }
  });

  it('סגירת טאב שממתין ואין בו מה לאבד אינה שואלת, ומורידה אותו מהרשומה', async () => {
    stub.storedSession = twoTabs();
    const wrapper = await mountShell();

    await wrapper.findAll('.word-doctab')[1]!.find('.word-doctab-close').trigger('click');
    await settle(12);

    expect(tabTitles(wrapper)).toEqual(['ראשון']);
    const last = lastPersistedSession() as { documents: Array<{ id: string }> };
    expect(last.documents.map((entry) => entry.id)).toEqual(['doc-1']);
  });

  it('סגירת הטאב הפעיל עוברת לטאב שממתין — וטוענת אותו', async () => {
    stub.storedSession = twoTabs();
    const wrapper = await mountShell();

    await wrapper.findAll('.word-doctab')[0]!.find('.word-doctab-close').trigger('click');
    await settle(14);

    expect(tabTitles(wrapper)).toEqual(['שני']);
    expect(stub.openSources, 'הטאב שנבחר במקומו נטען, ולא נשאר ריק').toEqual([
      'loopback://first',
      'loopback://second',
    ]);
  });

  /**
   * „יציאה” סוגרת — וזה מה שלא היה נכון קודם: הכפתור שאל „לצאת בלי לשמור?
   * השינויים יימחקו”, ואז ניווט לספרייה בלבד; המסמכים המשיכו לחכות פתוחים
   * (ה-WebView מושהה ואינו נהרס), ושום דבר לא נמחק. שאלה שמתארת פעולה שאינה
   * קורית היא שאלה שלומדים להתעלם ממנה, ודווקא זו.
   *
   * שלוש טענות בבדיקה אחת, כי „נסגר” הוא צירוף שלהן: הרצועה, הרשומה שההפעלה
   * הבאה תקרא (הטאב שברקע יורד ממנה גם הוא — לא רק זה שעל המסך), והמעבר
   * לספרייה שקורה רק אחרי הסגירה.
   */
  it('„יציאה” סוגרת את כל הטאבים ואז עוברת לספרייה', async () => {
    stub.storedSession = twoTabs();
    const wrapper = await mountShell();

    // ה-SDK מוזרק לרגע הזה בלבד: בשאר הקובץ `window.Otzaria` אינו קיים
    // בכוונה (ראו ראש הקובץ), וכאן צריך למדוד שהמעבר לספרייה אכן נקרא —
    // סגירה שאינה מוציאה את המשתמש מהמסך אינה „יציאה”.
    const hostCalls: string[] = [];
    window.Otzaria = {
      call: (method: string) => {
        hostCalls.push(method);
        return Promise.resolve({ success: true, data: true });
      },
    } as never;

    try {
      const fileTab = wrapper.findAll('[role="tab"]').find((tab) => tab.text().includes('קובץ'));
      await fileTab!.trigger('click');
      await settle();

      await buttonByTip(wrapper, 'סגירת המסמך').trigger('click');
      await settle(16);

      expect(tabTitles(wrapper), 'נשאר מסמך פתוח אחרי היציאה').toEqual(['מסמך חדש']);
      const last = lastPersistedSession() as { documents: Array<{ id: string }> };
      expect(last.documents.map((entry) => entry.id)).not.toContain('doc-1');
      expect(last.documents.map((entry) => entry.id)).not.toContain('doc-2');
      expect(hostCalls, 'היציאה לא הוציאה מהמסך').toContain('navigation.goTo');
    } finally {
      window.Otzaria = undefined as never;
    }
  });

  it('„יציאה” על טאב שיש בו עבודה שלא נשמרה מגבה אותה לפני שהיא נמחקת', async () => {
    // מסלול שני ונפרד מ„×” על טאב: היציאה שואלת על כל הטאבים ואז סוגרת את
    // כולם (`closeAllSessions`), והמחיקה שם היא של `destroy` ולא של
    // `discardDraft`. גיבוי שהיה מחובר רק למסלול הראשון היה מאבד כאן עבודה
    // בשקט — בדיוק כמו שהיה לפניו.
    stub.storedSession = twoTabs(SECOND_DRAFT);
    stub.draftsByPath[SECOND_DRAFT.path] = new Uint8Array([80, 75, 3, 4]);
    const wrapper = await mountShell();

    const hostCalls: string[] = [];
    window.Otzaria = {
      call: (method: string) => {
        hostCalls.push(method);
        return Promise.resolve({ success: true, data: true });
      },
    } as never;

    try {
      const fileTab = wrapper.findAll('[role="tab"]').find((tab) => tab.text().includes('קובץ'));
      await fileTab!.trigger('click');
      await settle();
      await buttonByTip(wrapper, 'סגירת המסמך').trigger('click');
      await settle(16);

      // הטאב הראשון נקי ואינו נשאל; השאלה היא על השני, שיש בו טיוטה.
      expect(unsavedDialog(), 'היציאה שאלה').not.toBeNull();
      await answerUnsaved('discard');
      // התשובה משחררת שרשרת שלמה: סגירת כל הטאבים, פתיחת טאב ריק במקומם,
      // ורק בסופה המעבר לספרייה.
      await settle(16);

      const backup = stub.workspaceWrites.find((write) => write.path.startsWith('discarded-'));
      expect([...(backup?.bytes ?? [])], 'העבודה שלא נשמרה נשמרה לשחזור').toEqual([80, 75, 3, 4]);
      expect(lastDiscardBackups().map((entry) => entry.name)).toEqual(['שני']);
      expect(hostCalls, 'ורק אחר כך יצאנו').toContain('navigation.goTo');
    } finally {
      window.Otzaria = undefined as never;
    }
  });

  it('אחרי „לא לשמור” הכפתור מופיע, ופתיחה משם מעלה את העבודה בטאב חדש', async () => {
    // המעגל השלם: „לא לשמור” → העותק נכתב → הכפתור ב„פתח מסמך” מופיע →
    // המסמך חוזר. כל חוליה בו שבורה בשקט אם היא לבדה — הכפתור מוסתר כשהמונה
    // 0, והשורה פותחת מסמך ריק אם הבייטים לא הגיעו.
    stub.storedSession = twoTabs(SECOND_DRAFT);
    stub.draftsByPath[SECOND_DRAFT.path] = new Uint8Array([80, 75, 3, 4]);
    const wrapper = await mountShell();

    await wrapper.findAll('.word-doctab')[1]!.find('.word-doctab-close').trigger('click');
    await settle(12);
    await answerUnsaved('discard');

    // הבייטים שנכתבו לגיבוי הם מה שייקרא ממנו — הכפיל מחזיר לפי נתיב.
    const backup = stub.workspaceWrites.find((write) => write.path.startsWith('discarded-'));
    stub.draftsByPath[backup!.path] = backup!.bytes;
    stub.openSources.length = 0;

    // דרך המסך שהמשתמש עובר בו: „פתח קובץ” ואז „נסגרו בלי לשמור”.
    pressCtrl('KeyO');
    await settle(8);
    const entryPoint = document.querySelector<HTMLButtonElement>('.open-discarded');
    expect(entryPoint, 'הכפתור מופיע ברגע שיש מה לשחזר').not.toBeNull();
    expect(entryPoint!.textContent).toContain('(1)');

    entryPoint!.click();
    await settle(8);
    expect(document.querySelector('.discarded-name')?.textContent).toBe('שני');

    document.querySelector<HTMLButtonElement>('.discarded-open')!.click();
    await settle(16);

    expect(stub.openSources[0], 'הבייטים של הגיבוי, לא קובץ מהדיסק').toBeInstanceOf(Blob);
    expect(tabTitles(wrapper), 'ובטאב נוסף, על שם המסמך שממנו בא').toContain('שני');
    expect(wrapper.find('.word-statusbar').text()).toContain('נפתח מהעותק שנשמר בסגירה');
  });

  it('לחיצה כפולה על „פתח” של עותק שחזור פותחת אותו פעם אחת בלבד', async () => {
    // קריאה מהמרחב הפרטי היא אסינכרונית. בלי נעילה, שתי לחיצות לפני סיומה
    // פותחות שתי פעולות `openDocument`, והאחרונה עשויה להחליף את הראשונה.
    stub.storedDiscardBackups = [
      { slot: 0, name: 'עותק', size: 4, discardedAt: Date.now(), token: null },
    ];
    stub.draftsByPath['discarded-0.docx'] = new Uint8Array([80, 75, 3, 4]);
    let releaseRead: (() => void) | undefined;
    stub.workspaceReadGate = new Promise<void>((resolve) => {
      releaseRead = resolve;
    });
    await mountShell();
    stub.openSources.length = 0;

    pressCtrl('KeyO');
    await settle(8);
    document.querySelector<HTMLButtonElement>('.open-discarded')!.click();
    await settle(8);

    const open = document.querySelector<HTMLButtonElement>('.discarded-open')!;
    open.click();
    open.click();
    await settle(4);
    expect(open.disabled, 'הפעולה מסומנת כעסוקה בזמן הקריאה').toBe(true);
    expect(stub.openSources, 'הפתיחה עדיין ממתינה לבייטים').toEqual([]);

    releaseRead?.();
    await settle(16);
    expect(stub.openSources, 'נפתחה פעולה אחת בלבד').toHaveLength(1);
  });

  it('טאב שיש בו מנוע נקי וטיוטה — הגיבוי לוקח את הטיוטה, לא את מה שבמנוע', async () => {
    /*
     * המסלול שאבד בו מידע, והוא היחיד שבו הגיבוי יכול להיכשל בשקט: טאב שיש
     * בו מנוע שאינו „מלוכלך” בעוד הרשומה מחזיקה טיוטה עם העבודה. כך נראה טאב
     * שהטעינה שלו נכשלה ונפל לתוכו מסמך ריק (`remember: false`), וכך נראה גם
     * טאב שנפתח מהטיוטה. `exportDocx` עליו **מצליח** ומחזיר את מה שבמנוע —
     * ולכן גיבוי שמעדיף את המנוע היה כותב את הריקנות ומוחק את העבודה
     * ב-`destroy({ removeDraft: true })` שבא אחריו.
     *
     * הכפיל חייב לייצא בהצלחה, אחרת הבדיקה עוברת מהסיבה הלא נכונה: בלי
     * `export` על המופע, `exportDocx` זורק והמעטפת נופלת לטיוטה ממילא.
     */
    stub.storedSession = twoTabs(SECOND_DRAFT);
    stub.draftsByPath[SECOND_DRAFT.path] = new Uint8Array([80, 75, 3, 4]);
    const wrapper = await mountShell();
    /*
     * `arrayBuffer` מוזרק ידנית, ובלעדיו הבדיקה חסרת ערך.
     *
     * **jsdom אינו מממש `Blob.prototype.arrayBuffer` כלל** (נמדד: אפס אזכורים
     * ב-`node_modules/jsdom/lib/jsdom/living/generated/Blob.js`). בלי ההזרקה
     * הזאת `exportDocx` מצליח, המעטפת נופלת בשורה שאחריו על
     * `exported.arrayBuffer is not a function`, ה-`catch` שלה נופל לטיוטה —
     * והבדיקה עוברת **גם על הקוד השבור**. זה נמדד בפועל בסבב מוטציה, ולא
     * הונח: בדיקה שעוברת על הקוד השבור גרועה מאין בדיקה.
     */
    const engineBytes = new Uint8Array([0, 0, 0, 0]);
    const engineDoc = Object.assign(new Blob([engineBytes]), {
      arrayBuffer: () => Promise.resolve(engineBytes.buffer),
    });
    (stub.superdoc!.host as unknown as { export: () => Promise<Blob> }).export = () =>
      Promise.resolve(engineDoc);

    // מעבר לטאב טוען את הטיוטה לתוך מנוע אמיתי, והמצביע ברשומה נשאר
    // (`keepDraft`) — כלומר מכאן והלאה יש גם מנוע וגם טיוטה.
    await wrapper.findAll('.word-doctab')[1]!.trigger('click');
    await settle(14);

    await wrapper.findAll('.word-doctab')[1]!.find('.word-doctab-close').trigger('click');
    await settle(12);
    await answerUnsaved('discard');

    const backup = stub.workspaceWrites.find((write) => write.path.startsWith('discarded-'));
    expect([...(backup?.bytes ?? [])], 'העבודה שבטיוטה — ולא המסמך שבמנוע').toEqual([
      80, 75, 3, 4,
    ]);
  });

  it('Escape סוגר את מסך השחזור גם כשהמיקוד כבר לא בתוכו', async () => {
    // „הסר” מוריד את הכפתור הממוקד מה-DOM, המיקוד נופל ל-`body`, ומאותו רגע
    // ההקשה מגיעה ל-`window` ולא לחלון. בלי ענף במעטפת היא הייתה ממקדת את
    // המסמך שמאחורי מודאל פתוח, והחלון היה נשאר על המסך.
    stub.storedDiscardBackups = [
      { slot: 0, name: 'ראשון', size: 4, discardedAt: Date.now(), token: null },
      { slot: 1, name: 'שני', size: 4, discardedAt: Date.now() - 1000, token: null },
    ];
    await mountShell();

    pressCtrl('KeyO');
    await settle(8);
    document.querySelector<HTMLButtonElement>('.open-discarded')!.click();
    await settle(8);
    document.querySelector<HTMLButtonElement>('.discarded-forget')!.click();
    await settle(12);
    expect(document.querySelector('.discarded-dialog'), 'עדיין פתוח — נשארה שורה').not.toBeNull();

    window.dispatchEvent(new KeyboardEvent('keydown', { code: 'Escape', key: 'Escape', bubbles: true }));
    await settle(8);

    expect(document.querySelector('.discarded-dialog'), 'נסגר').toBeNull();
  });

  it('Escape על „המסמך לא נשמר” הוא „ביטול” — גם מ-window', async () => {
    // הכיוון הבטוח: הקשה שאיש לא תפס אינה יכולה להיות אישור למחיקה.
    stub.storedSession = twoTabs(SECOND_DRAFT);
    stub.draftsByPath[SECOND_DRAFT.path] = new Uint8Array([80, 75, 3, 4]);
    const wrapper = await mountShell();

    await wrapper.findAll('.word-doctab')[1]!.find('.word-doctab-close').trigger('click');
    await settle(12);
    expect(unsavedDialog()).not.toBeNull();

    window.dispatchEvent(new KeyboardEvent('keydown', { code: 'Escape', key: 'Escape', bubbles: true }));
    await settle(12);

    expect(unsavedDialog(), 'נסגר').toBeNull();
    expect(wrapper.findAll('.word-doctab'), 'והטאב נשאר').toHaveLength(2);
    expect(stub.removedDrafts, 'ולא נמחק דבר').toEqual([]);
  });

  it('„הסר” מוחק את העותק מהדיסק ומהרשומה', async () => {
    stub.storedDiscardBackups = [
      { slot: 3, name: 'ישן', size: 4, discardedAt: Date.now() - 60_000, token: null },
    ];
    const wrapper = await mountShell();

    pressCtrl('KeyO');
    await settle(8);
    document.querySelector<HTMLButtonElement>('.open-discarded')!.click();
    await settle(8);
    document.querySelector<HTMLButtonElement>('.discarded-forget')!.click();
    await settle(12);

    expect(lastDiscardBackups(), 'הרשומה התרוקנה').toEqual([]);
    expect(stub.removedDrafts, 'והקובץ עצמו נמחק').toContain('discarded-3.docx');
    expect(document.querySelector('.discarded-dialog'), 'ומסך ריק נסגר').toBeNull();
    // ובלי גיבויים הכפתור אינו מופיע יותר.
    pressCtrl('KeyO');
    await settle(8);
    expect(document.querySelector('.open-discarded')).toBeNull();
    expect(wrapper.findAll('.word-doctab')).toHaveLength(1);
  });

  it.each(['forget', 'open'])('כשל בקריאת היסטוריית גיבויים ב-%s משאיר את הרשומה לניסיון נוסף', async (action) => {
    stub.storedDiscardBackups = [
      { slot: 1, name: 'עותק', size: 10, discardedAt: Date.now(), token: null },
    ];
    const wrapper = await mountShell();
    pressCtrl('KeyO');
    await settle(8);
    document.querySelector<HTMLButtonElement>('.open-discarded')!.click();
    await settle(8);
    stub.discardHistoryFailure = true;
    document.querySelector<HTMLButtonElement>(`.discarded-${action}`)!.click();
    await settle(12);
    expect(stub.persistedDiscardBackups).toHaveLength(0);
    expect(stub.removedDrafts).not.toContain('discarded-1.docx');
    expect(document.querySelector('.discarded-name')?.textContent).toBe('עותק');
    expect(wrapper.find('.word-statusbar').text()).toContain('אפשר לנסות שוב');
    expect(wrapper.find('.word-statusbar').text()).not.toContain('השורה הוסרה');
    expect(document.querySelector<HTMLButtonElement>('.discarded-forget')!.disabled).toBe(false);
    stub.discardHistoryFailure = false;
    document.querySelector<HTMLButtonElement>('.discarded-forget')!.click();
    await settle(12);
    expect(stub.removedDrafts).toContain('discarded-1.docx');
  });

  it('שורה שהעותק שלה נעלם מוסרת במקום לפתוח מסמך ריק', async () => {
    // קובץ שנמחק מבחוץ. שורה שאי אפשר לפתוח היא שורה שילחצו עליה שוב.
    stub.storedDiscardBackups = [
      { slot: 1, name: 'נעלם', size: 10, discardedAt: Date.now(), token: null },
    ];
    const wrapper = await mountShell();
    stub.openSources.length = 0;

    pressCtrl('KeyO');
    await settle(8);
    document.querySelector<HTMLButtonElement>('.open-discarded')!.click();
    await settle(8);
    document.querySelector<HTMLButtonElement>('.discarded-open')!.click();
    await settle(12);

    expect(stub.openSources, 'לא נפתח מסמך').toEqual([]);
    expect(lastDiscardBackups(), 'והשורה ירדה').toEqual([]);
    expect(wrapper.find('.word-statusbar').text()).toContain('לא נמצא');
  });

  it('„ביטול” על השאלה ביציאה משאיר את הכול — ואינו מגבה דבר', async () => {
    // גיבוי שנכתב על טאב שלא נסגר הוא רעש שגונב משבצת מאחד שכן נסגר.
    stub.storedSession = twoTabs(SECOND_DRAFT);
    stub.draftsByPath[SECOND_DRAFT.path] = new Uint8Array([80, 75, 3, 4]);
    const wrapper = await mountShell();

    const fileTab = wrapper.findAll('[role="tab"]').find((tab) => tab.text().includes('קובץ'));
    await fileTab!.trigger('click');
    await settle();
    await buttonByTip(wrapper, 'סגירת המסמך').trigger('click');
    await settle(16);
    await answerUnsaved('cancel');

    expect(tabTitles(wrapper), 'שני הטאבים נשארו').toEqual(['ראשון', 'שני']);
    expect(stub.persistedDiscardBackups, 'ולא נכתב גיבוי').toEqual([]);
    expect(stub.removedDrafts, 'ולא נמחקה טיוטה').toEqual([]);
  });
});

/**
 * התקרה על מספר המנועים החיים — ה-RAM של Issue #26.
 *
 * ## מה נמדד
 *
 * „טאב נרדם” הוא בדיוק אותו מצב של טאב ששוחזר ולא נטען, ולכן די בשתי שאלות:
 * האם המנוע באמת שוחרר (`swap.close`), והאם חזרה לטאב פותחת אותו מחדש עם
 * מה שהיה בו. השאלה השלישית היא הגבול: מסמך שאין לו קובץ אינו נרדם לעולם —
 * העותק היחיד שלו הוא הטיוטה שלו.
 */
describe('תקרת המסמכים החיים', () => {
  const OPENED = { token: 'tok-open', url: 'loopback://opened', name: 'פתוח.docx', size: 100 };

  /** „+”: טאב חדש, ובתוכו מסמך ריק — בלי קובץ. */
  async function openEmptyTab(wrapper: Awaited<ReturnType<typeof mountShell>>): Promise<void> {
    await wrapper.find('.word-doctabs-new').trigger('click');
    await settle(12);
  }

  it('ארבעה מסמכים חדשים בלי קובץ — אף אחד אינו נרדם', async () => {
    // הגבול של המנגנון, ולא היכולת שלו: מסמך שאין לו קובץ הוא מסמך שכל
    // קיומו תלוי בטיוטה שלו, והוא נשאר בזיכרון גם מעבר לתקרה.
    const wrapper = await mountShell();
    for (let index = 0; index < 3; index += 1) await openEmptyTab(wrapper);

    expect(wrapper.findAll('.word-doctab'), 'ארבעה טאבים, מעל התקרה של שלושה').toHaveLength(4);
    expect(stub.closedEngines).toBe(0);
  });

  it('מסמך שנפתח מקובץ נרדם, וחזרה אליו פותחת אותו מחדש', async () => {
    stub.storedSession = {
      version: 2,
      documents: [
        { id: 'doc-1', document: { token: 't1', name: 'א.docx', writable: true }, caret: null, draft: null },
        { id: 'doc-2', document: { token: 't2', name: 'ב.docx', writable: true }, caret: null, draft: null },
        { id: 'doc-3', document: { token: 't3', name: 'ג.docx', writable: true }, caret: null, draft: null },
        { id: 'doc-4', document: { token: 't4', name: 'ד.docx', writable: true }, caret: null, draft: null },
      ],
      activeId: 'doc-1',
      view: { zoom: null, focusMode: false, ribbonTab: null, ribbonCollapsed: false },
    };
    const NAMES: Record<string, string> = { t1: 'א.docx', t2: 'ב.docx', t3: 'ג.docx', t4: 'ד.docx' };
    stub.resolvedFile = (token: string) => ({
      ...OPENED,
      token,
      url: `loopback://${token}`,
      name: NAMES[token] ?? OPENED.name,
    });

    const wrapper = await mountShell();
    // הראשון נטען בעלייה; שלושת האחרים נטענים כשעוברים אליהם — והרביעי הוא
    // זה שמפיל את התקרה על הראשון.
    for (const index of [1, 2, 3]) {
      await wrapper.findAll('.word-doctab')[index]!.trigger('click');
      await settle(14);
    }

    expect(stub.closedEngines, 'בדיוק מנוע אחד שוחרר').toBe(1);
    expect(stub.openSources).toEqual([
      'loopback://t1',
      'loopback://t2',
      'loopback://t3',
      'loopback://t4',
    ]);

    // חזרה לטאב שנרדם פותחת אותו שוב, מהקובץ שלו.
    await wrapper.findAll('.word-doctab')[0]!.trigger('click');
    await settle(14);

    expect(stub.openSources[4]).toBe('loopback://t1');
    expect(wrapper.find('.word-doctab.active .word-doctab-title').text()).toBe('א');
  });
});

/**
 * מיקום הגלילה של מסמך במעבר טאב.
 *
 * הבאג שדווח: „עוברים טאב וחוזרים — הוא זוכר איפה אני, וברגע שמתחילים לגלול
 * הוא חוזר לראש”. הסיבה היא שהפאנל של טאב שאינו פעיל מוסתר ב-`display: none`,
 * וזה מוחק את מיקום הגלילה של מיכל הגלילה שבתוכו (ההנמקה המלאה
 * ב-sessions/pane-scroll.ts). הבדיקה כאן מודדת את החיווט: שהמעטפת קוראת את
 * המיקום **לפני** ההסתרה, ומחזירה אותו בכניסה.
 */
describe('מיקום הגלילה בין טאבים', () => {
  /** מיכל הגלילה של הטאב הפעיל, כפי שהמעטפת מוצאת אותו. */
  function activeHost(): HTMLElement | null {
    const panes = document.querySelectorAll<HTMLElement>('.document-pane');
    for (const pane of panes) {
      if (pane.style.display !== 'none') return pane.querySelector('.editor-stack__host');
    }
    return null;
  }

  it('גלילה בטאב אחד חוזרת אליו אחרי מעבר וחזרה', async () => {
    const wrapper = await mountShell();
    const first = activeHost();
    expect(first, 'למסמך הפתוח יש מיכל גלילה').not.toBeNull();
    first!.scrollTop = 640;

    // טאב שני, ואז חזרה לראשון.
    await wrapper.find('.word-doctabs-new').trigger('click');
    await settle(12);
    expect(activeHost(), 'הטאב החדש הוא מיכל אחר').not.toBe(first);

    /**
     * מה שהדפדפן עושה, ידנית.
     *
     * `display: none` הורס את קופסת הפריסה, ואיתה את מיקום הגלילה — אבל
     * ל-jsdom אין פריסה בכלל, ולכן `scrollTop` שורד אצלו בכל מקרה. בלי
     * האיפוס הזה הבדיקה הייתה עוברת בירוק גם אילו המעטפת לא שמרה דבר,
     * כלומר בדיקה שאינה מודדת את מה שהיא טוענת למדוד.
     */
    first!.scrollTop = 0;

    await wrapper.findAll('.word-doctab')[0]!.trigger('click');
    await settle(12);

    expect(activeHost()).toBe(first);
    expect(first!.scrollTop, 'המיקום חזר ולא אופס').toBe(640);
  });

  it('חזרה מהרקע מתקנת מיקום שאבד — ואינה נוגעת במיקום ששרד', async () => {
    /**
     * החיווט של „המשתמש חזר”: `onPluginShown` ← `repairPaneScroll`.
     *
     * שני הצדדים נבדקים בבידוד (tests/unit/lifecycle.test.ts,
     * tests/unit/pane-scroll.test.ts), אבל בדיוק ההרכבה היא מה שהקובץ הזה
     * קיים בשבילו: מחיקת הבלוק כולו מ-`onMounted` הייתה עוברת בירוק בשניהם.
     *
     * ל-jsdom אין פריסה, ולכן האיפוס נעשה כאן ביד — בדיוק כמו בבדיקת מעבר
     * הטאב שמעל, ומאותו טעם.
     */
    const wrapper = await mountShell();
    const host = activeHost()!;
    host.scrollTop = 480;

    window.dispatchEvent(new Event('pagehide'));
    await settle(6);
    host.scrollTop = 0;

    window.dispatchEvent(new Event('pageshow'));
    await settle(6);

    expect(host.scrollTop, 'המיקום שאבד תוקן').toBe(480);
    expect(wrapper.exists()).toBe(true);
  });

  it('חזרה מהרקע אינה גוררת מסמך שכבר במקום אחר', async () => {
    // הכלל שמונע את הנזק ההפוך: המסמך התעמד מחדש והמשתמש כבר במקום אחר.
    await mountShell();
    const host = activeHost()!;
    host.scrollTop = 480;

    window.dispatchEvent(new Event('pagehide'));
    await settle(6);
    host.scrollTop = 120;

    window.dispatchEvent(new Event('pageshow'));
    await settle(6);

    expect(host.scrollTop).toBe(120);
  });

  it('הטאב השני נשאר בראש המסמך שלו', async () => {
    // הכיוון ההפוך: המיקום שנשמר שייך לטאב שלו, ואינו נגרר לשכן.
    const wrapper = await mountShell();
    activeHost()!.scrollTop = 640;

    await wrapper.find('.word-doctabs-new').trigger('click');
    await settle(12);

    expect(activeHost()!.scrollTop).toBe(0);
  });
});

/**
 * קישוריות בין טאבים — מה שהמעטפת **מציגה** ומי שהיא **מדברת איתו** אחרי מעבר.
 *
 * ## למה הבדיקות האלה קיימות
 *
 * המעטפת מחזיקה עותק של מצב הטאב הפעיל (28 refs), ובכל מעבר טאב מעתיקה אותו
 * לסשן וחזרה. הבדיקות שמעל מאמתות כותרות, נקודות ומיקום גלילה — ואף אחת אינה
 * מפעילה פקד אחרי מעבר, ואף אחת אינה מבחינה בין מנועים: הכפיל הוא אובייקט אחד
 * לכל הפתיחות. נמדד: מחיקת `commandAdapter.value = ui.commandAdapter` מהשחזור
 * — כלומר רצועה שמדברת עם המנוע של הטאב הקודם — עוברת את כל 48 הבדיקות בירוק.
 *
 * כאן כל פתיחה מקבלת מנוע, אדפטר ומטריקות **משלה**, והאימות הוא דרך ה-DOM
 * ודרך מה שהאדפטר קיבל — לא דרך `vm`. זה נבנה לפני פיצול `App.vue` ונשאר תקף
 * אחריו, כי אף אחת מהבדיקות אינה תלויה בשם של פונקציה פנימית.
 */
describe('קישוריות בין טאבים', () => {
  /**
   * מנוע מזויף שמריץ את ה-disposers שנרשמו בו — בסדר הרשמה, כמו
   * `createEditor` (create-editor.ts: `disposers.splice(0)`). ברירת המחדל של
   * הקובץ (`onDispose: () => {}`) בולעת אותם, ולכן סדר השחרור לא היה מדיד.
   */
  function engineSession(): unknown {
    const disposers: Array<() => void> = [];
    const superdoc = createSuperdocDouble();
    superdoc.host.setDocumentMode = vi.fn();
    return {
      superdoc: superdoc.host,
      container: document.createElement('div'),
      ui: {
        selection: { observe: () => () => {}, apply: () => ({ ok: true }) },
        viewport: { scrollIntoView: async () => ({ success: true }) },
      },
      onDispose: (dispose: () => void) => {
        disposers.push(dispose);
      },
      destroy: () => {
        for (const dispose of disposers.splice(0)) dispose();
      },
    };
  }

  const metricsOf = (totalPages: number): DocMetrics => ({ words: totalPages * 100, totalPages, currentPage: 1 });

  /** מה שפס המצב מציג בתא „עמודי המסמך”. */
  function pageStatus(wrapper: Awaited<ReturnType<typeof mountShell>>): string {
    return wrapper.find('[data-tip-title="עמודי המסמך"]').text();
  }

  /** שני טאבים חיים עם כפילים נפרדים; מחזירה את האדפטרים לפי סדר הפתיחה. */
  async function twoLiveTabs(): Promise<{
    wrapper: Awaited<ReturnType<typeof mountShell>>;
    adapters: CommandDouble[];
  }> {
    const adapters: CommandDouble[] = [];
    const pages = [3, 7];
    stub.sessionFactory = engineSession;
    stub.adapterFactory = () => {
      const adapter = createCommandDouble({ states: { 'document-mode': { value: 'editing' } } });
      adapters.push(adapter);
      return adapter;
    };
    stub.metricsFactory = () => metricsOf(pages[adapters.length - 1] ?? 1);

    const wrapper = await mountShell();
    await wrapper.find('.word-doctabs-new').trigger('click');
    await settle(12);
    return { wrapper, adapters };
  }

  it('אחרי חזרה לטאב, פס המצב והרצועה מחוברים למסמך **שלו** ולא לזה שנעזב', async () => {
    const { wrapper, adapters } = await twoLiveTabs();
    expect(adapters, 'שני מנועים נפרדים').toHaveLength(2);
    expect(pageStatus(wrapper), 'הטאב השני פעיל ומציג את המספרים שלו').toContain('7');

    await wrapper.findAll('.word-doctab')[0]!.trigger('click');
    await settle(8);

    expect(pageStatus(wrapper), 'הפס חזר למספרים של הטאב הראשון').toContain('3');
    expect(pageStatus(wrapper)).not.toContain('7');

    // הלחיצה ברצועה מגיעה למנוע של הטאב הפעיל — ולא לזה שנעזב.
    const before = adapters.map((adapter) => adapter.calls.length);
    await buttonByTip(wrapper, 'מודגש').trigger('click');
    await settle();

    expect(adapters[0]!.calls.length - before[0]!, 'המנוע של הטאב הראשון קיבל את הפקודה').toBe(1);
    expect(adapters[0]!.calls[adapters[0]!.calls.length - 1]!.id).toBe('bold');
    expect(adapters[1]!.calls.length - before[1]!, 'המנוע של הטאב השני לא קיבל דבר').toBe(0);
  });

  it('בחירה ב„גבולות עמוד” וב„מספרי שורות” מרעננת את המודל מיד', async () => {
    // `applyPageBorders`/`applyLineNumbering` אינן מפעילות `onUpdate` של המנוע,
    // ולכן בלי `refreshNow()` המפורש ב-`reportCommand` הבחירה אינה מצטיירת עד
    // העריכה הבאה — „גבול רפאים”. הבדיקה עוברת את המסלול האמיתי: לשונית,
    // תפריט, פריט, `report` שהרצועה מזריקה, ואז המודל.
    const wrapper = await mountShell();
    stub.refreshCalls.length = 0;

    const layoutTab = wrapper.findAll('[role="tab"]').find((tab) => tab.text().includes('פריסה'));
    expect(layoutTab, 'לשונית „פריסה” קיימת').toBeDefined();
    await layoutTab!.trigger('click');
    await settle();

    for (const [tip, model] of [
      ['מסגרת סביב העמוד', 'pageBorders'],
      ['מספור השורות בשולי הדף', 'lineNumbers'],
    ] as const) {
      await buttonByTip(wrapper, tip).trigger('click');
      await settle();
      const item = wrapper.find('.ribbon-menu__item');
      expect(item.exists(), `התפריט של „${tip}” נפתח`).toBe(true);
      await item.trigger('click');
      await settle();
      expect(stub.refreshCalls, `המודל ${model} רוענן אחרי הבחירה`).toContain(model);
    }
  });

  it('ה-refs של התבנית מאוכלסים — המעטפת, ערימת העורכים ושכבת האיות', async () => {
    // `ref="name"` בתבנית נפתר מול הקישורים של App.vue עצמה. שם שמפסיק להיות
    // קישור ברמה העליונה (למשל אחרי העברה לקומפוזבל) משאיר את האובייקט `null`
    // לנצח, בשקט — ו„הוסף למילון” בתפריט ההקשר מת בלי אף שגיאה.
    const wrapper = await mountShell();
    const vm = wrapper.vm as unknown as Record<string, unknown>;

    expect(vm.shellRef, 'shellRef').not.toBeNull();
    expect(vm.editorStackRef, 'editorStackRef').not.toBeNull();
    expect(vm.spellingOverlayRef, 'spellingOverlayRef').not.toBeNull();
  });

  it('סגירת טאב ברקע משחררת את המודלים **שלו**, בסדר הרשמה, ואינה נוגעת בטאב הגלוי', async () => {
    const { wrapper } = await twoLiveTabs();
    await wrapper.findAll('.word-doctab')[0]!.trigger('click');
    await settle(8);
    expect(pageStatus(wrapper)).toContain('3');
    stub.disposeOrder.length = 0;

    // סגירת הטאב השני בזמן שהראשון פעיל — התרחיש שהשומרים הא-סימטריים
    // ב-`openDocumentInto` קיימים בשבילו, ושאף בדיקה לא הפעילה.
    await wrapper.findAll('.word-doctab')[1]!.find('.word-doctab-close').trigger('click');
    await settle(12);

    expect(wrapper.findAll('.word-doctab'), 'נשאר טאב אחד').toHaveLength(1);
    // בדיוק המודלים של הטאב שנסגר, ובסדר שבו `openDocumentInto` רשם אותם.
    expect(stub.disposeOrder).toEqual(['metrics', 'pageBorders', 'lineNumbers', 'search']);
    // והטאב הגלוי לא הרגיש דבר: הפס עדיין מציג את המסמך **שלו**.
    expect(pageStatus(wrapper)).toContain('3');
  });

  /**
   * באג חי, שנמדד ומתועד כאן עד שיתוקן: הנקודה של טאב **ברקע** אינה מתנקה
   * (ואינה מופיעה) כשהקואורדינטור שלו מדווח — עד מעבר הטאב הבא.
   *
   * הסיבה: `documentTabs` קורא `session.ui.saveSnapshot.isDirty` של טאב רקע
   * מאובייקט פשוט שאין לו תלות ריאקטיבית, בעוד הקואורדינטור כותב אליו
   * `session.ui.saveSnapshot = snapshot`.
   *
   * **תוקן.** `tabStripRevision` ב-App.vue הוא התלות המפורשת שהייתה חסרה,
   * ו-`onStateChange` מקדם אותו — כלומר הקואורדינטור של טאב ברקע מזיז את
   * הנקודה שלו מיד. הבדיקה עברה מ-`it.fails` ל-`it` יחד עם התיקון.
   */
  it('שמירה אוטומטית של טאב ברקע מעדכנת את הנקודה שלו', async () => {
    const { wrapper } = await twoLiveTabs();
    await wrapper.findAll('.word-doctab')[0]!.trigger('click');
    await settle(8);
    expect(stub.saveDepsList, 'קואורדינטור לכל טאב').toHaveLength(2);

    // `state` הוא שלב הפעולה (idle/exporting/…), לא „מלוכלך”: זה `isDirty`.
    const dirty: SaveSnapshot = {
      state: 'idle',
      isDirty: true,
      isSaving: false,
      targetToken: null,
      name: null,
      lastError: null,
    };
    stub.saveDepsList[1]!.onStateChange?.(dirty);
    await settle();

    expect(
      wrapper.findAll('.word-doctab')[1]!.find('.word-doctab-dirty').exists(),
      'הטאב ברקע מסומן כשהקואורדינטור שלו דיווח',
    ).toBe(true);
  });
});

/**
 * קיצורי הטאבים.
 *
 * הם נבדקים כאן ולא ב-`shortcuts-core.test.ts` מפני שאין להם מה למדוד בלי
 * רצועת טאבים אמיתית: „הבא” הוא המקום הבא **ברצועה**, ו„סגירה” היא המסלול
 * שמפרק session ומחליט מי מקבל את המקום. שתי ההכרעות האלה חיות ב-App.vue,
 * ומעטפת מזויפת הייתה בודקת את המזויפת.
 *
 * מה שנבדק כאן הוא ההבטחה שהמשתמש מכיר מכל תוכנה אחרת — ולא רק ש„המטפל
 * נקרא”: הגלישה מהסוף להתחלה, המיקום המוחלט, `Alt+9` שהוא „האחרון” ולא
 * „תשיעי”, והזהות בין שלושת הצירופים של „הבא”.
 */
describe('קיצורי הטאבים', () => {
  /** אינדקס הטאב הפעיל ברצועה, לפי הסדר שהמשתמש רואה. */
  function activeTab(wrapper: Awaited<ReturnType<typeof mountShell>>): number {
    return wrapper.findAll('.word-doctab').findIndex((tab) => tab.classes('active'));
  }

  function tabCount(wrapper: Awaited<ReturnType<typeof mountShell>>): number {
    return wrapper.findAll('.word-doctab').length;
  }

  /** מקיש צירוף על החלון — בדיוק כמו המשתמש, דרך המנתב האמיתי. */
  async function press(over: Partial<KeyboardEventInit> & { code: string }): Promise<void> {
    window.dispatchEvent(
      new KeyboardEvent('keydown', { cancelable: true, bubbles: true, ...over }),
    );
    await settle(12);
  }

  /** מעטפת עם `count` טאבים פתוחים, כשהאחרון הוא הפעיל — כמו אחרי „+”. */
  async function shellWithTabs(count: number): Promise<Awaited<ReturnType<typeof mountShell>>> {
    const wrapper = await mountShell();
    for (let index = 1; index < count; index += 1) {
      await wrapper.find('.word-doctabs-new').trigger('click');
      await settle(12);
    }
    return wrapper;
  }

  it('Ctrl+T פותח טאב נוסף ומפעיל אותו', async () => {
    const wrapper = await shellWithTabs(1);
    expect(tabCount(wrapper)).toBe(1);

    await press({ code: 'KeyT', ctrlKey: true });

    expect(tabCount(wrapper)).toBe(2);
    expect(activeTab(wrapper), 'הטאב החדש הוא הפעיל').toBe(1);
  });

  it('Ctrl+Tab מתקדם ברצועה, וגולש מהסוף להתחלה', async () => {
    // הגלישה אינה קישוט: היא מה שהופך את הצירוף לניתן ללחיצה חוזרת בלי
    // להסתכל, וכל דפדפן עושה אותה.
    const wrapper = await shellWithTabs(3);
    expect(activeTab(wrapper), 'אחרי „+” פעמיים הפעיל הוא האחרון').toBe(2);

    await press({ code: 'Tab', ctrlKey: true });
    expect(activeTab(wrapper)).toBe(0);

    await press({ code: 'Tab', ctrlKey: true });
    expect(activeTab(wrapper)).toBe(1);
  });

  it('Ctrl+Shift+Tab חוזר אחורה, וגולש מההתחלה לסוף', async () => {
    const wrapper = await shellWithTabs(3);

    await press({ code: 'Tab', ctrlKey: true, shiftKey: true });
    expect(activeTab(wrapper)).toBe(1);

    await press({ code: 'Tab', ctrlKey: true, shiftKey: true });
    expect(activeTab(wrapper)).toBe(0);

    await press({ code: 'Tab', ctrlKey: true, shiftKey: true });
    expect(activeTab(wrapper), 'מהראשון אחורה — לאחרון').toBe(2);
  });

  it('Ctrl+Page Down ו-Ctrl+F6 הם אותו „הבא” בדיוק', async () => {
    // שלושה צירופים לאותה פעולה, כי שלוש קהילות משתמשים שונות: הדפדפן,
    // VSCode ו-Word. אם הם מתפצלים, אחד מהם משקר.
    const wrapper = await shellWithTabs(3);

    await press({ code: 'PageDown', ctrlKey: true });
    expect(activeTab(wrapper)).toBe(0);

    await press({ code: 'F6', ctrlKey: true });
    expect(activeTab(wrapper)).toBe(1);

    await press({ code: 'PageUp', ctrlKey: true });
    expect(activeTab(wrapper)).toBe(0);

    await press({ code: 'F6', ctrlKey: true, shiftKey: true });
    expect(activeTab(wrapper)).toBe(2);
  });

  it('F6 בלי Ctrl אינו מזיז טאב — הוא מעבר בין אזורי הממשק', async () => {
    // שתי רשומות על אותו מקש פיזי, ורק המודיפייר מבדיל. `match.ts` דורש
    // התאמה מדויקת, וזו הבדיקה ששומרת עליה מהצד של המשתמש.
    const wrapper = await shellWithTabs(2);
    const before = activeTab(wrapper);

    await press({ code: 'F6' });

    expect(activeTab(wrapper)).toBe(before);
  });

  it('Alt+1 ו-Alt+2 עוברים למיקום מוחלט ברצועה', async () => {
    const wrapper = await shellWithTabs(3);

    await press({ code: 'Digit1', altKey: true });
    expect(activeTab(wrapper)).toBe(0);

    await press({ code: 'Digit2', altKey: true });
    expect(activeTab(wrapper)).toBe(1);
  });

  it('Alt+9 הוא הטאב האחרון, ולא „טאב מספר תשע”', async () => {
    const wrapper = await shellWithTabs(3);
    await press({ code: 'Digit1', altKey: true });
    expect(activeTab(wrapper)).toBe(0);

    await press({ code: 'Digit9', altKey: true });

    expect(activeTab(wrapper)).toBe(2);
  });

  it('Alt+ספרה שאין לה טאב אינו מזיז דבר', async () => {
    const wrapper = await shellWithTabs(2);
    await press({ code: 'Digit1', altKey: true });

    await press({ code: 'Digit7', altKey: true });

    expect(activeTab(wrapper)).toBe(0);
    expect(tabCount(wrapper)).toBe(2);
  });

  it('Ctrl+W סוגר את הטאב שעל המסך', async () => {
    const wrapper = await shellWithTabs(3);

    await press({ code: 'KeyW', ctrlKey: true });

    expect(tabCount(wrapper)).toBe(2);
  });

  it('Ctrl+F4 סוגר גם הוא — הצירוף של Word לאותה פעולה', async () => {
    const wrapper = await shellWithTabs(3);

    await press({ code: 'F4', ctrlKey: true });

    expect(tabCount(wrapper)).toBe(2);
  });

  it('Ctrl+W על הטאב האחרון משאיר מסמך ריק ולא מעטפת בלי טאבים', async () => {
    // הרצועה, הפס והסרגלים כולם מניחים מסמך. „אפס טאבים” אינו מצב שקיים כאן,
    // ובדפדפן המקבילה שלו היא סגירת החלון — מה שאין לנו רשות לעשות.
    const wrapper = await shellWithTabs(1);

    await press({ code: 'KeyW', ctrlKey: true });

    expect(tabCount(wrapper)).toBe(1);
  });

  it('Ctrl+Shift+T פותח מחדש את הקובץ של הטאב שנסגר', async () => {
    stub.storedSession = {
      version: 2,
      documents: [
        { id: 'doc-1', document: FIRST_DOC, caret: null, draft: null },
        { id: 'doc-2', document: SECOND_DOC, caret: null, draft: null },
      ],
      activeId: 'doc-1',
      view: { zoom: null, focusMode: false, ribbonTab: null, ribbonCollapsed: false },
    };
    stub.resolvedFile = (token: string) => TAB_FILES[token] ?? null;

    const wrapper = await mountShell();
    expect(stub.openSources).toEqual(['loopback://first']);

    // הטאב הפעיל נקי (נטען זה עתה מהדיסק), ולכן אין שאלה — והשני נטען
    // כשהוא מקבל את המקום.
    await press({ code: 'KeyW', ctrlKey: true });
    expect(tabCount(wrapper)).toBe(1);
    expect(stub.openSources).toEqual(['loopback://first', 'loopback://second']);

    await press({ code: 'KeyT', ctrlKey: true, shiftKey: true });

    expect(tabCount(wrapper), 'הקובץ חוזר לטאב משלו').toBe(2);
    expect(stub.openSources).toEqual([
      'loopback://first',
      'loopback://second',
      'loopback://first',
    ]);
  });

  it('Ctrl+Shift+T בלי טאב שנסגר אומר זאת, ואינו פותח דבר', async () => {
    const wrapper = await shellWithTabs(1);
    const opened = stub.openSources.length;

    await press({ code: 'KeyT', ctrlKey: true, shiftKey: true });

    expect(tabCount(wrapper)).toBe(1);
    expect(stub.openSources).toHaveLength(opened);
    expect(wrapper.text()).toContain('אין טאב שנסגר');
  });

  it('טאב שנפתח מחדש נשאר בר-כתיבה — הבאג שהיה', async () => {
    /*
     * הבאג: `reopenClosedTab` פתח דרך `resolveFileUrl`, שמחזירה
     * `{token, url, name, size}` **בלי `access`**. הפתיחה נראתה מוצלחת
     * לגמרי, והמסמך הפך בשקט לקריאה-בלבד — `save` בלי יעד כתיבה, ו-
     * `writable: false` שנכתב לרשומת ההפעלה, כלומר גם אחרי הפעלה מחדש.
     *
     * הסימן הנצפה הוא „פתוח לקריאה” בשורת המצב, אותו אחד שנמדד ב„הודעת
     * השלב המקדים”. ה-stub מחזיר כאן בדיוק את מה שהגשר האמיתי מחזיר — בלי
     * `access` — ולכן הבדיקה נופלת על התיקון ולא על זיוף.
     */
    stub.storedSession = {
      version: 2,
      documents: [{ id: 'doc-1', document: FIRST_DOC, caret: null, draft: null }],
      activeId: 'doc-1',
      view: { zoom: null, focusMode: false, ribbonTab: null, ribbonCollapsed: false },
    };
    stub.resolvedFile = (token: string) => TAB_FILES[token] ?? null;

    const wrapper = await mountShell();
    const status = () => wrapper.find('.word-statusbar').text();
    expect(status(), 'נפתח בכתיבה מלכתחילה').not.toContain('פתוח לקריאה');

    await press({ code: 'KeyW', ctrlKey: true });
    await press({ code: 'KeyT', ctrlKey: true, shiftKey: true });

    expect(stub.openSources, 'הקובץ אכן נפתח מחדש').toContain('loopback://first');
    expect(status(), 'ההרשאה שרדה את הסגירה והפתיחה').not.toContain('פתוח לקריאה');
  });

  it('טאב שהיה קריאה-בלבד נפתח מחדש כקריאה-בלבד', async () => {
    // הכיוון ההפוך, ומאותה סיבה: ההרשאה נלקחת ממה שהטאב ידע, ולכן היא חייבת
    // לשמר גם „לא” ולא רק „כן”. בלי זה התיקון היה יכול להיות „תמיד כתיב”,
    // וזה כשל חמור יותר — ניסיון כתיבה ל-token שאין עליו הרשאה.
    stub.storedSession = {
      version: 2,
      documents: [
        {
          id: 'doc-1',
          document: { ...FIRST_DOC, writable: false },
          caret: null,
          draft: null,
        },
      ],
      activeId: 'doc-1',
      view: { zoom: null, focusMode: false, ribbonTab: null, ribbonCollapsed: false },
    };
    stub.resolvedFile = (token: string) => TAB_FILES[token] ?? null;

    const wrapper = await mountShell();

    await press({ code: 'KeyW', ctrlKey: true });
    await press({ code: 'KeyT', ctrlKey: true, shiftKey: true });

    expect(wrapper.find('.word-statusbar').text()).toContain('פתוח לקריאה');
  });

  it('טאב חדש שלא נשמר מעולם אינו נכנס למחסנית „נסגר”', async () => {
    // אין לו token, ולכן אין ממה לפתוח אותו מחדש. רישום שלו היה מייצר
    // `Ctrl+Shift+T` שנכשל תמיד — הבטחה שאי אפשר לקיים.
    const wrapper = await shellWithTabs(2);
    const opened = stub.openSources.length;

    await press({ code: 'KeyW', ctrlKey: true });
    await press({ code: 'KeyT', ctrlKey: true, shiftKey: true });

    expect(tabCount(wrapper)).toBe(1);
    expect(stub.openSources).toHaveLength(opened);
    expect(wrapper.text()).toContain('אין טאב שנסגר');
  });
});

describe('פעולה כבדה — עותק לפני, והקפאה לאורכה', () => {
  /**
   * זו הבדיקה מקצה לקצה של המסלול שכל התכונה נבנתה בשבילו: פעולה אחת שמשנה
   * את כל המסמך, ושאם התוסף קורס תוך כדי — אין ממנה דרך חזרה.
   *
   * `arrayBuffer` מוזרק, ובלי זה הבדיקה חסרת ערך: jsdom אינו מממש
   * `Blob.prototype.arrayBuffer`, ולכן כתיבת הצ׳קפוינט הייתה נכשלת בשקט
   * וה-`catch` שלה היה מחזיר „הצליח”. ראו את ההנמקה המלאה ליד הגיבוי של
   * „לא לשמור”.
   */
  function fakeExport(): void {
    const bytes = new Uint8Array([80, 75, 3, 4]);
    const doc = Object.assign(new Blob([bytes]), {
      arrayBuffer: () => Promise.resolve(bytes.buffer),
    });
    (stub.superdoc!.host as unknown as { export: () => Promise<Blob> }).export = () =>
      Promise.resolve(doc);
  }

  it('„החלף הכל” כותב עותק לפני, ומקפיא ומשחרר את השמירה האוטומטית', async () => {
    const wrapper = await mountShell();
    fakeExport();
    const before = stub.workspaceWrites.length;

    wrapper.findComponent(FindReplaceDialog).vm.$emit('replace-all', 'רשב״י');
    await settle(12);

    const written = stub.workspaceWrites.slice(before).filter((write) => write.path.startsWith('session-draft-'));
    expect(written.length, 'עותק אחד לפני הפעולה').toBe(1);
    expect([...(written[0]?.bytes ?? [])], 'הבייטים של המסמך, לא מסמך ריק').toEqual([80, 75, 3, 4]);
    expect(stub.holdCalls, 'ההקפאה נלקחה').toBe(1);
    expect(stub.releaseCalls, 'ושוחררה בסוף').toBe(1);
  });

  it('ההקפאה משוחררת גם כשהפעולה נכשלת', async () => {
    // בלי `finally` המסמך היה נשאר בלי שמירה אוטומטית עד מעבר מסמך.
    const wrapper = await mountShell();
    fakeExport();
    stub.replaceAllFails = true;

    wrapper.findComponent(FindReplaceDialog).vm.$emit('replace-all', 'רשב״י');
    await settle(12);

    expect(stub.holdCalls).toBe(1);
    expect(stub.releaseCalls).toBe(1);
  });
});

describe('הגרסה שלפני השמירה האחרונה', () => {
  /** רשומה שבה המסמך נשמר, והעותק שלפני השמירה עדיין במקומו. */
  function sessionWithPrevious(documentToken: string | null = 'tok'): unknown {
    return {
      version: 2,
      documents: [
        {
          id: 'doc-1',
          document: { token: 'tok', name: 'חידושים.docx', writable: true },
          caret: null,
          draft: null,
          previousDraft: {
            path: 'session-draft-doc-1-a.docx',
            savedAt: Date.now() - 120_000,
            documentToken,
            sourceSize: 120,
          },
        },
      ],
      activeId: 'doc-1',
      view: { zoom: null, focusMode: false, ribbonTab: null, ribbonCollapsed: false },
    };
  }

  it('נפתחת בטאב חדש, והקובץ עצמו אינו נוגע', async () => {
    // הבטיחות של המסלול היא בדיוק בשני אלה: המסמך שעל המסך אינו מוחלף,
    // ומה שנפתח אין לו יעד כתיבה — „שמור” בו פותח „שמור בשם”.
    stub.storedSession = sessionWithPrevious();
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };
    stub.draftsByPath['session-draft-doc-1-a.docx'] = new Uint8Array([80, 75, 3, 4]);

    const wrapper = await mountShell();
    stub.openSources.length = 0;

    pressCtrl('KeyO');
    await settle(8);
    const entryPoint = document.querySelector<HTMLButtonElement>('.open-previous');
    expect(entryPoint, 'הכפתור מופיע כשיש גרסה קודמת').not.toBeNull();

    entryPoint!.click();
    await settle(16);

    expect(stub.openSources[0], 'הבייטים של הגרסה הקודמת').toBeInstanceOf(Blob);
    // שתי טענות ולא אחת: „הכותרת מופיעה” נכון גם כשהטאב הפתוח **הוחלף**,
    // וזו בדיוק התוצאה שהמסלול הזה אסור לו להגיע אליה.
    const titles = wrapper.findAll('.word-doctab-title').map((el) => el.text());
    expect(titles, 'שני טאבים — הפתוח נשאר').toHaveLength(2);
    expect(titles[0], 'המסמך שהיה על המסך לא זז').toBe('חידושים');
    expect(titles[1]).toContain('גרסה קודמת');
    expect(wrapper.find('.word-statusbar').text()).toContain('מלפני השמירה האחרונה');
    expect(stub.draftRemovals, 'פתיחה אינה מוחקת את העותק').toBe(0);
  });

  it('גרסה קודמת של מסמך אחר אינה מוצעת', async () => {
    // המצביע שורד מעבר מסמך בכוונה, ולכן בלי בדיקת ה-token הכפתור היה מציע
    // לפתוח גרסה של משהו אחר לגמרי.
    stub.storedSession = sessionWithPrevious('other');
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };
    stub.draftsByPath['session-draft-doc-1-a.docx'] = new Uint8Array([80, 75, 3, 4]);

    await mountShell();

    pressCtrl('KeyO');
    await settle(8);

    expect(document.querySelector('.open-previous')).toBeNull();
  });

  it('כשהעותק אינו במקומו — נאמר, ולא נפתח מסמך ריק', async () => {
    stub.storedSession = sessionWithPrevious();
    stub.resolvedFile = { token: 'tok', url: 'loopback://fresh', name: 'חידושים.docx', size: 120 };
    // אין רשומה ב-`draftsByPath` — הקריאה תחזיר null.

    const wrapper = await mountShell();
    stub.openSources.length = 0;

    pressCtrl('KeyO');
    await settle(8);
    document.querySelector<HTMLButtonElement>('.open-previous')!.click();
    await settle(12);

    expect(stub.openSources, 'לא נפתח דבר').toHaveLength(0);
    expect(wrapper.find('.word-statusbar').text()).toContain('לא נמצאה');
  });
});

describe('explicit startup opening before engine restoration', () => {
  const scope = window as unknown as Record<string, unknown>;
  const view = { zoom: null, focusMode: false, ribbonTab: null, ribbonCollapsed: false };
  const draft = { path: 'session-draft-doc-1.docx', savedAt: 1, documentToken: 'tok-1', sourceSize: 100 };

  function previous(options: { draft?: boolean; writable?: boolean; caret?: boolean } = {}) {
    const stored = {
      version: 2,
      documents: [{
        id: 'doc-1', document: { ...FIRST_DOC, writable: options.writable !== false },
        caret: options.caret ? { start: { blockId: 'b9', ordinal: 8, offset: 4 }, end: null } : null,
        draft: options.draft ? draft : null,
      }],
      activeId: 'doc-1', view,
    };
    stub.storedSession = stored;
    stub.resolvedFile = (token: string) => TAB_FILES[token] ?? null;
    if (options.draft) stub.draftsByPath[draft.path] = new Uint8Array([80, 75, 3, 4]);
    return stored;
  }

  function book(token = 'tok-2') {
    return {
      itemId: 'otzaria-word-edit-book', context: 'reader-text', bookUid: token,
      currentBook: token === 'tok-1' ? 'ראשון' : 'שני', currentType: 'docx', currentSource: 'user',
    };
  }

  function host(reply?: (token: string) => Promise<unknown>) {
    const listeners = new Map<string, (event: unknown) => void>();
    window.Otzaria = {
      call: async (method: string, payload: { bookUid?: string }) => {
        if (method === 'library.openBookFile') {
          const token = payload.bookUid!;
          return reply ? reply(token) : { success: true, data: { ...TAB_FILES[token] as object, access: 'readwrite' } };
        }
        return { success: true, data: {} };
      },
      on: (name: string, listener: (event: unknown) => void) => listeners.set(name, listener),
      off: (name: string) => listeners.delete(name),
    } as never;
    return (name: string, event: unknown) => listeners.get(name)!(event);
  }

  function pendingBooks(...events: unknown[]) {
    scope.__otzariaToolbarClicks = { queue: events, live: false };
  }

  afterEach(() => {
    delete scope.__otzariaToolbarClicks;
    delete scope.__otzariaContextMenuClicks;
    delete scope.__otzariaPageOpened;
    delete (window as Partial<Window>).Otzaria;
  });

  it('opens B first while A and its draft stay unloaded until A is selected', async () => {
    previous({ draft: true }); host(); pendingBooks(book());
    const wrapper = await mountShell();
    await settle(12);
    expect(stub.openSources).toEqual(['loopback://second']);
    expect(wrapper.findAll('.word-doctab')).toHaveLength(2);
    expect(wrapper.find('.word-doctab.active .word-doctab-title').text()).toBe('שני');
    const record = stub.persistedSessions[stub.persistedSessions.length - 1] as {
      documents: Array<{ id: string; draft: unknown }>;
    };
    expect(record.documents.find(item => item.id === 'doc-1')?.draft).toEqual(draft);
    expect(stub.removedDrafts).toEqual([]);
    await wrapper.findAll('.word-doctab')[0]!.trigger('click');
    await settle(16);
    expect(stub.openSources).toHaveLength(2);
    expect(stub.openSources[1]).toBeInstanceOf(Blob);
    expect(stub.markRestoredCalls).toBe(1);
  });

  it('keeps an unloaded read-only A instead of replacing its metadata with B', async () => {
    previous({ writable: false }); host(); pendingBooks(book());
    const wrapper = await mountShell(); await settle(12);
    expect(stub.openSources).toEqual(['loopback://second']);
    expect(wrapper.findAll('.word-doctab')).toHaveLength(2);
    await wrapper.findAll('.word-doctab')[0]!.trigger('click'); await settle(16);
    expect(stub.openSources).toEqual(['loopback://second', 'loopback://first']);
  });

  it.each([true, false])('deduplicates two requests for A and restores its draft (writable=%s)', async (writable) => {
    previous({ draft: true, writable }); host(); pendingBooks(book('tok-1'), book('tok-1'));
    const wrapper = await mountShell(); await settle(20);
    expect(stub.openSources).toHaveLength(1);
    expect(stub.openSources[0]).toBeInstanceOf(Blob);
    expect(wrapper.findAll('.word-doctab')).toHaveLength(1);
    expect(stub.markRestoredCalls).toBe(1);
    expect(stub.removedDrafts).toEqual([]);
  });

  it('accepts an explicit request arriving while startup preferences are pending', async () => {
    const stored = previous();
    let release!: (value: unknown) => void;
    stub.storedSession = new Promise(resolve => { release = resolve; });
    const emit = host();
    const wrapper = await mountShell();
    try {
      expect(stub.openSources).toEqual([]);
      emit('reader.toolbar_item_clicked', book());
      release(stored); await settle(20);
      expect(stub.openSources).toEqual(['loopback://second']);
    } finally { release(stored); await settle(4); wrapper.unmount(); }
  });

  it.each(['rejected', 'host-failed', 'engine-failed'])('restores A once after a %s request', async (failure) => {
    previous();
    host(failure === 'host-failed' ? async () => ({ success: false, error: { message: 'RPC failed' } }) : undefined);
    pendingBooks(failure === 'rejected' ? { ...book(), currentType: 'txt' } : book());
    if (failure === 'engine-failed') stub.openFailures = 1;
    const wrapper = await mountShell(); await settle(20);
    expect(stub.openSources.filter(source => source === 'loopback://first')).toHaveLength(1);
    expect(stub.openSources).toHaveLength(failure === 'engine-failed' ? 2 : 1);
    expect(wrapper.findAll('.word-doctab')).toHaveLength(1);
    expect(wrapper.find('.word-statusbar').text()).toMatch(/אינו קובץ Word|RPC failed|לא נפתח/);
  });

  it('does not retry a failed explicit opening of the old active A', async () => {
    previous({ draft: true }); host(); pendingBooks(book('tok-1'));
    stub.openFailures = 1;
    const wrapper = await mountShell(); await settle(20);
    expect(stub.openSources).toHaveLength(2);
    expect(stub.openSources[0]).toBeInstanceOf(Blob);
    expect(stub.openSources[1]).toBeUndefined();
    expect(wrapper.findAll('.word-doctab')).toHaveLength(1);
    expect(stub.removedDrafts).toEqual([]);
  });

  it('a received file opens first and preserves both its recovery draft and old A', async () => {
    previous({ draft: true }); host();
    scope.__otzariaPageOpened = {
      queue: [{ param: { name: 'מתוסף.docx', dataUrl: 'data:application/octet-stream;base64,UEsDBA==' } }], live: false,
    };
    const descriptor = Object.getOwnPropertyDescriptor(Blob.prototype, 'arrayBuffer');
    Object.defineProperty(Blob.prototype, 'arrayBuffer', {
      configurable: true,
      value(this: Blob) {
        return new Promise<ArrayBuffer>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result as ArrayBuffer);
          reader.onerror = () => reject(reader.error);
          reader.readAsArrayBuffer(this);
        });
      },
    });
    let wrapper: Awaited<ReturnType<typeof mountShell>> | undefined;
    try {
      wrapper = await mountShell(); await settle(24);
      expect(stub.openSources).toHaveLength(1);
      expect(stub.openSources[0]).toBeInstanceOf(Blob);
      expect(wrapper.findAll('.word-doctab')).toHaveLength(2);
      expect(wrapper.find('.word-doctab.active .word-doctab-title').text()).toBe('מתוסף');
      await vi.waitFor(() => expect(stub.workspaceWrites.some(write => write.path.startsWith('session-draft-'))).toBe(true));
      const record = stub.persistedSessions[stub.persistedSessions.length - 1] as {
        documents: Array<{ id: string; document: unknown; draft: unknown }>;
      };
      expect(record.documents.find(item => item.id === 'doc-1')?.draft).toEqual(draft);
      expect(record.documents.find(item => item.id !== 'doc-1')).toMatchObject({ document: null, draft: expect.any(Object) });
      expect(stub.removedDrafts).toEqual([]);
    } finally {
      wrapper?.unmount();
      if (descriptor) Object.defineProperty(Blob.prototype, 'arrayBuffer', descriptor);
      else delete (Blob.prototype as { arrayBuffer?: unknown }).arrayBuffer;
    }
  });

  it('an earlier blank fallback is not mistaken for an already opened file', async () => {
    previous({ writable: false }); const emit = host();
    stub.openFailures = 1;
    const wrapper = await mountShell(); await settle(12);
    expect(stub.openSources).toEqual(['loopback://first', undefined]);
    emit('reader.toolbar_item_clicked', book('tok-1')); await settle(20);
    expect(stub.openSources).toEqual(['loopback://first', undefined, 'loopback://first']);
    expect(wrapper.find('.word-doctab.active .word-doctab-title').text()).toBe('ראשון');
  });

  it('preserves request order and skips the old engine for the first explicit request', async () => {
    previous(); host(); pendingBooks(book(), book('tok-1'), book());
    const wrapper = await mountShell(); await settle(28);
    expect(stub.openSources).toEqual(['loopback://second', 'loopback://first']);
    expect(wrapper.findAll('.word-doctab')).toHaveLength(2);
    expect(wrapper.find('.word-doctab.active .word-doctab-title').text()).toBe('שני');
  });

  it('sends text once to restored A after caret restoration, before processing B', async () => {
    previous({ caret: true }); host(); pendingBooks(book());
    scope.__otzariaContextMenuClicks = {
      queue: [{ itemId: 'otzaria-word-send-to-document', selectedText: 'קטע לבדיקה' }], live: false,
    };
    const inserted: Array<{ source: unknown; caret: unknown }> = [];
    const doc = stub.superdoc!.host.activeEditor!.doc as unknown as { insert(input: unknown): unknown };
    doc.insert = () => {
      inserted.push({ source: stub.openSources[stub.openSources.length - 1], caret: stub.caretApplied[0] });
      return { success: true };
    };
    await mountShell(); await settle(28);
    expect(stub.openSources).toEqual(['loopback://first', 'loopback://second']);
    expect(inserted).toEqual([{
      source: 'loopback://first',
      caret: { kind: 'selection', start: { kind: 'text', blockId: 'b9', offset: 4 }, end: { kind: 'text', blockId: 'b9', offset: 4 } },
    }]);
  });

  it('independent QA: early B with no stored session reuses the initial placeholder', async () => {
    host(); pendingBooks(book());
    const wrapper = await mountShell();
    await vi.waitFor(() => expect(stub.openSources).toEqual(['loopback://second']));
    expect(wrapper.findAll('.word-doctab')).toHaveLength(1);
    expect(wrapper.find('.word-doctab.active .word-doctab-title').text()).toBe('שני');
    const record = stub.persistedSessions[stub.persistedSessions.length - 1] as {
      documents: Array<{ document: { token: string } | null }>;
    };
    expect(record.documents).toHaveLength(1);
    expect(record.documents[0]?.document?.token).toBe('tok-2');
  });

  it('independent QA: two received blobs keep their order and individual recovery drafts', async () => {
    previous({ draft: true }); host();
    scope.__otzariaPageOpened = {
      queue: [
        { param: { name: 'בית.docx', dataUrl: 'data:application/octet-stream;base64,UEsDBA==' } },
        { param: { name: 'גימל.docx', dataUrl: 'data:application/octet-stream;base64,UEsDBQ==' } },
      ], live: false,
    };
    const descriptor = Object.getOwnPropertyDescriptor(Blob.prototype, 'arrayBuffer');
    Object.defineProperty(Blob.prototype, 'arrayBuffer', {
      configurable: true,
      value(this: Blob) {
        return new Promise<ArrayBuffer>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result as ArrayBuffer);
          reader.onerror = () => reject(reader.error);
          reader.readAsArrayBuffer(this);
        });
      },
    });
    let wrapper: Awaited<ReturnType<typeof mountShell>> | undefined;
    try {
      wrapper = await mountShell();
      await vi.waitFor(() => {
        const record = stub.persistedSessions[stub.persistedSessions.length - 1] as {
          documents: Array<{ id: string; draft: unknown }>;
        };
        expect(record.documents.filter(item => item.id !== 'doc-1' && item.draft)).toHaveLength(2);
      });
      expect(stub.openSources).toHaveLength(2);
      expect(Array.from(new Uint8Array(await (stub.openSources[0] as Blob).arrayBuffer()))).toEqual([80, 75, 3, 4]);
      expect(Array.from(new Uint8Array(await (stub.openSources[1] as Blob).arrayBuffer()))).toEqual([80, 75, 3, 5]);
      const record = stub.persistedSessions[stub.persistedSessions.length - 1] as {
        documents: Array<{ id: string; document: unknown; draft: { path: string } | null }>;
      };
      expect(record.documents.find(item => item.id === 'doc-1')?.draft).toEqual(draft);
      const added = record.documents.filter(item => item.id !== 'doc-1');
      expect(added.every(item => item.document === null)).toBe(true);
      expect(new Set(added.map(item => item.draft!.path)).size).toBe(2);
      expect(stub.removedDrafts).toEqual([]);
      expect(wrapper.findAll('.word-doctab')).toHaveLength(3);
      expect(wrapper.find('.word-doctab.active .word-doctab-title').text()).toBe('גימל');
    } finally {
      wrapper?.unmount();
      if (descriptor) Object.defineProperty(Blob.prototype, 'arrayBuffer', descriptor);
      else delete (Blob.prototype as { arrayBuffer?: unknown }).arrayBuffer;
    }
  });

  it('independent QA: pending startup opening does not create a new session after unmount', async () => {
    previous();
    let release!: () => void;
    let resolving = false;
    const gate = new Promise<void>(resolve => { release = resolve; });
    host(async token => { resolving = true; await gate; return { success: true, data: { ...TAB_FILES[token] as object, access: 'readwrite' } }; });
    pendingBooks(book());
    const wrapper = await mountShell();
    try {
      await vi.waitFor(() => expect(resolving).toBe(true));
      const sessions = scope.__otzariaEditors as Map<string, unknown>;
      expect(stub.openSources).toEqual([]);
      const count = sessions.size;
      wrapper.unmount();
      release(); await settle(28);
      expect(stub.openSources, 'no engine may be created after its owner unmounted').toEqual([]);
      expect(sessions.size, 'no new detached DocumentSession may survive').toBe(count);
    } finally { release(); await settle(4); wrapper.unmount(); }
  });

  it('independent QA: late send keeps an originally empty target separate from explicit B', async () => {
    stub.storedSession = {
      version: 2, activeId: 'doc-1', view,
      documents: [{ id: 'doc-1', document: null, draft: null, caret: null }],
    };
    let release!: () => void;
    let resolving = false;
    const gate = new Promise<void>(resolve => { release = resolve; });
    const emit = host(async token => { resolving = true; await gate; return { success: true, data: { ...TAB_FILES[token] as object, access: 'readwrite' } }; });
    pendingBooks(book());
    const inserted: unknown[] = [];
    const doc = stub.superdoc!.host.activeEditor!.doc as unknown as { insert(input: unknown): unknown };
    doc.insert = () => { inserted.push(stub.openSources[stub.openSources.length - 1]); return { success: true }; };
    const wrapper = await mountShell();
    try {
      await vi.waitFor(() => expect(resolving).toBe(true));
      emit('contextMenu.itemClicked', { itemId: 'otzaria-word-send-to-document', selectedText: 'קטע' });
      release(); await settle(28);
      expect(inserted).toEqual([undefined]);
      expect(stub.openSources).toEqual(['loopback://second', undefined]);
    } finally { release(); await settle(4); wrapper.unmount(); }
  });

  it('independent QA: invalid empty send does not preempt a valid explicit opening', async () => {
    previous(); host(); pendingBooks(book());
    scope.__otzariaContextMenuClicks = {
      queue: [{ itemId: 'otzaria-word-send-to-document', selectedText: '' }], live: false,
    };
    await mountShell(); await settle(28);
    expect(stub.openSources).toEqual(['loopback://second']);
  });

  it('a send arriving during B resolution still goes to the original A', async () => {
    previous();
    let release!: () => void;
    const gate = new Promise<void>(resolve => { release = resolve; });
    const emit = host(async token => { await gate; return { success: true, data: { ...TAB_FILES[token] as object, access: 'readwrite' } }; });
    pendingBooks(book());
    const inserted: unknown[] = [];
    const doc = stub.superdoc!.host.activeEditor!.doc as unknown as { insert(input: unknown): unknown };
    doc.insert = () => { inserted.push(stub.openSources[stub.openSources.length - 1]); return { success: true }; };
    const wrapper = await mountShell();
    try {
      emit('contextMenu.itemClicked', { itemId: 'otzaria-word-send-to-document', selectedText: 'קטע' });
      release(); await settle(28);
      expect(stub.openSources).toEqual(['loopback://second', 'loopback://first']);
      expect(inserted).toEqual(['loopback://first']);
    } finally { release(); await settle(4); wrapper.unmount(); }
  });
});

describe('פתיחת ספר Word מסרגל הקורא', () => {
  function host(bookReply: () => Promise<unknown>) {
    const listeners = new Map<string, (event: unknown) => void>();
    window.Otzaria = {
      call: (method: string) => method === 'library.openBookFile'
        ? bookReply()
        : Promise.resolve({ success: true, data: {} }),
      on: (name: string, listener: (event: unknown) => void) => listeners.set(name, listener),
      off: (name: string) => listeners.delete(name),
    } as never;
    return () => listeners.get('reader.toolbar_item_clicked')!({
      itemId: 'otzaria-word-edit-book', context: 'reader-text',
      currentBook: 'ראשון', currentBookId: 'ראשון', bookUid: 'user:42',
      currentId: 42, currentType: 'docx', currentSource: 'user',
    });
  }

  const fileReply = () => Promise.resolve({
    success: true, data: { ...TAB_FILES['tok-1'] as object, access: 'readwrite' },
  });

  function storedTabs(activeId = 'doc-1', draft: unknown = null) {
    stub.storedSession = {
      version: 2,
      documents: [
        { id: 'doc-1', document: FIRST_DOC, caret: null, draft },
        { id: 'doc-2', document: SECOND_DOC, caret: null, draft: null },
      ],
      activeId,
      view: { zoom: null, focusMode: false, ribbonTab: null, ribbonCollapsed: false },
    };
    stub.resolvedFile = (token: string) => TAB_FILES[token] ?? null;
  }

  it('לחיצה חוזרת עוברת לטאב הכתיב בלי לפתוח מנוע נוסף', async () => {
    const emit = host(fileReply);
    const wrapper = await mountShell();
    try {
      emit(); await settle(16);
      expect(wrapper.find('.word-statusbar').text()).toContain('„שמור” כותב לקובץ עצמו');
      await wrapper.find('.word-doctabs-new').trigger('click');
      await settle(12);
      const opened = stub.openSources.length;
      emit(); await settle(16);
      expect(wrapper.findAll('.word-doctab')).toHaveLength(2);
      expect(wrapper.find('.word-doctab.active .word-doctab-title').text()).toBe('ראשון');
      expect(stub.openSources).toHaveLength(opened);
    } finally { wrapper.unmount(); window.Otzaria = undefined as never; }
  });

  it('גם ספר שנפתח מ„אחרונים” מזוהה לפי יעד הכתיבה', async () => {
    stub.storedRecents = [{ ...FIRST_DOC, size: 100, openedAt: 1, pinned: false }];
    stub.resolvedFile = TAB_FILES['tok-1'];
    const emit = host(fileReply);
    const wrapper = await mountShell();
    try {
      pressCtrl('KeyO'); await settle(8);
      const { default: OpenDialog } = await import('../../src/ui/panels/OpenDocumentDialog.vue');
      wrapper.findComponent(OpenDialog).vm.$emit('open-recent', 'tok-1');
      await settle(16);
      expect(stub.openSources[stub.openSources.length - 1]).toBe('loopback://first');
      const opened = stub.openSources.length;
      emit(); await settle(16);
      expect(wrapper.findAll('.word-doctab')).toHaveLength(1);
      expect(stub.openSources).toHaveLength(opened);
    } finally { wrapper.unmount(); window.Otzaria = undefined as never; }
  });

  it.each(['doc-1', 'doc-2'])('שחזור טאב כתיב, פעיל או ממתין (%s), אינו יוצר עותק נוסף', async (activeId) => {
    storedTabs(activeId);
    const emit = host(fileReply);
    const wrapper = await mountShell();
    try {
      emit(); await settle(20);
      expect(wrapper.findAll('.word-doctab')).toHaveLength(2);
      expect(wrapper.find('.word-doctab.active .word-doctab-title').text()).toBe('ראשון');
      expect(stub.openSources.filter(source => source === 'loopback://first')).toHaveLength(1);
    } finally { wrapper.unmount(); window.Otzaria = undefined as never; }
  });

  it('טאב ממתין עם שינויים משוחזר מהטיוטה ולא מהעותק שבדיסק', async () => {
    const draft = { path: 'session-draft-doc-1.docx', savedAt: 1, documentToken: 'tok-1', sourceSize: 100 };
    storedTabs('doc-2', draft);
    stub.draftsByPath[draft.path] = new Uint8Array([80, 75, 3, 4]);
    const emit = host(fileReply);
    const wrapper = await mountShell();
    try {
      emit(); await settle(20);
      expect(wrapper.findAll('.word-doctab')).toHaveLength(2);
      expect(stub.openSources[stub.openSources.length - 1]).toBeInstanceOf(Blob);
      expect(stub.openSources).not.toContain('loopback://first');
      expect(stub.markRestoredCalls).toBe(1);
    } finally { wrapper.unmount(); window.Otzaria = undefined as never; }
  });

  it.each(['preflight', 'autoRun'])('התראת %s קודמת להנחיית השמירה', async (warning) => {
    const emit = host(fileReply);
    const wrapper = await mountShell();
    try {
      if (warning === 'preflight') stub.preflight.notice = COMPLEX_SCRIPT_BOLD_NOTICE;
      else stub.preflight.vba = {
        ...NO_VBA, hasMacroPart: true, autoRun: ['AutoOpen'], status: 'התראת מאקרו בפתיחה',
      };
      emit(); await settle(16);
      expect(wrapper.find('.word-statusbar').text()).toContain(
        warning === 'preflight' ? COMPLEX_SCRIPT_BOLD_NOTICE : 'התראת מאקרו בפתיחה',
      );
      expect(wrapper.find('.word-statusbar').text()).not.toContain('„שמור” כותב לקובץ עצמו');
    } finally { wrapper.unmount(); window.Otzaria = undefined as never; }
  });

  it('ספר לקריאה בלבד אינו מקבל יעד כתיבה', async () => {
    const emit = host(async () => ({ success: true, data: { ...TAB_FILES['tok-1'] as object, access: 'read' } }));
    const wrapper = await mountShell();
    try {
      emit(); await settle(16);
      expect(wrapper.find('.word-statusbar').text()).toContain('„שמור” ישמור עותק במקום אחר');
      const record = stub.persistedSessions[stub.persistedSessions.length - 1] as {
        documents: Array<{ document: typeof FIRST_DOC }>;
      };
      expect(record.documents[0]?.document.writable).toBe(false);
    } finally { wrapper.unmount(); window.Otzaria = undefined as never; }
  });

  it('סבב פתירת הספר חוסם פתיחה, סגירה ומעבר טאב ידניים ומשחרר את השער בסיום', async () => {
    storedTabs();
    let release!: () => void;
    const gate = new Promise<void>(resolve => { release = resolve; });
    let started = false;
    const emit = host(async () => { started = true; await gate; return fileReply(); });
    const wrapper = await mountShell();
    try {
      emit(); await settle(8);
      expect(started).toBe(true);
      const opened = stub.openSources.length;
      await wrapper.find('.word-doctabs-new').trigger('click');
      pressCtrl('KeyW');
      await wrapper.findAll('.word-doctab')[1]!.trigger('click');
      await settle(8);
      expect(wrapper.findAll('.word-doctab')).toHaveLength(2);
      expect(wrapper.find('.word-doctab.active .word-doctab-title').text()).toBe('ראשון');
      expect(stub.openSources).toHaveLength(opened);
      release(); await settle(16);
      await wrapper.find('.word-doctabs-new').trigger('click'); await settle(12);
      expect(wrapper.findAll('.word-doctab')).toHaveLength(3);
    } finally { release(); await settle(8); wrapper.unmount(); window.Otzaria = undefined as never; }
  });

  it.each(['rejected', 'throw'])('כשל RPC מסוג %s משחרר את שער הפתיחה', async (failure) => {
    const emit = host(async () => {
      if (failure === 'throw') throw new Error('RPC failed');
      return { success: false, error: { message: 'אין הרשאה' } };
    });
    const wrapper = await mountShell();
    try {
      emit(); await settle(16);
      await wrapper.find('.word-doctabs-new').trigger('click'); await settle(12);
      expect(wrapper.findAll('.word-doctab')).toHaveLength(2);
    } finally { wrapper.unmount(); window.Otzaria = undefined as never; }
  });
});

describe('startup history outside the opening barrier', () => {
  it('opens a stored file before either history reply and merges the early recent entry', async () => {
    let releaseRecents!: (value: unknown) => void;
    let releaseBackups!: (value: unknown) => void;
    stub.storedRecents = new Promise(resolve => { releaseRecents = resolve; });
    stub.storedDiscardBackups = new Promise(resolve => { releaseBackups = resolve; });
    stub.storedSession = {
      version: 2,
      documents: [{ id: 'doc-1', document: FIRST_DOC, caret: null, draft: null }],
      activeId: 'doc-1',
      view: { zoom: null, focusMode: false, ribbonTab: null, ribbonCollapsed: false },
    };
    stub.resolvedFile = TAB_FILES['tok-1'];
    const oldRecents = [
      { ...FIRST_DOC, size: 1, openedAt: 1, pinned: true },
      { token: 'older', name: 'ישן.docx', writable: false, size: 3, openedAt: 2, pinned: false },
    ];
    const wrapper = await mountShell();
    try {
      expect(stub.openSources).toContain('loopback://first');
      expect(stub.persistedRecents).toHaveLength(0);
      pressCtrl('KeyO');
      await settle(8);
      expect(wrapper.text()).toContain('טוען מסמכים אחרונים');
      expect(wrapper.find('.open-browse').attributes('disabled')).toBeUndefined();
      releaseRecents(oldRecents);
      releaseBackups([]);
      await settle(20);
      const saved = stub.persistedRecents[stub.persistedRecents.length - 1] as Array<Record<string, unknown>>;
      expect(saved.map(item => item.token)).toEqual(['tok-1', 'older']);
      expect(saved[0]).toMatchObject({ pinned: true, writable: true, size: 100 });
      expect(saved[0]!.openedAt).toBeGreaterThan(2);
      expect(wrapper.findAll('.rec-open')).toHaveLength(2);
      expect(wrapper.text()).not.toContain('טוען מסמכים אחרונים');
    } finally {
      releaseRecents(oldRecents);
      releaseBackups([]);
      await settle(12);
      wrapper.unmount();
    }
  });
});
