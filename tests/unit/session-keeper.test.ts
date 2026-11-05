/**
 * זוכר ההפעלה.
 *
 * ההבטחה שהקובץ הזה שומר עליה מנוסחת במספר אחד: **לכל היותר דקה של עבודה
 * נמצאת באוויר** — ולא משנה אם המשתמש עוצר להקליד או לא. סביבה: זמן מזויף,
 * כי כל ההתנהגות כאן היא מתי, ולא מה.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  DRAFT_DELAY_MS,
  DRAFT_MAX_WAIT_MS,
  PERSIST_DELAY_MS,
  createSessionKeeper,
  type SessionKeeper,
  type SessionKeeperDeps,
} from '../../src/sessions/session-keeper';
import {
  activeEntry,
  draftSlotPaths,
  emptySession,
  withActiveEntry,
  type SessionDraft,
  type SessionState,
} from '../../src/sessions/session-state';
import type { CaretAnchor } from '../../src/engine/caret-anchor';
import type { WorkspaceWrite } from '../../src/host/workspace';

const anchor: CaretAnchor = { start: { blockId: 'b3', ordinal: 2, offset: 5 }, end: null };

/** שתי המשבצות של המסמך שהבדיקות משתמשות בו. ראו `nextDraftPath`. */
const DRAFT_PATHS = draftSlotPaths('test-doc');
/** המשבצת הראשונה — לאן נכתבת הטיוטה הראשונה של רשומה נקייה. */
const DRAFT_PATH = DRAFT_PATHS[0];

/** הרשומה האחרונה שנכתבה. `at(-1)` אינו ב-lib של הפרויקט. */
function last(records: SessionState[]): SessionState {
  return records[records.length - 1];
}

/** הרשומה הפעילה מתוך הרשומה האחרונה שנכתבה — הצורה החדשה היא אוסף. */
function lastEntry(records: SessionState[]) {
  return activeEntry(last(records));
}

interface Harness {
  keeper: SessionKeeper;
  deps: SessionKeeperDeps;
  persisted: SessionState[];
  drafts: number[];
  /** הנתיבים שנכתבו, לפי הסדר. */
  draftPaths: string[];
  removals: number;
  /** הנתיבים שנמחקו, לפי הסדר. */
  removedPaths: string[];
  /** הטיוטות שהועברו ל„נסגרו בלי לשמור”. */
  rescued: SessionDraft[];
  /** מה ש-`rescueDraft` הבא יחזיר. */
  rescueResult: boolean;
  /** האם פעולה כבדה רצה — `SaveCoordinator.hold`. */
  held: boolean;
  /** האם הייצוא הבא ייכשל. */
  exportFails: boolean;
  exports: number;
  dirty: boolean;
  saving: boolean;
  caret: CaretAnchor | null;
  writeResult: WorkspaceWrite;
  /** כמה פעמים נאמר למשתמש שהמסמך גדול מדי לטיוטה. */
  tooLargeReports: number;
  /** מה שיוחזר מ-`settleSave`; ברירת המחדל היא הבטחה שנפתרת מיד. */
  onSettleSave: (() => Promise<void>) | null;
  /** מריצה את הטיימרים שהבשילו וממתינה לשרשרת האסינכרונית שהם פתחו. */
  tick: (ms: number) => Promise<void>;
}

function harness(): Harness {
  const state = {
    persisted: [] as SessionState[],
    drafts: [] as number[],
    draftPaths: [] as string[],
    removals: 0,
    removedPaths: [] as string[],
    rescued: [] as SessionDraft[],
    rescueResult: true,
    held: false,
    exportFails: false,
    exports: 0,
    dirty: false,
    saving: false,
    caret: null as CaretAnchor | null,
    writeResult: 'written' as WorkspaceWrite,
    tooLargeReports: 0,
    onSettleSave: null as (() => Promise<void>) | null,
  };

  const deps: SessionKeeperDeps = {
    id: 'test-doc',
    persist: async (snapshot) => {
      state.persisted.push(structuredClone(snapshot));
    },
    exportDocument: async () => {
      state.exports += 1;
      if (state.exportFails) throw new Error('הייצוא נכשל');
      return new Blob([new Uint8Array([1, 2, 3])]);
    },
    writeDraft: async (content, path) => {
      if (state.writeResult !== 'written') return state.writeResult;
      state.drafts.push(content.size);
      state.draftPaths.push(path);
      return 'written';
    },
    removeDraft: async (path) => {
      state.removals += 1;
      state.removedPaths.push(path);
    },
    rescueDraft: async (draft) => {
      state.rescued.push(draft);
      return state.rescueResult;
    },
    draftPaths: DRAFT_PATHS,
    readCaret: async () => state.caret,
    isDirty: () => state.dirty,
    isSaving: () => state.saving,
    isHeld: () => state.held,
    settleSave: () => state.onSettleSave?.() ?? Promise.resolve(),
    onDraftTooLarge: () => {
      state.tooLargeReports += 1;
    },
  };

  const keeper = createSessionKeeper(deps);

  return {
    keeper,
    deps,
    get persisted() {
      return state.persisted;
    },
    get drafts() {
      return state.drafts;
    },
    get draftPaths() {
      return state.draftPaths;
    },
    get removals() {
      return state.removals;
    },
    get removedPaths() {
      return state.removedPaths;
    },
    get rescued() {
      return state.rescued;
    },
    get rescueResult() {
      return state.rescueResult;
    },
    set rescueResult(value: boolean) {
      state.rescueResult = value;
    },
    get held() {
      return state.held;
    },
    set held(value: boolean) {
      state.held = value;
    },
    get exportFails() {
      return state.exportFails;
    },
    set exportFails(value: boolean) {
      state.exportFails = value;
    },
    get exports() {
      return state.exports;
    },
    get dirty() {
      return state.dirty;
    },
    set dirty(value: boolean) {
      state.dirty = value;
    },
    get saving() {
      return state.saving;
    },
    set saving(value: boolean) {
      state.saving = value;
    },
    get caret() {
      return state.caret;
    },
    set caret(value: CaretAnchor | null) {
      state.caret = value;
    },
    get writeResult() {
      return state.writeResult;
    },
    set writeResult(value: WorkspaceWrite) {
      state.writeResult = value;
    },
    get tooLargeReports() {
      return state.tooLargeReports;
    },
    get onSettleSave() {
      return state.onSettleSave;
    },
    set onSettleSave(value: (() => Promise<void>) | null) {
      state.onSettleSave = value;
    },
    async tick(ms: number) {
      await vi.advanceTimersByTimeAsync(ms);
      // הטיימר פותח שרשרת של await-ים (ייצוא, כתיבה, כתיבת רשומה); צריך
      // לתת לה להתנקז לפני שבודקים.
      await Promise.resolve();
      await Promise.resolve();
      await Promise.resolve();
    },
  };
}

beforeEach(() => {
  vi.useFakeTimers();
});

afterEach(() => {
  vi.useRealTimers();
});

describe('הרשומה', () => {
  it('נכתבת בהשהיה אחת אחרי סדרת שינויים, ולא אחת לכל שינוי', async () => {
    const h = harness();
    h.caret = anchor;

    h.keeper.noteChange();
    h.keeper.noteChange();
    h.keeper.noteChange();
    expect(h.persisted, 'לפני ההשהיה לא נכתב דבר').toHaveLength(0);

    await h.tick(PERSIST_DELAY_MS);
    expect(h.persisted).toHaveLength(1);
    expect(activeEntry(h.persisted[0])?.caret).toEqual(anchor);
  });

  it('סמן שלא נקרא אינו מוחק את מה שכבר ידענו', async () => {
    // הבחירה עשויה להיות מחוץ למסמך ברגע הכתיבה — מיקוד בשדה בדיאלוג — וזו
    // אינה סיבה לשכוח איפה המשתמש היה.
    const h = harness();
    h.caret = anchor;
    h.keeper.noteChange();
    await h.tick(PERSIST_DELAY_MS);

    h.caret = null;
    h.keeper.noteChange();
    await h.tick(PERSIST_DELAY_MS);

    expect(activeEntry(h.persisted[1])?.caret).toEqual(anchor);
  });

  it('תזוזת סמן כותבת את הרשומה — שם הסמן יושב', async () => {
    const h = harness();
    h.caret = anchor;

    h.keeper.noteCaretMoved();
    await h.tick(PERSIST_DELAY_MS);

    expect(activeEntry(h.persisted[0])?.caret).toEqual(anchor);
  });

  it('תזוזת סמן אינה כותבת טיוטה — הטיוטה היא בייטים של מסמך', async () => {
    // הבחירה מדווחת בכל תו **בנוסף** ל-`onUpdate` של המסמך, וקודם שניהם
    // קראו ל-`noteChange`: מונה השינויים זז גם על תזוזת סמן, ולכן טיוטה
    // נכתבה שוב על מסמך שלא השתנה מאז הקודמת.
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    expect(h.drafts).toHaveLength(1);

    h.keeper.noteCaretMoved();
    expect(h.keeper.hasUnwrittenWork, 'הסמן זז — אין עבודה חדשה').toBe(false);
    await h.tick(DRAFT_DELAY_MS);

    expect(h.drafts, 'הסמן זז ולא המסמך — אין מה לייצא שוב').toHaveLength(1);
  });

  it('מצב תצוגה שלא השתנה אינו מייצר כתיבה', async () => {
    const h = harness();
    h.keeper.updateView({ zoom: 120 });
    await h.tick(PERSIST_DELAY_MS);
    expect(h.persisted).toHaveLength(1);

    h.keeper.updateView({ zoom: 120 });
    await h.tick(PERSIST_DELAY_MS);
    expect(h.persisted, 'אותו ערך — אין מה לכתוב').toHaveLength(1);
  });

  it('מצב התצוגה נאסף מכמה מקורות לרשומה אחת', async () => {
    const h = harness();
    h.keeper.updateView({ zoom: 150 });
    h.keeper.updateView({ focusMode: true });
    h.keeper.updateView({ ribbonTab: 'review', ribbonCollapsed: true });

    await h.tick(PERSIST_DELAY_MS);

    expect(last(h.persisted).view).toEqual({
      zoom: 150,
      focusMode: true,
      ribbonTab: 'review',
      ribbonCollapsed: true,
    });
  });
});

describe('הטיוטה', () => {
  it('נכתבת רק כשיש עבודה שאינה בדיסק', async () => {
    const h = harness();
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    expect(h.exports, 'מסמך נקי — אין מה לייצא').toBe(0);

    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    expect(h.drafts).toEqual([3]);
  });

  it('הרשומה מצביעה לטיוטה שנכתבה, עם המסמך שהיא שייכת לו', async () => {
    const h = harness();
    h.keeper.setDocument({ token: 'tok', name: 'א.docx', writable: true }, { sourceSize: 4_096 });
    h.dirty = true;
    h.keeper.noteChange();

    await h.tick(DRAFT_DELAY_MS);

    expect(lastEntry(h.persisted)?.draft).toMatchObject({
      path: DRAFT_PATH,
      documentToken: 'tok',
      sourceSize: 4_096,
    });
  });

  it('הקלדה רצופה אינה דוחה את הכתיבה לנצח', async () => {
    // זו ההבטחה שהמודול קיים בשבילה: debounce לבדו היה נדחה שוב ושוב, ומי
    // שכותב ברצף עשר דקות היה מגלה שלא נכתב דבר.
    const h = harness();
    h.dirty = true;

    for (let elapsed = 0; elapsed < DRAFT_MAX_WAIT_MS; elapsed += DRAFT_DELAY_MS / 2) {
      h.keeper.noteChange();
      await h.tick(DRAFT_DELAY_MS / 2);
    }

    expect(h.drafts.length, 'נכתבה טיוטה למרות שההקלדה לא נעצרה').toBeGreaterThan(0);
  });

  it('שמירה שרצה דוחה את הייצוא ואינה מבטלת אותו', async () => {
    const h = harness();
    h.dirty = true;
    h.saving = true;
    h.keeper.noteChange();

    await h.tick(DRAFT_DELAY_MS);
    expect(h.exports, 'לא רצים שני ייצואים במקביל על אותו מסמך').toBe(0);

    h.saving = false;
    await h.tick(DRAFT_DELAY_MS);
    expect(h.drafts).toEqual([3]);
  });

  it('כתיבה שנכשלה **משאירה** את הטיוטה הקודמת', async () => {
    // טיוטה שעל הדיסק מאוחרת תמיד לשמירה האחרונה, ולכן היא מחזיקה עבודה
    // שאין בקובץ. „ישנה” היא ביחס למה שעל המסך, לא ביחס לדיסק — ומחיקתה
    // בגלל כתיבה שנכשלה משמידה את העותק היחיד ששרד.
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    const written = lastEntry(h.persisted)?.draft;
    expect(written).not.toBeNull();

    h.writeResult = 'failed';
    h.keeper.noteChange();
    await h.tick(DRAFT_MAX_WAIT_MS);

    expect(h.removals, 'אין למחוק עבודה בגלל כתיבה שנכשלה').toBe(0);
    expect(lastEntry(h.persisted)?.draft).toEqual(written);
  });

  it('כתיבה שנכשלה מנסה שוב בסבב הבא', async () => {
    // `draftedRevision` אינו מתקדם על כישלון, ולכן אותה עבודה עדיין ממתינה.
    const h = harness();
    h.dirty = true;
    h.writeResult = 'failed';
    h.keeper.noteChange();
    await h.tick(DRAFT_MAX_WAIT_MS);
    expect(h.drafts).toHaveLength(0);

    h.writeResult = 'written';
    h.keeper.noteChange();
    await h.tick(DRAFT_MAX_WAIT_MS);
    expect(h.drafts).toHaveLength(1);
  });

  it('מסמך גדול מהמכסה מדווח למשתמש פעם אחת, ולא בכל סבב', async () => {
    // התשובה קבועה — המסמך לא יקטן מעצמו — ולכן חזרה עליה כל עשר שניות היא
    // הטרדה. מסמך אחר מאפס: עליו התשובה אינה ידועה.
    const h = harness();
    h.dirty = true;
    h.writeResult = 'too-large';

    h.keeper.noteChange();
    await h.tick(DRAFT_MAX_WAIT_MS);
    h.keeper.noteChange();
    await h.tick(DRAFT_MAX_WAIT_MS);
    expect(h.tooLargeReports).toBe(1);

    h.keeper.setDocument({ token: 'other', name: 'אחר.docx', writable: true });
    h.keeper.noteChange();
    await h.tick(DRAFT_MAX_WAIT_MS);
    expect(h.tooLargeReports).toBe(2);
  });

  it('כשל חולף אינו מדווח למשתמש', async () => {
    // הסבב הבא עשוי להצליח, ואין מה להטריד בו את מי שרק מקליד.
    const h = harness();
    h.dirty = true;
    h.writeResult = 'failed';

    h.keeper.noteChange();
    await h.tick(DRAFT_MAX_WAIT_MS);

    expect(h.tooLargeReports).toBe(0);
  });

  it('שמירה מוצלחת מעבירה את הטיוטה לקודמת ואינה מוחקת קובץ', async () => {
    // הרגרסיה שהשינוי הזה מונע: עד כאן הטיוטה נמחקה בכל שמירה, ולכן פעולה
    // שהמשתמש לא רצה נכנסה לקובץ בלי שנשאר בשום מקום עותק של „לפני”.
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    const written = lastEntry(h.persisted)?.draft;

    h.dirty = false;
    await h.keeper.noteSaved(h.keeper.captureSave(), { sourceSize: 8_192 });

    expect(h.removals, 'אין קריאת מחיקה — ההעברה היא החלפת מצביע').toBe(0);
    expect(lastEntry(h.persisted)?.draft).toBeNull();
    expect(lastEntry(h.persisted)?.previousDraft).toEqual(written);
  });

  it('הטיוטה שאחרי השמירה נכתבת למשבצת השנייה ואינה דורסת את הקודמת', async () => {
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    h.dirty = false;
    await h.keeper.noteSaved(h.keeper.captureSave(), { sourceSize: 8_192 });
    const previous = lastEntry(h.persisted)?.previousDraft?.path;

    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    expect(h.draftPaths).toEqual([DRAFT_PATHS[0], DRAFT_PATHS[1]]);
    expect(lastEntry(h.persisted)?.previousDraft?.path).toBe(previous);
  });

  it('שתי שמירות רצופות מחזירות את הכתיבה למשבצת הראשונה', async () => {
    const h = harness();
    for (let round = 0; round < 3; round += 1) {
      h.dirty = true;
      h.keeper.noteChange();
      await h.tick(DRAFT_DELAY_MS);
      h.dirty = false;
      await h.keeper.noteSaved(h.keeper.captureSave(), { sourceSize: 8_192 });
    }

    expect(h.draftPaths).toEqual([DRAFT_PATHS[0], DRAFT_PATHS[1], DRAFT_PATHS[0]]);
  });

  it('הקלדה רצופה בלי שמירה כותבת שוב לאותה משבצת', async () => {
    // אין טעם לבזבז את השנייה: שם יושב העותק שהתכונה נבנתה כדי לשמור.
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    h.keeper.noteChange();
    await h.tick(DRAFT_MAX_WAIT_MS);

    expect(h.draftPaths).toEqual([DRAFT_PATHS[0], DRAFT_PATHS[0]]);
  });

  it('שמירה בלי טיוטה משאירה את הקודמת במקומה', async () => {
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    h.dirty = false;
    await h.keeper.noteSaved(h.keeper.captureSave(), { sourceSize: 8_192 });
    const previous = lastEntry(h.persisted)?.previousDraft;

    await h.keeper.noteSaved(h.keeper.captureSave(), { sourceSize: 9_000 });

    expect(lastEntry(h.persisted)?.previousDraft).toEqual(previous);
  });

  it('הקודמת שורדת מעבר מסמך עם keepDraft', async () => {
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true });
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    h.dirty = false;
    await h.keeper.noteSaved(h.keeper.captureSave(), { sourceSize: 8_192 });

    h.keeper.setDocument(null, { keepDraft: true });

    expect(lastEntry(h.persisted)?.previousDraft?.documentToken).toBe('a');
  });

  it('מעבר למסמך אחר **אינו** מוחק את הטיוטה', async () => {
    // מחיקה כאן נראתה סבירה, ובפועל השמידה עבודה בשלושה מסלולים שאיש לא
    // התכוון אליהם — פתיחה שנכשלה, טיוטה של מסמך שיחזרו אליו, ותשובת „לא”
    // לשאלה על קובץ שהשתנה. המצביע נושא את ה-token של הבעלים, ולכן הוא
    // מזדהה בעצמו (`decideDraftRecovery`) ואינו זקוק לניקוי מונע.
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true });
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    h.keeper.setDocument({ token: 'b', name: 'ב.docx', writable: true });
    await h.tick(0);

    expect(h.removals).toBe(0);
    expect(lastEntry(h.persisted)?.draft, 'המצביע נשאר, ועדיין מסומן בבעליו').toMatchObject({
      documentToken: 'a',
    });
  });

  it('„למחוק את השינויים” הוא המסלול היחיד שמוחק', async () => {
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true });
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    await h.keeper.discardDraft();

    // שתי המשבצות ולא רק מה שברשומה: קובץ שנכתב ורגע לפני שנרשם נקטע היה
    // נשאר יתום לנצח, ואין ניקוי יתומים בעלייה.
    expect(new Set(h.removedPaths)).toEqual(new Set(DRAFT_PATHS));
    expect(lastEntry(h.persisted)?.draft).toBeNull();
    expect(lastEntry(h.persisted)?.previousDraft).toBeNull();
  });

  it('„למחוק” מוחק גם את הקודמת ואת נתיב הרשומה מגרסה ישנה', async () => {
    const h = harness();
    const legacy = 'session-draft-test-doc.docx';
    h.keeper.adopt(
      withActiveEntry(emptySession(), {
        document: { token: 'a', name: 'א.docx', writable: true },
        draft: { path: legacy, savedAt: 1, documentToken: 'a', sourceSize: 10 },
      }),
    );

    await h.keeper.discardDraft();

    expect(new Set(h.removedPaths)).toEqual(new Set([legacy, ...DRAFT_PATHS]));
  });

  it('„למחוק” בלי דבר ברשומה אינו נוגע בגשר', async () => {
    const h = harness();
    await h.keeper.discardDraft();

    expect(h.removals).toBe(0);
  });

  it('גודל הקובץ מהשמירה מגיע לטיוטה הבאה', async () => {
    // בלעדיו הרשומה נשארת עם הגודל שהיה בפתיחה, ו„הקובץ השתנה מבחוץ” נשאל
    // אחרי כל שמירה רגילה — כלומר בדיוק כשהוא שגוי.
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true }, { sourceSize: 1_000 });
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    h.dirty = false;
    await h.keeper.noteSaved(h.keeper.captureSave(), { sourceSize: 1_500 });

    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    expect(lastEntry(h.persisted)?.draft).toMatchObject({ sourceSize: 1_500 });
  });

  it('שמירה שהסתיימה בזמן כתיבת טיוטה אינה משאירה מצביע לקובץ שנמחק', async () => {
    // `noteSaved` קורא את המצב בתוך התור ולא לפניו: לפניו הוא תיאר את הרגע
    // שבו השמירה הסתיימה, בעוד המחיקה רצה אחרי סבב טיוטה שכבר המתין.
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true });
    h.dirty = true;
    h.keeper.noteChange();

    const writing = h.tick(DRAFT_DELAY_MS);
    const saving = h.keeper.noteSaved(h.keeper.captureSave(), { sourceSize: 2_000 });
    await Promise.all([writing, saving]);
    await h.tick(0);

    expect(lastEntry(h.persisted)?.draft, 'הרשומה והדיסק חייבים להסכים').toBeNull();
  });

  it('טיוטה שנשמרת עוברת לבעלות המסמך שנפתח ממנה', async () => {
    // המסלול: הקובץ לא נמצא, והעבודה שלא נשמרה נפתחה כמסמך חדש. בלי ההעברה
    // הרשומה הייתה ממשיכה לטעון שהטיוטה שייכת ל-token שאינו נפתר, והפתיחה
    // הבאה הייתה פוסלת אותה.
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true }, { sourceSize: 500 });
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    expect(lastEntry(h.persisted)?.draft).toMatchObject({ documentToken: 'a', sourceSize: 500 });

    h.keeper.setDocument(null, { keepDraft: true });
    await h.tick(0);

    expect(lastEntry(h.persisted)?.draft).toMatchObject({ documentToken: null, sourceSize: null });
  });

  it('פתיחה **מתוך** הטיוטה משאירה אותה על הדיסק', async () => {
    // המסלול היחיד שבו המחיקה הרגילה הייתה משמידה עבודה: הטיוטה היא בדיוק
    // מה שעל המסך, ובחלון שבין המחיקה לכתיבה הבאה קריסה מוחקת הכול.
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true });
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true }, { keepDraft: true });
    await h.tick(0);

    expect(h.removals).toBe(0);
    expect(lastEntry(h.persisted)?.draft).not.toBeNull();
  });
});

describe('הסמן מול החלפת מסמך', () => {
  it('נמחק כשהמסמך התחלף', async () => {
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true });
    h.caret = anchor;
    h.keeper.noteChange();
    await h.tick(PERSIST_DELAY_MS);

    h.keeper.setDocument({ token: 'b', name: 'ב.docx', writable: true });
    await h.tick(0);

    expect(lastEntry(h.persisted)?.caret, 'סמן של מסמך אחד ברשומה של אחר').toBeNull();
  });

  it('נשמר כשאותו מסמך נפתח מחדש', async () => {
    // זה בדיוק המסלול שהתכונה נכתבה בשבילו; מחיקה תמידית הייתה מאבדת אותו.
    const h = harness();
    h.keeper.adopt(
      withActiveEntry(emptySession(), {
        document: { token: 'a', name: 'א.docx', writable: true },
        caret: anchor,
      }),
    );
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true });
    await h.tick(0);

    expect(lastEntry(h.persisted)?.caret).toEqual(anchor);
  });
});

describe('נקודת השמירה', () => {
  it('נקודה של זוכר אחר אינה משנה את היעד או צורכת את הנקודה אצל בעליה', async () => {
    const a = harness(); const b = harness();
    a.keeper.setDocument({ token: 'a', name: 'A.docx', writable: true });
    b.keeper.setDocument({ token: 'b', name: 'B.docx', writable: true });
    a.dirty = true; a.keeper.noteChange();
    await a.tick(DRAFT_DELAY_MS);
    const point = a.keeper.captureSave();
    const before = b.keeper.state;

    await b.keeper.noteSaved(point, {
      sourceSize: 999,
      document: { token: 'foreign', name: 'Foreign.docx', writable: true },
    });
    expect(b.keeper.state).toBe(before);
    a.dirty = false;
    await a.keeper.noteSaved(point, { sourceSize: 100 });
    expect(activeEntry(a.keeper.state)?.draft).toBeNull();
    expect(activeEntry(a.keeper.state)?.previousDraft).not.toBeNull();
  });

  it('נקודה ממסמך קודם אינה משנה את היעד, הגודל או תזמון הטיוטה החדשה', async () => {
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'A.docx', writable: true }, { sourceSize: 100 });
    h.dirty = true; h.keeper.noteChange();
    const point = h.keeper.captureSave();
    h.keeper.setDocument({ token: 'b', name: 'B.docx', writable: true }, { sourceSize: 200 });
    h.keeper.noteChange();

    await h.keeper.noteSaved(point, {
      sourceSize: 999,
      document: { token: 'stale', name: 'Stale.docx', writable: true },
    });
    await h.tick(DRAFT_DELAY_MS);
    expect(activeEntry(h.keeper.state)?.document?.token).toBe('b');
    expect(activeEntry(h.keeper.state)?.draft).toMatchObject({ documentToken: 'b', sourceSize: 200 });
    expect(h.keeper.hasUnwrittenWork).toBe(false);
  });

  it.each(['replace', 'dispose'] as const)('נקודה שהתיישנה בזמן ההמתנה בתור אינה נרשמת (%s)', async (operation) => {
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'A.docx', writable: true });
    await h.tick(0);
    const point = h.keeper.captureSave();
    let release!: (value: CaretAnchor | null) => void;
    const pending = new Promise<CaretAnchor | null>(resolve => { release = resolve; });
    h.deps.readCaret = vi.fn().mockReturnValueOnce(pending).mockResolvedValue(null);
    h.keeper.noteCaretMoved();
    await h.tick(PERSIST_DELAY_MS);
    expect(h.deps.readCaret).toHaveBeenCalledOnce();
    const saved = h.keeper.noteSaved(point, {
      sourceSize: 999,
      document: { token: 'stale', name: 'Stale.docx', writable: true },
    });
    if (operation === 'replace') {
      h.keeper.setDocument({ token: 'b', name: 'B.docx', writable: true });
    } else h.keeper.dispose();
    const expected = h.keeper.state;
    release(null); await saved; await h.tick(0);
    expect(h.keeper.state).toBe(expected);
    expect(h.persisted.every(state => activeEntry(state)?.document?.token !== 'stale')).toBe(true);
  });

  it('נקודה שכבר נרשמה אינה ניתנת להפעלה חוזרת אחרי עריכה חדשה', async () => {
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'A.docx', writable: true });
    const point = h.keeper.captureSave();
    await h.keeper.noteSaved(point, { sourceSize: 100 });
    h.dirty = true; h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    const before = h.keeper.state;
    await h.keeper.noteSaved(point, {
      sourceSize: 999,
      document: { token: 'replayed', name: 'Replayed.docx', writable: true },
    });
    expect(h.keeper.state).toBe(before);
    expect(activeEntry(h.keeper.state)?.draft).toMatchObject({ documentToken: 'a', sourceSize: 100 });
  });

  it('כשל טיוטה מאוחרת שומר את העותק הקודם ומאפשר ניסיון חוזר', async () => {
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'A.docx', writable: true });
    h.dirty = true; h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    const oldDraft = activeEntry(h.keeper.state)?.draft;
    const point = h.keeper.captureSave();
    h.keeper.noteChange();
    await h.keeper.noteSaved(point, { sourceSize: 200 });
    h.writeResult = 'failed';
    await h.tick(DRAFT_DELAY_MS);

    expect(activeEntry(h.keeper.state)?.previousDraft).toEqual(oldDraft);
    expect(activeEntry(h.keeper.state)?.draft).toBeNull();
    expect(h.keeper.hasUnwrittenWork).toBe(true);
    h.writeResult = 'written';
    await h.keeper.flush();
    expect(activeEntry(h.keeper.state)?.previousDraft).toEqual(oldDraft);
    expect(activeEntry(h.keeper.state)?.draft).toMatchObject({ documentToken: 'a', sourceSize: 200 });
    expect(h.keeper.hasUnwrittenWork).toBe(false);
    expect(h.draftPaths).toEqual([DRAFT_PATHS[0], DRAFT_PATHS[1]]);
  });

  it('שמירת snapshot ישן אינה מאריכה את תקרת ההמתנה של הקלדה רצופה', async () => {
    const h = harness();
    h.dirty = true; h.keeper.noteChange();
    const point = h.keeper.captureSave();
    for (let edit = 0; edit < 6; edit += 1) {
      await h.tick(9_000);
      h.keeper.noteChange();
    }
    await h.tick(1_000);
    await h.keeper.noteSaved(point, { sourceSize: 100 });
    await h.tick(DRAFT_MAX_WAIT_MS - 55_000 - 1);
    expect(h.exports).toBe(0);
    await h.tick(1);
    expect(h.exports).toBe(1);
    expect(activeEntry(h.keeper.state)?.draft).not.toBeNull();
  });
});

describe('flush', () => {
  it('כותב מיד, בלי להמתין להשהיה', async () => {
    const h = harness();
    h.caret = anchor;
    h.keeper.noteChange();

    await h.keeper.flush();

    expect(h.persisted).toHaveLength(1);
  });

  it('כותב גם את הטיוטה כשיש עבודה שאינה בדיסק', async () => {
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();

    await h.keeper.flush();

    expect(h.drafts).toEqual([3]);
  });

  it('אידמפוטנטי — שלושת מקורות היציאה עשויים לירות יחד', async () => {
    // ניווט באוצריא מייצר גם `plugin.suspended` וגם `visibilitychange`, ולעיתים
    // גם `pagehide`. בלי הקיזוז, העזיבה הייתה עולה שלושה ייצואים מלאים של
    // אותו מסמך בדיוק.
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();

    await Promise.all([h.keeper.flush(), h.keeper.flush(), h.keeper.flush()]);

    expect(h.drafts.length, 'ייצוא אחד לכל יציאה, לא שלושה').toBe(1);
  });

  it('טיוטה נכתבת שוב אחרי שינוי נוסף', async () => {
    // הקיזוז שלמעלה אינו רשאי להפוך ל„נכתבה פעם אחת, ודי”.
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();
    await h.keeper.flush();

    h.keeper.noteChange();
    await h.keeper.flush();

    expect(h.drafts).toHaveLength(2);
  });

  it('בזמן שמירה ממתין לה, ואז כותב את מה שהשמירה לא הספיקה', async () => {
    // ביציאה אין „סבב הבא”. דחייה כאן הייתה תולה את כל העבודה שמאז הטיוטה
    // האחרונה בהצלחת השמירה שרצה — וכשל שלה היה משאיר אותה בלי עותק.
    const h = harness();
    h.dirty = true;
    h.saving = true;
    h.keeper.noteChange();

    // השמירה נכשלה: המסמך נשאר מלוכלך גם אחריה.
    h.onSettleSave = async () => {
      h.saving = false;
    };

    await h.keeper.flush();

    expect(h.drafts, 'הטיוטה נכתבת אחרי שהשמירה הסתיימה').toEqual([3]);
  });

  it('בזמן שמירה שהצליחה אינו מייצא לחינם', async () => {
    // `noteSaved` כבר ניקה, והמסמך נקי — אין עבודה שאינה בדיסק, ולכן אין מה
    // לכתוב. ייצוא כאן היה עבודה מיותרת בדיוק ברגע הרגיש ביותר.
    const h = harness();
    h.dirty = true;
    h.saving = true;
    h.keeper.noteChange();

    h.onSettleSave = async () => {
      h.saving = false;
      h.dirty = false;
    };

    await h.keeper.flush();

    expect(h.exports).toBe(0);
    expect(h.drafts).toHaveLength(0);
  });

  it('אחרי dispose אינו כותב', async () => {
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();
    h.keeper.dispose();

    await h.keeper.flush();
    await h.tick(DRAFT_MAX_WAIT_MS);

    expect(h.persisted).toHaveLength(0);
    expect(h.drafts).toHaveLength(0);
  });
});

/**
 * `hasUnwrittenWork` — השער שמאפשר ל„טאב נרדם” להיות בטוח.
 *
 * ## למה הבדיקות האלה קיימות
 *
 * שחרור המנוע של טאב ברקע (App.vue, `sleepTab`) מוחק מהזיכרון את כל מה
 * שהמסמך מחזיק. מה שמתיר אותו הוא התשובה `false` כאן — ההבטחה שכל מה שנמחק
 * ניתן לקריאה בחזרה מהטיוטה או מהדיסק. ביקורת מדדה שמוטציה שמוחקת את השער
 * הזה מ-`sleepTab` עוברת בירוק את כל בדיקות המעטפת, מפני שהכפיל של רכז
 * השמירה שם לעולם אינו „מלוכלך”. לכן הן כאן, על הדבר עצמו.
 *
 * שלושת המסלולים הם שלוש התשובות האפשריות של כתיבת הטיוטה: נכתבה, לא נכתבה,
 * ונכתבה אך העבודה המשיכה.
 */
describe('בייטים שנפתחו בלי קובץ מקור', () => {
  it('כותב ומאמץ את המקור גם כשהשחזור מקפיא שמירה ובלי לייצא מחדש', async () => {
    const h = harness();
    h.dirty = true;
    h.held = true;
    h.keeper.setDocument(null, { keepDraft: true });
    expect(await h.keeper.storeOpenedDraft(new Blob(['received bytes']), 'מקור.docm')).toBe('written');
    expect(h.drafts).toEqual([14]);
    expect(lastEntry(h.persisted)?.draft).toMatchObject({ path: DRAFT_PATH, name: 'מקור.docm', documentToken: null });
    expect(h.keeper.hasUnwrittenWork).toBe(false);
    await h.keeper.flush();
    expect(h.exports).toBe(0);
    expect(h.drafts).toHaveLength(1);
  });

  it.each(['failed', 'too-large'] as const)('כשל %s אינו מרשה לפרק עבודה שאינה מגובה', async (result) => {
    const h = harness();
    h.dirty = true;
    h.held = true;
    h.writeResult = result;
    expect(await h.keeper.storeOpenedDraft(new Blob(['bytes']), 'שם.docx')).toBe(result);
    expect(activeEntry(h.keeper.state)?.draft).toBeNull();
    expect(h.keeper.hasUnwrittenWork).toBe(true);
  });

  it('עריכה במהלך כתיבת המקור נשארת ממתינה לייצוא ושומרת את השם', async () => {
    const h = harness();
    h.dirty = true;
    let finish!: (value: WorkspaceWrite) => void;
    h.deps.writeDraft = () => new Promise((resolve) => { finish = resolve; });
    const writing = h.keeper.storeOpenedDraft(new Blob(['bytes']), 'שם.docx');
    await Promise.resolve();
    h.keeper.noteChange();
    finish('written');
    await writing;
    expect(h.keeper.hasUnwrittenWork).toBe(true);
    h.deps.writeDraft = async () => 'written';
    await h.keeper.flush();
    expect(h.keeper.hasUnwrittenWork).toBe(false);
    expect(activeEntry(h.keeper.state)?.draft?.name).toBe('שם.docx');
  });

  it('עריכה שקדמה לכתיבת המקור אינה נחשבת כאילו נכללה בבייטים שהתקבלו', async () => {
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();
    expect(await h.keeper.storeOpenedDraft(new Blob(['original']), 'שם.docx')).toBe('written');
    expect(h.keeper.hasUnwrittenWork).toBe(true);
    await h.keeper.flush();
    expect(h.exports).toBe(1);
    expect(h.keeper.hasUnwrittenWork).toBe(false);
  });

  it('כשל כתיבת הרשומה אינו מדווח הצלחה ואינו מתיר הרדמה', async () => {
    const h = harness();
    h.dirty = true;
    h.deps.persist = async () => { throw new Error('storage unavailable'); };
    expect(await h.keeper.storeOpenedDraft(new Blob(['bytes']), 'שם.docx')).toBe('failed');
    expect(h.keeper.hasUnwrittenWork).toBe(true);
    expect(activeEntry(h.keeper.state)?.draft?.name).toBe('שם.docx');
  });

  it('כתיבה שחזרה אחרי החלפת מסמך אינה מאומצת על המסמך החדש', async () => {
    const h = harness();
    let finish!: (value: WorkspaceWrite) => void;
    h.deps.writeDraft = () => new Promise((resolve) => { finish = resolve; });
    const writing = h.keeper.storeOpenedDraft(new Blob(['bytes']), 'שם.docx');
    await Promise.resolve();
    h.keeper.setDocument({ token: 'other', name: 'אחר.docx', writable: true });
    finish('written');
    expect(await writing).toBe('failed');
    expect(activeEntry(h.keeper.state)?.draft).toBeNull();
  });
});

describe('hasUnwrittenWork', () => {
  it('מסמך נקי — אין עבודה באוויר', () => {
    const h = harness();

    expect(h.keeper.hasUnwrittenWork).toBe(false);
  });

  it('עריכה שטרם נכתבה לטיוטה — יש', async () => {
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();

    expect(h.keeper.hasUnwrittenWork, 'ההשהיה עוד לא הבשילה').toBe(true);
  });

  it('אחרי `flush` שכתב את הטיוטה — אין', async () => {
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();

    await h.keeper.flush();

    expect(h.drafts.length, 'הטיוטה אכן נכתבה').toBe(1);
    expect(h.keeper.hasUnwrittenWork).toBe(false);
  });

  it('טיוטה שלא נכתבה — נשאר „יש”, וזה מה שמונע שחרור מנוע', async () => {
    // המסלול שמציל מסמך גדול מהמכסה: הוא לעולם לא יירדם, וישלם בזיכרון
    // במקום בעבודה. אותה תשובה בדיוק גם לכשל גשר (`failed`).
    for (const result of ['too-large', 'failed'] as const) {
      const h = harness();
      h.writeResult = result;
      h.dirty = true;
      h.keeper.noteChange();

      await h.keeper.flush();

      expect(h.keeper.hasUnwrittenWork, result).toBe(true);
    }
  });

  it('עריכה נוספת אחרי טיוטה מוצלחת — חוזר ל„יש”', async () => {
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();
    await h.keeper.flush();

    h.keeper.noteChange();

    expect(h.keeper.hasUnwrittenWork).toBe(true);
  });

  it('מסמך שנשמר לדיסק — אין, גם אם משהו זז מאז', async () => {
    // שני התנאים ולא אחד: מונה השינויים לבדו זז גם מתזוזת סמן במסמך שנשמר.
    const h = harness();
    h.dirty = false;
    h.keeper.noteChange();

    expect(h.keeper.hasUnwrittenWork).toBe(false);
  });
});

describe('הקפאה בזמן פעולה כבדה', () => {
  it('התקרה אינה יורה ייצוא בזמן הקפאה', async () => {
    // התקרה אינה נדחית על ידי הקלדה, ולכן בלי השער הזה ייצוא מלא היה יורה
    // באמצע כלי שרץ דקה על כל המסמך — בדיוק על מכונה חלשה, ובדיוק ברגע
    // שהמנוע כבר עמוס.
    const h = harness();
    h.dirty = true;
    h.held = true;

    h.keeper.noteChange();
    await h.tick(DRAFT_MAX_WAIT_MS * 2);

    expect(h.exports).toBe(0);
  });

  it('אחרי השחרור הטיוטה נכתבת בסבב הבא ואינה מבוטלת', async () => {
    const h = harness();
    h.dirty = true;
    h.held = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_MAX_WAIT_MS);

    h.held = false;
    await h.tick(DRAFT_DELAY_MS);

    expect(h.drafts).toHaveLength(1);
  });

  it('flush בזמן הקפאה אינו דורס את הצ׳קפוינט', async () => {
    // מסלול אמיתי: מעבר ללשונית אחרת באוצריא באמצע כלי מפעיל את `flush`,
    // והוא היה מייצא מסמך חצי-מוחל לאותה משבצת שבה יושב העותק שלפני הכלי.
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    expect(h.drafts).toHaveLength(1);

    h.held = true;
    h.keeper.noteChange();
    await h.keeper.flush();

    expect(h.drafts, 'הייצוא נעצר').toHaveLength(1);
    expect(h.persisted.length, 'הרשומה עצמה כן נכתבה').toBeGreaterThan(1);
  });

  it('flush בזמן הקפאה משאיר עבודה לא-כתובה, כדי שהטאב לא יירדם', async () => {
    const h = harness();
    h.dirty = true;
    h.held = true;
    h.keeper.noteChange();

    await h.keeper.flush();

    expect(h.keeper.hasUnwrittenWork).toBe(true);
  });
});

describe('checkpoint', () => {
  it('מסמך נקי נכתב כקודמת ולא כטיוטה', async () => {
    // התרחיש שבשבילו זה נבנה: „שמור” ואז כלי כבד. המסמך נקי בדיוק ברגע
    // שהכלי מתחיל, ולכן `writeDraftNow` הרגיל היה חוזר בלי לכתוב דבר.
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true });
    h.dirty = false;

    await h.keeper.checkpoint();

    expect(h.drafts).toHaveLength(1);
    expect(lastEntry(h.persisted)?.draft, 'מסמך נקי אינו „עבודה לא-שמורה”').toBeNull();
    expect(lastEntry(h.persisted)?.previousDraft?.path).toBe(DRAFT_PATHS[0]);
  });

  it('מסמך מלוכלך נכתב כטיוטה', async () => {
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true });
    h.dirty = true;
    h.keeper.noteChange();

    await h.keeper.checkpoint();

    expect(h.drafts).toHaveLength(1);
    expect(lastEntry(h.persisted)?.draft?.path).toBe(DRAFT_PATHS[0]);
    expect(h.keeper.hasUnwrittenWork).toBe(false);
  });

  it('מלוכלך שכבר יש לו טיוטה עדכנית אינו מייצא שוב', async () => {
    const h = harness();
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    await h.keeper.checkpoint();

    expect(h.exports).toBe(1);
  });

  it('כותב גם בזמן הקפאה — הקורא מקפיא לפניו בכוונה', async () => {
    const h = harness();
    h.dirty = false;
    h.held = true;

    await h.keeper.checkpoint();

    expect(h.drafts).toHaveLength(1);
  });

  it('ממתין לשמירה שרצה ואז כותב', async () => {
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true });
    h.dirty = false;
    h.saving = true;
    h.onSettleSave = async () => {
      h.saving = false;
    };

    await h.keeper.checkpoint();

    expect(h.drafts).toHaveLength(1);
  });

  it('התרחיש המלא: שמור ← צ׳קפוינט ← כלי ← שמירה — והקודמת היא הצ׳קפוינט', async () => {
    const h = harness();
    h.keeper.setDocument({ token: 'a', name: 'א.docx', writable: true });

    // המסמך נשמר, ולכן אין טיוטה.
    h.dirty = false;
    await h.keeper.noteSaved(h.keeper.captureSave(), { sourceSize: 1_000 });

    // לפני הכלי — עותק.
    await h.keeper.checkpoint();
    const checkpoint = lastEntry(h.persisted)?.previousDraft?.path;
    expect(checkpoint).toBe(DRAFT_PATHS[0]);

    // הכלי רץ, והשמירה האוטומטית כותבת את תוצאתו לקובץ.
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    h.dirty = false;
    await h.keeper.noteSaved(h.keeper.captureSave(), { sourceSize: 2_000 });

    // וזו הנקודה: אחרי שהתוצאה הרעה בקובץ, עדיין יש לאן לחזור.
    expect(lastEntry(h.persisted)?.previousDraft?.path).toBe(DRAFT_PATHS[1]);
    expect(lastEntry(h.persisted)?.draft).toBeNull();
  });

  it('כשל ייצוא אינו נזרק לקורא', async () => {
    // רשת ביטחון, לא תנאי: מי שביקש להריץ כלי יריץ אותו גם בלעדיה.
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const h = harness();
    h.exportFails = true;
    h.dirty = false;

    await expect(h.keeper.checkpoint()).resolves.toBeUndefined();
    warn.mockRestore();
  });
});

describe('טיוטה שהמנוע לא נטען ממנה', () => {
  // הרשומה שהגיעה מהשטח אחרי שדרוג ל-7.0: מסמך חדש שמעולם לא נשמר, והעבודה
  // כולה בטיוטה. השחזור לא הביא אותה, הטאב נפתח ריק — והקלדה ראשונה בו
  // הייתה כותבת מסמך ריק בדיוק על המשבצת שלה.
  const unsaved: SessionDraft = { path: DRAFT_PATH, savedAt: 1, documentToken: null, sourceSize: null };

  function restoredBlank(): Harness {
    const h = harness();
    h.keeper.adopt(withActiveEntry(emptySession(), { document: null, draft: unsaved }));
    h.keeper.setDocument(null);
    return h;
  }

  it('מועברת לגיבוי לפני הכתיבה הראשונה', async () => {
    const h = restoredBlank();
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    expect(h.rescued).toEqual([unsaved]);
    expect(h.drafts).toHaveLength(1);
    // והכתיבות הבאות כבר אינן שואלות: הטיוטה עכשיו של מה שבמנוע.
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    expect(h.rescued).toHaveLength(1);
  });

  it('כשההעברה נכשלה, אף כתיבה אינה נוחתת עליה — גם לא השנייה, וגם אחרי שמירה', async () => {
    const h = harness();
    // והרשומה מחזיקה גם „קודמת” במשבצת השנייה: היא זו שמוותרים עליה.
    h.keeper.adopt(withActiveEntry(emptySession(), {
      document: null,
      draft: unsaved,
      previousDraft: { path: DRAFT_PATHS[1], savedAt: 0, documentToken: null, sourceSize: null },
    }));
    h.keeper.setDocument(null);
    h.rescueResult = false;
    h.dirty = true;
    for (let i = 0; i < 3; i++) {
      h.keeper.noteChange();
      await h.tick(DRAFT_DELAY_MS);
    }
    const point = h.keeper.captureSave();
    h.dirty = false;
    await h.keeper.noteSaved(point);
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    expect(h.draftPaths.length).toBeGreaterThan(1);
    expect(h.draftPaths).not.toContain(DRAFT_PATH);
    // וכל הזדמנות מנסה להעביר אותה שוב.
    expect(h.rescued.length).toBeGreaterThan(1);
  });

  it('הצלחה מאוחרת משחררת את המשבצת', async () => {
    const h = restoredBlank();
    h.rescueResult = false;
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    h.rescueResult = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    expect(h.rescued).toHaveLength(2);
    expect(h.draftPaths).toEqual([DRAFT_PATHS[1], DRAFT_PATHS[1], DRAFT_PATHS[1]]);
  });

  it('מסמך ריק שנפתח בלי setDocument (פתיחה שנכשלה) — הטיוטה שנכתבה קודם אינה שלו', async () => {
    // טאב שנרדם: הטיוטה נכתבה ממנו עצמו, ואז ההתעוררות נכשלה ונפתח ריק
    // ב-`remember: false` — בלי `adopt` ובלי `setDocument`.
    const h = harness();
    h.keeper.setDocument(null);
    h.dirty = true;
    h.keeper.noteChange();
    await h.keeper.flush();
    expect(h.draftPaths).toEqual([DRAFT_PATH]);

    h.keeper.noteDraftNotLoaded();
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    expect(h.rescued.map((draft) => draft.path)).toEqual([DRAFT_PATH]);
  });

  it('„לא לשמור” בזמן שכתיבת טיוטה באוויר — הכתיבה אינה דורסת את מה שעוד לא הועבר', async () => {
    const h = restoredBlank();
    h.dirty = true;
    let release!: () => void;
    const exported = new Promise<void>((resolve) => { release = resolve; });
    const original = h.deps.exportDocument;
    h.deps.exportDocument = async () => { await exported; return original(); };
    h.keeper.noteChange();
    await vi.advanceTimersByTimeAsync(DRAFT_DELAY_MS);

    const discarding = h.keeper.discardDraft();
    release();
    await discarding;
    await h.tick(0);

    expect(h.draftPaths).toEqual([]);
    expect(h.rescued).toEqual([unsaved]);
  });

  it('טיוטה שנפתחה ממנה אינה מועברת, ונכתבת באותה משבצת', async () => {
    const h = harness();
    h.keeper.adopt(withActiveEntry(emptySession(), { document: null, draft: unsaved }));
    h.keeper.setDocument(null, { keepDraft: true });
    h.dirty = true;
    h.keeper.noteChange();
    await h.tick(DRAFT_DELAY_MS);

    expect(h.rescued).toEqual([]);
    expect(h.draftPaths).toEqual([DRAFT_PATH]);
  });

  it('„לא לשמור” על מנוע מלוכלך מעביר אותה לפני המחיקה', async () => {
    const h = restoredBlank();
    h.dirty = true;
    await h.keeper.discardDraft();

    expect(h.rescued).toEqual([unsaved]);
  });

  it('„לא לשמור” שההעברה בו נכשלה אינו מוחק אותה', async () => {
    const h = restoredBlank();
    h.rescueResult = false;
    h.dirty = true;
    await h.keeper.discardDraft();

    expect(h.removedPaths).not.toContain(DRAFT_PATH);
  });

  it('„לא לשמור” על מנוע נקי אינו מעביר — הסוגר כבר גיבה מהטיוטה עצמה', async () => {
    const h = restoredBlank();
    await h.keeper.discardDraft();

    expect(h.rescued).toEqual([]);
  });

  it('שמירה של מה שבמנוע מעבירה אותה במקום להפוך אותה ל„קודמת”', async () => {
    const h = restoredBlank();
    h.dirty = true;
    h.keeper.noteChange();
    const point = h.keeper.captureSave();
    h.dirty = false;
    await h.keeper.noteSaved(point, { document: { token: 'n', name: 'חדש.docx', writable: true } });

    expect(h.rescued).toEqual([unsaved]);
    expect(activeEntry(h.keeper.state)?.previousDraft).toBeNull();
  });
});
