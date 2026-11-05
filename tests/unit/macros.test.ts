/**
 * מערכת המאקרו — ההתקנה, מתגי ההקלטה והפירוק.
 *
 * המנוע מוחלף בכפיל מבני (אותה תבנית כמו בדיקות האדפטרים): מה שנבדק כאן הוא
 * החיווט שלנו — לא superdoc-macros עצמה, שמגיעה עם חבילת הבדיקות שלה.
 * הקיצורים (Ctrl+Alt+R וכו') אינם נבדקים כאן: הם רשומות רג'יסטרי שמנותבות
 * ב-App.vue, ובדיקות הרג'יסטרי מכסות אותן.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { SuperDoc } from 'superdoc';
import type { EditorSession } from '../../src/engine/create-editor';
import { createEditorInteraction } from '../../src/engine/editor-interaction';
import {
  MACRO_STATUS,
  SCRIPTS_FLAG_KEY,
  installMacros,
  recordingSavedText,
  type MacrosSession,
} from '../../src/engine/macros';

interface FakeEngine {
  ui: {
    commands: {
      has(id: string): boolean;
      get(id: string): { getState(): { reason?: string } };
      executeAsync(id: string, payload?: unknown): Promise<unknown>;
    };
  };
  activeEditor: null;
}

function createFakeSession(): { session: MacrosSession; engine: FakeEngine } {
  const engine: FakeEngine = {
    ui: {
      commands: {
        has: (id) => id === 'bold',
        get: () => ({ getState: () => ({}) }),
        executeAsync: async () => ({ success: true }),
      },
    },
    activeEditor: null,
  };
  // כפיל מבני של המופע: ההתקנה צורכת רק את המשטחים שלמעלה.
  return { session: { superdoc: engine as unknown as SuperDoc }, engine };
}

function typeChar(container: HTMLElement, char: string): void {
  container.dispatchEvent(
    new InputEvent('beforeinput', { inputType: 'insertText', data: char, bubbles: true }),
  );
}

describe('installMacros', () => {
  beforeEach(() => {
    // המאקרו נשמרים ב-localStorage; בלי ניקוי בדיקה אחת מזהמת את הבאה.
    localStorage.clear();
  });

  it('an in-flight replay cannot change document-mode through the raw controller after opening starts', async () => {
    const { session, engine } = createFakeSession();
    engine.ui.commands.has = () => true;
    const container = document.createElement('div');
    let finishFirst!: () => void;
    const first = new Promise<void>(resolve => { finishFirst = resolve; });
    const run = vi.fn(async () => { await first; return { ok: true as const }; });
    const interaction = createEditorInteraction({
      superdoc: { setDocumentMode: () => {} }, container, onDispose: () => {},
    } as unknown as EditorSession, {
      has: () => true,
      getState: () => ({ supported: true, enabled: true, active: false, value: 'editing' }),
      observe: () => () => {}, run,
    });
    const handle = installMacros(session, container, () => {}, { commands: interaction.commands });
    handle.toggleRecording();
    await engine.ui.commands.executeAsync('bold');
    await engine.ui.commands.executeAsync('document-mode', { mode: 'editing' });
    handle.toggleRecording();
    const replay = handle.kit.replayRecording(handle.kit.listRecordings()[0]!.id);
    await Promise.resolve();
    expect(run).toHaveBeenCalledOnce();
    const release = interaction.block();
    finishFirst();
    expect(await replay).toMatchObject({ ok: false });
    expect(run).toHaveBeenCalledOnce();
    release();
    handle.dispose();
  });

  it('toggleRecording מקליט הקלדה ופקודות ושומר בשם רץ', async () => {
    const { session, engine } = createFakeSession();
    const container = document.createElement('div');
    const statuses: string[] = [];
    const handle = installMacros(session, container, (message) => statuses.push(message));

    handle.toggleRecording();
    expect(handle.recording.value).toBe(true);
    expect(statuses).toEqual([MACRO_STATUS.recordingStarted]);

    typeChar(container, 'ש');
    typeChar(container, 'לום'); // גם קלט רב-תווי (הדבקה) נקלט, ומתלכד עם ההקשה שלפניו
    await engine.ui.commands.executeAsync('bold');
    handle.toggleRecording();

    expect(handle.recording.value).toBe(false);
    expect(statuses[1]).toBe(recordingSavedText('מאקרו 1'));
    const recordings = handle.kit.listRecordings();
    expect(recordings).toHaveLength(1);
    expect(recordings[0].steps).toEqual([
      { type: 'insert-text', text: 'שלום' },
      { type: 'command', id: 'bold' },
    ]);

    handle.dispose();
  });

  it('הדבקה נקלטת להקלטה — גם כשהטקסט מגיע ב-dataTransfer', () => {
    const { session } = createFakeSession();
    const container = document.createElement('div');
    const handle = installMacros(session, container, () => undefined);

    handle.toggleRecording();
    container.dispatchEvent(
      new InputEvent('beforeinput', { inputType: 'insertFromPaste', data: 'מודבק', bubbles: true }),
    );
    handle.toggleRecording();

    expect(handle.kit.listRecordings()[0]!.steps).toEqual([{ type: 'insert-text', text: 'מודבק' }]);
    handle.dispose();
  });

  it('הקלדת IME מתלכדת לאירוע אחד מ-compositionend', () => {
    const { session } = createFakeSession();
    const container = document.createElement('div');
    const handle = installMacros(session, container, () => undefined);

    handle.toggleRecording();
    container.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    // אירועי הביניים של ההרכבה — כל אחד נושא את המועמד המצטבר — חייבים
    // להיבלע, אחרת המילה הייתה מוקלטת פעם על כל הקשה.
    container.dispatchEvent(
      new InputEvent('beforeinput', { inputType: 'insertCompositionText', data: 'ש', bubbles: true }),
    );
    container.dispatchEvent(
      new InputEvent('beforeinput', { inputType: 'insertCompositionText', data: 'שלום', bubbles: true }),
    );
    container.dispatchEvent(new CompositionEvent('compositionend', { data: 'שלום', bubbles: true }));
    handle.toggleRecording();

    expect(handle.kit.listRecordings()[0]!.steps).toEqual([{ type: 'insert-text', text: 'שלום' }]);
    handle.dispose();
  });

  it('ברירת המחדל: סקריפטים כבויים — ה-kit עצמו מסרב, לא רק הלשונית', async () => {
    const { session } = createFakeSession();
    const container = document.createElement('div');
    const handle = installMacros(session, container, () => undefined);

    expect(handle.scriptsEnabled).toBe(false);
    const result = await handle.kit.runSource('return 1');
    expect(result.ok).toBe(false);

    handle.dispose();
  });

  it('הדלקת הדגל ב-localStorage מדליקה את הסקריפטים', () => {
    localStorage.setItem(SCRIPTS_FLAG_KEY, 'on');
    const { session } = createFakeSession();
    const container = document.createElement('div');
    const handle = installMacros(session, container, () => undefined);

    expect(handle.scriptsEnabled).toBe(true);
    handle.dispose();
  });

  it('replayLast בלי הקלטה מדווח שאין מה לנגן', () => {
    const { session } = createFakeSession();
    const container = document.createElement('div');
    const statuses: Array<{ message: string; isError?: boolean }> = [];
    const handle = installMacros(session, container, (message, isError) =>
      statuses.push({ message, isError }),
    );

    handle.replayLast();

    expect(statuses).toEqual([{ message: MACRO_STATUS.noRecordings, isError: true }]);
    handle.dispose();
  });

  it('שותל קטע ברירת מחדל פעם אחת בלבד', () => {
    const { session } = createFakeSession();
    const container = document.createElement('div');
    const first = installMacros(session, container, () => undefined);
    expect(first.kit.listSnippets()).toHaveLength(1);
    first.dispose();

    const second = installMacros(session, container, () => undefined);
    expect(second.kit.listSnippets()).toHaveLength(1);
    second.dispose();
  });

  it('מי שמחק את כל הקטעים לא מקבל את קטע הדוגמה בחזרה', () => {
    const { session } = createFakeSession();
    const container = document.createElement('div');
    const first = installMacros(session, container, () => undefined);
    first.kit.removeSnippet(first.kit.listSnippets()[0]!.id);
    first.dispose();

    const second = installMacros(session, container, () => undefined);
    expect(second.kit.listSnippets()).toHaveLength(0);
    second.dispose();
  });

  it('פעולה שאינה ניתנת להקלטה: שמירה רק אחרי אישור מפורש', async () => {
    const { session, engine } = createFakeSession();
    const container = document.createElement('div');
    const statuses: string[] = [];
    let consent = false;
    const asked: string[] = [];
    const handle = installMacros(session, container, (message) => statuses.push(message), {
      confirmIncomplete: async (title) => {
        asked.push(title);
        return consent;
      },
    });

    // "הכנסת תמונה": פקודה שה-payload שלה גדול מהתקרה.
    handle.toggleRecording();
    typeChar(container, 'א');
    await engine.ui.commands.executeAsync('bold', { src: 'x'.repeat(20_000) });
    await new Promise((resolve) => setTimeout(resolve, 0));

    handle.toggleRecording(); // עצירה — המשתמש מסרב לשמירה חלקית
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(asked).toHaveLength(1);
    expect(handle.kit.listRecordings()).toHaveLength(0);
    expect(statuses).toContain(MACRO_STATUS.incompleteDiscarded);

    // סבב שני — הפעם עם הסכמה.
    consent = true;
    handle.toggleRecording();
    typeChar(container, 'ב');
    await engine.ui.commands.executeAsync('bold', { src: 'x'.repeat(20_000) });
    await new Promise((resolve) => setTimeout(resolve, 0));
    handle.toggleRecording();
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(handle.kit.listRecordings()).toHaveLength(1);
    expect(handle.kit.listRecordings()[0]!.steps).toEqual([{ type: 'insert-text', text: 'ב' }]);
    handle.dispose();
  });

  it('dispose מחזיר את ה-controller לקדמותו וזורק הקלטה פתוחה', async () => {
    const { session, engine } = createFakeSession();
    const original = engine.ui.commands.executeAsync;
    const container = document.createElement('div');
    const handle = installMacros(session, container, () => undefined);

    expect(engine.ui.commands.executeAsync).not.toBe(original);

    handle.toggleRecording();
    handle.dispose();

    // ההקלטה שנשארה פתוחה נזרקה — לא נשמרה, והמצב התאפס.
    expect(handle.recording.value).toBe(false);
    expect(handle.kit.listRecordings()).toHaveLength(0);
    expect(engine.ui.commands.executeAsync).toBe(original);
  });
});
