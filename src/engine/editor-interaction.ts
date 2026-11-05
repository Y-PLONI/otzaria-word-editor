import { shallowRef, type Ref } from 'vue';
import type { CommandState } from 'superdoc/ui';
import type { EditorSession } from './create-editor';
import { reasonText, type CommandAdapter } from './command-adapter';

export interface EditorInteraction {
  readonly blocked: Readonly<Ref<boolean>>;
  readonly pendingModeChanges: Readonly<Ref<number>>;
  readonly commands: CommandAdapter;
  /** Only this editor is locked; a replacement can initialize independently. */
  block(): () => void;
}

/**
 * A temporary, non-document lock. `inert` alone misses Ribbon commands and
 * Document API callers holding this editor. SuperDoc's public viewing mode
 * rejects those writes too (see protection.ts), without changing DOCX protection.
 * The adapter also blocks `document-mode`, so UI commands cannot undo the lock.
 */
export function createEditorInteraction(editor: EditorSession, adapter: CommandAdapter): EditorInteraction {
  const blocked = shallowRef(false);
  const pendingModeChanges = shallowRef(0);
  const listeners = new Set<() => void>();
  let disposed = false;
  let leases = 0;
  let restore: (() => void) | undefined;

  const stateWhileBlocked = (state: CommandState): CommandState =>
    blocked.value ? { ...state, enabled: false, reason: 'not-ready' } : state;

  const publish = (value: boolean): void => {
    blocked.value = value;
    for (const listener of listeners) listener();
  };

  editor.onDispose(() => {
    // Disposers run before SuperDoc.destroy(). Never restore the mode of a
    // destroyed editor, or allow a retained adapter to dispatch to it.
    disposed = true;
    publish(true);
    listeners.clear();
  });

  return {
    blocked,
    pendingModeChanges,
    commands: {
      beginModeChange() {
        if (blocked.value) return null;
        pendingModeChanges.value += 1;
        let released = false;
        return () => {
          if (released) return;
          released = true;
          pendingModeChanges.value -= 1;
        };
      },
      has: (id) => adapter.has(id),
      getState: (id) => stateWhileBlocked(adapter.getState(id)),
      observe(id, listener) {
        const notify = (): void => listener(stateWhileBlocked(adapter.getState(id)));
        const stop = adapter.observe(id, (state) => listener(stateWhileBlocked(state)));
        listeners.add(notify);
        return () => { listeners.delete(notify); stop(); };
      },
      run(id, payload) {
        if (blocked.value) return Promise.resolve({ ok: false, reason: 'not-ready', message: reasonText('not-ready') });
        return adapter.run(id, payload);
      },
    },
    block() {
      if (disposed) throw new Error('המסמך כבר נסגר');
      if (pendingModeChanges.value > 0) throw new Error('שינוי הגנת המסמך עדיין מתבצע');
      if (leases === 0) {
        const previousMode = adapter.getState('document-mode').value;
        if (previousMode !== 'editing' && previousMode !== 'suggesting' && previousMode !== 'viewing') {
          throw new Error('לא ניתן לנעול את המסמך בזמן הפתיחה');
        }
        const previousInert = editor.container.inert;
        restore = () => {
          try {
            if (!disposed && previousMode !== 'viewing') editor.superdoc.setDocumentMode(previousMode);
          } finally {
            editor.container.inert = previousInert;
            if (!disposed) publish(false);
          }
        };
        publish(true);
        editor.container.inert = true;
        try {
          if (previousMode !== 'viewing') editor.superdoc.setDocumentMode('viewing');
        } catch (error) {
          restore();
          restore = undefined;
          throw error;
        }
      }
      leases += 1;
      let released = false;
      return () => {
        if (released) return;
        released = true;
        if (--leases !== 0) return;
        const release = restore;
        restore = undefined;
        release?.();
      };
    },
  };
}
