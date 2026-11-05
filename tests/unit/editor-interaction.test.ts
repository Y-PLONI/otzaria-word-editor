import { describe, expect, it, vi } from 'vitest';
import type { CommandState } from 'superdoc/ui';
import type { EditorSession } from '../../src/engine/create-editor';
import type { CommandAdapter } from '../../src/engine/command-adapter';
import { createEditorInteraction } from '../../src/engine/editor-interaction';

type Mode = 'editing' | 'suggesting' | 'viewing';

function fixture(initialMode: Mode = 'editing', inert = false) {
  let mode = initialMode;
  const disposers: Array<() => void> = [];
  const container = document.createElement('div');
  container.inert = inert;
  const run = vi.fn(async () => ({ ok: true as const }));
  const setDocumentMode = vi.fn((next: Mode) => { mode = next; });
  const getState = (id: string): CommandState => ({
    supported: true, enabled: true, active: false, value: id === 'document-mode' ? mode : undefined,
  });
  const adapter: CommandAdapter = { has: () => true, getState, observe: () => () => {}, run };
  const editor = {
    container,
    superdoc: { setDocumentMode },
    onDispose: (dispose: () => void) => { disposers.push(dispose); },
  } as unknown as EditorSession;
  return {
    interaction: createEditorInteraction(editor, adapter), container, run, setDocumentMode,
    mode: () => mode,
    dispose: () => { for (const dispose of disposers) dispose(); },
  };
}

describe('old editor interaction lease', () => {
  it.each(['editing', 'suggesting', 'viewing'] as const)('restores the original %s mode and prior inert state', async (mode) => {
    const old = fixture(mode, true);
    const listener = vi.fn();
    const stop = old.interaction.commands.observe('bold', listener);
    const release = old.interaction.block();
    expect(old.mode()).toBe('viewing');
    expect(old.container.inert).toBe(true);
    expect(old.interaction.commands.getState('bold')).toMatchObject({ enabled: false, reason: 'not-ready' });
    expect(listener).toHaveBeenLastCalledWith(expect.objectContaining({ enabled: false }));
    // This adapter may have been captured before the UI became inert. In
    // particular, changing document-mode must not escape the engine's lock.
    expect(await old.interaction.commands.run('document-mode', { mode: 'editing' })).toMatchObject({ ok: false });
    expect(await old.interaction.commands.run('bold')).toMatchObject({ ok: false });
    expect(old.run).not.toHaveBeenCalled();

    release();
    expect(old.mode()).toBe(mode);
    expect(old.container.inert).toBe(true);
    expect(listener).toHaveBeenLastCalledWith(expect.objectContaining({ enabled: true }));
    await old.interaction.commands.run('bold');
    expect(old.run).toHaveBeenCalledOnce();
    stop();
  });

  it('a stale release cannot unlock a newer lease; another editor remains writable', async () => {
    const old = fixture('suggesting'), candidate = fixture();
    const releaseA = old.interaction.block();
    releaseA();
    const releaseB = old.interaction.block();
    releaseA();
    expect(old.mode()).toBe('viewing');
    expect(old.container.inert).toBe(true);
    await candidate.interaction.commands.run('bold');
    expect(candidate.run).toHaveBeenCalledOnce();
    expect(candidate.mode()).toBe('editing');
    expect(candidate.setDocumentMode).not.toHaveBeenCalled();
    releaseB();
    expect(old.mode()).toBe('suggesting');
    expect(old.container.inert).toBe(false);
  });

  it('does not restore mode or dispatch through a retained adapter after disposal', async () => {
    const old = fixture();
    const release = old.interaction.block();
    old.dispose();
    release();
    expect(old.setDocumentMode.mock.calls).toEqual([['viewing']]);
    expect(old.container.inert).toBe(false);
    expect(await old.interaction.commands.run('bold')).toMatchObject({ ok: false });
    expect(old.run).not.toHaveBeenCalled();
  });

  it('rolls back a failed acquisition without leaving the editor or commands locked', async () => {
    const old = fixture('suggesting');
    old.setDocumentMode.mockImplementationOnce(() => { throw new Error('mode failed'); });
    expect(() => old.interaction.block()).toThrow('mode failed');
    expect(old.mode()).toBe('suggesting');
    expect(old.container.inert).toBe(false);
    expect(old.interaction.blocked.value).toBe(false);
    const release = old.interaction.block();
    expect(old.mode()).toBe('viewing');
    release();
    expect(old.mode()).toBe('suggesting');
  });
});
