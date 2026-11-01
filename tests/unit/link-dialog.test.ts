/**
 * זרימת הקישור. עד שהיא עברה למעטפת היא ישבה בתוך `InsertTab` **ולא הייתה
 * מכוסה בכלל**: נבדקו הדיאלוג לבדו (`tests/component/dialogs.test.ts`) ובנאי
 * ה-payload לבדו (`tests/contract/link-payload.test.ts`), ואילו החוליה
 * שמחברת ביניהם — תצלום הבחירה, ההכרעה אם לשלוח `text`, וחסימת כתובת פסולה —
 * לא נבדקה מעולם.
 */
import { describe, it, expect, vi } from 'vitest';
import { createLinkDialog, type LinkDialogDeps } from '../../src/composables/use-link-dialog';
import { emptySelectionSnapshot, type DocSelectionSnapshot } from '../../src/engine/doc-selection';
import type { CommandOutcome } from '../../src/engine/command-adapter';

function snapshot(over: Partial<DocSelectionSnapshot> = {}): DocSelectionSnapshot {
  return { ...emptySelectionSnapshot(), ...over };
}

function setup(over: Partial<LinkDialogDeps> = {}) {
  const runLink = vi.fn();
  const reports: Array<{ outcome: CommandOutcome; id: string }> = [];
  const deps: LinkDialogDeps = {
    canOpen: () => true,
    readSelection: async () => snapshot(),
    runLink,
    report: (outcome, id) => reports.push({ outcome, id }),
    ...over,
  };
  return { runLink, reports, dialog: createLinkDialog(deps) };
}

describe('createLinkDialog', () => {
  it('נפתח סגור', () => {
    const { dialog } = setup();
    expect(dialog.isOpen.value).toBe(false);
  });

  it('הפתיחה מצלמת את הבחירה', async () => {
    // הדיאלוג גוזל את המיקוד מהעורך ברגע שמקלידים בו. בלי התצלום הקישור היה
    // נכתב על טווח שכבר אינו קיים.
    const selection = snapshot({ hasRange: true, text: 'רש"י', target: { blockId: 'b1' } });
    const { dialog } = setup({ readSelection: async () => selection });

    await dialog.open();

    expect(dialog.isOpen.value).toBe(true);
    expect(dialog.selection.value).toEqual(selection);
  });

  it('הבחירה נשארת כפי שהייתה גם אם המסמך השתנה מאז', async () => {
    let live = snapshot({ hasRange: true, text: 'ראשון' });
    const { dialog, runLink } = setup({ readSelection: async () => live });

    await dialog.open();
    live = snapshot({ hasRange: false, text: '' });

    dialog.submit({ href: 'https://example.org', text: 'טקסט חדש' });

    // עם טווח מסומן אין `text` — וזה מה שהתצלום קובע, לא המצב החי.
    expect(runLink).toHaveBeenCalledWith({ href: 'https://example.org/' });
  });

  it('עם טווח מסומן: אין text — המנוע עוטף את מה שכבר מסומן', async () => {
    const { dialog, runLink } = setup({
      readSelection: async () => snapshot({ hasRange: true, text: 'גמרא', target: { id: 't' } }),
    });

    await dialog.open();
    dialog.submit({ href: 'https://sefaria.org', text: 'לא אמור להישלח' });

    expect(runLink).toHaveBeenCalledWith({
      href: 'https://sefaria.org/',
      target: { id: 't' },
    });
  });

  it('בלי טווח: text נשלח, כי אין מה לעטוף', async () => {
    const { dialog, runLink } = setup({
      readSelection: async () => snapshot({ hasRange: false, target: { id: 't' } }),
    });

    await dialog.open();
    dialog.submit({ href: 'https://otzaria.org', text: 'אוצריא' });

    expect(runLink).toHaveBeenCalledWith({
      href: 'https://otzaria.org/',
      text: 'אוצריא',
      target: { id: 't' },
    });
  });

  it('אישור מוצלח סוגר את הדיאלוג', async () => {
    const { dialog } = setup();

    await dialog.open();
    dialog.submit({ href: 'https://example.org', text: 'כאן' });

    expect(dialog.isOpen.value).toBe(false);
  });

  it('כתובת פסולה מדווחת ואינה מגיעה לפקודה', async () => {
    const { dialog, runLink, reports } = setup();

    await dialog.open();
    dialog.submit({ href: '   ', text: 'כאן' });

    expect(runLink).not.toHaveBeenCalled();
    expect(reports).toHaveLength(1);
    expect(reports[0]!.id).toBe('link');
    expect(reports[0]!.outcome.ok).toBe(false);
    // הדיאלוג נשאר פתוח: יש מה לתקן.
    expect(dialog.isOpen.value).toBe(true);
  });

  it('סגירה אינה מריצה פקודה', async () => {
    const { dialog, runLink } = setup();

    await dialog.open();
    dialog.close();

    expect(dialog.isOpen.value).toBe(false);
    expect(runLink).not.toHaveBeenCalled();
  });

  it('בלי מסמך פתוח — אין דיאלוג ריק, יש הודעה', async () => {
    const { dialog, reports } = setup({ canOpen: () => false });

    await dialog.open();

    expect(dialog.isOpen.value).toBe(false);
    expect(reports).toHaveLength(1);
    expect(reports[0]!.outcome.ok).toBe(false);
  });

  it('פתיחה חוזרת מצלמת מחדש', async () => {
    let current = snapshot({ text: 'ראשון' });
    const { dialog } = setup({ readSelection: async () => current });

    await dialog.open();
    dialog.close();
    current = snapshot({ text: 'שני', hasRange: true });
    await dialog.open();

    expect(dialog.selection.value.text).toBe('שני');
  });

  describe('סמן בלי טווח: insertAtCaret ולא פקודת המנוע', () => {
    // פקודת `link` של המנוע כותבת כאן ב-`hyperlinks.insert`, שנדחית בכל פסקה
    // שכבר יש בה קישור — „קישור אחד עובד, השני לא”.
    const at = { kind: 'selection' };

    it('הכתיבה עוברת ל-insertAtCaret, והתוצאה מדווחת', async () => {
      const insertAtCaret = vi.fn(async () => ({ ok: true as const }));
      const { dialog, runLink, reports } = setup({
        readSelection: async () => snapshot({ selectionTarget: at, target: { id: 't' } }),
        insertAtCaret,
      });

      await dialog.open();
      dialog.submit({ href: 'https://b.org', text: 'שני' });
      await vi.waitFor(() => expect(reports).toHaveLength(1));

      expect(insertAtCaret).toHaveBeenCalledWith({ href: 'https://b.org/', text: 'שני', at });
      expect(runLink).not.toHaveBeenCalled();
      expect(reports[0]).toEqual({ outcome: { ok: true }, id: 'link' });
    });

    it('null — נופל לפקודת המנוע', async () => {
      const { dialog, runLink } = setup({
        readSelection: async () => snapshot({ selectionTarget: at }),
        insertAtCaret: async () => null,
      });

      await dialog.open();
      dialog.submit({ href: 'https://b.org', text: 'שני' });

      await vi.waitFor(() => expect(runLink).toHaveBeenCalledWith({ href: 'https://b.org/', text: 'שני' }));
    });

    it('עם טווח מסומן — עדיין פקודת המנוע (wrap)', async () => {
      const insertAtCaret = vi.fn();
      const { dialog, runLink } = setup({
        readSelection: async () => snapshot({ hasRange: true, selectionTarget: at, target: { id: 't' } }),
        insertAtCaret,
      });

      await dialog.open();
      dialog.submit({ href: 'https://b.org', text: '' });

      expect(insertAtCaret).not.toHaveBeenCalled();
      expect(runLink).toHaveBeenCalled();
    });
  });
});
