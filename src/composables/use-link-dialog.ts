/**
 * דיאלוג הקישור, כמצב של המעטפת.
 *
 * ## למה הוא עבר לכאן מ-`InsertTab`
 *
 * לשוניות הרצועה מורכבות רק כשהן פעילות („mount on active”), ולכן דיאלוג
 * שחי בתוך לשונית אינו קיים כשלשונית אחרת פתוחה — ו-`Ctrl+K` פשוט לא היה
 * יכול לפתוח אותו. שאר הדיאלוגים של הממשק (חיפוש, אודות) כבר יושבים במעטפת
 * מאותה סיבה; הקישור היה היוצא מן הכלל.
 *
 * המודול הזה הוא מצב טהור בלי תבנית, כדי שההכרעות — ובראשן תצלום הבחירה —
 * ייבדקו בלי להרכיב קומפוננטה.
 */
import { ref, shallowRef, type Ref, type ShallowRef } from 'vue';
import type { CommandOutcome } from '../engine/command-adapter';
import { emptySelectionSnapshot, type DocSelectionSnapshot } from '../engine/doc-selection';
import { LINK_HREF_HINT, linkPayload } from '../engine/payloads';

export interface LinkDialogDeps {
  /** האם יש מסמך שאפשר לקשר בו. בלעדיו הדיאלוג לא נפתח כלל. */
  canOpen: () => boolean;
  /** קורא את הבחירה הנוכחית מהמסמך. */
  readSelection: () => Promise<DocSelectionSnapshot>;
  /** מריץ את פקודת `link` עם ה-payload. */
  runLink: (payload: unknown) => void;
  /**
   * כותב קישור בסמן, כשאין טווח מסומן. `null` = אין סמן שאפשר לכתוב בו,
   * ואז הדיאלוג נופל ל-`runLink`. ראו `insertHyperlinkAtCaret` למה פקודת
   * המנוע אינה המסלול כאן.
   */
  insertAtCaret?: (link: { href: string; text?: string; at: unknown }) => Promise<CommandOutcome | null>;
  /** מדווח כשל שלא הגיע מהפקודה עצמה. */
  report: (outcome: CommandOutcome, commandId: string) => void;
}

export interface LinkDialogState {
  isOpen: Ref<boolean>;
  /**
   * הבחירה **כפי שהייתה ברגע הפתיחה**. השדה בדיאלוג לוקח את המיקוד מהעורך
   * ברגע שמקלידים בו, ובלי התצלום הזה הקישור היה נכתב על טווח שאינו קיים.
   */
  selection: ShallowRef<DocSelectionSnapshot>;
  open: () => Promise<void>;
  close: () => void;
  submit: (link: { href: string; text: string }) => void;
}

export function createLinkDialog(deps: LinkDialogDeps): LinkDialogState {
  const isOpen = ref(false);
  const selection = shallowRef<DocSelectionSnapshot>(emptySelectionSnapshot());

  async function open(): Promise<void> {
    // בלי מסמך הדיאלוג היה נפתח ריק, והאישור היה נכשל סגור אחרי שהמשתמש כבר
    // הקליד כתובת. עדיף לומר זאת מראש.
    if (!deps.canOpen()) {
      deps.report({ ok: false, message: 'אין מסמך פתוח להוספת קישור', reason: 'not-ready' }, 'link');
      return;
    }

    selection.value = await deps.readSelection();
    isOpen.value = true;
  }

  function close(): void {
    isOpen.value = false;
  }

  function submit(link: { href: string; text: string }): void {
    const snapshot = selection.value;

    const payload = linkPayload({
      href: link.href,
      // עם טווח מסומן המסלול הוא `hyperlinks.wrap`, שמתעלם מ-`text`. שליחתו
      // הייתה יוצרת ציפייה שהטקסט המסומן יוחלף.
      text: snapshot.hasRange ? undefined : link.text,
      target: snapshot.target ?? undefined,
    });

    if (!payload) {
      // הדיאלוג אינו מאפשר אישור של כתובת פסולה, ולכן זו הגנה על החוזה ולא
      // מצב שאפשר להגיע אליו דרך הממשק.
      deps.report({ ok: false, message: LINK_HREF_HINT, reason: 'invalid-href' }, 'link');
      return;
    }

    isOpen.value = false;

    if (!snapshot.hasRange && snapshot.selectionTarget && deps.insertAtCaret) {
      void runAtCaret(payload, snapshot.selectionTarget);
      return;
    }
    deps.runLink(payload);
  }

  async function runAtCaret(payload: { href: string; text?: string }, at: unknown): Promise<void> {
    let outcome: CommandOutcome | null;
    try {
      outcome = await deps.insertAtCaret!({ href: payload.href, text: payload.text, at });
    } catch (error) {
      outcome = { ok: false, message: error instanceof Error ? error.message : 'הוספת הקישור נכשלה', reason: 'threw' };
    }
    if (outcome === null) deps.runLink(payload);
    else deps.report(outcome, 'link');
  }

  return { isOpen, selection, open, close, submit };
}
