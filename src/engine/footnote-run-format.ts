/**
 * הקלדה אחרי סימן הערה שהוכנס זה עתה ממשיכה בגופן של הגוף, ולא בברירת
 * המחדל של המסמך.
 *
 * ## הבעיה
 *
 * `footnotes.insert` כותב את סימן ההפניה כריצה שה-`rPr` שלה הוא
 * `rStyle=FootnoteReference` בלבד, בלי העיצוב הישיר של המקום שבו ההערה
 * הוכנסה. טקסט שמוקלד מיד אחרי הסימן יורש את הסימונים של מה שלפני הסמן —
 * כלומר של הסימן — ולכן אין לו גופן וגודל משלו, והוא נופל ל-docDefaults
 * (Arial 12). Enter שאחריו מעביר את זה לפסקה הבאה, ומשם „הכתב במסמך נהיה
 * כמו הכתב של ההערה” (Arial הוא גם מה שתוכן ההערה מקבל). Word מעתיק לריצת
 * ההפניה את העיצוב הישיר של נקודת ההכנסה, ולכן אצלו זה לא קורה.
 *
 * נמדד על 2.23 בדפדפן אמיתי, על גוף ב-TaameyDavidCLM 16:
 *
 * - ההקלדה אחרי הסימן נכתבה כ-`<w:r><w:t> המשך</w:t></w:r>` — בלי `rPr`.
 * - הרצועה הציגה במקום הזה „Arial / 8” — הקשר של הסימן ולא של הגוף.
 * - Enter מיד אחרי החזרה לגוף, **בלי** הקלדה קודם, לא דולף.
 *
 * ## למה לא לתקן את ריצת ההפניה עצמה
 *
 * דווח: superdoc/docx-editor#4126. כשהמנוע יעתיק את העיצוב הישיר לריצת
 * ההפניה, המודול הזה מיותר וצריך להסירו.
 *
 * זה התיקון הנכון, ואינו אפשרי מבחוץ: `format.apply` על התו של הסימן
 * מחזיר `INVALID_TARGET` („format target could not be resolved.”), ועל טווח
 * שמכיל גם את התו שלפניו הוא מצליח — ומדלג על הסימן. נמדד על 2.17 ועל 2.23.
 *
 * ## מה כן
 *
 * לפני ההכנסה נקראים המקום של הסמן והגופן והגודל שהרצועה מדווחת שם. כשהסמן
 * חוזר לגוף **בדיוק** אל המקום שאחרי הסימן, פקודות `font-family` ו-`font-size`
 * רצות פעם אחת על הסמן המכווץ. על סמן מכווץ הן אינן משנות טקסט אלא קובעות
 * את הסימונים של ההקלדה הבאה — אותו מסלול שבו „בוחרים גופן ואז מקלידים”.
 * נמדד: אחרי זה גם ההמשך בשורה וגם השורה שאחרי Enter נכתבו ב-
 * `<w:sz w:val="32"/><w:rFonts w:ascii="TaameyDavidCLM" w:hAnsi="TaameyDavidCLM"/>`,
 * זהה לריצה שלפני ההערה.
 *
 * התיקון תקף רק בחזרה הראשונה לגוף, ורק אם הגוף לא השתנה בינתיים.
 * שינוי בגוף מבטל את העוגן במקום לנחש לאן ההפניה זזה; עריכה בתוך ההערה
 * אינה מבטלת אותו. כך עוגן של הערה שנמחקה אינו דורס בחירת גופן מאוחרת.
 */
import { defineSuperDocExtension, type SuperDoc } from 'superdoc';
import { fontFamilyPayload, fontSizePayload } from './payloads';
import type { MaybePromise } from './document-api';

interface CommandsLike {
  has(id: string): boolean;
  get(id: string): { getState(): { value?: unknown } };
  executeAsync(id: string, payload?: unknown): Promise<unknown>;
}

interface TextPointLike {
  kind?: string;
  blockId?: unknown;
  offset?: unknown;
  story?: { storyType?: unknown } | null;
}

interface SelectionTargetLike {
  kind?: string;
  start?: TextPointLike;
  end?: TextPointLike;
  story?: { storyType?: unknown } | null;
}

interface SelectionSliceLike {
  status?: string;
  selectionTarget?: SelectionTargetLike | null;
}

export interface NoteRunFormatHost {
  activeEditor?: {
    doc?: {
      selection?: { current?: () => MaybePromise<{ selectionTarget?: unknown } | undefined> };
      footnotes?: {
        get?: (input: { target: { kind: 'entity'; entityType: 'footnote'; noteId: string } }) => MaybePromise<{ noteId?: string; type?: string } | undefined>;
      };
    } | null;
  } | null;
  ui?: {
    commands?: CommandsLike | null;
    selection?: {
      /** יורה מיד עם הצילום הנוכחי, ואחר כך בכל שינוי. */
      observe?: (listener: (snapshot: SelectionSliceLike) => void) => () => void;
    } | null;
  } | null;
}

export interface InsertedNoteIdentity {
  noteId: string;
  type: 'footnote' | 'endnote';
}

export type NoteRunFormatTarget = SuperDoc | NoteRunFormatHost | null | undefined;

/** מה שנלכד לפני ההכנסה. `null` = אין מה לשאת הלאה. */
export interface NoteAnchor {
  blockId: string;
  /** ההיסט של הסמן לפני ההכנסה. הסימן יושב בו, והמקום שאחריו הוא `offset + 1`. */
  offset: number;
  family: string | null;
  size: number | null;
}

/** סמן מכווץ בגוף המסמך. ההערה עצמה היא story אחר, ושם אין מה לשאת. */
function bodyCaret(target: unknown): { blockId: string; offset: number } | null {
  const t = target as SelectionTargetLike | null | undefined;
  if (!t || t.kind !== 'selection') return null;
  const { start, end } = t;
  if (start?.kind !== 'text' || end?.kind !== 'text') return null;
  if (typeof start.blockId !== 'string' || start.blockId !== end.blockId) return null;
  if (typeof start.offset !== 'number' || start.offset !== end.offset) return null;
  const storyType = (t.story ?? start.story)?.storyType;
  if (storyType !== undefined && storyType !== 'body') return null;
  return { blockId: start.blockId, offset: start.offset };
}

function stateValue(commands: CommandsLike, id: string): unknown {
  try {
    return commands.has(id) ? commands.get(id).getState().value : undefined;
  } catch {
    return undefined;
  }
}

/**
 * קוראת, **לפני** ההכנסה, את המקום ואת הגופן והגודל שם. אחרי ההכנסה הסמן
 * כבר בתוך ההערה, והרצועה מדווחת את ההקשר שלה.
 */
export async function captureNoteAnchor(host: NoteRunFormatTarget): Promise<NoteAnchor | null> {
  const h = host as NoteRunFormatHost | null | undefined;
  const commands = h?.ui?.commands;
  const current = h?.activeEditor?.doc?.selection?.current;
  if (!commands || typeof current !== 'function' || typeof h?.ui?.selection?.observe !== 'function') return null;

  let caret: { blockId: string; offset: number } | null;
  try {
    caret = bodyCaret((await current())?.selectionTarget);
  } catch {
    return null;
  }
  if (!caret) return null;

  const family = fontFamilyPayload(String(stateValue(commands, 'font-family') ?? ''));
  const rawSize = stateValue(commands, 'font-size');
  const size = rawSize === undefined || rawSize === null || rawSize === '' ? null : fontSizePayload(rawSize as string);
  if (family === null && size === null) return null;
  return { ...caret, family, size };
}

/** המעקב הפעיל לכל מסמך. הערה חדשה מחליפה את הקודמת: רק האחרונה רלוונטית. */
const pending = new WeakMap<object, () => void>();

/**
 * אחרי הכנסה מוצלחת: מחכה שהסמן יחזור למקום שאחרי הסימן, ומריצה שם את
 * פקודות הגופן פעם אחת. חזרה למקום אחר בגוף או שינוי בגוף מבטלים אותו.
 */
export function carryFormatPastNote(host: NoteRunFormatTarget, anchor: NoteAnchor, note: InsertedNoteIdentity): void {
  const h = host as NoteRunFormatHost | null | undefined;
  const commands = h?.ui?.commands;
  const observe = h?.ui?.selection?.observe;
  const get = h?.activeEditor?.doc?.footnotes?.get;
  if (!h || !commands || typeof observe !== 'function' || typeof get !== 'function') return;

  pending.get(h)?.();
  const editor = h.activeEditor;
  const after = anchor.offset + 1;
  let applying = false;
  let enteredNote = false;
  let allowPreInsertionSnapshot = true;
  let done = false;
  let unsubscribe: (() => void) | null = null;
  const stop = (): void => {
    done = true;
    unsubscribe?.();
    unsubscribe = null;
    if (pending.get(h) === stop) pending.delete(h);
  };
  pending.set(h, stop);

  unsubscribe = observe.call(h.ui!.selection, (slice) => {
    if (done) return;
    if (slice.status !== 'ready') return;
    if (h.activeEditor !== editor) return stop();
    const target = slice.selectionTarget;
    const storyType = (target?.story ?? target?.start?.story)?.storyType;
    if (storyType === 'footnote' || storyType === 'endnote') {
      enteredNote = true;
      if (applying) stop();
      return;
    }
    const caret = bodyCaret(target);
    if (!enteredNote) {
      // The first settled snapshot can still describe the pre-insertion body caret.
      // Ignore that snapshot once; later body selections must not leave a dormant anchor.
      if (allowPreInsertionSnapshot && caret?.blockId === anchor.blockId && caret.offset === anchor.offset) {
        allowPreInsertionSnapshot = false;
        return;
      }
      return stop();
    }
    if (!caret || caret.blockId !== anchor.blockId || caret.offset !== after) return stop();
    if (applying) return;
    applying = true;
    void (async () => {
      // Some 2.23 note removals omit the body's affected-story event. Check the
      // inserted entity as well; ambiguous footnote/endnote ids fail closed.
      try {
        const live = await get({ target: { kind: 'entity', entityType: 'footnote', noteId: note.noteId } });
        if (live?.noteId !== note.noteId || live.type !== note.type) return;
      } catch {
        return;
      }
      await applyCaretFormat(commands, anchor, () => !done && h.activeEditor === editor);
    })().finally(stop);
  });
  // `observe` יורה מיד, עוד לפני שההשמה למעלה הסתיימה.
  if (done) unsubscribe?.();
}

async function applyCaretFormat(commands: CommandsLike, anchor: NoteAnchor, isCurrent: () => boolean): Promise<void> {
  try {
    if (isCurrent() && anchor.family !== null && commands.has('font-family')) await commands.executeAsync('font-family', anchor.family);
    if (isCurrent() && anchor.size !== null && commands.has('font-size')) await commands.executeAsync('font-size', anchor.size);
  } catch (error) {
    // ההערה כבר נכתבה; גופן ההקלדה הבאה אינו שווה הודעת שגיאה.
    console.warn('[otzaria-word] הערה: החזרת גופן הגוף אחרי הסימן נכשלה', error);
  }
}

/**
 * הגוף הוא מרחב הקואורדינטות של העוגן. כל שינוי בו (כולל Undo, מחיקה
 * והחלפת מסמך) מבטל את התיקון; הקלדה בתוכן ההערה רשאית להמשיך.
 * אירועי המוטציה הציבוריים כוללים את ה-stories שהשתנו, בלי סריקת מסמך.
 */
export function createNoteRunFormatExtension() {
  let host: NoteRunFormatTarget = null;
  const cancel = (): void => {
    if (host) pending.get(host)?.();
  };
  return {
    extension: defineSuperDocExtension({
      id: 'otzaria.note-run-format',
      activate(context) {
        cancel();
        context.onMutation({}, (event) => {
          if (event.stories.length === 0 || event.stories.some((story) => story.storyType === 'body')) cancel();
        });
        return cancel;
      },
    }),
    bind(target: NoteRunFormatTarget): void {
      cancel();
      host = target;
    },
    dispose(): void {
      cancel();
      host = null;
    },
  };
}
