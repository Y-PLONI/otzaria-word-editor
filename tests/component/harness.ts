/**
 * תשתית ההרכבה של בדיקות הקומפוננטות.
 *
 * ## למה הקבצים כאן קיימים
 *
 * כל בדיקת `.vue` במאגר עד עכשיו הייתה **סריקת מקור** — regex על טקסט הקובץ.
 * היא תפסה דברים אמיתיים (שלוש-עשרה כפתורים בלי `@click`), אבל היא אינה יכולה
 * לתפוס „לחצתי והפקודה לא הגיעה למנוע”: `doCut(){}` הוא HTML ו-JavaScript
 * תקינים לחלוטין, ו-`{ fontFamily: 'X' }` נראה נכון בכל סריקה. שתי משפחות
 * הבאגים האלה הן מה שההרכבה כאן מודדת.
 *
 * ## הכפיל של האדפטר, וההבדל בין כפיל לכפיל
 *
 * הבדיקה שנמחקה (`tests/unit/ribbon-commands.test.ts`) הריצה את ה-payloads מול
 * `executeAsync(id, payload) { calls.push(...); return true; }` — כפיל שמסכים
 * לכל דבר, ולכן אישר בירוק את חמשת ה-payloads שהמנוע דוחה בשקט. הכפיל כאן
 * עושה את ההפוך: הוא **מריץ את ה-payload דרך הוולידטורים האמיתיים של
 * superdoc** (tests/support/superdoc-engine.ts), ומחזיר כשל על מה שהמנוע היה
 * דוחה. כלומר לחיצה על בורר גופן שמעבירה `{ fontFamily }` נכשלת כאן — וזו
 * בדיוק הבדיקה שהייתה חסרה.
 *
 * הוא גם מדגם את שאר מה שה-controller עושה לפני שהוא נוגע במסמך: מזהה פקודה
 * שאינו בקטלוג נדחה, ופקודה שהמצב שלה `enabled: false` אינה מנותבת. שלוש
 * הרשימות (`calls` / `applied` / `rejected` / `blocked`) הן ההבחנה בין „הלחיצה
 * הגיעה לאדפטר”, „המנוע היה משנה את המסמך” ו„המנוע היה מסרב”.
 *
 * ## הכפיל של המופע
 *
 * ל-`ACTIVE_SUPERDOC` יש מסלול שני, שאינו עובר בפקודות בכלל: שוליים, כיוון דף,
 * עמודות, הערות שוליים, לוח, בחירת הכל, מעבר עמוד וציטוט — כולם קוראים
 * ל-Document API ישירות. הכפיל מקליט כל קריאה כזאת, וזה מה שהופך „לחצתי על
 * „שוליים צרים”” לבדיקה.
 *
 * זמינות הפעולות נמסרת דרך Proxy ולא דרך רשימה כתובה: מרחב השאלות של
 * engine/doc-capabilities.ts גדל, ורשימה קשיחה כאן הייתה משאירה פקד חדש
 * מנוטרל בלי שאיש ישים לב — כלומר בדיקה שמאשרת בירוק כפתור מת.
 */
import { afterEach } from 'vitest';
import { nextTick, ref, shallowRef, type Component, type Ref } from 'vue';
import { enableAutoUnmount, mount, type DOMWrapper, type VueWrapper } from '@vue/test-utils';
import type { SuperDoc } from 'superdoc';
import type { CommandState } from 'superdoc/ui';
import {
  reasonText,
  type CommandAdapter,
  type CommandOutcome,
} from '../../src/engine/command-adapter';
import {
  COMMAND_ADAPTER,
  COMMAND_REPORTER,
  DOCUMENT_GENERATION,
  FONT_MEMORY,
  FONT_OPTIONS,
  HEAVY_ACTION_GUARD,
  READOUT_SELECTION,
  SPELLCHECK,
  STYLE_DOCUMENT,
  STYLE_GALLERY,
  STYLE_SHORTCUTS,
  type HeavyActionGuard,
  type SpellcheckHandle,
  type StyleDocumentHandle,
  type StyleShortcutsHandle,
} from '../../src/composables/keys';
import {
  comboConflict,
  conflictMessage,
  removeShortcut,
  styleShortcutFor,
  upsertShortcut,
  type CustomShortcut,
} from '../../src/ui/shortcuts/custom-shortcuts';
import type { StyleEdit } from '../../src/engine/docx-style-edit';
import { createFontMemory, type FontMemory } from '../../src/composables/use-font-controls';
import { ACTIVE_SUPERDOC } from '../../src/engine/document-api';
import { ACTIVE_MACROS, type MacrosHandle } from '../../src/engine/macros';
import { readTip, type TipContent } from '../../src/ui/tooltip/tooltip-content';
import { fallbackFontOptions, type FontOptions } from '../../src/engine/font-options';
import { fallbackStyleGallery, type StyleGalleryState } from '../../src/engine/style-gallery';
import type { ReadoutSelection } from '../../src/engine/readout-hold';
import { checkPayload, commandDescriptor } from '../support/superdoc-engine';

/* ------------------------------------------------------------------ */
/* המתנה                                                              */
/* ------------------------------------------------------------------ */

/**
 * מחזירה את הבקרה עד שכל ה-watchers הא-סינכרוניים נרגעו.
 *
 * שתי הלשוניות שקוראות יכולות (`readDocCapabilities`) עושות זאת ב-`watch`
 * עם `immediate`, כלומר בשרשרת של כמה microtasks. בלי ההמתנה כל פקד היה נמדד
 * במצבו הראשוני — מנוטרל — וכל בדיקה כאן הייתה עוברת מהסיבה הלא נכונה.
 */
export async function settle(rounds = 6): Promise<void> {
  for (let round = 0; round < rounds; round += 1) {
    await Promise.resolve();
    await nextTick();
  }
}

/* ------------------------------------------------------------------ */
/* כפיל האדפטר                                                        */
/* ------------------------------------------------------------------ */

export interface CommandCall {
  id: string;
  payload?: unknown;
}

export interface RejectedCall extends CommandCall {
  /** איזה ולידטור של המנוע דחה, לצורך הודעת כשל שאפשר לעשות איתה משהו. */
  route: string;
}

export interface CommandDoubleOptions {
  /** מצב התחלתי לפקודות מסוימות. ברירת המחדל: נתמכת, זמינה, לא פעילה. */
  states?: Record<string, Partial<CommandState>>;
  /** ברירת המחדל לכל הפקודות שלא נזכרו ב-`states`. */
  defaultState?: Partial<CommandState>;
  /** מזהים שהמנוע „אינו מכיר”, כדי לבדוק את המסלול הזה. */
  unknown?: readonly string[];
  /** פקודות שהמנוע מנתב אך הן נכשלות, לפי `reason` של ה-controller. */
  failures?: Record<string, string>;
  /**
   * פקודות שהתשובה שלהן מושהית עד `release(id)`.
   *
   * למה זה נדרש: פקד שנשען על „מה שנבחר וטרם נענה” צריך להיבדק גם במצב הזה
   * בדיוק — בחירה שנייה בזמן שהראשונה באוויר. בלי שליטה על רגע התשובה אין דרך
   * להעמיד את המצב הזה, וההתנהגות היחידה שנמדדת היא זו של מנוע מיידי.
   */
  held?: readonly string[];
}

export interface CommandDouble extends CommandAdapter {
  /** כל קריאה ל-`run`, כולל כאלה שנחסמו — „הלחיצה הגיעה לאדפטר”. */
  readonly calls: CommandCall[];
  /** הקריאות שהמנוע היה מבצע בפועל. */
  readonly applied: CommandCall[];
  /** payload שהמנוע היה דוחה. רשימה שאינה ריקה = כפתור שנלחץ ולא קרה כלום. */
  readonly rejected: RejectedCall[];
  /** קריאה לפקודה שמצבה `enabled: false`. */
  readonly blocked: CommandCall[];
  /** ה-payloads שהוחלו עבור מזהה מסוים, לפי הסדר. */
  payloads(id: string): unknown[];
  /** מעדכנת מצב פקודה ומודיעה למי שמאזין — כמו המנוע כשהבחירה משתנה. */
  setState(id: string, patch: Partial<CommandState>): void;
  /** משחררת את התשובה המושהית **הראשונה** של הפקודה (FIFO). */
  release(id: string): void;
  reset(): void;
}

const BASE_STATE: CommandState = {
  supported: true,
  enabled: true,
  active: false,
  value: undefined,
};

export function createCommandDouble(options: CommandDoubleOptions = {}): CommandDouble {
  const calls: CommandCall[] = [];
  const applied: CommandCall[] = [];
  const rejected: RejectedCall[] = [];
  const blocked: CommandCall[] = [];

  const unknown = new Set(options.unknown ?? []);
  const failures = options.failures ?? {};
  const held = new Set(options.held ?? []);
  const waiting = new Map<string, Array<() => void>>();
  const states = new Map<string, CommandState>();
  const listeners = new Map<string, Set<(state: CommandState) => void>>();

  const has = (id: string): boolean => !unknown.has(id) && commandDescriptor(id) !== null;

  const stateOf = (id: string): CommandState => {
    const existing = states.get(id);
    if (existing) return existing;
    const initial: CommandState = {
      ...BASE_STATE,
      ...options.defaultState,
      ...options.states?.[id],
    };
    states.set(id, initial);
    return initial;
  };

  const double: CommandDouble = {
    calls,
    applied,
    rejected,
    blocked,

    has,

    getState(id) {
      // המנוע אינו מציע מצב לפקודה שאינה בקטלוג, ופקד שקורא מצב בלי לשאול
      // `has` תחילה הוא באג שקט — עדיף שייפול כאן.
      if (!has(id)) throw new Error(`אין לקרוא מצב של הפקודה ${id} — היא אינה מוכרת למנוע`);
      return stateOf(id);
    },

    observe(id, listener) {
      const set = listeners.get(id) ?? new Set();
      set.add(listener);
      listeners.set(id, set);
      return () => {
        set.delete(listener);
      };
    },

    async run(id, payload): Promise<CommandOutcome> {
      calls.push({ id, payload });

      if (!has(id)) {
        return {
          ok: false,
          message: `הפעולה ${id} אינה מוכרת למנוע`,
          reason: 'unknown-command',
        };
      }

      const state = stateOf(id);
      if (!state.enabled) {
        blocked.push({ id, payload });
        const reason = state.reason ?? 'selection-required';
        return { ok: false, message: reasonText(reason), reason };
      }

      // כאן ההבדל בין הכפיל הזה לכפיל שהיה: ה-payload עובר דרך הוולידטורים
      // של superdoc עצמו, ולא דרך `return true`.
      const verdict = checkPayload(id, payload);
      if (!verdict.accepted) {
        rejected.push({ id, payload, route: verdict.route });
        return {
          ok: false,
          message: `המנוע דחה את ה-payload של ${id} (${verdict.route})`,
          reason: 'payload-rejected',
        };
      }

      applied.push({ id, payload });

      if (held.has(id)) {
        await new Promise<void>((resolve) => {
          const queue = waiting.get(id) ?? [];
          queue.push(resolve);
          waiting.set(id, queue);
        });
      }

      const failure = failures[id];
      if (failure) return { ok: false, message: reasonText(failure), reason: failure };
      return { ok: true };
    },

    payloads(id) {
      return applied.filter((call) => call.id === id).map((call) => call.payload);
    },

    setState(id, patch) {
      const next: CommandState = { ...stateOf(id), ...patch };
      states.set(id, next);
      for (const listener of listeners.get(id) ?? []) listener(next);
    },

    release(id) {
      const next = waiting.get(id)?.shift();
      if (!next) throw new Error(`אין תשובה מושהית לשחרור עבור ${id}`);
      next();
    },

    reset() {
      calls.length = 0;
      applied.length = 0;
      rejected.length = 0;
      blocked.length = 0;
    },
  };

  return double;
}

/* ------------------------------------------------------------------ */
/* כפיל המופע (Document API)                                          */
/* ------------------------------------------------------------------ */

export interface DocCall {
  op: string;
  input?: unknown;
}

export interface SuperdocDoubleOptions {
  /** פעולות שהמנוע מדווח כלא-זמינות. משמש לבדוק פקד מנוטרל ואת ה-tooltip שלו. */
  denied?: readonly string[];
  /** מסלולים שאינם קיימים בפאסדה בכלל — גרסת מנוע שאין לה את היכולת. */
  missing?: readonly string[];
  /** מסלול שנכשל, עם קוד הכשל שהקבלה מחזירה. */
  failures?: Record<string, { code: string; message?: string }>;
  /**
   * מסלולים שהתשובה שלהם נפתרת **מעבר לגבול macrotask**, כמו במנוע האמיתי.
   *
   * למה זה נדרש: הכפיל פותר כל קריאה באותו tick, ולכן שום בדיקה אינה יכולה
   * להעמיד מצב שבו קריאה אחת עדיין באוויר בזמן שהמשתמש לוחץ פקד אחר. זה
   * בדיוק החלון שבו קריאת אימות (`footnotes.get`) מתיישנת — נמדד ~10ms
   * במסמך ריק וקר — ובלי ההשהיה כאן המסלול הזה אינו נגיש לבדיקה בכלל.
   *
   * התשובה מחושבת **ברגע הקריאה** ונמסרת אחר כך, כמו במנוע: מה שהקורא
   * מקבל הוא תצלום של המסמך כפי שהיה כשהוא שאל, ולא כפי שהוא בזמן התשובה.
   * חישוב מאוחר היה מסתיר את המרוץ במקום למדוד אותו.
   */
  deferred?: readonly string[];
  /** מה שהבחירה במסמך מדווחת. */
  selection?: { blockId?: string | null; hasRange?: boolean; text?: string };
  /** קישורים קיימים ש-`hyperlinks.list` ידווח (גל 22). */
  hyperlinks?: readonly Record<string, unknown>[];
  /**
   * מה שהמקטע במסמך **כבר** נושא. שני השדות האלה הם היחידים שהתוסף קורא
   * מהמקטע ואז שולח בחזרה, ולכן הם היחידים שכפיל חייב לדעת לייצר: מקטע
   * שיש בו `countBy: 5` הוא בדיוק המקרה שבו „מספרי שורות → רציף” שאינו
   * משמר מוחק את מה שנקבע ב-Word.
   */
  /**
   * מה שהמקטע במסמך **כבר** נושא. שני השדות האלה הם היחידים שהתוסף קורא
   * מהמקטע ואז שולח בחזרה, ולכן הם היחידים שכפיל חייב לדעת לייצר: מקטע
   * שיש בו `countBy: 5` הוא בדיוק המקרה שבו „מספרי שורות → רציף” שאינו
   * משמר מוחק את מה שנקבע ב-Word.
   */
  sections?: {
    lineNumbering?: {
      enabled?: boolean;
      countBy?: number;
      start?: number;
      distance?: number;
      restart?: string;
    };
    pageNumbering?: { start?: number; format?: string };
    /**
     * מידות נייר ב**אינצ'ים**, בצורה שהמנוע האמיתי מפרויקט (ראו engine/print.ts:
     * הפרויקציה הציבורית היא twips/1440). ברירת המחדל היא A4 — והיא **הייתה**
     * צורת ה-twips הגולמית, בהערה שאף צרן אינו קורא את `width`. מאז יש אחד:
     * `layOutTocRows` גוזר את מיקום עצירת הטאב מרוחב אזור הטקסט, וכפיל
     * שמחזיר 11906 „אינצ'ים” היה נותן לו עמוד ברוחב שמונה מטרים.
     */
    pageSize?: { width?: number; height?: number };
    /**
     * שולי הדף ב**אינצ'ים**, כמו `pageSize` לידו. ברירת המחדל היא השוליים
     * של מסמך חדש; `null` מחסיר אותם במפורש כדי לדמות מסמך שעדיין נטען.
     */
    margins?: { top?: number; right?: number; bottom?: number; left?: number } | null;
    /**
     * `sectionDirection` של המקטע — מה שפסקה שאינה מצהירה `<w:bidi>` יורשת.
     * ברירת המחדל היא RTL של מסמך חדש; `null` מחסיר את ההצהרה.
     */
    direction?: 'rtl' | 'ltr' | null;
  };
  /**
   * תכונות הפסקה שבה הסמן, במודל SDM/1 — **נקודות**, לא twips. אלה מה ש-
   * `doc.get` מחזיר ל-`readParagraphFormat`, ולכן הדיאלוג „פסקה” ממלא את
   * עצמו מהם. ברירת המחדל היא פסקה בלי עיצוב ישיר.
   */
  paragraphProps?: Record<string, unknown>;
  /**
   * תוכן העניינים שבמסמך. ברירת המחדל היא מסמך **בלי** טבלה — ראו ההסבר
   * ליד מסלולי `toc` למטה — ומי שבודק את הקבוצה מעמיד כאן מסמך שיש בו אחת.
   */
  toc?: {
    /** מזהי הטבלאות. יותר מאחד = המצב הדו-משמעי שההסרה מסרבת לו. */
    ids?: readonly string[];
    /** כמה שורות תוצאה נוספות יש בכל טבלה, מעבר לבלוק הראשון. */
    rowsPerToc?: number;
    /** ערכי `TC` שסומנו במסמך. */
    entries?: readonly { nodeId: string; text: string; level: number }[];
  };
  /**
   * המפתח שבמסמך. ברירת המחדל היא מסמך **בלי** מפתח ובלי ערכים מסומנים —
   * אותה החלטה כמו ב-`toc`, ומאותו טעם: זה המצב שבו „עדכן מפתח” אמור לומר
   * למה הוא אינו יכול, ומסמך שכבר יש בו מפתח היה מסתיר את ההבדל בין „נוצר”
   * ל„היה שם”.
   */
  index?: {
    /** מזהי המפתחות. יותר מאחד = המצב הדו-משמעי שההסרה מסרבת לו. */
    ids?: readonly string[];
    /** מספר הטורים של המפתח הראשון, כפי ש-`index.list` מצהיר עליו. */
    columns?: number;
    runIn?: boolean;
    /** ערכי `XE` שסומנו במסמך. הכתובת מיקומית, כמו במנוע האמיתי. */
    entries?: readonly { blockId: string; offset: number; text: string; subEntry?: string }[];
  };
  /**
   * הציטוטים שבמסמך. אותה החלטה כמו ב-`toc` וב-`index`: מסמך **בלי** מקורות,
   * בלי ציטוטים ובלי ביבליוגרפיה. זה המצב שבו „עדכן ביבליוגרפיה” אמור לומר
   * למה הוא אינו יכול, ומסמך שכבר יש בו אחת היה מסתיר את ההבדל בין „נוצרה”
   * ל„הייתה שם”.
   */
  citations?: {
    /** המקורות בחלק הביבליוגרפיה. */
    sources?: readonly { sourceId: string; type?: string; title?: string; year?: string }[];
    /** שדות `CITATION` שבמסמך, כל אחד והמקור שהוא מפנה אליו. */
    cited?: readonly string[];
    /**
     * מזהי הבלוקים של הביבליוגרפיות. יותר מאחד = המצב הדו-משמעי שההסרה
     * מסרבת לו. מוצגים דרך `fields.list` ולא דרך `blocks.list`, בדיוק כמו
     * במנוע האמיתי — ראו engine/citations.ts.
     */
    bibliographyIds?: readonly string[];
  };
  /**
   * ההערות שבמסמך. אותה החלטה כמו ב-`toc` וב-`captions`: מסמך **בלי** הערות.
   *
   * `type` נמסר לכל הערה, ו-`noteId` **אינו** — הוא מוקצה כאן לכל סוג בנפרד
   * ומתחיל מ-1, כמו במנוע האמיתי. זה מה שמייצר את ההתנגשות שנמדדה: הערת
   * שוליים והערת סיום שנושאות את אותו מספר ואת אותה כתובת בדיוק. כפיל
   * שהיה מקצה מזהה אחד רץ לשני הסוגים לא היה יודע לייצר את המסלול הזה
   * בכלל — וזה המסלול שבו „הסר” על הערת סיום מוחק הערת שוליים.
   */
  notes?: {
    items?: readonly { type: 'footnote' | 'endnote'; content: string }[];
    /**
     * כמה הערות עמוד אחד מחזיר, וכשהוא נמסר כל שאיבה שאחריו **זורקת** —
     * כשל אמצע-שאיבה. `total` ממשיך לדווח את הכול, כלומר זה המצב שבו רשימה
     * חלקית עלולה להיראות מלאה. הכפיל אינו יכול לייצר אותו אחרת: המנוע
     * שואב ב-200, ומסמך של מאתיים הערות בבדיקת קומפוננטה הוא רעש.
     */
    pageLimit?: number;
  };
  /**
   * הכיתובים שבמסמך, בסדר הופעתם. אותה החלטה כמו ב-`toc` וב-`citations`:
   * מסמך **בלי** כיתובים. המספור אינו נמסר כאן אלא מחושב ב-`captions.list`
   * מסדר הפריטים — בדיוק כמו במנוע האמיתי, שבו כיתוב שנוסף באמצע מזיז את
   * המספרים שאחריו. כפיל שהיה מקבל את המספר מבחוץ לא היה מודד את זה.
   */
  captions?: {
    /**
     * `tableBefore` שם לפני הכיתוב בלוק `tbl:*` מסוג `table`, ו-`tableAfter`
     * אחריו — הצורה השכיחה של כיתוב במסמך אמיתי („טבלה 1: …” שמתחת ללוח),
     * ומאז 2.15.0-next.15 גם העוגן שנשלח בפועל. `tocBefore`/`tocAfter` שמים
     * שם בלוק `tableOfContents`, שהוא הסוג שהמנוע עדיין דוחה — ובלעדיו אין
     * מה שיוכיח שהסירוב שנשאר עדיין חוסם. כפיל של פסקאות בלבד לא היה מודד
     * אף אחד מהמסלולים האלה.
     */
    items?: readonly {
      nodeId: string;
      label: string;
      text?: string;
      tableBefore?: boolean;
      tableAfter?: boolean;
      tocBefore?: boolean;
      tocAfter?: boolean;
    }[];
  };
}

export interface SuperdocDouble {
  remainingBlocks: () => string[];
  /** מה שמוזרק ל-`ACTIVE_SUPERDOC`. אינו SuperDoc אמיתי, ומטופס כך בכוונה. */
  readonly host: SuperDoc;
  readonly calls: DocCall[];
  /** הקלטים שמסלול מסוים קיבל, לפי הסדר. */
  inputs(op: string): unknown[];
  ops(): string[];
  /**
   * משנה את מה שהבחירה תדווח מהקריאה הבאה ואילך.
   *
   * הבחירה במסמך אינה קבועה, וממשק שקורא אותה יותר מפעם אחת אינו ניתן
   * לבדיקה בלי זה — ראו ההערה ליד `hasRange` ב-`createSuperdocDouble`.
   */
  setSelection(next: { hasRange?: boolean; text?: string }): void;
  reset(): void;
}

export function createSuperdocDouble(options: SuperdocDoubleOptions = {}): SuperdocDouble {
  const calls: DocCall[] = [];
  const denied = new Set(options.denied ?? []);
  const missing = new Set(options.missing ?? []);
  const deferred = new Set(options.deferred ?? []);
  const failures = options.failures ?? {};

  const blockId = options.selection?.blockId === undefined ? 'block-1' : options.selection.blockId;
  /*
   * הבחירה משתנה, ואינה קבועה בזמן ההרכבה.
   *
   * המנוע האמיתי מדווח את הבחירה **ברגע שנשאל**, והממשק שואל אותו יותר
   * מפעם אחת: „גופן מתקדם”, למשל, אינו חוסם את המסמך ושואל מחדש בכל שינוי
   * בחירה. כפיל שקופא על מה שנמסר לו בהרכבה אינו יכול להעמיד את המצב הזה
   * בכלל — ובלעדיו „המשתמש סימן טקסט בזמן שהדיאלוג פתוח” אינו ניתן לבדיקה.
   */
  let hasRange = options.selection?.hasRange ?? false;
  let selectionText = options.selection?.text ?? '';

  /**
   * זמינות הפעולות כ-Proxy. `readDocCapabilities` שואל מפה לפי שם הפעולה,
   * ורשימה כתובה כאן הייתה מתיישנת בכל פקד חדש — ואז הבדיקה של אותו פקד הייתה
   * מודדת אותו מנוטרל.
   */
  const operations = new Proxy(
    {},
    {
      get: (_target, key: string | symbol) =>
        typeof key === 'string' && denied.has(key)
          ? { available: false, reasons: ['OPERATION_UNAVAILABLE'] }
          : { available: true, reasons: [] },
    },
  );

  const globalFlags = new Proxy(
    {},
    {
      get: (_target, key: string | symbol) =>
        typeof key === 'string' && denied.has(key)
          ? { enabled: false, reasons: ['OPERATION_UNAVAILABLE'] }
          : { enabled: true, reasons: [] },
    },
  );

  /** מסלול שנרשם, מקליט, ומחזיר קבלה — או `undefined` אם הוא „חסר במנוע”. */
  function route<T>(op: string, impl: (input: unknown) => T): ((input: unknown) => T) | undefined {
    if (missing.has(op)) return undefined;
    return ((input: unknown) => {
      calls.push({ op, input });
      const value = impl(input);
      if (!deferred.has(op)) return value;
      // `setTimeout` ולא `Promise.resolve`: microtask נפתר לפני שהדפדפן
      // מספיק לעבד אירוע מצביע, ולכן הוא אינו מייצג את החלון שנמדד.
      return new Promise((resolve) => setTimeout(() => resolve(value), 0));
    }) as (input: unknown) => T;
  }

  function receipt(op: string): { success: boolean; failure?: { code: string; message?: string } } {
    const failure = failures[op];
    if (failure) return { success: false, failure };
    return { success: true };
  }

  const selectionTarget = {
    kind: 'text',
    story: { kind: 'body' },
    segments: blockId
      ? [{ blockId, range: { start: 0, end: hasRange ? selectionText.length || 4 : 0 } }]
      : [],
  };

  /**
   * ה-`SelectionTarget` — מודל אחר מ-`TextTarget` שמעל, ולא כינוי שלו: החוזה
   * קובע ש-`target` הוא ה-`TextTarget` לצריכה של תגובות, ו-`selectionTarget`
   * הוא „the public selection-target model the write APIs consume directly”.
   * `format.*` מקבל את השני, ולכן כפיל שמחזיר רק את הראשון היה מודד פקד
   * שנכשל סגור על „יש לסמן טקסט”.
   */
  const selectionEnvelope = {
    kind: 'selection',
    start: { kind: 'text', blockId, offset: 0 },
    end: { kind: 'text', blockId, offset: selectionText.length || 4 },
  };

  const tocIds = options.toc?.ids ?? [];
  const tocEntries = options.toc?.entries ?? [];
  const indexIds = options.index?.ids ?? [];
  const indexEntries = options.index?.entries ?? [];
  const citationSources = options.citations?.sources ?? [];
  const citedSourceIds = options.citations?.cited ?? [];
  const bibliographyIds = options.citations?.bibliographyIds ?? [];
  /**
   * ההערות שבמסמך. המזהה מוקצה לפי סוג ו**אינו** מחושב מחדש בהסרה: נמדד
   * שהמנוע משאיר את `displayNumber` של השאר כמו שהיה (הסרה של הערה 2
   * הותירה 1 ו-3), והמספר ש-Word יציג נקבע אצלו בפתיחה ולא כאן.
   */
  let footnoteSeq = 0;
  let endnoteSeq = 0;
  const notes = (options.notes?.items ?? []).map((item) => ({
    type: item.type,
    noteId: String(item.type === 'endnote' ? ++endnoteSeq : ++footnoteSeq),
    content: item.content,
  }));

  /** מונה מזהים לכיתוב שנוצר בכפיל. `nodeId` של פסקה הוא כתובת חולפת. */
  let captionSeq = 0;
  /** משתנה: `captions.insert` ו-`captions.remove` מוסיפים ומורידים ממנו. */
  const captionItems = (options.captions?.items ?? []).map((item) => ({
    nodeId: item.nodeId,
    label: item.label,
    text: item.text ?? '',
  }));
  const captionOptions = options.captions?.items ?? [];

  /**
   * הבלוקים שבמסמך, במבנה שהמנוע האמיתי מחזיק: **בלוק אחד** מסוג
   * `tableOfContents` לכל טבלה, ואחריו שורותיה כפסקאות `TOC1`. זה מה שהופך
   * את „הסר” לבדיקה אמיתית — ב-2.23 `toc.remove` מסיר את כל שורות השדה,
   * בלי שהאדפטר יצטרך למחוק פסקאות נוספות.
   */
  let blocks = tocIds
    .flatMap((nodeId) => [
      { nodeId, nodeType: 'tableOfContents', styleId: 'TOC1' },
      ...Array.from({ length: options.toc?.rowsPerToc ?? 0 }, (_, index) => ({
        nodeId: `${nodeId}-row-${index}`,
        nodeType: 'paragraph',
        styleId: 'TOC1',
      })),
    ])
    .concat([{ nodeId: 'block-1', nodeType: 'paragraph', styleId: 'Normal' }])
    // הכיתובים הם פסקאות ככל פסקה אחרת, ובאות **אחרי** `block-1`: זה מה
    // שהופך את „ערוך כיתוב” לבדיקה אמיתית — העריכה מחזירה את הכיתוב לפי
    // הבלוק שלפניו, וכיתוב ראשון ברשימה נשען על `block-1` ולא על עצמו.
    .concat(
      captionOptions.flatMap((item) => [
        ...(item.tableBefore
          ? [{ nodeId: `tbl:${item.nodeId}-before`, nodeType: 'table', styleId: '' }]
          : []),
        ...(item.tocBefore
          ? [{ nodeId: `toc:${item.nodeId}-before`, nodeType: 'tableOfContents', styleId: '' }]
          : []),
        { nodeId: item.nodeId, nodeType: 'paragraph', styleId: 'Caption' },
        ...(item.tableAfter
          ? [{ nodeId: `tbl:${item.nodeId}-after`, nodeType: 'table', styleId: '' }]
          : []),
        ...(item.tocAfter
          ? [{ nodeId: `toc:${item.nodeId}-after`, nodeType: 'tableOfContents', styleId: '' }]
          : []),
      ]),
    )
    .map((block, ordinal) => ({ ...block, ordinal }));

  const clipboardPayload = {
    source: 'superdoc',
    items: [{ type: 'text/plain', kind: 'string', data: selectionText || 'טקסט' }],
  };

  /**
   * הערת השוליים שמספרה `noteId`, ובהיעדרה הערת הסיום. זה הסדר שנמדד
   * במנוע — ולא בחירה של הכפיל.
   */
  const resolveNote = (noteId: string) =>
    notes.find((note) => note.type === 'footnote' && note.noteId === noteId) ??
    notes.find((note) => note.noteId === noteId);

  const doc = {
    capabilities: {
      get: route('capabilities.get', () => ({ operations, global: globalFlags })),
    },
    /**
     * המקטע שהכפיל מציג נושא גם את ההגדרות שנקראות ממנו, ולא רק את מידות
     * הדף: `lineNumbering` הוא מה ש„מספרי שורות” משמר במקום למחוק,
     * ו-`headerFooterMargins`/`pageNumbering` הם מה ששני הדיאלוגים נפתחים
     * עליו. כפיל שמחזיר מקטע בלי השדות האלה היה מאשר בירוק דיאלוג שנפתח
     * תמיד על ברירות המחדל — כלומר בדיוק הבאג שהקריאה נועדה למנוע.
     *
     * הערכים באינצ'ים, כמו שהמנוע האמיתי מחזיר: `0.5"` = 720 twips,
     * ו-`headerFooterMargins` של המסמך הריק הוא בדיוק זה.
     */
    sections: {
      list: route('sections.list', () => ({
        items: [
          {
            address: { sectionIndex: 0 },
            pageSetup: {
              width: options.sections?.pageSize?.width ?? 11906 / 1440,
              height: options.sections?.pageSize?.height ?? 16838 / 1440,
              orientation: 'portrait',
            },
            ...(options.sections?.margins === null
              ? {}
              : { margins: options.sections?.margins ?? { left: 1, right: 1, top: 1, bottom: 1 } }),
            ...(options.sections?.direction === null
              ? {}
              : { sectionDirection: options.sections?.direction ?? 'rtl' }),
            headerFooterMargins: { header: 0.5, footer: 0.5 },
            lineNumbering: options.sections?.lineNumbering,
            pageNumbering: options.sections?.pageNumbering,
            verticalAlign: 'top',
          },
        ],
      })),
      setPageMargins: route('sections.setPageMargins', () => receipt('sections.setPageMargins')),
      setPageSetup: route('sections.setPageSetup', () => receipt('sections.setPageSetup')),
      setColumns: route('sections.setColumns', () => receipt('sections.setColumns')),
      setTitlePage: route('sections.setTitlePage', () => receipt('sections.setTitlePage')),
      setOddEvenHeadersFooters: route('sections.setOddEvenHeadersFooters', () =>
        receipt('sections.setOddEvenHeadersFooters'),
      ),
      setLineNumbering: route('sections.setLineNumbering', () =>
        receipt('sections.setLineNumbering'),
      ),
      setVerticalAlign: route('sections.setVerticalAlign', () =>
        receipt('sections.setVerticalAlign'),
      ),
      setHeaderFooterMargins: route('sections.setHeaderFooterMargins', () =>
        receipt('sections.setHeaderFooterMargins'),
      ),
      setPageNumbering: route('sections.setPageNumbering', () =>
        receipt('sections.setPageNumbering'),
      ),
      setPageBorders: route('sections.setPageBorders', () => receipt('sections.setPageBorders')),
      clearPageBorders: route('sections.clearPageBorders', () =>
        receipt('sections.clearPageBorders'),
      ),
    },
    /**
     * המסמך של הכפיל נפתח **בלי** כותרות: `resolve` מחזיר `status: 'none'`
     * ו-`parts.list` ריק. זה המצב שהפקד נמדד בו — „אין כותרת, לחיצה יוצרת
     * אחת” — ולא מסמך שכבר יש בו כותרת, שבו „עריכה” היא no-op ולכן אינה
     * מוכיחה דבר.
     */
    headerFooters: {
      resolve: route('headerFooters.resolve', () => ({ status: 'none' })),
      refs: {
        set: route('headerFooters.refs.set', () => receipt('headerFooters.refs.set')),
        clear: route('headerFooters.refs.clear', () => receipt('headerFooters.refs.clear')),
        setLinkedToPrevious: route('headerFooters.refs.setLinkedToPrevious', () =>
          receipt('headerFooters.refs.setLinkedToPrevious'),
        ),
      },
      parts: {
        list: route('headerFooters.parts.list', () => ({ items: [] })),
        create: route('headerFooters.parts.create', () => {
          const result = receipt('headerFooters.parts.create');
          // הקבלה של `parts.create` נושאת גם את מזהה החלק; בלעדיו אין למה
          // להפנות, והקורא נכשל על „המנוע לא החזיר מזהה”.
          return result.success ? { ...result, refId: 'rId9', partPath: 'word/header1.xml' } : result;
        }),
        delete: route('headerFooters.parts.delete', () => receipt('headerFooters.parts.delete')),
      },
    },
    /**
     * הערות שוליים והערות סיום. שלוש הבחנות מכוונות, וכולן נמדדו במנוע
     * האמיתי:
     *
     * 1. **הכתובת אינה נושאת את הסוג.** `entityType` הוא `'footnote'` גם
     *    עבור הערת סיום, ולכן `get`/`update`/`remove` פותרים **תחילה** את
     *    הערת השוליים שמספרה זהה, ורק בהיעדרה את הערת הסיום. זה מה שהופך
     *    את האימות שלפני ההסרה לבדיקה אמיתית.
     * 2. **`get` זורק** על כתובת שאינה קיימת ואינו מחזיר קבלה.
     * 3. **`list` מכבד `limit`/`offset` באמת**, ומחזיר את שני הסוגים יחד
     *    עם `type` כמפריד היחיד ביניהם.
     */
    footnotes: {
      insert: route('footnotes.insert', (input) => {
        const failed = receipt('footnotes.insert');
        if (!failed.success) return failed;
        // ההוספה **נכנסת לרשימה**, ולא רק מחזירה קבלה: היא זו שמייצרת את
        // הכתובת המתנגשת. כפיל שמאשר בלי להוסיף היה מסתיר את מה שקורה
        // כשהוספה נכנסת בזמן שאימות של הסרה עדיין באוויר.
        const { type, content } = input as { type: 'footnote' | 'endnote'; content: string };
        notes.push({
          type,
          noteId: String(type === 'endnote' ? ++endnoteSeq : ++footnoteSeq),
          content,
        });
        return { success: true };
      }),
      list: route('footnotes.list', (input) => {
        const query = (input ?? {}) as { type?: string; limit?: number; offset?: number };
        const pageLimit = options.notes?.pageLimit;
        if (pageLimit !== undefined && (query.offset ?? 0) >= pageLimit) {
          throw new Error('footnotes.list failed.');
        }
        const matching = notes
          .filter((note) => query.type === undefined || note.type === query.type)
          .map((note) => ({
            id: note.noteId,
            handle: { ref: `footnote:${note.noteId}`, refStability: 'stable' },
            address: { kind: 'entity', entityType: 'footnote', noteId: note.noteId },
            type: note.type,
            noteId: note.noteId,
            displayNumber: note.noteId,
            content: note.content,
          }));
        const offset = query.offset ?? 0;
        const limit = Math.min(query.limit ?? matching.length, pageLimit ?? matching.length);
        return { items: matching.slice(offset, offset + limit), total: matching.length };
      }),
      get: route('footnotes.get', (input) => {
        const { target } = input as { target: { noteId: string } };
        const found = resolveNote(target.noteId);
        if (!found) throw new Error('footnote/endnote was not found.');
        return {
          address: { kind: 'entity', entityType: 'footnote', noteId: found.noteId },
          type: found.type,
          noteId: found.noteId,
          displayNumber: found.noteId,
          content: found.content,
        };
      }),
      update: route('footnotes.update', (input) => {
        const failed = receipt('footnotes.update');
        if (!failed.success) return failed;
        const { target, patch } = input as {
          target: { noteId: string };
          patch: { content: string };
        };
        const found = resolveNote(target.noteId);
        if (!found) {
          return {
            success: false,
            failure: { code: 'TARGET_NOT_FOUND', message: 'footnote/endnote was not found.' },
          };
        }
        // מחליף ואינו מוסיף — זה מה שנמדד, וזה ההבדל מ-`captions.update`.
        found.content = patch.content;
        return { success: true, footnote: target };
      }),
      remove: route('footnotes.remove', (input) => {
        const failed = receipt('footnotes.remove');
        if (!failed.success) return failed;
        const { target } = input as { target: { noteId: string } };
        const found = resolveNote(target.noteId);
        if (!found) {
          return {
            success: false,
            failure: { code: 'TARGET_NOT_FOUND', message: 'footnote/endnote was not found.' },
          };
        }
        notes.splice(notes.indexOf(found), 1);
        return { success: true, footnote: target };
      }),
    },
    /**
     * המסמך של הכפיל נפתח **בלי** שדות: `fields.list` ריק. זה המצב שהפקדים
     * נמדדים בו — „אין שדות, לחיצה מכניסה אחד” — ומסמך שכבר יש בו שדות היה
     * מסתיר את ההבדל בין „הוכנס” ל„היה שם”.
     */
    fields: {
      // הביבליוגרפיות מגיעות מכאן ולא מ-`blocks.list`, בדיוק כמו במנוע
      // האמיתי: אין ל-`citations.bibliography` פעולת `list`, והבלוק מופיע
      // ב-`blocks.list` כפסקה רגילה. ההנמקה ב-engine/citations.ts.
      list: route('fields.list', () => ({
        items: bibliographyIds.map((nodeId) => ({
          address: { kind: 'field', blockId: nodeId, occurrenceIndex: 0, nestingDepth: 0 },
          instruction: 'BIBLIOGRAPHY',
          fieldType: 'BIBLIOGRAPHY',
        })),
        total: bibliographyIds.length,
      })),
      insert: route('fields.insert', () => {
        const result = receipt('fields.insert');
        // הקבלה של `fields.insert` נושאת את כתובת השדה שנוצר; היא מה שמאפשר
        // את ה-rebuild שמחשב את התוצאה מיד אחרי ההכנסה.
        return result.success
          ? { ...result, field: { kind: 'field', blockId, occurrenceIndex: 0, nestingDepth: 0 } }
          : result;
      }),
      rebuild: route('fields.rebuild', () => receipt('fields.rebuild')),
    },
    /**
     * המסמך של הכפיל נפתח **בלי** סימניות: `bookmarks.list` ריק. זה המצב
     * שהפקד נמדד בו — „אין סימניות, הדיאלוג מוסיף אחת” — ומסמך שכבר יש בו
     * סימניות היה מסתיר את ההבדל בין „נוספה” ל„הייתה שם”.
     */
    bookmarks: {
      list: route('bookmarks.list', () => ({ items: [], total: 0 })),
      insert: route('bookmarks.insert', () => receipt('bookmarks.insert')),
      rename: route('bookmarks.rename', () => receipt('bookmarks.rename')),
      remove: route('bookmarks.remove', () => receipt('bookmarks.remove')),
    },
    /**
     * אותה החלטה כמו ב-`fields`: מסמך **בלי** תוכן עניינים ובלי ערכים
     * מסומנים. זה המצב שהפקדים נמדדים בו — „אין טבלה, ולכן הפעולה מדווחת
     * למה” — ומסמך שכבר יש בו טבלה היה מסתיר את ההבדל בין „נוצרה” ל„הייתה
     * שם”. בדיקת הקומפוננטה של הקבוצה מרכיבה כפיל משלה עם טבלה אחת.
     */
    toc: {
      list: route('toc.list', () => ({
        items: tocIds.map((nodeId) => ({
          address: { kind: 'block', nodeType: 'tableOfContents', nodeId },
          sourceConfig: {},
          displayConfig: { hyperlinks: true },
        })),
        total: tocIds.length,
      })),
      update: route('toc.update', () => receipt('toc.update')),
      remove: route('toc.remove', (input) => {
        const { target } = input as { target: { nodeId: string } };
        const result = receipt('toc.remove');
        if (result.success) {
          blocks = blocks.filter((block) =>
            block.nodeId !== target.nodeId && !block.nodeId.startsWith(`${target.nodeId}-row-`),
          );
        }
        return result;
      }),
      configure: route('toc.configure', () => receipt('toc.configure')),
      markEntry: route('toc.markEntry', () => receipt('toc.markEntry')),
      unmarkEntry: route('toc.unmarkEntry', () => receipt('toc.unmarkEntry')),
      listEntries: route('toc.listEntries', () => ({
        items: tocEntries.map((entry) => ({
          address: { kind: 'inline', nodeType: 'tableOfContentsEntry', nodeId: entry.nodeId },
          text: entry.text,
          level: entry.level,
        })),
        total: tocEntries.length,
      })),
    },
    /**
     * המפתח. שלושה הבדלים מכוונים מ-`toc` שמעל, וכולם נמדדו במנוע האמיתי:
     * `index.insert` מקבל `at` ו-`config` (ולא כתובת של עצם קיים), הכתובת של
     * ערך `XE` היא **עוגן מיקומי** ולא `nodeId`, ו-`index.remove` מפיל את
     * הבלוק כולו — ולכן אין כאן שיירים ואין תלות ב-`blocks.*`.
     * ההנמקה המלאה ב-engine/index-field.ts.
     */
    index: {
      list: route('index.list', () => ({
        items: indexIds.map((nodeId) => ({
          id: nodeId,
          address: { kind: 'block', nodeType: 'index', nodeId },
          config: { columns: options.index?.columns, runIn: options.index?.runIn === true },
          entryCount: indexEntries.length,
        })),
        total: indexIds.length,
      })),
      insert: route('index.insert', () => receipt('index.insert')),
      configure: route('index.configure', () => receipt('index.configure')),
      rebuild: route('index.rebuild', () => receipt('index.rebuild')),
      remove: route('index.remove', () => receipt('index.remove')),
      entries: {
        list: route('index.entries.list', () => ({
          items: indexEntries.map((entry) => ({
            id: `${entry.blockId}#${entry.offset}`,
            address: {
              kind: 'inline',
              nodeType: 'indexEntry',
              anchor: {
                start: { blockId: entry.blockId, offset: entry.offset },
                end: { blockId: entry.blockId, offset: entry.offset + 1 },
              },
            },
            text: entry.text,
            subEntry: entry.subEntry,
          })),
          total: indexEntries.length,
        })),
        insert: route('index.entries.insert', () => receipt('index.entries.insert')),
        remove: route('index.entries.remove', () => receipt('index.entries.remove')),
      },
    },
    /** רשימת הבלוקים מאפשרת לאמת תוצאות פקודות מול מצב המסמך. */
    blocks: {
      list: route('blocks.list', (input) => {
        const query = (input ?? {}) as { limit?: number; offset?: number };
        const offset = query.offset ?? 0;
        const end = query.limit === undefined ? undefined : offset + query.limit;
        return { blocks: blocks.slice(offset, end), total: blocks.length };
      }),
      deleteRange: route('blocks.deleteRange', (input) => {
        const { start, end } = input as { start: { nodeId: string }; end: { nodeId: string } };
        const from = blocks.findIndex((block) => block.nodeId === start.nodeId);
        const to = blocks.findIndex((block) => block.nodeId === end.nodeId);
        if (from !== -1 && to !== -1) blocks.splice(from, to - from + 1);
        return receipt('blocks.deleteRange');
      }),
    },
    /**
     * כיתובים. שלוש הבחנות מכוונות, וכולן נמדדו במנוע האמיתי:
     *
     * 1. **המספור מחושב ולא נמסר.** כל תווית מנהלת רצף משלה, לפי סדר
     *    הכיתובים במסמך. זה מה שהופך „הוסף באמצע” ו„הסר” לבדיקות.
     * 2. **`list` מכבד `limit`/`offset` באמת.** engine/captions.ts שואב
     *    עמודים עד `total`, וכפיל שמחזיר תמיד את הכול היה מאשר בירוק
     *    שאיבה שאינה קיימת.
     * 3. **`captions.update` אינו כאן בכלל.** המודול אינו קורא לו — הוא
     *    מוסיף את הטקסט החדש על הישן במקום להחליף אותו — ומסלול בכפיל
     *    לפעולה שאיש אינו קורא לה הוא הזמנה לקרוא לה.
     */
    captions: {
      list: route('captions.list', (input) => {
        const query = (input ?? {}) as { label?: string; limit?: number; offset?: number };
        const counters = new Map<string, number>();
        const numbered = captionItems.map((item) => {
          const next = (counters.get(item.label) ?? 0) + 1;
          counters.set(item.label, next);
          return {
            id: item.nodeId,
            address: { kind: 'block', nodeType: 'paragraph', nodeId: item.nodeId },
            label: item.label,
            number: next,
            text: item.text,
            instruction: `SEQ ${item.label} \\* ARABIC`,
          };
        });
        const matching =
          query.label === undefined
            ? numbered
            : numbered.filter((item) => item.label === query.label);
        const offset = query.offset ?? 0;
        const end = query.limit === undefined ? undefined : offset + query.limit;
        return { items: matching.slice(offset, end), total: matching.length };
      }),
      insert: route('captions.insert', (input) => {
        const failed = receipt('captions.insert');
        if (!failed.success) return failed;
        const { adjacentTo, position, label, text } = input as {
          adjacentTo: { nodeId: string; nodeType?: string };
          position: 'above' | 'below';
          label: string;
          text?: string;
        };
        const at = blocks.findIndex((block) => block.nodeId === adjacentTo.nodeId);
        // הכתובת חייבת לשאת את הסוג **האמיתי** של הבלוק, וזה מה שנמדד על
        // 2.15.0-next.15: טבלה שנשלחה כ-`paragraph`, ופסקה שנשלחה כ-`table`,
        // מוחזרות שתיהן `TARGET_NOT_FOUND` באותו נוסח.
        if (at === -1 || blocks[at].nodeType !== adjacentTo.nodeType) {
          return {
            success: false,
            failure: {
              code: 'TARGET_NOT_FOUND',
              message: `target ${adjacentTo.nodeType ?? 'paragraph'} ${adjacentTo.nodeId} was not found`,
            },
          };
        }
        // פסקה וטבלה עליונה, ותו לא: בלוק `tableOfContents` נמדד `INVALID_TARGET`
        // גם כשהכתובת נושאת את סוגו, בנוסח של המנוע עצמו.
        if (blocks[at].nodeType !== 'paragraph' && blocks[at].nodeType !== 'table') {
          return {
            success: false,
            failure: {
              code: 'INVALID_TARGET',
              message: `create.paragraph only supports before/after placement relative to paragraph-like blocks or top-level tables; received ${blocks[at].nodeType}.`,
            },
          };
        }
        const nodeId = `caption-${++captionSeq}`;
        const index = position === 'above' ? at : at + 1;
        blocks.splice(index, 0, { nodeId, nodeType: 'paragraph', styleId: 'Caption', ordinal: index });
        // הכיתובים נשמרים בסדר שבו הם יושבים בבלוקים, ולא בסדר ההוספה:
        // המספור נגזר מהמסמך, ורשימה שמסודרת אחרת הייתה ממציאה אותו.
        const before = blocks
          .slice(0, index)
          .filter((block) => captionItems.some((item) => item.nodeId === block.nodeId)).length;
        captionItems.splice(before, 0, { nodeId, label, text: text ?? '' });
        return { success: true, caption: { kind: 'block', nodeType: 'paragraph', nodeId } };
      }),
      remove: route('captions.remove', (input) => {
        const failed = receipt('captions.remove');
        if (!failed.success) return failed;
        const { target } = input as { target: { nodeId: string } };
        const index = captionItems.findIndex((item) => item.nodeId === target.nodeId);
        if (index === -1) {
          return {
            success: false,
            failure: { code: 'TARGET_NOT_FOUND', message: `caption "${target.nodeId}" was not found.` },
          };
        }
        captionItems.splice(index, 1);
        blocks = blocks.filter((block) => block.nodeId !== target.nodeId);
        return { success: true, caption: target };
      }),
    },
    /**
     * `paragraphs.setDirection` — הצעד שמסובב את פסקת הכיתוב לימין-לשמאל.
     * המנוע האמיתי כותב פסקת `Caption` **בלי** `<w:bidi/>` (נמדד ב-docx),
     * והמודול מוסיף אותו. כאן הוא רק מוקלט: מה שנבדק הוא שהוא נשלח, ושכשל
     * שלו אינו מפיל את הכיתוב.
     */
    paragraphs: {
      setDirection: route('paragraphs.setDirection', () => receipt('paragraphs.setDirection')),
    },
    /**
     * ציטוטים. שתי הבחנות מכוונות מול השכנים, ושתיהן נמדדו במנוע האמיתי:
     * המקורות הם **ישויות** ולא בלוקים (`kind: 'entity'`), והביבליוגרפיה
     * נמצאת דרך `fields.list` בלבד. ההנמקה המלאה ב-engine/citations.ts.
     */
    citations: {
      list: route('citations.list', () => ({
        items: citedSourceIds.map((sourceId) => ({
          address: { kind: 'inline', nodeType: 'citation' },
          sourceIds: [sourceId],
          instruction: `CITATION ${sourceId}`,
        })),
        total: citedSourceIds.length,
      })),
      insert: route('citations.insert', () => receipt('citations.insert')),
      sources: {
        list: route('citations.sources.list', () => ({
          items: citationSources.map((source) => ({
            id: source.sourceId,
            address: {
              kind: 'entity',
              entityType: 'citationSource',
              sourceId: source.sourceId,
            },
            sourceId: source.sourceId,
            tag: source.title ?? source.sourceId,
            type: source.type ?? 'book',
            fields: { title: source.title ?? '', year: source.year ?? '' },
          })),
          total: citationSources.length,
        })),
        insert: route('citations.sources.insert', () => receipt('citations.sources.insert')),
        update: route('citations.sources.update', () => receipt('citations.sources.update')),
        remove: route('citations.sources.remove', () => receipt('citations.sources.remove')),
      },
      bibliography: {
        insert: route('citations.bibliography.insert', () =>
          receipt('citations.bibliography.insert'),
        ),
        rebuild: route('citations.bibliography.rebuild', () =>
          receipt('citations.bibliography.rebuild'),
        ),
        remove: route('citations.bibliography.remove', () =>
          receipt('citations.bibliography.remove'),
        ),
      },
    },
    /** אותה החלטה כמו ב-`fields`: מסמך בלי הפניות מקושרות. */
    crossRefs: {
      list: route('crossRefs.list', () => ({ items: [], total: 0 })),
      rebuild: route('crossRefs.rebuild', () => receipt('crossRefs.rebuild')),
    },
    format: {
      paragraph: {
        setFlowOptions: route('format.paragraph.setFlowOptions', () =>
          receipt('format.paragraph.setFlowOptions'),
        ),
        // גל 11 — תפריט „פסקה”. אותן חתימות של המנוע: יעד + מצב מלא, והכשל
        // מוכרע לפי `failures` בדיוק כמו בכל מסלול אחר.
        setIndentation: route('format.paragraph.setIndentation', () =>
          receipt('format.paragraph.setIndentation'),
        ),
        clearIndentation: route('format.paragraph.clearIndentation', () =>
          receipt('format.paragraph.clearIndentation'),
        ),
        setSpacing: route('format.paragraph.setSpacing', () => receipt('format.paragraph.setSpacing')),
        clearSpacing: route('format.paragraph.clearSpacing', () => receipt('format.paragraph.clearSpacing')),
        setKeepOptions: route('format.paragraph.setKeepOptions', () =>
          receipt('format.paragraph.setKeepOptions'),
        ),
        setTabStop: route('format.paragraph.setTabStop', () => receipt('format.paragraph.setTabStop')),
        clearTabStop: route('format.paragraph.clearTabStop', () => receipt('format.paragraph.clearTabStop')),
        clearAllTabStops: route('format.paragraph.clearAllTabStops', () =>
          receipt('format.paragraph.clearAllTabStops'),
        ),
      },
      vertAlign: route('format.vertAlign', () => receipt('format.vertAlign')),
      // גל 12 — „גופן מתקדם". החתימה: { target: SelectionTarget, inline }.
      apply: route('format.apply', () => receipt('format.apply')),
    },
    // גל 14א — פעולות רשימה. היעד: ListItemAddress; הכשל לפי `failures`.
    lists: {
      setLevelNumberStyle: route('lists.setLevelNumberStyle', () =>
        receipt('lists.setLevelNumberStyle'),
      ),
      applyStyle: route('lists.applyStyle', () => receipt('lists.applyStyle')),
      restartAt: route('lists.restartAt', () => receipt('lists.restartAt')),
      continuePrevious: route('lists.continuePrevious', () => receipt('lists.continuePrevious')),
      convertToText: route('lists.convertToText', () => receipt('lists.convertToText')),
    },
    // גל 22 — ניהול קישורים. list מחזיר stories; remove לפי within.
    hyperlinks: {
      list: route('hyperlinks.list', () => ({
        stories: [{ storyId: 'main', hyperlinks: options.hyperlinks ?? [] }],
      })),
      wrap: route('hyperlinks.wrap', () => receipt('hyperlinks.wrap')),
      remove: route('hyperlinks.remove', () => receipt('hyperlinks.remove')),
    },
    // גל 19 — הגנת מסמך. get מחזיר מצב; set/clear לפי `failures`.
    protection: {
      get: route('protection.get', () => ({
        editingRestriction: { mode: 'none', enforced: false },
      })),
      setEditingRestriction: route('protection.setEditingRestriction', () =>
        receipt('protection.setEditingRestriction'),
      ),
      clearEditingRestriction: route('protection.clearEditingRestriction', () =>
        receipt('protection.clearEditingRestriction'),
      ),
    },
    /**
     * `doc.get` — המסמך במודל SDM/1, מצומצם לבלוק אחד שהבחירה מצביעה עליו.
     * `readParagraphFormat` הוא הקורא היחיד כרגע, והוא מחפש `id === nodeId`
     * ואת ה-props תחת המפתח של סוג הבלוק.
     */
    getNodeById: route('getNodeById', () => ({
      node: { id: blockId, kind: 'paragraph', paragraph: { inlines: [], props: options.paragraphProps ?? {} } },
      address: { kind: 'block', nodeType: 'paragraph', nodeId: blockId },
    })),
    getNode: route('getNode', (input) => ({
      node: { id: blockId, kind: 'paragraph', paragraph: { inlines: [], props: options.paragraphProps ?? {} } },
      address: input,
    })),
    get: route('get', () => ({
      body: [
        {
          id: blockId,
          kind: 'paragraph',
          paragraph: { inlines: [], props: options.paragraphProps ?? {} },
        },
      ],
    })),
    selection: {
      current: route('selection.current', () => ({
        empty: !hasRange,
        target: selectionTarget,
        // נמסר רק כשיש טווח: המנוע מקרין `null` כשאין מה להקרין, וכתיבה
        // דורשת טווח.
        selectionTarget: hasRange ? selectionEnvelope : null,
        text: selectionText,
      })),
    },
    clipboard: {
      serializeSelection: route('clipboard.serializeSelection', () => ({
        payload: clipboardPayload,
      })),
      parse: route('clipboard.parse', () => ({
        success: true,
        plan: { fragment: { blocks: [] }, diagnostics: [] },
      })),
      insert: route('clipboard.insert', () => receipt('clipboard.insert')),
    },
    delete: route('delete', () => receipt('delete')),
    ranges: {
      resolve: route('ranges.resolve', () => ({ target: selectionTarget })),
    },
    // גל 13 — „ברירות מחדל למסמך". dryRun נתמך ומחזיר before, כמו במנוע.
    styles: {
      apply: route('styles.apply', (input) => {
        const options = (input as { options?: { dryRun?: boolean } }).options;
        if (options?.dryRun) {
          return { success: true, changed: false, before: { fontSize: 24 }, after: {}, dryRun: true };
        }
        return receipt('styles.apply');
      }),
    },
    insert: route('insert', () => receipt('insert')),
  };

  const host = {
    activeEditor: { doc },
    /**
     * `superdoc.focus()` — הדרך היחידה להחזיר את הסמן לטקסט, ומה ש-`F6`
     * ו-`Escape` נשענים עליה. הכפיל חושף אותה כדי שהמסלול הזה יהיה נמדד ולא
     * מונח.
     */
    focus: route('focus', () => undefined),
    ui: {
      selection: {
        getSnapshot: () => {
          calls.push({ op: 'ui.selection.getSnapshot' });
          return {
            status: 'ready',
            empty: !hasRange,
            selectionTarget: hasRange ? selectionTarget : null,
          };
        },
        apply: route('ui.selection.apply', () => ({ ok: true })),
      },
    },
  };

  return {
    host: host as unknown as SuperDoc,
    remainingBlocks: () => blocks.map((block) => block.nodeId),
    calls,
    inputs: (op) => calls.filter((call) => call.op === op).map((call) => call.input),
    ops: () => calls.map((call) => call.op),
    /** משנה את מה שהבחירה תדווח מהקריאה הבאה ואילך — ראו ההערה ליד `hasRange`. */
    setSelection: (next: { hasRange?: boolean; text?: string }) => {
      if (next.hasRange !== undefined) hasRange = next.hasRange;
      if (next.text !== undefined) selectionText = next.text;
    },
    reset: () => {
      calls.length = 0;
    },
  };
}

/* ------------------------------------------------------------------ */
/* ההרכבה                                                             */
/* ------------------------------------------------------------------ */

/**
 * מונה `DOCUMENT_GENERATION` (composables/keys.ts) לכפילי הבדיקה.
 *
 * מודול-רמה ולא מקומי ל-`mountUi`: `PageBreakTracker` (engine/page-break.ts)
 * הוא מופע יחיד המשותף לכל התכנית — כולל בין קבצי בדיקה שרצים באותו worker —
 * ובלי מונה שעולה בכל הרכבה, כל מסמכי הבדיקה היו משתפים את אותה ברירת מחדל
 * `0` מ-`inject(DOCUMENT_GENERATION, shallowRef(0))`, ו-`syncDocument`
 * (שמשווה לפי ערך) לא הייתה מזהה "מסמך אחר" אחרי ההרכבה הראשונה בכל התהליך —
 * בדיוק הבאג שה-QA שהוסיף את המונה הזה נועד למנוע. עולה בכל `mountUi` **וגם**
 * בכל `setSuperdoc`, ובכוונה שמרנית מדי (גם `setSuperdoc(sameDouble)` מעלה):
 * איפוס-יתר של המעקב המקומי הוא בדיוק הכשל הבטוח (`isOn` חוזר `false`),
 * ואיפוס-חסר הוא זה שהיה שקט ומטעה.
 */
let documentGenerationCounter = 0;

export interface ReportedOutcome {
  outcome: CommandOutcome;
  commandId: string;
}

export interface HarnessOptions {
  /** ברירת המחדל: כפיל אדפטר חדש עם כל הפקודות זמינות. `null` = אין מנוע. */
  adapter?: CommandDouble | null;
  /** ברירת המחדל: כפיל מופע עם כל היכולות. `null` = אין מסמך פתוח. */
  superdoc?: SuperdocDouble | null;
  fontOptions?: FontOptions;
  /**
   * הזיכרון של בוררי הגופן. ברירת המחדל היא זיכרון חדש לכל הרכבה — כמו
   * ברירת המחדל של ה-inject עצמו. בדיקה שמוסרת **אותו** זיכרון לשתי הרכבות
   * מודדת בדיוק את מה שהמעטפת עושה: רצועה ותפריט הקשר שחולקים מצב.
   */
  fontMemory?: FontMemory;
  styleGallery?: StyleGalleryState;
  /** ‏`styles.xml` שכפיל גיליון הסגנונות מחזיר. ברירת המחדל: `STYLE_SHEET_DOUBLE`. */
  stylesXml?: string;
  /** הקיצורים האישיים בהרכבה. */
  customShortcuts?: CustomShortcut[];
  /**
   * מצב הבחירה שהחזקת החיווי נשענת עליו. ברירת המחדל היא סמן שהתיישב —
   * המצב שרצועה במסמך פתוח נמצאת בו רוב הזמן. ראו engine/readout-hold.ts.
   */
  readoutSelection?: ReadoutSelection;
  /**
   * השומר לפעולה כבדה, כפי שהמעטפת מזריקה אותו. ברירת המחדל היא הפעולה
   * כמות שהיא — אבל היא **מוזרקת**, כי פקד שנשען על מפתח שלא הוזרק נופל
   * לברירת המחדל של ה-inject ונראה עובד.
   */
  heavyActionGuard?: HeavyActionGuard;
  props?: Record<string, unknown>;
}

/** סמן במסמך, קריאה שהתיישבה. */
export const SETTLED_CARET: ReadoutSelection = { empty: true, settled: true };

export interface Harness {
  wrapper: VueWrapper;
  /** העריכות שנשלחו לגיליון הסגנונות, בסדר שבו נשלחו. */
  styleEdits: { edit: StyleEdit; options?: { applyToSelection?: boolean } }[];
  /** הקיצורים האישיים אחרי מה שהרכיב הצמיד או הסיר — אותה לוגיקה של המעטפת. */
  shortcuts: { list: CustomShortcut[] };
  adapter: CommandDouble;
  superdoc: SuperdocDouble;
  /** מה שהמדווח קיבל — המסלול שמגיע למשתמש כהודעה בעברית. */
  reports: ReportedOutcome[];
  /** הכשלים בלבד. */
  failures(): ReportedOutcome[];
  /** מחליפה את המסמך הפעיל, כמו פתיחת מסמך חדש במעטפת. */
  setSuperdoc(next: SuperdocDouble | null): Promise<void>;
  /** מחליפה את האדפטר, כמו החלפת session — מאפסת את זיכרון החיווי. */
  setAdapter(next: CommandDouble | null): Promise<void>;
  /** מזיזה את הבחירה: סמן/טווח, התיישבה או לא. */
  setReadoutSelection(next: ReadoutSelection): Promise<void>;
}

/** גיליון סגנונות מינימלי: „רגיל”, „כותרת 1”, וסגנון של המשתמש. */
export const STYLE_SHEET_DOUBLE =
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:next w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Mine"><w:name w:val="שלי"/><w:basedOn w:val="Normal"/></w:style>' +
  '</w:styles>';

/**
 * מרכיבה קומפוננטה עם בדיוק מה שהמעטפת מזריקה, ולא פחות: פקד שנשען על מפתח
 * שלא הוזרק נופל לברירת המחדל של ה-inject ונראה עובד — וזה באג שקט שבדיקה
 * חייבת לא לחזור עליו.
 */
export function mountUi(component: Component, options: HarnessOptions = {}): Harness {
  const adapter = options.adapter === undefined ? createCommandDouble() : options.adapter;
  const superdoc = options.superdoc === undefined ? createSuperdocDouble() : options.superdoc;

  const adapterRef: Ref<CommandAdapter | null> = shallowRef(adapter ?? null);
  const superdocRef = shallowRef<SuperDoc | null>(superdoc ? superdoc.host : null);
  const reports: ReportedOutcome[] = [];

  const provide: Record<symbol, unknown> = {};
  provide[COMMAND_ADAPTER as unknown as symbol] = adapterRef;
  provide[COMMAND_REPORTER as unknown as symbol] = (
    outcome: CommandOutcome,
    commandId: string,
  ): void => {
    reports.push({ outcome, commandId });
  };
  provide[FONT_OPTIONS as unknown as symbol] = ref(options.fontOptions ?? fallbackFontOptions());
  provide[FONT_MEMORY as unknown as symbol] = options.fontMemory ?? createFontMemory();
  provide[STYLE_GALLERY as unknown as symbol] = shallowRef(
    options.styleGallery ?? fallbackStyleGallery(),
  );
  provide[ACTIVE_SUPERDOC as unknown as symbol] = superdocRef;
  provide[HEAVY_ACTION_GUARD as unknown as symbol] =
    options.heavyActionGuard ?? ((action: () => Promise<unknown>) => action());
  /**
   * מערכת המאקרו מלווה כל מסמך פתוח (App.vue קובע אותה יחד עם
   * `activeSuperdoc`), ולכן הכפיל עוקב אחרי אותו כלל: יש מופע — יש מערכת.
   * הכפתורים ברצועה קוראים ממנה רק את `recording`; ה-kit עצמו אינו נצרך
   * בהרכבת לשונית, ולכן אינו ממומש — בדיקת הדיאלוג בונה kit אמיתי בעצמה.
   */
  const macrosStub: MacrosHandle = {
    kit: undefined as unknown as MacrosHandle['kit'],
    scriptsEnabled: true,
    recording: shallowRef(false),
    toggleRecording: () => undefined,
    replayLast: () => undefined,
    dispose: () => undefined,
  };
  provide[ACTIVE_MACROS as unknown as symbol] = shallowRef<MacrosHandle | null>(
    superdoc ? macrosStub : null,
  );
  const readoutSelectionRef = shallowRef<ReadoutSelection>(
    options.readoutSelection ?? SETTLED_CARET,
  );
  provide[READOUT_SELECTION as unknown as symbol] = readoutSelectionRef;
  const documentGenerationRef = shallowRef((documentGenerationCounter += 1));
  provide[DOCUMENT_GENERATION as unknown as symbol] = documentGenerationRef;

  /**
   * מתג בדיקת האיות. כפיל **מתפקד** ולא no-op: `toggle` מהפך את `enabled`
   * בדיוק כמו המעטפת, וזה מה שמאפשר לשער „אין כפתור מת” (ribbon-tabs) למדוד
   * את הכפתור הזה — לחיצה שאינה משנה דבר בכפיל הייתה נראית שם כפקד שבור.
   * המילון עצמו אינו נטען כאן: זו הרכבת רכיב, ואין בה DOM של מסמך לסמן.
   */
  const spellcheckEnabled = ref(false);
  provide[SPELLCHECK as unknown as symbol] = {
    enabled: spellcheckEnabled,
    busy: ref(false),
    toggle: () => {
      spellcheckEnabled.value = !spellcheckEnabled.value;
    },
  } satisfies SpellcheckHandle;

  /**
   * גיליון הסגנונות. כפיל שרושם ולא מחיל: הכתיבה עצמה (ייצוא, `styles.xml`,
   * פתיחה מחדש) היא של המעטפת ונבדקת ב-tests/unit/docx-style-*.test.ts.
   * בלי מסמך — `null`, כמו במעטפת.
   */
  const styleEdits: Harness['styleEdits'] = [];
  provide[STYLE_DOCUMENT as unknown as symbol] = {
    readStyles: async () => (superdocRef.value ? options.stylesXml ?? STYLE_SHEET_DOUBLE : null),
    applyEdit: async (edit, editOptions) => {
      styleEdits.push({ edit, options: editOptions });
      return { ok: true };
    },
  } satisfies StyleDocumentHandle;

  /**
   * קיצורי הסגנונות: כפיל **מתפקד** על אותן פונקציות שהמעטפת קוראת להן
   * (`upsertShortcut`, `comboConflict`), כדי שהתנגשות תיבדק כמו באמת.
   */
  const shortcuts = { list: [...(options.customShortcuts ?? [])] };
  provide[STYLE_SHORTCUTS as unknown as symbol] = {
    comboFor: (styleId) => styleShortcutFor(shortcuts.list, styleId)?.combo ?? null,
    problem: (styleId, combo) => {
      const conflict = comboConflict(combo, shortcuts.list, styleShortcutFor(shortcuts.list, styleId)?.id);
      return conflict ? conflictMessage(conflict) : null;
    },
    assign: (style, combo) => {
      const existing = styleShortcutFor(shortcuts.list, style.id);
      if (combo === null) {
        if (existing) shortcuts.list = removeShortcut(shortcuts.list, existing.id);
        return null;
      }
      const result = upsertShortcut(shortcuts.list, {
        ...(existing ? { id: existing.id } : {}),
        name: style.name,
        combo,
        preset: {},
        style,
      });
      if (!result.ok) return result.message;
      shortcuts.list = result.list;
      return null;
    },
    renamed: (styleId, name) => {
      shortcuts.list = shortcuts.list.map((entry) =>
        entry.kind === 'apply-style' && entry.style.id === styleId ? { ...entry, name, style: { id: styleId, name } } : entry,
      );
    },
    removed: (styleId) => {
      const existing = styleShortcutFor(shortcuts.list, styleId);
      if (existing) shortcuts.list = removeShortcut(shortcuts.list, existing.id);
    },
  } satisfies StyleShortcutsHandle;

  const wrapper = mount(component, {
    props: options.props,
    attachTo: document.body,
    global: { provide },
  });

  return {
    wrapper,
    styleEdits,
    shortcuts,
    // הכפילים מוחזרים גם כשלא הוזרקו, כדי שבדיקה תוכל לאשר שלא נגעו בהם.
    adapter: adapter ?? createCommandDouble(),
    superdoc: superdoc ?? createSuperdocDouble(),
    reports,
    failures: () => reports.filter((report) => !report.outcome.ok),
    async setSuperdoc(next) {
      superdocRef.value = next ? next.host : null;
      documentGenerationRef.value = documentGenerationCounter += 1;
      await settle();
    },
    async setAdapter(next) {
      adapterRef.value = next;
      await settle();
    },
    async setReadoutSelection(next) {
      readoutSelectionRef.value = next;
      await settle();
    },
  };
}

/**
 * מספר ה-events שהקומפוננטה **פלטה** — בלי אירועי DOM שרק עברו דרכה.
 *
 * @vue/test-utils רושם ב-`emitted()` לא רק `emit` של הקומפוננטה אלא גם כל
 * אירוע DOM מקורי שעולה לשורש שלה (חיקוי של fallthrough listeners; ראו
 * `recordEvent` ב-VTU). התוצאה: כל `trigger('click')` מופיע שם כ-`click` —
 * גם כשנלחץ כפתור שאין לו מטפל בכלל. שער „אין כפתור מת” שסופר אותו היה
 * מאשר בירוק בדיוק את מה שהוא נבנה לתפוס.
 */
/**
 * כפתור לפי הטקסט שהטולטיפ שלו מתחיל בו.
 *
 * ## למה שני שדות ולא אחד
 *
 * אותו prop `tooltip` נוחת בשדה אחר לפי הפקד: בכפתור אייקון הוא ה**כותרת**
 * („מודגש”), ובכפתור שיש לו תווית גלויה הוא יורד ל**הסבר** („העתק עיצוב ממקום
 * אחד…”, „הפעולה אינה זמינה בגרסה הזאת של המנוע”) והכותרת נשארת התווית. זה
 * הכלל של `tipParts`, והוא נכון — אבל הוא אומר שבדיקה שמחפשת „הפקד שהטולטיפ
 * שלו אומר X” חייבת להסתכל בשניהם. אחרת כל שינוי של `label` בפקד היה מפיל
 * בדיקות שאינן נוגעות בו.
 *
 * ## למה תחילת המחרוזת
 *
 * זכר לתקופה שבה ה-`title` נשא גם את הצירוף בסוגריים. היום `data-tip-shortcut`
 * נפרד, אבל ההתאמה לפי תחילית עדיין מונעת נעילה על סופי מחרוזות שנוטים לזוז.
 *
 * ## ולמה לא `title`
 *
 * התכונה הזאת אינה קיימת עוד באף אלמנט בתוכנה — היא מה שצייר טולטיפ שני, אפור,
 * מעל הכרטיס המעוצב. ראו tests/unit/native-title.test.ts.
 */
export function findButtonByTip(
  wrapper: VueWrapper,
  tipPrefix: string,
): DOMWrapper<Element> | undefined {
  return wrapper.findAll('button').find((button) => {
    const attributes = button.attributes();
    return (
      (attributes['data-tip-title'] ?? '').startsWith(tipPrefix) ||
      (attributes['data-tip-desc'] ?? '').startsWith(tipPrefix)
    );
  });
}

/** כמו `findButtonByTip`, ונופלת עם שם הכפתור כשאין כזה. */
export function buttonByTip(wrapper: VueWrapper, tipPrefix: string): DOMWrapper<Element> {
  const button = findButtonByTip(wrapper, tipPrefix);
  if (!button) throw new Error(`אין כפתור שהטולטיפ שלו מתחיל ב"${tipPrefix}"`);
  return button;
}

/**
 * סלקטור CSS לפקד לפי הטולטיפ שלו — לשימוש ב-`find`/`findAll` ישירים.
 *
 * שני השדות מאותה סיבה בדיוק כמו ב-`findButtonByTip`.
 */
export function tipSelector(tip: string, tag = 'button'): string {
  return `${tag}[data-tip-title="${tip}"],${tag}[data-tip-desc="${tip}"]`;
}

/** אותו דבר לפי תחילית, למקרים שהטקסט המלא נושא גם מספר או מידה שזזים. */
export function tipStartsSelector(prefix: string, tag = 'button'): string {
  return `${tag}[data-tip-title^="${prefix}"],${tag}[data-tip-desc^="${prefix}"]`;
}

/**
 * הבוררים ברצועה, בלי לדעת מאיזה סוג הם.
 *
 * שני מימושים חיים זה לצד זה: `<select>` נייטיב (מרווח שורות), ובורר
 * שאפשר להקליד בו — `<input role="combobox">` (גופן וגודל גופן; ראו
 * RibbonCombo.vue).
 * לבדיקה שמודדת „מה הבורר מציג” ו„מה קורה כשבוחרים” ההבדל בין השניים אינו
 * העניין, ובלי העטיפה הזאת כל מעבר בין המימושים היה מפיל אותה מחדש.
 */
function pickerOf(wrapper: VueWrapper, tip: string): DOMWrapper<Element> {
  const picker = wrapper.find(
    `select[data-tip-title="${tip}"],input[role="combobox"][data-tip-title="${tip}"]`,
  );
  if (!picker.exists()) throw new Error(`אין בורר עם הטולטיפ „${tip}”`);
  return picker;
}

/** מה שהבורר מציג בפועל ב-DOM. */
export function pickerValue(wrapper: VueWrapper, tip: string): string {
  return (pickerOf(wrapper, tip).element as HTMLSelectElement | HTMLInputElement).value;
}

/**
 * בחירה בבורר, בדרך שהמשתמש עובר בה.
 *
 * ב-`<select>` זו השמה ואירוע `change`. בבורר החיפוש אין רשימה קבועה לבחור
 * ממנה: מקלידים ומאשרים ב-Enter — וזה גם מה שמפעיל את הדירוג, כלומר את אותו
 * מסלול שהמשתמש עובר בו ולא קיצור סביבו.
 */
export async function setPicker(
  wrapper: VueWrapper,
  tip: string,
  value: string,
): Promise<void> {
  const picker = pickerOf(wrapper, tip);
  if (picker.element.tagName === 'SELECT') {
    await picker.setValue(value);
    return;
  }
  await picker.trigger('focus');
  await picker.setValue(value);
  await picker.trigger('keydown', { key: 'Enter' });
}

/* ------------------------------------------------------------------ */
/* תפריטי הרצועה (RibbonMenuButton)                                    */
/* ------------------------------------------------------------------ */

/**
 * פותחת תפריט של פקד ברצועה לפי הטולטיפ שלו, ומחזירה את תוויות הפריטים.
 *
 * כאן ולא בקובץ בדיקה יחיד מפני שיש כבר שלושה צרכנים (רשימות, שדות, מרווח
 * שורות), וכל אחד מהם כתב לעצמו את אותן שתי שורות. במצב מפוצל הלחיצה היא על
 * החץ ולא על הגוף — הגוף מפעיל פעולה — ולכן העוטפת מחפשת אותו קודם.
 */
export async function openRibbonMenu(wrapper: VueWrapper, tip: string): Promise<string[]> {
  const button = buttonByTip(wrapper, tip);
  const menu = button.element.closest('.ribbon-menu');
  // לחיצה על פקד פתוח **סוגרת** אותו. בדיקה שקוראת „מה מוצג” פעמיים בריצה
  // אחת הייתה מקבלת רשימה ריקה בפעם השנייה, ונראית ככשל של הפקד.
  if (!menu?.querySelector('.ribbon-menu__popover')) {
    const arrow = wrapper
      .findAll('.word-split__arrow')
      .find((node) => node.element.closest('.ribbon-menu') === menu);
    await (arrow ?? button).trigger('click');
    await settle();
  }
  return wrapper.findAll('.ribbon-menu__item-label').map((node) => node.text());
}

/** לוחצת על פריט בתפריט פתוח, לפי התווית שלו. */
export async function clickRibbonMenuItem(wrapper: VueWrapper, label: string): Promise<void> {
  const item = wrapper
    .findAll('.ribbon-menu__item')
    .find((node) => node.find('.ribbon-menu__item-label').text() === label);
  if (!item) throw new Error(`הפריט „${label}” לא נמצא בתפריט`);
  await item.trigger('click');
  await settle();
}

/**
 * התווית של הפריט המסומן בתפריט הפתוח, או `''` כשאין כזה.
 *
 * זה מה שהחליף את `pickerValue` למרווח השורות: הערך הנוכחי אינו כתוב בפקד
 * אלא מסומן בתוך התפריט, בדיוק כמו ב-Word.
 */
export function checkedRibbonMenuItem(wrapper: VueWrapper): string {
  const checked = wrapper.find('.ribbon-menu__item--checked');
  return checked.exists() ? checked.find('.ribbon-menu__item-label').text() : '';
}

const NO_TIP: TipContent = { title: '', shortcut: '', description: '' };

/**
 * שלושת שדות הטולטיפ של פקד, דרך הקוד שהתוכנה עצמה קוראת בו.
 *
 * `readTip` ולא קריאת תכונות ידנית: כך בדיקה שמאשרת „הכפתור אומר למה הוא
 * מנוטרל” מודדת בדיוק את מה שהמשתמש יראה, ולא ייצוג מקביל שעלול להתפצל ממנו.
 */
export function tipOf(target: Element | { element: Element }): TipContent {
  const element = target instanceof Element ? target : target.element;
  return readTip(element) ?? NO_TIP;
}

/**
 * שורת ההסבר של הפקד — ההסבר אם יש, אחרת הכותרת.
 *
 * למה איחוד ולא שדה אחד: אותו טקסט („הפעולה אינה זמינה בגרסה הזאת של המנוע”)
 * יורד להסבר בפקד שיש לו תווית גלויה, ונשאר כותרת בכפתור אייקון. זה הכלל של
 * `tipParts`, והוא נכון — אבל בדיקה ששואלת „מה הפקד אומר למשתמש” אינה אמורה
 * להיות תלויה בו.
 */
export function tipMessage(target: Element | { element: Element }): string {
  const tip = tipOf(target);
  return tip.description || tip.title;
}

export function emittedCount(wrapper: VueWrapper, ignore: readonly string[] = ['click']): number {
  return Object.entries(wrapper.emitted())
    .filter(([name]) => !ignore.includes(name))
    .reduce((total, [, occurrences]) => total + occurrences.length, 0);
}

/** מפרקת אוטומטית כל הרכבה בסוף בדיקה. נקראת פעם אחת בראש קובץ בדיקה. */
export function autoUnmount(): void {
  enableAutoUnmount(afterEach);
}

/* ------------------------------------------------------------------ */
/* לוח המערכת                                                          */
/* ------------------------------------------------------------------ */

/**
 * ל-jsdom אין `navigator.clipboard`, ובלעדיו כל פעולות הלוח נכשלות ב„לוח
 * המערכת חסם” — כלומר הבדיקה הייתה מודדת את מסלול החסימה ולא את הפעולה.
 */
export function installSystemClipboard(text = 'טקסט מהלוח'): () => void {
  const original = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  Object.defineProperty(navigator, 'clipboard', {
    configurable: true,
    value: {
      writeText: async () => {},
      readText: async () => text,
    },
  });
  return () => {
    if (original) Object.defineProperty(navigator, 'clipboard', original);
    // ב-jsdom המאפיין אינו קיים מראש, ולכן „החזרה למצב הקודם” היא הסרתו.
    else Reflect.deleteProperty(navigator, 'clipboard');
  };
}

/** לחיצה מחוץ לפקד, כפי שמאזין ה-`pointerdown` הגלובלי רואה אותה. */
export function clickOutside(): void {
  // jsdom אינו מממש PointerEvent, ו-MouseEvent בשם הזה מפעיל את אותו מאזין.
  document.body.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }));
}
