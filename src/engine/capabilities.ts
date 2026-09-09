/**
 * ה-registry של פקודות SuperDoc שהתוסף משתמש בהן.
 *
 * זו רשימה אחת ויחידה: ה-Ribbon ייבנה ממנה, ובדיקת החוזה
 * (tests/contract/superdoc-commands.test.ts) מריצה אותה מול המנוע. מזהה שהמנוע
 * לא מכיר מפיל את הבדיקה — ולא נשאר כפתור מת בממשק.
 *
 * `BUILT_IN_COMMAND_IDS` הציבורי מכסה 16 מזהים קנוניים בלבד, לא את הקטלוג
 * המלא, ו-`COMMAND_CATALOG` אינו export ציבורי. הגילוי הציבורי היחיד הוא
 * `ui.commands.ids` ו-`ui.commands.has(id)`.
 */

/**
 * המזהים לפי תחום. החלוקה היא לפי משפחת הפעולה במנוע, לא לפי לשונית ב-Ribbon:
 * לשונית שתשתמש בפקודה מכמה תחומים תרכיב אותה מכאן.
 */
export const COMMAND_GROUPS = {
  history: ['undo', 'redo'],
  character: [
    'bold',
    'italic',
    'underline',
    'strikethrough',
    'clear-formatting',
    'copy-format',
  ],
  font: ['font-family', 'font-size'],
  color: ['text-color', 'highlight-color'],
  paragraph: ['text-align', 'line-height', 'linked-style', 'direction-rtl', 'direction-ltr'],
  lists: ['bullet-list', 'numbered-list', 'indent-increase', 'indent-decrease'],
  insert: ['link', 'image', 'table-of-contents-insert', 'table-insert'],
  table: [
    'table-add-row-before',
    'table-add-row-after',
    'table-delete-row',
    'table-add-column-before',
    'table-add-column-after',
    'table-delete-column',
    'table-delete',
    'table-merge-cells',
    'table-split-cell',
    'table-remove-borders',
  ],
  review: [
    'acceptChange',
    'rejectChange',
    'acceptAllChanges',
    'rejectAllChanges',
    'track-changes-accept-selection',
    'track-changes-reject-selection',
  ],
  view: [
    'zoom',
    'zoom-fit-width',
    'ruler',
    'formatting-marks',
    'document-mode',
    'measurement-unit',
  ],
} as const;

export type CommandGroup = keyof typeof COMMAND_GROUPS;
export type CommandId = (typeof COMMAND_GROUPS)[CommandGroup][number];

export const COMMAND_IDS: readonly CommandId[] = Object.values(COMMAND_GROUPS).flat();

/**
 * פקודות ההקשר של טבלה. בלי תא חוקי המנוע מחזיר
 * `table-context-unavailable`, ולכן הפקד מוצג disabled ולא נעלם.
 */
export const TABLE_CONTEXT_COMMANDS: readonly CommandId[] = COMMAND_GROUPS.table;

/**
 * מזהים שקיימים בקטלוג אך המנוע מסמן כ-unsupported. הם לא נכנסים ל-registry
 * ולא מוצגים בממשק; הרשימה כאן כדי שבדיקת החוזה תשמור על ההנחה הזאת ותתריע
 * אם גרסה עתידית תתמוך בהם.
 */
export const KNOWN_UNSUPPORTED_COMMANDS = {
  'table-fix': 'תיקון טבלה אינו נתמך במנוע 2.13.0',
} as const;

/** צורת המינימום מ-`superdoc.ui` שנדרשת לבדיקת יכולות. */
export interface CommandProbe {
  commands: {
    has(id: string): boolean;
    get(id: string): { getState(): { source?: string } };
  };
}

export interface CommandSupportReport {
  /** המנוע אינו מכיר את המזהה בכלל. */
  unknown: string[];
  /** המזהה בקטלוג אך מסומן unsupported. */
  unsupported: string[];
  /** מזהים שהמנוע מנתב. */
  routed: string[];
}

/**
 * בודקת את ה-registry מול ה-controller.
 *
 * ההבחנה היא לפי `source` ולא לפי `supported`: לפני `onReady` המנוע מדווח
 * `supported: false` על כל פקודה, אבל `source` נשאר `builtin` לפקודה מנותבת
 * ו-`unsupported` לפקודה שאינה. כך אפשר לבדוק יכולות גם בזמן boot.
 */
export function inspectCommandSupport(
  ui: CommandProbe,
  ids: readonly string[] = COMMAND_IDS,
): CommandSupportReport {
  const report: CommandSupportReport = { unknown: [], unsupported: [], routed: [] };

  for (const id of ids) {
    if (!ui.commands.has(id)) {
      report.unknown.push(id);
      continue;
    }
    if (ui.commands.get(id).getState().source === 'unsupported') {
      report.unsupported.push(id);
      continue;
    }
    report.routed.push(id);
  }

  return report;
}
