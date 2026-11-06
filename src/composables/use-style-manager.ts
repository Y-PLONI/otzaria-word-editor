/**
 * ניהול סגנונות מהגלריה — „שנה”, „צור”, „עדכן כך שיתאים לבחירה”, שינוי שם
 * לסגנון מותאם. הלשונית מחזיקה רק את הדיאלוג; כל השאר כאן.
 *
 * כל פעולה מתחילה בקריאת `styles.xml` הטרי (`readStyles`), ולא ממטמון:
 * „ברירות מחדל למסמך” ופעולות אחרות כותבות לגיליון בלי לעבור כאן, ודיאלוג
 * שנפתח על גיליון ישן היה כותב בחזרה ערכים שכבר אינם נכונים.
 */
import { computed, inject, shallowRef, watch } from 'vue';
import type { StyleCatalogItem } from 'superdoc/ui';
import {
  COMMAND_ADAPTER,
  COMMAND_REPORTER,
  STATUS_NOTIFIER,
  STYLE_DOCUMENT,
  STYLE_SHORTCUTS,
  type CommandReporter,
} from './keys';
import { comboSignature, type KeyCombo } from '../ui/shortcuts/combo';
import { ACTIVE_SUPERDOC } from '../engine/document-api';
import {
  listStyles,
  newStyleId,
  resolveFormat,
  type StyleEdit,
  type StyleFormat,
  type StyleInfo,
} from '../engine/docx-style-edit';
import { formatPatch, selectionStyleFormat, withoutThemedFont } from '../engine/style-from-selection';
import { readParagraphFormat } from '../engine/paragraph-format';
import { styleDisplayLabel } from '../engine/style-gallery';
import { restoreComplexScriptStyles } from '../engine/docx-cs-display';
import type { StyleAction } from '../ui/ribbon/common/StyleGallery.vue';
import type { StyleChoice, StyleDialogMode, StyleDialogResult } from '../ui/panels/StyleDialog.vue';
import type { CommandOutcome } from '../engine/command-adapter';
import { rememberStyleDefault, templateHasStyle, UNSUPPORTED_STYLE_DEFAULTS_MESSAGE } from './new-document-defaults';

export interface StyleDialogState {
  /** The originating instance; a dialog never follows the active tab. */
  origin: object;
  mode: StyleDialogMode;
  styleId: string | null;
  styleName: string;
  styleType: string;
  styles: readonly StyleChoice[];
  basedOn: string | null;
  next: string | null;
  format: StyleFormat;
  /** מה שהעיצוב יורש — הטלאי נכתב ביחס אליו. */
  baseline: StyleFormat;
  /** גיליון הסגנונות כפי שנקרא בפתיחה. */
  xml: string;
  /** קיצור המקשים של הסגנון בפתיחה. */
  shortcut: KeyCombo | null;
  /** הסגנון קיים בתבנית של מסמך חדש — ראו „גם לכל מסמך חדש” בדיאלוג. */
  newDocumentsOption: boolean;
}

function sameCombo(a: KeyCombo | null, b: KeyCombo | null): boolean {
  if (a === null || b === null) return a === b;
  return comboSignature(a) === comboSignature(b);
}

/** השם של הסגנון החדש הבא: „סגנון1”, „סגנון2”… — כמו ב-Word. */
export function nextStyleName(styles: readonly StyleInfo[]): string {
  const taken = new Set(styles.map((style) => style.name.trim().toLowerCase()));
  for (let index = 1; ; index++) {
    const candidate = `סגנון${index}`;
    if (!taken.has(candidate)) return candidate;
  }
}

function labelOf(style: StyleInfo): string {
  return styleDisplayLabel({ id: style.id, name: style.name, custom: style.custom } as StyleCatalogItem);
}

function choicesOf(styles: readonly StyleInfo[]): StyleChoice[] {
  return styles
    .filter((style) => style.type === 'paragraph' || style.type === 'character')
    .map((style) => ({ id: style.id, label: labelOf(style), type: style.type }));
}

const silentReporter: CommandReporter = () => {};

export function useStyleManager(activeStyleId: () => string | null) {
  const handle = inject(STYLE_DOCUMENT, null);
  const adapter = inject(COMMAND_ADAPTER, () => shallowRef(null), true);
  const superdoc = inject(ACTIVE_SUPERDOC, () => shallowRef(null), true);
  const report = inject(COMMAND_REPORTER, silentReporter);
  const notify = inject(STATUS_NOTIFIER, () => {});
  const shortcuts = inject(STYLE_SHORTCUTS, null);

  function comboProblem(styleId: string | null, combo: KeyCombo): string | null {
    return shortcuts ? shortcuts.problem(styleId ?? '', combo) : null;
  }

  const dialog = shallowRef<StyleDialogState | null>(null);
  const busy = shallowRef(false);
  const error = shallowRef<string | null>(null);
  watch(superdoc, () => close(), { flush: 'sync' });

  const available = computed(() => handle !== null && superdoc.value !== null);

  async function readSelection(): Promise<StyleFormat> {
    const paragraph = await readParagraphFormat(superdoc.value);
    return selectionStyleFormat(adapter.value, paragraph.ok ? paragraph : null);
  }

  async function run<T>(action: () => Promise<T>): Promise<T | undefined> {
    if (busy.value) return undefined;
    busy.value = true;
    try {
      return await action();
    } finally {
      busy.value = false;
    }
  }

  function fail(message: string, id: string): CommandOutcome {
    const outcome: CommandOutcome = { ok: false, message, reason: 'style-manager' };
    report(outcome, id);
    return outcome;
  }

  async function readSheet(): Promise<string | null> {
    const xml = (await handle?.readStyles()) ?? null;
    if (xml === null) fail('גיליון הסגנונות של המסמך אינו קריא', 'style-manager');
    return xml;
  }

  let appliedStyleId: string | undefined;
  async function apply(edit: StyleEdit, success: string, options?: { applyToSelection?: boolean }): Promise<boolean> {
    if (!handle) return false;
    appliedStyleId = undefined;
    const origin = superdoc.value;
    const outcome = await handle.applyEdit(edit, options);
    if (superdoc.value !== origin) return false;
    report(outcome, `style-${edit.kind}`);
    if (!outcome.ok) {
      error.value = outcome.message;
      return false;
    }
    appliedStyleId = outcome.styleId;
    notify(success);
    return true;
  }

  /** פעולה מהתפריט של הגלריה. */
  async function onAction(action: StyleAction, styleId: string | null): Promise<void> {
    if (!available.value) return;
    const origin = superdoc.value!;
    error.value = null;
    await run(async () => {
      const xml = await readSheet();
      if (xml === null || superdoc.value !== origin) return;
      const styles = listStyles(xml);
      const id = styleId ?? activeStyleId() ?? styles.find((s) => s.type === 'paragraph' && s.isDefault)?.id ?? null;
      const info = id ? styles.find((style) => style.id === id) : undefined;

      if (action === 'create') {
        const base = info && info.type === 'paragraph' ? info : styles.find((s) => s.type === 'paragraph' && s.isDefault);
        const baseline = base ? resolveFormat(xml, base.id) : {};
        const selection = await readSelection();
        if (superdoc.value !== origin) return;
        dialog.value = {
          origin,
          mode: 'create',
          styleId: null,
          styleName: nextStyleName(styles),
          styleType: 'paragraph',
          styles: choicesOf(styles),
          basedOn: base?.id ?? null,
          next: null,
          format: { ...baseline, ...withoutThemedFont(selection, baseline) },
          baseline,
          xml,
          shortcut: null,
          newDocumentsOption: false,
        };
        return;
      }

      if (!info) {
        fail('הסגנון אינו במסמך — ייתכן שהוא מוצע בגלריה ועדיין לא הוגדר', `style-${action}`);
        return;
      }
      const resolved = resolveFormat(xml, info.id);

      if (action === 'update') {
        const selection = await readSelection();
        if (superdoc.value !== origin) return;
        const patch = formatPatch(withoutThemedFont(selection, resolved), resolved);
        if (Object.keys(patch).length === 0) {
          notify(`„${labelOf(info)}” כבר תואם לבחירה`);
          return;
        }
        await apply({ kind: 'modify', id: info.id, format: patch }, `הסגנון „${labelOf(info)}” עודכן`);
        return;
      }

      const newDocumentsOption = action === 'modify' && await templateHasStyle(info.id);
      if (superdoc.value !== origin) return;
      dialog.value = {
        origin,
        newDocumentsOption,
        mode: action as StyleDialogMode,
        styleId: info.id,
        styleName: action === 'rename' ? info.name : labelOf(info),
        styleType: info.type,
        styles: choicesOf(styles),
        basedOn: info.basedOn,
        next: info.next ?? info.id,
        format: resolved,
        baseline: resolved,
        xml,
        shortcut: shortcuts?.comboFor(info.id) ?? null,
      };
    });
  }

  async function onSubmit(result: StyleDialogResult): Promise<void> {
    const state = dialog.value;
    if (!state || state.origin !== superdoc.value) { close(); return; }
    error.value = null;
    // Validate the shortcut before making any document transaction.
    const comboChanged = !sameCombo(result.combo, state.shortcut);
    if (comboChanged && result.combo) {
      const problem = comboProblem(state.styleId, result.combo);
      if (problem) {
        error.value = problem;
        return;
      }
    }
    const assignCombo = (id: string, name: string): void => {
      if (!comboChanged || !shortcuts) return;
      const failure = shortcuts.assign({ id, name }, result.combo);
      if (failure) fail(failure, 'style-shortcut');
    };

    // „גם לכל מסמך חדש”: העיצוב השלם שהדיאלוג החזיר, אחרי שהמסמך עצמו עודכן.
    const rememberForNewDocuments = async (): Promise<boolean> => {
      if (!result.forNewDocuments || !state.styleId) return true;
      // הדיאלוג מולא מהגיליון החי, שבו החריץ הלטיני מציג את העברי
      // (engine/docx-cs-display.ts). להגדרות נשמר מה שהסגנון אומר באמת, ועליו
      // רק מה שהמשתמש שינה — אחרת ערך שהוצג ולא נבחר הופך לברירת מחדל.
      const authored = resolveFormat(restoreComplexScriptStyles(state.xml) ?? state.xml, state.styleId);
      const remembered = { ...authored, ...formatPatch(result.format, state.baseline) };
      if (!await rememberStyleDefault(state.styleId, remembered)) {
        error.value = UNSUPPORTED_STYLE_DEFAULTS_MESSAGE;
        fail(UNSUPPORTED_STYLE_DEFAULTS_MESSAGE, 'style-new-documents');
        return false;
      }
      notify(`„${state.styleName}” יחול גם על כל מסמך חדש`);
      return true;
    };

    await run(async () => {
      let done = false;
      switch (state.mode) {
        case 'create': {
          const createdId = newStyleId(state.xml, result.name);
          // הטלאי ביחס לבסיס **שנבחר בדיאלוג**, ולא לזה שהדיאלוג נפתח עליו.
          const baseline = result.basedOn ? resolveFormat(state.xml, result.basedOn) : {};
          done = await apply(
            {
              kind: 'create',
              name: result.name,
              type: 'paragraph',
              basedOn: result.basedOn,
              next: result.next,
              format: formatPatch(result.format, baseline),
            },
            `נוצר הסגנון „${result.name}”`,
            { applyToSelection: true },
          );
          if (done) assignCombo(appliedStyleId ?? createdId, result.name);
          break;
        }
        case 'modify': {
          const changedBase = result.basedOn !== state.basedOn;
          const patch = formatPatch(result.format, state.baseline);
          const changedNext = state.styleType === 'paragraph' && result.next !== state.next;
          if (!changedBase && !changedNext && Object.keys(patch).length === 0) {
            // רק הקיצור השתנה — אין מה לכתוב למסמך, ואין טעינה מחדש.
            assignCombo(state.styleId as string, state.styleName);
            if (comboChanged) notify(`קיצור המקשים של „${state.styleName}” נשמר`);
            done = await rememberForNewDocuments();
            break;
          }
          done = await apply(
            {
              kind: 'modify',
              id: state.styleId as string,
              basedOn: changedBase ? result.basedOn : undefined,
              next: changedNext ? result.next : undefined,
              format: patch,
            },
            `הסגנון „${state.styleName}” עודכן`,
          );
          if (done) {
            assignCombo(state.styleId as string, state.styleName);
            done = await rememberForNewDocuments();
          }
          break;
        }
        case 'rename':
          done = result.name === state.styleName || await apply(
            { kind: 'rename', id: state.styleId as string, name: result.name },
            `שם הסגנון שונה ל„${result.name}”`,
          );
          if (done) shortcuts?.renamed(state.styleId as string, result.name);
          break;
      }
      if (done) dialog.value = null;
    });
  }

  function close(): void {
    dialog.value = null;
    error.value = null;
  }

  return { dialog, busy, error, available, onAction, onSubmit, close, comboProblem };
}
