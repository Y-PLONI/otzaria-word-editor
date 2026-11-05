/**
 * דיאלוג ניהול המאקרו, כפי שהמשתמש פוגש אותו.
 *
 * ה-kit כאן אמיתי (MacroKit של superdoc-macros, עם אחסון בזיכרון) והמארח
 * כפיל — כלומר מה שנבדק הוא בדיוק מה שהדיאלוג עושה: שמירה, בחירה, הרצה
 * וייבוא/ייצוא מול ה-API האמיתי, בלי מנוע. הדיאלוג מרונדר ב-Teleport לגוף
 * הדף, ולכן הבדיקות ניגשות אליו דרך ה-document — כמו LinkDialog.
 */
import { describe, expect, it, vi } from 'vitest';
import { shallowRef } from 'vue';
import { DOMWrapper } from '@vue/test-utils';
import {
  MacroKit,
  createMemoryStorage,
  type MacroHost,
  type MacroOutcome,
  type VbaModule,
} from 'superdoc-macros';
import MacrosDialog from '../../src/ui/panels/MacrosDialog.vue';
import type { MacrosHandle } from '../../src/engine/macros';
import { NO_VBA, WARNING_TEXT, type DocumentVba } from '../../src/engine/vba-import';
import { autoUnmount, mountUi, settle } from './harness';

autoUnmount();

/** מארח כפיל: מסמך כמחרוזת. מספיק לכל מה שהדיאלוג מפעיל. */
function createFakeHost(): MacroHost & { text: string } {
  const host = {
    text: '',
    commands: {
      has: (id: string) => id === 'bold',
      ids: () => ['bold'] as const,
      async execute(): Promise<MacroOutcome> {
        return { ok: true };
      },
    },
    async insertText(value: string): Promise<MacroOutcome> {
      host.text += value;
      return { ok: true };
    },
    async deleteBackward(count: number): Promise<MacroOutcome> {
      host.text = host.text.slice(0, -count);
      return { ok: true };
    },
    async deleteForward(): Promise<MacroOutcome> {
      return { ok: true };
    },
    async getSelection() {
      return { text: '', hasRange: false, blockId: null, selectionTarget: null, empty: true };
    },
    async replaceAll() {
      return { ok: true, replaced: 0 };
    },
    async getDocumentText() {
      return host.text;
    },
    onCommand: () => () => undefined,
    onTextInput: () => () => undefined,
  };
  return host;
}

function createHandle(options: { scriptsEnabled?: boolean } = {}): {
  handle: MacrosHandle;
  host: ReturnType<typeof createFakeHost>;
} {
  const host = createFakeHost();
  const kit = new MacroKit({ host, storage: createMemoryStorage(), runner: 'eval' });
  const handle: MacrosHandle = {
    kit,
    scriptsEnabled: options.scriptsEnabled ?? true,
    recording: shallowRef(false),
    toggleRecording: () => undefined,
    replayLast: () => undefined,
    dispose: () => undefined,
  };
  return { handle, host };
}

/** אלמנט מתוך ה-Teleport, כ-wrapper שאפשר ללחוץ עליו. */
function dialog(): DOMWrapper<Element> {
  const element = document.querySelector('.macros-dialog');
  if (!element) throw new Error('הדיאלוג אינו בגוף הדף');
  return new DOMWrapper(element);
}

function buttonByText(text: string): DOMWrapper<Element> {
  const button = dialog()
    .findAll('button')
    .find((candidate) => candidate.text() === text);
  if (!button) throw new Error(`אין כפתור "${text}"`);
  return button;
}

async function switchTab(title: string): Promise<void> {
  const tab = dialog()
    .findAll('[role="tab"]')
    .find((candidate) => candidate.text() === title);
  if (!tab) throw new Error(`אין לשונית "${title}"`);
  await tab.trigger('click');
  await settle();
}

describe('MacrosDialog', () => {
  it('בלי מסמך — הסבר במקום פקדים', async () => {
    mountUi(MacrosDialog, { props: { isOpen: true, handle: null } });
    await settle();

    expect(dialog().text()).toContain('יש לפתוח מסמך');
    expect(dialog().findAll('[role="tab"]')).toHaveLength(0);
  });

  it('קטע טקסט: הוספה, בחירה ומחיקה מול ה-kit', async () => {
    const { handle } = createHandle();
    mountUi(MacrosDialog, { props: { isOpen: true, handle } });
    await settle();
    await switchTab('קטעי טקסט');

    await dialog().find('#md-snip-name').setValue('חתימה');
    await dialog().find('#md-snip-text').setValue('בברכה');
    await dialog().find('#md-snip-trigger').setValue('חתמ');
    await buttonByText('הוסף').trigger('click');
    await settle();

    expect(handle.kit.listSnippets()).toEqual([
      expect.objectContaining({ name: 'חתימה', text: 'בברכה', trigger: 'חתמ' }),
    ]);
    // אחרי שמירה הפריט נבחר, והכפתור מתחלף ל„עדכן”.
    expect(buttonByText('עדכן').exists()).toBe(true);

    await buttonByText('מחק').trigger('click');
    await settle();
    expect(handle.kit.listSnippets()).toHaveLength(0);
  });

  it('כלים מובנים: הלשונית מופיעה רק כשנרשמו, הרצה וקיצור עובדים מול ה-kit', async () => {
    const { handle } = createHandle();
    let ran = 0;
    handle.kit.registerTool({
      id: 'shulchan.demo',
      name: 'כלי הדגמה',
      description: 'תיאור הכלי',
      run: () => {
        ran += 1;
        return { ok: true };
      },
    });
    mountUi(MacrosDialog, { props: { isOpen: true, handle } });
    await settle();
    await switchTab('כלים');

    const item = dialog()
      .findAll('.md-list-item')
      .find((candidate) => candidate.text().includes('כלי הדגמה'));
    expect(item).toBeDefined();
    await item!.trigger('click');
    await settle();
    expect(dialog().text()).toContain('תיאור הכלי');

    await buttonByText('הרץ').trigger('click');
    await settle();
    expect(ran).toBe(1);

    await dialog().find('#md-tool-shortcut').setValue('Ctrl+Alt+5');
    await buttonByText('שמור קיצור').trigger('click');
    await settle();
    expect(handle.kit.listTools()[0]?.shortcut).toBe('Ctrl+Alt+5');
  });

  it('בלי כלים רשומים — אין לשונית „כלים”', async () => {
    const { handle } = createHandle();
    mountUi(MacrosDialog, { props: { isOpen: true, handle } });
    await settle();
    const titles = dialog()
      .findAll('[role="tab"]')
      .map((tab) => tab.text());
    expect(titles).not.toContain('כלים');
  });

  it('לשונית הסקריפטים מוסתרת כשהדגל כבוי', async () => {
    const { handle } = createHandle({ scriptsEnabled: false });
    mountUi(MacrosDialog, { props: { isOpen: true, handle } });
    await settle();

    const titles = dialog()
      .findAll('[role="tab"]')
      .map((tab) => tab.text());
    expect(titles).toEqual(['הקלטות', 'קטעי טקסט', 'ייבוא וייצוא']);
  });

  it('קיצור שמתנגש בפריט שמור נחסם עם שם הפריט', async () => {
    const { handle } = createHandle();
    handle.kit.saveSnippet({ name: 'חתימה', text: 'בברכה', shortcut: 'Ctrl+Alt+5' });

    mountUi(MacrosDialog, { props: { isOpen: true, handle } });
    await settle();
    await switchTab('קטעי טקסט');

    await dialog().find('#md-snip-name').setValue('אחר');
    await dialog().find('#md-snip-text').setValue('טקסט');
    await dialog().find('#md-snip-shortcut').setValue('Ctrl+Alt+5');
    await settle();

    expect(dialog().find('[role="alert"]').text()).toContain('חתימה');
    expect((buttonByText('הוסף').element as HTMLButtonElement).disabled).toBe(true);
  });

  it('קיצור פסול חוסם שמירה ומציג שגיאה', async () => {
    const { handle } = createHandle();
    mountUi(MacrosDialog, { props: { isOpen: true, handle } });
    await settle();
    await switchTab('קטעי טקסט');

    await dialog().find('#md-snip-name').setValue('א');
    await dialog().find('#md-snip-text').setValue('ב');
    // `Ctrl+` — יש modifier ואין מקש; זה מה ש-parseShortcut דוחה.
    await dialog().find('#md-snip-shortcut').setValue('Ctrl+');
    await settle();

    expect(dialog().find('[role="alert"]').exists()).toBe(true);
    expect((buttonByText('הוסף').element as HTMLButtonElement).disabled).toBe(true);
    expect(handle.kit.listSnippets()).toHaveLength(0);
  });

  it('סקריפט: „הרץ” מריץ את מה שבעורך מול המסמך', async () => {
    const { handle, host } = createHandle();
    mountUi(MacrosDialog, { props: { isOpen: true, handle } });
    await settle();
    await switchTab('סקריפטים');

    await dialog().find('#md-scr-source').setValue(`await api.insertText('שלום');`);
    await buttonByText('הרץ').trigger('click');
    await settle();
    await settle();

    expect(host.text).toBe('שלום');
    expect(dialog().text()).toContain('המאקרו הסתיים בהצלחה');
  });

  it.each([
    ['הקלטות', 'נגן'], ['קטעי טקסט', 'הכנס למסמך'],
    ['סקריפטים', 'הרץ'], ['כלים', 'הרץ'],
  ])('דיאלוג פתוח אינו משנה מסמך שננעל (%s), והטופס נשמר עד השחרור', async (section, action) => {
    const { handle, host } = createHandle();
    const text = 'בדיקת ביצוע';
    handle.kit.importState(JSON.stringify({
      version: 1, scripts: [], snippets: [],
      recordings: [{ version: 1, id: 'pending', name: 'בדיקה', steps: [{ type: 'insert-text', text }] }],
    }));
    handle.kit.saveSnippet({ name: 'בדיקה', text });
    handle.kit.registerTool({
      id: 'pending', name: 'בדיקה', description: 'בדיקת נעילה',
      run: () => { host.text += text; return { ok: true }; },
    });
    let blocked = false;
    const { wrapper } = mountUi(MacrosDialog, {
      props: { isOpen: true, handle, busy: false, isDocumentBlocked: () => blocked },
    });
    await settle();
    await switchTab(section);
    const source = `await api.insertText('${text}');`;
    if (section === 'סקריפטים') await dialog().find('#md-scr-source').setValue(source);
    else await dialog().find('.md-list-item').trigger('click');
    await settle();

    // The live guard changes before a rendering update can disable the button.
    blocked = true;
    await buttonByText(action).trigger('click');
    await settle(6);
    expect(host.text).toBe('');
    await wrapper.setProps({ busy: true });
    expect((buttonByText(action).element as HTMLButtonElement).disabled).toBe(true);
    if (section === 'סקריפטים') expect((dialog().find('#md-scr-source').element as HTMLTextAreaElement).value).toBe(source);

    blocked = false;
    await wrapper.setProps({ busy: false });
    await buttonByText(action).trigger('click');
    await vi.waitFor(() => expect(host.text).toBe(text));
  });

  it('הקלטה: בחירה, ניגון ועדכון שם', async () => {
    const { handle, host } = createHandle();
    handle.kit.importState(
      JSON.stringify({
        version: 1,
        scripts: [],
        snippets: [],
        recordings: [
          {
            version: 1,
            id: 'rec-1',
            name: 'פתיח',
            steps: [{ type: 'insert-text', text: 'בס"ד' }],
          },
        ],
      })
    );

    const wrapper = mountUi(MacrosDialog, { props: { isOpen: true, handle } }).wrapper;
    await settle();

    const item = dialog()
      .findAll('[role="option"]')
      .find((candidate) => candidate.text().includes('פתיח'));
    expect(item).toBeDefined();
    await item!.trigger('click');
    await settle();

    await buttonByText('נגן').trigger('click');
    await settle();
    await settle();
    expect(host.text).toBe('בס"ד');
    expect(wrapper.emitted('status')).toBeTruthy();

    await dialog().find('#md-rec-name').setValue('פתיח דבר תורה');
    await buttonByText('עדכן').trigger('click');
    await settle();
    expect(handle.kit.listRecordings()[0]!.name).toBe('פתיח דבר תורה');
  });

  it('ייצוא ממלא את התיבה וייבוא ממזג', async () => {
    const { handle } = createHandle();
    handle.kit.saveSnippet({ name: 'בס"ד', text: 'בס"ד', trigger: 'בסד' });

    mountUi(MacrosDialog, { props: { isOpen: true, handle } });
    await settle();
    await switchTab('ייבוא וייצוא');

    await buttonByText('ייצא לכאן').trigger('click');
    await settle();
    const exported = (dialog().find('#md-transfer').element as HTMLTextAreaElement).value;
    expect(exported).toContain('בסד');

    // מיזוג של אותו ייצוא חזרה — לא מכפיל: פריט עם אותו מזהה מוחלף.
    await buttonByText('ייבא מכאן').trigger('click');
    await settle();
    expect(handle.kit.listSnippets()).toHaveLength(1);
  });
});

/**
 * לשונית „VBA במסמך”.
 *
 * מה שנבדק כאן הוא בעיקר מה ש**אינו** קורה: אין כפתור שמריץ, אין כפתור
 * שמייבא לתוך ה-kit, והקוד מוצג בתיבה לקריאה בלבד. זו לשונית שמתארת את הקובץ
 * שנפתח, לא יכולת של העורך.
 */
describe('MacrosDialog — VBA שבמסמך', () => {
  const vbaModule = (
    name: string,
    source: string,
    kind: VbaModule['kind'] = 'standard',
  ): VbaModule => ({ name, kind, source, truncated: false });

  const withModules = (overrides: Partial<DocumentVba> = {}): DocumentVba => ({
    hasMacroPart: true,
    modules: [
      vbaModule('Module1', 'Sub AutoOpen()\r\n    MsgBox "שלום"\r\nEnd Sub'),
      vbaModule('ThisDocument', 'Attribute VB_Name = "ThisDocument"', 'document'),
    ],
    autoRun: ['Module1.AutoOpen'],
    warnings: [WARNING_TEXT['auto-run-macros']],
    status: 'במסמך יש מאקרו',
    unreadable: false,
    ...overrides,
  });

  it('הלשונית אינה קיימת במסמך בלי מאקרו', async () => {
    const { handle } = createHandle();
    mountUi(MacrosDialog, { props: { isOpen: true, handle, documentVba: NO_VBA } });
    await settle();

    // לשונית ריקה על מסמך רגיל הייתה רק מבלבלת.
    const titles = dialog()
      .findAll('[role="tab"]')
      .map((tab) => tab.text());
    expect(titles).not.toContain('VBA במסמך');
  });

  it('מציגה את קוד המודול, לקריאה בלבד', async () => {
    const { handle } = createHandle();
    mountUi(MacrosDialog, { props: { isOpen: true, handle, documentVba: withModules() } });
    await settle();
    await switchTab('VBA במסמך');

    const source = dialog().find('#md-vba-source').element as HTMLTextAreaElement;
    expect(source.value).toContain('Sub AutoOpen()');
    // קריאה בלבד: הקוד הזה הגיע מקובץ חיצוני, ואין שום מסלול שעורך או מריץ
    // אותו כאן.
    expect(source.readOnly).toBe(true);
  });

  it('אומרת במפורש שהמאקרו אינם מורצים, ומזהירה על הרצה אוטומטית', async () => {
    const { handle } = createHandle();
    mountUi(MacrosDialog, { props: { isOpen: true, handle, documentVba: withModules() } });
    await settle();
    await switchTab('VBA במסמך');

    const text = dialog().text();
    expect(text).toContain('אינם מורצים כאן');
    // האזהרה הביטחונית: Word מריץ `AutoOpen` מעצמו, וכאן לא.
    expect(text).toContain(WARNING_TEXT['auto-run-macros']);
  });

  it('מחליפה מודול לפי הבחירה', async () => {
    const { handle } = createHandle();
    mountUi(MacrosDialog, { props: { isOpen: true, handle, documentVba: withModules() } });
    await settle();
    await switchTab('VBA במסמך');

    await dialog().find('#md-vba-module').setValue('ThisDocument');
    await settle();

    const source = dialog().find('#md-vba-source').element as HTMLTextAreaElement;
    expect(source.value).toContain('VB_Name = "ThisDocument"');
  });

  it('חלק מאקרו שלא נקרא — הלשונית קיימת ומסבירה', async () => {
    const { handle } = createHandle();
    const unreadable: DocumentVba = {
      hasMacroPart: true,
      modules: [],
      autoRun: [],
      warnings: ['במסמך יש מאקרו, אבל פרויקט המאקרו אינו קריא.'],
      status: null,
      unreadable: true,
    };

    mountUi(MacrosDialog, { props: { isOpen: true, handle, documentVba: unreadable } });
    await settle();
    await switchTab('VBA במסמך');

    // הלשונית קיימת דווקא כאן: זה המצב שבו למשתמש הכי חשוב לדעת שהקובץ שלו
    // נושא מאקרו — הם נשמרים, גם אם אי אפשר להראות אותם.
    expect(dialog().text()).toContain('אינו קריא');
    expect(dialog().find('#md-vba-source').exists()).toBe(false);
  });

  it('אין בלשונית שום כפתור שמריץ או מייבא', async () => {
    const { handle } = createHandle();
    mountUi(MacrosDialog, { props: { isOpen: true, handle, documentVba: withModules() } });
    await settle();
    await switchTab('VBA במסמך');

    const labels = dialog()
      .findAll('.md-footer button')
      .map((button) => button.text());
    // „סגור” בלבד. כל כפתור אחר כאן היה הופך טקסט מקובץ חיצוני למשהו שקורה.
    expect(labels.filter((label) => label !== 'סגור')).toEqual([]);
    expect(handle.kit.listScripts()).toEqual([]);
  });
});
