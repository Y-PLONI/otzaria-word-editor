import { describe, expect, it, vi } from 'vitest';
import HomeTab from '../../src/ui/ribbon/tabs/HomeTab.vue';
import { autoUnmount, mountUi, settle } from './harness';
import { NO_THEME_FONTS, showComplexScriptStyles } from '../../src/engine/docx-cs-display';

const preference = vi.hoisted(() => ({ remember: vi.fn(async () => false) }));
vi.mock('../../src/composables/new-document-defaults', () => ({
  templateHasStyle: async () => true,
  rememberStyleDefault: preference.remember,
  UNSUPPORTED_STYLE_DEFAULTS_MESSAGE: 'עיצוב הסגנונות אינו נתמך במסמכים חדשים בלי הבדל בין העורך ל-Word; ההעדפה לא נשמרה',
}));
autoUnmount();

describe('remembering an unchanged imported style', () => {
  it('shows an unsupported-default refusal in the still-open dialog and reports no saved success', async () => {
    const harness = mountUi(HomeTab, { styleGallery: {
      items: [{ id: 'Normal', label: 'רגיל', previewText: 'AaBbCc', previewStyle: {}, custom: false }],
      activeId: 'Normal', fromDocument: true,
    } });
    await settle();
    await harness.wrapper.find('.style-card').trigger('contextmenu', { clientX: 500, clientY: 100 });
    await settle();
    document.querySelector<HTMLButtonElement>('[data-action="modify"]')!.click();
    await settle(4);
    const checkbox = document.querySelector<HTMLInputElement>('[data-testid="style-new-documents"]')!;
    expect(checkbox).not.toBeNull();
    checkbox.checked = true;
    checkbox.dispatchEvent(new Event('change', { bubbles: true }));
    await settle();
    document.querySelector<HTMLButtonElement>('[data-testid="style-dialog"] .sd-btn-primary')!.click();
    await settle(4);
    expect(preference.remember).toHaveBeenCalledOnce();
    expect(harness.styleEdits).toHaveLength(0);
    expect(document.querySelector('[data-testid="style-dialog"]')?.textContent).toContain('ההעדפה לא נשמרה');
    expect(harness.failures().some(item => item.commandId === 'style-new-documents')).toBe(true);
  });
});

describe('remembering a style from a sheet shown through docx-cs-display', () => {
  it('stores what the style authors (Courier New 20), not the Hebrew values the live sheet shows', async () => {
    preference.remember.mockClear();
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const authored = `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="David"/><w:sz w:val="22"/><w:szCs w:val="28"/></w:rPr></w:rPrDefault></w:docDefaults>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="40"/></w:rPr></w:style></w:styles>`;
    const live = showComplexScriptStyles(authored, NO_THEME_FONTS)!;
    const harness = mountUi(HomeTab, {
      stylesXml: live,
      styleGallery: {
        items: [{ id: 'Heading1', label: 'כותרת 1', previewText: 'AaBbCc', previewStyle: {}, custom: false }],
        activeId: 'Heading1', fromDocument: true,
      },
    });
    await settle();
    await harness.wrapper.find('.style-card').trigger('contextmenu', { clientX: 500, clientY: 100 });
    await settle();
    document.querySelector<HTMLButtonElement>('[data-action="modify"]')!.click();
    await settle(4);
    const checkbox = document.querySelector<HTMLInputElement>('[data-testid="style-new-documents"]')!;
    checkbox.checked = true;
    checkbox.dispatchEvent(new Event('change', { bubbles: true }));
    await settle();
    document.querySelector<HTMLButtonElement>('[data-testid="style-dialog"] .sd-btn-primary')!.click();
    await settle(4);
    expect(preference.remember).toHaveBeenCalledOnce();
    const [id, format] = preference.remember.mock.calls[0] as unknown as [string, Record<string, unknown>];
    expect(id).toBe('Heading1');
    expect(format).toMatchObject({ fontFamily: 'Courier New', fontSizePt: 20 });
  });
});
