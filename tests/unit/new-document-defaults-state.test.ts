// @vitest-environment node
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { deriveHebrewBlankDocx, readZip } from '../../scripts/blank-docx';
import { patchBlankDocumentXml, patchBlankStylesXml } from '../../src/engine/blank-document';
import { resolveFormat } from '../../src/engine/docx-style-edit';
import { readTemplateStyles } from '../../src/engine/new-document-defaults';

const host = vi.hoisted(() => ({ stored: null as unknown, save: vi.fn(), load: vi.fn() }));
vi.mock('../../src/host/settings', () => ({ loadNewDocumentDefaults: async () => { host.load(); return host.stored; }, saveNewDocumentDefaults: host.save }));
vi.mock('../../src/engine/blank-document-source', () => ({ blankDocumentSource: () => new Blob([new Uint8Array(DOCX)]) }));
const DOCX = deriveHebrewBlankDocx({ patchDocument: patchBlankDocumentXml, patchStyles: patchBlankStylesXml });
const TEMPLATE = readZip(DOCX).find(entry => entry.name === 'word/styles.xml')!.data.toString('utf8');

beforeEach(() => { vi.resetModules(); host.stored = null; host.save.mockReset(); host.load.mockClear(); });

describe('new document preference acceptance and migration', () => {
  it('shares an early pending load and waits for its defaults before creating the document', async () => {
    let release!: (value: unknown) => void;
    host.stored = new Promise(resolve => { release = resolve; });
    const state = await import('../../src/composables/new-document-defaults');
    const ready = state.loadNewDocumentDefaults();
    let created = false;
    const source = state.newDocumentSource().then(blob => { created = true; return blob; });
    await Promise.resolve();
    expect(created).toBe(false);
    expect(host.load).toHaveBeenCalledOnce();
    release({ run: { fontFamily: 'David', fontSizePt: 18 }, styles: {} });
    await ready;
    const xml = (await readTemplateStyles((await source)!))!;
    expect(resolveFormat(xml, 'Normal')).toMatchObject({ fontFamily: 'David', fontSizePt: 18 });
    expect(host.load).toHaveBeenCalledOnce();
  });

  it('leaves safe saved settings intact and reports refusal of the imported no-op italic preference', async () => {
    host.stored = { run: { fontSizePt: 15 }, styles: { Heading1: { color: '123456' } } };
    const state = await import('../../src/composables/new-document-defaults');
    await state.loadNewDocumentDefaults();
    const previous = state.newDocumentDefaults.value;
    expect(await state.rememberStyleDefault('Normal', { ...resolveFormat(TEMPLATE, 'Normal'), italic: true })).toBe(false);
    expect(state.newDocumentDefaults.value).toBe(previous);
    expect(host.save).not.toHaveBeenCalled();
    expect(await state.rememberStyleDefault('Heading1', { ...resolveFormat(TEMPLATE, 'Heading1'), fontSizePt: 21 })).toBe(true);
    expect(host.save).toHaveBeenCalledOnce();
    expect(state.newDocumentDefaults.value.styles.Heading1.fontSizePt).toBe(21);
  });

  it('warns about incompatible stored styles, keeps the safe font choice, and allows reset', async () => {
    host.stored = { run: { fontFamily: 'David', fontSizePt: 15 }, styles: { Normal: { italic: true }, Heading1: { color: '123456' } } };
    const state = await import('../../src/composables/new-document-defaults');
    const warning = vi.fn();
    const blank = (await state.newDocumentSource(warning))!;
    expect(warning).toHaveBeenCalledWith(state.STORED_STYLE_DEFAULTS_REJECTED_MESSAGE);
    expect(state.newDocumentDefaults.value).toEqual({ run: { fontFamily: 'David', fontSizePt: 15 }, styles: {} });
    const xml = (await readTemplateStyles(blank))!;
    expect(resolveFormat(xml, 'Normal')).toMatchObject({ fontFamily: 'David', fontSizePt: 15 });
    expect(resolveFormat(xml, 'Heading4').italic).toBe(true);
    expect(await state.describeNewDocumentDefaults()).toContain(state.STORED_STYLE_DEFAULTS_REJECTED_MESSAGE);
    expect(host.save).not.toHaveBeenCalled();
    state.resetNewDocumentDefaults();
    warning.mockClear();
    await state.newDocumentSource(warning);
    expect(warning).not.toHaveBeenCalled();
    expect(await state.describeNewDocumentDefaults()).toBeNull();
    expect(host.save).toHaveBeenCalledWith(null);
  });
});
