/** Read the destination's authored run formatting without flattening script fonts.
 * The clipboard contract can represent a single family, but cannot represent
 * separate Latin/CS families or a direct theme font. Reject those cases before
 * insertion. Uniform fixed families are captured from the paragraph cascade.
 * Theme/split channels remain available through the original keep-source mode.
 */
import type { SelectionTarget } from 'superdoc/ui';
import type { PasteRunMarks } from './paste-formatting';
import { restoreComplexScriptStyles } from './docx-cs-display';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export interface DestinationCapture {
  evaluatedRevision?: string;
  fragment: { xml: string; namespaces?: Record<string, string> };
  source?: { paragraphStyleId?: string | null };
  context?: { xml?: readonly string[] };
  dependencies?: readonly { target: string; dataBase64?: string }[];
}

const elements = (element: Element, name: string) => Array.from(element.children).filter(
  (child) => child.namespaceURI === W && child.localName === name,
);
const attr = (element: Element | undefined | false, name: string) => element ? element.getAttributeNS(W, name) : null;

function parseXml(xml: string): Element {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('Invalid destination XML');
  return doc.documentElement;
}

/** Only the font properties need the cascade: a partial direct font override
 * cannot be copied into the scalar clipboard family without checking its peers.
 */
function inheritedFonts(capture: DestinationCapture): Record<string, string> {
  const dependency = capture.dependencies?.find((item) => item.target.endsWith('/styles.xml'));
  if (!dependency?.dataBase64) throw new Error('Destination styles unavailable');
  const binary = atob(dependency.dataBase64);
  const xml = new TextDecoder().decode(Uint8Array.from(binary, (char) => char.charCodeAt(0)));
  // The live sheet shows each level's Hebrew font/size in its Latin slot
  // (docx-cs-display.ts); read the document's own values, as Word would.
  const styles = parseXml(restoreComplexScriptStyles(xml) ?? xml);
  const fonts: Record<string, string> = { b: '0', bCs: '0', i: '0', iCs: '0' };
  function add(rPr: Element | undefined, styleToggle = false): void {
    const node = rPr && elements(rPr, 'rFonts')[0];
    if (rPr) for (const key of ['sz', 'szCs', 'b', 'bCs', 'i', 'iCs', 'highlight']) {
      const child = elements(rPr, key)[0];
      if (child) {
        const value = attr(child, 'val') ?? '1';
        const toggle = ['b', 'bCs', 'i', 'iCs'].includes(key);
        if (styleToggle && toggle) {
          if (!['0', 'false', 'off'].includes(value)) fonts[key] = fonts[key] === '1' ? '0' : '1';
        } else fonts[key] = toggle ? (['0', 'false', 'off'].includes(value) ? '0' : '1') : value;
      }
    }
    if (node) for (const attribute of Array.from(node.attributes)) {
      if (attribute.namespaceURI === W) fonts[attribute.localName] = attribute.value;
    }
  }
  const defaults = elements(styles, 'docDefaults')[0];
  const runDefault = defaults && elements(defaults, 'rPrDefault')[0];
  add(runDefault && elements(runDefault, 'rPr')[0]);
  const all = elements(styles, 'style');
  const defaultStyle = all.find((style) => attr(style, 'type') === 'paragraph' && ['1', 'true', 'on'].includes(attr(style, 'default') ?? ''));
  const defaultCharacter = all.find((style) => attr(style, 'type') === 'character' && ['1', 'true', 'on'].includes(attr(style, 'default') ?? ''));
  if (defaultCharacter && (elements(defaultCharacter, 'rPr')[0]?.children.length || elements(defaultCharacter, 'basedOn').length)) throw new Error('Default character formatting unsupported');
  const seen = new Set<string>();
  function addStyle(id: string | null | undefined): void {
    if (!id) return;
    if (seen.has(id)) throw new Error('Cyclic destination style');
    seen.add(id);
    const style = all.find((node) => attr(node, 'styleId') === id);
    if (!style) throw new Error('Missing destination style');
    const paragraphProps = elements(style, 'pPr')[0];
    if (paragraphProps?.getElementsByTagNameNS(W, 'numPr').length) throw new Error('Numbered destination style unsupported');
    addStyle(attr(elements(style, 'basedOn')[0], 'val'));
    add(elements(style, 'rPr')[0], true);
  }
  addStyle(capture.source?.paragraphStyleId ?? attr(defaultStyle, 'styleId'));
  return fonts;
}

export function marksFromDestinationCapture(capture: DestinationCapture): PasteRunMarks {
  const namespaces = { w: W, ...capture.fragment.namespaces };
  const bindings = Object.entries(namespaces).map(([key, value]) => `xmlns:${key}="${value.replace(/&/g, '&amp;').replace(/"/g, '&quot;')}"`).join(' ');
  const root = parseXml(`<capture ${bindings}>${capture.fragment.xml}</capture>`);
  for (const xml of capture.context?.xml ?? []) {
    const context = parseXml(`<capture ${bindings}>${xml}</capture>`);
    // The paragraph mark's own run properties (`pPr/rPr`, `<w:rtl/>` in every
    // Hebrew Word paragraph) format the mark only, never text inserted mid-run.
    for (const mark of Array.from(context.getElementsByTagNameNS(W, 'pPr')).flatMap((pPr) => elements(pPr, 'rPr'))) mark.remove();
    if (['tblPr', 'tblGrid', 'tcPr', 'numPr', 'rPr'].some((name) => context.getElementsByTagNameNS(W, name).length)) {
      throw new Error('Structural destination formatting unsupported');
    }
  }
  const run = root.getElementsByTagNameNS(W, 'r')[0];
  if (!run) throw new Error('Destination run unavailable');
  const props = elements(run, 'rPr')[0];
  // `rtl` declares direction, `kern`/`noProof`/`ligatures` are
  // typography and proofing: none is appearance that merge must carry over.
  const supported = new Set(['rFonts', 'b', 'bCs', 'i', 'iCs', 'sz', 'szCs', 'color', 'highlight', 'u', 'strike', 'caps', 'smallCaps', 'vertAlign', 'lang', 'rtl', 'kern', 'noProof', 'ligatures']);
  if (props && Array.from(props.children).some((node) => !supported.has(node.localName))) throw new Error('Authored destination formatting unsupported');
  const directHighlight = props && elements(props, 'highlight')[0];
  if (attr(directHighlight, 'val') === 'none') throw new Error('Explicit highlight cancellation unsupported');
  const underline = props && elements(props, 'u')[0];
  if (underline && !['single', 'none'].includes(attr(underline, 'val') ?? 'single')) throw new Error('Unsupported destination underline');
  if (props && elements(props, 'rStyle').length) throw new Error('Destination character style unsupported');
  const marks: PasteRunMarks = {};
  const inherited = inheritedFonts(capture);
  const directFonts = props && elements(props, 'rFonts')[0];
  if (directFonts) {
    const direct = Object.fromEntries(Array.from(directFonts.attributes).filter((item) => item.namespaceURI === W).map((item) => [item.localName, item.value]));
    // `hint` and the East Asian channel never select a Hebrew or Latin font.
    if (Object.keys(direct).some((key) => /theme$/i.test(key) && key !== 'eastAsiaTheme')) throw new Error('Direct theme font unsupported');
    for (const channel of ['ascii', 'hAnsi', 'cs']) {
      if (direct[channel]) {
        inherited[channel] = direct[channel]!;
        delete inherited[channel === 'cs' ? 'cstheme' : `${channel}Theme`];
        delete inherited[`${channel}Theme`];
      }
    }
  }
  // Latin and Hebrew may differ (Times New Roman + David is the common Word
  // pairing): the rewrite gives each script segment its own family. The two
  // Latin channels, and theme fonts, still cannot be represented.
  const [ascii, hAnsi, cs] = ['ascii', 'hAnsi', 'cs'].map((channel) => inherited[channel]);
  if (!ascii || !hAnsi || !cs || ascii.toLowerCase() !== hAnsi.toLowerCase() ||
      ['asciiTheme', 'hAnsiTheme', 'csTheme', 'cstheme'].some((key) => inherited[key])) {
    throw new Error('Destination font channels cannot be captured faithfully');
  }
  marks.fontFamily = ascii;
  if (cs.toLowerCase() !== ascii.toLowerCase()) marks.fontFamilyCs = cs;
  const highlight = attr(directHighlight, 'val') ?? inherited.highlight;
  if (highlight && highlight !== 'none') marks.highlight = highlight;
  for (const [xmlName, key] of [['sz', 'fontSizePt'], ['szCs', 'fontSizeCsPt']] as const) {
    const value = attr(props && elements(props, xmlName)[0], 'val') ?? inherited[xmlName] ?? null;
    if (value !== null) {
      const size = Number(value) / 2;
      if (!Number.isFinite(size) || size <= 0) throw new Error('Invalid destination size');
      marks[key] = size;
    }
  }
  for (const [xmlName, key] of [['b', 'bold'], ['bCs', 'boldCs'], ['i', 'italic'], ['iCs', 'italicCs'], ['caps', 'caps'], ['smallCaps', 'smallCaps']] as const) {
    const node = props && elements(props, xmlName)[0];
    const value = node ? attr(node, 'val') ?? '1' : inherited[xmlName];
    if (value !== undefined) marks[key] = !['0', 'false', 'off'].includes(value);
  }
  const vertical = attr(props && elements(props, 'vertAlign')[0], 'val');
  if (vertical === 'baseline' || vertical === 'superscript' || vertical === 'subscript') marks.verticalAlign = vertical;
  return marks;
}

/** A one-character read at the insertion point; it never changes live selection. */
export function destinationSample(target: SelectionTarget): SelectionTarget | null {
  if (target.start.kind !== 'text') return null;
  const offset = target.start.offset;
  return { ...target, start: { ...target.start, offset: Math.max(0, offset - 1) }, end: { ...target.start, offset: offset > 0 ? offset : 1 } };
}

/** Empty documents have no legal one-character capture. Inspect their small
 * exported package instead; this fallback is never used for a populated document.
 */
export async function marksFromEmptyDocument(contentBase64: string, blockId: string): Promise<PasteRunMarks> {
  const { readZip, readEntryText } = await import('./docx-parts');
  const binary = atob(contentBase64);
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  const entries = readZip(bytes);
  if (!entries) throw new Error('Empty document package unavailable');
  const document = entries.find((entry) => entry.name === 'word/document.xml');
  const styles = entries.find((entry) => entry.name === 'word/styles.xml');
  if (!document || !styles) throw new Error('Empty document parts unavailable');
  const documentXml = await readEntryText(document);
  const stylesXml = await readEntryText(styles);
  if (!documentXml || !stylesXml) throw new Error('Empty document XML unavailable');
  const root = parseXml(documentXml);
  const body = elements(root, 'body')[0];
  if (!body || elements(body, 'p').length !== 1 || Array.from(body.children).some((child) => !['p', 'sectPr'].includes(child.localName))) throw new Error('Not a standalone empty document');
  const paragraph = elements(body, 'p')[0]!;
  if (paragraph.getAttributeNS('http://schemas.microsoft.com/office/word/2010/wordml', 'paraId') !== blockId) throw new Error('Empty destination block changed');
  if (Array.from(paragraph.children).some((child) => child.localName !== 'pPr')) throw new Error('Empty document contains run context');
  const pPr = elements(paragraph, 'pPr')[0];
  if (pPr && ['numPr', 'rPr'].some((name) => elements(pPr, name).length)) throw new Error('Empty paragraph formatting unsupported');
  const styleId = pPr && attr(elements(pPr, 'pStyle')[0], 'val');
  const styleBytes = new TextEncoder().encode(stylesXml);
  let styleBinary = '';
  for (const byte of styleBytes) styleBinary += String.fromCharCode(byte);
  return marksFromDestinationCapture({
    fragment: { xml: '<w:r><w:t/></w:r>' },
    source: { paragraphStyleId: styleId || null },
    dependencies: [{ target: '/word/styles.xml', dataBase64: btoa(styleBinary) }],
  });
}
