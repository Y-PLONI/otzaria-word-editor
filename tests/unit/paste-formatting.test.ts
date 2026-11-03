import { describe, expect, it } from 'vitest';
import {
  destinationMarks,
  mergeFragment,
  scriptSegments,
  textOnlyFragment,
  type PasteFragment,
} from '../../src/engine/paste-formatting';

/**
 * ה-fragment שה-`parse` של המנוע בנה מה-HTML שנמדד (Arial 20 אדום, מודגש,
 * נטוי, קו תחתון, סימון, מרכוז, כותרת) — בצורה שהחוזה מגדיר.
 */
function sourceFragment(): PasteFragment {
  const look = { fontFamily: 'Arial', fontSizePt: 20, color: '#FF0000' };
  return {
    kind: 'superdoc.clipboard.fragment',
    version: 'v2.1',
    blocks: [
      {
        kind: 'paragraph',
        paragraphProps: { alignment: 'center' },
        pPrXml: '<w:pPr><w:jc w:val="center"/></w:pPr>',
        styleRef: 'Quote',
        runs: [
          { text: 'מודגש', marks: { ...look, bold: true } },
          { text: ' ', marks: look },
          { text: 'נטוי', marks: { ...look, italic: true } },
          { text: ' רקע', marks: { ...look, highlight: 'yellow' } },
          { text: 'm', marks: look },
          { text: '2', marks: { ...look, verticalAlign: 'superscript' } },
        ],
      },
      { kind: 'heading', level: 1, runs: [{ text: 'כותרת', marks: { bold: true } }] },
      {
        kind: 'paragraph',
        listIntent: { kind: 'bullet', level: 0 },
        numbering: { numId: '3', ilvl: '0' },
        runs: [{ text: 'פריט', marks: { underline: true, color: '#00FF00' } }],
      },
    ],
  };
}

const blocks = (fragment: PasteFragment) => fragment.blocks as unknown as Array<Record<string, unknown>>;
const runs = (block: Record<string, unknown>) => block.runs as Array<{ text: string; marks?: object }>;

describe('mergeFragment — „מזג עיצוב”', () => {
  it('משאיר מהמקור רק הדגשות: מודגש, נטוי, קו תחתון וכתב עילי', () => {
    const merged = mergeFragment(sourceFragment(), {});
    const first = runs(blocks(merged)[0]!);

    expect(first.map((run) => run.marks)).toEqual([
      { bold: true },
      undefined,
      { italic: true },
      undefined,
      undefined,
      { verticalAlign: 'superscript' },
    ]);
    expect(runs(blocks(merged)[2]!)[0]!.marks).toEqual({ underline: true });
  });

  it('מוותר על עיצוב הפסקה של המקור — היא יורשת את סגנון היעד', () => {
    const [first, heading] = blocks(mergeFragment(sourceFragment(), {}));

    for (const key of ['paragraphProps', 'pPrXml', 'styleRef']) expect(first).not.toHaveProperty(key);
    // „כותרת 1” של המקור היא סגנון, ולא תוכן.
    expect(heading!.kind).toBe('paragraph');
    expect(heading).not.toHaveProperty('level');
  });

  it('רשימה נשארת רשימה', () => {
    const list = blocks(mergeFragment(sourceFragment(), {}))[2]!;
    expect(list.listIntent).toEqual({ kind: 'bullet', level: 0 });
    expect(list.numbering).toEqual({ numId: '3', ilvl: '0' });
  });

  it('עיצוב היעד על כל ריצה, וההדגשה של המקור מוסיפה עליו', () => {
    const merged = mergeFragment(sourceFragment(), { fontFamily: 'David', fontSizePt: 14, fontSizeCsPt: 14 });
    const first = runs(blocks(merged)[0]!);

    expect(first[0]!.marks).toEqual({ fontFamily: 'David', fontSizePt: 14, fontSizeCsPt: 14, bold: true });
    expect(first[1]!.marks).toEqual({ fontFamily: 'David', fontSizePt: 14, fontSizeCsPt: 14 });
  });

  it('הדגשה שחלה על כל הטקסט אינה הדגשה אלא מראה — ונזרקת (כמו ב-Word)', () => {
    const fragment: PasteFragment = {
      blocks: [
        {
          kind: 'paragraph',
          runs: [
            { text: 'כולו', marks: { bold: true, italic: true } },
            // רווח אינו מכריע: הוא אינו הופך את המשפט ל„מודגש בחלקו”.
            { text: ' ' },
            { text: 'מודגש', marks: { bold: true } },
          ],
        },
      ],
    };

    const merged = runs(blocks(mergeFragment(fragment, {}))[0]!);
    expect(merged[0]!.marks).toEqual({ italic: true });
    expect(merged[2]!.marks).toBeUndefined();
  });

  it('קישור נשאר קישור, וההדגשה שבתוכו נשמרת', () => {
    const fragment: PasteFragment = {
      blocks: [
        {
          kind: 'paragraph',
          runs: [{ text: 'ראו ' }, { text: 'כאן', marks: { underline: true, color: '#0563C1' } }],
          inlines: [
            { kind: 'text', text: 'ראו ', marks: { fontFamily: 'Arial' } },
            {
              kind: 'hyperlink',
              target: { kind: 'external', url: 'https://example.org' },
              inlines: [{ kind: 'text', text: 'כאן', marks: { underline: true, color: '#0563C1' } }],
            },
          ],
        },
      ],
    };

    const paragraph = blocks(mergeFragment(fragment, {}))[0]!;
    const inlines = paragraph.inlines as Array<Record<string, unknown>>;
    expect(inlines[0]).toEqual({ kind: 'text', text: 'ראו ' });
    expect(inlines[1]!.target).toEqual({ kind: 'external', url: 'https://example.org' });
    expect(inlines[1]!.inlines).toEqual([{ kind: 'text', text: 'כאן', marks: { underline: true } }]);
  });

  it('טבלה נשארת טבלה, בלי צבע המילוי והריפוד של המקור', () => {
    const fragment: PasteFragment = {
      blocks: [
        {
          kind: 'table',
          rows: [
            {
              cells: [
                {
                  gridSpan: 2,
                  shadingColor: 'FFFF00',
                  padding: { topPt: 4 },
                  blocks: [{ kind: 'paragraph', styleRef: 'X', runs: [{ text: 'תא', marks: { color: '#FF0000' } }] }],
                },
              ],
            },
          ],
        },
      ],
    };

    const table = blocks(mergeFragment(fragment, { fontFamily: 'David' }))[0]!;
    const cell = (table.rows as Array<{ cells: Array<Record<string, unknown>> }>)[0]!.cells[0]!;
    expect(cell.gridSpan).toBe(2);
    expect(cell).not.toHaveProperty('shadingColor');
    expect(cell).not.toHaveProperty('padding');
    expect(cell.blocks).toEqual([{ kind: 'paragraph', runs: [{ text: 'תא', marks: { fontFamily: 'David' } }] }]);
  });

  it('אינו משנה את ה-fragment שקיבל', () => {
    const fragment = sourceFragment();
    const before = JSON.stringify(fragment);
    mergeFragment(fragment, { fontFamily: 'David' });
    expect(JSON.stringify(fragment)).toBe(before);
  });
});

describe('textOnlyFragment — „שמור טקסט בלבד”', () => {
  it('שום עיצוב מהמקור — גם לא הדגשות ולא רשימות; רק עיצוב היעד', () => {
    const only = textOnlyFragment(sourceFragment(), { color: '#0000FF' });

    for (const block of blocks(only)) {
      expect(block.kind).toBe('paragraph');
      expect(block).not.toHaveProperty('listIntent');
      expect(block).not.toHaveProperty('numbering');
      for (const run of runs(block)) expect(run.marks).toEqual({ color: '#0000FF' });
    }
  });
});

describe('HTML emphasis script semantics', () => {
  const source: PasteFragment = { blocks: [{ kind: 'paragraph', runs: [
    { text: 'plain' }, { text: 'bold', marks: { bold: true } }, { text: 'italic', marks: { italic: true } },
  ] }] };
  it('HTML local emphasis overrides destination off for both script channels', () => {
    const output = mergeFragment(source, { bold: false, boldCs: false, italic: false, italicCs: false }, { mirrorHtmlEmphasis: true });
    expect(runs(blocks(output)[0]!).map((run) => run.marks)).toEqual([
      { bold: false, boldCs: false, italic: false, italicCs: false },
      { bold: true, boldCs: true, italic: false, italicCs: false },
      { bold: false, boldCs: false, italic: true, italicCs: true },
    ]);
  });
  it('private fragments retain independent Latin-only emphasis', () => {
    const output = mergeFragment(source, { boldCs: false, italicCs: false });
    expect(runs(blocks(output)[0]!)[1]!.marks).toEqual({ boldCs: false, italicCs: false, bold: true });
  });
  it('uniform HTML bold is still discarded in both channels', () => {
    const output = mergeFragment({ blocks: [{ kind: 'paragraph', runs: [{ text: 'all', marks: { bold: true } }] }] }, {}, { mirrorHtmlEmphasis: true });
    expect(runs(blocks(output)[0]!)[0]!.marks).toBeUndefined();
  });
});

describe('destinationMarks — explicit destination states', () => {
  it('retains black and every explicit false rather than falling back to inherited defaults', () => {
    expect(destinationMarks({ color: '#000000', bold: false, italic: false, underline: false, strikethrough: false }))
      .toEqual({ color: '#000000', bold: false, italic: false, underline: false, strike: false });
  });
  it('does not flatten font/size channels from the toolbar', () => {
    expect(destinationMarks({ fontFamily: 'Arial', fontSizePt: 12 })).toEqual({});
  });
  it('keeps actual underline/highlight and leaves unknown states unknown', () => {
    expect(destinationMarks({ underline: true, highlight: '#ffff00' })).toEqual({ underline: true, highlight: 'yellow' });
  });
  it('rewrites nested content-control text and counts it for partial emphasis', () => {
    const source: PasteFragment = { blocks: [{ kind: 'paragraph', runs: [], inlines: [
      { kind: 'sdt', tag: 'keep-me', inlines: [{ kind: 'sdt', inlines: [{ kind: 'text', text: 'bold', marks: { bold: true, fontFamily: 'Source' } }] }] },
      { kind: 'text', text: 'plain', marks: { fontFamily: 'Source' } },
    ] }] };
    const output = mergeFragment(source, { color: '#000000', bold: false });
    expect(JSON.stringify(output)).not.toContain('Source');
    expect(JSON.stringify(output)).toContain('keep-me');
    expect(JSON.stringify(output)).toContain('"bold":true');
    expect(JSON.stringify(textOnlyFragment(source, { bold: false }))).not.toContain('"bold":true');
  });
});

describe('scriptSegments', () => {
  it('splits Hebrew from Latin and attaches neutrals to the preceding script', () => {
    expect(scriptSegments('שלום, world 12 עוד')).toEqual([
      { text: 'שלום, ', complex: true },
      { text: 'world 12 ', complex: false },
      { text: 'עוד', complex: true },
    ]);
  });
  it('gives leading neutrals to the first letter after them, and letterless text to Hebrew', () => {
    expect(scriptSegments('(abc) א')).toEqual([{ text: '(abc) ', complex: false }, { text: 'א', complex: true }]);
    expect(scriptSegments('12.')).toEqual([{ text: '12.', complex: true }]);
  });
  it('splits a merged run per script so each script keeps its own destination font', () => {
    const fragment = { blocks: [{ kind: 'paragraph', runs: [{ text: 'דף lamed', marks: { bold: true } }] }] };
    const merged = mergeFragment(fragment, { fontFamily: 'Times New Roman', fontFamilyCs: 'David', fontSizePt: 12, fontSizeCsPt: 14 });
    expect((merged.blocks[0] as { runs: unknown[] }).runs).toEqual([
      { text: 'דף ', marks: { fontFamily: 'David', fontSizePt: 14, fontSizeCsPt: 14 } },
      { text: 'lamed', marks: { fontFamily: 'Times New Roman', fontSizePt: 12, fontSizeCsPt: 12 } },
    ]);
  });
});
