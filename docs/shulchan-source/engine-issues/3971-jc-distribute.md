### Summary

`format.paragraph.setAlignment` only accepts `PARAGRAPH_ALIGNMENTS = ['left', 'center', 'right', 'justify']` (`document-api/src/paragraphs/paragraphs.types.d.ts`), so there is no way to write `w:jc="distribute"`.

The model already *reads* it: `types/paragraph.types.d.ts` has `alignment?: ... | 'distributed'` and `types/sd-props.d.ts` lists `'distributed' | 'thaiDistribute' | ...`. A document that comes in with `distribute` round-trips, but an integrator cannot set it.

### Use case

Hebrew religious-text typesetting conventionally justifies every line **including the last one** ("Distributed" in Word's Paragraph dialog). The de-facto Word tooling used by editors of these texts sets `wdAlignParagraphDistribute` on body paragraphs. We are porting that tooling to SuperDoc and this is the one alignment we cannot produce.

### Expected

- Write side: `format.paragraph.setAlignment({ target, alignment: 'distribute' })` (or `'distributed'`, matching the read-side name), emitting `<w:jc w:val="distribute"/>`.
- Layout: the last line is stretched like the other lines when `jc=distribute` (i.e. `ResolvedTextLineItem.skipJustify` would be `false` for the last line in that mode).

`thaiDistribute` and the kashida variants are out of scope here; `distribute` alone covers the Hebrew case.

### Versions

`superdoc@2.10.0` and `2.11.0` — `PARAGRAPH_ALIGNMENTS` unchanged.
