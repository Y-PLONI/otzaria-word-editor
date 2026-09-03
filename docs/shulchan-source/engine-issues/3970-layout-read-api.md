### Summary

Feature request: a **read-only, public layout query API** exposing what the layout engine already computes — per-paragraph line boundaries (as text offsets) and per-page block boundaries — plus a way to await "pagination settled" after a mutation.

The engine has this data internally. `dist/layout-engine/contracts/src/resolved-layout.d.ts` declares `ResolvedPage`, `ResolvedFragmentItem` and `ResolvedTextLineItem { line, lineIndex, availableWidth, skipJustify, ... }`, and the paginator lives in `dist/layout-engine/layout-engine/src/paginator.d.ts`. None of it is reachable from the `superdoc` public surface: `SuperDoc` exposes `pagination-update` (`{ totalPages }`) and `viewport-change` / `getViewportMetrics()` (available width, fit zoom) only, and `doc.*` has no layout namespace.

### Use cases

We are building a Hebrew typesetting toolset on top of SuperDoc (porting a widely used set of Word macros for editors of rabbinic texts). Several of the most requested tools are pure functions of line/page geometry and cannot be built without it:

1. **Hanging first word ("window")** — the second line starts under the second word of the first line. Needs: where line 1 breaks (text offset) and the rendered width of the first word.
2. **Last-line balance** — avoid a single orphan word on the last line by widening spaces on the tightest earlier line. Needs: line boundaries and per-line slack.
3. **Page / column alignment** — nudge paragraph spacing so a page (or both columns of a two-column section) ends flush. Needs: which page a block ends on, bottom of content per page/column.
4. **Page-break marking / verification** — mark or verify the first/last word of every page. Needs: page → (blockId, offset) boundaries.
5. **Skip rules** — "apply only to paragraphs with 2+ lines". Needs: line count per paragraph.

Today the only way to approximate these is to read the rendered DOM (`[data-page-index]`, `Range.getClientRects()` on page elements) and re-associate rects with `doc.blocks.list()` heuristically, then wait on ad-hoc settle delays after every mutation. That is fragile (zoom via CSS transform, virtualized pages, paragraphs split across pages, RTL bidi runs) and couples integrators to internal DOM structure.

### Proposed shape (illustrative)

```ts
// read-only; lengths in pt (or twips), text positions as offsets in the block's text
doc.layout.lines({ blockId }): Promise<Array<{
  index: number;
  start: number; end: number;               // text offsets within the block
  pageIndex: number; columnIndex?: number;
  widthPt: number; availableWidthPt: number; // slack = available - width
  top: number; height: number;              // relative to the page content box
}>>;

doc.layout.pages(): Promise<Array<{
  index: number;
  contentBottomPt: number;                  // bottom of the last laid-out line
  blocks: Array<{ blockId: string; fromLine: number; toLine: number; startsHere: boolean; endsHere: boolean }>;
}>>;

// resolves after the pagination that reflects the latest mutation
doc.layout.settled(): Promise<{ totalPages: number }>;
```

Even a subset (lines per block + page boundaries + a settled promise) would unblock all five tools above. Read-only is enough; we do not need to drive layout.

### Versions

`superdoc@2.10.0` and `2.11.0` — the public `.d.ts` in both has no layout query surface.
