### Summary

`doc.mutations.apply({ atomic: true, steps: [{ op: 'text.rewrite', where: { by: 'target', target }, ... }] })` is rejected at runtime, although `TargetWhere` is part of the public `StepWhere` union in `document-api/src/types/mutation-plan.types.d.ts`:

```ts
export type TargetWhere = { by: 'target'; target: SelectionTarget };
export type StepWhere = SelectWhere | RefWhere | TargetWhere | BlockWhere;
```

The call fails with:

```
DocumentApiValidationError: v2 text.rewrite currently requires a ref produced by query.match/find or a single text selector.
```

So the types advertise a by-offset target, but the v2 implementation only accepts `by: 'ref'` and `by: 'select'`.

### Versions

- Measured against the packaged `dist` of `superdoc@2.10.0`.
- `superdoc@2.11.0`: the `TargetWhere` declaration in `mutation-plan.types.d.ts` is unchanged.

### Why this matters

A "fix all" style tool (hundreds of small, offset-addressed text edits across a document, each already resolved to a `SelectionTarget`) has exactly two options today:

- `where: { by: 'select', select: { type: 'text', pattern }, within: block, require: 'all' }` — works and gives a single undo step, **but** it addresses by pattern, so it cannot express "the same pattern gets two different replacements in the same paragraph" (e.g. `" ."` becomes `"."` at one place and `". "` at another, depending on what follows).
- `where: { by: 'block' }` — rewrites the whole block text. Works, but on a paragraph with mixed run formatting the rewritten text is re-styled with `onNonUniform: 'majority'`, silently flattening bold/italic runs. That is a worse regression than losing atomic undo.

Result: we fall back to N separate `doc.replace({ target, text })` calls, which means N undo steps for the user (300 fixes = 300 Ctrl+Z).

### Expected

Either `by: 'target'` is accepted by `text.rewrite` (it is the most natural input for callers that already computed exact ranges), or the member is removed from `StepWhere` so the gap shows up at compile time instead of at runtime. The first would be far more useful.

### Minimal reproduction

```js
const doc = superdoc.activeEditor.doc;
const { blocks } = await doc.blocks.list({ includeText: true });
const block = blocks[0];
// a text range inside `block`, as returned by doc.find / query.match or computed from block.text
const target = {
  kind: 'selection',
  start: { kind: 'text', blockId: block.nodeId, offset: 0 },
  end:   { kind: 'text', blockId: block.nodeId, offset: 3 },
};
await doc.mutations.apply({
  atomic: true,
  changeMode: 'direct',
  steps: [{ op: 'text.rewrite', where: { by: 'target', target }, text: 'xyz' }],
});
// -> DocumentApiValidationError: v2 text.rewrite currently requires a ref produced by query.match/find or a single text selector.
```

The same step with `where: { by: 'select', select: { type: 'text', pattern: '...' }, within: { kind: 'block', nodeType: 'paragraph', nodeId: block.nodeId }, require: 'all' }` succeeds.
