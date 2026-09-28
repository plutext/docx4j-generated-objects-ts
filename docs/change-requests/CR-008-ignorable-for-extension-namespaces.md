# CR-008: A marshalled part declares the Office extension namespaces it uses as ignorable

**Status:** Proposed 2026-09-28
**Depends on:** CR-001 (`fixRootNamespaceDeclarations`, which owns the root's declarations and its
`mc:Ignorable`), CR-006 (the prefix aliases it resolves through)
**Requested by:** `plutext/docx4j-core-ts` CR-002 section 29 and its check 23 (2026-09-28, at
Jason's direction), which needs `w16du:dateUtc` on every tracked revision it creates
**Counterpart:** docx4j's `Paginate.declareW14Ignorable` and `DocumentSettingsPart`, which add
`w14`/`w15` to a part's `ignorable` when they write content in those namespaces (docx4j CR-023)

## 1. Summary

A part this package marshals declares the namespaces its tree uses, and says nothing about which of
them a reader may ignore. For Office's extension namespaces that is the wrong half of the
statement: under ECMA-376 Part 3, a consumer that does not understand a namespace **not** declared
ignorable must treat the markup as an error, and Word declares every extension namespace it knows.

Reproduced here, marshalling a created document whose `w:ins` carries `dateUtc`:

```
<w:document xmlns:w="…/wordprocessingml/2006/main" xmlns:w16du="…/word/2023/wordml/word16du">
```

`w16du:dateUtc` is written and `xmlns:w16du` is declared - and there is no `mc:Ignorable` at all.
What Word writes for the same content, in `loadAndSave.docx`:

```
mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14"
```

`fixRootNamespaceDeclarations` already *keeps* and declares whatever a root's `mc:Ignorable` names
(CR-001 section 7), and resolves a prefix the table cannot produce (CR-006). It has never **added**
one, because until now nothing in this package wrote content that needed it: a loaded document
carried Word's list, and a created one wrote no extension content.

## 2. What Word actually lists, which is not what the document uses

Measured across the fidelity corpus, and it bears on the design:

| Document | `mc:Ignorable` on `word/document.xml` |
|---|---|
| `NumberingImplicitNumId.docx` | `w14 wp14` |
| `numPicBullet-word2019-pict.docx` | `w14 w15 w16se w16cid wp14` |
| `loadAndSave.docx` | `w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14` |

None of these lists is the set the document *uses*; each is the set that version of Word knows how
to write. So "what should be ignorable" cannot be derived from a tree, and matching Word exactly
would mean picking a Word version to imitate. This CR therefore proposes the narrower and
defensible rule - **list what this part actually uses** - and says plainly that it is not what Word
does.

## 3. The two shapes, and a recommendation

**A. The marshaller adds it.** When the root's type has an `ignorable` property, the tree uses a
namespace in a declared table of Office extension namespaces, and the root's `mc:Ignorable` does
not already name its prefix, the prefix is appended. Protects every writer - core-ts, the editor,
an add-in - without any of them knowing to ask.

**B. A helper the writer calls**, `declareIgnorable(root, prefix)`, mirroring docx4j's
`Paginate.declareW14Ignorable`. Explicit, and the writer decides.

**Recommended: A, with B's table exported.** docx4j can rely on B because its writers live in the
same codebase as the helper; this package's writers are other repositories, and a rule that
protects only the callers who remember it will not protect the third one. The risk of A is that the
facade makes a claim on the writer's behalf - "a reader may ignore this" - so it is confined to a
published table of namespaces for which the claim is true by construction: Office's `w14`, `w15`,
`w16*`, `wp14` and their spreadsheet and presentation equivalents, which exist precisely to be
ignorable extensions. Anything else is the writer's business, and `B` remains available for it.

Both are marshal-side only. Nothing changes on unmarshal, and a loaded document's own list is
**unioned, never replaced** - dropping a prefix Word wrote would be a regression of exactly the kind
CR-001 section 7 exists to prevent.

## 4. Where it applies

Roots whose type declares an `ignorable` property, which docx4j CR-018 and CR-023 extended: at
least `Document`, `Hdr`, `Ftr`, `CTFootnotes`, `CTEndnotes`, `Comments`, `CTCommentsEx`, `CTPeople`,
`Styles`, `Numbering`, `CTSettings`, and the SpreadsheetML roots CR-006 concerns. A root without one
is left alone: the schema does not admit the attribute there, and writing it anyway would produce a
part Office rejects - a worse outcome than the one this CR fixes.

## 5. Tests

- `test/smoke.mjs`: a created `w:document` whose `w:ins` carries `dateUtc` marshals with
  `mc:Ignorable="w16du"` and `xmlns:w16du` declared; a loaded document whose root lists ten prefixes
  keeps all ten and gains nothing it does not use; a root whose type has no `ignorable` property
  gains no attribute; and the prefix added is the one `NAMESPACE_PREFIXES` gives, so CR-006's alias
  path is exercised by a part using the SpreadsheetML main namespace.
- `test/fidelity.mjs`: the corpus must not move. Every part there is one Office wrote, so each
  already lists what it needs; a part that gains a prefix would mean the rule is adding something
  Word did not, and the test would fail. That is the check that A is conservative.

## 6. Open questions

1. **Order.** Word writes its list in a conventional order (`w14 w15 w16se … wp14`). Appending is
   simplest and the attribute is a set, so order should not matter to a reader - but "should not"
   is not "does not", and nobody here has checked what Word does with an unfamiliar order.
2. **`w16du` specifically**: whether a Word version that does not know it refuses a part that
   declares it ignorable, or ignores it as the specification says. Untested by either repository,
   and the reason core-ts writes the attribute only into documents that already list the prefix.
3. Whether the table should follow docx4j's `NamespacePrefixMappings` mechanically or be a
   hand-kept subset. Mechanical risks claiming ignorability for a namespace that is not an
   extension; hand-kept risks going stale. The CR assumes hand-kept, small, and exported.
