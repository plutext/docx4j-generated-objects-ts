# CR-008: A marshalled part declares the Office extension namespaces it uses as ignorable

**Status:** Implemented 2026-10-04 (shape A, section 9)
**Depends on:** CR-001 (`fixRootNamespaceDeclarations`, which owns the root's declarations and its
`mc:Ignorable`), CR-006 (the prefix aliases it resolves through)
**Requested by:** `plutext/docx4j-core-ts` CR-002 section 29 and its check 23 (2026-09-28, at
Jason's direction), which needs `w16du:dateUtc` on every tracked revision it creates - **though no
longer blocked on this**: see section 7
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
   **Answered 2026-10-04 by core-ts Word check 36 (section 9.7)**: Word 2010, which knows neither
   w15 nor w16du, opened both declared files; undeclared, it refused the w15 element and tolerated
   the w16du attribute.
3. Whether the table should follow docx4j's `NamespacePrefixMappings` mechanically or be a
   hand-kept subset. Mechanical risks claiming ignorability for a namespace that is not an
   extension; hand-kept risks going stale. The CR assumes hand-kept, small, and exported.

## 7. Nobody is waiting for it (2026-09-28)

The day this was filed, core-ts took docx4j's pattern for itself (its `257d40c`): before its
tracker writes `w16du:dateUtc` into a part, it adds `w16du` to that root's `ignorable`, and this
package's marshaller then declares it - which it already does. So `core-ts/CR-002.revision-dates`
records this CR as related rather than depending on it, and **no consumer is blocked**.

That is the right outcome for core-ts and it does not retire the CR, but it does change what the CR
is for, and the distinction is worth keeping straight. It is no longer "the thing that makes
tracked dates work"; it is a safeguard for **the writer who does not know to ask** - the editor, an
add-in, any consumer writing `w14` or `w15` content into a part it created - and a second line
under core-ts, whose per-part call is correct but easy to omit in a new code path.

A safeguard with no one waiting is a weaker case for scheduling than a blocker, and should be
judged as one. What would strengthen it is evidence for open question 2: if some Word version does
refuse a part whose extension content is not declared ignorable, this stops being a tidiness
measure. Neither repository has tested it, and neither can without Word.

(2026-10-04: it has happened. Word 2010 refuses an undeclared w15 element - section 9.7.)

## 8. Someone is waiting now: content a caller brings in (core-ts, 2026-10-04)

Asked by Jason from the core-ts session on 2026-10-04, after core-ts CR-002 section 40 (the
editor's requests from Word check 35). Section 7's "second line under core-ts, easy to omit in a new
code path" has happened, and in a path where core-ts's own pattern cannot reach.

**The case.** core-ts now adds `w15` to a root's `ignorable` wherever *it* writes a `w15` element
(its `f44f9a0`: the content control setters and `insertContentControl`), as it does `w16du` for
revision dates. But `insertOoxml` and `insertXml` splice in content the **caller** wrote, in any
extension namespace. Reproduced on core-ts `fb58e29` with this package at 0.3.0: a fragment holding
`<w15:appearance w15:val="tags"/>` in a `w:sdtPr`, inserted into a created document, marshals as

```
<w:document xmlns:w="…/wordprocessingml/2006/main" xmlns:w15="…/word/2012/wordml">
```

with the element written and no `mc:Ignorable`. core-ts could scan what it inserts for a table of
namespaces, but that table is this CR's, and the editor's own save (`exportedSdtPr`, its ED-005
12.58) already carries a third copy of the same rule for `w15`. Shape A does it once, for every
namespace in the table and every writer.

**Asked:** implement shape A as section 3 recommends, with the table exported (B's helper is
optional; core-ts has its own `declareIgnorable(part, prefix)` over the root's `ignorable`).
Section 4's roots and section 5's tests stand. Two additions to the tests from this case: a
`w:document` whose `w:sdtPr` holds a `w15` element gains `mc:Ignorable="w15"`; and one whose root
already lists `w15` (core-ts having declared it) is unchanged, so the two mechanisms do not fight.

**What core-ts assumes meanwhile, and does afterwards.** Meanwhile `insertOoxml` and `insertXml`
declare nothing, and core-ts's per-writer calls stay. Once this lands and core-ts takes the release,
those calls become redundant; core-ts will decide then whether to keep them as documentation of
intent or remove them, and its tests that assert `mc:Ignorable="w15"` and `"w16du"` must pass
unchanged either way.

**What is still not measured (open question 2).** Word check 35 opened, in Word 2010, a file whose
`w15` content **was** declared ignorable, and it opened and dropped the `w15` elements on saving.
Nobody has opened one whose `w15` content was **not** declared, in Word 2010 or any other version.
So the case for this CR is still ECMA-376 Part 3's rule and Word's own practice, not an observed
refusal. One file settles it: a document with `w15:appearance` on a control and no `mc:Ignorable`,
opened in Word 2010. core-ts can make that file on request.

(Made and opened the same day, with three companions: Word 2010 refused it - section 9.7.)


## 9. As implemented (2026-10-04)

Shape A, as section 8 asked. `marshalToDocument` - so `marshalString`, `marshalNode` and each typed
part of `marshalPackage` - calls `declareExtensionsIgnorable` on the marshalled root before
`fixRootNamespaceDeclarations`, which then declares what was added. B's helper is not added: core-ts
has its own, and nothing else has asked.

### 9.1 The table: what Office lists, measured

`IGNORABLE_EXTENSION_NAMESPACES` is exported beside `NAMESPACE_PREFIXES`: twenty namespace URIs, no
prefixes, since the prefix listed is the one the call's table gives. Open question 3 is answered
"hand-kept", and kept by a criterion rather than by judgement: **a namespace is listed when Office
itself names it in a part root's `mc:Ignorable`**, which the fidelity corpus measures. Word's eleven
(`w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl cr w16du wp14`, in the order Word writes them on
`commentsExtensible.xml`) and Excel's nine (`x14ac x15 x16r2 xr xr2 xr3 xr6 xr10 xr16`; Excel's own
order varies by part). The `x` token Excel lists on slicer parts is SpreadsheetML main, not an
extension, and stays CR-006's alias.

The measurement is what kept section 3's "and their spreadsheet and presentation equivalents" from
being taken literally. `x15ac` is an extension namespace by any reading, and Excel uses it on every
workbook (`x15ac:absPath`), but only inside an `mc:Choice Requires="x15"`, and never lists it;
in the table, it would have been added to every workbook the corpus holds. Likewise `a14`, `a16`,
`c14` and the PowerPoint namespaces, which Office writes only inside a Choice requiring them or
inside an extension list. No PresentationML root binds `ignorable` in any case (section 4's list
holds; the full set is the twelve WordprocessingML roots, `CTCommentsEx`, `CTPeople`,
`CTCommentsIds`, `CTCommentsExtensible`, the ten SpreadsheetML roots, and the x14 slicer and x15
timeline roots), so a presentation entry could never apply.

### 9.2 "Uses": three places Office does not count

Office does not list every extension namespace a part uses, and the gap is not one of version: it
is where the use is. A use counts unless it is

- **the root's own namespace.** Excel's x15 `timelineCacheDefinition` and `timelines` parts list
  `xr10` and `x xr10`, not x15. Without this exemption both fail the fidelity test (tried). Word's w15, w16cid
  and w16cex parts do list their own namespace; the union keeps that, and a created one gains it
  only from content in another namespace.
- **below an `mc:Choice` whose `Requires` names it.** Only a reader that understands the namespace
  takes that branch, which is why Word never lists `wps`. ECMA-376 Part 3's own rule, and the only
  one of the three no corpus part needs - no table namespace appears there - so the smoke tests it:
  w14 inside `Requires="w14"` adds nothing, inside `Requires="wps"` adds `w14`.
- **below an extension list's `ext`** (an `ext` whose parent is an `extLst`, in any namespace). A
  reader skips one by its `uri`, not by `mc:Ignorable`: Excel's `styles.xml` carries
  `x15:timelineStyles` in an `ext` and lists `x14ac x16r2 xr`. Without this, sixteen SpreadsheetML
  parts fail (tried). The `ext` element itself still counts.

Section 2's rule stands: this lists what the part uses, less those three, and does not imitate a
Word version's full list.

### 9.3 Roots, prefixes and order

The root's type is the one its `TYPE_NAME` names, or for a literal without one, the one its element
name declares; it qualifies when it or a base declares an attribute property named `mc:Ignorable`.
The markup-compatibility elements bind one too, and are excluded, since none is ever a part root.

Prefixes the root already lists are kept in their order, and a namespace already listed under any
prefix is not listed again; new prefixes are appended in the table's order. Open question 1 is not
settled by this - nobody has checked what Word does with an order it would not write - but a created
part now lists in Word's order, and a loaded one keeps Word's list exactly unless the tree uses
something it does not name. A namespace the call's table writes as the default has no prefix to
list and is left unlisted; none in the table is, with the default table.

### 9.4 Section 5's SpreadsheetML test, as read

"The prefix added is the one `NAMESPACE_PREFIXES` gives, so CR-006's alias path is exercised by a
part using the SpreadsheetML main namespace" is taken as two checks, since the main namespace is
not an extension and is never added: a worksheet (main namespace as the default) using
`x14ac:dyDescent` gains `x14ac`, declared, with the default untouched; and CR-006's slicer cache
keeps `x xr10` exactly, the alias still resolving `x` and its x14 root namespace not added. The
prefix is shown to be the call's table's by a third: a per-call table mapping w16du to `du` lists
`du`.

### 9.5 Tests

`test/smoke.mjs`, one block: section 5's four (with 9.4's reading), section 8's two (a literal
`w:document` holding a content control built by `wml` with `w15:appearance` gains `w15`; one
already listing `w14 w15` is unchanged), and one line for each exemption of 9.2. With the call to
`declareExtensionsIgnorable` removed from the build, the block's first assertion fails.

`test/fidelity.mjs` is unchanged and is section 5's conservativeness check: 268 checks, every part
unchanged, the resolved forms included. It is load-bearing - removing the root exemption fails 2
parts, removing the `ext` exemption 16.

### 9.6 Consumers

core-ts: its per-writer `declareIgnorable` calls (w16du on revisions, w15 on content controls)
become redundant once it takes the release, and stay harmless - a namespace already listed is not
listed again, the second of section 8's tests. Its tests asserting `mc:Ignorable="w15"` and
`"w16du"` on created documents should pass unchanged either way, which is the thing to confirm
before release. A part that core-ts writes and whose tree uses another table namespace (w14 on a
copied paragraph, x14ac on a row) now gains that prefix too: correct by this CR, and visible in
core-ts's marshal hashes where its fixtures lack a list Office would have written.

**Verified by core-ts against this working tree (2026-10-04)**, core-ts at `fb58e29` plus check 36's
files: its whole suite (744 tests, typecheck, zero parity differences over 62 docx4j goldens); every
XML and rels part of its 183 fixture packages after `unmarshalAll()` and a save - 2849 parts -
compared as text under 0.3.0 and this checkout, of which exactly 2 differ, both check 36's
undeclared files, and only in the root start tag (`xmlns:mc` and `mc:Ignorable="w15"` / `"w16du"`
added). So on core-ts's fixtures the rule adds nothing to any part Office or docx4j wrote, the
same result as the fidelity corpus here. The consumer case: `insertOoxml` into a created document
now lists w15 for `w15:appearance`, `w15 w16sdtdh` for a `w15:dataBinding` carrying
`w16sdtdh:storeItemChecksum`, w14 for `w14:paraId` and `w14:checkbox`, and w16du for `dateUtc`;
core-ts's own writers beside it give `w15 w16du` once each. Not verified: a browser, and the
published artifact, which core-ts re-checks the same three ways when it exists.

### 9.7 Open question 2, answered for Word 2010: core-ts Word check 36

Measured the same day (core-ts's `test/README.md` check 36, run by Jason; files in its
`test/fixtures/revisions/check36/`). Open question 2 is answered for Word 2010:

- an extension **element** not declared ignorable is refused. A document whose `w:sdtPr` holds
  `<w15:appearance w15:val="tags"/>` with `xmlns:w15` and no `mc:Ignorable` did not open: "problems
  with the contents", an unspecified error at the element's position, no offer to repair. The same
  document with `mc:Ignorable="w15"` opened;
- an extension **attribute** not declared ignorable is tolerated. A `w:ins` carrying
  `w16du:dateUtc` with no `mc:Ignorable` opened and showed its tracked insertion, as its declared
  twin did;
- Word 2010 drops both on saving, declared or not. Word 365, which knows both namespaces, opened the
  undeclared files without a message.

One element and one attribute were tried, in Word 2010 only among versions that know neither
namespace. So this CR is a **correctness fix** for extension elements, no longer only a safeguard:
section 7's "what would strengthen it" has happened. For attributes the declaration remains what
the specification asks and Word writes, without an observed failure behind it.

**Attributes still count, deliberately.** Check 36 shows one attribute tolerated by one version,
which is not a licence to stop: ECMA-376 Part 3 draws no line between an element and an attribute
in a non-understood namespace, Office itself lists a namespace it uses only in attributes (every
worksheet in the corpus lists `x14ac`, used there only as `x14ac:dyDescent`), and the cost of
listing one is a token.

### 9.8 Known limits

- **Content held as DOM is looked into.** The walk is over the marshalled DOM, not the typed
  objects, so a node the model kept from a wildcard - inside an unresolved `mc:Choice`, or any
  other `any` slot - counts like typed content, its attributes included. The smoke's
  `Requires="wps"` case is one: `w14:x` is unknown to the model and held as a DOM node. Tried
  as well with a w15 element three levels down a DOM subtree and a `w14:paraId` on a DOM element:
  each root gained the prefix.
- **A part held whole as DOM gains nothing.** `marshalPackage` passes a part whose root the model
  does not know through untouched; it has no type, so nothing says its root admits `mc:Ignorable`.
- **Below an extension list's `ext`, nothing is looked at** (section 9.2), in typed or DOM content.
- **Word 2010 is the only version tried** for a refusal, with one element and one attribute.

