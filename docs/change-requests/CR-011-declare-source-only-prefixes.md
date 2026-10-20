# CR-011: Declare from the source the prefixes only the source bound

**Status:** Proposed 2026-10-20
**Requested by:** core-ts, at Jason's word (2026-10-20: "Send this declaration defect to the
objects session as a CR there"). The editor found it while building ED-009's fixture, and core-ts
reproduced it in core-ts 0.5.0 (main f8617e8) with this package at 0.4.2 and `@docx4j/jsonix`
3.4.2. core-ts records it as its CR-010 section 14.9.
**Depends on:** CR-001 sections 3.1 and 3.2, whose `declare()` this extends. Nothing outside this
repository.
**Counterpart:** docx4j records the prefixes at unmarshal. `Docx4jUnmarshallerListener` passes
each `mc:Choice`'s `Requires`, `Ignorable` and `MustUnderstand` to
`JaxbXmlPart.addMcChoiceNamespace`, and has since 6.0.0. It records prefixes only, though, not
namespaces. On marshal, `NamespacePrefixMapperUtils.declareNamespaces` resolves each prefix through
`NamespacePrefixMappings.getNamespaceURIStatic`, its static table, and otherwise logs "No mapping for
prefix" and declares nothing. Reading the code, docx4j has the same gap; it has not been run here.
Since docx4j CR-046, `McSelection.recordRequiresNamespaces` records each `Requires` prefix by its
namespace, but only to select a branch; the marshal doesn't use it.

## 1. Summary

When the facade unmarshals a part, the typed tree keeps no namespace declarations. Those
declarations are rebuilt on marshal (CR-001 section 3):
- the runtime declares every `NAMESPACE_PREFIXES` entry on the root;
- `fixRootNamespaceDeclarations` removes the ones nothing uses;
- for each prefix that `mc:Ignorable` or an `mc:Choice`'s `Requires` names and that is still
  undeclared, `declare()` looks in the table, then in `IGNORABLE_PREFIX_ALIASES` (CR-006).

A prefix that neither has can't be declared. Since 0.4.2 it is kept in a `Requires` with a warning,
and dropped from `mc:Ignorable` with a warning. The source declared it, but that declaration is
gone by then. The part then names a prefix it never declares, which ECMA-376 Part 3 forbids. Whether
Word refuses such a file hasn't been checked.

Three shapes, all from the requests:
- **Declared on the root:** `xmlns:zz9="urn:example:zz9" mc:Ignorable="zz9"`, with a
  `<mc:Choice Requires="zz9">` in the body (the editor's fixture). The Ignorable entry is dropped,
  and the Requires is left unbound.
- **Declared below the root:** docx4j's `extraction-probe.docx` (docx4j 4a89a5967) declares
  `xmlns:zz9="http://example.com/zz9"` on each of three `mc:AlternateContent` elements. CR-001
  section 3.2 says "a declaration below the root that binds the prefix answers for its Choice", but
  that holds only while the declaration is in the tree, and the typed tree keeps none.
- **A second prefix for a namespace the table knows:** the same probe's `shape2010` is bound to the
  `wps` namespace. The table maps namespaces to prefixes, so it can't say what `shape2010` means; only
  the source can.

Office declares only the prefixes it knows, so none of Office's own files has these shapes
(section 2). Any part with a private `Requires` or `Ignorable` prefix does, if it is saved with its
branches kept. Keeping branches is core-ts's default load since its 0.5.0. With `mcePreprocess` the
Fallback is taken, so no Choice is written, but the `mc:Ignorable` entry is still dropped.

**Proposed** (section 3):
- record the source's bindings when a part is unmarshalled;
- find that record when any element of the part is marshalled;
- declare a prefix from the record before trying the table.

Nothing changes for Office's files: in a prototype, the corpus hashes are identical.

## 2. Measured (2026-10-20)

**The defect, at the facade alone.** I ran `unmarshalString` then `marshalString` on 0.4.2's
`dist/`, without core-ts:
- the editor's shape: the root keeps neither `xmlns:zz9` nor `mc:Ignorable`, the Choice's
  `Requires="zz9"` resolves nowhere, and both warnings appear;
- the probe's `word/document.xml`: 11 Choices, of which 4 resolve nowhere once marshalled. Three
  name zz9 (one as `Requires="wps zz9"`) and one names shape2010. The warning names `zz9, shape2010`.

**The fidelity corpus** (341 parts that Office wrote) never needs the source:
- It has 1,673 root declarations, 520 `mc:Ignorable` prefixes and 29 `Requires` prefixes. Of the 29
  Choices, 13 declare their prefix themselves (`a14`, `c14`, `cx1`, `sle15`, `tsle`).
- Every prefix an `mc:Ignorable` or `Requires` names is in the table or the aliases.
- Five root prefixes are in neither: `dcmitype` (17 parts), `aink` (22), `clbl` (3), `c16r2` (3)
  and `p188` (2). None of them is named by an `mc:Ignorable` or a `Requires`.
- One root binding disagrees with what the table and aliases would declare. Three VML parts bind
  `x` to `urn:schemas-microsoft-com:office:excel`, but the aliases map `x` to SpreadsheetML main (CR-006),
  and the table writes the Excel VML namespace as `xvml`. Nothing names `x` in those parts, but if
  something did, the table and the aliases would declare the wrong namespace for it. That is why
  section 3.3 puts the source first.

**A prototype** of section 3, built in a scratch copy and not committed (86 changed lines in
`src/index.mts`):
- both repro shapes come back with every prefix bound and no warning;
- the smoke passes once the one case section 5 changes is changed;
- fidelity passes: 359 checks, and 631 prefixes resolve;
- the bundle test passes and the typecheck is clean;
- `marshal-hashes.mjs` gives the same hashes as 0.4.2 for all 341 parts (333 marshalled).

**Cost:** I parsed a synthetic 2.1 MB `document.xml` (140,002 elements, 500 Choices) with xmldom. The
record's scan (section 3.1) took 18 ms (median of 9 runs), while `unmarshalString` of the same part
takes about 1.2 s. That is about 1.5%, paid at unmarshal only.

## 3. Proposal

### 3.1 Record the source's bindings at unmarshal

The facade's unmarshal paths each record the source's bindings: `unmarshalString`, `unmarshalNode`,
and `unmarshalPackage` for each part. For each, the record holds:
- **the bindings in scope at the source root:** its own declarations and its ancestors', since a
  flat OPC part also inherits `pkg:package`'s. The nearer one wins, and the default namespace is
  left out;
- **for each `mc:Choice`, the binding in scope there of each prefix its `Requires` names**, found
  with `lookupNamespaceURI`. This is what catches the probe's declarations on `mc:AlternateContent`.

The record is kept in a module-private `WeakMap` keyed by the root's value object, which an edit
keeps. The typed tree is not changed: it gains no property, its JSON is the same, and the record is
collected with the tree.

To record from text, `unmarshalString` now parses the text itself, with `Jsonix.DOM.parse`, and
then calls `unmarshalDocument`. That is the runtime's own `unmarshalString`, so parse errors and
the located errors of `@docx4j/jsonix` 3.4.0 are unchanged.

### 3.2 Find the record when marshalling

A marshal looks for a record on the element's value, and if there is none, on its nearest ancestor
by `PARENT`. So a paragraph or content control marshalled on its own finds its part's record. core-ts
does this in `Paragraph`, `ContentControl`, `office-js/fragment.mts` and `toApiScript`. `deepCopy`
of a recorded root gives the copy the same record.

### 3.3 Declare from the record first

In `fixRootNamespaceDeclarations`, each prefix that `mc:Ignorable` or a `Requires` names is
resolved as follows:
1. A declaration that an element or attribute in the tree uses stays as it is. It can't be rebound
   without changing that content's namespace.
2. Otherwise the source's binding is used, then the table's, then the aliases'.

A declaration the runtime made from the table, and that is kept only because `mc:Ignorable` or a
`Requires` names its prefix, gives way to the source's binding. Without that rule, a source binding
`w14` to something else would get the table's `w14`.

The source comes first because its binding is what the prefix meant. The `x` in the VML parts of
section 2 is the measured case: there, the aliases would give the wrong namespace.

### 3.4 Warnings

- **Nothing resolves a prefix:** the existing two warnings, with the same wording. This now means
  the source didn't bind the prefix either, or the root was built by hand and has no record.
- **New:** the source bound a prefix that `mc:Ignorable` or a `Requires` names to one namespace, but
  the tree uses the prefix for another. Root declarations can't keep both meanings. Declaring it on
  the Choice would also rebind the Choice's content. So the warning says the `Requires` or
  `mc:Ignorable` now means the tree's namespace.

### 3.5 One binding per prefix per part

The record holds one binding for each prefix, as docx4j's per-part record does. Suppose a part
binds one prefix to two namespaces, on the root and on a Choice, or on two Choices. Then the record
keeps the root's binding, or else the first in document order. The marshal warns where a
`Requires` names such a prefix. Nothing measured binds a prefix two ways. The prototype doesn't
have this warning yet; it keeps the first binding silently.

### 3.6 Release

This is a fix that changes output only for parts with these shapes, like 0.4.2, so a patch fits
(0.4.3). The version is Jason's call.

## 4. Not proposed

- **A record per Choice, by object.** It would be exact where section 3.5 is not. But it needs to
  know which typed object each DOM Choice became, and `@docx4j/jsonix` 3.4 has no unmarshal
  callback that says so. That would take a jsonix CR, and nothing measured needs it.
- **Declarations kept in the typed tree,** as a property or in `otherAttributes`. That would change
  the objects' shape, `deepCopy` and JSON. Declaring on demand in the runtime (jsonix-CR-003 part 2,
  deferred there) is the route that would retire `fixRootNamespaceDeclarations` altogether.
- **One registry for the whole process** of every prefix seen. It would leak one document's
  bindings into another's marshal.
- **`MustUnderstand`, a Choice's own `mc:Ignorable`, and `ProcessContent`.** The facade doesn't
  declare their prefixes today, and no part in the corpus carries any of them.
- **Content moved in from another source.** Suppose a fragment is unmarshalled on its own and its
  children are moved into a part, as an `insertXml` or `insertOoxml` can do. Its bindings stay with
  the fragment's root. If core-ts finds a case, an export such as `adoptSourceBindings(from, to)`
  would cover it. Word's `getOoxml()` writes the table's prefixes.
- **A root built by hand.** There is no source to record, so it warns as in 0.4.2.

## 5. Tests

- **Smoke, in CR-001 section 3.2's block:**
  - **The case that pins the defect today:** `kept('zz', 'xmlns:zz="urn:zz"')` asserts the
    `Requires` is kept and warned about. It changes to "declared from the source, no warning". The
    "nothing resolves" case becomes a Choice whose source declares nothing.
- **Smoke, new cases:**
  - the editor's shape (root-declared, with `mc:Ignorable` and `Requires`);
  - the probe's shape (declared on `mc:AlternateContent`, including the `shape2010` second prefix);
  - a flat OPC part through `unmarshalPackage` and `marshalPackage`;
  - a paragraph marshalled alone;
  - `deepCopy` of the root;
  - a hand-built root, which warns as before;
  - the superseded table declaration;
  - the tree-used conflict, which warns;
  - a prefix bound two ways, which warns.
- **Fidelity:** unchanged. The corpus never triggers the rule, which section 2 measured, and its
  prefix-resolution check (631) holds. The probe isn't a part Office wrote, so it belongs in the
  smoke's inline cases, not in the corpus.
- **Release step 6:** expect the tree's hashes to match the 0.4.2 tarball's, apart from parts that
  have these shapes, of which the corpus has none.
- **core-ts:** its section 14.9 test asserts the broken behaviour, so it fails on this release.
  That failure is the signal that the fix is in.

## 6. Risks

- **Output changes, as intended,** for a part whose source bound a private prefix that
  `mc:Ignorable` or a `Requires` names. Before, `mc:Ignorable` lost the entry and the `Requires`
  was left unbound. Now both are declared. A reader that doesn't understand `zz9` ignores it, as the
  source told it to.
- **Faithful to an odd source.** Suppose a source binds a known prefix to another namespace and
  names it in `mc:Ignorable`. That meaning is now kept, where before the table's was written. A
  reader of the source saw the source's meaning, so it is the right one to keep.
- **`PARENT`:** the walk in section 3.2 relies on parent pointers, which the facade's context sets
  by default. A context built with `parentPointers: false` finds only the root's own record.
- **Memory:** one small map per unmarshalled part root, typically 20 to 40 entries, collected with
  the tree.

## 7. Consumers

- **core-ts:**
  - **No code change is needed.** `XmlPart` unmarshals through `unmarshalNode` and marshals
    through `marshalNode`, so it gets the record without changes.
  - **Its own `recordDeclarations` and `declareIgnorablePrefixes`** (`src/parts/XmlPart.mts`, its
    CR-001 section 17) can never act, as I read them. `fixRootNamespaceDeclarations` already
    declares or drops every `mc:Ignorable` prefix inside `marshalNode`, and has since 0.1.5
    (CR-001 section 3.1). So by the time `declareIgnorablePrefixes` runs, nothing is left
    undeclared. Its comment, that the facade "can only *remove* unused declarations", has been out
    of date since then. After this CR, core-ts can delete both.
  - **Its section 14.9 test flips,** and it raises its range once this is released.
- **The editor:** ED-009's fixture keeps its declarations when saved.
- **Office JS add-ins using this package alone:** an `unmarshalString`/`marshalString` round trip
  keeps a part's private prefixes.
- **docx4j:** the same gap, from reading the code (section header). A docx4j CR could pass the
  URIs that CR-046's `McSelection` records to the marshal side. Whether to raise one is Jason's call.
