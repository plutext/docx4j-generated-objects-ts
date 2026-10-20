# CR-011: Declare from the source the prefixes only the source bound

**Status:** Proposed 2026-10-20
**Requested by:** core-ts, at Jason's word (2026-10-20: "Send this declaration defect to the
objects session as a CR there"). The editor found it while building ED-009's fixture. core-ts
reproduced it in core-ts 0.5.0 (main f8617e8) with this package at 0.4.2 and `@docx4j/jsonix`
3.4.2, and records it as its CR-010 section 14.9. On the first draft, core-ts measured that both of
its `insertOoxml` forms bring in Choices whose prefix only the inserted OOXML declares. It suggested
keeping each Choice's record on the Choice itself, which this revision adopts (section 3.1).
**Depends on:** CR-001 sections 3.1 and 3.2, whose `declare()` this extends. For recording at
unmarshal, it also depends on a new `@docx4j/jsonix` CR: an unmarshal-side `onElement`, the
counterpart of jsonix-CR-006's marshal one (section 3.2). Not yet filed.
**Counterpart:** docx4j records the prefixes at unmarshal. `Docx4jUnmarshallerListener` passes
each `mc:Choice`'s `Requires`, `Ignorable` and `MustUnderstand` to
`JaxbXmlPart.addMcChoiceNamespace`, and has since 6.0.0. It records the prefixes only, per part, not
their namespaces. On marshal, `NamespacePrefixMapperUtils.declareNamespaces` resolves each prefix
through `NamespacePrefixMappings.getNamespaceURIStatic`, its static table, and otherwise logs "No
mapping for prefix" and declares nothing. Reading the code, docx4j has the same gap; it has not been
run here. Since docx4j CR-046, `McSelection.recordRequiresNamespaces` records each `Requires` prefix
by its namespace, but only to select a branch; the marshal doesn't use it.

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

Four shapes, all from the requests:
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
- **Moved in from another source:** core-ts's `Body.insertOoxml` with a bare `<w:p>` fragment, or
  with a flat OPC package. The inserted content is unmarshalled on its own, then moved into the
  destination's tree. Whatever the fragment declared stays behind with the fragment.

Office declares only the prefixes it knows, so none of Office's own files has these shapes
(section 2). Any part with a private `Requires` or `Ignorable` prefix does, if it is saved with its
branches kept. Keeping branches is core-ts's default load since its 0.5.0. With `mcePreprocess` the
Fallback is taken, so no Choice is written, but the `mc:Ignorable` entry is still dropped.

**Proposed** (section 3):
- **record** each source binding that the table and the aliases would not give. A Choice's
  bindings are kept on the Choice itself, so they go wherever an edit moves it. The root's bindings
  are kept on the root, for its `mc:Ignorable`;
- **find** the records at marshal through the marshaller's `onElement`;
- **declare** from them before the table.

Nothing changes for Office's files: in a prototype, the corpus hashes are identical, and the cost
is within measurement noise.

## 2. Measured (2026-10-20)

**The defect, at the facade alone.** I ran `unmarshalString` then `marshalString` on 0.4.2's
`dist/`, without core-ts:
- the editor's shape: the root keeps neither `xmlns:zz9` nor `mc:Ignorable`, the Choice's
  `Requires="zz9"` resolves nowhere, and both warnings appear;
- the probe's `word/document.xml`: 11 Choices, of which 4 resolve nowhere once marshalled. Three
  name zz9 (one as `Requires="wps zz9"`) and one names shape2010. The warning names `zz9, shape2010`.

**The move, in core-ts** (core-ts measured this, on f8617e8 with objects 0.4.2): `Body.insertOoxml(...,
'End')` into `loadAndSave.docx`, then a save. Both insertion forms leave the inserted Choice's
`Requires="zz9"` unbound:
- a bare `<w:p>` whose root declares `xmlns:zz9`. This goes through `builders/wml`'s `wml()`, and so
  through `unmarshalNode`;
- the same paragraph in a flat OPC package. That package is loaded as its own
  `WordprocessingMLPackage`, and its body content is moved, with `PARENT` relinked.

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
  section 3.4 puts the source first.

**A prototype** of section 3, built in a scratch copy and not committed (164 changed lines in
`src/index.mts`). It stands in for the runtime callback with a wrapper on the Choice class's
`unmarshal` (section 3.2).
- Twelve cases, each run on the prototype and on 0.4.2. On the prototype all twelve behave as
  section 3 says. On 0.4.2, ten of them leave a prefix unbound or drop it from `mc:Ignorable`:
  - the editor's shape, and the probe's;
  - a paragraph moved in from another unmarshalled source, and a moved `wml()` fragment;
  - a flat OPC part;
  - a paragraph marshalled on its own;
  - `deepCopy` of a paragraph, and of the root;
  - the superseded table declaration (section 3.4);
  - a prefix bound two ways (section 3.5).
- The two that need no fix are a Choice nothing binds, which warns in both builds, and the
  tree-used conflict, whose namespace is the tree's in both. The prototype also gives that conflict
  its new warning.
- The smoke passes once the one case section 5 changes is changed.
- Fidelity passes: 359 checks, and 631 prefixes resolve.
- The bundle test passes and the typecheck is clean.
- `marshal-hashes.mjs` gives the same hashes as 0.4.2 for all 341 parts (333 marshalled).

**Cost:** measured on a synthetic 2.1 MB `document.xml` (140,002 elements, 500 Choices requiring
`w14`), using the median of 7 runs.
- **unmarshal, marshal and `deepCopy` are the same as 0.4.2,** within run-to-run noise: unmarshal
  1,016 to 1,029 ms against 1,006 to 1,034 ms, and the same for marshal and `deepCopy`. Office's
  bindings are all ones the table gives, so nothing is recorded.
- **When records exist, marshal and `deepCopy` cost more.** An early version recorded every
  binding, and with those records marshal took about 5 to 8% longer and `deepCopy` of the whole
  document about 60% longer: 280 to 300 ms against 173 to 189 ms. The extra time is the walk that
  carries records across a copy (section 3.3), and it doesn't depend on how many records there are.
  The flag that turns the walk on is process-wide, so this applies to every copy once any file with
  a source-only binding has been loaded.

## 3. Proposal

### 3.1 What is recorded, and where

A binding is recorded only if the table and the aliases would not give it. That makes the record
empty for every file Office wrote. Records are kept in a module-private `WeakMap` that is keyed by
typed objects. The tree gains no property, its JSON is the same, and the records are collected with
the objects.

- **Each `mc:Choice`:** the binding in scope at the source Choice of each prefix its `Requires`
  names, found with `lookupNamespaceURI`. This catches the probe's declarations on
  `mc:AlternateContent`. It is keyed by the Choice's own typed object, so it goes wherever an edit
  moves the Choice, with nothing for the caller to do: core-ts's `insertOoxml` and `insertXml`, or
  a cut and paste.
- **Each part root unmarshalled through the facade** (`unmarshalString`, `unmarshalNode`, and
  `unmarshalPackage` for each part): the bindings in scope at the source root, meaning its own
  declarations and its ancestors'. A flat OPC part inherits `pkg:package`'s. The nearer one wins,
  and the default namespace is left out. This record serves the root's `mc:Ignorable`, and is a
  fallback for a `Requires` whose Choice has no record of its own.

To record from text, `unmarshalString` now parses the text itself, with `Jsonix.DOM.parse`, and then
calls `unmarshalDocument`. That is the runtime's own `unmarshalString`, so parse errors and the
located errors of `@docx4j/jsonix` 3.4.0 are unchanged.

### 3.2 Recording each Choice as it is unmarshalled

The typed Choice and the DOM element it came from are paired only inside the runtime. Its
unmarshal reports nothing per element: jsonix-CR-006 gave it `onUnexpectedElement` and
`onUnexpectedAttribute`, and gave only the marshaller an `onElement`. Two ways to get the pair:

1. **Recommended: an unmarshal-side `onElement(node, value)` in `@docx4j/jsonix`.** This is the
   counterpart of the marshal callback, a new jsonix CR. The facade would set it on the
   unmarshallers it creates, composed with the caller's own. CR-011 would then wait for that jsonix
   release, and this package's floor would rise to it.
2. **What the prototype does:** when the context is built, wrap the `unmarshal` of the
   `org_docx4j_mce.AlternateContent.Choice` class info, and read `input.node` before it runs. This
   needs no runtime release, so it could ship in a patch now. But it depends on internals: the
   method's signature and `Jsonix.XML.Input`'s `node`. Today the facade only *reads* one internal,
   `TypeInfo.properties`, and has noted it as a candidate for a jsonix typings CR. It patches none.
   The smoke would catch a runtime change that broke the wrapper.

The choice between them is Jason's: (1) is the cleaner design, (2) is the faster release.

### 3.3 Finding the records at marshal, and carrying them across a copy

- **At marshal**, the facade passes the marshaller an `onElement`, composed with any the caller
  passes (core-ts's `binder.mts` passes one). It is set only once some record exists. For each
  written element whose object is a recorded Choice, it notes the DOM Choice and that Choice's
  bindings. Wildcard DOM content is reported against its original node, which is still in its
  source document, so the in-scope bindings of the Choices inside it are looked up there. This also
  makes a paragraph marshalled on its own declare what its Choices need: core-ts does that in
  `Paragraph`, `ContentControl`, `office-js/fragment.mts` and `toApiScript`.
- **`deepCopy`**, and so `deepCopyAs`, gives each copied object the record of its original. The copy
  has the original's shape, so this is one walk over both. The walk is skipped while no record exists.

### 3.4 Declaring from the record first

In `fixRootNamespaceDeclarations`, each prefix that `mc:Ignorable` or a `Requires` names is
resolved as follows:
1. A declaration that an element or attribute in the tree uses stays as it is. It can't be rebound
   without changing that content's namespace.
2. Otherwise the source's binding is used, then the table's, then the aliases'. For a `Requires`,
   the source's binding is its Choice's record, else the root's. For `mc:Ignorable`, it is the
   root's record.

A declaration the runtime made from the table, and that is kept only because `mc:Ignorable` or a
`Requires` names its prefix, gives way to the source's binding. Without that rule, a source binding
`w14` to something else would get the table's `w14`.

The source comes first because its binding is what the prefix meant. The `x` in the VML parts of
section 2 is the measured case: there, the aliases would give the wrong namespace.

### 3.5 Warnings

Each is given once per prefix per marshal.
- **Nothing resolves a prefix:** the existing two warnings, with the same wording. This now means
  no source bound the prefix either.
- **New, the tree's namespace wins:** the source bound a named prefix to one namespace, but the tree
  uses the prefix for another. Root declarations can't keep both meanings. Declaring it on the Choice
  would also rebind the Choice's content. So the warning says the `Requires` or `mc:Ignorable` now
  means the tree's namespace.
- **New, one prefix bound two ways:** two of a part's Choices, or a Choice and the root, mean
  different namespaces by one prefix that the tree doesn't use. A root declaration can carry only
  one of them, so the first is declared, with this warning. Nothing measured binds a prefix two ways.

### 3.6 Release

This fix changes output only for parts with these shapes, as 0.4.2 did. With section 3.2's way 2, it
is a patch (0.4.3). With way 1 it waits for the runtime release, and raises the floor as 0.4.1 did.
Both choices are Jason's.

## 4. Not proposed

- **Declarations kept in the typed tree,** as a property or in `otherAttributes`. That would change
  the objects' shape and JSON. Jsonix's `deepCopy` copies own enumerable properties, so it would
  carry them, but that is the only gain. Declaring on demand in the runtime (jsonix-CR-003 part 2,
  deferred there) is the route that would retire `fixRootNamespaceDeclarations` altogether.
- **One record per part, found through `PARENT`.** This was the first draft. It fails for content
  moved in from another source, which section 2 measured in both of core-ts's `insertOoxml` forms.
  It also needs an export such as `adoptSourceBindings(from, to)` for a caller to call; keeping the
  record on the Choice needs nothing.
- **One registry for the whole process** of every prefix seen. It would leak one document's
  bindings into another's marshal.
- **`MustUnderstand`, a Choice's own `mc:Ignorable`, and `ProcessContent`.** The facade doesn't
  declare their prefixes today, and no part in the corpus carries any of them.
- **Merging an inserted root's `mc:Ignorable` into the destination's.** core-ts's `insertOoxml`
  doesn't do it. That is a separate question, and core-ts hasn't raised it.

## 5. Tests

- **Smoke, in CR-001 section 3.2's block:**
  - **The case that pins the defect today:** `kept('zz', 'xmlns:zz="urn:zz"')` asserts the
    `Requires` is kept and warned about. It changes to "declared from the source, no warning". The
    "nothing resolves" case becomes a Choice whose source declares nothing.
- **Smoke, new cases:** the twelve of section 2, with these checks:
  - **nothing recorded for the table's own bindings:** a part whose bindings are all the table's
    leaves no record, so marshal installs no `onElement` and `deepCopy` does no walk;
  - **a caller's `onElement` is still called** for every element.
- **Fidelity:** unchanged. The corpus never triggers the rule, which section 2 measured, and its
  prefix-resolution check (631) holds. The probe isn't a part Office wrote, so it belongs in the
  smoke's inline cases, not in the corpus.
- **Release step 6:** expect the tree's hashes to match the 0.4.2 tarball's, apart from parts that
  have these shapes, of which the corpus has none.
- **core-ts:** its section 14.9 test asserts the broken behaviour, so it fails on this release.
  That failure is the signal that the fix is in. Its two `insertOoxml` measurements should then
  declare zz9.

## 6. Risks

- **Output changes, as intended,** for a part whose source bound a private prefix that
  `mc:Ignorable` or a `Requires` names. Before, `mc:Ignorable` lost the entry and the `Requires`
  was left unbound. Now both are declared. A reader that doesn't understand `zz9` ignores it, as the
  source told it to.
- **Faithful to an odd source.** Suppose a source binds a known prefix to another namespace and
  names it. That meaning is now kept, where before the table's was written. A reader of the source
  saw the source's meaning, so it is the right one to keep.
- **Once a record exists in the process,** every marshal runs the `onElement` and every `deepCopy`
  walks. That costs about 5 to 8% on marshal and 60% on a whole-document `deepCopy` (section 2).
  core-ts says the `deepCopy` cost doesn't reach it (2026-10-20). Its `deepCopy` calls copy at
  most a paragraph: run and paragraph properties, a style, a cell's properties, a content control.
  Its `pkg.clone()` saves the package and loads it again, and never copies a tree. The marshal cost
  falls on save.
- **Way 2 of section 3.2** depends on runtime internals, as that section says.
- **Records go with typed objects.** A Choice rebuilt from scratch, for instance from a
  `toSource` (CR-003 phase B) or by hand, has no record, and warns as it does today.

## 7. Consumers

- **core-ts:**
  - **No code change is needed.** `XmlPart` unmarshals through `unmarshalNode` and marshals through
    `marshalNode`. Both `insertOoxml` forms unmarshal through the same context, and so record their
    Choices. Its `binder.mts` `onElement` keeps working, composed with the facade's.
  - **Its own `recordDeclarations` and `declareIgnorablePrefixes`** (`src/parts/XmlPart.mts`, its
    CR-001 section 17) can never act. `fixRootNamespaceDeclarations` already declares or drops every
    `mc:Ignorable` prefix inside `marshalNode`, and has since 0.1.5 (CR-001 section 3.1). core-ts
    confirmed this on the first draft: its zz9 repro logged this package's "dropped" warning, never
    its own. It removes both when it raises its range to this release, and records that in CR-010
    section 14.9.
  - **Its section 14.9 test flips.**
- **The editor:** ED-009's fixture keeps its declarations when saved.
- **Office JS add-ins using this package alone:** an `unmarshalString`/`marshalString` round trip
  keeps a part's private prefixes, and so does content moved between parts.
- **`@docx4j/jsonix`,** under way 1 of section 3.2: a new CR for the unmarshal-side `onElement`.
- **docx4j:** the same gap, from reading the code (see the header). A docx4j CR could pass the
  URIs that CR-046's `McSelection` records to the marshal side. Whether to raise one is Jason's call.
