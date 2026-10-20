# CR-011: Declare from the source the prefixes only the source bound

**Status:** Implemented 2026-10-20, released in 0.4.3 (2026-10-20), on `@docx4j/jsonix` 3.5.0's jsonix-CR-008; the runtime route decided 2026-10-20 (Jason: "objects CR-011: the jsonix CR")
**Requested by:** core-ts, at Jason's word (2026-10-20: "Send this declaration defect to the
objects session as a CR there"). The editor found it while building ED-009's fixture. core-ts
reproduced it in core-ts 0.5.0 (main f8617e8) with this package at 0.4.2 and `@docx4j/jsonix`
3.4.2, and records it as its CR-010 section 14.9. On the first draft, core-ts measured that both of
its `insertOoxml` forms bring in Choices whose prefix only the inserted OOXML declares. It suggested
keeping each Choice's record on the Choice itself, which this revision adopts (section 3.1).
**Depends on:** CR-001 sections 3.1 and 3.2, whose `declare()` this extends. For recording at
unmarshal, it also depends on **jsonix-CR-008**, an unmarshal-side `onElement` that is the
counterpart of jsonix-CR-006's marshal one (section 3.2). Jason chose this route on 2026-10-20, and
core-ts relayed the decision. The jsonix session filed and implemented it the same day, and released
it in `@docx4j/jsonix` 3.5.0 (tag 6807bbe), which is this package's floor (`^3.5.0`) from this CR.
**Counterpart:** docx4j CR-052 (proposed 2026-10-20, drafted by docx4j-25 at this CR's request).
docx4j-25 measured docx4j on VERSION_17_3_4 at cf1abf001 before drafting it:
- **A part saved as loaded already keeps its bindings.** `extraction-probe.docx` saves with all 12
  `Requires` prefixes resolving, and docx4j writes `xmlns:zz9` and `xmlns:shape2010` on each
  `mc:AlternateContent`. A root `mc:Ignorable="zz9"` loaded from its package is declared on the root.
  This package lacks exactly that, and this CR adds it.
- **Moved, copied or built content does not:**
  - a Choice built in code with `Requires="cx1"` is undeclared on save, though the table knows `cx1`;
  - the probe's zz9 paragraphs, moved or deep-copied into another package, are undeclared;
  - a root `mc:Ignorable="zz9"` set from code is written with no declaration.
  This package already declares a table prefix such as `cx1` for a built Choice (CR-001 section 3.2).
  This CR's per-Choice record and its `deepCopy` carry cover the moved and copied cases.

Which docx4j record a loaded part's bindings come from is CR-052's phase 0. There are two
candidates: CR-046's per-Choice `McSelection.REQUIRES_NAMESPACES`, and the part's
`addMcChoiceNamespace` string. docx4j-25 rates CR-052 medium priority, high if Word refuses such a
file. An earlier draft of this header said, from reading the code alone, that docx4j had this CR's
whole gap. The measurement shows that is true only for content that has moved.

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

**Implemented** (section 3):
- **record** each source binding that the table and the aliases would not give. A Choice's
  bindings are kept on the Choice itself, so they go wherever an edit moves it. The root's bindings
  are kept on the root, for its `mc:Ignorable`;
- **find** the records at marshal through the marshaller's `onElement`;
- **declare** from them before the table.

Nothing changes for Office's files: the corpus hashes are identical to 0.4.2's, and the cost is
within measurement noise.

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

**The implementation** of section 3 is in `src/index.mts`. Before it, a prototype was built in a
scratch copy: first on `@docx4j/jsonix` 3.4.2, with a wrapper on the Choice class's `unmarshal`
standing in for the callback, then on jsonix-CR-008's real callback (jsonix 080fbb0, linked). The
results below were measured on the prototype and hold on the implementation, which runs on the
published 3.5.0.
- Unchanged objects 0.4.2 on jsonix 080fbb0: the typecheck and `npm test` pass, and all 341
  corpus hashes are identical, so the runtime change alone moves nothing.
- On the real callback, the caller's own unmarshal `onElement` still receives every element in
  post-order, and the caller's marshal `onElement` every written element.
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
`onUnexpectedAttribute`, and gave only the marshaller an `onElement`.

So `@docx4j/jsonix` gains an unmarshal-side `onElement(element, value)`, the counterpart of the marshal
callback: jsonix-CR-008 (`docs/change-requests/jsonix-CR-008-unmarshal-on-element.md` there). It
has the shape this CR asked for:
- **The option:** a per-call option of `createUnmarshaller`, beside the other two callbacks.
- **When it fires:** once for each element that a class info unmarshals to a typed object, after
  the object is complete. By then its attributes, content and `PARENT` are set, and its children
  have already been reported.
- **What `node` is:** the source element itself, so `lookupNamespaceURI` sees its ancestors.
- **Not reported:** simple-typed elements, element wrappers, and wildcard DOM content. That content
  stays the source node itself in the tree, which section 3.3 relies on. The typings promise this.
- **Under `xsi:type`:** the derived type's object is reported once.

The callback has the same name as the marshal one. So a caller who passes one options object to
both an unmarshal and a marshal now gets reports from both (measured: 9 and 9 for a nine-element
part). The facade keeps the two option objects separate.

The facade passes its callback to every unmarshaller it creates, composed with the caller's own.
That covers `unmarshalString`, `unmarshalNode` and each part of `unmarshalPackage`, and through
`unmarshalNode` also `builders/wml`'s `wml()` and core-ts. An unmarshaller a caller creates directly
from the context records nothing, which is one more reason to unmarshal through the facade.

The rejected alternative, which the prototype used to stand in for the callback, is a wrapper on
the `unmarshal` of the `org_docx4j_mce.AlternateContent.Choice` class info. It would have shipped
without a runtime release, but it patches runtime internals: the method's signature and
`Jsonix.XML.Input`'s `node`. Jason chose the callback (2026-10-20).

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

Released in 0.4.3 (tag 3b4f4ba, 2026-10-20), a patch, as 0.4.2 was: the fix changes output only
for parts with these shapes. It raised this package's floor to `@docx4j/jsonix` ^3.5.0, which carries
section 3.2's callback, as 0.4.1 raised it to 3.4.2.
- **core-ts's go:** core-ts tested 9ecd17a first. 877 of 878 tests passed, the one failure being
  its pin of the old behaviour. Its survey of 750 packages differed only in the four copies of
  docx4j's probe, each by the two root declarations it gains.
- **Release step 6:** the npm tarball's hashes are identical to the tree's.

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

- **Smoke, in CR-001 section 3.2's block:** `kept('zz', 'xmlns:zz="urn:zz"')` asserted that the
  `Requires` was kept and warned about, which pinned the defect. It now asserts the prefix is
  declared from the source with no warning. The "nothing resolves" case is now a Choice whose
  source declares nothing.
- **Smoke, CR-011's block:**
  - the twelve cases of section 2, with the warnings each should or shouldn't give;
  - a caller's `onElement` on both sides, the unmarshal one seeing children before parents.
  - Not tested: that a part with only the table's bindings leaves no record. That is internal state,
    and section 2's cost measurement is what shows it.
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
- **`@docx4j/jsonix`:** jsonix-CR-008, the unmarshal-side `onElement` (section 3.2), released in
  3.5.0; this package's floor is now ^3.5.0.
- **docx4j:** CR-052 (see the header) covers docx4j's half: moved, copied and built content, and a
  tree set from code. Where a source-only `Requires` prefix is declared is for Jason. There are
  three places, all equivalent XML, and Word 365 reads all three alike (core-ts check 50, below).
  - **The root:** this CR, following CR-001's convention.
  - **The `mc:AlternateContent`:** CR-052's D4 recommendation, and what docx4j does for a loaded
    part.
  - **The `mc:Choice` that names it:** what Word does. core-ts counted, over 2,728 Word 365 saves
    (24,075 WordprocessingML parts, 2026-10-20), where Word declares. A `Requires` prefix in Word's
    fixed root set goes on the root, 664 of 664. One outside it goes on the Choice: `zz9`, 3 of 3,
    in Word's resaves of docx4j's probes (build 16.0.20326), whose sources declare it on
    `mc:AlternateContent`. Word moved the declaration onto the Choice and onto a sibling
    `mc:Fallback`, and didn't add `zz9` to `mc:Ignorable`. In a WordprocessingML part, Word declares
    nothing on `mc:AlternateContent`.

  Jason ran core-ts's check 50 in Word 365 on 2026-10-20. Its file has five runs, each with a text
  box as its Choice and plain text as its Fallback. Case 1 is Word's own `Requires="wps"`, declared
  on the root. Cases 2 to 4 bind the `wps` namespace to another prefix, used on the shape elements
  too, and declare it on the root, on `mc:AlternateContent` and on `mc:Choice` respectively. Case 5
  requires `zz9`. Word showed the text box in cases 1 to 4 and the Fallback in case 5. So Word reads
  this CR's root placement, and the choice between the three is about matching Word's output, not
  about Word reading them. On save, Word renamed the three aliases to `wps` on the root and resolved
  case 5. Each kept `mc:Fallback` carried unused copies of the source's declarations outside Word's
  fixed set: `zz9` and the root alias from the root, and case 3's from its `mc:AlternateContent`.

  Either local form would also keep both meanings in section 3.5's case of one prefix bound two
  ways, because the tree doesn't use that prefix. If the ports should agree, the root rule in
  section 3.4 would change, for a prefix the tree doesn't use, to the place Jason chooses. For a
  prefix in the table, which CR-001 governs and this CR doesn't change, the root matches Word in
  WordprocessingML parts but not in chart parts: there Word declares `xmlns:mc` on
  `mc:AlternateContent` and `c14` on the Choice (37 of 37), as Excel and PowerPoint do. core-ts
  records the counts and check 50 in its CR-010 section 14.9, and sent the counts to docx4j-25 for
  CR-052.
