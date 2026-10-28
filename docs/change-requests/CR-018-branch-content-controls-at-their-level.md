# CR-018: A content control in a kept `mc:AlternateContent` branch, typed at its host's level

**Status:** Implemented 2026-10-28 on `@docx4j/jsonix` 3.11.0, unreleased, for 0.7.0 with CR-017 (sections 9 and 10).
Accepted the same day, B1 to B4 as recommended (Jason: "CR-018 decisions as recommended"; "then implement CR-018").
Proposed the same day, and revised before the decision to take a runtime hook for both B1 and B2 (Jason: "B2 if it
would be neater to ask for a jsonix change, we can do that"). The hook was prototyped, then measured against jsonix's
candidate (section 4.1).
**Requested by:** Jason, 2026-10-28 ("write the CR please"). The gap came up while CR-017's branch fix was being
implemented (202b58d). docx4j-12 then answered that docx4j isn't exposed, so the mistyping is this package's own
(CR-017 section 3, d4d34d2).
**Depends on:** jsonix-CR-014, `@docx4j/jsonix` 3.11.0 (released 2026-10-28, tag 222c2bd). It adds `wildcardTypeInfo`,
a per-call option through which a caller declares the type of an element a wildcard holds, on unmarshal and on marshal
(section 4.1). objects' downstream check of its candidate passed the same day, and the published `jsonix-factory.js` is
byte-identical to the bytes tested. The fallback not taken (B1 (b), B2 (b)) would have needed nothing new. B3 changes
CR-017, which is implemented and unreleased.
**Counterparts:**
- docx4j CR-051 section 11.3 (`McBranchSdt`, 17.3.4, phase 1, 2026-10-20) is the same fix in Java. docx4j binds a
  `w:sdt` directly in a branch by the kind its container asks for (section 3). The route decided is the same
  design: the declaration is chosen before the element is read.
- core-ts is the consumer (section 8).

## 1. Summary

The runtime types a kept `mc:Choice` or `mc:Fallback`'s children by their global declarations, since a branch's content
is a lax wildcard (CR-009). The global `w:sdt` is the block-level class, `SdtBlock`. So a content control in a branch is
always block-level, whatever its host:
- a run-level control in a paragraph's branch;
- a row-level control in a table's branch;
- a cell-level control in a row's branch.

Everything below such a control is typed at block level too. A run-level control's text is then left out by `textOf`
("before  after", where the same control outside a branch gives "before IN CONTROL after"). Code that asks a control's
class is told the wrong level. And CR-017 has to allow for the wrong class when it tests a branch.

docx4j fixed the same gap in its CR-051 section 11.3, by binding such a control by its container's kind. This CR does
the same here:
- **The runtime** resolves an element a wildcard holds through one lookup, on each side. jsonix-CR-014's
  `wildcardTypeInfo` lets a caller answer first.
- **The facade's answer:** a `w:sdt` directly in a branch is declared as its host's level asks (`SdtRun`, `CTSdtRow` or
  `CTSdtCell`).
- **On read,** the control and everything below it are then typed by the right lists from the start.
- **On save,** that same declaration means the runtime writes no `xsi:type`.

Prototyped by patching the runtime (section 4.1), the hook gives:
- the right classes at each level, a branch inside a branch's control included;
- `textOf`'s text back;
- saves byte-identical to today's.

Without a runtime change, the facade alone can reach the same result (section 4.2):
- retype the values in place after the read, which is exact because the four classes of each family share their
  properties;
- remove the `xsi:type` the runtime writes on save.

Either way, CR-017's two allowances have nothing left to allow for, and can retire (B3).

## 2. What happens today (measured 2026-10-28 at d4d34d2)

### 2.1 The typing

| A branch in | Its `w:sdt` reads as | Below it | The host's own list, and docx4j since 17.3.4 |
|---|---|---|---|
| a paragraph, a hyperlink, a simple field, `w:ins`/`w:del`, a smart tag (run level) | `SdtBlock` | `SdtContentBlock`; custom XML in it is `CTCustomXmlBlock`, a control in it `SdtBlock` again | `SdtRun`, `CTSdtContentRun`, `CTCustomXmlRun`, `SdtRun` |
| a table (row level) | `SdtBlock` | `SdtContentBlock`, its `w:tr` a `Tr` by the global declaration | `CTSdtRow`, `CTSdtContentRow` |
| a row (cell level) | `SdtBlock` | `SdtContentBlock`, its `w:tc` a `Tc` | `CTSdtCell`, `CTSdtContentCell` |
| the body, a cell, a header, ... (block level) | `SdtBlock` | right | the same |

ECMA-376 Part 3 has a reader take a Choice it understands as if its content stood in the host. So the host's level is
the right one, and docx4j binds by it. `w:sdt` is the only element whose class a branch gets wrong. Every other element
that a CR-017 container's list names falls into one of two groups (measured over all 22 lists):
- **Declared globally with the class the list gives:** the paragraph, the run, the table, the bookmarks, CR-009's 61,
  and M7's nine.
- **Not declared globally:** `w:customXml`, `w:moveFrom`, `w:moveTo`, the move and custom XML range markers, and math in
  `w:ins`/`w:del`. These stay DOM in a branch, as in docx4j (its CR-051 section 7).

### 2.2 What it costs

- **`textOf`** reads a paragraph's run-level controls but not a block-level one. So it leaves out the text of a
  run-level control in a paragraph's branch. For a control holding a run, and custom XML with a nested control:
  "before  after" today, "before IN CONTROL NESTED after" when typed right. At row and cell level `textOf` reads no
  branch at all, however typed (section 5), so the typing costs it nothing there.
- **The class.** `TYPE_NAME` is docx4j's class, which the declarations promise (CLAUDE.md, "What the declarations
  promise"). A caller asking "is this a run-level control?" is told block-level. core-ts keeps every branch on load (its
  preprocessor, CR-010 phase 4), its content API reads the selected one, and several of its sets key on the class
  (section 8).
- **CR-017** has two allowances that exist only for this (its section 3). A branch's children are tested leniently: one
  is legal if the host's list names it. Sdt content or custom XML below a branch is tested at the level the host gives.
  The leniency costs a report: a block-level control that code puts in a paragraph's branch isn't reported, since the
  paragraph's list names `w:sdt`.
- **The save** costs nothing today. The global declaration is the class that was read, so no `xsi:type` is written, and
  a round trip is byte-identical.

### 2.3 Where it occurs

No measured corpus holds the form:
- **The fidelity corpus:** of the 341 parts, the one with both a `w:sdt` and a branch has the control outside the
  branch.
- **core-ts's survey** of 9,075 parts at bb18596 reported nothing outside check 72's files. CR-017 had no allowances
  yet at that commit, and it reports such a control twice at each of the three levels (rebuilt and measured
  2026-10-28): the control on the name rule, and its content's row, cell or run as misplaced in block-level sdt
  content.
- **docx4j** found none in CR-021's 835 files or in Jason's probes (its CR-051 section 11.2).

It comes from producers that wrap their own content in a branch, and from code: core-ts's `insertOoxml` keeps a
fragment's branches (CR-011).

## 3. What docx4j does

docx4j's measured failure was worse (its CR-051 section 11.2). JAXB's lax binding picks one of the four `w:sdt` classes
for every level, so a block control in a block branch silently lost its paragraphs, and with a `w:tag` the part failed
to load. Its fix, `McBranchSdt`, works in two steps:
1. **On the way in,** a `w:sdt` directly in `mc:Choice` or `mc:Fallback` is put in a private namespace, so the lax list
   keeps it as a DOM `Element`. This is done on every unmarshal route: a `StreamReaderDelegate` on StAX, a rename on
   the DOM paths.
2. **Then it is bound by its container's kind,** with `unmarshal(Node, Class)`: `SdtBlock`, `SdtRun`, `CTSdtRow` or
   `CTSdtCell`. One that can't be bound goes back to the `w` namespace and stays DOM, so the private name is never
   written.

docx4j-12 measured CR-017's test document on 2026-10-28, in both of docx4j's load modes. Nothing was refused, and the
save was byte-identical.

JAXB gives docx4j no way to choose a lax element's class while reading, hence the rename. The runtime here has one place
where it chooses (section 4.1), so a hook there does in one step what docx4j does in two.

## 4. Proposal

### 4.1 A jsonix hook, and the facade's answer (B1 (a), B2 (a), decided)

**Where the runtime decides.** An element in a wildcard (`AnyElementPropertyInfo`, the content of `mc:Choice`,
`mc:Fallback`, `w:pict`, an extension list, ...) is resolved through one lookup on each side:
- **On unmarshal,** `getTypeInfoByInputElement`: an `xsi:type` if the source has one, else the element's declaration,
  in the property's scope and then globally.
- **On marshal,** `convertToTypedNamedValue` resolves the declaration by name. `marshalElement` then writes the value's
  own class as an `xsi:type` wherever it differs from that declaration.

**jsonix-CR-014** (2026-10-28, at this CR's ask; W1 to W3 decided, implemented as the 3.11.0 candidate) adds
`wildcardTypeInfo`, one per-call option on `createUnmarshaller` and `createMarshaller`, called at those two lookups:
- **Its arguments (its W1):** `(name, parent, propertyInfo, classInfo, elementOrValue)`, the same on both sides. The
  parent is the element's parent in the source on unmarshal, and the parent being written on marshal, so one function
  serves both factories. The last argument is the element being read, or the value being written.
- **Its answer:** a type info is the declared type, and `undefined` or `null` gives today's lookup. On marshal an
  `xsi:type` is written only where the value's class differs from the answer.
- **What comes first (its W2):** an `xsi:type` in the source that resolves still comes first, and the option isn't
  called then. Office never writes one.
- **Where it is called (its W3):** only for a typed `AnyElementPropertyInfo`. It isn't called for DOM-only wildcards,
  for the root, or for an element-ref list's fallback to a global declaration.
- **When:** on unmarshal before the element is read, so before its `onElement`; on marshal before `omitElement` and
  before the element is written.

Unset, nothing changes and nothing is called.

**The facade's answer** is a name test, then a walk over the ancestors (`branchSdtType`, `src/branches.mts`, the walk
CR-017's `effectiveContainer` made):
- Only a `w:sdt` whose parent is an `mc:Choice` or `mc:Fallback` is answered.
- **Skipped in the walk:** `mc:Choice`, `mc:Fallback`, `mc:AlternateContent`, `w:sdt`, `w:sdtContent` and
  `w:customXml`.
- **Fixing the level (`HOST_LEVEL`):** a paragraph, a hyperlink, a simple field, `w:ins`/`w:del`,
  `w:moveFrom`/`w:moveTo`, a smart tag, `w:dir` or `w:bdo` is run level; a table is row level; a row is cell level; the
  body, a glossary's `w:docPartBody`, a cell, a text box, a header, a footer, a note or a comment is block level. That
  is every element whose list declares `w:sdt`, at the level of the class it declares, but `w:sdtContent` and
  `w:customXml`, whose level is their host's (measured from the context; the smoke holds it so).
- The answer is that level's class: `SdtBlock`, `SdtRun`, `CTSdtRow` or `CTSdtCell`.

On read the control is then typed by that class, its content by that class's `sdtContent`, and what is below by those
lists, as outside a branch. Nothing is retyped after. On save, the declaration is the class the value has, so the
runtime writes no `xsi:type`. A control of any other class still gets one, as it would outside a branch, and CR-017
reports it.

**Prototyped 2026-10-28** by patching the two lookups on `AnyElementPropertyInfo`'s prototype in a scratch script,
with the facade's answer taking the parent on both sides, as jsonix-CR-014's W1 has it. The forms were:
- a run-level control holding a run, and custom XML with a nested control;
- a row-level control;
- a cell-level control holding a cell with a block-level control in it;
- a block-level control;
- a branch inside a run-level branch's control.

Every class came out as the host's list gives it: `SdtRun`, `CTSdtContentRun`, `CTCustomXmlRun` and `SdtRun` again;
`CTSdtRow`; `CTSdtCell`, with the cell's own control `SdtBlock`. `textOf` gave "before IN CONTROL NESTED after" and
"DEEP". The results:
- no CR-017 report;
- no `xsi:type`;
- each save byte-identical to today's save of the same input, and stable when read and saved again.

**Measured against jsonix-CR-014's 3.11.0 candidate** the same day, at jsonix's ask before its release. The candidate's
`jsonix-factory.js` has sha256 `906bc9f2...7061d`. The answer was wired into a scratch copy of the facade at d4d34d2,
composed after a caller's own:
- **The forms above,** plus a control below a paragraph inside a branch: the same classes as the prototype, without
  patching the runtime, and the paragraph's control typed by its list as before. No CR-017 report, no `xsi:type`. Each
  save was byte-identical to 3.10.0's save of the same input, and stable.
- **A caller's own answer** is taken over the facade's.
- **A control of another class** in a branch (a `CTSdtRow` in a paragraph's Choice) is written with its `xsi:type`.
  CR-017's leniency doesn't report it, which is the cost B3 removes.
- **The full suite** passes and typechecks, with the option unset and with the answer set.
- **The fidelity corpus** (341 parts, 333 marshalled): every marshalled hash is identical to 3.10.0's, unset and set.

It is always on, through every facade read and save:
- `unmarshalString`, `unmarshalNode` and `unmarshalPackage`, and `wml`, which reads through `unmarshalNode`;
- `marshalString`, `marshalNode` and `marshalPackage`.

A caller's own option of the same name is asked first, as CR-013's normaliser asks a caller's adapters first. The cost
is a name test per element a wildcard holds.

### 4.2 The fallback, not taken: without a runtime change (B1 (b), B2 (b))

The facade alone gives the same classes and output, measured the same day. It stays recorded in case jsonix-CR-014
doesn't ship.

**On read, retype in place.** A first step in the composed unmarshal `onElement` runs when a value of a level-dependent
class closes (`SdtBlock`, `SdtContentBlock` or `CTCustomXmlBlock`). It walks the source element's ancestors as above.
If the walk crossed a branch and the level isn't block, `TYPE_NAME` becomes that level's class:

| Read as | Run | Row | Cell |
|---|---|---|---|
| `SdtBlock` | `SdtRun` | `CTSdtRow` | `CTSdtCell` |
| `SdtContentBlock` | `CTSdtContentRun` | `CTSdtContentRow` | `CTSdtContentCell` |
| `CTCustomXmlBlock` | `CTCustomXmlRun` | `CTCustomXmlRow` | `CTCustomXmlCell` |

This is exact only because of two facts, measured from the context:
- **The same properties.** The four classes of each family declare the same properties:
  - for a control, `sdtPr` (`SdtPr`), `sdtEndPr` (`CTSdtEndPr`) and `sdtContent`, which is the level's content class
    and is retyped in its turn;
  - `content` for its content;
  - for custom XML, `customXmlPr` (`CTCustomXmlPr`), `content`, and the attributes `element` and `uri`.
- **The same lists but two names.** The run-, row- and cell-level content lists (35, 29 and 29 names) give every name
  the class that `SdtContentBlock`'s list or the global declaration gave on read. The exceptions are `w:sdt` and
  `w:customXml`, and those are retyped in their turn.

The callbacks fire bottom-up, so each value is retyped as it closes, after its children and before its container. In
place keeps the object's identity, `PARENT`, and CR-011's record on a nested Choice.

**On save, remove the `xsi:type`.** The runtime writes one for the retyped control (measured: `w:CT_SdtRun`,
`w:CT_SdtRow`, `w:CT_SdtCell`). The composed marshal `onElement` removes it from a `w:sdt` directly in a branch whose
class is its host's level's. The root's `xsi` declaration is then unused, and CR-001's strip removes it. Measured at the
three levels, the save was byte-identical to today's.

This route rests on four couplings the hook doesn't have:
- the shape facts above;
- that the runtime writes the attribute before it calls `onElement`;
- an edit of the output DOM;
- CR-001's strip.

The smoke would hold all four.

### 4.3 CR-017's two allowances retire (B3)

With either route, a branch's control and everything below it carry their level's class. CR-017's two allowances then
have nothing left to allow for. Retired, a branch's child is tested as the host's own children are:
- by its class, against the host's set;
- by the name rule, against the host's list, which is the declaration the save now writes by.

`effectiveContainer` and the `lenient` parameter go from CR-017's tests; under 4.1 the walk lives on in the facade's
answer. The one change in what CR-017 reports: a block-level control that code puts in a paragraph's branch is
reported, as it is outside one: "You can't add a block-level content control (SdtBlock) to a paragraph (w:p); ...". A
file read can't produce one any more. CR-017 section 3 is revised to say so.

### 4.4 What doesn't change

- **No option of the facade's own.** This is always on, as CR-009's typing is.
- **Untouched helpers:** `textOf`'s rules, `mcBranchOf`, the builders, `deepCopy`.
- **What stays DOM:** a `w:customXml` directly in a branch, as in docx4j.
- **The corpora.** Neither holds the form, so nothing in them changes. Every fidelity part's marshalled hash must stay
  identical (341 parts, 333 marshalled; measured so against the candidate, section 4.1).
- **What `textOf` returns** changes only where a run-level control was mistyped. It now reads as it does outside a
  branch.

## 5. Not proposed

- **`textOf` at block, row and cell level.** Found while measuring, and not about typing. `textOf` treats every
  `mc:AlternateContent` as run content (`RUN_CONTAINERS`). It gives:
  - "before\nONETWO\nafter" for two paragraphs in a body branch;
  - "before\n\nafter" for a block-level control there, typed right or not;
  - "A1\n" for a table whose second row is in a branch.

  docx4j's `TextUtils` reads such content after its load resolves the branch. Its output is kept stable because core-ts's
  readers match it (Jason, 2026-10-19), so a change is its own question, for Jason and core-ts.
- **Typing a `w:customXml`, `w:moveFrom` or `w:moveTo` directly in a branch.** None is declared globally, and none is a
  docx4j root, so each stays DOM in both ports. The hook could type them, but that would be a divergence beyond CR-009
  and M7.
- **docx4j's route, a private namespace before the read.** It would need a DOM pass before every read, plus a restore
  of a caller's DOM, for what the hook does at the lookup.

## 6. Decisions

**Decided 2026-10-28 as recommended** (Jason: "CR-018 decisions as recommended"): B1 (a), B2 (a), B3 (a), B4 (a). The
options and reasons stand below. jsonix-CR-014's W1 to W3 were decided the same day, in jsonix.

B1 and B2 go together. The hook serves both sides, and taking it for one side only would leave two mechanisms for one
rule.

- **B1. The read:**
  - (a) jsonix-CR-014's `wildcardTypeInfo`, which the facade answers for a branch's `w:sdt` (section 4.1);
  - (b) retype in place as each value closes (section 4.2);
  - (c) read the branch's element again with the host's scope;
  - (d) leave the classes, and teach `textOf` and CR-017 the level.

  **Recommendation:** (a). The type is chosen before the element is read, as docx4j chooses it, so nothing depends on
  the four classes keeping one shape. (b) is the fallback if the runtime change is refused: measured exact, but only by
  those shape facts. (c) fires every callback twice, so `onWordReading` and `onOutOfRange` would report twice. (d) leaves
  `TYPE_NAME` wrong for every other caller.
- **B2. The save:**
  - (a) the same hook on marshal, so the declaration is the value's class and no `xsi:type` is written;
  - (b) the facade's marshal `onElement` removes the `xsi:type` after the runtime writes it.

  **Recommendation:** (a). It is the neater of the two, as Jason asked. (b) rests on the runtime's order of writing, an
  edit of the output DOM and CR-001's strip; it is the fallback with B1 (b).
- **B3. CR-017's allowances:** (a) retire them (section 4.3); (b) keep them, redundant. **Recommendation:** (a). They
  were there only for this, and the leniency hides a real mistake.
- **B4. The release:**
  - (a) with CR-017 in 0.7.0, so CR-017's allowances never ship, and core-ts's gate runs once for both. Under B1 (a)
    0.7.0 then waits for `@docx4j/jsonix` 3.11.0;
  - (b) 0.7.0 as it stands, and this in 0.8.0.

  It changes classes and `textOf`'s output, so it is a minor either way. **Recommendation:** (a), unless 0.7.0 is
  wanted before jsonix can release. jsonix-CR-013 went from proposal to release in a day. 0.6.1 (73b9606) is unaffected
  either way.

## 7. Tests

The smoke gains a CR-018 block:
- **The context.** `w:sdt` is the only name in a CR-017 container's list whose global declaration gives another class.
  A regeneration that adds a second fails here, since the answer covers `w:sdt` alone. Under B1 (b) the smoke also holds
  4.2's two shape facts.
- **The read**, at each level:
  - a run-level control holding a run, and custom XML with a nested control;
  - a row-level and a cell-level control;
  - a block-level one, which is unchanged;
  - a branch inside a branch's control;
  - a control below a paragraph inside a branch, typed by the paragraph's list as today.

  It also checks the classes, `PARENT`, a caller's `onElement` seeing the final class, and a nested Choice's CR-011
  record.
- **`textOf`:** "before IN CONTROL NESTED after".
- **The save:**
  - each level's part saved as read: canonically equal to the source, with no `xsi:type`;
  - a control of another class in a branch keeps its `xsi:type`.
- **CR-017:**
  - its branch regressions stand;
  - a block-level control built into a paragraph's Choice is reported, on save, with docx4j's message (B3).

`fidelity.mjs` passes unchanged, with every fidelity part's marshalled hash identical. The jsonix CR's own tests are
jsonix's.
core-ts reruns its survey (35 reports expected, all from check 72) and its readers (section 8).

## 8. Consumers and counterparts

- **core-ts** keeps every branch on load and reads the selected one (its preprocessor's header, CR-010 phase 4). Its
  sets key on the class:
  - `Paragraph.mts` makes an `SdtBlock` a block control's view;
  - `tableRules.mts` and `tree.mts` count `SdtBlock` among the blocks;
  - `spanEdits.mts` treats it as a control.

  A branch's control now reaches these at its level, and its readers that match `textOf` change with it. Which readers
  change, and whether any of its checks expected the old class, is core-ts's to measure. Under B3 (a), its CR-016
  phase 2 refuses a save holding a block-level control in a paragraph's branch, as it would outside one. With
  `mcePreprocess: true` it isn't exposed: the branch is resolved on the DOM, and the control is typed by its new host's
  list.
- **jsonix:** jsonix-CR-014 is the hook (B1 (a), B2 (a)). Its W1 to W3 are decided, and its 3.11.0 candidate passed
  objects' downstream check (section 4.1). It has no other consumer today. The facade's answer is objects'.
- **docx4j** has section 11.3 already, so nothing is asked of it. This CR brings the two ports' classes level.
- **docx4j-python** has its own model, so its exposure is its own question.

## 9. Release

A minor, since classes and `textOf`'s output change where the form occurs. It rides in 0.7.0 with CR-017 (B4 (a)), and
the floor of `@docx4j/jsonix` rises to ^3.11.0 (B1 (a)). The order:
1. jsonix releases 3.11.0 (done 2026-10-28);
2. this CR is implemented, with the floor raised and the lockfile updated (done 2026-10-28, section 10);
3. core-ts reruns its suite and survey against `main`, and says go (RELEASING.md, and Jason's rule);
4. 0.7.0 is released.

0.6.1 (73b9606) can still be cut alone before then.

## 10. Implemented (2026-10-28)

B1 to B4 as recommended, on `@docx4j/jsonix` 3.11.0 (`package.json` ^3.11.0; the lockfile at 3.11.0, whose
`jsonix-factory.js` has sha256 `906bc9f2...7061d`):
- **`src/branches.mts`** (internal, not an exported path): `HOST_LEVEL`, `SDT_AT_LEVEL` and `branchSdtType(name,
  parent)`, which gives the class for a `w:sdt` whose parent is an `mc:Choice` or `mc:Fallback`, and nothing for any
  other element.
- **The facade** (`src/index.mts`): `branchTypes` composes a caller's `wildcardTypeInfo` first, then the facade's
  answer, with the type info looked up on the context the unmarshaller or marshaller came from. It is passed by every
  facade read (`recordingOptions`, which now takes the context: `unmarshalString`, `unmarshalNode`, `unmarshalPackage`,
  and so `wml`) and every facade save (`marshalToDocument`: `marshalString`, `marshalNode`, `marshalPackage`).
- **B3:** `LEVEL_OF`, `BY_LEVEL` and `effectiveContainer` are gone from the facade. So is `testList`'s and
  `misplacement`'s `lenient` parameter, with `Declarations.listDeclares`, which only it used. A branch's children are
  tested as the host's own, on read and on save.
- **The smoke** gains section 7's CR-018 block. In it, `HOST_LEVEL` is held to the context, and `w:sdt` is checked to
  be the only name a CR-017 list declares otherwise than the global declaration. CR-017's existing branch cases pass
  unchanged without the allowances. A block-level control built into a paragraph's Choice is reported on save with
  docx4j's message, and written with `xsi:type="w:SdtBlock"`, the mapping's name for the type.
- **The gate:** `npm run typecheck` and `npm test` pass. Every fidelity part's marshalled hash (341 parts, 333
  marshalled) is identical to the build before this CR on 3.10.0. `npm pack` gains `dist/branches.mjs` and its
  declarations, 851 files.

Also measured through the facade, beside the smoke: each form's save is byte-identical to the 3.10.0 build's save of
the same input, and stable when read and saved again.
