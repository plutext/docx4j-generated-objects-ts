# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

`@docx4j/generated-objects-ts`: Office Open XML as typed JavaScript. `modules/` holds 103 Jsonix mappings
(UMD `.js`, ES module `.mjs`) with TypeScript declarations (`.d.ts`, `.d.mts`), **generated** by
[jsonix-schema-compiler](https://github.com/plutext/jsonix-schema-compiler) from
[docx4j](https://github.com/plutext/docx4j)'s `xsd/ROOT.xsd`. `src/` holds the only hand-written
code: a facade with docx4j's names (`index.mts`), the module registry (`modules.mts`), `wml-globals.mts` (CR-009: one more
mapping declaring globally the 61 WordprocessingML elements docx4j declares as roots, so kept `mc` branches type their
`w:drawing`/`w:pict`, and nine more beyond docx4j's roots, CR-017 M7, so a misplaced one reads typed; `getContext` and
`modulesFor` add it), `misplaced.mts` (CR-017, internal: docx4j CR-064's `checkParent` sets and messages, which the
facade's `onMisplaced` reports and the builders refuse with), `helpers/wml.mts` (per-type docx4j decisions,
compiler CR-007; Word 365's colour reading, `WORD_COLOR_NAMES`/`wordColor`, CR-013; Word's highlight palette and the
highlight Word stores for a colour Office JS writes, `highlightNameForColor`, CR-014; Word's storage of each integer
attribute and what it reads, `WORD_INTEGER_STORAGE`/`wordIntegerReading`, CR-016) and `builders/wml.mts` (CR-002: `wml` fragments, text sugar, the run mapping,
`textOf`, `runItemsOf`, content controls, rows and cells, inline pictures, the `w:rPrChange` element
list, `walkAll` and `mcBranchOf` (CR-003 phase A and section 3.8), traversal; imports the helpers, never the reverse). The runtime is
[`@docx4j/jsonix`](https://github.com/plutext/jsonix) 3.10.0+ (on/off booleans, jsonix-CR-007; base64Binary in strict mode, 3.4.2; the unmarshal-side `onElement` CR-011 records through, jsonix-CR-008, 3.5.0; `onInvalidValue`, which CR-013's Word reading goes through, jsonix-CR-010, 3.7.0; the marshal-side `onAttributeValue` and `omitElement` CR-013's save normaliser goes through, jsonix-CR-011, 3.8.0; `strictIntegers`, which CR-015's phase 2 passes, jsonix-CR-012, 3.9.0; `onUnexpectedText`, which CR-017's `'#text'` read report goes through, jsonix-CR-013, 3.10.0). This is the counterpart of docx4j's
`docx4j-generated-objects` module and is usable on its own (Office JS add-ins); the engine layer
(OPC packaging, parts, style/numbering resolution, the counterpart of `docx4j-core`) is the separate
repository `plutext/docx4j-core-ts`, package `@docx4j/core-ts`, which depends on this one. Tree-only helpers belong here;
anything needing parts or relationships belongs there.

## Commands

```
npm ci              # typescript, esbuild, @xmldom/xmldom and the @docx4j/jsonix runtime, as locked in package-lock.json (committed)
npm run build       # tsc -p tsconfig.build.json: src/ -> dist/ (index, helpers/wml, builders/wml: .mjs and .d.mts)
npm run typecheck   # tsc --strict, noEmit, over modules/*.d.ts, modules/*.d.mts, src/, test/*.ts
npm test            # build, the nodenext consumer check, then test/smoke.mjs, fidelity.mjs, word365.mjs, word365-integers.mjs, word365-ranges.mjs, bundle.mjs
node test/smoke.mjs # the single runtime test, when dist/ is already built
npm pack --dry-run  # ships dist/, modules/, LICENSE, NOTICE, README.md, package.json only
```

`@xmldom/xmldom` is declared although `@docx4j/jsonix` brings it: five tests import it, and installing a jsonix
checkout with `npm install --no-save` (jsonix's downstream check) prunes a package nothing here declares.

`test/lib/marshal-hashes.mjs` hashes the marshalled output of every fidelity part for the build
`OBJECTS_TS_DIST` names, defaulting to `dist/`; `RELEASING.md` step 6 runs it against the published
tarball and diffs, which is the only check that the artifact npm served is the one that was tested.

`prepublishOnly` runs `typecheck` then `test`. CI (`.github/workflows/test.yml`) runs both on Node
18, 20 and 22, after `npm ci`; `npm update` then a lockfile commit picks up newer dependency versions.
Releases publish to npm from `.github/workflows/push-to-npm.yml` on a GitHub release (trusted
publishing, tag = `package.json` version); see `RELEASING.md`. The GitHub CLI `gh` is installed on
this machine and logged in as `plutext` (scopes `repo`, `workflow`; confirmed 2026-10-04), so CI
results (`gh run list --commit <sha>`) and RELEASING.md step 5's `gh release create` work from the
shell.

**Never release without core-ts's confirmation.** Before `npm version` (step 2 of `RELEASING.md`),
ask a `docx4j-core-ts` agent (`SendMessage`, see `ListAgents`) to run its typecheck and full test
suite against the proposed release (the pushed `main` commit, by `file:` or `git` dependency), and
wait for its reply saying it is good to release. This package's tests are a self-written
`document.xml` and two compile-only checks; core-ts's fixtures are real Office parts and its tests
take `mc:Choice` branches, so it sees what the smoke cannot. 0.1.5 (2026-09-20) went out without
that check and regressed every pptx/xlsx with an equation in a text body (`a14:m` admitted, `a:rPr`
unknown to OMML's `CT_R`), while core-ts 0.1.0's `^0.1.1` range pulled it into fresh installs.

There are eight tests and no test framework:

- `test/smoke.mjs` (runtime, plain `node:assert`) unmarshals `test/fixtures/document.xml` through the facade and checks `TYPE_NAME`,
  `PARENT`, `deepCopy` (children re-linked, the copy's own `PARENT` unset), a marshal round trip, the
  `v:line` attribute order `id style from to`, that mappings load via `require()` (directly and
  by package self-reference) and via `import()` through `./modules/*`, and the package round trip
  (`unmarshalPackage` types the known parts, `marshalPackage` leaves its input untouched).
  CR-002's block checks `builders/wml`: every wrapper row, the content control inference (nested,
  forced), the tagged form, the sugar, the run mapping round trip, `textOf`, `walk` / `find` /
  `linkParents`. CR-009's block checks that each of the 61 resolves globally to its type and that
  the generated module still declares none of them itself (so a regeneration that does retires the
  entry), that `modulesFor` adds the mapping with wml and not without, and that a kept Choice's
  `w:drawing` and Fallback's `w:pict` (over a `v:shapetype` with `w14:anchorId`) unmarshal typed and
  marshal whole. docx4j CR-051's block checks that the block, row and cell content lists name
  `mc:AlternateContent` (docx4j 1cff2cd5c), and that one at each level unmarshals typed and marshals
  whole. CR-011's block checks that a prefix only the source bound is declared from the
  source: on the root, on `mc:AlternateContent`, a second prefix for a known namespace, after a
  move between trees, a `wml()` fragment, `deepCopy`, a flat OPC part, a paragraph marshalled
  alone; the superseded, tree-used and two-ways cases; and that a caller's `onElement` is composed.
  It also checks where each declaration goes (CR-012): a `Requires` prefix on its Choice, ahead of
  `Requires` and not on the root, unless the root declares that binding already; the same for a
  Choice's own `mc:Ignorable` and `mc:MustUnderstand`, and what happens to one nothing resolves.
  CR-013's block checks `readRunOptions`' highlight and colour readings, `wordColor`, the table's 140
  names (held to docx4j's `WordColorValues` when `../docx4j` is checked out), and the facade's
  `onInvalidValue`: refused booleans on, refused enumeration values dropped, `w:u`'s on/off words
  `none`, the scope (w14's `CT_OnOff`; DrawingML and other simple types still fail), every facade
  unmarshal, the `onWordReading` record and a caller's own callback taking precedence. Its phase C
  block writes docx4j's `ColourNormalFormTest` synthetic cases the same (a theme colour, revisions'
  and a paragraph mark's former properties) and adds a style, the document defaults, a border, a
  pattern and the page colour; the tree unchanged, the 19 records and their paths, the opt-out per
  call and process-wide, a caller's adapters asked first, `marshalNode` and `marshalPackage`.
  CR-014's block holds `HIGHLIGHT_COLORS`, `readRunOptions`, `highlightNameForColor` and `applyRunOptions` to
  core-ts's checks 61 (Word 365 through Office JS: 18 reads, 45 writes) and 62 (7,225 colours written to
  `font.highlightColor`, every one stored as Word stored it), copied to `test/fixtures/check61/` and `check62/`
  with a `SOURCE.md` each, and checks CSS's 147 names and the forms refused.
  CR-015's block checks that every integer attribute of Word's own mappings (178, found from the context) reads Word's
  forms as their integer part and drops the forms Word refuses, while other mappings still fail. It also checks the
  record, a caller's precedence, that a save writes integers while a caller's own non-integer still throws, and that
  an empty rsid is left out on save with its path.
  CR-016's block checks that each entry of `WORD_INTEGER_STORAGE` (49, 3 unmeasured after round 6; `w:legacy`, which Word drops, left out) is an integer `w:` attribute of
  its class, so a regeneration that renames one fails. It also checks `wordIntegerReading`'s wrap, clamp, rule and
  certain range, and `onOutOfRange` on read and on save. Nothing is rewritten, with or without the normaliser. The
  builders' `RangeError` is checked at the boundaries.
  CR-017's block holds `misplaced.mts`'s tables to the context: each container's list, no subtype of a container or a
  refused class, each refused class's element names, and `w:sdt` as each list declares it. It checks that `p`, `tc`,
  `tr`, `sdt` and `wml`'s interpolation refuse with docx4j's messages and accept the legal children. On save it checks
  each mistake's path, rule and message, property elements counted in a step, a string reported before the runtime's
  throw, a callback's throw, the name rule (and the class rule first), and sdt content in sdt content. On read it checks
  the same file, property elements absorbed, mc branches by their host's set, M7's `w:instrText`, and core-ts's
  check 72 forms.
- `test/readme-examples.ts` (compile-only, via `typecheck`) holds the README snippets against
  minimal Office JS stubs. Change a README example and this file together.
- `test/fidelity.mjs` (runtime, CR-004 phases A and B) unmarshals and marshals **341 parts** Office 365
  wrote - every XML part of seventeen documents from docx4j's test resources, under
  `test/fixtures/fidelity/<document>/<part path>`, each directory with a `SOURCE.md` naming the file
  and commit - and compares canonically with `test/lib/canonical.mjs` (QNames not prefixes,
  attributes sorted, namespace declarations ignored, `1`/`true`, two spellings of one
  `xsd:double` and base64 with and without line breaks the same value, whitespace-only text dropped, element order compared **except** under
  the two `xsd:all` roots of `docProps`). Parts are read by their own encoding, since one is UTF-16.
  Every part carrying an `mc:AlternateContent` is also checked with its Choices resolved, as a
  consumer understanding `test/lib/understood.mjs` would resolve them - the stronger check, since
  resolving converts DOM into typed content and so can expose a binding gap the unresolved form hid
  (CR-004 sections 10 and 12). Differences there are classified by the unmarshaller's own
  `onUnexpectedElement`, not by reading the diff. A part that throws is a failure.
  Three tables record what is not a new finding, each strict in both directions so it cannot outlive
  its reason (the device is core-ts CR-001 section 19): `KNOWN`, a difference in one part, matched on
  the difference itself; `KNOWN_MISSING_ATTRIBUTES`, an attribute the model does not bind and so
  drops everywhere (seventeen today, all docx4j's - `a14:legacySpreadsheetColorIndex` with the
  `mc:Ignorable` beside it, and `w:stylePaneFormatFilter`'s fifteen booleans; the `xr*` `uid`
  family and `b:Sources/@Version` left the table with the docx4j 17.2.1 regeneration), which fails
  if no part loses it any more; and `UNMODELLED`, eight parts
  whose roots docx4j deliberately does not bind (sensitivity labels, PowerPoint's 2018 comments,
  Excel's threaded comments, a Power Query blob), which fails if the model ever learns one. Every
  differing line is classified, not only the first, so a recorded loss cannot mask an unrecorded one;
  the classifier has its own self-checks, as does the canonicaliser. Since the comparison ignores
  namespace declarations, every prefix an `mc:Ignorable`, an `mc:MustUnderstand` or an `mc:Choice`'s
  `Requires` names in a marshalled part is also checked to resolve where it is written (CR-001
  section 3.2).
  Run against 0.1.5's `modules/` the test reports all five losses that release carried. A part that
  reaches CR-013's `onWordReading` or `onNormalised`, CR-016's `onOutOfRange` or CR-017's `onMisplaced` fails too:
  Office writes only values the runtime reads, colours in their normal form, integers within its own storage, and no
  child where its container never admits it.
- `test/word365.mjs` (runtime, CR-013) reads Word 365's run property probes (rounds 1 to 7 of
  2026-10-24, `test/fixtures/word365-probes/`, each probe's `word/document.xml` beside Word's save
  of it and, for fifteen, docx4j's): each marker run's direct toggles, underline and its colour,
  `vertAlign`, `em`, colour, highlight, shading and border, and the page colour, must read as in
  Word's save, the oracle, as docx4j's `RunPropertiesWord365Test` has it. 388 markers, 385 equal;
  the three differences (Word's save drops a direct `w:u` repeating a style's) and the three
  readings Word's PDFs give where its save changed what it drew are stated, strict both ways. Each
  probe saved under phase C must read the same again, have every colour in the schema's form, and
  write the 251 markers docx4j saved as docx4j does, value for value; docx4j's rule and opt-out
  cases are written the same.
- `test/word365-integers.mjs` (runtime, CR-015) reads docx4j CR-065's shared Word probe (Jason, Word 365, rounds 2 to
  4 of 2026-10-26; `test/fixtures/cr065-probe/`, each probe part beside Word's save of it, and six files Word refused
  outright). Its 64 cases put a decimal or an odd form in an integer attribute: paragraph, run, border, tab, table and
  section properties, a style, the settings and a header. Each case is read through the facade and written again, and
  must write what Word's save wrote, the oracle. 50 are equal, and 6 differences are stated, strict both ways: four
  truncations to a default Word's save leaves out, and two `w:gridCol` Word recomputes. In the refused files, all 8
  forms are dropped and reported, four of them through phase 2's `strictIntegers`.
- `test/word365-ranges.mjs` (runtime, CR-016) reads docx4j CR-066's rounds 5 and 6 (Jason, Word 365, 2026-10-26;
  `test/fixtures/cr066-probe/` and `cr066-round6/`, the group files and settings singles beside Word's saves).
  - **The cases:** 652 integer cases at and beyond each attribute's storage width. For each, `wordIntegerReading` must
    give what Word's save wrote, the oracle: the value, the rule for a line, a row or a frame's height, and the
    attribute for `firstLine`/`hanging`.
  - **What it can't compare:** 16 read as 0 and are left out by Word. 30 are stated: `w:gridCol`, which Word recomputes,
    and `w:wBefore`/`w:wAfter`, which it dropped.
  - **The report:** `onOutOfRange` must report exactly the 432 cases Word reads differently, on read and on save, and the
    save must write every value as held.
  - **Group E:** round 6's group E holds CR-015's no-digit forms to Word's 0.
- `test/bundle.mjs` (runtime, CR-005) bundles an `import` consumer and a `require()` consumer with
  esbuild, so a bundle takes the registry's mappings and no more; its ESM consumer reads and writes
  `o:gfxdata` in strict mode (base64Binary, jsonix 3.4.2).
- `test/nodenext/consumer.mts` (compile-only, via `npm test` after the build) imports every public
  path by the package's own name under `module`/`moduleResolution: nodenext`, as a Node ES module
  consumer does. The repository's tsconfigs use `bundler`, which accepts extensionless relative
  imports in the emitted declarations; `nodenext` rejects them, and with `skipLibCheck` they
  silently become `any` (0.1.0 shipped that in `helpers/wml` and `builders/wml`). Import `modules/`
  from `src/` with the `.mjs` extension. Its `@ts-expect-error` lines go unused, and fail, if those types
  resolve to `any`.

Fixtures are extracted part XML (and core-ts's Word measurements as JSON, CR-014; docx4j CR-065's case list, CR-015), never archives (no zip dependency) and never edited by hand; they
cost consumers nothing, since `files` ships `dist/` and `modules/` only.

`typecheck` has `skipLibCheck: false` and includes every generated `.d.ts`, so it is also the check
that a regeneration's declarations compile; `lib` includes `dom` because the runtime typings need
`Node`/`Document`.

## Layout and rules

- `modules/<Module>.factory.{mjs,d.mts}` and `<Module>.el.{mjs,d.mts}` are the generated element
  factories (compiler CR-010: docx4j's `ObjectFactory` names, `createRT`, `createPElement`,
  `el.p`), public as `./factory/*` and `./el/*`. They are generated with the modules.
- `modules/` is generated output plus a reference copy of `bindings.xjb`. **Never edit it by hand.**
  Regenerate from the compiler repository (`OfficeOpenXML/generate.sh ../docx4j/xsd/ROOT.xsd
  ../docx4j-generated-objects-ts`, see `generate.md`); the output is deterministic, so a regeneration with unchanged
  inputs is an empty diff. Commits that regenerate cite the compiler and docx4j commits used.
- `package.json` deliberately has **no `"type": "module"`**: the UMD `.js` mappings must stay
  CommonJS-loadable; ES modules are marked by `.mjs`/`.mts`. Adding it breaks `require()` of every
  mapping (ERR_REQUIRE_ESM).
- Public paths are the `exports` map only: `.` (facade), `./helpers/wml`, `./builders/wml`, `./modules/*`
  (`import` → `.mjs`, `require` → `.js`, `types` → `.d.ts`). Keep them stable.
- The facade builds one `Jsonix.Context` over all modules lazily (`getContext`, first use) with
  `parentPointers: true`; modules depend on each other, so the context is all-or-nothing.
  `resetContext()` drops the cached context so the next `getContext(options)` rebuilds it with
  those options; `getContextSync()` returns the built context or throws, for `deepCopyAsSync` and
  other callers that cannot await (CR-003 section 2). The facade is asynchronous because modules load via dynamic `import()`.
  `unmarshalPackage` / `marshalPackage` handle flat OPC packages (Office JS `getOoxml()`): parts are
  `xsd:any processContents="skip"`, hence DOM by the schema, and are converted to typed elements
  where the model knows the root element (and back to DOM on marshal, without modifying the input).
  `src/modules.mts` is the registry of the 104 mappings, imported by literal specifiers so a
  bundler can see them (CR-005); it is maintained by hand when modules appear or disappear
  (`generate.sh` does not write it), and `MODULE_NAMES` and `ModuleName` derive from it, so it is
  the only list. The smoke fails on a stale registry (a mapping references a missing dependency).
  `getContext({ modules })` builds a context over fewer mappings, with `modulesFor(root)` for the
  closure.
- `src/` and `test/` import from `../modules/...` and `../dist/...`; `dist/` is git-ignored and
  built by `npm run build`.

## Change requests

`docs/change-requests/` holds numbered proposals (`CR-NNN-<slug>.md`, indexed with status in its
`README.md`) for the hand-written part of this package only: the facade, the layout, the tests.
Proposals for the generated `modules/` go to the compiler repository's `docs/change-requests/`,
for the runtime to `plutext/jsonix`. Each CR names its docx4j counterpart, what it depends on,
the tests it adds, and the consumers it affects (`docx4j-core-ts` files CRs against this package).
Check the index before changing facade behaviour: a CR may already specify it. CR-001
(implemented 2026-09-10) makes docx4j's prefix table (`NAMESPACE_PREFIXES`) the context default;
the facade's marshal functions derive a per-root table (a `Relationships` root takes the default
namespace, since the runtime allows one default) and strip unused root `xmlns` declarations,
keeping, and declaring from the table, what `mc:Ignorable` and each `mc:Choice`'s `Requires` name
(CR-001 sections 3.1 and 3.2). CR-011 (implemented 2026-10-20) declares those from the source's own
binding first: the facade's unmarshal functions record, through jsonix-CR-008's `onElement`, each
binding the table and the aliases would not give, on the `mc:Choice`'s typed object and on the
part root (a module `WeakMap`; `deepCopy` carries it), and the marshal finds them through its
`onElement`. Unmarshal through the facade too, or nothing is recorded. CR-012 (implemented
2026-10-20, Jason's choice of Word's placement for both ports) declares such a `Requires` prefix
on the `mc:Choice` that names it, not on the root, unless the root already declares that binding,
and treats a Choice's own `mc:Ignorable` and `mc:MustUnderstand` as its `Requires` (docx4j CR-052's
D1 (b), Jason 2026-10-24); a root's `mc:Ignorable` prefixes, and the prefixes the table or aliases
give, stay on the root. CR-008 (implemented 2026-10-04) has them first append to a
part root's `mc:Ignorable` the Office extension namespaces its tree uses
(`IGNORABLE_EXTENSION_NAMESPACES`, the twenty Office lists, measured on the fidelity corpus, which
must not move). Marshal through the facade, not `createMarshaller()` directly, to get that output.
CR-013 (implemented 2026-10-24) reads run properties as Word 365 reads them: every facade
unmarshal passes `wordInvalidValue` as the runtime's `onInvalidValue` unless the caller passes one,
scoped by `classInfo.module.name` to Word's own WordprocessingML mappings, and `readRunOptions`
reads highlight and colour through `helpers/wml`'s `wordColor`. Its phase C has every facade
marshal write run properties in their normal form through jsonix-CR-011's `onAttributeValue` and
`omitElement`, keyed by type (`normalising` in `src/index.mts`; `normaliseOnSave`,
`setNormaliseOnSave`, `onNormalised`). The rules are Jason's probes'; docx4j CR-060 and CR-061 are
the Java port, value for value, core-ts CR-015 the consumer. CR-014 (implemented 2026-10-24) gives
`HIGHLIGHT_COLORS` Word's palette (`WdColorIndex`, as Office JS reads it) and has `applyRunOptions` take Office JS's
`font.highlightColor` input (a CSS name or any `#RRGGBB`), stored as Word stores it: `highlightNameForColor`'s
rule (Windows' integer HLS hue sectors, a dark bound, grey bands) with check 62's 149 grey-edge answers as data.
CR-015 (phase 1 implemented 2026-10-26) has `wordInvalidValue` read a refused value of any integer attribute of Word's
own mappings as Word 365 does, measured by docx4j CR-065's shared probe:
- Word's form (`[+-]?(digits[.digits]|.digits)`, untrimmed) gives its integer part, truncated toward zero;
- any other form, which makes Word refuse the whole document, is dropped;
- both are reported as `'value.truncated'` or `'value.refused'`.

Its normaliser also leaves out an empty `w:rsid*` on save. Phase 2 (implemented the same day) passes jsonix-CR-012's
`strictIntegers` on every facade unmarshal that uses the facade's own `onInvalidValue`. So the forms the runtime's
`Number()` would read first (`" 240 "`, `0x10`, `1e3`, `""`, `10206.0`) come to the same rule, and every other mapping
reads them through `typeInfo.parse`, unchanged. docx4j CR-065's D3 is the same rule in Java.
CR-016 (implemented 2026-10-26) gives Word's storage of each integer attribute (`helpers/wml`'s `WORD_INTEGER_STORAGE`,
from docx4j CR-066's round 5): int16, uint16, 8 or 5 bits, or int32, with clamps on tab stops and page size, and the
rule in the sign for exact lines and row heights. `wordIntegerReading` gives what Word reads for a value.
- **The report:** the facade's `onOutOfRange` reports each value Word would read differently, on read (through the
  composed unmarshal `onElement`) and on save (an adapter on `onAttributeValue`, outside the normaliser). It never
  rewrites. Each report carries the table entry it comes from (`typeName`, `property`, `storage`) and the rule held
  beside a sign-carried attribute (core-ts's ask, 2026-10-27), since an element's name does not give its class.
- **Why per attribute:** docx4j's type-bound print method can be exact only per type, but the adapters know the class
  and the property, so this report is exact per attribute.
- **The builders** throw a `RangeError` for a width or size Word can't hold.

CR-017 (implemented 2026-10-28, M1 to M7 as recommended) refuses and reports misplaced children, with docx4j CR-064's
`checkParent` sets (`src/misplaced.mts`):
- **The rules:** a child is tested by its class (its value's `TYPE_NAME`, or the class its name gives). There is also
  core-ts CR-016's name-and-value rule: a value whose class isn't the one its name declares there, which the marshal
  would write with `xsi:type`. An `mc:AlternateContent` branch is tested by its host's set.
- **The builders:** `p`, `tc`, `tr`, `sdt` and `wml`'s interpolation throw a `TypeError` in docx4j's wording. Without
  the context, the builders apply the name rule to `w:sdt` alone. The tagged form marks each interpolated element and
  refuses only those.
- **The report:** `onMisplaced` on read and on save, each container's list tested once. On save that is as the
  container is opened, so a string is reported before the runtime's throw and a callback's throw stops the save; the
  `path` is in `OutOfRange`'s form. On read, a string a file holds in a container, which the runtime drops, is
  reported as `'#text'` through jsonix-CR-013's `onUnexpectedText` (CR-017.text); text inside a class that is no
  container is not. core-ts CR-016 phase 2 refuses a save from it.
- **M7:** nine run and paragraph elements are declared globally beyond docx4j's roots (`wml-globals.mts` list C), so a
  misplaced one reads typed rather than failing its part. docx4j called the divergence harmless.
- **CR-064's addendum** (docx4j ce87ec6d7, mirrored): deleted text and a tab refused where text and a break are, and
  a content control at the wrong level by class in every set.
- **Branches:** a kept `mc:Choice` types its children by their global declarations (the global `w:sdt` is
  block-level), so a branch's children are tested leniently, legal if the host's list names them, and sdt content or
  custom XML below a branch is tested at the level the branch's host gives (`effectiveContainer`).

## What the declarations promise

Module and `TYPE_NAME` names are docx4j's packages with underscores (`org_docx4j_wml.P` is
`org.docx4j.wml.P`). Required properties are non-optional, collections are arrays, `elementRef`
properties are unions of `TypedNamedValue<T>`, enums are literal unions, `TYPE_NAME` is a literal
union over the type and its subtypes, `readonly PARENT?` is the union of possible containers,
docx4j's hand-written interfaces (`ContentAccessor`, `SdtElement`, ...) are union aliases, dates
are Jsonix calendars, and attributes are emitted sorted by name except `v:line`. The rationale for
each lives in the compiler repository's `docs/change-requests/` (CR-005 to CR-007, CR-009 for this
split) and in `plutext/jsonix`'s `jsonix-CR-001`/`002`.

## Related repositories (checked out side by side)

- `../jsonix-schema-compiler`: generates `modules/`; owns `OfficeOpenXML/bindings.xjb` and
  `generate.sh`.
- `../jsonix`: the `@docx4j/jsonix` runtime (`nodejs/scripts`), typings in `types/main.d.ts`.
- `../docx4j`: the schemas (`xsd/ROOT.xsd`) and the Java model this package mirrors.
- `../docx4j-core-ts`: the engine (`@docx4j/core-ts`), the main consumer of the facade.

## Portfolio task registry

This repository's change requests are indexed, with their dependencies on work in the other
docx4j repositories, in `../docx4j-portfolio/tasks.yaml` (ids `<repo>/<CR>[.<phase>]`; this
repository's key is `objects-ts`).

- **Registry changes go through the docx4j-portfolio session** (Jason's rule, 2026-10-17). Do
  not edit `../docx4j-portfolio/tasks.yaml`, `scripts/status-snapshots.yaml` or
  `docs/dependencies.md` directly. When a CR's status changes (a phase lands; a CR is proposed,
  deferred or abandoned) or its dependencies change, send the change in the same session
  (`SendMessage`) to the session whose working directory is `~/git/docx4j-portfolio` (find it in
  `ListAgents` by that directory; its name suffix changes): the entry id, the exact new `status`,
  `depends_on` and `notes` (or the whole new entry), and any CR Status line to re-baseline. It
  applies the change, runs `tasks.py check`, `graph` and `accept`, and commits naming the
  requesting session.
- Only if no such session is running may you edit directly: then run
  `python3 ../docx4j-portfolio/scripts/tasks.py check` (it reports `CHANGED` for each CR whose
  Status line was edited; once the registry entry agrees, `tasks.py accept`, and `tasks.py graph`
  if dependencies changed), and commit only your own hunks through a temporary index built from
  HEAD (`GIT_INDEX_FILE`), never the shared index; apply the edit to a copy of HEAD's file and to
  the working tree separately, never truncating the shared file (it was emptied twice that way).
- Before starting a CR or phase, check `python3 ../docx4j-portfolio/scripts/tasks.py blocked`
  (read-only): it may be waiting on work in another repository.
