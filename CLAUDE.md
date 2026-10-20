# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

`@docx4j/generated-objects-ts`: Office Open XML as typed JavaScript. `modules/` holds 103 Jsonix mappings
(UMD `.js`, ES module `.mjs`) with TypeScript declarations (`.d.ts`, `.d.mts`), **generated** by
[jsonix-schema-compiler](https://github.com/plutext/jsonix-schema-compiler) from
[docx4j](https://github.com/plutext/docx4j)'s `xsd/ROOT.xsd`. `src/` holds the only hand-written
code: a facade with docx4j's names (`index.mts`), the module registry (`modules.mts`), `wml-globals.mts` (CR-009: one more
mapping declaring globally the 61 WordprocessingML elements docx4j declares as roots, so kept `mc` branches type their
`w:drawing`/`w:pict`; `getContext` and `modulesFor` add it), `helpers/wml.mts` (per-type docx4j decisions,
compiler CR-007) and `builders/wml.mts` (CR-002: `wml` fragments, text sugar, the run mapping,
`textOf`, `runItemsOf`, content controls, rows and cells, inline pictures, the `w:rPrChange` element
list, `walkAll` and `mcBranchOf` (CR-003 phase A and section 3.8), traversal; imports the helpers, never the reverse). The runtime is
[`@docx4j/jsonix`](https://github.com/plutext/jsonix) 3.5.0+ (on/off booleans, jsonix-CR-007; base64Binary in strict mode, 3.4.2; the unmarshal-side `onElement` CR-011 records through, jsonix-CR-008, 3.5.0). This is the counterpart of docx4j's
`docx4j-generated-objects` module and is usable on its own (Office JS add-ins); the engine layer
(OPC packaging, parts, style/numbering resolution, the counterpart of `docx4j-core`) is the separate
repository `plutext/docx4j-core-ts`, package `@docx4j/core-ts`, which depends on this one. Tree-only helpers belong here;
anything needing parts or relationships belongs there.

## Commands

```
npm ci              # typescript and the @docx4j/jsonix runtime, as locked in package-lock.json (committed)
npm run build       # tsc -p tsconfig.build.json: src/ -> dist/ (index, helpers/wml, builders/wml: .mjs and .d.mts)
npm run typecheck   # tsc --strict, noEmit, over modules/*.d.ts, modules/*.d.mts, src/, test/*.ts
npm test            # build, then the nodenext consumer check, then node test/smoke.mjs
node test/smoke.mjs # the single runtime test, when dist/ is already built
npm pack --dry-run  # ships dist/, modules/, LICENSE, NOTICE, README.md, package.json only
```

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

There are four tests and no test framework:

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
  marshal whole. CR-011's block checks that a prefix only the source bound is declared from the
  source: on the root, on `mc:AlternateContent`, a second prefix for a known namespace, after a
  move between trees, a `wml()` fragment, `deepCopy`, a flat OPC part, a paragraph marshalled
  alone; the superseded, tree-used and two-ways cases; and that a caller's `onElement` is composed.
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
  namespace declarations, every prefix an `mc:Ignorable` or an `mc:Choice`'s `Requires` names in a
  marshalled part is also checked to resolve where it is written (CR-001 section 3.2).
  Run against 0.1.5's `modules/` the test reports all five losses that release carried.
- `test/nodenext/consumer.mts` (compile-only, via `npm test` after the build) imports every public
  path by the package's own name under `module`/`moduleResolution: nodenext`, as a Node ES module
  consumer does. The repository's tsconfigs use `bundler`, which accepts extensionless relative
  imports in the emitted declarations; `nodenext` rejects them, and with `skipLibCheck` they
  silently become `any` (0.1.0 shipped that in `helpers/wml` and `builders/wml`). Import `modules/`
  from `src/` with the `.mjs` extension. Its `@ts-expect-error` lines go unused, and fail, if those types
  resolve to `any`.

Fixtures are extracted part XML, never archives (no zip dependency) and never edited by hand; they
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
`onElement`. Unmarshal through the facade too, or nothing is recorded. CR-008 (implemented 2026-10-04) has them first append to a
part root's `mc:Ignorable` the Office extension namespaces its tree uses
(`IGNORABLE_EXTENSION_NAMESPACES`, the twenty Office lists, measured on the fidelity corpus, which
must not move). Marshal through the facade, not `createMarshaller()` directly, to get that output.

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
