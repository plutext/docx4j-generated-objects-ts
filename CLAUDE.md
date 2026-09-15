# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

`@docx4j/generated-objects-ts`: Office Open XML as typed JavaScript. `modules/` holds 94 Jsonix mappings
(UMD `.js`, ES module `.mjs`) with TypeScript declarations (`.d.ts`, `.d.mts`), **generated** by
[jsonix-schema-compiler](https://github.com/plutext/jsonix-schema-compiler) from
[docx4j](https://github.com/plutext/docx4j)'s `xsd/ROOT.xsd`. `src/` holds the only hand-written
code: a facade with docx4j's names (`index.mts`), `helpers/wml.mts` (per-type docx4j decisions,
compiler CR-007) and `builders/wml.mts` (CR-002: `wml` fragments, text sugar, the run mapping,
`textOf`, traversal; imports the helpers, never the reverse). The runtime is
[`@docx4j/jsonix`](https://github.com/plutext/jsonix) 3.2.1+. This is the counterpart of docx4j's
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

`prepublishOnly` runs `typecheck` then `test`. CI (`.github/workflows/test.yml`) runs both on Node
18, 20 and 22, after `npm ci`; `npm update` then a lockfile commit picks up newer dependency versions.
Releases publish to npm from `.github/workflows/push-to-npm.yml` on a GitHub release (trusted
publishing, tag = `package.json` version); see `RELEASING.md`.

There are three tests and no test framework:

- `test/smoke.mjs` (runtime, plain `node:assert`) unmarshals `test/fixtures/document.xml` through the facade and checks `TYPE_NAME`,
  `PARENT`, `deepCopy` (children re-linked, the copy's own `PARENT` unset), a marshal round trip, the
  `v:line` attribute order `id style from to`, that mappings load via `require()` (directly and
  by package self-reference) and via `import()` through `./modules/*`, and the package round trip
  (`unmarshalPackage` types the known parts, `marshalPackage` leaves its input untouched).
  CR-002's block checks `builders/wml`: every wrapper row, the content control inference (nested,
  forced), the tagged form, the sugar, the run mapping round trip, `textOf`, `walk` / `find` /
  `linkParents`.
- `test/readme-examples.ts` (compile-only, via `typecheck`) holds the README snippets against
  minimal Office JS stubs. Change a README example and this file together.
- `test/nodenext/consumer.mts` (compile-only, via `npm test` after the build) imports every public
  path by the package's own name under `module`/`moduleResolution: nodenext`, as a Node ES module
  consumer does. The repository's tsconfigs use `bundler`, which accepts extensionless relative
  imports in the emitted declarations; `nodenext` rejects them, and with `skipLibCheck` they
  silently become `any` (0.1.0 shipped that in `helpers/wml` and `builders/wml`). Import `modules/`
  from `src/` with the `.mjs` extension. Its `@ts-expect-error` lines go unused, and fail, if those types
  resolve to `any`.

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
  those options. The facade is asynchronous because modules load via dynamic `import()`.
  `unmarshalPackage` / `marshalPackage` handle flat OPC packages (Office JS `getOoxml()`): parts are
  `xsd:any processContents="skip"`, hence DOM by the schema, and are converted to typed elements
  where the model knows the root element (and back to DOM on marshal, without modifying the input).
  `MODULE_NAMES` in `src/index.mts` is maintained by hand when modules appear or disappear
  (`generate.sh` does not write it); the smoke fails on a stale list (a mapping references a
  missing dependency).
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
keeping what `mc:Ignorable` names. Marshal through the facade, not `createMarshaller()` directly,
to get that output.

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

- When a CR's status changes (a phase lands; a CR is proposed, deferred or abandoned) or its
  dependencies change, update the matching entry in `tasks.yaml` in the same session (`status`,
  `depends_on`; add an entry for a new CR or phase).
- Then run `python3 ../docx4j-portfolio/scripts/tasks.py check`. It reports `CHANGED` for each CR
  whose Status line was edited; once the registry entry agrees, run `tasks.py accept` (and
  `tasks.py graph` if dependencies changed).
- Before starting a CR or phase, check `python3 ../docx4j-portfolio/scripts/tasks.py blocked`: it
  may be waiting on work in another repository.
