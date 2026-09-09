# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

`@docx4j/generated-objects-ts`: Office Open XML as typed JavaScript. `modules/` holds 94 Jsonix mappings
(UMD `.js`, ES module `.mjs`) with TypeScript declarations (`.d.ts`, `.d.mts`), **generated** by
[jsonix-schema-compiler](https://github.com/plutext/jsonix-schema-compiler) from
[docx4j](https://github.com/plutext/docx4j)'s `xsd/ROOT.xsd`. `src/` holds the only hand-written
code: a facade with docx4j's names (`index.mts`) and `helpers/wml.mts`. The runtime is
[`@docx4j/jsonix`](https://github.com/plutext/jsonix) 3.2.0+. This is the counterpart of docx4j's
`docx4j-generated-objects` module and is usable on its own (Office JS add-ins); the engine layer
(OPC packaging, parts, style/numbering resolution, the counterpart of `docx4j-core`) is the separate
repository `plutext/docx4j-core-ts`, package `@docx4j/core-ts`, which depends on this one. Tree-only helpers belong here;
anything needing parts or relationships belongs there.

## Commands

```
# Until @docx4j/jsonix 3.2.0 is on npm, install the runtime from a sibling checkout of plutext/jsonix
npm install --no-save typescript@5.6.3 ../jsonix/nodejs/scripts
# (afterwards: npm install)

npm run build       # tsc -p tsconfig.build.json: src/ -> dist/ (index.mjs, index.d.mts, helpers/wml.mjs, .d.mts)
npm run typecheck   # tsc --strict over modules/*.d.ts, modules/*.d.mts, src/, test/ (lib includes dom: the runtime typings need Node/Document)
npm test            # build, then node test/smoke.mjs
npm pack --dry-run  # ships dist/, modules/, LICENSE, NOTICE, README.md, package.json only
```

`test/smoke.mjs` unmarshals `test/fixtures/document.xml` through the facade and checks `TYPE_NAME`,
`PARENT`, `deepCopy` (children re-linked, the copy's own `PARENT` unset), a marshal round trip, the
`v:line` attribute order `id style from to`, and that mappings load via `require()` (directly and
by package self-reference) and via `import()` through `./modules/*`.

## Layout and rules

- `modules/` is generated output plus a reference copy of `bindings.xjb`. **Never edit it by hand.**
  Regenerate from the compiler repository (`OfficeOpenXML/generate.sh ../docx4j/xsd/ROOT.xsd
  ../docx4j-generated-objects-ts`, see `generate.md`); the output is deterministic, so a regeneration with unchanged
  inputs is an empty diff. Commits that regenerate cite the compiler and docx4j commits used.
- `package.json` deliberately has **no `"type": "module"`**: the UMD `.js` mappings must stay
  CommonJS-loadable; ES modules are marked by `.mjs`/`.mts`. Adding it breaks `require()` of every
  mapping (ERR_REQUIRE_ESM).
- Public paths are the `exports` map only: `.` (facade), `./helpers/wml`, `./modules/*`
  (`import` → `.mjs`, `require` → `.js`, `types` → `.d.ts`). Keep them stable.
- The facade builds one `Jsonix.Context` over all modules lazily (`getContext`, first use) with
  `parentPointers: true`; modules depend on each other, so the context is all-or-nothing.
  `unmarshalPackage` / `marshalPackage` handle flat OPC packages (Office JS `getOoxml()`): parts are
  `xsd:any processContents="skip"`, hence DOM by the schema, and are converted to typed elements
  where the model knows the root element (and back to DOM on marshal, without modifying the input).
  README examples are compile-checked in `test/readme-examples.ts` against minimal Office JS stubs.
  `MODULE_NAMES` in `src/index.mts` is maintained by hand when modules appear or disappear; the
  smoke fails on a stale list (a mapping references a missing dependency).
- `src/` and `test/` import from `../modules/...` and `../dist/...`; `dist/` is git-ignored and
  built by `npm run build` (also on `prepublishOnly`).

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
