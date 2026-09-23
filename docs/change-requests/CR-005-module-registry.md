# CR-005: A bundler-friendly way to load the mapping modules

**Status:** Proposed 2026-09-24 (reviewed here 2026-09-24, section 5: accepted in shape, with the registry hand-written in `src/`)
**Depends on:** nothing in this package
**Requested by:** `plutext/docx4j-ts-editor` ED-002 section 10.3 item 1 (the browser editor,
2026-09-23), which had to teach its bundler where the modules are before the first open worked
**Counterpart:** none in docx4j (JAXB contexts are built from class names on the classpath)

## 1. Summary

`getContext()` loads the 103 mapping modules with one computed dynamic import:

```ts
Promise.all(MODULE_NAMES.map((name) => import(`../modules/${name}.mjs`)))
```

Node resolves that at run time and is content. A bundler is not: the specifier is a template
string, so the bundler cannot see which files it names, leaves the `import()` as a runtime URL
relative to the bundle, and the browser asks for `modules/org_docx4j_wml.mjs` beside a bundle
that has no such neighbour. The editor saw 404s on its first open and fixed it on its side
(`vite.config.ts`, `dynamicImportVarsOptions.include` pointing into this package's `dist`), which
makes Vite's dynamic-import-vars pass expand the template into a glob over `../modules/*.mjs`.
That works, at a price: the glob also matches the `.el.mjs`, `.factory.mjs` and `.js` siblings,
so the editor's build carries some 575 chunks of which 103 ever load, and every other consumer
(the add-in guide's esbuild command, a webpack user, a Rollup library build) has to find the
same trick or serve the modules beside the bundle.

The package should not need the trick. Two ways, not exclusive:

1. **A static registry.** A generated `modules/index.mjs` (or `src/modules.mts`) that imports
   every mapping module by a literal specifier and exports them as one array or map. Bundlers
   see 103 literal imports and chunk them as they like; `getContext()` imports the registry.
   The cost is that the registry pulls every module into the graph even for a consumer that
   wants one, which is what today's `Promise.all` does anyway (it loads all 103 on first use).
   `generate.sh` writes it; `MODULE_NAMES` stays for callers that want the names.
2. **A `modules` option on `getContext`.** `getContext({ modules })` takes the loaded modules
   (or a loader function) from the caller, so an add-in that wants a smaller context (wml and
   its dependencies only) passes those and nothing else is fetched. The default stays the
   registry of item 1.

Recommendation: both; item 1 fixes every bundler with no consumer change, item 2 is what a
bundle-size-conscious add-in wants, and it costs one option.

## 2. What changes

- `generate.sh` writes `modules/index.mjs` exporting `MODULES` (an array of the mapping module
  namespaces in `MODULE_NAMES` order) and a `.d.mts` for it.
- `getContext()` builds from `MODULES` when no `modules` option is given; the computed import
  goes.
- `Jsonix.ContextOptions` gains `modules?: readonly Record<string, unknown>[] | (() =>
  Promise<readonly Record<string, unknown>[]>)` on this package's facade (the runtime's options
  type is not touched; the facade strips the option before passing the rest on).
- The README's bundling note is rewritten: no bundler configuration needed; the editor deletes
  its `dynamicImportVarsOptions` once it consumes the release.

## 3. Tests

- The existing context tests pass unchanged (the default path).
- A test builds a context from `modules: [wml, relationships, ...]` and marshals a `w:p`.
- A build test: `esbuild --bundle` over a one-line consumer produces a bundle that unmarshals a
  document with no filesystem beside it (run in `test/`, esbuild as a dev dependency; the
  point of the CR is that this passes without configuration).

## 4. Open questions

1. Whether `MODULE_NAMES` stays exported once the registry exists (recommendation: yes; the
   editor's console lists the names to load declarations lazily). **Answered in section 5.2: it
   stays, but derived from the registry rather than kept beside it.**
2. Whether the registry is a JavaScript module with 103 static imports or a JSON manifest the
   facade reads (recommendation: the module; a manifest still needs a computed import).
   **Agreed: the module.** A manifest buys nothing, since reading it still leaves the facade with
   a computed specifier, which is the whole defect.

## 5. Review (this repository, 2026-09-24)

The diagnosis holds against the shipped code: `dist/index.mjs` carries
``import(`../modules/${name}.mjs`)`` verbatim, so a bundler has a template literal and no way to
know what it names. The cost of the editor's glob workaround is as described: `modules/` holds 309
`.mjs` files of which 103 are mappings, the other 206 being the `.el.mjs` and `.factory.mjs`
siblings of compiler CR-010. Both proposals are accepted in shape, with three changes.

### 5.1 The registry is hand-written in `src/`, not generated into `modules/`

Section 2 has `generate.sh` write `modules/index.mjs`. That crosses a repository boundary for no
gain: `generate.sh` and `bindings.xjb` live in `plutext/jsonix-schema-compiler`, `modules/` is
generated output this repository must never edit by hand, and a change there needs a CR in the
compiler repository, a compiler release, and a regeneration - to ship a fix that has nothing to do
with the schemas. The alternative section 1 already offers in parentheses becomes the
recommendation: **a hand-written `src/modules.mts`** with 103 literal `import` statements,
compiled into `dist/` with the rest of the facade.

Nothing is lost by hand-maintaining it, because the list is hand-maintained today: `MODULE_NAMES`
lives in `src/index.mts` precisely because `generate.sh` does not write it (see `CLAUDE.md`), and
the smoke test fails when it goes stale, since a mapping then references a module the context
never loaded. The registry inherits that check.

### 5.2 One list, not two: `MODULE_NAMES` derives from the registry

A registry of 103 literal imports beside a hand-kept array of the same 103 names is two lists that
must agree, and only one of them is watched. The registry is therefore the single source: it maps
name to module, and `MODULE_NAMES` becomes its keys (`Object.keys(MODULES)`, or the array built
from it), keeping the export and the `ModuleName` type that callers and `test/nodenext/consumer.mts`
use. A module added or dropped by a regeneration is then one edit, and the smoke still catches a
miss.

### 5.3 The `modules` option obeys the singleton rule, and says so

`getContext` caches one context and returns it for every later call, so `getContext({ modules })`
after something has already built the default context returns that one and ignores the option, as
it already does for `namespacePrefixes` and the rest. This is a footgun for exactly the caller item
2 is for - the add-in that wants a small context will often be the second caller, not the first.
The option is fine as proposed, but its documentation must carry the same sentence `resetContext`
carries: to build a context with different options, call `resetContext()` first. A test should pin
it (`getContext({ modules: [...] })` after a default build returns the full context; after
`resetContext()`, the small one).

### 5.4 Not blocking

None of this changes the recommendation to do both items. The defect is real for the standalone
consumers this package exists to serve - an Office JS add-in bundled with esbuild meets it on its
first run, as the editor did.
