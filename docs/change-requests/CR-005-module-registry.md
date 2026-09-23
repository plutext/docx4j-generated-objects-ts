# CR-005: A bundler-friendly way to load the mapping modules

**Status:** Proposed 2026-09-24
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
   editor's console lists the names to load declarations lazily).
2. Whether the registry is a JavaScript module with 103 static imports or a JSON manifest the
   facade reads (recommendation: the module; a manifest still needs a computed import).
