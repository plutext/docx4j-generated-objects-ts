# Change requests

Numbered proposals for changes to the hand-written part of this package (the facade in `src/`,
the package layout, the tests). Changes to the generated `modules/` are proposed in the compiler
repository (`plutext/jsonix-schema-compiler`, `docs/change-requests/`), changes to the runtime in
`plutext/jsonix`.

| CR | Title | Status |
|---|---|---|
| [CR-001](CR-001-namespace-prefixes.md) | The facade ships docx4j's namespace prefix table as the context default (`NAMESPACE_PREFIXES`; unused root declarations stripped until the runtime declares on demand) | Implemented 2026-09-10 |
