# loadAndSave.docx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/loadAndSave.docx` at 49e66dd85 (the same
file `@docx4j/core-ts` uses as `test/fixtures/loadAndSave.docx`): a Word 365 document with
comments and a chart.

- `word/commentsExtensible.xml` - a `w16cex` root, typed by docx4j CR-018.
- `word/charts/chart1.xml` - carries `<c16r3:dispNaAsBlank val="1"/>`, whose `val` was dropped on a
  round trip (the element came back bare, flipping the option) until docx4j `c363a969f`, released
  in 0.1.6.

Phase B extracted **every** XML part of this document, not only those listed above; the listed
ones are what each was chosen for. Parts are the archive's bytes, never edited by hand, and
`[Content_Types].xml` is excluded (the model does not carry it).
