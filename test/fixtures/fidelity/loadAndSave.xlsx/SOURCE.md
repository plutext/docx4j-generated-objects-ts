# loadAndSave.xlsx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/loadAndSave.xlsx` at 49e66dd85 (the same
file `@docx4j/core-ts` uses): an Excel 365 workbook with a comment, a chart and a drawing.

- `xl/drawings/drawing1.xml` - an `mc:AlternateContent` whose `a14` branch holds an equation;
  checked as written and with that branch taken, as for the pptx slide above.
- `xl/charts/chart1.xml` - a second `c16r3:dispNaAsBlank`, from Excel rather than Word.

Phase B extracted **every** XML part of this document, not only those listed above; the listed
ones are what each was chosen for. Parts are the archive's bytes, never edited by hand, and
`[Content_Types].xml` is excluded (the model does not carry it).
