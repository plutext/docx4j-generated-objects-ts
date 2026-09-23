# loadAndSave.xlsx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/loadAndSave.xlsx` at 49e66dd85 (the same
file `@docx4j/core-ts` uses): an Excel 365 workbook with a comment, a chart and a drawing.

- `xl/drawings/drawing1.xml` - an `mc:AlternateContent` whose `a14` branch holds an equation;
  checked as written and with that branch taken, as for the pptx slide above.
- `xl/charts/chart1.xml` - a second `c16r3:dispNaAsBlank`, from Excel rather than Word.

Extracted as part XML, not the archive (CR-004 section 3). Never edited by hand.
