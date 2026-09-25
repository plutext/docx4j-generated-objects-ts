# cr022-checkbox.xlsx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/cr022-checkbox.xlsx` at 49e66dd85:
Excel 365's save of a worksheet with a form-control checkbox linked to a cell.

Parts here, and what each demonstrates:

- `xl/workbook.xml` - `xr2:uid` on `workbookView` (a known difference, see `test/fidelity.mjs`).
- `xl/worksheets/sheet1.xml` - the `x14ac` attributes docx4j CR-022 typed, and the control.
- `xl/styles.xml` - `mc:Ignorable` naming `x16r2`, the prefix `7a74fb0` added to the table.
- `xl/drawings/drawing1.xml` - an `xdr:wsDr` whose only child is an `mc:AlternateContent`. This is
  the part that emptied silently (59 elements in, one out) until docx4j CR-024; it is the reason
  this test exists.
- `xl/ctrlProps/ctrlProp1.xml` - an `x14` `formControlPr` root, typed by docx4j CR-022.
- `xl/drawings/vmlDrawing1.vml` - the legacy drawing behind the control, whose root is `<xml>` in no
  namespace. Every release up to 0.1.6 threw on it; docx4j CR-026 fixed the binding.

Phase B extracted **every** XML part of this document, not only those listed above; the listed
ones are what each was chosen for. Parts are the archive's bytes, never edited by hand, and
`[Content_Types].xml` is excluded (the model does not carry it).
