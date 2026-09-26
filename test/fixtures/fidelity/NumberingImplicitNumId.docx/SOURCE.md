# NumberingImplicitNumId.docx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/NumberingImplicitNumId.docx` at
tag `docx4j-17.2.1` (0e8e7633e). Added 2026-09-27, at the core-ts session's report, for one part:

- `word/settings.xml` carries `w:stylePaneFormatFilter` with **sixteen** attributes, where
  every other document in this corpus writes the element with `w:val` alone. `wml.xsd` types it
  `CT_ShortHexNumber`, which declares only `w:val`, so the other fifteen are dropped: 16 in, 1
  out. Recorded in `test/fidelity.mjs` against docx4j.

Every XML part of the archive is here except `[Content_Types].xml`, which the model does not
carry. Parts are the archive's bytes and are never edited by hand.
