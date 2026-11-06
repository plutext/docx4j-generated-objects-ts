# NumberingImplicitNumId.docx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/NumberingImplicitNumId.docx` at
tag `docx4j-17.2.1` (0e8e7633e). Added 2026-09-27, at the core-ts session's report, for one part:

- `word/settings.xml` carries `w:stylePaneFormatFilter` with **sixteen** attributes, and it is
  the only part in this corpus that has the element. `wml.xsd` typed it `CT_ShortHexNumber`, which
  declares only `w:val`, so the other fifteen were dropped: 16 in, 1 out. docx4j CR-078 types it
  `CT_StylePaneFilter`, and since this package's regeneration (CR-019) all sixteen round-trip.

Every XML part of the archive is here except `[Content_Types].xml`, which the model does not
carry. Parts are the archive's bytes and are never edited by hand.
