# cr022-data-model.xlsx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/cr022-data-model.xlsx` at 0d076bdde:
Excel 365's save of a Power Query data model; its customXml/item1.xml is a UTF-16 DataMashup blob that nobody models, and its pivot caches and connections carry the xr and xr16 uid attributes. Added whole in CR-004 phase B.

Every XML part of the archive is here except `[Content_Types].xml`, which the model does not
carry. Parts are the archive's bytes and are never edited by hand.
