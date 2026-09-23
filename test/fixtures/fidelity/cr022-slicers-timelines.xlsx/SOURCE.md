# cr022-slicers-timelines.xlsx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/cr022-slicers-timelines.xlsx` at 49e66dd85:
Excel 365's save of a pivot table with a slicer and a timeline.

- `xl/slicerCaches/slicerCache1.xml` - an `x14` `slicerCacheDefinition` root (docx4j CR-022), with
  `mc:Ignorable="x xr10"`: `x` is bound to the SpreadsheetML main namespace, which this package
  writes as the default. A known difference, recorded in `test/fidelity.mjs`.
- `xl/timelineCaches/timelineCache1.xml` - an `x15` `timelineCacheDefinition` root (docx4j CR-022).

Extracted as part XML, not the archive (CR-004 section 3). Never edited by hand.
