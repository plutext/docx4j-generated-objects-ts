# cr022-slicers-timelines.xlsx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/cr022-slicers-timelines.xlsx` at 49e66dd85:
Excel 365's save of a pivot table with a slicer and a timeline.

- `xl/slicerCaches/slicerCache1.xml` - an `x14` `slicerCacheDefinition` root (docx4j CR-022), with
  `mc:Ignorable="x xr10"`: `x` is bound to the SpreadsheetML main namespace, which this package
  writes as the default. A known difference, recorded in `test/fidelity.mjs`.
- `xl/timelineCaches/timelineCache1.xml` - an `x15` `timelineCacheDefinition` root (docx4j CR-022).
- `xl/drawings/drawing1.xml` - an `a:graphicData` framing an `sle:slicer`, inside the `a14`
  `mc:Choice` every consumer takes. As Office wrote it the part round-trips (the
  `mc:AlternateContent` is kept whole); with the branch taken it throws, because
  `CT_GraphicalObjectData`'s wildcard is `processContents="strict"` and a graphic the model does
  not bind cannot stay DOM. Found by core-ts CR-004 phase A (2026-09-24), recorded in
  `test/fidelity.mjs` against docx4j.

Phase B extracted **every** XML part of this document, not only those listed above; the listed
ones are what each was chosen for. Parts are the archive's bytes, never edited by hand, and
`[Content_Types].xml` is excluded (the model does not carry it).
