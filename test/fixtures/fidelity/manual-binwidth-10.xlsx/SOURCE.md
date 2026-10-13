# manual-binwidth-10.xlsx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/charts/chartex/manual-binwidth-10.xlsx` at
`28bdf831c`: Excel 365's save of a histogram over synthetic data with a manual bin width 10 (`cx:binSize val="10"`),
the oracle of docx4j CR-040 (`ChartExBinningTest`). Added whole 2026-10-14 with the
regeneration that binds the two elements as Excel writes them; before it, this package kept
`<cx:binning intervalClosed="r"/>` and lost the bins, and the fidelity test reports exactly that.

Every XML part of the archive is here except `[Content_Types].xml`, which the model does not
carry. Parts are the archive's bytes and are never edited by hand.
