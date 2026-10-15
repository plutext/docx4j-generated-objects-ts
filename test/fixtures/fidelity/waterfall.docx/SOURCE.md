# waterfall.docx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/charts/waterfall.docx` at `f9708f5b3`:
Word's save of a waterfall chart, a chartex (`cx:chartSpace`) part whose `word/charts/chartEx1.xml`
carries `<cx:externalData r:id="rId1" cx:autoUpdate="0"/>`, the oracle of docx4j CR-040 phase 2
(`ChartExExternalDataTest`). Added whole 2026-10-15 with the regeneration that binds the
attribute qualified, as Office writes it; before it, this package read the attribute past and
wrote none, and the fidelity test reports exactly that.

Every XML part of the archive is here except `[Content_Types].xml`, which the model does not
carry; the media, the embedded workbook and the thumbnail are binary and stay out. Parts are the
archive's bytes and are never edited by hand.
