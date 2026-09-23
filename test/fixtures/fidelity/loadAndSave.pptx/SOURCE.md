# loadAndSave.pptx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/loadAndSave.pptx` at 49e66dd85 (the same
file `@docx4j/core-ts` uses): a PowerPoint 365 deck.

- `ppt/slides/slide2.xml` - a placeholder holding PowerPoint's usual `mc:AlternateContent`, whose
  `mc:Choice Requires="a14"` carries an `a14:m` equation and whose `mc:Fallback` is a picture.
  Checked twice: as Office wrote it, and with the `a14` branch taken, which is what a consumer that
  resolves markup compatibility unmarshals. The second is the case 0.1.5 regressed on - `a14:m`
  typed, and OMML's `CT_R` rejecting the `a:rPr` DrawingML text carries - and the first is not,
  which is why both are checked.

Extracted as part XML, not the archive (CR-004 section 3). Never edited by hand.
