# tracked-changes-equations.docx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/tracked-changes-equations.docx` at 49e66dd85:
Word 365's save of a document with tracked changes, comments and OMML equations.

- `word/document.xml` - tracked changes carrying `w16du:dateUtc` (docx4j CR-023) and equations.
- `word/numbering.xml` - `w16cid:durableId` and `restartNumberingAfterBreak` (docx4j CR-023).
- `word/settings.xml` - the `w14:docId` / `w15:chartTrackingRefBased` order Word writes, which the
  schema declares the other way round. A known difference, recorded in `test/fidelity.mjs`.
- `word/people.xml` - a `w15` people part.

Extracted as part XML, not the archive (CR-004 section 3). Never edited by hand.
