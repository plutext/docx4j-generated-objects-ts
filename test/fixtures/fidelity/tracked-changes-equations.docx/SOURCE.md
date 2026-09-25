# tracked-changes-equations.docx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/tracked-changes-equations.docx` at 49e66dd85:
Word 365's save of a document with tracked changes, comments and OMML equations.

- `word/document.xml` - tracked changes carrying `w16du:dateUtc` (docx4j CR-023) and equations.
- `word/numbering.xml` - `w16cid:durableId` and `restartNumberingAfterBreak` (docx4j CR-023).
- `word/settings.xml` - the `w14:docId` / `w15:chartTrackingRefBased` order Word writes, which the
  schema declares the other way round. A known difference, recorded in `test/fidelity.mjs`.
- `word/people.xml` - a `w15` people part.

Phase B extracted **every** XML part of this document, not only those listed above; the listed
ones are what each was chosen for. Parts are the archive's bytes, never edited by hand, and
`[Content_Types].xml` is excluded (the model does not carry it).
