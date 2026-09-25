# numPicBullet-word2019-pict.docx

From `plutext/docx4j`, `docx4j-core-tests/src/test/resources/numPicBullet-word2019-pict.docx` at
0d076bdde: Word 2019's save of a document with a picture bullet written as a `w:pict`, which is what
makes it a useful fixture - inline VML in a numbering part. Added whole in CR-004 phase B.

Every XML part of the archive is here except `[Content_Types].xml`, which the model does not
carry. Parts are the archive's bytes and are never edited by hand.
