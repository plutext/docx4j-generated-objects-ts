# Word's reading of a child given more than once (CR-022)

Copied from docx4j's `docx4j-core-tests/src/test/resources/repeated/` at the tag `docx4j-17.3.6` (`4afcf4be8`), the files docx4j's
`RepeatedChildrenTest` reads, by name. Each case is a directory:

- `source/word/...` holds the source's parts this test reads: `document.xml`, and where the source has them, its
  `_rels/document.xml.rels` and its header and footer parts. Not `[Content_Types].xml`, `_rels/.rels` or `numbering.xml`.
- `word365/word/...` holds the same parts of Word's save of it (`<case>-word365.docx`): `document.xml`, and the relationships,
  headers and footers where the source has any.
- `word365.txt` is the text Word saved from the source (`<case>-word365.txt`), as it is.

The parts were extracted from the archives with nothing changed. `MANIFEST.sha1` gives the SHA-1 of each of the 303 files they come
from, so `sha1sum -c` run in docx4j's `repeated/` checks them.

Word 16.0.20326.20232, through the Word bridge, each round as docx4j's CR-079 records it (WD REQ-322).

| Round | Cases | What it holds |
|---|---|---|
| `rc-*` | 16 | Two copies: `w:body`, block `w:sdtContent`, `w:txbxContent` in a VML box, `w:pPr`, `w:rPr`, `w:jc`, `w:b`, `w:sz`, `w:sectPr`, `w:tblPr`, and three controls |
| `rt-*` | 20 | Three copies of the same, four of `w:txbxContent`, bodies with and without a `w:sectPr`, and `rt-50`/`rt-51`, an `mc:AlternateContent` root holding three `w:document` |
| `sm-*` | 40 | `sectpr-merge`: one difference a file, among `w:sectPr`'s attributes and children, `w:tblPr`, `w:ind`, `w:spacing`, `w:rFonts`, `w:vanish` and the header references. `sm-52` has an `mc:AlternateContent` root |
| `ft-*` | 12 | `footer-refs`: header and footer references across copies and joined bodies |
| `c79-*` | 7 | docx4j CR-079's phase 0: `w:sdtContent` at run, row and cell level, `wps:txbx`, a `w:pPr`'s `w:rPr`, `w:numPr` and `w:pPrChange` |
| `op-*` | 6 | docx4j-python's `one-parent-word-check`: `w:pgSz`, `w:pgMar`, `w:cols`, `w:numPr` and a paragraph mark's `w:rPr` twice in one parent |

Not copied: docx4j's `nd-*` (11 cases), a `w:body` or `w:document` nested in another. Five have no save, since Word refuses them.
A `w:body` in a `w:body` is core-ts's to read (its CR-028 D12), as are `rt-50`, `rt-51` and `sm-52`, whose root the facade
leaves to core-ts's root pass.

Every case is a main document part. Nothing here says what Word does with a repeat in a header, a note, a style or a numbering
level.
