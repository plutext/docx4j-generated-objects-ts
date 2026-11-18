# Word's reading of a child repeated in a note, a comment and a header's text box (CR-022 D7)

The portfolio session's round `notes-comments-repeats`, the one docx4j CR-079's addendum D10 (a) calls for (registry
`portfolio/notes-comments-repeat-word-check`). Built by the portfolio's `scripts/probes/ncrepeat_build.py` (SHA-256 `56c70fee…`,
untracked there when copied) and run through the Word bridge on 10 November 2026: Word 16.0.20326.20232, bridge `ed94a27`, the
`save` verb to docx, txt and pdf, each file twice. The two runs of each file agree. The `-a` run is copied here.

Every package holds a footnote, an endnote, a comment and a header whose one paragraph holds a VML text box. Each holds plain words
but the container under test, which holds one of the shapes this package's header rounds measured (`test/fixtures/header-footer-probe/`),
as it is:

| Cases | Container | Shapes |
|---|---|---|
| `nr-fn-*` | the footnote, `word/footnotes.xml` | `11` `w:sdtContent`, `12` `w:txbxContent`, `20` `w:pPr` values, `21` `w:pPr` merge, `23` `w:rPr`, `60` `w:tblPr`, `70` `w:ind` and `73` `w:rFonts` twice in one parent |
| `nr-en-*` | the endnote, `word/endnotes.xml` | the same eight |
| `nr-cm-*` | the comment, `word/comments.xml` | the same eight |
| `nr-tb-*` | the header's text box, `word/header1.xml` | the same but `12` |
| `nr-00-control` | every container plain | |

The bridge's own control, `nr-0-mu-control`, isn't copied.

Each case is a directory:
- `source/word/...` holds the container's part as built, and for the control all four.
- `word365/word/...` holds the same parts of Word's save (`<case>.word.docx` in `bridge/portfolio/<case>-a/`).
- `word365.txt` is the text Word saved (`<case>.txt`).

`cases.json` gives each case's story, its parts, the header case its shape comes from, and what the main part and a header read
before the round. `MANIFEST.sha256` gives the SHA-256 of each probe, each `-a` save and each text, by its path in `bridge/portfolio/`,
so `sha256sum -c` run there checks all 96.

The parts were extracted from the archives with nothing changed.

What they show: in a footnote, an endnote and a comment every shape reads as in the main part, two `w:pPr` merging. A paragraph in a
header's text box reads as the header: two `w:pPr` keep the first copy (`nr-tb-20`, `nr-tb-21`). Two `w:txbxContent` in a VML box in
a note or a comment made Word repair the file, each job reporting unreadable content, and its save loses the box's words and the
container's later lines. Nothing here measures a DrawingML text box, or a VML box with one `w:txbxContent` in a note or a comment.
