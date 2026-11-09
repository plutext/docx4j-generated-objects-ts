# Word's reading of a child repeated in a header or a footer (CR-022 D7)

This package's two Word rounds, built in this repository's session and run through the Word bridge by the portfolio session on
10 November 2026: Word 16.0.20326.20232, bridge `da85101`, the `save` verb to docx, txt and pdf, each file twice. Every job opened
silently, and each file's two runs saved the same parts but for `w14:paraId`, `w14:textId` and the rsids, and the same text. The
`-a` run is copied here.

| Round | Cases | What it holds |
|---|---|---|
| `header-repeats` (`hr-*`) | 8 | A control, and seven measured main-part cases (`rc-11`, `rc-12`, `rc-21`, `rc-23`, `sm-60`, `sm-70`, `sm-73`) with their body content in `word/header1.xml` as it is |
| `header-footer-repeats` (`hf-*`) | 20 | A control with a header and a footer; a header's two `w:pPr` as values, reversed, three copies, three values, an empty first copy, no `w:jc` and a mark's `w:rPr` (`hf-10` to `hf-16`); three copies of `w:rPr`, `w:sdtContent` and a VML `w:txbxContent` in a header (`hf-17` to `hf-19`); the same shapes in a footer (`hf-20` to `hf-25`); and main-part twins of the new markup (`hf-51`, `hf-54`, `hf-55`) |

In each source the main part is "BODY WORDS", and its `w:sectPr` refers to the header or the footer, or the twin's markup is the body.

Each case is a directory:
- `source/word/...` holds the probe's `document.xml`, `_rels/document.xml.rels` and header or footer, as built.
- `word365/word/...` holds the same parts of Word's save (`<case>.word.docx` in `bridge/portfolio/<case>-a/`).
- `word365.txt` is the text Word saved (`<case>.txt`).

`cases.json` gives each case's story, its source and what was known before the round. `MANIFEST.sha256` gives the SHA-256 of each
probe, each `-a` save and each text, by its path in `bridge/portfolio/` (the probe as the bridge copied it into `<case>-a/`, the same
bytes as built), so `sha256sum -c` run there checks all 84.

The parts were extracted from the archives with nothing changed.

What they show: in a header and a footer every shape reads as in the main part but a paragraph's repeated `w:pPr`. There, the
first copy's paragraph properties stand and each later copy is ignored (11 of 11), an empty first copy included, where the
main-part twins merge. In `hf-16` a later copy's paragraph-mark `w:rPr` lands on the paragraph's run. Nothing here measures a
note, a comment, a header's text box's paragraphs, or a paragraph in a header's table.
