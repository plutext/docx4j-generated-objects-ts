# Word 365 run property probes (CR-013)

`<probe>.docx/word/document.xml` is the main document part of a probe, and `<probe>-word365.docx/word/document.xml`
the same part of Word 365's save of it: rounds 1 to 7 of Jason Harrop's probes of 24 October 2026, 33 probes and 33
saves. The archives are on the share, `Office 2016/onoff-fun/`, with Word's PDFs and screenshots and the share's
`README.txt`, which says what each marker tests. Rounds 1 to 4 were built by the docx4j-portfolio session, rounds 5
to 7 by docx4j-25 (`build5.py`, `build6.py`, `build7.py` there), and all were opened and saved in Word 365 by Jason
the same day.

`<probe>-docx4j.docx/word/document.xml`, for fifteen probes, is docx4j's own save of the probe, made by docx4j at
aaadd93f9 (CR-061 phases 0 to 2) for round 7's Word check, from the same share. CR-013's phase C writes what these
do, value for value.

docx4j holds the probes and Word's saves, byte for byte (compared 2026-10-24), in
`docx4j-core-tests/src/test/resources/runprops/word365/` (`plutext/docx4j`: 87e6ace56 for rounds 1 to 4, 1d1b824a9
for rounds 5 and 6, 5163c2986 for round 7), read by its `RunPropertiesWord365Test` and `ColourNormalFormTest`.

Only `word/document.xml` is here: the tests read direct formatting, so the styles part is not needed. Parts are the
archives' bytes, extracted with `unzip -p`, and are never edited by hand.
