# Word 365 run property probes (CR-013)

`<probe>.docx/word/document.xml` is the main document part of a probe, and `<probe>-word365.docx/word/document.xml`
the same part of Word 365's save of it: rounds 1 to 6 of Jason Harrop's probes of 24 October 2026, 28 probes and 28
saves. The archives are on the share, `Office 2016/onoff-fun/`, with Word's PDFs and screenshots and the share's
`README.txt`, which says what each marker tests. Rounds 1 to 4 were built by the docx4j-portfolio session, rounds 5
and 6 by docx4j-25 (`build5.py`, `build6.py` there), and all were opened and saved in Word 365 by Jason the same day.

docx4j holds the same archives, byte for byte (compared 2026-10-24), in
`docx4j-core-tests/src/test/resources/runprops/word365/` (`plutext/docx4j`, 87e6ace56 for rounds 1 to 4 and
1d1b824a9 for rounds 5 and 6), read by its `RunPropertiesWord365Test`. Round 7 (`colour-pattern`,
`colour-background-*`) tests what a save writes, and waits for CR-013's phase C.

Only `word/document.xml` is here: the test reads direct formatting, so the styles part is not needed. Parts are the
archives' bytes, extracted with `unzip -p`, and are never edited by hand.
