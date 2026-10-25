# docx4j CR-065's Word probe: decimals and odd forms in integer attributes (CR-015)

These files are the shared Word probe of docx4j CR-065 phase 0, which docx4j-25 built for docx4j, core-ts and
docx4j-python. Jason Harrop opened each one in Word 365 on 26 October 2026, in rounds 2 to 4, and saved each one that
Word opened.

What is here:
- `cr065-<family>.docx/word/document.xml` is the probe part;
- `cr065-<family>-word365.docx/...` is the same part of Word's save of it;
- for family I, `styles.xml`, `settings.xml` and the header are here too. The header is `header.xml` in the probe and
  `header1.xml` in Word's save;
- `cases.txt` is the case list: family, id, attribute and value.

The families:
- **A1 to A7** split family A's text forms:
  - A1 (plain decimals), A2a to A2d (`240.`, `.5`, `+240`, `+240.5`) and A6 (`""`) opened clean;
  - A2e (`" 240.5 "`), A2f (`" 240 "`), A3 (`240,5`), A4 (`0x10`), A5 (`1.5e2`, `2.5e0`, `1e3`) and A7 (`abc`) were
    each refused outright, with no repair offered, so they have no save.
- **B to I** opened clean: signed twips, docx4j's rounding list, borders, tab stops, run properties, tables, sections,
  and a style, the settings and a header.

The archives are on the share, `Office 2016/__tmp/cr065/` and its `families/`, with the probe's `README.txt`. docx4j
records the rounds in `docs/developer/change-requests/CR-065-integer-attributes-read-as-word-reads-them.md` section 3
(`plutext/docx4j` `VERSION_17_3_5`: 02b1cae86 for round 2, e1968a1ac for round 3 and aacaa4c45 for round 4).

The parts are the archives' bytes, extracted with `unzip -p`, and are never edited by hand.
