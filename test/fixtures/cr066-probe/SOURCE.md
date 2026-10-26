# docx4j CR-066's Word probe, round 5: integers at and beyond each attribute's storage width (CR-016)

These files are round 5 of the shared Word probe, docx4j CR-066 phase 0, which docx4j-25 built for docx4j, core-ts and
docx4j-python. Jason Harrop opened each one in Word 365 on 26 October 2026; all six opened with no repair offered, and
he saved each one.

What is here:
- `cr066-<group>.docx/word/document.xml` is the probe part;
- `cr066-<group>-word365.docx/word/document.xml` is the same part of Word's save of it;
- `cases.txt` is the case list: id, group, attribute, value, and what a 16-bit wrap or a clamp would give;
- `families.txt` is the family list.

The groups:
- **P:** `w:spacing/@w:before`, `w:ind/@w:left`, `w:spacing/@w:line` under `auto` and `exact`, and `w:tab/@w:pos`;
- **B:** a border's `@w:sz` and `@w:space`;
- **R:** `w:sz`, `w:position`, the run's `w:spacing` and text scale `w:w`;
- **T:** `w:tblW`, `w:tcW` and `w:tblInd` under `dxa` and `pct`, and `w:trHeight`;
- **S:** `w:pgSz` and `w:pgMar/@w:top`, and `w:pgNumType/@w:start`;
- **N:** negatives in attributes the schema calls unsigned.

Each case paragraph is labelled with its id. The value is in the label paragraph (P, B), in the run after "then:" (R),
in the table after the label (T), or in the section the paragraph after the label ends (S).

The archives are on the share, `Office 2016/__tmp/cr066/groups/`, with the probe's `README.txt` and the family and single
files, which weren't needed. docx4j records the result in
`docs/developer/change-requests/CR-066-integer-attributes-within-words-range.md` section 3.1 (`plutext/docx4j`
`VERSION_17_3_5`, 58f193fb5).

The parts are the archives' bytes, extracted with `unzip -p`, and are never edited by hand.
