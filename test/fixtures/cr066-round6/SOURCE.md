# docx4j CR-066's Word probe, round 6: the integer attributes round 5 left unmeasured or inferred (CR-016)

These files are round 6 of the shared Word probe, which docx4j-25 built at this package's request (docx4j CR-066
section 3.3). Jason Harrop opened each one in Word 365 on 26 October 2026: every file opened with no repair offered,
and he saved each one.

What is here:
- `cr066-r6-<group>.docx/word/document.xml` is the probe part, and `cr066-r6-<group>-word365.docx/...` the same part of
  Word's save, for groups A, E, F, P, R, S and T;
- `cr066-r6-D1-<n>` and `D2-<n>` are one value each of `w:defaultTabStop` and `w:hyphenationZone`, so only their
  `word/settings.xml` is here;
- `cases.txt` is the case list: id, group, attribute, place, value, and what a signed or unsigned 16-bit wrap would give;
- `families.txt` is the family list.

The groups:
- **A:** `w:ind` `firstLine`, `hanging`, `right`, `start` and `end`, `w:spacing/@w:after`, and `w:spacing/@w:line` under
  `atLeast`;
- **E:** the forms with no digit (`.`, `+`, `-`, `+.`) in `w:spacing/@w:before` and `w:ind/@w:left`, under a style whose
  `before` and `left` are 100 (CR-015);
- **F:** `w:framePr` `w`, `h`, `hSpace`, `vSpace`, `x` and `y`;
- **P:** `w:tblpPr`'s four distances from text, and `tblpX` and `tblpY`;
- **R:** `w:fitText`, `w:szCs`, `w:kern`, and a run border's `sz` and `space`;
- **S:** `w:pgMar` `left`, `right`, `header`, `footer`, `gutter` and `bottom`, `w:cols` and `w:col`, `w:lnNumType/@w:distance`,
  and a page border;
- **T:** `w:trHeight` under `exact` and `auto`, `w:gridCol`, table and cell borders, `w:tblCellSpacing`, `w:tblCellMar`,
  `w:tcMar`, `w:wBefore` and `w:wAfter`.

Group N (a numbering level's `w:legacy`) is left out: Word 365 dropped the element and renumbered the lists, so its save
measures nothing. The case layout is round 5's (`../cr066-probe/SOURCE.md`).

The archives are on the share, `Office 2016/__tmp/cr066/round6/` (`groups/`, `singles/`), with the probe's `README.txt`.
docx4j records the result in `docs/developer/change-requests/CR-066-integer-attributes-within-words-range.md` section
3.4.

The parts are the archives' bytes, extracted with `unzip -p`, and are never edited by hand.
