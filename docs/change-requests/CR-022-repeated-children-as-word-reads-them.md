# CR-022: A repeated child read as Word reads it, through jsonix-CR-016's callback

**Status:** Implemented 2026-11-10 at Jason's route (c) in core-ts CR-028 D2 ("Objects CR-022 (Recommended)", relayed
by core-ts the same day). It reads as docx4j 17.3.6's `RepeatedChildren` does, rule for rule, against the same Word saves (section 6),
but in a header or a footer, where D7 follows Word. Section 4 is wider than the CR as accepted, by what Word was measured on since
(section 4.1). Accepted 2026-11-07, D1 to D5 as recommended (Jason: "CR-022 as recommended") and D6 the same day (Jason: "D6: widen CR-022 to REQ-322"). D7 decided (c) 2026-11-10
(Jason: "D7 (c) as recommended"), after two Word rounds, and built: in a header or a footer a paragraph's repeated `w:pPr` keeps the first copy.
In their text boxes too since 2026-11-13 (Jason: "do The CR-022 text box fix"), after the portfolio's round in notes, comments and a
header's text box (section 3.2). Released in 0.9.0 (2026-11-18, tag `54c9966`; Jason: "release 0.9.0"), after core-ts's go on `7958467`.
**Requested by:** Jason, 2026-11-07 ("ask objects for that CR"), relayed by jsonix-4b; built at his route (c), relayed by core-ts.
**Depends on:** `jsonix/CR-016`, released in `@docx4j/jsonix` 3.13.0 (2026-11-07, tag `a5359c4`). Its `onOverwrittenElement` is the
hook, so the floor here is `^3.13.0`.
**Counterparts:**
- WD REQ-322 (WD 0.101, portfolio `7150868`) is the reading, for a main document part.
- docx4j's CR-079 phases 2 and 3 (`RepeatedChildren`, af50deffe and 8c23c444f, released in 17.3.6) and docx4j-python's CR-039 are the
  other ports'. The comparisons in section 6 are docx4j's `RepeatedChildrenTest`, which ported python's.
- core-ts CR-028 takes this reading for its section 3.2 (D2's route (c), so D5 here is settled) and keeps the root pass and a
  `w:body` in a `w:body` (its D12).
- jsonix stays at the report, because the runtime holds no Word rules: the same split as jsonix-CR-010 and CR-013.

## 1. Summary

A single-valued property given a second element keeps the last. jsonix-CR-016 reports that, but a report leaves the first
`w:body`'s words out of the tree, and Word shows them.

Every facade read now passes `wordOverwritten` as the runtime's `onOverwrittenElement`, which reads the measured pairs as WD REQ-322,
Word and docx4j do:
- **Joined:** `w:body` and `w:sdtContent` at every level, in order, and `w:txbxContent` in a VML box in reverse.
- **Applied in order to one object:** `w:pPr`, `w:rPr`, `w:sectPr` and `w:tblPr`, and the children measured twice in one parent.
- **In a header or a footer,** and in a text box in one, a paragraph's repeated `w:pPr` keeps the first copy whole, as Word reads it
  there (D7).

Anything else stays as the runtime leaves it.

Against Word's saves of docx4j's 101 files, 98 match on every comparison, against 33 with the runtime's reading. The other three
are `mc:AlternateContent` roots, core-ts's root pass. In a header or a footer, 25 of this package's 28 match, two more to the same
effect, and the last but for one run's italic. In notes, comments and a header's text box, the portfolio's 32 files match but the three
Word repairs. No fidelity hash moved.

## 2. What was there

`PropertyInfo.setProperty` overwrites, so the last copy stands, with no report. That matches Word for an element with one attribute.
Elsewhere it loses what the earlier copies stated: 68 of the 101 files differed from Word's save, and in 29 `textOf` didn't give
Word's text, where it now does in all but the 17 section 6 names.

## 3. What Word reads

Word 16.0.20326.20232, each file twice, all silent, the runs equal apart from generated ids. On the share under
`bridge/portfolio/`: `rc-*` two copies, `rt-*` three and four, `sm-*` and `ft-*` one difference a file. Then docx4j CR-079's phase 0
(`c79-*`) and python's `one-parent-word-check` (`op-*`). `rt-50`, `rt-51` and `sm-52`, which are mc roots, aren't this CR's.

| Given more than once | Word |
|---|---|
| `w:body` | Joined in order, the closing `w:sectPr` applied in order to one object, a body with none passed over. One in a `w:pPr` stays |
| `w:sdtContent` | Joined in order, at block level (`rc-11`, `rt-11`) and at run, row and cell level (`c79-30` to `c79-32`) |
| `w:txbxContent`, in a VML box | Joined in reverse, with two, three and four copies. Unexplained. In `wps:txbx` the last stands (`c79-33`) |
| `w:pPr`, `w:rPr`, `w:sectPr`, `w:tblPr` | Applied in order to one object, and the properties within them too (section 3.1) |
| `w:ind`, `w:spacing`, `w:numPr` and a mark's `w:rPr` in one `w:pPr`; `w:rFonts` in one `w:rPr`; `w:sectPr` in one `w:pPr`; `w:pgSz`, `w:pgMar` and `w:cols` in one `w:sectPr` | The same, as two copies of the parent (`sm-53`, `sm-70`, `sm-72`, `sm-73`, `op-10` to `op-21`) |

**The limits:** one small file per shape, one level in, all in a main document part but section 3.2's. Unmeasured: a `w:pPr` or
`w:rPr` not adjacent to its copy, a repeat in a `w:style` or a `w:lvl`, a DrawingML text box outside the main part, a second `w:tcPr` or `w:trPr`,
`a:ext`, `w:rFonts` twice in a mark's `w:rPr`, and a `w:col` list. core-ts CR-028's phase 0 lists most of these.

### 3.1 An attribute carried across `w:sectPr` copies

`rt-10-body` and `rt-40-sectPr` hold three copies whose `w:pgSz` are `11906 x 16838`, `15840 x 12240` with `w:orient="landscape"`,
and `12240 x 15840`. Word's save, in both files on both runs, is `w:w="12240" w:h="15840" w:orient="landscape"`: the third's size
with the second's orientation.

So the later value doesn't replace the whole of what came before. core-ts and the portfolio counted it in six files (`rt-10`,
`rt-40`, `rt-50`, `rt-51`, `nd-22`, `nd-24`), in nested bodies and a `Fallback` too.

The `sectpr-merge` round (`sm-*`, 40 files twice, read by docx4j-python, spot-checked by the portfolio) then settled it, and REQ-322
states it at WD 0.101. Word applies each copy in order to one object: what a copy states wins and what it doesn't state stays, for
attributes and child elements alike, one level in. An on/off element with no `w:val` states on (`sm-81`). Any other attribute left
out stays, even where the schema gives it a default (`w:num`, `sm-21`; `w:orient`, `sm-01`).

- **What it covers:** `w:pPr`, `w:rPr`, `w:sectPr` and `w:tblPr`, and within them `w:ind`, `w:spacing`, `w:rFonts`, `w:pgSz`,
  `w:pgMar`, `w:cols`, `w:jc`, `w:sz`, `w:b` and `w:vanish`. Phase 0 adds a `w:pPr`'s `w:rPr` (`c79-40`) and `w:numPr`, the first's
  `w:ilvl` with the second's `w:numId` (`c79-41`), and a first copy's `w:pPrChange` stays (`c79-42`).
- **A merge, not a sticky value:** a later `w:orient="portrait"` wins (`sm-07`).
- **Joined bodies:** their closing `w:sectPr` are applied the same way, a body with none passed over.
- **The exception:** a `w:headerReference` or `w:footerReference` on a later copy is ignored, with its part, whatever its `w:type` and
  across joined bodies. Only the first copy's stand, and a first copy with none leaves none (`sm-40` to `sm-42`, and the `footer-refs`
  round, `ft-*`, 12 files twice, each checked by the kept part's text). `ft-07`'s later footer goes even beside an earlier header,
  and `ft-08` and `ft-09`'s later `w:type="first"` goes with `w:titlePg` kept.
- **Unexplained:** `w:tblW`. `w:w="4000" w:type="dxa"` then `w:type="auto"` saved as `w:w="0" w:type="dxa"` (`sm-61`, one file). That
  may be how Word reads `w:type="auto"` with no `w:w` on its own, and REQ-322 says nothing of it.

### 3.2 In a header, a footer, a note and a comment (D7)

Two rounds of this package's, run by the portfolio on 10 November (bridge `da85101`, every job silent, the runs equal but for generated
ids, the PDFs agreeing with the saves; `test/fixtures/header-footer-probe/`). `header-repeats` (`hr-*`, 8 files) put seven measured
main-part cases into `word/header1.xml` as they are. `header-footer-repeats` (`hf-*`, 20 files) took a header's `w:pPr` through
seven forms, three copies of the other shapes, the same shapes in a footer, and main-part twins of the new markup.

| Given more than once, in a header or a footer | Word |
|---|---|
| `w:sdtContent`, `w:txbxContent` in a VML box, `w:rPr`, `w:tblPr`; `w:ind` and `w:rFonts` twice in one parent | As in the main part, with two and three copies (12 of 12) |
| A paragraph's `w:pPr` | The first copy's paragraph properties stand, and every later copy is ignored: values (center, then right, gives center), reversed (`ind`, then `jc`, gives `ind`), three copies, three values, an empty first copy (which leaves none), and no `w:jc` (11 of 11). The main-part twins merge |
| A paragraph mark's `w:rPr` in each `w:pPr` (`hf-16`) | The first's on the mark, and the later's on the paragraph's run, which the PDF draws in italic. One file |

Then the portfolio's round `notes-comments-repeats` (`nr-*`, 31 files and a control, built by its `scripts/probes/ncrepeat_build.py`,
bridge `ed94a27`, 10 November, every pair agreeing; `test/fixtures/notes-comments-probe/`): the header rounds' shapes in a footnote, an
endnote, a comment and a VML text box in a header, each container else one plain paragraph.

| Given more than once | Word |
|---|---|
| In a footnote, an endnote or a comment: `w:sdtContent`, a paragraph's `w:pPr` (values and merge), `w:rPr`, `w:tblPr`; `w:ind` and `w:rFonts` twice in one parent | As in the main part (7 shapes of 7 in each) |
| Two `w:txbxContent` in a VML box in a note or a comment | Repaired: each job reports unreadable content, and the box's words and the container's later lines are lost (3 of 3). Whether any VML box there does the same is unmeasured |
| In a VML text box in a header | As in the header: a paragraph's `w:pPr` keeps the first copy (`nr-tb-20`, `nr-tb-21`), and the other five shapes read as everywhere |

## 4. The change

`src/repeated.mts`, keyed by holder class (or a class it's based on) and property, as `wordInvalidValue` is scoped:

| Class (`org_docx4j_wml`, but for `CTTextbox`), property | Reading |
|---|---|
| `Document`, `body` | Join in order, the earlier's content at the head of the later's, and merge the closing `sectPr` |
| `SdtBlock`, `SdtRun`, `CTSdtRow`, `CTSdtCell`, `sdtContent` | Join in order |
| `org_docx4j_vml.CTTextbox`, `txbxContent` | Join in reverse: the earlier's content at the later's tail |
| `P`, `pPr`; `R`, `rPr`; `PPr`, `rPr` and `sectPr`; `PPrBase`, `ind`, `spacing` and `numPr`; `RPr`, `rFonts`; `Body`, `sectPr`; `SectPr`, `pgSz`, `pgMar` and `cols`; `Tbl`, `tblPr` | Merge: the later applied to the earlier |

**The merge,** as docx4j's `RepeatedChildren.merge`: for each property the earlier states, the later takes it if it states none.
Where both state one of a class in `MERGED` (`PPr`, `RPr`, `ParaRPr`, `SectPr`, `TblPr`, `PPrBase.Ind`, `PPrBase.Spacing`,
`PPrBase.NumPr`, `RFonts`, `SectPr.PgSz`, `SectPr.PgMar`, `CTColumns`), the two merge in turn. Otherwise the later's stands whole.
- **An on/off element** needs no rule of its own: a later `<w:b/>` replaces `<w:b w:val="0"/>` whole, and the model never has to say
  it states on. `w:jc` and `w:sz`, with one attribute, need none either.
- **References:** `SectPr` holds header and footer references in one list, `egHdrFtrReferences`. The later takes the earlier's list
  whole, or none if it had none, so only the first copy's reach the end.
- **A list** is the later's, as docx4j's merge leaves it (a list has no setter there): `w:cols`' `w:col`, unmeasured.
- **`w:tblW`** is the later's, or none if the later states none (docx4j's `AS_JAXB_LEAVES`, its CR-079 section 4).
- **In a header or a footer** (D7 (c), `FIRST_IN_HEADER_FOOTER`): a paragraph's repeated `w:pPr` keeps the first copy whole, the later
  copy's own properties removed, so an empty first copy keeps nothing. Where the paragraph stands is read from `keptNode`'s ancestors:
  the nearest WordprocessingML story element, `w:hdr` or `w:ftr`, through tables, controls, mc branches and text boxes, since a text
  box's paragraph reads as its host story's (`nr-tb-20`, `nr-tb-21`). `w:body`, a note and a comment end the search, so their
  paragraphs merge, in their text boxes too. A paragraph read with no such ancestor, in a `wml()` fragment say, merges. A later copy's mark `w:rPr`, which Word puts on the
  paragraph's run (`hf-16`), isn't followed.

`wordOverwritten(node, keptNode, propertyInfo, classInfo, value, keptValue)` edits `keptValue`, the copy the runtime keeps, in place,
and re-links what it moves. It returns whether it applied a reading, which the runtime ignores. `recordingOptions` passes it as
`onOverwrittenElement` unless the caller passes one (D3). With three copies the callback fires twice, so the kept copy gathers the
earlier ones in turn. The save then writes one container.

**Order against the other reports:** the runtime reports an overwrite once both copies are complete, so the later copy's `onElement`
has run. CR-017's `onMisplaced` and CR-016's `onOutOfRange` test each copy once, before the join, and a CR-011 record moves with its
`mc:Choice`.

### 4.1 Beyond the CR as accepted

docx4j's 17.3.6 holds what Word was measured on after D6, so D2's rule ("add a pair only when Word is measured") adds it here:
`w:sdtContent` at run, row and cell level, and `w:txbxContent` left alone in `wps:txbx` (`c79-30` to `c79-33`); a mark's `w:rPr`
and `w:numPr` merging in turn (`c79-40`, `c79-41`); and `w:numPr`, a mark's `w:rPr`, `w:sectPr`, `w:pgSz`, `w:pgMar` and `w:cols`
twice in one parent (`sm-53`, `op-10` to `op-21`). As accepted, section 4 named `w:tblW` "taken whole" and was silent on a list;
both follow docx4j now.

## 5. What jsonix asked

1. **Merge depth.** Property by property in the classes `MERGED` names, and attribute by attribute where those are attributes. An
   earlier `w:val="0"` can't survive a later `<w:b/>`, so `rc-31`, `rc-32`, `rt-31`, `rt-32`, `sm-80` and `sm-81` read to Word's
   effect.
2. **Unmeasured elements.** Unchanged: the runtime keeps the last, and a caller's own callback can see each one. A rule is added
   only where Word was measured (D2).
3. **In-place editing.** It is enough. All 101 cases were read through it. Nothing needed a returned value, so this gives no reason
   to change R3's recommendation.
4. **`parentPointers`.** Moved children are re-linked: over the 101 cases, 1,562 values walked, no `PARENT` wrong.

## 6. Results

`test/word365-repeated.mjs` reads docx4j's 101 files (section 3's rounds) and compares the facade's save of each with Word's, by
docx4j's `RepeatedChildrenTest`. Those comparisons are the words in order, every section whole, the target paragraph's properties,
the `sm-*` facts and the references by their part's text, plus each text box's words in order.
- **Equal:** every comparison but three: `rt-50`'s and `rt-51`'s words and `sm-52`'s sections. Each is an `mc:AlternateContent` root,
  whose `w:document` children the facade reads as a list, so nothing repeats until core-ts's root pass makes it one.
- **With the runtime's reading,** 73 comparisons in 68 files differed.
- **Word's own text:** `textOf` gives the text Word saved from each source in 84 files. The other 17 are 12 whose text holds a
  header's or a footer's words, two a list's label, and the three mc roots, all named, strict both ways.
- **Python's tolerances, as docx4j has them:** an off `w:b` or `w:vanish` reads as one not stated, and a portrait `w:orient` and a
  `w:cols w:space="720"` as Word's defaults, since Word's save writes nothing there (`rc-31`, `rt-32`, `sm-80`, `sm-07`, `sm-32`).
  `sm-61`'s `w:tblW` is held to the later copy's.
- **In a header and a footer** (D7, section 3.2): this package's 28 files, 29 stories, compared by each control's and text box's words
  and the target paragraph's every property by its path. 26 of the 29 stories match, in 25 of the 28 files. `hf-13` and `hf-14` keep a first copy's `w:jc left` and an empty
  first `w:pPr`, which Word's save leaves out, to the same effect, and `hf-16`'s later mark `w:rPr` isn't put on the run. All three
  are named, strict both ways. With the merge alone (D7 (a)) 8 more differ. Each table's `w:tblPr`, as `RepeatedChildrenTest`'s
  facts give it, is compared too, with no difference.
- **In notes, comments and a header's text box** (section 3.2): the portfolio's 32 files, 35 parts, compared as the header rounds
  are. 32 parts match. The three are the `w:txbxContent` pairs Word repairs in a note or a comment, where the facade joins them in
  reverse; named, strict both ways. With a text box as a story of its own, as first built, `nr-tb-20` and `nr-tb-21` differ too.
- **What a save does:** read again, no file repeats a child. A `deepCopy` writes the same.
- **The gate:** the 341 fidelity hashes are identical to those of `7a85af2` (main), on jsonix 3.13.0. `npm run typecheck` and `npm test` pass on
  Node 22.23.4 and 24.15.0; CI runs 18, 20 and 22.
- **Mutants:** 34 of `dist/repeated.mjs`, one rule each, all caught by the smoke or the Word test. Seven are D7's: the rule off, a
  footer not a story, a text box one, a table cell or a `w:pict` ending the search, and the later copy's own properties or `PARENT`
  kept.

**core-ts,** 2026-11-10, on a tarball of the working tree before D7 (c) (sha1 `514e7665`), its main `b25f9e3`, Node 22, jsonix 3.13.0.
It found nothing against it:
- **Its suite:** typecheck passes, and 1,982 of 1,982, the same as on 0.8.0, no test moving.
- **Its bytes:** a load, `unmarshalAll` and save of 1,615 packages (its `test/`, docx4j's `charts/` and `repeated/`), every entry
  hashed, 0.8.0 against the candidate. 69 of 19,719 entries differ, each a `word/document.xml` of a `repeated/` source. Nothing else
  moved, Word's saves there included.
- **Through its own load:** Word's text, line for line with list labels, in 95 of 107 sources, 75 on 0.8.0, none lost. Of the 12, three
  are the mc roots, seven hold Word's lines in another order (its text model puts headers first), and two have the merged list label
  separated by a tab where Word wrote a space.
- **Its traps:** `isUnchangedSinceRead` is true on all 112 main parts after `unmarshalAll` (section 9's third), and a save read back
  gives the same text in all 112.
- **Not verified there:** Node 18 and 20, strict mode, and the published tarball, which its step 3 reads again at its range raise.

**core-ts again,** on the candidate with D7 (c) (sha1 `e8576a62`), the same day and setup: it passes, nothing moving beyond the rule.
- **Its suite and bytes:** 1,982 of 1,982. The same 69 entries change against 0.8.0, and none against the first candidate. docx4j's
  `repeated/` reads through its load exactly as on the first.
- **This package's 28 header and footer files, through its load:** each story's paragraphs (`w:pPr`, the mark's and each run's
  `w:rPr`, the text) held to Word's save. 26 of 28 agree, against 4 on 0.8.0. The two are `hf-13`'s `w:jc left`, to the same effect,
  and `hf-16`'s run italic, both stated here. Its comparison takes `hf-14`'s empty `w:pPr` as Word's.
- **Not run there:** unchanged-since-read on the 28, and Node 18 and 20.

**core-ts on the text box fix,** 2026-11-13, on a third candidate (sha1 `5905202b`), its main `1168069`, Node 22.23.4: it passes,
nothing in its corpus moving.
- **Its suite:** typecheck passes, and 2,032 of 2,032, as on its stock install.
- **Its bytes:** 1,686 packages loaded, read whole and saved, 20,718 entries hashed, against `e8576a62` on the same build: none differ.
  Nothing there has a repeated `w:pPr` in a header's text box.
- **Through its own load:** the 28 header and footer files as before, 26 of 28. The portfolio's 32: 29 match Word's save, 27 on
  `e8576a62`, the two that moved being `nr-tb-20` and `nr-tb-21`; the other three are the files Word repairs. docx4j's `repeated/`
  reads as on the first two candidates, and all 112 main parts are unchanged since read.
- **The release:** core-ts packed the pushed `7958467` itself on 2026-11-18 and got this candidate, its go standing. 0.9.0's
  tarball (`c9b6483b`) differs from it in `package.json` alone, and its 341 marshalled hashes equal the tree's (`RELEASING.md` step 6).

## 7. Decisions

D1 to D5 decided as recommended, and D6 as recommended with the change it states, all 2026-11-07. D5 is settled from core-ts's side
by its D2's route (c), 2026-11-10. D7 decided (c) as recommended, 2026-11-10 (Jason: "D7 (c) as recommended").

| D | Question | Decided |
|---|---|---|
| D1 | Which pairs? | The five first in section 4, which D6 and section 4.1 added to, including `w:txbxContent`'s reverse order, since the key is the one shape measured and its words are otherwise lost. The rule is by class, so it also applies outside the main part, where REQ-322 speaks of the main part alone: a header and a footer, measured since and refined by D7, and a note or a comment, measured since, which read as the main part |
| D2 | Everything else? | Leave it to the runtime, and add a pair only when Word is measured |
| D3 | On by default? | Yes, as CR-013's reading is: it changes only a file with the shape. A caller's `onOverwrittenElement` replaces it, and `wordOverwritten` is exported so a caller can call it too, as `wordInvalidValue` is |
| D4 | Is the join reported? | Not by the facade yet. core-ts CR-028 D8, the code, is open with docx4j and python, and docx4j logs one WARN a part with a count and no code. A caller counts through `wordOverwritten`'s return meanwhile |
| D5 | This or core-ts CR-028's containers pass? | This: core-ts's D2 took route (c), 2026-11-10. Its root pass stays |
| D6 | Widen section 4 to REQ-322, as it stood at WD 0.100? | Yes, as D2 said to add what Word was measured on, and `sm-*` measures it. References are the exception, only the first copy's standing (`ft-*`) |
| D7 | By class, or the main document part's alone, as docx4j? | **(c), as recommended** (Jason, 2026-11-10: "D7 (c) as recommended"), built as section 4 has it. Measured in two rounds first (Jason, 2026-11-10: "D7 let's measure a header", then "run a further round: a footer, three copies, and other properties"). Word 16.0.20326.20232, bridge da85101, every job silent, each file's two runs equal but for generated ids. `header-repeats` (`hr-*`, 8 files) put seven measured main-part cases into `word/header1.xml` as they are; `header-footer-repeats` (`hf-*`, 20 files) took a header's `w:pPr` through values, reversed order, three copies, an empty first copy, no `w:jc` and a mark's `w:rPr`, three copies of the other shapes, the same shapes in a footer, and main-part twins of the new markup. **Every shape but `w:pPr` reads in a header and a footer as in the main part**, 12 of 12: `w:sdtContent` joined in order and `w:txbxContent` in reverse (two and three copies), `w:rPr` merged (two and three), `w:tblPr`, and `w:ind` and `w:rFonts` twice in one parent. **A paragraph's repeated `w:pPr` doesn't:** in all 11 cases, header and footer alike, the first copy's paragraph properties stand and every later copy's are ignored, even where the first is empty (`hf-14`), where the main part merges them (the twins `hf-51`, `hf-54` and `hf-55` merge). The saves and the PDFs agree. In `hf-16`, the later copy's mark `w:rPr` (`w:i`) lands on the paragraph's run, which the PDF draws in italic. Against the 23: by class, as built, matches 12. docx4j's reading, the main part's alone and the last copy elsewhere, matches none. Python's, the main part's alone and its fork's first copy elsewhere, matches 11, the `w:pPr` files but `hf-16`, and `hr-60`; it loses the later `w:sdtContent`'s words with no warning and refuses an edited save (python CR-039 section 10.3, measured at `9f565b6`). *Options:* (a) by class as built; (b) the main part's alone, as docx4j; (c) by class, and in a `w:hdr` or `w:ftr` a paragraph's repeated `w:pPr` keeps the first copy whole: 12 of 12 and the `w:pPr` cases in 8 exactly and 2 to the same effect (a first `jc left` and an empty first copy, the defaults Word leaves out), `hf-16`'s run italic not followed. *Recommended:* (c). Notes and comments stay by class, as the portfolio's round then measured them (section 3.2). A text box was first taken as a story of its own, unmeasured, and so merged in a header; the same round showed a header's text box reads as the header (`nr-tb-20`, `nr-tb-21`), and it does here since 2026-11-13 (Jason: "do The CR-022 text box fix"). Python proposes to read headers and footers as (c) does (its CR-039 phases E and F, proposed, nothing built) |

## 8. Tests and docs

- **`test/word365-repeated.mjs`:** section 6, on `test/fixtures/repeated-probe/` (each source's parts beside Word's save, a
  `SOURCE.md`, and a SHA-1 manifest of docx4j's files), `test/fixtures/header-footer-probe/` (this package's two rounds, the same
  layout, `cases.json`, and a SHA-256 manifest of the probes and Word's saves) and `test/fixtures/notes-comments-probe/` (the
  portfolio's round, laid out the same, each case's story part alone).
- **Smoke:** the tables held to the context; three bodies joined with their `w:sectPr` applied in order and only the first copy's
  references, headers and footers together; `PARENT` after a join and after merges at each level; the children twice in one parent;
  a list the later's; run-level `w:sdtContent`; a VML `w:txbxContent` in reverse; `w:tblW` the later's; a pair not in the table left
  to the runtime; a caller's callback first; `unmarshalNode`, `unmarshalPackage`, `wml()` and a footnote alike; a save read again
  repeating nothing; `deepCopy`; and CR-017's report and CR-011's record through a join. D7: a header's and a footer's paragraph
  keeping the first of three `w:pPr`, `PARENT` re-linked by identity, an empty first copy, through a table and a header's text box; a
  text box in the main part, a footnote and a comment, and a `wml()` fragment, merging; a header's `w:rPr` still merging; the story
  names held to the context.
- **Docs:** `CLAUDE.md` and a README bullet.

## 9. Consumers

- **core-ts:** its CR-028 phase 2 takes 0.9.0, counting through `wordOverwritten`'s return at its read and re-read once
  CR-028 D8's code is agreed. Verified on both candidates, before D7 (c) and with it (section 6). Three traps its section 3.2
  names, and how this meets them:
  - Carrying "what the lost copy holds and the kept doesn't" would carry `sm-81`'s `w:val="0"` over a `<w:vanish/>`. The merge takes
    nothing the later states, and `w:vanish` isn't in `MERGED`, so the later's stands whole.
  - That test never fires on `sm-41` or `ft-02`, where the kept copy holds the reference and has to lose it. The references are
    replaced by the earlier's, or removed, whatever the later holds.
  - Its read (`XmlPart.mts:441`) and re-read (`:335`, `isUnchangedSinceRead`) must read alike. Both go through the facade with the
    default, so both join. A read that passes its own `onOverwrittenElement` must call `wordOverwritten` from it, in both.
  A later copy's references leave the tree. This package holds no relationships, so each part and relationship stays unless
  core-ts drops it, as Word's save does.
- **docx4j-ts-editor and add-ins:** a file with the shape reads whole, and a save writes one container.
