# CR-022: A repeated child read as Word reads it, through jsonix-CR-016's callback

**Status:** Accepted 2026-11-07, D1 to D5 as recommended (Jason: "CR-022 as recommended") and D6 the same day (Jason: "D6: widen
CR-022 to REQ-322"; section 7). Nothing implemented, and `jsonix/CR-016` is released. Proposed the same day. Low priority: no file in
either corpus has the shape.
**Requested by:** Jason, 2026-11-07 ("ask objects for that CR"), relayed by jsonix-4b. The CR, not the implementation.
**Depends on:** `jsonix/CR-016`, released in `@docx4j/jsonix` 3.13.0 (2026-11-07, tag `a5359c4`). Its `onOverwrittenElement` is the
hook, so this needs a raise of the floor here.
**Counterparts:**
- WD REQ-322 (WD 0.101, portfolio `7150868`) is the reading, for a main document part.
- core-ts CR-028 section 3.2 proposes the same reading as DOM passes in its load, and may be deferred (D5).
- docx4j's CR-079 and docx4j-python's CR-039 are the other ports'.
- jsonix stays at the report, because the runtime holds no Word rules: the same split as jsonix-CR-010 and CR-013.

## 1. Summary

A single-valued property given a second element keeps the last. jsonix-CR-016 reports that, but a report leaves the first
`w:body`'s words out of the tree, and Word shows them.

This CR has every facade read pass a `wordOverwritten` callback that reads the measured pairs as WD REQ-322 and Word do: `w:body`
and block `w:sdtContent` joined in order, `w:txbxContent` in a VML box joined in reverse, and `w:pPr`, `w:rPr`, `w:sectPr` and
`w:tblPr` applied in order to one object, one level in. Anything else stays as the runtime leaves it.

In a scratch prototype 80 of Word's 85 measured saves match, against 25 today, with no fidelity hash moved. Four of the other five
hold a value Word's save leaves out, to the same effect. The fifth is `w:tblW`, which nothing explains yet (section 3.1).

## 2. What is there

`PropertyInfo.setProperty` overwrites, so the last copy stands, with no report. That matches Word for an element with one attribute.
Elsewhere it loses what the earlier copies stated: 60 of the 85 cases differ from Word's save today.

## 3. What Word reads

Word 16.0.20326.20232, each file twice, all silent, the runs equal apart from generated ids. On the share under
`bridge/portfolio/`: `rc-*` two copies, `rt-*` three and four, `sm-*` and `ft-*` one difference a file. `rt-50`, `rt-51` and `sm-52`,
which are mc roots, aren't this CR's.

| Given more than once | Word |
|---|---|
| `w:body` | Joined in order, the closing `w:sectPr` applied in order to one object, a body with none passed over. One in a `w:pPr` stays |
| `w:sdtContent`, block level | Joined in order |
| `w:txbxContent`, in a VML box | Joined in reverse, with two, three and four copies. Unexplained, and the only shape tried |
| `w:pPr`, `w:rPr`, `w:sectPr`, `w:tblPr` | Applied in order to one object, and the properties within them too (section 3.1) |

**The limits:** one small file per shape, one level in. Unmeasured: a `w:pPr` or `w:rPr` not adjacent to its copy, a repeat in a
header, a note or a comment, in a `w:style` or a `w:lvl`, a second `w:tcPr` or `w:trPr`, and `a:ext`. core-ts CR-028's phase 0 lists
most of these. No case holds a `w:rPr` in a `w:pPr` or a `w:col`.

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
  `w:pgMar`, `w:cols`, `w:jc`, `w:sz`, `w:b` and `w:vanish`.
- **A merge, not a sticky value:** a later `w:orient="portrait"` wins (`sm-07`).
- **Joined bodies:** their closing `w:sectPr` are applied the same way, a body with none passed over.
- **The exception:** a `w:headerReference` or `w:footerReference` on a later copy is ignored, with its part, whatever its `w:type` and
  across joined bodies. Only the first copy's stand, and a first copy with none leaves none (`sm-40` to `sm-42`, and the `footer-refs`
  round, `ft-*`, 12 files twice, each checked here by the kept part's text). `ft-07`'s later footer goes even beside an earlier
  header, and `ft-08` and `ft-09`'s later `w:type="first"` goes with `w:titlePg` kept.
- **Unexplained:** `w:tblW`. `w:w="4000" w:type="dxa"` then `w:type="auto"` saved as `w:w="0" w:type="dxa"` (`sm-61`, one file). That
  may be how Word reads `w:type="auto"` with no `w:w` on its own, so it may be no exception, and REQ-322 says nothing of it.

## 4. The change

In a new `src/repeated.mts`, keyed by class and property, as `wordInvalidValue` is scoped:

| Class, property (`org_docx4j_wml`, but for `CTTextbox`) | Reading |
|---|---|
| `Document`, `body` | Join in order, and apply the closing `sectPr` in order |
| `SdtBlock`, `sdtContent` | Join in order |
| `org_docx4j_vml.CTTextbox`, `txbxContent` | Join in reverse |
| `P`, `pPr`; `R`, `rPr`; `Body`, `sectPr`; `PPr`, `sectPr`; `Tbl`, `tblPr` | Apply in order: for each property the kept copy lacks, take the lost copy's |
| `PPr`, `ind` and `spacing`; `RPr`, `rFonts`; `SectPr`, `pgSz`, `pgMar` and `cols` | The same, attribute by attribute, whether the property repeats in one copy or its holder repeats |

- **Every other property a later copy states replaces the earlier whole.** So an on/off element with no `w:val` needs no rule of its
  own: a later `<w:b/>` replaces `<w:b w:val="0"/>`, and the model never has to say it states on. `w:jc` and `w:sz`, with one
  attribute, need none either.
- **References:** `SectPr` holds header and footer references in one list, `egHdrFtrReferences`. When copies are applied, the kept
  copy takes the lost copy's list whole, or none if it had none, so only the first copy's reach the end.
- **`w:tblW`** isn't in the table, so it is taken whole.

`wordOverwritten(node, keptNode, propertyInfo, classInfo, value, keptValue)` edits `keptValue` in place and re-links what it moves.
`recordingOptions` passes it as `onOverwrittenElement` unless the caller passes one (D3). With three copies the callback fires
twice, so the kept copy gathers the earlier ones in turn. The save then writes one container.

## 5. What jsonix asked

1. **Merge depth.** Property by property in `w:pPr`, `w:rPr`, `w:sectPr` and `w:tblPr`, and attribute by attribute in the six
   properties section 4 names. An earlier `w:val="0"` can't survive a later `<w:b/>`, so `rc-31`, `rc-32`, `rt-31`, `rt-32`, `sm-80`
   and `sm-81` read to Word's effect.
2. **Unmeasured elements.** Unchanged: the runtime keeps the last, and a caller's own callback can see each one. A rule is added
   only where Word was measured (D2).
3. **In-place editing.** It is enough. All 85 cases were read through it, the merges at both levels, the references and
   `w:txbxContent` among them, with two, three and four copies. Nothing needed a returned value, so this gives no reason to change
   R3's recommendation.
4. **`parentPointers`.** Moved children need re-linking, and the prototype does it. Over the 85 cases, 1,328 values walked, no
   `PARENT` was wrong. `PPr` and `RPr` hold no collections. `SectPr`'s references are the one list a merge moves, and they
   re-link too.

## 6. The prototype

In a scratch copy of `106319f` on npm's jsonix 3.13.0, with `src/repeated.mts` as section 4 has it and one line in `recordingOptions`.

- Marshalled output against Word's save, on 85 cases (34 `rc-*` and `rt-*`, 39 `sm-*`, 12 `ft-*`): the words in order, and every
  child of each `w:pPr`, `w:rPr`, `w:sectPr` and `w:tblPr` with its attributes, beside Word's own additions (`w:cols w:space="720"`,
  `w:lang`, `w:tblCellMar`, `w:tblInd`, `w:tblLook`). Word renumbers references, so they are compared by type, and by the part's
  text in `sm-40` to `sm-42` and `ft-*`. **80 match, against 25 today:** 32 of the 34 (17 today), 36 of the 39 (5 today) and all 12
  (3 today).
- The five. `rc-31`, `rt-32` and `sm-80` keep a `w:val="false"` element where Word's save drops it, and `sm-07` keeps
  `w:orient="portrait"`, the default, where Word's save leaves it out, all to the same effect. `sm-61` holds `w:tblW w:type="auto"`
  where Word's save has `w:w="0" w:type="dxa"`.
- A save read again reports no overwrite, and a `deepCopy` marshals the same.
- `npm run typecheck` and `npm test` pass on Node 18.20.8, 20.20.2 and 22.23.4, and all 341 marshalled hashes are identical to
  `106319f`'s on 3.13.0 on each.

**Cost,** on the same runtime: one process per run, the median of 25 reads of 8,000 runs, six runs each. The median of the six is 166.2
ms unset, 165.5 with a no-op callback and 168.0 with `wordOverwritten`, and the ranges overlap. That is about 2 ms in 166.

core-ts hasn't been run against it.

## 7. Decisions

D1 to D5 decided as recommended, and D6 as recommended with the change it states, all 2026-11-07.

| D | Question | Decided |
|---|---|---|
| D1 | Which pairs? | The five first in section 4, which D6 added to, including `w:txbxContent`'s reverse order, since the key is the one shape measured and its words are otherwise lost. WD REQ-322 requires nothing of it, and core-ts's D7 would wait for a second shape. The rule is by class, so it also applies in a header or a note, unmeasured, where REQ-322 speaks of the main part |
| D2 | Everything else, with `w:sectPr` in view after section 3.1? | Leave it to the runtime, and add a pair only when Word is measured. Don't build a `w:sectPr` merge on one attribute; put different properties and attributes in a Word round first |
| D3 | On by default? | Yes, as CR-013's reading is: it changes only a file with the shape. A caller's `onOverwrittenElement` replaces it, and `wordOverwritten` is exported so a caller can call it too, as `wordInvalidValue` is |
| D4 | Is the join reported? | Not by the facade yet. The WD's join warning is open for Jason, and a file read as Word reads it has lost nothing. A caller can wrap `wordOverwritten` meanwhile |
| D5 | This or core-ts CR-028's containers pass? | One, not both. core-ts reads through this facade, so this would give it section 3.2 with no pass of its own, and its root pass stays. That is core-ts's call with Jason, and CR-028 may be deferred anyway |
| D6 | Widen section 4 to REQ-322, as it stood at WD 0.100: `w:sectPr` and `w:tblPr` added, the merge taken one level in, and a join's closing `w:sectPr` applied in order? | Yes, as D2 said to add what Word was measured on, and `sm-*` measures it. As first accepted, a join kept the last copy's `w:sectPr` whole, its references included, where Word keeps the first's. References are the exception, only the first copy's standing (`ft-*`), and `w:tblW` stays out of the merge until it is explained. The recommendation expected a rule for an on/off element with no `w:val`. The prototype needed none, since the merge goes by attribute only in six named properties |

## 8. Tests and docs

- **`test/word365-repeated.mjs`, a further test:** the `rc-*`, `rt-*`, `sm-*` and `ft-*` parts under `test/fixtures/repeated-probe/`, each
  beside Word's `word/document.xml`, with a `SOURCE.md`. Marshalled output compared as in section 6, the five differences stated,
  strict both ways.
- **Smoke:** the table held to the context, a caller's callback taking precedence, `PARENT` after a join and a merge at each level, a
  pair not in the table left to the runtime, an on/off element and `w:tblW` taken whole, and references with three copies, headers and
  footers together.
- **The gate:** every fidelity hash identical.
- **Docs:** `CLAUDE.md` and a README bullet.

## 9. Consumers

- **core-ts:** D5. Its gate compares a re-marshal of the source with the tree, and both pass through this reading, so a file with
  the shape should still count as unchanged. A later copy's references leave the tree. This package holds no relationships, so
  each part and relationship stays unless core-ts drops it, as Word's save does. Not run.
- **docx4j-ts-editor and add-ins:** a file with the shape reads whole, and a save writes one container.
