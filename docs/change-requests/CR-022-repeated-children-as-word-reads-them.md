# CR-022: A repeated child read as Word reads it, through jsonix-CR-016's callback

**Status:** Accepted 2026-11-07, D1 to D5 as recommended (Jason: "CR-022 as recommended"; section 7). Nothing implemented, and it
waits on `jsonix/CR-016`. Proposed the same day. Low priority: no file in either corpus has the shape.
**Requested by:** Jason, 2026-11-07 ("ask objects for that CR"), relayed by jsonix-4b. The CR, not the implementation.
**Depends on:** `jsonix/CR-016` (`ec11e3e`, local; R1 to R3 open; not implemented). Its `onOverwrittenElement` is the hook, so a
change to its signature changes section 4, and this needs a jsonix release and a raise of the floor here.
**Counterparts:**
- WD REQ-322 (WD 0.99) is the reading, for a main document part.
- core-ts CR-028 section 3.2 proposes the same reading as DOM passes in its load, and may be deferred (D5).
- docx4j's CR-079 and docx4j-python's CR-039 are the other ports'.
- jsonix stays at the report, because the runtime holds no Word rules: the same split as jsonix-CR-010 and CR-013.

## 1. Summary

A single-valued property given a second element keeps the last. jsonix-CR-016 reports that, but a report leaves the first
`w:body`'s words out of the tree, and Word shows them.

This CR has every facade read pass a `wordOverwritten` callback that reads five measured pairs as Word does: `w:body` and block
`w:sdtContent` joined in order, `w:txbxContent` in a VML box joined in reverse, and `w:pPr` and `w:rPr` merged with the later
value winning. Anything else stays as the runtime leaves it.

In a scratch prototype 30 of Word's 34 measured saves match, against 17 today, with no fidelity hash moved. Of the other four,
two are a `w:b` Word drops to the same effect. The other two are a finding: Word carried an attribute across `w:sectPr` copies
(section 3.1).

## 2. What is there

`PropertyInfo.setProperty` overwrites, so the last copy stands, with no report. That matches Word for a leaf. In a container it
loses the earlier copies: 17 of the 34 cases differ from Word's save today, 4 of them the two kinds above.

## 3. What Word reads

Word 16.0.20326.20232, each file twice, all silent, the runs equal apart from generated ids. On the share under
`bridge/portfolio/`: `rc-*` two copies, `rt-*` three and four. `rt-50` and `rt-51`, which are mc roots, aren't this CR's.

| Given more than once | Word |
|---|---|
| `w:body` | Joined in order. The last closing `w:sectPr` stated stands, a body with none passed over. One in a `w:pPr` stays |
| `w:sdtContent`, block level | Joined in order |
| `w:txbxContent`, in a VML box | Joined in reverse, with two, three and four copies. Unexplained, and the only shape tried |
| `w:pPr`, `w:rPr` | Merged, the later value winning |
| `w:jc`, `w:b`, `w:sz` | The last |
| `w:sectPr`, `w:tblPr` | The later value, with section 3.1's exception. Each copy states the same property names |

**The limits:** one small file per shape, one level deep, one property each way. Unmeasured: a `w:pPr` or `w:rPr` not adjacent to
its copy, a repeat in a header, a note or a comment, in a `w:style` or a `w:lvl`, a second `w:tcPr` or `w:trPr`, `a:ext`, and a
repeated property with different attributes. core-ts CR-028's phase 0 lists most of these.

### 3.1 An attribute carried across `w:sectPr` copies

`rt-10-body` and `rt-40-sectPr` hold three copies whose `w:pgSz` are `11906 x 16838`, `15840 x 12240` with `w:orient="landscape"`,
and `12240 x 15840`. Word's save, in both files on both runs, is `w:w="12240" w:h="15840" w:orient="landscape"`: the third's size
with the second's orientation.

So the later value doesn't replace the whole of what came before, at least for this attribute. Word appears to apply each copy
over what it already holds. core-ts and the portfolio have since counted it in six files (`rt-10`, `rt-40`, `rt-50`, `rt-51`,
`nd-22`, `nd-24`), in nested bodies and a `Fallback` too, and WD REQ-322's NOTE records it. No file has a later copy stating
`w:orient="portrait"`, and it is still one attribute, so the rule is an inference. If it holds generally, "the last" for a leaf with several attributes may be wrong too (`w:ind`, `w:spacing`, `w:rFonts`).

## 4. The change

In a new `src/repeated.mts`, keyed by class and property, as `wordInvalidValue` is scoped:

| Class, property | Reading |
|---|---|
| `org_docx4j_wml.Document`, `body` | Join in order, and take the lost copy's `sectPr` if the kept one has none |
| `org_docx4j_wml.SdtBlock`, `sdtContent` | Join in order |
| `org_docx4j_vml.CTTextbox`, `txbxContent` | Join in reverse |
| `org_docx4j_wml.P`, `pPr`; `org_docx4j_wml.R`, `rPr` | For each property the kept copy lacks, take the lost copy's |

`wordOverwritten(node, keptNode, propertyInfo, classInfo, value, keptValue)` edits `keptValue` in place and re-links what it moves.
`recordingOptions` passes it as `onOverwrittenElement` unless the caller passes one (D3). With three copies the callback fires
twice, so the kept copy gathers the earlier ones in turn. The save then writes one container.

## 5. What jsonix asked

1. **Merge depth.** Property by property in `w:pPr` and `w:rPr`, each property whole, so an earlier `w:val="0"` can't survive a
   later `<w:b/>`. `rc-31`, `rc-32`, `rt-31` and `rt-32` read as Word does. Section 3.1 is the caution: Word itself may go deeper
   for a property with several attributes, which only a Word round can settle.
2. **Unmeasured elements.** Unchanged: the runtime keeps the last, and a caller's own callback can see each one. A rule is added
   only where Word was measured (D2).
3. **In-place editing.** It is enough. All 34 cases were read through it, the `w:pPr` and `w:rPr` merges and `w:txbxContent` among
   them, with two, three and four copies. Nothing needed a returned value, so this gives no reason to change R3's recommendation.
4. **`parentPointers`.** Moved children need re-linking, and the prototype does it. Over the 34 cases, 589 values walked, no
   `PARENT` was wrong. `PPr` and `RPr` hold no collections, so a merge never has to join a list.

## 6. The prototype

In a scratch copy of the worktree (CR-021 in it), on jsonix 3.12.0 with jsonix-CR-016's sketch patched into its factory, Node 22.

- Marshalled output against Word's save, on 34 cases: the words in order, every `w:pPr` and `w:rPr` child with its attributes, and
  the `w:sectPr` and `w:tblPr` beside Word's own additions (`w:cols`, `w:tblCellMar`, `w:tblLook`). **30 match, against 17 today.**
- The four: `rc-31` and `rt-32` keep `w:b w:val="false"` where Word's save drops it, to the same effect. `rt-10` and `rt-40` lack
  section 3.1's `w:orient`.
- A save read again reports no overwrite, and a `deepCopy` marshals the same.
- `npm test` passes, and all 341 marshalled hashes are identical.

core-ts hasn't been run against it. The prototype casts the option, since jsonix's typings don't have it yet.

## 7. Decisions

All five decided as recommended, 2026-11-07.

| D | Question | Decided |
|---|---|---|
| D1 | Which pairs? | The five in section 4, including `w:txbxContent`'s reverse order, since the key is the one shape measured and its words are otherwise lost. WD REQ-322 requires nothing of it, and core-ts's D7 would wait for a second shape. The rule is by class, so it also applies in a header or a note, unmeasured, where REQ-322 speaks of the main part |
| D2 | Everything else, with `w:sectPr` in view after section 3.1? | Leave it to the runtime, and add a pair only when Word is measured. Don't build a `w:sectPr` merge on one attribute; put different properties and attributes in a Word round first |
| D3 | On by default? | Yes, as CR-013's reading is: it changes only a file with the shape. A caller's `onOverwrittenElement` replaces it, and `wordOverwritten` is exported so a caller can call it too, as `wordInvalidValue` is |
| D4 | Is the join reported? | Not by the facade yet. The WD's join warning is open for Jason, and a file read as Word reads it has lost nothing. A caller can wrap `wordOverwritten` meanwhile |
| D5 | This or core-ts CR-028's containers pass? | One, not both. core-ts reads through this facade, so this would give it section 3.2 with no pass of its own, and its root pass stays. That is core-ts's call with Jason, and CR-028 may be deferred anyway |

## 8. Tests and docs

- **`test/word365-repeated.mjs`, a further test:** the `rc-*` and `rt-*` parts under `test/fixtures/repeated-probe/`, each beside
  Word's `word/document.xml`, with a `SOURCE.md`. Marshalled output compared as in section 6, the four differences stated, strict
  both ways.
- **Smoke:** the table held to the context, a caller's callback taking precedence, `PARENT` after a join and a merge, a pair not
  in the table left to the runtime, and a Word-measured leaf kept whole.
- **The gate:** every fidelity hash identical.
- **Docs:** `CLAUDE.md` and a README bullet.

## 9. Consumers

- **core-ts:** D5. Its gate compares a re-marshal of the source with the tree, and both pass through this reading, so a file with
  the shape should still count as unchanged. Not run.
- **docx4j-ts-editor and add-ins:** a file with the shape reads whole, and a save writes one container.
