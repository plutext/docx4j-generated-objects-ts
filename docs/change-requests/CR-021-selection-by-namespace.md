# CR-021: `mcBranchOf` by namespace, over the understood set, selecting nothing as Word does

**Status:** Implemented 2026-11-07, D1 to D5 as decided. Not released. Accepted the same day, D1 to D5 as recommended (Jason:
"CR-021 decisions as recommended. please implement"; section 7). Proposed the same day.
**Requested by:** Jason, 2026-11-07 ("write the objects CR for the four items: removing the third step, the
mce.nothing_selected warning, the new understood set, and matching by URI"). It follows his rulings recorded in the WD:
"REQ-305 itself changes for all three engines", "selecting nothing should raise a warning" and "two sets".
**Depends on:**
- WD REQ-302 to REQ-306 and REQ-801, at WD 0.99 (portfolio `0309843`). REQ-305's order holds: the WD, then docx4j's list,
  then each engine.
- docx4j's published understood set, `understood-namespaces.txt`, 114 URIs, at `824565363` (local, unpushed when read).
- Nothing from the runtime, the compiler or `modules/`.

**Counterparts:**
- docx4j's `McSelection` already matches by URI for text (`requiresNamespaces`, `understands`). Its
  `selectedBranch(McPurpose.TEXT, ...)` still takes the first Choice at `824565363`, and moves on its own.
- core-ts CR-010 D4 item 1 asked for URIs here in October, and this package never made the change. core-ts raises
  `mce.nothing_selected` itself, and has to change its one call of `mcBranchOf` (section 9).
- docx4j-python moves on its own.

## 1. Summary

`mcBranchOf` chooses an `mc:AlternateContent`'s branch, and `textOf` reads through it. It differed from Word and the WD in
four ways, and this CR changed all four:

1. **The third step went.** With nothing understood and no Fallback it took the first Choice, where Word shows neither branch.
2. **Selecting nothing is reported,** through a callback core-ts can raise `mce.nothing_selected` from.
3. **The default set is the understood set:** docx4j's 114 URIs in place of `NAMESPACE_PREFIXES`, which stays for
   marshalling.
4. **`Requires` is matched by URI,** through CR-011's records, where prefixes were compared.

`textOf` now matches Word on all 142 measured files, against 118 before, and no marshalled hash moved.

## 2. What was there

`mcBranchOf` took the first Choice whose `Requires` prefixes were all in its set, else the Fallback, else the first Choice,
else nothing. Its default set was every prefix in `NAMESPACE_PREFIXES`.

- **The third step** was WD REQ-305 until 0.97, a documented departure from ISO 29500-3 9.4, pinned by the smoke's "rule 3".
- **The set is the wrong one.** The table was built for marshalling. Of its 136 prefixes, 27 aren't in the understood set:
  11 Word was measured not to understand (`properties`, `vt`, `cp`, `dc`, `dcterms`, `cppr`, `xvml`, `pvml`, `prop`, `cx5`,
  `cx8`) and 16 never measured. Six of the 114 have no entry in it: SpreadsheetML main, MathML, InkML, `xr9`, `c16r2`, `ax`.
- **By prefix.** REQ-303 allows the conventional table only where the model keeps no declarations. Since CR-011 a read
  through the facade keeps them: each `mc:Choice` holds the binding of any prefix the table wouldn't give. So the allowance
  now fits a hand-built tree alone. core-ts measured the effect: all 98 URIs in both sets, each `ul-*` file given a
  Fallback and using the prefix `q`, took the Fallback through its kept load.
- **Nothing reported.** `mcBranchOf` and `textOf` had no way to say a branch was lost, and neither knows the part's name.

## 3. What Word reads

Word 16.0.20326.20232, each file twice, all silent, the runs agreeing, on the share under `bridge/portfolio/`:

| Round | Files | Shows |
|---|---|---|
| `ul-*` | The 118 URIs of the old list, `w14` among them as a control, and a made-up control | The Choice for 105; neither branch for 13 and the made-up control |
| `uc-*` | 11 namespaces the list lacked, and two controls | The Choice for `cx4`, `cx6`, `cx7`, `c16r2`, `ax`, `xsi`; neither for `cx5`, `cx8`, `mo`, `arto`, `dcmitype` |
| `c305-*` | Block and run level, `wpi`, `oel`, `cx3`, `w16sdtfl`, unknown namespaces | The Choice where understood; neither otherwise |
| `rw-25` | An unknown namespace in a body, no Fallback | Neither |

**The limit:** every file is `Choice`-only, in a main document part, on one Word build. REQ-306 says so: the set is silent on
spreadsheets and presentations, and a namespace not measured counts as not understood. That a Fallback is taken where a
Choice isn't understood is REQ-302 and ISO; these rounds didn't put one in.

## 4. The change

- **`src/understood.mts`:** `UNDERSTOOD_NAMESPACES`, the 114 URIs in docx4j's order, public from `./builders/wml`.
- **`requiresNamespaces(choice)`,** docx4j's name, public from the facade, where `sourceBindings` lives: for each `Requires`
  prefix, the source's binding, else the table's, else the aliases', else `undefined`.
- **`mcBranchOf`:**

```ts
export interface NothingSelected {
  /** The namespaces the element's Choices require, each once; a prefix nothing resolves has none. */
  namespaces: string[];
  value: Mce.AlternateContent;
}

export interface McOptions {
  /** The namespace URIs this reader understands; `UNDERSTOOD_NAMESPACES` by default. */
  understoodNamespaces?: Iterable<string>;
  /** Called when an element has Choices, none understood, and no Fallback. */
  onNothingSelected?(report: NothingSelected): void;
}
```

  A Choice is understood when its `Requires` names at least one prefix and every prefix resolves to a namespace in the set
  (REQ-302, REQ-304). Else the Fallback. Else `undefined`, with the report if the element has Choices.
- **`CONVENTIONAL_PREFIXES` went.** `mcBranchOf`'s and `McOptions`' comments are rewritten. Both said docx4j compares prefixes
  and `mcBranchOf` follows `McSelection`, which is true of its drawing path only: that path prefers the Fallback by default,
  and its text path resolves URIs.
- **Unchanged:** `NAMESPACE_PREFIXES` and everything a save writes, `walkAll` (every branch), and `test/lib/understood.mjs`,
  which is a consumer's judgement of what survives a resolved save and not this set.

## 5. What changes for a reader

| A Choice requiring | With | Before | After |
|---|---|---|---|
| A namespace not in the set | No Fallback | The first Choice | Nothing, reported |
| A table prefix outside the set (`cx5`, `dc`) | A Fallback | The Choice | The Fallback |
| A table prefix outside the set | No Fallback | The Choice | Nothing, reported |
| An understood namespace under another prefix (`q`) | A Fallback | The Fallback | The Choice |
| An understood namespace with no table entry (`c16r2`) | A Fallback | The Fallback | The Choice |

`textOf` changed with it (D1).

## 6. Results

- `textOf` against Word's saved text, over `ul-*`, `uc-*`, `c305-*` and `rw-25`: **142 of 142 match**. 118 did before: the 24
  were 14 `ul`, 6 `uc`, 3 `c305` and `rw-25`. Word showed the Choice in 118 and neither branch in 24, each of those reported.
- `npm run typecheck` and `npm test` pass on Node 18.20.8, 20.20.2 and 22.23.4, with jsonix 3.11.0.
- All 341 marshalled hashes are identical to `53e5a49`'s.
- Every `Requires` in the fidelity fixtures, Excel's `x15` and `x14` with no Fallback among them, resolves into the set, so
  no fixture's selection moved.

core-ts ran its suite in a scratch copy at its `1cae4ef`, both builds on jsonix 3.12.0, with `tree.mts:77` changed to
`mcBranchOf(value)`. Typecheck passes. Against `53e5a49`'s build, 8 more tests fail, all pinned to the old selection: 5 `PENDING`
req303 entries, the `extraction-probe-inspect` golden, and 2 `kept-branches` tests using `cx3` as not understood. Each stops at its
first failed assertion, so editing them may show more. All 119 `ul-*` files match Word through a kept and a resolved load.
`onNothingSelected` wasn't exercised there. docx4j-ts-editor has not been run against it.


## 7. Decisions

All five decided as recommended, 2026-11-07.

| D | Question | Decided |
|---|---|---|
| D1 | `textOf`'s output has been frozen since 2026-10-19, with a Word difference moving the core-ts reader rather than `textOf`. Does it follow `mcBranchOf` here, or keep the old selection privately? | Follow, as a stated exception. A private copy would keep returning words Word doesn't show, which is what REQ-305 changed to stop, and `textOf` and `mcBranchOf` would disagree. |
| D2 | Where does the warning surface? | A callback on `mcBranchOf` only. `textOf` takes no options (the same ruling), so it stays silent, and it isn't the text model REQ-801 binds. core-ts raises `mce.nothing_selected` with the part's name. A report on read for an add-in alone could follow if one asks. |
| D3 | `understood` held prefixes. Change its meaning, or rename it? | Rename to `understoodNamespaces` and drop `understood`. Prefixes passed to a URI option would silently select nothing; a rename fails core-ts's typecheck instead. |
| D4 | A prefix with no record and no table entry, as in a hand-built tree? | Not understood, and documented, as REQ-303 requires. The six URIs with no table entry get no prefix, since a read has its record. |
| D5 | The set is silent on spreadsheets and presentations. One default for every part? | One default, as `mcBranchOf` is WordprocessingML's helper. A caller can pass its own set, and no fixture's selection moves. |

## 8. Tests and docs

- **`test/word365-understood.mjs`, a ninth test:** the 142 `word/document.xml` under `test/fixtures/understood-probe/`, each
  beside Word's text, with a `SOURCE.md`. `textOf` must equal Word's, and `mcBranchOf` must report exactly where Word showed
  neither branch.
- **Smoke:** the CR-003 block's rule 3 and its two `understood` lines. A new block for `requiresNamespaces` (a source
  binding, the table, an alias, a prefix bound to another namespace than the table's, an unresolved prefix, a `deepCopy`, a
  `wml()` fragment, a hand-built tree), an empty `Requires`, `onNothingSelected` with and without a Fallback and at a root, and
  `UNDERSTOOD_NAMESPACES` held to docx4j's file when `../docx4j` is checked out.
- **The gate:** every fidelity hash identical.
- **Docs:** `CLAUDE.md`, the README's `mcBranchOf` entry, and CR-003 section 3.8's rule.

## 9. Consumers

- **core-ts:** `selectedBranchOf` (`src/model/content/tree.mts:77`) passes `understood: READ_UNDERSTOOD_PREFIXES`. The rename
  fails its typecheck until it changes, and the change can't land ahead of this release, since 0.7.0 has no such default, so it
  goes in the commit that raises its dependency. `mcBranchOf(value)` with no option uses the understood set. Its load's own sets, and when it adopts REQ-305, are its own work. Its readers change
  with D1, and it raises `mce.nothing_selected` from `onNothingSelected`. Its load must read through the facade for a prefix
  like `q` to resolve.
- **docx4j-ts-editor:** it calls `textOf`, so its text changes with D1. It doesn't call `mcBranchOf`. Not asked yet.
- **An add-in on this package alone:** `textOf` and `mcBranchOf` changed as section 5 has it.

## 10. Not in this CR

- **REQ-317, `mc:AlternateContent` as a main part's root:** core-ts CR-028's. `mcBranchOf` on such a root does select
  nothing after this CR, since it has no third step to take.
- **`mc:Ignorable`:** REQ-306 has the understood set decide it too, but this package carries it and doesn't act on it.
- **A second set here:** REQ-306's traversal set is for what a reader can read, and nothing in this package selects by it.
