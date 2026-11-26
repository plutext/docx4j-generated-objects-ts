# CR-024: A save's report paths found only for a report, in time linear in the part

**Status:** Implemented 2026-11-27, D1 and D2 as built (section 5), and committed at Jason's word ("make the commits then push"). Not
released: it goes out as a patch, 0.9.1, after core-ts's go on the pushed commit. core-ts and the editor verified the candidate
(section 4).
**Requested by:** Jason, through docx4j-ts-editor's ED-015 P4, accepted 2026-11-27 (editor `9d9efe6`): "The fix goes to
`@docx4j/generated-objects-ts` as a change request of its own (section 5, with the position cache as its second half), released as a
patch (0.9.1)". It was relayed by the editor's session the same day.
**Depends on:** nothing. No change to the runtime, the compiler or `modules/`.
**Counterparts:** the editor's ED-015, which depends on this CR (its S2 and S3). core-ts's binder keeps one path cache for its
`boundsOf`, as the editor notes. docx4j has no such report on save.

## 1. Summary

Every facade save that passes `onMisplaced` found the path of every container element it wrote, though only a report uses one, and
an ordinary document has none. Finding a path counted back over every earlier same-named sibling at each level, so the cost was
quadratic in the body's length. core-ts's `savePart` always passes `onMisplaced` (its CR-016 D2), so every save of every part paid it.
In Chromium, where each `previousSibling` is a call into the native DOM, the editor's 255-page save took 34.6 s, 93.5% of it there.

Now:
- **`onMisplaced`'s path is found at a container's first report,** not as the container is opened (D1). With nothing misplaced, no path
  is found.
- **One path memo per marshal serves all three reports** (`onMisplaced`, `onNormalised`, `onOutOfRange`; D2). It remembers each element's
  path and position. A position counts back only to the nearest same-named sibling already placed, and an ancestor's path is found
  once. A part with thousands of reports costs time linear in its size.

Every report and every saved byte is unchanged: no fidelity hash moved, and a stress save with 18,832 reports gives the same reports
and output as 0.9.0.

## 2. What was there

`pathOf` (`src/index.mts`, since CR-013's phase C, `9bd7289`) walked from an element to the root. At each step it counted the element's
earlier siblings of the same name, giving `/w:document/w:body[1]/w:p[512]/...`. Four places called it:
- **`reportingMisplaced`** (CR-017) called it in `onElement` for every container it opened (every paragraph, run, cell, row, table and
  content control), before testing the container's list. The path was used only in a report. For the k-th paragraph of a body it
  counted back over k − 1 paragraphs, and every run and cell in a paragraph counted back over them again, through its ancestors.
- **`normalising`** (CR-013) called it twice, and **`reportingOutOfRange`** (CR-016) once, per report. Each report counted back afresh, so
  a long document with a normalised value in every paragraph was quadratic again. The editor sets `normaliseOnSave` off and didn't
  meet that; an application normalising a long document would.

ED-015 section 3 has the profile: in Chromium 153, `pathOf` took 32.4 s of a 34.6 s save. In Node 24, with xmldom, it took 1.4 s of 3.0 s.

## 3. The change

`src/index.mts` only:
- **`pathsOf()`** replaces `pathOf`. It returns a path function holding two `WeakMap`s, an element's path and its position among
  same-named siblings, and is made afresh for each marshal (`marshalToDocument`). The marshal only appends: jsonix appends an element
  to its parent before `onElement` is called, and appends imported DOM in `writeNode` before it is reported. So a position, once found,
  stands for the rest of the marshal. The paths are the same strings as before.
- **`normalising`, `reportingOutOfRange` and `reportingMisplaced`** take the marshal's path function as a parameter.
- **In `reportingMisplaced`,** `base` is found at the first report (`base ??= pathOf(written)`). `testList` is synchronous and runs in
  `onElement`, before the container's children are written, so the path found then is the one found up front.

Nothing public changes: no export, option or type, and no report's content.

## 4. Results

The editor's `anon-250.docx` (255 pages, anonymised; its main part, 7,218 body blocks), read through the facade and saved unchanged in
Node 22.23.4, the two builds run back to back (0.9.0 from npm's tarball). Three runs each, in ms:

| Save | 0.9.0 | This CR |
|---|---|---|
| No callback | 614 to 657 | 629 to 684 |
| `onMisplaced`, nothing misplaced | 2,885 to 3,296 | 717 to 739 |
| `onNormalised`, nothing normalised | 676 to 718 | 659 to 765 |
| All three callbacks, nothing reported | 3,144 to 3,384 | 878 to 903 |

A stress save of the same part with a report of each kind across the body took 9,210 ms on 0.9.0 and 1,182 ms here. It had a run
straight in the body after every paragraph, a table in every paragraph in a cell (9,178 misplaced), an empty `w:rsidR` on every
paragraph (9,178 normalised) and a `w:sz` of 70,000 in every run property with one (476 out of range). The reports, in order, and the
output's SHA-256 are identical to 0.9.0's.

The editor measured the candidate (sha1 `b59f42f1`, this change on `54c9966`) on 2026-11-27, built into the editor over it (ED-015
section 7.4, editor `d4ddd84`). Three runs each, in ms:

| The editor's benchmark | 0.9.0 | The candidate | Its target |
|---|---|---|---|
| Chromium 153, `saveEditedMs` (the export) | 32,198 | 1,408 to 1,817 | 2,000 |
| Chromium 153, `saveEditedCheckedMs` (with its self-check) | 35,997 | 3,748 to 3,928 | 4,000 |
| Node 24, the same two | 2,773 and 5,625 | 1,249 and 3,402 | |

Its editor-model suite (87 files, 1,672 passed, 8 skipped, every save self-checked) is the same over the candidate as over 0.9.0. Its
S3 measures the release again.

The 341 fidelity parts' marshalled hashes are identical to 0.9.0's published tarball.

**core-ts on the candidate,** 2026-11-27, Node 22: identical to npm's 0.9.0 in everything it compared.
- **The tarball:** only `dist/index.mjs` differs from 0.9.0's, so no declaration changes.
- **Its suite:** typecheck clean, and 2,077 of 2,077.
- **Its corpus:** 1,716 files, read and saved twice on each version: at the defaults, and with its checks off, recording every report.
  The files are every package under its `test/`, docx4j's `charts/`, its check 72's 18 and `anon-250`, plain and with all three
  kinds of report injected. The two versions' dumps are byte-identical: 42,117 saved parts, 33 refused saves with 1,050 findings,
  and 1,050 misplaced, 3,975 out-of-range and 4,272 normalised reports, each with its path, in order. The paths reach `w:p[7088]`,
  1,400 of them inside cells or controls.
- **Not checked:** a browser, Node 18 and 20, and the published tarball. It compares that after the release, before raising its
  range at Jason's word.

## 5. Decisions

| | Question | Decision |
|---|---|---|
| D1 | `onMisplaced`'s path: when? | At the first report of each container, as ED-015 section 5 proposes: every report the same, a container with nothing misplaced finding none |
| D2 | The other three reports' paths | A memo per marshal, shared by all three reports, ED-015's "second half": linear, where a cache per report kind would find an ancestor's path once per kind. Per marshal, not per module, so nothing outlives the DOM it describes |

## 6. Tests and docs

- **`test/smoke.mjs`, CR-024's block.** It counts, not times. While it runs, every element the marshal creates reads its
  `previousSibling` through a counter (xmldom's `createElementNS` wrapped, and restored after), so the counts are exact on any machine.
  - **Nothing misplaced:** a save with `onMisplaced` reads exactly what a save without it reads.
  - **A report of each kind for every paragraph:** 200 and 400 paragraphs. Every path is the expected one, and the reads double:
    2,402 against 4,802, where 0.9.0's code reads 42,003 against 164,003.
- **Mutants** of `dist/index.mjs`, each caught (scratchpad `cr024/mutants.py`):
  - **By the new block:** the eager `base`; a position counting every sibling; a position not remembered.
  - **By CR-013's path checks:** a remembered position off by one; an ancestor's path taken for the element's.
- **Unchanged:** the full suite, and the 341 fidelity hashes against 0.9.0's tarball.
- **Docs:** `CLAUDE.md`'s smoke description and the CR index.

## 7. Consumers

- **docx4j-ts-editor:** ED-015's S2 takes 0.9.1 by `npm update`, since core-ts 0.9.0 depends on `^0.9.0`. Its S3 runs `saveEditedMs` and
  `saveEditedCheckedMs` against it.
- **core-ts:** no change. Its `savePart` passes `onMisplaced` on every save and gets the time back. Its go on the pushed commit comes
  before the release, as for every release.

## 8. Not in this CR

- **The read side.** A report on read carries no path: `path` is set on save only (`OutOfRange`, `MisplacedChild`).
- **A path cache across marshals,** for an application saving one tree many times. A marshal writes a new DOM each time, so there is
  nothing to carry.
- **A browser test here.** Every test in this package runs on xmldom. The browser figures are the editor's (section 4).
