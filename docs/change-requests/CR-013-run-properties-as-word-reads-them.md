# CR-013: Run properties read as Word reads them (`readRunOptions`, a colour table, the facade's part of the parse)

**Status:** Proposed 2026-10-24, at core-ts's request; not scheduled, implementation waits for Jason
**Requested by:** core-ts (session docx4j-core-ts-5b), 2026-10-24, for its CR-015 section 6, which the portfolio
session asked it to raise here. core-ts's CR was drafted at Jason's word ("yes please", to one CR each in docx4j,
core-ts and docx4j-python against the same rules and probes).
**Depends on:** the parse half of core-ts CR-015 section 6, which core-ts raised with jsonix-d2 on 2026-10-24 and
which isn't filed yet: an `ST_OnOff` value read by Word's rule, and a value outside its enumeration kept loading.
Sections 3.1, 3.2, 3.5 and 3.6 depend on its design. Sections 3.3 and 3.4 depend on nothing.
**Counterpart:** docx4j CR-060 (proposed, 38ccaedeb), the same rules in Java; core-ts CR-015 and docx4j-python's
CR are drafted in parallel.

## 1. Summary

Word 365 is lenient where this package is strict. Jason's probes of 2026-10-24 found three rules. The probes, his
PDFs and Word's saves are in `~/vbShares/Office 2016/onoff-fun/`; core-ts CR-015 section 2 and docx4j CR-060
section 1 give the rules in full.
- **On/off:** off only for exactly `false`, `off` or `0`; anything else is on (`yes`, `FALSE`, `""`, `" false"`).
- **Enumerations** (`w:u`, `w:highlight`, `w:vertAlign`, `w:em`): a value outside the list, case-exact, is unset.
- **`w:color`:** trimmed; six hex digits in any case, with an optional `#`; or a classic HTML colour name,
  case-insensitive. `auto` in any case, and anything else, is automatic.

This package's part:
- `builders/wml`'s `readRunOptions`, which core-ts's `Font`, `toApiScript` and `spanEdits` read through, reads
  `w:highlight` and `w:color` raw today. It should read them as Word does, and do the same for the on/off and
  enumeration values once they load.
- A colour table that core-ts can share, rather than keep a copy.
- Whatever the facade needs so that the runtime's leniency applies to WordprocessingML only.

## 2. What this package does (measured 2026-10-24, at 1e4e895)

Each value was unmarshalled as a run's `w:rPr` through `unmarshalString`, then read with `readRunOptions`:
- **The whole part fails** for an on/off value other than `true`, `false`, `1`, `0`, `on` or `off` (`yes`, `FALSE`,
  `" false"`, `""`). It also fails for any `w:u` or `w:vertAlign` value outside its enumeration (`Single`, `bogus`,
  `Superscript`). The runtime refuses each ("Either [true], [1], [on], [false], [0] or [off] expected", "Value
  [Single] is invalid for the enum type").
- **Read raw:** `w:highlight` and `w:color` are strings in the model, so they load.
  - `w:highlight` `YELLOW` and `bogus` read as `"YELLOW"` and `"bogus"`, where Word shows none. `yellow` reads
    `#FFFF00`.
  - `w:color` `red`, `#ff0000`, `" FF0000 "`, `AUTO` and `F00` each read as their own string. Word reads them as
    `#FF0000`, `#FF0000`, `#FF0000`, automatic and automatic. Only bare six-digit hex (`ff0000`) reads `#FF0000`.
- **Already as Word reads it:** every valid value; bare means on; `off` is off; a bare `<w:u/>` reads `None`.
  `readRunOptions`' `on(v)` is `v !== undefined && v.val !== false`, so it is right for any boolean the parse gives.

## 3. Proposal

### 3.1 On/off

No change here if the parse gives Word's boolean (core-ts CR-015's D1 (a)): `on()` is already right, and so is every
reader in core-ts. If the runtime keeps the source string instead, `on()` takes Word's rule: off only for `false`,
`"false"`, `"off"` or `"0"`, exactly.

### 3.2 Enumerations

If the runtime drops a value outside the enumeration, leaving the attribute absent, `readRunOptions` already reads
it as Word does: `w:u` reads `None`, and `w:vertAlign` reads neither superscript nor subscript. docx4j CR-060's D3
("read as absent") does the same. The generated declarations also stay the literal unions they are. That is the
form this package would ask for.

If the runtime keeps the string, `readRunOptions` changes in one place: underline's fallback becomes `None`, not
`Mixed`. Every valid `ST_Underline` value has an Office JS name (`UNDERLINE_TO_WML` covers all seventeen besides
`none`), so the fallback is reached only by an invalid value. `vertAlign` already compares exactly.

### 3.3 `w:highlight`

A value outside `ST_HighlightColor`, case-exact, reads as no highlight: `highlightHexValue(h) ?? null`, not
`?? h`. `highlightHexValue` already matches case-exactly. Word repairs such a file, dropping the element, so this
is what Word shows after the repair. It needs no runtime change, since the model types the value as a string.

### 3.4 `w:color`, and one table for both TypeScript ports

`helpers/wml` exports the table and a reader:
- **`WORD_COLOR_NAMES`:** the classic HTML (Netscape) colour names Word reads, with their values. That is CSS
  Color 3's 147 keywords, less `grey`, `darkgrey`, `darkslategrey`, `dimgrey`, `lightslategrey`, `slategrey` (Word
  spells grays with an a, except `lightgrey`) and `lightgray`: 140 names. The names measured in Word's saves (41 by
  core-ts's count) are held to Word's values by tests, and the rest come from CSS. It is not DrawingML's preset
  table (core-ts CR-008's `PRESET`), which spells and values some names differently.
- **`wordColor(value)`:** returns `'#RRGGBB'`, `'auto'` or `undefined`. It trims the value. `auto` in any case
  gives `'auto'`. An optional `#` and six hex digits in any case give upper-case `#RRGGBB`. A table name,
  case-insensitive, gives its value. Anything else, a value Word reads as automatic but that names no colour,
  gives `undefined`. That leaves the caller to choose between "automatic" and "as if absent" (O1, section 5).

`readRunOptions`' `color` then reads `''` when `w:color` is absent, as today, and otherwise
`wordColor(val) ?? <O1>` (section 5). core-ts's `ResolvedColor.ofHex` and `WordColors` can call `wordColor` instead
of keeping a second table. If docx4j adds the same table (CR-060's D5), the test can compare the two name lists.

While there: `RunFormatting.color`'s comment says "or a theme name", but `readRunOptions` reads only `w:color/@w:val`.
The comment is corrected.

### 3.5 The facade's part of the parse

This depends on the runtime's design, which is jsonix-d2's:
- **(a) A global rule** (Word's on/off for every `Boolean`, every enumeration lenient): nothing changes here.
- **(b) Per mapping:** the WordprocessingML mappings (`org_docx4j_wml` and CR-009's `org_docx4j_wml_globals`) are
  marked lenient. That could be done by the compiler, or by the facade before it builds its context.
- **(c) Per call:** an unmarshaller option. The facade sets it in `unmarshalPackage` by content type, and in
  `unmarshalString`/`unmarshalNode` by the root's namespace. core-ts passes it itself through CR-007's per-call
  options where it unmarshals a part.

This package's preference is (b). A WordprocessingML part also holds DrawingML (`wp:`, `a:`, `pic:`) and VML,
whose booleans Word's rule isn't measured for. Per-part scoping, as in (c), would apply it to them too.

### 3.6 What a save writes (core-ts CR-015's D2, Jason's)

The rule is that a save never changes meaning:
- `applyRunOptions` already writes only valid forms.
- An on/off or enumeration value read leniently is written in the runtime's form, whatever its design: as
  `"true"`/`"false"`, and an invalid enumeration value dropped.
- **If D2 (a):** an element left with no required value is written bare. That is `w:vertAlign` or `w:em` whose
  invalid value the parse dropped, and `w:highlight` whose value is missing or outside its list. Word repairs a
  file with one, dropping the element. The facade's marshal can drop them first. That is a tree rule, so it would
  live here, and every consumer gets it. Office never writes such an element, so the fidelity corpus can't move.

## 4. Not proposed

- **Word's own save form:** on bare, off dropped or `0`, names and lower-case hex as upper-case hex. That is
  core-ts's D2 (c), possible later; neither changes meaning.
- **Office JS's own answers** for these values (core-ts's D4, its check 60), and rendering.
- **The parse itself:** that's the runtime's, or the compiler's under (b).

## 5. Decisions (for Jason)

- **O1. An unusable `w:color` in `readRunOptions`:** `''`, as if not set, or `'auto'`, which is what Word shows with
  no style. Whether such a value overrides a style is core-ts's D5, a round 5 probe. **Recommendation:** `''`
  until that probe, as core-ts reads it until then, then follow it.
- **O2. The colour table's home:** `helpers/wml`, exported, with core-ts calling `wordColor` rather than keeping its
  own table. **Recommendation:** yes.
- **O3. Dropping valueless `w:highlight`, `w:vertAlign` and `w:em` on marshal**, if D2 (a): in the facade, or in
  core-ts. **Recommendation:** the facade.
- **The scoping (section 3.5)** is the runtime's design. This package's preference is per mapping.

## 6. Tests

Smoke, a block of its own:
- **Without the runtime:**
  - `readRunOptions` over the probe values that load today: `w:highlight` `YELLOW` and `bogus` read null.
  - For `w:color`: hex with `#`, lower case, surrounding whitespace, `auto`/`Auto`/`AUTO`, the measured names
    (`red`, `darkGray`, `lightgrey`, `lightgoldenrodyellow`, ...), and the non-names (`grey`, `lightGray`,
    `darkYellow`, `rebeccapurple`, `"dark red"`, `F00`, `0xFF0000`, `GG0000`, `""`) read per O1.
  - `wordColor` itself.
  - The table's 140 names, and the seven excluded spellings missing from it.
- **With the runtime:** each on/off value of core-ts CR-015's rule 1 reads as Word does. So does each invalid
  `w:u`/`w:vertAlign`/`w:em`, through unmarshal and `readRunOptions`. Until the runtime lands, each such value is
  asserted to fail the part, with the upstream commit cited, so that the fix announces itself.
- **The probes:** the marker paragraphs' `document.xml` from `onoff-fun/` are extracted as fixtures, with a
  `SOURCE.md` naming each file. Extracted part XML only, never the archives, and never edited by hand. Each marker
  then reads as Word read it.
- **If O3:** a marshalled valueless `w:vertAlign` is dropped, and the fidelity corpus's 341 hashes don't move.

## 7. Consumers

- **core-ts:** its `Font`, `toApiScript` and `spanEdits` read through `readRunOptions`, and its CR-015 depends on this.
  With O2 it drops its own colour table in favour of `wordColor`.
- **Office JS add-ins using this package alone:** the same readings, through `readRunOptions`.

## 8. Phases and release

- **Phase A**, with no dependency: sections 3.3 and 3.4, with O1 and O2. Half a day.
- **Phase B**, once the runtime lands: section 3.1 or 3.2 (whichever its design needs), section 3.5, and the probe
  fixtures. Half a day, plus a regeneration if the compiler marks the mappings.
- **Phase C**, if D2 (a) and O3: section 3.6's drop on marshal. Two hours.

Each phase is a patch, as 0.3.1 and 0.4.1 were: a fix to readings, an additive export, possibly a raised runtime
floor. No API is removed.
