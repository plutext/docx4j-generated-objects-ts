# CR-013: Run properties read as Word reads them (`readRunOptions`, a colour table, the facade's part of the parse)

**Status:** Proposed 2026-10-24, at core-ts's request; not scheduled, implementation waits for Jason
**Requested by:** core-ts (session docx4j-core-ts-5b), 2026-10-24, for its CR-015 section 6, which the portfolio
session asked it to raise here. core-ts's CR was drafted at Jason's word ("yes please", to one CR each in docx4j,
core-ts and docx4j-python against the same rules and probes).
**Depends on:** jsonix-CR-010 for sections 3.1, 3.2, 3.5 and 3.6. Jason decided its J1 to J3 as recommended. It is
implemented and pushed (jsonix master 5df33ad, CI green, 2026-10-24) but not released; npm's latest is 3.6.0, so
phase B waits for the release. It hands an attribute value its type refuses to a per-call `onInvalidValue` instead of failing the part,
and a value its type can't print to `onUnprintableValue`. The facade sets both, with Word's rules. Sections 3.3 and
3.4 depend on nothing.
**Counterpart:** docx4j CR-060, the same rules in Java, implemented through phase 3 on `VERSION_17_3_5`
(868a1797b, unpushed, 2026-10-24):
- an `OnOffAdapter` on 46 of `wml.xsd`'s 48 boolean attributes, as annox annotations. A regeneration from that xsd
  is byte-identical to this package's `modules/`, checked 2026-10-24;
- a value outside an enumeration read as absent;
- marshal-only adapters that leave out a `w:highlight`, `w:vertAlign` or `w:em` whose value is missing or invalid.

core-ts CR-015 and docx4j-python's CR are drafted in parallel.

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
- jsonix-CR-010's two callbacks, set by the facade with Word's rules, for WordprocessingML only.

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

The facade's `onInvalidValue` (section 3.5) returns `true` for a refused `Boolean` in `org_docx4j_wml`. The
runtime already reads `false`, `0` and `off` as false, and refuses only what Word reads as on (`yes`, `FALSE`,
`" false"`, `""`). The model then holds Word's boolean, so `on()` is already right, and so is every reader in
core-ts (core-ts CR-015's D1 (a)). A save writes `true`, which Word reads the same way, as jsonix-CR-007's `off`
is saved `false`.

### 3.2 Enumerations

The facade's `onInvalidValue` either keeps a refused enumeration value as its string, or drops it, leaving the
attribute absent. Which one is core-ts CR-015's D1/D2, and jsonix-CR-010 section 5.4's question, for Jason:
- **Drop:** `readRunOptions` already reads it as Word does: `w:u` reads `None`, and `w:vertAlign` reads neither
  superscript nor subscript. docx4j CR-060's D3 ("read as absent") does the same, and the model stays inside its
  declared types. But `<w:u w:val="Single"/>` saves as `<w:u/>`. Over a style that underlines, that shows the
  style's underline (Word's S11), which changes meaning if core-ts's D5 finds that an invalid value overrides a
  style.
- **Keep** (jsonix-d2's recommendation until D5 is measured): a save writes the source's spelling back, through
  the facade's `onUnprintableValue`, so it means to Word what the source meant, whatever D5 finds. The cost is a
  typed property holding a string outside its declared union, such as `'Single'` where the declaration says
  `UnderlineEnumeration`. Every reader must then treat a non-member as unset. A TypeScript `switch` over the union
  meets a value it can't name, and the declarations don't say so.

`readRunOptions` changes in one place either way: underline's fallback becomes `None`, not `Mixed`. Every valid
`ST_Underline` value has an Office JS name (`UNDERLINE_TO_WML` covers all seventeen besides `none`), so the
fallback is reached only by an invalid value. `vertAlign` already compares exactly.

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

### 3.5 The facade sets jsonix-CR-010's callbacks

Every facade unmarshal (`unmarshalString`, `unmarshalNode`, `unmarshalPackage`) passes an `onInvalidValue`. Every
facade marshal passes an `onUnprintableValue`, if the enumerations are kept. Each is scoped by
`classInfo.module.name === 'org_docx4j_wml'`:
- **On/off:** a refused `Boolean` gives `true` (section 3.1).
- **Enumerations:** a refused value is kept or dropped (section 3.2). On marshal, a kept string is written back.
- **Anything else rethrows,** so it fails as today. That covers other simple types in wml, and every type of
  every other mapping, such as DrawingML (`wp:`, `a:`, `pic:`), VML and w14 inside a `document.xml`, which Word's
  rule isn't measured for.

The rule covers every `Boolean` in `org_docx4j_wml`, including the two that docx4j's `OnOffAdapter` leaves strict:
`w:writingStyle/@w:checkStyle` and `w:readModeInkLockDown/@w:actualPg`, both in the settings part. docx4j leaves
them out for a Java reason: they are required, so they bind to a primitive `boolean`, which an adapter can't bind.
They are `ST_OnOff` in the spec, and Word's reading of them is unmeasured in either port. Applying Word's rule
keeps `settings.xml` loading where docx4j would refuse it, so this is a known difference between the ports rather
than one to match.

Scoping by the class's mapping, not by the part, needs no per-part choice. A `document.xml` gets Word's rule for
its WordprocessingML and today's strictness for the DrawingML it holds. CR-009's `org_docx4j_wml_globals`
declares elements of `org_docx4j_wml`'s types, so their class infos answer `org_docx4j_wml` too.

A caller's own `onInvalidValue` or `onUnprintableValue`, passed through CR-007's per-call options, takes precedence.
The facade's are exported (`wordInvalidValue`, `wordUnprintableValue`) for a caller that wants to compose them.

Word's enumeration rule is measured for `w:u`, `w:vertAlign` and `w:em` only, and `w:highlight` is a string here.
Whether every wml enumeration is lenient, or only those three, is O4.

### 3.6 What a save writes (core-ts CR-015's D2, Jason's)

The rule is that a save never changes meaning:
- `applyRunOptions` already writes only valid forms.
- An on/off value read leniently is written `"true"`.
- An invalid enumeration value is written back as read if kept, or not written if dropped (section 3.2).
- **If D2 (a):** Word repairs a file with a `w:vertAlign` or `w:em` whose value is missing or invalid, or a
  `w:highlight` whose value is missing or outside its list, by dropping the element. The facade's marshal can drop
  them first. That is a tree rule, so it would live here, and every consumer gets it. Office never writes such an
  element, so the fidelity corpus can't move.

## 4. Not proposed

- **Word's own save form:** on bare, off dropped or `0`, names and lower-case hex as upper-case hex. That is
  core-ts's D2 (c), possible later; neither changes meaning.
- **Office JS's own answers** for these values (core-ts's D4, its check 60), and rendering.
- **The callback mechanism itself:** that's jsonix-CR-010's.
- **Widening the generated enumeration declarations** to admit a kept string. That's a compiler question, if
  keeping is chosen and a consumer asks.

## 5. Decisions (for Jason)

- **O1. An unusable `w:color` in `readRunOptions`:** `''`, as if not set, or `'auto'`, which is what Word shows with
  no style. Whether such a value overrides a style is core-ts's D5, a round 5 probe. **Recommendation:** `''`
  until that probe, as core-ts reads it until then, then follow it.
- **O2. The colour table's home:** `helpers/wml`, exported, with core-ts calling `wordColor` rather than keeping its
  own table. **Recommendation:** yes.
- **O3. Dropping valueless `w:highlight`, `w:vertAlign` and `w:em` on marshal**, if D2 (a): in the facade, or in
  core-ts. **Recommendation:** the facade.
- **O4. Which wml enumerations are lenient:** every one (a refused value anywhere in `org_docx4j_wml` loads), or
  only `w:u`, `w:vertAlign` and `w:em`, the ones Word was measured for. Refusing the part is the worse half
  (core-ts CR-015 section 1). But every reader of another enumeration (`w:fldChar/@w:fldCharType`, say) must then
  treat a non-member as unset. **Recommendation:** every one, read as unset, since Word opened every measured case
  and a part that fails can't be read at all.
- **Keep or drop** an invalid enumeration value is core-ts's D1/D2 and jsonix-CR-010's section 5.4, not decided
  here. docx4j reads it as absent, which is the same as dropping. Keeping would make the ports differ on save for
  `w:u`, whose source spelling this package would write back where docx4j writes a bare `<w:u/>`.

## 6. Tests

Smoke, a block of its own:
- **Without the runtime:**
  - `readRunOptions` over the probe values that load today: `w:highlight` `YELLOW` and `bogus` read null.
  - For `w:color`: hex with `#`, lower case, surrounding whitespace, `auto`/`Auto`/`AUTO`, the measured names
    (`red`, `darkGray`, `lightgrey`, `lightgoldenrodyellow`, ...), and the non-names (`grey`, `lightGray`,
    `darkYellow`, `rebeccapurple`, `"dark red"`, `F00`, `0xFF0000`, `GG0000`, `""`) read per O1.
  - `wordColor` itself.
  - The table's 140 names, and the seven excluded spellings missing from it.
- **With the runtime:**
  - Each on/off value of core-ts CR-015's rule 1 reads as Word does, and so does each invalid
    `w:u`/`w:vertAlign`/`w:em`, through unmarshal and `readRunOptions`.
  - The same values in a DrawingML attribute of the same part still fail.
  - A caller's own callback takes precedence.
  - If kept, the value saves as read; if dropped, the attribute goes.
  - Until jsonix-CR-010 is released, each such value is asserted to fail the part, with the upstream commit cited,
    so that the fix announces itself.
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
- **Phase B**, once jsonix-CR-010 is released: sections 3.1, 3.2 and 3.5, with O4 and the keep-or-drop decision,
  and the probe fixtures, with the runtime floor raised to that release. Half a day.
- **Phase C**, if D2 (a) and O3: section 3.6's drop on marshal. Two hours.

Each phase is a patch, as 0.3.1 and 0.4.1 were: a fix to readings, an additive export, possibly a raised runtime
floor. No API is removed.
