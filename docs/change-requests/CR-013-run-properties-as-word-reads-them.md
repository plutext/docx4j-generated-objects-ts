# CR-013: Run properties read as Word reads them (`readRunOptions`, a colour table, the facade's part of the parse)

**Status:** Proposed 2026-10-24, at core-ts's request; not scheduled, implementation waits for Jason
**Requested by:** core-ts (session docx4j-core-ts-5b), 2026-10-24, for its CR-015 section 6, which the portfolio
session asked it to raise here. core-ts's CR was drafted at Jason's word ("yes please", to one CR each in docx4j,
core-ts and docx4j-python against the same rules and probes).
**Depends on:** jsonix-CR-010 for sections 3.1, 3.2, 3.5 and 3.6. It hands an attribute value its type refuses to
a per-call `onInvalidValue` instead of failing the part, and a value its type can't print to `onUnprintableValue`;
the facade sets them with Word's rules. Jason decided its J1 to J3 as recommended, and it was released in
`@docx4j/jsonix` 3.7.0 (tag d4e8af8, 2026-10-24), so phase B raises the floor to ^3.7.0. Sections 3.3 and 3.4
depend on nothing.
**Counterpart:** docx4j CR-060, the same rules in Java, implemented through phase 3 on `VERSION_17_3_5`
(868a1797b, unpushed, 2026-10-24):
- an `OnOffAdapter` on 46 of `wml.xsd`'s 48 boolean attributes, as annox annotations. A regeneration from that xsd
  is byte-identical to this package's `modules/`, checked 2026-10-24;
- a value outside an enumeration read as absent;
- marshal-only adapters that leave out a `w:highlight`, `w:vertAlign` or `w:em` whose value is missing or invalid.

core-ts CR-015 and docx4j-python's CR are drafted in parallel.

## 1. Summary

Word 365 is lenient where this package is strict. Jason's six rounds of probes, all on 2026-10-24, found these
rules. The probes, his PDFs and Word's saves are in `~/vbShares/Office 2016/onoff-fun/`; core-ts CR-015 section 2
and docx4j CR-060 sections 1, 12 and 13 give them in full.
- **On/off:** off only for exactly `false`, `off` or `0`; anything else is on (`yes`, `FALSE`, `""`, `" false"`).
- **Enumerations** (`w:u`, `w:highlight`, `w:vertAlign`, `w:em`): a value outside the list, case-exact, is as if
  absent, so a style's value shows through it (round 5).
  - The exception: a `w:u` whose value is exactly `true`, `false`, `on`, `off`, `1` or `0` reads as `none`, and
    removes a style's underline (round 6). `True`, `TRUE`, `yes` and `no` are invalid like any other value.
- **`w:color`:** six hex digits in any case, with an optional `#`, surrounding whitespace trimmed; or a classic HTML
  colour name, case-insensitive. Only exactly `auto` is automatic and overrides a style's colour (round 6). `Auto`,
  `AUTO`, `" auto"` and anything else are as if absent (round 5).
  - Where an unusable colour sits matters (round 6). It is as if absent in `w:color` and `w:u/@w:color`, and
    automatic, overriding the style, in `w:shd/@w:fill` and a border's `w:color`.

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
  - `w:color` `red`, `#ff0000`, `" FF0000 "`, `AUTO` and `F00` each read as their own string. Word reads the
    first three as `#FF0000`, and the last two as absent (automatic where no style gives a colour). Only bare
    six-digit hex (`ff0000`) reads `#FF0000`.
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

The facade's `onInvalidValue` drops a refused enumeration value, leaving the attribute absent. Round 5 showed
that this keeps meaning: an invalid value over a style is as if absent, and the style's value shows. The
exception is `w:u`, whose six on/off words (`true`, `false`, `on`, `off`, `1`, `0`, exactly) give `'none'`, a member
of the enumeration, since Word reads them as none (round 6). So:
- `readRunOptions` already reads the result as Word does: `w:u` reads `None`, and `w:vertAlign` reads neither
  superscript nor subscript.
- The model stays inside its declared types, so readers never meet a string outside a union, and no
  `onUnprintableValue` is needed.
- docx4j CR-060 reads invalid values as absent (its D3), so the ports agree on save. That includes `w:u true`,
  which Word itself saves as `none`.

Dropping is core-ts CR-015's D2 (a), recommended there and confirmed by round 5; it is Jason's decision. The
alternative is keeping each invalid value's spelling and writing it back through `onUnprintableValue` (D2 (b),
jsonix-CR-010 section 5.4). That would put strings outside the declared unions into the model, and make the ports
differ on save.

`readRunOptions` also changes underline's fallback to `None`, not `Mixed`, for an object built by hand with an
invalid value. Every valid `ST_Underline` value has an Office JS name (`UNDERLINE_TO_WML` covers all seventeen
besides `none`), so the fallback is reached only by an invalid value. `vertAlign` already compares exactly.

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
- **`wordColor(value)`:** returns `'#RRGGBB'`, `'auto'` or `undefined`.
  - Exactly `auto` gives `'auto'`, untrimmed and case-sensitive (round 6).
  - Six hex digits in any case, with an optional `#` and surrounding whitespace trimmed, give upper-case `#RRGGBB`.
  - A table name, case-insensitive, gives its value. Trimming is measured for hex only, so names aren't trimmed.
  - Anything else gives `undefined`. Word reads such a value as absent in `w:color` and `w:u/@w:color`, but as
    automatic in a fill or a border colour (round 6), so the caller applies its own place's rule.

`readRunOptions`' `color` then reads `''` when `w:color` is absent or unusable, and otherwise `wordColor(val)` (O1,
section 5). core-ts's `ResolvedColor.ofHex` and `WordColors` can call `wordColor` instead of keeping a second table.
docx4j has the same table as `ColorTables.HTML` (CR-060's D5), and a test holds the two name lists equal.

While there: `RunFormatting.color`'s comment says "or a theme name", but `readRunOptions` reads only `w:color/@w:val`.
The comment is corrected.

### 3.5 The facade sets jsonix-CR-010's callbacks

Every facade unmarshal (`unmarshalString`, `unmarshalNode`, `unmarshalPackage`) passes an `onInvalidValue`. A facade
marshal passes an `onUnprintableValue` only if Jason chooses to keep invalid values (section 3.2). Each is scoped
by `classInfo.module.name === 'org_docx4j_wml'`:
- **On/off:** a refused `Boolean` gives `true` (section 3.1).
- **Enumerations:** a refused value is dropped, but for `w:u`'s six on/off words, which give `'none'`
  (section 3.2).
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
Whether every wml enumeration is lenient, or only those three, is O4. Whether Word's own extension mappings join
`org_docx4j_wml` in the lenient set is O5.

### 3.6 What a save writes (core-ts CR-015's D2, Jason's)

The rule is that a save never changes meaning:
- `applyRunOptions` already writes only valid forms.
- An on/off value read leniently is written `"true"`.
- An invalid enumeration value is not written (dropped, section 3.2). `w:u`'s on/off words are written `none`, as
  Word saves them.
- **If D2 (a):** Word repairs a file with a `w:vertAlign` or `w:em` whose value is missing or invalid, or a
  `w:highlight` whose value is missing or outside its list, by dropping the element. The facade's marshal can drop
  them first. That is a tree rule, so it would live here, and every consumer gets it. Office never writes such an
  element, so the fidelity corpus can't move.

## 4. Not proposed

- **Word's own save form:** on bare, off dropped or `0`, names and lower-case hex as upper-case hex. That is
  core-ts's D2 (c), possible later; neither changes meaning.
- **Office JS's own answers** for these values (core-ts's D4, its check 60), and rendering.
- **The callback mechanism itself:** that's jsonix-CR-010's.
- **Widening the generated enumeration declarations** to admit a kept string. Dropping makes it unnecessary.
- **A save normaliser** for colours. If one is added later: an unusable `w:color/@w:val` beside a `w:themeColor`
  must be written `auto`, not dropped, since `val` is required in `CT_Color`. Without a theme it is dropped
  (docx4j-python's review, through core-ts).
- **A character style's highlight,** which Word doesn't draw (round 6). That is style resolution, which belongs
  to core-ts and docx4j; `readRunOptions` reads direct formatting only.

## 5. Decisions (for Jason)

- **O1. An unusable `w:color` in `readRunOptions`:** `''`, as if not set, or `'auto'`. Rounds 5 and 6 answered it:
  over a style's colour, an unusable value is as if absent and the style's colour shows. Only exactly `auto`
  overrides. **Recommendation:** `''`.
- **O2. The colour table's home:** `helpers/wml`, exported, with core-ts calling `wordColor` rather than keeping its
  own table. **Recommendation:** yes.
- **O3. Dropping valueless `w:highlight`, `w:vertAlign` and `w:em` on marshal**, if D2 (a): in the facade, or in
  core-ts. **Recommendation:** the facade.
- **O4. Which wml enumerations are lenient:** every one (a refused value anywhere in `org_docx4j_wml` loads), or
  only `w:u`, `w:vertAlign` and `w:em`, the ones Word was measured for. Refusing the part is the worse half
  (core-ts CR-015 section 1). But every reader of another enumeration (`w:fldChar/@w:fldCharType`, say) must then
  treat a non-member as unset. With dropping, no reader meets one: the attribute is absent. **Recommendation:**
  every one, since Word opened every measured case and a part that fails can't be read at all.
- **O5. Word's own extension mappings in the lenient set** (core-ts CR-015's D10, proposed 2026-10-24): `org_docx4j_w14`,
  `w15`, `w15symex`, `w16`, `w16cex` and `w16cid` alongside `org_docx4j_wml`. Scoped to `org_docx4j_wml` alone, a
  `w14:checked w:val="TRUE"` still fails its part, where docx4j-python reads it by Word's rule. Neither round measured
  w14. In this package, only w14 has on/off attributes:
  - `CT_OnOff.val`, a `Boolean` (`w14:checkbox/w14:checked`, `w14:conflictMode`, `w14:discardImageEditingData`).
    It takes the `Boolean` rule.
  - Two on/off values typed as an anonymous enumeration of `true`, `false`, `0` and `1`:
    `CT_LinearShadeProperties/@scaled`, and `CT_StyleSet/@val` (`w14:stylisticSets`). Under the enumeration rule
    they would be dropped. The `Boolean` rule mapped onto the members would read them as Word reads on/off: `'false'`
    for exactly `false`, `off` or `0`, else `'true'`.
  - w14's 13 other enumerations (ligatures, number forms, text effects) would become lenient too.
  The other five mappings have no booleans and one enumeration between them. DrawingML, VML, SpreadsheetML and
  PresentationML stay strict. **Recommendation:** with D10, if Jason agrees, and with the on/off enumerations read by
  the `Boolean` rule.
- **Keep or drop** an invalid enumeration value is core-ts CR-015's D2, Jason's. This CR assumes D2 (a), drop,
  which core-ts recommends and round 5 confirmed keeps meaning (section 3.2).

## 6. Tests

Smoke, a block of its own:
- **Without the runtime:**
  - `readRunOptions` over the probe values that load today: `w:highlight` `YELLOW` and `bogus` read null.
  - For `w:color`: hex with `#`, lower case and surrounding whitespace; `auto` (`'auto'`); `Auto`, `AUTO`,
    `" auto"` and `#auto` (`''`); the measured names (`red`, `darkGray`, `lightgrey`, `lightgoldenrodyellow`,
    ...); and the non-names (`grey`, `lightGray`, `darkYellow`, `rebeccapurple`, `"dark red"`, `F00`,
    `0xFF0000`, `GG0000`, `""`), which read `''`.
  - `wordColor` itself.
  - The table's 140 names, and the seven excluded spellings missing from it.
- **With the runtime:**
  - Each on/off value of core-ts CR-015's rule 1 reads as Word does, and so does each invalid
    `w:u`/`w:vertAlign`/`w:em`, through unmarshal and `readRunOptions`. `w:u`'s six on/off words read and save
    as `none`, and `True`/`yes` as absent.
  - The same values in a DrawingML attribute of the same part still fail.
  - A caller's own callback takes precedence.
  - A dropped value's attribute isn't written on save.
  - Until jsonix-CR-010 is released, each such value is asserted to fail the part, with the upstream commit cited,
    so that the fix announces itself.
- **The probes:** the marker paragraphs' `document.xml` from `onoff-fun/` are extracted as fixtures, with a
  `SOURCE.md` naming each file. Extracted part XML only, never the archives, and never edited by hand. Each marker
  then reads as Word read it.
  - A scratch prototype of phase B, with the drop policy on `@docx4j/jsonix` 3.7.0, was compared with Word's saves
    over all 28 probes of rounds 1 to 6, on the toggles, underline, `vertAlign` and `em` (2026-10-24). 320 of 323
    markers agree. The three that differ (`colour-elsewhere-styles` U2 to U4) are direct-formatting artefacts:
    - in U2 and U3, Word's save drops a direct `w:u w:val="single"` that repeats the style's;
    - U4 is one of the two saves docx4j CR-060 section 13 found Word itself to change meaning.
    The test states those three from the PDF, as docx4j's does.
  - After a save and reload, all 323 read the same. The 341 corpus hashes don't move, and no corpus part reaches
    either callback.
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
