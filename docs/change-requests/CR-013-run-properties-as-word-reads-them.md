# CR-013: Run properties read as Word reads them (`readRunOptions`, a colour table, the facade's part of the parse)

**Status:** Proposed 2026-10-24, at core-ts's request; its rules decided 2026-10-24 with core-ts CR-015's D1 to D11 (Jason: "CR-015 D1 to D11 decided as now recommended"), O4 and O6 decided as recommended (Jason, 2026-10-24); implementation waits for Jason's go
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

docx4j CR-061 (accepted 2026-10-24) writes colours in their normal form on save, which phase C follows value for
value. core-ts CR-015 (D1 to D11 decided) and docx4j-python's CR-018 are the other ports.

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
- **Phase A:** `builders/wml`'s `readRunOptions`, which core-ts's `Font`, `toApiScript` and `spanEdits` read through,
  reads `w:highlight` and `w:color` raw today. It should read them as Word does, through a colour table that core-ts
  shares rather than keeps a copy of.
- **Phase B:** jsonix-CR-010's `onInvalidValue`, set by the facade with Word's rules for Word's own WordprocessingML
  mappings, so the on/off and enumeration values load and read as Word reads them. It also records what was read so.
- **Phase C:** a save normaliser in the facade's marshal, as docx4j's CR-060 and CR-061 write, value for value:
  marshal-only, on by default, with an opt-out (core-ts CR-015's D2).

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

Dropping is decided: core-ts CR-015's D1 and D2, as docx4j does (Jason, 2026-10-24). The alternative, keeping
each invalid value's spelling and writing it back through `onUnprintableValue`, was set aside. It would have put
strings outside the declared unions into the model, and made the ports differ on save.

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
  - A table name, case-insensitive and trimmed, gives its value. Trimming is measured for hex only; names are
    trimmed too for parity with docx4j's `WordColorValues.normalise` (core-ts CR-015's D7).
  - Anything else gives `undefined`. Word reads such a value as absent in `w:color` and `w:u/@w:color`, but as
    automatic in a fill or a border colour (round 6), so the caller applies its own place's rule.

`readRunOptions`' `color` then reads `''` when `w:color` is absent or unusable, and otherwise `wordColor(val)` (O1,
section 5). core-ts's `ResolvedColor.ofHex` and `WordColors` can call `wordColor` instead of keeping a second table.
docx4j has the same table, 140 names, in `org.docx4j.wml.WordColorValues` (CR-060's D5; `ColorTables.HTML`
before 17.3.5). `WORD_COLOR_NAMES` is copied from it, and a test holds the two name lists equal.

While there: `RunFormatting.color`'s comment says "or a theme name", but `readRunOptions` reads only `w:color/@w:val`.
The comment is corrected.

### 3.5 The facade sets jsonix-CR-010's callbacks

Every facade unmarshal (`unmarshalString`, `unmarshalNode`, `unmarshalPackage`) passes an `onInvalidValue`. No
`onUnprintableValue` is needed, since dropping keeps every value inside its type. The callback is scoped by
`classInfo.module.name`. The scope is `org_docx4j_wml` and, by extension, Word's own WordprocessingML mappings:
`org_docx4j_w14`, `w15`, `w15symex`, `w16`, `w16cex` and `w16cid` (core-ts CR-015's D10, decided).
- **On/off:** a refused `Boolean` gives `true` (section 3.1). In w14 that is `CT_OnOff.val`
  (`w14:checkbox/w14:checked`, `w14:conflictMode`, `w14:discardImageEditingData`).
- **w14's two on/off values typed as an anonymous enumeration** of `true`, `false`, `0` and `1`
  (`CT_LinearShadeProperties/@scaled`, and `CT_StyleSet/@val` in `w14:stylisticSets`) take the `Boolean` rule,
  mapped onto the members: `'false'` for exactly `false`, `off` or `0`, else `'true'`.
- **Enumerations:** a refused value is dropped, but for `w:u`'s six on/off words, which give `'none'`
  (section 3.2).
- **Anything else rethrows,** so it fails as today. That covers other simple types in those mappings, and every
  type of every other mapping: DrawingML (`wp:`, `a:`, `pic:`), VML, SpreadsheetML and PresentationML, which
  Word's rule isn't measured for. Neither round measured w14; refusing a part is worse than an unmeasured reading
  by Word's rule (D10's reason).

The rule covers every `Boolean` in `org_docx4j_wml`, including the two that docx4j's `OnOffAdapter` leaves strict:
`w:writingStyle/@w:checkStyle` and `w:readModeInkLockDown/@w:actualPg`, both in the settings part. docx4j leaves
them out for a Java reason: they are required, so they bind to a primitive `boolean`, which an adapter can't bind.
They are `ST_OnOff` in the spec, and Word's reading of them is unmeasured in either port. Applying Word's rule
keeps `settings.xml` loading where docx4j would refuse it, so this is a known difference between the ports rather
than one to match.

Scoping by the class's mapping, not by the part, needs no per-part choice. A `document.xml` gets Word's rule for
its WordprocessingML and today's strictness for the DrawingML it holds. CR-009's `org_docx4j_wml_globals`
declares elements of `org_docx4j_wml`'s types, so their class infos answer `org_docx4j_wml` too.

A caller's own `onInvalidValue`, passed through CR-007's per-call options, takes precedence. The facade's is
exported (`wordInvalidValue`) for a caller that wants to compose it.

**The record** (core-ts CR-015's D8). Each value read by Word's rule is reported to a facade option,
`onWordReading(record)`. The record holds:
- the source element and the attribute's name;
- the text as read, and the value the model holds instead (`true`, `'none'`, or none if dropped);
- its kind: `'value.kept.exact'` if the text would be a lexical form once trimmed (`" false"`, `" yellow"`), else
  `'value.kept.invalid'` (`"yes"`, `"Yellow"`, `"bogus"`). jsonix doesn't collapse whitespace, so both kinds reach
  the callback, and the facade tells them apart by trimming.

core-ts turns each record into a `PartWarning` on the part, since the part is core-ts's to know. Whether the facade
raises warnings itself or hands over the data was O6: the data, to callbacks (decided).

Word's enumeration rule is measured for `w:u`, `w:vertAlign` and `w:em` only, and `w:highlight` is a string here.
Every enumeration in the scope is lenient, not only those three (O4, decided).

### 3.6 What a save writes: the normaliser (core-ts CR-015's D2, decided)

Jason's principle: "they are entitled to assume we don't write dodgy content". A save never changes Word's
reading, and never writes content the schema rejects. Two parts are already settled by the parse:
- **On/off** is written `true`/`false`, as jsonix writes a boolean, and a bare value stays bare. That holds with
  the normaliser on or off (D11: the ports agree on readings, not on spelling).
- **An invalid enumeration value** is already absent from the model, so it isn't written. `w:u`'s on/off words
  are written `none`, as Word saves them, and any other invalid `w:u` is written bare.

The rest is the normaliser: a pass over the marshalled DOM, in `marshalToDocument` beside
`fixRootNamespaceDeclarations`. It works on WordprocessingML elements wherever they sit, including a revision's
former properties (`w:rPrChange/w:rPr`, a paragraph mark's), which Reject All restores (D9). It writes docx4j
CR-060 with CR-061's normal form, value for value:

| Where | Accepted | `auto` | Rejected, or missing |
|---|---|---|---|
| `w:color/@w:val` | upper-case `RRGGBB` | `auto` | the `w:color` dropped where it has no theme attribute; beside a `w:themeColor`, or a `w:themeTint` or `w:themeShade` without one, kept with `w:val="auto"`, since `w:val` is required |
| `w:u/@w:color` | upper-case `RRGGBB` | `auto` | the attribute dropped; the `w:u` keeps its `w:val` |
| `w:shd/@w:fill`, `w:shd/@w:color`, `w:background/@w:color`, a border's `w:color` | upper-case `RRGGBB` | `auto` | `auto` |
| `w:highlight`, `w:vertAlign`, `w:em` | kept | | the element dropped, as Word's repair does |

"Accepted" is `wordColor`'s: a name, hex in any case, with an optional `#` and padding. A value already in the
normal form is written as it is. The theme attributes are untouched. `w:shd/@w:color` and `w:background/@w:color`
follow the fill's rule, pending docx4j CR-061 phase 3: Jason's Word check of docx4j's saves, and a round 7 probe of
those two attributes. docx4j's implementation (aaadd93f9, unpushed) and its tests, `ColourNormalFormTest` and
`ColourNormalFormOptOutTest` in core-tests, are the reference for each case.

**Marshal-only** (docx4j's D6): the tree in memory keeps the source's value, and only the output is normalised.
So a caller's reading of the tree is the same before and after a save.

**The opt-out** (docx4j's D1 and D2):
- A process-wide default, `setNormaliseOnSave(false)`, and a per-call marshal option, `normaliseOnSave`. core-ts's
  per-package setting (`pkg.normaliseOnSave`) passes the per-call option.
- Off means colours are written as read, and the `w:highlight`/`w:vertAlign`/`w:em` drop is off.
- On/off and `w:u`, being settled at the parse, are written the same either way.

**The report** (D8): each rewrite is passed to a marshal option, `onNormalised(record)`. The record holds the
written element's path, the attribute (or the element, if dropped), the value before and the value after (none if
dropped). core-ts adds the part and gathers the records into its save report (`value.normalised`).

**A part never unmarshalled** is copied byte for byte by core-ts, as now.

**The fidelity corpus can't move.** A scan of all 341 parts (2026-10-24) found no colour outside the normal form
(Office writes upper-case hex), and Office never writes a valueless or invalid `w:highlight`, `w:vertAlign` or
`w:em`. The test holds the hashes to that.

## 4. Not proposed

- **Word's spelling of on/off** (bare, and off as `0` or dropped). Word reads `true`/`false` exactly, and the
  ports agree on readings, not on bytes (core-ts CR-015's D11).
- **Office JS's own answers** for these values (core-ts's D4, its check 60), and rendering.
- **The callback mechanism itself:** that's jsonix-CR-010's.
- **Widening the generated enumeration declarations** to admit a kept string. Dropping makes it unnecessary.
- **A character style's highlight,** which Word doesn't draw (round 6). That is style resolution, which belongs
  to core-ts and docx4j; `readRunOptions` reads direct formatting only.

## 5. Decisions

**Settled by core-ts CR-015's D1 to D11** (Jason, 2026-10-24), where they overlap:
- **O1, an unusable `w:color` in `readRunOptions`:** `''`, as if absent (D5, D7: an unusable run colour is absent
  over a style).
- **O2, the colour table:** `helpers/wml`'s `WORD_COLOR_NAMES`, held to docx4j's `ColorTables.HTML`, with core-ts
  calling `wordColor` (D3).
- **O3, the drop on marshal:** in the facade, as the normaliser (D2: "objects CR-013.C grown into the normaliser").
- **O5, Word's extension mappings in the lenient set:** yes, with w14's on/off enumerations read by the `Boolean`
  rule (D10).
- **Keep or drop:** drop (D1, D2), with `w:u`'s on/off words read as `none` (D6).

**Decided by Jason as recommended** (2026-10-24, "CR-013 O4 and O6 as recommended"):
- **O4. Which wml enumerations are lenient:** every one (a refused value anywhere in `org_docx4j_wml` loads), or
  only `w:u`, `w:vertAlign` and `w:em`, the ones Word was measured for. Refusing the part is the worse half
  (core-ts CR-015 section 1). But every reader of another enumeration (`w:fldChar/@w:fldCharType`, say) must then
  treat a non-member as unset. With dropping, no reader meets one: the attribute is absent. **Recommendation:**
  every one, since Word opened every measured case and a part that fails can't be read at all.
- **O6. The record and the report** (D8): the facade hands the data to callbacks (`onWordReading`, `onNormalised`),
  and core-ts makes the `PartWarning`s and its save report; or the facade raises warnings itself. This package has
  no warnings model, so raising means `console.warn` per value, which is noise to an Office JS add-in. docx4j
  only logs at DEBUG. **Recommendation:** callbacks, with nothing written to the console.

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
- **The normaliser (phase C):**
  - Each row of section 3.6's table, in a run's `w:rPr`, a paragraph mark's, a style's and a `w:rPrChange`'s former
    properties. That covers a name, lower-case hex, `#` and padding written upper-case; `auto` kept; `Auto` and
    `F00` dropped from `w:color`, and written `auto` beside a `w:themeColor`, and beside a `w:themeTint` or
    `w:themeShade` alone; a rejected underline colour dropped;
    a rejected fill or border colour written `auto`; a valueless or invalid `w:highlight`, `w:vertAlign` and `w:em`
    dropped.
  - The tree is unchanged after the marshal.
  - The opt-out, both process-wide and per call, writes colours as read and keeps the three elements.
  - `onNormalised` receives one record per rewrite.
  - Every probe saved and reread reads as the probe did, as Word's PDFs show for U4 and H7, the two saves in which
    Word changed meaning.
  - The fidelity corpus's 341 hashes don't move.
  - The cases of docx4j's `ColourNormalFormTest` and `ColourNormalFormOptOutTest`, written the same.
- **The record (phase B):** `onWordReading` receives `value.kept.exact` for `" false"` and `value.kept.invalid` for
  `"yes"`.
- **The scope (phase B):** `w14:checked w:val="TRUE"` reads true; w14's `scaled="on"` reads `'true'`; a DrawingML
  boolean in the same part still fails.

## 7. Consumers

- **core-ts:** its `Font`, `toApiScript` and `spanEdits` read through `readRunOptions`, and its CR-015 depends on this.
  It drops its own colour table in favour of `wordColor` (O2). It passes `pkg.normaliseOnSave` as the marshal option,
  and turns `onWordReading` and `onNormalised` records into its `PartWarning`s and save report (O6).
- **Every consumer that marshals through the facade:** colours outside the normal form are written in it, on by
  default. The meaning to Word is unchanged, and Office-written parts are untouched.
- **Office JS add-ins using this package alone:** the same readings, through `readRunOptions`.

## 8. Phases and release

- **Phase A**, with no dependency: sections 3.3 and 3.4. Half a day.
- **Phase B**, on jsonix 3.7.0: sections 3.1, 3.2 and 3.5, the scope of D10, the record, and the probe fixtures,
  with the runtime floor raised to ^3.7.0. With O4 and O6. Half a day.
- **Phase C**, after A (`wordColor`) and B (the invalid values load): section 3.6's normaliser, its opt-out and its
  report. One day.

Phases A and B are patches, as 0.3.1 and 0.4.1 were: a fix to readings, an additive export, a raised runtime
floor. Phase C changes marshal output by default for parts with colours outside the normal form. That may make it a
minor, as CR-009's behaviour and output change made 0.4.0; or a patch, as CR-008's added `mc:Ignorable` was in
0.3.1. Jason's call. No API is removed.
