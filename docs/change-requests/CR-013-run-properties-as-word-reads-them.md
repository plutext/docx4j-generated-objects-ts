# CR-013: Run properties read as Word reads them (`readRunOptions`, a colour table, the facade's part of the parse)

**Status:** Phases A and B implemented 2026-10-24, not yet released (with `@docx4j/jsonix` ^3.7.0); phase C waits for jsonix-CR-011's release
**Requested by:** core-ts (session docx4j-core-ts-5b), 2026-10-24, for its CR-015 section 6, which the portfolio
session asked it to raise here. core-ts's CR was drafted at Jason's word ("yes please", to one CR each in docx4j,
core-ts and docx4j-python against the same rules and probes).
**Depends on:** jsonix-CR-010 for sections 3.1, 3.2, 3.5 and 3.6. It hands an attribute value its type refuses to
a per-call `onInvalidValue` instead of failing the part, and a value its type can't print to `onUnprintableValue`;
the facade sets them with Word's rules. Jason decided its J1 to J3 as recommended, and it was released in
`@docx4j/jsonix` 3.7.0 (tag d4e8af8, 2026-10-24), so phase B raised the floor to ^3.7.0. Section 3.6, phase C,
depends on jsonix-CR-011: marshal-side adapters, requested for it at Jason's word. Jason decided its K1 to K3 as
recommended, and it was implemented on jsonix master (6c24d4c, 2026-10-24), not yet released. Sections 3.3 and 3.4
depend on nothing.
**Counterpart:** docx4j CR-060, the same rules in Java, implemented through phase 3 on `VERSION_17_3_5`
(868a1797b, unpushed, 2026-10-24):
- an `OnOffAdapter` on 46 of `wml.xsd`'s 48 boolean attributes, as annox annotations. A regeneration from that xsd
  is byte-identical to this package's `modules/`, checked 2026-10-24;
- a value outside an enumeration read as absent;
- marshal-only adapters that leave out a `w:highlight`, `w:vertAlign` or `w:em` whose value is missing or invalid.

docx4j CR-061 writes colours in their normal form on save, which phase C follows value for value. It was
completed on 2026-10-24 (5163c2986, unpushed). Word 365 opened docx4j's saves of fifteen probes with no repair
prompt, and read them at all 251 markers as it read the sources. Its `wml.xsd` annotations
(`NormaliseOnSave.*`, `UnderlineValueAdapter`) regenerate byte-identically here too. core-ts CR-015 (D1 to D11 decided) and docx4j-python's CR-018 are the other ports.

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
- **Phase A** (implemented): `builders/wml`'s `readRunOptions`, which core-ts's `Font`, `toApiScript` and
  `spanEdits` read through, read `w:highlight` and `w:color` raw. It reads them as Word does now, through a colour
  table that core-ts shares rather than keeps a copy of.
- **Phase B** (implemented): jsonix-CR-010's `onInvalidValue`, set by the facade with Word's rules for Word's own
  WordprocessingML mappings, so the on/off and enumeration values load and read as Word reads them. It also records
  what was read so.
- **Phase C:** a save normaliser in the facade's marshal, as docx4j's CR-060 and CR-061 write, value for value:
  marshal-only, on by default, with an opt-out (core-ts CR-015's D2).

## 2. What this package did before (measured 2026-10-24, at 1e4e895)

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
`" false"`, `""`). That relies on jsonix comparing the text exactly, with none of the whitespace collapse XSD
gives `xs:boolean`: a conforming parse would read `" false"` as false, where Word reads it as on. jsonix records
the exact comparison as part of its contract (jsonix-CR-010 section 13, jsonix 5258d81), with tests that fail if
the parse ever collapses or folds case, for booleans and enumerations alike. The model then holds Word's boolean, so `on()` is already right, and so is every reader in
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
    trimmed too for parity with docx4j's `WordColorValues.normalise` (core-ts CR-015's D7). The trim is Java's
    `String.trim`, every character up to U+0020, as docx4j's is.
  - Anything else gives `undefined`. Word reads such a value as absent in `w:color` and `w:u/@w:color`, but as
    automatic in a fill or a border colour (round 6), so the caller applies its own place's rule.

`readRunOptions`' `color` then reads `''` when `w:color` is absent or unusable, and otherwise `wordColor(val)` (O1,
section 5). core-ts's `ResolvedColor.ofHex` and `WordColors` can call `wordColor` instead of keeping a second table.
docx4j has the same table, 140 names, in `org.docx4j.wml.WordColorValues` (CR-060's D5; `ColorTables.HTML`
before 17.3.5). `WORD_COLOR_NAMES` is copied from it by script, a frozen object with no prototype, so that a lookup
of `constructor` finds nothing. The smoke holds the two tables equal, name for name and value for value, where
`../docx4j` is checked out beside this repository; CI has no such checkout, and says so.

While there: `RunFormatting.color`'s comment says "or a theme name", but `readRunOptions` reads only `w:color/@w:val`.
The comment is corrected.

### 3.5 The facade sets jsonix-CR-010's callbacks

Every facade unmarshal (`unmarshalString`, `unmarshalNode`, `unmarshalPackage`) passes an `onInvalidValue`. No
`onUnprintableValue` is needed, since dropping keeps every value inside its type. The callback is scoped by
`classInfo.module.name`. The scope is `org_docx4j_wml` and, by extension, Word's own WordprocessingML mappings:
`org_docx4j_w14`, `w15`, `w15symex`, `w16`, `w16cex` and `w16cid` (core-ts CR-015's D10, decided).
- **On/off:** a refused `Boolean` gives `true` (section 3.1). In w14 that is `CT_OnOff.val`
  (`w14:checkbox/w14:checked`, `w14:conflictMode`, `w14:discardImageEditingData`).
- **w14's two on/off values declared as an enumeration** of `true`, `false`, `0` and `1`
  (`CT_LinearShadeProperties/@scaled`, and `CT_StyleSet/@val` in `w14:stylisticSets`) are strings in the model,
  found in implementation. The compiler writes the enumeration on the property, which the runtime doesn't read, and
  the declarations type them `string`. So they are never refused, and hold their text as written (`scaled="on"`
  reads `'on'`). docx4j holds both as a `String` too. The rule proposed here, the `Boolean` rule mapped onto the
  members, has nothing to apply to. The smoke pins the strings, so a regeneration that types them as an enumeration
  fails it, and brings the rule back.
- **Enumerations:** a refused value is dropped, but for `w:u`'s six on/off words, which give `'none'`
  (section 3.2).
- **Anything else rethrows,** so it fails as before. That covers other simple types in those mappings, and every
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
its WordprocessingML and the runtime's strictness for the DrawingML it holds. CR-009's `org_docx4j_wml_globals`
declares elements of `org_docx4j_wml`'s types, so their class infos answer `org_docx4j_wml` too.

A caller's own `onInvalidValue`, passed through CR-007's per-call options, takes precedence, and `onWordReading`
is then not called. The facade's is exported (`wordInvalidValue`) for a caller that wants to compose it.
`unmarshalPackage` takes no options, so it reads by Word's rule and reports to no one.

**The record** (core-ts CR-015's D8). Each value read by Word's rule is reported to a facade option,
`onWordReading(record)`, which `unmarshalString` and `unmarshalNode` take beside the runtime's options (the type
is `UnmarshalOptions`; the record's is `WordReading`). The record holds:
- the source element and the attribute's name;
- the text as read, and the value the model holds instead (`true`, `'none'`, or undefined if dropped);
- its kind: `'value.kept.exact'` if the text, trimmed of XML whitespace, is a form its type reads (`" false"`,
  `" superscript"`), else `'value.kept.invalid'` (`"yes"`, `"Single"`, `"bogus"`). A boolean's forms are the
  runtime's: `true`, `false`, `1`, `0`, `on`, `off`. jsonix doesn't collapse whitespace, so both kinds reach the
  callback, and the facade tells them apart by trimming. `w:u="true"` is `'value.kept.invalid'`, since `true` is
  not an `ST_Underline` value, though Word reads it as `none`.

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

The rest is the normaliser, on jsonix-CR-011's marshal-side adapters, as docx4j's is on JAXB's `XmlAdapter`s:
- **`onAttributeValue`**, called after a value prints, rewrites the colour attributes by their class and property
  (`CT_Color.val`, `CT_Underline`'s colour, `CT_Shd`'s fill and colour, `CT_Background`'s colour, `CT_Border`'s
  colour). It is given the owning object, so `CT_Color.val` can see a theme attribute beside it.
- **`omitElement`**, called before an element is written, leaves out a `w:highlight`, `w:vertAlign` or `w:em` with
  a missing or invalid value, and a `w:color` whose value is rejected and that has no theme attribute.

Both are keyed by type, not by element name, so a `CT_Border` is caught under any of its dozen element names, and
both are scoped by `classInfo.module.name` as `onInvalidValue` is. `omitElement` covers every typed element path,
so a revision's former properties (`w:rPrChange/w:rPr`, a paragraph mark's), which Reject All restores (D9), come
for free. It writes docx4j CR-060 with CR-061's normal form, value for value:

| Where | Accepted | `auto` | Rejected, or missing |
|---|---|---|---|
| `w:color/@w:val` | upper-case `RRGGBB` | `auto` | the `w:color` dropped where it has no theme attribute; beside a `w:themeColor`, or a `w:themeTint` or `w:themeShade` without one, kept with `w:val="auto"`, since `w:val` is required |
| `w:u/@w:color` | upper-case `RRGGBB` | `auto` | the attribute dropped; the `w:u` keeps its `w:val` |
| `w:shd/@w:fill`, `w:shd/@w:color`, `w:background/@w:color`, a border's `w:color` | upper-case `RRGGBB` | `auto` | `auto` |
| `w:highlight`, `w:vertAlign`, `w:em` | kept | | the element dropped, as Word's repair does |

"Accepted" is `wordColor`'s: a name, hex in any case, with an optional `#` and padding. A value already in the
normal form is written as it is. The theme attributes are untouched. Round 7 (CR-061 phase 3) measured the other
two:
- `w:shd/@w:color`, the pattern colour, follows the fill's rule: a rejected value is automatic and saved `auto`.
- `w:background/@w:color`, the page colour, reads names. Word drops a `w:background` whose colour it rejects, where
  docx4j writes `w:color="auto"`. Both are valid, and both mean no page colour. This package writes `auto`, value
  for value with docx4j.

docx4j's implementation (aaadd93f9 to 5163c2986, unpushed) and its tests, `ColourNormalFormTest` and
`ColourNormalFormOptOutTest` in core-tests, are the reference for each case.

**Marshal-only** (docx4j's D6): the tree in memory keeps the source's value, and only the output is normalised.
So a caller's reading of the tree is the same before and after a save.

**Not reached:** WordprocessingML held as DOM. That happens only beneath an element the model doesn't know, such
as a producer's own element inside a kept `mc:Choice`; a `w:sdt` there is typed, as `SdtBlock`. jsonix-CR-011's
hooks don't see DOM, and docx4j's adapters don't either, so it is written as read in both ports. None of the 317
parts of the corpus and the probes holds a phase C target as DOM (checked 2026-10-24). No mapping here has an
element map.

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

Phases A and B are tested as below (2026-10-24); phase C's tests are proposed.

The smoke, a block of its own:
- **Without the runtime:**
  - `readRunOptions`' `w:highlight`: `YELLOW`, `bogus` and `none` read null, `yellow` reads `#FFFF00`.
  - For `w:color`: hex with `#`, in lower and mixed case and padded; `auto` (`'auto'`); `Auto`, `AUTO`, `" auto"`,
    `"auto "` and `#auto` (`''`); names (`red`, `RED`, `" red "`, `darkGray`, `lightgrey`,
    `lightgoldenrodyellow`, `Navy`, ...); and the non-names (`grey`, `lightGray`, `darkYellow`, `darkgrey`,
    `rebeccapurple`, `transparent`, `windowtext`, `"dark red"`, `#red`, `F00`, `0xFF0000`, `FF00000`, `GG0000`,
    `""`), which read `''`.
  - `wordColor` itself, a tab and a newline trimmed, and `constructor`, `toString` and `__proto__` not names.
  - The table's 140 names, frozen, the seven excluded spellings missing from it, and docx4j's `WordColorValues`
    equal to it where `../docx4j` is checked out.
  - `w:u` with a value outside `ST_Underline`, and bare, reads `None`.
- **With the runtime:**
  - Each on/off value of core-ts CR-015's rule 1 reads as Word does through unmarshal and `readRunOptions` (13 on,
    3 off), on `w:caps` too; each invalid `w:u`/`w:vertAlign`/`w:em` is dropped. `w:u`'s six on/off words read and
    save as `none`, and `True`, `TRUE`, `yes` and `no` as absent.
  - A dropped value's attribute isn't written on save, and a refused boolean is written `true`.
  - The scope: `w14:checked w14:val="TRUE"` reads true; w14's two enumerations hold `on` and `yes` as strings
    (section 3.5). `wp:anchor behindDoc="FALSE"` and `wp:docPr hidden="yes"` in a part with a lenient `w:b` still
    fail it, and so does a `w:sz` that isn't a number.
  - `unmarshalPackage` and `wml()` read by the rule too.
  - The record: `onWordReading` receives `value.kept.exact` for `" false"` and `" superscript"`, and
    `value.kept.invalid` for `"yes"`, `w:u="true"` and `"DOT"`, each with its element, attribute, text and value.
  - A caller's own callback takes precedence, `onWordReading` is then not called, `wordInvalidValue` composes, and
    a caller that rethrows has the runtime's strictness back.

**The probes**, `test/word365.mjs`: each probe's `word/document.xml` and Word 365's save of it, rounds 1 to 6, 28
probes, are fixtures in `test/fixtures/word365-probes/`. They are extracted part XML only, never the archives, and
never edited by hand, with a `SOURCE.md`. A marker is the middle run of each three-run paragraph after the title,
and its saved run the last run of the same paragraph in the save holding its text, as docx4j's
`RunPropertiesWord365Test` finds them. Each marker's direct toggles (`b`, `bCs`, `i`, `iCs`, `caps`, `smallCaps`,
`strike`, `dstrike`, `outline`, `shadow`, `emboss`, `imprint`, `vanish`), `w:u` value and colour, `vertAlign`,
`em`, colour and highlight are compared with the save's (2026-10-24):
- 376 markers, 373 equal. A valueless `w:u` reads as unset, as Word's save drops one. Where neither the run nor its
  paragraph has a style, a direct `none` underline and an `auto` colour read as unset, since Word's save drops
  them and nothing beneath differs (those probes have no styles part).
- The three that differ are `colour-elsewhere-styles` U2 to U4, over a character style's single red underline:
  Word's save drops a direct `w:u w:val="single"` that repeats the style's. They are stated, strict both ways.
- As docx4j's test does, Word's reading from its PDF replaces its save's where the save changed what Word drew:
  `props-enum-styles` M6 and `highlight-char-styles` H7 (a direct green highlight over a character style's), and
  U4's underline colour (`auto`, drawn black).
- 88 values in the probes are read by Word's rule, and none in Word's saves.
- After a save through the facade and a reload, all 376 read the same.
- Before phase B, 13 of the 28 probes failed to load at all. The prototype's count, 323 markers with 320 equal,
  came from a scratch script lost with its session; the three differences were the same.

**The corpus:** the 341 marshal hashes are identical to 0.4.3's (`test/lib/marshal-hashes.mjs` against the npm
tarball, both on jsonix 3.7.0), and `test/fidelity.mjs` fails a part that reaches `onWordReading`. None does.

**The normaliser (phase C)**, proposed:
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

## 7. Consumers

- **core-ts:** its `Font`, `toApiScript` and `spanEdits` read through `readRunOptions`, and its CR-015 depends on this.
  It drops its own colour table in favour of `wordColor` (O2). It passes `pkg.normaliseOnSave` as the marshal option,
  and turns `onWordReading` and `onNormalised` records into its `PartWarning`s and save report (O6).
- **Every consumer that marshals through the facade:** colours outside the normal form are written in it, on by
  default. The meaning to Word is unchanged, and Office-written parts are untouched.
- **Office JS add-ins using this package alone:** the same readings, through `readRunOptions`.

## 8. Phases and release

- **Phase A**, with no dependency: sections 3.3 and 3.4. Implemented 2026-10-24.
- **Phase B**, on jsonix 3.7.0: sections 3.1, 3.2 and 3.5, the scope of D10, the record, and the probe fixtures,
  with the runtime floor raised to ^3.7.0. With O4 and O6. Implemented 2026-10-24, with phase A, on `main` for
  0.4.4 (Jason, 2026-10-24).
- **Phase C**, after A (`wordColor`), B (the invalid values load) and jsonix-CR-011's release: section 3.6's
  normaliser, its opt-out and its report, with the runtime floor raised to that release. One day.

Phases A and B are patches, as 0.3.1 and 0.4.1 were: a fix to readings, an additive export, a raised runtime
floor. Phase C changes marshal output by default for parts with colours outside the normal form. That may make it a
minor, as CR-009's behaviour and output change made 0.4.0; or a patch, as CR-008's added `mc:Ignorable` was in
0.3.1. Jason's call. No API is removed.
