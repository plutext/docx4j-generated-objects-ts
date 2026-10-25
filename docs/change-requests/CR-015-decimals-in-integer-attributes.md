# CR-015: A decimal in an integer attribute, read as Word reads it rather than failing the part (`wordInvalidValue`)

**Status:** Implemented 2026-10-26 (Jason: "implement CR-015 as recommended"): phase 1 (N1 to N7), and phase 2 (N8 (b))
on `@docx4j/jsonix` 3.9.0's `strictIntegers` (jsonix-CR-012)
**Requested by:** Jason, through the docx4j-portfolio session (docx4j-portfolio-9d), 2026-10-26. The source is the
portfolio's survey of 2026-09-21, `docs/broken_producer_repairs_across_ports.md` section 2, row 1, which found the part
fails here where docx4j repairs it.
**Depends on:** docx4j CR-065 phase 0, the shared Word probe (registry `docx4j/CR-065.0`):
- round 2 (Jason, Word 365, 2026-10-26) measured the rule: truncation;
- rounds 3 and 4 (the same day) measured which text Word reads (N4, N8).

Phase 0 is complete. Phase 2 also depends on jsonix-CR-012's `strictIntegers`, released in `@docx4j/jsonix` 3.9.0
(tag 612cf08, 2026-10-26), which raised this package's floor to ^3.9.0.

Phase 1 needed nothing new in the runtime: jsonix-CR-010's `onInvalidValue`, through which CR-013 reads Word's values, is in
`@docx4j/jsonix` ^3.8.0 already.
**Counterpart:** docx4j CR-065 (docx4j-25, 2026-10-26). Jason decided its D1 to D6 as recommended, and its D3 is this
CR's rule: Word's form exactly, truncation, `""` as `0`, other forms dropped with a warning, no trimming. Its phase 1
is committed (9ea65238a, local). docx4j today repairs through
`mc-preprocessor.xslt` (`docx4j-core/src/main/resources/org/docx4j/jaxb/`, last changed 1cff2cd5c), whose rounding
templates date from 2014 (Google Docs), 2015, 2017 (`w:hanging="141.99999999999994"`), 2018 (pandoc 2.2.1, docx4j issue
#298) and 2025 (`w:firstLine`, 11.5.7). CR-065:
- phase 1 stops a part that can't be read being replaced by an empty one;
- phase 2 gives every integer attribute a lenient parse through an xjb `javaType`, the Java counterpart of section
  3.1's branch.

docx4j-python proposes keeping the text and rounding it in its readers and on save (section 3.1). All three ports take
the rule from the shared probe.

## 1. Summary

Real producers write decimals where `wml.xsd` wants an integer:
- Google Docs: `<w:tblW w:w="10206.0" w:type="dxa"/>`, `<w:pgSz w:h="16839.0" w:w="11907.0"/>`;
- pandoc 2.2.1, on `w:tblW`;
- the `w:hanging="141.99999999999994"` family;
- `w:spacing w:line="259.2"`.

Word opens all of them. Here, any decimal other than `.0` fails the whole part (`Argument [259.2] must be an
integer.`), so the part can be neither read nor re-marshalled. docx4j loads the part when the decimal is in one of its
eight attribute shapes, and rounds it half-even. That is Xalan's `format-number`, not Word's rule.

**Word truncates toward zero.** Round 2 of the shared probe tested every integer family, in the document, the styles,
the settings and a header, and Word's save wrote the integer part of each value:
- `720.5` to `720` and `721.5` to `721`;
- `-121.5` to `-121` and `-0.5` to `0`;
- `141.99999999999994` to `141`, `1.8` to `1` and `400.7` to `400`.

Phase 1: the facade's `wordInvalidValue`, which every facade unmarshal passes to the runtime (CR-013 section 3.5),
reads a decimal in an integer attribute of Word's own WordprocessingML mappings as its integer part, as Word does.
- **The model:** it then holds an integer, so a save writes one.
- **Reporting:** `onWordReading` reports each value read this way.
- **Text Word refuses.** Word reads a sign, digits and a fraction (`+240.5`, `240.`, `.5`), or `""` as `0`. Any other
  form makes it refuse the whole document, with no repair offered (rounds 3 and 4). That includes whitespace around
  the value (`" 240 "`), a comma, hex, an exponent and letters. Such a value that reaches the facade is
  dropped and reported, so the part loads and a save writes a document Word opens (N4). docx4j's CR-065 recommends the
  same.
- **Forms the runtime would read first.** The runtime's `Number()` reads some forms Word refuses, before the facade
  sees them: `" 240 "` as `240` and `0x10` as `16`. Phase 2 passes the runtime's opt-in strict parse, so they come to
  the same rule (N8).

## 2. What happens today, and what Word does (measured 2026-10-26)

**This package and docx4j.** 44 generated cases, each a `w:body` holding one shape. They were read through this package
at ea8a5ba (0.5.0's code plus CR-014), and through docx4j's own `MainDocumentPart.unmarshal` (docx4j `VERSION_17_3_5` at
2f27dca82, jars built that day). docx4j's output is marshalled by `XmlUtils.marshaltoString`.

| Input | docx4j | here (0.5.0) |
|---|---|---|
| `w:tblW w:w="10206.0"`, `w:pgSz w:h="16839.0" w:w="11907.0"` (Google Docs) | `10206`, `16839`, `11907` | the same: the runtime parses an integer with `Number()`, and `10206.0` is one |
| `w:spacing w:line="259.2"`, `w:ind w:hanging="141.99999999999994"`, `w:gridCol w:w="1234.5"` | `259`, `142`, `1234` | **the part fails** |
| half-way: `9000.5`, `9001.5`, `240.5`, `241.5`, `12.5`, `-0.5`, `-1.5` | `9000`, `9002`, `240`, `242`, `12`, `0`, `-2`: **half-even** | fails |
| `w:spacing w:before`, `w:ind w:left`, `w:sz`, `w:pgMar w:top`, `w:tab w:pos`, a `w:between` border's `w:space` (outside docx4j's list) | **the part fails** (`For input string: "120.5"`) | fails |
| a decimal on docx4j's list and one off it, in one part | fails | fails |
| `1e3`, `1.5e2`, `0x10`, `""` | fails (XPath's `number()` gives `NaN`) | `1000`, `150`, `16`, `0`: the runtime's `Number()` reads them |
| `240,5` | fails | fails |
| `+240`, `" 240 "`, `240.` | `240` | `240` |

So the portfolio was right that Xalan's `format-number` rounds half-even. One correction: `w:tblW/@w:w` is an integer
here now, not a string; the survey's `MeasurementOrPercent` was objects 0.1.6. So `"9000.5"` fails the part, as `259.2`
does.

docx4j's list, exactly (each `format-number(., '#')`):
- `@w:w` where `../@w:type='dxa'`, or on `w:pgSz`, `w:gridCol`, `w:tblW` or `w:tblInd`;
- `w:pgSz/@w:h`;
- `w:spacing/@w:line` and `@w:after`;
- `w:ind/@w:hanging` and `@w:firstLine`;
- `w:trHeight/@w:val`;
- `@w:space` on `w:top`, `w:left`, `w:bottom` and `w:right`.

The XSLT runs only after a plain unmarshal fails, or always for the parts `docx4j.jaxb.preprocess.always` names. The
second unmarshal continues on validation events, but a number format error still fails it. So in docx4j a decimal
anywhere else fails the part, as here.

**Word: round 2 of the shared probe** (docx4j CR-065 phase 0; Jason, Word 365, 2026-10-26). The cases are 65 in all,
on the share at `Office 2016/__tmp/cr065/`, with `cases.txt` listing them. Families B to I each opened with no repair
prompt, and Word's saves are `families/cr065-{B..I}-word365.docx`. Read here against the probe, case by case:
- **Every value Word wrote is the truncation of the one in the probe:**
  - signed twips (`w:ind/@w:left`: `720.5`, `720.4`, `721.5`, `-120.5`, `-121.5`, `-1.5`);
  - docx4j's list (`141.5`, `120.5`, `259.2`, `240.5`, `241.5`, `141.99999999999994`, `12.5`, `1.8`, `7.2`, `9000.5`,
    `9001.5`, `100.5`, `10.5`, `400.7`, `300.4`, `11907.5`, `16838.5`);
  - eighth-points (`w:sz="7.5"`) and the `w:between` space (`1.8`);
  - tab stops (`1440.5`, `720.5`) and a run's `w:w`, `w:spacing` and `w:position` (`90.5`, `20.5`, `6.5`);
  - a `pct` `w:tcW` and `w:tblW` (`50.5`) and `w:pgMar/@w:top` (`1133.5`, `1440.5`);
  - a style's indent, `w:defaultTabStop` in the settings, and a header's spacing.
- **Two values Word wrote nowhere:** `w:ind/@w:left="-0.5"` truncates to `0`, the default, and run `w:sz="22.5"` to `22`,
  the document default. Word's save leaves out a direct value equal to what lies beneath, as CR-013's probes found for
  `w:u`. Rounding half-up would have given `23`, which Word would have written.
- **`w:gridCol`:** Word recomputes it, so the save says nothing about its reading.
- **`""`:** written as `0` in `w:pgMar`'s `header`, `footer`, `gutter` and `top` (SSRS's shape, and an empty `top`
  alone). Empty `rsid*` attributes are dropped.
- **Family A**, `w:spacing/@w:before` with the odd forms, was refused as a whole. So round 3 (the same day) split it
  into seven files:
  - A1, plain decimals (`120.5`, `120.4`, `121.5`, `122.5`, `240.0`, `141.99999999999994`), opened clean and
    truncated;
  - A6, `""`, opened clean and read as `0`, the default, so the save wrote nothing;
  - **each of the others was refused outright, with no repair offered:** A2 (`240.`, `.5`, `+240`, `+240.5`,
    `" 240.5 "`, `" 240 "`, in one file), A3 (`240,5`), A4 (`0x10`), A5 (`1.5e2`, `2.5e0`, `1e3`) and A7 (`abc`).

  Round 4 (the same day) tried A2's six forms one at a time:
  - `240.`, `+240` and `+240.5` opened clean and read as `240`;
  - `.5` read as `0`, the default, so the save wrote nothing;
  - `" 240.5 "` and `" 240 "` were each refused outright.

  So **Word's form** is an optional `+` or `-`, then digits with an optional `.` and fraction, or a `.` and a fraction,
  with no whitespace around it; or the empty string, read as `0`. The fraction is truncated toward zero. Anything else
  refuses the whole document: whitespace around the value, a comma, hex, an exponent, letters. A lone `.` or sign is
  unmeasured.

So docx4j's repair differs from Word's reading in 7 of the probe's values on its own list: `141.5`, `241.5`,
`141.99999999999994`, `1.8`, `9001.5`, `400.7` and `11907.5`, each one more than Word's.

Word was also found to keep twips beyond 16 bits modulo 2^16 (`7205001` as an indent reads `-3959`). That is section 4's.

## 3. Proposal

### 3.1 Where: the facade's `wordInvalidValue` (N1)

The runtime calls `onInvalidValue` exactly when it refuses an attribute value, with the property's type and the
declaring class's module. Its return is stored. CR-013 already sets it on every facade unmarshal (`unmarshalString`,
`unmarshalNode`, `unmarshalPackage`), scoped to Word's own mappings (`WORD_READ_MODULES`: `wml`, `w14`, `w15`,
`w15symex`, `w16`, `w16cex`, `w16cid`). One more branch is needed:
- **When:** the type is one of the runtime's integer types.
- **Result:** for Word's form (section 3.3), the integer part (section 3.4). For any other text, which Word refuses,
  the attribute is dropped (N4 (b)).

Nothing is re-parsed. There is no stylesheet, and no second pass over the part: a value is read once, where it is
read. A caller's own `onInvalidValue` replaces the facade's, as for CR-013, and can compose `wordInvalidValue`.

A load-time pass like the MCE preprocessor would have to know each attribute's type, which the mapping already knows
here. It would also cost a pass over every part, where this costs nothing until a value is refused.

docx4j-python has proposed the opposite for its port: keep the text as written, and round it in its one integer
reader and in its save normaliser. Here that would put the string `"259.2"` in a property the declarations type
`number`. Every reader, core-ts's resolver arithmetic among them, would then have to read both types, and the runtime
can't print it, so every save would depend on the normaliser. In a Python port that keeps raw strings already, the
choice costs less. Here, the model holding the integer keeps the declarations true.

### 3.2 Which attributes (N2)

There are two choices:
- **(a) docx4j's list**, conditions included: `@w:w` only with `dxa` or on its four parents, and `@w:space` only on the
  four sides. A decimal elsewhere still fails the part, and the conditions must be checked by element name.
- **(b) every integer attribute of Word's own mappings**: 178 attribute properties on 100 types (103 `Integer`, 37
  `UnsignedLong`, 29 `Int`, 9 `Long`), keyed by the runtime's type.

**Recommendation: (b).** Word truncates in every family round 2 tried, docx4j's list and beyond it, including the
styles, settings and header parts. The alternative is losing the whole part, and docx4j's list is only the shapes
reported so far, not a rule. Keying by type also needs no element names, so it can't fall behind the schema. The ids
and counts among the 178 (`w:id`, `w:numId`, `w:ilvl`) weren't in the probe; reading `w:id="3.5"` as `3` is no worse
than failing the part.

Other mappings (DrawingML, VML, SpreadsheetML, ...) still fail as before. Their integers are another producer's
defect, unreported and unmeasured.

### 3.3 Which text (N4, N8)

**Read by this CR:** Word's form, exactly (rounds 3 and 4):
- the pattern is `^[+-]?(\d+(\.\d*)?|\.\d+)$`, then truncated (section 3.4);
- nothing is trimmed, since Word refuses `" 240.5 "`. docx4j's CR-065 takes the same form, with no trimming (its D3
  (a)), so the three ports read identically.

**A form Word refuses that reaches the facade (N4).** Examples are `" 240.5 "`, `240,5`, `2.5e0` and `abc`, and a lone
`.` or sign, which are unmeasured and outside Word's form. Word refuses the whole document for any of them. There are two choices:
- **(a) fail the part**, as 0.5.0 does;
- **(b) drop the attribute, reported.** The kind is `'value.refused'` and the value undefined. The part loads, the
  attribute reads as absent, and a save writes a document Word opens. docx4j's CR-065 recommends the same (its D3).

**Recommendation: (b), with docx4j.** Failing the part loses all of it for one attribute, and Word gives no reading
that failing could match, since it refuses the document. The report says what was dropped and why.

**Forms the runtime reads first (N8).** `@docx4j/jsonix` parses an integer with `Number()`, so some forms Word refuses
never reach `onInvalidValue`:
- `" 240 "` reads as `240`, and `0x10` as `16`;
- `1e3` reads as `1000`, and `1.5e2` as `150`.

So this package reads documents Word won't open, with values Word never read. The runtime's other lenient reads are
Word's own: `""` as `0`, `+240` and `240.` as `240`, and `10206.0` as `10206`. There are two choices:
- **(a) leave the runtime as it is.**
- **(b) a jsonix CR for an opt-in strict integer parse** (an unmarshaller option) that accepts exactly `[+-]?\d+`. Note
  that `xsd:integer` collapses whitespace, so XSD would read `" 240 "`, but Word refuses it; the option must therefore
  not trim. The forms above would then reach the facade:
  - Word's own forms, including `""` as `0`, are read as Word reads them;
  - the rest are dropped as in N4 (b);
  - outside Word's mappings, `wordInvalidValue` reads them as today, so nothing changes there.

**Recommendation: (b), as a second phase.** It makes the ports agree, since docx4j's lenient parse drops such forms
too (its D3). But no producer is known to write them, so it needn't hold back the first phase.

**Phase 2: jsonix-CR-012** (jsonix-d2, 2026-10-26; L1 and L2 decided as recommended by Jason; released in
`@docx4j/jsonix` 3.9.0). It adds a per-call unmarshaller option, `strictIntegers`, off by default.
- **What it accepts:** with the option on, an attribute of the 13 integer types reads only `/^[+-]?[0-9]+$/`, with no
  trimming. Everything else goes to `onInvalidValue` as written: `""`, `" 240 "`, `240.`, `10206.0`, `0x10`, `1e3`.
- **Where (L1, recommended):** attribute values only. Word's mappings have no integer element content or integer
  lists, and there `onInvalidValue` can receive the refusal.

The facade (phase 2, implemented 2026-10-26):
- **sets the option** on every facade unmarshal that uses its own `onInvalidValue`. A caller who passes its own
  `onInvalidValue` keeps the runtime's reading unless it passes `strictIntegers: true` too; otherwise its callback
  would start receiving `""` and `10206.0`, which it never saw before. `strictIntegers: false` asks for the runtime's
  reading with the facade's callback;
- **in other mappings,** reads a refused value with `propertyInfo.typeInfo.parse(text)`, the runtime's parse unchanged
  (jsonix-CR-012 section 4.5). That is exactly the reading before phase 2, `0x10` as `16` and a refusal of `1.5`, where
  `Number()` would not be;
- **in Word's mappings,** reads Word's forms and `""` as Word does, and drops the rest:
  - `10206.0` and `240.` are reported as `'value.truncated'`. Google Docs' widths, read silently today, are then
    reported: the value is unchanged, and the text is still one the schema refuses;
  - `""` is reported as `'value.kept.invalid'` with the value `0`, Word's reading of an invalid value.

It was verified against jsonix master before the release, then against the npm 3.9.0 tarball, whose runtime is
byte-identical to the verified checkout:
- the 341 corpus parts marshal identically, with the option off and on;
- the fallback to the runtime's parse ran 0 times over the corpus, since no Office part has an integer attribute the
  strict parse refuses.

### 3.4 The rule: truncation toward zero (N3, measured)

The rule is the integer part of the value: `Math.trunc(Number(text))`, with `-0` written as `0`. In JavaScript,
`Number()` of a plain decimal, then truncation, gives the integer part of its digits, exactly for any value a document
holds:
- `141.99999999999994` gives `141`, since the double is just below `142`;
- `-0.5` gives `0`.

This is Word's rule, measured (section 2). docx4j's half-even was Xalan's. The ports move together: docx4j's CR-065 and
docx4j-python take the same round 2.

### 3.5 Reporting (N5)

`onWordReading` receives a `WordReading` for each value read this way:
- `kind: 'value.truncated'`, with `value` the integer (`WordReading.value` widens to
  `boolean | string | number | undefined`); or `kind: 'value.refused'` (N4 (b)), with `value` undefined;
- `text`: the attribute as written;
- `element` and `attribute`, as for CR-013.

The union widens, so a consumer that switches exhaustively on `kind` fails to compile until it handles the new kind.
That is the intent: core-ts's `ValueKeptWarning` copies `kind` into its `code`.

The other choice reuses `'value.kept.invalid'`, so nothing in the types changes. But a consumer then can't tell "read
as Word reads an invalid value" from "the decimal part dropped".

**Recommendation:** the new kinds. docx4j-python CR-018 names the kinds; counterpart names there are that session's.

### 3.6 What a save writes (N6)

The model holds the integer, so every save writes an integer. A loaded document never reaches the save normaliser with a
decimal, whether `normaliseOnSave` is on or off.
- This is what Word's save writes: round 2's oracle is exactly the saves.
- The read is like CR-013's on/off reading, where `yes` is held as `true` and saved `true`.
- Jason's aim, never writing what the schema rejects, holds by default for every part this CR loads.

That leaves a caller who sets a non-integer itself (`spacing.line = 12.96 * 20`). Today the marshal throws
`Argument [259.2] must be an integer.`, through jsonix-CR-010's `onUnprintableValue`, which the facade doesn't set.
There are two choices:
- **(a) keep the throw.** A non-integer in an integer property is the caller's bug, and the throw names it;
- **(b) under the normaliser** (on by default), truncate it and report it to `onNormalised`. A save then never fails on
  it, but the caller's bug goes unseen unless they listen.

**Recommendation: (a).** The normaliser writes Word's reading of documents; it is not there to repair a caller's
arithmetic. core-ts's setters round what they compute.

### 3.7 SSRS 2012's empty attributes (N7)

SSRS 2012's Word renderer writes `<w:sectPr w:rsidRPr="" w:rsidDel="" w:rsidR="" w:rsidSect="">` and
`<w:pgMar ... w:header="" w:footer="" w:gutter=""/>`.
- **Here:** the part loads already. The runtime reads `""` as `0` in all three `w:pgMar` values, and keeps the rsids as
  `""`, writing them back as `""`.
- **Word** (round 2, H60): writes `header`, `footer` and `gutter` as `0`, exactly as read here, and drops the empty
  rsids.
- **docx4j:** writes `0` for `w:header` and `w:footer`, drops `w:gutter` (which `CT_PageMar` requires), and drops the
  empty rsids. It does this only because the empty `w:header` failed its plain unmarshal; a lone `<w:p w:rsidR="">` it
  keeps as `""`.

So the margins already read and write as Word's save does. The difference is the rsids: `ST_LongHexNumber` is four bytes
of `hexBinary`, so an empty one is schema-invalid, and Word's save leaves it out.

**Recommendation:** ride along, on the save side only. Under the normaliser, leave out an empty `rsid*` attribute,
reported to `onNormalised`. There are 20 such attribute properties in `wml`, all strings in the mapping, keyed by name.
The tree keeps `""`, as phase C keeps what it read. This is what Word's save does.

## 4. Not proposed

- **Word's 16-bit measures.** Word keeps a twips value beyond 16 bits modulo 2^16, so `7205001` as an indent reads
  `-3959`. The schema allows such a value, the runtime keeps the number as written, and no producer is known to write
  one. Reading it as Word does belongs to a resolver (core-ts), not to the parse.
- **Other mappings' integer attributes** (section 3.2).
- **docx4j's list.** docx4j's CR-065 replaces it with a lenient parse. Matching N4 and N7 there is docx4j's to decide.

## 5. Decisions

**Decided as recommended** (Jason, 2026-10-26: "implement CR-015 as recommended"). N1 to N7 are phase 1. N8 (b) is
phase 2: jsonix-CR-012, released in 3.9.0, and the facade's reading of what it brings. Both were implemented the same
day. The options and reasons stand below.

- **N1. Where:** the facade's `wordInvalidValue` (CR-013's `onInvalidValue`); a load-time pass; or, as docx4j-python
  proposes for its port, the text kept and rounded by readers and on save (section 3.1). **Recommendation:**
  `wordInvalidValue`.
- **N2. Which attributes:** (a) docx4j's list with its conditions; (b) every integer attribute of Word's own mappings
  (178). **Recommendation:** (b).
- **N3. The rule:** truncation toward zero, as Word's saves show (round 2), in all three ports. Half-even (docx4j's
  XSLT) differs from Word in 7 of the probe's values on docx4j's own list. **Recommendation:** truncation.
- **N4. A form Word refuses** (rounds 3 and 4), reaching the facade: (a) fail the part; (b) drop the attribute, reported as
  `'value.refused'`. **Recommendation:** (b), as docx4j's CR-065 D3.
- **N5. Reporting:** new kinds, `'value.truncated'` with a numeric `value` and `'value.refused'`; or reuse
  `'value.kept.invalid'`. **Recommendation:** the new kinds.
- **N6. A caller's non-integer on save:** keep the throw, or truncate it under the normaliser. **Recommendation:** keep
  the throw.
- **N7. Empty rsids:** leave them out on save under the normaliser, as Word's save does, or leave as is.
  **Recommendation:** leave them out on save.
- **N8. Forms the runtime reads first** (`0x10` as `16`, `1e3` as `1000`), which Word refuses: (a) leave the runtime as
  it is; (b) a jsonix CR for an opt-in strict integer parse, so that the facade reads them by N4. **Recommendation:**
  (b), as a second phase.

## 6. Tests

Phase 1 adds one test, `test/word365-integers.mjs`, and a CR-015 block in the smoke.

- **Fixtures:**
  - `test/fixtures/cr065-probe/` holds families A1, A2a to A2d, A6 and B to I and Word's saves of them, as extracted
    part XML. For family I it also holds `styles.xml`, `settings.xml` and the header.
  - It also holds the refused families A2e, A2f, A3 to A5 and A7, without saves; Word's answer for them is the
    refusal itself.
  - `cases.txt` is the probe's case list, and a `SOURCE.md` names the share, docx4j CR-065 section 3 and Jason's runs.
  - Never edited by hand.
- **Word as the oracle** (`test/word365-integers.mjs`). Each of the 64 cases is read through the facade from the probe
  and written again, and the case's attribute values are compared with those in Word's save.
  - 50 cases write what Word's save writes.
  - 6 are stated, strict both ways. Four are a truncation equal to what lies beneath, which Word's save leaves out:
    `.5`, `""`, `-0.5` and run `w:sz="22.5"` (this package writes the `0` or `22` it read). Two are `w:gridCol`, which
    Word recomputes.
  - In the refused files, all 8 forms are dropped and reported as `'value.refused'`. Four of them (`" 240 "`, `0x10`,
    `1.5e2`, `1e3`) the runtime's `Number()` read until phase 2.
  - Every record is `'value.truncated'` with the integer part (55, among them `10206.0` and `240.`), `'value.kept.invalid'`
    for `""` read as `0` (5), or `'value.refused'` (8).
  - Word's saves themselves need neither Word's reading nor the normaliser.
  - SSRS's empty rsids are left out on save, with their records.
  - Rounding instead of truncating fails the test (checked by changing the build).
- **The scope** (smoke): every integer attribute of Word's own mappings, 178 of them, found from the context, reads
  - Word's forms as their integer part (`7.9`, `-7.9`, `+240.5`, `.5`, `240.`, and `-0.5` as `0`, never `-0`);
  - each refused form as dropped (`" 7.9 "`, `7,9`, `2.5e0`, `abc`, `.`, `-`, `Infinity`, `" 240 "`, `0x10`, `1e3`), and
    `""` as `0`.

  Every other mapping's integer attributes read as before phase 2: `0x10`, `""`, `" 240 "` and `1e3` give `16`, `0`,
  `240` and `1000`, and `7.9` fails with the runtime's error. A `wp:extent` in a part still fails.

  This replaces the generated per-attribute fixture first proposed, since the scope is a property of the mapping,
  which the context lists.
- **The record and precedence** (smoke): `onWordReading`'s `kind`, `value` and `text`; the reading without
  `onWordReading`; a caller's own `onInvalidValue` taking precedence; and a refused `wml` dateTime still failing the
  part. For phase 2: `" 240 "`, `""`, `10206.0`, `0x10` and `+240` read and reported; `strictIntegers: false` giving
  the runtime's reading; a caller's own `onInvalidValue` keeping it unless it asks for `strictIntegers`. CR-013's block, which held a refused `wml` integer to failing, now holds it to being dropped.
- **The save** (smoke): an integer is written, and a caller's own non-integer still throws, with the normaliser on or
  off (N6). An empty rsid is left out and reported with its path, while the tree keeps `""`, and the opt-out writes it.
- **Found while testing:** phase C's `onNormalised` gave the wrong path for an attribute on an element with content.
  The runtime writes an element's attributes after its content, so the record named the last element written inside
  it. The record now takes the element written from the attribute's owner. Phase C's colours are on elements without
  content, apart from a `w:background` holding a `v:background`, which is why the fidelity corpus never showed it.
- **Fidelity:** unchanged. Office writes no decimals and no empty rsids, and `test/fidelity.mjs` fails a part that
  reaches `onWordReading` or `onNormalised`.
- `test/nodenext/consumer.mts` checks the two new kinds and a numeric `value`.

## 7. Consumers

- **core-ts:**
  - Parts that failed now load with no change in core-ts, since every facade unmarshal passes `wordInvalidValue`.
  - `ValueKeptWarning.code` gains `'value.truncated'` and `'value.refused'`. Its `valueKept` message reads "read as
    259" for the first and "dropped: Word 365 refuses a document with this value" for the second. Its `value` type
    widens to `number`.
  - Its corpus survey would say how many parts newly load.
  - The 16-bit wrap is its resolver's, if ever needed (section 4).
- **docx4j:** none from this CR. Its CR-065 phase 2 is the counterpart. Its XSLT's half-even differs from Word, and it
  records this CR's `w:between` border space as one more shape its fixed list misses.
- **docx4j-python:** keeps the raw string with a `ConverterWarning` (the survey's row 1). Its counterpart is its own.

## 8. Release

- **Changes:** parts that failed now load, `WordReading` gains two kinds, and, with N7, a save leaves out empty rsids.
  Google Docs' `.0` widths, read silently before, are reported as `'value.truncated'`.
- **Version:** under Jason's rule (minors for a substantive change of API or output), a minor.
- **Timing:** both phases are implemented, and can go out with CR-014 in 0.6.0. The runtime floor rises to ^3.9.0.
