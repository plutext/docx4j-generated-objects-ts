# CR-015: A decimal in an integer attribute, rounded as docx4j rounds it (`wordInvalidValue`)

**Status:** Proposed 2026-10-26
**Requested by:** Jason, through the docx4j-portfolio session (docx4j-portfolio-9d), 2026-10-26. The source is the
portfolio's survey of 2026-09-21, `docs/broken_producer_repairs_across_ports.md` section 2, row 1, which found the part
fails here where docx4j repairs it.
**Depends on:** nothing new. jsonix-CR-010's `onInvalidValue`, through which CR-013 reads Word's values, is in
`@docx4j/jsonix` ^3.8.0 already.
**Counterpart:** docx4j's `mc-preprocessor.xslt` (`docx4j-core/src/main/resources/org/docx4j/jaxb/`, last changed
1cff2cd5c). Its rounding templates date from 2014 (Google Docs), 2015, 2017 (`w:hanging="141.99999999999994"`), 2018
(pandoc 2.2.1, docx4j issue #298) and 2025 (`w:firstLine`, 11.5.7). Extending them past their list (section 3.2) is
docx4j's to file, with Jason.

## 1. Summary

Real producers write decimals where `wml.xsd` wants an integer:
- Google Docs: `<w:tblW w:w="10206.0" w:type="dxa"/>`, `<w:pgSz w:h="16839.0" w:w="11907.0"/>`;
- pandoc 2.2.1, on `w:tblW`;
- the `w:hanging="141.99999999999994"` family;
- `w:spacing w:line="259.2"`.

Word opens all of them. Here, any decimal other than `.0` fails the whole part (`Argument [259.2] must be an
integer.`), so the part can be neither read nor re-marshalled. docx4j loads the part when the decimal is in one of
its eight attribute shapes: it rounds the value half-even and reads the rest typed.

Proposed: the facade's `wordInvalidValue`, which every facade unmarshal already passes to the runtime (CR-013 section
3.5), rounds a decimal in an integer attribute of Word's own WordprocessingML mappings, half-even, as docx4j does. The
model then holds an integer, so a save writes one. `onWordReading` reports each value it rounds.
- **docx4j's list:** this matches docx4j value for value.
- **Other integer attributes:** docx4j fails the part on these, and this package would load it (N2).
- **Text that is no decimal number** (`240,5`) still fails the part, as it does in docx4j.

## 2. What happens today (measured 2026-10-26)

44 generated cases, each a `w:body` holding one shape. They were read through this package at ea8a5ba (0.5.0's code
plus CR-014), and through docx4j's own `MainDocumentPart.unmarshal` (docx4j `VERSION_17_3_5` at 2f27dca82, jars built
that day). docx4j's output is marshalled by `XmlUtils.marshaltoString`.

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

So the portfolio's inference about the rounding mode is right: Xalan's `format-number` rounds half-even. One
correction: `w:tblW/@w:w` is an integer here now, not a string. The survey's `MeasurementOrPercent` was objects 0.1.6.
So `"9000.5"` fails the part, as `259.2` does.

docx4j's list, exactly (each `format-number(., '#')`):
- `@w:w` where `../@w:type='dxa'`, or on `w:pgSz`, `w:gridCol`, `w:tblW` or `w:tblInd`. So a `pct` table width
  `50.5` gives `50`, and a `w:tcW` only with `dxa`;
- `w:pgSz/@w:h`;
- `w:spacing/@w:line` and `@w:after`;
- `w:ind/@w:hanging` and `@w:firstLine`;
- `w:trHeight/@w:val`;
- `@w:space` on `w:top`, `w:left`, `w:bottom` and `w:right`.

The XSLT runs only after a plain unmarshal fails, or always for the parts `docx4j.jaxb.preprocess.always` names. The
second unmarshal then continues on validation events, but a number format error still fails it. So in docx4j a decimal
anywhere else fails the part, as here.

## 3. Proposal

### 3.1 Where: the facade's `wordInvalidValue` (N1)

The runtime calls `onInvalidValue` exactly when it refuses an attribute value, with the property's type and the
declaring class's module. Its return is stored. CR-013 already sets it on every facade unmarshal (`unmarshalString`,
`unmarshalNode`, `unmarshalPackage`), scoped to Word's own mappings (`WORD_READ_MODULES`: `wml`, `w14`, `w15`,
`w15symex`, `w16`, `w16cex`, `w16cid`). One more branch is needed:
- **When:** the type is one of the runtime's integer types, and the text is a decimal number in XPath's form (section
  3.3).
- **Result:** the value, rounded (section 3.4).

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
- **(a) docx4j's list**, conditions included: `@w:w` only with `dxa` or on its four parents, `@w:space` only on the four
  sides. This is value for value with docx4j, a decimal elsewhere still fails the part, and the conditions must be
  checked by element name.
- **(b) every integer attribute of Word's own mappings**: 178 attribute properties on 100 types (103 `Integer`, 37
  `UnsignedLong`, 29 `Int`, 9 `Long`), keyed by the runtime's type. On docx4j's list this is value for value too, since
  the same rounding is applied (the prototype agrees on every case docx4j loads, section 6). Elsewhere it loads a part
  docx4j fails: the rest of `w:spacing` and `w:ind`, `w:sz`, `w:pgMar`, `w:tab/@w:pos`, `w:between`'s space, and ids
  and counts (`w:id="3.5"` reads as `4`).

**Recommendation: (b).** The alternative to rounding is losing the whole part, and a decimal in `w:before` is the same
defect as one in `w:after`; docx4j's list is the shapes reported so far, not a rule. (b) agrees with docx4j wherever
docx4j loads. It also needs no element names, so it can't fall behind the schema.

Other mappings (DrawingML, VML, SpreadsheetML, ...) still fail as before. Their integers are another producer's
defect, and unreported.

### 3.3 Which text (N4)

The text must be a decimal number as XPath 1.0's `number()` reads it, which is what docx4j rounds:
- optional whitespace, an optional `-`, digits with an optional `.` and digits, or `.` and digits, then whitespace;
- in a pattern, `^[ \t\r\n]*-?(\d+(\.\d*)?|\.\d+)[ \t\r\n]*$`.

Anything else that the runtime refuses still fails the part, as it does in docx4j: `240,5`, `2.5e0`, `+240.5`,
`Infinity`, `abc`.

The other choice is to drop such an attribute, as CR-013 drops an invalid enumeration value. Word's reading of such
text is unmeasured, so failing as before is the conservative choice.

**Recommendation:** fail as before, until a Word probe says otherwise (section 3.4).

The runtime reads some text that never reaches `onInvalidValue` (section 4): `""`, `1e3` and `0x10` give `0`, `1000`
and `16`.

### 3.4 Rounding (N3)

The rounding is half-even, as docx4j's `format-number` does it (measured, section 2): `240.5` gives `240`, `241.5`
`242`, `-0.5` `0` (never `-0`) and `141.99999999999994` `142`. In JavaScript, with `f = Math.floor(v)` and
`d = v - f` (exact for any value a document holds):
- `d < 0.5` gives `f`;
- `d > 0.5` gives `f + 1`;
- a tie gives whichever of the two is even.

Word's own reading is unmeasured: whether it rounds, truncates or ignores a decimal. "As Word reads it", rather than
"as docx4j repairs it", needs a Word probe, which only Jason can run. As in CR-013's rounds, Word's save would be the
oracle: a document with each family's decimals at `.2`, `.5` on an even and an odd integer, `.8`, a negative `.5`,
`141.99999999999994`, and the forms of section 3.3, opened and saved in Word 365. If Word differs, docx4j's XSLT and
this rule change together.

**Recommendation:** half-even now, value for value with docx4j. The probe is worth running only if a consumer needs
Word's answer for the half-way cases; the reported producers write `.0`, `.2`, `.8` and `.99999999999994`, where every
rounding rule agrees and truncation doesn't.

### 3.5 Reporting (N5)

`onWordReading` receives a `WordReading` for each value rounded:
- `kind: 'value.rounded'`;
- `value`: the integer (`WordReading.value` widens to `boolean | string | number | undefined`);
- `text`: the attribute as written;
- `element` and `attribute`, as for CR-013.

The union widens, so a consumer that switches exhaustively on `kind` fails to compile until it handles the new kind.
That is the intent: core-ts's `ValueKeptWarning` copies `kind` into its `code`.

The other choice reuses `'value.kept.invalid'`, so nothing in the types changes. But a consumer then can't tell "read
as Word reads an invalid value" from "rounded as docx4j does".

**Recommendation:** the new kind. docx4j-python CR-018 names the kinds; a counterpart name there is that session's.

### 3.6 What a save writes (N6)

The model holds the rounded integer, so every save writes an integer. A loaded document never reaches the save
normaliser with a decimal, whether `normaliseOnSave` is on or off. The read is like CR-013's on/off reading, where
`yes` is held as `true` and saved `true`. Jason's aim, never writing what the schema rejects, holds by default for
every part this CR loads.

That leaves a caller who sets a non-integer itself (`spacing.line = 12.96 * 20`). Today the marshal throws
`Argument [259.2] must be an integer.`, through jsonix-CR-010's `onUnprintableValue`, which the facade doesn't set.
There are two choices:
- **(a) keep the throw.** A non-integer in an integer property is the caller's bug, and the throw names it;
- **(b) under the normaliser** (on by default), round it half-even and report it to `onNormalised`. A save then never
  fails on it, but the caller's bug goes unseen unless they listen.

**Recommendation: (a).** The normaliser writes Word's reading of documents; it is not there to repair a caller's
arithmetic. core-ts's setters round what they compute.

### 3.7 SSRS 2012's empty attributes (N7, optional)

SSRS 2012's Word renderer writes `<w:sectPr w:rsidRPr="" w:rsidDel="" w:rsidR="" w:rsidSect="">` and
`<w:pgMar ... w:header="" w:footer="" w:gutter=""/>`. Measured as section 2:
- **Here:** the part loads already. The runtime reads `""` as `0` in all three `w:pgMar` values, and keeps the rsids as
  `""`, writing them back as `""`.
- **docx4j:** an empty `w:header` fails its plain unmarshal (`Zero length BigInteger`), so its XSLT runs. That writes
  `0` for `w:header` and `w:footer`, drops `w:gutter`, and drops every empty rsid. A lone `<w:p w:rsidR="">` in a part
  that otherwise loads is kept as `""` by docx4j too, since the XSLT never runs.

So the margins already read as docx4j's do. This package writes `w:gutter="0"` where docx4j writes none, and
`CT_PageMar` requires `w:gutter`, so this package's output is the schema-valid one. The difference is the rsids:
`ST_LongHexNumber` is four bytes of `hexBinary`, so an empty one is schema-invalid. That is the "dodgy content" a save
shouldn't write.

**Recommendation:** ride along, on the save side only. Under the normaliser, leave out an empty `rsid*` attribute (20
attribute properties of `wml`, all strings in the mapping, keyed by name), reported to `onNormalised`. The tree keeps
`""`, as phase C keeps what it read. An rsid is a revision-session marker that Word doesn't draw, so leaving it out
changes no reading. This agrees with docx4j for SSRS's files, and writes less than docx4j for a lone empty rsid.

## 4. Not proposed

- **The runtime's lenient integer parse.** `@docx4j/jsonix` parses an integer with `Number()`:
  - `""` and whitespace give `0`;
  - `0x10`, `0b1` and `0o7` are read as hex, binary and octal;
  - `1e3` and `1.5e2` are read as exponents;
  - `10206.0` is an integer.

  None of these reach `onInvalidValue`. docx4j fails `""`, `1e3` and `0x10` and repairs `10206.0`; Word is unmeasured.
  Making the parse strict, so that these reach the facade's rules, would be a jsonix CR. It should wait for the Word
  probe, since `""` as `0` already gives SSRS's margins docx4j's reading.
- **Other mappings' integer attributes** (section 3.2).
- **docx4j's list.** Extending it to (b), or matching N4 and N7, is docx4j's to file.

## 5. Decisions (for Jason)

- **N1. Where:** the facade's `wordInvalidValue` (CR-013's `onInvalidValue`); a load-time pass; or, as docx4j-python
  proposes for its port, the text kept and rounded by readers and on save (section 3.1). **Recommendation:**
  `wordInvalidValue`.
- **N2. Which attributes:** (a) docx4j's list with its conditions; (b) every integer attribute of Word's own mappings
  (178). **Recommendation:** (b).
- **N3. Rounding:** half-even as docx4j does, or wait for a Word probe. **Recommendation:** half-even now; the probe
  only if a consumer needs Word's half-way answer.
- **N4. Text that is no decimal number:** fail as before (docx4j fails too), or drop the attribute.
  **Recommendation:** fail as before.
- **N5. Reporting:** a new `kind: 'value.rounded'` with a numeric `value`, or reuse `'value.kept.invalid'`.
  **Recommendation:** the new kind.
- **N6. A caller's non-integer on save:** keep the throw, or round it under the normaliser. **Recommendation:** keep
  the throw.
- **N7. SSRS's empty rsids:** leave them out on save under the normaliser, or leave as is. **Recommendation:** leave
  them out on save.

## 6. Tests

- **Fixtures:**
  - `test/fixtures/integer-decimals/cases.json`, the shapes of section 2 and one decimal on each of the 178
    attributes, generated by a script beside it.
  - docx4j's output for each case, generated by a small Java program run against docx4j's jars at a named commit,
    beside its `SOURCE.md`. The program is `MainDocumentPart.unmarshal`, then `XmlUtils.marshaltoString`; it was
    prototyped for section 2.
  - Never edited by hand.
- **docx4j as the oracle:** wherever docx4j loads a case, the facade's marshalled output equals docx4j's, attributes
  compared as sets. Where docx4j fails, the case is listed, strict both ways: rounded here (N2 (b)), read by the
  runtime's `Number()` (section 4), or failing in both (N4). A prototype (the proposed branch as a caller's
  `onInvalidValue`, composed with `wordInvalidValue`) gives:
  - of the 30 cases docx4j loads: 29 identical. The 30th is SSRS's sectPr, whose rsids are kept (N7) and whose gutter
    is `0` (section 3.7);
  - of docx4j's 14 failures: 8 rounded here, 5 read by `Number()`, and 1 (`240,5`) failing in both.
- **Rounding:** the half-way, negative, `-0` and `141.99999999999994` cases.
- **The scope:** each of the 178 rounds. A DrawingML integer (`a:off/@x="12.5"`) still fails, and so does a refused
  value of each other mapping.
- **The record:** `onWordReading`'s `kind`, `value` and `text`. A caller's own `onInvalidValue` takes precedence.
- **The save:** an integer is written. With N7, an empty rsid is left out and reported to `onNormalised`, while the
  tree keeps `""`.
- **Fidelity:** unchanged. Office writes no decimals, and `test/fidelity.mjs` already fails a part that reaches
  `onWordReading`.

## 7. Consumers

- **core-ts:**
  - Parts that failed now load with no change in core-ts, since every facade unmarshal passes `wordInvalidValue`.
  - `ValueKeptWarning.code` gains `'value.rounded'`. Its `valueKept` message reads "read as 259", and its `value` type
    widens to `number`.
  - Its corpus survey would say how many parts newly load.
  - Nothing else.
- **docx4j:** none from this CR. Extending its list, if Jason wants (b) in both ports, is docx4j's.
- **docx4j-python:** keeps the raw string with a `ConverterWarning` (the survey's row 1). Its counterpart is its own.

## 8. Release

- **Changes:** parts that failed now load, `WordReading` gains a kind, and, with N7, a save leaves out empty rsids.
- **Version:** under Jason's rule (minors for a substantive change of API or output), a minor. It can ride with CR-014
  in 0.6.0 (implemented, unreleased).
- **Effort:** half a day with the oracle fixtures, once N1 to N7 are decided.
