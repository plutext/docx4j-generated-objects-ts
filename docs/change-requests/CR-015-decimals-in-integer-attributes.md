# CR-015: A decimal in an integer attribute, read as Word reads it rather than failing the part (`wordInvalidValue`)

**Status:** Proposed 2026-10-26; Word's rule measured (docx4j CR-065's probe, rounds 2 and 3); A2's six forms singly
wait for an optional round 4
**Requested by:** Jason, through the docx4j-portfolio session (docx4j-portfolio-9d), 2026-10-26. The source is the
portfolio's survey of 2026-09-21, `docs/broken_producer_repairs_across_ports.md` section 2, row 1, which found the part
fails here where docx4j repairs it.
**Depends on:** docx4j CR-065 phase 0, the shared Word probe (registry `docx4j/CR-065.0`):
- round 2 (Jason, Word 365, 2026-10-26) measured the rule: truncation;
- round 3 (the same day) measured which text Word reads (N4);
- an optional round 4 would test A2's six forms one at a time.

Nothing new is needed in the runtime: jsonix-CR-010's `onInvalidValue`, through which CR-013 reads Word's values, is in
`@docx4j/jsonix` ^3.8.0 already.
**Counterpart:** docx4j CR-065 (docx4j-25, proposed to Jason 2026-10-26). docx4j today repairs through
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

Proposed: the facade's `wordInvalidValue`, which every facade unmarshal already passes to the runtime (CR-013 section
3.5), reads a decimal in an integer attribute of Word's own WordprocessingML mappings as its integer part, as Word does.
- **The model:** it then holds an integer, so a save writes one.
- **Reporting:** `onWordReading` reports each value read this way.
- **Text Word refuses.** Word reads only `-?digits[.digits]` and `""`. Any other form (`240,5`, `0x10`, `1e3`, `abc`)
  makes it refuse the whole document, with no repair offered (round 3). Such a value that reaches the facade is
  dropped and reported, so the part loads and a save writes a document Word opens (N4). docx4j's CR-065 recommends the
  same.
- **Forms the runtime reads first.** The runtime's `Number()` reads some forms Word refuses, `0x10` as `16`, before the
  facade sees them. A strict parse, opt-in through a jsonix CR, would bring them to the same rule (N8, a second
  phase).

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

  So Word's form is an optional `-`, digits, and optionally `.` and digits; or empty. Which of A2's six Word refuses on
  its own is an optional round 4, whose six files are on the share.

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

**Read by this CR:** Word's form (round 3), with XML whitespace around it trimmed:
- the pattern is `^[ \t\r\n]*-?\d+(\.\d+)?[ \t\r\n]*$`, then truncated (section 3.4);
- trimming follows docx4j's CR-065 D3, unless round 4 finds that Word refuses `" 240.5 "` on its own.

**A form Word refuses that reaches the facade (N4).** Examples are `240,5`, `2.5e0`, `+240.5`, `.5` and `abc`. Word
refuses the whole document for any of them. There are two choices:
- **(a) fail the part**, as 0.5.0 does;
- **(b) drop the attribute, reported.** The kind is `'value.refused'` and the value undefined. The part loads, the
  attribute reads as absent, and a save writes a document Word opens. docx4j's CR-065 recommends the same (its D3).

**Recommendation: (b), with docx4j.** Failing the part loses all of it for one attribute, and Word gives no reading
that failing could match, since it refuses the document. The report says what was dropped and why.

**Forms the runtime reads first (N8).** `@docx4j/jsonix` parses an integer with `Number()`, so some forms Word refuses
never reach `onInvalidValue`:
- `0x10` reads as `16`, `1e3` as `1000` and `1.5e2` as `150`;
- `+240`, `240.` and `" 240 "` read as `240`, though Word refused them in A2; round 4 would say which it refuses on its
  own.

So this package reads documents Word won't open, with values Word never read. By contrast, `""` as `0` and `10206.0` as
`10206` are Word's own readings. There are two choices:
- **(a) leave the runtime as it is.**
- **(b) a jsonix CR for an opt-in strict integer parse** (an unmarshaller option) that refuses anything outside
  `xsd:integer`'s lexical form. The forms above would then reach the facade:
  - Word's own forms, including `""` as `0`, are read as Word reads them;
  - the rest are dropped as in N4 (b);
  - outside Word's mappings, `wordInvalidValue` reads them with `Number()` as today, so nothing changes there.

**Recommendation: (b), as a second phase.** It makes the ports agree, since docx4j's lenient parse drops such forms
too (its D3). But no producer is known to write them, so it needn't hold back the first phase.

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

## 5. Decisions (for Jason)

- **N1. Where:** the facade's `wordInvalidValue` (CR-013's `onInvalidValue`); a load-time pass; or, as docx4j-python
  proposes for its port, the text kept and rounded by readers and on save (section 3.1). **Recommendation:**
  `wordInvalidValue`.
- **N2. Which attributes:** (a) docx4j's list with its conditions; (b) every integer attribute of Word's own mappings
  (178). **Recommendation:** (b).
- **N3. The rule:** truncation toward zero, as Word's saves show (round 2), in all three ports. Half-even (docx4j's
  XSLT) differs from Word in 7 of the probe's values on docx4j's own list. **Recommendation:** truncation.
- **N4. A form Word refuses** (round 3), reaching the facade: (a) fail the part; (b) drop the attribute, reported as
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

- **Fixtures:**
  - Families A1, A6 and B to I and Word's saves of them, as extracted part XML, under `test/fixtures/cr065-probe/`,
    with a `SOURCE.md` naming the share, docx4j-25's builder (docx4j CR-065 phase 0) and Jason's runs. The parts are
    `word/document.xml`, and for I `styles.xml`, `settings.xml` and the header.
  - The refused families A2 to A5 and A7, without saves. Word's answer for them is the refusal itself.
  - `test/fixtures/integer-decimals/`: one plain decimal on each of the 178 attributes, generated by a script beside
    it.
  - Never edited by hand.
- **Word as the oracle:** each case read through the facade from the probe gives what Word's save wrote for it, as in
  `test/word365.mjs`. The known differences are stated, strict both ways:
  - a truncation equal to what lies beneath, which Word's save leaves out (`-0.5`, run `w:sz="22.5"`);
  - `w:gridCol`, which Word recomputes.
- **docx4j's XSLT:** its 7 differences from Word are listed. When docx4j's CR-065 follows round 2, the list empties.
- **The rule:** truncation of positive and negative values, `-0`, `.5` and `141.99999999999994`.
- **The scope:** each of the 178 is read. A DrawingML integer (`a:off/@x="12.5"`) still fails, and so does a refused
  value of each other mapping.
- **The forms:**
  - each refused family's values are dropped and reported as `'value.refused'`, or read by the runtime first (N8,
    stated until phase 2);
  - A1 is truncated and A6's `""` is `0`, as in Word's saves.
- **The record:** `onWordReading`'s `kind`, `value` and `text`. A caller's own `onInvalidValue` takes precedence.
- **The save:** an integer is written. With N7, an empty rsid is left out and reported to `onNormalised`, while the
  tree keeps `""`.
- **Fidelity:** unchanged. Office writes no decimals, and `test/fidelity.mjs` already fails a part that reaches
  `onWordReading`.

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
- **Version:** under Jason's rule (minors for a substantive change of API or output), a minor.
- **Timing:** the first phase (N1 to N7) is measured and can be built now. It can ride with CR-014 in 0.6.0 if Jason
  wants it there. The second phase (N8) waits for a jsonix release, and round 4 can only refine the whitespace and
  sign forms.
- **Effort:** half a day with the Word fixtures, once N1 to N7 are decided.
