# CR-016: Integer attributes within Word's range: reported on read and on save, refused by the builders

**Status:** Proposed 2026-10-26; G1 to G6 for Jason
**Requested by:** Jason, 2026-10-26 ("write the guard CR"). He raised it through the portfolio session, relayed from
docx4j-python-cd: Word 365 stores twips in 16 bits and wraps an integer too large to fit, so every engine should guard
against code setting a measure Word can't hold.
**Depends on:** docx4j CR-066 phase 0, round 5 of the shared Word probe (Jason, Word 365, 2026-10-26; done). Nothing new
in the runtime: the read report goes through jsonix-CR-008's unmarshal-side `onElement` (3.5.0), and the save report
through jsonix-CR-011's `onAttributeValue` (3.8.0).
**Counterparts:**
- docx4j CR-066 (docx4j-25, proposed, D1 to D8 for Jason) reports at marshal and unmarshal in the print and parse
  methods bound to each integer simple type, and guards `TblFactory`;
- docx4j-python CR-022 raises at every setter of its content API;
- core-ts's content API is its own (section 8).

All three keep a file's value as written.

## 1. Summary

Word stores each WordprocessingML measure in a fixed width and reads a value outside it modulo that width, sometimes
then clamped. It neither refuses nor repairs the document. In CR-065's round 2, `w:ind/@w:left="7205001"` read as
`-3959`, and `w:spacing/@w:before="1205003"` in a header became 17.6 inches before every paragraph. Round 5 measured the
width attribute by attribute (section 2). It is 16 bits signed or unsigned, 8 and 5 bits for a border, 32 bits for a
page number, with clamps on page width and tab stops, and the rule carried in the sign for row height and exact line
spacing.

Today this package holds such a value as written, writes it as held, and its builders take any number (section 3).

Proposed:
- **A table** of Word's storage per attribute, exported from `helpers/wml` with a function giving what Word reads for a
  value (G5, G6).
- **A report:** on read and on save, through a new callback `onOutOfRange`, for each value Word would read differently
  (G1 to G3). The value is never rewritten, and the tree keeps it as written.
- **A guard:** the builders that take a measure (`tbl`, `tr` and `tc` widths, `applyRunOptions`'s `size`) throw a
  `RangeError` for a value Word can't hold (G4).

docx4j's report knows only the simple type, so it can be exact only where a type's attributes agree (its D1 and D7,
amended). This package's adapters know the class and the property, so its report is exact for every measured
attribute.

## 2. What Word does (round 5; docx4j CR-066 section 3.1)

The probe is on the share at `Office 2016/__tmp/cr066/`: 190 cases in `cases.txt`, over 18 attributes (21 counting a width's `w:type` and the line rule), in six group
files (`groups/cr066-{P,B,R,T,S,N}.docx`), each opened clean, with no repair, and saved by Word. Word's saves are the
oracle. Mapped to this package's classes:

| Attribute (class, property) | Word's storage, from the saves |
|---|---|
| `w:ind` `left`, `right`, `start`, `end` (`PPrBase.Ind`); the run's `w:spacing` (`CTSignedTwipsMeasure.val`); `w:position` (`CTSignedHpsMeasure.val`) | int16, wraps |
| `w:tblW`, `w:tcW`, `w:tblInd` `@w:w` (`TblWidth.w`), under `dxa` and under `pct` | int16, wraps |
| `w:pgMar/@w:top` (`SectPr.PgMar.top`) | int16, wraps |
| `w:spacing/@w:before` (`PPrBase.Spacing.before`) | uint16, wraps; `65535` is written `-1` |
| `w:sz` (`HpsMeasure.val`); text scale `w:w` (`CTTextScale.val`, schema 0 to 600) | uint16, wraps; no clamp at the UI limits |
| `w:tab/@w:pos` (`CTTabStop.pos`) | int16, wraps, then clamped to ±31680 |
| `w:pgSz/@w:w` (`SectPr.PgSz.w`) | uint16, wraps, then clamped to 144 to 31681 |
| `w:pgSz/@w:h` (`SectPr.PgSz.h`) | uint16, wraps, then at least 144 |
| `w:spacing/@w:line` (`PPrBase.Spacing.line`) | under `auto`, int16, wraps; under `exact`, stored as the negative, so the wrapped sign re-derives the rule (`exact 65535` reads `atLeast 1`) |
| `w:trHeight/@w:val` (`CTHeight.val`) | int16, the sign is the rule (negative is `exact`): `32768` reads `exact 32768`, and `65536` is dropped |
| a border's `@w:sz` (`CTBorder.sz`) | 8 bits, wraps (mod 256); no clamp at the UI's 48 |
| a border's `@w:space` (`CTBorder.space`) | 5 bits, wraps (mod 32) |
| `w:pgNumType/@w:start` (`CTPageNumber.start`) | int32, wraps |

No value was refused. A negative in an attribute the schema calls unsigned is accepted (`w:sz="-1"` reads `65535`).

The widths are those of Word's binary format (DOC's `PAP`, `CHP`, `BRC`, `SEP`), so the table is unlikely to move
between Word versions.

**Unmeasured:** the other ST_TwipsMeasure attributes: `w:gridCol`, the other page margins, `hanging`, `firstLine`,
`after`, frame and drawing distances. Of the 44 twips attributes of Word's own mappings, derived from docx4j's xsd and
cited in CR-066 section 2, ST_TwipsMeasure's are mixed:
- `before` is uint16;
- `trHeight` is int16 with a rule;
- `pgSz` is clamped.

So for an unmeasured one only the certain range holds. A value outside -32768 to 65535 changes in every measured
attribute.

One correction to CR-066's table: `w:pgMar/@w:top` (and `bottom`) is `ST_SignedTwipsMeasure` in docx4j's `wml.xsd`, as
in ECMA-376, not `ST_TwipsMeasure`. Word's int16 is its schema type's sign.

## 3. What this package does today (measured 2026-10-26, at 05d87c5)

- **Read:** an out-of-range integer is a valid integer, so it is held as written, with nothing reported: a tree says
  7205001 where Word shows -3959.
- **Save:** written as held.
- **The builders** take any number, and the runtime's print checks only that it is an integer:
  - `applyRunOptions({ size: 1e6 })` writes `w:sz="2000000"`, and `size: -5` writes `w:sz="-10"`;
  - `tbl(rows, { width: 7205001 })` writes `w:tblW`, `w:gridCol` and `w:tcW` as `7205001`, and `width: -100` writes
    `w:gridCol w:w="-100"`;
  - a non-integer, `NaN` or `Infinity` throws only at marshal.

## 4. Proposal

### 4.1 The table (G5, G6)

`helpers/wml` exports Word's storage per attribute, keyed by `TYPE_NAME` and property, as section 2 gives it:
- **an entry:** the width and sign, plus a clamp (tab stops, page width and height) or a rule (`line` with `lineRule`,
  `trHeight` with `hRule`);
- **a function** of the type, the property, the value and, for a rule, the owner, giving what Word reads: the value it
  writes back, and for a rule the rule too. It returns `undefined` for an attribute not in the table.
- **in range** means Word reads the value as written. `before`'s `-1`, Word's own spelling of `65535`, is in range.
- **Unmeasured ST_TwipsMeasure attributes (G6):** the certain range only. A value outside -32768 to 65535 is reported,
  and the reading is given as unknown, not guessed.
- **Not in the table:** `ST_DecimalNumber`'s ids and counts, the `*Chars` and `*Lines` attributes, w14's DrawingML-style
  units, and every other mapping. Their width is unmeasured, and EMUs are 64-bit.

core-ts's content API can use the same table at its setters, so the two TypeScript packages refuse alike.

### 4.2 Read: reported, held as written (G2)

The facade's unmarshal functions already compose an `onElement` (CR-011). For each typed object whose `TYPE_NAME` has
entries, it checks those properties, a map lookup per element. A value Word reads differently is reported, and the tree
keeps it as written. Emulating Word's reading would swap one wrong number for another and lose the evidence, as
CR-066's D3 and docx4j-python agree.

### 4.3 Save: reported, never rewritten (G3)

The marshal adapter (`onAttributeValue`, jsonix-CR-011) knows the class, the property and the owner, so the check is
exact at save too. It returns `undefined`, writing the value as held: a failed save loses more than a wrong measure,
and a rewrite writes a number nobody asked for (CR-066 D2). The check runs whether or not `normaliseOnSave` is on,
since it changes nothing written. It catches what code set between read and save.

### 4.4 The channel (G1)

A new callback, `onOutOfRange(report)`, on `UnmarshalOptions` and `MarshalOptions`, called for each value only when the
caller passes it. That is this package's rule for reporting: no reporting callback by default. The report gives:
- `element`, `attribute` and, on save, `path` (as `Normalisation`'s);
- `value`, the number held and written;
- `wordReads`, what Word reads (with the rule where one applies), or `undefined` for an unmeasured attribute outside
  the certain range;
- `storage`, for the message (`int16`, `uint16, clamped to 144..31681`, ...).

`onWordReading` and `onNormalised` are not used. Their records say what the tree holds instead of the text, and what a
save rewrote, and here neither changes.

### 4.5 The builders refuse (G4)

The four inputs where a caller's number becomes a measure throw a `RangeError` naming Word's range, before anything is
built:
- `tbl`'s `widths` and `width`, `tr`'s `widths` and `tc`'s `width`: each width and the table's total (`w:tblW`) within
  0 to 32767 twips (`TblWidth.w` is int16);
- `applyRunOptions`'s `size`: finite, and its half-points within 0 to 65535 (`w:sz` is uint16).

A negative is refused there, though Word keeps one: a width or a font size can't be negative, so a negative from a
caller is a mistake. Non-integers keep today's rounding (`size`) or the marshal's throw (widths).

## 5. Not proposed

- **Emulating Word's reading** on read, or rewriting on save (CR-066 D2, D3).
- **Word's UI limits** (1638 pt, 22 in, 600 %, a 6 pt border). Word keeps a value between a UI limit and its storage
  limit, so it is not reported. The two clamps Word applies on read (tab stops, page width) are in the table (CR-066 D4).
- **A whole-tree walker.** The read report covers what was loaded, and the save report what was set since. A consumer
  with its own walk can call the table's function.
- **The other measures a caller sets in this package:** `inlinePicture`'s `cx`/`cy` are DrawingML EMUs, and content
  control ids are decimal numbers, not measured.

## 6. Decisions (for Jason)

- **G1. The channel:** (a) a new `onOutOfRange` callback on unmarshal and marshal; (b) new kinds on `onWordReading` and
  `onNormalised`. **Recommendation:** (a). The existing records describe a change to the value, and here there is none.
- **G2. Read:** (a) reported per attribute through `onElement`, held as written; (b) not reported on read.
  **Recommendation:** (a). It is where a loaded document's defect shows, and costs one lookup per element.
- **G3. Save:** (a) reported per attribute through `onAttributeValue`, never rewritten, with or without
  `normaliseOnSave`; (b) not reported on save. **Recommendation:** (a).
- **G4. The builders:** (a) a `RangeError` outside Word's storage range, and for a negative width or size; (b) the
  storage range only, negatives allowed where Word keeps them; (c) nothing. **Recommendation:** (a). A width outside
  16 bits, or a negative one, is always the caller's mistake, as CR-066's D5 says of `TblFactory`.
- **G5. The table:** (a) exported from `helpers/wml` with its reading function, for core-ts's setters; (b) internal.
  **Recommendation:** (a).
- **G6. Unmeasured ST_TwipsMeasure attributes:** (a) the certain range (outside -32768 to 65535), as CR-066 does; (b)
  assume uint16; (c) nothing until measured. **Recommendation:** (a). It never reports a value Word keeps, and a later
  round can narrow it.

## 7. Tests

- **Fixtures:**
  - round 5's six group files and Word's saves of them, as extracted part XML, under `test/fixtures/cr066-probe/`;
  - `cases.txt`;
  - a `SOURCE.md` naming the share and docx4j CR-066 section 3.1;
  - never edited by hand.
- **Word as the oracle** (`test/word365-ranges.mjs`):
  - for each of the 190 cases, the table's reading of the written value equals what Word's save wrote, with the rule
    for `line` and `trHeight`;
  - the read report fires for exactly the cases whose save differs;
  - `w:gridCol`, which Word recomputes, is stated.
- **The certain range:** an unmeasured ST_TwipsMeasure attribute (`w:spacing/@w:after`, `w:gridCol`) is reported
  outside -32768 to 65535 and not inside it.
- **Smoke:**
  - the channel on read and save, with `path` and `wordReads`;
  - nothing rewritten, and the tree unchanged;
  - reported with `normaliseOnSave: false`;
  - not called when not passed;
  - each builder's `RangeError` and its message, and the values at each boundary accepted.
- **Fidelity:** a part that reaches `onOutOfRange` fails, as for `onWordReading` and `onNormalised`. Word writes nothing
  out of its own range.
- `test/nodenext/consumer.mts`: the callback, the report type and the table's types.

## 8. Consumers

- **core-ts:** its content API's setters can refuse with the exported table, as docx4j-python's CR-022 does at its
  setters. It can also map `onOutOfRange` to part warnings beside its `ValueKeptWarning`. That is its own CR.
- **docx4j:** CR-066 is the counterpart. The one correction (section 2) is told to docx4j-25.
- **docx4j-python:** CR-022. It refuses a negative in a schema-unsigned attribute at its setters; this CR refuses only
  where its own builders take a measure (G4).

## 9. Release

- **Changes:** new API (the table, its function, `onOutOfRange`), and the builders throw where they wrote a wrong
  number before.
- **Version:** under Jason's rule, a minor. It can ride with CR-014 and CR-015 in 0.6.0, or follow.
- **Effort:** a day, with the fixtures.
