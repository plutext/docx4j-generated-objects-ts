# CR-019: `w:stylePaneFormatFilter`'s fifteen attributes, kept on a re-marshal

**Status:** Accepted 2026-11-06, D1 to D4 as recommended (Jason: "CR-019 decisions as recommended; please commit"; section 8).
Proposed 2026-11-05. Nothing is implemented, and the regeneration waits on docx4j's schema change (section 4).
**Requested by:** Jason, 2026-11-05 ("similarly the objects agent its CR"), through core-ts CR-025 D6 (core-ts `5c75758`,
local). He asked for the CR, not the implementation.
**Depends on:** `docx4j/style-pane-format-filter` (registry, planned, raised 2026-10-11; no docx4j CR yet). `modules/`
follows docx4j's `wml.xsd`, so the schema changes there first (section 4).
**Counterparts:**
- docx4j has the same gap. `CTShortHexNumber` holds only `val`, and `CTSettings.getStylePaneFormatFilter()` returns it
  (`javap` on docx4j's built classes, `VERSION_17_3_6` at `b267f24ce`).
- docx4j-python already binds the attributes, with a patch to its own schema copy (`a6bf811`, 2026-09-12; its CR-002
  section 5.7), as `xsd:boolean` with `w:val` optional.
- core-ts is the consumer (section 10).

## 1. Summary

A read and a re-marshal take `w:stylePaneFormatFilter` from 16 attributes to 1, because `wml.xsd` types it
`CT_ShortHexNumber`, which declares only `w:val`. Later editions of the standard type it `CT_StylePaneFilter`.

The decision is to give docx4j's `wml.xsd` that type, with the 15 attributes as `s:ST_OnOff` and `w:val` optional, then
regenerate. Nothing in `src/` changes. A prototype changes exactly 1 of the 341 fidelity parts, the intended one
(section 6).

**Priority: low.** The form is lost but not the state, since the 15 attributes repeat `w:val`'s bits in all three files
that have them (section 2.2). core-ts is protected meanwhile by its CR-025.

The schema comparison also found five other elements with the same kind of gap, two of them settings elements. They are
outside this CR, and section 7 records them and what comes next (D4).

## 2. What is there today

### 2.1 The loss

Measured at 0.7.0, `@docx4j/jsonix` 3.11.0, on `test/fixtures/fidelity/NumberingImplicitNumId.docx/word/settings.xml`:
the element becomes `<w:stylePaneFormatFilter w:val="3F01"/>`, and `onUnexpectedAttribute` fires for each of the 15.

### 2.2 The attributes repeat `w:val`

The standard's table for `w:val` defines 15 bits, and each matches one attribute by its description:

| Bit | Attribute | Bit | Attribute |
|---|---|---|---|
| `0x0001` | `allStyles` | `0x0100` | `directFormattingOnRuns` |
| `0x0002` | `customStyles` | `0x0200` | `directFormattingOnParagraphs` |
| `0x0004` | `latentStyles` | `0x0400` | `directFormattingOnNumbering` |
| `0x0008` | `stylesInUse` | `0x0800` | `directFormattingOnTables` |
| `0x0010` | undefined | `0x1000` | `clearFormatting` |
| `0x0020` | `headingStyles` | `0x2000` | `top3HeadingStyles` |
| `0x0040` | `numberingStyles` | `0x4000` | `visibleStyles` |
| `0x0080` | `tableStyles` | `0x8000` | `alternateStyleNames` |

All three files with the attributes hold `3F01`, and their attributes give `3F01` when computed from this table.

**Not measured:** whether Word reads `w:val` when the attributes are absent, and which it believes when they disagree.

### 2.3 Who writes which form

Every Word package under core-ts's `test/`, docx4j-python's `samples/`, and docx4j's `sample-docs`,
`docx4j-samples-resources` and `docx4j-core-tests/src` that has the element:

| Package | `AppVersion` | Form |
|---|---|---|
| `NumberingImplicitNumId.docx` | 14.0000 | 16 attributes |
| `same_as_previous_numbering.docx` | 14.0000 | 16 attributes |
| `toc.docx` | 14.0000 | 16 attributes |
| `jaxb-binder-issue.docx` | 12.0000 | `w:val` alone |
| `pandoc-code.docx` | 12.0000 | `w:val` alone |

No file here shows what Word 365 writes. The element is in 1 of this repository's 433 fixture XML files, so
`test/fixtures/fidelity/NumberingImplicitNumId.docx/SOURCE.md` is wrong to say every other document writes it with
`w:val` alone (section 9).

## 3. An edition difference

Yes, this is one. docx4j's `wml.xsd` keeps the first edition's `CT_ShortHexNumber` here. ISO/IEC 29500-4:2016 transitional
has `CT_StylePaneFilter`, with the 15 attributes as `s:ST_OnOff` and an optional `w:val`. Strict has the 15 and no
`w:val`. docx4j-python's CR-002 places the type's arrival at the second edition. Only the 2016 text and its strict
counterpart were compared here.

docx4j has taken this same step before: `CT_TblLook`, `w:val` beside named on/off attributes, came "from transitional 4ed"
in `d77736089` (2013), typed `s:ST_OnOff`, with `w:val` optional.

## 4. The route

The change has to start in docx4j. Java's model (`docx4j-generated-objects`, xjc over `xsd/ROOT.xsd`) and `modules/`
(`generate.sh` over the same file) both come from `xsd/wml/wml.xsd`, and `modules/` is never edited by hand. One schema
change serves both, and docx4j-python can retire its patch at its next refresh:

```xml
<xsd:complexType name="CT_StylePaneFilter">
  <xsd:attribute name="allStyles" type="s:ST_OnOff"/>
  <!-- customStyles ... alternateStyleNames, as section 2.2 -->
  <xsd:attribute name="val" type="ST_ShortHexNumber"/>
</xsd:complexType>
```

`CT_Settings/stylePaneFormatFilter` then takes that type. In docx4j, `getStylePaneFormatFilter()`'s return type changes;
no hand-written Java names it.

`CT_ShortHexNumber` stays for `w:stylePaneSortMethod`. ISO gives that element `CT_StyleSort` with names (`name`,
`priority`, ...), but `val` is a plain string on both ports, and `w:val="name"` and `w:val="0000"` both round-trip here.
Nothing is lost, so nothing changes.

## 5. The attribute type and `w:val` (D2, D3)

Both types were prototyped:

| | `s:ST_OnOff` (decided) | `xsd:boolean` (docx4j-python) |
|---|---|---|
| The fixture writes | `w:allStyles="1"`, as Word wrote it | `w:allStyles="true"` |
| Held as | `'1'`, one of six spellings | `true` |
| `on` | kept as `on` | written `true` |
| A refused value (`maybe`, empty) | dropped and reported | read as on and reported (CR-013) |
| Agrees with | ISO, and `CT_TblLook` in the same file | docx4j-python |

D2 is `s:ST_OnOff`. Nothing reads these attributes, and the point of the change is that they survive, so
keeping the source's spelling matters more than a convenient declaration. No new form reaches Word either.

This departs from the registry's note that docx4j should follow docx4j-python. A `boolean` is the better declaration if
anyone comes to read them.

`w:val` stays and becomes optional, as in ISO and `CT_TblLook`. It is not derived from the attributes, nor they from it: a
save writes what it was given. A caller who changes one and not the other writes a file whose reading is unmeasured
(section 2.2).

## 6. The prototype

Measured in a scratch copy of 0.7.0's `dist/` and `modules/`, with the type patched into `org_docx4j_wml` by hand. The real
output is the regeneration's.

- **Reading and writing:** the fixture reports nothing dropped, and all 16 attributes are written. An element with no
  `w:val` reads and writes.
- **`test/lib/marshal-hashes.mjs`:** 1 of 341 lines differs for either type, `NumberingImplicitNumId.docx/word/settings.xml`.
- **`test/fidelity.mjs`:** it fails with its 15 `KNOWN_MISSING_ATTRIBUTES` entries, each "recorded as not bound, but no part
  loses it any more", and nothing else.

## 7. The same kind of gap elsewhere (D4)

`wml.xsd` compared with the ISO 2016 transitional schema, type by type, then each candidate tried through the facade at
0.7.0. The held count is a text search of the Word packages in section 2.3's trees, which are test fixtures rather than a
sample of the world.

| Element | In ISO, not in docx4j | Through the facade | Held |
|---|---|---|---|
| `w:documentProtection` (settings) | `algorithmName`, `hashValue`, `saltValue`, `spinCount` | all 4 dropped; `w:edit` and `w:enforcement` stay | element in 1 package, without them |
| `w:writeProtection` (settings) | the same 4 | all 4 dropped | element in none |
| `w:cnfStyle` | 12 on/off attributes beside `w:val`, this element's shape | all 12 dropped | element in 2 packages, `w:val` alone |
| `w:numFmt` | `w:format`, and the value `custom` | both go: `<w:numFmt/>` | none with `w:format` |
| `w:tc` | `w:id` | dropped | none |

- **`w:documentProtection` matters most.** A file using ISO's attributes keeps its restriction and loses the password hash
  behind it, so the loss is one of state, unlike section 2.2's. The older `w:hash`, `w:salt` and `w:crypt*` attributes do
  round-trip (measured).
- **`w:numFmt`:** an `mc:AlternateContent` holding one in a `w:lvl` is dropped whole, as CR-025's host list says of `w:lvl`.
  That is believed to be Word's form for a custom format, but no file here shows it.
- **Not attribute gaps:** ISO also has child elements docx4j lacks (`CT_TcMar/start`, `CT_TblBorders/start`, `CT_SdtPr/label`
  and others). Those are CR-026's kind of case and were not tried here.

So core-ts's "only real loss" holds for what its corpus contains, not for what the model can lose. How often these occur
isn't known. The portfolio's catalogue holds 2,274,603 settings and 1,618,539 numbering blobs, which can count everything
above except `w:cnfStyle` and `w:tc`, whose elements live in bodies. Its scan doesn't record these attributes today, so D4 has that count made first.

## 8. Decisions

| | Question | Decided 2026-11-06, as recommended |
|---|---|---|
| D1 | The route: docx4j's `wml.xsd`, then regenerate? | Yes. It needs a small docx4j CR under `docx4j/style-pane-format-filter`. |
| D2 | The attributes' type? | `s:ST_OnOff` (section 5). docx4j-python's view is still wanted, since this departs from its patch. |
| D3 | `w:val`? | Stays, optional, unsynchronised (section 5). |
| D4 | Section 7's gaps? | Keep them out of this CR. Count the settings and numbering blobs first, then take `w:documentProtection` and `w:writeProtection` into the docx4j CR if they occur. |

## 9. Tests and docs, at the regeneration

- **`test/fidelity.mjs`:** remove `STYLE_PANE_FILTER` and its 15 entries. The table's own strictness forces this (section 6),
  and `KNOWN_MISSING_ATTRIBUTES` goes from seventeen to two.
- **`test/smoke.mjs`:** a block holding the 16 attributes read and written in the source's spelling, an element with no
  `w:val`, and `w:stylePaneSortMethod` still `CTShortHexNumber`.
- **The gate:** the marshalled hashes differ in that one part only, and `typecheck` compiles the regenerated declarations.
- **Docs:** `CLAUDE.md`'s seventeen and its list, and the `SOURCE.md` sentence (section 2.3). CR-004 section 11 is a dated
  record and stays.

## 10. Consumers

- **core-ts:** it doesn't use the property (no hits in `src/` or `test/`). CR-025's one real file stops reporting, so
  its 15 dropped attributes go to 0 and a `part.skipped` test needs another source. That file's settings hash changes.
- **`@docx4j/generated-objects-ts` callers:** `stylePaneFormatFilter` becomes `CTStylePaneFilter`, so reading `.val` gives
  `string | undefined`. A value built with the old `TYPE_NAME` is written with `xsi:type="w:CT_ShortHexNumber"` (measured).
- **docx4j-python:** D2 departs from its patch, which it can retire at its next refresh.

The regeneration goes out in the next release, with core-ts's go as usual. Patch or minor is Jason's call: one property's
declared type and one element's output change, and 0.3.2 was a patch for a change of that kind.
