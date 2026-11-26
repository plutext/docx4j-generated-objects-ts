# CR-019: `w:stylePaneFormatFilter`'s fifteen attributes, kept on a re-marshal

**Status:** Implemented 2026-11-06, D1 to D3 as decided (Jason: "do the real regeneration"), on docx4j CR-078 (`f960df59e`), released
in 0.8.0 (2026-11-08). Accepted 2026-11-06, D1 to D4 as recommended (Jason: "CR-019 decisions as recommended; please commit"; section 8).
Proposed 2026-11-05.
**Requested by:** Jason, 2026-11-05 ("similarly the objects agent its CR"), through core-ts CR-025 D6 (core-ts `5c75758`,
local).
**Depends on:** `docx4j/style-pane-format-filter`, done as docx4j CR-078 phase 1 (`f960df59e` on `VERSION_17_3_6`, unpushed when
this was written). `modules/` follows docx4j's `wml.xsd`, so the schema changed there first (section 4).
**Counterparts:**
- docx4j had the same gap: `CTShortHexNumber` holds only `val`, and `CTSettings.getStylePaneFormatFilter()` returned it
  (`javap` on docx4j's built classes, `VERSION_17_3_6` at `b267f24ce`). Its CR-078 closes it with the same schema change.
- docx4j-python already binds the attributes, with a patch to its own schema copy (`a6bf811`, 2026-09-12; its CR-002
  section 5.7), as `xsd:boolean` with `w:val` optional.
- core-ts is the consumer (section 10).

## 1. Summary

A read and a re-marshal took `w:stylePaneFormatFilter` from 16 attributes to 1, because `wml.xsd` typed it
`CT_ShortHexNumber`, which declares only `w:val`. Later editions of the standard type it `CT_StylePaneFilter`.

docx4j's `wml.xsd` now has that type, with the 15 attributes as `s:ST_OnOff` and `w:val` optional, and `modules/` is regenerated
from it. Nothing in `src/` changes. Five generated files change, and exactly 1 of the 341 fidelity parts, the intended one
(section 6).

**Priority: low.** For anything Word writes, the form was lost but not the state, since Word's `w:val` and attributes always
agree (section 2.2). A file where they disagree did change its reading, because dropping the attributes left `w:val` alone for
Word to read. core-ts was protected meanwhile by its CR-025.

The schema comparison also found five other elements with the same kind of gap, two of them settings elements. They are
outside this CR, and section 7 records them and what comes next (D4).

## 2. What was there

### 2.1 The loss

Measured at 0.7.0, `@docx4j/jsonix` 3.11.0, on `test/fixtures/fidelity/NumberingImplicitNumId.docx/word/settings.xml`:
the element became `<w:stylePaneFormatFilter w:val="3F01"/>`, and `onUnexpectedAttribute` fired for each of the 15.

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

**What Word reads** (docx4j CR-078 section 1: Word 16.0.20326.20232 through the bridge, 64 purpose-built files, 84 saves, no
dialogs). Word holds one value, and `w:val` and the attributes agree in every save.

- **`w:val` counts only with no attribute beside it.** Then each bit reads as this table has it, and `0x0010`, which has no
  attribute, goes: `0010` saves as `0000`, and `FFFF` as `FFEF`.
- **With any attribute there, `w:val` is ignored,** and an attribute left out reads as on. `w:val="3F01" w:allStyles="0"` reads
  `FFEE`, not `3F00`. An empty element reads `FFEF`.
- **Attribute order and compatibility modes 12 and 14 change nothing.**

A save shows the value Word holds. That the Styles pane uses it is inferred.

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

Word 365 writes all sixteen, in every compatibility mode (docx4j CR-078). The element is in 1 of this repository's 433 fixture
XML files, so `test/fixtures/fidelity/NumberingImplicitNumId.docx/SOURCE.md` was wrong to say every other document writes it with
`w:val` alone (section 9).

## 3. An edition difference

Yes, this is one. docx4j's `wml.xsd` kept the first edition's `CT_ShortHexNumber` here. ISO/IEC 29500-4:2016 transitional
has `CT_StylePaneFilter`, with the 15 attributes as `s:ST_OnOff` and an optional `w:val`. Strict has the 15 and no
`w:val`. docx4j-python's CR-002 places the type's arrival at the second edition. Only the 2016 text and its strict
counterpart were compared here.

docx4j has taken this same step before: `CT_TblLook`, `w:val` beside named on/off attributes, came "from transitional 4ed"
in `d77736089` (2013), typed `s:ST_OnOff`, with `w:val` optional.

## 4. The route

The change had to start in docx4j. Java's model (`docx4j-generated-objects`, xjc over `xsd/ROOT.xsd`) and `modules/`
(`generate.sh` over the same file) both come from `xsd/wml/wml.xsd`, and `modules/` is never edited by hand. One schema
change, docx4j CR-078 phase 1, serves both, and docx4j-python can retire its patch at its next refresh:

```xml
<xsd:complexType name="CT_StylePaneFilter">
  <xsd:attribute name="allStyles" type="s:ST_OnOff"/>
  <!-- customStyles ... alternateStyleNames, as section 2.2 -->
  <xsd:attribute name="val" type="ST_ShortHexNumber"/>
</xsd:complexType>
```

`CT_Settings/stylePaneFormatFilter` takes that type. In docx4j, `getStylePaneFormatFilter()`'s return type changes;
no hand-written Java names it (confirmed there).

`CT_ShortHexNumber` stays for `w:stylePaneSortMethod`. ISO gives that element `CT_StyleSort` with names (`name`,
`priority`, ...), but `val` is a plain string on both ports, and `w:val="name"` and `w:val="0000"` both round-trip here.
Nothing is lost, so nothing changes.

## 5. The attribute type and `w:val` (D2, D3)

Both types were prototyped here, through jsonix:

| | `s:ST_OnOff` (decided) | `xsd:boolean` |
|---|---|---|
| The fixture writes | `w:allStyles="1"`, as Word wrote it | `w:allStyles="true"` |
| Held as | `'1'`, one of six spellings | `true` |
| `on` | kept as `on` | written `true` |
| A refused value (`maybe`, empty) | dropped and reported | read as on and reported (CR-013) |
| Agrees with | ISO, and `CT_TblLook` in the same file | docx4j-python's patch |

D2 is `s:ST_OnOff`. Nothing reads these attributes, and the point of the change is that they survive, so
keeping the source's spelling matters more than a convenient declaration. No new form reaches Word either.

This departs from the registry's note that docx4j should follow docx4j-python, which has no objection (2026-11-06): none of
its hand-written code or tests read these attributes. Its own `xsd:boolean` writes `1`/`0`, so the table's `true` is
jsonix's spelling, not docx4j-python's. A `boolean` is the better declaration if anyone comes to read them.

`w:val` stays and is optional, as in ISO and `CT_TblLook`. It is not derived from the attributes, nor they from it: a save
writes what it was given. So on a loaded file with the attributes, a caller who sets `val` alone changes nothing Word reads
(section 2.2), where at 0.7.0, with the attributes dropped, the same call did. Jason kept D3 knowing this (docx4j CR-078), and
`README.md` says so.

## 6. The regeneration

Compiler `a314237`, docx4j `f960df59e`, 2026-11-06. It ran first as a dry run into scratch copies, from three trees of `xsd/`, then
for real, byte-identical to the third:

| `xsd/` from | Shows | Generated files that differ |
|---|---|---|
| `ac97b56c5`, behind 0.7.0's `modules/` | the tooling reproduces them | 0 |
| `f960df59e`'s parent | docx4j's changes since (375 `wml.xsd` lines of annotations and `ST_DecimalNumber`) reach nothing here | 0 |
| `f960df59e` | this change alone | 5, all `org_docx4j_wml`: 161 lines in, 3 out |

- **Reading and writing:** the fixture reports nothing dropped, all 16 attributes are written as spelled, and an element with
  no `w:val` reads and writes.
- **`test/lib/marshal-hashes.mjs`:** 1 of 341 lines differs, `NumberingImplicitNumId.docx/word/settings.xml`, as a prototype by
  hand had shown for either type.
- **`test/fidelity.mjs`:** before section 9's change it failed with its 15 `KNOWN_MISSING_ATTRIBUTES` entries, each "recorded as
  not bound, but no part loses it any more", and nothing else.
- **`npm test` and `typecheck`:** both pass on Node 18, 20 and 22 with `@docx4j/jsonix` 3.11.0, and on Node 22 with 3.12.0 from
  npm, the hashes the same.

## 7. The same kind of gap elsewhere (D4)

`wml.xsd` compared with the ISO 2016 transitional schema, type by type, then each candidate tried through the facade at
0.7.0, and again at `71ef2e1` (CR-022's commit) with the same result. The portfolio counted them on 12 November 2026 on the
gather host, 3,242,806 kept documents of 3,722,083: the settings and numbering parts by their blobs, and the main part,
headers, footers and notes by a scan of each kept document.

| Element | In ISO, not in docx4j | Through the facade | Kept documents |
|---|---|---|---|
| `w:documentProtection` (settings) | `algorithmName`, `hashValue`, `saltValue`, `spinCount` | all 4 dropped; `w:edit` and `w:enforcement` stay | 132 with them, of 146,005 with the element; 49,906 hold the older `w:hash` form |
| `w:writeProtection` (settings) | the same 4 | all 4 dropped | none with them, of 2,861; 256 hold the older form |
| `w:cnfStyle` | 12 on/off attributes beside `w:val`, this element's shape | all 12 dropped | 41,958 (6,363,324 elements), each attribute in about 41,800 |
| `w:numFmt` | `w:format`, and the value `custom` | both go: `<w:numFmt/>`. In an `mc:AlternateContent` in a `w:lvl`, the whole `mc:AlternateContent` goes, Fallback too | 2,111 with `w:format`, 2,115 with `custom`; 2,098 with an `mc:AlternateContent` in a `w:lvl`, every one `Requires="w14"` and holding a `w:numFmt` |
| `w:tc` | `w:id` | dropped | none |

- **`w:documentProtection` is the gravest per file.** A file using ISO's attributes keeps its restriction and loses the password
  hash behind it, so the loss is one of state, unlike section 2.2's. It's rare: 132 documents, against 49,906 using the older
  `w:hash`, `w:salt` and `w:crypt*` attributes, which round-trip (measured). docx4j's two types share its `AG_Password`, so one
  schema change would cover `w:writeProtection` too, though nothing counted uses ISO's form there.
- **`w:numFmt` is the largest loss of state:** the count shows the `mc:AlternateContent` form, `w:numFmt w:val="custom"` with
  `w:format` in a `w14` Choice, in 2,098 documents. A save through the facade drops it whole, as CR-025's host list says of
  `w:lvl`, leaving the level with no `w:numFmt`, which ISO reads as decimal. Keeping it takes `w:format`, `custom`, and an
  `mc:AlternateContent` in a `w:lvl`.
- **`w:cnfStyle` is the most frequent.** Whether `w:val` repeats the attributes, as `w:stylePaneFormatFilter`'s do (section
  2.2), wasn't counted. Where it does, the loss is of form.
- **Not attribute gaps:** ISO also has child elements docx4j lacks (`CT_TcMar/start`, `CT_TblBorders/start`, `CT_SdtPr/label`
  and others). Those are CR-026's kind of case and were not tried here.

So core-ts's "only real loss" holds for what its corpus contains, not for what the model can lose. By D4, `w:documentProtection`
goes to docx4j, since its ISO form occurs. No docx4j CR covered any of the four, open or planned (docx4j, 18 November 2026, at
`6405804c8`; CR-078's section 3 left the protection pair for this count), so Jason had docx4j draft one (2026-11-18: "have it
draft it"): docx4j CR-088, proposed the same day and awaiting his review. After a measuring phase it has three parts: `AG_Password`'s
four attributes, which cover both protection elements; `CT_Cnf`'s twelve, `w:val` optional and not kept in step with them; and
`CT_NumFmt`'s `w:format` with `custom` (and ISO's `bahtText` and `dollarText`), with an `mc:AlternateContent` admitted in a `w:lvl`
as a choice with `w:numFmt`. `w:tc`'s `w:id` stays out. This package follows it with a regeneration, as it followed CR-078, with
nothing in `src/` depending on these types. The first two parts change attributes only; the third changes `CT_Lvl`'s content, so a
level holding the branch has no `numFmt`, which core-ts's numbering would read.

## 8. Decisions

| | Question | Decided 2026-11-06, as recommended |
|---|---|---|
| D1 | The route: docx4j's `wml.xsd`, then regenerate? | Yes, as a small docx4j CR under `docx4j/style-pane-format-filter`: its CR-078. |
| D2 | The attributes' type? | `s:ST_OnOff` (section 5). docx4j-python has no objection, though this departs from its patch. |
| D3 | `w:val`? | Stays, optional, unsynchronised (section 5). |
| D4 | Section 7's gaps? | Keep them out of this CR. Count the settings and numbering blobs first, then take `w:documentProtection` and `w:writeProtection` into the docx4j CR if they occur. Counted 2026-11-12 (section 7): `w:documentProtection`'s ISO form occurs, in 132 documents. No docx4j CR covered it, so docx4j drafted CR-088 for it, with `w:cnfStyle` and `w:numFmt` (Jason 2026-11-18; proposed) |

## 9. Tests and docs

- **`test/fidelity.mjs`:** `STYLE_PANE_FILTER` and its 15 entries are gone, as the table's own strictness forced (section 6),
  and `KNOWN_MISSING_ATTRIBUTES` goes from seventeen to two.
- **`test/smoke.mjs`:** CR-019's block holds the 16 attributes read and written in the source's spelling, an element with no
  `w:val`, `w:val` changed alone and written as given (D3), and `w:stylePaneSortMethod` still `CTShortHexNumber`.
- **The gate:** the marshalled hashes differ in that one part only, and `typecheck` compiles the regenerated declarations.
- **Docs:** `CLAUDE.md`'s seventeen and its list, `README.md`'s note on `val` (section 5), and the `SOURCE.md` sentence
  (section 2.3). CR-004 section 11 is a dated record and stays.

## 10. Consumers

- **core-ts:** it doesn't use the property (no hits in `src/` or `test/`). CR-025's one real file stops reporting, so
  its 15 dropped attributes go to 0 and a `part.skipped` test needs another source. That file's settings hash changes.
- **`@docx4j/generated-objects-ts` callers:** `stylePaneFormatFilter` becomes `CTStylePaneFilter`, so reading `.val` gives
  `string | undefined`. A value built with the old `TYPE_NAME` is written with `xsi:type="w:CT_ShortHexNumber"` (measured).
  Setting `val` alone on a loaded file with the attributes no longer changes what Word reads (section 5).
- **docx4j-python:** it can retire its patch at its next refresh. Its generated class's 15 fields go from `bool` to the enum,
  which takes a CHANGELOG line, and its normaliser, which picks on/off attributes by `bool`, stops covering them.

The regeneration went out in 0.8.0, a minor for CR-020, CR-021 and CR-023 beside it, with core-ts's go at `9a2a945`.
