# CR-004: A round-trip fidelity test over parts Office wrote

**Status:** Phase A implemented 2026-09-23 (section 8), phase B 2026-09-25 (section 9)
**Depends on:** nothing in this package; the fixtures come from `plutext/docx4j`
**Requested by:** the docx4j session (2026-09-20), after CR-021 to CR-023 closed a run of losses
that none of this package's tests could have found
**Counterpart:** docx4j's `ExcelExtensionsTest`, `WordExtensionAttributesTest` and the byte
comparison in its CR gates

## 1. Summary

Every loss docx4j's CR-021, CR-022 and CR-023 closed was of one kind: unmarshal a part Office wrote,
marshal it again, and something is missing. None was visible in the declarations, and none of the
three tests here would have caught any of them:

- `test/smoke.mjs` round-trips one small `document.xml` this repository wrote itself;
- `test/readme-examples.ts` and `test/nodenext/consumer.mts` are compile-only.

This CR adds the missing kind: take the XML parts of documents Office 365 saved, unmarshal and
marshal each through the facade, and compare canonically.

**A trial run on 2026-09-20 found a serious bug in twenty minutes**, which is the argument for the
CR. Over 17 parts from docx4j's `cr022-checkbox.xlsx` and `tracked-changes-equations.docx`, 12 were
identical and four differed:

| Part | Finding |
|---|---|
| `xl/drawings/drawing1.xml` | **empties**: 59 elements in, one out, silently. Excel's `xdr:wsDr` holds an `mc:AlternateContent`, which `CT_Drawing` does not admit, so it is dropped. Reported to docx4j, whose schema it is. |
| `xl/styles.xml` | `mc:Ignorable` lost `x16r2`, a prefix the table lacked (fixed, `7a74fb0`) |
| `xl/workbook.xml` | the `xr2:uid` loss docx4j has logged as its own xlsx4j CR, seen from a second direction |
| `word/settings.xml` | same content, different element order: to be understood, then either fixed or justified |

`[Content_Types].xml` throws, since the model does not carry it; the test excludes package-level
parts, which are `@docx4j/core-ts`'s ground.

## 2. What the test does

For each fixture part:

1. `unmarshalString`, then `marshalString`, both through the facade (not `createMarshaller`, so the
   namespace rules of CR-001 are exercised too);
2. canonicalise both sides and compare;
3. on a difference, report the part, the counts, and the first differing line.

A part that throws is a failure with its message, since Office wrote it and the model claims to
cover it.

### 2.1 Canonicalisation

Raw bytes differ legitimately, so the comparison normalises, and each rule is a decision this CR
records rather than a convenience:

- **QNames, not prefixes**: `{ns}local`, since a prefix table choice is not a difference (CR-001
  tests the prefixes themselves).
- **Attributes sorted** by QName: the declarations promise sorted output, Office's order is its own.
- **Namespace declarations ignored** on comparison: CR-001 section 3.1's tests cover which
  declarations appear, and `mc:Ignorable` is compared as an attribute value.
- **`xs:boolean` attributes normalised**: `1` and `true` are the same value. The CR first said "only
  for attributes the model types as boolean, not for every attribute that looks like one"; the
  implementation normalises every attribute's lexical form instead, which is **safe in this
  comparison and not in general**, and the reason is worth keeping: the two sides are a part Office
  wrote and that same part marshalled back, so an attribute the model does not type as boolean is
  copied through verbatim and cannot present two spellings. Only the model's own typing can produce
  a `1` against a `true`. The list came first and silently missed `noGrp`, `noChangeAspect`,
  `anchorCtr` and `knownFonts` - a canonicaliser with a hand-kept list of what to normalise is a
  canonicaliser that hides differences by omission.
- **Whitespace-only text dropped**, and text otherwise compared exactly. `xml:space="preserve"`
  content is compared exactly, which the canonicaliser must respect.

A difference in element order is a difference: docx4j's `propertyOrder` (compiler CR-007) is there
to keep schema order, so `word/settings.xml` above is a real question, not noise.

## 3. Fixtures

Extracted part XML as text files under `test/fixtures/fidelity/<document>/<part path>.xml`, not the
`.docx` and `.xlsx` archives:

- no zip dependency (this package has one devDependency, TypeScript, and no runtime dependency but
  the Jsonix runtime);
- a failing test shows a readable diff;
- `package.json`'s `files` ships `dist` and `modules` only, so fixtures cost consumers nothing.

Sizes measured 2026-09-20: the six `cr022-*.xlsx` hold 103 XML parts, 222 kB uncompressed
(`cr022-data-model.xlsx` is 245 kB of archive, almost all the binary `xl/model/item.data`, which is
not XML and stays out). The Word fixtures add `tracked-changes-equations.docx` (9 parts),
`numPicBullet-word2019-pict.docx` and `loadAndSave.docx`.

Each fixture directory carries a `SOURCE.md`: the docx4j path it came from, the commit, and what it
demonstrates. Fixtures are never edited by hand; a part that needs trimming is regenerated from the
source document and that is recorded.

## 4. Phases

- **A** (implemented 2026-09-23): the canonicaliser, the runner, and the parts that between them
  cover every loss CR-021 to CR-023 closed: `xl/workbook.xml`, `xl/worksheets/sheet1.xml` (x14ac
  attributes, a control), `xl/styles.xml`, `xl/drawings/drawing1.xml`, a slicer cache, a timeline
  cache, `xl/ctrlProps/ctrlProp1.xml`, `word/document.xml` (tracked changes with `dateUtc`),
  `word/numbering.xml` (`durableId`, `restartNumberingAfterBreak`), `word/settings.xml`,
  `word/people.xml`, `word/commentsExtensible.xml`. Four more were added as the CR was implemented,
  each from a loss found after the CR was written: `xl/drawings/vmlDrawing1.vml` (docx4j CR-026),
  two `chart1.xml` carrying `c16r3:dispNaAsBlank` (docx4j `c363a969f`), and `ppt/slides/slide2.xml`
  with `xl/drawings/drawing1.xml` from `loadAndSave.xlsx`, both also checked with their `a14`
  `mc:Choice` taken (0.1.5's regression). 17 parts, 19 checks.
- **B** (implemented 2026-09-25): every XML part of the fixtures, triaged into bugs here, schema
  gaps for docx4j, or canonicalisation rules with their justification. 237 parts from twelve
  documents; see section 9.

## 5. Tests and layout

`test/fidelity.mjs`, run from `npm test` after the build, beside `test/smoke.mjs`; the canonicaliser
in `test/lib/canonical.mjs` so it can be unit-tested itself (a difference the canonicaliser hides is
worse than one it reports). `CLAUDE.md`'s test list gains it. No `src/` change: this is a test.

## 6. What it does not cover

- Parts and packaging: content types, relationships, the zip. `@docx4j/core-ts` owns those, and its
  own round-trip tests cover them.
- Binary parts, including the data model's `item.data`.
- Whether Office opens the result. Only Word and Excel can answer that, which is what docx4j's
  manual checks are for; this catches the losses before it gets that far.

## 7. Open questions

1. How large a corpus before the runtime cost matters? Phase A is a dozen parts; phase B is ~110 and
   builds the context once, so the cost is parsing, not context construction.
   **Settled by measurement (2026-09-23)**: 19 checks over 17 parts add about a second to `npm test`,
   almost all of it the one context build the smoke already pays for. Phase B can proceed on size.
2. Should a known, accepted difference be recorded per part (an expectations file), or should the
   test fail until the difference is fixed or the canonicaliser justified?
   **Settled 2026-09-23, and neither purely**: two of the three differences phase A found are owned
   by docx4j and deferred there, so failing until fixed would mean a red suite indefinitely, and a
   red suite teaches people to ignore it. The runner keeps a `KNOWN` table instead, and it is strict
   in both directions: an entry names the owner and the difference, is matched on the difference
   itself (a part that starts differing some other way fails), and **fails when the part becomes
   identical**, so a fix elsewhere forces the entry out rather than passing silently. An
   expectations file that absorbs whatever it finds would defeat the test.
3. Do the fixtures belong here, or in a repository both packages can share? core-ts will want the
   same documents for its package-level tests.
   **Deferred, deliberately**: the four documents phase A uses are already identical in both
   repositories (checked byte for byte), so the cost of the duplication today is 212 kB and the
   `SOURCE.md` note that says where each came from. A shared repository is worth proposing when the
   corpus grows in phase B, not before.

## 8. What phase A found (2026-09-23)

17 parts from six documents, 19 checks (two parts are checked a second time with their `mc:Choice`
taken); 18 parts and 21 checks after the addition below. Sixteen round-trip identically. Three
differences, each recorded in `KNOWN`, and a fourth added the next day:

| Part | Difference | Owner |
|---|---|---|
| `cr022-checkbox.xlsx` `xl/workbook.xml` | `xr2:uid` on `workbookView` is dropped: not in the model | docx4j, logged as an xlsx4j CR |
| `cr022-slicers-timelines.xlsx` `xl/slicerCaches/slicerCache1.xml` | `mc:Ignorable` loses `x` | **this package**, see below |
| `tracked-changes-equations.docx` `word/settings.xml` | `w14:docId` and `w15:chartTrackingRefBased` swap | docx4j, deferred past 17.1.1 |
| `cr022-slicers-timelines.xlsx` `xl/drawings/drawing1.xml` (a14 taken, added 2026-09-24) | **throws**: an `sle:slicer` graphic cannot stay DOM | docx4j, see below |

The second is this package's own: Excel binds `x` to the SpreadsheetML main
namespace on its slicer, slicer cache and timeline parts and names it in `mc:Ignorable="x xr10"`,
but this package writes that namespace as the default (CR-001), so the prefix table cannot produce
`x` and the facade drops the token from `mc:Ignorable` rather than declaring the prefix. docx4j met
the same problem and answered it in its CR-024 by pre-declaring the prefix beside the default. The
fix here is a facade change and so wants its own CR; until it lands, a re-marshalled slicer cache
tells a reader to ignore one prefix fewer than Excel did.

**An eighteenth part, added 2026-09-24.** `cr022-slicers-timelines.xlsx` `xl/drawings/drawing1.xml`,
after `@docx4j/core-ts` met it implementing its own CR-004: an `a:graphicData` framing an
`sle:slicer`, inside the `a14` `mc:Choice` that every consumer takes. As Office wrote it the part
round-trips; with the branch taken it **throws**, because `CT_GraphicalObjectData`'s wildcard is
`processContents="strict"` in `dml-graphicalObject.xsd`, so a graphic no module binds cannot stay
DOM. Confirmed here and relayed to docx4j, whose schema it is; the remedy is `lax`, as its CR-021
did for the mce wildcards, and XJC gives `@XmlAnyElement(lax = true)` either way, which is why
docx4j never sees it. Recording it taught `KNOWN` to carry a throw: a part the model cannot read at
all is the same finding as one it reads lossily, and wants the same owner and the same insistence -
the entry is matched on the message and fails if the part throws differently or stops throwing.

**Two of the four closed on 2026-09-25**, by docx4j `8e8f6ea83` (both strict wildcards to lax,
CR-024 section 10) and `dae2dfc8b` (`CT_Settings` in Word's order), regenerated here: the slicer
graphic and `word/settings.xml` both round-trip, and the `KNOWN` entries failed until they were
removed, which is the behaviour the table exists for. CR-006 closed a third. One remains, the
`xr2:uid` docx4j has logged for xlsx4j.

**The device the `KNOWN` table is an instance of** - recording a dependency's defect as a test that
asserts the broken behaviour, so that the fix announces itself by failing - is written up once, in
`plutext/docx4j-core-ts` CR-001 section 19, "Recording a defect so that its fix cannot pass
unnoticed" (2026-09-25). core-ts arrived at the same thing independently in its
`test/ignorable.test.mjs` `UNREADABLE` table, and both have now caught a fix rather than a
regression. Its two boundaries apply here: the table is for defects in a **dependency**, never in
this package - a defect of our own is fixed or the test fails - and an entry without its reason and
its upstream reference decays into the exclusion it was meant to replace, which is why every entry
here carries an owner and a commit.

**The test was checked against the releases it was built for.** Run with 0.1.5's `modules/`, it
reports all five losses that release carried: the `a14` equation throwing in both the pptx slide and
the xlsx drawing (only in the resolved check - as Office wrote them, both round-trip, which is why
the resolved variant exists), `c16r3:dispNaAsBlank` losing its `val` in both charts, and the
vmlDrawing part throwing. Those cost two releases and a regression between them.


## 9. Phase B (2026-09-25)

Every XML part of twelve documents: the six `cr022-*.xlsx`, `tracked-changes-equations.docx`,
`numPicBullet-word2019-pict.docx`, and `loadAndSave.docx`/`.pptx`/`.xlsx`. **237 parts, 745 kB, and
the whole run takes under a second**, which settles open question 1 for good. 184 were identical at
the first run; 53 differed, in four groups, and the triage below is what phase B is for.

### 9.1 Two canonicalisation rules, both narrow (23 parts)

`docProps/core.xml` and `docProps/app.xml` differed in **every** document, by sibling order alone -
verified as a pure permutation in all 23, so nothing is lost. Both roots are **`xsd:all`** in
docx4j's schemas (`opc-coreProperties.xsd`, and `Properties` in
`shared-documentPropertiesExtended.xsd`), where no order is privileged. Jsonix cannot represent an
unordered model, so the compiler flattens the members into an ordered property list; and no order
would match anyway, because Word, Excel and PowerPoint each write `app.xml` differently (Word puts
`Application` after `Characters`, Excel first, PowerPoint third). There is nothing to fix in a
schema or in the facade, so this is section 2.1's premise - "the declarations promise schema order"
- being true of `xsd:sequence` and false of `xsd:all`. The canonicaliser now sorts the **subtrees**
of exactly those two elements, which keeps a value swapped between two differently named siblings
visible; `docProps/custom.xml` is an `xsd:sequence` and must never be added.

The second rule: `xsd:double` has many lexical forms per value (`1E-4` against `0.0001`,
`46285.360326851849` against `46285.36032685185`), so two attribute values that both parse as
finite numbers and compare `===` are the same value. As with the boolean rule, this can only hide a
difference the model's own typing created, since an untyped attribute is copied through verbatim.

### 9.2 Three losses, all docx4j's, all missing declarations (22 parts)

| Attribute | Dropped from | Instances |
|---|---|---|
| `xr:uid` (2014/revision) | `autoFilter`, `hyperlink`, `table`, `pivotCacheDefinition`, `pivotTableDefinition`, `comment` | 17 |
| `xr2:uid` (2015/revision2) | `workbookView` | 8 |
| `xr3:uid` (2016/revision3) | `tableColumn` | 13 |
| `xr16:uid` (2017/revision16) | `connection` | 5 |
| `Version` | `b:Sources` | 1 |

69 `uid` attributes in the corpus, **43 lost**. The mechanism is sound and only declarations are
missing: the 26 that survive are exactly the elements docx4j has already declared them on -
`CT_Worksheet`, and the x14/x15 slicer and timeline roots. `loadAndSave.xlsx`'s `sheet1.xml` makes
the point on its own, with two `xr:uid` in and one out: the worksheet's survives, the hyperlink's
does not. None of the losses is inside an `extLst` or an `mc:AlternateContent`. Relayed to docx4j
with the draft schema changes.

### 9.3 Eight parts nobody models, recorded as such

`docMetadata/LabelInfo.xml` (three documents), PowerPoint's `authors.xml` and modern comments,
Excel's `persons` and threaded comments, and a Power Query `DataMashup` blob. docx4j's
`ContentTypes.java` says of each "not bound, a `DefaultXmlPart`", and `unmarshalPackage` here
already keeps such a part as DOM, so the throw was the harness asking a question the model never
claimed to answer. They are now listed in `UNMODELLED` and checked the opposite
way round: **the root must still be unknown**, so if docx4j ever binds one the entry fails and has
to go. Each cites the docx4j `ContentTypes` **constant name** rather than its comment, at that
session's request: the names are the interface it keeps stable, the comments get reworded.

docx4j confirmed both losses of section 9.2 independently, on its own round trip with every part
forced to unmarshal - the same nine `xr*:uid` in `loadAndSave.xlsx` with one surviving, and
`b:Sources/@Version` - and is taking them as the "one small CR" its CR-022 section 20 left open,
with the two attribute-only schema files (2016/revision3, 2017/revision16) the drafts call for.

The `DataMashup` part also found a bug in the harness rather than the model: it is UTF-16 LE with a
BOM, and fixtures were read as UTF-8, so it arrived as mojibake. Parts are now read by their own
encoding.

### 9.4 What phase B changed about the method

Recording a loss **per part** stops working at this size: one missing declaration costs the same
attribute on 21 parts, and 21 entries would rot separately. Losses are now recorded per
**attribute**, in `KNOWN_MISSING_ATTRIBUTES`, and an entry no part exercises fails at the end of the
run - the same insistence as `KNOWN`, applied to a class.

That change forced another, and it is the more important of the two: the comparison used to report
only the **first** differing line of a part, which was enough when a part had one finding. With
findings recorded by class, the first difference can mask a later one - `xr3:uid` on `tableColumn`
was invisible behind the `table` element's own `xr:uid`, and only the staleness check revealed it.
Every differing line is now classified, and a line is explained only if the recorded attributes
account for **all** of it. That classifier is the code most able to hide a real loss, so it has its
own self-checks: a line that drops a recorded attribute *and* changes another value is not
explained.
