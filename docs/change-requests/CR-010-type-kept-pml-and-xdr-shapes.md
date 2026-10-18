# CR-010: Type the shapes a kept `mc:AlternateContent` holds in slides and spreadsheet drawings

**Status:** Proposed 2026-10-18
**Requested by:** nobody yet. core-ts CR-010 section 14.7 records it as not done: `p:sp` and the
`xdr:` shapes in a kept branch stay DOM, docx4j types `p:sp` there, and no consumer has asked.
Drafted at Jason's request, 2026-10-18.
**Depends on:** CR-009, whose form (a hand-written mapping of global declarations beside the
generated ones) this repeats. Proposal 3 also depends on docx4j binding `a14:legacySpreadsheetColorIndex`
on `a:srgbClr` (section 3.2); no docx4j CR asks for that yet.
**Counterpart:** docx4j's `@XmlRootElement` on `org.pptx4j.pml.Shape` (`p:sp`) and
`CTCommentAuthor` (`p:cmAuthor`). docx4j declares no other PresentationML shape and no
SpreadsheetML drawing shape or anchor as a root, so its JAXB context leaves those DOM in a kept
branch, as this one does today.

## 1. Summary

CR-009 made the context declare globally the WordprocessingML elements docx4j declares as roots.
The lax wildcard of a kept `mc:Choice` or `mc:Fallback` then types `w:drawing` and `w:pict`, as
docx4j's does. It left PresentationML and SpreadsheetML out (its section 2): nothing read them
through the content API.

A slide's shape tree and a spreadsheet drawing hold `mc:AlternateContent` too:
- PowerPoint puts a chartex `p:graphicFrame` in a Choice with a `p:pic` Fallback.
- Excel does the same with `xdr:graphicFrame` and `xdr:sp` for slicers, timelines and chartex
  charts.
- Excel puts a whole form-control `xdr:twoCellAnchor` in an `a14` Choice.

Inside a kept branch all of these are DOM today. Since core-ts CR-010 phase 4, keeping is core-ts's
default load. So `placeholdersOf` does not see a placeholder `p:sp` inside a Choice, and a mutator
cannot edit those shapes.

Proposed, in three parts, measured in section 3:
1. **`p:sp` and `p:cmAuthor`:** the two PresentationML roots docx4j declares and this mapping
   doesn't. They come back whole, and this is parity with docx4j. Ready.
2. **The other shapes:** `p:pic`, `p:graphicFrame`, `p:grpSp`, `p:cxnSp`, `p:contentPart`, and
   the `xdr:` shapes `sp`, `pic`, `graphicFrame`, `grpSp` and `cxnSp`. Typed, every one comes back
   whole, but docx4j declares none of them, so this goes beyond docx4j.
3. **The three `xdr:` anchors** (`twoCellAnchor`, `oneCellAnchor`, `absoluteAnchor`): beyond docx4j
   too. Typing a kept anchor loses `a14:legacySpreadsheetColorIndex` in both form-control
   workbooks of the corpus. Kept as DOM, the attribute survives today.

## 2. What docx4j declares, and what this context declares

Source: docx4j's generated objects as `VERSION_17_3_4` built them on 2026-10-18
(`docx4j-generated-objects/target/generated-sources/xjc`, checkout at `f59828e62`).
`org.pptx4j.pml` has 14 `@XmlRootElement` classes and 4 unscoped `@XmlElementDecl`s;
`org.docx4j.dml.spreadsheetdrawing` has 1 and 3.

| | docx4j declares globally | here globally | missing here |
|---|---|---|---|
| PresentationML | `cmAuthor`, `cmAuthorLst`, `cmLst`, `handoutMaster`, `notes`, `notesMaster`, `oleObj`, `presentation`, `presentationPr`, `sld`, `sldLayout`, `sldMaster`, `sldSyncPr`, `sp`, `tagLst`, `viewPr` | the same, except `sp` and `cmAuthor` | `sp` (`Shape`), `cmAuthor` (`CTCommentAuthor`) |
| SpreadsheetML drawing | `wsDr`, `from`, `to` | the same | none |

docx4j makes `p:sp` a root by schema annotation, not by a global element: `pml-2ed.xsd` line 3721
carries `<annox:annotate target="class">@jakarta.xml.bind.annotation.XmlRootElement(name="sp")`,
which xjc applies to `CT_Shape`'s class. The Jsonix compiler ignores `annox`, which is why this
mapping lacks it.

The mapping declares none of the shapes within a parent scope either (`scopedElementInfosMap`
holds no `p:` or `xdr:` shape name). They exist only as members of their parents' element
properties:
- `org_pptx4j_pml.GroupShape.spOrGrpSpOrGraphicFrame` (`sp` → `Shape`, `grpSp` → `GroupShape`,
  `graphicFrame` → `CTGraphicalObjectFrame`, `cxnSp` → `CxnSp`, `pic` → `Pic`, `contentPart` →
  `CTRel`);
- `org_docx4j_dml_spreadsheetdrawing.CTGroupShape.spOrGrpSpOrGraphicFrame` (`sp` → `CTShape`,
  `grpSp` → `CTGroupShape`, `graphicFrame` → `CTGraphicalObjectFrame`, `cxnSp` → `CTConnector`,
  `pic` → `CTPicture`);
- `CTDrawing.egAnchor` (`twoCellAnchor` → `CTTwoCellAnchor`, `oneCellAnchor` → `CTOneCellAnchor`,
  `absoluteAnchor` → `CTAbsoluteAnchor`).

Each name has one type across its properties. Those types are what proposals 2 and 3 would
declare.

## 3. Measured (2026-10-18)

The corpus is CR-004's: 341 parts of seventeen documents under `test/fixtures/fidelity/`. Each part
was unmarshalled and marshalled through the facade, with branches kept (the facade's own
behaviour), twice:
- over the default context;
- with one extra mapping of CR-009's form declaring the names.

Branch children were counted typed or DOM. Each output was compared with its source by
`test/lib/canonical.mjs`, and losses were collected by `onUnexpectedElement` and
`onUnexpectedAttribute`.

### 3.1 Where the shapes sit in a branch

Nine parts hold one of these names as a child of a Choice or Fallback:

| element in a branch | parts | count | branch | typed as |
|---|---|---|---|---|
| `p:sp` | `loadAndSave.pptx` slide 2 | 2 | `a14` Choice and its Fallback | `org_pptx4j_pml.Shape` |
| `p:graphicFrame` | `waterfall.pptx` slide 1 | 1 | `cx1` Choice (a chartex chart) | `org_pptx4j_pml.CTGraphicalObjectFrame` |
| `p:pic` | `waterfall.pptx` slide 1 | 1 | its Fallback | `org_pptx4j_pml.Pic` |
| `xdr:graphicFrame` | `cr022-slicers-timelines.xlsx` (two drawings), `manual-bincount-6.xlsx`, `manual-binwidth-10.xlsx` | 5 | `a14`/`tsle` Choices (slicer, timeline), `cx1` Choices (chartex) | `org_docx4j_dml_spreadsheetdrawing.CTGraphicalObjectFrame` |
| `xdr:sp` | the same four drawings, and `loadAndSave.xlsx` | 7 | their Fallbacks; `loadAndSave.xlsx`'s `a14` Choice and Fallback | `org_docx4j_dml_spreadsheetdrawing.CTShape` |
| `xdr:twoCellAnchor` | `cr022-checkbox.xlsx`, `cr022-checkbox-linked.xlsx` | 2 | `a14` Choice under `xdr:wsDr` (a form-control checkbox) | `org_docx4j_dml_spreadsheetdrawing.CTTwoCellAnchor` |

core-ts's larger corpus counted much the same in CR-009 section 3: `xdr:sp` 9,
`xdr:graphicFrame` 7, `p:graphicFrame` 2, `p:pic` 2, `p:sp` 2, `xdr:twoCellAnchor` 2.

### 3.2 The round trip

- **Proposal 1 alone** (`p:sp`, `p:cmAuthor`): one part changes, `loadAndSave.pptx` slide 2. Both
  `p:sp` come back typed, and the part comes back whole (no canonical difference, nothing
  reported dropped).
- **All three proposals:** nine parts change, the nine of section 3.1.
  - Every listed element comes back typed.
  - Seven of the nine come back whole: both slides, the slicer and timeline drawings, the two
    chartex drawings and `loadAndSave.xlsx`'s drawing. The slicer and timeline content inside
    `a:graphicData` stays DOM, since `sle` and `tsle` are unbound, and comes back whole as DOM
    does.
  - The two checkbox drawings each lose two attributes and the `mc:Ignorable` beside each:
    `<a:srgbClr val="FFFFFF" mc:Ignorable="a14" a14:legacySpreadsheetColorIndex="65"/>` (and
    `000000`/`64`) inside `a14:hiddenFill` and `a14:hiddenLine` come back as
    `<a:srgbClr val="FFFFFF"/>`. `CT_SRgbColor` has no any-attribute.
- **The other 332 parts:** byte-identical output and the same branch typing under each set. No other
  lax wildcard in the corpus meets one of these names.

The checkbox loss is not new to the model. It is one of `test/fidelity.mjs`'s
`KNOWN_MISSING_ATTRIBUTES` (with the `mc:Ignorable` entry), recorded for the **resolved** check:
once the `a14` Choice is taken, the content is typed and the attribute goes. Unresolved it
survives, because the whole branch is DOM. core-ts CR-001 section 21.1 now lists it as kept, for
that reason, since core-ts CR-010 phase 4 keeps the branch by default. Proposal 3 would end that:
- a kept checkbox anchor would lose the attribute too;
- core-ts's default load of a form-control workbook would regress to the resolved form's loss.

The `xdr:sp` inside that anchor is typed only once the anchor is. Under proposal 2 alone the
checkbox drawings stay DOM below the anchor, and come back whole.

## 4. Proposal

Each proposal is a hand-written mapping beside the generated ones, as CR-009's
`org_docx4j_wml_globals` is: `typeInfos: []`, one `elementInfos` entry per name with the qualified
type name. `getContext` adds it, and `modulesFor` adds it when its module is in the closure. Its
name table is exported as data, as `WML_GLOBAL_ELEMENTS` is.

1. **Declare `p:sp` (`Shape`) and `p:cmAuthor` (`CTCommentAuthor`) globally:**
   `PML_GLOBAL_ELEMENTS` and `org_pptx4j_pml_globals`, in a new `src/pml-globals.mts`.
   - **What it gives:** a kept slide branch types its `p:sp`, so core-ts's `placeholdersOf` could
     read one through `selectedBranchOf` (core-ts CR-010 section 14.7).
   - **A side effect:** `unmarshalString('<p:sp ...>')` works as a root. That is what docx4j's root
     is for: `docx4j-samples-pptx4j`'s `Table.java` (line 75) unmarshals its `SAMPLE_SHAPE`, a
     `p:sp`, with `XmlUtils.unmarshalString(SAMPLE_SHAPE, Context.jcPML)`. Today it throws here
     (`could not be unmarshalled`).
   - **Status:** parity with docx4j, measured whole. Ready to implement.
2. **The other shapes, beyond docx4j:** `p:pic`, `p:graphicFrame`, `p:grpSp`, `p:cxnSp`,
   `p:contentPart`, and `xdr:sp`, `xdr:pic`, `xdr:graphicFrame`, `xdr:grpSp`, `xdr:cxnSp`. All
   measured whole. There are two routes:
   - **a. The facade declares them here**, as a recorded departure from docx4j. The tables would
     then say, entry by entry, which names docx4j declares and which it does not. CR-009's are
     all docx4j's today.
   - **b. docx4j declares them first**, and this package follows. A docx4j CR would add `annox`
     root annotations like `p:sp`'s. They regenerate to an empty diff here, since the compiler
     ignores `annox`, and this package adds the names to its tables.

     docx4j would type the same kept branches. It keeps every branch under its CR-021, so the
     gain is the same there.

   Recommended: **b**. Every global declaration here is docx4j's today, and CR-009's test and
   rationale rest on that. One name list across both packages keeps this context typing what
   docx4j's types, which is CR-009's promise. Route b costs a docx4j release cycle. Nothing reads these elements yet, so the wait costs nothing now.
3. **The anchors, beyond docx4j and gated on the attribute:** `xdr:twoCellAnchor`,
   `xdr:oneCellAnchor` and `xdr:absoluteAnchor` (the `egAnchor` members). Not before the model
   binds `a14:legacySpreadsheetColorIndex` on `a:srgbClr` (section 3.2). That binding is a docx4j
   schema change, like docx4j CR-048's `w14:anchorId`. It fits the same docx4j CR as 2b.

   The `mc:Ignorable` on the `a:srgbClr` would still go. The model binds `mc:Ignorable` on part
   roots only. CR-008 then declares `a14` ignorable on the drawing's root, so the output stays
   valid, but the fidelity comparison reports it. Whether that is acceptable belongs in that docx4j
   CR.

Recommended order:
- proposal 1 when a consumer wants it, or with the next minor;
- proposals 2 and 3 through one docx4j CR. This package then adds docx4j's new roots to its
  tables. The annotations regenerate to an empty diff, and proposal 3's attribute binding
  regenerates `modules/`.

None of this is a patch. It changes what a kept branch holds and what is a valid root, so it is a
minor, like CR-009's 0.4.0.

## 5. Tests

- **Smoke**, as CR-009's block:
  - each entry resolves globally (`getElementInfo` with no scope) to its type;
  - the generated module declares none of them itself, so a regeneration that does retires the
    entry;
  - `modulesFor('org_pptx4j_pml')` adds the mapping, and a closure without the module does not.
- **Smoke:** a `p:spTree` holding an `a14` Choice with a `p:sp` and a Fallback with a `p:sp`
  unmarshals both typed and marshals whole. `unmarshalString` takes a bare `p:sp`.
- **Fidelity:** unchanged for proposals 1 and 2, which section 3.2 measured. Proposal 3 needs one
  more check, because `KNOWN_MISSING_ATTRIBUTES` is keyed by attribute, not by part or check: a loss
  newly exposed in an unresolved part would be absorbed by the existing entry without a word. The
  unresolved checks of the two checkbox drawings must stay whole, or the table must record which
  checks lose the attribute.
- **core-ts:** re-run CR-010's survey on the release, and decide whether `placeholdersOf` reads a
  kept `p:spTree` branch.

## 6. Risks

- **Every lax wildcard sees the declarations,** as CR-009 section 6 says of its 61: `a:graphicData`,
  the extension lists' `ext`, `w:pict`'s content. Across the corpus nothing outside section 3.1
  changed. A file whose extension content carries one of these names would have it typed, which is
  lossless only where the type binds everything.
- **Typing exposes binding gaps the DOM form hid.** The checkbox attribute is the one in this
  corpus, and the reason proposal 3 waits. The corpus has seventeen documents. The other shapes
  carry other extension attributes in files it doesn't have, which core-ts's survey over its
  larger corpus would show first.
- **A shape becomes a valid root,** so `marshalString` of a bare `p:sp` writes it without the
  `xsi:type="p:CT_Shape"` it gets today, as CR-009 found for `w:tbl` and `w:tc`. Checked on a
  title placeholder: XML saved in the old form still unmarshals as `org_pptx4j_pml.Shape` and
  writes back in the new form.

## 7. Consumers

- **core-ts:**
  - `placeholdersOf` (`src/parts/pml/placeholders.mts`) iterates
    `cSld.spTree.spOrGrpSpOrGraphicFrame` and matches `TYPE_NAME === 'org_pptx4j_pml.Shape'`.
    Typed `p:sp` in a kept branch is what lets it look inside one.
  - CR-010 section 14.7 names this CR's subject, and the survey is core-ts's.
  - The default load keeps branches since phase 4, so typing applies to core-ts's ordinary load,
    not only to a caller who keeps branches.
- **The editor:** nothing reads slides or spreadsheet drawings there.
- **Office JS add-ins using this package alone:** a kept branch in a slide or drawing types where
  it was DOM. Code that tested for `nodeType` there would see objects instead.
