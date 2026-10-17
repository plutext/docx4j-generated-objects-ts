# CR-009: Type what a kept `mc:AlternateContent` branch holds, as docx4j does

**Status:** Proposed 2026-10-17
**Requested by:** `plutext/docx4j-core-ts` CR-010 phase 2 (Jason asked for phase 2 in the core-ts
session, 2026-10-17). CR-010 keeps `mc:AlternateContent` and saves every branch, docx4j CR-021's
policy. Its phase 1 is done (core-ts `6b178af`); phase 3 puts the content API over kept branches.
**Depends on:** nothing new. The measurements below ran on 0.3.2 and `@docx4j/jsonix` 3.4.1.
**Counterpart:** docx4j's `@XmlRootElement` classes and unscoped `ObjectFactory` `@XmlElementDecl`s,
which its JAXB context types in a lax wildcard.

## 1. Summary

The object model puts what a kept `mc:Choice` or `mc:Fallback` holds in a lax wildcard (`any`,
`allowTypedObject`). The runtime types an element there only when the context declares it
globally: `getElementInfo(name, scope)` looks in the Choice's own scope, then in `##global`. The
WordprocessingML mapping declares `w:drawing`, `w:pict`, `w:tbl`, `w:t`, `w:hyperlink` and others
only within their parents (`scope: '.R'`, ...), so inside a kept branch they stay DOM. The content
API can't see them, and a mutator can't edit them. docx4j's context types the same elements,
because its classes declare them as root elements.

Proposed: the context declares globally the 61 WordprocessingML elements docx4j declares globally
and this mapping doesn't (section 4). Measured by registering them in a live context (no module
edited):
- Inside kept branches, `w:drawing` (13) and `w:pict` (9) come back typed.
- Today's default load, which resolves branches, changes in none of 8,235 parts.
- Every part still writes back whole, but for one VML attribute that docx4j loses too (section 3).

## 2. What docx4j declares that this mapping doesn't

Source: docx4j's generated objects, as its checkout built them on 2026-10-17 (`VERSION_17_3_4` at
`23cb5bf66`, `docx4j-generated-objects/target/generated-sources/xjc`). There are 191
`@XmlRootElement` classes and 513 `@XmlElementDecl`s with no `scope`, which make 640 global names.
Compared with the context's `scopedElementInfosMap['##global']`:

- **564** are global here too.
- **46 are declared here only within parents.** All are WordprocessingML, and each has one type
  across its scopes (list A, section 4).
- **30 have no element declaration here at all.** The mapping knows them only as their parents'
  properties.
  - 19 are WordprocessingML. Four are `w:sdtContent`, which docx4j maps to four classes
    (`CTSdtContentRun`, `CTSdtContentCell`, `CTSdtContentRow`, `SdtContentBlock`), ambiguous there
    too, so they're left out. The other 15 are list B.
  - 11 are in other vocabularies, which nothing reads through the content API:
    - PresentationML `p:sp` and `p:cmAuthor`;
    - SpreadsheetML `col`, `cols`, `row`, `sheetData`, `sheetFormatPr`, `sheetView` and
      `sheetViews`;
    - the custom properties' `property`;
    - the flat OPC `pkg:part` and `pkg:xmlData`;
    - the task pane's `taskpane`;
    - `{}xml`.

    They're left out here. A later CR can take them if a consumer needs pptx or xlsx branches
    typed.

## 3. Measured with the 61 registered (2026-10-17)

The corpus is the 459 files core-ts's release checks hash: its `test/fixtures/` and docx4j's
`docx4j-core-tests/src/test/resources/charts/`. Each global was registered by copying its scoped
declaration into `##global` with no scope, or, for list B, by building an `ElementInfo` over the
type named.

**With every branch kept** (core-ts `mcePreprocess: false`, then `unmarshalAll` and a save; 281
parts hold the element):

| inside kept branches | today | with the 61 |
|---|---|---|
| `w:drawing` | DOM 13 | typed 13 |
| `w:pict` | DOM 9 | typed 9 |
| `c14:style`, `x15ac:absPath` | typed | typed |
| `c:style` 227, `xdr:sp` 9, `xdr:graphicFrame` 7, `p:graphicFrame` 2, `p:pic` 2, `p:sp` 2, `controls` 2, `xdr:twoCellAnchor` 2 | DOM | DOM (none of them is in list A or B) |

**The round trip.** Each `mc:AlternateContent` was compared with its source (names, attributes,
children; prefix declarations aside):
- Today, all 281 parts come back whole. With the 61, 277 do.
- The other four (`loadAndSave.docx` and core-ts's three `cr043-*.docx`) lose **`w14:anchorId` on
  the `v:shapetype` inside a `w:pict` Fallback**, and nothing else. `org_docx4j_vml.CTShapetype`
  has no any-attribute.
- docx4j has the same gap, checked with docx4j 17.3.2: loading `loadAndSave.docx` and marshalling
  its main document part, the Fallback is kept and the attribute is gone. docx4j also writes
  `o:spt="202"` back as `"202.0"`, which this package doesn't.

Two rewrites are lexical, not losses, and are what a resolved Choice already gets today:
- `xsd:boolean` values written `1`/`0` come back `true`/`false`;
- `o:gfxdata`, base64, loses its line breaks.

The comparison above treats both as equal.

**With branches resolved** (today's default load), every part of all 459 files was hashed after
`unmarshalAll` and a save. None of the 8,235 parts changes when the 61 are registered.

## 4. Proposal

1. **Declare the 61 globally,** each with the type docx4j's root class has. Lists A and B give the
   mapping's type names, without the `org_docx4j_wml.` prefix.

   List A, declared here only within parents (copy the scoped declaration, without its scope):

   `w:alias` (SdtPr.Alias), `w:altChunk` (CTAltChunk), `w:annotationRef` (R.AnnotationRef),
   `w:bibliography` (SdtPr.Bibliography), `w:bookmarkEnd` (CTMarkupRange), `w:bookmarkStart`
   (CTBookmark), `w:citation` (SdtPr.Citation), `w:commentReference` (R.CommentReference),
   `w:continuationSeparator` (R.ContinuationSeparator), `w:cr` (R.Cr), `w:dataBinding` (CTDataBinding),
   `w:dayLong` (R.DayLong), `w:dayShort` (R.DayShort), `w:drawing` (Drawing), `w:endnoteRef`
   (R.EndnoteRef), `w:equation` (SdtPr.Equation), `w:fldChar` (FldChar), `w:fldSimple` (CTSimpleField),
   `w:footnoteRef` (R.FootnoteRef), `w:group` (SdtPr.Group), `w:hyperlink` (P.Hyperlink), `w:jc` (Jc),
   `w:lastRenderedPageBreak` (R.LastRenderedPageBreak), `w:lock` (CTLock), `w:monthLong` (R.MonthLong),
   `w:monthShort` (R.MonthShort), `w:noBreakHyphen` (R.NoBreakHyphen), `w:permEnd` (CTPerm),
   `w:permStart` (RangePermissionStart), `w:pgNum` (R.PgNum), `w:pict` (Pict), `w:picture`
   (SdtPr.Picture), `w:ptab` (R.Ptab), `w:rPr` (RPr), `w:richText` (SdtPr.RichText), `w:separator`
   (R.Separator), `w:smartTag` (CTSmartTagRun), `w:softHyphen` (R.SoftHyphen), `w:spacing`
   (CTSignedTwipsMeasure), `w:sym` (R.Sym), `w:t` (Text), `w:tab` (R.Tab), `w:tbl` (Tbl), `w:tc` (Tc),
   `w:yearLong` (R.YearLong), `w:yearShort` (R.YearShort)

   List B, declared here only as properties (a new declaration over the type named). For `pPr`,
   `sectPr`, `tblPr` and `tblGrid` this is docx4j's root class, not the `...Base` type their
   properties use within `w:pPrChange` and the like:

   `w:ind` (PPrBase.Ind), `w:numPr` (PPrBase.NumPr), `w:outlineLvl` (PPrBase.OutlineLvl), `w:pBdr`
   (PPrBase.PBdr), `w:tblPr` (TblPr), `w:abstractNum` (Numbering.AbstractNum), `w:num` (Numbering.Num),
   `w:lvlOverride` (Numbering.Num.LvlOverride), `w:numPicBullet` (Numbering.NumPicBullet), `w:sectPr`
   (SectPr), `w:pPr` (PPr), `w:tabs` (Tabs), `w:sdtPr` (SdtPr), `w:lvl` (Lvl), `w:tblGrid` (TblGrid)

   There are two places to do it:
   - **The facade** registers them when it builds the context, as the measurement did. It's
     hand-written code here, needs no regeneration, and keeps the list next to its source.
   - **The generated mapping** declares them, through a compiler CR (`jsonix-schema-compiler`)
     and a regeneration, since `modules/` is generated.

   Recommended: the facade, with the list committed as data (docx4j's names and the mapping's
   types) and a test that every entry resolves to a type in the context. A regeneration that
   renames a type then announces itself. The form that needs no runtime mutation is **one more
   mapping**, hand-written in `src/` (say `org_docx4j_wml_globals`): `dependencies:
   ['org_docx4j_wml']`, no `typeInfos`, and an `elementInfos` entry per name whose `typeInfo` is
   the qualified type name (`'org_docx4j_wml.Drawing'`), the way the generated modules reference
   each other's types. `getContext` passes it to `new Jsonix.Context` beside the generated
   modules, and `modulesFor` includes it whenever `org_docx4j_wml` is in the closure. It is not a
   generated module, so it does not go in `modules/` or the CR-005 registry's count; it is facade
   code, like the prefix table. Verified on a kept branch, section 7.
2. **Keep `w14:anchorId` on `v:shapetype`.** Either give `CT_Shapetype` an any-attribute (and its
   VML siblings if the fidelity test, CR-004, finds the same there), or declare the attribute.
   Either way it's a generated-type change, through the compiler or `bindings.xjb`. The
   loss also exists in docx4j's binding; the core-ts session tells Jason, who decides whether
   docx4j hears of it.

## 5. Tests

- Each of the 61 resolves: `getElementInfo` with no scope returns its type.
- A `w:drawing` and a `w:pict` inside a kept `mc:Choice` and `mc:Fallback` (a fragment under `w:r`)
  unmarshal typed and marshal back whole.
- CR-004's fidelity test runs unchanged, since it resolves.
- On the core-ts side, CR-010 phase 2's test is section 3's survey:
  - every `w:drawing` and `w:pict` in a kept branch typed;
  - every kept element back whole, booleans and base64 compared by value;
  - default-mode output unchanged.

  core-ts records today's DOM state as a test asserting it, which fails when this release
  arrives.

## 6. Risks

- **Every lax wildcard in the model sees these declarations,** not only `mc:Choice` and
  `mc:Fallback`. Counted over the mappings' `propertyInfos` of `type: 'anyElement'`, the context
  has 31, of which 28 allow a typed object (the runtime's built-in `xs:anyType` adds a 29th typed
  one in a live context, declared by no module); among them `a:graphicData` (`GraphicData.any`), the content of
  `w:pict` and `w:object` (`CTPictureBase.anyAndAny`), and the extension lists' `ext`
  (`CTExtension.any` in chart, chartex, w16 and pml). Across the corpus nothing changed
  with branches resolved, so the effect today is confined to kept branches. A file whose extension
  content carries one of these elements would now have it typed, which is lossless only where its
  type binds all of it. CR-004's fidelity test is the guard.
- **Typing exposes binding gaps the DOM form hid:** the `w14:anchorId` above is one. core-ts
  re-runs section 3's survey on each objects release before CR-010 phase 4 makes keeping the
  default.

## 7. Verified here (2026-10-17)

The objects-ts session built the context from the 104 generated modules plus one extra mapping
of the recommended form (section 4), declaring `w:drawing`, `w:pict` and `w:pPr` by qualified type
name, and unmarshalled a `w:r` holding an `mc:AlternateContent` whose Choice carries a `w:drawing`
and whose Fallback a `w:pict` over a `v:shapetype` with `w14:anchorId` and `o:spt="202"`:

| | today | with the extra mapping |
|---|---|---|
| Choice content | DOM | `org_docx4j_wml.Drawing` |
| Fallback content | DOM | `org_docx4j_wml.Pict` |
| `w14:anchorId` on the marshalled `v:shapetype` | kept | **lost** (proposal 2) |
| `o:spt` written back | `202` | `202` |
| `unmarshalString('<w:drawing/>')` at the root | throws | `org_docx4j_wml.Drawing` |

The runtime accepts a mapping with no `typeInfos` whose `elementInfos` name another module's
types, so no context mutation is needed. The last row is a side effect worth stating: each of
the 61 becomes a valid document root for `unmarshalString` and `marshalString`, which is what
docx4j's `@XmlRootElement` gives Java callers. CR-002's `wml` wrapper stays, since it exists for
every scoped element, not only these, and its wrapping is harmless for a now-global one.
`getElementInfo`'s fallback to `##global` is at `jsonix-factory.js:6086` in `@docx4j/jsonix` 3.4.1.

