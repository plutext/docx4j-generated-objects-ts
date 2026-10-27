# CR-009: Type what a kept `mc:AlternateContent` branch holds, as docx4j does

**Status:** Implemented 2026-10-18 (proposal 2 by regeneration from docx4j CR-048, `97c43b9`; proposal 1 in the facade, section 9); released the same day in 0.4.0 (tag `4a4bf9b`); list C, nine elements beyond docx4j's roots, added 2026-10-28 by CR-017 M7, a recorded divergence (section 10)
**Requested by:** `plutext/docx4j-core-ts` CR-010 phase 2 (Jason asked for phase 2 in the core-ts
session, 2026-10-17). CR-010 keeps `mc:AlternateContent` and saves every branch, docx4j CR-021's
policy. Its phase 1 is done (core-ts `6b178af`); phase 3 puts the content API over kept branches.
**Depends on:** docx4j CR-048 (`CR-048-vml-attributes-word-writes.md`), implemented in docx4j as `acd0defc9` on VERSION_17_3_4, pushed 2026-10-18, to ship in docx4j 17.3.4; regenerated here as `97c43b9`. Nothing else. The measurements below ran on 0.3.2 and `@docx4j/jsonix` 3.4.1.
**Counterpart:** docx4j's `@XmlRootElement` classes and unscoped `ObjectFactory` `@XmlElementDecl`s,
which its JAXB context types in a lax wildcard. Since 2026-10-28 the mapping also declares nine elements docx4j doesn't
root (section 10).

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
2. **Keep `w14:anchorId` on `v:shapetype`.** Done (2026-10-18): docx4j CR-048 declares
   `w14:anchorId` on every VML shape type and `v:group`, and `o:gfxdata` on the nine shape types
   that lacked it, in `vml-main.xsd`, and `97c43b9` regenerates `modules/` from it. The loss was
   in docx4j's own binding too, and is closed there by the same change. (As first written this
   said the fix went through the compiler or `bindings.xjb`; section 8 corrects that: `modules/`
   is generated from docx4j's schemas, so the schema was the place.)

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


## 8. Readiness (2026-10-18)

**Proposal 1 is ready to implement.** Checked against the mapping at 0.3.2: all 61 names resolve
to a type the WML mapping has, none is already global, and each is bound to one type across its
scopes (the two homonyms, `t` and `rPr`, are `m:` elements of the math module under the same local
name, distinct QNames). The facade's one scopeless `getElementInfo` call, `bindsIgnorable`'s
lookup for a literal root, gains answers for the 61 and is unaffected (none binds
`mc:Ignorable`). Section 7 settles the route.

**Proposal 2 was a docx4j schema change, and had to land first or together.** It has landed (see the end of this section). The two findings that decided it:

- The attribute belongs in docx4j's schema, not the compiler or `bindings.xjb`: `modules/` is
  generated from `../docx4j/xsd/vml/vml-main.xsd`, whose `CT_Shapetype` (line 662) and every
  other VML type lack an any-attribute (no VML xsd has one), so the fix is an `xsd:anyAttribute`
  on the core attribute group, or a declared `w14:anchorId`, in docx4j, then a regeneration
  here, the shape of CR-040. That also closes docx4j's own loss (section 3). Which VML elements
  Word writes `w14:anchorId` on decides the form: in the corpus here it is `v:shapetype` only,
  the `v:shape` beside it carrying none. core-ts measured its 459 survey inputs (2026-10-18,
  core-ts CR-010 section 10.2, `328439e`): every attribute on a `v:`, `o:`, `w10:`, `x:` or
  `pvml:` element outside the no-namespace and `o:` namespaces is either `v:ext`, which the
  schema declares, or `w14:anchorId` on `v:shapetype`, four times, all inside an
  `mc:AlternateContent`; none of 12 `v:shape` or the 3 `v:shapetype` outside a branch carries it.
  Thin evidence (the corpus has little WordprocessingML VML; nothing opened in Word for it), but
  it says a declared `w14:anchorId` on `CT_Shapetype` would cover what is seen, where an
  any-attribute on the core group covers what is not. docx4j was told, at Jason's go, by
  the core-ts session; docx4j 17.3.2's `CTShapetype` had neither the property nor an
  `@XmlAnyAttribute`.
- **Shipping proposal 1 alone regresses the facade's round trip.** The facade keeps every
  branch (resolution is core-ts's and the fidelity test's, not the unmarshaller's), and the
  fidelity corpus holds the case: `loadAndSave.docx`'s `document.xml` has a `w:pict` Fallback
  over a `v:shapetype` with `w14:anchorId`, which round-trips whole today and would not with
  `w:pict` typed. CR-004 would report it, and recording it in `KNOWN` would be recording a loss
  this CR introduced. So either docx4j's VML change regenerates here before or with proposal 1,
  or proposal 1 ships without `w:pict` (60 elements; the Choice's `w:drawing` typed, the
  Fallback's `w:pict` DOM until the regeneration), which costs core-ts phase 2 the 9 typed
  `w:pict`s meanwhile. The choice did not arise: CR-048 landed the same day.

**docx4j CR-048 (2026-10-18).** docx4j's own fixtures (36 VML parts in 29 files, 1,839 `v:`
elements) widen the picture: `w14:anchorId` lost on `v:shape` (32, Word 16 regenerating a
Fallback from its Choice), `v:shapetype` (13) and `v:rect` (1); `o:gfxdata` lost on `v:oval` (4)
and `v:rect` (2), which ECMA admits on `CT_Shape` only; [MS-DOCX] 2.2.6 documents `w14:anchorId`
on `w:pict` and `w:object` only, so VML placement is a named departure. CR-048 proposes explicit
`<xsd:attribute ref="w14:anchorId"/>` references (the form `wml.xsd` uses for `CT_Picture` and
`CT_Object`, which the compiler already generates), no any-attribute, `o:spt` left `xsd:float`;
this package's position, sent the same day: the broad set (every VML shape type plus `v:group`)
and `o:gfxdata` on the types lacking it, since optional attributes cost nothing here and the
"seen" set keeps moving.

**CR-048 implemented; the gate is open (2026-10-18).** docx4j `acd0defc9` (VERSION_17_3_4, pushed the same day; ships in docx4j 17.3.4, date not set)
adds `w14:anchorId` to `AG_AllCoreAttributes`, so all eleven VML shape types bind it, and
`o:gfxdata` to the nine that lacked it; `o:spt` stays `xsd:float`. Regenerated here as `97c43b9`:
only the VML mapping changes (eleven `anchorId`, nine `gfxdata`), typecheck, tests and the
fidelity corpus unchanged, as expected while kept branches are DOM. Section 7's probe rerun on it
keeps `w14:anchorId` on the Fallback's `v:shapetype`, and a `v:rect` with both attributes
round-trips through the facade. Word's own saves keep every VML `anchorId` and `gfxdata` (Jason's
Office check, via the docx4j session). So proposal 2 is done by regeneration, the sequencing
question of section 8 is moot, and proposal 1 can ship with all 61, `w:pict` included.

## 9. Implementation notes (2026-10-18)

Proposal 1 is `src/wml-globals.mts`: `WML_GLOBAL_ELEMENTS`, the 61 as local name to type
(lists A and B as data), and `org_docx4j_wml_globals`, the mapping of section 4's recommended form
(`dependencies: ['org_docx4j_wml']`, no `typeInfos`, one `elementInfos` entry per name by qualified
type name). `getContext()` passes it beside the 104 generated modules; `modulesFor` appends it to
any closure holding `org_docx4j_wml`; a caller's own `modules` list gets it only if it includes
it. Both are exported from the facade. The CR-005 registry and `MODULE_NAMES` are unchanged: it is
facade code, not a module to pick.

Tests, in the smoke: each of the 61 resolves with no scope to `org_docx4j_wml.<type>`; the
generated module declares none of them globally itself, so a regeneration that starts to retires
the entry here (strict in both directions, as the fidelity tables are); `modulesFor` adds the
mapping with wml and not without; a `w:r` holding an `mc:AlternateContent` whose Choice carries a
`w:drawing` and whose Fallback a `w:pict` over a `v:shapetype` with `w14:anchorId` unmarshals both
typed, the VML inside typed too, and marshals whole, the attribute kept (docx4j CR-048) and
`o:spt` as read; `<w:drawing/>` unmarshals at the root. The nodenext consumer check imports the
two exports.

Two things the tests exposed, both expected by section 3 and now recorded:

- **A bare `w:tc` no longer carries `xsi:type`.** CR-003's smoke expected
  `<w:tc xsi:type="w:CT_Tc">` from `marshalString(tc(...))`: with no global declaration the runtime
  marshalled the root by its type and said so. Now `w:tc`, `w:tbl`, `w:drawing` and the rest are
  declared, so they marshal as docx4j's root classes do, without one. The expectation changed;
  `deepCopyAs`'s `xsi:type` on a `w:pPr` inside `w:pPrChange` is a different mechanism (a subtype
  where the property declares the base) and is unaffected.
- **The fidelity corpus now types its kept branches**, since the facade keeps every branch and
  CR-004's unresolved check marshals what it unmarshalled. One line differed: the `v:shape`
  inside `loadAndSave.docx`'s `w:pict` Fallback writes `o:gfxdata` unwrapped where Word wraps it
  at 76 characters with CRLF, the same bytes. `test/lib/canonical.mjs` gained a third value rule
  beside the boolean and double ones, with the same argument: two values differing by line breaks
  alone, the remainder base64 of whole quanta, are one `xsd:base64Binary`; spaces and text of the
  wrong shape stay differences, and the self-checks say so. The 359 checks pass; the
  `v:shapetype`'s `w14:anchorId` round-trips because of CR-048 (without `97c43b9` this line would
  have reported it lost, section 8).

Not done here: core-ts's survey of section 5 on this build, and its CR-010 phase 2, which this
unblocks; the 11 non-WordprocessingML roots of section 2 remain a later CR.

**core-ts's survey and verdict (2026-10-18): good to release.** Against an npm pack of `abcb881`
installed over core-ts `328439e`: 837 of 838 tests pass, the one failure the announcer that
asserted a kept `w:drawing` is DOM; parity, XPath parity, typecheck and the nodenext consumer check
green. The section 5 survey over 459 inputs with branches kept: `w:drawing` 13 and `w:pict` 9 now
typed, everything else as before; 281 of 281 parts write every `mc:AlternateContent` back whole
(booleans, `o:gfxdata` whitespace and whitespace-only text normalised), so the four `w14:anchorId`
losses of section 3 are closed, `loadAndSave.docx`'s Fallback `v:shapetype` keeping
`w14:anchorId="7969C50B"` and `o:spt="202"`. Hashes after load, unmarshalAll and save: with
branches resolved, 0 of 8,235 entries differ (the CR-048 regeneration included); with branches
kept, 9 `document.xml` parts differ, all of it booleans `1`/`0` written `true`/`false`, four
`o:gfxdata` unwrapped, attribute order, namespace declarations and whitespace-only text between
elements - nothing added, dropped or changed. Bare-element marshalling over 8,089 elements: `w:p`,
`w:sdt` and `w:tr` unchanged; `w:tbl`, `w:tc`, `w:drawing`, `w:bookmarkStart` and `w:bookmarkEnd`
differ by `xsi:type` and its declaration only, and XML saved with the old `xsi:type` still
unmarshals and re-marshals to the new form (2,574 each way), so scripts saved before the change
read back. No core-ts test or golden expects the `xsi:type`; the editor may marshal bare `w:tbl`
or `w:tc`, and has been told. core-ts's range (`^0.3.2`) did not admit 0.4.0; at Jason's word it
raised to `^0.4.0` in `779b824` (2026-10-18, lockfile resolving 0.4.0 from npm), flipping its
announcer to the typed form: `mce.test.mjs` asserts the typed `w:drawing` and `w:pict` and the
`v:shapetype`'s `w14:anchorId` through a save, and fails on 0.3.2. Suite 838/838, typecheck green;
core-ts CR-010 section 10.4 records the acceptance.

## 10. Nine more, beyond docx4j's roots (CR-017 M7, 2026-10-28)

The rule of this CR was that the mapping declares exactly the elements docx4j declares as roots and the generated
mapping doesn't declare globally. Reconciled on 2026-10-28 against docx4j's generated sources at HEAD: 99 wml root names
(103 `@XmlRootElement` annotations, nested classes included, and 7 unscoped declarations), of which the generated mapping
declares 37, this mapping the 61 of lists A and B, and `w:sdtContent`, with four classes, neither. No root is missing.

CR-017's M7 (d), decided by Jason on 2026-10-28, adds list C: `w:instrText` and `w:delInstrText` (`Text`), `w:object`
(`CTObject`), `w:ruby` (`CTRuby`), `w:footnoteReference` and `w:endnoteReference` (`CTFtnEdnRef`), `w:dir` (`P.Dir`),
`w:bdo` (`P.Bdo`) and `w:subDoc` (`CTRel`). None is a docx4j root. A content list types any element the context
declares globally, so without them one of these misplaced, such as a `w:instrText` straight in a `w:p`, failed the read
of its whole part, where a misplaced `w:t` reads typed. With them it reads typed, is kept, and CR-017's `onMisplaced`
reports it.

This is a divergence from docx4j, recorded here as docx4j asked (docx4j-12, 2026-10-28). docx4j won't root the nine on
its own account: its lax wildcards never legitimately hold them, so a root would type them only where they are
misplaced. It called the divergence harmless to itself and its consumers, since its root set is an artefact of the
schema's global declarations, not a contract. The divergence shows only on misplaced content, which this package keeps
and docx4j's JAXB drops from a strict container. The marshalled output of all 342 fidelity parts is unchanged by it.
