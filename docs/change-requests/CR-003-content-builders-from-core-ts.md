# CR-003: `builders/wml` additions from core-ts's content API, and `deepCopyAs`

**Status:** Proposed 2026-09-16
**Depends on:** CR-002 (`builders/wml`), compiler CR-010 (the generated `factory/*` and `el/*`)
**Requested by:** `plutext/docx4j-core-ts` CR-002 (the content API), phases C, E, F, G and I, which
worked around each gap below with a private copy and recorded it ("Objects-package gaps" in its
sections 8, 10, 12 and 13). Every copy is marked there as the code to delete when this lands.
**Counterpart:** docx4j `BinaryPartAbstractImage.createImageInline`, `TblFactory`,
`XmlUtils.deepCopy`, `TraversalUtil`, and the content-control code in `org.docx4j.model.sdt`

## 1. Summary

core-ts's content API needed ten things that use only the object model, so, by this repository's
rule (tree-only helpers here, anything needing parts or relationships in core-ts), they belong
here. Each is a port of code that already runs in core-ts:

| core-ts gap | Here | core-ts source |
|---|---|---|
| deepCopy of `w:pPr` into `w:pPrChange` writes `xsi:type="w:CT_PPr"` | `deepCopyAs(value, typeName)` (facade) | `tracking.mts` (sets `TYPE_NAME` by hand) |
| no `w:sdt` builder (block, run, row, cell) | `sdt`, `sdtPr`, `nextSdtId` | `model/customxml/insert.mts` |
| `SdtPr.rPrOrAliasOrLock` is searched by local name and cast | `sdtProperty`, `sdtKindOf` | `ContentControl.findProperty`, `TYPE_BY_ELEMENT` |
| no inline-picture constructor | `inlinePicture(relId, options)` | `InlinePicture.mts` `drawingFor` |
| `tbl` builds whole tables only | `tr`, `tc` (and `tbl` over them) | `rowElement()` |
| run lists live under three property names | `runItemsOf` | `model/content/tree.mts` |
| `w:rPr` (named) and `w:rPrChange/w:rPr` (element list) do not convert | `rPrToElements`, `rPrFromElements` | `tracking.mts` `rPrElements` / `rPrFromElements` |
| `walk` does not enter DOM held by `xs:any` | `walkAll(root, visitor, domVisitor)` | reference rewriting in `insertOoxml` |
| `toApiScript` falls back to a marshalled XML string | `toSource`, `isSugarExpressible` | `office-js/toApiScript.mts` |

A bug found while checking, in the published 0.1.2:

- **`textOf` drops text inside `w:moveTo` (and `w:moveFrom`).** `RunTrackChange` keeps its runs under
  `accOrBarOrBox`, which `textOf`'s child lookup does not read: for
  `<w:r><w:t>kept </w:t></w:r><w:moveTo><w:r><w:t>moved</w:t></w:r></w:moveTo><w:ins><w:r><w:t> inserted</w:t></w:r></w:ins>`
  it returns `"kept  inserted"`. **Fixed 2026-09-16** ahead of the rest of this CR, with the
  `runItemsOf` of section 3.4: `textOf` now reads `w:moveTo` and skips `w:moveFrom`, and
  `runItemsOf` is exported. That part of phase A is therefore done; the rest is unimplemented.

Two corrections to the core-ts notes, checked 2026-09-16 against 0.1.2:

- **The w14 run effects do have wrappers**, in the factory rather than in `el`: `el/org_docx4j_wml`
  holds only elements global to wml's namespace, while `factory/org_docx4j_wml` has the scoped
  wrappers for every w14 member of `EG_RPrBase` (46 of them: `createCTRPrChangeRPrGlow`,
  `createCTParaRPrOriginalLigatures`, ...). `rPrToElements` uses them, so nothing is dropped.
- **`SdtPr.rPrOrAliasOrLock` is typed** element by element; what is missing is an accessor. The
  single choice list is also docx4j's model (`SdtPr.getRPrOrAliasOrLock()`), so this CR adds
  helpers rather than asking the compiler to split the property.

Gaps that are not tree helpers are routed elsewhere (section 7).

## 2. `deepCopyAs` (facade)

```ts
export function deepCopyAs<T extends { TYPE_NAME?: string }>(value: object, typeName: NonNullable<T['TYPE_NAME']>, parent?: unknown): T;
```

`deepCopy`, then the copy's `TYPE_NAME` is set to `typeName` and the copy's own properties that the
target type does not declare are removed (the marshaller would ignore them; removing them keeps the
tree honest). `typeName` must be the value's type or one of its base types in the context
(`PPr` to `PPrBase`, `RPr` to `ParaRPr`'s base, ...); anything else throws. Children keep their own
types. The `w:pPrChange` case becomes `pPrChange.pPr = deepCopyAs<PPrBase>(pPr, 'org_docx4j_wml.PPrBase')`,
which drops `rPr`, `sectPr` and `pPrChange` from the copy and marshals `<w:pPr>` without `xsi:type`,
as Word writes it. docx4j: `XmlUtils.deepCopy` into a `PPrBase` built by hand.

## 3. `builders/wml` additions

### 3.1 Content controls

```ts
export type SdtKind = 'RichText' | 'PlainText' | 'Picture' | 'BuildingBlockGallery' | 'CheckBox' | 'ComboBox'
  | 'DropDownList' | 'DatePicker' | 'RepeatingSection' | 'RepeatingSectionItem' | 'Group' | 'Citation' | 'Bibliography' | 'Equation';
export type SdtForm = 'block' | 'run' | 'row' | 'cell';
export interface SdtOptions { kind?: SdtKind; id?: number; tag?: string; title?: string; form?: SdtForm }

export function sdtPr(options?: SdtOptions): M.SdtPr;
export function sdt(content: Element[], options?: SdtOptions): Element<M.SdtBlock | M.SdtRun | M.CTSdtRow | M.CTSdtCell>;
export function nextSdtId(root: unknown): number;
export function sdtProperty<T = unknown>(sdtPr: M.SdtPr | undefined, localPart: string, namespaceURI?: string): Element<T> | undefined;
export function sdtKindOf(sdtPr: M.SdtPr | undefined): SdtKind | 'Unknown';
```

- Kind names are Office JS's `Word.ContentControlType` values, so core-ts passes them through.
  Rich text is the untyped control (no kind element), as Word writes it. The kind element is the
  one core-ts's `kindElement` writes, including the w14 checkbox with its MS Gothic ballot boxes
  and the date picker's `dateFormat`/`storeMappedDataAs`. Since docx4j CR-018 the checkbox's
  `w14:checked` is a real boolean (`{ val: false }`), not one of the string spellings.
- `sdtPr` writes `w:alias` (title), `w:tag`, `w:id` and the kind element, in that order.
- `sdt`'s form is inferred from the content, as `wml` infers a wrapper (paragraphs and tables:
  block; runs, hyperlinks, run-level controls: run; `w:tr`: row; `w:tc`: cell) unless `form` is
  given. `RepeatingSection` with a run form throws, as core-ts's `checkKind` does.
- `id` defaults to a random 31-bit id. Builders see no document, so `nextSdtId(root)` gives one
  unused among the `w:id`s of the controls under `root` (core-ts's `nextControlId`).
- `sdtProperty` finds a `w:sdtPr` child by name (default namespace wml; w14 and w15 for the
  checkbox, appearance, repeating section). `sdtKindOf` is the inverse of the kind element
  (`docPartObj` and `docPartList` are both `BuildingBlockGallery`).

### 3.2 Inline pictures

```ts
export interface InlinePictureOptions { cx: number; cy: number; id: number; name: string; descr?: string; title?: string }
export function inlinePicture(relId: string, options: InlinePictureOptions): Element<M.Drawing>;
```

A `w:drawing` with one `wp:inline` and a `pic:pic` whose `a:blip/@r:embed` is `relId`, extents in
EMU, `noChangeAspect`, a `rect` preset and a stretch fill, as docx4j's `createImageInline` writes
it (core-ts `drawingFor`, a copy over the generated factories). The relationship is a string here;
creating the image part and the relationship stays in core-ts. `title` sits beside `descr`, which
docx4j CR-018 added to `wp:docPr` (section 7).

### 3.3 Rows and cells

```ts
export interface CellOptions { width?: number }
export function tc(blocks: string | Element[], options?: CellOptions): Element<M.Tc>;
export function tr(cells: (string | Element<M.Tc>)[], options?: { widths?: number[] }): Element<M.Tr>;
```

A string cell is one paragraph of text, as in `tbl`; a cell always holds at least one paragraph
(an empty `w:p` when `blocks` is empty), as Word requires. `tbl` is re-expressed over `tr` and `tc`
with byte-identical output (the existing smoke assertions stay).

### 3.4 Run items

```ts
export function runItemsOf(value: object): Element[] | undefined;
```

The run-level list a holder keeps, under docx4j's property names: `content` for most,
`customXmlOrSmartTagOrSdt` for `w:ins` and `w:del`, `accOrBarOrBox` for `w:moveFrom` and `w:moveTo`
(`RunTrackChange`), `sdtContent.content` for a run-level control. `runItemsOf` is purely structural: it returns whatever list the
holder keeps and skips no kind, so a caller can read the original of a revision (core-ts's
`{ view: 'original' }` needs `w:moveFrom` and `w:del`). The filtering stays in `textOf`, whose own
child lookup learned `accOrBarOrBox` at the same time (done 2026-09-16): moved-to text is read like
inserted text, and `w:moveFrom` is skipped as `w:del` is.

### 3.5 Run properties as an element list

```ts
export function rPrToElements(rPr: M.RPr): NonNullable<M.CTRPrChange.RPr['egrPrBase']>;
export function rPrFromElements(list: M.CTRPrChange.RPr | M.CTRPrChange.RPr['egrPrBase'] | undefined): M.RPr;
```

`w:rPrChange/w:rPr` (`CTRPrChange.RPr`) keeps `EG_RPrBase` as an element list; `w:rPr` has named
properties. `rPrToElements` deep-copies each named property into its element in the schema's
`EG_RPrBase` order, the w14 effects included (`glow`, `shadow`, `reflection`, `textOutline`,
`textFill`, `scene3d`, `props3d`, `ligatures`, `numForm`, `numSpacing`, `stylisticSets`,
`cntxtAlts`), through the factory's scoped `createCTRPrChangeRPr*` wrappers. `rPrFromElements` is
the inverse and ignores `w:rPrChange` itself. The same pair serves `w:pPr/w:rPr` (`ParaRPr`) and
`CTParaRPrOriginal` in a later revision if core-ts needs it.

### 3.6 Walking into DOM

```ts
export function walkAll(root: unknown, visitor: (value: object, parent: object | undefined, key: string | number) => boolean | void,
  domVisitor: (node: Element, owner: object, key: string | number) => void): void;
```

`walk` plus the DOM elements held by `xs:any` properties (`any`, `xmlData`, ...): `domVisitor` is
called for each such element and its descendant elements (walked with `childNodes`, since xmldom
has no `firstElementChild`). `walk` keeps its documented behaviour. Use: rewriting `r:embed`,
`r:id` and `r:link` wherever they are, such as an SVG twin's `asvg:svgBlip` in `a:extLst`.

### 3.7 Source

```ts
export interface ToSourceOptions { variable?: string; factory?: 'el' | 'sugar' }
export function toSource(element: Element, options?: ToSourceOptions): string;
export function isSugarExpressible(element: Element): boolean;
```

`toSource` returns TypeScript statements that rebuild the element: `const p1 = p([r('x', { bold: true })]);`
where the sugar expresses it (`factory: 'sugar'`, the default when `isSugarExpressible`), else the
`el.*` calls with literals. Names are stable (`p1`, `r1`, `tbl1`, ...; `variable` names the root),
so core-ts's `toApiScript` can emit `body.insertElement(<source>, 'End')` instead of an XML string.
`isSugarExpressible` is true when `p`/`r`/`t`/`br`/`tab`/`tbl` and `readRunOptions` round-trip the
element exactly (the run formatting is the inverse of `applyRunOptions`). XML stays the fallback
only for what neither form types (DOM held by `xs:any`).

## 4. Phases

- **A**: sections 2 and 3.1 to 3.6. Ports of running code; core-ts deletes its copies. In the order
  core-ts asked for (2026-09-16), by the size of the copy each removes: `sdt` / `sdtPr` /
  `nextSdtId` / `sdtProperty` / `sdtKindOf` (its phase E `insert.mts`), then `tr` / `tc` and
  `inlinePicture` (phase C), then `rPrToElements` / `rPrFromElements` and `deepCopyAs` (phase F),
  then `walkAll`. The `runItemsOf` part landed early with the `textOf` fix (section 1).
- **B**: section 3.7. New code with its own design questions (literal formatting, calendars, QNames).
  No deadline: core-ts's `toApiScript` fallback works with marshalled XML.

## 5. Layout and exports

All of section 3 goes in `src/builders/wml.mts` (exported from `./builders/wml`); `deepCopyAs` in
`src/index.mts` (`.`). No new public paths. `test/nodenext/consumer.mts` gains a
`@ts-expect-error` line for one new export, and `test/readme-examples.ts` a snippet if the README
shows one. Version 0.2.0 (new API, no breaking change).

## 6. Tests

`test/smoke.mjs`:

- `deepCopyAs`: `w:pPr` into `w:pPrChange` marshals without `xsi:type`; `rPr`/`sectPr` dropped from
  the copy; the original untouched; a non-ancestor type throws.
- `sdt` in the four forms, each kind's element, `sdtPr` order, form inference and the
  repeating-section check; `nextSdtId` avoiding ids in the tree; `sdtProperty` for wml and w14
  names; `sdtKindOf` for every kind, round trip with `sdt`.
- `inlinePicture` marshals the `wp:inline` docx4j writes (compared with a fixture).
- `tr`/`tc`: string and element cells, an empty cell gets a paragraph; `tbl` output unchanged.
- `runItemsOf` for `P`, `w:ins`, `w:del`, `w:moveTo`, `w:moveFrom`, a run-level control; `textOf`
  reads text inside `w:moveTo` (the section 1 case gives `"kept moved inserted"`) and skips `w:moveFrom`.
- `rPrToElements` / `rPrFromElements` round trip with every `EG_RPrBase` member including the w14
  effects, in schema order.
- `walkAll` reaches an `r:embed` inside an `a:extLst` held as DOM.
- Phase B: `toSource` output evaluated back to an element marshals identically, for sugar and `el`.

## 7. Routed elsewhere

**Done**: docx4j CR-018 (merged as `33e9cb411` on `VERSION_17_1_1`) made all of it a schema change,
and the regeneration of 2026-09-16 brought it here:

- `ignorable` on `Comments`, w15 `CTCommentsEx` and `CTPeople` (a re-marshalled comments part keeps
  Word's `mc:Ignorable` again, checked through the facade);
- `title` on `CTNonVisualDrawingProps`, so `inlinePicture` (section 3.2) takes `title` from the
  start and Office JS's `altTextTitle` round-trips;
- `contact` optional on w15 `CTPerson`;
- `otherAttributes` (`xsd:anyAttribute lax`) on `DatastoreItem`;
- w14 `CTOnOff.val` is now `boolean`, so a checkbox reads `checked.val === true` rather than the
  three string spellings; `sdtPr`'s checkbox element (section 3.1) writes `{ val: false }`;
- two new modules, `org_docx4j_w16cex` (`w16cex:commentsExtensible`, a typed
  `word/commentsExtensible.xml`) and `org_docx4j_w16`, with their `MODULE_NAMES` entries: 96 modules.

core-ts can drop its DOM handling of `commentsExtensible.xml` once it depends on that release.

## 8. Consumers

`@docx4j/core-ts` imports these and deletes its copies: `kindElement`, `sdtPrFor`, `sdtBlockFor`,
`sdtRunFor`, `nextControlId`, `checkKind` (`insert.mts`); `findProperty` and `TYPE_BY_ELEMENT`
(`ContentControl.mts`); `drawingFor` (`InlinePicture.mts`); `rowElement`; `runItemsOf`
(`tree.mts`); `rPrElements` / `rPrFromElements` (`tracking.mts`); the `TYPE_NAME` override for
`w:pPrChange`; and, in phase B, `toApiScript`'s XML fallback.

## 9. Open questions

1. `walkAll` as a separate function, or an options object on `walk` (whose positional `parent` and
   `key` parameters make a third-position options argument a breaking change)?
2. Should `sdt` accept Office JS's `Word.ContentControlAppearance` (`w15:appearance`) and colour now,
   or leave them to core-ts's view?
3. `toSource`'s output for calendars, QNames and `xs:any` DOM: literals, helper calls, or a refusal?
