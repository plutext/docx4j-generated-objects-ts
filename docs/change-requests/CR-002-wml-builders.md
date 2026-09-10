# CR-002: `builders/wml`: XML fragments, text sugar and tree traversal for WordprocessingML

**Status:** Proposed 2026-09-10
**Depends on:** compiler CR-010 (the generated `factory/*` and `el/*`, landed 2026-09-10), CR-001
(the prefix table, `NAMESPACE_PREFIXES`)
**Requested by:** `plutext/docx4j-core-ts` CR-002 (the content API), whose phase A this is; that
package carries private copies of the pieces it needs until this lands, then imports them
**Counterpart:** docx4j `XmlUtils.unmarshalString` + `W_NAMESPACE_DECLARATION`, `TextUtils`,
`TraversalUtil`, `ClassFinder`

## 1. Summary

Three hand-written helpers that need only the object model, next to the generated factories:

1. **`wml`**: an XML fragment (one or more sibling elements, as they appear inside
   `document.xml`) to typed elements, with the standard namespace declarations added so the
   author writes `<w:p>` and not `<w:p xmlns:w="...">`.
2. **Text sugar** over `el`: `p('Hello')`, `r('bold', { bold: true })`, `t('  spaced ')` (sets
   `xml:space`), `tbl([['a', 'b']])`, and `textOf(p)` the other way.
3. **Traversal**: `walk(root, visitor)` over every typed object, `find(root, typeName)`,
   `linkParents(value, parent)`.

### 1.1 Relationship to `helpers/wml`

`src/helpers/wml.mts` (`./helpers/wml`, compiler CR-007 item 5) exists already and stays as
it is. It holds the decisions docx4j's `ModifyGeneratedSources` adds to the *generated* Java
model, property by property: the `w:highlight` colour table (`highlightHexValue`,
`highlightNameForColor`), `isQFormat` with its `defQFormat` fallback, `isCustomStyle`. Those
are facts about individual types that the schema does not state; they belong next to the
declarations and are maintained against CR-007.

`builders/wml` is the other side of docx4j: what `docx4j-core`'s `XmlUtils`, `TextUtils` and
`TraversalUtil` do with a tree once it exists (parse a fragment, build content, read text,
walk, link parents). Nothing in it overlaps with `helpers/wml`: `textOf`, `walk`, `find` and
`linkParents` have no counterpart there, and the constructors are new. The builders import
the helpers where a property needs them (`r(text, { highlightColor: '#FFFF00' })` goes through
`highlightNameForColor`; `p(text, { style })` does not check `isCustomStyle`). Two subpaths
rather than one keep the maintenance boundary visible: a change to `helpers/wml` follows a
change in docx4j's generated-source adjustments, a change to `builders/wml` follows
docx4j-core or the content API in core-ts. If the split ever reads as noise, `builders/wml`
can re-export `helpers/wml` so that one import serves both; the reverse direction (helpers
importing builders) is ruled out.

## 2. `wml`

```ts
// One function, callable both ways; a tagged call is recognised by its TemplateStringsArray.
export function wml(xml: string, options?: WmlOptions): Promise<TypedNamedValue[]>;            // the siblings
export function wml(strings: TemplateStringsArray, ...values: Interpolated[]): Promise<TypedNamedValue[]>;   // no options: core-ts uses the plain form
export function wmlOne<T = unknown>(xml: string, options?: WmlOptions): Promise<TypedNamedValue<T>>;   // exactly one, else throws
export function wmlOne<T = unknown>(strings: TemplateStringsArray, ...values: Interpolated[]): Promise<TypedNamedValue<T>>;
wml.raw(xml: string): Raw;                                                                     // marks a string as XML for interpolation

export interface WmlOptions {
  wrapper?: 'body' | 'p' | 'tr' | 'tc' | 'r' | 'none';   // force one instead of inferring from the first element
  preprocess?: (doc: Document) => void;                    // runs over the parsed DOM before unmarshalling (core-ts passes its MCE preprocessor)
}
/** A string marked as XML for interpolation; branded so that a plain string can never be mistaken for it. */
export interface Raw { readonly __wmlRaw: true; readonly xml: string }
type Interpolated = string | number | TypedNamedValue | Raw | Interpolated[];
```

Interpolation, in the tagged form: a `{ name, value }` pair is marshalled in place (so
`wml\`<w:p>${run}</w:p>\`` composes typed and textual content); a string or number is
**escaped as text** (`&`, `<`, `>` and quotes), so `${userInput}` cannot break the fragment or
smuggle markup in; `wml.raw(xml)` is inserted verbatim for callers who really hold XML; an
array interpolates each item in turn. The plain-string form is what the tagged form produces
after interpolation, so both meet in one implementation.

The runtime's `unmarshalDocument` resolves **global** elements only (`w:p`, `w:body`,
`w:document`, `w:settings` are global; `w:tbl`, `w:r`, `w:tr` are declared in a scope and
cannot be unmarshalled standalone: verified 2026-09-10, `w:tbl` throws "not known in this
context"). So the fragment is wrapped in the global container docx4j would put its first
element in, then that container's content is returned. (Not "the container the model
accepts": the generated model is permissive, a bare `w:r` or a run-level `w:sdt` unmarshals
under `w:body` without complaint, verified 2026-09-10, so acceptance decides nothing.)

| First element of the fragment | Wrapper | Returned |
|---|---|---|
| `w:p`, `w:tbl`, `w:altChunk`, ... (block level only) | `<w:body>` | `body.content` |
| `w:r`, `w:hyperlink`, `w:fldSimple`, `w:smartTag`, ... (run level only) | `<w:p>` | `p.content` |
| `w:sdt`, `w:customXml` (a different type at each level: `SdtBlock` / `SdtRun`, `CTCustomXmlBlock` / `CTCustomXmlRun`, and the row and cell forms) | by the first decisive descendant (below) | that wrapper's content |
| `w:bookmarkStart` / `End`, `w:proofErr`, `w:ins`, `w:del`, `w:moveFrom` / `To`, `w:commentRangeStart` / `End`, `w:permStart` / `End` (the same type at both levels) | `<w:body>` | `body.content` |
| `w:tr` | `<w:body><w:tbl>` | `tbl.content` |
| `w:tc` | `<w:body><w:tbl><w:tr>` | `tr.content` |
| `w:t`, `w:br`, `w:tab`, `w:drawing`, ... | `<w:p><w:r>` | `r.content` |
| a global element (`w:document`, `w:styles`, ...) | none | `[element]` |

The wrapper's element declares every prefix of `NAMESPACE_PREFIXES`, so `w14:paraId`,
`r:id`, `wp:inline` and `mc:AlternateContent` parse; declarations inside the fragment win.
The returned values have `PARENT` linked among themselves; the top-level ones point at the
throw-away wrapper, and the caller (core-ts `insertElement`) re-links them to the real
container. `wrapper` overrides the inference for every row. Malformed XML throws the
parser's error; an element the model does not know throws the runtime's.

**Content controls and `w:customXml`** are the one place where the wrapper changes the type:
a run-level `w:sdt` unmarshalled under `w:body` comes back as `SdtBlock` with a
`SdtContentBlock`, not `SdtRun` (verified 2026-09-10; trying `w:body` first and falling back
to `w:p` therefore never falls back, and was wrong). And the first child of `w:sdtContent` is
often another `w:sdt` (repeating sections, grouped controls), so looking one level down does
not settle it either. The wrapper is chosen by the **first decisive descendant**: scan the
children of `w:sdtContent` (or of `w:customXml`) in document order; a block-only element
(`w:p`, `w:tbl`, `w:altChunk`) decides `body`, a run-only element (`w:r`, `w:hyperlink`,
`w:fldSimple`, `w:smartTag`, `w:dir`, `w:bdo`) decides `p`, `w:tr` decides `tr` and `w:tc`
decides `tc` (the row and cell controls); a nested `w:sdt` or `w:customXml` is recursed into
and its decision, if any, is the answer; an element with the same type at both levels
(`w:bookmarkStart`, `w:proofErr`, `w:ins` holding runs is schema-legal in both) decides
nothing and the scan continues; nothing decisive anywhere (properties only, an empty
repeating section, a chain of empty nested controls) is `body`. The scan runs over the DOM,
which is parsed anyway for `preprocess`, and only the outermost decision is needed: once the
outer wrapper is right the runtime types every nested control from its container.

MCE: `wml` owns both steps, parsing and unmarshalling, so the caller cannot touch the DOM in
between; the `preprocess` hook is that point. Without it a fragment carrying
`mc:AlternateContent` is typed as `AlternateContent` where the model allows it and throws
where it does not (a `w:drawing` inside `mc:Choice` is a local element), as `unmarshalString`
does today; core-ts passes its MCE preprocessor through the hook when it inserts XML into a
part, so the fragment is resolved exactly as the part's own content was.

## 3. Text sugar

```ts
p(text: string, opts?: RunOptions & { style?: string }): TypedNamedValue<P>     // one run; `style` here is the paragraph style (w:pStyle), a run style is `runStyle`
p(runs: TypedNamedValue<R>[], opts?): TypedNamedValue<P>
r(text: string, opts?: RunOptions): TypedNamedValue<R>
// RunOptions, UnderlineType, the underline table, applyRunOptions and readRunOptions are DEFINED HERE
// (moved from core-ts src/model/content/runOptions.mts, where they were written on 2026-09-10 to
// make the move a file move): core-ts depends on this package, so the one mapping has to live here.
export type UnderlineType = 'None' | 'Single' | 'Word' | 'Double' | 'Thick' | 'Dotted' | 'DottedHeavy' | 'DashLine' | 'DashLineHeavy' | 'DashLineLong' | 'DashLineLongHeavy' | 'DotDashLine' | 'DotDashLineHeavy' | 'TwoDotDashLine' | 'TwoDotDashLineHeavy' | 'Wave' | 'WaveHeavy' | 'WaveDouble' | 'Mixed';
export const UNDERLINE_TO_WML: Record<Exclude<UnderlineType, 'None' | 'Mixed'>, UnderlineEnumeration>;   // seventeen entries, and the inverse map
export interface RunOptions {   // the vocabulary of core-ts's Font view (Office JS Word.Font); undefined leaves a property alone
  bold?: boolean; italic?: boolean;
  underline?: boolean | UnderlineType;   // true is 'Single'; false or 'None' removes w:u
  strikeThrough?: boolean; doubleStrikeThrough?: boolean; subscript?: boolean; superscript?: boolean;
  name?: string;                // w:rFonts ascii and hAnsi; clears the theme font attributes; '' removes w:rFonts
  size?: number;                // points; w:sz and w:szCs are half-points
  color?: string;               // '#RRGGBB' or 'auto'; '' removes w:color
  highlightColor?: string | null;   // a highlight name or one of its '#RRGGBB' values, through helpers/wml highlightNameForColor; null or '' removes; anything else throws
  style?: string;               // w:rStyle; '' removes
}
export function applyRunOptions(rPr: RPr, opts: RunOptions): RPr;        // writes in place: bold is w:b and w:bCs, italic w:i and w:iCs, size w:sz and w:szCs, sub/superscript share w:vertAlign
export function readRunOptions(rPr: RPr | undefined): RunFormatting;   // the inverse, direct formatting only
export interface RunFormatting {   // every RunOptions property, present; underline narrowed to its names (never a boolean on the read side)
  bold: boolean; italic: boolean; underline: UnderlineType; strikeThrough: boolean; doubleStrikeThrough: boolean; subscript: boolean; superscript: boolean;
  name: string;                 // '' when not set directly
  size: number;                 // points; 0 when not set directly
  color: string;                // '#RRGGBB', 'auto' or a theme name; '' when not set directly
  highlightColor: string | null;   // '#RRGGBB'; null when there is none
  style: string;                // '' when not set
}
t(text: string): TypedNamedValue<Text>          // xml:space="preserve" when the text starts or ends with whitespace or holds two spaces
br(type?: 'page' | 'column' | 'textWrapping'): TypedNamedValue<Br>
tab(): TypedNamedValue<R.Tab>
tbl(rows: string[][], opts?: { style?: string; widths?: number[] /* twips */ }): TypedNamedValue<Tbl>   // tblGrid from the widths or equal columns
textOf(value: P | R | Tc | Tbl | Body | Hdr | Ftr | ...): string   // w:t, w:tab (\t), w:br and w:cr (\n), w:noBreakHyphen, w:softHyphen, w:sym; runs inside hyperlinks, sdt, ins; del skipped; paragraphs joined with \n (docx4j TextUtils)
```

All built on `el` so there is one source of QNames; every returned value carries `TYPE_NAME`.
`RunOptions` uses the property names of core-ts's `Font` (Office JS `Word.Font`) rather than
the schema's (`name` not `rFonts`, `highlightColor` not `highlight`), so that what a builder
sets, a view reads back under the same name. The names alone would still leave two mappings,
so the mapping is one function pair, `applyRunOptions` / `readRunOptions`, defined here:
`r()` calls `applyRunOptions` on a fresh `w:rPr`, and core-ts's `Font` setters and getters
call the pair on each run in scope. core-ts wrote the pair first (its
`src/model/content/runOptions.mts`, 2026-09-10) so that this CR's phase A moves the file
rather than re-implementing it. `t` splits nothing; `r(text)` turns a `\t` or `\n` into
`w:tab` / `w:br` items, as docx4j's `addParagraphOfText` does not but users expect.

## 4. Traversal

```ts
walk(root: unknown, visitor: (value: object, parent: object | undefined, key: string | number) => boolean | void): void
  // depth first over typed objects (TYPE_NAME); { name, value } pairs and arrays are descended, not visited; QNames, calendars and DOM nodes are not entered at all (package parts are DOM by the schema); return false to not descend (docx4j TraversalUtil)
find<T>(root: unknown, typeName: string): T[]                    // ClassFinder: every object whose TYPE_NAME matches
linkParents(value: object, parent: object | undefined): void      // sets non-enumerable PARENT on value and, recursively, its typed descendants (what the unmarshaller does; deepCopy copies, this does not)
```

`PARENT` is defined as the unmarshaller defines it (non-enumerable, writable, configurable;
verified 2026-09-10), so linked and unmarshalled objects are indistinguishable.

## 5. Layout, tests

`src/builders/wml.mts` compiled to `dist/builders/wml.mjs`, exported as
`@docx4j/generated-objects-ts/builders/wml`. Tests in `test/smoke.mjs`:

- `wml`: each wrapper row of section 2 (a fragment in, `TYPE_NAME` and element names out,
  `PARENT` linked below the top level); the tagged form with an interpolated typed value, an
  escaped string and a `wml.raw`; `wmlOne` throwing on two siblings.
- The content control inference, since the wrapper decides the type there: a `w:sdt` holding
  `w:r` yields `SdtRun` and, with `wrapper: 'body'`, `SdtBlock`; a `w:sdt` holding a `w:sdt`
  holding `w:p` yields `SdtBlock` outside and in, the same with `w:r` inside yields `SdtRun`
  twice; a `w:bookmarkStart` before the run does not change that; a `w:sdt` holding `w:tr`
  yields the row control; an empty `w:sdtContent` and a properties-only `w:sdt` yield
  `SdtBlock`; `w:customXml` holding `w:r` yields `CTCustomXmlRun`.
- Text sugar: `textOf` on the README example; `t` with leading, trailing and double spaces
  (`xml:space`); `r('a\tb')` giving `w:tab`; `tbl` grid from widths and from equal columns.
- The run mapping, since after the move it lives here and a regeneration that renamed an
  `RPr` property must surface here, not downstream: `applyRunOptions` then `readRunOptions`
  round-trips every option; `underline: true` reads back as `Single` and `false` as `None`;
  `color: 'auto'`; each removal value (`false`, `''`, `null`, `'None'`) deletes its property;
  `highlightColor` with a non-highlight value throws; `readRunOptions(undefined)` is all
  defaults; and `r(text, opts)` produces the same `w:rPr` as `applyRunOptions({}, opts)`, that
  call being the builder's only use of the pair. core-ts's `content.test.mjs` has the same
  round trip through `Font` and keeps it as the consumer-side check.
- Traversal: `walk` with a skip, `find`, `linkParents` then `find` from the parent, and
  `walk` not entering a DOM node.

README: a paragraph under the factories section.

## 6. Open questions (decided 2026-09-10)

1. `p(text)` with `\n`: one paragraph with `w:br` items, as `r(text)` already does; `wml` and
   core-ts `insertParagraph` are the ways to get several paragraphs.
2. `wml` does not accept the `<pkg:package>` form: it is a fragment parser and stays tree-only.
   Flat OPC is the facade's already (`unmarshalPackage` / `marshalPackage` in `src/index.mts`;
   core-ts's `FlatOpcPartStore` sits above them). The case behind the question stands, though:
   Word's `getOoxml()` for something on the document surface returns `document.xml` plus
   styles, numbering, settings, fonts, theme and web settings, and when the snippet carries no
   relationships and styles need no resolving, only `document.xml` is wanted. That is a
   follow-up against the facade, not this CR: `unmarshalPackage(ooxml, { parts:
   ['/word/document.xml'] })`, typing the named parts only and leaving the rest as DOM (what the
   function already does for roots the model does not know). Whether it is worth having is a
   measurement, typing being the step that grows with the part count; measure on a real
   `getOoxml()` payload first, then propose it as CR-003 if it pays.
