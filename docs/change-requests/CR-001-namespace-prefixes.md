# CR-001: The facade ships docx4j's namespace prefix table as the context default

**Status:** Implemented 2026-09-10 (sections 2.1 and 2.3 amended on implementation: one default namespace per marshal, chosen by the root, as docx4j's per-part mappers)
**Depends on:** `@docx4j/jsonix` 3.2.0 (`Jsonix.Context` option `namespacePrefixes`, already supported)
**Requested by:** `plutext/docx4j-core-ts` CR-001 (the engine), whose Phase A cannot write a
Word-readable `document.xml` without this
**Counterpart:** docx4j `org.docx4j.jaxb.NamespacePrefixMappings` (the table),
`NamespacePrefixMapper` / `NamespacePrefixMapperRelationshipsPart` (how it is applied),
`McIgnorableNamespaceDeclarator` (the `mc:Ignorable` half)

## 1. Problem

`getContext()` builds the shared `Jsonix.Context` without a `namespacePrefixes` table, so the
marshaller invents prefixes. Reproduced on 2026-09-10 with the current facade:

```
in:  <w:document xmlns:w=".../wordprocessingml/2006/main" xmlns:w14=".../word/2010/wordml"
       xmlns:mc=".../markup-compatibility/2006" mc:Ignorable="w14">
       <w:body><w:p w14:paraId="1"><w:r><w:t xml:space="preserve"> x</w:t></w:r></w:p></w:body></w:document>

out: <w:document xmlns:w="..." p1:Ignorable="w14" xmlns:p1=".../markup-compatibility/2006">
       <w:body><w:p p0:paraId="1" xmlns:p0=".../word/2010/wordml"><w:r>
       <w:t p2:space="preserve" xmlns:p2="http://www.w3.org/XML/1998/namespace"> x</w:t></w:r></w:p></w:body></w:document>
```

Three things are wrong with the output, in increasing order of severity:

1. The prefixes are not the conventional ones (`w14`, `mc`), so the XML is hard to read and to
   diff against what Word writes. Cosmetic on its own.
2. `mc:Ignorable="w14"` names a prefix, and ECMA-376 Part 3 requires every prefix named there to
   be declared on the element carrying the attribute. Here `w14` is declared nowhere (the
   attribute that uses the namespace got `p0` instead), so Word reports the file as corrupt.
   This is the blocker for the engine.
3. `p2:space` binds a new prefix to the `xml` namespace, which the Namespaces in XML
   recommendation forbids (only the reserved prefix `xml` may be bound to it). Word rejects it;
   `@xmldom/xmldom` and browsers happen to let it through, which is why the smoke tests did not
   notice.

The engine could pass a table itself (`getContext({ namespacePrefixes })`), but the context is
built by whoever calls `getContext` first and the options of later callers are ignored. A
consumer that unmarshals something through the facade before loading a package would silently
defeat the engine's table. The default has to live here. This CR is the correction to
docx4j-core-ts CR-001 section 9, which claimed the facade needed nothing new.

## 2. Proposal

### 2.1 `NAMESPACE_PREFIXES`

`src/index.mts` exports docx4j's prefix table as a frozen object, namespace URI to prefix:

```ts
export const NAMESPACE_PREFIXES: Readonly<Record<string, string>> = Object.freeze({
  'http://www.w3.org/XML/1998/namespace': 'xml',
  'http://www.w3.org/2001/XMLSchema-instance': 'xsi',
  'http://www.w3.org/2001/XMLSchema': 'xs',
  'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
  'http://schemas.openxmlformats.org/package/2006/relationships': 'rel',    // 'rel' as docx4j's static table; the default namespace when the root is Relationships (see 2.3)
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main': 'w',
  'http://schemas.openxmlformats.org/presentationml/2006/main': 'p',
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main': '',          // SML uses the default namespace, as docx4j; 's' when the root is Relationships (see 2.3)
  'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
  'http://schemas.openxmlformats.org/drawingml/2006/picture': 'pic',
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing': 'wp',
  'http://schemas.openxmlformats.org/drawingml/2006/chart': 'c',
  'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing': 'cdr',
  'http://schemas.openxmlformats.org/drawingml/2006/diagram': 'dgm',
  'http://schemas.openxmlformats.org/drawingml/2006/compatibility': 'comp',
  'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas': 'lc',
  'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing': 'xdr',
  'http://schemas.openxmlformats.org/officeDocument/2006/math': 'm',
  'http://schemas.openxmlformats.org/officeDocument/2006/bibliography': 'b',
  'http://schemas.openxmlformats.org/officeDocument/2006/customXml': 'ds',
  'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties': 'prop',
  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties': 'properties',
  'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes': 'vt',
  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties': 'cp',
  'http://schemas.openxmlformats.org/package/2006/digital-signature': 'mdssi',
  'http://schemas.openxmlformats.org/schemaLibrary/2006/main': 'sl',
  'http://purl.org/dc/elements/1.1/': 'dc',
  'http://purl.org/dc/terms/': 'dcterms',
  'urn:schemas-microsoft-com:vml': 'v',
  'urn:schemas-microsoft-com:office:office': 'o',
  'urn:schemas-microsoft-com:office:word': 'w10',
  'urn:schemas-microsoft-com:office:excel': 'xvml',
  'urn:schemas-microsoft-com:office:powerpoint': 'pvml',
  'urn:schemas-microsoft-com:mac:vml': 'mv',
  'http://schemas.microsoft.com/office/2006/xmlPackage': 'pkg',
  'http://schemas.microsoft.com/office/2006/coverPageProps': 'cppr',
  'http://schemas.microsoft.com/office/2006/digsig': 'dssi',
  'http://schemas.microsoft.com/office/word/2003/auxHint': 'WX',
  'http://schemas.microsoft.com/office/word/2006/wordml': 'wne',
  'http://schemas.microsoft.com/office/word/2010/wordml': 'w14',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing': 'wp14',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas': 'wpc',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup': 'wpg',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingInk': 'wpi',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape': 'wps',
  'http://schemas.microsoft.com/office/word/2012/wordml': 'w15',
  'http://schemas.microsoft.com/office/word/2012/wordprocessingDrawing': 'wp15',
  'http://schemas.microsoft.com/office/word/2015/wordml/symex': 'w16se',
  'http://schemas.microsoft.com/office/word/2016/wordml/cid': 'w16cid',
  'http://schemas.microsoft.com/office/word/2018/wordml': 'w16',
  'http://schemas.microsoft.com/office/word/2018/wordml/cex': 'w16cex',
  'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash': 'w16sdtdh',
  'http://schemas.microsoft.com/office/word/2023/wordml/word16du': 'w16du',
  'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock': 'w16sdtfl',
  'http://schemas.microsoft.com/office/drawing/2007/8/2/chart': 'c14',
  'http://schemas.microsoft.com/office/drawing/2008/diagram': 'dsp',
  'http://schemas.microsoft.com/office/drawing/2010/main': 'a14',
  'http://schemas.microsoft.com/office/drawing/2010/picture': 'pic14',
  'http://schemas.microsoft.com/office/drawing/2010/chartDrawing': 'cdr14',
  'http://schemas.microsoft.com/office/drawing/2010/diagram': 'dgm14',
  'http://schemas.microsoft.com/office/drawing/2012/main': 'a15',
  'http://schemas.microsoft.com/office/drawing/2012/chart': 'c15',
  'http://schemas.microsoft.com/office/drawing/2012/chartStyle': 'cs',
  'http://schemas.microsoft.com/office/drawing/2013/main/command': 'a13cmd',
  'http://schemas.microsoft.com/office/drawing/2014/main': 'a16',
  'http://schemas.microsoft.com/office/drawing/2014/chart': 'c16',
  'http://schemas.microsoft.com/office/drawing/2014/chart/ac': 'c16ac',
  'http://schemas.microsoft.com/office/drawing/2014/chartex': 'cx',
  'http://schemas.microsoft.com/office/drawing/2016/11/main': 'a1611',
  'http://schemas.microsoft.com/office/drawing/2016/11/diagram': 'dgm1611',
  'http://schemas.microsoft.com/office/drawing/2016/12/diagram': 'dgm1612',
  'http://schemas.microsoft.com/office/drawing/2016/ink': 'ink16',
  'http://schemas.microsoft.com/office/drawing/2016/SVG/main': 'a16svg',
  'http://schemas.microsoft.com/office/drawing/2017/03/chart': 'c173',
  'http://schemas.microsoft.com/office/drawing/2017/decorative': 'adec',
  'http://schemas.microsoft.com/office/drawing/2017/model3d': 'am3d',
  'http://schemas.microsoft.com/office/drawing/2018/animation': 'an18',
  'http://schemas.microsoft.com/office/drawing/2018/animation/model3d': 'anam3d',
  'http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor': 'a18hc',
  'http://schemas.microsoft.com/office/thememl/2012/main': 'thm15',
  'http://schemas.microsoft.com/office/powerpoint/2010/main': 'p14',
  'http://schemas.microsoft.com/office/powerpoint/2012/main': 'p15',
  'http://schemas.microsoft.com/office/powerpoint/2013/main/command': 'p13cmd',
  'http://schemas.microsoft.com/office/powerpoint/2014/inkAction': 'iact',
  'http://schemas.microsoft.com/office/powerpoint/2015/main': 'p16',
  'http://schemas.microsoft.com/office/powerpoint/2015/09/main': 'p159',
  'http://schemas.microsoft.com/office/powerpoint/2015/10/main': 'p1510',
  'http://schemas.microsoft.com/office/powerpoint/2016/6/main': 'p166',
  'http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom': 'psez',
  'http://schemas.microsoft.com/office/powerpoint/2016/slidezoom': 'pslz',
  'http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom': 'psuz',
  'http://schemas.microsoft.com/office/powerpoint/2017/3/main': 'p173',
  'http://schemas.microsoft.com/office/powerpoint/2017/10/main': 'p1710',
  'http://schemas.microsoft.com/office/powerpoint/2018/4/main': 'p184',
  'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main': 'x14',
  'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac': 'x14ac',
  'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main': 'x15',
  'http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac': 'x15ac',
  'http://schemas.microsoft.com/office/spreadsheetml/2014/revision': 'xr',
  'http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing': 'xdr14',
  'http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11': 'wetp',
  'http://schemas.microsoft.com/office/webextensions/webextension/2010/11': 'we',
  'http://schemas.microsoft.com/ink/2010/main': 'msink',
  'http://schemas.microsoft.com/aml/2001/core': 'aml',
  'http://uri.etsi.org/01903/v1.3.2#': 'xd',
  'http://opendope.org/xpaths': 'odx',
  'http://opendope.org/conditions': 'odc',
  'http://opendope.org/components': 'odi',
  'http://opendope.org/questions': 'odq',
  'http://opendope.org/answers': 'oda',
  'http://opendope.org/SmartArt/DataHierarchy': 'odgm',
});
```

This is `NamespacePrefixMappings.getPreferredPrefixStatic` on docx4j `VERSION_17_1_1`, in the
same order, minus the entries docx4j itself has commented out; the prefixes docx4j marks
`// made up` are kept, since they are what docx4j writes. Namespaces of generated modules that
docx4j does not name (`org_docx4j_vml_root`, the encryption modules, MathML, InkML) are left to
the marshaller as today; add them here when a consumer needs them stable.

### 2.2 The default

```ts
export function getContext(options?: Jsonix.ContextOptions): Promise<Jsonix.Context> {
  ...
  new Jsonix.Context(mappings, { parentPointers: true, namespacePrefixes: NAMESPACE_PREFIXES, ...options })
```

A caller that passes its own `namespacePrefixes` replaces the table (spread semantics, no
merge); to extend it, spread `NAMESPACE_PREFIXES` into the option. The context is still built on
first use, so this is the one place where the default applies to every consumer.

### 2.3 Default namespaces

docx4j's static table has one default-namespace entry, SpreadsheetML (`<workbook xmlns="...">`;
`"s"` when JAXB requires a prefix), and maps package relationships to `rel`. The relationships
part gets `<Relationships xmlns="...">` from a second mapper, `NamespacePrefixMapperRelationshipsPart`,
that docx4j uses only when marshalling that part. The runtime forces the same shape: its output
seeds one prefix-to-namespace scope from the table, so two entries mapped to `''` collide
(verified 2026-09-10: the root's `xmlns` takes whichever entry comes last, and every child in the
other namespace redeclares `xmlns`; the earlier claim that the runtime would sort this out was
wrong). So `NAMESPACE_PREFIXES` follows docx4j's static table, and the facade's marshal functions
derive the table for each marshal from the root element (`namespacePrefixesFor`): a root in the
relationships namespace makes that namespace the default and gives SpreadsheetML `s`; every other
root uses the table as is. The derived table reaches the runtime through an object that inherits
from the shared context with its own `namespacePrefixes`; the context itself is not modified. A
consumer that calls `createMarshaller()` on the context directly gets the static table, hence
`rel:Relationships`, and every entry declared on the root (section 3).

## 3. What the runtime does with the table today, and the interim

`Jsonix.XML.Output` seeds its root namespace scope from `namespacePrefixes` and, when it writes
the document element, calls `declareNamespaces()` over that whole scope. So **every entry in the
table is declared on every root element**, used or not: about 110 `xmlns` attributes on a
`w:document`, on a `Relationships`, on a `w:styles`. It is well-formed and Word opens it (Word
itself declares some 40 namespaces on `w:document`), and it covers the `mc:Ignorable` requirement
for free, but it bloats every part and puts `xmlns:xml="http://www.w3.org/XML/1998/namespace"`
on every root, which the recommendation allows but no other producer writes.

Interim, in this package: `marshalString`, `marshalNode` and `marshalPackage` strip from the
root element every `xmlns` declaration that is neither used by an element or attribute in the
tree nor named in the root's `mc:Ignorable` attribute (the second rule is docx4j's
`McIgnorableNamespaceDeclarator`). One pass over the marshalled DOM
(`stripUnusedNamespaceDeclarations`); the same walk the runtime would need. A declaration below
the root shadows the root's for that subtree, so in a flat package each part's root keeps its own
declarations and `pkg:package` declares only `pkg`, as Word writes it. The `xml` prefix is never
declared. This gives the output docx4j gives. (The walk uses `childNodes`, not
`firstElementChild`: `@xmldom/xmldom` 0.9 does not implement the latter.)

Proper fix, in the runtime (filed as jsonix-CR-003, whose part 2 is deferred to 3.3.0; its part 1,
the `Jsonix.DOM` and `Context.namespacePrefixes` typings, shipped in 3.2.1 and removed this
package's casts): `namespacePrefixes` becomes a
preference only (declared where first used); a new marshaller option, say
`declareNamespaces: string[]`, lists the URIs to declare on the root regardless; and `xml` is
seeded as already declared. When that lands, the interim stripping becomes a no-op and can go.
Until then it is the behaviour consumers can rely on.

## 4. Tests

`test/smoke.mjs` gains:

- Marshal of the section 1 input reproduces the conventional prefixes, declares `w`, `w14` and
  `mc` on the root and nothing unused, and writes `xml:space` with no `xmlns:xml`.
- A `Relationships` root marshals with `xmlns="..."` and unprefixed children.
- The `mc:Ignorable` rule: a root with `mc:Ignorable="w14 w15"` and no `w15` content still
  declares `xmlns:w15`.
- A marshalled flat package declares only `pkg` on `pkg:package`.
- `getContext({ namespacePrefixes: {...} })` after `resetContext()` replaces the table.

All in `test/smoke.mjs` as of the implementation.

## 5. Consumers

- `docx4j-core-ts` raises its dependency to the release that carries this and drops the
  section 9 claim from its CR-001. Its `XmlPart` marshals through the facade and needs no
  namespace handling of its own.
- README: one paragraph under the facade, naming `NAMESPACE_PREFIXES` and how to extend it (done).
- The runtime's `Jsonix.DOM.serialize` is not in its typings; the facade casts to reach it.
  Adding `DOM.parse` / `DOM.serialize` to `types/main.d.ts` in `plutext/jsonix` would remove the cast.
