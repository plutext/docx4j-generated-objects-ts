// Runtime check of the facade against a small WordprocessingML document.
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';
import { unmarshalString, marshalString, unmarshalPackage, marshalPackage, unwrap, deepCopy, getContext, resetContext, NAMESPACE_PREFIXES, Jsonix } from '../dist/index.mjs';
import { createRequire } from 'node:module';
import { highlightHexValue, isCustomStyle } from '../dist/helpers/wml.mjs';

const xml = readFileSync(new URL('./fixtures/document.xml', import.meta.url), 'utf8');
const element = await unmarshalString(xml);
assert.equal(element.name.localPart, 'document');
const doc = unwrap(element);
assert.equal(doc.TYPE_NAME, 'org_docx4j_wml.Document');
assert.equal(doc.body.TYPE_NAME, 'org_docx4j_wml.Body');
assert.equal(doc.body.PARENT, doc, 'parent pointers are on by default');
const [p, tbl] = doc.body.content.map(unwrap);
assert.equal(p.TYPE_NAME, 'org_docx4j_wml.P');
assert.equal(tbl.TYPE_NAME, 'org_docx4j_wml.Tbl');
assert.equal(p.PARENT, doc.body);
const run = unwrap(p.content[0]);
assert.equal(run.TYPE_NAME, 'org_docx4j_wml.R');
assert.equal(unwrap(run.content[0]).value, 'Hello, generated-objects-ts');
assert.equal(highlightHexValue(run.rPr.highlight), '#FFFF00');
assert.equal(isCustomStyle({}), false);

const copy = deepCopy(p);
assert.notEqual(copy, p);
assert.equal(unwrap(copy.content[0]).PARENT, copy, 'children re-linked to the copy');
assert.equal(copy.PARENT, undefined, "a copy's own PARENT is unset");
assert.ok(Jsonix.Util.Type.isEqual(copy, p));

const out = await marshalString(element);
const again = unwrap(await unmarshalString(out));
assert.equal(again.body.content.length, doc.body.content.length);

// v:line marshals as id style from to (docx4j #469, compiler CR-007)
const ctx = await getContext();
const line = ctx.createMarshaller().marshalString({ name: { namespaceURI: 'urn:schemas-microsoft-com:vml', localPart: 'line' },
  value: { TYPE_NAME: 'org_docx4j_vml.CTLine', from: '0,0', to: '1,1', vmlId: 'l1', style: 'x' } }).replace(/ xmlns[^ >]*/g, '');
assert.match(line, /line id="l1" style="x" from="0,0" to="1,1"/);
// The UMD mappings stay CommonJS-loadable (no "type": "module" in package.json), also by package self-reference.
const require = createRequire(import.meta.url);
assert.equal(require('../modules/org_docx4j_wml.js').org_docx4j_wml.name, 'org_docx4j_wml');
assert.equal(require('@docx4j/generated-objects-ts/modules/org_docx4j_wml').org_docx4j_wml.name, 'org_docx4j_wml');
const { org_docx4j_wml } = await import('@docx4j/generated-objects-ts/modules/org_docx4j_wml');
assert.equal(org_docx4j_wml.name, 'org_docx4j_wml');
// Flat OPC package as Office JS getOoxml() returns it: the w:document inside pkg:xmlData is typed.
const flat = `<?xml version="1.0" standalone="yes"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>${xml.replace(/^<\?xml[^>]*>\s*/, '')}</pkg:xmlData></pkg:part></pkg:package>`;
const raw = await unmarshalString(flat);
assert.equal(raw.value.TYPE_NAME, 'org_docx4j_xmlPackage.Package');
assert.equal(typeof raw.value.part[1].xmlData.any.nodeType, 'number', 'the schema says skip: plain unmarshalling yields DOM');
const pkgElement = await unmarshalPackage(flat);
const docPart = pkgElement.value.part.find((p) => p.name === '/word/document.xml');
assert.equal(docPart.xmlData.any.value.TYPE_NAME, 'org_docx4j_wml.Document', 'unmarshalPackage types the known parts');
assert.equal(docPart.xmlData.any.value.body.content[0].value.TYPE_NAME, 'org_docx4j_wml.P');
const relsPart = pkgElement.value.part.find((p) => p.name === '/_rels/.rels');
assert.equal(relsPart.xmlData.any.value.TYPE_NAME, 'org_docx4j_relationships.Relationships');
unwrap(docPart.xmlData.any.value.body.content[0]).content[0].value.content[0].value.value = 'HELLO';
const packaged = await marshalPackage(pkgElement);
assert.match(packaged, /pkg:package/);
assert.match(packaged, /<w:t>HELLO<\/w:t>/);
assert.equal(docPart.xmlData.any.value.TYPE_NAME, 'org_docx4j_wml.Document', 'marshalPackage does not modify its input');
const roundTripped = await unmarshalPackage(packaged);
assert.equal(unwrap(roundTripped.value.part.find((p) => p.name === '/word/document.xml').xmlData.any).body.content.length, 2);

// CR-001: docx4j's prefixes; the root declares what the tree uses and what mc:Ignorable names, nothing else.
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const declarations = (xml) => Object.fromEntries([...xml.slice(0, xml.indexOf('>')).matchAll(/ xmlns(?::([^=]+))?="([^"]*)"/g)].map((m) => [m[1] ?? '', m[2]]));
const ignorable = await marshalString(await unmarshalString(
  `<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}" mc:Ignorable="w14 w15"><w:body><w:p w14:paraId="1"><w:r><w:t xml:space="preserve"> x</w:t></w:r></w:p></w:body></w:document>`,
));
assert.deepEqual(declarations(ignorable), { w: W, mc: MC, w14: W14, w15: W15 }, 'conventional prefixes; w15 declared because mc:Ignorable names it; nothing unused; no xmlns:xml');
assert.match(ignorable, /<w:p w14:paraId="1"><w:r><w:t xml:space="preserve"> x<\/w:t>/);
const relationships = await marshalString(await unmarshalString(`<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="http://x" Target="styles.xml"/></Relationships>`));
assert.deepEqual(declarations(relationships), { '': RELS }, 'a Relationships root uses the default namespace');
assert.match(relationships, /^<Relationships [^>]*><Relationship Id="rId1" Target="styles.xml" Type="http:\/\/x"\/><\/Relationships>$/, 'children unprefixed, nothing redeclared');
assert.deepEqual(declarations(packaged), { pkg: 'http://schemas.microsoft.com/office/2006/xmlPackage' }, 'a package root declares only pkg; each part declares its own');
resetContext();
await getContext({ namespacePrefixes: { ...NAMESPACE_PREFIXES, [W14]: 'wx14' } });
assert.match(await marshalString(await unmarshalString(`<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body><w:p w14:paraId="1"/></w:body></w:document>`)), /xmlns:wx14="[^"]*"[^>]*><w:body><w:p wx14:paraId="1"\/>/, 'a table passed to getContext replaces the default');
resetContext();

console.log('generated-objects-ts smoke: unmarshal, parent pointers, deepCopy, marshal, v:line order, flat OPC package round trip, namespace prefixes OK');

// Element factories (compiler CR-010): docx4j's ObjectFactory names, TYPE_NAME on wrapped literals,
// the same XML as a literal, and the public subpaths ./factory/* and ./el/*.
{
  const { createP, createR, createText, createRElement, createRT, createPElement, createSdtPrAlias, createSdtPrAliasElement } = await import('@docx4j/generated-objects-ts/factory/org_docx4j_wml');
  const el = await import('@docx4j/generated-objects-ts/el/org_docx4j_wml');
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const viaFactory = createPElement(createP({ content: [createRElement(createR({ content: [createRT(createText({ value: 'Hello' }))] }))] }));
  const viaEl = el.p({ content: [el.r({ content: [el.t({ value: 'Hello' })] })] });
  const literal = { name: { namespaceURI: W, localPart: 'p' }, value: { content: [{ name: { namespaceURI: W, localPart: 'r' }, value: { content: [{ name: { namespaceURI: W, localPart: 't' }, value: { value: 'Hello' } }] } }] } };
  const expected = await marshalString(literal);
  assert.equal(await marshalString(viaFactory), expected);
  assert.equal(await marshalString(viaEl), expected);
  assert.equal(viaFactory.value.TYPE_NAME, 'org_docx4j_wml.P');
  assert.equal(viaEl.value.TYPE_NAME, 'org_docx4j_wml.P', 'el.p sets TYPE_NAME on a literal');
  assert.equal(viaEl.value.content[0].value.content[0].value.TYPE_NAME, 'org_docx4j_wml.Text');
  assert.equal(el.sdt({}).value.TYPE_NAME, undefined, 'el.sdt has four types by scope and leaves TYPE_NAME alone');
  assert.equal(el.sdt({}).name.localPart, 'sdt');
  assert.equal(createSdtPrAliasElement(createSdtPrAlias({ val: 'x' })).value.TYPE_NAME, 'org_docx4j_wml.SdtPr.Alias');
  assert.equal(typeof el.object, 'function', 'a reserved-word element name is exported with `as`');
  assert.equal(el.r({}).name.namespaceURI, W, 'el.r is w:r; m:r (a foreign namespace declared in WML scopes) is not in el');
  console.log('factories: ObjectFactory names, TYPE_NAME and marshalling agree with literals');
}
