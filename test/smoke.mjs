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

// CR-002: builders/wml. Fragments wrapped in the container docx4j would use, text sugar over el,
// the run mapping shared with core-ts's Font view, and traversal.
{
  const { wml, wmlOne, p, r, t, br, tab, tbl, textOf, walk, find, linkParents, applyRunOptions, readRunOptions } = await import('@docx4j/generated-objects-ts/builders/wml');
  const el = await import('@docx4j/generated-objects-ts/el/org_docx4j_wml');
  const types = (elements) => elements.map((e) => e.value.TYPE_NAME);
  const names = (elements) => elements.map((e) => e.name.localPart);
  // Each wrapper row: the fragment's elements come back typed, in order, PARENT linked below the top level.
  const blocks = await wml('<w:p><w:r><w:t>a</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>');
  assert.deepEqual(types(blocks), ['org_docx4j_wml.P', 'org_docx4j_wml.Tbl']);
  assert.deepEqual(names(blocks), ['p', 'tbl']);
  assert.equal(blocks[0].value.PARENT.TYPE_NAME, 'org_docx4j_wml.Body', 'the top-level PARENT is the throw-away wrapper');
  assert.equal(blocks[0].value.content[0].value.PARENT, blocks[0].value, 'PARENT linked below the top level');
  assert.deepEqual(types(await wml('<w:r><w:t>a</w:t></w:r><w:hyperlink r:id="rId1"><w:r><w:t>b</w:t></w:r></w:hyperlink>')), ['org_docx4j_wml.R', 'org_docx4j_wml.P.Hyperlink']);
  assert.deepEqual(types(await wml('<w:tr><w:tc><w:p/></w:tc></w:tr>')), ['org_docx4j_wml.Tr']);
  assert.deepEqual(types(await wml('<w:tc><w:p/></w:tc>')), ['org_docx4j_wml.Tc']);
  assert.deepEqual(types(await wml('<w:t>a</w:t><w:br/><w:tab/>')), ['org_docx4j_wml.Text', 'org_docx4j_wml.Br', 'org_docx4j_wml.R.Tab']);
  assert.deepEqual(types(await wml('<w:document><w:body><w:p/></w:body></w:document>')), ['org_docx4j_wml.Document'], 'a global element is unmarshalled as it is');
  assert.deepEqual(types(await wml('<w:bookmarkStart w:id="0" w:name="x"/>')), ['org_docx4j_wml.CTBookmark'], 'the same type at both levels: w:body');
  const [withId] = await wml('<w:p w14:paraId="1A2B3C4D"><w:r><w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">x</w:t></w:r></w:p>');
  assert.equal(withId.value.paraId, '1A2B3C4D', 'the standard prefixes are declared for the fragment');
  assert.equal(withId.value.content[0].value.content[0].value.TYPE_NAME, 'org_docx4j_wml.Text', 'a declaration inside the fragment wins');
  await assert.rejects(wmlOne('<w:p/><w:p/>'), /Expected one element, found 2/);
  await assert.rejects(wml('<w:p><w:r></w:p>'), /mismatch/, 'malformed XML throws the parser\'s error');
  await assert.rejects(wml('<w:bogus/>'), /not known in this context/, 'an unknown element throws the runtime\'s error');
  // The content control inference: the wrapper decides the type there (a run-level w:sdt unmarshals under w:body, as SdtBlock).
  assert.deepEqual(types(await wml('<w:sdt><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>')), ['org_docx4j_wml.SdtRun']);
  assert.deepEqual(types(await wml('<w:sdt><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>', { wrapper: 'body' })), ['org_docx4j_wml.SdtBlock'], 'wrapper overrides the inference');
  const [nestedBlock] = await wml('<w:sdt><w:sdtContent><w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt></w:sdtContent></w:sdt>');
  assert.deepEqual([nestedBlock.value.TYPE_NAME, nestedBlock.value.sdtContent.content[0].value.TYPE_NAME], ['org_docx4j_wml.SdtBlock', 'org_docx4j_wml.SdtBlock'], 'nested: decided by the first decisive descendant');
  const [nestedRun] = await wml('<w:sdt><w:sdtContent><w:sdt><w:sdtContent><w:r/></w:sdtContent></w:sdt></w:sdtContent></w:sdt>');
  assert.deepEqual([nestedRun.value.TYPE_NAME, nestedRun.value.sdtContent.content[0].value.TYPE_NAME], ['org_docx4j_wml.SdtRun', 'org_docx4j_wml.SdtRun']);
  assert.deepEqual(types(await wml('<w:sdt><w:sdtContent><w:bookmarkStart w:id="0" w:name="x"/><w:r/></w:sdtContent></w:sdt>')), ['org_docx4j_wml.SdtRun'], 'a bookmark before the run is not decisive');
  assert.deepEqual(types(await wml('<w:sdt><w:sdtContent><w:tr><w:tc><w:p/></w:tc></w:tr></w:sdtContent></w:sdt>')), ['org_docx4j_wml.CTSdtRow'], 'a row control');
  assert.deepEqual(types(await wml('<w:sdt><w:sdtContent/></w:sdt>')), ['org_docx4j_wml.SdtBlock'], 'nothing decisive: body');
  assert.deepEqual(types(await wml('<w:sdt><w:sdtPr><w:alias w:val="x"/></w:sdtPr></w:sdt>')), ['org_docx4j_wml.SdtBlock']);
  assert.deepEqual(types(await wml('<w:customXml w:element="x"><w:r/></w:customXml>')), ['org_docx4j_wml.CTCustomXmlRun']);
  // The tagged form: a typed value marshals in place, text is escaped, wml.raw is verbatim.
  const [tagged] = await wml`<w:p>${r('bold', { bold: true })}${'<&>'}${wml.raw('<w:r><w:t>raw</w:t></w:r>')}</w:p>`;
  assert.equal(await marshalString(tagged), `<w:p xmlns:w="${W}"><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t>raw</w:t></w:r></w:p>`);
  assert.equal((await wmlOne`<w:p>${[r('a'), r('b')]}</w:p>`).value.content.length, 2, 'an array interpolates item by item');
  let hooked;
  await wml('<w:p/>', { preprocess: (doc) => { hooked = doc.documentElement.localName; } });
  assert.equal(hooked, 'body', 'preprocess sees the wrapped DOM before unmarshalling');
  // Text sugar.
  assert.deepEqual([t(' a').value.space, t('a ').value.space, t('a  b').value.space, t('a b').value.space], ['preserve', 'preserve', 'preserve', undefined]);
  assert.deepEqual(names(r('a\tb\nc').value.content), ['t', 'tab', 't', 'br', 't']);
  assert.equal(br('page').value.type, 'page');
  assert.equal(br().value.type, undefined);
  assert.equal(tab().value.TYPE_NAME, 'org_docx4j_wml.R.Tab');
  assert.equal(await marshalString(p('Hi', { style: 'Heading1', bold: true, runStyle: 'Strong' })),
    `<w:p xmlns:w="${W}"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Strong"/><w:b/><w:bCs/></w:rPr><w:t>Hi</w:t></w:r></w:p>`);
  assert.equal(p([r('a'), r('b')]).value.content.length, 2);
  assert.equal(p('x', { bold: false }).value.content[0].value.rPr, undefined, 'no empty w:rPr');
  const grid = tbl([['a', 'b'], ['c']], { widths: [1000, 2000], style: 'TableGrid' });
  assert.deepEqual(grid.value.tblGrid.gridCol.map((g) => g.w), [1000, 2000]);
  assert.equal(grid.value.tblPr.tblStyle.val, 'TableGrid');
  assert.equal(grid.value.content[1].value.content.length, 2, 'a short row is padded with empty cells');
  assert.deepEqual(tbl([['a', 'b', 'c']]).value.tblGrid.gridCol.map((g) => g.w), [3008, 3008, 3010], 'equal columns sum to the table width');
  assert.doesNotMatch(await marshalString(el.body({ content: [tbl([['x']])] })), /xsi:type/);
  const [rich] = await wml('<w:p><w:r><w:t xml:space="preserve">a </w:t><w:tab/></w:r><w:hyperlink><w:r><w:t>b</w:t></w:r></w:hyperlink><w:del><w:r><w:delText>X</w:delText></w:r></w:del><w:ins><w:r><w:t>c</w:t></w:r></w:ins><w:sdt><w:sdtContent><w:r><w:br/><w:t>d</w:t><w:sym w:font="Wingdings" w:char="F0FC"/></w:r></w:sdtContent></w:sdt></w:p>');
  assert.equal(textOf(rich), 'a \tbc\nd', 'textOf: t, tab, br, sym; hyperlink, ins and sdt read; del skipped');
  assert.equal(textOf(rich.value.content[0]), 'a \t', 'textOf a run');
  assert.equal(textOf(doc.body), 'Hello, generated-objects-ts\ncell ', 'textOf a body: paragraphs and table cells joined with \\n');
  assert.equal(textOf(element), textOf(doc.body), 'textOf a document element');
  assert.equal(textOf(tbl([['a', 'b'], ['c', 'd']])), 'a\nb\nc\nd');
  // The run mapping (moved from core-ts): the round trip, removals, the throw, the defaults, and r() using it.
  const full = { bold: true, italic: true, underline: true, strikeThrough: true, doubleStrikeThrough: true, subscript: true, superscript: false, name: 'Arial', size: 11.5, color: '#ff0000', highlightColor: '#FFFF00', style: 'Strong' };
  const rPr = applyRunOptions({}, full);
  assert.deepEqual(rPr, { b: {}, bCs: {}, i: {}, iCs: {}, u: { val: 'single' }, strike: {}, dstrike: {}, vertAlign: { val: 'subscript' }, rFonts: { ascii: 'Arial', hAnsi: 'Arial' }, sz: { val: 23 }, szCs: { val: 23 }, color: { val: 'FF0000' }, highlight: { val: 'yellow' }, rStyle: { val: 'Strong' } });
  assert.deepEqual(readRunOptions(rPr), { ...full, underline: 'Single', color: '#FF0000' });
  assert.equal(readRunOptions(applyRunOptions({}, { underline: 'Double' })).underline, 'Double');
  assert.equal(readRunOptions(applyRunOptions({}, { underline: false })).underline, 'None');
  assert.equal(applyRunOptions({}, { color: 'auto' }).color.val, 'auto');
  assert.deepEqual(applyRunOptions(applyRunOptions({}, { bold: true, underline: 'Double', name: 'A', color: 'auto', highlightColor: 'yellow', style: 'S' }), { bold: false, underline: 'None', name: '', color: '', highlightColor: null, style: '' }), {}, 'each removal value deletes its property');
  assert.throws(() => applyRunOptions({}, { highlightColor: '#123456' }), /Not a highlight colour/);
  assert.deepEqual(readRunOptions(undefined), { bold: false, italic: false, underline: 'None', strikeThrough: false, doubleStrikeThrough: false, subscript: false, superscript: false, name: '', size: 0, color: '', highlightColor: null, style: '' });
  assert.deepEqual({ ...r('x', { bold: true, size: 12 }).value.rPr, TYPE_NAME: undefined }, { ...applyRunOptions({}, { bold: true, size: 12 }), TYPE_NAME: undefined }, 'r() writes w:rPr through applyRunOptions');
  // Traversal.
  const [para] = await wml('<w:p><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>');
  const seen = [];
  walk(para, (v, parent, key) => { seen.push(`${v.TYPE_NAME}@${key}`); if (v.TYPE_NAME === 'org_docx4j_wml.R' && key === 0) return false; });
  assert.deepEqual(seen, ['org_docx4j_wml.P@', 'org_docx4j_wml.R@0', 'org_docx4j_wml.R@1', 'org_docx4j_wml.Text@0'], 'walk: typed objects with their key; false skips the subtree');
  assert.equal(find(para, 'org_docx4j_wml.Text').length, 2);
  const built = p('x', { bold: true });
  assert.equal(built.value.content[0].value.PARENT, undefined, 'el builders do not link parents');
  linkParents(built, undefined);
  const builtRun = built.value.content[0].value;
  assert.equal(builtRun.PARENT, built.value);
  assert.equal(builtRun.rPr.PARENT, builtRun);
  assert.deepEqual(Object.getOwnPropertyDescriptor(builtRun, 'PARENT'), { value: built.value, writable: true, enumerable: false, configurable: true }, 'PARENT as the unmarshaller defines it');
  assert.equal(find(built, 'org_docx4j_wml.Text').length, 1, 'find from the parent after linking');
  const domSeen = [];
  walk(pkgElement, (v) => { domSeen.push(v.nodeType === undefined ? 'typed' : 'DOM'); });
  assert.ok(domSeen.length > 0 && !domSeen.includes('DOM'), 'walk does not enter DOM nodes');
  console.log('builders/wml: fragments (content controls by first decisive descendant), tagged form, text sugar, run mapping, traversal OK');
}
