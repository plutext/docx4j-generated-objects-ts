// Runtime check of the facade against a small WordprocessingML document.
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';
import { unmarshalString, marshalString, unmarshalPackage, marshalPackage, unwrap, deepCopy, deepCopyAs, deepCopyAsSync, getContext, getContextSync, resetContext, NAMESPACE_PREFIXES, Jsonix } from '../dist/index.mjs';
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
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>${xml.replace(/^<\?xml[^>]*>\s*/, '')}</pkg:xmlData></pkg:part>
<pkg:part pkg:name="/word/commentsExtensible.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"><pkg:xmlData>
<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w16cex"><w16cex:commentExtensible w16cex:durableId="1A2B3C4D" w16cex:dateUtc="2026-09-16T10:30:00Z"/></w16cex:commentsExtensible></pkg:xmlData></pkg:part></pkg:package>`;
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
// The w16cex module (docx4j CR-018): word/commentsExtensible.xml is typed like any known root element.
const cexPart = pkgElement.value.part.find((p) => p.name === '/word/commentsExtensible.xml');
assert.equal(cexPart.xmlData.any.value.TYPE_NAME, 'org_docx4j_w16cex.CTCommentsExtensible', 'unmarshalPackage types a w16cex part');
assert.equal(cexPart.xmlData.any.value.commentExtensible[0].durableId, '1A2B3C4D');
assert.equal(cexPart.xmlData.any.value.commentExtensible[0].dateUtc.year, 2026, 'dateUtc is a calendar');
assert.equal(cexPart.xmlData.any.value.ignorable, 'w16cex', 'mc:Ignorable on the w16cex root (CR-018)');

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

// CR-001 section 7: every prefix mc:Ignorable names is declared on the root (docx4j's
// McIgnorableNamespaceDeclarator). Word and Excel repair a file where one is not; found by an
// @docx4j/core-ts Excel run over xl/workbook.xml, whose xr2, xr6 and xr10 nothing in the tree uses.
const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const declarationsOf = (xml) => {
  const root = xml.slice(0, xml.indexOf('>') + 1);
  const declared = [...root.matchAll(/xmlns:([a-z0-9]+)=/g)].map((m) => m[1]);
  const ignorable = (root.match(/mc:Ignorable="([^"]*)"/) || [, ''])[1].split(/\s+/).filter(Boolean);
  return { declared, ignorable, undeclared: ignorable.filter((p) => !declared.includes(p)) };
};
const workbook = `<workbook xmlns="${SML}" xmlns:mc="${MC}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" mc:Ignorable="x15 xr xr6 xr10 xr2"><sheets><sheet name="S1" sheetId="1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></sheets></workbook>`;
const workbookOut = declarationsOf(await marshalString(await unmarshalString(workbook)));
assert.deepEqual(workbookOut.undeclared, [], 'every mc:Ignorable prefix is declared (xr2, xr6, xr10 are not used by the tree)');
assert.ok(['xr2', 'xr6', 'xr10'].every((p) => workbookOut.declared.includes(p)), 'the revision prefixes are declared from NAMESPACE_PREFIXES');
const wordOut = declarationsOf(await marshalString(await unmarshalString(
  `<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}" mc:Ignorable="w14 w15"><w:body><w:p/></w:body></w:document>`)));
assert.deepEqual(wordOut.undeclared, [], 'w15 is declared although nothing uses it');
assert.deepEqual(wordOut.ignorable, ['w14', 'w15']);
const unknownOut = declarationsOf(await marshalString(await unmarshalString(
  `<w:document xmlns:w="${W}" xmlns:mc="${MC}" mc:Ignorable="w14 zz"><w:body><w:p/></w:body></w:document>`)));
assert.deepEqual(unknownOut.ignorable, ['w14'], 'a prefix no table entry resolves is dropped from mc:Ignorable');
assert.deepEqual(unknownOut.undeclared, []);

// CR-002: builders/wml. Fragments wrapped in the container docx4j would use, text sugar over el,
// the run mapping shared with core-ts's Font view, and traversal.
{
  const { wml, wmlOne, p, r, t, br, tab, tbl, tr, tc, inlinePicture, rPrToElements, rPrFromElements, walkAll, mcBranchOf, sdt, W15_NAMESPACE, sdtPr, nextSdtId, sdtProperty, sdtKindOf, textOf, runItemsOf, walk, find, linkParents, applyRunOptions, readRunOptions } = await import('@docx4j/generated-objects-ts/builders/wml');
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
  // CR-003 section 3.4: run lists live under three property names; w:moveTo is the document's text, w:moveFrom is not.
  const [moved] = await wml('<w:p><w:r><w:t xml:space="preserve">kept </w:t></w:r><w:moveTo w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>moved</w:t></w:r></w:moveTo><w:moveFrom w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>gone</w:t></w:r></w:moveFrom><w:ins w:id="3" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve"> inserted</w:t></w:r></w:ins></w:p>');
  assert.equal(textOf(moved), 'kept moved inserted', 'textOf reads w:moveTo (accOrBarOrBox) and skips w:moveFrom');
  const [moveTo, moveFrom, ins] = moved.value.content.slice(1);
  assert.equal(runItemsOf(moveTo.value).length, 1, 'runItemsOf: w:moveTo keeps its runs under accOrBarOrBox');
  assert.equal(runItemsOf(moveFrom.value).length, 1, 'runItemsOf: w:moveFrom too');
  assert.equal(runItemsOf(ins.value).length, 1, 'runItemsOf: w:ins keeps its runs under customXmlOrSmartTagOrSdt');
  assert.equal(runItemsOf(moved.value)[0].name.localPart, 'r', 'runItemsOf: a paragraph uses content');
  const [sdtRun] = await wml('<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent><w:r><w:t>in a control</w:t></w:r></w:sdtContent></w:sdt>', { wrapper: 'p' });
  assert.equal(runItemsOf(sdtRun.value).length, 1, 'runItemsOf: a run-level content control reads through sdtContent');
  assert.equal(runItemsOf({}), undefined, 'runItemsOf: no run list');
  // CR-003 section 3.1: content controls in the four forms, their kind elements, and the readers.
  const blockSdt = sdt([p('inside')], { kind: 'PlainText', tag: 'a-tag', title: 'A title', id: 7 });
  assert.equal(blockSdt.value.TYPE_NAME, 'org_docx4j_wml.SdtBlock', 'paragraph content gives the block form');
  assert.equal(blockSdt.value.sdtContent.TYPE_NAME, 'org_docx4j_wml.SdtContentBlock');
  assert.equal(await marshalString(blockSdt),
    `<w:sdt xmlns:w="${W}"><w:sdtPr><w:alias w:val="A title"/><w:tag w:val="a-tag"/><w:id w:val="7"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:sdtContent></w:sdt>`,
    'w:sdtPr writes alias, tag, id then the kind element');
  assert.equal(sdt([r('x')], { id: 1 }).value.TYPE_NAME, 'org_docx4j_wml.SdtRun', 'a run gives the run form');
  const [aRow] = await wml('<w:tr><w:tc><w:p/></w:tc></w:tr>');
  assert.equal(sdt([aRow], { id: 1 }).value.TYPE_NAME, 'org_docx4j_wml.CTSdtRow', 'a w:tr gives the row form');
  const [aCell] = await wml('<w:tc><w:p/></w:tc>');
  assert.equal(sdt([aCell], { id: 1 }).value.TYPE_NAME, 'org_docx4j_wml.CTSdtCell', 'a w:tc gives the cell form');
  assert.equal(sdt([r('x')], { id: 1, form: 'block' }).value.TYPE_NAME, 'org_docx4j_wml.SdtBlock', 'form overrides the inference');
  assert.throws(() => sdt([r('x')], { kind: 'RepeatingSection' }), /block-level control/);
  const checkbox = sdt([r('x')], { kind: 'CheckBox', id: 2 });
  assert.equal(sdtProperty(checkbox.value.sdtPr, 'checkbox', W14).value.checked.val, false, 'w14:checked is a boolean since 0.1.3');
  assert.equal(sdtKindOf(checkbox.value.sdtPr), 'CheckBox');
  assert.equal(sdtKindOf(sdt([p('x')], { id: 3 }).value.sdtPr), 'RichText', 'an untyped control is rich text');
  for (const kind of ['PlainText', 'Picture', 'BuildingBlockGallery', 'ComboBox', 'DropDownList', 'DatePicker', 'RepeatingSection', 'RepeatingSectionItem', 'Group', 'Citation', 'Bibliography', 'Equation']) {
    assert.equal(sdtKindOf(sdtPr({ kind })), kind, `sdtKindOf round trip: ${kind}`);
  }
  assert.equal(sdtProperty(blockSdt.value.sdtPr, 'tag').value.val, 'a-tag');
  assert.equal(sdtProperty(blockSdt.value.sdtPr, 'nope'), undefined);
  // Word writes a repeating section's binding as w15:dataBinding, so sdtProperty takes a list or '*'.
  const [boundSdt] = await wml('<w:sdt><w:sdtPr><w:id w:val="9"/><w15:dataBinding w:xpath="/root[1]/a[1]" w:storeItemID="{GUID}"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>');
  const boundPr = boundSdt.value.sdtPr;
  assert.equal(sdtProperty(boundPr, 'dataBinding'), undefined, 'the wml default does not see the w15 twin');
  assert.equal(sdtProperty(boundPr, 'dataBinding', W15_NAMESPACE)?.value.xpath, '/root[1]/a[1]');
  assert.equal(sdtProperty(boundPr, 'dataBinding', [W, W15_NAMESPACE])?.value.storeItemID, '{GUID}', 'an array looks in both');
  assert.equal(sdtProperty(boundPr, 'dataBinding', '*')?.value.xpath, '/root[1]/a[1]', "'*' is any namespace");
  assert.ok(nextSdtId(blockSdt) !== 7 && Number.isInteger(nextSdtId(blockSdt)), 'nextSdtId avoids the ids in the tree');
  assert.equal(textOf(blockSdt), 'inside', 'textOf reads a control built by sdt');

  // CR-003 sections 3.2 and 3.3: rows and cells (tbl rebuilt over them), and the inline picture
  // docx4j's createImageInline writes. Attributes are emitted sorted by name (see "What the
  // declarations promise"), and a standalone w:tc / w:tbl / w:drawing root carries xsi:type.
  assert.equal(await marshalString(tc('a', { width: 1000 })),
    `<w:tc xmlns:w="${W}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="w:CT_Tc"><w:tcPr><w:tcW w:type="dxa" w:w="1000"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>`);
  assert.equal(await marshalString(tc([])),
    `<w:tc xmlns:w="${W}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="w:CT_Tc"><w:p/></w:tc>`,
    'an empty cell still gets the w:p Word requires, and nothing else');
  assert.equal(await marshalString(tr(['a', 'b'], { widths: [100, 200] })),
    `<w:tr xmlns:w="${W}"><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="100"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="200"/></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>`);
  assert.equal(textOf(tr([tc('x'), 'y'])), 'x\ny', 'tr takes built cells and strings');
  assert.match(await marshalString(tbl([['Item', 'Qty'], ['Widget', '3']], { style: 'TableGrid' })),
    /<w:tblPr><w:tblStyle w:val="TableGrid"\/><w:tblW w:type="dxa" w:w="9026"\/><\/w:tblPr><w:tblGrid><w:gridCol w:w="4513"\/><w:gridCol w:w="4513"\/><\/w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4513"\/>/,
    'tbl over tr/tc writes the grid it always did');

  const picture = inlinePicture('rId9', { cx: 914400, cy: 457200, id: 3, name: 'image1.png', descr: 'alt', title: 'Title' });
  const pictureXml = await marshalString(picture);
  assert.match(pictureXml, /<wp:inline distB="0" distL="0" distR="0" distT="0"><wp:extent cx="914400" cy="457200"\/>/);
  assert.match(pictureXml, /<wp:docPr descr="alt" id="3" name="image1.png" title="Title"\/>/, 'docPr carries descr and the title docx4j CR-018 added');
  assert.match(pictureXml, /<a:graphicFrameLocks noChangeAspect="true"\/>/);
  assert.match(pictureXml, /<a:graphicData uri="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/picture"><pic:pic>/);
  assert.match(pictureXml, /<a:blip r:embed="rId9"\/><a:stretch><a:fillRect\/><\/a:stretch>/);
  assert.match(pictureXml, /<a:prstGeom prst="rect"><a:avLst\/><\/a:prstGeom>/);
  assert.equal(find(picture, 'org_docx4j_dml_picture.Pic').length, 1, 'the pic is reachable by TYPE_NAME');
  assert.equal(picture.value.anchorOrInline[0].docPr.PARENT, picture.value.anchorOrInline[0], 'inlinePicture links PARENT');

  // CR-003 section 2 and 3.5: deepCopyAs for w:pPrChange, and w:rPr as the EG_RPrBase element list.
  const [styled] = await wml('<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr></w:p>');
  const asBase = await deepCopyAs(styled.value.pPr, 'org_docx4j_wml.PPrBase');
  assert.equal(asBase.TYPE_NAME, 'org_docx4j_wml.PPrBase');
  assert.deepEqual(Object.keys(asBase).filter((k) => k !== 'TYPE_NAME'), ['pStyle', 'jc'], 'properties PPrBase does not declare are dropped');
  assert.ok(styled.value.pPr.rPr, 'the original is untouched');
  styled.value.pPr.pPrChange = { id: 1, author: 'A', pPr: asBase };
  const changed = await marshalString(styled);
  assert.match(changed, /<w:pPrChange w:id="1" w:author="A"><w:pPr><w:pStyle w:val="Heading1"\/><w:jc w:val="center"\/><\/w:pPr><\/w:pPrChange>/, 'no xsi:type, as Word writes it');
  await assert.rejects(deepCopyAs(styled.value.pPr, 'org_docx4j_wml.RPr'), /not org_docx4j_wml.PPr or one of its base types/);
  // core-ts records a w:pPrChange inside synchronous Office JS-shaped setters, so both forms exist.
  const syncCopy = deepCopyAsSync(styled.value.pPr, 'org_docx4j_wml.PPrBase');
  assert.deepEqual(Object.keys(syncCopy), Object.keys(asBase), 'deepCopyAsSync agrees with the async form');
  assert.throws(() => deepCopyAsSync(styled.value.pPr, 'org_docx4j_wml.RPr'), /not org_docx4j_wml.PPr or one of its base types/);
  assert.ok(getContextSync(), 'the context is built by now');

  const [rich2] = await wml('<w:r><w:rPr><w:rStyle w:val="Strong"/><w:b/><w:color w:val="FF0000"/><w:shadow/><w:u w:val="single"/><w14:glow w14:rad="1"><w14:srgbClr w14:val="FF0000"/></w14:glow><w14:shadow w14:blurRad="1"><w14:srgbClr w14:val="00FF00"/></w14:shadow><w14:ligatures w14:val="standard"/></w:rPr></w:r>', { wrapper: 'p' });
  const asElements = rPrToElements(rich2.value.rPr);
  assert.deepEqual(asElements.map((e) => (e.name.namespaceURI === W14 ? 'w14:' : '') + e.name.localPart),
    ['rStyle', 'b', 'shadow', 'color', 'u', 'w14:glow', 'w14:shadow', 'w14:ligatures'],
    'EG_RPrBase order, wml shadow before color, the w14 effects last');
  const backToRPr = rPrFromElements(asElements);
  assert.deepEqual(Object.keys(backToRPr).filter((k) => k !== 'TYPE_NAME' && k !== 'PARENT').sort(),
    ['b', 'color', 'glow', 'ligatures', 'rStyle', 'shadow', 'shadow14', 'u'], 'w14:shadow comes back as shadow14, not shadow');
  assert.equal(backToRPr.shadow14.blurRad, 1, 'the w14 effect keeps its value');
  assert.equal(rPrToElements(undefined).length, 0);
  assert.equal(rPrFromElements(undefined).TYPE_NAME, 'org_docx4j_wml.RPr');

  // CR-003 section 3.6: walkAll enters the DOM an xs:any property holds.
  const domSeenAll = [];
  walkAll(pkgElement, () => {}, (node) => { if (node.getAttribute && node.getAttribute('w:val')) domSeenAll.push(node.nodeName); });
  const typedSeen = [];
  walkAll(pkgElement, (v) => { typedSeen.push(v); }, () => {});
  assert.ok(typedSeen.length > 0, 'walkAll visits typed objects like walk');

  // CR-003 section 3.8: the mc:AlternateContent branch a reader takes (docx4j's McSelection).
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  const alt = (branches) => `<w:p xmlns:w="${W}" xmlns:mc="${MC}" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:zz="urn:not-known"><w:r><w:t xml:space="preserve">a </w:t></w:r><mc:AlternateContent>${branches}</mc:AlternateContent></w:p>`;
  const paragraphOf = async (xml) => unwrap(await unmarshalString(xml));
  const choiceWps = '<mc:Choice Requires="wps"><w:r><w:t>wps</w:t></w:r></mc:Choice>';
  const choiceUnknown = '<mc:Choice Requires="zz"><w:r><w:t>unknown</w:t></w:r></mc:Choice>';
  const fallback = '<mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback>';
  assert.equal(textOf(await paragraphOf(alt(choiceUnknown + choiceWps + fallback))), 'a wps', 'rule 1: the first understood Choice');
  assert.equal(textOf(await paragraphOf(alt(choiceUnknown + fallback))), 'a fallback', 'rule 2: the Fallback when no Choice is understood');
  assert.equal(textOf(await paragraphOf(alt(choiceUnknown))), 'a unknown', 'rule 3: Choices with no Fallback take the first, which docx4j TextUtils cannot');
  assert.equal(textOf(await paragraphOf(alt(''))), 'a ', 'rule 4: no branches, no text');
  const withBoth = await paragraphOf(alt(choiceUnknown + choiceWps + fallback));
  const ac = withBoth.content[1].value;
  assert.equal(mcBranchOf(ac).length, 1);
  assert.equal(textOf(mcBranchOf(ac)[0]), 'wps', 'mcBranchOf returns the branch items');
  assert.equal(textOf(mcBranchOf(ac, { understood: ['zz'] })[0]), 'unknown', 'a caller can say what it understands');
  assert.equal(textOf(mcBranchOf(ac, { understood: [] })[0]), 'fallback', 'nothing understood: the Fallback');
  assert.equal(mcBranchOf(undefined), undefined);

  console.log('builders/wml: fragments (content controls by first decisive descendant), tagged form, text sugar, run mapping, traversal OK');
}
