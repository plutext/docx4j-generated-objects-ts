// Runtime check of the facade against a small WordprocessingML document.
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';
import { unmarshalString, marshalString, unmarshalPackage, marshalPackage, unwrap, deepCopy, deepCopyAs, deepCopyAsSync, getContext, getContextSync, resetContext, NAMESPACE_PREFIXES, IGNORABLE_PREFIX_ALIASES, IGNORABLE_EXTENSION_NAMESPACES, MODULE_NAMES, MODULES, modulesFor, WML_GLOBAL_ELEMENTS, Jsonix } from '../dist/index.mjs';
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
// ECMA-376 Transitional ST_OnOff admits on/off beside true/false/1/0 (pandoc writes w:val="off"); the runtime reads them since
// @docx4j/jsonix 3.4.1 (jsonix-CR-007), the floor of 0.3.2, and writes them back as false/true, the forms ECMA-376 admits and Office itself writes.
const onOff = await unmarshalString(`<w:style xmlns:w="${W}" w:type="paragraph" w:styleId="SourceCode"><w:pPr><w:keepNext w:val="on"/><w:wordWrap w:val="off"/></w:pPr></w:style>`);
assert.equal(onOff.value.pPr.wordWrap.val, false, 'w:val="off" is false');
assert.equal(onOff.value.pPr.keepNext.val, true, 'w:val="on" is true');
assert.match(await marshalString(onOff), /<w:keepNext w:val="true"\/><w:wordWrap w:val="false"\/>/, 'written back as true and false');
resetContext();
await getContext({ namespacePrefixes: { ...NAMESPACE_PREFIXES, [W14]: 'wx14' } });
assert.match(await marshalString(await unmarshalString(`<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body><w:p w14:paraId="1"/></w:body></w:document>`)), /xmlns:wx14="[^"]*"[^>]*><w:body><w:p wx14:paraId="1"\/>/, 'a table passed to getContext replaces the default');
resetContext();

// CR-005: the registry is the list of modules, and `modules` builds a context over fewer of them.
assert.equal(MODULE_NAMES.length, Object.keys(MODULES).length, 'MODULE_NAMES is the registry keys');
assert.deepEqual([...MODULE_NAMES].sort(), Object.keys(MODULES).sort());
assert.equal(MODULES.org_docx4j_wml.name, 'org_docx4j_wml', 'the registry holds the mappings themselves');
// The option is read when the context is BUILT, so a caller wanting another one resets first; a
// caller that forgets gets the context that exists, as with every other option.
// The mappings reference each other, so a hand-picked list fails at context build; modulesFor
// takes the closure from each mapping's own `dependencies`.
const wmlOnly = modulesFor('org_docx4j_wml');
assert.ok(wmlOnly.length > 1 && wmlOnly.length < MODULE_NAMES.length, `a closure, not everything: ${wmlOnly.length} of ${MODULE_NAMES.length}`);
assert.ok(wmlOnly.some((m) => m.name === 'org_docx4j_com_microsoft_schemas_office_drawing_x2010_main'),
  'the closure reaches a14, which a hand-picked list of wml and its neighbours misses');
assert.throws(() => modulesFor('org_docx4j_not_a_module'), /not a module of this package/);
const full = await getContext();
assert.equal(await getContext({ modules: wmlOnly }), full, 'options are read when the context is built, not after');
resetContext();
const small = await getContext({ modules: wmlOnly });
assert.notEqual(small, full);
assert.match(await marshalString(await unmarshalString(`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>`)),
  /<w:t>x<\/w:t>/, 'a context over a subset of the modules marshals a w:p');
await assert.rejects(unmarshalString(`<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="http://x" Target="s.xml"/></Relationships>`)
  .then(() => {}), /not known in this context|Relationships/, 'and knows nothing of the modules it was not given');
resetContext();
const byLoader = await getContext({ modules: () => Promise.resolve(wmlOnly) });
assert.notEqual(byLoader, full, 'modules may be a function, for a caller that loads them itself');
resetContext();

// CR-007: per-call options reach the runtime, and nothing is wired by default.
{
  const W_ = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const SML_ = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const XR_ = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision';

  // An attribute no type declares is dropped; the callback names it where it was dropped.
  const dropped = [];
  await unmarshalString(`<w:document xmlns:w="${W_}" xmlns:zz="urn:zz"><w:body><w:p zz:mine="1"/></w:body></w:document>`,
    { onUnexpectedAttribute: (name) => dropped.push(`${name.namespaceURI}}${name.localPart}`) });
  assert.deepEqual(dropped, ['urn:zz}mine'], 'onUnexpectedAttribute reports the dropped attribute');

  // Nothing is reported when nothing is dropped, and a namespace declaration is not dropped content.
  const quiet = [];
  await unmarshalString(`<w:document xmlns:w="${W_}"><w:body><w:p/></w:body></w:document>`,
    { onUnexpectedAttribute: (name) => quiet.push(name.localPart), onUnexpectedElement: (name) => quiet.push(name.localPart) });
  assert.deepEqual(quiet, [], 'no callback fires for a part the model reads whole');

  // An element the parent does not accept is reported at the parent (the kind-2 case of CR-004 section 10.1).
  const elements = [];
  await unmarshalString(`<w:document xmlns:w="${W_}"><w:body><w:p/></w:body><w:sectPr/></w:document>`,
    { onUnexpectedElement: (name) => elements.push(name.localPart) });
  assert.deepEqual(elements, ['sectPr'], 'onUnexpectedElement reports an element its parent does not accept');

  // onElement records what a marshal wrote.
  const written = [];
  await marshalString(await unmarshalString(`<w:document xmlns:w="${W_}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>`),
    { onElement: (element) => written.push(element.localName) });
  assert.deepEqual(written.slice(0, 4), ['document', 'body', 'p', 'r'], 'onElement records the elements written');

  // A per-call prefix table applies to this call only. It governs names that carry no prefix of
  // their own: a QName unmarshalled from input keeps the prefix the input used, which is the
  // runtime's documented behaviour and was true of the hand-rolled derivation this replaced. So the
  // element here is built rather than parsed.
  const parsed = await unmarshalString(`<w:document xmlns:w="${W_}"><w:body><w:p/></w:body></w:document>`);
  const built = { name: { namespaceURI: W_, localPart: 'document' }, value: unwrap(parsed) };
  const mine = await marshalString(built, { namespacePrefixes: { ...NAMESPACE_PREFIXES, [W_]: 'ww' } });
  assert.match(mine, /^<ww:document /, 'the call sees its own table');
  const next = await marshalString(built);
  assert.match(next, /^<w:document /, 'and the shared context is unchanged by it');
  assert.match(await marshalString(parsed), /^<w:document /, 'a parsed root keeps the prefix its input used');
  const rels = await marshalString(await unmarshalString(`<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="http://x" Target="s.xml"/></Relationships>`),
    { namespacePrefixes: { ...NAMESPACE_PREFIXES, [SML_]: 'sx' } });
  assert.match(rels, /^<Relationships xmlns="/, "CR-001's per-root rule still applies over a caller's table");

  // The facade sets no reporting callback of its own: a drop is some callers' mechanism (CR-007 section 6). Its
  // onInvalidValue (CR-013) reads a value as Word does, and reports only to an onWordReading the caller passes.
  const sheet = `<worksheet xmlns="${SML_}" xmlns:xr="${XR_}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr"><sheetData><row r="1"><c r="A1" xr:uid="{1}"/></row></sheetData></worksheet>`;
  let fired = false;
  await unmarshalString(sheet);                                    // no options: silent
  await unmarshalString(sheet, { onUnexpectedAttribute: () => { fired = true; } });
  assert.equal(fired, true, 'xr:uid on a cell is not bound, and is reported when asked for');
  console.log('CR-007: per-call unmarshal and marshal options, and no default callback OK');
}

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
// An mc:Choice's Requires names prefixes too, declared whether or not its branch uses them: Word's
// chartex Choice says Requires="cx1" (2015/9/8) over content in cx (2014), and with nothing in cx1
// its declaration was stripped, leaving a Requires no reader resolves (core-ts, once its CR-010
// phase 4 kept branches by default, 2026-10-19; docx4j CR-049 declares the same).
{
  const CX1 = 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex';
  const kept = (requires, declaration) => unmarshalString(`<w:document xmlns:w="${W}" xmlns:mc="${MC}"><w:body><w:p><w:r><mc:AlternateContent>`
    + `<mc:Choice ${declaration} Requires="${requires}"><w:drawing/></mc:Choice><mc:Fallback><w:t>chart</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>`);
  const choiceXml = await marshalString(await kept('cx1', `xmlns:cx1="${CX1}"`));
  assert.ok(declarationsOf(choiceXml).declared.includes('cx1'), 'a Requires prefix nothing in its branch uses is declared on the root');
  assert.match(choiceXml, new RegExp(`xmlns:cx1="${CX1}"`), 'as NAMESPACE_PREFIXES has it');
  const warn = console.warn, warned = [];
  console.warn = (message) => warned.push(message);
  const fromSource = await marshalString(await kept('zz', 'xmlns:zz="urn:zz"'));
  assert.match(fromSource, /xmlns:zz="urn:zz"/, 'CR-011: a Requires prefix only the source declared is declared from the source');
  assert.equal(warned.length, 0);
  const unknownChoice = await marshalString(await kept('zz', '')).finally(() => { console.warn = warn; });
  assert.match(unknownChoice, /Requires="zz"/, 'a Requires prefix nothing resolves is kept, since dropping it would change the branch a reader takes');
  assert.match(warned.join('\n'), /Requires names zz/, 'and warned about');
}

// CR-011: a prefix only the source bound - one NAMESPACE_PREFIXES and the aliases cannot give, or
// give for another namespace - is recorded at unmarshal, on the Choice's own object (so it travels
// wherever an edit moves the Choice) and on the part root (for mc:Ignorable), and declared from
// that record before the table. The editor's fixture, docx4j's extraction-probe.docx (declared on
// mc:AlternateContent; shape2010 a second prefix for wps) and core-ts's insertOoxml are the cases.
// CR-012: a Requires prefix so declared goes on its mc:Choice, ahead of Requires, as Word writes it,
// unless the root already declares that binding; an mc:Ignorable prefix stays on the root. A
// Choice's own mc:Ignorable and mc:MustUnderstand follow its Requires (docx4j CR-052's D1 (b)).
{
  const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
  const { wml } = await import('../dist/builders/wml.mjs');
  const ac = (declaration = 'xmlns:zz9="urn:zz9"', requires = 'zz9') => `<mc:AlternateContent ${declaration}><mc:Choice Requires="${requires}"><w:t>round</w:t></mc:Choice><mc:Fallback><w:t>flat</w:t></mc:Fallback></mc:AlternateContent>`;
  const documentXml = (body, rootDeclarations = '') => `<w:document xmlns:w="${W}" xmlns:mc="${MC}" ${rootDeclarations}><w:body>${body}</w:body></w:document>`;
  const declares = (xml, prefix = 'zz9', namespaceURI = 'urn:zz9') => xml.includes(`xmlns:${prefix}="${namespaceURI}"`);
  const onChoice = (xml, prefix = 'zz9', namespaceURI = 'urn:zz9') => new RegExp(`<mc:Choice [^>]*xmlns:${prefix}="${namespaceURI}"[^>]* Requires="`).test(xml);
  const onRoot = (xml, prefix = 'zz9', namespaceURI = 'urn:zz9') => xml.slice(0, xml.indexOf('>')).includes(`xmlns:${prefix}="${namespaceURI}"`);
  const warn = console.warn, warned = [];
  console.warn = (message) => warned.push(String(message));
  try {
    const editors = await marshalString(await unmarshalString(documentXml(`<w:p><w:r>${ac('')}</w:r></w:p>`, 'xmlns:zz9="urn:zz9" mc:Ignorable="zz9"')));
    assert.ok(onRoot(editors) && /mc:Ignorable="zz9"/.test(editors), 'declared on the root: mc:Ignorable keeps zz9 and the Requires resolves');
    assert.equal(editors.split('xmlns:zz9=').length, 2, 'and the Choice doesn\'t repeat the root\'s binding');
    const rootOnly = await marshalString(await unmarshalString(documentXml(`<w:p><w:r>${ac('')}</w:r></w:p>`, 'xmlns:zz9="urn:zz9"')));
    assert.ok(onChoice(rootOnly) && !onRoot(rootOnly), 'a root declaration mc:Ignorable doesn\'t name moves to the Choice');
    const probe = await marshalString(await unmarshalString(documentXml(`<w:p><w:r>${ac()}</w:r><w:r>${ac(`xmlns:shape2010="${WPS}"`, 'shape2010')}</w:r></w:p>`)));
    assert.ok(onChoice(probe) && onChoice(probe, 'shape2010', WPS), 'declared on mc:AlternateContent, and a second prefix for a namespace the table knows: each on its Choice');
    assert.ok(!onRoot(probe) && !onRoot(probe, 'shape2010', WPS), 'not on the root');
    assert.equal(warned.length, 0, 'none of that warns');

    // Moved in from another source: a paragraph of another unmarshalled tree, and a wml() fragment.
    const destination = await unmarshalString(documentXml('<w:p/>'));
    const source = await unmarshalString(documentXml(`<w:p><w:r>${ac()}</w:r></w:p>`));
    const moved = source.value.body.content[0];
    destination.value.body.content.push(moved);
    const [fragment] = await wml(`<w:p xmlns:shape2010="${WPS}"><w:r>${ac('', 'shape2010')}</w:r></w:p>`);
    destination.value.body.content.push(fragment);
    const destinationXml = await marshalString(destination);
    assert.ok(onChoice(destinationXml) && onChoice(destinationXml, 'shape2010', WPS), 'the record travels with the Choice when content moves between trees');
    assert.ok(onChoice(await marshalString(moved)), 'a paragraph marshalled on its own declares what its Choice needs');
    const copied = await unmarshalString(documentXml('<w:p/>'));
    copied.value.body.content.push(deepCopy(moved, copied.value.body));
    assert.ok(onChoice(await marshalString(copied)), 'deepCopy carries the record');
    assert.match(await marshalString(deepCopy(await unmarshalString(documentXml('<w:p/>', 'xmlns:zz9="urn:zz9" mc:Ignorable="zz9"')))), /mc:Ignorable="zz9"/, 'and the root\'s');

    // A flat OPC part: its root's bindings and its Choices' are its own.
    const flat = `<pkg:package xmlns:pkg="${PKG}"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>${documentXml(`<w:p><w:r>${ac()}</w:r></w:p>`)}</pkg:xmlData></pkg:part></pkg:package>`;
    const flatXml = await marshalPackage(await unmarshalPackage(flat));
    assert.ok(onChoice(flatXml) && !/<w:document[^>]*xmlns:zz9=/.test(flatXml), 'a flat OPC part declares zz9 on its Choice');

    // A table declaration kept only because a Requires names its prefix goes; the Choice declares the source's.
    const superseded = await marshalString(await unmarshalString(documentXml(`<w:p><w:r>${ac('xmlns:w14="urn:not-w14"', 'w14')}</w:r></w:p>`)));
    assert.ok(onChoice(superseded, 'w14', 'urn:not-w14') && !onRoot(superseded, 'w14', W14), 'the source\'s w14, since nothing in the tree uses w14');
    assert.equal(warned.length, 0);

    // The tree's w14 outside the Choice doesn't stop the Choice declaring the source's.
    const both = await marshalString(await unmarshalString(documentXml(`<w:p w14:paraId="00000001"><w:r>${ac('xmlns:w14="urn:not-w14"', 'w14')}</w:r></w:p>`, `xmlns:w14="${W14}"`)));
    assert.ok(onRoot(both, 'w14', W14) && onChoice(both, 'w14', 'urn:not-w14'), 'w14:paraId keeps the tree\'s w14, and the Choice keeps the source\'s');
    assert.equal(warned.length, 0);

    // A declaration the Choice's own content uses is never rebound; the source's other meaning is warned about.
    const clash = await marshalString(await unmarshalString(documentXml(`<w:p><w:r><mc:AlternateContent xmlns:w14="urn:not-w14"><mc:Choice Requires="w14"><w:p xmlns:w14="${W14}" w14:paraId="00000001"/></mc:Choice><mc:Fallback><w:t>flat</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>`)));
    assert.ok(/w14:paraId="00000001"/.test(clash) && onRoot(clash, 'w14', W14) && !onChoice(clash, 'w14', 'urn:not-w14'), 'the Choice\'s w14:paraId keeps the tree\'s w14');
    assert.equal(warned.length, 1);
    assert.match(warned.splice(0)[0], /source bound w14 to urn:not-w14, but the tree uses it/);

    // One prefix bound two ways in one part keeps both meanings, each on its Choice; one root's
    // mc:Ignorable and a Choice, too.
    const twoWays = await marshalString(await unmarshalString(documentXml(`<w:p><w:r>${ac('xmlns:zz9="urn:A"')}</w:r><w:r>${ac('xmlns:zz9="urn:B"')}</w:r></w:p>`)));
    assert.ok(onChoice(twoWays, 'zz9', 'urn:A') && onChoice(twoWays, 'zz9', 'urn:B'));
    const rootAndChoice = await marshalString(await unmarshalString(documentXml(`<w:p><w:r>${ac('xmlns:zz9="urn:B"')}</w:r></w:p>`, 'xmlns:zz9="urn:A" mc:Ignorable="zz9"')));
    assert.ok(onRoot(rootAndChoice, 'zz9', 'urn:A') && onChoice(rootAndChoice, 'zz9', 'urn:B'));
    assert.equal(warned.length, 0);

    // A Choice's own mc:Ignorable and mc:MustUnderstand, as Requires: docx4j's
    // aChoicesOwnIgnorableAndMustUnderstandAreDeclared, whose table prefixes go on the root.
    const own = `<w:p><w:r><mc:AlternateContent xmlns:zz9="urn:zz9" xmlns:wps="${WPS}" xmlns:w14="${W14}"><mc:Choice Requires="wps" mc:Ignorable="zz9" mc:MustUnderstand="w14"><w:t>round</w:t></mc:Choice><mc:Fallback><w:t>flat</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>`;
    const ownDocument = await unmarshalString(documentXml(own));
    const ownXml = await marshalString(ownDocument);
    assert.ok(onRoot(ownXml, 'wps', WPS) && onRoot(ownXml, 'w14', W14), 'the table\'s wps and w14 on the root');
    assert.match(ownXml, /<mc:Choice xmlns:zz9="urn:zz9" Requires="wps" mc:Ignorable="zz9" mc:MustUnderstand="w14">/, 'the source\'s zz9 on the Choice, ahead of its attributes, which keep their order');
    assert.ok(!onRoot(ownXml), 'not on the root');
    assert.ok(onChoice(await marshalString(ownDocument.value.body.content[0])), 'the record travels with the Choice');
    assert.equal(warned.length, 0);
    const unresolved = await marshalString(await unmarshalString(documentXml(`<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14" mc:Ignorable="w14 qq" mc:MustUnderstand="rr"><w:t>round</w:t></mc:Choice><mc:Fallback><w:t>flat</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>`
      + `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14" mc:Ignorable="qq"><w:t>round</w:t></mc:Choice><mc:Fallback><w:t>flat</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>`)));
    assert.match(unresolved, /mc:Ignorable="w14" mc:MustUnderstand="rr">/, 'a MustUnderstand prefix nothing resolves is kept, and an Ignorable one dropped from the Choice');
    assert.match(unresolved, /<mc:Choice Requires="w14">/, 'with its mc:Ignorable, where it named nothing else');
    assert.deepEqual(warned.splice(0), [
      'marshal: an mc:Choice\'s mc:MustUnderstand names rr, which no declaration or NAMESPACE_PREFIXES entry resolves; kept',
      'marshal: an mc:Choice\'s mc:Ignorable names qq, which no declaration or NAMESPACE_PREFIXES entry resolves; dropped']);

    // A caller's own onElement still sees every element, on both sides.
    const unmarshalled = [], marshalled = [];
    const reported = await unmarshalString(documentXml(`<w:p><w:r>${ac()}</w:r></w:p>`), { onElement: (element) => unmarshalled.push(element.localName) });
    assert.ok(onChoice(await marshalString(reported, { onElement: (element) => marshalled.push(element.localName) })));
    assert.deepEqual(unmarshalled, ['t', 'Choice', 't', 'Fallback', 'AlternateContent', 'r', 'p', 'body', 'document'], 'children first (jsonix-CR-008)');
    assert.deepEqual(marshalled, ['document', 'body', 'p', 'r', 'AlternateContent', 'Choice', 't', 'Fallback', 't']);
  } finally {
    console.warn = warn;
  }
}
console.log('CR-011, CR-012: prefixes only the source bound are declared from the source, a Requires prefix on its Choice (and its own mc:Ignorable and mc:MustUnderstand), wherever the content goes OK');

// CR-006: a prefix mc:Ignorable names that the table cannot produce, because its namespace is
// written as the default. Excel's slicer cache: x is the SpreadsheetML main namespace, which this
// package writes as the default, so the prefix is declared beside the default rather than the token
// dropped (docx4j's getPreDeclaredNamespaceUris2, its CR-024). Found by test/fidelity.mjs.
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XR10 = 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10';
const slicerCacheXml = await marshalString(await unmarshalString(
  `<slicerCacheDefinition xmlns="${X14}" xmlns:mc="${MC}" mc:Ignorable="x xr10" xmlns:x="${SML}" xmlns:xr10="${XR10}" name="S" sourceName="R" xr10:uid="{1}"><pivotTables><pivotTable tabId="3" name="P"/></pivotTables><data><tabular pivotCacheId="1"><items count="0"/></tabular></data></slicerCacheDefinition>`));
const slicerCache = declarationsOf(slicerCacheXml);
assert.deepEqual(slicerCache.ignorable, ['x', 'xr10'], 'CR-006: the x token survives');
assert.deepEqual(slicerCache.undeclared, [], 'CR-006: and is declared');
assert.match(slicerCacheXml, new RegExp(`xmlns:x="${SML}"`), 'CR-006: x is declared as the SpreadsheetML main namespace');
// Excel writes that root's own namespace as the default; this package writes it with the table's
// prefix (CR-001), which is a prefix choice and not a difference in content.
assert.match(slicerCacheXml, /^<x14:slicerCacheDefinition /, "the root uses the table's prefix for x14");
assert.ok(!slicerCacheXml.slice(0, slicerCacheXml.indexOf('>')).includes(' xmlns='), 'and declares no default namespace');
assert.equal(IGNORABLE_PREFIX_ALIASES.x, SML, 'the alias table is exported, as NAMESPACE_PREFIXES is');

// CR-008: a part root that binds mc:Ignorable lists the Office extension namespaces its tree uses,
// by the prefix the call's table gives, after what it already lists. ECMA-376 Part 3 makes extension
// markup that is not declared ignorable an error for a reader that does not understand it; core-ts
// met it writing w16du:dateUtc, then in a caller's w15 content spliced in by insertOoxml.
{
  const W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
  const X14AC = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac';
  const X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
  const marshalled = async (xml, options) => marshalString(await unmarshalString(xml), options);
  const ins = '<w:ins w:id="1" w:author="A" w:date="2026-10-04T00:00:00Z" w16du:dateUtc="2026-10-04T00:00:00Z"><w:r><w:t>x</w:t></w:r></w:ins>';
  const datedXml = `<w:document xmlns:w="${W}" xmlns:w16du="${W16DU}"><w:body><w:p>${ins}</w:p></w:body></w:document>`;

  // Section 5's four. A document with no mc:Ignorable, as a created one has, whose w:ins carries dateUtc:
  const dated = await marshalled(datedXml);
  assert.deepEqual(declarationsOf(dated).ignorable, ['w16du'], 'CR-008: w16du:dateUtc lists w16du');
  assert.equal(declarations(dated).w16du, W16DU, 'CR-008: and declares it');
  // A loaded document keeps Word's ten, in order, and gains nothing it does not use.
  const WORD = 'w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14';
  const loaded = await marshalled(`<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:w16du="${W16DU}" xmlns:mc="${MC}" mc:Ignorable="${WORD}"><w:body><w:p w14:paraId="1">${ins}</w:p></w:body></w:document>`);
  assert.deepEqual(declarationsOf(loaded).ignorable, WORD.split(' '), 'CR-008: a loaded list is unioned, never replaced');
  // A root whose type binds no mc:Ignorable gains none: the schema does not admit it there.
  const paragraph = await marshalled(`<w:p xmlns:w="${W}" xmlns:w14="${W14}" xmlns:w16du="${W16DU}" w14:paraId="1">${ins}</w:p>`);
  assert.ok(!paragraph.includes('Ignorable'), 'CR-008: w:p binds no mc:Ignorable, so none is written');
  // The prefix is the call's table's. A SpreadsheetML part, whose main namespace is the default and
  // so has no prefix to list (CR-006), lists x14ac; the slicer cache above kept "x xr10" whole, and
  // its root's own x14 is not added. A per-call table's prefix is the one listed.
  const sheet = await marshalled(`<worksheet xmlns="${SML}" xmlns:x14ac="${X14AC}"><sheetData><row r="1" x14ac:dyDescent="0.25"/></sheetData></worksheet>`);
  assert.deepEqual(declarationsOf(sheet).ignorable, ['x14ac'], 'CR-008: x14ac:dyDescent lists x14ac');
  assert.equal(declarations(sheet).x14ac, X14AC);
  assert.equal(declarations(sheet)[''], SML, 'CR-008: and the default namespace is untouched');
  const renamed = await marshalString({ name: { namespaceURI: W, localPart: 'document' }, value: unwrap(await unmarshalString(datedXml)) },
    { namespacePrefixes: { ...NAMESPACE_PREFIXES, [W16DU]: 'du' } });
  assert.deepEqual(declarationsOf(renamed).ignorable, ['du'], "CR-008: the prefix listed is the call's table's");
  assert.equal(declarations(renamed).du, W16DU);

  // Section 8's two: content a caller brings in. A literal w:document (no TYPE_NAME, so the root's
  // type comes from its element name) whose content control was built from a fragment with a w15
  // element, as core-ts's insertOoxml splices one in:
  const { wml } = await import('@docx4j/generated-objects-ts/builders/wml');
  const control = await wml('<w:sdt><w:sdtPr><w15:appearance w15:val="tags"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>');
  const created = await marshalString({ name: { namespaceURI: W, localPart: 'document' }, value: { body: { content: control } } });
  assert.deepEqual(declarationsOf(created).ignorable, ['w15'], 'CR-008: a w15 element in w:sdtPr lists w15');
  assert.equal(declarations(created).w15, W15);
  // And a root that already lists w15, as core-ts's own declareIgnorable leaves it, is unchanged.
  const declared = await marshalled(`<w:document xmlns:w="${W}" xmlns:w15="${W15}" xmlns:mc="${MC}" mc:Ignorable="w14 w15"><w:body><w:sdt><w:sdtPr><w15:appearance w15:val="tags"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt></w:body></w:document>`);
  assert.deepEqual(declarationsOf(declared).ignorable, ['w14', 'w15'], 'CR-008: a list already naming w15 is unchanged');

  // Where Office does not list a namespace it uses, nor does this: in a branch whose mc:Choice
  // requires it (but in one requiring something else it counts), below an extension list's ext,
  // and as the root's own namespace (Excel's x15 timelines). w14:x is unknown to the model, so it is
  // held as a DOM node: the walk is over the marshalled DOM, and wildcard content counts like typed.
  const branch = (requires) => marshalled(`<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}"><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="${requires}"><w14:x/></mc:Choice><mc:Fallback><w:t>a</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>`);
  assert.deepEqual(declarationsOf(await branch('w14')).ignorable, [], 'CR-008: not inside an mc:Choice that requires it');
  assert.deepEqual(declarationsOf(await branch('wps')).ignorable, ['w14'], 'CR-008: but inside one requiring another namespace');
  const styles = await marshalled(`<styleSheet xmlns="${SML}"><extLst><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="${X15}"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>`);
  assert.ok(!styles.includes('Ignorable'), 'CR-008: not below an extension list\'s ext');
  const timelines = await marshalled(`<timelines xmlns="${X15}"><timeline name="Date" cache="C" caption="Date" level="2" selectionLevel="2" scrollPosition="2026-05-19T00:00:00"/></timelines>`);
  assert.ok(!timelines.includes('Ignorable'), "CR-008: not the root's own namespace");

  assert.ok(Object.isFrozen(IGNORABLE_EXTENSION_NAMESPACES) && IGNORABLE_EXTENSION_NAMESPACES.includes(W16DU), 'the table is exported');
  console.log('CR-008: extension namespaces a part root uses are listed in its mc:Ignorable OK');
}

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
  assert.throws(() => applyRunOptions({}, { highlightColor: 'darkYellow' }), /Not a highlight colour/);
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
  // declarations promise"). A standalone w:tc / w:tbl / w:drawing root carried xsi:type until
  // CR-009 declared them globally; now it marshals as docx4j's root classes do, without one.
  assert.equal(await marshalString(tc('a', { width: 1000 })),
    `<w:tc xmlns:w="${W}"><w:tcPr><w:tcW w:type="dxa" w:w="1000"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>`);
  assert.equal(await marshalString(tc([])),
    `<w:tc xmlns:w="${W}"><w:p/></w:tc>`,
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

// CR-009: the WordprocessingML elements docx4j declares as roots are declared globally by one more
// mapping, so a kept mc:Choice or mc:Fallback types its w:drawing and w:pict as docx4j's context does.
{
  resetContext();
  const ctx = await getContext();
  const names = Object.keys(WML_GLOBAL_ELEMENTS);
  assert.equal(names.length, 61, 'CR-009 lists 61');
  const wmlGlobals = new Set(MODULES.org_docx4j_wml.elementInfos.filter((e) => !e.scope).map((e) => e.elementName.localPart ?? e.elementName));
  for (const [localPart, type] of Object.entries(WML_GLOBAL_ELEMENTS)) {
    const info = ctx.getElementInfo({ key: `{${W}}${localPart}` }, undefined);
    assert.equal(info?.typeInfo?.name, `org_docx4j_wml.${type}`, `w:${localPart} resolves globally to ${type}`);
    assert.ok(!wmlGlobals.has(localPart), `w:${localPart} is not declared globally by the generated module itself (a regeneration that does retires it here)`);
  }
  assert.ok(modulesFor('org_docx4j_wml').some((m) => m.name === 'org_docx4j_wml_globals'), 'modulesFor adds the globals to a closure holding wml');
  assert.ok(!modulesFor('org_docx4j_dml').some((m) => m.name === 'org_docx4j_wml_globals'), 'and not to one without it');
  const V = 'urn:schemas-microsoft-com:vml';
  const kept = await unmarshalString(`<w:p xmlns:w="${W}" xmlns:mc="${MC}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:v="${V}" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w14="${W14}"><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0"><wp:extent cx="1" cy="2"/></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shapetype id="_x0000_t202" w14:anchorId="7969C50B" o:spt="202" path="m,l,21600r21600,l21600,xe"/></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>`);
  const ac = unwrap(kept).content[0].value.content[0].value;
  assert.equal(ac.choice[0].any[0].value.TYPE_NAME, 'org_docx4j_wml.Drawing', 'a kept Choice types its w:drawing');
  assert.equal(ac.fallback.any[0].value.TYPE_NAME, 'org_docx4j_wml.Pict', 'a kept Fallback types its w:pict');
  assert.equal(ac.fallback.any[0].value.anyAndAny[0].value.TYPE_NAME, 'org_docx4j_vml.CTShapetype', 'and the VML inside it');
  const written = await marshalString(kept);
  assert.match(written, /<w:drawing><wp:inline distT="0"><wp:extent cx="1" cy="2"\/><\/wp:inline><\/w:drawing>/, 'the Choice marshals whole');
  assert.match(written, /<v:shapetype [^>]*w14:anchorId="7969C50B"[^>]*o:spt="202"/, 'the Fallback keeps w14:anchorId (docx4j CR-048) and writes o:spt as read');
  assert.equal(unwrap(await unmarshalString(`<w:drawing xmlns:w="${W}"/>`)).TYPE_NAME, 'org_docx4j_wml.Drawing', 'each of the 61 is a valid root now, as docx4j\'s root classes are');
  console.log('CR-009: 61 WordprocessingML elements declared globally; kept mc branches typed as docx4j types them OK');
}

// docx4j CR-051 (xsd 1cff2cd5c): mc:AlternateContent is admitted in EG_ContentBlockContent,
// EG_ContentRowContent and EG_ContentCellContent, where every desktop Word reads it. The runtime
// already typed it there, since an elementRefs property resolves an element it doesn't list
// among the context's global ones; the regeneration lists it, and the declarations admit it.
{
  for (const local of ['Body', 'CTCustomXmlBlock', 'CTFtnEdn', 'CTTxbxContent', 'Comments.Comment', 'Hdr', 'Ftr', 'SdtContentBlock', 'Tc', 'Tbl', 'CTSdtContentRow', 'CTCustomXmlRow', 'Tr', 'CTSdtContentCell', 'CTCustomXmlCell']) {
    const typeInfo = MODULES.org_docx4j_wml.typeInfos.find((t) => t.localName === local);
    const refs = typeInfo.propertyInfos.find((pi) => pi.name === 'content').elementTypeInfos;
    assert.ok(refs.some((r) => r.elementName.localPart === 'AlternateContent' && r.elementName.namespaceURI === MC), `${local}'s content lists mc:AlternateContent`);
  }
  const branch = (choice, fallback) => `<mc:AlternateContent><mc:Choice Requires="w14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
  const block = branch('<w:p><w:r><w:t>new</w:t></w:r></w:p>', '<w:p><w:r><w:t>old</w:t></w:r></w:p>');
  const cell = branch('<w:tc><w:p/></w:tc>', '<w:tc><w:p/></w:tc>');
  const row = branch(`<w:tr>${cell}</w:tr>`, '<w:tr><w:tc><w:p/></w:tc></w:tr>');
  const xml = `<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:w14="${W14}"><w:body>${block}<w:tbl>${row}</w:tbl></w:body></w:document>`;
  const doc = await unmarshalString(xml);
  const [blockAc, tbl] = doc.value.body.content;
  assert.equal(blockAc.value.TYPE_NAME, 'org_docx4j_mce.AlternateContent', 'block level: in w:body');
  assert.equal(blockAc.value.PARENT, doc.value.body);
  assert.equal(blockAc.value.choice[0].any[0].value.TYPE_NAME, 'org_docx4j_wml.P', 'its branches typed (CR-009)');
  const rowAc = tbl.value.content[0];
  assert.equal(rowAc.value.TYPE_NAME, 'org_docx4j_mce.AlternateContent', 'row level: in w:tbl');
  const tr = rowAc.value.choice[0].any[0];
  assert.equal(tr.value.TYPE_NAME, 'org_docx4j_wml.Tr');
  assert.equal(tr.value.content[0].value.TYPE_NAME, 'org_docx4j_mce.AlternateContent', 'cell level: in w:tr');
  const written = await marshalString(doc);
  assert.ok(written.includes(`<w:body>${block}<w:tbl>${row}</w:tbl></w:body>`), 'all three marshal back as read');
  console.log('docx4j CR-051: mc:AlternateContent listed at block, row and cell level, typed and marshalled whole OK');
}

// CR-013: run properties read as Word 365 reads them (Jason's probes, 2026-10-24; test/word365.mjs reads the probes
// themselves). Phase A, helpers/wml's colour table and readRunOptions; phase B, the facade's onInvalidValue.
{
  const { readRunOptions } = await import('../dist/builders/wml.mjs');
  const { WORD_COLOR_NAMES, wordColor } = await import('../dist/helpers/wml.mjs');
  const { wordInvalidValue } = await import('../dist/index.mjs');
  const W_ = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const W14_ = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const WP_ = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';

  // Phase A, without the runtime: w:highlight outside ST_HighlightColor (case-exact) is none, as Word repairs it.
  for (const val of ['YELLOW', 'bogus', 'none']) assert.equal(readRunOptions({ highlight: { val } }).highlightColor, null, `highlight ${val}`);
  assert.equal(readRunOptions({ highlight: { val: 'yellow' } }).highlightColor, '#FFFF00');
  // w:color: hex with #, any case, padded; exactly auto; the names Word reads; anything else as if absent.
  const color = (val) => readRunOptions({ color: { val } }).color;
  for (const val of ['FF0000', 'ff0000', '#ff0000', '#FF0000', ' FF0000 ', 'FF0000 ', ' #Ff0000', 'red', 'RED', 'Red', ' red ']) assert.equal(color(val), '#FF0000', `color ${JSON.stringify(val)}`);
  assert.equal(color('auto'), 'auto');
  for (const [val, hex] of [['darkGray', '#A9A9A9'], ['lightgrey', '#D3D3D3'], ['lightgoldenrodyellow', '#FAFAD2'], ['Navy', '#000080'], ['DarkRed', '#8B0000'], ['papayawhip', '#FFEFD5'], ['dimgray', '#696969']]) assert.equal(color(val), hex, val);
  for (const val of ['Auto', 'AUTO', ' auto', 'auto ', '#auto', 'grey', 'lightGray', 'darkYellow', 'darkgrey', 'rebeccapurple', 'transparent', 'windowtext', 'dark red', '#red', 'F00', '0xFF0000', 'FF00000', 'FF000', 'GG0000', '', ' ']) {
    assert.equal(color(val), '', `color ${JSON.stringify(val)} reads as absent`);
  }
  assert.equal(readRunOptions({}).color, '', 'absent');
  // wordColor itself: the same rule; undefined where Word rejects the value, for the caller's place to read.
  assert.equal(wordColor(undefined), undefined);
  assert.equal(wordColor('auto'), 'auto');
  assert.equal(wordColor('\t#00ff7f\n'), '#00FF7F', 'trimmed as Java trims');
  assert.equal(wordColor(' SpringGreen '), '#00FF7F', 'a name, trimmed, in any case (docx4j CR-060 D7)');
  for (const val of ['Auto', 'grey', 'F00', '', 'constructor', 'toString', '__proto__']) assert.equal(wordColor(val), undefined, JSON.stringify(val));
  // The table: 140 classic HTML names, lower-case, to RRGGBB; CSS 3's 147 less six grey spellings and lightgray.
  const names = Object.keys(WORD_COLOR_NAMES);
  assert.equal(names.length, 140);
  assert.ok(Object.isFrozen(WORD_COLOR_NAMES) && Object.getPrototypeOf(WORD_COLOR_NAMES) === null);
  assert.ok(names.every((n) => /^[a-z]+$/.test(n)) && Object.values(WORD_COLOR_NAMES).every((h) => /^[0-9A-F]{6}$/.test(h)));
  for (const n of ['grey', 'darkgrey', 'darkslategrey', 'dimgrey', 'lightslategrey', 'slategrey', 'lightgray']) assert.ok(!(n in WORD_COLOR_NAMES), `${n} is not a name Word reads`);
  for (const n of ['lightgrey', 'gray', 'darkgray', 'darkslategray', 'dimgray', 'lightslategray', 'slategray']) assert.ok(n in WORD_COLOR_NAMES, n);
  // Held to docx4j's WordColorValues, name for name and value for value, where docx4j is checked out beside.
  const java = new URL('../../docx4j/docx4j-generated-objects/src/main/java/org/docx4j/wml/WordColorValues.java', import.meta.url);
  let docx4jTable = 'not checked (no ../docx4j)';
  try {
    const source = readFileSync(java, 'utf8');
    const body = source.slice(source.indexOf('HTML = table(') + 'HTML = table('.length);
    const strings = [...body.slice(0, body.indexOf(');')).matchAll(/"([^"]*)"/g)].map((m) => m[1]);
    const pairs = Object.fromEntries(strings.flatMap((s, i) => i % 2 === 0 ? [[s, strings[i + 1]]] : []));
    assert.deepEqual({ ...WORD_COLOR_NAMES }, pairs, "WORD_COLOR_NAMES is docx4j's WordColorValues table");
    docx4jTable = "docx4j's WordColorValues";
  } catch (e) {
    if (e.code !== 'ENOENT') throw e;
  }
  // underline: a value outside ST_Underline (set by hand, or kept by a caller's own onInvalidValue) is no underline.
  assert.equal(readRunOptions({ u: { val: 'bogus' } }).underline, 'None', 'not Mixed');
  assert.equal(readRunOptions({ u: {} }).underline, 'None', 'a bare w:u');

  // Phase B, with the runtime: on/off (core-ts CR-015 rule 1). Off is exactly false, off or 0; anything else is on.
  const rPrOf = async (inner, options) => unwrap(await unmarshalString(`<w:rPr xmlns:w="${W_}" xmlns:w14="${W14_}">${inner}</w:rPr>`, options));
  for (const val of ['true', '1', 'on', 'yes', 'FALSE', 'False', 'OFF', ' false', 'false ', '00', 'no', '', ' ']) {
    assert.equal(readRunOptions(await rPrOf(`<w:b w:val="${val}"/>`)).bold, true, `w:b ${JSON.stringify(val)} is on`);
  }
  for (const val of ['false', 'off', '0']) assert.equal(readRunOptions(await rPrOf(`<w:b w:val="${val}"/>`)).bold, false, `w:b ${val} is off`);
  assert.equal((await rPrOf('<w:caps w:val="yes"/>')).caps.val, true, 'every wml boolean, not only w:b');
  // An enumeration value outside its list is dropped (as if absent, so a style's value shows): w:u, w:vertAlign, w:em.
  const invalid = await rPrOf('<w:u w:val="Single" w:color="FF0000"/><w:vertAlign w:val="Superscript"/><w:em w:val="DOT"/>');
  assert.deepEqual([invalid.u.val, invalid.u.color, invalid.vertAlign.val, invalid.em.val], [undefined, 'FF0000', undefined, undefined]);
  const read = readRunOptions(invalid);
  assert.deepEqual([read.underline, read.superscript, read.subscript], ['None', false, false]);
  assert.equal((await rPrOf('<w:u w:val="bogus"/>')).u.val, undefined);
  // Except w:u's on/off words, exactly, which Word reads as none (round 6); True and yes are invalid like any other.
  for (const val of ['true', 'false', 'on', 'off', '1', '0']) {
    const rPr = await rPrOf(`<w:u w:val="${val}"/>`);
    assert.equal(rPr.u.val, 'none', `w:u ${val} is none`);
    assert.match(await marshalString({ name: { namespaceURI: W_, localPart: 'rPr' }, value: rPr }), /<w:u w:val="none"\/>/, `and saves as none`);
  }
  for (const val of ['True', 'TRUE', 'yes', 'no']) assert.equal((await rPrOf(`<w:u w:val="${val}"/>`)).u.val, undefined, `w:u ${val} is absent`);
  // A dropped value isn't written; a refused boolean is written true, which Word reads the same. The valueless
  // w:vertAlign and w:em are left out by phase C's normaliser, as Word's repair does, and kept with it off.
  const refused = await unmarshalString(`<w:rPr xmlns:w="${W_}"><w:b w:val="yes"/><w:u w:val="Single"/><w:vertAlign w:val="Superscript"/><w:em w:val="DOT"/></w:rPr>`);
  assert.match(await marshalString(refused, { normaliseOnSave: false }), /><w:b w:val="true"\/><w:u\/><w:vertAlign\/><w:em\/><\/w:rPr>$/, 'no refused value is written back');
  assert.match(await marshalString(refused), /><w:b w:val="true"\/><w:u\/><\/w:rPr>$/, 'and phase C leaves out the valueless ones');
  // Scope (core-ts CR-015's D10): Word's own WordprocessingML mappings. w14's CT_OnOff is a boolean, so TRUE is on.
  const checkbox = unwrap(await unmarshalString(`<w:sdtPr xmlns:w="${W_}" xmlns:w14="${W14_}"><w14:checkbox><w14:checked w14:val="TRUE"/></w14:checkbox></w:sdtPr>`));
  assert.equal(checkbox.rPrOrAliasOrLock[0].value.checked.val, true, 'w14:checked TRUE is on');
  // w14's own ST_OnOff enumeration is written on the property, where the runtime doesn't read it: a string, never refused.
  // (If a regeneration types them as an enumeration, this fails, and wordInvalidValue needs their rule: CR-013 section 3.5.)
  const w14 = await rPrOf('<w14:textFill><w14:gradFill><w14:lin w14:ang="0" w14:scaled="on"/></w14:gradFill></w14:textFill><w14:stylisticSets><w14:styleSet w14:id="1" w14:val="yes"/></w14:stylisticSets>');
  assert.deepEqual([w14.textFill.gradFill.lin.scaled, w14.stylisticSets.styleSet[0].val], ['on', 'yes']);
  // Every other mapping stays strict: a DrawingML boolean or a wml integer in the same part still fails it.
  const inDocument = (inner) => `<w:document xmlns:w="${W_}" xmlns:wp="${WP_}"><w:body><w:p><w:r><w:rPr><w:b w:val="yes"/></w:rPr>${inner}</w:r></w:p></w:body></w:document>`;
  await assert.rejects(unmarshalString(inDocument('<w:drawing><wp:anchor behindDoc="FALSE"/></w:drawing>')), /behindDoc/, 'wp:anchor behindDoc="FALSE" fails the part');
  await assert.rejects(unmarshalString(inDocument('<w:drawing><wp:inline><wp:docPr id="1" name="x" hidden="yes"/></wp:inline></w:drawing>')), /hidden/, 'wp:docPr hidden="yes" too');
  // A refused wml integer is CR-015's: a form Word refuses is dropped. Another simple type still fails the part.
  assert.equal((await rPrOf('<w:sz w:val="big"/>')).sz.val, undefined, 'a refused wml integer is dropped (CR-015)');
  await assert.rejects(unmarshalString(`<w:document xmlns:w="${W_}"><w:body><w:p><w:ins w:id="1" w:author="a" w:date="bogus"><w:r><w:t>x</w:t></w:r></w:ins></w:p></w:body></w:document>`), /date/, 'a refused wml dateTime fails as before');
  // Every facade unmarshal: a flat OPC package and a wml() fragment.
  const flatBold = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>${inDocument('')}</pkg:xmlData></pkg:part></pkg:package>`;
  const flatPart = (await unmarshalPackage(flatBold)).value.part[0].xmlData.any;
  assert.equal(flatPart.value.body.content[0].value.content[0].value.rPr.b.val, true, 'unmarshalPackage');
  const { wml } = await import('../dist/builders/wml.mjs');
  assert.equal((await wml('<w:r><w:rPr><w:i w:val="yes"/></w:rPr></w:r>'))[0].value.rPr.i.val, true, 'wml()');
  // The record (core-ts CR-015's D8, O6): onWordReading, the data for the caller's warnings; nothing to the console.
  const readings = [];
  await rPrOf('<w:b w:val=" false"/><w:i w:val="yes"/><w:u w:val="true"/><w:vertAlign w:val=" superscript"/><w:em w:val="DOT"/><w:sz w:val="24"/>', { onWordReading: (r) => readings.push(r) });
  assert.deepEqual(readings.map((r) => [r.element.localName, r.attribute.localPart, r.text, r.value, r.kind]), [
    ['b', 'val', ' false', true, 'value.kept.exact'],
    ['i', 'val', 'yes', true, 'value.kept.invalid'],
    ['u', 'val', 'true', 'none', 'value.kept.invalid'],
    ['vertAlign', 'val', ' superscript', undefined, 'value.kept.exact'],
    ['em', 'val', 'DOT', undefined, 'value.kept.invalid'],
  ]);
  // A caller's own onInvalidValue takes precedence, and onWordReading is then not called; wordInvalidValue composes.
  const mine = [];
  const own = await rPrOf('<w:b w:val="yes"/><w:u w:val="Single"/>', { onInvalidValue: (text) => { mine.push(text); return text === 'yes' ? false : 'double'; }, onWordReading: () => assert.fail('not called') });
  assert.deepEqual([own.b.val, own.u.val, mine], [false, 'double', ['yes', 'Single']]);
  const composed = await rPrOf('<w:b w:val="yes"/>', { onInvalidValue: (...args) => { mine.push('composed'); return wordInvalidValue(...args); } });
  assert.deepEqual([composed.b.val, mine.at(-1)], [true, 'composed']);
  await assert.rejects(rPrOf('<w:b w:val="yes"/>', { onInvalidValue: (text, p, c, n, error) => { throw error; } }), /w:b\/@w:val/, "a caller can restore the runtime's strictness");
  console.log(`CR-013: colours (${docx4jTable}), highlight, on/off, enumerations and w:u's on/off words read as Word 365 reads them; the scope, the record and the caller's precedence OK`);
}

// CR-013 phase C: what a save writes, run properties in their normal form (docx4j CR-060 with CR-061, value for value;
// test/word365.mjs holds every probe to docx4j's own saves). On jsonix-CR-011's marshal-side adapters, @docx4j/jsonix 3.8.0.
{
  const { setNormaliseOnSave, marshalNode } = await import('../dist/index.mjs');
  const W_ = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const W15_ = 'http://schemas.microsoft.com/office/word/2012/wordml';
  const snapshot = (v) => JSON.stringify(v, (k, x) => (k === 'PARENT' ? undefined : x));
  const runOf = (xml, marker) => { const i = xml.indexOf(`>${marker}<`); return xml.substring(Math.max(xml.lastIndexOf('<w:r>', i), xml.lastIndexOf('<w:r ', i)), xml.indexOf('</w:r>', i)); };
  // docx4j's ColourNormalFormTest.synthetic(): a theme colour, a revision's former properties, a paragraph mark's.
  const rev = ' w:author="A" w:date="2026-10-24T00:00:00Z"';
  const document = `<w:document xmlns:w="${W_}" xmlns:w15="${W15_}"><w:background w:color="navy"/><w:body>`
    + '<w:p><w:r><w:rPr><w:color w:val="bogus" w:themeColor="accent1"/></w:rPr><w:t>T1_THEME</w:t></w:r></w:p>'
    + `<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="1"${rev}><w:rPr><w:color w:val="F00"/><w:highlight w:val="YELLOW"/><w:vertAlign/><w:em w:val="DOT"/>`
    + '<w:u w:val="single" w:color="red"/><w:shd w:val="clear" w:color="auto" w:fill="bogus"/></w:rPr></w:rPrChange></w:rPr><w:t>T2_REVISION</w:t></w:r></w:p>'
    + `<w:p><w:r><w:rPr><w:i/><w:rPrChange w:id="2"${rev}><w:rPr><w:color w:val="red"/><w:highlight w:val="green"/></w:rPr></w:rPrChange></w:rPr><w:t>T3_REVISION_NAME</w:t></w:r></w:p>`
    + `<w:p><w:pPr><w:rPr><w:b/><w:rPrChange w:id="3"${rev}><w:rPr><w:color w:val="F00"/><w:highlight w:val="bogus"/><w:color w:val="navy"/></w:rPr></w:rPrChange></w:rPr></w:pPr><w:r><w:t>T4_MARK_REVISION</w:t></w:r></w:p>`
    // and beyond docx4j's: a theme tint or shade alone, a paragraph border, a pattern colour, w15's colour (not normalised, as docx4j).
    + '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:color="F00"/><w:bottom w:val="single" w:color="lime"/></w:pBdr></w:pPr><w:r><w:rPr><w:color w:val="" w:themeTint="99"/><w:shd w:val="pct50" w:color="Navy" w:fill="ffffff"/><w:highlight w:val="none"/></w:rPr><w:t>T5_TINT</w:t></w:r></w:p>'
    + '<w:p><w:r><w:rPr><w:color w:val="Auto" w:themeShade="BF"/><w:highlight/></w:rPr><w:t>T6_SHADE</w:t></w:r></w:p>'
    + '<w:sdt><w:sdtPr><w15:color w:val="red"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>'
    + '</w:body></w:document>';
  const element = await unmarshalString(document);
  const before = snapshot(element);
  const records = [];
  const xml = await marshalString(element, { onNormalised: (n) => records.push(n) });
  assert.equal(snapshot(element), before, 'marshal-only: the tree keeps what it read (docx4j D6)');
  assert.match(runOf(xml, 'T1_THEME'), /<w:color w:themeColor="accent1" w:val="auto"\/>/, 'a rejected value beside a theme colour is auto');
  const t2 = runOf(xml, 'T2_REVISION');
  for (const left of ['<w:color', '<w:highlight', '<w:vertAlign', '<w:em']) assert.ok(!t2.includes(left), `a revision's former ${left} is left out`);
  assert.ok(t2.includes('<w:u w:color="FF0000" w:val="single"/>') && t2.includes('w:fill="auto"'), "a revision's former properties are written as live ones");
  assert.ok(runOf(xml, 'T3_REVISION_NAME').includes('<w:color w:val="FF0000"/><w:highlight w:val="green"/>'));
  const mark = xml.substring(xml.lastIndexOf('<w:p>', xml.indexOf('T4_MARK_REVISION')), xml.indexOf('T4_MARK_REVISION'));
  assert.ok(!mark.includes('F00"') && !mark.includes('<w:highlight') && mark.includes('<w:color w:val="000080"/>'), "a paragraph mark's former properties");
  assert.ok(xml.includes('<w:pBdr><w:top w:color="auto" w:val="single"/><w:bottom w:color="00FF00" w:val="single"/></w:pBdr>'), 'a CT_Border under any of its names');
  assert.ok(runOf(xml, 'T5_TINT').includes('<w:color w:themeTint="99" w:val="auto"/><w:highlight w:val="none"/><w:shd w:color="000080" w:fill="FFFFFF" w:val="pct50"/>'), 'beside a theme tint alone; none kept; a pattern colour');
  assert.ok(runOf(xml, 'T6_SHADE').includes('<w:rPr><w:color w:themeShade="BF" w:val="auto"/></w:rPr>'), 'beside a theme shade alone; a valueless highlight left out');
  assert.ok(xml.includes('<w:background w:color="000080"/>') && xml.includes('<w15:color w:val="red"/>'), "the page colour; w15's colour as read, as docx4j");
  const former = (i) => unwrap(element).body.content[i].value.content[0].value.rPr.rPrChange.rPr.egrPrBase;
  assert.equal(former(2)[0].value.val, 'red', 'the former w:color still reads red in memory');
  assert.equal(former(1).length, 6, 'and the elements left out are still there');
  // The report: one record per value rewritten or left out, with where it is and what it was.
  assert.equal(records.length, 19);
  const at = (path) => records.find((n) => n.path === path);
  assert.deepEqual({ ...at('/w:document/w:body[1]/w:p[1]/w:r[1]/w:rPr[1]/w:color[1]/@w:val'), element: undefined }, { element: undefined, path: '/w:document/w:body[1]/w:p[1]/w:r[1]/w:rPr[1]/w:color[1]/@w:val', before: 'bogus', after: 'auto' });
  const highlight = at('/w:document/w:body[1]/w:p[2]/w:r[1]/w:rPr[1]/w:rPrChange[1]/w:rPr[1]/w:highlight');
  assert.deepEqual([highlight.before, highlight.after, highlight.element.localName], ['YELLOW', null, 'rPr'], 'an element left out: its parent, its w:val');
  assert.equal(at('/w:document/w:body[1]/w:p[2]/w:r[1]/w:rPr[1]/w:rPrChange[1]/w:rPr[1]/w:vertAlign').before, undefined);
  assert.equal(at('/w:document/w:background/@w:color')?.after, undefined, 'positions after the root only');
  assert.equal(at('/w:document/w:background[1]/@w:color').after, '000080');
  // A style's run properties and the document defaults.
  const styles = await unmarshalString(`<w:styles xmlns:w="${W_}"><w:docDefaults><w:rPrDefault><w:rPr><w:color w:val="maroon"/></w:rPr></w:rPrDefault></w:docDefaults>`
    + '<w:style w:type="character" w:styleId="S"><w:rPr><w:color w:val="F00"/><w:highlight w:val="Yellow"/><w:u w:val="double" w:color="bogus"/></w:rPr></w:style></w:styles>');
  const stylesXml = await marshalString(styles);
  assert.ok(stylesXml.includes('<w:rPrDefault><w:rPr><w:color w:val="800000"/></w:rPr></w:rPrDefault>') && stylesXml.includes('<w:rPr><w:u w:val="double"/></w:rPr></w:style>'), 'styles');
  // The opt-out: per call, and process-wide with a call's setting winning; off writes as read, with no records.
  const quiet = [];
  const asRead = await marshalString(element, { normaliseOnSave: false, onNormalised: (n) => quiet.push(n) });
  assert.ok(asRead.includes('<w:color w:themeColor="accent1" w:val="bogus"/>') && asRead.includes('<w:vertAlign/>') && asRead.includes('<w:background w:color="navy"/>'));
  setNormaliseOnSave(false);
  try {
    assert.equal(await marshalString(element, { onNormalised: (n) => quiet.push(n) }), asRead);
    assert.equal(await marshalString(element, { normaliseOnSave: true }), xml);
  } finally {
    setNormaliseOnSave(true);
  }
  assert.equal(quiet.length, 0);
  // A caller's own adapters are asked first: an answer wins, undefined falls through to the normaliser.
  const own = await marshalString(element, {
    onAttributeValue: (value, text, owner, propertyInfo, classInfo) => (classInfo.name === 'org_docx4j_wml.CTBackground' ? 'ABCDEF' : undefined),
    omitElement: (value, name) => (name.localPart === 'highlight' ? false : undefined),
  });
  assert.ok(own.includes('<w:background w:color="ABCDEF"/>') && own.includes('<w:highlight w:val="YELLOW"/>') && own.includes('<w:color w:themeColor="accent1" w:val="auto"/>'));
  // Every facade marshal: marshalNode, and marshalPackage's typed parts.
  assert.ok((await marshalNode(element)).getElementsByTagNameNS(W_, 'background')[0].getAttributeNS(W_, 'color') === '000080', 'marshalNode');
  const flatColour = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="${W_}"><w:body><w:p><w:r><w:rPr><w:color w:val="red"/></w:rPr></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
  assert.match(await marshalPackage(await unmarshalPackage(flatColour)), /<w:color w:val="FF0000"\/>/, 'marshalPackage');
  console.log('CR-013 phase C: colours in their normal form, the repaired elements left out, in run, paragraph mark, revision and style properties; the tree unchanged; the report, the opt-out and a caller\'s adapters OK');
}

// CR-014: the highlight palette as Word and Office JS read it. core-ts's check 61 (Word 365 through Office JS: 18
// reads, 45 writes) and check 62 (7,225 colours written to font.highlightColor), copied under test/fixtures/check61/
// and check62/.
{
  const { applyRunOptions, readRunOptions } = await import('../dist/builders/wml.mjs');
  const { HIGHLIGHT_COLORS, highlightNameForColor, WORD_COLOR_NAMES } = await import('../dist/helpers/wml.mjs');
  const check61 = JSON.parse(readFileSync(new URL('./fixtures/check61/61-highlights-word365.json', import.meta.url), 'utf8'));
  const check62 = JSON.parse(readFileSync(new URL('./fixtures/check62/62-highlight-snap-word365.json', import.meta.url), 'utf8'));
  assert.deepEqual([check61.reads.length, check61.writes.length, check62.points.length], [18, 45, 7225]);

  // Reads: the sixteen names in Word's palette, as font.highlightColor reads them; none and a shading alone, no highlight.
  const palette = Object.fromEntries(HIGHLIGHT_COLORS.map(([name, hex]) => [name, '#' + hex]));
  assert.deepEqual(palette, Object.fromEntries(check61.reads.filter((x) => x.highlightColor !== null).map((x) => [x.name, x.highlightColor])), "HIGHLIGHT_COLORS is Word's palette");
  for (const { name, highlightColor } of check61.reads) {
    const rPr = name === 'shd' ? { shd: { val: 'clear', color: 'auto', fill: 'FFFF00' } } : { highlight: { val: name } };
    assert.equal(highlightHexValue(name === 'shd' ? undefined : name) ?? null, highlightColor, `highlightHexValue ${name}`);
    assert.equal(readRunOptions(rPr).highlightColor, highlightColor, `read ${name}`);
  }

  // Writes: what Word stores and reads back for each value Office JS was given; refused as Office JS refuses; null removes.
  for (const w of check61.writes) {
    const before = { highlight: { val: 'red' } };
    if (w.value === null) {
      assert.deepEqual(applyRunOptions(before, { highlightColor: null }), {}, 'null removes');
    } else if (w.set !== undefined) {
      assert.match(w.set, /InvalidArgument/);
      assert.equal(highlightNameForColor(w.value), undefined, w.value);
      assert.throws(() => applyRunOptions(before, { highlightColor: w.value }), /Not a highlight colour/, w.value);
    } else {
      assert.equal(highlightNameForColor(w.value), w.stored.highlight, `stored for ${w.value}`);
      const rPr = applyRunOptions(before, { highlightColor: w.value });
      assert.deepEqual(rPr, { highlight: { val: w.stored.highlight } }, `written for ${w.value}`);
      assert.equal(readRunOptions(rPr).highlightColor, w.highlightColor, `read back for ${w.value}`);
    }
    assert.equal(w.stored.shd, null, 'never a shading');
  }

  // The snap (H4 (c)): every colour of check 62 is stored as Word stored it, the hex in either case.
  const nameOf = Object.fromEntries(HIGHLIGHT_COLORS.map(([name, hex]) => [hex, name]));
  let stored = 0;
  for (const [hex, read, extra] of check62.points) {
    const word = nameOf[read];
    assert.ok(word !== undefined && (extra === undefined || extra.stored === word), `${hex}: Word's read and store agree`);
    if (extra !== undefined) stored++;
    assert.equal(highlightNameForColor('#' + hex), word, `#${hex}`);
    assert.equal(highlightNameForColor('#' + hex.toLowerCase()), word, `#${hex.toLowerCase()}`);
  }
  assert.equal(stored, 75, 'the stored names spot-checked through getOoxml');
  // Every palette value round-trips through the run mapping.
  for (const [name, hex] of HIGHLIGHT_COLORS) assert.equal(readRunOptions(applyRunOptions({}, { highlightColor: '#' + hex.toLowerCase() })).highlightColor, '#' + hex, name);
  // A colour check 62 didn't measure follows the rule: the hue sectors, the dark bound and the grey bands.
  const measured = new Set(check62.points.map(([hex]) => hex));
  for (const [hex, name] of [['FF0012', 'magenta'], ['C00001', 'darkRed'], ['C10001', 'red'], ['3F3F40', 'black'], ['404041', 'darkGray'], ['A0A0A1', 'lightGray'], ['E0E0E1', 'white'], ['FFD701', 'yellow'], ['123457', 'darkCyan']]) {
    assert.ok(!measured.has(hex), `${hex} is unmeasured`);
    assert.equal(highlightNameForColor('#' + hex), name, `#${hex}`);
  }

  // Names (H3): CSS Color 3's 147, in any case, each stored as its colour is; the seven spellings Word's w:color refuses too.
  const extra = { darkgrey: 'darkgray', darkslategrey: 'darkslategray', dimgrey: 'dimgray', grey: 'gray', lightgray: 'lightgrey', lightslategrey: 'lightslategray', slategrey: 'slategray' };
  const css = [...Object.keys(WORD_COLOR_NAMES), ...Object.keys(extra)];
  assert.equal(new Set(css).size, 147);
  for (const name of css) {
    const stored = highlightNameForColor('#' + WORD_COLOR_NAMES[extra[name] ?? name]);
    for (const spelling of [name, name.toUpperCase(), name[0].toUpperCase() + name.slice(1)]) assert.equal(highlightNameForColor(spelling), stored, spelling);
  }
  for (const [name, stored] of [['green', 'darkGreen'], ['lime', 'green'], ['olive', 'darkYellow'], ['navy', 'darkBlue'], ['teal', 'darkCyan'], ['silver', 'lightGray'], ['gray', 'darkGray'], ['darkGray', 'darkGray'], ['lightGray', 'lightGray'], ['gold', 'yellow'], ['DarkMagenta', 'darkMagenta']]) {
    assert.equal(highlightNameForColor(name), stored, name);
  }
  // Refused: no CSS name (a highlight name CSS lacks; CSS 4's), a hex without its #, three digits, other forms; a padded name.
  for (const value of ['darkYellow', 'none', 'rebeccapurple', 'transparent', 'FFFF00', '#FF0', '#FFFF000', 'rgb(255, 255, 0)', '#GGGGGG', ' yellow', '', 'constructor', '__proto__']) {
    assert.equal(highlightNameForColor(value), undefined, JSON.stringify(value));
  }
  assert.equal(highlightNameForColor(' #ffff00 '), 'yellow', 'a hex is trimmed');
  assert.equal(highlightNameForColor(undefined), undefined);
  assert.deepEqual(applyRunOptions({ highlight: { val: 'red' } }, { highlightColor: '' }), {}, "'' removes");
  console.log("CR-014: Word's palette read (check 61's 18 reads), Office JS's input written (its 45 writes), the snap on check 62's 7,225 colours, CSS's 147 names OK");
}

// CR-015: a decimal in an integer attribute of Word's own mappings, read as Word 365 reads it: truncated toward zero; a
// form Word refuses the whole document for, dropped. test/word365-integers.mjs holds the readings to Word's own saves
// of docx4j CR-065's probe; here, the scope, the record, a caller's precedence and the save. Phase 2 passes jsonix-CR-012's
// strictIntegers (@docx4j/jsonix 3.9.0), so the forms the runtime's Number() would read reach the same rule.
{
  const { wordInvalidValue } = await import('../dist/index.mjs');
  const W_ = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const WP_ = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const WORD = new Set(['org_docx4j_wml', 'org_docx4j_w14', 'org_docx4j_w15', 'org_docx4j_w15symex', 'org_docx4j_w16', 'org_docx4j_w16cex', 'org_docx4j_w16cid']);
  const INTEGERS = new Set(['Integer', 'Long', 'Int', 'Short', 'Byte', 'NonNegativeInteger', 'NonPositiveInteger', 'NegativeInteger', 'PositiveInteger', 'UnsignedLong', 'UnsignedInt', 'UnsignedShort', 'UnsignedByte']);
  const context = await getContext();
  // The scope: every integer attribute of Word's own mappings, keyed by type; every other mapping's reads as before
  // strictIntegers, through the runtime's own parse, or fails with its error.
  let word = 0;
  let others = 0;
  for (const classInfo of Object.values(context.typeInfos)) {
    if (classInfo.CLASS_NAME !== 'Jsonix.Model.ClassInfo' || !classInfo.module) continue;
    for (const property of classInfo.properties ?? []) {
      if (property.CLASS_NAME !== 'Jsonix.Model.AttributePropertyInfo' || !INTEGERS.has(property.typeInfo.name)) continue;
      const error = new Error('refused');
      if (!WORD.has(classInfo.module.name)) {
        others++;
        assert.throws(() => wordInvalidValue('7.9', property, classInfo, null, error), /Argument \[7\.9\] must be an integer\./, `${classInfo.name}.${property.name}`);
        assert.deepEqual(['0x10', '', ' 240 ', '1e3'].map((t) => wordInvalidValue(t, property, classInfo, null, error)), [16, 0, 240, 1000], 'as before strictIntegers');
        continue;
      }
      word++;
      const read = (text) => wordInvalidValue(text, property, classInfo, null, error);
      assert.deepEqual([read('7.9'), read('-7.9'), read('+240.5'), read('.5'), read('240.')], [7, -7, 240, 0, 240], `${classInfo.name}.${property.name}`);
      assert.ok(Object.is(read('-0.5'), 0), 'never -0');
      for (const refused of [' 7.9 ', '7.9 ', '7,9', '2.5e0', '0x1.8', 'abc', 'Infinity', 'NaN', ' 240 ', '\t240', '0x10', '1e3', '1.5e2', '. ', '++']) assert.equal(read(refused), undefined, JSON.stringify(refused));
      assert.deepEqual(['', '.', '+', '-', '+.', '10206.0'].map(read), [0, 0, 0, 0, 0, 10206], 'the forms with no digit as 0 (round 6, group E), and a decimal with no fraction');
    }
  }
  assert.equal(word, 178, "178 integer attributes on Word's own mappings");
  assert.ok(others > 0);
  assert.equal(wordInvalidValue('141.99999999999994', { attributeName: {}, typeInfo: { name: 'Integer' } }, { module: { name: 'org_docx4j_wml' } }, null, new Error()), 141, 'the integer part, not the nearest');

  // The record: a truncated value and a refused form, each with its kind; a caller's own onInvalidValue takes precedence.
  const doc = (pPr, attrs = '') => `<w:document xmlns:w="${W_}" xmlns:wp="${WP_}"><w:body><w:p${attrs}><w:pPr>${pPr}</w:pPr><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>`;
  const readings = [];
  const read = unwrap(await unmarshalString(doc('<w:spacing w:before="120.5" w:after="240,5" w:line="259.2"/><w:ind w:left="-720.5"/>'), { onWordReading: (r) => readings.push(r) }));
  const pPr = read.body.content[0].value.pPr;
  assert.deepEqual([pPr.spacing.before, pPr.spacing.after, pPr.spacing.line, pPr.ind.left], [120, undefined, 259, -720]);
  assert.deepEqual(readings.map((r) => [r.element.localName, r.attribute.localPart, r.text, r.value, r.kind]), [
    ['spacing', 'after', '240,5', undefined, 'value.refused'],
    ['spacing', 'before', '120.5', 120, 'value.truncated'],
    ['spacing', 'line', '259.2', 259, 'value.truncated'],
    ['ind', 'left', '-720.5', -720, 'value.truncated'],
  ].sort((a, b) => readings.findIndex((r) => r.text === a[2]) - readings.findIndex((r) => r.text === b[2])));
  assert.equal(unwrap(await unmarshalString(doc('<w:spacing w:before="120.5"/>'))).body.content[0].value.pPr.spacing.before, 120, 'without onWordReading too');
  // Phase 2: what the runtime's Number() would read reaches Word's rule, and is reported.
  const strict = [];
  const strictRead = unwrap(await unmarshalString(doc('<w:spacing w:before=" 240 " w:after="" w:line="10206.0"/><w:ind w:left="0x10" w:right="+240"/>'), { onWordReading: (r) => strict.push(r) }));
  const strictPPr = strictRead.body.content[0].value.pPr;
  assert.deepEqual([strictPPr.spacing.before, strictPPr.spacing.after, strictPPr.spacing.line, strictPPr.ind.left, strictPPr.ind.right], [undefined, 0, 10206, undefined, 240]);
  assert.deepEqual(strict.map((r) => [r.attribute.localPart, r.text, r.value, r.kind]).sort(), [
    ['after', '', 0, 'value.kept.invalid'], ['before', ' 240 ', undefined, 'value.refused'],
    ['left', '0x10', undefined, 'value.refused'], ['line', '10206.0', 10206, 'value.truncated'],
  ]);
  // strictIntegers: false asks for the runtime's reading; a caller's own onInvalidValue keeps it unless it asks too.
  const lenient = unwrap(await unmarshalString(doc('<w:spacing w:before=" 240 " w:line="259.2"/>'), { strictIntegers: false })).body.content[0].value.pPr.spacing;
  assert.deepEqual([lenient.before, lenient.line], [240, 259]);
  const callers = [];
  const theirs = unwrap(await unmarshalString(doc('<w:spacing w:before=" 240 " w:line="259.2"/>'), { onInvalidValue: (text) => { callers.push(text); return 1; } })).body.content[0].value.pPr.spacing;
  assert.deepEqual([theirs.before, theirs.line, callers], [240, 1, ['259.2']], "a caller's own onInvalidValue: today's reading");
  callers.length = 0;
  await unmarshalString(doc('<w:spacing w:before=" 240 "/>'), { onInvalidValue: (text) => { callers.push(text); return 1; }, strictIntegers: true });
  assert.deepEqual(callers, [' 240 '], 'and strictIntegers when it asks');
  const own = unwrap(await unmarshalString(doc('<w:spacing w:before="120.5"/>'), { onInvalidValue: () => 121, onWordReading: () => assert.fail('not called beside its own') }));
  assert.equal(own.body.content[0].value.pPr.spacing.before, 121, "a caller's own onInvalidValue");
  // Other mappings fail as before: DrawingML's extent.
  await assert.rejects(unmarshalString(`<w:document xmlns:w="${W_}" xmlns:wp="${WP_}"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="12.5" cy="10"/></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`), /must be an integer/);
  // The save writes the integer; a caller's own non-integer still throws (N6), with the normaliser or without.
  assert.ok((await marshalString(await unmarshalString(doc('<w:spacing w:line="259.2"/>')))).includes('<w:spacing w:line="259"/>'));
  const set = await unmarshalString(doc('<w:spacing w:line="240"/>'));
  unwrap(set).body.content[0].value.pPr.spacing.line = 259.2;
  await assert.rejects(marshalString(set), /must be an integer/);
  await assert.rejects(marshalString(set, { normaliseOnSave: false }), /must be an integer/);

  // N7: an empty rsid is left out on save, as Word's save leaves it out, and reported; the tree keeps it; the opt-out writes it.
  const ssrs = await unmarshalString(doc('', ' w:rsidR="" w:rsidRDefault="00AB12CD" w:rsidP=""'));
  const normalised = [];
  const saved = await marshalString(ssrs, { onNormalised: (n) => normalised.push(n) });
  assert.ok(saved.includes('<w:p w:rsidRDefault="00AB12CD">'), 'empty rsids left out, a non-empty one kept');
  assert.deepEqual(normalised.map((n) => [n.path, n.before, n.after]), [
    ['/w:document/w:body[1]/w:p[1]/@w:rsidP', '', null], ['/w:document/w:body[1]/w:p[1]/@w:rsidR', '', null],
  ]);
  assert.equal(unwrap(ssrs).body.content[0].value.rsidR, '', 'the tree keeps what it read');
  assert.ok((await marshalString(ssrs, { normaliseOnSave: false })).includes('w:rsidR=""'), 'the opt-out writes as read');
  console.log("CR-015: 178 integer attributes of Word's own mappings read as Word does (truncated; a refused form dropped), other mappings as before; the record, a caller's precedence; the save writes integers and leaves out empty rsids OK");
}

// CR-016: integer attributes within Word's range. test/word365-ranges.mjs holds wordIntegerReading and onOutOfRange to
// Word's saves of docx4j CR-066's round 5; here, the table against the mapping, the channel and the builders' guard.
{
  const { WORD_INTEGER_STORAGE, wordIntegerReading } = await import('../dist/helpers/wml.mjs');
  const { applyRunOptions, tbl, tc, tr } = await import('../dist/builders/wml.mjs');
  const W_ = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const INTEGERS = new Set(['Integer', 'Long', 'Int', 'Short', 'Byte', 'NonNegativeInteger', 'NonPositiveInteger', 'NegativeInteger', 'PositiveInteger', 'UnsignedLong', 'UnsignedInt', 'UnsignedShort', 'UnsignedByte']);
  const context = await getContext();
  // The table names wml's own attributes: each an integer attribute of its class, in w's namespace, by its local name.
  let entries = 0;
  let unmeasured = 0;
  for (const [typeName, properties] of Object.entries(WORD_INTEGER_STORAGE)) {
    const classInfo = context.getTypeInfoByName(typeName);
    for (const [name, storage] of Object.entries(properties)) {
      const property = classInfo.properties.find((p) => p.name === name);
      assert.ok(property && property.CLASS_NAME === 'Jsonix.Model.AttributePropertyInfo' && INTEGERS.has(property.typeInfo.name), `${typeName}.${name}`);
      assert.deepEqual([property.attributeName.namespaceURI, property.attributeName.localPart], [W_, name]);
      entries++;
      if (!storage.measured) unmeasured++;
    }
  }
  assert.deepEqual([entries, unmeasured], [51, 4], '44 twips attributes (30 ST_TwipsMeasure, 14 signed) and 7 others; 4 unmeasured after round 6 (gridCol, legacySpace, dxaOrig, dyaOrig)');
  assert.equal(wordIntegerReading('org_docx4j_wml.Lvl.Start', 'val', 70000), undefined, 'not in the table');
  // The reading: wrapped, clamped, the rule in the sign; an unmeasured attribute only outside -32768..65535.
  assert.deepEqual(wordIntegerReading('org_docx4j_wml.PPrBase.Ind', 'left', 7205001), { value: -3959, written: -3959, kept: false, storage: 'int16' });
  assert.deepEqual(wordIntegerReading('org_docx4j_wml.SectPr.PgSz', 'w', 40000), { value: 31681, written: 31681, kept: false, storage: 'uint16, then 144..31681' });
  assert.deepEqual(wordIntegerReading('org_docx4j_wml.PPrBase.Spacing', 'line', 65535, 'exact'), { value: 1, written: 1, rule: 'atLeast', kept: false, storage: 'int16, the sign the rule' });
  assert.equal(wordIntegerReading('org_docx4j_wml.PPrBase.Spacing', 'before', -1).kept, true, "before's -1 is Word's own 65535");
  assert.deepEqual(wordIntegerReading('org_docx4j_wml.TblGridCol', 'w', 70000), { value: null, written: null, kept: false, storage: 'unmeasured: changed outside -32768..65535' });
  assert.equal(wordIntegerReading('org_docx4j_wml.TblGridCol', 'w', 40000).kept, true, 'unmeasured, inside the certain range');
  // Round 6: firstLine and hanging share one field; a frame's height keeps 15 bits; a column's negative width is computed.
  assert.deepEqual(wordIntegerReading('org_docx4j_wml.PPrBase.Ind', 'firstLine', 32768), { value: 32768, written: 32768, attribute: 'hanging', kept: false, storage: 'int16, shared by firstLine (+) and hanging (-)' });
  assert.equal(wordIntegerReading('org_docx4j_wml.PPrBase.Ind', 'hanging', 32768).kept, true);
  assert.deepEqual(wordIntegerReading('org_docx4j_wml.CTFramePr', 'h', 32768, 'exact'), { value: 0, written: 0, rule: 'auto', kept: false, storage: '15 bits, bit 15 clearing exact' });
  assert.deepEqual(wordIntegerReading('org_docx4j_wml.CTColumn', 'w', -1), { value: null, written: null, kept: false, storage: 'int16, a negative computed' });
  assert.equal(wordIntegerReading('org_docx4j_wml.CTHeight', 'val', 7205001, 'auto').kept, true, 'a row height under auto is ignored');

  // The channel: on read, the tree keeps the value; on save, written as held, with or without the normaliser.
  const doc = `<w:document xmlns:w="${W_}"><w:body><w:p><w:pPr><w:spacing w:before="240" w:after="70000" w:line="65535" w:lineRule="exact"/><w:ind w:left="7205001"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>`;
  const read = [];
  const tree = await unmarshalString(doc, { onOutOfRange: (r) => read.push(r) });
  assert.deepEqual(read.map((r) => [r.element.localName, r.attribute.localPart, r.value, r.wordReads.value, r.wordReads.rule]).sort(), [
    ['ind', 'left', 7205001, -3959, undefined], ['spacing', 'after', 70000, 4464, undefined], ['spacing', 'line', 65535, 1, 'atLeast'],
  ]);
  const pPr = unwrap(tree).body.content[0].value.pPr;
  assert.deepEqual([pPr.ind.left, pPr.spacing.after, pPr.spacing.line], [7205001, 70000, 65535], 'held as written');
  pPr.spacing.before = 70000; // set by code
  for (const normaliseOnSave of [true, false]) {
    const saved = [];
    const xml = await marshalString(tree, { normaliseOnSave, onOutOfRange: (r) => saved.push([r.path, r.value]) });
    assert.ok(xml.includes('w:left="7205001"') && xml.includes('w:before="70000"'), 'never rewritten');
    assert.deepEqual(saved.sort(), [
      ['/w:document/w:body[1]/w:p[1]/w:pPr[1]/w:ind[1]/@w:left', 7205001],
      ['/w:document/w:body[1]/w:p[1]/w:pPr[1]/w:spacing[1]/@w:after', 70000],
      ['/w:document/w:body[1]/w:p[1]/w:pPr[1]/w:spacing[1]/@w:before', 70000],
      ['/w:document/w:body[1]/w:p[1]/w:pPr[1]/w:spacing[1]/@w:line', 65535],
    ], `on save, normaliseOnSave ${normaliseOnSave}`);
  }
  assert.ok((await marshalString(tree)).includes('w:left="7205001"'), 'not passed: nothing reported, nothing changed');

  // The builders refuse a measure Word can't hold (G4), and take the boundaries.
  assert.equal(tc('a', { width: 32767 }).value.tcPr.tcW.w, 32767);
  for (const width of [32768, -1, NaN, Infinity]) assert.throws(() => tc('a', { width }), RangeError, `cell width ${width}`);
  assert.throws(() => tr(['a', 'b'], { widths: [100, 40000] }), /Cell width 40000 is outside Word's range: a table or cell width is stored in 16 bits, 0 to 32767 twips/);
  assert.throws(() => tbl([['a']], { width: 40000 }), /Table width 40000/);
  assert.throws(() => tbl([['a', 'b']], { widths: [20000, 20000] }), /Table width \(the columns' total\) 40000/);
  assert.equal(tbl([['a', 'b']], { widths: [20000, 12767] }).value.tblPr.tblW.w, 32767);
  assert.equal(applyRunOptions({}, { size: 32767.5 }).sz.val, 65535);
  for (const size of [32768, -1, NaN, Infinity]) assert.throws(() => applyRunOptions({}, { size }), /Font size .* is outside Word's range: w:sz is stored in 16 bits, 0 to 32767.5 points/, `size ${size}`);
  console.log("CR-016: Word's integer storage (51 attributes, 4 at the certain range) read as rounds 5 and 6 measured; onOutOfRange on read and on save, nothing rewritten; the builders' widths and size within Word's range OK");
}
