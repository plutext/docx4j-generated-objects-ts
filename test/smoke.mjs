// Runtime check of the facade against a small WordprocessingML document.
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';
import { unmarshalString, marshalString, unwrap, deepCopy, getContext, Jsonix } from '../dist/index.mjs';
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
assert.equal(unwrap(run.content[0]).value, 'Hello, docx4j-ts');
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
assert.equal(require('@docx4j/docx4j-ts/modules/org_docx4j_wml').org_docx4j_wml.name, 'org_docx4j_wml');
const { org_docx4j_wml } = await import('@docx4j/docx4j-ts/modules/org_docx4j_wml');
assert.equal(org_docx4j_wml.name, 'org_docx4j_wml');
console.log('docx4j-ts smoke: unmarshal, parent pointers, deepCopy, marshal, v:line order OK');
