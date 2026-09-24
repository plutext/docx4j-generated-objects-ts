// The CommonJS half of the bundle check. This package's facade is import-only, so a require()
// consumer uses the UMD mappings (`./modules/*` carries a "require" condition) and builds its own
// context - the path the smoke test's require() checks describe.
//
// It earns its place because a bundler picks the runtime's entry from the CONSUMER'S SYNTAX, not
// the output format (jsonix-CR-005 section 8): an `import` consumer bundled to cjs still gets the
// Node ESM entry, so only a require() consumer reaches jsonix.js's UMD footer and its literal
// requires. Without this witness both bundles would exercise the same entry.
//
// org_docx4j_relationships is the whole context on purpose: its module closure is itself (nothing
// else is referenced), so this stays a check of the runtime's DOM and entry resolution rather than
// a hand-picked module list, which is what modulesFor exists to avoid and is unreachable from CJS.
const { Jsonix } = require('@docx4j/jsonix');
const { org_docx4j_relationships } = require('@docx4j/generated-objects-ts/modules/org_docx4j_relationships');

const context = new Jsonix.Context([org_docx4j_relationships]);
const element = context.createUnmarshaller().unmarshalString(
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + '<Relationship Id="rId1" Type="http://x" Target="word/document.xml"/></Relationships>');
const out = context.createMarshaller().marshalString(element);
console.log(`${element.value.TYPE_NAME} ${element.value.relationship.length} ${out.includes('rId1')}`);
