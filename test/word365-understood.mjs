// Runtime check of CR-021 against Word 365: the rounds behind docx4j's understood set (test/fixtures/understood-probe/, with a
// SOURCE.md), each a Choice-only file beside the text Word saved from it. Word's text is the oracle: textOf must give it, and
// mcBranchOf must report nothing selected exactly where Word showed neither branch.
import { readdirSync, readFileSync } from 'node:fs';
import assert from 'node:assert/strict';
import { unmarshalString, unwrap } from '../dist/index.mjs';
import { mcBranchOf, textOf, walkAll } from '../dist/builders/wml.mjs';

const dir = new URL('fixtures/understood-probe/', import.meta.url);
const ids = readdirSync(dir).filter((name) => name.endsWith('.document.xml')).map((name) => name.replace('.document.xml', '')).sort();
assert.equal(ids.length, 142, 'the rounds: 119 ul, 13 uc, 9 c305, rw-25');

const linesOf = (text) => text.split(/\r?\n|\r|\f/).map((line) => line.trim()).filter(Boolean);
const alternateContentIn = (root) => {
  const found = [];
  walkAll(root, (value) => { if (value.TYPE_NAME === 'org_docx4j_mce.AlternateContent') found.push(value); }, () => {});
  return found;
};

let shown = 0;
let neither = 0;
for (const id of ids) {
  const word = linesOf(readFileSync(new URL(`${id}.word.txt`, dir), 'latin1'));
  const document = await unmarshalString(readFileSync(new URL(`${id}.document.xml`, dir), 'utf8'));
  assert.deepEqual(linesOf(textOf(unwrap(document).body)), word, `${id}: textOf gives Word's text`);

  const wordShowsChoice = word.some((line) => line.includes('CHOICE WORDS'));
  const reports = [];
  const elements = alternateContentIn(unwrap(document));
  assert.equal(elements.length, 1, `${id}: one element`);
  const branch = mcBranchOf(elements[0], { onNothingSelected: (report) => reports.push(report) });
  assert.equal(branch !== undefined, wordShowsChoice, `${id}: a branch where Word shows the Choice`);
  assert.equal(reports.length, wordShowsChoice ? 0 : 1, `${id}: reported where Word shows neither`);
  if (wordShowsChoice) shown++;
  else { neither++; assert.ok(reports[0].namespaces.length > 0, `${id}: the report names the namespace`); }
}

console.log(`word365 understood (CR-021): ${ids.length} files, textOf equal to Word's text in each; Word showed the Choice in ${shown}, neither branch in ${neither}, each of those reported`);
