// Runtime check of CR-016 against Word 365: docx4j CR-066's round 5 of the shared probe (Jason, 2026-10-26), integers
// at and beyond each attribute's storage width. test/fixtures/cr066-probe/ holds the six group files and Word's saves of
// them, with a SOURCE.md. Word's save is the oracle: it writes Word's reading of each value. So for each case,
// helpers/wml's wordIntegerReading must give what Word's save wrote, and the facade's onOutOfRange must report exactly
// the cases Word reads differently, on read and on save.
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { unmarshalString, marshalString } from '../dist/index.mjs';
import { wordIntegerReading } from '../dist/helpers/wml.mjs';

const DIR = new URL('./fixtures/cr066-probe/', import.meta.url);
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const GROUPS = ['P', 'B', 'R', 'T', 'S', 'N'];
const part = (file) => readFileSync(new URL(`${file}.docx/word/document.xml`, DIR), 'utf8');
const CASES = readFileSync(new URL('cases.txt', DIR), 'utf8').trim().split('\n').slice(1).map((l) => l.split('\t'))
  .filter(([, group]) => group !== 'Z').map(([id, group, attribute, value]) => ({ id, group, attribute, value: Number(value) }));
assert.equal(CASES.length, 190);

/** Each case's attribute: its class and property in this package, the rule beside it, and where it is in the part. */
const ATTRIBUTES = {
  'w:spacing/@w:before': ['org_docx4j_wml.PPrBase.Spacing', 'before', ['pPr', 'spacing']],
  'w:ind/@w:left': ['org_docx4j_wml.PPrBase.Ind', 'left', ['ind']],
  'w:spacing/@w:line[auto]': ['org_docx4j_wml.PPrBase.Spacing', 'line', ['pPr', 'spacing'], 'lineRule'],
  'w:spacing/@w:line[exact]': ['org_docx4j_wml.PPrBase.Spacing', 'line', ['pPr', 'spacing'], 'lineRule'],
  'w:tab/@w:pos': ['org_docx4j_wml.CTTabStop', 'pos', ['tab']],
  'w:pBdr/w:top/@w:sz': ['org_docx4j_wml.CTBorder', 'sz', ['pBdr', 'top']],
  'w:pBdr/w:top/@w:space': ['org_docx4j_wml.CTBorder', 'space', ['pBdr', 'top']],
  'w:sz/@w:val': ['org_docx4j_wml.HpsMeasure', 'val', ['rPr', 'sz']],
  'w:position/@w:val': ['org_docx4j_wml.CTSignedHpsMeasure', 'val', ['rPr', 'position']],
  'w:spacing/@w:val': ['org_docx4j_wml.CTSignedTwipsMeasure', 'val', ['rPr', 'spacing']],
  'w:w/@w:val': ['org_docx4j_wml.CTTextScale', 'val', ['rPr', 'w']],
  'w:tblW/@w:w[dxa]': ['org_docx4j_wml.TblWidth', 'w', ['tblW']], 'w:tblW/@w:w[pct]': ['org_docx4j_wml.TblWidth', 'w', ['tblW']],
  'w:tcW/@w:w[dxa]': ['org_docx4j_wml.TblWidth', 'w', ['tcW']], 'w:tcW/@w:w[pct]': ['org_docx4j_wml.TblWidth', 'w', ['tcW']],
  'w:tblInd/@w:w[dxa]': ['org_docx4j_wml.TblWidth', 'w', ['tblInd']],
  'w:trHeight/@w:val': ['org_docx4j_wml.CTHeight', 'val', ['trHeight'], 'hRule'],
  'w:pgSz/@w:w': ['org_docx4j_wml.SectPr.PgSz', 'w', ['pgSz']],
  'w:pgSz/@w:h': ['org_docx4j_wml.SectPr.PgSz', 'h', ['pgSz']],
  'w:pgMar/@w:top': ['org_docx4j_wml.SectPr.PgMar', 'top', ['pgMar']],
  'w:pgNumType/@w:start': ['org_docx4j_wml.CTPageNumber', 'start', ['pgNumType']],
};

const textOf = (p) => Array.from(p.getElementsByTagNameNS(W, 't')).map((t) => t.textContent).join('');
const nextElement = (n) => { for (let x = n.nextSibling; x; x = x.nextSibling) if (x.nodeType === 1) return x; return undefined; };
const previousElement = (n) => { for (let x = n.previousSibling; x; x = x.previousSibling) if (x.nodeType === 1) return x; return undefined; };
const holdsSection = (p) => p.localName === 'p' && p.getElementsByTagNameNS(W, 'sectPr').length > 0 && textOf(p) === '';

/** The case's block: its label paragraph and, where the value is in it, the table or the section-ending paragraph after. */
function blockOf(doc, id) {
  const p = Array.from(doc.getElementsByTagNameNS(W, 'p')).find((x) => textOf(x).startsWith(`${id} `));
  if (!p) return undefined;
  const next = nextElement(p);
  return next && (next.localName === 'tbl' || holdsSection(next)) ? [p, next] : [p];
}

/** The label of the case an element belongs to: the body's child holding it, or the label paragraph before that. */
function caseOf(element) {
  let child = element;
  while (child.parentNode && child.parentNode.localName !== 'body') child = child.parentNode;
  const label = child.localName === 'tbl' || holdsSection(child) ? previousElement(child) : child;
  return label ? textOf(label).split(' ')[0] : undefined;
}

/** What Word's save wrote for the case: the attribute in the block (its first value), with the rule beside it. */
function savedOf(doc, c) {
  const [, , steps, ruleAttribute] = ATTRIBUTES[c.attribute];
  const attribute = c.attribute.match(/@w:(\w+)/)[1];
  for (const root of blockOf(doc, c.id)) {
    for (const e of Array.from(root.getElementsByTagNameNS(W, steps.at(-1)))) {
      if (steps.length > 1 && e.parentNode.localName !== steps.at(-2)) continue;
      if (e.hasAttributeNS(W, attribute)) return { value: Number(e.getAttributeNS(W, attribute)), rule: ruleAttribute ? (e.getAttributeNS(W, ruleAttribute) || undefined) : undefined };
    }
  }
  return undefined;
}

const parse = (xml) => new DOMParser().parseFromString(xml, 'text/xml');
const known = [];
const changed = new Set();
for (const group of GROUPS) {
  const word = parse(part(`cr066-${group}-word365`));
  for (const c of CASES.filter((x) => x.group === group)) {
    const [typeName, property, , ruleAttribute] = ATTRIBUTES[c.attribute];
    const writtenRule = c.attribute.match(/\[(auto|exact)\]/)?.[1];
    const reading = wordIntegerReading(typeName, property, c.value, writtenRule);
    assert.ok(reading, `${c.id}: in the table`);
    const saved = savedOf(word, c);
    if (saved === undefined) {
      // Word leaves out a value that wraps to the default: tblInd and trHeight 65536 read as 0.
      assert.equal(reading.written, 0, `${c.id}: Word left it out, so it read the default`);
      known.push(c.id);
    } else {
      assert.equal(reading.written, saved.value, `${c.id} ${c.attribute}=${c.value}: Word's save wrote ${saved.value}`);
      if (ruleAttribute === 'lineRule') assert.equal(reading.rule, saved.rule, `${c.id}: the line rule`);
      if (ruleAttribute === 'hRule') assert.equal(reading.rule === 'exact' ? 'exact' : undefined, saved.rule, `${c.id}: the row rule`);
    }
    if (!reading.kept) changed.add(c.id);
    // Word reads the value as written exactly when its save writes it back, but for before's 65535, which it writes -1.
    const sameInSave = saved !== undefined && saved.value === c.value && (writtenRule === undefined || saved.rule === writtenRule) && (ruleAttribute !== 'hRule' || saved.rule === undefined);
    assert.equal(reading.kept, sameInSave || (c.id === 'P1-3'), `${c.id}: kept`);
  }
}
assert.deepEqual(known, ['T3-4', 'T6-4']);

// The facade: onOutOfRange on read reports exactly the cases Word reads differently, and on save the same, unchanged.
const readReported = new Set();
const savedReported = new Set();
for (const group of GROUPS) {
  const readReports = [];
  const tree = await unmarshalString(part(`cr066-${group}`), { onOutOfRange: (r) => readReports.push(r) });
  for (const r of readReports) {
    const id = caseOf(r.element);
    assert.ok(CASES.some((c) => c.id === id && c.value === r.value), `read: ${r.attribute.localPart}=${r.value} belongs to case ${id}`);
    assert.equal(r.wordReads.kept, false);
    readReported.add(id);
  }
  const saveReports = [];
  const xml = await marshalString(tree, { onOutOfRange: (r) => saveReports.push(r) });
  for (const r of saveReports) {
    assert.ok(r.path.startsWith('/w:document/w:body[1]/') && r.path.endsWith(`/@w:${r.attribute.localPart}`), r.path);
    savedReported.add(caseOf(r.element));
  }
  assert.equal(saveReports.length, readReports.length, `${group}: the same values on save`);
  // Never rewritten: every value is written as the probe holds it.
  for (const c of CASES.filter((x) => x.group === group)) assert.equal(savedOf(parse(xml), c).value, c.value, `${c.id}: written as held`);
}
assert.deepEqual([...readReported].sort(), [...changed].sort(), 'read: reported exactly where Word reads differently');
assert.deepEqual([...savedReported].sort(), [...changed].sort(), 'save: the same cases');

console.log(`word365 ranges (CR-016): ${CASES.length} cases of docx4j CR-066's round 5 in ${GROUPS.length} files, each read by wordIntegerReading as Word's save writes it `
  + `(${known.length} wrapped to a default Word leaves out); ${changed.size} reported by onOutOfRange on read and on save, the rest not; nothing rewritten OK`);
