// Runtime check of CR-016 against Word 365: docx4j CR-066's rounds 5 and 6 of the shared probe (Jason, 2026-10-26),
// integers at and beyond each attribute's storage width. test/fixtures/cr066-probe/ (round 5) and cr066-round6/ hold
// the probe files and Word's saves of them, with a SOURCE.md each. Word's save is the oracle: it writes Word's reading of
// each value. So for each case, helpers/wml's wordIntegerReading must give what Word's save wrote, and the facade's
// onOutOfRange must report exactly the cases Word reads differently, on read and on save. Round 6's group E (the forms
// with no digit, CR-015) rides on the same files.
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { unmarshalString, marshalString } from '../dist/index.mjs';
import { wordIntegerReading, WORD_INTEGER_STORAGE } from '../dist/helpers/wml.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const S = 'org_docx4j_wml.';
const ROUNDS = [
  { dir: 'cr066-probe', prefix: 'cr066-', groups: ['P', 'B', 'R', 'T', 'S', 'N'], cases: 190 },
  { dir: 'cr066-round6', prefix: 'cr066-r6-', groups: ['A', 'F', 'R', 'T', 'P', 'S'], singles: ['D'], cases: 481 },
];

/**
 * Each attribute of the case lists: its class and property here, where it is (element steps from the case's block), the
 * rule beside it, and, for the attributes Word writes back under another name, that name.
 */
const ATTRIBUTES = {
  'w:spacing/@w:before': [S + 'PPrBase.Spacing', 'before', ['pPr', 'spacing']],
  'w:spacing/@w:after': [S + 'PPrBase.Spacing', 'after', ['pPr', 'spacing']],
  'w:spacing/@w:line[auto]': [S + 'PPrBase.Spacing', 'line', ['pPr', 'spacing'], 'lineRule'],
  'w:spacing/@w:line[exact]': [S + 'PPrBase.Spacing', 'line', ['pPr', 'spacing'], 'lineRule'],
  'w:spacing/@w:line[atLeast]': [S + 'PPrBase.Spacing', 'line', ['pPr', 'spacing'], 'lineRule'],
  'w:ind/@w:left': [S + 'PPrBase.Ind', 'left', ['ind']],
  'w:ind/@w:right': [S + 'PPrBase.Ind', 'right', ['ind']],
  'w:ind/@w:start': [S + 'PPrBase.Ind', 'start', ['ind'], undefined, 'left'],
  'w:ind/@w:end': [S + 'PPrBase.Ind', 'end', ['ind'], undefined, 'right'],
  'w:ind/@w:firstLine': [S + 'PPrBase.Ind', 'firstLine', ['ind']],
  'w:ind/@w:hanging': [S + 'PPrBase.Ind', 'hanging', ['ind']],
  'w:tab/@w:pos': [S + 'CTTabStop', 'pos', ['tab']],
  'w:pBdr/w:top/@w:sz': [S + 'CTBorder', 'sz', ['pBdr', 'top']],
  'w:pBdr/w:top/@w:space': [S + 'CTBorder', 'space', ['pBdr', 'top']],
  'w:bdr/@w:sz': [S + 'CTBorder', 'sz', ['rPr', 'bdr']],
  'w:bdr/@w:space': [S + 'CTBorder', 'space', ['rPr', 'bdr']],
  'w:tblBorders/w:top/@w:sz': [S + 'CTBorder', 'sz', ['tblBorders', 'top']],
  'w:tblBorders/w:top/@w:space': [S + 'CTBorder', 'space', ['tblBorders', 'top']],
  'w:tcBorders/w:top/@w:sz': [S + 'CTBorder', 'sz', ['tcBorders', 'top']],
  'w:tcBorders/w:top/@w:space': [S + 'CTBorder', 'space', ['tcBorders', 'top']],
  'w:pgBorders/w:top/@w:sz': [S + 'CTBorder', 'sz', ['pgBorders', 'top']],
  'w:pgBorders/w:top/@w:space': [S + 'CTBorder', 'space', ['pgBorders', 'top']],
  'w:sz/@w:val': [S + 'HpsMeasure', 'val', ['rPr', 'sz']],
  'w:szCs/@w:val': [S + 'HpsMeasure', 'val', ['rPr', 'szCs']],
  'w:kern/@w:val': [S + 'HpsMeasure', 'val', ['rPr', 'kern']],
  'w:position/@w:val': [S + 'CTSignedHpsMeasure', 'val', ['rPr', 'position']],
  'w:spacing/@w:val': [S + 'CTSignedTwipsMeasure', 'val', ['rPr', 'spacing']],
  'w:w/@w:val': [S + 'CTTextScale', 'val', ['rPr', 'w']],
  'w:fitText/@w:val': [S + 'CTFitText', 'val', ['rPr', 'fitText']],
  'w:tblW/@w:w[dxa]': [S + 'TblWidth', 'w', ['tblW']], 'w:tblW/@w:w[pct]': [S + 'TblWidth', 'w', ['tblW']],
  'w:tcW/@w:w[dxa]': [S + 'TblWidth', 'w', ['tcW']], 'w:tcW/@w:w[pct]': [S + 'TblWidth', 'w', ['tcW']],
  'w:tblInd/@w:w[dxa]': [S + 'TblWidth', 'w', ['tblInd']],
  'w:tblCellSpacing/@w:w': [S + 'TblWidth', 'w', ['tblCellSpacing']],
  'w:tblCellMar/w:left/@w:w': [S + 'TblWidth', 'w', ['tblCellMar', 'left']],
  'w:tcMar/w:top/@w:w': [S + 'TblWidth', 'w', ['tcMar', 'top']],
  'w:wBefore/@w:w': [S + 'TblWidth', 'w', ['wBefore']],
  'w:wAfter/@w:w': [S + 'TblWidth', 'w', ['wAfter']],
  'w:gridCol/@w:w': [S + 'TblGridCol', 'w', ['gridCol']],
  'w:trHeight/@w:val': [S + 'CTHeight', 'val', ['trHeight'], 'hRule'],
  'w:trHeight/@w:val[exact]': [S + 'CTHeight', 'val', ['trHeight'], 'hRule'],
  'w:trHeight/@w:val[auto]': [S + 'CTHeight', 'val', ['trHeight'], 'hRule'],
  'w:framePr/@w:w': [S + 'CTFramePr', 'w', ['framePr']], 'w:framePr/@w:h': [S + 'CTFramePr', 'h', ['framePr'], 'hRule'],
  'w:framePr/@w:hSpace': [S + 'CTFramePr', 'hSpace', ['framePr']], 'w:framePr/@w:vSpace': [S + 'CTFramePr', 'vSpace', ['framePr']],
  'w:framePr/@w:x': [S + 'CTFramePr', 'x', ['framePr']], 'w:framePr/@w:y': [S + 'CTFramePr', 'y', ['framePr']],
  'w:tblpPr/@w:leftFromText': [S + 'CTTblPPr', 'leftFromText', ['tblpPr']], 'w:tblpPr/@w:rightFromText': [S + 'CTTblPPr', 'rightFromText', ['tblpPr']],
  'w:tblpPr/@w:topFromText': [S + 'CTTblPPr', 'topFromText', ['tblpPr']], 'w:tblpPr/@w:bottomFromText': [S + 'CTTblPPr', 'bottomFromText', ['tblpPr']],
  'w:tblpPr/@w:tblpX': [S + 'CTTblPPr', 'tblpX', ['tblpPr']], 'w:tblpPr/@w:tblpY': [S + 'CTTblPPr', 'tblpY', ['tblpPr']],
  'w:pgSz/@w:w': [S + 'SectPr.PgSz', 'w', ['pgSz']], 'w:pgSz/@w:h': [S + 'SectPr.PgSz', 'h', ['pgSz']],
  'w:pgMar/@w:top': [S + 'SectPr.PgMar', 'top', ['pgMar']], 'w:pgMar/@w:bottom': [S + 'SectPr.PgMar', 'bottom', ['pgMar']],
  'w:pgMar/@w:left': [S + 'SectPr.PgMar', 'left', ['pgMar']], 'w:pgMar/@w:right': [S + 'SectPr.PgMar', 'right', ['pgMar']],
  'w:pgMar/@w:header': [S + 'SectPr.PgMar', 'header', ['pgMar']], 'w:pgMar/@w:footer': [S + 'SectPr.PgMar', 'footer', ['pgMar']],
  'w:pgMar/@w:gutter': [S + 'SectPr.PgMar', 'gutter', ['pgMar']],
  'w:pgNumType/@w:start': [S + 'CTPageNumber', 'start', ['pgNumType']],
  'w:cols/@w:space': [S + 'CTColumns', 'space', ['cols']],
  'w:cols/w:col/@w:w': [S + 'CTColumn', 'w', ['cols', 'col']],
  'w:cols/w:col/@w:space': [S + 'CTColumn', 'space', ['cols', 'col']],
  'w:lnNumType/@w:distance': [S + 'CTLineNumber', 'distance', ['lnNumType']],
  'w:defaultTabStop/@w:val': [S + 'CTTwipsMeasure', 'val', ['defaultTabStop']],
  'w:hyphenationZone/@w:val': [S + 'CTTwipsMeasure', 'val', ['hyphenationZone']],
};

/** What Word's saves don't let this test hold, each with its reason, strict both ways. */
const NOT_MEASURED = {
  'w:gridCol/@w:w': 'Word recomputes the grid, even under a fixed layout with no cell widths',
  'w:wBefore/@w:w': 'Word dropped it with the w:gridBefore that did not fit the row',
  'w:wAfter/@w:w': 'Word dropped it with the w:gridAfter that did not fit the row',
};

const textOf = (p) => Array.from(p.getElementsByTagNameNS(W, 't')).map((t) => t.textContent).join('');
const nextElement = (n) => { for (let x = n.nextSibling; x; x = x.nextSibling) if (x.nodeType === 1) return x; return undefined; };
const previousElement = (n) => { for (let x = n.previousSibling; x; x = x.previousSibling) if (x.nodeType === 1) return x; return undefined; };
const holdsSection = (p) => p.localName === 'p' && p.getElementsByTagNameNS(W, 'sectPr').length > 0 && textOf(p) === '';
const parse = (xml) => new DOMParser().parseFromString(xml, 'text/xml');

/** The case's block: its label paragraph and, where the value is in it, the table or the section-ending paragraph after. */
function blockOf(doc, id) {
  if (doc.documentElement.localName === 'settings') return [doc.documentElement];
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

/** The first element the steps name in the block (each step the parent of the next). */
function elementIn(block, steps) {
  for (const root of block) {
    for (const e of [root, ...Array.from(root.getElementsByTagNameNS(W, steps.at(-1)))]) {
      if (e.localName !== steps.at(-1)) continue;
      let ok = true;
      for (let i = steps.length - 2, n = e.parentNode; i >= 0; i--, n = n?.parentNode) if (n?.localName !== steps[i]) { ok = false; break; }
      if (ok) return e;
    }
  }
  return undefined;
}

const integerAttribute = (e, name) => (e?.hasAttributeNS(W, name) ? Number(e.getAttributeNS(W, name)) : undefined);

/** Every case of both rounds, with the part it is in and Word's save of that part. */
const CASES = [];
for (const round of ROUNDS) {
  const dir = new URL(`./fixtures/${round.dir}/`, import.meta.url);
  const list = readFileSync(new URL('cases.txt', dir), 'utf8').trim().split('\n').slice(1).map((l) => l.split('\t'));
  for (const [id, group, attribute, ...rest] of list) {
    if (group === 'Z' || group === 'N' && round.dir === 'cr066-round6') continue; // round 6's N: Word dropped w:legacy
    const value = round.dir === 'cr066-probe' ? rest[0] : rest[1]; // round 6 has a place column first
    const file = round.singles?.includes(group) ? `${round.prefix}${id}` : `${round.prefix}${group}`;
    const part = round.singles?.includes(group) ? 'word/settings.xml' : 'word/document.xml';
    CASES.push({ round: round.dir, id, group, attribute, value, file, part, dir });
  }
}
const read = (c, suffix = '') => readFileSync(new URL(`${c.file}${suffix}.docx/${c.part}`, c.dir), 'utf8');
const integerCases = CASES.filter((c) => c.group !== 'E');
assert.deepEqual([CASES.filter((c) => c.round === 'cr066-probe').length, CASES.filter((c) => c.round === 'cr066-round6').length], [190, 478]);

const parsed = new Map();
const docOf = (c, suffix) => { const key = `${c.round} ${c.file}${suffix}`; if (!parsed.has(key)) parsed.set(key, parse(read(c, suffix))); return parsed.get(key); };

let compared = 0;
const leftOut = [];
const stated = {};
const changed = new Set();
for (const c of integerCases) {
  const [typeName, property, steps, ruleAttribute, savedAs] = ATTRIBUTES[c.attribute] ?? assert.fail(`${c.id}: ${c.attribute} unmapped`);
  const value = Number(c.value);
  const writtenRule = c.attribute.match(/\[(auto|exact|atLeast)\]/)?.[1] ?? (property === 'h' && typeName.endsWith('CTFramePr') ? 'exact' : undefined);
  const reading = wordIntegerReading(typeName, property, value, writtenRule);
  assert.ok(reading, `${c.id}: in the table`);
  const key = `${c.round} ${c.id}`;
  if (!reading.kept) changed.add(key);
  if (NOT_MEASURED[c.attribute]) { (stated[c.attribute] ??= []).push(c.id); continue; }
  const saved = elementIn(blockOf(docOf(c, '-word365'), c.id), steps);
  const storage = WORD_INTEGER_STORAGE[typeName][property];
  // What Word's save wrote: the value under its name (the pair's other name where the sign moved it), and its rule.
  let savedValue;
  let savedName = savedAs ?? property;
  if (storage.pair) {
    const positive = integerAttribute(saved, storage.pair.positive);
    const negative = integerAttribute(saved, storage.pair.negative);
    [savedName, savedValue] = negative !== undefined ? [storage.pair.negative, negative] : [storage.pair.positive, positive];
  } else {
    savedValue = integerAttribute(saved, savedName);
  }
  const savedRule = ruleAttribute ? (saved?.getAttributeNS(W, ruleAttribute) || undefined) : undefined;
  if (ruleAttribute === 'hRule' && writtenRule === 'auto' && storage.signRule) {
    assert.equal(savedValue, undefined, `${c.id}: a row height under auto, dropped`);
    compared++;
    continue;
  }
  if (savedValue === undefined) {
    // Word leaves out a value that reads as the default 0.
    assert.equal(reading.written, 0, `${c.id} ${c.attribute}=${value}: Word left it out, so it read 0`);
    leftOut.push(key);
    continue;
  }
  if (reading.written === null && storage.negative === 'computed') {
    assert.ok(savedValue > 0 && savedValue !== value, `${c.id}: a negative column width, computed (${savedValue})`);
    compared++;
    continue;
  }
  assert.equal(reading.written, savedValue, `${c.id} ${c.attribute}=${value}: Word's save wrote ${savedValue}`);
  if (storage.pair) assert.equal(reading.attribute, savedName, `${c.id}: written as ${savedName}`);
  if (ruleAttribute === 'lineRule') assert.equal(reading.rule, savedRule, `${c.id}: the line rule`);
  if (ruleAttribute === 'hRule' && storage.signRule && !storage.heightFlag) assert.equal(reading.rule === 'exact' ? 'exact' : undefined, savedRule, `${c.id}: the row rule`);
  if (storage.heightFlag) assert.equal(reading.rule === 'exact' ? 'exact' : undefined, savedRule, `${c.id}: the frame's rule`);
  // Word reads the value as written exactly when its save writes it back, but 65535 where Word writes it -1.
  const sameInSave = savedValue === value && savedName === (savedAs ?? property) && (writtenRule === undefined || savedRule === writtenRule
    || (writtenRule === 'atLeast' && ruleAttribute === 'lineRule' && savedRule === 'atLeast'));
  const rowRuleSame = ruleAttribute !== 'hRule' || storage.heightFlag || savedRule === (writtenRule === 'exact' ? 'exact' : undefined);
  assert.equal(reading.kept, (sameInSave && rowRuleSame) || (storage.maxWrittenAs !== undefined && value === 65535), `${c.id}: kept`);
  compared++;
}
assert.deepEqual(Object.fromEntries(Object.entries(stated).map(([k, v]) => [k, v.length])), { 'w:gridCol/@w:w': 10, 'w:wBefore/@w:w': 10, 'w:wAfter/@w:w': 10 });

// The facade: onOutOfRange on read reports exactly the cases Word reads differently, and on save the same, unchanged.
const reported = { read: new Set(), save: new Set() };
for (const [key, doc] of [...new Set(integerCases.map((c) => `${c.round}\t${c.file}\t${c.part}`))].map((k) => [k, k.split('\t')])) {
  const [round, file, part] = doc;
  const c = integerCases.find((x) => x.round === round && x.file === file);
  const reportsOn = { read: [], save: [] };
  const tree = await unmarshalString(read(c), { onOutOfRange: (r) => reportsOn.read.push(r) });
  const xml = await marshalString(tree, { onOutOfRange: (r) => reportsOn.save.push(r) });
  for (const side of ['read', 'save']) {
    for (const r of reportsOn[side]) {
      const id = part === 'word/settings.xml' ? file.slice(c.round === 'cr066-round6' ? 'cr066-r6-'.length : 'cr066-'.length) : caseOf(r.element);
      assert.ok(integerCases.some((x) => x.round === round && x.id === id && Number(x.value) === r.value), `${side}: ${r.attribute.localPart}=${r.value} belongs to case ${id} (${key})`);
      assert.equal(r.wordReads.kept, false);
      // The entry on the report (2026-10-27): the key pair names the table entry, which is the report's own.
      assert.equal(r.storage, WORD_INTEGER_STORAGE[r.typeName][r.property], `${side}: ${id}'s report carries its table entry`);
      assert.equal(r.property, r.attribute.localPart);
      reported[side].add(`${round} ${id}`);
    }
  }
  assert.equal(reportsOn.save.length, reportsOn.read.length, `${file}: the same values on save`);
  // Never rewritten: every case's value is written as the probe holds it.
  for (const x of integerCases.filter((y) => y.round === round && y.file === file)) {
    const [, property, steps] = ATTRIBUTES[x.attribute];
    const e = elementIn(blockOf(parse(xml), x.id), steps);
    assert.equal(integerAttribute(e, property), Number(x.value), `${x.id}: written as held`);
  }
}
assert.deepEqual([...reported.read].sort(), [...changed].sort(), 'read: reported exactly where Word reads differently');
assert.deepEqual([...reported.save].sort(), [...changed].sort(), 'save: the same cases');

// Round 6's group E (CR-015): the forms with no digit, read as 0, as Word's save writes them under a style's 100.
const groupE = CASES.filter((c) => c.group === 'E');
const eFile = { ...groupE[0], file: 'cr066-r6-E' };
const eReadings = [];
const eTree = await unmarshalString(read(eFile), { onWordReading: (r) => eReadings.push(r) });
const eXml = parse(await marshalString(eTree));
const eWord = parse(read(eFile, '-word365'));
for (const c of groupE) {
  const [, property, steps] = ATTRIBUTES[c.attribute];
  assert.equal(integerAttribute(elementIn(blockOf(eXml, c.id), steps), property), 0, `${c.id} "${c.value}": read as 0`);
  assert.equal(integerAttribute(elementIn(blockOf(eWord, c.id), steps), property), 0, `${c.id}: Word's save wrote 0`);
}
assert.deepEqual(eReadings.map((r) => [r.text, r.value, r.kind]).sort(), groupE.map((c) => [c.value, 0, 'value.kept.invalid']).sort());

console.log(`word365 ranges (CR-016): ${integerCases.length} cases of docx4j CR-066's rounds 5 and 6, each read by wordIntegerReading as Word's save writes it `
  + `(${compared} compared, ${leftOut.length} read as 0 and left out by Word, ${Object.values(stated).flat().length} stated: gridCol recomputed, wBefore/wAfter dropped); `
  + `${changed.size} reported by onOutOfRange on read and on save, the rest not; nothing rewritten; group E's ${groupE.length} no-digit forms read as 0 (CR-015) OK`);
