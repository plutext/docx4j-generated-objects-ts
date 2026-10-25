// Runtime check of CR-015 against Word 365: docx4j CR-065's shared probe (Jason, 2026-10-26, rounds 2 to 4), a decimal
// or an odd form in an integer attribute. test/fixtures/cr065-probe/ holds each probe part and, for each file Word
// opened, Word's save of it, with a SOURCE.md. Word's save is the oracle: every value it writes is its reading. So each
// case read through the facade and written again must write what Word's save wrote, except where stated below. For a
// file Word refused, the oracle is the refusal: each value is dropped and reported (phase 2's strictIntegers brings
// the forms the runtime's Number() would read, " 240 ", 0x10, 1e3, to the same rule).
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { unmarshalString, marshalString } from '../dist/index.mjs';

const DIR = new URL('./fixtures/cr065-probe/', import.meta.url);
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const FILES = readdirSync(DIR).filter((d) => d.endsWith('.docx') && !d.endsWith('-word365.docx')).map((d) => d.slice(0, -'.docx'.length)).sort();
assert.equal(FILES.length, 20, 'A1 to A7 split seven ways (A2 six), and B to I');
const part = (file, name) => readFileSync(new URL(`${file}.docx/${name}`, DIR), 'utf8');
const saved = (file) => existsSync(new URL(`${file}-word365.docx`, DIR));

/** The cases whose value cases.txt does not give as one path. */
const SPECIAL = {
  H56: { steps: ['pgSz'], attributes: ['w', 'h'] },
  H57: { steps: ['pgSz'], attributes: ['w', 'h'] },
  H60: { steps: ['pgMar'], attributes: ['header', 'footer', 'gutter'] },
  I62: { part: 'styles', steps: ['ind'], attributes: ['left'] },
  I63: { part: 'settings', steps: ['defaultTabStop'], attributes: ['val'] },
  I64: { part: 'header', steps: ['spacing'], attributes: ['before'] },
};

/** Each case of cases.txt: its id, and where its value is (element steps and attribute), or a reader of its own. */
const CASES = readFileSync(new URL('cases.txt', DIR), 'utf8').trim().split('\n').map((line) => line.split('\t')[1])
  .filter((c) => !c.startsWith('Z')).map((c) => {
    const id = c.split(' ')[0];
    const m = c.match(/^\S+ ((?:w:\w+\/)*w:\w+)\/@w:(\w+)="/);
    return m ? { id, steps: m[1].split('/').map((s) => s.slice(2)), attributes: [m[2]] } : { id, ...SPECIAL[id] };
  });
assert.equal(CASES.length, 64);

/**
 * Where ours and Word's save differ, each stated, strict both ways:
 * - a value truncated to what lies beneath, which Word's save leaves out (the paragraph's default 0, the document's
 *   default font size 22), where this package writes what it read;
 * - w:gridCol, which Word recomputes from the table's widths.
 */
const KNOWN = {
  A08: 'truncated to 0, the default; Word leaves it out',
  A17: '"" read as 0, the default; Word leaves it out',
  B25: 'truncated to 0, the default; Word leaves it out',
  F41: 'truncated to 22, the document default; Word leaves it out',
  G50: 'w:gridCol, recomputed by Word',
  G51: 'w:gridCol, recomputed by Word',
};

/** The forms of the files Word refused, each dropped and reported as 'value.refused'. */
const REFUSED = { A11: ' 240.5 ', A12: '240,5', A13: '0x10', A14: '1.5e2', A15: '2.5e0', A16: '1e3', A18: 'abc', A19: ' 240 ' };

const textOf = (p) => Array.from(p.getElementsByTagNameNS(W, 't')).map((t) => t.textContent).join('');
const elementsAfter = (node) => { const out = []; for (let n = node.nextSibling; n; n = n.nextSibling) if (n.nodeType === 1) { out.push(n); break; } return out; };

/** The label's paragraph and, where the case's value is in it, the table or the section-ending paragraph after it. */
function blockOf(doc, id) {
  const p = Array.from(doc.getElementsByTagNameNS(W, 'p')).find((x) => textOf(x).startsWith(`${id} `));
  if (!p) return undefined;
  const [next] = elementsAfter(p);
  const holds = next && (next.localName === 'tbl' || (next.localName === 'p' && next.getElementsByTagNameNS(W, 'sectPr').length > 0 && textOf(next) === ''));
  return holds ? [p, next] : [p];
}

/** Each value the case's attributes have in the block, in document order: the element named by the last step, under the steps before it. */
function valuesIn(scope, { steps, attributes }) {
  const out = [];
  for (const root of scope) {
    const candidates = [root, ...Array.from(root.getElementsByTagNameNS(W, steps[steps.length - 1]))].filter((e) => e.localName === steps[steps.length - 1]);
    for (const e of candidates) {
      let ok = true;
      for (let i = steps.length - 2, n = e.parentNode; i >= 0; i--, n = n?.parentNode) if (n?.localName !== steps[i]) { ok = false; break; }
      if (!ok) continue;
      for (const a of attributes) if (e.hasAttributeNS(W, a)) out.push(`${a}=${e.getAttributeNS(W, a)}`);
    }
  }
  return out;
}

const parse = (xml) => new DOMParser().parseFromString(xml, 'text/xml');

/** A part read through the facade and written again, with what was reported. */
async function throughFacade(xml) {
  const readings = [];
  const normalised = [];
  const out = await marshalString(await unmarshalString(xml, { onWordReading: (r) => readings.push(r) }), { onNormalised: (n) => normalised.push(n) });
  return { doc: parse(out), readings, normalised };
}

/** The file holding a case: the one whose probe part has its label. */
const fileOf = (c) => FILES.find((f) => f.startsWith(`cr065-${c.id[0]}`) && new RegExp(`>${c.id} `).test(part(f, 'word/document.xml')));
const PARTS = { styles: 'word/styles.xml', settings: 'word/settings.xml' };
const headerOf = (file) => readdirSync(new URL(`${file}.docx/word/`, DIR)).find((n) => /^header\d*\.xml$/.test(n));

const ours = new Map();
const words = new Map();
async function read(file, which) {
  const key = `${file} ${which}`;
  if (!ours.has(key)) {
    const name = which === 'document' ? 'word/document.xml' : which === 'header' ? `word/${headerOf(file)}` : PARTS[which];
    ours.set(key, await throughFacade(part(file, name)));
    if (saved(file)) {
      const savedName = which === 'header' ? `word/${headerOf(`${file}-word365`)}` : name;
      const word = await throughFacade(part(`${file}-word365`, savedName));
      assert.deepEqual([word.readings.length, word.normalised.length], [0, 0], `${file}: Word's save needs neither Word's reading nor the normaliser`);
      words.set(key, word);
    }
  }
  return { ours: ours.get(key), word: words.get(key) };
}

let equal = 0;
const differing = {};
for (const c of CASES) {
  const file = fileOf(c);
  assert.ok(file, `${c.id}: in a probe file`);
  const which = c.part ?? 'document';
  const { ours: mine, word } = await read(file, which);
  const scope = (doc) => (which === 'document' ? blockOf(doc, c.id)
    : which === 'styles' ? Array.from(doc.getElementsByTagNameNS(W, 'style')).filter((st) => st.getAttributeNS(W, 'styleId') === 'ProbeStyle')
      : [doc.documentElement]);
  const ourScope = scope(mine.doc);
  assert.ok(ourScope, `${c.id}: its label`);
  const ourValues = valuesIn(ourScope, c);
  if (!word) {
    // Word refused the file: each value is dropped and reported.
    assert.ok(c.id in REFUSED, `${c.id}: a refused form`);
    assert.deepEqual(ourValues, [], `${c.id}: dropped`);
    const reading = mine.readings.find((r) => r.text === REFUSED[c.id]);
    assert.ok(reading && reading.kind === 'value.refused' && reading.value === undefined, `${c.id}: reported as refused`);
    continue;
  }
  const wordValues = valuesIn(scope(word.doc), c);
  if (JSON.stringify(ourValues) === JSON.stringify(wordValues)) equal++;
  else differing[c.id] = `${JSON.stringify(ourValues)} vs Word ${JSON.stringify(wordValues)}`;
}
assert.deepEqual(Object.keys(differing).sort(), Object.keys(KNOWN).sort(), `the known differences, strict both ways: ${JSON.stringify(differing, null, 1)}`);

// What was reported: each value read by Word's rule, truncated, "" as 0, or refused.
let truncated = 0;
let refused = 0;
let empty = 0;
for (const { readings } of ours.values()) {
  for (const r of readings) {
    if (r.kind === 'value.refused') { refused++; assert.equal(r.value, undefined); continue; }
    if (r.kind === 'value.kept.invalid') { empty++; assert.deepEqual([r.text, r.value], ['', 0]); continue; }
    assert.equal(r.kind, 'value.truncated', `${r.attribute.localPart}=${JSON.stringify(r.text)}`);
    assert.equal(r.value, Math.trunc(Number(r.text)) || 0, `${r.attribute.localPart}=${JSON.stringify(r.text)} truncated`);
    truncated++;
  }
}
assert.equal(refused, Object.keys(REFUSED).length, 'one record per refused form');
assert.equal(empty, 5, '"" read as 0: A6, SSRS\'s header, footer and gutter, and an empty pgMar top');
// SSRS's empty rsids are left out on save, as Word's save leaves them out (N7); the margins read "" as 0, as Word does.
const ssrs = (await read(fileOf({ id: 'H60' }), 'document')).ours;
const ssrsSectPr = blockOf(ssrs.doc, 'H60')[1].getElementsByTagNameNS(W, 'sectPr')[0];
assert.ok(!Array.from(ssrsSectPr.attributes).some((a) => a.localName.startsWith('rsid')), 'no empty rsid written');
assert.deepEqual(ssrs.normalised.map((n) => [n.path.split('/@')[1], n.before, n.after]).sort(),
  [['w:rsidDel', '', null], ['w:rsidR', '', null], ['w:rsidRPr', '', null], ['w:rsidSect', '', null]]);

console.log(`word365 integers (CR-015): ${CASES.length} cases of docx4j CR-065's probe in ${FILES.length} files; ${equal} written as Word's saves write them, ${Object.keys(KNOWN).length} known (a default Word leaves out, gridCol recomputed); `
  + `${truncated} values truncated, ${empty} "" read as 0 and ${refused} refused forms dropped, each reported; SSRS's empty rsids left out OK`);
