// Runtime check of CR-022 against Word 365: a child given more than once where the model holds one (WD REQ-322). The files are
// docx4j's repeated/ at docx4j-17.3.6 (test/fixtures/repeated-probe/, with a SOURCE.md): the rounds rc-, rt-, sm-, ft-, c79- and op-,
// each source's parts beside Word's save of it. Word's save is the oracle. The comparisons are docx4j's RepeatedChildrenTest, which
// ported docx4j-python's, so the three engines are held to the same reading: each is read through the facade and written again.
import { readdirSync, readFileSync, existsSync } from 'node:fs';
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { unmarshalString, marshalString, unwrap } from '../dist/index.mjs';
import { textOf, walk } from '../dist/builders/wml.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const dir = new URL('fixtures/repeated-probe/', import.meta.url);

const cases = readdirSync(dir).filter((name) => existsSync(new URL(`${name}/source/word/document.xml`, dir))).sort();
const count = (prefix) => cases.filter((name) => name.startsWith(prefix)).length;
assert.deepEqual([count('rc-'), count('rt-'), count('sm-'), count('ft-'), count('c79-'), count('op-')], [16, 20, 40, 12, 7, 6],
  'the rounds: 16 rc, 20 rt, 40 sm, 12 ft, 7 c79, 6 op');
const named = (prefix) => cases.filter((name) => name.startsWith(prefix));

const read = (path) => readFileSync(new URL(path, dir), 'utf8');
const parse = (xml) => new DOMParser().parseFromString(xml, 'application/xml').documentElement;

/** The source's main part, read through the facade. */
const source = (name) => unmarshalString(read(`${name}/source/word/document.xml`));
/** The facade's save of the source's main part. */
const saved = async (name) => parse(await marshalString(await source(name)));
/** Word's save. */
const word = (name) => parse(read(`${name}/word365/word/document.xml`));

const all = (e, local) => Array.from(e.getElementsByTagNameNS(W, local));
const children = (e) => Array.from(e.childNodes).filter((n) => n.nodeType === 1);

/** An attribute's value as Word writes an on/off one. */
const spelled = (local, value) => local !== 'val' ? value
  : value === 'true' || value === 'on' ? '1' : value === 'false' || value === 'off' ? '0' : value;

function attributes(e) {
  const out = {};
  for (const a of Array.from(e.attributes)) {
    const local = a.localName ?? a.nodeName;
    if (local.startsWith('rsid') || a.prefix === 'xmlns' || a.nodeName === 'xmlns') continue;
    out[local] = spelled(local, a.nodeValue);
  }
  return out;
}

function target(root, words) {
  const found = all(root, 'p').find((p) => p.textContent.includes(words));
  assert.ok(found, `no paragraph holding ${words}`);
  return found;
}

const isOff = (facts) => facts !== undefined && Object.keys(facts).length === 1 && facts.val === '0';

/** python's target_properties: what the paragraph's w:pPr and w:rPr hold, each child's attributes, the last of a name. */
function targetProperties(root, words) {
  const out = {};
  const holders = [];
  const collect = (e) => {
    for (const c of children(e)) {
      if (c.namespaceURI === W && (c.localName === 'pPr' || c.localName === 'rPr')) holders.push(c);
      collect(c);
    }
  };
  collect(target(root, words));
  for (const properties of holders) for (const child of children(properties)) out[child.localName] = attributes(child);
  // an on/off property stated off reads as one not stated, where nothing beneath sets it: Word's save writes nothing (rc-31, sm-80)
  for (const onOff of ['b', 'vanish']) if (isOff(out[onOff])) delete out[onOff];
  return out;
}

/** python's _stated: what a w:sectPr, w:tblPr, w:pPr or w:rPr states, child by child, beside Word's own additions. */
function stated(element) {
  const facts = {};
  const section = element.localName === 'sectPr';
  const table = element.localName === 'tblPr';
  let headers = 0;
  for (const child of children(element)) {
    const name = child.localName;
    if (name === 'headerReference') { headers++; continue; }
    if (table && (name === 'tblCellMar' || name === 'tblLook')) continue; // Word's own defaults
    const said = attributes(child);
    if (said.orient === 'portrait') delete said.orient; // the default, which Word's save leaves out (sm-07)
    if (name === 'cols' && !('space' in said)) said.space = '720'; // likewise the other way (sm-32)
    facts[name] = said;
  }
  if (section) {
    if (!('cols' in facts)) facts.cols = { space: '720' };
    facts.headerReference = { count: String(headers) };
  }
  return facts;
}

/** python's _merged: every section and table, and the target paragraph's properties and its runs'. */
function merged(root) {
  const paragraph = [];
  const runs = [];
  for (const c of children(target(root, 'TARGET TEXT'))) {
    if (c.localName === 'pPr' && children(c).length > 0) paragraph.push(stated(c));
    if (c.localName === 'r') {
      for (const rPr of children(c)) {
        if (rPr.localName !== 'rPr' || children(rPr).length === 0) continue;
        const facts = stated(rPr);
        if (!(Object.keys(facts).length === 1 && isOff(facts.vanish))) runs.push(facts);
      }
    }
  }
  return { sections: all(root, 'sectPr').map(stated), tables: all(root, 'tblPr').map(stated), paragraph, runs };
}

const pageWidths = (root) => all(root, 'sectPr').flatMap((s) => children(s).filter((c) => c.localName === 'pgSz').map((c) => c.getAttributeNS(W, 'w')));

let checks = 0;
const differ = [];
/** A comparison that fails is recorded, so that every difference is listed, then held to the stated ones below. */
function compare(name, what, ours, theirs) {
  checks++;
  try { assert.deepStrictEqual(ours, theirs); } catch { differ.push(`${name}: ${what}`); }
}

// --- a property given twice: the target paragraph's properties -----------------------------------------------------------------
const PROPERTIES = ['rc-20-pPr-values', 'rc-21-pPr-merge', 'rc-22-rPr-values', 'rc-23-rPr-merge', 'rc-30-jc', 'rc-31-b-on-off',
  'rc-32-b-off-on', 'rc-33-sz', 'rt-20-pPr-values', 'rt-21-pPr-merge', 'rt-22-rPr-values', 'rt-23-rPr-merge', 'rt-30-jc',
  'rt-31-b-on-off-on', 'rt-32-b-off-on-off', 'rt-33-sz'];
for (const name of PROPERTIES) compare(name, 'the target paragraph\'s properties', targetProperties(await saved(name), 'TARGET TEXT'), targetProperties(word(name), 'TARGET TEXT'));
// c79-40: a w:pPr's w:rPr merges too; c79-42: a w:pPrChange in the first copy survives
compare('c79-40-ppr-nested-rpr', 'the marked paragraph\'s properties', targetProperties(await saved('c79-40-ppr-nested-rpr'), 'MARKED WORDS'), targetProperties(word('c79-40-ppr-nested-rpr'), 'MARKED WORDS'));
compare('c79-42-ppr-pprchange', 'the first copy\'s w:pPrChange', all(target(await saved('c79-42-ppr-pprchange'), 'CHANGED WORDS'), 'pPrChange').length, 1);

// --- c79-41 and op-20: a w:numPr applied to one object --------------------------------------------------------------------------
const numPrOf = (root, words) => {
  const numPr = all(words ? target(root, words) : root, 'numPr')[0];
  return Object.fromEntries(children(numPr).map((c) => [c.localName, c.getAttributeNS(W, 'val')]));
};
// c79-41: the first's w:ilvl with the second's w:numId; op-20: two w:numPr in one w:pPr, the same
for (const [name, words, expected] of [['c79-41-ppr-numpr', 'LISTED WORDS', { ilvl: '0', numId: '2' }], ['op-20-numpr-in-one-ppr', undefined, { ilvl: '1', numId: '2' }]]) {
  assert.deepEqual(numPrOf(word(name), words), expected, `${name}: Word's save`);
  compare(name, 'the w:numPr', numPrOf(await saved(name), words), expected);
}

// --- the words, in order: the facade's reading of the source against its reading of Word's save --------------------------------------
for (const name of [...named('rc-'), ...named('rt-'), ...named('c79-3')]) {
  const ours = textOf(unwrap(await source(name)));
  compare(name, 'the words', ours, textOf(unwrap(await unmarshalString(read(`${name}/word365/word/document.xml`)))));
}

// --- a text box's words, in order: textOf doesn't read a w:pict, so each w:txbxContent's text runs, ours against Word's save ---------
const boxes = (root) => all(root, 'txbxContent').map((box) => all(box, 't').map((t) => t.textContent));
for (const name of cases.filter((name) => read(`${name}/source/word/document.xml`).includes('txbxContent'))) {
  compare(name, 'the text boxes', boxes(await saved(name)), boxes(word(name)));
}

// --- and against the text Word saved from each source, every case: textOf reads the main story, so where Word's text holds a header's
// or a footer's words, or a list's label, it can't be compared, and those cases are named, strict both ways ---------------------------
const linesOf = (text) => text.split(/\r?\n|\r|\f|\x07|\x0b/).map((line) => line.trim()).filter(Boolean);
const NOT_THE_MAIN_STORY = {
  'a header\'s or a footer\'s words': ['ft-00-control', 'ft-01-footer-first-only', 'ft-03-footer-both', 'ft-04-three-first-and-third',
    'ft-06-first-both-later-header', 'ft-07-first-header-later-footer', 'ft-08-footer-default-then-first', 'ft-09-header-default-then-first',
    'ft-10-two-bodies-footers', 'ft-11-two-bodies-headers', 'sm-40-header-first', 'sm-42-header-both'],
  'a list\'s label': ['c79-41-ppr-numpr', 'op-20-numpr-in-one-ppr'],
};
const notCompared = [];
for (const name of cases) {
  checks++;
  const wordText = linesOf(readFileSync(new URL(`${name}/word365.txt`, dir), 'latin1'));
  try { assert.deepStrictEqual(linesOf(textOf(unwrap(await source(name)))), wordText); } catch { notCompared.push(name); }
}

// --- every value's PARENT, and a save read again repeats nothing ------------------------------------------------------------------------
let walked = 0;
for (const name of cases) {
  const document = await source(name);
  walk(unwrap(document), (value, parent) => {
    if (!value.TYPE_NAME || !parent) return;
    walked++;
    assert.equal(value.PARENT, parent, `${name}: ${value.TYPE_NAME}'s PARENT`);
  });
  let repeated = 0;
  await unmarshalString(await marshalString(document), { onOverwrittenElement() { repeated++; } });
  assert.equal(repeated, 0, `${name}: the save, read again, repeats nothing`);
}

// --- the sections a save holds -----------------------------------------------------------------------------------------------------
for (const name of [...named('rc-1'), ...named('rt-1')]) {
  const ours = await saved(name);
  const theirs = word(name);
  if (pageWidths(ours).length === 0) continue; // a file that states no section has none, where Word's save writes a default
  compare(name, 'the page widths', pageWidths(ours), pageWidths(theirs));
  compare(name, 'the sections', all(ours, 'sectPr').map(stated), all(theirs, 'sectPr').map(stated));
}

// --- the sectpr-merge round --------------------------------------------------------------------------------------------------------
for (const name of named('sm-')) {
  if (name.startsWith('sm-61')) continue; // below
  compare(name, 'the sections, tables and target properties', merged(await saved(name)), merged(word(name)));
}

// --- the one-parent-word-check round: a child given twice in one parent -------------------------------------------------------------
for (const name of ['op-10-pgsz-in-one-sectpr', 'op-11-pgmar-in-one-sectpr', 'op-12-cols-in-one-sectpr']) {
  compare(name, 'the sections', all(await saved(name), 'sectPr').map(stated), all(word(name), 'sectPr').map(stated));
}
{
  const markOf = (root) => { let found; for (const pPr of all(root, 'pPr')) for (const c of children(pPr)) if (c.localName === 'rPr') found = stated(c); return found; };
  const ours = markOf(await saved('op-21-rpr-in-one-ppr'));
  compare('op-21-rpr-in-one-ppr', 'the paragraph mark\'s properties', ours, markOf(word('op-21-rpr-in-one-ppr')));
  assert.ok('b' in markOf(word('op-21-rpr-in-one-ppr')), 'op-21: Word\'s save holds the mark\'s bold');
}

// --- the footer-refs round, and sm-40 to sm-42: only the first copy's references stand -------------------------------------------------
function references(root, at) {
  const rels = parse(read(`${at}/word/_rels/document.xml.rels`));
  const targets = Object.fromEntries(children(rels).map((r) => [r.getAttribute('Id'), r.getAttribute('Target')]));
  return all(root, 'sectPr').map((section) => children(section)
    .filter((ref) => ref.localName === 'headerReference' || ref.localName === 'footerReference')
    .map((ref) => `${ref.localName} ${ref.getAttributeNS(W, 'type')} ${parse(read(`${at}/word/${targets[ref.getAttributeNS(R, 'id')]}`)).textContent.trim()}`)
    .sort());
}
for (const name of [...named('ft-'), ...named('sm-4')]) {
  compare(name, 'the references', references(await saved(name), `${name}/source`), references(word(name), `${name}/word365`));
}

// --- sm-61: w:tblW stays the later copy's ------------------------------------------------------------------------------------------------
{
  // REQ-322 says nothing of w:tblW, and Word's save fits neither reading, so the later copy's whole
  const tblW = all(await saved('sm-61-tblpr-attributes'), 'tblW').at(-1);
  compare('sm-61-tblpr-attributes', 'w:tblW, the later copy\'s', attributes(tblW), { type: 'auto' });
  const table = '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblPr><w:jc w:val="center"/></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const tbl = unwrap(await unmarshalString(`<w:document xmlns:w="${W}"><w:body>${table}<w:p/></w:body></w:document>`)).body.content[0];
  const tblPr = (tbl.value ?? tbl).tblPr;
  // where the later copy states no w:tblW, none: docx4j CR-079 section 4's decision, as Word's reading isn't known
  compare('a table built here', 'no w:tblW, and the earlier w:jc', [tblPr.tblW, tblPr.jc?.val], [undefined, 'center']);
}

// --- CR-022 D7: a header and a footer, this package's rounds header-repeats and header-footer-repeats (test/fixtures/header-footer-probe/,
// with a SOURCE.md): seven main-part cases put into a header as they are, then a header's w:pPr in seven forms, three copies of the
// other shapes, the same shapes in a footer, and main-part twins of the new markup. Each story part is read through the facade and
// written again, against Word's save of it: its controls' and text boxes' words, and the target paragraph's properties by path ------
const hfDir = new URL('fixtures/header-footer-probe/', import.meta.url);
const hfCases = JSON.parse(readFileSync(new URL('cases.json', hfDir), 'utf8'));
const hfRead = (path) => readFileSync(new URL(path, hfDir), 'utf8');
/** The story part a case's main part refers to, or the main part itself. */
function storyPart(id, side, story) {
  if (story === 'main') return 'document.xml';
  const reference = all(parse(hfRead(`${id}/${side}/word/document.xml`)), story === 'header' ? 'headerReference' : 'footerReference')[0];
  const rels = children(parse(hfRead(`${id}/${side}/word/_rels/document.xml.rels`)));
  return rels.find((r) => r.getAttribute('Id') === reference.getAttributeNS(R, 'id')).getAttribute('Target');
}
/** The target paragraph's every property, keyed by its path from the paragraph (pPr/jc, pPr/rPr/b, r/rPr/i), an off w:b or w:vanish as one not stated. */
function propertiesByPath(p) {
  const out = {};
  const visit = (e, path) => {
    for (const k of children(e)) {
      if (k.localName === 't') continue;
      const here = path ? `${path}/${k.localName}` : k.localName;
      if (children(k).length === 0) out[here] = attributes(k); else visit(k, here);
    }
  };
  visit(p, '');
  for (const key of Object.keys(out)) if (/\/(b|vanish)$/.test(key) && isOff(out[key])) delete out[key];
  return out;
}
const storyFacts = (root) => {
  const p = all(root, 'p').find((e) => /TARGET TEXT|MARKED WORDS/.test(e.textContent) && all(e, 'p').length === 0);
  return { text: all(root, 'p').filter((e) => all(e, 'p').length === 0).map((e) => e.textContent),
    controls: all(root, 'sdtContent').map((c) => all(c, 'p').map((e) => e.textContent)),
    boxes: all(root, 'txbxContent').map((b) => all(b, 't').map((t) => t.textContent)), target: p && propertiesByPath(p) };
};
let hfStories = 0;
for (const [id, { story }] of Object.entries(hfCases)) {
  for (const kind of story === 'both' ? ['header', 'footer'] : [story]) {
    hfStories++;
    const ours = parse(await marshalString(await unmarshalString(hfRead(`${id}/source/word/${storyPart(id, 'source', kind)}`))));
    compare(id, `the ${kind}`, storyFacts(ours), storyFacts(parse(hfRead(`${id}/word365/word/${storyPart(id, 'word365', kind)}`))));
  }
}

// --- the stated differences, strict both ways --------------------------------------------------------------------------------------------
// An mc:AlternateContent root holding the copies, which core-ts's root pass reads (its CR-028 section 3.1): the facade reads a root
// branch's w:document children as a list, which repeats nothing, and textOf gives no text for such a root (CR-021).
const MC_ROOTS = ['rt-50-mc-root-three-in-choice', 'rt-51-mc-root-three-in-fallback', 'sm-52-branch-orient-first'];
// In a header, a first copy's w:jc left and an empty first w:pPr are kept, where Word's save leaves out what states only the default,
// to the same effect (hf-13, hf-14). And a later copy's paragraph-mark w:rPr, which Word puts on the paragraph's run, isn't followed
// (hf-16, one file).
const STATED = ['rt-50-mc-root-three-in-choice: the words', 'rt-51-mc-root-three-in-fallback: the words',
  'sm-52-branch-orient-first: the sections, tables and target properties',
  'hf-13-hdr-ppr-values-three: the header', 'hf-14-hdr-ppr-empty-first: the header', 'hf-16-hdr-ppr-nested-rpr: the header'];
assert.deepEqual(differ, STATED, 'every difference from Word\'s save is a stated one, and every stated one still differs');
assert.deepEqual(notCompared.sort(), [...Object.values(NOT_THE_MAIN_STORY).flat(), ...MC_ROOTS].sort(),
  'textOf gives Word\'s text but where Word\'s holds more than the main story, and for an mc root');

console.log(`word365 repeated (CR-022): docx4j's ${cases.length} files and this package's ${Object.keys(hfCases).length} in a header or a footer (${hfStories} stories), `
  + `${checks} checks against Word's saves and texts: ${differ.length} stated, the mc roots core-ts reads and three in a header; `
  + `textOf is Word's text in ${cases.length - notCompared.length}; ${walked} values' PARENT; no save repeats a child`);
