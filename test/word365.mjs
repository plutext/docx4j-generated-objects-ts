// Runtime check of CR-013 against Word 365: Jason's run property probes (2026-10-24), each read as Word read it, and
// each saved as docx4j saves it. test/fixtures/word365-probes/ holds each probe's word/document.xml, Word 365's save
// of it and, for fifteen, docx4j's save (CR-061, value for value with phase C), with a SOURCE.md.
// Every probe paragraph after the title is "label: MARKER (end)", three runs, and only the MARKER run carries the
// property tested. Word's save is the oracle: Word writes back only values it reads one way, so a marker's reading
// in the probe must equal its reading in the save, as docx4j's RunPropertiesWord365Test has it. The readings are
// of direct formatting (this package resolves no styles), so where Word's save changes direct formatting without
// changing what Word shows, or changes what Word shows, the difference is stated below.
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import assert from 'node:assert/strict';
import { unmarshalString, marshalString, unwrap, setNormaliseOnSave } from '../dist/index.mjs';
import { textOf, readRunOptions } from '../dist/builders/wml.mjs';
import { wordColor } from '../dist/helpers/wml.mjs';

const DIR = new URL('./fixtures/word365-probes/', import.meta.url);
const PROBES = readdirSync(DIR).filter((d) => d.endsWith('.docx') && !/-(word365|docx4j)\.docx$/.test(d)).map((d) => d.slice(0, -'.docx'.length)).sort();
assert.equal(PROBES.length, 33, 'rounds 1 to 7');

const TOGGLES = ['b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'vanish'];

/**
 * A marker run's direct formatting, as far as the probes test it. A valueless w:u is unset, as Word reads it (its
 * save drops one). A fill, pattern or border colour Word rejects is automatic (round 6). Where neither the run nor
 * its paragraph has a style, a direct value equal to the default (w:u none, w:color auto, a clear shading with an
 * automatic fill) is unset too: Word's save drops it, and with nothing beneath they read alike (the unstyled probes
 * have no styles part).
 */
function reading(run, paragraph) {
  const rPr = run.rPr ?? {};
  const unstyled = paragraph.pPr?.pStyle === undefined && rPr.rStyle === undefined;
  const options = readRunOptions(run.rPr);
  const m = {};
  for (const toggle of TOGGLES) m[toggle] = rPr[toggle] !== undefined && rPr[toggle].val !== false;
  m.u = rPr.u?.val ?? 'unset';
  if (m.u === 'none' && unstyled) m.u = 'unset';
  m.uColor = rPr.u === undefined ? '-' : (wordColor(rPr.u.color) ?? '-');
  m.vertAlign = rPr.vertAlign?.val ?? 'unset';
  m.em = rPr.em?.val ?? 'unset';
  m.color = options.color === 'auto' && unstyled ? '' : options.color;
  m.highlight = options.highlightColor;
  m.shading = rPr.shd === undefined ? 'none' : `${rPr.shd.val} ${wordColor(rPr.shd.color) ?? 'auto'} on ${wordColor(rPr.shd.fill) ?? 'auto'}`;
  if (m.shading === 'clear auto on auto' && unstyled) m.shading = 'none';
  m.border = rPr.bdr === undefined ? 'none' : `${rPr.bdr.val} ${wordColor(rPr.bdr.color) ?? 'auto'}`;
  return m;
}

/** The page colour (w:background), read as Word reads it: a rejected colour, like none, is no page colour. */
const pageColour = (doc) => (doc.background === undefined ? 'auto' : (wordColor(doc.background.color) ?? 'auto'));

/**
 * Where Word's own save changes what Word drew, Word's reading from its PDF (docx4j's WORD_DREW; CR-060 section 13):
 * a direct highlight over a character style's, drawn but dropped by the save; an underline colour "auto" over a
 * style's red, drawn black but dropped.
 */
const WORD_DREW = {
  'props-enum-styles M6_HIGHLIGHT_GREEN': { highlight: '#008000' },
  'highlight-char-styles H7_CHAR_STYLE_DIRECT_GREEN': { highlight: '#008000' },
  'colour-elsewhere-styles U4_UNDERLINE_AUTO': { uColor: 'auto' },
};

/**
 * The differences that stay, each in direct formatting only: over the character style's single red underline,
 * Word's save drops a direct w:u single that repeats it (U4's colour aside, which WORD_DREW states). Strict both
 * ways: a known difference that no longer differs fails too.
 */
const KNOWN = new Set([
  'colour-elsewhere-styles U2_UNDERLINE_F00 u',
  'colour-elsewhere-styles U3_UNDERLINE_NAVY u',
  'colour-elsewhere-styles U4_UNDERLINE_AUTO u',
]);

/** What a save wrote for a marker run, as far as phase C writes it: the colour attributes as written, and the elements it may leave out. */
function written(run) {
  const x = run.rPr ?? {};
  const as = (v, absent) => (v === undefined ? absent : (v.val ?? 'bare'));
  return {
    color: x.color === undefined ? 'absent' : `${x.color.val ?? '-'}${x.color.themeColor ? ` beside ${x.color.themeColor}` : ''}`,
    uColor: x.u === undefined ? 'no w:u' : (x.u.color ?? 'none'),
    shading: x.shd === undefined ? 'no w:shd' : `${x.shd.color ?? '-'} on ${x.shd.fill ?? '-'}`,
    border: x.bdr === undefined ? 'no w:bdr' : (x.bdr.color ?? '-'),
    highlight: as(x.highlight, 'absent'),
    vertAlign: as(x.vertAlign, 'absent'),
    em: as(x.em, 'absent'),
  };
}

/** Phase C writes every colour in the schema's form (docx4j's ColourNormalFormTest.outOfForm), and no w:color without a w:val. */
function outOfForm(xml) {
  const bad = [];
  for (const pattern of [/<w:color\b[^>]*?\bw:val="([^"]*)"/g, /<w:u\b[^>]*?\bw:color="([^"]*)"/g, /<w:shd\b[^>]*?\bw:fill="([^"]*)"/g,
    /<w:shd\b[^>]*?\bw:color="([^"]*)"/g, /<w:(?:top|left|bottom|right|start|end|between|bar|bdr|insideH|insideV|tl2br|tr2bl)\b[^>]*?\bw:color="([^"]*)"/g,
    /<w:background\b[^>]*?\bw:color="([^"]*)"/g]) {
    for (const m of xml.matchAll(pattern)) if (!/^(auto|[0-9A-F]{6})$/.test(m[1])) bad.push(m[0]);
  }
  for (const m of xml.matchAll(/<w:color(?:\s+(?!w:val=)[\w:]+="[^"]*")*\s*\/>/g)) bad.push(`${m[0]} (no w:val)`);
  return bad;
}

/** A marker's run in a part's XML (docx4j's ColourNormalFormTest.run). */
function runXml(xml, marker) {
  const i = xml.indexOf(`>${marker}<`);
  assert.ok(i > 0, marker);
  const start = Math.max(xml.lastIndexOf('<w:r>', i), xml.lastIndexOf('<w:r ', i));
  return xml.substring(start, xml.indexOf('</w:r>', i));
}

const paragraphs = (element) => unwrap(element).body.content.map(unwrap).filter((v) => v.TYPE_NAME === 'org_docx4j_wml.P');
const runs = (p) => p.content.map(unwrap).filter((v) => v.TYPE_NAME === 'org_docx4j_wml.R');
const source = (name) => readFileSync(new URL(`${name}.docx/word/document.xml`, DIR), 'utf8');
const load = async (name) => {
  const readings = [];
  const element = await unmarshalString(source(name), { onWordReading: (r) => readings.push(r) });
  return { element, readings };
};

let markers = 0;
let agree = 0;
let readByWordsRule = 0;
let againstDocx4j = 0;
let normalisations = 0;
const found = new Set();
const unexpected = [];
const reloaded = [];
const unlikeDocx4j = [];
const notInForm = [];
const savedXml = {};
for (const name of PROBES) {
  // Before CR-013 phase B every probe with a value Word reads and the runtime refuses failed to load at all.
  const probe = await load(name);
  const word = await load(`${name}-word365`);
  readByWordsRule += probe.readings.length;
  assert.equal(word.readings.length, 0, `${name}: Word's save writes only values the runtime reads`);
  // Saved through the facade (phase C on, by default): in the schema's form, and read again as the probe was.
  const ours = savedXml[name] = await marshalString(probe.element, { onNormalised: () => normalisations++ });
  notInForm.push(...outOfForm(ours).map((b) => `${name}: ${b}`));
  const again = await unmarshalString(ours);
  // Word's own saves, and docx4j's, are in the normal form already.
  await marshalString(word.element, { onNormalised: (n) => assert.fail(`${name}: Word's save normalised at ${n.path}`) });
  const docx4j = existsSync(new URL(`${name}-docx4j.docx`, DIR)) ? await unmarshalString(source(`${name}-docx4j`)) : undefined;
  if (docx4j) await marshalString(docx4j, { onNormalised: (n) => assert.fail(`${name}: docx4j's save normalised at ${n.path}`) });
  const page = pageColour(unwrap(probe.element));
  assert.equal(page, pageColour(unwrap(word.element)), `${name}: the page colour reads as Word read it`);
  assert.equal(pageColour(unwrap(again)), page, `${name}: and is kept by a save`);
  const [ps, ws, as] = [paragraphs(probe.element), paragraphs(word.element), paragraphs(again)];
  const ds = docx4j && paragraphs(docx4j);
  assert.equal(ws.length, ps.length, `${name}: Word keeps the paragraphs one for one`);
  for (let i = 1; i < ps.length; i++) {
    const rs = runs(ps[i]);
    if (rs.length !== 3) continue;
    const text = textOf(rs[1]);
    // Word merges runs whose formatting is equal, so in its save the marker can share a run with its label.
    const runIn = (p) => runs(p).filter((r) => textOf(r).includes(text)).at(-1);
    const wordRun = runIn(ws[i]);
    assert.ok(wordRun, `${name} ${text}: in Word's save`);
    markers++;
    const read = reading(rs[1], ps[i]);
    const words = { ...reading(wordRun, ws[i]), ...WORD_DREW[`${name} ${text}`] };
    const differing = Object.keys(words).filter((k) => read[k] !== words[k]);
    for (const k of differing) {
      const key = `${name} ${text} ${k}`;
      if (KNOWN.has(key)) found.add(key);
      else unexpected.push(`${key}: ours ${read[k]}, Word's ${words[k]}`);
    }
    if (differing.length === 0) agree++;
    const ourRun = runIn(as[i]);
    if (JSON.stringify(reading(ourRun, as[i])) !== JSON.stringify(read)) reloaded.push(`${name} ${text}`);
    if (ds) {
      againstDocx4j++;
      const [o, d] = [written(ourRun), written(runIn(ds[i]))];
      for (const k of Object.keys(o)) if (o[k] !== d[k]) unlikeDocx4j.push(`${name} ${text} ${k}: ours ${o[k]}, docx4j's ${d[k]}`);
    }
  }
}
assert.deepEqual(unexpected, [], 'every marker reads as Word read it');
assert.deepEqual([...KNOWN].filter((k) => !found.has(k)), [], 'each known difference still differs');
assert.deepEqual(reloaded, [], 'a save and reload keeps every reading');
assert.deepEqual(unlikeDocx4j, [], "phase C writes each marker as docx4j's CR-061 does, value for value");
assert.deepEqual(notInForm, [], "every colour saved is in the schema's form");
assert.equal(markers, 388);
assert.equal(agree, 385);
assert.equal(againstDocx4j, 251);

// Each rule as Word saves it (docx4j's ColourNormalFormTest.eachRuleAsWordSavesIt): a name or hex written upper-case,
// exactly auto kept, a rejected text colour left out, a rejected underline colour left out, a rejected fill, border,
// pattern or page colour written auto.
const cases = [
  ['colour-case', 'K4_Red', 'w:val="FF0000"'], ['colour-case', 'K1_auto', 'w:val="auto"'],
  ['colour-hex', 'X1_ff0000', 'w:val="FF0000"'], ['colour-hex', 'X10_FF0000', 'w:val="FF0000"'],
  ['colour-styles', 'R9_AUTO', 'w:val="auto"'],
  ['colour-elsewhere', 'U1_UNDERLINE_RED', 'w:color="FF0000"'], ['colour-elsewhere', 'U3_UNDERLINE_F00', '<w:u w:val="single"/>'],
  ['colour-elsewhere', 'S2_SHADING_AQUA', 'w:fill="00FFFF"'], ['colour-elsewhere', 'S3_SHADING_F00', 'w:fill="auto"'],
  ['colour-elsewhere', 'B1_BORDER_RED', 'w:color="FF0000"'], ['colour-elsewhere', 'B2_BORDER_F00', 'w:color="auto"'],
  ['colour-pattern', 'P2_PATTERN_NAVY', 'w:color="000080"'], ['colour-pattern', 'P3_PATTERN_F00', 'w:color="auto"'],
  ['colour-pattern', 'P7_STYLE_F00', 'w:color="auto"'],
];
for (const [name, marker, expected] of cases) assert.ok(runXml(savedXml[name], marker).includes(expected), `${name} ${marker}: ${expected}`);
for (const [name, marker] of [['colour-hex', 'X3_F00'], ['colour-styles', 'R2_F00']]) assert.ok(!runXml(savedXml[name], marker).includes('<w:color'), `${name} ${marker}: left out`);
assert.ok(savedXml['colour-background-red'].includes('<w:background w:color="FF0000"/>'));
// Word drops a w:background whose colour it rejects; docx4j writes auto, and so does this package (no page colour either way).
assert.ok(savedXml['colour-background-f00'].includes('<w:background w:color="auto"/>'));

// The opt-out (docx4j's ColourNormalFormOptOutTest): per call and process-wide, colours as read, the repaired elements kept.
const asRead = async (name, options) => marshalString(await unmarshalString(source(name)), options);
const hexAsRead = await asRead('colour-hex', { normaliseOnSave: false });
assert.ok(runXml(hexAsRead, 'X1_ff0000').includes('w:val="ff0000"') && runXml(hexAsRead, 'X3_F00').includes('w:val="F00"'), 'per call');
assert.ok(runXml(await asRead('props-enum-missing', { normaliseOnSave: false }), 'M1_HIGHLIGHT_BARE').includes('<w:highlight/>'));
setNormaliseOnSave(false);
try {
  assert.ok(runXml(await asRead('colour-case'), 'K4_Red').includes('w:val="Red"'), 'process-wide');
  assert.ok(runXml(await asRead('colour-case', { normaliseOnSave: true }), 'K4_Red').includes('w:val="FF0000"'), "a call's setting wins");
} finally {
  setNormaliseOnSave(true);
}
assert.ok(!runXml(await asRead('colour-case'), 'K4_Red').includes('w:val="Red"'), 'on again');

console.log(`word365 probes (CR-013): ${PROBES.length} probes, ${markers} markers, ${agree} read as Word read them and ${KNOWN.size} known; `
  + `${readByWordsRule} values read by Word's rule; saved, ${normalisations} values normalised, every colour in the schema's form, `
  + `all ${markers} read the same again, and ${againstDocx4j} written as docx4j's saves write them; the opt-out OK`);
