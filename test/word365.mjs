// Runtime check of CR-013 against Word 365: Jason's run property probes (2026-10-24), each read as Word read it.
// test/fixtures/word365-probes/ holds each probe's word/document.xml and Word 365's save of it (SOURCE.md there).
// Every probe paragraph after the title is "label: MARKER (end)", three runs, and only the MARKER run carries the
// property tested. Word's save is the oracle: Word writes back only values it reads one way, so a marker's reading
// in the probe must equal its reading in the save, as docx4j's RunPropertiesWord365Test has it. The readings are
// of direct formatting (this package resolves no styles), so where Word's save changes direct formatting without
// changing what Word shows, or changes what Word shows, the difference is stated below.
import { readFileSync, readdirSync } from 'node:fs';
import assert from 'node:assert/strict';
import { unmarshalString, marshalString, unwrap } from '../dist/index.mjs';
import { textOf, readRunOptions } from '../dist/builders/wml.mjs';
import { wordColor } from '../dist/helpers/wml.mjs';

const DIR = new URL('./fixtures/word365-probes/', import.meta.url);
const PROBES = readdirSync(DIR).filter((d) => d.endsWith('.docx') && !d.endsWith('-word365.docx')).map((d) => d.slice(0, -'.docx'.length)).sort();
assert.equal(PROBES.length, 28, 'rounds 1 to 6');

const TOGGLES = ['b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'vanish'];

/**
 * A marker run's direct formatting, as far as the probes test it. A valueless w:u is unset, as Word reads it (its
 * save drops one). Where neither the run nor its paragraph has a style, a direct value equal to the default (w:u
 * none, w:color auto) is unset too: Word's save drops it, and with nothing beneath they read alike (the
 * unstyled probes have no styles part).
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
  return m;
}

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

const paragraphs = (element) => unwrap(element).body.content.map(unwrap).filter((v) => v.TYPE_NAME === 'org_docx4j_wml.P');
const runs = (p) => p.content.map(unwrap).filter((v) => v.TYPE_NAME === 'org_docx4j_wml.R');
const load = async (name) => {
  const readings = [];
  const element = await unmarshalString(readFileSync(new URL(`${name}.docx/word/document.xml`, DIR), 'utf8'), { onWordReading: (r) => readings.push(r) });
  return { element, readings };
};

let markers = 0;
let agree = 0;
let readByWordsRule = 0;
const found = new Set();
const unexpected = [];
const reloaded = [];
for (const name of PROBES) {
  // Before CR-013 phase B every probe with a value Word reads and the runtime refuses failed to load at all.
  const probe = await load(name);
  const saved = await load(`${name}-word365`);
  readByWordsRule += probe.readings.length;
  assert.equal(saved.readings.length, 0, `${name}: Word's save writes only values the runtime reads`);
  // A save through the facade, read again, reads as the probe did.
  const again = paragraphs(await unmarshalString(await marshalString(probe.element)));
  const [ps, ss] = [paragraphs(probe.element), paragraphs(saved.element)];
  assert.equal(ss.length, ps.length, `${name}: Word keeps the paragraphs one for one`);
  for (let i = 1; i < ps.length; i++) {
    const rs = runs(ps[i]);
    if (rs.length !== 3) continue;
    const text = textOf(rs[1]);
    // Word merges runs whose formatting is equal, so in its save the marker can share a run with its label.
    const savedRun = runs(ss[i]).filter((r) => textOf(r).includes(text)).at(-1);
    assert.ok(savedRun, `${name} ${text}: in Word's save`);
    markers++;
    const ours = reading(rs[1], ps[i]);
    const word = { ...reading(savedRun, ss[i]), ...WORD_DREW[`${name} ${text}`] };
    const differing = Object.keys(word).filter((k) => ours[k] !== word[k]);
    for (const k of differing) {
      const key = `${name} ${text} ${k}`;
      if (KNOWN.has(key)) found.add(key);
      else unexpected.push(`${key}: ours ${ours[k]}, Word's ${word[k]}`);
    }
    if (differing.length === 0) agree++;
    if (JSON.stringify(reading(runs(again[i])[1], again[i])) !== JSON.stringify(ours)) reloaded.push(`${name} ${text}`);
  }
}
assert.deepEqual(unexpected, [], 'every marker reads as Word read it');
assert.deepEqual([...KNOWN].filter((k) => !found.has(k)), [], 'each known difference still differs');
assert.deepEqual(reloaded, [], 'a save and reload keeps every reading');
assert.equal(markers, 376);
assert.equal(agree, 373);
console.log(`word365 probes (CR-013): ${PROBES.length} probes, ${markers} markers, ${agree} read as Word read them and ${KNOWN.size} known; `
  + `all ${markers} the same after a save and reload; ${readByWordsRule} values read by Word's rule OK`);
