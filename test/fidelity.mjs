// CR-004 phase A: round-trip fidelity over parts Office 365 wrote.
//
// For each fixture part: unmarshal and marshal it through the facade (not `createMarshaller`, so
// CR-001's namespace rules are exercised too), and compare canonically. A part that throws is a
// failure with its message: Office wrote it and the model claims to cover it.
//
// Every loss docx4j's CR-021 to CR-026 closed was of this kind - invisible in the declarations, and
// invisible to a round trip of a document this repository wrote itself.
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import assert from 'node:assert/strict';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { unmarshalString, marshalString } from '../dist/index.mjs';
import { canonicaliseString, compareCanonically } from './lib/canonical.mjs';

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const root = fileURLToPath(new URL('./fixtures/fidelity/', import.meta.url));

// ---------------------------------------------------------------- the canonicaliser itself
//
// A difference the canonicaliser hides is worse than one it reports, so it is checked first, on
// both sides: what it must call equal, and what it must not.
{
  const same = (a, b, why) => assert.ok(compareCanonically(a, b).equal, `should be equal: ${why}`);
  const differs = (a, b, why) => assert.equal(compareCanonically(a, b).equal, false, `should differ: ${why}`);

  same('<a:p xmlns:a="urn:x"><a:r/></a:p>', '<p xmlns="urn:x"><r/></p>', 'prefixes are not content');
  same('<p x="1" a="2"/>', '<p a="2" x="1"/>', 'attribute order');
  same('<p xmlns:w="urn:w" w:val="x"/>', '<p xmlns:z="urn:w" z:val="x"/>', 'attribute prefixes');
  same('<p><r i="1"/></p>', '<p><r i="true"/></p>', 'boolean attribute spellings');
  same('<p>\n  <r/>\n</p>', '<p><r/></p>', 'whitespace between elements');
  same('<p xmlns:mc="urn:mc" mc:Ignorable="w14 w15"/>', '<p xmlns:mc="urn:mc" mc:Ignorable="w14 w15"/>', 'Ignorable as a value');

  differs('<p><a/><b/></p>', '<p><b/><a/></p>', 'element order: the declarations promise schema order');
  differs('<p><a/></p>', '<p/>', 'a dropped child');
  differs('<p a="1"/>', '<p/>', 'a dropped attribute');
  same('<p styleId="1"/>', '<p styleId="true"/>', 'any 1/true pair: only the model can spell one value two ways (see canonical.mjs)');
  differs('<p a="1"/>', '<p a="false"/>', 'a value that actually changed');
  differs('<p><t xml:space="preserve">a </t></p>', '<p><t xml:space="preserve">a</t></p>', 'preserved trailing space');
  differs('<p><t>text</t></p>', '<p><t/></p>', 'dropped text');
  differs('<p xmlns="urn:a"/>', '<p xmlns="urn:b"/>', 'a different namespace');

  // The line numbers and counts a failure reports must point at the difference.
  const report = compareCanonically('<p><a/><b/></p>', '<p><a/></p>');
  assert.equal(report.line, 3);
  assert.deepEqual(report.counts, { before: 3, after: 2 });
  console.log('fidelity: canonicaliser self-checks OK');
}

// ---------------------------------------------------------------- mc:AlternateContent resolution
//
// A part as Office wrote it holds `mc:AlternateContent` unresolved, and the model keeps it whole
// (docx4j CR-021, CR-024), so a round trip of it never enters the branch. A consumer that resolves
// markup compatibility first - which `@docx4j/core-ts` does on every part - unmarshals the chosen
// branch instead, and that is where 0.1.5's regression lived: a14:m typed, and OMML's CT_R rejecting
// the a:rPr DrawingML text carries. So the parts that carry one are also checked resolved.
//
// This is the reader's half of ECMA-376 Part 3's rule, done on the DOM rather than in the model,
// since the point is to unmarshal what a preprocessing consumer would hand us.
function takeChoice(xml, understood) {
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const alternates = [...doc.getElementsByTagNameNS(MC, 'AlternateContent')];
  for (const alternate of alternates) {
    const branches = [...alternate.childNodes].filter((n) => n.nodeType === 1);
    const chosen = branches.find((b) => b.localName === 'Choice'
      && b.getAttribute('Requires').split(/\s+/).every((p) => understood.includes(p)))
      ?? branches.find((b) => b.localName === 'Fallback')
      ?? branches[0];
    // The declarations a branch carries (Office puts xmlns:a14 on the Choice) go with its children,
    // since the element that declared them is about to be dropped.
    const declarations = [...chosen.attributes].filter((a) => a.name === 'xmlns' || a.prefix === 'xmlns');
    for (const child of [...chosen.childNodes].filter((n) => n.nodeType === 1)) {
      for (const d of declarations) child.setAttribute(d.name, d.value);
      alternate.parentNode.insertBefore(child, alternate);
    }
    alternate.parentNode.removeChild(alternate);
  }
  return { xml: new XMLSerializer().serializeToString(doc), count: alternates.length };
}

// ---------------------------------------------------------------- the fixtures
const parts = [];
const walk = (dir) => {
  for (const entry of readdirSync(dir)) {
    const full = path.join(dir, entry);
    if (statSync(full).isDirectory()) walk(full);
    else if (entry !== 'SOURCE.md') parts.push(full);
  }
};
walk(root);
parts.sort();
assert.ok(parts.length >= 17, `expected the phase A fixtures, found ${parts.length}`);

// Parts whose `mc:AlternateContent` is also checked resolved, with the prefixes a reader that
// understands this package's namespaces would accept. Word and PowerPoint write a14 (equations and
// shapes in DrawingML text); the list is the branch's `Requires`, not every prefix we know.
const RESOLVED = new Map([
  ['cr022-slicers-timelines.xlsx/xl/drawings/drawing1.xml', ['a14']],
  ['loadAndSave.pptx/ppt/slides/slide2.xml', ['a14']],
  ['loadAndSave.xlsx/xl/drawings/drawing1.xml', ['a14']],
]);

/**
 * Differences that are known, understood, and owned elsewhere (CR-004 open question 2).
 *
 * The device is written up in plutext/docx4j-core-ts CR-001 section 19: a dependency's defect is
 * recorded as a test that asserts the broken behaviour, so the fix announces itself by failing,
 * where a skip would be silent in both directions. It is for a DEPENDENCY's defect, never one of
 * ours, and an entry without its reason and upstream reference decays into the exclusion it
 * replaced.
 *
 * The test is otherwise strict: any difference fails. An entry here does not soften that, it
 * records a specific difference and its owner - and it is matched exactly, so the part still fails
 * if the difference changes shape, AND it fails when the part becomes identical, which is how the
 * entry gets removed when someone fixes it. An expectations file that quietly absorbs whatever it
 * finds would defeat the test.
 */
const KNOWN = new Map([
  ['cr022-checkbox.xlsx/xl/workbook.xml', {
    why: 'xr2:uid on workbookView is not in the model; docx4j has logged it as an xlsx4j CR of its own.',
    owner: 'plutext/docx4j (xlsx4j)',
    before: '{http://schemas.microsoft.com/office/spreadsheetml/2015/revision2}uid="{00000000-000D-0000-FFFF-FFFF00000000}"',
    after: '(absent)',
  }],
]);

const failures = [];
const known = [];
let checked = 0;

for (const file of parts) {
  const name = path.relative(root, file).split(path.sep).join('/');
  const source = readFileSync(file, 'utf8');
  const cases = [[name, source]];
  const resolve = RESOLVED.get(name);
  if (resolve) {
    const { xml, count } = takeChoice(source, resolve);
    assert.ok(count > 0, `${name}: no mc:AlternateContent to resolve; the fixture or the list is stale`);
    assert.ok(canonicaliseString(xml).length > 0);
    cases.push([`${name} [mc:Choice ${resolve.join(' ')} taken]`, xml]);
  }

  for (const [label, xml] of cases) {
    checked++;
    let out;
    try {
      out = await marshalString(await unmarshalString(xml));
    } catch (error) {
      // A throw can be a recorded difference too: a part Office wrote that the model cannot read
      // is the same finding as one it reads lossily, and wants the same owner and the same
      // insistence. Matched on the message, so a part that starts throwing differently fails.
      const expectedThrow = KNOWN.get(label)?.throws;
      if (expectedThrow && error.message.includes(expectedThrow)) known.push(`${label}: ${KNOWN.get(label).why} [${KNOWN.get(label).owner}]`);
      else failures.push(`${label}\n    threw: ${error.message.split('\n')[0]}`
        + (expectedThrow ? `\n    (KNOWN expects a different throw here: ${expectedThrow})` : ''));
      continue;
    }
    const report = compareCanonically(xml, out);
    const expected = KNOWN.get(label);
    if (report.equal) {
      // A known difference that has gone: the entry is stale and must be removed, or a canonicalisation
      // rule has started hiding it. Either way it needs a person, so it fails. A recorded throw that
      // no longer throws arrives here too, which is how a fix upstream forces the record out.
      if (expected) failures.push(`${label}\n    round-trips identically now, but KNOWN still records a difference`
        + `\n    (${expected.owner}): ${expected.why}\n    Remove the entry.`);
      continue;
    }
    // A known difference is matched on what actually differs, not merely on the part's name.
    const matches = expected
      && report.before.includes(expected.before === '(absent)' ? expected.after : expected.before)
      && (expected.after === '(absent)' ? !report.after.includes(expected.before) : report.after.includes(expected.after));
    if (matches) {
      known.push(`${label}: ${expected.why} [${expected.owner}]`);
    } else {
      failures.push(`${label}\n    line ${report.line} of ${report.counts.before} (${report.counts.after} out)`
        + `\n    Office: ${report.before}\n    ours  : ${report.after}`
        + (expected ? `\n    (KNOWN records a different difference here: ${expected.why})` : ''));
    }
  }
}

for (const entry of known) console.log(`fidelity: known difference, ${entry}`);

if (failures.length) {
  console.error(`\nfidelity: ${failures.length} of ${checked} checks lost something in a round trip:\n`);
  for (const failure of failures) console.error(`  ${failure}\n`);
  process.exitCode = 1;
} else {
  console.log(`fidelity: ${checked} parts Office wrote round-trip through the facade unchanged`
    + (known.length ? `, but for ${known.length} known difference${known.length === 1 ? '' : 's'} above` : ''));
}
