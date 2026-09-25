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
import { canonicaliseString, compareCanonically, differences, sameLine } from './lib/canonical.mjs';

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

  // CR-004 phase B: xs:all members have no order, and xsd:double has many spellings per value.
  const CORE = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
  const APP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
  same(`<cp:coreProperties xmlns:cp="${CORE}" xmlns:d="urn:d"><d:a>1</d:a><d:b>2</d:b></cp:coreProperties>`,
    `<cp:coreProperties xmlns:cp="${CORE}" xmlns:d="urn:d"><d:b>2</d:b><d:a>1</d:a></cp:coreProperties>`,
    'coreProperties is xsd:all: its children have no order');
  same(`<Properties xmlns="${APP}"><a>1</a><b>2</b></Properties>`, `<Properties xmlns="${APP}"><b>2</b><a>1</a></Properties>`,
    'extended Properties likewise');
  differs(`<cp:coreProperties xmlns:cp="${CORE}" xmlns:d="urn:d"><d:a>1</d:a><d:b>2</d:b></cp:coreProperties>`,
    `<cp:coreProperties xmlns:cp="${CORE}" xmlns:d="urn:d"><d:a>2</d:a><d:b>1</d:b></cp:coreProperties>`,
    'and a value swapped between two of them is still a difference: subtrees are sorted, not lines');
  differs(`<cp:coreProperties xmlns:cp="${CORE}" xmlns:d="urn:d"><d:a>1</d:a></cp:coreProperties>`,
    `<cp:coreProperties xmlns:cp="${CORE}" xmlns:d="urn:d"/>`, 'nor does the exception hide a dropped child');
  differs('<p><a/><b/></p>', '<p><b/><a/></p>', 'and it reaches only those two elements');

  same('<p v="1E-4"/>', '<p v="0.0001"/>', 'the same xsd:double written differently');
  same('<p v="46285.360326851849"/>', '<p v="46285.36032685185"/>', 'and at the precision Office writes');
  differs('<p v="1E-4"/>', '<p v="1E-3"/>', 'a different number is a difference');
  differs('<p v="1"/>', '<p v="1x"/>', 'and a value that is not a number is compared as text');
  differs('<p v="1" w="2"/>', '<p v="1"/>', 'a dropped attribute is not hidden by the numeric rule');

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

// ---------------------------------------------------------------- reading a part
//
// Fixtures are the archive's bytes, so a part is read by its own encoding rather than assumed to
// be UTF-8: Excel writes customXml/item1.xml of a Power Query workbook as UTF-16 LE with a BOM.
// Read as UTF-8 it arrives as mojibake and the parser fails, which would look like a model defect.
function readPart(file) {
  const bytes = readFileSync(file);
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  return bytes.toString('utf8').replace(/^\uFEFF/, '');
}

// Roots docx4j deliberately does not bind, so the model cannot type them and the facade keeps such
// a part as DOM (`unmarshalPackage` catches exactly this and passes the part through unchanged).
// Each names the docx4j ContentTypes constant that records the decision - the constant name rather
// than its comment, since docx4j keeps the names stable and rewords the comments. They are checked the opposite way round from
// everything else: the root must still be unknown, so if docx4j ever binds one this fails and the
// entry has to go - the same insistence as KNOWN, for a decision rather than a defect
// (core-ts CR-001 section 19).
const UNMODELLED = new Map([
  ['{http://schemas.microsoft.com/office/2020/mipLabelMetadata}labelList',
    'a sensitivity label; docx4j binds no schema for the namespace and has no ContentTypes constant for the part'],
  ['{http://schemas.microsoft.com/office/powerpoint/2018/8/main}authorLst',
    "PowerPoint's 2018 comment authors; docx4j ContentTypes.PRESENTATIONML_MODERN_COMMENT_AUTHORS, not bound (a DefaultXmlPart)"],
  ['{http://schemas.microsoft.com/office/powerpoint/2018/8/main}cmLst',
    "PowerPoint's 2018 modern comments; docx4j ContentTypes.PRESENTATIONML_MODERN_COMMENTS, not bound (a DefaultXmlPart)"],
  ['{http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments}personList',
    "the persons of Excel's threaded comments; docx4j ContentTypes.SPREADSHEETML_PERSONS, not bound (a DefaultXmlPart)"],
  ['{http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments}ThreadedComments',
    "Excel's threaded comments; docx4j ContentTypes.SPREADSHEETML_THREADED_COMMENTS, not bound (a DefaultXmlPart)"],
  ['{http://schemas.microsoft.com/DataMashup}DataMashup',
    "a Power Query blob (base64 in UTF-16), modelled by nobody"],
]);

const rootQNameOf = (xml) => {
  const root = new DOMParser().parseFromString(xml, 'text/xml').documentElement;
  return root.namespaceURI ? `{${root.namespaceURI}}${root.localName}` : root.localName;
};

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
/**
 * Attributes the model does not bind, so they are dropped from every element that carries them
 * (CR-004 phase B). Recorded by ATTRIBUTE rather than by part, because that is the shape of the
 * finding: one missing declaration costs the same attribute on 21 parts, and 21 entries saying the
 * same thing would rot separately.
 *
 * Kept honest the same way `KNOWN` is: an entry never observed during a run fails at the end, so a
 * declaration arriving upstream forces the entry out rather than passing unnoticed.
 */
const KNOWN_MISSING_ATTRIBUTES = new Map([
  ['{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}uid',
    'xr:uid is declared on CT_Worksheet alone, so it survives on a worksheet root and is dropped from '
    + 'autoFilter, hyperlink, table, pivotCacheDefinition, pivotTableDefinition and comment. [plutext/docx4j]'],
  ['{http://schemas.microsoft.com/office/spreadsheetml/2015/revision2}uid',
    'xr2:uid on workbookView: the loss docx4j has logged for xlsx4j, seen on every workbook here. [plutext/docx4j]'],
  ['{http://schemas.microsoft.com/office/spreadsheetml/2016/revision3}uid',
    'xr3:uid on tableColumn; docx4j binds no schema for the 2016/revision3 namespace. [plutext/docx4j]'],
  ['{http://schemas.microsoft.com/office/spreadsheetml/2017/revision16}uid',
    'xr16:uid on connection; docx4j binds no schema for the 2017/revision16 namespace. [plutext/docx4j]'],
  ['Version',
    'Version on b:Sources: CT_Sources in shared-bibliography.xsd declares SelectedStyle, StyleName and '
    + 'URI only, and Word writes Version="6". [plutext/docx4j]'],
]);

const seenMissing = new Set();

/**
 * The attributes of `KNOWN_MISSING_ATTRIBUTES` that account for the whole of a differing line, or
 * null if anything else differs too. Striking them from Office's line and comparing what is left
 * makes the rules compose - a line may lose an xr:uid AND respell an xsd:double - while keeping the
 * classification honest: a line that loses a recorded attribute and something else is not
 * explained, and fails. This is the code most able to hide a real loss, so it is checked below.
 */
export function explainedByMissingAttributes(office, ours) {
  const known = droppedAttributes(office, ours).filter((a) => KNOWN_MISSING_ATTRIBUTES.has(a));
  if (known.length === 0) return null;
  // Rebuild Office's line without the recorded attributes and compare what is left, structurally
  // rather than by pattern: an attribute name here is a QName in braces, which a regexp would have
  // to escape, and getting that wrong would silently explain everything.
  const head = (line) => line.slice(0, line.length - attributeTokens(line).join(' ').length).trimEnd();
  const remainder = `${head(office)} ${attributeTokens(office)
    .filter((token) => !known.includes(token.slice(0, token.indexOf('='))))
    .join(' ')}`.trimEnd();
  return sameLine(remainder, ours) ? known : null;
}

/** The attributes a canonical line has that its counterpart does not. */
const attributeTokens = (line) => (line ?? '').match(/(?:\{[^}]*\})?[\w:.-]+="(?:[^"\\]|\\.)*"/g) ?? [];

const droppedAttributes = (office, ours) => {
  const theirs = new Set(attributeTokens(ours));
  return attributeTokens(office).filter((a) => !theirs.has(a)).map((a) => a.slice(0, a.indexOf('=')));
};

{
  // The classifier that decides a difference is a recorded one. Over-permissive here means a real
  // loss reported as known, which is the failure mode this whole test exists to prevent.
  const XR = '{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}uid';
  const line = (attrs) => `  {ns}table name="T" ${attrs}`.trimEnd();
  assert.deepEqual(explainedByMissingAttributes(line(`${XR}="{A}" ref="A1"`), line('ref="A1"')), [XR],
    'a line losing only a recorded attribute is explained');
  assert.equal(explainedByMissingAttributes(line(`${XR}="{A}" ref="A1"`), line('ref="A2"')), null,
    'a line losing a recorded attribute AND changing another value is NOT explained');
  assert.equal(explainedByMissingAttributes(line(`${XR}="{A}" ref="A1" displayName="D"`), line('ref="A1"')), null,
    'nor one that also drops an unrecorded attribute');
  assert.equal(explainedByMissingAttributes(line('ref="A1" displayName="D"'), line('ref="A1"')), null,
    'an unrecorded attribute alone is never explained');
  assert.deepEqual(explainedByMissingAttributes(line(`${XR}="{A}" v="1E-4"`), line('v="0.0001"')), [XR],
    'the rules compose: a recorded loss and an xsd:double respelling together');

  console.log('fidelity: classifier self-checks OK');
}

const KNOWN = new Map([
]);

const failures = [];
const known = [];
const notModelled = [];
let checked = 0;

for (const file of parts) {
  const name = path.relative(root, file).split(path.sep).join('/');
  const source = readPart(file);
  const cases = [[name, source]];
  const resolve = RESOLVED.get(name);
  if (resolve) {
    const { xml, count } = takeChoice(source, resolve);
    assert.ok(count > 0, `${name}: no mc:AlternateContent to resolve; the fixture or the list is stale`);
    assert.ok(canonicaliseString(xml).length > 0);
    cases.push([`${name} [mc:Choice ${resolve.join(' ')} taken]`, xml]);
  }

  const unmodelled = UNMODELLED.get(rootQNameOf(source));
  if (unmodelled) {
    notModelled.push(`${name}: ${unmodelled}`);
    await assert.rejects(unmarshalString(source), /not known in this context/,
      `${name} is recorded as unmodelled, but the model now knows its root: remove the entry`);
    continue;
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
    const all = differences(xml, out);
    // Every differing line is classified, not only the first: a line that differs solely by an
    // attribute the model is known not to bind is that finding, recorded once per attribute; the
    // rest are failures. Stopping at the first would let a recorded loss mask an unrecorded one.
    const unexplained = [];
    for (const difference of all.differences) {
      const explained = explainedByMissingAttributes(difference.before, difference.after);
      if (explained) for (const a of explained) seenMissing.add(a);
      else unexplained.push(difference);
    }
    const report = { ...compareCanonically(xml, out), ...(unexplained[0] ?? {}), equal: unexplained.length === 0 };
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

for (const attribute of seenMissing) console.log(`fidelity: not bound, so dropped: ${attribute}\n  ${KNOWN_MISSING_ATTRIBUTES.get(attribute)}`);
for (const [attribute, why] of KNOWN_MISSING_ATTRIBUTES) {
  if (!seenMissing.has(attribute)) failures.push(`${attribute}\n    is recorded as not bound, but no part loses it any more`
    + `\n    (${why})\n    Remove the entry, and check the attribute now round-trips.`);
}
for (const entry of known) console.log(`fidelity: known difference, ${entry}`);
if (notModelled.length) console.log(`fidelity: ${notModelled.length} parts are not modelled by design and stay DOM:\n`
  + notModelled.map((e) => `  ${e}`).join('\n'));

if (failures.length) {
  console.error(`\nfidelity: ${failures.length} of ${checked} checks lost something in a round trip:\n`);
  for (const failure of failures) console.error(`  ${failure}\n`);
  process.exitCode = 1;
} else {
  console.log(`fidelity: ${checked} parts Office wrote round-trip through the facade unchanged`
    + (known.length ? `, but for ${known.length} known difference${known.length === 1 ? '' : 's'} above` : ''));
}
