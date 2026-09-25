// Canonical form of an XML part, for the round-trip fidelity comparison (CR-004 section 2.1).
//
// Each rule here is a decision the CR records, not a convenience: it says what is NOT a difference
// between what Office wrote and what this package writes back. A difference this hides is worse
// than one it reports, so the rules are narrow, and `canonicalise` reports the path of every node
// it normalises a value at, so the runner can show where a difference is.
import { DOMParser } from '@xmldom/xmldom';

const XML_NS = 'http://www.w3.org/XML/1998/namespace';

/** `{ns}local`, or `local` in no namespace: a prefix is a serialisation choice, not content. */
export const qname = (node) => (node.namespaceURI ? `{${node.namespaceURI}}${node.localName}` : node.localName);

/**
 * `xs:boolean` spells true as `1` or `true` and false as `0` or `false`, and the model writes back
 * the spelling of its own choosing, so `1` and `true` are the same value here.
 *
 * This normalises the lexical form of EVERY attribute rather than a list of the ones the schemas
 * type as boolean, which is safe in this comparison and not in general: the two sides are a part
 * Office wrote and the same part marshalled back, so an attribute the model does not type as
 * boolean is copied through verbatim and cannot present two spellings. A difference is therefore
 * only ever hidden where the model's own boolean typing created it - which is exactly the case the
 * rule is for. (A hand-kept list was the first attempt; it silently missed `noGrp`,
 * `noChangeAspect`, `anchorCtr`, `knownFonts` and more, which is the failure mode to avoid in a
 * canonicaliser.)
 */
const normaliseBoolean = (value) => (value === '1' ? 'true' : value === '0' ? 'false' : value);

/**
 * Whether two attribute values are the same `xsd:double` written differently: Office writes
 * `1E-4` and `46285.360326851849`, the model holds a JS number and writes `0.0001` and
 * `46285.36032685185`. The lexical space of `xsd:double` is many-to-one, so this is a spelling,
 * not a value - but only where BOTH sides parse as finite numbers and compare `===`, and only
 * where the two strings differ in the first place.
 *
 * What it hides, and all it hides: a string-typed attribute re-spelled into a numerically equal
 * form (a `ST_Xstring` cell value `"1.0"` becoming `"1"`, which Excel would see as a change). That
 * cannot arise between a part and its own re-marshal unless the model typed the attribute as a
 * number, which is the case the rule is for - the same argument as the boolean rule above.
 */
const sameNumber = (a, b) => {
  if (a === b || a === '' || b === '') return false;
  const x = Number(a), y = Number(b);
  return Number.isFinite(x) && Number.isFinite(y) && x === y;
};

/**
 * The attributes of an element, as sorted `qname=value` strings. Namespace declarations are
 * dropped: which declarations a root carries is CR-001's subject and has its own tests, and
 * `mc:Ignorable` is compared as an ordinary attribute value (it names prefixes, so it is compared
 * as the text Office wrote).
 */
const attributesOf = (element) => {
  const out = [];
  for (let i = 0; i < element.attributes.length; i++) {
    const attribute = element.attributes[i];
    if (attribute.name === 'xmlns' || attribute.prefix === 'xmlns') continue;
    const name = attribute.namespaceURI && attribute.namespaceURI !== XML_NS
      ? `{${attribute.namespaceURI}}${attribute.localName}`
      : attribute.namespaceURI === XML_NS ? `xml:${attribute.localName}` : attribute.localName;
    const value = normaliseBoolean(attribute.value);
    out.push(`${name}=${JSON.stringify(value)}`);
  }
  return out.sort();
};

/**
 * Whether an element's text is compared exactly including whitespace. `xml:space="preserve"` says
 * so, and the flag is inherited by descendants, as the XML specification has it.
 */
const preservesSpace = (element, inherited) => {
  const own = element.getAttributeNS?.(XML_NS, 'space');
  return own ? own === 'preserve' : inherited;
};

/**
 * Elements whose children have no order, so a reordering is not a difference (CR-004 phase B).
 *
 * The general rule is the opposite - element order IS compared, because the generated declarations
 * promise schema order (compiler CR-007) - but that premise holds only for `xsd:sequence`. These
 * two are `xsd:all` in docx4j's schemas (`opc-coreProperties.xsd`, and `Properties` in
 * `shared-documentPropertiesExtended.xsd`), where every order is equally valid. Jsonix has no
 * representation for an unordered model, so the compiler flattens the members into an ordered
 * property list and the marshaller writes that order; no order it could write would match the
 * corpus, because Word, Excel and PowerPoint each write `docProps/app.xml` differently (Word puts
 * `Application` after `Characters`, Excel first, PowerPoint third). There is nothing to fix in the
 * schema or the facade, which is why this is a canonicalisation rule and not a bug.
 *
 * Deliberately two QNames and not a category: `docProps/custom.xml` is an `xsd:sequence` of
 * repeated `property` elements, where order is meaningful, and must never be added here.
 */
const ORDER_FREE = new Set([
  '{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}coreProperties',
  '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Properties',
]);

/**
 * A canonical string for an element tree: one line per node, indented by depth, with the element's
 * QName, its sorted attributes, and its text. Element ORDER is preserved and therefore compared -
 * the declarations promise schema order (compiler CR-007), so a reordering is a real difference -
 * except under an `ORDER_FREE` element, where the children's whole subtrees are sorted.
 *
 * Sorting subtrees rather than lines is what keeps that exception narrow: a child keeps its own
 * descendants and attributes, so a value moved between two differently named siblings still shows
 * as a difference. What the exception does hide, and all it hides, is a reordering of those
 * children.
 */
export function canonicalise(node, { preserve = false, depth = 0, lines = [], path = '' } = {}) {
  const here = `${path}/${qname(node)}`;
  lines.push(`${'  '.repeat(depth)}${qname(node)} ${attributesOf(node).join(' ')}`.trimEnd());
  const space = preservesSpace(node, preserve);
  const orderFree = ORDER_FREE.has(qname(node));
  const blocks = [];
  for (let i = 0; i < node.childNodes.length; i++) {
    const child = node.childNodes[i];
    if (child.nodeType === 1) {
      if (orderFree) blocks.push(canonicalise(child, { preserve: space, depth: depth + 1, path: here }));
      else canonicalise(child, { preserve: space, depth: depth + 1, lines, path: here });
    } else if (child.nodeType === 3 || child.nodeType === 4) {
      // Whitespace-only text between elements is formatting, not content, unless xml:space says so.
      const text = space ? child.data : child.data.trim();
      if (text !== '') lines.push(`${'  '.repeat(depth + 1)}#text ${JSON.stringify(text)}`);
    }
    // Comments and processing instructions are not part of the model and are not compared.
  }
  if (orderFree) for (const block of blocks.sort((a, b) => (a.join('\n') < b.join('\n') ? -1 : 1))) lines.push(...block);
  return lines;
}

/** Parse and canonicalise a part, as an array of lines. */
export function canonicaliseString(xml) {
  const errors = [];
  const doc = new DOMParser({
    onError: (level, message) => { if (level !== 'warning') errors.push(`${level}: ${message}`); },
  }).parseFromString(xml, 'text/xml');
  if (errors.length) throw new Error(`could not parse: ${errors[0]}`);
  return canonicalise(doc.documentElement);
}

/**
 * Compare two parts canonically. Returns `{ equal }`, and on a difference the line number and the
 * two lines, plus the element counts, so a failure names what moved rather than only that it did.
 */
/** Two canonical lines differing only in the lexical form of numeric attribute values. */
export const sameLine = (a, b) => {
  if (a === b) return true;
  if (a === undefined || b === undefined) return false;
  const values = (line) => [...line.matchAll(/="((?:[^"\\]|\\.)*)"/g)].map((m) => m[1]);
  const [x, y] = [values(a), values(b)];
  if (x.length !== y.length || x.length === 0) return false;
  // The line with every value blanked must match, so only the values may differ.
  const blank = (line) => line.replace(/="(?:[^"\\]|\\.)*"/g, '=""');
  if (blank(a) !== blank(b)) return false;
  return x.every((value, i) => value === y[i] || sameNumber(value, y[i]));
};

/**
 * Every differing line between two parts, not merely the first. The first-difference form was
 * enough while a part had one finding; once findings are recorded by class (an attribute the model
 * does not bind, say), stopping at the first would let a known difference mask an unknown one
 * further down - the masking this test exists to prevent.
 */
export function differences(before, after) {
  const a = canonicaliseString(before);
  const b = canonicaliseString(after);
  const out = [];
  for (let i = 0; i < Math.max(a.length, b.length); i++) {
    if (!sameLine(a[i], b[i])) {
      out.push({ line: i + 1, before: a[i] ?? '(nothing: the part ends here)', after: b[i] ?? '(nothing: the part ends here)' });
    }
  }
  return { differences: out, counts: { before: a.length, after: b.length } };
}

export function compareCanonically(before, after) {
  const a = canonicaliseString(before);
  const b = canonicaliseString(after);
  for (let i = 0; i < Math.max(a.length, b.length); i++) {
    if (!sameLine(a[i], b[i])) {
      return {
        equal: false, line: i + 1, before: a[i] ?? '(nothing: the part ends here)', after: b[i] ?? '(nothing: the part ends here)',
        counts: { before: a.length, after: b.length },
      };
    }
  }
  return { equal: true, counts: { before: a.length, after: b.length } };
}
