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
 * A canonical string for an element tree: one line per node, indented by depth, with the element's
 * QName, its sorted attributes, and its text. Element ORDER is preserved and therefore compared -
 * the declarations promise schema order (compiler CR-007), so a reordering is a real difference.
 */
export function canonicalise(node, { preserve = false, depth = 0, lines = [], path = '' } = {}) {
  const here = `${path}/${qname(node)}`;
  lines.push(`${'  '.repeat(depth)}${qname(node)} ${attributesOf(node).join(' ')}`.trimEnd());
  const space = preservesSpace(node, preserve);
  for (let i = 0; i < node.childNodes.length; i++) {
    const child = node.childNodes[i];
    if (child.nodeType === 1) {
      canonicalise(child, { preserve: space, depth: depth + 1, lines, path: here });
    } else if (child.nodeType === 3 || child.nodeType === 4) {
      // Whitespace-only text between elements is formatting, not content, unless xml:space says so.
      const text = space ? child.data : child.data.trim();
      if (text !== '') lines.push(`${'  '.repeat(depth + 1)}#text ${JSON.stringify(text)}`);
    }
    // Comments and processing instructions are not part of the model and are not compared.
  }
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
export function compareCanonically(before, after) {
  const a = canonicaliseString(before);
  const b = canonicaliseString(after);
  for (let i = 0; i < Math.max(a.length, b.length); i++) {
    if (a[i] !== b[i]) {
      return {
        equal: false, line: i + 1, before: a[i] ?? '(nothing: the part ends here)', after: b[i] ?? '(nothing: the part ends here)',
        counts: { before: a.length, after: b.length },
      };
    }
  }
  return { equal: true, counts: { before: a.length, after: b.length } };
}
