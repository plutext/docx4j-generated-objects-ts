import { Jsonix } from '@docx4j/jsonix';

/**
 * Malformed XML refused in a browser (CR-023). A browser's `DOMParser` doesn't throw: it returns a document holding a
 * `parsererror` element, and read on, that document gives a tree cut off at the fault. xmldom throws, or for a few faults
 * recovers, which this leaves as it is.
 */

/** Chromium and WebKit's, then Firefox's. */
const PARSER_ERROR_NAMESPACES: ReadonlySet<string> = new Set([
  'http://www.w3.org/1999/xhtml',
  'http://www.mozilla.org/newlayout/xml/parsererror.xml',
]);

const firstElementOf = (node: Node): Element | undefined => {
  for (let child = node.firstChild; child; child = child.nextSibling) if (child.nodeType === 1) return child as Element;
  return undefined;
};

/**
 * The `parsererror` element of a document a browser failed to parse. It is the root in Firefox, the root's first element in
 * Chromium, and in `html` and `body` there when nothing parsed at all.
 */
function parserErrorOf(root: Element): Element | undefined {
  let element: Element | undefined = root;
  for (let depth = 0; element !== undefined && depth < 3; depth++, element = firstElementOf(element)) {
    if (element.localName === 'parsererror' && PARSER_ERROR_NAMESPACES.has(element.namespaceURI ?? '')) return element;
  }
  return undefined;
}

/** Throws if `node` is a browser's report of a parse that failed: a document, or the root element of one. */
export function refuseParserError(node: Node): void {
  const root = node.nodeType === 9 ? (node as Document).documentElement : node as Element;
  const error = root ? parserErrorOf(root) : undefined;
  if (error === undefined) return;
  // Chromium's message is its `div`. Firefox's is the text ahead of `sourcetext`, which repeats the source line.
  const message = error.getElementsByTagNameNS('http://www.w3.org/1999/xhtml', 'div')[0]?.textContent ?? error.firstChild?.nodeValue;
  throw new SyntaxError(`Malformed XML: ${message?.replace(/\s+/g, ' ').trim() || 'the parser reported an error'}`);
}

/** `Jsonix.DOM.parse`, refusing what a browser's parser reports rather than throws. */
export function parseXml(xml: string): Document {
  const doc = Jsonix.DOM.parse(xml);
  refuseParserError(doc);
  return doc;
}
