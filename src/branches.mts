// CR-018: a content control directly in a kept mc:AlternateContent branch, typed at its host's level. A branch's content
// is a lax wildcard, so the runtime types each child by its global declaration, and the global w:sdt is the block-level
// class: a run-level control in a paragraph's mc:Choice read as SdtBlock, with everything below it. ECMA-376 Part 3 has
// a reader take a branch's content as if it stood in the host, and docx4j binds such a control by its container's kind
// (CR-051 section 11.3, McBranchSdt). The facade answers jsonix-CR-014's `wildcardTypeInfo` (3.11.0) with this class, on
// read and on save, so the control is read by the host's lists and written with no xsi:type. Internal, not an exported
// path.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

export type Level = 'block' | 'run' | 'row' | 'cell';

/**
 * The level an element gives a content control below it: every WordprocessingML element whose list declares `w:sdt`,
 * by the class it declares there, but `w:sdtContent` and `w:customXml`, whose level is their host's (measured from the
 * context, 2026-10-28; the smoke holds it to the context).
 */
export const HOST_LEVEL: Readonly<Record<string, Level>> = {
  body: 'block', docPartBody: 'block', tc: 'block', txbxContent: 'block', hdr: 'block', ftr: 'block', footnote: 'block', endnote: 'block', comment: 'block',
  p: 'run', hyperlink: 'run', fldSimple: 'run', smartTag: 'run', ins: 'run', del: 'run', moveFrom: 'run', moveTo: 'run', dir: 'run', bdo: 'run',
  tbl: 'row',
  tr: 'cell',
};

/** The class of `w:sdt` at each level, as the host's list declares it. */
export const SDT_AT_LEVEL: Readonly<Record<Level, string>> = {
  block: 'org_docx4j_wml.SdtBlock', run: 'org_docx4j_wml.SdtRun', row: 'org_docx4j_wml.CTSdtRow', cell: 'org_docx4j_wml.CTSdtCell',
};

/**
 * The class a `w:sdt` directly in an `mc:Choice` or `mc:Fallback` takes: its host's, found by walking up from the branch
 * past the `mc` elements and any `w:sdt`, `w:sdtContent` and `w:customXml` to the first element in `HOST_LEVEL`.
 * Undefined for any other element, and where no ancestor gives a level (the runtime's own lookup then applies). `parent`
 * is the element being read's parent in the source, or the parent being written, as jsonix-CR-014's W1 gives it.
 */
export function branchSdtType(name: { readonly namespaceURI: string; readonly localPart: string }, parent: Node | null | undefined): string | undefined {
  if (name.namespaceURI !== W || name.localPart !== 'sdt' || !parent || parent.nodeType !== 1) return undefined;
  const branch = parent as Element;
  if (branch.namespaceURI !== MC || (branch.localName !== 'Choice' && branch.localName !== 'Fallback')) return undefined;
  let n: Node | null = branch;
  while (n !== null && n.nodeType === 1) {
    const e = n as Element;
    n = e.parentNode;
    if (e.namespaceURI === MC) continue;
    if (e.namespaceURI !== W) return undefined;
    if (e.localName === 'sdt' || e.localName === 'sdtContent' || e.localName === 'customXml') continue;
    const level = HOST_LEVEL[e.localName];
    return level === undefined ? undefined : SDT_AT_LEVEL[level];
  }
  return undefined;
}
