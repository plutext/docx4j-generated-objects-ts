// CR-022: a child given more than once where the model holds one, read as Word 365 reads it (WD REQ-322). The runtime keeps the
// last copy of a single-valued property, so a first w:body's words, a first w:pPr's properties and a first w:sectPr's page size are
// lost, where Word shows them. jsonix-CR-016's onOverwrittenElement (3.13.0) reports each overwrite while the earlier copy is still
// held, and wordOverwritten edits the copy the runtime keeps. The counterpart of docx4j's RepeatedChildren (17.3.6, af50deffe and
// 8c23c444f, CR-079 phases 2 and 3), rule for rule, measured by the same Word saves. Internal, not an exported path; the facade
// exports wordOverwritten.

import { Jsonix } from '@docx4j/jsonix';

const WML = 'org_docx4j_wml.';

type Typed = { TYPE_NAME?: string; [property: string]: unknown };

/**
 * The containers joined (holder class, property), and where the earlier copy's content goes: at the head of the later's (`w:body`
 * and `w:sdtContent` at every level, in order) or at its tail (`w:txbxContent` in a VML box, the later's content first: Word's
 * reading with two, three and four copies, D6, unexplained). In `wps:txbx` the runtime's last is Word's, so it has no rule.
 */
export const JOINED: Readonly<Record<string, Readonly<Record<string, 'head' | 'tail'>>>> = {
  [`${WML}Document`]: { body: 'head' },
  [`${WML}SdtBlock`]: { sdtContent: 'head' },
  [`${WML}SdtRun`]: { sdtContent: 'head' },
  [`${WML}CTSdtRow`]: { sdtContent: 'head' },
  [`${WML}CTSdtCell`]: { sdtContent: 'head' },
  'org_docx4j_vml.CTTextbox': { txbxContent: 'tail' },
};

/**
 * The properties merged where they repeat (holder class or a class based on it, property): `w:pPr` in a paragraph, `w:rPr` in a
 * run, `w:sectPr` in a body or a `w:pPr`, `w:tblPr` in a table, and a child given twice in one parent as the rounds `sectpr-merge`
 * and `one-parent-word-check` measured it: `w:ind`, `w:spacing`, `w:numPr` and a paragraph mark's `w:rPr` in one `w:pPr`,
 * `w:rFonts` in one run's `w:rPr`, `w:sectPr` in one `w:pPr`, and `w:pgSz`, `w:pgMar` and `w:cols` in one `w:sectPr` (`sm-53`,
 * `sm-70`, `sm-72`, `sm-73`, `op-10` to `op-21`). `w:rFonts` twice in a paragraph mark's `w:rPr` is unmeasured, and keeps the last.
 * `w:jc`, `w:sz`, `w:b` and `w:vanish` need nothing: the runtime's last is Word's.
 */
export const MERGED_IN: Readonly<Record<string, readonly string[]>> = {
  [`${WML}P`]: ['pPr'],
  [`${WML}R`]: ['rPr'],
  [`${WML}PPr`]: ['rPr', 'sectPr'],
  [`${WML}PPrBase`]: ['ind', 'spacing', 'numPr'],
  [`${WML}RPr`]: ['rFonts'],
  [`${WML}Body`]: ['sectPr'],
  [`${WML}SectPr`]: ['pgSz', 'pgMar', 'cols'],
  [`${WML}Tbl`]: ['tblPr'],
};

/** The classes whose copies merge, and whose children of these classes merge in turn, one level in and on (REQ-322). */
export const MERGED: ReadonlySet<string> = new Set(['PPr', 'RPr', 'ParaRPr', 'SectPr', 'TblPr', 'PPrBase.Ind', 'PPrBase.Spacing',
  'PPrBase.NumPr', 'RFonts', 'SectPr.PgSz', 'SectPr.PgMar', 'CTColumns'].map((name) => WML + name));

/**
 * Taken as the runtime leaves it, the later copy's, or none if it has none: REQ-322 says nothing of `w:tblW`, and Word's save of
 * `sm-61` fits neither reading (docx4j's `AS_JAXB_LEAVES`).
 */
export const AS_THE_RUNTIME_LEAVES: ReadonlySet<string> = new Set(['tblW']);

/**
 * Where a header's or a footer's story reads a property otherwise (D7 (c), Jason 2026-11-10): a paragraph's repeated `w:pPr` keeps
 * the first copy whole, every later copy ignored, an empty first copy too. Word, in the rounds `header-repeats` and
 * `header-footer-repeats` (11 of 11 cases in a header and a footer; the main-part twins merge), and in a header's text box
 * (`notes-comments-repeats`, 2 of 2). A later copy's mark `w:rPr`, which Word puts on the paragraph's run (`hf-16`), isn't followed.
 */
export const FIRST_IN_HEADER_FOOTER: Readonly<Record<string, readonly string[]>> = {
  [`${WML}P`]: ['pPr'],
};

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** A header's and a footer's story, by its element. */
export const HEADER_FOOTER_STORIES: ReadonlySet<string> = new Set(['hdr', 'ftr']);

/**
 * The other stories a paragraph can stand in, by element, which end the search for a header or a footer; none stands in one, so they
 * only shorten it. Notes and comments read as the main part (Word, the portfolio's round `notes-comments-repeats`, 7 shapes of 7 in
 * each). A text box isn't among them: its paragraph reads as its host story's, so a header's text box keeps the first `w:pPr` (the
 * same round, `nr-tb-20` and `nr-tb-21`).
 */
export const OTHER_STORIES: ReadonlySet<string> = new Set(['body', 'footnote', 'endnote', 'comment', 'docPartBody']);

/** Whether the element's nearest story is a header's or a footer's: through tables, controls, mc branches and text boxes. */
function inHeaderOrFooter(element: Node | null): boolean {
  for (let e = element; e !== null && e.nodeType === 1; e = e.parentNode) {
    const { namespaceURI, localName } = e as Element;
    if (namespaceURI !== W) continue;
    if (HEADER_FOOTER_STORIES.has(localName)) return true;
    if (OTHER_STORIES.has(localName)) return false;
  }
  return false;
}

/** The later copy given the earlier's properties whole, and nothing of its own, re-linked. */
function keepFirst(earlier: Typed, later: Typed): void {
  if (earlier.TYPE_NAME !== later.TYPE_NAME) return;
  for (const property of Object.keys(later)) if (property !== 'TYPE_NAME') delete later[property];
  for (const property of Object.keys(earlier)) {
    if (property === 'TYPE_NAME') continue;
    const theirs = earlier[property];
    if (Array.isArray(theirs)) for (const item of theirs) relink(item, later); else relink(theirs, later);
    later[property] = theirs;
  }
}

/** The holder's class, or a class it is based on (an `xsi:type` element is read as its derived type), named `name`. */
function isA(classInfo: Jsonix.TypeInfo | undefined, name: string): boolean {
  for (let c = classInfo; c; c = c.baseTypeInfo) if (c.name === name) return true;
  return false;
}

function rule(classInfo: Jsonix.ClassInfo, property: string): 'head' | 'tail' | 'merge' | undefined {
  for (let c: Jsonix.TypeInfo | undefined = classInfo; c; c = c.baseTypeInfo) {
    const join = JOINED[c.name]?.[property];
    if (join) return join;
    if (MERGED_IN[c.name]?.includes(property)) return 'merge';
  }
  return undefined;
}

function firstInHeaderFooter(classInfo: Jsonix.ClassInfo, property: string): boolean {
  for (let c: Jsonix.TypeInfo | undefined = classInfo; c; c = c.baseTypeInfo) if (FIRST_IN_HEADER_FOOTER[c.name]?.includes(property)) return true;
  return false;
}

function isTyped(value: unknown): value is Typed {
  return value !== null && typeof value === 'object' && typeof (value as Typed).TYPE_NAME === 'string';
}

/** `PARENT` set on a moved value as the unmarshaller sets it: on a typed object, or on a `{ name, value }` wrapper's value. */
function relink(item: unknown, owner: object): void {
  if (isTyped(item)) { Jsonix.Util.setParent(item, owner); return; }
  const wrapped = item !== null && typeof item === 'object' && 'name' in item ? (item as { value?: unknown }).value : undefined;
  if (isTyped(wrapped)) Jsonix.Util.setParent(wrapped, owner);
}

/** The earlier container's content moved into the later's, at its head or its tail, and re-linked. */
function moveContent(earlier: Typed, later: Typed, where: 'head' | 'tail'): void {
  const from = earlier.content;
  if (!Array.isArray(from) || from.length === 0) return;
  const moved = from.splice(0);
  for (const item of moved) relink(item, later);
  const to: unknown[] = Array.isArray(later.content) ? later.content : [];
  later.content = to;
  if (where === 'head') to.unshift(...moved); else to.push(...moved);
}

/**
 * Two bodies joined in order, the earlier's content at the head of the later's, and their closing `w:sectPr` merged, a body with
 * none passed over (WD REQ-322).
 */
function joinBodies(earlier: Typed, later: Typed): void {
  moveContent(earlier, later, 'head');
  const first = earlier.sectPr, last = later.sectPr;
  if (!isTyped(first)) return;
  delete earlier.sectPr;
  if (!isTyped(last)) { relink(first, later); later.sectPr = first; }
  else merge(first, last);
}

/**
 * The later copy applied to the earlier, the result in the later (WD REQ-322): what the later states wins and what it doesn't state
 * stays, attributes and child elements alike, and a child of a class in `MERGED` merges in turn. An on/off element with no `w:val`
 * states on, and so wins. A list is the later's, as docx4j's merge leaves it, but a `w:sectPr`'s header and footer references: only
 * the earlier's stand, any on the later being ignored, an earlier with none leaving none. `w:tblW` is the later's.
 */
function merge(earlier: Typed, later: Typed): void {
  if (earlier.TYPE_NAME !== later.TYPE_NAME) return;
  for (const property of Object.keys(earlier)) {
    if (property === 'TYPE_NAME' || AS_THE_RUNTIME_LEAVES.has(property)) continue;
    const theirs = earlier[property];
    if (theirs === undefined || theirs === null || Array.isArray(theirs)) continue;
    const mine = later[property];
    if (mine === undefined || mine === null) {
      relink(theirs, later);
      later[property] = theirs;
    } else if (isTyped(mine) && MERGED.has(mine.TYPE_NAME!) && isTyped(theirs)) {
      merge(theirs, mine);
    }
  }
  if (later.TYPE_NAME === `${WML}SectPr`) {
    const references = earlier.egHdrFtrReferences;
    if (Array.isArray(references) && references.length > 0) {
      for (const reference of references) relink(reference, later);
      later.egHdrFtrReferences = references;
    } else delete later.egHdrFtrReferences;
  }
}

/**
 * The facade's `onOverwrittenElement` (jsonix-CR-016), set on every facade unmarshal unless the caller passes one (CR-022 D3): a child
 * given more than once where the model holds one is read as Word 365 reads it (WD REQ-322), editing `keptValue`, the copy the runtime
 * keeps, in place, while `value`, the earlier, is still held. With three copies it is called twice, so the kept copy gathers the
 * earlier ones in turn, and a save writes one.
 * - **Joined:** `w:body` in a `w:document`, and `w:sdtContent` in a control at every level, in order, the earlier's content at the head
 *   of the later's; joined bodies' closing `w:sectPr` merged, a body with none passed over. `w:txbxContent` in a VML box in reverse,
 *   the later's content first.
 * - **Merged:** `w:pPr` in a paragraph, `w:rPr` in a run, `w:sectPr` in a body or a `w:pPr`, `w:tblPr` in a table, and `w:ind`,
 *   `w:spacing`, `w:numPr`, a paragraph mark's `w:rPr`, `w:rFonts` in a run's `w:rPr`, and `w:pgSz`, `w:pgMar` and `w:cols`, given
 *   twice in one parent: the later applied to the earlier, what it states winning and what it doesn't staying, the classes REQ-322
 *   names merging in turn. Only the first copy's `w:headerReference` and `w:footerReference` stand. `w:tblW` is the later's.
 * - **In a header or a footer,** a paragraph's repeated `w:pPr` keeps the first copy whole instead, as Word reads it there (D7 (c)):
 *   the paragraph's nearest story is a `w:hdr` or a `w:ftr`, found from `keptNode`'s ancestors through any text box, so a paragraph
 *   in a header's text box keeps the first copy too, and one in the main part's, a note's or a comment's merges.
 * - **Anything else** is left as the runtime leaves it, the last copy, as nothing measured says otherwise.
 *
 * By class, so in any part and in a `wml()` fragment too (CR-022 D1 and D7), where docx4j reads the main document part's alone. Returns
 * whether a reading was applied, for a caller that passes its own `onOverwrittenElement` and calls this from it (to count, say);
 * the runtime ignores it.
 */
export function wordOverwritten(node: Element, keptNode: Element, propertyInfo: Jsonix.PropertyInfo, classInfo: Jsonix.ClassInfo,
  value: unknown, keptValue: unknown): boolean {
  if (!isTyped(value) || !isTyped(keptValue) || value === keptValue) return false;
  const how = rule(classInfo, propertyInfo.name);
  if (how === undefined) return false;
  if (how === 'merge') {
    if (firstInHeaderFooter(classInfo, propertyInfo.name) && inHeaderOrFooter(keptNode)) keepFirst(value, keptValue);
    else merge(value, keptValue);
  } else if (isA(classInfo, `${WML}Document`)) joinBodies(value, keptValue);
  else moveContent(value, keptValue, how);
  return true;
}
