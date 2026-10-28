// CR-017: misplaced children. docx4j CR-064's `ArrayListDocx4j.checkParent` sets (befc57371), a child tested by its
// class as docx4j tests the value a JAXBElement wraps, with core-ts CR-016's name-and-value rule (its D3) and
// docx4j's message wording (its D4). Hand-written and internal: the facade reports through `onMisplaced` (it passes
// the context, for the name rule in full), and `builders/wml` refuses with a TypeError (without the context: the
// name rule for `w:sdt` only, from the table below). Not a public path.
//
// Each container refuses only children its schema never allows there, so a file read can't trip a class rule
// unless it holds a misplaced child; it never says "allow only" (CR-064 D4, D8).

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MCE_ALTERNATE_CONTENT = 'org_docx4j_mce.AlternateContent';

/** The five sets of `checkParent`; `sdtContent` is the separate rule for sdt content in sdt content. */
export type MisplacedSet = 'block-level' | 'run-level' | 'run' | 'table' | 'row';

export interface Container {
  /** The set of `checkParent` the container is in; none for row- and cell-level sdt content. */
  readonly set?: MisplacedSet;
  /** Whether it is sdt content, which refuses sdt content directly. */
  readonly sdtContent: boolean;
  /** The property holding its children. */
  readonly list: string;
  /** For a message: what a reader of the code calls it, with its element name. */
  readonly describe: string;
}

const block = (describe: string, sdtContent = false): Container => ({ set: 'block-level', sdtContent, list: 'content', describe });
const runLevel = (describe: string, list = 'content', sdtContent = false): Container => ({ set: 'run-level', sdtContent, list, describe });

/** The containers, by `TYPE_NAME` (none has a subtype; the smoke checks it). */
export const CONTAINERS: Readonly<Record<string, Container>> = {
  'org_docx4j_wml.Body': block('the body (w:body)'),
  'org_docx4j_wml.Tc': block('a table cell (w:tc)'),
  'org_docx4j_wml.SdtContentBlock': block('block-level sdt content (w:sdtContent)', true),
  'org_docx4j_wml.CTTxbxContent': block('a text box (w:txbxContent)'),
  'org_docx4j_wml.Hdr': block('a header (w:hdr)'),
  'org_docx4j_wml.Ftr': block('a footer (w:ftr)'),
  'org_docx4j_wml.CTFtnEdn': block('a footnote or endnote'),
  'org_docx4j_wml.Comments.Comment': block('a comment (w:comment)'),
  'org_docx4j_wml.CTCustomXmlBlock': block('block-level custom XML (w:customXml)'),
  'org_docx4j_wml.P': runLevel('a paragraph (w:p)'),
  'org_docx4j_wml.P.Hyperlink': runLevel('a hyperlink (w:hyperlink)'),
  'org_docx4j_wml.CTSimpleField': runLevel('a field (w:fldSimple)'),
  'org_docx4j_wml.RunIns': runLevel('an insertion (w:ins)', 'customXmlOrSmartTagOrSdt'),
  'org_docx4j_wml.RunDel': runLevel('a deletion (w:del)', 'customXmlOrSmartTagOrSdt'),
  'org_docx4j_wml.CTSmartTagRun': runLevel('a smart tag (w:smartTag)'),
  'org_docx4j_wml.CTCustomXmlRun': runLevel('run-level custom XML (w:customXml)'),
  'org_docx4j_wml.CTSdtContentRun': runLevel('run-level sdt content (w:sdtContent)', 'content', true),
  'org_docx4j_wml.R': { set: 'run', sdtContent: false, list: 'content', describe: 'a run (w:r)' },
  'org_docx4j_wml.Tbl': { set: 'table', sdtContent: false, list: 'content', describe: 'a table (w:tbl)' },
  'org_docx4j_wml.Tr': { set: 'row', sdtContent: false, list: 'content', describe: 'a table row (w:tr)' },
  'org_docx4j_wml.CTSdtContentRow': { sdtContent: true, list: 'content', describe: 'row-level sdt content (w:sdtContent)' },
  'org_docx4j_wml.CTSdtContentCell': { sdtContent: true, list: 'content', describe: 'cell-level sdt content (w:sdtContent)' },
};

const T = (name: string) => `org_docx4j_wml.${name}`;

/**
 * The classes refused, by set, as `checkParent` has them (CR-017 section 3), with docx4j CR-064's addendum (its section
 * 8.1, ce87ec6d7): deleted text where text is refused, a tab where a break is, and a content control at the wrong level.
 * A string is refused by every set.
 */
export const REFUSED: Readonly<Record<MisplacedSet, ReadonlySet<string>>> = {
  'block-level': new Set(['R', 'Text', 'DelText', 'Drawing', 'Br', 'R.Tab', 'Tc', 'Tr', 'SectPr', 'TcPr', 'SdtRun', 'CTSdtRow', 'CTSdtCell'].map(T)),
  'run-level': new Set(['Tbl', 'P', 'Text', 'DelText', 'Drawing', 'Br', 'R.Tab', 'PPr', 'SdtBlock', 'CTSdtRow', 'CTSdtCell'].map(T)),
  run: new Set(['R', 'P', 'Tbl', 'RPr', 'SdtBlock', 'SdtRun', 'CTSdtRow', 'CTSdtCell'].map(T)),
  table: new Set(['Tc', 'P', 'R', 'Tbl', 'TblPr', 'SdtBlock', 'SdtRun', 'CTSdtCell'].map(T)),
  row: new Set(['P', 'R', 'Tbl', 'TrPr', 'SdtBlock', 'SdtRun', 'CTSdtRow'].map(T)),
};

const SDT_CONTROLS: ReadonlySet<string> = new Set(['SdtBlock', 'SdtRun', 'CTSdtRow', 'CTSdtCell'].map(T));
/** Where a content control of each level goes, for a message (docx4j's `sdtGoes`). */
const SDT_GOES: Readonly<Record<string, string>> = {
  [T('SdtBlock')]: 'beside a paragraph (in the body, a cell, a header or the like)', [T('SdtRun')]: 'in a paragraph (w:p)',
  [T('CTSdtRow')]: 'in a table (w:tbl)', [T('CTSdtCell')]: 'in a table row (w:tr)',
};

/** Sdt content's four classes, refused directly in sdt content. */
export const SDT_CONTENT_TYPES: ReadonlySet<string> = new Set(['SdtContentBlock', 'CTSdtContentRun', 'CTSdtContentRow', 'CTSdtContentCell'].map(T));

/**
 * The element names of each refused class (the smoke holds them to the context): a hand-built value without a
 * `TYPE_NAME`, or a DOM element, is tested by the class its name gives.
 */
export const CLASS_NAMES: Readonly<Record<string, readonly string[]>> = {
  [T('R')]: ['r'], [T('P')]: ['p'], [T('Tbl')]: ['tbl'], [T('Tc')]: ['tc'], [T('Tr')]: ['tr'],
  [T('Text')]: ['t', 'instrText', 'delInstrText', 'fldData'], [T('DelText')]: ['delText'], [T('Drawing')]: ['drawing'], [T('Br')]: ['br'], [T('R.Tab')]: ['tab'],
  [T('PPr')]: ['pPr'], [T('RPr')]: ['rPr'], [T('SectPr')]: ['sectPr'], [T('TblPr')]: ['tblPr'], [T('TrPr')]: ['trPr'], [T('TcPr')]: ['tcPr'],
};
const CLASS_BY_NAME: ReadonlyMap<string, string> = new Map(Object.entries(CLASS_NAMES).flatMap(([type, names]) => names.map((n) => [n, type] as const)));

/** The class each set's list declares `w:sdt` as (the smoke holds it to the context): the name rule the builders apply. */
export const SDT_DECLARED: Readonly<Record<string, string>> = {
  'block-level': T('SdtBlock'), 'run-level': T('SdtRun'), table: T('CTSdtRow'), row: T('CTSdtCell'),
  [T('CTSdtContentRow')]: T('CTSdtRow'), [T('CTSdtContentCell')]: T('CTSdtCell'),
};

export type MisplacedRule = 'misplaced' | 'name-value';

export interface Misplacement {
  readonly rule: MisplacedRule;
  readonly message: string;
  /** The child's element name, or '#text' for a string. */
  readonly child: { readonly namespaceURI: string; readonly localPart: string; readonly prefix?: string } | '#text';
  /** The child's class, where it has one. */
  readonly childType?: string;
}

/** What the facade knows from the context: the class a list, or the context globally, declares for a name, and subtyping. */
export interface Declarations {
  declared(containerType: string, list: string, name: { namespaceURI: string; localPart: string }): string | undefined;
  isA(type: string, base: string): boolean;
}

interface Child {
  readonly name?: { namespaceURI: string; localPart: string; prefix?: string };
  readonly type?: string;
  /** The value's own TYPE_NAME, as against one its name gives. */
  readonly ownType: boolean;
  readonly text: boolean;
}

function isDomElement(item: unknown): item is { nodeType: number; namespaceURI: string | null; localName: string | null; prefix: string | null } {
  return typeof item === 'object' && item !== null && (item as { nodeType?: unknown }).nodeType === 1;
}

function childOf(item: unknown): Child {
  if (typeof item !== 'object' || item === null) return { text: true, ownType: false };
  if (isDomElement(item)) {
    const name = { namespaceURI: item.namespaceURI ?? '', localPart: item.localName ?? '', ...(item.prefix ? { prefix: item.prefix } : {}) };
    return { name, type: name.namespaceURI === W ? CLASS_BY_NAME.get(name.localPart) : undefined, ownType: false, text: false };
  }
  const pair = item as { name?: { namespaceURI?: string; localPart?: string; prefix?: string }; value?: unknown };
  if (typeof pair.name === 'object' && pair.name !== null && typeof pair.name.localPart === 'string') {
    const name = { namespaceURI: pair.name.namespaceURI ?? '', localPart: pair.name.localPart, ...(pair.name.prefix ? { prefix: pair.name.prefix } : {}) };
    const own = typeof pair.value === 'object' && pair.value !== null ? (pair.value as { TYPE_NAME?: unknown }).TYPE_NAME : undefined;
    if (typeof own === 'string') return { name, type: own, ownType: true, text: false };
    return { name, type: name.namespaceURI === W ? CLASS_BY_NAME.get(name.localPart) : undefined, ownType: false, text: false };
  }
  const own = (item as { TYPE_NAME?: unknown }).TYPE_NAME;
  return { type: typeof own === 'string' ? own : undefined, ownType: typeof own === 'string', text: false };
}

/** A name for a message: its prefix where it has one, `w:` for WordprocessingML, else the local name. */
function display(name: { namespaceURI: string; localPart: string; prefix?: string }): string {
  if (name.prefix) return `${name.prefix}:${name.localPart}`;
  return name.namespaceURI === W ? `w:${name.localPart}` : name.localPart;
}

const CHILD_DESCRIPTIONS: Readonly<Record<string, string>> = {
  [T('R')]: 'a run (w:r)', [T('P')]: 'a paragraph (w:p)', [T('Tbl')]: 'a table (w:tbl)', [T('Tc')]: 'a table cell (w:tc)', [T('Tr')]: 'a table row (w:tr)',
  [T('Drawing')]: 'a drawing (w:drawing)', [T('Br')]: 'a break (w:br)', [T('DelText')]: 'deleted text (w:delText)', [T('R.Tab')]: 'a tab (w:tab)',
  [T('PPr')]: 'paragraph properties (w:pPr)', [T('RPr')]: 'run properties (w:rPr)', [T('SectPr')]: 'section properties (w:sectPr)',
  [T('TblPr')]: 'table properties (w:tblPr)', [T('TrPr')]: 'row properties (w:trPr)', [T('TcPr')]: 'cell properties (w:tcPr)',
  [T('SdtBlock')]: 'a block-level content control (SdtBlock)', [T('SdtRun')]: 'a run-level content control (SdtRun)',
  [T('CTSdtRow')]: 'a row-level content control (CTSdtRow)', [T('CTSdtCell')]: 'a cell-level content control (CTSdtCell)',
};

function describeChild(type: string, name?: { namespaceURI: string; localPart: string; prefix?: string }): string {
  if (type === T('Text')) return `text (${name ? display(name) : 'w:t'})`;
  return CHILD_DESCRIPTIONS[type] ?? type.replace(/^org_docx4j_[a-z0-9_]*\./, '');
}

/** The message for a child its container's set refuses: docx4j's wording (CR-064 section 8), with element names. */
function refusalMessage(container: Container, containerType: string, child: Child): string {
  const c = container.describe;
  if (child.text) {
    switch (container.set) {
      case 'block-level': return `You can't add a string to ${c}; put it in text (w:t), in a run (w:r), in a paragraph (w:p), and add the paragraph.`;
      case 'run-level': return `You can't add a string to ${c}; put it in text (w:t), in a run (w:r), and add the run.`;
      case 'run': return `You can't add a string to a run (w:r); put it in text (w:t) and add that.`;
      case 'table': return `You can't add a string to a table (w:tbl); text goes in a run (w:r) in a paragraph (w:p) in a cell (w:tc).`;
      default: return `You can't add a string to a table row (w:tr); text goes in a run (w:r) in a paragraph (w:p) in a cell (w:tc).`;
    }
  }
  const type = child.type!;
  const what = describeChild(type, child.name);
  if (SDT_CONTROLS.has(type)) {
    const goes = `it goes ${SDT_GOES[type]}`;
    switch (container.set) {
      case 'run-level': return `You can't add ${what} to ${c}; ${goes}; a run-level content control (SdtRun) goes here.`;
      case 'block-level': return `You can't add ${what} to ${c}; ${goes}; a block-level content control (SdtBlock) goes here.`;
      case 'run': return `You can't add a content control (w:sdt) inside a run (w:r); a run-level content control (SdtRun) goes in the paragraph, beside the run.`;
      case 'table': return `You can't add ${what} to a table (w:tbl); ${goes}; a row-level content control (CTSdtRow) goes here.`;
      default: return `You can't add ${what} to a table row (w:tr); ${goes}; a cell-level content control (CTSdtCell) goes here.`;
    }
  }
  switch (container.set) {
    case 'run-level':
      if (type === T('Tbl')) return `You can't add a table (w:tbl) inside ${c}; it should be a sibling of the paragraph.`;
      if (type === T('P')) return `You can't nest a paragraph (w:p) inside ${c}.`;
      if (type === T('PPr')) return `You can't add paragraph properties (w:pPr) to the content of ${c}; set them as the paragraph's pPr.`;
      return `You can't add ${what} directly to ${c}; add it to a run (w:r) and add that.`;
    case 'run':
      if (type === T('R')) return `You can't nest a run (w:r) inside a run; add it to the paragraph.`;
      if (type === T('P')) return `You can't nest a paragraph (w:p) inside a run (w:r).`;
      if (type === T('Tbl')) return `You can't add a table (w:tbl) inside a run (w:r).`;
      return `You can't add run properties (w:rPr) to a run's content; set them as the run's rPr.`;
    case 'block-level':
      if (type === T('R')) return `You can't add a run (w:r) directly to ${c}; add it to a paragraph (w:p) and add that.`;
      if (type === T('Tc')) return `You can't add a table cell (w:tc) to ${c}; a cell goes in a row (w:tr), in a table (w:tbl).`;
      if (type === T('Tr')) return `You can't add a table row (w:tr) to ${c}; a row goes in a table (w:tbl).`;
      if (type === T('SectPr')) return `You can't add section properties (w:sectPr) to the content of ${c}; set the body's sectPr, or pPr.sectPr on the paragraph that ends the section.`;
      if (type === T('TcPr')) return containerType === T('Tc')
        ? `You can't add cell properties (w:tcPr) to a table cell's content; set them as the cell's tcPr.`
        : `You can't add cell properties (w:tcPr) to ${c}; cell properties belong to a table cell (w:tc), as its tcPr.`;
      return `You can't add ${what} directly to ${c}; add it to a run (w:r) in a paragraph (w:p) and add the paragraph.`;
    case 'table':
      if (type === T('Tc')) return `You can't add a table cell (w:tc) directly to a table (w:tbl); add it to a row (w:tr) and add that.`;
      if (type === T('Tbl')) return `You can't nest a table (w:tbl) directly in a table; a nested table goes in a cell (w:tc).`;
      if (type === T('TblPr')) return `You can't add table properties (w:tblPr) to a table's content; set them as the table's tblPr.`;
      return `You can't add ${what} directly to a table (w:tbl); add it to a cell (w:tc) in a row (w:tr).`;
    default:
      if (type === T('Tbl')) return `You can't nest a table (w:tbl) in a table row (w:tr); a nested table goes in a cell (w:tc).`;
      if (type === T('TrPr')) return `You can't add row properties (w:trPr) to a row's content; set them as the row's trPr.`;
      return `You can't add ${what} directly to a table row (w:tr); add it to a cell (w:tc) and add that.`;
  }
}

/**
 * Whether `item`, in `containerType`'s list, is misplaced (CR-017 section 3), and why: a class its container's set
 * refuses, sdt content directly in sdt content, or a name and value that disagree. With `declarations` (the facade's,
 * from the context) the name rule is tested for every child; without (the builders), for `w:sdt` alone.
 */
export function misplacement(containerType: string, item: unknown, declarations?: Declarations): Misplacement | undefined {
  const container = CONTAINERS[containerType];
  if (container === undefined) return undefined;
  const child = childOf(item);
  const reported = (rule: MisplacedRule, message: string): Misplacement => ({
    rule, message, child: child.text || !child.name ? '#text' : child.name, ...(child.type ? { childType: child.type } : {}),
  });
  if (child.text) return container.set ? reported('misplaced', refusalMessage(container, containerType, child)) : undefined;
  if (child.type && container.set && REFUSED[container.set].has(child.type)) return reported('misplaced', refusalMessage(container, containerType, child));
  if (child.type && container.sdtContent && SDT_CONTENT_TYPES.has(child.type)) {
    return reported('misplaced', `You can't nest sdt content (w:sdtContent) directly inside ${container.describe}; nest a content control (w:sdt) that holds it.`);
  }
  if (child.name && child.ownType && child.type) {
    const declared = declarations
      ? declarations.declared(containerType, container.list, child.name)
      : child.name.namespaceURI === W && child.name.localPart === 'sdt' ? SDT_DECLARED[container.set ?? containerType] : undefined;
    if (declared !== undefined && declared !== child.type && !declarations?.isA(child.type, declared)) {
      return reported('name-value', `You can't add a ${display(child.name)} whose value is ${describeChild(child.type, child.name)} to ${container.describe}; there a ${display(child.name)} is ${describeChild(declared, child.name)}.`);
    }
  }
  return undefined;
}

/** Whether a list item is an `mc:AlternateContent`, whose branches are tested by their host's set (core-ts D3). */
export function alternateContentOf(item: unknown): { choice?: unknown[]; fallback?: unknown } | undefined {
  if (typeof item !== 'object' || item === null) return undefined;
  const value = (item as { value?: unknown }).value ?? item;
  return typeof value === 'object' && value !== null && (value as { TYPE_NAME?: unknown }).TYPE_NAME === MCE_ALTERNATE_CONTENT
    ? value as { choice?: unknown[]; fallback?: unknown } : undefined;
}

/** A branch's children (`mc:Choice` and `mc:Fallback` hold them in `any`). */
export function branchItems(branch: unknown): unknown[] {
  const any = typeof branch === 'object' && branch !== null ? (branch as { any?: unknown }).any : undefined;
  return Array.isArray(any) ? any : [];
}
