// builders/wml (CR-002): hand-written helpers next to the generated factories, the counterpart of
// what docx4j-core's XmlUtils, TextUtils and TraversalUtil do with a tree once it exists:
//   - `wml`: an XML fragment (sibling elements as written inside document.xml) to typed elements,
//     with docx4j's namespace declarations added, wrapped in the container docx4j would put it in;
//   - text sugar over `el` (`p`, `r`, `t`, `br`, `tab`, `tbl`) and `textOf` the other way, with the
//     one name-to-w:rPr mapping (`RunOptions`, `applyRunOptions` / `readRunOptions`) that core-ts's
//     Font view shares;
//   - traversal: `walk`, `find`, `linkParents`.
// `helpers/wml` (the per-type docx4j decisions) is imported here; it never imports this module.
// Compiled to dist/builders/wml.mjs by `npm run build`.
import { Jsonix, unmarshalNode, marshalString, NAMESPACE_PREFIXES } from '../index.mjs';
import type { TypedNamedValue } from '../index.mjs';
import type * as M from '../../modules/org_docx4j_wml.mjs';
import * as el from '../../modules/org_docx4j_wml.el.mjs';
import { highlightHexValue, highlightNameForColor } from '../helpers/wml.mjs';

/** A `{ name, value }` pair as content arrays hold them. */
export type Element<T = unknown> = TypedNamedValue<T>;

export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** True for a `{ name, value }` element pair. */
export function isElement(value: unknown): value is Element {
  return typeof value === 'object' && value !== null && 'value' in value
    && typeof (value as Element).name === 'object' && (value as Element).name !== null && 'localPart' in (value as Element).name;
}

/** The TYPE_NAME of an element's value, or of a value. */
export function typeNameOf(value: unknown): string | undefined {
  const v = isElement(value) ? value.value : value;
  return typeof v === 'object' && v !== null ? (v as { TYPE_NAME?: string }).TYPE_NAME : undefined;
}

// ---------------------------------------------------------------------------------------------
// 1. wml: fragments
// ---------------------------------------------------------------------------------------------

/** The global container a fragment is wrapped in before unmarshalling (the runtime resolves global elements only). */
export type Wrapper = 'body' | 'p' | 'tr' | 'tc' | 'r' | 'none';

export interface WmlOptions {
  /** Force a wrapper instead of inferring one from the first element. */
  wrapper?: Wrapper;
  /** Runs over the parsed DOM before unmarshalling (core-ts passes its MCE preprocessor). */
  preprocess?: (doc: Document) => void;
}

/** A string marked as XML for interpolation; branded so that a plain string can never be mistaken for it. */
export interface Raw { readonly __wmlRaw: true; readonly xml: string }

/** What the tagged form interpolates: text (escaped), a typed element (marshalled), raw XML, or a list of those. */
export type Interpolated = string | number | boolean | null | undefined | Element | Raw | Interpolated[];

let declarations: string | undefined;
function standardDeclarations(): string {
  return (declarations ??= Object.entries(NAMESPACE_PREFIXES)
    .filter(([, prefix]) => prefix !== '' && prefix !== 'xml')
    .map(([ns, prefix]) => `xmlns:${prefix}="${ns}"`)
    .join(' '));
}

const RUN_LEVEL = new Set(['r', 'hyperlink', 'fldSimple', 'smartTag', 'dir', 'bdo']);
const RUN_CONTENT = new Set(['t', 'br', 'tab', 'cr', 'sym', 'drawing', 'pict', 'fldChar', 'instrText', 'delText', 'noBreakHyphen', 'softHyphen', 'object', 'ptab', 'lastRenderedPageBreak', 'footnoteReference', 'endnoteReference', 'commentReference']);
const BLOCK_ONLY = new Set(['p', 'tbl', 'altChunk']);
/** Declared at both block and run level with different types (SdtBlock / SdtRun, CTCustomXmlBlock / CTCustomXmlRun): the first decisive descendant decides. */
const DIFFERENT_AT_BOTH = new Set(['sdt', 'customXml']);
/** Global elements: unmarshalled as they are. Everything else (bookmarks, proofErr, ins, del, the range markers: the same type at both levels) goes to w:body. */
const GLOBAL = new Set(['document', 'body', 'styles', 'numbering', 'settings', 'fonts', 'hdr', 'ftr', 'footnotes', 'endnotes', 'comments', 'glossaryDocument', 'webSettings']);

const CHAINS: Record<Wrapper, string[]> = { none: [], body: ['body'], p: ['p'], tr: ['tbl', 'body'], tc: ['tr', 'tbl', 'body'], r: ['r', 'p'] };

function isTemplate(value: unknown): value is TemplateStringsArray {
  return Array.isArray(value) && 'raw' in value;
}

function escapeText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function isRaw(value: unknown): value is Raw {
  return typeof value === 'object' && value !== null && (value as Raw).__wmlRaw === true && typeof (value as Raw).xml === 'string';
}

async function interpolate(value: Interpolated): Promise<string> {
  if (value === null || value === undefined) return '';
  if (Array.isArray(value)) return (await Promise.all(value.map(interpolate))).join('');
  if (isRaw(value)) return value.xml;
  if (isElement(value)) return marshalString(value);
  return escapeText(String(value));
}

async function fromTemplate(strings: TemplateStringsArray, values: Interpolated[]): Promise<string> {
  let out = strings[0] ?? '';
  for (let i = 0; i < values.length; i++) out += (await interpolate(values[i])) + (strings[i + 1] ?? '');
  return out;
}

/**
 * docx4j XmlUtils.unmarshalString with W_NAMESPACE_DECLARATION: one or more sibling WML elements,
 * as written inside document.xml, to typed elements. The standard prefixes (`NAMESPACE_PREFIXES`)
 * are declared for the fragment; its own declarations win. The fragment is wrapped in the global
 * container docx4j would put its first element in (`w:body`, `w:p`, `w:tbl`, `w:tr`, `w:r`; for
 * `w:sdt` and `w:customXml` by the first decisive descendant, since the model accepts a run-level
 * control under `w:body` and types it `SdtBlock`), and that container's content is returned, with
 * PARENT linked below the top level (the top-level PARENT is the throw-away wrapper).
 *
 * Tagged form: `wml\`<w:p>${run}</w:p>\``. A typed element interpolates as its XML, a string or
 * number is escaped as text, `wml.raw(xml)` is inserted verbatim, an array item by item.
 */
export function wml(xml: string, options?: WmlOptions): Promise<Element[]>;
export function wml(strings: TemplateStringsArray, ...values: Interpolated[]): Promise<Element[]>;
export async function wml(first: string | TemplateStringsArray, ...rest: unknown[]): Promise<Element[]> {
  if (isTemplate(first)) return parseFragment(await fromTemplate(first, rest as Interpolated[]), {});
  return parseFragment(first, (rest[0] as WmlOptions | undefined) ?? {});
}
wml.raw = (xml: string): Raw => ({ __wmlRaw: true, xml });

/** Exactly one element from a fragment; throws otherwise. */
export function wmlOne<T = unknown>(xml: string, options?: WmlOptions): Promise<Element<T>>;
export function wmlOne<T = unknown>(strings: TemplateStringsArray, ...values: Interpolated[]): Promise<Element<T>>;
export async function wmlOne<T = unknown>(first: string | TemplateStringsArray, ...rest: unknown[]): Promise<Element<T>> {
  const all = isTemplate(first) ? await wml(first, ...(rest as Interpolated[])) : await wml(first, rest[0] as WmlOptions | undefined);
  if (all.length !== 1) throw new Error(`Expected one element, found ${all.length}`);
  return all[0] as Element<T>;
}

function parseXml(text: string): Document {
  return Jsonix.DOM.parse(text);
}

function firstElementChild(node: Node, after?: globalThis.Element): globalThis.Element | undefined {
  let c = after ? after.nextSibling : node.firstChild;
  for (; c; c = c.nextSibling) if (c.nodeType === 1) return c as globalThis.Element;
  return undefined;
}

async function parseFragment(xml: string, options: WmlOptions): Promise<Element[]> {
  if (!/<[A-Za-z_]/.test(xml)) throw new Error('Not an XML fragment');
  // Parsed once inside a neutral element carrying the declarations; the wrapper chain is then built around the fragment in the DOM.
  const doc = parseXml(`<w:fragment ${standardDeclarations()}>${xml}</w:fragment>`);
  const neutral = doc.documentElement;
  const firstEl = firstElementChild(neutral);
  if (!firstEl) throw new Error('Not an XML fragment');
  const wrapper = options.wrapper ?? inferWrapper(firstEl);
  const chain = CHAINS[wrapper];
  let outer: globalThis.Element;
  if (chain.length === 0) {
    if (firstElementChild(neutral, firstEl) !== undefined) throw new Error('A global element must be the only element of the fragment');
    doc.removeChild(neutral);
    for (const attr of Array.from(neutral.attributes)) if (attr.name.startsWith('xmlns:') && !firstEl.hasAttribute(attr.name)) firstEl.setAttribute(attr.name, attr.value);
    doc.appendChild(firstEl);
    outer = firstEl;
  } else {
    let inner: globalThis.Element | undefined;
    outer = neutral;
    for (const localName of chain) {
      const e = doc.createElementNS(W_NS, `w:${localName}`);
      if (inner) e.appendChild(inner); else for (const c of Array.from(neutral.childNodes)) e.appendChild(c);
      inner = e;
      outer = e;
    }
    for (const attr of Array.from(neutral.attributes)) outer.setAttribute(attr.name, attr.value);
    doc.replaceChild(outer, neutral);
  }
  options.preprocess?.(doc);
  let element: Element = await unmarshalNode(doc);
  if (chain.length === 0) return [element];
  for (let i = chain.length - 1; i >= 0; i--) {
    const content = (element.value as { content?: Element[] }).content ?? [];
    if (i === 0) return content;
    element = content[0]!;
  }
  return [];
}

function inferWrapper(e: globalThis.Element): Wrapper {
  const local = e.localName;
  if (GLOBAL.has(local)) return 'none';
  if (local === 'tr') return 'tr';
  if (local === 'tc') return 'tc';
  if (RUN_LEVEL.has(local)) return 'p';
  if (RUN_CONTENT.has(local)) return 'r';
  if (DIFFERENT_AT_BOTH.has(local)) return decisiveDescendant(e) ?? 'body';
  return 'body';
}

/**
 * The wrapper a w:sdt or w:customXml needs, from its first decisive descendant in document order:
 * block content decides body, run content p, a row tr, a cell tc; a nested control is recursed
 * into; an element with the same type at both levels is skipped; nothing decisive (properties
 * only, an empty section, a chain of empty controls) is undefined.
 */
function decisiveDescendant(control: globalThis.Element): Wrapper | undefined {
  let holder: globalThis.Element | undefined = control;
  if (control.localName === 'sdt') {
    holder = undefined;
    for (let c = firstElementChild(control); c; c = firstElementChild(control, c)) if (c.localName === 'sdtContent') { holder = c; break; }
  }
  if (!holder) return undefined;
  for (let c = firstElementChild(holder); c; c = firstElementChild(holder, c)) {
    const local = c.localName;
    if (BLOCK_ONLY.has(local)) return 'body';
    if (RUN_LEVEL.has(local)) return 'p';
    if (local === 'tr') return 'tr';
    if (local === 'tc') return 'tc';
    if (DIFFERENT_AT_BOTH.has(local)) {
      const nested = decisiveDescendant(c);
      if (nested) return nested;
    }
  }
  return undefined;
}

// ---------------------------------------------------------------------------------------------
// 2. The run mapping: RunOptions in the vocabulary of Office JS Word.Font (core-ts's Font view)
// ---------------------------------------------------------------------------------------------

/** Office JS Word.UnderlineType, the subset w:u expresses. */
export type UnderlineType = 'None' | 'Single' | 'Word' | 'Double' | 'Thick' | 'Dotted' | 'DottedHeavy' | 'DashLine' | 'DashLineHeavy' | 'DashLineLong' | 'DashLineLongHeavy' | 'DotDashLine' | 'DotDashLineHeavy' | 'TwoDotDashLine' | 'TwoDotDashLineHeavy' | 'Wave' | 'WaveHeavy' | 'WaveDouble' | 'Mixed';

export const UNDERLINE_TO_WML: Readonly<Record<Exclude<UnderlineType, 'None' | 'Mixed'>, M.UnderlineEnumeration>> = {
  Single: 'single', Word: 'words', Double: 'double', Thick: 'thick', Dotted: 'dotted', DottedHeavy: 'dottedHeavy', DashLine: 'dash', DashLineHeavy: 'dashedHeavy',
  DashLineLong: 'dashLong', DashLineLongHeavy: 'dashLongHeavy', DotDashLine: 'dotDash', DotDashLineHeavy: 'dashDotHeavy', TwoDotDashLine: 'dotDotDash', TwoDotDashLineHeavy: 'dashDotDotHeavy',
  Wave: 'wave', WaveHeavy: 'wavyHeavy', WaveDouble: 'wavyDouble',
};
const WML_TO_UNDERLINE = new Map<string, UnderlineType>(Object.entries(UNDERLINE_TO_WML).map(([k, v]) => [v, k as UnderlineType]));

/**
 * Run formatting in the vocabulary of Office JS `Word.Font`, so that what a builder sets a view
 * reads back under the same name. `undefined` leaves a property alone; `false`, `'None'`, `''`
 * or `null` removes it.
 */
export interface RunOptions {
  bold?: boolean;
  italic?: boolean;
  /** true is 'Single'; false or 'None' removes w:u. */
  underline?: boolean | UnderlineType;
  strikeThrough?: boolean;
  doubleStrikeThrough?: boolean;
  subscript?: boolean;
  superscript?: boolean;
  /** w:rFonts ascii and hAnsi; the theme font attributes are cleared; '' removes w:rFonts. */
  name?: string;
  /** Points; w:sz and w:szCs are half-points. */
  size?: number;
  /** '#RRGGBB' or 'auto'; '' removes w:color. */
  color?: string;
  /** A highlight name or one of its '#RRGGBB' values (helpers/wml); null or '' removes; anything else throws. */
  highlightColor?: string | null;
  /** w:rStyle; '' removes. */
  style?: string;
}

/** What `readRunOptions` returns: every RunOptions property, present; underline narrowed to its names. */
export interface RunFormatting {
  bold: boolean; italic: boolean; underline: UnderlineType; strikeThrough: boolean; doubleStrikeThrough: boolean; subscript: boolean; superscript: boolean;
  /** '' when not set directly. */
  name: string;
  /** Points; 0 when not set directly. */
  size: number;
  /** '#RRGGBB', 'auto' or a theme name; '' when not set directly. */
  color: string;
  /** '#RRGGBB'; null when there is none. */
  highlightColor: string | null;
  /** '' when not set. */
  style: string;
}

/** Writes the options onto run properties in place: the single mapping from names to w:rPr. */
export function applyRunOptions(rPr: M.RPr, opts: RunOptions): M.RPr {
  if (opts.bold !== undefined) { if (opts.bold) { rPr.b = {}; rPr.bCs = {}; } else { delete rPr.b; delete rPr.bCs; } }
  if (opts.italic !== undefined) { if (opts.italic) { rPr.i = {}; rPr.iCs = {}; } else { delete rPr.i; delete rPr.iCs; } }
  if (opts.underline !== undefined) {
    const u = opts.underline === true ? 'Single' : opts.underline === false ? 'None' : opts.underline;
    if (u === 'None' || u === 'Mixed') delete rPr.u; else rPr.u = { val: UNDERLINE_TO_WML[u] };
  }
  if (opts.strikeThrough !== undefined) { if (opts.strikeThrough) rPr.strike = {}; else delete rPr.strike; }
  if (opts.doubleStrikeThrough !== undefined) { if (opts.doubleStrikeThrough) rPr.dstrike = {}; else delete rPr.dstrike; }
  if (opts.subscript !== undefined) { if (opts.subscript) rPr.vertAlign = { val: 'subscript' }; else if (rPr.vertAlign?.val === 'subscript') delete rPr.vertAlign; }
  if (opts.superscript !== undefined) { if (opts.superscript) rPr.vertAlign = { val: 'superscript' }; else if (rPr.vertAlign?.val === 'superscript') delete rPr.vertAlign; }
  if (opts.name !== undefined) {
    if (opts.name === '') delete rPr.rFonts;
    else { rPr.rFonts = { ...rPr.rFonts, ascii: opts.name, hAnsi: opts.name }; delete rPr.rFonts.asciiTheme; delete rPr.rFonts.hAnsiTheme; }
  }
  if (opts.size !== undefined) { const half = Math.round(opts.size * 2); rPr.sz = { val: half }; rPr.szCs = { val: half }; }
  if (opts.color !== undefined) {
    const hex = opts.color.startsWith('#') ? opts.color.substring(1) : opts.color;
    if (hex === '') delete rPr.color;
    else rPr.color = { val: hex.toLowerCase() === 'auto' ? 'auto' : hex.toUpperCase() };
  }
  if (opts.highlightColor !== undefined) {
    if (opts.highlightColor === null || opts.highlightColor === '') delete rPr.highlight;
    else {
      const name = highlightNameForColor(opts.highlightColor);
      if (!name) throw new Error(`Not a highlight colour: ${opts.highlightColor} (use a highlight name or one of its #RRGGBB values)`);
      rPr.highlight = { val: name };
    }
  }
  if (opts.style !== undefined) { if (opts.style === '') delete rPr.rStyle; else rPr.rStyle = { val: opts.style }; }
  return rPr;
}

function on(v: M.BooleanDefaultTrue | undefined): boolean {
  return v !== undefined && v.val !== false;
}

/** Reads the direct formatting of run properties in the same vocabulary; the inverse of applyRunOptions. */
export function readRunOptions(rPr: M.RPr | undefined): RunFormatting {
  const u = rPr?.u?.val;
  const font = rPr?.rFonts;
  const color = rPr?.color?.val;
  const h = rPr?.highlight?.val;
  return {
    bold: on(rPr?.b),
    italic: on(rPr?.i),
    underline: u === undefined || u === 'none' ? 'None' : (WML_TO_UNDERLINE.get(u) ?? 'Mixed'),
    strikeThrough: on(rPr?.strike),
    doubleStrikeThrough: on(rPr?.dstrike),
    subscript: rPr?.vertAlign?.val === 'subscript',
    superscript: rPr?.vertAlign?.val === 'superscript',
    name: font?.ascii ?? font?.hAnsi ?? '',
    size: rPr?.sz?.val === undefined ? 0 : rPr.sz.val / 2,
    color: color === undefined ? '' : /^[0-9A-Fa-f]{6}$/.test(color) ? '#' + color.toUpperCase() : color,
    highlightColor: h === undefined || h === 'none' ? null : (highlightHexValue(h) ?? h),
    style: rPr?.rStyle?.val ?? '',
  };
}

// ---------------------------------------------------------------------------------------------
// 3. Text sugar over `el`
// ---------------------------------------------------------------------------------------------

/** `w:t`, with xml:space="preserve" when the text starts or ends with whitespace or holds two spaces. */
export function t(text: string): Element<M.Text> {
  const value: M.Text = { value: text };
  if (/^\s|\s$|\s\s/.test(text)) value.space = 'preserve';
  return el.t(value);
}

/** `w:br`; a text-wrapping break has no type attribute. */
export function br(type?: M.STBrType): Element<M.Br> {
  return el.br(type && type !== 'textWrapping' ? { type } : {});
}

/** `w:tab`. */
export function tab(): Element<M.R.Tab> {
  return el.tab({});
}

function runProperties(opts: RunOptions | undefined): M.RPr | undefined {
  if (!opts) return undefined;
  const rPr = applyRunOptions({ TYPE_NAME: 'org_docx4j_wml.RPr' }, opts);
  return Object.keys(rPr).some((k) => k !== 'TYPE_NAME') ? rPr : undefined;
}

/** A run of text; a `\t` or `\n` in the text becomes `w:tab` / `w:br`. */
export function r(text: string, opts?: RunOptions): Element<M.R> {
  const content: NonNullable<M.R['content']> = [];
  for (const piece of text.split(/(\t|\n)/)) {
    if (piece === '') continue;
    if (piece === '\t') content.push(tab());
    else if (piece === '\n') content.push(br());
    else content.push(t(piece));
  }
  const value: M.R = { content };
  const rPr = runProperties(opts);
  if (rPr) value.rPr = rPr;
  return el.r(value);
}

export interface ParagraphOptions {
  /** The paragraph style (w:pStyle). */
  style?: string;
}

/**
 * A paragraph of one run (`p('Hello', { bold: true })`) or of the given runs. `style` is the
 * paragraph style; a run style on the one-run form is `runStyle`.
 */
export function p(text: string, opts?: Omit<RunOptions, 'style'> & ParagraphOptions & { runStyle?: string }): Element<M.P>;
export function p(runs: Element[], opts?: ParagraphOptions): Element<M.P>;
export function p(textOrRuns: string | Element[], opts?: Omit<RunOptions, 'style'> & ParagraphOptions & { runStyle?: string }): Element<M.P> {
  let content: Element[];
  if (typeof textOrRuns === 'string') {
    const { style: _style, runStyle, ...runOpts } = opts ?? {};
    content = [r(textOrRuns, runStyle !== undefined ? { ...runOpts, style: runStyle } : runOpts)];
  } else {
    content = textOrRuns;
  }
  const value: M.P = { content: content as M.P['content'] };
  if (opts?.style) value.pPr = { pStyle: { val: opts.style } };
  return el.p(value);
}

export interface TableOptions {
  /** The table style (w:tblStyle). */
  style?: string;
  /** Column widths in twips; equal columns over `width` when absent. */
  widths?: number[];
  /** The total width in twips when `widths` is absent; default 9026 (A4, 2.54 cm margins). */
  width?: number;
}

/** A table of text cells: one paragraph per cell, a grid from the widths or equal columns. */
export function tbl(rows: string[][], opts: TableOptions = {}): Element<M.Tbl> {
  const columns = Math.max(1, opts.widths?.length ?? 0, ...rows.map((row) => row.length));
  const total = opts.width ?? 9026;
  const equal = Math.floor(total / columns);
  // Given widths as they are; missing ones equal, the last taking the rounding remainder so the grid sums to the table width.
  const widths = Array.from({ length: columns }, (_, i) => opts.widths?.[i] ?? (i === columns - 1 ? total - equal * (columns - 1) : equal));
  const tblPr: M.TblPr = { tblW: { w: widths.reduce((a, b) => a + b, 0), type: 'dxa' } };
  if (opts.style) tblPr.tblStyle = { val: opts.style };
  const value: M.Tbl = {
    tblPr,
    tblGrid: { gridCol: widths.map((w) => ({ w })) },
    content: rows.map((row) => el.tr({
      content: Array.from({ length: columns }, (_, i) => el.tc({ tcPr: { tcW: { w: widths[i]!, type: 'dxa' } }, content: [p(row[i] ?? '')] })),
    })),
  };
  return el.tbl(value);
}

/** Run containers: their children are scanned for runs. w:del (RunDel) is not among them, so deleted text is skipped. */
const RUN_CONTAINERS = new Set([
  'org_docx4j_wml.P', 'org_docx4j_wml.P.Hyperlink', 'org_docx4j_wml.SdtRun', 'org_docx4j_wml.CTSdtContentRun', 'org_docx4j_wml.RunIns',
  'org_docx4j_wml.RunTrackChange', 'org_docx4j_wml.CTSmartTagRun', 'org_docx4j_wml.CTCustomXmlRun', 'org_docx4j_wml.P.Dir', 'org_docx4j_wml.P.Bdo',
  'org_docx4j_wml.CTSimpleField',
]);

/** The child values of a container, under whichever property the model gives them (docx4j's names: content, customXmlOrSmartTagOrSdt, footnote, ...); element pairs unwrapped. */
function childValues(v: object): object[] {
  const o = v as Record<string, unknown>;
  if ((o.TYPE_NAME === 'org_docx4j_wml.SdtRun' || o.TYPE_NAME === 'org_docx4j_wml.SdtBlock' || o.TYPE_NAME === 'org_docx4j_wml.CTSdtRow' || o.TYPE_NAME === 'org_docx4j_wml.CTSdtCell') && typeof o.sdtContent === 'object' && o.sdtContent !== null) {
    return childValues(o.sdtContent);
  }
  if (o.TYPE_NAME === 'org_docx4j_wml.Document' && typeof o.body === 'object' && o.body !== null) return childValues(o.body);
  for (const key of ['content', 'customXmlOrSmartTagOrSdt', 'footnote', 'endnote', 'comment']) {
    const list = o[key];
    if (Array.isArray(list)) {
      return list.map((item: unknown) => (isElement(item) ? item.value : item)).filter((item): item is object => typeof item === 'object' && item !== null);
    }
  }
  return [];
}

function itemText(item: Element): string | undefined {
  const v = item.value as Record<string, unknown>;
  switch (item.name.localPart) {
    case 't': return String(v.value ?? '');
    case 'tab': return '\t';
    case 'br': return '\n';
    case 'cr': return '\n';
    case 'noBreakHyphen': return '\u2011';
    case 'softHyphen': return '\u00AD';
    case 'sym': {
      const code = v._char;
      return typeof code === 'string' && /^[0-9A-Fa-f]{4}$/.test(code) ? String.fromCharCode(parseInt(code, 16)) : '';
    }
    default: return undefined;
  }
}

function runText(run: M.R): string {
  let out = '';
  for (const item of run.content ?? []) out += itemText(item as Element) ?? '';
  return out;
}

function inlineText(container: object): string {
  let out = '';
  for (const item of childValues(container)) {
    const tn = typeNameOf(item);
    if (tn === 'org_docx4j_wml.R') out += runText(item as M.R);
    else if (tn !== undefined && RUN_CONTAINERS.has(tn)) out += inlineText(item);
  }
  return out;
}

function blockTexts(container: object, out: string[]): void {
  const tn = (container as { TYPE_NAME?: string }).TYPE_NAME;
  if (tn === 'org_docx4j_wml.R') { out.push(runText(container as M.R)); return; }
  if (tn !== undefined && RUN_CONTAINERS.has(tn)) { out.push(inlineText(container)); return; }
  for (const item of childValues(container)) blockTexts(item, out);
}

/**
 * docx4j TextUtils: the text of a run, a paragraph or a block container (cell, table, body,
 * header, footer, content control, document). w:t, w:tab (\t), w:br and w:cr (\n),
 * w:noBreakHyphen, w:softHyphen and w:sym contribute; runs inside hyperlinks, content controls,
 * fields and insertions are read; deletions are skipped; paragraphs are joined with \n.
 */
export function textOf(value: Element | object): string {
  const v = isElement(value) ? value.value : value;
  if (typeof v !== 'object' || v === null) return '';
  const out: string[] = [];
  blockTexts(v, out);
  return out.join('\n');
}

// ---------------------------------------------------------------------------------------------
// 4. Traversal
// ---------------------------------------------------------------------------------------------

function isQName(value: object): boolean {
  return 'localPart' in value && 'namespaceURI' in value && !('TYPE_NAME' in value);
}

function isLeaf(value: object): boolean {
  return isQName(value) || typeof (value as Node).nodeType === 'number' || value instanceof Date
    || typeof (value as { year?: unknown }).year === 'number' && !('TYPE_NAME' in value);
}

/**
 * docx4j TraversalUtil: depth first over every typed object under `root`. `{ name, value }` pairs
 * and arrays are descended, not visited; QNames, calendars and DOM nodes are not entered. The
 * visitor gets the object, its parent object and the key (property name or array index) it was
 * reached by; return false to not descend.
 */
export function walk(root: unknown, visitor: (value: object, parent: object | undefined, key: string | number) => boolean | void, parent?: object, key: string | number = ''): void {
  if (typeof root !== 'object' || root === null) return;
  if (Array.isArray(root)) {
    for (let i = 0; i < root.length; i++) walk(root[i], visitor, parent, i);
    return;
  }
  if (isElement(root)) {
    walk(root.value, visitor, parent, key);
    return;
  }
  if (isLeaf(root)) return;
  if (visitor(root, parent, key) === false) return;
  for (const k of Object.keys(root)) {
    if (k === 'TYPE_NAME' || k === 'PARENT') continue;
    walk((root as Record<string, unknown>)[k], visitor, root, k);
  }
}

/** docx4j ClassFinder: every object under `root` whose TYPE_NAME is `typeName`. */
export function find<T = unknown>(root: unknown, typeName: string): T[] {
  const out: T[] = [];
  walk(root, (v) => {
    if ((v as { TYPE_NAME?: string }).TYPE_NAME === typeName) out.push(v as T);
  });
  return out;
}

/** Sets PARENT (non-enumerable, writable, configurable: as the unmarshaller does) on `value` and, recursively, on its typed descendants. */
export function linkParents(value: unknown, parent: object | undefined): void {
  walk(value, (v, parentOfV) => { Jsonix.Util.setParent(v, parentOfV); }, parent);
}
