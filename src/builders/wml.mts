// builders/wml (CR-002): hand-written helpers next to the generated factories, the counterpart of
// what docx4j-core's XmlUtils, TextUtils and TraversalUtil do with a tree once it exists:
//   - `wml`: an XML fragment (sibling elements as written inside document.xml) to typed elements,
//     with docx4j's namespace declarations added, wrapped in the container docx4j would put it in;
//   - text sugar over `el` (`p`, `r`, `t`, `br`, `tab`, `tbl`) and `textOf` the other way, with the
//     one name-to-w:rPr mapping (`RunOptions`, `applyRunOptions` / `readRunOptions`) that core-ts's
//     Font view shares;
//   - content controls (CR-003): `sdt` in its four forms with a typed `w:sdtPr`, and the readers
//     `sdtProperty` / `sdtKindOf`; rows and cells (`tr`, `tc`) and `inlinePicture`, docx4j's
//     createImageInline over the generated dml factories;
//   - run properties as the element list `w:rPrChange/w:rPr` keeps (`rPrToElements` / `rPrFromElements`);
//   - traversal: `walk`, `walkAll` (DOM in `xs:any` properties too), `find`, `linkParents`.
// `helpers/wml` (the per-type docx4j decisions) is imported here; it never imports this module.
// Compiled to dist/builders/wml.mjs by `npm run build`.
import { Jsonix, unmarshalNode, marshalString, NAMESPACE_PREFIXES } from '../index.mjs';
import type { TypedNamedValue } from '../index.mjs';
import type * as M from '../../modules/org_docx4j_wml.mjs';
import type * as Dml from '../../modules/org_docx4j_dml.mjs';
import * as el from '../../modules/org_docx4j_wml.el.mjs';
import * as wmlFactory from '../../modules/org_docx4j_wml.factory.mjs';
import * as w14el from '../../modules/org_docx4j_w14.el.mjs';
import * as w15el from '../../modules/org_docx4j_w15.el.mjs';
import * as picEl from '../../modules/org_docx4j_dml_picture.el.mjs';
import { createInline, createCTEffectExtent } from '../../modules/org_docx4j_dml_wordprocessingDrawing.factory.mjs';
import { createPic, createCTPictureNonVisual } from '../../modules/org_docx4j_dml_picture.factory.mjs';
import {
  createCTPositiveSize2D, createCTNonVisualDrawingProps, createCTNonVisualGraphicFrameProperties, createCTGraphicalObjectFrameLocking,
  createGraphic, createGraphicData, createCTBlip, createCTBlipFillProperties, createCTStretchInfoProperties, createCTRelativeRect,
  createCTShapeProperties, createCTTransform2D, createCTPoint2D, createCTPresetGeometry2D, createCTGeomGuideList, createCTNonVisualPictureProperties,
} from '../../modules/org_docx4j_dml.factory.mjs';
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

export interface CellOptions {
  /** The cell width in twips (`w:tcW`, type dxa); none when absent. */
  width?: number;
}

/** A table cell: a string is one paragraph, an empty cell still gets the `w:p` Word requires. */
export function tc(blocks: string | Element[], opts: CellOptions = {}): Element<M.Tc> {
  const content = typeof blocks === 'string' ? [p(blocks)] : (blocks.length > 0 ? blocks : [el.p({})]);
  const value: M.Tc = { content: content as M.Tc['content'] };
  if (opts.width !== undefined) value.tcPr = { tcW: { w: opts.width, type: 'dxa' } };
  return el.tc(value);
}

/** A table row of cells; a string cell is one paragraph, and `widths` (twips) sets each `w:tcW`. */
export function tr(cells: (string | Element<M.Tc>)[], opts: { widths?: number[] } = {}): Element<M.Tr> {
  return el.tr({
    content: cells.map((cell, i) => {
      const width = opts.widths?.[i];
      return typeof cell === 'string' || !isElement(cell) ? tc(cell as string, width === undefined ? {} : { width }) : cell;
    }) as M.Tr['content'],
  });
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
    content: rows.map((row) => tr(Array.from({ length: columns }, (_, i) => tc(row[i] ?? '', { width: widths[i]! })))) as M.Tbl['content'],
  };
  return el.tbl(value);
}

/** Run containers: their children are scanned for runs. w:del (RunDel) is not among them, so deleted text is skipped. */
const RUN_CONTAINERS = new Set([
  'org_docx4j_wml.P', 'org_docx4j_wml.P.Hyperlink', 'org_docx4j_wml.SdtRun', 'org_docx4j_wml.CTSdtContentRun', 'org_docx4j_wml.RunIns',
  'org_docx4j_wml.RunTrackChange', 'org_docx4j_wml.CTSmartTagRun', 'org_docx4j_wml.CTCustomXmlRun', 'org_docx4j_wml.P.Dir', 'org_docx4j_wml.P.Bdo',
  'org_docx4j_wml.CTSimpleField',
]);

/** The child values of a container, under whichever property the model gives them (docx4j's names: content, customXmlOrSmartTagOrSdt, accOrBarOrBox, footnote, ...); element pairs unwrapped, w:moveFrom dropped (its runs are not the document's text, as w:del's are not). */
function childValues(v: object): object[] {
  const o = v as Record<string, unknown>;
  if ((o.TYPE_NAME === 'org_docx4j_wml.SdtRun' || o.TYPE_NAME === 'org_docx4j_wml.SdtBlock' || o.TYPE_NAME === 'org_docx4j_wml.CTSdtRow' || o.TYPE_NAME === 'org_docx4j_wml.CTSdtCell') && typeof o.sdtContent === 'object' && o.sdtContent !== null) {
    return childValues(o.sdtContent);
  }
  if (o.TYPE_NAME === 'org_docx4j_wml.Document' && typeof o.body === 'object' && o.body !== null) return childValues(o.body);
  for (const key of ['content', 'customXmlOrSmartTagOrSdt', 'accOrBarOrBox', 'footnote', 'endnote', 'comment']) {
    const list = o[key];
    if (Array.isArray(list)) {
      return list
        .filter((item: unknown) => !(isElement(item) && item.name.localPart === 'moveFrom'))
        .map((item: unknown) => (isElement(item) ? item.value : item))
        .filter((item): item is object => typeof item === 'object' && item !== null);
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

/** Office JS `Word.ContentControlType`, as Word reports a control from its `w:sdtPr` (CR-003 section 3.1). */
export type SdtKind = 'RichText' | 'PlainText' | 'Picture' | 'BuildingBlockGallery' | 'CheckBox' | 'ComboBox'
  | 'DropDownList' | 'DatePicker' | 'RepeatingSection' | 'RepeatingSectionItem' | 'Group' | 'Citation' | 'Bibliography' | 'Equation';

/** Which of the four `w:sdt` forms: the model types each differently (SdtBlock, SdtRun, CTSdtRow, CTSdtCell). */
export type SdtForm = 'block' | 'run' | 'row' | 'cell';

export interface SdtOptions {
  /** The kind element in `w:sdtPr`; rich text is the untyped control, as Word writes it. */
  kind?: SdtKind;
  /** `w:id`; a random 31-bit id when absent (`nextSdtId` gives one free in a tree). */
  id?: number;
  /** `w:tag`. */
  tag?: string;
  /** `w:alias`, which Word shows as the title. */
  title?: string;
  /** Overrides the form inferred from the content. */
  form?: SdtForm;
}

const CHECKBOX_FONT = 'MS Gothic';

/** The `w:sdtPr` child that types a control, by the kind Word reports; none for rich text. */
function kindElement(kind: SdtKind): Element | undefined {
  switch (kind) {
    case 'RichText': return undefined;
    case 'PlainText': return el.text({});
    case 'Picture': return el.picture({});
    case 'BuildingBlockGallery': return el.docPartObj({});
    case 'CheckBox': return w14el.checkbox({ checked: { val: false }, checkedState: { val: '2612', font: CHECKBOX_FONT }, uncheckedState: { val: '2610', font: CHECKBOX_FONT } }) as Element;
    case 'ComboBox': return el.comboBox({});
    case 'DropDownList': return el.dropDownList({});
    case 'DatePicker': return el.date({ dateFormat: { val: 'd/MM/yyyy' }, storeMappedDataAs: { val: 'dateTime' } });
    case 'RepeatingSection': return w15el.repeatingSection({}) as Element;
    case 'RepeatingSectionItem': return w15el.repeatingSectionItem({}) as Element;
    case 'Group': return el.group({});
    case 'Citation': return el.citation({});
    case 'Bibliography': return el.bibliography({});
    case 'Equation': return el.equation({});
  }
}

/** The kind element's local name back to the kind Word reports (docPartList is a gallery too). */
const KIND_BY_ELEMENT: Readonly<Record<string, SdtKind>> = {
  text: 'PlainText', richText: 'RichText', picture: 'Picture', docPartObj: 'BuildingBlockGallery', docPartList: 'BuildingBlockGallery',
  checkbox: 'CheckBox', comboBox: 'ComboBox', dropDownList: 'DropDownList', date: 'DatePicker',
  repeatingSection: 'RepeatingSection', repeatingSectionItem: 'RepeatingSectionItem', group: 'Group',
  citation: 'Citation', bibliography: 'Bibliography', equation: 'Equation',
};

/** Run-level items a `w:sdt` can wrap; `w:tr` and `w:tc` give the row and cell forms, everything else block. */
function formOf(content: Element[]): SdtForm {
  const first = content.find((item) => isElement(item));
  const name = first?.name.localPart;
  if (name === undefined) return 'block';
  if (name === 'tr') return 'row';
  if (name === 'tc') return 'cell';
  if (RUN_LEVEL.has(name) || RUN_CONTENT.has(name)) return 'run';
  return 'block';
}

/**
 * A `w:sdtPr`: `w:alias`, `w:tag`, `w:id` and the kind element, in that order (docx4j's
 * `org.docx4j.model.sdt` writes the same; core-ts's `insertContentControl` used to build it).
 */
export function sdtPr(options: SdtOptions = {}): M.SdtPr {
  const items: Element[] = [];
  if (options.title !== undefined) items.push(el.alias({ val: options.title }));
  if (options.tag !== undefined) items.push(el.tag({ val: options.tag }));
  items.push(el.id({ val: options.id ?? randomSdtId() }));
  const typed = options.kind === undefined ? undefined : kindElement(options.kind);
  if (typed) items.push(typed);
  return { TYPE_NAME: 'org_docx4j_wml.SdtPr', rPrOrAliasOrLock: items as M.SdtPr['rPrOrAliasOrLock'] };
}

const SDT_CONTENT_TYPE: Record<SdtForm, [sdt: string, content: string]> = {
  block: ['org_docx4j_wml.SdtBlock', 'org_docx4j_wml.SdtContentBlock'],
  run: ['org_docx4j_wml.SdtRun', 'org_docx4j_wml.CTSdtContentRun'],
  row: ['org_docx4j_wml.CTSdtRow', 'org_docx4j_wml.CTSdtContentRow'],
  cell: ['org_docx4j_wml.CTSdtCell', 'org_docx4j_wml.CTSdtContentCell'],
};

/**
 * A content control wrapping `content`. The form follows the content (`w:tr` a row, `w:tc` a cell,
 * runs and run content a run, everything else block) unless `form` says otherwise; `w:sdt` is one
 * element name with four types, so the value's `TYPE_NAME` is what tells them apart.
 * A repeating section is block-level, as in Word, and throws in the run form.
 */
export function sdt(content: Element[], options: SdtOptions = {}): Element<M.SdtBlock | M.SdtRun | M.CTSdtRow | M.CTSdtCell> {
  const form = options.form ?? formOf(content);
  if (options.kind === 'RepeatingSection' && form === 'run') {
    throw new Error('A repeating section is a block-level control; wrap paragraphs or a table, or pass form');
  }
  const [sdtType, contentType] = SDT_CONTENT_TYPE[form];
  const value = {
    TYPE_NAME: sdtType,
    sdtPr: sdtPr(options),
    sdtContent: { TYPE_NAME: contentType, content },
  } as unknown as M.SdtBlock | M.SdtRun | M.CTSdtRow | M.CTSdtCell;
  const element = el.sdt(value);
  linkParents(value, undefined);
  return element;
}

function randomSdtId(): number {
  return 1 + Math.floor(Math.random() * 0x3fffffff);
}

/** A free `w:id` for a new control: unique among the ids of the controls under `root`, as Word's are. */
export function nextSdtId(root: unknown): number {
  const used = new Set<number>();
  for (const pr of find<M.SdtPr>(root, 'org_docx4j_wml.SdtPr')) {
    for (const item of pr.rPrOrAliasOrLock ?? []) {
      if (item.name.localPart !== 'id') continue;
      const val = (item.value as { val?: number }).val;
      if (typeof val === 'number') used.add(val);
    }
  }
  let id = randomSdtId();
  while (used.has(id)) id = randomSdtId();
  return id;
}

/**
 * A `w:sdtPr` child by element name (`tag`, `alias`, `id`, `dataBinding`, w14's `checkbox`,
 * w15's `appearance`, ...): the model keeps them as one choice list, as docx4j does.
 */
export function sdtProperty<T = unknown>(pr: M.SdtPr | undefined, localPart: string, namespaceURI: string = W_NS): Element<T> | undefined {
  return (pr?.rPrOrAliasOrLock ?? []).find((item) => item.name.localPart === localPart && item.name.namespaceURI === namespaceURI) as Element<T> | undefined;
}

/** The kind Word reports for a control, from the kind element in its `w:sdtPr`; rich text when untyped. */
export function sdtKindOf(pr: M.SdtPr | undefined): SdtKind {
  for (const item of pr?.rPrOrAliasOrLock ?? []) {
    const kind = KIND_BY_ELEMENT[item.name.localPart];
    if (kind) return kind;
  }
  return 'RichText';
}

export interface InlinePictureOptions {
  /** The extent in EMU (914400 per inch): the width. */
  cx: number;
  /** The extent in EMU: the height. */
  cy: number;
  /** `wp:docPr/@id`, unique in the document. */
  id: number;
  /** `wp:docPr/@name`, what Word shows in the selection pane. */
  name: string;
  /** `wp:docPr/@descr`: Office JS `InlinePicture.altTextDescription`. */
  descr?: string;
  /** `wp:docPr/@title`: Office JS `altTextTitle` (the attribute docx4j CR-018 added). */
  title?: string;
}

/**
 * A `w:drawing` holding one `wp:inline` for an image part, as docx4j's
 * `BinaryPartAbstractImage.createImageInline` writes it: the blip's `r:embed` is `relId`, the
 * extents are in EMU, the frame locks the aspect ratio and the shape is a stretched `rect` preset.
 * The image part and its relationship belong to the package, so `relId` is a plain string here
 * (core-ts's `addImage` makes the part, then calls this).
 */
export function inlinePicture(relId: string, options: InlinePictureOptions): Element<M.Drawing> {
  const { cx, cy, id, name, descr, title } = options;
  const extent = (): Dml.CTPositiveSize2D => createCTPositiveSize2D({ cx, cy });
  const docProps = (): Dml.CTNonVisualDrawingProps => createCTNonVisualDrawingProps({
    id, name, ...(descr === undefined ? {} : { descr }), ...(title === undefined ? {} : { title }),
  });
  const picture = createPic({
    nvPicPr: createCTPictureNonVisual({ cNvPr: docProps(), cNvPicPr: createCTNonVisualPictureProperties({}) }),
    blipFill: createCTBlipFillProperties({
      blip: createCTBlip({ embed: relId }),
      stretch: createCTStretchInfoProperties({ fillRect: createCTRelativeRect({}) }),
    }),
    spPr: createCTShapeProperties({
      xfrm: createCTTransform2D({ off: createCTPoint2D({ x: 0, y: 0 }), ext: extent() }),
      prstGeom: createCTPresetGeometry2D({ prst: 'rect', avLst: createCTGeomGuideList({}) }),
    }),
  });
  const inline = createInline({
    distT: 0, distB: 0, distL: 0, distR: 0,
    extent: extent(),
    effectExtent: createCTEffectExtent({ l: 0, t: 0, r: 0, b: 0 }),
    docPr: docProps(),
    cNvGraphicFramePr: createCTNonVisualGraphicFrameProperties({ graphicFrameLocks: createCTGraphicalObjectFrameLocking({ noChangeAspect: true }) }),
    graphic: createGraphic({ graphicData: createGraphicData({ uri: PIC_NS, any: [picEl.pic(picture)] }) }),
  });
  const drawing = el.drawing({ anchorOrInline: [inline] });
  linkParents(drawing.value, undefined);
  return drawing;
}

/** The graphic data uri of a DrawingML picture. */
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

/**
 * docx4j TextUtils: the text of a run, a paragraph or a block container (cell, table, body,
 * header, footer, content control, document). w:t, w:tab (\t), w:br and w:cr (\n),
 * w:noBreakHyphen, w:softHyphen and w:sym contribute; runs inside hyperlinks, content controls,
 * fields, insertions and moved-to runs are read; deletions and w:moveFrom are skipped; paragraphs
 * are joined with \n.
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
 * The run-level items a holder keeps, under the model's own property names (CR-003 section 3.4):
 * `content` for most, `customXmlOrSmartTagOrSdt` for w:ins and w:del, `accOrBarOrBox` for
 * w:moveFrom and w:moveTo (docx4j's names for those choice groups), and `sdtContent`'s list for a
 * run-level content control. Undefined when the value holds no run list.
 */
export function runItemsOf(value: object): Element[] | undefined {
  const v = value as { content?: Element[]; customXmlOrSmartTagOrSdt?: Element[]; accOrBarOrBox?: Element[]; sdtContent?: object };
  if (Array.isArray(v.content)) return v.content;
  if (Array.isArray(v.customXmlOrSmartTagOrSdt)) return v.customXmlOrSmartTagOrSdt;
  if (Array.isArray(v.accOrBarOrBox)) return v.accOrBarOrBox;
  if (v.sdtContent) return runItemsOf(v.sdtContent);
  return undefined;
}

/**
 * `w:rPr`'s properties in `EG_RPrBase` order (the declaration order of `RPr` up to `oMath`, then
 * the w14 effects), each with the generated wrapper that writes it inside `w:rPrChange/w:rPr`:
 * the scoped `createCTRPrChangeRPr*` where XJC generated one, the module's `el` wrapper for the
 * five wml elements that have none (they are global elements), and w14's `el` for `w14:shadow`
 * (`shadow14` here, since `shadow` is wml's boolean). A property with no wrapper is not in the
 * group and is skipped, `rPrChange` among them: a recorded original has no original of its own.
 */
const RPR_ELEMENTS: ReadonlyArray<readonly [property: string, wrap: (value: never) => Element]> = [
  ['rStyle', el.rStyle], ['rFonts', el.rFonts],
  ['b', wmlFactory.createCTRPrChangeRPrB], ['bCs', wmlFactory.createCTRPrChangeRPrBCs],
  ['i', wmlFactory.createCTRPrChangeRPrI], ['iCs', wmlFactory.createCTRPrChangeRPrICs],
  ['caps', wmlFactory.createCTRPrChangeRPrCaps], ['smallCaps', wmlFactory.createCTRPrChangeRPrSmallCaps],
  ['strike', wmlFactory.createCTRPrChangeRPrStrike], ['dstrike', wmlFactory.createCTRPrChangeRPrDstrike],
  ['outline', wmlFactory.createCTRPrChangeRPrOutline], ['shadow', wmlFactory.createCTRPrChangeRPrShadow],
  ['emboss', wmlFactory.createCTRPrChangeRPrEmboss], ['imprint', wmlFactory.createCTRPrChangeRPrImprint],
  ['noProof', wmlFactory.createCTRPrChangeRPrNoProof], ['snapToGrid', wmlFactory.createCTRPrChangeRPrSnapToGrid],
  ['vanish', wmlFactory.createCTRPrChangeRPrVanish], ['webHidden', wmlFactory.createCTRPrChangeRPrWebHidden],
  ['color', el.color], ['spacing', wmlFactory.createCTRPrChangeRPrSpacing], ['w', wmlFactory.createCTRPrChangeRPrW],
  ['kern', wmlFactory.createCTRPrChangeRPrKern], ['position', wmlFactory.createCTRPrChangeRPrPosition],
  ['sz', wmlFactory.createCTRPrChangeRPrSz], ['szCs', wmlFactory.createCTRPrChangeRPrSzCs],
  ['highlight', el.highlight], ['u', el.u], ['effect', wmlFactory.createCTRPrChangeRPrEffect],
  ['bdr', wmlFactory.createCTRPrChangeRPrBdr], ['shd', wmlFactory.createCTRPrChangeRPrShd],
  ['fitText', wmlFactory.createCTRPrChangeRPrFitText], ['vertAlign', wmlFactory.createCTRPrChangeRPrVertAlign],
  ['rtl', wmlFactory.createCTRPrChangeRPrRtl], ['cs', wmlFactory.createCTRPrChangeRPrCs],
  ['em', wmlFactory.createCTRPrChangeRPrEm], ['lang', wmlFactory.createCTRPrChangeRPrLang],
  ['eastAsianLayout', wmlFactory.createCTRPrChangeRPrEastAsianLayout],
  ['specVanish', wmlFactory.createCTRPrChangeRPrSpecVanish], ['oMath', wmlFactory.createCTRPrChangeRPrOMath],
  ['glow', wmlFactory.createCTRPrChangeRPrGlow], ['shadow14', w14el.shadow],
  ['reflection', wmlFactory.createCTRPrChangeRPrReflection], ['textOutline', wmlFactory.createCTRPrChangeRPrTextOutline],
  ['textFill', wmlFactory.createCTRPrChangeRPrTextFill], ['scene3D', wmlFactory.createCTRPrChangeRPrScene3D],
  ['props3D', wmlFactory.createCTRPrChangeRPrProps3D], ['ligatures', wmlFactory.createCTRPrChangeRPrLigatures],
  ['numForm', wmlFactory.createCTRPrChangeRPrNumForm], ['numSpacing', wmlFactory.createCTRPrChangeRPrNumSpacing],
  ['stylisticSets', wmlFactory.createCTRPrChangeRPrStylisticSets], ['cntxtAlts', wmlFactory.createCTRPrChangeRPrCntxtAlts],
] as ReadonlyArray<readonly [string, (value: never) => Element]>;

/**
 * `w:rPr`'s named properties as the element list `w:rPrChange/w:rPr` keeps (`CTRPrChange.RPr`,
 * an `EG_RPrBase` list), in schema order, through the generated wrappers, the w14 effects
 * included. `w:rPrChange` itself is not carried over (a recorded original has no original).
 */
export function rPrToElements(rPr: M.RPr | undefined): NonNullable<M.CTRPrChange.RPr['egrPrBase']> {
  const out: Element[] = [];
  const source = rPr as unknown as Record<string, unknown> | undefined;
  if (!source) return out as NonNullable<M.CTRPrChange.RPr['egrPrBase']>;
  for (const [property, wrap] of RPR_ELEMENTS) {
    const value = source[property];
    if (value === undefined) continue;
    out.push((wrap as (v: unknown) => Element)(Jsonix.Util.deepCopy(value, undefined)));
  }
  return out as NonNullable<M.CTRPrChange.RPr['egrPrBase']>;
}

/** The `w:rPr` property an element of the group belongs to: the w14 names differ from the element names. */
function propertyOf(item: Element): string {
  const name = item.name.localPart;
  if (item.name.namespaceURI !== W14_NS) return name;
  if (name === 'shadow') return 'shadow14';
  if (name === 'scene3d') return 'scene3D';
  if (name === 'props3d') return 'props3D';
  return name;
}

const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';

/** The inverse: the element list back to `w:rPr`'s named properties (docx4j's reject of a formatting change). */
export function rPrFromElements(list: M.CTRPrChange.RPr | M.CTRPrChange.RPr['egrPrBase'] | undefined): M.RPr {
  const items = Array.isArray(list) ? list : list?.egrPrBase ?? [];
  const rPr = { TYPE_NAME: 'org_docx4j_wml.RPr' } as M.RPr;
  const target = rPr as unknown as Record<string, unknown>;
  for (const item of items) target[propertyOf(item)] = Jsonix.Util.deepCopy(item.value, undefined);
  linkParents(rPr, undefined);
  return rPr;
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

/**
 * `walk`, plus the DOM elements held by `xs:any` properties, which `walk` does not enter (a flat
 * OPC part, an `a:extLst` extension): `domVisitor` gets each such element and every element under
 * it, walked with `childNodes` (xmldom has no `firstElementChild`). Use it to rewrite references
 * wherever they are, such as an SVG twin's `r:embed` in an extension (CR-003 section 3.6).
 */
export function walkAll(
  root: unknown,
  visitor: (value: object, parent: object | undefined, key: string | number) => boolean | void,
  domVisitor: (node: globalThis.Element, owner: object | undefined, key: string | number) => void,
): void {
  walk(root, (value, parent, key) => {
    const result = visitor(value, parent, key);
    if (result === false) return false;
    for (const [property, held] of Object.entries(value)) {
      for (const node of (Array.isArray(held) ? held : [held])) {
        if (!isDomElement(node)) continue;
        walkDomElements(node, (element) => domVisitor(element, value, property));
      }
    }
    return undefined;
  });
}

function isDomElement(value: unknown): value is globalThis.Element {
  return typeof value === 'object' && value !== null && (value as Node).nodeType === 1;
}

function walkDomElements(node: globalThis.Element, visit: (element: globalThis.Element) => void): void {
  visit(node);
  for (let child = node.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1) walkDomElements(child as globalThis.Element, visit);
  }
}

/** Sets PARENT (non-enumerable, writable, configurable: as the unmarshaller does) on `value` and, recursively, on its typed descendants. */
export function linkParents(value: unknown, parent: object | undefined): void {
  walk(value, (v, parentOfV) => { Jsonix.Util.setParent(v, parentOfV); }, parent);
}
