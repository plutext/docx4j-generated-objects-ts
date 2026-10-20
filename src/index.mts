// The facade of @docx4j/generated-objects-ts: docx4j's names over the Jsonix runtime and the generated
// Office Open XML mappings in ../modules. Hand-written; the list of modules is ./modules.mts, which
// generate.sh does not write either, so add or remove an import there when modules appear or
// disappear (the smoke test fails on a stale registry).
import { Jsonix } from '@docx4j/jsonix';
import { MODULES } from './modules.mjs';
import { WML_GLOBAL_ELEMENTS, org_docx4j_wml_globals } from './wml-globals.mjs';
export { MODULES };
export { WML_GLOBAL_ELEMENTS, org_docx4j_wml_globals };
export { Jsonix };
export type { TypedNamedValue, XmlQName, XmlCalendar, XmlDuration, JsonixMapping } from '../modules/org_docx4j_wml.mjs';
import type { PackageElement, Package, Part } from '../modules/org_docx4j_xmlPackage.mjs';
export type { PackageElement, Package, Part };

/**
 * Every generated module, one Jsonix mapping each; they reference each other, so a context needs
 * them all. The registry in `./modules.mjs` is the list (CR-005): these names are its keys, so a
 * module added or dropped by a regeneration is one edit there, not two here.
 */
export const MODULE_NAMES = Object.keys(MODULES) as readonly ModuleName[] & ModuleName[];

export type ModuleName = keyof typeof MODULES;

const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const SML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/**
 * docx4j's namespace prefix table (org.docx4j.jaxb.NamespacePrefixMappings.getPreferredPrefixStatic,
 * docx4j VERSION_17_1_1, same order, minus the entries docx4j comments out), namespace URI to
 * prefix; the context default (CR-001). One namespace may be the default (''): SpreadsheetML, as
 * docx4j. A Relationships root instead gets its own namespace as the default and SpreadsheetML
 * 's' (docx4j's NamespacePrefixMapperRelationshipsPart and its requirePrefix fallback); see
 * `namespacePrefixesFor`. To extend: getContext({ namespacePrefixes: { ...NAMESPACE_PREFIXES, ... } })
 * before first use; a table passed to getContext replaces this one.
 */
export const NAMESPACE_PREFIXES: Readonly<Record<string, string>> = Object.freeze({
  'http://www.w3.org/XML/1998/namespace': 'xml',
  'http://www.w3.org/2001/XMLSchema-instance': 'xsi',
  'http://www.w3.org/2001/XMLSchema': 'xs',
  'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
  'http://schemas.openxmlformats.org/package/2006/relationships': 'rel',    // the default namespace when the root is Relationships, see namespacePrefixesFor
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main': 'w',
  'http://schemas.openxmlformats.org/presentationml/2006/main': 'p',
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main': '',          // SML uses the default namespace, as docx4j; 's' when the root is Relationships
  'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
  'http://schemas.openxmlformats.org/drawingml/2006/picture': 'pic',
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing': 'wp',
  'http://schemas.openxmlformats.org/drawingml/2006/chart': 'c',
  'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing': 'cdr',
  'http://schemas.openxmlformats.org/drawingml/2006/diagram': 'dgm',
  'http://schemas.openxmlformats.org/drawingml/2006/compatibility': 'comp',
  'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas': 'lc',
  'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing': 'xdr',
  'http://schemas.openxmlformats.org/officeDocument/2006/math': 'm',
  'http://schemas.openxmlformats.org/officeDocument/2006/bibliography': 'b',
  'http://schemas.openxmlformats.org/officeDocument/2006/customXml': 'ds',
  'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties': 'prop',
  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties': 'properties',
  'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes': 'vt',
  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties': 'cp',
  'http://schemas.openxmlformats.org/package/2006/digital-signature': 'mdssi',
  'http://schemas.openxmlformats.org/schemaLibrary/2006/main': 'sl',
  'http://purl.org/dc/elements/1.1/': 'dc',
  'http://purl.org/dc/terms/': 'dcterms',
  'urn:schemas-microsoft-com:vml': 'v',
  'urn:schemas-microsoft-com:office:office': 'o',
  'urn:schemas-microsoft-com:office:word': 'w10',
  'urn:schemas-microsoft-com:office:excel': 'xvml',
  'urn:schemas-microsoft-com:office:powerpoint': 'pvml',
  'urn:schemas-microsoft-com:mac:vml': 'mv',
  'http://schemas.microsoft.com/office/2006/xmlPackage': 'pkg',
  'http://schemas.microsoft.com/office/2006/coverPageProps': 'cppr',
  'http://schemas.microsoft.com/office/2006/digsig': 'dssi',
  'http://schemas.microsoft.com/office/word/2003/auxHint': 'WX',
  'http://schemas.microsoft.com/office/word/2006/wordml': 'wne',
  'http://schemas.microsoft.com/office/word/2010/wordml': 'w14',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing': 'wp14',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas': 'wpc',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup': 'wpg',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingInk': 'wpi',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape': 'wps',
  'http://schemas.microsoft.com/office/word/2012/wordml': 'w15',
  'http://schemas.microsoft.com/office/word/2012/wordprocessingDrawing': 'wp15',
  'http://schemas.microsoft.com/office/word/2015/wordml/symex': 'w16se',
  'http://schemas.microsoft.com/office/word/2016/wordml/cid': 'w16cid',
  'http://schemas.microsoft.com/office/word/2018/wordml': 'w16',
  'http://schemas.microsoft.com/office/word/2018/wordml/cex': 'w16cex',
  'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash': 'w16sdtdh',
  'http://schemas.microsoft.com/office/word/2023/wordml/word16du': 'w16du',
  'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock': 'w16sdtfl',
  'http://schemas.microsoft.com/office/drawing/2007/8/2/chart': 'c14',
  'http://schemas.microsoft.com/office/drawing/2008/diagram': 'dsp',
  'http://schemas.microsoft.com/office/drawing/2010/main': 'a14',
  'http://schemas.microsoft.com/office/drawing/2010/picture': 'pic14',
  'http://schemas.microsoft.com/office/drawing/2010/chartDrawing': 'cdr14',
  'http://schemas.microsoft.com/office/drawing/2010/diagram': 'dgm14',
  'http://schemas.microsoft.com/office/drawing/2012/main': 'a15',
  'http://schemas.microsoft.com/office/drawing/2012/chart': 'c15',
  'http://schemas.microsoft.com/office/drawing/2012/chartStyle': 'cs',
  'http://schemas.microsoft.com/office/drawing/2013/main/command': 'a13cmd',
  'http://schemas.microsoft.com/office/drawing/2014/main': 'a16',
  'http://schemas.microsoft.com/office/drawing/2014/chart': 'c16',
  'http://schemas.microsoft.com/office/drawing/2014/chart/ac': 'c16ac',
  'http://schemas.microsoft.com/office/drawing/2014/chartex': 'cx',
  'http://schemas.microsoft.com/office/drawing/2016/11/main': 'a1611',
  'http://schemas.microsoft.com/office/drawing/2016/11/diagram': 'dgm1611',
  'http://schemas.microsoft.com/office/drawing/2016/12/diagram': 'dgm1612',
  'http://schemas.microsoft.com/office/drawing/2016/ink': 'ink16',
  'http://schemas.microsoft.com/office/drawing/2016/SVG/main': 'a16svg',
  'http://schemas.microsoft.com/office/drawing/2017/03/chart': 'c16r3',
  'http://schemas.microsoft.com/office/drawing/2017/decorative': 'adec',
  'http://schemas.microsoft.com/office/drawing/2017/model3d': 'am3d',
  'http://schemas.microsoft.com/office/drawing/2018/animation': 'an18',
  'http://schemas.microsoft.com/office/drawing/2018/animation/model3d': 'anam3d',
  'http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor': 'a18hc',
  'http://schemas.microsoft.com/office/thememl/2012/main': 'thm15',
  'http://schemas.microsoft.com/office/powerpoint/2010/main': 'p14',
  'http://schemas.microsoft.com/office/powerpoint/2012/main': 'p15',
  'http://schemas.microsoft.com/office/powerpoint/2013/main/command': 'p13cmd',
  'http://schemas.microsoft.com/office/powerpoint/2014/inkAction': 'iact',
  'http://schemas.microsoft.com/office/powerpoint/2015/main': 'p16',
  'http://schemas.microsoft.com/office/powerpoint/2015/09/main': 'p159',
  'http://schemas.microsoft.com/office/powerpoint/2015/10/main': 'p1510',
  'http://schemas.microsoft.com/office/powerpoint/2016/6/main': 'p166',
  'http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom': 'psez',
  'http://schemas.microsoft.com/office/powerpoint/2016/slidezoom': 'pslz',
  'http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom': 'psuz',
  'http://schemas.microsoft.com/office/powerpoint/2017/3/main': 'p173',
  'http://schemas.microsoft.com/office/powerpoint/2017/10/main': 'p1710',
  'http://schemas.microsoft.com/office/powerpoint/2018/4/main': 'p184',
  'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main': 'x14',
  'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac': 'x14ac',
  'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main': 'x15',
  'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex': 'cx1',
  'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex': 'cx2',
  'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex': 'cx3',
  'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex': 'cx4',
  'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex': 'cx5',
  'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex': 'cx6',
  'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex': 'cx7',
  'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex': 'cx8',
  'http://schemas.microsoft.com/office/word/2026/wordml/cei': 'cei',
  'http://schemas.microsoft.com/office/comments/2020/reactions': 'cr',
  'http://schemas.microsoft.com/office/spreadsheetml/2015/02/main': 'x16r2',
  'http://schemas.microsoft.com/office/2019/extlst': 'oel',
  'http://schemas.microsoft.com/office/drawing/2010/slicer': 'sle',
  'http://schemas.microsoft.com/office/drawing/2012/slicer': 'sle15',
  'http://schemas.microsoft.com/office/drawing/2012/timeslicer': 'tsle',
  'http://schemas.microsoft.com/office/excel/2006/main': 'xm',
  'http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac': 'x12ac',
  'http://schemas.microsoft.com/office/spreadsheetml/2014/11/main': 'x16',
  'http://schemas.microsoft.com/office/spreadsheetml/2017/revision16': 'xr16',
  'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2': 'xr2',
  'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3': 'xr3',
  'http://schemas.microsoft.com/office/spreadsheetml/2016/revision6': 'xr6',
  'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10': 'xr10',
  'http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac': 'x15ac',
  'http://schemas.microsoft.com/office/spreadsheetml/2014/revision': 'xr',
  'http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing': 'xdr14',
  'http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11': 'wetp',
  'http://schemas.microsoft.com/office/webextensions/webextension/2010/11': 'we',
  'http://schemas.microsoft.com/ink/2010/main': 'msink',
  'http://schemas.microsoft.com/aml/2001/core': 'aml',
  'http://uri.etsi.org/01903/v1.3.2#': 'xd',
  'http://opendope.org/xpaths': 'odx',
  'http://opendope.org/conditions': 'odc',
  'http://opendope.org/components': 'odi',
  'http://opendope.org/questions': 'odq',
  'http://opendope.org/answers': 'oda',
  'http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures': 'xcalcf',
  'http://opendope.org/SmartArt/DataHierarchy': 'odgm',
});

/**
 * Prefixes Office names in a root's `mc:Ignorable` that `NAMESPACE_PREFIXES` cannot produce,
 * because that namespace is written as the default (CR-006). Prefix to namespace URI, the
 * opposite direction to `NAMESPACE_PREFIXES`, since this is the direction the question is asked
 * in: an `mc:Ignorable` token is a prefix, and the marshaller must declare it or drop it.
 *
 * Excel binds `x` to the SpreadsheetML main namespace on its slicer, slicer cache and timeline
 * parts (whose roots are x14 and x15) and writes `mc:Ignorable="x xr10"`; this package writes
 * that namespace as the default, so the prefix is declared beside the default declaration, both
 * being in scope. docx4j does the same through `NamespacePrefixMapperUtils
 * .getPreDeclaredNamespaceUris2` and `NamespacePrefixMappings.getNamespaceURIStatic` (its CR-024).
 *
 * One entry, because one namespace is written as the default.
 */
export const IGNORABLE_PREFIX_ALIASES: Readonly<Record<string, string>> = Object.freeze({
  x: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
});

/**
 * Office's extension namespaces, which the facade's marshal functions declare ignorable on a part
 * root whose tree uses them (CR-008): ECMA-376 Part 3 makes markup in a namespace a reader does not
 * understand an error unless the namespace is declared ignorable, and Word declares every one it
 * knows. A root whose type binds `mc:Ignorable` gains, for each of these its tree uses and its list
 * does not already name, the prefix the call's table gives; prefixes it already lists are kept, in
 * order.
 *
 * Being listed here is a claim that a reader may ignore the namespace, so the list is the
 * namespaces Office itself names in a part root's `mc:Ignorable` (measured on the fidelity corpus),
 * Word's in the order Word writes them, then Excel's. Nothing Office does not list: not x15ac,
 * which Excel writes only inside an `mc:Choice` requiring x15; not a14, a16, c14 or p14, which
 * Office writes only inside such a Choice or an extension list. No PresentationML root binds
 * `mc:Ignorable` in any case.
 */
export const IGNORABLE_EXTENSION_NAMESPACES: readonly string[] = Object.freeze([
  'http://schemas.microsoft.com/office/word/2010/wordml',                   // w14
  'http://schemas.microsoft.com/office/word/2012/wordml',                   // w15
  'http://schemas.microsoft.com/office/word/2015/wordml/symex',             // w16se
  'http://schemas.microsoft.com/office/word/2016/wordml/cid',               // w16cid
  'http://schemas.microsoft.com/office/word/2018/wordml',                   // w16
  'http://schemas.microsoft.com/office/word/2018/wordml/cex',               // w16cex
  'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',       // w16sdtdh
  'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock',     // w16sdtfl
  'http://schemas.microsoft.com/office/comments/2020/reactions',            // cr (Word lists it on commentsExtensible)
  'http://schemas.microsoft.com/office/word/2023/wordml/word16du',          // w16du
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',    // wp14
  'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',            // x14ac
  'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',         // x15
  'http://schemas.microsoft.com/office/spreadsheetml/2015/02/main',         // x16r2
  'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',        // xr
  'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2',       // xr2
  'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3',       // xr3
  'http://schemas.microsoft.com/office/spreadsheetml/2016/revision6',       // xr6
  'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10',      // xr10
  'http://schemas.microsoft.com/office/spreadsheetml/2017/revision16',      // xr16
]);

const EXTENSION_NAMESPACES: ReadonlySet<string> = new Set(IGNORABLE_EXTENSION_NAMESPACES);

/**
 * `getContext` options: the runtime's, plus `modules` for a context over fewer mappings than the
 * whole model (CR-005) - the mapping objects themselves, or a function returning them, so a caller
 * that wants wml alone can import those modules and nothing else is pulled in. The option is
 * stripped before the rest is passed to the runtime.
 *
 * Like every option here it is read only when the context is BUILT: `getContext` returns the one
 * it has built for every later call, so a caller that wants a different context calls
 * `resetContext()` first. An add-in asking for a small context is usually not the first caller.
 */
export interface ContextOptions extends Jsonix.ContextOptions {
  modules?: readonly unknown[] | (() => Promise<readonly unknown[]> | readonly unknown[]);
}

/**
 * The modules named, plus everything they depend on, transitively (CR-005): the mappings reference
 * each other, so a context over a hand-picked list usually fails at build time with "Type info
 * [...] is not known in this context" - a list of wml and the obvious neighbours misses the a14
 * module that `CT_TextParagraph` reached after docx4j CR-021, for instance.
 *
 * Each generated mapping declares its own `dependencies`, so the closure is computed rather than
 * guessed: `getContext({ modules: modulesFor('org_docx4j_wml') })` is the usable form of the
 * option. Note that this makes a SMALLER CONTEXT, not a smaller bundle - a bundler that sees the
 * facade takes the registry with it either way.
 */
export function modulesFor(...names: ModuleName[]): Jsonix.Mapping[] {
  const chosen = new Map<string, Jsonix.Mapping>();
  const add = (name: string): void => {
    if (chosen.has(name)) return;
    const module = (MODULES as Record<string, Jsonix.Mapping | undefined>)[name];
    if (module === undefined) throw new Error(`${name} is not a module of this package (see MODULE_NAMES)`);
    chosen.set(name, module);
    for (const dependency of referencedModules(module)) add(dependency);
  };
  for (const name of names) add(name);
  const closure = [...chosen.values()];
  // CR-009: a closure holding wml gets the global declarations it lacks, as the default context does.
  return chosen.has('org_docx4j_wml') ? [...closure, org_docx4j_wml_globals] : closure;
}

/**
 * The modules a mapping refers to, read from the type references themselves
 * (`typeInfo: 'org_docx4j_dml.CTTextCharacterProperties'`) rather than from its `dependencies`
 * field, which the compiler under-reports: `org_docx4j_wml` names six there and refers to eight,
 * missing `org_docx4j_dml_wordprocessingDrawing` and `org_docx4j_w15symex`, so a closure built
 * from `dependencies` fails at context build. Reported to the compiler; reading the references is
 * right either way, since they are what the runtime resolves.
 */
function referencedModules(module: Jsonix.Mapping): Set<string> {
  const found = new Set<string>();
  const seen = new Set<unknown>();
  const visit = (value: unknown): void => {
    if (typeof value === 'string') {
      const dot = value.indexOf('.');
      if (dot > 0 && value.startsWith('org_')) {
        const name = value.slice(0, dot);
        if (name in MODULES && name !== module.name) found.add(name);
      }
      return;
    }
    if (typeof value !== 'object' || value === null || seen.has(value)) return;
    seen.add(value);
    for (const each of Object.values(value)) visit(each);
  };
  visit(module);
  for (const declared of (module as { dependencies?: readonly string[] }).dependencies ?? []) {
    if (declared in MODULES) found.add(declared);
  }
  return found;
}

let contextPromise: Promise<Jsonix.Context> | undefined;
let builtContext: Jsonix.Context | undefined;

/**
 * The context over all Office Open XML mappings, built once on first use (the modules are
 * loaded lazily; 104 modules are several megabytes), plus `org_docx4j_wml_globals` (CR-009: the
 * WordprocessingML elements docx4j declares as roots, so that a kept `mc:Choice` or `mc:Fallback`
 * types its `w:drawing` or `w:pict` as docx4j does). Parent pointers are on by default, as in
 * docx4j's object model, and so is docx4j's prefix table (`NAMESPACE_PREFIXES`); pass options to
 * change that or anything else. A `modules` list of your own gets the globals only if it holds
 * them: `modulesFor` adds them to any closure containing `org_docx4j_wml`.
 */
export function getContext(options?: ContextOptions): Promise<Jsonix.Context> {
  if (contextPromise === undefined) {
    const { modules, ...contextOptions } = options ?? {};
    const mappings = Promise.resolve(typeof modules === 'function' ? modules() : modules ?? [...Object.values(MODULES), org_docx4j_wml_globals]);
    contextPromise = mappings.then((loaded) => {
      builtContext = new Jsonix.Context(loaded as Jsonix.Mapping[], { parentPointers: true, namespacePrefixes: NAMESPACE_PREFIXES, ...contextOptions });
      return builtContext;
    });
  }
  return contextPromise;
}

/** Forgets the shared context, so that the next call to getContext builds a new one (e.g. with other options). */
export function resetContext(): void {
  contextPromise = undefined;
  builtContext = undefined;
}

/**
 * The context, for a caller that cannot await: it must already be built, which it is once anything
 * has been unmarshalled or marshalled through the facade. Throws otherwise, since building it loads
 * the modules and that is asynchronous. `getContext()` is the form to prefer.
 */
export function getContextSync(): Jsonix.Context {
  if (builtContext === undefined) {
    throw new Error('The context is not built yet: await getContext() (or any unmarshal/marshal) once first');
  }
  return builtContext;
}

/**
 * docx4j XmlUtils.unmarshalString: parse XML into a typed element, e.g. unmarshalString<DocumentElement>(xml).
 * The source's prefix bindings that `NAMESPACE_PREFIXES` cannot give are recorded for the marshal
 * (CR-011), so the text is parsed here, as the runtime's own unmarshalString does.
 *
 * `options` are for this call only (CR-007): `onUnexpectedElement` and `onUnexpectedAttribute`
 * report what the model dropped, at the point it was dropped. They are never set by default - see
 * the note on `UnmarshallerOptions`. An `onElement` (jsonix-CR-008) is called after the facade's
 * own, which is how a Choice's bindings are recorded.
 */
export async function unmarshalString<E extends Jsonix.TypedNamedValue = Jsonix.TypedNamedValue>(
  xml: string, options?: Jsonix.UnmarshallerOptions,
): Promise<E> {
  const context = await getContext();
  const doc = Jsonix.DOM.parse(xml);
  return recordRootBindings(doc, context.createUnmarshaller(recordingOptions(options)).unmarshalDocument<E>(doc));
}

/** docx4j XmlUtils.marshaltoString: serialise an element to XML, with docx4j's prefixes and namespace declarations. */
export async function marshalString(element: Jsonix.TypedNamedValue, options?: Jsonix.MarshallerOptions): Promise<string> {
  return serialize(marshalToDocument(await getContext(), element, options));
}

/** docx4j XmlUtils.unwrap: the value of an element. */
export function unwrap<T>(element: Jsonix.TypedNamedValue<T>): T {
  return element.value;
}

/** docx4j XmlUtils.deepCopy: a structural copy with PARENT re-linked (Jsonix.Util.deepCopy), each object given its original's CR-011 record. */
export function deepCopy<T>(value: T, parent?: unknown): T {
  const copy = Jsonix.Util.deepCopy(value, parent);
  if (recorded) carryBindings(value, copy, new Set());
  return copy;
}

/** The copy has the original's shape, so each record is carried to the object in the same place. */
function carryBindings(original: unknown, copy: unknown, seen: Set<object>): void {
  if (typeof original !== 'object' || original === null || typeof copy !== 'object' || copy === null || isNode(original) || seen.has(original)) return;
  seen.add(original);
  const bindings = sourceBindings.get(original);
  if (bindings !== undefined) sourceBindings.set(copy, bindings);
  for (const key of Object.keys(original)) {
    if (key !== 'PARENT') carryBindings((original as Record<string, unknown>)[key], (copy as Record<string, unknown>)[key], seen);
  }
}

/**
 * The property names a type declares, its bases included, from the context's mapping model.
 * (`TypeInfo.properties` is not in the runtime's typings, hence the cast; a candidate for a
 * jsonix typings CR, as `Jsonix.DOM` was.)
 */
function declaredProperties(typeInfo: Jsonix.TypeInfo | undefined): Set<string> {
  const names = new Set<string>();
  for (let ti = typeInfo; ti; ti = ti.baseTypeInfo) {
    for (const property of (ti as unknown as { properties?: { name: string }[] }).properties ?? []) names.add(property.name);
  }
  return names;
}

/**
 * `deepCopy`, then the copy is typed as `typeName`: its `TYPE_NAME` is set and the properties that
 * type does not declare are removed. `typeName` must be the value's own type or one of its bases.
 *
 * docx4j does this with a hand-built object: a `w:pPrChange` holds a `PPrBase`, not the `PPr` a
 * paragraph has, and a copy that keeps `PPr` marshals as `<w:pPr xsi:type="w:CT_PPr">`, which is
 * valid but not what Word writes (CR-003 section 2). `deepCopyAsSync` is the same for a caller
 * that cannot await and knows the context is built.
 */
export async function deepCopyAs<T extends { TYPE_NAME?: string }>(value: object, typeName: NonNullable<T['TYPE_NAME']>, parent?: unknown): Promise<T> {
  await getContext();
  return deepCopyAsSync<T>(value, typeName, parent);
}

/**
 * `deepCopyAs` for a caller that cannot await (`@docx4j/core-ts`'s `recordPPrChange` runs inside
 * the synchronous property setters of its Office JS-shaped views). The context must already be
 * built, as it is once the document has been unmarshalled; otherwise this throws.
 */
export function deepCopyAsSync<T extends { TYPE_NAME?: string }>(value: object, typeName: NonNullable<T['TYPE_NAME']>, parent?: unknown): T {
  const context = getContextSync();
  const from = (value as { TYPE_NAME?: string }).TYPE_NAME;
  const target = context.getTypeInfoByName(typeName as string);
  if (!target) throw new Error(`Not a type in this context: ${String(typeName)}`);
  let ok = from === undefined;
  for (let ti: Jsonix.TypeInfo | undefined = from === undefined ? undefined : context.getTypeInfoByName(from); ti && !ok; ti = ti.baseTypeInfo) {
    if (ti.name === typeName) ok = true;
  }
  if (!ok) throw new Error(`${String(typeName)} is not ${from} or one of its base types`);
  const copy = deepCopy(value, parent) as Record<string, unknown>;
  const declared = declaredProperties(target);
  for (const key of Object.keys(copy)) {
    if (key === 'TYPE_NAME' || key === 'PARENT') continue;
    if (!declared.has(key)) delete copy[key];
  }
  copy.TYPE_NAME = typeName as string;
  return copy as T;
}

/** docx4j XmlUtils.unmarshal(Node): a DOM element or document, through the shared context, recording the source's bindings as `unmarshalString` does. */
export async function unmarshalNode<E extends Jsonix.TypedNamedValue = Jsonix.TypedNamedValue>(
  node: Node, options?: Jsonix.UnmarshallerOptions,
): Promise<E> {
  return recordRootBindings(node, (await getContext()).createUnmarshaller(recordingOptions(options)).unmarshalDocument<E>(node));
}

/** docx4j XmlUtils.marshaltoW3CDomDocument: the element as a DOM element, with docx4j's prefixes and namespace declarations. */
export async function marshalNode(element: Jsonix.TypedNamedValue, options?: Jsonix.MarshallerOptions): Promise<Element> {
  return marshalToDocument(await getContext(), element, options).documentElement;
}

/**
 * The prefix bindings a source declared, kept for the marshal (CR-011), since the typed tree keeps
 * no namespace declarations. Two kinds of object have a record: an `mc:Choice`, holding the
 * binding in scope there of each prefix its `Requires` names, so that it travels with the Choice
 * wherever an edit moves it; and a part's root value, holding the bindings in scope at the source
 * root, for its `mc:Ignorable`.
 */
const sourceBindings = new WeakMap<object, ReadonlyMap<string, string>>();
// Until a first record exists, a marshal adds no onElement and a deepCopy no walk.
let recorded = false;

const CHOICE_TYPE = 'org_docx4j_mce.AlternateContent.Choice';

const TABLE_BY_PREFIX: ReadonlyMap<string, string> = new Map(Object.entries(NAMESPACE_PREFIXES).map(([namespaceURI, prefix]) => [prefix, namespaceURI]));

/** Only what the table and the aliases would not give is worth recording: Office's own files leave no record. */
function sourceOnly(prefix: string, namespaceURI: string): boolean {
  return (TABLE_BY_PREFIX.get(prefix) ?? IGNORABLE_PREFIX_ALIASES[prefix]) !== namespaceURI;
}

function requiresOf(element: Element): string[] {
  return (element.getAttribute('Requires') ?? '').split(/\s+/).filter(Boolean);
}

/** The binding in scope at `element` of each prefix its Requires names, or undefined when it names none the source binds. */
function choiceBindingsAt(element: Element): Map<string, string> | undefined {
  let bindings: Map<string, string> | undefined;
  for (const prefix of requiresOf(element)) {
    const namespaceURI = element.lookupNamespaceURI(prefix);
    if (namespaceURI && sourceOnly(prefix, namespaceURI)) (bindings ??= new Map()).set(prefix, namespaceURI);
  }
  return bindings;
}

/**
 * The unmarshaller options with this package's `onElement` (jsonix-CR-008) composed in front of the
 * caller's: each typed mc:Choice records its source-only bindings as it is built.
 */
function recordingOptions(options?: Jsonix.UnmarshallerOptions): Jsonix.UnmarshallerOptions {
  const callerOnElement = options?.onElement;
  return {
    ...options,
    onElement(element: Element, value: object): void {
      if ((value as { TYPE_NAME?: string }).TYPE_NAME === CHOICE_TYPE) {
        const bindings = choiceBindingsAt(element);
        if (bindings) { sourceBindings.set(value, bindings); recorded = true; }
      }
      callerOnElement?.(element, value);
    },
  };
}

function rootValueOf(element: unknown): unknown {
  return typeof element === 'object' && element !== null && 'name' in element && 'value' in element ? (element as Jsonix.TypedNamedValue).value : element;
}

/** Records against `element`'s value the bindings in scope at the source root: its own declarations and its ancestors' (a flat OPC part inherits pkg:package's). */
function recordRootBindings<E extends Jsonix.TypedNamedValue>(node: Node, element: E): E {
  const root = node.nodeType === 9 ? (node as Document).documentElement : node as Element;
  const value = rootValueOf(element);
  if (!root || typeof value !== 'object' || value === null) return element;
  const bindings = new Map<string, string>();
  const seen = new Set<string>();
  for (let n: Node | null = root; n && n.nodeType === 1; n = n.parentNode) {
    for (const attr of Array.from((n as Element).attributes)) {
      const prefix = prefixOfDeclaration(attr);
      if (prefix && !seen.has(prefix)) {
        seen.add(prefix);
        if (sourceOnly(prefix, attr.value)) bindings.set(prefix, attr.value);
      }
    }
  }
  if (bindings.size > 0) { sourceBindings.set(value, bindings); recorded = true; }
  return element;
}

function isNode(value: unknown): value is Node {
  return typeof value === 'object' && value !== null && typeof (value as Node).nodeType === 'number';
}

/**
 * The prefix table for one marshal, as docx4j chooses a NamespacePrefixMapper per part: the
 * context's table, except that a Relationships root makes its namespace the default (what Word
 * writes) and SpreadsheetML takes docx4j's requirePrefix prefix, since the runtime can bind only
 * one namespace to the default prefix.
 */
function namespacePrefixesFor(context: Jsonix.Context, root: Jsonix.TypedNamedValue, base?: Record<string, string>): Record<string, string> {
  const table: Record<string, string> = { ...(base ?? context.namespacePrefixes) };
  if (root.name.namespaceURI === RELATIONSHIPS_NS) {
    table[RELATIONSHIPS_NS] = '';
    if (table[SML_NS] === '') table[SML_NS] = 's';
  }
  return table;
}

/**
 * Marshals through the context with the root-specific table, then declares on the root only what
 * the tree uses (CR-001 section 3): the runtime declares every entry of the table on the root
 * element, used or not.
 *
 * The context is not modified - the table and any other per-call option go to
 * `createMarshaller(options)`, which derives a context of its own (CR-007; until
 * `@docx4j/jsonix` 3.4.0 this package derived one by hand with `Object.create`). A caller's own
 * `namespacePrefixes` replaces the context's as the base, and the per-root rule is then applied to
 * it, since choosing a table per root is this package's decision rather than the runtime's.
 * jsonix-CR-003 part 2, still deferred there, is what would retire `fixRootNamespaceDeclarations`.
 *
 * Before that, a part root that binds `mc:Ignorable` lists the extension namespaces its tree uses
 * (CR-008). The markup-compatibility elements bind it too, but are never a part's root.
 *
 * Once any CR-011 record exists, the marshaller's `onElement`, composed in front of the caller's,
 * notes each written Choice that has a record. Wildcard DOM content is looked up through its
 * original node, which is still in its source document.
 */
function marshalToDocument(context: Jsonix.Context, element: Jsonix.TypedNamedValue, options?: Jsonix.MarshallerOptions): Document {
  const table = namespacePrefixesFor(context, element, options?.namespacePrefixes);
  const choices = new Map<Element, ReadonlyMap<string, string>>();
  const callerOnElement = options?.onElement;
  const onElement = !recorded ? callerOnElement : (written: Node, value: unknown): void => {
    if (isNode(value)) {
      // Wildcard DOM content: its original is still in its source document, which answers for its Choices.
      const originals = (value as Element).getElementsByTagNameNS?.(MC_NS, 'Choice');
      const copies = (written as Element).getElementsByTagNameNS?.(MC_NS, 'Choice');
      for (let i = 0; originals && copies && i < originals.length && i < copies.length; i++) {
        const bindings = choiceBindingsAt(originals[i]);
        if (bindings) choices.set(copies[i], bindings);
      }
    } else if (typeof value === 'object' && value !== null) {
      const bindings = sourceBindings.get(value);
      if (bindings !== undefined && (value as { TYPE_NAME?: string }).TYPE_NAME === CHOICE_TYPE) choices.set(written as Element, bindings);
    }
    callerOnElement?.(written, value);
  };
  const doc = context.createMarshaller({ ...options, namespacePrefixes: table, onElement }).marshalDocument(element);
  if (element.name.namespaceURI !== MC_NS && bindsIgnorable(context, element)) declareExtensionsIgnorable(doc.documentElement, table);
  const rootValue = rootValueOf(element);
  fixRootNamespaceDeclarations(doc.documentElement, table, typeof rootValue === 'object' && rootValue !== null ? sourceBindings.get(rootValue) : undefined, choices);
  return doc;
}

/**
 * Whether the root's type binds `mc:Ignorable` (CR-008 section 4): the type its `TYPE_NAME` names,
 * or for a literal without one, the type its element name declares. A root that does not is left
 * alone, since the schema does not admit the attribute there and Office rejects a part that has it.
 */
function bindsIgnorable(context: Jsonix.Context, element: Jsonix.TypedNamedValue): boolean {
  const typeName = (element.value as { TYPE_NAME?: unknown } | null)?.TYPE_NAME;
  // getElementInfo reads only the name's key; its typings say string.
  const key = (element.name.namespaceURI ? `{${element.name.namespaceURI}}` : '') + element.name.localPart;
  const typeInfo: Jsonix.TypeInfo | undefined = typeof typeName === 'string'
    ? context.getTypeInfoByName(typeName)
    : context.getElementInfo({ key } as unknown as string, undefined as unknown as string)?.typeInfo;
  for (let ti = typeInfo; ti; ti = ti.baseTypeInfo) {
    for (const property of (ti as unknown as { properties?: { attributeName?: Jsonix.XML.QName | null }[] }).properties ?? []) {
      if (property.attributeName?.namespaceURI === MC_NS && property.attributeName.localPart === 'Ignorable') return true;
    }
  }
  return false;
}

/**
 * Appends to the root's `mc:Ignorable` the table's prefix for each of
 * `IGNORABLE_EXTENSION_NAMESPACES` the tree uses and the list does not already name (CR-008);
 * `fixRootNamespaceDeclarations` then declares it. The list is unioned, never replaced (CR-001
 * section 7), and new prefixes follow the old in the order of `IGNORABLE_EXTENSION_NAMESPACES`.
 *
 * Used means where ECMA-376 Part 3 asks a reader to understand it, less three places Office relies
 * on something else, so that a part Office wrote gains nothing:
 * - the root's own namespace (Excel's x15 timeline parts do not list x15; a reader that does not
 *   understand it has no reason to open the part);
 * - below an `mc:Choice` whose `Requires` names it, a branch only a reader that understands it takes
 *   (why Word never lists wps);
 * - below an extension list's `ext`, which a reader skips by its `uri` (Excel's styles carry x15
 *   there and do not list it).
 */
function declareExtensionsIgnorable(root: Element, table: Record<string, string>): void {
  const namespaceFor = new Map(Object.entries(table).map(([namespaceURI, prefix]) => [prefix, namespaceURI]));
  const resolve = (prefix: string, scope: ReadonlyMap<string, string>): string | undefined =>
    scope.get(prefix) ?? namespaceFor.get(prefix) ?? IGNORABLE_PREFIX_ALIASES[prefix];
  const scopeOf = (element: Element, outer: ReadonlyMap<string, string>): ReadonlyMap<string, string> => {
    let scope: Map<string, string> | undefined;
    for (const attr of Array.from(element.attributes)) {
      const prefix = prefixOfDeclaration(attr);
      if (prefix !== undefined) (scope ??= new Map(outer)).set(prefix, attr.value);
    }
    return scope ?? outer;
  };
  const used = new Set<string>();
  const note = (namespaceURI: string | null, required: ReadonlySet<string>): void => {
    if (namespaceURI && EXTENSION_NAMESPACES.has(namespaceURI) && namespaceURI !== root.namespaceURI && !required.has(namespaceURI)) used.add(namespaceURI);
  };
  const walk = (element: Element, outer: ReadonlyMap<string, string>, required: ReadonlySet<string>): void => {
    const scope = scopeOf(element, outer);
    let within = required;
    if (element.namespaceURI === MC_NS && element.localName === 'Choice') {
      const named = new Set(required);
      for (const prefix of (element.getAttribute('Requires') ?? '').split(/\s+/).filter(Boolean)) {
        const namespaceURI = resolve(prefix, scope);
        if (namespaceURI !== undefined) named.add(namespaceURI);
      }
      within = named;
    }
    note(element.namespaceURI, within);
    for (const attr of Array.from(element.attributes)) {
      if (prefixOfDeclaration(attr) === undefined) note(attr.namespaceURI, within);
    }
    if (element.localName === 'ext' && (element.parentNode as Element | null)?.localName === 'extLst') return;
    for (const child of Array.from(element.childNodes)) {
      if (child.nodeType === 1) walk(child as Element, scope, within);
    }
  };
  walk(root, new Map(), new Set());
  if (used.size === 0) return;

  const existing = root.getAttributeNodeNS(MC_NS, 'Ignorable');
  const listed = (existing?.value ?? '').split(/\s+/).filter(Boolean);
  const rootScope = scopeOf(root, new Map());
  for (const prefix of listed) {
    const namespaceURI = resolve(prefix, rootScope);
    if (namespaceURI !== undefined) used.delete(namespaceURI);
  }
  // A namespace the table writes as the default has no prefix to list, so it is left unlisted.
  const added = IGNORABLE_EXTENSION_NAMESPACES.filter((namespaceURI) => used.has(namespaceURI) && table[namespaceURI])
    .map((namespaceURI) => table[namespaceURI]);
  if (added.length === 0) return;
  const mc = existing?.prefix || table[MC_NS] || 'mc';
  if (!root.getAttributeNode(`xmlns:${mc}`)) root.setAttributeNS(XMLNS_NS, `xmlns:${mc}`, MC_NS);
  root.setAttributeNS(MC_NS, `${mc}:Ignorable`, [...listed, ...added].join(' '));
}

/** The runtime's serializer (xmldom in Node, XMLSerializer in browsers). */
function serialize(doc: Document): string {
  return Jsonix.DOM.serialize(doc);
}

function prefixOfDeclaration(attr: Attr): string | undefined {
  if (attr.name === 'xmlns') return '';
  if (attr.namespaceURI === XMLNS_NS || attr.name.startsWith('xmlns:')) return attr.localName ?? attr.name.slice(6);
  return undefined;
}

/**
 * The root's namespace declarations, as docx4j's McIgnorableNamespaceDeclarator leaves them:
 * every declaration that no element or attribute in the tree resolves to is removed unless the
 * root's mc:Ignorable or an mc:Choice's Requires names its prefix, and every prefix those name
 * that is not declared is added: from the source's own binding where one was recorded (CR-011:
 * `rootSource` for the part root, `choices` for each written Choice), else from `table`, else from
 * `IGNORABLE_PREFIX_ALIASES` for a prefix the table cannot produce because its namespace is
 * written as the default (CR-006).
 *
 * An mc:Ignorable prefix is declared on the root, as is a Requires prefix the table or the aliases
 * give. A Requires prefix only the source bound is declared on the mc:Choice that names it, ahead
 * of its Requires, unless the root already declares that binding (CR-012). That is Word's
 * placement: over 2,728 Word 365 saves it declared every Requires prefix of its own fixed set on
 * the root, and docx4j's private zz9 on the Choice, having read it on mc:AlternateContent; Word
 * reads all three places alike (core-ts, 2026-10-20). 0.4.3 declared it on the root.
 *
 * A Requires prefix counts because it names a namespace, not because anything is in it: Word's
 * chartex Choice says `Requires="cx1"` (2015/9/8) over content in `cx` (2014), so nothing in the
 * branch uses cx1 and its declaration went, leaving a Requires no reader can resolve (core-ts,
 * once its CR-010 phase 4 kept branches by default, 2026-10-19; docx4j CR-049 declares the same).
 * A declaration below the root that binds the prefix answers for its Choice, as everywhere here.
 * A Requires prefix nothing resolves is kept, with a warning: dropping it would change which
 * branch a reader takes.
 *
 * The source comes first because its binding is what the prefix meant: three Excel VML parts of
 * the fidelity corpus bind `x` to the Excel VML namespace, which the aliases would declare as
 * SpreadsheetML. A declaration the tree's own elements or attributes use is never rebound: where
 * the source meant another namespace by a prefix the root's tree uses, for mc:Ignorable, or the
 * Choice's own content uses, for a Requires, that is warned about. A use elsewhere in the tree
 * doesn't stop a Choice declaring its own binding, so a prefix bound two ways keeps both meanings.
 * A root declaration kept only because mc:Ignorable names it gives way to the source's.
 *
 * An mc:Ignorable prefix neither resolves is dropped
 * from mc:Ignorable with a warning: Word and Excel repair a file whose mc:Ignorable names an undeclared prefix, which is
 * how this was found (@docx4j/core-ts, an Excel acceptance run over xl/workbook.xml, 2026-09-19:
 * the model binds no xr:revisionPtr or xr2:uid, so nothing in the tree used xr2, xr6 or xr10 and
 * the marshaller declared none of them).
 *
 * `xmlns:xml` is never written. Declarations below the root shadow the root's, so a part's own
 * declarations inside a flat package count for the part, not for the package root.
 */
function fixRootNamespaceDeclarations(root: Element, table: Record<string, string>, rootSource?: ReadonlyMap<string, string>, choices?: ReadonlyMap<Element, ReadonlyMap<string, string>>): void {
  const declared = new Map<string, Attr>();
  for (const attr of Array.from(root.attributes)) {
    const prefix = prefixOfDeclaration(attr);
    if (prefix !== undefined) declared.set(prefix, attr);
  }
  const used = new Set<string>();
  const use = (prefix: string, namespaceURI: string, shadowed: Set<string>): void => {
    if (!shadowed.has(prefix) && declared.get(prefix)?.value === namespaceURI) used.add(prefix);
  };
  const required = new Set<string>();
  // The Requires prefixes only the source bound, each declared on its Choice (CR-012).
  const local: { choice: Element; prefix: string; namespaceURI: string }[] = [];
  const clashed = new Set<string>();
  const clash = (prefix: string, meant: string, tree: string): void => {
    if (clashed.has(prefix)) return;
    clashed.add(prefix);
    console.warn(`marshal: the source bound ${prefix} to ${meant}, but the tree uses it for ${tree}; mc:Ignorable or a Requires naming it now means the latter`);
  };
  const walk = (element: Element, shadowed: Set<string>): void => {
    let scope = shadowed;
    if (element !== root) {
      for (const attr of Array.from(element.attributes)) {
        const prefix = prefixOfDeclaration(attr);
        if (prefix !== undefined) {
          if (scope === shadowed) scope = new Set(shadowed);
          scope.add(prefix);
        }
      }
    }
    use(element.prefix ?? '', element.namespaceURI ?? '', scope);
    for (const attr of Array.from(element.attributes)) {
      if (attr.namespaceURI && prefixOfDeclaration(attr) === undefined) use(attr.prefix ?? '', attr.namespaceURI, scope);
    }
    if (element.namespaceURI === MC_NS && element.localName === 'Choice') {
      for (const prefix of requiresOf(element)) {
        if (scope.has(prefix)) continue;
        // What the prefix meant where it was required: the Choice's own source, else the root's.
        const namespaceURI = choices?.get(element)?.get(prefix) ?? rootSource?.get(prefix);
        if (namespaceURI !== undefined) {
          const tree = otherUse(element, prefix, namespaceURI);
          if (tree === undefined) {
            local.push({ choice: element, prefix, namespaceURI });
            continue;
          }
          clash(prefix, namespaceURI, tree);
        }
        required.add(prefix);
      }
    }
    for (const child of Array.from(element.childNodes)) {
      if (child.nodeType === 1) walk(child as Element, scope);   // childNodes: xmldom has no firstElementChild
    }
  };
  walk(root, new Set());
  const usedByTree = new Set(used);
  const ignorable = (root.getAttributeNS(MC_NS, 'Ignorable') ?? '').split(/\s+/).filter(Boolean);
  for (const prefix of ignorable) used.add(prefix);
  for (const prefix of required) used.add(prefix);
  for (const [prefix, attr] of declared) {
    // A declaration kept only because mc:Ignorable names its prefix gives way to the source's
    // binding, which is what that prefix meant (CR-011). One a Requires named has gone to its Choice.
    const meant = ignorable.includes(prefix) ? rootSource?.get(prefix) : undefined;
    const superseded = !usedByTree.has(prefix) && meant !== undefined && meant !== attr.value;
    if (!used.has(prefix) || attr.value === XML_NS || superseded) root.removeAttributeNode(attr);
  }
  if (ignorable.length > 0 || required.size > 0) declareOnRoot(root, table, ignorable, required, rootSource, clash);
  // After the root's, so that a binding the root declares isn't repeated on a Choice.
  for (const { choice, prefix, namespaceURI } of local) {
    if (root.getAttributeNode(`xmlns:${prefix}`)?.value === namespaceURI) continue;
    // Ahead of Requires, as Word writes it.
    const requires = choice.getAttributeNode('Requires');
    if (requires) choice.removeAttributeNode(requires);
    choice.setAttributeNS(XMLNS_NS, `xmlns:${prefix}`, namespaceURI);
    if (requires) choice.setAttributeNode(requires);
  }
}

/** The namespace other than `namespaceURI` that `element`'s content writes `prefix` for, where no declaration inside rebinds it. */
function otherUse(element: Element, prefix: string, namespaceURI: string): string | undefined {
  for (const child of Array.from(element.childNodes)) {
    if (child.nodeType !== 1) continue;   // childNodes: xmldom has no firstElementChild
    const inner = child as Element;
    const attrs = Array.from(inner.attributes);
    if (attrs.some((attr) => prefixOfDeclaration(attr) === prefix)) continue;
    if (inner.prefix === prefix && (inner.namespaceURI ?? '') !== namespaceURI) return inner.namespaceURI ?? '';
    const attr = attrs.find((attr) => attr.prefix === prefix && attr.namespaceURI && attr.namespaceURI !== namespaceURI);
    if (attr) return attr.namespaceURI!;
    const below = otherUse(inner, prefix, namespaceURI);
    if (below !== undefined) return below;
  }
  return undefined;
}

/** `fixRootNamespaceDeclarations`'s declarations on the root: the mc:Ignorable prefixes, and the Requires prefixes no source bound. */
function declareOnRoot(root: Element, table: Record<string, string>, ignorable: string[], required: ReadonlySet<string>, rootSource: ReadonlyMap<string, string> | undefined, clash: (prefix: string, meant: string, tree: string) => void): void {
  const namespaceFor = new Map(Object.entries(table).map(([namespaceURI, prefix]) => [prefix, namespaceURI]));
  // The table answers "which prefix for this namespace", which has no answer for the namespace
  // written as the default; the aliases answer "what does this prefix mean" (CR-006).
  // The source's own binding comes first: it is what the prefix meant (CR-011).
  const declare = (prefix: string): boolean => {
    const existing = root.getAttributeNode(`xmlns:${prefix}`);
    if (existing) {
      const meant = rootSource?.get(prefix);
      if (meant !== undefined && meant !== existing.value) clash(prefix, meant, existing.value);
      return true;
    }
    const namespaceURI = rootSource?.get(prefix) ?? namespaceFor.get(prefix) ?? IGNORABLE_PREFIX_ALIASES[prefix];
    if (namespaceURI === undefined) return false;
    root.setAttributeNS(XMLNS_NS, `xmlns:${prefix}`, namespaceURI);
    return true;
  };
  const unrequired = [...required].filter((prefix) => !declare(prefix));
  if (unrequired.length > 0) {
    console.warn(`marshal: an mc:Choice's Requires names ${unrequired.join(', ')}, which no declaration or NAMESPACE_PREFIXES entry resolves; kept`);
  }
  const unresolved = ignorable.filter((prefix) => !declare(prefix));
  if (unresolved.length > 0) {
    const kept = ignorable.filter((prefix) => !unresolved.includes(prefix));
    // An mc:Ignorable naming a prefix nothing declares is what Word and Excel repair, and the
    // table is the only place a declaration could come from.
    console.warn(`marshal: mc:Ignorable names ${unresolved.join(', ')}, which no declaration or NAMESPACE_PREFIXES entry resolves; dropped`);
    if (kept.length > 0) root.setAttributeNS(MC_NS, 'mc:Ignorable', kept.join(' '));
    else root.removeAttributeNS(MC_NS, 'Ignorable');
  }
}

/**
 * A flat OPC package, as Office JS `getOoxml()` returns it: `pkg:part/pkg:xmlData` content is
 * DOM in the schema (`processContents="skip"`); this unmarshals each XML part whose root element
 * the model knows (`w:document`, `w:styles`, relationships, ...) to a typed element, leaving the
 * rest as DOM, and records each part's source bindings as `unmarshalString` does (CR-011).
 * Reverse with `marshalPackage` before `insertOoxml()`.
 */
export async function unmarshalPackage(ooxml: string): Promise<PackageElement> {
  const pkg = await unmarshalString<PackageElement>(ooxml);
  const unmarshaller = (await getContext()).createUnmarshaller(recordingOptions());
  for (const part of pkg.value.part ?? []) {
    const any = part.xmlData?.any;
    if (part.xmlData && isNode(any)) {
      try {
        part.xmlData.any = recordRootBindings(any, unmarshaller.unmarshalDocument(any));
      } catch {
        // root element unknown to the model: keep the DOM
      }
    }
  }
  return pkg;
}

/** The inverse of `unmarshalPackage`: typed XML parts become DOM again, then the package is serialised. The input is not modified. */
export async function marshalPackage(pkg: PackageElement): Promise<string> {
  const context = await getContext();
  const parts: Part[] = (pkg.value.part ?? []).map((part) => {
    const any = part.xmlData?.any;
    if (part.xmlData && any && !isNode(any) && typeof any === 'object' && 'name' in any && 'value' in any) {
      return { ...part, xmlData: { ...part.xmlData, any: marshalToDocument(context, any as Jsonix.TypedNamedValue).documentElement } };
    }
    return part;
  });
  const value: Package = { ...pkg.value, part: parts };
  return serialize(marshalToDocument(context, { name: pkg.name, value }));
}
