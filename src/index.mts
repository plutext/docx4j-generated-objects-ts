// The facade of @docx4j/generated-objects-ts: docx4j's names over the Jsonix runtime and the generated
// Office Open XML mappings in ../modules. Hand-written; the list of modules is ./modules.mts, which
// generate.sh does not write either, so add or remove an import there when modules appear or
// disappear (the smoke test fails on a stale registry).
import { Jsonix } from '@docx4j/jsonix';
import { MODULES } from './modules.mjs';
import { WML_GLOBAL_ELEMENTS, org_docx4j_wml_globals } from './wml-globals.mjs';
import { wordColor, HIGHLIGHT_COLORS, WORD_INTEGER_STORAGE, wordIntegerReading, type WordIntegerReading, type WordIntegerStorage } from './helpers/wml.mjs';
import { CONTAINERS, misplacement, alternateContentOf, branchItems, type Declarations, type Misplacement, type MisplacedRule } from './misplaced.mjs';
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
const WML_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

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
 * own, which is how a Choice's bindings are recorded. `onInvalidValue` is `wordInvalidValue`
 * unless the caller passes one, so Word's own WordprocessingML loads as Word 365 reads it (CR-013),
 * and `onWordReading` is told of each value so read.
 */
export async function unmarshalString<E extends Jsonix.TypedNamedValue = Jsonix.TypedNamedValue>(
  xml: string, options?: UnmarshalOptions,
): Promise<E> {
  const context = await getContext();
  const doc = Jsonix.DOM.parse(xml);
  return recordRootBindings(doc, context.createUnmarshaller(recordingOptions(options)).unmarshalDocument<E>(doc));
}

/**
 * docx4j XmlUtils.marshaltoString: serialise an element to XML, with docx4j's prefixes and namespace declarations.
 * Run properties are written in their normal form (CR-013's normaliser), unless `normaliseOnSave` is false.
 */
export async function marshalString(element: Jsonix.TypedNamedValue, options?: MarshalOptions): Promise<string> {
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
  node: Node, options?: UnmarshalOptions,
): Promise<E> {
  return recordRootBindings(node, (await getContext()).createUnmarshaller(recordingOptions(options)).unmarshalDocument<E>(node));
}

/** docx4j XmlUtils.marshaltoW3CDomDocument: the element as a DOM element, with docx4j's prefixes and namespace declarations. */
export async function marshalNode(element: Jsonix.TypedNamedValue, options?: MarshalOptions): Promise<Element> {
  return marshalToDocument(await getContext(), element, options).documentElement;
}

/**
 * The prefix bindings a source declared, kept for the marshal (CR-011), since the typed tree keeps
 * no namespace declarations. Two kinds of object have a record: an `mc:Choice`, holding the
 * binding in scope there of each prefix its `Requires`, its own `mc:Ignorable` and its
 * `mc:MustUnderstand` name, so that it travels with the Choice wherever an edit moves it; and a
 * part's root value, holding the bindings in scope at the source root, for its `mc:Ignorable`.
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

type ChoiceAttribute = 'Requires' | 'mc:Ignorable' | 'mc:MustUnderstand';

/** The prefixes an mc:Choice names, each with the attribute naming it: its Requires, then its own mc:Ignorable and mc:MustUnderstand (docx4j CR-052's D1 (b)). */
function namedOnChoice(element: Element): [string, ChoiceAttribute][] {
  const named: [string, ChoiceAttribute][] = [];
  const lists: [ChoiceAttribute, string | null][] = [['Requires', element.getAttribute('Requires')],
    ['mc:Ignorable', element.getAttributeNS(MC_NS, 'Ignorable')], ['mc:MustUnderstand', element.getAttributeNS(MC_NS, 'MustUnderstand')]];
  for (const [attribute, value] of lists) {
    for (const prefix of (value ?? '').split(/\s+/).filter(Boolean)) named.push([prefix, attribute]);
  }
  return named;
}

/** The binding in scope at `element` of each prefix it names, or undefined when it names none the source binds. */
function choiceBindingsAt(element: Element): Map<string, string> | undefined {
  let bindings: Map<string, string> | undefined;
  for (const [prefix] of namedOnChoice(element)) {
    const namespaceURI = element.lookupNamespaceURI(prefix);
    if (namespaceURI && sourceOnly(prefix, namespaceURI)) (bindings ??= new Map()).set(prefix, namespaceURI);
  }
  return bindings;
}

/**
 * The unmarshaller options with this package's `onElement` (jsonix-CR-008) composed in front of the
 * caller's: each typed mc:Choice records its source-only bindings as it is built. And
 * `wordInvalidValue` as the `onInvalidValue` (CR-013), unless the caller passes one, reporting to
 * `onWordReading` when the caller asks; with it, jsonix-CR-012's `strictIntegers` (CR-015 phase 2),
 * so that an integer the runtime's `Number()` would read (`" 240 "`, `0x10`, `""`) reaches Word's
 * rule. A caller's own `onInvalidValue` keeps the runtime's reading unless it asks for
 * `strictIntegers` too.
 */
function recordingOptions(options?: UnmarshalOptions): Jsonix.UnmarshallerOptions {
  const { onWordReading, onOutOfRange, onMisplaced, ...runtimeOptions } = options ?? {};
  const callerOnElement = runtimeOptions.onElement;
  // CR-017: each typed value's source element, so a report can name a misplaced child's (its onElement comes first);
  // and the character data the runtime drops (jsonix-CR-013, 3.10.0), by its parent, reported with that container.
  const sources = onMisplaced ? new WeakMap<object, Element>() : undefined;
  const dropped = onMisplaced ? new WeakMap<Node, Text[]>() : undefined;
  const callerOnUnexpectedText = runtimeOptions.onUnexpectedText;
  return {
    ...runtimeOptions,
    ...(onMisplaced ? {
      onUnexpectedText(text: string, node: Node, classInfo: Jsonix.ClassInfo): void {
        const parent = node.parentNode;
        if (parent) { const texts = dropped!.get(parent); if (texts) texts.push(node as Text); else dropped!.set(parent, [node as Text]); }
        callerOnUnexpectedText?.(text, node, classInfo);
      },
    } : {}),
    onElement(element: Element, value: object): void {
      if ((value as { TYPE_NAME?: string }).TYPE_NAME === CHOICE_TYPE) {
        const bindings = choiceBindingsAt(element);
        if (bindings) { sourceBindings.set(value, bindings); recorded = true; }
      }
      if (onOutOfRange) reportOutOfRange(element, value, onOutOfRange);
      if (onMisplaced) { sources!.set(value, element); reportMisplacedOnRead(element, value, sources!, dropped!, onMisplaced); }
      callerOnElement?.(element, value);
    },
    onInvalidValue: runtimeOptions.onInvalidValue ?? (onWordReading ? reportingWordReadings(onWordReading) : wordInvalidValue),
    strictIntegers: runtimeOptions.strictIntegers ?? runtimeOptions.onInvalidValue === undefined,
  };
}

// ---------------------------------------------------------------------------------------------
// CR-013: attribute values read as Word 365 reads them
// ---------------------------------------------------------------------------------------------

/**
 * Word's own WordprocessingML mappings, whose refused attribute values `wordInvalidValue` reads as
 * Word does (CR-013 section 3.5; the extension mappings are core-ts CR-015's D10). CR-009's
 * `org_docx4j_wml_globals` declares elements of wml's types, so its class infos answer `org_docx4j_wml`.
 */
const WORD_READ_MODULES: ReadonlySet<string> = new Set(['org_docx4j_wml', 'org_docx4j_w14', 'org_docx4j_w15', 'org_docx4j_w15symex', 'org_docx4j_w16', 'org_docx4j_w16cex', 'org_docx4j_w16cid']);

/** What the runtime's boolean reads (jsonix-CR-007): exactly these, case and spacing as written. Also `w:u`'s on/off words. */
const BOOLEAN_FORMS: ReadonlySet<string> = new Set(['true', 'false', '1', '0', 'on', 'off']);

/** The runtime's integer types: `Jsonix.Schema.XSD.Integer` and its restrictions, by name (CR-015). */
const INTEGER_TYPES: ReadonlySet<string> = new Set(['Integer', 'Long', 'Int', 'Short', 'Byte', 'NonNegativeInteger',
  'NonPositiveInteger', 'NegativeInteger', 'PositiveInteger', 'UnsignedLong', 'UnsignedInt', 'UnsignedShort', 'UnsignedByte']);

/**
 * A number in an integer attribute, in the form Word 365 reads (CR-015 section 3.3, docx4j CR-065's
 * probe): a sign, digits and a fraction (`+240.5`, `240.`, `.5`), with no whitespace around it. Word
 * refuses the whole document for any other form: whitespace, a comma, hex, an exponent, letters.
 */
const WORD_NUMBER = /^[+-]?(\d+(\.\d*)?|\.\d+)$/;

/** Word's form with no digit: `""`, `.`, a sign, or a sign and a dot, which Word reads as 0 (round 6's group E). */
const NO_DIGITS = /^[+-]?\.?$/;

/**
 * An integer attribute's text as Word 365 reads it: `""` and the forms with no digit (`.`, `+`, `-`,
 * `+.`; docx4j CR-066's round 6, group E) as 0; Word's form as its integer part, truncated toward
 * zero (never `-0`); undefined for a form Word refuses.
 */
function wordInteger(text: string): number | undefined {
  if (NO_DIGITS.test(text)) return 0;
  if (!WORD_NUMBER.test(text)) return undefined;
  const value = Math.trunc(Number(text));
  return value === 0 ? 0 : value;
}

/**
 * The facade's `onInvalidValue` (jsonix-CR-010), set on every facade unmarshal unless the caller
 * passes one (CR-013 section 3.5): an attribute value of Word's own WordprocessingML that the type
 * refuses is read as Word 365 reads it, rather than failing the part.
 * - A refused boolean is on (`true`): Word reads only exactly `false`, `off` and `0` as off, and the
 *   runtime already reads those, so what reaches here (`yes`, `FALSE`, `" false"`, `""`) is on.
 *   That includes w14's `CT_OnOff` (`w14:checked`). w14's own `ST_OnOff` enumeration
 *   (`CT_LinearShadeProperties/@scaled`, `CT_StyleSet/@val`) is a string in the mapping, which
 *   writes it on the property where the runtime doesn't read it: never refused, held as written.
 * - A refused enumeration value is dropped, so the attribute is absent, as Word reads it (a style's
 *   value shows through); except `w:u`'s `true`, `false`, `on`, `off`, `1` and `0`, exactly, which
 *   Word reads as `'none'`.
 * - A refused integer (CR-015) in Word's form (`259.2`, `-120.5`, `+240.5`, `.5`, `10206.0`) is its
 *   integer part, truncated toward zero, as Word 365 reads it, and `""`, `.`, `+`, `-` and `+.` are 0. Any other form
 *   (`" 240 "`, `240,5`, `0x10`, `1e3`, `abc`), for which Word refuses the whole document, is
 *   dropped. The facade's unmarshal functions pass jsonix-CR-012's `strictIntegers`, so every text
 *   but an optional sign and digits arrives here.
 * - In every other mapping (DrawingML, VML, SpreadsheetML, ...), an integer `strictIntegers` refused
 *   is read by the runtime's own parse, as before it (`0x10` as 16, `""` as 0), or refused with its
 *   error. Anything else is refused as before (the error is rethrown): other simple types, whose
 *   Word reading is unmeasured, and every other type of every other mapping.
 * Exported for a caller with its own `onInvalidValue` that wants to compose it.
 */
export function wordInvalidValue(text: string, propertyInfo: Jsonix.AttributePropertyInfo, classInfo: Jsonix.ClassInfo, node: Element, error: Error): unknown {
  const typeInfo = propertyInfo.typeInfo;
  if (!WORD_READ_MODULES.has(classInfo.module.name)) {
    // jsonix-CR-012: strictIntegers refused what the runtime's own parse reads; read it as before.
    if (INTEGER_TYPES.has(typeInfo.name) && typeInfo.parse) return typeInfo.parse(text);
    throw error;
  }
  if (typeInfo.name === 'Boolean') return true;
  if (INTEGER_TYPES.has(typeInfo.name)) return wordInteger(text);
  if (typeInfo.CLASS_NAME !== 'Jsonix.Model.EnumLeafInfo') throw error;
  if (typeInfo.name === 'org_docx4j_wml.UnderlineEnumeration' && BOOLEAN_FORMS.has(text)) return 'none';
  return undefined;
}

/** A value read by Word's rule, as `onWordReading` receives it (CR-013 section 3.5, core-ts CR-015's D8). */
export interface WordReading {
  /** The element carrying the attribute, in the source document. */
  readonly element: Element;
  readonly attribute: Jsonix.XML.QName;
  /** The attribute's text, as written. */
  readonly text: string;
  /**
   * What the model holds instead: `true` for a boolean; `'none'` for `w:u`'s on/off words; an integer's
   * integer part (CR-015); undefined where the value was dropped.
   */
  readonly value: boolean | string | number | undefined;
  /**
   * For a boolean or an enumeration, `'value.kept.exact'` when the text, trimmed of XML whitespace,
   * is a form its type reads (`" false"`), else `'value.kept.invalid'` (`"yes"`, `"Single"`). For an
   * integer (CR-015), `'value.truncated'` (`"259.2"` read as 259, `"10206.0"` as 10206),
   * `'value.kept.invalid'` for a form with no digit (`""`, `.`, `+`, `-`, `+.`) read as 0, or
   * `'value.refused'` where the form is one Word refuses the whole document for and the value was dropped.
   */
  readonly kind: 'value.kept.exact' | 'value.kept.invalid' | 'value.truncated' | 'value.refused';
}

// ---------------------------------------------------------------------------------------------
// CR-016: integer attributes within Word's range
// ---------------------------------------------------------------------------------------------

/**
 * A value Word 365 would read differently from the tree (CR-016), as `onOutOfRange` receives it: an integer outside
 * the width Word stores its attribute in (`helpers/wml`'s `WORD_INTEGER_STORAGE`). The tree keeps the value, and a
 * save writes it as it is.
 */
export interface OutOfRange {
  /**
   * On read, the element in the source; on save, the element written, as it stands when the attribute is written: its
   * later attributes (`w:lineRule` after `w:line`) are not on it yet. `rule` and `storage` say what the element will.
   */
  readonly element: Element;
  readonly attribute: Jsonix.XML.QName;
  /** On save, from the part root, as `Normalisation`'s: `/w:document/w:body[1]/w:p[3]/w:pPr[1]/w:ind[1]/@w:left`. */
  readonly path?: string;
  /** The value the tree holds. */
  readonly value: number;
  /** What Word reads for it, and how Word stores the attribute. */
  readonly wordReads: WordIntegerReading;
  /**
   * The `WORD_INTEGER_STORAGE` entry the report comes from, and its key pair: the `TYPE_NAME` declaring the property
   * (`org_docx4j_wml.CTBorder` for `w:top`, `w:bdr` and every other border; `org_docx4j_wml.TblWidth` for `w:tblW` and
   * `w:tcW`), which the element's name alone cannot give, and the property, the attribute's local name.
   */
  readonly typeName: string;
  readonly property: string;
  readonly storage: WordIntegerStorage;
  /**
   * For an attribute whose sign carries a rule (`storage.signRule`: `w:spacing/@w:line`, `w:trHeight`, `w:framePr/@w:h`),
   * the rule the tree holds beside it (`lineRule` or `hRule`'s value), or undefined when absent (`auto`, or `atLeast`
   * for `hRule`). Undefined for every other attribute.
   */
  readonly rule?: string;
}

/** For each type, its properties in `WORD_INTEGER_STORAGE` (its own and its base types'), with the attribute each is. */
const rangedByType = new Map<string, readonly (readonly [typeName: string, property: string, attribute: Jsonix.XML.QName])[]>();

function rangedProperties(typeName: string): readonly (readonly [string, string, Jsonix.XML.QName])[] {
  let found = rangedByType.get(typeName);
  if (found === undefined) {
    const list: [string, string, Jsonix.XML.QName][] = [];
    for (let ti: Jsonix.TypeInfo | undefined = getContextSync().getTypeInfoByName(typeName); ti; ti = ti.baseTypeInfo) {
      const ranged = WORD_INTEGER_STORAGE[ti.name];
      if (ranged === undefined) continue;
      for (const property of (ti as unknown as { properties?: { name: string; attributeName?: Jsonix.XML.QName | null }[] }).properties ?? []) {
        if (property.attributeName && ranged[property.name] !== undefined) list.push([ti.name, property.name, property.attributeName]);
      }
    }
    rangedByType.set(typeName, found = list);
  }
  return found;
}

/**
 * The report for an owner's integer property, without its element and path, when Word would read it differently: its
 * reading with the rule beside it where its sign carries one, and the table entry it comes from. Undefined when the
 * value isn't a number, isn't ranged, or Word reads it as held.
 */
function outOfRangeOf(typeName: string, property: string, owner: unknown, value: unknown): Omit<OutOfRange, 'element' | 'attribute' | 'path'> | undefined {
  if (typeof value !== 'number') return undefined;
  const storage = WORD_INTEGER_STORAGE[typeName]?.[property];
  if (storage === undefined) return undefined;
  const held = storage.measured && storage.signRule ? (owner as Record<string, unknown>)[storage.signRule] : undefined;
  const rule = typeof held === 'string' ? held : undefined;
  const wordReads = wordIntegerReading(typeName, property, value, rule);
  if (wordReads === undefined || wordReads.kept) return undefined;
  return { value, wordReads, typeName, property, storage, ...(rule === undefined ? {} : { rule }) };
}

/** The read side (CR-016 section 4.2): each ranged property of a typed object Word would read differently. */
function reportOutOfRange(element: Element, value: object, onOutOfRange: (report: OutOfRange) => void): void {
  const typeName = (value as { TYPE_NAME?: string }).TYPE_NAME;
  if (typeof typeName !== 'string') return;
  for (const [declaring, property, attribute] of rangedProperties(typeName)) {
    const report = outOfRangeOf(declaring, property, value, (value as Record<string, unknown>)[property]);
    if (report) onOutOfRange({ ...report, element, attribute });
  }
}

// ---------------------------------------------------------------------------------------------
// CR-017: misplaced children
// ---------------------------------------------------------------------------------------------

/**
 * A child its container never admits (CR-017), as `onMisplaced` receives it: a class docx4j CR-064's `checkParent`
 * refuses there (a run in a cell, a table in a paragraph, a string in a run, ...), sdt content directly in sdt content,
 * or a name and value that disagree (a `w:p` whose value is a run; a block-level content control in a paragraph). The
 * tree keeps it, and a save writes it as held unless the callback throws.
 */
export interface MisplacedChild {
  /** The container's element: on read in the source; on save the element written, open, its children not yet written. */
  readonly element: Element;
  /** The container's `TYPE_NAME`. */
  readonly container: string;
  /** The child's element name, or `'#text'` for a string. */
  readonly child: { readonly namespaceURI: string; readonly localPart: string; readonly prefix?: string } | '#text';
  /** The child's class, where it has one: its value's `TYPE_NAME`, or the class its element name gives. */
  readonly childType?: string;
  /** Its position in the list holding it. */
  readonly index: number;
  /**
   * For a child of an `mc:AlternateContent` branch, tested by its host's set: the steps from the container to the
   * branch, `mc:AlternateContent[1]/mc:Choice[1]`.
   */
  readonly via?: string;
  /** `'misplaced'`, a class the container refuses; `'name-value'`, a value whose class isn't the one its name declares there. */
  readonly rule: MisplacedRule;
  /** docx4j's wording, with element names: "You can't add a run (w:r) directly to a table cell (w:tc); ...". */
  readonly message: string;
  /** On save, from the part root, in `OutOfRange`'s form: `/w:document/w:body[1]/w:r[1]`; the container's path for a string. */
  readonly path?: string;
  /** On read, the child's element in the source. */
  readonly childElement?: Element;
  /**
   * On read, for `'#text'`: the string a file held where its container admits none, which the runtime drops
   * (jsonix-CR-013's `onUnexpectedText`, 3.10.0); `index` is where it stood among the list's items.
   */
  readonly text?: string;
}

/** The context's declarations, for the name rule: the class a list (or, where it names none, the context) gives a name. */
let declarationsOf: { context: Jsonix.Context; declarations: Declarations } | undefined;
function declarations(): Declarations {
  const context = getContextSync();
  if (declarationsOf?.context === context) return declarationsOf.declarations;
  const cache = new Map<string, string | undefined>();
  const typed = (info: unknown): string | undefined => {
    const name = typeof info === 'object' && info !== null ? (info as { name?: unknown }).name : undefined;
    return typeof name === 'string' ? name : undefined;
  };
  const result: Declarations = {
    declared(containerType, list, name) {
      const key = `${containerType} ${list} {${name.namespaceURI}}${name.localPart}`;
      if (cache.has(key)) return cache.get(key);
      const property = (context.getTypeInfoByName(containerType) as unknown as { properties?: { name: string; elementTypeInfos?: { elementName: { namespaceURI: string; localPart: string }; typeInfo: unknown }[] }[] } | undefined)
        ?.properties?.find((p) => p.name === list);
      const local = property?.elementTypeInfos?.find((e) => e.elementName.namespaceURI === name.namespaceURI && e.elementName.localPart === name.localPart);
      const found = local ? typed(local.typeInfo)
        : typed(context.getElementInfo({ key: `{${name.namespaceURI}}${name.localPart}` } as unknown as string, undefined as unknown as string)?.typeInfo);
      cache.set(key, found);
      return found;
    },
    listDeclares(containerType, list, name) {
      const property = (context.getTypeInfoByName(containerType) as unknown as { properties?: { name: string; elementTypeInfos?: { elementName: { namespaceURI: string; localPart: string } }[] }[] } | undefined)
        ?.properties?.find((p) => p.name === list);
      return property?.elementTypeInfos?.some((e) => e.elementName.namespaceURI === name.namespaceURI && e.elementName.localPart === name.localPart) ?? false;
    },
    isA(type, base) {
      for (let ti: Jsonix.TypeInfo | undefined = context.getTypeInfoByName(type); ti; ti = ti.baseTypeInfo) if (ti.name === base) return true;
      return false;
    },
  };
  declarationsOf = { context, declarations: result };
  return result;
}

const W_NS_ = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
/** The level an ancestor fixes for a content control or custom XML below it. */
const LEVEL_OF: Readonly<Record<string, 'block' | 'run' | 'row' | 'cell'>> = {
  body: 'block', tc: 'block', txbxContent: 'block', hdr: 'block', ftr: 'block', footnote: 'block', endnote: 'block', comment: 'block',
  p: 'run', hyperlink: 'run', fldSimple: 'run', smartTag: 'run', ins: 'run', del: 'run', dir: 'run', bdo: 'run', tbl: 'row', tr: 'cell',
};
/** Containers whose class depends on the level they stand at, and that class per level. */
const BY_LEVEL: Readonly<Record<string, Readonly<Partial<Record<'block' | 'run' | 'row' | 'cell', string>>>>> = Object.fromEntries([
  ...['SdtContentBlock', 'CTSdtContentRun', 'CTSdtContentRow', 'CTSdtContentCell'].map((t) => [`org_docx4j_wml.${t}`,
    { block: 'org_docx4j_wml.SdtContentBlock', run: 'org_docx4j_wml.CTSdtContentRun', row: 'org_docx4j_wml.CTSdtContentRow', cell: 'org_docx4j_wml.CTSdtContentCell' }]),
  ...['CTCustomXmlBlock', 'CTCustomXmlRun'].map((t) => [`org_docx4j_wml.${t}`, { block: 'org_docx4j_wml.CTCustomXmlBlock', run: 'org_docx4j_wml.CTCustomXmlRun' }]),
]);

/**
 * The class a container is tested as. A kept `mc:Choice` or `mc:Fallback` types its children by their global
 * declarations, and the global `w:sdt` is the block-level class, so a run-level control in a paragraph's branch, and
 * everything below it, is typed block-level. So for sdt content and custom XML whose path up to the element that fixes
 * its level (a paragraph, a cell, a table, a row, ...) crosses a branch, the level is that element's, and its children
 * are tested leniently: one its list names is legal whatever class the runtime gave it. Elsewhere the runtime typed
 * it by the list it stands in, which is right. Undefined: a level with no container of that kind (custom XML in a row).
 */
function effectiveContainer(typed: string, element: Node): { type: string | undefined; lenient: boolean } {
  const byLevel = BY_LEVEL[typed];
  if (byLevel === undefined) return { type: typed, lenient: false };
  let crossed = false;
  let n: Node | null = element.parentNode;
  while (n !== null && n.nodeType === 1) {
    const e = n as Element;
    n = e.parentNode;
    if (e.namespaceURI === MC_NS && (e.localName === 'Choice' || e.localName === 'Fallback' || e.localName === 'AlternateContent')) {
      crossed ||= e.localName !== 'AlternateContent';
      continue;
    }
    if (e.namespaceURI !== W_NS_) break;
    if (e.localName === 'sdt' || e.localName === 'sdtContent' || e.localName === 'customXml') continue;
    const level = LEVEL_OF[e.localName];
    if (level === undefined || !crossed) break;
    return { type: byLevel[level], lenient: true };
  }
  return { type: typed, lenient: false };
}

/**
 * Tests a container's list, and each `mc:AlternateContent` branch in it by the container's set, calling `found` with
 * each misplaced child, its index in its own list, and the steps to it from the container (each `name[n]`, from
 * `stepName`; `offset` adds the same-named property elements written before the list).
 */
function testList(containerType: string, items: readonly unknown[], stepName: (item: unknown) => string | undefined, offset: (step: string) => number,
  mc: string, found: (misplaced: Misplacement, item: unknown, index: number, via: readonly string[], step: string | undefined) => void, via: readonly string[] = [],
  lenient = false): void {
  const context = declarations();
  const seen = new Map<string, number>();
  items.forEach((item, index) => {
    const name = stepName(item);
    let step: string | undefined;
    if (name !== undefined) {
      const n = (seen.get(name) ?? 0) + 1;
      seen.set(name, n);
      step = `${name}[${n + offset(name)}]`;
    }
    const misplaced = misplacement(containerType, item, context, lenient);
    if (misplaced) found(misplaced, item, index, via, step);
    const alternate = alternateContentOf(item);
    if (alternate && step !== undefined) {
      // A branch's children are typed by their global declarations, not by the host's list: tested leniently.
      (alternate.choice ?? []).forEach((choice, k) => testList(containerType, branchItems(choice), stepName, () => 0, mc, found, [...via, step!, `${mc}:Choice[${k + 1}]`], true));
      if (alternate.fallback) testList(containerType, branchItems(alternate.fallback), stepName, () => 0, mc, found, [...via, step, `${mc}:Fallback[1]`], true);
    }
  });
}

/** A list item's element name, as a report gives it. */
function childName(item: unknown): MisplacedChild['child'] {
  if (typeof item !== 'object' || item === null) return '#text';
  if (isNode(item)) {
    const e = item as Element;
    return { namespaceURI: e.namespaceURI ?? '', localPart: e.localName, ...(e.prefix ? { prefix: e.prefix } : {}) };
  }
  const name = (item as { name?: MisplacedChild['child'] }).name;
  return typeof name === 'object' && name !== null ? name : '#text';
}

/**
 * The strings the runtime dropped from a container (jsonix-CR-013), with its list's items: each reported as `'#text'`
 * where the container's set refuses a string, at the index it stood at among the items, and an `mc:AlternateContent`
 * branch's by its host's set.
 */
function reportDroppedText(container: string, element: Element, items: readonly unknown[], via: readonly string[], sources: WeakMap<object, Element>,
  dropped: WeakMap<Node, Text[]>, mc: string, onMisplaced: (report: MisplacedChild) => void): void {
  const sourceOf = (item: unknown): Node | undefined => {
    if (isNode(item)) return item as Node;
    const v = typeof item === 'object' && item !== null ? ((item as { value?: unknown }).value ?? item) : undefined;
    return typeof v === 'object' && v !== null ? sources.get(v) : undefined;
  };
  const texts = dropped.get(element);
  if (texts) {
    const misplaced = misplacement(container, '');
    if (misplaced) {
      const itemNodes = new Set(items.map(sourceOf).filter((n): n is Node => n !== undefined));
      for (const text of texts) {
        let index = 0;
        for (let sibling = text.previousSibling; sibling; sibling = sibling.previousSibling) if (itemNodes.has(sibling)) index++;
        onMisplaced({
          element, container, child: '#text', index, ...(via.length ? { via: via.join('/') } : {}), rule: misplaced.rule, message: misplaced.message,
          text: text.data,
        });
      }
    }
  }
  items.forEach((item) => {
    const alternate = alternateContentOf(item);
    const acElement = alternate ? sourceOf(item) : undefined;
    if (!alternate || !acElement) return;
    const acStep = `${mc}:AlternateContent[${items.filter((other, i) => alternateContentOf(other) && i <= items.indexOf(item)).length}]`;
    (alternate.choice ?? []).forEach((choice, k) => {
      const choiceElement = sources.get(choice as object);
      if (choiceElement) reportDroppedText(container, choiceElement as Element, branchItems(choice), [...via, acStep, `${mc}:Choice[${k + 1}]`], sources, dropped, mc, (report) => onMisplaced({ ...report, element }));
    });
    const fallbackElement = alternate.fallback ? sources.get(alternate.fallback as object) : undefined;
    if (fallbackElement) reportDroppedText(container, fallbackElement as Element, branchItems(alternate.fallback), [...via, acStep, `${mc}:Fallback[1]`], sources, dropped, mc, (report) => onMisplaced({ ...report, element }));
  });
}

/** The read side (CR-017 section 4.2): a typed container's list, once its value is complete. */
function reportMisplacedOnRead(element: Element, value: object, sources: WeakMap<object, Element>, dropped: WeakMap<Node, Text[]>, onMisplaced: (report: MisplacedChild) => void): void {
  const typed = (value as { TYPE_NAME?: string }).TYPE_NAME;
  const spec = typed === undefined ? undefined : CONTAINERS[typed];
  if (spec === undefined) return;
  const { type: container, lenient } = effectiveContainer(typed!, element);
  if (container === undefined) return;
  // The runtime leaves the list unset when it read no element: a string that was the container's only child is still
  // to be reported, so an unset list is an empty one here.
  const listed = (value as Record<string, unknown>)[spec.list];
  const items: readonly unknown[] = Array.isArray(listed) ? listed : [];
  const table = getContextSync().namespacePrefixes;
  const stepName = (item: unknown): string | undefined => {
    const name = childName(item);
    return name === '#text' ? undefined : isNode(item) ? (item as Element).nodeName : qualifiedName(name as Jsonix.XML.QName, table);
  };
  testList(container, items, stepName, () => 0, table[MC_NS] || 'mc',
    (misplaced, item, index, via) => {
      const childValue = typeof item === 'object' && item !== null ? ((item as { value?: unknown }).value ?? item) : undefined;
      const childElement = isNode(item) ? item as Element : typeof childValue === 'object' && childValue !== null ? sources.get(childValue) : undefined;
      onMisplaced({
        element, container: container!, child: childName(item), ...(misplaced.childType ? { childType: misplaced.childType } : {}), index,
        ...(via.length ? { via: via.join('/') } : {}), rule: misplaced.rule, message: misplaced.message, ...(childElement ? { childElement } : {}),
      });
    }, [], lenient);
  reportDroppedText(container, element, items, [], sources, dropped, table[MC_NS] || 'mc', onMisplaced);
}

/** The facade's unmarshal options: the runtime's, for this call only (CR-007), and `onWordReading` (CR-013). */
export interface UnmarshalOptions extends Jsonix.UnmarshallerOptions {
  /**
   * Called for each attribute value `wordInvalidValue` read by Word's rule rather than refused,
   * after it was read. Not called when the caller passes its own `onInvalidValue`.
   */
  onWordReading?(reading: WordReading): void;
  /**
   * Called for each integer attribute value Word 365 would read differently (CR-016): outside the width Word stores
   * it in. The tree keeps the value as written. Not called unless passed.
   */
  onOutOfRange?(report: OutOfRange): void;
  /**
   * Called for each misplaced child (CR-017): a class its container never admits, or a name and value that disagree.
   * Each container's list is tested once its value is complete. The tree keeps what it read. Not called unless passed.
   */
  onMisplaced?(report: MisplacedChild): void;
}

function readOnceTrimmed(text: string, typeInfo: Jsonix.TypeInfo): boolean {
  const trimmed = text.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '');
  return typeInfo.name === 'Boolean' ? BOOLEAN_FORMS.has(trimmed) : Object.values((typeInfo as Jsonix.EnumLeafInfo).keys).includes(trimmed);
}

function reportingWordReadings(onWordReading: (reading: WordReading) => void): NonNullable<Jsonix.UnmarshallerOptions['onInvalidValue']> {
  return (text, propertyInfo, classInfo, node, error) => {
    const value = wordInvalidValue(text, propertyInfo, classInfo, node, error) as WordReading['value'];
    const kind: WordReading['kind'] = INTEGER_TYPES.has(propertyInfo.typeInfo.name)
      ? (value === undefined ? 'value.refused' : NO_DIGITS.test(text) ? 'value.kept.invalid' : 'value.truncated')
      : (readOnceTrimmed(text, propertyInfo.typeInfo) ? 'value.kept.exact' : 'value.kept.invalid');
    onWordReading({ element: node, attribute: propertyInfo.attributeName, text, value, kind });
    return value;
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
function marshalToDocument(context: Jsonix.Context, element: Jsonix.TypedNamedValue, facadeOptions?: MarshalOptions): Document {
  const { normaliseOnSave, onNormalised, onOutOfRange, onMisplaced, ...options } = facadeOptions ?? {};
  const table = namespacePrefixesFor(context, element, options.namespacePrefixes);
  const choices = new Map<Element, ReadonlyMap<string, string>>();
  const callerOnElement = options.onElement;
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
  const marshallerOptions: Jsonix.MarshallerOptions = { ...options, namespacePrefixes: table, onElement };
  const normalised = (normaliseOnSave ?? normaliseOnSaveByDefault) ? normalising(marshallerOptions, table, onNormalised) : marshallerOptions;
  const ranged = onOutOfRange ? reportingOutOfRange(normalised, table, onOutOfRange) : normalised;
  const doc = context.createMarshaller(onMisplaced ? reportingMisplaced(ranged, table, onMisplaced) : ranged)
    .marshalDocument(element);
  if (element.name.namespaceURI !== MC_NS && bindsIgnorable(context, element)) declareExtensionsIgnorable(doc.documentElement, table);
  const rootValue = rootValueOf(element);
  fixRootNamespaceDeclarations(doc.documentElement, table, typeof rootValue === 'object' && rootValue !== null ? sourceBindings.get(rootValue) : undefined, choices);
  return doc;
}

// ---------------------------------------------------------------------------------------------
// CR-013 phase C: what a save writes, run properties in their normal form (docx4j CR-060 with CR-061)
// ---------------------------------------------------------------------------------------------

let normaliseOnSaveByDefault = true;

/**
 * Whether the facade's marshal functions write run properties in their normal form (CR-013 section
 * 3.6), for every call that does not say: on unless set off, as docx4j's `docx4j.save.normalise`.
 * A call's `normaliseOnSave` takes precedence.
 */
export function setNormaliseOnSave(on: boolean): void {
  normaliseOnSaveByDefault = on;
}

/** A value the normaliser rewrote or left out, as `onNormalised` receives it (CR-013 section 3.6, core-ts CR-015's D8). */
export interface Normalisation {
  /** The element written with the rewritten attribute, or the one a left-out element was left out of. */
  readonly element: Element;
  /** From the part root: `/w:document/w:body[1]/w:p[3]/w:r[1]/w:rPr[1]/w:color[1]/@w:val`, or `.../w:rPr[1]/w:highlight` for an element left out. */
  readonly path: string;
  /** As read: the attribute's text, or the left-out element's `w:val` (undefined where it had none). */
  readonly before: string | undefined;
  /** As written: the attribute's new text; null where the attribute or the element was left out. */
  readonly after: string | null;
}

/** The facade's marshal options: the runtime's, for this call only (CR-007), and the normaliser's (CR-013). */
export interface MarshalOptions extends Jsonix.MarshallerOptions {
  /** Whether this call writes run properties in their normal form; where unset, `setNormaliseOnSave`'s setting, on by default. */
  normaliseOnSave?: boolean;
  /** Called for each value the normaliser rewrote or left out. */
  onNormalised?(normalisation: Normalisation): void;
  /**
   * Called for each integer attribute value written that Word 365 would read differently (CR-016); the value is
   * written as held, whether or not `normaliseOnSave` is on. Not called unless passed.
   */
  onOutOfRange?(report: OutOfRange): void;
  /**
   * Called for each misplaced child (CR-017), with its path. Each container's list is tested as the container is
   * opened, before its children are written, so a string in a run is reported before the runtime refuses it, and a
   * callback that throws stops the save there with its own error. Otherwise written as held. Not called unless passed.
   */
  onMisplaced?(report: MisplacedChild): void;
}

/**
 * The colour attributes written in their normal form, by the class declaring them and the property,
 * with what a value Word rejects becomes: `'auto'`, which Word reads as automatic in a fill, a
 * pattern, a border or a page colour; or null, leaving out an underline colour, which Word reads as
 * absent. A run's `w:color` is the third case: left out whole (`leftOutOnSave`), and written `auto`
 * only beside a theme attribute, since its `w:val` is required.
 */
const COLOUR_ATTRIBUTES: ReadonlyMap<string, 'auto' | null> = new Map<string, 'auto' | null>([
  ['org_docx4j_wml.Color val', 'auto'],
  ['org_docx4j_wml.U color', null],
  ['org_docx4j_wml.CTShd fill', 'auto'],
  ['org_docx4j_wml.CTShd color', 'auto'],
  ['org_docx4j_wml.CTBorder color', 'auto'],
  ['org_docx4j_wml.CTBackground color', 'auto'],
]);

/** `wml`'s `ST_LongHexNumber` revision-session attributes: an empty one is schema-invalid, and Word's save leaves it out (CR-015 N7). */
const RSID_ATTRIBUTES: ReadonlySet<string> = new Set(['rsidDel', 'rsidP', 'rsidR', 'rsidRDefault', 'rsidRPr', 'rsidSect', 'rsidTr']);

/** ST_HighlightColor: the sixteen names and `none`. */
const HIGHLIGHT_VALUES: ReadonlySet<string> = new Set([...HIGHLIGHT_COLORS.map(([name]) => name), 'none']);

/** A colour's normal form (docx4j's NormaliseOnSave.plain): exactly `auto` as it is; a value Word accepts as upper-case RRGGBB; else `rejected`. */
function normalColour(text: string, rejected: 'auto' | null): string | null {
  if (text === 'auto') return text;
  const hex = wordColor(text);
  return hex === undefined ? rejected : hex.substring(1);
}

/**
 * The elements a save leaves out, by type: a `w:color` Word reads as absent (its value rejected or
 * missing, and no theme attribute beside it), and a `w:highlight`, `w:vertAlign` or `w:em` that
 * Word's repair drops (a highlight outside ST_HighlightColor; a vertAlign or em with no value, the
 * parse having dropped an invalid one).
 */
function leftOutOnSave(value: unknown, typeInfo: Jsonix.TypeInfo): boolean {
  if (typeof value !== 'object' || value === null) return false;
  const v = value as { val?: string; themeColor?: string; themeTint?: string; themeShade?: string };
  switch (typeInfo.name) {
    case 'org_docx4j_wml.Color': return wordColor(v.val) === undefined && v.themeColor === undefined && v.themeTint === undefined && v.themeShade === undefined;
    case 'org_docx4j_wml.Highlight': return v.val === undefined || !HIGHLIGHT_VALUES.has(v.val);
    case 'org_docx4j_wml.CTVerticalAlignRun':
    case 'org_docx4j_wml.CTEm': return v.val === undefined;
    default: return false;
  }
}

/** An element's path from the part root, each step after the root with its position among same-named siblings. */
function pathOf(node: Node): string {
  const steps: string[] = [];
  for (let n: Node | null = node; n && n.nodeType === 1; n = n.parentNode) {
    if (n.parentNode?.nodeType !== 1) { steps.unshift(n.nodeName); continue; }
    let position = 1;
    for (let sibling = n.previousSibling; sibling; sibling = sibling.previousSibling) if (sibling.nodeName === n.nodeName) position++;
    steps.unshift(`${n.nodeName}[${position}]`);
  }
  return '/' + steps.join('/');
}

function qualifiedName(name: Jsonix.XML.QName, table: Readonly<Record<string, string>>): string {
  const prefix = name.namespaceURI ? (table[name.namespaceURI] ?? name.prefix) : '';
  return prefix ? `${prefix}:${name.localPart}` : name.localPart;
}

/** Whether `owner`'s property holds `value`, itself or in a `{ name, value }` pair, alone or in a list. */
function holds(owner: unknown, property: string, value: unknown): boolean {
  if (typeof owner !== 'object' || owner === null) return false;
  const held = (owner as Record<string, unknown>)[property];
  return [held].flat().some((item) => item === value || (typeof item === 'object' && item !== null && (item as { value?: unknown }).value === value));
}

/**
 * The marshaller options with CR-013's normaliser (section 3.6), on jsonix-CR-011's marshal-side
 * adapters, keyed by type rather than element name and scoped to `org_docx4j_wml`. The caller's own
 * `onAttributeValue` and `omitElement` are asked first; where they fall through (undefined), the
 * normaliser applies. Marshal-only: the tree keeps what it read. Where `onNormalised` is passed,
 * `onElement` also tracks what has been written, so that each record can say where it is.
 */
function normalising(options: Jsonix.MarshallerOptions, table: Readonly<Record<string, string>>, onNormalised?: (normalisation: Normalisation) => void): Jsonix.MarshallerOptions {
  const { onAttributeValue: callerOnAttributeValue, omitElement: callerOmitElement, onElement } = options;
  const writtenFrom = new WeakMap<Node, unknown>();
  let lastWritten: Node | undefined;
  return {
    ...options,
    onElement: !onNormalised ? onElement : (element: Node, value: unknown): void => {
      writtenFrom.set(element, value);
      lastWritten = element;
      onElement?.(element, value);
    },
    onAttributeValue(value, text, owner, propertyInfo, classInfo) {
      const callers = callerOnAttributeValue?.(value, text, owner, propertyInfo, classInfo);
      if (callers !== undefined || classInfo.module.name !== 'org_docx4j_wml') return callers;
      let normal: string | null;
      if (text === '' && propertyInfo.attributeName.namespaceURI === WML_NS && RSID_ATTRIBUTES.has(propertyInfo.attributeName.localPart)) {
        normal = null;
      } else {
        const rejected = COLOUR_ATTRIBUTES.get(`${classInfo.name} ${propertyInfo.name}`);
        if (rejected === undefined) return undefined;
        normal = normalColour(text, rejected);
        if (normal === text) return undefined;
      }
      if (onNormalised) {
        // The owner's element is open, and every element written since is inside it: the runtime writes an
        // element's attributes after its content (a w:p's rsids after its runs).
        let element: Node | null | undefined = lastWritten;
        while (element && writtenFrom.get(element) !== owner) element = element.parentNode;
        element ??= lastWritten;
        if (element) onNormalised({ element: element as Element, path: `${pathOf(element)}/@${qualifiedName(propertyInfo.attributeName, table)}`, before: text, after: normal });
      }
      return normal;
    },
    omitElement(value, name, typeInfo, propertyInfo, classInfo) {
      const callers = callerOmitElement?.(value, name, typeInfo, propertyInfo, classInfo);
      if (callers !== undefined || typeInfo.module?.name !== 'org_docx4j_wml' || !leftOutOnSave(value, typeInfo)) return callers;
      if (onNormalised) {
        // The owner's element is open, and every element written since is inside it.
        let parent: Node | null | undefined = lastWritten;
        while (parent && !holds(writtenFrom.get(parent), propertyInfo.name, value)) parent = parent.parentNode;
        parent ??= lastWritten;
        if (parent) onNormalised({ element: parent as Element, path: `${pathOf(parent)}/${qualifiedName(name, table)}`, before: (value as { val?: string }).val, after: null });
      }
      return true;
    },
  };
}

/**
 * The marshaller options reporting each integer attribute Word would read differently (CR-016 section 4.3), on
 * jsonix-CR-011's `onAttributeValue`, composed around the options it is given (the normaliser's, or the caller's).
 * It never changes what is written: it returns what the options it wraps return.
 */
function reportingOutOfRange(options: Jsonix.MarshallerOptions, table: Readonly<Record<string, string>>, onOutOfRange: (report: OutOfRange) => void): Jsonix.MarshallerOptions {
  const { onAttributeValue: innerOnAttributeValue, onElement } = options;
  const writtenFrom = new WeakMap<Node, unknown>();
  let lastWritten: Node | undefined;
  return {
    ...options,
    onElement(element: Node, value: unknown): void {
      writtenFrom.set(element, value);
      lastWritten = element;
      onElement?.(element, value);
    },
    onAttributeValue(value, text, owner, propertyInfo, classInfo) {
      const result = innerOnAttributeValue?.(value, text, owner, propertyInfo, classInfo);
      // classInfo is the class declaring the property (the runtime marshals a base type's properties through the base).
      const report = outOfRangeOf(classInfo.name, propertyInfo.name, owner, value);
      if (report) {
        // The owner's element is open, and every element written since is inside it.
        let element: Node | null | undefined = lastWritten;
        while (element && writtenFrom.get(element) !== owner) element = element.parentNode;
        element ??= lastWritten;
        if (element) {
          onOutOfRange({
            ...report, element: element as Element, attribute: propertyInfo.attributeName,
            path: `${pathOf(element)}/@${qualifiedName(propertyInfo.attributeName, table)}`,
          });
        }
      }
      return result;
    },
  };
}

/**
 * The number of the container's own property elements named `step` that the marshal writes before its list (a
 * `w:pPr` before a paragraph's content), so a misplaced one's step counts them, as the written document will.
 */
function writtenBefore(container: object, list: string, step: string, table: Readonly<Record<string, string>>): number {
  const chain: Jsonix.TypeInfo[] = [];
  for (let ti: Jsonix.TypeInfo | undefined = getContextSync().getTypeInfoByName((container as { TYPE_NAME: string }).TYPE_NAME); ti; ti = ti.baseTypeInfo) chain.unshift(ti);
  let count = 0;
  for (const ti of chain) {
    for (const property of (ti as unknown as { properties?: { name: string; elementName?: Jsonix.XML.QName }[] }).properties ?? []) {
      if (property.name === list) return count;
      const held = (container as Record<string, unknown>)[property.name];
      if (property.elementName && held !== undefined && held !== null && qualifiedName(property.elementName, table) === step) count += Array.isArray(held) ? held.length : 1;
    }
  }
  return count;
}

/**
 * The marshaller options reporting each misplaced child (CR-017 section 4.2), on the marshal's `onElement`, which the
 * runtime calls as it opens an element, before its children: each container's list is tested then. Composed around the
 * options it is given; it never changes what is written.
 */
function reportingMisplaced(options: Jsonix.MarshallerOptions, table: Readonly<Record<string, string>>, onMisplaced: (report: MisplacedChild) => void): Jsonix.MarshallerOptions {
  const { onElement } = options;
  return {
    ...options,
    onElement(written: Node, value: unknown): void {
      const typed = typeof value === 'object' && value !== null && !isNode(value) ? (value as { TYPE_NAME?: string }).TYPE_NAME : undefined;
      const spec = typed === undefined ? undefined : CONTAINERS[typed];
      const items = spec ? (value as Record<string, unknown>)[spec.list] : undefined;
      const { type: container, lenient } = spec ? effectiveContainer(typed!, written) : { type: undefined, lenient: false };
      if (spec && container !== undefined && Array.isArray(items)) {
        const base = pathOf(written);
        const stepName = (item: unknown): string | undefined => {
          const name = childName(item);
          return name === '#text' ? undefined : isNode(item) ? (item as Element).nodeName : qualifiedName(name as Jsonix.XML.QName, table);
        };
        testList(container!, items, stepName, (step) => writtenBefore(value as object, spec.list, step, table), table[MC_NS] || 'mc',
          (misplaced, item, index, via, step) => onMisplaced({
            element: written as Element, container: container!, child: childName(item), ...(misplaced.childType ? { childType: misplaced.childType } : {}), index,
            ...(via.length ? { via: via.join('/') } : {}), rule: misplaced.rule, message: misplaced.message,
            path: [base, ...via, ...(step ? [step] : [])].join('/'),
          }), [], lenient);
      }
      onElement?.(written, value);
    },
  };
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
 * reads all three places alike (core-ts, 2026-10-20). 0.4.3 declared it on the root. A Choice's own
 * mc:Ignorable and mc:MustUnderstand prefixes follow its Requires, as docx4j's do (its CR-052, D1 (b)).
 *
 * A Requires prefix counts because it names a namespace, not because anything is in it: Word's
 * chartex Choice says `Requires="cx1"` (2015/9/8) over content in `cx` (2014), so nothing in the
 * branch uses cx1 and its declaration went, leaving a Requires no reader can resolve (core-ts,
 * once its CR-010 phase 4 kept branches by default, 2026-10-19; docx4j CR-049 declares the same).
 * A declaration below the root that binds the prefix answers for its Choice, as everywhere here.
 * A Requires prefix nothing resolves is kept, with a warning: dropping it would change which
 * branch a reader takes. So is a Choice's mc:MustUnderstand prefix, which a reader that doesn't
 * understand it must refuse; one in the Choice's mc:Ignorable is dropped from it, as from the root's.
 *
 * The source comes first because its binding is what the prefix meant: three Excel VML parts of
 * the fidelity corpus bind `x` to the Excel VML namespace, which the aliases would declare as
 * SpreadsheetML. A declaration the tree's own elements or attributes use is never rebound: where
 * the source meant another namespace by a prefix the root's tree uses, for mc:Ignorable, or the
 * Choice's own content uses, for a prefix the Choice names, that is warned about. A use elsewhere
 * in the tree doesn't stop a Choice declaring its own binding, so a prefix bound two ways keeps
 * both meanings.
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
  // The prefixes a Choice names that go on the root: by Requires or mc:MustUnderstand (kept if
  // nothing resolves them), and by the Choice's own mc:Ignorable (dropped from it if nothing does).
  const required = new Map<string, ChoiceAttribute>();
  const choiceIgnorable: { choice: Element; prefix: string }[] = [];
  // The prefixes a Choice names that only the source bound, each declared on its Choice (CR-012).
  const local: { choice: Element; prefix: string; namespaceURI: string }[] = [];
  const clashed = new Set<string>();
  const clash = (prefix: string, meant: string, tree: string): void => {
    if (clashed.has(prefix)) return;
    clashed.add(prefix);
    console.warn(`marshal: the source bound ${prefix} to ${meant}, but the tree uses it for ${tree}; an mc:Ignorable, Requires or mc:MustUnderstand naming it now means the latter`);
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
      for (const [prefix, attribute] of namedOnChoice(element)) {
        if (scope.has(prefix)) continue;
        // What the prefix meant where the Choice named it: the Choice's own source, else the root's.
        const namespaceURI = choices?.get(element)?.get(prefix) ?? rootSource?.get(prefix);
        if (namespaceURI !== undefined) {
          const tree = otherUse(element, prefix, namespaceURI);
          if (tree === undefined) {
            local.push({ choice: element, prefix, namespaceURI });
            continue;
          }
          clash(prefix, namespaceURI, tree);
        }
        if (attribute === 'mc:Ignorable') choiceIgnorable.push({ choice: element, prefix });
        else if (!required.has(prefix)) required.set(prefix, attribute);
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
  for (const prefix of required.keys()) used.add(prefix);
  for (const { prefix } of choiceIgnorable) used.add(prefix);
  for (const [prefix, attr] of declared) {
    // A declaration kept only because mc:Ignorable names its prefix gives way to the source's
    // binding, which is what that prefix meant (CR-011). One a Requires named has gone to its Choice.
    const meant = ignorable.includes(prefix) ? rootSource?.get(prefix) : undefined;
    const superseded = !usedByTree.has(prefix) && meant !== undefined && meant !== attr.value;
    if (!used.has(prefix) || attr.value === XML_NS || superseded) root.removeAttributeNode(attr);
  }
  if (ignorable.length > 0 || required.size > 0 || choiceIgnorable.length > 0) declareOnRoot(root, table, ignorable, required, choiceIgnorable, rootSource, clash);
  // After the root's, so that a binding the root declares isn't repeated on a Choice.
  for (const { choice, prefix, namespaceURI } of local) {
    if (root.getAttributeNode(`xmlns:${prefix}`)?.value === namespaceURI || choice.getAttributeNode(`xmlns:${prefix}`)) continue;
    // Ahead of Requires and the Choice's other attributes, as Word writes it.
    const after = Array.from(choice.attributes).filter((attr) => prefixOfDeclaration(attr) === undefined);
    for (const attr of after) choice.removeAttributeNode(attr);
    choice.setAttributeNS(XMLNS_NS, `xmlns:${prefix}`, namespaceURI);
    for (const attr of after) choice.setAttributeNode(attr);
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

/** `fixRootNamespaceDeclarations`'s declarations on the root: the mc:Ignorable prefixes, and the prefixes a Choice names that no source bound. */
function declareOnRoot(root: Element, table: Record<string, string>, ignorable: string[], required: ReadonlyMap<string, ChoiceAttribute>,
  choiceIgnorable: readonly { choice: Element; prefix: string }[], rootSource: ReadonlyMap<string, string> | undefined,
  clash: (prefix: string, meant: string, tree: string) => void): void {
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
  const unrequired = [...required].filter(([prefix]) => !declare(prefix));
  for (const attribute of ['Requires', 'mc:MustUnderstand']) {
    const prefixes = unrequired.filter(([, by]) => by === attribute).map(([prefix]) => prefix);
    if (prefixes.length > 0) console.warn(`marshal: an mc:Choice's ${attribute} names ${prefixes.join(', ')}, which no declaration or NAMESPACE_PREFIXES entry resolves; kept`);
  }
  // A Choice's own mc:Ignorable naming a prefix nothing declares is dropped from it, as the root's is.
  const dropped = new Map<Element, Set<string>>();
  for (const { choice, prefix } of choiceIgnorable) {
    if (declare(prefix)) continue;
    let prefixes = dropped.get(choice);
    if (!prefixes) dropped.set(choice, prefixes = new Set());
    prefixes.add(prefix);
  }
  if (dropped.size > 0) {
    const all = new Set([...dropped.values()].flatMap((prefixes) => [...prefixes]));
    console.warn(`marshal: an mc:Choice's mc:Ignorable names ${[...all].join(', ')}, which no declaration or NAMESPACE_PREFIXES entry resolves; dropped`);
    for (const [choice, prefixes] of dropped) {
      const attr = choice.getAttributeNodeNS(MC_NS, 'Ignorable')!;
      const kept = attr.value.split(/\s+/).filter((prefix) => prefix && !prefixes.has(prefix));
      if (kept.length > 0) choice.setAttributeNS(MC_NS, attr.name, kept.join(' '));
      else choice.removeAttributeNode(attr);
    }
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
