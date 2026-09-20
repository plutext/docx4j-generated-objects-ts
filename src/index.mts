// The facade of @docx4j/generated-objects-ts: docx4j's names over the Jsonix runtime and the generated
// Office Open XML mappings in ../modules. Hand-written, MODULE_NAMES included: generate.sh does not write
// it, so add or remove names when modules appear or disappear (the smoke test fails on a stale list).
import { Jsonix } from '@docx4j/jsonix';
export { Jsonix };
export type { TypedNamedValue, XmlQName, XmlCalendar, XmlDuration, JsonixMapping } from '../modules/org_docx4j_wml.mjs';
import type { PackageElement, Package, Part } from '../modules/org_docx4j_xmlPackage.mjs';
export type { PackageElement, Package, Part };

/** All generated modules, one Jsonix mapping each; they reference each other, so a context needs them all. */
export const MODULE_NAMES = [
  'org_docx4j_bibliography',
  'org_docx4j_cei',
  'org_docx4j_com_microsoft_schemas_ink_x2010_main',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2010_chartDrawing',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2010_diagram',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2010_main',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2012_chart',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2012_chartStyle',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2012_main',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2014_main',
  'org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram',
  'org_docx4j_com_microsoft_schemas_office_drawing_x201611_main',
  'org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink',
  'org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2017_model3d',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor',
  'org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction',
  'org_docx4j_com_microsoft_schemas_office_thememl_x2012_main',
  'org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11',
  'org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11',
  'org_docx4j_com_microsoft_schemas_office_word_x2006_wordml',
  'org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas',
  'org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing',
  'org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingGroup',
  'org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape',
  'org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing',
  'org_docx4j_com_microsoft_schemas_office_x2006_encryption',
  'org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate',
  'org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password',
  'org_docx4j_customXmlProperties',
  'org_docx4j_customxml',
  'org_docx4j_dml',
  'org_docx4j_dml_chart',
  'org_docx4j_dml_chartDrawing',
  'org_docx4j_dml_chart_x2007',
  'org_docx4j_dml_compatibility',
  'org_docx4j_dml_diagram',
  'org_docx4j_dml_diagram2008',
  'org_docx4j_dml_lockedCanvas',
  'org_docx4j_dml_picture',
  'org_docx4j_dml_spreadsheetdrawing',
  'org_docx4j_dml_wordprocessingDrawing',
  'org_docx4j_docProps_core',
  'org_docx4j_docProps_core_dc_elements',
  'org_docx4j_docProps_core_dc_terms',
  'org_docx4j_docProps_coverPageProps',
  'org_docx4j_docProps_custom',
  'org_docx4j_docProps_extended',
  'org_docx4j_docProps_variantTypes',
  'org_docx4j_math',
  'org_docx4j_mce',
  'org_docx4j_org_w3_x1998_math_mathML',
  'org_docx4j_org_w3_x2003_inkML',
  'org_docx4j_relationships',
  'org_docx4j_sharedtypes',
  'org_docx4j_vml',
  'org_docx4j_vml_officedrawing',
  'org_docx4j_vml_presentationDrawing',
  'org_docx4j_vml_root',
  'org_docx4j_vml_spreadsheetDrawing',
  'org_docx4j_vml_wordprocessingDrawing',
  'org_docx4j_w14',
  'org_docx4j_w15',
  'org_docx4j_w15symex',
  'org_docx4j_w16',
  'org_docx4j_w16cex',
  'org_docx4j_w16cid',
  'org_docx4j_wml',
  'org_docx4j_xmlPackage',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x2010_main',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_sectionzoom',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main',
  'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main',
  'org_pptx4j_pml',
  'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main',
  'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac',
  'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_main',
  'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2011_x1_ac',
  'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision',
  'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_x11_main',
  'org_xlsx4j_schemas_microsoft_com_office_excel_2006_main',
  'org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main',
  'org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing',
  'org_xlsx4j_sml',
] as const;

export type ModuleName = (typeof MODULE_NAMES)[number];

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
  'http://opendope.org/SmartArt/DataHierarchy': 'odgm',
});

let contextPromise: Promise<Jsonix.Context> | undefined;
let builtContext: Jsonix.Context | undefined;

/**
 * The context over all Office Open XML mappings, built once on first use (the modules are
 * loaded lazily; 103 modules are several megabytes). Parent pointers are on by default, as in
 * docx4j's object model, and so is docx4j's prefix table (`NAMESPACE_PREFIXES`); pass options to
 * change that or anything else.
 */
export function getContext(options?: Jsonix.ContextOptions): Promise<Jsonix.Context> {
  if (contextPromise === undefined) {
    contextPromise = Promise.all(MODULE_NAMES.map((name) => import(`../modules/${name}.mjs`))).then(
      (modules) => {
        builtContext = new Jsonix.Context(modules.flatMap((m) => Object.values(m) as Jsonix.Mapping[]), { parentPointers: true, namespacePrefixes: NAMESPACE_PREFIXES, ...options });
        return builtContext;
      },
    );
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

/** docx4j XmlUtils.unmarshalString: parse XML into a typed element, e.g. unmarshalString<DocumentElement>(xml). */
export async function unmarshalString<E extends Jsonix.TypedNamedValue = Jsonix.TypedNamedValue>(xml: string): Promise<E> {
  return (await getContext()).createUnmarshaller().unmarshalString<E>(xml);
}

/** docx4j XmlUtils.marshaltoString: serialise an element to XML, with docx4j's prefixes and namespace declarations. */
export async function marshalString(element: Jsonix.TypedNamedValue): Promise<string> {
  return serialize(marshalToDocument(await getContext(), element));
}

/** docx4j XmlUtils.unwrap: the value of an element. */
export function unwrap<T>(element: Jsonix.TypedNamedValue<T>): T {
  return element.value;
}

/** docx4j XmlUtils.deepCopy: a structural copy with PARENT re-linked (Jsonix.Util.deepCopy). */
export function deepCopy<T>(value: T, parent?: unknown): T {
  return Jsonix.Util.deepCopy(value, parent);
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

/** docx4j XmlUtils.unmarshal(Node): a DOM element or document, through the shared context. */
export async function unmarshalNode<E extends Jsonix.TypedNamedValue = Jsonix.TypedNamedValue>(node: Node): Promise<E> {
  return (await getContext()).createUnmarshaller().unmarshalDocument<E>(node);
}

/** docx4j XmlUtils.marshaltoW3CDomDocument: the element as a DOM element, with docx4j's prefixes and namespace declarations. */
export async function marshalNode(element: Jsonix.TypedNamedValue): Promise<Element> {
  return marshalToDocument(await getContext(), element).documentElement;
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
function namespacePrefixesFor(context: Jsonix.Context, root: Jsonix.TypedNamedValue): Record<string, string> {
  const table: Record<string, string> = { ...context.namespacePrefixes };
  if (root.name.namespaceURI === RELATIONSHIPS_NS) {
    table[RELATIONSHIPS_NS] = '';
    if (table[SML_NS] === '') table[SML_NS] = 's';
  }
  return table;
}

/**
 * Marshals through the context with the root-specific table, then declares on the root only what
 * the tree uses (CR-001 section 3): the runtime declares every entry of the table on the root
 * element, used or not. The context is not modified: the marshaller sees the table through a
 * derived object, the runtime's `namespacePrefixes` field being the documented option (jsonix-CR-003
 * part 2, deferred to 3.3.0, is the runtime fix; when it lands this reduces to createMarshaller().marshalDocument).
 */
function marshalToDocument(context: Jsonix.Context, element: Jsonix.TypedNamedValue): Document {
  const derived = Object.create(context, { namespacePrefixes: { value: namespacePrefixesFor(context, element) } }) as Jsonix.Context;
  const doc = derived.createMarshaller().marshalDocument(element);
  fixRootNamespaceDeclarations(doc.documentElement, namespacePrefixesFor(context, element));
  return doc;
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
 * root's mc:Ignorable names its prefix, and every prefix mc:Ignorable names that is not declared
 * is added from `table`. A prefix the table cannot resolve is dropped from mc:Ignorable with a
 * warning: Word and Excel repair a file whose mc:Ignorable names an undeclared prefix, which is
 * how this was found (@docx4j/core-ts, an Excel acceptance run over xl/workbook.xml, 2026-09-19:
 * the model binds no xr:revisionPtr or xr2:uid, so nothing in the tree used xr2, xr6 or xr10 and
 * the marshaller declared none of them).
 *
 * `xmlns:xml` is never written. Declarations below the root shadow the root's, so a part's own
 * declarations inside a flat package count for the part, not for the package root.
 */
function fixRootNamespaceDeclarations(root: Element, table: Record<string, string>): void {
  const declared = new Map<string, Attr>();
  for (const attr of Array.from(root.attributes)) {
    const prefix = prefixOfDeclaration(attr);
    if (prefix !== undefined) declared.set(prefix, attr);
  }
  const used = new Set<string>();
  const use = (prefix: string, namespaceURI: string, shadowed: Set<string>): void => {
    if (!shadowed.has(prefix) && declared.get(prefix)?.value === namespaceURI) used.add(prefix);
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
    for (const child of Array.from(element.childNodes)) {
      if (child.nodeType === 1) walk(child as Element, scope);   // childNodes: xmldom has no firstElementChild
    }
  };
  walk(root, new Set());
  const ignorable = (root.getAttributeNS(MC_NS, 'Ignorable') ?? '').split(/\s+/).filter(Boolean);
  for (const prefix of ignorable) used.add(prefix);
  for (const [prefix, attr] of declared) {
    if (!used.has(prefix) || attr.value === XML_NS) root.removeAttributeNode(attr);
  }
  if (ignorable.length === 0) return;

  const namespaceFor = new Map(Object.entries(table).map(([namespaceURI, prefix]) => [prefix, namespaceURI]));
  const unresolved: string[] = [];
  for (const prefix of ignorable) {
    if (root.getAttributeNode(`xmlns:${prefix}`)) continue;
    const namespaceURI = namespaceFor.get(prefix);
    if (namespaceURI === undefined) { unresolved.push(prefix); continue; }
    root.setAttributeNS(XMLNS_NS, `xmlns:${prefix}`, namespaceURI);
  }
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
 * rest as DOM. Reverse with `marshalPackage` before `insertOoxml()`.
 */
export async function unmarshalPackage(ooxml: string): Promise<PackageElement> {
  const pkg = await unmarshalString<PackageElement>(ooxml);
  const unmarshaller = (await getContext()).createUnmarshaller();
  for (const part of pkg.value.part ?? []) {
    const any = part.xmlData?.any;
    if (part.xmlData && isNode(any)) {
      try {
        part.xmlData.any = unmarshaller.unmarshalDocument(any);
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
