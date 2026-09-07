// The facade of @docx4j/docx4j-ts: docx4j's names over the Jsonix runtime and the generated
// Office Open XML mappings in ../modules. Hand-written; MODULE_NAMES is regenerated with the mappings.
import { Jsonix } from '@docx4j/jsonix';
export { Jsonix };
export type { TypedNamedValue, XmlQName, XmlCalendar, XmlDuration, JsonixMapping } from '../modules/org_docx4j_wml.mjs';

/** All generated modules, one Jsonix mapping each; they reference each other, so a context needs them all. */
export const MODULE_NAMES = [
  'org_docx4j_bibliography',
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
  'org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink',
  'org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main',
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
  'org_docx4j_customxml',
  'org_docx4j_customXmlProperties',
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
  'org_xlsx4j_schemas_microsoft_com_office_excel_2006_main',
  'org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main',
  'org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing',
  'org_xlsx4j_sml'
] as const;

export type ModuleName = (typeof MODULE_NAMES)[number];

let contextPromise: Promise<Jsonix.Context> | undefined;

/**
 * The context over all Office Open XML mappings, built once on first use (the modules are
 * loaded lazily; 94 modules are several megabytes). Parent pointers are on by default, as in
 * docx4j's object model; pass options to change that or anything else.
 */
export function getContext(options?: Jsonix.ContextOptions): Promise<Jsonix.Context> {
  if (contextPromise === undefined) {
    contextPromise = Promise.all(MODULE_NAMES.map((name) => import(`../modules/${name}.mjs`))).then(
      (modules) => new Jsonix.Context(modules.flatMap((m) => Object.values(m) as Jsonix.Mapping[]), { parentPointers: true, ...options }),
    );
  }
  return contextPromise;
}

/** Forgets the shared context, so that the next call to getContext builds a new one (e.g. with other options). */
export function resetContext(): void {
  contextPromise = undefined;
}

/** docx4j XmlUtils.unmarshalString: parse XML into a typed element, e.g. unmarshalString<DocumentElement>(xml). */
export async function unmarshalString<E extends Jsonix.TypedNamedValue = Jsonix.TypedNamedValue>(xml: string): Promise<E> {
  return (await getContext()).createUnmarshaller().unmarshalString<E>(xml);
}

/** docx4j XmlUtils.marshaltoString: serialise an element to XML. */
export async function marshalString(element: Jsonix.TypedNamedValue): Promise<string> {
  return (await getContext()).createMarshaller().marshalString(element);
}

/** docx4j XmlUtils.unwrap: the value of an element. */
export function unwrap<T>(element: Jsonix.TypedNamedValue<T>): T {
  return element.value;
}

/** docx4j XmlUtils.deepCopy: a structural copy with PARENT re-linked (Jsonix.Util.deepCopy). */
export function deepCopy<T>(value: T, parent?: unknown): T {
  return Jsonix.Util.deepCopy(value, parent);
}
