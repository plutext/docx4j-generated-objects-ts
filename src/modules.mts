// The mapping registry: every generated module, imported by a literal specifier (CR-005).
//
// Hand-written, like the list it replaces: generate.sh does not write this file, so add or remove
// an import when a regeneration adds or drops a module (the smoke test fails on a stale registry,
// since a mapping then references a module the context never loaded). MODULE_NAMES and the
// ModuleName type in the facade derive from MODULES, so this is the only list to keep.
//
// Literal specifiers, not `import(`../modules/${name}.mjs`)`, because a bundler cannot see what a
// template names: Vite leaves it a runtime URL and 404s unless configured, and esbuild expands it
// to a glob over ../modules/*.mjs, which drags in the `.el.mjs` and `.factory.mjs` siblings of
// every module - three files bundled for each one that loads. Here a bundler sees 103 imports and
// takes exactly those.

import { org_docx4j_bibliography } from '../modules/org_docx4j_bibliography.mjs';
import { org_docx4j_cei } from '../modules/org_docx4j_cei.mjs';
import { org_docx4j_com_microsoft_schemas_ink_x2010_main } from '../modules/org_docx4j_com_microsoft_schemas_ink_x2010_main.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2010_chartDrawing } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2010_chartDrawing.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2010_diagram } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2010_diagram.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2010_main } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2010_main.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2012_chart } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2012_chart.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2012_chartStyle } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2012_chartStyle.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2012_main } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2012_main.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2014_main } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2014_main.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x201611_main } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x201611_main.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2017_model3d } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2017_model3d.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d.mjs';
import { org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor } from '../modules/org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor.mjs';
import { org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction } from '../modules/org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction.mjs';
import { org_docx4j_com_microsoft_schemas_office_thememl_x2012_main } from '../modules/org_docx4j_com_microsoft_schemas_office_thememl_x2012_main.mjs';
import { org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11 } from '../modules/org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11.mjs';
import { org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11 } from '../modules/org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11.mjs';
import { org_docx4j_com_microsoft_schemas_office_word_x2006_wordml } from '../modules/org_docx4j_com_microsoft_schemas_office_word_x2006_wordml.mjs';
import { org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas } from '../modules/org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas.mjs';
import { org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing } from '../modules/org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing.mjs';
import { org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingGroup } from '../modules/org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingGroup.mjs';
import { org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape } from '../modules/org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape.mjs';
import { org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing } from '../modules/org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing.mjs';
import { org_docx4j_com_microsoft_schemas_office_x2006_encryption } from '../modules/org_docx4j_com_microsoft_schemas_office_x2006_encryption.mjs';
import { org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate } from '../modules/org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate.mjs';
import { org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password } from '../modules/org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password.mjs';
import { org_docx4j_customXmlProperties } from '../modules/org_docx4j_customXmlProperties.mjs';
import { org_docx4j_customxml } from '../modules/org_docx4j_customxml.mjs';
import { org_docx4j_dml } from '../modules/org_docx4j_dml.mjs';
import { org_docx4j_dml_chart } from '../modules/org_docx4j_dml_chart.mjs';
import { org_docx4j_dml_chartDrawing } from '../modules/org_docx4j_dml_chartDrawing.mjs';
import { org_docx4j_dml_chart_x2007 } from '../modules/org_docx4j_dml_chart_x2007.mjs';
import { org_docx4j_dml_compatibility } from '../modules/org_docx4j_dml_compatibility.mjs';
import { org_docx4j_dml_diagram } from '../modules/org_docx4j_dml_diagram.mjs';
import { org_docx4j_dml_diagram2008 } from '../modules/org_docx4j_dml_diagram2008.mjs';
import { org_docx4j_dml_lockedCanvas } from '../modules/org_docx4j_dml_lockedCanvas.mjs';
import { org_docx4j_dml_picture } from '../modules/org_docx4j_dml_picture.mjs';
import { org_docx4j_dml_spreadsheetdrawing } from '../modules/org_docx4j_dml_spreadsheetdrawing.mjs';
import { org_docx4j_dml_wordprocessingDrawing } from '../modules/org_docx4j_dml_wordprocessingDrawing.mjs';
import { org_docx4j_docProps_core } from '../modules/org_docx4j_docProps_core.mjs';
import { org_docx4j_docProps_core_dc_elements } from '../modules/org_docx4j_docProps_core_dc_elements.mjs';
import { org_docx4j_docProps_core_dc_terms } from '../modules/org_docx4j_docProps_core_dc_terms.mjs';
import { org_docx4j_docProps_coverPageProps } from '../modules/org_docx4j_docProps_coverPageProps.mjs';
import { org_docx4j_docProps_custom } from '../modules/org_docx4j_docProps_custom.mjs';
import { org_docx4j_docProps_extended } from '../modules/org_docx4j_docProps_extended.mjs';
import { org_docx4j_docProps_variantTypes } from '../modules/org_docx4j_docProps_variantTypes.mjs';
import { org_docx4j_math } from '../modules/org_docx4j_math.mjs';
import { org_docx4j_mce } from '../modules/org_docx4j_mce.mjs';
import { org_docx4j_org_w3_x1998_math_mathML } from '../modules/org_docx4j_org_w3_x1998_math_mathML.mjs';
import { org_docx4j_org_w3_x2003_inkML } from '../modules/org_docx4j_org_w3_x2003_inkML.mjs';
import { org_docx4j_relationships } from '../modules/org_docx4j_relationships.mjs';
import { org_docx4j_sharedtypes } from '../modules/org_docx4j_sharedtypes.mjs';
import { org_docx4j_vml } from '../modules/org_docx4j_vml.mjs';
import { org_docx4j_vml_officedrawing } from '../modules/org_docx4j_vml_officedrawing.mjs';
import { org_docx4j_vml_presentationDrawing } from '../modules/org_docx4j_vml_presentationDrawing.mjs';
import { org_docx4j_vml_root } from '../modules/org_docx4j_vml_root.mjs';
import { org_docx4j_vml_spreadsheetDrawing } from '../modules/org_docx4j_vml_spreadsheetDrawing.mjs';
import { org_docx4j_vml_wordprocessingDrawing } from '../modules/org_docx4j_vml_wordprocessingDrawing.mjs';
import { org_docx4j_w14 } from '../modules/org_docx4j_w14.mjs';
import { org_docx4j_w15 } from '../modules/org_docx4j_w15.mjs';
import { org_docx4j_w15symex } from '../modules/org_docx4j_w15symex.mjs';
import { org_docx4j_w16 } from '../modules/org_docx4j_w16.mjs';
import { org_docx4j_w16cex } from '../modules/org_docx4j_w16cex.mjs';
import { org_docx4j_w16cid } from '../modules/org_docx4j_w16cid.mjs';
import { org_docx4j_wml } from '../modules/org_docx4j_wml.mjs';
import { org_docx4j_xmlPackage } from '../modules/org_docx4j_xmlPackage.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x2010_main } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x2010_main.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_sectionzoom } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_sectionzoom.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main.mjs';
import { org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main } from '../modules/org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main.mjs';
import { org_pptx4j_pml } from '../modules/org_pptx4j_pml.mjs';
import { org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main } from '../modules/org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main.mjs';
import { org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac } from '../modules/org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac.mjs';
import { org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_main } from '../modules/org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_main.mjs';
import { org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2011_x1_ac } from '../modules/org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2011_x1_ac.mjs';
import { org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision } from '../modules/org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision.mjs';
import { org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_x11_main } from '../modules/org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_x11_main.mjs';
import { org_xlsx4j_schemas_microsoft_com_office_excel_2006_main } from '../modules/org_xlsx4j_schemas_microsoft_com_office_excel_2006_main.mjs';
import { org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main } from '../modules/org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main.mjs';
import { org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing } from '../modules/org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing.mjs';
import { org_xlsx4j_sml } from '../modules/org_xlsx4j_sml.mjs';

/** Every mapping module, by name, in MODULE_NAMES order. The context needs them all: they reference each other. */
export const MODULES = {
  org_docx4j_bibliography,
  org_docx4j_cei,
  org_docx4j_com_microsoft_schemas_ink_x2010_main,
  org_docx4j_com_microsoft_schemas_office_drawing_x2010_chartDrawing,
  org_docx4j_com_microsoft_schemas_office_drawing_x2010_diagram,
  org_docx4j_com_microsoft_schemas_office_drawing_x2010_main,
  org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture,
  org_docx4j_com_microsoft_schemas_office_drawing_x2012_chart,
  org_docx4j_com_microsoft_schemas_office_drawing_x2012_chartStyle,
  org_docx4j_com_microsoft_schemas_office_drawing_x2012_main,
  org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command,
  org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart,
  org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac,
  org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex,
  org_docx4j_com_microsoft_schemas_office_drawing_x2014_main,
  org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram,
  org_docx4j_com_microsoft_schemas_office_drawing_x201611_main,
  org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram,
  org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main,
  org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink,
  org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart,
  org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative,
  org_docx4j_com_microsoft_schemas_office_drawing_x2017_model3d,
  org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation,
  org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d,
  org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor,
  org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction,
  org_docx4j_com_microsoft_schemas_office_thememl_x2012_main,
  org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11,
  org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11,
  org_docx4j_com_microsoft_schemas_office_word_x2006_wordml,
  org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas,
  org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing,
  org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingGroup,
  org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape,
  org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing,
  org_docx4j_com_microsoft_schemas_office_x2006_encryption,
  org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate,
  org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password,
  org_docx4j_customXmlProperties,
  org_docx4j_customxml,
  org_docx4j_dml,
  org_docx4j_dml_chart,
  org_docx4j_dml_chartDrawing,
  org_docx4j_dml_chart_x2007,
  org_docx4j_dml_compatibility,
  org_docx4j_dml_diagram,
  org_docx4j_dml_diagram2008,
  org_docx4j_dml_lockedCanvas,
  org_docx4j_dml_picture,
  org_docx4j_dml_spreadsheetdrawing,
  org_docx4j_dml_wordprocessingDrawing,
  org_docx4j_docProps_core,
  org_docx4j_docProps_core_dc_elements,
  org_docx4j_docProps_core_dc_terms,
  org_docx4j_docProps_coverPageProps,
  org_docx4j_docProps_custom,
  org_docx4j_docProps_extended,
  org_docx4j_docProps_variantTypes,
  org_docx4j_math,
  org_docx4j_mce,
  org_docx4j_org_w3_x1998_math_mathML,
  org_docx4j_org_w3_x2003_inkML,
  org_docx4j_relationships,
  org_docx4j_sharedtypes,
  org_docx4j_vml,
  org_docx4j_vml_officedrawing,
  org_docx4j_vml_presentationDrawing,
  org_docx4j_vml_root,
  org_docx4j_vml_spreadsheetDrawing,
  org_docx4j_vml_wordprocessingDrawing,
  org_docx4j_w14,
  org_docx4j_w15,
  org_docx4j_w15symex,
  org_docx4j_w16,
  org_docx4j_w16cex,
  org_docx4j_w16cid,
  org_docx4j_wml,
  org_docx4j_xmlPackage,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x2010_main,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_sectionzoom,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main,
  org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main,
  org_pptx4j_pml,
  org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main,
  org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac,
  org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_main,
  org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2011_x1_ac,
  org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision,
  org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_x11_main,
  org_xlsx4j_schemas_microsoft_com_office_excel_2006_main,
  org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main,
  org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing,
  org_xlsx4j_sml,
} as const;
