// CR-009: the WordprocessingML elements docx4j declares as roots (an @XmlRootElement class or an
// unscoped ObjectFactory @XmlElementDecl) and the generated mapping declares only within their
// parents (`w:drawing` in `.R`, `w:tbl` in `.Body`, ...) or only as properties (`w:pPr`). The
// runtime types an element inside a lax wildcard - a kept `mc:Choice` or `mc:Fallback`, a
// `w:pict` body, an extension `ext` - only if the context declares it globally
// (`getElementInfo` looks in the parent's scope, then in `##global`), so without this a kept
// branch's `w:drawing` stays DOM, where docx4j's context types it. One more mapping, hand-written,
// with no types of its own: it names the generated module's types, as the generated modules name
// each other's. Not in `modules/` (it is not generated) nor in the CR-005 registry (it is not a
// module to pick); `getContext` adds it to the default context and `modulesFor` to any closure
// holding `org_docx4j_wml`. A side effect: each of these is a valid root for `unmarshalString`
// and `marshalString`, as docx4j's root classes are for JAXB. `w:sdtContent` is left out: docx4j
// maps it to four classes, so it is ambiguous there too.
import type { Jsonix } from '@docx4j/jsonix';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * The 61 elements, local name to the mapping's type (an `org_docx4j_wml` type name, docx4j's
 * class name with `.` for nesting). List A of CR-009 (46 declared only within parents, each with
 * one type across its scopes) then list B (15 declared only as properties; for `pPr`, `sectPr`,
 * `tblPr` and `tblGrid` the root class, not the `...Base` type their `w:pPrChange` siblings use).
 * Source: docx4j VERSION_17_3_4 generated objects, compared with the mapping's `##global` scope.
 */
export const WML_GLOBAL_ELEMENTS: Readonly<Record<string, string>> = Object.freeze({
  // List A: declared here only within parents
  alias: 'SdtPr.Alias',
  altChunk: 'CTAltChunk',
  annotationRef: 'R.AnnotationRef',
  bibliography: 'SdtPr.Bibliography',
  bookmarkEnd: 'CTMarkupRange',
  bookmarkStart: 'CTBookmark',
  citation: 'SdtPr.Citation',
  commentReference: 'R.CommentReference',
  continuationSeparator: 'R.ContinuationSeparator',
  cr: 'R.Cr',
  dataBinding: 'CTDataBinding',
  dayLong: 'R.DayLong',
  dayShort: 'R.DayShort',
  drawing: 'Drawing',
  endnoteRef: 'R.EndnoteRef',
  equation: 'SdtPr.Equation',
  fldChar: 'FldChar',
  fldSimple: 'CTSimpleField',
  footnoteRef: 'R.FootnoteRef',
  group: 'SdtPr.Group',
  hyperlink: 'P.Hyperlink',
  jc: 'Jc',
  lastRenderedPageBreak: 'R.LastRenderedPageBreak',
  lock: 'CTLock',
  monthLong: 'R.MonthLong',
  monthShort: 'R.MonthShort',
  noBreakHyphen: 'R.NoBreakHyphen',
  permEnd: 'CTPerm',
  permStart: 'RangePermissionStart',
  pgNum: 'R.PgNum',
  pict: 'Pict',
  picture: 'SdtPr.Picture',
  ptab: 'R.Ptab',
  rPr: 'RPr',
  richText: 'SdtPr.RichText',
  separator: 'R.Separator',
  smartTag: 'CTSmartTagRun',
  softHyphen: 'R.SoftHyphen',
  spacing: 'CTSignedTwipsMeasure',
  sym: 'R.Sym',
  t: 'Text',
  tab: 'R.Tab',
  tbl: 'Tbl',
  tc: 'Tc',
  yearLong: 'R.YearLong',
  yearShort: 'R.YearShort',
  // List B: declared here only as properties
  ind: 'PPrBase.Ind',
  numPr: 'PPrBase.NumPr',
  outlineLvl: 'PPrBase.OutlineLvl',
  pBdr: 'PPrBase.PBdr',
  tblPr: 'TblPr',
  abstractNum: 'Numbering.AbstractNum',
  num: 'Numbering.Num',
  lvlOverride: 'Numbering.Num.LvlOverride',
  numPicBullet: 'Numbering.NumPicBullet',
  sectPr: 'SectPr',
  pPr: 'PPr',
  tabs: 'Tabs',
  sdtPr: 'SdtPr',
  lvl: 'Lvl',
  tblGrid: 'TblGrid',
});

/**
 * The mapping that declares them: `dependencies` on the generated module, no types, one
 * `elementInfos` entry per name. Passed to `new Jsonix.Context` beside the generated modules by
 * `getContext`; a caller building its own list adds it to type kept branches.
 */
export const org_docx4j_wml_globals: Jsonix.Mapping = Object.freeze({
  name: 'org_docx4j_wml_globals',
  dependencies: ['org_docx4j_wml'],
  typeInfos: [],
  elementInfos: Object.entries(WML_GLOBAL_ELEMENTS).map(([localPart, type]) => ({
    elementName: { namespaceURI: W, localPart },
    typeInfo: `org_docx4j_wml.${type}`,
  })),
});
