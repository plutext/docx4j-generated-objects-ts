var org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing_Module_Factory = function () {
  var org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing = {
    name: 'org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2010\/spreadsheetDrawing',
    dependencies: ['org_docx4j_com_microsoft_schemas_office_drawing_x2010_main', 'org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTApplicationNonVisualDrawingProps',
        typeName: 'CT_ApplicationNonVisualDrawingProps',
        propertyInfos: [{
            name: 'fPublished',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'fPublished'
            },
            type: 'attribute'
          }, {
            name: 'macro',
            attributeName: {
              localPart: 'macro'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTContentPart',
        typeName: 'CT_ContentPart',
        propertyInfos: [{
            name: 'nvContentPartPr',
            typeInfo: '.CTContentPartNonVisual'
          }, {
            name: 'nvPr',
            typeInfo: '.CTApplicationNonVisualDrawingProps'
          }, {
            name: 'xfrm',
            typeInfo: 'org_docx4j_dml.CTTransform2D'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'bwMode',
            typeInfo: 'org_docx4j_dml.STBlackWhiteMode',
            defaultValue: 'auto',
            attributeName: {
              localPart: 'bwMode'
            },
            type: 'attribute'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTContentPartNonVisual',
        typeName: 'CT_ContentPartNonVisual',
        propertyInfos: [{
            name: 'cNvPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualDrawingProps'
          }, {
            name: 'cNvContentPartPr',
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2010_main.CTNonVisualInkContentPartProperties'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTContentPart',
        elementName: 'contentPart'
      }]
  };
  return {
    org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing: org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing_Module_Factory);
}
else {
  var org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing_Module = org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing = org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing_Module.org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing;
  }
  else {
    var org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing = org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing_Module.org_xlsx4j_schemas_microsoft_com_office_excel_x2010_spreadsheetDrawing;
  }
}