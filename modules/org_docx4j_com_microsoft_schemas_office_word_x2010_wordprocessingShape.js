var org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape = {
    name: 'org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordprocessingShape',
    dependencies: ['org_docx4j_dml', 'org_docx4j_wml'],
    typeInfos: [{
        localName: 'CTLinkedTextboxInformation',
        typeName: 'CT_LinkedTextboxInformation',
        propertyInfos: [{
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'id',
            required: true,
            typeInfo: 'UnsignedShort',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'seq',
            required: true,
            typeInfo: 'UnsignedShort',
            attributeName: {
              localPart: 'seq'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTTextboxInfo',
        typeName: 'CT_TextboxInfo',
        propertyInfos: [{
            name: 'txbxContent',
            elementName: {
              localPart: 'txbxContent',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
            },
            typeInfo: 'org_docx4j_wml.CTTxbxContent'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'id',
            typeInfo: 'UnsignedShort',
            defaultValue: 0,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTWordprocessingShape',
        typeName: 'CT_WordprocessingShape',
        propertyInfos: [{
            name: 'cNvPr',
            typeInfo: 'org_docx4j_dml.CTNonVisualDrawingProps'
          }, {
            name: 'cNvSpPr',
            typeInfo: 'org_docx4j_dml.CTNonVisualDrawingShapeProps'
          }, {
            name: 'cNvCnPr',
            typeInfo: 'org_docx4j_dml.CTNonVisualConnectorProperties'
          }, {
            name: 'spPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'style',
            typeInfo: 'org_docx4j_dml.CTShapeStyle'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'txbx',
            typeInfo: '.CTTextboxInfo'
          }, {
            name: 'linkedTxbx',
            typeInfo: '.CTLinkedTextboxInformation'
          }, {
            name: 'bodyPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTTextBodyProperties'
          }, {
            name: 'normalEastAsianFlow',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'normalEastAsianFlow'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTWordprocessingShape',
        elementName: 'wsp'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape: org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape_Module = org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape = org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape_Module.org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape = org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape_Module.org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape;
  }
}