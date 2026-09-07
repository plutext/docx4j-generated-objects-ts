var org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas = {
    name: 'org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordprocessingCanvas',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTWordprocessingCanvas',
        typeName: 'CT_WordprocessingCanvas',
        propertyInfos: [{
            name: 'bg',
            typeInfo: 'org_docx4j_dml.CTBackgroundFormatting'
          }, {
            name: 'whole',
            typeInfo: 'org_docx4j_dml.CTWholeE2OFormatting'
          }, {
            name: 'wspOrPicOrContentPart',
            minOccurs: 0,
            collection: true,
            elementTypeInfos: [{
                elementName: {
                  localPart: 'wsp',
                  namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordprocessingShape'
                },
                typeInfo: 'org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingShape.CTWordprocessingShape'
              }, {
                elementName: {
                  localPart: 'pic',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/picture'
                },
                typeInfo: 'org_docx4j_dml_picture.Pic'
              }, {
                elementName: {
                  localPart: 'contentPart',
                  namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordml'
                },
                typeInfo: 'org_docx4j_w14.CTWordContentPart'
              }, {
                elementName: {
                  localPart: 'wgp',
                  namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordprocessingGroup'
                },
                typeInfo: 'org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingGroup.CTWordprocessingGroup'
              }, {
                elementName: 'graphicFrame',
                typeInfo: 'org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingGroup.CTGraphicFrame'
              }],
            type: 'elements'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTWordprocessingCanvas',
        elementName: 'wpc'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas: org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas_Module = org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas = org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas_Module.org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas = org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas_Module.org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingCanvas;
  }
}