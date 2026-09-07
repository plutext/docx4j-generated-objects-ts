var org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing = {
    name: 'org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordprocessingDrawing',
    typeInfos: [{
        localName: 'CTSizeRelH',
        typeName: 'CT_SizeRelH',
        propertyInfos: [{
            name: 'pctWidth',
            required: true,
            typeInfo: 'Int'
          }, {
            name: 'relativeFrom',
            required: true,
            typeInfo: '.STSizeRelFromH',
            attributeName: {
              localPart: 'relativeFrom'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSizeRelV',
        typeName: 'CT_SizeRelV',
        propertyInfos: [{
            name: 'pctHeight',
            required: true,
            typeInfo: 'Int'
          }, {
            name: 'relativeFrom',
            required: true,
            typeInfo: '.STSizeRelFromV',
            attributeName: {
              localPart: 'relativeFrom'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STSizeRelFromH',
        baseTypeInfo: 'Token',
        values: ['margin', 'page', 'leftMargin', 'rightMargin', 'insideMargin', 'outsideMargin']
      }, {
        type: 'enumInfo',
        localName: 'STSizeRelFromV',
        baseTypeInfo: 'Token',
        values: ['margin', 'page', 'topMargin', 'bottomMargin', 'insideMargin', 'outsideMargin']
      }],
    elementInfos: [{
        typeInfo: 'Int',
        elementName: 'pctPosHOffset'
      }, {
        typeInfo: 'Int',
        elementName: 'pctPosVOffset'
      }, {
        typeInfo: '.CTSizeRelH',
        elementName: 'sizeRelH'
      }, {
        typeInfo: '.CTSizeRelV',
        elementName: 'sizeRelV'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing: org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing_Module = org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing = org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing_Module.org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing = org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing_Module.org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing;
  }
}