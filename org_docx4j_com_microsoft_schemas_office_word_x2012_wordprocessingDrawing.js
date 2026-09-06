var org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing = {
    name: 'org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2012\/wordprocessingDrawing',
    typeInfos: [{
        localName: 'CTWebVideoPr',
        typeName: 'CT_WebVideoPr',
        propertyInfos: [{
            name: 'embeddedHtml',
            defaultValue: '',
            attributeName: {
              localPart: 'embeddedHtml'
            },
            type: 'attribute'
          }, {
            name: 'h',
            typeInfo: 'UnsignedInt',
            defaultValue: 0,
            attributeName: {
              localPart: 'h'
            },
            type: 'attribute'
          }, {
            name: 'w',
            typeInfo: 'UnsignedInt',
            defaultValue: 0,
            attributeName: {
              localPart: 'w'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTWebVideoPr',
        elementName: 'webVideoPr'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing: org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing_Module = org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing = org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing_Module.org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing = org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing_Module.org_docx4j_com_microsoft_schemas_office_word_x2012_wordprocessingDrawing;
  }
}