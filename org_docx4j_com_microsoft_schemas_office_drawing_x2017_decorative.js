var org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2017\/decorative',
    typeInfos: [{
        localName: 'CTDecorative',
        typeName: 'CT_Decorative',
        propertyInfos: [{
            name: 'val',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTDecorative',
        elementName: 'decorative'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative: org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative = org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative = org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2017_decorative;
  }
}