var org_docx4j_com_microsoft_schemas_office_drawing_x201611_main_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x201611_main = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x201611_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2016\/11\/main',
    defaultAttributeNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships',
    typeInfos: [{
        localName: 'CTPictureAttributionSourceURL',
        typeName: 'CT_PictureAttributionSourceURL',
        propertyInfos: [{
            name: 'id',
            required: true,
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTPictureAttributionSourceURL',
        elementName: 'picAttrSrcUrl'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x201611_main: org_docx4j_com_microsoft_schemas_office_drawing_x201611_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x201611_main_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x201611_main_Module = org_docx4j_com_microsoft_schemas_office_drawing_x201611_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x201611_main = org_docx4j_com_microsoft_schemas_office_drawing_x201611_main_Module.org_docx4j_com_microsoft_schemas_office_drawing_x201611_main;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x201611_main = org_docx4j_com_microsoft_schemas_office_drawing_x201611_main_Module.org_docx4j_com_microsoft_schemas_office_drawing_x201611_main;
  }
}