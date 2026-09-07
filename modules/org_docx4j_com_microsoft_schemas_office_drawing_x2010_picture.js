var org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2010\/picture',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [],
    elementInfos: [{
        typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList',
        elementName: 'extLst'
      }, {
        typeInfo: 'org_docx4j_dml.CTShapeStyle',
        elementName: 'style'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture: org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture = org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture = org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2010_picture;
  }
}