var org_docx4j_com_microsoft_schemas_office_thememl_x2012_main_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_thememl_x2012_main = {
    name: 'org_docx4j_com_microsoft_schemas_office_thememl_x2012_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/thememl\/2012\/main',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTThemeFamily',
        typeName: 'CT_ThemeFamily',
        propertyInfos: [{
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'id',
            required: true,
            typeInfo: 'Token',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'vid',
            required: true,
            typeInfo: 'Token',
            attributeName: {
              localPart: 'vid'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTThemeFamily',
        elementName: 'themeFamily'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_thememl_x2012_main: org_docx4j_com_microsoft_schemas_office_thememl_x2012_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_thememl_x2012_main_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_thememl_x2012_main_Module = org_docx4j_com_microsoft_schemas_office_thememl_x2012_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_thememl_x2012_main = org_docx4j_com_microsoft_schemas_office_thememl_x2012_main_Module.org_docx4j_com_microsoft_schemas_office_thememl_x2012_main;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_thememl_x2012_main = org_docx4j_com_microsoft_schemas_office_thememl_x2012_main_Module.org_docx4j_com_microsoft_schemas_office_thememl_x2012_main;
  }
}