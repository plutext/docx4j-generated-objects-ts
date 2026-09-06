var org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main_Module_Factory = function () {
  var org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main = {
    name: 'org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2008\/2\/main',
    dependencies: ['org_xlsx4j_schemas_microsoft_com_office_excel_2006_main'],
    typeInfos: [{
        localName: 'CTDefinedName',
        typeName: 'CT_DefinedName',
        propertyInfos: [{
            name: 'argumentDescriptions',
            typeInfo: '.CTDefinedNameArgumentDescriptions'
          }, {
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDefinedNameArgumentDescription',
        typeName: 'CT_DefinedNameArgumentDescription',
        propertyInfos: [{
            name: 'value',
            type: 'value'
          }, {
            name: 'index',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'index'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDefinedNameArgumentDescriptions',
        typeName: 'CT_DefinedNameArgumentDescriptions',
        propertyInfos: [{
            name: 'argumentDescription',
            required: true,
            collection: true,
            typeInfo: '.CTDefinedNameArgumentDescription'
          }, {
            name: 'count',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'count'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDefinedNames',
        typeName: 'CT_DefinedNames',
        propertyInfos: [{
            name: 'definedName',
            required: true,
            collection: true,
            typeInfo: '.CTDefinedName'
          }]
      }],
    elementInfos: [{
        typeInfo: 'org_xlsx4j_schemas_microsoft_com_office_excel_2006_main.CTWorksheet',
        elementName: 'macrosheet'
      }]
  };
  return {
    org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main: org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main_Module_Factory);
}
else {
  var org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main_Module = org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main = org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main_Module.org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main;
  }
  else {
    var org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main = org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main_Module.org_xlsx4j_schemas_microsoft_com_office_excel_2008_2_main;
  }
}