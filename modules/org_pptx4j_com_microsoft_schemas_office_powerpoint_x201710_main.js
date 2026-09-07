var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main_Module_Factory = function () {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main = {
    name: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/powerpoint\/2017\/10\/main',
    typeInfos: [{
        localName: 'CTReadonlyRecommended',
        typeName: 'CT_ReadonlyRecommended',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTReadonlyRecommended',
        elementName: 'readonlyRecommended'
      }]
  };
  return {
    org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main: org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main_Module_Factory);
}
else {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main_Module = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main;
  }
  else {
    var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201710_main;
  }
}