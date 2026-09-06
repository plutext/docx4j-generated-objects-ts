var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main_Module_Factory = function () {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main = {
    name: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/powerpoint\/2015\/main',
    typeInfos: [{
        localName: 'CTDesignElement',
        typeName: 'CT_DesignElement',
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
        typeInfo: '.CTDesignElement',
        elementName: 'designElem'
      }]
  };
  return {
    org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main: org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main_Module_Factory);
}
else {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main_Module = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main;
  }
  else {
    var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2015_main;
  }
}