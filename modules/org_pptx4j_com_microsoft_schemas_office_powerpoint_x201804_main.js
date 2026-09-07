var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main_Module_Factory = function () {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main = {
    name: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/powerpoint\/2018\/4\/main',
    typeInfos: [{
        localName: 'CTClassificationOutcome',
        typeName: 'CT_ClassificationOutcome',
        propertyInfos: [{
            name: 'val',
            typeInfo: '.STClassificationOutcomeType',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STClassificationOutcomeType',
        baseTypeInfo: 'Token',
        values: ['none', 'hdr', 'ftr', 'watermark']
      }],
    elementInfos: [{
        typeInfo: '.CTClassificationOutcome',
        elementName: 'classification'
      }]
  };
  return {
    org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main: org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main_Module_Factory);
}
else {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main_Module = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main;
  }
  else {
    var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201804_main;
  }
}