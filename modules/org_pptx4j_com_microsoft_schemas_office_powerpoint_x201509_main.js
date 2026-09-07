var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main_Module_Factory = function () {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main = {
    name: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/powerpoint\/2015\/09\/main',
    typeInfos: [{
        localName: 'CTMorphTransition',
        typeName: 'CT_MorphTransition',
        propertyInfos: [{
            name: 'option',
            required: true,
            typeInfo: '.STTransitionMorphOption',
            attributeName: {
              localPart: 'option'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STTransitionMorphOption',
        baseTypeInfo: 'Token',
        values: ['byObject', 'byWord', 'byChar']
      }],
    elementInfos: [{
        typeInfo: '.CTMorphTransition',
        elementName: 'morph'
      }]
  };
  return {
    org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main: org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main_Module_Factory);
}
else {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main_Module = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main;
  }
  else {
    var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201509_main;
  }
}