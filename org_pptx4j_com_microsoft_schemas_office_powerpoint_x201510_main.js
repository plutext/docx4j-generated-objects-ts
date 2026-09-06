var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main_Module_Factory = function () {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main = {
    name: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/powerpoint\/2015\/10\/main',
    dependencies: ['org_pptx4j_pml'],
    typeInfos: [{
        localName: 'CTClientRevision',
        typeName: 'CT_ClientRevision',
        propertyInfos: [{
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'v',
            typeInfo: 'UnsignedInt',
            defaultValue: 0,
            attributeName: {
              localPart: 'v'
            },
            type: 'attribute'
          }, {
            name: 'vWet',
            typeInfo: 'UnsignedInt',
            defaultValue: 0,
            attributeName: {
              localPart: 'vWet'
            },
            type: 'attribute'
          }, {
            name: 'dt',
            required: true,
            typeInfo: 'DateTime',
            attributeName: {
              localPart: 'dt'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTClientRevisionList',
        typeName: 'CT_ClientRevisionList',
        propertyInfos: [{
            name: 'client',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTClientRevision'
          }]
      }, {
        localName: 'CTRevisionInfo',
        typeName: 'CT_RevisionInfo',
        propertyInfos: [{
            name: 'revLst',
            typeInfo: '.CTClientRevisionList'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTRevisionInfo',
        elementName: 'revInfo'
      }]
  };
  return {
    org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main: org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main_Module_Factory);
}
else {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main_Module = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main;
  }
  else {
    var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201510_main;
  }
}