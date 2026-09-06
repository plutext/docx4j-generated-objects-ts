var org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11 = {
    name: 'org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/webextensions\/taskpanes\/2010\/11',
    dependencies: ['org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11', 'org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTOsfTaskpane',
        typeName: 'CT_OsfTaskpane',
        propertyInfos: [{
            name: 'webextensionref',
            required: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11.CTWebExtensionPartRef'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'dockstate',
            required: true,
            attributeName: {
              localPart: 'dockstate'
            },
            type: 'attribute'
          }, {
            name: 'visibility',
            required: true,
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'visibility'
            },
            type: 'attribute'
          }, {
            name: 'width',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'width'
            },
            type: 'attribute'
          }, {
            name: 'row',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'row'
            },
            type: 'attribute'
          }, {
            name: 'locked',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'locked'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOsfTaskpanes',
        typeName: 'CT_OsfTaskpanes',
        propertyInfos: [{
            name: 'taskpane',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTOsfTaskpane'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTOsfTaskpanes',
        elementName: 'taskpanes'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11: org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11_Module = org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11 = org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11_Module.org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11 = org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11_Module.org_docx4j_com_microsoft_schemas_office_webextensions_taskpanes_2010_11;
  }
}