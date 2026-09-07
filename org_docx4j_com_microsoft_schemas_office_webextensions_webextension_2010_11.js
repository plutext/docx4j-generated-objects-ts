var org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11 = {
    name: 'org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/webextensions\/webextension\/2010\/11',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTOsfWebExtension',
        typeName: 'CT_OsfWebExtension',
        propertyInfos: [{
            name: 'reference',
            required: true,
            typeInfo: '.CTOsfWebExtensionReference'
          }, {
            name: 'alternateReferences',
            typeInfo: '.CTOsfWebExtensionReferenceList'
          }, {
            name: 'properties',
            required: true,
            typeInfo: '.CTOsfWebExtensionPropertyBag'
          }, {
            name: 'bindings',
            required: true,
            typeInfo: '.CTOsfWebExtensionBindingList'
          }, {
            name: 'snapshot',
            typeInfo: 'org_docx4j_dml.CTBlip'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'frozen',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'frozen'
            },
            type: 'attribute'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOsfWebExtensionBinding',
        typeName: 'CT_OsfWebExtensionBinding',
        propertyInfos: [{
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'appref',
            required: true,
            attributeName: {
              localPart: 'appref'
            },
            type: 'attribute'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'type',
            required: true,
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOsfWebExtensionBindingList',
        typeName: 'CT_OsfWebExtensionBindingList',
        propertyInfos: [{
            name: 'binding',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTOsfWebExtensionBinding'
          }]
      }, {
        localName: 'CTOsfWebExtensionProperty',
        typeName: 'CT_OsfWebExtensionProperty',
        propertyInfos: [{
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'value',
            required: true,
            attributeName: {
              localPart: 'value'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOsfWebExtensionPropertyBag',
        typeName: 'CT_OsfWebExtensionPropertyBag',
        propertyInfos: [{
            name: 'property',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTOsfWebExtensionProperty'
          }]
      }, {
        localName: 'CTOsfWebExtensionReference',
        typeName: 'CT_OsfWebExtensionReference',
        propertyInfos: [{
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'store',
            attributeName: {
              localPart: 'store'
            },
            type: 'attribute'
          }, {
            name: 'storeType',
            attributeName: {
              localPart: 'storeType'
            },
            type: 'attribute'
          }, {
            name: 'version',
            required: true,
            attributeName: {
              localPart: 'version'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOsfWebExtensionReferenceList',
        typeName: 'CT_OsfWebExtensionReferenceList',
        propertyInfos: [{
            name: 'reference',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTOsfWebExtensionReference'
          }]
      }, {
        localName: 'CTWebExtensionPartRef',
        typeName: 'CT_WebExtensionPartRef',
        propertyInfos: [{
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTOsfWebExtension',
        elementName: 'webextension'
      }, {
        typeInfo: '.CTWebExtensionPartRef',
        elementName: 'webextensionref'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11: org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11_Module = org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11 = org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11_Module.org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11 = org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11_Module.org_docx4j_com_microsoft_schemas_office_webextensions_webextension_2010_11;
  }
}