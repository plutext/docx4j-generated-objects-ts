var org_docx4j_w16_Module_Factory = function () {
  var org_docx4j_w16 = {
    name: 'org_docx4j_w16',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2018\/wordml',
    defaultAttributeNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2018\/wordml',
    typeInfos: [{
        localName: 'CTExtension',
        typeName: 'CT_Extension',
        propertyInfos: [{
            name: 'any',
            required: true,
            mixed: false,
            type: 'anyElement'
          }, {
            name: 'uri',
            typeInfo: 'Token',
            type: 'attribute'
          }]
      }, {
        localName: 'CTExtensionList',
        typeName: 'CT_ExtensionList',
        propertyInfos: [{
            name: 'ext',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTExtension'
          }]
      }],
    elementInfos: []
  };
  return {
    org_docx4j_w16: org_docx4j_w16
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_w16_Module_Factory);
}
else {
  var org_docx4j_w16_Module = org_docx4j_w16_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_w16 = org_docx4j_w16_Module.org_docx4j_w16;
  }
  else {
    var org_docx4j_w16 = org_docx4j_w16_Module.org_docx4j_w16;
  }
}