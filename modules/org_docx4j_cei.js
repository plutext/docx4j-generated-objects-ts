var org_docx4j_cei_Module_Factory = function () {
  var org_docx4j_cei = {
    name: 'org_docx4j_cei',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2026\/wordml\/cei',
    defaultAttributeNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2026\/wordml\/cei',
    dependencies: ['org_docx4j_w16'],
    typeInfos: [{
        localName: 'CTCommentEntityInfo',
        typeName: 'CT_CommentEntityInfo',
        propertyInfos: [{
            name: 'extLst',
            typeInfo: 'org_docx4j_w16.CTExtensionList'
          }, {
            name: 'entityType',
            typeInfo: 'UnsignedInt',
            defaultValue: 0,
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTCommentEntityInfo',
        elementName: 'commentEntityInfo'
      }]
  };
  return {
    org_docx4j_cei: org_docx4j_cei
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_cei_Module_Factory);
}
else {
  var org_docx4j_cei_Module = org_docx4j_cei_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_cei = org_docx4j_cei_Module.org_docx4j_cei;
  }
  else {
    var org_docx4j_cei = org_docx4j_cei_Module.org_docx4j_cei;
  }
}