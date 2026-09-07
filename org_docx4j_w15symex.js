var org_docx4j_w15symex_Module_Factory = function () {
  var org_docx4j_w15symex = {
    name: 'org_docx4j_w15symex',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2015\/wordml\/symex',
    defaultAttributeNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2015\/wordml\/symex',
    typeInfos: [{
        localName: 'CTSymEx',
        typeName: 'CT_SymEx',
        propertyInfos: [{
            name: '_char',
            attributeName: 'char',
            type: 'attribute'
          }, {
            name: 'font',
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTSymEx',
        elementName: 'symEx'
      }]
  };
  return {
    org_docx4j_w15symex: org_docx4j_w15symex
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_w15symex_Module_Factory);
}
else {
  var org_docx4j_w15symex_Module = org_docx4j_w15symex_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_w15symex = org_docx4j_w15symex_Module.org_docx4j_w15symex;
  }
  else {
    var org_docx4j_w15symex = org_docx4j_w15symex_Module.org_docx4j_w15symex;
  }
}