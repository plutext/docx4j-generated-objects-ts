var org_docx4j_vml_root_Module_Factory = function () {
  var org_docx4j_vml_root = {
    name: 'org_docx4j_vml_root',
    typeInfos: [{
        localName: 'Xml',
        typeName: null,
        propertyInfos: [{
            name: 'any',
            required: true,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }],
    elementInfos: [{
        typeInfo: '.Xml',
        elementName: {
          localPart: 'xml'
        }
      }]
  };
  return {
    org_docx4j_vml_root: org_docx4j_vml_root
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_vml_root_Module_Factory);
}
else {
  var org_docx4j_vml_root_Module = org_docx4j_vml_root_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_vml_root = org_docx4j_vml_root_Module.org_docx4j_vml_root;
  }
  else {
    var org_docx4j_vml_root = org_docx4j_vml_root_Module.org_docx4j_vml_root;
  }
}