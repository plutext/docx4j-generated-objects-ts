var org_docx4j_dml_compatibility_Module_Factory = function () {
  var org_docx4j_dml_compatibility = {
    name: 'org_docx4j_dml_compatibility',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/compatibility',
    typeInfos: [{
        localName: 'CTCompat',
        typeName: 'CT_Compat',
        propertyInfos: [{
            name: 'spid',
            required: true,
            typeInfo: 'Token',
            attributeName: {
              localPart: 'spid'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTCompat',
        elementName: 'legacyDrawing'
      }]
  };
  return {
    org_docx4j_dml_compatibility: org_docx4j_dml_compatibility
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_dml_compatibility_Module_Factory);
}
else {
  var org_docx4j_dml_compatibility_Module = org_docx4j_dml_compatibility_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_dml_compatibility = org_docx4j_dml_compatibility_Module.org_docx4j_dml_compatibility;
  }
  else {
    var org_docx4j_dml_compatibility = org_docx4j_dml_compatibility_Module.org_docx4j_dml_compatibility;
  }
}