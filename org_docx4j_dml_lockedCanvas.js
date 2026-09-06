var org_docx4j_dml_lockedCanvas_Module_Factory = function () {
  var org_docx4j_dml_lockedCanvas = {
    name: 'org_docx4j_dml_lockedCanvas',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/lockedCanvas',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [],
    elementInfos: [{
        typeInfo: 'org_docx4j_dml.CTGvmlGroupShape',
        elementName: 'lockedCanvas'
      }]
  };
  return {
    org_docx4j_dml_lockedCanvas: org_docx4j_dml_lockedCanvas
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_dml_lockedCanvas_Module_Factory);
}
else {
  var org_docx4j_dml_lockedCanvas_Module = org_docx4j_dml_lockedCanvas_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_dml_lockedCanvas = org_docx4j_dml_lockedCanvas_Module.org_docx4j_dml_lockedCanvas;
  }
  else {
    var org_docx4j_dml_lockedCanvas = org_docx4j_dml_lockedCanvas_Module.org_docx4j_dml_lockedCanvas;
  }
}