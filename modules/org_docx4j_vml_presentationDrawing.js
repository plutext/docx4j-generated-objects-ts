var org_docx4j_vml_presentationDrawing_Module_Factory = function () {
  var org_docx4j_vml_presentationDrawing = {
    name: 'org_docx4j_vml_presentationDrawing',
    defaultElementNamespaceURI: 'urn:schemas-microsoft-com:office:powerpoint',
    typeInfos: [{
        localName: 'CTEmpty',
        typeName: 'CT_Empty'
      }, {
        localName: 'CTRel',
        typeName: 'CT_Rel',
        propertyInfos: [{
            name: 'id',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTEmpty',
        elementName: 'iscomment'
      }, {
        typeInfo: '.CTRel',
        elementName: 'textdata'
      }]
  };
  return {
    org_docx4j_vml_presentationDrawing: org_docx4j_vml_presentationDrawing
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_vml_presentationDrawing_Module_Factory);
}
else {
  var org_docx4j_vml_presentationDrawing_Module = org_docx4j_vml_presentationDrawing_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_vml_presentationDrawing = org_docx4j_vml_presentationDrawing_Module.org_docx4j_vml_presentationDrawing;
  }
  else {
    var org_docx4j_vml_presentationDrawing = org_docx4j_vml_presentationDrawing_Module.org_docx4j_vml_presentationDrawing;
  }
}