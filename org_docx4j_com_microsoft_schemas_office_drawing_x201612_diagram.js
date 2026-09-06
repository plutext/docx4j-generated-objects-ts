var org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2016\/12\/diagram',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [],
    elementInfos: [{
        typeInfo: 'org_docx4j_dml.CTTextListStyle',
        elementName: 'lstStyle'
      }, {
        typeInfo: 'org_docx4j_dml.CTShapeProperties',
        elementName: 'spPr'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram: org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram_Module = org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram = org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram_Module.org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram = org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram_Module.org_docx4j_com_microsoft_schemas_office_drawing_x201612_diagram;
  }
}