var org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2014\/chart\/ac',
    dependencies: ['org_docx4j_dml_chart'],
    typeInfos: [],
    elementInfos: [{
        typeInfo: 'org_docx4j_dml_chart.CTMultiLvlStrData',
        elementName: 'multiLvlStrLit'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac: org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac = org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac = org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2014_chart_ac;
  }
}