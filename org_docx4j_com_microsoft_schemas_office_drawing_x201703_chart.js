var org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2017\/03\/chart',
    defaultAttributeNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2017\/03\/chart',
    typeInfos: [{
        localName: 'CTBooleanFalse',
        typeName: 'CT_BooleanFalse',
        propertyInfos: [{
            name: 'val',
            typeInfo: 'Boolean',
            defaultValue: false,
            type: 'attribute'
          }]
      }, {
        localName: 'CTDataDisplayOptions16',
        typeName: 'CT_DataDisplayOptions16',
        propertyInfos: [{
            name: 'dispNaAsBlank',
            typeInfo: '.CTBooleanFalse'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTDataDisplayOptions16',
        elementName: 'dataDisplayOptions16'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart: org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart_Module = org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart = org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart_Module.org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart = org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart_Module.org_docx4j_com_microsoft_schemas_office_drawing_x201703_chart;
  }
}