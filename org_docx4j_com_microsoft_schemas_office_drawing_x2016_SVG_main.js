var org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2016\/SVG\/main',
    defaultAttributeNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships',
    typeInfos: [{
        localName: 'CTSVGBlip',
        typeName: 'CT_SVGBlip',
        propertyInfos: [{
            name: 'embed',
            defaultValue: '',
            type: 'attribute'
          }, {
            name: 'link',
            defaultValue: '',
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTSVGBlip',
        elementName: 'svgBlip'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main: org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main = org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main = org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2016_SVG_main;
  }
}