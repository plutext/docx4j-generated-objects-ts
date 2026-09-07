var org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2018\/hyperlinkcolor',
    typeInfos: [{
        localName: 'CTHyperlinkColor',
        typeName: 'CT_HyperlinkColor',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: '.STHyperlinkColor',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STHyperlinkColor',
        baseTypeInfo: 'Token',
        values: ['hlink', 'tx']
      }],
    elementInfos: [{
        typeInfo: '.CTHyperlinkColor',
        elementName: 'hlinkClr'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor: org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor = org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor = org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2018_hyperlinkcolor;
  }
}