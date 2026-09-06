var org_docx4j_dml_picture_Module_Factory = function () {
  var org_docx4j_dml_picture = {
    name: 'org_docx4j_dml_picture',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/picture',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTPictureNonVisual',
        typeName: 'CT_PictureNonVisual',
        propertyInfos: [{
            name: 'cNvPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualDrawingProps'
          }, {
            name: 'cNvPicPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualPictureProperties'
          }]
      }, {
        localName: 'Pic',
        typeName: 'CT_Picture',
        propertyInfos: [{
            name: 'nvPicPr',
            required: true,
            typeInfo: '.CTPictureNonVisual'
          }, {
            name: 'blipFill',
            required: true,
            typeInfo: 'org_docx4j_dml.CTBlipFillProperties'
          }, {
            name: 'spPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }]
      }],
    elementInfos: [{
        typeInfo: '.Pic',
        elementName: 'pic'
      }]
  };
  return {
    org_docx4j_dml_picture: org_docx4j_dml_picture
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_dml_picture_Module_Factory);
}
else {
  var org_docx4j_dml_picture_Module = org_docx4j_dml_picture_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_dml_picture = org_docx4j_dml_picture_Module.org_docx4j_dml_picture;
  }
  else {
    var org_docx4j_dml_picture = org_docx4j_dml_picture_Module.org_docx4j_dml_picture;
  }
}