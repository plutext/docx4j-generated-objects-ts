var org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink',
    typeInfos: [{
        type: 'enumInfo',
        localName: 'STExtendedBrushPropertyName',
        values: ['inkEffects', 'anchorX', 'anchorY', 'scaleFactor']
      }, {
        type: 'enumInfo',
        localName: 'STInkEffectsType',
        values: ['none', 'pencil', 'rainbow', 'galaxy', 'gold', 'silver', 'lava', 'ocean', 'rosegold', 'bronze']
      }],
    elementInfos: []
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink: org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink = org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink = org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2016_ink;
  }
}