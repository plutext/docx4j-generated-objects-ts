var org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2018\/animation\/model3d',
    dependencies: ['org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation', 'org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTEmbeddedAnimation',
        typeName: 'CT_EmbeddedAnimation',
        propertyInfos: [{
            name: 'animPr',
            required: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation.CTAnimationProperties'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'animId',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'animId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPosterFrame',
        typeName: 'CT_PosterFrame',
        propertyInfos: [{
            name: 'animId',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'animId'
            },
            type: 'attribute'
          }, {
            name: 'frame',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'frame'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTEmbeddedAnimation',
        elementName: 'embedAnim'
      }, {
        typeInfo: '.CTPosterFrame',
        elementName: 'posterFrame'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d: org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d = org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d = org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_model3d;
  }
}