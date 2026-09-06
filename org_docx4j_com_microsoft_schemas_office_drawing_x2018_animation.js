var org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2018\/animation',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTAnimationProperties',
        typeName: 'CT_AnimationProperties',
        propertyInfos: [{
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'name',
            defaultValue: '',
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'length',
            required: true,
            attributeName: {
              localPart: 'length'
            },
            type: 'attribute'
          }, {
            name: 'count',
            defaultValue: '1',
            attributeName: {
              localPart: 'count'
            },
            type: 'attribute'
          }, {
            name: 'auto',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'auto'
            },
            type: 'attribute'
          }, {
            name: 'offset',
            defaultValue: '0',
            attributeName: {
              localPart: 'offset'
            },
            type: 'attribute'
          }, {
            name: 'st',
            defaultValue: '0',
            attributeName: {
              localPart: 'st'
            },
            type: 'attribute'
          }, {
            name: 'end',
            defaultValue: '0',
            attributeName: {
              localPart: 'end'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STIndefinite',
        baseTypeInfo: 'Token',
        values: ['indefinite']
      }],
    elementInfos: []
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation: org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation = org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation = org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2018_animation;
  }
}