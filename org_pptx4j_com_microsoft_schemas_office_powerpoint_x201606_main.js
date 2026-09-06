var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main_Module_Factory = function () {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main = {
    name: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/powerpoint\/2016\/6\/main',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTZoomObjectProperties',
        typeName: 'CT_ZoomObjectProperties',
        propertyInfos: [{
            name: 'blipFill',
            required: true,
            typeInfo: 'org_docx4j_dml.CTBlipFillProperties'
          }, {
            name: 'spPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'id',
            required: true,
            typeInfo: 'Token',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'returnToParent',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'returnToParent'
            },
            type: 'attribute'
          }, {
            name: 'imageType',
            typeInfo: '.STZoomObjectImageType',
            defaultValue: 'preview',
            attributeName: {
              localPart: 'imageType'
            },
            type: 'attribute'
          }, {
            name: 'transitionDur',
            attributeName: {
              localPart: 'transitionDur'
            },
            type: 'attribute'
          }, {
            name: 'showBg',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'showBg'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STZoomObjectImageType',
        baseTypeInfo: 'Token',
        values: ['preview', 'cover']
      }],
    elementInfos: []
  };
  return {
    org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main: org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main_Module_Factory);
}
else {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main_Module = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main;
  }
  else {
    var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main;
  }
}