var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom_Module_Factory = function () {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom = {
    name: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/powerpoint\/2016\/slidezoom',
    dependencies: ['org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main', 'org_pptx4j_pml'],
    typeInfos: [{
        localName: 'CTSlideZoom',
        typeName: 'CT_SlideZoom',
        propertyInfos: [{
            name: 'sldZmObj',
            required: true,
            typeInfo: '.CTSlideZoomObject'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }]
      }, {
        localName: 'CTSlideZoomObject',
        typeName: 'CT_SlideZoomObject',
        propertyInfos: [{
            name: 'zmPr',
            required: true,
            typeInfo: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main.CTZoomObjectProperties'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }, {
            name: 'cId',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'cId'
            },
            type: 'attribute'
          }, {
            name: 'sldId',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'sldId'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTSlideZoom',
        elementName: 'sldZm'
      }]
  };
  return {
    org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom: org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom_Module_Factory);
}
else {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom_Module = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom;
  }
  else {
    var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_slidezoom;
  }
}