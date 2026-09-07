var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom_Module_Factory = function () {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom = {
    name: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/powerpoint\/2016\/summaryzoom',
    dependencies: ['org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main', 'org_pptx4j_pml'],
    typeInfos: [{
        localName: 'CTFixedLayout',
        typeName: 'CT_FixedLayout'
      }, {
        localName: 'CTGridLayout',
        typeName: 'CT_GridLayout'
      }, {
        localName: 'CTSummaryZoom',
        typeName: 'CT_SummaryZoom',
        propertyInfos: [{
            name: 'summaryZmObj',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTSummaryZoomObject'
          }, {
            name: 'gridLayout',
            typeInfo: '.CTGridLayout'
          }, {
            name: 'fixedLayout',
            typeInfo: '.CTFixedLayout'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }]
      }, {
        localName: 'CTSummaryZoomObject',
        typeName: 'CT_SummaryZoomObject',
        propertyInfos: [{
            name: 'zmPr',
            required: true,
            typeInfo: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201606_main.CTZoomObjectProperties'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }, {
            name: 'descr',
            defaultValue: '',
            attributeName: {
              localPart: 'descr'
            },
            type: 'attribute'
          }, {
            name: 'offsetFactorX',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'offsetFactorX'
            },
            type: 'attribute'
          }, {
            name: 'offsetFactorY',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'offsetFactorY'
            },
            type: 'attribute'
          }, {
            name: 'scaleFactorX',
            typeInfo: 'Int',
            defaultValue: 100000,
            attributeName: {
              localPart: 'scaleFactorX'
            },
            type: 'attribute'
          }, {
            name: 'scaleFactorY',
            typeInfo: 'Int',
            defaultValue: 100000,
            attributeName: {
              localPart: 'scaleFactorY'
            },
            type: 'attribute'
          }, {
            name: 'sectionId',
            required: true,
            typeInfo: 'Token',
            attributeName: {
              localPart: 'sectionId'
            },
            type: 'attribute'
          }, {
            name: 'title',
            defaultValue: '',
            attributeName: {
              localPart: 'title'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTSummaryZoom',
        elementName: 'summaryZm'
      }]
  };
  return {
    org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom: org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom_Module_Factory);
}
else {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom_Module = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom;
  }
  else {
    var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2016_summaryzoom;
  }
}