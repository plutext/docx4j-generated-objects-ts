var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main_Module_Factory = function () {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main = {
    name: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/powerpoint\/2012\/main',
    dependencies: ['org_docx4j_dml', 'org_pptx4j_pml'],
    typeInfos: [{
        localName: 'CTChartTrackingRefBased',
        typeName: 'CT_ChartTrackingRefBased',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCommentThreading',
        typeName: 'CT_CommentThreading',
        propertyInfos: [{
            name: 'parentCm',
            typeInfo: '.CTParentCommentIdentifier'
          }, {
            name: 'timeZoneBias',
            typeInfo: 'Int',
            attributeName: {
              localPart: 'timeZoneBias'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTExtendedGuide',
        typeName: 'CT_ExtendedGuide',
        propertyInfos: [{
            name: 'clr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTColor'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }, {
            name: 'id',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'name',
            defaultValue: '',
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'orient',
            typeInfo: 'org_pptx4j_pml.STDirection',
            defaultValue: 'vert',
            attributeName: {
              localPart: 'orient'
            },
            type: 'attribute'
          }, {
            name: 'pos',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'pos'
            },
            type: 'attribute'
          }, {
            name: 'userDrawn',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'userDrawn'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTExtendedGuideList',
        typeName: 'CT_ExtendedGuideList',
        propertyInfos: [{
            name: 'guide',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTExtendedGuide'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }]
      }, {
        localName: 'CTParentCommentIdentifier',
        typeName: 'CT_ParentCommentIdentifier',
        propertyInfos: [{
            name: 'authorId',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'authorId'
            },
            type: 'attribute'
          }, {
            name: 'idx',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'idx'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPresenceInfo',
        typeName: 'CT_PresenceInfo',
        propertyInfos: [{
            name: 'userId',
            required: true,
            attributeName: {
              localPart: 'userId'
            },
            type: 'attribute'
          }, {
            name: 'providerId',
            required: true,
            attributeName: {
              localPart: 'providerId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPresetTransition',
        typeName: 'CT_PresetTransition',
        propertyInfos: [{
            name: 'prst',
            attributeName: {
              localPart: 'prst'
            },
            type: 'attribute'
          }, {
            name: 'invX',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'invX'
            },
            type: 'attribute'
          }, {
            name: 'invY',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'invY'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTChartTrackingRefBased',
        elementName: 'chartTrackingRefBased'
      }, {
        typeInfo: '.CTExtendedGuideList',
        elementName: 'notesGuideLst'
      }, {
        typeInfo: '.CTPresenceInfo',
        elementName: 'presenceInfo'
      }, {
        typeInfo: '.CTPresetTransition',
        elementName: 'prstTrans'
      }, {
        typeInfo: '.CTExtendedGuideList',
        elementName: 'sldGuideLst'
      }, {
        typeInfo: '.CTCommentThreading',
        elementName: 'threadingInfo'
      }]
  };
  return {
    org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main: org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main_Module_Factory);
}
else {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main_Module = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main;
  }
  else {
    var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2012_main;
  }
}