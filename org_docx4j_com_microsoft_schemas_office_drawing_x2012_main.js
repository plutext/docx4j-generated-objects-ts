var org_docx4j_com_microsoft_schemas_office_drawing_x2012_main_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2012_main = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2012_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2012\/main',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTBackgroundPr',
        typeName: 'CT_BackgroundPr',
        propertyInfos: [{
            name: 'bwMode',
            typeInfo: 'org_docx4j_dml.STBlackWhiteMode',
            attributeName: {
              localPart: 'bwMode'
            },
            type: 'attribute'
          }, {
            name: 'bwNormal',
            typeInfo: 'org_docx4j_dml.STBlackWhiteMode',
            attributeName: {
              localPart: 'bwNormal'
            },
            type: 'attribute'
          }, {
            name: 'bwPure',
            typeInfo: 'org_docx4j_dml.STBlackWhiteMode',
            attributeName: {
              localPart: 'bwPure'
            },
            type: 'attribute'
          }, {
            name: 'targetScreenSize',
            typeInfo: 'Token',
            values: ['544x376', '640x480', '720x512', '800x600', '1024x768', '1152x882', '1152x900', '1280x1024', '1600x1200', '1800x1440', '1920x1200'],
            attributeName: {
              localPart: 'targetScreenSize'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTNonVisualGroupProps',
        typeName: 'CT_NonVisualGroupProps',
        propertyInfos: [{
            name: 'isLegacyGroup',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'isLegacyGroup'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTObjectPr',
        typeName: 'CT_ObjectPr',
        propertyInfos: [{
            name: 'isActiveX',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'isActiveX'
            },
            type: 'attribute'
          }, {
            name: 'linkType',
            attributeName: {
              localPart: 'linkType'
            },
            type: 'attribute'
          }, {
            name: 'objectId',
            attributeName: {
              localPart: 'objectId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSignatureLine',
        typeName: 'CT_SignatureLine',
        propertyInfos: [{
            name: 'addlXml',
            attributeName: {
              localPart: 'addlXml'
            },
            type: 'attribute'
          }, {
            name: 'allowComments',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'allowComments'
            },
            type: 'attribute'
          }, {
            name: 'id',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'isSignatureLine',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'isSignatureLine'
            },
            type: 'attribute'
          }, {
            name: 'provId',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'provId'
            },
            type: 'attribute'
          }, {
            name: 'showSignDate',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'showSignDate'
            },
            type: 'attribute'
          }, {
            name: 'sigProvUrl',
            attributeName: {
              localPart: 'sigProvUrl'
            },
            type: 'attribute'
          }, {
            name: 'signingInstructions',
            attributeName: {
              localPart: 'signingInstructions'
            },
            type: 'attribute'
          }, {
            name: 'signingInstructionsSet',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'signingInstructionsSet'
            },
            type: 'attribute'
          }, {
            name: 'suggestedSigner',
            attributeName: {
              localPart: 'suggestedSigner'
            },
            type: 'attribute'
          }, {
            name: 'suggestedSigner2',
            attributeName: {
              localPart: 'suggestedSigner2'
            },
            type: 'attribute'
          }, {
            name: 'suggestedSignerEmail',
            attributeName: {
              localPart: 'suggestedSignerEmail'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTBackgroundPr',
        elementName: 'backgroundPr'
      }, {
        typeInfo: '.CTNonVisualGroupProps',
        elementName: 'nonVisualGroupProps'
      }, {
        typeInfo: '.CTObjectPr',
        elementName: 'objectPr'
      }, {
        typeInfo: '.CTSignatureLine',
        elementName: 'signatureLine'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2012_main: org_docx4j_com_microsoft_schemas_office_drawing_x2012_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2012_main_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2012_main_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2012_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2012_main = org_docx4j_com_microsoft_schemas_office_drawing_x2012_main_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2012_main;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2012_main = org_docx4j_com_microsoft_schemas_office_drawing_x2012_main_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2012_main;
  }
}