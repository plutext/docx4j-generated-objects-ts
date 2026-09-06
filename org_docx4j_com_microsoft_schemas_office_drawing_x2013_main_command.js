var org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2013\/main\/command',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTChangesData',
        typeName: 'CT_ChangesData',
        propertyInfos: [{
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'actId',
            typeInfo: 'Int',
            attributeName: {
              localPart: 'actId'
            },
            type: 'attribute'
          }, {
            name: 'dt',
            typeInfo: 'DateTime',
            attributeName: {
              localPart: 'dt'
            },
            type: 'attribute'
          }, {
            name: 'v',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'v'
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
            name: 'name',
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'userId',
            attributeName: {
              localPart: 'userId'
            },
            type: 'attribute'
          }, {
            name: 'providerId',
            attributeName: {
              localPart: 'providerId'
            },
            type: 'attribute'
          }, {
            name: 'clId',
            attributeName: {
              localPart: 'clId'
            },
            type: 'attribute'
          }, {
            name: 'email',
            attributeName: {
              localPart: 'email'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTConnectorChanges',
        typeName: 'CT_ConnectorChanges',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: '.CTChangesData'
          }, {
            name: 'cxnSpMkLst',
            required: true,
            typeInfo: '.CTConnectorMonikerList'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'chg',
            required: true,
            typeInfo: {
              type: 'list',
              baseTypeInfo: '.STConnectorChangeBit'
            },
            attributeName: {
              localPart: 'chg'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTConnectorMoniker',
        typeName: 'CT_ConnectorMoniker',
        propertyInfos: [{
            name: 'id',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'creationId',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'creationId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTConnectorMonikerList',
        typeName: 'CT_ConnectorMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'CTGraphicFrameChanges',
        typeName: 'CT_GraphicFrameChanges',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: '.CTChangesData'
          }, {
            name: 'graphicFrameMkLst',
            required: true,
            typeInfo: '.CTGraphicFrameMonikerList'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'chg',
            required: true,
            typeInfo: {
              type: 'list',
              baseTypeInfo: '.STGraphicFrameChangeBit'
            },
            attributeName: {
              localPart: 'chg'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGraphicFrameMoniker',
        typeName: 'CT_GraphicFrameMoniker',
        propertyInfos: [{
            name: 'id',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'creationId',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'creationId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGraphicFrameMonikerList',
        typeName: 'CT_GraphicFrameMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'CTGroupShapeChanges',
        typeName: 'CT_GroupShapeChanges',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: '.CTChangesData'
          }, {
            name: 'grpSpMkLst',
            required: true,
            typeInfo: '.CTGroupShapeMonikerList'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'chg',
            required: true,
            typeInfo: {
              type: 'list',
              baseTypeInfo: '.STGroupShapeChangeBit'
            },
            attributeName: {
              localPart: 'chg'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGroupShapeMoniker',
        typeName: 'CT_GroupShapeMoniker',
        propertyInfos: [{
            name: 'id',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'creationId',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'creationId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGroupShapeMonikerList',
        typeName: 'CT_GroupShapeMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'CTInkChanges',
        typeName: 'CT_InkChanges',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: '.CTChangesData'
          }, {
            name: 'inkMkLst',
            required: true,
            typeInfo: '.CTInkMonikerList'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'chg',
            required: true,
            typeInfo: {
              type: 'list',
              baseTypeInfo: '.STInkChangeBit'
            },
            attributeName: {
              localPart: 'chg'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTInkMoniker',
        typeName: 'CT_InkMoniker',
        propertyInfos: [{
            name: 'id',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'creationId',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'creationId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTInkMonikerList',
        typeName: 'CT_InkMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'CTPictureChanges',
        typeName: 'CT_PictureChanges',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: '.CTChangesData'
          }, {
            name: 'picMkLst',
            required: true,
            typeInfo: '.CTPictureMonikerList'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'chg',
            required: true,
            typeInfo: {
              type: 'list',
              baseTypeInfo: '.STPictureChangeBit'
            },
            attributeName: {
              localPart: 'chg'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureMoniker',
        typeName: 'CT_PictureMoniker',
        propertyInfos: [{
            name: 'id',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'creationId',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'creationId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureMonikerList',
        typeName: 'CT_PictureMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'CTShapeChanges',
        typeName: 'CT_ShapeChanges',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: '.CTChangesData'
          }, {
            name: 'spMkLst',
            required: true,
            typeInfo: '.CTShapeMonikerList'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'chg',
            required: true,
            typeInfo: {
              type: 'list',
              baseTypeInfo: '.STShapeChangeBit'
            },
            attributeName: {
              localPart: 'chg'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTShapeMoniker',
        typeName: 'CT_ShapeMoniker',
        propertyInfos: [{
            name: 'id',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'creationId',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'creationId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTShapeMonikerList',
        typeName: 'CT_ShapeMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STConnectorChangeBit',
        baseTypeInfo: 'Token',
        values: ['add', 'del', 'mod', 'ord', 'topLvl', 'modVis', 'replST', 'delST', 'replId']
      }, {
        type: 'enumInfo',
        localName: 'STGraphicFrameChangeBit',
        baseTypeInfo: 'Token',
        values: ['add', 'del', 'mod', 'ord', 'topLvl', 'modVis', 'replST', 'delST', 'replId', 'modGraphic']
      }, {
        type: 'enumInfo',
        localName: 'STGroupShapeChangeBit',
        baseTypeInfo: 'Token',
        values: ['add', 'del', 'mod', 'ord', 'topLvl', 'modVis', 'replST', 'delST', 'replId']
      }, {
        type: 'enumInfo',
        localName: 'STInkChangeBit',
        baseTypeInfo: 'Token',
        values: ['add', 'del', 'mod', 'ord', 'topLvl', 'modVis', 'replST', 'delST', 'replId', 'reco', 'modStrokes']
      }, {
        type: 'enumInfo',
        localName: 'STPictureChangeBit',
        baseTypeInfo: 'Token',
        values: ['add', 'del', 'mod', 'ord', 'topLvl', 'modVis', 'replST', 'delST', 'replId', 'modCrop']
      }, {
        type: 'enumInfo',
        localName: 'STShapeChangeBit',
        baseTypeInfo: 'Token',
        values: ['add', 'del', 'mod', 'ord', 'topLvl', 'modVis', 'replST', 'delST', 'replId', 'modCrop']
      }],
    elementInfos: [{
        typeInfo: '.CTConnectorMonikerList',
        elementName: 'cxnSpMkLst'
      }, {
        typeInfo: '.CTGraphicFrameMonikerList',
        elementName: 'graphicFrameMkLst'
      }, {
        typeInfo: '.CTGroupShapeMonikerList',
        elementName: 'grpMkLst'
      }, {
        typeInfo: '.CTInkMonikerList',
        elementName: 'inkMkLst'
      }, {
        typeInfo: '.CTPictureMonikerList',
        elementName: 'picMkLst'
      }, {
        typeInfo: '.CTShapeMonikerList',
        elementName: 'spMkLst'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command: org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command = org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command = org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command;
  }
}