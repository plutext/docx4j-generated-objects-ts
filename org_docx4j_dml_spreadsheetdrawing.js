var org_docx4j_dml_spreadsheetdrawing_Module_Factory = function () {
  var org_docx4j_dml_spreadsheetdrawing = {
    name: 'org_docx4j_dml_spreadsheetdrawing',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/spreadsheetDrawing',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTAbsoluteAnchor',
        typeName: 'CT_AbsoluteAnchor',
        propertyInfos: [{
            name: 'pos',
            required: true,
            typeInfo: 'org_docx4j_dml.CTPoint2D'
          }, {
            name: 'ext',
            required: true,
            typeInfo: 'org_docx4j_dml.CTPositiveSize2D'
          }, {
            name: 'sp',
            required: true,
            typeInfo: '.CTShape'
          }, {
            name: 'grpSp',
            required: true,
            typeInfo: '.CTGroupShape'
          }, {
            name: 'graphicFrame',
            required: true,
            typeInfo: '.CTGraphicalObjectFrame'
          }, {
            name: 'cxnSp',
            required: true,
            typeInfo: '.CTConnector'
          }, {
            name: 'pic',
            required: true,
            typeInfo: '.CTPicture'
          }, {
            name: 'clientData',
            required: true,
            typeInfo: '.CTAnchorClientData'
          }]
      }, {
        localName: 'CTAnchorClientData',
        typeName: 'CT_AnchorClientData',
        propertyInfos: [{
            name: 'fLocksWithSheet',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'fLocksWithSheet'
            },
            type: 'attribute'
          }, {
            name: 'fPrintsWithSheet',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'fPrintsWithSheet'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTConnector',
        typeName: 'CT_Connector',
        propertyInfos: [{
            name: 'nvCxnSpPr',
            required: true,
            typeInfo: '.CTConnectorNonVisual'
          }, {
            name: 'spPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'style',
            typeInfo: 'org_docx4j_dml.CTShapeStyle'
          }, {
            name: 'macro',
            attributeName: {
              localPart: 'macro'
            },
            type: 'attribute'
          }, {
            name: 'fPublished',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'fPublished'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTConnectorNonVisual',
        typeName: 'CT_ConnectorNonVisual',
        propertyInfos: [{
            name: 'cNvPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualDrawingProps'
          }, {
            name: 'cNvCxnSpPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualConnectorProperties'
          }]
      }, {
        localName: 'CTDrawing',
        typeName: 'CT_Drawing',
        propertyInfos: [{
            name: 'egAnchor',
            minOccurs: 0,
            collection: true,
            elementTypeInfos: [{
                elementName: 'twoCellAnchor',
                typeInfo: '.CTTwoCellAnchor'
              }, {
                elementName: 'oneCellAnchor',
                typeInfo: '.CTOneCellAnchor'
              }, {
                elementName: 'absoluteAnchor',
                typeInfo: '.CTAbsoluteAnchor'
              }],
            type: 'elements'
          }]
      }, {
        localName: 'CTGraphicalObjectFrame',
        typeName: 'CT_GraphicalObjectFrame',
        propertyInfos: [{
            name: 'nvGraphicFramePr',
            required: true,
            typeInfo: '.CTGraphicalObjectFrameNonVisual'
          }, {
            name: 'xfrm',
            required: true,
            typeInfo: 'org_docx4j_dml.CTTransform2D'
          }, {
            name: 'graphic',
            required: true,
            elementName: {
              localPart: 'graphic',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/main'
            },
            typeInfo: 'org_docx4j_dml.Graphic'
          }, {
            name: 'macro',
            attributeName: {
              localPart: 'macro'
            },
            type: 'attribute'
          }, {
            name: 'fPublished',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'fPublished'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGraphicalObjectFrameNonVisual',
        typeName: 'CT_GraphicalObjectFrameNonVisual',
        propertyInfos: [{
            name: 'cNvPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualDrawingProps'
          }, {
            name: 'cNvGraphicFramePr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualGraphicFrameProperties'
          }]
      }, {
        localName: 'CTGroupShape',
        typeName: 'CT_GroupShape',
        propertyInfos: [{
            name: 'nvGrpSpPr',
            required: true,
            typeInfo: '.CTGroupShapeNonVisual'
          }, {
            name: 'grpSpPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTGroupShapeProperties'
          }, {
            name: 'spOrGrpSpOrGraphicFrame',
            minOccurs: 0,
            collection: true,
            elementTypeInfos: [{
                elementName: 'sp',
                typeInfo: '.CTShape'
              }, {
                elementName: 'grpSp',
                typeInfo: '.CTGroupShape'
              }, {
                elementName: 'graphicFrame',
                typeInfo: '.CTGraphicalObjectFrame'
              }, {
                elementName: 'cxnSp',
                typeInfo: '.CTConnector'
              }, {
                elementName: 'pic',
                typeInfo: '.CTPicture'
              }],
            type: 'elements'
          }]
      }, {
        localName: 'CTGroupShapeNonVisual',
        typeName: 'CT_GroupShapeNonVisual',
        propertyInfos: [{
            name: 'cNvPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualDrawingProps'
          }, {
            name: 'cNvGrpSpPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualGroupDrawingShapeProps'
          }]
      }, {
        localName: 'CTMarker',
        typeName: 'CT_Marker',
        propertyInfos: [{
            name: 'col',
            required: true,
            typeInfo: 'Int'
          }, {
            name: 'colOff',
            required: true,
            typeInfo: 'Long'
          }, {
            name: 'row',
            required: true,
            typeInfo: 'Int'
          }, {
            name: 'rowOff',
            required: true,
            typeInfo: 'Long'
          }]
      }, {
        localName: 'CTOneCellAnchor',
        typeName: 'CT_OneCellAnchor',
        propertyInfos: [{
            name: 'from',
            required: true,
            typeInfo: '.CTMarker'
          }, {
            name: 'ext',
            required: true,
            typeInfo: 'org_docx4j_dml.CTPositiveSize2D'
          }, {
            name: 'sp',
            required: true,
            typeInfo: '.CTShape'
          }, {
            name: 'grpSp',
            required: true,
            typeInfo: '.CTGroupShape'
          }, {
            name: 'graphicFrame',
            required: true,
            typeInfo: '.CTGraphicalObjectFrame'
          }, {
            name: 'cxnSp',
            required: true,
            typeInfo: '.CTConnector'
          }, {
            name: 'pic',
            required: true,
            typeInfo: '.CTPicture'
          }, {
            name: 'clientData',
            required: true,
            typeInfo: '.CTAnchorClientData'
          }]
      }, {
        localName: 'CTPicture',
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
          }, {
            name: 'style',
            typeInfo: 'org_docx4j_dml.CTShapeStyle'
          }, {
            name: 'macro',
            defaultValue: '',
            attributeName: {
              localPart: 'macro'
            },
            type: 'attribute'
          }, {
            name: 'fPublished',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'fPublished'
            },
            type: 'attribute'
          }]
      }, {
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
        localName: 'CTShape',
        typeName: 'CT_Shape',
        propertyInfos: [{
            name: 'nvSpPr',
            required: true,
            typeInfo: '.CTShapeNonVisual'
          }, {
            name: 'spPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'style',
            typeInfo: 'org_docx4j_dml.CTShapeStyle'
          }, {
            name: 'txBody',
            typeInfo: 'org_docx4j_dml.CTTextBody'
          }, {
            name: 'macro',
            attributeName: {
              localPart: 'macro'
            },
            type: 'attribute'
          }, {
            name: 'textlink',
            attributeName: {
              localPart: 'textlink'
            },
            type: 'attribute'
          }, {
            name: 'fLocksText',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'fLocksText'
            },
            type: 'attribute'
          }, {
            name: 'fPublished',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'fPublished'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTShapeNonVisual',
        typeName: 'CT_ShapeNonVisual',
        propertyInfos: [{
            name: 'cNvPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualDrawingProps'
          }, {
            name: 'cNvSpPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualDrawingShapeProps'
          }]
      }, {
        localName: 'CTTwoCellAnchor',
        typeName: 'CT_TwoCellAnchor',
        propertyInfos: [{
            name: 'from',
            required: true,
            typeInfo: '.CTMarker'
          }, {
            name: 'to',
            required: true,
            typeInfo: '.CTMarker'
          }, {
            name: 'sp',
            required: true,
            typeInfo: '.CTShape'
          }, {
            name: 'grpSp',
            required: true,
            typeInfo: '.CTGroupShape'
          }, {
            name: 'graphicFrame',
            required: true,
            typeInfo: '.CTGraphicalObjectFrame'
          }, {
            name: 'cxnSp',
            required: true,
            typeInfo: '.CTConnector'
          }, {
            name: 'pic',
            required: true,
            typeInfo: '.CTPicture'
          }, {
            name: 'clientData',
            required: true,
            typeInfo: '.CTAnchorClientData'
          }, {
            name: 'editAs',
            typeInfo: '.STEditAs',
            defaultValue: 'twoCell',
            attributeName: {
              localPart: 'editAs'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STEditAs',
        baseTypeInfo: 'Token',
        values: ['twoCell', 'oneCell', 'absolute']
      }],
    elementInfos: [{
        typeInfo: '.CTMarker',
        elementName: 'from'
      }, {
        typeInfo: '.CTMarker',
        elementName: 'to'
      }, {
        typeInfo: '.CTDrawing',
        elementName: 'wsDr'
      }]
  };
  return {
    org_docx4j_dml_spreadsheetdrawing: org_docx4j_dml_spreadsheetdrawing
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_dml_spreadsheetdrawing_Module_Factory);
}
else {
  var org_docx4j_dml_spreadsheetdrawing_Module = org_docx4j_dml_spreadsheetdrawing_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_dml_spreadsheetdrawing = org_docx4j_dml_spreadsheetdrawing_Module.org_docx4j_dml_spreadsheetdrawing;
  }
  else {
    var org_docx4j_dml_spreadsheetdrawing = org_docx4j_dml_spreadsheetdrawing_Module.org_docx4j_dml_spreadsheetdrawing;
  }
}