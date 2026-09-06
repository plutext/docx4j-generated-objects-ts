var org_docx4j_dml_wordprocessingDrawing_Module_Factory = function () {
  var org_docx4j_dml_wordprocessingDrawing = {
    name: 'org_docx4j_dml_wordprocessingDrawing',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/wordprocessingDrawing',
    dependencies: ['org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing', 'org_docx4j_dml'],
    typeInfos: [{
        localName: 'Anchor',
        typeName: 'CT_Anchor',
        propertyInfos: [{
            name: 'simplePos',
            required: true,
            typeInfo: 'org_docx4j_dml.CTPoint2D'
          }, {
            name: 'positionH',
            required: true,
            typeInfo: '.CTPosH'
          }, {
            name: 'positionV',
            required: true,
            typeInfo: '.CTPosV'
          }, {
            name: 'extent',
            required: true,
            typeInfo: 'org_docx4j_dml.CTPositiveSize2D'
          }, {
            name: 'effectExtent',
            typeInfo: '.CTEffectExtent'
          }, {
            name: 'wrapNone',
            required: true,
            typeInfo: '.CTWrapNone'
          }, {
            name: 'wrapSquare',
            required: true,
            typeInfo: '.CTWrapSquare'
          }, {
            name: 'wrapTight',
            required: true,
            typeInfo: '.CTWrapTight'
          }, {
            name: 'wrapThrough',
            required: true,
            typeInfo: '.CTWrapThrough'
          }, {
            name: 'wrapTopAndBottom',
            required: true,
            typeInfo: '.CTWrapTopBottom'
          }, {
            name: 'docPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualDrawingProps'
          }, {
            name: 'cNvGraphicFramePr',
            typeInfo: 'org_docx4j_dml.CTNonVisualGraphicFrameProperties'
          }, {
            name: 'graphic',
            required: true,
            elementName: {
              localPart: 'graphic',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/main'
            },
            typeInfo: 'org_docx4j_dml.Graphic'
          }, {
            name: 'sizeRelH',
            elementName: {
              localPart: 'sizeRelH',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordprocessingDrawing'
            },
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing.CTSizeRelH'
          }, {
            name: 'sizeRelV',
            elementName: {
              localPart: 'sizeRelV',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordprocessingDrawing'
            },
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_word_x2010_wordprocessingDrawing.CTSizeRelV'
          }, {
            name: 'distT',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distT'
            },
            type: 'attribute'
          }, {
            name: 'distB',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distB'
            },
            type: 'attribute'
          }, {
            name: 'distL',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distL'
            },
            type: 'attribute'
          }, {
            name: 'distR',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distR'
            },
            type: 'attribute'
          }, {
            name: 'simplePosAttr',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'simplePos'
            },
            type: 'attribute'
          }, {
            name: 'relativeHeight',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'relativeHeight'
            },
            type: 'attribute'
          }, {
            name: 'behindDoc',
            required: true,
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'behindDoc'
            },
            type: 'attribute'
          }, {
            name: 'locked',
            required: true,
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'locked'
            },
            type: 'attribute'
          }, {
            name: 'layoutInCell',
            required: true,
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'layoutInCell'
            },
            type: 'attribute'
          }, {
            name: 'hidden',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'hidden'
            },
            type: 'attribute'
          }, {
            name: 'allowOverlap',
            required: true,
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'allowOverlap'
            },
            type: 'attribute'
          }, {
            name: 'anchorId',
            typeInfo: 'HexBinary',
            attributeName: {
              localPart: 'anchorId',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordprocessingDrawing'
            },
            type: 'attribute'
          }, {
            name: 'editId',
            typeInfo: 'HexBinary',
            attributeName: {
              localPart: 'editId',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordprocessingDrawing'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTEffectExtent',
        typeName: 'CT_EffectExtent',
        propertyInfos: [{
            name: 'l',
            required: true,
            typeInfo: 'Long',
            attributeName: {
              localPart: 'l'
            },
            type: 'attribute'
          }, {
            name: 't',
            required: true,
            typeInfo: 'Long',
            attributeName: {
              localPart: 't'
            },
            type: 'attribute'
          }, {
            name: 'r',
            required: true,
            typeInfo: 'Long',
            attributeName: {
              localPart: 'r'
            },
            type: 'attribute'
          }, {
            name: 'b',
            required: true,
            typeInfo: 'Long',
            attributeName: {
              localPart: 'b'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPosH',
        typeName: 'CT_PosH',
        propertyInfos: [{
            name: 'align',
            required: true,
            typeInfo: '.STAlignH'
          }, {
            name: 'posOffset',
            required: true,
            typeInfo: 'Int'
          }, {
            name: 'relativeFrom',
            required: true,
            typeInfo: '.STRelFromH',
            attributeName: {
              localPart: 'relativeFrom'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPosV',
        typeName: 'CT_PosV',
        propertyInfos: [{
            name: 'align',
            required: true,
            typeInfo: '.STAlignV'
          }, {
            name: 'posOffset',
            required: true,
            typeInfo: 'Int'
          }, {
            name: 'relativeFrom',
            required: true,
            typeInfo: '.STRelFromV',
            attributeName: {
              localPart: 'relativeFrom'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTWrapNone',
        typeName: 'CT_WrapNone'
      }, {
        localName: 'CTWrapPath',
        typeName: 'CT_WrapPath',
        propertyInfos: [{
            name: 'start',
            required: true,
            typeInfo: 'org_docx4j_dml.CTPoint2D'
          }, {
            name: 'lineTo',
            required: true,
            minOccurs: 2,
            collection: true,
            typeInfo: 'org_docx4j_dml.CTPoint2D'
          }, {
            name: 'edited',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'edited'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTWrapSquare',
        typeName: 'CT_WrapSquare',
        propertyInfos: [{
            name: 'effectExtent',
            typeInfo: '.CTEffectExtent'
          }, {
            name: 'wrapText',
            required: true,
            typeInfo: '.STWrapText',
            attributeName: {
              localPart: 'wrapText'
            },
            type: 'attribute'
          }, {
            name: 'distT',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distT'
            },
            type: 'attribute'
          }, {
            name: 'distB',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distB'
            },
            type: 'attribute'
          }, {
            name: 'distL',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distL'
            },
            type: 'attribute'
          }, {
            name: 'distR',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distR'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTWrapThrough',
        typeName: 'CT_WrapThrough',
        propertyInfos: [{
            name: 'wrapPolygon',
            required: true,
            typeInfo: '.CTWrapPath'
          }, {
            name: 'wrapText',
            required: true,
            typeInfo: '.STWrapText',
            attributeName: {
              localPart: 'wrapText'
            },
            type: 'attribute'
          }, {
            name: 'distL',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distL'
            },
            type: 'attribute'
          }, {
            name: 'distR',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distR'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTWrapTight',
        typeName: 'CT_WrapTight',
        propertyInfos: [{
            name: 'wrapPolygon',
            required: true,
            typeInfo: '.CTWrapPath'
          }, {
            name: 'wrapText',
            required: true,
            typeInfo: '.STWrapText',
            attributeName: {
              localPart: 'wrapText'
            },
            type: 'attribute'
          }, {
            name: 'distL',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distL'
            },
            type: 'attribute'
          }, {
            name: 'distR',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distR'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTWrapTopBottom',
        typeName: 'CT_WrapTopBottom',
        propertyInfos: [{
            name: 'effectExtent',
            typeInfo: '.CTEffectExtent'
          }, {
            name: 'distT',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distT'
            },
            type: 'attribute'
          }, {
            name: 'distB',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distB'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'Inline',
        typeName: 'CT_Inline',
        propertyInfos: [{
            name: 'extent',
            required: true,
            typeInfo: 'org_docx4j_dml.CTPositiveSize2D'
          }, {
            name: 'effectExtent',
            typeInfo: '.CTEffectExtent'
          }, {
            name: 'docPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualDrawingProps'
          }, {
            name: 'cNvGraphicFramePr',
            typeInfo: 'org_docx4j_dml.CTNonVisualGraphicFrameProperties'
          }, {
            name: 'graphic',
            required: true,
            elementName: {
              localPart: 'graphic',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/main'
            },
            typeInfo: 'org_docx4j_dml.Graphic'
          }, {
            name: 'distT',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distT'
            },
            type: 'attribute'
          }, {
            name: 'distB',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distB'
            },
            type: 'attribute'
          }, {
            name: 'distL',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distL'
            },
            type: 'attribute'
          }, {
            name: 'distR',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'distR'
            },
            type: 'attribute'
          }, {
            name: 'anchorId',
            typeInfo: 'HexBinary',
            attributeName: {
              localPart: 'anchorId',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordprocessingDrawing'
            },
            type: 'attribute'
          }, {
            name: 'editId',
            typeInfo: 'HexBinary',
            attributeName: {
              localPart: 'editId',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordprocessingDrawing'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STAlignH',
        baseTypeInfo: 'Token',
        values: ['left', 'right', 'center', 'inside', 'outside']
      }, {
        type: 'enumInfo',
        localName: 'STAlignV',
        baseTypeInfo: 'Token',
        values: ['top', 'bottom', 'center', 'inside', 'outside']
      }, {
        type: 'enumInfo',
        localName: 'STRelFromH',
        baseTypeInfo: 'Token',
        values: ['margin', 'page', 'column', 'character', 'leftMargin', 'rightMargin', 'insideMargin', 'outsideMargin']
      }, {
        type: 'enumInfo',
        localName: 'STRelFromV',
        baseTypeInfo: 'Token',
        values: ['margin', 'page', 'paragraph', 'line', 'topMargin', 'bottomMargin', 'insideMargin', 'outsideMargin']
      }, {
        type: 'enumInfo',
        localName: 'STWrapText',
        baseTypeInfo: 'Token',
        values: ['bothSides', 'left', 'right', 'largest']
      }],
    elementInfos: [{
        typeInfo: '.Anchor',
        elementName: 'anchor'
      }, {
        typeInfo: '.Inline',
        elementName: 'inline'
      }]
  };
  return {
    org_docx4j_dml_wordprocessingDrawing: org_docx4j_dml_wordprocessingDrawing
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_dml_wordprocessingDrawing_Module_Factory);
}
else {
  var org_docx4j_dml_wordprocessingDrawing_Module = org_docx4j_dml_wordprocessingDrawing_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_dml_wordprocessingDrawing = org_docx4j_dml_wordprocessingDrawing_Module.org_docx4j_dml_wordprocessingDrawing;
  }
  else {
    var org_docx4j_dml_wordprocessingDrawing = org_docx4j_dml_wordprocessingDrawing_Module.org_docx4j_dml_wordprocessingDrawing;
  }
}