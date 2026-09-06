var org_docx4j_com_microsoft_schemas_ink_x2010_main_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_ink_x2010_main = {
    name: 'org_docx4j_com_microsoft_schemas_ink_x2010_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/ink\/2010\/main',
    typeInfos: [{
        localName: 'CTCtxLink',
        typeName: 'CT_CtxLink',
        propertyInfos: [{
            name: 'direction',
            typeInfo: '.STDir',
            attributeName: {
              localPart: 'direction'
            },
            type: 'attribute'
          }, {
            name: 'ref',
            attributeName: {
              localPart: 'ref'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCtxNode',
        typeName: 'CT_CtxNode',
        propertyInfos: [{
            name: 'property',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTProperty'
          }, {
            name: 'sourceLink',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTCtxLink'
          }, {
            name: 'destinationLink',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTCtxLink'
          }, {
            name: 'id',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'type',
            required: true,
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: 'rotatedBoundingBox',
            typeInfo: {
              type: 'list'
            },
            attributeName: {
              localPart: 'rotatedBoundingBox'
            },
            type: 'attribute'
          }, {
            name: 'alignmentLevel',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'alignmentLevel'
            },
            type: 'attribute'
          }, {
            name: 'contentType',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'contentType'
            },
            type: 'attribute'
          }, {
            name: 'ascender',
            typeInfo: {
              type: 'list'
            },
            attributeName: {
              localPart: 'ascender'
            },
            type: 'attribute'
          }, {
            name: 'descender',
            typeInfo: {
              type: 'list'
            },
            attributeName: {
              localPart: 'descender'
            },
            type: 'attribute'
          }, {
            name: 'baseline',
            typeInfo: {
              type: 'list'
            },
            attributeName: {
              localPart: 'baseline'
            },
            type: 'attribute'
          }, {
            name: 'midline',
            typeInfo: {
              type: 'list'
            },
            attributeName: {
              localPart: 'midline'
            },
            type: 'attribute'
          }, {
            name: 'customRecognizerId',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'customRecognizerId'
            },
            type: 'attribute'
          }, {
            name: 'mathML',
            defaultValue: '',
            attributeName: {
              localPart: 'mathML'
            },
            type: 'attribute'
          }, {
            name: 'mathStruct',
            defaultValue: '',
            attributeName: {
              localPart: 'mathStruct'
            },
            type: 'attribute'
          }, {
            name: 'mathSymbol',
            defaultValue: '',
            attributeName: {
              localPart: 'mathSymbol'
            },
            type: 'attribute'
          }, {
            name: 'beginModifierType',
            defaultValue: '',
            attributeName: {
              localPart: 'beginModifierType'
            },
            type: 'attribute'
          }, {
            name: 'endModifierType',
            defaultValue: '',
            attributeName: {
              localPart: 'endModifierType'
            },
            type: 'attribute'
          }, {
            name: 'rotationAngle',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'rotationAngle'
            },
            type: 'attribute'
          }, {
            name: 'hotPoints',
            typeInfo: {
              type: 'list'
            },
            attributeName: {
              localPart: 'hotPoints'
            },
            type: 'attribute'
          }, {
            name: 'centroid',
            attributeName: {
              localPart: 'centroid'
            },
            type: 'attribute'
          }, {
            name: 'semanticType',
            defaultValue: 'none',
            attributeName: {
              localPart: 'semanticType'
            },
            type: 'attribute'
          }, {
            name: 'shapeName',
            defaultValue: '',
            attributeName: {
              localPart: 'shapeName'
            },
            type: 'attribute'
          }, {
            name: 'shapeGeometry',
            typeInfo: {
              type: 'list'
            },
            attributeName: {
              localPart: 'shapeGeometry'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTProperty',
        typeName: 'CT_Property',
        propertyInfos: [{
            name: 'value',
            typeInfo: 'HexBinary',
            type: 'value'
          }, {
            name: 'type',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STDir',
        values: ['to', 'from', 'with']
      }, {
        type: 'enumInfo',
        localName: 'STKnownCtxNodeType',
        values: ['root', 'unclassifiedInk', 'writingRegion', 'analysisHint', 'object', 'inkDrawing', 'image', 'paragraph', 'line', 'inkBullet', 'inkWord', 'textWord', 'customRecognizer', 'mathRegion', 'mathEquation', 'mathStruct', 'mathSymbol', 'mathIdentifier', 'mathOperator', 'mathNumber', 'nonInkDrawing', 'groupNode', 'mixedDrawing']
      }, {
        type: 'enumInfo',
        localName: 'STKnownSemanticType',
        values: ['none', 'underline', 'strikethrough', 'highlight', 'scratchOut', 'verticalRange', 'callout', 'enclosure', 'comment', 'container', 'connector']
      }],
    elementInfos: [{
        typeInfo: '.CTCtxNode',
        elementName: 'context'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_ink_x2010_main: org_docx4j_com_microsoft_schemas_ink_x2010_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_ink_x2010_main_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_ink_x2010_main_Module = org_docx4j_com_microsoft_schemas_ink_x2010_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_ink_x2010_main = org_docx4j_com_microsoft_schemas_ink_x2010_main_Module.org_docx4j_com_microsoft_schemas_ink_x2010_main;
  }
  else {
    var org_docx4j_com_microsoft_schemas_ink_x2010_main = org_docx4j_com_microsoft_schemas_ink_x2010_main_Module.org_docx4j_com_microsoft_schemas_ink_x2010_main;
  }
}