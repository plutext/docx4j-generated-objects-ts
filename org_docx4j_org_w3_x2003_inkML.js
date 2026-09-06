var org_docx4j_org_w3_x2003_inkML_Module_Factory = function () {
  var org_docx4j_org_w3_x2003_inkML = {
    name: 'org_docx4j_org_w3_x2003_inkML',
    defaultElementNamespaceURI: 'http:\/\/www.w3.org\/2003\/InkML',
    typeInfos: [{
        localName: 'ActiveAreaType',
        typeName: 'activeArea.type',
        propertyInfos: [{
            name: 'size',
            attributeName: {
              localPart: 'size'
            },
            type: 'attribute'
          }, {
            name: 'height',
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'height'
            },
            type: 'attribute'
          }, {
            name: 'width',
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'width'
            },
            type: 'attribute'
          }, {
            name: 'units',
            attributeName: {
              localPart: 'units'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'AffineType',
        typeName: 'affine.type',
        propertyInfos: [{
            name: 'value',
            type: 'value'
          }, {
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'AnnotationType',
        typeName: 'annotation.type',
        propertyInfos: [{
            name: 'otherAttributes',
            type: 'anyAttribute'
          }, {
            name: 'value',
            type: 'value'
          }, {
            name: 'type',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: 'encoding',
            attributeName: {
              localPart: 'encoding'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'AnnotationXMLType',
        typeName: 'annotationXML.type',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowTypedObject: false,
            mixed: false,
            type: 'anyElement'
          }, {
            name: 'type',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: 'encoding',
            attributeName: {
              localPart: 'encoding'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'BindType',
        typeName: 'bind.type',
        propertyInfos: [{
            name: 'source',
            attributeName: {
              localPart: 'source'
            },
            type: 'attribute'
          }, {
            name: 'target',
            attributeName: {
              localPart: 'target'
            },
            type: 'attribute'
          }, {
            name: 'column',
            attributeName: {
              localPart: 'column'
            },
            type: 'attribute'
          }, {
            name: 'variable',
            attributeName: {
              localPart: 'variable'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'BrushPropertyType',
        typeName: 'brushProperty.type',
        propertyInfos: [{
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'value',
            required: true,
            attributeName: {
              localPart: 'value'
            },
            type: 'attribute'
          }, {
            name: 'units',
            attributeName: {
              localPart: 'units'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'BrushType',
        typeName: 'brush.type',
        propertyInfos: [{
            name: 'brushProperty',
            minOccurs: 0,
            collection: true,
            typeInfo: '.BrushPropertyType'
          }, {
            name: 'annotation',
            minOccurs: 0,
            collection: true,
            typeInfo: '.AnnotationType'
          }, {
            name: 'annotationXML',
            minOccurs: 0,
            collection: true,
            typeInfo: '.AnnotationXMLType'
          }, {
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }, {
            name: 'brushRef',
            attributeName: {
              localPart: 'brushRef'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTMatrix',
        typeName: 'CT_Matrix',
        propertyInfos: [{
            name: 'value',
            type: 'value'
          }, {
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CanvasTransformType',
        typeName: 'canvasTransform.type',
        propertyInfos: [{
            name: 'content',
            required: true,
            maxOccurs: 2,
            collection: true,
            mixed: false,
            allowDom: false,
            elementName: 'mapping',
            typeInfo: '.MappingType',
            type: 'elementRef'
          }, {
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }, {
            name: 'invertible',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'invertible'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CanvasType',
        typeName: 'canvas.type',
        propertyInfos: [{
            name: 'traceFormat',
            typeInfo: '.TraceFormatType'
          }, {
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }, {
            name: 'traceFormatRef',
            attributeName: {
              localPart: 'traceFormatRef'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ChannelPropertiesType',
        typeName: 'channelProperties.type',
        propertyInfos: [{
            name: 'channelProperty',
            minOccurs: 0,
            collection: true,
            typeInfo: '.ChannelPropertiesType.ChannelProperty'
          }]
      }, {
        localName: 'ChannelPropertiesType.ChannelProperty',
        typeName: null,
        propertyInfos: [{
            name: 'channel',
            required: true,
            attributeName: {
              localPart: 'channel'
            },
            type: 'attribute'
          }, {
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'value',
            required: true,
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'value'
            },
            type: 'attribute'
          }, {
            name: 'units',
            attributeName: {
              localPart: 'units'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ChannelType',
        typeName: 'channel.type',
        propertyInfos: [{
            name: 'mapping',
            typeInfo: '.MappingType'
          }, {
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }, {
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'type',
            values: ['integer', 'decimal', 'double', 'boolean'],
            defaultValue: 'decimal',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: '_default',
            defaultValue: '0',
            attributeName: {
              localPart: 'default'
            },
            type: 'attribute'
          }, {
            name: 'min',
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'min'
            },
            type: 'attribute'
          }, {
            name: 'max',
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'max'
            },
            type: 'attribute'
          }, {
            name: 'orientation',
            values: ['+ve', '-ve'],
            defaultValue: '+ve',
            attributeName: {
              localPart: 'orientation'
            },
            type: 'attribute'
          }, {
            name: 'respectTo',
            attributeName: {
              localPart: 'respectTo'
            },
            type: 'attribute'
          }, {
            name: 'units',
            attributeName: {
              localPart: 'units'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'ContextType',
        typeName: 'context.type',
        propertyInfos: [{
            name: 'canvas',
            typeInfo: '.CanvasType'
          }, {
            name: 'canvasTransform',
            typeInfo: '.CanvasTransformType'
          }, {
            name: 'traceFormat',
            typeInfo: '.TraceFormatType'
          }, {
            name: 'inkSource',
            typeInfo: '.InkSourceType'
          }, {
            name: 'brush',
            typeInfo: '.BrushType'
          }, {
            name: 'timestamp',
            typeInfo: '.TimestampType'
          }, {
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }, {
            name: 'contextRef',
            attributeName: {
              localPart: 'contextRef'
            },
            type: 'attribute'
          }, {
            name: 'canvasRef',
            attributeName: {
              localPart: 'canvasRef'
            },
            type: 'attribute'
          }, {
            name: 'canvasTransformRef',
            attributeName: {
              localPart: 'canvasTransformRef'
            },
            type: 'attribute'
          }, {
            name: 'traceFormatRef',
            attributeName: {
              localPart: 'traceFormatRef'
            },
            type: 'attribute'
          }, {
            name: 'inkSourceRef',
            attributeName: {
              localPart: 'inkSourceRef'
            },
            type: 'attribute'
          }, {
            name: 'brushRef',
            attributeName: {
              localPart: 'brushRef'
            },
            type: 'attribute'
          }, {
            name: 'timestampRef',
            attributeName: {
              localPart: 'timestampRef'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'DefinitionsType',
        typeName: 'definitions.type',
        propertyInfos: [{
            name: 'brushOrCanvasOrCanvasTransform',
            required: true,
            collection: true,
            elementTypeInfos: [{
                elementName: 'brush',
                typeInfo: '.BrushType'
              }, {
                elementName: 'canvas',
                typeInfo: '.CanvasType'
              }, {
                elementName: 'canvasTransform',
                typeInfo: '.CanvasTransformType'
              }, {
                elementName: 'context',
                typeInfo: '.ContextType'
              }, {
                elementName: 'inkSource',
                typeInfo: '.InkSourceType'
              }, {
                elementName: 'mapping',
                typeInfo: '.MappingType'
              }, {
                elementName: 'timestamp',
                typeInfo: '.TimestampType'
              }, {
                elementName: 'trace',
                typeInfo: '.TraceType'
              }, {
                elementName: 'traceFormat',
                typeInfo: '.TraceFormatType'
              }, {
                elementName: 'traceGroup',
                typeInfo: '.TraceGroupType'
              }, {
                elementName: 'traceView',
                typeInfo: '.TraceViewType'
              }],
            type: 'elements'
          }]
      }, {
        localName: 'InkSourceType',
        typeName: 'inkSource.type',
        propertyInfos: [{
            name: 'traceFormat',
            required: true,
            typeInfo: '.TraceFormatType'
          }, {
            name: 'sampleRate',
            typeInfo: '.SampleRateType'
          }, {
            name: 'latency',
            typeInfo: '.LatencyType'
          }, {
            name: 'activeArea',
            typeInfo: '.ActiveAreaType'
          }, {
            name: 'sourceProperty',
            minOccurs: 0,
            collection: true,
            typeInfo: '.SourcePropertyType'
          }, {
            name: 'channelProperties',
            typeInfo: '.ChannelPropertiesType'
          }, {
            name: 'id',
            required: true,
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }, {
            name: 'manufacturer',
            attributeName: {
              localPart: 'manufacturer'
            },
            type: 'attribute'
          }, {
            name: 'model',
            attributeName: {
              localPart: 'model'
            },
            type: 'attribute'
          }, {
            name: 'serialNo',
            attributeName: {
              localPart: 'serialNo'
            },
            type: 'attribute'
          }, {
            name: 'specificationRef',
            attributeName: {
              localPart: 'specificationRef'
            },
            type: 'attribute'
          }, {
            name: 'description',
            attributeName: {
              localPart: 'description'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'InkType',
        typeName: 'ink.type',
        propertyInfos: [{
            name: 'definitionsOrContextOrTrace',
            minOccurs: 0,
            collection: true,
            elementTypeInfos: [{
                elementName: 'definitions',
                typeInfo: '.DefinitionsType'
              }, {
                elementName: 'context',
                typeInfo: '.ContextType'
              }, {
                elementName: 'trace',
                typeInfo: '.TraceType'
              }, {
                elementName: 'traceGroup',
                typeInfo: '.TraceGroupType'
              }, {
                elementName: 'traceView',
                typeInfo: '.TraceViewType'
              }, {
                elementName: 'annotation',
                typeInfo: '.AnnotationType'
              }, {
                elementName: 'annotationXML',
                typeInfo: '.AnnotationXMLType'
              }],
            type: 'elements'
          }, {
            name: 'documentID',
            attributeName: {
              localPart: 'documentID'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'IntermittentChannelsType',
        typeName: 'intermittentChannels.type',
        propertyInfos: [{
            name: 'channel',
            required: true,
            collection: true,
            typeInfo: '.ChannelType'
          }]
      }, {
        localName: 'LatencyType',
        typeName: 'latency.type',
        propertyInfos: [{
            name: 'value',
            required: true,
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'value'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'MappingType',
        typeName: 'mapping.type',
        propertyInfos: [{
            name: 'content',
            collection: true,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'bind',
                typeInfo: '.BindType'
              }, {
                elementName: 'table',
                typeInfo: '.TableType'
              }, {
                elementName: 'affine',
                typeInfo: '.AffineType'
              }, {
                elementName: {
                  localPart: 'math',
                  namespaceURI: 'http:\/\/www.w3.org\/1998\/Math\/MathML'
                },
                typeInfo: 'org_docx4j_org_w3_x1998_math_mathML.MathType'
              }, {
                elementName: 'mapping',
                typeInfo: '.MappingType'
              }],
            type: 'elementRefs'
          }, {
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }, {
            name: 'type',
            values: ['identity', 'product', 'table', 'affine', 'mathml', 'unknown'],
            defaultValue: 'unknown',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: 'mappingRef',
            attributeName: {
              localPart: 'mappingRef'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'SampleRateType',
        typeName: 'sampleRate.type',
        propertyInfos: [{
            name: 'uniform',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'uniform'
            },
            type: 'attribute'
          }, {
            name: 'value',
            required: true,
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'value'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'SourcePropertyType',
        typeName: 'sourceProperty.type',
        propertyInfos: [{
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'value',
            required: true,
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'value'
            },
            type: 'attribute'
          }, {
            name: 'units',
            attributeName: {
              localPart: 'units'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'TableType',
        typeName: 'table.type',
        propertyInfos: [{
            name: 'value',
            type: 'value'
          }, {
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }, {
            name: 'apply',
            values: ['absolute', 'relative'],
            defaultValue: 'absolute',
            attributeName: {
              localPart: 'apply'
            },
            type: 'attribute'
          }, {
            name: 'interpolation',
            values: ['floor', 'middle', 'ceiling', 'linear', 'cubic'],
            defaultValue: 'linear',
            attributeName: {
              localPart: 'interpolation'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'TimestampType',
        typeName: 'timestamp.type',
        propertyInfos: [{
            name: 'id',
            required: true,
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }, {
            name: 'time',
            typeInfo: 'UnsignedLong',
            attributeName: {
              localPart: 'time'
            },
            type: 'attribute'
          }, {
            name: 'timestampRef',
            attributeName: {
              localPart: 'timestampRef'
            },
            type: 'attribute'
          }, {
            name: 'timeString',
            typeInfo: 'DateTime',
            attributeName: {
              localPart: 'timeString'
            },
            type: 'attribute'
          }, {
            name: 'timeOffset',
            typeInfo: 'Decimal',
            defaultValue: 0,
            attributeName: {
              localPart: 'timeOffset'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'TraceFormatType',
        typeName: 'traceFormat.type',
        propertyInfos: [{
            name: 'channel',
            minOccurs: 0,
            collection: true,
            typeInfo: '.ChannelType'
          }, {
            name: 'intermittentChannels',
            typeInfo: '.IntermittentChannelsType'
          }, {
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'TraceGroupType',
        typeName: 'traceGroup.type',
        propertyInfos: [{
            name: 'traceAndTraceGroupAndTraceView',
            minOccurs: 0,
            collection: true,
            elementTypeInfos: [{
                elementName: 'trace',
                typeInfo: '.TraceType'
              }, {
                elementName: 'traceGroup',
                typeInfo: '.TraceGroupType'
              }, {
                elementName: 'traceView',
                typeInfo: '.TraceViewType'
              }, {
                elementName: 'annotation',
                typeInfo: '.AnnotationType'
              }, {
                elementName: 'annotationXML',
                typeInfo: '.AnnotationXMLType'
              }],
            type: 'elements'
          }, {
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }, {
            name: 'contextRef',
            attributeName: {
              localPart: 'contextRef'
            },
            type: 'attribute'
          }, {
            name: 'brushRef',
            attributeName: {
              localPart: 'brushRef'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'TraceType',
        typeName: 'trace.type',
        propertyInfos: [{
            name: 'value',
            type: 'value'
          }, {
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }, {
            name: 'type',
            values: ['penDown', 'penUp', 'indeterminate'],
            defaultValue: 'penDown',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: 'continuation',
            values: ['begin', 'end', 'middle'],
            defaultValue: 'begin',
            attributeName: {
              localPart: 'continuation'
            },
            type: 'attribute'
          }, {
            name: 'priorRef',
            attributeName: {
              localPart: 'priorRef'
            },
            type: 'attribute'
          }, {
            name: 'contextRef',
            attributeName: {
              localPart: 'contextRef'
            },
            type: 'attribute'
          }, {
            name: 'brushRef',
            attributeName: {
              localPart: 'brushRef'
            },
            type: 'attribute'
          }, {
            name: 'duration',
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'duration'
            },
            type: 'attribute'
          }, {
            name: 'timeOffset',
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'timeOffset'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'TraceViewType',
        typeName: 'traceView.type',
        propertyInfos: [{
            name: 'id',
            typeInfo: 'ID',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }, {
            name: 'traceDataRef',
            attributeName: {
              localPart: 'traceDataRef'
            },
            type: 'attribute'
          }, {
            name: 'from',
            attributeName: {
              localPart: 'from'
            },
            type: 'attribute'
          }, {
            name: 'to',
            attributeName: {
              localPart: 'to'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'BooleanStrType',
        values: ['F', 'T']
      }, {
        type: 'enumInfo',
        localName: 'BrushRasterOpType',
        values: ['noOperation', 'copyPen', 'maskPen']
      }, {
        type: 'enumInfo',
        localName: 'BrushTipType',
        values: ['ellipse', 'rectangle', 'drop']
      }, {
        type: 'enumInfo',
        localName: 'StandardBrushPropertyNameType',
        values: ['width', 'height', 'color', 'transparency', 'tip', 'rasterOp', 'antiAliased', 'fitToCurve', 'ignorePressure']
      }, {
        type: 'enumInfo',
        localName: 'StandardChannelNameType',
        values: ['X', 'Y', 'Z', 'F', 'S', 'B1', 'B2', 'B3', 'B4', 'OTx', 'OTy', 'OA', 'OE', 'OR', 'C', 'CR', 'CG', 'CB', 'CC', 'CM', 'CY', 'CK', 'A', 'W', 'BW', 'BH', 'T']
      }, {
        type: 'enumInfo',
        localName: 'StandardChannelPropertyNameType',
        values: ['threshold', 'resolution', 'quantization', 'noise', 'accuracy', 'crossCoupling', 'skew', 'minBandwidth', 'peakRate', 'distortion']
      }],
    elementInfos: [{
        typeInfo: '.AffineType',
        elementName: 'affine',
        scope: '.MappingType'
      }, {
        typeInfo: '.BindType',
        elementName: 'bind',
        scope: '.MappingType'
      }, {
        typeInfo: '.DefinitionsType',
        elementName: 'definitions'
      }, {
        typeInfo: '.InkType',
        elementName: 'ink'
      }, {
        typeInfo: '.MappingType',
        elementName: 'mapping',
        scope: '.CanvasTransformType'
      }, {
        typeInfo: '.MappingType',
        elementName: 'mapping',
        scope: '.MappingType'
      }, {
        typeInfo: '.TableType',
        elementName: 'table',
        scope: '.MappingType'
      }, {
        typeInfo: '.TraceType',
        elementName: 'trace'
      }, {
        typeInfo: '.TraceViewType',
        elementName: 'traceView'
      }]
  };
  return {
    org_docx4j_org_w3_x2003_inkML: org_docx4j_org_w3_x2003_inkML
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_org_w3_x2003_inkML_Module_Factory);
}
else {
  var org_docx4j_org_w3_x2003_inkML_Module = org_docx4j_org_w3_x2003_inkML_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_org_w3_x2003_inkML = org_docx4j_org_w3_x2003_inkML_Module.org_docx4j_org_w3_x2003_inkML;
  }
  else {
    var org_docx4j_org_w3_x2003_inkML = org_docx4j_org_w3_x2003_inkML_Module.org_docx4j_org_w3_x2003_inkML;
  }
}