var org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2014\/chartex',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTAddress',
        typeName: 'CT_Address',
        propertyInfos: [{
            name: 'address1',
            attributeName: {
              localPart: 'address1'
            },
            type: 'attribute'
          }, {
            name: 'adminDistrict1',
            attributeName: {
              localPart: 'adminDistrict1'
            },
            type: 'attribute'
          }, {
            name: 'adminDistrict2',
            attributeName: {
              localPart: 'adminDistrict2'
            },
            type: 'attribute'
          }, {
            name: 'countryRegion',
            attributeName: {
              localPart: 'countryRegion'
            },
            type: 'attribute'
          }, {
            name: 'isoCountryCode',
            attributeName: {
              localPart: 'isoCountryCode'
            },
            type: 'attribute'
          }, {
            name: 'locality',
            attributeName: {
              localPart: 'locality'
            },
            type: 'attribute'
          }, {
            name: 'postalCode',
            attributeName: {
              localPart: 'postalCode'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTAggregation',
        typeName: 'CT_Aggregation'
      }, {
        localName: 'CTAxis',
        typeName: 'CT_Axis',
        propertyInfos: [{
            name: 'catScaling',
            typeInfo: '.CTCategoryAxisScaling'
          }, {
            name: 'valScaling',
            typeInfo: '.CTValueAxisScaling'
          }, {
            name: 'title',
            typeInfo: '.CTAxisTitle'
          }, {
            name: 'units',
            typeInfo: '.CTAxisUnits'
          }, {
            name: 'majorGridlines',
            typeInfo: '.CTGridlines'
          }, {
            name: 'minorGridlines',
            typeInfo: '.CTGridlines'
          }, {
            name: 'majorTickMarks',
            typeInfo: '.CTTickMarks'
          }, {
            name: 'minorTickMarks',
            typeInfo: '.CTTickMarks'
          }, {
            name: 'tickLabels',
            typeInfo: '.CTTickLabels'
          }, {
            name: 'numFmt',
            typeInfo: '.CTNumberFormat'
          }, {
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'txPr',
            typeInfo: 'org_docx4j_dml.CTTextBody'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }, {
            name: 'hidden',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'hidden'
            },
            type: 'attribute'
          }, {
            name: 'id',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTAxisId',
        typeName: 'CT_AxisId',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTAxisTitle',
        typeName: 'CT_AxisTitle',
        propertyInfos: [{
            name: 'tx',
            typeInfo: '.CTText'
          }, {
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'txPr',
            typeInfo: 'org_docx4j_dml.CTTextBody'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }]
      }, {
        localName: 'CTAxisUnits',
        typeName: 'CT_AxisUnits',
        propertyInfos: [{
            name: 'unitsLabel',
            typeInfo: '.CTAxisUnitsLabel'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }, {
            name: 'unit',
            typeInfo: '.STAxisUnit',
            attributeName: {
              localPart: 'unit'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTAxisUnitsLabel',
        typeName: 'CT_AxisUnitsLabel',
        propertyInfos: [{
            name: 'tx',
            typeInfo: '.CTText'
          }, {
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'txPr',
            typeInfo: 'org_docx4j_dml.CTTextBody'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }]
      }, {
        localName: 'CTBinning',
        typeName: 'CT_Binning',
        propertyInfos: [{
            name: 'binSize',
            typeInfo: 'Double'
          }, {
            name: 'binCount',
            typeInfo: 'UnsignedInt'
          }, {
            name: 'intervalClosed',
            typeInfo: '.STIntervalClosedSide',
            attributeName: {
              localPart: 'intervalClosed'
            },
            type: 'attribute'
          }, {
            name: 'overflow',
            attributeName: {
              localPart: 'overflow'
            },
            type: 'attribute'
          }, {
            name: 'underflow',
            attributeName: {
              localPart: 'underflow'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCategoryAxisScaling',
        typeName: 'CT_CategoryAxisScaling',
        propertyInfos: [{
            name: 'gapWidth',
            attributeName: {
              localPart: 'gapWidth'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTChart',
        typeName: 'CT_Chart',
        propertyInfos: [{
            name: 'title',
            typeInfo: '.CTChartTitle'
          }, {
            name: 'plotArea',
            required: true,
            typeInfo: '.CTPlotArea'
          }, {
            name: 'legend',
            typeInfo: '.CTLegend'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }]
      }, {
        localName: 'CTChartData',
        typeName: 'CT_ChartData',
        propertyInfos: [{
            name: 'externalData',
            typeInfo: '.CTExternalData'
          }, {
            name: 'data',
            required: true,
            collection: true,
            typeInfo: '.CTData'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }]
      }, {
        localName: 'CTChartSpace',
        typeName: 'CT_ChartSpace',
        propertyInfos: [{
            name: 'chartData',
            required: true,
            typeInfo: '.CTChartData'
          }, {
            name: 'chart',
            required: true,
            typeInfo: '.CTChart'
          }, {
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'txPr',
            typeInfo: 'org_docx4j_dml.CTTextBody'
          }, {
            name: 'clrMapOvr',
            typeInfo: 'org_docx4j_dml.CTColorMapping'
          }, {
            name: 'fmtOvrs',
            typeInfo: '.CTFormatOverrides'
          }, {
            name: 'printSettings',
            typeInfo: '.CTPrintSettings'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }]
      }, {
        localName: 'CTChartTitle',
        typeName: 'CT_ChartTitle',
        propertyInfos: [{
            name: 'tx',
            typeInfo: '.CTText'
          }, {
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'txPr',
            typeInfo: 'org_docx4j_dml.CTTextBody'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }, {
            name: 'align',
            typeInfo: '.STPosAlign',
            defaultValue: 'ctr',
            attributeName: {
              localPart: 'align'
            },
            type: 'attribute'
          }, {
            name: 'overlay',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'overlay'
            },
            type: 'attribute'
          }, {
            name: 'pos',
            typeInfo: '.STSidePos',
            defaultValue: 't',
            attributeName: {
              localPart: 'pos'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTClear',
        typeName: 'CT_Clear',
        propertyInfos: [{
            name: 'geoLocationQueryResults',
            typeInfo: '.CTGeoLocationQueryResults'
          }, {
            name: 'geoDataEntityQueryResults',
            typeInfo: '.CTGeoDataEntityQueryResults'
          }, {
            name: 'geoDataPointToEntityQueryResults',
            typeInfo: '.CTGeoDataPointToEntityQueryResults'
          }, {
            name: 'geoChildEntitiesQueryResults',
            typeInfo: '.CTGeoChildEntitiesQueryResults'
          }, {
            name: 'geoParentEntitiesQueryResults',
            typeInfo: '.CTGeoParentEntitiesQueryResults'
          }]
      }, {
        localName: 'CTCopyrights',
        typeName: 'CT_Copyrights',
        propertyInfos: [{
            name: 'copyright',
            minOccurs: 0,
            collection: true
          }]
      }, {
        localName: 'CTData',
        typeName: 'CT_Data',
        propertyInfos: [{
            name: 'numDimOrStrDim',
            required: true,
            collection: true,
            elementTypeInfos: [{
                elementName: 'numDim',
                typeInfo: '.CTNumericDimension'
              }, {
                elementName: 'strDim',
                typeInfo: '.CTStringDimension'
              }],
            type: 'elements'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }, {
            name: 'id',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDataId',
        typeName: 'CT_DataId',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDataLabel',
        typeName: 'CT_DataLabel',
        propertyInfos: [{
            name: 'numFmt',
            typeInfo: '.CTNumberFormat'
          }, {
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'txPr',
            typeInfo: 'org_docx4j_dml.CTTextBody'
          }, {
            name: 'visibility',
            typeInfo: '.CTDataLabelVisibilities'
          }, {
            name: 'separator'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }, {
            name: 'idx',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'idx'
            },
            type: 'attribute'
          }, {
            name: 'pos',
            typeInfo: '.STDataLabelPos',
            attributeName: {
              localPart: 'pos'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDataLabelHidden',
        typeName: 'CT_DataLabelHidden',
        propertyInfos: [{
            name: 'idx',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'idx'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDataLabelVisibilities',
        typeName: 'CT_DataLabelVisibilities',
        propertyInfos: [{
            name: 'categoryName',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'categoryName'
            },
            type: 'attribute'
          }, {
            name: 'seriesName',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'seriesName'
            },
            type: 'attribute'
          }, {
            name: 'value',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'value'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDataLabels',
        typeName: 'CT_DataLabels',
        propertyInfos: [{
            name: 'numFmt',
            typeInfo: '.CTNumberFormat'
          }, {
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'txPr',
            typeInfo: 'org_docx4j_dml.CTTextBody'
          }, {
            name: 'visibility',
            typeInfo: '.CTDataLabelVisibilities'
          }, {
            name: 'separator'
          }, {
            name: 'dataLabel',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTDataLabel'
          }, {
            name: 'dataLabelHidden',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTDataLabelHidden'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }, {
            name: 'pos',
            typeInfo: '.STDataLabelPos',
            attributeName: {
              localPart: 'pos'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDataPoint',
        typeName: 'CT_DataPoint',
        propertyInfos: [{
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }, {
            name: 'idx',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'idx'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTExtension',
        typeName: 'CT_Extension',
        propertyInfos: [{
            name: 'any',
            required: true,
            mixed: false,
            type: 'anyElement'
          }, {
            name: 'uri',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'uri'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTExtensionList',
        typeName: 'CT_ExtensionList',
        propertyInfos: [{
            name: 'ext',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTExtension'
          }]
      }, {
        localName: 'CTExternalData',
        typeName: 'CT_ExternalData',
        propertyInfos: [{
            name: 'autoUpdate',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'autoUpdate'
            },
            type: 'attribute'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTExtremeValueColorPosition',
        typeName: 'CT_ExtremeValueColorPosition'
      }, {
        localName: 'CTFormatOverride',
        typeName: 'CT_FormatOverride',
        propertyInfos: [{
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }, {
            name: 'idx',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'idx'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTFormatOverrides',
        typeName: 'CT_FormatOverrides',
        propertyInfos: [{
            name: 'fmtOvr',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTFormatOverride'
          }]
      }, {
        localName: 'CTFormula',
        typeName: 'CT_Formula',
        propertyInfos: [{
            name: 'value',
            type: 'value'
          }, {
            name: 'dir',
            typeInfo: '.STFormulaDirection',
            defaultValue: 'col',
            attributeName: {
              localPart: 'dir'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoCache',
        typeName: 'CT_GeoCache',
        propertyInfos: [{
            name: 'binaryOrClear',
            minOccurs: 0,
            collection: true,
            elementTypeInfos: [{
                elementName: 'binary',
                typeInfo: 'Base64Binary'
              }, {
                elementName: 'clear',
                typeInfo: '.CTClear'
              }],
            type: 'elements'
          }, {
            name: 'provider',
            required: true,
            attributeName: {
              localPart: 'provider'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoChildEntities',
        typeName: 'CT_GeoChildEntities',
        propertyInfos: [{
            name: 'geoHierarchyEntity',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTGeoHierarchyEntity'
          }]
      }, {
        localName: 'CTGeoChildEntitiesQuery',
        typeName: 'CT_GeoChildEntitiesQuery',
        propertyInfos: [{
            name: 'geoChildTypes',
            typeInfo: '.CTGeoChildTypes'
          }, {
            name: 'entityId',
            required: true,
            attributeName: {
              localPart: 'entityId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoChildEntitiesQueryResult',
        typeName: 'CT_GeoChildEntitiesQueryResult',
        propertyInfos: [{
            name: 'geoChildEntitiesQuery',
            typeInfo: '.CTGeoChildEntitiesQuery'
          }, {
            name: 'geoChildEntities',
            typeInfo: '.CTGeoChildEntities'
          }]
      }, {
        localName: 'CTGeoChildEntitiesQueryResults',
        typeName: 'CT_GeoChildEntitiesQueryResults',
        propertyInfos: [{
            name: 'geoChildEntitiesQueryResult',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTGeoChildEntitiesQueryResult'
          }]
      }, {
        localName: 'CTGeoChildTypes',
        typeName: 'CT_GeoChildTypes',
        propertyInfos: [{
            name: 'entityType',
            minOccurs: 0,
            collection: true,
            typeInfo: '.STEntityType'
          }]
      }, {
        localName: 'CTGeoData',
        typeName: 'CT_GeoData',
        propertyInfos: [{
            name: 'geoPolygons',
            typeInfo: '.CTGeoPolygons'
          }, {
            name: 'copyrights',
            typeInfo: '.CTCopyrights'
          }, {
            name: 'east',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'east'
            },
            type: 'attribute'
          }, {
            name: 'entityId',
            required: true,
            attributeName: {
              localPart: 'entityId'
            },
            type: 'attribute'
          }, {
            name: 'entityName',
            required: true,
            attributeName: {
              localPart: 'entityName'
            },
            type: 'attribute'
          }, {
            name: 'north',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'north'
            },
            type: 'attribute'
          }, {
            name: 'south',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'south'
            },
            type: 'attribute'
          }, {
            name: 'west',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'west'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoDataEntityQuery',
        typeName: 'CT_GeoDataEntityQuery',
        propertyInfos: [{
            name: 'entityId',
            required: true,
            attributeName: {
              localPart: 'entityId'
            },
            type: 'attribute'
          }, {
            name: 'entityType',
            required: true,
            typeInfo: '.STEntityType',
            attributeName: {
              localPart: 'entityType'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoDataEntityQueryResult',
        typeName: 'CT_GeoDataEntityQueryResult',
        propertyInfos: [{
            name: 'geoDataEntityQuery',
            typeInfo: '.CTGeoDataEntityQuery'
          }, {
            name: 'geoData',
            typeInfo: '.CTGeoData'
          }]
      }, {
        localName: 'CTGeoDataEntityQueryResults',
        typeName: 'CT_GeoDataEntityQueryResults',
        propertyInfos: [{
            name: 'geoDataEntityQueryResult',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTGeoDataEntityQueryResult'
          }]
      }, {
        localName: 'CTGeoDataPointQuery',
        typeName: 'CT_GeoDataPointQuery',
        propertyInfos: [{
            name: 'entityType',
            required: true,
            typeInfo: '.STEntityType',
            attributeName: {
              localPart: 'entityType'
            },
            type: 'attribute'
          }, {
            name: 'latitude',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'latitude'
            },
            type: 'attribute'
          }, {
            name: 'longitude',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'longitude'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoDataPointToEntityQuery',
        typeName: 'CT_GeoDataPointToEntityQuery',
        propertyInfos: [{
            name: 'entityId',
            required: true,
            attributeName: {
              localPart: 'entityId'
            },
            type: 'attribute'
          }, {
            name: 'entityType',
            required: true,
            typeInfo: '.STEntityType',
            attributeName: {
              localPart: 'entityType'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoDataPointToEntityQueryResult',
        typeName: 'CT_GeoDataPointToEntityQueryResult',
        propertyInfos: [{
            name: 'geoDataPointQuery',
            typeInfo: '.CTGeoDataPointQuery'
          }, {
            name: 'geoDataPointToEntityQuery',
            typeInfo: '.CTGeoDataPointToEntityQuery'
          }]
      }, {
        localName: 'CTGeoDataPointToEntityQueryResults',
        typeName: 'CT_GeoDataPointToEntityQueryResults',
        propertyInfos: [{
            name: 'geoDataPointToEntityQueryResult',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTGeoDataPointToEntityQueryResult'
          }]
      }, {
        localName: 'CTGeoEntity',
        typeName: 'CT_GeoEntity',
        propertyInfos: [{
            name: 'entityName',
            required: true,
            attributeName: {
              localPart: 'entityName'
            },
            type: 'attribute'
          }, {
            name: 'entityType',
            required: true,
            typeInfo: '.STEntityType',
            attributeName: {
              localPart: 'entityType'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoHierarchyEntity',
        typeName: 'CT_GeoHierarchyEntity',
        propertyInfos: [{
            name: 'entityId',
            required: true,
            attributeName: {
              localPart: 'entityId'
            },
            type: 'attribute'
          }, {
            name: 'entityName',
            required: true,
            attributeName: {
              localPart: 'entityName'
            },
            type: 'attribute'
          }, {
            name: 'entityType',
            required: true,
            typeInfo: '.STEntityType',
            attributeName: {
              localPart: 'entityType'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoLocation',
        typeName: 'CT_GeoLocation',
        propertyInfos: [{
            name: 'address',
            typeInfo: '.CTAddress'
          }, {
            name: 'entityName',
            required: true,
            attributeName: {
              localPart: 'entityName'
            },
            type: 'attribute'
          }, {
            name: 'entityType',
            required: true,
            typeInfo: '.STEntityType',
            attributeName: {
              localPart: 'entityType'
            },
            type: 'attribute'
          }, {
            name: 'latitude',
            typeInfo: 'Double',
            attributeName: {
              localPart: 'latitude'
            },
            type: 'attribute'
          }, {
            name: 'longitude',
            typeInfo: 'Double',
            attributeName: {
              localPart: 'longitude'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoLocationQuery',
        typeName: 'CT_GeoLocationQuery',
        propertyInfos: [{
            name: 'adminDistrict1',
            attributeName: {
              localPart: 'adminDistrict1'
            },
            type: 'attribute'
          }, {
            name: 'adminDistrict2',
            attributeName: {
              localPart: 'adminDistrict2'
            },
            type: 'attribute'
          }, {
            name: 'countryRegion',
            attributeName: {
              localPart: 'countryRegion'
            },
            type: 'attribute'
          }, {
            name: 'entityType',
            required: true,
            typeInfo: '.STEntityType',
            attributeName: {
              localPart: 'entityType'
            },
            type: 'attribute'
          }, {
            name: 'postalCode',
            attributeName: {
              localPart: 'postalCode'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoLocationQueryResult',
        typeName: 'CT_GeoLocationQueryResult',
        propertyInfos: [{
            name: 'geoLocationQuery',
            typeInfo: '.CTGeoLocationQuery'
          }, {
            name: 'geoLocations',
            typeInfo: '.CTGeoLocations'
          }]
      }, {
        localName: 'CTGeoLocationQueryResults',
        typeName: 'CT_GeoLocationQueryResults',
        propertyInfos: [{
            name: 'geoLocationQueryResult',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTGeoLocationQueryResult'
          }]
      }, {
        localName: 'CTGeoLocations',
        typeName: 'CT_GeoLocations',
        propertyInfos: [{
            name: 'geoLocation',
            typeInfo: '.CTGeoLocation'
          }]
      }, {
        localName: 'CTGeoParentEntitiesQuery',
        typeName: 'CT_GeoParentEntitiesQuery',
        propertyInfos: [{
            name: 'entityId',
            required: true,
            attributeName: {
              localPart: 'entityId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoParentEntitiesQueryResult',
        typeName: 'CT_GeoParentEntitiesQueryResult',
        propertyInfos: [{
            name: 'geoParentEntitiesQuery',
            required: true,
            typeInfo: '.CTGeoParentEntitiesQuery'
          }, {
            name: 'geoEntity',
            typeInfo: '.CTGeoEntity'
          }, {
            name: 'geoParentEntity',
            typeInfo: '.CTGeoParentEntity'
          }]
      }, {
        localName: 'CTGeoParentEntitiesQueryResults',
        typeName: 'CT_GeoParentEntitiesQueryResults',
        propertyInfos: [{
            name: 'geoParentEntitiesQueryResult',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTGeoParentEntitiesQueryResult'
          }]
      }, {
        localName: 'CTGeoParentEntity',
        typeName: 'CT_GeoParentEntity',
        propertyInfos: [{
            name: 'entityId',
            required: true,
            attributeName: {
              localPart: 'entityId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoPolygon',
        typeName: 'CT_GeoPolygon',
        propertyInfos: [{
            name: 'numPoints',
            required: true,
            typeInfo: 'Integer',
            attributeName: {
              localPart: 'numPoints'
            },
            type: 'attribute'
          }, {
            name: 'pcaRings',
            required: true,
            attributeName: {
              localPart: 'pcaRings'
            },
            type: 'attribute'
          }, {
            name: 'polygonId',
            required: true,
            attributeName: {
              localPart: 'polygonId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGeoPolygons',
        typeName: 'CT_GeoPolygons',
        propertyInfos: [{
            name: 'geoPolygon',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTGeoPolygon'
          }]
      }, {
        localName: 'CTGeography',
        typeName: 'CT_Geography',
        propertyInfos: [{
            name: 'geoCache',
            typeInfo: '.CTGeoCache'
          }, {
            name: 'attribution',
            required: true,
            attributeName: {
              localPart: 'attribution'
            },
            type: 'attribute'
          }, {
            name: 'cultureLanguage',
            required: true,
            typeInfo: 'Language',
            attributeName: {
              localPart: 'cultureLanguage'
            },
            type: 'attribute'
          }, {
            name: 'cultureRegion',
            required: true,
            attributeName: {
              localPart: 'cultureRegion'
            },
            type: 'attribute'
          }, {
            name: 'projectionType',
            typeInfo: '.STGeoProjectionType',
            attributeName: {
              localPart: 'projectionType'
            },
            type: 'attribute'
          }, {
            name: 'viewedRegionType',
            typeInfo: '.STGeoMappingLevel',
            attributeName: {
              localPart: 'viewedRegionType'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGridlines',
        typeName: 'CT_Gridlines',
        propertyInfos: [{
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }]
      }, {
        localName: 'CTHeaderFooter',
        typeName: 'CT_HeaderFooter',
        propertyInfos: [{
            name: 'oddHeader'
          }, {
            name: 'oddFooter'
          }, {
            name: 'evenHeader'
          }, {
            name: 'evenFooter'
          }, {
            name: 'firstHeader'
          }, {
            name: 'firstFooter'
          }, {
            name: 'alignWithMargins',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'alignWithMargins'
            },
            type: 'attribute'
          }, {
            name: 'differentFirst',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'differentFirst'
            },
            type: 'attribute'
          }, {
            name: 'differentOddEven',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'differentOddEven'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTLegend',
        typeName: 'CT_Legend',
        propertyInfos: [{
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'txPr',
            typeInfo: 'org_docx4j_dml.CTTextBody'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }, {
            name: 'align',
            typeInfo: '.STPosAlign',
            defaultValue: 'ctr',
            attributeName: {
              localPart: 'align'
            },
            type: 'attribute'
          }, {
            name: 'overlay',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'overlay'
            },
            type: 'attribute'
          }, {
            name: 'pos',
            typeInfo: '.STSidePos',
            defaultValue: 'r',
            attributeName: {
              localPart: 'pos'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTNumberColorPosition',
        typeName: 'CT_NumberColorPosition',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTNumberFormat',
        typeName: 'CT_NumberFormat',
        propertyInfos: [{
            name: 'formatCode',
            required: true,
            attributeName: {
              localPart: 'formatCode'
            },
            type: 'attribute'
          }, {
            name: 'sourceLinked',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'sourceLinked'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTNumericDimension',
        typeName: 'CT_NumericDimension',
        propertyInfos: [{
            name: 'content',
            required: true,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'f',
                typeInfo: '.CTFormula'
              }, {
                elementName: 'nf',
                typeInfo: '.CTFormula'
              }, {
                elementName: 'lvl',
                typeInfo: '.CTNumericLevel'
              }],
            type: 'elementRefs'
          }, {
            name: 'type',
            required: true,
            typeInfo: '.STNumericDimensionType',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTNumericLevel',
        typeName: 'CT_NumericLevel',
        propertyInfos: [{
            name: 'pt',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTNumericValue'
          }, {
            name: 'formatCode',
            attributeName: {
              localPart: 'formatCode'
            },
            type: 'attribute'
          }, {
            name: 'name',
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'ptCount',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'ptCount'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTNumericValue',
        typeName: 'CT_NumericValue',
        propertyInfos: [{
            name: 'value',
            typeInfo: 'Double',
            type: 'value'
          }, {
            name: 'idx',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'idx'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPageMargins',
        typeName: 'CT_PageMargins',
        propertyInfos: [{
            name: 'b',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'b'
            },
            type: 'attribute'
          }, {
            name: 'footer',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'footer'
            },
            type: 'attribute'
          }, {
            name: 'header',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'header'
            },
            type: 'attribute'
          }, {
            name: 'l',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'l'
            },
            type: 'attribute'
          }, {
            name: 'r',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'r'
            },
            type: 'attribute'
          }, {
            name: 't',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 't'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPageSetup',
        typeName: 'CT_PageSetup',
        propertyInfos: [{
            name: 'blackAndWhite',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'blackAndWhite'
            },
            type: 'attribute'
          }, {
            name: 'copies',
            typeInfo: 'UnsignedInt',
            defaultValue: 1,
            attributeName: {
              localPart: 'copies'
            },
            type: 'attribute'
          }, {
            name: 'draft',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'draft'
            },
            type: 'attribute'
          }, {
            name: 'firstPageNumber',
            typeInfo: 'UnsignedInt',
            defaultValue: 1,
            attributeName: {
              localPart: 'firstPageNumber'
            },
            type: 'attribute'
          }, {
            name: 'horizontalDpi',
            typeInfo: 'Int',
            defaultValue: 600,
            attributeName: {
              localPart: 'horizontalDpi'
            },
            type: 'attribute'
          }, {
            name: 'orientation',
            typeInfo: '.STPageOrientation',
            defaultValue: 'default',
            attributeName: {
              localPart: 'orientation'
            },
            type: 'attribute'
          }, {
            name: 'paperSize',
            typeInfo: 'UnsignedInt',
            defaultValue: 1,
            attributeName: {
              localPart: 'paperSize'
            },
            type: 'attribute'
          }, {
            name: 'useFirstPageNumber',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'useFirstPageNumber'
            },
            type: 'attribute'
          }, {
            name: 'verticalDpi',
            typeInfo: 'Int',
            defaultValue: 600,
            attributeName: {
              localPart: 'verticalDpi'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTParentLabelLayout',
        typeName: 'CT_ParentLabelLayout',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: '.STParentLabelLayout',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPercentageColorPosition',
        typeName: 'CT_PercentageColorPosition',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'Double',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPlotArea',
        typeName: 'CT_PlotArea',
        propertyInfos: [{
            name: 'plotAreaRegion',
            required: true,
            typeInfo: '.CTPlotAreaRegion'
          }, {
            name: 'axis',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTAxis'
          }, {
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }]
      }, {
        localName: 'CTPlotAreaRegion',
        typeName: 'CT_PlotAreaRegion',
        propertyInfos: [{
            name: 'plotSurface',
            typeInfo: '.CTPlotSurface'
          }, {
            name: 'series',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTSeries'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }]
      }, {
        localName: 'CTPlotSurface',
        typeName: 'CT_PlotSurface',
        propertyInfos: [{
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }]
      }, {
        localName: 'CTPrintSettings',
        typeName: 'CT_PrintSettings',
        propertyInfos: [{
            name: 'headerFooter',
            typeInfo: '.CTHeaderFooter'
          }, {
            name: 'pageMargins',
            typeInfo: '.CTPageMargins'
          }, {
            name: 'pageSetup',
            typeInfo: '.CTPageSetup'
          }]
      }, {
        localName: 'CTRegionLabelLayout',
        typeName: 'CT_RegionLabelLayout',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: '.STRegionLabelLayout',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTRelId',
        typeName: 'CT_RelId',
        propertyInfos: [{
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSeries',
        typeName: 'CT_Series',
        propertyInfos: [{
            name: 'tx',
            typeInfo: '.CTText'
          }, {
            name: 'spPr',
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }, {
            name: 'valueColors',
            typeInfo: '.CTValueColors'
          }, {
            name: 'valueColorPositions',
            typeInfo: '.CTValueColorPositions'
          }, {
            name: 'dataPt',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTDataPoint'
          }, {
            name: 'dataLabels',
            typeInfo: '.CTDataLabels'
          }, {
            name: 'dataId',
            typeInfo: '.CTDataId'
          }, {
            name: 'layoutPr',
            typeInfo: '.CTSeriesLayoutProperties'
          }, {
            name: 'axisId',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTAxisId'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }, {
            name: 'formatIdx',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'formatIdx'
            },
            type: 'attribute'
          }, {
            name: 'hidden',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'hidden'
            },
            type: 'attribute'
          }, {
            name: 'layoutId',
            required: true,
            typeInfo: '.STSeriesLayout',
            attributeName: {
              localPart: 'layoutId'
            },
            type: 'attribute'
          }, {
            name: 'ownerIdx',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'ownerIdx'
            },
            type: 'attribute'
          }, {
            name: 'uniqueId',
            attributeName: {
              localPart: 'uniqueId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSeriesElementVisibilities',
        typeName: 'CT_SeriesElementVisibilities',
        propertyInfos: [{
            name: 'connectorLines',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'connectorLines'
            },
            type: 'attribute'
          }, {
            name: 'meanLine',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'meanLine'
            },
            type: 'attribute'
          }, {
            name: 'meanMarker',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'meanMarker'
            },
            type: 'attribute'
          }, {
            name: 'nonoutliers',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'nonoutliers'
            },
            type: 'attribute'
          }, {
            name: 'outliers',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'outliers'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSeriesLayoutProperties',
        typeName: 'CT_SeriesLayoutProperties',
        propertyInfos: [{
            name: 'parentLabelLayout',
            typeInfo: '.CTParentLabelLayout'
          }, {
            name: 'regionLabelLayout',
            typeInfo: '.CTRegionLabelLayout'
          }, {
            name: 'visibility',
            typeInfo: '.CTSeriesElementVisibilities'
          }, {
            name: 'aggregation',
            typeInfo: '.CTAggregation'
          }, {
            name: 'binning',
            typeInfo: '.CTBinning'
          }, {
            name: 'geography',
            typeInfo: '.CTGeography'
          }, {
            name: 'statistics',
            typeInfo: '.CTStatistics'
          }, {
            name: 'subtotals',
            typeInfo: '.CTSubtotals'
          }, {
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }]
      }, {
        localName: 'CTStatistics',
        typeName: 'CT_Statistics',
        propertyInfos: [{
            name: 'quartileMethod',
            typeInfo: '.STQuartileMethod',
            attributeName: {
              localPart: 'quartileMethod'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTStringDimension',
        typeName: 'CT_StringDimension',
        propertyInfos: [{
            name: 'content',
            required: true,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'f',
                typeInfo: '.CTFormula'
              }, {
                elementName: 'nf',
                typeInfo: '.CTFormula'
              }, {
                elementName: 'lvl',
                typeInfo: '.CTStringLevel'
              }],
            type: 'elementRefs'
          }, {
            name: 'type',
            required: true,
            typeInfo: '.STStringDimensionType',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTStringLevel',
        typeName: 'CT_StringLevel',
        propertyInfos: [{
            name: 'pt',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTStringValue'
          }, {
            name: 'name',
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'ptCount',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'ptCount'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTStringValue',
        typeName: 'CT_StringValue',
        propertyInfos: [{
            name: 'value',
            type: 'value'
          }, {
            name: 'idx',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'idx'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSubtotalIndex',
        typeName: 'CT_SubtotalIndex',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSubtotals',
        typeName: 'CT_Subtotals',
        propertyInfos: [{
            name: 'idx',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTSubtotalIndex'
          }]
      }, {
        localName: 'CTText',
        typeName: 'CT_Text',
        propertyInfos: [{
            name: 'txData',
            typeInfo: '.CTTextData'
          }, {
            name: 'rich',
            typeInfo: 'org_docx4j_dml.CTTextBody'
          }]
      }, {
        localName: 'CTTextData',
        typeName: 'CT_TextData',
        propertyInfos: [{
            name: 'content',
            required: true,
            maxOccurs: 2,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'f',
                typeInfo: '.CTFormula'
              }, {
                elementName: 'v'
              }],
            type: 'elementRefs'
          }]
      }, {
        localName: 'CTTickLabels',
        typeName: 'CT_TickLabels',
        propertyInfos: [{
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }]
      }, {
        localName: 'CTTickMarks',
        typeName: 'CT_TickMarks',
        propertyInfos: [{
            name: 'extLst',
            typeInfo: '.CTExtensionList'
          }, {
            name: 'type',
            typeInfo: '.STTickMarksType',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTValueAxisScaling',
        typeName: 'CT_ValueAxisScaling',
        propertyInfos: [{
            name: 'majorUnit',
            attributeName: {
              localPart: 'majorUnit'
            },
            type: 'attribute'
          }, {
            name: 'max',
            attributeName: {
              localPart: 'max'
            },
            type: 'attribute'
          }, {
            name: 'min',
            attributeName: {
              localPart: 'min'
            },
            type: 'attribute'
          }, {
            name: 'minorUnit',
            attributeName: {
              localPart: 'minorUnit'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTValueColorEndPosition',
        typeName: 'CT_ValueColorEndPosition',
        propertyInfos: [{
            name: 'extremeValue',
            typeInfo: '.CTExtremeValueColorPosition'
          }, {
            name: 'number',
            typeInfo: '.CTNumberColorPosition'
          }, {
            name: 'percent',
            typeInfo: '.CTPercentageColorPosition'
          }]
      }, {
        localName: 'CTValueColorMiddlePosition',
        typeName: 'CT_ValueColorMiddlePosition',
        propertyInfos: [{
            name: 'number',
            typeInfo: '.CTNumberColorPosition'
          }, {
            name: 'percent',
            typeInfo: '.CTPercentageColorPosition'
          }]
      }, {
        localName: 'CTValueColorPositions',
        typeName: 'CT_ValueColorPositions',
        propertyInfos: [{
            name: 'min',
            typeInfo: '.CTValueColorEndPosition'
          }, {
            name: 'mid',
            typeInfo: '.CTValueColorMiddlePosition'
          }, {
            name: 'max',
            typeInfo: '.CTValueColorEndPosition'
          }, {
            name: 'count',
            typeInfo: 'Int',
            defaultValue: 2,
            attributeName: {
              localPart: 'count'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTValueColors',
        typeName: 'CT_ValueColors',
        propertyInfos: [{
            name: 'minColor',
            typeInfo: 'org_docx4j_dml.CTSolidColorFillProperties'
          }, {
            name: 'midColor',
            typeInfo: 'org_docx4j_dml.CTSolidColorFillProperties'
          }, {
            name: 'maxColor',
            typeInfo: 'org_docx4j_dml.CTSolidColorFillProperties'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STAxisUnit',
        values: ['hundreds', 'thousands', 'tenThousands', 'hundredThousands', 'millions', 'tenMillions', 'hundredMillions', 'billions', 'trillions', 'percentage']
      }, {
        type: 'enumInfo',
        localName: 'STDataLabelPos',
        values: ['bestFit', 'b', 'ctr', 'inBase', 'inEnd', 'l', 'outEnd', 'r', 't']
      }, {
        type: 'enumInfo',
        localName: 'STEntityType',
        values: ['Address', 'AdminDistrict', 'AdminDistrict2', 'AdminDistrict3', 'Continent', 'CountryRegion', 'Locality', 'Ocean', 'Planet', 'PostalCode', 'Region', 'Unsupported']
      }, {
        type: 'enumInfo',
        localName: 'STFormulaDirection',
        values: ['col', 'row']
      }, {
        type: 'enumInfo',
        localName: 'STGeoMappingLevel',
        values: ['dataOnly', 'postalCode', 'county', 'state', 'countryRegion', 'countryRegionList', 'world']
      }, {
        type: 'enumInfo',
        localName: 'STGeoProjectionType',
        values: ['mercator', 'miller', 'robinson', 'albers']
      }, {
        type: 'enumInfo',
        localName: 'STIntervalClosedSide',
        values: ['l', 'r']
      }, {
        type: 'enumInfo',
        localName: 'STNumericDimensionType',
        values: ['val', 'x', 'y', 'size', 'colorVal']
      }, {
        type: 'enumInfo',
        localName: 'STPageOrientation',
        values: ['default', 'portrait', 'landscape']
      }, {
        type: 'enumInfo',
        localName: 'STParentLabelLayout',
        values: ['none', 'banner', 'overlapping']
      }, {
        type: 'enumInfo',
        localName: 'STPosAlign',
        values: ['min', 'ctr', 'max']
      }, {
        type: 'enumInfo',
        localName: 'STQuartileMethod',
        values: ['inclusive', 'exclusive']
      }, {
        type: 'enumInfo',
        localName: 'STRegionLabelLayout',
        values: ['none', 'bestFitOnly', 'showAll']
      }, {
        type: 'enumInfo',
        localName: 'STSeriesLayout',
        values: ['boxWhisker', 'clusteredColumn', 'funnel', 'paretoLine', 'regionMap', 'sunburst', 'treemap', 'waterfall']
      }, {
        type: 'enumInfo',
        localName: 'STSidePos',
        values: ['l', 't', 'r', 'b']
      }, {
        type: 'enumInfo',
        localName: 'STStringDimensionType',
        values: ['cat', 'colorStr', 'entityId']
      }, {
        type: 'enumInfo',
        localName: 'STTickMarksType',
        values: ['in', 'out', 'cross', 'none']
      }],
    elementInfos: [{
        typeInfo: '.CTRelId',
        elementName: 'chart'
      }, {
        typeInfo: '.CTChartSpace',
        elementName: 'chartSpace'
      }, {
        typeInfo: '.CTFormula',
        elementName: 'f',
        scope: '.CTNumericDimension'
      }, {
        typeInfo: '.CTFormula',
        elementName: 'f',
        scope: '.CTStringDimension'
      }, {
        typeInfo: '.CTFormula',
        elementName: 'f',
        scope: '.CTTextData'
      }, {
        typeInfo: '.CTNumericLevel',
        elementName: 'lvl',
        scope: '.CTNumericDimension'
      }, {
        typeInfo: '.CTStringLevel',
        elementName: 'lvl',
        scope: '.CTStringDimension'
      }, {
        typeInfo: '.CTFormula',
        elementName: 'nf',
        scope: '.CTNumericDimension'
      }, {
        typeInfo: '.CTFormula',
        elementName: 'nf',
        scope: '.CTStringDimension'
      }, {
        elementName: 'v',
        scope: '.CTTextData'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex: org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex = org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex = org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2014_chartex;
  }
}