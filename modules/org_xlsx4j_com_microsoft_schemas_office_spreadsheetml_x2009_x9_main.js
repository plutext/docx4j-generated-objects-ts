var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main_Module_Factory = function () {
  var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main = {
    name: 'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2009\/9\/main',
    dependencies: ['org_xlsx4j_schemas_microsoft_com_office_excel_2006_main', 'org_xlsx4j_sml'],
    typeInfos: [{
        localName: 'CTCacheField',
        typeName: 'CT_CacheField',
        propertyInfos: [{
            name: 'ignore',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'ignore'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCacheHierarchy',
        typeName: 'CT_CacheHierarchy',
        propertyInfos: [{
            name: 'setLevels',
            typeInfo: '.CTSetLevels'
          }, {
            name: 'flattenHierarchies',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'flattenHierarchies'
            },
            type: 'attribute'
          }, {
            name: 'hierarchizeDistinct',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'hierarchizeDistinct'
            },
            type: 'attribute'
          }, {
            name: 'ignore',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'ignore'
            },
            type: 'attribute'
          }, {
            name: 'measuresSet',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'measuresSet'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCacheSourceExt',
        typeName: 'CT_CacheSourceExt',
        propertyInfos: [{
            name: 'sourceConnection',
            required: true,
            typeInfo: '.CTSourceConnection'
          }]
      }, {
        localName: 'CTCalculatedMember',
        typeName: 'CT_CalculatedMember',
        propertyInfos: [{
            name: 'tupleSet',
            typeInfo: '.CTTupleSet'
          }, {
            name: 'displayFolder',
            attributeName: {
              localPart: 'displayFolder'
            },
            type: 'attribute'
          }, {
            name: 'dynamicSet',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'dynamicSet'
            },
            type: 'attribute'
          }, {
            name: 'flattenHierarchies',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'flattenHierarchies'
            },
            type: 'attribute'
          }, {
            name: 'hierarchizeDistinct',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'hierarchizeDistinct'
            },
            type: 'attribute'
          }, {
            name: 'mdxLong',
            attributeName: {
              localPart: 'mdxLong'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCfIcon',
        typeName: 'CT_CfIcon',
        propertyInfos: [{
            name: 'iconId',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'iconId'
            },
            type: 'attribute'
          }, {
            name: 'iconSet',
            required: true,
            values: ['3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters', '3Stars', '3Triangles', '5Boxes', 'NoIcons'],
            attributeName: {
              localPart: 'iconSet'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCfRule',
        typeName: 'CT_CfRule',
        propertyInfos: [{
            name: 'f',
            minOccurs: 0,
            maxOccurs: 3,
            collection: true,
            elementName: {
              localPart: 'f',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2006\/main'
            }
          }, {
            name: 'colorScale',
            typeInfo: '.CTColorScale'
          }, {
            name: 'dataBar',
            typeInfo: '.CTDataBar'
          }, {
            name: 'iconSet',
            typeInfo: '.CTIconSet'
          }, {
            name: 'dxf',
            typeInfo: 'org_xlsx4j_sml.CTDxf'
          }, {
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }, {
            name: 'aboveAverage',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'aboveAverage'
            },
            type: 'attribute'
          }, {
            name: 'activePresent',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'activePresent'
            },
            type: 'attribute'
          }, {
            name: 'bottom',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'bottom'
            },
            type: 'attribute'
          }, {
            name: 'equalAverage',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'equalAverage'
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
            name: 'operator',
            typeInfo: 'org_xlsx4j_sml.STConditionalFormattingOperator',
            attributeName: {
              localPart: 'operator'
            },
            type: 'attribute'
          }, {
            name: 'percent',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'percent'
            },
            type: 'attribute'
          }, {
            name: 'priority',
            typeInfo: 'Int',
            attributeName: {
              localPart: 'priority'
            },
            type: 'attribute'
          }, {
            name: 'rank',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'rank'
            },
            type: 'attribute'
          }, {
            name: 'stdDev',
            typeInfo: 'Int',
            attributeName: {
              localPart: 'stdDev'
            },
            type: 'attribute'
          }, {
            name: 'stopIfTrue',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'stopIfTrue'
            },
            type: 'attribute'
          }, {
            name: 'text',
            attributeName: {
              localPart: 'text'
            },
            type: 'attribute'
          }, {
            name: 'timePeriod',
            typeInfo: 'org_xlsx4j_sml.STTimePeriod',
            attributeName: {
              localPart: 'timePeriod'
            },
            type: 'attribute'
          }, {
            name: 'type',
            typeInfo: 'org_xlsx4j_sml.STCfType',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCfvo',
        typeName: 'CT_Cfvo',
        propertyInfos: [{
            name: 'f',
            elementName: {
              localPart: 'f',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2006\/main'
            }
          }, {
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }, {
            name: 'gte',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'gte'
            },
            type: 'attribute'
          }, {
            name: 'type',
            required: true,
            typeInfo: '.STCfvoType',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTColorScale',
        typeName: 'CT_ColorScale',
        propertyInfos: [{
            name: 'cfvo',
            required: true,
            minOccurs: 2,
            collection: true,
            typeInfo: '.CTCfvo'
          }, {
            name: 'color',
            required: true,
            minOccurs: 2,
            collection: true,
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }]
      }, {
        localName: 'CTConditionalFormat',
        typeName: 'CT_ConditionalFormat',
        propertyInfos: [{
            name: 'pivotAreas',
            typeInfo: 'org_xlsx4j_sml.CTPivotAreas'
          }, {
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }, {
            name: 'id',
            required: true,
            typeInfo: 'Token',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'priority',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'priority'
            },
            type: 'attribute'
          }, {
            name: 'scope',
            typeInfo: 'org_xlsx4j_sml.STScope',
            defaultValue: 'selection',
            attributeName: {
              localPart: 'scope'
            },
            type: 'attribute'
          }, {
            name: 'type',
            typeInfo: 'org_xlsx4j_sml.STType',
            defaultValue: 'none',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTConditionalFormats',
        typeName: 'CT_ConditionalFormats',
        propertyInfos: [{
            name: 'conditionalFormat',
            required: true,
            collection: true,
            typeInfo: '.CTConditionalFormat'
          }, {
            name: 'count',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'count'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTConditionalFormatting',
        typeName: 'CT_ConditionalFormatting',
        propertyInfos: [{
            name: 'cfRule',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTCfRule'
          }, {
            name: 'sqref',
            elementName: {
              localPart: 'sqref',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2006\/main'
            },
            typeInfo: 'org_xlsx4j_schemas_microsoft_com_office_excel_2006_main.CTSqref'
          }, {
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }, {
            name: 'pivot',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'pivot'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTConditionalFormattings',
        typeName: 'CT_ConditionalFormattings',
        propertyInfos: [{
            name: 'conditionalFormatting',
            required: true,
            collection: true,
            typeInfo: '.CTConditionalFormatting'
          }]
      }, {
        localName: 'CTConnection',
        typeName: 'CT_Connection',
        propertyInfos: [{
            name: 'calculatedMembers',
            typeInfo: 'org_xlsx4j_sml.CTCalculatedMembers'
          }, {
            name: 'culture',
            attributeName: {
              localPart: 'culture'
            },
            type: 'attribute'
          }, {
            name: 'embeddedDataId',
            attributeName: {
              localPart: 'embeddedDataId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCustomFilter',
        typeName: 'CT_CustomFilter',
        propertyInfos: [{
            name: 'operator',
            typeInfo: 'org_xlsx4j_sml.STFilterOperator',
            defaultValue: 'equal',
            attributeName: {
              localPart: 'operator'
            },
            type: 'attribute'
          }, {
            name: 'val',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCustomFilters',
        typeName: 'CT_CustomFilters',
        propertyInfos: [{
            name: 'customFilter',
            required: true,
            maxOccurs: 2,
            collection: true,
            typeInfo: '.CTCustomFilter'
          }, {
            name: 'and',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'and'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDataBar',
        typeName: 'CT_DataBar',
        propertyInfos: [{
            name: 'cfvo',
            required: true,
            minOccurs: 2,
            maxOccurs: 2,
            collection: true,
            typeInfo: '.CTCfvo'
          }, {
            name: 'fillColor',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'borderColor',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'negativeFillColor',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'negativeBorderColor',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'axisColor',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'axisPosition',
            typeInfo: '.STDataBarAxisPosition',
            defaultValue: 'automatic',
            attributeName: {
              localPart: 'axisPosition'
            },
            type: 'attribute'
          }, {
            name: 'border',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'border'
            },
            type: 'attribute'
          }, {
            name: 'direction',
            typeInfo: '.STDataBarDirection',
            defaultValue: 'context',
            attributeName: {
              localPart: 'direction'
            },
            type: 'attribute'
          }, {
            name: 'gradient',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'gradient'
            },
            type: 'attribute'
          }, {
            name: 'maxLength',
            typeInfo: 'UnsignedInt',
            defaultValue: 90,
            attributeName: {
              localPart: 'maxLength'
            },
            type: 'attribute'
          }, {
            name: 'minLength',
            typeInfo: 'UnsignedInt',
            defaultValue: 10,
            attributeName: {
              localPart: 'minLength'
            },
            type: 'attribute'
          }, {
            name: 'negativeBarBorderColorSameAsPositive',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'negativeBarBorderColorSameAsPositive'
            },
            type: 'attribute'
          }, {
            name: 'negativeBarColorSameAsPositive',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'negativeBarColorSameAsPositive'
            },
            type: 'attribute'
          }, {
            name: 'showValue',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'showValue'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDataField',
        typeName: 'CT_DataField',
        propertyInfos: [{
            name: 'pivotShowAs',
            typeInfo: '.STPivotShowAs',
            attributeName: {
              localPart: 'pivotShowAs'
            },
            type: 'attribute'
          }, {
            name: 'sourceField',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'sourceField'
            },
            type: 'attribute'
          }, {
            name: 'uniqueName',
            attributeName: {
              localPart: 'uniqueName'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDataValidation',
        typeName: 'CT_DataValidation',
        propertyInfos: [{
            name: 'formula1',
            typeInfo: '.CTDataValidationFormula'
          }, {
            name: 'formula2',
            typeInfo: '.CTDataValidationFormula'
          }, {
            name: 'sqref',
            required: true,
            elementName: {
              localPart: 'sqref',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2006\/main'
            },
            typeInfo: 'org_xlsx4j_schemas_microsoft_com_office_excel_2006_main.CTSqref'
          }, {
            name: 'allowBlank',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'allowBlank'
            },
            type: 'attribute'
          }, {
            name: 'error',
            attributeName: {
              localPart: 'error'
            },
            type: 'attribute'
          }, {
            name: 'errorStyle',
            typeInfo: 'org_xlsx4j_sml.STDataValidationErrorStyle',
            defaultValue: 'stop',
            attributeName: {
              localPart: 'errorStyle'
            },
            type: 'attribute'
          }, {
            name: 'errorTitle',
            attributeName: {
              localPart: 'errorTitle'
            },
            type: 'attribute'
          }, {
            name: 'imeMode',
            typeInfo: 'org_xlsx4j_sml.STDataValidationImeMode',
            defaultValue: 'noControl',
            attributeName: {
              localPart: 'imeMode'
            },
            type: 'attribute'
          }, {
            name: 'operator',
            typeInfo: 'org_xlsx4j_sml.STDataValidationOperator',
            defaultValue: 'between',
            attributeName: {
              localPart: 'operator'
            },
            type: 'attribute'
          }, {
            name: 'prompt',
            attributeName: {
              localPart: 'prompt'
            },
            type: 'attribute'
          }, {
            name: 'promptTitle',
            attributeName: {
              localPart: 'promptTitle'
            },
            type: 'attribute'
          }, {
            name: 'showDropDown',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'showDropDown'
            },
            type: 'attribute'
          }, {
            name: 'showErrorMessage',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'showErrorMessage'
            },
            type: 'attribute'
          }, {
            name: 'showInputMessage',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'showInputMessage'
            },
            type: 'attribute'
          }, {
            name: 'type',
            typeInfo: 'org_xlsx4j_sml.STDataValidationType',
            defaultValue: 'none',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: 'uid',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'uid',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2014\/revision'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDataValidationFormula',
        typeName: 'CT_DataValidationFormula',
        propertyInfos: [{
            name: 'f',
            required: true,
            elementName: {
              localPart: 'f',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2006\/main'
            }
          }]
      }, {
        localName: 'CTDataValidations',
        typeName: 'CT_DataValidations',
        propertyInfos: [{
            name: 'dataValidation',
            required: true,
            collection: true,
            typeInfo: '.CTDataValidation'
          }, {
            name: 'count',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'count'
            },
            type: 'attribute'
          }, {
            name: 'disablePrompts',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'disablePrompts'
            },
            type: 'attribute'
          }, {
            name: 'xWindow',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'xWindow'
            },
            type: 'attribute'
          }, {
            name: 'yWindow',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'yWindow'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDatastoreItem',
        typeName: 'CT_DatastoreItem',
        propertyInfos: [{
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTFilter',
        typeName: 'CT_Filter',
        propertyInfos: [{
            name: 'val',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTFormControlPr',
        typeName: 'CT_FormControlPr',
        propertyInfos: [{
            name: 'itemLst',
            typeInfo: '.CTListItems'
          }, {
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }, {
            name: 'checked',
            typeInfo: '.STChecked',
            attributeName: {
              localPart: 'checked'
            },
            type: 'attribute'
          }, {
            name: 'colored',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'colored'
            },
            type: 'attribute'
          }, {
            name: 'dropLines',
            typeInfo: 'UnsignedInt',
            defaultValue: 8,
            attributeName: {
              localPart: 'dropLines'
            },
            type: 'attribute'
          }, {
            name: 'dropStyle',
            typeInfo: '.STDropStyle',
            attributeName: {
              localPart: 'dropStyle'
            },
            type: 'attribute'
          }, {
            name: 'dx',
            typeInfo: 'UnsignedInt',
            defaultValue: 80,
            attributeName: {
              localPart: 'dx'
            },
            type: 'attribute'
          }, {
            name: 'editVal',
            typeInfo: '.STEditValidation',
            attributeName: {
              localPart: 'editVal'
            },
            type: 'attribute'
          }, {
            name: 'firstButton',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'firstButton'
            },
            type: 'attribute'
          }, {
            name: 'fmlaGroup',
            attributeName: {
              localPart: 'fmlaGroup'
            },
            type: 'attribute'
          }, {
            name: 'fmlaLink',
            attributeName: {
              localPart: 'fmlaLink'
            },
            type: 'attribute'
          }, {
            name: 'fmlaRange',
            attributeName: {
              localPart: 'fmlaRange'
            },
            type: 'attribute'
          }, {
            name: 'fmlaTxbx',
            attributeName: {
              localPart: 'fmlaTxbx'
            },
            type: 'attribute'
          }, {
            name: 'horiz',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'horiz'
            },
            type: 'attribute'
          }, {
            name: 'inc',
            typeInfo: 'UnsignedInt',
            defaultValue: 1,
            attributeName: {
              localPart: 'inc'
            },
            type: 'attribute'
          }, {
            name: 'justLastX',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'justLastX'
            },
            type: 'attribute'
          }, {
            name: 'lockText',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'lockText'
            },
            type: 'attribute'
          }, {
            name: 'max',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'max'
            },
            type: 'attribute'
          }, {
            name: 'min',
            typeInfo: 'UnsignedInt',
            defaultValue: 0,
            attributeName: {
              localPart: 'min'
            },
            type: 'attribute'
          }, {
            name: 'multiLine',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'multiLine'
            },
            type: 'attribute'
          }, {
            name: 'multiSel',
            attributeName: {
              localPart: 'multiSel'
            },
            type: 'attribute'
          }, {
            name: 'noThreeD',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'noThreeD'
            },
            type: 'attribute'
          }, {
            name: 'noThreeD2',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'noThreeD2'
            },
            type: 'attribute'
          }, {
            name: 'objectType',
            typeInfo: '.STObjectType',
            attributeName: {
              localPart: 'objectType'
            },
            type: 'attribute'
          }, {
            name: 'page',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'page'
            },
            type: 'attribute'
          }, {
            name: 'passwordEdit',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'passwordEdit'
            },
            type: 'attribute'
          }, {
            name: 'sel',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'sel'
            },
            type: 'attribute'
          }, {
            name: 'seltype',
            typeInfo: '.STSelType',
            defaultValue: 'single',
            attributeName: {
              localPart: 'seltype'
            },
            type: 'attribute'
          }, {
            name: 'textHAlign',
            typeInfo: '.STTextHAlign',
            defaultValue: 'left',
            attributeName: {
              localPart: 'textHAlign'
            },
            type: 'attribute'
          }, {
            name: 'textVAlign',
            typeInfo: '.STTextVAlign',
            defaultValue: 'top',
            attributeName: {
              localPart: 'textVAlign'
            },
            type: 'attribute'
          }, {
            name: 'val',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }, {
            name: 'verticalBar',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'verticalBar'
            },
            type: 'attribute'
          }, {
            name: 'widthMin',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'widthMin'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTIconFilter',
        typeName: 'CT_IconFilter',
        propertyInfos: [{
            name: 'iconId',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'iconId'
            },
            type: 'attribute'
          }, {
            name: 'iconSet',
            required: true,
            values: ['3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters', '3Stars', '3Triangles', '5Boxes', 'NoIcons'],
            attributeName: {
              localPart: 'iconSet'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTIconSet',
        typeName: 'CT_IconSet',
        propertyInfos: [{
            name: 'cfvo',
            required: true,
            minOccurs: 2,
            collection: true,
            typeInfo: '.CTCfvo'
          }, {
            name: 'cfIcon',
            minOccurs: 0,
            maxOccurs: 5,
            collection: true,
            typeInfo: '.CTCfIcon'
          }, {
            name: 'custom',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'custom'
            },
            type: 'attribute'
          }, {
            name: 'iconSet',
            values: ['3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters', '3Stars', '3Triangles', '5Boxes', 'NoIcons'],
            defaultValue: '3TrafficLights1',
            attributeName: {
              localPart: 'iconSet'
            },
            type: 'attribute'
          }, {
            name: 'percent',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'percent'
            },
            type: 'attribute'
          }, {
            name: 'reverse',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'reverse'
            },
            type: 'attribute'
          }, {
            name: 'showValue',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'showValue'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTIgnoredError',
        typeName: 'CT_IgnoredError',
        propertyInfos: [{
            name: 'sqref',
            required: true,
            elementName: {
              localPart: 'sqref',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2006\/main'
            },
            typeInfo: 'org_xlsx4j_schemas_microsoft_com_office_excel_2006_main.CTSqref'
          }, {
            name: 'calculatedColumn',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'calculatedColumn'
            },
            type: 'attribute'
          }, {
            name: 'emptyCellReference',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'emptyCellReference'
            },
            type: 'attribute'
          }, {
            name: 'evalError',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'evalError'
            },
            type: 'attribute'
          }, {
            name: 'formula',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'formula'
            },
            type: 'attribute'
          }, {
            name: 'formulaRange',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'formulaRange'
            },
            type: 'attribute'
          }, {
            name: 'listDataValidation',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'listDataValidation'
            },
            type: 'attribute'
          }, {
            name: 'numberStoredAsText',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'numberStoredAsText'
            },
            type: 'attribute'
          }, {
            name: 'twoDigitTextYear',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'twoDigitTextYear'
            },
            type: 'attribute'
          }, {
            name: 'unlockedFormula',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'unlockedFormula'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTIgnoredErrors',
        typeName: 'CT_IgnoredErrors',
        propertyInfos: [{
            name: 'ignoredError',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTIgnoredError'
          }, {
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }]
      }, {
        localName: 'CTListItem',
        typeName: 'CT_ListItem',
        propertyInfos: [{
            name: 'val',
            required: true,
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTListItems',
        typeName: 'CT_ListItems',
        propertyInfos: [{
            name: 'item',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTListItem'
          }, {
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }]
      }, {
        localName: 'CTOlapSlicerCache',
        typeName: 'CT_OlapSlicerCache',
        propertyInfos: [{
            name: 'levels',
            required: true,
            typeInfo: '.CTOlapSlicerCacheLevelsData'
          }, {
            name: 'selections',
            required: true,
            typeInfo: '.CTOlapSlicerCacheSelections'
          }, {
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }, {
            name: 'pivotCacheId',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'pivotCacheId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOlapSlicerCacheItem',
        typeName: 'CT_OlapSlicerCacheItem',
        propertyInfos: [{
            name: 'p',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTOlapSlicerCacheItemParent'
          }, {
            name: 'c',
            attributeName: {
              localPart: 'c'
            },
            type: 'attribute'
          }, {
            name: 'n',
            required: true,
            attributeName: {
              localPart: 'n'
            },
            type: 'attribute'
          }, {
            name: 'nd',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'nd'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOlapSlicerCacheItemParent',
        typeName: 'CT_OlapSlicerCacheItemParent',
        propertyInfos: [{
            name: 'n',
            required: true,
            attributeName: {
              localPart: 'n'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOlapSlicerCacheLevelData',
        typeName: 'CT_OlapSlicerCacheLevelData',
        propertyInfos: [{
            name: 'ranges',
            typeInfo: '.CTOlapSlicerCacheRanges'
          }, {
            name: 'count',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'count'
            },
            type: 'attribute'
          }, {
            name: 'crossFilter',
            typeInfo: '.STSlicerCacheCrossFilter',
            defaultValue: 'showItemsWithDataAtTop',
            attributeName: {
              localPart: 'crossFilter'
            },
            type: 'attribute'
          }, {
            name: 'sortOrder',
            typeInfo: '.STOlapSlicerCacheSortOrder',
            defaultValue: 'natural',
            attributeName: {
              localPart: 'sortOrder'
            },
            type: 'attribute'
          }, {
            name: 'sourceCaption',
            attributeName: {
              localPart: 'sourceCaption'
            },
            type: 'attribute'
          }, {
            name: 'uniqueName',
            required: true,
            attributeName: {
              localPart: 'uniqueName'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOlapSlicerCacheLevelsData',
        typeName: 'CT_OlapSlicerCacheLevelsData',
        propertyInfos: [{
            name: 'level',
            required: true,
            collection: true,
            typeInfo: '.CTOlapSlicerCacheLevelData'
          }, {
            name: 'count',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'count'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOlapSlicerCacheRange',
        typeName: 'CT_OlapSlicerCacheRange',
        propertyInfos: [{
            name: 'i',
            required: true,
            collection: true,
            typeInfo: '.CTOlapSlicerCacheItem'
          }, {
            name: 'startItem',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'startItem'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOlapSlicerCacheRanges',
        typeName: 'CT_OlapSlicerCacheRanges',
        propertyInfos: [{
            name: 'range',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTOlapSlicerCacheRange'
          }]
      }, {
        localName: 'CTOlapSlicerCacheSelection',
        typeName: 'CT_OlapSlicerCacheSelection',
        propertyInfos: [{
            name: 'p',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTOlapSlicerCacheItemParent'
          }, {
            name: 'n',
            required: true,
            attributeName: {
              localPart: 'n'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOlapSlicerCacheSelections',
        typeName: 'CT_OlapSlicerCacheSelections',
        propertyInfos: [{
            name: 'selection',
            required: true,
            collection: true,
            typeInfo: '.CTOlapSlicerCacheSelection'
          }, {
            name: 'count',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'count'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTOleItem',
        typeName: 'CT_OleItem',
        propertyInfos: [{
            name: 'values',
            typeInfo: 'org_xlsx4j_sml.CTDdeValues'
          }, {
            name: 'advise',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'advise'
            },
            type: 'attribute'
          }, {
            name: 'icon',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'icon'
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
            name: 'preferPic',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'preferPic'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPivotCacheDefinition',
        typeName: 'CT_PivotCacheDefinition',
        propertyInfos: [{
            name: 'pivotCacheId',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'pivotCacheId'
            },
            type: 'attribute'
          }, {
            name: 'slicerData',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'slicerData'
            },
            type: 'attribute'
          }, {
            name: 'supportAddCalcMems',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'supportAddCalcMems'
            },
            type: 'attribute'
          }, {
            name: 'supportSubqueryCalcMem',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'supportSubqueryCalcMem'
            },
            type: 'attribute'
          }, {
            name: 'supportSubqueryNonVisual',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'supportSubqueryNonVisual'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPivotChange',
        typeName: 'CT_PivotChange',
        propertyInfos: [{
            name: 'editValue',
            required: true,
            typeInfo: '.CTPivotEditValue'
          }, {
            name: 'tupleItems',
            required: true,
            typeInfo: '.CTTupleItems'
          }, {
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }, {
            name: 'allocationMethod',
            typeInfo: '.STAllocationMethod',
            defaultValue: 'equalAllocation',
            attributeName: {
              localPart: 'allocationMethod'
            },
            type: 'attribute'
          }, {
            name: 'weightExpression',
            attributeName: {
              localPart: 'weightExpression'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPivotChanges',
        typeName: 'CT_PivotChanges',
        propertyInfos: [{
            name: 'pivotChange',
            required: true,
            collection: true,
            typeInfo: '.CTPivotChange'
          }]
      }, {
        localName: 'CTPivotEdit',
        typeName: 'CT_PivotEdit',
        propertyInfos: [{
            name: 'userEdit',
            required: true,
            typeInfo: '.CTPivotUserEdit'
          }, {
            name: 'tupleItems',
            required: true,
            typeInfo: '.CTTupleItems'
          }, {
            name: 'pivotArea',
            required: true,
            typeInfo: 'org_xlsx4j_sml.CTPivotArea'
          }, {
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }]
      }, {
        localName: 'CTPivotEditValue',
        typeName: 'CT_PivotEditValue',
        propertyInfos: [{
            name: 'value',
            type: 'value'
          }, {
            name: 'valueType',
            required: true,
            typeInfo: '.STPivotEditValueType',
            attributeName: {
              localPart: 'valueType'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPivotEdits',
        typeName: 'CT_PivotEdits',
        propertyInfos: [{
            name: 'pivotEdit',
            required: true,
            collection: true,
            typeInfo: '.CTPivotEdit'
          }]
      }, {
        localName: 'CTPivotField',
        typeName: 'CT_PivotField',
        propertyInfos: [{
            name: 'fillDownLabels',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'fillDownLabels'
            },
            type: 'attribute'
          }, {
            name: 'ignore',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'ignore'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPivotHierarchy',
        typeName: 'CT_PivotHierarchy',
        propertyInfos: [{
            name: 'ignore',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'ignore'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPivotTableDefinition',
        typeName: 'CT_PivotTableDefinition',
        propertyInfos: [{
            name: 'pivotEdits',
            typeInfo: '.CTPivotEdits'
          }, {
            name: 'pivotChanges',
            typeInfo: '.CTPivotChanges'
          }, {
            name: 'conditionalFormats',
            typeInfo: '.CTConditionalFormats'
          }, {
            name: 'allocationMethod',
            typeInfo: '.STAllocationMethod',
            defaultValue: 'equalAllocation',
            attributeName: {
              localPart: 'allocationMethod'
            },
            type: 'attribute'
          }, {
            name: 'altText',
            attributeName: {
              localPart: 'altText'
            },
            type: 'attribute'
          }, {
            name: 'altTextSummary',
            attributeName: {
              localPart: 'altTextSummary'
            },
            type: 'attribute'
          }, {
            name: 'autoApply',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'autoApply'
            },
            type: 'attribute'
          }, {
            name: 'calculatedMembersInFilters',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'calculatedMembersInFilters'
            },
            type: 'attribute'
          }, {
            name: 'enableEdit',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'enableEdit'
            },
            type: 'attribute'
          }, {
            name: 'fillDownLabelsDefault',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'fillDownLabelsDefault'
            },
            type: 'attribute'
          }, {
            name: 'hideValuesRow',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'hideValuesRow'
            },
            type: 'attribute'
          }, {
            name: 'visualTotalsForSets',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'visualTotalsForSets'
            },
            type: 'attribute'
          }, {
            name: 'weightExpression',
            attributeName: {
              localPart: 'weightExpression'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPivotUserEdit',
        typeName: 'CT_PivotUserEdit',
        propertyInfos: [{
            name: 'f',
            elementName: {
              localPart: 'f',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2006\/main'
            }
          }, {
            name: 'editValue',
            typeInfo: '.CTPivotEditValue'
          }]
      }, {
        localName: 'CTProtectedRange',
        typeName: 'CT_ProtectedRange',
        propertyInfos: [{
            name: 'sqref',
            required: true,
            elementName: {
              localPart: 'sqref',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2006\/main'
            },
            typeInfo: 'org_xlsx4j_schemas_microsoft_com_office_excel_2006_main.CTSqref'
          }, {
            name: 'algorithmName',
            attributeName: {
              localPart: 'algorithmName'
            },
            type: 'attribute'
          }, {
            name: 'hashValue',
            typeInfo: 'Base64Binary',
            attributeName: {
              localPart: 'hashValue'
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
            name: 'password',
            typeInfo: 'HexBinary',
            attributeName: {
              localPart: 'password'
            },
            type: 'attribute'
          }, {
            name: 'saltValue',
            typeInfo: 'Base64Binary',
            attributeName: {
              localPart: 'saltValue'
            },
            type: 'attribute'
          }, {
            name: 'securityDescriptor',
            attributeName: {
              localPart: 'securityDescriptor'
            },
            type: 'attribute'
          }, {
            name: 'spinCount',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'spinCount'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTProtectedRanges',
        typeName: 'CT_ProtectedRanges',
        propertyInfos: [{
            name: 'protectedRange',
            required: true,
            collection: true,
            typeInfo: '.CTProtectedRange'
          }]
      }, {
        localName: 'CTSetLevel',
        typeName: 'CT_SetLevel',
        propertyInfos: [{
            name: 'hierarchy',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'hierarchy'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSetLevels',
        typeName: 'CT_SetLevels',
        propertyInfos: [{
            name: 'setLevel',
            required: true,
            collection: true,
            typeInfo: '.CTSetLevel'
          }, {
            name: 'count',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'count'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSlicer',
        typeName: 'CT_Slicer',
        propertyInfos: [{
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }, {
            name: 'cache',
            required: true,
            attributeName: {
              localPart: 'cache'
            },
            type: 'attribute'
          }, {
            name: 'caption',
            attributeName: {
              localPart: 'caption'
            },
            type: 'attribute'
          }, {
            name: 'columnCount',
            typeInfo: 'UnsignedInt',
            defaultValue: 1,
            attributeName: {
              localPart: 'columnCount'
            },
            type: 'attribute'
          }, {
            name: 'level',
            typeInfo: 'UnsignedInt',
            defaultValue: 0,
            attributeName: {
              localPart: 'level'
            },
            type: 'attribute'
          }, {
            name: 'lockedPosition',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'lockedPosition'
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
            name: 'rowHeight',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'rowHeight'
            },
            type: 'attribute'
          }, {
            name: 'showCaption',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'showCaption'
            },
            type: 'attribute'
          }, {
            name: 'startItem',
            typeInfo: 'UnsignedInt',
            defaultValue: 0,
            attributeName: {
              localPart: 'startItem'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'uid',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'uid',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2016\/revision10'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSlicerCache',
        typeName: 'CT_SlicerCache',
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
        localName: 'CTSlicerCacheData',
        typeName: 'CT_SlicerCacheData',
        propertyInfos: [{
            name: 'olap',
            typeInfo: '.CTOlapSlicerCache'
          }, {
            name: 'tabular',
            typeInfo: '.CTTabularSlicerCache'
          }]
      }, {
        localName: 'CTSlicerCacheDefinition',
        typeName: 'CT_SlicerCacheDefinition',
        propertyInfos: [{
            name: 'pivotTables',
            typeInfo: '.CTSlicerCachePivotTables'
          }, {
            name: 'data',
            typeInfo: '.CTSlicerCacheData'
          }, {
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }, {
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'sourceName',
            required: true,
            attributeName: {
              localPart: 'sourceName'
            },
            type: 'attribute'
          }, {
            name: 'uid',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'uid',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2016\/revision10'
            },
            type: 'attribute'
          }, {
            name: 'ignorable',
            attributeName: {
              localPart: 'Ignorable',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/markup-compatibility\/2006'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSlicerCachePivotTable',
        typeName: 'CT_SlicerCachePivotTable',
        propertyInfos: [{
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'tabId',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'tabId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSlicerCachePivotTables',
        typeName: 'CT_SlicerCachePivotTables',
        propertyInfos: [{
            name: 'pivotTable',
            required: true,
            collection: true,
            typeInfo: '.CTSlicerCachePivotTable'
          }]
      }, {
        localName: 'CTSlicerCaches',
        typeName: 'CT_SlicerCaches',
        propertyInfos: [{
            name: 'slicerCache',
            required: true,
            collection: true,
            typeInfo: '.CTSlicerCache'
          }]
      }, {
        localName: 'CTSlicerRef',
        typeName: 'CT_SlicerRef',
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
        localName: 'CTSlicerRefs',
        typeName: 'CT_SlicerRefs',
        propertyInfos: [{
            name: 'slicer',
            required: true,
            collection: true,
            typeInfo: '.CTSlicerRef'
          }]
      }, {
        localName: 'CTSlicerStyle',
        typeName: 'CT_SlicerStyle',
        propertyInfos: [{
            name: 'slicerStyleElements',
            typeInfo: '.CTSlicerStyleElements'
          }, {
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSlicerStyleElement',
        typeName: 'CT_SlicerStyleElement',
        propertyInfos: [{
            name: 'dxfId',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'dxfId'
            },
            type: 'attribute'
          }, {
            name: 'type',
            required: true,
            typeInfo: '.STSlicerStyleType',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSlicerStyleElements',
        typeName: 'CT_SlicerStyleElements',
        propertyInfos: [{
            name: 'slicerStyleElement',
            required: true,
            collection: true,
            typeInfo: '.CTSlicerStyleElement'
          }]
      }, {
        localName: 'CTSlicerStyles',
        typeName: 'CT_SlicerStyles',
        propertyInfos: [{
            name: 'slicerStyle',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTSlicerStyle'
          }, {
            name: 'defaultSlicerStyle',
            required: true,
            attributeName: {
              localPart: 'defaultSlicerStyle'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSlicers',
        typeName: 'CT_Slicers',
        propertyInfos: [{
            name: 'slicer',
            required: true,
            collection: true,
            typeInfo: '.CTSlicer'
          }, {
            name: 'ignorable',
            attributeName: {
              localPart: 'Ignorable',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/markup-compatibility\/2006'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSortCondition',
        typeName: 'CT_SortCondition',
        propertyInfos: [{
            name: 'customList',
            attributeName: {
              localPart: 'customList'
            },
            type: 'attribute'
          }, {
            name: 'descending',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'descending'
            },
            type: 'attribute'
          }, {
            name: 'dxfId',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'dxfId'
            },
            type: 'attribute'
          }, {
            name: 'iconId',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'iconId'
            },
            type: 'attribute'
          }, {
            name: 'iconSet',
            values: ['3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters', '3Stars', '3Triangles', '5Boxes', 'NoIcons'],
            defaultValue: '3Arrows',
            attributeName: {
              localPart: 'iconSet'
            },
            type: 'attribute'
          }, {
            name: 'ref',
            required: true,
            attributeName: {
              localPart: 'ref'
            },
            type: 'attribute'
          }, {
            name: 'sortBy',
            typeInfo: 'org_xlsx4j_sml.STSortBy',
            defaultValue: 'value',
            attributeName: {
              localPart: 'sortBy'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSourceConnection',
        typeName: 'CT_SourceConnection',
        propertyInfos: [{
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSparkline',
        typeName: 'CT_Sparkline',
        propertyInfos: [{
            name: 'f',
            elementName: {
              localPart: 'f',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2006\/main'
            }
          }, {
            name: 'sqref',
            required: true,
            elementName: {
              localPart: 'sqref',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2006\/main'
            },
            typeInfo: 'org_xlsx4j_schemas_microsoft_com_office_excel_2006_main.CTSqref'
          }]
      }, {
        localName: 'CTSparklineGroup',
        typeName: 'CT_SparklineGroup',
        propertyInfos: [{
            name: 'colorSeries',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'colorNegative',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'colorAxis',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'colorMarkers',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'colorFirst',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'colorLast',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'colorHigh',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'colorLow',
            typeInfo: 'org_xlsx4j_sml.CTColor'
          }, {
            name: 'f',
            elementName: {
              localPart: 'f',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/excel\/2006\/main'
            }
          }, {
            name: 'sparklines',
            required: true,
            typeInfo: '.CTSparklines'
          }, {
            name: 'dateAxis',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'dateAxis'
            },
            type: 'attribute'
          }, {
            name: 'displayEmptyCellsAs',
            typeInfo: '.STDispBlanksAs',
            defaultValue: 'zero',
            attributeName: {
              localPart: 'displayEmptyCellsAs'
            },
            type: 'attribute'
          }, {
            name: 'displayHidden',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'displayHidden'
            },
            type: 'attribute'
          }, {
            name: 'displayXAxis',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'displayXAxis'
            },
            type: 'attribute'
          }, {
            name: 'first',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'first'
            },
            type: 'attribute'
          }, {
            name: 'high',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'high'
            },
            type: 'attribute'
          }, {
            name: 'last',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'last'
            },
            type: 'attribute'
          }, {
            name: 'lineWeight',
            typeInfo: 'Double',
            defaultValue: 0.75,
            attributeName: {
              localPart: 'lineWeight'
            },
            type: 'attribute'
          }, {
            name: 'low',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'low'
            },
            type: 'attribute'
          }, {
            name: 'manualMax',
            typeInfo: 'Double',
            attributeName: {
              localPart: 'manualMax'
            },
            type: 'attribute'
          }, {
            name: 'manualMin',
            typeInfo: 'Double',
            attributeName: {
              localPart: 'manualMin'
            },
            type: 'attribute'
          }, {
            name: 'markers',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'markers'
            },
            type: 'attribute'
          }, {
            name: 'maxAxisType',
            typeInfo: '.STSparklineAxisMinMax',
            defaultValue: 'individual',
            attributeName: {
              localPart: 'maxAxisType'
            },
            type: 'attribute'
          }, {
            name: 'minAxisType',
            typeInfo: '.STSparklineAxisMinMax',
            defaultValue: 'individual',
            attributeName: {
              localPart: 'minAxisType'
            },
            type: 'attribute'
          }, {
            name: 'negative',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'negative'
            },
            type: 'attribute'
          }, {
            name: 'rightToLeft',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'rightToLeft'
            },
            type: 'attribute'
          }, {
            name: 'type',
            typeInfo: '.STSparklineType',
            defaultValue: 'line',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: 'uid',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'uid',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2015\/revision2'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSparklineGroups',
        typeName: 'CT_SparklineGroups',
        propertyInfos: [{
            name: 'sparklineGroup',
            required: true,
            collection: true,
            typeInfo: '.CTSparklineGroup'
          }]
      }, {
        localName: 'CTSparklines',
        typeName: 'CT_Sparklines',
        propertyInfos: [{
            name: 'sparkline',
            required: true,
            collection: true,
            typeInfo: '.CTSparkline'
          }]
      }, {
        localName: 'CTTable',
        typeName: 'CT_Table',
        propertyInfos: [{
            name: 'altText',
            attributeName: {
              localPart: 'altText'
            },
            type: 'attribute'
          }, {
            name: 'altTextSummary',
            attributeName: {
              localPart: 'altTextSummary'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTTabularSlicerCache',
        typeName: 'CT_TabularSlicerCache',
        propertyInfos: [{
            name: 'items',
            typeInfo: '.CTTabularSlicerCacheItems'
          }, {
            name: 'extLst',
            typeInfo: 'org_xlsx4j_sml.CTExtensionList'
          }, {
            name: 'crossFilter',
            typeInfo: '.STSlicerCacheCrossFilter',
            defaultValue: 'showItemsWithDataAtTop',
            attributeName: {
              localPart: 'crossFilter'
            },
            type: 'attribute'
          }, {
            name: 'customListSort',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'customListSort'
            },
            type: 'attribute'
          }, {
            name: 'pivotCacheId',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'pivotCacheId'
            },
            type: 'attribute'
          }, {
            name: 'showMissing',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'showMissing'
            },
            type: 'attribute'
          }, {
            name: 'sortOrder',
            typeInfo: '.STTabularSlicerCacheSortOrder',
            defaultValue: 'ascending',
            attributeName: {
              localPart: 'sortOrder'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTTabularSlicerCacheItem',
        typeName: 'CT_TabularSlicerCacheItem',
        propertyInfos: [{
            name: 'nd',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'nd'
            },
            type: 'attribute'
          }, {
            name: 's',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 's'
            },
            type: 'attribute'
          }, {
            name: 'x',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'x'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTTabularSlicerCacheItems',
        typeName: 'CT_TabularSlicerCacheItems',
        propertyInfos: [{
            name: 'i',
            required: true,
            collection: true,
            typeInfo: '.CTTabularSlicerCacheItem'
          }, {
            name: 'count',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'count'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTTupleItems',
        typeName: 'CT_TupleItems',
        propertyInfos: [{
            name: 'tupleItem',
            required: true,
            collection: true
          }]
      }, {
        localName: 'CTTupleSet',
        typeName: 'CT_TupleSet',
        propertyInfos: [{
            name: 'headers',
            required: true,
            typeInfo: '.CTTupleSetHeaders'
          }, {
            name: 'rows',
            required: true,
            typeInfo: '.CTTupleSetRows'
          }, {
            name: 'columnCount',
            typeInfo: 'UnsignedInt',
            defaultValue: 1,
            attributeName: {
              localPart: 'columnCount'
            },
            type: 'attribute'
          }, {
            name: 'rowCount',
            typeInfo: 'UnsignedInt',
            defaultValue: 1,
            attributeName: {
              localPart: 'rowCount'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTTupleSetHeader',
        typeName: 'CT_TupleSetHeader',
        propertyInfos: [{
            name: 'hierarchyName',
            attributeName: {
              localPart: 'hierarchyName'
            },
            type: 'attribute'
          }, {
            name: 'uniqueName',
            attributeName: {
              localPart: 'uniqueName'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTTupleSetHeaders',
        typeName: 'CT_TupleSetHeaders',
        propertyInfos: [{
            name: 'header',
            required: true,
            collection: true,
            typeInfo: '.CTTupleSetHeader'
          }]
      }, {
        localName: 'CTTupleSetRow',
        typeName: 'CT_TupleSetRow',
        propertyInfos: [{
            name: 'rowItem',
            required: true,
            collection: true,
            typeInfo: '.CTTupleSetRowItem'
          }]
      }, {
        localName: 'CTTupleSetRowItem',
        typeName: 'CT_TupleSetRowItem',
        propertyInfos: [{
            name: 'd',
            attributeName: {
              localPart: 'd'
            },
            type: 'attribute'
          }, {
            name: 'u',
            attributeName: {
              localPart: 'u'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTTupleSetRows',
        typeName: 'CT_TupleSetRows',
        propertyInfos: [{
            name: 'row',
            required: true,
            collection: true,
            typeInfo: '.CTTupleSetRow'
          }]
      }, {
        localName: 'CTWorkbookPr',
        typeName: 'CT_WorkbookPr',
        propertyInfos: [{
            name: 'accuracyVersion',
            typeInfo: 'UnsignedInt',
            defaultValue: 0,
            attributeName: {
              localPart: 'accuracyVersion'
            },
            type: 'attribute'
          }, {
            name: 'defaultImageDpi',
            typeInfo: 'UnsignedInt',
            defaultValue: 220,
            attributeName: {
              localPart: 'defaultImageDpi'
            },
            type: 'attribute'
          }, {
            name: 'discardImageEditData',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'discardImageEditData'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STAllocationMethod',
        values: ['equalAllocation', 'equalIncrement', 'weightedAllocation', 'weightedIncrement']
      }, {
        type: 'enumInfo',
        localName: 'STCfvoType',
        values: ['num', 'percent', 'max', 'min', 'formula', 'percentile', 'autoMin', 'autoMax']
      }, {
        type: 'enumInfo',
        localName: 'STChecked',
        baseTypeInfo: 'Token',
        values: ['Unchecked', 'Checked', 'Mixed']
      }, {
        type: 'enumInfo',
        localName: 'STDataBarAxisPosition',
        values: ['automatic', 'middle', 'none']
      }, {
        type: 'enumInfo',
        localName: 'STDataBarDirection',
        values: ['context', 'leftToRight', 'rightToLeft']
      }, {
        type: 'enumInfo',
        localName: 'STDispBlanksAs',
        values: ['span', 'gap', 'zero']
      }, {
        type: 'enumInfo',
        localName: 'STDropStyle',
        baseTypeInfo: 'Token',
        values: ['combo', 'comboedit', 'simple']
      }, {
        type: 'enumInfo',
        localName: 'STEditValidation',
        baseTypeInfo: 'Token',
        values: ['text', 'integer', 'number', 'reference', 'formula']
      }, {
        type: 'enumInfo',
        localName: 'STObjectType',
        baseTypeInfo: 'Token',
        values: ['Button', 'CheckBox', 'Drop', 'GBox', 'Label', 'List', 'Radio', 'Scroll', 'Spin', 'EditBox', 'Dialog']
      }, {
        type: 'enumInfo',
        localName: 'STOlapSlicerCacheSortOrder',
        values: ['natural', 'ascending', 'descending']
      }, {
        type: 'enumInfo',
        localName: 'STPivotEditValueType',
        values: ['number', 'dateTime', 'string', 'boolean', 'error']
      }, {
        type: 'enumInfo',
        localName: 'STPivotShowAs',
        values: ['percentOfParent', 'percentOfParentRow', 'percentOfParentCol', 'percentOfRunningTotal', 'rankAscending', 'rankDescending']
      }, {
        type: 'enumInfo',
        localName: 'STSelType',
        baseTypeInfo: 'Token',
        values: ['single', 'multi', 'extended']
      }, {
        type: 'enumInfo',
        localName: 'STSlicerCacheCrossFilter',
        values: ['none', 'showItemsWithDataAtTop', 'showItemsWithNoData']
      }, {
        type: 'enumInfo',
        localName: 'STSlicerStyleType',
        values: ['unselectedItemWithData', 'selectedItemWithData', 'unselectedItemWithNoData', 'selectedItemWithNoData', 'hoveredUnselectedItemWithData', 'hoveredSelectedItemWithData', 'hoveredUnselectedItemWithNoData', 'hoveredSelectedItemWithNoData']
      }, {
        type: 'enumInfo',
        localName: 'STSparklineAxisMinMax',
        values: ['individual', 'group', 'custom']
      }, {
        type: 'enumInfo',
        localName: 'STSparklineType',
        values: ['line', 'column', 'stacked']
      }, {
        type: 'enumInfo',
        localName: 'STTabularSlicerCacheSortOrder',
        values: ['ascending', 'descending']
      }, {
        type: 'enumInfo',
        localName: 'STTextHAlign',
        values: ['left', 'center', 'right', 'justify', 'distributed']
      }, {
        type: 'enumInfo',
        localName: 'STTextVAlign',
        values: ['top', 'center', 'bottom', 'justify', 'distributed']
      }],
    elementInfos: [{
        typeInfo: '.CTCacheField',
        elementName: 'cacheField'
      }, {
        typeInfo: '.CTCacheHierarchy',
        elementName: 'cacheHierarchy'
      }, {
        typeInfo: '.CTCalculatedMember',
        elementName: 'calculatedMember'
      }, {
        typeInfo: '.CTConditionalFormattings',
        elementName: 'conditionalFormattings'
      }, {
        typeInfo: '.CTConnection',
        elementName: 'connection'
      }, {
        typeInfo: '.CTCustomFilters',
        elementName: 'customFilters'
      }, {
        typeInfo: '.CTDataField',
        elementName: 'dataField'
      }, {
        typeInfo: '.CTDataValidations',
        elementName: 'dataValidations'
      }, {
        typeInfo: '.CTDatastoreItem',
        elementName: 'datastoreItem'
      }, {
        typeInfo: 'org_xlsx4j_sml.CTDxfs',
        elementName: 'dxfs'
      }, {
        typeInfo: '.CTFilter',
        elementName: 'filter'
      }, {
        typeInfo: '.CTFormControlPr',
        elementName: 'formControlPr'
      }, {
        typeInfo: '.CTIconFilter',
        elementName: 'iconFilter'
      }, {
        typeInfo: 'Token',
        elementName: 'id'
      }, {
        typeInfo: '.CTIgnoredErrors',
        elementName: 'ignoredErrors'
      }, {
        typeInfo: '.CTOleItem',
        elementName: 'oleItem'
      }, {
        typeInfo: '.CTPivotCacheDefinition',
        elementName: 'pivotCacheDefinition'
      }, {
        typeInfo: 'org_xlsx4j_sml.CTPivotCaches',
        elementName: 'pivotCaches'
      }, {
        typeInfo: '.CTPivotField',
        elementName: 'pivotField'
      }, {
        typeInfo: '.CTPivotHierarchy',
        elementName: 'pivotHierarchy'
      }, {
        typeInfo: '.CTPivotTableDefinition',
        elementName: 'pivotTableDefinition'
      }, {
        typeInfo: '.CTProtectedRanges',
        elementName: 'protectedRanges'
      }, {
        typeInfo: '.CTSlicerCacheDefinition',
        elementName: 'slicerCacheDefinition'
      }, {
        typeInfo: '.CTSlicerCaches',
        elementName: 'slicerCaches'
      }, {
        typeInfo: '.CTSlicerRefs',
        elementName: 'slicerList'
      }, {
        typeInfo: '.CTSlicerStyles',
        elementName: 'slicerStyles'
      }, {
        typeInfo: '.CTSlicers',
        elementName: 'slicers'
      }, {
        typeInfo: '.CTSortCondition',
        elementName: 'sortCondition'
      }, {
        typeInfo: '.CTSourceConnection',
        elementName: 'sourceConnection'
      }, {
        typeInfo: '.CTSparklineGroups',
        elementName: 'sparklineGroups'
      }, {
        typeInfo: '.CTTable',
        elementName: 'table'
      }, {
        typeInfo: '.CTWorkbookPr',
        elementName: 'workbookPr'
      }]
  };
  return {
    org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main: org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main_Module_Factory);
}
else {
  var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main_Module = org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main = org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main_Module.org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main;
  }
  else {
    var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main = org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main_Module.org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2009_x9_main;
  }
}