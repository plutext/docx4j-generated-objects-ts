var org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction = {
    name: 'org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/powerpoint\/2014\/inkAction',
    dependencies: ['org_docx4j_org_w3_x2003_inkML'],
    typeInfos: [{
        localName: 'CTAction',
        typeName: 'CT_Action',
        propertyInfos: [{
            name: 'property',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTActionProperty'
          }, {
            name: 'actionDataOrActionDataGroup',
            minOccurs: 0,
            collection: true,
            elementTypeInfos: [{
                elementName: 'actionData',
                typeInfo: '.CTActionData'
              }, {
                elementName: 'actionDataGroup',
                typeInfo: '.CTActionDataGroup'
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
            name: 'type',
            required: true,
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: 'startTime',
            required: true,
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'startTime'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTActionData',
        typeName: 'CT_ActionData',
        propertyInfos: [{
            name: 'transform',
            typeInfo: 'org_docx4j_org_w3_x2003_inkML.CTMatrix'
          }, {
            name: 'traceOrTraceView',
            minOccurs: 0,
            collection: true,
            elementTypeInfos: [{
                elementName: {
                  localPart: 'trace',
                  namespaceURI: 'http:\/\/www.w3.org\/2003\/InkML'
                },
                typeInfo: 'org_docx4j_org_w3_x2003_inkML.TraceType'
              }, {
                elementName: {
                  localPart: 'traceView',
                  namespaceURI: 'http:\/\/www.w3.org\/2003\/InkML'
                },
                typeInfo: 'org_docx4j_org_w3_x2003_inkML.TraceViewType'
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
            name: 'name',
            defaultValue: 'stroke',
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'ref',
            defaultValue: '',
            attributeName: {
              localPart: 'ref'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTActionDataGroup',
        typeName: 'CT_ActionDataGroup',
        propertyInfos: [{
            name: 'actionData',
            required: true,
            collection: true,
            typeInfo: '.CTActionData'
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
            defaultValue: 'stroke',
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTActionGroup',
        typeName: 'CT_ActionGroup',
        propertyInfos: [{
            name: 'action',
            required: true,
            collection: true,
            typeInfo: '.CTAction'
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
            required: true,
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: 'startTime',
            required: true,
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'startTime'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTActionProperty',
        typeName: 'CT_ActionProperty',
        propertyInfos: [{
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'value',
            defaultValue: 'ink',
            attributeName: {
              localPart: 'value'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTActions',
        typeName: 'CT_Actions',
        propertyInfos: [{
            name: 'definitions',
            elementName: {
              localPart: 'definitions',
              namespaceURI: 'http:\/\/www.w3.org\/2003\/InkML'
            },
            typeInfo: 'org_docx4j_org_w3_x2003_inkML.DefinitionsType'
          }, {
            name: 'actionGroupOrAction',
            minOccurs: 0,
            collection: true,
            elementTypeInfos: [{
                elementName: 'actionGroup',
                typeInfo: '.CTActionGroup'
              }, {
                elementName: 'action',
                typeInfo: '.CTAction'
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
            name: 'lengthUnit',
            required: true,
            values: ['m', 'cm', 'mm', 'in', 'pt', 'pc', 'em', 'ex', '1\/m', '1\/cm', '1\/mm', '1\/in', '1\/pt', '1\/pc', '1\/em', '1\/ex'],
            attributeName: {
              localPart: 'lengthUnit'
            },
            type: 'attribute'
          }, {
            name: 'timeUnit',
            required: true,
            values: ['s', 'ms', '1\/s', '1\/ms'],
            attributeName: {
              localPart: 'timeUnit'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STActionTypeReserved',
        values: ['add', 'remove', 'transform']
      }, {
        type: 'enumInfo',
        localName: 'STDataNameReserved',
        values: ['stroke', 'path', 'target']
      }, {
        type: 'enumInfo',
        localName: 'STPropertyNameReserved',
        values: ['dataType', 'style']
      }, {
        type: 'enumInfo',
        localName: 'STPropertyValueReserved',
        values: ['ink', 'pointEraser', 'strokeEraser', 'instant']
      }],
    elementInfos: [{
        typeInfo: '.CTActions',
        elementName: 'actions'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction: org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction_Module = org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction = org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction_Module.org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction = org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction_Module.org_docx4j_com_microsoft_schemas_office_powerpoint_x2014_inkAction;
  }
}