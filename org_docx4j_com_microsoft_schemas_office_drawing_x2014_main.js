var org_docx4j_com_microsoft_schemas_office_drawing_x2014_main_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2014_main = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2014_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2014\/main',
    typeInfos: [{
        localName: 'CTConnectableReferences',
        typeName: 'CT_ConnectableReferences',
        propertyInfos: [{
            name: 'st',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'st'
            },
            type: 'attribute'
          }, {
            name: 'end',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'end'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCreationId',
        typeName: 'CT_CreationId',
        propertyInfos: [{
            name: 'id',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTIdentifier',
        typeName: 'CT_Identifier',
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
        localName: 'CTPredecessorDrawingElementReference',
        typeName: 'CT_PredecessorDrawingElementReference',
        propertyInfos: [{
            name: 'pred',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'pred'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTIdentifier',
        elementName: 'colId'
      }, {
        typeInfo: '.CTCreationId',
        elementName: 'creationId'
      }, {
        typeInfo: '.CTConnectableReferences',
        elementName: 'cxnDERefs'
      }, {
        typeInfo: '.CTPredecessorDrawingElementReference',
        elementName: 'predDERef'
      }, {
        typeInfo: '.CTIdentifier',
        elementName: 'rowId'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2014_main: org_docx4j_com_microsoft_schemas_office_drawing_x2014_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2014_main_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2014_main_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2014_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2014_main = org_docx4j_com_microsoft_schemas_office_drawing_x2014_main_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2014_main;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2014_main = org_docx4j_com_microsoft_schemas_office_drawing_x2014_main_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2014_main;
  }
}