var org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/main',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTDiagramAutoBullet',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2016\/11\/diagram',
          localPart: 'CT_DiagramAutoBullet'
        },
        propertyInfos: [{
            name: 'buNone',
            required: true,
            typeInfo: 'org_docx4j_dml.CTTextNoBullet'
          }, {
            name: 'buAutoNum',
            required: true,
            typeInfo: 'org_docx4j_dml.CTTextAutonumberBullet'
          }, {
            name: 'buChar',
            required: true,
            typeInfo: 'org_docx4j_dml.CTTextCharBullet'
          }, {
            name: 'buBlip',
            required: true,
            typeInfo: 'org_docx4j_dml.CTTextBlipBullet'
          }, {
            name: 'prefix',
            attributeName: {
              localPart: 'prefix'
            },
            type: 'attribute'
          }, {
            name: 'leadZeros',
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'leadZeros'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTNumberDiagramInfo',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2016\/11\/diagram',
          localPart: 'CT_NumberDiagramInfo'
        },
        propertyInfos: [{
            name: 'buPr',
            required: true,
            elementName: {
              localPart: 'buPr',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2016\/11\/diagram'
            },
            typeInfo: '.CTDiagramAutoBullet'
          }, {
            name: 'lvl',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'lvl'
            },
            type: 'attribute'
          }, {
            name: 'ptType',
            required: true,
            typeInfo: '.STSTorageType',
            attributeName: {
              localPart: 'ptType'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTNumberDiagramInfoList',
        typeName: {
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2016\/11\/diagram',
          localPart: 'CT_NumberDiagramInfoList'
        },
        propertyInfos: [{
            name: 'autoBuNodeInfo',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'autoBuNodeInfo',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2016\/11\/diagram'
            },
            typeInfo: '.CTNumberDiagramInfo'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STSTorageType',
        baseTypeInfo: 'Token',
        values: ['sibTrans', 'parTrans']
      }],
    elementInfos: [{
        typeInfo: '.CTNumberDiagramInfoList',
        elementName: {
          localPart: 'autoBuNodeInfoLst',
          namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2016\/11\/diagram'
        }
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram: org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram_Module = org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram = org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram_Module.org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram = org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram_Module.org_docx4j_com_microsoft_schemas_office_drawing_x201611_diagram;
  }
}