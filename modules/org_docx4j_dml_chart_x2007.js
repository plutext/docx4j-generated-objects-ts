var org_docx4j_dml_chart_x2007_Module_Factory = function () {
  var org_docx4j_dml_chart_x2007 = {
    name: 'org_docx4j_dml_chart_x2007',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2007\/8\/2\/chart',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTBooleanFalse',
        typeName: 'CT_BooleanFalse',
        propertyInfos: [{
            name: 'val',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTBooleanTrue',
        typeName: 'CT_BooleanTrue',
        propertyInfos: [{
            name: 'val',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTInvertSolidFillFmt',
        typeName: 'CT_InvertSolidFillFmt',
        propertyInfos: [{
            name: 'spPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTShapeProperties'
          }]
      }, {
        localName: 'CTPivotOptions',
        typeName: 'CT_PivotOptions',
        propertyInfos: [{
            name: 'dropZoneFilter',
            typeInfo: '.CTBooleanFalse'
          }, {
            name: 'dropZoneCategories',
            typeInfo: '.CTBooleanFalse'
          }, {
            name: 'dropZoneData',
            typeInfo: '.CTBooleanFalse'
          }, {
            name: 'dropZoneSeries',
            typeInfo: '.CTBooleanFalse'
          }, {
            name: 'dropZonesVisible',
            typeInfo: '.CTBooleanFalse'
          }]
      }, {
        localName: 'CTStyle',
        typeName: 'CT_Style',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'UnsignedByte',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTInvertSolidFillFmt',
        elementName: 'invertSolidFillFmt'
      }, {
        typeInfo: '.CTPivotOptions',
        elementName: 'pivotOptions'
      }, {
        typeInfo: '.CTStyle',
        elementName: 'style'
      }]
  };
  return {
    org_docx4j_dml_chart_x2007: org_docx4j_dml_chart_x2007
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_dml_chart_x2007_Module_Factory);
}
else {
  var org_docx4j_dml_chart_x2007_Module = org_docx4j_dml_chart_x2007_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_dml_chart_x2007 = org_docx4j_dml_chart_x2007_Module.org_docx4j_dml_chart_x2007;
  }
  else {
    var org_docx4j_dml_chart_x2007 = org_docx4j_dml_chart_x2007_Module.org_docx4j_dml_chart_x2007;
  }
}