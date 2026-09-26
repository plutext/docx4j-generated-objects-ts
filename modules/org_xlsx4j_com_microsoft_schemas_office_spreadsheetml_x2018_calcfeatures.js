var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures_Module_Factory = function () {
  var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures = {
    name: 'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2018\/calcfeatures',
    typeInfos: [{
        localName: 'CTCalcFeature',
        typeName: 'CT_CalcFeature',
        propertyInfos: [{
            name: 'name',
            required: true,
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCalcFeatures',
        typeName: 'CT_CalcFeatures',
        propertyInfos: [{
            name: 'feature',
            required: true,
            collection: true,
            typeInfo: '.CTCalcFeature'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTCalcFeatures',
        elementName: 'calcFeatures'
      }]
  };
  return {
    org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures: org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures_Module_Factory);
}
else {
  var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures_Module = org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures = org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures_Module.org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures;
  }
  else {
    var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures = org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures_Module.org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2018_calcfeatures;
  }
}