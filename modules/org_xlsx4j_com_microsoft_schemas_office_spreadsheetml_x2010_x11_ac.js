var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac_Module_Factory = function () {
  var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac = {
    name: 'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2010\/11\/ac',
    typeInfos: [{
        localName: 'CTAbsolutePath',
        typeName: 'CT_AbsolutePath',
        propertyInfos: [{
            name: 'url',
            required: true,
            attributeName: {
              localPart: 'url'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTAbsolutePath',
        elementName: 'absPath'
      }]
  };
  return {
    org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac: org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac_Module_Factory);
}
else {
  var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac_Module = org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac = org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac_Module.org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac;
  }
  else {
    var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac = org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac_Module.org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2010_x11_ac;
  }
}