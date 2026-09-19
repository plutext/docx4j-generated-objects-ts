var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision_Module_Factory = function () {
  var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision = {
    name: 'org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2014\/revision',
    typeInfos: [{
        localName: 'CTRevisionPtr',
        typeName: 'CT_RevisionPtr',
        propertyInfos: [{
            name: 'documentId',
            required: true,
            attributeName: {
              localPart: 'documentId'
            },
            type: 'attribute'
          }, {
            name: 'revIDLastSave',
            required: true,
            typeInfo: 'UnsignedLong',
            attributeName: {
              localPart: 'revIDLastSave'
            },
            type: 'attribute'
          }, {
            name: 'uidLastSave',
            typeInfo: 'Token',
            attributeName: {
              localPart: 'uidLastSave',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2016\/revision10'
            },
            type: 'attribute'
          }, {
            name: 'coauthVersionLast',
            typeInfo: 'UnsignedInt',
            defaultValue: 0,
            attributeName: {
              localPart: 'coauthVersionLast',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2016\/revision6'
            },
            type: 'attribute'
          }, {
            name: 'coauthVersionMax',
            typeInfo: 'UnsignedInt',
            defaultValue: 0,
            attributeName: {
              localPart: 'coauthVersionMax',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2016\/revision6'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTRevisionPtr',
        elementName: 'revisionPtr'
      }]
  };
  return {
    org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision: org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision_Module_Factory);
}
else {
  var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision_Module = org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision = org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision_Module.org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision;
  }
  else {
    var org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision = org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision_Module.org_xlsx4j_com_microsoft_schemas_office_spreadsheetml_x2014_revision;
  }
}