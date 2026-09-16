var org_docx4j_customXmlProperties_Module_Factory = function () {
  var org_docx4j_customXmlProperties = {
    name: 'org_docx4j_customXmlProperties',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/customXml',
    defaultAttributeNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/customXml',
    typeInfos: [{
        localName: 'DatastoreItem',
        typeName: null,
        propertyInfos: [{
            name: 'otherAttributes',
            type: 'anyAttribute'
          }, {
            name: 'schemaRefs',
            typeInfo: '.SchemaRefs'
          }, {
            name: 'itemID',
            required: true,
            type: 'attribute'
          }]
      }, {
        localName: 'SchemaRefs',
        typeName: null,
        propertyInfos: [{
            name: 'schemaRef',
            minOccurs: 0,
            collection: true,
            typeInfo: '.SchemaRefs.SchemaRef'
          }]
      }, {
        localName: 'SchemaRefs.SchemaRef',
        typeName: null,
        propertyInfos: [{
            name: 'uri',
            required: true,
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.DatastoreItem',
        elementName: 'datastoreItem'
      }, {
        typeInfo: '.SchemaRefs',
        elementName: 'schemaRefs'
      }]
  };
  return {
    org_docx4j_customXmlProperties: org_docx4j_customXmlProperties
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_customXmlProperties_Module_Factory);
}
else {
  var org_docx4j_customXmlProperties_Module = org_docx4j_customXmlProperties_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_customXmlProperties = org_docx4j_customXmlProperties_Module.org_docx4j_customXmlProperties;
  }
  else {
    var org_docx4j_customXmlProperties = org_docx4j_customXmlProperties_Module.org_docx4j_customXmlProperties;
  }
}