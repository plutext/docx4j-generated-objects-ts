var org_docx4j_customxml_Module_Factory = function () {
  var org_docx4j_customxml = {
    name: 'org_docx4j_customxml',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/schemaLibrary\/2006\/main',
    defaultAttributeNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/schemaLibrary\/2006\/main',
    typeInfos: [{
        localName: 'SchemaLibrary',
        typeName: null,
        propertyInfos: [{
            name: 'schema',
            minOccurs: 0,
            collection: true,
            typeInfo: '.SchemaLibrary.Schema'
          }]
      }, {
        localName: 'SchemaLibrary.Schema',
        typeName: null,
        propertyInfos: [{
            name: 'uri',
            defaultValue: '',
            type: 'attribute'
          }, {
            name: 'manifestLocation',
            type: 'attribute'
          }, {
            name: 'schemaLocation',
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.SchemaLibrary',
        elementName: 'schemaLibrary'
      }]
  };
  return {
    org_docx4j_customxml: org_docx4j_customxml
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_customxml_Module_Factory);
}
else {
  var org_docx4j_customxml_Module = org_docx4j_customxml_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_customxml = org_docx4j_customxml_Module.org_docx4j_customxml;
  }
  else {
    var org_docx4j_customxml = org_docx4j_customxml_Module.org_docx4j_customxml;
  }
}