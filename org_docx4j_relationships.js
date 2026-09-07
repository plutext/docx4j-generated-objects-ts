var org_docx4j_relationships_Module_Factory = function () {
  var org_docx4j_relationships = {
    name: 'org_docx4j_relationships',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/package\/2006\/relationships',
    typeInfos: [{
        localName: 'Relationship',
        typeName: null,
        propertyInfos: [{
            name: 'id',
            required: true,
            typeInfo: 'ID',
            attributeName: {
              localPart: 'Id'
            },
            type: 'attribute'
          }, {
            name: 'target',
            required: true,
            attributeName: {
              localPart: 'Target'
            },
            type: 'attribute'
          }, {
            name: 'targetMode',
            values: ['External', 'Internal'],
            attributeName: {
              localPart: 'TargetMode'
            },
            type: 'attribute'
          }, {
            name: 'type',
            required: true,
            attributeName: {
              localPart: 'Type'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'Relationships',
        typeName: null,
        propertyInfos: [{
            name: 'relationship',
            minOccurs: 0,
            collection: true,
            elementName: 'Relationship',
            typeInfo: '.Relationship'
          }]
      }],
    elementInfos: [{
        typeInfo: '.Relationship',
        elementName: 'Relationship'
      }, {
        typeInfo: '.Relationships',
        elementName: 'Relationships'
      }]
  };
  return {
    org_docx4j_relationships: org_docx4j_relationships
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_relationships_Module_Factory);
}
else {
  var org_docx4j_relationships_Module = org_docx4j_relationships_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_relationships = org_docx4j_relationships_Module.org_docx4j_relationships;
  }
  else {
    var org_docx4j_relationships = org_docx4j_relationships_Module.org_docx4j_relationships;
  }
}