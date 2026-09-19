var org_docx4j_mce_Module_Factory = function () {
  var org_docx4j_mce = {
    name: 'org_docx4j_mce',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/markup-compatibility\/2006',
    defaultAttributeNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/markup-compatibility\/2006',
    typeInfos: [{
        localName: 'AlternateContent',
        typeName: null,
        propertyInfos: [{
            name: 'choice',
            minOccurs: 0,
            collection: true,
            elementName: 'Choice',
            typeInfo: '.AlternateContent.Choice'
          }, {
            name: 'fallback',
            elementName: 'Fallback',
            typeInfo: '.AlternateContent.Fallback'
          }, {
            name: 'ignorable',
            attributeName: 'Ignorable',
            type: 'attribute'
          }, {
            name: 'mustUnderstand',
            attributeName: 'MustUnderstand',
            type: 'attribute'
          }, {
            name: 'processContent',
            attributeName: 'ProcessContent',
            type: 'attribute'
          }]
      }, {
        localName: 'AlternateContent.Choice',
        typeName: null,
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            mixed: false,
            type: 'anyElement'
          }, {
            name: 'requires',
            required: true,
            attributeName: {
              localPart: 'Requires'
            },
            type: 'attribute'
          }, {
            name: 'ignorable',
            attributeName: 'Ignorable',
            type: 'attribute'
          }, {
            name: 'mustUnderstand',
            attributeName: 'MustUnderstand',
            type: 'attribute'
          }, {
            name: 'processContent',
            attributeName: 'ProcessContent',
            type: 'attribute'
          }]
      }, {
        localName: 'AlternateContent.Fallback',
        typeName: null,
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            mixed: false,
            type: 'anyElement'
          }, {
            name: 'ignorable',
            attributeName: 'Ignorable',
            type: 'attribute'
          }, {
            name: 'mustUnderstand',
            attributeName: 'MustUnderstand',
            type: 'attribute'
          }, {
            name: 'processContent',
            attributeName: 'ProcessContent',
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.AlternateContent',
        elementName: 'AlternateContent'
      }]
  };
  return {
    org_docx4j_mce: org_docx4j_mce
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_mce_Module_Factory);
}
else {
  var org_docx4j_mce_Module = org_docx4j_mce_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_mce = org_docx4j_mce_Module.org_docx4j_mce;
  }
  else {
    var org_docx4j_mce = org_docx4j_mce_Module.org_docx4j_mce;
  }
}