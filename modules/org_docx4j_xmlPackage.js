var org_docx4j_xmlPackage_Module_Factory = function () {
  var org_docx4j_xmlPackage = {
    name: 'org_docx4j_xmlPackage',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/2006\/xmlPackage',
    defaultAttributeNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/2006\/xmlPackage',
    typeInfos: [{
        localName: 'Package',
        typeName: 'CT_Package',
        propertyInfos: [{
            name: 'part',
            minOccurs: 0,
            collection: true,
            typeInfo: '.Part'
          }]
      }, {
        localName: 'Part',
        typeName: 'CT_Part',
        propertyInfos: [{
            name: 'xmlData',
            typeInfo: '.XmlData'
          }, {
            name: 'binaryData',
            typeInfo: 'Base64Binary'
          }, {
            name: 'compression',
            values: ['store', 'deflateSuperFast', 'deflateFast', 'deflateNormal', 'deflateMaximum'],
            defaultValue: 'deflateSuperFast',
            type: 'attribute'
          }, {
            name: 'contentType',
            type: 'attribute'
          }, {
            name: 'name',
            required: true,
            type: 'attribute'
          }, {
            name: 'padding',
            typeInfo: 'UnsignedInt',
            defaultValue: 0,
            type: 'attribute'
          }]
      }, {
        localName: 'XmlData',
        typeName: 'CT_XmlData',
        propertyInfos: [{
            name: 'any',
            required: true,
            allowTypedObject: false,
            mixed: false,
            type: 'anyElement'
          }, {
            name: 'originalXmlEncoding',
            defaultValue: 'UTF-8',
            type: 'attribute'
          }, {
            name: 'originalXmlStandalone',
            defaultValue: 'yes',
            type: 'attribute'
          }, {
            name: 'originalXmlVersion',
            defaultValue: '1.0',
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.Package',
        elementName: 'package'
      }]
  };
  return {
    org_docx4j_xmlPackage: org_docx4j_xmlPackage
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_xmlPackage_Module_Factory);
}
else {
  var org_docx4j_xmlPackage_Module = org_docx4j_xmlPackage_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_xmlPackage = org_docx4j_xmlPackage_Module.org_docx4j_xmlPackage;
  }
  else {
    var org_docx4j_xmlPackage = org_docx4j_xmlPackage_Module.org_docx4j_xmlPackage;
  }
}