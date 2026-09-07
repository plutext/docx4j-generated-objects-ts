var org_docx4j_docProps_custom_Module_Factory = function () {
  var org_docx4j_docProps_custom = {
    name: 'org_docx4j_docProps_custom',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/docPropsVTypes',
    dependencies: ['org_docx4j_docProps_variantTypes'],
    typeInfos: [{
        localName: 'Properties',
        typeName: null,
        propertyInfos: [{
            name: 'property',
            minOccurs: 0,
            collection: true,
            elementName: {
              localPart: 'property',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/custom-properties'
            },
            typeInfo: '.Properties.Property'
          }]
      }, {
        localName: 'Properties.Property',
        typeName: null,
        propertyInfos: [{
            name: 'vector',
            typeInfo: 'org_docx4j_docProps_variantTypes.Vector'
          }, {
            name: 'array',
            typeInfo: 'org_docx4j_docProps_variantTypes.Array'
          }, {
            name: 'blob',
            typeInfo: 'Base64Binary'
          }, {
            name: 'oblob',
            typeInfo: 'Base64Binary'
          }, {
            name: 'empty',
            typeInfo: 'org_docx4j_docProps_variantTypes.Empty'
          }, {
            name: '_null',
            elementName: 'null',
            typeInfo: 'org_docx4j_docProps_variantTypes.Null'
          }, {
            name: 'i1',
            typeInfo: 'Byte'
          }, {
            name: 'i2',
            typeInfo: 'Short'
          }, {
            name: 'i4',
            typeInfo: 'Int'
          }, {
            name: 'i8',
            typeInfo: 'Long'
          }, {
            name: '_int',
            elementName: 'int',
            typeInfo: 'Int'
          }, {
            name: 'ui1',
            typeInfo: 'UnsignedByte'
          }, {
            name: 'ui2',
            typeInfo: 'UnsignedShort'
          }, {
            name: 'ui4',
            typeInfo: 'UnsignedInt'
          }, {
            name: 'ui8',
            typeInfo: 'UnsignedLong'
          }, {
            name: 'uint',
            typeInfo: 'UnsignedInt'
          }, {
            name: 'r4',
            typeInfo: 'Float'
          }, {
            name: 'r8',
            typeInfo: 'Double'
          }, {
            name: 'decimal',
            typeInfo: 'Decimal'
          }, {
            name: 'lpstr'
          }, {
            name: 'lpwstr'
          }, {
            name: 'bstr'
          }, {
            name: 'date',
            typeInfo: 'DateTime'
          }, {
            name: 'filetime',
            typeInfo: 'DateTime'
          }, {
            name: 'bool',
            typeInfo: 'Boolean'
          }, {
            name: 'cy'
          }, {
            name: 'error'
          }, {
            name: 'stream',
            typeInfo: 'Base64Binary'
          }, {
            name: 'ostream',
            typeInfo: 'Base64Binary'
          }, {
            name: 'storage',
            typeInfo: 'Base64Binary'
          }, {
            name: 'ostorage',
            typeInfo: 'Base64Binary'
          }, {
            name: 'vstream',
            typeInfo: 'org_docx4j_docProps_variantTypes.Vstream'
          }, {
            name: 'clsid'
          }, {
            name: 'cf',
            typeInfo: 'org_docx4j_docProps_variantTypes.Cf'
          }, {
            name: 'fmtid',
            required: true,
            attributeName: {
              localPart: 'fmtid'
            },
            type: 'attribute'
          }, {
            name: 'linkTarget',
            attributeName: {
              localPart: 'linkTarget'
            },
            type: 'attribute'
          }, {
            name: 'name',
            attributeName: {
              localPart: 'name'
            },
            type: 'attribute'
          }, {
            name: 'pid',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'pid'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.Properties',
        elementName: {
          localPart: 'Properties',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/custom-properties'
        }
      }]
  };
  return {
    org_docx4j_docProps_custom: org_docx4j_docProps_custom
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_docProps_custom_Module_Factory);
}
else {
  var org_docx4j_docProps_custom_Module = org_docx4j_docProps_custom_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_docProps_custom = org_docx4j_docProps_custom_Module.org_docx4j_docProps_custom;
  }
  else {
    var org_docx4j_docProps_custom = org_docx4j_docProps_custom_Module.org_docx4j_docProps_custom;
  }
}