var org_docx4j_docProps_core_dc_terms_Module_Factory = function () {
  var org_docx4j_docProps_core_dc_terms = {
    name: 'org_docx4j_docProps_core_dc_terms',
    defaultElementNamespaceURI: 'http:\/\/purl.org\/dc\/terms\/',
    dependencies: ['org_docx4j_docProps_core_dc_elements'],
    typeInfos: [{
        localName: 'Box',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'DCMIType',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'DDC',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'ElementOrRefinementContainer',
        typeName: 'elementOrRefinementContainer',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementName: {
              localPart: 'any',
              namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
            },
            typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
            type: 'elementRef'
          }]
      }, {
        localName: 'IMT',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'ISO3166',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'ISO6392',
        typeName: 'ISO639-2',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'LCC',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'LCSH',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'MESH',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'Period',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'Point',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'RFC1766',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'RFC3066',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'TGN',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'UDC',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'URI',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }, {
        localName: 'W3CDTF',
        baseTypeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
      }],
    elementInfos: [{
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'abstract',
        substitutionHead: {
          localPart: 'description',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'accessRights',
        substitutionHead: {
          localPart: 'rights',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'alternative',
        substitutionHead: {
          localPart: 'title',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'audience',
        substitutionHead: {
          localPart: 'any',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'available',
        substitutionHead: {
          localPart: 'date',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'bibliographicCitation',
        substitutionHead: {
          localPart: 'identifier',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'conformsTo',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'created',
        substitutionHead: {
          localPart: 'date',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'dateAccepted',
        substitutionHead: {
          localPart: 'date',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'dateCopyrighted',
        substitutionHead: {
          localPart: 'date',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'dateSubmitted',
        substitutionHead: {
          localPart: 'date',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'educationLevel',
        substitutionHead: 'audience'
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'extent',
        substitutionHead: {
          localPart: 'format',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'hasFormat',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'hasPart',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'hasVersion',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'isFormatOf',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'isPartOf',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'isReferencedBy',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'isReplacedBy',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'isRequiredBy',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'isVersionOf',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'issued',
        substitutionHead: {
          localPart: 'date',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'mediator',
        substitutionHead: 'audience'
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'medium',
        substitutionHead: {
          localPart: 'format',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'modified',
        substitutionHead: {
          localPart: 'date',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'references',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'replaces',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'requires',
        substitutionHead: {
          localPart: 'relation',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'spatial',
        substitutionHead: {
          localPart: 'coverage',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'tableOfContents',
        substitutionHead: {
          localPart: 'description',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'temporal',
        substitutionHead: {
          localPart: 'coverage',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }, {
        typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
        elementName: 'valid',
        substitutionHead: {
          localPart: 'date',
          namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
        }
      }]
  };
  return {
    org_docx4j_docProps_core_dc_terms: org_docx4j_docProps_core_dc_terms
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_docProps_core_dc_terms_Module_Factory);
}
else {
  var org_docx4j_docProps_core_dc_terms_Module = org_docx4j_docProps_core_dc_terms_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_docProps_core_dc_terms = org_docx4j_docProps_core_dc_terms_Module.org_docx4j_docProps_core_dc_terms;
  }
  else {
    var org_docx4j_docProps_core_dc_terms = org_docx4j_docProps_core_dc_terms_Module.org_docx4j_docProps_core_dc_terms;
  }
}