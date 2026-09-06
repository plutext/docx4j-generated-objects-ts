var org_docx4j_docProps_core_Module_Factory = function () {
  var org_docx4j_docProps_core = {
    name: 'org_docx4j_docProps_core',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/package\/2006\/metadata\/core-properties',
    dependencies: ['org_docx4j_docProps_core_dc_elements'],
    typeInfos: [{
        localName: 'CoreProperties',
        typeName: null,
        propertyInfos: [{
            name: 'category'
          }, {
            name: 'contentStatus'
          }, {
            name: 'contentType'
          }, {
            name: 'created',
            elementName: {
              localPart: 'created',
              namespaceURI: 'http:\/\/purl.org\/dc\/terms\/'
            },
            typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
          }, {
            name: 'creator',
            elementName: {
              localPart: 'creator',
              namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
            },
            typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
          }, {
            name: 'description',
            mixed: false,
            allowDom: false,
            elementName: {
              localPart: 'description',
              namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
            },
            typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
            type: 'elementRef'
          }, {
            name: 'identifier',
            mixed: false,
            allowDom: false,
            elementName: {
              localPart: 'identifier',
              namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
            },
            typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
            type: 'elementRef'
          }, {
            name: 'keywords'
          }, {
            name: 'language',
            elementName: {
              localPart: 'language',
              namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
            },
            typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
          }, {
            name: 'lastModifiedBy'
          }, {
            name: 'lastPrinted',
            typeInfo: 'DateTime'
          }, {
            name: 'modified',
            elementName: {
              localPart: 'modified',
              namespaceURI: 'http:\/\/purl.org\/dc\/terms\/'
            },
            typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
          }, {
            name: 'revision'
          }, {
            name: 'subject',
            elementName: {
              localPart: 'subject',
              namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
            },
            typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral'
          }, {
            name: 'title',
            mixed: false,
            allowDom: false,
            elementName: {
              localPart: 'title',
              namespaceURI: 'http:\/\/purl.org\/dc\/elements\/1.1\/'
            },
            typeInfo: 'org_docx4j_docProps_core_dc_elements.SimpleLiteral',
            type: 'elementRef'
          }, {
            name: 'version'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CoreProperties',
        elementName: 'coreProperties'
      }]
  };
  return {
    org_docx4j_docProps_core: org_docx4j_docProps_core
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_docProps_core_Module_Factory);
}
else {
  var org_docx4j_docProps_core_Module = org_docx4j_docProps_core_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_docProps_core = org_docx4j_docProps_core_Module.org_docx4j_docProps_core;
  }
  else {
    var org_docx4j_docProps_core = org_docx4j_docProps_core_Module.org_docx4j_docProps_core;
  }
}