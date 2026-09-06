var org_docx4j_docProps_coverPageProps_Module_Factory = function () {
  var org_docx4j_docProps_coverPageProps = {
    name: 'org_docx4j_docProps_coverPageProps',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/2006\/coverPageProps',
    typeInfos: [{
        localName: 'CoverPageProperties',
        typeName: null,
        propertyInfos: [{
            name: 'publishDate',
            required: true,
            elementName: 'PublishDate'
          }, {
            name: '_abstract',
            required: true,
            elementName: 'Abstract'
          }, {
            name: 'companyAddress',
            required: true,
            elementName: 'CompanyAddress'
          }, {
            name: 'companyPhone',
            required: true,
            elementName: 'CompanyPhone'
          }, {
            name: 'companyFax',
            required: true,
            elementName: 'CompanyFax'
          }, {
            name: 'companyEmail',
            required: true,
            elementName: 'CompanyEmail'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CoverPageProperties',
        elementName: 'CoverPageProperties'
      }]
  };
  return {
    org_docx4j_docProps_coverPageProps: org_docx4j_docProps_coverPageProps
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_docProps_coverPageProps_Module_Factory);
}
else {
  var org_docx4j_docProps_coverPageProps_Module = org_docx4j_docProps_coverPageProps_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_docProps_coverPageProps = org_docx4j_docProps_coverPageProps_Module.org_docx4j_docProps_coverPageProps;
  }
  else {
    var org_docx4j_docProps_coverPageProps = org_docx4j_docProps_coverPageProps_Module.org_docx4j_docProps_coverPageProps;
  }
}