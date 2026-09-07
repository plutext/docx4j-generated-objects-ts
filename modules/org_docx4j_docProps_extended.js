var org_docx4j_docProps_extended_Module_Factory = function () {
  var org_docx4j_docProps_extended = {
    name: 'org_docx4j_docProps_extended',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/extended-properties',
    dependencies: ['org_docx4j_docProps_variantTypes'],
    typeInfos: [{
        localName: 'Properties',
        typeName: null,
        propertyInfos: [{
            name: 'template',
            elementName: 'Template'
          }, {
            name: 'manager',
            elementName: 'Manager'
          }, {
            name: 'company',
            elementName: 'Company'
          }, {
            name: 'pages',
            elementName: 'Pages',
            typeInfo: 'Int'
          }, {
            name: 'words',
            elementName: 'Words',
            typeInfo: 'Int'
          }, {
            name: 'characters',
            elementName: 'Characters',
            typeInfo: 'Int'
          }, {
            name: 'presentationFormat',
            elementName: 'PresentationFormat'
          }, {
            name: 'lines',
            elementName: 'Lines',
            typeInfo: 'Int'
          }, {
            name: 'paragraphs',
            elementName: 'Paragraphs',
            typeInfo: 'Int'
          }, {
            name: 'slides',
            elementName: 'Slides',
            typeInfo: 'Int'
          }, {
            name: 'notes',
            elementName: 'Notes',
            typeInfo: 'Int'
          }, {
            name: 'totalTime',
            elementName: 'TotalTime',
            typeInfo: 'Int'
          }, {
            name: 'hiddenSlides',
            elementName: 'HiddenSlides',
            typeInfo: 'Int'
          }, {
            name: 'mmClips',
            elementName: 'MMClips',
            typeInfo: 'Int'
          }, {
            name: 'scaleCrop',
            elementName: 'ScaleCrop',
            typeInfo: 'Boolean'
          }, {
            name: 'headingPairs',
            elementName: 'HeadingPairs',
            typeInfo: '.Properties.HeadingPairs'
          }, {
            name: 'titlesOfParts',
            elementName: 'TitlesOfParts',
            typeInfo: '.Properties.TitlesOfParts'
          }, {
            name: 'linksUpToDate',
            elementName: 'LinksUpToDate',
            typeInfo: 'Boolean'
          }, {
            name: 'charactersWithSpaces',
            elementName: 'CharactersWithSpaces',
            typeInfo: 'Int'
          }, {
            name: 'sharedDoc',
            elementName: 'SharedDoc',
            typeInfo: 'Boolean'
          }, {
            name: 'hyperlinkBase',
            elementName: 'HyperlinkBase'
          }, {
            name: 'hLinks',
            elementName: 'HLinks',
            typeInfo: '.Properties.HLinks'
          }, {
            name: 'hyperlinksChanged',
            elementName: 'HyperlinksChanged',
            typeInfo: 'Boolean'
          }, {
            name: 'digSig',
            elementName: 'DigSig',
            typeInfo: '.Properties.DigSig'
          }, {
            name: 'application',
            elementName: 'Application'
          }, {
            name: 'appVersion',
            elementName: 'AppVersion'
          }, {
            name: 'docSecurity',
            elementName: 'DocSecurity',
            typeInfo: 'Int'
          }]
      }, {
        localName: 'Properties.DigSig',
        typeName: null,
        propertyInfos: [{
            name: 'blob',
            required: true,
            elementName: {
              localPart: 'blob',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/docPropsVTypes'
            },
            typeInfo: 'Base64Binary'
          }]
      }, {
        localName: 'Properties.HLinks',
        typeName: null,
        propertyInfos: [{
            name: 'vector',
            required: true,
            elementName: {
              localPart: 'vector',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/docPropsVTypes'
            },
            typeInfo: 'org_docx4j_docProps_variantTypes.Vector'
          }]
      }, {
        localName: 'Properties.HeadingPairs',
        typeName: null,
        propertyInfos: [{
            name: 'vector',
            required: true,
            elementName: {
              localPart: 'vector',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/docPropsVTypes'
            },
            typeInfo: 'org_docx4j_docProps_variantTypes.Vector'
          }]
      }, {
        localName: 'Properties.TitlesOfParts',
        typeName: null,
        propertyInfos: [{
            name: 'vector',
            required: true,
            elementName: {
              localPart: 'vector',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/docPropsVTypes'
            },
            typeInfo: 'org_docx4j_docProps_variantTypes.Vector'
          }]
      }],
    elementInfos: [{
        typeInfo: '.Properties',
        elementName: 'Properties'
      }]
  };
  return {
    org_docx4j_docProps_extended: org_docx4j_docProps_extended
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_docProps_extended_Module_Factory);
}
else {
  var org_docx4j_docProps_extended_Module = org_docx4j_docProps_extended_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_docProps_extended = org_docx4j_docProps_extended_Module.org_docx4j_docProps_extended;
  }
  else {
    var org_docx4j_docProps_extended = org_docx4j_docProps_extended_Module.org_docx4j_docProps_extended;
  }
}