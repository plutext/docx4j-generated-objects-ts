var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main_Module_Factory = function () {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main = {
    name: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/powerpoint\/2017\/3\/main',
    typeInfos: [{
        localName: 'CTTrack',
        typeName: 'CT_Track',
        propertyInfos: [{
            name: 'id',
            required: true,
            typeInfo: 'Token',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'label',
            required: true,
            attributeName: {
              localPart: 'label'
            },
            type: 'attribute'
          }, {
            name: 'lang',
            attributeName: {
              localPart: 'lang'
            },
            type: 'attribute'
          }, {
            name: 'embed',
            defaultValue: '',
            attributeName: {
              localPart: 'embed',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }, {
            name: 'link',
            defaultValue: '',
            attributeName: {
              localPart: 'link',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTTrackList',
        typeName: 'CT_TrackList',
        propertyInfos: [{
            name: 'track',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTTrack'
          }]
      }, {
        localName: 'CTTracksInfo',
        typeName: 'CT_TracksInfo',
        propertyInfos: [{
            name: 'trackLst',
            typeInfo: '.CTTrackList'
          }, {
            name: 'displayLoc',
            required: true,
            typeInfo: '.STDisplayLocation',
            attributeName: {
              localPart: 'displayLoc'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STDisplayLocation',
        values: ['media', 'slide']
      }],
    elementInfos: [{
        typeInfo: '.CTTracksInfo',
        elementName: 'tracksInfo'
      }]
  };
  return {
    org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main: org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main_Module_Factory);
}
else {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main_Module = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main;
  }
  else {
    var org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main = org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x201703_main;
  }
}