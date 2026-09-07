var org_docx4j_vml_wordprocessingDrawing_Module_Factory = function () {
  var org_docx4j_vml_wordprocessingDrawing = {
    name: 'org_docx4j_vml_wordprocessingDrawing',
    defaultElementNamespaceURI: 'urn:schemas-microsoft-com:office:word',
    typeInfos: [{
        localName: 'CTAnchorLock',
        typeName: 'CT_AnchorLock'
      }, {
        localName: 'CTBorder',
        typeName: 'CT_Border',
        propertyInfos: [{
            name: 'shadow',
            typeInfo: '.STBorderShadow',
            attributeName: {
              localPart: 'shadow'
            },
            type: 'attribute'
          }, {
            name: 'type',
            typeInfo: '.STBorderType',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: 'width',
            typeInfo: 'PositiveInteger',
            attributeName: {
              localPart: 'width'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTWrap',
        typeName: 'CT_Wrap',
        propertyInfos: [{
            name: 'anchorx',
            typeInfo: '.STHorizontalAnchor',
            attributeName: {
              localPart: 'anchorx'
            },
            type: 'attribute'
          }, {
            name: 'anchory',
            typeInfo: '.STVerticalAnchor',
            attributeName: {
              localPart: 'anchory'
            },
            type: 'attribute'
          }, {
            name: 'side',
            typeInfo: '.STWrapSide',
            attributeName: {
              localPart: 'side'
            },
            type: 'attribute'
          }, {
            name: 'type',
            typeInfo: '.STWrapType',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STBorderShadow',
        values: ['t', 'true', 'f', 'false']
      }, {
        type: 'enumInfo',
        localName: 'STBorderType',
        values: ['none', 'single', 'thick', 'double', 'hairline', 'dot', 'dash', 'dotDash', 'dashDotDot', 'triple', 'thinThickSmall', 'thickThinSmall', 'thickBetweenThinSmall', 'thinThick', 'thickThin', 'thickBetweenThin', 'thinThickLarge', 'thickThinLarge', 'thickBetweenThinLarge', 'wave', 'doubleWave', 'dashedSmall', 'dashDotStroked', 'threeDEmboss', 'threeDEngrave', 'HTMLOutset', 'HTMLInset']
      }, {
        type: 'enumInfo',
        localName: 'STHorizontalAnchor',
        values: ['margin', 'page', 'text', 'char']
      }, {
        type: 'enumInfo',
        localName: 'STVerticalAnchor',
        values: ['margin', 'page', 'text', 'line']
      }, {
        type: 'enumInfo',
        localName: 'STWrapSide',
        values: ['both', 'left', 'right', 'largest']
      }, {
        type: 'enumInfo',
        localName: 'STWrapType',
        values: ['topAndBottom', 'square', 'none', 'tight', 'through']
      }],
    elementInfos: [{
        typeInfo: '.CTAnchorLock',
        elementName: 'anchorlock'
      }, {
        typeInfo: '.CTBorder',
        elementName: 'borderbottom'
      }, {
        typeInfo: '.CTBorder',
        elementName: 'borderleft'
      }, {
        typeInfo: '.CTBorder',
        elementName: 'borderright'
      }, {
        typeInfo: '.CTBorder',
        elementName: 'bordertop'
      }, {
        typeInfo: '.CTWrap',
        elementName: 'wrap'
      }]
  };
  return {
    org_docx4j_vml_wordprocessingDrawing: org_docx4j_vml_wordprocessingDrawing
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_vml_wordprocessingDrawing_Module_Factory);
}
else {
  var org_docx4j_vml_wordprocessingDrawing_Module = org_docx4j_vml_wordprocessingDrawing_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_vml_wordprocessingDrawing = org_docx4j_vml_wordprocessingDrawing_Module.org_docx4j_vml_wordprocessingDrawing;
  }
  else {
    var org_docx4j_vml_wordprocessingDrawing = org_docx4j_vml_wordprocessingDrawing_Module.org_docx4j_vml_wordprocessingDrawing;
  }
}