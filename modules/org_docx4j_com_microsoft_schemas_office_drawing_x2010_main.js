var org_docx4j_com_microsoft_schemas_office_drawing_x2010_main_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2010_main = {
    name: 'org_docx4j_com_microsoft_schemas_office_drawing_x2010_main',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/drawing\/2010\/main',
    dependencies: ['org_docx4j_dml'],
    typeInfos: [{
        localName: 'CTCameraTool',
        typeName: 'CT_CameraTool',
        propertyInfos: [{
            name: 'cellRange',
            attributeName: {
              localPart: 'cellRange'
            },
            type: 'attribute'
          }, {
            name: 'spid',
            defaultValue: '0',
            attributeName: {
              localPart: 'spid'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCompatExt',
        typeName: 'CT_CompatExt',
        propertyInfos: [{
            name: 'spid',
            attributeName: {
              localPart: 'spid'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTContentPartLocking',
        typeName: 'CT_ContentPartLocking',
        propertyInfos: [{
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'noAdjustHandles',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'noAdjustHandles'
            },
            type: 'attribute'
          }, {
            name: 'noChangeArrowheads',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'noChangeArrowheads'
            },
            type: 'attribute'
          }, {
            name: 'noChangeAspect',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'noChangeAspect'
            },
            type: 'attribute'
          }, {
            name: 'noChangeShapeType',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'noChangeShapeType'
            },
            type: 'attribute'
          }, {
            name: 'noEditPoints',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'noEditPoints'
            },
            type: 'attribute'
          }, {
            name: 'noGrp',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'noGrp'
            },
            type: 'attribute'
          }, {
            name: 'noMove',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'noMove'
            },
            type: 'attribute'
          }, {
            name: 'noResize',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'noResize'
            },
            type: 'attribute'
          }, {
            name: 'noRot',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'noRot'
            },
            type: 'attribute'
          }, {
            name: 'noSelect',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'noSelect'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGvmlContentPart',
        typeName: 'CT_GvmlContentPart',
        propertyInfos: [{
            name: 'nvContentPartPr',
            typeInfo: '.CTGvmlContentPartNonVisual'
          }, {
            name: 'xfrm',
            typeInfo: 'org_docx4j_dml.CTTransform2D'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'bwMode',
            typeInfo: 'org_docx4j_dml.STBlackWhiteMode',
            attributeName: {
              localPart: 'bwMode'
            },
            type: 'attribute'
          }, {
            name: 'id',
            required: true,
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTGvmlContentPartNonVisual',
        typeName: 'CT_GvmlContentPartNonVisual',
        propertyInfos: [{
            name: 'cNvPr',
            required: true,
            typeInfo: 'org_docx4j_dml.CTNonVisualDrawingProps'
          }, {
            name: 'cNvContentPartPr',
            typeInfo: '.CTNonVisualInkContentPartProperties'
          }]
      }, {
        localName: 'CTIsGvmlCanvas',
        typeName: 'CT_IsGvmlCanvas',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'Boolean',
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTNonVisualInkContentPartProperties',
        typeName: 'CT_NonVisualInkContentPartProperties',
        propertyInfos: [{
            name: 'cpLocks',
            typeInfo: '.CTContentPartLocking'
          }, {
            name: 'extLst',
            typeInfo: 'org_docx4j_dml.CTOfficeArtExtensionList'
          }, {
            name: 'isComment',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'isComment'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPhoto',
        typeName: 'CT_Photo',
        propertyInfos: [{
            name: 'imgLayer',
            required: true,
            typeInfo: '.CTPictureLayer'
          }]
      }, {
        localName: 'CTPictureEffect',
        typeName: 'CT_PictureEffect',
        propertyInfos: [{
            name: 'artisticBlur',
            typeInfo: '.CTPictureEffectBlur'
          }, {
            name: 'artisticCement',
            typeInfo: '.CTPictureEffectCement'
          }, {
            name: 'artisticChalkSketch',
            typeInfo: '.CTPictureEffectChalkSketch'
          }, {
            name: 'artisticCrisscrossEtching',
            typeInfo: '.CTPictureEffectCrisscrossEtching'
          }, {
            name: 'artisticCutout',
            typeInfo: '.CTPictureEffectCutout'
          }, {
            name: 'artisticFilmGrain',
            typeInfo: '.CTPictureEffectFilmGrain'
          }, {
            name: 'artisticGlass',
            typeInfo: '.CTPictureEffectGlass'
          }, {
            name: 'artisticGlowDiffused',
            typeInfo: '.CTPictureEffectGlowDiffused'
          }, {
            name: 'artisticGlowEdges',
            typeInfo: '.CTPictureEffectGlowEdges'
          }, {
            name: 'artisticLightScreen',
            typeInfo: '.CTPictureEffectLightScreen'
          }, {
            name: 'artisticLineDrawing',
            typeInfo: '.CTPictureEffectLineDrawing'
          }, {
            name: 'artisticMarker',
            typeInfo: '.CTPictureEffectMarker'
          }, {
            name: 'artisticMosiaicBubbles',
            typeInfo: '.CTPictureEffectMosiaicBubbles'
          }, {
            name: 'artisticPaintStrokes',
            typeInfo: '.CTPictureEffectPaintStrokes'
          }, {
            name: 'artisticPaintBrush',
            typeInfo: '.CTPictureEffectPaintBrush'
          }, {
            name: 'artisticPastelsSmooth',
            typeInfo: '.CTPictureEffectPastelsSmooth'
          }, {
            name: 'artisticPencilGrayscale',
            typeInfo: '.CTPictureEffectPencilGrayscale'
          }, {
            name: 'artisticPencilSketch',
            typeInfo: '.CTPictureEffectPencilSketch'
          }, {
            name: 'artisticPhotocopy',
            typeInfo: '.CTPictureEffectPhotocopy'
          }, {
            name: 'artisticPlasticWrap',
            typeInfo: '.CTPictureEffectPlasticWrap'
          }, {
            name: 'artisticTexturizer',
            typeInfo: '.CTPictureEffectTexturizer'
          }, {
            name: 'artisticWatercolorSponge',
            typeInfo: '.CTPictureEffectWatercolorSponge'
          }, {
            name: 'backgroundRemoval',
            typeInfo: '.CTPictureEffectBackgroundRemoval'
          }, {
            name: 'brightnessContrast',
            typeInfo: '.CTPictureEffectBrightnessContrast'
          }, {
            name: 'colorTemperature',
            typeInfo: '.CTPictureEffectColorTemperature'
          }, {
            name: 'saturation',
            typeInfo: '.CTPictureEffectSaturation'
          }, {
            name: 'sharpenSoften',
            typeInfo: '.CTPictureEffectSharpenSoften'
          }, {
            name: 'visible',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'visible'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectBackgroundRemoval',
        typeName: 'CT_PictureEffectBackgroundRemoval',
        propertyInfos: [{
            name: 'foregroundMark',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTPictureEffectBackgroundRemovalForegroundMark'
          }, {
            name: 'backgroundMark',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTPictureEffectBackgroundRemovalBackgroundMark'
          }, {
            name: 'b',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'b'
            },
            type: 'attribute'
          }, {
            name: 'l',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'l'
            },
            type: 'attribute'
          }, {
            name: 'r',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'r'
            },
            type: 'attribute'
          }, {
            name: 't',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 't'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectBackgroundRemovalBackgroundMark',
        typeName: 'CT_PictureEffectBackgroundRemovalBackgroundMark',
        propertyInfos: [{
            name: 'x1',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'x1'
            },
            type: 'attribute'
          }, {
            name: 'x2',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'x2'
            },
            type: 'attribute'
          }, {
            name: 'y1',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'y1'
            },
            type: 'attribute'
          }, {
            name: 'y2',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'y2'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectBackgroundRemovalForegroundMark',
        typeName: 'CT_PictureEffectBackgroundRemovalForegroundMark',
        propertyInfos: [{
            name: 'x1',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'x1'
            },
            type: 'attribute'
          }, {
            name: 'x2',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'x2'
            },
            type: 'attribute'
          }, {
            name: 'y1',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'y1'
            },
            type: 'attribute'
          }, {
            name: 'y2',
            required: true,
            typeInfo: 'Int',
            attributeName: {
              localPart: 'y2'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectBlur',
        typeName: 'CT_PictureEffectBlur',
        propertyInfos: [{
            name: 'radius',
            typeInfo: 'Int',
            defaultValue: 10,
            attributeName: {
              localPart: 'radius'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectBrightnessContrast',
        typeName: 'CT_PictureEffectBrightnessContrast',
        propertyInfos: [{
            name: 'bright',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'bright'
            },
            type: 'attribute'
          }, {
            name: 'contrast',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'contrast'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectCement',
        typeName: 'CT_PictureEffectCement',
        propertyInfos: [{
            name: 'crackSpacing',
            typeInfo: 'Int',
            defaultValue: 24,
            attributeName: {
              localPart: 'crackSpacing'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectChalkSketch',
        typeName: 'CT_PictureEffectChalkSketch',
        propertyInfos: [{
            name: 'pressure',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'pressure'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectColorTemperature',
        typeName: 'CT_PictureEffectColorTemperature',
        propertyInfos: [{
            name: 'colorTemp',
            typeInfo: 'Int',
            defaultValue: 6500,
            attributeName: {
              localPart: 'colorTemp'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectCrisscrossEtching',
        typeName: 'CT_PictureEffectCrisscrossEtching',
        propertyInfos: [{
            name: 'pressure',
            typeInfo: 'Int',
            defaultValue: 30,
            attributeName: {
              localPart: 'pressure'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 75,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectCutout',
        typeName: 'CT_PictureEffectCutout',
        propertyInfos: [{
            name: 'numberOfShades',
            typeInfo: 'Int',
            defaultValue: 2,
            attributeName: {
              localPart: 'numberOfShades'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectFilmGrain',
        typeName: 'CT_PictureEffectFilmGrain',
        propertyInfos: [{
            name: 'grainSize',
            typeInfo: 'Int',
            defaultValue: 40,
            attributeName: {
              localPart: 'grainSize'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectGlass',
        typeName: 'CT_PictureEffectGlass',
        propertyInfos: [{
            name: 'scaling',
            typeInfo: 'Int',
            defaultValue: 34,
            attributeName: {
              localPart: 'scaling'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectGlowDiffused',
        typeName: 'CT_PictureEffectGlowDiffused',
        propertyInfos: [{
            name: 'intensity',
            typeInfo: 'Int',
            defaultValue: 5,
            attributeName: {
              localPart: 'intensity'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectGlowEdges',
        typeName: 'CT_PictureEffectGlowEdges',
        propertyInfos: [{
            name: 'smoothness',
            typeInfo: 'Int',
            defaultValue: 3,
            attributeName: {
              localPart: 'smoothness'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 15,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectLightScreen',
        typeName: 'CT_PictureEffectLightScreen',
        propertyInfos: [{
            name: 'gridSize',
            typeInfo: 'Int',
            defaultValue: 4,
            attributeName: {
              localPart: 'gridSize'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectLineDrawing',
        typeName: 'CT_PictureEffectLineDrawing',
        propertyInfos: [{
            name: 'pencilSize',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'pencilSize'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 25,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectMarker',
        typeName: 'CT_PictureEffectMarker',
        propertyInfos: [{
            name: 'size',
            typeInfo: 'Int',
            defaultValue: 97,
            attributeName: {
              localPart: 'size'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectMosiaicBubbles',
        typeName: 'CT_PictureEffectMosiaicBubbles',
        propertyInfos: [{
            name: 'pressure',
            typeInfo: 'Int',
            defaultValue: 14,
            attributeName: {
              localPart: 'pressure'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectPaintBrush',
        typeName: 'CT_PictureEffectPaintBrush',
        propertyInfos: [{
            name: 'brushSize',
            typeInfo: 'Int',
            defaultValue: 2,
            attributeName: {
              localPart: 'brushSize'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectPaintStrokes',
        typeName: 'CT_PictureEffectPaintStrokes',
        propertyInfos: [{
            name: 'intensity',
            typeInfo: 'Int',
            defaultValue: 5,
            attributeName: {
              localPart: 'intensity'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectPastelsSmooth',
        typeName: 'CT_PictureEffectPastelsSmooth',
        propertyInfos: [{
            name: 'scaling',
            typeInfo: 'Int',
            defaultValue: 34,
            attributeName: {
              localPart: 'scaling'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectPencilGrayscale',
        typeName: 'CT_PictureEffectPencilGrayscale',
        propertyInfos: [{
            name: 'pencilSize',
            typeInfo: 'Int',
            defaultValue: 27,
            attributeName: {
              localPart: 'pencilSize'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectPencilSketch',
        typeName: 'CT_PictureEffectPencilSketch',
        propertyInfos: [{
            name: 'pressure',
            typeInfo: 'Int',
            defaultValue: 22,
            attributeName: {
              localPart: 'pressure'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectPhotocopy',
        typeName: 'CT_PictureEffectPhotocopy',
        propertyInfos: [{
            name: 'detail',
            typeInfo: 'Int',
            defaultValue: 3,
            attributeName: {
              localPart: 'detail'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 30,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectPlasticWrap',
        typeName: 'CT_PictureEffectPlasticWrap',
        propertyInfos: [{
            name: 'smoothness',
            typeInfo: 'Int',
            defaultValue: 5,
            attributeName: {
              localPart: 'smoothness'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectSaturation',
        typeName: 'CT_PictureEffectSaturation',
        propertyInfos: [{
            name: 'sat',
            typeInfo: 'Int',
            defaultValue: 100,
            attributeName: {
              localPart: 'sat'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectSharpenSoften',
        typeName: 'CT_PictureEffectSharpenSoften',
        propertyInfos: [{
            name: 'amount',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'amount'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectTexturizer',
        typeName: 'CT_PictureEffectTexturizer',
        propertyInfos: [{
            name: 'scaling',
            typeInfo: 'Int',
            defaultValue: 34,
            attributeName: {
              localPart: 'scaling'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureEffectWatercolorSponge',
        typeName: 'CT_PictureEffectWatercolorSponge',
        propertyInfos: [{
            name: 'brushSize',
            typeInfo: 'Int',
            defaultValue: 2,
            attributeName: {
              localPart: 'brushSize'
            },
            type: 'attribute'
          }, {
            name: 'trans',
            typeInfo: 'Int',
            defaultValue: 0,
            attributeName: {
              localPart: 'trans'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTPictureLayer',
        typeName: 'CT_PictureLayer',
        propertyInfos: [{
            name: 'imgEffect',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTPictureEffect'
          }, {
            name: 'embed',
            defaultValue: '',
            attributeName: {
              localPart: 'embed',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTShadowObscured',
        typeName: 'CT_ShadowObscured',
        propertyInfos: [{
            name: 'val',
            typeInfo: 'Boolean',
            defaultValue: false,
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTTextMath',
        typeName: 'CT_TextMath'
      }, {
        localName: 'CTUseLocalDpi',
        typeName: 'CT_UseLocalDpi',
        propertyInfos: [{
            name: 'val',
            typeInfo: 'Boolean',
            defaultValue: true,
            attributeName: {
              localPart: 'val'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTCameraTool',
        elementName: 'cameraTool'
      }, {
        typeInfo: '.CTCompatExt',
        elementName: 'compatExt'
      }, {
        typeInfo: '.CTGvmlContentPart',
        elementName: 'contentPart'
      }, {
        typeInfo: 'org_docx4j_dml.CTEffectProperties',
        elementName: 'hiddenEffects'
      }, {
        typeInfo: 'org_docx4j_dml.CTFillProperties',
        elementName: 'hiddenFill'
      }, {
        typeInfo: 'org_docx4j_dml.CTLineProperties',
        elementName: 'hiddenLine'
      }, {
        typeInfo: 'org_docx4j_dml.CTScene3D',
        elementName: 'hiddenScene3d'
      }, {
        typeInfo: 'org_docx4j_dml.CTShape3D',
        elementName: 'hiddenSp3d'
      }, {
        typeInfo: '.CTPhoto',
        elementName: 'imgProps'
      }, {
        typeInfo: '.CTIsGvmlCanvas',
        elementName: 'isCanvas'
      }, {
        typeInfo: '.CTTextMath',
        elementName: 'm'
      }, {
        typeInfo: '.CTShadowObscured',
        elementName: 'shadowObscured'
      }, {
        typeInfo: '.CTUseLocalDpi',
        elementName: 'useLocalDpi'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_drawing_x2010_main: org_docx4j_com_microsoft_schemas_office_drawing_x2010_main
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_drawing_x2010_main_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_drawing_x2010_main_Module = org_docx4j_com_microsoft_schemas_office_drawing_x2010_main_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_drawing_x2010_main = org_docx4j_com_microsoft_schemas_office_drawing_x2010_main_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2010_main;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_drawing_x2010_main = org_docx4j_com_microsoft_schemas_office_drawing_x2010_main_Module.org_docx4j_com_microsoft_schemas_office_drawing_x2010_main;
  }
}