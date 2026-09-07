var org_docx4j_vml_Module_Factory = function () {
  var org_docx4j_vml = {
    name: 'org_docx4j_vml',
    defaultElementNamespaceURI: 'urn:schemas-microsoft-com:vml',
    defaultAttributeNamespaceURI: 'urn:schemas-microsoft-com:office:office',
    dependencies: ['org_docx4j_vml_officedrawing', 'org_docx4j_wml'],
    typeInfos: [{
        localName: 'CTArc',
        typeName: 'CT_Arc',
        propertyInfos: [{
            name: 'egShapeElements',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'path',
                typeInfo: '.CTPath'
              }, {
                elementName: 'formulas',
                typeInfo: '.CTFormulas'
              }, {
                elementName: 'handles',
                typeInfo: '.CTHandles'
              }, {
                elementName: 'fill',
                typeInfo: '.CTFill'
              }, {
                elementName: 'stroke',
                typeInfo: '.CTStroke'
              }, {
                elementName: 'shadow',
                typeInfo: '.CTShadow'
              }, {
                elementName: 'textbox',
                typeInfo: '.CTTextbox'
              }, {
                elementName: 'textpath',
                typeInfo: '.CTTextPath'
              }, {
                elementName: 'imagedata',
                typeInfo: '.CTImageData'
              }, {
                elementName: {
                  localPart: 'skew',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSkew'
              }, {
                elementName: {
                  localPart: 'extrusion',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTExtrusion'
              }, {
                elementName: {
                  localPart: 'callout',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTCallout'
              }, {
                elementName: {
                  localPart: 'lock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTLock'
              }, {
                elementName: {
                  localPart: 'clippath',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTClipPath'
              }, {
                elementName: {
                  localPart: 'signatureline',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSignatureLine'
              }, {
                elementName: {
                  localPart: 'wrap',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTWrap'
              }, {
                elementName: {
                  localPart: 'anchorlock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTAnchorLock'
              }, {
                elementName: {
                  localPart: 'bordertop',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderbottom',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderleft',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderright',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'ClientData',
                  namespaceURI: 'urn:schemas-microsoft-com:office:excel'
                },
                typeInfo: 'org_docx4j_vml_spreadsheetDrawing.CTClientData'
              }, {
                elementName: {
                  localPart: 'textdata',
                  namespaceURI: 'urn:schemas-microsoft-com:office:powerpoint'
                },
                typeInfo: 'org_docx4j_vml_presentationDrawing.CTRel'
              }],
            type: 'elementRefs'
          }, {
            name: 'alt',
            attributeName: {
              localPart: 'alt'
            },
            type: 'attribute'
          }, {
            name: 'chromakey',
            attributeName: {
              localPart: 'chromakey'
            },
            type: 'attribute'
          }, {
            name: 'clazz',
            attributeName: {
              localPart: 'class'
            },
            type: 'attribute'
          }, {
            name: 'coordorigin',
            attributeName: {
              localPart: 'coordorigin'
            },
            type: 'attribute'
          }, {
            name: 'coordsize',
            attributeName: {
              localPart: 'coordsize'
            },
            type: 'attribute'
          }, {
            name: 'endAngle',
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'endAngle'
            },
            type: 'attribute'
          }, {
            name: 'fillcolor',
            attributeName: {
              localPart: 'fillcolor'
            },
            type: 'attribute'
          }, {
            name: 'filled',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'filled'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'insetpen',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'insetpen'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'print',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'print'
            },
            type: 'attribute'
          }, {
            name: 'startAngle',
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'startAngle'
            },
            type: 'attribute'
          }, {
            name: 'strokecolor',
            attributeName: {
              localPart: 'strokecolor'
            },
            type: 'attribute'
          }, {
            name: 'stroked',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'stroked'
            },
            type: 'attribute'
          }, {
            name: 'strokeweight',
            attributeName: {
              localPart: 'strokeweight'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'target',
            attributeName: {
              localPart: 'target'
            },
            type: 'attribute'
          }, {
            name: 'title',
            attributeName: {
              localPart: 'title'
            },
            type: 'attribute'
          }, {
            name: 'wrapcoords',
            attributeName: {
              localPart: 'wrapcoords'
            },
            type: 'attribute'
          }, {
            name: 'allowincell',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'allowoverlap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'borderbottomcolor',
            type: 'attribute'
          }, {
            name: 'borderleftcolor',
            type: 'attribute'
          }, {
            name: 'borderrightcolor',
            type: 'attribute'
          }, {
            name: 'bordertopcolor',
            type: 'attribute'
          }, {
            name: 'bullet',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'button',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'bwmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwnormal',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwpure',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'clip',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'cliptowrap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'connectortype',
            typeInfo: 'org_docx4j_vml_officedrawing.STConnectorType',
            defaultValue: 'straight',
            type: 'attribute'
          }, {
            name: 'dgmlayout',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmlayoutmru',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmnodekind',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'doubleclicknotify',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'forcedash',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hr',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hralign',
            typeInfo: 'org_docx4j_vml_officedrawing.STHrAlign',
            defaultValue: 'left',
            type: 'attribute'
          }, {
            name: 'hrnoshade',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hrpct',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'hrstd',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'insetmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STInsetMode',
            defaultValue: 'custom',
            type: 'attribute'
          }, {
            name: 'ole',
            values: ['', 't', 'f', 'true', 'false'],
            type: 'attribute'
          }, {
            name: 'oleicon',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'oned',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'preferrelative',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'regroupid',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'spid',
            type: 'attribute'
          }, {
            name: 'spt',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'userdrawn',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'userhidden',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        localName: 'CTBackground',
        typeName: 'CT_Background',
        propertyInfos: [{
            name: 'fill',
            typeInfo: '.CTFill'
          }, {
            name: 'fillcolor',
            attributeName: {
              localPart: 'fillcolor'
            },
            type: 'attribute'
          }, {
            name: 'filled',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'filled'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'bwmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwnormal',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwpure',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'targetscreensize',
            values: ['544,376', '640,480', '720,512', '800,600', '1024,768', '1152,862'],
            type: 'attribute'
          }]
      }, {
        localName: 'CTCurve',
        typeName: 'CT_Curve',
        propertyInfos: [{
            name: 'egShapeElements',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'path',
                typeInfo: '.CTPath'
              }, {
                elementName: 'formulas',
                typeInfo: '.CTFormulas'
              }, {
                elementName: 'handles',
                typeInfo: '.CTHandles'
              }, {
                elementName: 'fill',
                typeInfo: '.CTFill'
              }, {
                elementName: 'stroke',
                typeInfo: '.CTStroke'
              }, {
                elementName: 'shadow',
                typeInfo: '.CTShadow'
              }, {
                elementName: 'textbox',
                typeInfo: '.CTTextbox'
              }, {
                elementName: 'textpath',
                typeInfo: '.CTTextPath'
              }, {
                elementName: 'imagedata',
                typeInfo: '.CTImageData'
              }, {
                elementName: {
                  localPart: 'skew',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSkew'
              }, {
                elementName: {
                  localPart: 'extrusion',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTExtrusion'
              }, {
                elementName: {
                  localPart: 'callout',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTCallout'
              }, {
                elementName: {
                  localPart: 'lock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTLock'
              }, {
                elementName: {
                  localPart: 'clippath',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTClipPath'
              }, {
                elementName: {
                  localPart: 'signatureline',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSignatureLine'
              }, {
                elementName: {
                  localPart: 'wrap',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTWrap'
              }, {
                elementName: {
                  localPart: 'anchorlock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTAnchorLock'
              }, {
                elementName: {
                  localPart: 'bordertop',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderbottom',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderleft',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderright',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'ClientData',
                  namespaceURI: 'urn:schemas-microsoft-com:office:excel'
                },
                typeInfo: 'org_docx4j_vml_spreadsheetDrawing.CTClientData'
              }, {
                elementName: {
                  localPart: 'textdata',
                  namespaceURI: 'urn:schemas-microsoft-com:office:powerpoint'
                },
                typeInfo: 'org_docx4j_vml_presentationDrawing.CTRel'
              }],
            type: 'elementRefs'
          }, {
            name: 'alt',
            attributeName: {
              localPart: 'alt'
            },
            type: 'attribute'
          }, {
            name: 'chromakey',
            attributeName: {
              localPart: 'chromakey'
            },
            type: 'attribute'
          }, {
            name: 'clazz',
            attributeName: {
              localPart: 'class'
            },
            type: 'attribute'
          }, {
            name: 'control1',
            attributeName: {
              localPart: 'control1'
            },
            type: 'attribute'
          }, {
            name: 'control2',
            attributeName: {
              localPart: 'control2'
            },
            type: 'attribute'
          }, {
            name: 'coordorigin',
            attributeName: {
              localPart: 'coordorigin'
            },
            type: 'attribute'
          }, {
            name: 'coordsize',
            attributeName: {
              localPart: 'coordsize'
            },
            type: 'attribute'
          }, {
            name: 'fillcolor',
            attributeName: {
              localPart: 'fillcolor'
            },
            type: 'attribute'
          }, {
            name: 'filled',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'filled'
            },
            type: 'attribute'
          }, {
            name: 'from',
            attributeName: {
              localPart: 'from'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'insetpen',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'insetpen'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'print',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'print'
            },
            type: 'attribute'
          }, {
            name: 'strokecolor',
            attributeName: {
              localPart: 'strokecolor'
            },
            type: 'attribute'
          }, {
            name: 'stroked',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'stroked'
            },
            type: 'attribute'
          }, {
            name: 'strokeweight',
            attributeName: {
              localPart: 'strokeweight'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'target',
            attributeName: {
              localPart: 'target'
            },
            type: 'attribute'
          }, {
            name: 'title',
            attributeName: {
              localPart: 'title'
            },
            type: 'attribute'
          }, {
            name: 'to',
            attributeName: {
              localPart: 'to'
            },
            type: 'attribute'
          }, {
            name: 'wrapcoords',
            attributeName: {
              localPart: 'wrapcoords'
            },
            type: 'attribute'
          }, {
            name: 'allowincell',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'allowoverlap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'borderbottomcolor',
            type: 'attribute'
          }, {
            name: 'borderleftcolor',
            type: 'attribute'
          }, {
            name: 'borderrightcolor',
            type: 'attribute'
          }, {
            name: 'bordertopcolor',
            type: 'attribute'
          }, {
            name: 'bullet',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'button',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'bwmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwnormal',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwpure',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'clip',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'cliptowrap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'connectortype',
            typeInfo: 'org_docx4j_vml_officedrawing.STConnectorType',
            defaultValue: 'straight',
            type: 'attribute'
          }, {
            name: 'dgmlayout',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmlayoutmru',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmnodekind',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'doubleclicknotify',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'forcedash',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hr',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hralign',
            typeInfo: 'org_docx4j_vml_officedrawing.STHrAlign',
            defaultValue: 'left',
            type: 'attribute'
          }, {
            name: 'hrnoshade',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hrpct',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'hrstd',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'insetmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STInsetMode',
            defaultValue: 'custom',
            type: 'attribute'
          }, {
            name: 'ole',
            values: ['', 't', 'f', 'true', 'false'],
            type: 'attribute'
          }, {
            name: 'oleicon',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'oned',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'preferrelative',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'regroupid',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'spid',
            type: 'attribute'
          }, {
            name: 'spt',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'userdrawn',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'userhidden',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        localName: 'CTF',
        typeName: 'CT_F',
        propertyInfos: [{
            name: 'eqn',
            attributeName: {
              localPart: 'eqn'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTFill',
        typeName: 'CT_Fill',
        propertyInfos: [{
            name: 'fill',
            elementName: {
              localPart: 'fill',
              namespaceURI: 'urn:schemas-microsoft-com:office:office'
            },
            typeInfo: 'org_docx4j_vml_officedrawing.CTFill'
          }, {
            name: 'alignshape',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'alignshape'
            },
            type: 'attribute'
          }, {
            name: 'angle',
            typeInfo: 'Decimal',
            attributeName: {
              localPart: 'angle'
            },
            type: 'attribute'
          }, {
            name: 'aspect',
            typeInfo: '.STImageAspect',
            attributeName: {
              localPart: 'aspect'
            },
            type: 'attribute'
          }, {
            name: 'color',
            attributeName: {
              localPart: 'color'
            },
            type: 'attribute'
          }, {
            name: 'color2',
            attributeName: {
              localPart: 'color2'
            },
            type: 'attribute'
          }, {
            name: 'colors',
            attributeName: {
              localPart: 'colors'
            },
            type: 'attribute'
          }, {
            name: 'focus',
            attributeName: {
              localPart: 'focus'
            },
            type: 'attribute'
          }, {
            name: 'focusposition',
            attributeName: {
              localPart: 'focusposition'
            },
            type: 'attribute'
          }, {
            name: 'focussize',
            attributeName: {
              localPart: 'focussize'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'method',
            typeInfo: '.STFillMethod',
            attributeName: {
              localPart: 'method'
            },
            type: 'attribute'
          }, {
            name: 'on',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'on'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'origin',
            attributeName: {
              localPart: 'origin'
            },
            type: 'attribute'
          }, {
            name: 'position',
            attributeName: {
              localPart: 'position'
            },
            type: 'attribute'
          }, {
            name: 'recolor',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'recolor'
            },
            type: 'attribute'
          }, {
            name: 'rotate',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'rotate'
            },
            type: 'attribute'
          }, {
            name: 'size',
            attributeName: {
              localPart: 'size'
            },
            type: 'attribute'
          }, {
            name: 'src',
            attributeName: {
              localPart: 'src'
            },
            type: 'attribute'
          }, {
            name: 'type',
            typeInfo: '.STFillType',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: 'id',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }, {
            name: 'althref',
            type: 'attribute'
          }, {
            name: 'detectmouseclick',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'href',
            type: 'attribute'
          }, {
            name: 'opacity2',
            type: 'attribute'
          }, {
            name: 'relid',
            type: 'attribute'
          }, {
            name: 'title',
            type: 'attribute'
          }]
      }, {
        localName: 'CTFormulas',
        typeName: 'CT_Formulas',
        propertyInfos: [{
            name: 'f',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTF'
          }]
      }, {
        localName: 'CTGroup',
        typeName: 'CT_Group',
        propertyInfos: [{
            name: 'pathOrFormulasOrHandles',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'path',
                typeInfo: '.CTPath'
              }, {
                elementName: 'formulas',
                typeInfo: '.CTFormulas'
              }, {
                elementName: 'handles',
                typeInfo: '.CTHandles'
              }, {
                elementName: 'fill',
                typeInfo: '.CTFill'
              }, {
                elementName: 'stroke',
                typeInfo: '.CTStroke'
              }, {
                elementName: 'shadow',
                typeInfo: '.CTShadow'
              }, {
                elementName: 'textbox',
                typeInfo: '.CTTextbox'
              }, {
                elementName: 'textpath',
                typeInfo: '.CTTextPath'
              }, {
                elementName: 'imagedata',
                typeInfo: '.CTImageData'
              }, {
                elementName: {
                  localPart: 'skew',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSkew'
              }, {
                elementName: {
                  localPart: 'extrusion',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTExtrusion'
              }, {
                elementName: {
                  localPart: 'callout',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTCallout'
              }, {
                elementName: {
                  localPart: 'lock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTLock'
              }, {
                elementName: {
                  localPart: 'clippath',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTClipPath'
              }, {
                elementName: {
                  localPart: 'signatureline',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSignatureLine'
              }, {
                elementName: {
                  localPart: 'wrap',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTWrap'
              }, {
                elementName: {
                  localPart: 'anchorlock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTAnchorLock'
              }, {
                elementName: {
                  localPart: 'bordertop',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderbottom',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderleft',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderright',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'ClientData',
                  namespaceURI: 'urn:schemas-microsoft-com:office:excel'
                },
                typeInfo: 'org_docx4j_vml_spreadsheetDrawing.CTClientData'
              }, {
                elementName: {
                  localPart: 'textdata',
                  namespaceURI: 'urn:schemas-microsoft-com:office:powerpoint'
                },
                typeInfo: 'org_docx4j_vml_presentationDrawing.CTRel'
              }, {
                elementName: 'group',
                typeInfo: '.CTGroup'
              }, {
                elementName: 'shape',
                typeInfo: '.CTShape'
              }, {
                elementName: 'shapetype',
                typeInfo: '.CTShapetype'
              }, {
                elementName: 'arc',
                typeInfo: '.CTArc'
              }, {
                elementName: 'curve',
                typeInfo: '.CTCurve'
              }, {
                elementName: 'image',
                typeInfo: '.CTImage'
              }, {
                elementName: 'line',
                typeInfo: '.CTLine'
              }, {
                elementName: 'oval',
                typeInfo: '.CTOval'
              }, {
                elementName: 'polyline',
                typeInfo: '.CTPolyLine'
              }, {
                elementName: 'rect',
                typeInfo: '.CTRect'
              }, {
                elementName: 'roundrect',
                typeInfo: '.CTRoundRect'
              }, {
                elementName: {
                  localPart: 'diagram',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTDiagram'
              }],
            type: 'elementRefs'
          }, {
            name: 'alt',
            attributeName: {
              localPart: 'alt'
            },
            type: 'attribute'
          }, {
            name: 'clazz',
            attributeName: {
              localPart: 'class'
            },
            type: 'attribute'
          }, {
            name: 'coordorigin',
            attributeName: {
              localPart: 'coordorigin'
            },
            type: 'attribute'
          }, {
            name: 'coordsize',
            attributeName: {
              localPart: 'coordsize'
            },
            type: 'attribute'
          }, {
            name: 'editas',
            typeInfo: '.STEditAs',
            attributeName: {
              localPart: 'editas'
            },
            type: 'attribute'
          }, {
            name: 'fillcolor',
            attributeName: {
              localPart: 'fillcolor'
            },
            type: 'attribute'
          }, {
            name: 'filled',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'filled'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'print',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'print'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'target',
            attributeName: {
              localPart: 'target'
            },
            type: 'attribute'
          }, {
            name: 'title',
            attributeName: {
              localPart: 'title'
            },
            type: 'attribute'
          }, {
            name: 'wrapcoords',
            attributeName: {
              localPart: 'wrapcoords'
            },
            type: 'attribute'
          }, {
            name: 'allowincell',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'allowoverlap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'borderbottomcolor',
            type: 'attribute'
          }, {
            name: 'borderleftcolor',
            type: 'attribute'
          }, {
            name: 'borderrightcolor',
            type: 'attribute'
          }, {
            name: 'bordertopcolor',
            type: 'attribute'
          }, {
            name: 'bullet',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'button',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'dgmlayout',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmlayoutmru',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmnodekind',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'doubleclicknotify',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hr',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hralign',
            typeInfo: 'org_docx4j_vml_officedrawing.STHrAlign',
            defaultValue: 'left',
            type: 'attribute'
          }, {
            name: 'hrnoshade',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hrpct',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'hrstd',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'insetmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STInsetMode',
            defaultValue: 'custom',
            type: 'attribute'
          }, {
            name: 'oned',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'regroupid',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'spid',
            type: 'attribute'
          }, {
            name: 'tablelimits',
            type: 'attribute'
          }, {
            name: 'tableproperties',
            type: 'attribute'
          }, {
            name: 'userdrawn',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'userhidden',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        localName: 'CTH',
        typeName: 'CT_H',
        propertyInfos: [{
            name: 'invx',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'invx'
            },
            type: 'attribute'
          }, {
            name: 'invy',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'invy'
            },
            type: 'attribute'
          }, {
            name: 'map',
            attributeName: {
              localPart: 'map'
            },
            type: 'attribute'
          }, {
            name: 'polar',
            attributeName: {
              localPart: 'polar'
            },
            type: 'attribute'
          }, {
            name: 'position',
            attributeName: {
              localPart: 'position'
            },
            type: 'attribute'
          }, {
            name: 'radiusrange',
            attributeName: {
              localPart: 'radiusrange'
            },
            type: 'attribute'
          }, {
            name: '_switch',
            values: ['t', 'f', 'true', 'false', ''],
            attributeName: {
              localPart: 'switch'
            },
            type: 'attribute'
          }, {
            name: 'xrange',
            attributeName: {
              localPart: 'xrange'
            },
            type: 'attribute'
          }, {
            name: 'yrange',
            attributeName: {
              localPart: 'yrange'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTHandles',
        typeName: 'CT_Handles',
        propertyInfos: [{
            name: 'h',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTH'
          }]
      }, {
        localName: 'CTImage',
        typeName: 'CT_Image',
        propertyInfos: [{
            name: 'egShapeElements',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'path',
                typeInfo: '.CTPath'
              }, {
                elementName: 'formulas',
                typeInfo: '.CTFormulas'
              }, {
                elementName: 'handles',
                typeInfo: '.CTHandles'
              }, {
                elementName: 'fill',
                typeInfo: '.CTFill'
              }, {
                elementName: 'stroke',
                typeInfo: '.CTStroke'
              }, {
                elementName: 'shadow',
                typeInfo: '.CTShadow'
              }, {
                elementName: 'textbox',
                typeInfo: '.CTTextbox'
              }, {
                elementName: 'textpath',
                typeInfo: '.CTTextPath'
              }, {
                elementName: 'imagedata',
                typeInfo: '.CTImageData'
              }, {
                elementName: {
                  localPart: 'skew',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSkew'
              }, {
                elementName: {
                  localPart: 'extrusion',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTExtrusion'
              }, {
                elementName: {
                  localPart: 'callout',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTCallout'
              }, {
                elementName: {
                  localPart: 'lock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTLock'
              }, {
                elementName: {
                  localPart: 'clippath',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTClipPath'
              }, {
                elementName: {
                  localPart: 'signatureline',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSignatureLine'
              }, {
                elementName: {
                  localPart: 'wrap',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTWrap'
              }, {
                elementName: {
                  localPart: 'anchorlock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTAnchorLock'
              }, {
                elementName: {
                  localPart: 'bordertop',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderbottom',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderleft',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderright',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'ClientData',
                  namespaceURI: 'urn:schemas-microsoft-com:office:excel'
                },
                typeInfo: 'org_docx4j_vml_spreadsheetDrawing.CTClientData'
              }, {
                elementName: {
                  localPart: 'textdata',
                  namespaceURI: 'urn:schemas-microsoft-com:office:powerpoint'
                },
                typeInfo: 'org_docx4j_vml_presentationDrawing.CTRel'
              }],
            type: 'elementRefs'
          }, {
            name: 'alt',
            attributeName: {
              localPart: 'alt'
            },
            type: 'attribute'
          }, {
            name: 'bilevel',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'bilevel'
            },
            type: 'attribute'
          }, {
            name: 'blacklevel',
            attributeName: {
              localPart: 'blacklevel'
            },
            type: 'attribute'
          }, {
            name: 'chromakey',
            attributeName: {
              localPart: 'chromakey'
            },
            type: 'attribute'
          }, {
            name: 'clazz',
            attributeName: {
              localPart: 'class'
            },
            type: 'attribute'
          }, {
            name: 'coordorigin',
            attributeName: {
              localPart: 'coordorigin'
            },
            type: 'attribute'
          }, {
            name: 'coordsize',
            attributeName: {
              localPart: 'coordsize'
            },
            type: 'attribute'
          }, {
            name: 'cropbottom',
            attributeName: {
              localPart: 'cropbottom'
            },
            type: 'attribute'
          }, {
            name: 'cropleft',
            attributeName: {
              localPart: 'cropleft'
            },
            type: 'attribute'
          }, {
            name: 'cropright',
            attributeName: {
              localPart: 'cropright'
            },
            type: 'attribute'
          }, {
            name: 'croptop',
            attributeName: {
              localPart: 'croptop'
            },
            type: 'attribute'
          }, {
            name: 'fillcolor',
            attributeName: {
              localPart: 'fillcolor'
            },
            type: 'attribute'
          }, {
            name: 'filled',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'filled'
            },
            type: 'attribute'
          }, {
            name: 'gain',
            attributeName: {
              localPart: 'gain'
            },
            type: 'attribute'
          }, {
            name: 'gamma',
            attributeName: {
              localPart: 'gamma'
            },
            type: 'attribute'
          }, {
            name: 'grayscale',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'grayscale'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'insetpen',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'insetpen'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'print',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'print'
            },
            type: 'attribute'
          }, {
            name: 'src',
            attributeName: {
              localPart: 'src'
            },
            type: 'attribute'
          }, {
            name: 'strokecolor',
            attributeName: {
              localPart: 'strokecolor'
            },
            type: 'attribute'
          }, {
            name: 'stroked',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'stroked'
            },
            type: 'attribute'
          }, {
            name: 'strokeweight',
            attributeName: {
              localPart: 'strokeweight'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'target',
            attributeName: {
              localPart: 'target'
            },
            type: 'attribute'
          }, {
            name: 'title',
            attributeName: {
              localPart: 'title'
            },
            type: 'attribute'
          }, {
            name: 'wrapcoords',
            attributeName: {
              localPart: 'wrapcoords'
            },
            type: 'attribute'
          }, {
            name: 'allowincell',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'allowoverlap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'borderbottomcolor',
            type: 'attribute'
          }, {
            name: 'borderleftcolor',
            type: 'attribute'
          }, {
            name: 'borderrightcolor',
            type: 'attribute'
          }, {
            name: 'bordertopcolor',
            type: 'attribute'
          }, {
            name: 'bullet',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'button',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'bwmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwnormal',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwpure',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'clip',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'cliptowrap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'connectortype',
            typeInfo: 'org_docx4j_vml_officedrawing.STConnectorType',
            defaultValue: 'straight',
            type: 'attribute'
          }, {
            name: 'dgmlayout',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmlayoutmru',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmnodekind',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'doubleclicknotify',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'forcedash',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hr',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hralign',
            typeInfo: 'org_docx4j_vml_officedrawing.STHrAlign',
            defaultValue: 'left',
            type: 'attribute'
          }, {
            name: 'hrnoshade',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hrpct',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'hrstd',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'insetmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STInsetMode',
            defaultValue: 'custom',
            type: 'attribute'
          }, {
            name: 'ole',
            values: ['', 't', 'f', 'true', 'false'],
            type: 'attribute'
          }, {
            name: 'oleicon',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'oned',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'preferrelative',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'regroupid',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'spid',
            type: 'attribute'
          }, {
            name: 'spt',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'userdrawn',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'userhidden',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        localName: 'CTImageData',
        typeName: 'CT_ImageData',
        propertyInfos: [{
            name: 'bilevel',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'bilevel'
            },
            type: 'attribute'
          }, {
            name: 'blacklevel',
            attributeName: {
              localPart: 'blacklevel'
            },
            type: 'attribute'
          }, {
            name: 'chromakey',
            attributeName: {
              localPart: 'chromakey'
            },
            type: 'attribute'
          }, {
            name: 'cropbottom',
            attributeName: {
              localPart: 'cropbottom'
            },
            type: 'attribute'
          }, {
            name: 'cropleft',
            attributeName: {
              localPart: 'cropleft'
            },
            type: 'attribute'
          }, {
            name: 'cropright',
            attributeName: {
              localPart: 'cropright'
            },
            type: 'attribute'
          }, {
            name: 'croptop',
            attributeName: {
              localPart: 'croptop'
            },
            type: 'attribute'
          }, {
            name: 'embosscolor',
            attributeName: {
              localPart: 'embosscolor'
            },
            type: 'attribute'
          }, {
            name: 'gain',
            attributeName: {
              localPart: 'gain'
            },
            type: 'attribute'
          }, {
            name: 'gamma',
            attributeName: {
              localPart: 'gamma'
            },
            type: 'attribute'
          }, {
            name: 'grayscale',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'grayscale'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'recolortarget',
            attributeName: {
              localPart: 'recolortarget'
            },
            type: 'attribute'
          }, {
            name: 'src',
            attributeName: {
              localPart: 'src'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }, {
            name: 'id',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }, {
            name: 'pict',
            attributeName: {
              localPart: 'pict',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }, {
            name: 'althref',
            type: 'attribute'
          }, {
            name: 'detectmouseclick',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'ohref',
            attributeName: 'href',
            type: 'attribute'
          }, {
            name: 'movie',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'oleid',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'relid',
            type: 'attribute'
          }, {
            name: 'title',
            type: 'attribute'
          }]
      }, {
        localName: 'CTLine',
        typeName: 'CT_Line',
        propertyInfos: [{
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'from',
            attributeName: {
              localPart: 'from'
            },
            type: 'attribute'
          }, {
            name: 'to',
            attributeName: {
              localPart: 'to'
            },
            type: 'attribute'
          }, {
            name: 'egShapeElements',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'path',
                typeInfo: '.CTPath'
              }, {
                elementName: 'formulas',
                typeInfo: '.CTFormulas'
              }, {
                elementName: 'handles',
                typeInfo: '.CTHandles'
              }, {
                elementName: 'fill',
                typeInfo: '.CTFill'
              }, {
                elementName: 'stroke',
                typeInfo: '.CTStroke'
              }, {
                elementName: 'shadow',
                typeInfo: '.CTShadow'
              }, {
                elementName: 'textbox',
                typeInfo: '.CTTextbox'
              }, {
                elementName: 'textpath',
                typeInfo: '.CTTextPath'
              }, {
                elementName: 'imagedata',
                typeInfo: '.CTImageData'
              }, {
                elementName: {
                  localPart: 'skew',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSkew'
              }, {
                elementName: {
                  localPart: 'extrusion',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTExtrusion'
              }, {
                elementName: {
                  localPart: 'callout',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTCallout'
              }, {
                elementName: {
                  localPart: 'lock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTLock'
              }, {
                elementName: {
                  localPart: 'clippath',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTClipPath'
              }, {
                elementName: {
                  localPart: 'signatureline',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSignatureLine'
              }, {
                elementName: {
                  localPart: 'wrap',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTWrap'
              }, {
                elementName: {
                  localPart: 'anchorlock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTAnchorLock'
              }, {
                elementName: {
                  localPart: 'bordertop',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderbottom',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderleft',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderright',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'ClientData',
                  namespaceURI: 'urn:schemas-microsoft-com:office:excel'
                },
                typeInfo: 'org_docx4j_vml_spreadsheetDrawing.CTClientData'
              }, {
                elementName: {
                  localPart: 'textdata',
                  namespaceURI: 'urn:schemas-microsoft-com:office:powerpoint'
                },
                typeInfo: 'org_docx4j_vml_presentationDrawing.CTRel'
              }],
            type: 'elementRefs'
          }, {
            name: 'alt',
            attributeName: {
              localPart: 'alt'
            },
            type: 'attribute'
          }, {
            name: 'chromakey',
            attributeName: {
              localPart: 'chromakey'
            },
            type: 'attribute'
          }, {
            name: 'clazz',
            attributeName: {
              localPart: 'class'
            },
            type: 'attribute'
          }, {
            name: 'coordorigin',
            attributeName: {
              localPart: 'coordorigin'
            },
            type: 'attribute'
          }, {
            name: 'coordsize',
            attributeName: {
              localPart: 'coordsize'
            },
            type: 'attribute'
          }, {
            name: 'fillcolor',
            attributeName: {
              localPart: 'fillcolor'
            },
            type: 'attribute'
          }, {
            name: 'filled',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'filled'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href'
            },
            type: 'attribute'
          }, {
            name: 'insetpen',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'insetpen'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'print',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'print'
            },
            type: 'attribute'
          }, {
            name: 'strokecolor',
            attributeName: {
              localPart: 'strokecolor'
            },
            type: 'attribute'
          }, {
            name: 'stroked',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'stroked'
            },
            type: 'attribute'
          }, {
            name: 'strokeweight',
            attributeName: {
              localPart: 'strokeweight'
            },
            type: 'attribute'
          }, {
            name: 'target',
            attributeName: {
              localPart: 'target'
            },
            type: 'attribute'
          }, {
            name: 'title',
            attributeName: {
              localPart: 'title'
            },
            type: 'attribute'
          }, {
            name: 'wrapcoords',
            attributeName: {
              localPart: 'wrapcoords'
            },
            type: 'attribute'
          }, {
            name: 'allowincell',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'allowoverlap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'borderbottomcolor',
            type: 'attribute'
          }, {
            name: 'borderleftcolor',
            type: 'attribute'
          }, {
            name: 'borderrightcolor',
            type: 'attribute'
          }, {
            name: 'bordertopcolor',
            type: 'attribute'
          }, {
            name: 'bullet',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'button',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'bwmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwnormal',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwpure',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'clip',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'cliptowrap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'connectortype',
            typeInfo: 'org_docx4j_vml_officedrawing.STConnectorType',
            defaultValue: 'straight',
            type: 'attribute'
          }, {
            name: 'dgmlayout',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmlayoutmru',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmnodekind',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'doubleclicknotify',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'forcedash',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hr',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hralign',
            typeInfo: 'org_docx4j_vml_officedrawing.STHrAlign',
            defaultValue: 'left',
            type: 'attribute'
          }, {
            name: 'hrnoshade',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hrpct',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'hrstd',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'insetmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STInsetMode',
            defaultValue: 'custom',
            type: 'attribute'
          }, {
            name: 'ole',
            values: ['', 't', 'f', 'true', 'false'],
            type: 'attribute'
          }, {
            name: 'oleicon',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'oned',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'preferrelative',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'regroupid',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'spid',
            type: 'attribute'
          }, {
            name: 'spt',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'userdrawn',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'userhidden',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        localName: 'CTOval',
        typeName: 'CT_Oval',
        propertyInfos: [{
            name: 'pathOrFormulasOrHandles',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'path',
                typeInfo: '.CTPath'
              }, {
                elementName: 'formulas',
                typeInfo: '.CTFormulas'
              }, {
                elementName: 'handles',
                typeInfo: '.CTHandles'
              }, {
                elementName: 'fill',
                typeInfo: '.CTFill'
              }, {
                elementName: 'stroke',
                typeInfo: '.CTStroke'
              }, {
                elementName: 'shadow',
                typeInfo: '.CTShadow'
              }, {
                elementName: 'textbox',
                typeInfo: '.CTTextbox'
              }, {
                elementName: 'textpath',
                typeInfo: '.CTTextPath'
              }, {
                elementName: 'imagedata',
                typeInfo: '.CTImageData'
              }, {
                elementName: {
                  localPart: 'skew',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSkew'
              }, {
                elementName: {
                  localPart: 'extrusion',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTExtrusion'
              }, {
                elementName: {
                  localPart: 'callout',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTCallout'
              }, {
                elementName: {
                  localPart: 'lock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTLock'
              }, {
                elementName: {
                  localPart: 'clippath',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTClipPath'
              }, {
                elementName: {
                  localPart: 'signatureline',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSignatureLine'
              }, {
                elementName: {
                  localPart: 'wrap',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTWrap'
              }, {
                elementName: {
                  localPart: 'anchorlock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTAnchorLock'
              }, {
                elementName: {
                  localPart: 'bordertop',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderbottom',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderleft',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderright',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'ClientData',
                  namespaceURI: 'urn:schemas-microsoft-com:office:excel'
                },
                typeInfo: 'org_docx4j_vml_spreadsheetDrawing.CTClientData'
              }, {
                elementName: {
                  localPart: 'textdata',
                  namespaceURI: 'urn:schemas-microsoft-com:office:powerpoint'
                },
                typeInfo: 'org_docx4j_vml_presentationDrawing.CTRel'
              }],
            type: 'elementRefs'
          }, {
            name: 'alt',
            attributeName: {
              localPart: 'alt'
            },
            type: 'attribute'
          }, {
            name: 'chromakey',
            attributeName: {
              localPart: 'chromakey'
            },
            type: 'attribute'
          }, {
            name: 'clazz',
            attributeName: {
              localPart: 'class'
            },
            type: 'attribute'
          }, {
            name: 'coordorigin',
            attributeName: {
              localPart: 'coordorigin'
            },
            type: 'attribute'
          }, {
            name: 'coordsize',
            attributeName: {
              localPart: 'coordsize'
            },
            type: 'attribute'
          }, {
            name: 'fillcolor',
            attributeName: {
              localPart: 'fillcolor'
            },
            type: 'attribute'
          }, {
            name: 'filled',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'filled'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'insetpen',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'insetpen'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'print',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'print'
            },
            type: 'attribute'
          }, {
            name: 'strokecolor',
            attributeName: {
              localPart: 'strokecolor'
            },
            type: 'attribute'
          }, {
            name: 'stroked',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'stroked'
            },
            type: 'attribute'
          }, {
            name: 'strokeweight',
            attributeName: {
              localPart: 'strokeweight'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'target',
            attributeName: {
              localPart: 'target'
            },
            type: 'attribute'
          }, {
            name: 'title',
            attributeName: {
              localPart: 'title'
            },
            type: 'attribute'
          }, {
            name: 'wrapcoords',
            attributeName: {
              localPart: 'wrapcoords'
            },
            type: 'attribute'
          }, {
            name: 'allowincell',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'allowoverlap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'borderbottomcolor',
            type: 'attribute'
          }, {
            name: 'borderleftcolor',
            type: 'attribute'
          }, {
            name: 'borderrightcolor',
            type: 'attribute'
          }, {
            name: 'bordertopcolor',
            type: 'attribute'
          }, {
            name: 'bullet',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'button',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'bwmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwnormal',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwpure',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'clip',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'cliptowrap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'connectortype',
            typeInfo: 'org_docx4j_vml_officedrawing.STConnectorType',
            defaultValue: 'straight',
            type: 'attribute'
          }, {
            name: 'dgmlayout',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmlayoutmru',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmnodekind',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'doubleclicknotify',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'forcedash',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hr',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hralign',
            typeInfo: 'org_docx4j_vml_officedrawing.STHrAlign',
            defaultValue: 'left',
            type: 'attribute'
          }, {
            name: 'hrnoshade',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hrpct',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'hrstd',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'insetmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STInsetMode',
            defaultValue: 'custom',
            type: 'attribute'
          }, {
            name: 'ole',
            values: ['', 't', 'f', 'true', 'false'],
            type: 'attribute'
          }, {
            name: 'oleicon',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'oned',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'preferrelative',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'regroupid',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'spid',
            type: 'attribute'
          }, {
            name: 'spt',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'userdrawn',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'userhidden',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        localName: 'CTPath',
        typeName: 'CT_Path',
        propertyInfos: [{
            name: 'arrowok',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'arrowok'
            },
            type: 'attribute'
          }, {
            name: 'fillok',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'fillok'
            },
            type: 'attribute'
          }, {
            name: 'gradientshapeok',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'gradientshapeok'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'insetpenok',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'insetpenok'
            },
            type: 'attribute'
          }, {
            name: 'limo',
            attributeName: {
              localPart: 'limo'
            },
            type: 'attribute'
          }, {
            name: 'shadowok',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'shadowok'
            },
            type: 'attribute'
          }, {
            name: 'strokeok',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'strokeok'
            },
            type: 'attribute'
          }, {
            name: 'textboxrect',
            attributeName: {
              localPart: 'textboxrect'
            },
            type: 'attribute'
          }, {
            name: 'textpathok',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'textpathok'
            },
            type: 'attribute'
          }, {
            name: 'v',
            attributeName: {
              localPart: 'v'
            },
            type: 'attribute'
          }, {
            name: 'connectangles',
            type: 'attribute'
          }, {
            name: 'connectlocs',
            type: 'attribute'
          }, {
            name: 'connecttype',
            typeInfo: 'org_docx4j_vml_officedrawing.STConnectType',
            type: 'attribute'
          }, {
            name: 'extrusionok',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        localName: 'CTPolyLine',
        typeName: 'CT_PolyLine',
        propertyInfos: [{
            name: 'pathOrFormulasOrHandles',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'path',
                typeInfo: '.CTPath'
              }, {
                elementName: 'formulas',
                typeInfo: '.CTFormulas'
              }, {
                elementName: 'handles',
                typeInfo: '.CTHandles'
              }, {
                elementName: 'fill',
                typeInfo: '.CTFill'
              }, {
                elementName: 'stroke',
                typeInfo: '.CTStroke'
              }, {
                elementName: 'shadow',
                typeInfo: '.CTShadow'
              }, {
                elementName: 'textbox',
                typeInfo: '.CTTextbox'
              }, {
                elementName: 'textpath',
                typeInfo: '.CTTextPath'
              }, {
                elementName: 'imagedata',
                typeInfo: '.CTImageData'
              }, {
                elementName: {
                  localPart: 'skew',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSkew'
              }, {
                elementName: {
                  localPart: 'extrusion',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTExtrusion'
              }, {
                elementName: {
                  localPart: 'callout',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTCallout'
              }, {
                elementName: {
                  localPart: 'lock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTLock'
              }, {
                elementName: {
                  localPart: 'clippath',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTClipPath'
              }, {
                elementName: {
                  localPart: 'signatureline',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSignatureLine'
              }, {
                elementName: {
                  localPart: 'wrap',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTWrap'
              }, {
                elementName: {
                  localPart: 'anchorlock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTAnchorLock'
              }, {
                elementName: {
                  localPart: 'bordertop',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderbottom',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderleft',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderright',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'ClientData',
                  namespaceURI: 'urn:schemas-microsoft-com:office:excel'
                },
                typeInfo: 'org_docx4j_vml_spreadsheetDrawing.CTClientData'
              }, {
                elementName: {
                  localPart: 'textdata',
                  namespaceURI: 'urn:schemas-microsoft-com:office:powerpoint'
                },
                typeInfo: 'org_docx4j_vml_presentationDrawing.CTRel'
              }, {
                elementName: {
                  localPart: 'ink',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTInk'
              }],
            type: 'elementRefs'
          }, {
            name: 'alt',
            attributeName: {
              localPart: 'alt'
            },
            type: 'attribute'
          }, {
            name: 'chromakey',
            attributeName: {
              localPart: 'chromakey'
            },
            type: 'attribute'
          }, {
            name: 'clazz',
            attributeName: {
              localPart: 'class'
            },
            type: 'attribute'
          }, {
            name: 'coordorigin',
            attributeName: {
              localPart: 'coordorigin'
            },
            type: 'attribute'
          }, {
            name: 'coordsize',
            attributeName: {
              localPart: 'coordsize'
            },
            type: 'attribute'
          }, {
            name: 'fillcolor',
            attributeName: {
              localPart: 'fillcolor'
            },
            type: 'attribute'
          }, {
            name: 'filled',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'filled'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'insetpen',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'insetpen'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'points',
            attributeName: {
              localPart: 'points'
            },
            type: 'attribute'
          }, {
            name: 'print',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'print'
            },
            type: 'attribute'
          }, {
            name: 'strokecolor',
            attributeName: {
              localPart: 'strokecolor'
            },
            type: 'attribute'
          }, {
            name: 'stroked',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'stroked'
            },
            type: 'attribute'
          }, {
            name: 'strokeweight',
            attributeName: {
              localPart: 'strokeweight'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'target',
            attributeName: {
              localPart: 'target'
            },
            type: 'attribute'
          }, {
            name: 'title',
            attributeName: {
              localPart: 'title'
            },
            type: 'attribute'
          }, {
            name: 'wrapcoords',
            attributeName: {
              localPart: 'wrapcoords'
            },
            type: 'attribute'
          }, {
            name: 'allowincell',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'allowoverlap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'borderbottomcolor',
            type: 'attribute'
          }, {
            name: 'borderleftcolor',
            type: 'attribute'
          }, {
            name: 'borderrightcolor',
            type: 'attribute'
          }, {
            name: 'bordertopcolor',
            type: 'attribute'
          }, {
            name: 'bullet',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'button',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'bwmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwnormal',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwpure',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'clip',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'cliptowrap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'connectortype',
            typeInfo: 'org_docx4j_vml_officedrawing.STConnectorType',
            defaultValue: 'straight',
            type: 'attribute'
          }, {
            name: 'dgmlayout',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmlayoutmru',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmnodekind',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'doubleclicknotify',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'forcedash',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hr',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hralign',
            typeInfo: 'org_docx4j_vml_officedrawing.STHrAlign',
            defaultValue: 'left',
            type: 'attribute'
          }, {
            name: 'hrnoshade',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hrpct',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'hrstd',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'insetmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STInsetMode',
            defaultValue: 'custom',
            type: 'attribute'
          }, {
            name: 'ole',
            values: ['', 't', 'f', 'true', 'false'],
            type: 'attribute'
          }, {
            name: 'oleicon',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'oned',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'preferrelative',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'regroupid',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'spid',
            type: 'attribute'
          }, {
            name: 'spt',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'userdrawn',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'userhidden',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        localName: 'CTRect',
        typeName: 'CT_Rect',
        propertyInfos: [{
            name: 'pathOrFormulasOrHandles',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'path',
                typeInfo: '.CTPath'
              }, {
                elementName: 'formulas',
                typeInfo: '.CTFormulas'
              }, {
                elementName: 'handles',
                typeInfo: '.CTHandles'
              }, {
                elementName: 'fill',
                typeInfo: '.CTFill'
              }, {
                elementName: 'stroke',
                typeInfo: '.CTStroke'
              }, {
                elementName: 'shadow',
                typeInfo: '.CTShadow'
              }, {
                elementName: 'textbox',
                typeInfo: '.CTTextbox'
              }, {
                elementName: 'textpath',
                typeInfo: '.CTTextPath'
              }, {
                elementName: 'imagedata',
                typeInfo: '.CTImageData'
              }, {
                elementName: {
                  localPart: 'skew',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSkew'
              }, {
                elementName: {
                  localPart: 'extrusion',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTExtrusion'
              }, {
                elementName: {
                  localPart: 'callout',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTCallout'
              }, {
                elementName: {
                  localPart: 'lock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTLock'
              }, {
                elementName: {
                  localPart: 'clippath',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTClipPath'
              }, {
                elementName: {
                  localPart: 'signatureline',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSignatureLine'
              }, {
                elementName: {
                  localPart: 'wrap',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTWrap'
              }, {
                elementName: {
                  localPart: 'anchorlock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTAnchorLock'
              }, {
                elementName: {
                  localPart: 'bordertop',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderbottom',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderleft',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderright',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'ClientData',
                  namespaceURI: 'urn:schemas-microsoft-com:office:excel'
                },
                typeInfo: 'org_docx4j_vml_spreadsheetDrawing.CTClientData'
              }, {
                elementName: {
                  localPart: 'textdata',
                  namespaceURI: 'urn:schemas-microsoft-com:office:powerpoint'
                },
                typeInfo: 'org_docx4j_vml_presentationDrawing.CTRel'
              }],
            type: 'elementRefs'
          }, {
            name: 'alt',
            attributeName: {
              localPart: 'alt'
            },
            type: 'attribute'
          }, {
            name: 'chromakey',
            attributeName: {
              localPart: 'chromakey'
            },
            type: 'attribute'
          }, {
            name: 'clazz',
            attributeName: {
              localPart: 'class'
            },
            type: 'attribute'
          }, {
            name: 'coordorigin',
            attributeName: {
              localPart: 'coordorigin'
            },
            type: 'attribute'
          }, {
            name: 'coordsize',
            attributeName: {
              localPart: 'coordsize'
            },
            type: 'attribute'
          }, {
            name: 'fillcolor',
            attributeName: {
              localPart: 'fillcolor'
            },
            type: 'attribute'
          }, {
            name: 'filled',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'filled'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'insetpen',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'insetpen'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'print',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'print'
            },
            type: 'attribute'
          }, {
            name: 'strokecolor',
            attributeName: {
              localPart: 'strokecolor'
            },
            type: 'attribute'
          }, {
            name: 'stroked',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'stroked'
            },
            type: 'attribute'
          }, {
            name: 'strokeweight',
            attributeName: {
              localPart: 'strokeweight'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'target',
            attributeName: {
              localPart: 'target'
            },
            type: 'attribute'
          }, {
            name: 'title',
            attributeName: {
              localPart: 'title'
            },
            type: 'attribute'
          }, {
            name: 'wrapcoords',
            attributeName: {
              localPart: 'wrapcoords'
            },
            type: 'attribute'
          }, {
            name: 'allowincell',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'allowoverlap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'borderbottomcolor',
            type: 'attribute'
          }, {
            name: 'borderleftcolor',
            type: 'attribute'
          }, {
            name: 'borderrightcolor',
            type: 'attribute'
          }, {
            name: 'bordertopcolor',
            type: 'attribute'
          }, {
            name: 'bullet',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'button',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'bwmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwnormal',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwpure',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'clip',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'cliptowrap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'connectortype',
            typeInfo: 'org_docx4j_vml_officedrawing.STConnectorType',
            defaultValue: 'straight',
            type: 'attribute'
          }, {
            name: 'dgmlayout',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmlayoutmru',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmnodekind',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'doubleclicknotify',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'forcedash',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hr',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hralign',
            typeInfo: 'org_docx4j_vml_officedrawing.STHrAlign',
            defaultValue: 'left',
            type: 'attribute'
          }, {
            name: 'hrnoshade',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hrpct',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'hrstd',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'insetmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STInsetMode',
            defaultValue: 'custom',
            type: 'attribute'
          }, {
            name: 'ole',
            values: ['', 't', 'f', 'true', 'false'],
            type: 'attribute'
          }, {
            name: 'oleicon',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'oned',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'preferrelative',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'regroupid',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'spid',
            type: 'attribute'
          }, {
            name: 'spt',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'userdrawn',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'userhidden',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        localName: 'CTRoundRect',
        typeName: 'CT_RoundRect',
        propertyInfos: [{
            name: 'pathOrFormulasOrHandles',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'path',
                typeInfo: '.CTPath'
              }, {
                elementName: 'formulas',
                typeInfo: '.CTFormulas'
              }, {
                elementName: 'handles',
                typeInfo: '.CTHandles'
              }, {
                elementName: 'fill',
                typeInfo: '.CTFill'
              }, {
                elementName: 'stroke',
                typeInfo: '.CTStroke'
              }, {
                elementName: 'shadow',
                typeInfo: '.CTShadow'
              }, {
                elementName: 'textbox',
                typeInfo: '.CTTextbox'
              }, {
                elementName: 'textpath',
                typeInfo: '.CTTextPath'
              }, {
                elementName: 'imagedata',
                typeInfo: '.CTImageData'
              }, {
                elementName: {
                  localPart: 'skew',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSkew'
              }, {
                elementName: {
                  localPart: 'extrusion',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTExtrusion'
              }, {
                elementName: {
                  localPart: 'callout',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTCallout'
              }, {
                elementName: {
                  localPart: 'lock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTLock'
              }, {
                elementName: {
                  localPart: 'clippath',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTClipPath'
              }, {
                elementName: {
                  localPart: 'signatureline',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSignatureLine'
              }, {
                elementName: {
                  localPart: 'wrap',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTWrap'
              }, {
                elementName: {
                  localPart: 'anchorlock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTAnchorLock'
              }, {
                elementName: {
                  localPart: 'bordertop',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderbottom',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderleft',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderright',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'ClientData',
                  namespaceURI: 'urn:schemas-microsoft-com:office:excel'
                },
                typeInfo: 'org_docx4j_vml_spreadsheetDrawing.CTClientData'
              }, {
                elementName: {
                  localPart: 'textdata',
                  namespaceURI: 'urn:schemas-microsoft-com:office:powerpoint'
                },
                typeInfo: 'org_docx4j_vml_presentationDrawing.CTRel'
              }],
            type: 'elementRefs'
          }, {
            name: 'alt',
            attributeName: {
              localPart: 'alt'
            },
            type: 'attribute'
          }, {
            name: 'arcsize',
            attributeName: {
              localPart: 'arcsize'
            },
            type: 'attribute'
          }, {
            name: 'chromakey',
            attributeName: {
              localPart: 'chromakey'
            },
            type: 'attribute'
          }, {
            name: 'clazz',
            attributeName: {
              localPart: 'class'
            },
            type: 'attribute'
          }, {
            name: 'coordorigin',
            attributeName: {
              localPart: 'coordorigin'
            },
            type: 'attribute'
          }, {
            name: 'coordsize',
            attributeName: {
              localPart: 'coordsize'
            },
            type: 'attribute'
          }, {
            name: 'fillcolor',
            attributeName: {
              localPart: 'fillcolor'
            },
            type: 'attribute'
          }, {
            name: 'filled',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'filled'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'insetpen',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'insetpen'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'print',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'print'
            },
            type: 'attribute'
          }, {
            name: 'strokecolor',
            attributeName: {
              localPart: 'strokecolor'
            },
            type: 'attribute'
          }, {
            name: 'stroked',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'stroked'
            },
            type: 'attribute'
          }, {
            name: 'strokeweight',
            attributeName: {
              localPart: 'strokeweight'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'target',
            attributeName: {
              localPart: 'target'
            },
            type: 'attribute'
          }, {
            name: 'title',
            attributeName: {
              localPart: 'title'
            },
            type: 'attribute'
          }, {
            name: 'wrapcoords',
            attributeName: {
              localPart: 'wrapcoords'
            },
            type: 'attribute'
          }, {
            name: 'allowincell',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'allowoverlap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'borderbottomcolor',
            type: 'attribute'
          }, {
            name: 'borderleftcolor',
            type: 'attribute'
          }, {
            name: 'borderrightcolor',
            type: 'attribute'
          }, {
            name: 'bordertopcolor',
            type: 'attribute'
          }, {
            name: 'bullet',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'button',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'bwmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwnormal',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwpure',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'clip',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'cliptowrap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'connectortype',
            typeInfo: 'org_docx4j_vml_officedrawing.STConnectorType',
            defaultValue: 'straight',
            type: 'attribute'
          }, {
            name: 'dgmlayout',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmlayoutmru',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmnodekind',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'doubleclicknotify',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'forcedash',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hr',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hralign',
            typeInfo: 'org_docx4j_vml_officedrawing.STHrAlign',
            defaultValue: 'left',
            type: 'attribute'
          }, {
            name: 'hrnoshade',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hrpct',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'hrstd',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'insetmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STInsetMode',
            defaultValue: 'custom',
            type: 'attribute'
          }, {
            name: 'ole',
            values: ['', 't', 'f', 'true', 'false'],
            type: 'attribute'
          }, {
            name: 'oleicon',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'oned',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'preferrelative',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'regroupid',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'spid',
            type: 'attribute'
          }, {
            name: 'spt',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'userdrawn',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'userhidden',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        localName: 'CTShadow',
        typeName: 'CT_Shadow',
        propertyInfos: [{
            name: 'color',
            attributeName: {
              localPart: 'color'
            },
            type: 'attribute'
          }, {
            name: 'color2',
            attributeName: {
              localPart: 'color2'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'matrix',
            attributeName: {
              localPart: 'matrix'
            },
            type: 'attribute'
          }, {
            name: 'obscured',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'obscured'
            },
            type: 'attribute'
          }, {
            name: 'offset',
            attributeName: {
              localPart: 'offset'
            },
            type: 'attribute'
          }, {
            name: 'offset2',
            attributeName: {
              localPart: 'offset2'
            },
            type: 'attribute'
          }, {
            name: 'on',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'on'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'origin',
            attributeName: {
              localPart: 'origin'
            },
            type: 'attribute'
          }, {
            name: 'type',
            typeInfo: '.STShadowType',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTShape',
        typeName: 'CT_Shape',
        propertyInfos: [{
            name: 'pathOrFormulasOrHandles',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'path',
                typeInfo: '.CTPath'
              }, {
                elementName: 'formulas',
                typeInfo: '.CTFormulas'
              }, {
                elementName: 'handles',
                typeInfo: '.CTHandles'
              }, {
                elementName: 'fill',
                typeInfo: '.CTFill'
              }, {
                elementName: 'stroke',
                typeInfo: '.CTStroke'
              }, {
                elementName: 'shadow',
                typeInfo: '.CTShadow'
              }, {
                elementName: 'textbox',
                typeInfo: '.CTTextbox'
              }, {
                elementName: 'textpath',
                typeInfo: '.CTTextPath'
              }, {
                elementName: 'imagedata',
                typeInfo: '.CTImageData'
              }, {
                elementName: {
                  localPart: 'skew',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSkew'
              }, {
                elementName: {
                  localPart: 'extrusion',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTExtrusion'
              }, {
                elementName: {
                  localPart: 'callout',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTCallout'
              }, {
                elementName: {
                  localPart: 'lock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTLock'
              }, {
                elementName: {
                  localPart: 'clippath',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTClipPath'
              }, {
                elementName: {
                  localPart: 'signatureline',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSignatureLine'
              }, {
                elementName: {
                  localPart: 'wrap',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTWrap'
              }, {
                elementName: {
                  localPart: 'anchorlock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTAnchorLock'
              }, {
                elementName: {
                  localPart: 'bordertop',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderbottom',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderleft',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderright',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'ClientData',
                  namespaceURI: 'urn:schemas-microsoft-com:office:excel'
                },
                typeInfo: 'org_docx4j_vml_spreadsheetDrawing.CTClientData'
              }, {
                elementName: {
                  localPart: 'textdata',
                  namespaceURI: 'urn:schemas-microsoft-com:office:powerpoint'
                },
                typeInfo: 'org_docx4j_vml_presentationDrawing.CTRel'
              }, {
                elementName: {
                  localPart: 'ink',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTInk'
              }, {
                elementName: {
                  localPart: 'iscomment',
                  namespaceURI: 'urn:schemas-microsoft-com:office:powerpoint'
                },
                typeInfo: 'org_docx4j_vml_presentationDrawing.CTEmpty'
              }],
            type: 'elementRefs'
          }, {
            name: 'adj',
            attributeName: {
              localPart: 'adj'
            },
            type: 'attribute'
          }, {
            name: 'alt',
            attributeName: {
              localPart: 'alt'
            },
            type: 'attribute'
          }, {
            name: 'chromakey',
            attributeName: {
              localPart: 'chromakey'
            },
            type: 'attribute'
          }, {
            name: 'clazz',
            attributeName: {
              localPart: 'class'
            },
            type: 'attribute'
          }, {
            name: 'coordorigin',
            attributeName: {
              localPart: 'coordorigin'
            },
            type: 'attribute'
          }, {
            name: 'coordsize',
            attributeName: {
              localPart: 'coordsize'
            },
            type: 'attribute'
          }, {
            name: 'equationxml',
            attributeName: {
              localPart: 'equationxml'
            },
            type: 'attribute'
          }, {
            name: 'fillcolor',
            attributeName: {
              localPart: 'fillcolor'
            },
            type: 'attribute'
          }, {
            name: 'filled',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'filled'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'insetpen',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'insetpen'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'path',
            attributeName: {
              localPart: 'path'
            },
            type: 'attribute'
          }, {
            name: 'print',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'print'
            },
            type: 'attribute'
          }, {
            name: 'strokecolor',
            attributeName: {
              localPart: 'strokecolor'
            },
            type: 'attribute'
          }, {
            name: 'stroked',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'stroked'
            },
            type: 'attribute'
          }, {
            name: 'strokeweight',
            attributeName: {
              localPart: 'strokeweight'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'target',
            attributeName: {
              localPart: 'target'
            },
            type: 'attribute'
          }, {
            name: 'title',
            attributeName: {
              localPart: 'title'
            },
            type: 'attribute'
          }, {
            name: 'type',
            attributeName: {
              localPart: 'type'
            },
            type: 'attribute'
          }, {
            name: 'wrapcoords',
            attributeName: {
              localPart: 'wrapcoords'
            },
            type: 'attribute'
          }, {
            name: 'allowincell',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'allowoverlap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'borderbottomcolor',
            type: 'attribute'
          }, {
            name: 'borderleftcolor',
            type: 'attribute'
          }, {
            name: 'borderrightcolor',
            type: 'attribute'
          }, {
            name: 'bordertopcolor',
            type: 'attribute'
          }, {
            name: 'bullet',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'button',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'bwmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwnormal',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwpure',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'clip',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'cliptowrap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'connectortype',
            typeInfo: 'org_docx4j_vml_officedrawing.STConnectorType',
            defaultValue: 'straight',
            type: 'attribute'
          }, {
            name: 'dgmlayout',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmlayoutmru',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmnodekind',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'doubleclicknotify',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'forcedash',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'gfxdata',
            typeInfo: 'Base64Binary',
            type: 'attribute'
          }, {
            name: 'hr',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hralign',
            typeInfo: 'org_docx4j_vml_officedrawing.STHrAlign',
            defaultValue: 'left',
            type: 'attribute'
          }, {
            name: 'hrnoshade',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hrpct',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'hrstd',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'insetmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STInsetMode',
            defaultValue: 'custom',
            type: 'attribute'
          }, {
            name: 'ole',
            values: ['', 't', 'f', 'true', 'false'],
            type: 'attribute'
          }, {
            name: 'oleicon',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'oned',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'preferrelative',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'regroupid',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'spid',
            type: 'attribute'
          }, {
            name: 'spt',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'userdrawn',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'userhidden',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        localName: 'CTShapetype',
        typeName: 'CT_Shapetype',
        propertyInfos: [{
            name: 'egShapeElements',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'path',
                typeInfo: '.CTPath'
              }, {
                elementName: 'formulas',
                typeInfo: '.CTFormulas'
              }, {
                elementName: 'handles',
                typeInfo: '.CTHandles'
              }, {
                elementName: 'fill',
                typeInfo: '.CTFill'
              }, {
                elementName: 'stroke',
                typeInfo: '.CTStroke'
              }, {
                elementName: 'shadow',
                typeInfo: '.CTShadow'
              }, {
                elementName: 'textbox',
                typeInfo: '.CTTextbox'
              }, {
                elementName: 'textpath',
                typeInfo: '.CTTextPath'
              }, {
                elementName: 'imagedata',
                typeInfo: '.CTImageData'
              }, {
                elementName: {
                  localPart: 'skew',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSkew'
              }, {
                elementName: {
                  localPart: 'extrusion',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTExtrusion'
              }, {
                elementName: {
                  localPart: 'callout',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTCallout'
              }, {
                elementName: {
                  localPart: 'lock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTLock'
              }, {
                elementName: {
                  localPart: 'clippath',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTClipPath'
              }, {
                elementName: {
                  localPart: 'signatureline',
                  namespaceURI: 'urn:schemas-microsoft-com:office:office'
                },
                typeInfo: 'org_docx4j_vml_officedrawing.CTSignatureLine'
              }, {
                elementName: {
                  localPart: 'wrap',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTWrap'
              }, {
                elementName: {
                  localPart: 'anchorlock',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTAnchorLock'
              }, {
                elementName: {
                  localPart: 'bordertop',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderbottom',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderleft',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'borderright',
                  namespaceURI: 'urn:schemas-microsoft-com:office:word'
                },
                typeInfo: 'org_docx4j_vml_wordprocessingDrawing.CTBorder'
              }, {
                elementName: {
                  localPart: 'ClientData',
                  namespaceURI: 'urn:schemas-microsoft-com:office:excel'
                },
                typeInfo: 'org_docx4j_vml_spreadsheetDrawing.CTClientData'
              }, {
                elementName: {
                  localPart: 'textdata',
                  namespaceURI: 'urn:schemas-microsoft-com:office:powerpoint'
                },
                typeInfo: 'org_docx4j_vml_presentationDrawing.CTRel'
              }],
            type: 'elementRefs'
          }, {
            name: 'complex',
            elementName: {
              localPart: 'complex',
              namespaceURI: 'urn:schemas-microsoft-com:office:office'
            },
            typeInfo: 'org_docx4j_vml_officedrawing.CTComplex'
          }, {
            name: 'adj',
            attributeName: {
              localPart: 'adj'
            },
            type: 'attribute'
          }, {
            name: 'alt',
            attributeName: {
              localPart: 'alt'
            },
            type: 'attribute'
          }, {
            name: 'chromakey',
            attributeName: {
              localPart: 'chromakey'
            },
            type: 'attribute'
          }, {
            name: 'clazz',
            attributeName: {
              localPart: 'class'
            },
            type: 'attribute'
          }, {
            name: 'coordorigin',
            attributeName: {
              localPart: 'coordorigin'
            },
            type: 'attribute'
          }, {
            name: 'coordsize',
            attributeName: {
              localPart: 'coordsize'
            },
            type: 'attribute'
          }, {
            name: 'fillcolor',
            attributeName: {
              localPart: 'fillcolor'
            },
            type: 'attribute'
          }, {
            name: 'filled',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'filled'
            },
            type: 'attribute'
          }, {
            name: 'href',
            attributeName: {
              localPart: 'href'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'insetpen',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'insetpen'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'path',
            attributeName: {
              localPart: 'path'
            },
            type: 'attribute'
          }, {
            name: 'print',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'print'
            },
            type: 'attribute'
          }, {
            name: 'strokecolor',
            attributeName: {
              localPart: 'strokecolor'
            },
            type: 'attribute'
          }, {
            name: 'stroked',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'stroked'
            },
            type: 'attribute'
          }, {
            name: 'strokeweight',
            attributeName: {
              localPart: 'strokeweight'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'target',
            attributeName: {
              localPart: 'target'
            },
            type: 'attribute'
          }, {
            name: 'title',
            attributeName: {
              localPart: 'title'
            },
            type: 'attribute'
          }, {
            name: 'wrapcoords',
            attributeName: {
              localPart: 'wrapcoords'
            },
            type: 'attribute'
          }, {
            name: 'allowincell',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'allowoverlap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'borderbottomcolor',
            type: 'attribute'
          }, {
            name: 'borderleftcolor',
            type: 'attribute'
          }, {
            name: 'borderrightcolor',
            type: 'attribute'
          }, {
            name: 'bordertopcolor',
            type: 'attribute'
          }, {
            name: 'bullet',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'button',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'bwmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwnormal',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'bwpure',
            typeInfo: 'org_docx4j_vml_officedrawing.STBWMode',
            type: 'attribute'
          }, {
            name: 'clip',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'cliptowrap',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'connectortype',
            typeInfo: 'org_docx4j_vml_officedrawing.STConnectorType',
            defaultValue: 'straight',
            type: 'attribute'
          }, {
            name: 'dgmlayout',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmlayoutmru',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'dgmnodekind',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'doubleclicknotify',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'forcedash',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hr',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hralign',
            typeInfo: 'org_docx4j_vml_officedrawing.STHrAlign',
            defaultValue: 'left',
            type: 'attribute'
          }, {
            name: 'hrnoshade',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'hrpct',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'hrstd',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'insetmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STInsetMode',
            defaultValue: 'custom',
            type: 'attribute'
          }, {
            name: 'master',
            type: 'attribute'
          }, {
            name: 'ole',
            values: ['', 't', 'f', 'true', 'false'],
            type: 'attribute'
          }, {
            name: 'oleicon',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'oned',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'preferrelative',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'regroupid',
            typeInfo: 'Integer',
            type: 'attribute'
          }, {
            name: 'spid',
            type: 'attribute'
          }, {
            name: 'spt',
            typeInfo: 'Float',
            type: 'attribute'
          }, {
            name: 'userdrawn',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'userhidden',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        localName: 'CTStroke',
        typeName: 'CT_Stroke',
        propertyInfos: [{
            name: 'left',
            elementName: {
              localPart: 'left',
              namespaceURI: 'urn:schemas-microsoft-com:office:office'
            },
            typeInfo: 'org_docx4j_vml_officedrawing.CTStrokeChild'
          }, {
            name: 'top',
            elementName: {
              localPart: 'top',
              namespaceURI: 'urn:schemas-microsoft-com:office:office'
            },
            typeInfo: 'org_docx4j_vml_officedrawing.CTStrokeChild'
          }, {
            name: 'right',
            elementName: {
              localPart: 'right',
              namespaceURI: 'urn:schemas-microsoft-com:office:office'
            },
            typeInfo: 'org_docx4j_vml_officedrawing.CTStrokeChild'
          }, {
            name: 'bottom',
            elementName: {
              localPart: 'bottom',
              namespaceURI: 'urn:schemas-microsoft-com:office:office'
            },
            typeInfo: 'org_docx4j_vml_officedrawing.CTStrokeChild'
          }, {
            name: 'column',
            elementName: {
              localPart: 'column',
              namespaceURI: 'urn:schemas-microsoft-com:office:office'
            },
            typeInfo: 'org_docx4j_vml_officedrawing.CTStrokeChild'
          }, {
            name: 'color',
            attributeName: {
              localPart: 'color'
            },
            type: 'attribute'
          }, {
            name: 'color2',
            attributeName: {
              localPart: 'color2'
            },
            type: 'attribute'
          }, {
            name: 'dashstyle',
            attributeName: {
              localPart: 'dashstyle'
            },
            type: 'attribute'
          }, {
            name: 'endarrow',
            typeInfo: '.STStrokeArrowType',
            attributeName: {
              localPart: 'endarrow'
            },
            type: 'attribute'
          }, {
            name: 'endarrowlength',
            typeInfo: '.STStrokeArrowLength',
            attributeName: {
              localPart: 'endarrowlength'
            },
            type: 'attribute'
          }, {
            name: 'endarrowwidth',
            typeInfo: '.STStrokeArrowWidth',
            attributeName: {
              localPart: 'endarrowwidth'
            },
            type: 'attribute'
          }, {
            name: 'endcap',
            typeInfo: '.STStrokeEndCap',
            attributeName: {
              localPart: 'endcap'
            },
            type: 'attribute'
          }, {
            name: 'filltype',
            typeInfo: '.STFillType',
            attributeName: {
              localPart: 'filltype'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'imagealignshape',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'imagealignshape'
            },
            type: 'attribute'
          }, {
            name: 'imageaspect',
            typeInfo: '.STImageAspect',
            attributeName: {
              localPart: 'imageaspect'
            },
            type: 'attribute'
          }, {
            name: 'imagesize',
            attributeName: {
              localPart: 'imagesize'
            },
            type: 'attribute'
          }, {
            name: 'insetpen',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'insetpen'
            },
            type: 'attribute'
          }, {
            name: 'joinstyle',
            typeInfo: '.STStrokeJoinStyle',
            attributeName: {
              localPart: 'joinstyle'
            },
            type: 'attribute'
          }, {
            name: 'linestyle',
            typeInfo: '.STStrokeLineStyle',
            attributeName: {
              localPart: 'linestyle'
            },
            type: 'attribute'
          }, {
            name: 'miterlimit',
            attributeName: {
              localPart: 'miterlimit'
            },
            type: 'attribute'
          }, {
            name: 'on',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'on'
            },
            type: 'attribute'
          }, {
            name: 'opacity',
            attributeName: {
              localPart: 'opacity'
            },
            type: 'attribute'
          }, {
            name: 'src',
            attributeName: {
              localPart: 'src'
            },
            type: 'attribute'
          }, {
            name: 'startarrow',
            typeInfo: '.STStrokeArrowType',
            attributeName: {
              localPart: 'startarrow'
            },
            type: 'attribute'
          }, {
            name: 'startarrowlength',
            typeInfo: '.STStrokeArrowLength',
            attributeName: {
              localPart: 'startarrowlength'
            },
            type: 'attribute'
          }, {
            name: 'startarrowwidth',
            typeInfo: '.STStrokeArrowWidth',
            attributeName: {
              localPart: 'startarrowwidth'
            },
            type: 'attribute'
          }, {
            name: 'weight',
            attributeName: {
              localPart: 'weight'
            },
            type: 'attribute'
          }, {
            name: 'id',
            attributeName: {
              localPart: 'id',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships'
            },
            type: 'attribute'
          }, {
            name: 'althref',
            type: 'attribute'
          }, {
            name: 'forcedash',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }, {
            name: 'href',
            type: 'attribute'
          }, {
            name: 'relid',
            type: 'attribute'
          }, {
            name: 'title',
            type: 'attribute'
          }]
      }, {
        localName: 'CTTextPath',
        typeName: 'CT_TextPath',
        propertyInfos: [{
            name: 'fitpath',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'fitpath'
            },
            type: 'attribute'
          }, {
            name: 'fitshape',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'fitshape'
            },
            type: 'attribute'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'on',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'on'
            },
            type: 'attribute'
          }, {
            name: 'string',
            attributeName: {
              localPart: 'string'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'trim',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'trim'
            },
            type: 'attribute'
          }, {
            name: 'xscale',
            typeInfo: '.STTrueFalse',
            attributeName: {
              localPart: 'xscale'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTTextbox',
        typeName: 'CT_Textbox',
        propertyInfos: [{
            name: 'txbxContent',
            elementName: {
              localPart: 'txbxContent',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
            },
            typeInfo: 'org_docx4j_wml.CTTxbxContent'
          }, {
            name: 'any',
            required: true,
            allowTypedObject: false,
            mixed: false,
            type: 'anyElement'
          }, {
            name: 'vmlId',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }, {
            name: 'inset',
            attributeName: {
              localPart: 'inset'
            },
            type: 'attribute'
          }, {
            name: 'style',
            attributeName: {
              localPart: 'style'
            },
            type: 'attribute'
          }, {
            name: 'insetmode',
            typeInfo: 'org_docx4j_vml_officedrawing.STInsetMode',
            defaultValue: 'custom',
            type: 'attribute'
          }, {
            name: 'singleclick',
            typeInfo: 'org_docx4j_vml_officedrawing.STTrueFalse',
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STEditAs',
        values: ['canvas', 'orgchart', 'radial', 'cycle', 'stacked', 'venn', 'bullseye']
      }, {
        type: 'enumInfo',
        localName: 'STExt',
        values: ['view', 'edit', 'backwardCompatible']
      }, {
        type: 'enumInfo',
        localName: 'STFillMethod',
        values: ['none', 'linear', 'sigma', 'any', 'linear sigma']
      }, {
        type: 'enumInfo',
        localName: 'STFillType',
        values: ['solid', 'gradient', 'gradientRadial', 'tile', 'pattern', 'frame']
      }, {
        type: 'enumInfo',
        localName: 'STImageAspect',
        values: ['ignore', 'atMost', 'atLeast']
      }, {
        type: 'enumInfo',
        localName: 'STShadowType',
        values: ['single', 'double', 'emboss', 'perspective']
      }, {
        type: 'enumInfo',
        localName: 'STStrokeArrowLength',
        values: ['short', 'medium', 'long']
      }, {
        type: 'enumInfo',
        localName: 'STStrokeArrowType',
        values: ['none', 'block', 'classic', 'oval', 'diamond', 'open']
      }, {
        type: 'enumInfo',
        localName: 'STStrokeArrowWidth',
        values: ['narrow', 'medium', 'wide']
      }, {
        type: 'enumInfo',
        localName: 'STStrokeEndCap',
        values: ['flat', 'square', 'round']
      }, {
        type: 'enumInfo',
        localName: 'STStrokeJoinStyle',
        values: ['round', 'bevel', 'miter']
      }, {
        type: 'enumInfo',
        localName: 'STStrokeLineStyle',
        values: ['single', 'thinThin', 'thinThick', 'thickThin', 'thickBetweenThin']
      }, {
        type: 'enumInfo',
        localName: 'STTrueFalse',
        values: ['t', 'f', 'true', 'false']
      }],
    elementInfos: [{
        typeInfo: '.CTArc',
        elementName: 'arc'
      }, {
        typeInfo: '.CTBackground',
        elementName: 'background'
      }, {
        typeInfo: '.CTCurve',
        elementName: 'curve'
      }, {
        typeInfo: '.CTFill',
        elementName: 'fill'
      }, {
        typeInfo: '.CTFormulas',
        elementName: 'formulas'
      }, {
        typeInfo: '.CTGroup',
        elementName: 'group'
      }, {
        typeInfo: '.CTHandles',
        elementName: 'handles'
      }, {
        typeInfo: '.CTImage',
        elementName: 'image'
      }, {
        typeInfo: '.CTImageData',
        elementName: 'imagedata'
      }, {
        typeInfo: '.CTLine',
        elementName: 'line'
      }, {
        typeInfo: '.CTOval',
        elementName: 'oval'
      }, {
        typeInfo: '.CTPath',
        elementName: 'path'
      }, {
        typeInfo: '.CTPolyLine',
        elementName: 'polyline'
      }, {
        typeInfo: '.CTRect',
        elementName: 'rect'
      }, {
        typeInfo: '.CTRoundRect',
        elementName: 'roundrect'
      }, {
        typeInfo: '.CTShadow',
        elementName: 'shadow'
      }, {
        typeInfo: '.CTShape',
        elementName: 'shape'
      }, {
        typeInfo: '.CTShapetype',
        elementName: 'shapetype'
      }, {
        typeInfo: '.CTStroke',
        elementName: 'stroke'
      }, {
        typeInfo: '.CTTextbox',
        elementName: 'textbox'
      }, {
        typeInfo: '.CTTextPath',
        elementName: 'textpath'
      }]
  };
  return {
    org_docx4j_vml: org_docx4j_vml
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_vml_Module_Factory);
}
else {
  var org_docx4j_vml_Module = org_docx4j_vml_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_vml = org_docx4j_vml_Module.org_docx4j_vml;
  }
  else {
    var org_docx4j_vml = org_docx4j_vml_Module.org_docx4j_vml;
  }
}