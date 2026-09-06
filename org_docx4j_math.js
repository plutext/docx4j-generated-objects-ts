var org_docx4j_math_Module_Factory = function () {
  var org_docx4j_math = {
    name: 'org_docx4j_math',
    defaultElementNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/math',
    defaultAttributeNamespaceURI: 'http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/math',
    dependencies: ['org_docx4j_sharedtypes', 'org_docx4j_wml'],
    typeInfos: [{
        localName: 'CTAcc',
        typeName: 'CT_Acc',
        propertyInfos: [{
            name: 'accPr',
            typeInfo: '.CTAccPr'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTAccPr',
        typeName: 'CT_AccPr',
        propertyInfos: [{
            name: 'chr',
            typeInfo: '.CTChar'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTBar',
        typeName: 'CT_Bar',
        propertyInfos: [{
            name: 'barPr',
            typeInfo: '.CTBarPr'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTBarPr',
        typeName: 'CT_BarPr',
        propertyInfos: [{
            name: 'pos',
            typeInfo: '.CTTopBot'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTBorderBox',
        typeName: 'CT_BorderBox',
        propertyInfos: [{
            name: 'borderBoxPr',
            typeInfo: '.CTBorderBoxPr'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTBorderBoxPr',
        typeName: 'CT_BorderBoxPr',
        propertyInfos: [{
            name: 'hideTop',
            typeInfo: '.CTOnOff'
          }, {
            name: 'hideBot',
            typeInfo: '.CTOnOff'
          }, {
            name: 'hideLeft',
            typeInfo: '.CTOnOff'
          }, {
            name: 'hideRight',
            typeInfo: '.CTOnOff'
          }, {
            name: 'strikeH',
            typeInfo: '.CTOnOff'
          }, {
            name: 'strikeV',
            typeInfo: '.CTOnOff'
          }, {
            name: 'strikeBLTR',
            typeInfo: '.CTOnOff'
          }, {
            name: 'strikeTLBR',
            typeInfo: '.CTOnOff'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTBox',
        typeName: 'CT_Box',
        propertyInfos: [{
            name: 'boxPr',
            typeInfo: '.CTBoxPr'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTBoxPr',
        typeName: 'CT_BoxPr',
        propertyInfos: [{
            name: 'opEmu',
            typeInfo: '.CTOnOff'
          }, {
            name: 'noBreak',
            typeInfo: '.CTOnOff'
          }, {
            name: 'diff',
            typeInfo: '.CTOnOff'
          }, {
            name: 'brk',
            typeInfo: '.CTManualBreak'
          }, {
            name: 'aln',
            typeInfo: '.CTOnOff'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTBreakBin',
        typeName: 'CT_BreakBin',
        propertyInfos: [{
            name: 'val',
            typeInfo: '.STBreakBin',
            type: 'attribute'
          }]
      }, {
        localName: 'CTBreakBinSub',
        typeName: 'CT_BreakBinSub',
        propertyInfos: [{
            name: 'val',
            values: ['--', '-+', '+-'],
            type: 'attribute'
          }]
      }, {
        localName: 'CTChar',
        typeName: 'CT_Char',
        propertyInfos: [{
            name: 'val',
            required: true,
            type: 'attribute'
          }]
      }, {
        localName: 'CTCtrlPr',
        typeName: 'CT_CtrlPr',
        propertyInfos: [{
            name: 'rPr',
            elementName: {
              localPart: 'rPr',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
            },
            typeInfo: 'org_docx4j_wml.RPr'
          }, {
            name: 'ins',
            required: true,
            elementName: {
              localPart: 'ins',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
            },
            typeInfo: 'org_docx4j_wml.CTRPrChange'
          }, {
            name: 'del',
            required: true,
            elementName: {
              localPart: 'del',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
            },
            typeInfo: 'org_docx4j_wml.CTRPrChange'
          }]
      }, {
        localName: 'CTD',
        typeName: 'CT_D',
        propertyInfos: [{
            name: 'dPr',
            typeInfo: '.CTDPr'
          }, {
            name: 'e',
            required: true,
            collection: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTDPr',
        typeName: 'CT_DPr',
        propertyInfos: [{
            name: 'begChr',
            typeInfo: '.CTChar'
          }, {
            name: 'sepChr',
            typeInfo: '.CTChar'
          }, {
            name: 'endChr',
            typeInfo: '.CTChar'
          }, {
            name: 'grow',
            typeInfo: '.CTOnOff'
          }, {
            name: 'shp',
            typeInfo: '.CTShp'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTEqArr',
        typeName: 'CT_EqArr',
        propertyInfos: [{
            name: 'eqArrPr',
            typeInfo: '.CTEqArrPr'
          }, {
            name: 'e',
            required: true,
            collection: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTEqArrPr',
        typeName: 'CT_EqArrPr',
        propertyInfos: [{
            name: 'baseJc',
            typeInfo: '.CTYAlign'
          }, {
            name: 'maxDist',
            typeInfo: '.CTOnOff'
          }, {
            name: 'objDist',
            typeInfo: '.CTOnOff'
          }, {
            name: 'rSpRule',
            typeInfo: '.CTSpacingRule'
          }, {
            name: 'rSp',
            typeInfo: '.CTUnSignedInteger'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTF',
        typeName: 'CT_F',
        propertyInfos: [{
            name: 'fPr',
            typeInfo: '.CTFPr'
          }, {
            name: 'num',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'den',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTFPr',
        typeName: 'CT_FPr',
        propertyInfos: [{
            name: 'type',
            typeInfo: '.CTFType'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTFType',
        typeName: 'CT_FType',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: '.STFType',
            type: 'attribute'
          }]
      }, {
        localName: 'CTFunc',
        typeName: 'CT_Func',
        propertyInfos: [{
            name: 'funcPr',
            typeInfo: '.CTFuncPr'
          }, {
            name: 'fName',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTFuncPr',
        typeName: 'CT_FuncPr',
        propertyInfos: [{
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTGroupChr',
        typeName: 'CT_GroupChr',
        propertyInfos: [{
            name: 'groupChrPr',
            typeInfo: '.CTGroupChrPr'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTGroupChrPr',
        typeName: 'CT_GroupChrPr',
        propertyInfos: [{
            name: 'chr',
            typeInfo: '.CTChar'
          }, {
            name: 'pos',
            typeInfo: '.CTTopBot'
          }, {
            name: 'vertJc',
            typeInfo: '.CTTopBot'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTInteger2',
        typeName: 'CT_Integer2',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'Integer',
            type: 'attribute'
          }]
      }, {
        localName: 'CTInteger255',
        typeName: 'CT_Integer255',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'Integer',
            type: 'attribute'
          }]
      }, {
        localName: 'CTLimLoc',
        typeName: 'CT_LimLoc',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: '.STLimLoc',
            type: 'attribute'
          }]
      }, {
        localName: 'CTLimLow',
        typeName: 'CT_LimLow',
        propertyInfos: [{
            name: 'limLowPr',
            typeInfo: '.CTLimLowPr'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'lim',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTLimLowPr',
        typeName: 'CT_LimLowPr',
        propertyInfos: [{
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTLimUpp',
        typeName: 'CT_LimUpp',
        propertyInfos: [{
            name: 'limUppPr',
            typeInfo: '.CTLimUppPr'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'lim',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTLimUppPr',
        typeName: 'CT_LimUppPr',
        propertyInfos: [{
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTM',
        typeName: 'CT_M',
        propertyInfos: [{
            name: 'mPr',
            typeInfo: '.CTMPr'
          }, {
            name: 'mr',
            required: true,
            collection: true,
            typeInfo: '.CTMR'
          }]
      }, {
        localName: 'CTMC',
        typeName: 'CT_MC',
        propertyInfos: [{
            name: 'mcPr',
            typeInfo: '.CTMCPr'
          }]
      }, {
        localName: 'CTMCPr',
        typeName: 'CT_MCPr',
        propertyInfos: [{
            name: 'count',
            typeInfo: '.CTInteger255'
          }, {
            name: 'mcJc',
            typeInfo: '.CTXAlign'
          }]
      }, {
        localName: 'CTMCS',
        typeName: 'CT_MCS',
        propertyInfos: [{
            name: 'mc',
            required: true,
            collection: true,
            typeInfo: '.CTMC'
          }]
      }, {
        localName: 'CTMPr',
        typeName: 'CT_MPr',
        propertyInfos: [{
            name: 'baseJc',
            typeInfo: '.CTYAlign'
          }, {
            name: 'plcHide',
            typeInfo: '.CTOnOff'
          }, {
            name: 'rSpRule',
            typeInfo: '.CTSpacingRule'
          }, {
            name: 'cGpRule',
            typeInfo: '.CTSpacingRule'
          }, {
            name: 'rSp',
            typeInfo: '.CTUnSignedInteger'
          }, {
            name: 'cSp',
            typeInfo: '.CTUnSignedInteger'
          }, {
            name: 'cGp',
            typeInfo: '.CTUnSignedInteger'
          }, {
            name: 'mcs',
            typeInfo: '.CTMCS'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTMR',
        typeName: 'CT_MR',
        propertyInfos: [{
            name: 'e',
            required: true,
            collection: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTManualBreak',
        typeName: 'CT_ManualBreak',
        propertyInfos: [{
            name: 'alnAt',
            typeInfo: 'Integer',
            type: 'attribute'
          }]
      }, {
        localName: 'CTMathPr',
        typeName: 'CT_MathPr',
        propertyInfos: [{
            name: 'mathFont',
            typeInfo: '.CTString'
          }, {
            name: 'brkBin',
            typeInfo: '.CTBreakBin'
          }, {
            name: 'brkBinSub',
            typeInfo: '.CTBreakBinSub'
          }, {
            name: 'smallFrac',
            typeInfo: '.CTOnOff'
          }, {
            name: 'dispDef',
            typeInfo: '.CTOnOff'
          }, {
            name: 'lMargin',
            typeInfo: '.CTTwipsMeasure'
          }, {
            name: 'rMargin',
            typeInfo: '.CTTwipsMeasure'
          }, {
            name: 'defJc',
            typeInfo: '.CTOMathJc'
          }, {
            name: 'preSp',
            typeInfo: '.CTTwipsMeasure'
          }, {
            name: 'postSp',
            typeInfo: '.CTTwipsMeasure'
          }, {
            name: 'interSp',
            typeInfo: '.CTTwipsMeasure'
          }, {
            name: 'intraSp',
            typeInfo: '.CTTwipsMeasure'
          }, {
            name: 'wrapIndent',
            required: true,
            typeInfo: '.CTTwipsMeasure'
          }, {
            name: 'wrapRight',
            required: true,
            typeInfo: '.CTOnOff'
          }, {
            name: 'intLim',
            typeInfo: '.CTLimLoc'
          }, {
            name: 'naryLim',
            typeInfo: '.CTLimLoc'
          }]
      }, {
        localName: 'CTNary',
        typeName: 'CT_Nary',
        propertyInfos: [{
            name: 'naryPr',
            typeInfo: '.CTNaryPr'
          }, {
            name: 'sub',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'sup',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTNaryPr',
        typeName: 'CT_NaryPr',
        propertyInfos: [{
            name: 'chr',
            typeInfo: '.CTChar'
          }, {
            name: 'limLoc',
            typeInfo: '.CTLimLoc'
          }, {
            name: 'grow',
            typeInfo: '.CTOnOff'
          }, {
            name: 'subHide',
            typeInfo: '.CTOnOff'
          }, {
            name: 'supHide',
            typeInfo: '.CTOnOff'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTOMath',
        typeName: 'CT_OMath',
        propertyInfos: [{
            name: 'egoMathElements',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'acc',
                typeInfo: '.CTAcc'
              }, {
                elementName: 'bar',
                typeInfo: '.CTBar'
              }, {
                elementName: 'box',
                typeInfo: '.CTBox'
              }, {
                elementName: 'borderBox',
                typeInfo: '.CTBorderBox'
              }, {
                elementName: 'd',
                typeInfo: '.CTD'
              }, {
                elementName: 'eqArr',
                typeInfo: '.CTEqArr'
              }, {
                elementName: 'f',
                typeInfo: '.CTF'
              }, {
                elementName: 'func',
                typeInfo: '.CTFunc'
              }, {
                elementName: 'groupChr',
                typeInfo: '.CTGroupChr'
              }, {
                elementName: 'limLow',
                typeInfo: '.CTLimLow'
              }, {
                elementName: 'limUpp',
                typeInfo: '.CTLimUpp'
              }, {
                elementName: 'm',
                typeInfo: '.CTM'
              }, {
                elementName: 'nary',
                typeInfo: '.CTNary'
              }, {
                elementName: 'phant',
                typeInfo: '.CTPhant'
              }, {
                elementName: 'rad',
                typeInfo: '.CTRad'
              }, {
                elementName: 'sPre',
                typeInfo: '.CTSPre'
              }, {
                elementName: 'sSub',
                typeInfo: '.CTSSub'
              }, {
                elementName: 'sSubSup',
                typeInfo: '.CTSSubSup'
              }, {
                elementName: 'sSup',
                typeInfo: '.CTSSup'
              }, {
                elementName: 'r',
                typeInfo: '.CTR'
              }, {
                elementName: {
                  localPart: 'proofErr',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.ProofErr'
              }, {
                elementName: {
                  localPart: 'permStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.RangePermissionStart'
              }, {
                elementName: {
                  localPart: 'permEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTPerm'
              }, {
                elementName: {
                  localPart: 'bookmarkStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTBookmark'
              }, {
                elementName: {
                  localPart: 'bookmarkEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMarkupRange'
              }, {
                elementName: {
                  localPart: 'moveFromRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMoveBookmark'
              }, {
                elementName: {
                  localPart: 'moveFromRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMoveFromRangeEnd'
              }, {
                elementName: {
                  localPart: 'moveToRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMoveToRangeEnd'
              }, {
                elementName: {
                  localPart: 'moveToRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMoveBookmark'
              }, {
                elementName: {
                  localPart: 'commentRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CommentRangeStart'
              }, {
                elementName: {
                  localPart: 'commentRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CommentRangeEnd'
              }, {
                elementName: {
                  localPart: 'customXmlInsRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTTrackChange'
              }, {
                elementName: {
                  localPart: 'customXmlInsRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMarkup'
              }, {
                elementName: {
                  localPart: 'customXmlDelRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTTrackChange'
              }, {
                elementName: {
                  localPart: 'customXmlDelRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMarkup'
              }, {
                elementName: {
                  localPart: 'customXmlMoveFromRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTTrackChange'
              }, {
                elementName: {
                  localPart: 'customXmlMoveFromRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMarkup'
              }, {
                elementName: {
                  localPart: 'customXmlMoveToRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTTrackChange'
              }, {
                elementName: {
                  localPart: 'customXmlMoveToRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMarkup'
              }, {
                elementName: {
                  localPart: 'ins',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.RunIns'
              }, {
                elementName: {
                  localPart: 'del',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.RunDel'
              }, {
                elementName: {
                  localPart: 'moveFrom',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.RunTrackChange'
              }, {
                elementName: {
                  localPart: 'moveTo',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.RunTrackChange'
              }, {
                elementName: 'oMathPara',
                typeInfo: '.CTOMathPara'
              }, {
                elementName: 'oMath',
                typeInfo: '.CTOMath'
              }],
            type: 'elementRefs'
          }]
      }, {
        localName: 'CTOMathArg',
        typeName: 'CT_OMathArg',
        propertyInfos: [{
            name: 'argPr',
            typeInfo: '.CTOMathArgPr'
          }, {
            name: 'egoMathElements',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'acc',
                typeInfo: '.CTAcc'
              }, {
                elementName: 'bar',
                typeInfo: '.CTBar'
              }, {
                elementName: 'box',
                typeInfo: '.CTBox'
              }, {
                elementName: 'borderBox',
                typeInfo: '.CTBorderBox'
              }, {
                elementName: 'd',
                typeInfo: '.CTD'
              }, {
                elementName: 'eqArr',
                typeInfo: '.CTEqArr'
              }, {
                elementName: 'f',
                typeInfo: '.CTF'
              }, {
                elementName: 'func',
                typeInfo: '.CTFunc'
              }, {
                elementName: 'groupChr',
                typeInfo: '.CTGroupChr'
              }, {
                elementName: 'limLow',
                typeInfo: '.CTLimLow'
              }, {
                elementName: 'limUpp',
                typeInfo: '.CTLimUpp'
              }, {
                elementName: 'm',
                typeInfo: '.CTM'
              }, {
                elementName: 'nary',
                typeInfo: '.CTNary'
              }, {
                elementName: 'phant',
                typeInfo: '.CTPhant'
              }, {
                elementName: 'rad',
                typeInfo: '.CTRad'
              }, {
                elementName: 'sPre',
                typeInfo: '.CTSPre'
              }, {
                elementName: 'sSub',
                typeInfo: '.CTSSub'
              }, {
                elementName: 'sSubSup',
                typeInfo: '.CTSSubSup'
              }, {
                elementName: 'sSup',
                typeInfo: '.CTSSup'
              }, {
                elementName: 'r',
                typeInfo: '.CTR'
              }, {
                elementName: {
                  localPart: 'proofErr',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.ProofErr'
              }, {
                elementName: {
                  localPart: 'permStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.RangePermissionStart'
              }, {
                elementName: {
                  localPart: 'permEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTPerm'
              }, {
                elementName: {
                  localPart: 'bookmarkStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTBookmark'
              }, {
                elementName: {
                  localPart: 'bookmarkEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMarkupRange'
              }, {
                elementName: {
                  localPart: 'moveFromRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMoveBookmark'
              }, {
                elementName: {
                  localPart: 'moveFromRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMoveFromRangeEnd'
              }, {
                elementName: {
                  localPart: 'moveToRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMoveToRangeEnd'
              }, {
                elementName: {
                  localPart: 'moveToRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMoveBookmark'
              }, {
                elementName: {
                  localPart: 'commentRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CommentRangeStart'
              }, {
                elementName: {
                  localPart: 'commentRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CommentRangeEnd'
              }, {
                elementName: {
                  localPart: 'customXmlInsRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTTrackChange'
              }, {
                elementName: {
                  localPart: 'customXmlInsRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMarkup'
              }, {
                elementName: {
                  localPart: 'customXmlDelRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTTrackChange'
              }, {
                elementName: {
                  localPart: 'customXmlDelRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMarkup'
              }, {
                elementName: {
                  localPart: 'customXmlMoveFromRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTTrackChange'
              }, {
                elementName: {
                  localPart: 'customXmlMoveFromRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMarkup'
              }, {
                elementName: {
                  localPart: 'customXmlMoveToRangeStart',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTTrackChange'
              }, {
                elementName: {
                  localPart: 'customXmlMoveToRangeEnd',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMarkup'
              }, {
                elementName: {
                  localPart: 'ins',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.RunIns'
              }, {
                elementName: {
                  localPart: 'del',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.RunDel'
              }, {
                elementName: {
                  localPart: 'moveFrom',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.RunTrackChange'
              }, {
                elementName: {
                  localPart: 'moveTo',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.RunTrackChange'
              }, {
                elementName: 'oMathPara',
                typeInfo: '.CTOMathPara'
              }, {
                elementName: 'oMath',
                typeInfo: '.CTOMath'
              }],
            type: 'elementRefs'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTOMathArgPr',
        typeName: 'CT_OMathArgPr',
        propertyInfos: [{
            name: 'argSz',
            typeInfo: '.CTInteger2'
          }]
      }, {
        localName: 'CTOMathJc',
        typeName: 'CT_OMathJc',
        propertyInfos: [{
            name: 'val',
            typeInfo: '.STJc',
            type: 'attribute'
          }]
      }, {
        localName: 'CTOMathPara',
        typeName: 'CT_OMathPara',
        propertyInfos: [{
            name: 'oMathParaPr',
            typeInfo: '.CTOMathParaPr'
          }, {
            name: 'oMath',
            required: true,
            collection: true,
            typeInfo: '.CTOMath'
          }]
      }, {
        localName: 'CTOMathParaPr',
        typeName: 'CT_OMathParaPr',
        propertyInfos: [{
            name: 'jc',
            typeInfo: '.CTOMathJc'
          }]
      }, {
        localName: 'CTOnOff',
        typeName: 'CT_OnOff',
        propertyInfos: [{
            name: 'val',
            typeInfo: 'org_docx4j_sharedtypes.STOnOff',
            type: 'attribute'
          }]
      }, {
        localName: 'CTPhant',
        typeName: 'CT_Phant',
        propertyInfos: [{
            name: 'phantPr',
            typeInfo: '.CTPhantPr'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTPhantPr',
        typeName: 'CT_PhantPr',
        propertyInfos: [{
            name: 'show',
            typeInfo: '.CTOnOff'
          }, {
            name: 'zeroWid',
            typeInfo: '.CTOnOff'
          }, {
            name: 'zeroAsc',
            typeInfo: '.CTOnOff'
          }, {
            name: 'zeroDesc',
            typeInfo: '.CTOnOff'
          }, {
            name: 'transp',
            typeInfo: '.CTOnOff'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTR',
        typeName: 'CT_R',
        propertyInfos: [{
            name: 'content',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'rPr',
                typeInfo: '.CTRPR'
              }, {
                elementName: {
                  localPart: 'rPr',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.RPr'
              }, {
                elementName: {
                  localPart: 'br',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.Br'
              }, {
                elementName: {
                  localPart: 't',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.Text'
              }, {
                elementName: {
                  localPart: 'delText',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.DelText'
              }, {
                elementName: {
                  localPart: 'instrText',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.Text'
              }, {
                elementName: {
                  localPart: 'delInstrText',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.Text'
              }, {
                elementName: {
                  localPart: 'noBreakHyphen',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.NoBreakHyphen'
              }, {
                elementName: {
                  localPart: 'softHyphen',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.SoftHyphen'
              }, {
                elementName: {
                  localPart: 'dayShort',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.DayShort'
              }, {
                elementName: {
                  localPart: 'monthShort',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.MonthShort'
              }, {
                elementName: {
                  localPart: 'yearShort',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.YearShort'
              }, {
                elementName: {
                  localPart: 'dayLong',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.DayLong'
              }, {
                elementName: {
                  localPart: 'monthLong',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.MonthLong'
              }, {
                elementName: {
                  localPart: 'yearLong',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.YearLong'
              }, {
                elementName: {
                  localPart: 'annotationRef',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.AnnotationRef'
              }, {
                elementName: {
                  localPart: 'footnoteRef',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.FootnoteRef'
              }, {
                elementName: {
                  localPart: 'endnoteRef',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.EndnoteRef'
              }, {
                elementName: {
                  localPart: 'separator',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.Separator'
              }, {
                elementName: {
                  localPart: 'continuationSeparator',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.ContinuationSeparator'
              }, {
                elementName: {
                  localPart: 'sym',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.Sym'
              }, {
                elementName: {
                  localPart: 'pgNum',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.PgNum'
              }, {
                elementName: {
                  localPart: 'cr',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.Cr'
              }, {
                elementName: {
                  localPart: 'tab',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.Tab'
              }, {
                elementName: {
                  localPart: 'object',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTObject'
              }, {
                elementName: {
                  localPart: 'pict',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.Pict'
              }, {
                elementName: {
                  localPart: 'fldChar',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.FldChar'
              }, {
                elementName: {
                  localPart: 'ruby',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTRuby'
              }, {
                elementName: {
                  localPart: 'footnoteReference',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTFtnEdnRef'
              }, {
                elementName: {
                  localPart: 'endnoteReference',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTFtnEdnRef'
              }, {
                elementName: {
                  localPart: 'commentReference',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.CommentReference'
              }, {
                elementName: {
                  localPart: 'drawing',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.Drawing'
              }, {
                elementName: {
                  localPart: 'ptab',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.Ptab'
              }, {
                elementName: {
                  localPart: 'lastRenderedPageBreak',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.R.LastRenderedPageBreak'
              }, {
                elementName: {
                  localPart: 'AlternateContent',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/markup-compatibility\/2006'
                },
                typeInfo: 'org_docx4j_mce.AlternateContent'
              }, {
                elementName: 't',
                typeInfo: '.CTText'
              }, {
                elementName: {
                  localPart: 'ins',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMathRunTrackChange'
              }, {
                elementName: {
                  localPart: 'del',
                  namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
                },
                typeInfo: 'org_docx4j_wml.CTMathRunTrackChange'
              }],
            type: 'elementRefs'
          }]
      }, {
        localName: 'CTRPR',
        typeName: 'CT_RPR',
        propertyInfos: [{
            name: 'lit',
            typeInfo: '.CTOnOff'
          }, {
            name: 'nor',
            typeInfo: '.CTOnOff'
          }, {
            name: 'scr',
            typeInfo: '.CTScript'
          }, {
            name: 'sty',
            typeInfo: '.CTStyle'
          }, {
            name: 'brk',
            typeInfo: '.CTManualBreak'
          }, {
            name: 'aln',
            typeInfo: '.CTOnOff'
          }]
      }, {
        localName: 'CTRad',
        typeName: 'CT_Rad',
        propertyInfos: [{
            name: 'radPr',
            typeInfo: '.CTRadPr'
          }, {
            name: 'deg',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTRadPr',
        typeName: 'CT_RadPr',
        propertyInfos: [{
            name: 'degHide',
            typeInfo: '.CTOnOff'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTSPre',
        typeName: 'CT_SPre',
        propertyInfos: [{
            name: 'sPrePr',
            typeInfo: '.CTSPrePr'
          }, {
            name: 'sub',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'sup',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTSPrePr',
        typeName: 'CT_SPrePr',
        propertyInfos: [{
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTSSub',
        typeName: 'CT_SSub',
        propertyInfos: [{
            name: 'sSubPr',
            typeInfo: '.CTSSubPr'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'sub',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTSSubPr',
        typeName: 'CT_SSubPr',
        propertyInfos: [{
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTSSubSup',
        typeName: 'CT_SSubSup',
        propertyInfos: [{
            name: 'sSubSupPr',
            typeInfo: '.CTSSubSupPr'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'sub',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'sup',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTSSubSupPr',
        typeName: 'CT_SSubSupPr',
        propertyInfos: [{
            name: 'alnScr',
            typeInfo: '.CTOnOff'
          }, {
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTSSup',
        typeName: 'CT_SSup',
        propertyInfos: [{
            name: 'sSupPr',
            typeInfo: '.CTSSupPr'
          }, {
            name: 'e',
            required: true,
            typeInfo: '.CTOMathArg'
          }, {
            name: 'sup',
            required: true,
            typeInfo: '.CTOMathArg'
          }]
      }, {
        localName: 'CTSSupPr',
        typeName: 'CT_SSupPr',
        propertyInfos: [{
            name: 'ctrlPr',
            typeInfo: '.CTCtrlPr'
          }]
      }, {
        localName: 'CTScript',
        typeName: 'CT_Script',
        propertyInfos: [{
            name: 'val',
            typeInfo: '.STScript',
            type: 'attribute'
          }]
      }, {
        localName: 'CTShp',
        typeName: 'CT_Shp',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: '.STShp',
            type: 'attribute'
          }]
      }, {
        localName: 'CTSpacingRule',
        typeName: 'CT_SpacingRule',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'Integer',
            type: 'attribute'
          }]
      }, {
        localName: 'CTString',
        typeName: 'CT_String',
        propertyInfos: [{
            name: 'val',
            type: 'attribute'
          }]
      }, {
        localName: 'CTStyle',
        typeName: 'CT_Style',
        propertyInfos: [{
            name: 'val',
            typeInfo: '.STStyle',
            type: 'attribute'
          }]
      }, {
        localName: 'CTText',
        typeName: 'CT_Text',
        propertyInfos: [{
            name: 'value',
            type: 'value'
          }, {
            name: 'space',
            values: ['default', 'preserve'],
            attributeName: {
              localPart: 'space',
              namespaceURI: 'http:\/\/www.w3.org\/XML\/1998\/namespace'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTTopBot',
        typeName: 'CT_TopBot',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: '.STTopBot',
            type: 'attribute'
          }]
      }, {
        localName: 'CTTwipsMeasure',
        typeName: 'CT_TwipsMeasure',
        propertyInfos: [{
            name: 'val',
            required: true,
            type: 'attribute'
          }]
      }, {
        localName: 'CTUnSignedInteger',
        typeName: 'CT_UnSignedInteger',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'UnsignedInt',
            type: 'attribute'
          }]
      }, {
        localName: 'CTXAlign',
        typeName: 'CT_XAlign',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'org_docx4j_sharedtypes.STXAlign',
            type: 'attribute'
          }]
      }, {
        localName: 'CTYAlign',
        typeName: 'CT_YAlign',
        propertyInfos: [{
            name: 'val',
            required: true,
            typeInfo: 'org_docx4j_sharedtypes.STYAlign',
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STBreakBin',
        values: ['before', 'after', 'repeat']
      }, {
        type: 'enumInfo',
        localName: 'STFType',
        values: ['bar', 'skw', 'lin', 'noBar']
      }, {
        type: 'enumInfo',
        localName: 'STJc',
        values: ['left', 'right', 'center', 'centerGroup']
      }, {
        type: 'enumInfo',
        localName: 'STLimLoc',
        values: ['undOvr', 'subSup']
      }, {
        type: 'enumInfo',
        localName: 'STScript',
        values: ['roman', 'script', 'fraktur', 'double-struck', 'sans-serif', 'monospace']
      }, {
        type: 'enumInfo',
        localName: 'STShp',
        values: ['centered', 'match']
      }, {
        type: 'enumInfo',
        localName: 'STStyle',
        values: ['p', 'b', 'i', 'bi']
      }, {
        type: 'enumInfo',
        localName: 'STTopBot',
        values: ['top', 'bot']
      }],
    elementInfos: [{
        typeInfo: '.CTAcc',
        elementName: 'acc',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTAcc',
        elementName: 'acc',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTBar',
        elementName: 'bar',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTBar',
        elementName: 'bar',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTBorderBox',
        elementName: 'borderBox',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTBorderBox',
        elementName: 'borderBox',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTBox',
        elementName: 'box',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTBox',
        elementName: 'box',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTD',
        elementName: 'd',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTD',
        elementName: 'd',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTEqArr',
        elementName: 'eqArr',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTEqArr',
        elementName: 'eqArr',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTF',
        elementName: 'f',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTF',
        elementName: 'f',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTFunc',
        elementName: 'func',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTFunc',
        elementName: 'func',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTGroupChr',
        elementName: 'groupChr',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTGroupChr',
        elementName: 'groupChr',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTLimLow',
        elementName: 'limLow',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTLimLow',
        elementName: 'limLow',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTLimUpp',
        elementName: 'limUpp',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTLimUpp',
        elementName: 'limUpp',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTM',
        elementName: 'm',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTM',
        elementName: 'm',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTMathPr',
        elementName: 'mathPr'
      }, {
        typeInfo: '.CTNary',
        elementName: 'nary',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTNary',
        elementName: 'nary',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTOMath',
        elementName: 'oMath'
      }, {
        typeInfo: '.CTOMathPara',
        elementName: 'oMathPara'
      }, {
        typeInfo: '.CTPhant',
        elementName: 'phant',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTPhant',
        elementName: 'phant',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTR',
        elementName: 'r',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTR',
        elementName: 'r',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTRPR',
        elementName: 'rPr',
        scope: '.CTR'
      }, {
        typeInfo: '.CTRad',
        elementName: 'rad',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTRad',
        elementName: 'rad',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTSPre',
        elementName: 'sPre',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTSPre',
        elementName: 'sPre',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTSSub',
        elementName: 'sSub',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTSSub',
        elementName: 'sSub',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTSSubSup',
        elementName: 'sSubSup',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTSSubSup',
        elementName: 'sSubSup',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTSSup',
        elementName: 'sSup',
        scope: '.CTOMath'
      }, {
        typeInfo: '.CTSSup',
        elementName: 'sSup',
        scope: '.CTOMathArg'
      }, {
        typeInfo: '.CTText',
        elementName: 't',
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.AnnotationRef',
        elementName: {
          localPart: 'annotationRef',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.CTMarkupRange',
        elementName: {
          localPart: 'bookmarkEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTMarkupRange',
        elementName: {
          localPart: 'bookmarkEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.CTBookmark',
        elementName: {
          localPart: 'bookmarkStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTBookmark',
        elementName: {
          localPart: 'bookmarkStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.R.CommentReference',
        elementName: {
          localPart: 'commentReference',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.ContinuationSeparator',
        elementName: {
          localPart: 'continuationSeparator',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.Cr',
        elementName: {
          localPart: 'cr',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.CTMarkup',
        elementName: {
          localPart: 'customXmlDelRangeEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTMarkup',
        elementName: {
          localPart: 'customXmlDelRangeEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.CTTrackChange',
        elementName: {
          localPart: 'customXmlDelRangeStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTTrackChange',
        elementName: {
          localPart: 'customXmlDelRangeStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.CTMarkup',
        elementName: {
          localPart: 'customXmlInsRangeEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTMarkup',
        elementName: {
          localPart: 'customXmlInsRangeEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.CTTrackChange',
        elementName: {
          localPart: 'customXmlInsRangeStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTTrackChange',
        elementName: {
          localPart: 'customXmlInsRangeStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.CTMarkup',
        elementName: {
          localPart: 'customXmlMoveFromRangeEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTMarkup',
        elementName: {
          localPart: 'customXmlMoveFromRangeEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.CTTrackChange',
        elementName: {
          localPart: 'customXmlMoveFromRangeStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTTrackChange',
        elementName: {
          localPart: 'customXmlMoveFromRangeStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.CTMarkup',
        elementName: {
          localPart: 'customXmlMoveToRangeEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTMarkup',
        elementName: {
          localPart: 'customXmlMoveToRangeEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.CTTrackChange',
        elementName: {
          localPart: 'customXmlMoveToRangeStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTTrackChange',
        elementName: {
          localPart: 'customXmlMoveToRangeStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.R.DayLong',
        elementName: {
          localPart: 'dayLong',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.DayShort',
        elementName: {
          localPart: 'dayShort',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.CTMathRunTrackChange',
        elementName: {
          localPart: 'del',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.Text',
        elementName: {
          localPart: 'delInstrText',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.Drawing',
        elementName: {
          localPart: 'drawing',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.EndnoteRef',
        elementName: {
          localPart: 'endnoteRef',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.CTFtnEdnRef',
        elementName: {
          localPart: 'endnoteReference',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.FldChar',
        elementName: {
          localPart: 'fldChar',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.FootnoteRef',
        elementName: {
          localPart: 'footnoteRef',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.CTFtnEdnRef',
        elementName: {
          localPart: 'footnoteReference',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.CTMathRunTrackChange',
        elementName: {
          localPart: 'ins',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.Text',
        elementName: {
          localPart: 'instrText',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.LastRenderedPageBreak',
        elementName: {
          localPart: 'lastRenderedPageBreak',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.MonthLong',
        elementName: {
          localPart: 'monthLong',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.MonthShort',
        elementName: {
          localPart: 'monthShort',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.RunTrackChange',
        elementName: {
          localPart: 'moveFrom',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.RunTrackChange',
        elementName: {
          localPart: 'moveFrom',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.CTMoveFromRangeEnd',
        elementName: {
          localPart: 'moveFromRangeEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTMoveFromRangeEnd',
        elementName: {
          localPart: 'moveFromRangeEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.CTMoveBookmark',
        elementName: {
          localPart: 'moveFromRangeStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTMoveBookmark',
        elementName: {
          localPart: 'moveFromRangeStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.RunTrackChange',
        elementName: {
          localPart: 'moveTo',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.RunTrackChange',
        elementName: {
          localPart: 'moveTo',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.CTMoveToRangeEnd',
        elementName: {
          localPart: 'moveToRangeEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTMoveToRangeEnd',
        elementName: {
          localPart: 'moveToRangeEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.CTMoveBookmark',
        elementName: {
          localPart: 'moveToRangeStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTMoveBookmark',
        elementName: {
          localPart: 'moveToRangeStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.R.NoBreakHyphen',
        elementName: {
          localPart: 'noBreakHyphen',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.CTObject',
        elementName: {
          localPart: 'object',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.CTPerm',
        elementName: {
          localPart: 'permEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.CTPerm',
        elementName: {
          localPart: 'permEnd',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.RangePermissionStart',
        elementName: {
          localPart: 'permStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMath'
      }, {
        typeInfo: 'org_docx4j_wml.RangePermissionStart',
        elementName: {
          localPart: 'permStart',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTOMathArg'
      }, {
        typeInfo: 'org_docx4j_wml.R.PgNum',
        elementName: {
          localPart: 'pgNum',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.Pict',
        elementName: {
          localPart: 'pict',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.Ptab',
        elementName: {
          localPart: 'ptab',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.RPr',
        elementName: {
          localPart: 'rPr',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.CTRuby',
        elementName: {
          localPart: 'ruby',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.Separator',
        elementName: {
          localPart: 'separator',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.SoftHyphen',
        elementName: {
          localPart: 'softHyphen',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.Sym',
        elementName: {
          localPart: 'sym',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.Text',
        elementName: {
          localPart: 't',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.Tab',
        elementName: {
          localPart: 'tab',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.YearLong',
        elementName: {
          localPart: 'yearLong',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }, {
        typeInfo: 'org_docx4j_wml.R.YearShort',
        elementName: {
          localPart: 'yearShort',
          namespaceURI: 'http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main'
        },
        scope: '.CTR'
      }]
  };
  return {
    org_docx4j_math: org_docx4j_math
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_math_Module_Factory);
}
else {
  var org_docx4j_math_Module = org_docx4j_math_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_math = org_docx4j_math_Module.org_docx4j_math;
  }
  else {
    var org_docx4j_math = org_docx4j_math_Module.org_docx4j_math;
  }
}