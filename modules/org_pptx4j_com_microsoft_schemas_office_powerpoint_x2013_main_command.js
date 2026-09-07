var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command_Module_Factory = function () {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command = {
    name: 'org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/powerpoint\/2013\/main\/command',
    dependencies: ['org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command', 'org_pptx4j_pml'],
    typeInfos: [{
        localName: 'CTChangesInfo',
        typeName: 'CT_ChangesInfo',
        propertyInfos: [{
            name: 'docChgLst',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTDocumentChangesList'
          }]
      }, {
        localName: 'CTCommentAuthorChanges',
        typeName: 'CT_CommentAuthorChanges',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTChangesData'
          }, {
            name: 'cmAuthorMkLst',
            required: true,
            typeInfo: '.CTCommentAuthorMonikerList'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }, {
            name: 'chg',
            required: true,
            typeInfo: {
              type: 'list',
              baseTypeInfo: '.STCommentAuthorChangeBit'
            },
            attributeName: {
              localPart: 'chg'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCommentAuthorMoniker',
        typeName: 'CT_CommentAuthorMoniker',
        propertyInfos: [{
            name: 'id',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'id'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCommentAuthorMonikerList',
        typeName: 'CT_CommentAuthorMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'CTCommentChanges',
        typeName: 'CT_CommentChanges',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTChangesData'
          }, {
            name: 'cmMkLst',
            required: true,
            typeInfo: '.CTCommentMonikerList'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }, {
            name: 'chg',
            required: true,
            typeInfo: {
              type: 'list',
              baseTypeInfo: '.STCommentChangeBit'
            },
            attributeName: {
              localPart: 'chg'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCommentMoniker',
        typeName: 'CT_CommentMoniker',
        propertyInfos: [{
            name: 'authorId',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'authorId'
            },
            type: 'attribute'
          }, {
            name: 'idx',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'idx'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTCommentMonikerList',
        typeName: 'CT_CommentMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'CTDocumentChanges',
        typeName: 'CT_DocumentChanges',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTChangesData'
          }, {
            name: 'docMkLst',
            required: true,
            typeInfo: '.CTDocumentMonikerList'
          }, {
            name: 'sldChg',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTSlideChanges'
          }, {
            name: 'sldMasterChg',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTMainMasterChanges'
          }, {
            name: 'cmAuthorChg',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTCommentAuthorChanges'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }, {
            name: 'chg',
            required: true,
            typeInfo: {
              type: 'list',
              baseTypeInfo: '.STDocumentChangeBit'
            },
            attributeName: {
              localPart: 'chg'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTDocumentChangesList',
        typeName: 'CT_DocumentChangesList',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTChangesData'
          }, {
            name: 'docChg',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTDocumentChanges'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }]
      }, {
        localName: 'CTDocumentMoniker',
        typeName: 'CT_DocumentMoniker'
      }, {
        localName: 'CTDocumentMonikerList',
        typeName: 'CT_DocumentMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'CTMainMasterChanges',
        typeName: 'CT_MainMasterChanges',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTChangesData'
          }, {
            name: 'sldMasterMkLst',
            required: true,
            typeInfo: '.CTMainMasterMonikerList'
          }, {
            name: 'spChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTShapeChanges'
          }, {
            name: 'grpChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTGroupShapeChanges'
          }, {
            name: 'graphicFrameChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTGraphicFrameChanges'
          }, {
            name: 'picChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTPictureChanges'
          }, {
            name: 'inkChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTInkChanges'
          }, {
            name: 'cxnChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTConnectorChanges'
          }, {
            name: 'sldLayoutChg',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTSlideLayoutChanges'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }, {
            name: 'chg',
            required: true,
            typeInfo: {
              type: 'list',
              baseTypeInfo: '.STMainMasterChangeBit'
            },
            attributeName: {
              localPart: 'chg'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTMainMasterMoniker',
        typeName: 'CT_MainMasterMoniker',
        propertyInfos: [{
            name: 'cId',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'cId'
            },
            type: 'attribute'
          }, {
            name: 'sldId',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'sldId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTMainMasterMonikerList',
        typeName: 'CT_MainMasterMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'CTSlideBaseMonikerList',
        typeName: 'CT_SlideBaseMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'CTSlideChanges',
        typeName: 'CT_SlideChanges',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTChangesData'
          }, {
            name: 'sldMkLst',
            required: true,
            typeInfo: '.CTSlideMonikerList'
          }, {
            name: 'spChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTShapeChanges'
          }, {
            name: 'grpChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTGroupShapeChanges'
          }, {
            name: 'graphicFrameChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTGraphicFrameChanges'
          }, {
            name: 'picChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTPictureChanges'
          }, {
            name: 'inkChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTInkChanges'
          }, {
            name: 'cxnChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTConnectorChanges'
          }, {
            name: 'cmChg',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTCommentChanges'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }, {
            name: 'chg',
            required: true,
            typeInfo: {
              type: 'list',
              baseTypeInfo: '.STSlideChangeBit'
            },
            attributeName: {
              localPart: 'chg'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSlideLayoutChanges',
        typeName: 'CT_SlideLayoutChanges',
        propertyInfos: [{
            name: 'chgData',
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTChangesData'
          }, {
            name: 'sldLayoutMkLst',
            required: true,
            typeInfo: '.CTSlideLayoutMonikerList'
          }, {
            name: 'spChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTShapeChanges'
          }, {
            name: 'grpChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTGroupShapeChanges'
          }, {
            name: 'graphicFrameChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTGraphicFrameChanges'
          }, {
            name: 'picChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTPictureChanges'
          }, {
            name: 'inkChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTInkChanges'
          }, {
            name: 'cxnChg',
            minOccurs: 0,
            collection: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_drawing_x2013_main_command.CTConnectorChanges'
          }, {
            name: 'extLst',
            typeInfo: 'org_pptx4j_pml.CTExtensionList'
          }, {
            name: 'chg',
            required: true,
            typeInfo: {
              type: 'list',
              baseTypeInfo: '.STSlideLayoutChangeBit'
            },
            attributeName: {
              localPart: 'chg'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSlideLayoutMoniker',
        typeName: 'CT_SlideLayoutMoniker',
        propertyInfos: [{
            name: 'cId',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'cId'
            },
            type: 'attribute'
          }, {
            name: 'sldId',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'sldId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSlideLayoutMonikerList',
        typeName: 'CT_SlideLayoutMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        localName: 'CTSlideMoniker',
        typeName: 'CT_SlideMoniker',
        propertyInfos: [{
            name: 'cId',
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'cId'
            },
            type: 'attribute'
          }, {
            name: 'sldId',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'sldId'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTSlideMonikerList',
        typeName: 'CT_SlideMonikerList',
        propertyInfos: [{
            name: 'any',
            minOccurs: 0,
            collection: true,
            allowDom: false,
            mixed: false,
            type: 'anyElement'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STCommentAuthorChangeBit',
        baseTypeInfo: 'Token',
        values: ['add', 'del', 'mod', 'replId']
      }, {
        type: 'enumInfo',
        localName: 'STCommentChangeBit',
        baseTypeInfo: 'Token',
        values: ['add', 'del', 'mod', 'replId']
      }, {
        type: 'enumInfo',
        localName: 'STDocumentChangeBit',
        baseTypeInfo: 'Token',
        values: ['undo', 'redo', 'ext', 'custSel', 'mod', 'addSld', 'delSld', 'modSld', 'sldOrd', 'addMainMaster', 'delMainMaster', 'modMainMaster', 'mainMasterOrd', 'addSection', 'delSection', 'modSection', 'addCmAuthor', 'delCmAuthor', 'modCmAuthor', 'replTag', 'delTag', 'addCustShow', 'delCustShow', 'modCustShow', 'modNotesMaster', 'modHandout', 'modShowInfo', 'addOsfTaskPaneApp', 'delOsfTaskPaneApp', 'setSldSz', 'modRtl', 'modChgInfo']
      }, {
        type: 'enumInfo',
        localName: 'STMainMasterChangeBit',
        baseTypeInfo: 'Token',
        values: ['addSp', 'delSp', 'modSp', 'spOrd', 'new', 'add', 'del', 'mod', 'ord', 'replId', 'modTransition', 'modMedia', 'setBg', 'setFolMasterAnim', 'setFolMasterObjs', 'modClrScheme', 'addAnim', 'delAnim', 'modAnim', 'replTag', 'delTag', 'setClrOvrMap', 'delDesignElem', 'modShow', 'addSldLayout', 'delSldLayout', 'modSldLayout', 'sldLayoutOrd']
      }, {
        type: 'enumInfo',
        localName: 'STSlideChangeBit',
        baseTypeInfo: 'Token',
        values: ['addSp', 'delSp', 'modSp', 'spOrd', 'new', 'add', 'del', 'mod', 'ord', 'replId', 'modTransition', 'modMedia', 'setBg', 'setFolMasterAnim', 'setFolMasterObjs', 'modClrScheme', 'addAnim', 'delAnim', 'modAnim', 'replTag', 'delTag', 'setClrOvrMap', 'delDesignElem', 'modShow', 'addCm', 'delCm', 'modCm', 'chgLayout', 'modNotes', 'modNotesTx', 'setSldSyncInfo', 'newSectionLinks']
      }, {
        type: 'enumInfo',
        localName: 'STSlideLayoutChangeBit',
        baseTypeInfo: 'Token',
        values: ['addSp', 'delSp', 'modSp', 'spOrd', 'new', 'add', 'del', 'mod', 'ord', 'replId', 'modTransition', 'modMedia', 'setBg', 'setFolMasterAnim', 'setFolMasterObjs', 'modClrScheme', 'addAnim', 'delAnim', 'modAnim', 'replTag', 'delTag', 'setClrOvrMap', 'delDesignElem', 'modShow']
      }],
    elementInfos: [{
        typeInfo: '.CTChangesInfo',
        elementName: 'chgInfo'
      }, {
        typeInfo: '.CTCommentAuthorMonikerList',
        elementName: 'cmAuthorMkLst'
      }, {
        typeInfo: '.CTCommentMonikerList',
        elementName: 'cmMkLst'
      }, {
        typeInfo: '.CTDocumentMonikerList',
        elementName: 'docMkLst'
      }, {
        typeInfo: '.CTSlideBaseMonikerList',
        elementName: 'sldBaseMkLst'
      }, {
        typeInfo: '.CTSlideLayoutMonikerList',
        elementName: 'sldLayoutMkLst'
      }, {
        typeInfo: '.CTMainMasterMonikerList',
        elementName: 'sldMasterMkLst'
      }, {
        typeInfo: '.CTSlideMonikerList',
        elementName: 'sldMkLst'
      }]
  };
  return {
    org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command: org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command_Module_Factory);
}
else {
  var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command_Module = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command;
  }
  else {
    var org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command = org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command_Module.org_pptx4j_com_microsoft_schemas_office_powerpoint_x2013_main_command;
  }
}