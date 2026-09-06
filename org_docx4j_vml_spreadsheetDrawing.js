var org_docx4j_vml_spreadsheetDrawing_Module_Factory = function () {
  var org_docx4j_vml_spreadsheetDrawing = {
    name: 'org_docx4j_vml_spreadsheetDrawing',
    defaultElementNamespaceURI: 'urn:schemas-microsoft-com:office:excel',
    typeInfos: [{
        localName: 'CTClientData',
        typeName: 'CT_ClientData',
        propertyInfos: [{
            name: 'moveWithCellsOrSizeWithCellsOrAnchor',
            minOccurs: 0,
            collection: true,
            mixed: false,
            allowDom: false,
            elementTypeInfos: [{
                elementName: 'MoveWithCells',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'SizeWithCells',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'Anchor'
              }, {
                elementName: 'Locked',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'DefaultSize',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'PrintObject',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'Disabled',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'AutoFill',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'AutoLine',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'AutoPict',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'FmlaMacro'
              }, {
                elementName: 'TextHAlign'
              }, {
                elementName: 'TextVAlign'
              }, {
                elementName: 'LockText',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'JustLastX',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'SecretEdit',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'Default',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'Help',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'Cancel',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'Dismiss',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'Accel',
                typeInfo: 'Integer'
              }, {
                elementName: 'Accel2',
                typeInfo: 'Integer'
              }, {
                elementName: 'Row',
                typeInfo: 'Integer'
              }, {
                elementName: 'Column',
                typeInfo: 'Integer'
              }, {
                elementName: 'Visible',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'RowHidden',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'ColHidden',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'VTEdit',
                typeInfo: 'Integer'
              }, {
                elementName: 'MultiLine',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'VScroll',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'ValidIds',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'FmlaRange'
              }, {
                elementName: 'WidthMin',
                typeInfo: 'Integer'
              }, {
                elementName: 'Sel',
                typeInfo: 'Integer'
              }, {
                elementName: 'NoThreeD2',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'SelType'
              }, {
                elementName: 'MultiSel'
              }, {
                elementName: 'LCT'
              }, {
                elementName: 'ListItem'
              }, {
                elementName: 'DropStyle'
              }, {
                elementName: 'Colored',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'DropLines',
                typeInfo: 'Integer'
              }, {
                elementName: 'Checked',
                typeInfo: 'Integer'
              }, {
                elementName: 'FmlaLink'
              }, {
                elementName: 'FmlaPict'
              }, {
                elementName: 'NoThreeD',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'FirstButton',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'FmlaGroup'
              }, {
                elementName: 'Val',
                typeInfo: 'Integer'
              }, {
                elementName: 'Min',
                typeInfo: 'Integer'
              }, {
                elementName: 'Max',
                typeInfo: 'Integer'
              }, {
                elementName: 'Inc',
                typeInfo: 'Integer'
              }, {
                elementName: 'Page',
                typeInfo: 'Integer'
              }, {
                elementName: 'Horiz',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'Dx',
                typeInfo: 'Integer'
              }, {
                elementName: 'MapOCX',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'CF',
                typeInfo: '.STCF'
              }, {
                elementName: 'Camera',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'RecalcAlways',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'AutoScale',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'DDE',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'UIObj',
                values: ['True', 't', 'False', 'f', '']
              }, {
                elementName: 'ScriptText'
              }, {
                elementName: 'ScriptExtended'
              }, {
                elementName: 'ScriptLanguage',
                typeInfo: 'NonNegativeInteger'
              }, {
                elementName: 'ScriptLocation',
                typeInfo: 'NonNegativeInteger'
              }, {
                elementName: 'FmlaTxbx'
              }],
            type: 'elementRefs'
          }, {
            name: 'objectType',
            required: true,
            typeInfo: '.STObjectType',
            attributeName: {
              localPart: 'ObjectType'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STCF',
        values: ['PictOld', 'Pict', 'Bitmap', 'PictPrint', 'PictScreen']
      }, {
        type: 'enumInfo',
        localName: 'STObjectType',
        values: ['Button', 'Checkbox', 'Dialog', 'Drop', 'Edit', 'GBox', 'Label', 'LineA', 'List', 'Movie', 'Note', 'Pict', 'Radio', 'RectA', 'Scroll', 'Spin', 'Shape', 'Group', 'Rect']
      }],
    elementInfos: [{
        typeInfo: 'Integer',
        elementName: 'Accel',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'Accel2',
        scope: '.CTClientData'
      }, {
        elementName: 'Anchor',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'AutoFill',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'AutoLine',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'AutoPict',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'AutoScale',
        scope: '.CTClientData'
      }, {
        typeInfo: '.STCF',
        elementName: 'CF',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'Camera',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'Cancel',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'Checked',
        scope: '.CTClientData'
      }, {
        typeInfo: '.CTClientData',
        elementName: 'ClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'ColHidden',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'Colored',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'Column',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'DDE',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'Default',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'DefaultSize',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'Disabled',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'Dismiss',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'DropLines',
        scope: '.CTClientData'
      }, {
        elementName: 'DropStyle',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'Dx',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'FirstButton',
        scope: '.CTClientData'
      }, {
        elementName: 'FmlaGroup',
        scope: '.CTClientData'
      }, {
        elementName: 'FmlaLink',
        scope: '.CTClientData'
      }, {
        elementName: 'FmlaMacro',
        scope: '.CTClientData'
      }, {
        elementName: 'FmlaPict',
        scope: '.CTClientData'
      }, {
        elementName: 'FmlaRange',
        scope: '.CTClientData'
      }, {
        elementName: 'FmlaTxbx',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'Help',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'Horiz',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'Inc',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'JustLastX',
        scope: '.CTClientData'
      }, {
        elementName: 'LCT',
        scope: '.CTClientData'
      }, {
        elementName: 'ListItem',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'LockText',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'Locked',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'MapOCX',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'Max',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'Min',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'MoveWithCells',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'MultiLine',
        scope: '.CTClientData'
      }, {
        elementName: 'MultiSel',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'NoThreeD',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'NoThreeD2',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'Page',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'PrintObject',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'RecalcAlways',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'Row',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'RowHidden',
        scope: '.CTClientData'
      }, {
        elementName: 'ScriptExtended',
        scope: '.CTClientData'
      }, {
        typeInfo: 'NonNegativeInteger',
        elementName: 'ScriptLanguage',
        scope: '.CTClientData'
      }, {
        typeInfo: 'NonNegativeInteger',
        elementName: 'ScriptLocation',
        scope: '.CTClientData'
      }, {
        elementName: 'ScriptText',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'SecretEdit',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'Sel',
        scope: '.CTClientData'
      }, {
        elementName: 'SelType',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'SizeWithCells',
        scope: '.CTClientData'
      }, {
        elementName: 'TextHAlign',
        scope: '.CTClientData'
      }, {
        elementName: 'TextVAlign',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'UIObj',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'VScroll',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'VTEdit',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'Val',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'ValidIds',
        scope: '.CTClientData'
      }, {
        values: ['True', 't', 'False', 'f', ''],
        elementName: 'Visible',
        scope: '.CTClientData'
      }, {
        typeInfo: 'Integer',
        elementName: 'WidthMin',
        scope: '.CTClientData'
      }]
  };
  return {
    org_docx4j_vml_spreadsheetDrawing: org_docx4j_vml_spreadsheetDrawing
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_vml_spreadsheetDrawing_Module_Factory);
}
else {
  var org_docx4j_vml_spreadsheetDrawing_Module = org_docx4j_vml_spreadsheetDrawing_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_vml_spreadsheetDrawing = org_docx4j_vml_spreadsheetDrawing_Module.org_docx4j_vml_spreadsheetDrawing;
  }
  else {
    var org_docx4j_vml_spreadsheetDrawing = org_docx4j_vml_spreadsheetDrawing_Module.org_docx4j_vml_spreadsheetDrawing;
  }
}