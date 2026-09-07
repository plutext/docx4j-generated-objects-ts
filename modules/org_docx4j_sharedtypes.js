var org_docx4j_sharedtypes_Module_Factory = function () {
  var org_docx4j_sharedtypes = {
    name: 'org_docx4j_sharedtypes',
    typeInfos: [{
        type: 'enumInfo',
        localName: 'STCalendarType',
        values: ['gregorian', 'gregorianUs', 'gregorianMeFrench', 'gregorianArabic', 'hijri', 'hebrew', 'taiwan', 'japan', 'thai', 'korea', 'saka', 'gregorianXlitEnglish', 'gregorianXlitFrench', 'none']
      }, {
        type: 'enumInfo',
        localName: 'STConformanceClass',
        values: ['strict', 'transitional']
      }, {
        type: 'enumInfo',
        localName: 'STOnOff',
        values: ['true', 'false', 'on', 'off', '0', '1']
      }, {
        type: 'enumInfo',
        localName: 'STVerticalAlignRun',
        values: ['baseline', 'superscript', 'subscript']
      }, {
        type: 'enumInfo',
        localName: 'STXAlign',
        values: ['left', 'center', 'right', 'inside', 'outside']
      }, {
        type: 'enumInfo',
        localName: 'STYAlign',
        values: ['inline', 'top', 'center', 'bottom', 'inside', 'outside']
      }],
    elementInfos: []
  };
  return {
    org_docx4j_sharedtypes: org_docx4j_sharedtypes
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_sharedtypes_Module_Factory);
}
else {
  var org_docx4j_sharedtypes_Module = org_docx4j_sharedtypes_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_sharedtypes = org_docx4j_sharedtypes_Module.org_docx4j_sharedtypes;
  }
  else {
    var org_docx4j_sharedtypes = org_docx4j_sharedtypes_Module.org_docx4j_sharedtypes;
  }
}