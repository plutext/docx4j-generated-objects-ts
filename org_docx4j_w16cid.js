var org_docx4j_w16cid_Module_Factory = function () {
  var org_docx4j_w16cid = {
    name: 'org_docx4j_w16cid',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2016\/wordml\/cid',
    defaultAttributeNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/word\/2016\/wordml\/cid',
    typeInfos: [{
        localName: 'CTCommentId',
        typeName: 'CT_CommentId',
        propertyInfos: [{
            name: 'paraId',
            type: 'attribute'
          }, {
            name: 'durableId',
            type: 'attribute'
          }]
      }, {
        localName: 'CTCommentsIds',
        typeName: 'CT_CommentsIds',
        propertyInfos: [{
            name: 'commentId',
            minOccurs: 0,
            collection: true,
            typeInfo: '.CTCommentId'
          }, {
            name: 'ignorable',
            attributeName: {
              localPart: 'Ignorable',
              namespaceURI: 'http:\/\/schemas.openxmlformats.org\/markup-compatibility\/2006'
            },
            type: 'attribute'
          }]
      }],
    elementInfos: [{
        typeInfo: '.CTCommentsIds',
        elementName: 'commentsIds'
      }]
  };
  return {
    org_docx4j_w16cid: org_docx4j_w16cid
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_w16cid_Module_Factory);
}
else {
  var org_docx4j_w16cid_Module = org_docx4j_w16cid_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_w16cid = org_docx4j_w16cid_Module.org_docx4j_w16cid;
  }
  else {
    var org_docx4j_w16cid = org_docx4j_w16cid_Module.org_docx4j_w16cid;
  }
}