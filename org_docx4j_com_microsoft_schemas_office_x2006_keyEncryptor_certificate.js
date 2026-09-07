var org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate = {
    name: 'org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/2006\/keyEncryptor\/certificate',
    typeInfos: [{
        localName: 'CTCertificateKeyEncryptor',
        typeName: 'CT_CertificateKeyEncryptor',
        propertyInfos: [{
            name: 'x509Certificate',
            required: true,
            typeInfo: 'Base64Binary',
            attributeName: {
              localPart: 'X509Certificate'
            },
            type: 'attribute'
          }, {
            name: 'certVerifier',
            required: true,
            typeInfo: 'Base64Binary',
            attributeName: {
              localPart: 'certVerifier'
            },
            type: 'attribute'
          }, {
            name: 'encryptedKeyValue',
            required: true,
            typeInfo: 'Base64Binary',
            attributeName: {
              localPart: 'encryptedKeyValue'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STPasswordKeyEncryptorUri',
        baseTypeInfo: 'Token',
        values: ['http:\/\/schemas.microsoft.com\/office\/2006\/keyEncryptor\/certificate']
      }],
    elementInfos: [{
        typeInfo: '.CTCertificateKeyEncryptor',
        elementName: 'encryptedKey'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate: org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate_Module = org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate = org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate_Module.org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate = org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate_Module.org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate;
  }
}