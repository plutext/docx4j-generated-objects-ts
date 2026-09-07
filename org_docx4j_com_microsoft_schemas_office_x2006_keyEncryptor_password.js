var org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password = {
    name: 'org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/2006\/keyEncryptor\/password',
    dependencies: ['org_docx4j_com_microsoft_schemas_office_x2006_encryption'],
    typeInfos: [{
        localName: 'CTPasswordKeyEncryptor',
        typeName: 'CT_PasswordKeyEncryptor',
        propertyInfos: [{
            name: 'blockSize',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'blockSize'
            },
            type: 'attribute'
          }, {
            name: 'cipherAlgorithm',
            required: true,
            typeInfo: 'Token',
            values: ['AES', 'RC2', 'RC4', 'DES', 'DESX', '3DES', '3DES_112'],
            attributeName: {
              localPart: 'cipherAlgorithm'
            },
            type: 'attribute'
          }, {
            name: 'cipherChaining',
            required: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_x2006_encryption.STCipherChaining',
            attributeName: {
              localPart: 'cipherChaining'
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
          }, {
            name: 'encryptedVerifierHashInput',
            required: true,
            typeInfo: 'Base64Binary',
            attributeName: {
              localPart: 'encryptedVerifierHashInput'
            },
            type: 'attribute'
          }, {
            name: 'encryptedVerifierHashValue',
            required: true,
            typeInfo: 'Base64Binary',
            attributeName: {
              localPart: 'encryptedVerifierHashValue'
            },
            type: 'attribute'
          }, {
            name: 'hashAlgorithm',
            required: true,
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_x2006_encryption.STHashAlgorithm',
            attributeName: {
              localPart: 'hashAlgorithm'
            },
            type: 'attribute'
          }, {
            name: 'hashSize',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'hashSize'
            },
            type: 'attribute'
          }, {
            name: 'keyBits',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'keyBits'
            },
            type: 'attribute'
          }, {
            name: 'saltSize',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'saltSize'
            },
            type: 'attribute'
          }, {
            name: 'saltValue',
            required: true,
            typeInfo: 'Base64Binary',
            attributeName: {
              localPart: 'saltValue'
            },
            type: 'attribute'
          }, {
            name: 'spinCount',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'spinCount'
            },
            type: 'attribute'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STPasswordKeyEncryptorUri',
        baseTypeInfo: 'Token',
        values: ['http:\/\/schemas.microsoft.com\/office\/2006\/keyEncryptor\/password']
      }],
    elementInfos: [{
        typeInfo: '.CTPasswordKeyEncryptor',
        elementName: 'encryptedKey'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password: org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password_Module = org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password = org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password_Module.org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password = org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password_Module.org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password;
  }
}