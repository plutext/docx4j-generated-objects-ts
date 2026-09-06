var org_docx4j_com_microsoft_schemas_office_x2006_encryption_Module_Factory = function () {
  var org_docx4j_com_microsoft_schemas_office_x2006_encryption = {
    name: 'org_docx4j_com_microsoft_schemas_office_x2006_encryption',
    defaultElementNamespaceURI: 'http:\/\/schemas.microsoft.com\/office\/2006\/encryption',
    dependencies: ['org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate', 'org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password'],
    typeInfos: [{
        localName: 'CTDataIntegrity',
        typeName: 'CT_DataIntegrity',
        propertyInfos: [{
            name: 'encryptedHmacKey',
            required: true,
            typeInfo: 'Base64Binary',
            attributeName: {
              localPart: 'encryptedHmacKey'
            },
            type: 'attribute'
          }, {
            name: 'encryptedHmacValue',
            required: true,
            typeInfo: 'Base64Binary',
            attributeName: {
              localPart: 'encryptedHmacValue'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTEncryption',
        typeName: 'CT_Encryption',
        propertyInfos: [{
            name: 'keyData',
            required: true,
            typeInfo: '.CTKeyData'
          }, {
            name: 'dataIntegrity',
            required: true,
            typeInfo: '.CTDataIntegrity'
          }, {
            name: 'keyEncryptors',
            required: true,
            typeInfo: '.CTKeyEncryptors'
          }]
      }, {
        localName: 'CTKeyData',
        typeName: 'CT_KeyData',
        propertyInfos: [{
            name: 'saltSize',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'saltSize'
            },
            type: 'attribute'
          }, {
            name: 'blockSize',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'blockSize'
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
            name: 'hashSize',
            required: true,
            typeInfo: 'UnsignedInt',
            attributeName: {
              localPart: 'hashSize'
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
            typeInfo: '.STCipherChaining',
            attributeName: {
              localPart: 'cipherChaining'
            },
            type: 'attribute'
          }, {
            name: 'hashAlgorithm',
            required: true,
            typeInfo: '.STHashAlgorithm',
            attributeName: {
              localPart: 'hashAlgorithm'
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
          }]
      }, {
        localName: 'CTKeyEncryptor',
        typeName: 'CT_KeyEncryptor',
        propertyInfos: [{
            name: 'encryptedPasswordKey',
            required: true,
            elementName: {
              localPart: 'encryptedKey',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/2006\/keyEncryptor\/password'
            },
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_password.CTPasswordKeyEncryptor'
          }, {
            name: 'encryptedCertificateKey',
            required: true,
            elementName: {
              localPart: 'encryptedKey',
              namespaceURI: 'http:\/\/schemas.microsoft.com\/office\/2006\/keyEncryptor\/certificate'
            },
            typeInfo: 'org_docx4j_com_microsoft_schemas_office_x2006_keyEncryptor_certificate.CTCertificateKeyEncryptor'
          }, {
            name: 'uri',
            values: ['http:\/\/schemas.microsoft.com\/office\/2006\/keyEncryptor\/password', 'http:\/\/schemas.microsoft.com\/office\/2006\/keyEncryptor\/certificate'],
            attributeName: {
              localPart: 'uri'
            },
            type: 'attribute'
          }]
      }, {
        localName: 'CTKeyEncryptors',
        typeName: 'CT_KeyEncryptors',
        propertyInfos: [{
            name: 'keyEncryptor',
            required: true,
            collection: true,
            typeInfo: '.CTKeyEncryptor'
          }]
      }, {
        type: 'enumInfo',
        localName: 'STCipherChaining',
        baseTypeInfo: 'Token',
        values: ['ChainingModeCBC', 'ChainingModeCFB']
      }, {
        type: 'enumInfo',
        localName: 'STHashAlgorithm',
        baseTypeInfo: 'Token',
        values: ['SHA1', 'SHA256', 'SHA384', 'SHA512', 'MD5', 'MD4', 'MD2', 'RIPEMD-128', 'RIPEMD-160', 'WHIRLPOOL']
      }],
    elementInfos: [{
        typeInfo: '.CTEncryption',
        elementName: 'encryption'
      }]
  };
  return {
    org_docx4j_com_microsoft_schemas_office_x2006_encryption: org_docx4j_com_microsoft_schemas_office_x2006_encryption
  };
};
if (typeof define === 'function' && define.amd) {
  define([], org_docx4j_com_microsoft_schemas_office_x2006_encryption_Module_Factory);
}
else {
  var org_docx4j_com_microsoft_schemas_office_x2006_encryption_Module = org_docx4j_com_microsoft_schemas_office_x2006_encryption_Module_Factory();
  if (typeof module !== 'undefined' && module.exports) {
    module.exports.org_docx4j_com_microsoft_schemas_office_x2006_encryption = org_docx4j_com_microsoft_schemas_office_x2006_encryption_Module.org_docx4j_com_microsoft_schemas_office_x2006_encryption;
  }
  else {
    var org_docx4j_com_microsoft_schemas_office_x2006_encryption = org_docx4j_com_microsoft_schemas_office_x2006_encryption_Module.org_docx4j_com_microsoft_schemas_office_x2006_encryption;
  }
}