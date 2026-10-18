// The one-line consumer of CR-005 section 3: a bundler must make this work with no configuration.
// Imported by package name, as a consumer writes it (package self-reference resolves the exports map).
import { unmarshalString, marshalString, unwrap } from '@docx4j/generated-objects-ts';

const doc = unwrap(await unmarshalString(
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>'));

// An ES-module bundle runs the CommonJS runtime inside it in strict mode, where @docx4j/jsonix
// before 3.4.2 threw on every xsd:base64Binary value (Base64Binary.decode assigned an undeclared
// variable); the editor met it on a VML Fallback's o:gfxdata. Node's own import never shows it,
// since it loads the runtime as CommonJS, so this is the only test here that can.
const pict = await unmarshalString('<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  + ' xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape id="s1" o:gfxdata="AQID"/></w:pict>');
const shape = unwrap(unwrap(pict).anyAndAny[0]);
console.log(`${doc.TYPE_NAME} ${doc.body.content.length} ${shape.gfxdata} ${(await marshalString(pict)).includes('o:gfxdata="AQID"')}`);
