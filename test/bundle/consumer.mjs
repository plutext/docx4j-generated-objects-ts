// The one-line consumer of CR-005 section 3: a bundler must make this work with no configuration.
// Imported by package name, as a consumer writes it (package self-reference resolves the exports map).
import { unmarshalString, unwrap } from '@docx4j/generated-objects-ts';

const doc = unwrap(await unmarshalString(
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>'));
console.log(`${doc.TYPE_NAME} ${doc.body.content.length}`);
