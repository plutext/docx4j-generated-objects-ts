// The README's examples, compile-checked (npm run typecheck). Office JS is stubbed minimally.
import { unmarshalString, marshalString, unmarshalPackage, marshalPackage, unwrap, deepCopy } from '../src/index.mjs';
import type { DocumentElement, P, R, Text, Tc } from '../modules/org_docx4j_wml';

declare namespace Word {
  interface ClientResult<T> { value: T; }
  interface RequestContext { sync(): Promise<void>; }
  interface Body { context: RequestContext; getOoxml(): ClientResult<string>; insertOoxml(ooxml: string, location: 'Replace' | 'Start' | 'End'): void; }
  interface CustomXmlPart { context: RequestContext; getXml(): ClientResult<string>; setXml(xml: string): void; }
}

export async function shout(body: Word.Body): Promise<void> {
  const ooxml = body.getOoxml();
  await body.context.sync();
  const pkg = await unmarshalPackage(ooxml.value);
  const part = unwrap(pkg).part?.find((p) => p.name === '/word/document.xml');
  const document = part?.xmlData?.any as DocumentElement | undefined;
  if (!document) return;
  for (const entry of unwrap(document).body?.content ?? []) {
    if (entry.value.TYPE_NAME !== 'org_docx4j_wml.P') continue;
    const paragraph: P = entry.value;
    for (const run of paragraph.content ?? []) {
      if (run.value.TYPE_NAME !== 'org_docx4j_wml.R') continue;
      for (const item of (run.value as R).content ?? []) {
        if (item.value.TYPE_NAME === 'org_docx4j_wml.Text') {
          const text = item.value as Text;
          text.value = text.value?.toUpperCase();
        }
      }
    }
  }
  body.insertOoxml(await marshalPackage(pkg), 'Replace');
  await body.context.sync();
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const paragraph: P = {
  pPr: { pStyle: { val: 'Heading1' } },
  content: [{
    name: { namespaceURI: W, localPart: 'r' },
    value: { rPr: { b: {} }, content: [{ name: { namespaceURI: W, localPart: 't' }, value: { value: 'Hello' } }] },
  }],
};
// @ts-expect-error no such property on run properties
export const wrong: P = { content: [{ name: { namespaceURI: W, localPart: 'r' }, value: { rPr: { bold: {} } } }] };
export const paragraphXml = (): Promise<string> => marshalString({ name: { namespaceURI: W, localPart: 'p' }, value: paragraph });

export async function navigate(documentXml: string): Promise<void> {
  const doc = unwrap(await unmarshalString<DocumentElement>(documentXml));
  const firstParagraph = doc.body?.content?.find((e) => e.value.TYPE_NAME === 'org_docx4j_wml.P')?.value as P | undefined;
  const container = firstParagraph?.PARENT;
  if (container?.TYPE_NAME === 'org_docx4j_wml.Tc') {
    const cell: Tc = container;
    void cell;
  }
  const copy = deepCopy(firstParagraph);
  void copy;
}

export async function readPart(part: Word.CustomXmlPart): Promise<void> {
  const xml = part.getXml();
  await part.context.sync();
  const element = await unmarshalString(xml.value);
  part.setXml(await marshalString(element));
}
