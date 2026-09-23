// Compile-only: the built package as a Node ES module consumer sees it (module/moduleResolution nodenext),
// through the package's own name and exports map. The repository's tsconfigs use moduleResolution bundler,
// which accepts extensionless relative imports in declarations; nodenext rejects them (TS2835), and with
// skipLibCheck they silently degrade to any. Run by npm test after the build.
import { unmarshalString, marshalString, unwrap, NAMESPACE_PREFIXES, IGNORABLE_PREFIX_ALIASES, MODULES, MODULE_NAMES, modulesFor, getContext } from '@docx4j/generated-objects-ts';
import type { DocumentElement, P } from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { createP, createPElement } from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import { p, r, textOf, find } from '@docx4j/generated-objects-ts/builders/wml';
import { isCustomStyle, highlightHexValue } from '@docx4j/generated-objects-ts/helpers/wml';

export async function consume(xml: string): Promise<string> {
  const doc = unwrap(await unmarshalString<DocumentElement>(xml));
  const paragraphs: P[] = find<P>(doc, 'org_docx4j_wml.P');
  const built: P = p([r('a', { bold: true })]).value;
  const same: P = el.p({ content: [] }).value;
  void getContext({ modules: modulesFor('org_docx4j_wml') });
  void [NAMESPACE_PREFIXES, IGNORABLE_PREFIX_ALIASES, MODULES.org_docx4j_wml, MODULE_NAMES[0], paragraphs, built, same, isCustomStyle({ customStyle: undefined }), highlightHexValue('yellow')];
  return textOf(doc) + (await marshalString(createPElement(createP())));
}

// The builders' and helpers' types must be real, not any: each line below is an error only while the module
// types resolve, so with skipLibCheck and a degraded import the directive goes unused and the check fails.
// @ts-expect-error p() returns Element<P>, whose value is not a number
export const notANumber: number = p('x').value;
// @ts-expect-error a highlight's val is a colour name, not a number (with Highlight as any, val would be any)
highlightHexValue({ val: 1 });
