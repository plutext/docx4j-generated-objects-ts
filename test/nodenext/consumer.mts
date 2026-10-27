// Compile-only: the built package as a Node ES module consumer sees it (module/moduleResolution nodenext),
// through the package's own name and exports map. The repository's tsconfigs use moduleResolution bundler,
// which accepts extensionless relative imports in declarations; nodenext rejects them (TS2835), and with
// skipLibCheck they silently degrade to any. Run by npm test after the build.
import { unmarshalString, marshalString, unwrap, NAMESPACE_PREFIXES, IGNORABLE_PREFIX_ALIASES, IGNORABLE_EXTENSION_NAMESPACES, MODULES, MODULE_NAMES, modulesFor, getContext, WML_GLOBAL_ELEMENTS, org_docx4j_wml_globals, wordInvalidValue, setNormaliseOnSave } from '@docx4j/generated-objects-ts';
import type { WordReading, Normalisation, OutOfRange } from '@docx4j/generated-objects-ts';
import type { DocumentElement, P } from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { createP, createPElement } from '@docx4j/generated-objects-ts/factory/org_docx4j_wml';
import * as el from '@docx4j/generated-objects-ts/el/org_docx4j_wml';
import { p, r, textOf, find } from '@docx4j/generated-objects-ts/builders/wml';
import { isCustomStyle, highlightHexValue, wordColor, WORD_COLOR_NAMES, WORD_INTEGER_STORAGE, wordIntegerReading, type WordIntegerReading, type WordIntegerStorage } from '@docx4j/generated-objects-ts/helpers/wml';

export async function consume(xml: string): Promise<string> {
  const doc = unwrap(await unmarshalString<DocumentElement>(xml));
  const paragraphs: P[] = find<P>(doc, 'org_docx4j_wml.P');
  const built: P = p([r('a', { bold: true })]).value;
  const same: P = el.p({ content: [] }).value;
  void getContext({ modules: modulesFor('org_docx4j_wml') });
  // CR-007: per-call options, typed by the runtime's own interfaces.
  void unmarshalString('<w:p/>', { onUnexpectedAttribute: (name) => void name.localPart });
  // CR-013: the facade's onWordReading beside the runtime's options, and its onInvalidValue to compose.
  const readings: WordReading[] = [];
  // CR-015: an integer's kinds and its numeric value.
  const integerKinds: WordReading['kind'][] = ['value.truncated', 'value.refused'];
  const truncatedTo: WordReading['value'] = 259;
  void unmarshalString('<w:p/>', { onWordReading: (reading) => readings.push(reading), onInvalidValue: wordInvalidValue });
  // CR-015 phase 2: jsonix-CR-012's strictIntegers, through the facade's options.
  void unmarshalString('<w:p/>', { strictIntegers: false });
  // CR-016: onOutOfRange on both sides, and the table and its reading.
  const ranges: OutOfRange[] = [];
  void unmarshalString('<w:p/>', { onOutOfRange: (r) => ranges.push(r) });
  void marshalString(createPElement(createP()), { onOutOfRange: (r) => void r.path });
  // The entry on the report (2026-10-27): the key pair finds it back in the table, and the entry is the table's type.
  const entry: WordIntegerStorage | undefined = ranges[0] && WORD_INTEGER_STORAGE[ranges[0].typeName]?.[ranges[0].property];
  const reported: WordIntegerStorage | undefined = ranges[0]?.storage;
  const rule: string | undefined = ranges[0]?.rule;
  void [entry, reported, rule];
  const reading: WordIntegerReading | undefined = wordIntegerReading('org_docx4j_wml.PPrBase.Ind', 'left', 7205001);
  const measured: boolean | undefined = WORD_INTEGER_STORAGE['org_docx4j_wml.PPrBase.Ind']?.['left']?.measured;
  const color: string | undefined = wordColor(WORD_COLOR_NAMES.red);
  void marshalString({ name: { namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', localPart: 'p' }, value: paragraphs[0]! }, { onElement: (element, value) => void [element, value] });
  // CR-013 phase C: the normaliser's options beside the runtime's, and its process-wide switch.
  const normalised: Normalisation[] = [];
  void marshalString(createPElement(createP()), { normaliseOnSave: true, onNormalised: (n) => normalised.push(n), omitElement: () => undefined });
  setNormaliseOnSave(true);
  // CR-008: the extension namespaces, a readonly list of URIs.
  const extensions: readonly string[] = IGNORABLE_EXTENSION_NAMESPACES;
  void [NAMESPACE_PREFIXES, IGNORABLE_PREFIX_ALIASES, extensions, MODULES.org_docx4j_wml, MODULE_NAMES[0], paragraphs, built, same, isCustomStyle({ customStyle: undefined }), highlightHexValue('yellow'), color, readings, normalised, integerKinds, truncatedTo, ranges, reading, measured, WML_GLOBAL_ELEMENTS.drawing, org_docx4j_wml_globals.name];
  return textOf(doc) + (await marshalString(createPElement(createP())));
}

// The builders' and helpers' types must be real, not any: each line below is an error only while the module
// types resolve, so with skipLibCheck and a degraded import the directive goes unused and the check fails.
// @ts-expect-error p() returns Element<P>, whose value is not a number
export const notANumber: number = p('x').value;
// @ts-expect-error a highlight's val is a colour name, not a number (with Highlight as any, val would be any)
highlightHexValue({ val: 1 });
// @ts-expect-error a reading's kind is one of four (with WordReading as any, anything would do)
export const notAKind: WordReading['kind'] = 'value.dropped';
// @ts-expect-error a range report's value is a number (with OutOfRange as any, a string would do)
export const notAValue: OutOfRange['value'] = '7205001';
// @ts-expect-error a normalisation's after is a string or null (with Normalisation as any, a number would do)
export const notAfter: Normalisation['after'] = 1;
