# CR-007: Per-call runtime options on the facade, and what unmarshalling dropped

**Status:** Proposed 2026-09-27
**Depends on:** `@docx4j/jsonix` 3.4.0 (jsonix-CR-006, implemented there as `fc44f0c`, unreleased)
**Requested by:** `plutext/docx4j-core-ts` CR-002 section 27 and its CR-006, relayed by the jsonix
session 2026-09-27; and CR-004's fidelity test here, which is a second consumer (section 4)
**Counterpart:** docx4j's `JaxbValidationEventHandler` and the `Unmarshaller`/`Marshaller` a part
is read and written with, which docx4j configures per call

## 1. Summary

The facade's four entry points take no options: `unmarshalString`, `unmarshalNode`, `marshalString`
and `marshalNode` each build their unmarshaller or marshaller from the shared context and offer the
caller no way to say anything about that one call. Everything is decided when the context is built,
which is right for the prefix table and wrong for a question that belongs to one document - *what
did you drop reading this part?*

`@docx4j/jsonix` 3.4.0 adds that: `createUnmarshaller(options)` and `createMarshaller(options)`,
copied onto a derived context rather than the shared one. Recognised for unmarshalling,
`onUnexpectedElement(name, node, classInfo)`, `onUnexpectedAttribute(name, node, classInfo)`,
`parentPointers`, `supportXsiType`; for marshalling, `onElement(element, value)`,
`namespacePrefixes`, `supportXsiType`. Errors arrive located (`t:style/@size (line 2): ...`), typed
as `Jsonix.LocatedError`.

This CR forwards those options through the facade.

## 2. What changes

```ts
export async function unmarshalString<E>(xml: string, options?: Jsonix.UnmarshallerOptions): Promise<E>;
export async function unmarshalNode<E>(node: Node, options?: Jsonix.UnmarshallerOptions): Promise<E>;
export async function marshalString(element: TypedNamedValue, options?: Jsonix.MarshallerOptions): Promise<string>;
export async function marshalNode(element: TypedNamedValue, options?: Jsonix.MarshallerOptions): Promise<Element>;
```

All four, not the two core-ts asked for: a facade whose node-level functions take options and whose
string-level ones do not would be a shape nobody could remember, and the string functions are the
ones this package's own tests and most consumers use.

The options are per call and reach the runtime through its derived context, so the shared context
is untouched and `resetContext()` is not involved - the opposite of `getContext`'s options, which
are read once when the context is built. That difference is the thing to document, since the two
kinds of option now sit next to each other in the README.

## 3. What it replaces here

`marshalToDocument` derives a context by hand for the per-root prefix table:

```ts
const derived = Object.create(context, { namespacePrefixes: { value: namespacePrefixesFor(context, element) } });
```

3.4.0's `namespacePrefixes` marshaller option is that, supported: the hand-rolled `Object.create`
goes, and `namespacePrefixesFor` stays, since choosing the table per root (CR-001) is this
package's decision, not the runtime's. Note the caveat the jsonix session gives: a QName
unmarshalled from input keeps its own prefix on marshal, so the table only governs names that
carry none - which is what the derivation does today, so nothing changes in behaviour.

Not to be confused with jsonix-CR-003 part 2, still deferred there, which would retire
`fixRootNamespaceDeclarations` - the pass that strips declarations the tree does not use. That is a
different promise and this CR does not deliver it.

## 4. The second consumer: the fidelity test

CR-004's phase B triage worked out what the model had dropped by **inference** - marshal the part
back, compare canonically, and reason from the diff to a missing attribute declaration. 43 lost
`uid` attributes were found that way, and the reasoning was only as good as the canonicaliser.

`onUnexpectedElement` and `onUnexpectedAttribute` report the same losses **directly**, at the point
the unmarshaller drops them, with the DOM node they stood on. That is a better instrument than a
diff: it names the loss instead of leaving it to be deduced, it catches a loss the canonicaliser
might call equal, and it would have found the `xr3:uid` on `tableColumn` that the first-difference
comparison masked. Worth using in `test/fidelity.mjs` once the option exists, as a cross-check that
the two methods agree - a difference the diff finds and the callback does not, or the reverse, is
itself a finding about the canonicaliser.

This is not a reason to hurry the CR; it is a reason to do it properly when the runtime ships.

## 5. Tests

- `test/smoke.mjs`: an `onUnexpectedElement` callback fires for an element the model does not
  admit, with the element name; `onElement` collects the elements a marshal writes; a
  `namespacePrefixes` option overrides for one call and leaves the next call's output unchanged
  (the shared context is not modified).
- The existing CR-001 prefix tests must pass unchanged when `marshalToDocument` stops deriving by
  hand: they are the check that the replacement is behaviour-preserving.
- `test/nodenext/consumer.mts` gains the new option types.
- `test/fidelity.mjs` (section 4) once the callbacks exist.

## 6. Open questions

1. Should the fidelity test **fail** on any `onUnexpectedElement` that its diff did not also find,
   or only report? Failing is stricter and likely right, on the argument of CR-004 section 7.2, but
   it cannot be decided until the callback has been run over the corpus once.
2. `parentPointers` as a per-call option: the facade sets it true when building the context
   (docx4j's model has parents), and a caller turning it off for one unmarshal would get objects
   whose `PARENT` is absent where the declarations say it may be present. Worth exposing, or worth
   refusing? Refusing is a facade decision, not a runtime one.
3. Does `deepCopyAs` want the same treatment? It marshals and unmarshals internally; a caller
   cannot reach those calls today, and it is not obvious anyone wants to.
