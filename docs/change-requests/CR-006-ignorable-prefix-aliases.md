# CR-006: An `mc:Ignorable` prefix the preferred-prefix table cannot produce

**Status:** Implemented 2026-09-24
**Depends on:** CR-001 (`NAMESPACE_PREFIXES` and `fixRootNamespaceDeclarations`)
**Requested by:** CR-004 phase A's fidelity test, which found it on 2026-09-23 over Excel's own
slicer cache
**Counterpart:** docx4j `NamespacePrefixMapperUtils.getPreDeclaredNamespaceUris2` with
`NamespacePrefixMappings.getNamespaceURIStatic`, as its CR-024 left them

## 1. Summary

A re-marshalled slicer cache tells a reader to ignore one prefix fewer than Excel did.

`xl/slicerCaches/slicerCache1.xml` has an `x14:slicerCacheDefinition` root carrying
`mc:Ignorable="x xr10"`, where Excel binds `x` to the SpreadsheetML main namespace. Marshalled back
through the facade it comes out as `mc:Ignorable="xr10"`, with a warning on the console:

```
marshal: mc:Ignorable names x, which no declaration or NAMESPACE_PREFIXES entry resolves; dropped
```

Nothing else in the part changes; this is the only difference the fidelity test finds in it.

## 2. Why

CR-001 made `fixRootNamespaceDeclarations` declare every prefix `mc:Ignorable` names, because Word
and Excel repair a file whose `mc:Ignorable` names an undeclared prefix. It resolves a prefix by
inverting `NAMESPACE_PREFIXES`:

```ts
const namespaceFor = new Map(Object.entries(table).map(([namespaceURI, prefix]) => [prefix, namespaceURI]));
```

That inversion is total for every namespace with a prefix, and empty for the one namespace this
package writes as the **default**:

```ts
'http://schemas.openxmlformats.org/spreadsheetml/2006/main': '',   // SML uses the default namespace, as docx4j
```

So the inverted map holds `'' -> spreadsheetml/2006/main` and no entry for `x`. The prefix is
unresolvable, and the facade drops the token rather than write an `mc:Ignorable` naming a prefix it
has not declared - which, on its own terms, is the right call between those two.

The third option is the one Excel takes and the one docx4j took in its CR-024: **declare the
prefix as well**. A namespace may be declared twice, once as the default and once with a prefix;
`xmlns="...spreadsheetml/2006/main"` and `xmlns:x="...spreadsheetml/2006/main"` on the same root
are both in scope, and `mc:Ignorable="x xr10"` then resolves. docx4j's
`getPreDeclaredNamespaceUris2` walks the `mc:Ignorable` tokens and pre-declares a pair for each,
resolving the prefix through `getNamespaceURIStatic`, which knows `x` for exactly this reason:

```java
// Excel binds "x" to the SpreadsheetML main namespace on its slicer, slicer cache and
// timeline parts (whose roots are x14/x15) and names it in mc:Ignorable="x xr10";
// docx4j writes that namespace as the default, so the prefix is declared beside it
// through NamespacePrefixMapperUtils.getPreDeclaredNamespaceUris2 (CR-024).
if (prefix.equals("x"))
    return "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
```

## 3. What changes

`src/index.mts` only; no generated output, no new public path.

1. **An alias table**, exported beside `NAMESPACE_PREFIXES` so a caller can see and extend it:

   ```ts
   /** Prefixes Office names in mc:Ignorable that NAMESPACE_PREFIXES cannot produce, because the
    *  namespace is written as the default. docx4j's NamespacePrefixMappings.getNamespaceURIStatic. */
   export const IGNORABLE_PREFIX_ALIASES: Record<string, string> = {
     x: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
   };
   ```

   One entry, because one namespace is written as the default. The table is keyed prefix to
   namespace (the opposite of `NAMESPACE_PREFIXES`), since that is the direction this question is
   asked in.

2. **`fixRootNamespaceDeclarations` consults it** after the inverted table and before giving up:
   `namespaceFor.get(prefix) ?? IGNORABLE_PREFIX_ALIASES[prefix]`. A prefix resolved this way is
   declared like any other, so the root gains `xmlns:x="...spreadsheetml/2006/main"` and keeps
   `mc:Ignorable="x xr10"`. Where the root is itself in the SpreadsheetML main namespace the
   declaration sits beside the default one, as docx4j's pair does; where it is an x14 or x15 root,
   as a slicer cache is, there is no default declaration to sit beside and `xmlns:x` simply joins
   the others. The warning and the dropping stay for a
   prefix neither table knows: that case is still a file Office would repair, and silence would
   hide it.

3. **The alias table follows the options table.** `marshalString` and friends take a
   `namespacePrefixes` option (CR-001); a caller that passes its own table still gets the aliases,
   since the aliases answer a different question (what a prefix in `mc:Ignorable` means) from the
   one the table answers (which prefix to write for a namespace). If that proves wrong, the option
   can carry aliases later; there is no consumer for it now.

## 4. Tests

- `test/smoke.mjs`: an `x14:slicerCacheDefinition` root with `mc:Ignorable="x xr10"` marshals with
  `xmlns:x` present and `mc:Ignorable` intact; and the existing check
  that an unknown prefix (`zz`) is still dropped with the warning keeps its behaviour.
- `test/fidelity.mjs`: the `KNOWN` entry for
  `cr022-slicers-timelines.xlsx/xl/slicerCaches/slicerCache1.xml` is removed, and the part joins the
  16 that round-trip identically. The table is strict in both directions, so leaving the entry in
  place would itself fail the test - which is how this CR closes.

## 5. Consumers

`@docx4j/core-ts` writes slicer, slicer cache and timeline parts through this facade (its CR-004
phase B), so it gets the fix by upgrading. Nothing to change there.

## 6. What this is not

Not a general aliasing mechanism for reading. The model never sees prefixes - it resolves
namespaces at unmarshal - and `mcBranchOf` (CR-003 section 3.8) judges `Requires` by prefix
against `NAMESPACE_PREFIXES` for a documented reason of its own. This CR touches one question on
the marshal side: what a prefix in a root's `mc:Ignorable` must be declared as.

## 7. As implemented (2026-09-24)

As proposed, in `src/index.mts`: `IGNORABLE_PREFIX_ALIASES` exported beside `NAMESPACE_PREFIXES`
with its one entry, and `fixRootNamespaceDeclarations` resolving an `mc:Ignorable` token through
`namespaceFor.get(prefix) ?? IGNORABLE_PREFIX_ALIASES[prefix]`. The warning and the dropping stay
for a prefix neither table knows, which the smoke still pins (`zz`).

Excel's `slicerCache1.xml` now round-trips: the root keeps `mc:Ignorable="x xr10"` and carries
`xmlns:x="...spreadsheetml/2006/main"`. It writes the root's own namespace with the table's prefix
(`<x14:slicerCacheDefinition xmlns:x14="...">`) where Excel writes it as the default - a prefix
choice under CR-001, not a difference in content, which the fidelity canonicaliser compares as
equal. A first draft of the smoke asserted the root kept a default declaration; that assertion was
wrong about this package's own rules, and the corrected one pins the prefixed form.

The `KNOWN` entry in `test/fidelity.mjs` is gone, and with it the third of the three differences
phase A found. The test's insistence - it fails while an entry records a difference the part no
longer has - is what closed this CR rather than leaving the entry to rot.

## 8. Open questions

1. Should the aliases instead be folded into `NAMESPACE_PREFIXES` as a second entry for the same
   namespace? No: that table is namespace-keyed and one namespace has one preferred prefix, which
   is what makes it a faithful copy of docx4j's `NamespacePrefixMappings`. A separate table says
   what it means.
2. Are there other prefixes Office names in `mc:Ignorable` that the table cannot produce? Only one
   namespace is written as the default today, so only `x` can arise. Phase B of CR-004 (every part
   of the fixture documents) is the measurement that would show another, and it is the reason to
   do phase B before assuming this is complete.
