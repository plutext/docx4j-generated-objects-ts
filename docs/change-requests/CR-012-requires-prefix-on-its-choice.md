# CR-012: Declare a source-only Requires prefix on its mc:Choice, as Word does

**Status:** Implemented 2026-10-20, not yet released; the placement decided 2026-10-20 (Jason: "(c) placement", for both ports); a Choice's own `mc:Ignorable` and `mc:MustUnderstand` added 2026-10-24 (Jason: "D1 (b): yes")
**Requested by:** Jason, through core-ts, which relayed the decision and measured Word's placement
(its CR-010 section 14.9, core-ts 2621f89).
**Depends on:** CR-011 (released in 0.4.3), whose records this places. Nothing new in the runtime.
**Counterpart:** docx4j CR-052, whose D4 the same decision settles.

## 1. Summary

CR-011 declares a prefix only the source bound from the source's binding, and declares it on the
part root. For an `mc:Ignorable` prefix the root is the only place. For a `Requires` prefix, Word
puts it on the `mc:Choice` that names it. Jason chose Word's placement for both ports:
- a `Requires` prefix the table or the aliases give is declared on the root, as today;
- any other `Requires` prefix is declared on the `mc:Choice` that names it, with the source's binding;
- an `mc:Ignorable` prefix stays on the root;
- nothing is copied onto `mc:Fallback`;
- chart parts follow the same rule.

A Choice's own `mc:Ignorable` and `mc:MustUnderstand` follow its `Requires`, as docx4j's do (its
CR-052, D1 (b); Jason, 2026-10-24). Until now their prefixes weren't declared at all.

This changes output only for parts with a source-only prefix that a Choice names, or with a
Choice's own `mc:Ignorable` or `mc:MustUnderstand`. The fidelity corpus has neither, and its 341
parts marshal byte-identically to 0.4.3.

## 2. Measured

core-ts counted where Word declares, over 2,728 Word 365 saves (24,075 WordprocessingML parts,
2026-10-20). A save counts when `app.xml` says Word 16 and `w:document` declares `cx8` and `w16du`.
- A `Requires` prefix in Word's fixed root set is declared on the root: 664 of 664 (`wps` 432, `cx1`
  140, `wpg` 73, `cx2` 10, `v` 5, `wpi` 2, `wpc` 1, `wp14` 1).
- A prefix outside that set goes on the Choice that names it: `zz9`, 3 of 3. These are Word's
  resaves (build 16.0.20326) of docx4j's fallback-rule probes, whose sources declare `zz9` on
  `mc:AlternateContent`. Word writes it ahead of `Requires`: `<mc:Choice
  xmlns:zz9="http://example.com/zz9" ... Requires="zz9">`. It also copies it onto a sibling
  `mc:Fallback`, and doesn't add `zz9` to `mc:Ignorable`.
- In a WordprocessingML part, Word declares nothing on `mc:AlternateContent`.
- In chart parts, Word declares `xmlns:mc` on `mc:AlternateContent` and `c14` on the Choice (37 of
  37), as Excel and PowerPoint do. `c14` is in the table, so the rule keeps it on the root.

**Word reads all three places alike** (core-ts check 50, Jason's run in Word 365, 2026-10-20). The
file has five runs, each with a text box as its Choice and plain text as its Fallback:
- case 1 is Word's own `Requires="wps"`, declared on the root;
- cases 2 to 4 bind the `wps` namespace to another prefix, used on the shape elements too, and
  declare it on the root, on `mc:AlternateContent` and on `mc:Choice` respectively;
- case 5 requires `zz9`.

Word showed the text box in cases 1 to 4 and the Fallback in case 5. So 0.4.3's root placement is
read by Word, and this CR is about matching Word's output. On save, Word renamed the three aliases
to `wps` on the root, and resolved case 5. Each kept `mc:Fallback` carried unused copies of the
source's declarations outside Word's fixed set.

## 3. Implemented

In `fixRootNamespaceDeclarations`, for each prefix a Choice names that no declaration below the
root already answers, by its `Requires`, its own `mc:Ignorable` or its `mc:MustUnderstand`:
1. **A source binding**, from the Choice's record or else the root's, is declared on that Choice,
   ahead of its `Requires` and its other attributes, as Word writes it. The source binding is the one the table and the
   aliases wouldn't give (CR-011 section 3.1). So `shape2010` for `wps`, and `w14` bound to another
   namespace, go on the Choice too.
2. **If the marshalled root declares that same binding,** the Choice doesn't repeat it. The root
   keeps a source-only binding only for its `mc:Ignorable`, or for the tree's own use. In the
   editor's shape, the root's `xmlns:zz9` serves its `mc:Ignorable` and the Choice's `Requires`
   alike, so it is written once. A `zz9` the source root declared that `mc:Ignorable` doesn't name
   moves to the Choice. Jason's rule said nothing of a binding both places need. He confirmed one
   declaration on the root, for both ports, on 2026-10-20 ("one declaration on the root right").
   Word's placement for such a prefix isn't measured: check 50's case 5 was resolved by Word, so
   its Choice wasn't written.
3. **Otherwise, as before:** the table or the aliases, on the root. A prefix nothing resolves:
   - in `Requires`, is kept, with the existing warning;
   - in `mc:MustUnderstand`, is kept, with the same warning naming that attribute. Dropping it
     would let a reader that doesn't understand the namespace go on where it must refuse;
   - in the Choice's own `mc:Ignorable`, is dropped from it, with a warning, as one in the root's
     `mc:Ignorable` is (CR-001). An undeclared prefix there identifies no namespace, and Office
     repairs a root that names one.

**A Choice's own `mc:Ignorable` and `mc:MustUnderstand`** (added 2026-10-24, Jason: "D1 (b): yes",
for parity with docx4j). The unmarshal records their source-only bindings with the `Requires`
ones, on the Choice's object, so they travel and copy the same way. On docx4j's
`aChoicesOwnIgnorableAndMustUnderstandAreDeclared` input, `<mc:Choice Requires="wps"
mc:Ignorable="zz9" mc:MustUnderstand="w14">` under an `mc:AlternateContent` declaring all three,
`wps` and `w14` go on the root and `zz9` on the Choice, as in docx4j. Before, this package wrote
neither `zz9` nor `w14`, and stripped the table's `w14` from the root, with no warning.

A Choice's declarations now go ahead of all its attributes, not only `Requires`. A typed Choice
writes `Requires`, `mc:Ignorable` and `mc:MustUnderstand` in that order, and moving only `Requires`
behind a declaration would have reordered them.

The root's own rules don't change. A declaration kept only because `mc:Ignorable` names its prefix
still gives way to the root's source binding. One a `Requires` named, from the table, now goes,
since the Choice declares the source's.

**Warnings narrow:**
- **The tree's namespace wins** now applies only where the Choice's own content writes the prefix
  for another namespace, since a declaration there would rebind that content, and to a root's
  `mc:Ignorable`. A use elsewhere in the tree no longer stops a Choice declaring its own binding.
  Its text now says "an mc:Ignorable, Requires or mc:MustUnderstand naming it now means the latter".
- **One prefix bound two ways** is gone. Each Choice declares its own binding, and a Choice's
  binding and the root's `mc:Ignorable` binding are both kept.

The cost is one walk of a Choice's content, only for a Choice with a source binding, and only once
a CR-011 record exists.

## 4. Not proposed

- **Renaming to the table's prefix,** as Word renamed `shape2010` to `wps`. That changes the
  `Requires` value, and Jason's rule keeps the source's binding.
- **Copies on `mc:Fallback`.** Word's are unused, and the rule says no.
- **A chart-part rule.** Word declares `c14` on the Choice there, but `c14` is in the table, so it
  stays on the root (Jason: "chart parts follow the same rule").
- **`ProcessContent`, and the own `mc:Ignorable` and `mc:MustUnderstand` of `mc:AlternateContent`
  and `mc:Fallback`:** not declared. docx4j doesn't declare them either: its CR-052 section 4 keeps
  `mc:AlternateContent`'s outside D1 (b). No part in the corpus carries any of them.

## 5. Tests

The smoke's CR-011 block asserts where each declaration goes:
- a `Requires` prefix only the source bound is on its Choice, ahead of `Requires`, and not on the
  root. This covers `zz9` and `shape2010` from `mc:AlternateContent`, content moved between trees, a
  `wml()` fragment, a paragraph marshalled alone, `deepCopy`, a flat OPC part, and the caller's
  `onElement` case;
- the editor's shape declares `zz9` once, on the root; a root `xmlns:zz9` that `mc:Ignorable`
  doesn't name moves to the Choice;
- the superseded table `w14` leaves the root, and the source's `w14` is on the Choice;
- the tree's `w14:paraId` outside the Choice keeps the root's `w14`, and the Choice declares the
  source's, with no warning;
- **new:** a Choice whose own content writes `w14:paraId` keeps the tree's `w14`, with the warning;
- two Choices binding `zz9` two ways keep both, and so do a root `mc:Ignorable` and a Choice, with
  no warning;
- a Choice's own `mc:Ignorable` and `mc:MustUnderstand`, on docx4j's input: `wps` and `w14` on the
  root, `zz9` on the Choice ahead of its attributes, which keep their order, and on the Choice when
  its paragraph is marshalled alone;
- an `mc:MustUnderstand` prefix nothing resolves is kept, and an `mc:Ignorable` one is dropped from
  the Choice (the attribute too, where it named nothing else), each with its warning.

The fidelity test's resolution check now reads `mc:MustUnderstand` on every element too, with a
self-check. It already read `mc:Ignorable` on every element.

`npm test` and the typecheck pass. Fidelity passes: 359 checks, 631 prefixes resolve. The 341
corpus hashes are identical to 0.4.3's. On docx4j's fixtures, nothing changes: in the probe's
`document.xml`, `zz9` is on each of 3 Choices and `shape2010` on its Choice, with `wps` and `w14`
on the root; in the charts, `c14` and `cx1` are on the root. Nothing is undeclared, and nothing warns.

## 6. Risks

- **A consumer asserting root placement.** core-ts says its CR-010 section 14.9 test asserts bindings
  through `lookupNamespaceURI`, and its one root assertion is an `mc:Ignorable` prefix. So it needs
  only its lockfile.
- **A serializer dropping the Choice's declaration.** A serializer may drop a declaration that
  repeats the binding in scope, but not one that changes it, which would change the document. The
  Choice's declaration always changes it, since one the root already declares isn't written.

## 7. Consumers

- **core-ts:** no code change; its survey's four copies of docx4j's probe move their two
  declarations from the root to the Choices.
- **docx4j:** CR-052 section 4 takes the same rule (docx4j 4b9c71288, 2026-10-21, unpushed):
  - one declaration on the root when the root's `mc:Ignorable` needs the same binding;
  - nothing on `mc:Fallback`;
  - the narrowed clash rule.

  docx4j's D1 (b), a Choice's own `mc:Ignorable` and `mc:MustUnderstand` following the `Requires`
  rule on that Choice, is here too since 2026-10-24 (section 3). Within it one difference
  remains: a prefix in a Choice's own `mc:Ignorable` that nothing resolves. docx4j keeps it
  undeclared, with a warning; this package drops it, with a warning, as it does from a root's
  `mc:Ignorable`. It arises only where neither the source, the table nor the aliases binds the
  prefix.

## 8. Release

A patch, 0.4.4, as 0.4.3 was: the output changes only for parts with a source-only prefix that a
Choice names, or with a Choice's own `mc:Ignorable` or `mc:MustUnderstand`.
There is no API change, and the runtime floor is unchanged. It needs core-ts's go on the pushed
`main` commit.
