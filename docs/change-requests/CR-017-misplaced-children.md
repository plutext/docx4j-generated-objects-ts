# CR-017: Misplaced children: refused by the builders, reported on read and on save

**Status:** Accepted 2026-10-28, M1 to M7 as recommended (Jason: "CR-017 decisions as recommended; commit then
implement CR-017"); proposed 2026-10-26, revised 2026-10-28 to docx4j CR-064 phase 1's code (befc57371), core-ts
CR-016's decisions (D1 to D6, 2026-10-27) and core-ts's check 72 in Word 365
**Requested by:** Jason, 2026-10-26 ("a CR for checking misplaced children (in the builders or as a validation pass)
please"). The question came from docx4j CR-064's hand-off (its section 7): whether the TypeScript port's content lists
refuse a misplaced child, with CR-064 section 2's table of mistakes as the list worth catching. Revised at Jason's word,
2026-10-28 ("revise CR-017"), after a review against what docx4j built and what core-ts decided.
**Depends on:** nothing new. The read report goes through jsonix-CR-008's unmarshal-side `onElement` (3.5.0), the save
report through the marshal's `onElement` (jsonix-CR-006). Both are already composed by the facade (CR-011, CR-016).
**Counterparts:**
- docx4j CR-064 (accepted 2026-10-26, phase 1 done in befc57371): `ArrayListDocx4j.checkParent` unwraps a `JAXBElement`
  and refuses a misplaced child by its class as it is added, with an `IllegalArgumentException` whose message gives the
  fix. Its phase 2, a tree validator, is optional; docx4j CR-066 took its own whole-tree need to the CR-050 inspector.
- core-ts CR-016 (D1 to D6 decided 2026-10-27) is the consumer. Its phase 1 is this CR, and its phase 2 refuses a save
  on this CR's save report, with an opt-out. Its D3 and D4 (the rules by type, a name and value that disagree,
  docx4j's messages) are taken here (section 3).
- docx4j-python has its own list types; its answer is its own.

## 1. Summary

A content list can hold a child its schema never allows there: a run straight in a cell, a table in a paragraph, a
paragraph in a run. core-ts's check 72 measured what Word 365 does with eighteen such forms (section 2.2): eleven don't
open at all, with no repair offered, and seven open with no message and are repaired silently, one losing its text.

docx4j refuses such a child where it is added, since Java's lists hold `Object` and its compiler can't help. This
package differs in four ways (section 2):
- **TypeScript's types already refuse CR-064's mistakes** when code pushes into a model's content array, because the
  declarations type each list as the union of what its schema allows. JavaScript callers, casts, a `{ name, value }`
  whose name and value disagree, and the builders' own `Element[]` parameters are not checked.
- **Nothing intercepts a push:** the content lists are plain arrays, so there is no add-time hook.
- **A misplaced child the context declares globally round-trips silently:** the facade reads it back typed and writes
  it as held, where docx4j's unmarshaller skips it as unexpected.
- **A misplaced child the context doesn't declare globally fails the read** of its whole part (a field instruction
  straight in a paragraph). The same child pushed from code is saved with an `xsi:type`, so the facade writes a part it
  can't read back.

Proposed, with docx4j CR-064's sets as its code has them and core-ts CR-016's rules and messages (section 3):
- **The builders refuse:** `p`, `tc`, `tr`, `sdt` and `wml`'s interpolation throw where a caller hands them a child
  that doesn't belong, before anything is built, with the fix in the message (M1, M3, M4).
- **A report on read and on save:** a new `onMisplaced` callback on the unmarshal and marshal options reports each
  misplaced child. Each container's list is tested as the container is met, on save before its children are written,
  so a callback can stop a save before the runtime's own error. The tree keeps what it read, and a save writes it as
  held (M2). core-ts's refused save is built on it.
- **Nine more elements declared globally:** the run and paragraph elements whose misplacement fails a read today are
  read typed and reported like the rest, a divergence from docx4j's roots that docx4j calls harmless (M7, section
  4.3).

## 2. What happens today

Measured 2026-10-26 at 86ae5cb, and again 2026-10-28 at 8d41d30.

### 2.1 In this package

**Pushed into a model's content array:**

| Mistake | TypeScript (strict) | marshal | read back |
|---|---|---|---|
| a run, `w:t`, `w:br`, a drawing, a cell or a row in the body or a cell | rejected | written as held | typed, no report |
| a table, a paragraph, `w:t` or `w:br` in a paragraph | rejected | written | typed |
| a paragraph, a run or a table in a run | rejected | written | typed |
| a paragraph, a run, a cell or a table straight in a table or a row | rejected | written | typed |
| `w:instrText`, `w:delInstrText` or `w:fldData` in a paragraph or the body | rejected | written with `xsi:type="w:CT_Text"` | **the read fails** (section 2.3) |
| `w:pPr`, `w:rPr`, `w:sectPr`, `w:tblPr`, `w:trPr` or `w:tcPr` in content | rejected | written where pushed (no `el` factory makes `w:pPr` or `w:sectPr`) | absorbed into its property and written first: no misplaced child is left |
| a string in a run, a paragraph or the body | rejected | throws `Property is not mixed, can't handle string values.` | - |
| a `{ name, value }` whose value is another class (`w:p` over a run) | accepted where the value's class belongs | `<w:p xsi:type="w:R">` | - |

**Through the builders,** which take `Element[]`, TypeScript accepts them all, and the facade writes them:
- `tc([r('a')])` writes `<w:tc><w:r><w:t>a</w:t></w:r></w:tc>`;
- `tc([tc('a')])` and `tc([tr(['a'])])` nest a cell or a row in a cell;
- `p([tbl(...)])` puts a table in a paragraph, and `p([inlinePicture(...)])` a drawing outside a run;
- `sdt([r('a')], { form: 'block' })` and `sdt([p('a'), r('b')])` write a block-level content control holding a run:
  the mistake docx4j's build found in its own export code (CR-064 section 8);
- a content control of one level placed at the other is written with an `xsi:type`: `<w:sdt xsi:type="w:SdtBlock">`
  in a paragraph, `<w:sdt xsi:type="w:CT_SdtRun">` in the body;
- ``wml`<w:p>${tbl([['a']])}</w:p>` `` writes a table in a paragraph, and ``wml`<w:tc>${r('z')}</w:tc>` `` a run in a
  cell.

### 2.2 In Word 365 (core-ts's check 72)

core-ts's check 72 (its CR-016 section 7.1; Jason, Word 365 16.0.20326.20232, 2026-10-27) is eighteen files, one
misplaced child each, covering every set of section 3:
- **Eleven don't open:** "Word experienced an error trying to open the file", with no repair offered. They hold a run,
  bare text, a cell or a row in the body; a run beside a cell's paragraph; a table, a paragraph or bare text in a
  paragraph; a run in a run; a cell beside a table's row; a run beside a row's cells.
- **Seven open with no message** and are repaired to the nearest place the child fits. A break in the body or in a
  paragraph goes into the next run. A paragraph in a run joins the outer paragraph. A paragraph beside a table's row or
  a row's cells moves before the table or into the first cell, and a table beside a row becomes a row. A string
  straight in a run is dropped, the only content lost.

So the forum's "unreadable content, then a repair" (CR-064 section 2) understated it: most of these files don't open.

### 2.3 Why reading doesn't object, and when it fails

The content lists are elementRefs properties that don't admit DOM. An element the list declares is typed as declared;
any other element is typed by its global declaration. The context declares `w:p`, `w:r`, `w:tbl`, `w:tc`, `w:tr`,
`w:t`, `w:br` and `w:drawing` globally, so the runtime types a misplaced one wherever it finds it, with no
`onUnexpectedElement`. (As proposed on 2026-10-26, this section also credited CR-021's lax wildcards; those are on
`mc:Choice` and `mc:Fallback`, not on these lists.)

An element with no global declaration fails the read instead. A run's `w:instrText`, `w:delInstrText`, `w:object`,
`w:ruby`, `w:footnoteReference` and `w:endnoteReference` are such, and so are a paragraph's `w:dir`, `w:bdo` and
`w:subDoc` in the body. One of them misplaced fails its whole part:

```
w:instrText (line 1): Element [{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w:instrText] could not be
unmarshalled as is not known in this context and the property does not allow DOM content.
```

Fifteen more fail the read only when misplaced inside a run: `w:customXml`, `w:moveFrom`, `w:moveTo`, and twelve range
markers for moves and custom XML.

docx4j's JAXB skips both kinds as unexpected elements (CR-064 section 8), so docx4j loads such a file without the
child, which never reaches a list. Its gate's clean result over 12,575 parts is therefore certain by construction, and
says nothing about this runtime, which keeps the child (M5).

## 3. The rules

docx4j's `checkParent` as befc57371 has it, with core-ts CR-016's D3 and D4. Each container refuses only children its
schema never allows there, so a legal document can't trip it (CR-064 D4, D9). It never says "allow only":
`mc:AlternateContent`, range markup, `w:altChunk`, the block- and run-level `w:sdt` and `w:customXml` wrappers,
revisions and extension elements all pass.

**A child is tested by its class,** as docx4j tests the value a `JAXBElement` wraps (CR-064 D1) and core-ts's D3
decided: the value's `TYPE_NAME`, or for a hand-built value without one, the class its element name declares. A class
can have several names. docx4j's `Text` is `w:t`, `w:instrText`, `w:delInstrText` and `w:fldData` here, all refused
where `Text` is; as proposed on 2026-10-26, the table tested `w:t` alone.

| Container | Refuses | The fix in the message |
|---|---|---|
| **block-level**: the body, a cell, block-level sdt content, a text box, a header, a footer, a footnote or endnote, a comment, block-level custom XML | a run; text, a drawing, a break; a cell, a row; `w:sectPr`; `w:tcPr`; a string | add a run to a paragraph, and text, a drawing or a break to a run in a paragraph, and add the paragraph; a cell goes in a row, in a table; use the body's `sectPr`, or `pPr.sectPr` on the paragraph that ends the section; use `tcPr` |
| **run-level**: a paragraph, a hyperlink, a simple field, `w:ins`/`w:del`, a smart tag, run-level custom XML, run-level sdt content | a table; a paragraph; text, a drawing, a break; `w:pPr`; a string | a table is a sibling of the paragraph; add it to a run and add that; use `pPr` |
| **a run** | a run, a paragraph, a table; `w:rPr`; a string | add the run to the paragraph; use `rPr`; put text in `w:t` |
| **a table** | a cell; a paragraph, a run; a table; `w:tblPr`; a string | a cell goes in a row; content goes in a cell in a row; a nested table goes in a cell; use `tblPr` |
| **a row** | a paragraph, a run, a table; `w:trPr`; a string | content goes in a cell; use `trPr` |
| **sdt content** | sdt content directly | nest a `w:sdt` that holds it |

As in docx4j, `w:tcPr` is refused in every block-level container; as proposed on 2026-10-26, only in a cell.

**A name and value that disagree** (core-ts D3): a child whose value's class isn't the one its element name declares
in that list, or globally where the list doesn't name it, such as `{ name: w:p, value: <a run> }`. A content control at
the wrong level is one: `w:sdt` in a paragraph declares the run-level class, so a block-level value there disagrees.
Such a child is exactly one the marshal writes with an `xsi:type`, which Office never writes. The runtime types each
element it reads by its name, so a file read always agrees, and this rule catches code only.

**`mc:AlternateContent`'s branches** are tested by their host's set (core-ts D3): a kept `mc:Choice` or `mc:Fallback`
in a paragraph is tested as run-level content.

**The messages** are docx4j's (core-ts D4), one wording across the ports, with the element name where docx4j names a
Java class: "You can't add a run (w:r) directly to a table cell (w:tc); add it to a paragraph (w:p) and add that." The
builders throw the message the report carries. Each names the container, since a person, or a model, fixes code from
it.

Containers are keyed by `TYPE_NAME` and their base types (`org_docx4j_wml.Body`, `.Tc`, `.SdtContentBlock`, `.P`,
`.P.Hyperlink`, `.R`, `.Tbl`, `.Tr`, ...), each with its content property. That is `content` for all but `RunIns` and
`RunDel`, whose runs are in `customXmlOrSmartTagOrSdt` (all checked in the context, 2026-10-26).

**Three cases neither port refuses** were sent to docx4j on 2026-10-28 (section 8):
- `w:delText`, whose docx4j class `DelText` is not a `Text`;
- a `w:tab` straight in a paragraph, which CR-064 section 2's row 7 names and its code doesn't refuse;
- a content control at the wrong level, which docx4j writes as a `w:sdt` holding the other level's content.

The name rule covers the last here. The other two are added when docx4j adds them, so the ports agree. docx4j
confirmed all three against `checkParent` the same day, and has put them to Jason as a CR-064 addendum, with the
message fix below; not yet ruled.

## 4. Proposal

### 4.1 The builders refuse (M1, M3, M4)

Where a caller hands a builder a child, the builder tests it against section 3 before building:
- **`p(runs)`:** each item, as run-level content;
- **`tc(blocks)`:** each item, as block-level content;
- **`tr(cells)`:** each item that isn't a string, as a row's. A string is already made a cell, and the types already
  ask for `Element<Tc>`, but JavaScript passes anything;
- **`sdt(content, options)`:** each item, as the content of the form given or inferred, block-level or run-level.
  `sdt([p('a'), r('b')])`, inferred block-level from its first item, refuses the run. The row and cell forms are tested
  by the name rule only;
- **`wml`'s tagged form:** an interpolated element, once the fragment is parsed and the element placed, as a child of
  the element it lands in. The plain form parses XML text, so a misplaced element there is the caller's markup,
  reported like any other read (section 4.2).

A refusal is a `TypeError` with the message of section 3, thrown before anything is built, so nothing half-made is
returned. `r`, `t` and `tbl` take strings and options only, so they can't receive a child.

### 4.2 A report on read and on save (M2)

`onMisplaced(report)` is a new callback on `UnmarshalOptions` and `MarshalOptions`, called only when passed, as
`onOutOfRange` is. The report gives:
- `element`, the container's element: on read the source element, on save the element written, open, its children
  not yet written;
- `container`, its `TYPE_NAME`;
- `child`, the child's element name, or `'#text'` for a string, and `childType`, its class where it has one;
- `index`, its position in the container's list;
- `rule`, `'misplaced'` or `'name-value'` (section 3), and `message`, section 3's;
- on save, `path`, in `Normalisation`'s and `OutOfRange`'s form and built by the same function: the container's path
  and the child's step (`/w:document/w:body[1]/w:r[1]`), or the container's path for a string. core-ts formats all
  three reports with one formatter (its ask, 2026-10-28).

Both sides test each container's list as a whole, once, when the container is met. On read that is the composed
unmarshal `onElement`, once the container's value is complete. On save it is the marshal's `onElement`, which the
runtime calls as it opens the element, before its children (measured 2026-10-28: `document`, `body`, `p`, `r`, `t`).
So on save:
- a string is reported before the runtime throws `Property is not mixed` over it;
- a callback that throws stops the save there with its own error. core-ts's refused save (its D2) can throw naming the
  part, the path, the rule and the fix, where its save fails today with "Problem saving part /word/document.xml" and
  nothing about the cause.

The tree keeps what it read, and a save writes it as held unless the callback throws. On read, a misplaced property
element never reaches a list (section 2.1), so those rules report a push from code only. A JavaScript caller who pushed
into a content array learns of it at save, the first point the package sees the tree.

### 4.3 Nine more elements declared globally (M7)

The nine elements of section 2.3 whose misplacement fails a read are declared in CR-009's hand-written global mapping,
`src/wml-globals.mts`, with their classes:
- `w:instrText` and `w:delInstrText` (`Text`), `w:object` (`CTObject`), `w:ruby` (`CTRuby`), `w:footnoteReference`
  and `w:endnoteReference` (`CTFtnEdnRef`);
- `w:dir` (`P.Dir`), `w:bdo` (`P.Bdo`) and `w:subDoc` (`CTRel`).

A misplaced one then reads typed, as a misplaced `w:t` does: kept, reported by `onMisplaced` where section 3 refuses
it, and saved as held. A pushed one is written without an `xsi:type`, so the facade reads back what it writes.

**A divergence from docx4j, recorded in CR-009 when implemented.** CR-009's mapping holds exactly the elements docx4j
declares as roots that the generated mapping doesn't declare globally. Reconciled on 2026-10-28 against docx4j's
generated sources at HEAD: 99 root names, of which the generated mapping declares 37, CR-009 the other 61, and
`w:sdtContent`, with four classes, neither. These nine are not docx4j roots. Asked the same day, docx4j won't make them
roots on its own account:
- its lax wildcards (`mc:AlternateContent`, `w:pict`'s and the shape defaults') are where none of the nine
  legitimately occurs, so a root would type them only where they are misplaced;
- the one gain, unmarshalling a fragment rooted at one of them, nobody has asked for;
- five of them would need unscoped `@XmlElementDecl`s, since `Text`, `CTFtnEdnRef` and `CTRel` serve several names,
  and every port's root list would change. If a measured document ever needs it, docx4j adds it then and tells the
  ports.

docx4j also called the divergence harmless to itself and its consumers. Its root set is an artefact of the schema's
global declarations, not a contract. Its inspector reads XML marshalled from JAXB, and its MCP server and docx4j-python
don't consult JAXB roots. The divergence shows only on misplaced content: this package reads, types and round-trips a
`w:instrText` straight in a `w:p`, where docx4j drops it from a strict container and keeps it as DOM in a lax one.
Both ports' parity fixtures are Word-written and hold no such thing.

**Measured on a prototype** (2026-10-28, the nine added to a copy of `dist/`): each of the nine, misplaced, reads typed
and saves as held, and a pushed `w:instrText` is written without an `xsi:type`. The fidelity test passes, and the
marshalled output of all 342 fidelity parts is byte-identical to the build without them. Each of the nine also becomes
a valid root for `unmarshalString` and `marshalString`, as CR-009's 61 are.

**Not declared:** the fifteen that fail only when misplaced inside a run. `w:customXml` can't be, with a different class
at each level, as `w:sdtContent` has. `w:moveFrom`, `w:moveTo` and the twelve range markers could be, but section 3
refuses none of them in a run, and they aren't measured.

### 4.4 What doesn't change

- The model's declarations, which already refuse at compile time every mistake but a name and value that disagree (M6).
- Reading: a misplaced child declared globally is still read typed, as today, not dropped as docx4j drops it. Word
  refuses most such files outright (section 2.2), and a silent drop would lose the content with no word to the caller;
  kept and reported, it can be fixed. core-ts's D6 refuses the save of such a part unless the caller opts out. (As
  proposed on 2026-10-26, this said Word repairs such a file rather than refusing it; check 72 measured otherwise.)
- The fifteen elements that fail only inside a run (section 2.3) still fail the read when misplaced there (M7).
- Saving: written as held. The facade doesn't refuse on its own account; a consumer that wants a refusal throws from
  the callback.

## 5. Not proposed

- **An allow-only check derived from the mapping** (each content property's element list). As CR-064 D8 found for
  docx4j's annotations, it would have to accept subtypes, `mc:AlternateContent` and extension elements, and a wrong
  refusal of a legal document is the worst outcome. The refuse-only sets can't make one. The name rule uses the mapping
  too, but only to compare a value with its own name, which every file read satisfies.
- **Order and completeness** (a cell's last block a paragraph, `pPr` first, a row needing a cell). Those are not about
  the child added; core-ts's `tableRules` covers the cell rule (its CR-002 section 55).
- **A standalone tree walker.** The read and save reports cover a loaded tree and a saved one. A caller wanting a
  check without a save marshals with `onMisplaced` and discards the result, as core-ts's on-request check (its D1) can.

## 6. Decisions

**Decided 2026-10-28 as recommended** (Jason: "CR-017 decisions as recommended; commit then implement CR-017"): M1 (a),
M2 (a), M3 (a), M4 (a), M5 yes, M6 (a), M7 (d). The options and reasons stand below.

M1, M5 and M6 stand as proposed on 2026-10-26, M1 with `sdt` added and M5 with the builders' own output. M2 to M4 are
revised to what core-ts decided for its phase 2 (its D2 to D4, 2026-10-27). M7 is new, and its recommendation moved
from (a) to (d) the same day, after docx4j's answer and a prototype (section 4.3).

- **M1. The builders:** (a) `p`, `tc`, `tr`, `sdt` and `wml`'s interpolation refuse a misplaced child, with section 3's
  rules and messages; (b) leave them as they are. **Recommendation:** (a). The builders are where TypeScript can't
  help, since they take `Element[]`, and where the caller is. `sdt` writes today the mistake docx4j found in its own
  code.
- **M2. The report:** (a) `onMisplaced` on read and on save, called only when passed, each container's list tested as
  the container is met (on save, before its children), never changing what is read or written; (b) on save only; (c)
  none. **Recommendation:** (a). core-ts's refused save is built on the save side, and a callback that runs before the
  children lets it refuse before the runtime's own error. The read side tells a caller at load what a save would
  refuse.
- **M3. The rules:** (a) docx4j's `checkParent` sets as befc57371 has them, a child tested by its class, with core-ts's
  name rule and `mc:AlternateContent`'s branches tested by their host's set (section 3); (b) the sets as proposed on
  2026-10-26, a child tested by its element name. **Recommendation:** (a). It is docx4j's code and core-ts's D3 at once;
  (b) misses three of `Text`'s four names, `w:tcPr` outside a cell, and every name and value that disagree.
- **M4. The error and the messages:** (a) a `TypeError` from the builders, with docx4j's wording and element names, the
  same message in the builder's error and in the report; (b) a plain `Error`, or messages naming this package's
  builders, as proposed on 2026-10-26. **Recommendation:** (a). The argument is of the wrong kind for where it is going,
  and one wording across the ports is core-ts's D4; a message naming `p()` reads oddly in a report on a loaded file.
- **M5. The gate:** the fidelity corpus's 341 parts produce no report, on read or on save, and neither does the
  builders' own output in the smoke test. **Recommendation:** yes. docx4j's clean gate was certain by construction
  (section 2.3); this runtime keeps a misplaced child, so the corpus here is real evidence. docx4j's one trip was in
  its own code, hence the builders' output. core-ts's survey of 9,071 parts is the larger gate, and its D6 turns a rule
  that trips there into a report.
- **M6. The builders' parameter types:** (a) keep `Element[]`; (b) narrow them to the model's own content unions
  (`p(runs: NonNullable<P['content']>)`), so TypeScript catches a misplaced child at compile time too.
  **Recommendation:** (a). (b) would break a caller passing `wml()`'s `Element[]`, which is legal at run time, and the
  runtime refusal catches every case.
- **M7. A misplaced child the context doesn't declare globally** (section 2.3): (a) the read fails as today, recorded
  and pinned by a test; (b) a jsonix CR so the runtime skips such an element and calls `onUnexpectedElement`, as
  docx4j's JAXB does; (c) keep it as DOM; (d) declare the nine run and paragraph elements globally in CR-009's mapping,
  a recorded divergence from docx4j's roots (section 4.3). **Recommendation:** (d).
  - (d) applies this CR's rule evenly: a misplaced field instruction is read, kept and reported as a misplaced `w:t`
    is, its part stays readable, and a pushed one saves without an `xsi:type`. It costs nine entries and an amended
    CR-009 rule, all in this repository. The prototype changes no byte of the fidelity output, and docx4j calls the
    divergence harmless.
  - (a) leaves the line between reading and failing where docx4j's root declarations happen to put it. core-ts keeps
    the bytes of a part that fails to read, so nothing is lost, but nothing of that part can be read either.
  - (b) is a runtime change in another repository, and once the caller saves, the skipped element is gone. The read
    report could carry the skip, so the cost is the loss, not silence.
  - (c) needs the runtime to keep DOM where the schema admits none.

  None of the four matters to an Office-written file: no part in the fidelity corpus holds such a child, since one
  that throws fails the test, and Word refuses the comparable `w:t` form outright (check 72's 72i). (Before docx4j's
  answer, this recommended (a), and gave (b)'s cost as silence.)

## 7. Tests

- **The builders** (smoke): each refused child of section 3 in `p`, `tc`, `tr`, `sdt` and `wml`'s interpolation throws a
  `TypeError` naming the container and the fix, `w:instrText` in `p` and `sdt([r('a')], { form: 'block' })` among
  them. The accepted children pass:
  - a hyperlink, a field, `w:ins`, a content control and `mc:AlternateContent` in a paragraph;
  - a table, a content control and custom XML in a cell;
  - a cell, a row-level content control and custom XML in a row.
- **The report on save** (smoke): each mistake of section 2.1, pushed into a model's array, is reported with its path
  and message, and written as held. The path has `OutOfRange`'s form, step for step. A string in a run is reported
  before the runtime's throw, and a callback that throws stops the save with its error. A name and value that disagree, and a content control at the wrong level, are
  reported by the name rule. With no callback, nothing is reported.
- **The report on read** (smoke): each globally declared mistake read from a file is reported, and the tree is
  unchanged. A misplaced property element is absorbed and not reported. A misplaced `w:instrText` reads typed and is
  reported (M7 (d)), and a `w:moveFrom` inside a run still fails the read.
- **CR-009's block** (smoke): 70 entries, the nine marked as this package's own, each resolving globally to its class.
  The generated module declares none of the 70 itself, so a regeneration that does retires the entry, as for the 61.
- **Check 72's forms** (smoke, written inline as core-ts's `fixtures/check72/build.mjs` writes them): each reported on
  read and on save, but the string in a run, which the read drops as Word does.
- **Fidelity:** a part that reaches `onMisplaced`, on read or on save, fails (M5).
- `test/nodenext/consumer.mts`: the callback and the report type.

## 8. Consumers and counterparts

- **core-ts:** its CR-016 phase 2 wires the save side to its refused save (D2, with `pkg.checkContentOnSave = false` to
  report instead), and the read side to part warnings, and runs its corpus as the gate (D6). Where its content API
  builds through these builders, a misplaced child it passes would throw where it was written before; its tests would
  find any such case. Told of this revision on 2026-10-28, core-ts confirmed the same day that the report gives its
  refused save everything it needs (the part, the path, the rule and the fix), asking only that `path` keep
  `OutOfRange`'s form (section 4.2). It also corrected its check 72 record: the bare breaks of 72e and 72j are written
  back as they were, so only the string in a run is lost.
- **docx4j:** CR-064 is the counterpart, and this CR takes its sets. Told on 2026-10-28: its `w:tcPr` message says "a
  cell's content" in any block-level container, the body included; the three cases neither port refuses (section 3);
  and check 72's measurement of Word, for CR-064 section 2. docx4j confirmed all four the same day. It cites check 72
  in CR-064 section 2 and records this CR in its section 7. The rule additions and the message fix wait on Jason's
  ruling there, and docx4j will say when he rules. Asked on 2026-10-28 for its view of the nine declarations, docx4j
  won't make them roots, and calls the divergence harmless (section 4.3).

## 9. Release

- **Changes:** the builders throw where they wrote a broken document before, and the API gains `onMisplaced`.
- **Version:** a minor under Jason's rule, 0.7.0. It can't ride with the 0.6.1 patch, `OutOfRange`'s table entry at
  core-ts's ask. (As proposed on 2026-10-26, it was to ride with 0.6.0, since released without it.)
- **Effort:** a day: the rules, both reports, the nine declarations and the gates.
