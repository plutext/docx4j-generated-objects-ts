# CR-017: Misplaced children: refused by the builders, reported on read and on save

**Status:** Proposed 2026-10-26; M1 to M6 for Jason
**Requested by:** Jason, 2026-10-26 ("a CR for checking misplaced children (in the builders or as a validation pass)
please"). The question came from docx4j CR-064's hand-off (its section 7): whether the TypeScript port's content lists
refuse a misplaced child, with CR-064 section 2's table of mistakes as the list worth catching.
**Depends on:** nothing new. The read report goes through jsonix-CR-008's unmarshal-side `onElement` (3.5.0), the save
report through the marshal's `onElement` (jsonix-CR-006). Both are already composed by the facade (CR-011, CR-016).
**Counterparts:**
- docx4j CR-064 (docx4j-25, accepted 2026-10-26, phase 1 done in befc57371): `ArrayListDocx4j.checkParent` refuses a
  misplaced child at add time, with the fix in the message. Its phase 2, a tree validator derived from the annotations,
  is optional and open.
- docx4j-python has its own list types; its answer is its own.

## 1. Summary

A content list can hold a child its schema never allows there: a run straight in a cell, a table in a paragraph, a
paragraph in a run. Word catches it on opening, with "unreadable content" and a repair that drops the content, far
from the line that caused it.

docx4j refuses such a child where it is added, since Java's lists hold `Object` and its compiler can't help. This
package differs in three ways (section 2):
- **TypeScript's types already refuse every one of CR-064's mistakes** when code pushes into a model's content array,
  because the declarations type each list as the union of what its schema allows. But JavaScript callers, casts and
  the builders' own `Element[]` parameters are not checked.
- **Nothing intercepts a push:** the content lists are plain arrays, so there is no add-time hook to check in.
- **A misplaced child round-trips silently:** the facade writes it as held and reads it back typed, with no
  `onUnexpectedElement`, where docx4j's unmarshaller would skip it.

Proposed, with docx4j CR-064's refuse-only table (section 3):
- **The builders refuse:** `p`, `tc`, `tr` and `wml`'s interpolation throw where a caller hands them a child that
  doesn't belong, before anything is built, with the fix in the message (M1, M3, M4).
- **A report on read and on save:** a new `onMisplaced` callback on the unmarshal and marshal options reports each
  misplaced child. The tree keeps it as read, and a save writes it as held (M2).

## 2. What happens today (measured 2026-10-26, at 86ae5cb)

**Pushed into a model's content array,** the 15 mistakes of CR-064 section 2's table:

| Mistake | TypeScript (strict) | marshal | read back |
|---|---|---|---|
| a run, or a `w:t`, in the body or a cell | rejected | written: `<w:tc><w:p/><w:r>…</w:r></w:tc>` | typed, no report |
| a drawing, a `w:br` or a `w:t` in a paragraph | rejected | written | typed |
| a table in a paragraph, a paragraph in a run | rejected | written | typed |
| a paragraph in a row, a cell or a table in a table | rejected | written | typed |
| `w:rPr` in a run's content | rejected | written, after the text | typed |
| a string in a run | rejected | throws `Property is not mixed, can't handle string values.` | - |
| `w:pPr` or `w:sectPr` in content | - | there is no `el` factory for them, so only by hand | - |

**Through the builders**, which take `Element[]`, TypeScript accepts them all, and the facade writes them:
- `tc([r('a')])` writes `<w:tc><w:r><w:t>a</w:t></w:r></w:tc>`;
- `tc([tc('a')])` and `tc([tr(['a'])])` nest a cell or a row in a cell;
- `p([tbl(...)])` puts a table in a paragraph;
- ``wml`<w:p>${tbl([['a']])}</w:p>` `` writes a table in a paragraph, and ``wml`<w:tc>${r('z')}</w:tc>` `` a run in a
  cell.

**Why reading doesn't object:** an elementRefs property resolves any element the context declares globally, which
includes `w:p`, `w:r`, `w:tbl` and `w:tc` (CR-009), and CR-021's lax wildcards admit them besides. So the runtime types
a misplaced child wherever it finds it. docx4j skips it as an unexpected element (CR-064 section 8).

## 3. The table (docx4j CR-064's, refuse-only)

Each container refuses only children its schema never allows there, so a legal document can't trip it (CR-064 D4,
D9). It never says "allow only": `mc:AlternateContent`, range markup, `w:altChunk`, the block- and run-level `w:sdt`
and `w:customXml` wrappers, revisions and extension elements all pass. A child is tested by its element name, the
`TypedNamedValue`'s `name`, as CR-064 D1 unwraps a `JAXBElement`.

| Container | Refuses | The fix in the message |
|---|---|---|
| **block-level**: the body, a cell, block-level sdt content, a text box, a header, a footer, a footnote or endnote, a comment, block-level custom XML | `w:r`, `w:t`, `w:drawing`, `w:br`; `w:tc`, `w:tr`; `w:sectPr`; `w:tcPr` in a cell; a string | add it to a paragraph (`p()`) and add that; a cell goes in a row (`tr()`), in a table; use `sectPr`, or the paragraph's `pPr.sectPr`; use `tcPr` |
| **run-level**: a paragraph, a hyperlink, a simple field, `w:ins`/`w:del`, a smart tag, run-level custom XML, run-level sdt content | `w:tbl`, `w:p`; `w:t`, `w:drawing`, `w:br`; `w:pPr`; a string | a table is a sibling of the paragraph; add it to a run (`r()`) and add that; use `pPr` |
| **a run** | `w:r`, `w:p`, `w:tbl`; `w:rPr`; a string | a run's content is text (`t()`), breaks and drawings; use `rPr` |
| **a table** | `w:tc`, `w:p`, `w:r`, `w:tbl`; `w:tblPr`; a string | a cell goes in a row (`tr()`); a nested table goes in a cell; use `tblPr` |
| **a row** | `w:p`, `w:r`, `w:tbl`; `w:trPr`; a string | content goes in a cell (`tc()`); use `trPr` |
| **sdt content** | sdt content directly | the sdt's content goes in its `sdtContent` once |

Containers are keyed by `TYPE_NAME` (`org_docx4j_wml.Body`, `.Tc`, `.SdtContentBlock`, `.P`, `.P.Hyperlink`, `.R`,
`.Tbl`, `.Tr`, ...), each with its content property. That is `content` for all but `RunIns` and `RunDel`, whose runs are
in `customXmlOrSmartTagOrSdt` (all checked in the context, 2026-10-26). Children are keyed by element name in `w`'s
namespace. The messages follow CR-064's form: "You
can't add a run (w:r) directly to a table cell (w:tc); add it to a paragraph (p()) and add that." The message names
the container, as a person, or a model, fixes code from it.

## 4. Proposal

### 4.1 The builders refuse (M1, M3, M4)

Where a caller hands a builder a child, the builder tests it against the table before building:
- **`p(runs)`:** each element, as a run-level child;
- **`tc(blocks)`:** each element, as a block-level child;
- **`tr(cells)`:** each element that isn't a string, as a row's child. A string is already made a cell, and the types
  already ask for `Element<Tc>`, but JavaScript passes anything;
- **`wml`'s tagged form:** an interpolated element, once the fragment is parsed and the element placed, as a child of
  the element it lands in. The plain form parses XML text, so a misplaced element there is the caller's markup,
  reported like any other read (section 4.2).

A refusal is a `TypeError` with the message of section 3, thrown before anything is built, so nothing half-made is
returned. `r`, `t` and `tbl` take strings and options only, so they can't receive a child.

### 4.2 A report on read and on save (M2)

`onMisplaced(report)` is a new callback on `UnmarshalOptions` and `MarshalOptions`, called only when passed, as
`onOutOfRange` is. The report gives:
- `parent`, the container's element;
- `child`, the element's name (or `'#text'` for a string);
- on save, `path`;
- `message`, as section 3's.

On read it runs in the composed `onElement`: for each typed container, its content list is tested once (a lookup by
`TYPE_NAME`, then one by name per child). On save it runs in the marshal's `onElement`, testing each element against
the one it was written into.

The tree keeps what it read, and a save writes it as held: the report is how a caller finds what Word would repair,
before Word does. A JavaScript caller who pushed into a content array learns of it at save, the first point the
package sees the tree.

### 4.3 What doesn't change

- The model's declarations, which already refuse every mistake at compile time for a typed caller (M6).
- Reading: a misplaced child in another producer's file is still read typed, as today, not dropped as docx4j drops it.
  Word repairs such a file rather than refusing it, and the content may be wanted; the report says it is there.
- Saving: written as held. A save that throws would lose more than the misplaced child.

## 5. Not proposed

- **An allow-only check derived from the mapping** (each content property's element list). As CR-064 D8 found for
  docx4j's annotations, it would have to accept subtypes, the lax wildcards, `mc:AlternateContent` and extension
  elements, and a wrong refusal of a legal document is the worst outcome. The refuse-only table can't make one.
- **Order and completeness** (a cell's last block a paragraph, `pPr` first, a row needing a cell). Those are not about
  the child added; core-ts's `tableRules` covers the cell rule (its CR-002 section 55).
- **A standalone tree walker.** The read and save reports cover a loaded tree and a saved one. A caller wanting a
  check without a save can marshal with `onMisplaced` and discard the result.

## 6. Decisions (for Jason)

- **M1. The builders:** (a) refuse a misplaced child, with section 3's table and messages; (b) leave them as they are.
  **Recommendation:** (a). The builders are where TypeScript can't help, since they take `Element[]`, and where the
  caller is.
- **M2. The report:** (a) `onMisplaced` on read and on save, called only when passed, never changing what is read or
  written; (b) on save only; (c) none. **Recommendation:** (a). It is the only check a JavaScript caller's direct push
  meets, and on read it finds another producer's misplaced content before Word repairs it.
- **M3. The table:** docx4j CR-064's refuse-only sets (section 3), so the ports agree, or a narrower one (body and cell
  only). **Recommendation:** CR-064's.
- **M4. The error:** a `TypeError` with the fix in the message, or a plain `Error`. **Recommendation:** `TypeError`. The
  argument is of the wrong kind for where it is going.
- **M5. The gate:** the fidelity corpus's 341 parts must produce no report, on read or on save. Office writes no
  misplaced child, and a rule that trips on one is dropped (CR-064 D9). **Recommendation:** yes. core-ts's corpus
  survey is the larger gate, offered at its next run.
- **M6. The builders' parameter types:** (a) keep `Element[]`; (b) narrow them to the model's own content unions
  (`p(runs: NonNullable<P['content']>)`), so TypeScript catches a misplaced child at compile time too.
  **Recommendation:** (a) now. (b) would break a caller passing `wml()`'s `Element[]`, which is legal at run time, and
  the runtime refusal catches every case.

## 7. Tests

- **The builders** (smoke): each refused child of section 3 in `p`, `tc`, `tr` and `wml`'s interpolation throws a
  `TypeError` naming the container and the fix. The accepted children pass:
  - a hyperlink, a field, `w:ins`, a content control and `mc:AlternateContent` in a paragraph;
  - a table, a content control and custom XML in a cell;
  - a cell, a row-level content control and custom XML in a row.
- **The report** (smoke): each mistake of section 2's table, pushed into a model's array, is reported on save with its
  path and message, and written as held. The same file read back is reported on read, and the tree is unchanged. With
  no callback, nothing is reported.
- **Fidelity:** a part that reaches `onMisplaced`, on read or on save, fails (M5).
- `test/nodenext/consumer.mts`: the callback and the report type.

## 8. Consumers

- **core-ts:** where its content API builds through these builders, a misplaced child it passes would throw where it
  was written before. Its tests would find any such case. It can also map `onMisplaced` to part warnings, as its own CR.
- **docx4j:** none. CR-064 is the counterpart, and this CR takes its table.

## 9. Release

- **Changes:** the builders throw where they wrote a broken document before, and the API gains `onMisplaced`.
- **Version:** a minor under Jason's rule. It can ride with 0.6.0.
- **Effort:** half a day to a day.
