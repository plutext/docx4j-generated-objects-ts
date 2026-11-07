# CR-023: Malformed XML refused on every facade read, in a browser as on Node

**Status:** Implemented 2026-11-07, D1 as decided, D2 to D5 as built (section 5). Not released. For 0.8.0 (Jason: "Add the guard to the
facade in 0.8.0").
**Requested by:** Jason, 2026-11-07, as above. It came out of docx4j-ts-editor's correction of CR-020 (its section 3.1).
**Depends on:** nothing. No change to the runtime, the compiler or `modules/`.
**Counterparts:**
- core-ts CR-002 section 33 is the same guard, in its `parseXml` (`src/xml/dom.mts`).
- CR-020's D4 is the other half: `wml()`'s interpolation takes the check, so an element holding such a character is refused by name
  before it reaches a parser.

## 1. Summary

A browser's `DOMParser` doesn't throw on malformed XML: it returns a document holding a `parsererror` element. The facade handed
that on. In Chromium a `w:document` or a flat OPC package then read as a tree cut off at the fault, with no error, and a save would
have written the loss.

`unmarshalString`, `unmarshalNode`, `unmarshalPackage` and `wml` now refuse such a document with a `SyntaxError` carrying the parser's
message. No fidelity hash moved, and nothing changes for well-formed XML.

## 2. What was there

Measured on 2026-11-07, a bundle of `b4373c8`'s `src` with jsonix 3.13.0, under Playwright. `document` is three paragraphs with the
fault in the second.

| Given | Node | Chromium | Firefox |
|---|---|---|---|
| A `w:document` holding U+000B, U+FFFE, `&#xB;` or a stray `<` | Read, the words kept | **The first paragraph alone, no error** | Throws, naming `parsererror` |
| A flat OPC package, the same in a part | Not run | **No error, 2 parts reported** | The same |
| A `w:p` or `w:r` root holding U+000B | Read | Throws, naming `parsererror` | The same |

- **The shape decides it.** Chromium puts `parsererror` first in the root and keeps what came before the fault. A `w:document` passes
  over an element it doesn't know, and a run's or paragraph's content refuses one, which is why docx4j-ts-editor saw a throw.
- **With nothing parsed,** Chromium gives `html`, `body`, then `parsererror`, one level below where core-ts looks. That threw anyway,
  as `html` is no known root.
- **Word** refused all eight faults in a main part measured on 2026-11-06, U+000B and a bare `&` among them, with no repair. A stray
  `<` and U+FFFE weren't among the eight.
- **It isn't new.** The parse line is the same at 0.7.0, by reading.
- **core-ts** reads through its own guard, so this reaches an add-in on this package alone, and any `unmarshalString` of a whole part.

## 3. The change

In a new internal `src/parse.mts`, not in the `exports` map:

- **`parseXml(xml)`:** `Jsonix.DOM.parse`, then the check. `unmarshalString` and `wml`'s fragment parse use it, and `unmarshalPackage`
  reads through `unmarshalString`.
- **`refuseParserError(node)`:** `unmarshalNode` calls it, for a document a caller parsed itself.
- **The test:** a `parsererror` in Chromium and WebKit's namespace or Firefox's, at the root, its first element, or that element's
  first. A `parsererror` in any other namespace is only an element.
- **The error:** `Malformed XML: ` and the parser's message, as core-ts words it. Chromium's is its `div`. Firefox's is the text ahead
  of `sourcetext`, which would repeat the whole source line.

## 4. Results

- In Chromium and Firefox: eight malformed shapes through `unmarshalString`, four faults in a `w:document` and in a package, a string
  `checkCharacters: false` wrote, `wml` given such text, and `unmarshalNode` of a document parsed by the caller. All are refused with
  the parser's line and column, or by CR-020's `RangeError` where `wml` interpolates an element. The controls read as before.
- `npm run typecheck` and `npm test` pass on Node 18.20.8, 20.20.2 and 22.23.4, with jsonix 3.11.0.
- All 341 marshalled hashes are identical to `106319f`'s.

WebKit wasn't run: its system libraries aren't on this machine. core-ts's `src/xml/dom.mts` records Chromium's form as WebKit's too.
core-ts and docx4j-ts-editor have not been run against this CR.

## 5. Decisions

D1 is Jason's. D2 to D5 were settled in building it and weren't put to him.

| D | Question | Decided |
|---|---|---|
| D1 | In 0.8.0? | Yes. |
| D2 | Here, or in jsonix's `DOM.parse`? | Here. It needs no runtime release, and core-ts's is at the same level. The runtime stays as it is for other mappings. |
| D3 | Which error? | `SyntaxError`. The facade throws `RangeError` and `TypeError` for what a caller gave it, and xmldom's own `ParseError` stays as it is. Nothing here makes the two one class. |
| D4 | `unmarshalNode` too? | Yes, since a caller may parse first and the loss is the same. It costs a look at three elements. |
| D5 | A browser in this repository's tests? | No. The smoke holds the documents the two browsers return, written out, so xmldom reads them. That checks the guard against their shape, not against a browser. A later browser release that changes the shape wouldn't fail here. |

## 6. Tests and docs

- **Smoke, a CR-023 block:** Chromium's document, with and without a root, and Firefox's, through `unmarshalString`, `unmarshalPackage`,
  `wml` and `unmarshalNode` (a document and its root). The messages exact. A `parsererror` in another namespace read as an element.
  xmldom's own throw left as it is.
- **The gate:** every fidelity hash identical.
- **Docs:** `CLAUDE.md` and a README bullet.

## 7. Consumers

- **core-ts:** its own reads refuse first, so nothing should change. A direct `unmarshalString` in a browser now throws where it read a
  truncated tree. Its release gate as usual.
- **docx4j-ts-editor:** it reads through core-ts. A read-back of such a string in a browser now fails with this message, in any root.
- **An add-in on this package alone:** a read that lost content silently now throws. That is the point of the CR, and it is still a
  new throw.

## 8. Not in this CR

- **xmldom's recovery.** On Node it throws for most faults, and for some it carries on: `a<b` reads as that text, `a & b` the same,
  an unquoted attribute reads, and text after the root is dropped (measured 2026-11-06). Those aren't refused here, though a browser
  refuses the first three, and Word `a & b` and text after the root. So a string can read on Node and throw in a browser.
- **A caller's own `createUnmarshaller()`,** which doesn't pass through the facade, as CR-013's and CR-018's readings don't.
