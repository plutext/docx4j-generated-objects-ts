# CR-020: A character XML 1.0 excludes, refused when the facade writes a string

**Status:** Implemented 2026-11-07, D1 to D3 as decided. Not released. Accepted the same day, D1 to D3 as recommended (Jason:
"CR-020's D1 to D3 as recommended; commit then implement"; section 6). Proposed the same day.
**Requested by:** Jason, 2026-11-07 ("yes", to writing it). It came out of the Word check of an ill-formed main part
(2026-11-06): Word refuses a part holding U+000B, this package reads one, and its save keeps the character.
**Depends on:** nothing. No change to the runtime, the compiler or `modules/`.
**Counterparts:**
- core-ts CR-002 section 60 D3 is the same guard at a part's save (`assertXmlCharacters`, `src/xml/dom.mts:72`).
- docx4j-python refuses through lxml at save, kept as its CR-013's D3.
- docx4j has no guard (section 2.3). Whether it gets one is a separate decision, not this CR's.

## 1. Summary

XML 1.0 excludes some characters: the C0 controls but tab, line feed and carriage return, U+FFFE, U+FFFF and an unpaired
surrogate. No escape can write them, so a part holding one isn't well-formed. `marshalString` and `marshalPackage` wrote
them as held and reported nothing.

Both now refuse such a string with a `RangeError` naming the character and where it is. `marshalString` has a per-call way
round it, since callers use it for more than saves (section 3). Nothing is rewritten. No fidelity hash moved.

## 2. What was there

### 2.1 This package

Before this CR, measured at `4f84e1e` on Node 22, jsonix 3.11.0. "Refused" means expat refuses the output.

| Route | Result |
|---|---|
| `marshalString`, U+0000, U+000B, U+001F, U+FFFE or U+FFFF in a `w:t` | Written raw, no report; refused |
| `marshalString`, U+000B in an attribute | Written raw, no report; refused |
| A file holding U+000B read, then `marshalString` | Written raw, no report; refused |
| `marshalPackage`, U+000B in a part | Written raw, no report (expat not run) |
| `marshalNode` | The returned `Element` holds it; no string is written |
| `r()` and `p()` given U+000B | Held in the tree |
| `marshalString`, an unpaired surrogate | Kept in the string; encoding it as UTF-8 then gives U+FFFD, silently |

### 2.2 Word

Word 16.0.20326.20232 refused a main part with U+000B in a `w:t`, on both runs ("Word experienced an error trying to open
the file"; the share's `bridge/portfolio/wf-*`). That was a built file: this package's output hasn't been run through
Word. What `insertOoxml` does with such a character isn't measured.

### 2.3 The other engines

- **core-ts:** `XmlPart` and `DefaultXmlPart` refuse at save. Its writers drop such characters first (CR-002 section 60
  E3), and the repair policy across the ports records the guard as its one save refusal with no option.
- **docx4j-python:** a save fails through lxml ("All strings must be XML compatible").
- **docx4j:** no guard. Measured here on 17.3.5 (JAXB reference implementation, Java 17, slf4j-simple at WARN), a `w:t`
  built through the API:
  - U+0000, U+000B, U+001F, U+FFFE and U+FFFF: `save` and `XmlUtils.marshaltoString` write each raw, with no exception
    and nothing logged. expat refuses the saved part, and docx4j's own reload fails at `getContent()` ("An invalid XML
    character (Unicode: 0xb)").
  - An unpaired high surrogate: `save` turned `a\uD83Db` into `a` and U+1F462, so the surrogate and the letter after it
    became one wrong character, silently. The part is well-formed and reloads.

## 3. Who reaches it

- **A caller's text.** Office JS's `Paragraph.text` reads a `w:br` as U+000B, a page break as U+000C, a column break as
  U+000E, the hyphens as U+001E and U+001F, and a note reference as U+0002 (core-ts check 48, CR-002 section 60). An
  add-in on this package alone that gives such text to `r()` and saves writes a bad part. `r()` maps `\t` and `\n` only,
  and that contract stays (section 60).
- **A file that already holds one.** This package reads it, and a save writes it back.
- **core-ts.** Its part saves are guarded, but they go through `marshalNode`. Four routes call `marshalString` with no
  guard: `Paragraph.getXml`, `ContentControl.getXml`, a range's blocks (`office-js/fragment.mts:69`) and the API script's
  fallback (`office-js/toApiScript.mts:274`). Its writers drop such characters, so only a value set by hand or read from
  a file gets there.
- **docx4j-ts-editor.** `marshalString` isn't only its save. `editor-model/src/session.mts:558` keeps a run's XML as a
  source entry, and `session.mts:1197` marshals an element and reads it straight back. Those round trips worked with such
  a character. A refusal there would throw mid-edit, far from any save, on a document that opened. This is why D2 isn't
  simply "no".
- **`wml()`'s interpolation.** It marshals an element only to parse it again (`src/builders/wml.mts`), so it is the same
  kind of round trip, inside this package. It passes `checkCharacters: false`.

## 4. The change

In `src/index.mts`, `refuseExcludedCharacters` runs on the string each function returns:

```ts
/** Any character outside XML 1.0's `Char` production; with the `u` flag an unpaired surrogate is one. */
const NOT_XML_CHAR = /[^\t\n\r -퟿-�\u{10000}-\u{10FFFF}]/u;
```

- **On the string, not the tree.** One scan covers attributes, text, and whatever stays DOM: a flat OPC part the model
  doesn't know, a branch's untyped content.
- **The form of the expression.** core-ts's uses a lookbehind for the low surrogate. This one is the production itself
  and needs none. Safari shipped lookbehind in 16.4, Office on a Mac runs add-ins in its engine, and a regular expression
  literal an engine can't parse fails the whole module. That reasoning is from documentation, not measured here.
- **The error.** A `RangeError`, as the builders throw for a width Word can't hold: `The XML holds U+000B, which XML 1.0
  excludes, so it can't be written as well-formed XML (character 107, after "..."); remove it, or write it as the element
  it stands for`. From `marshalPackage` it opens `Part /word/document.xml holds`, the name being the nearest `pkg:name`
  before the character.
- **The option.** `MarshalOptions.checkCharacters`, default `true`. `false` writes the string as held (D2).
- **Unchanged:** `marshalNode` (D3), `r()`, `p()`, every read, and what a save writes for anything else. Nothing is
  dropped or replaced.

A read report for a file holding such a character isn't in this CR. Word refuses the file, but reporting it is a
separate question from writing one.

## 5. Results

- `npm run typecheck` and `npm test` pass on Node 18.20.8, 20.20.2 and 22.23.4, with jsonix 3.11.0.
- All 341 marshalled hashes are identical to `4f84e1e`'s on each Node. The fixtures hold no such character, raw or as a
  reference.
- The expression agrees with core-ts's on 65,542 strings: every BMP code unit between letters, and the surrogate pairings.
- The scan takes about 10 ms over 10.8 million characters.

core-ts and docx4j-ts-editor have not been run against it.

## 6. Decisions

All three decided as recommended, 2026-11-07.

| D | Question | Decided |
|---|---|---|
| D1 | Refuse, or report through a callback and write the string anyway, as `onOutOfRange` and `onMisplaced` do? | Refuse. Those report a value that is still a file. This output can't be parsed, and core-ts and python already refuse. |
| D2 | A way round it? | Yes, per call: `checkCharacters: false` on `marshalString`, for callers using it as an internal round trip (section 3). The repair policy's "no option" is about a save, and `marshalString` isn't always one. `marshalPackage` takes no options and gets none. |
| D3 | Check `marshalNode` too? | No. It writes no string, and core-ts's guard follows it. Checking it would move core-ts's refusal away from the part save that names the part. |

## 7. Tests and docs

- **Smoke, a CR-020 block:** each class of excluded character in a `w:t`, and U+000B in an attribute, read from a file then
  written, and in a typed part and a DOM part of a flat OPC package. The allowed neighbours (tab, line feed, carriage
  return, U+007F, U+0085, U+00A0, the edges of the ranges, U+FFFD, a surrogate pair). The message's character, position and
  part. `checkCharacters: false`. `marshalNode` and `wml()`'s interpolation unchanged.
- **The gate:** every fidelity hash identical.
- **Docs:** `CLAUDE.md`'s smoke paragraph, and a README bullet by "What a save writes".

## 8. Consumers

- **core-ts:** the four routes in section 3 gain the refusal. Its part saves don't change. Release gate as usual.
- **docx4j-ts-editor:** eleven calls of `marshalString`, by its count. `session.mts:558` and `1197` and `stories.mts:355`, `811`
  and `864` build copies and source entries, and could throw mid-edit. It will opt out at copies and displays and keep the
  check on a real save. Filed against its install of the release, not yet run.
- **An add-in on this package alone:** a save that wrote a bad part now throws. That is the point of the CR, and it is
  still a new throw.
