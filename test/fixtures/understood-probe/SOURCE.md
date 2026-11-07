# Word's reading of an `mc:Choice` with no Fallback, by namespace (CR-021)

Each case is a `word/document.xml` (`<id>.document.xml`) beside the text Word saved from it (`<id>.word.txt`), never edited.
Word 16.0.20326.20232, through the Word bridge, 6 and 7 November 2026. Each file was opened twice, silently, and the two runs
saved the same text. Copied from the share's `bridge/portfolio/<id>-a`, with `-b` compared.

| Round | Cases | What it holds |
|---|---|---|
| `ul-*` | 119 | One `Choice`-only file for each of the 118 URIs on docx4j's old list, `w14` among them as a control, and a made-up namespace. Each `Requires` uses the prefix `q`. |
| `uc-*` | 13 | The namespaces that list lacked, and two controls. |
| `c305-*` | 9 | Block and run level, `wpi`, `oel`, `cx3`, `w16sdtfl` and unknown namespaces. |
| `rw-25` | 1 | An unknown namespace in a body, no Fallback. |

These are the measurements behind docx4j's `understood-namespaces.txt` (WD REQ-305, REQ-306). Every case is `Choice`-only, in a
main document part, so they say nothing of what Word does with a Fallback, or of spreadsheets and presentations.
