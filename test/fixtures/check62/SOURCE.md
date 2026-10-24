# Check 62: the highlight Word stores for a colour written to it (CR-014)

`62-highlight-snap-word365.json` is core-ts's check 62, run by Jason Harrop in Word 365 (16.0.20326.20232, PC, no
repair) on 24 October 2026. Each point records one colour written to `font.highlightColor`, `[written, read back]`;
75 points also have `{ stored }`, the `w:highlight` `getOoxml()` gave. The 7,225 colours are:
- the 16-level grid;
- the grey ramp;
- each channel's ramp over `00` and over `FF`;
- the colours around `#A9A9A9` and `#123456`.

`CHECK62_FAMILIES` in `build.mjs` there lists them.

Copied byte for byte from `plutext/docx4j-core-ts` at 8a0945c, `test/fixtures/check62/62-highlight-snap-word365.json`.
core-ts's CR-015 section 12.2 records the run. Never edited by hand.
