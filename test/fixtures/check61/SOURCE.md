# Check 61: Word's highlight palette through Office JS (CR-014)

`61-highlights-word365.json` is core-ts's check 61, run by Jason Harrop in Word 365 (16.0.20326.20232, PC, no repair)
on 24 October 2026. It records 18 reads, `font.highlightColor` of a run with each `w:highlight` (and of one with a
shading alone). It also records 45 writes: each value given to `font.highlightColor`, what Word read back, and the
`w:highlight` Word stored, from `getOoxml()`.

Copied byte for byte from `plutext/docx4j-core-ts` at dc7e37d, `test/fixtures/check61/61-highlights-word365.json`,
where `build.mjs` makes the document Word opened and `check61-script-lab.js` is the snippet that measured it. core-ts's
CR-015 section 12 records the run. Never edited by hand.
