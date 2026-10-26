# CR-014: The highlight palette as Word and Office JS read it (`HIGHLIGHT_COLORS`, `readRunOptions`, `applyRunOptions`)

**Status:** Implemented 2026-10-24 (Jason: "Implement CR-014"), released in 0.6.0 (2026-10-26); H1 to H4 decided as recommended (Jason, 2026-10-24, relayed by core-ts)
**Requested by:** core-ts (session docx4j-core-ts-5b), 2026-10-24. Its check 60 found the one difference between its
Word shim and Office JS in this package's highlight table (core-ts CR-015 section 11.6). Its check 61, written at
Jason's word ("write check 61") and run by him the same day, measured the palette both ways (CR-015 section 12).
**Depends on:** nothing. The data is check 61's.
**Counterpart:** docx4j's `Highlight.getHexVal()` and `setVal()` (injected by `ModifyGeneratedSources`). docx4j took
Word's palette in a3cfc9789 (17.3.5, after core-ts told docx4j-25 of check 61), so the two tables agree again. Its
`setVal("#RRGGBB")` takes the palette and its former CSS values, not Office JS's input (section 4).

## 1. Summary

`helpers/wml`'s `HIGHLIGHT_COLORS` was docx4j's `Highlight` table: CSS colour values for the sixteen
`ST_HighlightColor` names, with gold for `darkYellow`, which CSS lacks. Word's own palette differs in nine of them.
Office JS reports Word's palette, and `readRunOptions` and `applyRunOptions` speak Office JS's vocabulary (`Word.Font`),
so `readRunOptions` read nine highlights in colours Word doesn't use. And `applyRunOptions` took a highlight *name*
where Office JS takes a CSS colour *name*, so `'green'` wrote bright green where Office JS writes dark green.

The table now holds Word's palette. `applyRunOptions` takes what Office JS takes, through `highlightNameForColor`, and
stores what Word stores. That is exact for every input checks 61 and 62 measured, and follows the rule check 62 fits
elsewhere (H4 (c)).

## 2. What check 61 measured (Word 365 16.0.20326.20232, PC, 2026-10-24; no repair)

The data is core-ts's `test/fixtures/check61/61-highlights-word365.json`: 18 reads and 45 writes.

**Reads** (`font.highlightColor` of a run with each `w:highlight`):

| Name | Word (Office JS) | This package (docx4j) |
|---|---|---|
| `black`, `blue`, `cyan`, `magenta`, `red`, `yellow`, `white` | `000000`, `0000FF`, `00FFFF`, `FF00FF`, `FF0000`, `FFFF00`, `FFFFFF` | the same |
| `green` | `00FF00` | `008000` |
| `darkBlue` | `000080` | `00008B` |
| `darkCyan` | `008080` | `008B8B` |
| `darkGreen` | `008000` | `006400` |
| `darkMagenta` | `800080` | `8B008B` |
| `darkRed` | `800000` | `8B0000` |
| `darkYellow` | `808000` | `FFD700` |
| `darkGray` | `808080` | `A9A9A9` |
| `lightGray` | `C0C0C0` | `D3D3D3` |

`none`, and a character shading (`w:shd w:fill="FFFF00"`) with no highlight, read `null`. Word's palette is the one
usually given for `WdColorIndex`.

**Writes** (`font.highlightColor = value`; Word's stored `w:highlight` taken from `getOoxml()`):
- **A name is a CSS colour name, in any case,** not a `w:highlight` name, and its colour is stored as below. `green`
  is CSS's `#008000`, so Word stores `darkGreen`. `Yellow` gives `yellow`. `darkYellow`, which CSS lacks, is refused
  (`InvalidArgument`). `lightGray` is accepted, a spelling Word's `w:color` reading refuses (CR-013 section 3.4), so
  Office JS's list is CSS's own.
- **A hex, in any case, is always stored as a highlight**, never as a shading.
  - Each of the 16 palette values stores its own name.
  - Any other value snaps to a palette entry. CSS's values of the seven dark and grey names (`#00008B` and so on) snap
    to the same names, `#FFD700` and `#FF8000` to `yellow`, and `#123456` to `darkCyan`.
  - The snap is not the nearest in RGB: `#A9A9A9` goes to `darkGray` though `lightGray` is nearer, and `#123456` to
    `darkCyan` though `darkBlue` is nearer. No simple rule fits all ten: nearest RGB (euclidean, Manhattan,
    max-channel), luma-weighted RGB, HLS distance and per-channel thresholds to `00`, `80` and `FF` each get 8 at best,
    always missing those two (tried 2026-10-24).
- **`null`** removes the highlight.

Over the 63 cases, 0.5.0's `readRunOptions`/`applyRunOptions` differ from Word in 39; core-ts pins those 39.

### 2.1 Check 62: Word's snap, measured (Jason, 2026-10-24)

Jason asked why the full hex range shouldn't be allowed, and asked core-ts for a grid "so we can see whether there is
a simple rule". Check 62 (core-ts `test/fixtures/check62/`, the same Word 365) wrote 7,225 colours:
- the 16-level grid (every CSS three-digit hex colour);
- the grey ramp in steps of 1;
- each single-channel ramp, with the other two at `00` and at `FF`;
- ±8 around `#A9A9A9` and `#123456`.

Every colour was accepted, and each was stored as one of the sixteen. The 75 names spot-checked through `getOoxml`
agree with the colour read back.

The rule, as far as it is simple, was fitted on 2026-10-24. It agrees with Word on 7,075 of the 7,225 colours, and on
7,076 with `#A9A9A9`'s exception; the 149 left are all at the grey edge.
- **Hue** is Windows' `ColorRGBToHLS` hue, an integer from 0 to 239. The sectors are red below 20, yellow 20 to 60,
  green 61 to 100, cyan 101 to 140, blue 141 to 179, and magenta 180 and above. So magenta takes 270° to 360°, and
  `#FF0011` is magenta. Every hue decision in the data fits.
- **Dark or bright:** dark where `max + min ≤ C0` (HLS lightness at most 90 of 240), else bright. So `#010000` and
  `#C00000` are `darkRed`, and `#C10000` is `red`. Every decision fits.
- **Greys:** by `(max + min) / 2`. `00` to `3F` is `black`, `40` to `9F` `darkGray`, `A0` to `DF` `lightGray`, and
  `E0` to `FF` `white`, the midpoints between the palette greys. The one exception is `#A9A9A9`, CSS's `darkgray`,
  alone `darkGray`: Word appears to match a CSS name before any rule.
- **Grey or colour:** not simple. HLS saturation below 120 of 240 gives a grey, and that is the best single test, but
  149 points still differ. No measure tried separates them exactly: HLS and HSV saturation, chroma, Lab chroma and
  chroma over L, YCbCr chroma. The boundary bends with hue (a blue needs more chroma than a red to stay a colour) and
  with lightness. All 149 are low-saturation colours, which Word puts one side or the other of the grey edge.

## 3. Proposal

### 3.1 `HIGHLIGHT_COLORS` takes Word's palette

The nine values change to Word's (section 2's table). The names, the order and the types are unchanged.
`highlightHexValue` (docx4j's `getHexVal` counterpart) and `readRunOptions` read as Office JS reads. The comment "as in
docx4j (darkYellow uses gold)" now names Word's palette (`WdColorIndex`) and check 61. docx4j's table has the same
values since a3cfc9789.

### 3.2 `applyRunOptions` takes what Office JS takes

`highlightNameForColor(color)`, which `applyRunOptions` calls, takes Office JS's input (H2) and gives the name Word
stores:
- **A CSS colour name, any case:** CSS Color 3's 147 names. That is `WORD_COLOR_NAMES` plus the seven spellings Word's
  `w:color` reading refuses. Only `lightgray` among those seven is measured; the rest follow CSS (H3). The name gives
  its CSS hex, which then follows the hex rule.
- **A hex, `#RRGGBB` in any case,** trimmed. A palette value gives its own name. Any other value gives what Word snaps
  it to (H4 (c)):
  - for the 149 grey-edge colours the rule misses, check 62's answer, held as data;
  - for `#A9A9A9`, `darkGray`;
  - for every other colour, section 2.1's rule.

  That is Word's answer for every colour checks 61 and 62 measured.
- **Anything else is refused** (undefined; `applyRunOptions` throws, as before). That covers `darkYellow`, as Office
  JS refuses it, and the forms in section 4.
- `null` and `''` still remove the highlight.

`readRunOptions` and `applyRunOptions` agree with Word in all 63 of check 61's cases, and `highlightNameForColor` on all
7,225 of check 62's colours.

### 3.3 Not changed

- `readRunOptions` still reads a `w:highlight` outside `ST_HighlightColor`, and `none`, as `null` (CR-013 section 3.3).
- `w:highlight` stays a string in the model, so a document's values load as before. Only these readings and the
  builder's writes change.
- `WORD_COLOR_NAMES` stays Word's `w:color` list: 140 names, for reading documents. Office JS's input list is a
  separate, larger one.

## 4. Not proposed

- **Word's exact grey-or-colour test** for a low-saturation hex (section 2.1). H4 (c) uses check 62's answers where
  measured and the nearest simple rule between.
- **Office JS's other colour forms**, such as `rgb()`, a hex without `#`, a three-digit hex or a padded name. They are
  unmeasured, so they stay refused.
- **docx4j's input:** docx4j's `getHexVal()`, and its HTML and PDF, have Word's palette since a3cfc9789. Its
  `setVal("#RRGGBB")` takes the palette and its former CSS values only, not CSS names or the snap. That is docx4j's
  to change.

## 5. Decisions

**Decided as recommended** (Jason, 2026-10-24, in the core-ts session: "accept recs for decisions H1 to H4", relayed by
core-ts): H1 one table with Word's palette; H2 Office JS's input; H3 CSS's 147 names; H4 (c), section 2.1's rule with
check 62's 149 grey-edge answers as data, and no check 63 unless a consumer needs it. The options and reasons stand
below.

- **H1. One table, Word's values.** `HIGHLIGHT_COLORS` (and so `highlightHexValue`) takes Word's palette, rather than
  keeping docx4j's beside a second, Office JS table. **Recommendation:** one table. Word is the authority on what a
  `w:highlight` name means, docx4j's CSS values are a defect there too, and two tables would disagree forever.
- **H2. Office JS's input for `applyRunOptions` and `highlightNameForColor`.** CSS names in any case, with
  `green` giving `darkGreen` and `darkYellow` refused. The alternative keeps 0.5.0's highlight names (and adds hex
  case-insensitivity). **Recommendation:** Office JS's input, since the run mapping exists to speak `Word.Font`'s
  vocabulary, and core-ts's `Font` view is built on it. It changes behaviour for `'green'` (bright green becomes dark
  green, as in Office JS) and `'darkYellow'` (refused, as in Office JS; `'#808000'` sets it).
- **H3. The name list:** CSS's 147, measured only as far as `lightgray`; or `WORD_COLOR_NAMES` plus `lightgray` alone.
  **Recommendation:** CSS's 147. Office JS is a web API, and the one extra spelling measured is CSS's.
- **H4. A hex outside the palette,** which Office JS accepts and Word snaps to one of the sixteen (section 2.1):
  - (a) refuse it, except check 61's ten measured points;
  - (b) accept every hex by section 2.1's rule, with `#A9A9A9`'s exception. It agrees with Word on 7,076 of
    check 62's 7,225 points; the 149 that differ are low-saturation colours at the grey edge;
  - (c) as (b), with check 62's 149 grey-edge differences as a table of Word's answers. It agrees on every measured
    point, which includes all 4,096 CSS three-digit colours, and uses the rule between them;
  - (d) measure the grey edge more finely (a check 63), in the hope of a rule.

  **Recommendation:** (c). It's exact wherever measured, and elsewhere wrong only for a low-saturation colour near
  the grey edge, where Word's own choice looks arbitrary. (d) only if that edge matters to a consumer.

## 6. Tests

- **Fixtures:** check 61's and check 62's JSON are copied, byte for byte, to `test/fixtures/check61/` and `check62/`.
  Each has a `SOURCE.md` naming core-ts's file and commit (dc7e37d and 8a0945c).
- **Reads:** the smoke's CR-014 block holds `HIGHLIGHT_COLORS`, `highlightHexValue` and `readRunOptions` to check 61's
  18 reads.
- **Writes:** it holds `highlightNameForColor` and `applyRunOptions` to check 61's 45 writes: the stored name, the
  read-back, `InvalidArgument` as a throw, and `null` as removal.
- **The snap:** it holds `highlightNameForColor` to all 7,225 of check 62's colours, in either case, and nine
  unmeasured ones to the rule.
- **Names:** each of CSS's 147 names, in three cases, is stored as its colour is. A name CSS lacks (`darkYellow`,
  `none`, CSS 4's `rebeccapurple`) and the forms of section 4 are refused.
- **Probes:** `test/word365.mjs`'s highlight readings use Word's values. Its WORD_DREW entries for M6 and H7 (a green
  highlight) became `#00FF00`, as docx4j's did in a3cfc9789.
- `test/nodenext/consumer.mts` is unchanged; no export is added or removed.

## 7. Consumers

- **core-ts:** its `Font` shim's 39 pinned check 61 cases, and its check 62 pins, become agreements; the pins
  announce the release. Its own catalogue's highlight set (`catalogue.mts`) holds names only, so it needs no change.
- **Office JS add-ins using this package alone:** `readRunOptions` reads as Office JS does, and `applyRunOptions` takes
  what `font.highlightColor` takes.
- **docx4j:** none from this CR; it has its counterpart to decide.

## 8. Release

A change of readings and of the builder's accepted input:
- `'green'` writes `darkGreen`;
- `'darkYellow'` throws;
- every `#RRGGBB` is accepted;
- nine values read differently.

Under Jason's rule (minors for a substantive change of API or output), a minor: released in 0.6.0 (2026-10-26). No API
is removed.
