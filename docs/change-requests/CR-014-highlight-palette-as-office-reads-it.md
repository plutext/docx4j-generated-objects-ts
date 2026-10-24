# CR-014: The highlight palette as Word and Office JS read it (`HIGHLIGHT_COLORS`, `readRunOptions`, `applyRunOptions`)

**Status:** Proposed 2026-10-24; H1 to H4 for Jason
**Requested by:** core-ts (session docx4j-core-ts-5b), 2026-10-24. Its check 60 found the one difference between its
Word shim and Office JS in this package's highlight table (core-ts CR-015 section 11.6). Its check 61, written at
Jason's word ("write check 61") and run by him the same day, measured the palette both ways (CR-015 section 12).
**Depends on:** nothing. The data is check 61's.
**Counterpart:** docx4j's `Highlight.getHexVal()` and `setVal()` (injected by `ModifyGeneratedSources`), which have the
same table. core-ts told docx4j-25 on 2026-10-24. A docx4j CR is docx4j's to file, with Jason.

## 1. Summary

`helpers/wml`'s `HIGHLIGHT_COLORS` is docx4j's `Highlight` table: CSS colour values for the sixteen `ST_HighlightColor`
names, with gold for `darkYellow`, which CSS lacks. Word's own palette differs in nine of them. Office JS reports
Word's palette, and `readRunOptions` and `applyRunOptions` speak Office JS's vocabulary (`Word.Font`). So
`readRunOptions` reads nine highlights in colours Word doesn't use. And `applyRunOptions` takes a highlight *name*
where Office JS takes a CSS colour *name*, so `'green'` writes bright green where Office JS writes dark green.

Proposed: the table takes Word's palette. `applyRunOptions`, through `highlightNameForColor`, takes what Office JS
takes and stores what Word stores, for every input check 61 measured. A value Word would snap by a rule not yet
measured stays refused, as today.

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

Over the 63 cases, today's `readRunOptions`/`applyRunOptions` differ from Word in 39; core-ts pins those 39.

## 3. Proposal

### 3.1 `HIGHLIGHT_COLORS` takes Word's palette

The nine values change to Word's (section 2's table). The names, the order and the types are unchanged.
`highlightHexValue` (docx4j's `getHexVal` counterpart) and `readRunOptions` then read as Office JS reads. The
comment "as in docx4j (darkYellow uses gold)" becomes "Word's palette (`WdColorIndex`), as Office JS reads it; check
61". docx4j's table differs until docx4j follows.

### 3.2 `applyRunOptions` takes what Office JS takes

`highlightNameForColor(color)`, which `applyRunOptions` calls, takes Office JS's input (H2):
- **A CSS colour name, any case:** CSS Color 3's 147 names. That is `WORD_COLOR_NAMES` plus the seven spellings Word's
  `w:color` reading refuses. Only `lightgray` among those seven is measured; the rest follow CSS (H3). The name gives
  its CSS hex, which then follows the hex rule.
- **A hex, `#RRGGBB` in any case,** trimmed. A palette value gives its own name. One of the measured snaps gives
  Word's answer: the ten values in section 2 (H4).
- **Anything else is refused** (undefined; `applyRunOptions` throws, as now). That covers `darkYellow`, as Office JS
  refuses it, and any hex outside the palette and the measured points.
- `null` and `''` still remove the highlight.

Over check 61's 63 cases, the shim would then agree with Word in all 63.

### 3.3 Not changed

- `readRunOptions` still reads a `w:highlight` outside `ST_HighlightColor`, and `none`, as `null` (CR-013 section 3.3).
- `w:highlight` stays a string in the model, so a document's values load as before. Only these readings and the
  builder's writes change.
- `WORD_COLOR_NAMES` stays Word's `w:color` list: 140 names, for reading documents. Office JS's input list is a
  separate, larger one.

## 4. Not proposed

- **Word's snapping rule** for an arbitrary hex, until core-ts's check 62 measures it (H4). Until then such a value
  is refused, as today.
- **Office JS's other colour forms** (`rgb()`, a hex without `#`). They are unmeasured, so they stay refused.
- **docx4j's table:** its `getHexVal()`/`setVal()` are docx4j's to change. Its HTML and PDF draw the nine highlights in
  CSS's colours today, `darkYellow` gold where Word draws `#808000`.

## 5. Decisions for Jason

- **H1. One table, Word's values.** `HIGHLIGHT_COLORS` (and so `highlightHexValue`) takes Word's palette, rather than
  keeping docx4j's beside a second, Office JS table. **Recommendation:** one table. Word is the authority on what a
  `w:highlight` name means, docx4j's CSS values are a defect there too, and two tables would disagree forever.
- **H2. Office JS's input for `applyRunOptions` and `highlightNameForColor`.** CSS names in any case, with
  `green` giving `darkGreen` and `darkYellow` refused. The alternative keeps today's highlight names (and adds hex
  case-insensitivity). **Recommendation:** Office JS's input, since the run mapping exists to speak `Word.Font`'s
  vocabulary, and core-ts's `Font` view is built on it. It changes behaviour for `'green'` (bright green becomes dark
  green, as in Office JS) and `'darkYellow'` (refused, as in Office JS; `'#808000'` sets it).
- **H3. The name list:** CSS's 147, measured only as far as `lightgray`; or `WORD_COLOR_NAMES` plus `lightgray` alone.
  **Recommendation:** CSS's 147. Office JS is a web API, and the one extra spelling measured is CSS's.
- **H4. A hex outside the palette,** which Office JS accepts and Word snaps to one of the sixteen:
  - (a) the ten measured snaps as data, and every other value refused;
  - (b) every hex accepted, by Word's rule, once a grid check has measured it;
  - (c) every hex accepted now, by nearest RGB, an approximation that differs from Word for some greys and
    in-between colours.

  Refusing (a) is loud where a guess (c) would silently store another highlight than Office JS. Jason asked
  (2026-10-24) why the full hex range shouldn't be allowed, and asked core-ts to write **check 62**, a grid of Word's
  snaps, "so we can see whether there is a simple rule". **Recommendation:** (b). Use the rule if one fits every point
  of check 62. If none does, use the grid as a lookup table: exact on the grid, nearest grid point between. Until
  then, (a).

## 6. Tests

- check 61's JSON copied as a fixture (`test/fixtures/check61/`, with a `SOURCE.md` naming core-ts's file and commit).
- **Reads:** a smoke test holds `HIGHLIGHT_COLORS` and `readRunOptions` to check 61's 18 reads.
- **Writes:** a smoke test holds `highlightNameForColor` and `applyRunOptions` to its 45 writes: the stored name, the
  read-back, `InvalidArgument` as a throw, and `null` as removal.
- The probes (`test/word365.mjs`): the highlight readings now use Word's values. M6 and H7 (WORD_DREW `#008000`, docx4j's
  value for a green highlight) become `#00FF00`.
- `test/nodenext/consumer.mts` unchanged; no export is added or removed.

## 7. Consumers

- **core-ts:** its `Font` shim's 39 pinned cases become agreements, and the pins announce the release. Its own
  catalogue's highlight set (`catalogue.mts`) holds names only, so it needs no change.
- **Office JS add-ins using this package alone:** `readRunOptions` reads as Office JS does, and `applyRunOptions` takes
  what `font.highlightColor` takes.
- **docx4j:** none from this CR; it has its counterpart to decide.

## 8. Release

A change of readings and of the builder's accepted input: `'green'` writes `darkGreen`, `'darkYellow'` throws, and
nine values read differently. Under Jason's rule (minors for a substantive change of API or output), probably a
minor; Jason's call. No API is removed. Half a day, with the fixture and tests.
