// Hand-written helpers next to the generated WordprocessingML declarations (CR-007, item 5).
// They carry the docx4j decisions that its ModifyGeneratedSources adds to the generated Java:
// the w:highlight colour table (Highlight.setVal / getHexVal), the LsdException.isQFormat()
// fallback to the latent styles' defQFormat (needs parent pointers, jsonix-CR-002), and the
// Style.isCustomStyle() rule (docx4j issue 641). Compiled to dist/helpers/wml.mjs by `npm run build`.
import type { Highlight, Style, Styles } from '../../modules/org_docx4j_wml.mjs';

/** w:highlight names and their sRGB values, as in docx4j (darkYellow uses gold). */
export const HIGHLIGHT_COLORS: ReadonlyArray<readonly [name: string, hex: string]> = [
  ['black', '000000'], ['blue', '0000FF'], ['cyan', '00FFFF'], ['green', '008000'],
  ['magenta', 'FF00FF'], ['red', 'FF0000'], ['yellow', 'FFFF00'], ['white', 'FFFFFF'],
  ['darkBlue', '00008B'], ['darkCyan', '008B8B'], ['darkGreen', '006400'], ['darkMagenta', '8B008B'],
  ['darkRed', '8B0000'], ['darkYellow', 'FFD700'], ['darkGray', 'A9A9A9'], ['lightGray', 'D3D3D3'],
];

/** The `#RRGGBB` value of a highlight, or undefined for an unknown name (docx4j `Highlight.getHexVal()`). */
export function highlightHexValue(highlight: Pick<Highlight, 'val'> | string | undefined): string | undefined {
  const val = typeof highlight === 'string' ? highlight : highlight?.val;
  const entry = HIGHLIGHT_COLORS.find(([name]) => name === val);
  return entry ? '#' + entry[1] : undefined;
}

/**
 * The highlight name for a colour: a known name is returned as is, `#RRGGBB` is looked up in the
 * table; anything else (including rgb()) is undefined, as docx4j refuses it (`Highlight.setVal`).
 */
export function highlightNameForColor(color: string | undefined): string | undefined {
  if (color === undefined) return undefined;
  if (HIGHLIGHT_COLORS.some(([name]) => name === color)) return color;
  const trimmed = color.trim();
  if (trimmed.startsWith('#')) {
    const hex = trimmed.substring(1).toUpperCase();
    const entry = HIGHLIGHT_COLORS.find(([, value]) => value === hex);
    return entry ? entry[0] : undefined;
  }
  return undefined;
}

/**
 * Whether a latent style exception is a quick-format style, falling back to the latent styles'
 * `defQFormat` when the exception does not say (docx4j `LsdException.isQFormat()`). The fallback
 * requires `PARENT`, i.e. unmarshalling with `parentPointers: true`.
 */
export function isQFormat(exception: Styles.LatentStyles.LsdException): boolean {
  if (exception.qFormat !== undefined) return exception.qFormat;
  return exception.PARENT?.defQFormat ?? false;
}

/** An absent w:customStyle means a built-in style, whatever the schema default says (docx4j issue 641). */
export function isCustomStyle(style: Pick<Style, 'customStyle'>): boolean {
  return style.customStyle === true;
}
