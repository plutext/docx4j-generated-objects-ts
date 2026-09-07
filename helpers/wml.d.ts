import type { Highlight, Style, Styles } from '../org_docx4j_wml';
/** w:highlight names and their sRGB values, as in docx4j (darkYellow uses gold). */
export declare const HIGHLIGHT_COLORS: ReadonlyArray<readonly [name: string, hex: string]>;
/** The `#RRGGBB` value of a highlight, or undefined for an unknown name (docx4j `Highlight.getHexVal()`). */
export declare function highlightHexValue(highlight: Pick<Highlight, 'val'> | string | undefined): string | undefined;
/**
 * The highlight name for a colour: a known name is returned as is, `#RRGGBB` is looked up in the
 * table; anything else (including rgb()) is undefined, as docx4j refuses it (`Highlight.setVal`).
 */
export declare function highlightNameForColor(color: string | undefined): string | undefined;
/**
 * Whether a latent style exception is a quick-format style, falling back to the latent styles'
 * `defQFormat` when the exception does not say (docx4j `LsdException.isQFormat()`). The fallback
 * requires `PARENT`, i.e. unmarshalling with `parentPointers: true`.
 */
export declare function isQFormat(exception: Styles.LatentStyles.LsdException): boolean;
/** An absent w:customStyle means a built-in style, whatever the schema default says (docx4j issue 641). */
export declare function isCustomStyle(style: Pick<Style, 'customStyle'>): boolean;
