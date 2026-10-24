// Hand-written helpers next to the generated WordprocessingML declarations (CR-007, item 5).
// They carry the docx4j decisions that its ModifyGeneratedSources adds to the generated Java:
// the w:highlight colour table (Highlight.setVal / getHexVal), the LsdException.isQFormat()
// fallback to the latent styles' defQFormat (needs parent pointers, jsonix-CR-002), and the
// Style.isCustomStyle() rule (docx4j issue 641); and a colour value as Word 365 reads it, with
// the names it reads (CR-013, docx4j's WordColorValues). Compiled to dist/helpers/wml.mjs by `npm run build`.
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
 * The classic HTML (Netscape) colour names Word 365 reads in a colour value, lower-case, to RRGGBB:
 * CSS Color 3's 147 keywords less `grey`, `darkgrey`, `darkslategrey`, `dimgrey`, `lightslategrey`,
 * `slategrey` and `lightgray` (Word spells the grays with an a, but `lightgrey` with an e). 140
 * names, copied from docx4j's `WordColorValues` (CR-013 section 3.4), which the smoke holds them to.
 * Not DrawingML's preset colours, which spell and value some names differently.
 */
export const WORD_COLOR_NAMES: Readonly<Record<string, string>> = Object.freeze(Object.assign(Object.create(null) as Record<string, string>, {
  aliceblue: 'F0F8FF', antiquewhite: 'FAEBD7', aqua: '00FFFF', aquamarine: '7FFFD4', azure: 'F0FFFF',
  beige: 'F5F5DC', bisque: 'FFE4C4', black: '000000', blanchedalmond: 'FFEBCD', blue: '0000FF', blueviolet: '8A2BE2',
  brown: 'A52A2A', burlywood: 'DEB887', cadetblue: '5F9EA0', chartreuse: '7FFF00', chocolate: 'D2691E',
  coral: 'FF7F50', cornflowerblue: '6495ED', cornsilk: 'FFF8DC', crimson: 'DC143C', cyan: '00FFFF',
  darkblue: '00008B', darkcyan: '008B8B', darkgoldenrod: 'B8860B', darkgray: 'A9A9A9', darkgreen: '006400',
  darkkhaki: 'BDB76B', darkmagenta: '8B008B', darkolivegreen: '556B2F', darkorange: 'FF8C00', darkorchid: '9932CC',
  darkred: '8B0000', darksalmon: 'E9967A', darkseagreen: '8FBC8F', darkslateblue: '483D8B', darkslategray: '2F4F4F',
  darkturquoise: '00CED1', darkviolet: '9400D3', deeppink: 'FF1493', deepskyblue: '00BFFF', dimgray: '696969',
  dodgerblue: '1E90FF', firebrick: 'B22222', floralwhite: 'FFFAF0', forestgreen: '228B22', fuchsia: 'FF00FF',
  gainsboro: 'DCDCDC', ghostwhite: 'F8F8FF', gold: 'FFD700', goldenrod: 'DAA520', gray: '808080', green: '008000',
  greenyellow: 'ADFF2F', honeydew: 'F0FFF0', hotpink: 'FF69B4', indianred: 'CD5C5C', indigo: '4B0082',
  ivory: 'FFFFF0', khaki: 'F0E68C', lavender: 'E6E6FA', lavenderblush: 'FFF0F5', lawngreen: '7CFC00',
  lemonchiffon: 'FFFACD', lightblue: 'ADD8E6', lightcoral: 'F08080', lightcyan: 'E0FFFF',
  lightgoldenrodyellow: 'FAFAD2', lightgreen: '90EE90', lightgrey: 'D3D3D3', lightpink: 'FFB6C1',
  lightsalmon: 'FFA07A', lightseagreen: '20B2AA', lightskyblue: '87CEFA', lightslategray: '778899',
  lightsteelblue: 'B0C4DE', lightyellow: 'FFFFE0', lime: '00FF00', limegreen: '32CD32', linen: 'FAF0E6',
  magenta: 'FF00FF', maroon: '800000', mediumaquamarine: '66CDAA', mediumblue: '0000CD', mediumorchid: 'BA55D3',
  mediumpurple: '9370DB', mediumseagreen: '3CB371', mediumslateblue: '7B68EE', mediumspringgreen: '00FA9A',
  mediumturquoise: '48D1CC', mediumvioletred: 'C71585', midnightblue: '191970', mintcream: 'F5FFFA',
  mistyrose: 'FFE4E1', moccasin: 'FFE4B5', navajowhite: 'FFDEAD', navy: '000080', oldlace: 'FDF5E6', olive: '808000',
  olivedrab: '6B8E23', orange: 'FFA500', orangered: 'FF4500', orchid: 'DA70D6', palegoldenrod: 'EEE8AA',
  palegreen: '98FB98', paleturquoise: 'AFEEEE', palevioletred: 'DB7093', papayawhip: 'FFEFD5', peachpuff: 'FFDAB9',
  peru: 'CD853F', pink: 'FFC0CB', plum: 'DDA0DD', powderblue: 'B0E0E6', purple: '800080', red: 'FF0000',
  rosybrown: 'BC8F8F', royalblue: '4169E1', saddlebrown: '8B4513', salmon: 'FA8072', sandybrown: 'F4A460',
  seagreen: '2E8B57', seashell: 'FFF5EE', sienna: 'A0522D', silver: 'C0C0C0', skyblue: '87CEEB', slateblue: '6A5ACD',
  slategray: '708090', snow: 'FFFAFA', springgreen: '00FF7F', steelblue: '4682B4', tan: 'D2B48C', teal: '008080',
  thistle: 'D8BFD8', tomato: 'FF6347', turquoise: '40E0D0', violet: 'EE82EE', wheat: 'F5DEB3', white: 'FFFFFF',
  whitesmoke: 'F5F5F5', yellow: 'FFFF00', yellowgreen: '9ACD32',
}));

/** Java's String.trim, as docx4j's WordColorValues.normalise trims: characters up to U+0020 at either end. */
function trimAsJava(value: string): string {
  return value.replace(/^[\u0000-\u0020]+|[\u0000-\u0020]+$/g, '');
}

/**
 * A colour value (`w:color/@w:val`, `w:u/@w:color`, `w:shd/@w:fill`, a border's `w:color`) as
 * Word 365 reads it (CR-013 section 3.4): exactly `auto` gives `'auto'`, untrimmed and
 * case-sensitive; six hex digits in either case, with an optional `#` and surrounding whitespace
 * trimmed, give `'#RRGGBB'`, upper-case; a name of `WORD_COLOR_NAMES`, trimmed and in any case,
 * gives its value. Anything else (`Auto`, `F00`, `grey`, `""`) gives undefined, which Word reads as
 * absent in a text or underline colour and as automatic in a fill or a border colour, so the caller
 * applies its own place's rule. docx4j's `WordColorValues.normalise` and `stated`.
 */
export function wordColor(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  if (value === 'auto') return 'auto';
  const trimmed = trimAsJava(value);
  const hex = trimmed.startsWith('#') ? trimmed.substring(1) : trimmed;
  if (/^[0-9A-Fa-f]{6}$/.test(hex)) return '#' + hex.toUpperCase();
  const named = WORD_COLOR_NAMES[trimmed.toLowerCase()];
  return named === undefined ? undefined : '#' + named;
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
