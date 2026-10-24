// Hand-written helpers next to the generated WordprocessingML declarations (CR-007, item 5).
// They carry the docx4j decisions that its ModifyGeneratedSources adds to the generated Java:
// the w:highlight colour table (Highlight.getHexVal; Word's palette, CR-014), the LsdException.isQFormat()
// fallback to the latent styles' defQFormat (needs parent pointers, jsonix-CR-002), and the
// Style.isCustomStyle() rule (docx4j issue 641); a colour value as Word 365 reads it, with
// the names it reads (CR-013, docx4j's WordColorValues); and the highlight Word 365 stores for a
// colour Office JS writes (CR-014). Compiled to dist/helpers/wml.mjs by `npm run build`.
import type { Highlight, Style, Styles } from '../../modules/org_docx4j_wml.mjs';

/**
 * w:highlight names and their sRGB values: Word's palette (`WdColorIndex`), as Word 365 draws a highlight and
 * Office JS's `font.highlightColor` reads it (CR-014, core-ts check 61). docx4j's `Highlight` has the same since
 * 17.3.5; before, both had CSS's colours of the same names, and gold for darkYellow.
 */
export const HIGHLIGHT_COLORS: ReadonlyArray<readonly [name: string, hex: string]> = [
  ['black', '000000'], ['blue', '0000FF'], ['cyan', '00FFFF'], ['green', '00FF00'],
  ['magenta', 'FF00FF'], ['red', 'FF0000'], ['yellow', 'FFFF00'], ['white', 'FFFFFF'],
  ['darkBlue', '000080'], ['darkCyan', '008080'], ['darkGreen', '008000'], ['darkMagenta', '800080'],
  ['darkRed', '800000'], ['darkYellow', '808000'], ['darkGray', '808080'], ['lightGray', 'C0C0C0'],
];

/** The `#RRGGBB` value of a highlight, or undefined for an unknown name (docx4j `Highlight.getHexVal()`). */
export function highlightHexValue(highlight: Pick<Highlight, 'val'> | string | undefined): string | undefined {
  const val = typeof highlight === 'string' ? highlight : highlight?.val;
  const entry = HIGHLIGHT_COLORS.find(([name]) => name === val);
  return entry ? '#' + entry[1] : undefined;
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
 * CSS Color 3's 147 names, lower-case, to RRGGBB: what Office JS's `font.highlightColor` takes as a name (CR-014 H3).
 * Word's 140 (`WORD_COLOR_NAMES`) and the seven spellings its `w:color` reading refuses.
 */
const CSS_COLOR_NAMES: Readonly<Record<string, string>> = Object.freeze(Object.assign(Object.create(null) as Record<string, string>, WORD_COLOR_NAMES, {
  darkgrey: 'A9A9A9', darkslategrey: '2F4F4F', dimgrey: '696969', grey: '808080', lightgray: 'D3D3D3',
  lightslategrey: '778899', slategrey: '708090',
}));

/**
 * Windows' `ColorRGBToHLS` (shlwapi) hue and saturation, integers out of 240, by which Word snaps a colour to a
 * highlight: hue 0 is red, 40 yellow, 80 green, 120 cyan, 160 blue and 200 magenta.
 */
function hueAndSaturation(r: number, g: number, b: number): [hue: number, saturation: number] {
  const max = Math.max(r, g, b), min = Math.min(r, g, b), sum = max + min, diff = max - min;
  if (diff === 0) return [160, 0];
  const lightness = Math.trunc((sum * 240 + 255) / 510);
  const saturation = lightness <= 120
    ? Math.trunc((diff * 240 + (sum >> 1)) / sum)
    : Math.trunc((diff * 240 + ((510 - sum) >> 1)) / (510 - sum));
  const part = (c: number) => Math.trunc(((max - c) * 40 + (diff >> 1)) / diff);
  const hue = r === max ? part(b) - part(g) : g === max ? 80 + part(r) - part(b) : 160 + part(g) - part(r);
  return [hue < 0 ? hue + 240 : hue, saturation];
}

/** The hue sectors, each below its bound, with their bright and dark highlights (CR-014 section 2.1). */
const HUE_SECTORS: ReadonlyArray<readonly [below: number, bright: string, dark: string]> = [
  [20, 'red', 'darkRed'], [61, 'yellow', 'darkYellow'], [101, 'green', 'darkGreen'],
  [141, 'cyan', 'darkCyan'], [180, 'blue', 'darkBlue'], [240, 'magenta', 'darkMagenta'],
];

/**
 * Word's answer for the 149 colours of core-ts's check 62 that `snappedHighlight`'s rule puts on the other side of
 * the grey edge (CR-014 H4 (c)). All are low-saturation colours of the 16-level grid; no measure tried separates them.
 */
const GREY_EDGE: ReadonlyMap<string, string> = new Map(Object.entries({
  darkGray: `333399 334499 335599 336699 443399 553399 663399 6666CC 6677CC 6688CC 6699CC 7766CC 8866CC 9966CC`,
  lightGray: `9999DD 99AADD 99BBDD 99CCDD 99DDCC 99DDDD AA99DD BB99DD CC99DD CCCCEE CCDDEE CCEECC CCEEDD CCEEEE DD99CC
    DD99DD DDCCEE EECCDD EECCEE`,
  red: `884444 885544 994444 995544 995555 996644 AA4444 AA5544 AA5555 AA6644 AA6655 AA6666 AA7755 BB4444 BB5544
    BB5555 BB6644 BB6655 BB6666 BB7744 BB7755 BB7766 BB7777 BB8866 BB8877 CC7777 CC8877 CC8888 CC9977 CC9988
    CC9999 CCAA99 DDAAAA DDBBAA DDBBBB`,
  yellow: `77AA44 88AA44 88BB44 88BB55 997744 998844 999944 99AA44 99BB44 99BB55 AA7744 AA8844 AA9944 AAAA44 AABB44
    AABB55 AACC77 BB8844 BB8855 BB9944 BB9955 BB9966 BBAA44 BBAA55 BBAA66 BBBB44 BBBB55 BBBB66 BBCC77 CCAA77
    CCAA88 CCBB77 CCBB88 CCCC77 CCCC88 DDCCAA DDDDAA`,
  green: `44BB44 44BB55 44BB66 55BB44 66AA44 66BB44 77BB44 77BB55 77CC77 88CC77 99CC77`,
  magenta: `BB4455 BB4466`,
  darkRed: `221111 442222 552222 553322 663333 664433 773333 774433 883333 884433 885533`,
  darkYellow: `443322 445522 554422 555522 665533 667733 668833 775533 776633 777733 778833 886633 887733 888833`,
  darkGreen: `225522 335522 338833 338844 448833 558833`,
}).flatMap(([name, colours]) => colours.split(/\s+/).map((hex): [string, string] => [hex, name])));

/**
 * The highlight Word 365 stores for a colour outside its palette, `RRGGBB` upper-case (CR-014 section 2.1, fitted to
 * core-ts's check 62, 7,225 colours). A grey, by HLS saturation below 120, is banded by `(max + min) / 2` at 40, A0
 * and E0; a colour is its hue's sector, dark where `max + min` is at most C0. `#A9A9A9`, CSS's `darkgray`, is alone
 * `darkGray`: Word appears to match a CSS name first. Exact on every colour check 62 measured, with `GREY_EDGE`.
 */
function snappedHighlight(hex: string): string {
  const measured = GREY_EDGE.get(hex);
  if (measured !== undefined) return measured;
  if (hex === 'A9A9A9') return 'darkGray';
  const r = parseInt(hex.substring(0, 2), 16), g = parseInt(hex.substring(2, 4), 16), b = parseInt(hex.substring(4), 16);
  const sum = Math.max(r, g, b) + Math.min(r, g, b);
  const [hue, saturation] = hueAndSaturation(r, g, b);
  if (saturation < 120) {
    const lightness = sum >> 1;
    return lightness < 0x40 ? 'black' : lightness < 0xA0 ? 'darkGray' : lightness < 0xE0 ? 'lightGray' : 'white';
  }
  const [, bright, dark] = HUE_SECTORS.find(([below]) => hue < below) ?? HUE_SECTORS[HUE_SECTORS.length - 1];
  return sum <= 0xC0 ? dark : bright;
}

/**
 * The highlight Word 365 stores for a colour written to Office JS's `font.highlightColor` (CR-014 section 3.2; core-ts
 * checks 61 and 62): a CSS colour name in any case (`green` is CSS's `#008000`, so `darkGreen`), or `#RRGGBB` in any
 * case, trimmed. A palette value gives its own name, and any other colour the one of the sixteen Word snaps it to
 * (`snappedHighlight`). Anything else is undefined: `darkYellow`, which is no CSS name and which Office JS refuses,
 * a hex without its `#`, `rgb()`. `applyRunOptions` stores this name.
 */
export function highlightNameForColor(color: string | undefined): string | undefined {
  if (color === undefined) return undefined;
  const trimmed = color.trim();
  const hex = CSS_COLOR_NAMES[color.toLowerCase()]
    ?? (/^#[0-9A-Fa-f]{6}$/.test(trimmed) ? trimmed.substring(1).toUpperCase() : undefined);
  if (hex === undefined) return undefined;
  return HIGHLIGHT_COLORS.find(([, value]) => value === hex)?.[0] ?? snappedHighlight(hex);
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
