// Hand-written helpers next to the generated WordprocessingML declarations (CR-007, item 5).
// They carry the docx4j decisions that its ModifyGeneratedSources adds to the generated Java:
// the w:highlight colour table (Highlight.getHexVal; Word's palette, CR-014), the LsdException.isQFormat()
// fallback to the latent styles' defQFormat (needs parent pointers, jsonix-CR-002), and the
// Style.isCustomStyle() rule (docx4j issue 641); a colour value as Word 365 reads it, with
// the names it reads (CR-013, docx4j's WordColorValues); the highlight Word 365 stores for a
// colour Office JS writes (CR-014); and how Word stores an integer attribute, and so what it reads for
// a value outside that (CR-016). Compiled to dist/helpers/wml.mjs by `npm run build`.
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
 * How Word 365 stores an integer attribute (CR-016; docx4j CR-066's rounds 5 and 6, Jason's runs of 2026-10-26). Word
 * reads a value modulo its storage width, sometimes then clamped, and neither refuses nor repairs the document. A
 * measured entry gives the width, the sign and any clamp or rule; an unmeasured one (`w:gridCol`, which Word
 * recomputes, and an object's original size, not probed) is known only to change outside -32768 to 65535, where every
 * measured twips attribute does.
 */
export type WordIntegerStorage =
  | {
    readonly measured: true;
    readonly bits: 5 | 8 | 16 | 32;
    readonly signed: boolean;
    /** After the wrap, Word clamps to these bounds: tab stops to ±31680, page width to 144..31681, page height to at least 144. */
    readonly min?: number;
    readonly max?: number;
    /** The attribute beside it whose rule the sign carries: `lineRule` (exact and atLeast) or `hRule`. */
    readonly signRule?: 'lineRule' | 'hRule';
    /** How Word's save writes the largest value: 65535 as -1 (`w:spacing` `before` and `after`, four page margins). */
    readonly maxWrittenAs?: number;
    /**
     * Two attributes sharing one signed field (`w:ind`'s `firstLine`, stored as itself, and `hanging`, stored negated):
     * the sign of what is stored says which Word reads.
     */
    readonly pair?: { readonly positive: string; readonly negative: string };
    /** `w:framePr/@w:h`: the low 15 bits are the height, and bit 15 clears `hRule="exact"`. */
    readonly heightFlag?: boolean;
    /** What a negative becomes after the wrap: a width Word computes (`w:col/@w:w`), or -1 (`w:col/@w:space`). */
    readonly negative?: 'computed' | -1;
  }
  | { readonly measured: false };

const INT16: WordIntegerStorage = { measured: true, bits: 16, signed: true };
const UINT16: WordIntegerStorage = { measured: true, bits: 16, signed: false };
/** uint16, with 65535 written -1. */
const UINT16_MINUS_ONE: WordIntegerStorage = { measured: true, bits: 16, signed: false, maxWrittenAs: -1 };
const FIRST_LINE_HANGING = { positive: 'firstLine', negative: 'hanging' } as const;
const UNMEASURED: WordIntegerStorage = { measured: false };

/**
 * Word's storage of each integer attribute of `wml` whose width is known or bounded, by `TYPE_NAME` and property (the
 * attribute's local name), from Word's saves of docx4j CR-066's rounds 5 and 6. A class's other uses store as the
 * measured one: `CTBorder` was measured on paragraph, run, cell, table and page borders, `TblWidth` on `w:tblW`,
 * `w:tcW`, `w:tblInd`, `w:tblCellSpacing` and the cell margins, `HpsMeasure` on `w:sz`, `w:szCs` and `w:kern`, and
 * `CTTwipsMeasure` on the settings' `w:defaultTabStop` and `w:hyphenationZone`. `w:ind`'s `start` and `end` are read
 * as `left` and `right`. Not here: ids, counts, the `*Chars` and `*Lines` attributes and w14's DrawingML units, whose
 * width is unmeasured; and a numbering level's `w:legacy`, which Word 365 drops whole, so it reads none of its values
 * (docx4j CR-066's table ignores it too; Jason, 2026-10-28).
 */
export const WORD_INTEGER_STORAGE: Readonly<Record<string, Readonly<Record<string, WordIntegerStorage>>>> = {
  'org_docx4j_wml.PPrBase.Ind': {
    left: INT16, right: INT16, start: INT16, end: INT16,
    firstLine: { ...INT16, pair: FIRST_LINE_HANGING }, hanging: { ...INT16, pair: FIRST_LINE_HANGING },
  },
  'org_docx4j_wml.PPrBase.Spacing': { before: UINT16_MINUS_ONE, after: UINT16_MINUS_ONE, line: { ...INT16, signRule: 'lineRule' } },
  'org_docx4j_wml.CTSignedTwipsMeasure': { val: INT16 },
  'org_docx4j_wml.CTSignedHpsMeasure': { val: INT16 },
  'org_docx4j_wml.HpsMeasure': { val: UINT16 },
  'org_docx4j_wml.CTTextScale': { val: UINT16 },
  'org_docx4j_wml.CTTabStop': { pos: { ...INT16, min: -31680, max: 31680 } },
  'org_docx4j_wml.CTBorder': { sz: { measured: true, bits: 8, signed: false }, space: { measured: true, bits: 5, signed: false } },
  'org_docx4j_wml.TblWidth': { w: INT16 },
  'org_docx4j_wml.CTHeight': { val: { ...INT16, signRule: 'hRule' } },
  'org_docx4j_wml.SectPr.PgSz': { w: { ...UINT16, min: 144, max: 31681 }, h: { ...UINT16, min: 144 } },
  'org_docx4j_wml.SectPr.PgMar': {
    top: INT16, bottom: INT16, left: UINT16_MINUS_ONE, right: UINT16_MINUS_ONE,
    header: UINT16_MINUS_ONE, footer: UINT16_MINUS_ONE, gutter: UINT16_MINUS_ONE,
  },
  'org_docx4j_wml.CTPageNumber': { start: { measured: true, bits: 32, signed: true } },
  'org_docx4j_wml.CTFramePr': {
    x: INT16, y: INT16, w: INT16, hSpace: INT16, vSpace: INT16,
    h: { ...UINT16, heightFlag: true, signRule: 'hRule' },
  },
  'org_docx4j_wml.CTTblPPr': { tblpX: INT16, tblpY: INT16, leftFromText: INT16, rightFromText: INT16, topFromText: INT16, bottomFromText: INT16 },
  'org_docx4j_wml.TblGridCol': { w: UNMEASURED },
  'org_docx4j_wml.CTColumns': { space: INT16 },
  'org_docx4j_wml.CTColumn': { w: { ...INT16, negative: 'computed' }, space: { ...INT16, negative: -1 } },
  'org_docx4j_wml.CTFitText': { val: { measured: true, bits: 32, signed: true } },
  'org_docx4j_wml.CTLineNumber': { distance: INT16 },
  'org_docx4j_wml.CTObject': { dxaOrig: UNMEASURED, dyaOrig: UNMEASURED },
  'org_docx4j_wml.CTTwipsMeasure': { val: UINT16 },
};

/** What Word 365 reads for an integer attribute's value (CR-016). */
export interface WordIntegerReading {
  /** The value Word reads, wrapped and clamped as it stores it; null for an unmeasured attribute. */
  readonly value: number | null;
  /** The value Word's own save writes: `value`, except `w:spacing/@w:before`'s 65535, written -1. */
  readonly written: number | null;
  /** For `w:spacing/@w:line`, `w:trHeight` and `w:framePr/@w:h`, the rule Word reads with it (`auto`, `exact`, `atLeast`). */
  readonly rule?: 'auto' | 'exact' | 'atLeast';
  /** For `w:ind`'s `firstLine` and `hanging`, which share one field, the one Word reads it as. */
  readonly attribute?: string;
  /**
   * Whether Word reads the value, and its rule, as written. For an unmeasured attribute, false only outside
   * -32768 to 65535, where every measured attribute of its type changes.
   */
  readonly kept: boolean;
  /** How Word stores it, for a message: `int16`, `uint16, then 144..31681`, `int16, the sign the rule`, ... */
  readonly storage: string;
}

function wrapTo(value: number, bits: number, signed: boolean): number {
  const modulus = 2 ** bits;
  const unsigned = ((value % modulus) + modulus) % modulus;
  const wrapped = signed && unsigned >= modulus / 2 ? unsigned - modulus : unsigned;
  return wrapped === 0 ? 0 : wrapped;
}

function storageOf(s: WordIntegerStorage): string {
  if (!s.measured) return 'unmeasured: changed outside -32768..65535';
  const width = s.bits === 16 || s.bits === 32 ? `${s.signed ? 'int' : 'uint'}${s.bits}` : `${s.bits} bits`;
  if (s.pair) return `${width}, shared by ${s.pair.positive} (+) and ${s.pair.negative} (-)`;
  if (s.heightFlag) return '15 bits, bit 15 clearing exact';
  if (s.negative !== undefined) return `${width}, a negative ${s.negative === 'computed' ? 'computed' : 'written -1'}`;
  if (s.signRule) return `${width}, the sign the rule`;
  if (s.min !== undefined && s.max !== undefined) return s.min === -s.max ? `${width}, then ±${s.max}` : `${width}, then ${s.min}..${s.max}`;
  return s.min !== undefined ? `${width}, then at least ${s.min}` : width;
}

/**
 * What Word 365 reads for `value` in an integer attribute (CR-016, measured in docx4j CR-066's rounds 5 and 6): the
 * value modulo its storage width, then clamped where Word clamps. For `w:spacing/@w:line`, `w:trHeight` and
 * `w:framePr/@w:h` pass the rule beside it (`lineRule`, `hRule`).
 * - Under `exact` and `atLeast`, a line or a row height keeps its rule in the sign of what is stored (Word stores an
 *   `exact` one negated): `exact 65535` reads as `atLeast 1`, a `w:trHeight` of 32768 as `exact 32768`. An absent
 *   `hRule` is `atLeast`, and `lineRule` `auto`; under `auto` a line is plain int16 and a row height is ignored.
 * - `w:ind`'s `firstLine` and `hanging` are one field: `firstLine` 32768 reads as `hanging` 32768.
 * - A frame's height keeps 15 bits, and bit 15 clears `exact` (measured under `exact`).
 * Undefined for an attribute not in `WORD_INTEGER_STORAGE`.
 */
export function wordIntegerReading(typeName: string, property: string, value: number, rule?: string): WordIntegerReading | undefined {
  const s = WORD_INTEGER_STORAGE[typeName]?.[property];
  if (s === undefined) return undefined;
  const storage = storageOf(s);
  if (!s.measured) return { value: null, written: null, kept: value >= -32768 && value <= 65535, storage };
  if (s.pair) {
    const stored = wrapTo(property === s.pair.negative ? -value : value, 16, true);
    const read = Math.abs(stored);
    const attribute = stored < 0 ? s.pair.negative : s.pair.positive;
    return { value: read, written: read, attribute, kept: read === value && (attribute === property || value === 0), storage };
  }
  if (s.heightFlag) {
    const stored = wrapTo(value, 16, false);
    const read = stored & 0x7FFF;
    const written = (rule ?? 'auto') as 'auto' | 'exact' | 'atLeast';
    const readRule = stored & 0x8000 ? 'auto' : written;
    return { value: read, written: read, rule: readRule, kept: read === value && readRule === written, storage };
  }
  if (s.signRule === 'hRule' && rule === 'auto') return { value, written: value, rule: 'auto', kept: true, storage: `${storage}; ignored under auto` };
  if (s.negative !== undefined) {
    const stored = wrapTo(value, s.bits, s.signed);
    if (stored >= 0) return { value: stored, written: stored, kept: stored === value, storage };
    const read = s.negative === 'computed' ? null : s.negative;
    return { value: read, written: read, kept: read === value, storage };
  }
  if (s.signRule) {
    const written = (s.signRule === 'lineRule' ? rule ?? 'auto' : rule ?? 'atLeast') as 'auto' | 'exact' | 'atLeast';
    if (s.signRule === 'lineRule' && written === 'auto') {
      const read = wrapTo(value, 16, true);
      return { value: read, written: read, rule: 'auto', kept: read === value, storage: 'int16' };
    }
    const stored = wrapTo(written === 'exact' ? -value : value, 16, true);
    const read = Math.abs(stored);
    const readRule = stored < 0 ? 'exact' : 'atLeast';
    const kept = read === value && (readRule === written || (written === 'auto' && readRule === 'atLeast'));
    return { value: read, written: read, rule: readRule, kept, storage };
  }
  let read = wrapTo(value, s.bits, s.signed);
  if (s.min !== undefined) read = Math.max(s.min, read);
  if (s.max !== undefined) read = Math.min(s.max, read);
  const max = 2 ** s.bits - 1;
  const written = s.maxWrittenAs !== undefined && read === max ? s.maxWrittenAs : read;
  return { value: read, written, kept: read === value || (s.maxWrittenAs !== undefined && value === s.maxWrittenAs && read === max), storage };
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
