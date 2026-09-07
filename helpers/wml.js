/** w:highlight names and their sRGB values, as in docx4j (darkYellow uses gold). */
export const HIGHLIGHT_COLORS = [
    ['black', '000000'], ['blue', '0000FF'], ['cyan', '00FFFF'], ['green', '008000'],
    ['magenta', 'FF00FF'], ['red', 'FF0000'], ['yellow', 'FFFF00'], ['white', 'FFFFFF'],
    ['darkBlue', '00008B'], ['darkCyan', '008B8B'], ['darkGreen', '006400'], ['darkMagenta', '8B008B'],
    ['darkRed', '8B0000'], ['darkYellow', 'FFD700'], ['darkGray', 'A9A9A9'], ['lightGray', 'D3D3D3'],
];
/** The `#RRGGBB` value of a highlight, or undefined for an unknown name (docx4j `Highlight.getHexVal()`). */
export function highlightHexValue(highlight) {
    const val = typeof highlight === 'string' ? highlight : highlight === null || highlight === void 0 ? void 0 : highlight.val;
    const entry = HIGHLIGHT_COLORS.find(([name]) => name === val);
    return entry ? '#' + entry[1] : undefined;
}
/**
 * The highlight name for a colour: a known name is returned as is, `#RRGGBB` is looked up in the
 * table; anything else (including rgb()) is undefined, as docx4j refuses it (`Highlight.setVal`).
 */
export function highlightNameForColor(color) {
    if (color === undefined)
        return undefined;
    if (HIGHLIGHT_COLORS.some(([name]) => name === color))
        return color;
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
export function isQFormat(exception) {
    var _a, _b;
    if (exception.qFormat !== undefined)
        return exception.qFormat;
    return (_b = (_a = exception.PARENT) === null || _a === void 0 ? void 0 : _a.defQFormat) !== null && _b !== void 0 ? _b : false;
}
/** An absent w:customStyle means a built-in style, whatever the schema default says (docx4j issue 641). */
export function isCustomStyle(style) {
    return style.customStyle === true;
}
