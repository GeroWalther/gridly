'use strict';

// Excel stores colours three different ways: literal ARGB, an index into a
// legacy 56-colour palette, or a theme slot plus a tint. Only the first is
// self-describing, so the other two are resolved here against the default
// Office palette. A file carrying a custom theme will render with default
// theme colours; everything else round-trips exactly.

const INDEXED = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
  '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
  '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
  '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
  '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
];

// Office 2013+ default theme, in the slot order Excel writes into styles.xml
// (note lt1/dk1 are swapped relative to the theme XML itself).
const THEME = [
  '#FFFFFF', '#000000', '#E7E6E6', '#44546A',
  '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47',
  '#0563C1', '#954F72',
];

function clamp255(n) {
  return Math.max(0, Math.min(255, Math.round(n)));
}

function hexToRgb(hex) {
  return {
    r: parseInt(hex.slice(1, 3), 16),
    g: parseInt(hex.slice(3, 5), 16),
    b: parseInt(hex.slice(5, 7), 16),
  };
}

function rgbToHex(c) {
  const h = (n) => clamp255(n).toString(16).padStart(2, '0').toUpperCase();
  return '#' + h(c.r) + h(c.g) + h(c.b);
}

// Excel's tint: positive lightens toward white, negative darkens toward black.
function applyTint(hex, tint) {
  if (!tint) return hex;
  const c = hexToRgb(hex);
  const f = (v) => (tint > 0 ? v + (255 - v) * tint : v * (1 + tint));
  return rgbToHex({ r: f(c.r), g: f(c.g), b: f(c.b) });
}

// ExcelJS colour object -> "#RRGGBB", or null when there is no colour.
// Alpha is dropped: the grid composites opaque fills, and Excel itself only
// meaningfully uses alpha in a handful of legacy cases.
function toHex(color) {
  if (!color) return null;
  if (typeof color === 'string') {
    return /^#?[0-9A-Fa-f]{6}$/.test(color) ? '#' + color.replace('#', '').toUpperCase() : null;
  }
  if (color.argb) {
    const s = String(color.argb).toUpperCase();
    const rgb = s.length === 8 ? s.slice(2) : s.length === 6 ? s : null;
    return rgb ? '#' + rgb : null;
  }
  if (typeof color.theme === 'number') {
    const base = THEME[color.theme] || '#000000';
    return applyTint(base, color.tint || 0);
  }
  if (typeof color.indexed === 'number') {
    const base = INDEXED[color.indexed];
    if (!base) return null;
    return applyTint(base, color.tint || 0);
  }
  return null;
}

// "#RRGGBB" -> the ARGB string ExcelJS wants on the way back out.
function toArgb(hex) {
  if (!hex) return null;
  const s = hex.replace('#', '').toUpperCase();
  return s.length === 6 ? 'FF' + s : s.length === 8 ? s : null;
}

module.exports = { toHex, toArgb, applyTint, INDEXED, THEME };
