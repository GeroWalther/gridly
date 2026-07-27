'use strict';

// The workbook model. Deliberately plain JSON: it crosses the IPC boundary
// between the main process (which owns file parsing) and the renderer, so it
// must survive structured cloning. No class instances, no Dates.

const DEFAULT_COL_WIDTH = 8.43; // Excel character units
const DEFAULT_ROW_HEIGHT = 15; // points

const DEFAULT_FONT = {
  name: 'Calibri',
  size: 11,
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  color: '#000000',
};

const DEFAULT_STYLE = {
  font: DEFAULT_FONT,
  fill: null,
  align: { h: null, v: 'bottom', wrap: false, indent: 0 },
  border: null,
  numFmt: 'General',
};

// Styles are interned in a per-workbook table; cells hold an integer index.
// Formatting a thousand cells bold then costs a thousand ints, not a thousand
// style objects, and equality checks in the toolbar become `a.s === b.s`.
class StyleTable {
  constructor(styles) {
    this.styles = styles && styles.length ? styles.slice() : [clone(DEFAULT_STYLE)];
    this.index = new Map();
    this.styles.forEach((s, i) => this.index.set(stableKey(s), i));
  }

  get(i) {
    return this.styles[i] || this.styles[0];
  }

  intern(style) {
    const norm = normalize(style);
    const k = stableKey(norm);
    const hit = this.index.get(k);
    if (hit !== undefined) return hit;
    const i = this.styles.length;
    this.styles.push(norm);
    this.index.set(k, i);
    return i;
  }

  // Returns the index of `base` with `patch` merged over it.
  derive(baseIndex, patch) {
    const base = this.get(baseIndex);
    const next = {
      font: Object.assign({}, base.font, patch.font),
      fill: patch.fill !== undefined ? patch.fill : base.fill,
      align: Object.assign({}, base.align, patch.align),
      border: patch.border !== undefined ? patch.border : base.border,
      numFmt: patch.numFmt !== undefined ? patch.numFmt : base.numFmt,
    };
    return this.intern(next);
  }

  toJSON() {
    return this.styles;
  }
}

function normalize(s) {
  const src = s || {};
  return {
    font: Object.assign({}, DEFAULT_FONT, src.font),
    fill: src.fill || null,
    align: Object.assign({}, DEFAULT_STYLE.align, src.align),
    border: src.border || null,
    numFmt: src.numFmt || 'General',
  };
}

// JSON.stringify with sorted keys, so two structurally equal styles always
// produce the same intern key regardless of property insertion order.
function stableKey(v) {
  if (v === null || typeof v !== 'object') return JSON.stringify(v);
  if (Array.isArray(v)) return '[' + v.map(stableKey).join(',') + ']';
  const keys = Object.keys(v).sort();
  return '{' + keys.map((k) => JSON.stringify(k) + ':' + stableKey(v[k])).join(',') + '}';
}

function clone(v) {
  return JSON.parse(JSON.stringify(v));
}

function emptySheet(name) {
  return {
    name,
    cells: {},
    cols: {},
    rows: {},
    merges: [],
    frozen: { row: 0, col: 0 },
    maxRow: 0,
    maxCol: 0,
    tabColor: null,
    defaultColWidth: DEFAULT_COL_WIDTH,
    defaultRowHeight: DEFAULT_ROW_HEIGHT,
  };
}

function emptyWorkbook() {
  return {
    filePath: null,
    fileType: 'xlsx',
    readOnly: false,
    sheets: [emptySheet('Sheet1')],
    styles: [clone(DEFAULT_STYLE)],
    definedNames: [],
  };
}

module.exports = {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  DEFAULT_FONT,
  DEFAULT_STYLE,
  StyleTable,
  emptySheet,
  emptyWorkbook,
  clone,
  stableKey,
};
