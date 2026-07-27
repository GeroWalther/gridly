'use strict';

// Cell addressing helpers. Rows and columns are 1-based everywhere in the
// model, matching how Excel and ExcelJS number things. HyperFormula is
// 0-based, so conversions live in engine.js rather than here.

function colToLetter(col) {
  let s = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function letterToCol(letters) {
  let n = 0;
  const up = letters.toUpperCase();
  for (let i = 0; i < up.length; i++) {
    n = n * 26 + (up.charCodeAt(i) - 64);
  }
  return n;
}

const A1_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/;

// "B7" -> {row: 7, col: 2}; returns null when it isn't a cell reference.
function a1ToRC(a1) {
  const m = A1_RE.exec(String(a1).trim());
  if (!m) return null;
  return { row: parseInt(m[2], 10), col: letterToCol(m[1]) };
}

function rcToA1(row, col) {
  return colToLetter(col) + row;
}

// "A1:B3" (or a bare "A1") -> {r1, c1, r2, c2}, normalised so r1<=r2.
function parseRange(ref) {
  const parts = String(ref).replace(/^.*!/, '').split(':');
  const a = a1ToRC(parts[0]);
  if (!a) return null;
  const b = parts.length > 1 ? a1ToRC(parts[1]) : a;
  if (!b) return null;
  return {
    r1: Math.min(a.row, b.row),
    c1: Math.min(a.col, b.col),
    r2: Math.max(a.row, b.row),
    c2: Math.max(a.col, b.col),
  };
}

function rangeToA1(r) {
  if (r.r1 === r.r2 && r.c1 === r.c2) return rcToA1(r.r1, r.c1);
  return rcToA1(r.r1, r.c1) + ':' + rcToA1(r.r2, r.c2);
}

// Model cells are stored in a flat object keyed "row,col" so the whole
// workbook stays structured-clone friendly for IPC.
function key(row, col) {
  return row + ',' + col;
}

function unkey(k) {
  const i = k.indexOf(',');
  return { row: +k.slice(0, i), col: +k.slice(i + 1) };
}

module.exports = {
  colToLetter,
  letterToCol,
  a1ToRC,
  rcToA1,
  parseRange,
  rangeToA1,
  key,
  unkey,
};
