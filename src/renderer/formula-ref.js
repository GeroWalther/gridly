import { colToLetter, letterToCol } from '../shared/addr.js';

// Relative references have to move when a formula is copied or filled:
// copying "=A1+$B$2" one row down must yield "=A2+$B$2". This walks the
// formula rather than running a bare regex over it, so that string literals
// and sheet names are left alone — a naive pattern would happily rewrite the
// "Sheet1" in "Sheet1!A1" as if it were a cell reference.

const REF_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\w.(])/;
const MAX_ROW = 1048576;
const MAX_COL = 16384;

export function shiftFormula(formula, dr, dc) {
  if (!formula || (dr === 0 && dc === 0)) return formula;

  const s = String(formula);
  let out = '';
  let i = 0;

  while (i < s.length) {
    const ch = s[i];

    // Double-quoted string literal: copy verbatim, "" is an escaped quote.
    if (ch === '"') {
      let j = i + 1;
      while (j < s.length) {
        if (s[j] === '"') {
          if (s[j + 1] === '"') j += 2;
          else {
            j++;
            break;
          }
        } else j++;
      }
      out += s.slice(i, j);
      i = j;
      continue;
    }

    // Single-quoted sheet name, e.g. 'Q1 Results'!A1.
    if (ch === "'") {
      let j = i + 1;
      while (j < s.length) {
        if (s[j] === "'") {
          if (s[j + 1] === "'") j += 2;
          else {
            j++;
            break;
          }
        } else j++;
      }
      out += s.slice(i, j);
      i = j;
      continue;
    }

    if (/[A-Za-z_$]/.test(ch)) {
      const m = REF_RE.exec(s.slice(i));
      if (m) {
        out += shiftRef(m, dr, dc);
        i += m[0].length;
        continue;
      }
      // Not a reference — consume the whole identifier so that a bare sheet
      // name or function name is never partially rewritten.
      const idm = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(s.slice(i));
      if (idm) {
        out += idm[0];
        i += idm[0].length;
        continue;
      }
    }

    out += ch;
    i++;
  }

  return out;
}

function shiftRef(m, dr, dc) {
  const colAbs = m[1] === '$';
  const rowAbs = m[3] === '$';
  let col = letterToCol(m[2]);
  let row = parseInt(m[4], 10);

  if (!colAbs) col += dc;
  if (!rowAbs) row += dr;

  if (row < 1 || col < 1 || row > MAX_ROW || col > MAX_COL) return '#REF!';
  return (colAbs ? '$' : '') + colToLetter(col) + (rowAbs ? '$' : '') + row;
}
