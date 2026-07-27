import { StyleTable, emptySheet, emptyWorkbook } from '../shared/model.js';
import { key, parseRange, rangeToA1, rcToA1 } from '../shared/addr.js';
import { Engine } from './engine.js';
import { isDateFormat, dateToSerial, serialToDate, formatValue } from './format.js';

// Every mutation is expressed as an operation that knows how to produce its
// own inverse. Undo is then just "run the inverses in reverse order", which
// keeps the model and the formula engine in lockstep without either of them
// needing a private history.

export class Doc {
  constructor() {
    this.model = emptyWorkbook();
    this.styles = new StyleTable(this.model.styles);
    this.engine = new Engine();
    this.engine.load(this.model);
    this.activeSheet = 0;
    this.dirty = false;
    this.undoStack = [];
    this.redoStack = [];
    this.listeners = new Set();
  }

  // ------------------------------------------------------------- basics

  onChange(fn) {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  emit(detail) {
    for (const fn of this.listeners) fn(detail || {});
  }

  load(model) {
    this.model = model;
    this.styles = new StyleTable(model.styles);
    this.engine.load(model);
    this.activeSheet = 0;
    this.dirty = false;
    this.undoStack = [];
    this.redoStack = [];
    this.emit({ type: 'load' });
  }

  loadEmpty() {
    this.load(emptyWorkbook());
  }

  get sheet() {
    return this.model.sheets[this.activeSheet];
  }

  sheetAt(i) {
    return this.model.sheets[i];
  }

  cell(row, col, sheetIndex) {
    const sheet = this.model.sheets[sheetIndex === undefined ? this.activeSheet : sheetIndex];
    return sheet ? sheet.cells[key(row, col)] : undefined;
  }

  styleOf(row, col, sheetIndex) {
    const c = this.cell(row, col, sheetIndex);
    return this.styles.get(c ? c.s || 0 : 0);
  }

  // What the editor should show when you open a cell: the formula if there is
  // one, otherwise a round-trippable rendering of the stored value.
  editText(row, col) {
    const c = this.cell(row, col);
    if (!c) return '';
    if (c.f) return '=' + c.f;
    if (c.v === undefined || c.v === null) return '';
    const style = this.styles.get(c.s || 0);
    if (typeof c.v === 'number' && isDateFormat(style.numFmt)) {
      return formatValue(c.v, style.numFmt).text;
    }
    if (typeof c.v === 'boolean') return c.v ? 'TRUE' : 'FALSE';
    return String(c.v);
  }

  displayText(row, col) {
    const c = this.cell(row, col);
    if (!c || c.v === undefined || c.v === null) return '';
    return formatValue(c.v, this.styles.get(c.s || 0).numFmt).text;
  }

  // ------------------------------------------------------- transactions

  // Runs `fn` with a recorder; everything it does becomes one undo step.
  transact(label, fn) {
    const inverses = [];
    const tx = {
      apply: (op) => {
        const inv = this.applyOp(op);
        if (inv) inverses.push(inv);
        return inv;
      },
    };

    const result = fn(tx);
    if (!inverses.length) return result;

    this.undoStack.push({ label, inverses });
    if (this.undoStack.length > 200) this.undoStack.shift();
    this.redoStack.length = 0;
    this.markDirty();
    this.emit({ type: 'change' });
    return result;
  }

  undo() {
    const step = this.undoStack.pop();
    if (!step) return false;
    const redoInverses = [];
    for (let i = step.inverses.length - 1; i >= 0; i--) {
      const inv = this.applyOp(step.inverses[i]);
      if (inv) redoInverses.push(inv);
    }
    this.redoStack.push({ label: step.label, inverses: redoInverses.reverse() });
    this.markDirty();
    this.emit({ type: 'change', undo: true });
    return true;
  }

  redo() {
    const step = this.redoStack.pop();
    if (!step) return false;
    const undoInverses = [];
    for (let i = step.inverses.length - 1; i >= 0; i--) {
      const inv = this.applyOp(step.inverses[i]);
      if (inv) undoInverses.push(inv);
    }
    this.undoStack.push({ label: step.label, inverses: undoInverses.reverse() });
    this.markDirty();
    this.emit({ type: 'change', redo: true });
    return true;
  }

  markDirty() {
    this.dirty = true;
  }

  markClean() {
    this.dirty = false;
    this.emit({ type: 'clean' });
  }

  // ---------------------------------------------------------- operations

  applyOp(op) {
    switch (op.type) {
      case 'composite': {
        const inv = [];
        for (const sub of op.ops) inv.push(this.applyOp(sub));
        return { type: 'composite', ops: inv.reverse().filter(Boolean) };
      }
      case 'cells':
        return this.opCells(op);
      case 'style':
        return this.opStyle(op);
      case 'insertRows':
        return this.opInsertRows(op);
      case 'deleteRows':
        return this.opDeleteRows(op);
      case 'insertCols':
        return this.opInsertCols(op);
      case 'deleteCols':
        return this.opDeleteCols(op);
      case 'colWidth':
        return this.opColWidth(op);
      case 'rowHeight':
        return this.opRowHeight(op);
      case 'merges':
        return this.opMerges(op);
      case 'frozen':
        return this.opFrozen(op);
      case 'sheetAdd':
        return this.opSheetAdd(op);
      case 'sheetRemove':
        return this.opSheetRemove(op);
      case 'sheetRename':
        return this.opSheetRename(op);
      case 'sheetColor':
        return this.opSheetColor(op);
      default:
        return null;
    }
  }

  // Replace a set of cells wholesale. `cell: null` clears.
  opCells(op) {
    const si = op.sheetIndex;
    const sheet = this.model.sheets[si];
    if (!sheet) return null;

    const before = [];
    const engineEdits = [];

    for (const item of op.cells) {
      const k = key(item.row, item.col);
      const prev = sheet.cells[k];
      before.push({ row: item.row, col: item.col, cell: prev ? Object.assign({}, prev) : null });

      if (item.cell === null || item.cell === undefined) {
        delete sheet.cells[k];
        engineEdits.push({ sheetIndex: si, row: item.row, col: item.col, content: null });
      } else {
        sheet.cells[k] = Object.assign({}, item.cell);
        bump(sheet, item.row, item.col);
        engineEdits.push({
          sheetIndex: si,
          row: item.row,
          col: item.col,
          content: item.cell.f ? '=' + item.cell.f : item.cell.v === undefined ? null : item.cell.v,
        });
      }
    }

    this.applyEngineChanges(this.engine.setCells(engineEdits));
    return { type: 'cells', sheetIndex: si, cells: before };
  }

  // Writes the engine's recomputed values into the model. Formula cells keep
  // their formula; only the cached value changes.
  applyEngineChanges(changes) {
    for (const ch of changes) {
      const sheet = this.model.sheets[ch.sheetIndex];
      if (!sheet) continue;
      const k = key(ch.row, ch.col);
      const cell = sheet.cells[k];
      if (!cell) continue;
      if (!cell.f) continue; // literal cells are authoritative in the model
      if (ch.value === null) {
        delete cell.v;
        delete cell.t;
      } else {
        cell.v = ch.value;
        cell.t = ch.type;
      }
    }
  }

  opStyle(op) {
    const sheet = this.model.sheets[op.sheetIndex];
    if (!sheet) return null;
    const before = [];
    for (const item of op.cells) {
      const k = key(item.row, item.col);
      let cell = sheet.cells[k];
      before.push({ row: item.row, col: item.col, s: cell ? cell.s || 0 : null });
      if (!cell) {
        if (item.s === 0 || item.s === null) continue;
        cell = { s: item.s };
        sheet.cells[k] = cell;
        bump(sheet, item.row, item.col);
      } else if (item.s === null) {
        cell.s = 0;
      } else {
        cell.s = item.s;
      }
    }
    return { type: 'style', sheetIndex: op.sheetIndex, cells: before };
  }

  opInsertRows(op) {
    const sheet = this.model.sheets[op.sheetIndex];
    if (!sheet) return null;
    shiftCells(sheet, op.at, op.count, 'row');
    shiftSpecs(sheet, 'rows', op.at, op.count);
    sheet.merges = shiftMerges(sheet.merges, op.at, op.count, 'row');
    sheet.maxRow = Math.min(1048576, sheet.maxRow + op.count);
    this.engine.addRows(op.sheetIndex, op.at, op.count);
    this.applyEngineChanges(this.engine.refreshAll(this.model));
    return { type: 'deleteRows', sheetIndex: op.sheetIndex, at: op.at, count: op.count };
  }

  opDeleteRows(op) {
    const sheet = this.model.sheets[op.sheetIndex];
    if (!sheet) return null;

    // Everything about to disappear, captured so the inverse can put it back.
    const saved = [];
    for (const k of Object.keys(sheet.cells)) {
      const comma = k.indexOf(',');
      const r = +k.slice(0, comma);
      if (r >= op.at && r < op.at + op.count) {
        saved.push({ row: r, col: +k.slice(comma + 1), cell: Object.assign({}, sheet.cells[k]) });
      }
    }
    const savedRows = {};
    for (const rStr of Object.keys(sheet.rows)) {
      const r = +rStr;
      if (r >= op.at && r < op.at + op.count) savedRows[r] = Object.assign({}, sheet.rows[rStr]);
    }
    const savedMerges = sheet.merges.slice();

    for (const item of saved) delete sheet.cells[key(item.row, item.col)];
    for (const rStr of Object.keys(savedRows)) delete sheet.rows[rStr];

    shiftCells(sheet, op.at + op.count, -op.count, 'row');
    shiftSpecs(sheet, 'rows', op.at + op.count, -op.count);
    sheet.merges = shiftMerges(dropMerges(sheet.merges, op.at, op.count, 'row'), op.at + op.count, -op.count, 'row');
    sheet.maxRow = Math.max(0, sheet.maxRow - op.count);

    this.engine.removeRows(op.sheetIndex, op.at, op.count);
    this.applyEngineChanges(this.engine.refreshAll(this.model));

    return {
      type: 'composite',
      ops: [
        { type: 'insertRows', sheetIndex: op.sheetIndex, at: op.at, count: op.count },
        { type: 'cells', sheetIndex: op.sheetIndex, cells: saved },
        { type: 'rowHeight', sheetIndex: op.sheetIndex, rows: Object.keys(savedRows).map((r) => ({ row: +r, spec: savedRows[r] })) },
        { type: 'merges', sheetIndex: op.sheetIndex, merges: savedMerges },
      ],
    };
  }

  opInsertCols(op) {
    const sheet = this.model.sheets[op.sheetIndex];
    if (!sheet) return null;
    shiftCells(sheet, op.at, op.count, 'col');
    shiftSpecs(sheet, 'cols', op.at, op.count);
    sheet.merges = shiftMerges(sheet.merges, op.at, op.count, 'col');
    sheet.maxCol = Math.min(16384, sheet.maxCol + op.count);
    this.engine.addColumns(op.sheetIndex, op.at, op.count);
    this.applyEngineChanges(this.engine.refreshAll(this.model));
    return { type: 'deleteCols', sheetIndex: op.sheetIndex, at: op.at, count: op.count };
  }

  opDeleteCols(op) {
    const sheet = this.model.sheets[op.sheetIndex];
    if (!sheet) return null;

    const saved = [];
    for (const k of Object.keys(sheet.cells)) {
      const comma = k.indexOf(',');
      const c = +k.slice(comma + 1);
      if (c >= op.at && c < op.at + op.count) {
        saved.push({ row: +k.slice(0, comma), col: c, cell: Object.assign({}, sheet.cells[k]) });
      }
    }
    const savedCols = {};
    for (const cStr of Object.keys(sheet.cols)) {
      const c = +cStr;
      if (c >= op.at && c < op.at + op.count) savedCols[c] = Object.assign({}, sheet.cols[cStr]);
    }
    const savedMerges = sheet.merges.slice();

    for (const item of saved) delete sheet.cells[key(item.row, item.col)];
    for (const cStr of Object.keys(savedCols)) delete sheet.cols[cStr];

    shiftCells(sheet, op.at + op.count, -op.count, 'col');
    shiftSpecs(sheet, 'cols', op.at + op.count, -op.count);
    sheet.merges = shiftMerges(dropMerges(sheet.merges, op.at, op.count, 'col'), op.at + op.count, -op.count, 'col');
    sheet.maxCol = Math.max(0, sheet.maxCol - op.count);

    this.engine.removeColumns(op.sheetIndex, op.at, op.count);
    this.applyEngineChanges(this.engine.refreshAll(this.model));

    return {
      type: 'composite',
      ops: [
        { type: 'insertCols', sheetIndex: op.sheetIndex, at: op.at, count: op.count },
        { type: 'cells', sheetIndex: op.sheetIndex, cells: saved },
        { type: 'colWidth', sheetIndex: op.sheetIndex, cols: Object.keys(savedCols).map((c) => ({ col: +c, spec: savedCols[c] })) },
        { type: 'merges', sheetIndex: op.sheetIndex, merges: savedMerges },
      ],
    };
  }

  opColWidth(op) {
    const sheet = this.model.sheets[op.sheetIndex];
    if (!sheet) return null;
    const before = [];
    for (const item of op.cols) {
      const cur = sheet.cols[item.col];
      before.push({ col: item.col, spec: cur ? Object.assign({}, cur) : null });
      if (item.spec === null) delete sheet.cols[item.col];
      else sheet.cols[item.col] = Object.assign({}, item.spec);
    }
    return { type: 'colWidth', sheetIndex: op.sheetIndex, cols: before };
  }

  opRowHeight(op) {
    const sheet = this.model.sheets[op.sheetIndex];
    if (!sheet) return null;
    const before = [];
    for (const item of op.rows) {
      const cur = sheet.rows[item.row];
      before.push({ row: item.row, spec: cur ? Object.assign({}, cur) : null });
      if (item.spec === null) delete sheet.rows[item.row];
      else sheet.rows[item.row] = Object.assign({}, item.spec);
    }
    return { type: 'rowHeight', sheetIndex: op.sheetIndex, rows: before };
  }

  opMerges(op) {
    const sheet = this.model.sheets[op.sheetIndex];
    if (!sheet) return null;
    const before = sheet.merges.slice();
    sheet.merges = op.merges.slice();
    return { type: 'merges', sheetIndex: op.sheetIndex, merges: before };
  }

  opFrozen(op) {
    const sheet = this.model.sheets[op.sheetIndex];
    if (!sheet) return null;
    const before = Object.assign({}, sheet.frozen);
    sheet.frozen = Object.assign({}, op.frozen);
    return { type: 'frozen', sheetIndex: op.sheetIndex, frozen: before };
  }

  opSheetAdd(op) {
    const sheet = op.sheet || emptySheet(op.name);
    this.model.sheets.splice(op.index, 0, sheet);
    const actual = this.engine.addSheet(sheet.name, op.index);
    if (actual !== sheet.name) sheet.name = actual;
    if (op.sheet) {
      // Restoring a deleted sheet: push its cells back into the engine.
      const edits = [];
      for (const k of Object.keys(sheet.cells)) {
        const comma = k.indexOf(',');
        const cell = sheet.cells[k];
        edits.push({
          sheetIndex: op.index,
          row: +k.slice(0, comma),
          col: +k.slice(comma + 1),
          content: cell.f ? '=' + cell.f : cell.v === undefined ? null : cell.v,
        });
      }
      if (edits.length) this.applyEngineChanges(this.engine.setCells(edits));
    }
    return { type: 'sheetRemove', index: op.index };
  }

  opSheetRemove(op) {
    const sheet = this.model.sheets[op.index];
    if (!sheet || this.model.sheets.length <= 1) return null;
    this.model.sheets.splice(op.index, 1);
    this.engine.removeSheet(op.index);
    if (this.activeSheet >= this.model.sheets.length) this.activeSheet = this.model.sheets.length - 1;
    return { type: 'sheetAdd', index: op.index, sheet };
  }

  opSheetRename(op) {
    const sheet = this.model.sheets[op.index];
    if (!sheet) return null;
    const before = sheet.name;
    const actual = this.engine.renameSheet(op.index, op.name);
    if (actual === null) return null;
    sheet.name = actual;
    return { type: 'sheetRename', index: op.index, name: before };
  }

  opSheetColor(op) {
    const sheet = this.model.sheets[op.index];
    if (!sheet) return null;
    const before = sheet.tabColor;
    sheet.tabColor = op.color;
    return { type: 'sheetColor', index: op.index, color: before };
  }

  // ---------------------------------------------- high-level edit helpers

  setCellInput(row, col, text, sheetIndex) {
    const si = sheetIndex === undefined ? this.activeSheet : sheetIndex;
    const existing = this.cell(row, col, si);
    const baseStyle = existing ? existing.s || 0 : 0;
    const parsed = parseInput(text, this.styles.get(baseStyle));

    let s = baseStyle;
    if (parsed.numFmt) s = this.styles.derive(baseStyle, { numFmt: parsed.numFmt });

    const cell = parsed.empty ? (s === 0 ? null : { s }) : Object.assign({ s }, parsed.cell);

    this.transact('Edit Cell', (tx) => {
      tx.apply({ type: 'cells', sheetIndex: si, cells: [{ row, col, cell }] });
    });
  }

  // Applies a style patch to every cell of a range, creating cells as needed.
  applyStyleToRange(range, patch, label) {
    const si = this.activeSheet;
    const sheet = this.sheet;
    const items = [];
    for (let r = range.r1; r <= range.r2; r++) {
      for (let c = range.c1; c <= range.c2; c++) {
        const cur = sheet.cells[key(r, c)];
        const base = cur ? cur.s || 0 : 0;
        const next = this.styles.derive(base, typeof patch === 'function' ? patch(this.styles.get(base)) : patch);
        if (next !== base) items.push({ row: r, col: c, s: next });
      }
    }
    if (!items.length) return;
    this.transact(label || 'Format', (tx) => {
      tx.apply({ type: 'style', sheetIndex: si, cells: items });
    });
  }

  clearRange(range, what) {
    const si = this.activeSheet;
    const sheet = this.sheet;
    const cells = [];
    const styleItems = [];
    for (let r = range.r1; r <= range.r2; r++) {
      for (let c = range.c1; c <= range.c2; c++) {
        const cur = sheet.cells[key(r, c)];
        if (!cur) continue;
        if (what === 'formats') {
          if ((cur.s || 0) !== 0) styleItems.push({ row: r, col: c, s: 0 });
        } else if (what === 'all') {
          cells.push({ row: r, col: c, cell: null });
        } else {
          // Contents only: keep the cell's formatting.
          if (cur.v === undefined && !cur.f) continue;
          cells.push({ row: r, col: c, cell: (cur.s || 0) === 0 ? null : { s: cur.s } });
        }
      }
    }
    if (!cells.length && !styleItems.length) return;
    this.transact('Clear', (tx) => {
      if (cells.length) tx.apply({ type: 'cells', sheetIndex: si, cells });
      if (styleItems.length) tx.apply({ type: 'style', sheetIndex: si, cells: styleItems });
    });
  }

  // Writes a rectangular block of text inputs starting at (row, col), parsing
  // each one the way typing it would. Cells keep whatever formatting they had.
  writeBlock(row, col, rows) {
    const si = this.activeSheet;
    const cells = [];
    rows.forEach((cols, dr) => {
      cols.forEach((raw, dc) => {
        const r = row + dr;
        const c = col + dc;
        if (r > 1048576 || c > 16384) return;
        const existing = this.cell(r, c, si);
        const baseStyle = existing ? existing.s || 0 : 0;

        const parsed = parseInput(raw == null ? '' : String(raw), this.styles.get(baseStyle));
        let s = baseStyle;
        if (parsed.numFmt) s = this.styles.derive(baseStyle, { numFmt: parsed.numFmt });
        cells.push({ row: r, col: c, cell: parsed.empty ? (s === 0 ? null : { s }) : Object.assign({ s }, parsed.cell) });
      });
    });
    if (!cells.length) return;
    this.transact('Paste', (tx) => {
      tx.apply({ type: 'cells', sheetIndex: si, cells });
    });
  }
}

// ---------------------------------------------------------------- shifting

function bump(sheet, row, col) {
  if (row > sheet.maxRow) sheet.maxRow = row;
  if (col > sheet.maxCol) sheet.maxCol = col;
}

// Moves every cell at or beyond `at` by `delta` along one axis.
function shiftCells(sheet, at, delta, axis) {
  const entries = [];
  for (const k of Object.keys(sheet.cells)) {
    const comma = k.indexOf(',');
    const r = +k.slice(0, comma);
    const c = +k.slice(comma + 1);
    const v = axis === 'row' ? r : c;
    if (v >= at) entries.push({ r, c, cell: sheet.cells[k] });
  }
  for (const e of entries) delete sheet.cells[key(e.r, e.c)];
  // Sort so that a positive shift writes the far end first and never
  // overwrites a cell it still has to move.
  entries.sort((a, b) => (delta > 0 ? (axis === 'row' ? b.r - a.r : b.c - a.c) : axis === 'row' ? a.r - b.r : a.c - b.c));
  for (const e of entries) {
    const nr = axis === 'row' ? e.r + delta : e.r;
    const nc = axis === 'col' ? e.c + delta : e.c;
    if (nr < 1 || nc < 1) continue;
    sheet.cells[key(nr, nc)] = e.cell;
  }
}

function shiftSpecs(sheet, prop, at, delta) {
  const src = sheet[prop];
  const out = {};
  for (const kStr of Object.keys(src)) {
    const n = +kStr;
    if (n < at) out[n] = src[kStr];
    else if (n + delta >= 1) out[n + delta] = src[kStr];
  }
  sheet[prop] = out;
}

function shiftMerges(merges, at, delta, axis) {
  return merges
    .map((m) => {
      const r = parseRange(m);
      if (!r) return null;
      if (axis === 'row') {
        if (r.r1 >= at) r.r1 += delta;
        if (r.r2 >= at) r.r2 += delta;
      } else {
        if (r.c1 >= at) r.c1 += delta;
        if (r.c2 >= at) r.c2 += delta;
      }
      if (r.r1 < 1 || r.c1 < 1 || r.r2 < r.r1 || r.c2 < r.c1) return null;
      return rangeToA1(r);
    })
    .filter(Boolean);
}

// Drops merges that fall entirely inside a deleted band.
function dropMerges(merges, at, count, axis) {
  return merges.filter((m) => {
    const r = parseRange(m);
    if (!r) return false;
    const lo = axis === 'row' ? r.r1 : r.c1;
    const hi = axis === 'row' ? r.r2 : r.c2;
    return !(lo >= at && hi < at + count);
  });
}

// ------------------------------------------------------------ input parsing

const NUM_RE = /^-?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/;
const GROUPED_RE = /^-?\d{1,3}(,\d{3})+(\.\d+)?$/;
const PERCENT_RE = /^(-?[\d.,]+)\s*%$/;
const CURRENCY_RE = /^(-?)\s*([$€£¥])\s*([\d.,]+)$/;
const ISO_DATE_RE = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
const US_DATE_RE = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/;
const TIME_RE = /^(\d{1,2}):(\d{2})(:(\d{2}))?\s*(am|pm)?$/i;

// Mirrors how Excel interprets what you type: a leading "=" makes a formula,
// a leading apostrophe forces text, and otherwise the text is coerced to a
// number, percentage, currency amount, date or boolean where it can be —
// applying the matching number format when the cell had none.
export function parseInput(text, currentStyle) {
  const raw = text == null ? '' : String(text);

  if (raw === '') return { empty: true };

  if (raw.charAt(0) === '=') {
    const formula = raw.slice(1).trim();
    if (!formula) return { cell: { v: raw, t: 's' } };
    return { cell: { f: formula } };
  }

  if (raw.charAt(0) === "'") return { cell: { v: raw.slice(1), t: 's' } };

  const isTextFormatted = currentStyle && currentStyle.numFmt === '@';
  if (isTextFormatted) return { cell: { v: raw, t: 's' } };

  const trimmed = raw.trim();

  if (/^true$/i.test(trimmed)) return { cell: { v: true, t: 'b' } };
  if (/^false$/i.test(trimmed)) return { cell: { v: false, t: 'b' } };

  const hasFormat = currentStyle && currentStyle.numFmt && currentStyle.numFmt !== 'General';

  const pct = PERCENT_RE.exec(trimmed);
  if (pct) {
    const n = toNumber(pct[1]);
    if (n !== null) {
      return { cell: { v: n / 100, t: 'n' }, numFmt: hasFormat ? null : '0.00%' };
    }
  }

  const cur = CURRENCY_RE.exec(trimmed);
  if (cur) {
    const n = toNumber(cur[3]);
    if (n !== null) {
      const sign = cur[1] === '-' ? -1 : 1;
      const symbol = cur[2];
      return { cell: { v: sign * n, t: 'n' }, numFmt: hasFormat ? null : `"${symbol}"#,##0.00` };
    }
  }

  const n = toNumber(trimmed);
  if (n !== null) {
    // A grouped number like 1,234 implies a thousands format when the cell
    // had none, matching what Excel does.
    const numFmt = !hasFormat && GROUPED_RE.test(trimmed) ? '#,##0' + (trimmed.includes('.') ? '.00' : '') : null;
    return { cell: { v: n, t: 'n' }, numFmt };
  }

  const date = toDateSerial(trimmed);
  if (date !== null) {
    return { cell: { v: date.serial, t: 'n' }, numFmt: hasFormat ? null : date.numFmt };
  }

  return { cell: { v: raw, t: 's' } };
}

function toNumber(s) {
  const t = String(s).trim();
  if (NUM_RE.test(t)) return parseFloat(t);
  if (GROUPED_RE.test(t)) return parseFloat(t.replace(/,/g, ''));
  return null;
}

function toDateSerial(s) {
  let m = ISO_DATE_RE.exec(s);
  if (m) {
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    if (d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3]) {
      return { serial: dateToSerial(d), numFmt: 'yyyy-mm-dd' };
    }
    return null;
  }

  m = US_DATE_RE.exec(s);
  if (m) {
    let year = +m[3];
    if (year < 100) year += year < 30 ? 2000 : 1900;
    const d = new Date(Date.UTC(year, +m[1] - 1, +m[2]));
    if (d.getUTCMonth() === +m[1] - 1 && d.getUTCDate() === +m[2]) {
      return { serial: dateToSerial(d), numFmt: 'm/d/yyyy' };
    }
    return null;
  }

  m = TIME_RE.exec(s);
  if (m) {
    let h = +m[1];
    const min = +m[2];
    const sec = m[4] ? +m[4] : 0;
    const ampm = m[5] ? m[5].toLowerCase() : null;
    if (ampm === 'pm' && h < 12) h += 12;
    if (ampm === 'am' && h === 12) h = 0;
    if (h > 23 || min > 59 || sec > 59) return null;
    return {
      serial: (h * 3600 + min * 60 + sec) / 86400,
      numFmt: ampm ? 'h:mm AM/PM' : sec ? 'h:mm:ss' : 'h:mm',
    };
  }

  return null;
}
