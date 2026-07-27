import { HyperFormula, DetailedCellError } from 'hyperformula';
import { key } from '../shared/addr.js';

// HyperFormula is the calculation engine only. The workbook model stays the
// single source of truth for what is displayed and saved: every change the
// engine reports is written back into the model, so rendering never has to
// call into the engine.
//
// Addressing note: the model is 1-based (like Excel and ExcelJS), the engine
// is 0-based. Every conversion happens in this file and nowhere else.

const CONFIG = {
  licenseKey: 'gpl-v3',
  // Match Excel's grid limits; HyperFormula's own defaults are far smaller
  // and would reject perfectly ordinary workbooks.
  maxRows: 1048576,
  maxColumns: 16384,
  // Excel's serial-date origin.
  nullDate: { year: 1899, month: 12, day: 30 },
  smartRounding: true,
  useArrayArithmetic: true,
  functionArgSeparator: ',',
};

export class Engine {
  constructor() {
    this.hf = null;
    this.sheetIds = []; // model sheet index -> engine sheet id
  }

  // Builds the engine from a freshly loaded workbook model.
  load(model) {
    if (this.hf) this.hf.destroy();

    const sheets = {};
    const names = [];
    model.sheets.forEach((sheet, i) => {
      const name = uniqueName(sheet.name, names);
      names.push(name);
      sheets[name] = sheetToArray(sheet);
    });

    this.hf = HyperFormula.buildFromSheets(sheets, CONFIG);
    this.sheetIds = names.map((n) => this.hf.getSheetId(n));

    for (const dn of model.definedNames || []) {
      try {
        this.hf.addNamedExpression(dn.name, '=' + dn.ref);
      } catch (_) {
        // Names the engine cannot represent (sheet-scoped, 3-D refs) are
        // skipped; formulas using them will report #NAME?.
      }
    }

    return this.recalcAll(model);
  }

  sheetId(sheetIndex) {
    return this.sheetIds[sheetIndex];
  }

  // Writes the engine's computed value for every formula cell back into the
  // model. Called after load, since a file's cached results may be stale or
  // missing entirely.
  recalcAll(model) {
    model.sheets.forEach((sheet, si) => {
      const id = this.sheetIds[si];
      if (id === undefined) return;
      for (const k of Object.keys(sheet.cells)) {
        const cell = sheet.cells[k];
        if (!cell.f) continue;
        const comma = k.indexOf(',');
        const row = +k.slice(0, comma) - 1;
        const col = +k.slice(comma + 1) - 1;
        applyValue(cell, this.hf.getCellValue({ sheet: id, row, col }));
      }
    });
    return model;
  }

  // Sets one cell and returns the list of model cells whose values changed,
  // as [{sheetIndex, row, col, value, type}]. The caller applies them.
  setCell(sheetIndex, row, col, content) {
    const id = this.sheetIds[sheetIndex];
    if (id === undefined) return [];
    let changes;
    try {
      changes = this.hf.setCellContents({ sheet: id, row: row - 1, col: col - 1 }, content);
    } catch (err) {
      // An invalid formula still needs to land in the cell so the user can
      // fix it; report it as a #ERROR! value.
      return [{ sheetIndex, row, col, value: '#ERROR!', type: 'e' }];
    }
    return this.mapChanges(changes);
  }

  // Sets many cells in a single evaluation pass — one recalculation for the
  // whole block instead of one per cell. `cells` is [{sheetIndex,row,col,content}].
  setCells(cells) {
    const changes = this.hf.batch(() => {
      for (const c of cells) {
        const id = this.sheetIds[c.sheetIndex];
        if (id === undefined) continue;
        try {
          this.hf.setCellContents({ sheet: id, row: c.row - 1, col: c.col - 1 }, c.content);
        } catch (_) {
          /* skip cells the engine rejects; the model keeps the raw text */
        }
      }
    });
    return this.mapChanges(changes);
  }

  mapChanges(changes) {
    const out = [];
    for (const ch of changes || []) {
      if (!ch.address) continue;
      const si = this.sheetIds.indexOf(ch.address.sheet);
      if (si === -1) continue;
      const { value, type } = describe(ch.newValue);
      out.push({ sheetIndex: si, row: ch.address.row + 1, col: ch.address.col + 1, value, type });
    }
    return out;
  }

  getValue(sheetIndex, row, col) {
    const id = this.sheetIds[sheetIndex];
    if (id === undefined) return null;
    return describe(this.hf.getCellValue({ sheet: id, row: row - 1, col: col - 1 }));
  }

  // Re-reads every formula cell of the workbook. Used after bulk operations
  // where collecting incremental changes is not worth the bookkeeping.
  refreshAll(model) {
    const changed = [];
    model.sheets.forEach((sheet, si) => {
      const id = this.sheetIds[si];
      if (id === undefined) return;
      for (const k of Object.keys(sheet.cells)) {
        const cell = sheet.cells[k];
        if (!cell.f) continue;
        const comma = k.indexOf(',');
        const row = +k.slice(0, comma);
        const col = +k.slice(comma + 1);
        const { value, type } = describe(this.hf.getCellValue({ sheet: id, row: row - 1, col: col - 1 }));
        if (cell.v !== value) changed.push({ sheetIndex: si, row, col, value, type });
      }
    });
    return changed;
  }

  // Structural edits go through the engine so that formulas referencing the
  // moved cells are rewritten the way Excel rewrites them.
  addRows(sheetIndex, atRow, count) {
    this.hf.addRows(this.sheetIds[sheetIndex], [atRow - 1, count]);
  }

  removeRows(sheetIndex, atRow, count) {
    this.hf.removeRows(this.sheetIds[sheetIndex], [atRow - 1, count]);
  }

  addColumns(sheetIndex, atCol, count) {
    this.hf.addColumns(this.sheetIds[sheetIndex], [atCol - 1, count]);
  }

  removeColumns(sheetIndex, atCol, count) {
    this.hf.removeColumns(this.sheetIds[sheetIndex], [atCol - 1, count]);
  }

  addSheet(name, sheetIndex) {
    const actual = this.hf.addSheet(name);
    const id = this.hf.getSheetId(actual);
    this.sheetIds.splice(sheetIndex, 0, id);
    return actual;
  }

  removeSheet(sheetIndex) {
    const id = this.sheetIds[sheetIndex];
    if (id === undefined) return;
    this.hf.removeSheet(id);
    this.sheetIds.splice(sheetIndex, 1);
  }

  renameSheet(sheetIndex, name) {
    const id = this.sheetIds[sheetIndex];
    if (id === undefined) return name;
    try {
      this.hf.renameSheet(id, name);
      return name;
    } catch (_) {
      return null;
    }
  }

  // Validates a formula without committing it.
  isValidFormula(text) {
    try {
      return this.hf.validateFormula(text);
    } catch (_) {
      return false;
    }
  }

  destroy() {
    if (this.hf) this.hf.destroy();
    this.hf = null;
  }
}

// ------------------------------------------------------------- helpers

function uniqueName(name, taken) {
  let n = (name || 'Sheet').slice(0, 31) || 'Sheet';
  if (!taken.includes(n)) return n;
  let i = 2;
  while (taken.includes(n + ' (' + i + ')')) i++;
  return n + ' (' + i + ')';
}

// Model sheet -> the dense 2-D array HyperFormula wants at build time.
function sheetToArray(sheet) {
  const rows = [];
  for (let r = 1; r <= sheet.maxRow; r++) {
    const row = new Array(sheet.maxCol);
    let any = false;
    for (let c = 1; c <= sheet.maxCol; c++) {
      const cell = sheet.cells[key(r, c)];
      if (!cell) {
        row[c - 1] = null;
        continue;
      }
      if (cell.f) {
        row[c - 1] = '=' + cell.f;
        any = true;
      } else if (cell.v !== undefined && cell.v !== null) {
        row[c - 1] = cell.v;
        any = true;
      } else {
        row[c - 1] = null;
      }
    }
    rows.push(any ? row : []);
  }
  return rows;
}

// Engine value -> {value, type} in model terms.
function describe(v) {
  if (v === null || v === undefined) return { value: null, type: 'z' };
  if (v instanceof DetailedCellError) return { value: v.value, type: 'e' };
  if (typeof v === 'number') return { value: v, type: 'n' };
  if (typeof v === 'boolean') return { value: v, type: 'b' };
  if (typeof v === 'string') return { value: v, type: 's' };
  // Array formulas hand back nested arrays for the top-left cell; show the
  // corner value, which is what a non-spilling consumer expects.
  if (Array.isArray(v)) return describe(v[0] && v[0][0]);
  return { value: String(v), type: 's' };
}

function applyValue(cell, raw) {
  const { value, type } = describe(raw);
  if (value === null) delete cell.v;
  else cell.v = value;
  if (type === 'z') delete cell.t;
  else cell.t = type;
}

export { describe };
