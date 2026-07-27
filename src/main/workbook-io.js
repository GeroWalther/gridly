'use strict';

const path = require('path');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { toHex, toArgb } = require('./colors');
const { StyleTable, emptySheet, emptyWorkbook, DEFAULT_ROW_HEIGHT } = require('../shared/model');
const { rcToA1, parseRange, key } = require('../shared/addr');

// Excel's day 0. Serial dates are days since 1899-12-30 (the off-by-one is
// Lotus 1-2-3's phantom 1900 leap year, which Excel keeps for compatibility).
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86400000;

function dateToSerial(d) {
  return (d.getTime() - EXCEL_EPOCH_MS) / MS_PER_DAY;
}

function serialToDate(n) {
  return new Date(Math.round(n * MS_PER_DAY) + EXCEL_EPOCH_MS);
}

// A number format that renders as a date/time, i.e. one we must hand ExcelJS
// as a real Date rather than a bare number. Quoted literals and the colour /
// condition brackets are stripped first so that e.g. [Red]0.00 isn't mistaken
// for a date because of the "d" in "Red".
function isDateFormat(numFmt) {
  if (!numFmt || numFmt === 'General') return false;
  const stripped = String(numFmt)
    .replace(/\[[^\]]*\]/g, '')
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '');
  return /[ymdhs]/i.test(stripped);
}

// ---------------------------------------------------------------- reading

function styleFromCell(cell) {
  const s = cell.style || {};
  const f = s.font || {};
  const a = s.alignment || {};

  let fill = null;
  if (s.fill && s.fill.type === 'pattern' && s.fill.pattern !== 'none') {
    // Excel's "solid" pattern paints with fgColor, which reads backwards but
    // is what the spec says. Other patterns are approximated by their bg.
    fill = s.fill.pattern === 'solid' ? toHex(s.fill.fgColor) : toHex(s.fill.bgColor) || toHex(s.fill.fgColor);
  }

  let border = null;
  if (s.border) {
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const b = s.border[side];
      if (b && b.style) {
        border = border || {};
        border[side] = { style: b.style, color: toHex(b.color) || '#000000' };
      }
    }
  }

  return {
    font: {
      name: f.name || 'Calibri',
      size: f.size || 11,
      bold: !!f.bold,
      italic: !!f.italic,
      underline: !!f.underline,
      strike: !!f.strike,
      color: toHex(f.color) || '#000000',
    },
    fill,
    align: {
      h: a.horizontal || null,
      v: a.vertical || 'bottom',
      wrap: !!a.wrapText,
      indent: a.indent || 0,
    },
    border,
    numFmt: s.numFmt || 'General',
  };
}

// ExcelJS cell -> {v, f, t}. `t` mirrors SheetJS's type tags: n number,
// s string, b boolean, e error. Dates become serial numbers so that the
// formula engine and the number formatter both see one numeric type.
function valueFromCell(cell, numFmt) {
  const v = cell.value;

  if (v === null || v === undefined) return { v: null, f: null, t: 'z' };

  if (cell.type === ExcelJS.ValueType.Formula || (v && (v.formula || v.sharedFormula))) {
    const formula = cell.formula || null;
    let result = v.result;
    let t = 'n';
    if (result instanceof Date) {
      result = dateToSerial(result);
    } else if (result && typeof result === 'object' && result.error) {
      result = result.error;
      t = 'e';
    } else if (typeof result === 'string') {
      t = 's';
    } else if (typeof result === 'boolean') {
      t = 'b';
    } else if (result === null || result === undefined) {
      result = null;
      t = 'z';
    }
    return { v: result, f: formula, t };
  }

  if (v instanceof Date) return { v: dateToSerial(v), f: null, t: 'n' };
  if (typeof v === 'number') return { v, f: null, t: 'n' };
  if (typeof v === 'boolean') return { v, f: null, t: 'b' };
  if (typeof v === 'string') return { v, f: null, t: 's' };

  if (typeof v === 'object') {
    // Rich text collapses to its concatenated plain text; per-run formatting
    // is not represented in the model, and the cell keeps its base style.
    if (Array.isArray(v.richText)) {
      return { v: v.richText.map((r) => r.text).join(''), f: null, t: 's' };
    }
    if (v.error) return { v: v.error, f: null, t: 'e' };
    if (v.hyperlink !== undefined) {
      return { v: v.text != null ? String(v.text) : v.hyperlink, f: null, t: 's', link: v.hyperlink };
    }
    if (v.text !== undefined) return { v: String(v.text), f: null, t: 's' };
  }

  return { v: String(v), f: null, t: 's' };
}

async function readXlsx(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  const styles = new StyleTable();
  const model = emptyWorkbook();
  model.sheets = [];
  model.filePath = filePath;
  model.fileType = path.extname(filePath).slice(1).toLowerCase() || 'xlsx';

  wb.eachSheet((ws) => {
    const sheet = emptySheet(ws.name);
    sheet.tabColor = ws.properties && ws.properties.tabColor ? toHex(ws.properties.tabColor) : null;

    if (ws.properties && ws.properties.defaultRowHeight) {
      sheet.defaultRowHeight = ws.properties.defaultRowHeight;
    }
    if (ws.properties && ws.properties.defaultColWidth) {
      sheet.defaultColWidth = ws.properties.defaultColWidth;
    }

    (ws.columns || []).forEach((col, i) => {
      if (!col) return;
      const c = {};
      if (col.width != null) c.width = col.width;
      if (col.hidden) c.hidden = true;
      if (Object.keys(c).length) sheet.cols[i + 1] = c;
    });

    const view = (ws.views || [])[0];
    if (view && view.state === 'frozen') {
      sheet.frozen = { row: view.ySplit || 0, col: view.xSplit || 0 };
    }

    for (const range of ws.model && ws.model.merges ? ws.model.merges : []) {
      sheet.merges.push(range);
    }

    let maxRow = 0;
    let maxCol = 0;

    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (row.height != null && row.height !== sheet.defaultRowHeight) {
        sheet.rows[rowNumber] = { height: row.height };
      }
      if (row.hidden) {
        sheet.rows[rowNumber] = Object.assign(sheet.rows[rowNumber] || {}, { hidden: true });
      }

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const st = styleFromCell(cell);
        const si = styles.intern(st);
        const parsed = valueFromCell(cell, st.numFmt);

        // Skip cells that are both empty and unstyled — files routinely
        // materialise thousands of those and they cost us nothing to omit.
        if (parsed.v === null && !parsed.f && si === 0) return;

        const entry = { s: si };
        if (parsed.v !== null) entry.v = parsed.v;
        if (parsed.f) entry.f = parsed.f;
        if (parsed.t && parsed.t !== 'z') entry.t = parsed.t;
        if (parsed.link) entry.link = parsed.link;

        sheet.cells[key(rowNumber, colNumber)] = entry;
        if (rowNumber > maxRow) maxRow = rowNumber;
        if (colNumber > maxCol) maxCol = colNumber;
      });
    });

    sheet.maxRow = maxRow;
    sheet.maxCol = Math.max(maxCol, ws.columnCount || 0);
    model.sheets.push(sheet);
  });

  if (!model.sheets.length) model.sheets.push(emptySheet('Sheet1'));

  try {
    const dn = wb.definedNames && wb.definedNames.model ? wb.definedNames.model : [];
    model.definedNames = dn
      .filter((d) => d && d.name && Array.isArray(d.ranges) && d.ranges.length)
      .map((d) => ({ name: d.name, ref: d.ranges[0] }));
  } catch (_) {
    model.definedNames = [];
  }

  model.styles = styles.toJSON();

  // Macro-enabled workbooks are opened but never written back: ExcelJS drops
  // the vbaProject part, which would silently destroy the macros.
  if (model.fileType === 'xlsm') {
    model.readOnly = true;
    model.readOnlyReason = 'Macro-enabled workbooks open read-only so their macros are not lost. Use Save As to write a .xlsx copy.';
  }

  return model;
}

// Legacy .xls has no ExcelJS reader, so SheetJS handles it. Values and
// formulas come through; formatting does not, hence read-only.
function readLegacy(filePath, type) {
  const wb = XLSX.readFile(filePath, { cellFormula: true, cellStyles: false, cellDates: false });
  const styles = new StyleTable();
  const model = emptyWorkbook();
  model.sheets = [];
  model.filePath = filePath;
  model.fileType = type;

  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const sheet = emptySheet(name);
    if (!ws) {
      model.sheets.push(sheet);
      continue;
    }

    const ref = ws['!ref'] ? parseRange(ws['!ref']) : null;
    for (const addr of Object.keys(ws)) {
      if (addr[0] === '!') continue;
      const rc = parseRange(addr);
      if (!rc) continue;
      const c = ws[addr];
      if (c.v === undefined && !c.f) continue;

      const entry = { s: 0 };
      if (c.f) entry.f = c.f;
      if (c.v !== undefined && c.v !== null) {
        entry.v = c.v instanceof Date ? dateToSerial(c.v) : c.v;
        entry.t = c.t === 'd' ? 'n' : c.t || 's';
      }
      if (c.z) entry.s = styles.intern({ numFmt: c.z });

      sheet.cells[key(rc.r1, rc.c1)] = entry;
      if (rc.r1 > sheet.maxRow) sheet.maxRow = rc.r1;
      if (rc.c1 > sheet.maxCol) sheet.maxCol = rc.c1;
    }

    (ws['!merges'] || []).forEach((m) => {
      sheet.merges.push(rcToA1(m.s.r + 1, m.s.c + 1) + ':' + rcToA1(m.e.r + 1, m.e.c + 1));
    });
    (ws['!cols'] || []).forEach((col, i) => {
      if (col && col.wch != null) sheet.cols[i + 1] = { width: col.wch };
    });
    (ws['!rows'] || []).forEach((row, i) => {
      if (row && row.hpt != null) sheet.rows[i + 1] = { height: row.hpt };
    });

    if (ref) {
      sheet.maxRow = Math.max(sheet.maxRow, ref.r2);
      sheet.maxCol = Math.max(sheet.maxCol, ref.c2);
    }
    model.sheets.push(sheet);
  }

  if (!model.sheets.length) model.sheets.push(emptySheet('Sheet1'));
  model.styles = styles.toJSON();
  model.readOnly = true;
  model.readOnlyReason =
    'Excel 97-2003 (.xls) files open read-only. Use Save As to write a .xlsx copy.';
  return model;
}

function readCsv(filePath) {
  const wb = XLSX.readFile(filePath, { type: 'file', raw: false, cellDates: false });
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];
  const model = emptyWorkbook();
  model.filePath = filePath;
  model.fileType = 'csv';
  const sheet = emptySheet(path.basename(filePath, path.extname(filePath)).slice(0, 31) || 'Sheet1');

  for (const addr of Object.keys(ws || {})) {
    if (addr[0] === '!') continue;
    const rc = parseRange(addr);
    if (!rc) continue;
    const c = ws[addr];
    if (c.v === undefined || c.v === null) continue;
    sheet.cells[key(rc.r1, rc.c1)] = { v: c.v, s: 0, t: c.t || 's' };
    if (rc.r1 > sheet.maxRow) sheet.maxRow = rc.r1;
    if (rc.c1 > sheet.maxCol) sheet.maxCol = rc.c1;
  }

  model.sheets = [sheet];
  return model;
}

async function readWorkbook(filePath) {
  const ext = path.extname(filePath).slice(1).toLowerCase();
  if (ext === 'csv' || ext === 'txt' || ext === 'tsv') return readCsv(filePath);
  if (ext === 'xls') return readLegacy(filePath, 'xls');
  return readXlsx(filePath);
}

// ---------------------------------------------------------------- writing

function applyStyleToCell(cell, style) {
  if (!style) return;
  const f = style.font || {};
  cell.font = {
    name: f.name || 'Calibri',
    size: f.size || 11,
    bold: !!f.bold,
    italic: !!f.italic,
    underline: !!f.underline,
    strike: !!f.strike,
    color: { argb: toArgb(f.color || '#000000') },
  };

  if (style.fill) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: toArgb(style.fill) },
      bgColor: { argb: toArgb(style.fill) },
    };
  }

  const a = style.align || {};
  const alignment = {};
  if (a.h) alignment.horizontal = a.h;
  if (a.v && a.v !== 'bottom') alignment.vertical = a.v;
  if (a.wrap) alignment.wrapText = true;
  if (a.indent) alignment.indent = a.indent;
  if (Object.keys(alignment).length) cell.alignment = alignment;

  if (style.border) {
    const b = {};
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const s = style.border[side];
      if (s && s.style) b[side] = { style: s.style, color: { argb: toArgb(s.color || '#000000') } };
    }
    if (Object.keys(b).length) cell.border = b;
  }

  if (style.numFmt && style.numFmt !== 'General') cell.numFmt = style.numFmt;
}

async function writeWorkbook(model, filePath) {
  const ext = path.extname(filePath).slice(1).toLowerCase();
  if (ext === 'csv') return writeCsv(model, filePath);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Gridly';
  wb.lastModifiedBy = 'Gridly';
  wb.created = new Date();
  wb.modified = new Date();

  const styles = model.styles || [];

  for (const sheet of model.sheets) {
    const ws = wb.addWorksheet(sheet.name, {
      properties: {
        tabColor: sheet.tabColor ? { argb: toArgb(sheet.tabColor) } : undefined,
        defaultRowHeight: sheet.defaultRowHeight || DEFAULT_ROW_HEIGHT,
      },
      views:
        sheet.frozen && (sheet.frozen.row || sheet.frozen.col)
          ? [{ state: 'frozen', xSplit: sheet.frozen.col || 0, ySplit: sheet.frozen.row || 0 }]
          : undefined,
    });

    for (const colStr of Object.keys(sheet.cols || {})) {
      const spec = sheet.cols[colStr];
      const col = ws.getColumn(+colStr);
      if (spec.width != null) col.width = spec.width;
      if (spec.hidden) col.hidden = true;
    }

    for (const rowStr of Object.keys(sheet.rows || {})) {
      const spec = sheet.rows[rowStr];
      const row = ws.getRow(+rowStr);
      if (spec.height != null) row.height = spec.height;
      if (spec.hidden) row.hidden = true;
    }

    for (const k of Object.keys(sheet.cells)) {
      const comma = k.indexOf(',');
      const r = +k.slice(0, comma);
      const c = +k.slice(comma + 1);
      const cellModel = sheet.cells[k];
      const cell = ws.getCell(r, c);
      const style = styles[cellModel.s || 0];

      if (cellModel.f) {
        // The cached result travels with the formula so other apps can show a
        // value before they recalculate.
        let result = cellModel.v;
        if (result === undefined) result = null;
        if (cellModel.t === 'e' && typeof result === 'string') result = { error: result };
        cell.value = { formula: cellModel.f, result };
      } else if (cellModel.v !== undefined && cellModel.v !== null) {
        if (typeof cellModel.v === 'number' && style && isDateFormat(style.numFmt)) {
          cell.value = serialToDate(cellModel.v);
        } else if (cellModel.t === 'e') {
          cell.value = { error: cellModel.v };
        } else {
          cell.value = cellModel.v;
        }
      }

      if (cellModel.link) {
        cell.value = { text: cellModel.v != null ? String(cellModel.v) : cellModel.link, hyperlink: cellModel.link };
      }

      applyStyleToCell(cell, style);
    }

    for (const range of sheet.merges || []) {
      try {
        ws.mergeCells(range);
      } catch (_) {
        // Overlapping or already-merged ranges are skipped rather than
        // aborting the whole save.
      }
    }
  }

  for (const dn of model.definedNames || []) {
    try {
      wb.definedNames.add(dn.ref, dn.name);
    } catch (_) {
      /* invalid name refs are dropped */
    }
  }

  await wb.xlsx.writeFile(filePath);
}

async function writeCsv(model, filePath) {
  const fs = require('fs').promises;
  const sheet = model.sheets[model.activeSheet || 0] || model.sheets[0];
  const lines = [];
  for (let r = 1; r <= sheet.maxRow; r++) {
    const row = [];
    for (let c = 1; c <= sheet.maxCol; c++) {
      const cell = sheet.cells[key(r, c)];
      let v = cell && cell.v != null ? cell.v : '';
      v = String(v);
      if (/[",\n\r]/.test(v)) v = '"' + v.replace(/"/g, '""') + '"';
      row.push(v);
    }
    lines.push(row.join(','));
  }
  await fs.writeFile(filePath, lines.join('\r\n'), 'utf8');
}

module.exports = { readWorkbook, writeWorkbook, dateToSerial, serialToDate, isDateFormat };
