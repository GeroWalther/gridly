import { Doc } from './doc.js';
import { Grid, setGridTheme } from './grid.js';
import { Editor } from './editor.js';
import { shiftFormula } from './formula-ref.js';
import { formatValue, isDateFormat, widthToPx } from './format.js';
import { key, rcToA1, parseRange, rangeToA1 } from '../shared/addr.js';

const api = window.gridly;

const $ = (id) => document.getElementById(id);

const doc = new Doc();
const gridHost = $('gridhost');
const grid = new Grid(gridHost, doc);
const editor = new Editor({
  grid,
  doc,
  host: gridHost,
  formulaInput: $('formulaInput'),
  onCommit: (row, col, move) => {
    const delta = { up: [-1, 0], down: [1, 0], left: [0, -1], right: [0, 1] }[move] || [0, 0];
    grid.anchor = null;
    grid.selectCell(Math.max(1, row + delta[0]), Math.max(1, col + delta[1]));
    grid.focus();
  },
});

// The last block copied from within the app, kept so that an internal paste
// can carry formulas and formatting that the system clipboard cannot.
let internalClipboard = null;

// ---------------------------------------------------------------- refresh

let suppressRefresh = false;

doc.onChange(() => {
  if (suppressRefresh) return;
  grid.layout();
  grid.render();
  renderSheetTabs();
  updateStatus();
  updateToolbar();
  updateTitle();
  editor.syncFromSelection();
});

grid.onSelectionChange = () => {
  updateNameBox();
  updateStatus();
  updateToolbar();
  editor.syncFromSelection();
};

grid.onEditRequest = (row, col, initial) => {
  if (doc.model.readOnly) return flashReadOnly();
  editor.begin(row, col, { mode: initial == null ? 'edit' : 'enter', initial });
};

grid.onKeyDown = (e) => handleGridKey(e);
grid.onStructuralChange = (change) => applyResize(change);
grid.onFill = (source, target) => fillRange(source, target);
grid.onContextMenu = (e, hit) => showContextMenu(e, hit);

function refreshAll() {
  grid.layout();
  grid.render();
  renderSheetTabs();
  updateNameBox();
  updateStatus();
  updateToolbar();
  updateTitle();
  editor.syncFromSelection();
}

// ------------------------------------------------------------ title/status

function updateTitle() {
  const name = doc.model.filePath ? doc.model.filePath.split('/').pop() : 'Untitled';
  api.setTitle({
    title: name + (doc.dirty ? ' — Edited' : ''),
    filePath: doc.model.filePath,
    edited: doc.dirty,
  });
  $('readonlyBadge').hidden = !doc.model.readOnly;
}

function updateNameBox() {
  const sel = grid.selection;
  const a = grid.active;
  $('nameBox').value =
    sel.r1 === sel.r2 && sel.c1 === sel.c2 ? rcToA1(a.row, a.col) : `${rangeToA1(sel)}`;
}

// Excel's status-bar summary of the selection.
function updateStatus() {
  const sel = grid.selection;
  const sheet = doc.sheet;
  let count = 0;
  let numCount = 0;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;

  // Guard against summarising a whole-column selection cell by cell.
  const r2 = Math.min(sel.r2, Math.max(sheet.maxRow, sel.r1));
  const c2 = Math.min(sel.c2, Math.max(sheet.maxCol, sel.c1));

  for (let r = sel.r1; r <= r2; r++) {
    for (let c = sel.c1; c <= c2; c++) {
      const cell = sheet.cells[key(r, c)];
      if (!cell || cell.v === undefined || cell.v === null || cell.v === '') continue;
      count++;
      if (typeof cell.v === 'number' && cell.t !== 'e') {
        numCount++;
        sum += cell.v;
        if (cell.v < min) min = cell.v;
        if (cell.v > max) max = cell.v;
      }
    }
  }

  const parts = [];
  if (count) parts.push('Count: ' + count);
  if (numCount) {
    parts.push('Sum: ' + trim(sum));
    parts.push('Average: ' + trim(sum / numCount));
    parts.push('Min: ' + trim(min));
    parts.push('Max: ' + trim(max));
  }
  $('statusStats').textContent = parts.join('   ');
  $('statusLeft').textContent = doc.model.readOnly ? doc.model.readOnlyReason || 'Read-only' : 'Ready';
}

function trim(n) {
  if (!isFinite(n)) return '—';
  const r = Math.round(n * 1e10) / 1e10;
  return String(r);
}

function flashReadOnly() {
  $('statusLeft').textContent = doc.model.readOnlyReason || 'This workbook is read-only.';
}

// ------------------------------------------------------------- sheet tabs

function renderSheetTabs() {
  const host = $('sheetTabs');
  host.textContent = '';
  doc.model.sheets.forEach((sheet, i) => {
    const tab = document.createElement('button');
    tab.className = 'sheet-tab' + (i === doc.activeSheet ? ' active' : '');
    if (sheet.tabColor) {
      const dot = document.createElement('span');
      dot.className = 'dot';
      dot.style.background = sheet.tabColor;
      tab.appendChild(dot);
    }
    tab.appendChild(document.createTextNode(sheet.name));
    tab.onclick = () => switchSheet(i);
    tab.ondblclick = () => beginRenameTab(tab, i);
    tab.oncontextmenu = (e) => showSheetMenu(e, i);
    host.appendChild(tab);
  });
}

function switchSheet(i) {
  if (i === doc.activeSheet) return;
  if (editor.open) editor.cancel();
  grid.viewBySheet.set(doc.activeSheet, {
    scrollLeft: grid.scroller.scrollLeft,
    scrollTop: grid.scroller.scrollTop,
    selection: Object.assign({}, grid.selection),
    active: Object.assign({}, grid.active),
  });
  doc.activeSheet = i;
  const saved = grid.viewBySheet.get(i);
  grid.selection = saved ? saved.selection : { r1: 1, c1: 1, r2: 1, c2: 1 };
  grid.active = saved ? saved.active : { row: 1, col: 1 };
  grid.layout();
  grid.scroller.scrollLeft = saved ? saved.scrollLeft : 0;
  grid.scroller.scrollTop = saved ? saved.scrollTop : 0;
  refreshAll();
  grid.focus();
}

function beginRenameTab(tab, index) {
  if (doc.model.readOnly) return flashReadOnly();
  const input = document.createElement('input');
  input.value = doc.model.sheets[index].name;
  tab.textContent = '';
  tab.appendChild(input);
  input.focus();
  input.select();
  const finish = (commit) => {
    const name = input.value.trim();
    if (commit && name && name !== doc.model.sheets[index].name) {
      doc.transact('Rename Sheet', (tx) => tx.apply({ type: 'sheetRename', index, name }));
    }
    renderSheetTabs();
  };
  input.onblur = () => finish(true);
  input.onkeydown = (e) => {
    e.stopPropagation();
    if (e.key === 'Enter') finish(true);
    else if (e.key === 'Escape') finish(false);
  };
}

$('addSheet').onclick = () => {
  if (doc.model.readOnly) return flashReadOnly();
  const index = doc.model.sheets.length;
  const name = nextSheetName();
  doc.transact('New Sheet', (tx) => tx.apply({ type: 'sheetAdd', index, name }));
  switchSheet(index);
};

function nextSheetName() {
  const taken = new Set(doc.model.sheets.map((s) => s.name));
  let n = doc.model.sheets.length + 1;
  while (taken.has('Sheet' + n)) n++;
  return 'Sheet' + n;
}

function showSheetMenu(e, index) {
  e.preventDefault();
  openMenu(e.clientX, e.clientY, [
    { label: 'Rename', action: () => beginRenameTab($('sheetTabs').children[index], index) },
    {
      label: 'Duplicate',
      action: () => {
        const copy = JSON.parse(JSON.stringify(doc.model.sheets[index]));
        copy.name = copy.name + ' copy';
        doc.transact('Duplicate Sheet', (tx) => tx.apply({ type: 'sheetAdd', index: index + 1, sheet: copy }));
        switchSheet(index + 1);
      },
    },
    { sep: true },
    {
      label: 'Tab Colour…',
      action: () => {
        const picker = document.createElement('input');
        picker.type = 'color';
        picker.value = doc.model.sheets[index].tabColor || '#4472c4';
        picker.onchange = () => doc.transact('Tab Colour', (tx) => tx.apply({ type: 'sheetColor', index, color: picker.value }));
        picker.click();
      },
    },
    { sep: true },
    {
      label: 'Delete Sheet',
      disabled: doc.model.sheets.length <= 1,
      action: () => {
        doc.transact('Delete Sheet', (tx) => tx.apply({ type: 'sheetRemove', index }));
        if (doc.activeSheet >= doc.model.sheets.length) doc.activeSheet = doc.model.sheets.length - 1;
        refreshAll();
      },
    },
  ]);
}

// ------------------------------------------------------------------ menus

function openMenu(x, y, items) {
  const menu = $('contextMenu');
  menu.textContent = '';
  for (const item of items) {
    if (item.sep) {
      const sep = document.createElement('div');
      sep.className = 'ctx-sep';
      menu.appendChild(sep);
      continue;
    }
    const btn = document.createElement('button');
    btn.textContent = item.label;
    btn.disabled = !!item.disabled;
    btn.onclick = () => {
      closeMenu();
      item.action();
    };
    menu.appendChild(btn);
  }
  menu.hidden = false;
  const rect = menu.getBoundingClientRect();
  menu.style.left = Math.min(x, window.innerWidth - rect.width - 8) + 'px';
  menu.style.top = Math.min(y, window.innerHeight - rect.height - 8) + 'px';
}

function closeMenu() {
  $('contextMenu').hidden = true;
}

window.addEventListener('mousedown', (e) => {
  if (!$('contextMenu').hidden && !$('contextMenu').contains(e.target)) closeMenu();
  const bm = $('borderMenu');
  if (!bm.hidden && !bm.contains(e.target) && e.target !== $('borderBtn')) bm.hidden = true;
});

function showContextMenu(e, hit) {
  const ro = doc.model.readOnly;
  const items = [
    { label: 'Cut', disabled: ro, action: () => doCopy(true) },
    { label: 'Copy', action: () => doCopy(false) },
    { label: 'Paste', disabled: ro, action: () => pasteFromClipboard(false) },
    { sep: true },
  ];

  if (hit.kind === 'rowHeader') {
    const n = grid.selection.r2 - grid.selection.r1 + 1;
    items.push(
      { label: `Insert ${n} Row${n > 1 ? 's' : ''} Above`, disabled: ro, action: () => insertRows() },
      { label: `Delete ${n} Row${n > 1 ? 's' : ''}`, disabled: ro, action: () => deleteRows() },
      { label: 'Row Height…', disabled: ro, action: () => promptRowHeight() },
      { sep: true }
    );
  } else if (hit.kind === 'colHeader') {
    const n = grid.selection.c2 - grid.selection.c1 + 1;
    items.push(
      { label: `Insert ${n} Column${n > 1 ? 's' : ''} Left`, disabled: ro, action: () => insertCols() },
      { label: `Delete ${n} Column${n > 1 ? 's' : ''}`, disabled: ro, action: () => deleteCols() },
      { label: 'Column Width…', disabled: ro, action: () => promptColWidth() },
      { label: 'Autofit Width', disabled: ro, action: () => autofitSelection() },
      { sep: true }
    );
  } else {
    items.push(
      { label: 'Insert Rows Above', disabled: ro, action: () => insertRows() },
      { label: 'Insert Columns Left', disabled: ro, action: () => insertCols() },
      { label: 'Delete Rows', disabled: ro, action: () => deleteRows() },
      { label: 'Delete Columns', disabled: ro, action: () => deleteCols() },
      { sep: true }
    );
  }

  items.push(
    { label: 'Clear Contents', disabled: ro, action: () => doc.clearRange(grid.selection, 'contents') },
    { label: 'Clear Formatting', disabled: ro, action: () => doc.clearRange(grid.selection, 'formats') },
    { sep: true },
    { label: isMerged() ? 'Unmerge Cells' : 'Merge Cells', disabled: ro, action: () => toggleMerge() },
    { label: 'Freeze Panes Here', disabled: ro, action: () => freezeAtSelection() }
  );

  openMenu(e.clientX, e.clientY, items);
}

// ------------------------------------------------------------- edit paths

function handleGridKey(e) {
  const meta = e.metaKey || e.ctrlKey;

  if (meta) return; // menu accelerators own these

  if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault();
    if (doc.model.readOnly) return flashReadOnly();
    doc.clearRange(grid.selection, 'contents');
    return;
  }

  // A printable character starts editing with that character, as in Excel.
  if (e.key.length === 1 && !e.altKey) {
    if (doc.model.readOnly) return flashReadOnly();
    e.preventDefault();
    editor.begin(grid.active.row, grid.active.col, { mode: 'enter', initial: e.key });
  }
}

function applyResize(change) {
  if (doc.model.readOnly) return flashReadOnly();
  const si = doc.activeSheet;
  const sheet = doc.sheet;
  const sel = grid.selection;

  if (change.type === 'colWidth') {
    // Resizing one of several selected columns resizes them all.
    const cols =
      change.col >= sel.c1 && change.col <= sel.c2 && sel.r2 >= grid.nRows
        ? rangeList(sel.c1, sel.c2)
        : [change.col];
    doc.transact('Column Width', (tx) =>
      tx.apply({
        type: 'colWidth',
        sheetIndex: si,
        cols: cols.map((c) => ({ col: c, spec: Object.assign({}, sheet.cols[c], { width: change.width }) })),
      })
    );
  } else {
    const rows =
      change.row >= sel.r1 && change.row <= sel.r2 && sel.c2 >= grid.nCols
        ? rangeList(sel.r1, sel.r2)
        : [change.row];
    doc.transact('Row Height', (tx) =>
      tx.apply({
        type: 'rowHeight',
        sheetIndex: si,
        rows: rows.map((r) => ({ row: r, spec: Object.assign({}, sheet.rows[r], { height: change.height }) })),
      })
    );
  }
}

function rangeList(a, b) {
  const out = [];
  for (let i = a; i <= b; i++) out.push(i);
  return out;
}

function insertRows() {
  if (doc.model.readOnly) return flashReadOnly();
  const sel = grid.selection;
  const count = sel.r2 - sel.r1 + 1;
  doc.transact('Insert Rows', (tx) => tx.apply({ type: 'insertRows', sheetIndex: doc.activeSheet, at: sel.r1, count }));
}

function deleteRows() {
  if (doc.model.readOnly) return flashReadOnly();
  const sel = grid.selection;
  const count = sel.r2 - sel.r1 + 1;
  doc.transact('Delete Rows', (tx) => tx.apply({ type: 'deleteRows', sheetIndex: doc.activeSheet, at: sel.r1, count }));
}

function insertCols() {
  if (doc.model.readOnly) return flashReadOnly();
  const sel = grid.selection;
  const count = sel.c2 - sel.c1 + 1;
  doc.transact('Insert Columns', (tx) => tx.apply({ type: 'insertCols', sheetIndex: doc.activeSheet, at: sel.c1, count }));
}

function deleteCols() {
  if (doc.model.readOnly) return flashReadOnly();
  const sel = grid.selection;
  const count = sel.c2 - sel.c1 + 1;
  doc.transact('Delete Columns', (tx) => tx.apply({ type: 'deleteCols', sheetIndex: doc.activeSheet, at: sel.c1, count }));
}

function promptColWidth() {
  const sel = grid.selection;
  const cur = doc.sheet.cols[sel.c1];
  const value = window.prompt('Column width (characters):', String(cur && cur.width != null ? cur.width : doc.sheet.defaultColWidth));
  if (value === null) return;
  const w = parseFloat(value);
  if (!isFinite(w) || w < 0) return;
  doc.transact('Column Width', (tx) =>
    tx.apply({
      type: 'colWidth',
      sheetIndex: doc.activeSheet,
      cols: rangeList(sel.c1, sel.c2).map((c) => ({ col: c, spec: Object.assign({}, doc.sheet.cols[c], { width: w }) })),
    })
  );
}

function promptRowHeight() {
  const sel = grid.selection;
  const cur = doc.sheet.rows[sel.r1];
  const value = window.prompt('Row height (points):', String(cur && cur.height != null ? cur.height : doc.sheet.defaultRowHeight));
  if (value === null) return;
  const h = parseFloat(value);
  if (!isFinite(h) || h < 0) return;
  doc.transact('Row Height', (tx) =>
    tx.apply({
      type: 'rowHeight',
      sheetIndex: doc.activeSheet,
      rows: rangeList(sel.r1, sel.r2).map((r) => ({ row: r, spec: Object.assign({}, doc.sheet.rows[r], { height: h }) })),
    })
  );
}

function autofitSelection() {
  const sel = grid.selection;
  const cols = rangeList(sel.c1, Math.min(sel.c2, sel.c1 + 200));
  doc.transact('Autofit', (tx) =>
    tx.apply({
      type: 'colWidth',
      sheetIndex: doc.activeSheet,
      cols: cols.map((c) => ({ col: c, spec: Object.assign({}, doc.sheet.cols[c], { width: grid.autofitColumn(c) }) })),
    })
  );
}

function isMerged() {
  const sel = grid.selection;
  return (doc.sheet.merges || []).some((m) => {
    const r = parseRange(m);
    return r && r.r1 === sel.r1 && r.c1 === sel.c1 && r.r2 === sel.r2 && r.c2 === sel.c2;
  });
}

function toggleMerge() {
  if (doc.model.readOnly) return flashReadOnly();
  const sel = grid.selection;
  const ref = rangeToA1(sel);
  const merges = doc.sheet.merges || [];
  let next;
  if (isMerged()) {
    next = merges.filter((m) => m !== ref);
  } else {
    if (sel.r1 === sel.r2 && sel.c1 === sel.c2) return;
    // Drop any merge that overlaps the new one, as Excel does.
    next = merges.filter((m) => {
      const r = parseRange(m);
      return !r || r.r2 < sel.r1 || r.r1 > sel.r2 || r.c2 < sel.c1 || r.c1 > sel.c2;
    });
    next.push(ref);
  }
  doc.transact('Merge', (tx) => tx.apply({ type: 'merges', sheetIndex: doc.activeSheet, merges: next }));
}

function freezeAtSelection() {
  if (doc.model.readOnly) return flashReadOnly();
  const a = grid.active;
  doc.transact('Freeze Panes', (tx) =>
    tx.apply({ type: 'frozen', sheetIndex: doc.activeSheet, frozen: { row: a.row - 1, col: a.col - 1 } })
  );
  refreshAll();
}

function unfreeze() {
  doc.transact('Unfreeze Panes', (tx) => tx.apply({ type: 'frozen', sheetIndex: doc.activeSheet, frozen: { row: 0, col: 0 } }));
  refreshAll();
}

// ----------------------------------------------------------------- fill

// Drag-fill: a numeric or date run is extrapolated, anything else repeats.
function fillRange(source, target) {
  if (doc.model.readOnly) return flashReadOnly();
  const sheet = doc.sheet;
  const down = target.r2 > source.r2 || target.r1 < source.r1;
  const cells = [];

  const srcRows = source.r2 - source.r1 + 1;
  const srcCols = source.c2 - source.c1 + 1;

  for (let r = target.r1; r <= target.r2; r++) {
    for (let c = target.c1; c <= target.c2; c++) {
      if (r >= source.r1 && r <= source.r2 && c >= source.c1 && c <= source.c2) continue;

      // Index of this cell within the repeating source pattern.
      const dr = r - source.r1;
      const dc = c - source.c1;
      const sr = source.r1 + ((((dr % srcRows) + srcRows) % srcRows));
      const sc = source.c1 + ((((dc % srcCols) + srcCols) % srcCols));
      const src = sheet.cells[key(sr, sc)];
      if (!src) {
        cells.push({ row: r, col: c, cell: null });
        continue;
      }

      const next = Object.assign({}, src);
      if (src.f) {
        next.f = shiftFormula(src.f, r - sr, c - sc);
        delete next.v;
      } else if (typeof src.v === 'number') {
        const step = seriesStep(sheet, source, down);
        if (step) next.v = src.v + step * (down ? r - sr : c - sc);
      }
      cells.push({ row: r, col: c, cell: next });
    }
  }

  if (!cells.length) return;
  doc.transact('Fill', (tx) => tx.apply({ type: 'cells', sheetIndex: doc.activeSheet, cells }));
}

// A single numeric cell repeats (Excel's behaviour); two or more evenly
// spaced numbers define a step that is extrapolated.
function seriesStep(sheet, source, down) {
  const vals = [];
  if (down) {
    for (let r = source.r1; r <= source.r2; r++) {
      const cell = sheet.cells[key(r, source.c1)];
      if (!cell || typeof cell.v !== 'number' || cell.f) return null;
      vals.push(cell.v);
    }
  } else {
    for (let c = source.c1; c <= source.c2; c++) {
      const cell = sheet.cells[key(source.r1, c)];
      if (!cell || typeof cell.v !== 'number' || cell.f) return null;
      vals.push(cell.v);
    }
  }
  if (vals.length < 2) {
    // A lone date still increments by a day, which is what people expect.
    const cell = sheet.cells[key(source.r1, source.c1)];
    const style = cell ? doc.styles.get(cell.s || 0) : null;
    return style && isDateFormat(style.numFmt) ? 1 : 0;
  }
  const step = vals[1] - vals[0];
  for (let i = 2; i < vals.length; i++) {
    if (Math.abs(vals[i] - vals[i - 1] - step) > 1e-9) return 0;
  }
  return step;
}

// ------------------------------------------------------------- clipboard

function selectedBlock() {
  const sel = grid.expandToMerges(grid.selection);
  const sheet = doc.sheet;
  const rows = [];
  for (let r = sel.r1; r <= sel.r2; r++) {
    const row = [];
    for (let c = sel.c1; c <= sel.c2; c++) {
      const cell = sheet.cells[key(r, c)];
      row.push(cell ? Object.assign({}, cell) : null);
    }
    rows.push(row);
  }
  return { range: sel, rows };
}

function blockToTsv(block) {
  return block.rows
    .map((row) =>
      row
        .map((cell) => {
          if (!cell || cell.v === undefined || cell.v === null) return '';
          const text = formatValue(cell.v, doc.styles.get(cell.s || 0).numFmt).text;
          // Quote anything that would otherwise break the TSV grid.
          return /[\t\n\r"]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
        })
        .join('\t')
    )
    .join('\n');
}

function blockToHtml(block) {
  const rows = block.rows
    .map((row) => {
      const tds = row
        .map((cell) => {
          if (!cell) return '<td></td>';
          const style = doc.styles.get(cell.s || 0);
          const css = [];
          if (style.font.bold) css.push('font-weight:bold');
          if (style.font.italic) css.push('font-style:italic');
          if (style.font.underline) css.push('text-decoration:underline');
          if (style.font.color && style.font.color !== '#000000') css.push('color:' + style.font.color);
          if (style.fill) css.push('background-color:' + style.fill);
          if (style.align.h) css.push('text-align:' + style.align.h);
          css.push('font-family:' + style.font.name);
          css.push('font-size:' + style.font.size + 'pt');
          const text = cell.v == null ? '' : formatValue(cell.v, style.numFmt).text;
          return `<td style="${css.join(';')}">${escapeHtml(text)}</td>`;
        })
        .join('');
      return '<tr>' + tds + '</tr>';
    })
    .join('');
  return `<table>${rows}</table>`;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function isTextFieldFocused() {
  const el = document.activeElement;
  return !!el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA');
}

// Copies the selected range as tab-separated text plus an HTML table, the two
// flavours Excel, Numbers and Google Sheets all understand.
function doCopy(cut) {
  if (isTextFieldFocused()) {
    const el = document.activeElement;
    const text = el.value.slice(el.selectionStart, el.selectionEnd);
    if (text) api.writeClipboard({ text });
    if (cut && text) {
      const start = el.selectionStart;
      el.value = el.value.slice(0, start) + el.value.slice(el.selectionEnd);
      el.setSelectionRange(start, start);
      el.dispatchEvent(new Event('input', { bubbles: true }));
    }
    return;
  }

  if (cut && doc.model.readOnly) return flashReadOnly();

  const block = selectedBlock();
  const tsv = blockToTsv(block);
  internalClipboard = { block, tsv, cut, sheetIndex: doc.activeSheet };
  api.writeClipboard({ text: tsv, html: blockToHtml(block) });

  // A cut clears its source only once it has been pasted, matching Excel's
  // marching-ants behaviour.
}

async function pasteFromClipboard(valuesOnly) {
  const data = await api.readClipboard();

  if (isTextFieldFocused()) {
    const el = document.activeElement;
    const start = el.selectionStart;
    const text = data.text || '';
    el.value = el.value.slice(0, start) + text + el.value.slice(el.selectionEnd);
    el.setSelectionRange(start + text.length, start + text.length);
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return;
  }

  if (doc.model.readOnly) return flashReadOnly();
  applyPaste(data.text, data.html, valuesOnly);
}

function applyPaste(text, html, valuesOnly) {
  const target = grid.active;

  // An internal copy still matching the pasteboard round-trips formulas and
  // formatting; anything else comes in as text.
  if (!valuesOnly && internalClipboard && internalClipboard.tsv === text) {
    pasteInternal(internalClipboard, target);
    return;
  }

  let rows = null;
  if (!valuesOnly && html && /<table/i.test(html)) rows = parseHtmlTable(html);
  if (!rows) rows = parseTsv(text || '');
  if (!rows.length) return;

  doc.writeBlock(target.row, target.col, rows);
  grid.setSelection(
    { r1: target.row, c1: target.col, r2: target.row + rows.length - 1, c2: target.col + Math.max(...rows.map((r) => r.length)) - 1 },
    target
  );
}

function pasteInternal(clip, target) {
  const dr = target.row - clip.block.range.r1;
  const dc = target.col - clip.block.range.c1;
  const cells = [];

  clip.block.rows.forEach((row, ri) => {
    row.forEach((cell, ci) => {
      const r = target.row + ri;
      const c = target.col + ci;
      if (!cell) {
        cells.push({ row: r, col: c, cell: null });
        return;
      }
      const next = Object.assign({}, cell);
      if (next.f) {
        next.f = shiftFormula(next.f, dr, dc);
        delete next.v;
      }
      cells.push({ row: r, col: c, cell: next });
    });
  });

  doc.transact('Paste', (tx) => tx.apply({ type: 'cells', sheetIndex: doc.activeSheet, cells }));

  if (clip.cut) {
    // A cut only clears its source once it has actually been pasted.
    doc.clearRange(clip.block.range, 'contents');
    internalClipboard = Object.assign({}, clip, { cut: false });
  }

  grid.setSelection(
    {
      r1: target.row,
      c1: target.col,
      r2: target.row + clip.block.rows.length - 1,
      c2: target.col + (clip.block.rows[0] ? clip.block.rows[0].length : 1) - 1,
    },
    target
  );
}

// Excel writes quoted fields when a cell contains a tab or newline.
function parseTsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  let i = 0;

  while (i < text.length) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && field === '') {
      inQuotes = true;
      i++;
      continue;
    }
    if (ch === '\t') {
      row.push(field);
      field = '';
      i++;
      continue;
    }
    if (ch === '\r') {
      i++;
      continue;
    }
    if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      i++;
      continue;
    }
    field += ch;
    i++;
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  // A trailing newline should not create a phantom empty row.
  while (rows.length && rows[rows.length - 1].every((f) => f === '')) rows.pop();
  return rows;
}

// Pasting from Excel, Numbers or a web page: keep the text and the handful of
// styles that survive the HTML flavour of the pasteboard.
function parseHtmlTable(html) {
  let table;
  try {
    table = new DOMParser().parseFromString(html, 'text/html').querySelector('table');
  } catch (_) {
    return null;
  }
  if (!table) return null;

  const rows = [];
  for (const tr of table.querySelectorAll('tr')) {
    const row = [];
    for (const td of tr.querySelectorAll('td, th')) {
      const text = (td.textContent || '').replace(/ /g, ' ').trim();
      const span = parseInt(td.getAttribute('colspan') || '1', 10);
      row.push(text);
      for (let i = 1; i < span; i++) row.push('');
    }
    if (row.length) rows.push(row);
  }
  return rows.length ? rows : null;
}

// ----------------------------------------------------------- formatting

function applyFormat(patch, label) {
  if (doc.model.readOnly) return flashReadOnly();
  doc.applyStyleToRange(grid.expandToMerges(grid.selection), patch, label);
}

function toggleFormat(prop) {
  const cur = doc.styleOf(grid.active.row, grid.active.col);
  const next = !cur.font[prop];
  applyFormat({ font: { [prop]: next } }, prop);
}

function adjustDecimals(delta) {
  applyFormat((style) => {
    const fmt = style.numFmt === 'General' || !style.numFmt ? '0' : style.numFmt;
    const m = /^([^.]*)(?:\.(0*))?(.*)$/.exec(fmt);
    const decimals = m && m[2] ? m[2].length : 0;
    const next = Math.max(0, Math.min(10, decimals + delta));
    const head = fmt.includes('.') ? fmt.slice(0, fmt.indexOf('.')) : fmt;
    const tailStart = fmt.includes('.') ? fmt.indexOf('.') + 1 + decimals : fmt.length;
    const tail = fmt.slice(tailStart);
    return { numFmt: head + (next ? '.' + '0'.repeat(next) : '') + tail };
  }, 'Decimals');
}

function applyBorder(kind) {
  if (doc.model.readOnly) return flashReadOnly();
  const sel = grid.expandToMerges(grid.selection);
  const si = doc.activeSheet;
  const sheet = doc.sheet;
  const items = [];
  const line = { style: 'thin', color: '#000000' };
  const thick = { style: 'thick', color: '#000000' };

  for (let r = sel.r1; r <= sel.r2; r++) {
    for (let c = sel.c1; c <= sel.c2; c++) {
      const cur = sheet.cells[key(r, c)];
      const base = cur ? cur.s || 0 : 0;
      let border;
      if (kind === 'none') {
        border = null;
      } else {
        border = Object.assign({}, doc.styles.get(base).border);
        if (kind === 'all') Object.assign(border, { top: line, right: line, bottom: line, left: line });
        else if (kind === 'outline') {
          if (r === sel.r1) border.top = line;
          if (r === sel.r2) border.bottom = line;
          if (c === sel.c1) border.left = line;
          if (c === sel.c2) border.right = line;
        } else if (kind === 'bottom' && r === sel.r2) border.bottom = line;
        else if (kind === 'thickBottom' && r === sel.r2) border.bottom = thick;
        else if (kind === 'top' && r === sel.r1) border.top = line;
        else if (kind === 'left' && c === sel.c1) border.left = line;
        else if (kind === 'right' && c === sel.c2) border.right = line;
        if (!Object.keys(border).length) border = null;
      }
      const next = doc.styles.derive(base, { border });
      if (next !== base) items.push({ row: r, col: c, s: next });
    }
  }
  if (!items.length) return;
  doc.transact('Borders', (tx) => tx.apply({ type: 'style', sheetIndex: si, cells: items }));
}

function updateToolbar() {
  const style = doc.styleOf(grid.active.row, grid.active.col);
  setToggle('bold', style.font.bold);
  setToggle('italic', style.font.italic);
  setToggle('underline', style.font.underline);
  setToggle('wrap', style.align.wrap);
  setToggle('align:left', style.align.h === 'left');
  setToggle('align:center', style.align.h === 'center');
  setToggle('align:right', style.align.h === 'right');
  setToggle('merge', isMerged());

  $('fontFamily').value = style.font.name;
  if ($('fontFamily').selectedIndex === -1) {
    const opt = document.createElement('option');
    opt.textContent = style.font.name;
    $('fontFamily').appendChild(opt);
    $('fontFamily').value = style.font.name;
  }
  $('fontSize').value = String(style.font.size);
  $('textColor').value = /^#[0-9a-f]{6}$/i.test(style.font.color) ? style.font.color : '#000000';
  if (style.fill && /^#[0-9a-f]{6}$/i.test(style.fill)) $('fillColor').value = style.fill;

  const sel = $('numFmt');
  const match = Array.from(sel.options).find((o) => o.value === style.numFmt);
  sel.value = match ? style.numFmt : 'General';
}

function setToggle(cmd, on) {
  const btn = document.querySelector(`[data-cmd="${cmd}"]`);
  if (btn) btn.classList.toggle('on', !!on);
}

// -------------------------------------------------------------- autosum

function autosum() {
  if (doc.model.readOnly) return flashReadOnly();
  const a = grid.active;
  const sheet = doc.sheet;

  // Prefer the run of numbers directly above, then the run to the left.
  let r = a.row - 1;
  while (r >= 1 && isNumeric(sheet, r, a.col)) r--;
  if (r < a.row - 1) {
    doc.setCellInput(a.row, a.col, `=SUM(${rcToA1(r + 1, a.col)}:${rcToA1(a.row - 1, a.col)})`);
    return;
  }

  let c = a.col - 1;
  while (c >= 1 && isNumeric(sheet, a.row, c)) c--;
  if (c < a.col - 1) {
    doc.setCellInput(a.row, a.col, `=SUM(${rcToA1(a.row, c + 1)}:${rcToA1(a.row, a.col - 1)})`);
    return;
  }

  editor.begin(a.row, a.col, { mode: 'edit', initial: '=SUM()' });
}

function isNumeric(sheet, r, c) {
  const cell = sheet.cells[key(r, c)];
  return !!cell && typeof cell.v === 'number' && cell.t !== 'e';
}

// ----------------------------------------------------------------- sort

function sortSelection(ascending) {
  if (doc.model.readOnly) return flashReadOnly();
  let sel = grid.expandToMerges(grid.selection);
  const sheet = doc.sheet;

  // A single cell means "sort the block of data around it".
  if (sel.r1 === sel.r2 && sel.c1 === sel.c2) {
    sel = { r1: 1, c1: 1, r2: sheet.maxRow, c2: sheet.maxCol };
  }
  if (sel.r2 <= sel.r1) return;

  const sortCol = Math.min(Math.max(grid.active.col, sel.c1), sel.c2);
  const rows = [];
  for (let r = sel.r1; r <= sel.r2; r++) {
    const cells = [];
    for (let c = sel.c1; c <= sel.c2; c++) {
      const cell = sheet.cells[key(r, c)];
      cells.push(cell ? Object.assign({}, cell) : null);
    }
    rows.push({ cells, sortValue: sheet.cells[key(r, sortCol)] });
  }

  rows.sort((a, b) => compareCells(a.sortValue, b.sortValue) * (ascending ? 1 : -1));

  const cells = [];
  rows.forEach((row, i) => {
    const r = sel.r1 + i;
    row.cells.forEach((cell, j) => {
      // Formulas are not re-pointed when rows move, so a sorted formula would
      // silently reference the wrong row; store its computed value instead.
      let next = cell;
      if (cell && cell.f) {
        next = { s: cell.s };
        if (cell.v !== undefined) {
          next.v = cell.v;
          if (cell.t) next.t = cell.t;
        }
      }
      cells.push({ row: r, col: sel.c1 + j, cell: next });
    });
  });

  doc.transact('Sort', (tx) => tx.apply({ type: 'cells', sheetIndex: doc.activeSheet, cells }));
}

// Excel's sort order: numbers, then text, then booleans, then errors, then blanks.
function compareCells(a, b) {
  const ra = sortRank(a);
  const rb = sortRank(b);
  if (ra !== rb) return ra - rb;
  if (ra === 4) return 0;
  if (ra === 0) return a.v - b.v;
  if (ra === 1) return String(a.v).localeCompare(String(b.v), undefined, { numeric: true, sensitivity: 'base' });
  if (ra === 2) return (a.v ? 1 : 0) - (b.v ? 1 : 0);
  return String(a.v).localeCompare(String(b.v));
}

function sortRank(cell) {
  if (!cell || cell.v === undefined || cell.v === null || cell.v === '') return 4;
  if (cell.t === 'e') return 3;
  if (typeof cell.v === 'boolean') return 2;
  if (typeof cell.v === 'number') return 0;
  return 1;
}

// ----------------------------------------------------------------- find

let findMatches = [];
let findIndex = -1;

function openFind() {
  $('findBar').hidden = false;
  $('findInput').focus();
  $('findInput').select();
  runFind();
}

function closeFind() {
  $('findBar').hidden = true;
  findMatches = [];
  grid.focus();
}

function runFind() {
  const q = $('findInput').value.trim().toLowerCase();
  findMatches = [];
  findIndex = -1;
  if (q) {
    const sheet = doc.sheet;
    for (let r = 1; r <= sheet.maxRow; r++) {
      for (let c = 1; c <= sheet.maxCol; c++) {
        const cell = sheet.cells[key(r, c)];
        if (!cell || cell.v === undefined || cell.v === null) continue;
        const text = formatValue(cell.v, doc.styles.get(cell.s || 0).numFmt).text;
        const formula = cell.f ? '=' + cell.f : '';
        if (text.toLowerCase().includes(q) || formula.toLowerCase().includes(q)) {
          findMatches.push({ row: r, col: c });
        }
      }
    }
  }
  $('findCount').textContent = findMatches.length ? `0/${findMatches.length}` : q ? 'None' : '';
  if (findMatches.length) stepFind(1);
}

function stepFind(dir) {
  if (!findMatches.length) return;
  findIndex = (findIndex + dir + findMatches.length) % findMatches.length;
  const m = findMatches[findIndex];
  grid.anchor = { row: m.row, col: m.col };
  grid.selectCell(m.row, m.col);
  $('findCount').textContent = `${findIndex + 1}/${findMatches.length}`;
}

$('findInput').addEventListener('input', runFind);
$('findInput').addEventListener('keydown', (e) => {
  e.stopPropagation();
  if (e.key === 'Enter') stepFind(e.shiftKey ? -1 : 1);
  else if (e.key === 'Escape') closeFind();
});
$('findNext').onclick = () => stepFind(1);
$('findPrev').onclick = () => stepFind(-1);
$('findClose').onclick = closeFind;

// -------------------------------------------------------------- name box

$('nameBox').addEventListener('keydown', (e) => {
  e.stopPropagation();
  if (e.key !== 'Enter') return;
  const ref = $('nameBox').value.trim();
  const range = parseRange(ref);
  if (range) {
    grid.anchor = { row: range.r1, col: range.c1 };
    grid.setSelection(range, { row: range.r1, col: range.c1 });
    grid.focus();
  } else {
    updateNameBox();
  }
});
$('nameBox').addEventListener('blur', updateNameBox);

// -------------------------------------------------------------- toolbar

document.querySelectorAll('[data-cmd]').forEach((btn) => {
  btn.addEventListener('click', () => runCommand(btn.dataset.cmd));
});

$('fontFamily').onchange = (e) => applyFormat({ font: { name: e.target.value } }, 'Font');
$('fontSize').onchange = (e) => applyFormat({ font: { size: parseFloat(e.target.value) } }, 'Font Size');
$('textColor').oninput = (e) => applyFormat({ font: { color: e.target.value.toUpperCase() } }, 'Text Colour');
$('fillColor').oninput = (e) => applyFormat({ fill: e.target.value.toUpperCase() }, 'Fill');
$('numFmt').onchange = (e) => applyFormat({ numFmt: e.target.value }, 'Number Format');

$('borderBtn').onclick = (e) => {
  e.stopPropagation();
  const menu = $('borderMenu');
  menu.hidden = !menu.hidden;
};
$('borderMenu').querySelectorAll('[data-border]').forEach((btn) => {
  btn.onclick = () => {
    $('borderMenu').hidden = true;
    applyBorder(btn.dataset.border);
  };
});

function runCommand(cmd, arg) {
  if (cmd.startsWith('align:')) {
    const which = cmd.slice(6);
    const cur = doc.styleOf(grid.active.row, grid.active.col);
    return applyFormat({ align: { h: cur.align.h === which ? null : which } }, 'Align');
  }

  switch (cmd) {
    case 'bold':
    case 'italic':
    case 'underline':
      return toggleFormat(cmd);
    case 'wrap': {
      const cur = doc.styleOf(grid.active.row, grid.active.col);
      return applyFormat({ align: { wrap: !cur.align.wrap } }, 'Wrap');
    }
    case 'merge':
      return toggleMerge();
    case 'incDecimal':
      return adjustDecimals(1);
    case 'decDecimal':
      return adjustDecimals(-1);
    case 'autosum':
      return autosum();
    case 'sortAsc':
      return sortSelection(true);
    case 'sortDesc':
      return sortSelection(false);
    case 'align':
      return runCommand('align:' + arg);
    case 'numFmt':
      return applyFormat({ numFmt: arg }, 'Number Format');
    case 'undo':
      if (isTextFieldFocused()) return document.execCommand('undo');
      doc.undo();
      return refreshAll();
    case 'redo':
      if (isTextFieldFocused()) return document.execCommand('redo');
      doc.redo();
      return refreshAll();
    case 'selectAll':
      if (isTextFieldFocused()) return document.execCommand('selectAll');
      return grid.setSelection({ r1: 1, c1: 1, r2: grid.nRows, c2: grid.nCols }, grid.active, { scroll: false });
    case 'clear':
      if (doc.model.readOnly) return flashReadOnly();
      return doc.clearRange(grid.selection, 'contents');
    case 'copy':
      return doCopy(false);
    case 'cut':
      return doCopy(true);
    case 'paste':
      return pasteFromClipboard(false);
    case 'pasteValues':
      return pasteFromClipboard(true);
    case 'insertRow':
      return insertRows();
    case 'insertCol':
      return insertCols();
    case 'deleteRow':
      return deleteRows();
    case 'deleteCol':
      return deleteCols();
    case 'newSheet':
      return $('addSheet').click();
    case 'freeze':
      return freezeAtSelection();
    case 'unfreeze':
      return unfreeze();
    case 'autofit':
      return autofitSelection();
    case 'find':
      return openFind();
    case 'goto':
      $('nameBox').focus();
      return $('nameBox').select();
    case 'open':
      return openFile();
    case 'save':
      return saveFile(false);
    case 'saveAs':
      return saveFile(true);
    case 'exportCsv':
      return exportCsv();
    case 'print':
      return printSheet();
    case 'closeWindow':
      return requestClose();
    case 'shortcuts':
      return showShortcuts();
    case 'compat':
      return showCompat();
    case 'sample':
      return loadSample();
    default:
      return undefined;
  }
}

// ------------------------------------------------------------- file I/O

async function openFile(path) {
  if (!(await confirmDiscard())) return;
  const res = await api.openWorkbook(path);
  if (res.canceled) return;
  if (res.error) {
    await api.messageBox({ type: 'error', message: 'Could not open the file', detail: res.error, buttons: ['OK'] });
    return;
  }
  doc.load(res.model);
  grid.viewBySheet.clear();
  grid.selection = { r1: 1, c1: 1, r2: 1, c2: 1 };
  grid.active = { row: 1, col: 1 };
  grid.scroller.scrollLeft = 0;
  grid.scroller.scrollTop = 0;
  refreshAll();
  grid.focus();
}

async function saveFile(saveAs) {
  if (editor.open) editor.commit(null);
  const res = await api.saveWorkbook(serializeModel(), saveAs);
  if (res.canceled) return false;
  if (res.error) {
    await api.messageBox({ type: 'error', message: 'Could not save the file', detail: res.error, buttons: ['OK'] });
    return false;
  }
  doc.model.filePath = res.filePath;
  doc.model.fileType = res.filePath.split('.').pop().toLowerCase();
  // Saving a copy of a read-only source makes the copy editable.
  doc.model.readOnly = false;
  doc.model.readOnlyReason = null;
  doc.markClean();
  refreshAll();
  return true;
}

// The model crosses IPC by structured clone, so hand over a plain snapshot.
function serializeModel() {
  return {
    filePath: doc.model.filePath,
    fileType: doc.model.fileType,
    readOnly: doc.model.readOnly,
    activeSheet: doc.activeSheet,
    sheets: doc.model.sheets,
    styles: doc.styles.toJSON(),
    definedNames: doc.model.definedNames,
  };
}

async function exportCsv() {
  const path = await api.chooseCsvPath();
  if (!path) return;
  const model = serializeModel();
  model.filePath = path;
  const res = await api.saveWorkbook(Object.assign({}, model, { filePath: path, readOnly: false }), false);
  if (res.error) {
    await api.messageBox({ type: 'error', message: 'Could not export the CSV', detail: res.error, buttons: ['OK'] });
  }
}

async function confirmDiscard() {
  if (!doc.dirty) return true;
  const name = doc.model.filePath ? doc.model.filePath.split('/').pop() : 'Untitled';
  const answer = await api.confirmDiscard(name);
  if (answer === 'cancel') return false;
  if (answer === 'save') return await saveFile(false);
  return true;
}

async function requestClose() {
  if (await confirmDiscard()) api.closeWindow();
}

// --------------------------------------------------------------- printing

async function printSheet() {
  const sheet = doc.sheet;
  const maxR = Math.min(sheet.maxRow, 5000);
  const maxC = Math.min(sheet.maxCol, 200);
  const rows = [];

  for (let r = 1; r <= maxR; r++) {
    const tds = [];
    for (let c = 1; c <= maxC; c++) {
      const cell = sheet.cells[key(r, c)];
      const style = doc.styles.get(cell ? cell.s || 0 : 0);
      const css = [`font-family:${style.font.name}`, `font-size:${style.font.size}pt`];
      if (style.font.bold) css.push('font-weight:bold');
      if (style.font.italic) css.push('font-style:italic');
      if (style.font.color !== '#000000') css.push('color:' + style.font.color);
      if (style.fill) css.push('background-color:' + style.fill);
      const isNum = cell && typeof cell.v === 'number';
      css.push('text-align:' + (style.align.h || (isNum ? 'right' : 'left')));
      const text = cell && cell.v != null ? formatValue(cell.v, style.numFmt).text : '';
      tds.push(`<td style="${css.join(';')}">${escapeHtml(text)}</td>`);
    }
    rows.push('<tr>' + tds.join('') + '</tr>');
  }

  const widths = [];
  for (let c = 1; c <= maxC; c++) {
    const spec = sheet.cols[c];
    widths.push(`<col style="width:${widthToPx(spec && spec.width != null ? spec.width : sheet.defaultColWidth)}px">`);
  }

  const html = `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(sheet.name)}</title>
<style>
  @page { margin: 12mm; }
  body { font: 11pt -apple-system, sans-serif; margin: 0; }
  h1 { font-size: 13pt; margin: 0 0 8px; }
  table { border-collapse: collapse; table-layout: fixed; }
  td { border: 0.5pt solid #c9c9c9; padding: 1px 4px; overflow: hidden;
       white-space: nowrap; text-overflow: ellipsis; }
</style></head><body>
<h1>${escapeHtml(sheet.name)}</h1>
<table><colgroup>${widths.join('')}</colgroup>${rows.join('')}</table>
</body></html>`;

  const res = await api.print(html);
  if (res && !res.ok && res.error) {
    await api.messageBox({ type: 'error', message: 'Could not print', detail: res.error, buttons: ['OK'] });
  }
}

// ------------------------------------------------------------------ help

function showShortcuts() {
  api.messageBox({
    message: 'Keyboard shortcuts',
    detail: [
      'Cmd+N / Cmd+O / Cmd+S      New / Open / Save',
      'Cmd+Z / Shift+Cmd+Z        Undo / Redo',
      'Cmd+C / Cmd+X / Cmd+V      Copy / Cut / Paste',
      'Shift+Cmd+V                Paste values only',
      'Cmd+B / Cmd+I / Cmd+U      Bold / Italic / Underline',
      'Cmd+F                      Find',
      'Ctrl+G                     Go to cell',
      'F2                         Edit the active cell',
      'Alt+Enter                  Line break inside a cell',
      'Cmd+Arrow                  Jump to the edge of the data',
      'Shift+Arrow                Extend the selection',
      'Delete                     Clear contents',
      'Drag the square handle     Fill down or across',
      'Double-click a column edge Autofit that column',
    ].join('\n'),
    buttons: ['OK'],
  });
}

function showCompat() {
  api.messageBox({
    message: 'Excel compatibility',
    detail: [
      'Reads and writes .xlsx: values, formulas, number formats, fonts, fills,',
      'borders, alignment, merged cells, column widths, row heights, frozen',
      'panes, multiple sheets and defined names.',
      '',
      '.xls and macro-enabled .xlsm open read-only — use Save As for a .xlsx copy.',
      '',
      'Not supported: charts, pivot tables, macros, conditional formatting,',
      'data validation and images. Those parts of a file are left behind when',
      'a workbook is saved, which is why the originals are never overwritten',
      'in place unless you chose the destination yourself.',
    ].join('\n'),
    buttons: ['OK'],
  });
}

function loadSample() {
  const rows = [
    ['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Total', 'Share'],
    ['North', 128400, 141200, 137800, 159300, '=SUM(B2:E2)', '=F2/$F$6'],
    ['South', 98200, 102750, 111400, 118900, '=SUM(B3:E3)', '=F3/$F$6'],
    ['East', 174600, 168300, 181200, 195400, '=SUM(B4:E4)', '=F4/$F$6'],
    ['West', 87300, 94100, 89600, 101250, '=SUM(B5:E5)', '=F5/$F$6'],
    ['Total', '=SUM(B2:B5)', '=SUM(C2:C5)', '=SUM(D2:D5)', '=SUM(E2:E5)', '=SUM(F2:F5)', '=SUM(G2:G5)'],
  ];

  doc.loadEmpty();
  doc.writeBlock(1, 1, rows);

  const header = doc.styles.intern({
    font: { name: 'Calibri', size: 11, bold: true, color: '#FFFFFF' },
    fill: '#0D6B3F',
    align: { h: 'center', v: 'middle' },
  });
  const money = doc.styles.intern({ numFmt: '#,##0' });
  const pct = doc.styles.intern({ numFmt: '0.0%' });
  const totalRow = doc.styles.intern({
    font: { bold: true },
    numFmt: '#,##0',
    border: { top: { style: 'thin', color: '#000000' } },
  });

  doc.transact('Sample', (tx) => {
    const items = [];
    for (let c = 1; c <= 7; c++) items.push({ row: 1, col: c, s: header });
    for (let r = 2; r <= 5; r++) {
      for (let c = 2; c <= 6; c++) items.push({ row: r, col: c, s: money });
      items.push({ row: r, col: 7, s: pct });
    }
    for (let c = 1; c <= 7; c++) items.push({ row: 6, col: c, s: totalRow });
    tx.apply({ type: 'style', sheetIndex: 0, cells: items });
    tx.apply({
      type: 'colWidth',
      sheetIndex: 0,
      cols: [{ col: 1, spec: { width: 14 } }, { col: 6, spec: { width: 12 } }, { col: 7, spec: { width: 10 } }],
    });
    tx.apply({ type: 'frozen', sheetIndex: 0, frozen: { row: 1, col: 1 } });
  });

  doc.model.filePath = null;
  doc.markClean();
  refreshAll();
  grid.focus();
}

// ------------------------------------------------------------ main wiring

api.onMenuCommand(({ cmd, arg }) => runCommand(cmd, arg));
api.onOpenPath((p) => openFile(p));
api.onOpenPathIfEmpty((p) => {
  // An untouched blank window adopts the file; anything else hands it to a new
  // window, which has to carry the path or the file is silently dropped.
  if (!doc.dirty && !doc.model.filePath) openFile(p);
  else api.newWindow(p);
});
api.onRequestClose(() => requestClose());

// The canvas paints itself rather than inheriting CSS, so a theme change has
// to reach the grid explicitly and force a repaint.
api.onTheme((theme) => {
  document.documentElement.dataset.theme = theme;
  setGridTheme(theme);
  grid.refresh();
});

window.addEventListener('resize', () => grid.refresh());

// Follow hyperlinks out to the default browser instead of navigating the app.
document.addEventListener('click', (e) => {
  const a = e.target.closest && e.target.closest('a[href]');
  if (a) {
    e.preventDefault();
    api.openExternal(a.getAttribute('href'));
  }
});

refreshAll();
grid.focus();
