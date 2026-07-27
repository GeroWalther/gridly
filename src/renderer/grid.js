import { key, colToLetter, parseRange } from '../shared/addr.js';
import { formatValue, widthToPx, heightToPx, pxToWidth, pxToHeight, ptToPx } from './format.js';

// A canvas grid. The canvas is purely a painting surface; a transparent
// overflow:auto layer sits on top of it to provide real macOS scrollbars and
// trackpad momentum, and all pointer input is read from that layer.

const HEADER_H = 24;
const MIN_HEADER_W = 44;
const CELL_PAD = 4;
const RESIZE_GRAB = 5;
const FILL_HANDLE = 6;

const THEMES = {
  light: {
    dark: false,
    gridline: '#d9d9d9',
    headerBg: '#f5f5f5',
    headerBgActive: '#dfe6e0',
    headerText: '#4a4a4a',
    headerTextActive: '#0d6b3f',
    headerLine: '#c4c4c4',
    selBorder: '#0d6b3f',
    selFill: 'rgba(13, 107, 63, 0.09)',
    frozenLine: '#8a8a8a',
    canvasBg: '#ffffff',
    outside: '#fafafa',
    ink: '#000000',
    docBorder: '#000000',
  },
  dark: {
    dark: true,
    gridline: '#3a3a40',
    headerBg: '#242427',
    headerBgActive: '#2c3f35',
    headerText: '#a0a0a8',
    headerTextActive: '#4ecf95',
    headerLine: '#3f3f46',
    selBorder: '#4ecf95',
    selFill: 'rgba(78, 207, 149, 0.16)',
    frozenLine: '#6e6e78',
    canvasBg: '#1b1b1e',
    outside: '#131315',
    ink: '#e6e6ea',
    docBorder: '#8b8b96',
  },
};

let THEME = THEMES.light;

export function setGridTheme(name) {
  THEME = THEMES[name] || THEMES.light;
}

const NEAR_BLACK = 70;

function isNearBlack(color) {
  const m = /^#([0-9a-f]{6})$/i.exec(String(color));
  if (!m) return false;
  const n = parseInt(m[1], 16);
  return 0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255) < NEAR_BLACK;
}

// Workbook colours are authored against a white sheet. An unfilled cell takes
// the theme background, so on a dark canvas the near-black ink chosen for that
// white sheet would all but vanish — it maps to the theme's foreground instead.
// Ink the author gave a real colour, and anything sitting on an explicit fill,
// is left exactly as the file specifies.
function inkOn(fill, color, substitute) {
  if (!THEME.dark || fill) return color || substitute;
  if (color && !isNearBlack(color)) return color;
  return substitute;
}

const MAX_ROWS = 1048576;
const MAX_COLS = 16384;

export class Grid {
  constructor(container, doc) {
    this.container = container;
    this.doc = doc;

    this.canvas = document.createElement('canvas');
    this.canvas.className = 'grid-canvas';
    container.appendChild(this.canvas);

    this.scroller = document.createElement('div');
    this.scroller.className = 'grid-scroller';
    this.scroller.tabIndex = 0;
    this.spacer = document.createElement('div');
    this.spacer.className = 'grid-spacer';
    this.scroller.appendChild(this.spacer);
    container.appendChild(this.scroller);

    this.ctx = this.canvas.getContext('2d', { alpha: false });

    this.selection = { r1: 1, c1: 1, r2: 1, c2: 1 };
    this.active = { row: 1, col: 1 };
    this.viewBySheet = new Map();

    this.measureCache = new Map();
    this.drag = null;
    this.hover = null;
    this.rafPending = false;

    this.onSelectionChange = null;
    this.onEditRequest = null;
    this.onKeyDown = null;
    this.onContextMenu = null;
    this.onStructuralChange = null;
    this.onFill = null;

    this.bindEvents();
    this.resizeObserver = new ResizeObserver(() => this.refresh());
    this.resizeObserver.observe(container);
  }

  get sheet() {
    return this.doc.sheet;
  }

  focus() {
    this.scroller.focus({ preventScroll: true });
  }

  // ------------------------------------------------------------- layout

  refresh() {
    this.layout();
    this.render();
  }

  layout() {
    const sheet = this.sheet;
    if (!sheet) return;

    const rect = this.container.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    const cssW = Math.max(1, Math.floor(rect.width));
    const cssH = Math.max(1, Math.floor(rect.height));

    if (this.canvas.width !== Math.floor(cssW * dpr) || this.canvas.height !== Math.floor(cssH * dpr)) {
      this.canvas.width = Math.floor(cssW * dpr);
      this.canvas.height = Math.floor(cssH * dpr);
      this.canvas.style.width = cssW + 'px';
      this.canvas.style.height = cssH + 'px';
    }
    this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    this.viewW = cssW;
    this.viewH = cssH;

    // Extend past the used range so there is always empty grid to scroll into.
    this.nCols = Math.min(MAX_COLS, Math.max(sheet.maxCol + 20, 40));
    this.nRows = Math.min(MAX_ROWS, Math.max(sheet.maxRow + 80, 200));

    const defColPx = widthToPx(sheet.defaultColWidth || 8.43);
    const defRowPx = heightToPx(sheet.defaultRowHeight || 15);

    // colPos[c] is the left edge of column c in content pixels, so
    // colPos[c+1] - colPos[c] is the width of column c.
    this.colPos = new Float64Array(this.nCols + 2);
    for (let c = 2; c <= this.nCols + 1; c++) {
      this.colPos[c] = this.colPos[c - 1] + colWidthPx(sheet, c - 1, defColPx);
    }

    this.rowPos = new Float64Array(this.nRows + 2);
    for (let r = 2; r <= this.nRows + 1; r++) {
      this.rowPos[r] = this.rowPos[r - 1] + rowHeightPx(sheet, r - 1, defRowPx);
    }

    this.defColPx = defColPx;
    this.defRowPx = defRowPx;

    // Row header widens with the number of digits on screen.
    this.ctx.font = '12px -apple-system, sans-serif';
    const digits = String(Math.max(this.nRows, 100)).length;
    this.headerW = Math.max(MIN_HEADER_W, digits * 8 + 16);

    this.frozenRows = Math.min(sheet.frozen ? sheet.frozen.row || 0 : 0, this.nRows);
    this.frozenCols = Math.min(sheet.frozen ? sheet.frozen.col || 0 : 0, this.nCols);
    this.frozenW = this.colPos[this.frozenCols + 1];
    this.frozenH = this.rowPos[this.frozenRows + 1];

    this.totalW = this.colPos[this.nCols + 1];
    this.totalH = this.rowPos[this.nRows + 1];

    this.paneW = Math.max(0, this.viewW - this.headerW - this.frozenW);
    this.paneH = Math.max(0, this.viewH - HEADER_H - this.frozenH);

    // The transparent scroll layer covers the whole grid — including the
    // headers — so that every pointer event lands on one element. Its spacer
    // is sized so that scrolling bottoms out exactly when the last row and
    // column are flush against the frozen bands: scrollMax has to come out to
    // (totalW - frozenW) - paneW, and the layer is viewW wide, so the spacer
    // needs the header and frozen widths added back on.
    this.spacer.style.width = Math.max(0, this.totalW + this.headerW) + 'px';
    this.spacer.style.height = Math.max(0, this.totalH + HEADER_H) + 'px';

    this.buildMergeIndex(sheet);
  }

  buildMergeIndex(sheet) {
    this.mergeAnchor = new Map();
    this.mergeCovered = new Map();
    for (const ref of sheet.merges || []) {
      const r = parseRange(ref);
      if (!r) continue;
      const anchorKey = key(r.r1, r.c1);
      this.mergeAnchor.set(anchorKey, r);
      for (let rr = r.r1; rr <= r.r2; rr++) {
        for (let cc = r.c1; cc <= r.c2; cc++) {
          if (rr === r.r1 && cc === r.c1) continue;
          this.mergeCovered.set(key(rr, cc), anchorKey);
        }
      }
    }
  }

  // Expands a range so that it fully contains any merges it touches.
  expandToMerges(range) {
    let changed = true;
    const out = Object.assign({}, range);
    let guard = 0;
    while (changed && guard++ < 8) {
      changed = false;
      for (const r of this.mergeAnchor.values()) {
        const overlaps = !(r.r2 < out.r1 || r.r1 > out.r2 || r.c2 < out.c1 || r.c1 > out.c2);
        if (!overlaps) continue;
        if (r.r1 < out.r1) (out.r1 = r.r1), (changed = true);
        if (r.c1 < out.c1) (out.c1 = r.c1), (changed = true);
        if (r.r2 > out.r2) (out.r2 = r.r2), (changed = true);
        if (r.c2 > out.c2) (out.c2 = r.c2), (changed = true);
      }
    }
    return out;
  }

  get scrollX() {
    return this.scroller.scrollLeft;
  }

  get scrollY() {
    return this.scroller.scrollTop;
  }

  colX(c) {
    return this.headerW + this.colPos[c] - (c <= this.frozenCols ? 0 : this.scrollX);
  }

  rowY(r) {
    return HEADER_H + this.rowPos[r] - (r <= this.frozenRows ? 0 : this.scrollY);
  }

  colW(c) {
    return this.colPos[c + 1] - this.colPos[c];
  }

  rowH(r) {
    return this.rowPos[r + 1] - this.rowPos[r];
  }

  // First and last column drawn in the scrolling pane.
  visibleCols() {
    const startContent = this.scrollX + this.colPos[this.frozenCols + 1];
    let start = binarySearch(this.colPos, startContent, this.nCols);
    start = Math.max(this.frozenCols + 1, start);
    let end = start;
    const limit = startContent + this.paneW;
    while (end < this.nCols && this.colPos[end] < limit) end++;
    return { start, end: Math.min(this.nCols, end) };
  }

  visibleRows() {
    const startContent = this.scrollY + this.rowPos[this.frozenRows + 1];
    let start = binarySearch(this.rowPos, startContent, this.nRows);
    start = Math.max(this.frozenRows + 1, start);
    let end = start;
    const limit = startContent + this.paneH;
    while (end < this.nRows && this.rowPos[end] < limit) end++;
    return { start, end: Math.min(this.nRows, end) };
  }

  // ------------------------------------------------------------ rendering

  requestRender() {
    if (this.rafPending) return;
    this.rafPending = true;
    requestAnimationFrame(() => {
      this.rafPending = false;
      this.render();
    });
  }

  render() {
    const ctx = this.ctx;
    const sheet = this.sheet;
    if (!sheet || !this.colPos) return;

    ctx.fillStyle = THEME.outside;
    ctx.fillRect(0, 0, this.viewW, this.viewH);
    ctx.fillStyle = THEME.canvasBg;
    ctx.fillRect(this.headerW, HEADER_H, this.viewW - this.headerW, this.viewH - HEADER_H);

    const vc = this.visibleCols();
    const vr = this.visibleRows();

    const colBands = [];
    if (this.frozenCols > 0) colBands.push({ start: 1, end: this.frozenCols, frozen: true });
    colBands.push({ start: vc.start, end: vc.end, frozen: false });

    const rowBands = [];
    if (this.frozenRows > 0) rowBands.push({ start: 1, end: this.frozenRows, frozen: true });
    rowBands.push({ start: vr.start, end: vr.end, frozen: false });

    // Four panes: frozen corner, frozen row band, frozen column band, and the
    // scrolling body. Each is clipped so scrolled content cannot bleed into a
    // frozen band.
    for (const rb of rowBands) {
      for (const cb of colBands) {
        const clip = {
          x: cb.frozen ? this.headerW : this.headerW + this.frozenW,
          y: rb.frozen ? HEADER_H : HEADER_H + this.frozenH,
          w: cb.frozen ? this.frozenW : this.paneW,
          h: rb.frozen ? this.frozenH : this.paneH,
        };
        if (clip.w <= 0 || clip.h <= 0) continue;
        ctx.save();
        ctx.beginPath();
        ctx.rect(clip.x, clip.y, clip.w, clip.h);
        ctx.clip();
        this.drawPane(ctx, rb, cb);
        ctx.restore();
      }
    }

    this.drawSelection(ctx, rowBands, colBands);
    this.drawHeaders(ctx, rowBands, colBands);
    this.drawFrozenDividers(ctx);
  }

  drawPane(ctx, rb, cb) {
    const sheet = this.sheet;

    // Backgrounds first, so a fill never paints over a neighbour's text.
    for (let r = rb.start; r <= rb.end; r++) {
      if (isHiddenRow(sheet, r)) continue;
      const y = this.rowY(r);
      const h = this.rowH(r);
      if (h <= 0) continue;
      for (let c = cb.start; c <= cb.end; c++) {
        if (isHiddenCol(sheet, c)) continue;
        const k = key(r, c);
        if (this.mergeCovered.has(k)) continue;
        const cell = sheet.cells[k];
        if (!cell) continue;
        const style = this.doc.styles.get(cell.s || 0);
        if (!style.fill) continue;
        const merge = this.mergeAnchor.get(k);
        const x = this.colX(c);
        const w = merge ? this.colPos[merge.c2 + 1] - this.colPos[merge.c1] : this.colW(c);
        const hh = merge ? this.rowPos[merge.r2 + 1] - this.rowPos[merge.r1] : h;
        ctx.fillStyle = style.fill;
        ctx.fillRect(x, y, w, hh);
      }
    }

    this.drawGridlines(ctx, rb, cb);
    this.clearMergeInteriors(ctx, rb, cb);

    // Text.
    for (let r = rb.start; r <= rb.end; r++) {
      if (isHiddenRow(sheet, r)) continue;
      for (let c = cb.start; c <= cb.end; c++) {
        if (isHiddenCol(sheet, c)) continue;
        const k = key(r, c);
        if (this.mergeCovered.has(k)) continue;
        this.drawCellContent(ctx, r, c, sheet, cb);
      }
    }

    this.drawMergeOverflow(ctx, rb, cb);
    this.drawBorders(ctx, rb, cb);
  }

  // Whether this pane is responsible for painting the merge anchored at m:
  // either the anchor sits in the pane's own band, or the merge spills in from
  // a frozen band that is still flush against this one. Once that axis has
  // scrolled the panes no longer line up and the spill would land on top of
  // scrolled content, so it is dropped.
  paneShowsMerge(m, rb, cb) {
    if (m.r1 > rb.end || m.r2 < rb.start || m.c1 > cb.end || m.c2 < cb.start) return false;

    const anchorInBand = m.r1 >= rb.start && m.r1 <= rb.end && m.c1 >= cb.start && m.c1 <= cb.end;
    if (anchorInBand) return true;

    if ((m.c1 <= this.frozenCols) !== cb.frozen && this.scrollX !== 0) return false;
    if ((m.r1 <= this.frozenRows) !== rb.frozen && this.scrollY !== 0) return false;
    return true;
  }

  // A merge anchored inside a frozen band can reach past the split. The pane
  // that owns the anchor clips at the divider, so the pane on the other side
  // has to paint the same cell again to carry the text across.
  drawMergeOverflow(ctx, rb, cb) {
    if (!this.mergeAnchor.size) return;
    const sheet = this.sheet;

    for (const m of this.mergeAnchor.values()) {
      const anchorInBand = m.r1 >= rb.start && m.r1 <= rb.end && m.c1 >= cb.start && m.c1 <= cb.end;
      if (anchorInBand) continue;
      if (!this.paneShowsMerge(m, rb, cb)) continue;
      this.drawCellContent(ctx, m.r1, m.c1, sheet, cb);
    }
  }

  // Gridlines are painted straight across the band, so the interior lines of a
  // merged cell have to be wiped back out. The rect is inset by a pixel so the
  // merge's own outline stays part of the surrounding grid.
  clearMergeInteriors(ctx, rb, cb) {
    if (!this.mergeAnchor.size) return;
    const sheet = this.sheet;

    for (const m of this.mergeAnchor.values()) {
      if (m.r1 === m.r2 && m.c1 === m.c2) continue;
      if (!this.paneShowsMerge(m, rb, cb)) continue;

      const w = this.colPos[m.c2 + 1] - this.colPos[m.c1];
      const h = this.rowPos[m.r2 + 1] - this.rowPos[m.r1];
      if (w <= 2 || h <= 2) continue;

      const cell = sheet.cells[key(m.r1, m.c1)];
      const style = this.doc.styles.get(cell ? cell.s || 0 : 0);
      ctx.fillStyle = style.fill || THEME.canvasBg;
      ctx.fillRect(this.colX(m.c1) + 1, this.rowY(m.r1) + 1, w - 2, h - 2);
    }
  }

  drawGridlines(ctx, rb, cb) {
    const sheet = this.sheet;
    ctx.strokeStyle = THEME.gridline;
    ctx.lineWidth = 1;
    ctx.beginPath();

    const xStart = this.colX(cb.start);
    const xEnd = this.colX(cb.end) + this.colW(cb.end);
    const yStart = this.rowY(rb.start);
    const yEnd = this.rowY(rb.end) + this.rowH(rb.end);

    for (let r = rb.start; r <= rb.end + 1; r++) {
      if (r > this.nRows + 1) break;
      const y = Math.round(this.rowY(r)) + 0.5;
      ctx.moveTo(xStart, y);
      ctx.lineTo(xEnd, y);
    }
    for (let c = cb.start; c <= cb.end + 1; c++) {
      if (c > this.nCols + 1) break;
      const x = Math.round(this.colX(c)) + 0.5;
      ctx.moveTo(x, yStart);
      ctx.lineTo(x, yEnd);
    }
    ctx.stroke();
  }

  drawCellContent(ctx, r, c, sheet, cb) {
    const k = key(r, c);
    const cell = sheet.cells[k];
    if (!cell || cell.v === undefined || cell.v === null) return;

    const style = this.doc.styles.get(cell.s || 0);
    const merge = this.mergeAnchor.get(k);

    const x = this.colX(c);
    const y = this.rowY(r);
    const w = merge ? this.colPos[merge.c2 + 1] - this.colPos[merge.c1] : this.colW(c);
    const h = merge ? this.rowPos[merge.r2 + 1] - this.rowPos[merge.r1] : this.rowH(r);
    if (w <= 1 || h <= 1) return;

    const { text, color } = formatValue(cell.v, style.numFmt);
    if (text === '') return;

    const font = fontString(style.font);
    ctx.font = font;
    ctx.fillStyle = inkOn(style.fill, color || style.font.color, THEME.ink);

    const isNumeric = typeof cell.v === 'number';
    const isError = cell.t === 'e';
    let align = style.align.h;
    if (!align || align === 'general') {
      align = isError ? 'center' : isNumeric ? 'right' : typeof cell.v === 'boolean' ? 'center' : 'left';
    }

    const availW = w - CELL_PAD * 2;

    if (style.align.wrap && !merge) {
      this.drawWrapped(ctx, text, x, y, w, h, align, style);
      return;
    }

    const textW = this.measure(ctx, text, font);

    // Numbers that do not fit show ### rather than being clipped, as Excel does.
    if (isNumeric && !isError && textW > availW && style.numFmt !== 'General') {
      const hashes = '#'.repeat(Math.max(1, Math.floor(availW / this.measure(ctx, '#', font))));
      this.drawLine(ctx, hashes, x, y, w, h, align, style);
      return;
    }

    // Text wider than its column spills into adjacent empty cells, again
    // following Excel; the spill stops at the first occupied cell.
    let clipX = x;
    let clipW = w;
    if (!merge && textW > availW) {
      if (align === 'left') {
        let cc = c + 1;
        while (cc <= this.nCols && clipW - CELL_PAD * 2 < textW && !isOccupied(sheet, r, cc) && cc <= cb.end + 4) {
          clipW += this.colW(cc);
          cc++;
        }
      } else if (align === 'right') {
        let cc = c - 1;
        while (cc >= 1 && clipW - CELL_PAD * 2 < textW && !isOccupied(sheet, r, cc)) {
          clipW += this.colW(cc);
          clipX -= this.colW(cc);
          cc--;
        }
      }
    }

    const needsClip = textW > clipW - CELL_PAD * 2;
    if (needsClip) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, y, clipW, h);
      ctx.clip();
    }
    this.drawLine(ctx, text, clipX, y, clipW, h, align, style);
    if (needsClip) ctx.restore();
  }

  drawLine(ctx, text, x, y, w, h, align, style) {
    const indent = (style.align.indent || 0) * 8;
    let tx;
    if (align === 'right') tx = x + w - CELL_PAD - indent;
    else if (align === 'center') tx = x + w / 2;
    else tx = x + CELL_PAD + indent;

    ctx.textAlign = align === 'right' ? 'right' : align === 'center' ? 'center' : 'left';

    const v = style.align.v || 'bottom';
    let ty;
    if (v === 'top') {
      ctx.textBaseline = 'top';
      ty = y + 3;
    } else if (v === 'middle' || v === 'center') {
      ctx.textBaseline = 'middle';
      ty = y + h / 2;
    } else {
      ctx.textBaseline = 'alphabetic';
      ty = y + h - 5;
    }

    ctx.fillText(text, tx, ty);

    if (style.font.underline || style.font.strike) {
      const wpx = this.measure(ctx, text, ctx.font);
      const x0 = ctx.textAlign === 'right' ? tx - wpx : ctx.textAlign === 'center' ? tx - wpx / 2 : tx;
      const base = ctx.textBaseline === 'top' ? ty + ptToPx(style.font.size) : ctx.textBaseline === 'middle' ? ty + ptToPx(style.font.size) / 2.6 : ty;
      ctx.strokeStyle = ctx.fillStyle;
      ctx.lineWidth = 1;
      ctx.beginPath();
      if (style.font.underline) {
        ctx.moveTo(x0, Math.round(base + 1.5) + 0.5);
        ctx.lineTo(x0 + wpx, Math.round(base + 1.5) + 0.5);
      }
      if (style.font.strike) {
        const mid = base - ptToPx(style.font.size) * 0.3;
        ctx.moveTo(x0, Math.round(mid) + 0.5);
        ctx.lineTo(x0 + wpx, Math.round(mid) + 0.5);
      }
      ctx.stroke();
    }
  }

  drawWrapped(ctx, text, x, y, w, h, align, style) {
    const availW = w - CELL_PAD * 2;
    const lineH = ptToPx(style.font.size) * 1.25;
    const lines = this.wrapText(ctx, text, availW, ctx.font);

    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y, w, h);
    ctx.clip();

    const v = style.align.v || 'bottom';
    const totalH = lines.length * lineH;
    let startY;
    if (v === 'top') startY = y + 2;
    else if (v === 'middle' || v === 'center') startY = y + (h - totalH) / 2;
    else startY = y + h - totalH - 2;

    ctx.textBaseline = 'top';
    ctx.textAlign = align === 'right' ? 'right' : align === 'center' ? 'center' : 'left';
    const tx = align === 'right' ? x + w - CELL_PAD : align === 'center' ? x + w / 2 : x + CELL_PAD;

    lines.forEach((line, i) => {
      ctx.fillText(line, tx, startY + i * lineH);
    });
    ctx.restore();
  }

  wrapText(ctx, text, maxW, font) {
    const words = String(text).split(/(\s+)/);
    const lines = [];
    let cur = '';
    for (const word of words) {
      const test = cur + word;
      if (this.measure(ctx, test, font) > maxW && cur) {
        lines.push(cur.trimEnd());
        cur = word.trimStart();
      } else {
        cur = test;
      }
    }
    if (cur) lines.push(cur);
    return lines.length ? lines : [''];
  }

  drawBorders(ctx, rb, cb) {
    const sheet = this.sheet;
    for (let r = rb.start; r <= rb.end; r++) {
      if (isHiddenRow(sheet, r)) continue;
      for (let c = cb.start; c <= cb.end; c++) {
        if (isHiddenCol(sheet, c)) continue;
        const k = key(r, c);
        if (this.mergeCovered.has(k)) continue;
        const cell = sheet.cells[k];
        if (!cell) continue;
        const style = this.doc.styles.get(cell.s || 0);
        const border = style.border;
        if (!border) continue;

        const merge = this.mergeAnchor.get(k);
        const x = this.colX(c);
        const y = this.rowY(r);
        const w = merge ? this.colPos[merge.c2 + 1] - this.colPos[merge.c1] : this.colW(c);
        const h = merge ? this.rowPos[merge.r2 + 1] - this.rowPos[merge.r1] : this.rowH(r);

        for (const side of ['top', 'right', 'bottom', 'left']) {
          const b = border[side];
          if (!b) continue;
          ctx.strokeStyle = inkOn(style.fill, b.color, THEME.docBorder);
          ctx.lineWidth = borderWidth(b.style);
          ctx.setLineDash(borderDash(b.style));
          ctx.beginPath();
          const half = ctx.lineWidth % 2 === 1 ? 0.5 : 0;
          if (side === 'top') {
            ctx.moveTo(x, Math.round(y) + half);
            ctx.lineTo(x + w, Math.round(y) + half);
          } else if (side === 'bottom') {
            ctx.moveTo(x, Math.round(y + h) - half);
            ctx.lineTo(x + w, Math.round(y + h) - half);
          } else if (side === 'left') {
            ctx.moveTo(Math.round(x) + half, y);
            ctx.lineTo(Math.round(x) + half, y + h);
          } else {
            ctx.moveTo(Math.round(x + w) - half, y);
            ctx.lineTo(Math.round(x + w) - half, y + h);
          }
          ctx.stroke();
        }
        ctx.setLineDash([]);
      }
    }
  }

  drawSelection(ctx, rowBands, colBands) {
    const sel = this.expandToMerges(this.selection);
    const active = this.active;

    ctx.save();
    ctx.beginPath();
    ctx.rect(this.headerW, HEADER_H, this.viewW - this.headerW, this.viewH - HEADER_H);
    ctx.clip();

    const x1 = this.colX(sel.c1);
    const y1 = this.rowY(sel.r1);
    const x2 = this.colX(sel.c2) + this.colW(sel.c2);
    const y2 = this.rowY(sel.r2) + this.rowH(sel.r2);

    const isSingle = sel.r1 === sel.r2 && sel.c1 === sel.c2;
    if (!isSingle) {
      ctx.fillStyle = THEME.selFill;
      ctx.fillRect(x1, y1, x2 - x1, y2 - y1);

      // Punch the active cell back out so it stays visually distinct.
      const am = this.mergeAnchor.get(key(active.row, active.col));
      const ax = this.colX(active.col);
      const ay = this.rowY(active.row);
      const aw = am ? this.colPos[am.c2 + 1] - this.colPos[am.c1] : this.colW(active.col);
      const ah = am ? this.rowPos[am.r2 + 1] - this.rowPos[am.r1] : this.rowH(active.row);
      ctx.fillStyle = THEME.canvasBg;
      ctx.fillRect(ax + 1, ay + 1, aw - 2, ah - 2);
      this.repaintCell(ctx, active.row, active.col);
    }

    ctx.strokeStyle = THEME.selBorder;
    ctx.lineWidth = 2;
    ctx.strokeRect(Math.round(x1) + 1, Math.round(y1) + 1, Math.round(x2 - x1) - 2, Math.round(y2 - y1) - 2);

    // Fill handle.
    ctx.fillStyle = THEME.selBorder;
    ctx.fillRect(x2 - FILL_HANDLE / 2 - 1, y2 - FILL_HANDLE / 2 - 1, FILL_HANDLE, FILL_HANDLE);
    this.fillHandleRect = { x: x2 - FILL_HANDLE, y: y2 - FILL_HANDLE, w: FILL_HANDLE * 2, h: FILL_HANDLE * 2 };

    ctx.restore();
  }

  // Redraws one cell's fill, text and borders — used after the active cell is
  // knocked out of the selection tint.
  repaintCell(ctx, r, c) {
    const sheet = this.sheet;
    const cell = sheet.cells[key(r, c)];
    if (cell) {
      const style = this.doc.styles.get(cell.s || 0);
      if (style.fill) {
        const merge = this.mergeAnchor.get(key(r, c));
        const x = this.colX(c);
        const y = this.rowY(r);
        const w = merge ? this.colPos[merge.c2 + 1] - this.colPos[merge.c1] : this.colW(c);
        const h = merge ? this.rowPos[merge.r2 + 1] - this.rowPos[merge.r1] : this.rowH(r);
        ctx.fillStyle = style.fill;
        ctx.fillRect(x + 1, y + 1, w - 2, h - 2);
      }
    }
    this.drawCellContent(ctx, r, c, sheet, { start: c, end: c });
    this.drawBorders(ctx, { start: r, end: r }, { start: c, end: c });
  }

  drawHeaders(ctx, rowBands, colBands) {
    const sheet = this.sheet;
    const sel = this.expandToMerges(this.selection);

    ctx.fillStyle = THEME.headerBg;
    ctx.fillRect(0, 0, this.viewW, HEADER_H);
    ctx.fillRect(0, 0, this.headerW, this.viewH);

    ctx.font = '11px -apple-system, BlinkMacSystemFont, sans-serif';
    ctx.textBaseline = 'middle';

    // Column headers.
    ctx.save();
    ctx.beginPath();
    ctx.rect(this.headerW, 0, this.viewW - this.headerW, HEADER_H);
    ctx.clip();
    for (const cb of colBands) {
      for (let c = cb.start; c <= cb.end; c++) {
        if (isHiddenCol(sheet, c)) continue;
        const x = this.colX(c);
        const w = this.colW(c);
        if (w <= 0) continue;
        const activeCol = c >= sel.c1 && c <= sel.c2;
        ctx.fillStyle = activeCol ? THEME.headerBgActive : THEME.headerBg;
        ctx.fillRect(x, 0, w, HEADER_H);
        ctx.strokeStyle = THEME.headerLine;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(Math.round(x + w) + 0.5, 0);
        ctx.lineTo(Math.round(x + w) + 0.5, HEADER_H);
        ctx.stroke();
        ctx.fillStyle = activeCol ? THEME.headerTextActive : THEME.headerText;
        ctx.textAlign = 'center';
        ctx.fillText(colToLetter(c), x + w / 2, HEADER_H / 2 + 1);
      }
    }
    ctx.restore();

    // Row headers.
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, HEADER_H, this.headerW, this.viewH - HEADER_H);
    ctx.clip();
    for (const rb of rowBands) {
      for (let r = rb.start; r <= rb.end; r++) {
        if (isHiddenRow(sheet, r)) continue;
        const y = this.rowY(r);
        const h = this.rowH(r);
        if (h <= 0) continue;
        const activeRow = r >= sel.r1 && r <= sel.r2;
        ctx.fillStyle = activeRow ? THEME.headerBgActive : THEME.headerBg;
        ctx.fillRect(0, y, this.headerW, h);
        ctx.strokeStyle = THEME.headerLine;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(0, Math.round(y + h) + 0.5);
        ctx.lineTo(this.headerW, Math.round(y + h) + 0.5);
        ctx.stroke();
        ctx.fillStyle = activeRow ? THEME.headerTextActive : THEME.headerText;
        ctx.textAlign = 'right';
        ctx.fillText(String(r), this.headerW - 6, y + h / 2 + 1);
      }
    }
    ctx.restore();

    // Corner box and header separators.
    ctx.fillStyle = THEME.headerBg;
    ctx.fillRect(0, 0, this.headerW, HEADER_H);
    ctx.strokeStyle = THEME.headerLine;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(0, HEADER_H + 0.5);
    ctx.lineTo(this.viewW, HEADER_H + 0.5);
    ctx.moveTo(this.headerW + 0.5, 0);
    ctx.lineTo(this.headerW + 0.5, this.viewH);
    ctx.stroke();
  }

  drawFrozenDividers(ctx) {
    if (!this.frozenRows && !this.frozenCols) return;
    ctx.strokeStyle = THEME.frozenLine;
    ctx.lineWidth = 1;
    ctx.beginPath();
    if (this.frozenCols) {
      const x = Math.round(this.headerW + this.frozenW) + 0.5;
      ctx.moveTo(x, 0);
      ctx.lineTo(x, this.viewH);
    }
    if (this.frozenRows) {
      const y = Math.round(HEADER_H + this.frozenH) + 0.5;
      ctx.moveTo(0, y);
      ctx.lineTo(this.viewW, y);
    }
    ctx.stroke();
  }

  measure(ctx, text, font) {
    const k = font + '\u0000' + text;
    let w = this.measureCache.get(k);
    if (w === undefined) {
      const prev = ctx.font;
      ctx.font = font;
      w = ctx.measureText(text).width;
      ctx.font = prev;
      if (this.measureCache.size > 30000) this.measureCache.clear();
      this.measureCache.set(k, w);
    }
    return w;
  }

  // ------------------------------------------------------------ hit testing

  // Screen point -> what is under it.
  hitTest(px, py) {
    const inColHeader = py < HEADER_H;
    const inRowHeader = px < this.headerW;

    if (inColHeader && inRowHeader) return { kind: 'corner' };

    if (inColHeader) {
      const c = this.colAt(px);
      if (c === null) return { kind: 'none' };
      // Near a column edge? That's a resize grip.
      const right = this.colX(c) + this.colW(c);
      if (Math.abs(px - right) <= RESIZE_GRAB) return { kind: 'colResize', col: c };
      const left = this.colX(c);
      if (px - left <= RESIZE_GRAB && c > 1) {
        const prev = prevVisibleCol(this.sheet, c);
        if (prev) return { kind: 'colResize', col: prev };
      }
      return { kind: 'colHeader', col: c };
    }

    if (inRowHeader) {
      const r = this.rowAt(py);
      if (r === null) return { kind: 'none' };
      const bottom = this.rowY(r) + this.rowH(r);
      if (Math.abs(py - bottom) <= RESIZE_GRAB) return { kind: 'rowResize', row: r };
      const top = this.rowY(r);
      if (py - top <= RESIZE_GRAB && r > 1) {
        const prev = prevVisibleRow(this.sheet, r);
        if (prev) return { kind: 'rowResize', row: prev };
      }
      return { kind: 'rowHeader', row: r };
    }

    if (this.fillHandleRect && pointIn(px, py, this.fillHandleRect)) return { kind: 'fillHandle' };

    const r = this.rowAt(py);
    const c = this.colAt(px);
    if (r === null || c === null) return { kind: 'none' };
    return { kind: 'cell', row: r, col: c };
  }

  colAt(px) {
    if (px < this.headerW) return null;
    if (this.frozenCols && px < this.headerW + this.frozenW) {
      return binarySearch(this.colPos, px - this.headerW, this.frozenCols);
    }
    const content = px - this.headerW + this.scrollX;
    const c = binarySearch(this.colPos, content, this.nCols);
    return Math.min(this.nCols, Math.max(1, c));
  }

  rowAt(py) {
    if (py < HEADER_H) return null;
    if (this.frozenRows && py < HEADER_H + this.frozenH) {
      return binarySearch(this.rowPos, py - HEADER_H, this.frozenRows);
    }
    const content = py - HEADER_H + this.scrollY;
    const r = binarySearch(this.rowPos, content, this.nRows);
    return Math.min(this.nRows, Math.max(1, r));
  }

  localPoint(e) {
    const rect = this.container.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  }

  // ------------------------------------------------------------- selection

  setSelection(range, active, opts) {
    this.selection = {
      r1: Math.max(1, Math.min(range.r1, range.r2)),
      c1: Math.max(1, Math.min(range.c1, range.c2)),
      r2: Math.min(this.nRows, Math.max(range.r1, range.r2)),
      c2: Math.min(this.nCols, Math.max(range.c1, range.c2)),
    };
    if (active) this.active = { row: active.row, col: active.col };
    if (!opts || opts.scroll !== false) this.scrollIntoView(this.active.row, this.active.col);
    this.requestRender();
    if (this.onSelectionChange) this.onSelectionChange(this.selection, this.active);
  }

  selectCell(row, col) {
    this.setSelection({ r1: row, c1: col, r2: row, c2: col }, { row, col });
  }

  extendTo(row, col) {
    this.setSelection(
      { r1: this.anchor ? this.anchor.row : this.active.row, c1: this.anchor ? this.anchor.col : this.active.col, r2: row, c2: col },
      this.active,
      { scroll: false }
    );
    this.scrollIntoView(row, col);
  }

  scrollIntoView(row, col) {
    if (!this.colPos) return;
    // Frozen cells are always on screen.
    if (col > this.frozenCols) {
      const left = this.colPos[col] - this.colPos[this.frozenCols + 1];
      const right = this.colPos[col + 1] - this.colPos[this.frozenCols + 1];
      if (left < this.scroller.scrollLeft) this.scroller.scrollLeft = left;
      else if (right > this.scroller.scrollLeft + this.paneW) this.scroller.scrollLeft = right - this.paneW;
    }
    if (row > this.frozenRows) {
      const top = this.rowPos[row] - this.rowPos[this.frozenRows + 1];
      const bottom = this.rowPos[row + 1] - this.rowPos[this.frozenRows + 1];
      if (top < this.scroller.scrollTop) this.scroller.scrollTop = top;
      else if (bottom > this.scroller.scrollTop + this.paneH) this.scroller.scrollTop = bottom - this.paneH;
    }
  }

  // ---------------------------------------------------------------- events

  bindEvents() {
    this.scroller.addEventListener('scroll', () => this.requestRender(), { passive: true });

    this.scroller.addEventListener('mousedown', (e) => this.onMouseDown(e));
    this.scroller.addEventListener('mousemove', (e) => this.onMouseMove(e));
    this.scroller.addEventListener('dblclick', (e) => this.onDoubleClick(e));
    this.scroller.addEventListener('contextmenu', (e) => this.onContext(e));
    this.scroller.addEventListener('keydown', (e) => this.handleKeyDown(e));

    window.addEventListener('mouseup', (e) => this.onMouseUp(e));
    window.addEventListener('mousemove', (e) => this.onDragMove(e), true);
  }

  onMouseDown(e) {
    if (e.button === 2) return;
    const p = this.localPoint(e);
    const hit = this.hitTest(p.x, p.y);
    this.focus();

    if (hit.kind === 'colResize') {
      // The original spec is kept so the live preview can be rolled back
      // before the committed width goes through the undo stack.
      const spec = this.sheet.cols[hit.col];
      this.drag = {
        kind: 'colResize',
        col: hit.col,
        startX: p.x,
        startW: this.colW(hit.col),
        startSpec: spec ? Object.assign({}, spec) : null,
      };
      e.preventDefault();
      return;
    }
    if (hit.kind === 'rowResize') {
      const spec = this.sheet.rows[hit.row];
      this.drag = {
        kind: 'rowResize',
        row: hit.row,
        startY: p.y,
        startH: this.rowH(hit.row),
        startSpec: spec ? Object.assign({}, spec) : null,
      };
      e.preventDefault();
      return;
    }
    if (hit.kind === 'corner') {
      this.anchor = { row: 1, col: 1 };
      this.setSelection({ r1: 1, c1: 1, r2: this.nRows, c2: this.nCols }, { row: 1, col: 1 }, { scroll: false });
      return;
    }
    if (hit.kind === 'colHeader') {
      this.anchor = { row: 1, col: hit.col };
      this.drag = { kind: 'colSelect' };
      const c1 = e.shiftKey && this.lastHeaderCol ? this.lastHeaderCol : hit.col;
      this.lastHeaderCol = c1;
      this.setSelection({ r1: 1, c1, r2: this.nRows, c2: hit.col }, { row: 1, col: hit.col }, { scroll: false });
      return;
    }
    if (hit.kind === 'rowHeader') {
      this.anchor = { row: hit.row, col: 1 };
      this.drag = { kind: 'rowSelect' };
      const r1 = e.shiftKey && this.lastHeaderRow ? this.lastHeaderRow : hit.row;
      this.lastHeaderRow = r1;
      this.setSelection({ r1, c1: 1, r2: hit.row, c2: this.nCols }, { row: hit.row, col: 1 }, { scroll: false });
      return;
    }
    if (hit.kind === 'fillHandle') {
      const sel = this.expandToMerges(this.selection);
      this.drag = { kind: 'fill', source: sel, target: Object.assign({}, sel) };
      e.preventDefault();
      return;
    }
    if (hit.kind === 'cell') {
      // Clicking inside a merge selects the whole merge.
      const coveredBy = this.mergeCovered.get(key(hit.row, hit.col));
      let row = hit.row;
      let col = hit.col;
      if (coveredBy) {
        const parts = coveredBy.split(',');
        row = +parts[0];
        col = +parts[1];
      }
      if (e.shiftKey) {
        this.anchor = this.anchor || { row: this.active.row, col: this.active.col };
        this.extendTo(row, col);
      } else {
        this.anchor = { row, col };
        this.selectCell(row, col);
      }
      this.drag = { kind: 'cellSelect' };
    }
  }

  onMouseMove(e) {
    if (this.drag) return;
    const p = this.localPoint(e);
    const hit = this.hitTest(p.x, p.y);
    let cursor = 'cell';
    if (hit.kind === 'colResize') cursor = 'col-resize';
    else if (hit.kind === 'rowResize') cursor = 'row-resize';
    else if (hit.kind === 'colHeader' || hit.kind === 'rowHeader' || hit.kind === 'corner') cursor = 'default';
    else if (hit.kind === 'fillHandle') cursor = 'crosshair';
    this.scroller.style.cursor = cursor;
  }

  onDragMove(e) {
    if (!this.drag) return;
    const p = this.localPoint(e);

    if (this.drag.kind === 'colResize') {
      const w = Math.max(0, this.drag.startW + (p.x - this.drag.startX));
      this.previewColWidth(this.drag.col, w);
      return;
    }
    if (this.drag.kind === 'rowResize') {
      const h = Math.max(0, this.drag.startH + (p.y - this.drag.startY));
      this.previewRowHeight(this.drag.row, h);
      return;
    }

    this.autoScrollDuringDrag(p);

    const r = this.rowAt(Math.max(HEADER_H + 1, p.y));
    const c = this.colAt(Math.max(this.headerW + 1, p.x));
    if (r === null || c === null) return;

    if (this.drag.kind === 'cellSelect') {
      this.extendTo(r, c);
    } else if (this.drag.kind === 'colSelect') {
      this.setSelection({ r1: 1, c1: this.anchor.col, r2: this.nRows, c2: c }, this.active, { scroll: false });
    } else if (this.drag.kind === 'rowSelect') {
      this.setSelection({ r1: this.anchor.row, c1: 1, r2: r, c2: this.nCols }, this.active, { scroll: false });
    } else if (this.drag.kind === 'fill') {
      this.drag.target = fillTarget(this.drag.source, r, c);
      this.requestRender();
      this.drawFillPreview();
    }
  }

  autoScrollDuringDrag(p) {
    const edge = 30;
    let dx = 0;
    let dy = 0;
    if (p.x > this.viewW - edge) dx = (p.x - (this.viewW - edge)) / 2;
    else if (p.x < this.headerW + edge) dx = -((this.headerW + edge - p.x) / 2);
    if (p.y > this.viewH - edge) dy = (p.y - (this.viewH - edge)) / 2;
    else if (p.y < HEADER_H + edge) dy = -((HEADER_H + edge - p.y) / 2);
    if (dx) this.scroller.scrollLeft += dx;
    if (dy) this.scroller.scrollTop += dy;
  }

  drawFillPreview() {
    if (!this.drag || this.drag.kind !== 'fill') return;
    const t = this.drag.target;
    const ctx = this.ctx;
    ctx.save();
    ctx.beginPath();
    ctx.rect(this.headerW, HEADER_H, this.viewW - this.headerW, this.viewH - HEADER_H);
    ctx.clip();
    ctx.strokeStyle = '#6b6b6b';
    ctx.setLineDash([4, 3]);
    ctx.lineWidth = 1;
    const x = this.colX(t.c1);
    const y = this.rowY(t.r1);
    const w = this.colX(t.c2) + this.colW(t.c2) - x;
    const h = this.rowY(t.r2) + this.rowH(t.r2) - y;
    ctx.strokeRect(Math.round(x) + 0.5, Math.round(y) + 0.5, Math.round(w), Math.round(h));
    ctx.setLineDash([]);
    ctx.restore();
  }

  previewColWidth(col, px) {
    this.pendingResize = { kind: 'col', col, px };
    const sheet = this.sheet;
    sheet.cols[col] = Object.assign({}, sheet.cols[col], { width: pxToWidth(px) });
    this.layout();
    this.render();
  }

  previewRowHeight(row, px) {
    this.pendingResize = { kind: 'row', row, px };
    const sheet = this.sheet;
    sheet.rows[row] = Object.assign({}, sheet.rows[row], { height: pxToHeight(px) });
    this.layout();
    this.render();
  }

  onMouseUp() {
    const drag = this.drag;
    this.drag = null;
    if (!drag) return;

    if ((drag.kind === 'colResize' || drag.kind === 'rowResize') && this.pendingResize) {
      const pr = this.pendingResize;
      this.pendingResize = null;

      // Roll the live preview back to the pre-drag spec, then let the
      // controller apply the new size as a real, undoable operation.
      const sheet = this.sheet;
      if (pr.kind === 'col') {
        if (drag.startSpec) sheet.cols[pr.col] = drag.startSpec;
        else delete sheet.cols[pr.col];
      } else {
        if (drag.startSpec) sheet.rows[pr.row] = drag.startSpec;
        else delete sheet.rows[pr.row];
      }

      if (this.onStructuralChange) {
        this.onStructuralChange(
          pr.kind === 'col'
            ? { type: 'colWidth', col: pr.col, width: pxToWidth(pr.px) }
            : { type: 'rowHeight', row: pr.row, height: pxToHeight(pr.px) }
        );
      }
      return;
    }

    if (drag.kind === 'fill' && this.onFill) {
      const t = drag.target;
      if (t.r1 !== drag.source.r1 || t.c1 !== drag.source.c1 || t.r2 !== drag.source.r2 || t.c2 !== drag.source.c2) {
        this.onFill(drag.source, t);
      }
      this.requestRender();
    }
  }

  onDoubleClick(e) {
    const p = this.localPoint(e);
    const hit = this.hitTest(p.x, p.y);
    if (hit.kind === 'colResize') {
      const w = this.autofitColumn(hit.col);
      if (this.onStructuralChange) this.onStructuralChange({ type: 'colWidth', col: hit.col, width: w });
      return;
    }
    if (hit.kind === 'cell' && this.onEditRequest) {
      this.onEditRequest(this.active.row, this.active.col, null);
    }
  }

  onContext(e) {
    e.preventDefault();
    const p = this.localPoint(e);
    const hit = this.hitTest(p.x, p.y);
    if (hit.kind === 'cell') {
      const inside =
        hit.row >= this.selection.r1 && hit.row <= this.selection.r2 && hit.col >= this.selection.c1 && hit.col <= this.selection.c2;
      if (!inside) {
        this.anchor = { row: hit.row, col: hit.col };
        this.selectCell(hit.row, hit.col);
      }
    } else if (hit.kind === 'colHeader') {
      const inside = hit.col >= this.selection.c1 && hit.col <= this.selection.c2 && this.selection.r2 >= this.nRows;
      if (!inside) this.setSelection({ r1: 1, c1: hit.col, r2: this.nRows, c2: hit.col }, { row: 1, col: hit.col }, { scroll: false });
    } else if (hit.kind === 'rowHeader') {
      const inside = hit.row >= this.selection.r1 && hit.row <= this.selection.r2 && this.selection.c2 >= this.nCols;
      if (!inside) this.setSelection({ r1: hit.row, c1: 1, r2: hit.row, c2: this.nCols }, { row: hit.row, col: 1 }, { scroll: false });
    }
    if (this.onContextMenu) this.onContextMenu(e, hit);
  }

  // Computes the width in Excel units that fits the widest value in a column.
  autofitColumn(col, maxScanRows) {
    const sheet = this.sheet;
    const ctx = this.ctx;
    let widest = 0;
    const limit = Math.min(sheet.maxRow, maxScanRows || 5000);
    for (let r = 1; r <= limit; r++) {
      const cell = sheet.cells[key(r, col)];
      if (!cell || cell.v === undefined || cell.v === null) continue;
      const style = this.doc.styles.get(cell.s || 0);
      const { text } = formatValue(cell.v, style.numFmt);
      if (!text) continue;
      const w = this.measure(ctx, text, fontString(style.font));
      if (w > widest) widest = w;
    }
    const px = Math.min(600, Math.max(24, widest + CELL_PAD * 2 + 3));
    return pxToWidth(px);
  }

  handleKeyDown(e) {
    const meta = e.metaKey || e.ctrlKey;
    const sel = this.selection;
    const act = this.active;

    const moveTo = (row, col, extend) => {
      row = Math.max(1, Math.min(this.nRows, row));
      col = Math.max(1, Math.min(this.nCols, col));
      if (extend) {
        this.anchor = this.anchor || { row: act.row, col: act.col };
        this.setSelection({ r1: this.anchor.row, c1: this.anchor.col, r2: row, c2: col }, act, { scroll: false });
        this.scrollIntoView(row, col);
      } else {
        this.anchor = { row, col };
        this.selectCell(row, col);
      }
      e.preventDefault();
    };

    switch (e.key) {
      case 'ArrowUp':
        return moveTo(meta ? this.edgeRow(act.row, act.col, -1) : act.row - 1, act.col, e.shiftKey);
      case 'ArrowDown':
        return moveTo(meta ? this.edgeRow(act.row, act.col, 1) : act.row + 1, act.col, e.shiftKey);
      case 'ArrowLeft':
        return moveTo(act.row, meta ? this.edgeCol(act.row, act.col, -1) : act.col - 1, e.shiftKey);
      case 'ArrowRight':
        return moveTo(act.row, meta ? this.edgeCol(act.row, act.col, 1) : act.col + 1, e.shiftKey);
      case 'Tab':
        return moveTo(act.row, act.col + (e.shiftKey ? -1 : 1), false);
      case 'Enter':
        // Enter on a selected cell steps down (Shift+Enter steps up); F2 or
        // simply typing is what opens the editor.
        return moveTo(act.row + (e.shiftKey ? -1 : 1), act.col, false);
      case 'Home':
        return moveTo(meta ? 1 : act.row, 1, e.shiftKey);
      case 'End':
        return moveTo(meta ? Math.max(1, this.sheet.maxRow) : act.row, Math.max(1, this.sheet.maxCol), e.shiftKey);
      case 'PageDown':
        return moveTo(act.row + Math.max(1, Math.floor(this.paneH / this.defRowPx)), act.col, e.shiftKey);
      case 'PageUp':
        return moveTo(act.row - Math.max(1, Math.floor(this.paneH / this.defRowPx)), act.col, e.shiftKey);
      case 'F2':
        if (this.onEditRequest) {
          this.onEditRequest(act.row, act.col, null);
          e.preventDefault();
        }
        return;
      default:
        break;
    }

    if (this.onKeyDown) this.onKeyDown(e);
  }

  // Ctrl/Cmd+Arrow: jump to the edge of the current block of data.
  edgeRow(row, col, dir) {
    const sheet = this.sheet;
    const occupied = (r) => isOccupied(sheet, r, col);
    let r = row + dir;
    if (r < 1) return 1;
    if (occupied(row) && occupied(r)) {
      while (r + dir >= 1 && r + dir <= this.nRows && occupied(r + dir)) r += dir;
      return r;
    }
    while (r >= 1 && r <= this.nRows && !occupied(r)) r += dir;
    return Math.max(1, Math.min(this.nRows, r));
  }

  edgeCol(row, col, dir) {
    const sheet = this.sheet;
    const occupied = (c) => isOccupied(sheet, row, c);
    let c = col + dir;
    if (c < 1) return 1;
    if (occupied(col) && occupied(c)) {
      while (c + dir >= 1 && c + dir <= this.nCols && occupied(c + dir)) c += dir;
      return c;
    }
    while (c >= 1 && c <= this.nCols && !occupied(c)) c += dir;
    return Math.max(1, Math.min(this.nCols, c));
  }

  // Screen rectangle of a cell, for positioning the editor overlay.
  cellRect(row, col) {
    const merge = this.mergeAnchor.get(key(row, col));
    const r = merge || { r1: row, c1: col, r2: row, c2: col };
    return {
      x: this.colX(r.c1),
      y: this.rowY(r.r1),
      w: this.colPos[r.c2 + 1] - this.colPos[r.c1],
      h: this.rowPos[r.r2 + 1] - this.rowPos[r.r1],
    };
  }

  destroy() {
    this.resizeObserver.disconnect();
  }
}

// ------------------------------------------------------------------ helpers

function colWidthPx(sheet, col, defPx) {
  const spec = sheet.cols[col];
  if (spec && spec.hidden) return 0;
  if (spec && spec.width != null) return widthToPx(spec.width);
  return defPx;
}

function rowHeightPx(sheet, row, defPx) {
  const spec = sheet.rows[row];
  if (spec && spec.hidden) return 0;
  if (spec && spec.height != null) return heightToPx(spec.height);
  return defPx;
}

function isHiddenRow(sheet, r) {
  const spec = sheet.rows[r];
  return !!(spec && spec.hidden);
}

function isHiddenCol(sheet, c) {
  const spec = sheet.cols[c];
  return !!(spec && spec.hidden);
}

function prevVisibleCol(sheet, c) {
  for (let i = c - 1; i >= 1; i--) if (!isHiddenCol(sheet, i)) return i;
  return null;
}

function prevVisibleRow(sheet, r) {
  for (let i = r - 1; i >= 1; i--) if (!isHiddenRow(sheet, i)) return i;
  return null;
}

function isOccupied(sheet, r, c) {
  const cell = sheet.cells[key(r, c)];
  return !!(cell && cell.v !== undefined && cell.v !== null && cell.v !== '');
}

export function fontString(font) {
  const parts = [];
  if (font.italic) parts.push('italic');
  if (font.bold) parts.push('bold');
  parts.push(Math.round(ptToPx(font.size || 11)) + 'px');
  const name = font.name || 'Calibri';
  parts.push(`"${name}", "Helvetica Neue", Arial, sans-serif`);
  return parts.join(' ');
}

function borderWidth(style) {
  switch (style) {
    case 'thick':
    case 'double':
      return 3;
    case 'medium':
    case 'mediumDashed':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
      return 2;
    default:
      return 1;
  }
}

function borderDash(style) {
  switch (style) {
    case 'dashed':
    case 'mediumDashed':
      return [4, 3];
    case 'dotted':
      return [1, 2];
    case 'dashDot':
    case 'mediumDashDot':
      return [5, 2, 1, 2];
    case 'dashDotDot':
    case 'mediumDashDotDot':
      return [5, 2, 1, 2, 1, 2];
    case 'hair':
      return [1, 1];
    default:
      return [];
  }
}

// Largest i in [1, max] with pos[i] <= value.
function binarySearch(pos, value, max) {
  let lo = 1;
  let hi = max;
  let best = 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (pos[mid] <= value) {
      best = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return best;
}

function pointIn(x, y, r) {
  return x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h;
}

// Excel's fill handle extends along one axis only — whichever the pointer has
// travelled furthest in.
function fillTarget(source, row, col) {
  const t = Object.assign({}, source);
  const dRow = row > source.r2 ? row - source.r2 : row < source.r1 ? source.r1 - row : 0;
  const dCol = col > source.c2 ? col - source.c2 : col < source.c1 ? source.c1 - col : 0;
  if (dRow >= dCol) {
    if (row > source.r2) t.r2 = row;
    else if (row < source.r1) t.r1 = row;
  } else {
    if (col > source.c2) t.c2 = col;
    else if (col < source.c1) t.c1 = col;
  }
  return t;
}
