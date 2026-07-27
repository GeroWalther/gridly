import { fontString } from './grid.js';
import { rcToA1 } from '../shared/addr.js';

// Cell editing happens in a textarea floated over the grid canvas, mirrored
// live into the formula bar (and vice versa). Two modes, as in Excel:
//   "enter" — you started by typing, so arrow keys commit and move on
//   "edit"  — you started with F2 or a double-click, so arrows move the caret

export class Editor {
  constructor({ grid, doc, host, formulaInput, onCommit }) {
    this.grid = grid;
    this.doc = doc;
    this.host = host;
    this.formulaInput = formulaInput;
    this.onCommit = onCommit;

    this.open = false;
    this.mode = 'enter';
    this.row = 1;
    this.col = 1;

    this.input = document.createElement('textarea');
    this.input.className = 'cell-editor';
    this.input.spellcheck = false;
    this.input.style.display = 'none';
    host.appendChild(this.input);

    this.input.addEventListener('keydown', (e) => this.onKeyDown(e, false));
    this.input.addEventListener('input', () => {
      if (this.syncing) return;
      this.syncing = true;
      this.formulaInput.value = this.input.value;
      this.syncing = false;
      this.autosize();
    });

    this.formulaInput.addEventListener('keydown', (e) => this.onKeyDown(e, true));
    this.formulaInput.addEventListener('input', () => {
      if (this.syncing) return;
      if (!this.open) {
        const { row, col } = this.grid.active;
        this.begin(row, col, { mode: 'edit', initial: this.formulaInput.value, focus: 'formula' });
        return;
      }
      this.syncing = true;
      this.input.value = this.formulaInput.value;
      this.syncing = false;
      this.autosize();
    });
    this.formulaInput.addEventListener('focus', () => {
      if (!this.open) {
        const { row, col } = this.grid.active;
        this.begin(row, col, { mode: 'edit', initial: null, focus: 'formula' });
      }
    });
  }

  // Refreshes the formula bar to match the current selection.
  syncFromSelection() {
    if (this.open) return;
    const { row, col } = this.grid.active;
    this.formulaInput.value = this.doc.editText(row, col);
  }

  begin(row, col, opts) {
    const options = opts || {};
    this.row = row;
    this.col = col;
    this.mode = options.mode || 'enter';
    this.open = true;

    const text = options.initial != null ? options.initial : this.doc.editText(row, col);
    this.input.value = text;
    this.formulaInput.value = text;

    const style = this.doc.styleOf(row, col);
    this.input.style.font = fontString(style.font);
    this.input.style.color = style.font.color || '#000';
    this.input.style.background = style.fill || '#ffffff';
    this.input.style.textAlign = style.align.h && style.align.h !== 'general' ? style.align.h : 'left';

    this.position();
    this.input.style.display = 'block';

    if (options.focus === 'formula') {
      this.formulaInput.focus();
    } else {
      this.input.focus();
      const end = this.input.value.length;
      this.input.setSelectionRange(end, end);
    }
    this.autosize();
  }

  position() {
    const rect = this.grid.cellRect(this.row, this.col);
    this.input.style.left = rect.x + 'px';
    this.input.style.top = rect.y + 'px';
    this.minW = Math.max(rect.w, 30);
    this.minH = Math.max(rect.h, 18);
    this.input.style.width = this.minW + 'px';
    this.input.style.height = this.minH + 'px';
  }

  // Grows the editor past the cell when the content needs more room.
  autosize() {
    if (!this.open) return;
    this.input.style.height = 'auto';
    const needed = Math.max(this.minH, this.input.scrollHeight);
    this.input.style.height = needed + 'px';

    const probe = document.createElement('span');
    probe.style.cssText = 'position:absolute;visibility:hidden;white-space:pre;';
    probe.style.font = this.input.style.font;
    probe.textContent = this.input.value || ' ';
    document.body.appendChild(probe);
    const textW = probe.getBoundingClientRect().width + 14;
    probe.remove();
    this.input.style.width = Math.max(this.minW, Math.min(textW, 600)) + 'px';
  }

  commit(move) {
    if (!this.open) return;
    const value = this.input.value;
    const { row, col } = this;
    this.close();
    this.doc.setCellInput(row, col, value);
    if (this.onCommit) this.onCommit(row, col, move);
  }

  cancel() {
    if (!this.open) return;
    this.close();
    this.syncFromSelection();
    this.grid.focus();
  }

  close() {
    this.open = false;
    this.input.style.display = 'none';
    this.input.value = '';
  }

  onKeyDown(e, fromFormulaBar) {
    if (!this.open) return;
    const meta = e.metaKey || e.ctrlKey;

    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      this.cancel();
      return;
    }

    if (e.key === 'Enter') {
      if (e.altKey) {
        // Alt+Enter inserts a line break, as in Excel.
        e.preventDefault();
        const start = this.input.selectionStart;
        const end = this.input.selectionEnd;
        this.input.value = this.input.value.slice(0, start) + '\n' + this.input.value.slice(end);
        this.input.setSelectionRange(start + 1, start + 1);
        this.formulaInput.value = this.input.value;
        this.autosize();
        return;
      }
      e.preventDefault();
      e.stopPropagation();
      this.commit(e.shiftKey ? 'up' : 'down');
      return;
    }

    if (e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      this.commit(e.shiftKey ? 'left' : 'right');
      return;
    }

    // In "enter" mode the arrow keys behave like Enter with a direction.
    if (this.mode === 'enter' && !fromFormulaBar && !meta) {
      const dirs = { ArrowUp: 'up', ArrowDown: 'down', ArrowLeft: 'left', ArrowRight: 'right' };
      if (dirs[e.key]) {
        e.preventDefault();
        e.stopPropagation();
        this.commit(dirs[e.key]);
        return;
      }
    }

    // Everything else is ordinary text editing; keep it away from the grid.
    e.stopPropagation();
  }
}
