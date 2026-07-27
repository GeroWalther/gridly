# Gridly

A spreadsheet app for macOS that reads and writes real Excel files.

Gridly opens `.xlsx`, `.xlsm`, `.xls` and `.csv`, keeps the formatting Excel
authored, recalculates formulas, and writes `.xlsx` back out. It is a native
macOS app built on Electron — not a web page in a wrapper.

## Features

- **Real Excel I/O** — cell values, formulas with cached results, number
  formats, fonts, fills, borders, merges, frozen panes, column widths, row
  heights, hidden rows/columns, sheet tab colours and cross-sheet references
  survive a round trip.
- **Formula engine** — HyperFormula, with Excel-compatible function coverage.
- **Canvas grid** — the sheet paints to a canvas with a transparent scroll
  layer over it, so scrolling uses real macOS scrollbars and trackpad
  momentum while staying fast on large sheets.
- **Editing** — multi-cell selection, fill handle, cut/copy/paste (including
  formulas and formatting), undo/redo, insert/delete rows and columns, sort,
  find, and column/row resizing.
- **Dark mode** — follows the system appearance, or pin it under
  *View → Appearance*. Workbook colours are left exactly as authored; only
  ink that was written for a white sheet is remapped so it stays legible.
- **Printing** — prints the used range, which also covers Save as PDF.
- `.xls` and `.xlsm` open read-only and save to a `.xlsx` sibling.

## Requirements

macOS on Apple Silicon. Node 18+ to build.

## Development

```sh
npm install
npm start          # bundles the renderer and launches Electron
```

## Building

```sh
npm run pack       # signed .app in release/mac-arm64/
```

Packaging signs with a Developer ID and enables the hardened runtime, which
needs the JIT entitlements in `build/entitlements.mac.plist` — Chromium will
not start under the hardened runtime without them. The build is not notarized,
so a copy downloaded onto another Mac will be held by Gatekeeper until it is.

## Layout

```
src/main/        Electron main process — windows, menus, dialogs, file I/O
  workbook-io.js  xlsx/xls/csv read and write (ExcelJS + SheetJS)
src/preload/     the contextIsolated bridge; the renderer has no Node access
src/renderer/    the app itself
  grid.js         canvas grid: layout, painting, pointer and key handling
  doc.js          document model, undo/redo, structural edits
  engine.js       HyperFormula bridge
  editor.js       in-cell editor
src/shared/       code both processes use (addressing, empty-workbook model)
```

## License

MIT
