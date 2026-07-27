'use strict';

const { app, BrowserWindow, ipcMain, dialog, shell, Menu, clipboard, nativeTheme } = require('electron');
const path = require('path');
const fs = require('fs');
const { readWorkbook, writeWorkbook } = require('./workbook-io');
const { buildMenu } = require('./menu');

const isMac = process.platform === 'darwin';

// Files handed to us before the app finished starting (double-click on a
// .xlsx while Gridly is not running) wait here until we can open a window.
const pendingOpen = [];
let ready = false;

const SPREADSHEET_FILTERS = [
  { name: 'Excel Workbook', extensions: ['xlsx', 'xlsm'] },
  { name: 'Excel 97-2003', extensions: ['xls'] },
  { name: 'Comma Separated Values', extensions: ['csv'] },
  { name: 'All Spreadsheets', extensions: ['xlsx', 'xlsm', 'xls', 'csv'] },
];

// ----------------------------------------------------------------- theme
//
// themeSource is 'system' | 'light' | 'dark'. The choice is remembered between
// launches; what the renderer is told is always the resolved light/dark, so it
// never has to ask the OS itself.

const PREFS_FILE = () => path.join(app.getPath('userData'), 'prefs.json');

function readPrefs() {
  try {
    return JSON.parse(fs.readFileSync(PREFS_FILE(), 'utf8'));
  } catch {
    return {};
  }
}

function writePrefs(prefs) {
  try {
    fs.mkdirSync(path.dirname(PREFS_FILE()), { recursive: true });
    fs.writeFileSync(PREFS_FILE(), JSON.stringify(prefs, null, 2));
  } catch {
    // A prefs file we cannot write is not worth failing a launch over.
  }
}

function currentTheme() {
  return nativeTheme.shouldUseDarkColors ? 'dark' : 'light';
}

function windowBackground() {
  return currentTheme() === 'dark' ? '#1c1c1f' : '#f5f6f7';
}

function broadcastTheme() {
  const theme = currentTheme();
  for (const win of BrowserWindow.getAllWindows()) {
    win.setBackgroundColor(windowBackground());
    win.webContents.send('app:theme', theme);
  }
}

function setThemeSource(source) {
  nativeTheme.themeSource = source;
  writePrefs(Object.assign(readPrefs(), { themeSource: source }));
  Menu.setApplicationMenu(buildMenu(menuContext()));
  broadcastTheme();
}

function createWindow(filePath) {
  const win = new BrowserWindow({
    width: 1360,
    height: 860,
    minWidth: 720,
    minHeight: 480,
    show: false,
    backgroundColor: windowBackground(),
    titleBarStyle: 'hiddenInset',
    trafficLightPosition: { x: 14, y: 18 },
    webPreferences: {
      preload: path.join(__dirname, '..', 'preload', 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: false,
    },
  });

  win.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));

  win.once('ready-to-show', () => {
    win.webContents.send('app:theme', currentTheme());
    win.show();
    if (filePath) win.webContents.send('app:openPath', filePath);
  });

  // The renderer owns the unsaved-changes state, so closing has to round-trip
  // through it: veto the close, ask, then close for real if it says so.
  win.on('close', (e) => {
    if (win.__forceClose) return;
    e.preventDefault();
    win.webContents.send('app:requestClose');
  });

  return win;
}

function focusedWindow() {
  return BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
}

function menuContext() {
  return {
    createWindow,
    openWithDialog,
    focusedWindow,
    setThemeSource,
    themeSource: () => nativeTheme.themeSource,
  };
}

function openWithDialog(win) {
  const result = dialog.showOpenDialogSync(win, {
    title: 'Open Spreadsheet',
    properties: ['openFile'],
    filters: SPREADSHEET_FILTERS,
  });
  if (!result || !result.length) return null;
  return result[0];
}

app.whenReady().then(() => {
  app.setName('Gridly');

  const saved = readPrefs().themeSource;
  nativeTheme.themeSource = ['light', 'dark', 'system'].includes(saved) ? saved : 'system';
  // Fires when the OS appearance flips while "system" is selected.
  nativeTheme.on('updated', broadcastTheme);

  Menu.setApplicationMenu(buildMenu(menuContext()));

  ready = true;
  if (pendingOpen.length) {
    pendingOpen.splice(0).forEach((p) => createWindow(p));
  } else {
    createWindow(null);
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow(null);
  });
});

app.on('open-file', (event, filePath) => {
  event.preventDefault();
  if (!ready) {
    pendingOpen.push(filePath);
    return;
  }
  const win = focusedWindow();
  // An untouched blank window adopts the file; otherwise it gets its own.
  if (win) {
    win.webContents.send('app:openPathIfEmpty', filePath);
  } else {
    createWindow(filePath);
  }
});

app.on('window-all-closed', () => {
  if (!isMac) app.quit();
});

// ------------------------------------------------------------------- IPC

ipcMain.handle('wb:open', async (event, explicitPath) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const filePath = explicitPath || openWithDialog(win);
  if (!filePath) return { canceled: true };
  try {
    const model = await readWorkbook(filePath);
    app.addRecentDocument(filePath);
    return { canceled: false, model };
  } catch (err) {
    return { canceled: false, error: describeError(err, filePath) };
  }
});

ipcMain.handle('wb:save', async (event, { model, saveAs }) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  let filePath = model.filePath;

  // Save As, first save, and any read-only source all need a destination.
  // Read-only sources (.xls, .xlsm) default to a .xlsx sibling.
  if (saveAs || !filePath || model.readOnly) {
    const base = filePath
      ? path.join(path.dirname(filePath), path.basename(filePath, path.extname(filePath)) + '.xlsx')
      : path.join(app.getPath('documents'), 'Untitled.xlsx');
    const chosen = dialog.showSaveDialogSync(win, {
      title: saveAs ? 'Save As' : 'Save Workbook',
      defaultPath: base,
      filters: [
        { name: 'Excel Workbook', extensions: ['xlsx'] },
        { name: 'Comma Separated Values', extensions: ['csv'] },
      ],
    });
    if (!chosen) return { canceled: true };
    filePath = chosen;
  }

  try {
    await writeWorkbook(model, filePath);
    app.addRecentDocument(filePath);
    return { canceled: false, filePath };
  } catch (err) {
    return { canceled: false, error: describeError(err, filePath) };
  }
});

ipcMain.handle('wb:exportCsvPath', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const chosen = dialog.showSaveDialogSync(win, {
    title: 'Export Sheet as CSV',
    defaultPath: path.join(app.getPath('documents'), 'Sheet.csv'),
    filters: [{ name: 'Comma Separated Values', extensions: ['csv'] }],
  });
  return chosen || null;
});

ipcMain.handle('app:confirmDiscard', async (event, name) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const { response } = await dialog.showMessageBox(win, {
    type: 'warning',
    buttons: ['Save', "Don't Save", 'Cancel'],
    defaultId: 0,
    cancelId: 2,
    message: `Do you want to save the changes you made to ${name || 'this workbook'}?`,
    detail: "Your changes will be lost if you don't save them.",
  });
  return ['save', 'discard', 'cancel'][response];
});

ipcMain.handle('app:messageBox', async (event, opts) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const { response } = await dialog.showMessageBox(win, Object.assign({ type: 'info' }, opts));
  return response;
});

ipcMain.on('app:closeWindow', (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  if (win) {
    win.__forceClose = true;
    win.close();
  }
});

ipcMain.on('app:newWindow', (event, filePath) => createWindow(filePath || null));

ipcMain.on('app:setTitle', (event, { title, filePath, edited }) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  if (!win) return;
  win.setTitle(title || 'Gridly');
  win.setDocumentEdited(!!edited);
  if (filePath && fs.existsSync(filePath)) win.setRepresentedFilename(filePath);
});

ipcMain.handle('app:clipboardRead', () => ({
  text: clipboard.readText(),
  html: clipboard.readHTML(),
}));

ipcMain.handle('app:clipboardWrite', (event, data) => {
  const payload = {};
  if (data && typeof data.text === 'string') payload.text = data.text;
  if (data && typeof data.html === 'string') payload.html = data.html;
  if (Object.keys(payload).length) clipboard.write(payload);
  return true;
});

// Printing renders an HTML table of the used range in an offscreen window and
// hands it to the system print dialog, which also covers "Save as PDF".
ipcMain.handle('app:print', async (event, html) => {
  const tmp = path.join(app.getPath('temp'), `gridly-print-${Date.now()}.html`);
  await fs.promises.writeFile(tmp, html, 'utf8');

  const printWin = new BrowserWindow({
    show: false,
    webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true },
  });

  try {
    await printWin.loadFile(tmp);
    await new Promise((resolve) => {
      printWin.webContents.print({ silent: false, printBackground: true }, () => resolve());
    });
    return { ok: true };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  } finally {
    if (!printWin.isDestroyed()) printWin.destroy();
    fs.promises.unlink(tmp).catch(() => {});
  }
});

ipcMain.on('app:openExternal', (event, url) => {
  if (/^https?:\/\//i.test(url) || /^mailto:/i.test(url)) shell.openExternal(url);
});

function describeError(err, filePath) {
  const msg = String((err && err.message) || err);
  if (/ENOENT/.test(msg)) return `The file could not be found:\n${filePath}`;
  if (/EACCES|EPERM/.test(msg)) return `Permission denied writing to:\n${filePath}`;
  if (/zip|End of (central directory|data)|corrupt/i.test(msg)) {
    return 'This file is not a valid Excel workbook, or it is damaged.';
  }
  if (/password|encrypt/i.test(msg)) return 'This workbook is password-protected, which Gridly cannot open.';
  return msg;
}
