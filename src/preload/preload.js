'use strict';

const { contextBridge, ipcRenderer } = require('electron');

// The renderer runs with contextIsolation and no Node access. Everything it
// needs from the OS crosses this bridge as plain data.
contextBridge.exposeInMainWorld('gridly', {
  openWorkbook: (filePath) => ipcRenderer.invoke('wb:open', filePath || null),
  saveWorkbook: (model, saveAs) => ipcRenderer.invoke('wb:save', { model, saveAs: !!saveAs }),
  chooseCsvPath: () => ipcRenderer.invoke('wb:exportCsvPath'),
  confirmDiscard: (name) => ipcRenderer.invoke('app:confirmDiscard', name),
  messageBox: (opts) => ipcRenderer.invoke('app:messageBox', opts),
  print: (html) => ipcRenderer.invoke('app:print', html),

  closeWindow: () => ipcRenderer.send('app:closeWindow'),
  newWindow: (filePath) => ipcRenderer.send('app:newWindow', filePath || null),
  setTitle: (info) => ipcRenderer.send('app:setTitle', info),
  openExternal: (url) => ipcRenderer.send('app:openExternal', url),

  // A sandboxed preload cannot reach Electron's clipboard module, so both
  // directions go through the main process.
  readClipboard: () => ipcRenderer.invoke('app:clipboardRead'),
  writeClipboard: (data) => ipcRenderer.invoke('app:clipboardWrite', data),

  onMenuCommand: (fn) => ipcRenderer.on('menu:command', (_e, payload) => fn(payload)),
  onOpenPath: (fn) => ipcRenderer.on('app:openPath', (_e, p) => fn(p)),
  onOpenPathIfEmpty: (fn) => ipcRenderer.on('app:openPathIfEmpty', (_e, p) => fn(p)),
  onRequestClose: (fn) => ipcRenderer.on('app:requestClose', () => fn()),
  onTheme: (fn) => ipcRenderer.on('app:theme', (_e, theme) => fn(theme)),
});
