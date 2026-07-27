'use strict';

const { app, Menu, shell } = require('electron');

// Menu items mostly forward a command name to the focused renderer, which
// owns all document state. Only window/app-level items act here.
function buildMenu({ createWindow, focusedWindow, setThemeSource, themeSource }) {
  const send = (cmd, arg) => () => {
    const win = focusedWindow();
    if (win) win.webContents.send('menu:command', { cmd, arg });
  };

  const item = (label, cmd, accelerator, extra) =>
    Object.assign({ label, accelerator, click: send(cmd) }, extra || {});

  const template = [
    {
      label: 'Gridly',
      submenu: [
        { role: 'about' },
        { type: 'separator' },
        { role: 'services' },
        { type: 'separator' },
        { role: 'hide' },
        { role: 'hideOthers' },
        { role: 'unhide' },
        { type: 'separator' },
        { role: 'quit' },
      ],
    },
    {
      label: 'File',
      submenu: [
        { label: 'New Workbook', accelerator: 'Cmd+N', click: () => createWindow(null) },
        item('Open…', 'open', 'Cmd+O'),
        { label: 'Open Recent', role: 'recentDocuments', submenu: [{ label: 'Clear Menu', role: 'clearRecentDocuments' }] },
        { type: 'separator' },
        item('Close Window', 'closeWindow', 'Cmd+W'),
        item('Save', 'save', 'Cmd+S'),
        item('Save As…', 'saveAs', 'Shift+Cmd+S'),
        { type: 'separator' },
        item('Export Sheet as CSV…', 'exportCsv'),
        { type: 'separator' },
        item('Print…', 'print', 'Cmd+P'),
      ],
    },
    {
      label: 'Edit',
      submenu: [
        item('Undo', 'undo', 'Cmd+Z'),
        item('Redo', 'redo', 'Shift+Cmd+Z'),
        { type: 'separator' },
        // Handled in the renderer rather than by the native roles: the grid is
        // a canvas with focus on a plain div, and Chromium only delivers
        // clipboard events to editable targets, so a role-driven paste would
        // silently do nothing. The renderer switches between cell-range and
        // text-field semantics based on what actually has focus.
        item('Cut', 'cut', 'Cmd+X'),
        item('Copy', 'copy', 'Cmd+C'),
        item('Paste', 'paste', 'Cmd+V'),
        item('Paste Values Only', 'pasteValues', 'Shift+Cmd+V'),
        // Deliberately no accelerator: a Delete accelerator would out-rank the
        // cell editor and eat backspaces while typing.
        item('Clear Contents', 'clear'),
        { type: 'separator' },
        item('Select All', 'selectAll', 'Cmd+A'),
        { type: 'separator' },
        item('Find…', 'find', 'Cmd+F'),
        item('Go To…', 'goto', 'Ctrl+G'),
      ],
    },
    {
      label: 'Insert',
      submenu: [
        item('Rows Above', 'insertRow'),
        item('Columns Left', 'insertCol'),
        { type: 'separator' },
        item('Delete Rows', 'deleteRow'),
        item('Delete Columns', 'deleteCol'),
        { type: 'separator' },
        item('New Sheet', 'newSheet'),
      ],
    },
    {
      label: 'Format',
      submenu: [
        item('Bold', 'bold', 'Cmd+B'),
        item('Italic', 'italic', 'Cmd+I'),
        item('Underline', 'underline', 'Cmd+U'),
        { type: 'separator' },
        {
          label: 'Number Format',
          submenu: [
            ['General', 'General'],
            ['Number', '#,##0.00'],
            ['Currency', '"$"#,##0.00'],
            ['Percent', '0.00%'],
            ['Date', 'yyyy-mm-dd'],
            ['Time', 'h:mm:ss AM/PM'],
            ['Scientific', '0.00E+00'],
            ['Text', '@'],
          ].map(([label, fmt]) => ({ label, click: send('numFmt', fmt) })),
        },
        { type: 'separator' },
        { label: 'Align Left', click: send('align', 'left') },
        { label: 'Align Center', click: send('align', 'center') },
        { label: 'Align Right', click: send('align', 'right') },
        { type: 'separator' },
        item('Wrap Text', 'wrap'),
        item('Merge Cells', 'merge'),
        { type: 'separator' },
        item('Freeze Panes at Selection', 'freeze'),
        item('Unfreeze Panes', 'unfreeze'),
        { type: 'separator' },
        item('Autofit Column Width', 'autofit'),
      ],
    },
    {
      label: 'View',
      submenu: [
        {
          label: 'Appearance',
          submenu: ['system', 'light', 'dark'].map((source) => ({
            label: { system: 'Match System', light: 'Light', dark: 'Dark' }[source],
            type: 'radio',
            checked: themeSource() === source,
            click: () => setThemeSource(source),
          })),
        },
        { type: 'separator' },
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' },
        { type: 'separator' },
        { role: 'togglefullscreen' },
        { type: 'separator' },
        { role: 'toggleDevTools' },
      ],
    },
    {
      label: 'Window',
      submenu: [{ role: 'minimize' }, { role: 'zoom' }, { type: 'separator' }, { role: 'front' }],
    },
    {
      label: 'Help',
      submenu: [
        { label: 'Keyboard Shortcuts', click: send('shortcuts') },
        {
          label: 'About Excel Compatibility',
          click: send('compat'),
        },
        { type: 'separator' },
        {
          label: 'Open Sample Workbook',
          click: send('sample'),
        },
      ],
    },
  ];

  return Menu.buildFromTemplate(template);
}

module.exports = { buildMenu };
