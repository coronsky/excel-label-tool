const { app, BrowserWindow, dialog, ipcMain, Menu, globalShortcut } = require('electron');
const path = require('path');

let mainWindow;

const menuTemplate = [
  {
    label: 'ファイル',
    submenu: [
      {
        label: 'Excelファイルを開く',
        accelerator: 'CmdOrCtrl+O',
        click: async () => {
          const result = await dialog.showOpenDialog(mainWindow, {
            properties: ['openFile'],
            filters: [{ name: 'Excelファイル', extensions: ['xlsx', 'xls'] }]
          });
          if (!result.canceled && result.filePaths.length > 0) {
            mainWindow.webContents.send('open-file-path', result.filePaths[0]);
          }
        }
      },
      { type: 'separator' },
      { label: '終了', accelerator: 'Alt+F4', role: 'quit' }
    ]
  },
  {
    label: '編集',
    submenu: [
      { label: 'コピー', accelerator: 'CmdOrCtrl+C', role: 'copy' },
      { label: '貼り付け', accelerator: 'CmdOrCtrl+V', role: 'paste' },
      { label: '全選択', accelerator: 'CmdOrCtrl+A', role: 'selectAll' }
    ]
  },
  {
    label: '表示',
    submenu: [
      { label: '再読み込み', accelerator: 'CmdOrCtrl+R', role: 'reload' },
      { label: '拡大', accelerator: 'CmdOrCtrl+=', role: 'zoomIn' },
      { label: '縮小', accelerator: 'CmdOrCtrl+-', role: 'zoomOut' },
      { label: '等倍', accelerator: 'CmdOrCtrl+0', role: 'resetZoom' },
      { type: 'separator' },
      { label: '全画面', accelerator: 'F11', role: 'togglefullscreen' }
    ]
  }
];

function createWindow() {
  Menu.setApplicationMenu(Menu.buildFromTemplate(menuTemplate));

  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 800,
    minHeight: 500,
    resizable: true,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    title: 'Excelラベル抽出ツール',
    backgroundColor: '#f0f2f5'
  });

  mainWindow.loadFile('index.html');
}

app.whenReady().then(() => {
  createWindow();
  // PhotoScape X へ次のテキストを貼り付けるグローバルホットキー
  globalShortcut.register('CommandOrControl+Shift+V', () => {
    if (mainWindow) mainWindow.webContents.send('global-paste');
  });
});

app.on('will-quit', () => {
  globalShortcut.unregisterAll();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

ipcMain.handle('open-file-dialog', async () => {
  return dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }]
  });
});
