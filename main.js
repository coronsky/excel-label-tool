const { app, BrowserWindow, dialog, ipcMain, Menu, globalShortcut } = require('electron');
const path = require('path');
const fs   = require('fs');

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

ipcMain.handle('select-output-dir', async () => {
  return dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory'],
    title: 'PNG保存フォルダを選択'
  });
});

ipcMain.handle('select-paw-image', async () => {
  return dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [{ name: '画像ファイル', extensions: ['png', 'jpg', 'jpeg'] }],
    title: '肉球画像を選択'
  });
});

// ─── PNG Label Generator ──────────────────────────────────────────────────────

function resolveColor(color) {
  const map = {
    '黒': '#1a1008', '白': '#ffffff', '金': '#c8a000',
    '赤': '#cc0000', '青': '#0044aa', 'ピンク': '#ff69b4',
    '茶': '#5c2e00', '銀': '#888888', '緑': '#006600',
  };
  const c = (color || '黒').trim();
  return map[c] || (c.startsWith('#') ? c : '#1a1008');
}

function pawFilter(color) {
  const c = (color || '').trim();
  if (c === '白') return 'brightness(0) invert(1)';
  if (c === '金') return 'brightness(0) sepia(1) saturate(6) hue-rotate(10deg)';
  if (c === '銀') return 'brightness(0) invert(0.75)';
  if (c === '赤') return 'brightness(0) sepia(1) saturate(10) hue-rotate(340deg)';
  if (c === '青') return 'brightness(0) sepia(1) saturate(10) hue-rotate(200deg)';
  return 'none';
}

function buildLabelHTML({ name, color, font, deco }, { width, height, pawBase64 }) {
  const cssColor   = resolveColor(color);
  const filter     = pawFilter(color);
  const pawSrc     = pawBase64 ? `data:image/png;base64,${pawBase64}` : '';
  const decoH      = Math.round(height * 0.78);
  const initFontSz = Math.round(height * 0.68);
  const gap        = Math.round(height * 0.07);
  const d          = (deco || '').trim();

  const pawTag   = (side) => pawSrc
    ? `<img class="paw" src="${pawSrc}" data-side="${side}">`  : '';
  const charTag  = (ch) => `<span class="deco-char">${ch}</span>`;

  let L = '', R = '';
  if      (d === '肉球左右') { L = pawTag('L');  R = pawTag('R'); }
  else if (d === '肉球右')   { R = pawTag('R'); }
  else if (d === '肉球左')   { L = pawTag('L'); }
  else if (d === '星左右')   { L = charTag('★'); R = charTag('★'); }
  else if (d === 'ハート左右'){ L = charTag('❤'); R = charTag('❤'); }

  const esc   = s => s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const fam   = font && font.includes('行楷')
    ? '"AR P行楷書体L"'
    : '"MS PGothic","ＭＳ Ｐゴシック"';

  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
*{margin:0;padding:0;box-sizing:border-box}
html,body{width:${width}px;height:${height}px;overflow:hidden;background:#fff}
body{display:flex;align-items:center;justify-content:center;
     gap:${gap}px;padding:${Math.round(height*.05)}px ${Math.round(height*.08)}px}
.paw{height:${decoH}px;width:auto;flex-shrink:0;filter:${filter}}
.deco-char{font-size:${Math.round(height*.60)}px;color:${cssColor};line-height:1;flex-shrink:0}
#n{font-family:${fam},sans-serif;font-size:${initFontSz}px;
   color:${cssColor};white-space:nowrap;line-height:1.1;flex-shrink:1}
</style></head><body>
${L}<span id="n">${esc(name)}</span>${R}
<script>
(function(){
  var el=document.getElementById('n'),sz=${initFontSz};
  while(document.body.scrollWidth>${width}-1&&sz>14){sz-=2;el.style.fontSize=sz+'px';}
  document.title='done';
})();
</script>
</body></html>`;
}

ipcMain.handle('generate-labels', async (event, { labels, settings }) => {
  const { outputDir, width, height, pawImagePath } = settings;

  let pawBase64 = '';
  try {
    if (pawImagePath && fs.existsSync(pawImagePath))
      pawBase64 = fs.readFileSync(pawImagePath).toString('base64');
  } catch(e) { console.error(e); }

  const win = new BrowserWindow({
    width: width + 2, height: height + 2,
    show: false, frame: false,
    webPreferences: { nodeIntegration: false, contextIsolation: true }
  });

  let count = 0;
  const errors = [];

  for (const label of labels) {
    try {
      const html = buildLabelHTML(label, { width, height, pawBase64 });
      await win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));

      // wait for inline JS auto-resize to finish
      await new Promise(resolve => {
        const t0 = Date.now();
        const iv = setInterval(async () => {
          try {
            const t = await win.webContents.executeJavaScript('document.title');
            if (t === 'done' || Date.now() - t0 > 1500) { clearInterval(iv); resolve(); }
          } catch { clearInterval(iv); resolve(); }
        }, 30);
      });

      const img  = await win.webContents.capturePage({ x:0, y:0, width, height });
      const buf  = img.toPNG();
      const safe = label.name.replace(/[\\/:*?"<>|]/g, '_');
      const out  = path.join(outputDir, `${String(count+1).padStart(3,'0')}_${safe}.png`);
      fs.writeFileSync(out, buf);
      count++;
      mainWindow.webContents.send('label-gen-progress', { current: count, total: labels.length });
    } catch(e) {
      errors.push(`${label.name}: ${e.message}`);
    }
  }

  win.destroy();
  return { success: true, count, outputDir, errors };
});
