'use strict';

const XLSX = require('xlsx');
const { ipcRenderer, clipboard } = require('electron');
const path = require('path');

// ─── State ────────────────────────────────────────────────────────────────────
let allData = [];
let shopAggregation = {};
let copyIndex = 0;
let currentNameCol = 1;
let currentQtyCol  = 7;   // H列
let currentShopCol = 8;   // I列
let storedWorkbook  = null;
let storedFileName  = '';
let availableColumns = [];

// ─── DOM refs ─────────────────────────────────────────────────────────────────
const dropZone          = document.getElementById('drop-zone');
const errorMsg          = document.getElementById('error-msg');
const content           = document.getElementById('content');
const fileInfoEl        = document.getElementById('file-info');
const dataGrid          = document.getElementById('data-grid');
const copyStatusEl      = document.getElementById('copy-status');
const seqCopyBtn        = document.getElementById('seq-copy-btn');
const resetCopyBtn      = document.getElementById('reset-copy-btn');
const shopSection       = document.getElementById('shop-section');
const shopListEl        = document.getElementById('shop-list');
const bulkCopyBtn       = document.getElementById('bulk-copy-btn');
const openFileBtn       = document.getElementById('open-file-btn');
const reloadBtn         = document.getElementById('reload-btn');
const colSelectorsGroup = document.getElementById('col-selectors-group');

// ─── マウスホイールでデータグリッドをスクロール ───────────────────────────────
document.getElementById('data-grid').addEventListener('wheel', (e) => {
  const wrapper = document.querySelector('.data-wrapper');
  wrapper.scrollTop += e.deltaY;
}, { passive: true });

// ─── Menu: ファイルを開く（main.js から送信） ─────────────────────────────────
ipcRenderer.on('open-file-path', (_, filePath) => {
  processFile(filePath, path.basename(filePath));
});

// ─── Drag & Drop ──────────────────────────────────────────────────────────────
document.addEventListener('dragover', (e) => e.preventDefault());
document.addEventListener('drop', (e) => {
  e.preventDefault();
  const files = Array.from(e.dataTransfer.files);
  const xl = files.find(f => /\.xlsx?$/i.test(f.name));
  if (xl) processFile(xl.path, xl.name);
  else showError('Excelファイル (.xlsx / .xls) をドロップしてください');
});

dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', () => dropZone.classList.remove('drag-over'));

// ─── File dialog ──────────────────────────────────────────────────────────────
openFileBtn.addEventListener('click', async () => {
  const result = await ipcRenderer.invoke('open-file-dialog');
  if (!result.canceled && result.filePaths.length > 0)
    processFile(result.filePaths[0], path.basename(result.filePaths[0]));
});

reloadBtn.addEventListener('click', async () => {
  const result = await ipcRenderer.invoke('open-file-dialog');
  if (!result.canceled && result.filePaths.length > 0)
    processFile(result.filePaths[0], path.basename(result.filePaths[0]));
});

// ─── File processing ──────────────────────────────────────────────────────────

function processFile(filePath, fileName) {
  try {
    hideError();
    storedWorkbook = XLSX.readFile(filePath);
    storedFileName = fileName;

    availableColumns = discoverColumns(storedWorkbook);
    currentNameCol   = autoDetectCol(storedWorkbook, ['ラベル名', '名前', 'お名前', 'おなまえ', 'name', '氏名'], 1);
    currentQtyCol    = autoDetectCol(storedWorkbook, ['枚数', '数量', '枚', 'qty', 'quantity'], 7);
    currentShopCol   = autoDetectCol(storedWorkbook, ['ショップ', '店舗', '店', 'shop'], 8);

    dropZone.style.display = 'none';
    reloadBtn.style.display = '';
    content.style.display = 'flex';
    colSelectorsGroup.style.display = '';

    nameSelector.updateUI();
    qtySelector.updateUI();
    shopSelector.updateUI();
    reprocess();
  } catch (err) {
    showError(`読み込み失敗: ${err.message}`);
    console.error(err);
  }
}

function reprocess() {
  allData = [];
  shopAggregation = {};
  copyIndex = 0;

  storedWorkbook.SheetNames.forEach(name => {
    processSheet(storedWorkbook.Sheets[name], name);
  });

  fileInfoEl.textContent =
    `${storedFileName} — ${storedWorkbook.SheetNames.length}シート / ${allData.length}件`;

  renderData();
  renderShopAggregation();
  updateCopyStatus();
}

function processSheet(sheet, sheetName) {
  if (!sheet['!ref']) return;
  const range = XLSX.utils.decode_range(sheet['!ref']);

  for (let r = 3; r <= range.e.r; r++) {
    const nameCell = sheet[XLSX.utils.encode_cell({ r, c: currentNameCol })];
    if (!nameCell || nameCell.v === undefined || nameCell.v === null) continue;

    const original = String(nameCell.v).trim();
    if (!original) continue;

    const extracted = cleanName(original);

    const qtyCell  = sheet[XLSX.utils.encode_cell({ r, c: currentQtyCol })];
    const shopCell = sheet[XLSX.utils.encode_cell({ r, c: currentShopCol })];
    const qty  = qtyCell  && qtyCell.v ? Math.max(1, Math.round(Number(qtyCell.v) || 1)) : 1;
    const shop = shopCell && shopCell.v ? String(shopCell.v).trim() : '';

    allData.push({ sheet: sheetName, original, extracted });
    if (shop) shopAggregation[shop] = (shopAggregation[shop] || 0) + qty;
  }
}

// ─── Column discovery ─────────────────────────────────────────────────────────

function discoverColumns(wb) {
  const colMap = new Map();
  wb.SheetNames.forEach(sheetName => {
    const sheet = wb.Sheets[sheetName];
    if (!sheet['!ref']) return;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let c = range.s.c; c <= range.e.c; c++) {
      if (colMap.has(c)) continue;
      let header = '';
      for (let r = 0; r <= Math.min(2, range.e.r); r++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c })];
        if (cell && cell.v) { header = String(cell.v).trim(); break; }
      }
      colMap.set(c, { index: c, letter: colIndexToLetter(c), header });
    }
  });
  return Array.from(colMap.values()).sort((a, b) => a.index - b.index);
}

function autoDetectCol(wb, keys, defaultCol) {
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    if (!sheet['!ref']) continue;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let r = 0; r <= Math.min(2, range.e.r); r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c })];
        if (cell && cell.v) {
          const v = String(cell.v).toLowerCase();
          if (keys.some(k => v.includes(k.toLowerCase()))) return c;
        }
      }
    }
  }
  return defaultCol;
}

function colIndexToLetter(index) {
  let letter = '';
  let n = index;
  do {
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return letter;
}

// ─── Column selector factory ──────────────────────────────────────────────────
// 3つの列セレクターを同じロジックで生成する

function createColSelector(btnEl, dropdownEl, getCol, setCol, labelPrefix) {
  function updateUI() {
    const col = availableColumns.find(c => c.index === getCol());
    const label = col
      ? `${col.letter}列 — ${col.header || '（ヘッダーなし）'}`
      : `${colIndexToLetter(getCol())}列`;
    btnEl.textContent = `${labelPrefix}: ${label} ▼`;
  }

  function buildDropdown() {
    dropdownEl.innerHTML = '';
    availableColumns.forEach(col => {
      const isActive = col.index === getCol();
      const item = document.createElement('button');
      item.className = 'col-dropdown-item' + (isActive ? ' active' : '');
      item.innerHTML =
        `<span class="col-letter">${col.letter}</span>` +
        `<span class="col-header-label">${esc(col.header || '（ヘッダーなし）')}</span>` +
        (isActive ? '<span class="col-check">✓</span>' : '');

      item.addEventListener('click', () => {
        setCol(col.index);
        closeAllDropdowns();
        updateUI();
        reprocess();
      });
      dropdownEl.appendChild(item);
    });
  }

  btnEl.addEventListener('click', (e) => {
    e.stopPropagation();
    const isOpen = dropdownEl.style.display !== 'none';
    closeAllDropdowns();
    if (!isOpen) {
      dropdownEl.style.display = 'block';
      buildDropdown();
    }
  });

  dropdownEl.addEventListener('click', (e) => e.stopPropagation());

  return { updateUI };
}

function closeAllDropdowns() {
  document.querySelectorAll('.col-dropdown').forEach(d => { d.style.display = 'none'; });
}

document.addEventListener('click', closeAllDropdowns);

// セレクターを初期化（DOM構築後に実行）
const nameSelector = createColSelector(
  document.getElementById('col-select-btn'),
  document.getElementById('col-dropdown'),
  () => currentNameCol,
  (v) => { currentNameCol = v; },
  '抽出列'
);

const qtySelector = createColSelector(
  document.getElementById('qty-col-select-btn'),
  document.getElementById('qty-col-dropdown'),
  () => currentQtyCol,
  (v) => { currentQtyCol = v; },
  '枚数列'
);

const shopSelector = createColSelector(
  document.getElementById('shop-col-select-btn'),
  document.getElementById('shop-col-dropdown'),
  () => currentShopCol,
  (v) => { currentShopCol = v; },
  'ショップ列'
);

// ─── Name cleaning ────────────────────────────────────────────────────────────

function cleanName(name) {
  let s = String(name);

  s = s.replace(/（[^）]*）/g, '');
  s = s.replace(/\([^)]*\)/g, '');
  s = s.replace(/【[^】]*】/g, '');
  s = s.replace(/\[[^\]]*\]/g, '');
  s = s.replace(/「[^」]*」/g, '');
  s = s.replace(/〔[^〕]*〕/g, '');
  s = s.replace(/〈[^〉]*〉/g, '');

  const keywords = ['キラキラ', 'ハート', 'スター', 'リボン', '肉球', 'ほし', '星', '花'];

  let prev;
  do {
    prev = s;
    for (const kw of keywords) {
      s = s.replace(new RegExp('[\\s　・×,、。]*' + kw + '[\\s　]*$'), '');
      s = s.replace(
        new RegExp('(^|[\\s　・,、。])' + kw + '($|[\\s　・,、。])', 'g'),
        '$1$2'
      );
    }
  } while (s !== prev);

  s = s.replace(/[\s　]+/g, ' ').trim();
  return s;
}

// ─── Rendering ────────────────────────────────────────────────────────────────

function renderData() {
  dataGrid.innerHTML = '';

  const hExt = document.createElement('div');
  hExt.className = 'grid-cell grid-col-header';
  hExt.textContent = '抽出後（クリックでコピー）';

  const hOrig = document.createElement('div');
  hOrig.className = 'grid-cell grid-col-header';
  hOrig.textContent = 'オリジナル';

  dataGrid.appendChild(hExt);
  dataGrid.appendChild(hOrig);

  allData.forEach((item, index) => {
    const extBtn = document.createElement('button');
    extBtn.className = 'grid-cell extracted-btn';
    extBtn.dataset.index = index;
    extBtn.innerHTML =
      `<span class="row-num">${index + 1}</span>` +
      `<span class="cell-text">${esc(item.extracted)}</span>` +
      `<span class="copy-hint">コピー</span>`;

    extBtn.addEventListener('click', () => {
      clipboard.writeText(item.extracted);
      flash(extBtn, 'flash-click');
    });

    const origCell = document.createElement('div');
    origCell.className = 'grid-cell original-cell';
    origCell.innerHTML =
      `<span class="row-num">${index + 1}</span>` +
      `<span class="cell-text">${esc(item.original)}</span>`;

    dataGrid.appendChild(extBtn);
    dataGrid.appendChild(origCell);
  });
}

function renderShopAggregation() {
  shopListEl.innerHTML = '';
  const entries = Object.entries(shopAggregation);

  if (entries.length === 0) {
    shopSection.style.display = 'none';
    return;
  }

  // ショップ列のヘッダー名をタイトルに反映
  const shopColInfo = availableColumns.find(c => c.index === currentShopCol);
  const qtyColInfo  = availableColumns.find(c => c.index === currentQtyCol);
  const shopLabel   = shopColInfo && shopColInfo.header ? shopColInfo.header : `${colIndexToLetter(currentShopCol)}列`;
  const qtyLabel    = qtyColInfo  && qtyColInfo.header  ? qtyColInfo.header  : `${colIndexToLetter(currentQtyCol)}列`;
  document.querySelector('.shop-title').textContent =
    `集計（${qtyLabel} × ${shopLabel}）`;

  shopSection.style.display = 'flex';
  const sorted = entries.sort(([a], [b]) => a.localeCompare(b, 'ja'));
  const total = sorted.reduce((sum, [, qty]) => sum + qty, 0);

  sorted.forEach(([shop, qty]) => {
    const row = document.createElement('div');
    row.className = 'shop-row';
    row.innerHTML =
      `<span class="shop-name">${esc(shop)}</span>` +
      `<span class="shop-qty">${qty}</span>`;
    shopListEl.appendChild(row);
  });

  // 合計をスクロール外の固定エリアに表示
  const shopTotal = document.getElementById('shop-total');
  document.getElementById('shop-total-qty').textContent = total;
  shopTotal.style.display = '';
}

// ─── PhotoScape X 貼り付け ────────────────────────────────────────────────────

const { exec } = require('child_process');
let psCounting = false;

function pasteToPhotoScapeX(text) {
  clipboard.writeText(text);

  const psLines = [
    '$w = New-Object -ComObject wscript.shell',
    'if ($w.AppActivate("PhotoScape X")) {',
    '  Start-Sleep -Milliseconds 300',
    '  $w.SendKeys("^a")',
    '  Start-Sleep -Milliseconds 150',
    '  $w.SendKeys("^v")',
    '  exit 0',
    '} else { exit 1 }'
  ].join('\n');

  const encoded = Buffer.from(psLines, 'utf16le').toString('base64');
  const btn = document.getElementById('ps-paste-btn');

  exec(`powershell -NoProfile -EncodedCommand ${encoded}`, (err) => {
    psCounting = false;
    btn.disabled = false;
    if (err && err.code === 1) {
      showError('PhotoScape X が見つかりません。起動してテキストを編集モードにしてください。');
      btn.textContent = '× 見つかりません';
    } else {
      btn.textContent = '✓ 貼り付けました';
    }
    setTimeout(() => { btn.textContent = '▶ PhotoScapeに貼り付け'; }, 2000);
  });
}

document.getElementById('ps-paste-btn').addEventListener('click', () => {
  if (allData.length === 0 || psCounting) return;
  if (copyIndex >= allData.length) copyIndex = 0;

  const item = allData[copyIndex];
  clipboard.writeText(item.extracted);

  // 行をハイライト
  dataGrid.querySelectorAll('.extracted-btn').forEach(b => b.classList.remove('seq-active'));
  const activeBtn = dataGrid.querySelector(`[data-index="${copyIndex}"]`);
  if (activeBtn) {
    activeBtn.classList.add('seq-active');
    activeBtn.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  copyIndex++;
  updateCopyStatus();

  // 3秒カウントダウン後に自動貼り付け
  psCounting = true;
  const btn = document.getElementById('ps-paste-btn');
  btn.disabled = true;
  let count = 3;
  btn.textContent = `PhotoScapeへ切替えて... ${count}`;

  const timer = setInterval(() => {
    count--;
    if (count > 0) {
      btn.textContent = `PhotoScapeへ切替えて... ${count}`;
    } else {
      clearInterval(timer);
      btn.textContent = '貼り付け中...';
      pasteToPhotoScapeX(item.extracted);
    }
  }, 1000);
});

// ─── Copy actions ─────────────────────────────────────────────────────────────

seqCopyBtn.addEventListener('click', () => {
  if (allData.length === 0) return;
  if (copyIndex >= allData.length) copyIndex = 0;

  clipboard.writeText(allData[copyIndex].extracted);

  dataGrid.querySelectorAll('.extracted-btn').forEach(b => b.classList.remove('seq-active'));
  const activeBtn = dataGrid.querySelector(`[data-index="${copyIndex}"]`);
  if (activeBtn) {
    activeBtn.classList.add('seq-active');
    activeBtn.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    flash(activeBtn, 'flash-seq');
  }

  copyIndex++;
  updateCopyStatus();
});

resetCopyBtn.addEventListener('click', () => {
  copyIndex = 0;
  dataGrid.querySelectorAll('.extracted-btn').forEach(b => b.classList.remove('seq-active'));
  updateCopyStatus();
});

bulkCopyBtn.addEventListener('click', () => {
  const sorted = Object.entries(shopAggregation).sort(([a], [b]) => a.localeCompare(b, 'ja'));
  clipboard.writeText(sorted.map(([, qty]) => qty).join('\n'));

  const orig = bulkCopyBtn.textContent;
  bulkCopyBtn.textContent = '✓ コピーしました！';
  bulkCopyBtn.disabled = true;
  setTimeout(() => {
    bulkCopyBtn.textContent = orig;
    bulkCopyBtn.disabled = false;
  }, 1500);
});

// ─── Helpers ──────────────────────────────────────────────────────────────────

function updateCopyStatus() {
  copyStatusEl.textContent = `${copyIndex} / ${allData.length} コピー済み`;
}

function flash(el, cls) {
  el.classList.remove('flash-click', 'flash-seq');
  void el.offsetWidth;
  el.classList.add(cls);
  setTimeout(() => el.classList.remove(cls), 600);
}

function showError(msg) {
  errorMsg.textContent = msg;
  errorMsg.style.display = 'block';
  setTimeout(() => { errorMsg.style.display = 'none'; }, 4000);
}

function hideError() {
  errorMsg.style.display = 'none';
}

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
