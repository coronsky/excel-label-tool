'use strict';

const XLSX = require('xlsx');
const { ipcRenderer, clipboard } = require('electron');
const path = require('path');

// ─── State ────────────────────────────────────────────────────────────────────
let allData = [];
let shopAggregation = {};
let copyIndex = 0;
let currentNameCol  = 1;
let storedWorkbook  = null;
let storedFileName  = '';
let availableColumns = [];   // [{ index, letter, header }]

// ─── DOM refs ─────────────────────────────────────────────────────────────────
const dropZone     = document.getElementById('drop-zone');
const errorMsg     = document.getElementById('error-msg');
const content      = document.getElementById('content');
const fileInfoEl   = document.getElementById('file-info');
const dataGrid     = document.getElementById('data-grid');
const copyStatusEl = document.getElementById('copy-status');
const seqCopyBtn   = document.getElementById('seq-copy-btn');
const resetCopyBtn = document.getElementById('reset-copy-btn');
const shopSection  = document.getElementById('shop-section');
const shopListEl   = document.getElementById('shop-list');
const bulkCopyBtn  = document.getElementById('bulk-copy-btn');
const openFileBtn  = document.getElementById('open-file-btn');
const reloadBtn    = document.getElementById('reload-btn');
const colSelector  = document.getElementById('col-selector');
const colSelectBtn = document.getElementById('col-select-btn');
const colDropdown  = document.getElementById('col-dropdown');

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

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', () => dropZone.classList.remove('drag-over'));

// ─── File dialog ──────────────────────────────────────────────────────────────

openFileBtn.addEventListener('click', async () => {
  const result = await ipcRenderer.invoke('open-file-dialog');
  if (!result.canceled && result.filePaths.length > 0) {
    processFile(result.filePaths[0], path.basename(result.filePaths[0]));
  }
});

reloadBtn.addEventListener('click', async () => {
  const result = await ipcRenderer.invoke('open-file-dialog');
  if (!result.canceled && result.filePaths.length > 0) {
    processFile(result.filePaths[0], path.basename(result.filePaths[0]));
  }
});

// ─── File processing ──────────────────────────────────────────────────────────

function processFile(filePath, fileName) {
  try {
    hideError();
    storedWorkbook = XLSX.readFile(filePath);
    storedFileName = fileName;

    availableColumns = discoverColumns(storedWorkbook);
    currentNameCol   = autoDetectNameCol(storedWorkbook);

    dropZone.style.display = 'none';
    reloadBtn.style.display = '';
    content.style.display = 'block';
    colSelector.style.display = '';

    updateColSelectorUI();
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
    processSheet(storedWorkbook.Sheets[name], name, currentNameCol);
  });

  fileInfoEl.textContent =
    `${storedFileName} — ${storedWorkbook.SheetNames.length}シート / ${allData.length}件`;

  renderData();
  renderShopAggregation();
  updateCopyStatus();
}

function processSheet(sheet, sheetName, nameCol) {
  if (!sheet['!ref']) return;
  const range = XLSX.utils.decode_range(sheet['!ref']);

  for (let r = 3; r <= range.e.r; r++) {
    const nameCell = sheet[XLSX.utils.encode_cell({ r, c: nameCol })];
    if (!nameCell || nameCell.v === undefined || nameCell.v === null) continue;

    const original = String(nameCell.v).trim();
    if (!original) continue;

    const extracted = cleanName(original);

    // H col = index 7 (枚数), I col = index 8 (ショップ)
    const qtyCell  = sheet[XLSX.utils.encode_cell({ r, c: 7 })];
    const shopCell = sheet[XLSX.utils.encode_cell({ r, c: 8 })];
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

function autoDetectNameCol(wb) {
  const keys = ['ラベル名', '名前', 'お名前', 'おなまえ', 'name', '氏名'];
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
  return 1;
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

// ─── Column selector UI ───────────────────────────────────────────────────────

function updateColSelectorUI() {
  const col = availableColumns.find(c => c.index === currentNameCol);
  const label = col
    ? `${col.letter}列 — ${col.header || '（ヘッダーなし）'}`
    : `${colIndexToLetter(currentNameCol)}列`;
  colSelectBtn.textContent = `抽出列: ${label} ▼`;
}

colSelectBtn.addEventListener('click', (e) => {
  e.stopPropagation();
  const isOpen = colDropdown.style.display !== 'none';
  colDropdown.style.display = isOpen ? 'none' : 'block';
  if (!isOpen) buildColDropdown();
});

document.addEventListener('click', () => {
  colDropdown.style.display = 'none';
});

colDropdown.addEventListener('click', (e) => e.stopPropagation());

function buildColDropdown() {
  colDropdown.innerHTML = '';
  availableColumns.forEach(col => {
    const isActive = col.index === currentNameCol;
    const item = document.createElement('button');
    item.className = 'col-dropdown-item' + (isActive ? ' active' : '');
    item.innerHTML =
      `<span class="col-letter">${col.letter}</span>` +
      `<span class="col-header-label">${esc(col.header || '（ヘッダーなし）')}</span>` +
      (isActive ? '<span class="col-check">✓</span>' : '');

    item.addEventListener('click', () => {
      currentNameCol = col.index;
      colDropdown.style.display = 'none';
      updateColSelectorUI();
      reprocess();
    });

    colDropdown.appendChild(item);
  });
}

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

  const hOrig = document.createElement('div');
  hOrig.className = 'grid-cell grid-col-header';
  hOrig.textContent = 'オリジナル';

  const hExt = document.createElement('div');
  hExt.className = 'grid-cell grid-col-header';
  hExt.textContent = '抽出後（クリックでコピー）';

  dataGrid.appendChild(hOrig);
  dataGrid.appendChild(hExt);

  allData.forEach((item, index) => {
    const origCell = document.createElement('div');
    origCell.className = 'grid-cell original-cell';
    origCell.innerHTML =
      `<span class="row-num">${index + 1}</span>` +
      `<span class="cell-text">${esc(item.original)}</span>`;

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

    dataGrid.appendChild(origCell);
    dataGrid.appendChild(extBtn);
  });
}

function renderShopAggregation() {
  shopListEl.innerHTML = '';
  const entries = Object.entries(shopAggregation);

  if (entries.length === 0) {
    shopSection.style.display = 'none';
    return;
  }

  shopSection.style.display = 'block';
  const sorted = entries.sort(([a], [b]) => a.localeCompare(b, 'ja'));

  sorted.forEach(([shop, qty]) => {
    const row = document.createElement('div');
    row.className = 'shop-row';
    row.innerHTML =
      `<span class="shop-name">${esc(shop)}</span>` +
      `<span class="shop-qty">${qty}</span>`;
    shopListEl.appendChild(row);
  });
}

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
