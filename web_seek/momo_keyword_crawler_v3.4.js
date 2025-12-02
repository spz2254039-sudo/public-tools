// ==UserScript==
// @name         momo 關鍵字搜尋結果抓取器 V3.4（搜尋版+quota 對齊史料寫入）
// @namespace    http://tampermonkey.net/
// @version      3.4
// @description  在 momo「搜尋結果 / 列表頁」抓取商品網址與名稱，可載入 TXT 關鍵字規則與歷史 URL.csv 排除清單，支援自動跨頁直到達到「目標最多 N 筆」，去重後可分別下載「序號+商品網址」與「序號+商品名稱」CSV，並可選擇是否使用 / 寫入 localStorage 歷史網址。
// @match        https://www.momoshop.com.tw/*
// @grant        GM_addStyle
// ==/UserScript==

(function () {
  'use strict';

  // ============================
  //  全域常數與變數
  // ============================
  const LS_KEY_HISTORY = 'momo_search_seen_urls_v1'; // localStorage 歷史網址

  let includeKeywords = [];
  let excludeKeywords = [];
  let rulesLoaded = false;

  // 本輪（自腳本啟動以來）抓到的資料，用於匯出 CSV
  /** @type {{url: string, title: string}[]} */
  let records = [];

  // 匯出上限（0 表示不限，>0 表示最多 N 筆）
  let maxExportLimit = 0;

  // 歷史網址集合（localStorage + 本輪新增）
  /** @type {Set<string>} */
  let historyUrlSet = new Set();

  // 外部匯入歷史 URL.csv（排除用，不寫回 localStorage）
  /** @type {Set<string>} */
  let externalExcludeSet = new Set();

  // 方便顯示「最近一頁 已抓取 / 不重複」統計
  let lastGrabTotal = 0;
  let lastGrabUnique = 0;

  // 自動跨頁抓取中的旗標
  let isAutoPaging = false;

  const settings = {
    excludeTpUrls: false,
  };

  // ============================
  //  啟動：插入右下角按鈕
  // ============================
  function init() {
    if (document.readyState === 'complete' || document.readyState === 'interactive') {
      setTimeout(createFloatingButton, 1200);
    } else {
      window.addEventListener('DOMContentLoaded', () => setTimeout(createFloatingButton, 1200));
    }
  }

  function createFloatingButton() {
    if (document.getElementById('momo-toggle-btn')) return;

    const btn = document.createElement('button');
    btn.id = 'momo-toggle-btn';
    btn.textContent = 'momo 搜尋抓取';

    GM_addStyle(`
      #momo-toggle-btn {
        position: fixed;
        right: 16px;
        bottom: 16px;
        z-index: 999999;
        padding: 10px 14px;
        font-size: 13px;
        background: #ff4081;
        color: #fff;
        border: none;
        border-radius: 6px;
        cursor: pointer;
        box-shadow: 0 2px 6px rgba(0,0,0,0.2);
        user-select: none;
      }
      #momo-toggle-btn:hover {
        opacity: 0.92;
      }
    `);

    document.body.appendChild(btn);
    makeFloatingButtonDraggable(btn);

    btn.addEventListener('click', (e) => {
      if (btn._dragMoved) {
        btn._dragMoved = false;
        e.preventDefault();
        e.stopPropagation();
        return;
      }
      openPanel();
    });
  }

  function makeFloatingButtonDraggable(btn) {
    let isDragging = false;
    let startX = 0, startY = 0;
    let startLeft = 0, startTop = 0;
    let moved = false;

    btn.addEventListener('mousedown', (e) => {
      isDragging = true;
      moved = false;
      btn._dragMoved = false;

      startX = e.clientX;
      startY = e.clientY;
      const rect = btn.getBoundingClientRect();
      startLeft = rect.left;
      startTop = rect.top;

      e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
      if (!isDragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;

      if (Math.abs(dx) > 3 || Math.abs(dy) > 3) {
        moved = true;
      }

      btn.style.left = (startLeft + dx) + 'px';
      btn.style.top = (startTop + dy) + 'px';
      btn.style.right = 'auto';
      btn.style.bottom = 'auto';
    });

    document.addEventListener('mouseup', () => {
      if (isDragging && moved) {
        btn._dragMoved = true;
      }
      isDragging = false;
    });
  }

  // ============================
  //  控制面板
  // ============================
  function openPanel() {
    let panel = document.getElementById('momo-panel');
    if (panel) {
      panel.style.display = 'block';
      updateHistorySummary();
      updateCountStatus();
      return;
    }
    panel = createPanel();
    panel.style.display = 'block';
  }

  function createPanel() {
    let panel = document.getElementById('momo-panel');
    if (panel) return panel;

    // 先載入歷史網址
    historyUrlSet = loadHistoryFromLocalStorage();

    panel = document.createElement('div');
    panel.id = 'momo-panel';
    panel.innerHTML = `
      <div class="momo-header">
        <span>momo 關鍵字搜尋結果抓取器 （搜尋版）</span>
        <button type="button" id="momo-close-btn">×</button>
      </div>
      <div class="momo-body">
        <!-- 1. 抓取設定 -->
        <section class="momo-section">
          <h3>1. 抓取設定（匯入 TXT 規則 / 歷史 URL）</h3>
          <div class="momo-row">
            <span>規則文字檔（+包含 / -排除）：</span>
            <input type="file" id="momo-rules-file" accept=".txt" />
          </div>
          <div id="momo-rules-summary" class="momo-summary">
            尚未載入規則（未載入則不做關鍵字過濾，只做去重）。
          </div>

          <div class="momo-row" style="margin-top:6px;">
            <span>匯入歷史 URL.csv（排除清單，可選）：</span>
            <input type="file" id="momo-history-file" accept=".csv" />
          </div>
          <div id="momo-history-file-summary" class="momo-summary">
            尚未匯入外部歷史 URL（如需排除既有清單，可載入「序號,商品網址」格式的 CSV）。
          </div>
        </section>

        <!-- 2. 抓取目前頁面商品 -->
        <section class="momo-section">
          <h3>2. 抓取目前頁面商品（JSON-LD 擷取）</h3>
          <div class="momo-row-inline">
            <button type="button" id="momo-grab-btn">抓取目前頁面商品</button>
          </div>
          <div class="momo-row-inline" style="margin-top:6px;">
            <span class="momo-inline-label">目標總筆數（最多）：</span>
            <input type="text" id="momo-target-count" class="momo-target-input" placeholder="例如 300，空白或 0 表示不限" />
            <button type="button" id="momo-auto-grab-btn">自動跨頁抓取至目標</button>
          </div>
          <div class="momo-row-inline">
            <label class="momo-inline-label">
              <input type="checkbox" id="momo-exclude-tp" />
              排除 TP 型 momo 商品網址（/TP/.../goodsDetail/...）
            </label>
          </div>
          <div id="momo-count-status" class="momo-count-status">
            最近一頁：尚未抓取。 本次腳本累積可匯出：0 筆（匯出上限：不限）。
          </div>
          <div class="momo-small-tip">
            小提醒：請在「搜尋結果 / 列表頁」使用，本工具會從頁面中的 JSON-LD 結構（ItemList）抓取商品資訊；自動跨頁會嘗試點擊頁面右下角的「下一頁」按鈕。
          </div>
        </section>

        <!-- 3. 匯出 CSV -->
        <section class="momo-section">
          <h3>3. 匯出 CSV </h3>
          <div class="momo-row-inline">
            <button type="button" id="momo-download-url-btn" disabled>下載「序號 + 商品網址」</button>
            <button type="button" id="momo-download-title-btn" disabled>下載「序號 + 商品名稱」</button>
          </div>
          <div class="momo-small-tip">
            只會匯出「本次腳本啟動後」累積的新抓取結果；若設定了目標總筆數，匯出時將只取前 N 筆。
          </div>
        </section>

        <!-- 4. 歷史紀錄 -->
        <section class="momo-section">
          <h3>4. 歷史紀錄（localStorage）</h3>
          <div class="momo-row-inline" style="margin-bottom:4px;">
            <label class="momo-inline-label">
              <input type="checkbox" id="momo-use-history-filter" checked />
              使用 localStorage 歷史網址排除
            </label>
            <label class="momo-inline-label">
              <input type="checkbox" id="momo-write-history" checked />
              將新抓到的網址寫入 localStorage
            </label>
          </div>
          <div id="momo-history-summary" class="momo-summary">
            讀取中...
          </div>
          <div class="momo-row-inline">
            <button type="button" id="momo-download-history-btn">下載歷史網址.csv</button>
            <button type="button" id="momo-clear-history-btn">清空歷史網址</button>
          </div>
        </section>

        <!-- 5. 執行紀錄 -->
        <section class="momo-section">
          <h3>5. 執行紀錄（log）</h3>
          <div id="momo-log" class="momo-log"></div>
        </section>
      </div>
    `;

    document.body.appendChild(panel);

    GM_addStyle(`
      #momo-panel {
        position: fixed;
        right: 16px;
        bottom: 16px;
        z-index: 999998;
        width: 460px;
        min-width: 360px;
        max-width: 680px;
        max-height: 720px;
        background: #ffffff;
        border-radius: 8px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.22);
        font-size: 14px;
        font-family: -apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
        overflow: auto;
      }
      #momo-panel .momo-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 10px 12px;
        background: #ff4081;
        color: #fff;
        border-radius: 8px 8px 0 0;
        cursor: move;
        user-select: none;
      }
      #momo-panel .momo-header span {
        font-weight: 600;
        font-size: 15px;
      }
      #momo-panel .momo-header #momo-close-btn {
        background: transparent;
        border: none;
        color: #fff;
        font-size: 20px;
        cursor: pointer;
        padding: 0 4px;
        line-height: 1;
      }
      #momo-panel .momo-body {
        padding: 10px 12px 12px 12px;
      }
      #momo-panel .momo-section {
        margin-bottom: 12px;
        border-bottom: 1px solid #eee;
        padding-bottom: 8px;
      }
      #momo-panel .momo-section:last-of-type {
        border-bottom: none;
        padding-bottom: 4px;
      }
      #momo-panel h3 {
        margin: 6px 0 6px 0;
        font-size: 14px;
        font-weight: 600;
        color: #e91e63;
      }
      #momo-panel .momo-row {
        display: flex;
        flex-direction: column;
        margin-bottom: 8px;
        gap: 4px;
      }
      #momo-panel .momo-row span {
        font-weight: 500;
      }
      #momo-panel .momo-row-inline {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        align-items: center;
        margin-bottom: 6px;
      }
      #momo-panel input[type="file"] {
        font-size: 13px;
      }
      #momo-panel .momo-target-input {
        font-size: 12px;
        padding: 2px 4px;
        width: 160px;
        box-sizing: border-box;
      }
      #momo-panel .momo-inline-label {
        font-size: 12px;
        display: flex;
        align-items: center;
        gap: 4px;
      }
      #momo-panel button {
        padding: 7px 10px;
        font-size: 12px;
        background: #ff4081;
        color: #fff;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        white-space: nowrap;
      }
      #momo-panel button[disabled] {
        opacity: 0.55;
        cursor: default;
      }
      #momo-panel #momo-grab-btn {
        background: #ff6e40;
      }
      #momo-panel #momo-auto-grab-btn {
        background: #ff7043;
      }
      #momo-panel #momo-download-url-btn {
        background: #1976d2;
      }
      #momo-panel #momo-download-title-btn {
        background: #388e3c;
      }
      #momo-panel #momo-download-history-btn {
        background: #455a64;
      }
      #momo-panel #momo-clear-history-btn {
        background: #9e9e9e;
      }
      #momo-panel .momo-summary {
        font-size: 13px;
        line-height: 1.45;
        color: #444;
        background: #f7f7f7;
        padding: 6px 8px;
        border-radius: 4px;
        max-height: 96px;
        overflow-y: auto;
        white-space: pre-line;
      }
      #momo-panel .momo-count-status {
        margin-top: 6px;
        font-size: 13px;
        line-height: 1.4;
        background: #fff3f8;
        color: #c2185b;
        padding: 6px 8px;
        border-radius: 4px;
        white-space: pre-line;
      }
      #momo-panel .momo-small-tip {
        margin-top: 4px;
        font-size: 12px;
        color: #777;
      }
      #momo-panel .momo-log {
        font-size: 12px;
        line-height: 1.4;
        background: #fafafa;
        padding: 6px 8px;
        border-radius: 4px;
        max-height: 160px;
        overflow-y: auto;
        white-space: pre-line;
      }
    `);

    // 面板拖曳
    makePanelDraggable(panel);

    // 關閉按鈕
    document.getElementById('momo-close-btn').addEventListener('click', () => {
      panel.style.display = 'none';
    });

    // 規則檔
    document.getElementById('momo-rules-file').addEventListener('change', onRulesFileChange);

    // 匯入歷史 URL.csv
    document.getElementById('momo-history-file').addEventListener('change', onHistoryFileChange);

    // 抓取按鈕
    document.getElementById('momo-grab-btn').addEventListener('click', onGrabCurrentPage);
    document.getElementById('momo-auto-grab-btn').addEventListener('click', onAutoGrabAcrossPages);

    // 排除 TP 型網址
    const excludeTpCheckbox = document.getElementById('momo-exclude-tp');
    if (excludeTpCheckbox) {
      excludeTpCheckbox.checked = settings.excludeTpUrls;
      excludeTpCheckbox.addEventListener('change', (e) => {
        settings.excludeTpUrls = !!e.target.checked;
      });
    }

    // 匯出 CSV
    document.getElementById('momo-download-url-btn').addEventListener('click', onDownloadUrlCsv);
    document.getElementById('momo-download-title-btn').addEventListener('click', onDownloadTitleCsv);

    // 歷史紀錄
    document.getElementById('momo-download-history-btn').addEventListener('click', onDownloadHistoryCsv);
    document.getElementById('momo-clear-history-btn').addEventListener('click', onClearHistory);

    // 初始化歷史摘要 & 狀態
    updateHistorySummary();
    updateCountStatus();

    return panel;
  }

  function makePanelDraggable(panel) {
    const header = panel.querySelector('.momo-header');
    let isDragging = false;
    let startX, startY, startLeft, startTop;

    header.addEventListener('mousedown', (e) => {
      isDragging = true;
      startX = e.clientX;
      startY = e.clientY;
      const rect = panel.getBoundingClientRect();
      startLeft = rect.left;
      startTop = rect.top;
      e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
      if (!isDragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      panel.style.left = (startLeft + dx) + 'px';
      panel.style.top = (startTop + dy) + 'px';
      panel.style.right = 'auto';
      panel.style.bottom = 'auto';
    });

    document.addEventListener('mouseup', () => {
      isDragging = false;
    });
  }

  // ============================
  //  規則 TXT 讀取與解析
  // ============================
  function onRulesFileChange(e) {
    const file = e.target.files && e.target.files[0];
    const summaryEl = document.getElementById('momo-rules-summary');
    if (!file) {
      includeKeywords = [];
      excludeKeywords = [];
      rulesLoaded = false;
      summaryEl.textContent = '尚未載入規則（未載入則不做關鍵字過濾，只做去重）。';
      log('[提醒] 規則檔已清空。');
      return;
    }

    const reader = new FileReader();
    reader.onload = function (evt) {
      const text = String(evt.target.result || '');
      parseRulesText(text);
      rulesLoaded = true;

      const includeStr = includeKeywords.length
        ? '包含(+): ' + includeKeywords.join('、')
        : '包含(+): （未設定，表示不啟用白名單）';
      const excludeStr = excludeKeywords.length
        ? '排除(-): ' + excludeKeywords.join('、')
        : '排除(-): （未設定）';

      summaryEl.textContent =
        '已載入規則檔：' + file.name + '\n' +
        includeStr + '\n' +
        excludeStr;

      log(`[完成] 已載入規則檔：${file.name}。`);
    };
    reader.onerror = function () {
      rulesLoaded = false;
      includeKeywords = [];
      excludeKeywords = [];
      summaryEl.textContent = '讀取規則檔失敗，請重試或更換檔案。';
      log('[錯誤] 讀取規則檔失敗。');
    };
    reader.readAsText(file, 'utf-8');
  }

  function parseRulesText(text) {
    includeKeywords = [];
    excludeKeywords = [];

    const lines = text.split(/\r?\n/);
    for (let raw of lines) {
      if (!raw) continue;
      const line = raw.trim();
      if (!line) continue;
      if (line.startsWith('#')) continue;

      const firstChar = line.charAt(0);
      if (firstChar === '+' || firstChar === '＋') {
        const kw = line.slice(1).trim();
        if (kw) includeKeywords.push(kw.toLowerCase());
      } else if (firstChar === '-' || firstChar === '－') {
        const kw = line.slice(1).trim();
        if (kw) excludeKeywords.push(kw.toLowerCase());
      } else {
        // 其他行忽略
      }
    }
  }

  function matchesTitleFilter(title) {
    const text = (title || '').toLowerCase();

    // 白名單：有設定就必須命中其中一個
    if (includeKeywords.length > 0) {
      let hit = false;
      for (const kw of includeKeywords) {
        if (!kw) continue;
        if (text.includes(kw)) {
          hit = true;
          break;
        }
      }
      if (!hit) return false;
    }

    // 黑名單：有任一命中即排除
    for (const kw of excludeKeywords) {
      if (!kw) continue;
      if (text.includes(kw)) {
        return false;
      }
    }

    return true;
  }

  // ============================
  //  匯入歷史 URL.csv（排除清單）
  // ============================
  function onHistoryFileChange(e) {
    const file = e.target.files && e.target.files[0];
    const summaryEl = document.getElementById('momo-history-file-summary');
    if (!file) {
      externalExcludeSet = new Set();
      summaryEl.textContent = '尚未匯入外部歷史 URL（如需排除既有清單，可載入「序號,商品網址」格式的 CSV）。';
      log('[提醒] 外部歷史 URL 檔已清空。');
      return;
    }

    const reader = new FileReader();
    reader.onload = function (evt) {
      const text = String(evt.target.result || '');
      externalExcludeSet = parseHistoryCsvText(text);
      const count = externalExcludeSet.size;
      summaryEl.textContent =
        `已匯入外部歷史 URL 檔：${file.name}\n` +
        `可排除網址數量：${count} 筆（僅作為排除，不會寫回 localStorage）。`;
      log(`[完成] 已匯入外部歷史 URL：${file.name}，共 ${count} 筆。`);
    };
    reader.onerror = function () {
      externalExcludeSet = new Set();
      summaryEl.textContent = '讀取外部歷史 URL 檔失敗，請重試或更換檔案。';
      log('[錯誤] 讀取外部歷史 URL 檔失敗。');
    };
    reader.readAsText(file, 'utf-8');
  }

  function parseHistoryCsvText(text) {
    const set = new Set();
    const lines = text.split(/\r?\n/);
    for (const raw of lines) {
      if (!raw) continue;
      const line = raw.trim();
      if (!line) continue;
      // 跳過標題列
      if (/^序號[,，]/.test(line)) continue;

      const parts = line.split(',');
      for (let part of parts) {
        part = part.trim().replace(/^"|"$/g, '');
        if (!part) continue;
        if (/^https?:\/\//i.test(part)) {
          set.add(part);
          break;
        }
      }
    }
    return set;
  }

  // ============================
  //  歷史網址（localStorage）
  // ============================
  function loadHistoryFromLocalStorage() {
    try {
      const raw = localStorage.getItem(LS_KEY_HISTORY);
      if (!raw) return new Set();
      const arr = JSON.parse(raw);
      if (!Array.isArray(arr)) return new Set();
      return new Set(arr);
    } catch (e) {
      console.warn('[momo] 載入歷史網址失敗：', e);
      return new Set();
    }
  }

  function saveHistoryToLocalStorage() {
    try {
      const arr = Array.from(historyUrlSet);
      localStorage.setItem(LS_KEY_HISTORY, JSON.stringify(arr));
    } catch (e) {
      console.warn('[momo] 儲存歷史網址失敗：', e);
    }
  }

  function updateHistorySummary() {
    const el = document.getElementById('momo-history-summary');
    if (!el) return;
    const count = historyUrlSet.size;
    el.textContent =
      '目前 localStorage 歷史網址數量：' + count + ' 筆。\n' +
      '勾選「使用 localStorage 歷史網址排除」時，每次抓取都會先用此清單去重。\n' +
      '勾選「將新抓到的網址寫入 localStorage」時，符合條件的新網址才會加入此清單。';
  }

  function onDownloadHistoryCsv() {
    if (!historyUrlSet || historyUrlSet.size === 0) {
      log('[提醒] 目前沒有任何歷史網址可下載。');
      return;
    }
    const lines = [];
    lines.push('序號,商品網址');
    Array.from(historyUrlSet).forEach((url, idx) => {
      const n = idx + 1;
      lines.push(`${n},${url}`);
    });
    const csvText = lines.join('\n');
    const filename = 'momo_search_history_urls_' + getTimestampString() + '.csv';
    downloadCsv(csvText, filename);
    log(`[完成] 已下載歷史網址 CSV，共 ${historyUrlSet.size} 筆。`);
  }

  function onClearHistory() {
    try {
      localStorage.removeItem(LS_KEY_HISTORY);
    } catch (e) {
      console.warn('[momo] 清空歷史網址失敗：', e);
    }
    historyUrlSet = new Set();
    records = [];
    maxExportLimit = 0;
    lastGrabTotal = 0;
    lastGrabUnique = 0;
    updateHistorySummary();
    updateCountStatus();
    log('[完成] 已清空 localStorage 歷史網址，並清空本次抓取紀錄與目標上限。');
  }

  // ============================
  //  抓取目前頁面（JSON-LD） - 核心
  // ============================
  function onGrabCurrentPage() {
    const btn = document.getElementById('momo-grab-btn');
    if (btn) btn.disabled = true;

    try {
      const before = records.length;
      const stats = grabCurrentPageCore(null);
      if (!stats) return;
      const { listCount, pageTotal, pageUnique, addedThisPage } = stats;
      const added = typeof addedThisPage === 'number' ? addedThisPage : records.length - before;

      log(
        `[完成] 本頁 JSON-LD 共讀到 ${listCount} 筆商品，` +
        `有效卡片 ${pageTotal} 筆，其中「不重複且通過篩選」為 ${pageUnique} 筆（實際新增 ${added} 筆）。` +
        `\n目前本次腳本累積記錄：${records.length} 筆；歷史網址總數：${historyUrlSet.size} 筆。`
      );
    } catch (e) {
      console.error('[momo] 抓取過程錯誤：', e);
      log('[錯誤] 抓取過程發生例外，詳情請看 Console。');
    } finally {
      if (btn) btn.disabled = false;
    }
  }

  /**
   * 單頁抓取核心邏輯，可供「單頁抓取」及「自動跨頁」共用
   * @param {number|null|undefined} perPageQuota 單頁最多新增筆數（null/undefined 表示不限）
   * @returns {{listCount:number,pageTotal:number,pageUnique:number,addedThisPage:number}|null}
   */
  function grabCurrentPageCore(perPageQuota) {
    const list = extractProductsFromJsonLd();
    if (list.length === 0) {
      log('[提醒] 本頁面未在 JSON-LD 中找到任何商品 ItemList，請確認是否為搜尋結果頁。');
      lastGrabTotal = 0;
      lastGrabUnique = 0;
      updateCountStatus();
      return null;
    }

    const useHistoryFilter = !!(document.getElementById('momo-use-history-filter')?.checked);
    const writeHistory = !!(document.getElementById('momo-write-history')?.checked);

    let pageTotal = 0;
    let pageUnique = 0;
    let addedThisPage = 0;

    const quotaLimit = perPageQuota === null || perPageQuota === undefined ? null : perPageQuota;

    for (const item of list) {
      if (quotaLimit !== null && addedThisPage >= quotaLimit) {
        break;
      }

      const rawUrl = item.url;
      const title = item.title || '';

      if (!rawUrl) return;
      const url = canonicalizeMomoUrl(rawUrl);
      if (!url) return;

      const isTpUrl = isTpGoodsUrl(url);
      if (settings.excludeTpUrls && isTpUrl) {
        return;
      }

      pageTotal++;

      // 標題關鍵字過濾
      if (!matchesTitleFilter(title)) return;

      // 外部匯入歷史 URL.csv 排除
      if (externalExcludeSet.has(url)) return;

      // localStorage 歷史排除（可選）
      if (useHistoryFilter && historyUrlSet.has(url)) {
        return;
      }

      // 保留本次紀錄
      records.push({ url, title });
      pageUnique++;
      addedThisPage++;

      // 視設定決定是否寫回 localStorage
      if (writeHistory) {
        historyUrlSet.add(url);
      }
    }

    if (writeHistory) {
      saveHistoryToLocalStorage();
    }

    lastGrabTotal = pageTotal;
    lastGrabUnique = pageUnique;
    updateCountStatus();
    updateHistorySummary();

    if (records.length > 0) {
      const urlBtn = document.getElementById('momo-download-url-btn');
      const titleBtn = document.getElementById('momo-download-title-btn');
      if (urlBtn) urlBtn.disabled = false;
      if (titleBtn) titleBtn.disabled = false;
    }

    return { listCount: list.length, pageTotal, pageUnique, addedThisPage };
  }

  /**
   * 從頁面上的 <script type="application/ld+json"> 解析商品 ItemList
   * @returns {{url:string, title:string}[]}
   */

    /**
 * 嘗試把 ld+json 內容轉成物件：
 * - 先移除 BOM
 * - 移除多餘的結尾逗號（,] 或 ,}）
 * - 把 undefined 換成 null
 * - 再丟給 JSON.parse
 */
function safeParseLdJson(raw) {
  if (!raw) return null;

  let text = raw.trim();

  // 移除開頭 BOM
  text = text.replace(/^\uFEFF/, "");

  // 把 undefined 換成 null（預防之後 momo 亂塞）
  text = text.replace(/\bundefined\b/g, "null");

  // 移除物件 / 陣列尾端多餘的逗號：例如 [1,2,] 或 { "a":1, }
  text = text.replace(/,\s*([}\]])/g, "$1");

  try {
    return JSON.parse(text);
  } catch (e) {
    // 解析失敗就回傳 null，讓呼叫端自己略過
    // console.debug("safeParseLdJson 解析失敗:", e, text);
    return null;
  }
}

/**
 * 從 momo 搜尋結果頁的 JSON-LD 取出商品清單
 * 回傳 [{ url, title }]
 */
function extractProductsFromJsonLd() {
  const scripts = Array.from(
    document.querySelectorAll('script[type="application/ld+json"]')
  );

  /** @type {{url: string, title: string}[]} */
  const results = [];

  for (const s of scripts) {
    const raw = s.textContent;
    const data = safeParseLdJson(raw);
    if (!data) continue;

    const main = data.mainEntity;
    if (!main || main["@type"] !== "ItemList") continue;

    const list = Array.isArray(main.itemListElement)
      ? main.itemListElement
      : [];

    for (const p of list) {
      if (!p || p["@type"] !== "Product") continue;

      let url = (p.url || "").toString().trim();
      const title = (p.name || "").toString().trim();
      if (!url || !title) continue;

      // 有些 url 前後會多空白，順便清掉中間奇怪空白
      url = url.replace(/\s+/g, "");

      results.push({ url, title });
    }
  }

  return results;
}

  /**
   * momo 商品網址標準化，避免同商品不同追蹤參數造成重複
   * @param {string} href
   */
  function canonicalizeMomoUrl(href) {
    if (!href) return null;
    href = href.trim();
    try {
      // 某些 JSON-LD 可能是相對路徑或前面有空白
      const u = new URL(href, window.location.origin);

      if (!/momoshop\.com\.tw$/i.test(u.hostname)) return null;

      const origin = u.origin;
      const path = u.pathname;

      // 1) 一般 goodsDetail：/goods/GoodsDetail.jsp?i_code=14086840&Area=...
      const code = u.searchParams.get('i_code');
      if (path.startsWith('/goods/GoodsDetail.jsp') && code) {
        return `${origin}/goods/GoodsDetail.jsp?i_code=${code}`;
      }

      // 2) TP 型態：/TP/TP0002639/goodsDetail/TP00026390001044?xxx
      if (path.startsWith('/TP/')) {
        return origin + path;
      }

      // 3) 其他情況：退而求其次，保留 path + i_code（若有）
      if (code) {
        return `${origin}${path}?i_code=${code}`;
      }

      // 最後 fallback：直接 origin + path
      return origin + path;
    } catch (e) {
      return null;
    }
  }

  function isTpGoodsUrl(url) {
    if (!url) return false;
    try {
      const u = new URL(url, window.location.origin);
      const path = u.pathname || '';
      return /^\/TP\/TP\d+\/goodsDetail\//i.test(path);
    } catch (e) {
      return false;
    }
  }

  // ============================
  //  自動跨頁抓取（最多 N 筆）
  // ============================
  async function onAutoGrabAcrossPages() {
    if (isAutoPaging) {
      log('[提醒] 已經在自動跨頁抓取中，請先等待目前流程結束。');
      return;
    }

    const targetInput = document.getElementById('momo-target-count');
    let targetTotal = 0;
    if (targetInput && targetInput.value.trim() !== '') {
      const n = parseInt(targetInput.value.trim(), 10);
      if (!Number.isNaN(n) && n > 0) {
        targetTotal = n;
      }
    }

    maxExportLimit = targetTotal; // 設定「匯出上限 = 最多 N 筆」

    const targetDesc = targetTotal > 0 ? `${targetTotal} 筆（最多）` : '不限（抓完最後一頁為止）';
    log(`[開始] 自動跨頁抓取啟動，目標總筆數：${targetDesc}。`);

    isAutoPaging = true;

    const grabBtn = document.getElementById('momo-grab-btn');
    const autoBtn = document.getElementById('momo-auto-grab-btn');
    if (grabBtn) grabBtn.disabled = true;
    if (autoBtn) autoBtn.disabled = true;

    try {
      let safetyPageCount = 0;

      while (isAutoPaging) {
        safetyPageCount++;
        if (safetyPageCount > 200) {
          log('[警告] 自動跨頁已達安全頁數上限 200，為避免無限循環，已強制終止。');
          break;
        }

        const alreadyCollected = records.length;
        let perPageQuota = null;
        if (targetTotal > 0) {
          const quotaRemaining = targetTotal - alreadyCollected;
          if (quotaRemaining <= 0) {
            log(`[完成] 已達目標總筆數上限 ${targetTotal} 筆，自動跨頁結束。`);
            break;
          }
          perPageQuota = quotaRemaining;
        }

        const before = records.length;
        const stats = grabCurrentPageCore(perPageQuota);
        if (!stats) {
          log('[提醒] 無法從本頁抓取 JSON-LD 商品，自動跨頁流程結束。');
          break;
        }

        const added = typeof stats.addedThisPage === 'number' ? stats.addedThisPage : records.length - before;
        log(`[進度] 本頁新增可匯出筆數：${added} 筆；累積總筆數：${records.length} 筆。`);

        if (targetTotal > 0 && records.length >= targetTotal) {
          updateCountStatus();
          log(`[完成] 已達目標總筆數上限 ${targetTotal} 筆，自動跨頁結束。`);
          break;
        }

        const nextLink = getNextPageLink();
        if (!nextLink) {
          log('[完成] 找不到「下一頁」按鈕，可能已到最後一頁，自動跨頁流程結束。');
          break;
        }

        if (added === 0) {
          log('[提醒] 本頁沒有新增符合條件的商品，嘗試前往下一頁繼續。');
        }

        log('[進度] 已點擊「下一頁」，等待新頁面內容載入後繼續抓取...');
        nextLink.click();

        // momo 通常會部分更新內容（或重新載入），保守等 2 秒
        await sleep(2000);
      }
    } catch (e) {
      console.error('[momo] 自動跨頁流程錯誤：', e);
      log('[錯誤] 自動跨頁流程發生例外，詳細可見 Console。');
    } finally {
      isAutoPaging = false;
      if (grabBtn) grabBtn.disabled = false;
      if (autoBtn) autoBtn.disabled = false;
      log('[資訊] 自動跨頁流程已結束。');
      updateCountStatus();
    }
  }

  function getNextPageLink() {
    // <div class="page-control">
    //   <div class="page-total-products">...</div>
    //   <div class="page-btn page-next"><a>下一頁</a></div>
    // </div>
    const container = document.querySelector('.page-control .page-btn.page-next');
    if (!container) return null;

    if (container.classList.contains('disabled')) return null;

    const link = container.querySelector('a');
    if (!link) return null;

    return link;
  }

  // ============================
  //  匯出 CSV（尊重「最多 N 筆」）
  // ============================
  function onDownloadUrlCsv() {
    if (!records || records.length === 0) {
      log('[提醒] 目前沒有可匯出的資料（序號+商品網址），請先抓取頁面。');
      return;
    }
    const effectiveRecords =
      maxExportLimit > 0 ? records.slice(0, maxExportLimit) : records;

    const lines = [];
    lines.push('序號,商品網址');
    effectiveRecords.forEach((rec, idx) => {
      const n = idx + 1;
      lines.push(`${n},${rec.url}`);
    });
    const csvText = lines.join('\n');
    const filename = 'momo_search_urls_v1_' + getTimestampString() + '.csv';
    downloadCsv(csvText, filename);
    log(
      `[完成] 已下載「序號 + 商品網址」CSV，實際匯出 ${effectiveRecords.length} 筆` +
      (maxExportLimit > 0 ? `（目標上限：${maxExportLimit} 筆）` : '（未設定上限）') +
      '。'
    );
  }

  function onDownloadTitleCsv() {
    if (!records || records.length === 0) {
      log('[提醒] 目前沒有可匯出的資料（序號+商品名稱），請先抓取頁面。');
      return;
    }
    const effectiveRecords =
      maxExportLimit > 0 ? records.slice(0, maxExportLimit) : records;

    const lines = [];
    lines.push('序號,商品名稱');
    effectiveRecords.forEach((rec, idx) => {
      const n = idx + 1;
      const safeTitle = (rec.title || '').replace(/"/g, '""');
      lines.push(`${n},"${safeTitle}"`);
    });
    const csvText = lines.join('\n');
    const filename = 'momo_search_titles_v1_' + getTimestampString() + '.csv';
    downloadCsv(csvText, filename);
    log(
      `[完成] 已下載「序號 + 商品名稱」CSV，實際匯出 ${effectiveRecords.length} 筆` +
      (maxExportLimit > 0 ? `（目標上限：${maxExportLimit} 筆）` : '（未設定上限）') +
      '。'
    );
  }

  function downloadCsv(csvText, filename) {
    const BOM = '\uFEFF';
    const blob = new Blob([BOM + csvText], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    URL.revokeObjectURL(url);
  }

  // ============================
  //  顯示狀態 & log
  // ============================
  function updateCountStatus() {
    const el = document.getElementById('momo-count-status');
    if (!el) return;

    const exportable =
      maxExportLimit > 0 ? Math.min(records.length, maxExportLimit) : records.length;
    const limitText =
      maxExportLimit > 0 ? `匯出上限：${maxExportLimit} 筆` : '匯出上限：不限';

    if (lastGrabTotal === 0 && lastGrabUnique === 0 && records.length === 0) {
      el.textContent =
        `最近一頁：尚未抓取。\n本次腳本累積可匯出：0 筆（${limitText}）。`;
      return;
    }

    el.textContent =
      `最近一頁：共抓取 ${lastGrabTotal} 筆；其中「不重複且通過篩選」為 ${lastGrabUnique} 筆。\n` +
      `本次腳本累積記錄：${records.length} 筆；目前可匯出：${exportable} 筆（${limitText}）。`;
  }

  function log(msg) {
    const el = document.getElementById('momo-log');
    if (!el) return;
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, '0');
    const mm = String(now.getMinutes()).padStart(2, '0');
    const ss = String(now.getSeconds()).padStart(2, '0');
    const line = `${hh}:${mm}:${ss} ${msg}`;
    el.textContent = (el.textContent ? el.textContent + '\n' : '') + line;
    el.scrollTop = el.scrollHeight;
  }

  function getTimestampString() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const yyyy = d.getFullYear();
    const mm = pad(d.getMonth() + 1);
    const dd = pad(d.getDate());
    const hh = pad(d.getHours());
    const mi = pad(d.getMinutes());
    const ss = pad(d.getSeconds());
    return `${yyyy}${mm}${dd}_${hh}${mi}${ss}`;
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  // ============================
  //  啟動
  // ============================
  init();
})();
