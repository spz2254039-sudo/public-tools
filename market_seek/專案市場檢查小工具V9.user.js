// ==UserScript==
// @name         專案市場檢查小工具V9
// @namespace    http://tampermonkey.net/
// @version      V9
// @description  匯入 Excel / CSV（momo 匯出）建立待輸入清單，自動填寫市場檢查資料、支援自動下一筆、自動分類、自動送出、資料庫紀錄功能。
// @match        https://tp-masap-01.net.bsmi.gov.tw/Mas3103.action*
// @match        https://tp-masap-02.net.bsmi.gov.tw/Mas3103.action*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_setClipboard
// ==/UserScript==

(function () {
  'use strict';

  // V9 變更重點：
  // 1. 匯出 Excel 優先使用 raw.header，保留匯入檔所有標題欄（含自訂「型號」欄）。
  // 2. CSV 匯出增加「型號」欄位，匯出 r.model。
  // 3. 其餘功能沿用 V8：型號有值才帶入 MAS goods_modelno，空白不動原欄位。

  const DB_KEY = "mc_records";
  const PENDING_KEY = "pendingRecord";
  const QUEUE_KEY = "mc_queue";
  const QUEUE_INDEX_KEY = "mc_queueIndex";
  const CATEGORY_LAST_KEY = "mc_lastCategory";
  const AUTO_NEXT_KEY = "mc_autoNext";
  const AUTO_NEXT_FILL_KEY = "mc_autoNextFill";

  const AUTO_NAME_KEY = "mc_autoName";
  const AUTO_SUBMIT_KEY = "mc_autoSubmit";

  const RAW_ROWS_KEY = "mc_raw_rows";
  const ROW_CASE_MAP_KEY = "mc_row_case_map";

  const STANDARD_HEADER = [
    "編號",
    "檢查案號",
    "查核日期",
    "網路名稱/店家名稱",
    "賣家帳號或拍賣代碼",
    "商品名稱",
    "再查核日期",
    "是否下架",
    "是否改正",
    "調查結果",
    "網址/地址",
    "商檢標識",
    "已宣導",
    "已下架"
  ];

  const COLUMN_WIDTH_LIMITS = [
    12, // 編號
    22, // 檢查案號
    14, // 查核日期
    40, // 網路名稱/店家名稱
    30, // 賣家帳號或拍賣代碼
    60, // 商品名稱
    14, // 再查核日期
    12, // 是否下架
    12, // 是否改正
    60, // 調查結果
    100, // 網址/地址
    30, // 商檢標識
    12, // 已宣導
    12  // 已下架
  ];

  const WRAP_COL_INDEXES = [3, 5, 9, 10, 11];
  const WIDTH_PADDING = 10;

  // ============================
  //   商品類別對照表
  // ============================
  const categoryMap = {
    // A 類
    '風扇':         { bigCode: 'A', middleCode: 'A101', kindCode: '12' },
    '插頭及插座':   { bigCode: 'A', middleCode: 'A201', kindCode: '03-05' },
    '延長線':       { bigCode: 'A', middleCode: 'A201', kindCode: '03-06' },
    '燈具':         { bigCode: 'A', middleCode: 'A601', kindCode: '06-07' },
    '吹風機':       { bigCode: 'A', middleCode: 'A101', kindCode: '62' },
    '電鍋':         { bigCode: 'A', middleCode: 'A101', kindCode: '20-01' },
    'LED燈':        { bigCode: 'A', middleCode: 'A601', kindCode: '06-12' },
    '捲髮棒':       { bigCode: 'A', middleCode: 'A101', kindCode: '56-01' },
    '果汁機':       { bigCode: 'A', middleCode: 'A101', kindCode: '53' },
    '電壺':         { bigCode: 'A', middleCode: 'A101', kindCode: '26-01' },

    // B 類
    '行動電源':     { bigCode: 'B', middleCode: 'B201', kindCode: '92-29' },
    '鍵盤':         { bigCode: 'B', middleCode: 'B201', kindCode: '92-11' },
    '滑鼠':         { bigCode: 'B', middleCode: 'B201', kindCode: '92-13A' },
    '無人機':       { bigCode: 'B', middleCode: 'B201', kindCode: '92-36' },
    '充電器':       { bigCode: 'B', middleCode: 'B301', kindCode: '43' },
    '耳機':         { bigCode: 'B', middleCode: 'B101', kindCode: '38-10' },
    '音響':         { bigCode: 'B', middleCode: 'B101', kindCode: '38-01' },
    '計算器':       { bigCode: 'B', middleCode: 'B201', kindCode: '92-03' },

    // J 類
    '玩具':         { bigCode: 'J', middleCode: 'J501', kindCode: '65' },
    '寢具':         { bigCode: 'J', middleCode: 'JL07', kindCode: 'L07' },
    '毛巾':         { bigCode: 'J', middleCode: 'JL05', kindCode: 'L05' },
    '嬰幼兒服裝':   { bigCode: 'J', middleCode: 'JL01', kindCode: 'L01' },
    '安全帽':       { bigCode: 'J', middleCode: 'JE01', kindCode: '39-01' },
    '機車鏡片':     { bigCode: 'J', middleCode: 'JF01', kindCode: '17-02' },
    '筆擦':         { bigCode: 'J', middleCode: 'JT01', kindCode: 'JT01' },

    // M 類
    '瓦斯罐':       { bigCode: 'M', middleCode: 'MCP0', kindCode: '09-04' }
  };

  const style = document.createElement("style");
  style.textContent = `
    #mc-panel {
      position: fixed;
      top: 70px;
      right: 20px;
      width: 340px;
      background: #fff;
      border: 1px solid #ccc;
      box-shadow: 0 2px 6px rgba(0,0,0,.2);
      z-index: 2147483647;
      font-size: 12px;
      font-family: Arial, sans-serif;
    }
    #mc-header {
      background: #f0f0f0;
      padding: 5px;
      font-weight: bold;
      cursor: move;
    }
    #mc-body {
      padding: 6px;
    }
    #mc-category {
      width: 100%;
      margin-top: 4px;
    }
    .mc-btn {
      width: 100%;
      margin-top: 5px;
      font-size: 12px;
      padding: 3px;
      box-sizing: border-box;
    }
    #mc-autoNextWrap, #mc-autoNameWrap, #mc-autoSubmitWrap {
      margin-top: 6px;
    }
    #mc-toast {
      position: fixed;
      top: 10px;
      left: 50%;
      transform: translateX(-50%);
      background: #333;
      color: #fff;
      padding: 6px 10px;
      border-radius: 4px;
      z-index: 2147483647;
      font-size: 12px;
      white-space: pre-line;
    }
  `;
  document.head.appendChild(style);

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  function toast(msg) {
    const old = document.getElementById("mc-toast");
    if (old) old.remove();
    const d = document.createElement("div");
    d.id = "mc-toast";
    d.textContent = msg;
    document.body.appendChild(d);
    setTimeout(() => d.remove(), 3000);
  }

  function loadQueue() { return GM_getValue(QUEUE_KEY, []); }
  function saveQueue(v) { GM_setValue(QUEUE_KEY, v); }
  function loadIdx() { return GM_getValue(QUEUE_INDEX_KEY, 0); }
  function saveIdx(v) { GM_setValue(QUEUE_INDEX_KEY, v); }
  function loadAutoNext() { return GM_getValue(AUTO_NEXT_KEY, false); }
  function saveAutoNext(v) { GM_setValue(AUTO_NEXT_KEY, v); }
  function loadAutoName() { return GM_getValue(AUTO_NAME_KEY, false); }

  function isBlank(val) {
    return val === undefined || val === null || String(val).trim() === "";
  }

  function setCheckboxState(elem, checked) {
    if (!elem) return;
    if (elem.checked !== checked) elem.click();
  }

  function normalizeInspDateStr(val) {
    const t = (val || "").trim();
    if (!t) return "";
    const parts = t.split("/");
    if (parts.length === 3) {
      const [y, m = "", d = ""] = parts;
      return `${y.padStart(3, "0")}/${m.padStart(2, "0")}/${d.padStart(2, "0")}`;
    }
    return t;
  }

  function escapeXml(str) {
    return (str || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\"/g, "&quot;").replace(/'/g, "&apos;");
  }

  function colNumToName(num) {
    let n = num + 1;
    let name = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      name = String.fromCharCode(65 + rem) + name;
      n = Math.floor((n - 1) / 26);
    }
    return name;
  }

  const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[i] = c >>> 0;
    }
    return table;
  })();

  function crc32(arr) {
    let crc = 0 ^ (-1);
    for (let i = 0; i < arr.length; i++) {
      crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ arr[i]) & 0xFF];
    }
    return (crc ^ (-1)) >>> 0;
  }

  const textEncoder = new TextEncoder();

  function concatUint8(chunks) {
    let total = 0;
    for (const c of chunks) {
      if (!(c instanceof Uint8Array)) throw new Error("concatUint8 expects Uint8Array");
      total += c.length;
    }
    const out = new Uint8Array(total);
    let offset = 0;
    for (const c of chunks) {
      out.set(c, offset);
      offset += c.length;
    }
    if (offset !== total) throw new Error("concatUint8 length mismatch");
    return out;
  }

  function makeZip(files) {
    const centralParts = [];
    const fileParts = [];
    let offset = 0;

    for (const name of Object.keys(files)) {
      const contentStr = files[name];
      const data = contentStr instanceof Uint8Array ? contentStr : textEncoder.encode(contentStr);
      const crc = crc32(data);
      const nameBytes = textEncoder.encode(name);

      const localHeader = new ArrayBuffer(30);
      const lh = new DataView(localHeader);
      lh.setUint32(0, 0x04034b50, true);
      lh.setUint16(4, 20, true);
      lh.setUint16(6, 0, true);
      lh.setUint16(8, 0, true);
      lh.setUint16(10, 0, true);
      lh.setUint16(12, 0, true);
      lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true);
      lh.setUint32(22, data.length, true);
      lh.setUint16(26, nameBytes.length, true);
      lh.setUint16(28, 0, true);

      const localChunk = new Uint8Array(30 + nameBytes.length + data.length);
      localChunk.set(new Uint8Array(localHeader), 0);
      localChunk.set(nameBytes, 30);
      localChunk.set(data, 30 + nameBytes.length);
      fileParts.push(localChunk);

      const centralHeader = new ArrayBuffer(46);
      const ch = new DataView(centralHeader);
      ch.setUint32(0, 0x02014b50, true);
      ch.setUint16(4, 0x14, true);
      ch.setUint16(6, 20, true);
      ch.setUint16(8, 0, true);
      ch.setUint16(10, 0, true);
      ch.setUint16(12, 0, true);
      ch.setUint16(14, 0, true);
      ch.setUint32(16, crc, true);
      ch.setUint32(20, data.length, true);
      ch.setUint32(24, data.length, true);
      ch.setUint16(28, nameBytes.length, true);
      ch.setUint16(30, 0, true);
      ch.setUint16(32, 0, true);
      ch.setUint16(34, 0, true);
      ch.setUint16(36, 0, true);
      ch.setUint32(38, offset, true);

      const centralChunk = new Uint8Array(46 + nameBytes.length);
      centralChunk.set(new Uint8Array(centralHeader), 0);
      centralChunk.set(nameBytes, 46);
      centralParts.push(centralChunk);

      offset += localChunk.length;
    }

    const centralDir = concatUint8(centralParts);
    const fileData = concatUint8(fileParts);

    const eocd = new ArrayBuffer(22);
    const ev = new DataView(eocd);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, centralParts.length, true);
    ev.setUint16(10, centralParts.length, true);
    ev.setUint32(12, centralDir.length, true);
    ev.setUint32(16, offset, true);
    ev.setUint16(20, 0, true);

    return concatUint8([fileData, centralDir, new Uint8Array(eocd)]);
  }

  function computeMaxWidths(header, rows) {
    const maxWidth = header.map(h => (h || "").length);
    rows.forEach(cols => {
      cols.forEach((v, idx) => {
        const len = (v || "").toString().length;
        if (len > (maxWidth[idx] || 0)) maxWidth[idx] = len;
      });
    });
    return maxWidth;
  }

  function buildWorksheetXml(header, rows) {
    const widths = computeMaxWidths(header, rows);
    const colsXml = widths.map((w, idx) => {
      const limit = COLUMN_WIDTH_LIMITS[idx] || 20;
      const width = Math.min((w || 0) + WIDTH_PADDING, limit);
      const colIdx = idx + 1;
      return `<col min="${colIdx}" max="${colIdx}" width="${width}" customWidth="1" />`;
    }).join("");

    const rowXmlParts = [];
    const allRows = [header, ...rows];
    allRows.forEach((cols, rowIdx) => {
      const cells = cols.map((v, colIdx) => {
        const cellRef = `${colNumToName(colIdx)}${rowIdx + 1}`;
        const style = WRAP_COL_INDEXES.includes(colIdx) ? 1 : 0;
        const text = escapeXml(v || "");
        return `<c r="${cellRef}" t="inlineStr" s="${style}"><is><t xml:space="preserve">${text}</t></is></c>`;
      }).join("");
      rowXmlParts.push(`<row r="${rowIdx + 1}">${cells}</row>`);
    });

    const totalRows = allRows.length || 1;
    const lastColName = colNumToName(header.length - 1);
    const dimensionRef = `A1:${lastColName}${totalRows}`;

    return `<?xml version="1.0" encoding="UTF-8"?>` +
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<dimension ref="${dimensionRef}"/>` +
      `<sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
      `<sheetFormatPr defaultRowHeight="15"/>` +
      `<cols>${colsXml}</cols>` +
      `<sheetData>${rowXmlParts.join("")}</sheetData>` +
      `</worksheet>`;
  }

  function buildWorkbookXml() {
    return `<?xml version="1.0" encoding="UTF-8"?>` +
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>` +
      `</workbook>`;
  }

  function buildWorkbookRels() {
    return `<?xml version="1.0" encoding="UTF-8"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `</Relationships>`;
  }

  function buildRootRels() {
    return `<?xml version="1.0" encoding="UTF-8"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
      `</Relationships>`;
  }

  function buildContentTypes() {
    return `<?xml version="1.0" encoding="UTF-8"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
      `</Types>`;
  }

  function buildStylesXml() {
    return `<?xml version="1.0" encoding="UTF-8"?>` +
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
      `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="2">` +
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>` +
      `</cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
      `</styleSheet>`;
  }

  function buildXlsxFile(header, rows) {
    const sheetXml = buildWorksheetXml(header, rows);
    const files = {
      "[Content_Types].xml": buildContentTypes(),
      "_rels/.rels": buildRootRels(),
      "xl/workbook.xml": buildWorkbookXml(),
      "xl/_rels/workbook.xml.rels": buildWorkbookRels(),
      "xl/worksheets/sheet1.xml": sheetXml,
      "xl/styles.xml": buildStylesXml()
    };
    return makeZip(files);
  }

  function exportCsvFallback() {
    const db = GM_getValue(DB_KEY, []);
    if (!db.length) {
      toast("資料庫為空，無資料可匯出");
      return;
    }

    const header = "序號,檢查案號,商品名稱,型號,商檢標識,賣家帳號(品號),網址";
    const lines = db.map(r => [
      r.seq || "",
      r.caseNo || "",
      r.name || "",
      r.model || "",
      r.mark || "",
      r.seller_account || "",
      r.buy_site || ""
    ].join(","));

    const csv = [header, ...lines].join("\n");
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "市場檢查_已完成紀錄_fallback.csv";
    a.click();
    URL.revokeObjectURL(url);

    toast("已下載 CSV 備援匯出");
  }

  // -----------------------------
  //  Minimal XLSX reader (embedded)
  // -----------------------------
  const textDecoder = new TextDecoder("utf-8");

  async function decompressDeflateRaw(data) {
    if (typeof DecompressionStream === "undefined") {
      throw new Error("瀏覽器不支援 DecompressionStream，無法解壓 Excel");
    }
    const ds = new DecompressionStream("deflate-raw");
    const stream = new Response(new Blob([data]).stream().pipeThrough(ds));
    const buf = await stream.arrayBuffer();
    return new Uint8Array(buf);
  }

  async function unzipXlsx(arrayBuffer) {
    const view = new DataView(arrayBuffer);
    const u8 = new Uint8Array(arrayBuffer);
    let eocd = -1;
    for (let i = Math.max(0, u8.length - 0xFFFF); i <= u8.length - 22; i++) {
      if (view.getUint32(i, true) === 0x06054b50) {
        eocd = i;
      }
    }
    if (eocd === -1) throw new Error("無法找到 EOCD");

    const cdSize = view.getUint32(eocd + 12, true);
    const cdOffset = view.getUint32(eocd + 16, true);
    const files = {};
    let offset = cdOffset;
    while (offset < cdOffset + cdSize) {
      if (view.getUint32(offset, true) !== 0x02014b50) break;
      const compSize = view.getUint32(offset + 20, true);
      const nameLen = view.getUint16(offset + 28, true);
      const extraLen = view.getUint16(offset + 30, true);
      const commentLen = view.getUint16(offset + 32, true);
      const localOffset = view.getUint32(offset + 42, true);
      const name = textDecoder.decode(u8.slice(offset + 46, offset + 46 + nameLen));

      const lhNameLen = view.getUint16(localOffset + 26, true);
      const lhExtraLen = view.getUint16(localOffset + 28, true);
      const compMethod = view.getUint16(localOffset + 8, true);
      const dataStart = localOffset + 30 + lhNameLen + lhExtraLen;
      const compData = u8.slice(dataStart, dataStart + compSize);
      let content;
      if (compMethod === 0) content = compData;
      else if (compMethod === 8) content = await decompressDeflateRaw(compData);
      else throw new Error("不支援的壓縮格式");
      files[name] = content;
      offset += 46 + nameLen + extraLen + commentLen;
    }
    return files;
  }

  function parseSharedStrings(files) {
    const key = "xl/sharedStrings.xml";
    if (!files[key]) return [];
    const xml = textDecoder.decode(files[key]);
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const list = [];
    doc.querySelectorAll("si").forEach(si => {
      let text = "";
      si.querySelectorAll("t").forEach(t => { text += t.textContent || ""; });
      list.push(text);
    });
    return list;
  }

  function letterToIndex(ref) {
    const letters = ref.replace(/\d+/g, "");
    let idx = 0;
    for (let i = 0; i < letters.length; i++) {
      idx = idx * 26 + (letters.charCodeAt(i) - 64);
    }
    return idx - 1;
  }

  function parseSheetRows(sheetXml, sharedStrings) {
    const doc = new DOMParser().parseFromString(sheetXml, "application/xml");
    const rows = [];
    doc.querySelectorAll("sheetData > row").forEach(row => {
      const cells = [];
      row.querySelectorAll("c").forEach(c => {
        const ref = c.getAttribute("r") || "";
        const t = c.getAttribute("t") || "";
        const v = c.querySelector("v");
        let value = v ? v.textContent || "" : "";
        if (t === "s") {
          const idx = Number(value);
          value = Number.isFinite(idx) && sharedStrings[idx] !== undefined ? sharedStrings[idx] : "";
        }
        const colIdx = letterToIndex(ref || "");
        cells[colIdx] = value;
      });
      rows.push(cells);
    });
    return rows;
  }

  function getByHeaderKeys(row, headerIndex, keys) {
    for (const k of keys) {
      const idx = headerIndex[k];
      if (idx === undefined) continue;
      const val = row[idx];
      if (val !== undefined && val !== null && String(val).trim() !== "") {
        return String(val).trim();
      }
    }
    return "";
  }

  async function readFirstSheet(arrayBuffer) {
    const files = await unzipXlsx(arrayBuffer);
    const workbookXml = files["xl/workbook.xml"] ? textDecoder.decode(files["xl/workbook.xml"]) : null;
    if (!workbookXml) throw new Error("缺少 workbook 資料");

    const wbDoc = new DOMParser().parseFromString(workbookXml, "application/xml");
    const firstSheet = wbDoc.querySelector("sheet");
    if (!firstSheet) throw new Error("找不到工作表");
    const relId = firstSheet.getAttribute("r:id");

    const relsXml = files["xl/_rels/workbook.xml.rels"] ? textDecoder.decode(files["xl/_rels/workbook.xml.rels"]) : "";
    const relDoc = new DOMParser().parseFromString(relsXml, "application/xml");
    let target = "worksheets/sheet1.xml";
    relDoc.querySelectorAll("Relationship").forEach(rel => {
      if (rel.getAttribute("Id") === relId) target = rel.getAttribute("Target") || target;
    });
    if (!target.startsWith("xl/")) target = `xl/${target}`;

    const sheetFile = files[target];
    if (!sheetFile) throw new Error("找不到工作表內容");
    const sheetXml = textDecoder.decode(sheetFile);
    const sharedStrings = parseSharedStrings(files);
    const rows = parseSheetRows(sheetXml, sharedStrings);
    return { rows };
  }

  async function parseExcelFile(file) {
    try {
      const buffer = await file.arrayBuffer();
      const { rows } = await readFirstSheet(buffer);
      if (!rows.length) return [];
      const headerRow = rows[0].map(v => (v || "").toString().trim());
      const headerIndex = {};
      headerRow.forEach((h, i) => { if (h) headerIndex[h] = i; });

      const queue = [];
      const rawRows = [];
      const autoName = loadAutoName();
      const headerToUse = headerRow.length ? headerRow : STANDARD_HEADER.slice();

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i] || [];
        if (row.every(v => v === undefined || v === null || String(v).trim() === "")) continue;

        const seqVal = getByHeaderKeys(row, headerIndex, ["編號"]) || String(i);
        const markVal = getByHeaderKeys(row, headerIndex, ["商檢標識"]);
        const nameVal = getByHeaderKeys(row, headerIndex, ["商品名稱"]);
        const sellerVal = getByHeaderKeys(row, headerIndex, ["賣家帳號或拍賣代碼", "品號", "賣家帳號"]);
        const urlVal = getByHeaderKeys(row, headerIndex, ["網址/地址"]);
        const inspDateVal = getByHeaderKeys(row, headerIndex, ["查核日期"]);
        const modelVal = getByHeaderKeys(row, headerIndex, ["型號", "商品型號", "型號/規格"]);

        if (!markVal && !nameVal && !urlVal) continue;

        const paddedRow = Array(headerToUse.length).fill("");
        for (let c = 0; c < headerToUse.length; c++) paddedRow[c] = row[c] !== undefined && row[c] !== null ? String(row[c]) : "";
        rawRows.push({ rowId: seqVal, cols: paddedRow });

        const rec = {
          seq: seqVal,
          mark: markVal,
          name: nameVal,
          seller_account: sellerVal,
          buy_site: urlVal,
          inspDateStr: inspDateVal,
          model: modelVal,
          origin: "ZZ",
          status: "符合"
        };

        if (autoName) {
          const cat = guessCategoryFromName(rec.name);
          if (cat) rec.categoryKey = cat;
        }

        queue.push(rec);
      }

      GM_setValue(RAW_ROWS_KEY, { header: headerToUse, rows: rawRows });
      GM_setValue(ROW_CASE_MAP_KEY, {});

      return queue;
    } catch (e) {
      console.error(e);
      toast("Excel 解析失敗，請改用 CSV 匯入");
      return null;
    }
  }

  // ------------------------------------------------
  //  CSV 解析
  // ------------------------------------------------
  function parseCsvFile(file) {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onerror = () => {
        toast(`匯入失敗：${reader.error || "讀取 CSV 失敗"}`);
        resolve(null);
      };
      reader.onload = () => {
        const text = reader.result;
        if (!text) return resolve([]);
        let lines = String(text).split(/\r?\n/).map(s => s.trim()).filter(Boolean);
        if (!lines.length) return resolve([]);

        const headerCols = lines[0].split(",").map(s => s.trim());
        const headerIndex = {};
        headerCols.forEach((h, i) => { if (h) headerIndex[h] = i; });

        const queue = [];
        const rawRows = [];
        const autoName = loadAutoName();
        const headerToUse = headerCols.length ? headerCols : STANDARD_HEADER.slice();

        const getVal = (cols, keys) => getByHeaderKeys(cols, headerIndex, keys);

        for (let i = 1; i < lines.length; i++) {
          const raw = lines[i];
          if (!raw) continue;
          const cols = raw.split(",").map(s => s.trim());
          if (cols.every(v => !v)) continue;

          const seqVal = getVal(cols, ["編號"]) || String(i);
          const markVal = getVal(cols, ["商檢標識"]);
          const nameVal = getVal(cols, ["商品名稱"]);
          const sellerVal = getVal(cols, ["賣家帳號或拍賣代碼", "品號", "賣家帳號"]);
          const urlVal = getVal(cols, ["網址/地址"]);
          const inspDateVal = getVal(cols, ["查核日期"]);
          const modelVal = getVal(cols, ["型號", "商品型號", "型號/規格"]);

          if (!markVal && !nameVal && !urlVal) continue;

          const paddedRow = Array(headerToUse.length).fill("");
          for (let c = 0; c < headerToUse.length; c++) paddedRow[c] = cols[c] !== undefined && cols[c] !== null ? String(cols[c]) : "";
          rawRows.push({ rowId: seqVal, cols: paddedRow });

          const rec = {
            seq: seqVal,
            mark: markVal,
            name: nameVal,
            seller_account: sellerVal,
            buy_site: urlVal,
            inspDateStr: inspDateVal,
            model: modelVal,
            origin: "ZZ",
            status: "符合"
          };

          if (autoName) {
            const cat = guessCategoryFromName(rec.name);
            if (cat) rec.categoryKey = cat;
          }

          queue.push(rec);
        }

        GM_setValue(RAW_ROWS_KEY, { header: headerToUse, rows: rawRows });
        GM_setValue(ROW_CASE_MAP_KEY, {});

        resolve(queue);
      };

      reader.readAsText(file);
    });
  }

  function guessCategoryFromName(name) {
    if (!name) return "";
    if (categoryMap[name]) return name;
    for (const key of Object.keys(categoryMap)) {
      if (name.includes(key)) return key;
    }
    return "";
  }

  function autoClickConfirmIfEnabled() {
    if (!GM_getValue(AUTO_SUBMIT_KEY, false)) return;
    setTimeout(() => {
      const btns = document.querySelectorAll("input[type='button'], button, input[type='submit']");
      for (const b of btns) {
        const v = (b.value || b.textContent || "").trim();
        if (v.includes("確定")) {
          b.click();
          toast("已自動按下「確定」");
          break;
        }
      }
    }, 1000);
  }

  function getForm() {
    const noMarkCheckbox = document.evaluate("/html/body/div[2]/table/tbody/tr[4]/td/div/div/form/div[6]/fieldset/table/tbody/tr[5]/td[2]/input[1]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    return {
      status1: document.querySelector('#addInspCaseA\\.inspExec\\.status1'),
      status2: document.querySelector('#addInspCaseA\\.inspExec\\.status2'),
      status6: document.querySelector('#addInspCaseA\\.inspExec\\.status6'),
      unmark: document.querySelector('#addInspCaseA\\.inspExec\\.unmark_no'),
      noMarkCheckbox,
      name: document.querySelector('#addInspCaseA\\.goods_name'),
      sellerAccount: document.querySelector('#addInspCaseA\\.seller_account'),
      buySite: document.querySelector('#addInspCaseA\\.buy_site'),
      prd1: document.querySelector('#prd1_SCN3'),
      prd2: document.querySelector('#prd2_SCN3'),
      prd3: document.querySelector('#prd3_SCN3'),
      model: document.querySelector('#addInspCaseA\\.goods_modelno'),
      prodCountryHelpers: document.querySelectorAll('input#prod_country'),
      prodCountrySelect: document.querySelector('#addInspCaseA\\.prod_country'),
      conform: document.querySelector('#addInspCaseA\\.conform'),
      insp_dateStr: document.querySelector('#addInspCaseA\\.insp_dateStr')
    };
  }

  async function fillForm(rec, cat) {
    const el = getForm();
    const hasMark = !isBlank(rec.mark);

    setCheckboxState(el.status1, false);
    setCheckboxState(el.status2, false);
    setCheckboxState(el.noMarkCheckbox, false);

    if (el.unmark) {
      if (hasMark) {
        el.unmark.value = rec.mark;
        el.unmark.dispatchEvent(new Event("blur", { bubbles: true }));
        setCheckboxState(el.status6, false);
      } else {
        el.unmark.value = "";
        el.unmark.dispatchEvent(new Event("blur", { bubbles: true }));
        setCheckboxState(el.status6, true);
      }
    }

    if (hasMark) {
      setCheckboxState(el.status1, false);
      setCheckboxState(el.status2, true);
      setCheckboxState(el.noMarkCheckbox, false);
    } else {
      setCheckboxState(el.status1, true);
      setCheckboxState(el.status2, false);
      setCheckboxState(el.noMarkCheckbox, true);
    }

    if (el.name && rec.name) {
      el.name.value = rec.name;
      el.name.dispatchEvent(new Event("change", { bubbles: true }));
    }

    if (el.sellerAccount && rec.seller_account) {
      el.sellerAccount.value = rec.seller_account;
      el.sellerAccount.dispatchEvent(new Event("change", { bubbles: true }));
    }

    if (el.buySite) {
      const url = rec.buy_site || "";
      el.buySite.value = url;
      el.buySite.dispatchEvent(new Event("change", { bubbles: true }));
    }

    if (el.insp_dateStr) {
      const dateVal = normalizeInspDateStr(rec.inspDateStr || "");
      el.insp_dateStr.value = dateVal;
      el.insp_dateStr.dispatchEvent(new Event("blur", { bubbles: true }));
    }

    if (cat && categoryMap[cat]) {
      const { bigCode, middleCode, kindCode } = categoryMap[cat];

      if (el.prd1) {
        el.prd1.value = bigCode;
        el.prd1.dispatchEvent(new Event("keyup", { bubbles: true }));
      }
      await sleep(500);

      if (el.prd2) {
        el.prd2.value = middleCode;
        el.prd2.dispatchEvent(new Event("keyup", { bubbles: true }));
      }
      await sleep(500);

      if (el.prd3) {
        el.prd3.value = kindCode;
        el.prd3.dispatchEvent(new Event("keyup", { bubbles: true }));
      }
    }

    const hasModel = !isBlank(rec.model);
    if (el.model && hasModel) {
      el.model.value = rec.model;
      el.model.dispatchEvent(new Event("blur", { bubbles: true }));
    }

    const code = "ZZ";
    if (el.prodCountryHelpers.length) {
      el.prodCountryHelpers.forEach(input => {
        input.value = code;
        input.dispatchEvent(new Event("keyup", { bubbles: true }));
        input.dispatchEvent(new Event("blur", { bubbles: true }));
      });
    }
    if (el.prodCountrySelect) {
      el.prodCountrySelect.value = code;
      el.prodCountrySelect.dispatchEvent(new Event("change", { bubbles: true }));
    }

    if (el.conform && !el.conform.checked) el.conform.click();

    if (GM_getValue(AUTO_SUBMIT_KEY, false)) {
      toast("欄位已填入，準備自動送出…");
      autoClickConfirmIfEnabled();
    } else toast("欄位已填入（尚未送出）");
  }

  function buildPanel() {
    const d = document.createElement("div");
    d.id = "mc-panel";
    d.innerHTML = `
      <div id="mc-header">市場檢查小工具（Excel / CSV 版）</div>
      <div id="mc-body">
        <button class="mc-btn" id="mc-btn-importFile">匯入檔案（momo Excel / CSV）</button>
        <input type="file" id="mc-file-excel" accept=".xlsx,.csv" style="display:none;">
        <div style="margin-top:4px;font-size:11px;color:#555;">
          說明：可直接選擇 momo 匯出的 .xlsx，或你另存的 .csv 檔。
        </div>

        <select id="mc-category">
          <option value="">（選商品類別）</option>
          ${Object.keys(categoryMap).map(k => `<option value="${k}">${k}</option>`).join("")}
        </select>

        <div id="mc-autoNextWrap">
          <input type="checkbox" id="mc-autoNext">
          <label for="mc-autoNext">自動下一筆（確定後自動複製＋清空＋填入）</label>
        </div>

        <div id="mc-autoNameWrap">
          <input type="checkbox" id="mc-autoName">
          <label for="mc-autoName">依品名自動帶入商品類別</label>
        </div>

        <div id="mc-autoSubmitWrap">
          <input type="checkbox" id="mc-autoSubmit">
          <label for="mc-autoSubmit">自動送出（自動按「確定」）</label>
        </div>

        <hr>
        <div>待輸入清單</div>
        <div id="mc-queue-info" style="margin-top:6px;"></div>
        <button class="mc-btn" id="mc-btn-next">填下一筆</button>
        <button class="mc-btn" id="mc-btn-skip">跳過這筆</button>
        <button class="mc-btn" id="mc-btn-clearQueue">清空清單</button>

        <hr>
        <div>已完成紀錄</div>
        <div id="mc-db-count" style="margin-top:6px;"></div>
        <button class="mc-btn" id="mc-btn-export">下載紀錄（Excel）</button>
        <button class="mc-btn" id="mc-btn-clearDb">清空紀錄</button>
      </div>
    `;
    document.body.appendChild(d);

    makeDraggable(d, d.querySelector("#mc-header"));
    bindPanelActions();
    refreshPanel();

    const lastCat = GM_getValue(CATEGORY_LAST_KEY, "");
    if (lastCat && categoryMap[lastCat]) {
      document.getElementById("mc-category").value = lastCat;
    }

    document.getElementById("mc-autoNext").checked = loadAutoNext();
    document.getElementById("mc-autoNext").onchange = e => saveAutoNext(e.target.checked);

    document.getElementById("mc-autoName").checked = GM_getValue(AUTO_NAME_KEY, false);
    document.getElementById("mc-autoName").onchange = e => GM_setValue(AUTO_NAME_KEY, e.target.checked);

    document.getElementById("mc-autoSubmit").checked = GM_getValue(AUTO_SUBMIT_KEY, false);
    document.getElementById("mc-autoSubmit").onchange = e => GM_setValue(AUTO_SUBMIT_KEY, e.target.checked);
  }

  function makeDraggable(elem, handle) {
    let drag = false, ox = 0, oy = 0;

    handle.onmousedown = e => {
      drag = true;
      ox = e.clientX - elem.offsetLeft;
      oy = e.clientY - elem.offsetTop;
    };

    document.onmousemove = e => {
      if (!drag) return;
      elem.style.left = (e.clientX - ox) + "px";
      elem.style.top = (e.clientY - oy) + "px";
      elem.style.right = "auto";
    };

    document.onmouseup = () => drag = false;
  }

  function refreshPanel() {
    const q = loadQueue();
    const idx = loadIdx();
    const div = document.getElementById("mc-queue-info");

    if (!q.length) div.textContent = "待輸入清單：0 筆";
    else if (idx >= q.length) div.textContent = `待輸入清單：${q.length} 筆（已全部處理完）`;
    else div.textContent = `待輸入清單：${q.length} 筆，下一筆：${idx + 1} / ${q.length}`;

    const dbCount = GM_getValue(DB_KEY, []).length;
    document.getElementById("mc-db-count").textContent = `資料庫：${dbCount} 筆`;
  }

  // ------------------------------------------------
  //  匯入 Excel / CSV
  // ------------------------------------------------
  function bindPanelActions() {

    const importBtn = document.getElementById("mc-btn-importFile");
    const fileInput = document.getElementById("mc-file-excel");

    importBtn.onclick = () => {
      fileInput.value = "";
      fileInput.click();
    };

    fileInput.onchange = async () => {
      const f = fileInput.files?.[0];
      if (!f) return;
      let q = null;
      try {
        const nameLower = f.name.toLowerCase();
        if (nameLower.endsWith(".xlsx")) {
          q = await parseExcelFile(f);
          if (q === null) return;
        } else if (nameLower.endsWith(".csv")) {
          q = await parseCsvFile(f);
          if (q === null) return;
        } else {
          toast("不支援的檔案格式，請選擇 .xlsx 或 .csv");
          return;
        }
      } catch (e) {
        console.error(e);
        toast(`匯入失敗：${e.message || e}`);
        return;
      }

      if (!Array.isArray(q)) return;

      saveQueue(q);
      saveIdx(0);
      toast(`已匯入 ${q.length} 筆資料`);
      const rawCheck = GM_getValue(RAW_ROWS_KEY, null);
      if (!rawCheck || !rawCheck.rows || !rawCheck.rows.length) {
        toast("匯入看似成功，但 RAW_ROWS 未正確建立，請重新整理頁面後再匯入一次。");
      }
      refreshPanel();
    };

    document.getElementById("mc-btn-next").onclick = () => {
      const q = loadQueue();
      const idx = loadIdx();
      if (!q.length) return toast("待輸入清單是空的");
      if (idx >= q.length) return toast("已全部處理完");

      const rec = q[idx];
      const autoName = loadAutoName();
      let cat = document.getElementById("mc-category").value;

      if (!cat && autoName) {
        cat = rec.categoryKey || guessCategoryFromName(rec.name);
        if (cat) document.getElementById("mc-category").value = cat;
      }

      if (!cat) return toast("請選商品類別，或勾選依品名自動帶入");

      GM_setValue(CATEGORY_LAST_KEY, cat);
      rec.categoryKey = cat;
      GM_setValue(PENDING_KEY, rec);

      fillForm(rec, cat);
    };

    document.getElementById("mc-btn-skip").onclick = () => {
      const q = loadQueue();
      let idx = loadIdx();

      if (!q.length) return toast("待輸入清單是空的");
      if (idx >= q.length) return toast("已在最後一筆");

      idx++;
      saveIdx(idx);
      refreshPanel();
      toast("已跳過這筆");
    };

    document.getElementById("mc-btn-clearQueue").onclick = () => {
      saveQueue([]);
      saveIdx(0);
      refreshPanel();
      toast("已清空待輸入清單");
    };

    document.getElementById("mc-btn-export").onclick = () => {
      const raw = GM_getValue(RAW_ROWS_KEY, null);
      if (!raw || !raw.rows || !raw.rows.length) {
        toast("匯出 Excel 需要先匯入一次 momo Excel，請先按「匯入檔案（momo Excel / CSV）」。");
        return;
      }

      const db = GM_getValue(DB_KEY, []);
      const completed = db.filter(r => !isBlank(r.caseNo));
      if (!completed.length) {
        toast("目前沒有已完成案件，無可匯出");
        return;
      }

      try {
        const header = (raw.header && raw.header.length) ? raw.header.slice() : STANDARD_HEADER.slice();
        const headerLen = header.length;
        const headerIndexRaw = {};
        (raw.header || []).forEach((h, i) => { headerIndexRaw[h] = i; });
        const caseIdx = 1; // 「檢查案號」欄位固定第二欄

        const rows = completed.map(rec => {
          const srcRow = (raw.rows || []).find(r => String(r.rowId) === String(rec.seq));
          const srcCols = srcRow && Array.isArray(srcRow.cols) ? srcRow.cols : [];

          // 若找不到原始列，則以空白欄位填滿，僅填入編號與檢查案號（避免資料遺失）。
          const cols = Array(headerLen).fill("");
          header.forEach((h, idx) => {
            if (headerIndexRaw[h] !== undefined && srcCols[headerIndexRaw[h]] !== undefined && srcCols[headerIndexRaw[h]] !== null) {
              cols[idx] = String(srcCols[headerIndexRaw[h]]);
            }
          });
          cols[0] = rec.seq || cols[0] || "";
          cols[caseIdx] = rec.caseNo || "";
          return cols;
        });

        const xlsxBuffer = buildXlsxFile(header, rows);
        const blob = new Blob([xlsxBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "市場檢查_已完成紀錄.xlsx";
        a.click();
        URL.revokeObjectURL(url);

        toast("已下載紀錄 Excel");
      } catch (e) {
        console.error(e);
        toast(`匯出 Excel 失敗：${e.message || e}`);
        toast("Excel 匯出失敗，改用 CSV 匯出。");
        exportCsvFallback();
      }
    };

    document.getElementById("mc-btn-clearDb").onclick = () => {
      GM_setValue(DB_KEY, []);
      refreshPanel();
      toast("已清空紀錄");
    };
  }

  function getCaseNo() {
    const tds = document.querySelectorAll("td.item");
    for (const td of tds) {
      if (td.textContent.trim() === "檢查案號") {
        const v = td.nextElementSibling;
        return v ? v.textContent.trim() : null;
      }
    }
    return null;
  }

  function handleCasePage() {
    const rec = GM_getValue(PENDING_KEY, null);
    if (!rec) return;

    window.addEventListener("load", async () => {
      await sleep(500);

      const caseNo = getCaseNo();
      if (!caseNo) {
        GM_setValue(PENDING_KEY, null);
        return;
      }

      const now = new Date().toISOString();
      const record = {
        seq: rec.seq,
        caseNo,
        name: rec.name,
        mark: rec.mark,
        seller_account: rec.seller_account,
        buy_site: rec.buy_site,
        origin: "ZZ",
        status: rec.status,
        categoryKey: rec.categoryKey,
        createdAt: now
      };

      const db = GM_getValue(DB_KEY, []);
      const existingIndex = db.findIndex(r => r.seq === record.seq);
      if (existingIndex >= 0) db[existingIndex] = record;
      else db.push(record);
      GM_setValue(DB_KEY, db);

      const rowCaseMap = GM_getValue(ROW_CASE_MAP_KEY, {});
      rowCaseMap[rec.seq] = caseNo;
      GM_setValue(ROW_CASE_MAP_KEY, rowCaseMap);

      try { GM_setClipboard(caseNo, "text"); } catch {}

      let idx = loadIdx();
      const q = loadQueue();
      if (idx < q.length) idx++;
      saveIdx(idx);

      GM_setValue(PENDING_KEY, null);

      toast(`檢查案號：${caseNo}`);

      if (loadAutoNext()) {
        GM_setValue(AUTO_NEXT_FILL_KEY, true);
        const btn = document.querySelector("#doCopyA");
        if (btn) btn.click();
      }
    });
  }

  function handleNewPageAutoFill() {
    window.addEventListener("load", async () => {
      const need = GM_getValue(AUTO_NEXT_FILL_KEY, false);
      if (!need) return;

      GM_setValue(AUTO_NEXT_FILL_KEY, false);

      await sleep(600);

      const clearBtn = document.querySelector("#clsDetailA");
      if (clearBtn) {
        clearBtn.click();
        await sleep(500);
      }

      const q = loadQueue();
      let idx = loadIdx();

      if (!q.length || idx >= q.length) {
        toast("待輸入清單已處理完");
        return;
      }

      const rec = q[idx];
      const autoName = loadAutoName();

      let cat;
      if (autoName)
        cat = rec.categoryKey || guessCategoryFromName(rec.name) || GM_getValue(CATEGORY_LAST_KEY, "");
      else
        cat = GM_getValue(CATEGORY_LAST_KEY, "");

      if (!cat) {
        toast("無分類資訊，無法自動填入");
        return;
      }

      rec.categoryKey = cat;
      GM_setValue(PENDING_KEY, rec);

      await fillForm(rec, cat);

      toast(`已自動填入第 ${idx + 1} 筆`);
      refreshPanel();
    });
  }

  buildPanel();
  handleCasePage();
  handleNewPageAutoFill();

})();
