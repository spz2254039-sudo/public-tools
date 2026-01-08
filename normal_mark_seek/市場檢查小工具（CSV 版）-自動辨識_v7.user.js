// ==UserScript==
// @name         市場檢查小工具（CSV 版）-自動辨識-v7
// @namespace    http://tampermonkey.net/
// @version      7
// @description  匯入 CSV 建立待輸入清單，自動填寫市場檢查資料、支援自動下一筆、自動分類、自動送出、資料庫紀錄功能（自動辨識 UTF-8 / Big5 編碼）。
// @match        https://tp-masap-01.net.bsmi.gov.tw/Mas3103.action*
// @match        https://tp-masap-02.net.bsmi.gov.tw/Mas3103.action*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_setClipboard
// ==/UserScript==

(function () {
  'use strict';

  const DB_KEY = "mc_records";
  const PENDING_KEY = "pendingRecord";
  const QUEUE_KEY = "mc_queue";
  const QUEUE_INDEX_KEY = "mc_queueIndex";
  const CATEGORY_LAST_KEY = "mc_lastCategory";
  const AUTO_NEXT_KEY = "mc_autoNext";
  const AUTO_NEXT_FILL_KEY = "mc_autoNextFill";

  const AUTO_NAME_KEY = "mc_autoName";
  const AUTO_SUBMIT_KEY = "mc_autoSubmit";

  // ============================
  //   商品類別對照表
  // ============================
  const categoryMap = {
    // A 類
    '風扇':         { bigCode: 'A', middleCode: 'A101', kindCode: '12' },
    '插頭及插座':   { bigCode: 'A', middleCode: 'A201', kindCode: '03-05' },
    '延長線':     { bigCode: 'A', middleCode: 'A201', kindCode: '03-06' },
    '燈具':         { bigCode: 'A', middleCode: 'A601', kindCode: '06-07' },
    '吹風機':       { bigCode: 'A', middleCode: 'A101', kindCode: '62' },
    '電鍋':         { bigCode: 'A', middleCode: 'A101', kindCode: '20-01' },
    'LED燈':        { bigCode: 'A', middleCode: 'A601', kindCode: '06-12' },
    '捲髮棒':       { bigCode: 'A', middleCode: 'A101', kindCode: '56-01' },
    '果汁機':       { bigCode: 'A', middleCode: 'A101', kindCode: '53' },
    '電壺':         { bigCode: 'A', middleCode: 'A101', kindCode: '26-01' },
    '電保暖器':      { bigCode: 'A', middleCode: 'A101', kindCode: '92-31' },

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
    #mc-paste {
      width: 100%;
      height: 80px;
      margin-top: 6px;
      box-sizing: border-box;
      font-size: 12px;
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

  function pad2(n) { return n.toString().padStart(2, "0"); }

  function formatDateTime(d) {
    const y = d.getFullYear();
    const m = pad2(d.getMonth() + 1);
    const day = pad2(d.getDate());
    const h = pad2(d.getHours());
    const min = pad2(d.getMinutes());
    const s = pad2(d.getSeconds());
    return `${y}-${m}-${day} ${h}:${min}:${s}`;
  }

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

  function parseRecordLine(line) {
    const parts = line.split("\t").slice(0, 4).map(s => s.trim());
    while (parts.length < 4) parts.push("");
    return { mark: parts[0], name: parts[1], model: parts[2], origin: parts[3], status: "符合" };
  }

  function normalize(o) {
    if (!o) return "";
    if (/台|臺/.test(o)) return "台灣";
    if (/中|陸/.test(o)) return "中國大陸";
    return o;
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
    return {
      status2: document.querySelector('#addInspCaseA\\.inspExec\\.status2'),
      unmark: document.querySelector('#addInspCaseA\\.inspExec\\.unmark_no'),
      name: document.querySelector('#addInspCaseA\\.goods_name'),
      prd1: document.querySelector('#prd1_SCN3'),
      prd2: document.querySelector('#prd2_SCN3'),
      prd3: document.querySelector('#prd3_SCN3'),
      model: document.querySelector('#addInspCaseA\\.goods_modelno'),
      unmarkGoodsNo: document.querySelector('#addInspCaseA\\.inspExec\\.unmark_goods_no'),
      prodCountryHelpers: document.querySelectorAll('input#prod_country'),
      prodCountrySelect: document.querySelector('#addInspCaseA\\.prod_country'),
      conform: document.querySelector('#addInspCaseA\\.conform')
    };
  }

  async function fillForm(rec, cat) {
    const el = getForm();
    if (el.status2 && !el.status2.checked) el.status2.click();

    if (el.unmark && rec.mark) {
      el.unmark.value = rec.mark;
      el.unmark.dispatchEvent(new Event("blur", { bubbles: true }));
    }

    if (el.name && rec.name) {
      el.name.value = rec.name;
      el.name.dispatchEvent(new Event("change", { bubbles: true }));
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

    if (rec.model) {
      const mark = (rec.mark || "").trim();
      const modelTrim = rec.model.trim();

      const isMMark = /^M/i.test(mark);
      const isSevenDigits = /^\d{7}$/.test(modelTrim);

      if (isMMark && isSevenDigits && el.unmarkGoodsNo) {
        el.unmarkGoodsNo.value = modelTrim;
        el.unmarkGoodsNo.dispatchEvent(new Event("blur", { bubbles: true }));

        if (el.model) {
          el.model.value = "";
          el.model.dispatchEvent(new Event("blur", { bubbles: true }));
        }
      } else if (el.model) {
        el.model.value = rec.model;
        el.model.dispatchEvent(new Event("blur", { bubbles: true }));
      }
    }

    if (rec.origin) {
      const o = normalize(rec.origin);
      let code = "";
      if (/台/.test(o)) code = "TW";
      else if (/中/.test(o)) code = "CN";

      if (code) {
        if (el.prodCountryHelpers.length) {
          const input = code === "TW" ? el.prodCountryHelpers[0] : el.prodCountryHelpers[1];
          if (input) {
            input.value = code;
            input.dispatchEvent(new Event("keyup", { bubbles: true }));
            input.dispatchEvent(new Event("blur", { bubbles: true }));
          }
        }
        if (el.prodCountrySelect) {
          el.prodCountrySelect.value = code;
          el.prodCountrySelect.dispatchEvent(new Event("change", { bubbles: true }));
        }
      }
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
      <div id="mc-header">市場檢查小工具（CSV 版）</div>
      <div id="mc-body">
        <textarea id="mc-paste"></textarea>

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
        <button class="mc-btn" id="mc-btn-export">下載紀錄（CSV）</button>
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

  function bindPanelActions() {
    const pasteArea = document.getElementById("mc-paste");
    let pasteTimer;

    function parsePastedData(raw) {
      if (!raw || !raw.trim()) {
        saveQueue([]);
        saveIdx(0);
        refreshPanel();
        return;
      }

      const lines = raw.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      if (!lines.length) {
        saveQueue([]);
        saveIdx(0);
        refreshPanel();
        return;
      }

      const autoName = loadAutoName();
      const q = [];

      lines.forEach((line, i) => {
        const rec = parseRecordLine(line);
        rec.seq = String(i + 1);

        if (autoName) {
          const cat = guessCategoryFromName(rec.name);
          if (cat) rec.categoryKey = cat;
        }

        q.push(rec);
      });

      saveQueue(q);
      saveIdx(0);
      refreshPanel();
    }

    pasteArea.addEventListener("input", () => {
      clearTimeout(pasteTimer);
      pasteTimer = setTimeout(() => parsePastedData(pasteArea.value), 200);
    });

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
      const db = GM_getValue(DB_KEY, []);
      if (!db.length) return toast("資料庫為空");

      const header = "序號,檢查案號,商品名稱,商品檢驗標識,型號,產地,違規狀態,商品類別,建立時間";
      const lines = db.map(r =>
        [
          r.seq || "",
          r.caseNo || "",
          r.name || "",
          r.mark || "",
          r.model || "",
          r.origin || "",
          r.status || "",
          r.categoryKey || "",
          r.createdAt || ""
        ].join(",")
      );

      const csv = [header, ...lines].join("\n");
      const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "市場檢查_已完成紀錄.csv";
      a.click();
      URL.revokeObjectURL(url);

      toast("已下載紀錄 CSV");
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

      const now = formatDateTime(new Date());
      const record = {
        seq: rec.seq,
        caseNo,
        name: rec.name,
        mark: rec.mark,
        model: rec.model,
        origin: normalize(rec.origin),
        status: rec.status,
        categoryKey: rec.categoryKey,
        createdAt: now
      };

      const db = GM_getValue(DB_KEY, []);
      db.push(record);
      GM_setValue(DB_KEY, db);

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
