# AGENTS.md – momo_check 專案最高指導文件 v1

本文件是 **momo_check 系列工具** 的最高指引，  
適用於 **人類開發者** 以及 **AI 助手（例如 GitHub Copilot / OpenAI Codex / ChatGPT）**。

請在修改任一版本程式碼前，先閱讀並遵守本文件。

---

## 1. 專案總覽

專案名稱：`momo_check`（位於 `momo_tools/` 資料夾下）

目的：

- 從 CSV 讀取 momo 商品網址（含「序號」「商品網址」）
- 以 HTTP 爬蟲抓取商品資訊
- 依「查核紀錄彙整表」格式輸出 Excel，供公務查核紀錄使用

主要執行環境：

- Google Colab / Python 3.x
- 使用者會上傳 CSV，執行程式，下載 Excel

---

## 2. 不可破壞的核心原則（Hard Rules）

1. **不得整支重寫**
   - 新功能或大改版一律在現有檔案上「疊高樓」，保持程式架構與風格。
   - 禁止丟棄舊版邏輯重新寫一支完全不同的腳本。

2. **版本命名固定為 v1 / v2 / v3 / v4...**
   - 僅用簡單整數版本，例如：`momo_check_v1.py`、`momo_check_v2.py`、`momo_check_v3.py`、`momo_check_v4.py`。
   - 新版本檔名必須保留舊版本檔案，不得覆蓋舊檔。

3. **`parse_momo_simple()` 是核心爬蟲，不可移除**
   - 可在其內部「增強」功能（例如多抓欄位、調整正則），
   - 但不可隨意改變其回傳結構的基本 key：`商品名稱`、`品號`、`商檢字號`、`網址`。

4. **Excel 欄位順序不得更動**
   - 匯出的 DataFrame 欄位順序必須固定（見第 4 節）。
   - 即使新增功能，也應在內部重用現有欄位，不隨意新增欄位或改名。

5. **對 momo 伺服器要友善**
   - 每處理一個商品網址後必須 `time.sleep(5)` 秒以上。
   - 不得移除或縮短 sleep 以密集攻擊網站。

---

## 3. 目前程式版本說明

### 3.1 `momo_check_v1.py`（歷史版本）

- 早期版本，Word 報告為主，流程依賴 python-docx。
- 已退居歷史參考用，不建議再改動。

### 3.2 `momo_check_v2.py`

- 改為 **CSV 匯入 → Excel 匯出**。
- 保留原本 `parse_momo_simple()` 的爬蟲邏輯。
- 不處理商檢字號。

### 3.3 `momo_check_v3.py`（本 AGENTS 首次建立時的基準版）

- 在 v2 基礎上新增欄位：
  - 從商品頁抓取「商檢字號」，寫入 Excel「商檢標識」欄位。
- Excel 匯出改用 openpyxl 調整欄寬、設定換行。

### 3.4 `momo_check_v4.py`（預期行為）

- 將 `parse_momo_simple()` 中「商檢字號抓取」邏輯升級：
  - 優先從 `#panel-2` 規格表中尋找「商檢字號」那一列。
  - 再從 `#panel-2` 全文字搜尋。
  - 最後才對整頁全文做 fallback。
- 商檢字號格式放寬為：**M/R/D/T + 5 個英數字**。

---

## 4. 欄位與資料映射規則

### 4.1 `parse_momo_simple()` 回傳格式

函式簽名：

```python
def parse_momo_simple(url: str, max_retries: int = 3) -> Dict[str, str]
必須至少回傳以下欄位（key 名稱固定）：

"商品名稱"：momo 商品名稱（通常來自 og:title）

"品號"：商品品號（例如 TP00050990000014）

"商檢字號"：商檢標識（例如 R3G804，若抓不到則為空字串）

"網址"：原始商品網址

錯誤時：

"商品名稱" 可以是 "錯誤：{訊息}" 型式，

"品號" 可為 "錯誤"，

"商檢字號" 為 ""，

"網址" 照填。

4.2 Excel 欄位順序（從左到右）
DataFrame 欄位順序必須為：

編號（來自 CSV「序號」）

檢查案號

查核日期

網路名稱/店家名稱

賣家帳號或拍賣代碼

商品名稱

再查核日期

是否下架

是否改正

調查結果

網址/地址

商檢標識

已宣導

已下架

4.3 欄位填寫規則
對每一筆輸入「商品網址」：

編號：輸入 CSV 的「序號」

檢查案號：空字串

查核日期：當天日期轉 ROC，格式 YYY/M/D（由 to_roc_date() 計算）

網路名稱/店家名稱：固定填 "momo購物網"

賣家帳號或拍賣代碼：填入 parse_momo_simple() 抓到的 "品號"

商品名稱："商品名稱" 內容（若是錯誤格式則清空）

再查核日期：空字串

是否下架：空字串

是否改正：空字串

調查結果：

正常抓到商品名稱時留空。

若 "商品名稱" 以 "錯誤：" 開頭，去掉前綴後寫入此欄位。

網址/地址：原始商品網址

商檢標識：填入 "商檢字號"（已轉為大寫，如 R3G804）；抓不到則空字串。

已宣導：空字串

已下架：空字串

5. 商檢字號規則與抓取策略
5.1 字軌格式
商檢字號必須符合：

第一碼：M、R、D、T 其中一個（不分大小寫）

後面：5 碼英數字（A–Z 或 0–9）

總長度：6 字元

建議 regex（Python）：

python
複製程式碼
MRDT_REGEX = re.compile(r"\b[MRDT][A-Za-z0-9]{5}\b", re.IGNORECASE)
寫入 Excel 前請統一轉為大寫。

5.2 抓取優先順序（v4 之後必須遵守）
商品規格 tab（panel-2）表格中的「商檢字號」列

在 div#panel-2 下尋找文字為「商檢字號」的 label div。

取其右側兄弟節點（value div）中的文字，用 MRDT_REGEX 搜尋。

整個 panel-2 內容全文掃描

若 1 抓不到，對 panel-2 的 get_text() 使用 MRDT_REGEX.search()。

全頁 fallback

若仍抓不到，再對整個 soup 的文字做一次 MRDT_REGEX.search()。

三層都沒 match → 商檢字號 回傳空字串。

6. Excel 匯出規範
6.1 檔名格式
python
複製程式碼
filename = f"momo_check_output_ROC{roc_date.replace('/', '')}.xlsx"
例如 ROC 114 年 12 月 1 日 → momo_check_output_ROC1141201.xlsx

6.2 欄寬與換行
使用 openpyxl，匯出後需調整：

欄寬（由左至右），建議值：

欄位	寬度
編號	6
檢查案號	15
查核日期	12
網路名稱/店家名稱	16
賣家帳號或拍賣代碼	18
商品名稱	30
再查核日期	12
是否下架	10
是否改正	10
調查結果	30
網址/地址	40
商檢標識	14
已宣導	10
已下架	10

開啟自動換行（wrap_text = True）至少套用在：

商品名稱

調查結果

網址/地址

7. 給 AI / Codex 的開發指引
修改方式

新功能、新抓法、新欄位邏輯 → 請建立新版本檔案 momo_check_vX.py，不要覆寫舊檔。

在新版本中盡量保留舊版函式名稱與參數，維持相容。

必須保留的函式與介面

parse_momo_simple(url, max_retries=3) -> Dict[str, str]

fetch_momo_product(url) -> Dict[str, str]

load_input_csv() -> pd.DataFrame

build_output_rows(df, roc_date) -> pd.DataFrame

export_to_excel(df, roc_date) -> str

to_roc_date(date_obj) -> str

main() + if __name__ == "__main__": main()

當不確定使用者需求時

寧可留下原本行為不變，只在新版本檔案內增加功能，

不要刪除原有欄位、改名、或隨意更動輸出格式。

錯誤處理

對 momo 連線失敗時，務必捕捉例外並回傳錯誤訊息至 調查結果 欄位，

不可讓整批執行因單一商品失敗而中止。

8. 版本變更摘要（簡表）
v1：Word 報告版（python-docx），文字貼上 → Word 報告。

v2：改為 CSV 匯入 → Excel 匯出，去 Word 化。

v3：在 v2 基礎上新增商檢字號欄位，簡單從全文用正則 [MR]\d{5} 抓字軌。

v4：商檢字號改為 M/R/D/T + 5 英數字，改為優先從 panel-2 規格表抓，再從 panel-2 全文與全頁 fallback。

未來版本請在此處追加簡短說明。

yaml
複製程式碼

---

你可以先把這版 `AGENTS.md` 貼進 repo，之後如果我們再加功能（例如支援蝦皮、露天、或加自動分案邏輯），就一起在這份
