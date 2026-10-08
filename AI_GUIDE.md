# 🤖 AI 協作開發指南 (AI Guide & Guardrails)

本文件是專門為**未來接手本專案的 AI 協作助手（如 Claude, ChatGPT, Gemini, Cursor 等）**所準備的工程規格與防踩坑（Guardrails）指南。

在修改本專案任何程式碼前，**請務必詳讀本指南**，遵守既有的業務邏輯與邊界防護，避免改動引發歷史迴歸（Regression）問題。

---

## 🏗️ 1. 專案核心定位與架構

1. **純本機、零外部 API（Zero-LLM Runtime）**：
   - 本系統所有比對邏輯在使用者本機瀏覽器內 100% 透過純 JavaScript 與演算法執行。
   - **嚴禁在 `app.js` 中引入任何外部第三方 AI API、雲端連線或未授權之網路請求**，確保企業內訓報價與機密檔案 100% 安全。
2. **核心檔案分工**：
   - `index.html`：單頁前端介面（TailwindCSS 離線版）。
   - `app.js`：核心解析與比對引擎（包含 Word 解析、PDF 版面幾何還原、LCS 差異演算法、紅黃綠燈號判定）。
   - `lib/`：本地依賴函式庫（`jszip.min.js`, `pdf.min.js`, `pdf.worker.js`）。
   - `verify_10_catalogs.js`：全量 10 大系列手冊（132 門課程）自動化回歸檢驗腳本。
   - `CHANGELOG.md`：版本更新歷程紀錄。

---

## 🛡️ 2. 核心比對業務規則與防踩坑清單 (Crucial Guardrails)

修改程式碼時，**絕對不能違反**以下 8 大核心規則：

### 規則 1：課名代碼前綴嚴格保護（AIPM 案例）
* **問題背景**：PDF 美編常在左側排有代碼色塊（如 `AIPM`），但在右側標題又多打了一次代碼（如 `AIPM AI時代的專案經理...`）。
* **防護邏輯**：
  * **嚴格禁止在 PDF 標題解析時無條件濾除開頭代碼**！
  * 只有在代碼後緊接冒號標籤（例如 `AZ-900：`）時才允許剝離。
  * 若 PDF 標題多排代碼且後面接空格或文字，必須保留原始文字，讓比對引擎抓出 🟡 **「中文課名文字微差 (PDF 多排『xxx』)」**，絕不可幫美編自動吃掉錯誤。

### 規則 2：縮寫字首結巴去重（CCN CCNA 案例）
* **問題背景**：Word 原稿因排版負邊距凸排（hanging indent），有時在儲存格開頭多打了殘留縮寫（如 `CCN  CCNA Exam Preparation and Workshop`）。
* **防護邏輯**：
  * 當英文課名開頭字詞是緊鄰下一個完整字詞的前綴（如 `CCN` 是 `CCNA` 的前綴，或 `AWS AWS` 重複），自動智慧收攏為單一完整單字。

### 規則 3：項目符號形態差異嚴格提醒（黃色警示）
* **規則**：
  * 當 Word 原稿採用數字編號（`1. 2. 3.` 或 `1234`），而 PDF 美編排版改用圓點（`•` 或其他項目符號）時：
  * 雖然文字內文相符，但**必須標記為 🟡 黃色提醒**：`內文相符，但項目符號形式不同 (Word: 編號 ⇄ PDF: 圓點)`。

### 規則 4：章節範圍與實驗編號保護（Lab 10-11 案例）
* **規則**：
  * 大綱或清單中包含如 `Lab 10-11`、`Day 1-2`、`Module 1-3` 等範圍連字號時，**嚴禁將連字號後方文字切斷為新清單項目**，必須完整保留為單一主題。

### 規則 5：停開課程與刪除線防護（90% 刪除線規則）
* **規則**：
  * 當 PM 停開某門課程時，常在 Word 整張表格畫上刪除線（`<w:strike>` / `<w:dstrike>`）。
  * 若表格內刪除線字元比例達 **90% 以上**，系統直接將該課視為已廢棄課程排除，**不得報警 PDF 漏排**。

### 規則 6：宣傳品版面簡化與容錯規則
* **規則**：
  * **『學會技能』與『課程目標』**：美編排版若版面不足未排入 PDF，標記為灰色提示（正常），**絕不可報紅燈錯誤**。
  * **『後續推薦課程』**：Word 原稿常列 3~5 門，美編依規則通常僅排首門。**只要 PDF 排的首門推薦課程與 Word 第一門相符，即判定為通過（綠燈）**。

### 規則 7：原生 Word 多階層編號（numbering.xml）
* **規則**：
  * Word 清單編號來自 `word/numbering.xml` 階層樹與 `%1..%9` 格式，各課程獨立表格計數器必須獨立重置，防止上一張表的序號外溢至下一張表。

### 規則 8：三欄式緊湊版面適應（3-Column Gutter Detection）
* **規則**：
  * 區塊鏈等緊湊手冊採用三欄排版，欄間距判定需具備自適應欄寬（Gutter Detection），避免跨欄內容張冠李戴。

---

## 🧪 3. 測試與驗證體系 (Verification & Testing)

每次修改 `app.js` 後，**必須執行全量回歸測試**，確保既有的 10 本手冊維持 0 誤報。

### 執行回歸測試指令
在終端機執行：
```powershell
node verify_10_catalogs.js
```
*(若在 Windows 本機無獨立 node 環境，可透過 VS Code 內建之 Electron node 執行：`$env:ELECTRON_RUN_AS_NODE=1; & 'C:\Users\lexiewang\AppData\Local\Programs\Microsoft VS Code\Code.exe' verify_10_catalogs.js`)*

### 驗證通過標準 (Acceptance Criteria)
執行測試後輸出必須符合：
1. **Total Red Fields: 必須為 0**（不允許任何嚴重紅燈誤報）。
2. **Total Catalogs Tested: 10/10 全部通過**（資安、Cisco、RedHat、Python、MSSE、Power Platform、OffSec、OpenSource、Oracle DB、Oracle Java）。

---

## 📦 4. 版本發布與提交規範

1. **版本號同步**：
   * 在 `index.html` 的版本徽章更新版本號（例如 `v3.5.2` -> `v3.5.3`）。
2. **記錄更新日誌**：
   * 在 `CHANGELOG.md` 頂部依據 Keep a Changelog 格式詳細記載本次修復項目（`Fixed`、`Added`、`Changed`）。
3. **提交與推送**：
   * 使用語意化 Commit Message（例如 `fix(parser): ...`、`feat(alignment): ...`）。
   * 提交後推送到 GitHub `origin main`。
