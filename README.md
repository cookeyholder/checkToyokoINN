# 東橫 INN 空房監控系統

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-V8%20Runtime-4285F4?logo=google&logoColor=white)](https://developers.google.com/apps-script)
[![Language](https://img.shields.io/badge/Language-JavaScript%20%28ES6%29-F7DF1E?logo=javascript&logoColor=black)](https://developer.mozilla.org/en-US/docs/Web/JavaScript)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

## 📌 專案簡介

**東橫 INN 空房監控系統** 是一個運行在 Google Apps Script 平台上的自動化工具,用於監控 [東橫 INN](https://www.toyoko-inn.com/china/) (日本連鎖商務旅館) 的房間可用性。系統能夠:

- ✅ **自動監控**:定期檢查指定分店和房型的房間可用性
- 📧 **智能通知**:當房間可用時,自動發送郵件通知給使用者
- ⏰ **時間篩選**:支援在特定時間範圍內進行監控
- 🏨 **多分店支援**:同時監控全國多間分店
- 📊 **詳細歷史**:記錄每次檢查結果和通知狀態

---

## 🎯 主要功能

### 1. 房間可用性監控
- 定期爬取東橫 INN 官方網站的房間資訊
- 支援多個分店、房型和日期的組合監控
- 實時查詢房間庫存和價格資訊

### 2. 智能提醒系統
- 使用者可建立自訂提醒規則
- 支援暫停/啟用/刪除提醒
- 自動篩選過期提醒

### 3. 郵件通知機制
- 房間有空時立即發送郵件
- 防止過度通知:同一提醒最多 1 小時通知一次
- 支援自訂通知模板和樣式

### 4. 資料管理和檢查歷史
- 完整記錄每次檢查結果
- 追蹤通知發送狀態
- 檢視錯誤日誌和除錯資訊

### 5. 網頁應用程式介面
- 友善的前端使用者介面
- 支援建立、編輯和管理提醒
- 檢視檢查歷史和統計資訊

---

## 🚀 快速開始

### 前置需求

- 🔑 **Google 帳戶**:使用 Google 帳戶
- 📊 **Google 試算表**:至少有編輯權限
- 🌐 **網路連線**:能夠存取東橫 INN 官方網站

### 安裝步驟

#### 方案一:從 Google Apps Script 編輯器部署

1. **建立新的 Apps Script 專案**
   - 開啟 [Google Apps Script 控制台](https://script.google.com)
   - 點選「新增專案」

2. **複製程式碼**
   - 複製 `程式碼.js` 的所有內容到 Apps Script 編輯器中的 `程式碼.gs`
   - 複製 `scrapers.gs` 的所有內容到新建的 `scrapers.gs` 檔案

3. **設定 HTML 前端**
   - 建立新的 HTML 檔案 (點選「+」按鈕)
   - 命名為 `index.html`
   - 複製 `index.html` 的內容

4. **設定 Apps Script 配置**
   - 複製 `appsscript.json` 的內容到 Apps Script 編輯器的「專案設定」

5. **部署為網頁應用程式**
   - 點選「部署」➜「新增部署」
   - 選擇「類型」為「網頁應用程式」
   - 「執行身份為」選擇你的 Google 帳戶
   - 「誰可以存取」選擇「任何人」
   - 點選「部署」

#### 方案二:使用 Google 試算表直接開啟

1. **建立 Google 試算表**
   - 開啟 [Google 試算表](https://sheets.google.com)
   - 建立新試算表

2. **啟用 Apps Script**
   - 選單「擴充功能」➜「Apps Script」
   - 進入 Apps Script 編輯器

3. 按照「方案一」的步驟 2-4 進行

4. **建立 Apps Script 觸發器**
   - 執行 `menuSetupTimeTrigger()` 以設定定期檢查任務

### 初始設定

在首次使用前,需要進行以下初始化:

```javascript
// 1. 初始化試算表 ID (自動執行)
setupWithCurrentSpreadsheet()

// 2. 初始化工作表結構
menuInitializeAllSheets()

// 3. 爬取東橫 INN 分店資料
menuScrapeToyokoInnBranches()

// 4. 新增使用者帳號
menuAddCurrentUser()

// 5. 設定定期檢查觸發器
menuSetupTimeTrigger()
```

這些功能都可以通過試算表中的「東橫INN房間監控」自訂選單存取。

---

## 📖 使用指南

### 建立提醒

1. **開啟試算表或網頁應用程式**
   - 在試算表中執行 Apps Script 部署的網頁應用程式,或直接開啟試算表

2. **填寫提醒資訊**
   - **分店名稱**:選擇欲監控的東橫 INN 分店
   - **房型**:選擇一個或多個房型 (支援複選)
   - **成人人數**:設定房間需求的人數
   - **房間數**:設定需要的房間數量
   - **入住日期**:選擇入住日期
   - **退房日期**:選擇退房日期
   - **提醒開始時間**:監控何時開始檢查 (格式:HH:MM)
   - **提醒結束時間**:監控何時停止檢查 (格式:HH:MM)
   - **通知郵件**:輸入用於接收通知的郵件地址

3. **建立提醒**
   - 點選「建立提醒」按鈕
   - 系統會為每個選定的房型建立一條提醒記錄

### 管理提醒

| 操作 | 說明 |
|------|------|
| **啟用** | 恢復提醒的監控,系統會定期檢查並在有空房時發送通知 |
| **暫停** | 停止監控,不發送任何通知 |
| **歷史** | 檢視該提醒的檢查記錄和通知發送歷史 |
| **刪除** | 永久刪除提醒記錄 |

### 檢視檢查歷史

- **檢查歷史**工作表記錄每次房間檢查的詳細資訊:
  - 檢查時間
  - 是否找到空房
  - 通知狀態
  - 錯誤訊息 (如有)

- 使用此資訊診斷問題和追蹤系統行為

### 進階功能

#### 時間篩選

- **提醒開始時間**:監控何時開始檢查
  - 例如:「08:00」表示每天 8 點開始監控
  
- **提醒結束時間**:監控何時停止檢查
  - 例如:「23:00」表示每天 23 點停止監控

- **支援跨日時間範圍**
  - 例如:「23:00」至「01:00」表示每天晚上 23 點到次日凌晨 1 點進行監控

#### 通知冷卻機制

- 同一提醒在 1 小時內最多發送一次通知
- 防止因房間資訊更新頻繁導致的多次通知轟炸

#### 批量操作

- 建立提醒時可同時選擇多個房型
- 系統會為每個房型自動建立獨立的提醒記錄

---

## 📊 資料表結構

### 核心工作表

#### 1. **提醒清單** (`Reminders`)
記錄所有使用者建立的監控提醒

| 欄位 | 說明 |
|------|------|
| 提醒 ID | 唯一識別符 (UUID) |
| 分店編號 | 東橫 INN 分店代碼 |
| 分店名稱 | 東橫 INN 分店名稱 |
| 房型代號 | 房型識別代碼 |
| 房型名稱 | 房型描述 |
| 成人人數 | 房間需求的成人人數 |
| 房間數 | 需要的房間數量 |
| 入住日期 | 入住日期 (YYYY-MM-DD) |
| 退房日期 | 退房日期 (YYYY-MM-DD) |
| 提醒開始時間 | 監控開始時間 (HH:MM) |
| 提醒結束時間 | 監控結束時間 (HH:MM) |
| 使用者 Email | 通知接收郵件地址 |
| 建立時間 | 提醒建立時間戳記 |
| 最後通知時間 | 最後一次發送通知的時間 |
| 通知狀態 | `success` / `pending` / `failed` |
| 提醒狀態 | `active` / `paused` / `expired` |

#### 2. **檢查歷史** (`CheckHistory`)
記錄每次房間可用性檢查的結果

| 欄位 | 說明 |
|------|------|
| 檢查 ID | 唯一識別符 |
| 提醒 ID | 關聯的提醒識別碼 |
| 檢查時間 | 檢查進行的時間戳記 |
| 檢查結果 | `available` / `unavailable` / `error` |
| 錯誤訊息 | 如有錯誤,記錄詳細訊息 |
| 是否發送通知 | `true` / `false` |

#### 3. **東橫 INN 分店** (`Branches`)
東橫 INN 所有分店的資料庫

| 欄位 | 說明 |
|------|------|
| 分店編號 | 東橫 INN 官方編號 |
| 分店名稱 | 分店名稱 |
| 地區 | 地理位置區域 |
| 都道府縣 | 日本的行政區劃 |
| 地址 | 完整地址 |
| 電話 | 分店電話 |

#### 4. **房型代號** (`RoomTypes`)
系統支援的房型和對應代碼

| 欄位 | 說明 |
|------|------|
| 房型代號 | API 使用的房型代碼 (10, 20, 30, 40) |
| 房型名稱 | 房型描述 (單人房、雙人房等) |

#### 5. **網站參數設定** (`Parameters`)
爬蟲和 API 呼叫的設定參數

| 欄位 | 說明 |
|------|------|
| 參數項目 | 參數名稱 |
| 參數內容 | 參數值 |
| 參數說明 | 參數用途說明 |

#### 6. **帳號管理** (`Accounts`)
系統使用者帳號資訊

| 欄位 | 說明 |
|------|------|
| Email | 使用者郵件地址 |
| 姓名 | 使用者姓名 |
| 人員編號 | 內部識別號 |
| 部門單位 | 所屬部門 |
| 群組 | 使用者群組 |
| 狀態 | `active` / `inactive` |
| 備註 | 其他說明 |

---

## 🏗️ 系統架構

### 程式碼組織

系統採用 **分階段函式導向架構**,按功能劃分為 9 個階段:

```
第 1 階段: 試算表整合 (1.1-1.14)
  ├─ CRUD 操作
  └─ 使用者驗證

第 2 階段: 房間檢查器 (2.1-2.6)
  ├─ API 查詢
  └─ HTML 解析

第 3 階段: 提醒管理系統 (3.1-3.8)
  ├─ 時間篩選
  ├─ 冷卻機制
  └─ 主循環

第 4 階段: 郵件通知 (4.1-4.7)
  ├─ HTML 生成
  └─ GmailApp 整合

第 5 階段: 網頁應用程式介面 (5.1-5.8)
  ├─ doGet()
  └─ API 處理器

第 6 階段: 資料驗證 (6.1-6.8)
  ├─ 輸入驗證
  └─ 錯誤訊息

第 7 階段: 權限與部署 (7.1-7.4)
  ├─ OAuth 驗證
  └─ 初始化

第 8 階段: 測試 (8.1-8.16)
  ├─ 單元測試
  └─ 整合測試

第 9 階段: 文件和最佳化 (9.1-9.9)
  ├─ 文件
  ├─ 指南
  └─ 最佳化
```

### 檔案結構

```
checkToyokoINN/
├── 程式碼.js              # 主要業務邏輯和菜單功能
├── scrapers.gs            # 東橫 INN 官網爬蟲和數據採集
├── index.html             # 前端網頁應用程式介面
├── email_styles.html      # 郵件通知的 HTML 樣式
├── appsscript.json        # Apps Script 配置文件
├── README.md              # 本檔案 - 專案說明文檔
├── USER_GUIDE.md          # 詳細使用者指南
├── ARCHITECTURE.md        # 系統架構文檔
├── AGENTS.md              # 開發指南
└── openspec/              # OpenSpec 規範和變更追蹤
    ├── project.md         # 專案上下文
    ├── AGENTS.md          # OpenSpec 指令
    └── specs/             # 各模塊的詳細規範
        ├── room-checker/
        ├── reminder-management/
        ├── email-notification/
        ├── spreadsheet-data/
        └── web-ui/
```

### 關鍵函式流程

#### 主監控循環 (`checkRemindersAndNotify`)

```
1. 取得所有啟用的提醒
    ↓
2. 對每個提醒檢查時間篩選
    ↓
3. 構建查詢 URL 並查詢房間可用性
    ↓
4. 解析 HTML 回應判斷有無空房
    ↓
5. 如有空房且滿足通知條件
    ├─ 生成郵件內容
    ├─ 發送郵件通知
    └─ 記錄通知狀態
    ↓
6. 更新提醒的最後通知時間
    ↓
7. 記錄檢查歷史
```

#### 房間查詢流程 (`checkRoomAvailability`)

```
1. 構建查詢參數 (分店編號、房型、人數、日期等)
    ↓
2. 生成查詢 URL
    ↓
3. 發送 HTTP GET 請求
    ↓
4. 解析 HTML 回應
    ↓
5. 檢查是否有可預訂房間
    ↓
6. 回傳布林值或房間詳情
```

---

## ⚙️ 技術細節

### 技術棧

- **執行環境**: [Google Apps Script](https://developers.google.com/apps-script) (V8 Runtime)
- **程式語言**: JavaScript (ES6)
- **資料存儲**: [Google 試算表](https://sheets.google.com)
- **郵件服務**: Gmail API (通過 GmailApp)
- **網頁爬蟲**: UrlFetchApp + HTML 解析

### 依賴項

本專案使用 Google Apps Script 內建服務,無外部相依套件:

- **SpreadsheetApp**: 操作 Google 試算表
- **UrlFetchApp**: 發送 HTTP 請求到東橫 INN API
- **GmailApp**: 發送郵件通知
- **PropertiesService**: 儲存使用者設定和狀態
- **CacheService**: 快取常用資料以提高效能
- **Utilities**: 工具函式 (編碼、雜湊等)
- **ScriptApp**: 建立時間觸發器

### 配置信息

系統在 `appsscript.json` 中配置:

```json
{
  "timeZone": "Asia/Taipei",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/userinfo.email"
  ],
  "webapp": {
    "access": "ANYONE_ANONYMOUS",
    "executeAs": "USER_DEPLOYING"
  }
}
```

### Google Apps Script 配額

系統設計時已考慮 Google Apps Script 的配額限制:

| 資源 | 限制 | 策略 |
|------|------|------|
| **執行時間** | 6 分鐘/執行 | 批量查詢,定期執行 |
| **郵件服務** | 100/天 (免費) | 最多 1 小時通知一次 |
| **URL Fetch** | ~20,000/天 | 合理分散查詢 |
| **試算表讀寫** | 無硬限制 | 最小化 getValues() 呼叫 |

---

## 🔧 開發指南

### 設定開發環境

1. **在 Google Apps Script 編輯器中開啟專案**
   - 進入 [Google Apps Script 控制台](https://script.google.com)
   - 開啟現有專案或建立新專案

2. **使用內建編輯器或同步本地文件**
   - 推薦使用 [clasp](https://github.com/google/clasp) 工具進行本地開發

3. **測試函式**
   - 在編輯器中選擇函式並點選「執行」
   - 查看執行日誌了解詳細信息

### 編碼規範

- **命名規範**: 使用駝峰命名法 (`functionName`, `variableName`)
- **註解**:使用繁體中文註解說明邏輯
- **函式設計**: 單一職責,避免過度複雜
- **錯誤處理**: 使用 try-catch 和有意義的錯誤訊息
- **效能**: 最小化試算表操作,合理使用快取

### 除錯技巧

1. **查看執行日誌**
   ```javascript
   Logger.log("Debug message:", value);
   ```

2. **使用提示框進行測試**
   ```javascript
   const ui = SpreadsheetApp.getUi();
   ui.alert("Message", "Content", ui.ButtonSet.OK);
   ```

3. **試算表中檢查資料**
   - 查看各工作表中的資料
   - 驗證資料格式和內容

4. **監控執行時間**
   - Google Apps Script 每次執行最多 6 分鐘
   - 監控長時間運行的操作

### 常見問題排查

#### 問題: 未能連線到東橫 INN API

**可能原因:**
- 網路連接問題
- 官方 API 已變更或下線
- IP 被限制或速率限制

**解決方案:**
- 檢查網路連接
- 查看爬蟲日誌了解失敗原因
- 嘗試使用代理或備用 API 端點

#### 問題: 郵件未發送

**可能原因:**
- Gmail 配額已達限制
- 郵件地址錯誤
- 權限設定不正確

**解決方案:**
- 檢查 Gmail 每日配額
- 驗證使用者郵件地址
- 確認 OAuth 權限設定

#### 問題: 試算表操作緩慢

**可能原因:**
- 工作表資料過多
- 過多的 getValues() 呼叫
- 網路連接較慢

**解決方案:**
- 最小化試算表操作
- 使用快取 (CacheService)
- 定期歸檔舊資料

---

## 📈 效能最佳化

### 已實現的最佳化策略

1. **批量讀取試算表**
   - 一次讀取多筆資料,而非逐筆查詢

2. **選擇性欄位讀取**
   - 只讀取所需欄位,減少資料傳輸

3. **試算表物件快取**
   - 快取常用工作表物件,避免重複初始化

4. **及時清理大型物件參考**
   - 釋放不再需要的物件參考,節省記憶體

5. **監控執行時間配額**
   - 記錄執行時間,避免超出 6 分鐘限制

### 進一步最佳化建議

- [ ] 實現增量同步 (只同步更改的資料)
- [ ] 使用 Firestore 替代試算表存儲高頻資料
- [ ] 實現多個獨立的觸發器以並行處理
- [ ] 利用 CacheService 存儲查詢結果

---

## 🔒 安全性考慮

### OAuth 權限

系統申請的最小必要權限:

```javascript
"oauthScopes": [
  "https://www.googleapis.com/auth/spreadsheets",           // 試算表讀寫
  "https://www.googleapis.com/auth/script.external_request", // HTTP 請求
  "https://www.googleapis.com/auth/gmail.send",              // 發送郵件
  "https://www.googleapis.com/auth/script.scriptapp",       // 觸發器管理
  "https://www.googleapis.com/auth/userinfo.email"          // 取得郵件地址
]
```

### 資料安全建議

1. **定期備份試算表**
   - 定期匯出重要資料
   - 保留版本歷史

2. **限制試算表存取**
   - 只授予必要的編輯權限
   - 使用 Google Drive 的共用設定

3. **郵件隱私**
   - 不在郵件中洩露敏感信息
   - 定期清理檢查歷史

4. **API 端點安全**
   - 驗證 HTTPS 連接
   - 遵守目標網站的爬蟲政策和服務條款

---

## 📝 文件和資源

- [USER_GUIDE.md](./USER_GUIDE.md) - 詳細使用者指南
- [ARCHITECTURE.md](./ARCHITECTURE.md) - 系統架構文檔
- [AGENTS.md](./AGENTS.md) - 開發指南和最佳實踐
- [Google Apps Script 官方文檔](https://developers.google.com/apps-script)
- [Google 試算表 API 文檔](https://developers.google.com/sheets/api)
- [Gmail API 文檔](https://developers.google.com/gmail/api)

---

## 🤝 貢獻指南

我們歡迎貢獻!以下是貢獻流程:

1. **提交功能請求或報告問題**
   - 使用 GitHub Issues 提交 (如適用)
   - 清楚地描述需求或問題

2. **提交程式碼變更**
   - 遵循編碼規範
   - 添加合適的註解和文件
   - 使用 OpenSpec 工作流程管理重大變更

3. **更新文件**
   - 更新相關文件以反映變更
   - 確保文件保持同步

---

## 📄 許可證

本專案採用 MIT 許可證。詳見 [LICENSE](LICENSE) 檔案。

---

## 💡 常見問題解答

### Q: 系統多久檢查一次房間?

A: 預設配置下,系統每 5 分鐘檢查一次啟用的提醒。可以通過調整觸發器設定改變檢查間隔。

### Q: 可以監控多少個提醒?

A: 理論上沒有限制,但受 Google Apps Script 的執行時間配額 (6 分鐘) 影響。大約可以在 6 分鐘內檢查 500-1000 個提醒,具體取決於網路速度和系統負載。

### Q: 房間有空時會立即發送通知嗎?

A: 系統會在下次定期檢查時發現有空房,然後發送通知。通常在 5 分鐘內發送。

### Q: 可以修改郵件通知的內容嗎?

A: 可以。郵件模板在 `email_styles.html` 中定義。編輯該檔案可自訂郵件樣式和內容。

### Q: 系統支援哪些房型?

A: 系統支援東橫 INN 的所有標準房型:
- 10: 單人房
- 20: 雙人房
- 30: 雙床房
- 40: 三人以上房

### Q: 如何刪除所有資料並重新開始?

A: 
1. 選擇「初始化所有工作表」會清空並重新建立所有工作表
2. 注意:此操作無法撤銷

### Q: 系統是否會爬蟲東橫 INN 的網站?

A: 是的。系統會定期從東橫 INN 官方 API 和網站爬取分店資訊和房間可用性。所有爬蟲操作遵守禮貌爬蟲原則和網站的服務條款。

---

## 📞 支援和回饋

如有任何問題或建議:

1. 查看 [USER_GUIDE.md](./USER_GUIDE.md) 的常見問題部分
2. 檢查 [ARCHITECTURE.md](./ARCHITECTURE.md) 了解系統設計
3. 查看執行日誌和檢查歷史進行除錯
4. 參考 openspec 文件夾中的詳細規範

---

## 🔄 更新日誌

### v1.0.0 (2025-11-13)

- ✨ 初始版本
- ✅ 完成房間監控基本功能
- ✅ 實現郵件通知機制
- ✅ 建立提醒管理系統
- ✅ 開發網頁應用程式介面
- ✅ 添加詳細文件和指南

---

<div align="center">

**東橫 INN 空房監控系統** 由 AI 助手和開發團隊共同開發

[回到頂部](#東橫-inn-空房監控系統)

</div>
