# Project Context

## Purpose
自動檢查東橫 INN (Toyoko Inn) 飯店的可用房間資訊和價格,協助使用者追蹤訂房機會。

## Tech Stack
- **Google Apps Script (GAS)** - 主要執行環境
- **JavaScript (ES5/ES6)** - 程式語言
- **V8 Runtime** - JavaScript 執行引擎
- **Google Services** - 可能整合的服務 (例如: Sheets, Gmail, Calendar)

## Project Conventions

### Code Style
- 使用駝峰命名法 (camelCase) 作為函式名稱
- 使用繁體中文註解說明程式邏輯
- 保持程式碼簡潔,優先選擇清晰易讀的寫法
- 函式應有明確的單一職責

### Architecture Patterns
- 採用簡單的函式導向程式設計
- 避免過度抽象,優先考慮可維護性
- 使用 Google Apps Script 內建服務而非外部函式庫
- 適當使用快取減少 API 呼叫次數

### Testing Strategy
- 在 GAS 編輯器中手動測試主要功能
- 使用 Logger.log() 進行除錯和追蹤
- 測試不同的錯誤情境和邊界條件

### Git Workflow
- 使用有意義的提交訊息 (繁體中文)
- 在本地編輯後同步到 GAS 專案
- 使用 OpenSpec 工作流程管理重大變更

## Domain Context
- **東橫 INN**: 日本連鎖商務旅館品牌
- **訂房系統**: 需要了解飯店預訂網站的結構和 API (如果有)
- **房型資訊**: 不同房型、價格、可用日期等資料
- **時區**: 專案設定為 Asia/Taipei (台北時區)

## Important Constraints
- **GAS 配額限制**: 注意執行時間、URL Fetch 次數等限制
- **無外部相依套件**: 目前沒有使用任何外部函式庫
- **V8 Runtime**: 支援現代 JavaScript 功能,但仍有 GAS 環境的特定限制
- **網路請求**: 需遵守目標網站的爬蟲政策和速率限制

## External Dependencies
- **東橫 INN 官方網站/API**: 房間資訊的資料來源
- **Google Apps Script Services**: 
  - UrlFetchApp (HTTP 請求)
  - SpreadsheetApp (可能用於記錄資料)
  - GmailApp (可能用於通知)
  - PropertiesService (儲存設定)
  - CacheService (快取資料)
  - Utilities (工具函式)
