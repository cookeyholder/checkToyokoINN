## Why

目前當程式碼新增欄位（如先前加入 `notificationEmail`），若使用者未手動在試算表插入對應欄位，`getReminders()` 會將舊資料「建立時間」讀成「提醒收件 Email」，造成資料讀寫嚴重錯位。需要自動偵測處理而非完全依賴使用者手動 migration。

## What Changes

- 在 `initializeRemindersSheet()` 或讀取提醒前，動態偵測工作表欄位數
- 若欄位數小於新版預期值，自動在對應位置插入新欄位
- 保留既有資料列的欄位值，不破壞舊記錄

## Capabilities

### New Capabilities
- _(無新能力)_

### Modified Capabilities
- `reminder-management`: 提醒清單工作表初始化增加自動欄位 migration 功能

## Impact

- `程式碼.js`: `initializeRemindersSheet()` 新增自動補欄邏輯；或新增 `migrateRemindersSheet()` 獨立函式
