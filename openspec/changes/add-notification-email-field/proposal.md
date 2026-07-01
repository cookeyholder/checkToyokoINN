## Why

目前提醒通知信件固定寄到 Google Apps Script 執行者的 Email 信箱，無法讓使用者在網頁上指定通知信要寄到的 Email 地址。

## What Changes

- **新增「提醒收件 Email」欄位**：在「提醒清單」工作表的「使用者 Email」欄位之後新增一欄
- **新增通知 Email 輸入欄位**：在 Web 表單上增加一個 input，讓使用者輸入提醒信要寄送的 Email
- **後端改用新欄位寄送通知**：`sendNotification()` 改讀「提醒收件 Email」決定寄送對象
- **保留「使用者 Email」**：Session Email 欄位維持不變，繼續用於擁有權驗證
- **更新規格文件**：反映新欄位的存在與行為

## Capabilities

### New Capabilities

- _(無新能力，本次僅修改現有功能)_

### Modified Capabilities

- `spreadsheet-data`: 提醒清單工作表新增「提醒收件 Email」欄位（位於「使用者 Email」與「建立時間」之間）
- `web-ui`: 提醒建立表單新增「提醒收件 Email」文字輸入欄位；成功提交時會一併送出該值
- `email-notification`: 通知發送改為讀取「提醒收件 Email」；若該欄位為空則 fallback 到「使用者 Email」

## Impact

- `程式碼.js`: COLUMN_INDICES、addReminder、submitReminder、getReminders、getReminderList、sendNotification、initializeRemindersSheet 等函式需調整
- `index.html`: 表單 HTML 與 handleFormSubmit 需加入 email 欄位
- `openspec/specs/spreadsheet-data/spec.md`: 更新欄位定義
- `openspec/specs/web-ui/spec.md`: 更新表單欄位描述
- `openspec/specs/email-notification/spec.md`: 更新通知行為描述
