## Why

批次建立提醒時，`submitReminder()` 和 `addBatchReminders()` 在迴圈中逐筆呼叫 `sheet.appendRow()`，造成多次試算表 I/O 請求，在選取多個房型時易發生 HTTP timeout。此外 `addBatchReminders` 與 `submitReminder` 有大量重複邏輯且未被使用，增加維護成本。

## What Changes

- **批次寫入優化**：`submitReminder()` 改為收集所有新資料列後以 `setValues()` 一次寫入
- **移除 `addBatchReminders()`**：該函式未被任何程式碼呼叫且已遺漏 `notificationEmail` 欄位，移除以避免死碼維護負擔
- **統一批量寫入流程**：移除 `addReminder()` 內部無意義的 `SpreadsheetApp.flush()`（已於前次修正完成）

## Capabilities

### New Capabilities
- _(無新能力)_

### Modified Capabilities
- `reminder-management`: 批次新增提醒的實作方式改為 `setValues()` 一次寫入，不再逐筆 `appendRow`

## Impact

- `程式碼.js`: `submitReminder()` 改為批次 `setValues()`；`addBatchReminders()` 整段移除；`addReminder()` 維持現狀（已無 flush）
