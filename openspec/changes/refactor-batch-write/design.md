## Context

目前 `submitReminder()` 對每個房型逐筆呼叫 `sheet.appendRow()`，建立 N 筆提醒就進行 N 次試算表 I/O。當使用者選取多個房型（常見：3-4 種）時，加上其他外部位請求（官網查房），HTTP 執行時間容易超過 Google Apps Script 的 30 秒 timeout。

同時 `addBatchReminders()` 存在相同問題且有重複邏輯，但實際未被任何程式碼呼叫。

## Goals / Non-Goals

**Goals:**
- `submitReminder()` 改為收集所有資料列後一次 `setValues()` 寫入
- 移除未使用的 `addBatchReminders()` 函式
- 改善批次新增提醒時的效能與 timeout 風險

**Non-Goals:**
- 不改變 `addReminder()` 的單筆寫入界面（仍可被其他程式碼呼叫）
- 不改變 `sendNotification()`、`getReminders()` 等讀取邏輯
- 不修改試算表欄位結構

## Decisions

**決策一：收集陣列後一次性 `setValues()`**
- **理由**：將 N 次 `appendRow` 降為 1 次 `setValues`，減少 N-1 次試算表 I/O。
- 替代方案 `BatchUpdate` API：GAS 基層無法直接呼叫 Sheets API batchUpdate，`setValues` 是最佳解。

**決策二：移除 `addBatchReminders()`**
- **理由**：該函式僅在 `addBatchReminders` 本身定義中被提及，無任何呼叫點，且已遺漏 `notificationEmail` 欄位。
- 替代方案保留並修正：仍需同步維護重複邏輯，增加後續修改成本。

## Risks / Trade-offs

- **大於 1 筆寫入錯誤** → 若 `setValues` 因單列資料錯誤而失敗，整批都不寫入。風險極低（資料來自已驗證的表單輸入）。
- **現有相依 `addReminder`** → 保留 `addReminder` 不變，任何現有或未來的單筆呼叫不受影響。
