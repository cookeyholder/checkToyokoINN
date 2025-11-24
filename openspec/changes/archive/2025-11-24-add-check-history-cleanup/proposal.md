# Change Proposal: add-check-history-cleanup

## Why
檢查歷史工作表會隨著系統長期運行而不斷累積記錄,當記錄超過 10 萬列時可能影響試算表效能和讀寫速度。需要自動清理機制來保持工作表在可控範圍內。

## What Changes
- 在 `checkAllReminders()` 函式執行完所有提醒檢查並寫入檢查歷史後,新增自動清理邏輯
- 檢查「檢查歷史」工作表的總列數
- 如果列數超過 100,000 列,刪除第 2 列到第 97,000 列的資料 (保留標題列和最近約 3,000 筆記錄)
- 在 Logger 中記錄清理操作的執行情況

## Impact
- **Affected specs**: `spreadsheet-data`
- **Affected code**:
  - `程式碼.js:2530` - `checkAllReminders()` 函式需要新增清理邏輯
  - `程式碼.js:1459` - `logCheckHistory()` 函式的行為保持不變
- **Benefits**:
  - 防止試算表效能下降
  - 自動化維護,無需手動清理
  - 保留最近約 3,000 筆記錄供查詢
- **Risks**:
  - 舊記錄會被永久刪除 (可考慮先備份)
  - 刪除操作需要額外的執行時間
