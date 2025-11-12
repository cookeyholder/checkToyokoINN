## Why
目前的提醒管理流程仍然依賴 rowIndex 定位試算表列。當資料列被重新排序、插入或刪除時，rowIndex 容易漂移並導致刪除/狀態切換等操作指向錯誤的提醒。雖然我們已經在資料列中加入 UUID，前端與後端仍未完整改用 UUID 進行操作，存在資料不一致與歷史記錄錯誤的風險。

## What Changes
- 調整提醒讀取與操作流程，全面改用提醒 UUID 定位試算表資料列
- 更新刪除、暫停/啟用、排程檢查等 GAS 邏輯，透過 UUID 查找列並執行軟刪除
- 移除前後端對 rowIndex 的依賴與欄位傳遞，確保 API 與 UI 皆只暴露 UUID
- 擴充提醒管理規格，明確要求提醒資料須具備穩定的 UUID 識別並以此進行後續操作

## Impact
- Affected specs: reminder-management
- Affected code: 程式碼.js, index.html
