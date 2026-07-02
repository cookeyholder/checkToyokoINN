## Context

`COLUMN_INDICES` 是硬編碼的數字常數：

```javascript
var COLUMN_INDICES = {
    reminders: {
        uuid: 0,
        branchCode: 1,
        // ... 每個欄位對應一個數字
        notificationEmail: 12,
        createdAt: 13,
        // ...
    }
};
```

每次試算表欄位變動就需手動更新此物件及所有相依函式，這種集中式常數已成為維護痛點。

## Goals / Non-Goals

**Goals:**
- 新增 `getColumnIndices(sheet)` 根據標題列名稱動態解析索引
- 所有讀取提醒資料的位置改為使用動態解析（`getReminders`、`getReminderList`、`sendNotification`）
- 保留 `COLUMN_INDICES` 作為初始化試算表表頭時的參照，不全部移除

**Non-Goals:**
- 不改寫檢查歷史工作表的欄位解析（非本次範圍）
- 不改變任何現有行為或資料格式
- 不將寫入路徑改為動態對應：寫入時仍使用 `COLUMN_INDICES` 固定順序，因為標題列永遠由 `initializeRemindersSheet()` 以 canonical 順序建立，使用者幾乎不會手動調整試算表欄位，不產生讀寫不一致的問題

## Decisions

**決策一：每次讀取時重新解析**
- **理由**：GAS 專案的記憶體生命週期短，每次執行都是獨立的。每次讀取時解析一次即可，無需快取。
- 替代方案「全域快取」：在 `getColumnIndices` 內用靜態變數快取結果，但 GAS 執行緒模型不保證跨呼叫共享狀態。

**決策二：不修改 COLUMN_INDICES 的使用方式**
- **理由**：保留常數作為權威定義，僅在「讀取已存在的資料列」時使用動態解析。初始化寫入時仍用 `COLUMN_INDICES` 確保順序正確。

**決策三：寫入路徑維持固定順序，不改成動態對應**
- **理由**：`initializeRemindersSheet()` 永遠以 canonical 順序建立標題列，且使用者幾乎不會手動調整試算表欄位，因此寫入時的欄位順序永遠與表頭一致。若改成動態對應，`addReminder()` 的陣列建構會從 19 行的簡潔陣列字面值膨脹為 ~30 行的 position-aware 賦值，降低可讀性卻沒有實際效益。
- 替代方案「寫入全動態」：`new Array(cols).fill("")` 再逐欄 `newRow[indices.xxx] = value`。優點是欄位順序完全獨立，缺點是程式碼更冗長，且在使用者不手動調欄的前提下無實際場景受惠。

## Risks / Trade-offs

- **標題列不一致** → 若試算表標題列被修改或刪除，`indexOf` 會傳回 -1。由於使用者不會手動修改標題，不一致只能來自程式碼錯誤或意外破壞。`getColumnIndices` 在找不到時直接拋出明確錯誤，讓問題在開發階段就被發現，而非無聲 fallback 到錯誤資料。
- **寫入與讀取校準** → 寫入維持 `COLUMN_INDICES` 固定順序，當且僅當使用者手動調換試算表欄位時會產生不一致。但此情境已被使用者確認為幾乎不會發生，因此不投入額外複雜度處理。
- **效能** → `getRange(1, 1, 1, lastColumn).getValues()[0]` + `indexOf` 對每批讀取（約數十列）影響可忽略。
