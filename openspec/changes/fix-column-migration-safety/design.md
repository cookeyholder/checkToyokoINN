## Context

目前 `COLUMN_INDICES` 定義了 17 個欄位（索引 0-16）。若程式碼新增欄位而使用者未手動插入試算表欄位，所有在插入點之後的資料都會被錯誤解讀。例如在 `userEmail`（L 欄）與 `createdAt`（N 欄）之間插入 `notificationEmail`（M 欄）後，舊資料的 N 欄會被讀成 M 欄，O 讀成 N，以此類推。

## Goals / Non-Goals

**Goals:**
- `initializeRemindersSheet()` 自動偵測目前欄位數，若小於預期則插入缺失欄位
- 新功能也應自動補欄，不依賴使用者手動操作

**Non-Goals:**
- 不處理已錯位的舊資料修復（系統只在初始化或執行時補欄，不回頭 rewrite）
- 不套用到其他工作表（檢查歷史、帳號管理等）

## Decisions

**決策一：在 `initializeRemindersSheet()` 內處理 migration**
- **理由**：該函式已在每次部署或手動初始化時執行，是自然擴充點。
- 先用 `sheet.getLastColumn()` 取得目前欄位數，若小於 `Object.keys(COLUMN_INDICES.reminders).length` 則在對應位置插入 `sheet.insertColumnAfter(lastColumnIndex)` 並設定表頭。

**決策二：僅補欄不搬移舊資料**
- **理由**：`insertColumnAfter` 會自動將後續欄位移動，插入的新欄位的儲存格會是空的。這比試圖逐列搬移資料更安全且不會觸發試算表 quota 限制。

## Risks / Trade-offs

- **使用者已手動補欄** → `getLastColumn()` 會回傳正確值，補欄邏輯直接跳過，無副作用。
- **插入位置錯誤** → 若使用者已在欄位中間自行插入其他無關欄位，自動補欄可能在錯誤位置插入。需確保只在「現有欄位數小於預期」時才在末端補欄。
