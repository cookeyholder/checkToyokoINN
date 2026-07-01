## 1. 批次寫入改造

- [x] 1.1 修改 `submitReminder()`：在迴圈中為每個房型產生 UUID（`Utilities.getUuid()`），收集所有資料列到二維陣列
- [x] 1.2 每列的 `branchCode` 必須比照 `addReminder` 加上 `"'"` 前綴以強制試算表文字格式，避免 `00066` 變成 `66`
- [x] 1.3 在迴圈結束後以 `sheet.getRange().setValues()` 一次寫入
- [x] 1.4 保留單筆房型情況的相容性

## 2. 移除死碼

- [x] 2.1 整段移除 `addBatchReminders()` 函式定義
- [x] 2.2 確認無其他程式碼參照 `addBatchReminders`

## 3. 驗證

- [x] 3.1 測試單個房型建立提醒
- [x] 3.2 測試多個房型批次建立提醒
- [x] 3.3 確認無新錯誤（Deploy & 觸發測試）
