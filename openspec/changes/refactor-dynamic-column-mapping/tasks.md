## 1. 實作動態索引解析

- [ ] 1.1 新增 `getColumnIndices(sheet)` 工具函式
- [ ] 1.2 在函式內檢查每個必要欄位名稱是否存在，遺漏時拋出明確錯誤

## 2. 替換所有讀取位置

- [ ] 2.1 `getReminders()`：改用 `getColumnIndices` 解析各欄位
- [ ] 2.2 `getReminderList()`：改用 `getColumnIndices`
- [ ] 2.3 `sendNotification()` 內讀取使用者 Email 欄位改用動態解析
- [ ] 2.4 確認無其他函式直接以數字索引讀取 reminders 工作表

## 3. 驗證

- [ ] 3.1 測試讀取現有提醒資料無錯位
- [ ] 3.2 測試新增提醒後可正確讀回
- [ ] 3.3 確認初始化試算表功能不受影響
