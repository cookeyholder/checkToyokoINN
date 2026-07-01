## 1. 後端：COLUMN_INDICES 與試算表表頭

- [x] 1.1 在 `COLUMN_INDICES.reminders` 新增 `notificationEmail: 12`，並將 `createdAt` 以後的 index 全部 +1
- [x] 1.2 在 `initializeRemindersSheet()` 的表頭陣列中，於「使用者 Email」之後插入「提醒收件 Email」
- [x] 1.3 調整 `columnWidths` 陣列與欄寬設定（新增一組提醒收件 Email 欄寬 200，其餘對應位移）

## 2. 後端：讀寫邏輯

- [x] 2.1 修改 `addReminder()`：在 `newRow` 陣列中新增 `reminderData.notificationEmail` 欄位
- [x] 2.2 修改 `getReminders()`：讀取試算表時新增 `notificationEmail` 欄位（`rowData[COLUMN_INDICES.reminders.notificationEmail]`）
- [x] 2.3 修改 `submitReminder()`：改用 `formData.notificationEmail`（從表單取得）取代 Session Email
- [x] 2.4 修改 `getReminderList()`：在 return 的提醒物件中新增 `notificationEmail` 欄位

## 3. 後端：通知發送

- [x] 3.1 修改 `sendNotification()`：收件人改為 `reminder.notificationEmail`，若為空則 fallback 到 `reminder.userEmail`

## 4. 前端：表單新增 Email 欄位

- [x] 4.1 在 `index.html` 的表單中，於「提醒結束時間」之後新增「提醒收件 Email」文字輸入欄位（type="email"、必填）
- [x] 4.2 修改 `handleFormSubmit()`：將 `notificationEmail` 加入 `formData` 物件
- [x] 4.3 在表單驗證邏輯中確保 `notificationEmail` 不為空且格式正確

## 5. 規格文件

- [x] 5.1 更新 `openspec/specs/spreadsheet-data/spec.md`：提醒清單欄位格式加入「提醒收件 Email」
- [x] 5.2 更新 `openspec/specs/web-ui/spec.md`：表單欄位與提交行為加入「提醒收件 Email」
- [x] 5.3 更新 `openspec/specs/email-notification/spec.md`：通知發送行為改為優先讀取「提醒收件 Email」
- [x] 5.4 更新 `openspec/specs/reminder-management/spec.md`：提醒資料欄位定義加入「使用者 Email」與「提醒收件 Email」

## 6. 專案文件

- [x] 6.1 更新 `README.md`：提醒清單欄位表加入「提醒收件 Email」（位於使用者 Email 之後）
- [x] 6.2 更新 `ARCHITECTURE.md`：提醒工作表欄位列表加入「提醒收件 Email」
- [x] 6.3 更新 `USER_GUIDE.md`：建立提醒步驟加入填寫提醒收件 Email 的說明
