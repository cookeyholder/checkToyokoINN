# Email Notification 規格

## ADDED Requirements

### Requirement: 空房通知郵件發送
系統 SHALL 在發現空房時自動發送 Email 通知給使用者。

#### Scenario: 發送空房通知
- **WHEN** 系統檢查到某提醒有可用的禁菸空房
- **AND** 當前時間在提醒的時間範圍內
- **THEN** 系統應發送 Email 通知
- **AND** 收件人為該提醒記錄中的使用者 Email

#### Scenario: 不在提醒時間範圍內
- **WHEN** 發現空房但當前時間不在提醒範圍內
- **THEN** 系統不應發送通知
- **AND** 應記錄到 Logger 以供追蹤

### Requirement: 郵件內容格式
系統 SHALL 發送包含完整訂房資訊的格式化郵件。

#### Scenario: 郵件主旨
- **WHEN** 發送空房通知
- **THEN** 郵件主旨應為:
  ```
  【東橫 INN 空房通知】{分店名稱} - {入住日期}
  ```
  例如: 【東橫 INN 空房通知】東京新宿店 - 2026-02-04

#### Scenario: 郵件內容
- **WHEN** 發送空房通知
- **THEN** 郵件內文應包含:
  - 分店名稱 (不包含分店編號)
  - 房型
  - 入住日期和退房日期
  - 成人人數和房間數
  - 直接訂房連結 (包含所有查詢參數的 URL)
  - 提醒設定時間
  - 檢查時間 (發現空房的時間)

#### Scenario: 郵件格式
- **WHEN** 撰寫郵件內容
- **THEN** 應使用 HTML 格式提高可讀性
- **AND** 包含清楚的標題、段落和連結
- **AND** 使用合適的字體和顏色

### Requirement: 郵件發送錯誤處理
系統 SHALL 妥善處理郵件發送過程中的錯誤。

#### Scenario: 郵件發送成功
- **WHEN** GmailApp.sendEmail() 成功執行
- **THEN** 系統應記錄發送成功到 Logger
- **AND** 更新提醒的最後通知時間

#### Scenario: 郵件發送失敗
- **WHEN** 郵件發送時發生錯誤 (例如: 配額用盡、權限不足)
- **THEN** 系統應記錄錯誤詳情到 Logger
- **AND** 不應中斷其他提醒的檢查
- **AND** 繼續處理下一個提醒

#### Scenario: Email 地址無效
- **WHEN** 提醒記錄中的使用者 Email 為空值或無效
- **THEN** 系統應記錄警告到 Logger
- **AND** 跳過該次通知發送
- **AND** 在檢查歷史中記錄為錯誤

### Requirement: 通知去重機制
系統 SHALL 避免在短時間內重複發送相同的空房通知。

#### Scenario: 首次發現空房
- **WHEN** 首次檢查到某提醒有空房
- **THEN** 應立即發送通知
- **AND** 記錄最後通知時間

#### Scenario: 持續有空房
- **WHEN** 上次通知後 X 分鐘內再次檢查到相同空房
- **THEN** 系統不應重複發送通知
- **AND** 應在 Logger 記錄「空房持續可用」

#### Scenario: 通知冷卻時間
- **WHEN** 設計去重機制
- **THEN** 應定義冷卻時間為 1 小時 (固定值,不開放使用者設定)
- **AND** 超過冷卻時間後可再次通知

### Requirement: 訂房連結生成
系統 SHALL 在通知郵件中包含可直接開啟訂房頁面的 URL。

#### Scenario: 生成完整訂房 URL
- **WHEN** 建立通知郵件
- **THEN** 應生成包含所有查詢參數的 URL
- **AND** 格式與房間檢查時使用的相同
- **AND** 使用者點擊連結後應直接看到可預訂的房間

#### Scenario: URL 參數編碼
- **WHEN** 生成訂房 URL
- **THEN** 所有參數值應正確 URL 編碼
- **AND** 特殊字元應正確轉換

### Requirement: 郵件配額管理
系統 SHALL 監控並管理 Google Apps Script 的郵件發送配額。

#### Scenario: 接近配額限制
- **WHEN** 當日已發送接近配額上限的郵件數量
- **THEN** 系統應記錄警告到 Logger
- **AND** 可選擇性降低檢查頻率或暫停通知

#### Scenario: 配額用盡
- **WHEN** 嘗試發送郵件但已達配額上限
- **THEN** 系統應捕捉錯誤
- **AND** 記錄清楚的錯誤訊息
- **AND** 不中斷其他功能的執行
