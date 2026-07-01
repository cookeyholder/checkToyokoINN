## MODIFIED Requirements

### Requirement: 空房通知郵件發送
系統 SHALL 在發現空房時自動發送 Email 通知給使用者。

#### Scenario: 發送空房通知
- **WHEN** 系統檢查到某提醒有可用的禁菸空房
- **AND** 當前時間在提醒的時間範圍內
- **THEN** 系統應發送 Email 通知
- **AND** 收件人優先為該提醒記錄中的「提醒收件 Email」
- **AND** 若「提醒收件 Email」為空則 fallback 到「使用者 Email」

### Requirement: 郵件發送錯誤處理
系統 SHALL 妥善處理郵件發送過程中的錯誤。

#### Scenario: Email 地址無效
- **WHEN** 提醒記錄中的「提醒收件 Email」與「使用者 Email」皆為空值或無效
- **THEN** 系統應記錄警告到 Logger
- **AND** 跳過該次通知發送
- **AND** 在檢查歷史中記錄為錯誤
