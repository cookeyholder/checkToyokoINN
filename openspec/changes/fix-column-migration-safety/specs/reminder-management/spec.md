## ADDED Requirements

### Requirement: 自動欄位 migration
系統 SHALL 在初始化提醒清單工作表時自動偵測並補足缺失的欄位。

#### Scenario: 工作表欄位數不足
- **WHEN** 提醒清單工作表的欄位數小於程式碼預期的欄位數
- **THEN** 系統應自動在試算表末端插入缺失的欄位
- **AND** 設定新欄位的表頭名稱
- **AND** 不破壞既有資料

#### Scenario: 工作表欄位數正確
- **WHEN** 提醒清單工作表的欄位數等於或大於預期
- **THEN** 系統應跳過補欄
- **AND** 不進行任何修改
