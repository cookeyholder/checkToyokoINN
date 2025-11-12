# Room Checker 規格

## ADDED Requirements

### Requirement: 房間可用性查詢
系統 SHALL 能夠向東橫 INN API 發送查詢請求,檢查指定條件下是否有禁菸空房。

#### Scenario: 成功查詢空房狀態
- **WHEN** 提供有效的查詢參數 (分店編號、房型、人數、日期等)
- **THEN** 系統應構建正確的查詢 URL
- **AND** 發送 HTTP GET 請求到 `https://www.toyoko-inn.com/china/search/result/room_plan/`
- **AND** 解析回應內容判斷是否有空房

#### Scenario: 有空房可訂
- **WHEN** API 回應包含可預訂的房間資料
- **THEN** 函式應回傳 `true` 或包含房間詳情的物件

#### Scenario: 無空房
- **WHEN** API 回應頁面沒有房間資料
- **THEN** 函式應回傳 `false` 或空陣列

### Requirement: 查詢參數建構
系統 SHALL 正確建構包含所有必要參數的查詢 URL。

#### Scenario: 建構完整查詢 URL
- **WHEN** 提供查詢參數物件
- **THEN** 應產生格式如下的 URL:
  ```
  https://www.toyoko-inn.com/china/search/result/room_plan/?hotel=00066&roomType=30&people=2&room=1&smoking=noSmoking&r_avail_only=true&start=2026-02-04&end=2026-02-05
  ```
- **AND** 所有參數值應正確編碼

#### Scenario: 參數映射
- **WHEN** 建構查詢 URL
- **THEN** 應正確映射參數:
  - `hotel`: 分店編號 (例如: "00066")
  - `roomType`: 單一房型代號 (10=單人房, 20=雙人房, 30=雙床房, 40=三人以上房)
  - `people`: 成人人數
  - `room`: 房間數
  - `smoking`: 固定為 "noSmoking" (禁菸房)
  - `r_avail_only`: 固定為 "true" (僅顯示空房)
  - `start`: 入住日期 (YYYY-MM-DD)
  - `end`: 退房日期 (YYYY-MM-DD)
- **AND** 每次查詢僅包含一個 roomType 值

### Requirement: 網頁內容解析
系統 SHALL 解析東橫 INN API 回應的 HTML 內容,判斷是否有可用房間。

#### Scenario: 解析有空房的回應
- **WHEN** 回應 HTML 包含房間資訊區塊
- **THEN** 系統應識別為有空房

#### Scenario: 解析無空房的回應
- **WHEN** 回應 HTML 不包含房間資訊或顯示無空房訊息
- **THEN** 系統應識別為無空房

### Requirement: 錯誤處理
系統 SHALL 妥善處理查詢過程中的各種錯誤情況。

#### Scenario: 網路請求失敗
- **WHEN** HTTP 請求失敗 (逾時、網路錯誤等)
- **THEN** 系統應記錄錯誤到 Logger
- **AND** 回傳預設值 (無空房)
- **AND** 不應中斷整個檢查流程

#### Scenario: 無效的回應格式
- **WHEN** 收到非預期格式的 HTML 回應
- **THEN** 系統應記錄警告
- **AND** 安全地回傳無空房狀態

#### Scenario: HTTP 錯誤狀態碼
- **WHEN** 收到 4xx 或 5xx 錯誤回應
- **THEN** 系統應記錄具體的錯誤碼和訊息
- **AND** 適當處理 (例如: 429 Too Many Requests 應延遲重試)
