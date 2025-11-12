# 技術設計文件

## Context

這是一個從零開始的 Google Apps Script 專案,需要建立完整的房間可用性監控系統。系統涉及多個關注點:網頁介面、資料儲存、定期執行、外部 API 呼叫和郵件通知。

**主要限制:**
- GAS 執行時間限制 (最多 6 分鐘)
- UrlFetch 和 Email 每日配額
- 需要簡單易維護的架構
- 無法使用外部函式庫或 npm 套件

**利益相關者:**
- 使用者: 需要即時得知東橫 INN 空房資訊
- 系統: 需要穩定執行且不超過配額限制

## Goals / Non-Goals

### Goals
- 建立可靠的自動化房間檢查系統
- 提供直覺易用的網頁介面
- 確保在時間範圍內正確發送通知
- 避免重複通知造成困擾
- 妥善處理錯誤,不影響其他提醒的執行

### Non-Goals
- 不支援即時推送通知 (僅 Email)
- 不提供歷史記錄或統計分析
- 不整合其他訂房平台
- 不支援多使用者共用 (每個使用者需部署自己的副本)

## Decisions

### 決策 1: 單一檔案架構
**決定**: 將所有後端邏輯放在 `程式碼.js` 單一檔案中,使用函式分組組織。

**理由**:
- GAS 專案檔案數量較少時更易維護
- 不需要複雜的模組系統
- 所有函式可直接互相呼叫
- 符合專案慣例的簡單優先原則

**替代方案**:
- 多檔案分離: 會增加複雜度,對於此規模的專案不必要

### 決策 2: 試算表作為資料庫
**決定**: 使用 Google Sheets 作為唯一的資料儲存層。

**理由**:
- 使用者可直接檢視和編輯資料
- 不需額外的資料庫設定
- 符合 GAS 生態系統的最佳實踐
- 資料量小,不會有效能問題

**資料模型**:
```
東橫INN分店工作表:
- 標題列: 分店名稱 | 分店編號
- 欄 A: 分店名稱 (字串,例如: "東京新宿店")
- 欄 B: 分店編號 (字串,例如: "00066")

房型代號工作表:
- 欄 A: 房型代號 (數字: 10/20/30/40)
- 欄 B: 房型名稱 (字串)

提醒清單工作表:
- 欄 A-J: 提醒資料 (分店、房型、日期、時間等)
- 欄 K: 使用者 Email (建立提醒時記錄)
- 欄 L: 建立時間
- 欄 M: 最後通知時間 (ISO 8601)
- 欄 N: 通知狀態 (字串: "未通知"/"已通知")
- 欄 O: 提醒狀態 (字串: "啟用"/"暫停")

檢查歷史工作表 (新增):
- 欄 A: 提醒 ID (關聯到提醒清單)
- 欄 B: 檢查時間 (時間戳記)
- 欄 C: 檢查結果 ("有空房"/"無空房"/"錯誤")
- 欄 D: 錯誤訊息 (如果有)
- 欄 E: 是否發送通知 (布林值)

網站參數設定工作表 (新增):
- 欄 A: 參數項目
- 欄 B: 參數內容
- 欄 C: 參數說明

帳號管理工作表 (新增):
- 欄 A: Email
- 欄 B: 姓名
- 欄 C: 人員編號
- 欄 D: 部門單位
- 欄 E: 群組
- 欄 F: 狀態 ("啟用"/"停用")
- 欄 G: 備註
```

### 決策 3: HTML 解析策略
**決定**: 使用簡單的字串搜尋判斷是否有空房,而非完整的 DOM 解析。

**理由**:
- GAS 沒有內建的 DOM 解析器
- 東橫 INN 網站結構相對穩定
- 只需判斷有無空房,不需提取詳細資料
- 效能更好,程式碼更簡單

**實作方式**:
```javascript
function hasAvailableRooms(htmlContent) {
  // 尋找特定的標記或文字模式
  // 例如: 檢查是否包含 "空室" 或特定的 CSS class
  return htmlContent.indexOf('空室') !== -1 || 
         htmlContent.indexOf('available') !== -1;
}
```

**風險**: 如果網站結構改變,需要更新解析邏輯。
**緩解**: 加入詳細的 Logger 輸出,當解析失敗時容易追蹤。

### 決策 4: 通知去重機制
**決定**: 使用「最後通知時間」欄位實作冷卻期 (1 小時)。

**理由**:
- 避免使用者在短時間內收到大量重複郵件
- 實作簡單,只需比較時間戳記
- 1 小時的冷卻期提供更即時的通知

**邏輯**:
```javascript
function shouldSendNotification(lastNotificationTime) {
  if (!lastNotificationTime) return true;
  
  const now = new Date();
  const lastNotified = new Date(lastNotificationTime);
  const hoursDiff = (now - lastNotified) / (1000 * 60 * 60);
  
  return hoursDiff >= 1; // 1 小時冷卻期
}
```

### 決策 5: 時間觸發器策略
**決定**: 使用 ScriptApp.newTrigger() 建立每 5 分鐘執行的觸發器。

**理由**:
- 5 分鐘是反應速度與配額消耗的平衡點
- 每天最多 288 次檢查,遠低於 UrlFetch 配額
- 在函式中過濾提醒時間範圍,不在觸發器層級處理

**設定函式**:
```javascript
function setupTimeTrigger() {
  // 刪除現有觸發器
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // 建立新觸發器
  ScriptApp.newTrigger('checkAllReminders')
    .timeBased()
    .everyMinutes(5)
    .create();
}
```

### 決策 6: 網頁應用程式架構
**決定**: 使用簡單的 Server-Side Rendering,在後端生成完整的 HTML。

**理由**:
- 符合 GAS HtmlService 的設計模式
- 不需要複雜的前端框架
- 資料量小,不需要客戶端狀態管理
- 使用 google.script.run 與後端通訊

**架構**:
```
index.html (前端)
├── 靜態 HTML 結構
├── CSS 樣式 (內嵌)
└── JavaScript
    ├── 表單處理
    ├── google.script.run 呼叫
    └── UI 更新

程式碼.js (後端)
├── doGet() - 回傳 HTML
├── API 函式 (由前端呼叫)
│   ├── getBranchOptions()
│   ├── submitReminder(data)
│   ├── getReminderList()
│   └── deleteReminder(id)
└── 核心邏輯函式
```

### 決策 7: 錯誤處理策略
**決定**: 使用 try-catch 包裹關鍵操作,記錄錯誤但不中斷流程。

**理由**:
- 單一提醒失敗不應影響其他提醒
- 網路錯誤是暫時性的,下次執行可能成功
- Logger 提供足夠的除錯資訊

**模式**:
```javascript
function checkAllReminders() {
  const reminders = getReminders();
  
  reminders.forEach(reminder => {
    try {
      // 檢查邏輯
      if (checkRoomAvailability(reminder)) {
        sendNotification(reminder);
      }
    } catch (error) {
      Logger.log(`提醒 ${reminder.id} 檢查失敗: ${error.message}`);
      // 繼續處理下一個提醒
    }
  });
}
```

## Risks / Trade-offs

### 風險 1: 東橫 INN 網站結構變更
**影響**: HTML 解析邏輯可能失效,無法正確判斷空房狀態。

**緩解**:
- 在解析函式中加入多種識別模式作為備案
- 定期測試並更新解析邏輯
- Logger 記錄解析失敗的情況,便於快速發現問題

### 風險 2: GAS 配額限制
**影響**: 大量提醒可能超過 UrlFetch 或 Email 配額。

**緩解**:
- 每 5 分鐘執行一次,每天最多 288 次,遠低於限制
- 通知去重機制減少郵件數量
- 可在程式碼中加入配額監控邏輯

### 風險 3: 時區處理
**影響**: 提醒時間範圍可能因時區設定而錯誤。

**緩解**:
- 使用 appsscript.json 中設定的 Asia/Taipei 時區
- 在時間比較邏輯中明確使用相同時區
- 在 UI 上清楚標示使用的時區

### Trade-off: 即時性 vs 配額
**選擇**: 5 分鐘檢查頻率,而非更頻繁的檢查。

**好處**: 節省配額,降低對目標網站的負擔。

**壞處**: 可能錯過非常短時間內出現又消失的空房。

**判斷**: 對於飯店訂房場景,5 分鐘的延遲是可接受的。

## Migration Plan

此為全新專案,無需遷移。

**初次部署步驟**:
1. 建立 Google Sheets 試算表,設定三個工作表
2. 在試算表中填入分店資料和房型代號
3. 複製試算表 ID 到程式碼設定中
4. 部署為網頁應用程式
5. 執行 `setupTimeTrigger()` 建立時間觸發器
6. 授權必要的權限
7. 測試建立提醒和接收通知

**Rollback**: 刪除觸發器和網頁應用程式部署即可。

## Open Questions

目前沒有未決問題,所有主要設計決策已確定:
- ✅ 過期提醒: 保留但不檢查
- ✅ 郵件通知冷卻時間: 固定為 1 小時,不開放使用者自訂
- ✅ 檢查歷史: 永久保存,不自動清理舊記錄
