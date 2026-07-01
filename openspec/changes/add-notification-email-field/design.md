## Context

「提醒清單」工作表目前僅有「使用者 Email」欄位，由後端自動填入 GAS 執行者 Email。通知信固定寄到此欄位。使用者無法在網頁上指定其他收件地址。

## Goals / Non-Goals

**Goals:**
- 在「提醒清單」工作表新增「提醒收件 Email」欄位（位於「使用者 Email」與「建立時間」之間）
- 在網頁表單加入「提醒收件 Email」輸入欄位
- 後端 `sendNotification()` 改讀新欄位決定寄送對象
- 保留既有「使用者 Email」邏輯不變

**Non-Goals:**
- 不更動現有 Session Email 擁有權驗證行為
- 不更動通知冷卻機制與寄送時機
- 不處理 migration（使用者自行調整試算表）

## Decisions

| 決策 | 選擇 | 理由 |
|------|------|------|
| 欄位位置 | 「使用者 Email」與「建立時間」之間 | 語意上通知收件與使用者 Email 相關，且不影響後續既有欄位索引 |
| 新欄位為空時的行為 | fallback 到「使用者 Email」 | 相容既有資料，不讓舊提醒因無新欄位值而中斷通知 |
| 內部變數命名 | `notificationEmail` | 與「使用者 Email」(`userEmail`) 明確區分職責 |
| 是否需要 migration script | 否 | 使用者自行在試算表插入欄位並補填舊資料 |

## Risks / Trade-offs

- 使用者若填寫錯誤 Email 將收不到通知 → 無 mitigation，與一般 Email 輸入行為相同
- 新增欄位後 COLUMN_INDICES 中「建立時間」以後的 index 需全部 +1 → 影響範圍集中需仔細核對
