## Why
目前系統在爬取東橫 INN 分店資訊時，會將「日本以外」地區的分店過濾掉，導致使用者無法設定或接收日本以外地區（例如：韓國、蒙古等）的分店提醒。

## What Changes
- 修改分店更新邏輯，取消排除「日本以外」分店的限制。
- 保留完整的地區與都道府縣資訊，讓所有分店皆能正確寫入「東橫INN分店」工作表。

## Impact
- 影響檔案：[scrapers.gs](file:///home/cookeyholder/Projects/checkToyokoINN/scrapers.gs) 裡的 [updateBranchesSheet](file:///home/cookeyholder/Projects/checkToyokoINN/scrapers.gs#L548) 函數。
- 影響規格：新增對「東橫INN分店」工作表包含日本以外分店的規格描述。
