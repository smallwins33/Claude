# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 永久授權限制 ⚠️

**工作目錄只限 `/Users/mi/Developer`。任何情況下不得讀寫此目錄以外的路徑。**

- 不可執行 `pip` / `pip3 install`（會寫入 `/Users/mi/Library`）
- 不可讀寫 `/Users/mi/Library` 或任何非 `/Users/mi/Developer` 的目錄
- 套件問題請告知使用者自行處理，不代為執行
- 此限制適用於所有對話，無例外

## 執行原則 ⚠️

使用者交代的事一次做到位，不讓使用者 double check 或追問才補完。

- 永久規則 → 記憶檔＋CLAUDE.md 兩個都寫，不分開做
- 執行前先確認「完整做完的定義是什麼」，再動手
- 遇到技術問題先自己查清楚再行動，不要讓使用者幫忙發現錯誤

## Systeme MCP 注意事項 ⚠️

`get_contacts` API 的 `registeredAfter`、`tags` filter 無效，`order:asc` 分頁永遠 loop。
唯一可靠方式：`order:desc + 手動日期截止 + seen_ids`。
- Layer 1（本期）：`--mode new --since PERIOD_START`
- Layer 2（歷史，從 2026-03-01 起）：`--mode new --since 2026-03-01 --out /tmp/systeme_leads.json`

## Notion 諮詢資料期間篩選 ⚠️

抓取 Notion 諮詢記錄時，**不抓全量**，必須自動依廣告期間篩選：
- 只取 `諮詢時間 >= since`（報告期起始日）的記錄
- 本期廣告的名單，諮詢時間不可能是上上週以前，不需使用者再說區間
- **永遠不呼叫 notion-fetch**，用 notion-query-data-sources 拿 email、狀態、諮詢時間三欄即可
- 比對進諮詢：email 對上即可；比對成交：才看 狀態 欄位

## 語言設定

與使用者溝通時，請一律使用**繁體中文**回應。程式碼、指令、變數名稱等技術內容維持英文，但說明文字請使用繁體中文。
