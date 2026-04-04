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

## 語言設定

與使用者溝通時，請一律使用**繁體中文**回應。程式碼、指令、變數名稱等技術內容維持英文，但說明文字請使用繁體中文。
