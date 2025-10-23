# Repository Guidelines

## 專案結構與模組安排
本專案鎖定 Google Apps Script，所有執行程式碼置於存放庫根目錄。`appsscript.json` 定義部署設定，功能模組依職責拆分為多個檔案：`menu.js` 建立工作列與流程指令（如 `runFullSchedulingPipeline`、`resumePipelineAfterManualAdjustments`），`dataPreparation.js` 處理資料匯入與篩選，`scheduling.js` 集中節次與試場編排邏輯，`reportGeneration.js` 產生公告與列印資料，`printing.js` 封裝 PDF 合併輸出，`helpers.js` 提供共用工具。新增功能時請優先尋找對應模組，必要時再建立新檔案並於文件補充說明。

## 建置、測試與開發指令
開發流程採用 Apps Script CLI：先執行 `npm install -g @google/clasp` 完成安裝，之後以 `clasp login` 認證帳號；`clasp push` 上傳目前檔案至連結的腳本專案，`clasp pull` 則同步遠端變更。臨時驗證可在試算表中開啟 Apps Script 編輯器，但請確保最終變更回存到版本庫並透過 `clasp push` 部署。

## 程式風格與命名慣例
沿用 Apps Script 預設：兩空格縮排，分號使用需在同一檔案內保持一致，變數宣告以 `const`、`let` 為主。函式與變數採駝峰式命名（如 `assignExamRooms`, `populateSessionTimes`）；讀取參數表的常數則使用大寫蛇形。共用邏輯宜集中在 `helpers.js`，資料或排程專用的程式碼請放回對應模組以維持職責單一。

## 測試指引
目前尚未建立自動化測試，請於試算表的測試副本進行手動驗證。透過 Apps Script 編輯器的 Run 功能執行新進入點，確認產出的資料範圍（例如「註冊組補考名單」、「參數區」）符合預期。請在 PR 敘述中列出手動測試情境，方便後續重現。

## 提交與 Pull Request 守則
用繁體中文撰寫符合 conventional commit 規範的提交訊息，必要時補上英文說明。內文可連結相關議題或試算表，並註明任何工作表結構調整。Pull Request 需包含變更概述、測試紀錄，以及影響輸出成果時的截圖。

## 試算表設定提示
執行腳本前請確認「參數區」中的 B2、B5-B8 等參數範圍及各模組引用的欄位標題未被修改。若學年度或工作表名稱異動，請同步更新程式中的對應常數並在 PR 說明，以維持部署一致性。
