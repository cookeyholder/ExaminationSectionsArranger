# Chrome DevTools MCP 整合指南

本指南說明如何在 VS Code 中使用 Chrome DevTools MCP 伺服器來測試和除錯 Google Apps Script Web App。

## 什麼是 Chrome DevTools MCP？

Chrome DevTools MCP 是一個 Model Context Protocol 伺服器，允許 AI 助手（如 GitHub Copilot）透過程式化方式控制 Chrome 瀏覽器的 DevTools。

## 功能特色

- 🌐 **頁面控制**：導航、截圖、腳本執行
- 🔍 **元素互動**：點擊、填寫表單、滾動
- 📊 **效能分析**：追蹤載入時間、Core Web Vitals
- 🌐 **網路監控**：列出請求、檢查回應
- 🎨 **快照擷取**：文字快照、視覺截圖

## 前置需求

1. Chrome 瀏覽器已安裝
2. VS Code 已安裝 MCP 擴充功能
3. Web App 已部署

## 基本使用流程

### 1. 部署 Web App

在 Apps Script 編輯器中：

```
部署 → 新增部署作業 → 網頁應用程式
執行身分：我
存取權：所有人
→ 部署
```

取得 Web App URL（例如）：
```
https://script.google.com/macros/s/AKfycby.../exec
```

### 2. 在 Chrome 開啟 Web App

```bash
# 手動在 Chrome 開啟 URL
# 或使用指令
open -a "Google Chrome" "YOUR_WEB_APP_URL"
```

### 3. 透過 MCP 與頁面互動

在 VS Code 中，您可以透過 AI 助手執行以下操作：

#### 列出開啟的頁面

```
AI: 列出所有開啟的瀏覽器頁面
```

#### 選擇頁面

```
AI: 選擇第 1 個頁面
```

#### 擷取頁面快照

```
AI: 擷取當前頁面的文字快照
```

#### 執行 JavaScript

```
AI: 在頁面中執行 JavaScript: loadData()
```

#### 擷取截圖

```
AI: 擷取當前頁面的截圖
```

## 實際測試範例

### 範例 1：驗證資料載入

```javascript
// 在 VS Code Chat 中
「請在 Web App 頁面執行以下測試：
1. 擷取頁面快照確認介面已載入
2. 執行 loadData() 載入資料
3. 檢查 currentData 是否包含 exam.population
4. 截圖儲存結果」
```

MCP 會自動：
1. 擷取頁面內容
2. 執行 JavaScript
3. 驗證回傳值
4. 產生截圖

### 範例 2：效能測試

```javascript
「請測試 Web App 的載入效能：
1. 重新載入頁面
2. 啟動效能追蹤
3. 執行 loadData()
4. 停止追蹤並分析結果」
```

MCP 會：
1. 導航到頁面
2. 開始效能記錄
3. 觸發操作
4. 產生效能報告（LCP, FID, CLS 等）

### 範例 3：網路請求監控

```javascript
「請監控 API 請求：
1. 列出所有網路請求
2. 找到 mode=api 的請求
3. 檢查回應內容
4. 驗證 JSON 結構」
```

MCP 會：
1. 列出網路請求
2. 篩選特定請求
3. 取得回應內容
4. 解析並驗證

## 測試腳本範例

### 自動化測試流程

透過 AI 助手執行完整測試：

```markdown
請執行以下 GAS Web App 自動化測試：

**前置作業**
1. 開啟頁面：[YOUR_WEB_APP_URL]
2. 等待頁面載入完成

**測試 1：介面驗證**
- 擷取頁面快照
- 確認包含「GAS 開發測試工具」標題
- 確認包含統計卡片區域

**測試 2：資料載入**
- 點擊「重新載入資料」按鈕
- 等待 2 秒
- 執行 JS: currentData
- 驗證 currentData.exam.population > 0

**測試 3：JSON 下載**
- 執行 JS: downloadJSON()
- 確認瀏覽器觸發下載

**測試 4：效能測試**
- 重新載入頁面
- 測量 loadData() 執行時間
- 確認 < 2000ms

**產出**
- 每個測試的截圖
- 效能數據摘要
- 通過/失敗報告
```

## 進階用法

### 1. 模擬慢速網路

```javascript
「請在慢速 3G 網路下測試：
1. 啟用 Slow 3G 節流
2. 重新載入頁面
3. 測量載入時間
4. 停用節流」
```

### 2. 模擬低效能裝置

```javascript
「請模擬低階裝置：
1. 啟用 4x CPU 節流
2. 執行 loadData()
3. 測量執行時間
4. 停用節流」
```

### 3. 視覺回歸測試

```javascript
「請進行視覺回歸測試：
1. 擷取當前頁面截圖（baseline）
2. 執行 loadData()
3. 再次擷取截圖
4. 比較差異」
```

## 與傳統開發流程整合

### 完整 TDD 流程

```bash
# 1. 在本地編輯測試
code testRunner.js

# 2. 推送至 GAS
npm run push

# 3. 在 GAS 執行測試
# 執行 runAllTests()

# 4. 在瀏覽器開啟 Web App

# 5. 透過 MCP 自動化驗證
# AI: 「執行 Web App 整合測試」

# 6. 檢視測試報告
```

### 監視模式 + 自動測試

```bash
# Terminal 1: 監視模式
npm run watch

# VS Code: 編輯檔案並儲存
# → 自動推送至 GAS

# VS Code Chat: 請求自動測試
「檔案已更新，請重新載入 Web App 並執行測試」

# MCP 自動：
# 1. 重新整理頁面
# 2. 執行測試
# 3. 回報結果
```

## 除錯工作流程

### 當測試失敗時

```javascript
「測試失敗時請協助除錯：
1. 擷取頁面快照查看 DOM 結構
2. 開啟 Console 查看錯誤訊息
3. 執行 debugExamObject() 檢查資料
4. 截圖儲存錯誤狀態
5. 取得 Network 請求列表
6. 產生除錯報告」
```

### 效能瓶頸分析

```javascript
「請分析效能瓶頸：
1. 啟動效能追蹤
2. 執行完整排程流程
3. 停止追蹤
4. 分析 LCPBreakdown insight
5. 找出最慢的操作
6. 提供優化建議」
```

## 最佳實踐

### ✅ 建議做法

- 使用文字快照而非截圖（更快、更準確）
- 在測試前先驗證頁面已載入
- 使用 `wait_for` 等待動態內容
- 記錄測試步驟以便重現
- 定期清理測試資料

### ❌ 避免事項

- 不要在 MCP 中執行需要驗證的操作
- 不要過度依賴視覺比對（易受影響）
- 不要在生產環境執行破壞性測試
- 不要忽略網路錯誤

## 常見問題

### Q: MCP 找不到頁面？

A: 確認：
1. Chrome 瀏覽器已開啟
2. 頁面 URL 正確
3. VS Code MCP 擴充功能已啟用

### Q: JavaScript 執行失敗？

A: 檢查：
1. 頁面是否完全載入
2. 函式是否定義在全域範圍
3. 是否有 JavaScript 錯誤（查看 Console）

### Q: 效能追蹤無資料？

A: 確認：
1. 頁面已重新載入
2. 追蹤已正確啟動和停止
3. 等待足夠時間讓追蹤完成

## 相關資源

- [Chrome DevTools Protocol](https://chromedevtools.github.io/devtools-protocol/)
- [Model Context Protocol](https://modelcontextprotocol.io/)
- 本專案 [DEVELOPMENT.md](./DEVELOPMENT.md)

## 實戰演練

試試看以下挑戰：

1. **建立自動化測試套件**
   - 使用 MCP 自動執行所有測試
   - 產生測試報告

2. **效能基準測試**
   - 測量各函式執行時間
   - 建立效能基準線

3. **視覺回歸測試**
   - 擷取 baseline 截圖
   - 偵測 UI 變更

4. **整合 CI/CD**
   - 在部署前自動執行測試
   - 驗證 Web App 功能正常

開始使用 Chrome DevTools MCP 提升您的開發效率！🚀
