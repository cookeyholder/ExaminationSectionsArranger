# GAS 開發與測試指南

本文件說明如何使用 clasp 搭配 Chrome DevTools MCP 進行 Google Apps Script 開發和測試。

## 目錄
- [環境設定](#環境設定)
- [開發工作流程](#開發工作流程)
- [測試方法](#測試方法)
- [除錯技巧](#除錯技巧)
- [常用指令](#常用指令)

---

## 環境設定

### 1. 安裝 clasp

```bash
npm install -g @google/clasp
```

### 2. 登入 Google 帳號

```bash
clasp login
```

### 3. 驗證專案連結

檢查 `.clasp.json` 確認已連結到正確的 Apps Script 專案：

```bash
cat .clasp.json
```

---

## 開發工作流程

### 基本流程

```
本地編輯 → clasp push → GAS 執行 → 檢視結果 → 重複
```

### 1. 本地開發

在 VS Code 中編輯 `.js` 檔案：

```bash
# 開啟專案
code .

# 編輯檔案
# - domainModels.js
# - examService.js
# - scheduling.js
# 等等...
```

### 2. 推送至 GAS

```bash
# 推送所有變更
npm run push

# 或使用 clasp 原生指令
clasp push
```

### 3. 監視模式（自動推送）

開發時可啟用監視模式，檔案變更時自動推送：

```bash
npm run watch
```

> ⚠️ **注意**：監視模式會在每次檔案儲存時推送，請確保程式碼可執行。

### 4. 開啟 Apps Script 編輯器

```bash
npm run open
```

或直接在瀏覽器中開啟試算表，進入「擴充功能 > Apps Script」。

---

## 測試方法

### 方法 1: 使用內建測試執行器（推薦）

本專案提供完整的測試框架（`testRunner.js`）：

#### 快速測試
在 Apps Script 編輯器中執行：

```javascript
quickTest()
```

驗證項目：
- ✅ 工作表連線
- ✅ Exam 物件建立
- ✅ 統計功能

#### 完整測試套件
執行所有測試：

```javascript
runAllTests()
```

包含：
- 領域模型測試
- ExamService 測試
- 排程邏輯測試

#### 單一測試套件
```javascript
testDomainModels()    // 只測試領域模型
testExamService()     // 只測試 ExamService
testSchedulingLogic() // 只測試排程邏輯
```

### 方法 2: 使用開發工具進行視覺化測試

#### 部署 Web App

1. 在 Apps Script 編輯器中：
   - 點擊「部署 > 新增部署作業」
   - 選擇「網頁應用程式」
   - 執行身分：選擇自己
   - 存取權：「所有人」或「僅限自己」
   - 部署

2. 取得 Web App URL

3. 在瀏覽器中開啟該 URL，即可看到視覺化測試介面

#### 使用 Chrome DevTools MCP

開啟 Web App 後，使用 Chrome DevTools 進行互動式除錯：

```bash
# 在 VS Code 中，MCP 伺服器會自動連接 Chrome
# 可以直接與頁面互動、檢視資料、執行測試
```

### 方法 3: 手動測試

1. 開啟試算表
2. 在「補考排程」選單中執行功能
3. 檢查輸出資料

---

## 除錯技巧

### 1. 使用 Logger

在程式碼中加入 Logger 輸出：

```javascript
Logger.log('除錯訊息: ' + JSON.stringify(data));
```

查看日誌：

```bash
npm run logs
```

或在 Apps Script 編輯器中：「檢視 > 執行紀錄」

### 2. 使用除錯工具函式

`devTools.js` 提供多個除錯函式：

```javascript
// 顯示當前 Exam 物件資訊
debugExamObject()

// 建立資料快照
const snapshot = createDataSnapshot()
Logger.log(JSON.stringify(snapshot, null, 2))

// 測量效能
measurePerformance('assignSessionTimesForExam')

// 比較兩個快照差異
const before = createDataSnapshot()
// ... 執行操作 ...
const after = createDataSnapshot()
const diff = compareExamSnapshots(before, after)
```

### 3. 清除測試資料

```javascript
// 清除排程資料（保留學生清單）
clearSchedulingData()

// 建立測試資料
createTestData()
```

### 4. Web App 即時檢視

部署 Web App 後，在瀏覽器中可以：

- 🔄 重新載入資料
- 💾 下載 JSON 快照
- 📊 查看統計圖表
- 🧪 執行測試（開發中）

### 5. 使用 Chrome DevTools

透過 MCP 整合，可以：

1. 在 Web App 中設定中斷點
2. 檢視變數值
3. 單步執行程式碼
4. 監控網路請求
5. 分析效能

---

## 常用指令

### 開發指令

```bash
# 推送程式碼至 GAS
npm run push

# 拉取 GAS 的最新程式碼
npm run pull

# 監視模式（自動推送）
npm run watch

# 開啟 Apps Script 編輯器
npm run open

# 開啟 Web App
npm run open-webapp
```

### 版本管理

```bash
# 檢視當前版本與部署狀態
npm run status

# 建立新版本
npm run create-version

# 部署（推送 + 建立部署）
npm run deploy
```

### 測試指令

```bash
# 執行測試（需在 GAS 編輯器中手動執行）
npm run test

# 快速測試（需在 GAS 編輯器中手動執行）
npm run quick-test
```

### 除錯指令

```bash
# 查看執行日誌
npm run logs

# 開啟 Web App 進行除錯
npm run debug
```

---

## 開發最佳實踐

### 1. 程式碼變更流程

```bash
# 1. 拉取最新程式碼
npm run pull

# 2. 在本地編輯

# 3. 推送至 GAS
npm run push

# 4. 在 GAS 編輯器中測試
# 執行 quickTest() 或 runAllTests()

# 5. 確認無誤後提交 git
git add .
git commit -m "feat: 新增功能"
git push
```

### 2. 測試驅動開發

```bash
# 1. 先寫測試
# 在 testRunner.js 中新增測試案例

# 2. 推送並執行測試
npm run push
# 在 GAS 中執行測試，確認失敗

# 3. 實作功能

# 4. 推送並執行測試
npm run push
# 確認測試通過

# 5. 重構（如需要）
```

### 3. 使用 Web App 進行即時開發

```bash
# 1. 部署 Web App（只需一次）

# 2. 開啟監視模式
npm run watch

# 3. 在瀏覽器中開啟 Web App

# 4. 編輯程式碼
# 儲存後自動推送

# 5. 在 Web App 中重新載入
# 立即看到變更結果
```

### 4. 除錯工作流程

當遇到問題時：

1. **加入 Logger**
   ```javascript
   Logger.log('檢查點 1: ' + JSON.stringify(data));
   ```

2. **推送並執行**
   ```bash
   npm run push
   # 在 GAS 執行函式
   ```

3. **查看日誌**
   ```bash
   npm run logs
   ```

4. **使用除錯工具**
   ```javascript
   debugExamObject()  // 查看完整物件狀態
   ```

5. **Web App 視覺化**
   - 在 Web App 中查看資料快照
   - 下載 JSON 進行詳細分析

---

## 常見問題

### Q: 推送後程式碼沒有更新？

A: 確認 `.clasp.json` 中的 `scriptId` 正確，並嘗試：

```bash
clasp pull  # 先拉取
clasp push  # 再推送
```

### Q: 如何查看執行錯誤？

A: 使用以下方法：

1. Apps Script 編輯器的「檢視 > 執行紀錄」
2. 執行 `npm run logs`
3. 在程式碼中加入 try-catch 並記錄錯誤

### Q: Web App 更新後沒反應？

A: 需要重新部署：

1. 在 Apps Script 編輯器中
2. 「部署 > 管理部署作業」
3. 點擊編輯圖示
4. 選擇「新版本」
5. 部署

### Q: 如何在本地執行測試？

A: GAS 程式碼無法直接在 Node.js 環境執行，必須：

1. 推送至 GAS：`npm run push`
2. 在 Apps Script 編輯器中執行測試函式
3. 或使用 Web App 進行視覺化測試

### Q: 如何整合 Chrome DevTools？

A: 當 Web App 部署後：

1. 在瀏覽器開啟 Web App
2. 按 F12 開啟 Chrome DevTools
3. 可使用 Console、Network、Performance 等工具
4. 透過 MCP 整合可在 VS Code 中控制

---

## 進階技巧

### 條件編譯（開發/正式環境）

在 `globals.js` 中：

```javascript
const IS_DEVELOPMENT = true;  // 手動切換

function log(message) {
  if (IS_DEVELOPMENT) {
    Logger.log(message);
  }
}
```

### 效能分析

使用 `measurePerformance`：

```javascript
const result = measurePerformance('assignSessionTimesForExam');
Logger.log(`執行時間: ${result.duration}ms`);
```

### 資料快照比較

追蹤變更：

```javascript
const before = createDataSnapshot();

// 執行操作
assignSessionTimesForExam();

const after = createDataSnapshot();
const diff = compareExamSnapshots(before, after);

Logger.log('變更詳情: ' + JSON.stringify(diff, null, 2));
```

---

## 相關資源

- [clasp 官方文件](https://github.com/google/clasp)
- [Apps Script 文件](https://developers.google.com/apps-script)
- [本專案 AGENTS.md](./AGENTS.md) - AI 協作指引
- [REFACTORING_PLAN.md](./REFACTORING_PLAN.md) - 重構計畫

---

## 授權

請參考專案根目錄的 LICENSE 檔案。
