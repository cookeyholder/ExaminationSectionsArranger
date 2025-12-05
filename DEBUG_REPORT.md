# 🔍 除錯報告：Web App 按鈕無反應

## 診斷日期
2025-11-06

## 🐛 問題描述

在 Web App 中：
- ✅ 頁面正常載入
- ✅ 介面顯示完整（標題、按鈕、資料區）
- ❌ 「重新載入資料」按鈕無反應
- ❌ 「下載 JSON」按鈕無反應  
- ❌ 「執行測試」按鈕無反應
- ❌ 資料區顯示錯誤：`Unexpected token '<', "<!doctype "... is not valid JSON`

## 🔍 Chrome DevTools 診斷結果

### 1. 頁面結構檢查
✅ 頁面成功載入
✅ 按鈕元素存在：
- `uid=2_13` - 🔄 重新載入資料
- `uid=2_14` - 💾 下載 JSON
- `uid=2_15` - 🧪 執行測試

### 2. Network 請求分析
發現關鍵問題：

**請求 URL：**
```
https://n-qbcnzpz3sj7ob6sbbdrg2bjaswhe5nwiayedd5q-0lu-script.googleusercontent.com/userCodeAppPanel?mode=api
```

**預期回應：** JSON
**實際回應：** HTML

**回應內容：**
```html
<!doctype html>
<style nonce="...">...</style>
<iframe id="userHtmlFrame" ...></iframe>
```

**Content-Type：** `text/html; charset=utf-8` ❌  
**應該是：** `application/json` ✅

### 3. 根本原因

Google Apps Script **Web App 使用快取版本**！

即使已經：
- ✅ 修正 `devTools.js` 程式碼
- ✅ 推送至 GAS (`clasp push`)
- ✅ 重新載入瀏覽器頁面

**但是** Web App 仍使用舊的部署版本，該版本的 `doGet(e)` 函式處理 `mode=api` 時回傳 HTML 而非 JSON。

## ✅ 解決方案

### 方法 1：重新部署（推薦）

1. 開啟 Apps Script 編輯器：
   ```bash
   npm run open
   ```

2. 點擊右上角「部署」→「管理部署作業」

3. 點擊目前部署旁的 ✏️ 編輯圖示

4. 在「版本」下拉選單選擇「新版本」

5. 輸入版本說明：「修正 API 回應格式」

6. 點擊「部署」

7. 重新載入 Web App 頁面（Ctrl+Shift+R 或 Cmd+Shift+R）

### 方法 2：建立新部署

1. 在 Apps Script 編輯器：「部署」→「新增部署作業」

2. 類型：網頁應用程式

3. 執行身分：我

4. 存取權：所有人

5. 點擊「部署」

6. 使用新的 Web App URL

### 方法 3：使用測試部署（開發推薦）

1. 在 Apps Script 編輯器：「部署」→「測試部署作業」

2. 複製測試 URL

3. 測試部署會自動使用最新程式碼，無需重新部署

## 📊 修正前後對比

### 修正前（舊版本）
```javascript
// doGet 回傳 HTML 字串串接
var html = '<!DOCTYPE html>' + '...';
return HtmlService.createHtmlOutput(html);
```

**問題：** `mode=api` 也被包裝在 iframe 中

### 修正後（新版本）
```javascript
function doGet(e) {
  var mode = e && e.parameter && e.parameter.mode ? e.parameter.mode : 'viewer';
  
  if (mode === 'api') {
    return serveDataSnapshot(); // 直接回傳 JSON
  }
  
  return HtmlService.createHtmlOutputFromFile('index');
}
```

**改進：**
- ✅ `mode=api` 回傳純 JSON
- ✅ 使用外部 HTML 檔案（更清晰）
- ✅ Content-Type 正確設定為 `application/json`

## 🎯 驗證步驟

重新部署後，執行以下檢查：

### 1. 測試 API 端點

在瀏覽器開啟：
```
https://your-webapp-url/exec?mode=api
```

**預期結果：**
- 顯示純 JSON 資料
- 包含 `timestamp`, `exam.population`, `exam.statistics` 等欄位

### 2. 測試按鈕功能

在 Web App 中：
- 點擊「🔄 重新載入資料」→ 應顯示統計卡片
- 點擊「💾 下載 JSON」→ 應下載 `.json` 檔案
- 點擊「🧪 執行測試」→ 應顯示提示訊息

### 3. 透過 MCP 自動驗證

```
請測試 Web App 功能：
1. 重新載入頁面
2. 等待 3 秒
3. 檢查是否顯示統計數字
4. 點擊「重新載入資料」按鈕
5. 確認資料更新
6. 截圖結果
```

## 📝 學習重點

### Google Apps Script Web App 部署機制

1. **`clasp push`** - 只更新腳本編輯器中的程式碼
2. **部署版本** - Web App 使用特定版本，不會自動更新
3. **測試部署** - 永遠使用最新程式碼（適合開發）
4. **正式部署** - 使用固定版本（適合生產環境）

### 開發建議

- **開發階段**：使用「測試部署作業」URL
- **測試階段**：建立新版本並驗證
- **生產階段**：使用穩定版本部署

## 🔄 下次開發流程

```bash
# 1. 修改程式碼
code devTools.js

# 2. 推送至 GAS
npm run push

# 3. 選項 A：使用測試部署（推薦）
# 在編輯器：部署 > 測試部署作業

# 3. 選項 B：建立新版本
# 在編輯器：部署 > 管理部署作業 > 編輯 > 新版本

# 4. 重新載入頁面（強制重新整理）
Cmd+Shift+R (Mac) 或 Ctrl+Shift+R (Windows/Linux)
```

## 📸 截圖證據

詳見：`debug-screenshot.png`

## ✅ 修正狀態

- ✅ 問題已診斷
- ✅ 程式碼已修正
- ✅ 已推送至 GAS
- ⏳ 等待重新部署
- ⏳ 等待驗證

## 🚀 後續行動

1. **立即**：重新部署 Web App
2. **驗證**：測試所有按鈕功能
3. **文件**：更新 QUICKSTART.md 加入部署說明
4. **自動化**：考慮建立部署腳本

---

**診斷工具：** Chrome DevTools MCP  
**問題嚴重度：** 中（功能無法使用）  
**修復難度：** 低（只需重新部署）  
**預估修復時間：** 2 分鐘  

**報告產生時間：** 2025-11-06 10:09 CST
