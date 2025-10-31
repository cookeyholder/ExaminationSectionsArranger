# 實作計畫總結

## 📋 已建立的文件

### 1. **design.md** - 技術設計文件
記錄所有重要的技術決策：
- ✅ 使用組合模式而非繼承
- ✅ 保留三層容器的 students 陣列（降低風險）
- ✅ 建立 examService.js 服務層
- ✅ 分階段重寫，每階段獨立提交
- ✅ 資料流向圖和模組依賴圖
- ✅ 風險評估與回滾計畫

### 2. **tasks.md** - 詳細任務清單（0/52 完成）
包含 9 個階段共 52 個任務：
- **階段 0**: 準備基準測試資料（3 個任務）
- **階段 1**: 建立核心服務層（5 個任務）
- **階段 2**: 重寫排程邏輯（5 個任務）
- **階段 3**: 重寫輔助函式（5 個任務）
- **階段 4**: 重寫排序函式（5 個任務）
- **階段 5**: 清理舊程式碼（5 個任務）
- **階段 6**: 整合測試與效能驗證（8 個任務）
- **階段 7**: 文件更新（4 個任務）
- **階段 8**: 部署準備（6 個任務）

每個任務都包含：
- 明確的驗收標準
- 具體的實作指引
- 測試方法

### 3. **IMPLEMENTATION_GUIDE.md** - 執行指南
包含：
- 快速開始步驟
- 每個階段的詳細說明（目標、關鍵檔案、實作順序、驗收標準）
- 程式碼範例
- 常見問題 FAQ
- 檢查清單速查表
- 預估時間（總計 13.5 小時）

---

## 🎯 實作計畫特點

### 1. **漸進式重構**
不是一次性重寫所有程式碼，而是：
1. 先建立新的服務層（examService.js）
2. 逐步替換現有函式
3. 最後移除舊程式碼

**好處**：每個階段都可以獨立測試和提交，降低風險。

### 2. **嚴格的測試基準**
階段 0 建立基準測試資料：
- 匯出 3 個 CSV 檔案作為比對基準
- 記錄執行時間
- 每個階段都與基準比對

**好處**：確保重構後輸出完全一致。

### 3. **明確的驗收標準**
每個任務都有可驗證的標準，例如：
- ✅ `exam.population` 等於工作表總列數 - 1
- ✅ `grep "createEmptyClassroomRecord" *.js` 無結果
- ✅ 執行時間差異在 ±10% 內

**好處**：清楚知道何時算完成。

### 4. **完整的文件**
提供 3 個層次的文件：
- **design.md**: 給未來維護者看的技術決策
- **tasks.md**: 給實作者看的檢查清單
- **IMPLEMENTATION_GUIDE.md**: 給執行者看的操作手冊

---

## 📊 進度追蹤

使用以下指令查看進度：

```bash
# 查看 OpenSpec 狀態
openspec list

# 查看詳細任務
cat openspec/changes/refactor-domain-models/tasks.md | grep -c "\[ \]"  # 未完成任務數
cat openspec/changes/refactor-domain-models/tasks.md | grep -c "\[x\]"  # 已完成任務數
```

---

## 🚀 開始實作

### 立即執行

```bash
# 1. 確認在正確分支
git checkout refactor/domain-models
git status

# 2. 查看實作指南
open openspec/changes/refactor-domain-models/IMPLEMENTATION_GUIDE.md

# 3. 開始階段 0
# - 建立試算表測試副本（手動操作）
# - 執行現有流程並匯出 CSV
```

### 建議工作節奏

**第 1 天**（3-4 小時）：
- ✅ 階段 0: 準備基準（0.5 小時）
- ✅ 階段 1: 建立服務層（2 小時）
- ✅ 階段 2: 重寫排程邏輯（部分，1.5 小時）

**第 2 天**（3-4 小時）：
- ✅ 階段 2: 重寫排程邏輯（完成，1.5 小時）
- ✅ 階段 3: 重寫輔助函式（2 小時）

**第 3 天**（3-4 小時）：
- ✅ 階段 4: 重寫排序函式（1.5 小時）
- ✅ 階段 5: 清理舊程式碼（0.5 小時）
- ✅ 階段 6: 整合測試（1 小時）

**第 4 天**（2-3 小時）：
- ✅ 階段 6: 整合測試（完成，1 小時）
- ✅ 階段 7: 文件更新（1 小時）
- ✅ 階段 8: 部署準備（1 小時）

---

## 📝 關鍵檔案位置

```
/Users/cookeyholder/projects/ExaminationSectionsArranger/
├── openspec/changes/refactor-domain-models/
│   ├── design.md                 ← 技術設計文件
│   ├── tasks.md                  ← 任務檢查清單
│   ├── IMPLEMENTATION_GUIDE.md   ← 執行指南（本文件的詳細版）
│   ├── proposal.md               ← 變更提案
│   └── specs/
│       ├── data-models/spec.md   ← 資料模型規範
│       └── scheduling/spec.md    ← 排程邏輯規範
├── domainModels.js               ← 領域模型實作（已完成）
├── examService.js                ← 服務層（待建立）
├── scheduling.js                 ← 排程邏輯（待重構）
└── REFACTORING_PLAN.md           ← 原始重構計畫
```

---

## ✅ 驗證清單

實作前檢查：
- [x] `domainModels.js` 已存在於 `refactor/domain-models` 分支
- [x] OpenSpec 提案已驗證通過（`openspec validate --strict`）
- [x] `design.md` 已建立
- [x] `tasks.md` 已更新為 52 個詳細任務
- [x] `IMPLEMENTATION_GUIDE.md` 已建立
- [ ] 試算表測試副本已建立（手動操作）
- [ ] 基準測試資料已準備

---

## 🎓 學習重點

透過這次重構，你將學到：

1. **OpenSpec 工作流程**
   - 如何撰寫 design.md 記錄技術決策
   - 如何分解任務並追蹤進度
   - 如何使用規範驗證（spec-driven development）

2. **領域驅動設計**
   - 如何設計三層階層式領域模型
   - 如何使用組合模式建立通用容器
   - 如何定義 Single Source of Truth

3. **重構技巧**
   - 如何進行漸進式重構
   - 如何建立測試基準
   - 如何處理風險和回滾

4. **Google Apps Script 最佳實踐**
   - 如何組織模組和管理依賴
   - 如何使用 Clasp CLI 部署
   - 如何處理執行時間限制

---

## 💡 下一步行動

**現在就開始**：

1. 開啟 `IMPLEMENTATION_GUIDE.md` 查看詳細步驟
2. 建立試算表測試副本
3. 執行階段 0 的 3 個任務
4. 開始階段 1 建立 `examService.js`

**遇到問題時**：

1. 查看 `IMPLEMENTATION_GUIDE.md` 的常見問題區
2. 檢查 `design.md` 的技術決策說明
3. 參考 `REFACTORING_PLAN.md` 的程式碼範例

**需要協助時**：

- 使用 `Logger.log()` 追蹤資料流向
- 在 Apps Script 編輯器中執行測試函式
- 比對基準 CSV 檔案找出差異

---

準備好了嗎？開始實作吧！ 🚀
