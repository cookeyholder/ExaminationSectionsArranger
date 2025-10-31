# 設計文件：領域模型重構

## Context

現有排程系統使用手動建立的物件（`createEmptySessionRecord()` 和 `createEmptyClassroomRecord()`），每個物件都有自己的統計 getter，導致程式碼重複、難以維護和擴充。

**當前問題**：
- 統計邏輯散落在多個函式中
- 新增統計維度需要修改多處
- Session 和 Classroom 結構相似但實作獨立
- 缺乏 Exam 層級的統計能力

**利害關係人**：
- 開發者：需要維護和擴充系統
- 使用者（註冊組）：依賴系統產生正確的補考安排

## Goals / Non-Goals

### Goals
1. **統一領域模型**：建立通用的統計容器，消除重複程式碼
2. **三層架構**：實作 Exam → Session → Classroom 階層結構
3. **服務層抽象**：封裝 Exam 物件的建立、載入與儲存邏輯
4. **保持功能一致**：重構後的輸出結果與現有系統完全相同

### Non-Goals
- ❌ 改變業務規則（科別年級互斥、容量限制等）
- ❌ 優化執行效能（目前未遇到效能瓶頸）
- ❌ 建立自動化測試框架（仍採用手動測試）
- ❌ 向後相容舊的 API（完全重寫，不保留相容層）

## Decisions

### Decision 1: 使用組合模式而非繼承

**Why**：JavaScript 的原型繼承不適合這種情境，組合模式更靈活

**Implementation**：
```javascript
function createStatisticsContainer(config) {
  // 動態配置統計維度
  const { statisticsDimensions = [], children = null } = config;
  
  // 透過 Object.defineProperty 動態建立 getter
  statisticsDimensions.forEach(dimension => {
    Object.defineProperty(container, dimension.name, {
      get: function() { /* 統計邏輯 */ }
    });
  });
}
```

**Alternatives considered**：
- 類別繼承：需要定義抽象基類，但 Apps Script 的 ES5 環境支援有限
- Mixin 模式：會讓物件結構更複雜，不如組合直觀

### Decision 2: 保留三層容器的 students 陣列

**Why**：現有排程邏輯是先填充 Session.students，再分配到 Classroom

**Implementation**：
- Classroom.students：真實資料來源（Single Source of Truth）
- Session.students：中間處理用
- Exam.students：中間處理用

**Alternatives considered**：
- 計算屬性：Session.students 從 Classroom 聚合而來
  - 優點：避免資料重複
  - 缺點：需要大幅修改現有排程邏輯，風險高
  - **決策**：為了降低重構風險，保持現有資料流

### Decision 3: 建立 examService.js 服務層

**Why**：封裝 Exam 物件與工作表之間的轉換邏輯

**API**：
```javascript
createExamFromSheet()     // 從工作表建立 Exam
saveExamToSheet(exam)     // 儲存 Exam 到工作表
getColumnIndices()        // 取得欄位索引對映
```

**Alternatives considered**：
- 將服務方法放在 domainModels.js：會混合模型定義與資料存取
- 將服務方法放在 scheduling.js：會讓 scheduling.js 職責過重
- **決策**：獨立的服務層更符合單一職責原則

### Decision 4: 分階段重寫，每階段獨立提交

**Why**：降低風險，便於追蹤問題

**階段劃分**：
1. 建立 examService.js（新增，無破壞性）
2. 重寫排程函式（內部實作變更，外部 API 不變）
3. 重寫輔助函式（同上）
4. 重寫排序函式（同上）
5. 移除舊函式（破壞性變更）
6. 整合測試

## Architecture

### 資料流向

```
┌─────────────────┐
│ 工作表資料      │
└────────┬────────┘
         │ createExamFromSheet()
         ▼
┌─────────────────┐
│ Exam 物件       │
│  sessions[]     │
└────────┬────────┘
         │ 排程邏輯
         ▼
┌─────────────────┐
│ Session 物件    │
│  students[]     │
│  classrooms[]   │
└────────┬────────┘
         │ distributeToChildren()
         ▼
┌─────────────────┐
│ Classroom 物件  │
│  students[]     │ ← Single Source of Truth
└────────┬────────┘
         │ saveExamToSheet()
         ▼
┌─────────────────┐
│ 工作表資料      │
└─────────────────┘
```

### 模組依賴

```
globals.js
  └─ helpers.js
      └─ domainModels.js
          └─ examService.js
              └─ scheduling.js
                  └─ menu.js
```

## Implementation Details

### 關鍵技術點

**1. 動態統計屬性**
```javascript
Object.defineProperty(container, 'departmentDistribution', {
  get: function() {
    const stats = {};
    this.students.forEach(student => {
      const key = student[0]; // 科別
      stats[key] = (stats[key] || 0) + 1;
    });
    return stats;
  }
});
```

**2. 子容器分配**
```javascript
session.distributeToChildren((student, classrooms) => {
  return student[roomIndex]; // 回傳目標 classroom 索引
});
```

**3. 欄位索引抽象**
```javascript
const columns = getColumnIndices();
student[columns.subject]  // 取代 student[7]
```

### 檔案載入順序

在 `appsscript.json` 中確保順序（Apps Script 會按字母順序載入，需要重新命名或使用打包工具）：

```
1. globals.js
2. helpers.js
3. domainModels.js
4. examService.js
5. dataPreparation.js
6. scheduling.js
7. reportGeneration.js
8. printing.js
9. menu.js
```

## Risks / Trade-offs

### Risk 1: 統計結果不一致

**Likelihood**: 中
**Impact**: 高

**Mitigation**：
- 在階段 0 建立基準測試資料
- 每個階段完成後比對輸出結果
- 重點檢查人數統計、節次分配、試場分配

### Risk 2: 效能下降

**Likelihood**: 低
**Impact**: 中

**Mitigation**：
- 記錄重構前的執行時間
- 監控每階段的執行時間變化
- 若 getter 屬性造成效能問題，改為快取結果

**Trade-off**：動態統計 vs 預計算
- 選擇：動態統計（每次存取時計算）
- 原因：程式碼更簡潔，Apps Script 執行時間限制 6 分鐘足夠

### Risk 3: 資料遺失

**Likelihood**: 低
**Impact**: 極高

**Mitigation**：
- 在測試副本上操作
- 保留原始資料備份
- 每階段提交前驗證資料完整性

## Migration Plan

### Phase 1: 準備階段
1. 在測試副本上執行現有流程
2. 匯出所有輸出工作表為 CSV
3. 記錄執行時間

### Phase 2: 實作階段
依照 tasks.md 逐步執行，每完成一個階段：
1. 執行 `clasp push` 部署
2. 在測試副本上執行流程
3. 比對輸出結果
4. 記錄執行時間
5. Git commit

### Phase 3: 驗證階段
1. 執行完整流程 3 次
2. 確認所有輸出一致
3. 效能差異在 ±10% 內
4. 合併到 master

### Phase 4: 部署階段
1. 通知使用者（註冊組）
2. 在正式環境執行一次測試
3. 監控第一次實際使用
4. 準備快速回滾方案（切回 master 分支）

## Rollback Plan

若發現重大問題：

```bash
# 1. 立即切回 master
git checkout master
clasp push

# 2. 通知使用者暫停使用

# 3. 在 refactor/domain-models 分支修復問題

# 4. 完整測試後再次部署
```

## Open Questions

1. **Q**: 是否需要在排程過程中驗證資料一致性？
   **A**: 暫不實作，保持現有行為

2. **Q**: 是否要重構 reportGeneration.js 和 printing.js？
   **A**: 本次重構範圍僅限 scheduling.js，報表模組保持不變

3. **Q**: 是否要建立單元測試？
   **A**: 暫不建立，持續使用手動測試

## Success Criteria

✅ 所有現有功能正常運作
✅ 輸出結果與重構前完全一致
✅ 無新增錯誤或警告
✅ 執行時間差異在 ±10% 內
✅ 舊函式已完全移除
✅ 程式碼通過 `clasp push` 檢查
✅ 文件已更新（AGENTS.md）
