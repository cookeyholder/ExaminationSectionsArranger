# 重構提案：領域模型統一化

## Why

現有的排程系統使用手動建立的物件結構（`createEmptySessionRecord()` 和 `createEmptyClassroomRecord()`），導致：

1. **重複程式碼**: 統計邏輯在多處重複實作
2. **維護困難**: 新增統計維度需要修改多個函式
3. **缺乏一致性**: Session 和 Classroom 結構相似但實作方式不同
4. **擴充性差**: 無法輕鬆新增 Exam 層級的統計

需要建立統一的領域模型系統，提供：
- 通用的統計容器建構器
- 三層階層結構（Exam → Session → Classroom）
- 可配置的統計維度
- 向下分配學生的標準機制

## What Changes

- **新增** `domainModels.js`: 實作通用統計容器與三層領域模型
  - `createStatisticsContainer()`: 通用建構器
  - `createClassroomRecord()`: 試場模型
  - `createSessionRecord()`: 節次模型  
  - `createExamRecord()`: 考試模型

- **新增** `examService.js`: Exam 物件的服務層
  - `createExamFromSheet()`: 從工作表建立 Exam 物件
  - `saveExamToSheet()`: 將 Exam 物件存回工作表
  - `getColumnIndices()`: 取得欄位索引對映

- **重寫** `scheduling.js` 中的所有排程函式:
  - `scheduleCommonSubjectSessions()`: 使用 Exam 模型
  - `scheduleSpecializedSubjectSessions()`: 使用 Exam 模型
  - `assignExamRooms()`: 使用 `distributeToChildren()`
  - `allocateBagIdentifiers()`: 使用 Exam 模型
  - `populateSessionTimes()`: 使用 Exam 模型
  - `updateBagAndClassPopulations()`: 使用 Exam 模型

- **重寫** 排序函式統一使用 Exam 模型:
  - `sortFilteredStudentsBySubject()`
  - `sortFilteredStudentsByClassSeat()`
  - `sortFilteredStudentsBySessionRoom()`

- **移除** 舊的建立函式:
  - `createEmptyClassroomRecord()`
  - `createEmptySessionRecord()`
  - `buildSessionStatistics()`

- **更新** `AGENTS.md`: 新增領域模型章節說明

## Impact

### Affected Specs
- `scheduling`: 核心排程邏輯完全重寫
- `data-models`: 新增領域模型規範

### Affected Code
- `scheduling.js`: 約 300 行程式碼重寫
- `domainModels.js`: 新增約 260 行
- `examService.js`: 新增約 150 行
- `AGENTS.md`: 新增領域模型章節

### Breaking Changes
- **BREAKING**: 移除 `createEmptySessionRecord()` 和 `createEmptyClassroomRecord()`
- **BREAKING**: 移除 `buildSessionStatistics()`，改用 `createExamFromSheet()`
- **BREAKING**: 所有排程函式的內部實作完全改變（但對外 API 保持不變）

### Migration Path
1. 部署新的 `domainModels.js` 和 `examService.js`
2. 逐一重寫排程函式
3. 完整測試所有流程
4. 移除舊函式

### Risks
- 統計結果可能因實作差異而不同（需嚴格驗證）
- 效能可能受 getter 屬性影響（需監控執行時間）
- 大規模程式碼變更可能引入 bug

### Rollback Plan
保留 `master` 分支穩定版本，重構在 `refactor/domain-models` 分支進行。若發現問題可立即切回 master。
